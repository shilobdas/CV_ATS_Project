import os
import re
from pathlib import Path

import pymupdf # PyMuPDF
import ollama
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib import colors as pdf_colors
from reportlab.pdfbase.pdfmetrics import stringWidth


# ============================================================
# CONFIGURATION
# ============================================================

MODEL = os.environ.get("OLLAMA_MODEL", "llama3.1:8b")

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

# ------------------------------------------------------------
# Layout tuning constants
# ------------------------------------------------------------

# A block at least this wide (relative to the page's content width)
# is treated as a full-width banner rather than as column content.
BANNER_WIDTH_RATIO = 0.85

# A vertical white gap must be at least this fraction of the content
# width before it is accepted as a real column gutter.
MIN_GUTTER_RATIO = 0.025

# Resolution used only when a PDF page has no embedded text and must be
# rasterized for OCR. 300 DPI is a good balance for resume-sized type.
OCR_DPI = 300


def _block_text(block) -> str:
    """
    Join the lines of one PyMuPDF text block into a single string.
    """

    lines = []

    for line in block.get("lines", []):

        parts = [
            span.get("text", "")
            for span in line.get("spans", [])
        ]

        text = "".join(parts).strip()

        if text:
            lines.append(text)

    return "\n".join(lines)


def _collect_blocks(page, textpage=None):
    """
    Return every non-empty text block on the page with its position.
    """

    raw = page.get_text("dict", textpage=textpage)

    blocks = []

    for block in raw.get("blocks", []):

        # type 0 = text block, type 1 = image
        if block.get("type") != 0:
            continue

        text = _block_text(block)

        if not text:
            continue

        x0, y0, x1, y1 = block["bbox"]

        blocks.append(
            {
                "text": text,
                "x0": x0,
                "y0": y0,
                "x1": x1,
                "y1": y1,
            }
        )

    return blocks


def _configure_tessdata() -> None:
    """Locate Tesseract language data for PyMuPDF's OCR integration."""

    configured = os.environ.get("TESSDATA_PREFIX")

    if configured and Path(configured).is_dir():
        return

    candidates = []

    for variable in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        root = os.environ.get(variable)

        if not root:
            continue

        if variable == "LOCALAPPDATA":
            candidates.append(
                Path(root) / "Programs" / "Tesseract-OCR" / "tessdata"
            )
        else:
            candidates.append(Path(root) / "Tesseract-OCR" / "tessdata")

    for candidate in candidates:
        if candidate.is_dir():
            os.environ["TESSDATA_PREFIX"] = str(candidate)
            return


def _ocr_textpage(page):
    """Return a positioned OCR text page for an image-only PDF page."""

    _configure_tessdata()

    try:
        return page.get_textpage_ocr(
            language="eng",
            dpi=OCR_DPI,
            full=True,
        )
    except Exception as exc:
        raise RuntimeError(
            "This PDF page contains no embedded text and OCR could not run. "
            "Install Tesseract OCR with English language data, or set "
            "TESSDATA_PREFIX to its tessdata directory. "
            f"OCR error: {exc}"
        ) from exc


def _find_scanned_page_gutter(page):
    """Find a persistent vertical whitespace gutter in a scanned page."""

    pix = page.get_pixmap(colorspace=pymupdf.csGRAY, alpha=False)
    width = pix.width
    height = pix.height
    samples = pix.samples

    # Ignore the title/header area, where text commonly spans both columns.
    body_top = int(height * 0.16)
    body_bottom = int(height * 0.95)
    max_ink = max(4, int((body_bottom - body_top) * 0.008))

    ink_counts = []

    for x in range(width):
        count = sum(
            samples[(y * width) + x] < 200
            for y in range(body_top, body_bottom)
        )
        ink_counts.append(count)

    search_left = int(width * 0.15)
    search_right = int(width * 0.85)
    runs = []
    start = None

    for x in range(search_left, search_right + 1):
        low_ink = x < search_right and ink_counts[x] <= max_ink

        if low_ink and start is None:
            start = x
        elif not low_ink and start is not None:
            runs.append((start, x))
            start = None

    min_width = width * MIN_GUTTER_RATIO
    max_width = width * 0.15
    candidates = [
        (end - start, start, end)
        for start, end in runs
        if min_width <= end - start <= max_width
    ]

    if not candidates:
        return None

    _, start, end = max(candidates)
    return page.rect.x0 + (((start + end) / 2) / width) * page.rect.width


def _ocr_clip_text(page, clip) -> str:
    """OCR one clipped page region and return its plain text."""

    pix = page.get_pixmap(dpi=OCR_DPI, clip=clip, alpha=False)
    ocr_pdf = pymupdf.open(
        "pdf",
        pix.pdfocr_tobytes(language="eng"),
    )

    try:
        ocr_page = ocr_pdf[0]
        return ocr_page.get_text("text").strip()
    finally:
        ocr_pdf.close()


def _find_gutter(blocks, content_left, content_right):
    """
    Look for a vertical white gap running down the page.

    Projects every block onto the horizontal axis, then finds the
    widest empty stretch that has content on both sides of it.

    Returns the x position of the gutter centre, or None if the page
    does not look like a multi-column layout.
    """

    span = content_right - content_left

    if span <= 0:
        return None

    resolution = 400
    step = span / resolution

    occupied = [False] * (resolution + 1)

    for block in blocks:

        start = int((block["x0"] - content_left) / step)
        end = int((block["x1"] - content_left) / step)

        start = max(0, min(resolution, start))
        end = max(0, min(resolution, end))

        for i in range(start, end + 1):
            occupied[i] = True

    best = None

    i = 0

    while i <= resolution:

        if occupied[i]:
            i += 1
            continue

        j = i

        while j <= resolution and not occupied[j]:
            j += 1

        # Ignore gaps touching the left or right edge - those are
        # page margins, not gutters.
        if i > 0 and j <= resolution:

            width = (j - i) * step

            if best is None or width > best[0]:
                centre = content_left + ((i + j) / 2) * step
                best = (width, centre)

        i = j

    if best is None:
        return None

    width, centre = best

    if width < span * MIN_GUTTER_RATIO:
        return None

    return centre


def _extract_page(page, page_number: int) -> str:
    """
    Extract one page in true reading order.

    Single-column pages are read straight down.

    Two-column pages are split at the gutter and read one whole
    column at a time, so a heading always stays with the text that
    belongs to it.
    """

    blocks = _collect_blocks(page)

    if not blocks:
        print(f"Page {page_number}: no embedded text; running OCR...")
        _configure_tessdata()
        image_gutter = _find_scanned_page_gutter(page)

        if image_gutter is not None:
            header_bottom = page.rect.y0 + page.rect.height * 0.16
            regions = (
                (
                    "PAGE HEADER",
                    pymupdf.Rect(
                        page.rect.x0,
                        page.rect.y0,
                        page.rect.x1,
                        header_bottom,
                    ),
                ),
                (
                    "LEFT COLUMN",
                    pymupdf.Rect(
                        page.rect.x0,
                        header_bottom,
                        image_gutter,
                        page.rect.y1,
                    ),
                ),
                (
                    "RIGHT COLUMN",
                    pymupdf.Rect(
                        image_gutter,
                        header_bottom,
                        page.rect.x1,
                        page.rect.y1,
                    ),
                ),
            )
            parts = []

            for label, clip in regions:
                text = _ocr_clip_text(page, clip)

                if text:
                    parts.append(f"[{label} - OCR]\n{text}")

            if parts:
                return f"--- PAGE {page_number} ---\n" + "\n\n".join(parts)

        ocr_textpage = _ocr_textpage(page)
        blocks = _collect_blocks(page, textpage=ocr_textpage)

    if not blocks:
        return f"--- PAGE {page_number} ---\n[OCR FOUND NO TEXT]"

    content_left = min(b["x0"] for b in blocks)
    content_right = max(b["x1"] for b in blocks)
    content_width = content_right - content_left

    # ---------------------------------------
    # Separate obvious full-width banners
    # ---------------------------------------

    banners = []
    candidates = []

    for block in blocks:

        width = block["x1"] - block["x0"]

        if width >= content_width * BANNER_WIDTH_RATIO:
            banners.append(block)
        else:
            candidates.append(block)

    gutter = None

    if candidates:
        gutter = _find_gutter(
            candidates,
            content_left,
            content_right
        )

    # ---------------------------------------
    # Single column: simple top-to-bottom read
    # ---------------------------------------

    if gutter is None:

        ordered = sorted(
            blocks,
            key=lambda b: (round(b["y0"], 1), b["x0"])
        )

        body = "\n\n".join(b["text"] for b in ordered)

        return f"--- PAGE {page_number} ---\n{body}"

    # ---------------------------------------
    # Two columns
    # ---------------------------------------

    # Anything crossing the gutter is full-width after all.
    crossing = [
        b for b in candidates
        if b["x0"] < gutter < b["x1"]
    ]

    column_blocks = [
        b for b in candidates
        if not (b["x0"] < gutter < b["x1"])
    ]

    banners = banners + crossing

    left_column = sorted(
        [
            b for b in column_blocks
            if (b["x0"] + b["x1"]) / 2 < gutter
        ],
        key=lambda b: (round(b["y0"], 1), b["x0"])
    )

    right_column = sorted(
        [
            b for b in column_blocks
            if (b["x0"] + b["x1"]) / 2 >= gutter
        ],
        key=lambda b: (round(b["y0"], 1), b["x0"])
    )

    first_column_y = min(b["y0"] for b in column_blocks)
    last_column_y = max(b["y1"] for b in column_blocks)

    header = sorted(
        [b for b in banners if b["y1"] <= first_column_y],
        key=lambda b: b["y0"]
    )

    footer = sorted(
        [b for b in banners if b["y0"] >= last_column_y],
        key=lambda b: b["y0"]
    )

    placed = {id(b) for b in header} | {id(b) for b in footer}

    middle = sorted(
        [b for b in banners if id(b) not in placed],
        key=lambda b: b["y0"]
    )

    parts = [
        f"--- PAGE {page_number} (TWO-COLUMN LAYOUT) ---"
    ]

    def add(label, group):

        if group:
            parts.append(
                f"[{label}]\n"
                + "\n\n".join(b["text"] for b in group)
            )

    add("PAGE HEADER", header)
    add("LEFT COLUMN", left_column)
    add("RIGHT COLUMN", right_column)
    add("FULL WIDTH SECTION", middle)
    add("PAGE FOOTER", footer)

    return "\n\n".join(parts)


def read_pdf(file_path):
    """
    Extract text from a PDF in true reading order.

    Plain page.get_text("text") walks the page roughly top to bottom
    across the full page width, so a two-column CV comes out
    interleaved: sidebar content lands in the middle of work
    experience, and bullet points get separated from the job they
    belong to.

    This version reads the page as positioned blocks, detects whether
    there is a real column gutter, and emits one whole column at a
    time.

    Text-based PDFs are extracted directly. Scanned/image-only pages
    automatically fall back to local Tesseract OCR while preserving word
    positions for the same column-ordering logic.
    """

    try:

        print(f"Reading PDF: {file_path}")

        with pymupdf.open(file_path) as pdf:

            print(f"Pages: {len(pdf)}")

            all_pages = []

            for page_number, page in enumerate(pdf, start=1):

                all_pages.append(
                    _extract_page(page, page_number)
                )

            return "\n\n".join(all_pages)

    except Exception as e:
        raise RuntimeError(
            f"Failed to read PDF: {e}"
        ) from e


# ============================================================
# DOCX TEXT EXTRACTION
# ============================================================

def read_docx(file_path: Path) -> str:
    """
    Extract paragraphs and tables from DOCX.
    """

    print(f"\nReading DOCX: {file_path.name}")

    try:
        doc = Document(file_path)

        text_parts = []

        # -------------------------
        # Paragraphs
        # -------------------------

        for paragraph in doc.paragraphs:

            text = paragraph.text.strip()

            if text:
                text_parts.append(text)

        # -------------------------
        # Tables
        # -------------------------

        for table_index, table in enumerate(doc.tables):

            text_parts.append(
                f"\n--- TABLE {table_index + 1} ---"
            )

            for row in table.rows:

                cells = []

                for cell in row.cells:

                    cell_text = cell.text.strip()

                    if cell_text:
                        cells.append(cell_text)

                if cells:
                    text_parts.append(
                        " | ".join(cells)
                    )

        return "\n".join(text_parts).strip()

    except Exception as e:
        raise RuntimeError(
            f"Failed to read DOCX: {e}"
        )


# ============================================================
# UNIVERSAL FILE READER
# ============================================================

def read_cv(file_path: Path) -> str:

    extension = file_path.suffix.lower()

    if extension == ".pdf":

        return read_pdf(file_path)

    elif extension == ".docx":

        return read_docx(file_path)

    else:

        raise ValueError(
            "Unsupported file type. "
            "Only PDF and DOCX are supported."
        )


# ============================================================
# CLEAN EXTRACTED TEXT
# ============================================================

def clean_text(text: str) -> str:

    # Remove excessive spaces
    text = re.sub(r"[ \t]+", " ", text)

    # Reduce excessive blank lines
    text = re.sub(r"\n{3,}", "\n\n", text)

    return text.strip()


def prepare_cv_for_model(text: str) -> str:
    """Remove OCR layout labels and obvious decorative noise before prompting."""

    cleaned_lines = []
    noise_words = {"q", "ty", "ws", "ti", "mat", "-_"}

    for raw_line in text.splitlines():
        line = raw_line.strip()

        if not line:
            cleaned_lines.append("")
            continue

        if re.match(r"^---\s*PAGE\s+\d+\s*---$", line, re.IGNORECASE):
            continue

        if re.match(
            r"^\[(PAGE HEADER|LEFT COLUMN|RIGHT COLUMN).*OCR\]$",
            line,
            re.IGNORECASE,
        ):
            continue

        if line.lower() in noise_words:
            continue

        # Tesseract commonly reads graphical skill-rating circles as long
        # repetitions of O/Q/R/@ and mojibake symbols.
        if re.search(r"(?:OR){3}|(?:OQ){2}|[@Ââ€™®©]{3}", line, re.IGNORECASE):
            continue

        if line.lower() == "e":
            cleaned_lines.append("-")
            continue

        line = re.sub(r"^\$?\d+\s+KEY SKILLS$", "KEY SKILLS", line, flags=re.I)
        line = re.sub(r"^\d+\.\s+SUMMARY$", "SUMMARY", line, flags=re.I)
        line = re.sub(r"\b250kK\+", "250K+", line)
        cleaned_lines.append(line)

    cleaned = "\n".join(cleaned_lines)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


# ============================================================
# LLM PROMPT
# ============================================================

def build_rewrite_prompt(cv_text: str, job_description: str = "") -> str:
    """
    Build a prompt that asks the language model to rewrite the CV directly into
    ATS-friendly text, in a strict, easy-to-parse plain-text layout.

    The output of this prompt is the final CV content. It gets
    parsed by parse_ats_text() below and rendered into a properly
    designed single-column DOCX/PDF - so the exact line shapes
    described here matter, not just the wording.
    """

    jd_block = (
        job_description.strip()
        if job_description.strip()
        else "No job description provided."
    )

    return f"""Rewrite this CV as the final plain-text ATS resume. Output the
resume only: no analysis, instructions, Markdown, or placeholders. Use only
stated facts and omit any missing name or field. Use uppercase section headings,
"Title | Company | Dates" for jobs, and "- " bullets for duties and skills.
Job target: {jd_block}

CV:
{cv_text}"""


# ============================================================
# CALL LOCAL OLLAMA
# ============================================================

def generate_ats_text(
    cv_text: str,
    job_description: str = "",
    retry_feedback: str = "",
) -> str:
    """
    Send the CV to the local Ollama model and return ATS-friendly plain text.
    """

    prompt = build_rewrite_prompt(
        cv_text=prepare_cv_for_model(cv_text),
        job_description=job_description
    )

    if retry_feedback:
        prompt += (
            "\n\nCORRECTION: The previous reply was rejected because "
            f"{retry_feedback} Return the complete resume itself now; do not "
            "return a placeholder or describe what the response should contain."
        )

    print(f"\nSending CV to local Ollama model {MODEL}...")
    print("This may take some time on CPU.\n")

    try:
        response = ollama.generate(
            model=MODEL,
            prompt=prompt,
            options={"temperature": 0.2},
        )
        rewritten = response["response"].strip()
    except Exception as exc:
        raise RuntimeError(
            f"Local Ollama request failed for model {MODEL}: {exc}"
        ) from exc

    if not rewritten:
        raise RuntimeError("Ollama returned no usable text.")

    # Strip accidental Markdown code fences, if the model adds them
    # despite being told not to.

    rewritten = re.sub(
        r"^```(?:text)?\s*",
        "",
        rewritten,
        flags=re.IGNORECASE
    )

    rewritten = re.sub(
        r"\s*```$",
        "",
        rewritten
    )

    return rewritten.strip()


def validate_rewritten_text(rewritten_text: str) -> None:
    """Reject empty, placeholder, or structurally incomplete model output."""

    normalized = rewritten_text.strip()
    lowered = normalized.lower()
    placeholder_markers = (
        "[formatted content]",
        "[resume content]",
        "[resume text]",
        "insert resume",
    )

    if len(normalized) < 200:
        raise RuntimeError(
            f"Model output is too short to be a resume ({len(normalized)} characters)."
        )

    if any(marker in lowered for marker in placeholder_markers):
        raise RuntimeError("Model returned placeholder text instead of a resume.")

    heading_count = sum(
        1
        for line in normalized.splitlines()
        if line.strip().lower().rstrip(":") in KNOWN_HEADINGS
    )

    if heading_count < 2:
        raise RuntimeError(
            "Model output is incomplete: fewer than two resume sections were found."
        )


# ============================================================
# CLEAN UP MODEL OUTPUT
# ============================================================
# Small local models rarely follow formatting instructions 100% of
# the time. Everything in this section defensively normalizes
# common deviations (markdown bold, tab-separated bullets, smart
# bullet characters, stray links, placeholder contact values) BEFORE
# / AFTER the text is parsed into sections, so the final DOCX/PDF
# stays clean even when the model does not fully comply.

BULLET_MARKERS = ("-", "*", "\u2022", "\u2023", "\u25cf", "\u2043")

KNOWN_HEADINGS = {
    "professional summary": "PROFESSIONAL SUMMARY",
    "summary": "PROFESSIONAL SUMMARY",
    "technical skills": "TECHNICAL SKILLS",
    "skills": "TECHNICAL SKILLS",
    "professional experience": "PROFESSIONAL EXPERIENCE",
    "experience": "PROFESSIONAL EXPERIENCE",
    "work experience": "PROFESSIONAL EXPERIENCE",
    "education": "EDUCATION",
    "projects": "PROJECTS",
    "certifications": "CERTIFICATIONS",
    "achievements": "ACHIEVEMENTS",
    "languages": "LANGUAGES",
    "contact": "CONTACT",
    "contact information": "CONTACT",
}

JOB_HEADER_PIPE_RE = re.compile(
    r"^(?P<a>[^|]+)\|(?P<b>[^|]+)\|(?P<c>[^|]+)$"
)

JOB_HEADER_FALLBACK_RE = re.compile(
    r"^(?P<a>.+?)[,\u2013\-]\s*(?P<b>.+?)\s*\((?P<c>[^)]+)\)\s*$"
)

GPA_RE = re.compile(r"^gpa\s*:?\s*(.+)$", re.IGNORECASE)

# A trailing ", GPA: 3.8 / 4.0" (or similar) breaks the entry regexes
# above, since they expect the line to end right after "(...)". Pull
# it off before matching, then re-attach it as its own line.
GPA_SUFFIX_RE = re.compile(
    r",?\s*gpa\s*:?\s*([\d.]+\s*/\s*[\d.]+)\s*$",
    re.IGNORECASE
)

# Conversational filler some local models prepend despite being told
# not to ("Here is the rewritten resume:", "Sure, here's..."). Any
# line matching this is dropped rather than mistaken for the
# candidate's name.
PREAMBLE_RE = re.compile(
    r"^(here'?s|here is|below is|sure[,!]?|certainly[,!]?|"
    r"i'?ve rewritten|the following is)\b",
    re.IGNORECASE
)

# Placeholder values the model sometimes writes for a missing
# contact field instead of just leaving the field out entirely, e.g.
# "LinkedIn: (not present)", "GitHub: N/A", "Website: (no website
# provided)", "LinkedIn: no LinkedIn found". Matched case-insensitively
# after surrounding parens/brackets are stripped.
PLACEHOLDER_VALUE_RE = re.compile(
    r"^\(?("
    r"not present|not provided|not available|not applicable|"
    r"n/?a|none|unknown|not specified|not given|not listed|"
    r"no \w+ (provided|found|given|listed|available)|"
    r"-"
    r")\)?$",
    re.IGNORECASE
)


def _split_gpa_suffix(line: str):
    """
    Split a trailing GPA clause off an entry line.
    Returns (line_without_gpa, gpa_value_or_None).
    """

    match = GPA_SUFFIX_RE.search(line)

    if not match:
        return line, None

    return line[:match.start()].rstrip(", "), match.group(1).strip()


def _strip_markdown(line: str) -> str:

    # **bold** or __bold__ -> bold
    line = re.sub(r"\*\*(.+?)\*\*", r"\1", line)
    line = re.sub(r"__(.+?)__", r"\1", line)

    # remaining stray * or _ used as emphasis markers
    line = line.replace("**", "").replace("__", "")

    # leading markdown heading markers (#, ##, ...)
    line = re.sub(r"^#{1,6}\s*", "", line)

    # markdown links [text](url) -> text (url)
    line = re.sub(r"\[([^\]]+)\]\(([^)]+)\)", r"\1 (\2)", line)

    return line.strip()


def _split_bullet(line: str):
    """
    If the line is a bullet in any common form (-, *, bullet char,
    optionally followed by a tab instead of a space), return
    (True, text-without-marker). Otherwise (False, original-line).
    """

    stripped = line.strip()

    if not stripped:
        return False, line

    if stripped[0] in BULLET_MARKERS:

        rest = stripped[1:]
        rest = rest.lstrip("\t ")

        return True, rest.strip()

    return False, line


def _normalize_line(raw_line: str) -> str:

    line = raw_line.replace("\t", " ").strip()
    line = _strip_markdown(line)

    return line


def _clean_contact_line(contact_line: str) -> str:
    """
    Drop any "Label: value" segment whose value is a placeholder like
    "(not present)", "N/A", "None", etc. - defensive cleanup for when
    the model writes the field instead of leaving it out entirely, as
    instructed in the prompt.

    Also drops a segment with no value at all (e.g. a trailing
    "LinkedIn:" with nothing after it).
    """

    if not contact_line:
        return contact_line

    segments = [s.strip() for s in contact_line.split("|")]
    kept = []

    for segment in segments:

        if not segment:
            continue

        label, sep, value = segment.partition(":")
        value = value.strip().strip("()[]").strip()

        if not sep:
            # No "Label: value" shape at all - keep as-is, it may be
            # a bare value the model added outside the expected
            # format (e.g. just a phone number with no label).
            kept.append(segment)
            continue

        if not value or PLACEHOLDER_VALUE_RE.match(value):
            continue

        kept.append(segment)

    return " | ".join(kept)


def _looks_like_contact(line: str) -> bool:
    """Return whether a header line contains contact-style fields."""

    lowered = line.lower()
    return any(
        token in lowered
        for token in (
            "@", "phone", "email", "http", "linkedin", "github", "|", "+"
        )
    )


def remove_unsupported_name(rewritten_text: str, source_text: str) -> str:
    """Remove a model-generated header name not evidenced by the source CV."""

    lines = rewritten_text.splitlines()
    index = 0

    while index < len(lines):
        candidate = _normalize_line(lines[index])

        if candidate and not PREAMBLE_RE.match(candidate):
            break

        index += 1

    if index >= len(lines):
        return rewritten_text

    candidate = _normalize_line(lines[index])

    if (
        candidate.lower().rstrip(":") in KNOWN_HEADINGS
        or _looks_like_contact(candidate)
    ):
        return rewritten_text

    candidate_tokens = {
        token.lower()
        for token in re.findall(r"[A-Za-z]+", candidate)
        if len(token) > 1
    }
    source_tokens = {
        token.lower()
        for token in re.findall(r"[A-Za-z]+", source_text)
    }

    if candidate_tokens and candidate_tokens.issubset(source_tokens):
        return rewritten_text

    print(
        "WARNING: Removed an unsupported candidate name generated by the model: "
        f"{candidate}"
    )
    del lines[index]
    return "\n".join(lines).strip()


def parse_ats_text(rewritten_text: str) -> dict:
    """
    Parse the model's rewritten CV text into a structure the DOCX/PDF
    writers can render with real design (tab-aligned dates,
    underlined headings, etc.).

    Deliberately forgiving: if a line does not match the requested
    shape, it degrades to a plain paragraph instead of being dropped,
    so nothing from the CV silently disappears.
    """

    lines = [
        _normalize_line(l)
        for l in rewritten_text.splitlines()
    ]

    result = {
        "name": "",
        "contact": "",
        "sections": [],
        # each section: {"heading": str, "items": [...]}
        # item kinds:
        #   {"type": "text", "text": ...}
        #   {"type": "bullet", "text": ...}
        #   {"type": "entry", "title": ..., "sub": ..., "dates": ...,
        #    "bullets": [...], "extra": [...]}
    }

    i = 0
    n = len(lines)

    # ---------------------------------------
    # Header: name + contact line
    # ---------------------------------------

    while i < n and (not lines[i] or PREAMBLE_RE.match(lines[i])):
        i += 1

    if i < n and lines[i].lower().rstrip(":") not in KNOWN_HEADINGS:
        if not _looks_like_contact(lines[i]):
            result["name"] = lines[i]
            i += 1

    while i < n and not lines[i]:
        i += 1

    if i < n and lines[i].lower().rstrip(":") not in KNOWN_HEADINGS:

        looks_like_contact = _looks_like_contact(lines[i])

        if looks_like_contact:
            result["contact"] = lines[i]
            i += 1

    # ---------------------------------------
    # Sections
    # ---------------------------------------

    current_section = None

    def start_section(heading):
        nonlocal current_section
        current_section = {"heading": heading, "items": []}
        result["sections"].append(current_section)

    while i < n:

        line = lines[i]
        i += 1

        if not line:
            continue

        heading_key = line.lower().rstrip(":")

        if heading_key in KNOWN_HEADINGS:
            start_section(KNOWN_HEADINGS[heading_key])
            continue

        if current_section is None:
            # Content before any recognized heading - keep it, under
            # an implicit summary section, rather than lose it.
            start_section("PROFESSIONAL SUMMARY")

        is_bullet, bullet_text = _split_bullet(line)

        if current_section["heading"] == "PROFESSIONAL EXPERIENCE":

            if not is_bullet:

                line_for_match, gpa_value = _split_gpa_suffix(line)

                match = (
                    JOB_HEADER_PIPE_RE.match(line_for_match)
                    or JOB_HEADER_FALLBACK_RE.match(line_for_match)
                )

                if match:

                    entry = {
                        "type": "entry",
                        "title": match.group("a").strip(),
                        "sub": match.group("b").strip(),
                        "dates": match.group("c").strip(),
                        "bullets": [],
                    }

                    if gpa_value:
                        entry["bullets"].append(f"GPA: {gpa_value}")

                    current_section["items"].append(entry)

                    continue

                # Not a recognized job header - if there is an open
                # entry, treat it as a continuation line under it;
                # otherwise it is a stray paragraph.

                if (
                    current_section["items"]
                    and current_section["items"][-1]["type"] == "entry"
                ):
                    current_section["items"][-1]["bullets"].append(line)
                else:
                    current_section["items"].append(
                        {"type": "text", "text": line}
                    )

                continue

            # It is a bullet - attach to the most recent entry if one
            # is open, otherwise keep as a loose bullet.

            if (
                current_section["items"]
                and current_section["items"][-1]["type"] == "entry"
            ):
                current_section["items"][-1]["bullets"].append(
                    bullet_text
                )
            else:
                current_section["items"].append(
                    {"type": "bullet", "text": bullet_text}
                )

            continue

        if current_section["heading"] == "EDUCATION":

            if not is_bullet:

                line_for_match, gpa_value = _split_gpa_suffix(line)

                match = (
                    JOB_HEADER_PIPE_RE.match(line_for_match)
                    or JOB_HEADER_FALLBACK_RE.match(line_for_match)
                )

                if match:

                    entry = {
                        "type": "entry",
                        "title": match.group("a").strip(),
                        "sub": match.group("b").strip(),
                        "dates": match.group("c").strip(),
                        "bullets": [],
                    }

                    if gpa_value:
                        entry["bullets"].append(f"GPA: {gpa_value}")

                    current_section["items"].append(entry)

                    continue

                gpa_match = GPA_RE.match(line)

                if (
                    gpa_match
                    and current_section["items"]
                    and current_section["items"][-1]["type"] == "entry"
                ):
                    current_section["items"][-1]["bullets"].append(
                        f"GPA: {gpa_match.group(1).strip()}"
                    )
                    continue

                if (
                    current_section["items"]
                    and current_section["items"][-1]["type"] == "entry"
                ):
                    current_section["items"][-1]["bullets"].append(line)
                else:
                    current_section["items"].append(
                        {"type": "text", "text": line}
                    )

                continue

            if (
                current_section["items"]
                and current_section["items"][-1]["type"] == "entry"
            ):
                current_section["items"][-1]["bullets"].append(
                    bullet_text
                )
            else:
                current_section["items"].append(
                    {"type": "bullet", "text": bullet_text}
                )

            continue

        # Any other section (skills, summary, projects, etc.)

        if is_bullet:
            current_section["items"].append(
                {"type": "bullet", "text": bullet_text}
            )
        else:
            current_section["items"].append(
                {"type": "text", "text": line}
            )

    # ---------------------------------------
    # Fold a stray CONTACT section into the header
    # ---------------------------------------
    # Despite the prompt asking for contact info on line 2, some
    # models put it in its own section instead. If the header ended
    # up without contact info, pull it from there and drop the
    # section so it isn't shown twice.

    if not result["contact"]:

        for idx, sect in enumerate(result["sections"]):

            if sect["heading"] != "CONTACT":
                continue

            pieces = [
                item["text"]
                for item in sect["items"]
                if item["type"] in ("text", "bullet") and item["text"]
            ]

            if pieces:
                result["contact"] = " | ".join(pieces)

            del result["sections"][idx]
            break

    # ---------------------------------------
    # Strip placeholder contact values
    # ---------------------------------------
    # Defensive net for when the model writes "LinkedIn: (not
    # present)" / "GitHub: N/A" etc. instead of leaving the field out,
    # regardless of whether the contact line came from line 2 above
    # or from a folded CONTACT section.

    result["contact"] = _clean_contact_line(result["contact"])

    return result


# ============================================================
# PLAIN, DESIGNED ATS-FRIENDLY DOCX OUTPUT
# ============================================================
# Single column throughout, no tables, no text boxes, no images -
# the things that confuse ATS parsers. Visual polish (underlined
# section headings, right-aligned dates) comes from paragraph
# borders and a right tab stop, both of which are plain-paragraph
# features that ATS software reads as ordinary text.

HEADING_COLOR = RGBColor(0x1F, 0x3B, 0x57)


def _set_bottom_border(paragraph, color: str = "1F3B57", size: int = 6):
    """
    Add a thin bottom border to a paragraph - used as the underline
    under the name/contact block and under each section heading.
    """

    p_pr = paragraph._p.get_or_add_pPr()

    p_bdr = OxmlElement("w:pBdr")

    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), color)

    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def _right_tab_stop(doc) -> Pt:

    section = doc.sections[0]

    return (
        section.page_width
        - section.left_margin
        - section.right_margin
    )


def _add_name(doc, name: str):

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = paragraph.add_run(name)
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = HEADING_COLOR

    paragraph.space_after = Pt(2)


def _add_contact(doc, contact_line: str):

    display = contact_line.replace("|", "  \u2022  ")

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = paragraph.add_run(display)
    run.font.size = Pt(10)

    paragraph.space_after = Pt(10)

    _set_bottom_border(paragraph)


def _add_section_heading(doc, text: str):

    paragraph = doc.add_paragraph()
    paragraph.space_before = Pt(14)
    paragraph.space_after = Pt(4)

    run = paragraph.add_run(text)
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = HEADING_COLOR

    _set_bottom_border(paragraph)


def _add_entry_header(doc, title: str, sub: str, dates: str, tab_pos):

    title_paragraph = doc.add_paragraph()
    title_paragraph.space_before = Pt(6)
    title_paragraph.space_after = Pt(0)

    title_run = title_paragraph.add_run(title)
    title_run.bold = True
    title_run.font.size = Pt(11)

    sub_paragraph = doc.add_paragraph()
    sub_paragraph.space_after = Pt(2)

    tab_stops = sub_paragraph.paragraph_format.tab_stops
    tab_stops.add_tab_stop(tab_pos, WD_TAB_ALIGNMENT.RIGHT)

    sub_run = sub_paragraph.add_run(sub)
    sub_run.bold = True
    sub_run.font.size = Pt(10)

    if dates:

        sub_paragraph.add_run("\t")

        dates_run = sub_paragraph.add_run(dates)
        dates_run.font.size = Pt(10)
        dates_run.italic = True


def _add_bullet(doc, text: str):

    doc.add_paragraph(text, style="List Bullet")


def _add_plain(doc, text: str):

    doc.add_paragraph(text)


def write_ats_docx(rewritten_text: str, original_file: Path) -> Path:
    """
    Parse the model's rewritten CV text and render it into a properly
    designed, single-column, ATS-friendly DOCX: centered name and
    contact block, underlined section headings, and right-aligned
    dates next to each job/education entry.
    """

    data = parse_ats_text(rewritten_text)

    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    section = doc.sections[0]
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    section.top_margin = Inches(0.6)
    section.bottom_margin = Inches(0.6)

    tab_pos = _right_tab_stop(doc)

    if data["name"]:
        _add_name(doc, data["name"])

    if data["contact"]:
        _add_contact(doc, data["contact"])

    for sect in data["sections"]:

        _add_section_heading(doc, sect["heading"])

        for item in sect["items"]:

            if item["type"] == "entry":

                _add_entry_header(
                    doc,
                    item["title"],
                    item["sub"],
                    item["dates"],
                    tab_pos
                )

                for bullet_text in item["bullets"]:
                    _add_bullet(doc, bullet_text)

            elif item["type"] == "bullet":
                _add_bullet(doc, item["text"])

            else:
                _add_plain(doc, item["text"])

    output_path = (
        OUTPUT_DIR /
        f"{original_file.stem}_ATS_Friendly.docx"
    )

    doc.save(output_path)

    return output_path


# ============================================================
# PLAIN, DESIGNED ATS-FRIENDLY PDF OUTPUT
# ============================================================
# Same idea as write_ats_docx: single column, no tables, no text
# boxes, no images. Built with reportlab (pure Python, no Word or
# LibreOffice required on the machine running this script) so the
# final deliverable can be a PDF directly instead of a DOCX.
#
# Requires: pip install reportlab

PDF_PAGE_WIDTH, PDF_PAGE_HEIGHT = LETTER

PDF_MARGIN_LEFT = 0.7 * inch
PDF_MARGIN_RIGHT = 0.7 * inch
PDF_MARGIN_TOP = 0.6 * inch
PDF_MARGIN_BOTTOM = 0.6 * inch

PDF_CONTENT_WIDTH = (
    PDF_PAGE_WIDTH - PDF_MARGIN_LEFT - PDF_MARGIN_RIGHT
)

PDF_HEADING_COLOR = pdf_colors.HexColor("#1F3B57")

FONT_REGULAR = "Helvetica"
FONT_BOLD = "Helvetica-Bold"
FONT_ITALIC = "Helvetica-Oblique"


class _PdfCursor:
    """
    Tracks the current write position on the page and starts a new
    page automatically when content would run off the bottom.
    """

    def __init__(self, canvas_obj):
        self.c = canvas_obj
        self.y = PDF_PAGE_HEIGHT - PDF_MARGIN_TOP

    def new_page(self):
        self.c.showPage()
        self.y = PDF_PAGE_HEIGHT - PDF_MARGIN_TOP

    def ensure_space(self, height: float):

        if self.y - height < PDF_MARGIN_BOTTOM:
            self.new_page()


def _pdf_wrap_lines(text: str, font: str, size: int, max_width: float):
    """
    Break text into lines that fit max_width at the given font/size,
    the way a word processor would - reportlab draws single lines
    only, so this has to be done by hand.
    """

    words = text.split()

    if not words:
        return [""]

    lines = []
    current = ""

    for word in words:

        candidate = f"{current} {word}".strip()

        if stringWidth(candidate, font, size) <= max_width:
            current = candidate
        else:
            if current:
                lines.append(current)
            current = word

    if current:
        lines.append(current)

    return lines or [""]


def _pdf_draw_line(
    cur: _PdfCursor,
    text: str,
    font: str,
    size: int,
    color=pdf_colors.black,
    align: str = "left",
    x=None,
):

    cur.ensure_space(size * 1.4)
    cur.y -= size * 1.2

    cur.c.setFont(font, size)
    cur.c.setFillColor(color)

    if align == "center":
        cur.c.drawCentredString(
            PDF_PAGE_WIDTH / 2, cur.y, text
        )
    elif align == "right":
        cur.c.drawRightString(
            x if x is not None else PDF_PAGE_WIDTH - PDF_MARGIN_RIGHT,
            cur.y,
            text
        )
    else:
        cur.c.drawString(
            x if x is not None else PDF_MARGIN_LEFT,
            cur.y,
            text
        )


def _pdf_add_name(cur: _PdfCursor, name: str):

    _pdf_draw_line(
        cur, name, FONT_BOLD, 20,
        color=PDF_HEADING_COLOR, align="center"
    )

    cur.y -= 4


def _pdf_add_contact(cur: _PdfCursor, contact_line: str):

    display = contact_line.replace("|", "   \u2022   ")

    lines = _pdf_wrap_lines(
        display, FONT_REGULAR, 10, PDF_CONTENT_WIDTH
    )

    for line in lines:
        _pdf_draw_line(
            cur, line, FONT_REGULAR, 10, align="center"
        )

    cur.y -= 4

    cur.c.setStrokeColor(PDF_HEADING_COLOR)
    cur.c.setLineWidth(0.75)
    cur.c.line(
        PDF_MARGIN_LEFT, cur.y,
        PDF_PAGE_WIDTH - PDF_MARGIN_RIGHT, cur.y
    )

    cur.y -= 10


def _pdf_add_section_heading(cur: _PdfCursor, text: str):

    # Reserve room for the heading itself PLUS roughly one line of
    # whatever follows it, so a heading never ends up alone at the
    # bottom of a page with its content pushed to the next one.
    cur.ensure_space(48)
    cur.y -= 10

    _pdf_draw_line(
        cur, text, FONT_BOLD, 12,
        color=PDF_HEADING_COLOR, align="left"
    )

    cur.c.setStrokeColor(PDF_HEADING_COLOR)
    cur.c.setLineWidth(0.75)
    cur.c.line(
        PDF_MARGIN_LEFT, cur.y - 2,
        PDF_PAGE_WIDTH - PDF_MARGIN_RIGHT, cur.y - 2
    )

    cur.y -= 6


def _pdf_add_entry_header(
    cur: _PdfCursor, title: str, sub: str, dates: str
):

    _pdf_draw_line(cur, title, FONT_BOLD, 11)

    cur.ensure_space(14)
    cur.y -= 12

    cur.c.setFont(FONT_BOLD, 10)
    cur.c.setFillColor(pdf_colors.black)
    cur.c.drawString(PDF_MARGIN_LEFT, cur.y, sub)

    if dates:
        cur.c.setFont(FONT_ITALIC, 10)
        cur.c.drawRightString(
            PDF_PAGE_WIDTH - PDF_MARGIN_RIGHT, cur.y, dates
        )

    cur.y -= 4


def _pdf_add_bullet(cur: _PdfCursor, text: str):

    bullet_indent = 0.18 * inch

    lines = _pdf_wrap_lines(
        text, FONT_REGULAR, 11,
        PDF_CONTENT_WIDTH - bullet_indent
    )

    for index, line in enumerate(lines):

        cur.ensure_space(14)
        cur.y -= 13

        cur.c.setFont(FONT_REGULAR, 11)
        cur.c.setFillColor(pdf_colors.black)

        if index == 0:
            cur.c.drawString(PDF_MARGIN_LEFT, cur.y, "\u2022")

        cur.c.drawString(
            PDF_MARGIN_LEFT + bullet_indent, cur.y, line
        )


def _pdf_add_paragraph(cur: _PdfCursor, text: str):

    lines = _pdf_wrap_lines(
        text, FONT_REGULAR, 11, PDF_CONTENT_WIDTH
    )

    for line in lines:
        _pdf_draw_line(cur, line, FONT_REGULAR, 11)


def write_ats_pdf(rewritten_text: str, original_file: Path) -> Path:
    """
    Parse the model's rewritten CV text and render it directly into a
    single-column, ATS-friendly PDF - centered name and contact
    block, underlined section headings, right-aligned dates next to
    each job/education entry. Same layout as write_ats_docx, but no
    Word/LibreOffice dependency: reportlab writes the PDF bytes
    itself.
    """

    data = parse_ats_text(rewritten_text)

    if not data["sections"]:
        raise ValueError("Refusing to create a PDF with no resume sections.")

    output_path = (
        OUTPUT_DIR /
        f"{original_file.stem}_ATS_Friendly.pdf"
    )

    temp_output_path = output_path.with_suffix(".tmp.pdf")

    c = pdf_canvas.Canvas(str(temp_output_path), pagesize=LETTER)
    cur = _PdfCursor(c)

    if data["name"]:
        _pdf_add_name(cur, data["name"])

    if data["contact"]:
        _pdf_add_contact(cur, data["contact"])

    for sect in data["sections"]:

        _pdf_add_section_heading(cur, sect["heading"])

        for item in sect["items"]:

            if item["type"] == "entry":

                cur.ensure_space(55)

                _pdf_add_entry_header(
                    cur,
                    item["title"],
                    item["sub"],
                    item["dates"]
                )

                for bullet_text in item["bullets"]:
                    _pdf_add_bullet(cur, bullet_text)

            elif item["type"] == "bullet":
                _pdf_add_bullet(cur, item["text"])

            else:
                _pdf_add_paragraph(cur, item["text"])

    c.save()

    try:
        with pymupdf.open(temp_output_path) as rendered:
            page_count = len(rendered)
            text_length = sum(len(page.get_text().strip()) for page in rendered)

        if page_count < 1 or text_length < 1:
            raise ValueError(
                "Generated PDF validation failed: the document has no readable pages."
            )

        temp_output_path.replace(output_path)
    except Exception:
        temp_output_path.unlink(missing_ok=True)
        raise

    return output_path


def main():

    print("=" * 60)
    print("LOCAL ATS CV CONVERTER")
    print("=" * 60)

    # ---------------------------------------
    # Find CV files
    # ---------------------------------------

    cv_files = [
        file
        for file in INPUT_DIR.iterdir()
        if file.suffix.lower()
        in [".pdf", ".docx"]
    ]

    if not cv_files:

        print(
            "\nNo PDF or DOCX files found."
        )

        print(
            f"Put your CV inside:\n{INPUT_DIR}"
        )

        return

    print("\nAvailable CV files:")

    for index, file in enumerate(
        cv_files,
        start=1
    ):

        print(
            f"{index}. {file.name}"
        )

    # ---------------------------------------
    # Select CV
    # ---------------------------------------

    while True:

        try:

            choice = int(
                input(
                    "\nSelect CV number: "
                )
            )

            if 1 <= choice <= len(cv_files):
                break

        except ValueError:
            pass

        print(
            "Invalid selection."
        )

    cv_file = cv_files[
        choice - 1
    ]

    print(
        f"\nSelected: {cv_file.name}"
    )

    # ---------------------------------------
    # Optional JD
    # ---------------------------------------

    use_jd = input(
        "\nDo you want to provide a Job Description? "
        "(y/n): "
    ).strip().lower()

    job_description = ""

    if use_jd == "y":

        print(
            "\nPaste the Job Description."
        )

        print(
            "Type END on a new line when finished.\n"
        )

        lines = []

        while True:

            line = input()

            if line.strip() == "END":
                break

            lines.append(line)

        job_description = "\n".join(lines)

    # ---------------------------------------
    # Read CV
    # ---------------------------------------

    cv_text = read_cv(cv_file)

    cv_text = clean_text(cv_text)

    if not cv_text:

        print(
            "\nERROR: Could not extract text."
        )

        print(
            "If this is an image/scanned PDF, "
            "you will need an OCR step."
        )

        return

    # Save extracted text for debugging

    extracted_file = (
        OUTPUT_DIR /
        f"{cv_file.stem}_extracted.txt"
    )

    extracted_file.write_text(
        cv_text,
        encoding="utf-8"
    )

    print(
        f"\nExtracted text saved:\n{extracted_file}"
    )

    # ---------------------------------------
    # Language model: rewrite into ATS-friendly text
    # ---------------------------------------

    try:
        rewritten_text = generate_ats_text(
            cv_text,
            job_description
        )
    except RuntimeError as exc:
        if "no usable text" not in str(exc).lower():
            print(f"\nERROR: {exc}")
            return

        print(f"\nModel response rejected: {exc}")
        print("Retrying the chat API once with corrective instructions...")

        try:
            rewritten_text = generate_ats_text(
                cv_text,
                job_description,
                retry_feedback="the prior response contained reasoning only",
            )
        except RuntimeError as retry_exc:
            print(f"\nERROR after retry: {retry_exc}")
            print("The existing ATS-friendly PDF was not overwritten.")
            return

    # Always keep a copy of what the model actually returned, useful if
    # the DOCX/PDF step below needs debugging.

    raw_file = (
        OUTPUT_DIR /
        f"{cv_file.stem}_model_output.txt"
    )

    raw_file.write_text(
        rewritten_text,
        encoding="utf-8"
    )

    if not rewritten_text:

        print(
            "\nERROR: The language model returned an empty response."
        )

        print(
            f"Nothing to write. Raw output saved to:\n{raw_file}"
        )

        return

    try:
        validate_rewritten_text(rewritten_text)
    except RuntimeError as exc:
        print(f"\nModel response rejected: {exc}")
        print("Retrying the chat API once with corrective instructions...")

        try:
            rewritten_text = generate_ats_text(
                cv_text,
                job_description,
                retry_feedback=str(exc),
            )
        except RuntimeError as retry_exc:
            print(f"\nERROR: {retry_exc}")
            print("The existing ATS-friendly PDF was not overwritten.")
            return

        raw_file.write_text(
            rewritten_text,
            encoding="utf-8"
        )

        try:
            validate_rewritten_text(rewritten_text)
        except RuntimeError as retry_exc:
            print(f"\nERROR after retry: {retry_exc}")
            print("The existing ATS-friendly PDF was not overwritten.")
            return

    rewritten_text = remove_unsupported_name(
        rewritten_text,
        cv_text,
    )

    # ---------------------------------------
    # Write plain ATS-friendly PDF
    # ---------------------------------------

    pdf_path = write_ats_pdf(
        rewritten_text,
        cv_file
    )

    print("\n" + "=" * 60)
    print("DONE")
    print("=" * 60)

    print(
        f"\nATS-friendly CV saved to:\n{pdf_path}"
    )


if __name__ == "__main__":
    main()
