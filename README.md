# CV ATS Converter

A local, privacy-friendly tool that rewrites a candidate's CV (PDF or DOCX)
into a clean, single-column, ATS-friendly resume — optionally tailored to a
target job description — and renders it as a polished PDF.

Everything runs locally. The CV never leaves your machine: text extraction,
rewriting (via a local Llama model through [Ollama](https://ollama.com)),
and PDF generation all happen on-device.

---

## How it works

```
 input/*.pdf, *.docx
         │
         ▼
 ┌───────────────────┐
 │  Text extraction   │  PyMuPDF (PDF) / python-docx (DOCX)
 │                    │  Detects two-column layouts and reads them
 │                    │  column-by-column so content stays in order.
 └────────┬───────────┘
          ▼
 ┌───────────────────┐
 │  LLM rewrite       │  Llama 3.1 8B (via Ollama, local inference)
 │                    │  Rewrites the CV into a strict plain-text
 │                    │  format, optionally aligned to a job
 │                    │  description. Never invents facts.
 └────────┬───────────┘
          ▼
 ┌───────────────────┐
 │  Parse + clean     │  Normalizes markdown, bullets, headings, and
 │                    │  strips placeholder text the model may add
 │                    │  for missing contact fields.
 └────────┬───────────┘
          ▼
 ┌───────────────────┐
 │  PDF rendering     │  ReportLab draws a single-column, ATS-safe
 │                    │  PDF: no tables, no text boxes, no images.
 └────────┬───────────┘
          ▼
 output/*_ATS_Friendly.pdf
```

## Project structure

```
.
├── cv_ats_pdf.py        # Entire pipeline: extraction → LLM rewrite → PDF
├── requirements.txt
├── input/                # Drop your source CV here (.pdf or .docx)
├── output/               # Generated PDF + debug artifacts land here
└── README.md
```

`input/` and `output/` are tracked as empty folders in git (via
`.gitkeep`); their contents are gitignored, since CVs are personal data
that should never be committed.

## Prerequisites

- Python 3.10+
- [Ollama](https://ollama.com) installed and running locally, with the
  `llama3.1:8b` model pulled:
  ```bash
  ollama pull llama3.1:8b
  ```

## Setup

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

1. Place your CV (`.pdf` or `.docx`) in `input/`.
2. Run:
   ```bash
   python cv_ats_pdf.py
   ```
3. Select the CV from the list.
4. Optionally paste a target job description (type `END` on its own line
   to finish).
5. Find the result in `output/`:
   - `<name>_ATS_Friendly.pdf` — the final resume
   - `<name>_extracted.txt` — raw extracted text (debugging)
   - `<name>_llama_output.txt` — raw model output before parsing (debugging)

## Design notes

- **Factual accuracy is enforced in the prompt**, not just requested: the
  model is explicitly told never to invent numbers, titles, dates, or
  skills, and job-description keywords are only added to Skills if the
  original CV already demonstrates them.
- **Output parsing is deliberately forgiving.** Local 8B models don't
  follow formatting instructions perfectly. `parse_ats_text()` degrades
  gracefully — unrecognized lines become plain paragraphs instead of being
  dropped — so nothing from the original CV silently disappears.
- **Missing contact fields are omitted, never shown as placeholders.**
  If the source CV has no LinkedIn or GitHub, the final PDF simply won't
  mention it — the parser strips any placeholder text (`N/A`,
  `(not present)`, `(no LinkedIn provided)`, etc.) the model adds despite
  being told not to.
- **ATS-safe by construction.** The PDF/DOCX renderers use no tables, text
  boxes, or images — just plain paragraphs, borders, and tab stops, which
  ATS parsers read reliably as ordinary text.

## Roadmap ideas

- OCR fallback for scanned/image-only PDFs
- DOCX output alongside PDF
- Swap `ollama.generate` for a config-driven model choice (local or API)

## License

Private project — add a license here if you plan to share this publicly.
