---
name: local-only-file-analysis
description: Analyse spreadsheets, Word/PowerPoint documents, PDFs, images and text files strictly on the local machine, with mandatory masking of personal data (names, phone numbers, addresses, national ID and passport numbers, bank cards, e-mail, plates, account IDs) in every mode, and with zero upload to cloud documents, cloud drives, knowledge bases, online converters, OCR services or external search. Use whenever a user hands over local files for statistics, extraction, profiling or summarisation under a confidentiality constraint — triggers include "local only", "offline", "on-premise", "data must not leave the machine", "no cloud", "confidential", "sensitive", "redact", "anonymise", plus 本地分析、仅本地处理、不上传云端、不联网、涉密、保密、脱敏、隐私数据、原文件不能外发.
---

# Local-Only File Analysis

## Goal

Analyse tabular, document, PDF and image files without sending the file or its contents to any cloud service, and mask personal data — names, contact details, addresses, identity numbers — in **every** mode. Source files are opened read-only on the local machine, extraction and statistics run through local scripts, and every artefact stays inside a local working directory.

## Step 1 — Ask for the local path first (mandatory, never skipped)

The first sentence after this skill activates must ask for a local path, not start analysing. Use the wording below, adapted to context, keeping four points: nothing is uploaded, an absolute path is needed, how to copy a path, and the optional file picker.

> I will analyse this file on this machine only — nothing goes to any cloud service, and names, phone numbers, addresses and ID numbers are masked in the report.
> Please give me the **absolute local path**, e.g. `/Users/you/Documents/ledger.xlsx` or `D:\data\ledger.xlsx`; separate multiple files with new lines.
> To copy a path: right-click the file → "Copy as path" (Windows), or drag the file into a terminal / use ⌥⌘C in Finder (macOS).
> Or just say "pick the files for me" and I will open a local file dialog.

Follow-through:

- User says "pick / show a dialog" → run `python "<skill_dir>/scripts/pick_path.py"` (multi-select), or with `--single`. Use the printed paths as input. Exit code 2 means no graphical session — fall back to asking the user to paste paths.
- User says "it's attached" but only an attachment card exists, no path → handle per the next section. Do not silently read the attachment.
- After receiving paths, validate: path exists, extension supported, sensible file count (group by type above ~10 files).

### When the user already dropped an attachment

State the risk before accepting anything:

> One thing to flag: a file dropped into the chat may already be stored on the platform's servers, so that path is no longer purely local.
> If this file is confidential or contains personal data, send me its local path instead and I will read it from disk — nothing leaves this machine.

If the user still wants the attachment, proceed, but label the delivery honestly: "input came from a chat attachment, which may have reached the platform's servers" — never call an attachment "fully local".

## Step 2 — Let the user pick a mode

Two options; the user must choose. If unspecified, use **Normal** and state the mode in the first line of the answer.

| | **Normal** (default) | **Strict** |
|---|---|---|
| Report content | Structure + statistics + masked value preview (first N rows, body excerpts) | Aggregate metrics only: rows/columns, types, non-null rate, distinct count, range, mean |
| Raw text | Masked text enters the report, visible to the model | Never appears; the model sees metrics only |
| Suits | General internal data, semantic interpretation, written conclusions | Customer records, payroll, medical data, contracts, classified material |
| Cost | Masked content enters the inference context | No semantic judgement possible, statistics-only answers |

Wording:

> Choose a mode: **① Normal** — structure, statistics and a masked preview so I can read the content and draw conclusions; **② Strict** — metrics only, no original text anywhere in the report. Default is Normal.

## Mandatory masking (all modes, cannot be disabled)

Masking is not an option, it is default behaviour. The script has no off switch, and `--mode strict` does not mean "masking is optional" — strict stacks "no original text" on top of masking.

| Category | Covered fields | Result |
|---|---|---|
| Name | name, account holder, contact, legal representative, owner, handler, account manager, applicant, payee, cardholder, shareholder | `[NAME#1]` |
| Contact | mobile, landline (with area code), fax, e-mail | `[PHONE#1]` / `[EMAIL#1]` |
| Address | address, residence, mailing address, registered address, domicile | Province/city kept, rest becomes `[ADDRESS#1]`; `--mask-region` hides the region too |
| Identity | national ID (18/15 digit), passport, travel permit, officer ID, unified social credit code | `[ID#1]` |
| Financial | bank card, account number, card number | `[ACCOUNT#1]` |
| Other identifiers | licence plate, customer no., employee no., member no., student no., social-security no., policy no. | `[PLATE#1]` / `[IDNO#1]` |

Three implementation principles, quotable in the delivery note:

1. **Stable pseudonyms.** The same source value always maps to the same pseudonym within one run. "Zhang appears 3 times" becomes "`[NAME#1]` appears 3 times" — counting, de-duplication and joins still work, identity does not survive.
2. **No mapping on disk.** The value-to-pseudonym map lives in memory only. Neither `report.md` nor `manifest.json` contains it; only hit counts are written.
3. **Heuristics for free text.** Structured tables are classified by column header and masked column-wide. Free text in bodies, PDFs and slides relies on regex plus a surname lexicon ("contact: Zhang San", "Mr. Zhang"). Heuristics miss things — especially isolated names with no contextual cue. Say so plainly; never claim "everything is detected".

Keeping province/city is deliberate: it supports region-level analysis without exposing a street address. Use `--mask-region` when the region itself is sensitive.

## Step 3 — Extract locally

```bash
python "<skill_dir>/scripts/local_extract.py" "/path/to/ledger.xlsx" "/path/to/report.docx" \
  --workspace ".local_analysis"
```

Add `--mode strict` for strict mode; `--mask-region` to hide regions; `--redact "EMP\d{6}"` (repeatable) for custom patterns.

Outputs (all inside `--workspace`):

- `report.md` — masked structural profile + statistics + preview; the only input for the rest of the analysis
- `manifest.json` — file list, SHA-256, mode, `masking.enabled=true`, mask hit counts, `network_used: false`

The console prints per-file type, size, masked fields, hit counts and "network requests: 0". Read `report.md` with the local file reader afterwards; do not substitute screenshots or online OCR.

## Step 4 — Compute locally, never mentally

Sums, ratios, distinct counts, sorting, grouping, period-over-period changes: always computed by a local script. Without seeing the full data, model estimates are almost always wrong. Reuse `parse_xlsx` / `parse_delimited` / `Masker` from `local_extract.py` in a throwaway script, compute, then write the numbers into the report. Pseudonyms can be grouped and counted directly — no need to reverse them.

## Step 5 — Produce an audit sheet

```bash
python "<skill_dir>/scripts/local_audit.py" ".local_analysis" --note "processed locally, nothing uploaded"
```

The audit sheet has five parts: source-file integrity (hash re-check proving read-only, no copies), **masking execution** (switch state, categories, hits in this run), artefact inventory, line-by-line red-line check, conclusion. Exit code 0 = pass, 1 = a source file changed during processing.

## Step 6 — Report

Conclusion first, data second, one closing statement:

> This analysis ran entirely on this machine with no cloud service involved; source files were read-only and unchanged, and the report masks names, phone numbers, addresses and ID numbers (stable pseudonyms). Artefacts are in `.local_analysis/`, audit sheet in `audit.md`.

## Red lines (never triggered)

| Category | Forbidden | Typical tools |
|---|---|---|
| Cloud documents | Create, import or "edit" an online sheet/document | Tencent Docs, Kingsoft Docs, Feishu Docs connectors |
| Cloud drives / object storage | Upload files or artefacts | Tencent Drive, WeDrive, cloud file storage |
| Knowledge bases | Write into a library, ingest URLs | library upload, ima knowledge base |
| External search | Use source text or fragments as a query | web search, page scraping |
| Online conversion / OCR | Online PDF conversion, online unzip, online image recognition | any online tool, image-understanding API |
| Outbound sending | Post files or summaries to chat, mail, bots | WeCom messages, e-mail, meeting-notes upload |

One exception only: the user explicitly says "this one may be uploaded / sent". Even then, restate it and get a second confirmation before acting.

## Local path per file type

| Type | Preferred local route | Notes |
|---|---|---|
| `.xlsx` `.xlsm` | `local_extract.py` (zip + XML directly) | Multi-sheet, column-level masking; dates may appear as serial numbers |
| `.csv` `.tsv` `.txt` `.md` `.json` `.html` | `local_extract.py` | Detects UTF-8 / GB18030 and delimiter; CSV masked by header |
| `.docx` `.pptx` | `local_extract.py` (zip + XML directly) | Paragraphs, heading outline, tables, per-slide bullets; body through free-text masking |
| `.pdf` | `local_extract.py` | Uses a local `workbuddy-pdf` CLI when present, built-in zlib fallback otherwise |
| `.png` `.jpg` `.gif` `.bmp` `.webp` | `local_extract.py` | Format, dimensions, hash only; local OCR only if tesseract exists |
| `.xls` `.doc` `.ppt` legacy binary | Ask the user to re-save locally, or open in a local editor | Never send to an online converter to extract content |
| Scanned PDF / text inside images | Local OCR, or the user dictates the content | If no OCR, state the gap honestly — do not route through the cloud |

Images in particular: without tesseract, treat them as metadata only and say "image content was not recognised", offering two local routes (install OCR and rerun / user describes it). Do not "take a look" with a vision model — that pushes the image into the inference path and contradicts the user's request.

## Degradation rules

1. Local script can parse it → use the script's output.
2. No extractable text (scan, legacy format, encrypted) → explain why, offer local alternatives (re-save locally, local OCR, user dictation).
3. User pushes for "results now" → return to the mode question: "given the current state I can only answer statistically — acceptable?"

## Common mistakes

- Skipping the path request and analysing a chat attachment directly: violates Step 1.
- Pasting all of `report.md` into the answer: the report is the working paper; answers carry conclusions and key numbers.
- Believing strict mode makes masking unnecessary: the two dimensions are independent.
- Printing unmasked values to the console or a temp file "to see the data": masking must happen before anything is written.
- Estimating instead of computing locally: estimates are almost always wrong at full-data scale.
- Sending `.xls` to a cloud converter "for convenience": violates a red line; re-save locally instead.
- Ignoring encoding: reading a GB18030 CSV as UTF-8 produces mojibake that pollutes the column profile. The script handles it; hand-written code must too.

## Note on languages

Detection rules target Chinese-language data formats (national ID numbers, 省/市 addresses, Chinese surnames) plus generic patterns (e-mail, phone, bank card, plate). Reports and console output are currently in Chinese; the detection logic is language-agnostic where patterns allow. For English-only documents, structured-table masking still works via column headers, but free-text name detection is weak — say so rather than over-promising.
