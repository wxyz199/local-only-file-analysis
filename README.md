# Local-Only File Analysis

**Analyse business files without uploading them. Ever.**

A drop-in agent skill that reads spreadsheets, Word/PowerPoint documents, PDFs and images **on your own machine**, masks personal data before anything enters an AI context, and produces an audit sheet proving the source files were never copied, modified or transmitted.

No API keys. No account. No telemetry. No network calls. Pure Python standard library — zero third-party dependencies.

```
$ python scripts/local_extract.py examples/sample-ledger.csv --workspace .local_analysis
[local-only] 模式=普通 文件=1 合计=694 B
  - sample-ledger.csv  [文本]  694 B  5行  已脱敏:姓名/电话/邮箱/地址/证件
[local-only] 脱敏命中（按出现次数）：姓名 5 处，电话 5 处，邮箱 5 处，详细地址 5 处，证件 5 处
[local-only] 网络请求：0（无任何上传/外发调用）
```

*(Console output is currently in Chinese — see [Honest limitations](#honest-limitations). The masking itself is what matters here: 25 personal-data hits replaced before the agent read anything.)*

---

## Why this exists

Every "chat with your data" workflow silently uploads your file. For a payslip, a customer ledger, a contract or a medical record, that upload is the compliance problem — not the analysis.

Most privacy tooling answers this with a policy. This answers it with a **local extraction step**: the file is parsed by a script on your disk, the script writes a masked profile, and the agent only ever reads that profile. In strict mode the profile contains no original text at all.

| | Typical "analyse this file" | This skill |
|---|---|---|
| File bytes | uploaded to a service | read locally, read-only |
| What the model sees | the whole document | a masked profile (or metrics only) |
| Names, phones, IDs | exposed verbatim | replaced with stable pseudonyms |
| Proof | none | `manifest.json` + `audit.md` with SHA-256 re-check |
| Network | assumed | verified zero, recorded as `network_used: false` |

---

## Install

The skill is a directory with a `SKILL.md` and a `scripts/` folder — the format understood by Claude Code, Codex, Cursor, and other agent runtimes that read `SKILL.md` files.

**Claude Code / skills-compatible agents**

```bash
git clone https://github.com/<your-account>/local-only-file-analysis.git \
  ~/.claude/skills/local-only-file-analysis
```

**Manual (any runtime)**
Clone anywhere and point your agent at the folder, or copy `SKILL.md` and `scripts/` into your existing skills directory:

| Runtime | Typical location |
|---|---|
| Claude Code | `~/.claude/skills/` |
| Codex | `~/.codex/skills/` |
| Cursor / others | `~/.agents/skills/` |

**Requirements:** Python 3.10+ on the machine. Nothing to `pip install`. Optional: `tesseract` for local image OCR, and a local `workbuddy-pdf` CLI for PDF text extraction (a built-in zlib fallback runs without it).

---

## Use

Ask your agent something that implies a confidential local file:

> "Profile this ledger locally — don't upload it. `/Users/me/Documents/q3-ledger.xlsx`"

The skill will ask for the absolute path first (by design), ask you to choose a mode, then run:

```bash
python scripts/local_extract.py <files...> --workspace .local_analysis
python scripts/local_audit.py .local_analysis --note "processed locally"
```

You get:

```
.local_analysis/
├── report.md      # masked structure + statistics + preview (the agent reads this)
├── manifest.json  # hashes, mode, mask hit counts, network_used: false
└── audit.md       # integrity re-check + red-line checklist
```

### Modes

| | Normal (default) | Strict |
|---|---|---|
| Report | structure, stats, masked preview | aggregate metrics only |
| Original text | masked, then visible | never present |
| Use for | general internal data | customer records, payroll, contracts, classified material |

`--mode strict --mask-region` is the tightest setting: metrics only, and even the province/city is hidden.

### Flags

| Flag | Effect |
|---|---|
| `--mode normal\|strict` | report depth |
| `--workspace DIR` | output directory (default `.local_analysis`) |
| `--rows N` | preview rows |
| `--mask-region` | mask the region instead of keeping province/city |
| `--redact REGEX` | mask additional custom patterns (repeatable) |
| `--mask-headers` | mask column headers too |
| `--quiet` | suppress console summary |

---

## What gets masked — always, in every mode

Masking has **no off switch**. Strict mode is masking *plus* no-original-text, never instead of it.

| Category | Fields | Becomes |
|---|---|---|
| Name | name, account holder, contact, legal representative, handler, account manager, applicant, payee, cardholder, shareholder | `[NAME#1]` |
| Contact | mobile, landline, fax, e-mail | `[PHONE#1]`, `[EMAIL#1]` |
| Address | residence, mailing, registered, domicile | `Zhejiang Hangzhou [ADDRESS#1]` |
| Identity | national ID (18/15 digit), passport, permit, credit code | `[ID#1]` |
| Financial | bank card, account, card number | `[ACCOUNT#1]` |
| Other | licence plate, customer/employee/member/student/social-security/policy number | `[PLATE#1]`, `[IDNO#1]` |

**Stable pseudonyms** are the point: the same value maps to the same tag for the whole run, so "top 10 customers by volume" still works — you just can't tell who they are. The value↔pseudonym map never touches disk.

---

## Supported formats

| Input | Route |
|---|---|
| `.xlsx` `.xlsm` | zip + XML parsed directly — multi-sheet, column-level masking |
| `.csv` `.tsv` `.txt` `.md` `.json` `.html` | encoding auto-detect (UTF-8 / GB18030), delimiter sniffing |
| `.docx` `.pptx` | zip + XML — heading outline, paragraphs, tables, per-slide bullets |
| `.pdf` | local PDF CLI if present, built-in zlib text extraction otherwise |
| `.png` `.jpg` `.gif` `.bmp` `.webp` | format, dimensions, SHA-256 (OCR only if `tesseract` is installed) |
| `.xls` `.doc` `.ppt` legacy | asks you to re-save locally — never sent to an online converter |

---

## Honest limitations

This is a privacy tool, so the limits matter more than the features.

1. **The inference path is not isolated.** The skill guarantees no upload to cloud documents, drives, knowledge bases, converters, OCR services or search engines, and in strict mode no original text reaches the model. It does **not** claim "zero bytes leave the machine" — the model still runs remotely, and the masked report is its input. Any tool promising otherwise is lying.
2. **Free-text name detection is heuristic.** Structured tables mask column-wide by header, which is reliable. Names buried in prose are found via regex plus surname/context patterns, so isolated names with no cue can be missed. Structured identifiers (ID numbers, phones, e-mails, cards) are pattern-based and much more reliable.
3. **Chinese-language data is the design centre.** National ID numbers, 省/市 addresses and Chinese surnames are first-class; reports and console output are in Chinese. English documents get reliable table masking via headers, weaker prose masking.
4. **No OCR out of the box.** Without `tesseract`, images are metadata only, and the skill says so instead of quietly routing the image to a vision API.
5. **Attachments may already be gone.** A file dragged into a chat may already sit on the platform's servers. The skill asks for a local path first for exactly this reason and labels attachment-sourced input honestly.

---

## Verify it yourself

```bash
python scripts/local_audit.py .local_analysis
```

Re-hashes every source file, compares against the values recorded before processing, lists every artefact written, checks six red-line categories, and exits non-zero if a source file changed while it was being analysed.

```jsonc
// manifest.json (excerpt)
{
  "mode": "normal",
  "masking": { "enabled": true, "mandatory": true, "hits": { "phone": 14 } },
  "network_used": false,
  "files": [{ "sha256": "…", "kind": "sheet" }]
}
```

---

## License

MIT. Use it, fork it, ship it inside your own product.
