# Examples

`sample-ledger.csv` is synthetic — invented names, ID numbers, phone numbers and addresses, shaped like a real Chinese customer ledger.

```bash
python scripts/local_extract.py examples/sample-ledger.csv --workspace .local_analysis
cat .local_analysis/report.md    # every personal field is already a pseudonym
python scripts/local_audit.py .local_analysis
```

Expected console line: `已脱敏:姓名/电话/邮箱/地址/证件` — 25 hits across 5 rows, 0 raw values in the report.
