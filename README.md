# Medicine RFQ - Comparative Statement (CST)

`CST_RFQ_Medicines.xlsx` compares the vendor quotations received against the
medicines RFQ, item by item and vendor by vendor.

| Folder / file | What it is |
|---|---|
| `quotes/` | Vendor-filled RFQ sheets as received (Highnoon, Hoechst, Horizon, Jamil Traders, MP) |
| `reference/` | The earlier hand-made CST. Source of the MCOLSON and SAMI quotes (no files received), of recorded pack sizes, and of the previous awards used for review |
| `build_cst.py` | Rebuilds the CST from the files above |
| `CST_RFQ_Medicines.xlsx` | The output |

## Rebuild

```bash
pip install xlrd openpyxl
python3 build_cst.py
```

If LibreOffice is installed the script also recalculates the formulas so the
saved file shows values in viewers that do not calculate. Excel recalculates on
open regardless.

## Sheets in the output

1. **CST** - one row per RFQ item; per vendor: Price, Pack, Unit Price, Total, Brand.
   Summary columns give number of quotes, L1 per pack, L1 per unit and a review flag.
   Green = lowest pack price, blue = lowest unit price, pink = no valid quote.
2. **Vendor Summary** - items quoted, value, L1 counts per vendor.
3. **Review vs Previous CST** - previous award against the new L1 for every quoted item.
4. **Review Notes** - findings to settle before awarding.
5. **Source Quotes** - every priced row exactly as found in the sources.

All figures on sheets 1-2 are formulas. Correcting a price or filling in a Pack
cell updates L1, totals and the summary.
