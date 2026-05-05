# MB Calculator

This project exposes the workbook `MB_Calc2_Apr 2026_auto.xlsx` through a small Flask web app.

Workflow:

1. Upload an Excel workbook that contains a single input sheet with columns `A:BD`.
2. The app reads the first sheet from the uploaded workbook and uses it as raw input data.
3. The calculation uses the local reference workbook `MB_Calc2_Apr 2026_auto.xlsx` for rates, output headers, and output formatting.
4. The app exports a clean output workbook with:
   - input columns `A:BD`
   - final output columns `CO:CW`

## Run

```powershell
venv\Scripts\python.exe -m pip install -r requirements.txt
venv\Scripts\python.exe app.py
```

Open `http://127.0.0.1:5000`.

## Render / Gunicorn

This project can be served with Gunicorn using the WSGI entrypoint in `wsgi.py`.

Example start command for Render:

```bash
gunicorn wsgi:application
```

If you want a more explicit command on Render, use:

```bash
gunicorn --bind 0.0.0.0:$PORT wsgi:application
```

## Important constraint

This implementation computes the workbook in Python. The uploaded file does not need to contain sheets like `MAINCPF CPH DEP SEP END RBC CSR`, `DEP_BONUS CALC`, or `Rate Table`.

Instead, the project keeps those rules in the local reference workbook:

- `MAINCPF CPH DEP SEP END RBC CSR`
- `Rate Table`
- `DEP_BONUS CALC`
- `SEP_BONUS CALC`
- `END_BONUS CALC`
- `RBC_BONUS CALC`
- `CSR_BONUS CALC`

The app does not depend on Microsoft Excel.

## Output behavior

- The uploaded workbook only needs one sheet. The processor reads the first sheet automatically.
- Source columns `C:Y` remain visible in the downloaded workbook.
- If the uploaded workbook has the `CO:CW` header cells removed or blanked, the app restores those output column names in the generated file.
- The generated workbook copies the original sheet's formatting for the exported columns, including widths, fills, borders, alignment, row heights, and number formats.

## Main-sheet formula map

The main calculation block starts at `BF` and runs to `CM`. The row formulas are consistent across the sample rows.

| Column | Header | Formula pattern summary |
| --- | --- | --- |
| `BF` | `Status` | mirrors column `B` |
| `BG` | `Policy Year` | `DATEDIF(AF, AH, "Y")` |
| `BH` | `PPT * Freq` | `AD * AE` |
| `BI` | `Prem Paid * Freq` | rounded elapsed premium count from `AF` to `AG` |
| `BJ` | `Paidup Factor` | `BI / BH` |
| `BK` | `Annualised premium CSR, END` | array formula based on premium fields |
| `BL` | `Updated SA` | depends on policy status and paid-up factor |
| `BM` | `For Current Month Accrued Loan Interest` | interest from `M`, `N`, cutoff date `BM1`, rate `BN1` |
| `BN` | `Updated LN_LA` | `N + BM` |
| `BO` | `Premium Paid Years` | rounded `DATEDIF(AF, AG, "Y")` |
| `BP` | `Lapsed Death Case` | conditional status based on death, lapse, and cess dates |
| `BQ` | `Policy Inforce Tenure for CPF and CPH` | rounded elapsed years from `AF` to `AG` |
| `BR` | `Benefit Calculator Option...` | `VLOOKUP` to `Rate Table!B16:G21` |
| `BS` | `GA` | CPF/CPH guaranteed addition |
| `BT` | `GMA` | extra CPF/CPH addition for `SS` |
| `BU` | `Base Benefit` | array formula for CPF/CPH base payout |
| `BV` | `Maturity Benefit TOTAL` | sum of `BS:BU` with `IFNA` guards |
| `BW` | `Guaranteed Additions (RBC)` | RBC guaranteed addition |
| `BX` | `Terminal Bonus (RBC)` | `Rate Table!A9:G13` with `HLOOKUP` by policy year |
| `BY` | `Rev. Bonus for RBC` | lookup to `RBC_BONUS CALC` |
| `BZ` | `Maturity Benefit RBC` | RBC output amount |
| `CA` | `Rev. Bonus for DEP` | lookup to `DEP_BONUS CALC` |
| `CB` | `Maturity Benefit DEP` | DEP output amount |
| `CC` | `SA Due or Rev. Bonus Due SEP` | decides SEP branch from policy year vs term |
| `CD` | `Maturity Benefit or Rev. Bonus Amount for SEP` | array formula from SEP logic |
| `CE` | `Rev. Bonus CSR` | lookup to `CSR_BONUS CALC` |
| `CF` | `High SA Additions Amt` | `Rate Table!A2:F6` lookup times updated SA |
| `CG` | `Maturity Benefit CSR` | `CE + CF + BL` branch logic |
| `CH` | `CSR SB Paid` | array formula for previously paid survival benefits |
| `CI` | `100.1% CSR` | top-up check against `100.1%` of annualised premium |
| `CJ` | `Rev. Bonus for END (1)` | lookup to `END_BONUS CALC` |
| `CK` | `100.1 condition bonus (2)` | END top-up check |
| `CL` | `Maturity Benefit END` | END output amount |

The hidden complexity is in the bonus tabs. Those sheets calculate year-wise bonus accrual tables by policy year and then the main sheet pulls the final totals through `VLOOKUP`.

## Exported output block

The workbook's final user-facing output block is `CO:CW`:

- `CO`: `LN_LA Updated`
- `CP`: `Guaranteed Addition`
- `CQ`: `GMA`
- `CR`: `Overall Revisionary Bonus`
- `CS`: `Special Bonus`
- `CT`: `Terminal Bonus`
- `CU`: `Maturity Benefit`
- `CV`: `Condition of 100.1% , additional bonus for CSR and END`
- `CW`: `Total Maturity Benefit (inclusive of Rev. Bonus wherever applicable)`
