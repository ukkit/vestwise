# VestWise

RSU/ESPP capital gains tracker with Indian tax compliance. Processes a Schwab `BenefitHistory.xlsx` export and produces a formatted Excel summary.

> [!CAUTION]
> **For personal reference only.** This tool is not a substitute for professional tax advice. Do not use its output for actual tax filings. Tax laws change, calculations may be incorrect, and individual circumstances vary. Always consult a qualified chartered accountant for your ITR.

## Installation

Requires [uv](https://docs.astral.sh/uv/getting-started/installation/).

```bash
# Install uv (if not already installed)
curl -LsSf https://astral.sh/uv/install.sh | sh

# Clone and run — uv installs dependencies automatically
git clone <repo-url>
cd vestwise
uv run script.py
```

No need to create a virtualenv or run `pip install` — `uv` reads `pyproject.toml` and handles everything on first run.

## Usage

Place `BenefitHistory.xlsx` (downloaded from Schwab) in the project directory, then:

```bash
uv run script.py
```

Writes a timestamped output file, e.g. `20260227_210804_rsu_summary.xlsx`.

## Output Sheets

| Sheet | Contents |
|---|---|
| **Summary** | One row per grant — quantities, sellable shares, next vest date |
| **Vesting Schedule** | All past and future vest tranches with days-to-vest |
| **Sales History** | Every sale with capital gain, tax type/rate, INR amounts |
| **Tax Summary** | FY-wise capital gains totals (STCG / LTCG) |
| **Tax Withholding** | RSU withholding events with INR exchange rates |

## Indian Tax Rules Applied

- **Acquisition date** = vest date (not grant date)
- **Cost basis** = FMV on vest date
- **LTCG threshold** = 24 months (foreign/unlisted shares)
- **LTCG rate** = 12.5%, **STCG rate** = 30% slab
- **Exchange rate** = SBI TTBR on last business day of the preceding month (Rule 115)

## Data Files

### `data/SBI_REFERENCE_RATES_USD.csv`
SBI Telegraphic Transfer Buying Rates, auto-downloaded from [sahilgupta/sbi-fx-ratekeeper](https://github.com/sahilgupta/sbi-fx-ratekeeper) on first run and cached locally.

### `data/sale_price_overrides.csv`
Persists actual sale execution prices across runs. Auto-populated on first run using the brokerage price (if present in the xlsx) or the yfinance closing price. Sorted newest-first.

| Column | Description |
|---|---|
| `grant_id` | Grant number |
| `sale_date` | `YYYY-MM-DD` |
| `sale_seq` | 1-based; disambiguates multiple sales on the same date for the same grant |
| `sale_price_usd` | Price per share (2 decimal places) |
| `sale_quantity` | Shares sold |
| `source` | `xlsx`, `yfinance`, or `manual` |
| `notes` | Free-text, empty by default |

**To correct a price:** edit `sale_price_usd` directly, set `source=manual`, optionally add a note. Existing entries are never overwritten by the script — only new sales get appended.

## Sample Output:

```bash
# First Rtime execution
Reading file: BenefitHistory.xlsx
Reading file: BenefitHistory.xlsx
Found ESPP sheet
Found Restricted Stock sheet

Processing Restricted Stock sheet...
[WARNING] SBI TTBR not available for 06/04/2018, using yfinance market rate
[WARNING] SBI TTBR not available for 06/01/2017, using yfinance market rate
[WARNING] SBI TTBR not available for 06/17/2016, using yfinance market rate
[WARNING] SBI TTBR not available for 07/03/2019, using yfinance market rate
[WARNING] SBI TTBR not available for 06/04/2018, using yfinance market rate
[WARNING] SBI TTBR not available for 06/01/2017, using yfinance market rate
[WARNING] SBI TTBR not available for 06/17/2016, using yfinance market rate
[WARNING] SBI TTBR not available for 07/03/2019, using yfinance market rate
[WARNING] SBI TTBR not available for 07/03/2019, using yfinance market rate
[WARNING] SBI TTBR not available for 06/04/2018, using yfinance market rate
[WARNING] SBI TTBR not available for 06/17/2016, using yfinance market rate
[WARNING] SBI TTBR not available for 06/08/2015, using yfinance market rate
[WARNING] SBI TTBR not available for 03/24/2015, using yfinance market rate
Found 19 Restricted Stock grants

Processing ESPP sheet...
Found 3 ESPP grants
[OK] Saved 35 sale price overrides to ~/vestwise/data/sale_price_overrides.csv
Found 22 total grants
Summary created successfully: 20260227_213037_rsu_summary.xlsx

[OK] All grants passed validation checks!
```
