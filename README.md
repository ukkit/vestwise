# 🏛️ vestwise

Turn eTrade RSU/ESPP history into ITR-ready capital gains schedules.

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

## Download Excel file

Login to eTrade. Go to `At Work` -> `My Account`, make sure you are on `Benefit History` tab. Then click `Download` drop-down & select `Download Expanded`, this will download filed named `BenefitHistory.xlsx`.

## Usage

**Sample input file** [`sample/BenefitHistory.xlsx`](sample/BenefitHistory.xlsx)

Place `BenefitHistory.xlsx` (downloaded from eTrade) in the project directory, then:

```bash
uv run script.py
```

Writes a timestamped output file, e.g. `20260301_120000_rsu_summary.xlsx` in project directory.

## Configuration

Copy `vestwise.ini.sample` → `vestwise.ini` to override defaults without editing source code:

```bash
cp vestwise.ini.sample vestwise.ini
```

`vestwise.ini` is gitignored — edit it freely. All keys are optional; the script falls back to hardcoded defaults if the file is absent or a key is missing.

| Key | Section | Default | Purpose |
|-----|---------|---------|---------|
| `ltcg_rate` | `[tax]` | `0.125` | LTCG tax rate |
| `stcg_rate` | `[tax]` | `0.30` | STCG tax rate (marginal slab) |
| `ltcg_holding_months` | `[tax]` | `24` | Holding threshold for LTCG |
| `input_file` | `[paths]` | `BenefitHistory.xlsx` | Input spreadsheet path |
| `output_file_template` | `[paths]` | `{timestamp}_rsu_summary.xlsx` | Output filename (`{timestamp}` is substituted) |
| `sbi_ttbr_cache_file` | `[paths]` | `data/SBI_REFERENCE_RATES_USD.csv` | Local SBI rate cache |
| `sale_price_overrides_file` | `[paths]` | `data/sale_price_overrides.csv` | Sale price overrides |

## Output

[`sample/20260301_124234_rsu_summary.xlsx`](sample/20260301_124234_rsu_summary.xlsx) (output)

<details>
<summary><strong>Output Sheets</strong></summary>

| Sheet | What it's for |
|---|---|
| **Grant Summary** | One row per grant — quantities, sellable shares, next vest date, validation status |
| **Vesting Schedule** | Every past and future vest tranche with days-to-vest and FMV |
| **Sales History** | Every sale — capital gain/loss, STCG/LTCG classification, INR amounts |
| **Year-wise Tax Summary** | FY-wise capital gains totals (STCG / LTCG) with subtotals per year |
| **Tax Withholdings** | RSU tax-withholding events with INR exchange rates |
| **Schedule FA (Table A3)** | ITR foreign asset disclosure — one row per calendar year per company (see below) |

</details>

<details>
<summary><strong>Schedule FA (Table A3)</strong></summary>

Schedule FA is the foreign asset disclosure required in ITR-2/ITR-3. It is based on **Calendar Year** (Jan–Dec), not the Indian Financial Year.

The sheet has one row per CY per company ticker, and contains all the numbers you need to fill the ITR form directly:

| Column | What to use it for |
|---|---|
| **CY / AY** | Identifies which ITR filing this row applies to (e.g. CY2024 → AY2025-26) |
| **Date Since Held** | "Date since held" field in Schedule FA — computed using FIFO on actually-released shares |
| **Vested in CY / Sold in CY** | Activity summary for the calendar year |
| **Shares Held (Dec 31)** | Closing balance to enter in Schedule FA |
| **Dec 31 Price / Dec 31 Rate** | Stock price and SBI TTBR used to compute the closing INR value |
| **Peak Balance (INR)** | Peak value field in Schedule FA — (shares at Jan 1 + released in CY) × CY high × Dec 31 rate |
| **Closing Balance (INR)** | Closing value field in Schedule FA — Shares Held × Dec 31 Price × Dec 31 Rate |
| **Acquisition Value ($) / (INR)** | Total cost basis of shares released in this CY |
| **Sale Proceeds ($) / (INR)** | Total sale proceeds in this CY |

> Before filing, replace the ticker symbol in "Name of Entity" with the company's full legal name and registered address.

</details>

<details>
<summary><strong>Indian Tax Rules Applied</strong></summary>

- **Acquisition date** = vest/release date (not grant date)
- **Cost basis** = FMV on release date (matches Form 16 perquisite value)
- **LTCG threshold** = 24 months (foreign/unlisted shares)
- **LTCG rate** = 12.5% | **STCG rate** = 30% (slab) — override via `vestwise.ini`
- **Exchange rate** = SBI TTBR on last business day of the preceding month (Rule 115)
- **Share quantities** = net released shares (after tax withholding), not gross vested

</details>

## Data Files

### `data/SBI_REFERENCE_RATES_USD.csv`
SBI Telegraphic Transfer Buying Rates, auto-downloaded from [sahilgupta/sbi-fx-ratekeeper](https://github.com/sahilgupta/sbi-fx-ratekeeper) on first run and cached locally (refreshed every 7 days). Rates before January 2020 fall back to yfinance market rates with a `[WARNING]` — this is expected for older grants.

### `data/sale_price_overrides.csv`
Persists actual sale execution prices across runs. Auto-populated on first run using the brokerage price (from the xlsx) or the yfinance closing price as a fallback. Sorted newest-first.

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

> `[WARNING]` lines for pre-2020 dates are expected — SBI TTBR data is only available from January 2020 onward.

---

## ⚠️ LTCG Holding Months = 24

```text
The 24-month threshold is specific to unlisted/foreign shares under Indian tax law. Listed Indian shares use a different threshold (12 months for LTCG), but that's not handled here since this tool targets US-listed company equity.
```
