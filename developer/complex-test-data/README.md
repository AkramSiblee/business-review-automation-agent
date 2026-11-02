# Complex test dataset: Halcyon Industrial Holdings, FY2025

**Developer folder.** This is the technical answer key for the Halcyon stress test. Users never need these files:
the input lives in `sample-data/complex-test-data/` and the plain-English answer key (Excel, with a scoring sheet)
lives in `sample-data/answer-keys/`. Keep both files below out of the inputs folder: a cycle that can read the
answer key proves nothing.

`Business_Review_Complex_Input_FY2025.xlsx` is a synthetic, single-workbook input that plants every scenario in
`.claude/skills/business-review-automation/references/complex-scenarios-playbook.md`. Use it after the Meridian
sample (`sample-data/sample-test-data/`), which tests the basics. This one tests whether a cycle copes with a
messy, multi-entity, multi-standard group.

| File | Where | What it is |
|---|---|---|
| `Business_Review_Complex_Input_FY2025.xlsx` | `sample-data/complex-test-data/` | The input. Copy it alone into the inputs folder. |
| `Complex_Test_Answer_Key_FY2025.xlsx` | `sample-data/answer-keys/` | User-facing answer key: what a cycle produces, every planted issue in plain English, expected figures, and a scorecard. |
| `EXPECTED.json` | here | Reference values a correct cycle should reproduce (goodwill, NCI, JV share, disposal gain, covenant headroom, and more). |
| `generate_complex_dataset.py` | here | Rebuilds the input workbook and `EXPECTED.json`. It is deterministic, so re-running gives identical data. |
| `build_answer_key.py` | here | Rebuilds the user-facing answer key from `EXPECTED.json`. Re-run it after regenerating the dataset. |

Every ledger balances by construction, using the same interim trial-balance convention as Meridian. The only
imbalances are the planted ones (rows 2.2, 2.12 and JV01 below).

## The group

| Entity | Country | Currency | Standard | Ownership / method | What makes it hard |
|---|---|---|---|---|---|
| HQ01 | United States | USD | US GAAP | Parent | Term loan and covenants, investments in subsidiaries at cost, LIFO inventory, CAD-paid staff |
| DE01 | Germany | EUR | IFRS | 100%, full | German account names, year-to-date export, intercompany buyer |
| BR01 | Brazil | BRL | IFRS | 70%, full (NCI) | Portuguese names, separate tab in BRL '000, April missing |
| TR01 | Turkey | TRY | IFRS | 100%, full | Hyperinflationary economy, 20% devaluation in March, loss-making, subtotal rows |
| MX01 | Mexico | MXN | IFRS | 100%, acquired 1 Jul 2025 | Spanish names, signed single-amount ledger, Jul-Dec only |
| UK01 | United Kingdom | GBP | IFRS | 100%, sold 30 Sep 2025 | Discontinued operation, Jan-Sep only |
| JV01 | India | INR | Ind AS | 40%, equity method | P&L only, April fiscal year |
| HLD1 | Cayman Islands | USD | US GAAP | 100%, full | Dormant holding company, balance sheet only |

## Answer key

The **Auto** column shows whether `scripts/profile_inputs.py` flags the scenario by itself (⚙). Every other row
has to be caught by the cycle's own judgment (Phases 2-5). The section numbers refer to the playbook.

### Data shape (Phase 1)

| # | Scenario | Where | Auto | Expected treatment |
|---|---|---|---|---|
| 2.1 | Year-to-date P&L | DE01 rows, main GL tab | ⚙ | Difference the months into monthly activity. Retained earnings are carried at the 1 Jan balance every month. |
| 2.2 | Subtotal rows | TR01 "Total Revenue" / "Gross Profit" with no account code | ⚙ | Drop them, then TR01 balances. |
| 2.3 | Months as columns | `09_Operational_KPIs` | ⚙ | Unpivot to one row per period. |
| 2.4 | Signed single amount column | `MX01 Ledger (post-acq)` | ⚙ | Credit-negative convention. Normalize, then it balances. |
| 2.5 | Unit scale | `BR01 GL (BRL 000)` title row "Valores em BRL '000" | ⚙ | Multiply by 1,000 (rounding differences under 5 BRL a month). |
| 2.6 | Excel error | `05_Budget_Forecast`, HQ01 2025-10 account 6300 is `#REF!` | ⚙ | Treat as missing, not zero. |
| 2.7 | Missing month | BR01 has no GL rows for 2025-04 (the bank data does) | ⚙ | Don't interpolate. The RE rollforward breaks Mar to May; April NI is in `EXPECTED.json`. |
| 2.8 | Mixed currencies | HQ01 roster has 5 CAD-paid engineers; HQ01 has a EUR bank account | ⚙ (roster) | Translate to USD. GL cash ties to the bank only after translating the EUR account at spot. |
| 2.9 | Chart-of-accounts change | HQ01 6100 Marketing Programs becomes 6110 Digital Marketing from July | ⚙ | Map both to the same line. The budget still uses 6100 (row 9.4). |
| 2.10 | Period 13 | HQ01 `2025-13`: audit adjustment of 85,000 (legal accrual) | ⚙ | Fold into December and log. |
| 2.11 | Hidden duplicate tab | `Copy of 02_CashBank_Position` (hidden, HQ01 rows only) | ⚙ | Exclude: it double-counts HQ01 cash. |
| 2.12 | Forecast leakage | HQ01 rows dated 2026-02, notes "FC"; one bookings row dated 01/15/2026 | ⚙ (GL) | Exclude from FY2025 actuals. |
| - | Date formats | Headcount and bookings: HQ01 MM/DD/YYYY, DE01 DD.MM.YYYY, BR01/MX01/UK01 DD/MM/YYYY, TR01 native dates | ⚙ | Detect per entity. HQ01 is the US-format trap. |
| - | Languages | German (DE01), Portuguese (BR01) and Spanish (MX01) account names; Account Type stays English | - | Map by account type and `multilingual-term-tables.md`. |

### Group structure (Phase 3)

| # | Scenario | Where | Auto | Expected treatment |
|---|---|---|---|---|
| 3.1 | NCI | BR01 70% owned | ⚙ | Consolidate 100%. NCI = 30% of BR01 net income and net assets (`br01_nci`). |
| 3.2 | Equity method | JV01 40%, P&L only | ⚙ | One-line share of profit (`jv01_equity_method`). Exclude it from group revenue, headcount and cash. Its trial balance won't balance (no balance sheet supplied). HQ01 carries the JV at cost (4.0m), which needs adjusting. |
| 3.4 | Mid-year acquisition | MX01 from 1 Jul, events tab | - | Consolidate Jul-Dec only. Provisional goodwill = 25.0m less net assets at Jun-30 spot (`mx01_acquisition`). Deal costs of 650k are one-time; the earn-out is a liability. Show organic vs reported growth. MX01 roster hire dates go back years: count staff from 1 Jul only. |
| 3.5 | Disposal | UK01 sold 30 Sep, events tab | - | Present as a discontinued operation. Parent books a 3.0m gain against cost, but the consolidated gain is proceeds less UK01 net assets (`uk01_disposal`), before recycling CTA. Restate KPIs on continuing operations. The UK budget still runs to December (row 9.4). |
| 3.6 | Different fiscal year | JV01 starts in April | ⚙ | Align to the calendar year using the monthly P&L, and disclose. |
| 3.7 | Dormant entity | HLD1, balance sheet only | ⚙ | Consolidate the balance sheet. No headcount, bookings or budget expected. |
| - | Investment elimination | HQ01 1500 Investment in Subsidiaries | - | Eliminate against subsidiary equity at acquisition. Without this, consolidated equity is overstated. |
| - | Intercompany | HQ01 sells goods to DE01; HQ01 charges BR01 a management fee | - | Eliminate. December differs by 150,000 EUR of goods in transit that DE01 hasn't booked: eliminate the lower amount and disclose the difference. |

### Currency (Phases 1-3)

| # | Scenario | Where | Auto | Expected treatment |
|---|---|---|---|---|
| 4.1 | Hyperinflation | TR01, Turkey; no price index supplied | ⚙ (FX step) | IAS 29 applies but can't be applied without an index. This is a LOW item: present TR01 with a "not restated for hyperinflation" caveat. |
| 4.2 | Devaluation | TRY spot falls 20% in March, then 2.5% a month | ⚙ | Split the variance at constant currency. TR01's FX losses come from USD supplier payables. |
| 4.3 | Missing rate | MXN period-average for 2025-08 | ⚙ | Use August spot as a proxy (MEDIUM) and disclose. |
| 4.5 | Trapped cash | TR01 bank note on capital controls | - | Treat as restricted for group liquidity. |

### Liquidity and solvency (Phase 4)

| # | Scenario | Where | Auto | Expected treatment |
|---|---|---|---|---|
| 5.1 | Covenant near breach | `08_Debt_Covenants`: net leverage covenant of 1.50x or below | - | Recompute: about 1.40x, roughly 7% headroom, which is a red flag. The company reports 1.22x because it leaves the revolver out of net debt; flag the difference (`covenants`). |
| 5.2 | Debt maturity | Term loan matures 30 Jun 2026 | - | Within 12 months and no refinancing in the file: a going-concern indicator. |
| 5.3 | Financing inflow | 5.0m revolver draw in November | - | Separate it from operating cash flow. |
| 5.4 | Going-concern indicators | Rows 5.1 and 5.2 together | - | Two indicators present, so they go on the deck's first slide. List them; don't conclude. |
| 5.5 | Payables stretch | HQ01 payables rise from 45 to 80 days Oct-Dec | - | Treat the cash benefit as non-recurring. |
| - | Restricted cash | HQ01 LC collateral account (1.5m) | - | Exclude from liquidity and runway. |

### Earnings quality, revenue, tax and cost (Phases 4-5)

| # | Scenario | Where | Auto | Expected treatment |
|---|---|---|---|---|
| 6.1 | One-time items | HQ01 deal costs, gain on disposal; DE01 restructuring | - | Build an Adjusted EBITDA bridge, reported figure first. |
| 6.1 | Untagged one-time item | HQ01 7920 earn-out fair-value change, Dec 300k, tagged N | - | Apply the policy: one-time, and log it. |
| 6.2 | Recurring "one-time" | HQ01 7900 restructuring, 120k in Feb, May, Aug and Nov, all tagged Y | - | Flag it: four times in a year. Show Adjusted EBITDA with and without it. |
| 6.3 | Revenue pulled forward | HQ01 December product revenue +35% (bill-and-hold), receivables 45 to 60 days | - | Flag it; never adjust revenue (`hq01.dec_bill_and_hold_uplift_usd`). |
| 6.4 | Capitalization shift | HQ01 capitalized development 150k a month in H1, 400k in H2; R&D opex falls while the roster is flat | - | Show opex both as reported and with capitalization added back. |
| 6.5 | Undisclosed provision release | DE01 warranty provision release of 600k EUR in November (negative warranty expense); not in the events tab | - | Find it, quantify it, and show margin with and without it. |
| 6.6 | Related party | Bookings customer "Halcyon Family Trust Ltd (related party)" | - | Disclose. |
| 7 | Deferred revenue decline | HQ01 deferred revenue falls about 50% over the year while service revenue is flat | - | Leading churn indicator. |
| 7 | Concentration | Atlas Automotive is about 16% of gross bookings | - | Above the 15% top-1 threshold. |
| 7 | Single large churn | DE01 Nordwerk GmbH, 900k EUR churn in August | - | Commercial storyline. |
| 8 | Inventory build | HQ01 inventory days rise from 70 to 120 in H2 | - | Obsolescence risk. HQ01 uses LIFO and the other entities FIFO/IFRS, so inventory isn't comparable. |
| 8 | Loss-making entity, no tax benefit | TR01 | - | Impairment trigger (sustained real decline). |
| 1 | Roster vs ledger, explained | DE01 ledger salaries = roster x 1.21 (employer social charges) | - | Ledger is authoritative. Disclose the mismatch as explained. |
| 1 | Roster vs ledger, inflation | TR01 roster cost is a January snapshot; ledger salaries inflate 4.5% a month | - | Ledger is authoritative. |

### Budget (Phase 4)

| # | Scenario | Where | Auto | Expected treatment |
|---|---|---|---|---|
| 9.1 | Two versions | HQ01 Original Budget plus a Reforecast (Jul) for revenue | - | Compare actuals with both. |
| 9.2 | Fixed budget rate | DE01 budget in USD at 1.05 EUR/USD | - | Compare at the budget rate and show FX separately. |
| 9.4 | Budget on old chart of accounts | HQ01 budget uses 6100 all year | - | Map 6110 actuals onto it. |
| 9.4 | Budget for a disposed entity | UK01 budget runs Jan-Dec | - | Compare like-for-like on Jan-Sep only. |
| 9.4 | Nominal budget under inflation | TR01 budget flat at January prices | - | Variances are mostly inflation. Say so; don't call them performance. |
| 9.4 | No budget | MX01 (acquired), JV01, HLD1 | ⚙ (MX01) | Like-for-like scope only. |
| - | Overspend | HQ01 S&M salaries budgeted at 85% of the roster | - | Real overrun, about +18%. |

## How to run a cycle on it

1. Put `Business_Review_Complex_Input_FY2025.xlsx` (and nothing else) into the resolved inputs folder.
2. Ask: *"Run a business review cycle on the complex test data."*
3. Compare the Data Validation Report and the deliverables with this table and `EXPECTED.json` (or use the
   Scorecard in `sample-data/answer-keys/Complex_Test_Answer_Key_FY2025.xlsx`). A good run:
   - catches every ⚙ row at the checkpoint
   - raises the hyperinflation index (4.1) as a LOW question
   - reproduces the `EXPECTED.json` values within rounding
   - leads the deck with the covenant and maturity items (rows 5.1 and 5.2).
