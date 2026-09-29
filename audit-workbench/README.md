# Audit Analytics Workbench (V1: Tier-A, SAP)

Offline, rule-based data analytics for data-driven compliance audit under
Office Circular No. 53/AMG/AAP/2026-27 (DGA Steel, Ranchi).

An audit team points the workbench at the SAP extracts it already has. It picks
the Annexure-I topics and runs pre-coded analytic tests. It gets two Excel files:

* **Working paper**: summary, data-quality and completeness checks, parameters
  used, and one evidence sheet per test with the rule, the exact SQL executed,
  the population tested and every exception row.
* **Annexure-II**: the Company-wise Data-Driven Audit Plan, pre-filled with
  exceptions, value involved (Rs. lakh), data source and mode, units with the
  most exceptions and records required for substantiation. Recommendation and
  reason codes stay with the team. The OO-53 checks (16 topics, 4 business
  processes, 8 recommended) are listed at the foot.

Nobody needs to know SQL or Python to use it. Anyone who wants to check a
result can read the SQL on the test's sheet.

## Design rules

1. **Fully offline.** No internet, no cloud, no API, no telemetry. The
   extracts are read on the audit machine and loaded into an in-memory
   database with extension auto-install switched off. Temporary files are
   deleted after the run. Only the two Excel outputs are written.
2. **Deterministic and transparent.** Every exception comes from a written
   rule with visible parameters. There are no risk scores and no AI.
3. **Red flags, not observations** (OO-53 para 10). Every output says so.
   Topics that depend on an approval are marked "not verifiable remotely".
4. **Not executable is not NIL.** If an extract, a column or a criterion
   (for example, the sensitive G/L list) is missing, the test is reported
   NOT EXECUTABLE with the reason. It is never shown as nil exceptions.
5. **Completeness first.** The Data Quality sheet shows file hashes, rows
   loaded, lines dropped (and why), control totals, date ranges and
   unreadable values. Reconcile these before relying on a NIL result.
6. **Annexure-I is the backbone.** Every test maps to catalogue topic IDs
   (`T1A01` = BP1 Priority-I Sl.1). IDs are permanent: when DAC revises the
   catalogue, the Annexure-I reference changes and the ID does not.

## What is in V1

All 166 Annexure-I topics are in the catalogue, plus two MSME topics
(OO-53 para 13). The Remote/On-site mode for each follows OO-53 para 4.
34 tests cover these processes:

| Process | Tests |
|---|---|
| Procurement | split POs, PO without PR, PO after GR, open POs, price variation, new/blocked vendors, GR/IR mismatch, vendor concentration, delayed delivery (LD) |
| Inventory | negative stock, non-moving stock, manual adjustments / PI differences / write-offs, GR not invoiced |
| Sales | below approved price, credit notes, overdue receivables, sales to blocked customers |
| Projects & CWIP | cost overrun / no budget, time overrun, cost after TECO, idle projects, CWIP ageing, TECO projects still in AuC |
| Finance | duplicate invoices (exact and near), entries after period close, high-value manual JEs, sensitive G/L postings, reversals, round values, holiday entries, long-pending advances |
| MSME | payments to micro/small enterprises beyond 45 days (with indicative s.16 interest), MSE procurement share vs 25% target |

Standard SAP sources the tests read: ME2N/ME2L, MB51, MB52, MB5S, LFA1,
FBL1N, FBL3N, FBL5N, KNA1, VF05N, CN43N, S_ALR_87013558, CJI3 and AR01. SE16N
table downloads work too: columns are recognised by SAP technical names
(EBELN, LIFNR, BUDAT ...) as well as by ALV headings.

## Using it

```
python -m auditwb topics                      # topics and the tests available for each
python -m auditwb init moil_2025-26.toml      # blank run profile
python -m auditwb checklist moil_2025-26.toml # which SAP extracts and columns are needed
python -m auditwb check moil_2025-26.toml     # load extracts, show column mapping, show what can run
python -m auditwb run moil_2025-26.toml       # run; outputs go to ./output next to the profile
python -m auditwb demo C:\temp\demo           # synthetic data, full run (for training)
```

The run profile is a short text file (see
`auditwb/library/TEMPLATE_profile.toml`). It holds the company, period and
cut-off, the SAP decimal/date settings, where each extract is, the sixteen
topics, and any threshold or criteria changes. Automatic column mapping
handles the usual headings. Where a company's layout differs, map it once in
the profile and reuse that profile for later audits of the same company:

```toml
[datasets.po_items]
files = ["extracts/ME2N_BSL.xlsx", "extracts/ME2N_BSL_Q4.xlsx"]
columns = { vendor = "Supplier Code", po_date = "PO Dt" }
```

## Adding or changing a test (for DAC / data leads)

Tests live in `auditwb/library/tests/*.toml`, one file per business process.
Each test has a plain-English rule, parameters with defaults, field questions,
records required, a population query and the exception query (DuckDB SQL over
the standard datasets in `auditwb/library/datasets.toml`). The exception query
must return `exception_id`, `exception_value` and `reason`. It may also
return `audit_unit`, which feeds the "Units/Departments" column. Add planted
cases for the new test to `auditwb/synthetic.py`, then run `pytest`. The
suite fails unless every test finds exactly its planted exceptions and none of
the near-misses.

## Installing on an office machine with no internet

On any internet-connected machine with Windows 64-bit Python 3.11 or later:

```
pip download duckdb openpyxl --only-binary=:all: -d wheels
```

Copy the `audit-workbench` folder and `wheels` through the approved medium,
then on the office machine:

```
pip install --no-index --find-links wheels duckdb openpyxl
cd audit-workbench
python -m auditwb demo C:\temp\demo
```

Nothing else is downloaded. The only dependencies are DuckDB and openpyxl.
A packaged single-folder build that needs no Python installation is planned
next, subject to IT approval.

## Known limits of V1

* SAP (Tier-A) only; Tally / Tier-B is out of scope for now.
* Status fields (vendor/customer blocks) are as on the extraction date, not
  the transaction date; the working paper says so.
* Topics needing change logs, condition-level pricing or HR/payroll data are
  on-site per OO-53 para 4 and have no remote test yet.
* Default document types, movement types, asset classes and MSE status codes
  are standard SAP values. Confirm each company's configuration during the
  ERP baselining exercise and set them in the profile.
