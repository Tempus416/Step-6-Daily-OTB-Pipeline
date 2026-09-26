# Step-6-Daily-OTB-Pipeline
Consolidated, production-ready Python pipeline combining reference data loading, ingestion, cleaning, business logic, and validation for daily hotel OTB reporting across multiple properties — replaces Steps 1-5's exploratory notebooks with a single runnable pipeline, available as both a script and a notebook.


# Step 6 Daily OTB Pipeline

The consolidated, production-facing version of the Daily OTB pipeline — combines everything validated across Steps 1-5 (reference data, ingestion, cleaning, business logic, validation) into a single runnable pipeline instead of five separate exploratory notebooks.

## Two Formats, Same Logic

- **`Step 6 Daily OTB Pipeline.py`** — a standalone script, meant for unattended/scheduled daily execution: `python "Step 6 Daily OTB Pipeline.py"`
- **`Step 6 Daily OTB Pipeline.ipynb`** — the same logic in notebook form, for interactive use and inspecting intermediate results (`results[63638]["monthly"]`, etc.)

Both were tested independently and confirmed to produce identical output.

## What It Does

For each property (currently The Villages and Mansion Hotel):
1. Loads that property's reference/dimension workbook
2. Discovers and ingests every standardized daily OTB file
3. Cleans the data (removes summary rows, deduplicates, validates rate codes, checks nulls, verifies historical consistency)
4. Computes pickup, pace variance (STLY, both DOW- and Date-aligned), monthly roll-ups (including RevPAR and Occupancy %), and market segment roll-ups
5. Reconciles output against the source system's own totals (industry-standard ±3% tolerance)
6. Produces a recurring **unmapped rate code report** — a sorted, row-counted, date-ranged list of any rate code not yet in the reference mapping table

Returns a dict per property with every intermediate stage (`raw`, `clean`, `pickup`, `variance`, `monthly`, `segment`, `reconciliation`, `unmapped_codes`, `log`) available for inspection.

## Validated At Production Scale

Run against the full real dataset for both properties: 43 daily snapshots each, ~232K and ~252K raw rows respectively. Reconciliation passed at 43/43 snapshots for both properties, well within the ±3% industry-standard tolerance (actual variance: ~0.001%).

## A Deliberate Design Decision: What's *Not* Here

The hand-verified example spot-checks from Step 5 (specific hardcoded cases like a particular rate code/stay date combination proven to pick up +5 rooms between two known snapshots) are intentionally excluded. Those were one-time proof points used to validate the *logic* during development — they don't generalize to new data and would eventually produce false failures if run against future dates. They remain in Step 5's original notebook as a development record. `reconcile_totals`, which compares against the source's own data on any given day, is the reusable, generic check included here instead.

## Known Limitations / Not Yet Addressed

- **PropertyID is currently each property's STR ID** — a third-party identifier system reused as an internal key. Flagged early in this project as needing its own internal ID scheme; not yet implemented.
- **Functions are not yet split into importable modules** (`readers.py`, `cleaning.py`, `business_logic.py`). Considered and deliberately deferred — see project notes for rationale.
- **Unmapped rate code report is currently console/in-memory output only.** No file export or notification mechanism yet — relevant once this runs on an actual unattended schedule rather than being run manually.
- Reference data for the two properties is currently maintained as **two separate workbooks** with significant content overlap (room types, market segments, rate codes), rather than one shared file — a deliberate tradeoff made earlier in the project.

## Requirements

- Python 3.x, `pandas`, `numpy`, `openpyxl`
- Per-property reference/dimension workbook and standardized daily OTB files (see Steps 1-5 repos for how these are structured and validated)
