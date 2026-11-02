---
name: fiori-requirement-analysis
description: Extract grounded SAP Fiori requirements from uploaded TDD, BRD, FRD, functional specifications, fit-gap notes, spreadsheets, or available Context Studio documents, and generate completed Excel analytical and/or transactional Fiori input templates. Preserve the supplied template sheets and instructions, map source information to structured rows, mark every missing field NA, and report coverage and gaps. Use for client-independent Fiori analytical dashboards and transactional applications.
---

# SAP Fiori Requirements to Excel Input Templates

## Purpose and output contract

Transform supplied input document(s) into **populated copies of the two bundled Excel input templates** (NOT a Word functional specification):

- `assets/analytical_template.xlsx` -> `<subject>_Analytical_Fiori_Input.xlsx` (16 worksheets)
- `assets/transactional_template.xlsx` -> `<subject>_Transactional_Fiori_Input.xlsx` (17 worksheets)

For clearly analytical requirements, output the analytical workbook; for clearly transactional requirements, output the transactional workbook. For mixed/hybrid requirements or when the user explicitly requests both, output **both**, partitioning requirements appropriately and repeating shared, applicable information in each. When application type cannot be determined from the content, produce both and indicate the ambiguity in the summary. Do not make either workbook artificially contain requirements of the wrong type. Do not generate an additional functional specification unless separately requested.

Preserve all template worksheet names, header labels, styles, validations, and `START_HERE` / `SHEET_INDEX` instructions. Use *actual rows* in each relevant worksheet. An unfilled source value must be written literally as **`NA`**, not blank, TBD, unknown, not found, N/A, a guessed value, or an invented technical object. The rule applies to every column of each populated data row, including identifiers where none is supplied (but see the identifier rules below). Blank unused rows may remain blank; avoid hundreds of purposeless `NA` rows. Remove existing illustrative/sample data in row 5 and replace it with real extracted information where available, so examples are not mistaken for sourced facts.

## Step 1: Gather documents

Accept user-uploaded PDF, DOCX, XLSX, CSV, TXT, MD and archives containing such sources. Prefer the user-supplied files. A connected Context Studio MCP source may also be used when the user indicates that context; do not require that connector when an attachment is available.

For Context Studio, if available, discover contexts and metadata with the available context-broker tools; query Ready documents relevant to requirements, architecture, KPIs, mapping, UX, roles and reusable content. Use retrieval iteratively, especially for segmented documents. Record skipped or unavailable sources. If connectors are unavailable, proceed with attachments rather than failing.

Read full relevant sections/tables, not just snippets. Use source document and section/page/table/row references where available. Ignore client identity or names unless necessary as a functional requirement; do not copy organization-specific names into the generic output. Never copy secrets, passwords, private endpoints, personal information, or credentials.

## Step 2: Classify scope and extract an evidence ledger

Determine whether the information relates to Analytical, Transactional, or Both:

- Analytical: KPIs, measures, aggregates, dimensions, visualization, reporting filters, drilldown, reporting refresh and analytical data sources.
- Transactional: business object lifecycle, create/change/display, form fields, buttons, actions, workflows, validations, backend operations, approvals and status transitions.
- Shared: business context, roles, requirements, source inventory, field mappings, existing assets, decisions, stories, acceptance tests, traceability.

Create an internal evidence ledger for each extractable fact: `source_file`, `section/page` (or `NA`), `verbatim_or_close_paraphrase`, `target_sheet`, `target_column`, `requirement_reference` (or `NA`), `confidence`. Do not treat common SAP conventions, examples, or proposed implementation ideas as source facts.

Explicitly extract:
1. Project context, objectives, scope and process.
2. Requirements, priority and scope status.
3. Roles, authorization requirements and responsibilities.
4. KPIs, formulas, target and units; measures, dimensions, hierarchies and analytical UI.
5. Transactional objects, screens, actions, workflows, lifecycle, validations and interfaces.
6. Source systems, tables, CDS views, OData/API objects, key fields and refresh/integration details.
7. Source-to-target field mapping, transformation and derivation rules.
8. Reusable existing Fiori content, evidence for reuse, fit and limitations.
9. Gaps, constraints, dependencies, open decisions, user stories and testable acceptance criteria.

If a source refers to an object without its technical identifier, write the stated description in the suitable column and `NA` in the technical-ID column. Never invent KPI formulas, table names, CDS entities, OData APIs, roles, app IDs or acceptance thresholds.

## Step 3: Map evidence into the appropriate workbook

Use the bundled Excel sheets as the authoritative schema. The header labels in **row 4** of each data worksheet are the canonical field names. `START_HERE` and `SHEET_INDEX` are instructional sheets and must not be populated with data rows.

Shared worksheets in each workbook:
- `01_Project`: business overview, goals, ownership and scope
- `02_Requirements`: requirement descriptions, priority and context
- `03_Roles`: user roles, data scope, allowed actions and authorizations
- `04_Data_Sources`: systems, technical objects, entities and keys
- `05_Mappings`: target properties, origin, transformation, mappings
- `06_Reuse_Inventory`: existing reports, Fiori applications, CDS views, APIs and evidence of reuse
- `07_Gaps_Decisions`: *documented* gaps, decisions, risks, assumptions or missing critical requirement details
- `08_User_Stories`: actual requirements reformulated as stories **only if the role/capability/outcome are documented**; other missing clauses `NA`
- `09_Acceptance`: documented verifiable conditions and test scenarios; do not fabricate test criteria
- `10_Traceability`: links between documented IDs (and source-derived relationships)

Analytical-specific:
- `A1_KPIs`: measures, business definitions, formulas, aggregation, thresholds
- `A2_Dimensions`: labels, source elements, hierarchy, filters and granularity
- `A3_Analytics_UI`: tiles, charts, views, prompts, sorting and drilldown
- `A4_Analytics_Rules`: inclusions, exclusions, currency and time rules

Transactional-specific:
- `T1_Business_Objects`: object keys, relationships and lifecycle
- `T2_Screens_Actions`: UI screens, fields, actions, defaults and navigation
- `T3_Process_Workflow`: triggers, authorization and status transitions
- `T4_Interfaces`: operations, payloads, APIs and endpoints
- `T5_Validation_Audit`: validation conditions, warnings, access and audit

Build one record per entity/requirement/KPI/story/mapping/etc., not one record per input document. Preserve precise links across sheets. Keep explicit source IDs unchanged. If no ID exists, generate stable local reference IDs such as `REQ-001`, `KPI-001`, `SRC-001`, `STORY-001`, annotate in the summary that IDs were generated solely for workbook traceability, and never imply they existed in the source. For other unknown fields—including related IDs when the association is unsupported—use `NA`.

Do not infer missing features from the fact that a design document uses a certain technology. A field that is **not applicable** to a supported record is also `NA` by the user's convention. Do not create artificial gap rows for every `NA` cell; capture high-impact missing information as concise missing-input gap records, clearly labeled as identified documentation gaps, not approved business decisions.

## Step 4: Populate Excel deterministically

Use the bundled `scripts/populate_templates.py` helper with an evidence-grounded JSON payload. The helper loads the original workbook, retains formatting/instructions, clears illustrative row 5, fills cells in each populated row with supplied values or the literal `NA`, and saves a copy. Its input JSON uses exactly the worksheet names and row-4 header strings. Inspect the workbook's headers before composing payload keys; do not silently introduce new columns or sheets.

Example input JSON structure:

```json
{
  "analytical": {
    "01_Project": [{"Project ID": "NA", "Business process": "Order to cash"}],
    "A1_KPIs": [{"KPI ID": "KPI-001", "KPI / measure name": "Billing amount"}]
  },
  "transactional": {
    "T1_Business_Objects": [{"Object ID": "OBJ-001", "Business object name": "Billing document"}]
  }
}
```

Run e.g. `python scripts/populate_templates.py --input extracted.json --output-dir output --subject Billing --mode both`. For a single type use `--mode analytical` or `--mode transactional`.

For source materials without enough information to populate a specific sheet, leave the corresponding sheet's unused body blank; do not fabricate records. Preserve instructions and headings so the user can manually complete it later. For every **existing populated row**, all columns must contain a factual value or `NA`.

## Step 5: Validate and deliver

Verify:
- Required workbook(s) exist and open with openpyxl / equivalent.
- Correct 16/17-sheet lists and order match assets.
- All data rows are filled across all defined header columns with no blanks.
- No original example rows were left as data.
- IDs and referenced IDs are consistent; unresolved links are `NA`.
- No content not grounded in supplied documents is presented as confirmed.
- No secrets or client-specific branding unintentionally carried over.
- Formula and formatting integrity preserved.

Provide downloadable Excel workbooks (and a ZIP containing both when both requested), plus a concise report of source files used, classification, row counts per sheet, missing high-priority inputs, skipped sources and generated traceability IDs. The templates are a structured extraction artifact; do not call them a validated SAP design or assert existence of reusable SAP content without confirmation.
