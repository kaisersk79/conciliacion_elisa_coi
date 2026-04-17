# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Accounting reconciliation system that compares balances between **Elisa/Odoo** (ERP) and **COI** (internal chart of accounts). The goal is to detect discrepancies, unmatched accounts, and hierarchy validation errors. All scripts and data files live in the root (referred to as the "mensual" folder per convention).

## Execution Order

Run scripts sequentially — each depends on the previous output:

```bash
python3 libro_mayor_plano.py    # Input: libro_mayor_dic.xlsx  → Output: Reporte_Contable_Final.xlsx
python3 clean_coi.py            # Input: aux_coi_dic.xlsx       → Output: COI_Final_SumaCorrecta.xlsx
python3 conciliacion_coi.py     # Inputs: both outputs above   → Output: Analisis_Comparativo_*.xlsx
```

No build step, no dependencies to install beyond `openpyxl` and `pandas`.

## Architecture

### Data Flow

```
libro_mayor_dic.xlsx (Odoo raw) ──► libro_mayor_plano.py ──► Reporte_Contable_Final.xlsx
                                                                          │
aux_coi_dic.xlsx (COI raw) ──► clean_coi.py ──► COI_Final_SumaCorrecta.xlsx
                                                                          │
                                              conciliacion_coi.py ◄──────┘
                                                      │
                                              Analisis_Comparativo_*.xlsx
```

### Script Responsibilities

**`libro_mayor_plano.py`** — Processes Odoo general ledger:
- Builds 3-level hierarchy (Level 1: group like `101`, Level 2: sub-group like `107.05`, Level 3: leaf account)
- Reclassifies `205.xx` accounts containing "Samuel" / "Villa Rodríguez" into group `107.05` (Otros deudores diversos)
- Computes special line: `107 total − 107.05.01` (excludes merchandise sent but not invoiced from debtors)
- Uses `MAJOR_NAME_MAP` (~80 entries) to assign human-readable names to account codes

**`clean_coi.py`** — Processes COI chart of accounts:
- Parses unstructured format: lines matching `"Cuenta : XXXX-XXX-XXX ..."`
- Three hierarchy levels: **Madre Suprema** (`-000-000` suffix), **Padre** (has children), **Hoja** (leaf)
- Validates parent ↔ children sum consistency via a Check column (OK / DIF)
- Groups accounts by `Grupo` via `obtener_nombre_rubro()` (~80 mappings), summing only leaf accounts to avoid double-counting

**`conciliacion_coi.py`** — Reconciliation logic:
- `HEADER_MAP` (244 entries): maps Odoo section codes → COI account codes
- `VIRTUAL_COI_SUMS`: computes on-the-fly totals (e.g., `SUMA-CLIENTES-NACIONALES` = sum of `1150-001-000 + 1150-004-000 + ...`)
- Match fallback chain: direct HEADER_MAP → regex extraction from description → virtual sum name
- Tolerance: differences `< 0.01` are treated as OK (floating-point rounding)
- Statuses: `OK`, `DIFERENCIA`, `NO EN COI`, `NO EN COI POR ESTRUCTURA`, `NO EN ELISA`
- **Abuela (grandparent) check**: for accounts in `CHECK_ABUELAS_LIST` (23 key accounts), validates parent balance = sum of children with same 4-digit prefix

### Key Business Rules

- `107.05.06` and similar accounts without their own COI entry must aggregate into their parent group — same pattern as "PG sin cuenta propia" where all movements are lumped into a single catch-all account.
- COI accounts with balance but no Odoo match → `NO EN ELISA` (currently partially implemented; accounts found in COI but missing in Odoo must also be flagged systematically).
- **Parent sum check** (mamas/abuelas): if the sum of child accounts doesn't match the declared parent total, it signals a missing child account was not pulled from COI. This check is the primary way to detect omissions.
- The `recibo` + `pendiente` filter skips those rows intentionally (not real balances).

## Known Issues / Pending Improvements

1. **Cuentas en COI pero no en Elisa**: `conciliacion_coi.py` currently flags Elisa accounts missing in COI but does **not** reliably flag COI accounts that have a balance and have no Odoo counterpart. The orphaned-account detection for this direction needs to be completed.
2. **Check de mamas y abuelas**: The abuela check in `CHECK_ABUELAS_LIST` must also verify that the sum of leaf balances matches the declared parent total in COI — this catches cases where a child account exists in COI but was never added to `HEADER_MAP` or `VIRTUAL_COI_SUMS`.
3. **Cuenta `107.05.06`**: Like the "PG sin cuenta propia" grouping, any COI accounts that match this Odoo code by name (even without an explicit COI entry) should be bucketed into `107.05` rather than reported as unmatched.

## Configuration Hotspots

When a new account appears or mappings change, update these locations:

| Location | Purpose |
|---|---|
| `conciliacion_coi.py` → `HEADER_MAP` | Odoo code → COI code mappings |
| `conciliacion_coi.py` → `VIRTUAL_COI_SUMS` | Computed multi-account totals |
| `conciliacion_coi.py` → `CHECK_ABUELAS_LIST` | Accounts subject to parent-sum validation |
| `clean_coi.py` → `obtener_nombre_rubro()` | COI code → business category |
| `libro_mayor_plano.py` → `MAJOR_NAME_MAP` | Odoo code → display name |
