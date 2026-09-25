# Multi-Initiative_resource_allocation_excel
# Multi-Initiative Resource Allocation & Prioritization Tracker

An Excel-based operations tool that tracks tasks across multiple concurrent initiatives sharing a common resource pool (Budget, Staff-Hours, Vendor Slots), flags bottlenecks automatically, and recommends which tasks should get resources first when capacity is constrained.

## Problem

Teams running several initiatives at once (a product launch, a hiring drive, a corporate event, etc.) usually pull from the same limited pool of budget, people, and vendors. Without a shared view, it's easy to over-commit one resource while another sits underused — and to lose track of which tasks matter most when something has to give.

## What it does

- **Task log** — 24 tasks across 3 initiatives, each scored on Impact × Urgency to produce a live Priority Score and Rank
- **Shared capacity model** — a single quarter-level pool of Budget, Staff-Hours, and Vendor Slots that all three initiatives draw from
- **Allocation tracker** — `SUMIFS`-driven utilization by resource and by initiative, with automatic flags: *Over-allocated / Near Capacity / Healthy*
- **Prioritization logic** — a tie-broken ranking formula (`RANK` + `COUNTIF`) surfaces the top 5 tasks that should be resourced first under a constraint
- **Dashboard** — utilization, task-status, and budget-by-initiative charts, plus headline KPIs

## Key finding (sample data)

| Resource | Allocated | Capacity | Utilization | Status |
|---|---|---|---|---|
| Budget | ₹8,24,000 | ₹5,00,000 | 165% | 🔴 Over-allocated |
| Staff Hours | 507 | 550 | 92% | 🟡 Near Capacity |
| Vendor Slots | 27 | 35 | 77% | 🟢 Healthy |

Budget is the binding constraint, not people or vendors — so the priority ranking (Impact × Urgency) is what decides which tasks get funded first, not just who's free.

## How it's built

- **Tool:** Microsoft Excel (openpyxl for generation, LibreOffice-verified formulas)
- **Core formulas:** `SUMIFS`, `INDEX`/`MATCH`, `RANK` + `COUNTIF` tiebreak, `IF`/nested `IF` flagging, native Excel charts
- **Sheets:** `Tasks` → `Resource Capacity` → `Allocation Tracker` → `Dashboard`
- All calculated cells are live formulas (not hardcoded), so editing any task or capacity input recalculates the whole tracker.

## File

- `Resource_Allocation_Tracker.xlsx` — the full workbook
