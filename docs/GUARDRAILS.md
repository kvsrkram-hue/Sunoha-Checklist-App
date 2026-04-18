# SUNOHA CHECKLISTS APP — DEVELOPMENT GUARD RAILS
> Version 1.0 | April 2026
> Read this file BEFORE writing any code. These rules are non-negotiable.
> Violating any guard rail is a critical bug regardless of how minor it seems.

---

## GUARD RAIL 1 — SINGLE SOURCE OF TRUTH

Every data point has exactly ONE place it is read from. Never compute the same value from two different sources.

| Data Type | Single Source | Never Use |
|---|---|---|
| Remaining quantity | `getAllocatedQuantityForAutoId()` → QuantityAllocations table | `getUsedQuantity()`, response sheet sums, `tagged_quantity` subtraction |
| Inventory balance | `InventoryItems.current_stock` (updated by `createInventoryTransaction`) | Manual ledger sum, `balance_after` on latest row |
| Response data | `UntaggedChecklists.responses` JSON column | Per-checklist response sheet tabs (those are for human audit only) |
| Approval status | `UntaggedChecklists.is_approved` column | Re-evaluating response fields |
| Deletion status | `UntaggedChecklists.is_deleted` via `isDeleted()` helper | Any other field or flag |
| Field values in responses | Match by `questionText` string | Match by position index (e.g. `responses[2]`) |

---

## GUARD RAIL 2 — INVENTORY SYNC (CRITICAL)

The inventory balance and untagged quantities are TWO VIEWS of the SAME data. They must ALWAYS be in sync.

### The Complete Flow:

```
GREEN BEANS QC SUBMITTED (qty = 100kg):
  Inventory:  Green Beans item       +100kg  (IN ledger entry)
  Untagged:   New entry appears       total=100, tagged=0, remaining=100
  Dashboard:  Green Beans total       +100kg

ROASTED BEANS QC SUBMITTED (input=80kg, output=75kg):
  Inventory:  Green Beans item        -80kg  (OUT ledger entry)
  Inventory:  Roasted Beans item      +75kg  (IN ledger entry)
  Untagged:   Green Bean entry        tagged += 80, remaining -= 80
  Untagged:   New RB entry appears    total=75, tagged=0, remaining=75
  Dashboard:  Green Beans             -80kg, Roasted Beans +75kg

GRINDING SUBMITTED (input=70kg, output=68kg):
  Inventory:  Roasted Beans item      -70kg  (OUT ledger entry)
  Inventory:  Packing Items           +68kg  (IN ledger entry)
  Untagged:   Roasted Bean entry      tagged += 70, remaining -= 70
  Untagged:   New grinding entry      total=68, tagged=0, remaining=68
  Dashboard:  Roasted Beans -70kg, Packing Items +68kg

ORDER DELIVERED:
  Inventory:  Packing Items           -dispatched qty (OUT ledger entry)
  Untagged:   Grinding entry          tagged += dispatched qty
  Dashboard:  Packing Items decreases

EDIT RULE:
  Old qty=100, New qty=120 → create IN entry for +20
  Old qty=100, New qty=80  → create OUT entry for -20
  Validation: new qty CANNOT be less than sum of downstream tagged qty
```

### Mandatory Checks:
- NEVER write a ledger entry without updating `current_stock` on `InventoryItems`
- NEVER update `current_stock` without writing a ledger entry
- `current_stock` MUST be updated atomically with every ledger write
- Every IN ledger entry MUST create or correspond to an untagged entry
- Every OUT ledger entry MUST reduce `tagged_quantity` on source entry

---

## GUARD RAIL 3 — CHECKLISTS THAT AFFECT INVENTORY

Only these checklists affect inventory. All others have NO inventory impact unless explicitly configured by admin.

| Checklist ID | IN | OUT | total_quantity |
|---|---|---|---|
| `ck_green_beans` | Green Beans + qty received | — | qty received |
| `ck_roasted_beans` | Roasted Beans + sum(outputQty) | Green Beans - sum(inputQty) | sum(outputQty) |
| `ck_grinding` | Packing Items + Total Net weight | Roasted Beans - Input Weight | Total Net weight |
| Order Delivered | — | Packing Items - dispatched qty | — |

**NO inventory impact:**
- `ck_sample_qc` — sample reference only
- `ck_sample_retention` — not configured
- `ck_chicory` — not configured
- `ck_blending` — not configured
- Any new checklist — unless admin explicitly configures `inventoryLink` on a question

---

## GUARD RAIL 4 — NO POSITIONAL INDEX READS

NEVER read data by position number. ALWAYS read by name.

```javascript
// ❌ WRONG — position-based, breaks when template reorders
const beanType = responses[2];
const qty = row[5];

// ✅ CORRECT — name-based, immune to reordering
const beanType = responses.find(r => r.questionText === "Type of Beans")?.response;
const qty = rowObject.quantity; // via getRows() which maps by header name
```

**This applies to:**
- Reading response fields → always match by `questionText`
- Reading sheet columns → always use `getRows()` which maps by header
- Writing sheet columns → always use `appendToSheet()` or `updateSheetRow()` which write by header name
- NEVER use `row[3]` or `sheet.getRange(r, 4)` with hardcoded column numbers

---

## GUARD RAIL 5 — NO SILENT FAILURES

Every error must be visible and traceable. Never swallow exceptions.

```javascript
// ❌ WRONG — silent failure
function createInventoryTransaction(itemId, qty) {
  var item = findItem(itemId);
  if (!item) return; // silently does nothing
}

// ✅ CORRECT — logged failure
function createInventoryTransaction(itemId, qty) {
  var item = findItem(itemId);
  if (!item) {
    writeAuditLog({ action: "warn", notes: "Item not found: " + itemId });
    return { warning: "item_not_found", itemId: itemId };
  }
}
```

**Rules:**
- If `createInventoryTransaction` cannot find item: log to AuditLog and return `{ warning: "item_not_found" }`
- If quantity validation fails: return explicit error message with exact amounts
- If reversal already exists: return `{ skipped: true, reason: "Reversal already exists" }` — do NOT create duplicate
- Every warning must appear in AuditLog tab in Google Sheets

---

## GUARD RAIL 6 — QUANTITY BOUNDARIES

These checks MUST run on every submit, edit, and tag operation. No exceptions.

```
tagged_quantity     ≤  total_quantity          (ALWAYS)
remaining_quantity  ≥  0                       (show 0 as minimum, never negative display)
edit_quantity       ≥  sum(downstream tagged)  (cannot reduce below committed)
inputQty per batch  ≤  remaining on source     (cannot use more than available)
outputQty per batch ≤  inputQty                (cannot produce more than input)
```

**On negative inventory balance:**
- Do NOT block the OUT transaction (production reality — discrepancies happen)
- DO write the ledger entry
- DO update `current_stock` even if it goes negative
- DO log warning to AuditLog
- DO show negative items in RED on inventory dashboard
- Never silently ignore

---

## GUARD RAIL 7 — SHEET WRITES

ALWAYS write to sheets by column header name, never by position.

```javascript
// ❌ WRONG — position-based, causes column misalignment
sheet.getRange(row, 4).setValue(categoryValue); // position 4 might not be "category"

// ✅ CORRECT — header-based
appendToSheet(SHEETS.INVENTORY_LEDGER, {
  id: ledgerId,
  item_id: itemId,
  category: categoryValue,  // appendToSheet maps by header name
  quantity: qty
});
```

**Rules:**
- ALWAYS call `ensureSheetHasAllColumns(sheetName)` before writing to any sheet
- ALWAYS append new columns to the END of existing columns — NEVER insert in middle
- ALWAYS use `appendToSheet()` or `updateSheetRow()` — both write by header name
- NEVER use `sheet.appendRow([v1, v2, v3])` with positional array
- NEVER use `sheet.getRange(r, colNumber).setValue()` with hardcoded column numbers

---

## GUARD RAIL 8 — BACKWARD COMPATIBILITY

Never break existing data or API contracts.

- NEVER remove existing API actions from `doGet`/`doPost` routing switch
- NEVER change existing API response format (only ADD new fields)
- NEVER physically delete sheet rows — always use `is_deleted=true`
- NEVER rename Google Sheet tab names
- NEVER rename existing column headers
- NEVER change the meaning of an existing column
- Old submissions (single-batch RB format) MUST still display correctly
- New code paths MUST detect old format and handle gracefully

---

## GUARD RAIL 9 — NO DUPLICATE WRITES

Every submission must create exactly the right number of ledger entries — no more, no less.

| Checklist | Expected ledger entries |
|---|---|
| Green Beans QC | Exactly 1 IN entry |
| Roasted Beans QC — 1 batch | Exactly 2 entries (1 OUT Green Beans + 1 IN Roasted Beans) |
| Roasted Beans QC — N batches | Exactly 2N entries (N OUT + N IN) |
| Grinding & Packing | Exactly 2 entries (1 OUT Roasted Beans + 1 IN Packing Items) |
| Sample QC | Exactly 0 entries |

**Before any submit handler writes inventory:**
- Check: is there already a ledger entry for this `reference_id`?
- If yes: do NOT write again (idempotency check)
- This prevents double-writes from retries or duplicate submissions

**Before any reversal:**
- Check: does a `[REVERSAL]` entry already exist for this `reference_id`?
- If yes: return `{ skipped: true }` — do NOT create duplicate reversal

---

## GUARD RAIL 10 — AUTO-FILL MAPPINGS

- Each question has AT MOST one auto-fill mapping
- `autoFillMapping` is a single object `{ sourceFieldIdx, readOnly }` — NEVER an array
- Target label in mapping UI ALWAYS shows the CURRENT question's own name (`questions[qi].text`)
- NEVER show another question's name as the target label
- Mapping UI: show mapping row OR add button — NEVER both at same time
- If mapping exists: show "Pull value from: [dropdown] + Read-only checkbox + Remove ×"
- If no mapping: show "+ Auto-fill from [source]" button only

---

## GUARD RAIL 11 — IS_DELETED FILTERING

Every function that reads `UntaggedChecklists` rows MUST filter deleted entries.

```javascript
// ✅ Required pattern in every loop over UntaggedChecklists
for (var i = 0; i < rows.length; i++) {
  if (isDeleted(rows[i])) continue;  // THIS LINE IS MANDATORY
  // ... process row
}
```

**Functions that MUST have this check:**
- `handleGetUntagged`
- `handleGetUntaggedResponse`
- `getSubmissionByAutoId`
- `getApprovedEntriesForChecklist`
- `handleTagUntagged`
- `handleEditUntaggedResponse`
- `handleCreateAllocation`
- `findSubmissionByAutoId`
- `handleTagChecklistToStage`
- Any NEW function that loops over UntaggedChecklists

---

## GUARD RAIL 12 — TEST VERIFICATION (MANDATORY)

NEVER declare anything fixed without runtime proof.

```bash
# Run after EVERY code change — no exceptions
node scripts/test-inventory-writes.js

# All scenarios must show PASS
# If ANY fail: debug and fix before moving on
# NEVER skip this step

# Also syntax check both files
node --check server/google-apps-script.js
node -e "const s=require('fs').readFileSync('order-checklist-manager.jsx','utf8');let d=0;for(const c of s){if(c==='(')d++;if(c===')')d--;}console.log(d===0?'FRONTEND: OK':'FRONTEND: FAIL')"
```

**If simulation passes but live tests fail:**
- The gap is between mock data and real Google Sheets data
- Add diagnostic logging to find the exact difference
- NEVER assume the fix works without live test confirmation

---

## GUARD RAIL 13 — REFERENCE ID FORMAT

All inventory ledger entries MUST use the `UntaggedChecklists` row id (`ut_xxx` format) as `reference_id`.

```javascript
// ❌ WRONG — using auto_id
createInventoryTransaction(itemId, qty, "untagged", "GB-130426-ACAB-003", ...);

// ✅ CORRECT — using ut_ id
createInventoryTransaction(itemId, qty, "untagged", untaggedRowId, ...);
// untaggedRowId = the "id" column value from UntaggedChecklists (ut_xxx format)
```

This ensures every ledger entry can be traced back to its source submission.

---

## GUARD RAIL 14 — DEPLOYMENT CHECKLIST

Run through this after every deployment before sharing with team:

```
□ node scripts/test-inventory-writes.js → all PASS
□ Settings → Run Tests → 12/12 PASS
□ Submit one Green Beans QC → check ledger IN entry created
□ Submit one Roasted Beans QC → check ledger OUT+IN entries created
□ Check inventory dashboard — no unexpected negatives
□ Check dates show as DD-MM-YYYY not ISO format
□ Check delivered orders NOT in tag dropdowns
□ Check edit mode shows correct values in correct fields
```

---

## QUICK REFERENCE — WHAT EACH FUNCTION DOES

| Function | Purpose | Guard Rails |
|---|---|---|
| `createInventoryTransaction` | Writes ledger entry + updates current_stock | GR1, GR2, GR5, GR6 |
| `applyLegacyInventoryForChecklist` | Routes inventory writes for non-inventoryLink templates | GR3, GR9 |
| `applyRoastBatchInventory` | Writes per-batch OUT+IN for multi-batch roasting | GR3, GR9 |
| `validateRoastBatches` | Checks inputQty ≤ remaining before roasting | GR6 |
| `validateQuantityEdit` | Checks edit qty ≥ downstream committed | GR6 |
| `reverseInventoryLedgerForRef` | Creates offsetting entries on delete | GR5, GR9 |
| `getAllocatedQuantityForAutoId` | SSOT for remaining quantity | GR1 |
| `appendToSheet` | Writes new row by header name | GR7 |
| `updateSheetRow` | Updates existing row by header name | GR7 |
| `isDeleted` | Checks is_deleted flag | GR11 |
| `getRows` | Reads sheet rows as named objects | GR4 |

---

*Last updated: April 2026 | Update this file whenever a new rule is established*
