# SUNOHA CHECKLISTS APP — CLAUDE CODE PROMPT TEMPLATE
> Copy everything between the dashed lines and paste at the TOP of every Claude Code prompt.
> Then add your actual request below the "YOUR PROMPT HERE" section.

---

```
CONTEXT: Sunoha Checklist App — Enterprise-grade coffee production 
operations system. Used daily by production team and management.
Full chain: green bean procurement → roasting → grinding → packing 
→ order fulfillment.
Project location: C:\Users\Dell\Documents\Buz Docs\Application development\Checklists

MANDATORY PRE-READ — Before writing a single line of code:
1. Read docs/GUARDRAILS.md — all 14 rules are non-negotiable
2. Read docs/Sunoha-Checklists-PRD-v2.docx — full product spec
3. Read docs/Sunoha-Feature-Tracker.docx — what is built vs pending
4. Read scripts/test-inventory-writes.js — understand all test scenarios

STANDARDS:
- Zero bugs. Preserve ALL existing data and features.
- Read-merge-write on ALL sheet saves. Never overwrite entire sheets.
- Every destructive action writes to AuditLog before executing.
- All new sheets/columns must be backward compatible with existing data.
- Scalable: new checklist types and features will be added in future.
- NEVER declare anything fixed without running simulation and showing proof.
- NEVER use positional index to read sheet columns or response fields.
- NEVER write to sheets by position — always by column header name.
- NEVER create duplicate inventory ledger entries for same submission.
- NEVER skip is_deleted check when looping over UntaggedChecklists.

MANDATORY VERIFICATION after every change — show output as proof:
node scripts/test-inventory-writes.js
node --check server/google-apps-script.js
node -e "const s=require('fs').readFileSync('order-checklist-manager.jsx','utf8');let d=0;for(const c of s){if(c==='(')d++;if(c===')')d--;}console.log(d===0?'FRONTEND: OK':'FRONTEND: FAIL')"

Do NOT declare anything fixed without showing all 3 outputs passing.

=========================================
[YOUR PROMPT GOES HERE]
=========================================
```

---

## How to use this template

### Step 1 — Copy the block above
Copy everything between the triple backticks (` ``` `).

### Step 2 — Open Claude Code
In VS Code terminal, navigate to project folder and type:
```
cd "C:\Users\Dell\Documents\Buz Docs\Application development\Checklists"
claude
```

### Step 3 — Paste and add your request
Paste the template, then replace `[YOUR PROMPT GOES HERE]` with your actual request.

### Example of a complete prompt:

```
CONTEXT: Sunoha Checklist App — Enterprise-grade coffee production 
operations system...
[full template]
...
=========================================
Fix the auto-fill mapping label showing wrong question name.
The "Roast Date" question shows "Type of Bean" as target label.
Fix it to always show the current question's own name.
Only change what is needed. Do not touch anything else.
=========================================
```

---

## Rules for writing good prompts

| Rule | Good | Bad |
|---|---|---|
| Scope | "Fix only the auto-fill label bug" | "Fix everything" |
| Max fixes per prompt | 3 fixes maximum | 9 fixes at once |
| Be specific | "In EditChecklistView, the target label reads questions[0].text instead of questions[qi].text" | "The label is wrong" |
| Separate features from bugs | One prompt for bugs, separate prompt for new features | Mix bug fixes with new features |
| Include context | "The mapping is stored as question.linkedSource.autoFillMapping" | No context given |

---

## After every deployment run this sequence

```
1. Settings → Fix Ledger Alignment
2. Settings → Recalculate Inventory Balances  
3. Settings → Run Tests → must show 12/12
4. Fill one Green Beans QC → check ledger IN entry
5. Check inventory dashboard — no unexpected negatives
```

---

*Last updated: April 2026*
*Save this file to: docs/PROMPT_TEMPLATE.md in project root*
