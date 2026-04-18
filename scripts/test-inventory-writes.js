// Local simulation of inventory write paths from google-apps-script.js
// Tests the 4 failing scenarios without needing Google Sheets

// ═══════════════════════════════════════════════════════════════
// MOCK GOOGLE SHEETS ENVIRONMENT
// ═══════════════════════════════════════════════════════════════

const SHEETS = {
  INVENTORY_ITEMS: "InventoryItems",
  INVENTORY_LEDGER: "InventoryLedger",
  ROAST_CLASSIFICATIONS: "RoastClassifications",
  UNTAGGED_CHECKLISTS: "UntaggedChecklists",
};

const HEADERS = {
  InventoryItems: ["id", "category", "name", "unit", "opening_stock", "current_stock", "min_stock_alert", "created_at", "is_active", "abbreviation", "equivalent_items", "classification_id"],
  InventoryLedger: ["id", "item_id", "item_name", "category", "date", "type", "quantity", "balance_after", "reference_type", "reference_id", "notes", "done_by", "created_at", "question_index", "classification_id"],
};

// In-memory sheet data: { sheetName: [ [header_row], [data_row1], ... ] }
const sheetData = {};
const writeLog = [];

function resetSheets() {
  writeLog.length = 0;
  // InventoryItems with test data
  sheetData[SHEETS.INVENTORY_ITEMS] = [
    HEADERS.InventoryItems,
    ["inv_test_gb", "Green Beans", "TEST_GB_ITEM", "kg", 0, 0, 0, "2026-01-01", "true", "TGBI",
      JSON.stringify([{ category: "Roasted Beans", itemId: "inv_test_rb" }, { category: "Packing Items", itemId: "inv_test_pk" }]), ""],
    ["inv_test_rb", "Roasted Beans", "TEST_RB_ITEM", "kg", 0, 0, 0, "2026-01-01", "true", "TRBI",
      JSON.stringify([{ category: "Green Beans", itemId: "inv_test_gb" }, { category: "Packing Items", itemId: "inv_test_pk" }]), ""],
    ["inv_test_pk", "Packing Items", "TEST_PK_ITEM", "kg", 0, 0, 0, "2026-01-01", "true", "TPKI",
      JSON.stringify([{ category: "Green Beans", itemId: "inv_test_gb" }, { category: "Roasted Beans", itemId: "inv_test_rb" }]), ""],
  ];
  sheetData[SHEETS.INVENTORY_LEDGER] = [HEADERS.InventoryLedger];
}

// ═══════════════════════════════════════════════════════════════
// MOCK FUNCTIONS (matching google-apps-script.js behavior)
// ═══════════════════════════════════════════════════════════════

const _rowsCache = {};
function clearRowsCache() { for (const k in _rowsCache) delete _rowsCache[k]; }
function invalidateCache(name) { delete _rowsCache[name]; }

function getRows(sheetName) {
  if (_rowsCache[sheetName]) return _rowsCache[sheetName];
  const data = sheetData[sheetName];
  if (!data || data.length <= 1) { _rowsCache[sheetName] = []; return []; }
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    for (let j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j] !== undefined ? data[i][j] : "";
    rows.push(obj);
  }
  _rowsCache[sheetName] = rows;
  return rows;
}

function findRowIndex(sheetName, id) {
  const data = sheetData[sheetName];
  if (!data) return -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-based sheet row
  }
  return -1;
}

function appendToSheet(sheetName, obj) {
  const headers = HEADERS[sheetName];
  if (!headers) { console.log("  !! appendToSheet: no headers for " + sheetName); return; }
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : "");
  if (!sheetData[sheetName]) sheetData[sheetName] = [headers];
  sheetData[sheetName].push(row);
  invalidateCache(sheetName);
  writeLog.push({ action: "append", sheet: sheetName, obj: { ...obj } });
}

function updateSheetRow(sheetName, rowIndex, obj) {
  const headers = HEADERS[sheetName];
  if (!headers) return;
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : "");
  if (sheetData[sheetName] && sheetData[sheetName][rowIndex - 1]) {
    sheetData[sheetName][rowIndex - 1] = row;
  }
  invalidateCache(sheetName);
  writeLog.push({ action: "update", sheet: sheetName, row: rowIndex, id: obj.id });
}

let _nextIdCounter = 1000;
function nextId() { return String(++_nextIdCounter); }

function safeParseJSON(str, fallback) {
  try { if (typeof str === "string" && str.length > 0) return JSON.parse(str); return fallback; } catch(e) { return fallback; }
}

const Logger = { log: function(msg) { console.log("  [LOG] " + msg); } };

// ═══════════════════════════════════════════════════════════════
// EXTRACTED FUNCTIONS (from google-apps-script.js)
// ═══════════════════════════════════════════════════════════════

function findInventoryItemForCategory(sourceRef, targetCategory) {
  if (!sourceRef) return { item: null, warning: "No source inventory reference provided" };
  var items = getRows(SHEETS.INVENTORY_ITEMS);
  var srcStr = String(sourceRef).trim();
  var source = null;
  for (var i = 0; i < items.length; i++) {
    if (String(items[i].id) === srcStr || String(items[i].name) === srcStr) {
      source = items[i]; break;
    }
  }
  if (!source) return { item: null, warning: "Inventory item '" + sourceRef + "' not found" };
  if (!targetCategory || String(source.category) === String(targetCategory)) {
    return { item: source, isEquivalent: false, warning: "" };
  }
  var eqList = safeParseJSON(source.equivalent_items, []);
  for (var e = 0; e < eqList.length; e++) {
    if (String(eqList[e].category) === String(targetCategory) && eqList[e].itemId) {
      for (var j = 0; j < items.length; j++) {
        if (String(items[j].id) === String(eqList[e].itemId)) {
          return { item: items[j], isEquivalent: true, warning: "" };
        }
      }
    }
  }
  return { item: source, isEquivalent: false, warning: "[WARN] No equivalent" };
}

function createInventoryTransaction(itemId, type, quantity, refType, refId, notes, doneBy, questionIndex, classificationId) {
  var idx = findRowIndex(SHEETS.INVENTORY_ITEMS, itemId);
  if (idx < 0) {
    Logger.log("⚠ createInventoryTransaction: item '" + itemId + "' not found — SKIPPED");
    return { warning: "item not found: " + itemId };
  }
  var rows = getRows(SHEETS.INVENTORY_ITEMS);
  var item = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(itemId)) { item = rows[i]; break; } }
  if (!item) return { warning: "item row not found: " + itemId };

  var currentStock = parseFloat(item.current_stock) || 0;
  var newStock = type === "IN" ? currentStock + quantity : currentStock - quantity;
  item.current_stock = newStock;
  updateSheetRow(SHEETS.INVENTORY_ITEMS, idx, item);
  invalidateCache(SHEETS.INVENTORY_ITEMS);

  appendToSheet(SHEETS.INVENTORY_LEDGER, {
    id: "led_" + nextId(), item_id: itemId, item_name: item.name,
    category: item.category || "",
    date: new Date().toISOString().split("T")[0], type: type,
    quantity: (type === "OUT" ? -quantity : quantity), balance_after: newStock,
    reference_type: refType || "manual", reference_id: refId || "",
    notes: notes || "", done_by: doneBy || "", created_at: new Date().toISOString(),
    question_index: (questionIndex === undefined || questionIndex === null || questionIndex === "") ? "" : String(questionIndex),
    classification_id: classificationId || item.classification_id || "",
  });
}

function applyLegacyInventoryForChecklist(ck, respMap, refType, refId, person, fallbackInItemId, fallbackOutItemId, isEdit, grindClassificationId) {
  if (!ck) return;
  var suffix = isEdit ? " (edited)" : "";
  function readField(fieldName) {
    if (respMap[fieldName] !== undefined && respMap[fieldName] !== "") return respMap[fieldName];
    if (ck && ck.questions) {
      for (var fi = 0; fi < ck.questions.length; fi++) {
        if (ck.questions[fi].text === fieldName && respMap[fi] !== undefined && respMap[fi] !== "") return respMap[fi];
      }
    }
    return respMap[fieldName] || "";
  }
  function readQty(fieldName) {
    var raw = readField(fieldName);
    var v = parseFloat(raw);
    if (isNaN(v) || v === 0) Logger.log("readQty('" + fieldName + "'): raw='" + raw + "' parsed=" + v);
    return isNaN(v) ? 0 : v;
  }

  if (ck.id === "ck_green_beans") {
    var qtyReceived = readQty("Quantity received");
    if (qtyReceived <= 0) { Logger.log("ck_green_beans: qtyReceived <= 0, returning"); return; }
    var gbRef = readField("Type of Beans") || fallbackInItemId;
    Logger.log("ck_green_beans: gbRef = '" + gbRef + "'");
    var gbResolved = findInventoryItemForCategory(gbRef, "Green Beans");
    Logger.log("ck_green_beans: resolved item = " + (gbResolved.item ? gbResolved.item.id : "NULL") + " warning=" + (gbResolved.warning || "none"));
    if (gbResolved.item) {
      createInventoryTransaction(gbResolved.item.id, "IN", qtyReceived, refType, refId, "Green Bean shipment received" + suffix, person);
    }
    return;
  }

  if (ck.id === "ck_grinding") {
    var qIn = readQty("Quantity input");
    var qOut = readQty("Quantity output");
    var netWeight = readQty("Total Net weight");
    if (qIn <= 0 && netWeight > 0) qIn = netWeight;
    if (qOut <= 0 && netWeight > 0) qOut = netWeight;

    Logger.log("ck_grinding: qIn=" + qIn + " qOut=" + qOut + " netWeight=" + netWeight);

    // For test: skip the lookupChecklist/findSubmissionByAutoId chain.
    // Use fallback items directly.
    var roastAutoId = readField("Roast ID") || "";
    var beanRefG = "";

    // In real code this calls lookupChecklist + findSubmissionByAutoId.
    // For simulation, resolve directly from the mock data.
    if (roastAutoId && mockSubmissions[roastAutoId]) {
      beanRefG = mockSubmissions[roastAutoId].beanRef || "";
    }

    Logger.log("ck_grinding: roastAutoId='" + roastAutoId + "' beanRefG='" + beanRefG + "'");

    if (qIn > 0) {
      var rbOut = beanRefG
        ? findInventoryItemForCategory(beanRefG, "Roasted Beans")
        : (fallbackInItemId ? findInventoryItemForCategory(fallbackInItemId, "Roasted Beans") : { item: null, warning: "No Roast ID or fallback" });
      Logger.log("ck_grinding OUT: resolved = " + (rbOut.item ? rbOut.item.id : "NULL"));
      if (rbOut.item) createInventoryTransaction(rbOut.item.id, "OUT", qIn, refType, refId, "Used for grinding" + suffix, person);
    }
    if (qOut > 0) {
      var pkIn = beanRefG
        ? findInventoryItemForCategory(beanRefG, "Packing Items")
        : (fallbackOutItemId ? findInventoryItemForCategory(fallbackOutItemId, "Packing Items") : { item: null, warning: "No fallback" });
      Logger.log("ck_grinding IN: resolved = " + (pkIn.item ? pkIn.item.id : "NULL"));
      if (pkIn.item) createInventoryTransaction(pkIn.item.id, "IN", qOut, refType, refId, "Packed goods produced" + suffix, person, "", grindClassificationId || "");
    }
    return;
  }
}

function applyRoastBatchInventory(processed, refType, refId, person) {
  for (var i = 0; i < processed.length; i++) {
    var p = processed[i];
    Logger.log("applyRoastBatchInventory batch " + (i+1) + ": gbItemId=" + p.greenBeanItemId + " rbItemId=" + p.roastedBeanItemId + " in=" + p.inputQty + " out=" + p.outputQty);
    if (p.greenBeanItemId && p.inputQty > 0) {
      createInventoryTransaction(p.greenBeanItemId, "OUT", p.inputQty, refType, refId, "Roast batch " + (i+1), person, "rb_" + i + "_out");
    }
    if (p.roastedBeanItemId && p.outputQty > 0) {
      createInventoryTransaction(p.roastedBeanItemId, "IN", p.outputQty, refType, refId, "Roast batch " + (i+1), person, "rb_" + i + "_in", p.classificationId);
    }
  }
}

function reverseInventoryLedgerForRef(refType, refId, doneBy) {
  if (!refId) return 0;
  var ledger = getRows(SHEETS.INVENTORY_LEDGER);
  var matches = [];
  for (var i = 0; i < ledger.length; i++) {
    var row = ledger[i];
    if (String(row.reference_id) !== String(refId)) continue;
    if (refType && String(row.reference_type) !== String(refType)) continue;
    if (String(row.notes || "").indexOf("[REVERSAL]") === 0) continue;
    matches.push(row);
  }
  if (matches.length === 0) return 0;
  for (var j = 0; j < matches.length; j++) {
    var orig = matches[j];
    var origType = String(orig.type || "");
    var origQty = Math.abs(parseFloat(orig.quantity) || 0);
    if (!orig.item_id || origQty <= 0) continue;
    var reverseType = origType === "IN" ? "OUT" : "IN";
    createInventoryTransaction(orig.item_id, reverseType, origQty, refType || orig.reference_type, refId, "[REVERSAL] of " + orig.id, doneBy || "", orig.question_index);
  }
  invalidateCache(SHEETS.INVENTORY_LEDGER);
  return matches.length;
}

// Mock submission lookup for grinding test
const mockSubmissions = {};

// ═══════════════════════════════════════════════════════════════
// TEST SCENARIOS
// ═══════════════════════════════════════════════════════════════

function runScenario(name, fn) {
  console.log("\n" + "═".repeat(60));
  console.log("SCENARIO: " + name);
  console.log("═".repeat(60));
  try {
    const result = fn();
    console.log(result ? "✅ PASS" : "❌ FAIL");
    return result;
  } catch (e) {
    console.log("❌ FAIL (exception): " + e.message);
    console.log(e.stack);
    return false;
  }
}

function scenario1() {
  resetSheets(); clearRowsCache();
  console.log("  Calling applyLegacyInventoryForChecklist for ck_green_beans...");

  const ck = { id: "ck_green_beans", questions: [
    { text: "Source Sample", type: "text" },
    { text: "Type of Beans", type: "inventory_item" },
    { text: "Quantity received", type: "number" },
    { text: "Bags stored in which location", type: "text" },
    { text: "Shipment Approved?", type: "yesno" },
  ]};

  const respMap = { "Type of Beans": "inv_test_gb", "Quantity received": "100" };
  applyLegacyInventoryForChecklist(ck, respMap, "untagged", "test-gb-001", "TEST", "", "", false);

  const ledger = getRows(SHEETS.INVENTORY_LEDGER);
  console.log("  Ledger rows after: " + ledger.length);

  if (ledger.length === 0) { console.log("  !! No ledger rows created"); return false; }
  const row = ledger[0];
  const checks = [
    { name: "type=IN", ok: row.type === "IN", got: row.type },
    { name: "qty=100", ok: parseFloat(row.quantity) === 100, got: row.quantity },
    { name: "category=Green Beans", ok: row.category === "Green Beans", got: row.category },
    { name: "item_id=inv_test_gb", ok: row.item_id === "inv_test_gb", got: row.item_id },
    { name: "reference_id=test-gb-001", ok: row.reference_id === "test-gb-001", got: row.reference_id },
  ];
  let allPass = true;
  checks.forEach(c => {
    console.log("  " + (c.ok ? "✓" : "✗") + " " + c.name + (c.ok ? "" : " (got: " + c.got + ")"));
    if (!c.ok) allPass = false;
  });
  return allPass;
}

function scenario2() {
  resetSheets(); clearRowsCache();
  console.log("  Calling applyRoastBatchInventory...");

  const processed = [{
    sourceAutoId: "GB-TEST-001",
    inputQty: 40, outputQty: 35,
    greenBeanItemId: "inv_test_gb",
    greenBeanWarning: "",
    roastedBeanItemId: "inv_test_rb",
    roastedBeanWarning: "",
    classificationId: "",
    beanRef: "inv_test_gb",
  }];

  applyRoastBatchInventory(processed, "untagged", "test-rb-001", "TEST");

  const ledger = getRows(SHEETS.INVENTORY_LEDGER);
  console.log("  Ledger rows after: " + ledger.length);

  if (ledger.length < 2) { console.log("  !! Expected 2 rows, got " + ledger.length); return false; }
  const outRow = ledger.find(r => r.type === "OUT");
  const inRow = ledger.find(r => r.type === "IN");
  const checks = [
    { name: "OUT row exists", ok: !!outRow, got: outRow ? "yes" : "no" },
    { name: "IN row exists", ok: !!inRow, got: inRow ? "yes" : "no" },
    { name: "OUT qty=-40", ok: outRow && parseFloat(outRow.quantity) === -40, got: outRow ? outRow.quantity : "N/A" },
    { name: "IN qty=35", ok: inRow && parseFloat(inRow.quantity) === 35, got: inRow ? inRow.quantity : "N/A" },
    { name: "OUT category=Green Beans", ok: outRow && outRow.category === "Green Beans", got: outRow ? outRow.category : "N/A" },
    { name: "IN category=Roasted Beans", ok: inRow && inRow.category === "Roasted Beans", got: inRow ? inRow.category : "N/A" },
    { name: "OUT ref_id=test-rb-001", ok: outRow && outRow.reference_id === "test-rb-001", got: outRow ? outRow.reference_id : "N/A" },
  ];
  let allPass = true;
  checks.forEach(c => {
    console.log("  " + (c.ok ? "✓" : "✗") + " " + c.name + (c.ok ? "" : " (got: " + c.got + ")"));
    if (!c.ok) allPass = false;
  });
  return allPass;
}

function scenario3() {
  resetSheets(); clearRowsCache();
  console.log("  Calling applyLegacyInventoryForChecklist for ck_grinding...");

  // Mock the submission lookup: when grinding looks up Roast ID "RB-TEST-001",
  // it should find beanRef = "inv_test_gb" (the green bean item used in that roast)
  mockSubmissions["RB-TEST-001"] = { beanRef: "inv_test_gb" };

  const ck = { id: "ck_grinding", questions: [
    { text: "Roast ID", type: "text", linkedSource: { checklistId: "ck_roasted_beans" } },
    { text: "Invoice/SO", type: "text" },
    { text: "Client name", type: "text" },
    { text: "Grind size", type: "text" },
    { text: "Is the correct stickers applied", type: "yesno" },
    { text: "Total Net weight", type: "number" },
  ]};

  const respMap = { "Roast ID": "RB-TEST-001", "Total Net weight": "28" };
  applyLegacyInventoryForChecklist(ck, respMap, "untagged", "test-rg-001", "TEST", "", "", false);

  const ledger = getRows(SHEETS.INVENTORY_LEDGER);
  console.log("  Ledger rows after: " + ledger.length);

  if (ledger.length < 2) { console.log("  !! Expected 2 rows, got " + ledger.length); return false; }
  const outRow = ledger.find(r => r.type === "OUT");
  const inRow = ledger.find(r => r.type === "IN");
  const checks = [
    { name: "OUT row exists", ok: !!outRow, got: outRow ? "yes" : "no" },
    { name: "IN row exists", ok: !!inRow, got: inRow ? "yes" : "no" },
    { name: "OUT category=Roasted Beans", ok: outRow && outRow.category === "Roasted Beans", got: outRow ? outRow.category : "N/A" },
    { name: "IN category=Packing Items", ok: inRow && inRow.category === "Packing Items", got: inRow ? inRow.category : "N/A" },
    { name: "OUT qty=-28", ok: outRow && parseFloat(outRow.quantity) === -28, got: outRow ? outRow.quantity : "N/A" },
    { name: "IN qty=28", ok: inRow && parseFloat(inRow.quantity) === 28, got: inRow ? inRow.quantity : "N/A" },
  ];
  let allPass = true;
  checks.forEach(c => {
    console.log("  " + (c.ok ? "✓" : "✗") + " " + c.name + (c.ok ? "" : " (got: " + c.got + ")"));
    if (!c.ok) allPass = false;
  });
  return allPass;
}

function scenario4() {
  // Pre-populate with scenario 3 data, then reverse
  resetSheets(); clearRowsCache();
  mockSubmissions["RB-TEST-001"] = { beanRef: "inv_test_gb" };
  const ck = { id: "ck_grinding", questions: [
    { text: "Roast ID", type: "text", linkedSource: { checklistId: "ck_roasted_beans" } },
    { text: "Total Net weight", type: "number" },
  ]};
  const respMap = { "Roast ID": "RB-TEST-001", "Total Net weight": "28" };
  applyLegacyInventoryForChecklist(ck, respMap, "untagged", "test-rg-001", "TEST", "", "", false);
  clearRowsCache();

  console.log("  Ledger rows before reversal: " + getRows(SHEETS.INVENTORY_LEDGER).length);
  console.log("  Calling reverseInventoryLedgerForRef...");

  const reversed = reverseInventoryLedgerForRef("untagged", "test-rg-001", "TEST_REVERSAL");
  clearRowsCache();
  const ledger = getRows(SHEETS.INVENTORY_LEDGER);
  console.log("  Ledger rows after reversal: " + ledger.length);
  console.log("  Reversed count: " + reversed);

  const reversals = ledger.filter(r => String(r.notes || "").indexOf("[REVERSAL]") === 0);
  const checks = [
    { name: "2 reversal rows", ok: reversals.length === 2, got: reversals.length },
    { name: "reversed count = 2", ok: reversed === 2, got: reversed },
  ];
  reversals.forEach((r, i) => {
    checks.push({ name: "reversal " + (i+1) + " has [REVERSAL] in notes", ok: true, got: r.notes });
  });
  let allPass = true;
  checks.forEach(c => {
    console.log("  " + (c.ok ? "✓" : "✗") + " " + c.name + (c.ok ? "" : " (got: " + c.got + ")"));
    if (!c.ok) allPass = false;
  });
  return allPass;
}

// ═══════════════════════════════════════════════════════════════
// RUN ALL SCENARIOS
// ═══════════════════════════════════════════════════════════════

console.log("\n🔬 INVENTORY WRITE SIMULATION TEST\n");

const results = [
  runScenario("1: Green Beans QC ledger write", scenario1),
  runScenario("2: Roasted Beans multi-batch ledger write", scenario2),
  runScenario("3: Grinding ledger write", scenario3),
  runScenario("4: Soft delete reversal", scenario4),
];

console.log("\n" + "═".repeat(60));
console.log("SUMMARY");
console.log("═".repeat(60));
const passed = results.filter(r => r).length;
const failed = results.filter(r => !r).length;
console.log(`Passed: ${passed}  Failed: ${failed}`);
console.log(passed === 4 ? "✅ ALL PASS" : "❌ " + failed + " FAILURES");

// Print all writes
console.log("\n📝 Sheet writes during tests: " + writeLog.length);
writeLog.forEach((w, i) => {
  if (w.action === "append" && w.sheet === SHEETS.INVENTORY_LEDGER) {
    console.log(`  ${i+1}. APPEND InventoryLedger: type=${w.obj.type} qty=${w.obj.quantity} item=${w.obj.item_name} cat=${w.obj.category} ref=${w.obj.reference_id}`);
  }
});
