// ═══════════════════════════════════════════════════════════════
// Order Checklist Manager — Google Apps Script Backend
// Copy this entire file into Google Apps Script editor
// ═══════════════════════════════════════════════════════════════

/**
 * @OnlyCurrentDoc
 */

// ─── Sheet Names ───────────────────────────────────────────────
var SHEETS = {
  ORDERS: "Orders",
  ORDER_CHECKLISTS: "OrderChecklists",
  CHECKLISTS: "Checklists",
  ORDER_TYPES: "OrderTypes",
  CUSTOMERS: "Customers",
  RULES: "AssignmentRules",
  USERS: "Users",
  SESSIONS: "Sessions",
  AUDIT_LOG: "AuditLog",
  CONFIG: "Config",
  MASTER_SUMMARY: "Master Summary",
  ARCHIVED_ORDERS: "ArchivedOrders",
  ARCHIVED_ORDER_CHECKLISTS: "ArchivedOrderChecklists",
  ARCHIVED_RESPONSES: "ArchivedResponses",
  ARCHIVES_META: "ArchivesMeta",
  UNTAGGED_CHECKLISTS: "UntaggedChecklists",
  INVENTORY_ITEMS: "InventoryItems",
  INVENTORY_LEDGER: "InventoryLedger",
  INVENTORY_CATEGORIES: "InventoryCategories",
  ID_SEQUENCES: "IDSequences",
  QUANTITY_ALLOCATIONS: "QuantityAllocations",
  BLENDS: "Blends",
  DRAFTS: "Drafts",
  ROAST_CLASSIFICATIONS: "RoastClassifications",
};

// ─── Headers for standard sheets ──────────────────────────────
var HEADERS = {
  Orders: ["id", "name", "customer_id", "assigned_to", "order_type", "created_at", "status", "invoice_so", "order_type_detail", "order_lines", "product_type", "missing_checklist_reasons", "stages", "delivered_at"],
  OrderChecklists: ["id", "order_id", "checklist_id", "status", "completed_at", "completed_by", "work_date", "auto_id"],
  Checklists: ["id", "name", "subtitle", "form_url", "questions", "auto_id_config", "can_tag_to"],
  OrderTypes: ["id", "label"],
  Customers: ["id", "label"],
  AssignmentRules: ["id", "order_type_id", "customer_id", "checklist_ids"],
  Users: ["id", "username", "password_hash", "display_name", "role", "status", "created_at", "created_by"],
  Sessions: ["token", "user_id", "created_at", "last_active"],
  AuditLog: ["id", "user_id", "user_name", "action", "entity_type", "entity_id", "details", "timestamp"],
  Config: ["key", "value"],
  ArchivedOrders: ["id", "name", "customer_id", "assigned_to", "order_type", "created_at", "status", "invoice_so", "order_type_detail", "order_lines", "archive_id"],
  ArchivedOrderChecklists: ["id", "order_id", "checklist_id", "status", "completed_at", "completed_by", "work_date", "archive_id"],
  ArchivedResponses: ["archive_id", "order_id", "order_name", "customer", "checklist_name", "person", "date", "question", "response", "remark", "submitted_at"],
  ArchivesMeta: ["id", "date_range_start", "date_range_end", "orders_count", "created_at", "created_by", "order_ids"],
  UntaggedChecklists: ["id", "checklist_id", "checklist_name", "person", "date", "submitted_at", "tagged_order_id", "responses", "remarks", "submitted_by_user_id", "total_quantity", "tagged_quantity", "allocations", "auto_id", "is_deleted"],
  InventoryItems: ["id", "category", "name", "unit", "opening_stock", "current_stock", "min_stock_alert", "created_at", "is_active", "abbreviation", "equivalent_items", "classification_id"],
  InventoryLedger: ["id", "item_id", "item_name", "category", "date", "type", "quantity", "balance_after", "reference_type", "reference_id", "notes", "done_by", "created_at", "question_index", "classification_id"],
  InventoryCategories: ["id", "name"],
  IDSequences: ["prefix", "last_sequence"],
  QuantityAllocations: ["id", "source_checklist_id", "source_auto_id", "total_quantity", "destination_type", "destination_id", "destination_auto_id", "allocated_quantity", "allocated_at", "allocated_by"],
  Blends: ["id", "name", "customer", "description", "components", "is_active", "created_at"],
  Drafts: ["id", "checklist_id", "checklist_name", "user_id", "user_name", "responses", "linked_sources", "linked_orders", "remarks", "batch_allocations", "person", "work_date", "created_at", "updated_at"],
  RoastClassifications: ["id", "name", "type", "description", "created_by", "created_at", "updated_at", "is_active"],
};

// Common columns for per-checklist response tabs (Row 2 headers, data starts Row 3)
var RESPONSE_COMMON = ["Order ID", "Order Name", "Customer", "Person", "Date", "Submitted At"];
var RESPONSE_COMMON_COUNT = RESPONSE_COMMON.length; // 6

// Master Summary data columns (Row 4 headers, data starts Row 5)
var MASTER_SUMMARY_HEADERS = ["Order ID", "Order Name", "Customer", "Order Type", "Checklist", "Person", "Date", "Submitted At"];

// ─── Request-Level Cache ──────────────────────────────────────
var _rowsCache = {};
function clearRowsCache() { _rowsCache = {}; _columnMigrationDone = false; }
function invalidateCache(sheetName) { delete _rowsCache[sheetName]; }

// ─── Helpers ───────────────────────────────────────────────────

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet && HEADERS[name]) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(HEADERS[name]);
    sheet.getRange(1, 1, 1, HEADERS[name].length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Ensures a sheet has all the columns declared in HEADERS by appending any missing columns at the end.
// Used to migrate existing sheets when new columns are added in code without losing data.
function ensureSheetHasAllColumns(name) {
  if (!HEADERS[name]) return;
  var sheet = getSheet(name);
  if (!sheet) return;
  var expected = HEADERS[name];
  var lastCol = sheet.getLastColumn();
  var current = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  var changed = false;
  for (var i = 0; i < expected.length; i++) {
    if (current.indexOf(expected[i]) < 0) {
      sheet.getRange(1, current.length + 1).setValue(expected[i]).setFontWeight("bold");
      current.push(expected[i]);
      changed = true;
    }
  }
  if (changed) invalidateCache(name);
}

// Run column migration for all known sheets that may have evolved over time.
var _columnMigrationDone = false;
function ensureAllSheetColumnsMigrated() {
  if (_columnMigrationDone) return;
  _columnMigrationDone = true;
  try {
    ensureSheetHasAllColumns(SHEETS.ORDERS);
    ensureSheetHasAllColumns(SHEETS.ORDER_CHECKLISTS);
    ensureSheetHasAllColumns(SHEETS.CHECKLISTS);
    ensureSheetHasAllColumns(SHEETS.UNTAGGED_CHECKLISTS);
    ensureSheetHasAllColumns(SHEETS.INVENTORY_ITEMS);
    ensureSheetHasAllColumns(SHEETS.INVENTORY_LEDGER);
    ensureSheetHasAllColumns(SHEETS.ID_SEQUENCES);
    ensureSheetHasAllColumns(SHEETS.QUANTITY_ALLOCATIONS);
    ensureSheetHasAllColumns(SHEETS.BLENDS);
    ensureSheetHasAllColumns(SHEETS.DRAFTS);
    ensureSheetHasAllColumns(SHEETS.ROAST_CLASSIFICATIONS);
    // After ensuring the InventoryLedger has the new "category" column, backfill
    // historical rows once per deploy. Idempotent — only updates rows where
    // category is missing.
    backfillInventoryLedgerCategories();
  } catch (e) { /* non-fatal */ }
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getRows(sheetName) {
  if (_rowsCache[sheetName]) return _rowsCache[sheetName];
  var sheet = getSheet(sheetName);
  if (!sheet) { _rowsCache[sheetName] = []; return []; }
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) { _rowsCache[sheetName] = []; return []; }
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    rows.push(obj);
  }
  _rowsCache[sheetName] = rows;
  return rows;
}

function findRowIndex(sheetName, id) {
  var sheet = getSheet(sheetName);
  if (!sheet) return -1;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function appendToSheet(sheetName, obj) {
  var sheet = getSheet(sheetName);
  var headers = HEADERS[sheetName];
  var row = headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ""; });
  sheet.appendRow(row);
  invalidateCache(sheetName);
}

function updateSheetRow(sheetName, rowIndex, obj) {
  var sheet = getSheet(sheetName);
  var headers = HEADERS[sheetName];
  var row = headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ""; });
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  invalidateCache(sheetName);
}

function deleteSheetRow(sheetName, rowIndex) {
  var sheet = getSheet(sheetName);
  sheet.deleteRow(rowIndex);
  invalidateCache(sheetName);
}

function nextId() {
  return String(new Date().getTime()) + String(Math.floor(Math.random() * 1000));
}

// ─── Auto ID (sequence + builder) ─────────────────────────────

// Default prefix recommendations by checklist name (used as fallback when admin hasn't set autoIdConfig.prefix).
var DEFAULT_AUTO_ID_PREFIXES = {
  "Green Bean QC Sample Check": "GBS",
  "Green Beans Quality Check": "GB",
  "Roasted Beans Quality Check": "RB",
  "Grinding & Packing Checklist": "RG",
  "Tagging Roasted Beans": "TG",
  "Sample Retention Checklist": "SR",
  "Coffee with Chicory Mix": "CC",
};

function getDefaultPrefixForChecklist(name) {
  return DEFAULT_AUTO_ID_PREFIXES[name] || "";
}

// Atomically read+increment the sequence counter for a given prefix.
function getNextSequenceForPrefix(prefix) {
  ensureSheetHasAllColumns(SHEETS.ID_SEQUENCES);
  var sheet = getSheet(SHEETS.ID_SEQUENCES);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(prefix)) {
      var num = (parseInt(data[i][1]) || 0) + 1;
      sheet.getRange(i + 1, 2).setValue(num);
      invalidateCache(SHEETS.ID_SEQUENCES);
      return num;
    }
  }
  sheet.appendRow([prefix, 1]);
  invalidateCache(SHEETS.ID_SEQUENCES);
  return 1;
}

function formatDateDDMMYY(value) {
  if (!value) return "";
  var d;
  if (value instanceof Date) d = value;
  else {
    var s = String(value);
    // Accept yyyy-mm-dd or ISO
    d = new Date(s);
    if (isNaN(d.getTime())) return "";
  }
  var dd = String(d.getDate()).padStart(2, "0");
  var mm = String(d.getMonth() + 1).padStart(2, "0");
  var yy = String(d.getFullYear()).slice(-2);
  return dd + mm + yy;
}

// Look up an inventory item's abbreviation by id, falling back to a sanitized version of the name.
function getInventoryAbbreviation(itemIdOrName) {
  if (!itemIdOrName) return "";
  var items = getRows(SHEETS.INVENTORY_ITEMS);
  for (var i = 0; i < items.length; i++) {
    if (String(items[i].id) === String(itemIdOrName) || String(items[i].name) === String(itemIdOrName)) {
      var ab = String(items[i].abbreviation || "").trim().toUpperCase();
      if (ab) return ab;
      // Fallback: take initials of the item name
      var nm = String(items[i].name || "");
      return nm.replace(/[^A-Za-z0-9]/g, "").slice(0, 4).toUpperCase();
    }
  }
  return "";
}

// Sanitize an arbitrary text response into an item-code-style token (uppercase alphanumeric, max 6 chars).
function sanitizeItemCodeToken(value) {
  if (!value) return "";
  return String(value).replace(/[^A-Za-z0-9]/g, "").slice(0, 6).toUpperCase();
}

// Build the [PREFIX]-[DATE]-[ITEM_CODE]-[SEQ] string. Sequence is allocated atomically.
// responses: array of {questionIndex, response} OR an object {questionIndex: value}.
function generateAutoId(checklist, responsesByIndex, fallbackDate) {
  if (!checklist) return "";
  var cfg = checklist.autoIdConfig;
  if (!cfg || !cfg.enabled) return "";
  var prefix = (cfg.prefix || getDefaultPrefixForChecklist(checklist.name) || "AUTO").toUpperCase();

  // Resolve date portion
  var dateStr = "";
  if (cfg.dateFieldIdx !== null && cfg.dateFieldIdx !== undefined) {
    var dateVal = responsesByIndex[cfg.dateFieldIdx];
    dateStr = formatDateDDMMYY(dateVal);
  }
  if (!dateStr) dateStr = formatDateDDMMYY(fallbackDate || new Date());

  // Resolve item-code portion: if the configured field is a linkedSource or text, use sanitized response;
  // if the response value matches an inventory item id/name, use its abbreviation.
  var itemCode = "";
  if (cfg.itemCodeFieldIdx !== null && cfg.itemCodeFieldIdx !== undefined) {
    var raw = responsesByIndex[cfg.itemCodeFieldIdx];
    var ab = getInventoryAbbreviation(raw);
    itemCode = ab || sanitizeItemCodeToken(raw);
  }
  if (!itemCode) itemCode = "X";

  var seq = getNextSequenceForPrefix(prefix);
  var seqStr = String(seq).padStart(3, "0");
  return prefix + "-" + dateStr + "-" + itemCode + "-" + seqStr;
}

// Helper: build a {questionIndex: response} map from the responses array used by handleSubmitChecklist.
function responsesArrayToMap(responses) {
  var map = {};
  if (!Array.isArray(responses)) return map;
  for (var i = 0; i < responses.length; i++) {
    var r = responses[i];
    if (r && r.questionIndex !== undefined) map[r.questionIndex] = r.response || "";
  }
  return map;
}

function safeParseJSON(str, fallback) {
  try {
    if (typeof str === "string" && str.length > 0) return JSON.parse(str);
    return fallback;
  } catch (e) { return fallback; }
}

// ─── Order ID Generator (ORD-001 format) ──────────────────────

function getNextOrderId() {
  var sheet = getSheet(SHEETS.CONFIG);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === "next_order_number") {
      var num = parseInt(data[i][1]) || 1;
      sheet.getRange(i + 1, 2).setValue(num + 1);
      invalidateCache(SHEETS.CONFIG);
      return "ORD-" + String(num).padStart(3, "0");
    }
  }
  // Not found — create it
  sheet.appendRow(["next_order_number", 2]);
  invalidateCache(SHEETS.CONFIG);
  return "ORD-001";
}

// ─── Question Normalization ───────────────────────────────────
// Converts old string[] questions to new object[] format for backward compat.

function normalizeQuestions(questions) {
  if (!Array.isArray(questions)) return [];
  var DEFAULTS = { text: "", type: "text", formula: null, ideal: null, remarkCondition: null, isApprovalGate: false, linkedSource: null, inventoryLink: null, isMasterQuantity: false, inventoryCategory: "", idealLabel: "", idealUnit: "", remarksTargetIdx: null, autoFillMapping: null, dateComparison: null };
  return questions.map(function(q) {
    if (typeof q === "string") return { text: q, type: "text", formula: null, ideal: null, remarkCondition: null, isApprovalGate: false, linkedSource: null, inventoryLink: null, isMasterQuantity: false, inventoryCategory: "", idealLabel: "", idealUnit: "", remarksTargetIdx: null, autoFillMapping: null, dateComparison: null };
    // Start with defaults, then overlay ALL existing fields from q to preserve any extra metadata
    var result = {};
    var k;
    for (k in DEFAULTS) { if (DEFAULTS.hasOwnProperty(k)) result[k] = DEFAULTS[k]; }
    for (k in q) { if (q.hasOwnProperty(k) && q[k] !== undefined) result[k] = q[k]; }
    // Normalize specific fields
    if (!result.text) result.text = "";
    if (!result.type) result.type = "text";
    result.remarksTargetIdx = (result.remarksTargetIdx === null || result.remarksTargetIdx === undefined || result.remarksTargetIdx === "") ? null : Number(result.remarksTargetIdx);
    return result;
  });
}

function normalizeAutoIdConfig(cfg) {
  if (!cfg || typeof cfg !== "object") return null;
  return {
    enabled: cfg.enabled === true,
    prefix: String(cfg.prefix || "").toUpperCase(),
    dateFieldIdx: (cfg.dateFieldIdx === null || cfg.dateFieldIdx === undefined || cfg.dateFieldIdx === "") ? null : Number(cfg.dateFieldIdx),
    itemCodeFieldIdx: (cfg.itemCodeFieldIdx === null || cfg.itemCodeFieldIdx === undefined || cfg.itemCodeFieldIdx === "") ? null : Number(cfg.itemCodeFieldIdx),
  };
}

function questionTexts(nq) { return nq.map(function(q) { return q.text; }); }

function getRemarkIndices(nq) {
  var indices = [];
  for (var i = 0; i < nq.length; i++) { if (nq[i].remarkCondition) indices.push(i); }
  return indices;
}

// ─── Per-Checklist Response Tab Helpers ───────────────────────

function getOrCreateResponseSheet(checklistName, questions) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(checklistName);
  if (sheet) return sheet;

  var nq = normalizeQuestions(questions);
  var qTexts = questionTexts(nq);
  var remarkIdx = getRemarkIndices(nq);
  var remarkHeaders = remarkIdx.map(function(i) { return "Remarks: " + nq[i].text; });
  var dataCols = qTexts.length + remarkHeaders.length;
  var totalCols = RESPONSE_COMMON_COUNT + dataCols;

  sheet = ss.insertSheet(checklistName);

  // Row 1: common headers + checklist name merged across question+remark columns
  var row1 = RESPONSE_COMMON.slice();
  row1.push(checklistName);
  for (var i = 1; i < dataCols; i++) row1.push("");
  sheet.getRange(1, 1, 1, totalCols).setValues([row1]);
  if (dataCols > 1) sheet.getRange(1, RESPONSE_COMMON_COUNT + 1, 1, dataCols).merge();
  sheet.getRange(1, RESPONSE_COMMON_COUNT + 1, 1, 1)
    .setBackground("#D4A574").setFontColor("#FFFFFF").setHorizontalAlignment("center").setFontWeight("bold");

  // Row 2: common headers + question texts + remark headers
  var row2 = RESPONSE_COMMON.slice().concat(qTexts).concat(remarkHeaders);
  sheet.getRange(2, 1, 1, totalCols).setValues([row2]);

  // Formatting
  sheet.getRange(1, 1, 2, totalCols).setFontWeight("bold");
  sheet.setFrozenRows(2);
  sheet.getRange(2, RESPONSE_COMMON_COUNT + 1, 1, dataCols).setWrap(true);
  for (var k = 1; k <= RESPONSE_COMMON_COUNT; k++) sheet.setColumnWidth(k, 120);
  for (var m = RESPONSE_COMMON_COUNT + 1; m <= totalCols; m++) sheet.setColumnWidth(m, 160);
  // Color remark header columns amber
  if (remarkHeaders.length > 0) {
    sheet.getRange(2, RESPONSE_COMMON_COUNT + qTexts.length + 1, 1, remarkHeaders.length)
      .setBackground("#FFF3CD").setFontColor("#856404");
  }

  return sheet;
}

// Ensures an existing response sheet has ALL question + remark columns.
// NEVER removes existing columns — only appends missing ones.
function ensureResponseSheetColumns(checklistName, questions) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(checklistName);
  if (!sheet) return getOrCreateResponseSheet(checklistName, questions);

  var nq = normalizeQuestions(questions);
  var qTexts = questionTexts(nq);
  var remarkIdx = getRemarkIndices(nq);
  var remarkHeaders = remarkIdx.map(function(i) { return "Remarks: " + nq[i].text; });

  var currentCols = sheet.getLastColumn();
  if (currentCols < 1) return getOrCreateResponseSheet(checklistName, questions);
  var existingRow2 = sheet.getRange(2, 1, 1, currentCols).getValues()[0].map(String);

  var nextCol = currentCols + 1;
  // Add missing question columns
  for (var i = 0; i < qTexts.length; i++) {
    if (existingRow2.indexOf(qTexts[i]) < 0) {
      sheet.getRange(1, nextCol).setValue("");
      sheet.getRange(2, nextCol).setValue(qTexts[i]).setFontWeight("bold").setWrap(true);
      sheet.setColumnWidth(nextCol, 160);
      existingRow2.push(qTexts[i]);
      nextCol++;
    }
  }
  // Add missing remark columns
  for (var j = 0; j < remarkHeaders.length; j++) {
    if (existingRow2.indexOf(remarkHeaders[j]) < 0) {
      sheet.getRange(1, nextCol).setValue("");
      sheet.getRange(2, nextCol).setValue(remarkHeaders[j]).setFontWeight("bold").setWrap(true)
        .setBackground("#FFF3CD").setFontColor("#856404");
      sheet.setColumnWidth(nextCol, 160);
      existingRow2.push(remarkHeaders[j]);
      nextCol++;
    }
  }
  return sheet;
}

// Helper: resolve an inventory item id to its display name. Returns the original value if not found.
function inventoryItemNameById(value) {
  if (!value) return value;
  var items = getRows(SHEETS.INVENTORY_ITEMS);
  for (var i = 0; i < items.length; i++) {
    if (String(items[i].id) === String(value)) return String(items[i].name || value);
  }
  return value;
}

// Helper: resolve an inventory item display name back to its id. Returns the original value if not found.
function inventoryItemIdByName(value) {
  if (!value) return value;
  var items = getRows(SHEETS.INVENTORY_ITEMS);
  for (var i = 0; i < items.length; i++) {
    if (String(items[i].name) === String(value)) return String(items[i].id);
  }
  return value;
}

function writeResponseRow(checklistName, questions, data) {
  var nq = normalizeQuestions(questions);
  var sheet = ensureResponseSheetColumns(checklistName, nq);
  var qTexts = questionTexts(nq);
  var remarkIdx = getRemarkIndices(nq);
  var remarkHeaders = remarkIdx.map(function(ri) { return "Remarks: " + nq[ri].text; });

  // Read row 2 headers to build a column index map
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0 ? sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(String) : [];

  // Build row with correct number of columns
  var row = [];
  for (var c = 0; c < headers.length; c++) row.push("");

  // Fill common columns (positional — always first 6)
  var common = [data.orderId, data.orderName, data.customer, data.person, data.date, data.submittedAt];
  for (var ci = 0; ci < common.length && ci < headers.length; ci++) row[ci] = common[ci];

  // Fill question columns by header name lookup — write ALL field types
  for (var i = 0; i < qTexts.length; i++) {
    var colIdx = headers.indexOf(qTexts[i]);
    if (colIdx < 0) continue;
    var raw = (data.responses[i] !== undefined && data.responses[i] !== null) ? String(data.responses[i]) : "";
    // Inventory item fields: store display name (not raw id)
    if (nq[i] && nq[i].type === "inventory_item" && raw) raw = inventoryItemNameById(raw);
    row[colIdx] = raw;
  }
  // Fill remark columns by header name lookup
  for (var j = 0; j < remarkIdx.length; j++) {
    var qi = remarkIdx[j];
    var rmkColIdx = headers.indexOf(remarkHeaders[j]);
    if (rmkColIdx < 0) continue;
    row[rmkColIdx] = data.remarks ? (data.remarks[qi] || "") : "";
  }
  sheet.appendRow(row);
}

function readResponseRow(checklistName, questions, orderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(checklistName);
  if (!sheet || sheet.getLastRow() < 3) return null;

  var nq = normalizeQuestions(questions);
  var qTexts = questionTexts(nq);
  var remarkIdx = getRemarkIndices(nq);
  var remarkHeaders = remarkIdx.map(function(ri) { return "Remarks: " + nq[ri].text; });
  var data = sheet.getDataRange().getValues();

  // Read row 2 headers to build column index map
  var headers = data.length >= 2 ? data[1].map(String) : [];

  for (var i = 2; i < data.length; i++) {
    if (String(data[i][0]) === String(orderId)) {
      var responses = [];
      for (var q = 0; q < nq.length; q++) {
        var colIdx = headers.indexOf(qTexts[q]);
        var val = colIdx >= 0 ? String(data[i][colIdx] || "") : "";
        responses.push({
          questionIndex: q, questionText: nq[q].text,
          response: val, remark: "",
        });
      }
      // Read remark columns by header name
      for (var r = 0; r < remarkIdx.length; r++) {
        var rmkColIdx = headers.indexOf(remarkHeaders[r]);
        if (rmkColIdx >= 0) {
          responses[remarkIdx[r]].remark = String(data[i][rmkColIdx] || "");
        }
      }
      return {
        person: String(data[i][3] || ""), date: String(data[i][4] || ""),
        submittedAt: String(data[i][5] || ""), responses: responses,
      };
    }
  }
  return null;
}

function deleteResponseRow(checklistName, orderId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(checklistName);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 2; i--) {
    if (String(data[i][0]) === String(orderId)) sheet.deleteRow(i + 1);
  }
}

function deleteResponsesByOcId(ocId) {
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  var oc = null;
  for (var i = 0; i < ocs.length; i++) {
    if (String(ocs[i].id) === String(ocId)) { oc = ocs[i]; break; }
  }
  if (!oc) return;
  var ckRows = getRows(SHEETS.CHECKLISTS);
  for (var j = 0; j < ckRows.length; j++) {
    if (String(ckRows[j].id) === String(oc.checklist_id)) {
      deleteResponseRow(ckRows[j].name, String(oc.order_id));
      deleteFromMasterSummary(String(oc.order_id), ckRows[j].name);
      return;
    }
  }
}

// ─── Master Summary Helpers ───────────────────────────────────

function addToMasterSummary(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MASTER_SUMMARY);
  if (!sheet) return;
  sheet.appendRow([data.orderId, data.orderName, data.customer, data.orderType, data.checklistName, data.person, data.date, data.submittedAt]);
}

function deleteFromMasterSummary(orderId, checklistName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MASTER_SUMMARY);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  // Data starts at row 5 (index 4), after filter rows (1-2), blank row (3), header row (4)
  for (var i = data.length - 1; i >= 4; i--) {
    var match = String(data[i][0]) === String(orderId);
    if (checklistName) match = match && String(data[i][4]) === String(checklistName);
    if (match) sheet.deleteRow(i + 1);
  }
}

// ─── Lookup Helpers ───────────────────────────────────────────

function lookupOrder(orderId) {
  var orders = getRows(SHEETS.ORDERS);
  for (var i = 0; i < orders.length; i++) {
    if (String(orders[i].id) === String(orderId)) return orders[i];
  }
  return null;
}

function lookupCustomerLabel(customerId) {
  var customers = getRows(SHEETS.CUSTOMERS);
  for (var i = 0; i < customers.length; i++) {
    if (String(customers[i].id) === String(customerId)) return customers[i].label;
  }
  return "";
}

function lookupOrderTypeLabel(orderTypeId) {
  var ots = getRows(SHEETS.ORDER_TYPES);
  for (var i = 0; i < ots.length; i++) {
    if (String(ots[i].id) === String(orderTypeId)) return ots[i].label;
  }
  return "";
}

function lookupChecklist(checklistId) {
  var rows = getRows(SHEETS.CHECKLISTS);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(checklistId)) {
      var nq = normalizeQuestions(safeParseJSON(rows[i].questions, []));
      return {
        id: rows[i].id, name: rows[i].name, questions: nq,
        autoIdConfig: normalizeAutoIdConfig(safeParseJSON(rows[i].auto_id_config, null)),
      };
    }
  }
  return null;
}

// ─── Auth Helpers ──────────────────────────────────────────────

function hashPassword(password) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return digest.map(function(b) {
    var hex = (b < 0 ? b + 256 : b).toString(16);
    return hex.length === 1 ? "0" + hex : hex;
  }).join("");
}

function generateToken() { return Utilities.getUuid() + "_" + nextId(); }

function validateToken(tokenString) {
  if (!tokenString) return null;
  var sessions = getRows(SHEETS.SESSIONS);
  var session = null;
  for (var i = 0; i < sessions.length; i++) {
    if (String(sessions[i].token) === String(tokenString)) { session = sessions[i]; break; }
  }
  if (!session) return null;

  var lastActive = new Date(session.last_active);
  var now = new Date();
  if ((now.getTime() - lastActive.getTime()) / (1000 * 60 * 60) > 8) {
    var idx = findRowIndex(SHEETS.SESSIONS, session.token);
    if (idx > 0) deleteSheetRow(SHEETS.SESSIONS, idx);
    return null;
  }

  var sIdx = findRowIndex(SHEETS.SESSIONS, session.token);
  if (sIdx > 0) { session.last_active = now.toISOString(); updateSheetRow(SHEETS.SESSIONS, sIdx, session); }

  var users = getRows(SHEETS.USERS);
  for (var j = 0; j < users.length; j++) {
    if (String(users[j].id) === String(session.user_id) && users[j].status === "active")
      return { id: users[j].id, username: users[j].username, displayName: users[j].display_name, role: users[j].role };
  }
  return null;
}

function requireAdmin(user) { if (user.role !== "admin") return { error: "FORBIDDEN" }; return null; }

function writeAuditLog(user, action, entityType, entityId, details) {
  appendToSheet(SHEETS.AUDIT_LOG, {
    id: "audit_" + nextId(), user_id: user.id, user_name: user.displayName || user.username,
    action: action, entity_type: entityType, entity_id: entityId || "", details: details || "",
    timestamp: new Date().toISOString(),
  });
  // Also append to per-module audit tab
  try {
    var tabName = resolveAuditTabName(entityType, entityId);
    appendAuditLog(tabName, {
      timestamp: new Date().toISOString(),
      user: user.displayName || user.username || "",
      action: action,
      recordId: entityId || "",
      fieldChanged: "",
      oldValue: "",
      newValue: "",
      notes: details || "",
    });
  } catch (e) { /* non-fatal */ }
}

// Resolve which audit tab a log entry should go into based on entity type/id
function resolveAuditTabName(entityType, entityId) {
  var et = String(entityType || "").toLowerCase();
  if (et === "order" || et === "orders") return "Audit - Orders";
  if (et === "inventoryitem" || et === "inventorycategory" || et === "inventoryledger") return "Audit - Inventory";
  if (et === "checklist" || et === "checklistresponse") {
    // Try to resolve checklist name from entity id (OC id)
    if (entityId) {
      var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
      for (var i = 0; i < ocs.length; i++) {
        if (String(ocs[i].id) === String(entityId) || String(ocs[i].auto_id) === String(entityId)) {
          var ck = lookupChecklist(ocs[i].checklist_id);
          if (ck) return "Audit - " + ck.name;
        }
      }
    }
    return "Audit - Checklists";
  }
  if (et === "untaggedchecklist") return "Audit - Untagged";
  return "Audit - Other";
}

// Append a row to a per-module audit tab. Creates the tab with headers if missing.
function appendAuditLog(tabName, entry) {
  if (!tabName) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(tabName);
  var headers = ["Timestamp", "User", "Action", "Record ID", "Field Changed", "Old Value", "New Value", "Notes"];
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    entry.timestamp || new Date().toISOString(),
    entry.user || "",
    entry.action || "",
    entry.recordId || "",
    entry.fieldChanged || "",
    entry.oldValue || "",
    entry.newValue || "",
    entry.notes || "",
  ]);
}

// ─── Request Handlers ──────────────────────────────────────────

function doGet(e) {
  clearRowsCache();
  var action = (e && e.parameter && e.parameter.action) || "";
  try {
    var token = (e && e.parameter && e.parameter.token) || "";
    var user = validateToken(token);
    if (!user) return jsonResponse({ error: "AUTH_EXPIRED" });
    switch (action) {
      case "init":              return jsonResponse(handleInit(user));
      case "getMe":             return jsonResponse({ user: user });
      case "getOrderTypes":     return jsonResponse(handleGetOrderTypes());
      case "getCustomers":      return jsonResponse(handleGetCustomers());
      case "getChecklists":     return jsonResponse(handleGetChecklists());
      case "getRules":          return jsonResponse(handleGetRules());
      case "resolveChecklists": return jsonResponse(handleResolveChecklists(e.parameter));
      case "getOrders":         return jsonResponse(handleGetOrders());
      case "getOrder":          return jsonResponse(handleGetOrder(e.parameter));
      case "getResponses":      return jsonResponse(handleGetResponses(e.parameter));
      case "getUsers":          var err1 = requireAdmin(user); if (err1) return jsonResponse(err1); return jsonResponse(handleGetUsers());
      case "getAllResponses":   return jsonResponse(handleGetAllResponses());
      case "getAuditLog":       var err3 = requireAdmin(user); if (err3) return jsonResponse(err3); return jsonResponse(handleGetAuditLog(e.parameter));
      case "getArchives":       var err4 = requireAdmin(user); if (err4) return jsonResponse(err4); return jsonResponse(handleGetArchives());
      case "getUntagged":       return jsonResponse(handleGetUntagged());
      case "getUntaggedResponse": return jsonResponse(handleGetUntaggedResponse(e.parameter));
      case "getApprovedEntries": return jsonResponse(handleGetApprovedEntries(e.parameter));
      case "getLinkedEntries":  return jsonResponse(handleGetLinkedEntries(e.parameter));
      case "getInventoryItems": return jsonResponse(handleGetInventoryItems());
      case "getInventoryCategories": return jsonResponse(handleGetInventoryCategories());
      case "getInventoryLedger": return jsonResponse(handleGetInventoryLedger(e.parameter));
      case "getInventorySummary": return jsonResponse(handleGetInventorySummary());
      case "getBlends":         return jsonResponse(handleGetBlends(e.parameter));
      case "getDrafts":         return jsonResponse(handleGetDrafts(e.parameter, user));
      case "getClassifications": return jsonResponse(handleGetClassifications());
      case "getResponseChain": return jsonResponse(handleGetResponseChain(e.parameter));
      case "response-chain":   return jsonResponse(handleGetResponseChain(e.parameter));
      case "getOrderStageTemplates": return jsonResponse(getOrderStageTemplatesConfig());
      default:                  return jsonResponse({ error: "Unknown action: " + action });
    }
  } catch (err) { return jsonResponse({ error: err.message }); }
}

function doPost(e) {
  clearRowsCache();
  var body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) { return jsonResponse({ error: "Invalid JSON body" }); }
  var action = body.action || "";
  try {
    if (action === "login") return jsonResponse(handleLogin(body));
    var token = body.token || "";
    var user = validateToken(token);
    if (!user) return jsonResponse({ error: "AUTH_EXPIRED" });
    switch (action) {
      case "logout":            return jsonResponse(handleLogout(body));
      case "createOrder":       return jsonResponse(handleCreateOrder(body, user));
      case "deleteOrder":       var e1 = requireAdmin(user); if (e1) return jsonResponse(e1); return jsonResponse(handleDeleteOrder(body, user));
      case "editOrder":         return jsonResponse(handleEditOrder(body, user));
      case "updateOrderStatus": return jsonResponse(handleUpdateOrderStatus(body, user));
      case "submitChecklist":   return jsonResponse(handleSubmitChecklist(body, user));
      case "editResponse":      var e3 = requireAdmin(user); if (e3) return jsonResponse(e3); return jsonResponse(handleEditResponse(body, user));
      case "revertChecklist":   var e4 = requireAdmin(user); if (e4) return jsonResponse(e4); return jsonResponse(handleRevertChecklist(body, user));
      case "createOrderType":   var e5 = requireAdmin(user); if (e5) return jsonResponse(e5); return jsonResponse(handleCreateOrderType(body, user));
      case "deleteOrderType":   var e6 = requireAdmin(user); if (e6) return jsonResponse(e6); return jsonResponse(handleDeleteOrderType(body, user));
      case "createCustomer":    var e7 = requireAdmin(user); if (e7) return jsonResponse(e7); return jsonResponse(handleCreateCustomer(body, user));
      case "deleteCustomer":    var e8 = requireAdmin(user); if (e8) return jsonResponse(e8); return jsonResponse(handleDeleteCustomer(body, user));
      case "createChecklist":   var e9 = requireAdmin(user); if (e9) return jsonResponse(e9); return jsonResponse(handleCreateChecklist(body, user));
      case "updateChecklist":   var e10 = requireAdmin(user); if (e10) return jsonResponse(e10); return jsonResponse(handleUpdateChecklist(body, user));
      case "deleteChecklist":   var e11 = requireAdmin(user); if (e11) return jsonResponse(e11); return jsonResponse(handleDeleteChecklist(body, user));
      case "createRule":        var e12 = requireAdmin(user); if (e12) return jsonResponse(e12); return jsonResponse(handleCreateRule(body, user));
      case "updateRule":        var e13 = requireAdmin(user); if (e13) return jsonResponse(e13); return jsonResponse(handleUpdateRule(body, user));
      case "deleteRule":        var e14 = requireAdmin(user); if (e14) return jsonResponse(e14); return jsonResponse(handleDeleteRule(body, user));
      case "createUser":        var e15 = requireAdmin(user); if (e15) return jsonResponse(e15); return jsonResponse(handleCreateUser(body, user));
      case "updateUser":        var e16 = requireAdmin(user); if (e16) return jsonResponse(e16); return jsonResponse(handleUpdateUser(body, user));
      case "resetPassword":     var e17 = requireAdmin(user); if (e17) return jsonResponse(e17); return jsonResponse(handleResetPassword(body, user));
      case "changePassword":    return jsonResponse(handleChangePassword(body, user));
      case "archiveOrders":     var e18 = requireAdmin(user); if (e18) return jsonResponse(e18); return jsonResponse(handleArchiveOrders(body, user));
      case "submitUntagged":    return jsonResponse(handleSubmitUntagged(body, user));
      case "tagUntagged":       return jsonResponse(handleTagUntagged(body, user));
      case "createInventoryItem": var ei1 = requireAdmin(user); if (ei1) return jsonResponse(ei1); return jsonResponse(handleCreateInventoryItem(body, user));
      case "updateInventoryItem": var ei2 = requireAdmin(user); if (ei2) return jsonResponse(ei2); return jsonResponse(handleUpdateInventoryItem(body, user));
      case "addInventoryAdjustment": var ei3 = requireAdmin(user); if (ei3) return jsonResponse(ei3); return jsonResponse(handleAddInventoryAdjustment(body, user));
      case "editInventoryLedger": var ei4 = requireAdmin(user); if (ei4) return jsonResponse(ei4); return jsonResponse(handleEditInventoryLedger(body, user));
      case "createInventoryCategory": var ei5 = requireAdmin(user); if (ei5) return jsonResponse(ei5); return jsonResponse(handleCreateInventoryCategory(body, user));
      case "createBlend":       var eb1 = requireAdmin(user); if (eb1) return jsonResponse(eb1); return jsonResponse(handleCreateBlend(body, user));
      case "updateBlend":       var eb2 = requireAdmin(user); if (eb2) return jsonResponse(eb2); return jsonResponse(handleUpdateBlend(body, user));
      case "deleteBlend":       var eb3 = requireAdmin(user); if (eb3) return jsonResponse(eb3); return jsonResponse(handleDeleteBlend(body, user));
      case "saveDraft":         return jsonResponse(handleSaveDraft(body, user));
      case "deleteDraft":       return jsonResponse(handleDeleteDraft(body, user));
      case "createAllocation":  return jsonResponse(handleCreateAllocation(body, user));
      case "saveOrderTypeRequirements": var eOTR = requireAdmin(user); if (eOTR) return jsonResponse(eOTR); return jsonResponse(handleSaveOrderTypeRequirements(body, user));
      case "saveOrderStageTemplates": var eOST = requireAdmin(user); if (eOST) return jsonResponse(eOST); return jsonResponse(handleSaveOrderStageTemplates(body, user));
      case "tagChecklistToStage": return jsonResponse(handleTagChecklistToStage(body, user));
      case "untagChecklistFromStage": return jsonResponse(handleUntagChecklistFromStage(body, user));
      case "deliverOrder": return jsonResponse(handleDeliverOrder(body, user));
      case "editUntaggedResponse": return jsonResponse(handleEditUntaggedResponse(body, user));
      case "softDeleteChecklist": return jsonResponse(handleSoftDeleteChecklist(body, user));
      case "addClassification":
        var eC1 = requireAdmin(user); if (eC1) return jsonResponse(eC1);
        return jsonResponse(handleAddClassification(body, user));
      case "editClassification":
        var eC2 = requireAdmin(user); if (eC2) return jsonResponse(eC2);
        return jsonResponse(handleEditClassification(body, user));
      case "deactivateClassification":
        var eC3 = requireAdmin(user); if (eC3) return jsonResponse(eC3);
        return jsonResponse(handleDeactivateClassification(body, user));
      case "fixRoastedBeansTemplateOrder":
        var eFix = requireAdmin(user); if (eFix) return jsonResponse(eFix);
        return jsonResponse(handleFixRoastedBeansTemplateOrder(body, user));
      case "backfillLedgerCategories":
        var eBf = requireAdmin(user); if (eBf) return jsonResponse(eBf);
        _ledgerCategoryBackfillDone = false;
        return jsonResponse(backfillInventoryLedgerCategories());
      case "runTests":
        var eTest = requireAdmin(user); if (eTest) return jsonResponse(eTest);
        return jsonResponse(handleRunTests(user));
      default:                  return jsonResponse({ error: "Unknown action: " + action });
    }
  } catch (err) { return jsonResponse({ error: err.message }); }
}

// ─── Init (Batch Load) ────────────────────────────────────────

function handleInit(user) {
  ensureAllSheetColumnsMigrated();
  var checklists = handleGetChecklists();
  return {
    orders: handleGetOrders(),
    checklists: checklists,
    orderTypes: handleGetOrderTypes(),
    customers: handleGetCustomers(),
    rules: handleGetRules(),
    untaggedChecklists: handleGetUntagged(),
    approvedEntries: buildApprovedEntriesCache(checklists),
    inventoryItems: handleGetInventoryItems(),
    inventoryCategories: handleGetInventoryCategories(),
    inventorySummary: handleGetInventorySummary(),
    blends: handleGetBlends({}),
    drafts: user ? handleGetDrafts({}, user) : [],
    orderTypeRequirements: getOrderTypeRequirementsConfig(),
    orderStageTemplates: getOrderStageTemplatesConfig(),
    classifications: handleGetClassifications(),
  };
}

// ─── Auth ──────────────────────────────────────────────────────

function handleLogin(body) {
  var username = (body.username || "").trim().toLowerCase();
  var password = body.password || "";
  if (!username || !password) return { error: "Username and password required" };
  var users = getRows(SHEETS.USERS);
  var user = null;
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].username).toLowerCase() === username && users[i].status === "active") { user = users[i]; break; }
  }
  if (!user) return { error: "Invalid username or password" };
  if (String(user.password_hash) !== hashPassword(password)) return { error: "Invalid username or password" };
  var token = generateToken();
  var now = new Date().toISOString();
  appendToSheet(SHEETS.SESSIONS, { token: token, user_id: user.id, created_at: now, last_active: now });
  return { token: token, user: { id: user.id, username: user.username, displayName: user.display_name, role: user.role } };
}

function handleLogout(body) {
  var token = body.token || "";
  if (token) { var idx = findRowIndex(SHEETS.SESSIONS, token); if (idx > 0) deleteSheetRow(SHEETS.SESSIONS, idx); }
  return { success: true };
}

// ─── User Management ───────────────────────────────────────────

function handleGetUsers() {
  return getRows(SHEETS.USERS).map(function(u) {
    return { id: u.id, username: u.username, displayName: u.display_name, role: u.role, status: u.status, createdAt: u.created_at };
  });
}

function handleCreateUser(body, adminUser) {
  var username = (body.username || "").trim().toLowerCase();
  if (!username) return { error: "Username required" };
  var existing = getRows(SHEETS.USERS);
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i].username).toLowerCase() === username) return { error: "Username already exists" };
  }
  var id = "user_" + nextId();
  var obj = {
    id: id, username: username, password_hash: hashPassword(body.password || "changeme"),
    display_name: body.displayName || username, role: body.role || "user",
    status: "active", created_at: new Date().toISOString(), created_by: adminUser.id,
  };
  appendToSheet(SHEETS.USERS, obj);
  writeAuditLog(adminUser, "create_user", "User", id, "Created user: " + username);
  return { id: id, username: username, displayName: obj.display_name, role: obj.role, status: "active" };
}

function handleUpdateUser(body, adminUser) {
  var id = body.id;
  var idx = findRowIndex(SHEETS.USERS, id);
  if (idx < 0) return { error: "User not found" };
  var users = getRows(SHEETS.USERS);
  var user = null;
  for (var i = 0; i < users.length; i++) { if (String(users[i].id) === String(id)) { user = users[i]; break; } }
  if (!user) return { error: "User not found" };
  if (body.displayName !== undefined) user.display_name = body.displayName;
  if (body.role !== undefined) user.role = body.role;
  if (body.status !== undefined) user.status = body.status;
  updateSheetRow(SHEETS.USERS, idx, user);
  writeAuditLog(adminUser, "update_user", "User", id, "Updated: " + user.username);
  return { success: true };
}

function handleResetPassword(body, adminUser) {
  var id = body.id;
  var idx = findRowIndex(SHEETS.USERS, id);
  if (idx < 0) return { error: "User not found" };
  var users = getRows(SHEETS.USERS);
  var user = null;
  for (var i = 0; i < users.length; i++) { if (String(users[i].id) === String(id)) { user = users[i]; break; } }
  if (!user) return { error: "User not found" };
  user.password_hash = hashPassword(body.newPassword || "changeme");
  updateSheetRow(SHEETS.USERS, idx, user);
  writeAuditLog(adminUser, "reset_password", "User", id, "Reset password for: " + user.username);
  return { success: true };
}

// Self-service password change. Token already validates the caller; we re-verify the current password against the stored hash.
function handleChangePassword(body, sessionUser) {
  var currentPassword = body.currentPassword || "";
  var newPassword = body.newPassword || "";
  if (!currentPassword || !newPassword) return { error: "Current and new password required" };
  if (String(newPassword).length < 6) return { error: "New password must be at least 6 characters" };
  var idx = findRowIndex(SHEETS.USERS, sessionUser.id);
  if (idx < 0) return { error: "User not found" };
  var users = getRows(SHEETS.USERS);
  var user = null;
  for (var i = 0; i < users.length; i++) { if (String(users[i].id) === String(sessionUser.id)) { user = users[i]; break; } }
  if (!user) return { error: "User not found" };
  if (String(user.password_hash) !== hashPassword(currentPassword)) return { error: "Current password is incorrect" };
  user.password_hash = hashPassword(newPassword);
  updateSheetRow(SHEETS.USERS, idx, user);
  writeAuditLog(sessionUser, "change_password", "User", sessionUser.id, "Self password change");
  return { success: true };
}

// ─── Order Types ───────────────────────────────────────────────

function handleGetOrderTypes() { return getRows(SHEETS.ORDER_TYPES); }

function handleCreateOrderType(body, user) {
  var id = body.id || ("ot_" + nextId());
  var obj = { id: id, label: body.label };
  appendToSheet(SHEETS.ORDER_TYPES, obj);
  writeAuditLog(user, "create", "OrderType", id, body.label);
  return obj;
}

function handleDeleteOrderType(body, user) {
  var id = body.id;
  var idx = findRowIndex(SHEETS.ORDER_TYPES, id);
  if (idx > 0) deleteSheetRow(SHEETS.ORDER_TYPES, idx);
  var rules = getRows(SHEETS.RULES);
  for (var i = rules.length - 1; i >= 0; i--) {
    if (rules[i].order_type_id === id) { var rIdx = findRowIndex(SHEETS.RULES, rules[i].id); if (rIdx > 0) deleteSheetRow(SHEETS.RULES, rIdx); }
  }
  writeAuditLog(user, "delete", "OrderType", id, "");
  return { success: true };
}

// ─── Customers ─────────────────────────────────────────────────

function handleGetCustomers() { return getRows(SHEETS.CUSTOMERS); }

function handleCreateCustomer(body, user) {
  var id = body.id || ("cust_" + nextId());
  var obj = { id: id, label: body.label };
  appendToSheet(SHEETS.CUSTOMERS, obj);
  writeAuditLog(user, "create", "Customer", id, body.label);
  return obj;
}

function handleDeleteCustomer(body, user) {
  var id = body.id;
  var idx = findRowIndex(SHEETS.CUSTOMERS, id);
  if (idx > 0) deleteSheetRow(SHEETS.CUSTOMERS, idx);
  var rules = getRows(SHEETS.RULES);
  for (var i = rules.length - 1; i >= 0; i--) {
    if (rules[i].customer_id === id) { var rIdx = findRowIndex(SHEETS.RULES, rules[i].id); if (rIdx > 0) deleteSheetRow(SHEETS.RULES, rIdx); }
  }
  writeAuditLog(user, "delete", "Customer", id, "");
  return { success: true };
}

// ─── Checklists ────────────────────────────────────────────────

function handleGetChecklists() {
  ensureSheetHasAllColumns(SHEETS.CHECKLISTS);
  return getRows(SHEETS.CHECKLISTS).map(function(r) {
    return {
      id: r.id, name: r.name, subtitle: r.subtitle || "", formUrl: r.form_url || "",
      questions: normalizeQuestions(safeParseJSON(r.questions, [])),
      autoIdConfig: normalizeAutoIdConfig(safeParseJSON(r.auto_id_config, null)),
      canTagTo: safeParseJSON(r.can_tag_to, []),
    };
  });
}

function handleCreateChecklist(body, user) {
  ensureSheetHasAllColumns(SHEETS.CHECKLISTS);
  var id = body.id || ("ck_" + nextId());
  var questions = normalizeQuestions(body.questions || []);
  var autoIdConfig = normalizeAutoIdConfig(body.autoIdConfig || null);
  var canTagTo = Array.isArray(body.canTagTo) ? body.canTagTo : [];
  var obj = {
    id: id, name: body.name, subtitle: body.subtitle || "", form_url: body.formUrl || "",
    questions: JSON.stringify(questions),
    auto_id_config: autoIdConfig ? JSON.stringify(autoIdConfig) : "",
    can_tag_to: JSON.stringify(canTagTo),
  };
  appendToSheet(SHEETS.CHECKLISTS, obj);
  getOrCreateResponseSheet(body.name, questions);
  writeAuditLog(user, "create", "Checklist", id, body.name);
  return { id: id, name: body.name, subtitle: body.subtitle || "", formUrl: body.formUrl || "", questions: questions, autoIdConfig: autoIdConfig, canTagTo: canTagTo };
}

function handleUpdateChecklist(body, user) {
  ensureSheetHasAllColumns(SHEETS.CHECKLISTS);
  var id = body.id;
  var idx = findRowIndex(SHEETS.CHECKLISTS, id);
  if (idx < 0) return { error: "Checklist not found" };
  // Read existing row first to preserve all stored data
  var rows = getRows(SHEETS.CHECKLISTS);
  var existing = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { existing = rows[i]; break; } }
  if (!existing) return { error: "Checklist not found" };
  // Parse existing questions so we can merge metadata from them
  var existingQuestions = normalizeQuestions(safeParseJSON(existing.questions, []));
  // Merge incoming questions with existing question metadata to prevent data loss
  var incomingQuestions = body.questions || [];
  var mergedQuestions = normalizeQuestions(incomingQuestions.map(function(q, qi) {
    if (typeof q === "string") return q;
    // Try to find matching existing question by text to preserve metadata not in the incoming payload
    var match = null;
    for (var m = 0; m < existingQuestions.length; m++) {
      if (existingQuestions[m].text === q.text) { match = existingQuestions[m]; break; }
    }
    if (match) {
      // Merge: existing fields are the base, incoming fields override
      var merged = {};
      for (var k in match) { if (match.hasOwnProperty(k)) merged[k] = match[k]; }
      for (var k2 in q) { if (q.hasOwnProperty(k2) && q[k2] !== undefined) merged[k2] = q[k2]; }
      return merged;
    }
    return q;
  }));
  var autoIdConfig = normalizeAutoIdConfig(body.autoIdConfig !== undefined ? body.autoIdConfig : safeParseJSON(existing.auto_id_config, null));
  var canTagTo = body.canTagTo !== undefined ? (Array.isArray(body.canTagTo) ? body.canTagTo : []) : safeParseJSON(existing.can_tag_to, []);
  // Merge into existing row to preserve any fields not in the update payload
  existing.name = body.name !== undefined ? body.name : existing.name;
  existing.subtitle = body.subtitle !== undefined ? body.subtitle : (existing.subtitle || "");
  existing.form_url = body.formUrl !== undefined ? body.formUrl : (existing.form_url || "");
  existing.questions = JSON.stringify(mergedQuestions);
  existing.auto_id_config = autoIdConfig ? JSON.stringify(autoIdConfig) : "";
  existing.can_tag_to = JSON.stringify(canTagTo);
  updateSheetRow(SHEETS.CHECKLISTS, idx, existing);
  // Ensure response sheet has remark columns if questions changed
  ensureResponseSheetColumns(existing.name, mergedQuestions);
  writeAuditLog(user, "update", "Checklist", id, existing.name);
  // Reference check: collect all auto_ids for responses of this checklist and see if anything downstream references them.
  var refWarning = "";
  try {
    var refAutoIds = [];
    var ocsForCk = getRows(SHEETS.ORDER_CHECKLISTS);
    for (var oi = 0; oi < ocsForCk.length; oi++) {
      if (String(ocsForCk[oi].checklist_id) === String(id) && ocsForCk[oi].auto_id) refAutoIds.push(String(ocsForCk[oi].auto_id));
    }
    var utsForCk = getRows(SHEETS.UNTAGGED_CHECKLISTS);
    for (var ui = 0; ui < utsForCk.length; ui++) {
      if (String(utsForCk[ui].checklist_id) === String(id) && utsForCk[ui].auto_id) refAutoIds.push(String(utsForCk[ui].auto_id));
    }
    var seenRefs = {};
    var allRefs = [];
    for (var ai = 0; ai < refAutoIds.length; ai++) {
      var refs = findUpstreamReferencesForAutoId(refAutoIds[ai]);
      for (var ri = 0; ri < refs.length; ri++) {
        if (!seenRefs[refs[ri]]) { seenRefs[refs[ri]] = true; allRefs.push(refs[ri]); }
      }
    }
    if (allRefs.length > 0) refWarning = "Referenced by: " + allRefs.join(", ");
  } catch (e) { /* non-fatal */ }
  var result = { id: id, name: existing.name, subtitle: existing.subtitle || "", formUrl: existing.form_url || "", questions: mergedQuestions, autoIdConfig: autoIdConfig, canTagTo: canTagTo, success: true };
  if (refWarning) result.warning = refWarning;
  return result;
}

function handleDeleteChecklist(body, user) {
  var id = body.id;
  var idx = findRowIndex(SHEETS.CHECKLISTS, id);
  if (idx > 0) deleteSheetRow(SHEETS.CHECKLISTS, idx);
  var rules = getRows(SHEETS.RULES);
  for (var i = 0; i < rules.length; i++) {
    var ckIds = safeParseJSON(rules[i].checklist_ids, []);
    var filtered = ckIds.filter(function(x) { return x !== id; });
    if (filtered.length !== ckIds.length) {
      var rIdx = findRowIndex(SHEETS.RULES, rules[i].id);
      if (rIdx > 0) { rules[i].checklist_ids = JSON.stringify(filtered); updateSheetRow(SHEETS.RULES, rIdx, rules[i]); }
    }
  }
  writeAuditLog(user, "delete", "Checklist", id, "");
  return { success: true };
}

// ─── Assignment Rules ──────────────────────────────────────────

function handleGetRules() {
  return getRows(SHEETS.RULES).map(function(r) {
    return { id: r.id, orderTypeId: r.order_type_id, customerId: r.customer_id, checklistIds: safeParseJSON(r.checklist_ids, []) };
  });
}

function handleCreateRule(body, user) {
  var id = body.id || ("rule_" + nextId());
  var obj = { id: id, order_type_id: body.orderTypeId, customer_id: body.customerId, checklist_ids: JSON.stringify(body.checklistIds || []) };
  appendToSheet(SHEETS.RULES, obj);
  writeAuditLog(user, "create", "Rule", id, "");
  return { id: id, orderTypeId: body.orderTypeId, customerId: body.customerId, checklistIds: body.checklistIds || [] };
}

function handleUpdateRule(body, user) {
  var id = body.id;
  var idx = findRowIndex(SHEETS.RULES, id);
  if (idx < 0) return { error: "Rule not found" };
  // Read existing row first, then merge
  var rows = getRows(SHEETS.RULES);
  var rule = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { rule = rows[i]; break; } }
  if (!rule) return { error: "Rule not found" };
  if (body.orderTypeId !== undefined) rule.order_type_id = body.orderTypeId;
  if (body.customerId !== undefined) rule.customer_id = body.customerId;
  if (body.checklistIds !== undefined) rule.checklist_ids = JSON.stringify(body.checklistIds || []);
  updateSheetRow(SHEETS.RULES, idx, rule);
  writeAuditLog(user, "update", "Rule", id, "");
  return { id: id, orderTypeId: rule.order_type_id, customerId: rule.customer_id, checklistIds: safeParseJSON(rule.checklist_ids, []) };
}

function handleDeleteRule(body, user) {
  var idx = findRowIndex(SHEETS.RULES, body.id);
  if (idx > 0) deleteSheetRow(SHEETS.RULES, idx);
  writeAuditLog(user, "delete", "Rule", body.id, "");
  return { success: true };
}

// ─── Resolve Checklists ────────────────────────────────────────

function handleResolveChecklists(params) {
  var orderTypeId = params.order_type_id || "";
  var customerId = params.customer_id || "";
  var allRules = getRows(SHEETS.RULES).map(function(r) {
    return { order_type_id: r.order_type_id, customer_id: r.customer_id, checklistIds: safeParseJSON(r.checklist_ids, []) };
  });
  var specific = allRules.filter(function(r) { return r.order_type_id === orderTypeId && r.customer_id === customerId; });
  if (specific.length > 0) return dedup(specific);
  var byType = allRules.filter(function(r) { return r.order_type_id === orderTypeId && r.customer_id === "any"; });
  if (byType.length > 0) return dedup(byType);
  var byCust = allRules.filter(function(r) { return r.order_type_id === "any" && r.customer_id === customerId; });
  if (byCust.length > 0) return dedup(byCust);
  var fallback = allRules.filter(function(r) { return r.order_type_id === "any" && r.customer_id === "any"; });
  return dedup(fallback);
}

function dedup(rules) {
  var all = [];
  rules.forEach(function(r) { all = all.concat(r.checklistIds); });
  return all.filter(function(v, i, a) { return a.indexOf(v) === i; });
}

// ─── Orders ────────────────────────────────────────────────────

function handleGetOrders() {
  var orders = getRows(SHEETS.ORDERS);
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  orders.sort(function(a, b) { return String(b.created_at).localeCompare(String(a.created_at)); });
  return orders.map(function(o) { return formatOrder(o, ocs); });
}

function handleGetOrder(params) {
  var id = params.id;
  var orders = getRows(SHEETS.ORDERS);
  var order = null;
  for (var i = 0; i < orders.length; i++) { if (String(orders[i].id) === String(id)) { order = orders[i]; break; } }
  if (!order) return { error: "Order not found" };
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  return formatOrder(order, ocs);
}

function formatOrder(o, ocs) {
  var orderOcs = ocs.filter(function(c) { return String(c.order_id) === String(o.id); });
  var rawStatus = o.status || "active";
  // Backward compat: treat "active" as "beans_not_roasted"
  var orderStatus = rawStatus === "active" ? "beans_not_roasted" : rawStatus;
  // Phase 6 Fix 2: canTag flag — frontend should use this to filter "Tag to Order" dropdowns
  var canTag = (orderStatus !== "delivered" && orderStatus !== "cancelled");
  return {
    id: o.id, name: o.name, customerId: o.customer_id, assignedTo: o.assigned_to || "",
    orderType: o.order_type, createdAt: o.created_at,
    invoiceSo: o.invoice_so || "", orderTypeDetail: o.order_type_detail || "",
    status: orderStatus, canTag: canTag,
    orderLines: safeParseJSON(o.order_lines, []),
    productType: o.product_type || "",
    missingChecklistReasons: safeParseJSON(o.missing_checklist_reasons, {}),
    stages: safeParseJSON(o.stages, []),
    deliveredAt: o.delivered_at || "",
    checklists: orderOcs.map(function(c) {
      return { id: c.id, checklistId: c.checklist_id, status: c.status || "pending", completedAt: c.completed_at || null, completedBy: c.completed_by || null, workDate: c.work_date || null };
    }),
  };
}

function handleCreateOrder(body, user) {
  var orderId = getNextOrderId(); // ORD-001 format
  var orderLines = body.orderLines || [];
  // Ensure each line has taggedQuantity initialized to 0; persist blend snapshot for versioning
  orderLines = orderLines.map(function(l) { return {
    blendId: l.blendId || "",
    blend: l.blend || "",
    blendComponents: Array.isArray(l.blendComponents) ? l.blendComponents : [],
    quantity: parseFloat(l.quantity) || 0,
    deliveryDate: l.deliveryDate || "",
    taggedQuantity: 0,
  }; });
  // Auto-populate stages from product type template if not supplied
  var stages = Array.isArray(body.stages) ? body.stages : [];
  if ((!stages || stages.length === 0) && body.productType) {
    var templates = getOrderStageTemplatesConfig();
    var tpl = templates[body.productType] || [];
    stages = tpl.map(function(s, i) {
      return {
        id: "stage_" + nextId() + "_" + i,
        name: s.name || ("Stage " + (i+1)),
        checklistId: s.checklistId || "",
        quantityField: s.quantityField || "",
        requiredQty: parseFloat(s.requiredQty) || 0,
        position: i,
        taggedEntries: [],
        advanced: false,
      };
    });
  } else if (Array.isArray(stages)) {
    stages = stages.map(function(s, i) {
      return {
        id: s.id || ("stage_" + nextId() + "_" + i),
        name: s.name || ("Stage " + (i+1)),
        checklistId: s.checklistId || "",
        quantityField: s.quantityField || "",
        requiredQty: parseFloat(s.requiredQty) || 0,
        position: i,
        taggedEntries: Array.isArray(s.taggedEntries) ? s.taggedEntries : [],
        advanced: s.advanced === true,
      };
    });
  }
  var orderObj = {
    id: orderId, name: body.name, customer_id: body.customerId,
    assigned_to: body.assignedTo || "", order_type: body.orderType,
    created_at: body.createdAt || new Date().toISOString(), status: "beans_not_roasted",
    invoice_so: body.invoiceSo || "", order_type_detail: body.orderTypeDetail || "",
    order_lines: JSON.stringify(orderLines),
    product_type: body.productType || "",
    missing_checklist_reasons: "",
    stages: JSON.stringify(stages),
    delivered_at: "",
  };
  appendToSheet(SHEETS.ORDERS, orderObj);
  var checklists = body.checklists || [];
  var createdOcs = [];
  for (var i = 0; i < checklists.length; i++) {
    var ocId = "oc_" + nextId() + "_" + i;
    var oc = { id: ocId, order_id: orderId, checklist_id: checklists[i].checklistId, status: "pending", completed_at: "", completed_by: "", work_date: "" };
    appendToSheet(SHEETS.ORDER_CHECKLISTS, oc);
    createdOcs.push({ id: ocId, checklistId: checklists[i].checklistId, status: "pending", completedAt: null, completedBy: null, workDate: null });
  }
  writeAuditLog(user, "create", "Order", orderId, body.name);
  return { id: orderId, name: body.name, customerId: body.customerId, assignedTo: body.assignedTo || "", orderType: body.orderType, createdAt: orderObj.created_at, invoiceSo: orderObj.invoice_so, orderTypeDetail: orderObj.order_type_detail, status: "beans_not_roasted", orderLines: orderLines, productType: body.productType || "", missingChecklistReasons: {}, stages: stages, deliveredAt: "", checklists: createdOcs };
}

function handleDeleteOrder(body, user) {
  var orderId = body.id;
  // Reference check: block if any stage has tagged checklists
  var ordersForCheck = getRows(SHEETS.ORDERS);
  for (var oi = 0; oi < ordersForCheck.length; oi++) {
    if (String(ordersForCheck[oi].id) === String(orderId)) {
      var stageArr = safeParseJSON(ordersForCheck[oi].stages, []);
      var hasTagged = false;
      for (var si = 0; si < stageArr.length; si++) {
        if (Array.isArray(stageArr[si].taggedEntries) && stageArr[si].taggedEntries.length > 0) { hasTagged = true; break; }
      }
      if (hasTagged) {
        writeAuditLog(user, "delete_blocked", "Order", orderId, "Blocked: stages have tagged checklists");
        return { error: "Remove all tagged checklists from order stages first." };
      }
      break;
    }
  }
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS).filter(function(c) { return String(c.order_id) === String(orderId); });

  // Delete responses from per-checklist tabs + Master Summary
  for (var i = 0; i < ocs.length; i++) {
    var ck = lookupChecklist(ocs[i].checklist_id);
    if (ck) {
      deleteResponseRow(ck.name, orderId);
    }
  }
  deleteFromMasterSummary(orderId, null); // Delete all entries for this order

  // Delete OrderChecklists
  for (var j = ocs.length - 1; j >= 0; j--) {
    var ocIdx = findRowIndex(SHEETS.ORDER_CHECKLISTS, ocs[j].id);
    if (ocIdx > 0) deleteSheetRow(SHEETS.ORDER_CHECKLISTS, ocIdx);
  }
  var idx = findRowIndex(SHEETS.ORDERS, orderId);
  if (idx > 0) deleteSheetRow(SHEETS.ORDERS, idx);
  writeAuditLog(user, "delete", "Order", orderId, "");
  return { success: true };
}

function handleEditOrder(body, user) {
  var id = body.id;
  var idx = findRowIndex(SHEETS.ORDERS, id);
  if (idx < 0) return { error: "Order not found" };
  var orders = getRows(SHEETS.ORDERS);
  var order = null;
  for (var i = 0; i < orders.length; i++) { if (String(orders[i].id) === String(id)) { order = orders[i]; break; } }
  if (!order) return { error: "Order not found" };
  var changes = [];
  if (body.name !== undefined && body.name !== order.name) { changes.push("name: " + order.name + " -> " + body.name); order.name = body.name; }
  if (body.customerId !== undefined && body.customerId !== order.customer_id) { changes.push("customer changed"); order.customer_id = body.customerId; }
  if (body.assignedTo !== undefined && body.assignedTo !== order.assigned_to) { changes.push("assigned: " + order.assigned_to + " -> " + body.assignedTo); order.assigned_to = body.assignedTo; }
  if (body.invoiceSo !== undefined && body.invoiceSo !== order.invoice_so) { changes.push("invoice_so changed"); order.invoice_so = body.invoiceSo; }
  if (body.orderTypeDetail !== undefined && body.orderTypeDetail !== order.order_type_detail) { changes.push("order_type_detail changed"); order.order_type_detail = body.orderTypeDetail; }
  if (body.orderLines !== undefined) { changes.push("order_lines updated"); order.order_lines = JSON.stringify(body.orderLines); }
  if (body.productType !== undefined && body.productType !== order.product_type) { changes.push("product_type: " + (order.product_type || "(none)") + " -> " + body.productType); order.product_type = body.productType; }
  if (body.missingChecklistReasons !== undefined) { changes.push("missing_checklist_reasons updated"); order.missing_checklist_reasons = JSON.stringify(body.missingChecklistReasons); }
  if (body.stages !== undefined && Array.isArray(body.stages)) {
    changes.push("stages updated");
    var normalizedStages = body.stages.map(function(s, i) {
      return {
        id: s.id || ("stage_" + nextId() + "_" + i),
        name: s.name || ("Stage " + (i+1)),
        checklistId: s.checklistId || "",
        quantityField: s.quantityField || "",
        requiredQty: parseFloat(s.requiredQty) || 0,
        position: typeof s.position === "number" ? s.position : i,
        taggedEntries: Array.isArray(s.taggedEntries) ? s.taggedEntries : [],
        advanced: s.advanced === true,
      };
    });
    order.stages = JSON.stringify(normalizedStages);
  }
  if (body.status !== undefined && body.status !== order.status) { changes.push("status: " + order.status + " -> " + body.status); order.status = body.status; }
  updateSheetRow(SHEETS.ORDERS, idx, order);
  writeAuditLog(user, "edit", "Order", id, changes.join("; "));
  return { success: true };
}

// ─── Update Order Status ──────────────────────────────────────

function handleUpdateOrderStatus(body, user) {
  var id = body.id;
  var newStatus = body.status;
  var validStatuses = ["beans_not_roasted", "beans_roasted", "packed", "completed", "delivered"];
  if (validStatuses.indexOf(newStatus) < 0) return { error: "Invalid status: " + newStatus };

  var idx = findRowIndex(SHEETS.ORDERS, id);
  if (idx < 0) return { error: "Order not found" };
  var orders = getRows(SHEETS.ORDERS);
  var order = null;
  for (var i = 0; i < orders.length; i++) { if (String(orders[i].id) === String(id)) { order = orders[i]; break; } }
  if (!order) return { error: "Order not found" };

  var oldStatus = order.status || "beans_not_roasted";
  order.status = newStatus;
  updateSheetRow(SHEETS.ORDERS, idx, order);

  // When moved to "delivered", trigger inventory OUT for packed goods
  if (newStatus === "delivered" && oldStatus !== "delivered") {
    var orderLines = safeParseJSON(order.order_lines, []);
    for (var j = 0; j < orderLines.length; j++) {
      var line = orderLines[j];
      if (line.taggedQuantity > 0) {
        // Find packed goods inventory items and create OUT transaction
        var invItems = getRows(SHEETS.INVENTORY_ITEMS);
        for (var k = 0; k < invItems.length; k++) {
          if (String(invItems[k].category).toLowerCase() === "packing items" && invItems[k].is_active !== "false") {
            createInventoryTransaction(invItems[k].id, "OUT", line.taggedQuantity, "order_delivery", id, "Delivered: " + line.blend + " (" + order.name + ")", user.displayName || user.username);
            break; // One OUT per blend line from first active packing item
          }
        }
      }
    }
  }

  writeAuditLog(user, "update_status", "Order", id, "Status: " + oldStatus + " -> " + newStatus);
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  return formatOrder(order, ocs);
}

// ─── Quantity Allocations / Inventory Link Helpers ────────────

// Returns the trackable batch quantity for a checklist submission. Recognizes:
//   - any number question with inventoryLink.txType === "IN" (the new unified concept)
//   - legacy isMasterQuantity flag (kept for backward compat with old data)
function getMasterQuantityFromResponses(questions, responsesMap) {
  if (!Array.isArray(questions)) return 0;
  for (var i = 0; i < questions.length; i++) {
    var q = questions[i];
    if (!q) continue;
    if (q.inventoryLink && q.inventoryLink.enabled && q.inventoryLink.txType === "IN") {
      return parseFloat(responsesMap[i]) || 0;
    }
    if (q.isMasterQuantity) return parseFloat(responsesMap[i]) || 0;
  }
  return 0;
}

function getAllocatedQuantityForAutoId(sourceAutoId, excludeDestAutoId) {
  if (!sourceAutoId) return 0;
  ensureSheetHasAllColumns(SHEETS.QUANTITY_ALLOCATIONS);
  var rows = getRows(SHEETS.QUANTITY_ALLOCATIONS);
  var total = 0;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].source_auto_id) === String(sourceAutoId)) {
      if (excludeDestAutoId && String(rows[i].destination_auto_id) === String(excludeDestAutoId)) continue;
      total += parseFloat(rows[i].allocated_quantity) || 0;
    }
  }
  return total;
}

function createQuantityAllocation(sourceChecklistId, sourceAutoId, totalQty, destType, destId, destAutoId, allocatedQty, allocatedBy) {
  ensureSheetHasAllColumns(SHEETS.QUANTITY_ALLOCATIONS);
  appendToSheet(SHEETS.QUANTITY_ALLOCATIONS, {
    id: "qa_" + nextId(),
    source_checklist_id: sourceChecklistId || "",
    source_auto_id: sourceAutoId || "",
    total_quantity: totalQty || 0,
    destination_type: destType || "checklist",
    destination_id: destId || "",
    destination_auto_id: destAutoId || "",
    allocated_quantity: allocatedQty || 0,
    allocated_at: new Date().toISOString(),
    allocated_by: allocatedBy || "",
  });
}

// Reverse all allocations created with a given destination_auto_id (used on edit/delete).
function reverseAllocationsForDestination(destAutoId) {
  if (!destAutoId) return;
  var sheet = getSheet(SHEETS.QUANTITY_ALLOCATIONS);
  if (!sheet || sheet.getLastRow() < 2) return;
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][6]) === String(destAutoId)) sheet.deleteRow(i + 1);
  }
  invalidateCache(SHEETS.QUANTITY_ALLOCATIONS);
}

// Look up a source submission by its auto id. Returns { totalQuantity, responses } or null.
function getSubmissionByAutoId(autoId, sourceCk) {
  if (!autoId || !sourceCk) return null;
  ensureSheetHasAllColumns(SHEETS.ORDER_CHECKLISTS);
  ensureSheetHasAllColumns(SHEETS.UNTAGGED_CHECKLISTS);
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  for (var i = 0; i < ocs.length; i++) {
    if (String(ocs[i].auto_id) === String(autoId) && String(ocs[i].checklist_id) === String(sourceCk.id)) {
      var resp = readResponseRow(sourceCk.name, sourceCk.questions, String(ocs[i].order_id));
      if (resp) {
        var rmap = {};
        for (var r = 0; r < resp.responses.length; r++) rmap[resp.responses[r].questionIndex] = resp.responses[r].response;
        return { totalQuantity: getMasterQuantityFromResponses(sourceCk.questions, rmap), responses: rmap };
      }
      return { totalQuantity: 0, responses: {} };
    }
  }
  var uts = getRows(SHEETS.UNTAGGED_CHECKLISTS);
  for (var j = 0; j < uts.length; j++) {
    if (isDeleted(uts[j])) continue;
    if (String(uts[j].auto_id) === String(autoId) && String(uts[j].checklist_id) === String(sourceCk.id)) {
      var responses = safeParseJSON(uts[j].responses, []);
      var rmap2 = responsesArrayToMap(responses);
      return {
        totalQuantity: parseFloat(uts[j].total_quantity) || getMasterQuantityFromResponses(sourceCk.questions, rmap2),
        responses: rmap2,
      };
    }
  }
  return null;
}

// Legacy per-checklist inventory routing. Applies the correct IN/OUT split for checklists
// that do not yet have per-question inventoryLink configs. All lookups go through
// findInventoryItemForCategory so that Green Beans / Roasted Beans / Packing Items are
// always separated into the correct category rows via equivalent_items.
//
//   • ck_green_beans    : IN  against Green Beans item (resolved from "Type of Beans")
//   • ck_roasted_beans  : OUT Green Beans (input) + IN Roasted Beans (equivalent of input)
//   • ck_grinding       : OUT Roasted Beans (resolved from tagged Roast ID's bean type) +
//                         IN  Packing Items (equivalent of that same bean type)
//
// `fallbackInItemId` / `fallbackOutItemId` are only used when the response data doesn't
// carry enough context to resolve an item directly — preserves the older client payload.
function applyLegacyInventoryForChecklist(ck, respMap, refType, refId, person, fallbackInItemId, fallbackOutItemId, isEdit, grindClassificationId) {
  if (!ck) return;
  var suffix = isEdit ? " (edited)" : "";

  // Helper: read a value from respMap by question text. Falls back to searching by
  // question index if the caller built the map from questionIndex keys (which happens
  // when the response payload uses index-keyed format). This fixes the "qty=0" bug.
  function readField(fieldName) {
    if (respMap[fieldName] !== undefined && respMap[fieldName] !== "") return respMap[fieldName];
    // Fall back: scan ck.questions for the matching text and try by index
    if (ck && ck.questions) {
      for (var fi = 0; fi < ck.questions.length; fi++) {
        if (ck.questions[fi].text === fieldName && respMap[fi] !== undefined && respMap[fi] !== "") {
          return respMap[fi];
        }
      }
    }
    return respMap[fieldName] || "";
  }
  function readQty(fieldName) {
    var raw = readField(fieldName);
    var v = parseFloat(raw);
    if (isNaN(v) || v === 0) {
      Logger.log("applyLegacyInventory readQty('" + fieldName + "'): raw='" + raw + "' parsed=" + v);
    }
    return isNaN(v) ? 0 : v;
  }

  if (ck.id === "ck_green_beans") {
    var qtyReceived = readQty("Quantity received");
    if (qtyReceived <= 0) return;
    var gbRef = readField("Type of Beans") || fallbackInItemId;
    var gbResolved = findInventoryItemForCategory(gbRef, "Green Beans");
    if (gbResolved.item) {
      var gbNotes = "Green Bean shipment received" + suffix;
      if (gbResolved.warning) gbNotes = gbResolved.warning + " | " + gbNotes;
      createInventoryTransaction(gbResolved.item.id, "IN", qtyReceived, refType, refId, gbNotes, person);
    } else {
      Logger.log("applyLegacyInventory ck_green_beans: " + (gbResolved.warning || "no item resolved"));
    }
    return;
  }

  if (ck.id === "ck_roasted_beans") {
    var qtyInput = readQty("Quantity input");
    var qtyOutput = readQty("Quantity output");
    var beanRef = readField("Type of Beans") || fallbackInItemId;
    if (qtyInput > 0) {
      var gbOut = findInventoryItemForCategory(beanRef, "Green Beans");
      if (gbOut.item) {
        var outNotes = "Used for roasting" + suffix;
        if (gbOut.warning) outNotes = gbOut.warning + " | " + outNotes;
        createInventoryTransaction(gbOut.item.id, "OUT", qtyInput, refType, refId, outNotes, person);
      } else {
        Logger.log("applyLegacyInventory ck_roasted_beans OUT: " + (gbOut.warning || "no item resolved"));
      }
    }
    if (qtyOutput > 0) {
      // Prefer an explicit roasted-beans fallback if caller provided one; otherwise resolve
      // from the bean reference via equivalent_items.
      var rbIn = fallbackOutItemId
        ? findInventoryItemForCategory(fallbackOutItemId, "Roasted Beans")
        : findInventoryItemForCategory(beanRef, "Roasted Beans");
      if (rbIn.item) {
        var inNotes = "Roast batch output" + suffix;
        if (rbIn.warning) inNotes = rbIn.warning + " | " + inNotes;
        createInventoryTransaction(rbIn.item.id, "IN", qtyOutput, refType, refId, inNotes, person);
      } else {
        Logger.log("applyLegacyInventory ck_roasted_beans IN: " + (rbIn.warning || "no item resolved"));
      }
    }
    return;
  }

  if (ck.id === "ck_grinding") {
    var qIn = readQty("Quantity input");
    var qOut = readQty("Quantity output");
    var netWeight = readQty("Total Net weight");
    if (qIn <= 0 && netWeight > 0) qIn = netWeight;
    if (qOut <= 0 && netWeight > 0) qOut = netWeight;

    // Resolve the underlying bean type. Grinding doesn't carry it directly, so follow the
    // tagged Roast ID (auto-id of a prior Roasted Beans QC) to its "Type of Beans" field.
    var roastAutoId = readField("Roast ID") || "";
    var beanRefG = "";
    if (roastAutoId) {
      var roastCk = lookupChecklist("ck_roasted_beans");
      if (roastCk) {
        var roastSub = findSubmissionByAutoId(roastAutoId, roastCk);
        if (roastSub && roastSub.responses) {
          // "Type of Beans" is field index 2 on ck_roasted_beans
          beanRefG = roastSub.responses[2] || "";
        }
      }
    }

    if (qIn > 0) {
      // Roasted Beans item = equivalent of the green beans reference (or the roasted item
      // itself if beanRefG already resolves to a Roasted Beans row).
      var rbOut = beanRefG
        ? findInventoryItemForCategory(beanRefG, "Roasted Beans")
        : (fallbackInItemId ? findInventoryItemForCategory(fallbackInItemId, "Roasted Beans") : { item: null, warning: "No Roast ID or fallback provided" });
      if (rbOut.item) {
        var gOutNotes = "Used for grinding" + suffix;
        if (rbOut.warning) gOutNotes = rbOut.warning + " | " + gOutNotes;
        createInventoryTransaction(rbOut.item.id, "OUT", qIn, refType, refId, gOutNotes, person);
      } else {
        Logger.log("applyLegacyInventory ck_grinding OUT: " + (rbOut.warning || "no item resolved"));
      }
    }
    if (qOut > 0) {
      var pkIn = beanRefG
        ? findInventoryItemForCategory(beanRefG, "Packing Items")
        : (fallbackOutItemId ? findInventoryItemForCategory(fallbackOutItemId, "Packing Items") : { item: null, warning: "No Roast ID or fallback provided" });
      if (pkIn.item) {
        var gInNotes = "Packed goods produced" + suffix;
        if (pkIn.warning) gInNotes = pkIn.warning + " | " + gInNotes;
        createInventoryTransaction(pkIn.item.id, "IN", qOut, refType, refId, gInNotes, person, "", grindClassificationId || "");
      } else {
        Logger.log("applyLegacyInventory ck_grinding IN: " + (pkIn.warning || "no item resolved"));
      }
    }
    return;
  }
}

// Resolve an inventory item by reference (id or name), optionally mapped to a target
// category via the source item's equivalent_items list.
// Returns { item, isEquivalent, warning }:
//   • item: the inventory-items row to write the ledger entry against (may be the source
//     itself if no category mapping is required, or null when nothing resolves).
//   • isEquivalent: true when the returned item is a different row reached via equivalent_items.
//   • warning: non-empty string when the lookup could not find a clean match in the target
//     category — the caller should prepend this to the ledger entry notes so the mismatch
//     is visible to operators.
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
  return {
    item: source,
    isEquivalent: false,
    warning: "[WARN] No '" + targetCategory + "' equivalent linked to '" + source.name + "' — ledger written against source item; stock for " + targetCategory + " needs manual correction"
  };
}

// Process per-question inventoryLink configs to create inventory transactions.
function processInventoryLinks(checklist, responsesMap, refType, refId, doneBy, isEdit) {
  if (!checklist) {
    Logger.log("processInventoryLinks: no checklist, skipping");
    return;
  }
  var nq = checklist.questions || [];
  Logger.log("processInventoryLinks: " + checklist.name + " (" + nq.length + " questions)" + (isEdit ? " [edit]" : ""));

  // Build a fast lookup of existing non-reversal ledger entries keyed by refId + questionIndex + txType.
  // Used as a duplicate-write guard on new submissions.
  var existingKey = {};
  if (refId) {
    var ledgerRows = getRows(SHEETS.INVENTORY_LEDGER);
    for (var li = 0; li < ledgerRows.length; li++) {
      var lrow = ledgerRows[li];
      if (String(lrow.reference_id) !== String(refId)) continue;
      if (String(lrow.notes || "").indexOf("[REVERSAL]") === 0) continue;
      var k = String(refId) + "|" + String(lrow.question_index || "") + "|" + String(lrow.type || "");
      existingKey[k] = true;
    }
  }

  for (var i = 0; i < nq.length; i++) {
    var q = nq[i];
    if (!q.inventoryLink || !q.inventoryLink.enabled) continue;
    var rawQty = responsesMap[i];
    var qty = parseFloat(rawQty);
    Logger.log("  q[" + i + "] " + q.text + " inventoryLink=" + JSON.stringify(q.inventoryLink) + " qty=" + qty);
    // Reject negative or non-numeric quantities — checklist submissions must never write a
    // negative inventory change. The legitimate way to reduce stock below a prior entry is
    // a manual OUT adjustment.
    if (isNaN(qty) || qty < 0) {
      Logger.log("  → WARNING: Skipped negative/invalid qty for " + (q.text || "question " + i) + ": " + rawQty);
      continue;
    }
    if (qty === 0) { Logger.log("  → qty=0, skipping"); continue; }
    var link = q.inventoryLink;
    var itemId = "";
    var linkWarning = "";
    if (link.itemSource && link.itemSource.type === "fixed") {
      itemId = link.itemSource.itemId || "";
      Logger.log("  → fixed item: " + itemId);
    } else if (link.itemSource && link.itemSource.type === "field") {
      var srcVal = responsesMap[link.itemSource.fieldIdx];
      Logger.log("  → field source idx=" + link.itemSource.fieldIdx + " value=" + srcVal);
      if (srcVal) {
        var resolved = findInventoryItemForCategory(srcVal, link.category || "");
        if (resolved.item) {
          itemId = resolved.item.id;
          if (resolved.warning) linkWarning = resolved.warning;
          Logger.log("  → resolved " + (resolved.isEquivalent ? "(via equivalent_items)" : "(direct)") +
                     " to " + itemId + (linkWarning ? " [" + linkWarning + "]" : ""));
        } else {
          Logger.log("  → " + (resolved.warning || "no item resolved"));
        }
      }
    }
    if (itemId) {
      var txType = link.txType || "IN";
      var dupKey = String(refId || "") + "|" + String(i) + "|" + String(txType);
      if (!isEdit && existingKey[dupKey]) {
        Logger.log("  → WARNING: Duplicate inventory entry prevented for refId: " + refId + " (questionIndex=" + i + ", type=" + txType + ")");
        continue;
      }
      var noteBase = "Auto from " + (q.text || "checklist field");
      var notes = linkWarning ? (linkWarning + " | " + noteBase) : noteBase;
      Logger.log("  → createInventoryTransaction(" + itemId + ", " + txType + ", " + qty + ")");
      createInventoryTransaction(itemId, txType, qty, refType || "checklist", refId || "", notes, doneBy || "", i);
      existingKey[dupKey] = true;
    } else {
      Logger.log("  → no itemId resolved, skipping");
    }
  }
}

// New unified allocation processor. Consumes a batchAllocations payload supplied by the client:
//   batchAllocations = { [questionIdx]: [{ sourceAutoId, quantity }, ...] }
// Validates: each row's quantity ≤ remaining-for-that-source (excluding any prior rows from this same destAutoId, for re-submit cases),
// and the OUT-linked field's value (if any) ≤ sum of all allocations.
// Returns { ok, error }. If `commit` is true, also writes the QuantityAllocations rows.
function processQuantityAllocationsForSubmission(checklist, responsesMap, batchAllocations, destAutoId, destType, destId, allocatedBy, excludeDestAutoId, commit) {
  if (!checklist) return { ok: true };
  var nq = checklist.questions || [];
  if (!batchAllocations || typeof batchAllocations !== "object") return { ok: true };
  var queued = [];
  var totalAlloc = 0;
  // Track the running per-source allocation across rows so two rows for the same source don't double-count remaining
  var pendingBySource = {};
  for (var qi = 0; qi < nq.length; qi++) {
    var q = nq[qi];
    if (!q || !q.linkedSource || !q.linkedSource.checklistId) continue;
    var allocs = batchAllocations[qi];
    if (!allocs) allocs = batchAllocations[String(qi)];
    if (!Array.isArray(allocs) || allocs.length === 0) continue;
    var sourceCk = lookupChecklist(q.linkedSource.checklistId);
    if (!sourceCk) continue;
    for (var ai = 0; ai < allocs.length; ai++) {
      var srcAutoId = String((allocs[ai] && allocs[ai].sourceAutoId) || "").trim();
      var amt = parseFloat(allocs[ai] && allocs[ai].quantity) || 0;
      if (!srcAutoId || amt <= 0) continue;
      var srcInfo = getSubmissionByAutoId(srcAutoId, sourceCk);
      var srcTotal = srcInfo ? srcInfo.totalQuantity : 0;
      if (srcTotal <= 0) {
        return { ok: false, error: "Source batch " + srcAutoId + " has no trackable quantity" };
      }
      var alreadyAllocated = getAllocatedQuantityForAutoId(srcAutoId, excludeDestAutoId);
      var pendingHere = pendingBySource[srcAutoId] || 0;
      var remaining = srcTotal - alreadyAllocated - pendingHere;
      if (amt > remaining) {
        return { ok: false, error: "Cannot use " + amt + " from " + srcAutoId + " — only " + remaining + " available" };
      }
      pendingBySource[srcAutoId] = pendingHere + amt;
      queued.push({ sourceCkId: sourceCk.id, sourceAutoId: srcAutoId, srcTotal: srcTotal, amount: amt });
      totalAlloc += amt;
    }
  }
  // Validate the OUT-linked quantity field (if any) against the total of all batch allocations
  if (totalAlloc > 0) {
    for (var oi = 0; oi < nq.length; oi++) {
      var oq = nq[oi];
      if (oq && oq.inventoryLink && oq.inventoryLink.enabled && oq.inventoryLink.txType === "OUT") {
        var outVal = parseFloat(responsesMap[oi]) || 0;
        if (outVal > totalAlloc + 0.0001) {
          return { ok: false, error: "Input quantity (" + outVal + ") exceeds available from tagged batches (" + totalAlloc + "). Add more batches or reduce input." };
        }
      }
    }
  }
  if (commit) {
    for (var k = 0; k < queued.length; k++) {
      createQuantityAllocation(queued[k].sourceCkId, queued[k].sourceAutoId, queued[k].srcTotal, destType, destId, destAutoId, queued[k].amount, allocatedBy);
    }
  }
  return { ok: true };
}

// ─── Multi-Batch Roasting Helper ─────────────────────────────

// Validate and process an array of roast batch objects for ck_roasted_beans.
// Each batch: { sourceAutoId, inputQty, outputQty, reasonForLoss, classificationId }
// Returns { ok, error, processed[] } where processed has resolved item ids for inventory writes.
function validateRoastBatches(batches, excludeDestAutoId) {
  if (!Array.isArray(batches) || batches.length === 0) return { ok: false, error: "No roast batches provided" };
  if (batches.length > 6) return { ok: false, error: "Maximum 6 batches allowed" };
  var greenBeanCk = lookupChecklist("ck_green_beans");
  if (!greenBeanCk) return { ok: false, error: "Green Bean QC template not found" };
  var processed = [];
  var pendingBySource = {};
  for (var i = 0; i < batches.length; i++) {
    var b = batches[i];
    var srcAutoId = String(b.sourceAutoId || "").trim();
    var inputQty = parseFloat(b.inputQty) || 0;
    var outputQty = parseFloat(b.outputQty) || 0;
    if (!srcAutoId) return { ok: false, error: "Batch " + (i + 1) + ": missing source batch" };
    if (inputQty <= 0) return { ok: false, error: "Batch " + (i + 1) + ": input quantity must be > 0" };
    if (outputQty < 0) return { ok: false, error: "Batch " + (i + 1) + ": output quantity cannot be negative" };
    if (outputQty > inputQty) return { ok: false, error: "Batch " + (i + 1) + ": output (" + outputQty + ") cannot exceed input (" + inputQty + ")" };
    // Check remaining qty of source green bean batch
    var srcInfo = getSubmissionByAutoId(srcAutoId, greenBeanCk);
    if (!srcInfo) return { ok: false, error: "Batch " + (i + 1) + ": source batch " + srcAutoId + " not found" };
    var srcTotal = srcInfo.totalQuantity || 0;
    var alreadyAllocated = getAllocatedQuantityForAutoId(srcAutoId, excludeDestAutoId || "");
    var pendingHere = pendingBySource[srcAutoId] || 0;
    var remaining = srcTotal - alreadyAllocated - pendingHere;
    if (inputQty > remaining + 0.01) {
      return { ok: false, error: "Batch " + (i + 1) + ": cannot use " + inputQty + "kg from " + srcAutoId + " — only " + Math.round(remaining * 100) / 100 + "kg available" };
    }
    pendingBySource[srcAutoId] = pendingHere + inputQty;
    // Resolve the green bean inventory item from the source submission's "Type of Beans" field
    var beanRef = srcInfo.responses ? (srcInfo.responses[1] || srcInfo.responses["1"] || "") : "";
    var gbItem = findInventoryItemForCategory(beanRef, "Green Beans");
    var rbItem = findInventoryItemForCategory(beanRef, "Roasted Beans");
    processed.push({
      sourceAutoId: srcAutoId,
      inputQty: inputQty,
      outputQty: outputQty,
      lossQty: Math.round((inputQty - outputQty) * 100) / 100,
      lossPercent: inputQty > 0 ? Math.round((inputQty - outputQty) / inputQty * 1000) / 10 : 0,
      reasonForLoss: b.reasonForLoss || "",
      classificationId: b.classificationId || "",
      greenBeanItemId: gbItem.item ? gbItem.item.id : "",
      greenBeanWarning: gbItem.warning || "",
      roastedBeanItemId: rbItem.item ? rbItem.item.id : "",
      roastedBeanWarning: rbItem.warning || "",
      beanRef: beanRef,
    });
  }
  return { ok: true, processed: processed };
}

// Apply inventory writes for validated roast batches. Called after validation passes.
function applyRoastBatchInventory(processed, refType, refId, person) {
  for (var i = 0; i < processed.length; i++) {
    var p = processed[i];
    if (p.greenBeanItemId && p.inputQty > 0) {
      var outNotes = "Roast batch " + (i + 1) + ": " + p.inputQty + "kg from " + p.sourceAutoId;
      if (p.greenBeanWarning) outNotes = p.greenBeanWarning + " | " + outNotes;
      createInventoryTransaction(p.greenBeanItemId, "OUT", p.inputQty, refType, refId, outNotes, person, "rb_" + i + "_out");
    }
    if (p.roastedBeanItemId && p.outputQty > 0) {
      var inNotes = "Roast batch " + (i + 1) + ": " + p.outputQty + "kg output" + (p.classificationId ? " [" + lookupClassificationLabel(p.classificationId) + "]" : "");
      if (p.roastedBeanWarning) inNotes = p.roastedBeanWarning + " | " + inNotes;
      createInventoryTransaction(p.roastedBeanItemId, "IN", p.outputQty, refType, refId, inNotes, person, "rb_" + i + "_in", p.classificationId);
    }
  }
}

// Write QuantityAllocation rows to track green bean usage by roast batches.
function applyRoastBatchAllocations(processed, destAutoId, destType, destId, allocatedBy) {
  var greenBeanCk = lookupChecklist("ck_green_beans");
  if (!greenBeanCk) return;
  for (var i = 0; i < processed.length; i++) {
    var p = processed[i];
    var srcInfo = getSubmissionByAutoId(p.sourceAutoId, greenBeanCk);
    var srcTotal = srcInfo ? srcInfo.totalQuantity : 0;
    createQuantityAllocation(greenBeanCk.id, p.sourceAutoId, srcTotal, destType, destId, destAutoId, p.inputQty, allocatedBy);
  }
}

// Build auto-id for multi-batch roast: RB-DDMMYY-[ABBR|MIX]-NNN
function buildRoastAutoId(processed, dateStr) {
  var prefix = "RB";
  var dateFormatted = formatDateDDMMYY(dateStr || new Date());
  var itemCode = "X";
  if (processed.length > 0) {
    var firstRef = processed[0].beanRef;
    var ab = getInventoryAbbreviation(firstRef);
    itemCode = ab || sanitizeItemCodeToken(firstRef) || "X";
    if (processed.length > 1) {
      var allSame = processed.every(function(p) { return p.beanRef === firstRef; });
      if (!allSame) itemCode = itemCode + "MIX";
    }
  }
  var seq = getNextSequenceForPrefix(prefix);
  return prefix + "-" + dateFormatted + "-" + itemCode + "-" + String(seq).padStart(3, "0");
}

// ─── Submit Checklist & Responses ──────────────────────────────

function handleSubmitChecklist(body, user) {
  ensureSheetHasAllColumns(SHEETS.ORDER_CHECKLISTS);
  var ocId = body.id;
  var date = body.date || "";
  var person = body.person || "";
  var responses = body.responses || [];
  var now = new Date().toISOString();

  var ocIdx = findRowIndex(SHEETS.ORDER_CHECKLISTS, ocId);
  if (ocIdx < 0) return { error: "Order checklist not found" };
  var ocRows = getRows(SHEETS.ORDER_CHECKLISTS);
  var oc = null;
  for (var i = 0; i < ocRows.length; i++) { if (String(ocRows[i].id) === String(ocId)) { oc = ocRows[i]; break; } }
  if (!oc) return { error: "Order checklist not found" };

  // Look up related data
  var ck = lookupChecklist(oc.checklist_id);
  if (!ck) return { error: "Checklist template not found" };
  var order = lookupOrder(oc.order_id);
  var customerLabel = order ? lookupCustomerLabel(order.customer_id) : "";
  var orderTypeLabel = order ? lookupOrderTypeLabel(order.order_type) : "";

  // Build responses map for downstream helpers
  var responsesMap = responsesArrayToMap(responses);
  var batchAllocations = body.batchAllocations || null;

  // Validate quantity allocations BEFORE mutating anything (read-only pass; excludes prior allocations from this OC)
  var preAlloc = processQuantityAllocationsForSubmission(ck, responsesMap, batchAllocations, "", "checklist", ocId, person, oc.auto_id || "", false);
  if (!preAlloc.ok) return { error: preAlloc.error };

  // Reverse any prior allocations from a previous submission of the same OC (re-submit case)
  if (oc.auto_id) reverseAllocationsForDestination(oc.auto_id);

  // Generate auto ID (if checklist has it enabled)
  var autoId = generateAutoId(ck, responsesMap, date || now);

  // Update OC status + auto_id
  oc.status = "completed"; oc.completed_at = now; oc.completed_by = person; oc.work_date = date;
  if (autoId) oc.auto_id = autoId;
  updateSheetRow(SHEETS.ORDER_CHECKLISTS, ocIdx, oc);

  // Delete old responses if re-submitting
  deleteResponseRow(ck.name, String(oc.order_id));
  deleteFromMasterSummary(String(oc.order_id), ck.name);

  // Build response array for sheet write
  var respArray = responses.map(function(r) { return (r.response !== undefined && r.response !== null) ? String(r.response) : ""; });
  var remarks = body.remarks || {};

  // ── Multi-batch roasting path (ck_roasted_beans with roast_batches array) ──
  var roastBatches = body.roast_batches || null;
  if (ck.id === "ck_roasted_beans" && Array.isArray(roastBatches) && roastBatches.length > 0) {
    var rbValidation = validateRoastBatches(roastBatches, oc.auto_id || "");
    if (!rbValidation.ok) return { error: rbValidation.error };
    autoId = buildRoastAutoId(rbValidation.processed, date || now);
    oc.auto_id = autoId;
    updateSheetRow(SHEETS.ORDER_CHECKLISTS, ocIdx, oc);
    // Store roast_batches JSON in "Shipment number used" column (index 0) BEFORE writing
    if (respArray.length > 0) respArray[0] = JSON.stringify(roastBatches);
    // Fill summary quantities into template fields so they are persisted and readable
    var totalIn = 0, totalOut = 0;
    for (var ti = 0; ti < rbValidation.processed.length; ti++) { totalIn += rbValidation.processed[ti].inputQty; totalOut += rbValidation.processed[ti].outputQty; }
    if (respArray.length > 3) respArray[3] = String(totalIn);   // Quantity input
    if (respArray.length > 4) respArray[4] = String(totalOut);  // Quantity output
    if (respArray.length > 6) respArray[6] = String(Math.round((totalIn - totalOut) * 100) / 100); // Loss in weight
    // Write response row with batch data embedded
    writeResponseRow(ck.name, ck.questions, {
      orderId: String(oc.order_id), orderName: order ? order.name : "", customer: customerLabel,
      person: person, date: date, submittedAt: now, responses: respArray, remarks: remarks,
    });
    addToMasterSummary({
      orderId: String(oc.order_id), orderName: order ? order.name : "", customer: customerLabel,
      orderType: orderTypeLabel, checklistName: ck.name, person: person, date: date, submittedAt: now,
    });
    applyRoastBatchInventory(rbValidation.processed, "checklist", ocId, person);
    if (autoId) applyRoastBatchAllocations(rbValidation.processed, autoId, "checklist", ocId, person);
    for (var bi = 0; bi < rbValidation.processed.length; bi++) {
      var bp = rbValidation.processed[bi];
      writeAuditLog(user, "tag", "roast_batch_link", ocId,
        "Batch " + (bi + 1) + ": Input " + bp.inputQty + "kg from " + bp.sourceAutoId + " → Output " + bp.outputQty + "kg" + (bp.classificationId ? " [" + lookupClassificationLabel(bp.classificationId) + "]" : ""));
    }
    writeAuditLog(user, "submit", "Checklist", ocId, ck.name + " [" + autoId + "] (" + rbValidation.processed.length + " batches)");
    return { success: true, autoId: autoId };
  }

  // Standard path — write responses and master summary
  writeResponseRow(ck.name, ck.questions, {
    orderId: String(oc.order_id), orderName: order ? order.name : "", customer: customerLabel,
    person: person, date: date, submittedAt: now, responses: respArray, remarks: remarks,
  });
  addToMasterSummary({
    orderId: String(oc.order_id), orderName: order ? order.name : "", customer: customerLabel,
    orderType: orderTypeLabel, checklistName: ck.name, person: person, date: date, submittedAt: now,
  });

  // ── Inventory processing: new inventoryLink system takes precedence over legacy ──
  // If any question on this checklist has inventoryLink.enabled, we ONLY run the new
  // processInventoryLinks() and skip the legacy Input Item / Output Item dropdown path —
  // running both caused every quantity to be written twice into the ledger.
  var nqForInvCheck = (ck && ck.questions) || [];
  var hasInventoryLinkQuestions = nqForInvCheck.some(function(q) { return q.inventoryLink && q.inventoryLink.enabled; });

  if (hasInventoryLinkQuestions) {
    processInventoryLinks(ck, responsesMap, "checklist", ocId, person, false);
  }

  // ── Quantity allocations (Feature 3) — commit batch rows now that pre-validation passed ──
  if (autoId) processQuantityAllocationsForSubmission(ck, responsesMap, batchAllocations, autoId, "checklist", ocId, person, "", true);

  // ── Legacy hard-coded inventory tracking (only for checklists NOT migrated to inventoryLink) ──
  if (!hasInventoryLinkQuestions) {
    var respMap = {};
    for (var ri = 0; ri < responses.length; ri++) {
      respMap[responses[ri].questionText] = responses[ri].response || "";
    }
    applyLegacyInventoryForChecklist(ck, respMap, "checklist", ocId, person, body.inventoryItemId || "", body.inventoryOutputItemId || "", false, body.grindClassificationId || "");
  }

  writeAuditLog(user, "submit", "Checklist", ocId, ck.name + (autoId ? " [" + autoId + "]" : ""));
  return { success: true, autoId: autoId };
}

function handleEditResponse(body, user) {
  var ocId = body.id;
  var newResponses = body.responses || [];
  var newDate = body.date;
  var newPerson = body.person;

  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  var oc = null;
  for (var j = 0; j < ocs.length; j++) { if (String(ocs[j].id) === String(ocId)) { oc = ocs[j]; break; } }
  if (!oc) return { error: "Order checklist not found" };

  var ck = lookupChecklist(oc.checklist_id);
  if (!ck) return { error: "Checklist template not found" };

  // ── Phase 5: Validate downstream quantity constraints before saving ──
  if (oc.auto_id) {
    var editResponsesMapForValidation = responsesArrayToMap(newResponses);
    var newMasterQty = getMasterQuantityFromResponses(ck.questions, editResponsesMapForValidation);
    if (newMasterQty > 0) {
      var qvResult = validateQuantityEdit(oc.checklist_id, oc.auto_id, newMasterQty);
      if (!qvResult.allowed) return { error: qvResult.reason };
    }
  }

  if (newPerson) oc.completed_by = newPerson;
  if (newDate) oc.work_date = newDate;
  var ocIdx = findRowIndex(SHEETS.ORDER_CHECKLISTS, ocId);
  if (ocIdx > 0) updateSheetRow(SHEETS.ORDER_CHECKLISTS, ocIdx, oc);
  var order = lookupOrder(oc.order_id);
  var customerLabel = order ? lookupCustomerLabel(order.customer_id) : "";
  var orderTypeLabel = order ? lookupOrderTypeLabel(order.order_type) : "";

  // Delete old and write new
  deleteResponseRow(ck.name, String(oc.order_id));
  deleteFromMasterSummary(String(oc.order_id), ck.name);

  var respArray = newResponses.map(function(r) { return (r.response !== undefined && r.response !== null) ? String(r.response) : ""; });
  var remarks = body.remarks || {};
  writeResponseRow(ck.name, ck.questions, {
    orderId: String(oc.order_id), orderName: order ? order.name : "", customer: customerLabel,
    person: newPerson || oc.completed_by, date: newDate || oc.work_date,
    submittedAt: oc.completed_at, responses: respArray, remarks: remarks,
  });

  addToMasterSummary({
    orderId: String(oc.order_id), orderName: order ? order.name : "", customer: customerLabel,
    orderType: orderTypeLabel, checklistName: ck.name,
    person: newPerson || oc.completed_by, date: newDate || oc.work_date, submittedAt: oc.completed_at,
  });

  // ── Multi-batch roasting edit path ──
  var editRoastBatches = body.roast_batches || null;
  var editPerson = newPerson || oc.completed_by;
  if (ck.id === "ck_roasted_beans" && Array.isArray(editRoastBatches) && editRoastBatches.length > 0) {
    // Capture before state from existing response sheet before reversal
    var beforeBatches = "";
    try {
      var oldResp = readResponseRow(ck.name, ck.questions, String(oc.order_id));
      if (oldResp && oldResp.responses && oldResp.responses[0]) beforeBatches = oldResp.responses[0].response || "";
    } catch(e) {}
    // Reverse all prior inventory + allocations
    reverseInventoryLedgerForRef("checklist", ocId, editPerson);
    if (oc.auto_id) reverseAllocationsForDestination(oc.auto_id);
    // Validate and apply new batches
    var editRbVal = validateRoastBatches(editRoastBatches, oc.auto_id || "");
    if (!editRbVal.ok) return { error: editRbVal.error };
    applyRoastBatchInventory(editRbVal.processed, "checklist", ocId, editPerson);
    if (oc.auto_id) applyRoastBatchAllocations(editRbVal.processed, oc.auto_id, "checklist", ocId, editPerson);
    writeAuditLog(user, "edit", "ChecklistResponse", ocId,
      "Roast batches edited | before=" + beforeBatches + " | after=" + JSON.stringify(editRoastBatches));
    var refWarning = "";
    try { if (oc.auto_id) { var refs = findUpstreamReferencesForAutoId(oc.auto_id); if (refs.length > 0) refWarning = "Referenced by: " + refs.join(", "); } } catch(e) {}
    var result = { success: true };
    if (refWarning) result.warning = refWarning;
    return result;
  }

  // ── Inventory re-processing for edit ─────────────────────────
  // Reverse prior ledger entries for this OC, then re-apply based on the updated responses.
  // This prevents accumulation of inventory entries on every edit.
  var editResponsesMap = responsesArrayToMap(newResponses);
  var nqForInvEdit = (ck && ck.questions) || [];
  var hasInventoryLinkQuestionsEdit = nqForInvEdit.some(function(q) { return q.inventoryLink && q.inventoryLink.enabled; });

  if (hasInventoryLinkQuestionsEdit) {
    reverseInventoryLedgerForRef("checklist", ocId, editPerson);
    processInventoryLinks(ck, editResponsesMap, "checklist", ocId, editPerson, true);
  } else {
    // Legacy hard-coded path: reverse prior entries, then re-apply using updated values.
    var reversedCount = reverseInventoryLedgerForRef("checklist", ocId, editPerson);
    var respMapE = {};
    for (var riE = 0; riE < newResponses.length; riE++) {
      respMapE[newResponses[riE].questionText] = newResponses[riE].response || "";
    }
    applyLegacyInventoryForChecklist(ck, respMapE, "checklist", ocId, editPerson, body.inventoryItemId || "", body.inventoryOutputItemId || "", true, body.grindClassificationId || "");
    Logger.log("handleEditResponse: legacy inventory — reversed " + reversedCount + " prior entries");
  }

  writeAuditLog(user, "edit_response", "ChecklistResponse", ocId, "Edited responses");
  // Reference check: if this response's autoId is referenced downstream, return a warning (non-blocking)
  var refWarning = "";
  try {
    if (oc.auto_id) {
      var refs = findUpstreamReferencesForAutoId(oc.auto_id);
      if (refs.length > 0) refWarning = "Referenced by: " + refs.join(", ");
    }
  } catch (e) { /* non-fatal */ }
  var result = { success: true };
  if (refWarning) result.warning = refWarning;
  return result;
}

function handleRevertChecklist(body, user) {
  var ocId = body.id;
  var ocIdx = findRowIndex(SHEETS.ORDER_CHECKLISTS, ocId);
  if (ocIdx < 0) return { error: "Order checklist not found" };
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  var oc = null;
  for (var i = 0; i < ocs.length; i++) { if (String(ocs[i].id) === String(ocId)) { oc = ocs[i]; break; } }
  if (!oc) return { error: "Order checklist not found" };

  // Reference check: is this checklist's autoId referenced by any other checklist or order stage?
  if (oc.auto_id) {
    var refs = findUpstreamReferencesForAutoId(oc.auto_id);
    if (refs.length > 0) {
      writeAuditLog(user, "revert_blocked", "Checklist", ocId, "Blocked: referenced by " + refs.join(", "));
      return { error: "Cannot delete — " + refs.join(", ") + " is linked to this. Remove that link first." };
    }
  }

  // Delete responses from per-checklist tab + Master Summary
  var ck = lookupChecklist(oc.checklist_id);
  if (ck) {
    deleteResponseRow(ck.name, String(oc.order_id));
    deleteFromMasterSummary(String(oc.order_id), ck.name);
  }

  oc.status = "pending"; oc.completed_at = ""; oc.completed_by = ""; oc.work_date = "";
  updateSheetRow(SHEETS.ORDER_CHECKLISTS, ocIdx, oc);
  writeAuditLog(user, "revert", "Checklist", ocId, "Reverted to pending");
  return { success: true };
}

function handleGetResponses(params) {
  var ocId = params.id;
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  var oc = null;
  for (var i = 0; i < ocs.length; i++) { if (String(ocs[i].id) === String(ocId)) { oc = ocs[i]; break; } }
  if (!oc) return { error: "Order checklist not found" };

  var ck = lookupChecklist(oc.checklist_id);
  if (!ck) return { workDate: oc.work_date || "", person: oc.completed_by || "", completedAt: oc.completed_at || "", responses: [] };

  var responseData = readResponseRow(ck.name, ck.questions, String(oc.order_id));
  if (!responseData) {
    return { workDate: oc.work_date || "", person: oc.completed_by || "", completedAt: oc.completed_at || "", responses: [] };
  }

  return {
    workDate: responseData.date || oc.work_date || "",
    person: responseData.person || oc.completed_by || "",
    completedAt: oc.completed_at || "",
    autoId: oc.auto_id || "",
    responses: responseData.responses.map(function(r, idx) {
      var resp = r.response;
      // For inventory_item fields, the sheet stores display name; convert back to id so the dropdown can pre-select.
      // Old rows may already have an id ("inv_…") — leave those as-is.
      if (ck.questions[idx] && ck.questions[idx].type === "inventory_item" && resp) {
        if (String(resp).indexOf("inv_") !== 0) resp = inventoryItemIdByName(resp);
      }
      return { questionIndex: idx, questionText: r.questionText, response: resp, remark: r.remark || "", respondedBy: responseData.person, respondedAt: responseData.submittedAt };
    }),
  };
}

function handleGetAllResponses() {
  var checklists = handleGetChecklists();
  var allOcs = getRows(SHEETS.ORDER_CHECKLISTS);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = [];

  for (var c = 0; c < checklists.length; c++) {
    var ck = checklists[c];
    var sheet = ss.getSheetByName(ck.name);
    if (!sheet || sheet.getLastRow() < 3) continue;

    var data = sheet.getDataRange().getValues();
    // Build header-based column map from row 2
    var hdrs = data.length >= 2 ? data[1].map(String) : [];
    var nq = ck.questions; // Already normalized by handleGetChecklists
    var qTxts = questionTexts(nq);
    var remarkIdx = getRemarkIndices(nq);
    var rmkHdrs = remarkIdx.map(function(ri2) { return "Remarks: " + nq[ri2].text; });
    var qcMap = {}; for (var qi2 = 0; qi2 < qTxts.length; qi2++) qcMap[qi2] = hdrs.indexOf(qTxts[qi2]);
    var rcMap = {}; for (var rj = 0; rj < rmkHdrs.length; rj++) rcMap[rj] = hdrs.indexOf(rmkHdrs[rj]);

    for (var i = 2; i < data.length; i++) {
      var orderId = String(data[i][0] || "");
      if (!orderId) continue;

      // Find corresponding orderChecklistId
      var ocId = "";
      for (var j = 0; j < allOcs.length; j++) {
        if (String(allOcs[j].order_id) === orderId && String(allOcs[j].checklist_id) === ck.id) {
          ocId = String(allOcs[j].id); break;
        }
      }

      var responses = [];
      for (var q = 0; q < nq.length; q++) {
        var rawVal = qcMap[q] >= 0 ? String(data[i][qcMap[q]] || "") : "";
        if (nq[q] && nq[q].type === "inventory_item" && rawVal && rawVal.indexOf("inv_") === 0) {
          rawVal = inventoryItemNameById(rawVal);
        }
        responses.push({
          question: nq[q].text,
          response: rawVal,
          remark: "",
          originalResponse: null,
        });
      }
      // Read remark columns by header name
      for (var ri = 0; ri < remarkIdx.length; ri++) {
        if (rcMap[ri] >= 0) responses[remarkIdx[ri]].remark = String(data[i][rcMap[ri]] || "");
      }

      result.push({
        orderChecklistId: ocId,
        orderName: String(data[i][1] || ""),
        customer: String(data[i][2] || ""),
        checklistName: ck.name,
        person: String(data[i][3] || ""),
        date: String(data[i][4] || ""),
        submittedAt: String(data[i][5] || ""),
        editedAt: null, editedBy: null,
        responses: responses,
      });
    }
  }

  result.sort(function(a, b) { return String(b.submittedAt).localeCompare(String(a.submittedAt)); });
  return result;
}

function handleGetAuditLog(params) {
  params = params || {};
  var rows = getRows(SHEETS.AUDIT_LOG);
  // Apply optional filters
  var filtered = rows.filter(function(r) {
    if (params.action && String(r.action) !== String(params.action)) return false;
    if (params.entityType && String(r.entity_type) !== String(params.entityType)) return false;
    if (params.entityId && String(r.entity_id).indexOf(String(params.entityId)) < 0) return false;
    if (params.performedBy && String(r.user_name || "").toLowerCase().indexOf(String(params.performedBy).toLowerCase()) < 0) return false;
    if (params.dateFrom) {
      var ts = String(r.timestamp || "").split("T")[0];
      if (ts < String(params.dateFrom)) return false;
    }
    if (params.dateTo) {
      var ts2 = String(r.timestamp || "").split("T")[0];
      if (ts2 > String(params.dateTo)) return false;
    }
    return true;
  });
  filtered.sort(function(a, b) { return String(b.timestamp).localeCompare(String(a.timestamp)); });
  var offset = parseInt(params.offset) || 0;
  var limit = Math.min(parseInt(params.limit) || 50, 100);
  var page = filtered.slice(offset, offset + limit);
  return {
    entries: page.map(function(r) {
      return {
        id: r.id, timestamp: r.timestamp, action: r.action,
        entityType: r.entity_type, entityId: r.entity_id,
        performedBy: r.user_name || "", details: r.details || "",
      };
    }),
    total: filtered.length,
    hasMore: (offset + limit) < filtered.length,
  };
}

// ─── Archiving ─────────────────────────────────────────────────

function handleArchiveOrders(body, user) {
  var daysOld = body.daysOld || 30;
  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - daysOld);
  var cutoffStr = cutoff.toISOString();

  var orders = getRows(SHEETS.ORDERS);
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);

  var toArchive = orders.filter(function(o) {
    var orderOcs = ocs.filter(function(c) { return String(c.order_id) === String(o.id); });
    var allCompleted = orderOcs.length > 0 && orderOcs.every(function(c) { return c.status === "completed"; });
    return allCompleted && String(o.created_at) < cutoffStr;
  });

  if (toArchive.length === 0) return { error: "No completed orders older than " + daysOld + " days found" };

  var archiveId = "arch_" + nextId();
  var now = new Date().toISOString();
  var orderIds = toArchive.map(function(o) { return String(o.id); });

  // Copy orders to archive
  for (var i = 0; i < toArchive.length; i++) {
    toArchive[i].archive_id = archiveId;
    appendToSheet(SHEETS.ARCHIVED_ORDERS, toArchive[i]);
  }

  // Copy OCs to archive + collect response data
  var archivedOcIds = [];
  var checklists = handleGetChecklists();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  for (var j = 0; j < ocs.length; j++) {
    if (orderIds.indexOf(String(ocs[j].order_id)) >= 0) {
      ocs[j].archive_id = archiveId;
      appendToSheet(SHEETS.ARCHIVED_ORDER_CHECKLISTS, ocs[j]);
      archivedOcIds.push(String(ocs[j].id));

      // Archive response data from per-checklist tab
      var ck = null;
      for (var c = 0; c < checklists.length; c++) {
        if (String(checklists[c].id) === String(ocs[j].checklist_id)) { ck = checklists[c]; break; }
      }
      if (ck) {
        var respData = readResponseRow(ck.name, ck.questions, String(ocs[j].order_id));
        if (respData) {
          for (var q = 0; q < respData.responses.length; q++) {
            appendToSheet(SHEETS.ARCHIVED_RESPONSES, {
              archive_id: archiveId, order_id: String(ocs[j].order_id),
              order_name: "", customer: "", checklist_name: ck.name,
              person: respData.person, date: respData.date,
              question: respData.responses[q].questionText,
              response: respData.responses[q].response,
              remark: respData.responses[q].remark || "",
              submitted_at: respData.submittedAt,
            });
          }
        }
      }
    }
  }

  // Write archive metadata
  var dates = toArchive.map(function(o) { return String(o.created_at); }).sort();
  appendToSheet(SHEETS.ARCHIVES_META, {
    id: archiveId, date_range_start: dates[0] || "", date_range_end: dates[dates.length - 1] || "",
    orders_count: toArchive.length, created_at: now, created_by: user.displayName, order_ids: JSON.stringify(orderIds),
  });

  // Delete from active: per-checklist tabs + Master Summary
  for (var d = 0; d < toArchive.length; d++) {
    var ordId = String(toArchive[d].id);
    for (var e = 0; e < checklists.length; e++) {
      deleteResponseRow(checklists[e].name, ordId);
    }
    deleteFromMasterSummary(ordId, null);
  }

  // Delete OCs
  for (var m = ocs.length - 1; m >= 0; m--) {
    if (orderIds.indexOf(String(ocs[m].order_id)) >= 0) {
      var ocIdx = findRowIndex(SHEETS.ORDER_CHECKLISTS, ocs[m].id);
      if (ocIdx > 0) deleteSheetRow(SHEETS.ORDER_CHECKLISTS, ocIdx);
    }
  }

  // Delete Orders
  for (var n = toArchive.length - 1; n >= 0; n--) {
    var oIdx = findRowIndex(SHEETS.ORDERS, toArchive[n].id);
    if (oIdx > 0) deleteSheetRow(SHEETS.ORDERS, oIdx);
  }

  writeAuditLog(user, "archive", "Archive", archiveId, "Archived " + toArchive.length + " orders older than " + daysOld + " days");
  return { success: true, archived: toArchive.length, archiveId: archiveId };
}

function handleGetArchives() {
  var rows = getRows(SHEETS.ARCHIVES_META);
  rows.sort(function(a, b) { return String(b.created_at).localeCompare(String(a.created_at)); });
  return rows.map(function(r) {
    return { id: r.id, dateRangeStart: r.date_range_start, dateRangeEnd: r.date_range_end, ordersCount: r.orders_count, createdAt: r.created_at, createdBy: r.created_by };
  });
}

// ─── Untagged Checklists ──────────────────────────────────────

function isDeleted(row) {
  return row.is_deleted === true || row.is_deleted === "true";
}

function handleGetUntagged() {
  ensureSheetHasAllColumns(SHEETS.UNTAGGED_CHECKLISTS);
  // Show a checklist in the dashboard while its output quantity has NOT been fully consumed.
  // Consumption = max(tagged_quantity, total allocations in QuantityAllocations).
  return getRows(SHEETS.UNTAGGED_CHECKLISTS).filter(function(r) { return !isDeleted(r); }).map(function(r) {
    var totalQ = parseFloat(r.total_quantity) || 0;
    var taggedQ = parseFloat(r.tagged_quantity) || 0;
    var allocatedFromQA = r.auto_id ? getAllocatedQuantityForAutoId(r.auto_id) : 0;
    var effectiveTagged = Math.max(taggedQ, allocatedFromQA);
    return {
      id: r.id, checklistId: r.checklist_id, checklistName: r.checklist_name,
      person: r.person, date: r.date, submittedAt: r.submitted_at,
      taggedOrderId: r.tagged_order_id || "", responses: safeParseJSON(r.responses, []),
      remarks: safeParseJSON(r.remarks, {}), submittedByUserId: r.submitted_by_user_id || "",
      totalQuantity: totalQ, taggedQuantity: effectiveTagged, remainingQuantity: totalQ - effectiveTagged,
      allocations: safeParseJSON(r.allocations, []),
      autoId: r.auto_id || "",
    };
  }).filter(function(entry) {
    // Visible while remaining > 0, OR if no tracking at all and untagged (legacy rows without total_quantity)
    if (entry.totalQuantity > 0) return entry.remainingQuantity > 0;
    // No total quantity configured — fall back to classic untagged-or-partial behavior
    var isUntagged = !entry.taggedOrderId || String(entry.taggedOrderId).trim() === "";
    return isUntagged;
  });
}

// Fetch a single untagged checklist row by its id. Used by the edit dialog on the
// dashboard so it can pre-populate fields with the saved responses & remarks.
// Returns the same shape as a row in handleGetUntagged() plus the parsed batchAllocations
// when present, so the edit form can re-show batch tags.
function handleGetUntaggedResponse(params) {
  var id = (params && (params.id || params.responseId)) || "";
  if (!id) return { error: "Missing untagged response id" };
  ensureSheetHasAllColumns(SHEETS.UNTAGGED_CHECKLISTS);
  var rows = getRows(SHEETS.UNTAGGED_CHECKLISTS);
  var r = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { r = rows[i]; break; } }
  if (!r) return { error: "Untagged response not found: " + id };
  if (isDeleted(r)) return { error: "Untagged response has been deleted: " + id };
  var totalQ = parseFloat(r.total_quantity) || 0;
  var taggedQ = parseFloat(r.tagged_quantity) || 0;
  var allocatedFromQA = r.auto_id ? getAllocatedQuantityForAutoId(r.auto_id) : 0;
  var effectiveTagged = Math.max(taggedQ, allocatedFromQA);
  return {
    id: r.id, checklistId: r.checklist_id, checklistName: r.checklist_name,
    person: r.person, date: r.date, submittedAt: r.submitted_at,
    taggedOrderId: r.tagged_order_id || "",
    responses: safeParseJSON(r.responses, []),
    remarks: safeParseJSON(r.remarks, {}),
    submittedByUserId: r.submitted_by_user_id || "",
    totalQuantity: totalQ, taggedQuantity: effectiveTagged, remainingQuantity: totalQ - effectiveTagged,
    allocations: safeParseJSON(r.allocations, []),
    autoId: r.auto_id || "",
  };
}

function handleSubmitUntagged(body, user) {
  ensureSheetHasAllColumns(SHEETS.UNTAGGED_CHECKLISTS);
  var checklistId = body.checklistId;
  var ck = lookupChecklist(checklistId);
  if (!ck) return { error: "Checklist template not found" };

  var id = "ut_" + nextId();
  var now = new Date().toISOString();
  var person = body.person || "";
  var date = body.date || "";
  var responses = body.responses || [];
  var remarks = body.remarks || {};
  var orderId = body.orderId || ""; // optional tag-to-order
  var responsesMap = responsesArrayToMap(responses);
  var batchAllocations = body.batchAllocations || null;

  // Validate quantity allocations BEFORE mutating anything
  var preAlloc = processQuantityAllocationsForSubmission(ck, responsesMap, batchAllocations, "", "checklist", id, person, "", false);
  if (!preAlloc.ok) return { error: preAlloc.error };

  // Generate auto ID
  var autoId = generateAutoId(ck, responsesMap, date || now);

  // ── Multi-batch roasting path for ck_roasted_beans ──
  var utRoastBatches = body.roast_batches || null;
  var isMultiBatchRoast = checklistId === "ck_roasted_beans" && Array.isArray(utRoastBatches) && utRoastBatches.length > 0;
  if (isMultiBatchRoast) {
    var utRbVal = validateRoastBatches(utRoastBatches, "");
    if (!utRbVal.ok) return { error: utRbVal.error };
    autoId = buildRoastAutoId(utRbVal.processed, date || now);
  }

  // Extract total quantity
  var totalQuantity = 0;
  if (isMultiBatchRoast) {
    for (var tqi = 0; tqi < utRoastBatches.length; tqi++) totalQuantity += parseFloat(utRoastBatches[tqi].outputQty) || 0;
  } else {
    totalQuantity = getMasterQuantityFromResponses(ck.questions, responsesMap);
    if (totalQuantity <= 0) {
      var nq = ck.questions;
      for (var qi = 0; qi < nq.length; qi++) {
        if ((nq[qi].type === "number" || nq[qi].type === "text_number") &&
            (nq[qi].text.toLowerCase().indexOf("quantity") >= 0 || nq[qi].text.toLowerCase().indexOf("weight") >= 0 || nq[qi].text.toLowerCase().indexOf("net") >= 0)) {
          var qtyVal = parseFloat(responsesMap[qi]) || 0;
          if (qtyVal > totalQuantity) totalQuantity = qtyVal;
        }
      }
    }
  }

  var obj = {
    id: id, checklist_id: checklistId, checklist_name: ck.name,
    person: person, date: date, submitted_at: now,
    tagged_order_id: orderId, responses: JSON.stringify(responses),
    remarks: JSON.stringify(remarks), submitted_by_user_id: user.id,
    total_quantity: totalQuantity, tagged_quantity: orderId ? totalQuantity : 0,
    allocations: orderId ? JSON.stringify([{orderId: orderId, quantity: totalQuantity}]) : "[]",
    auto_id: autoId || "",
  };
  appendToSheet(SHEETS.UNTAGGED_CHECKLISTS, obj);

  // Build respArray for per-checklist response tab
  var respArray = responses.map(function(r) { return (r.response !== undefined && r.response !== null) ? String(r.response) : ""; });

  if (isMultiBatchRoast) {
    // Store batch JSON in Shipment column + fill summary fields
    if (respArray.length > 0) respArray[0] = JSON.stringify(utRoastBatches);
    var utTotalIn = 0, utTotalOut = 0;
    for (var uti = 0; uti < utRbVal.processed.length; uti++) { utTotalIn += utRbVal.processed[uti].inputQty; utTotalOut += utRbVal.processed[uti].outputQty; }
    if (respArray.length > 3) respArray[3] = String(utTotalIn);
    if (respArray.length > 4) respArray[4] = String(utTotalOut);
    if (respArray.length > 6) respArray[6] = String(Math.round((utTotalIn - utTotalOut) * 100) / 100);
    // Inventory per-batch (skip legacy processInventoryLinks)
    applyRoastBatchInventory(utRbVal.processed, "untagged", id, person);
    if (autoId) applyRoastBatchAllocations(utRbVal.processed, autoId, "checklist", id, person);
    for (var utbi = 0; utbi < utRbVal.processed.length; utbi++) {
      var utbp = utRbVal.processed[utbi];
      writeAuditLog(user, "tag", "roast_batch_link", id,
        "Batch " + (utbi + 1) + ": Input " + utbp.inputQty + "kg from " + utbp.sourceAutoId + " → Output " + utbp.outputQty + "kg");
    }
  } else {
    // Standard inventory processing
    var nqForInvCheckUt = (ck && ck.questions) || [];
    var hasInventoryLinkQuestionsUt = nqForInvCheckUt.some(function(q) { return q.inventoryLink && q.inventoryLink.enabled; });
    if (hasInventoryLinkQuestionsUt) {
      processInventoryLinks(ck, responsesMap, "untagged", id, person, false);
    }
    if (autoId) processQuantityAllocationsForSubmission(ck, responsesMap, batchAllocations, autoId, "checklist", id, person, "", true);
  }

  var orderName = "", customerLabel = "", orderTypeLabel = "";
  if (orderId) {
    var order = lookupOrder(orderId);
    if (order) {
      orderName = order.name;
      customerLabel = lookupCustomerLabel(order.customer_id);
      orderTypeLabel = lookupOrderTypeLabel(order.order_type);
    }
  }

  writeResponseRow(ck.name, ck.questions, {
    orderId: orderId || "UNTAGGED", orderName: orderName, customer: customerLabel,
    person: person, date: date, submittedAt: now, responses: respArray, remarks: remarks,
  });

  // If tagged to an order, also add to Master Summary and create OrderChecklist
  if (orderId) {
    // Create an OrderChecklist entry so it shows in the order
    var ocId = "oc_" + nextId() + "_ut";
    appendToSheet(SHEETS.ORDER_CHECKLISTS, {
      id: ocId, order_id: orderId, checklist_id: checklistId,
      status: "completed", completed_at: now, completed_by: person, work_date: date,
      auto_id: autoId || "",
    });
    addToMasterSummary({
      orderId: orderId, orderName: orderName, customer: customerLabel,
      orderType: orderTypeLabel, checklistName: ck.name, person: person, date: date, submittedAt: now,
    });
    // Mark as tagged immediately
    var utIdx = findRowIndex(SHEETS.UNTAGGED_CHECKLISTS, id);
    if (utIdx > 0) {
      obj.tagged_order_id = orderId;
      updateSheetRow(SHEETS.UNTAGGED_CHECKLISTS, utIdx, obj);
    }
  }

  // ── Inventory transactions for untagged submissions (legacy path; only when no inventoryLink) ──
  if (!hasInventoryLinkQuestionsUt) {
    var respMap2 = {};
    for (var ri2 = 0; ri2 < responses.length; ri2++) {
      respMap2[responses[ri2].questionText] = responses[ri2].response || "";
    }
    applyLegacyInventoryForChecklist(ck, respMap2, "checklist", id, person, body.inventoryItemId || "", body.inventoryOutputItemId || "", false, body.grindClassificationId || "");
  }

  writeAuditLog(user, "submit_untagged", "UntaggedChecklist", id, ck.name + (autoId ? " [" + autoId + "]" : "") + (orderId ? " (tagged to " + orderId + ")" : ""));
  return {
    id: id, checklistId: checklistId, checklistName: ck.name,
    person: person, date: date, submittedAt: now, taggedOrderId: orderId,
    responses: responses, remarks: remarks, submittedByUserId: user.id,
    autoId: autoId,
  };
}

function handleTagUntagged(body, user) {
  var utId = body.id;
  var orderId = body.orderId;
  var tagQuantity = parseFloat(body.tagQuantity) || 0;
  if (!utId || !orderId) return { error: "Missing id or orderId" };

  var rows = getRows(SHEETS.UNTAGGED_CHECKLISTS);
  var ut = null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(utId)) { ut = rows[i]; break; }
  }
  if (!ut) return { error: "Untagged checklist not found" };
  if (isDeleted(ut)) return { error: "This entry has been deleted and cannot be tagged." };

  // Permission check: admin or own submission
  if (user.role !== "admin" && String(ut.submitted_by_user_id) !== String(user.id)) {
    return { error: "You can only tag your own untagged checklists" };
  }

  var totalQ = parseFloat(ut.total_quantity) || 0;
  var currentTaggedQ = parseFloat(ut.tagged_quantity) || 0;

  // If no quantity tracking (totalQ=0) or no partial quantity specified, do full tag
  if (totalQ <= 0 || tagQuantity <= 0) tagQuantity = totalQ > 0 ? (totalQ - currentTaggedQ) : 0;

  // Check if already fully tagged
  if (totalQ > 0 && currentTaggedQ >= totalQ) {
    return { error: "This checklist is already fully tagged" };
  }

  // Cap tagQuantity to remaining
  var remaining = totalQ > 0 ? totalQ - currentTaggedQ : 0;
  if (totalQ > 0 && tagQuantity > remaining) tagQuantity = remaining;

  var ck = lookupChecklist(ut.checklist_id);
  if (!ck) return { error: "Checklist template not found" };
  var order = lookupOrder(orderId);
  if (!order) return { error: "Order not found" };

  var customerLabel = lookupCustomerLabel(order.customer_id);
  var orderTypeLabel = lookupOrderTypeLabel(order.order_type);
  var responses = safeParseJSON(ut.responses, []);
  var remarks = safeParseJSON(ut.remarks, {});
  var now = new Date().toISOString();

  // Update the per-checklist response tab: change "UNTAGGED" to actual order ID (only on first tag)
  if (!ut.tagged_order_id || String(ut.tagged_order_id).trim() === "") {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var respSheet = ss.getSheetByName(ck.name);
    if (respSheet && respSheet.getLastRow() >= 3) {
      var data = respSheet.getDataRange().getValues();
      for (var r = data.length - 1; r >= 2; r--) {
        if (String(data[r][0]) === "UNTAGGED" && String(data[r][5]) === String(ut.submitted_at)) {
          respSheet.getRange(r + 1, 1).setValue(orderId);
          respSheet.getRange(r + 1, 2).setValue(order.name);
          respSheet.getRange(r + 1, 3).setValue(customerLabel);
          break;
        }
      }
    }
  }

  // Create OrderChecklist entry
  var ocId = "oc_" + nextId() + "_ut";
  appendToSheet(SHEETS.ORDER_CHECKLISTS, {
    id: ocId, order_id: orderId, checklist_id: ut.checklist_id,
    status: "completed", completed_at: ut.submitted_at, completed_by: ut.person, work_date: ut.date,
    auto_id: ut.auto_id || "",
  });

  // Add to Master Summary
  addToMasterSummary({
    orderId: orderId, orderName: order.name, customer: customerLabel,
    orderType: orderTypeLabel, checklistName: ck.name,
    person: ut.person, date: ut.date, submittedAt: ut.submitted_at,
  });

  // Update allocation tracking
  var allocations = safeParseJSON(ut.allocations, []);
  allocations.push({ orderId: orderId, quantity: tagQuantity, taggedAt: now });
  var newTaggedQ = currentTaggedQ + tagQuantity;
  var isFullyTagged = totalQ > 0 ? newTaggedQ >= totalQ : true;

  ut.tagged_order_id = isFullyTagged ? orderId : (ut.tagged_order_id || "");
  ut.tagged_quantity = newTaggedQ;
  ut.allocations = JSON.stringify(allocations);
  var utIdx = findRowIndex(SHEETS.UNTAGGED_CHECKLISTS, utId);
  if (utIdx > 0) updateSheetRow(SHEETS.UNTAGGED_CHECKLISTS, utIdx, ut);

  // Update order lines tagged quantity if applicable
  if (tagQuantity > 0) {
    var orderLines = safeParseJSON(order.order_lines, []);
    if (orderLines.length > 0) {
      // Distribute tagged quantity across blend lines (simple: fill first available)
      var remaining2 = tagQuantity;
      for (var ol = 0; ol < orderLines.length && remaining2 > 0; ol++) {
        var lineRemaining = (parseFloat(orderLines[ol].quantity) || 0) - (parseFloat(orderLines[ol].taggedQuantity) || 0);
        if (lineRemaining > 0) {
          var alloc = Math.min(remaining2, lineRemaining);
          orderLines[ol].taggedQuantity = (parseFloat(orderLines[ol].taggedQuantity) || 0) + alloc;
          remaining2 -= alloc;
        }
      }
      var orderIdx = findRowIndex(SHEETS.ORDERS, orderId);
      if (orderIdx > 0) {
        order.order_lines = JSON.stringify(orderLines);
        updateSheetRow(SHEETS.ORDERS, orderIdx, order);
      }
    }
  }

  writeAuditLog(user, "tag_untagged", "UntaggedChecklist", utId, "Tagged " + tagQuantity + " to " + orderId);
  return { success: true, fullyTagged: isFullyTagged };
}

// ─── Cross-Checklist Linking (Approval Gate) ──────────────────

function handleGetApprovedEntries(params) {
  var checklistId = params.id || params.checklist_id || "";
  if (!checklistId) return { error: "Missing checklist id" };
  return getApprovedEntriesForChecklist(checklistId);
}

// Returns rich entry objects: { linkedId, autoId, orderId, orderName, person, date, submittedAt, responses: [{question,response}] }
function getApprovedEntriesForChecklist(checklistId) {
  var ck = lookupChecklist(checklistId);
  if (!ck) return [];

  var nq = ck.questions;
  var approvalGateIdx = -1, masterQtyIdx = -1;
  for (var i = 0; i < nq.length; i++) {
    if (nq[i].isApprovalGate) approvalGateIdx = i;
    if (nq[i].isMasterQuantity || (nq[i].inventoryLink && nq[i].inventoryLink.enabled && nq[i].inventoryLink.txType === "IN")) masterQtyIdx = i;
  }
  // Auto ID is the only identifier used for linking
  var autoIdEnabled = ck.autoIdConfig && ck.autoIdConfig.enabled;
  if (approvalGateIdx < 0) return [];
  if (!autoIdEnabled) return [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ck.name);
  if (!sheet || sheet.getLastRow() < 3) return [];

  // Build {orderId -> autoId} map from OrderChecklists for this checklist
  ensureSheetHasAllColumns(SHEETS.ORDER_CHECKLISTS);
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  var orderIdToAutoId = {};
  for (var oi = 0; oi < ocs.length; oi++) {
    if (String(ocs[oi].checklist_id) === String(checklistId) && ocs[oi].auto_id) {
      orderIdToAutoId[String(ocs[oi].order_id)] = String(ocs[oi].auto_id);
    }
  }
  // Also look up by submitted_at for untagged entries
  ensureSheetHasAllColumns(SHEETS.UNTAGGED_CHECKLISTS);
  var uts = getRows(SHEETS.UNTAGGED_CHECKLISTS);
  var submittedAtToAutoId = {};
  for (var ui = 0; ui < uts.length; ui++) {
    if (isDeleted(uts[ui])) continue;
    if (String(uts[ui].checklist_id) === String(checklistId) && uts[ui].auto_id) {
      submittedAtToAutoId[String(uts[ui].submitted_at)] = String(uts[ui].auto_id);
    }
  }

  var remarkIdx = getRemarkIndices(nq);
  var qTexts = questionTexts(nq);
  var remarkHeaders = remarkIdx.map(function(ri2) { return "Remarks: " + nq[ri2].text; });
  var data = sheet.getDataRange().getValues();
  // Build header-based column map from row 2
  var headers = data.length >= 2 ? data[1].map(String) : [];
  var qColMap = {};
  for (var hi = 0; hi < qTexts.length; hi++) { qColMap[hi] = headers.indexOf(qTexts[hi]); }
  var rmkColMap = {};
  for (var hj = 0; hj < remarkHeaders.length; hj++) { rmkColMap[hj] = headers.indexOf(remarkHeaders[hj]); }

  var results = [];
  var seenIds = {};
  var approvalColIdx = qColMap[approvalGateIdx];
  var masterQtyColIdx = masterQtyIdx >= 0 ? qColMap[masterQtyIdx] : -1;
  for (var r = 2; r < data.length; r++) {
    var approvalVal = approvalColIdx >= 0 ? String(data[r][approvalColIdx] || "").trim().toLowerCase() : "";
    if (approvalVal === "yes") {
      var rowOrderIdEarly = String(data[r][0] || "");
      var rowSubmittedAtEarly = String(data[r][5] || "");
      var rowAutoIdEarly = orderIdToAutoId[rowOrderIdEarly] || submittedAtToAutoId[rowSubmittedAtEarly] || "";
      // Auto ID is the sole identifier for linking
      var linkedIdVal = rowAutoIdEarly;
      if (linkedIdVal && !seenIds[linkedIdVal]) {
        seenIds[linkedIdVal] = true;
        var responses = [];
        for (var q = 0; q < nq.length; q++) {
          var remark = "";
          for (var ri = 0; ri < remarkIdx.length; ri++) {
            if (remarkIdx[ri] === q && rmkColMap[ri] >= 0) remark = String(data[r][rmkColMap[ri]] || "");
          }
          var rspVal = qColMap[q] >= 0 ? String(data[r][qColMap[q]] || "") : "";
          if (nq[q] && nq[q].type === "inventory_item" && rspVal && rspVal.indexOf("inv_") === 0) {
            rspVal = inventoryItemNameById(rspVal);
          }
          responses.push({ question: nq[q].text, response: rspVal, remark: remark });
        }
        var rowOrderId = String(data[r][0] || "");
        var rowSubmittedAt = String(data[r][5] || "");
        var autoIdForRow = orderIdToAutoId[rowOrderId] || submittedAtToAutoId[rowSubmittedAt] || "";

        // If this checklist has a master quantity, compute total and allocated
        var totalMasterQty = 0;
        if (masterQtyColIdx >= 0) totalMasterQty = parseFloat(data[r][masterQtyColIdx]) || 0;
        var allocatedQty = autoIdForRow ? getAllocatedQuantityForAutoId(autoIdForRow) : 0;

        results.push({
          linkedId: linkedIdVal,
          autoId: autoIdForRow,
          orderId: rowOrderId,
          orderName: String(data[r][1] || ""),
          person: String(data[r][3] || ""),
          date: String(data[r][4] || ""),
          submittedAt: rowSubmittedAt,
          responses: responses,
          masterQuantity: totalMasterQty,
          allocatedQuantity: allocatedQty,
          remainingMasterQuantity: totalMasterQty > 0 ? (totalMasterQty - allocatedQty) : 0,
        });
      }
    }
  }
  return results;
}

// Compute used quantity: how much of a source entry's quantity field has been consumed by referencing checklists
// sourceChecklistId: the source checklist (e.g., Green Bean QC)
// linkedIdValue: the selected linked ID value (e.g., "SHP-001")
// quantityFieldName: the name of the quantity field in the SOURCE checklist (e.g., "Quantity received")
// consumerChecklistId: the consuming checklist (e.g., Roasted Beans QC)
// consumerQuantityFieldName: the field name in the consumer that holds "quantity used" (e.g., "Quantity input")
// consumerLinkedFieldName: the field in the consumer that holds the linked ID reference
function getUsedQuantity(consumerChecklistId, linkedIdValue, consumerQuantityFieldName, consumerLinkedFieldName) {
  var ck = lookupChecklist(consumerChecklistId);
  if (!ck) return 0;

  var nq = ck.questions;
  var linkedFieldIdx = -1, qtyFieldIdx = -1;
  for (var i = 0; i < nq.length; i++) {
    if (nq[i].text === consumerLinkedFieldName || nq[i].linkedSource) linkedFieldIdx = i;
    if (nq[i].text === consumerQuantityFieldName) qtyFieldIdx = i;
  }
  // Better: find linked field by linkedSource property
  for (var j = 0; j < nq.length; j++) {
    if (nq[j].linkedSource) linkedFieldIdx = j;
  }
  if (linkedFieldIdx < 0 || qtyFieldIdx < 0) return 0;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(ck.name);
  if (!sheet || sheet.getLastRow() < 3) return 0;

  var data = sheet.getDataRange().getValues();
  // Use header-based column lookup
  var hdrs = data.length >= 2 ? data[1].map(String) : [];
  var linkedColIdx = hdrs.indexOf(nq[linkedFieldIdx].text);
  var qtyColIdx = hdrs.indexOf(nq[qtyFieldIdx].text);
  if (linkedColIdx < 0 || qtyColIdx < 0) return 0;

  var total = 0;
  for (var r = 2; r < data.length; r++) {
    var refVal = String(data[r][linkedColIdx] || "").trim();
    if (refVal === String(linkedIdValue)) {
      var qtyVal = parseFloat(data[r][qtyColIdx]) || 0;
      total += qtyVal;
    }
  }
  return total;
}

// CHAIN_CONFIG: defines the coffee workflow chain and quantity tracking relationships
var CHAIN_CONFIG = {
  // checklistId -> { sourceChecklistId, sourceQuantityField, consumerQuantityField }
  // This is populated dynamically from checklist questions, but we define the chain order here
  // Chain: Sample QC -> Green Bean Shipment QC -> Roasting QC -> Grinding & Packing -> Dispatch
  chainOrder: ["ck_sample_qc", "ck_green_beans", "ck_roasted_beans", "ck_grinding"],
  // Quantity tracking pairs: consumer checklist -> { sourceField (in source), consumerField (in consumer) }
  quantityTracking: {
    "ck_roasted_beans": { sourceQuantityField: "Quantity received", consumerQuantityField: "Quantity input" },
    "ck_grinding": { sourceQuantityField: "Quantity output", consumerQuantityField: "Total Net weight" },
  },
};

// handleGetLinkedEntries: returns approved entries with remaining quantity info
function handleGetLinkedEntries(params) {
  var checklistId = params.checklist_id || params.id || "";
  if (!checklistId) return { error: "Missing checklist_id" };

  var ck = lookupChecklist(checklistId);
  if (!ck) return { error: "Checklist not found" };

  // Find the linked source field
  var nq = ck.questions;
  var linkedField = null;
  for (var i = 0; i < nq.length; i++) {
    if (nq[i].linkedSource && nq[i].linkedSource.checklistId) {
      linkedField = nq[i];
      break;
    }
  }
  if (!linkedField) return [];

  var sourceChecklistId = linkedField.linkedSource.checklistId;
  var entries = getApprovedEntriesForChecklist(sourceChecklistId);

  // Add quantity tracking if configured — uses QuantityAllocations as single source of truth
  // (same data source as backend validation in processQuantityAllocationsForSubmission)
  var qtyConfig = CHAIN_CONFIG.quantityTracking[checklistId];
  if (qtyConfig) {
    for (var e = 0; e < entries.length; e++) {
      var totalQty = 0;
      for (var r = 0; r < entries[e].responses.length; r++) {
        if (entries[e].responses[r].question === qtyConfig.sourceQuantityField) {
          totalQty = parseFloat(entries[e].responses[r].response) || 0;
          break;
        }
      }
      var entryAutoId = entries[e].autoId || entries[e].linkedId;
      var allocQty = getAllocatedQuantityForAutoId(entryAutoId);
      entries[e].totalQuantity = totalQty;
      entries[e].usedQuantity = allocQty;
      entries[e].remainingQuantity = totalQty - allocQty;
    }
  }

  // Universal quantity tracking via QuantityAllocations + master quantity
  // (entries already have masterQuantity, allocatedQuantity, remainingMasterQuantity from getApprovedEntriesForChecklist)
  for (var k = 0; k < entries.length; k++) {
    if (entries[k].masterQuantity > 0 && entries[k].totalQuantity === undefined) {
      entries[k].totalQuantity = entries[k].masterQuantity;
      entries[k].usedQuantity = entries[k].allocatedQuantity || 0;
      entries[k].remainingQuantity = entries[k].remainingMasterQuantity;
    }
  }

  // Phase 6 Fix 1: Mark fully-allocated entries so the frontend can grey them out
  for (var m = 0; m < entries.length; m++) {
    var rem = entries[m].remainingQuantity;
    var tot = entries[m].totalQuantity;
    entries[m].fullyAllocated = (tot > 0 && rem <= 0);
  }

  return entries;
}

function buildApprovedEntriesCache(checklists) {
  var cache = {};
  for (var i = 0; i < checklists.length; i++) {
    var ck = checklists[i];
    var hasGate = false;
    for (var q = 0; q < ck.questions.length; q++) {
      if (ck.questions[q].isApprovalGate) hasGate = true;
    }
    var hasAutoId = ck.autoIdConfig && ck.autoIdConfig.enabled;
    if (hasGate && hasAutoId) {
      cache[ck.id] = getApprovedEntriesForChecklist(ck.id);
      // Project masterQuantity to legacy fields for the dropdown UI
      var arr = cache[ck.id];
      for (var ai = 0; ai < arr.length; ai++) {
        if (arr[ai].masterQuantity > 0 && arr[ai].totalQuantity === undefined) {
          arr[ai].totalQuantity = arr[ai].masterQuantity;
          arr[ai].usedQuantity = arr[ai].allocatedQuantity || 0;
          arr[ai].remainingQuantity = arr[ai].remainingMasterQuantity;
        }
      }
    }
  }
  // Also build linked entries with quantity for consumer checklists
  for (var j = 0; j < checklists.length; j++) {
    var ck2 = checklists[j];
    for (var q2 = 0; q2 < ck2.questions.length; q2++) {
      if (ck2.questions[q2].linkedSource && ck2.questions[q2].linkedSource.checklistId) {
        var srcId = ck2.questions[q2].linkedSource.checklistId;
        if (!cache[srcId]) {
          cache[srcId] = getApprovedEntriesForChecklist(srcId);
        }
        // Add quantity info — unified to QuantityAllocations as single source of truth
        var qtyConfig = CHAIN_CONFIG.quantityTracking[ck2.id];
        if (qtyConfig && cache[srcId]) {
          for (var e = 0; e < cache[srcId].length; e++) {
            var entry = cache[srcId][e];
            var totalQty = 0;
            for (var r = 0; r < entry.responses.length; r++) {
              if (entry.responses[r].question === qtyConfig.sourceQuantityField) {
                totalQty = parseFloat(entry.responses[r].response) || 0; break;
              }
            }
            var entryAutoIdC = entry.autoId || entry.linkedId;
            var allocQtyC = getAllocatedQuantityForAutoId(entryAutoIdC);
            entry.totalQuantity = totalQty;
            entry.usedQuantity = allocQtyC;
            entry.remainingQuantity = totalQty - allocQtyC;
          }
        }
        break;
      }
    }
  }
  return cache;
}

// ─── Inventory Management ─────────────────────────────────────

function handleGetInventoryCategories() {
  return getRows(SHEETS.INVENTORY_CATEGORIES).map(function(r) { return { id: r.id, name: r.name }; });
}

function handleCreateInventoryCategory(body, user) {
  var id = body.id || ("icat_" + nextId());
  var obj = { id: id, name: body.name };
  appendToSheet(SHEETS.INVENTORY_CATEGORIES, obj);
  writeAuditLog(user, "create", "InventoryCategory", id, body.name);
  return obj;
}

function handleGetInventoryItems() {
  ensureSheetHasAllColumns(SHEETS.INVENTORY_ITEMS);
  return getRows(SHEETS.INVENTORY_ITEMS).map(function(r) {
    return {
      id: r.id, category: r.category, name: r.name, unit: r.unit || "kg",
      openingStock: parseFloat(r.opening_stock) || 0,
      currentStock: parseFloat(r.current_stock) || 0,
      minStockAlert: parseFloat(r.min_stock_alert) || 0,
      createdAt: r.created_at, isActive: r.is_active !== "false" && r.is_active !== false,
      classificationId: r.classification_id || "",
      classificationLabel: lookupClassificationLabel(r.classification_id),
      abbreviation: String(r.abbreviation || "").toUpperCase(),
      equivalentItems: safeParseJSON(r.equivalent_items, []),
    };
  });
}

function validateAbbreviation(ab) {
  var s = String(ab || "").toUpperCase();
  if (!/^[A-Z0-9]{2,6}$/.test(s)) return null;
  return s;
}

function handleCreateInventoryItem(body, user) {
  ensureSheetHasAllColumns(SHEETS.INVENTORY_ITEMS);
  var id = body.id || ("inv_" + nextId());
  var openingStock = parseFloat(body.openingStock) || 0;
  var ab = validateAbbreviation(body.abbreviation);
  if (!ab) return { error: "Abbreviation is required (2-6 uppercase letters/digits)" };
  var abDup = checkAbbreviationUniqueness(ab, body.category, "");
  if (abDup) return { error: abDup };
  var obj = {
    id: id, category: body.category, name: body.name, unit: body.unit || "kg",
    opening_stock: openingStock, current_stock: openingStock,
    min_stock_alert: parseFloat(body.minStockAlert) || 0,
    created_at: new Date().toISOString(), is_active: "true",
    abbreviation: ab,
    equivalent_items: JSON.stringify(Array.isArray(body.equivalentItems) ? body.equivalentItems : []),
    classification_id: body.classificationId || "",
  };
  appendToSheet(SHEETS.INVENTORY_ITEMS, obj);
  // Create opening stock ledger entry if > 0
  if (openingStock > 0) {
    appendToSheet(SHEETS.INVENTORY_LEDGER, {
      id: "led_" + nextId(), item_id: id, item_name: body.name,
      category: body.category || "",
      date: new Date().toISOString().split("T")[0], type: "IN",
      quantity: openingStock, balance_after: openingStock,
      reference_type: "manual", reference_id: "", notes: "Opening stock",
      done_by: user.displayName || user.username, created_at: new Date().toISOString(),
    });
  }
  writeAuditLog(user, "create", "InventoryItem", id, body.name);
  return { id: id, category: body.category, name: body.name, unit: body.unit || "kg", openingStock: openingStock, currentStock: openingStock, minStockAlert: obj.min_stock_alert, createdAt: obj.created_at, isActive: true, abbreviation: ab, classificationId: obj.classification_id || "" };
}

function handleUpdateInventoryItem(body, user) {
  ensureSheetHasAllColumns(SHEETS.INVENTORY_ITEMS);
  var id = body.id;
  var idx = findRowIndex(SHEETS.INVENTORY_ITEMS, id);
  if (idx < 0) return { error: "Item not found" };
  var rows = getRows(SHEETS.INVENTORY_ITEMS);
  var item = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { item = rows[i]; break; } }
  if (!item) return { error: "Item not found" };
  if (body.name !== undefined) item.name = body.name;
  if (body.category !== undefined) item.category = body.category;
  if (body.unit !== undefined) item.unit = body.unit;
  if (body.minStockAlert !== undefined) item.min_stock_alert = parseFloat(body.minStockAlert) || 0;
  if (body.isActive !== undefined) item.is_active = body.isActive ? "true" : "false";
  if (body.abbreviation !== undefined) {
    var ab2 = validateAbbreviation(body.abbreviation);
    if (!ab2) return { error: "Abbreviation must be 2-6 uppercase letters/digits" };
    var abDup2 = checkAbbreviationUniqueness(ab2, item.category, id);
    if (abDup2) return { error: abDup2 };
    item.abbreviation = ab2;
  }
  if (body.classificationId !== undefined) item.classification_id = body.classificationId || "";
  // Two-way equivalent linking
  if (body.equivalentItems !== undefined) {
    var oldEqList = safeParseJSON(item.equivalent_items, []);
    var newEqList = Array.isArray(body.equivalentItems) ? body.equivalentItems : [];
    item.equivalent_items = JSON.stringify(newEqList);
    // The current item's category (read from existing record before any updates)
    var currentCategory = item.category;

    // Detect added: items in newEqList whose itemId is NOT in oldEqList
    var addedItems = newEqList.filter(function(n) {
      for (var a = 0; a < oldEqList.length; a++) { if (oldEqList[a].itemId === n.itemId) return false; }
      return true;
    });
    // Detect removed: items in oldEqList whose itemId is NOT in newEqList
    var removedItems = oldEqList.filter(function(o) {
      for (var a = 0; a < newEqList.length; a++) { if (newEqList[a].itemId === o.itemId) return false; }
      return true;
    });

    // For each newly added equivalent, add the reverse link on the target item
    for (var ai = 0; ai < addedItems.length; ai++) {
      syncEquivalentLink(addedItems[ai].itemId, id, currentCategory, true);
    }
    // For each removed equivalent, remove the reverse link from the target item
    for (var ri = 0; ri < removedItems.length; ri++) {
      syncEquivalentLink(removedItems[ri].itemId, id, currentCategory, false);
    }
    // Invalidate cache since we may have modified other rows
    if (addedItems.length > 0 || removedItems.length > 0) invalidateCache(SHEETS.INVENTORY_ITEMS);
  }
  updateSheetRow(SHEETS.INVENTORY_ITEMS, idx, item);
  writeAuditLog(user, "update", "InventoryItem", id, item.name);
  // Return full updated item so frontend can sync state accurately
  // Also return all inventory items so frontend can update reverse links in UI
  var allItems = handleGetInventoryItems();
  return {
    success: true,
    id: item.id, category: item.category, name: item.name, unit: item.unit || "kg",
    openingStock: parseFloat(item.opening_stock) || 0,
    currentStock: parseFloat(item.current_stock) || 0,
    minStockAlert: parseFloat(item.min_stock_alert) || 0,
    createdAt: item.created_at, isActive: item.is_active !== "false" && item.is_active !== false,
    abbreviation: String(item.abbreviation || "").toUpperCase(),
    equivalentItems: safeParseJSON(item.equivalent_items, []),
    allItems: allItems,
  };
}

// Helper: add or remove a reverse equivalent link on a target item.
// equivalentItems are objects: { category: "Green Beans", itemId: "inv_xxx" }
function syncEquivalentLink(targetItemId, sourceItemId, sourceCategory, shouldAdd) {
  var tIdx = findRowIndex(SHEETS.INVENTORY_ITEMS, targetItemId);
  if (tIdx < 0) return;
  // Re-read rows fresh (cache may be stale after prior writes)
  invalidateCache(SHEETS.INVENTORY_ITEMS);
  var tRows = getRows(SHEETS.INVENTORY_ITEMS);
  var target = null;
  for (var i = 0; i < tRows.length; i++) { if (String(tRows[i].id) === String(targetItemId)) { target = tRows[i]; break; } }
  if (!target) return;
  var eqList = safeParseJSON(target.equivalent_items, []);
  var alreadyHas = false;
  for (var j = 0; j < eqList.length; j++) { if (eqList[j].itemId === sourceItemId) { alreadyHas = true; break; } }
  if (shouldAdd && !alreadyHas) {
    eqList.push({ category: sourceCategory, itemId: sourceItemId });
    target.equivalent_items = JSON.stringify(eqList);
    updateSheetRow(SHEETS.INVENTORY_ITEMS, tIdx, target);
    invalidateCache(SHEETS.INVENTORY_ITEMS);
  } else if (!shouldAdd && alreadyHas) {
    eqList = eqList.filter(function(e) { return e.itemId !== sourceItemId; });
    target.equivalent_items = JSON.stringify(eqList);
    updateSheetRow(SHEETS.INVENTORY_ITEMS, tIdx, target);
    invalidateCache(SHEETS.INVENTORY_ITEMS);
  }
}

function handleGetInventoryLedger(params) {
  var itemId = params.item_id || params.id || "";
  var rows = getRows(SHEETS.INVENTORY_LEDGER);
  var filtered = itemId ? rows.filter(function(r) { return String(r.item_id) === String(itemId); }) : rows;
  filtered.sort(function(a, b) { return String(b.created_at).localeCompare(String(a.created_at)); });
  return filtered.map(function(r) {
    return {
      id: r.id, itemId: r.item_id, itemName: r.item_name,
      category: r.category || "",
      date: r.date,
      type: r.type, quantity: parseFloat(r.quantity) || 0,
      balanceAfter: parseFloat(r.balance_after) || 0,
      referenceType: r.reference_type, referenceId: r.reference_id,
      notes: r.notes, doneBy: r.done_by, createdAt: r.created_at,
      classificationId: r.classification_id || "",
    };
  });
}

function handleAddInventoryAdjustment(body, user) {
  var itemId = body.itemId;
  var idx = findRowIndex(SHEETS.INVENTORY_ITEMS, itemId);
  if (idx < 0) return { error: "Item not found" };
  var rows = getRows(SHEETS.INVENTORY_ITEMS);
  var item = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(itemId)) { item = rows[i]; break; } }
  if (!item) return { error: "Item not found" };

  var qty = parseFloat(body.quantity);
  if (isNaN(qty) || qty === 0) return { error: "Quantity must be non-zero" };
  var adjType = body.adjustmentType === "reduction" ? "OUT" : "IN";
  // IN (addition) must be strictly positive. OUT (reduction) may be negative — a negative
  // reduction is a manual correction that restores stock (e.g. reversing an over-deduction).
  if (adjType === "IN" && qty <= 0) return { error: "Addition quantity must be positive" };
  var currentStock = parseFloat(item.current_stock) || 0;
  var newStock = adjType === "IN" ? currentStock + qty : currentStock - qty;

  item.current_stock = newStock;
  updateSheetRow(SHEETS.INVENTORY_ITEMS, idx, item);
  invalidateCache(SHEETS.INVENTORY_ITEMS);

  var ledgerEntry = {
    id: "led_" + nextId(), item_id: itemId, item_name: item.name,
    category: item.category || "",
    date: body.date || new Date().toISOString().split("T")[0], type: "ADJUSTMENT",
    quantity: (adjType === "OUT" ? -qty : qty), balance_after: newStock,
    reference_type: "manual", reference_id: "",
    notes: (adjType === "IN" ? "Addition" : "Reduction") + (body.notes ? ": " + body.notes : ""),
    done_by: user.displayName || user.username, created_at: new Date().toISOString(),
    classification_id: item.classification_id || "",
  };
  appendToSheet(SHEETS.INVENTORY_LEDGER, ledgerEntry);
  writeAuditLog(user, "inventory_adjustment", "InventoryItem", itemId, adjType + " " + qty + " " + item.unit);
  return { success: true, newStock: newStock };
}

function handleEditInventoryLedger(body, user) {
  var id = body.id;
  var idx = findRowIndex(SHEETS.INVENTORY_LEDGER, id);
  if (idx < 0) return { error: "Ledger entry not found" };
  var rows = getRows(SHEETS.INVENTORY_LEDGER);
  var entry = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { entry = rows[i]; break; } }
  if (!entry) return { error: "Ledger entry not found" };
  if (body.notes !== undefined) entry.notes = body.notes;
  if (body.date !== undefined) entry.date = body.date;
  updateSheetRow(SHEETS.INVENTORY_LEDGER, idx, entry);
  writeAuditLog(user, "edit_ledger", "InventoryLedger", id, "Edited");
  return { success: true };
}

function lookupClassificationLabel(classificationId) {
  if (!classificationId) return "";
  var rows = getRows(SHEETS.ROAST_CLASSIFICATIONS);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(classificationId)) return rows[i].name || "";
  }
  return "";
}

function handleGetInventorySummary() {
  var items = getRows(SHEETS.INVENTORY_ITEMS);
  var summary = { greenBeans: 0, roastedBeans: 0, packedGoods: 0, lowStockCount: 0 };
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    if (item.is_active === "false" || item.is_active === false) continue;
    var stock = parseFloat(item.current_stock) || 0;
    var cat = String(item.category || "").trim().toLowerCase();
    if (cat === "green beans") summary.greenBeans += stock;
    else if (cat === "roasted beans") summary.roastedBeans += stock;
    else if (cat === "packing items") summary.packedGoods += stock;
    var minAlert = parseFloat(item.min_stock_alert) || 0;
    if (minAlert > 0 && stock < minAlert) summary.lowStockCount++;
  }
  return summary;
}

// Helper: create inventory transaction (used by checklist submissions)
function createInventoryTransaction(itemId, type, quantity, refType, refId, notes, doneBy, questionIndex, classificationId) {
  var idx = findRowIndex(SHEETS.INVENTORY_ITEMS, itemId);
  if (idx < 0) {
    Logger.log("⚠ createInventoryTransaction: item '" + itemId + "' not found in InventoryItems — ledger entry skipped (type=" + type + ", qty=" + quantity + ", refId=" + (refId || "") + ")");
    return { warning: "item not found: " + itemId };
  }
  var rows = getRows(SHEETS.INVENTORY_ITEMS);
  var item = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(itemId)) { item = rows[i]; break; } }
  if (!item) {
    Logger.log("⚠ createInventoryTransaction: item '" + itemId + "' found by index but not in rows cache — ledger entry skipped");
    return { warning: "item row not found: " + itemId };
  }

  var currentStock = parseFloat(item.current_stock) || 0;
  var newStock = type === "IN" ? currentStock + quantity : currentStock - quantity;
  // Never block the transaction — just flag it so the user is aware the item is oversold.
  if (type === "OUT" && newStock < 0) {
    Logger.log("⚠ NEGATIVE STOCK WARNING: " + item.name + " (" + itemId + ") — OUT " + quantity + " " + (item.unit || "") + " brings stock from " + currentStock + " to " + newStock + " (refType=" + (refType || "manual") + ", refId=" + (refId || "") + ", by=" + (doneBy || "") + ")");
  }
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

// Look up an inventory item's category by id. Used to populate the InventoryLedger
// category column on writes and to backfill historical rows.
function getInventoryCategoryForItemId(itemId) {
  if (!itemId) return "";
  var rows = getRows(SHEETS.INVENTORY_ITEMS);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(itemId)) return rows[i].category || "";
  }
  return "";
}

// One-time backfill: populate the "category" column on every InventoryLedger row that
// currently has it blank. Looks up category from InventoryItems via item_id. Safe to
// re-run — only writes when the cell is empty and the lookup succeeds. Returns a
// summary { scanned, updated, missing } so an admin call can see what happened.
var _ledgerCategoryBackfillDone = false;
function backfillInventoryLedgerCategories() {
  if (_ledgerCategoryBackfillDone) return { scanned: 0, updated: 0, missing: 0, skipped: true };
  _ledgerCategoryBackfillDone = true;
  var sheet = getSheet(SHEETS.INVENTORY_LEDGER);
  if (!sheet) return { scanned: 0, updated: 0, missing: 0 };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { scanned: 0, updated: 0, missing: 0 };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  var catCol = headers.indexOf("category");
  var itemIdCol = headers.indexOf("item_id");
  if (catCol < 0 || itemIdCol < 0) return { scanned: 0, updated: 0, missing: 0 };

  // Build itemId → category lookup once
  var items = getRows(SHEETS.INVENTORY_ITEMS);
  var catByItemId = {};
  for (var ii = 0; ii < items.length; ii++) {
    catByItemId[String(items[ii].id)] = items[ii].category || "";
  }

  var range = sheet.getRange(2, 1, lastRow - 1, headers.length);
  var values = range.getValues();
  var scanned = 0, updated = 0, missing = 0;
  for (var r = 0; r < values.length; r++) {
    scanned++;
    var existing = String(values[r][catCol] || "").trim();
    if (existing) continue; // already populated
    var itemId = String(values[r][itemIdCol] || "").trim();
    if (!itemId) { missing++; continue; }
    var cat = catByItemId[itemId];
    if (!cat) { missing++; continue; }
    values[r][catCol] = cat;
    updated++;
  }
  if (updated > 0) {
    range.setValues(values);
    invalidateCache(SHEETS.INVENTORY_LEDGER);
  }
  Logger.log("backfillInventoryLedgerCategories: scanned=" + scanned + " updated=" + updated + " missing=" + missing);
  return { scanned: scanned, updated: updated, missing: missing };
}

// Helper: reverse all inventory ledger entries that were written for a given reference
// (order checklist id, untagged response id, etc). Writes compensating entries with the
// opposite sign and decrements/increments current_stock back. Does NOT delete existing rows
// so the audit trail is preserved.
// Returns number of entries reversed.
function reverseInventoryLedgerForRef(refType, refId, doneBy) {
  if (!refId) return 0;
  var ledger = getRows(SHEETS.INVENTORY_LEDGER);
  var matches = [];
  for (var i = 0; i < ledger.length; i++) {
    var row = ledger[i];
    if (String(row.reference_id) !== String(refId)) continue;
    if (refType && String(row.reference_type) !== String(refType)) continue;
    // Skip entries that are themselves reversals (notes prefix marker) to prevent double-reverse
    if (String(row.notes || "").indexOf("[REVERSAL]") === 0) continue;
    matches.push(row);
  }
  if (matches.length === 0) return 0;
  Logger.log("reverseInventoryLedgerForRef: reversing " + matches.length + " entries for " + refType + "/" + refId);
  for (var j = 0; j < matches.length; j++) {
    var orig = matches[j];
    var origType = String(orig.type || "");
    var origQty = Math.abs(parseFloat(orig.quantity) || 0);
    if (!orig.item_id || origQty <= 0) continue;
    var reverseType = origType === "IN" ? "OUT" : "IN";
    createInventoryTransaction(
      orig.item_id, reverseType, origQty,
      refType || orig.reference_type || "manual",
      refId,
      "[REVERSAL] of led:" + orig.id + " — " + (orig.notes || ""),
      doneBy || "",
      orig.question_index
    );
  }
  invalidateCache(SHEETS.INVENTORY_LEDGER);
  return matches.length;
}

// ═══════════════════════════════════════════════════════════════
// ─── PHASE 2: Classification System ─────────────────────────
// ═══════════════════════════════════════════════════════════════

function handleGetClassifications() {
  ensureSheetHasAllColumns(SHEETS.ROAST_CLASSIFICATIONS);
  var rows = getRows(SHEETS.ROAST_CLASSIFICATIONS);
  var result = { roast_degree: [], grind_size: [] };
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (r.is_active === "false" || r.is_active === false) continue;
    var entry = { id: r.id, name: r.name, type: r.type, description: r.description || "", createdBy: r.created_by || "", createdAt: r.created_at || "" };
    var t = String(r.type || "").toLowerCase();
    if (t === "roast_degree" && result.roast_degree) result.roast_degree.push(entry);
    else if (t === "grind_size" && result.grind_size) result.grind_size.push(entry);
  }
  return result;
}

function handleAddClassification(body, user) {
  ensureSheetHasAllColumns(SHEETS.ROAST_CLASSIFICATIONS);
  var name = String(body.name || "").trim();
  if (!name) return { error: "Name is required" };
  var type = String(body.type || "").toLowerCase();
  if (type !== "roast_degree" && type !== "grind_size") return { error: "Type must be 'roast_degree' or 'grind_size'" };
  // Duplicate check (case-insensitive within same type)
  var existing = getRows(SHEETS.ROAST_CLASSIFICATIONS);
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i].type) === type && String(existing[i].name).toLowerCase() === name.toLowerCase() && existing[i].is_active !== "false") {
      return { error: "'" + name + "' already exists in " + type };
    }
  }
  var now = new Date().toISOString();
  var id = "rc_" + nextId();
  var obj = { id: id, name: name, type: type, description: body.description || "", created_by: user.displayName || user.username, created_at: now, updated_at: now, is_active: "true" };
  appendToSheet(SHEETS.ROAST_CLASSIFICATIONS, obj);
  writeAuditLog(user, "create", "RoastClassification", id, name + " (" + type + ")");
  return { id: id, name: name, type: type, description: obj.description, createdBy: obj.created_by, createdAt: now };
}

function handleEditClassification(body, user) {
  var id = body.id;
  if (!id) return { error: "Missing classification id" };
  var idx = findRowIndex(SHEETS.ROAST_CLASSIFICATIONS, id);
  if (idx < 0) return { error: "Classification not found" };
  var rows = getRows(SHEETS.ROAST_CLASSIFICATIONS);
  var entry = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { entry = rows[i]; break; } }
  if (!entry) return { error: "Classification not found" };
  var beforeState = JSON.stringify({ name: entry.name, description: entry.description });
  if (body.name !== undefined) entry.name = String(body.name).trim();
  if (body.description !== undefined) entry.description = String(body.description || "");
  entry.updated_at = new Date().toISOString();
  updateSheetRow(SHEETS.ROAST_CLASSIFICATIONS, idx, entry);
  invalidateCache(SHEETS.ROAST_CLASSIFICATIONS);
  var afterState = JSON.stringify({ name: entry.name, description: entry.description });
  writeAuditLog(user, "edit", "RoastClassification", id, "before=" + beforeState + " after=" + afterState);
  return { success: true, id: id, name: entry.name, description: entry.description };
}

function handleDeactivateClassification(body, user) {
  var id = body.id;
  if (!id) return { error: "Missing classification id" };
  var idx = findRowIndex(SHEETS.ROAST_CLASSIFICATIONS, id);
  if (idx < 0) return { error: "Classification not found" };
  var rows = getRows(SHEETS.ROAST_CLASSIFICATIONS);
  var entry = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { entry = rows[i]; break; } }
  if (!entry) return { error: "Classification not found" };
  // Check if used in any inventory item or ledger entry
  var invItems = getRows(SHEETS.INVENTORY_ITEMS);
  var usedIn = [];
  for (var j = 0; j < invItems.length; j++) {
    if (String(invItems[j].classification_id) === String(id)) usedIn.push(invItems[j].name);
  }
  if (usedIn.length > 0) {
    return { error: "Cannot deactivate — used by inventory items: " + usedIn.join(", ") + ". Remove classification from those items first." };
  }
  entry.is_active = "false";
  entry.updated_at = new Date().toISOString();
  updateSheetRow(SHEETS.ROAST_CLASSIFICATIONS, idx, entry);
  invalidateCache(SHEETS.ROAST_CLASSIFICATIONS);
  writeAuditLog(user, "deactivate", "RoastClassification", id, entry.name);
  return { success: true };
}

// ═══════════════════════════════════════════════════════════════
// ─── PHASE 5: Quantity Validation — Universal ───────────────
// ═══════════════════════════════════════════════════════════════

// Traverses the downstream chain for a given entry and checks whether a new quantity
// is safe. Returns { allowed: bool, reason: string, downstreamTotal: number }.
function validateQuantityEdit(checklistType, entryAutoId, newQuantity) {
  if (!entryAutoId) return { allowed: true, reason: "", downstreamTotal: 0 };
  newQuantity = parseFloat(newQuantity) || 0;

  // Determine what downstream checklists consume this entry
  var downstreamConfig = {
    "ck_green_beans": { consumerCkId: "ck_roasted_beans", fieldName: "Quantity input" },
    "ck_roasted_beans": { consumerCkId: "ck_grinding", fieldName: "Total Net weight" },
  };

  var cfg = downstreamConfig[checklistType];
  if (!cfg) {
    // Also check QuantityAllocations
    var allocTotal = getAllocatedQuantityForAutoId(entryAutoId);
    if (newQuantity < allocTotal) {
      return { allowed: false, reason: "Cannot reduce to " + newQuantity + ". " + allocTotal + "kg already allocated downstream.", downstreamTotal: allocTotal };
    }
    return { allowed: true, reason: "", downstreamTotal: allocTotal };
  }

  // Sum downstream usage from QuantityAllocations (preferred) and from response sheets
  var allocTotal2 = getAllocatedQuantityForAutoId(entryAutoId);
  var consumerCk = lookupChecklist(cfg.consumerCkId);
  var sheetUsed = 0;
  if (consumerCk) {
    sheetUsed = getUsedQuantity(cfg.consumerCkId, entryAutoId, cfg.fieldName, "");
  }
  var downstreamTotal = Math.max(allocTotal2, sheetUsed);

  if (newQuantity < downstreamTotal) {
    return {
      allowed: false,
      reason: "Cannot reduce quantity to " + newQuantity + ". " + downstreamTotal + "kg is already committed downstream.",
      downstreamTotal: downstreamTotal,
    };
  }
  return { allowed: true, reason: "", downstreamTotal: downstreamTotal };
}

// ═══════════════════════════════════════════════════════════════
// ─── PHASE 4 (partial): Soft Delete with Audit ─────────────
// ═══════════════════════════════════════════════════════════════

function handleSoftDeleteChecklist(body, user) {
  var id = body.id;
  var reason = body.reason || "";
  var entityType = body.entityType || ""; // "untagged" or "order_checklist"
  if (!id) return { error: "Missing id" };

  // ── Untagged checklist soft-delete ──
  if (entityType === "untagged" || String(id).indexOf("ut_") === 0) {
    ensureSheetHasAllColumns(SHEETS.UNTAGGED_CHECKLISTS);
    var utRows = getRows(SHEETS.UNTAGGED_CHECKLISTS);
    var ut = null;
    for (var i = 0; i < utRows.length; i++) { if (String(utRows[i].id) === String(id)) { ut = utRows[i]; break; } }
    if (!ut) return { error: "Untagged checklist not found" };
    if (isDeleted(ut)) return { error: "Already deleted" };

    // Check downstream dependencies
    if (ut.auto_id) {
      var refs = findUpstreamReferencesForAutoId(ut.auto_id);
      if (refs.length > 0) {
        return { error: "Cannot delete — first untag from: " + refs.join(", ") };
      }
    }
    // Check if tagged to an order
    if (ut.tagged_order_id && String(ut.tagged_order_id).trim() !== "") {
      return { error: "Cannot delete — first remove from order: " + ut.tagged_order_id };
    }

    var beforeState = JSON.stringify({ id: ut.id, checklistName: ut.checklist_name, person: ut.person, date: ut.date, totalQuantity: ut.total_quantity });
    // Soft delete
    ut.is_deleted = "true";
    var utIdx = findRowIndex(SHEETS.UNTAGGED_CHECKLISTS, id);
    if (utIdx > 0) updateSheetRow(SHEETS.UNTAGGED_CHECKLISTS, utIdx, ut);
    invalidateCache(SHEETS.UNTAGGED_CHECKLISTS);

    // Reverse inventory ledger entries
    var reversed = reverseInventoryLedgerForRef("untagged", id, user.displayName || user.username);
    var reversed2 = reverseInventoryLedgerForRef("checklist", id, user.displayName || user.username);

    writeAuditLog(user, "delete", "UntaggedChecklist", id, "Reason: " + reason + " | beforeState=" + beforeState + " | reversed " + (reversed + reversed2) + " ledger entries");
    return { success: true, reversed: reversed + reversed2 };
  }

  // ── Order checklist soft-delete (revert to pending + audit log) ──
  if (entityType === "order_checklist" || String(id).indexOf("oc_") === 0) {
    var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
    var oc = null;
    for (var j = 0; j < ocs.length; j++) { if (String(ocs[j].id) === String(id)) { oc = ocs[j]; break; } }
    if (!oc) return { error: "Order checklist not found" };

    // Check downstream
    if (oc.auto_id) {
      var refs2 = findUpstreamReferencesForAutoId(oc.auto_id);
      if (refs2.length > 0) {
        return { error: "Cannot delete — first untag from: " + refs2.join(", ") };
      }
    }

    var ck = lookupChecklist(oc.checklist_id);
    var beforeState2 = JSON.stringify({ id: oc.id, checklistId: oc.checklist_id, status: oc.status, completedBy: oc.completed_by, workDate: oc.work_date });

    // Delete responses
    if (ck) {
      deleteResponseRow(ck.name, String(oc.order_id));
      deleteFromMasterSummary(String(oc.order_id), ck.name);
    }

    // Reverse inventory
    var rev = reverseInventoryLedgerForRef("checklist", id, user.displayName || user.username);

    // Revert to pending
    oc.status = "pending"; oc.completed_at = ""; oc.completed_by = ""; oc.work_date = "";
    var ocIdx = findRowIndex(SHEETS.ORDER_CHECKLISTS, id);
    if (ocIdx > 0) updateSheetRow(SHEETS.ORDER_CHECKLISTS, ocIdx, oc);

    writeAuditLog(user, "delete", "Checklist", id, "Reason: " + reason + " | beforeState=" + beforeState2 + " | reversed " + rev + " ledger entries");
    return { success: true, reversed: rev };
  }

  return { error: "Unknown entity type. Pass entityType='untagged' or 'order_checklist'" };
}

// ═══════════════════════════════════════════════════════════════
// ─── PHASE 8: Edit Fixes ────────────────────────────────────
// ═══════════════════════════════════════════════════════════════

function handleEditUntaggedResponse(body, user) {
  var id = body.id;
  if (!id) return { error: "Missing untagged response id" };
  ensureSheetHasAllColumns(SHEETS.UNTAGGED_CHECKLISTS);
  var rows = getRows(SHEETS.UNTAGGED_CHECKLISTS);
  var ut = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { ut = rows[i]; break; } }
  if (!ut) return { error: "Untagged response not found" };
  if (isDeleted(ut)) return { error: "Cannot edit a deleted entry" };

  // Permission: admin or own submission
  if (user.role !== "admin" && String(ut.submitted_by_user_id) !== String(user.id)) {
    return { error: "You can only edit your own untagged checklists" };
  }

  var ck = lookupChecklist(ut.checklist_id);
  if (!ck) return { error: "Checklist template not found" };

  var newResponses = body.responses || [];
  var newDate = body.date;
  var newPerson = body.person;
  var remarks = body.remarks || {};

  // Build responses map and validate quantity downstream
  var responsesMap = responsesArrayToMap(newResponses);
  var oldTotalQty = parseFloat(ut.total_quantity) || 0;
  var newTotalQty = getMasterQuantityFromResponses(ck.questions, responsesMap);
  if (newTotalQty <= 0) {
    var nq = ck.questions;
    for (var qi = 0; qi < nq.length; qi++) {
      if ((nq[qi].type === "number" || nq[qi].type === "text_number") &&
          (nq[qi].text.toLowerCase().indexOf("quantity") >= 0 || nq[qi].text.toLowerCase().indexOf("weight") >= 0)) {
        var qtyVal = parseFloat(responsesMap[qi]) || 0;
        if (qtyVal > newTotalQty) newTotalQty = qtyVal;
      }
    }
  }

  // Validate downstream quantity constraints
  if (ut.auto_id && newTotalQty < oldTotalQty) {
    var vResult = validateQuantityEdit(ut.checklist_id, ut.auto_id, newTotalQty);
    if (!vResult.allowed) return { error: vResult.reason };
  }

  var beforeState = JSON.stringify({ person: ut.person, date: ut.date, totalQuantity: oldTotalQty });

  // Update fields
  if (newPerson) ut.person = newPerson;
  if (newDate) ut.date = newDate;
  ut.responses = JSON.stringify(newResponses);
  ut.remarks = JSON.stringify(remarks);
  ut.total_quantity = newTotalQty;

  var utIdx = findRowIndex(SHEETS.UNTAGGED_CHECKLISTS, id);
  if (utIdx > 0) updateSheetRow(SHEETS.UNTAGGED_CHECKLISTS, utIdx, ut);
  invalidateCache(SHEETS.UNTAGGED_CHECKLISTS);

  // Re-process inventory: reverse prior, then re-apply
  var nqForInv = (ck && ck.questions) || [];
  var hasInvLinks = nqForInv.some(function(q) { return q.inventoryLink && q.inventoryLink.enabled; });
  var editPerson = newPerson || ut.person;

  if (hasInvLinks) {
    reverseInventoryLedgerForRef("untagged", id, editPerson);
    reverseInventoryLedgerForRef("checklist", id, editPerson);
    processInventoryLinks(ck, responsesMap, "untagged", id, editPerson, true);
  } else {
    reverseInventoryLedgerForRef("untagged", id, editPerson);
    reverseInventoryLedgerForRef("checklist", id, editPerson);
    var respMapByText = {};
    for (var ri = 0; ri < newResponses.length; ri++) {
      respMapByText[newResponses[ri].questionText] = newResponses[ri].response || "";
    }
    applyLegacyInventoryForChecklist(ck, respMapByText, "checklist", id, editPerson, body.inventoryItemId || "", body.inventoryOutputItemId || "", true, body.grindClassificationId || "");
  }

  // Update per-checklist response tab if exists
  var orderId = ut.tagged_order_id || "UNTAGGED";
  deleteResponseRow(ck.name, orderId);
  var respArray = newResponses.map(function(r) { return (r.response !== undefined && r.response !== null) ? String(r.response) : ""; });
  var orderName = "", customerLabel = "";
  if (ut.tagged_order_id) {
    var order = lookupOrder(ut.tagged_order_id);
    if (order) { orderName = order.name; customerLabel = lookupCustomerLabel(order.customer_id); }
  }
  writeResponseRow(ck.name, ck.questions, {
    orderId: orderId, orderName: orderName, customer: customerLabel,
    person: ut.person, date: ut.date, submittedAt: ut.submitted_at, responses: respArray, remarks: remarks,
  });

  var afterState = JSON.stringify({ person: ut.person, date: ut.date, totalQuantity: newTotalQty });
  writeAuditLog(user, "edit", "UntaggedChecklist", id, "before=" + beforeState + " after=" + afterState);
  return { success: true };
}

// ═══════════════════════════════════════════════════════════════
// ─── PHASE 9: Abbreviation Uniqueness ───────────────────────
// ═══════════════════════════════════════════════════════════════

// Validates that an abbreviation is unique within its category. Called during
// create and update. Returns null if OK, or an error string if duplicate.
function checkAbbreviationUniqueness(abbreviation, category, excludeItemId) {
  if (!abbreviation) return null;
  var ab = String(abbreviation).toUpperCase();
  var items = getRows(SHEETS.INVENTORY_ITEMS);
  for (var i = 0; i < items.length; i++) {
    if (String(items[i].id) === String(excludeItemId)) continue;
    if (String(items[i].abbreviation || "").toUpperCase() === ab) {
      return "Abbreviation '" + ab + "' already exists in inventory (used by: " + items[i].name + " in " + items[i].category + "). Abbreviations must be unique across all categories.";
    }
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════
// ─── Automated Test Suite ────────────────────────────────────
// ═══════════════════════════════════════════════════════════════

function handleRunTests(user) {
  var testUser = { id: user.id, username: user.username, displayName: "TEST_RUNNER", role: "admin" };
  var results = [];
  var testIds = { items: [], untagged: [], ledger: [], allocations: [], classifications: [], auditStart: new Date().toISOString() };

  function check(desc, actual, expected) {
    var pass = actual === expected;
    return { check: desc, result: pass ? "PASS" : "FAIL", detail: pass ? String(actual) : "Expected " + JSON.stringify(expected) + ", got " + JSON.stringify(actual) };
  }
  function checkTruthy(desc, val) {
    return { check: desc, result: val ? "PASS" : "FAIL", detail: String(val || "(falsy)") };
  }
  function checkGte(desc, val, min) {
    var pass = val >= min;
    return { check: desc, result: pass ? "PASS" : "FAIL", detail: val + (pass ? " >= " : " < ") + min };
  }

  // ── TEST 1: Inventory Items ──
  var t1Checks = [];
  try {
    var gbItem = handleCreateInventoryItem({ name: "TEST_GB_ITEM", abbreviation: "TGBI", category: "Green Beans", unit: "kg", openingStock: 0, equivalentItems: [] }, testUser);
    testIds.items.push(gbItem.id);
    var rbItem = handleCreateInventoryItem({ name: "TEST_RB_ITEM", abbreviation: "TRBI", category: "Roasted Beans", unit: "kg", openingStock: 0, equivalentItems: [] }, testUser);
    testIds.items.push(rbItem.id);
    handleUpdateInventoryItem({ id: gbItem.id, equivalentItems: [{ category: "Roasted Beans", itemId: rbItem.id }] }, testUser);
    handleUpdateInventoryItem({ id: rbItem.id, equivalentItems: [{ category: "Green Beans", itemId: gbItem.id }] }, testUser);
    invalidateCache(SHEETS.INVENTORY_ITEMS);
    var items = getRows(SHEETS.INVENTORY_ITEMS);
    var foundGb = items.find(function(r) { return r.id === gbItem.id; });
    var foundRb = items.find(function(r) { return r.id === rbItem.id; });
    t1Checks.push(checkTruthy("Green Bean item created", !!foundGb));
    t1Checks.push(checkTruthy("Roasted Bean item created", !!foundRb));
    var gbEq = safeParseJSON(foundGb ? foundGb.equivalent_items : "[]", []);
    t1Checks.push(check("GB equivalent links to RB", gbEq.length > 0 && gbEq[0].itemId === rbItem.id, true));
  } catch (e) { t1Checks.push({ check: "Inventory creation", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 1, testName: "Inventory Items", status: t1Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t1Checks });

  // ── TEST 2: Green Bean QC Sample Check ──
  var t2Checks = [], sampleAutoId = "";
  try {
    var sampleCk = lookupChecklist("ck_sample_qc");
    var sampleNq = sampleCk ? sampleCk.questions : [];
    var sampleResp = sampleNq.map(function(q, qi) {
      if (q.text === "Type of Beans") return { questionIndex: qi, questionText: q.text, response: gbItem.id };
      if (q.text === "Sample Quantity") return { questionIndex: qi, questionText: q.text, response: "1" };
      if (q.text === "Sample Approved?") return { questionIndex: qi, questionText: q.text, response: "Yes" };
      return { questionIndex: qi, questionText: q.text, response: "TEST_VALUE" };
    });
    var s2 = handleSubmitUntagged({ checklistId: "ck_sample_qc", date: "2026-01-01", person: "TEST_RUNNER", responses: sampleResp, remarks: {}, orderId: "" }, testUser);
    testIds.untagged.push(s2.id);
    sampleAutoId = s2.autoId || "";
    invalidateCache(SHEETS.UNTAGGED_CHECKLISTS);
    var utRow = getRows(SHEETS.UNTAGGED_CHECKLISTS).find(function(r) { return r.id === s2.id; });
    t2Checks.push(checkTruthy("Entry created in UntaggedChecklists", !!utRow));
    t2Checks.push(checkTruthy("Auto ID starts with GBS-", sampleAutoId.indexOf("GBS-") === 0));
    t2Checks.push(check("total_quantity = 1", parseFloat(utRow ? utRow.total_quantity : 0), 1));
  } catch (e) { t2Checks.push({ check: "Sample QC submission", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 2, testName: "Green Bean QC Sample Check", status: t2Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t2Checks });

  // ── TEST 3: Green Beans Quality Check ──
  var t3Checks = [], gbAutoId = "", gbUtId = "";
  try {
    var gbCk = lookupChecklist("ck_green_beans");
    var gbNq = gbCk ? gbCk.questions : [];
    var gbResp = gbNq.map(function(q, qi) {
      if (q.text === "Source Sample") return { questionIndex: qi, questionText: q.text, response: sampleAutoId };
      if (q.text === "Type of Beans") return { questionIndex: qi, questionText: q.text, response: gbItem.id };
      if (q.text === "Quantity received") return { questionIndex: qi, questionText: q.text, response: "100" };
      if (q.text === "Shipment Approved?") return { questionIndex: qi, questionText: q.text, response: "Yes" };
      return { questionIndex: qi, questionText: q.text, response: "TEST_VALUE" };
    });
    var s3 = handleSubmitUntagged({ checklistId: "ck_green_beans", date: "2026-01-02", person: "TEST_RUNNER", responses: gbResp, remarks: {}, orderId: "", batchAllocations: {} }, testUser);
    gbUtId = s3.id;
    testIds.untagged.push(s3.id);
    gbAutoId = s3.autoId || "";
    invalidateCache(SHEETS.UNTAGGED_CHECKLISTS);
    invalidateCache(SHEETS.INVENTORY_LEDGER);
    var utRow3 = getRows(SHEETS.UNTAGGED_CHECKLISTS).find(function(r) { return r.id === s3.id; });
    t3Checks.push(checkTruthy("Entry created", !!utRow3));
    t3Checks.push(checkTruthy("Auto ID starts with GB-", gbAutoId.indexOf("GB-") === 0));
    t3Checks.push(check("total_quantity = 100", parseFloat(utRow3 ? utRow3.total_quantity : 0), 100));
    var ledger3 = getRows(SHEETS.INVENTORY_LEDGER).filter(function(r) { return String(r.reference_id) === String(s3.id) && r.type === "IN"; });
    t3Checks.push(checkTruthy("Ledger IN entry exists", ledger3.length > 0));
    if (ledger3.length > 0) {
      t3Checks.push(check("Ledger IN qty = 100", parseFloat(ledger3[0].quantity), 100));
      t3Checks.push(check("Ledger category = Green Beans", String(ledger3[0].category || "").trim(), "Green Beans"));
    }
  } catch (e) { t3Checks.push({ check: "GB QC submission", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 3, testName: "Green Beans Quality Check", status: t3Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t3Checks });

  // ── TEST 4: Roasted Beans QC (multi-batch) ──
  var t4Checks = [], rbAutoId = "", rbUtId = "";
  try {
    var rbCk = lookupChecklist("ck_roasted_beans");
    var rbNq = rbCk ? rbCk.questions : [];
    var rbResp = rbNq.map(function(q, qi) {
      if (q.text === "Date of Roast") return { questionIndex: qi, questionText: q.text, response: "2026-01-03" };
      if (q.text === "Roast profile") return { questionIndex: qi, questionText: q.text, response: "Medium" };
      if (q.text === "Roast Approved?") return { questionIndex: qi, questionText: q.text, response: "Yes" };
      return { questionIndex: qi, questionText: q.text, response: "" };
    });
    var s4 = handleSubmitUntagged({
      checklistId: "ck_roasted_beans", date: "2026-01-03", person: "TEST_RUNNER",
      responses: rbResp, remarks: {}, orderId: "",
      roast_batches: [{ sourceAutoId: gbAutoId, inputQty: 40, outputQty: 35, reasonForLoss: "Moisture loss", classificationId: "" }],
    }, testUser);
    rbUtId = s4.id;
    testIds.untagged.push(s4.id);
    rbAutoId = s4.autoId || "";
    invalidateCache(SHEETS.UNTAGGED_CHECKLISTS);
    invalidateCache(SHEETS.INVENTORY_LEDGER);
    invalidateCache(SHEETS.QUANTITY_ALLOCATIONS);
    var utRow4 = getRows(SHEETS.UNTAGGED_CHECKLISTS).find(function(r) { return r.id === s4.id; });
    t4Checks.push(checkTruthy("Entry created", !!utRow4));
    t4Checks.push(checkTruthy("Auto ID starts with RB-", rbAutoId.indexOf("RB-") === 0));
    t4Checks.push(check("total_quantity = 35", parseFloat(utRow4 ? utRow4.total_quantity : 0), 35));
    // Check roast_batches stored in responses
    var storedResp = safeParseJSON(utRow4 ? utRow4.responses : "[]", []);
    var shipField = storedResp.find(function(r) { return r.questionText === "Shipment number used"; });
    t4Checks.push(checkTruthy("roast_batches JSON in responses", shipField && String(shipField.response || "").indexOf("[") === 0));
    // Check ledger
    var ledger4 = getRows(SHEETS.INVENTORY_LEDGER).filter(function(r) { return String(r.reference_id) === String(s4.id); });
    var outEntries = ledger4.filter(function(r) { return r.type === "OUT"; });
    var inEntries = ledger4.filter(function(r) { return r.type === "IN"; });
    t4Checks.push(checkTruthy("Ledger OUT entry (Green Beans)", outEntries.length > 0));
    t4Checks.push(checkTruthy("Ledger IN entry (Roasted Beans)", inEntries.length > 0));
    if (outEntries.length > 0) t4Checks.push(check("OUT qty = -40", parseFloat(outEntries[0].quantity), -40));
    if (inEntries.length > 0) t4Checks.push(check("IN qty = 35", parseFloat(inEntries[0].quantity), 35));
    // Check GB remaining
    var gbAllocated = getAllocatedQuantityForAutoId(gbAutoId);
    t4Checks.push(check("GB allocated = 40", gbAllocated, 40));
    t4Checks.push(check("GB remaining = 60", 100 - gbAllocated, 60));
  } catch (e) { t4Checks.push({ check: "RB QC multi-batch", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 4, testName: "Roasted Beans QC (multi-batch)", status: t4Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t4Checks });

  // ── TEST 5: Quantity Validation ──
  var t5Checks = [], rb2UtId = "";
  try {
    var rbResp5 = rbNq.map(function(q, qi) {
      if (q.text === "Roast Approved?") return { questionIndex: qi, questionText: q.text, response: "Yes" };
      return { questionIndex: qi, questionText: q.text, response: "" };
    });
    // Attempt over-allocation (70 > 60 remaining)
    var s5fail = handleSubmitUntagged({
      checklistId: "ck_roasted_beans", date: "2026-01-04", person: "TEST_RUNNER",
      responses: rbResp5, remarks: {}, orderId: "",
      roast_batches: [{ sourceAutoId: gbAutoId, inputQty: 70, outputQty: 60, reasonForLoss: "test", classificationId: "" }],
    }, testUser);
    t5Checks.push(checkTruthy("Over-allocation rejected", !!s5fail.error));
    // Exact remaining (60)
    var s5ok = handleSubmitUntagged({
      checklistId: "ck_roasted_beans", date: "2026-01-04", person: "TEST_RUNNER",
      responses: rbResp5, remarks: {}, orderId: "",
      roast_batches: [{ sourceAutoId: gbAutoId, inputQty: 60, outputQty: 55, reasonForLoss: "test", classificationId: "" }],
    }, testUser);
    t5Checks.push(checkTruthy("Exact allocation accepted", !!s5ok.id));
    if (s5ok.id) { rb2UtId = s5ok.id; testIds.untagged.push(s5ok.id); }
    invalidateCache(SHEETS.QUANTITY_ALLOCATIONS);
    var gbAlloc5 = getAllocatedQuantityForAutoId(gbAutoId);
    t5Checks.push(check("GB fully allocated = 100", gbAlloc5, 100));
  } catch (e) { t5Checks.push({ check: "Quantity validation", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 5, testName: "Quantity Validation", status: t5Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t5Checks });

  // ── TEST 6: Grinding & Packing ──
  var t6Checks = [], grindUtId = "";
  try {
    var grCk = lookupChecklist("ck_grinding");
    var grNq = grCk ? grCk.questions : [];
    var grResp = grNq.map(function(q, qi) {
      if (q.text === "Roast ID") return { questionIndex: qi, questionText: q.text, response: rbAutoId };
      if (q.text === "Total Net weight") return { questionIndex: qi, questionText: q.text, response: "28" };
      return { questionIndex: qi, questionText: q.text, response: "TEST_VALUE" };
    });
    var s6 = handleSubmitUntagged({
      checklistId: "ck_grinding", date: "2026-01-05", person: "TEST_RUNNER",
      responses: grResp, remarks: {}, orderId: "",
      batchAllocations: { "0": [{ sourceAutoId: rbAutoId, quantity: 30 }] },
    }, testUser);
    grindUtId = s6.id;
    testIds.untagged.push(s6.id);
    invalidateCache(SHEETS.UNTAGGED_CHECKLISTS);
    invalidateCache(SHEETS.INVENTORY_LEDGER);
    var utRow6 = getRows(SHEETS.UNTAGGED_CHECKLISTS).find(function(r) { return r.id === s6.id; });
    t6Checks.push(checkTruthy("Entry created", !!utRow6));
    var ledger6 = getRows(SHEETS.INVENTORY_LEDGER).filter(function(r) { return String(r.reference_id) === String(s6.id); });
    t6Checks.push(checkGte("Ledger entries created", ledger6.length, 1));
  } catch (e) { t6Checks.push({ check: "Grinding submission", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 6, testName: "Grinding & Packing", status: t6Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t6Checks });

  // ── TEST 7: Soft Delete ──
  var t7Checks = [];
  try {
    var del7 = handleSoftDeleteChecklist({ id: grindUtId, entityType: "untagged", reason: "Automated test cleanup" }, testUser);
    t7Checks.push(checkTruthy("Delete succeeded", !!del7.success));
    invalidateCache(SHEETS.UNTAGGED_CHECKLISTS);
    invalidateCache(SHEETS.INVENTORY_LEDGER);
    var utRow7 = getRows(SHEETS.UNTAGGED_CHECKLISTS).find(function(r) { return r.id === grindUtId; });
    t7Checks.push(check("is_deleted = true", isDeleted(utRow7), true));
    // Check it doesn't appear in approved entries
    var approved7 = getApprovedEntriesForChecklist("ck_grinding");
    var found7 = approved7.find(function(e) { return e.autoId === (utRow7 ? utRow7.auto_id : ""); });
    t7Checks.push(check("Not in approved entries", !!found7, false));
    // Reversal entries
    var reversals = getRows(SHEETS.INVENTORY_LEDGER).filter(function(r) { return String(r.reference_id) === String(grindUtId) && String(r.notes || "").indexOf("[REVERSAL]") === 0; });
    t7Checks.push(checkGte("Reversal ledger entries exist", reversals.length, 1));
  } catch (e) { t7Checks.push({ check: "Soft delete", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 7, testName: "Soft Delete", status: t7Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t7Checks });

  // ── TEST 8: Edit Validation ──
  var t8Checks = [];
  try {
    // Try reducing GB qty below committed (should fail)
    var edit8fail = handleEditUntaggedResponse({
      id: gbUtId, person: "TEST_RUNNER", date: "2026-01-02",
      responses: gbNq.map(function(q, qi) {
        if (q.text === "Quantity received") return { questionIndex: qi, questionText: q.text, response: "30" };
        if (q.text === "Type of Beans") return { questionIndex: qi, questionText: q.text, response: gbItem.id };
        if (q.text === "Shipment Approved?") return { questionIndex: qi, questionText: q.text, response: "Yes" };
        return { questionIndex: qi, questionText: q.text, response: "TEST_VALUE" };
      }), remarks: {},
    }, testUser);
    t8Checks.push(checkTruthy("Reduce below committed rejected", !!edit8fail.error));
    // Increase to 150 (should succeed)
    var edit8ok = handleEditUntaggedResponse({
      id: gbUtId, person: "TEST_RUNNER", date: "2026-01-02",
      responses: gbNq.map(function(q, qi) {
        if (q.text === "Quantity received") return { questionIndex: qi, questionText: q.text, response: "150" };
        if (q.text === "Type of Beans") return { questionIndex: qi, questionText: q.text, response: gbItem.id };
        if (q.text === "Shipment Approved?") return { questionIndex: qi, questionText: q.text, response: "Yes" };
        return { questionIndex: qi, questionText: q.text, response: "TEST_VALUE" };
      }), remarks: {},
    }, testUser);
    t8Checks.push(checkTruthy("Increase succeeded", !!edit8ok.success));
    invalidateCache(SHEETS.UNTAGGED_CHECKLISTS);
    var utRow8 = getRows(SHEETS.UNTAGGED_CHECKLISTS).find(function(r) { return r.id === gbUtId; });
    t8Checks.push(check("total_quantity updated to 150", parseFloat(utRow8 ? utRow8.total_quantity : 0), 150));
  } catch (e) { t8Checks.push({ check: "Edit validation", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 8, testName: "Edit Validation", status: t8Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t8Checks });

  // ── TEST 9: Dropdown Remaining Quantity ──
  var t9Checks = [];
  try {
    var linked9 = handleGetLinkedEntries({ checklist_id: "ck_roasted_beans" });
    var gbEntry9 = Array.isArray(linked9) ? linked9.find(function(e) { return e.autoId === gbAutoId; }) : null;
    t9Checks.push(checkTruthy("GB entry found in linked entries", !!gbEntry9));
    if (gbEntry9) {
      t9Checks.push(check("totalQuantity = 150", gbEntry9.totalQuantity, 150));
      t9Checks.push(check("remainingQuantity = 50", gbEntry9.remainingQuantity, 50));
    }
  } catch (e) { t9Checks.push({ check: "Dropdown quantity", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 9, testName: "Dropdown Remaining Quantity", status: t9Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t9Checks });

  // ── TEST 10: Audit Log ──
  var t10Checks = [];
  try {
    var auditRows = getRows(SHEETS.AUDIT_LOG).filter(function(r) { return String(r.timestamp) >= testIds.auditStart && String(r.user_name) === "TEST_RUNNER"; });
    var hasTag = auditRows.some(function(r) { return r.action === "tag"; });
    var hasDelete = auditRows.some(function(r) { return r.action === "delete"; });
    var hasEdit = auditRows.some(function(r) { return r.action === "edit"; });
    t10Checks.push(checkTruthy("Audit log has tag entries", hasTag));
    t10Checks.push(checkTruthy("Audit log has delete entries", hasDelete));
    t10Checks.push(checkTruthy("Audit log has edit entries", hasEdit));
    t10Checks.push(checkGte("Total audit entries from test run", auditRows.length, 3));
  } catch (e) { t10Checks.push({ check: "Audit log", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 10, testName: "Audit Log", status: t10Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t10Checks });

  // ── TEST 11: Classifications ──
  var t11Checks = [], testClassId = "";
  try {
    var cls = handleAddClassification({ name: "TEST_CLASS", type: "roast_degree", description: "Test classification" }, testUser);
    testClassId = cls.id || "";
    testIds.classifications.push(testClassId);
    invalidateCache(SHEETS.ROAST_CLASSIFICATIONS);
    var allCls = handleGetClassifications();
    var found11 = (allCls.roast_degree || []).find(function(c) { return c.id === testClassId; });
    t11Checks.push(checkTruthy("Classification created", !!found11));
    handleDeactivateClassification({ id: testClassId }, testUser);
    invalidateCache(SHEETS.ROAST_CLASSIFICATIONS);
    var allCls2 = handleGetClassifications();
    var found11b = (allCls2.roast_degree || []).find(function(c) { return c.id === testClassId; });
    t11Checks.push(check("Not in active after deactivation", !!found11b, false));
  } catch (e) { t11Checks.push({ check: "Classifications", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 11, testName: "Classifications", status: t11Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t11Checks });

  // ── TEST 12: is_deleted filtering ──
  var t12Checks = [];
  try {
    var utRow12 = getRows(SHEETS.UNTAGGED_CHECKLISTS).find(function(r) { return r.id === grindUtId; });
    var grindAutoId = utRow12 ? String(utRow12.auto_id) : "";
    if (grindAutoId) {
      var sub12 = getSubmissionByAutoId(grindAutoId, lookupChecklist("ck_grinding"));
      t12Checks.push(check("Deleted not in getSubmissionByAutoId", sub12, null));
    } else {
      t12Checks.push({ check: "getSubmissionByAutoId", result: "FAIL", detail: "No auto_id found for grinding entry" });
    }
    var linked12 = handleGetLinkedEntries({ checklist_id: "ck_grinding" });
    var grindInLinked = Array.isArray(linked12) ? linked12.find(function(e) { return e.autoId === grindAutoId; }) : null;
    t12Checks.push(check("Deleted not in getLinkedEntries", !!grindInLinked, false));
  } catch (e) { t12Checks.push({ check: "is_deleted filtering", result: "FAIL", detail: e.message }); }
  results.push({ testNumber: 12, testName: "is_deleted Filtering", status: t12Checks.every(function(c) { return c.result === "PASS"; }) ? "PASS" : "FAIL", checks: t12Checks });

  // ── CLEANUP ──
  var cleanupDeleted = 0;
  var sheetsAffected = [];
  try {
    // Delete test untagged entries
    var utSheet = getSheet(SHEETS.UNTAGGED_CHECKLISTS);
    if (utSheet) {
      var utData = utSheet.getDataRange().getValues();
      for (var u = utData.length - 1; u >= 1; u--) {
        var utPerson = String(utData[u][3] || "");
        if (utPerson === "TEST_RUNNER") { utSheet.deleteRow(u + 1); cleanupDeleted++; }
      }
      if (cleanupDeleted > 0) sheetsAffected.push("UntaggedChecklists");
    }
    invalidateCache(SHEETS.UNTAGGED_CHECKLISTS);
    // Delete test ledger entries
    var ledSheet = getSheet(SHEETS.INVENTORY_LEDGER);
    if (ledSheet) {
      var ledData = ledSheet.getDataRange().getValues();
      var ledHeaders = ledData[0].map(String);
      var refCol = ledHeaders.indexOf("reference_id");
      var doneByCol = ledHeaders.indexOf("done_by");
      var ledDel = 0;
      for (var l = ledData.length - 1; l >= 1; l--) {
        if (String(ledData[l][doneByCol] || "") === "TEST_RUNNER") { ledSheet.deleteRow(l + 1); ledDel++; }
      }
      cleanupDeleted += ledDel;
      if (ledDel > 0) sheetsAffected.push("InventoryLedger");
    }
    invalidateCache(SHEETS.INVENTORY_LEDGER);
    // Delete test inventory items
    var invSheet = getSheet(SHEETS.INVENTORY_ITEMS);
    if (invSheet) {
      var invData = invSheet.getDataRange().getValues();
      var invDel = 0;
      for (var iv = invData.length - 1; iv >= 1; iv--) {
        if (String(invData[iv][2] || "").indexOf("TEST_") === 0) { invSheet.deleteRow(iv + 1); invDel++; }
      }
      cleanupDeleted += invDel;
      if (invDel > 0) sheetsAffected.push("InventoryItems");
    }
    invalidateCache(SHEETS.INVENTORY_ITEMS);
    // Delete test quantity allocations
    var qaSheet = getSheet(SHEETS.QUANTITY_ALLOCATIONS);
    if (qaSheet) {
      var qaData = qaSheet.getDataRange().getValues();
      var qaDel = 0;
      for (var qa = qaData.length - 1; qa >= 1; qa--) {
        if (String(qaData[qa][8] || "") === "TEST_RUNNER") { qaSheet.deleteRow(qa + 1); qaDel++; }
      }
      cleanupDeleted += qaDel;
      if (qaDel > 0) sheetsAffected.push("QuantityAllocations");
    }
    invalidateCache(SHEETS.QUANTITY_ALLOCATIONS);
    // Delete test audit log entries
    var auditSheet = getSheet(SHEETS.AUDIT_LOG);
    if (auditSheet) {
      var auditData = auditSheet.getDataRange().getValues();
      var auditDel = 0;
      for (var a = auditData.length - 1; a >= 1; a--) {
        if (String(auditData[a][2] || "") === "TEST_RUNNER") { auditSheet.deleteRow(a + 1); auditDel++; }
      }
      cleanupDeleted += auditDel;
      if (auditDel > 0) sheetsAffected.push("AuditLog");
    }
    invalidateCache(SHEETS.AUDIT_LOG);
    // Delete test classifications
    var clsSheet = getSheet(SHEETS.ROAST_CLASSIFICATIONS);
    if (clsSheet) {
      var clsData = clsSheet.getDataRange().getValues();
      var clsDel = 0;
      for (var cl = clsData.length - 1; cl >= 1; cl--) {
        if (String(clsData[cl][1] || "").indexOf("TEST_") === 0) { clsSheet.deleteRow(cl + 1); clsDel++; }
      }
      cleanupDeleted += clsDel;
      if (clsDel > 0) sheetsAffected.push("RoastClassifications");
    }
    invalidateCache(SHEETS.ROAST_CLASSIFICATIONS);
    // Delete test response rows from per-checklist tabs
    var testCkNames = ["Green Bean QC Sample Check", "Green Beans Quality Check", "Roasted Beans Quality Check", "Grinding & Packing Checklist"];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    for (var tn = 0; tn < testCkNames.length; tn++) {
      var rSheet = ss.getSheetByName(testCkNames[tn]);
      if (rSheet && rSheet.getLastRow() > 2) {
        var rData = rSheet.getDataRange().getValues();
        for (var rr = rData.length - 1; rr >= 2; rr--) {
          if (String(rData[rr][3] || "") === "TEST_RUNNER") { rSheet.deleteRow(rr + 1); cleanupDeleted++; }
        }
      }
    }
  } catch (e) { Logger.log("Test cleanup error: " + e.message); }

  var passed = results.filter(function(r) { return r.status === "PASS"; }).length;
  var failed = results.filter(function(r) { return r.status === "FAIL"; }).length;
  return {
    totalTests: results.length,
    passed: passed,
    failed: failed,
    results: results,
    cleanupSummary: { rowsDeleted: cleanupDeleted, sheetsAffected: sheetsAffected },
  };
}

// ─── One-time corrective: deduplicate ledger entries and recompute current_stock ──
//
// Confirmed duplicate pattern: for each submission, two ledger rows share the SAME
// reference_id but different reference_type:
//   • reference_type="untagged" — written by the legacy hard-coded tracking path
//   • reference_type="checklist" — written by processInventoryLinks()
//
// The "checklist"-typed rows are the duplicates to remove (the legacy "untagged" rows
// were already handling inventory before processInventoryLinks started firing).
//
// Algorithm:
//   1. Load all ledger rows.
//   2. Group by reference_id.
//   3. For each group: if BOTH an "untagged" and a "checklist" row exist for the same
//      (item_id, direction), mark the "checklist" row for deletion.
//   4. Delete flagged rows from the sheet bottom-up (so row indices above stay valid).
//   5. Recompute current_stock per affected item chronologically from the surviving rows
//      (opening_stock + Σ signed quantities).
//   6. Log before/after for each item.
//
// Run manually from the Apps Script editor — NOT auto-run.
function fixDuplicateInventoryEntries() {
  ensureSheetHasAllColumns(SHEETS.INVENTORY_LEDGER);
  ensureSheetHasAllColumns(SHEETS.INVENTORY_ITEMS);
  clearRowsCache();

  var ledgerSheet = getSheet(SHEETS.INVENTORY_LEDGER);
  if (!ledgerSheet) {
    Logger.log("fixDuplicateInventoryEntries: InventoryLedger sheet missing");
    return { error: "InventoryLedger sheet missing" };
  }
  var data = ledgerSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("fixDuplicateInventoryEntries: ledger is empty");
    return { ledgerRows: 0, duplicatesRemoved: 0, itemsFixed: 0 };
  }
  var headers = data[0].map(String);
  var colIndex = {};
  for (var h = 0; h < headers.length; h++) colIndex[headers[h]] = h;
  var required = ["id", "item_id", "type", "quantity", "reference_type", "reference_id", "created_at"];
  for (var rq = 0; rq < required.length; rq++) {
    if (colIndex[required[rq]] === undefined) {
      Logger.log("fixDuplicateInventoryEntries: missing required column: " + required[rq]);
      return { error: "Missing column: " + required[rq] };
    }
  }

  // Build parallel arrays: rows (object form) + their sheet row number (1-based, incl. header).
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[headers[j]] = data[i][j];
    obj.__rowIndex = i + 1; // sheet row number
    rows.push(obj);
  }
  Logger.log("fixDuplicateInventoryEntries: scanning " + rows.length + " ledger rows");

  // ── Step 1+2+3: group by reference_id, flag "checklist" duplicates when a sibling
  //               "untagged" row exists for the same (item_id, direction).
  var byRef = {};
  for (var r1 = 0; r1 < rows.length; r1++) {
    var rr = rows[r1];
    var refId = String(rr.reference_id || "");
    if (!refId) continue;
    if (!byRef[refId]) byRef[refId] = [];
    byRef[refId].push(rr);
  }

  var toDelete = []; // sheet row numbers to delete
  var dupIds = {};   // ledger ids flagged duplicate
  Object.keys(byRef).forEach(function(refId) {
    var group = byRef[refId];
    // Split by reference_type
    var untagged = [], checklist = [];
    for (var g = 0; g < group.length; g++) {
      var rt = String(group[g].reference_type || "").toLowerCase();
      if (rt === "untagged") untagged.push(group[g]);
      else if (rt === "checklist") checklist.push(group[g]);
    }
    if (untagged.length === 0 || checklist.length === 0) return;

    // For each checklist row, look for a matching untagged row with same item_id + same
    // direction (IN/OUT). If found, flag the checklist row as duplicate.
    for (var ci = 0; ci < checklist.length; ci++) {
      var cRow = checklist[ci];
      var cItem = String(cRow.item_id || "");
      var cDir = String(cRow.type || "").toUpperCase();
      if (!cItem || !cDir) continue;
      var matched = false;
      for (var ui = 0; ui < untagged.length; ui++) {
        var uRow = untagged[ui];
        if (String(uRow.item_id || "") !== cItem) continue;
        if (String(uRow.type || "").toUpperCase() !== cDir) continue;
        matched = true; break;
      }
      if (matched) {
        dupIds[String(cRow.id)] = true;
        toDelete.push(cRow.__rowIndex);
        Logger.log("  DUPLICATE flagged: refId=" + refId + " item=" + cItem + " dir=" + cDir + " ledgerId=" + cRow.id + " (reference_type=checklist)");
      }
    }
  });

  // ── Step 4: delete flagged rows from the bottom up so indices above remain stable.
  toDelete.sort(function(a, b) { return b - a; });
  var duplicatesRemoved = 0;
  for (var d = 0; d < toDelete.length; d++) {
    ledgerSheet.deleteRow(toDelete[d]);
    duplicatesRemoved++;
  }
  invalidateCache(SHEETS.INVENTORY_LEDGER);
  Logger.log("fixDuplicateInventoryEntries: deleted " + duplicatesRemoved + " duplicate row(s)");

  // ── Step 5: recompute current_stock per item from the surviving ledger, in chronological
  //           order (by created_at asc; fall back to row order for ties).
  var surviving = rows.filter(function(r) { return !dupIds[String(r.id)]; });
  surviving.sort(function(a, b) {
    var ta = Date.parse(String(a.created_at || "")) || 0;
    var tb = Date.parse(String(b.created_at || "")) || 0;
    if (ta !== tb) return ta - tb;
    return (a.__rowIndex || 0) - (b.__rowIndex || 0);
  });

  var perItemDelta = {};
  for (var s = 0; s < surviving.length; s++) {
    var sr = surviving[s];
    var itemId = String(sr.item_id || "");
    if (!itemId) continue;
    var txType = String(sr.type || "").toUpperCase();
    var q = parseFloat(sr.quantity) || 0;
    // Normalize sign from type so legacy rows with inconsistent signs are handled correctly.
    var signed;
    if (txType === "IN") signed = Math.abs(q);
    else if (txType === "OUT") signed = -Math.abs(q);
    else if (txType === "ADJUSTMENT") signed = q; // sign preserved on adjustments
    else signed = q;
    perItemDelta[itemId] = (perItemDelta[itemId] || 0) + signed;
  }

  var items = getRows(SHEETS.INVENTORY_ITEMS);
  var itemsFixed = 0;
  for (var ii = 0; ii < items.length; ii++) {
    var item = items[ii];
    var opening = parseFloat(item.opening_stock) || 0;
    var before = parseFloat(item.current_stock) || 0;
    var delta = perItemDelta[String(item.id)] || 0;
    var after = opening + delta;
    var changed = Math.abs(before - after) > 0.0001;
    Logger.log("item " + item.id + " (" + item.name + "): opening=" + opening + ", current(before)=" + before + ", delta=" + delta + ", current(after)=" + after + (changed ? "  ← CHANGED" : ""));
    if (changed) {
      var idx = findRowIndex(SHEETS.INVENTORY_ITEMS, item.id);
      if (idx > 0) {
        item.current_stock = after;
        updateSheetRow(SHEETS.INVENTORY_ITEMS, idx, item);
        itemsFixed++;
      }
    }
  }
  invalidateCache(SHEETS.INVENTORY_ITEMS);
  Logger.log("fixDuplicateInventoryEntries: completed. duplicatesRemoved=" + duplicatesRemoved + ", itemsFixed=" + itemsFixed);
  return {
    ledgerRows: rows.length,
    duplicatesRemoved: duplicatesRemoved,
    survivingRows: surviving.length,
    itemsFixed: itemsFixed,
  };
}

// ─── One-time corrective: redirect delivery deductions from Green Beans to Roasted Beans ──
//
// Prior versions of handleDeliverOrder() deducted from the IN-tx (Green Beans) item instead
// of the OUT-tx (Roasted Beans / Packing Items) item. This walks every ledger entry with
// reference_type = "order_delivery" and, for any entry whose item is in the Green Beans
// category, writes:
//   • a +qty compensating entry (reversal) against the Green Beans item
//   • a -qty entry against the equivalent Roasted Beans item (or Packing Items)
// Both entries also update current_stock on InventoryItems.
//
// Entries whose notes already contain "[DELIVERY-CORRECTION]" are skipped to keep the
// function idempotent across multiple runs. Reversal rows (notes starting with "[REVERSAL]")
// are also skipped.
//
// Run manually from the Apps Script editor — NOT auto-run.
function fixDeliveryLedgerEntries() {
  ensureSheetHasAllColumns(SHEETS.INVENTORY_LEDGER);
  ensureSheetHasAllColumns(SHEETS.INVENTORY_ITEMS);
  clearRowsCache();

  var ledger = getRows(SHEETS.INVENTORY_LEDGER);
  var items = getRows(SHEETS.INVENTORY_ITEMS);
  var itemById = {};
  for (var i = 0; i < items.length; i++) itemById[String(items[i].id)] = items[i];

  Logger.log("fixDeliveryLedgerEntries: scanning " + ledger.length + " ledger rows");

  // Build a set of ledger ids that have already been corrected (to avoid double correcting
  // across multiple runs). An entry is considered already-corrected if a "[DELIVERY-CORRECTION]"
  // row exists whose notes reference its ledger id.
  var alreadyCorrectedOriginId = {};
  for (var lc = 0; lc < ledger.length; lc++) {
    var notes = String(ledger[lc].notes || "");
    if (notes.indexOf("[DELIVERY-CORRECTION]") !== 0) continue;
    var m = /of led:([^\s—-]+)/.exec(notes);
    if (m && m[1]) alreadyCorrectedOriginId[m[1]] = true;
  }

  var correctionsMade = 0;
  var skippedNoEquivalent = 0;
  var skippedAlreadyCorrected = 0;

  for (var r = 0; r < ledger.length; r++) {
    var row = ledger[r];
    if (String(row.reference_type || "") !== "order_delivery") continue;
    var rowNotes = String(row.notes || "");
    if (rowNotes.indexOf("[REVERSAL]") === 0) continue;
    if (rowNotes.indexOf("[DELIVERY-CORRECTION]") === 0) continue;
    if (alreadyCorrectedOriginId[String(row.id)]) { skippedAlreadyCorrected++; continue; }

    var origItem = itemById[String(row.item_id || "")];
    if (!origItem) continue;
    var origCat = String(origItem.category || "").toLowerCase();
    if (origCat !== "green beans") continue;

    var qty = Math.abs(parseFloat(row.quantity) || 0);
    if (qty <= 0) continue;

    // Find equivalent in Roasted Beans (preferred) or Packing Items.
    var equiv = safeParseJSON(origItem.equivalent_items, []);
    var preferred = ["Roasted Beans", "Packing Items"];
    var targetItem = null;
    for (var pi = 0; pi < preferred.length; pi++) {
      for (var ei = 0; ei < equiv.length; ei++) {
        if (String(equiv[ei].category).toLowerCase() === preferred[pi].toLowerCase() && equiv[ei].itemId) {
          targetItem = itemById[String(equiv[ei].itemId)] || null;
          if (targetItem) break;
        }
      }
      if (targetItem) break;
    }
    if (!targetItem) {
      Logger.log("  SKIP ledger " + row.id + " — " + origItem.name + " has no Roasted Beans / Packing Items equivalent");
      skippedNoEquivalent++;
      continue;
    }

    var orderRefId = String(row.reference_id || "");
    var doneBy = String(row.done_by || "system-correction");
    Logger.log("  CORRECT led:" + row.id + " — reverse " + qty + " " + origItem.unit + " to " + origItem.name + ", re-deduct from " + targetItem.name);

    // Write the reversal: put qty BACK onto the Green Beans item.
    createInventoryTransaction(
      origItem.id, "IN", qty,
      "order_delivery_correction", orderRefId,
      "[DELIVERY-CORRECTION] reversal of led:" + row.id + " — originally deducted from wrong item",
      doneBy, ""
    );

    // Write the corrected deduction: remove qty from the Roasted Beans equivalent.
    createInventoryTransaction(
      targetItem.id, "OUT", qty,
      "order_delivery_correction", orderRefId,
      "[DELIVERY-CORRECTION] redirect of led:" + row.id + " — correctly deducted from " + targetItem.category,
      doneBy, ""
    );

    correctionsMade++;
  }

  invalidateCache(SHEETS.INVENTORY_LEDGER);
  invalidateCache(SHEETS.INVENTORY_ITEMS);
  Logger.log("fixDeliveryLedgerEntries: completed. corrections=" + correctionsMade + ", skippedNoEquivalent=" + skippedNoEquivalent + ", skippedAlreadyCorrected=" + skippedAlreadyCorrected);
  return {
    corrections: correctionsMade,
    skippedNoEquivalent: skippedNoEquivalent,
    skippedAlreadyCorrected: skippedAlreadyCorrected,
  };
}

// ─── One-time corrective: recompute current_stock for all items from the ledger ──
// Used as a final settle step after other corrective functions. Sorts surviving ledger
// entries chronologically and rebuilds current_stock = opening_stock + Σ signed quantities.
function recalculateAllInventoryStock() {
  ensureSheetHasAllColumns(SHEETS.INVENTORY_LEDGER);
  ensureSheetHasAllColumns(SHEETS.INVENTORY_ITEMS);
  clearRowsCache();

  var ledger = getRows(SHEETS.INVENTORY_LEDGER);
  var sorted = ledger.slice().sort(function(a, b) {
    var ta = Date.parse(String(a.created_at || "")) || 0;
    var tb = Date.parse(String(b.created_at || "")) || 0;
    return ta - tb;
  });
  var perItemDelta = {};
  for (var s = 0; s < sorted.length; s++) {
    var sr = sorted[s];
    var itemId = String(sr.item_id || "");
    if (!itemId) continue;
    var txType = String(sr.type || "").toUpperCase();
    var q = parseFloat(sr.quantity) || 0;
    var signed;
    if (txType === "IN") signed = Math.abs(q);
    else if (txType === "OUT") signed = -Math.abs(q);
    else if (txType === "ADJUSTMENT") signed = q;
    else signed = q;
    perItemDelta[itemId] = (perItemDelta[itemId] || 0) + signed;
  }

  var items = getRows(SHEETS.INVENTORY_ITEMS);
  var itemsFixed = 0;
  for (var ii = 0; ii < items.length; ii++) {
    var item = items[ii];
    var opening = parseFloat(item.opening_stock) || 0;
    var before = parseFloat(item.current_stock) || 0;
    var delta = perItemDelta[String(item.id)] || 0;
    var after = opening + delta;
    var changed = Math.abs(before - after) > 0.0001;
    Logger.log("item " + item.id + " (" + item.name + "): opening=" + opening + ", before=" + before + ", delta=" + delta + ", after=" + after + (changed ? "  ← CHANGED" : "") + (after < 0 ? "  ⚠ NEGATIVE" : ""));
    if (changed) {
      var idx = findRowIndex(SHEETS.INVENTORY_ITEMS, item.id);
      if (idx > 0) {
        item.current_stock = after;
        updateSheetRow(SHEETS.INVENTORY_ITEMS, idx, item);
        itemsFixed++;
      }
    }
  }
  invalidateCache(SHEETS.INVENTORY_ITEMS);
  Logger.log("recalculateAllInventoryStock: itemsFixed=" + itemsFixed);
  return { itemsScanned: items.length, itemsFixed: itemsFixed };
}

// ─── Combined corrective: run delivery fix → duplicate fix → recompute all stock ──
// The single button to press when you want to settle everything. Run manually.
function fixAllInventoryIssues() {
  Logger.log("═══ fixAllInventoryIssues: START ═══");
  Logger.log("── Step 1: fixDeliveryLedgerEntries ──");
  var step1 = fixDeliveryLedgerEntries();
  Logger.log("── Step 2: fixDuplicateInventoryEntries ──");
  var step2 = fixDuplicateInventoryEntries();
  Logger.log("── Step 3: recalculateAllInventoryStock ──");
  var step3 = recalculateAllInventoryStock();

  Logger.log("── Step 4: final inventory state ──");
  clearRowsCache();
  var finalItems = getRows(SHEETS.INVENTORY_ITEMS);
  var negativeItems = [];
  for (var i = 0; i < finalItems.length; i++) {
    var it = finalItems[i];
    var stock = parseFloat(it.current_stock) || 0;
    Logger.log("  " + it.id + " " + it.name + " [" + it.category + "]: stock=" + stock + (stock < 0 ? "  ⚠ NEGATIVE" : ""));
    if (stock < 0) negativeItems.push({ id: it.id, name: it.name, stock: stock });
  }
  Logger.log("═══ fixAllInventoryIssues: DONE ═══");
  return {
    deliveryFix: step1,
    duplicateFix: step2,
    recalc: step3,
    negativeItems: negativeItems,
  };
}

// ─── Blends ────────────────────────────────────────────────────

function normalizeBlendComponents(arr) {
  if (!Array.isArray(arr)) return [];
  return arr.map(function(c) {
    return {
      category: String(c.category || ""),
      itemId: String(c.itemId || ""),
      itemName: String(c.itemName || ""),
      percentage: parseFloat(c.percentage) || 0,
    };
  });
}

function validateBlendComponents(components) {
  if (!Array.isArray(components) || components.length === 0) return { ok: false, error: "Blend must have at least one component" };
  var total = 0;
  for (var i = 0; i < components.length; i++) {
    var p = parseFloat(components[i].percentage) || 0;
    if (p < 0) return { ok: false, error: "Component percentages must be positive" };
    total += p;
  }
  if (Math.abs(total - 100) > 0.001) return { ok: false, error: "Component percentages must total 100% (currently " + total + "%)" };
  return { ok: true };
}

function handleGetBlends(params) {
  ensureSheetHasAllColumns(SHEETS.BLENDS);
  var rows = getRows(SHEETS.BLENDS);
  var customerFilter = params && (params.customer || params.customerId) || "";
  var includeInactive = params && params.includeInactive === "true";
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var active = r.is_active !== "false" && r.is_active !== false;
    if (!includeInactive && !active) continue;
    var blend = {
      id: r.id,
      name: r.name,
      customer: r.customer || "General",
      description: r.description || "",
      components: normalizeBlendComponents(safeParseJSON(r.components, [])),
      isActive: active,
      createdAt: r.created_at,
    };
    if (customerFilter) {
      // Match either explicit customer label or "General"
      if (String(blend.customer) !== String(customerFilter) && String(blend.customer).toLowerCase() !== "general") continue;
    }
    out.push(blend);
  }
  return out;
}

function handleCreateBlend(body, user) {
  ensureSheetHasAllColumns(SHEETS.BLENDS);
  var name = String(body.name || "").trim();
  if (!name) return { error: "Blend name is required" };
  var components = normalizeBlendComponents(body.components || []);
  var v = validateBlendComponents(components);
  if (!v.ok) return { error: v.error };
  var id = body.id || ("blend_" + nextId());
  var obj = {
    id: id,
    name: name,
    customer: body.customer || "General",
    description: body.description || "",
    components: JSON.stringify(components),
    is_active: "true",
    created_at: new Date().toISOString(),
  };
  appendToSheet(SHEETS.BLENDS, obj);
  writeAuditLog(user, "create", "Blend", id, name);
  return { id: id, name: name, customer: obj.customer, description: obj.description, components: components, isActive: true, createdAt: obj.created_at };
}

function handleUpdateBlend(body, user) {
  ensureSheetHasAllColumns(SHEETS.BLENDS);
  var id = body.id;
  var idx = findRowIndex(SHEETS.BLENDS, id);
  if (idx < 0) return { error: "Blend not found" };
  var rows = getRows(SHEETS.BLENDS);
  var blend = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { blend = rows[i]; break; } }
  if (!blend) return { error: "Blend not found" };
  if (body.name !== undefined) blend.name = String(body.name).trim();
  if (body.customer !== undefined) blend.customer = body.customer || "General";
  if (body.description !== undefined) blend.description = body.description || "";
  if (body.components !== undefined) {
    var components = normalizeBlendComponents(body.components);
    var v = validateBlendComponents(components);
    if (!v.ok) return { error: v.error };
    blend.components = JSON.stringify(components);
  }
  if (body.isActive !== undefined) blend.is_active = body.isActive ? "true" : "false";
  updateSheetRow(SHEETS.BLENDS, idx, blend);
  writeAuditLog(user, "update", "Blend", id, blend.name);
  return {
    id: blend.id, name: blend.name, customer: blend.customer || "General",
    description: blend.description || "",
    components: normalizeBlendComponents(safeParseJSON(blend.components, [])),
    isActive: blend.is_active !== "false" && blend.is_active !== false,
    createdAt: blend.created_at,
  };
}

function handleDeleteBlend(body, user) {
  ensureSheetHasAllColumns(SHEETS.BLENDS);
  var id = body.id;
  var idx = findRowIndex(SHEETS.BLENDS, id);
  if (idx < 0) return { error: "Blend not found" };
  var rows = getRows(SHEETS.BLENDS);
  var blend = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { blend = rows[i]; break; } }
  if (!blend) return { error: "Blend not found" };
  blend.is_active = "false";
  updateSheetRow(SHEETS.BLENDS, idx, blend);
  writeAuditLog(user, "delete", "Blend", id, blend.name);
  return { success: true };
}

// ─── Drafts ────────────────────────────────────────────────────

function handleGetDrafts(params, user) {
  ensureSheetHasAllColumns(SHEETS.DRAFTS);
  var rows = getRows(SHEETS.DRAFTS);
  var isAdmin = user && user.role === "admin";
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (!isAdmin && String(r.user_id) !== String(user.id)) continue;
    out.push({
      id: r.id,
      checklistId: r.checklist_id,
      checklistName: r.checklist_name,
      userId: r.user_id,
      userName: r.user_name,
      responses: safeParseJSON(r.responses, {}),
      linkedSources: safeParseJSON(r.linked_sources, []),
      linkedOrders: safeParseJSON(r.linked_orders, []),
      remarks: safeParseJSON(r.remarks, {}),
      batchAllocations: safeParseJSON(r.batch_allocations, {}),
      person: r.person || "",
      workDate: r.work_date || "",
      createdAt: r.created_at,
      updatedAt: r.updated_at,
    });
  }
  out.sort(function(a, b) { return String(b.updatedAt).localeCompare(String(a.updatedAt)); });
  return out;
}

function handleSaveDraft(body, user) {
  ensureSheetHasAllColumns(SHEETS.DRAFTS);
  var checklistId = body.checklistId || "";
  if (!checklistId) return { error: "checklistId required" };
  var ck = lookupChecklist(checklistId);
  var checklistName = ck ? ck.name : (body.checklistName || "");
  var now = new Date().toISOString();
  var draftId = body.id || ("draft_" + nextId());

  // If updating an existing draft, locate the row
  var existingIdx = -1;
  if (body.id) existingIdx = findRowIndex(SHEETS.DRAFTS, body.id);

  var obj = {
    id: draftId,
    checklist_id: checklistId,
    checklist_name: checklistName,
    user_id: user.id,
    user_name: user.displayName || user.username,
    responses: JSON.stringify(body.responses || {}),
    linked_sources: JSON.stringify(body.linkedSources || []),
    linked_orders: JSON.stringify(body.linkedOrders || []),
    remarks: JSON.stringify(body.remarks || {}),
    batch_allocations: JSON.stringify(body.batchAllocations || {}),
    person: body.person || "",
    work_date: body.workDate || "",
    created_at: existingIdx > 0 ? (body.createdAt || now) : now,
    updated_at: now,
  };

  if (existingIdx > 0) {
    // Permission check: admin or owner
    var rows = getRows(SHEETS.DRAFTS);
    var existing = null;
    for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(body.id)) { existing = rows[i]; break; } }
    if (existing && user.role !== "admin" && String(existing.user_id) !== String(user.id)) {
      return { error: "You can only edit your own drafts" };
    }
    if (existing) obj.created_at = existing.created_at || obj.created_at;
    updateSheetRow(SHEETS.DRAFTS, existingIdx, obj);
  } else {
    appendToSheet(SHEETS.DRAFTS, obj);
  }
  return {
    id: obj.id, checklistId: obj.checklist_id, checklistName: obj.checklist_name,
    userId: obj.user_id, userName: obj.user_name,
    responses: safeParseJSON(obj.responses, {}),
    linkedSources: safeParseJSON(obj.linked_sources, []),
    linkedOrders: safeParseJSON(obj.linked_orders, []),
    remarks: safeParseJSON(obj.remarks, {}),
    batchAllocations: safeParseJSON(obj.batch_allocations, {}),
    person: obj.person, workDate: obj.work_date,
    createdAt: obj.created_at, updatedAt: obj.updated_at,
  };
}

function handleDeleteDraft(body, user) {
  ensureSheetHasAllColumns(SHEETS.DRAFTS);
  var id = body.id;
  var idx = findRowIndex(SHEETS.DRAFTS, id);
  if (idx < 0) return { success: true }; // already gone
  var rows = getRows(SHEETS.DRAFTS);
  var existing = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === String(id)) { existing = rows[i]; break; } }
  if (existing && user.role !== "admin" && String(existing.user_id) !== String(user.id)) {
    return { error: "You can only delete your own drafts" };
  }
  deleteSheetRow(SHEETS.DRAFTS, idx);
  return { success: true };
}

// Generic standalone allocation: tag a source autoId to a destination (checklist autoId or order id) with a quantity.
// Validates against the source's remaining (sourceTotal - already-allocated). Used by the untagged tagging UI.
function handleCreateAllocation(body, user) {
  ensureSheetHasAllColumns(SHEETS.QUANTITY_ALLOCATIONS);
  var sourceAutoId = String(body.sourceAutoId || "").trim();
  var sourceChecklistId = body.sourceChecklistId || "";
  var destinationType = body.destinationType || "checklist";
  var destinationId = body.destinationId || "";
  var destinationAutoId = body.destinationAutoId || "";
  var quantity = parseFloat(body.quantity) || 0;
  if (!sourceAutoId || !destinationId || quantity <= 0) {
    return { error: "sourceAutoId, destinationId and positive quantity are required" };
  }
  // Resolve source total via the source checklist
  var sourceCk = sourceChecklistId ? lookupChecklist(sourceChecklistId) : null;
  if (!sourceCk) {
    // Fallback: scan OrderChecklists/UntaggedChecklists for the autoId and find its checklist
    var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
    for (var i = 0; i < ocs.length; i++) {
      if (String(ocs[i].auto_id) === sourceAutoId) { sourceCk = lookupChecklist(ocs[i].checklist_id); break; }
    }
    if (!sourceCk) {
      var uts = getRows(SHEETS.UNTAGGED_CHECKLISTS);
      for (var j = 0; j < uts.length; j++) {
        if (String(uts[j].auto_id) === sourceAutoId) { sourceCk = lookupChecklist(uts[j].checklist_id); break; }
      }
    }
  }
  if (!sourceCk) return { error: "Source checklist not found for " + sourceAutoId };
  var srcInfo = getSubmissionByAutoId(sourceAutoId, sourceCk);
  var srcTotal = srcInfo ? srcInfo.totalQuantity : 0;
  if (srcTotal <= 0) return { error: "Source has no trackable quantity" };
  var alreadyAllocated = getAllocatedQuantityForAutoId(sourceAutoId);
  var remaining = srcTotal - alreadyAllocated;
  if (quantity > remaining + 0.0001) {
    return { error: "Insufficient quantity. Only " + remaining + " available from " + sourceAutoId };
  }
  createQuantityAllocation(sourceCk.id, sourceAutoId, srcTotal, destinationType, destinationId, destinationAutoId, quantity, user.displayName || user.username);
  writeAuditLog(user, "tag_allocation", "QuantityAllocation", sourceAutoId, "→ " + destinationType + " " + (destinationAutoId || destinationId) + " (" + quantity + ")");
  return { success: true, remaining: remaining - quantity };
}

// ─── Response Chain (Feature: Universal chain tagging) ─────────

function handleGetResponseChain(params) {
  var checklistId = params.checklist_id || params.checklistId || "";
  var responseId = params.response_id || params.responseId || "";
  if (!checklistId || !responseId) return { error: "Missing checklist_id or response_id (autoId)" };
  var chain = [];
  var visited = {};
  var curCkId = checklistId;
  var curAutoId = responseId;
  for (var depth = 0; depth < 10; depth++) {
    if (!curCkId || !curAutoId) break;
    var key = curCkId + "::" + curAutoId;
    if (visited[key]) break;
    visited[key] = true;
    var ck = lookupChecklist(curCkId);
    if (!ck) { chain.push({ checklistId: curCkId, checklistName: "Unknown", responseId: curAutoId, autoId: curAutoId, fields: [], error: "unavailable" }); break; }
    // Find the submission by autoId
    var submission = findSubmissionByAutoId(curAutoId, ck);
    if (!submission) { chain.push({ checklistId: curCkId, checklistName: ck.name, responseId: curAutoId, autoId: curAutoId, fields: [], error: "unavailable" }); break; }
    var fields = [];
    var nq = ck.questions;
    for (var qi = 0; qi < nq.length; qi++) {
      fields.push({ question: nq[qi].text, response: submission.responses[qi] || "", type: nq[qi].type || "text" });
    }
    chain.push({ checklistId: curCkId, checklistName: ck.name, responseId: curAutoId, autoId: curAutoId, fields: fields });
    // Find the linked source field to continue the chain
    var nextCkId = "", nextAutoId = "";
    for (var li = 0; li < nq.length; li++) {
      if (nq[li].linkedSource && nq[li].linkedSource.checklistId) {
        nextCkId = nq[li].linkedSource.checklistId;
        nextAutoId = String(submission.responses[li] || "").trim();
        // If batch-allocated (comma-separated), take the first one
        if (nextAutoId.indexOf(",") >= 0) nextAutoId = nextAutoId.split(",")[0].trim();
        break;
      }
    }
    curCkId = nextCkId;
    curAutoId = nextAutoId;
  }
  return chain;
}

// Finds a submission by autoId from OrderChecklists or UntaggedChecklists response tabs.
// Returns { responses: {0: "val", 1: "val", ...} } or null.
function findSubmissionByAutoId(autoId, ck) {
  if (!autoId || !ck) return null;
  // Check OrderChecklists
  ensureSheetHasAllColumns(SHEETS.ORDER_CHECKLISTS);
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  for (var i = 0; i < ocs.length; i++) {
    if (String(ocs[i].auto_id) === String(autoId) && String(ocs[i].checklist_id) === String(ck.id)) {
      var resp = readResponseRow(ck.name, ck.questions, String(ocs[i].order_id));
      if (resp) {
        var rmap = {};
        for (var r = 0; r < resp.responses.length; r++) rmap[resp.responses[r].questionIndex] = resp.responses[r].response;
        return { responses: rmap };
      }
    }
  }
  // Check UntaggedChecklists
  ensureSheetHasAllColumns(SHEETS.UNTAGGED_CHECKLISTS);
  var uts = getRows(SHEETS.UNTAGGED_CHECKLISTS);
  for (var j = 0; j < uts.length; j++) {
    if (String(uts[j].auto_id) === String(autoId) && String(uts[j].checklist_id) === String(ck.id)) {
      var responses = safeParseJSON(uts[j].responses, []);
      return { responses: responsesArrayToMap(responses) };
    }
  }
  return null;
}

// ─── Order Type Requirements Config ──────────────────────────

function getOrderTypeRequirementsConfig() {
  var rows = getRows(SHEETS.CONFIG);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].key) === "order_type_requirements") {
      return safeParseJSON(rows[i].value, {});
    }
  }
  return {};
}

function handleSaveOrderTypeRequirements(body, user) {
  var requirements = body.requirements || {};
  var sheet = getSheet(SHEETS.CONFIG);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === "order_type_requirements") {
      sheet.getRange(i + 1, 2).setValue(JSON.stringify(requirements));
      invalidateCache(SHEETS.CONFIG);
      writeAuditLog(user, "update", "Config", "order_type_requirements", "Updated order type requirements");
      return { success: true, requirements: requirements };
    }
  }
  // Not found — create it
  sheet.appendRow(["order_type_requirements", JSON.stringify(requirements)]);
  invalidateCache(SHEETS.CONFIG);
  writeAuditLog(user, "create", "Config", "order_type_requirements", "Created order type requirements");
  return { success: true, requirements: requirements };
}

// Find all upstream references to an autoId.
// Searches QuantityAllocations (where source_auto_id === autoId) and Orders.stages.taggedEntries.
// Returns an array of referencing auto IDs / stage identifiers.
function findUpstreamReferencesForAutoId(autoId) {
  var refs = [];
  if (!autoId) return refs;
  // Quantity allocations where this autoId is the SOURCE (something else is tagged to it)
  ensureSheetHasAllColumns(SHEETS.QUANTITY_ALLOCATIONS);
  var qa = getRows(SHEETS.QUANTITY_ALLOCATIONS);
  for (var i = 0; i < qa.length; i++) {
    if (String(qa[i].source_auto_id) === String(autoId)) {
      var destAid = qa[i].destination_auto_id || qa[i].destination_id;
      if (destAid && refs.indexOf(String(destAid)) < 0) refs.push(String(destAid));
    }
  }
  // Order stages referencing this autoId
  var orders = getRows(SHEETS.ORDERS);
  for (var oi = 0; oi < orders.length; oi++) {
    var stages = safeParseJSON(orders[oi].stages, []);
    for (var si = 0; si < stages.length; si++) {
      var tagged = Array.isArray(stages[si].taggedEntries) ? stages[si].taggedEntries : [];
      for (var ti = 0; ti < tagged.length; ti++) {
        if (String(tagged[ti].responseId || tagged[ti].autoId) === String(autoId)) {
          var label = (orders[oi].id || "?") + "/" + (stages[si].name || "stage");
          if (refs.indexOf(label) < 0) refs.push(label);
        }
      }
    }
  }
  return refs;
}

// ─── Order Stage Templates Config ────────────────────────────

function getOrderStageTemplatesConfig() {
  var rows = getRows(SHEETS.CONFIG);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].key) === "order_stage_templates") {
      return safeParseJSON(rows[i].value, {});
    }
  }
  return {};
}

function handleSaveOrderStageTemplates(body, user) {
  var templates = body.templates || {};
  // Normalize each template entry
  Object.keys(templates).forEach(function(pt) {
    var arr = Array.isArray(templates[pt]) ? templates[pt] : [];
    templates[pt] = arr.map(function(s, i) {
      return {
        name: String(s.name || ("Stage " + (i+1))),
        checklistId: s.checklistId || "",
        quantityField: s.quantityField || "",
        requiredQty: parseFloat(s.requiredQty) || 0,
        position: i,
      };
    });
  });
  var sheet = getSheet(SHEETS.CONFIG);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === "order_stage_templates") {
      sheet.getRange(i + 1, 2).setValue(JSON.stringify(templates));
      invalidateCache(SHEETS.CONFIG);
      writeAuditLog(user, "update", "Config", "order_stage_templates", "Updated order stage templates");
      return { success: true, templates: templates };
    }
  }
  sheet.appendRow(["order_stage_templates", JSON.stringify(templates)]);
  invalidateCache(SHEETS.CONFIG);
  writeAuditLog(user, "create", "Config", "order_stage_templates", "Created order stage templates");
  return { success: true, templates: templates };
}

// ─── Stage Tag / Untag Handlers ──────────────────────────────

function handleTagChecklistToStage(body, user) {
  var orderId = body.orderId;
  var stageId = body.stageId;
  var qty = parseFloat(body.quantity) || 0;
  var isMixedBlend = body.isMixedBlend === true;
  if (!orderId || !stageId || qty <= 0) return { error: "Missing orderId, stageId or quantity" };

  var idx = findRowIndex(SHEETS.ORDERS, orderId);
  if (idx < 0) return { error: "Order not found" };
  var orders = getRows(SHEETS.ORDERS);
  var order = null;
  for (var i = 0; i < orders.length; i++) { if (String(orders[i].id) === String(orderId)) { order = orders[i]; break; } }
  if (!order) return { error: "Order not found" };

  var stages = safeParseJSON(order.stages, []);
  var stage = null;
  for (var s = 0; s < stages.length; s++) { if (String(stages[s].id) === String(stageId)) { stage = stages[s]; break; } }
  if (!stage) return { error: "Stage not found in order" };

  // Blend-related metadata (optional, used for blend-order tagging)
  var blendLineIndex = (body.blendLineIndex !== undefined && body.blendLineIndex !== null && body.blendLineIndex !== "") ? Number(body.blendLineIndex) : null;
  var componentItemId = body.componentItemId || "";
  var componentItemName = body.componentItemName || "";
  var now = new Date().toISOString();

  if (!Array.isArray(stage.taggedEntries)) stage.taggedEntries = [];

  if (isMixedBlend) {
    // ── Mixed-blend tag: direct inventory tag, no checklist. ──
    var mixedItemId = body.mixedInventoryItemId || "";
    var mixedItemName = body.mixedInventoryItemName || "";
    var mixedBlendId = body.mixedBlendId || "";
    if (!mixedItemId) return { error: "Missing mixedInventoryItemId for mixed blend tag" };
    if (blendLineIndex === null || blendLineIndex < 0) return { error: "blendLineIndex required for mixed blend tag" };
    // Validate the inventory item exists
    var invItems = getRows(SHEETS.INVENTORY_ITEMS);
    var invItem = null;
    for (var ii = 0; ii < invItems.length; ii++) {
      if (String(invItems[ii].id) === String(mixedItemId)) { invItem = invItems[ii]; break; }
    }
    if (!invItem) return { error: "Mixed inventory item not found: " + mixedItemId };
    if (!mixedItemName) mixedItemName = invItem.name;
    // Validate the order's blend line ratio matches the Blend record ratio exactly
    var orderLines = safeParseJSON(order.order_lines, []);
    if (blendLineIndex >= orderLines.length) return { error: "blendLineIndex out of range" };
    var targetLine = orderLines[blendLineIndex];
    if (!Array.isArray(targetLine.blendComponents) || targetLine.blendComponents.length === 0) return { error: "Target blend line has no components" };
    // Look up the Blend record by matching the mixedBlendId (preferred) or by inventory item name
    var blendRows = getRows(SHEETS.BLENDS);
    var matchedBlend = null;
    if (mixedBlendId) {
      for (var bi = 0; bi < blendRows.length; bi++) {
        if (String(blendRows[bi].id) === String(mixedBlendId)) { matchedBlend = blendRows[bi]; break; }
      }
    }
    if (!matchedBlend) {
      for (var bj = 0; bj < blendRows.length; bj++) {
        if (String(blendRows[bj].name).toLowerCase() === String(mixedItemName).toLowerCase()) { matchedBlend = blendRows[bj]; break; }
      }
    }
    if (!matchedBlend) return { error: "No blend definition found for '" + mixedItemName + "' — cannot validate ratios" };
    var blendComps = safeParseJSON(matchedBlend.components, []);
    // Compare ratios exactly (keyed by itemId OR lowercase name)
    var ratioMismatch = compareBlendRatiosExact_(blendComps, targetLine.blendComponents);
    if (ratioMismatch) return { error: "Blend ratio mismatch — " + ratioMismatch };

    stage.taggedEntries.push({
      qty: qty,
      blendLineIndex: blendLineIndex,
      isMixed: true,
      mixedItemId: mixedItemId,
      mixedItemName: mixedItemName,
      mixedBlendId: matchedBlend.id,
      taggedAt: now,
      taggedBy: user.displayName || user.username,
    });

    order.stages = JSON.stringify(stages);
    updateSheetRow(SHEETS.ORDERS, idx, order);

    writeAuditLog(user, "tag_stage_mixed", "Order", orderId, "Tagged mixed blend " + mixedItemName + " (" + qty + "kg) to stage " + (stage.name || stageId) + " line " + blendLineIndex);
    return { success: true, stages: stages };
  }

  // ── Standard checklist-entry tag ──
  var responseId = body.responseId || body.autoId || "";
  var sourceChecklistId = body.sourceChecklistId || "";
  if (!responseId) return { error: "Missing responseId" };

  var sourceCk = sourceChecklistId ? lookupChecklist(sourceChecklistId) : null;
  if (!sourceCk) {
    var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
    for (var oi2 = 0; oi2 < ocs.length; oi2++) {
      if (String(ocs[oi2].auto_id) === String(responseId)) { sourceCk = lookupChecklist(ocs[oi2].checklist_id); break; }
    }
    if (!sourceCk) {
      var uts2 = getRows(SHEETS.UNTAGGED_CHECKLISTS);
      for (var ui2 = 0; ui2 < uts2.length; ui2++) {
        if (String(uts2[ui2].auto_id) === String(responseId)) { sourceCk = lookupChecklist(uts2[ui2].checklist_id); break; }
      }
    }
  }
  if (!sourceCk) return { error: "Source checklist not found for " + responseId };
  var srcInfo = getSubmissionByAutoId(responseId, sourceCk);
  var srcTotal = srcInfo ? srcInfo.totalQuantity : 0;
  if (srcTotal <= 0) return { error: "Source has no trackable quantity" };
  var alreadyAllocated = getAllocatedQuantityForAutoId(responseId);
  var remaining = srcTotal - alreadyAllocated;
  if (qty > remaining + 0.0001) return { error: "Insufficient quantity. Only " + remaining + " available from " + responseId };

  createQuantityAllocation(sourceCk.id, responseId, srcTotal, "order_stage", orderId + ":" + stageId, responseId, qty, user.displayName || user.username);

  stage.taggedEntries.push({
    responseId: responseId,
    checklistId: sourceCk.id,
    autoId: responseId,
    qty: qty,
    quantityFieldValue: body.quantityFieldValue !== undefined ? body.quantityFieldValue : "",
    blendLineIndex: blendLineIndex !== null ? blendLineIndex : undefined,
    componentItemId: componentItemId || undefined,
    componentItemName: componentItemName || undefined,
    taggedAt: now,
    taggedBy: user.displayName || user.username,
  });

  order.stages = JSON.stringify(stages);
  updateSheetRow(SHEETS.ORDERS, idx, order);

  writeAuditLog(user, "tag_stage", "Order", orderId, "Tagged " + responseId + " (" + qty + ") to stage " + (stage.name || stageId) + (componentItemName?" as "+componentItemName:""));
  return { success: true, stages: stages };
}

// Compare two blend component arrays for exact ratio match. Returns null on match, or a descriptive mismatch string.
// Each side is an array of { itemId?, itemName?, percentage }.
function compareBlendRatiosExact_(blendA, blendB) {
  var keyFor = function(c) { return c.itemId ? "id:" + c.itemId : "n:" + String(c.itemName || "").toLowerCase().trim(); };
  var totalA = 0, totalB = 0;
  var mapA = {}, mapB = {};
  (blendA || []).forEach(function(c) {
    var k = keyFor(c);
    mapA[k] = (mapA[k] || 0) + (parseFloat(c.percentage) || 0);
    totalA += parseFloat(c.percentage) || 0;
  });
  (blendB || []).forEach(function(c) {
    var k = keyFor(c);
    mapB[k] = (mapB[k] || 0) + (parseFloat(c.percentage) || 0);
    totalB += parseFloat(c.percentage) || 0;
  });
  var keys = {};
  Object.keys(mapA).forEach(function(k) { keys[k] = true; });
  Object.keys(mapB).forEach(function(k) { keys[k] = true; });
  var labelFor = function(c) { return c.itemName || c.itemId || "unknown"; };
  var aDesc = (blendA || []).map(function(c){return (parseFloat(c.percentage)||0) + "% " + labelFor(c);}).join(", ");
  var bDesc = (blendB || []).map(function(c){return (parseFloat(c.percentage)||0) + "% " + labelFor(c);}).join(", ");
  for (var k in keys) {
    var a = parseFloat(mapA[k]) || 0;
    var b = parseFloat(mapB[k]) || 0;
    if (Math.abs(a - b) > 0.01) {
      return "selected blend uses " + aDesc + " but order requires " + bDesc;
    }
  }
  return null;
}

function handleUntagChecklistFromStage(body, user) {
  var orderId = body.orderId;
  var stageId = body.stageId;
  var responseId = body.responseId || body.autoId || "";
  // Optional precise locators for cases where responseId alone isn't unique (e.g. multiple tags of same autoId
  // for different ingredients, or mixed-blend tags with no responseId at all).
  var blendLineIndex = (body.blendLineIndex !== undefined && body.blendLineIndex !== null && body.blendLineIndex !== "") ? Number(body.blendLineIndex) : null;
  var componentItemId = body.componentItemId || "";
  var mixedItemId = body.mixedItemId || "";
  var isMixedBlend = body.isMixedBlend === true;
  if (!orderId || !stageId) return { error: "Missing orderId or stageId" };
  if (!responseId && !isMixedBlend) return { error: "Missing responseId or isMixedBlend" };

  var idx = findRowIndex(SHEETS.ORDERS, orderId);
  if (idx < 0) return { error: "Order not found" };
  var orders = getRows(SHEETS.ORDERS);
  var order = null;
  for (var i = 0; i < orders.length; i++) { if (String(orders[i].id) === String(orderId)) { order = orders[i]; break; } }
  if (!order) return { error: "Order not found" };

  var stages = safeParseJSON(order.stages, []);
  var stage = null;
  for (var s = 0; s < stages.length; s++) { if (String(stages[s].id) === String(stageId)) { stage = stages[s]; break; } }
  if (!stage) return { error: "Stage not found in order" };

  var removed = null;
  if (Array.isArray(stage.taggedEntries)) {
    for (var ti = 0; ti < stage.taggedEntries.length; ti++) {
      var te = stage.taggedEntries[ti];
      var match = false;
      if (isMixedBlend) {
        if (te.isMixed === true
          && (!mixedItemId || String(te.mixedItemId) === String(mixedItemId))
          && (blendLineIndex === null || Number(te.blendLineIndex) === blendLineIndex)) match = true;
      } else if (responseId && String(te.responseId || te.autoId) === String(responseId)) {
        // If componentItemId was provided, require a match (multiple tags of the same autoId possible in blend orders)
        if (componentItemId) {
          if (String(te.componentItemId || "") === String(componentItemId)) match = true;
        } else if (blendLineIndex !== null) {
          if (Number(te.blendLineIndex) === blendLineIndex) match = true;
        } else {
          match = true;
        }
      }
      if (match) {
        removed = te;
        stage.taggedEntries.splice(ti, 1);
        break;
      }
    }
  }
  if (!removed) return { error: "Tagged entry not found in stage" };

  // Reverse the corresponding allocation for this destination (only for checklist-based tags)
  if (!removed.isMixed && (removed.autoId || removed.responseId)) {
    var destKey = orderId + ":" + stageId;
    var srcAid = removed.autoId || removed.responseId;
    var qaSheet = getSheet(SHEETS.QUANTITY_ALLOCATIONS);
    if (qaSheet && qaSheet.getLastRow() >= 2) {
      var qaData = qaSheet.getDataRange().getValues();
      // Delete the single best-matching allocation row (by source_auto_id + dest + allocated_quantity)
      for (var qi = qaData.length - 1; qi >= 1; qi--) {
        if (String(qaData[qi][2]) === String(srcAid)
          && String(qaData[qi][5]) === String(destKey)
          && Math.abs((parseFloat(qaData[qi][7]) || 0) - (parseFloat(removed.qty) || 0)) < 0.0001) {
          qaSheet.deleteRow(qi + 1);
          invalidateCache(SHEETS.QUANTITY_ALLOCATIONS);
          break;
        }
      }
    }
  }

  order.stages = JSON.stringify(stages);
  updateSheetRow(SHEETS.ORDERS, idx, order);

  var logLabel = removed.isMixed ? ("mixed " + (removed.mixedItemName || removed.mixedItemId)) : (removed.autoId || removed.responseId || "entry");
  writeAuditLog(user, "untag_stage", "Order", orderId, "Removed " + logLabel + " from stage " + (stage.name || stageId));
  return { success: true, stages: stages };
}

// Resolve the OUTPUT inventory item for a tagged checklist entry (backend).
// Priority: IN-tx inventoryLink (with category-to-equivalent mapping if raw item is in a different category),
// then any inventoryLink, then plain inventory_item question. Matches the frontend logic + processInventoryLinks behavior.
// Returns { id, name } or null.
function resolveTaggedEntryItemBackend_(te) {
  if (!te || !te.checklistId) return null;
  var ck = lookupChecklist(te.checklistId);
  if (!ck) return null;
  var nq = ck.questions || [];

  // Prefer IN-tx inventoryLink
  var chosen = null;
  for (var i = 0; i < nq.length; i++) {
    var q = nq[i];
    if (q && q.inventoryLink && q.inventoryLink.enabled && q.inventoryLink.txType === "IN") { chosen = { q: q, idx: i }; break; }
  }
  if (!chosen) {
    for (var i2 = 0; i2 < nq.length; i2++) {
      var q2 = nq[i2];
      if (q2 && q2.inventoryLink && q2.inventoryLink.enabled) { chosen = { q: q2, idx: i2 }; break; }
    }
  }
  if (!chosen) {
    for (var j = 0; j < nq.length; j++) {
      if (nq[j] && nq[j].type === "inventory_item") { chosen = { q: nq[j], idx: j, plainInv: true }; break; }
    }
  }
  if (!chosen) return null;

  var link = chosen.q.inventoryLink || null;
  var targetCategory = link && link.category ? link.category : "";
  var items = getRows(SHEETS.INVENTORY_ITEMS);

  // Fixed item source — use directly
  if (link && link.itemSource && link.itemSource.type === "fixed" && link.itemSource.itemId) {
    for (var fi = 0; fi < items.length; fi++) {
      if (String(items[fi].id) === String(link.itemSource.itemId)) return { id: items[fi].id, name: items[fi].name };
    }
    return { id: link.itemSource.itemId, name: link.itemSource.itemId };
  }

  // Resolve the source field index
  var itemFieldIdx = -1;
  if (link && link.itemSource && link.itemSource.type === "field" && link.itemSource.fieldIdx !== undefined && link.itemSource.fieldIdx !== null && link.itemSource.fieldIdx !== "") {
    itemFieldIdx = Number(link.itemSource.fieldIdx);
  } else if (chosen.plainInv) {
    itemFieldIdx = chosen.idx;
  }
  if (itemFieldIdx < 0) return null;

  var autoId = te.autoId || te.responseId;
  var sub = findSubmissionByAutoId(autoId, ck);
  if (!sub || !sub.responses) return null;
  var itemVal = String(sub.responses[itemFieldIdx] || "").trim();
  if (!itemVal) return null;

  // Resolve the raw item, then map to target category via equivalent_items if needed
  var raw = null;
  for (var k = 0; k < items.length; k++) {
    if (String(items[k].id) === itemVal || String(items[k].name) === itemVal) { raw = items[k]; break; }
  }
  if (!raw) return { id: "", name: itemVal };

  if (targetCategory && raw.category && String(raw.category) !== String(targetCategory)) {
    var equiv = safeParseJSON(raw.equivalent_items, []);
    for (var ei = 0; ei < equiv.length; ei++) {
      if (String(equiv[ei].category) === String(targetCategory) && equiv[ei].itemId) {
        for (var xi = 0; xi < items.length; xi++) {
          if (String(items[xi].id) === String(equiv[ei].itemId)) return { id: items[xi].id, name: items[xi].name };
        }
        return { id: equiv[ei].itemId, name: equiv[ei].itemId };
      }
    }
  }
  return { id: raw.id, name: raw.name };
}

// ─── Deliver Order (marks delivered + deducts inventory) ────

function handleDeliverOrder(body, user) {
  var orderId = body.id;
  if (!orderId) return { error: "Missing order id" };
  var confirmed = body.confirmed === true;

  var idx = findRowIndex(SHEETS.ORDERS, orderId);
  if (idx < 0) return { error: "Order not found" };
  var orders = getRows(SHEETS.ORDERS);
  var order = null;
  for (var i = 0; i < orders.length; i++) { if (String(orders[i].id) === String(orderId)) { order = orders[i]; break; } }
  if (!order) return { error: "Order not found" };

  var stages = safeParseJSON(order.stages, []);
  // Compute planned deductions: for each tagged entry across stages, find output inventory item from that checklist
  var deductions = [];
  for (var si = 0; si < stages.length; si++) {
    var tagged = Array.isArray(stages[si].taggedEntries) ? stages[si].taggedEntries : [];
    for (var ti = 0; ti < tagged.length; ti++) {
      var te = tagged[ti];
      var blendLineIdx = (te.blendLineIndex !== undefined && te.blendLineIndex !== null) ? Number(te.blendLineIndex) : null;

      // ── Mixed blend tag: deduct from the chosen pre-blended inventory item directly ──
      if (te.isMixed === true && te.mixedItemId) {
        var mixName = te.mixedItemName || "";
        if (!mixName) {
          var invRowsMx = getRows(SHEETS.INVENTORY_ITEMS);
          for (var mxi = 0; mxi < invRowsMx.length; mxi++) {
            if (String(invRowsMx[mxi].id) === String(te.mixedItemId)) { mixName = invRowsMx[mxi].name; break; }
          }
        }
        deductions.push({
          itemId: te.mixedItemId, itemName: mixName, category: "",
          qty: parseFloat(te.qty) || 0,
          stageName: stages[si].name || "Stage",
          checklistAutoId: "mixed:" + mixName,
          blendLineIndex: blendLineIdx,
          isMixed: true,
        });
        continue;
      }

      // ── Checklist-based tag: resolve via inventoryLink ──
      if (!te.checklistId) continue;
      var ck = lookupChecklist(te.checklistId);
      if (!ck) continue;

      // Order delivery must deduct from the OUTPUT item of the production stage.
      // By convention in this system: the OUT-tx inventoryLink question tracks the
      // OUTPUT item (what was produced, e.g. roasted beans / packed goods). Prefer
      // that question. If no OUT-tx question exists, fall back to the IN-tx question
      // and remap to an output-category equivalent if the resolved item is a raw input.
      var resolveItemForQuestion = function(q) {
        if (!q || !q.inventoryLink || !q.inventoryLink.enabled || !q.inventoryLink.itemSource) return null;
        var src = q.inventoryLink.itemSource;
        if (src.type === "fixed" && src.itemId) {
          return { itemId: src.itemId, itemName: "", category: "" };
        }
        if (src.type === "field") {
          var sub = findSubmissionByAutoId(te.autoId || te.responseId, ck);
          if (!sub) return null;
          var itemVal = sub.responses[src.fieldIdx];
          if (!itemVal) return null;
          var invRows = getRows(SHEETS.INVENTORY_ITEMS);
          for (var ir = 0; ir < invRows.length; ir++) {
            if (String(invRows[ir].id) === String(itemVal) || String(invRows[ir].name) === String(itemVal)) {
              return { itemId: invRows[ir].id, itemName: invRows[ir].name, category: invRows[ir].category };
            }
          }
        }
        return null;
      };

      var outItemId = "", outItemName = "", outCategory = "";
      var usedFallback = false;

      // Pass 1: prefer OUT-tx question (tracks the output item directly)
      for (var qi = 0; qi < ck.questions.length; qi++) {
        var qOut = ck.questions[qi];
        if (!qOut || !qOut.inventoryLink || !qOut.inventoryLink.enabled) continue;
        if (qOut.inventoryLink.txType !== "OUT") continue;
        var resO = resolveItemForQuestion(qOut);
        if (resO && resO.itemId) {
          outItemId = resO.itemId; outItemName = resO.itemName || ""; outCategory = resO.category || "";
          break;
        }
      }

      // Pass 2: fall back to IN-tx question (legacy config). Remap to output-category
      // equivalent via equivalent_items when the resolved item is a raw input like Green Beans.
      if (!outItemId) {
        usedFallback = true;
        for (var qi2 = 0; qi2 < ck.questions.length; qi2++) {
          var qIn = ck.questions[qi2];
          if (!qIn || !qIn.inventoryLink || !qIn.inventoryLink.enabled) continue;
          if (qIn.inventoryLink.txType !== "IN") continue;
          var resI = resolveItemForQuestion(qIn);
          if (resI && resI.itemId) {
            outItemId = resI.itemId; outItemName = resI.itemName || ""; outCategory = resI.category || "";
            break;
          }
        }
        if (outItemId) {
          var invRowsAll = getRows(SHEETS.INVENTORY_ITEMS);
          var resolvedItem = null;
          for (var irR = 0; irR < invRowsAll.length; irR++) {
            if (String(invRowsAll[irR].id) === String(outItemId)) { resolvedItem = invRowsAll[irR]; break; }
          }
          if (resolvedItem) {
            var OUTPUT_CATEGORIES = ["Packing Items", "Roasted Beans"];
            var resolvedCat = String(resolvedItem.category || "");
            var isOutputCat = false;
            for (var oc = 0; oc < OUTPUT_CATEGORIES.length; oc++) {
              if (resolvedCat.toLowerCase() === OUTPUT_CATEGORIES[oc].toLowerCase()) { isOutputCat = true; break; }
            }
            if (!isOutputCat) {
              var equiv = safeParseJSON(resolvedItem.equivalent_items, []);
              var foundEquivId = "", foundEquivCat = "";
              for (var pref = 0; pref < OUTPUT_CATEGORIES.length; pref++) {
                for (var eqi = 0; eqi < equiv.length; eqi++) {
                  if (String(equiv[eqi].category).toLowerCase() === OUTPUT_CATEGORIES[pref].toLowerCase() && equiv[eqi].itemId) {
                    foundEquivId = equiv[eqi].itemId; foundEquivCat = equiv[eqi].category; break;
                  }
                }
                if (foundEquivId) break;
              }
              if (foundEquivId) {
                outItemId = foundEquivId; outItemName = ""; outCategory = foundEquivCat;
                for (var irE = 0; irE < invRowsAll.length; irE++) {
                  if (String(invRowsAll[irE].id) === String(outItemId)) { outItemName = invRowsAll[irE].name; outCategory = invRowsAll[irE].category || outCategory; break; }
                }
              } else {
                Logger.log("handleDeliverOrder: WARNING — IN-tx resolved item " + resolvedItem.id + " (" + resolvedCat + ") has no output-category equivalent");
              }
            }
          }
        }
      }

      if (outItemId) {
        if (!outItemName) {
          var invRows2 = getRows(SHEETS.INVENTORY_ITEMS);
          for (var ir2 = 0; ir2 < invRows2.length; ir2++) {
            if (String(invRows2[ir2].id) === String(outItemId)) { outItemName = invRows2[ir2].name; outCategory = invRows2[ir2].category; break; }
          }
        }
        if (usedFallback) Logger.log("handleDeliverOrder: using IN-tx fallback for checklist " + ck.id);
        deductions.push({
          itemId: outItemId, itemName: outItemName, category: outCategory,
          qty: parseFloat(te.qty) || 0,
          stageName: stages[si].name || "Stage",
          checklistAutoId: te.autoId || te.responseId,
          blendLineIndex: blendLineIdx,
          componentItemName: te.componentItemName || outItemName,
        });
      }
    }
  }

  // ── Blend composition validation: per blend line, verify ingredients tagged with exact ratios ──
  var blendErrors = [];
  var blendWarnings = [];
  try {
    var orderLinesBlendV = safeParseJSON(order.order_lines, []).filter(function(l) {
      return Array.isArray(l.blendComponents) && l.blendComponents.length > 0;
    });
    if (orderLinesBlendV.length > 0) {
      var rawLines = safeParseJSON(order.order_lines, []);
      for (var lineIdx = 0; lineIdx < rawLines.length; lineIdx++) {
        var rawLine = rawLines[lineIdx];
        if (!Array.isArray(rawLine.blendComponents) || rawLine.blendComponents.length === 0) continue;
        var lineQty = parseFloat(rawLine.quantity) || 0;
        var components = rawLine.blendComponents;
        var keyFor = function(c) { return c.itemId ? ("id:" + c.itemId) : ("n:" + String(c.itemName || "").toLowerCase().trim()); };

        // Required per ingredient for this line
        var perIngredient = components.map(function(c) {
          return { key: keyFor(c), itemId: c.itemId, itemName: c.itemName, required: ((parseFloat(c.percentage) || 0) / 100) * lineQty, percentage: parseFloat(c.percentage) || 0 };
        });

        // Gather all taggedEntries matching this blend line from all stages
        var lineTags = [];
        for (var si2 = 0; si2 < stages.length; si2++) {
          var tagged2 = Array.isArray(stages[si2].taggedEntries) ? stages[si2].taggedEntries : [];
          for (var ti2 = 0; ti2 < tagged2.length; ti2++) {
            var te2 = tagged2[ti2];
            if (Number(te2.blendLineIndex) === lineIdx) lineTags.push(te2);
          }
        }

        // Mixed contribution for this line
        var mixedTotal = 0;
        for (var mxi2 = 0; mxi2 < lineTags.length; mxi2++) {
          if (lineTags[mxi2].isMixed === true) mixedTotal += parseFloat(lineTags[mxi2].qty) || 0;
        }

        // Direct (non-mixed) tagged qty per ingredient
        var directPerKey = {};
        for (var dti = 0; dti < lineTags.length; dti++) {
          var dte = lineTags[dti];
          if (dte.isMixed === true) continue;
          var dKey = dte.componentItemId ? ("id:" + dte.componentItemId) : ("n:" + String(dte.componentItemName || "").toLowerCase().trim());
          directPerKey[dKey] = (directPerKey[dKey] || 0) + (parseFloat(dte.qty) || 0);
        }
        var directTotal = 0; Object.keys(directPerKey).forEach(function(k){ directTotal += directPerKey[k]; });

        // Per-ingredient: mixed contribution + direct must meet required
        var missing = [];
        for (var pi2 = 0; pi2 < perIngredient.length; pi2++) {
          var p = perIngredient[pi2];
          var mixedContribution = mixedTotal * (p.percentage / 100);
          var totalForItem = (directPerKey[p.key] || 0) + mixedContribution;
          var tolerance = Math.max(p.required * 0.005, 0.01);
          if (totalForItem < p.required - tolerance) {
            missing.push({ itemName: p.itemName, required: Math.round(p.required*100)/100, tagged: Math.round(totalForItem*100)/100 });
          }
        }
        if (missing.length > 0) {
          missing.forEach(function(m) {
            blendErrors.push({ blendName: rawLine.blend || ("Blend " + (lineIdx+1)), itemName: m.itemName, expected: m.required, actual: m.tagged });
          });
          continue; // skip ratio check for this line since ingredients are incomplete
        }

        // Ratio validation (only when there are direct tags — mixed tags are pre-validated at tag time)
        if (directTotal > 0) {
          for (var pi3 = 0; pi3 < perIngredient.length; pi3++) {
            var p2 = perIngredient[pi3];
            var directForItem = directPerKey[p2.key] || 0;
            var actualPct = directTotal > 0 ? (directForItem / directTotal * 100) : 0;
            if (Math.abs(actualPct - p2.percentage) > 0.5) {
              var actStr = perIngredient.map(function(px){ var v = directPerKey[px.key]||0; return (directTotal>0?Math.round(v/directTotal*100):0) + "% " + (px.itemName||""); }).join(" : ");
              var reqStr = perIngredient.map(function(px){ return px.percentage + "% " + (px.itemName||""); }).join(" : ");
              blendErrors.push({ blendName: rawLine.blend || ("Blend " + (lineIdx+1)), itemName: "(ratio)", ratioActual: actStr, ratioRequired: reqStr });
              break;
            }
          }
        }
      }
    }
  } catch (bcErr) { /* non-fatal */ }

  if (blendErrors.length > 0) {
    // BLOCK — build a descriptive error message per offender
    var errPieces = blendErrors.map(function(b) {
      if (b.ratioActual) return b.blendName + ": ratio mismatch — tagged " + b.ratioActual + " but requires " + b.ratioRequired;
      return b.blendName + ": " + b.itemName + " requires " + b.expected + "kg but only " + b.actual + "kg is tagged";
    });
    return { error: "Cannot deliver — " + errPieces.join("; "), blendErrors: blendErrors };
  }

  if (!confirmed) {
    return { preview: true, deductions: deductions, blendWarnings: blendWarnings, orderId: orderId, orderName: order.name };
  }

  // Apply deductions atomically: write all transactions
  for (var d = 0; d < deductions.length; d++) {
    var ded = deductions[d];
    if (!ded.itemId || ded.qty <= 0) continue;
    createInventoryTransaction(
      ded.itemId, "OUT", ded.qty,
      "order_delivery", orderId,
      "Order fulfilled - " + (order.name || orderId) + " [" + ded.checklistAutoId + "]",
      user.displayName || user.username
    );
    appendAuditLog("Audit - Inventory", {
      timestamp: new Date().toISOString(),
      user: user.displayName || user.username || "",
      action: "deduction_on_delivery",
      recordId: orderId,
      fieldChanged: ded.itemName || ded.itemId,
      oldValue: "",
      newValue: String(-ded.qty),
      notes: "Order fulfilled - " + (order.name || orderId) + " [" + ded.checklistAutoId + "]",
    });
  }

  order.status = "delivered";
  order.delivered_at = new Date().toISOString();
  updateSheetRow(SHEETS.ORDERS, idx, order);

  writeAuditLog(user, "deliver", "Order", orderId, "Delivered with " + deductions.length + " deductions");
  var ocs = getRows(SHEETS.ORDER_CHECKLISTS);
  return { success: true, order: formatOrder(order, ocs), deductions: deductions };
}

// ─── Master Summary Setup ──────────────────────────────────────

function setupMasterSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MASTER_SUMMARY);
  if (!sheet) sheet = ss.insertSheet(SHEETS.MASTER_SUMMARY);
  if (sheet.getLastRow() > 0) return; // Already set up

  // Row 1: Filter labels
  var labels = ["Filter by Customer:", "Filter by Person:", "Filter by Checklist:", "Filter by Order Type:", "Filter by Date From:", "Filter by Date To:"];
  sheet.getRange(1, 1, 1, labels.length).setValues([labels]).setFontWeight("bold").setFontSize(10).setBackground("#f3f3f3");

  // Row 2: Dropdown defaults
  sheet.getRange("A2").setValue("All");
  sheet.getRange("B2").setValue("All");
  sheet.getRange("C2").setValue("All");
  sheet.getRange("D2").setValue("All");
  sheet.getRange("E2").setValue("");
  sheet.getRange("F2").setValue("");
  sheet.getRange("A2:D2").setBackground("#ffffff");
  sheet.getRange("E2:F2").setBackground("#ffffff").setNumberFormat("yyyy-mm-dd");

  // Data validation: Customers
  var custValues = ["All"];
  try {
    var custSheet = ss.getSheetByName("Customers");
    if (custSheet && custSheet.getLastRow() > 1) {
      var cd = custSheet.getRange(2, 2, custSheet.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < cd.length; i++) { if (cd[i][0]) custValues.push(String(cd[i][0])); }
    }
  } catch(e) {}
  sheet.getRange("A2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(custValues).setAllowInvalid(true).build());

  // Data validation: Users/Person
  var personValues = ["All"];
  try {
    var userSheet = ss.getSheetByName("Users");
    if (userSheet && userSheet.getLastRow() > 1) {
      var ud = userSheet.getRange(2, 4, userSheet.getLastRow() - 1, 1).getValues(); // display_name col 4
      for (var j = 0; j < ud.length; j++) { if (ud[j][0]) personValues.push(String(ud[j][0])); }
    }
  } catch(e) {}
  sheet.getRange("B2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(personValues).setAllowInvalid(true).build());

  // Data validation: Checklists
  var ckValues = ["All", "Green Beans Quality Check", "Roasted Beans Quality Check", "Tagging Roasted Beans",
    "Grinding & Packing Checklist", "Sample Retention Checklist", "Coffee with Chicory Mix"];
  sheet.getRange("C2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(ckValues).setAllowInvalid(true).build());

  // Data validation: Order Types
  var otValues = ["All"];
  try {
    var otSheet = ss.getSheetByName("OrderTypes");
    if (otSheet && otSheet.getLastRow() > 1) {
      var od = otSheet.getRange(2, 2, otSheet.getLastRow() - 1, 1).getValues();
      for (var k = 0; k < od.length; k++) { if (od[k][0]) otValues.push(String(od[k][0])); }
    }
  } catch(e) {}
  sheet.getRange("D2").setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(otValues).setAllowInvalid(true).build());

  // Date validation
  sheet.getRange("E2").setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());
  sheet.getRange("F2").setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build());

  // Row 3: blank separator

  // Row 4: Data column headers
  sheet.getRange(4, 1, 1, MASTER_SUMMARY_HEADERS.length).setValues([MASTER_SUMMARY_HEADERS])
    .setFontWeight("bold").setBackground("#D4A574").setFontColor("#FFFFFF");

  sheet.setFrozenRows(4);
  for (var w = 1; w <= MASTER_SUMMARY_HEADERS.length; w++) sheet.setColumnWidth(w, 140);
}

// Canonical question array for the Roasted Beans Quality Check checklist. Defined as a
// builder function so both the seed-data path and the updateChecklistTemplates()
// re-seed path use the same source of truth — eliminates the label/index drift that
// caused "Type of Bean" to display dates and "Date of Roast" to display roast profiles.
//
// Field index contract (do not reorder without also updating dependents):
//   0: Shipment number used     (linkedSource → ck_green_beans)
//   1: Roast profile            (text)
//   2: Type of Beans            (inventory_item, autoFill from green beans QC)
//                                ← inventoryLink fieldIdx targets this position
//   3: Quantity input           (number, OUT Green Beans via field 2)
//   4: Quantity output          (number, IN Roasted Beans via field 2 → equivalent_items)
//   5: Date of Roast            (date) ← autoIdConfig.dateFieldIdx
//   6: Loss in weight
//   7: Reason for loss
//   8: How is the roasted beans stored?
//   9: Roast Approved?          (approval gate)
//  10: Remarks
function ROASTED_BEANS_CANONICAL_QUESTIONS() {
  return [
    { text: "Shipment number used", type: "text", linkedSource: { checklistId: "ck_green_beans", type: "approved_only" } },
    { text: "Roast profile", type: "text" },
    { text: "Type of Beans", type: "inventory_item", inventoryCategory: "Green Beans", autoFillMapping: { sourceFieldIdx: "1", readOnly: true } },
    { text: "Quantity input", type: "number",
      inventoryLink: { enabled: true, txType: "OUT", category: "Green Beans", itemSource: { type: "field", fieldIdx: 2, itemId: "" } } },
    { text: "Quantity output", type: "number",
      inventoryLink: { enabled: true, txType: "IN", category: "Roasted Beans", itemSource: { type: "field", fieldIdx: 2, itemId: "" } } },
    { text: "Date of Roast", type: "date" },
    { text: "Loss in weight", type: "number" },
    { text: "Reason for loss", type: "text" },
    { text: "How is the roasted beans stored?", type: "text" },
    { text: "Roast Approved?", type: "yesno", isApprovalGate: true },
    { text: "Remarks", type: "text" },
  ];
}

// One-time corrective: re-write the Roasted Beans Quality Check template stored in the
// Checklists sheet so its question array matches ROASTED_BEANS_CANONICAL_QUESTIONS().
// This fixes the bug where labels and values were misaligned ("Type of Bean" showing a
// date, "Date of Roast" showing a roast profile) due to inconsistent template definitions
// across the codebase.
//
// Preservation rules (enforced):
//   • Existing per-row response data in per-checklist response tabs is NOT modified —
//     only the template's questions JSON is rewritten.
//   • Existing inventoryLink / linkedSource / autoFillMapping configs from the canonical
//     definition are written verbatim.
//   • Returns a diff summary so admins can verify the change.
function handleFixRoastedBeansTemplateOrder(body, user) {
  var ckId = "ck_roasted_beans";
  var idx = findRowIndex(SHEETS.CHECKLISTS, ckId);
  if (idx < 0) return { error: "Roasted Beans Quality Check template not found in Checklists sheet" };
  var rows = getRows(SHEETS.CHECKLISTS);
  var existing = null;
  for (var i = 0; i < rows.length; i++) { if (String(rows[i].id) === ckId) { existing = rows[i]; break; } }
  if (!existing) return { error: "Roasted Beans Quality Check template row not found" };

  var oldQuestions = safeParseJSON(existing.questions, []);
  var oldOrder = oldQuestions.map(function(q) { return q.text || ""; });

  var canonical = ROASTED_BEANS_CANONICAL_QUESTIONS();
  var canonicalOrder = canonical.map(function(q) { return q.text; });

  // Preserve any custom questions already present in the deployed template that aren't
  // in the canonical set — append them at the end so no admin-added field is dropped.
  var canonicalTexts = {};
  for (var c = 0; c < canonical.length; c++) canonicalTexts[canonical[c].text] = true;
  for (var o = 0; o < oldQuestions.length; o++) {
    var qt = oldQuestions[o].text || "";
    if (qt && !canonicalTexts[qt]) {
      canonical.push(oldQuestions[o]);
      canonicalTexts[qt] = true;
    }
  }

  existing.questions = JSON.stringify(canonical);
  // Ensure autoIdConfig points at the right indices for the canonical layout
  var autoIdCfg = safeParseJSON(existing.auto_id_config, null) || { enabled: true, prefix: "RB" };
  autoIdCfg.itemCodeFieldIdx = 2;
  autoIdCfg.dateFieldIdx = 5;
  if (autoIdCfg.enabled === undefined) autoIdCfg.enabled = true;
  if (!autoIdCfg.prefix) autoIdCfg.prefix = "RB";
  existing.auto_id_config = JSON.stringify(autoIdCfg);
  updateSheetRow(SHEETS.CHECKLISTS, idx, existing);
  invalidateCache(SHEETS.CHECKLISTS);

  // Refresh the per-checklist response sheet header row so new submissions land in the
  // right columns. Existing data rows are NOT moved — admins must reconcile historical
  // rows manually if column positions shifted.
  try { getOrCreateResponseSheet(existing.name, canonical); } catch (e) { /* non-fatal */ }

  writeAuditLog(user, "fix_template_order", "Checklist", ckId, "Roasted Beans QC question order normalized");
  return {
    success: true,
    checklistId: ckId,
    oldOrder: oldOrder,
    newOrder: canonical.map(function(q) { return q.text; }),
    note: "Template question order has been normalized. Per-checklist response tab headers updated. Historical response rows were not moved — please review manually if old labels/data were misaligned."
  };
}

// ─── Update Checklist Templates (run once manually) ──────────
// Overwrites ALL checklist question configs in the Checklists sheet
// with the correct linked source, approval gate, and linked ID settings.
// Safe to re-run — it uses upsert logic (update if exists, insert if not).

function updateChecklistTemplates() {
  var templateMap = {
    "ck_sample_qc": {
      name: "Green Bean QC Sample Check", subtitle: "Per sample received", form_url: "",
      questions: [
        { text: "Supplier/Origin", type: "text" },
        { text: "Type of Beans", type: "text" },
        { text: "Sample Quantity", type: "number" },
        { text: "Date Received", type: "date" },
        { text: "Visual Inspection Notes", type: "text" },
        { text: "Cupping Score", type: "number" },
        { text: "Sample Approved?", type: "yesno", isApprovalGate: true },
        { text: "Remarks", type: "text" },
      ],
    },
    "ck_green_beans": {
      name: "Green Beans Quality Check", subtitle: "Per green bean shipment", form_url: "",
      questions: [
        // "Sample Code Used" removed — Auto ID from source is shown in the linked dropdown
        { text: "Source Sample", type: "text", linkedSource: { checklistId: "ck_sample_qc", type: "approved_only" } },
        { text: "Type of Beans", type: "text", autoFillMapping: { sourceFieldIdx: "1", readOnly: true } },
        { text: "Quantity received", type: "number" },
        { text: "Bags stored in which location", type: "text" },
        { text: "Shipment Approved?", type: "yesno", isApprovalGate: true },
        { text: "Remarks", type: "text" },
      ],
    },
    "ck_roasted_beans": {
      name: "Roasted Beans Quality Check", subtitle: "Per roast batch", form_url: "",
      // Source of truth lives in ROASTED_BEANS_CANONICAL_QUESTIONS() — kept consistent
      // with the seed-data path so labels and field indices never drift.
      questions: ROASTED_BEANS_CANONICAL_QUESTIONS(),
    },
    "ck_tagging": {
      name: "Tagging Roasted Beans", subtitle: "Per roast batch tagging", form_url: "",
      questions: [
        { text: "Roast tag number", type: "text" },
        { text: "Grade", type: "text" },
        { text: "Quantity", type: "number" },
        { text: "Date of roast", type: "date" },
      ],
    },
    "ck_grinding": {
      name: "Grinding & Packing Checklist", subtitle: "For each packing lot", form_url: "",
      questions: [
        { text: "Roast ID", type: "text", linkedSource: { checklistId: "ck_roasted_beans", type: "approved_only" } },
        { text: "Invoice/SO", type: "text" },
        { text: "Client name", type: "text" },
        { text: "Type of Bean", type: "text", autoFillMapping: { sourceFieldIdx: "1", readOnly: true } },
        { text: "Grind size", type: "text" },
        { text: "Is the correct stickers applied", type: "yesno" },
        { text: "Total Net weight", type: "number" },
        { text: "Is the Net total weight matching the order", type: "yesno" },
        { text: "Is the package labelled for the right customer", type: "yesno" },
        { text: "Stickers have been properly attached", type: "yesno" },
        { text: "Did we keep a sample with us for reference", type: "yesno" },
        { text: "Sample quantity", type: "number" },
        { text: "Sample code", type: "text" },
      ],
    },
    "ck_sample_retention": {
      name: "Sample Retention Checklist", subtitle: "Each new sample", form_url: "",
      questions: [
        { text: "Sample Code", type: "text" },
        { text: "Client name", type: "text" },
        { text: "Grind size", type: "text" },
        { text: "Grinder used", type: "text" },
        { text: "Roast IDs (Multiple)", type: "text" },
        { text: "Sample quantity", type: "number" },
        { text: "Date of Roast", type: "date" },
        { text: "Date of Grind", type: "date" },
        { text: "Approved sample?", type: "yesno" },
      ],
    },
    "ck_chicory_mix": {
      name: "Coffee with Chicory Mix", subtitle: "Per chicory mix lot", form_url: "",
      questions: [
        { text: "Client name", type: "text" },
        { text: "Invoice/SO", type: "text" },
        { text: "Weight of coffee", type: "number" },
        { text: "Weight of chicory", type: "number" },
        { text: "Expected ratio", type: "text" },
        { text: "Have we checked the right proportions of coffee and chicory before mixing", type: "yesno" },
        { text: "Has the mixing of coffee and chicory been done completely", type: "yesno" },
        { text: "Have we cross checked the taste and aroma of the lot vs sample", type: "yesno" },
        { text: "Sample ID used for lot", type: "text" },
        { text: "Weight of the sample", type: "number" },
        { text: "Result of the sample test", type: "text" },
        { text: "If not satisfied reasons", type: "text" },
        { text: "Rectification action", type: "text" },
        { text: "Whether the taste and aroma is matching now?", type: "yesno" },
        { text: "Weight of packed coffee", type: "number" },
      ],
    },
  };

  var ids = Object.keys(templateMap);
  for (var i = 0; i < ids.length; i++) {
    var id = ids[i];
    var def = templateMap[id];
    var questions = normalizeQuestions(def.questions);
    var idx = findRowIndex(SHEETS.CHECKLISTS, id);
    if (idx > 0) {
      // Update existing
      var obj = { id: id, name: def.name, subtitle: def.subtitle || "", form_url: def.form_url || "", questions: JSON.stringify(questions) };
      updateSheetRow(SHEETS.CHECKLISTS, idx, obj);
      Logger.log("Updated: " + def.name + " (" + id + ")");
    } else {
      // Insert new
      var obj2 = { id: id, name: def.name, subtitle: def.subtitle || "", form_url: def.form_url || "", questions: JSON.stringify(questions) };
      appendToSheet(SHEETS.CHECKLISTS, obj2);
      Logger.log("Inserted: " + def.name + " (" + id + ")");
    }
    // Ensure response sheet exists
    getOrCreateResponseSheet(def.name, questions);
  }

  // ── FIX 1: Remove duplicate checklists (keep canonical IDs only) ──
  var allCkRows = getRows(SHEETS.CHECKLISTS);
  var nameMap = {};
  for (var d = 0; d < allCkRows.length; d++) {
    var lowerName = String(allCkRows[d].name).toLowerCase().trim();
    if (!nameMap[lowerName]) nameMap[lowerName] = [];
    nameMap[lowerName].push(allCkRows[d]);
  }
  var canonicalIds = Object.keys(templateMap);
  for (var dupName in nameMap) {
    if (nameMap[dupName].length > 1) {
      // Keep the one with a canonical ID, delete the rest
      for (var dd = 0; dd < nameMap[dupName].length; dd++) {
        var dupRow = nameMap[dupName][dd];
        if (canonicalIds.indexOf(String(dupRow.id)) < 0) {
          var dupIdx = findRowIndex(SHEETS.CHECKLISTS, dupRow.id);
          if (dupIdx > 0) {
            deleteSheetRow(SHEETS.CHECKLISTS, dupIdx);
            Logger.log("Deleted duplicate checklist: " + dupRow.name + " (" + dupRow.id + ")");
            // Also clean up any assignment rules referencing the deleted ID
            var dupRules = getRows(SHEETS.RULES);
            for (var dr = 0; dr < dupRules.length; dr++) {
              var rCkIds = safeParseJSON(dupRules[dr].checklist_ids, []);
              var filtered = rCkIds.filter(function(x) { return x !== dupRow.id; });
              // Replace deleted ID with canonical ID if there's one with same name
              var canonical = nameMap[dupName].find(function(c) { return canonicalIds.indexOf(String(c.id)) >= 0; });
              if (canonical && filtered.indexOf(canonical.id) < 0 && rCkIds.indexOf(dupRow.id) >= 0) filtered.push(canonical.id);
              if (filtered.length !== rCkIds.length || (rCkIds.indexOf(dupRow.id) >= 0)) {
                var rIdx = findRowIndex(SHEETS.RULES, dupRules[dr].id);
                if (rIdx > 0) { dupRules[dr].checklist_ids = JSON.stringify(filtered); updateSheetRow(SHEETS.RULES, rIdx, dupRules[dr]); }
              }
            }
          }
        }
      }
    }
  }

  Logger.log("All checklist templates updated with correct linking config.");
}

// ─── Clean Up Users (run once manually) ──────────────────────
// Deletes ALL users except admin (user_admin)

function cleanupUsers() {
  var sheet = getSheet(SHEETS.USERS);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) !== "user_admin") {
      sheet.deleteRow(i + 1);
    }
  }
  // Also clean up sessions for deleted users
  var sessSheet = getSheet(SHEETS.SESSIONS);
  var sessData = sessSheet.getDataRange().getValues();
  for (var j = sessData.length - 1; j >= 1; j--) {
    if (String(sessData[j][1]) !== "user_admin") {
      sessSheet.deleteRow(j + 1);
    }
  }
  invalidateCache(SHEETS.USERS);
  invalidateCache(SHEETS.SESSIONS);
  Logger.log("Cleaned up all non-admin users and their sessions.");
}

// ─── Seed Data (run once manually) ─────────────────────────────

// ═══════════════════════════════════════════════════════════════
// SAFE SEED: only runs if sheet is empty, never overwrites existing data.
// Each section checks if the sheet already has data rows (beyond
// the header).  If ANY data rows exist the section is skipped
// entirely — no overwriting, no duplicate appends.
// ═══════════════════════════════════════════════════════════════
function seedData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Helper: returns true when a sheet has NO data rows (header-only or empty)
  function sheetIsEmpty(sheet) { return !sheet || sheet.getLastRow() <= 1; }

  // Create standard sheets with headers (safe — only adds header if missing)
  var sheetNames = Object.keys(HEADERS);
  for (var i = 0; i < sheetNames.length; i++) {
    var name = sheetNames[i];
    var sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS[name]);
      sheet.getRange(1, 1, 1, HEADERS[name].length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  }

  // SAFE SEED: Config — only if empty
  var configSheet = getSheet(SHEETS.CONFIG);
  if (sheetIsEmpty(configSheet)) {
    appendToSheet(SHEETS.CONFIG, { key: "next_order_number", value: 1 });
  }

  // SAFE SEED: Users — only if empty
  var usersSheet = getSheet(SHEETS.USERS);
  if (sheetIsEmpty(usersSheet)) {
    appendToSheet(SHEETS.USERS, {
      id: "user_admin", username: "admin", password_hash: hashPassword("admin123"),
      display_name: "Administrator", role: "admin", status: "active",
      created_at: new Date().toISOString(), created_by: "system",
    });
  }

  // SAFE SEED: Order Types — only if empty
  var otSheet = getSheet(SHEETS.ORDER_TYPES);
  if (sheetIsEmpty(otSheet)) {
    [{ id: "rg_coffee", label: "Roast & Ground Coffee" },
     { id: "rg_chicory", label: "Coffee with Chicory Mix" },
     { id: "green_beans", label: "Green Beans" },
     { id: "roasted_beans", label: "Roasted Beans" },
     { id: "general", label: "General Order" }]
    .forEach(function(ot) { appendToSheet(SHEETS.ORDER_TYPES, ot); });
  }

  // SAFE SEED: Customers — only if empty
  var custSheet = getSheet(SHEETS.CUSTOMERS);
  if (sheetIsEmpty(custSheet)) {
    appendToSheet(SHEETS.CUSTOMERS, { id: "cust_default", label: "Default / Walk-in" });
  }

  // SAFE SEED: Checklists — only if empty. Also creates per-checklist response tabs.
  var ckSheet = getSheet(SHEETS.CHECKLISTS);
  var checklistDefs = [
    { id: "ck_sample_qc", name: "Green Bean QC Sample Check", subtitle: "Per sample received", form_url: "",
      autoIdConfig: { enabled: true, prefix: "GBS", dateFieldIdx: 3, itemCodeFieldIdx: 1 },
      questions: [
        { text: "Supplier/Origin", type: "text" },
        { text: "Type of Beans", type: "inventory_item", inventoryCategory: "Green Beans" },
        { text: "Sample Quantity", type: "number",
          inventoryLink: { enabled: true, txType: "IN", category: "Green Beans", itemSource: { type: "field", fieldIdx: 1, itemId: "" } } },
        { text: "Date Received", type: "date" },
        { text: "Visual Inspection Notes", type: "text" },
        { text: "Cupping Score", type: "number" },
        { text: "Sample Approved?", type: "yesno", isApprovalGate: true },
        { text: "Remarks", type: "text" },
      ] },
    { id: "ck_green_beans", name: "Green Beans Quality Check", subtitle: "Per green bean shipment", form_url: "",
      autoIdConfig: { enabled: true, prefix: "GB", dateFieldIdx: null, itemCodeFieldIdx: 1 },
      questions: [
        { text: "Source Sample", type: "text", linkedSource: { checklistId: "ck_sample_qc", type: "approved_only" } },
        { text: "Type of Beans", type: "inventory_item", inventoryCategory: "Green Beans", autoFillMapping: { sourceFieldIdx: "1", readOnly: true } },
        { text: "Quantity received", type: "number",
          inventoryLink: { enabled: true, txType: "IN", category: "Green Beans", itemSource: { type: "field", fieldIdx: 1, itemId: "" } } },
        { text: "Bags stored in which location", type: "text" },
        { text: "Shipment Approved?", type: "yesno", isApprovalGate: true },
        { text: "Remarks", type: "text" },
      ] },
    { id: "ck_roasted_beans", name: "Roasted Beans Quality Check", subtitle: "Per roast batch", form_url: "",
      autoIdConfig: { enabled: true, prefix: "RB", dateFieldIdx: 5, itemCodeFieldIdx: 2 },
      questions: ROASTED_BEANS_CANONICAL_QUESTIONS() },
    { id: "ck_tagging", name: "Tagging Roasted Beans", subtitle: "Per roast batch tagging", form_url: "",
      questions: [
        { text: "Roast tag number", type: "text" },
        { text: "Grade", type: "text" },
        { text: "Quantity", type: "number" },
        { text: "Date of roast", type: "date" },
      ] },
    { id: "ck_grinding", name: "Grinding & Packing Checklist", subtitle: "For each packing lot", form_url: "",
      questions: [
        { text: "Roast ID", type: "text", linkedSource: { checklistId: "ck_roasted_beans", type: "approved_only" } },
        { text: "Invoice/SO", type: "text" },
        { text: "Client name", type: "text" },
        { text: "Grind size", type: "text" },
        { text: "Is the correct stickers applied", type: "yesno" },
        { text: "Total Net weight", type: "number" },
        { text: "Is the Net total weight matching the order", type: "yesno" },
        { text: "Is the package labelled for the right customer", type: "yesno" },
        { text: "Stickers have been properly attached", type: "yesno" },
        { text: "Did we keep a sample with us for reference", type: "yesno" },
        { text: "Sample quantity", type: "number" },
        { text: "Sample code", type: "text" },
      ] },
    { id: "ck_sample_retention", name: "Sample Retention Checklist", subtitle: "Each new sample", form_url: "",
      questions: [
        { text: "Sample Code", type: "text" },
        { text: "Client name", type: "text" },
        { text: "Grind size", type: "text" },
        { text: "Grinder used", type: "text" },
        { text: "Roast IDs (Multiple)", type: "text" },
        { text: "Sample quantity", type: "number" },
        { text: "Date of Roast", type: "date" },
        { text: "Date of Grind", type: "date" },
        { text: "Approved sample?", type: "yesno" },
      ] },
    { id: "ck_chicory_mix", name: "Coffee with Chicory Mix", subtitle: "Per chicory mix lot", form_url: "",
      questions: [
        { text: "Client name", type: "text" },
        { text: "Invoice/SO", type: "text" },
        { text: "Weight of coffee", type: "number" },
        { text: "Weight of chicory", type: "number" },
        { text: "Expected ratio", type: "text" },
        { text: "Have we checked the right proportions of coffee and chicory before mixing", type: "yesno" },
        { text: "Has the mixing of coffee and chicory been done completely", type: "yesno" },
        { text: "Have we cross checked the taste and aroma of the lot vs sample", type: "yesno" },
        { text: "Sample ID used for lot", type: "text" },
        { text: "Weight of the sample", type: "number" },
        { text: "Result of the sample test", type: "text" },
        { text: "If not satisfied reasons", type: "text" },
        { text: "Rectification action", type: "text" },
        { text: "Whether the taste and aroma is matching now?", type: "yesno" },
        { text: "Weight of packed coffee", type: "number" },
      ] },
  ];

  if (sheetIsEmpty(ckSheet)) {
    for (var c = 0; c < checklistDefs.length; c++) {
      var def = checklistDefs[c];
      appendToSheet(SHEETS.CHECKLISTS, {
        id: def.id, name: def.name, subtitle: def.subtitle, form_url: def.form_url,
        questions: JSON.stringify(def.questions),
        auto_id_config: def.autoIdConfig ? JSON.stringify(def.autoIdConfig) : "",
      });
    }
  }

  // Create per-checklist response tabs (always ensure they exist)
  for (var t = 0; t < checklistDefs.length; t++) {
    getOrCreateResponseSheet(checklistDefs[t].name, checklistDefs[t].questions);
  }

  // Create Master Summary tab
  setupMasterSummary();

  // SAFE SEED: Assignment Rules — only if empty
  var rulesSheet = getSheet(SHEETS.RULES);
  if (sheetIsEmpty(rulesSheet)) {
    [
      { id: "rule_1", order_type_id: "rg_coffee", customer_id: "any", checklist_ids: JSON.stringify(["ck_sample_qc", "ck_green_beans", "ck_roasted_beans", "ck_tagging", "ck_grinding", "ck_sample_retention"]) },
      { id: "rule_2", order_type_id: "rg_chicory", customer_id: "any", checklist_ids: JSON.stringify(["ck_sample_qc", "ck_green_beans", "ck_roasted_beans", "ck_tagging", "ck_grinding", "ck_chicory_mix", "ck_sample_retention"]) },
      { id: "rule_3", order_type_id: "green_beans", customer_id: "any", checklist_ids: JSON.stringify(["ck_sample_qc", "ck_green_beans"]) },
      { id: "rule_4", order_type_id: "roasted_beans", customer_id: "any", checklist_ids: JSON.stringify(["ck_roasted_beans", "ck_tagging"]) },
      { id: "rule_5", order_type_id: "general", customer_id: "any", checklist_ids: JSON.stringify(["ck_grinding"]) },
    ].forEach(function(r) { appendToSheet(SHEETS.RULES, r); });
  }

  // SAFE SEED: Inventory Categories — only if empty
  var icatSheet = getSheet(SHEETS.INVENTORY_CATEGORIES);
  if (sheetIsEmpty(icatSheet)) {
    [{ id: "icat_green", name: "Green Beans" },
     { id: "icat_roasted", name: "Roasted Beans" },
     { id: "icat_packing", name: "Packing Items" },
     { id: "icat_others", name: "Others" }]
    .forEach(function(c) { appendToSheet(SHEETS.INVENTORY_CATEGORIES, c); });
  }

  // SAFE SEED: Inventory Items — only if empty
  var invSheet = getSheet(SHEETS.INVENTORY_ITEMS);
  if (sheetIsEmpty(invSheet)) {
    [
      { id: "inv_acab",   category: "Green Beans",   name: "Arabica Cherry AB", unit: "kg", opening_stock: 0, current_stock: 0, min_stock_alert: 0, created_at: new Date().toISOString(), is_active: "true", abbreviation: "ACAB",
        equivalent_items: JSON.stringify([{ category: "Roasted Beans", itemId: "inv_acab_r" }, { category: "Packing Items", itemId: "inv_acab_g" }]) },
      { id: "inv_rcpb",   category: "Green Beans",   name: "Robusta Cherry PB", unit: "kg", opening_stock: 0, current_stock: 0, min_stock_alert: 0, created_at: new Date().toISOString(), is_active: "true", abbreviation: "RCPB",
        equivalent_items: JSON.stringify([{ category: "Roasted Beans", itemId: "inv_rcpb_r" }, { category: "Packing Items", itemId: "inv_rcpb_g" }]) },
      { id: "inv_acab_r", category: "Roasted Beans", name: "Arabica Cherry AB (Roasted)", unit: "kg", opening_stock: 0, current_stock: 0, min_stock_alert: 0, created_at: new Date().toISOString(), is_active: "true", abbreviation: "ACABR",
        equivalent_items: JSON.stringify([{ category: "Green Beans", itemId: "inv_acab" }, { category: "Packing Items", itemId: "inv_acab_g" }]) },
      { id: "inv_rcpb_r", category: "Roasted Beans", name: "Robusta Cherry PB (Roasted)", unit: "kg", opening_stock: 0, current_stock: 0, min_stock_alert: 0, created_at: new Date().toISOString(), is_active: "true", abbreviation: "RCPBR",
        equivalent_items: JSON.stringify([{ category: "Green Beans", itemId: "inv_rcpb" }, { category: "Packing Items", itemId: "inv_rcpb_g" }]) },
      { id: "inv_acab_g", category: "Packing Items", name: "Arabica Cherry AB Ground", unit: "kg", opening_stock: 0, current_stock: 0, min_stock_alert: 0, created_at: new Date().toISOString(), is_active: "true", abbreviation: "ACABG",
        equivalent_items: JSON.stringify([{ category: "Green Beans", itemId: "inv_acab" }, { category: "Roasted Beans", itemId: "inv_acab_r" }]) },
      { id: "inv_rcpb_g", category: "Packing Items", name: "Robusta Cherry PB Ground", unit: "kg", opening_stock: 0, current_stock: 0, min_stock_alert: 0, created_at: new Date().toISOString(), is_active: "true", abbreviation: "RCPBG",
        equivalent_items: JSON.stringify([{ category: "Green Beans", itemId: "inv_rcpb" }, { category: "Roasted Beans", itemId: "inv_rcpb_r" }]) },
      { id: "inv_pouch250", category: "Packing Items", name: "250g Pouch", unit: "pieces", opening_stock: 0, current_stock: 0, min_stock_alert: 50, created_at: new Date().toISOString(), is_active: "true", abbreviation: "P250", equivalent_items: "[]" },
      { id: "inv_pouch500", category: "Packing Items", name: "500g Pouch", unit: "pieces", opening_stock: 0, current_stock: 0, min_stock_alert: 50, created_at: new Date().toISOString(), is_active: "true", abbreviation: "P500", equivalent_items: "[]" },
      { id: "inv_bag1kg",   category: "Packing Items", name: "1kg Bag",    unit: "pieces", opening_stock: 0, current_stock: 0, min_stock_alert: 25, created_at: new Date().toISOString(), is_active: "true", abbreviation: "B1KG", equivalent_items: "[]" },
    ].forEach(function(item) { appendToSheet(SHEETS.INVENTORY_ITEMS, item); });
  }

  // Ensure InventoryLedger sheet exists
  getSheet(SHEETS.INVENTORY_LEDGER);

  // SAFE SEED: Blends — only if empty
  var blendsSheet = getSheet(SHEETS.BLENDS);
  if (sheetIsEmpty(blendsSheet)) {
    var nowIso = new Date().toISOString();
    [
      { id: "blend_pure_arabica", name: "Pure Arabica", customer: "General", description: "100% Arabica roasted coffee",
        components: JSON.stringify([{ category: "Roasted Beans", itemId: "", itemName: "Any Arabica", percentage: 100 }]),
        is_active: "true", created_at: nowIso },
      { id: "blend_classic_80_20", name: "Classic 80-20", customer: "General", description: "80% Roasted Coffee + 20% Chicory",
        components: JSON.stringify([{ category: "Roasted Beans", itemId: "", itemName: "Roasted Coffee", percentage: 80 }, { category: "Others", itemId: "", itemName: "Chicory", percentage: 20 }]),
        is_active: "true", created_at: nowIso },
    ].forEach(function(b) { appendToSheet(SHEETS.BLENDS, b); });
  }

  // SAFE SEED: Roast Classifications — only if empty
  var rcSheet = getSheet(SHEETS.ROAST_CLASSIFICATIONS);
  if (sheetIsEmpty(rcSheet)) {
    var rcNow = new Date().toISOString();
    [
      { id: "rc_light", name: "Light Roast", type: "roast_degree", description: "Light brown, no oil on surface", created_by: "system", created_at: rcNow, updated_at: rcNow, is_active: "true" },
      { id: "rc_medium", name: "Medium Roast", type: "roast_degree", description: "Medium brown, balanced flavor", created_by: "system", created_at: rcNow, updated_at: rcNow, is_active: "true" },
      { id: "rc_dark", name: "Dark Roast", type: "roast_degree", description: "Dark brown, oily surface", created_by: "system", created_at: rcNow, updated_at: rcNow, is_active: "true" },
      { id: "rc_fine", name: "Fine", type: "grind_size", description: "Fine grind for espresso", created_by: "system", created_at: rcNow, updated_at: rcNow, is_active: "true" },
      { id: "rc_medium_grind", name: "Medium", type: "grind_size", description: "Medium grind for drip/pour-over", created_by: "system", created_at: rcNow, updated_at: rcNow, is_active: "true" },
      { id: "rc_coarse", name: "Coarse", type: "grind_size", description: "Coarse grind for French press", created_by: "system", created_at: rcNow, updated_at: rcNow, is_active: "true" },
    ].forEach(function(c) { appendToSheet(SHEETS.ROAST_CLASSIFICATIONS, c); });
  }

  // Delete the default "Sheet1" if it exists and is unused
  var sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1 && ss.getSheets().length > 1) {
    try { ss.deleteSheet(sheet1); } catch(e) {}
  }

  Logger.log("Seed data complete! All tabs created and populated.");
}

// ═══════════════════════════════════════════════════════════════
// ─── MIGRATION: Run this manually after restoring old sheet ───
// ═══════════════════════════════════════════════════════════════
//
// migrateData() reads your existing old-format data and creates
// the new structure alongside it. It does NOT delete any old tabs.
//
// Steps to use:
//   1. Restore your old Google Sheet from version history
//   2. Paste this entire new code into Apps Script
//   3. Run migrateData() from the editor
//   4. Verify the new tabs look correct
//   5. Optionally run cleanupOldTabs() to remove old ChecklistResponses tab
//
// ═══════════════════════════════════════════════════════════════

function migrateData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = [];

  log.push("=== Migration started at " + new Date().toISOString() + " ===");

  // ── Step 0: Ensure new infrastructure tabs exist ──────────────
  // Create Config tab if missing
  var configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Config");
    configSheet.appendRow(["key", "value"]);
    configSheet.getRange(1, 1, 1, 2).setFontWeight("bold");
    configSheet.setFrozenRows(1);
    log.push("Created Config tab");
  }

  // Create ArchivedResponses tab if missing (new name)
  var archivedRespSheet = ss.getSheetByName("ArchivedResponses");
  if (!archivedRespSheet) {
    archivedRespSheet = ss.insertSheet("ArchivedResponses");
    archivedRespSheet.appendRow(HEADERS.ArchivedResponses);
    archivedRespSheet.getRange(1, 1, 1, HEADERS.ArchivedResponses.length).setFontWeight("bold");
    archivedRespSheet.setFrozenRows(1);
    log.push("Created ArchivedResponses tab");
  }

  // ── Step 1: Read all existing data from old tabs ─────────────
  log.push("\n--- Reading existing data ---");

  var ordersSheet = ss.getSheetByName("Orders");
  var ocSheet = ss.getSheetByName("OrderChecklists");
  var oldResponsesSheet = ss.getSheetByName("ChecklistResponses");
  var checklistsSheet = ss.getSheetByName("Checklists");
  var customersSheet = ss.getSheetByName("Customers");
  var orderTypesSheet = ss.getSheetByName("OrderTypes");

  if (!ordersSheet) { log.push("ERROR: Orders tab not found!"); Logger.log(log.join("\n")); return; }
  if (!ocSheet) { log.push("ERROR: OrderChecklists tab not found!"); Logger.log(log.join("\n")); return; }
  if (!checklistsSheet) { log.push("ERROR: Checklists tab not found!"); Logger.log(log.join("\n")); return; }

  // Read orders
  var ordersData = readSheetAsObjects_(ordersSheet);
  log.push("Found " + ordersData.length + " orders");

  // Read order checklists
  var ocsData = readSheetAsObjects_(ocSheet);
  log.push("Found " + ocsData.length + " order checklists");

  // Read old responses (may not exist if fresh)
  var oldResponses = [];
  if (oldResponsesSheet && oldResponsesSheet.getLastRow() > 1) {
    oldResponses = readSheetAsObjects_(oldResponsesSheet);
  }
  log.push("Found " + oldResponses.length + " old response rows");

  // Read checklists for name/question lookup
  var checklistRows = readSheetAsObjects_(checklistsSheet);
  var checklistMap = {}; // id -> { name, questions[] }
  for (var i = 0; i < checklistRows.length; i++) {
    var cr = checklistRows[i];
    checklistMap[String(cr.id)] = {
      name: String(cr.name),
      questions: safeParseJSON(cr.questions, []),
    };
  }
  log.push("Found " + checklistRows.length + " checklist templates");

  // Read customers for label lookup
  var customerMap = {}; // id -> label
  if (customersSheet && customersSheet.getLastRow() > 1) {
    var custRows = readSheetAsObjects_(customersSheet);
    for (var c = 0; c < custRows.length; c++) customerMap[String(custRows[c].id)] = String(custRows[c].label);
  }

  // Read order types for label lookup
  var orderTypeMap = {}; // id -> label
  if (orderTypesSheet && orderTypesSheet.getLastRow() > 1) {
    var otRows = readSheetAsObjects_(orderTypesSheet);
    for (var t = 0; t < otRows.length; t++) orderTypeMap[String(otRows[t].id)] = String(otRows[t].label);
  }

  // ── Step 2: Build ID mapping (old ord_xxx -> new ORD-001) ────
  log.push("\n--- Building order ID mapping ---");

  // Sort orders by created_at so ORD numbers are chronological
  ordersData.sort(function(a, b) {
    return String(a.created_at).localeCompare(String(b.created_at));
  });

  var idMap = {}; // old id -> new ORD-xxx id
  var nextNum = 1;

  // Check if Config already has a counter (in case of partial migration)
  var configData = configSheet.getDataRange().getValues();
  for (var ci = 1; ci < configData.length; ci++) {
    if (String(configData[ci][0]) === "next_order_number") {
      nextNum = parseInt(configData[ci][1]) || 1;
      break;
    }
  }

  for (var oi = 0; oi < ordersData.length; oi++) {
    var oldId = String(ordersData[oi].id);
    // Skip if already in ORD-xxx format (already migrated)
    if (oldId.indexOf("ORD-") === 0) {
      idMap[oldId] = oldId;
      var existingNum = parseInt(oldId.replace("ORD-", "")) || 0;
      if (existingNum >= nextNum) nextNum = existingNum + 1;
    } else {
      var newId = "ORD-" + String(nextNum).padStart(3, "0");
      idMap[oldId] = newId;
      nextNum++;
    }
  }

  log.push("Mapped " + Object.keys(idMap).length + " order IDs");
  log.push("Next order number will be: " + nextNum);

  // ── Step 3: Update Orders sheet with new IDs ─────────────────
  log.push("\n--- Updating Orders with new IDs ---");

  var ordersHeaders = ordersSheet.getDataRange().getValues()[0];
  var idColIdx = ordersHeaders.indexOf("id");
  if (idColIdx < 0) { log.push("ERROR: 'id' column not found in Orders!"); Logger.log(log.join("\n")); return; }

  var ordersAllData = ordersSheet.getDataRange().getValues();
  var ordersUpdated = 0;
  for (var r = 1; r < ordersAllData.length; r++) {
    var currentId = String(ordersAllData[r][idColIdx]);
    if (idMap[currentId] && idMap[currentId] !== currentId) {
      ordersSheet.getRange(r + 1, idColIdx + 1).setValue(idMap[currentId]);
      ordersUpdated++;
    }
  }
  log.push("Updated " + ordersUpdated + " order IDs in Orders tab");

  // ── Step 4: Update OrderChecklists with new order IDs ────────
  log.push("\n--- Updating OrderChecklists with new order IDs ---");

  var ocHeaders = ocSheet.getDataRange().getValues()[0];
  var ocOrderIdColIdx = ocHeaders.indexOf("order_id");
  if (ocOrderIdColIdx < 0) { log.push("ERROR: 'order_id' column not found in OrderChecklists!"); Logger.log(log.join("\n")); return; }

  var ocAllData = ocSheet.getDataRange().getValues();
  var ocsUpdated = 0;
  for (var ocr = 1; ocr < ocAllData.length; ocr++) {
    var ocOldOrderId = String(ocAllData[ocr][ocOrderIdColIdx]);
    if (idMap[ocOldOrderId] && idMap[ocOldOrderId] !== ocOldOrderId) {
      ocSheet.getRange(ocr + 1, ocOrderIdColIdx + 1).setValue(idMap[ocOldOrderId]);
      ocsUpdated++;
    }
  }
  log.push("Updated " + ocsUpdated + " order_id references in OrderChecklists tab");

  // ── Step 5: Save next_order_number to Config ─────────────────
  var configFound = false;
  var cfgData = configSheet.getDataRange().getValues();
  for (var cfgi = 1; cfgi < cfgData.length; cfgi++) {
    if (String(cfgData[cfgi][0]) === "next_order_number") {
      configSheet.getRange(cfgi + 1, 2).setValue(nextNum);
      configFound = true;
      break;
    }
  }
  if (!configFound) {
    configSheet.appendRow(["next_order_number", nextNum]);
  }
  log.push("Saved next_order_number = " + nextNum + " to Config tab");

  // ── Step 6: Create per-checklist response tabs ───────────────
  log.push("\n--- Creating per-checklist response tabs ---");

  var createdTabs = [];
  var checklistIds = Object.keys(checklistMap);
  for (var cki = 0; cki < checklistIds.length; cki++) {
    var ckId = checklistIds[cki];
    var ck = checklistMap[ckId];
    if (ck.questions.length === 0) {
      log.push("  Skipping '" + ck.name + "' (no questions)");
      continue;
    }
    var existingTab = ss.getSheetByName(ck.name);
    if (existingTab) {
      log.push("  Tab '" + ck.name + "' already exists (skipping creation)");
    } else {
      getOrCreateResponseSheet(ck.name, ck.questions);
      log.push("  Created tab: '" + ck.name + "'");
    }
    createdTabs.push(ck.name);
  }

  // ── Step 7: Migrate response data to per-checklist tabs ──────
  log.push("\n--- Migrating response data ---");

  if (oldResponses.length === 0) {
    log.push("No old responses to migrate.");
  } else {
    // Group old responses by order_checklist_id
    var grouped = {};
    for (var ri = 0; ri < oldResponses.length; ri++) {
      var resp = oldResponses[ri];
      var ocId = String(resp.order_checklist_id);
      if (!grouped[ocId]) {
        grouped[ocId] = {
          orderChecklistId: ocId,
          orderName: String(resp.order_name || ""),
          customer: String(resp.customer || ""),
          checklistName: String(resp.checklist_name || ""),
          person: String(resp.person || ""),
          date: String(resp.date || ""),
          submittedAt: String(resp.submitted_at || ""),
          responses: [],
        };
      }
      grouped[ocId].responses.push({
        question: String(resp.question || ""),
        response: String(resp.response || ""),
      });
    }

    var groupedKeys = Object.keys(grouped);
    log.push("Found " + groupedKeys.length + " unique checklist submissions to migrate");

    // Re-read OC data (now with updated order_id values)
    var freshOcData = readSheetAsObjects_(ocSheet);
    var ocLookup = {}; // oc_id -> { order_id, checklist_id }
    for (var fi = 0; fi < freshOcData.length; fi++) {
      ocLookup[String(freshOcData[fi].id)] = {
        order_id: String(freshOcData[fi].order_id),
        checklist_id: String(freshOcData[fi].checklist_id),
      };
    }

    // Re-read orders (now with updated IDs)
    var freshOrders = readSheetAsObjects_(ordersSheet);
    var orderLookup = {}; // order_id -> { name, customer_id, order_type }
    for (var foi = 0; foi < freshOrders.length; foi++) {
      orderLookup[String(freshOrders[foi].id)] = {
        name: String(freshOrders[foi].name),
        customer_id: String(freshOrders[foi].customer_id),
        order_type: String(freshOrders[foi].order_type),
      };
    }

    var migratedCount = 0;
    var skippedCount = 0;

    for (var gi = 0; gi < groupedKeys.length; gi++) {
      var entry = grouped[groupedKeys[gi]];
      var ocInfo = ocLookup[entry.orderChecklistId];

      // Determine checklist name and questions
      var ckName = entry.checklistName;
      var ckQuestions = [];
      if (ocInfo && checklistMap[ocInfo.checklist_id]) {
        ckName = checklistMap[ocInfo.checklist_id].name;
        ckQuestions = checklistMap[ocInfo.checklist_id].questions;
      } else {
        // Try to find by name match
        for (var cid in checklistMap) {
          if (checklistMap[cid].name === ckName) {
            ckQuestions = checklistMap[cid].questions;
            break;
          }
        }
      }

      if (!ckName || ckQuestions.length === 0) {
        log.push("  SKIP: No checklist found for OC " + entry.orderChecklistId);
        skippedCount++;
        continue;
      }

      // Check if this tab exists
      var targetSheet = ss.getSheetByName(ckName);
      if (!targetSheet) {
        getOrCreateResponseSheet(ckName, ckQuestions);
        targetSheet = ss.getSheetByName(ckName);
      }

      // Determine the new order ID
      var newOrderId = "";
      if (ocInfo) {
        newOrderId = ocInfo.order_id; // Already updated to ORD-xxx
      }

      // Check for duplicate: skip if this order already has a row in this tab
      var existingData = targetSheet.getDataRange().getValues();
      var alreadyExists = false;
      for (var de = 2; de < existingData.length; de++) {
        if (String(existingData[de][0]) === newOrderId) { alreadyExists = true; break; }
      }
      if (alreadyExists) {
        log.push("  SKIP: '" + ckName + "' already has row for " + newOrderId);
        skippedCount++;
        continue;
      }

      // Build response array aligned to question order
      var respArray = [];
      for (var qi = 0; qi < ckQuestions.length; qi++) {
        var qText = ckQuestions[qi];
        var found = "";
        for (var rri = 0; rri < entry.responses.length; rri++) {
          if (entry.responses[rri].question === qText) {
            found = entry.responses[rri].response;
            break;
          }
        }
        respArray.push(found);
      }

      // Determine order name and customer from fresh data
      var orderName = entry.orderName;
      var customerLabel = entry.customer;
      if (ocInfo && orderLookup[ocInfo.order_id]) {
        orderName = orderLookup[ocInfo.order_id].name;
        var custId = orderLookup[ocInfo.order_id].customer_id;
        if (customerMap[custId]) customerLabel = customerMap[custId];
      }

      // Write the row
      writeResponseRow(ckName, ckQuestions, {
        orderId: newOrderId,
        orderName: orderName,
        customer: customerLabel,
        person: entry.person,
        date: entry.date,
        submittedAt: entry.submittedAt,
        responses: respArray,
      });
      migratedCount++;
    }

    log.push("Migrated " + migratedCount + " submissions to per-checklist tabs");
    if (skippedCount > 0) log.push("Skipped " + skippedCount + " (already migrated or no matching checklist)");
  }

  // ── Step 8: Populate Master Summary from migrated data ───────
  log.push("\n--- Setting up Master Summary ---");
  setupMasterSummary();

  // Read all per-checklist tabs and populate Master Summary
  var summarySheet = ss.getSheetByName(SHEETS.MASTER_SUMMARY);
  if (summarySheet) {
    // Only populate if no data rows exist yet (row 5+)
    if (summarySheet.getLastRow() <= 4) {
      var freshOcData2 = readSheetAsObjects_(ocSheet);
      var freshOrders2 = readSheetAsObjects_(ordersSheet);
      var summaryRows = [];

      for (var si = 0; si < freshOcData2.length; si++) {
        var soc = freshOcData2[si];
        if (String(soc.status) !== "completed") continue;

        var sOrderId = String(soc.order_id);
        var sOrder = null;
        for (var soi = 0; soi < freshOrders2.length; soi++) {
          if (String(freshOrders2[soi].id) === sOrderId) { sOrder = freshOrders2[soi]; break; }
        }

        var sCk = checklistMap[String(soc.checklist_id)];
        if (!sCk) continue;

        var sOrderName = sOrder ? String(sOrder.name) : "";
        var sCustLabel = sOrder ? (customerMap[String(sOrder.customer_id)] || "") : "";
        var sOtLabel = sOrder ? (orderTypeMap[String(sOrder.order_type)] || "") : "";

        summaryRows.push([sOrderId, sOrderName, sCustLabel, sOtLabel, sCk.name,
          String(soc.completed_by || ""), String(soc.work_date || ""), String(soc.completed_at || "")]);
      }

      if (summaryRows.length > 0) {
        summarySheet.getRange(5, 1, summaryRows.length, MASTER_SUMMARY_HEADERS.length).setValues(summaryRows);
        log.push("Added " + summaryRows.length + " rows to Master Summary");
      } else {
        log.push("No completed checklists to add to Master Summary");
      }
    } else {
      log.push("Master Summary already has data, skipping population");
    }
  }

  // ── Step 9: Summary ──────────────────────────────────────────
  log.push("\n=== Migration complete! ===");
  log.push("IMPORTANT: Verify the new tabs before running cleanupOldTabs()");
  log.push("  - Check per-checklist tabs have correct response data");
  log.push("  - Check Master Summary has all completed checklists");
  log.push("  - Check Orders tab has ORD-xxx IDs");
  log.push("  - Check OrderChecklists tab references ORD-xxx IDs");
  log.push("  - Old ChecklistResponses tab is still intact");

  Logger.log(log.join("\n"));
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Migration complete! Check Execution Log for details.",
    "Migration Done", 10
  );
}

// ─── Helper: Read sheet as array of objects ────────────────────

function readSheetAsObjects_(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var headers = data[0];
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) obj[String(headers[j])] = data[i][j];
    rows.push(obj);
  }
  return rows;
}

// ─── Cleanup: Remove old tabs AFTER verifying migration ────────
//
// Run this ONLY after you've verified migrateData() worked correctly.
// It renames the old ChecklistResponses tab to "ChecklistResponses_OLD"
// so you can delete it manually when you're confident.

function cleanupOldTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var oldSheet = ss.getSheetByName("ChecklistResponses");
  if (oldSheet) {
    oldSheet.setName("ChecklistResponses_OLD");
    Logger.log("Renamed 'ChecklistResponses' to 'ChecklistResponses_OLD'");
    Logger.log("You can now manually delete this tab if everything looks good.");
    ss.toast("Renamed to ChecklistResponses_OLD. Delete manually when ready.", "Cleanup Done", 10);
  } else {
    Logger.log("No 'ChecklistResponses' tab found — nothing to clean up.");
    ss.toast("No old ChecklistResponses tab found.", "Nothing to do", 5);
  }

  // Also rename old ArchivedChecklistResponses if it exists
  var oldArchived = ss.getSheetByName("ArchivedChecklistResponses");
  if (oldArchived) {
    oldArchived.setName("ArchivedChecklistResponses_OLD");
    Logger.log("Renamed 'ArchivedChecklistResponses' to 'ArchivedChecklistResponses_OLD'");
  }
}
