/**
 * Zotero → Google Sheets Reading List Importer
 * Sheet header (row 1): Paper | Authors | Year | Theme | Status | Notes
 * We also store ZoteroKey in hidden column G (no header rewrite).
 */

const ZOTERO_API_BASE = "https://api.zotero.org";
const READING_LIST_TAG = "reading list";
const READ_TAG = "Read"
const SKIMMED_TAG = "Skimmed"
const PRIORITY_TAG = "Priority"
const NOT_STARTED_TAG = "Not started"
const NOT_FINISHED_TAG = "Not finished"
const PAGE_SIZE = 100;
const SNAPSHOT_PROP_KEY = "ZOTERO_READING_LIST_SNAPSHOT_V1";

// Legacy internal note tags (kept only for migration cleanup; never write them again).
const LEGACY_TAG_IMPORTED_TO_SHEETS = "imported_to_sheets";
const LEGACY_TAG_SHEETS_ORIGIN = "origin_sheets";

// Column indices (1-based)
const COL_PAPER = 1;   // A
const COL_AUTHORS = 2; // B
const COL_YEAR = 3;    // C
const COL_THEME = 4;   // D
const COL_STATUS = 5;  // E
const COL_NOTES = 6;   // F
const COL_KEY = 7;     // G (hidden)
const COL_HASH = 8;    // H (hidden)
const COL_LINKURL = 9; // I (hidden)

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Zotero")
    .addItem("Import reading list from Zotero", "importReadingList")
    .addItem("Export changes to Zotero", "pushSheetEditsToZotero")
    .addSeparator()
    //.addItem("Import Zotero tags", "refreshThemeOptionsFromZotero")
    .addItem("Import new Zotero notes", "importNewZoteroNotes")
    .addSeparator()
    .addItem("Help", "showZoteroHelp")
    .addToUi();
}

function importReadingList() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const cfg = getConfig_();

  // Always refresh tags at start of import (silent)
  try {
    refreshThemeOptionsFromZotero(false);
  } catch (e) {
    // If tag refresh fails, you probably still want the import to run,
    // but notify the user once.
    SpreadsheetApp.getUi().alert("Warning: failed to refresh Zotero tags. Import will continue.\n\n" + e);
  }

  const t = getTableInfo_(sheet);
  ensureKeyColumnHidden_(sheet, t.colKey);
  ensureHashColumnHidden_(sheet, t.colHash);
  ensureLinkUrlColumnHidden_(sheet, t.colLinkUrl);
  ensureThemeOptionsSheet_();

  ui.alert(
    "Zotero full sync started. This will refresh the entire reading list."
  );

  // ---------- Helpers (local, so this function is self-consistent) ----------
  function getCellString_(r, c) {
    return (sheet.getRange(r, c).getValue() || "").toString();
  }

  function sheetRowSnapshot_(r) {
    const paperCell = sheet.getRange(r, t.colPaper);

    const paperText = (paperCell.getValue() || "").toString();
    const authors = getCellString_(r, t.colAuthors);
    const year = getCellString_(r, t.colYear);
    const theme = getCellString_(r, t.colTheme);
    const status = getCellString_(r, t.colStatus);
    const notes = getCellString_(r, t.colNotes);

    const linkFromPaper = getCellLinkUrl_(paperCell);
    const linkFromCol = t.colLinkUrl ? getCellString_(r, t.colLinkUrl) : "";
    const effectiveLinkUrl = normalizeLinkForHash_(linkFromPaper || linkFromCol);

    const hash = rowFingerprint_(paperText, authors, year, theme, status, notes, effectiveLinkUrl);
    return { paperText, authors, year, theme, status, notes, effectiveLinkUrl, hash };
  }

  function zoteroIncomingHash_(z, statusKeep, notesKeep) {
    const zLink = normalizeLinkForHash_(z.linkUrl);
    return rowFingerprint_(z.title, z.authors, z.year, z.themeValue, statusKeep, notesKeep, zLink);
  }

  function diffColumns_(sheetSnap, z) {
    const diffs = [];

    const zTitle = (z.title || "").toString();
    const zAuthors = (z.authors || "").toString();
    const zYear = (z.year || "").toString();
    const zTheme = (z.themeValue || "").toString();
    const zLink = normalizeLinkForHash_(z.linkUrl);

    const sTitle = (sheetSnap.paperText || "").toString();
    const sAuthors = (sheetSnap.authors || "").toString();
    const sYear = (sheetSnap.year || "").toString();
    const sTheme = (sheetSnap.theme || "").toString();
    const sLink = (sheetSnap.effectiveLinkUrl || "").toString();

    if (sTitle !== zTitle) diffs.push("Paper");
    if (sAuthors !== zAuthors) diffs.push("Authors");
    if (sYear !== zYear) diffs.push("Year");
    if (sTheme !== zTheme) diffs.push("Theme");
    if (sLink !== zLink) diffs.push("URL");

    return diffs;
  }

  function prettyDiffLine_(rowNumber, diffs) {
    return `Row ${rowNumber}: ${diffs.join(", ")}`;
  }

  function applyZoteroToRowPreservingStatusNotes_(r, z) {
    sheet.getRange(r, t.colAuthors).setValue(z.authors || "");
    sheet.getRange(r, t.colYear).setValue(z.year || "");
    sheet.getRange(r, t.colTheme).setValue(z.themeValue || "");

    const paperCell = sheet.getRange(r, t.colPaper);
    const zLink = normalizeLinkForHash_(z.linkUrl);

    if (zLink) setTitleHyperlink_(paperCell, z.title || "", zLink);
    else clearTitleHyperlink_(paperCell, z.title || "");

    if (t.colLinkUrl) sheet.getRange(r, t.colLinkUrl).setValue(zLink);

    const statusKeep = (sheet.getRange(r, t.colStatus).getValue() || "").toString();
    const notesKeep = (sheet.getRange(r, t.colNotes).getValue() || "").toString();
    const finalHash = rowFingerprint_(z.title || "", z.authors || "", z.year || "", z.themeValue || "", statusKeep, notesKeep, zLink);
    sheet.getRange(r, t.colHash).setValue(finalHash);
  }
  // -------------------------------------------------------------------------

  // -------------------- (1) Refresh hash for EVERY existing row --------------------
  const lastRow = sheet.getLastRow();
  const dataStart = t.dataStartRow;
  const existingRowCount = Math.max(0, lastRow - dataStart + 1);

  if (existingRowCount > 0) {
    const hashOut = [];
    const linkOut = [];

    for (let i = 0; i < existingRowCount; i++) {
      const r = dataStart + i;
      const snap = sheetRowSnapshot_(r);
      hashOut.push([snap.hash]);
      if (t.colLinkUrl) linkOut.push([snap.effectiveLinkUrl]);
    }

    sheet.getRange(dataStart, t.colHash, existingRowCount, 1).setValues(hashOut);
    if (t.colLinkUrl) sheet.getRange(dataStart, t.colLinkUrl, existingRowCount, 1).setValues(linkOut);
  }
  // -------------------------------------------------------------------------------

  // 1) Pull current reading list from Zotero
  const items = fetchAllItemsByTag_(cfg, READING_LIST_TAG)
    .filter(it => it?.data?.itemType !== "note");

  const zoteroRows = new Map();
  const allThemesToAdd = [];

  const STATUS_TAGS_SET = new Set([
    READ_TAG.toLowerCase(),
    SKIMMED_TAG.toLowerCase(),
    PRIORITY_TAG.toLowerCase(),
    NOT_STARTED_TAG.toLowerCase(),
    NOT_FINISHED_TAG.toLowerCase()
  ]);

  for (const item of items) {
    const key = item.key;
    if (!key) continue;

    const title = (item.data.title || "").trim();
    const authors = formatCreators_(item.data.creators || []);
    const year = parseYear_(item.data.date || "");
    const linkUrl = bestItemUrl_(item.data);

    const tags = (item.data.tags || [])
      .map(x => (x.tag || "").trim())
      .filter(x => {
        if (!x) return false;
        const lower = x.toLowerCase();
        return lower !== READING_LIST_TAG &&
          !STATUS_TAGS_SET.has(lower) &&
          !isInternalZoteroTag_(lower);
      });

    const uniqueTags = [...new Set(tags)].sort((a, b) => a.localeCompare(b));
    allThemesToAdd.push(...uniqueTags);

    const themeValue = uniqueTags.join(", ");
    zoteroRows.set(key, { title, authors, year, themeValue, linkUrl, key });
  }

  appendMissingThemeOptions_([...new Set(allThemesToAdd)]);

  // Build current sheet index: key -> rowNumber
  const keyToRow = new Map();
  if (existingRowCount > 0) {
    const keyVals = sheet.getRange(dataStart, t.colKey, existingRowCount, 1).getValues();
    for (let i = 0; i < keyVals.length; i++) {
      const k = (keyVals[i][0] || "").toString().trim();
      if (k) keyToRow.set(k, dataStart + i);
    }
  }

  // -------------------- (2) Compare incoming Zotero vs CURRENT sheet hash --------------------
  const updatedRowNumbers = [];
  const changeLines = [];

  for (const [key, z] of zoteroRows.entries()) {
    if (!keyToRow.has(key)) continue;

    const r = keyToRow.get(key);

    const statusKeep = (sheet.getRange(r, t.colStatus).getValue() || "").toString();
    const notesKeep  = (sheet.getRange(r, t.colNotes).getValue() || "").toString();

    const currentHash = (sheet.getRange(r, t.colHash).getValue() || "").toString();
    const incomingHash = zoteroIncomingHash_(z, statusKeep, notesKeep);

    if (currentHash === incomingHash) continue;

    const before = sheetRowSnapshot_(r);
    const diffs = diffColumns_(before, z);
    if (diffs.length) {
      changeLines.push(prettyDiffLine_(r, diffs));
    }

    applyZoteroToRowPreservingStatusNotes_(r, z);

    updatedRowNumbers.push(r);
  }

  if (changeLines.length) {
    const maxLines = 25;
    const shown = changeLines.slice(0, maxLines);
    const more = changeLines.length > maxLines ? `\n…and ${changeLines.length - maxLines} more` : "";
    ui.alert("Incoming Zotero changes will be applied to the sheet:\n\n" + shown.join("\n") + more);
  }
  // -----------------------------------------------------------------------------------------

  // -------------------- Append NEW rows --------------------
  const addedLabels = [];
  const addedWithImportedNotes = [];
  const addedWithNoteImportErrors = [];
  let appendAt = t.appendRow;

  for (const [key, z] of zoteroRows.entries()) {
    if (keyToRow.has(key)) continue;

    sheet.getRange(appendAt, t.colAuthors).setValue(z.authors || "");
    sheet.getRange(appendAt, t.colYear).setValue(z.year || "");
    sheet.getRange(appendAt, t.colTheme).setValue(z.themeValue || "");
    sheet.getRange(appendAt, t.colStatus).setValue("");
    sheet.getRange(appendAt, t.colNotes).setValue("");
    sheet.getRange(appendAt, t.colKey).setValue(z.key);

    const zLink = normalizeLinkForHash_(z.linkUrl);
    const paperCell = sheet.getRange(appendAt, t.colPaper);
    if (zLink) setTitleHyperlink_(paperCell, z.title || "", zLink);
    else clearTitleHyperlink_(paperCell, z.title || "");

    if (t.colLinkUrl) sheet.getRange(appendAt, t.colLinkUrl).setValue(zLink);

    if (cfg.includeNotes) {
      try {
        const notesCell = sheet.getRange(appendAt, t.colNotes);
        const stats = appendNewZoteroNotesToSheetInline_(cfg, key, notesCell, {
          includeSheetsOriginAsContent: true
        });
        if (stats.appended > 0) {
          addedWithImportedNotes.push(makeReferenceLabel_(z.title, z.authors, z.year));
        }
      } catch (e) {
        const label = makeReferenceLabel_(z.title, z.authors, z.year);
        addedWithNoteImportErrors.push(label);
        Logger.log(`WARNING: Failed importing notes on initial row import for ${label} (${key}). Error: ${e}`);
      }
    }

    const notesFinal = (sheet.getRange(appendAt, t.colNotes).getValue() || "").toString();
    const finalHash = rowFingerprint_(z.title || "", z.authors || "", z.year || "", z.themeValue || "", "", notesFinal, zLink);
    sheet.getRange(appendAt, t.colHash).setValue(finalHash);

    addedLabels.push(makeReferenceLabel_(z.title, z.authors, z.year));
    appendAt++;
  }
  // --------------------------------------------

  // -------------------- Delete rows not in Zotero reading list --------------------
  const deletedLabels = []; // NEW

  const lastRowAfterAppends = sheet.getLastRow();
  const rowCountAfter = Math.max(0, lastRowAfterAppends - dataStart + 1);

  if (rowCountAfter > 0) {
    // Pull title/authors/year/keys so we can label deletes BEFORE deleting rows
    const keyVals = sheet.getRange(dataStart, t.colKey, rowCountAfter, 1).getValues()
      .map(r => (r[0] || "").toString().trim());

    const titleVals = sheet.getRange(dataStart, t.colPaper, rowCountAfter, 1).getValues()
      .map(r => (r[0] || "").toString());

    const authorVals = sheet.getRange(dataStart, t.colAuthors, rowCountAfter, 1).getValues()
      .map(r => (r[0] || "").toString());

    const yearVals = sheet.getRange(dataStart, t.colYear, rowCountAfter, 1).getValues()
      .map(r => (r[0] || "").toString());

    const rowsToDelete = [];

    for (let i = 0; i < keyVals.length; i++) {
      const k = keyVals[i];
      if (!k) continue;
      if (!zoteroRows.has(k)) {
        rowsToDelete.push(dataStart + i);
        deletedLabels.push(makeReferenceLabel_(titleVals[i], authorVals[i], yearVals[i]));
      }
    }

    rowsToDelete.sort((a, b) => b - a);
    for (const r of rowsToDelete) {
      sheet.deleteRow(r);
    }
  }
  // -------------------------------------------------------------------------------

  // Snapshot after import so export can detect references later deleted from Sheets.
  saveReadingListSnapshot_(buildSnapshotFromZoteroRows_(zoteroRows));

  updatedRowNumbers.sort((a, b) => a - b);

  ui.alert(
    "Import complete.\n" +
    `Updated row(s): ${updatedRowNumbers.length ? updatedRowNumbers.join(", ") : "None"}\n` +
    `References added: ${formatReferenceList_(addedLabels)}\n` +
    `References removed: ${formatReferenceList_(deletedLabels)}` +
    (cfg.includeNotes && addedWithNoteImportErrors.length
      ? `\nIssues: note import failed for ${formatReferenceList_(addedWithNoteImportErrors)}`
      : "")
  );
}

/* -------------------- Read the current sheet content -------------------- */
function getSheetRowSnapshot_(sheet, rowNumber, t) {
  const paperCell = sheet.getRange(rowNumber, t.colPaper);

  // Values from cells (normalize/trim for stable hashes)
  const paper   = (getCellText_(paperCell) || "").toString().trim();
  const authors = (sheet.getRange(rowNumber, t.colAuthors).getValue() || "").toString().trim();
  const year    = (sheet.getRange(rowNumber, t.colYear).getValue() || "").toString().trim();
  const theme   = (sheet.getRange(rowNumber, t.colTheme).getValue() || "").toString().trim();
  const status  = (sheet.getRange(rowNumber, t.colStatus).getValue() || "").toString().trim();
  const notes   = (sheet.getRange(rowNumber, t.colNotes).getValue() || "").toString().trim();
  const key     = (sheet.getRange(rowNumber, t.colKey).getValue() || "").toString().trim();

  // Paper cell rich-text link info (authoritative if rich text exists)
  const linkInfo = getCellLinkInfo_(paperCell);

  // Stored LinkUrl col (fallback only when NO rich text exists)
  const linkFromCol = t.colLinkUrl
    ? (sheet.getRange(rowNumber, t.colLinkUrl).getValue() || "").toString().trim()
    : "";

  // If rich text exists, use its URL even if it's ""
  const chosenRaw = linkInfo.hasRichText ? linkInfo.url : linkFromCol;

  const effectiveLinkUrl = normalizeUrl_(chosenRaw);

  const hash = rowFingerprint_(paper, authors, year, theme, status, notes, effectiveLinkUrl);

  return { paper, authors, year, theme, status, notes, key, effectiveLinkUrl, hash };
}

/* -------------------- Import new Zotero notes -------------------- */

function importNewZoteroNotes() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const cfg = getConfig_();
  const t = getTableInfo_(sheet);

  if (!cfg.includeNotes) {
    ui.alert("Notes import is disabled (ZOTERO_INCLUDE_NOTES=false).");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < t.dataStartRow) {
    ui.alert("No rows found.");
    return;
  }

  const numRows = lastRow - t.dataStartRow + 1;

  // Read keys + notes + label fields in one go
  const keyVals = sheet.getRange(t.dataStartRow, t.colKey, numRows, 1).getValues();
  const paperVals = sheet.getRange(t.dataStartRow, t.colPaper, numRows, 1).getValues();
  const authorVals = sheet.getRange(t.dataStartRow, t.colAuthors, numRows, 1).getValues();
  const yearVals = sheet.getRange(t.dataStartRow, t.colYear, numRows, 1).getValues();
  const notesRange = sheet.getRange(t.dataStartRow, t.colNotes, numRows, 1);

  let appendedSnippets = 0;
  let processedItems = 0;

  const changedRefs = [];
  const skippedRefs = [];

  ui.alert("Appending new Zotero notes. This may take a moment.");

  for (let i = 0; i < numRows; i++) {
    const key = (keyVals[i][0] || "").toString().trim();
    if (!key) continue;

    const notesCell = notesRange.getCell(i + 1, 1);
    const stats = appendNewZoteroNotesToSheetInline_(cfg, key, notesCell);
    const refLabel = makeReferenceLabel_(
      (paperVals[i][0] || "").toString(),
      (authorVals[i][0] || "").toString(),
      (yearVals[i][0] || "").toString()
    );

    appendedSnippets += stats.appended;
    processedItems++;

    if (stats.appended > 0) changedRefs.push(refLabel);
    else skippedRefs.push(refLabel);

    if (stats.appended === 0 && stats.skippedImported > 0) {
      Logger.log(`Item ${key}: notes exist but were already imported earlier (skipped ${stats.skippedImported}).`);
    } else if (stats.appended > 0) {
      Logger.log(`Item ${key}: appended ${stats.appended} new note snippet(s).`);
    }
  }

  ui.alert(
    "Notes import complete.\n" +
    `Processed references: ${processedItems}\n` +
    `Imported note snippets: ${appendedSnippets}\n` +
    `References changed: ${formatReferenceList_(changedRefs)}\n` +
    `References unchanged: ${formatReferenceList_(skippedRefs)}`
  );
}

function getCellText_(cell) {
  // Paper cell may be rich text; getValue() is fine for the displayed text
  return (cell.getValue() || "").toString();
}

function getCellLinkUrl_(cell) {
  try {
    const rt = cell.getRichTextValue();
    if (!rt) return "";
    return (rt.getLinkUrl() || "").toString().trim();
  } catch (e) {
    return "";
  }
}

function normalizeLinkForHash_(url) {
  // Keep simple + stable. Don’t force doi.org.
  let u = (url || "").toString().trim();
  if (!u) return "";
  // Add scheme if missing (so hyperlinks work)
  if (!/^https?:\/\//i.test(u)) u = "https://" + u;
  // Trim wrapping punctuation
  u = u.replace(/^[<(\[]+/, "").replace(/[>\])\].,;:]+$/, "");
  return u;
}

/**
 * Compute "current sheet hash" for a row using actual sheet state.
 * Uses hyperlink from Paper cell if present, else LinkUrl column.
 */
function computeSheetRowHash_(sheet, rowNumber, t) {
  const paperCell = sheet.getRange(rowNumber, t.colPaper);

  const paper = getCellText_(paperCell).trim();
  const authors = (sheet.getRange(rowNumber, t.colAuthors).getValue() || "").toString().trim();
  const year = (sheet.getRange(rowNumber, t.colYear).getValue() || "").toString().trim();
  const theme = (sheet.getRange(rowNumber, t.colTheme).getValue() || "").toString().trim();
  const status = (sheet.getRange(rowNumber, t.colStatus).getValue() || "").toString().trim();
  const notes = (sheet.getRange(rowNumber, t.colNotes).getValue() || "").toString().trim();

  // Link logic: Paper cell rich-text overrides stored link column (even if empty).
  const linkInfo = getCellLinkInfo_(paperCell);

  const stored = (t.colLinkUrl
    ? (sheet.getRange(rowNumber, t.colLinkUrl).getValue() || "").toString().trim()
    : ""
  );

  const chosenRaw = linkInfo.hasRichText ? linkInfo.url : stored;

  // Normalize consistently (adds https:// if missing, converts bare DOI to https://doi.org/...)
  const effectiveLinkUrl = normalizeUrl_(chosenRaw);

  const hash = rowFingerprint_(paper, authors, year, theme, status, notes, effectiveLinkUrl);
  return { hash, effectiveLinkUrl };
}

function getCellLinkInfo_(cell) {
  try {
    const rt = cell.getRichTextValue();
    if (!rt) return { hasRichText: false, hasLink: false, url: "" };

    const url = (rt.getLinkUrl() || "").toString().trim();
    const hasLink = !!url;

    return { hasRichText: true, hasLink, url };
  } catch (e) {
    return { hasRichText: false, hasLink: false, url: "" };
  }
}

/* -------------------- Zotero key in sheet -------------------- */

function ensureKeyColumnHidden_(sheet, colKey) {
  try { sheet.hideColumns(colKey); } catch (e) {}
}

function ensureLinkUrlColumnHidden_(sheet, col) {
  try { sheet.hideColumns(col); } catch (e) {}
}

function getExistingKeysFromSheet_(sheet, dataStartRow, colKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return new Set();

  const values = sheet.getRange(dataStartRow, colKey, lastRow - dataStartRow + 1, 1).getValues();
  const keys = values.map(r => (r[0] || "").toString().trim()).filter(Boolean);
  return new Set(keys);
}

/* -------------------- DOI hyperlink -------------------- */

function setTitleHyperlink_(cell, title, url) {
  const rich = SpreadsheetApp.newRichTextValue()
    .setText(title)
    .setLinkUrl(url)
    .build();
  cell.setRichTextValue(rich);
}

function extractDoi_(data) {
  // 1) Prefer explicit fields
  const direct = (data.DOI || data.doi || "").toString().trim();
  const cleanedDirect = normalizeDoi_(direct);
  if (cleanedDirect) return cleanedDirect;

  // 2) Fallback: find DOI anywhere in "extra"
  const extra = (data.extra || "").toString();

  // Match a DOI-ish token, not just "non-space" (avoids trailing punctuation)
  // DOI prefix is always 10.<digits>/
  const m = extra.match(/\b10\.\d{4,9}\/[^\s"<>()]+/i);
  if (m && m[0]) return normalizeDoi_(m[0]);

  return "";
}

function normalizeDoi_(doi) {
  let d = (doi || "").toString().trim();
  if (!d) return "";

  // Remove URL prefix
  d = d.replace(/^https?:\/\/(dx\.)?doi\.org\//i, "");

  // Remove common trailing punctuation and wrapping
  // (very common in Zotero "extra" fields)
  d = d.replace(/^[<(\[]+/, "");
  d = d.replace(/[>\])\].,;:]+$/, "");

  // Also strip trailing period(s) that sometimes stick after cleanup
  d = d.replace(/\.+$/, "");

  return d.trim();
}

function bestItemUrl_(itemData) {
  // 1) Prefer Zotero URL field (best for theses/PDFs)
  const url = (itemData.url || "").toString().trim();
  if (url) return normalizeUrl_(url);

  // 2) Else prefer DOI (explicit fields or extra)
  const doi = extractDoi_(itemData);
  if (doi) return normalizeUrl_(doi); // normalizeUrl_ will convert 10.xxxx/... to https://doi.org/...

  // 3) Else try "extra" for URL:
  const extra = (itemData.extra || "").toString();
  const mUrl = extra.match(/URL:\s*(\S+)/i);
  if (mUrl && mUrl[1]) return normalizeUrl_(mUrl[1].trim());

  return "";
}

function normalizeUrl_(url) {
  let u = (url || "").toString().trim();
  if (!u) return "";

  // Trim surrounding punctuation
  u = u.replace(/^[<(\[]+/, "");
  u = u.replace(/[>\])\].,;:]+$/, "");
  u = u.replace(/\s/g, "");

  // If it looks like a DOI string (10.xxxx/...), convert to doi.org
  if (/^10\.\d{4,9}\//i.test(u)) {
    u = "https://doi.org/" + u;
  }

  // If missing scheme, add https://
  if (!/^https?:\/\//i.test(u)) {
    u = "https://" + u;
  }

  // Normalize doi.org URLs only (optional)
  u = u.replace(/^(https?:\/\/)(dx\.)?doi\.org\/+/i, "https://doi.org/");
  u = u.replace(/\/+$/, ""); // remove trailing slash

  return u;
}

// Hash needs to include the link URL (and later status too)
function rowFingerprint_(paper, authors, year, theme, status, notes, linkUrl) {
  const s = [
    (paper || "").toString().trim(),
    (authors || "").toString().trim(),
    (year || "").toString().trim(),
    (theme || "").toString().trim(),
    (status || "").toString().trim(),
    (notes || "").toString().trim(),
    (linkUrl || "").toString().trim()
  ].join("||");
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return raw.map(b => (b & 0xff).toString(16).padStart(2, "0")).join("");
}

function getCellLinkUrl_(cell) {
  try {
    const rt = cell.getRichTextValue();
    if (!rt) return "";
    return (rt.getLinkUrl() || "").toString().trim();
  } catch (e) {
    return "";
  }
}

function clearTitleHyperlink_(cell, title) {
  const rich = SpreadsheetApp.newRichTextValue()
    .setText(title)
    .build(); // no link
  cell.setRichTextValue(rich);
}

/* -------------------- Notes: HTML -> plain text, append -------------------- */

// -------------------- Notes sync markers + inline helpers --------------------

const NOTE_MARK_SHEETS_ORIGIN = "<!--ZSHEET:ORIGIN=SHEETS-->";
const NOTE_MARK_IMPORTED = "<!--ZSHEET:IMPORTED_TO_SHEETS-->";
const SHEET_NOTE_HEADER = "Imported from Google Sheets";

function hasMarker_(html, marker) {
  return (html || "").toString().includes(marker);
}

function addMarkerIfMissing_(html, marker) {
  const s = (html || "").toString();
  return hasMarker_(s, marker) ? s : (s + "\n" + marker);
}

function isInternalZoteroTag_(tag) {
  const t = (tag || "").toString().trim().toLowerCase();
  return t === LEGACY_TAG_IMPORTED_TO_SHEETS || t === LEGACY_TAG_SHEETS_ORIGIN;
}

function stripInternalNoteTags_(tagsField) {
  const before = (tagsField || []).map(x => (x?.tag || "").toString().trim()).filter(Boolean);
  const after = before.filter(tag => !isInternalZoteroTag_(tag));
  return {
    changed: before.length !== after.length,
    tags: after.map(tag => ({ tag }))
  };
}

// Inline append to keep the cell short: "existing; new; new2"
function appendInlineText_(existing, addition) {
  const a = (existing || "").toString().trim();
  const b = (addition || "").toString().trim();
  if (!b) return a;
  if (!a) return b;
  return `${a}; ${b}`;
}

function escapeRegex_(s) {
  return (s || "").toString().replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function hasSheetsOriginHeaderPrefix_(html) {
  const plain = htmlToPlainText_(html).replace(/\s+/g, " ").trim();
  const re = new RegExp(`^${escapeRegex_(SHEET_NOTE_HEADER)}\\b`, "i");
  return re.test(plain);
}

function plainFromSheetsOriginHtml_(html) {
  const raw = (html || "").toString();
  const header = escapeRegex_(SHEET_NOTE_HEADER);

  // Remove known marker and header paragraph in HTML form first.
  const withoutHeaderHtml = raw
    .replace(/<!--\s*ZSHEET:ORIGIN=SHEETS\s*-->/gi, "")
    .replace(
      new RegExp(`<p[^>]*>\\s*(?:<strong[^>]*>)?\\s*${header}\\s*(?:<\\/strong>)?\\s*<\\/p>`, "i"),
      ""
    );

  // Then remove any remaining leading plain-text header variant.
  let plain = htmlToPlainText_(withoutHeaderHtml).replace(/\s+/g, " ").trim();
  plain = plain.replace(new RegExp(`^${header}\\s*[:\\-–—]?\\s*`, "i"), "").trim();
  return plain;
}

function appendNewZoteroNotesToSheetInline_(cfg, parentKey, notesCell, options) {
  const includeSheetsOriginAsContent = !!(options && options.includeSheetsOriginAsContent);
  const url = buildItemChildrenUrl_(cfg, parentKey, { itemType: "note", limit: 100 });
  const children = zoteroFetch_(cfg, url);
  const notes = Array.isArray(children) ? children : [];

  const toAppendPlain = [];
  const notesToUpdate = [];

  let skippedImported = 0;
  let skippedOrigin = 0;
  let emptyNotes = 0;

  for (const n of notes) {
    const html = (n?.data?.note || "").toString();
    const tags = n?.data?.tags || [];

    const hasOriginMarker = hasMarker_(html, NOTE_MARK_SHEETS_ORIGIN);
    const hasImportedMarker = hasMarker_(html, NOTE_MARK_IMPORTED);
    const hasLegacyOriginTag = hasTag_(tags, LEGACY_TAG_SHEETS_ORIGIN);
    const hasLegacyImportedTag = hasTag_(tags, LEGACY_TAG_IMPORTED_TO_SHEETS);
    const cleanedTags = stripInternalNoteTags_(tags);

    let migratedHtml = html;
    if (hasLegacyOriginTag && !hasOriginMarker) {
      migratedHtml = addMarkerIfMissing_(migratedHtml, NOTE_MARK_SHEETS_ORIGIN);
    }
    if (hasLegacyImportedTag && !hasImportedMarker) {
      migratedHtml = addMarkerIfMissing_(migratedHtml, NOTE_MARK_IMPORTED);
    }

    const hasOriginHeaderPrefix = hasSheetsOriginHeaderPrefix_(migratedHtml);
    if (hasOriginHeaderPrefix && !hasOriginMarker) {
      migratedHtml = addMarkerIfMissing_(migratedHtml, NOTE_MARK_SHEETS_ORIGIN);
    }

    const isOrigin = hasOriginMarker || hasLegacyOriginTag || hasOriginHeaderPrefix;
    const isImported = hasImportedMarker || hasLegacyImportedTag;

    if (isOrigin) {
      if (includeSheetsOriginAsContent) {
        const plain = plainFromSheetsOriginHtml_(migratedHtml);
        if (plain) toAppendPlain.push(plain);
        else emptyNotes++;
      } else {
        skippedOrigin++;
      }
      if (migratedHtml !== html || cleanedTags.changed) {
        notesToUpdate.push({
          key: n.key,
          version: n.version,
          data: n.data,
          note: migratedHtml,
          tags: cleanedTags.tags
        });
      }
      continue;
    }

    if (isImported) {
      skippedImported++;
      if (migratedHtml !== html || cleanedTags.changed) {
        notesToUpdate.push({
          key: n.key,
          version: n.version,
          data: n.data,
          note: migratedHtml,
          tags: cleanedTags.tags
        });
      }
      continue;
    }

    const plain = htmlToPlainText_(html).replace(/\s+/g, " ").trim();
    if (!plain) { emptyNotes++; continue; }

    toAppendPlain.push(plain);
    notesToUpdate.push({
      key: n.key,
      version: n.version,
      data: n.data,
      note: addMarkerIfMissing_(migratedHtml, NOTE_MARK_IMPORTED),
      tags: cleanedTags.tags
    });
  }

  if (toAppendPlain.length) {
    const existing = (notesCell.getValue() || "").toString();
    const chunk = toAppendPlain.join("; ");
    notesCell.setValue(appendInlineText_(existing, chunk));
  }

  if (notesToUpdate.length) {
    for (const m of notesToUpdate) {
      try {
        const updated = { ...m.data };
        updated.note = m.note;
        updated.tags = m.tags;
        updated.version = m.version;
        zoteroPutItemData_(cfg, m.key, updated);
      } catch (e) {
        Logger.log(
          `WARNING: Failed to update note markers for parent ${parentKey}, note ${m.key}. ` +
          `Error: ${e}`
        );
      }
    }
  }

  return {
    appended: toAppendPlain.length,
    skippedImported,
    skippedOrigin,
    emptyNotes,
    totalChildNotes: notes.length
  };
}

function hasTag_(tagsField, tag) {
  const t = (tag || "").toString().trim().toLowerCase();
  return (tagsField || []).some(x => (x?.tag || "").toString().trim().toLowerCase() === t);
}

function htmlToPlainText_(html) {
  let s = html;
  s = s.replace(/<\/(p|div|br|li|h\d)>/gi, "\n");
  s = s.replace(/<li>/gi, "• ");
  s = s.replace(/<[^>]+>/g, "");
  s = s.replace(/&nbsp;/g, " ");
  s = s.replace(/&amp;/g, "&");
  s = s.replace(/&lt;/g, "<");
  s = s.replace(/&gt;/g, ">");
  s = s.replace(/&quot;/g, "\"");
  s = s.replace(/&#39;/g, "'");
  s = s.replace(/\r/g, "");
  s = s.replace(/\n{3,}/g, "\n\n");
  return s.trim();
}

/* -------------------- Zotero API -------------------- */

function fetchAllItemsByTag_(cfg, tag) {
  const out = [];
  let start = 0;

  while (true) {
    const url = buildItemsUrl_(cfg, { tag, start, limit: PAGE_SIZE });
    const resp = zoteroFetch_(cfg, url);
    if (!Array.isArray(resp) || resp.length === 0) break;

    out.push(...resp);
    if (resp.length < PAGE_SIZE) break;
    start += PAGE_SIZE;
  }
  return out;
}

function zoteroFetch_(cfg, url) {
  const res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      "Zotero-API-Key": cfg.apiKey,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code < 200 || code >= 300) throw new Error(`Zotero API error ${code}: ${text}`);
  return JSON.parse(text);
}

function buildItemsUrl_(cfg, { tag, start, limit }) {
  const base = `${ZOTERO_API_BASE}/users/${encodeURIComponent(cfg.libraryId)}/items`;
  const params = { tag, start: String(start), limit: String(limit) };
  return base + "?" + toQuery_(params);
}

function buildItemChildrenUrl_(cfg, itemKey, { itemType, limit }) {
  const base = `${ZOTERO_API_BASE}/users/${encodeURIComponent(cfg.libraryId)}/items/${encodeURIComponent(itemKey)}/children`;
  const params = { itemType, limit: String(limit) };
  return base + "?" + toQuery_(params);
}

/* -------------------- Sheet header behavior -------------------- */

function ensureHeaderOnlyIfMissing_(sheet) {
  // DO NOT rewrite header if it exists. Only write if row 1 is empty.
  const expected = ["Paper", "Authors", "Year", "Theme", "Status", "Notes"];
  const current = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
  const allEmpty = current.every(v => !v);

  if (allEmpty) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    sheet.setFrozenRows(1);
  }

  // Do not touch existing header even if it differs.
  // Do not write ZoteroKey header — keys live in hidden col G.
}

/* -------------------- Formatting helpers -------------------- */

function formatCreators_(creators) {
  return (creators || [])
    .filter(c => c && (c.lastName || c.name))
    .map(c => c.name ? c.name : ((c.firstName || "").trim() ? `${(c.lastName || "").trim()}, ${(c.firstName || "").trim()}` : (c.lastName || "").trim()))
    .filter(Boolean)
    .join("; ");
}

function parseYear_(dateStr) {
  const m = String(dateStr || "").match(/\b(18|19|20|21)\d{2}\b/);
  return m ? m[0] : "";
}

function sanitizeLabelToken_(value, fallback) {
  const token = (value || "").toString().trim().replace(/[^\p{L}\p{N}_-]+/gu, "");
  return token || fallback;
}

function firstMeaningfulWordOver3_(title) {
  const words = (title || "").toString().trim()
    .split(/\s+/)
    .map(w => sanitizeLabelToken_(w, ""))
    .filter(Boolean);

  if (!words.length) return "untitled";
  return words.find(w => w.length > 3) || words[0];
}

function firstAuthorLastNameForLabel_(authorsStr) {
  const a = (authorsStr || "").toString().trim();
  if (!a) return "noauthor";

  const first = a.split(";")[0].trim();
  if (!first) return "noauthor";

  if (first.includes(",")) {
    return sanitizeLabelToken_(first.split(",")[0], "noauthor");
  }

  const parts = first.split(/\s+/).filter(Boolean);
  return sanitizeLabelToken_(parts.length ? parts[parts.length - 1] : "", "noauthor");
}

function yearForLabel_(value) {
  const s = (value || "").toString().trim();
  const m = s.match(/\b(18|19|20|21)\d{2}\b/);
  return m ? m[0] : "noyear";
}

function makeReferenceLabel_(title, authorsStr, yearStr) {
  return `${firstMeaningfulWordOver3_(title)}_${firstAuthorLastNameForLabel_(authorsStr)}_${yearForLabel_(yearStr)}`;
}

function formatReferenceList_(labels, maxItems) {
  const max = maxItems || 25;
  const deduped = [];
  const seen = new Set();

  for (const raw of (labels || [])) {
    const label = (raw || "").toString().trim();
    if (!label || seen.has(label)) continue;
    seen.add(label);
    deduped.push(label);
  }

  if (!deduped.length) return "None";
  const shown = deduped.slice(0, max);
  const more = deduped.length > max ? `, ...(+${deduped.length - max} more)` : "";
  return shown.join(", ") + more;
}

function buildSnapshotFromZoteroRows_(zoteroRows) {
  const byKey = {};
  for (const [key, z] of zoteroRows.entries()) {
    if (!key) continue;
    const title = (z?.title || "").toString();
    const authors = (z?.authors || "").toString();
    const year = (z?.year || "").toString();
    byKey[key] = {
      label: makeReferenceLabel_(title, authors, year),
      title,
      authors,
      year
    };
  }
  return byKey;
}

function buildSnapshotFromSheet_(sheet, t) {
  const byKey = {};
  const lastRow = sheet.getLastRow();
  if (lastRow < t.dataStartRow) return byKey;

  const numRows = lastRow - t.dataStartRow + 1;
  const keys = sheet.getRange(t.dataStartRow, t.colKey, numRows, 1).getValues();
  const papers = sheet.getRange(t.dataStartRow, t.colPaper, numRows, 1).getValues();
  const authors = sheet.getRange(t.dataStartRow, t.colAuthors, numRows, 1).getValues();
  const years = sheet.getRange(t.dataStartRow, t.colYear, numRows, 1).getValues();

  for (let i = 0; i < numRows; i++) {
    const key = (keys[i][0] || "").toString().trim();
    if (!key) continue;

    const title = (papers[i][0] || "").toString();
    const author = (authors[i][0] || "").toString();
    const year = (years[i][0] || "").toString();

    byKey[key] = {
      label: makeReferenceLabel_(title, author, year),
      title,
      authors: author,
      year
    };
  }

  return byKey;
}

function saveReadingListSnapshot_(byKey) {
  const props = getSnapshotPropertiesStore_();
  const payload = {
    savedAt: new Date().toISOString(),
    byKey: byKey || {}
  };
  props.setProperty(SNAPSHOT_PROP_KEY, JSON.stringify(payload));
}

function loadReadingListSnapshot_() {
  const props = getSnapshotPropertiesStore_();
  const raw = props.getProperty(SNAPSHOT_PROP_KEY);
  if (!raw) return { exists: false, savedAt: "", byKey: {} };

  try {
    const parsed = JSON.parse(raw);
    const byKey = (parsed && typeof parsed.byKey === "object" && parsed.byKey) ? parsed.byKey : {};
    return {
      exists: true,
      savedAt: (parsed?.savedAt || "").toString(),
      byKey
    };
  } catch (e) {
    Logger.log(`WARNING: invalid ${SNAPSHOT_PROP_KEY}; resetting snapshot.`);
    return { exists: false, savedAt: "", byKey: {} };
  }
}

function getSnapshotPropertiesStore_() {
  return PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties();
}

function pickTheme_(tags) {
  if (!tags || !tags.length) return "";
  const sorted = [...new Set(tags)].sort((a, b) => a.localeCompare(b));
  return sorted[0];
}

function toQuery_(params) {
  return Object.keys(params)
    .filter(k => params[k] !== undefined && params[k] !== null && params[k] !== "")
    .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
    .join("&");
}

/* -------------------- Config -------------------- */

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  const libraryId = props.getProperty("ZOTERO_LIBRARY_ID");
  const apiKey = props.getProperty("ZOTERO_API_KEY");
  const includeNotes = (props.getProperty("ZOTERO_INCLUDE_NOTES") || "true").toLowerCase() === "true";

  if (!libraryId || !apiKey) {
    throw new Error(
      "Missing Zotero config. Set Script Properties:\n" +
      "- ZOTERO_LIBRARY_ID\n" +
      "- ZOTERO_API_KEY\n" +
      "- ZOTERO_INCLUDE_NOTES (true|false)"
    );
  }
  return { libraryId, apiKey, includeNotes };
}

function findHeaderRow_(sheet) {
  const requiredMain = ["paper", "authors", "year", "theme", "status", "notes"]; // required
  const optional = ["key", "hash", "linkurl"]; // optional
  const maxScan = Math.min(50, sheet.getMaxRows());
  const width = Math.min(50, sheet.getMaxColumns());

  const values = sheet.getRange(1, 1, maxScan, width).getValues();

  for (let r = 0; r < values.length; r++) {
    const row = values[r].map(v => (v || "").toString().trim().toLowerCase());

    const colMap = {};
    for (let c = 0; c < row.length; c++) {
      const name = row[c];
      if (requiredMain.includes(name) || optional.includes(name)) colMap[name] = c + 1;
    }

    const hasMain = requiredMain.every(h => !!colMap[h]);
    if (!hasMain) continue;

    // If optional columns missing, fall back to "notes+1, +2, +3"
    const colPaper   = colMap.paper;
    const colAuthors = colMap.authors;
    const colYear    = colMap.year;
    const colTheme   = colMap.theme;
    const colStatus  = colMap.status;
    const colNotes   = colMap.notes;

    const colKey     = colMap.key     || (colNotes + 1);
    const colHash    = colMap.hash    || (colKey + 1);
    const colLinkUrl = colMap.linkurl || (colHash + 1);

    return {
      headerRow: r + 1,
      colPaper, colAuthors, colYear, colTheme, colStatus, colNotes,
      colKey, colHash, colLinkUrl
    };
  }

  throw new Error(
    `Could not find a header row containing at least: Paper, Authors, Year, Theme, Status, Notes (case-insensitive).`
  );
}

function getTableInfo_(sheet) {
  const h = findHeaderRow_(sheet);
  const headerRow = h.headerRow;
  const dataStartRow = headerRow + 1;

  // Determine appendRow by last non-empty Paper cell
  const lastRow = sheet.getLastRow();
  let appendRow = dataStartRow;

  if (lastRow >= dataStartRow) {
    const paperVals = sheet.getRange(dataStartRow, h.colPaper, lastRow - dataStartRow + 1, 1)
      .getValues()
      .map(r => (r[0] || "").toString().trim());

    let lastNonEmptyOffset = -1;
    for (let i = paperVals.length - 1; i >= 0; i--) {
      if (paperVals[i]) { lastNonEmptyOffset = i; break; }
    }
    appendRow = lastNonEmptyOffset === -1 ? dataStartRow : (dataStartRow + lastNonEmptyOffset + 1);
  }

  return { ...h, dataStartRow, appendRow };
}

const THEME_OPTIONS_SHEET = "ThemeOptions";
const THEME_OPTIONS_COL = 1; // column A

function ensureThemeOptionsSheet_() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(THEME_OPTIONS_SHEET);
  if (!s) {
    s = ss.insertSheet(THEME_OPTIONS_SHEET);
    s.getRange(1, 1).setValue("Theme options");
  }
  return s;
}

function getThemeOptionsSet_() {
  const s = ensureThemeOptionsSheet_();
  const last = s.getLastRow();
  if (last < 2) return new Set();
  const vals = s.getRange(2, THEME_OPTIONS_COL, last - 1, 1).getValues()
    .map(r => (r[0] || "").toString().trim())
    .filter(Boolean);
  return new Set(vals);
}

function appendMissingThemeOptions_(tags) {
  const s = ensureThemeOptionsSheet_();
  const existing = getThemeOptionsSet_();

  const missing = [...new Set(tags)]
    .map(t => t.trim())
    .filter(t => t && !existing.has(t));

  if (!missing.length) return 0;

  const startRow = s.getLastRow() + 1;
  s.getRange(startRow, THEME_OPTIONS_COL, missing.length, 1)
    .setValues(missing.map(x => [x]));
  return missing.length;
}

function refreshThemeOptionsFromZotero(showAlerts) {
  const ui = SpreadsheetApp.getUi();
  const cfg = getConfig_();
  const doAlerts = (showAlerts !== false);

  if (doAlerts) ui.alert("Refreshing ThemeOptions from Zotero tags…");

  // Fetch *all* tags in your Zotero library
  const tags = fetchAllTags_(cfg)
    .map(t => t.trim())
    .filter(Boolean);

  const STATUS_TAGS_SET = new Set([
    READ_TAG.toLowerCase(),
    SKIMMED_TAG.toLowerCase(),
    PRIORITY_TAG.toLowerCase(),
    NOT_STARTED_TAG.toLowerCase(),
    NOT_FINISHED_TAG.toLowerCase()
  ]);

  const cleaned = tags.filter(t => {
    const tl = t.toLowerCase();
    return tl !== READING_LIST_TAG &&
      !STATUS_TAGS_SET.has(tl) &&
      !isInternalZoteroTag_(tl);
  });

  cleaned.sort((a, b) => a.localeCompare(b));

  // Write to ThemeOptions!A2:A
  const s = ensureThemeOptionsSheet_();

  // Clear existing A2:A
  const last = s.getLastRow();
  if (last >= 2) {
    s.getRange(2, THEME_OPTIONS_COL, last - 1, 1).clearContent();
  }

  if (cleaned.length) {
    s.getRange(2, THEME_OPTIONS_COL, cleaned.length, 1)
      .setValues(cleaned.map(x => [x]));
  }

  // ---- NEW: store sync metadata in hidden column B ----
  // B1: last sync timestamp, B2: count
  s.getRange(1, 2).setValue("Last tag sync"); // B1 label (optional)
  s.getRange(2, 2).setValue(new Date());      // B2 timestamp
  s.getRange(3, 2).setValue(cleaned.length);  // B3 count

  // Hide column B (safe if already hidden)
  try { s.hideColumns(2); } catch (e) {}

  if (doAlerts) ui.alert(`Done. Wrote ${cleaned.length} tag(s) to ${THEME_OPTIONS_SHEET}!A2:A`);
}

function fetchAllTags_(cfg) {
  const out = [];
  let start = 0;

  while (true) {
    const url = buildTagsUrl_(cfg, { start, limit: PAGE_SIZE });
    const resp = zoteroFetch_(cfg, url);

    if (!Array.isArray(resp) || resp.length === 0) break;

    // Zotero returns objects like { tag: "X", meta: { numItems: ... } }
    for (const obj of resp) {
      const tag = (obj && obj.tag) ? obj.tag.toString() : "";
      if (tag) out.push(tag);
    }

    if (resp.length < PAGE_SIZE) break;
    start += PAGE_SIZE;
  }

  return out;
}

function buildTagsUrl_(cfg, { start, limit }) {
  // Personal library
  const base = `${ZOTERO_API_BASE}/users/${encodeURIComponent(cfg.libraryId)}/tags`;
  const params = { start: String(start), limit: String(limit) };
  return base + "?" + toQuery_(params);
}

/* -------------------- Write to Zotero -------------------- */


function pushSheetEditsToZotero() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const cfg = getConfig_();
  const t = getTableInfo_(sheet);

  const changedRows = [];
  const conflictRefs = [];
  const failedRefs = [];

  const lastRow = sheet.getLastRow();
  const numRows = Math.max(0, lastRow - t.dataStartRow + 1);

  const keys = numRows > 0
    ? sheet.getRange(t.dataStartRow, t.colKey, numRows, 1).getValues().map(r => (r[0] || "").toString().trim())
    : [];
  const hashes = numRows > 0
    ? sheet.getRange(t.dataStartRow, t.colHash, numRows, 1).getValues().map(r => (r[0] || "").toString())
    : [];

  const baselineSnapshot = loadReadingListSnapshot_();
  const currentSnapshot = buildSnapshotFromSheet_(sheet, t);
  let deletionNote = "";

  // -------- WARNING ONLY FOR Paper/Authors/Year --------
  const refsWithCoreChanges = [];
  for (let i = 0; i < numRows; i++) {
    const rowNumber = t.dataStartRow + i;
    const itemKey = keys[i];
    if (!itemKey) continue;

    const snap = getSheetRowSnapshot_(sheet, rowNumber, t);

    const hashCell = sheet.getRange(rowNumber, t.colHash);
    const lastCore = getCoreHashCheckpoint_(hashCell);
    const currentCore = coreFingerprint_(snap.paper, snap.authors, snap.year);

    // If we have never checkpointed this row before, checkpoint now (no warning)
    if (!lastCore) {
      setCoreHashCheckpoint_(hashCell, currentCore);
      continue;
    }

    if (lastCore !== currentCore) {
      refsWithCoreChanges.push(makeReferenceLabel_(snap.paper, snap.authors, snap.year));
    }
  }

  if (refsWithCoreChanges.length) {
    const resp = ui.alert(
      "Export to Zotero",
      `Changes detected in Paper/Authors/Year for reference(s): ${formatReferenceList_(refsWithCoreChanges, 12)}.\n` +
      "Continue exporting to Zotero?",
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return;
  }
  // -------- END WARNING --------

  for (let i = 0; i < numRows; i++) {
    const rowNumber = t.dataStartRow + i;
    const itemKey = keys[i];
    if (!itemKey) continue;

    const snap = getSheetRowSnapshot_(sheet, rowNumber, t);
    const lastHash = hashes[i];

    // Full-row hash gating (so Notes/Status edits still push notes/tag changes)
    if (lastHash && lastHash === snap.hash) continue;

    try {
      const itemObj = zoteroGetItem_(cfg, itemKey);
      const data = itemObj.data;

      // --- Update title ---
      const newTitle = (snap.paper || "").toString().trim();
      if (newTitle) data.title = newTitle;

      // --- Update URL FROM SHEET HYPERLINK (including deletion) ---
      // snap.effectiveLinkUrl is already normalized by your normalizeUrl_
      const newUrl = (snap.effectiveLinkUrl || "").toString().trim();
      data.url = newUrl; // if "", Zotero URL becomes empty (deletion)

      // Optional: if URL is doi.org/<doi>, also set DOI field
      const doi = doiFromUrl_(newUrl);
      if (doi) data.DOI = doi;
      // If user deleted the link and you want DOI cleared too, uncomment:
      // if (!newUrl) data.DOI = "";

      // --- Theme/Status -> tags ---
      const themeTags = (snap.theme || "").toString()
        .split(",")
        .map(s => s.trim())
        .filter(Boolean);

      const statusTag = (snap.status || "").toString().trim();

      const tagsOut = new Set(themeTags.filter(tag => !isInternalZoteroTag_(tag)));
      if (statusTag && !isInternalZoteroTag_(statusTag)) tagsOut.add(statusTag);
      tagsOut.add(READING_LIST_TAG);

      data.tags = Array.from(tagsOut).map(tag => ({ tag }));

      data.version = itemObj.version;
      zoteroPutItemData_(cfg, itemKey, data);

      // Notes push (Sheets-origin note)
      upsertSheetNotesChild_(cfg, itemKey, (snap.notes || "").toString());

      // ✅ After success: store full hash + linkUrl + core checkpoint
      sheet.getRange(rowNumber, t.colHash).setValue(snap.hash);
      sheet.getRange(rowNumber, t.colLinkUrl).setValue(newUrl);

      const hashCell = sheet.getRange(rowNumber, t.colHash);
      setCoreHashCheckpoint_(hashCell, coreFingerprint_(snap.paper, snap.authors, snap.year));

      changedRows.push(rowNumber);

    } catch (e) {
      const msg = (e && e.message) ? e.message : String(e);
      const rowLabel = makeReferenceLabel_(snap.paper, snap.authors, snap.year);

      if (msg.includes("Zotero API error 412")) {
        conflictRefs.push(rowLabel);
        sheet.getRange(rowNumber, t.colStatus).setValue("CONFLICT (Zotero updated elsewhere)");
      } else {
        failedRefs.push(rowLabel);
        sheet.getRange(rowNumber, t.colStatus).setValue("PUSH FAILED");
      }
    }
  }

  const removedFromReadingList = [];
  const removeReadingListFailed = [];
  const failedDeletedKeys = [];

  if (!baselineSnapshot.exists) {
    deletionNote = "Deletion check skipped (no baseline snapshot found). A baseline was created now.";
  } else {
    const deletedKeys = Object.keys(baselineSnapshot.byKey).filter(key => !currentSnapshot[key]);

    for (const key of deletedKeys) {
      const ref = baselineSnapshot.byKey[key] || {};
      const label = (ref.label || key).toString();

      try {
        const result = removeReadingListTagFromItem_(cfg, key);
        if (result.removed) removedFromReadingList.push(label);
      } catch (e) {
        removeReadingListFailed.push(label);
        failedDeletedKeys.push(key);
        Logger.log(`WARNING: Failed removing "${READING_LIST_TAG}" for deleted sheet item ${key}. Error: ${e}`);
      }
    }
  }

  const snapshotToSave = { ...currentSnapshot };
  if (baselineSnapshot.exists) {
    for (const key of failedDeletedKeys) {
      snapshotToSave[key] = baselineSnapshot.byKey[key] || { label: key, title: "", authors: "", year: "" };
    }
  }
  saveReadingListSnapshot_(snapshotToSave);

  changedRows.sort((a, b) => a - b);
  const issueParts = [];
  if (conflictRefs.length) issueParts.push(`conflicts: ${formatReferenceList_(conflictRefs)}`);
  if (failedRefs.length) issueParts.push(`push failures: ${formatReferenceList_(failedRefs)}`);
  if (removeReadingListFailed.length) issueParts.push(`reading-list removal failures: ${formatReferenceList_(removeReadingListFailed)}`);
  if (deletionNote) issueParts.push(deletionNote);

  ui.alert(
    "Export to Zotero complete.\n" +
    `Updated row(s): ${changedRows.length ? changedRows.join(", ") : "None"}\n` +
    `Removed from reading list: ${formatReferenceList_(removedFromReadingList)}\n` +
    `Issues: ${issueParts.length ? issueParts.join(" | ") : "None"}`
  );
}

function removeReadingListTagFromItem_(cfg, itemKey) {
  const itemObj = zoteroGetItem_(cfg, itemKey);
  const data = itemObj.data || {};
  const tags = Array.isArray(data.tags) ? data.tags : [];

  const filtered = tags.filter(t => {
    const tag = (t?.tag || "").toString().trim().toLowerCase();
    return tag !== READING_LIST_TAG;
  });

  if (filtered.length === tags.length) return { removed: false };

  data.tags = filtered
    .map(t => ({ tag: (t?.tag || "").toString().trim() }))
    .filter(t => !!t.tag);
  data.version = itemObj.version;

  zoteroPutItemData_(cfg, itemKey, data);
  return { removed: true };
}

function coreFingerprint_(paper, authors, year) {
  const s = [
    (paper || "").toString().trim(),
    (authors || "").toString().trim(),
    (year || "").toString().trim()
  ].join("||");
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return raw.map(b => (b & 0xff).toString(16).padStart(2, "0")).join("");
}

// Store/retrieve core-hash checkpoint in the NOTE of the Hash cell
function getCoreHashCheckpoint_(hashCell) {
  const note = (hashCell.getNote() || "").toString().trim();
  const m = note.match(/COREHASH:([0-9a-f]{64})/i);
  return m ? m[1] : "";
}

function setCoreHashCheckpoint_(hashCell, coreHash) {
  const note = (hashCell.getNote() || "").toString();
  // Remove existing COREHASH line if present
  const cleaned = note
    .split("\n")
    .filter(line => !/^COREHASH:/i.test(line.trim()))
    .join("\n")
    .trim();

  const next = (cleaned ? cleaned + "\n" : "") + `COREHASH:${coreHash}`;
  hashCell.setNote(next);
}

// Optional: if URL is doi.org/<doi>, also set Zotero DOI field
function doiFromUrl_(u) {
  const url = (u || "").toString().trim();
  const m = url.match(/^https?:\/\/doi\.org\/(.+)$/i);
  if (!m) return "";
  return normalizeDoi_(m[1]); // you already have normalizeDoi_
}

function zoteroGetItem_(cfg, itemKey) {
  const url = `${ZOTERO_API_BASE}/users/${encodeURIComponent(cfg.libraryId)}/items/${encodeURIComponent(itemKey)}?format=json`;
  const obj = zoteroFetchRaw_(cfg, url, "get");
  return JSON.parse(obj.text); // includes key/version/data
}

function zoteroPutItemData_(cfg, itemKey, dataObj) {
  const url = `${ZOTERO_API_BASE}/users/${encodeURIComponent(cfg.libraryId)}/items/${encodeURIComponent(itemKey)}`;

  // Per Zotero docs: for single-object PUT, upload full JSON OR just the `data` object; only `data` is processed.  [oai_citation:5‡zotero.org](https://www.zotero.org/support/dev/web_api/v3/write_requests)
  const payload = JSON.stringify({ data: dataObj });

  const res = UrlFetchApp.fetch(url, {
    method: "put",
    contentType: "application/json",
    headers: {
      "Zotero-API-Key": cfg.apiKey,
      "Accept": "application/json"
    },
    payload,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(`Zotero API error ${code}: ${text}`);
  }
}

function zoteroDeleteItem_(cfg, itemKey, version) {
  const url = `${ZOTERO_API_BASE}/users/${encodeURIComponent(cfg.libraryId)}/items/${encodeURIComponent(itemKey)}`;

  const headers = {
    "Zotero-API-Key": cfg.apiKey,
    "Accept": "application/json"
  };

  // Safe deletion: only delete if version matches what we saw
  // Zotero supports If-Unmodified-Since-Version
  if (version !== undefined && version !== null) {
    headers["If-Unmodified-Since-Version"] = String(version);
  }

  const res = UrlFetchApp.fetch(url, {
    method: "delete",
    headers,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(`Zotero API error ${code}: ${text}`);
  }
}

function zoteroFetchRaw_(cfg, url, method) {
  const res = UrlFetchApp.fetch(url, {
    method: method || "get",
    headers: {
      "Zotero-API-Key": cfg.apiKey,
      "Accept": "application/json"
    },
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(`Zotero API error ${code}: ${text}`);
  }
  return { code, text };
}

function buildSheetOriginNoteHtml_(plainText) {
  const escaped = escapeHtml_((plainText || "").toString().trim());
  return `<div>
${NOTE_MARK_SHEETS_ORIGIN}
<p><strong>${SHEET_NOTE_HEADER}</strong></p>
<p>${escaped.replace(/\n/g, "<br>")}</p>
</div>`;
}

function upsertSheetNotesChild_(cfg, parentKey, plainText) {
  const childrenUrl = buildItemChildrenUrl_(cfg, parentKey, { itemType: "note", limit: 100 });
  const children = zoteroFetch_(cfg, childrenUrl);
  const notes = Array.isArray(children) ? children : [];

  // Delete notes previously imported to Sheets, but never delete the Sheets-origin note.
  // Supports marker-based logic and legacy internal tags.
  let deletedImportedNotes = 0;

  for (const n of notes) {
    const html = (n?.data?.note || "").toString();
    const tags = n?.data?.tags || [];

    const isSheetsOrigin = hasMarker_(html, NOTE_MARK_SHEETS_ORIGIN) || hasTag_(tags, LEGACY_TAG_SHEETS_ORIGIN);
    if (isSheetsOrigin) continue;

    const isImportedToSheets = hasTag_(tags, LEGACY_TAG_IMPORTED_TO_SHEETS) || hasMarker_(html, NOTE_MARK_IMPORTED);
    if (!isImportedToSheets) continue;

    try {
      zoteroDeleteItem_(cfg, n.key, n.version);
      deletedImportedNotes++;
    } catch (e) {
      Logger.log(`WARNING: Failed deleting imported-to-sheets note ${n.key} (parent ${parentKey}). Error: ${e}`);
    }
  }

  // Find existing Sheets-origin note by marker (or legacy internal tag)
  let target = null;
  for (const n of notes) {
    const html = (n?.data?.note || "").toString();
    const tags = n?.data?.tags || [];
    if (hasMarker_(html, NOTE_MARK_SHEETS_ORIGIN) || hasTag_(tags, LEGACY_TAG_SHEETS_ORIGIN)) {
      target = n;
      break;
    }
  }

  const htmlNote = buildSheetOriginNoteHtml_(plainText);

  if (target) {
    const noteData = { ...target.data };
    noteData.tags = stripInternalNoteTags_(noteData.tags).tags;
    noteData.note = htmlNote;
    noteData.version = target.version;
    zoteroPutItemData_(cfg, target.key, noteData);
  } else {
    // Create a new child note
    const createUrl = `${ZOTERO_API_BASE}/users/${encodeURIComponent(cfg.libraryId)}/items`;
    const newNote = {
      itemType: "note",
      parentItem: parentKey,
      note: htmlNote
    };

    const res = UrlFetchApp.fetch(createUrl, {
      method: "post",
      contentType: "application/json",
      headers: { "Zotero-API-Key": cfg.apiKey, "Accept": "application/json" },
      payload: JSON.stringify([newNote]),
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const text = res.getContentText();
    if (code < 200 || code >= 300) throw new Error(`Zotero API error ${code}: ${text}`);
  }

  if (deletedImportedNotes > 0) {
    Logger.log(`Parent ${parentKey}: deleted ${deletedImportedNotes} child note(s) previously imported to Sheets.`);
  }

}

function escapeHtml_(s) {
  return (s || "").toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/* -------------------- Row fingerprint -------------------- */


function ensureHashColumnHidden_(sheet, colHash) {
  try { sheet.hideColumns(colHash); } catch (e) {}
}

/* -------------------- User Guide -------------------- */

function showZoteroHelp() {
  const html = HtmlService.createHtmlOutput(`
  <html>
    <head>
      <base target="_top">
      <style>
        :root{
          --font: 16px;           /* <-- bump this */
          --fontSmall: 14px;
        }
        body{
          font-family: Arial, sans-serif;
          margin: 0;
          font-size: var(--font);
          color: #111;
        }
        .wrap{ padding: 16px 18px 18px 18px; }
        h2{ margin: 0 0 12px 0; font-size: 20px; }
        .hint{ color:#444; font-size: var(--fontSmall); margin-bottom: 12px; }

        /* Accordion */
        details{
          border: 1px solid #ddd;
          border-radius: 10px;
          padding: 10px 12px;
          margin: 10px 0;
          background: #fff;
        }
        summary{
          cursor: pointer;
          font-weight: 700;
          font-size: 16px;
          list-style: none;
          outline: none;
        }
        /* Remove default disclosure triangle (Chrome) and add our own */
        summary::-webkit-details-marker { display:none; }
        summary::before{
          content: "▸";
          display: inline-block;
          width: 18px;
          color: #333;
        }
        details[open] summary::before{ content: "▾"; }

        .section{ margin-top: 10px; }
        h3{ margin: 12px 0 6px 0; font-size: 16px; }
        p{ margin: 6px 0; line-height: 1.4; }
        li{ margin: 6px 0; }
        code{
          background: #f3f3f3;
          padding: 2px 6px;
          border-radius: 6px;
          font-size: 0.95em;
        }
        .note{
          margin-top: 10px;
          background: #fff8e1;
          border: 1px solid #ffe08a;
          padding: 10px 12px;
          border-radius: 10px;
        }
        .small{ font-size: var(--fontSmall); color:#555; }
        .footer{
          margin-top: 14px;
          padding-top: 10px;
          border-top: 1px solid #eee;
          font-size: var(--fontSmall);
          color:#444;
        }
      </style>
    </head>

    <body>
      <div class="wrap">
        <h2>Zotero ↔ Google Sheets Plugin — Help</h2>
        <p class="hint"><a href="https://www.loom.com/share/39dc7ec5ae504a048e516bc89dd1aa61" target="_blank" rel="noopener">Tutorial here</a></p>

        <details open>
          <summary>1) Setting up</summary>
          <div class="section">
            <ol>
              <li>
                Get your Zotero credentials:
                <ul>
                  <li><b>API key</b>: <a href="https://www.zotero.org" target="_blank" rel="noopener">www.zotero.org</a> → <b>Settings</b> → <b>Security</b> → <b>Create new private key</b></li> → <b>Allow library, notes, and write access</b>
                  <li><b>Library ID</b>: Zotero web → <b>Settings</b> → <b>Security</b> (shown as your <b>User ID</b>)</li>
                </ul>
              </li>
              <li>Open <b>Extensions → Apps Script</b>.</li>
              <li>Set Script Properties:
                <ul>
                  <li><code>ZOTERO_LIBRARY_ID</code></li>
                  <li><code>ZOTERO_API_KEY</code></li>
                  <li><code>ZOTERO_INCLUDE_NOTES</code> <code>true</code></li>
                  <li><code>ZOTERO_LIBRARY_TYPE</code> <code>user</code></li>
                </ul>
              </li>
              <li>Reload the spreadsheet so the <b>Zotero</b> menu appears.</li>
            </ol>
          </div>
        </details>

        <details>
          <summary>2) Import reading list from Zotero</summary>
          <div class="section">
            <ol>
              <li>In Zotero, tag references you want to import with <b><code>reading list</code></b>.</li>
              <li>Run <b>Zotero → Import reading list from Zotero</b> when you want to import your references into Sheets.</li>
            </ol>
            <p>This will:</p>
            <ul>
              <li>Pull all items tagged <code>reading list</code></li>
              <li>Fill: <b>Paper</b> (hyperlinked), <b>Authors</b>, <b>Year</b>, <b>Theme</b> (from Zotero tags)<b>, Notes</b></li>
            </ul>
          </div>
        </details>

        <details>
          <summary>3) Export changes to Zotero</summary>
          <div class="section">
            <p>Run <b>Zotero → Export changes to Zotero</b> when you want Sheets edits reflected in Zotero.</p>
            <p>This will:</p>
            <ul>
              <li>Update Zotero title, URL (including hyperlink deletions), and tags (Theme + Status)</li>
              <li>Export your Sheet Notes into Zotero as a note headed <b>“Imported from Google Sheets”</b></li>
              <li>Remove <code>reading list</code> in Zotero for references you deleted from the Sheet (snapshot-based)</li>
            </ul>
            <div class="note">
              <b>Note:</b> If you edit core bibliographic fields (Paper/Authors/Year), you will be prompted to confirm export.
            </div>
          </div>
        </details>

        <details>
          <summary>4) Import new Zotero notes</summary>
          <div class="section">
            <p>Run <b>Zotero → Import new Zotero notes</b> if you created notes directly in Zotero and want them in Sheets.</p>
            <ul>
              <li>Scans notes under each <code>reading list</code> item</li>
              <li>Appends new notes to the <b>Notes</b> cell in Sheets</li>
              <li>Marks imported notes in Zotero so they won’t be appended again</li>
            </ul>
          </div>
        </details>

        <details>
          <summary>5) Usage notes & troubleshooting</summary>
          <div class="section">
            <h3>Usage notes</h3>
            <ul>
              <li>Use Sheets as your dashboard: update <b>Status</b> and keep <b>Notes</b> here.</li>
              <li>Imports won’t overwrite Status/Notes.</li>
              <li>Theme dropdown options refresh from Zotero tags at import.</li>
              <li>You can adjust Theme styling via <b>Data → Data validation rules</b>.</li>
            </ul>

            <h3>Troubleshooting</h3>
            <ul>
              <li><b>Menu missing:</b> reload the spreadsheet.</li>
              <li><b>Config error:</b> confirm Script Properties exist: <code>ZOTERO_LIBRARY_ID</code>, <code>ZOTERO_API_KEY</code>.</li>
              <li><b>Hyperlinks broken:</b> ensure URLs include <code>https://</code>.</li>
              <li><b>Group library:</b> endpoints must use <code>/groups/&lt;id&gt;/...</code>.</li>
            </ul>

            <div class="footer">
              Questions / feedback: <b>frdfaa2@cam.ac.uk</b>
            </div>
          </div>
        </details>
      </div>
    </body>
  </html>
  `).setWidth(760).setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, "Zotero Plugin Help");
}
