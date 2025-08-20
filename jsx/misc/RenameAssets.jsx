#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * RenameAssets.jsx
 *
 * 概要 / Overview:
 *   グラフィックスタイル／ブラシ／スウォッチ／シンボルの「名前」を、
 *   ダイアログで指定した「検索→置換」で一括変更します。プレビューで事前確認可能。
 *   Renames names of Graphic Styles, Brushes, Swatches, and Symbols in bulk
 *   using a dialog-based Find/Replace. Preview lets you inspect changes before applying.
 *
 * 特徴 / Features:
 *   - 検索・置換／正規表現／大文字小文字無視
 *     Find/Replace, Regex, and Ignore Case
 *   - 対象の種類をラジオボタンで選択（スタイル／ブラシ／スウォッチ／シンボル）
 *     Choose target collection via radio buttons (Style/Brush/Swatch/Symbol)
 *   - プレビューはボタン押下時のみ実行（キャンセルで元に復帰）
 *     Preview runs only when pressing the button (Cancel restores original names)
 *   - 既存名との重複は (2),(3)… を自動付番で回避
 *     Auto-deduplicate by appending (2), (3), …
 *   - 角括弧で囲まれた既定名はスキップ
 *     Bracketed default names are skipped
 *
 * 更新履歴 / Update History:
 *   - 2025-08-20: v1.0 初版。ダイアログ（検索/置換、正規表現、大小無視、対象選択）、
 *                 プレビューボタン、ボタン横並びレイアウト、透明度/位置調整、
 *                 ローカライズ（日本語/英語）、成功時メッセージ非表示を実装。
 */

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var lang = getCurrentLang();

/* ローカライズ定義（キー→{ja,en} 形式） / Localization in key→{ja,en} form */
var LABELS = {
  dialogTitle: { ja: "グラフィック名の置換 " + SCRIPT_VERSION, en: "Rename Graphic Names " + SCRIPT_VERSION },
  panelTarget: { ja: "対象", en: "Target" },
  radioStyle:  { ja: "グラフィックスタイル", en: "Graphic Styles" },
  radioBrush:  { ja: "ブラシ", en: "Brushes" },
  radioSwatch: { ja: "スウォッチ", en: "Swatches" },
  radioSymbol: { ja: "シンボル", en: "Symbols" },
  find:        { ja: "検索文字列:", en: "Find:" },
  replace:     { ja: "置換文字列:", en: "Replace:" },
  regex:       { ja: "正規表現", en: "Regular Expression" },
  ignoreCase:  { ja: "大文字／小文字を無視", en: "Ignore Case" },
  statusNone:  { ja: "—", en: "—" },
  statusNoFind:{ ja: "検索文字列なし", en: "Find string is empty" },
  statusPreview:{
    ja: function(changed,total){ return "プレビュー: 変更 " + changed + "／全 " + total; },
    en: function(changed,total){ return "Preview: " + changed + " changed / " + total + " total"; }
  },
  btnCancel:   { ja: "キャンセル", en: "Cancel" },
  btnPreview:  { ja: "プレビュー", en: "Preview" },
  btnOK:       { ja: "OK", en: "OK" },
  alertNoDoc:  { ja: "ドキュメントが開かれていません。\nNo document is open.", en: "No document is open." },
  alertNoAny:  { ja: "このドキュメントには対象となる項目（スタイル／ブラシ／スウォッチ／シンボル）がありません。\nNo styles/brushes/swatches/symbols in this document.", en: "No styles/brushes/swatches/symbols in this document." },
  alertNoFind: { ja: "「検索」文字列が空です。\nFind string is empty.", en: "Find string is empty." },
  alertNoTarget:{ ja: "選択した対象に名前付き項目がありません。\nNo items found for selected target.", en: "No items found for selected target." }
};

// Localize helpers
function t(key){ var m = LABELS[key]; return m ? (lang === 'ja' ? m.ja : m.en) : key; }
function tf(key){ var m = LABELS[key]; return m ? (lang === 'ja' ? m.ja : m.en) : function(){ return ""; }; }

function main() {
    if (app.documents.length === 0) {
        alert(t('alertNoDoc'));
        return;
    }
    var doc = app.activeDocument;
    // Validate that the document has at least one eligible collection
    var hasAny = (doc.graphicStyles && doc.graphicStyles.length) ||
        (doc.brushes && doc.brushes.length) ||
        (doc.swatches && doc.swatches.length) ||
        (doc.symbols && doc.symbols.length);
    if (!hasAny) {
        alert(t('alertNoAny'));
        return;
    }

    var params = showDialog();
    if (!params) return; // canceled

    var findStr = params.find;
    var replStr = params.repl;
    var useRegex = params.regex;
    var ignoreCase = params.ignoreCase;
    var target = params.target;

    if (!findStr) {
        alert(t('alertNoFind'));
        return;
    }

    var collection, totalCount;
    if (target === "brush") {
        collection = app.activeDocument.brushes;
        totalCount = app.activeDocument.brushes.length;
    } else if (target === "swatch") {
        collection = app.activeDocument.swatches;
        totalCount = app.activeDocument.swatches.length;
    } else if (target === "symbol") {
        collection = app.activeDocument.symbols;
        totalCount = app.activeDocument.symbols.length;
    } else {
        collection = app.activeDocument.graphicStyles;
        totalCount = app.activeDocument.graphicStyles.length;
    }

    if (!collection || totalCount === 0) {
        alert(t('alertNoTarget'));
        return;
    }

    renameItemsByCollection(collection, findStr, replStr, useRegex, ignoreCase);
}

// entry point
main();

/* ---------------- Helpers ---------------- */


/*
 * 置換関数のビルド / Build a replacer function
 * returns: function(oldName) -> newName or null
 */
function buildReplacer(findStr, replStr, useRegex, ignoreCase) {
    if (useRegex) {
        var re = null;
        try {
            re = new RegExp(findStr, ignoreCase ? 'gi' : 'g');
        } catch (e) {
            re = null;
        }
        if (!re) return function() {
            return null;
        };
        return function(oldName) {
            return re.test(oldName) ? oldName.replace(re, replStr) : null;
        };
    } else if (ignoreCase) {
        // compile once
        var reCI = null;
        try {
            reCI = new RegExp(findStr, 'gi');
        } catch (e) {
            reCI = null;
        }
        if (!reCI) return function() {
            return null;
        };
        return function(oldName) {
            return reCI.test(oldName) ? oldName.replace(reCI, replStr) : null;
        };
    } else {
        return function(oldName) {
            return oldName.indexOf(findStr) > -1 ? oldName.split(findStr).join(replStr) : null;
        };
    }
}

/*
 * 変更案の作成（未適用）/ Compute proposed names without mutating
 * returns: { names:Array, changed:Number, skipped:Number }
 */
function computeProposedNames(collection, findStr, replStr, useRegex, ignoreCase) {
    var items = collection;
    var total = items.length;
    var changed = 0,
        skipped = 0;

    // Snapshot original names for collision map
    var original = [];
    var used = {};
    for (var i = 0; i < total; i++) {
        var n = items[i].name;
        var nStr = (n === undefined || n === null) ? "" : String(n);
        original.push(nStr);
        used[nStr] = true;
    }

    var proposed = original.slice(0);
    var replacer = buildReplacer(findStr, replStr, useRegex, ignoreCase);
    for (var j = 0; j < total; j++) {
        var oldName = original[j];
        if (oldName === undefined || oldName === null) {
            skipped++;
            continue;
        }
        var oldStr = String(oldName);
        // Skip bracketed default (use inline regex to avoid scope issues in some ExtendScript versions)
        if (/^\[.*\]$/.test(oldStr)) {
            skipped++;
            continue;
        }

        var next = replacer(oldStr);

        if (next !== null && next !== oldStr) {
            // Ensure uniqueness against used map (based on original names)
            var unique = ensureUniqueName(next, used);
            proposed[j] = unique;
            used[unique] = true;
            changed++;
        }
    }
    return {
        names: proposed,
        changed: changed,
        skipped: skipped
    };
}

// 変更の適用 / Apply proposed names
function applyNames(collection, names) {
    var count = 0;
    for (var i = 0; i < collection.length && i < names.length; i++) {
        var cur = collection[i].name;
        var next = names[i];
        if (cur !== next) {
            try {
                collection[i].name = next;
                count++;
            } catch (e) {
                /* ignore */
            }
        }
    }
    return count;
}

/*
 * ダイアログ生成 / Build dialog
 */
function showDialog() {
    var dlg = new Window("dialog", t('dialogTitle'));
    // Adjust dialog position and opacity
    var offsetX = 300;
    var offsetY = 0;
    var dialogOpacity = 0.98;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function() {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        try {
            dlg.opacity = opacityValue;
        } catch (e) {}
    }

    setDialogOpacity(dlg, dialogOpacity);
    shiftDialogPosition(dlg, offsetX, offsetY);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];
    dlg.margins = 16;
    dlg.spacing = 12;

    // Target radio group (top)
    var panelTarget = dlg.add("panel", undefined, t('panelTarget'));
    panelTarget.orientation = "row";
    panelTarget.alignChildren = ["left", "center"];
    panelTarget.margins = [15, 20, 15, 10];
    panelTarget.spacing = 16;

    var rbStyles = panelTarget.add("radiobutton", undefined, t('radioStyle'));
    var rbBrushes = panelTarget.add("radiobutton", undefined, t('radioBrush'));
    var rbSwatch = panelTarget.add("radiobutton", undefined, t('radioSwatch'));
    var rbSymbol = panelTarget.add("radiobutton", undefined, t('radioSymbol'));
    rbStyles.value = true; // default

    var labelWidth = 120; // ラベル幅を揃える

    function addLabeledEdit(parent, labelText, width) {
        var row = parent.add('group');
        var lbl = row.add('statictext', undefined, labelText);
        lbl.preferredSize.width = width;
        lbl.justify = 'right';
        var et = row.add('edittext', undefined, '');
        et.characters = 30;
        return et;
    }

    var inputFind = addLabeledEdit(dlg, t('find'), labelWidth);
    var inputRepl = addLabeledEdit(dlg, t('replace'), labelWidth);

    // Regex + Case-insensitive options (side by side, centered)
    var rowOpts = dlg.add("group");
    rowOpts.orientation = "row";
    rowOpts.alignChildren = ["center", "center"];
    rowOpts.alignment = "center";
    var cbRegex = rowOpts.add("checkbox", undefined, t('regex'));
    cbRegex.value = false; // default OFF
    var cbCase = rowOpts.add("checkbox", undefined, t('ignoreCase'));
    cbCase.value = false; // default OFF

    // Preview status (button-driven preview only)
    var rowStatus = dlg.add('group');
    rowStatus.orientation = 'row';
    rowStatus.alignChildren = ['center', 'center'];
    rowStatus.alignment = 'center';
    var stStatus = rowStatus.add('statictext', undefined, '');
    stStatus.characters = 30;
    stStatus.alignment = 'center';
    try {
        stStatus.justify = 'center';
    } catch (e) {}

    // --- Preview engine ---
    var snapshot = null; // original names
    var snapTarget = null; // which target the snapshot belongs to

    function getTargetKey() {
        return rbBrushes.value ? 'brush' : (rbSwatch.value ? 'swatch' : (rbSymbol.value ? 'symbol' : 'style'));
    }

    function getCollectionByKey(key) {
        var d = app.activeDocument;
        if (!d) return null;
        if (key === 'brush') return d.brushes;
        if (key === 'swatch') return d.swatches;
        if (key === 'symbol') return d.symbols;
        return d.graphicStyles;
    }

    // Consolidated snapshot helper: build (if needed), restore, then run work
    function withSnapshot(targetKey, workFn){
        var col = getCollectionByKey(targetKey);
        if (!col) return null;
        // take snapshot if missing or target changed
        if (!snapshot || snapTarget !== targetKey){
            snapshot = [];
            for (var i = 0; i < col.length; i++) snapshot[i] = col[i].name;
            snapTarget = targetKey;
        }
        // always restore to snapshot first (avoid cumulative preview)
        for (var j = 0; j < col.length && j < snapshot.length; j++){
            try { col[j].name = snapshot[j]; } catch(e) {}
        }
        // run the work on restored collection
        try { workFn && workFn(col); } catch(e) {}
        return col;
    }

    // Keep a thin restore for Cancel button
    function restoreSnapshot(){
        if (!snapshot) return;
        var col = getCollectionByKey(snapTarget);
        if (!col) return;
        for (var i = 0; i < col.length && i < snapshot.length; i++){
            try { col[i].name = snapshot[i]; } catch(e){}
        }
    }

    function updatePreview() {
        var key = getTargetKey();
        var findTxt = String(inputFind.text || '');
        var replTxt = String(inputRepl.text || '');
        var useRe = cbRegex.value === true;
        var ic = cbCase.value === true;
        if (!findTxt) { stStatus.text = t('statusNoFind'); return; }

        var lastPlan = null;
        var col = withSnapshot(key, function(restoredCol){
            lastPlan = computeProposedNames(restoredCol, findTxt, replTxt, useRe, ic);
            applyNames(restoredCol, lastPlan.names);
        });
        if (!col) { stStatus.text = t('statusNone'); return; }
        stStatus.text = tf('statusPreview')(lastPlan.changed, col.length);
        app.redraw();
    }

    // Remove automatic/live preview event wiring

    // Keep radio-button handlers but stop auto-preview
    var rbs = [rbStyles, rbBrushes, rbSwatch, rbSymbol];
    for (var i = 0; i < rbs.length; i++) {
        rbs[i].onClick = function() {
            snapshot = null; // clear snapshot so next manual preview rebuilds
        };
    }
    // Remove checkbox click handler entirely

    // Buttons layout (main row with left/right groups and stretch spacer)
    var btnRowGroup = dlg.add("group");
    btnRowGroup.orientation = "row";
    // 
    btnRowGroup.margins = [10, 10, 10, 0];
    btnRowGroup.alignment = ["fill", "bottom"];

    // Left-side buttons
    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    var btnCancel = btnLeftGroup.add("button", undefined, t('btnCancel'), {
        name: "cancel"
    });

    // Stretch spacer
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    // Right-side buttons
    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    var btnPreview = btnRightGroup.add("button", undefined, t('btnPreview'), {
        name: "preview"
    });
    var btnOK = btnRightGroup.add("button", undefined, t('btnOK'), {
        name: "ok"
    });

    // Handlers
    btnOK.onClick = function() {
        // keep previewed names as final
        accepted = true;
        snapshot = null; // drop snapshot to avoid unintended restore
        dlg.close();
    };
    btnCancel.onClick = function() {
        // revert to original names if preview applied
        try {
            restoreSnapshot();
        } catch (e) {}
        dlg.close();
    };
    btnPreview.onClick = function() {
        updatePreview();
    };

    // 

    // UX niceties
    inputFind.active = true;

    var accepted = false;

    dlg.show();

    return accepted ? {
        find: String(inputFind.text || ""),
        repl: String(inputRepl.text || ""),
        regex: cbRegex.value,
        ignoreCase: cbCase.value,
        target: rbBrushes.value ? "brush" : (rbSwatch.value ? "swatch" : (rbSymbol.value ? "symbol" : "style"))
    } : null;
}

function renameItemsByCollection(collection, findStr, replStr, useRegex, ignoreCase) {
    var items = collection;
    var total = items.length;
    var changed = 0;
    var skipped = 0;

    // Snapshot existing names for fast collision checks
    var used = {};
    for (var i = 0; i < total; i++) {
        var nm = items[i].name;
        var nmStr = (nm === undefined || nm === null) ? "" : String(nm);
        used[nmStr] = true;
    }

    var replacer = buildReplacer(findStr, replStr, useRegex, ignoreCase);

    for (var j = 0; j < total; j++) {
        var item = items[j];
        var oldName = item.name;
        if (oldName === undefined || oldName === null) {
            skipped++;
            continue;
        }
        var oldStr = String(oldName);
        // Skip bracketed default (inline regex)
        if (/^\[.*\]$/.test(oldStr)) {
            skipped++;
            continue;
        }
        var proposed = replacer(oldStr);

        if (proposed !== null && proposed !== oldStr) {
            var unique = ensureUniqueName(proposed, used);
            if (unique !== oldStr) {
                try {
                    item.name = unique;
                    used[unique] = true;
                    // old name stays in map, but that's fine (we just avoid new collisions)
                    changed++;
                } catch (e) {
                    // If rename fails, continue gracefully
                }
            }
        }
    }
    return {
        changed: changed,
        skipped: skipped
    };
}

// Append " (2)", " (3)", ... until unique
function ensureUniqueName(base, usedMap) {
    if (!usedMap[base]) return base;
    var i = 2;
    var candidate = base + " (" + i + ")";
    while (usedMap[candidate]) {
        i++;
        candidate = base + " (" + i + ")";
        // safeguard
        if (i > 9999) break;
    }
    return candidate;
}