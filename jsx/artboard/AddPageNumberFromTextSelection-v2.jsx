#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

AddPageNumberFromTextSelection.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- 選択中のポイントテキスト（_pagenumber レイヤー推奨）を雛形に、すべてのアートボードへ連番テキストを配置します。
- 値の変更（接頭辞・開始番号・接尾辞・ゼロパディング・総ページ）や ↑↓/Shift+↑↓ による数値変更で、ドキュメント上のプレビューが即時更新されます（OK は確定のみ、キャンセルでプレビューを破棄）。
- _pagenumber レイヤーが無い場合は自動作成し、プレビュー中は一時的にロック解除・可視化・最前面移動して安全に処理します。OK/キャンセル時に元のロック／可視状態と重ね順を確実に復元します。

### 主な機能：

- 開始番号／接頭辞／接尾辞／ゼロパディング／総ページの指定
- 3カラムUI、ダイアログの透明度・表示位置の調整
- テンプレートが1件のみでも、プレビュー中に全アートボードへ仮増殖して確認可能（内部: pasteInAllArtboard）
- _pagenumber レイヤーの自動作成・一時解除・最前面化・確定/取消時の復元
- ダイアログタイトルにバージョン番号を表示（SCRIPT_VERSION）

### 処理の流れ：

1) 適当なポイントテキストを選択（_pagenumber レイヤー推奨）
2) ダイアログで各値を設定（↑↓/Shift+↑↓で数値調整、Zでゼロ埋め、Aで総ページ表示を切替）
3) 値変更ごとにライブプレビューを確認（プレビュー中は _pagenumber を一時解除＋最前面化）
4) OK で確定（プレビュー用バックアップ破棄＆状態復元）／キャンセルでプレビューを破棄して完全復元

### note：

- 対象はポイントテキストです。段落揃えなどの体裁変更は行いません。

### 更新履歴：

- v1.9 (20250810) : _pagenumber の自動作成、ロック/可視の一時解除、最前面化、OK/キャンセル時の重ね順・可視・ロックの完全復元を追加。プレビュー耐性を強化。
- v1.8 (20250625) : リファクタリング、ライブプレビューと全ABプレビュー複製に対応、UI整理
- v1.0 (20250625) : 初期バージョン

---

### Script Name：

AddPageNumberFromTextSelection.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### Overview：

- Places sequential page-number text on all artboards using the currently selected point text as a template (the _pagenumber layer is recommended).
- Live preview updates instantly when values change (prefix, start number, suffix, zero padding, total pages) and on Up/Down or Shift+Up/Down. OK commits; Cancel discards the preview.
- If the _pagenumber layer doesn’t exist, it is auto-created. During preview, the layer is temporarily unlocked, shown, and moved to the top; on OK/Cancel, its lock/visibility and stacking order are fully restored.

### Main Features：

- Configure start number / prefix / suffix / zero padding / total pages
- 3-column dialog UI; adjustable dialog opacity and position
- Preview can temporarily duplicate a single template across all artboards (internally: pasteInAllArtboard)
- Auto-create, temporarily unlock/show/move-to-top for _pagenumber; fully restore state on commit/cancel
- Version number shown in the dialog title (SCRIPT_VERSION)

### Flow：

1) Select a point text object (the _pagenumber layer is recommended)
2) Set values in the dialog (Up/Down or Shift+Up/Down; Z toggles zero-pad; A toggles total pages)
3) Live preview updates every change (_pagenumber is temporarily prepared and brought to front)
4) Click OK to commit (discard preview backup & restore state) or Cancel to discard and fully restore

### Notes：

- Target must be a point text. This script does not change paragraph alignment or styling.

### Update History：

- v1.9 (2025-08-10): Added auto-create for _pagenumber, temporary unlock/visibility, move-to-top during preview, and full restoration of stacking order & visibility/lock on OK/Cancel; improved preview robustness.
- v1.8 (2025-06-25): Refactor; added live preview and all-artboards preview duplication; UI cleanup
- v1.0 (2025-06-25): Initial release
*/

var SCRIPT_VERSION = "v1.9";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}

// プレビュー用一時レイヤー名 / Temp layer for preview backup
var TMP_LAYER_NAME = "_pagenumber_preview";
var lang = getCurrentLang();

var LABELS = {
    // 1) ダイアログタイトル / Dialog title
    dialogTitle: {
        ja: "ページ番号を追加 " + SCRIPT_VERSION,
        en: "Add Page Numbers " + SCRIPT_VERSION
    },
    // 2) 接頭辞 / Prefix
    prefixLabel: {
        ja: "接頭辞",
        en: "Prefix"
    },
    // 3) 開始番号 / Starting number
    promptMessage: {
        ja: "開始番号",
        en: "Starting number"
    },
    // 4) ゼロパディング / Zero pad
    zeroPadLabel: {
        ja: "ゼロパディング",
        en: "Zero pad"
    },
    // 5) 接尾辞 / Suffix
    suffixLabel: {
        ja: "接尾辞",
        en: "Suffix"
    },
    // 6) 総ページ数を表示 / Show total pages
    totalPageLabel: {
        ja: "総ページ数を表示",
        en: "Show total pages"
    },
    // 7) OK / Cancel ボタン
    okLabel: {
        ja: "OK",
        en: "OK"
    },
    cancelLabel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    // 8) エラーメッセージ / Error messages
    errorNotNumber: {
        ja: "有効な数字を入力してください",
        en: "Please enter a valid number"
    },
    errorInvalidSelection: {
        ja: "ポイントテキストを選択してください（_pagenumber レイヤー推奨）",
        en: "Please select a point text (the _pagenumber layer is recommended)."
    }
};

/* レイヤー取得 or 作成 / Get or create a layer by name */
function getOrCreateLayer(doc, name) {
    var layer = null;
    try {
        layer = doc.layers.getByName(name);
    } catch (e) {
        layer = doc.layers.add();
        layer.name = name;
    }
    return layer;
}

/* レイヤーを確実に削除（unlock→内容削除→remove）/ Force-remove a layer by name */
function forceRemoveLayerByName(doc, name) {
    try {
        var ly = doc.layers.getByName(name);
        try {
            ly.locked = false;
        } catch (_l0) {}
        try {
            ly.visible = true;
        } catch (_v0) {}
        try {
            for (var i = ly.pageItems.length - 1; i >= 0; i--) {
                trySet(ly.pageItems[i], 'locked', false);
                tryCall(function() {
                    ly.pageItems[i].remove();
                });
            }
            for (var j = ly.layers.length - 1; j >= 0; j--) {
                tryCall(function() {
                    ly.layers[j].remove();
                });
            }
        } catch (_purge) {}
        try {
            ly.remove();
        } catch (_r3) {}
    } catch (_no) {}
}

// --- Safe helpers to reduce repetitive try/catch ---
function trySet(obj, prop, value) {
    try {
        if (obj) obj[prop] = value;
    } catch (_) {}
}

function tryCall(fn) {
    try {
        fn && fn();
    } catch (_) {}
}

/* 型ガード / Type guards */
function isTextFrame(obj) {
    return obj && obj.typename === "TextFrame";
}

/* 選択中のTextFrameを返す（なければnull）/ Get selected TextFrame or null */
function getSelectedTextFrame() {
    try {
        if (app.selection && app.selection.length > 0 && app.selection[0]) {
            return isTextFrame(app.selection[0]) ? app.selection[0] : null;
        }
    } catch (_) {}
    return null;
}


/* =========================================================
   処理ヘルパー / Processing Helpers (no behavior change)
   ========================================================= */

/* =========================================================
   _pagenumber レイヤー状態の保存・一時解除・復元 / Save, prepare, and restore state
   ========================================================= */

function getLayerIndex(doc, layer) {
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i] === layer) return i;
    }
    return -1;
}

function getLayerNameAtIndex(doc, idx) {
    try {
        return doc.layers[idx].name;
    } catch (e) {}
    return null;
}

function capturePagenumberState(doc, layer) {
    var st = {
        locked: true,
        visible: true,
        index: -1,
        neighborAboveName: null
    };
    try {
        st.locked = layer.locked;
    } catch (_1) {}
    try {
        st.visible = layer.visible;
    } catch (_2) {}
    st.index = getLayerIndex(doc, layer);
    // ひとつ上（前面側）のレイヤー名を記録（復元基準）/ remember neighbor above for restore
    if (st.index > 0) {
        st.neighborAboveName = getLayerNameAtIndex(doc, st.index - 1);
    } else {
        st.neighborAboveName = null; // もともと最前面
    }
    return st;
}

function preparePagenumberForWork(doc, layer) {
    // ロック解除＆可視化 / unlock & show
    try {
        layer.locked = false;
    } catch (_l0) {}
    try {
        layer.visible = true;
    } catch (_v0) {}
    // 最上位へ / move to top
    try {
        layer.move(doc, ElementPlacement.PLACEATBEGINNING);
    } catch (_m0) {}
}

function restorePagenumberState(doc, layer, st) {
    if (!st) return;

    // レイヤー順序を復元 / restore stacking order
    try {
        if (st.neighborAboveName) {
            // 直上レイヤーの「後ろ（下）」へ移動 / place after the neighbor that was above
            var above = null;
            try {
                above = doc.layers.getByName(st.neighborAboveName);
            } catch (_nf) {
                above = null;
            }
            if (above) {
                try {
                    layer.move(above, ElementPlacement.PLACEAFTER);
                } catch (_ma) {}
            }
        } else {
            // もともと最前面なら最前面へ / if originally top, keep at top
            try {
                layer.move(doc, ElementPlacement.PLACEATBEGINNING);
            } catch (_mt) {}
        }
    } catch (_ord) {}

    // 可視・ロック状態を復元 / restore visibility & lock
    try {
        layer.visible = st.visible;
    } catch (_rv) {}
    try {
        layer.locked = st.locked;
    } catch (_rl) {}
}

/* _pagenumber レイヤー取得 or 作成 / Get or create the _pagenumber layer */
function getOrCreatePagenumberLayer(doc) {
    var ly = null;
    try {
        ly = doc.layers.getByName("_pagenumber");
    } catch (e) {
        ly = doc.layers.add();
        ly.name = "_pagenumber";
    }
    return ly;
}

/* 選択を指定レイヤーへ移動（先頭へ）/ Move current selection into layer (place at beginning) */
function moveSelectionToLayer(doc, layer) {
    var tf = getSelectedTextFrame();
    if (tf) {
        try {
            tf.move(layer, ElementPlacement.PLACEATBEGINNING);
        } catch (_) {}
        return tf;
    }
    return null;
}

/* 指定アートボード上の最初のTextFrameを取得 / Get first TextFrame on a given artboard */
function getFirstTextFrameOnArtboard(layer, doc, abIndex) {
    var rect = doc.artboards[abIndex].artboardRect;
    for (var i = 0; i < layer.textFrames.length; i++) {
        var tf = layer.textFrames[i];
        var p = tf.position;
        if (p[0] >= rect[0] && p[0] <= rect[2] && p[1] <= rect[1] && p[1] >= rect[3]) {
            return tf;
        }
    }
    return null;
}

/* どのアートボードでも最初に見つかったTextFrameを返す / Find first TextFrame on any artboard */
function findTextFrameOnAnyArtboard(layer, doc) {
    var tf = getFirstTextFrameOnArtboard(layer, doc, 0);
    if (tf) return tf;
    for (var j = 1; j < doc.artboards.length; j++) {
        tf = getFirstTextFrameOnArtboard(layer, doc, j);
        if (tf) return tf;
    }
    return null;
}

/* 位置を基準ABへ合わせる / Rebase a text frame to the base artboard (index) */
function rebaseToArtboard(doc, textFrame, targetABIndex) {
    var pos = textFrame.position;
    var baseRect = doc.artboards[targetABIndex].artboardRect;
    var currentIndex = getArtboardIndexByPosition(doc, pos);
    if (currentIndex >= 0) {
        var currentRect = doc.artboards[currentIndex].artboardRect;
        var dx = baseRect[0] - currentRect[0];
        var dy = baseRect[1] - currentRect[1];
        textFrame.position = [pos[0] + dx, pos[1] + dy];
    }
}

/* カット＆全AB貼付（安全化付き）/ Cut & paste in all artboards (temporarily unlock/show) */
function seedAndPasteAll(doc, targetText, startNum) {
    // 内容を初期化 / Seed contents
    trySet(targetText, 'contents', String(startNum));

    var srcLayer = null;
    var prevLyLocked = null,
        prevLyVisible = null;
    try {
        srcLayer = targetText.layer;
    } catch (_) {
        srcLayer = null;
    }
    if (srcLayer) {
        try {
            prevLyLocked = srcLayer.locked;
        } catch (_) {
            prevLyLocked = null;
        }
        try {
            prevLyVisible = srcLayer.visible;
        } catch (_) {
            prevLyVisible = null;
        }
    }
    trySet(targetText, 'locked', false);
    if (srcLayer) {
        trySet(srcLayer, 'locked', false);
        trySet(srcLayer, 'visible', true);
    }

    // 元のアートボードをアクティブ化 / Activate the source artboard
    try {
        var srcIdx = getArtboardIndexByPosition(doc, targetText.position);
        if (srcIdx >= 0) {
            doc.artboards.setActiveArtboardIndex(srcIdx);
        }
    } catch (_) {}

    // 選択→カット→全AB貼付 / Select → Cut → Paste in all artboards
    tryCall(function() {
        app.selection = null;
    });
    trySet(targetText, 'selected', true);
    tryCall(function() {
        app.cut();
    });

    // 貼付先レイヤーを明示（対象テキストの所属レイヤーへ）/ ensure paste target layer
    try {
        if (targetText.layer) doc.activeLayer = targetText.layer;
    } catch (_) {}

    tryCall(function() {
        app.executeMenuCommand('pasteInAllArtboard');
    });

    // レイヤーの一時状態を復元 / Restore layer state
    if (srcLayer) {
        if (prevLyLocked !== null) trySet(srcLayer, 'locked', prevLyLocked);
        if (prevLyVisible !== null) trySet(srcLayer, 'visible', prevLyVisible);
    }
}

/* _pagenumber 上のTFを AB順に収集 / Collect text frames on _pagenumber sorted by AB index */
function collectTextFramesSorted(doc, layer, exclude) {
    var list = [];
    for (var i = 0; i < layer.textFrames.length; i++) {
        var tf = layer.textFrames[i];
        if (exclude && tf === exclude) continue;
        var idx = getArtboardIndexByPosition(doc, tf.position);
        list.push({
            frame: tf,
            abIdx: idx
        });
    }
    list.sort(function(a, b) {
        return a.abIdx - b.abIdx;
    });
    return list;
}

/* 連番を流し込む / Apply numbering contents */
function applyNumbering(framesList, startNum, maxDigits, prefix, suffix, zeroPad, maxNum, showTotal) {
    for (var i = 0; i < framesList.length; i++) {
        var num = startNum + i;
        var txt = buildPageNumberString(num, maxDigits, prefix, suffix, zeroPad, maxNum, showTotal);
        try {
            framesList[i].frame.contents = txt;
        } catch (_) {}
    }
}

function removeOtherTextFrames(layer, except) {
    var frames = layer.textFrames;
    for (var i = frames.length - 1; i >= 0; i--) {
        var f = frames[i];
        if (f === except) continue; // keep target
        if (f.typename === "TextFrame") {
            try {
                f.locked = false;
            } catch (_) {}
            try {
                f.remove();
            } catch (_) {}
        }
    }
}

function getArtboardIndexByPosition(doc, pos) {
    for (var i = 0; i < doc.artboards.length; i++) {
        var abRect = doc.artboards[i].artboardRect;
        if (pos[0] >= abRect[0] && pos[0] <= abRect[2] && pos[1] <= abRect[1] && pos[1] >= abRect[3]) {
            return i;
        }
    }
    return -1;
}


function buildPageNumberString(num, maxDigits, prefix, suffix, zeroPad, totalPageNum, showTotal) {
    var s = String(num);
    if (zeroPad && s.length < maxDigits) {
        s = Array(maxDigits - s.length + 1).join("0") + s; // ES3-safe zero pad
    }
    var out = (prefix || "") + s + (suffix || "");
    if (showTotal) out += "/" + totalPageNum;
    return out;
}

/* 値の増減（↑↓／Shift+↑↓）/ Increment-Decrement by Arrow Keys (Shift for ±10) */
function changeValueByArrowKey(editText, onChanged) {
    if (!editText || !editText.addEventListener) return;
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = keyboard.shiftKey ? 10 : 1;
        var k = event.keyName;
        var isUp = (k === "Up" || k === "UpArrow");
        var isDown = (k === "Down" || k === "DownArrow");

        if (isUp) {
            if (keyboard.shiftKey) {
                value = Math.ceil((value + 1) / delta) * delta; // 10の倍数へスナップ / snap to 10s
            } else {
                value += delta;
            }
            if (value < 0) value = 0;
            editText.text = Math.round(value);
            event.preventDefault();
            if (typeof onChanged === "function") {
                try {
                    onChanged();
                } catch (_e1) {}
            }
        } else if (isDown) {
            if (keyboard.shiftKey) {
                value = Math.floor((value - 1) / delta) * delta; // 10の倍数へスナップ / snap to 10s
            } else {
                value -= delta;
            }
            if (value < 0) value = 0;
            editText.text = Math.round(value);
            event.preventDefault();
            if (typeof onChanged === "function") {
                try {
                    onChanged();
                } catch (_e2) {}
            }
        }
    });
}

/* キー入力でゼロパディングON/OFF / Toggle zero-padding via keyboard (Z) */
/* 任意キーでチェックボックスをトグル / Generic checkbox toggle by key */
function addToggleKeyHandler(dlg, keyChar, checkbox, onChanged, skipWhenEditTextFocus) {
    if (!dlg || !checkbox || !dlg.addEventListener) return;
    dlg.addEventListener("keydown", function(event) {
        // Aキーなど、入力欄フォーカス中は無視したい場合に限りスキップ / Skip only when requested
        if (skipWhenEditTextFocus) {
            try {
                if (event.target && event.target.type === "edittext") return;
            } catch (_) {}
        }
        var k = (event.keyName || "").toUpperCase();
        if (k === String(keyChar).toUpperCase()) {
            checkbox.value = !checkbox.value;
            if (typeof onChanged === "function") {
                try {
                    onChanged();
                } catch (_) {}
            }
            event.preventDefault();
        }
    });
}




/* =============================
   UI ヘルパー / UI Helpers
   ============================= */
function addColumn(parent, align) {
    var col = parent.add("group");
    col.orientation = "column";
    col.alignChildren = align || "left";
    return col;
}

function addLabeledEditText(parent, label, initValue, width) {
    parent.add("statictext", undefined, label);
    var field = parent.add("edittext", undefined, initValue);
    field.characters = width;
    return field;
}

function addCheckbox(parent, label) {
    return parent.add("checkbox", undefined, label);
}

/* ライブプレビュー本体 / Live preview core (decoupled from UI) */
function updatePreview(doc, layerName, start, prefix, suffix, zeroPad, showTotal) {
    if (!doc) return;
    var layer = null;
    try {
        layer = doc.layers.getByName(layerName);
    } catch (e) {
        layer = null;
    }
    if (!layer) return;

    var frames = layer.textFrames;
    if (isNaN(start)) return;

    // まず選択中のテキストを優先（レイヤー無関係）/ Prefer the currently selected TextFrame (layer-agnostic)
    var selectedTemplate = null;
    try {
        selectedTemplate = getSelectedTextFrame();
    } catch (_sel) {}

    // 選択が無い場合は従来どおり _pagenumber の最初（AB順）/ otherwise fallback to first in _pagenumber by AB order
    var needToPickFromLayer = !selectedTemplate;

    // 常にプレビューでクリーン複製 / Always rebuild via preview duplication
    if (selectedTemplate || (frames && frames.length >= 1)) {
        var tmpLayer = getOrCreateLayer(doc, TMP_LAYER_NAME);
        trySet(tmpLayer, 'visible', false);
        trySet(tmpLayer, 'locked', false);

        // テンプレートを決定 / decide template
        var template = selectedTemplate;
        if (!template) {
            var pick = [];
            for (var pi = 0; pi < frames.length; pi++) {
                var pf = frames[pi];
                pick.push({
                    f: pf,
                    idx: getArtboardIndexByPosition(doc, pf.position)
                });
            }
            pick.sort(function(a, b) {
                return a.idx - b.idx;
            });
            template = pick.length ? pick[0].f : null;
        }
        if (!template) return;
        // バックアップ作成（非表示）/ make hidden backup
        var backup = null;
        tryCall(function() {
            backup = template.duplicate(tmpLayer, ElementPlacement.PLACEATBEGINNING);
        });
        trySet(backup, 'visible', false);
        trySet(backup, 'locked', true);

        // 一時アンロック＆可視化 / temporarily unlock/show
        trySet(template, 'locked', false);
        var srcLayer = null;
        try {
            srcLayer = template.layer;
        } catch (_) {
            srcLayer = null;
        }
        if (srcLayer) {
            trySet(srcLayer, 'locked', false);
            trySet(srcLayer, 'visible', true);
        }

        // テンプレの属するABをアクティブ化 / activate source AB
        try {
            var srcIdx = getArtboardIndexByPosition(doc, template.position);
            if (srcIdx >= 0) doc.artboards.setActiveArtboardIndex(srcIdx);
        } catch (_) {}

        // select → cut（クリップボードへ）/ select → cut (to clipboard)
        tryCall(function() {
            app.selection = null;
        });
        trySet(template, 'selected', true);
        tryCall(function() {
            app.cut();
        });

        // 既存のテキストを全削除（クリーンにしてから貼付）/ clear existing before paste
        try {
            removeOtherTextFrames(layer, null);
        } catch (_) {}

        // 貼付先レイヤーを明示 / ensure paste target layer
        try {
            doc.activeLayer = layer;
        } catch (_) {}

        // 全ABに貼付 / paste to all artboards
        tryCall(function() {
            app.executeMenuCommand('pasteInAllArtboard');
        });

        // 取得し直し / refresh refs
        frames = layer.textFrames;
    }

    // 計算準備 / Precompute
    var totalAB = doc.artboards.length;
    var maxNum = start + totalAB - 1;
    var maxDigits = String(maxNum).length;
    prefix = prefix || "";
    suffix = suffix || "";
    zeroPad = !!zeroPad;
    showTotal = !!showTotal;

    // AB順で並べ替え / Sort by artboard index
    var list = [];
    for (var i = 0; i < frames.length; i++) {
        var f = frames[i];
        var abIdx = getArtboardIndexByPosition(doc, f.position);
        list.push({
            frame: f,
            abIdx: abIdx
        });
    }
    list.sort(function(a, b) {
        return a.abIdx - b.abIdx;
    });

    for (var j = 0; j < list.length; j++) {
        var num = start + j;
        list[j].frame.contents = buildPageNumberString(num, maxDigits, prefix, suffix, zeroPad, maxNum, showTotal);
    }
    try {
        app.redraw();
    } catch (_e) {}
}

function main() {
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    var PG_STATE = null;
    var pgl = null;
    // _pagenumber レイヤーの事前作成＋状態保存＆一時解除・最前面化
    var docForPg = null;
    if (app.documents.length > 0) {
        try {
            docForPg = app.activeDocument;
            pgl = getOrCreatePagenumberLayer(docForPg);
            // 元状態をキャプチャ / capture original state
            PG_STATE = capturePagenumberState(docForPg, pgl);
            // 作業のためにロック解除＆可視化＆最前面化 / prepare for work
            preparePagenumberForWork(docForPg, pgl);
        } catch (_prep) {}
    }
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    // 親（3カラム）/ 3-column layout
    var columnsGroup = dialog.add("group");
    columnsGroup.orientation = "row";
    columnsGroup.alignChildren = "top";

    // 左: 接頭辞 / Left: Prefix
    var leftGroup = addColumn(columnsGroup);
    var prefixField = addLabeledEditText(leftGroup, LABELS.prefixLabel[lang], "", 10);

    // 中央: 開始番号 + ゼロ埋め / Center: Starting number + Zero pad
    var centerGroup = addColumn(columnsGroup);
    var inputField = addLabeledEditText(centerGroup, LABELS.promptMessage[lang], "1", 6);
    var zeroPadCheckbox = addCheckbox(centerGroup, LABELS.zeroPadLabel[lang]);
    // 矢印キーでの増減を有効化 / Enable arrow-key increment-decrement
    changeValueByArrowKey(inputField, function() {
        triggerPreview();
    });

    // 右: 接尾辞 + 総ページ / Right: Suffix + Total pages
    var rightGroup = addColumn(columnsGroup);
    var suffixField = addLabeledEditText(rightGroup, LABELS.suffixLabel[lang], "", 10);
    var totalPageCheckbox = addCheckbox(rightGroup, LABELS.totalPageLabel[lang]);

    // --- ライブプレビューラッパー / Live preview wrapper ---
    function triggerPreview() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;
        var start = parseInt(inputField.text, 10);
        if (isNaN(start)) return;
        updatePreview(
            doc,
            "_pagenumber",
            start,
            (prefixField.text || ""),
            (suffixField.text || ""),
            !!zeroPadCheckbox.value,
            !!totalPageCheckbox.value
        );
    }

    // 値変更でプレビュー更新 / Update preview on value changes
    if (prefixField && prefixField.addEventListener) {
        prefixField.addEventListener("changing", function() {
            triggerPreview();
        });
    }
    if (inputField && inputField.addEventListener) {
        inputField.addEventListener("changing", function() {
            triggerPreview();
        });
    }
    if (suffixField && suffixField.addEventListener) {
        suffixField.addEventListener("changing", function() {
            triggerPreview();
        });
    }
    if (zeroPadCheckbox) {
        zeroPadCheckbox.onClick = function() {
            triggerPreview();
        };
    }
    if (totalPageCheckbox) {
        totalPageCheckbox.onClick = function() {
            triggerPreview();
        };
    }
    // Zキーでゼロパディング、Aキーで総ページ表示を切替 / Toggle with keys
    addToggleKeyHandler(dialog, "Z", zeroPadCheckbox, function() {
        triggerPreview();
    }, false);
    addToggleKeyHandler(dialog, "A", totalPageCheckbox, function() {
        triggerPreview();
    }, true);

    // メイングループ（横並び） / Main group (horizontal layout)
    var btnRowGroup = dialog.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.alignChildren = ["fill", "center"];
    btnRowGroup.margins = [10, 10, 10, 0];
    btnRowGroup.alignment = ["fill", "bottom"];

    // 左側グループ / Left-side button group
    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    // キャンセルボタン / Cancel button
    var btnCancel = btnLeftGroup.add("button", undefined, LABELS.cancelLabel[lang], {
        name: "cancel"
    });

    // スペーサー（伸縮）/ Spacer (stretchable)
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    // 右側グループ / Right-side button group
    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    // OKボタン / OK button
    var btnOK = btnRightGroup.add("button", undefined, LABELS.okLabel[lang], {
        name: "ok"
    });

    // ハンドラ / Handlers
    btnOK.onClick = function() {
        // プレビュー用バックアップを破棄 / discard preview backup if any
        try {
            if (app.documents.length) {
                forceRemoveLayerByName(app.activeDocument, TMP_LAYER_NAME);
            }
        } catch (_) {}

        // 追加：_pagenumber の状態を復元 / restore _pagenumber layer state
        try {
            if (app.documents.length && PG_STATE) {
                var __doc = app.activeDocument;
                var __pgl = pgl || getOrCreatePagenumberLayer(__doc);
                restorePagenumberState(__doc, __pgl, PG_STATE);
            }
        } catch (_restOK) {}

        dialog.close(1);
    };
    btnCancel.onClick = function() {
        // （既存）プレビューの復元ロジック
        try {
            if (!app.documents.length) {
                dialog.close(0);
                return;
            }
            var doc = app.activeDocument;
            var tmp = null;
            try {
                tmp = doc.layers.getByName(TMP_LAYER_NAME);
            } catch (_nf) {
                tmp = null;
            }
            if (!tmp) {
                // バックアップが無い場合でも状態復元は行う
            } else {
                var pgl = getOrCreateLayer(doc, "_pagenumber");
                try {
                    removeOtherTextFrames(pgl, null);
                } catch (_r) {}
                try {
                    tmp.locked = false;
                } catch (_l) {}
                try {
                    tmp.visible = true;
                } catch (_v) {}
                for (var bi = tmp.pageItems.length - 1; bi >= 0; bi--) {
                    try {
                        var it = tmp.pageItems[bi];
                        try {
                            it.locked = false;
                        } catch (_il) {}
                        it.move(pgl, ElementPlacement.PLACEATBEGINNING);
                        try {
                            it.visible = true;
                        } catch (_iv) {}
                    } catch (_mv) {}
                }
                forceRemoveLayerByName(doc, TMP_LAYER_NAME);
                try {
                    app.redraw();
                } catch (_rd) {}
            }
        } catch (_e) {}

        // 追加：_pagenumber の状態を復元 / restore _pagenumber layer state
        try {
            if (app.documents.length && PG_STATE) {
                var __doc = app.activeDocument;
                var __pgl = pgl || getOrCreatePagenumberLayer(__doc);
                restorePagenumberState(__doc, __pgl, PG_STATE);
            }
        } catch (_restCancel) {}

        dialog.close(0);
    };

    inputField.active = true;

    // ダイアログの透明度と位置調整 / Adjust dialog opacity and position
    var offsetX = 300;
    var dialogOpacity = 0.97;

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function() {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        dlg.opacity = opacityValue;
    }

    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);

    // 初回プレビュー / First preview pass
    try {
        triggerPreview();
    } catch (_initPrev) {}

    if (dialog.show() !== 1) {
        return;
    }

    var startNum = parseInt(inputField.text, 10);
    if (isNaN(startNum)) {
        alert(LABELS.errorNotNumber[lang]);
        return;
    }

    if (app.documents.length === 0) return;
    var doc = app.activeDocument;

    pgl = pgl || getOrCreatePagenumberLayer(doc);

    var moved = moveSelectionToLayer(doc, pgl);
    removeOtherTextFrames(pgl, moved);

    // 選択を優先。なければ _pagenumber 内から探索 / Prefer selected text, otherwise pick from _pagenumber
    var targetText = getSelectedTextFrame();
    if (!targetText) {
        targetText = findTextFrameOnAnyArtboard(pgl, doc);
    }

    if (!targetText) {
        alert(LABELS.errorInvalidSelection[lang]);
        return;
    }

    // _pagenumber レイヤーに置く（確定時）/ Ensure it's on _pagenumber (for commit)
    try {
        if (targetText.layer.name !== '_pagenumber') targetText.move(pgl, ElementPlacement.PLACEATBEGINNING);
    } catch (_mv) {}

    // ここではアートボードの付け替えは行わない / Do not rebase to artboard 1 here
    removeOtherTextFrames(pgl, targetText);

    if (targetText.layer.name !== "_pagenumber") {
        alert(LABELS.errorInvalidSelection[lang]);
        return;
    }

    if (targetText.typename !== "TextFrame") {
        alert(LABELS.errorInvalidSelection[lang]);
        return;
    }

    seedAndPasteAll(doc, targetText, startNum);

    var framesList = collectTextFramesSorted(doc, pgl, targetText);
    var maxNum = startNum + doc.artboards.length - 1;
    var maxDigits = String(maxNum).length;
    applyNumbering(framesList, startNum, maxDigits, prefixField.text, suffixField.text, zeroPadCheckbox.value, maxNum, totalPageCheckbox.value);

}

main();