#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "DialogEngine"

/*

### スクリプト名：

FitArtboardWithMargin.jsx

### 概要

- 更新日：20260116
- 選択オブジェクトまたはすべてのオブジェクトのバウンディングボックスにマージンを加え、アートボードを自動調整します。
- 定規単位に応じた初期マージン値と即時プレビュー付きダイアログを提供します。
- ピクセル整数値に丸めてアートボードを設定します。

### 主な機能

- 定規単位ごとのマージン初期値設定
- 外接バウンディングボックス計算
- 即時プレビュー付きダイアログ
- ピクセル整数丸め

### 処理の流れ

1. 対象（選択オブジェクトまたはアートボード）を選択
2. ダイアログでマージン値を設定（即時プレビュー対応）
3. 設定に基づきアートボードを自動調整

### オリジナル、謝辞

Gorolib Design
https://gorolib.blog.jp/archives/71820861.html

### オリジナルからの変更点

- ダイアログボックスを閉じずにプレビュー更新
- 単位系（mm、px など）によってデフォルト値を切り替え
- アートボードの座標・サイズをピクセルベースで整数値に
- オブジェクトを選択していない場合には、すべてのオブジェクトを対象に
- ↑↓キー、shift + ↑↓キーによる入力

### note

https://note.com/dtp_transit/n/n15d3c6c5a1e5

### 更新履歴

- v1.0 (20250420) : 初期バージョン
- v1.1 (20250708) : UI改善、ポイント初期値変更
- v1.2 (20250709) : UI改善とバグ修正
- v1.3 (20250710) : 「対象：現在のアートボード、すべてのアートボード」を追加
- v1.4 (20250713) : 矢印キーによる値変更機能を追加、UI改善
- v1.5 (20250715) : 上下・左右個別に設定できるように
- v1.6 (20250716) : テキストを含む場合、実行時にアウトライン化して再計測
- v1.7 (20250717) : プレビュー境界チェックボックスを追加
- v1.7.1 (20250817): 微調整
- v1.8 (20260116) : プレビュー時のUndo履歴汚染を抑制、OK後に一括Undo可能な形に調整
- v1.8.1 (20260116) : マージンに負の値を入力・矢印キー操作で許容

---

### Script Name:

FitArtboardWithMargin.jsx

### Overview

- Last Updated: 20260116
- Automatically resize the artboard to fit the bounding box of selected or all objects with margin.
- Provides unit-based default margin values and an instant preview dialog.
- Sets the artboard size rounded to pixel integers.

### Main Features

- Default margin values based on ruler units
- Bounding box calculation
- Dialog with live preview
- Integer pixel rounding

### Workflow

1. Select target (selection or artboard)
2. Set margin value in dialog (with live preview)
3. Automatically adjust artboard size based on settings

### Changelog

- v1.0 (20250420): Initial version
- v1.1 (20250708): UI improvements, updated default point value
- v1.2 (20250709): UI improvements and bug fixes
- v1.3 (20250710): Added "Target: Current Artboard, All Artboards" options
- v1.4 (20250713): Added arrow key value change feature, UI improvements
- v1.5 (20250715): Enabled separate settings for vertical and horizontal margins
- v1.6 (20250716): Outlines text at runtime for accurate measurement
- v1.7 (20250817): Added preview bounds checkbox
- v1.7.1 (20250818): Minor adjustments
- v1.8 (20260116): Prevents preview from polluting Undo history; keeps final action as a single Undo step
- v1.8.1 (20260116): Allow negative margin values (including arrow-key changes)

*/

var SCRIPT_VERSION = "v1.8.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

// -------------------------------
// 日英ラベル定義 / Japanese-English label definitions
// -------------------------------


var LABELS = {
    dialogTitle: {
        ja: "アートボードサイズを調整 " + SCRIPT_VERSION,
        en: "Adjust Artboard Size " + SCRIPT_VERSION
    },
    targetSelection: {
        ja: "選択したオブジェクト",
        en: "Selected Objects"
    },
    targetArtboard: {
        ja: "現在のアートボード",
        en: "Current Artboard"
    },
    targetAllArtboards: {
        ja: "すべてのアートボード",
        en: "All Artboards"
    },
    marginLabel: {
        ja: "マージン",
        en: "Margin"
    },
    marginVertical: {
        ja: "上下",
        en: "Vertical"
    },
    marginHorizontal: {
        ja: "左右",
        en: "Horizontal"
    },
    linked: {
        ja: "連動",
        en: "Linked"
    },
    // 最下部プレビュー境界チェックボックス用 / For preview bounds checkbox
    previewBounds: {
        ja: "プレビュー境界",
        en: "Preview bounds"
    },
    numberAlert: {
        ja: "数値を入力してください。",
        en: "Please enter a number."
    },
    errorOccurred: {
        ja: "エラーが発生しました: ",
        en: "An error occurred: "
    }
};

/*
エラー整形ヘルパー / Error formatting helper
Illustrator/ExtendScript の Error から行番号などを含めて読みやすく整形。
*/
function formatError(e) {
    try {
        var msg = (e && e.message) ? String(e.message) : String(e);
        var ln = (e && e.line) ? (" line " + e.line) : "";
        var fn = (e && e.fileName) ? (" (" + e.fileName + ")") : "";
        return msg + ln + fn;
    } catch (_) {
        return String(e);
    }
}


/**
 * プレビュー時の履歴管理と一括Undoを制御するクラス / Preview Undo/History manager
 *
 * - updatePreview() のたびに rollback() してから addStep() で最新状態を1回だけ適用
 * - OK/Cancel 時に rollback() してプレビュー操作を履歴から取り除く
 */
function PreviewManager() {
    this.undoDepth = 0; // プレビュー中に実行されたアクションの回数

    /**
     * 変更操作を実行し、履歴としてカウントする
     * @param {Function} func - 実行したい処理（無名関数で渡す）
     */
    this.addStep = function(func) {
        try {
            func();
            this.undoDepth++;
            app.redraw();
        } catch (e) {
            // 失敗時はカウントしない
            try { $.writeln("[PreviewManager] addStep error: " + e); } catch (_) {}
        }
    };

    /**
     * プレビューのために行った変更を全て取り消す（キャンセル時など）
     */
    this.rollback = function() {
        try {
            while (this.undoDepth > 0) {
                app.undo();
                this.undoDepth--;
            }
        } catch (e) {
            // undo が効かない/失敗した場合はここで止める
            try { $.writeln("[PreviewManager] rollback error: " + e); } catch (_) {}
            this.undoDepth = 0;
        }
        try { app.redraw(); } catch (_) {}
    };

    /**
     * 現在の状態を確定する（OK時）
     * このスクリプトでは、OK時に一度 rollback() してから main() 側で本処理を1回だけ実行するため、
     * ここでは rollback のみを行う。
     */
    this.confirm = function() {
        this.rollback();
    };
}


/*
共通エンジン名を使用 / Use a common engine name
複数スクリプト間で位置記憶を共有しつつ、key で保存先を分離します。
Share session state across scripts; separate each dialog by key.
*/
function _getSavedLoc(key) {
    return $.global[key] && $.global[key].length === 2 ? $.global[key] : null;
}

function _setSavedLoc(key, loc) {
    $.global[key] = [loc[0], loc[1]];
}

function _clampToScreen(loc) {
    try {
        var vb = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
        var x = Math.max(vb[0] + 10, Math.min(loc[0], vb[2] - 10));
        var y = Math.max(vb[1] + 10, Math.min(loc[1], vb[3] - 10));
        return [x, y];
    } catch (e) {
        return loc;
    }
}



// -------------------------------
// 設定定数 / Configuration constants
// -------------------------------
var CONFIG = {
    supportedUnits: ['inch', 'mm', 'pt', 'pica', 'cm', 'H', 'px'],
    defaultMarginByUnit: {
        mm: '5',
        px: '20',
        pt: '10',
        _fallback: '0'
    },
    previewBoundsDefault: true,
    linkDefault: true,
    dialogOpacity: 0.95,
    offsetX: 300
};

function getDefaultMargin(unit) {
    return CONFIG.defaultMarginByUnit.hasOwnProperty(unit) ?
        CONFIG.defaultMarginByUnit[unit] :
        CONFIG.defaultMarginByUnit._fallback;
}

/*
マージンダイアログ表示 / Show margin input dialog with live preview
*/
function showMarginDialog(defaultValue, unit, artboardCount, hasSelection) {
    var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
    // ---- Dialog position persistence / 位置の記憶 ----
    var dlgPositionKey = "__FitArtboardWithMargin_Dialog"; // 固有キー / unique key for this dialog
    if ($.global[dlgPositionKey] === undefined) $.global[dlgPositionKey] = null; // ensure slot
    var __savedLoc = _getSavedLoc(dlgPositionKey);

    // apply saved location (fallback to existing centering/offset if none)
    if (__savedLoc) {
        dlg.onShow = (function(prev) {
            return function() {
                try {
                    if (typeof prev === 'function') prev();
                } catch (_) {}
                dlg.location = _clampToScreen(__savedLoc);
            };
        })(dlg.onShow);
    }

    // save on move
    var __saveDlgLoc = function() {
        _setSavedLoc(dlgPositionKey, [dlg.location[0], dlg.location[1]]);
    };
    dlg.onMove = (function(prev) {
        return function() {
            try {
                if (typeof prev === 'function') prev();
            } catch (_) {}
            __saveDlgLoc();
        };
    })(dlg.onMove);
    /* ダイアログ位置と不透明度のカスタマイズ / Customize dialog offset and opacity */

    function shiftDialogPosition(dlg, offsetX, offsetY) {
        dlg.onShow = function() {
            dlg.layout.layout(true);
            var dialogWidth = dlg.bounds.width;
            var dialogHeight = dlg.bounds.height;

            var screenWidth = $.screens[0].right - $.screens[0].left;
            var screenHeight = $.screens[0].bottom - $.screens[0].top;

            var centerX = screenWidth / 2 - dialogWidth / 2;
            var centerY = screenHeight / 2 - dialogHeight / 2;

            dlg.location = [centerX + offsetX, centerY + offsetY];
        };
    }

    function setDialogOpacity(dlg, opacityValue) {
        dlg.opacity = opacityValue;
    }

    setDialogOpacity(dlg, CONFIG.dialogOpacity);
    if (!__savedLoc) {
        // 初回のみセンターからのオフセットを適用 / Apply offset only on first run (no saved location)
        shiftDialogPosition(dlg, CONFIG.offsetX, 0);
    }
    /* ダイアログ位置と不透明度のカスタマイズ: ここまで / Dialog offset & opacity: end */

    dlg.orientation = "column";
    dlg.alignChildren = "fill";
    dlg.margins = 15;

    /* 対象選択パネル / Target selection panel */
    var targetPanel = dlg.add("panel", undefined, (lang === "ja" ? "対象" : "Target"));
    targetPanel.orientation = "column";
    targetPanel.alignChildren = "left";
    targetPanel.margins = [15, 20, 15, 10];

    var radioGroup = targetPanel.add("group");
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";

    var radioSelection = radioGroup.add("radiobutton", undefined, LABELS.targetSelection[lang]);
    radioSelection.enabled = hasSelection;
    var radioArtboard = radioGroup.add("radiobutton", undefined, LABELS.targetArtboard[lang]);
    var radioAllArtboards = radioGroup.add("radiobutton", undefined, LABELS.targetAllArtboards[lang]);

    /* 対象選択の初期値設定 / Set initial radio selection */
    if (!hasSelection && artboardCount === 1) {
        radioArtboard.value = true;
    } else if (!hasSelection && artboardCount > 1) {
        radioAllArtboards.value = true;
    } else {
        radioSelection.value = true;
    }

    /* マージン入力パネル / Margin input panel */
    var marginPanel = dlg.add("panel", undefined, (lang === "ja" ? "マージン" : "Margin") + " (" + unit + ")");
    marginPanel.orientation = "row";
    marginPanel.alignChildren = ["fill", "top"];
    marginPanel.margins = [15, 20, 15, 10];

    var leftColumn = marginPanel.add("group");
    leftColumn.orientation = "column";
    leftColumn.alignChildren = ["left", "center"];

    var rightColumn = marginPanel.add("group");
    rightColumn.orientation = "column";
    rightColumn.alignChildren = ["left", "center"];
    rightColumn.alignment = ["fill", "center"];

    /* 上下マージン入力欄 / Vertical margin input */
    var verticalGroup = leftColumn.add("group");
    verticalGroup.orientation = "row";
    var labelV = verticalGroup.add("statictext", undefined, LABELS.marginVertical[lang] + ":");
    var inputV = verticalGroup.add("edittext", undefined, defaultValue);
    inputV.characters = 4;

    /* 左右マージン入力欄 / Horizontal margin input */
    var horizontalGroup = leftColumn.add("group");
    horizontalGroup.orientation = "row";
    var labelH = horizontalGroup.add("statictext", undefined, LABELS.marginHorizontal[lang] + ":");
    var inputH = horizontalGroup.add("edittext", undefined, defaultValue);
    inputH.characters = 4;
    inputH.enabled = !CONFIG.linkDefault; // 連動ONなら左右はディム

    /* 連動チェックボックス / Linked checkbox */
    var linkCheckbox = rightColumn.add("checkbox", undefined, LABELS.linked[lang]);
    linkCheckbox.value = CONFIG.linkDefault;

    /* 現在のアートボードrectと全アートボードrectを保存（プレビュー用に復元） / Save current and all artboard rects for preview restore */
    var abIndex = app.activeDocument.artboards.getActiveArtboardIndex();
    var originalRects = [];
    for (var i = 0; i < app.activeDocument.artboards.length; i++) {
        originalRects.push(app.activeDocument.artboards[i].artboardRect.slice());
    }

    // プレビュー時のUndo履歴管理 / Preview undo/history manager
    var previewMgr = new PreviewManager();

    // --- Preview helpers / プレビュー用小関数 ---
    function previewAllArtboards(previewMarginV, previewMarginH) {
        for (var i = 0; i < app.activeDocument.artboards.length; i++) {
            var baseRect = originalRects[i].slice();
            var bounds = baseRect;
            bounds[0] -= previewMarginH;
            bounds[1] += previewMarginV;
            bounds[2] += previewMarginH;
            bounds[3] -= previewMarginV;
            app.activeDocument.artboards[i].artboardRect = bounds;
        }
        app.redraw();
    }

    function previewArtboard(previewMarginV, previewMarginH) {
        var baseRect = originalRects[abIndex].slice();
        var bounds = baseRect;
        bounds[0] -= previewMarginH;
        bounds[1] += previewMarginV;
        bounds[2] += previewMarginH;
        bounds[3] -= previewMarginV;
        app.activeDocument.artboards[abIndex].artboardRect = bounds;
        app.redraw();
    }

    function previewSelection(previewMarginV, previewMarginH) {
        var previewItems = app.activeDocument.selection.length === 0 ? app.activeDocument.pageItems : app.activeDocument.selection;
        var tempItems = collectEffectiveItems(previewItems);
        if (tempItems.length === 0) return;
        var previewBounds = getMaxBounds(tempItems, previewBoundsCheckbox.value);
        previewBounds[0] -= previewMarginH;
        previewBounds[1] += previewMarginV;
        previewBounds[2] += previewMarginH;
        previewBounds[3] -= previewMarginV;
        app.activeDocument.artboards[abIndex].artboardRect = previewBounds;
        app.redraw();
    }

    /*
    プレビュー更新関数 / Update artboard preview for dialog
    Undoによる履歴クリーンアップとプレビュー反映をPreviewManager経由で行う
    */
    function updatePreview(valueV, valueH) {
        // 直前のプレビューをUndoで巻き戻して履歴を汚さない
        previewMgr.rollback();

        var parsed = parseMarginPair(valueV, valueH, unit);
        if (!parsed.valid) return;
        var previewMarginV = parsed.vPt;
        var previewMarginH = parsed.hPt;

        var targetMode = radioSelection.value ? "selection" : (radioArtboard.value ? "artboard" : "allArtboards");

        previewMgr.addStep(function() {
            if (targetMode === "allArtboards") {
                previewAllArtboards(previewMarginV, previewMarginH);
                return;
            }
            if (targetMode === "artboard") {
                previewArtboard(previewMarginV, previewMarginH);
                return;
            }
            // default: selection
            previewSelection(previewMarginV, previewMarginH);
        });
    }

    /* 入力欄で矢印キーによる増減を可能に / Enable arrow key increment/decrement in input */
    changeValueByArrowKey(inputV, function(val) {
        if (linkCheckbox.value) inputH.text = val;
        updatePreview(inputV.text, inputH.text);
    });
    changeValueByArrowKey(inputH, function(val) {
        if (linkCheckbox.value) return;
        updatePreview(inputV.text, inputH.text);
    });
    inputV.active = true;
    inputV.onChanging = function() {
        if (linkCheckbox.value) inputH.text = inputV.text;
        updatePreview(inputV.text, inputH.text);
    };
    inputH.onChanging = function() {
        if (linkCheckbox.value) return; // 連動中は水平の直接編集は無効
        updatePreview(inputV.text, inputH.text);
    };
    linkCheckbox.onClick = function() {
        var linked = linkCheckbox.value;
        // 連動ONのときは左右をディム、値は上下に追随
        if (linked) {
            inputH.text = inputV.text;
            inputH.enabled = false;
        } else {
            inputH.enabled = true;
        }
        updatePreview(inputV.text, inputH.text);
    };
    radioSelection.onClick = function() {
        updatePreview(inputV.text, inputH.text);
    };
    radioArtboard.onClick = function() {
        updatePreview(inputV.text, inputH.text);
    };
    radioAllArtboards.onClick = function() {
        updatePreview(inputV.text, inputH.text);
    };

    // 最下部（ボタンの上）に配置するチェックボックス / Checkbox above buttons at the bottom
    var previewBoundsCheckbox = dlg.add("checkbox", undefined, LABELS.previewBounds[lang]);
    previewBoundsCheckbox.alignment = "left";
    previewBoundsCheckbox.value = CONFIG.previewBoundsDefault; // デフォルト
    previewBoundsCheckbox.margins = [0, 5, 0, 0];
    // チェック状態の変更でプレビューを更新 / Refresh preview when checkbox toggled
    previewBoundsCheckbox.onClick = function() {
        updatePreview(inputV.text, inputH.text);
    };

    /* ボタングループ / Button group */
    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    var cancelBtn = btnGroup.add("button", undefined, (lang === "ja" ? "キャンセル" : "Cancel"), {
        name: "cancel"
    });
    var okBtn = btnGroup.add("button", undefined, (lang === "ja" ? "OK" : "OK"), {
        name: "ok"
    });
    btnGroup.margins = [0, 5, 0, 0];

    var result = null;
    okBtn.onClick = function() {
        var parsed = parseMarginPair(inputV.text, inputH.text, unit);
        if (parsed && parsed.valid) {
            result = {
                marginV: inputV.text,
                marginH: inputH.text,
                target: radioSelection.value ? "selection" : (radioArtboard.value ? "artboard" : "allArtboards"),
                previewBounds: previewBoundsCheckbox.value
            };
            // プレビューによる履歴を一括Undoしてから閉じる
            previewMgr.confirm();
            dlg.close();
        } else {
            alert(LABELS.numberAlert[lang]);
        }
    };
    cancelBtn.onClick = function() {
        // キャンセル時は必ずロールバックして閉じる（プレビューによる履歴を残さない）
        previewMgr.rollback();
        dlg.close();
    };

    // persist location on close
    okBtn.onClick = (function(prev) {
        return function() {
            try {
                __saveDlgLoc();
            } catch (_) {}
            if (typeof prev === 'function') return prev();
            dlg.close(1);
        };
    })(okBtn.onClick);

    if (typeof cancelBtn !== 'undefined' && cancelBtn) {
        cancelBtn.onClick = (function(prev) {
            return function() {
                try {
                    __saveDlgLoc();
                } catch (_) {}
                if (typeof prev === 'function') return prev();
                dlg.close(0);
            };
        })(cancelBtn.onClick);
    }
    updatePreview(inputV.text, inputH.text);
    dlg.show();
    return result;
}

/* メイン処理 / Main process */
function main() {
    var selectedItems, artboards, rulerType, marginUnit;
    var defaultMarginValue, artboardIndex, selectedBounds;

    try {
        var doc = app.activeDocument;
        selectedItems = doc.selection;
        if (selectedItems.length === 0) {
            selectedItems = doc.pageItems;
            if (selectedItems.length === 0) return;
        }

        artboards = doc.artboards;
        rulerType = app.preferences.getIntegerPreference("rulerType");
        var su = CONFIG.supportedUnits;
        marginUnit = su[rulerType];

        /* 単位ごとの初期マージン値設定 / Set default margin value based on unit */
        defaultMarginValue = getDefaultMargin(marginUnit);

        /* 選択なし・複数アートボード時は allArtboards をデフォルトに / Default to allArtboards if no selection and multiple artboards */
        var isAllArtboardsDefault = (doc.selection.length === 0 && artboards.length > 1);
        if (isAllArtboardsDefault) {
            defaultMarginValue = '0';
        }

        /* ユーザーにマージンを入力させる / Show margin input dialog */
        var userInput = showMarginDialog(defaultMarginValue, marginUnit, artboards.length, doc.selection.length > 0);
        if (!userInput) return;

        /* allArtboards選択時は計算用マージンを0に / Set margin to 0 for allArtboards calculation */
        if (userInput.target === "allArtboards") {
            defaultMarginValue = '0';
        }

        // 新しいUI: userInput.marginV, userInput.marginH
        var marginV = parseFloat(userInput.marginV);
        var marginH = parseFloat(userInput.marginH);
        var targetMode = userInput.target;
        var marginVInPoints = toPt(marginV, marginUnit);
        var marginHInPoints = toPt(marginH, marginUnit);

        if (targetMode === "artboard") {
            var artRect = artboards[artboards.getActiveArtboardIndex()].artboardRect;
            var bounds = artRect.slice();
            bounds[0] -= marginHInPoints;
            bounds[1] += marginVInPoints;
            bounds[2] += marginHInPoints;
            bounds[3] -= marginVInPoints;

            var x0 = Math.round(bounds[0]);
            var y1 = Math.round(bounds[1]);
            var x2 = Math.round(bounds[2]);
            var y3 = Math.round(bounds[3]);

            var width = Math.round(x2 - x0);
            var height = Math.round(y1 - y3);

            x2 = x0 + width;
            y3 = y1 - height;

            bounds[0] = x0;
            bounds[1] = y1;
            bounds[2] = x2;
            bounds[3] = y3;

            artboards[artboards.getActiveArtboardIndex()].artboardRect = bounds;
            return;
        }

        if (targetMode === "allArtboards") {
            for (var i = 0; i < artboards.length; i++) {
                var artRect = artboards[i].artboardRect;
                var bounds = artRect.slice();
                bounds[0] -= marginHInPoints;
                bounds[1] += marginVInPoints;
                bounds[2] += marginHInPoints;
                bounds[3] -= marginVInPoints;

                var x0 = Math.round(bounds[0]);
                var y1 = Math.round(bounds[1]);
                var x2 = Math.round(bounds[2]);
                var y3 = Math.round(bounds[3]);

                var width = Math.round(x2 - x0);
                var height = Math.round(y1 - y3);

                x2 = x0 + width;
                y3 = y1 - height;

                bounds[0] = x0;
                bounds[1] = y1;
                bounds[2] = x2;
                bounds[3] = y3;

                artboards[i].artboardRect = bounds;
            }
            return;
        }

        if (targetMode === "selection") {

            // 選択アイテムの正規化（クリップグループ→クリッピングパス）/ Normalize selection (clipped group → clipping path)
            selectedItems = collectEffectiveItems(selectedItems);
            if (selectedItems.length === 0) return;

            // TextFrame のアウトライン化結果とその他アイテムを明確に分離
            var originalTextFrames = [];
            var outlinedTextFrames = [];
            var otherItems = [];

            for (var i = 0; i < selectedItems.length; i++) {
                var item = selectedItems[i];
                if (item.typename === "TextFrame") {
                    var dup = item.duplicate();
                    dup.hidden = true;
                    originalTextFrames.push(dup);

                    var outlined = item.createOutline();
                    // createOutline() が複数パス返す場合に対応
                    if (outlined.length && outlined.length > 0) {
                        for (var j = 0; j < outlined.length; j++) {
                            outlinedTextFrames.push(outlined[j]);
                        }
                    } else {
                        outlinedTextFrames.push(outlined);
                    }
                } else {
                    otherItems.push(item);
                }
            }

            // アウトライン化したTextFrameとその他アイテムを合成してバウンディングボックスを取得
            var tempItemsForBounds = outlinedTextFrames.concat(otherItems);
            selectedBounds = getMaxBounds(tempItemsForBounds, userInput.previewBounds);
            selectedBounds[0] -= marginHInPoints;
            selectedBounds[1] += marginVInPoints;
            selectedBounds[2] += marginHInPoints;
            selectedBounds[3] -= marginVInPoints;

            // outlinedTextFrames のみ remove()
            for (var i = 0; i < outlinedTextFrames.length; i++) {
                try {
                    outlinedTextFrames[i].remove();
                } catch (e) {}
            }
            // originalTextFrames のみ hidden=false で復元
            for (var i = 0; i < originalTextFrames.length; i++) {
                try {
                    originalTextFrames[i].hidden = false;
                } catch (e) {}
            }

            // 座標と幅・高さを整数に丸める
            var x0 = Math.round(selectedBounds[0]);
            var y1 = Math.round(selectedBounds[1]);
            var x2 = Math.round(selectedBounds[2]);
            var y3 = Math.round(selectedBounds[3]);

            var width = Math.round(x2 - x0);
            var height = Math.round(y1 - y3);

            x2 = x0 + width;
            y3 = y1 - height;

            selectedBounds[0] = x0;
            selectedBounds[1] = y1;
            selectedBounds[2] = x2;
            selectedBounds[3] = y3;

            // アートボードの更新
            artboardIndex = artboards.getActiveArtboardIndex();
            artboards[artboardIndex].artboardRect = selectedBounds;
        }

    } catch (e) {
        try {
            $.writeln("[FitArtboardWithMargin] ERROR: " + formatError(e));
        } catch (_) {}
        alert(LABELS.errorOccurred[lang] + formatError(e));
    }
}

/*
選択アイテムの正規化 / Normalize selected items
- クリップグループ(GroupItem.clipped=true)はクリッピングパスのみを採用
- それ以外はそのまま
This unifies item collection for preview/runtime to avoid duplication.
*/
function collectEffectiveItems(items) {
    var out = [];
    for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (it && it.typename === "GroupItem" && it.clipped) {
            // クリップグループはクリッピングパスのみ採用 / For clipped group, include only the clipping path
            try {
                for (var j = 0; j < it.pageItems.length; j++) {
                    var child = it.pageItems[j];
                    if (child.clipping) {
                        out.push(child);
                        break;
                    }
                }
            } catch (e) {
                /* ignore */ }
        } else {
            out.push(it);
        }
    }
    return out;
}

/*
単位→pt変換ユーティリティ / Unit to pt conversion utility
`new UnitValue(val, unit).as('pt')` の共通化 / Factor common conversion
*/
function toPt(val, unit) {
    try {
        var n = Number(val);
        if (isNaN(n)) return NaN;
        return new UnitValue(n, unit).as('pt');
    } catch (e) {
        return NaN;
    }
}

/*
入力検証の一元化 / Centralize input validation for margins
- `h` が未入力/NaN の場合は `v` を採用 / If `h` is NaN, fall back to `v`
- 単位変換まで実施し pt 値も返す
*/
function parseMarginPair(textV, textH, unit) {
    var v = parseFloat(textV);
    var h = parseFloat(textH);
    if (isNaN(v)) return {
        valid: false
    };
    if (isNaN(h)) h = v; // 連動相当のフォールバック / fallback like linked
    var vPt = toPt(v, unit);
    var hPt = toPt(h, unit);
    if (isNaN(vPt) || isNaN(hPt)) return {
        valid: false
    };
    return {
        valid: true,
        v: v,
        h: h,
        vPt: vPt,
        hPt: hPt
    };
}

/* 選択オブジェクト群から最大のバウンディングボックスを取得 / Get maximum bounding box from multiple items */
function getMaxBounds(items, usePreviewBounds) {
    var bounds = getBounds(items[0], usePreviewBounds);
    for (var i = 1; i < items.length; i++) {
        var itemBounds = getBounds(items[i], usePreviewBounds);
        bounds[0] = Math.min(bounds[0], itemBounds[0]);
        bounds[1] = Math.max(bounds[1], itemBounds[1]);
        bounds[2] = Math.max(bounds[2], itemBounds[2]);
        bounds[3] = Math.min(bounds[3], itemBounds[3]);
    }
    return bounds;
}

/* オブジェクトのバウンディングボックスを取得 / Get bounding box of a single object
   usePreviewBounds=true なら visibleBounds（プレビュー境界: 塗り/線を含む）
   usePreviewBounds=false なら geometricBounds（幾何境界: パス外形のみ） */
function getBounds(item, usePreviewBounds) {
    return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
}

/* edittextに矢印キーで値を増減する機能を追加 / Add arrow key increment/decrement to edittext */
function changeValueByArrowKey(editText, onUpdate) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;

        if (event.keyName == "Up" || event.keyName == "Down") {
            if (keyboard.shiftKey) {
                // 10の倍数にスナップ
                var base = Math.round(value / 10) * 10;
                if (event.keyName == "Up") {
                    value = base + 10;
                } else {
                    value = base - 10;
                }
            } else {
                var delta = event.keyName == "Up" ? 1 : -1;
                value += delta;
            }

            event.preventDefault();
            editText.text = value;
            if (typeof onUpdate === "function") {
                onUpdate(editText.text);
            }
        }
    });
}

main();
app.selectTool("Adobe Select Tool");