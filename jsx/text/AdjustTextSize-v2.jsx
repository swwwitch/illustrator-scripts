#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

AdjustTextScaleBaseline.jsx

### Readme （GitHub）：

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AdjustTextScaleBaseline.md

### 概要：

- 更新日：20260203
- Illustrator上で選択文字のフォントサイズ、拡大率、ベースラインシフト、カーニング、トラッキングを調整
- ダイアログボックスで数値入力し、即時プレビューと適用が可能

### 主な機能：

- 見かけのサイズとフォントサイズの変換
- 水平・垂直比率を個別調整可能
- ベースラインシフト、カーニング、トラッキングの指定
- OKで確定、リセットで初期状態に復元

### 更新履歴：

- v1.0 (20250723) : 初期バージョン
- v1.1 (20250724) : ベースラインシフトとカーニング対応
- v1.2 (20250725) : プレビュー反映の強化、100%時の見かけディム処理、shiftキー刻み修正
- v1.3 (20250726) : トラッキング機能追加、UI微調整
- v1.4 (20260203) : 更新履歴の日付整合、ベースライン適用条件の修正、軽微なクリーンアップ
*/

/*

### Script Name:

AdjustTextScaleBaseline.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AdjustTextScaleBaseline.md

### Description:

- Last Updated: 2026-02-03
- Adjust font size, horizontal/vertical scale, baseline shift, kerning, and tracking of selected text in Illustrator
- Realtime preview and numeric input via dialog box

### Main Features:

- Convert between apparent size and font size
- Individually adjust horizontal/vertical scale
- Modify baseline shift, kerning, and tracking
- Confirm with OK, restore with Reset

### Processing Flow:

1. Build UI and define labels (ja/en)
2. Get initial state of selected text
3. Apply changes in realtime on input
4. Confirm with OK, cancel or reset to revert

### Change Log:

- v1.0 (2025-07-23): Initial release
- v1.1 (2025-07-24): Added baseline shift and kerning
- v1.2 (2025-07-25): Enhanced preview handling, dim apparent size at 100% scale, fixed shift key increments
- v1.3 (2025-07-26): Added tracking feature, UI adjustments
- v1.4 (2026-02-03): Normalized changelog dates, fixed baseline target condition, minor cleanup

*/

var SCRIPT_VERSION = "v1.4";

// UI要素をまとめるオブジェクト / UI element references
var uiElements = {
    sizeInput: null,
    hScaleInput: null,
    baselineInput: null,
    kerningInput: null,
    apparentSizeText: null
};

var apparentSizeDisplay; // グローバルに定義 / Global definition

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: "フォントサイズとベースライン調整 " + SCRIPT_VERSION,
        en: "Font Size & Baseline Adjuster " + SCRIPT_VERSION
    },
    targetChar: {
        ja: "対象文字：",
        en: "Target Char:"
    },
    adjust: {
        ja: "調整",
        en: "Adjust"
    },
    fontSize: {
        ja: "フォントサイズ",
        en: "Font Size"
    },
    scale: {
        ja: "水平比率/垂直比率",
        en: "Scale"
    },
    apparent: {
        ja: "見かけ",
        en: "Apparent"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    reset: {
        ja: "リセット",
        en: "Reset"
    },
    selectTextAlert: {
        ja: "テキストを選択してください",
        en: "Please select text"
    },
    baselineShiftLabel: {
        ja: "ベースラインシフト",
        en: "Baseline Shift"
    },
    kerning: {
        ja: "カーニング",
        en: "Kerning"
    },
    tracking: {
        ja: "トラッキング",
        en: "Tracking"
    }
};

/* 文字選択状態を取得 / Get selected characters from current document */
function isTargetChar(ch, targetChar, hasTargetChar) {
    return !hasTargetChar || targetChar.indexOf(ch.contents) !== -1;
}

/* 文字選択状態を取得 / Get selected characters from current document */
function getTextSelection() {
    var sel = app.selection;
    var res = [];
    if (!sel || sel.length === 0) return res;
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item instanceof TextFrame) {
            res.push(item.textRange);
        } else if (item instanceof TextRange) {
            res.push(item);
        }
    }
    if (!(sel instanceof Array) && sel instanceof TextRange) {
        res.push(sel);
    }
    return res;
}

/* 英数字を除いたユニークな文字列を取得 / Get unique non-alphanumeric characters from text */
function getUniqueNonAlphanumerics(text) {
    var stripped = text.replace(/[0-9A-Za-z]/g, "");
    var result = "";
    for (var i = 0; i < stripped.length; i++) {
        var ch = stripped.charAt(i);
        if (result.indexOf(ch) === -1) {
            result += ch;
        }
    }
    return result;
}

/* 対象文字に一致する文字だけに処理を適用 / Apply callback to matching characters in text range */
function applyToTargetCharacters(range, targetChar, hasTargetChar, callback) {
    var chars = range.characters;
    for (var j = 0; j < chars.length; j++) {
        var ch = chars[j];
        if (isTargetChar(ch, targetChar, hasTargetChar)) {
            callback(ch);
        }
    }
}

/* 指定文字にだけベースラインシフトを適用 / Apply baseline shift only to matching characters */
function applyBaselineShiftToChars(range, targetChar, hasTargetChar) {
    if (!range || !range.characters || targetChar === undefined) return;

    var val = parseFloat(uiElements.baselineInput.text);
    if (isNaN(val)) return;

    var chars = range.characters;
    for (var i = 0; i < chars.length; i++) {
        var ch = chars[i];
        if (isTargetChar(ch, targetChar, hasTargetChar)) {
            ch.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
            ch.characterAttributes.baselineShift = val;
        }
    }
}

/* ダイアログの表示位置をオフセット / Shift dialog position by offsetX and offsetY */
function shiftDialogPosition(dlg, offsetX, offsetY) {
    dlg.onShow = function() {
        var currentX = dlg.location[0];
        var currentY = dlg.location[1];
        dlg.location = [currentX + offsetX, currentY + offsetY];
    };
}

/* ダイアログの透明度を設定 / Set dialog opacity */
function setDialogOpacity(dlg, opacityValue) {
    dlg.opacity = opacityValue;
}

/* 上下キーによる数値の増減を有効にする関数 / Enable value increment/decrement by arrow keys */
function changeValueByArrowKey(editText, options) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;
        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;
        var precision = 0;

        if (keyboard.shiftKey) {
            delta = options.shiftStep || 10;
        } else if (keyboard.altKey) {
            delta = options.altStep || 0.1;
            precision = 1;
        } else {
            delta = options.step || 1;
        }

        if (event.keyName == "Up") {
            if (keyboard.shiftKey) {
                value = Math.floor(value / delta) * delta + delta;
            } else {
                value += delta;
            }
        } else if (event.keyName == "Down") {
            if (keyboard.shiftKey) {
                value = Math.ceil(value / delta) * delta - delta;
            } else {
                value -= delta;
            }
        } else {
            return;
        }

        if (precision > 0) {
            var factor = Math.pow(10, precision);
            value = Math.round(value * factor) / factor;
        } else {
            value = Math.round(value);
        }

        event.preventDefault();
        editText.text = value;
        editText.notify("onChange");
    });
}

/* 単位コードとプリファレンスキーに応じて単位ラベルを返す関数 / Get unit label for code and pref key */
function getUnitLabel(code, prefKey) {
    var unitMap = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };
    if (code === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return unitMap[code] || "不明";
}

/* 現在のテキスト単位ラベルを取得 / Get current text unit label */
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("text/units");
    return getUnitLabel(unitCode, "text/units");
}

/* =========================================
 * PreviewManager
 * プレビュー時にUndo履歴を汚さないための小さな管理クラス。
 * - addStep(): 変更処理を実行してundoDepthをカウント
 * - rollback(): プレビューで行った変更をすべて取り消し
 * - confirm(finalAction): OK時に「一度戻して1回だけ本番処理」を実行し、Undoを1回にまとめる
 * ========================================= */
function PreviewManager() {
    this.undoDepth = 0;

    this.addStep = function(func) {
        try {
            func();
            this.undoDepth++;
            app.redraw();
        } catch (e) {
            alert("Preview Error: " + e);
        }
    };

    this.rollback = function() {
        while (this.undoDepth > 0) {
            app.undo();
            this.undoDepth--;
        }
        app.redraw();
    };

    this.confirm = function(finalAction) {
        if (typeof finalAction === "function") {
            this.rollback();
            finalAction();
        } else {
            this.undoDepth = 0;
        }
    };
}

/* メイン処理 / Main function */
function main() {
    // 初期フォントサイズを記録する変数
    var initialFontSize = null;

    if (app.documents.length <= 0) {
        return;
    }

    var targetRanges = getTextSelection();
    if (targetRanges.length === 0) {
        alert(LABELS.selectTextAlert[lang]);
        return;
    }

    var allSelText = "";
    for (var i = 0; i < targetRanges.length; i++) {
        allSelText += targetRanges[i].contents;
    }

    function getCommonCharacters(textArray) {
        if (textArray.length === 0) return "";
        var common = textArray[0];
        for (var i = 1; i < textArray.length; i++) {
            var current = textArray[i];
            var temp = "";
            for (var j = 0; j < common.length; j++) {
                var ch = common.charAt(j);
                if (current.indexOf(ch) !== -1 && temp.indexOf(ch) === -1) {
                    temp += ch;
                }
            }
            common = temp;
            if (common.length === 0) break;
        }
        return common;
    }

    var textArray = [];
    for (var i = 0; i < targetRanges.length; i++) {
        textArray.push(targetRanges[i].contents);
    }
    var commonChars = getCommonCharacters(textArray);
    var uniqueNonAN = getUniqueNonAlphanumerics(commonChars);

    // プレビューUndo管理 / Preview undo management
    var previewMgr = new PreviewManager();

    /* 入力値取得 / Get current state from inputs */
    function getTargetCharState() {
        var targetChar = targetCharInput.text;
        var hasTargetChar = targetChar && targetChar.length > 0;
        return {
            targetChar: targetChar,
            hasTargetChar: hasTargetChar
        };
    }

    function getOriginalCharSize(targetChar, hasTargetChar) {
        for (var i = 0, len = targetRanges.length; i < len; i++) {
            var chars = targetRanges[i].characters;
            for (var j = 0; j < chars.length; j++) {
                var ch = chars[j];
                if (isTargetChar(ch, targetChar, hasTargetChar)) {
                    return ch.size;
                }
            }
        }
        return null;
    }

    function findFirstMatchingChar(targetChar, hasTargetChar) {
        for (var i = 0, len = targetRanges.length; i < len; i++) {
            var chars = targetRanges[i].characters;
            for (var j = 0; j < chars.length; j++) {
                var ch = chars[j];
                if (isTargetChar(ch, targetChar, hasTargetChar)) {
                    return ch;
                }
            }
        }
        return null;
    }

    function updateInfoText() {
        var state = getTargetCharState();
        var exampleChar = findFirstMatchingChar(state.targetChar, state.hasTargetChar);
        if (!exampleChar) {
            uiElements.sizeInput.text = "";
            uiElements.hScaleInput.text = "";
            uiElements.apparentSizeText.text = "--";
            return;
        }
        var size = exampleChar.size;
        if (initialFontSize === null) {
            initialFontSize = size;
        }
        var h = exampleChar.characterAttributes.horizontalScale;
        size = Math.round(size * 10) / 10;
        h = Math.round(h * 10) / 10;
        uiElements.sizeInput.text = size + "";
        uiElements.hScaleInput.text = h + "";
        updateApparentSizeDisplay(unitLabel);
    }

    function calculateApparentSize(size, scale) {
        return Math.round(size * scale) / 100;
    }

    function updateApparentSizeDisplay(unitLabel) {
        var size = parseFloat(uiElements.sizeInput.text, 10);
        var scale = parseFloat(uiElements.hScaleInput.text, 10);
        if (isNaN(size) || isNaN(scale)) {
            uiElements.apparentSizeText.text = "--";
        } else {
            var apparent = calculateApparentSize(size, scale);
            uiElements.apparentSizeText.text = apparent + "";
        }

        var isDimmed = (scale === 100);
        apparentLabel.enabled = !isDimmed;
        uiElements.apparentSizeText.enabled = !isDimmed;
        displayGroup.children[2].enabled = !isDimmed;
    }

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.alignChildren = "left";

    var offsetX = 300;
    var dialogOpacity = 0.97;
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);

    var contentGroup = dialog.add("group");
    contentGroup.orientation = "row";
    contentGroup.alignChildren = "top";
    contentGroup.spacing = 20;

    var leftCol = contentGroup.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = "left";
    leftCol.spacing = 20;

    var rightCol = contentGroup.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = "right";

    var targetTextGroup = leftCol.add("group");
    targetTextGroup.orientation = "row";
    targetTextGroup.add("statictext", undefined, LABELS.targetChar[lang]);
    var targetCharInput = targetTextGroup.add("edittext", undefined, uniqueNonAN);
    targetCharInput.characters = 10;

    var infoPanel = leftCol.add("panel", undefined, LABELS.adjust[lang]);
    infoPanel.orientation = "column";
    infoPanel.alignChildren = ["left", "top"];
    infoPanel.margins = [15, 10, 15, 10];

    var modeGroup = infoPanel.add("group");
    modeGroup.orientation = "column";
    modeGroup.alignChildren = "left";

    var radioGroup = modeGroup.add("group");
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";

    var sizeGroup = radioGroup.add("group", undefined, { orientation: "row" });
    sizeGroup.margins = [0, 10, 0, 0];
    sizeGroup.alignChildren = "left";
    var sizeLabel = sizeGroup.add("statictext", undefined, LABELS.fontSize[lang]);
    sizeLabel.justify = "right";
    uiElements.sizeInput = sizeGroup.add("edittext", undefined, "0");
    uiElements.sizeInput.characters = 4;
    uiElements.sizeInput.enabled = true;

    var unitLabel = getUnitLabel(app.preferences.getIntegerPreference("text/units"), "text/units");
    sizeGroup.add("statictext", undefined, unitLabel);

    var scaleGroup = radioGroup.add("group", undefined, { orientation: "row" });
    scaleGroup.margins = [0, 0, 0, 6];
    var scaleLabel = scaleGroup.add("statictext", undefined, LABELS.scale[lang]);
    scaleLabel.justify = "right";
    uiElements.hScaleInput = scaleGroup.add("edittext", undefined, "100");
    uiElements.hScaleInput.characters = 4;
    scaleGroup.add("statictext", undefined, "%");
    uiElements.sizeInput.enabled = true;
    uiElements.hScaleInput.enabled = true;

    var displayGroup = infoPanel.add("group", undefined, { orientation: "row" });
    var apparentLabel = displayGroup.add("statictext", undefined, LABELS.apparent[lang]);
    apparentLabel.justify = "right";
    uiElements.apparentSizeText = displayGroup.add("statictext", undefined, "--");
    uiElements.apparentSizeText.characters = 5;
    displayGroup.add("statictext", undefined, unitLabel);

    var baselineGroup = infoPanel.add("group");
    baselineGroup.orientation = "row";
    baselineGroup.alignChildren = ["right", "center"];

    var baselineLabel = baselineGroup.add("statictext", undefined, LABELS.baselineShiftLabel[lang]);
    baselineLabel.justify = "right";

    function setUnifiedLabelWidth() {
        var labelWidth = 120;
        sizeLabel.preferredSize.width = labelWidth;
        scaleLabel.preferredSize.width = labelWidth;
        apparentLabel.preferredSize.width = labelWidth;
        baselineLabel.preferredSize.width = labelWidth;
        kerningLabel.preferredSize.width = labelWidth;
    }

    uiElements.baselineInput = baselineGroup.add("edittext", undefined, "0");
    uiElements.baselineInput.characters = 4;
    uiElements.baselineInput.justify = "right";
    baselineGroup.add("statictext", undefined, unitLabel);

    /* 共通適用ロジック: サイズまたはスケールを対象文字に適用 / Apply size or scale to target characters */
    function applyPropertyToCharacters(propertyName, value, state) {
        for (var i = 0, len = targetRanges.length; i < len; i++) {
            applyToTargetCharacters(targetRanges[i], state.targetChar, state.hasTargetChar, function(ch) {
                if (propertyName === "size") {
                    ch.size = value;
                } else if (propertyName === "scale") {
                    ch.characterAttributes.horizontalScale = value;
                    ch.characterAttributes.verticalScale = value;
                }
            });
        }
    }

    /* 現在UIの値をまとめて取得 / Collect current UI values */
    function getCurrentParams() {
        var sizeVal = parseFloat(uiElements.sizeInput.text);
        var scaleVal = parseFloat(uiElements.hScaleInput.text);
        var baselineVal = parseFloat(uiElements.baselineInput.text);
        var kerningVal = parseFloat(uiElements.kerningInput.text);
        var trackingVal = parseFloat(trackingInput.text);
        return {
            size: isNaN(sizeVal) ? null : sizeVal,
            scale: isNaN(scaleVal) ? null : scaleVal,
            baseline: isNaN(baselineVal) ? null : baselineVal,
            kerning: isNaN(kerningVal) ? null : kerningVal,
            tracking: isNaN(trackingVal) ? null : trackingVal
        };
    }

    /* 現在のUI状態を対象文字にまとめて適用 / Apply current UI state to target characters */
    function applyAllCurrentValues() {
        var state = getTargetCharState();
        var params = getCurrentParams();

        if (params.size !== null) applyPropertyToCharacters("size", params.size, state);
        if (params.scale !== null) applyPropertyToCharacters("scale", params.scale, state);

        if (params.baseline !== null) {
            for (var i = 0; i < targetRanges.length; i++) {
                applyBaselineShiftToChars(targetRanges[i], state.targetChar, state.hasTargetChar);
            }
        }
        if (params.kerning !== null) {
            for (var k = 0; k < targetRanges.length; k++) {
                applyKerningToChars(targetRanges[k], state.targetChar, state.hasTargetChar);
            }
        }
        if (params.tracking !== null) {
            for (var t = 0; t < targetRanges.length; t++) {
                applyToTargetCharacters(targetRanges[t], state.targetChar, state.hasTargetChar, function(ch) {
                    ch.characterAttributes.tracking = params.tracking;
                });
            }
        }
    }

    /* プレビュー更新 / Update preview without polluting undo history */
    function updatePreview() {
        previewMgr.rollback();
        previewMgr.addStep(function() {
            applyAllCurrentValues();
        });
    }

    /* onChanging の連打で Undo/再適用が過剰にならないよう間引く / Throttle heavy preview updates */
    var __lastPreviewTime = 0;
    var __previewIntervalMs = 80;
    function updatePreviewThrottled() {
        var now = (new Date()).getTime();
        if (now - __lastPreviewTime < __previewIntervalMs) return;
        __lastPreviewTime = now;
        updatePreview();
    }

    function applyBaselineShiftToTargetChar() {
        updatePreview();
    }

    uiElements.baselineInput.onChange = applyBaselineShiftToTargetChar;
    uiElements.baselineInput.onChanging = function() {
        updatePreviewThrottled();
    };
    changeValueByArrowKey(uiElements.baselineInput, { step: 1, shiftStep: 10, altStep: 0.1 });

    function addKeyEvents(input) {
        changeValueByArrowKey(input, { step: 1, shiftStep: 10, altStep: 0.1 });
    }
    addKeyEvents(uiElements.baselineInput);

    var kerningGroup = infoPanel.add("group");
    kerningGroup.orientation = "row";
    kerningGroup.alignChildren = ["right", "center"];

    var kerningLabel = kerningGroup.add("statictext", undefined, LABELS.kerning[lang]);
    kerningLabel.justify = "right";

    uiElements.kerningInput = kerningGroup.add("edittext", undefined, "0");
    uiElements.kerningInput.characters = 4;
    uiElements.kerningInput.justify = "right";
    kerningGroup.add("statictext", undefined, "/1000");

    changeValueByArrowKey(uiElements.kerningInput, { step: 1, shiftStep: 10, altStep: 0.1 });

    uiElements.kerningInput.onChange = function() {
        updatePreview();
    };
    uiElements.kerningInput.onChanging = function() {
        updatePreviewThrottled();
    };

    var trackingGroup = infoPanel.add("group");
    trackingGroup.orientation = "row";
    trackingGroup.alignChildren = ["right", "center"];

    var trackingLabel = trackingGroup.add("statictext", undefined, LABELS.tracking[lang]);
    trackingLabel.justify = "right";
    trackingLabel.preferredSize.width = 120;

    var trackingInput = trackingGroup.add("edittext", undefined, "0");
    trackingInput.characters = 4;
    trackingInput.justify = "right";
    trackingGroup.add("statictext", undefined, "/1000");

    changeValueByArrowKey(trackingInput, { step: 1, shiftStep: 10, altStep: 0.1 });

    trackingInput.onChange = function() {
        updatePreview();
    };
    trackingInput.onChanging = function() {
        updatePreviewThrottled();
    };

    var buttonGroup = rightCol.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.orientation = "column";
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang]);
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    var cancelResetSpacer = buttonGroup.add("statictext", undefined, "");
    cancelResetSpacer.preferredSize.height = 50;
    var resetBtn = buttonGroup.add("button", undefined, LABELS.reset[lang]);
    resetBtn.preferredSize.width = 90;
    cancelBtn.preferredSize.width = 90;
    okBtn.preferredSize.width = 90;

    function resetToInitial() {
        targetCharInput.text = uniqueNonAN;
        if (initialFontSize !== null) {
            uiElements.sizeInput.text = initialFontSize;
        }
        uiElements.hScaleInput.text = "100";
        uiElements.baselineInput.text = "0";
        uiElements.kerningInput.text = "0";
        trackingInput.text = "0";

        updatePreview();
        updateInfoText();
    }
    resetBtn.onClick = function() {
        resetToInitial();
    };

    cancelBtn.onClick = function() {
        previewMgr.rollback();
        dialog.close(2);
    };

    targetCharInput.onChange = function() {
        updateInfoText();
        updatePreview();
    };

    okBtn.onClick = function() {
        // Undoを1回にまとめて確定 / Confirm as a single undo step
        previewMgr.confirm(function() {
            applyAllCurrentValues();
        });
        targetCharInput.onChange = null;
        dialog.close();
    };

    var state = getTargetCharState();
    var originalSize = getOriginalCharSize(state.targetChar, state.hasTargetChar);
    if (originalSize) {
        uiElements.sizeInput.text = originalSize.toFixed(2);
    }

    setUnifiedLabelWidth();
    updateInfoText();

    uiElements.sizeInput.onChange = function() {
        updatePreview();
        updateInfoText();
        updateApparentSizeDisplay(unitLabel);
    };
    changeValueByArrowKey(uiElements.sizeInput, { step: 1, shiftStep: 10, altStep: 0.1 });

    uiElements.hScaleInput.onChange = function() {
        updatePreview();
        updateInfoText();
        updateApparentSizeDisplay(unitLabel);
    };
    changeValueByArrowKey(uiElements.hScaleInput, { step: 1, shiftStep: 10, altStep: 5 });

    uiElements.sizeInput.onChanging = function() {
        updateApparentSizeDisplay(unitLabel);
    };
    uiElements.hScaleInput.onChanging = function() {
        updateApparentSizeDisplay(unitLabel);
    };

    updateApparentSizeDisplay(unitLabel);

    dialog.onShow = (function(origOnShow) {
        return function() {
            if (typeof origOnShow === "function") origOnShow();
            uiElements.hScaleInput.active = true;
        };
    })(dialog.onShow);

    // 初期状態も管理下で1回プレビュー適用（この時点でundoDepth=1になる）
    updatePreview();

    dialog.show();
    return;
}

main();

/* カーニングを対象文字に適用 / Apply kerning to matching characters */
function applyKerningToChars(range, targetChar, hasTargetChar) {
    if (!range || !range.characters || targetChar === undefined) return;
    var val = parseFloat(uiElements.kerningInput.text);
    if (isNaN(val)) return;
    var chars = range.characters;
    for (var i = 0; i < chars.length; i++) {
        var ch = chars[i];
        if (isTargetChar(ch, targetChar, hasTargetChar)) {
            ch.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
            ch.kerning = val;
        }
    }
}