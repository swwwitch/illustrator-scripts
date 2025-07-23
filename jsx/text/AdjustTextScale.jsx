#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script name:
AdjustTextScale.jsx

### Readme （GitHub）/ Readme (GitHub):
https://github.com/swwwitch/illustrator-scripts

### 概要 / Overview:
- テキストのフォントサイズまたは横比率を数値入力で調整する / Adjust font size or horizontal scale by numeric input
- キーボードでの微調整にも対応（上下キー） / Supports fine-tuning with keyboard (up/down keys)

### 主な機能 / Main features:
- フォントサイズまたはスケール比率での編集 / Edit by font size or scale ratio
- 見かけのサイズ表示 / Apparent size display
- リアルタイムプレビュー / Real-time preview
- リセット、キャンセル対応 / Reset and cancel support

### 処理の流れ / Processing flow:
1. ダイアログボックスを表示し、数値を入力 / Show dialog and input value
2. 現在の選択テキストのサイズ情報を取得 / Get size info of current selected text
3. 入力値に応じてスケーリング処理 / Scale processing by input value
4. OK押下時に値を確定、Cancel/Reset対応 / Confirm on OK, support Cancel/Reset

### 更新履歴 / Update history:
- v1.0 (20250723) : 初期バージョン / Initial version
*/

var SCRIPT_VERSION = "v1.0"; // スクリプトのバージョン / Script version

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

var LABELS = {
    dialogTitle: {
        ja: "フォント情報を追加",
        en: "Add Font Info"
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
        ja: "比率",
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
    }
};

var apparentSizeDisplay; // グローバルに定義 / Global definition

/*
文字選択状態を取得 / Get selected characters from current document
*/
function isTargetChar(ch, targetChar, hasTargetChar) {
    return !hasTargetChar || targetChar.indexOf(ch.contents) !== -1;
}

/*
文字選択状態を取得 / Get selected characters from current document
*/
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
    /*
    文字選択状態を取得
    Get selected characters from current document
    */
    if (!(sel instanceof Array) && sel instanceof TextRange) {
        res.push(sel);
    }
    return res;
}

/*
英数字を除いたユニークな文字列を取得 / Get unique non-alphanumeric characters from text
*/
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

/*
対象文字に一致する文字だけに処理を適用 / Apply callback to matching characters in text range
*/
function applyToTargetCharacters(range, targetChar, hasTargetChar, callback) {
    var chars = range.characters;
    for (var j = 0; j < chars.length; j++) {
        var ch = chars[j];
        if (isTargetChar(ch, targetChar, hasTargetChar)) {
            callback(ch);
        }
    }
}

/*
指定文字にだけベースラインシフトを適用 / Apply baseline shift only to matching characters
*/
function applyBaselineShiftToChars(range, targetChar, hasTargetChar) {
    if (!range || !range.characters || targetChar === undefined) return;

    var val = parseFloat(baselineInput.text);
    if (isNaN(val)) return;

    var chars = range.characters;
    for (var i = 0; i < chars.length; i++) {
        var ch = chars[i];
        if (isTargetChar(ch, targetChar, hasTargetChar)) {
            ch.characterAttributes.baselineShift = val;
        }
    }
}

/*
ダイアログの表示位置をオフセット / Shift dialog position by offsetX and offsetY
*/
function shiftDialogPosition(dlg, offsetX, offsetY) {
    dlg.onShow = function() {
        var currentX = dlg.location[0];
        var currentY = dlg.location[1];
        dlg.location = [currentX + offsetX, currentY + offsetY];
    };
}

/*
ダイアログの透明度を設定 / Set dialog opacity
*/
function setDialogOpacity(dlg, opacityValue) {
    dlg.opacity = opacityValue;
}

/*
上下キーによる数値の増減を有効にする関数 / Enable value increment/decrement by arrow keys
*/
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
            value += delta;
        } else if (event.keyName == "Down") {
            value -= delta;
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
        /*
        数値変更時にリアルタイムプレビューを実行
        Apply real-time preview when value changes
        */
        if (
            (typeof sizeInput !== "undefined" && editText === sizeInput) ||
            (typeof hScaleInput !== "undefined" && editText === hScaleInput) ||
            (typeof baselineInput !== "undefined" && editText === baselineInput)
        ) {
            /*
            ベースラインシフト入力時は即時プレビューを適用
            Apply baseline shift preview on change
            */
            if (typeof baselineInput !== "undefined" && editText === baselineInput) {
                baselineInput.text = editText.text; // 明示的に更新
                applyBaselineShiftToTargetChar();
            }
        }
    });
}

/*
単位コードとプリファレンスキーに応じて単位ラベルを返す関数 / Get unit label for code and pref key
*/
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

/*
現在のテキスト単位ラベルを取得 / Get current text unit label
*/
function getCurrentUnitLabel() {
    var unitCode = app.preferences.getIntegerPreference("text/units");
    return getUnitLabel(unitCode, "text/units");
}


/* メイン処理 / Main function */
function main() {
    // 事前宣言：sizeInput, hScaleInput, baselineInput（未定義エラー回避のため）
    var sizeInput; // フォントサイズ入力
    var hScaleInput; // 水平スケール入力
    var baselineInput; // ベースラインシフト入力
    // ドキュメントが開かれていなければ終了 / Exit if no document is open
    if (app.documents.length <= 0) {
        return;
    }

    var targetRanges = getTextSelection();
    // 選択が空のときに警告を表示 / Show alert if no text selected
    if (targetRanges.length === 0) {
        alert(LABELS.selectTextAlert[lang]);
        return;
    }

    var allSelText = "";
    for (var i = 0; i < targetRanges.length; i++) {
        allSelText += targetRanges[i].contents;
    }

    /* 選択テキストから共通の文字列を取得 / Get common characters from selected texts */
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

    /* 入力値取得 / Get current state from inputs */
    function getState() {
        var targetChar = targetCharInput.text;
        var hasTargetChar = targetChar && targetChar.length > 0;
        // originalSize: 選択範囲から対象文字の最初のサイズを取得 / Get first size of target character from selection
        var originalSize = null;
        for (var i = 0, len = targetRanges.length; i < len; i++) {
            var chars = targetRanges[i].characters;
            for (var j = 0; j < chars.length; j++) {
                var ch = chars[j];
                if (isTargetChar(ch, targetChar, hasTargetChar)) {
                    originalSize = ch.size;
                    break;
                }
            }
            if (originalSize !== null) break;
        }
        return {
            targetChar: targetChar,
            hasTargetChar: hasTargetChar,
            originalSize: originalSize
            // chars: chars
        };
    }

    /* 対象文字に一致する最初の文字を取得 / Find first character matching target */
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

    // applyChanges() 関数は不要のため削除 / applyChanges() function not needed, removed

    /* 情報テキストを更新 / Update info text with example character */
    function updateInfoText() {
        var state = getState();
        var exampleChar = findFirstMatchingChar(state.targetChar, state.hasTargetChar);
        if (!exampleChar) {
            /* 情報をクリア / Clear info fields */
            sizeInput.text = "";
            hScaleInput.text = "";
            apparentSizeText.text = "--";
            return;
        }
        var size = exampleChar.size;
        var h = exampleChar.characterAttributes.horizontalScale;
        var v = exampleChar.characterAttributes.verticalScale;
        var visualSize = Math.round(size * (h / 100) * 10) / 10;
        size = Math.round(size * 10) / 10;
        h = Math.round(h * 10) / 10;
        /* 情報を更新 / Update information */
        sizeInput.text = size + "";
        hScaleInput.text = h + "";
        updateApparentSizeDisplay(unitLabel);
    }

    /* 見かけサイズ計算ロジックを共通関数化 / Function to update apparent size display */
    function updateApparentSizeDisplay(unitLabel) {
        var size = parseFloat(sizeInput.text, 10);
        var scale = parseFloat(hScaleInput.text, 10);
        if (isNaN(size) || isNaN(scale)) {
            apparentSizeText.text = "--";
        } else {
            var apparent = Math.round(size * scale) / 100;
            apparentSizeText.text = apparent + "";
        }
    }

    // handleInputChange() は updateInfoText() に統合するため削除 / handleInputChange() merged into updateInfoText, removed

    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.alignChildren = "left";

    /*
    位置と透明度の調整 / Adjust position and opacity
    */
    var offsetX = 300;
    var dialogOpacity = 0.97;
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);

    // キーボードでラジオボタン切り替えを有効にする / Enable radio button switching by keyboard
    // dialog, modeFontRadio, modeScaleRadio は下で定義されるため、ラジオボタン定義後に呼ぶ / Call after radio button definition

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

    /*
    対象文字入力グループ / Target characters input group
    */
    var targetTextGroup = leftCol.add("group");
    targetTextGroup.orientation = "row";
    targetTextGroup.add("statictext", undefined, LABELS.targetChar[lang]);
    var targetCharInput = targetTextGroup.add("edittext", undefined, uniqueNonAN);
    targetCharInput.characters = 10;

    /*
    文字サイズ・比率グループ / Font size & scale group
    */
    var infoPanel = leftCol.add("panel", undefined, LABELS.adjust[lang]);
    infoPanel.orientation = "column";
    infoPanel.alignChildren = ["left", "top"];
    infoPanel.margins = [15, 10, 15, 10];

    // --- ラジオボタングループを作成し、infoPanelの直下に追加 / Add radio button group under infoPanel ---
    var modeGroup = infoPanel.add("group");
    modeGroup.orientation = "column";
    modeGroup.alignChildren = "left";

    // ラジオボタングループを削除し、statictext追加 / Remove radio button group and add statictext
    var radioGroup = modeGroup.add("group");
    radioGroup.orientation = "column";
    radioGroup.alignChildren = "left";

    /*
    サイズ入力グループ / Font size input group
    */
    var labelWidth = 120;
    var sizeGroup = radioGroup.add("group", undefined, {
        orientation: "row"
    });
    sizeGroup.margins = [0, 10, 0, 0];
    sizeGroup.alignChildren = "left";
    var sizeLabel = sizeGroup.add("statictext", undefined, LABELS.fontSize[lang]);
    sizeLabel.justify = "right";
    sizeInput = sizeGroup.add("edittext", undefined, "0");
    sizeInput.characters = 4;
    sizeInput.enabled = true;

    // 単位ラベルを一度だけ取得し再利用 / Get unit label once and reuse
    var unitLabel = getUnitLabel(app.preferences.getIntegerPreference("text/units"), "text/units");
    sizeGroup.add("statictext", undefined, unitLabel);

    /*
    比率入力グループ / Scale input group
    */
    var scaleGroup = radioGroup.add("group", undefined, {
        orientation: "row"
    });
    scaleGroup.margins = [0, 0, 0, 6];
    var scaleLabel = scaleGroup.add("statictext", undefined, LABELS.scale[lang]);
    scaleLabel.justify = "right";
    hScaleInput = scaleGroup.add("edittext", undefined, "100");
    hScaleInput.characters = 4;
    scaleGroup.add("statictext", undefined, "%");
    // 入力欄はどちらも有効化 / Both input fields enabled
    sizeInput.enabled = true;
    hScaleInput.enabled = true;




    /*
    見かけサイズ表示グループ / Apparent size display group
    */
    var displayGroup = infoPanel.add("group", undefined, {
        orientation: "row"
    });
    var apparentLabel = displayGroup.add("statictext", undefined, LABELS.apparent[lang]);
    apparentLabel.justify = "right";
    var apparentSizeText = displayGroup.add("statictext", undefined, "--");
    apparentSizeText.characters = 5;
    displayGroup.add("statictext", undefined, unitLabel);


    /*
    ベースラインシフト / Baseline Shift
    */
    var baselineGroup = infoPanel.add("group");
    baselineGroup.orientation = "row";
    baselineGroup.alignChildren = ["right", "center"];

    var baselineLabel = baselineGroup.add("statictext", undefined, LABELS.baselineShiftLabel[lang]);
    baselineLabel.justify = "right";

    // ラベル幅を共通化する関数
    function setUnifiedLabelWidth() {
        var labelWidth = 100;
        sizeLabel.preferredSize.width = labelWidth;
        scaleLabel.preferredSize.width = labelWidth;
        apparentLabel.preferredSize.width = labelWidth;
        baselineLabel.preferredSize.width = labelWidth;
    }

    baselineInput = baselineGroup.add("edittext", undefined, "0");
    baselineInput.characters = 4;
    baselineInput.justify = "right";
    baselineGroup.add("statictext", undefined, unitLabel);


    /*
    ベースラインシフトのプレビュー反映処理
    Apply baseline shift preview on change
    */
    function applyBaselineShiftToTargetChar() {
        var state = getState();
        for (var i = 0; i < targetRanges.length; i++) {
            applyBaselineShiftToChars(targetRanges[i], state.targetChar, state.hasTargetChar);
        }
        app.redraw();
    }
    baselineInput.onChange = applyBaselineShiftToTargetChar;
    baselineInput.onChanging = function() {
        applyBaselineShiftToTargetChar();
        updateBaselineShift();
    };
    // 上下キー対応
    changeValueByArrowKey(baselineInput, {
        step: 1,
        shiftStep: 10,
        altStep: 0.1
    });

    /*
    入力欄で上下キーによる数値調整を有効化
    Enable up/down key adjustment for input fields
    */
    function addKeyEvents(input) {
        changeValueByArrowKey(input, {
            step: 1,
            shiftStep: 10,
            altStep: 0.1
        });
    }
    addKeyEvents(baselineInput);



    /*
    ボタングループ / Button group
    */
    var buttonGroup = rightCol.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.orientation = "column";
    var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang]);
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
    // スペーサーの高さを大きく (例: 20) / Increase spacer height (e.g. 20)
    var cancelResetSpacer = buttonGroup.add("statictext", undefined, "");
    cancelResetSpacer.preferredSize.height = 50;
    var resetBtn = buttonGroup.add("button", undefined, LABELS.reset[lang]);
    // 3ボタンの幅を揃える / Set same width for 3 buttons
    resetBtn.preferredSize.width = 90;
    cancelBtn.preferredSize.width = 90;
    okBtn.preferredSize.width = 90;

    /*
    入力値と選択文字を初期状態にリセット
    Reset all fields and selection to initial state
    */
    function resetToInitial() {
        targetCharInput.text = uniqueNonAN;
        var state = getState();
        var exampleChar = findFirstMatchingChar(state.targetChar, state.hasTargetChar);
        if (!exampleChar) return;
        var initialSize = exampleChar.size;
        sizeInput.text = Math.round(initialSize * 10) / 10 + "";
        hScaleInput.text = "100";
        baselineInput.text = "0"; // ベースラインシフトも初期値 0 にリセット
        var sel = app.activeDocument.selection;
        if (sel.length > 0 && sel[0].typename === "TextFrame") {
            sel[0].textRange.characterAttributes.baselineShift = 0;
        }
        for (var i = 0, len = targetRanges.length; i < len; i++) {
            var chars = targetRanges[i].characters;
            for (var j = 0; j < chars.length; j++) {
                var ch = chars[j];
                ch.characterAttributes.size = initialSize;
                ch.characterAttributes.horizontalScale = 100;
                ch.characterAttributes.verticalScale = 100;
                ch.characterAttributes.baselineShift = 0;
            }
        }
        updateInfoText();
        app.redraw();
    }
    resetBtn.onClick = function() {
        resetToInitial();
    };
    cancelBtn.onClick = function() {
        dialog.close(2);
    };


    /*
    対象文字欄の変更で情報を更新
    Update info fields when target character input changes
    */
    targetCharInput.onChange = updateInfoText;

    /*
    OKボタン押下時にイベントハンドラ解除・ダイアログ終了
    Remove handlers and close dialog on OK
    */
    okBtn.onClick = function() {
        targetCharInput.onChange = null;
        dialog.close();
    };

    // 初期設定: getState直後に originalSize があれば sizeInput.text に設定 / Set sizeInput.text if originalSize exists after getState
    var state = getState();
    if (state.originalSize) {
        sizeInput.text = state.originalSize.toFixed(2);
    }
    // ラベル幅を統一
    setUnifiedLabelWidth();
    // 初期設定として明示的に呼び出し / Explicitly call as initial setting
    updateInfoText();

    /*
    共通適用ロジック: サイズまたはスケールを対象文字に適用
    Apply size or scale to target characters
    */
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

    /*
    文字サイズ直接入力時の処理
    Handle direct input of font size
    */
    sizeInput.onChange = function() {
        var val = Number(sizeInput.text) || 12;
        if (!isNaN(val)) {
            var state = getState();
            applyPropertyToCharacters("size", val, state);
            app.redraw();
            updateInfoText();
        }
        updateApparentSizeDisplay(unitLabel);
    };
    // 上下キーによる数値の増減を有効にする（小数第1位まで対応）
    // Enable up/down key for value adjustment (to 1 decimal place)
    changeValueByArrowKey(sizeInput, {
        step: 1,
        shiftStep: 10,
        altStep: 0.1
    });

    /*
    スケール比率直接入力時の処理
    Handle direct input of scale ratio
    */
    hScaleInput.onChange = function() {
        var val = Number(hScaleInput.text) || 100;
        if (!isNaN(val)) {
            var state = getState();
            applyPropertyToCharacters("scale", val, state);
            app.redraw();
            updateInfoText();
        }
        updateApparentSizeDisplay(unitLabel);
    };
    // 上下キーによる数値の増減を有効にする（整数単位）
    // Enable up/down key for value adjustment (integer)
    changeValueByArrowKey(hScaleInput, {
        step: 1,
        shiftStep: 10,
        altStep: 5
    });

    /*
    sizeInput, hScaleInput の onChanging で見かけサイズを更新
    Update apparent size on changing of size or scale input
    */
    sizeInput.onChanging = function() {
        updateApparentSizeDisplay(unitLabel);
    };
    hScaleInput.onChanging = function() {
        updateApparentSizeDisplay(unitLabel);
    };

    // ダイアログ初期化直後にも見かけサイズ表示を更新 / Update apparent size display immediately after dialog initialization
    updateApparentSizeDisplay(unitLabel);
    // ダイアログ表示後に比率入力欄をハイライト / Highlight scale input after dialog is shown
    dialog.onShow = (function(origOnShow) {
        return function() {
            if (typeof origOnShow === "function") origOnShow();
            hScaleInput.active = true;
        };
    })(dialog.onShow);
    dialog.show();
    // プレビュー済みの内容を確定 / Confirm previewed content
    return;
}

/*
メイン実行部 / Entry point
*/
main();