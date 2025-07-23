#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名 / Script Name：
AdjustTextScale.jsx

### 概要 / Overview：
選択されたテキストに対して、指定文字に一致する部分のみ、
- 水平／垂直スケールを均等に変更
- またはフォントサイズを加算／減算可能
- ダイアログでリアルタイムにプレビュー反映しながら調整できる

### 主な機能 / Main Features：
- 指定文字のみを対象にスケーリングまたはサイズ変更
- 上下キー、Shift、Alt による数値変更に対応
- スケール初期化（リセット）ボタン付き
- 日本語・英語ラベル切替対応

### 更新履歴 / Change Log：
- v0.1 (20250723): 初期バージョン
- v0.2 (20250723): 調整
- v0.3 (20250723): 調整
*/

var SCRIPT_VERSION = "v0.2";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* ラベル定義を日本語／英語形式で統一 / Unified label definitions (Japanese/English) */
var LABELS = {
    dialogTitle: {
        ja: "文字サイズと比率の調整 " + SCRIPT_VERSION,
        en: "Adjust Font Size and Scale " + SCRIPT_VERSION
    },
    targetCharsLabel: {
        ja: "対象文字：",
        en: "Target Characters:"
    },
    applyTargetPanel: {
        ja: "適用対象",
        en: "Apply To"
    },
    modeScale: {
        ja: "水平比率／垂直比率",
        en: "Horizontal / Vertical Scale"
    },
    modeFontSize: {
        ja: "文字サイズ",
        en: "Font Size"
    },
    incrementLabel: {
        ja: "増減値：",
        en: "Increment:"
    },
    reset: {
        ja: "リセット",
        en: "Reset"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    alertNoText: {
        ja: "テキストを選択してください",
        en: "Please select a text object"
    },
    alertNegativeScale: {
        ja: "負のスケール値は非推奨です",
        en: "Negative scale may flip text"
    }
};

/* 対象文字の一致判定 / Check if character matches target */
function isTargetChar(ch, targetChar, hasTargetChar) {
    return !hasTargetChar || targetChar.indexOf(ch.contents) !== -1;
}

/* 選択されたテキスト範囲（TextFrame / TextRange）を取得 / Get selected text ranges (TextFrame / TextRange) */
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
    /* sel が配列でない（単体 TextRange の場合）も対応 / Also handle single TextRange if sel is not array */
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

/* ダイアログの表示位置をオフセットする / Shift dialog position by offsetX and offsetY */
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

/* 入力欄で上下キーによる値の増減を可能にする / Enable arrow key increment/decrement on edittext
   Shiftキー押下で10単位増減 / Shift key changes by 10
   Altキー押下で小数点1桁まで許可 / Alt key allows one decimal place */
function changeValueByArrowKey(editText) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;
        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;
        if (keyboard.shiftKey) {
            delta = 10;
            if (event.keyName == "Up") {
                value = Math.ceil((value + 1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 1) / delta) * delta;
                event.preventDefault();
            }
        } else if (keyboard.altKey) {
            delta = 0.1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        } else {
            delta = 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
        }
        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10; /* 小数第1位まで / Round to 1 decimal */
        } else {
            value = Math.round(value); /* 整数に丸め / Round to integer */
        }
        event.preventDefault();
        /* 値を更新してから onChange を通知 / Update value first, then notify */
        editText.text = value;
        editText.notify("onChange");
        /* プレビュー更新 / Update preview */
        if (typeof handleInputChange === "function") {
            handleInputChange(); /* 最新の値で再評価 / Re-evaluate with latest value */
        }
    });
}

/* メイン処理 / Main function */
function main() {
    /* ドキュメントが開かれていなければ終了 / Exit if no document is open */
    if (app.documents.length <= 0) {
        return;
    }

    var targetRanges = getTextSelection();
    /* 選択が空のときに警告を表示 / Show alert if no text selected */
    if (targetRanges.length === 0) {
        alert(getLabelText(LABELS.alertNoText));
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

    /* 状態を取得 / Get current state from inputs */
    function getState() {
        var mode = (modeScale.value) ? "scale" : "fontSize";
        var increment = parseFloat(input.text);
        var targetChar = targetCharInput.text;
        var hasTargetChar = targetChar && targetChar.length > 0;
        return {
            mode: mode,
            increment: increment,
            targetChar: targetChar,
            hasTargetChar: hasTargetChar
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

    /* 変更を適用 / Apply changes based on mode and increment */
    function applyChanges(mode, increment, targetChar, hasTargetChar) {
        /* 負のスケール値で警告を表示 / Warn if negative scale */
        if (mode === "scale" && increment < -99) {
            alert(getLabelText(LABELS.alertNegativeScale));
        }
        if (increment === 0) return; // 0のときは処理しない / Skip if increment is zero
        var changed = false; // 変更があれば redraw
        for (var i = 0, len = targetRanges.length; i < len; i++) {
            var range = targetRanges[i];
            if (mode === "fontSize") {
                applyToTargetCharacters(range, targetChar, hasTargetChar, function(ch) {
                    var before = ch.size;
                    ch.size += increment;
                    if (ch.size !== before) changed = true;
                });
            } else {
                applyToTargetCharacters(range, targetChar, hasTargetChar, function(ch) {
                    var v = ch.characterAttributes.verticalScale;
                    var h = ch.characterAttributes.horizontalScale;
                    /* 縦横比を統一してから調整 / Unify aspect ratio before applying scale */
                    if (v !== h) h = v;
                    var newV = v + increment;
                    var newH = h + increment;
                    if (ch.characterAttributes.verticalScale !== newV || ch.characterAttributes.horizontalScale !== newH) {
                        ch.characterAttributes.verticalScale = newV;
                        ch.characterAttributes.horizontalScale = newH;
                        changed = true;
                    }
                });
            }
        }
        if (changed) app.redraw();
        updateInfoText();
    }

    /* 情報テキストを更新 / Update info text with example character */
    function updateInfoText() {
        var state = getState();
        var exampleChar = findFirstMatchingChar(state.targetChar, state.hasTargetChar);
        if (!exampleChar) {
            /* 情報をクリア / Clear info fields */
            sizeInput.text = "";
            hScaleInput.text = "";
            apparentSizeDisplay.text = "—";
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
        updateApparentSize();
    }

    // 「見かけ」サイズを計算して表示する関数
    function updateApparentSize() {
        var size = parseFloat(sizeInput.text);
        var scale = parseFloat(hScaleInput.text);
        if (!isNaN(size) && !isNaN(scale)) {
            var apparent = size * scale / 100;
            apparentSizeDisplay.text = apparent.toFixed(2);
        } else {
            apparentSizeDisplay.text = "—";
        }
    }

    /* 入力変更時の処理 / Handle input changes */
    function handleInputChange() {
        var state = getState();
        // 入力値が不正な場合エラー表示 / Show error if NaN
        if (isNaN(state.increment)) {
            /* 無効な数値をクリア / Clear on invalid number */
            sizeInput.text = "";
            hScaleInput.text = "";
            return;
        }
        // input.text = "0" のときも updateInfoText() を実行 / Also update info on "0"
        if (!isNaN(state.increment)) {
            if (state.increment !== 0) {
                applyChanges(state.mode, state.increment, state.targetChar, state.hasTargetChar);
            }
            updateInfoText();
        }
        // --- 増減値変更時に「サイズ」「比率」も再計算してフィールド反映 / Also recalc fields on increment change ---
        // currentStateの更新 / Update currentState
        var exampleChar = findFirstMatchingChar(state.targetChar, state.hasTargetChar);
        if (exampleChar) {
            currentState.size = exampleChar.size;
            currentState.h = exampleChar.characterAttributes.horizontalScale;
        }
        // modeScale/radio選択時は比率を、modeFont/radio選択時はサイズを再表示 / Show scale or size by mode
        if (modeScale.value) {
            hScaleInput.text = Math.round(currentState.h);
        } else if (modeFont.value) {
            sizeInput.text = Math.round(currentState.size * 10) / 10;
        }
    }

    var dialog = new Window("dialog", getLabelText(LABELS.dialogTitle));
    dialog.alignChildren = "left";

    /* 位置と透明度の調整 / Adjust position and opacity */
    var offsetX = 300;
    var dialogOpacity = 0.97;
    setDialogOpacity(dialog, dialogOpacity);
    shiftDialogPosition(dialog, offsetX, 0);

    var contentGroup = dialog.add("group");
    contentGroup.orientation = "row";
    contentGroup.alignChildren = "top";

    var leftCol = contentGroup.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = "left";

    var rightCol = contentGroup.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = "right";

    // 1. 対象文字グループ
    var targetTextGroup = leftCol.add("group");
    targetTextGroup.orientation = "row";
    targetTextGroup.add("statictext", undefined, getLabelText(LABELS.targetCharsLabel));
    var targetCharInput = targetTextGroup.add("edittext", undefined, uniqueNonAN);
    targetCharInput.characters = 10;

    // 2. 適用対象グループ
    var modeGroup = leftCol.add("panel", undefined, getLabelText(LABELS.applyTargetPanel));
    modeGroup.orientation = "column";
    modeGroup.alignChildren = "left";
    modeGroup.margins = [15, 20, 15, 10];
    var modeScale = modeGroup.add("radiobutton", undefined, getLabelText(LABELS.modeScale));
    var modeFont = modeGroup.add("radiobutton", undefined, getLabelText(LABELS.modeFontSize));
    modeScale.value = true;

    // 3. 増減値グループ
    var stepGroup = leftCol.add("group");
    stepGroup.orientation = "row";
    stepGroup.add("statictext", undefined, getLabelText(LABELS.incrementLabel));
    var input = stepGroup.add("edittext", undefined, "0");
    input.characters = 5;

    // 4. 文字サイズ・比率グループ
    var infoPanel = leftCol.add("panel", undefined, LABELS.scaleRatioPanel ? LABELS.scaleRatioPanel[lang] : "サイズと比率");
    infoPanel.orientation = "column";
    infoPanel.alignChildren = ["left", "top"];
    infoPanel.margins = [15, 10, 15, 10];

    // サイズ・比率を横並びで
    var rowGroup = infoPanel.add("group");
    rowGroup.orientation = "row";

    // サイズ入力フィールド
    var sizeGroup = rowGroup.add("group");
    sizeGroup.orientation = "column";
    sizeGroup.add("statictext", undefined, LABELS.size ? LABELS.size[lang] || "サイズ:" : "サイズ:");
    var sizeInput = sizeGroup.add("edittext", undefined, "0");
    sizeInput.characters = 3;
    sizeGroup.add("statictext", undefined, getUnitLabel(app.preferences.getIntegerPreference("typeUnits"), "typeUnits"));

    // 比率入力フィールド
    var scaleGroup = rowGroup.add("group");
    scaleGroup.orientation = "column";
    scaleGroup.alignChildren = ["center", "top"];
    scaleGroup.alignment = ["left", "top"];
    scaleGroup.add("statictext", undefined, LABELS.scale ? LABELS.scale[lang] || "比率:" : "比率:");
    var hScaleInput = scaleGroup.add("edittext", undefined, "100");
    hScaleInput.characters = 4;
    scaleGroup.add("statictext", undefined, "%");

    // 見かけのサイズ表示グループ（右）
    var displayGroup = rowGroup.add("group");
    displayGroup.orientation = "column";

    displayGroup.alignChildren = ["center", "top"];
    scaleGroup.alignment = ["left", "top"];

    displayGroup.add("statictext", undefined, LABELS.apparentSizeLabel ? LABELS.apparentSizeLabel[lang] : "見かけ");
    var apparentSizeDisplay = displayGroup.add("statictext", undefined, "—");
    apparentSizeDisplay.characters = 6;


    changeValueByArrowKey(input);

    var buttonGroup = rightCol.add("group");
    buttonGroup.alignment = "right";
    buttonGroup.orientation = "column";
    var okBtn = buttonGroup.add("button", undefined, getLabelText(LABELS.ok));
    var cancelBtn = buttonGroup.add("button", undefined, getLabelText(LABELS.cancel));
    buttonGroup.add("statictext", undefined, " ");
    var resetBtn = buttonGroup.add("button", undefined, getLabelText(LABELS.reset));
    resetBtn.onClick = function() {
        input.text = "0";
        targetCharInput.text = uniqueNonAN;
        modeScale.value = true;
        modeFont.value = false;
        /* スケールをリセット / Reset scale to 100% */
        for (var i = 0, len = targetRanges.length; i < len; i++) {
            var chars = targetRanges[i].characters;
            for (var j = 0; j < chars.length; j++) {
                chars[j].characterAttributes.horizontalScale = 100;
                chars[j].characterAttributes.verticalScale = 100;
            }
        }
        updateInfoText(); // プレビュー更新のみ / Preview only, no reapplication
        app.redraw();
    };
    cancelBtn.onClick = function() {
        dialog.close(2);
    };

    /* --- 即時プレビュー用イベントハンドラの登録（初回のみ）/ Set live preview event handlers (only once) --- */
    input.onChange = handleInputChange;
    modeScale.onClick = handleInputChange;
    modeFont.onClick = handleInputChange;
    targetCharInput.onChange = handleInputChange;

    // 「OK」ボタン押下時に onChange/onClick を無効化 / Disable onChange/onClick on OK
    okBtn.onClick = function() {
        input.onChange = null;
        modeScale.onClick = null;
        modeFont.onClick = null;
        targetCharInput.onChange = null;
        dialog.close();
    };

    // 対象文字の現在のサイズ・比率を初期値として設定 / Set current size/scale as initial values
    (function setInitialSizeInput() {
        var state = getState();
        var exampleChar = findFirstMatchingChar(state.targetChar, state.hasTargetChar);
        if (exampleChar) {
            var size = Math.round(exampleChar.size * 10) / 10;
            var h = Math.round(exampleChar.characterAttributes.horizontalScale * 10) / 10;
            sizeInput.text = size + "";
            hScaleInput.text = h + "";
        }
    })();

    // --- sizeInput, hScaleInput の onChange: 値変更で applyChanges を呼び出す / Call applyChanges on value change ---
    var currentState = {
        mode: (modeScale.value) ? "scale" : "fontSize",
        increment: parseFloat(input.text),
        targetChar: targetCharInput.text,
        hasTargetChar: targetCharInput.text && targetCharInput.text.length > 0,
        size: parseFloat(sizeInput.text),
        h: parseFloat(hScaleInput.text)
    };

    sizeInput.onChange = function() {
        var val = parseFloat(sizeInput.text);
        if (!isNaN(val)) {
            currentState.size = val;
            // 例: サイズ変更の適用
            // mode: "fontSize", increment: 差分, etc.
            // ここでは選択文字のうち対象文字に対してサイズを val にする
            for (var i = 0, len = targetRanges.length; i < len; i++) {
                applyToTargetCharacters(targetRanges[i], currentState.targetChar, currentState.hasTargetChar, function(ch) {
                    ch.size = val;
                });
            }
            app.redraw();
            updateInfoText();
        }
        updateApparentSize();
    };

    // サイズフィールドのキー操作によるプレビュー更新 / Preview update on key in size field
    sizeInput.addEventListener("keydown", function(event) {
        if (event.keyName == "Up" || event.keyName == "Down") {
            var current = parseFloat(sizeInput.text);
            if (!isNaN(current)) {
                if (event.keyName == "Up") {
                    current += 0.1;
                } else if (event.keyName == "Down") {
                    current -= 0.1;
                }
                current = Math.round(current * 10) / 10;
                sizeInput.text = current;
                currentState.size = current;
                // 適用
                for (var i = 0, len = targetRanges.length; i < len; i++) {
                    applyToTargetCharacters(targetRanges[i], currentState.targetChar, currentState.hasTargetChar, function(ch) {
                        ch.size = current;
                    });
                }
                app.redraw();
                updateInfoText();
            }
            updateApparentSize();
            event.preventDefault();
        }
    });

    hScaleInput.onChange = function() {
        var val = parseFloat(hScaleInput.text);
        if (!isNaN(val)) {
            currentState.h = val;
            // 対象文字の horizontalScale/verticalScale を val にする
            for (var i = 0, len = targetRanges.length; i < len; i++) {
                applyToTargetCharacters(targetRanges[i], currentState.targetChar, currentState.hasTargetChar, function(ch) {
                    ch.characterAttributes.horizontalScale = val;
                    ch.characterAttributes.verticalScale = val;
                });
            }
            app.redraw();
            updateInfoText();
        }
        updateApparentSize();
    };

    // 比率フィールドのキー操作によるプレビュー更新 / Preview update on key in scale field
    hScaleInput.addEventListener("keydown", function(event) {
        if (event.keyName == "Up" || event.keyName == "Down") {
            var current = parseFloat(hScaleInput.text);
            if (!isNaN(current)) {
                if (event.keyName == "Up") {
                    current += 1;
                } else if (event.keyName == "Down") {
                    current -= 1;
                }
                current = Math.round(current);
                hScaleInput.text = current;
                currentState.h = current;
                // 適用
                for (var i = 0, len = targetRanges.length; i < len; i++) {
                    applyToTargetCharacters(targetRanges[i], currentState.targetChar, currentState.hasTargetChar, function(ch) {
                        ch.characterAttributes.horizontalScale = current;
                        ch.characterAttributes.verticalScale = current;
                    });
                }
                app.redraw();
                updateInfoText();
            }
            updateApparentSize();
            event.preventDefault();
        }
    });

    // sizeInput, hScaleInput の onChanging で見かけサイズを更新
    sizeInput.onChanging = updateApparentSize;
    hScaleInput.onChanging = updateApparentSize;

    // 「増減値」フィールドを選択状態にする
    input.active = true;
    dialog.show();
    /* プレビュー済みの内容を確定 / No re-processing needed on OK */
    return;
}

/* ラベルを言語に応じて取得 / Get label text according to language */
function getLabelText(labelObj) {
    try {
        if (app.locale && app.locale.indexOf("ja") === 0) {
            return labelObj.ja;
        }
    } catch (e) {}
    return labelObj.en;
}
// 単位コードとラベルのマップ
var unitMap = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    5: "Q/H", // 特別処理される
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

// 単位コードとプリファレンスキーに応じて単位ラベルを返す関数
function getUnitLabel(code, prefKey) {
    if (code === 5) {
        var hKeys = {
            "text/asianunits": true,
            "rulerType": true,
            "strokeUnits": true
        };
        return hKeys[prefKey] ? "H" : "Q";
    }
    return unitMap[code] || "pt";
}

// メイン実行部 / Entry point
main();