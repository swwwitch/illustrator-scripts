#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 円とテキストを 1 つずつ選択し、テキストを指定回数繰り返して円のパス上文字に変換する
- 連結文字（半角スペース／半角スペース x2／中黒）を選択でき、中黒の場合は末尾にも連結文字を付与
- 円周に合わせて文字サイズを自動調整（フィット補正率で開始・終了の隙間を微調整）
- 数値フィールドは ↑↓ キーで増減（Shift で ±10・10 の倍数にスナップ、Option で ±0.1）
- プレビュー対応

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.0.0";

// =========================================
// ユーザー設定 / User settings
// =========================================

/* 繰り返し数の初期値 / Default repeat count */
var DEFAULT_REPEAT_COUNT = 3;

/* フィット補正率の初期値（%）/ Default fit correction (%) */
var DEFAULT_CORRECTION = 103;

// =========================================
// ローカライズ / Localization
// =========================================

/* 現在のロケールから言語を判定 / Detect language from the current locale */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var currentLanguage = getCurrentLang();

var LABELS = {
    dialog: {
        title: { ja: "パス上文字に変換", en: "Convert to Path Type" }
    },
    panel: {
        settings: { ja: "設定", en: "Settings" },
        fit:      { ja: "フィット", en: "Fit" }
    },
    label: {
        repeatCount: { ja: "繰り返し数", en: "Repeat count" },
        separator:   { ja: "連結文字", en: "Separator" },
        correction:  { ja: "フィット補正率(%)", en: "Fit correction (%)" }
    },
    radio: {
        space1: { ja: "半角スペース", en: "Space" },
        space2: { ja: "半角スペース x2", en: "Space x2" },
        bullet: { ja: "•（中黒）", en: "• (bullet)" }
    },
    checkbox: {
        preview: { ja: "プレビュー", en: "Preview" },
        fitSize: { ja: "文字サイズを調整してフィット", en: "Adjust font size to fit" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    alert: {
        noDocument: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        selectTwo: {
            ja: "円とテキストを1つずつ選択してください。",
            en: "Please select one circle and one text frame."
        },
        needTextAndPath: {
            ja: "テキストフレームと円のパスを1つずつ選択してください。",
            en: "Please select one text frame and one circle path."
        },
        emptyText: {
            ja: "テキストが空です。",
            en: "The text is empty."
        },
        invalidCount: {
            ja: "繰り返し数には1以上の整数を入力してください。",
            en: "Enter an integer of 1 or more for the repeat count."
        },
        invalidCorrection: {
            ja: "フィット補正率には0より大きい数値を入力してください。",
            en: "Enter a number greater than 0 for the fit correction."
        }
    }
};

/* ネストキーからローカライズ文字列を取得 / Resolve a localized string from a dotted key */
function getLabel(key) {
    var parts = key.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node == null) return key;
        node = node[parts[i]];
    }
    if (node == null) return key;
    return (node[currentLanguage] != null) ? node[currentLanguage] : node.en;
}

/* ローカライズ文字列を返す / Return a localized string */
function L(key) {
    return getLabel(key);
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLabel(key) + (currentLanguage === "ja" ? "：" : ":");
}

(function () {

    /* ドキュメントの有無を確認 / Ensure a document is open */
    if (app.documents.length === 0) {
        alert(L("alert.noDocument"));
        return;
    }

    var activeDoc = app.activeDocument;

    /* 選択数を確認（円とテキストの2つ）/ Ensure exactly two objects are selected */
    if (activeDoc.selection.length !== 2) {
        alert(L("alert.selectTwo"));
        return;
    }

    var selectedItems = activeDoc.selection;
    var sourceTextFrame = null;
    var circlePath = null;

    /* 選択からテキストフレームと円のパスを振り分け / Pick the text frame and the circle path from the selection */
    for (var i = 0; i < selectedItems.length; i++) {
        if (selectedItems[i].typename === "TextFrame") {
            sourceTextFrame = selectedItems[i];
        } else if (selectedItems[i].typename === "PathItem") {
            circlePath = selectedItems[i];
        }
    }

    if (sourceTextFrame === null || circlePath === null) {
        alert(L("alert.needTextAndPath"));
        return;
    }

    if (sourceTextFrame.contents === "") {
        alert(L("alert.emptyText"));
        return;
    }

    var originalText = sourceTextFrame.contents;
    var originalTextHidden = sourceTextFrame.hidden;
    var originalPathHidden = circlePath.hidden;

    var previewTextFrame = null;
    var isUpdatingPreview = false;

    /* 繰り返し文字列を作成（trailing が真なら末尾にも連結文字）/ Build the repeated string (append a trailing separator when trailing is true) */
    function makeRepeatedText(text, repeatCount, separator, trailing) {
        var segments = [];
        for (var i = 0; i < repeatCount; i++) {
            segments.push(text);
        }
        var joined = segments.join(separator);
        if (trailing) {
            joined += separator;
        }
        return joined;
    }

    /* 円・楕円のおおよその周長を取得（Ramanujan 近似）/ Approximate the ellipse perimeter (Ramanujan approximation) */
    function getEllipsePerimeter(ellipseItem) {
        var bounds = ellipseItem.geometricBounds;
        var width = Math.abs(bounds[2] - bounds[0]);
        var height = Math.abs(bounds[1] - bounds[3]);

        var radiusX = width / 2;
        var radiusY = height / 2;

        /* Ramanujan 近似式（h は (a-b)^2/(a+b)^2）/ Ramanujan approximation (ramanujanH = (a-b)^2/(a+b)^2) */
        var ramanujanH = Math.pow(radiusX - radiusY, 2) / Math.pow(radiusX + radiusY, 2);
        var perimeter = Math.PI * (radiusX + radiusY) *
            (1 + (3 * ramanujanH) / (10 + Math.sqrt(4 - 3 * ramanujanH)));

        return perimeter;
    }

    /* 文字属性をコピー / Copy character attributes */
    function copyTextAttributes(sourceFrame, targetFrame) {
        try {
            targetFrame.textRange.characterAttributes.size =
                sourceFrame.textRange.characterAttributes.size;

            targetFrame.textRange.characterAttributes.textFont =
                sourceFrame.textRange.characterAttributes.textFont;

            targetFrame.textRange.characterAttributes.fillColor =
                sourceFrame.textRange.characterAttributes.fillColor;

            targetFrame.textRange.characterAttributes.tracking =
                sourceFrame.textRange.characterAttributes.tracking;

            targetFrame.textRange.characterAttributes.horizontalScale =
                sourceFrame.textRange.characterAttributes.horizontalScale;

            targetFrame.textRange.characterAttributes.verticalScale =
                sourceFrame.textRange.characterAttributes.verticalScale;

            targetFrame.textRange.characterAttributes.baselineShift =
                sourceFrame.textRange.characterAttributes.baselineShift;
        } catch (e) {}
    }

    /* 段落属性をコピー / Copy paragraph attributes */
    function copyParagraphAttributes(sourceFrame, targetFrame) {
        try {
            targetFrame.textRange.paragraphAttributes.justification =
                sourceFrame.textRange.paragraphAttributes.justification;
        } catch (e) {}
    }

    /* 一時テキストで文字幅を測る / Measure text width using a temporary text frame */
    function measureTextWidth(text, fontSize) {
        var tempFrame = activeDoc.textFrames.add();
        tempFrame.contents = text;

        copyTextAttributes(sourceTextFrame, tempFrame);
        tempFrame.textRange.characterAttributes.size = fontSize;

        app.redraw();

        var bounds = tempFrame.geometricBounds;
        var width = Math.abs(bounds[2] - bounds[0]);

        tempFrame.remove();

        return width;
    }

    /* 円周に合う文字サイズを計算 / Calculate the font size that fits the circumference */
    function calculateFitFontSize(text, ellipseItem, correctionPercent) {
        var perimeter = getEllipsePerimeter(ellipseItem);

        var correctionRatio = correctionPercent / 100;

        var baseSize;
        try {
            baseSize = sourceTextFrame.textRange.characterAttributes.size;
        } catch (e) {
            baseSize = 12;
        }

        var textWidth = measureTextWidth(text, baseSize);

        if (textWidth <= 0) {
            return baseSize;
        }

        var fittedSize = baseSize * ((perimeter * correctionRatio) / textWidth);

        return fittedSize;
    }

    /* プレビューを削除して元の表示状態へ戻す / Remove the preview and restore the original visibility */
    function removePreview() {
        try {
            if (previewTextFrame !== null) {
                previewTextFrame.remove();
                previewTextFrame = null;
            }
        } catch (e) {}

        sourceTextFrame.hidden = originalTextHidden;
        circlePath.hidden = originalPathHidden;

        app.redraw();
    }

    /* 選択中の連結文字とその設定を返す（trailing は中黒のときのみ真）/ Return the selected separator and whether it trails (only the bullet trails) */
    function getSeparatorInfo() {
        if (separatorDoubleSpaceRadio.value) {
            return { text: "  ", trailing: false };   /* 半角スペース x2 / two spaces */
        }
        if (separatorBulletRadio.value) {
            return { text: " • ", trailing: true };    /* 前後に半角スペース付きの中黒 / bullet padded with spaces */
        }
        return { text: " ", trailing: false };          /* 半角スペース / single space */
    }

    /* パス上文字を作成 / Create the path type */
    function createPathTypeText(repeatCount, correctionPercent, shouldFit, isPreview) {
        var separatorInfo = getSeparatorInfo();
        var repeatedText = makeRepeatedText(originalText, repeatCount, separatorInfo.text, separatorInfo.trailing);

        var duplicatedPath = circlePath.duplicate();

        var pathTypeFrame = activeDoc.textFrames.pathText(duplicatedPath);
        pathTypeFrame.contents = repeatedText;

        copyTextAttributes(sourceTextFrame, pathTypeFrame);
        copyParagraphAttributes(sourceTextFrame, pathTypeFrame);

        /* フィット ON のときのみ円周に合わせて文字サイズを調整 / Resize to fit only when fitting is enabled */
        if (shouldFit) {
            var fittedFontSize = calculateFitFontSize(repeatedText, circlePath, correctionPercent);
            pathTypeFrame.textRange.characterAttributes.size = fittedFontSize;
        }

        if (isPreview) {
            previewTextFrame = pathTypeFrame;

            sourceTextFrame.hidden = true;
            circlePath.hidden = true;
        }

        app.redraw();

        return pathTypeFrame;
    }

    /* 修飾キーに応じて次の値を計算（通常 ±1 / Shift ±10・10 の倍数にスナップ / Option ±0.1）/ Compute the next value by modifier (±1, Shift ±10 snapped, Option ±0.1) */
    function stepArrowValue(value, direction, keyboard) {
        if (keyboard.shiftKey) {
            /* Shift：±10 で 10 の倍数にスナップ / Shift: step ±10 snapped to a multiple of 10 */
            value = (direction > 0)
                ? Math.ceil((value + 1) / 10) * 10
                : Math.floor((value - 1) / 10) * 10;
            return Math.round(value);
        }
        if (keyboard.altKey) {
            /* Option：±0.1（小数第1位に丸め）/ Option: step ±0.1 rounded to one decimal */
            return Math.round((value + direction * 0.1) * 10) / 10;
        }
        /* 通常：±1（整数に丸め）/ Default: step ±1 rounded to an integer */
        return Math.round(value + direction);
    }

    /* 複数ラベルの幅を最大値に揃える / Align the widths of the given labels to the widest */
    function alignLabelWidths(labels) {
        var maxWidth = 0;
        for (var i = 0; i < labels.length; i++) {
            if (labels[i].preferredSize.width > maxWidth) {
                maxWidth = labels[i].preferredSize.width;
            }
        }
        for (var j = 0; j < labels.length; j++) {
            labels[j].preferredSize.width = maxWidth;
        }
    }

    /* ↑↓キーで値を増減 / Increment value with arrow keys */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var direction = 0;
            if (event.keyName === "Up") direction = 1;
            else if (event.keyName === "Down") direction = -1;
            else return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            event.preventDefault();

            var keyboard = ScriptUI.environment.keyboardState;
            value = stepArrowValue(value, direction, keyboard);

            /* 減算時のみ 0 で下げ止め（Option は小数微調整のため除外）/ Floor at 0 only when decreasing (Option excluded for fine decimals) */
            if (direction < 0 && !keyboard.altKey && value < 0) {
                value = 0;
            }

            editText.text = value;

            /* .text への代入では onChanging が発火しないため明示的に呼ぶ / Setting .text does not fire onChanging, so call it manually */
            if (typeof editText.onChanging === "function") {
                editText.onChanging();
            }
        });
    }

    // -----------------------------------------
    // ダイアログ / Dialog
    // -----------------------------------------

    /* タイトルバーにバージョンを表示 / Show the version in the title bar */
    var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);

    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var settingsPanel = dialog.add("panel", undefined, L("panel.settings"));
    settingsPanel.orientation = "column";
    settingsPanel.alignChildren = "left";
    settingsPanel.margins = 15;

    var repeatCountGroup = settingsPanel.add("group");
    repeatCountGroup.orientation = "row";
    var repeatCountLabel = repeatCountGroup.add("statictext", undefined, labelText("label.repeatCount"));

    var repeatCountInput = repeatCountGroup.add("edittext", undefined, String(DEFAULT_REPEAT_COUNT));
    repeatCountInput.characters = 6;

    var separatorGroup = settingsPanel.add("group");
    separatorGroup.orientation = "row";
    separatorGroup.alignChildren = "top";
    var separatorLabel = separatorGroup.add("statictext", undefined, labelText("label.separator"));

    var separatorRadioGroup = separatorGroup.add("group");
    separatorRadioGroup.orientation = "column";
    separatorRadioGroup.alignChildren = "left";

    var separatorSpaceRadio = separatorRadioGroup.add("radiobutton", undefined, L("radio.space1"));
    var separatorDoubleSpaceRadio = separatorRadioGroup.add("radiobutton", undefined, L("radio.space2"));
    var separatorBulletRadio = separatorRadioGroup.add("radiobutton", undefined, L("radio.bullet"));
    separatorSpaceRadio.value = true;

    /* 「繰り返し数」「連結文字」のラベル幅を揃える / Align the repeat-count and separator label widths */
    alignLabelWidths([repeatCountLabel, separatorLabel]);

    var fitPanel = dialog.add("panel", undefined, L("panel.fit"));
    fitPanel.orientation = "column";
    fitPanel.alignChildren = "left";
    fitPanel.margins = 15;

    var fitSizeCheckbox = fitPanel.add("checkbox", undefined, L("checkbox.fitSize"));
    fitSizeCheckbox.value = true;

    var correctionGroup = fitPanel.add("group");
    correctionGroup.orientation = "row";
    correctionGroup.add("statictext", undefined, labelText("label.correction"));

    var correctionInput = correctionGroup.add("edittext", undefined, String(DEFAULT_CORRECTION));
    correctionInput.characters = 6;

    var footerGroup = dialog.add("group");
    footerGroup.orientation = "row";
    footerGroup.alignment = "fill";

    /* 左：プレビュー / Left: preview */
    var previewArea = footerGroup.add("group");
    previewArea.alignment = ["left", "center"];

    var previewCheckbox = previewArea.add("checkbox", undefined, L("checkbox.preview"));
    previewCheckbox.value = true;

    /* 中央：spacer（余白を伸縮）/ Center: spacer that absorbs extra width */
    var spacerArea = footerGroup.add("group");
    spacerArea.alignment = ["fill", "center"];

    /* 右：ボタン2つ / Right: two buttons */
    var actionButtonGroup = footerGroup.add("group");
    actionButtonGroup.alignment = ["right", "center"];

    var cancelButton = actionButtonGroup.add("button", undefined, L("button.cancel"));
    var okButton = actionButtonGroup.add("button", undefined, "OK");

    /* 繰り返し数を取得（1 以上の整数のみ）/ Get the repeat count (integer >= 1 only) */
    function getRepeatCount() {
        var parsedCount = parseInt(repeatCountInput.text, 10);

        if (isNaN(parsedCount) || parsedCount < 1) {
            return null;
        }

        return parsedCount;
    }

    /* フィット補正率を取得（0 より大きい数値のみ）/ Get the fit correction (number > 0 only) */
    function getFitCorrection() {
        var parsedCorrection = parseFloat(correctionInput.text);

        if (isNaN(parsedCorrection) || parsedCorrection <= 0) {
            return null;
        }

        return parsedCorrection;
    }

    /* プレビューを更新 / Refresh the preview */
    function updatePreview() {
        removePreview();

        if (!previewCheckbox.value) {
            return;
        }

        var repeatCount = getRepeatCount();
        var correctionPercent = getFitCorrection();

        if (repeatCount === null || correctionPercent === null) {
            return;
        }

        createPathTypeText(repeatCount, correctionPercent, fitSizeCheckbox.value, true);
    }

    updatePreview();

    /* 入力中のプレビュー更新（再入防止）/ Refresh preview while typing (guarded against re-entry) */
    function onInputChanging() {
        if (isUpdatingPreview) return;

        isUpdatingPreview = true;
        updatePreview();
        isUpdatingPreview = false;
    }

    repeatCountInput.onChanging = onInputChanging;
    correctionInput.onChanging = onInputChanging;

    previewCheckbox.onClick = function () {
        updatePreview();
    };

    fitSizeCheckbox.onClick = function () {
        correctionInput.enabled = fitSizeCheckbox.value;
        updatePreview();
    };

    separatorSpaceRadio.onClick = updatePreview;
    separatorDoubleSpaceRadio.onClick = updatePreview;
    separatorBulletRadio.onClick = updatePreview;

    /* 初期状態を反映 / Apply the initial state */
    correctionInput.enabled = fitSizeCheckbox.value;

    changeValueByArrowKey(repeatCountInput);
    changeValueByArrowKey(correctionInput);

    okButton.onClick = function () {
        var repeatCount = getRepeatCount();
        var correctionPercent = getFitCorrection();

        if (repeatCount === null) {
            alert(L("alert.invalidCount"));
            return;
        }

        if (correctionPercent === null) {
            alert(L("alert.invalidCorrection"));
            return;
        }

        removePreview();

        var resultTextFrame = createPathTypeText(repeatCount, correctionPercent, fitSizeCheckbox.value, false);

        sourceTextFrame.remove();
        circlePath.remove();

        activeDoc.selection = null;
        resultTextFrame.selected = true;

        dialog.close();
    };

    cancelButton.onClick = function () {
        removePreview();

        activeDoc.selection = null;
        sourceTextFrame.selected = true;
        circlePath.selected = true;

        dialog.close();
    };

    dialog.onClose = function () {
        removePreview();
    };

    dialog.show();

})();
