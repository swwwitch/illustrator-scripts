#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 円（パス）とテキストを 1 つずつ選択し、テキストを指定回数繰り返して、円を複製したパス上の文字に変換する
- 連結文字は「スペース」または任意の入力文字（初期値は欧文 bullet「•」）を選択でき、末尾にも連結文字を付与
- 連結文字の前後に入れる半角スペース数を指定可能（スペースのみの場合は区切りの間隔になる）
- スペース以外の連結文字は、スケール（水平・垂直比率）とベースライン（環境設定のテキスト単位）を調整可能
- 円周に合わせた文字サイズの自動調整は ON/OFF 可能（補正率で開始・終了の隙間を微調整）
- 生成結果を円の中心基準で回転
- 数値フィールドは ↑↓ キーで増減（Shift で ±10・10 の倍数にスナップ、Option で ±0.1）
- プレビュー対応（確定までは元のテキスト・円を保持し、OK で元を削除して生成結果を選択）

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

/* 連結文字の初期値（欧文 bullet）/ Default separator character (bullet) */
var DEFAULT_SEPARATOR_CHAR = "•";

/* 連結スペース数の初期値（文字は左右にこの数ずつ）/ Default number of separator spaces (the character uses this many on each side) */
var DEFAULT_SPACE_COUNT = 1;

/* 連結文字（スペース以外）のスケール初期値（%）/ Default scale for the separator's non-space characters (%) */
var DEFAULT_SCALE = 100;

/* 連結文字（スペース以外）のベースライン初期値（環境設定のテキスト単位）/ Default baseline for the separator's non-space characters (in the preferences text unit) */
var DEFAULT_BASELINE = 0;

/* 補正率の初期値（%）/ Default correction (%) */
var DEFAULT_CORRECTION = 103;

/* 回転角度の初期値（度）/ Default rotation angle (degrees) */
var DEFAULT_ROTATION = 0;

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
        title: { ja: "円周文字の作成", en: "Create Circular Path Text" }
    },
    panel: {
        settings: { ja: "基本設定", en: "Basic Settings" },
        join: { ja: "連結文字", en: "Separator" },
        fit: { ja: "文字サイズ", en: "Font Size" }
    },
    label: {
        repeatCount: { ja: "繰り返し数", en: "Repeat count" },
        separator: { ja: "種類", en: "Type" },
        spaceCount: { ja: "前後のスペース", en: "Side spaces" },
        scale: { ja: "文字スケール", en: "Character scale" },
        baseline: { ja: "ベースラインシフト", en: "Baseline shift" },
        correction: { ja: "円周の補正", en: "Circumference correction" },
        rotation: { ja: "回転角度", en: "Rotation angle" }
    },
    radio: {
        space: { ja: "スペース", en: "Space" },
        character: { ja: "文字", en: "Character" }
    },
    checkbox: {
        preview: { ja: "プレビュー", en: "Preview" },
        fitSize: { ja: "円周に合わせて文字サイズを調整", en: "Adjust font size to circumference" }
    },
    button: {
        cancel: { ja: "キャンセル", en: "Cancel" }
    },
    tooltip: {
        repeatCount: {
            ja: "元のテキストを円周上に繰り返す回数です。",
            en: "Number of times to repeat the source text around the path."
        },
        rotation: {
            ja: "生成したパス上文字を、円の中心を基準に回転します。",
            en: "Rotates the generated path text around the circle center."
        },
        separatorSpace: {
            ja: "テキスト同士をスペースだけで区切ります。",
            en: "Separates repeated text with spaces only."
        },
        separatorCharacter: {
            ja: "入力した文字を連結文字として使用します。初期値は欧文 bullet です。",
            en: "Uses the typed character as the separator. The default is a bullet."
        },
        spaceCount: {
            ja: "連結文字の前後に入れる半角スペース数です。スペースのみの場合は区切りのスペース数になります。",
            en: "Number of half-width spaces on each side of the separator. For Space mode, this is the separator spacing."
        },
        scale: {
            ja: "スペース以外の連結文字にだけ適用する水平・垂直比率です。",
            en: "Horizontal and vertical scale applied only to non-space separator characters."
        },
        baseline: {
            ja: "スペース以外の連結文字にだけ適用するベースラインシフトです。単位はIllustratorの文字単位設定に従います。",
            en: "Baseline shift applied only to non-space separator characters. The unit follows Illustrator's text unit preference."
        },
        fitSize: {
            ja: "円周に収まるように、元テキストの文字サイズを基準に自動調整します。",
            en: "Automatically adjusts font size based on the source text so the repeated text fits the circumference."
        },
        correction: {
            ja: "文字サイズ自動調整時の隙間を微調整します。100%が基準です。",
            en: "Fine-tunes the gap when auto-adjusting font size. 100% is the baseline."
        },
        preview: {
            ja: "設定変更を一時的に反映して確認します。確定するまでは元のオブジェクトを保持します。",
            en: "Temporarily previews the result while keeping the original objects until you click OK."
        }
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
            ja: "補正率には0より大きい数値を入力してください。",
            en: "Enter a number greater than 0 for the correction."
        }
    }
};

/* ネストキー（例 "panel.fit"）からローカライズ文字列を取得 / Resolve a localized string from a dotted key (e.g. "panel.fit") */
function getLocalizedText(key) {
    var parts = key.split(".");
    var node = LABELS;
    for (var i = 0; i < parts.length; i++) {
        if (node == null) return key;
        node = node[parts[i]];
    }
    if (node == null) return key;
    return (node[currentLanguage] != null) ? node[currentLanguage] : node.en;
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLocalizedText(key) + (currentLanguage === "ja" ? "：" : ":");
}

/* tooltip を設定 / Set tooltip text */
function setTooltip(control, key) {
    control.helpTip = getLocalizedText(key);
}

// =========================================
// 単位 / Units
// =========================================

/* 単位コード→ラベルと 1 単位あたりの pt 数（環境設定の text/units 用）/ Unit code -> label and points-per-unit (for the preferences text/units) */
function getUnitInfo(unitCode) {
    /* 0=inch / 1=mm / 2=pt / 3=pica / 4=cm / 5=Q / 6=px */
    switch (unitCode) {
        case 0: return { label: "in", ptPerUnit: 72 };
        case 1: return { label: "mm", ptPerUnit: 72 / 25.4 };
        case 3: return { label: "pica", ptPerUnit: 12 };
        case 4: return { label: "cm", ptPerUnit: 72 / 2.54 };
        case 5: return { label: "Q", ptPerUnit: (72 / 25.4) * 0.25 };  /* 1Q = 0.25mm */
        case 6: return { label: "px", ptPerUnit: 1 };                    /* 72ppi 前提 / assumes 72ppi */
        default: return { label: "pt", ptPerUnit: 1 };                    /* 2=pt ほか / 2=pt and fallback */
    }
}

// =========================================
// レイアウト / Layout
// =========================================

/* パネルの余白と間隔 / Panel margins and spacing */
var PANEL_MARGINS = [16, 20, 16, 12];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Apply shared panel layout */
function setupColumnPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ["fill", "top"];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

(function () {

    /* ドキュメントの有無を確認 / Ensure a document is open */
    if (app.documents.length === 0) {
        alert(getLocalizedText("alert.noDocument"));
        return;
    }

    var activeDoc = app.activeDocument;

    /* 選択数を確認（円とテキストの2つ）/ Ensure exactly two objects are selected */
    if (activeDoc.selection.length !== 2) {
        alert(getLocalizedText("alert.selectTwo"));
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
        alert(getLocalizedText("alert.needTextAndPath"));
        return;
    }

    if (sourceTextFrame.contents === "") {
        alert(getLocalizedText("alert.emptyText"));
        return;
    }

    var originalText = sourceTextFrame.contents;

    var measureFrame = null;        /* 画面外に常駐させる計測用フレーム（使い回し）/ Reusable off-canvas measurement frame */
    var isUpdatingPreview = false;
    var committed = false;          /* OK で確定したか / Whether OK has committed the result */

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

        /* Ramanujan 近似式（ramanujanH は (a-b)^2/(a+b)^2）/ Ramanujan approximation (ramanujanH = (a-b)^2/(a+b)^2) */
        var ramanujanH = Math.pow(radiusX - radiusY, 2) / Math.pow(radiusX + radiusY, 2);
        var perimeter = Math.PI * (radiusX + radiusY) *
            (1 + (3 * ramanujanH) / (10 + Math.sqrt(4 - 3 * ramanujanH)));

        return perimeter;
    }

    /* 属性を1つだけ安全にコピー（1つ失敗しても他へ波及させない）/ Copy a single attribute safely (one failure won't affect the others) */
    function safeCopyAttribute(sourceAttributes, targetAttributes, attributeName) {
        try {
            targetAttributes[attributeName] = sourceAttributes[attributeName];
        } catch (e) { }
    }

    /* 文字属性をコピー（属性ごとに独立してコピー）/ Copy character attributes (each attribute copied independently) */
    function copyTextAttributes(sourceFrame, targetFrame) {
        var sourceAttributes = sourceFrame.textRange.characterAttributes;
        var targetAttributes = targetFrame.textRange.characterAttributes;
        var attributeNames = ["size", "textFont", "fillColor", "tracking",
            "horizontalScale", "verticalScale", "baselineShift"];
        for (var i = 0; i < attributeNames.length; i++) {
            safeCopyAttribute(sourceAttributes, targetAttributes, attributeNames[i]);
        }
    }

    /* 段落属性をコピー / Copy paragraph attributes */
    function copyParagraphAttributes(sourceFrame, targetFrame) {
        safeCopyAttribute(
            sourceFrame.textRange.paragraphAttributes,
            targetFrame.textRange.paragraphAttributes,
            "justification"
        );
    }

    /* 計測用フレームを1枚だけ用意（画面外に配置して使い回す）/ Ensure a single reusable measurement frame, placed off-canvas */
    function ensureMeasureFrame() {
        var valid = false;
        if (measureFrame !== null) {
            try {
                measureFrame.contents;   /* 参照できれば有効 / accessible means still valid */
                valid = true;
            } catch (e) {
                valid = false;
            }
        }
        if (!valid) {
            measureFrame = activeDoc.textFrames.add();
            measureFrame.top = -100000;
            measureFrame.left = -100000;
            /* 生成を確定し、プレビューの app.undo() で消えないようにする / Commit the creation so the preview's app.undo() cannot remove it */
            app.redraw();
        }
        return measureFrame;
    }

    /* 計測用フレームを削除 / Remove the measurement frame */
    function removeMeasureFrame() {
        try {
            if (measureFrame !== null) {
                measureFrame.remove();
            }
        } catch (e) { }
        measureFrame = null;
    }

    /* 使い回しフレームで文字幅を測る（毎回 add/remove しないので一時フレームが溜まらない）/ Measure text width with the reusable frame (no per-call add/remove, so temp frames cannot accumulate) */
    function measureTextWidth(text, fontSize) {
        var frame = ensureMeasureFrame();
        frame.contents = text;

        copyTextAttributes(sourceTextFrame, frame);
        frame.textRange.characterAttributes.size = fontSize;

        /* geometricBounds を正しく更新するため redraw が必要。常駐フレームなので毎回の add/remove は無く、蓄積しない。
           A redraw is required for geometricBounds to update correctly. The frame is reused (no per-call add/remove), so nothing accumulates. */
        app.redraw();

        var bounds = frame.geometricBounds;
        return Math.abs(bounds[2] - bounds[0]);
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

    /* 指定個数の半角スペース文字列を作る / Build a string of the given number of half-width spaces */
    function makeSpaces(count) {
        return new Array(count + 1).join(" ");
    }

    /* 選択中の連結文字とその設定を返す（scalableChars はスペース以外の文字／無ければ ""）/ Return the selected separator info (scalableChars are the non-space chars, or "") */
    function getSeparatorInfo() {
        var spaces = makeSpaces(getSpaceCount());
        if (separatorCharRadio.value) {
            /* 入力文字の左右にスペースを付与し、末尾にも繰り返す / Pad the typed character with spaces on both sides and repeat it at the end */
            var separatorChar = separatorCharInput.text;
            return { text: spaces + separatorChar + spaces, trailing: true, scalableChars: separatorChar };
        }
        /* スペースのみ（末尾にもスペースを付与し、円の折り返し位置の間隔も揃える）/ Spaces only (also trail spaces so the wrap-around gap matches) */
        return { text: spaces, trailing: true, scalableChars: "" };
    }

    /* 連結文字（スペース以外）に水平・垂直比率とベースラインを適用 / Apply H/V scale and baseline to the separator's non-space characters */
    function styleSeparatorChars(textFrame, targetChars, scalePercent, baselineShiftPt) {
        var characters = textFrame.textRange.characters;
        for (var i = 0; i < characters.length; i++) {
            var ch = characters[i].contents;
            if (ch !== " " && targetChars.indexOf(ch) !== -1) {
                characters[i].characterAttributes.horizontalScale = scalePercent;
                characters[i].characterAttributes.verticalScale = scalePercent;
                characters[i].characterAttributes.baselineShift = baselineShiftPt;
            }
        }
    }

    /* 指定点を中心にアイテムを回転 / Rotate an item around the given point */
    function rotateAroundCenter(item, angle, centerX, centerY) {
        var rotationMatrix = app.getRotationMatrix(angle);
        item.translate(-centerX, -centerY);
        item.transform(rotationMatrix);
        item.translate(centerX, centerY);
    }

    /* フィット文字サイズを計算（計測のため redraw を伴うので適用バッチの前に実行）/ Compute the fitted font size (involves a redraw, so run it before the apply batch) */
    function computeFitSize(repeatCount, correctionPercent, shouldFit) {
        if (!shouldFit) {
            return null;
        }
        var separatorInfo = getSeparatorInfo();
        var repeatedText = makeRepeatedText(originalText, repeatCount, separatorInfo.text, separatorInfo.trailing);
        return calculateFitFontSize(repeatedText, circlePath, correctionPercent);
    }

    /* パス上文字を作成（計測・redraw は含めない＝1 undo グループにするため）/ Create the path type (no measurement or redraw, so it forms a single undo group) */
    function createPathTypeText(repeatCount, fontSize, scale, baseline, rotation, isPreview) {
        var separatorInfo = getSeparatorInfo();
        var repeatedText = makeRepeatedText(originalText, repeatCount, separatorInfo.text, separatorInfo.trailing);

        var duplicatedPath = circlePath.duplicate();

        var pathTypeFrame = activeDoc.textFrames.pathText(duplicatedPath);
        pathTypeFrame.contents = repeatedText;

        copyTextAttributes(sourceTextFrame, pathTypeFrame);
        copyParagraphAttributes(sourceTextFrame, pathTypeFrame);

        /* 事前計算したフィットサイズを適用（null はフィット OFF）/ Apply the precomputed fit size (null means fitting is off) */
        if (fontSize !== null) {
            pathTypeFrame.textRange.characterAttributes.size = fontSize;
        }

        /* 連結文字（スペース以外）のスケール・ベースラインを適用 / Apply scale and baseline to the separator's non-space characters */
        if (separatorInfo.scalableChars !== "" && (scale !== 100 || baseline !== 0)) {
            styleSeparatorChars(pathTypeFrame, separatorInfo.scalableChars, scale, baseline);
        }

        /* 円の中心を基準に回転 / Rotate around the circle center */
        if (rotation !== 0) {
            var bounds = circlePath.geometricBounds;
            var centerX = (bounds[0] + bounds[2]) / 2;
            var centerY = (bounds[1] + bounds[3]) / 2;
            rotateAroundCenter(pathTypeFrame, rotation, centerX, centerY);
        }

        /* プレビュー時は元のテキスト・円を一時的に隠す（undo で復帰）/ Hide the originals during preview (restored by undo) */
        if (isPreview) {
            sourceTextFrame.hidden = true;
            circlePath.hidden = true;
        }

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

    /* ↑↓キーで値を増減（allowNegative が真なら負値も許可）/ Increment value with arrow keys (allowNegative permits values below 0) */
    function changeValueByArrowKey(editText, allowNegative) {
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

            /* 減算時のみ 0 で下げ止め（Option と allowNegative は除外）/ Floor at 0 only when decreasing (Option and allowNegative excluded) */
            if (!allowNegative && direction < 0 && !keyboard.altKey && value < 0) {
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

    /* 環境設定のテキスト単位（ベースライン入力に使用）/ Preferences text unit (used by the baseline field) */
    var textUnitInfo = getUnitInfo(app.preferences.getIntegerPreference("text/units"));

    /* タイトルバーにバージョンを表示 / Show the version in the title bar */
    var dialog = new Window("dialog", getLocalizedText("dialog.title") + " " + SCRIPT_VERSION);

    dialog.orientation = "column";
    dialog.alignChildren = "fill";

    var settingsPanel = dialog.add("panel", undefined, getLocalizedText("panel.settings"));
    setupColumnPanel(settingsPanel);

    var repeatCountGroup = settingsPanel.add("group");
    repeatCountGroup.orientation = "row";
    var repeatCountLabel = repeatCountGroup.add("statictext", undefined, labelText("label.repeatCount"));

    var repeatCountInput = repeatCountGroup.add("edittext", undefined, String(DEFAULT_REPEAT_COUNT));
    repeatCountInput.characters = 3;
    setTooltip(repeatCountInput, "tooltip.repeatCount");

    var rotationGroup = settingsPanel.add("group");
    rotationGroup.orientation = "row";
    var rotationLabel = rotationGroup.add("statictext", undefined, labelText("label.rotation"));

    var rotationInput = rotationGroup.add("edittext", undefined, String(DEFAULT_ROTATION));
    rotationInput.characters = 3;
    setTooltip(rotationInput, "tooltip.rotation");

    rotationGroup.add("statictext", undefined, "°");

    var joinPanel = dialog.add("panel", undefined, getLocalizedText("panel.join"));
    setupColumnPanel(joinPanel);

    var separatorGroup = joinPanel.add("group");
    separatorGroup.orientation = "row";
    separatorGroup.alignChildren = "top";
    var separatorLabel = separatorGroup.add("statictext", undefined, labelText("label.separator"));

    var separatorRadioGroup = separatorGroup.add("group");
    separatorRadioGroup.orientation = "column";
    separatorRadioGroup.alignChildren = "left";

    var separatorSpaceRadio = separatorRadioGroup.add("radiobutton", undefined, getLocalizedText("radio.space"));
    setTooltip(separatorSpaceRadio, "tooltip.separatorSpace");

    /* 「文字」ラジオ＋自由入力フィールド（初期値は欧文 bullet）/ "Character" radio plus a free-input field (defaults to the bullet) */
    var separatorCharRow = separatorRadioGroup.add("group");
    separatorCharRow.orientation = "row";
    var separatorCharRadio = separatorCharRow.add("radiobutton", undefined, getLocalizedText("radio.character"));
    setTooltip(separatorCharRadio, "tooltip.separatorCharacter");
    var separatorCharInput = separatorCharRow.add("edittext", undefined, DEFAULT_SEPARATOR_CHAR);
    separatorCharInput.characters = 3;
    setTooltip(separatorCharInput, "tooltip.separatorCharacter");

    /* スペースと文字は親グループが異なるため排他制御は手動 / Space and character radios live in different parents, so exclusivity is handled manually */
    separatorSpaceRadio.value = false;
    separatorCharRadio.value = true;

    var spaceCountGroup = joinPanel.add("group");
    spaceCountGroup.orientation = "row";
    var spaceCountLabel = spaceCountGroup.add("statictext", undefined, labelText("label.spaceCount"));

    var spaceCountInput = spaceCountGroup.add("edittext", undefined, String(DEFAULT_SPACE_COUNT));
    spaceCountInput.characters = 6;
    setTooltip(spaceCountInput, "tooltip.spaceCount");

    var scaleGroup = joinPanel.add("group");
    scaleGroup.orientation = "row";
    var scaleLabel = scaleGroup.add("statictext", undefined, labelText("label.scale"));

    var scaleInput = scaleGroup.add("edittext", undefined, String(DEFAULT_SCALE));
    scaleInput.characters = 6;
    setTooltip(scaleInput, "tooltip.scale");

    scaleGroup.add("statictext", undefined, "%");

    var baselineGroup = joinPanel.add("group");
    baselineGroup.orientation = "row";
    var baselineLabel = baselineGroup.add("statictext", undefined, labelText("label.baseline"));

    var baselineInput = baselineGroup.add("edittext", undefined, String(DEFAULT_BASELINE));
    baselineInput.characters = 6;
    setTooltip(baselineInput, "tooltip.baseline");

    /* 単位ラベルは環境設定のテキスト単位を表示 / The unit label shows the preferences text unit */
    baselineGroup.add("statictext", undefined, textUnitInfo.label);

    var fitPanel = dialog.add("panel", undefined, getLocalizedText("panel.fit"));
    setupColumnPanel(fitPanel);

    var fitSizeCheckbox = fitPanel.add("checkbox", undefined, getLocalizedText("checkbox.fitSize"));
    fitSizeCheckbox.value = true;
    setTooltip(fitSizeCheckbox, "tooltip.fitSize");

    var correctionGroup = fitPanel.add("group");
    correctionGroup.orientation = "row";
    var correctionLabel = correctionGroup.add("statictext", undefined, labelText("label.correction"));

    var correctionInput = correctionGroup.add("edittext", undefined, String(DEFAULT_CORRECTION));
    correctionInput.characters = 6;
    setTooltip(correctionInput, "tooltip.correction");

    correctionGroup.add("statictext", undefined, "%");

    /* ダイアログ内すべてのテキストフィールド前ラベルを同一幅に揃える / Align all field labels across the dialog to one width */
    alignLabelWidths([repeatCountLabel, rotationLabel, separatorLabel, spaceCountLabel, scaleLabel, baselineLabel, correctionLabel]);

    var footerGroup = dialog.add("group");
    footerGroup.orientation = "row";
    footerGroup.alignment = "fill";

    /* 左：プレビュー / Left: preview */
    var previewArea = footerGroup.add("group");
    previewArea.alignment = ["left", "center"];

    var previewCheckbox = previewArea.add("checkbox", undefined, getLocalizedText("checkbox.preview"));
    previewCheckbox.value = true;
    setTooltip(previewCheckbox, "tooltip.preview");

    /* 中央：spacer（余白を伸縮）/ Center: spacer that absorbs extra width */
    var spacerArea = footerGroup.add("group");
    spacerArea.alignment = ["fill", "center"];

    /* 右：ボタン2つ / Right: two buttons */
    var actionButtonGroup = footerGroup.add("group");
    actionButtonGroup.alignment = ["right", "center"];

    var cancelButton = actionButtonGroup.add("button", undefined, getLocalizedText("button.cancel"));
    var okButton = actionButtonGroup.add("button", undefined, "OK");

    /* 繰り返し数を取得（1 以上の整数のみ）/ Get the repeat count (integer >= 1 only) */
    function getRepeatCount() {
        var parsedCount = parseInt(repeatCountInput.text, 10);

        if (isNaN(parsedCount) || parsedCount < 1) {
            return null;
        }

        return parsedCount;
    }

    /* 連結スペース数を取得（1 以上の整数、不正時は 1）/ Get the separator space count (integer >= 1, falls back to 1) */
    function getSpaceCount() {
        var parsedSpaces = parseInt(spaceCountInput.text, 10);

        if (isNaN(parsedSpaces) || parsedSpaces < 1) {
            return 1;
        }

        return parsedSpaces;
    }

    /* 連結文字スケールを取得（0 より大きい数値、不正時は 100）/ Get the separator scale (number > 0, falls back to 100) */
    function getScale() {
        var parsedScale = parseFloat(scaleInput.text);

        if (isNaN(parsedScale) || parsedScale <= 0) {
            return 100;
        }

        return parsedScale;
    }

    /* ベースラインシフトを pt で取得（入力は環境設定のテキスト単位、負値・小数可）/ Get the baseline shift in points (input is in the preferences text unit; negatives and decimals allowed) */
    function getBaselineShiftPt() {
        var parsedBaseline = parseFloat(baselineInput.text);

        if (isNaN(parsedBaseline)) {
            return 0;
        }

        return parsedBaseline * textUnitInfo.ptPerUnit;
    }

    /* 補正率を取得（0 より大きい数値のみ）/ Get the correction (number > 0 only) */
    function getCorrectionPercent() {
        var parsedCorrection = parseFloat(correctionInput.text);

        if (isNaN(parsedCorrection) || parsedCorrection <= 0) {
            return null;
        }

        return parsedCorrection;
    }

    /* 回転角度を取得（負値・小数可、不正時は 0）/ Get the rotation angle (negatives and decimals allowed, falls back to 0) */
    function getRotation() {
        var parsedRotation = parseFloat(rotationInput.text);

        if (isNaN(parsedRotation)) {
            return 0;
        }

        return parsedRotation;
    }

    /* プレビュー：適用 → redraw（見せる）→ undo（内部を未適用へ戻す）/ Preview: apply -> redraw (show) -> undo (revert the model) */
    function runPreview() {
        if (!previewCheckbox.value) {
            app.redraw();   /* 未適用状態を表示 / show the clean state */
            return;
        }

        var repeatCount = getRepeatCount();
        if (repeatCount === null) {
            app.redraw();
            return;
        }

        /* 補正率はフィット ON のときだけ必須 / The correction is required only when fitting is on */
        var correctionPercent = getCorrectionPercent();
        if (fitSizeCheckbox.value && correctionPercent === null) {
            app.redraw();
            return;
        }

        /* 計測は redraw を含むため適用バッチの前に実行 / Measure before the apply batch (it involves a redraw) */
        var fontSize = computeFitSize(repeatCount, correctionPercent, fitSizeCheckbox.value);

        /* 仮アイテムで強制的に変化を起こし、undo の空振りを防ぐ。画面外に作り、変数にも保持しない（undo で消える）/ Force a change with an off-canvas dummy so undo cannot misfire (not kept in a variable since undo removes it) */
        activeDoc.pathItems.rectangle(-100000, -100000, 1, 1);

        createPathTypeText(repeatCount, fontSize, getScale(), getBaselineShiftPt(), getRotation(), true);

        app.redraw();   /* 見せる / show the applied result */
        app.undo();     /* 内部を未適用へ戻す（画面は適用後のまま）/ revert the model (screen keeps showing it) */
    }

    /* 入力中のプレビュー更新（再入防止）/ Refresh preview while typing (guarded against re-entry) */
    function onInputChanging() {
        if (isUpdatingPreview) return;

        isUpdatingPreview = true;
        runPreview();
        isUpdatingPreview = false;
    }

    repeatCountInput.onChanging = onInputChanging;
    separatorCharInput.onChanging = onInputChanging;
    spaceCountInput.onChanging = onInputChanging;
    scaleInput.onChanging = onInputChanging;
    baselineInput.onChanging = onInputChanging;
    correctionInput.onChanging = onInputChanging;
    rotationInput.onChanging = onInputChanging;

    previewCheckbox.onClick = function () {
        runPreview();
    };

    fitSizeCheckbox.onClick = function () {
        correctionInput.enabled = fitSizeCheckbox.value;
        runPreview();
    };

    /* 連結文字の変更：スケール・ベースラインは「文字」のときだけ有効 / Separator change: scale and baseline apply only to the character option */
    function onSeparatorChange() {
        var charSelected = separatorCharRadio.value;
        separatorCharInput.enabled = charSelected;
        scaleInput.enabled = charSelected;
        baselineInput.enabled = charSelected;
        runPreview();
    }

    /* クリックした側を選択し、もう一方を解除（手動排他）/ Select the clicked radio and clear the other (manual exclusivity) */
    separatorSpaceRadio.onClick = function () {
        separatorSpaceRadio.value = true;
        separatorCharRadio.value = false;
        onSeparatorChange();
    };
    separatorCharRadio.onClick = function () {
        separatorCharRadio.value = true;
        separatorSpaceRadio.value = false;
        onSeparatorChange();
    };

    /* 初期状態を反映 / Apply the initial state */
    correctionInput.enabled = fitSizeCheckbox.value;
    separatorCharInput.enabled = separatorCharRadio.value;
    scaleInput.enabled = separatorCharRadio.value;
    baselineInput.enabled = separatorCharRadio.value;

    changeValueByArrowKey(repeatCountInput);
    changeValueByArrowKey(spaceCountInput);
    changeValueByArrowKey(scaleInput);
    changeValueByArrowKey(baselineInput, true);
    changeValueByArrowKey(correctionInput);
    changeValueByArrowKey(rotationInput, true);

    okButton.onClick = function () {
        var repeatCount = getRepeatCount();
        var correctionPercent = getCorrectionPercent();

        if (repeatCount === null) {
            alert(getLocalizedText("alert.invalidCount"));
            return;
        }

        /* 補正率はフィット ON のときだけ必須 / The correction is required only when fitting is on */
        if (fitSizeCheckbox.value && correctionPercent === null) {
            alert(getLocalizedText("alert.invalidCorrection"));
            return;
        }

        /* 確定：未適用のクリーンな状態から 1 回だけ適用（計測→適用は redraw を挟まず一塊）/ Commit: apply once from the clean state (measure then apply with no redraw in between) */
        var fontSize = computeFitSize(repeatCount, correctionPercent, fitSizeCheckbox.value);
        removeMeasureFrame();   /* 計測フレームを片付けてから確定 / Clean up the measurement frame before committing */

        var resultTextFrame = createPathTypeText(repeatCount, fontSize, getScale(), getBaselineShiftPt(), getRotation(), false);

        sourceTextFrame.remove();
        circlePath.remove();

        activeDoc.selection = null;
        resultTextFrame.selected = true;

        committed = true;
        app.redraw();
        dialog.close();
    };

    cancelButton.onClick = function () {
        /* 内部は undo 済み。計測フレームを片付け、元の選択へ戻して閉じる / Model already reverted; clean up the measurement frame, restore the selection, and close */
        removeMeasureFrame();
        activeDoc.selection = null;
        sourceTextFrame.selected = true;
        circlePath.selected = true;
        app.redraw();
        dialog.close();
    };

    dialog.onClose = function () {
        /* 計測フレームを片付け、未確定なら未適用状態を画面に反映 / Clean up the measurement frame; if not committed, refresh the screen to the reverted state */
        removeMeasureFrame();
        if (!committed) {
            app.redraw();
        }
    };

    /* 初期プレビューはウィンドウ表示後に起動（同期実行を避ける）/ Start the initial preview after the window is shown (avoid running it synchronously) */
    dialog.onShow = function () {
        runPreview();
    };

    dialog.show();

})();
