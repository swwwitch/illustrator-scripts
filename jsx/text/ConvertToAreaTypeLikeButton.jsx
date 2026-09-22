#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形・エリア内文字を対象に、エリア内文字の作成と調整を行います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToAreaTypeLikeButton.md

### Overview

Creates and adjusts area text from point text, text on a path, shapes, or existing area text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToAreaTypeLikeButton.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ConvertToAreaTypeLikeButton";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToAreaTypeLikeButton.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToAreaTypeLikeButton.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ボタン風の拡大倍率 / Button-style expansion ratios */
    var BUTTON_WIDTH_RATIO  = 1.2;  /* 元の幅に対する倍率 / Ratio of original width */
    var BUTTON_HEIGHT_RATIO = 1.6;  /* 元の高さに対する倍率 / Ratio of original height */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS        = 20;                /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS         = [16, 20, 16, 12];  /* パネルの余白 / panel margins */
    var PANEL_SPACING         = 8;                 /* パネル内の間隔 / panel spacing */
    var SIZE_LABEL_WIDTH      = 44;                /* 幅・高さのラベル幅 / width of the width/height labels */
    var INDENT_CHECKBOX_WIDTH = 52;                /* 左右インデントのチェックボックス幅 / width of the indent checkboxes */

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - パネル内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループの共通設定を適用する（row/column で整列を切り替え）
     * @param {Group} targetGroup - 対象のグループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - グループ内の間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(targetGroup, orientation, spacing) {
        var groupOrientation = orientation || "column";
        targetGroup.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        targetGroup.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        targetGroup.alignment = "fill";
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義（カテゴリ構造）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        /* ダイアログ / Dialog */
        dialog: {
            title: { ja: "AreaType Toolkit", en: "AreaType Toolkit" }
        },
        /* パネル見出し / Panel titles */
        panel: {
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            frameSize: { ja: "フレームサイズ", en: "Frame size" },
            indent: { ja: "インデント", en: "Indent" },
            options: { ja: "オプション", en: "Options" }
        },
        /* 入力ラベル / Field labels */
        fieldLabel: {
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" }
        },
        /* チェックボックス / Checkboxes */
        checkbox: {
            indentLeft: { ja: "左", en: "Left" },
            indentRight: { ja: "右", en: "Right" },
            sync: { ja: "連動", en: "Link" },
            margin: { ja: "外側からの間隔", en: "Spacing" }
        },
        /* ボタン / Buttons */
        button: {
            overset: { ja: "文字あふれ解消", en: "Make overset" },
            run: { ja: "実行", en: "Run" },
            close: { ja: "閉じる", en: "Close" }
        },
        /* ツールチップ / Tooltips */
        tooltip: {
            fontSize: { ja: "エリア内文字のフォントサイズです。", en: "Font size of the area text." },
            overset: { ja: "文字があふれないところまでフォントサイズを下げます。", en: "Lowers the font size until the text no longer overflows." },
            width: { ja: "テキストフレームの幅です。", en: "Width of the text frame." },
            height: { ja: "テキストフレームの高さです。", en: "Height of the text frame." },
            indentLeft: { ja: "段落の左インデントを設定します。", en: "Sets the left indent of the paragraphs." },
            indentRight: { ja: "段落の右インデントを設定します。", en: "Sets the right indent of the paragraphs." },
            indentValue: { ja: "インデントの量です。", en: "Amount of the indent." },
            sync: { ja: "左右のインデントを同じ値にします。", en: "Keeps the left and right indents the same." },
            margin: { ja: "テキストフレームの内側に空ける余白です。", en: "Inset kept inside the text frame." }
        },
        /* 警告メッセージ / Alerts */
        alert: {
            selectText: {
                ja: "ポイント文字・パス上文字・エリア内文字を選択してください。",
                en: "Please select point text, path text, or area text."
            },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * "category.key" 形式のキーからラベルを取得する
     * @param {string} labelPath - ラベルキー（例: "panel.fontSize"）
     * @returns {string} 現在の言語のラベル文字列（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            if (labelNode && typeof labelNode[labelPathKeys[i]] !== "undefined") { labelNode = labelNode[labelPathKeys[i]]; }
            else { return labelPath; }
        }
        if (labelNode[uiLang]) return labelNode[uiLang];
        if (labelNode.en) return labelNode.en;
        return labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelSet - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 選択の判定 / Selection checks
    // =========================================

    /**
     * パス上文字かを返す
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} パス上文字なら true
     */
    function isPathTextFrame(pageItem) {
        return pageItem.typename === "TextFrame" && pageItem.kind === TextType.PATHTEXT;
    }

    /**
     * エリア内文字かを返す
     * @param {PageItem} pageItem - 調べるオブジェクト
     * @returns {boolean} エリア内文字なら true
     */
    function isAreaTextFrame(pageItem) {
        return pageItem.typename === "TextFrame" && pageItem.kind === TextType.AREATEXT;
    }

    /**
     * 選択にポイント文字（パス上文字を含む）・エリア内文字・パスが含まれるかを調べる
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {{hasPointText: boolean, hasAreaText: boolean, hasPathItem: boolean}} 含まれる種類
     */
    function classifySelection(selectedItems) {
        var selectionKinds = { hasPointText: false, hasAreaText: false, hasPathItem: false };
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (selectedItem.typename === "TextFrame") {
                if (selectedItem.kind === TextType.POINTTEXT || selectedItem.kind === TextType.PATHTEXT) selectionKinds.hasPointText = true;
                if (selectedItem.kind === TextType.AREATEXT) selectionKinds.hasAreaText = true;
            }
            if (selectedItem.typename === "PathItem" || selectedItem.typename === "CompoundPathItem") {
                selectionKinds.hasPathItem = true;
            }
        }
        return selectionKinds;
    }

    /**
     * 選択から最初のエリア内文字を返す
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {TextFrame|null} 最初のエリア内文字（無ければ null）
     */
    function findFirstAreaText(selectedItems) {
        for (var i = 0; i < selectedItems.length; i++) {
            if (isAreaTextFrame(selectedItems[i])) return selectedItems[i];
        }
        return null;
    }

    // =========================================
    // パス上文字 → ポイント文字（変換前処理）/ Path text → Point text (pre-process)
    // =========================================

    /**
     * 関数を実行し、例外は握りつぶす（属性ごとに失敗しても残りを続けるため）
     * @param {Function} attemptAction - 実行する処理
     * @returns {*} 処理の戻り値（例外時は undefined）
     */
    function runIgnoringErrors(attemptAction) {
        try { return attemptAction(); } catch (e) { return undefined; }
    }

    /**
     * 文字ごとの属性を控える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Object[]} 文字ごとの属性
     */
    function snapshotCharacterAttributes(textFrame) {
        var attributeSnapshots = [];
        for (var i = 0; i < textFrame.characters.length; i++) {
            var sourceAttributes = textFrame.characters[i].characterAttributes;
            attributeSnapshots.push({
                font: sourceAttributes.textFont,
                size: sourceAttributes.size,
                fillColor: sourceAttributes.fillColor,
                strokeColor: sourceAttributes.strokeColor,
                strokeWeight: sourceAttributes.strokeWeight,
                autoLeading: sourceAttributes.autoLeading,
                leading: sourceAttributes.leading
            });
        }
        return attributeSnapshots;
    }

    /**
     * 控えた文字属性を書き戻す（ベースライン移動と比率はリセット）。属性ごとに失敗しても続ける
     * @param {TextFrame} textFrame - 書き戻し先のテキストフレーム
     * @param {Object[]} attributeSnapshots - snapshotCharacterAttributes() の戻り値
     * @returns {void}
     */
    function restoreCharacterAttributes(textFrame, attributeSnapshots) {
        var restoreCount = Math.min(textFrame.characters.length, attributeSnapshots.length);
        for (var i = 0; i < restoreCount; i++) {
            var targetAttributes = textFrame.characters[i].characterAttributes;
            var savedAttributes = attributeSnapshots[i];

            runIgnoringErrors(function () { targetAttributes.textFont = savedAttributes.font; });
            runIgnoringErrors(function () { targetAttributes.size = savedAttributes.size; });
            runIgnoringErrors(function () { targetAttributes.fillColor = savedAttributes.fillColor; });
            runIgnoringErrors(function () {
                var savedStrokeColor = savedAttributes.strokeColor;
                targetAttributes.strokeColor = savedStrokeColor;
                targetAttributes.strokeWeight = (savedStrokeColor && savedStrokeColor.typename === "NoColor") ? 0 : savedAttributes.strokeWeight;
            });
            runIgnoringErrors(function () { targetAttributes.baselineShift = 0; });
            runIgnoringErrors(function () { targetAttributes.horizontalScale = 100; });
            runIgnoringErrors(function () { targetAttributes.verticalScale = 100; });
            runIgnoringErrors(function () { targetAttributes.autoLeading = savedAttributes.autoLeading; });
            if (!savedAttributes.autoLeading) runIgnoringErrors(function () { targetAttributes.leading = savedAttributes.leading; });
        }
    }

    /**
     * パス上文字を字形を保ったままポイント文字へ分離する
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame[]} pathTextFrames - 分離するパス上文字
     * @returns {TextFrame[]} 作成したポイント文字
     */
    function detachPathTextToPointText(doc, pathTextFrames) {
        var createdPointTexts = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) return createdPointTexts;

        /* 新規テキストだけ選べるよう選択を解除 / Clear selection */
        runIgnoringErrors(function () { doc.selection = null; });

        for (var j = pathTextFrames.length - 1; j >= 0; j--) {
            var pathText = pathTextFrames[j];
            if (!pathText || !isPathTextFrame(pathText)) continue;

            var originalPath = null;
            runIgnoringErrors(function () { originalPath = pathText.textPath; });
            if (!originalPath) continue;

            /* 1) 文字ごとの属性を退避 / Snapshot per-character attributes */
            var attributeSnapshots = snapshotCharacterAttributes(pathText);

            var textContents = "";
            runIgnoringErrors(function () { textContents = pathText.contents; });

            var justification = null;
            runIgnoringErrors(function () {
                if (pathText.paragraphs && pathText.paragraphs.length > 0) {
                    justification = pathText.paragraphs[0].paragraphAttributes.justification;
                }
            });

            /* 2) パス始点にポイント文字を新規作成 / Create new point text at path start anchor */
            var pointText = doc.textFrames.add();
            var anchorPoint = null;
            runIgnoringErrors(function () {
                if (originalPath.pathPoints && originalPath.pathPoints.length > 0) {
                    anchorPoint = originalPath.pathPoints[0].anchor;
                }
            });
            if (anchorPoint) {
                pointText.position = [anchorPoint[0], anchorPoint[1]];
            }

            pointText.contents = textContents;

            if (justification !== null && pointText.paragraphs && pointText.paragraphs.length > 0) {
                runIgnoringErrors(function () { pointText.paragraphs[0].paragraphAttributes.justification = justification; });
            }

            /* 既定の線を一旦消し、後で文字ごとに復元 / Clear default stroke, restore per-character later */
            runIgnoringErrors(function () {
                pointText.textRange.characterAttributes.strokeColor = new NoColor();
                pointText.textRange.characterAttributes.strokeWeight = 0;
            });

            /* 文字ごとの属性を復元 / Restore per-character attributes */
            restoreCharacterAttributes(pointText, attributeSnapshots);

            /* 3) 元のパス上文字を削除（パスも一緒に消える）/ Remove original path text */
            runIgnoringErrors(function () { pathText.remove(); });

            /* 4) 新規テキストを選択して返す / Select and return new text */
            runIgnoringErrors(function () { pointText.selected = true; });
            createdPointTexts.push(pointText);
        }

        return createdPointTexts;
    }

    /**
     * 選択内のパス上文字をポイント文字へ置き換え、置き換えた選択配列を返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @returns {PageItem[]} 置き換え後の選択（パス上文字が無ければ元の選択）
     */
    function preprocessPathTextSelection(doc, currentSelection) {
        if (!doc || !currentSelection || !currentSelection.length) return currentSelection;

        var pathTexts = [];
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            /* 無効オブジェクト（削除済み等）は読めない / Invalid (deleted) objects cannot be read */
            try {
                if (selectedItem && isPathTextFrame(selectedItem)) pathTexts.push(selectedItem);
            } catch (e0) { }
        }
        if (!pathTexts.length) return currentSelection;

        var createdPointTexts = detachPathTextToPointText(doc, pathTexts);
        if (!createdPointTexts.length) return currentSelection;

        /* パス上文字を新ポイント文字に差し替えた新しい選択配列を構築 / Build replaced selection array */
        var replacedSelection = [];
        for (var j = 0; j < currentSelection.length; j++) {
            var remainingItem = currentSelection[j];
            /* 削除したパス上文字は読めないので飛ばす / Removed path text cannot be read, so it is skipped */
            try {
                if (remainingItem && !isPathTextFrame(remainingItem)) replacedSelection.push(remainingItem);
            } catch (e1) { }
        }
        for (var k = 0; k < createdPointTexts.length; k++) replacedSelection.push(createdPointTexts[k]);

        try { doc.selection = replacedSelection; } catch (e) { }
        app.redraw();

        return replacedSelection;
    }

    // =========================================
    // オーバーセット判定・文字サイズ調整 / Overset detection & font sizing
    // =========================================

    /**
     * overflows プロパティを安全に取得する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean|null} あふれていれば true（読めなければ null）
     */
    function readOverflows(textFrame) {
        try {
            if (textFrame && typeof textFrame.overflows !== "undefined") return !!textFrame.overflows;
        } catch (e) { }
        return null;
    }

    /**
     * 表示されている行に収まらない文字があるか判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} あふれているとき true
     */
    function hasHiddenCharacters(textFrame) {
        var lineCount = textFrame.lines.length;
        if (lineCount === 0) return textFrame.characters.length > 0;
        var visibleCharacters = 0;
        for (var i = 0; i < lineCount; i++) { visibleCharacters += textFrame.lines[i].characters.length; }
        return visibleCharacters < textFrame.characters.length;
    }

    /**
     * テキストフレームがあふれているか判定する（エリア内文字は overflows を優先）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} あふれているとき true
     */
    function isOversetFrame(textFrame) {
        if (textFrame && textFrame.kind === TextType.AREATEXT) {
            var overflowState = readOverflows(textFrame);
            if (overflowState !== null) return overflowState;
        }
        try { return hasHiddenCharacters(textFrame); } catch (e) { return false; }
    }

    /**
     * 行送り比率（行送り/サイズ）を取得する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {{ratio: number}|null} 行送りの比率（自動行送りなどのときは null）
     */
    function getLeadingInfo(textFrame) {
        try {
            var textAttributes = textFrame.textRange.characterAttributes;
            if (textAttributes.autoLeading) return null;
            var fontSize = textAttributes.size, leading = textAttributes.leading;
            if (fontSize > 0 && leading > 0) return { ratio: leading / fontSize };
        } catch (e) { }
        return null;
    }

    /**
     * 比率を保ったまま行送りを更新する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} newSize - 変更後の文字サイズ（pt）
     * @param {{ratio: number}|null} leadingInfo - getLeadingInfo() の戻り値
     * @returns {void}
     */
    function applyLeading(textFrame, newSize, leadingInfo) {
        if (!leadingInfo) return;
        try { textFrame.textRange.characterAttributes.leading = newSize * leadingInfo.ratio; } catch (e) { }
    }

    /**
     * 文字サイズを設定し、行送りを比率に合わせる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fontSize - 文字サイズ（pt）
     * @param {{ratio: number}|null} leadingInfo - getLeadingInfo() の戻り値
     * @returns {void}
     */
    function setFontSizeKeepingLeading(textFrame, fontSize, leadingInfo) {
        textFrame.textRange.characterAttributes.size = fontSize;
        applyLeading(textFrame, fontSize, leadingInfo);
    }

    /**
     * あふれなくなる最大サイズを二分探索で探して縮小する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function shrinkFont(textFrame) {
        if (textFrame.characters.length <= 0 || !isOversetFrame(textFrame)) return;
        var leadingInfo = getLeadingInfo(textFrame);
        var upperSize = textFrame.textRange.characterAttributes.size;
        var lowerSize = 0.1;

        /* 最小でもあふれるならそのまま終了 / If even lowerSize overflows, keep min size */
        setFontSizeKeepingLeading(textFrame, lowerSize, leadingInfo);
        if (isOversetFrame(textFrame)) return;

        /* 二分探索 / Binary search */
        for (var i = 0; i < 40; i++) {
            var middleSize = (lowerSize + upperSize) / 2;
            setFontSizeKeepingLeading(textFrame, middleSize, leadingInfo);
            if (isOversetFrame(textFrame)) {
                upperSize = middleSize;
            } else {
                lowerSize = middleSize;
            }
            if (upperSize - lowerSize < 0.1) break;
        }

        /* あふれない側（lowerSize）に確定 / Settle on the non-overset side */
        setFontSizeKeepingLeading(textFrame, lowerSize, leadingInfo);
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    // =========================================

    /**
     * 文字列を16進数表現へ変換する
     * @param {string} text - 変換する文字列
     * @returns {string} 16進文字列
     */
    function asciiToHex(text) {
        var hexText = "";
        for (var i = 0; i < text.length; i++) {
            var hexPair = text.charCodeAt(i).toString(16);
            if (hexPair.length < 2) hexPair = "0" + hexPair;
            hexText += hexPair;
        }
        return hexText;
    }

    /**
     * /name ブロック（ASCII）を生成する
     * @param {string} actionName - アクション名またはセット名
     * @returns {string} 名前ブロックの文字列
     */
    function buildActionNameBlock(actionName) {
        return "/name [ " + actionName.length + " " + asciiToHex(actionName).toUpperCase() + " ]";
    }

    /**
     * アクションセット定義（.aia）文字列を組み立てる
     * @param {string} setName - アクションセット名
     * @param {string} internalName - イベントの内部名
     * @param {string} localizedNameHex - ローカライズ名（長さと16進。空なら省略）
     * @param {number} parameterKey - パラメーターのキー
     * @param {Array<{name: string, value: number}>} actionDefinitions - アクション名と値の組
     * @returns {string} .aia 形式のアクションセット定義
     */
    function buildActionSetAia(setName, internalName, localizedNameHex, parameterKey, actionDefinitions) {
        var aiaText = "/version 3" +
            buildActionNameBlock(setName) +
            "/isOpen 1" +
            "/actionCount " + actionDefinitions.length;

        for (var i = 0; i < actionDefinitions.length; i++) {
            var actionDef = actionDefinitions[i];
            aiaText += "/action-" + (i + 1) + " {" +
                " " + buildActionNameBlock(actionDef.name) +
                " /keyIndex 0" +
                " /colorIndex 0" +
                " /isOpen 1" +
                " /eventCount 1" +
                " /event-1 {" +
                " /useRulersIn1stQuadrant 0" +
                " /internalName (" + internalName + ")" +
                (localizedNameHex ? (" /localizedName [ " + localizedNameHex + " ]") : "") +
                " /isOpen 0" +
                " /isOn 1" +
                " /hasDialog 0" +
                " /parameterCount 1" +
                " /parameter-1 {" +
                " /key " + parameterKey +
                " /showInPalette 4294967295" +
                " /type (integer)" +
                " /value " + actionDef.value +
                " }" +
                " }" +
                "}";
        }
        return aiaText;
    }

    /* フレーム整列アクションセット名 / Frame-alignment action set name */
    var AREA_TEXT_ACTION_SET = "AreaText";

    /* フレーム整列のアクション名（添字がアクションの値：0=上, 1=中央, 2=下, 3=均等）
       Frame-alignment action names; the index is the action value (0=top, 1=center, 2=bottom, 3=justify) */
    var FRAME_ALIGNMENT_ACTIONS = ["AlignTop", "AlignCenter", "AlignBottom", "AlignJustify"];

    /**
     * アクションセットを読み込む（temp に .aia を書き出して loadAction。既存があれば先に外す）
     * @param {string} setName - アクションセット名
     * @param {string} aiaString - .aia 形式のアクションセット定義
     * @returns {void}
     */
    function loadActionSet(setName, aiaString) {
        unloadActionSet(setName);
        /* 一時ファイルの書き出し・読み込みは失敗しうる / Writing the temp file or loading it can fail */
        try {
            var actionFile = new File(Folder.temp + "/AreaTypeToolkit_action_" + setName + ".aia");
            actionFile.open("w");
            actionFile.write(aiaString);
            actionFile.close();

            app.loadAction(actionFile);
            actionFile.remove();
        } catch (e) { }
    }

    /**
     * アクションセットを破棄する
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadActionSet(setName) {
        /* 読み込まれていなければ例外になる / Throws when the set is not loaded */
        try { app.unloadAction(setName, ""); } catch (e) { }
    }

    /**
     * フレーム整列アクション（AlignTop/Center/Bottom/Justify）を読み込む
     * @returns {void}
     */
    function loadAreaTextActions() {
        var actionDefinitions = [];
        for (var i = 0; i < FRAME_ALIGNMENT_ACTIONS.length; i++) {
            actionDefinitions.push({ name: FRAME_ALIGNMENT_ACTIONS[i], value: i });
        }
        var aiaText = buildActionSetAia(
            AREA_TEXT_ACTION_SET,
            "adobe_frameAlignment",
            "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
            1717660782,
            actionDefinitions
        );
        loadActionSet(AREA_TEXT_ACTION_SET, aiaText);
    }

    /**
     * フレーム整列アクションを破棄する
     * @returns {void}
     */
    function unloadAreaTextActions() {
        unloadActionSet(AREA_TEXT_ACTION_SET);
    }

    /**
     * テキストの配置（垂直方向のフレーム整列）をアクションで変更する
     * @param {number} alignValue - 0=上 / 1=中央 / 2=下 / 3=均等
     * @returns {void}
     */
    function runFrameAlignmentAction(alignValue) {
        if (alignValue !== 0 && alignValue !== 1 && alignValue !== 2 && alignValue !== 3) return;
        try { app.doScript(FRAME_ALIGNMENT_ACTIONS[alignValue], AREA_TEXT_ACTION_SET, false); } catch (e) { }
    }

    /**
     * 指定フレームにフレーム整列を適用する（プレビュー中はスキップ）
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @param {number} alignValue - 0=上 / 1=中央 / 2=下 / 3=均等
     * @param {boolean} forPreview - プレビュー中か
     * @returns {void}
     */
    function applyAreaTextFrameAlignment(areaTextFrame, alignValue, forPreview) {
        /* app.doScript はプレビュー中に呼ぶと不安定なためスキップ / Unstable during preview, so skip */
        if (forPreview) return;
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [areaTextFrame];
            app.redraw(); /* 選択状態を確定 / Commit the selection */
            runFrameAlignmentAction(alignValue);
        } catch (e) { }
    }

    // =========================================
    // UIユーティリティ / UI utilities
    // =========================================

    /**
     * ↑↓キーで値を増減する（Shift=10 / Alt=0.1）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すか
     * @param {Function} [onChangeCallback] - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onChangeCallback) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboardState = ScriptUI.environment.keyboardState;

            if (keyboardState.shiftKey) {
                var delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else if (keyboardState.altKey) {
                if (event.keyName == "Up") {
                    value += 0.1;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= 0.1;
                    event.preventDefault();
                }
            } else {
                if (event.keyName == "Up") {
                    value += 1;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= 1;
                    event.preventDefault();
                }
            }

            value = keyboardState.altKey ? Math.round(value * 10) / 10 : Math.round(value);
            if (!allowNegative && value < 0) value = 0;

            editText.text = value;
            if (typeof onChangeCallback === "function") { onChangeCallback(); }
        });
    }

    /**
     * 幅・高さの行（ラベル＋入力欄＋単位）を追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} labelKey - LABELS.fieldLabel と LABELS.tooltip のキー
     * @param {string} unitLabel - 単位の表示
     * @returns {EditText} 追加した入力欄
     */
    function addSizeRow(parentContainer, labelKey, unitLabel) {
        var sizeRowGroup = parentContainer.add("group");
        var sizeLabel = sizeRowGroup.add("statictext", undefined, labelText("fieldLabel." + labelKey));
        sizeLabel.preferredSize.width = SIZE_LABEL_WIDTH;
        var sizeInput = sizeRowGroup.add("edittext", undefined, "");
        sizeInput.characters = 5;
        sizeInput.helpTip = getLabel("tooltip." + labelKey);
        sizeRowGroup.add("statictext", undefined, unitLabel);
        return sizeInput;
    }

    /**
     * インデントの行（チェックボックス＋入力欄＋単位）を追加する（入力欄は無効で始まる）
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} labelKey - LABELS.checkbox と LABELS.tooltip のキー
     * @param {string} unitLabel - 単位の表示
     * @returns {{checkbox: Checkbox, input: EditText}} 追加したチェックボックスと入力欄
     */
    function addIndentRow(parentContainer, labelKey, unitLabel) {
        var indentRowGroup = parentContainer.add("group");
        var indentCheckbox = indentRowGroup.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        indentCheckbox.preferredSize.width = INDENT_CHECKBOX_WIDTH;
        indentCheckbox.helpTip = getLabel("tooltip." + labelKey);
        var indentInput = indentRowGroup.add("edittext", undefined, "0");
        indentInput.characters = 4;
        indentInput.helpTip = getLabel("tooltip.indentValue");
        indentRowGroup.add("statictext", undefined, unitLabel);
        indentInput.enabled = false;
        return { checkbox: indentCheckbox, input: indentInput };
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /**
     * テキストのフォントとサイズを読む
     * @param {TextFrame} textFrame - 読み取り元
     * @returns {{font: TextFont|null, size: number}} フォントとサイズ（読めなければ null / 0）
     */
    function readFontAndSize(textFrame) {
        var fontStyle = { font: null, size: 0 };
        try {
            fontStyle.font = textFrame.textRange.characterAttributes.textFont;
            fontStyle.size = textFrame.textRange.characterAttributes.size;
        } catch (e) { }
        return fontStyle;
    }

    /**
     * フォントとサイズを適用する（未インストールのフォントなどは失敗しうる）
     * @param {TextFrame} textFrame - 適用先
     * @param {{font: TextFont|null, size: number}} fontStyle - readFontAndSize() の戻り値
     * @returns {void}
     */
    function applyFontAndSize(textFrame, fontStyle) {
        try {
            if (fontStyle.font) textFrame.textRange.characterAttributes.textFont = fontStyle.font;
            if (fontStyle.size > 0) textFrame.textRange.characterAttributes.size = fontStyle.size;
        } catch (e) { }
    }

    /**
     * 選択からモードを判定して変換し、調整ダイアログを開く
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} currentSelection - 現在の選択
     * @returns {void}
     */
    function convertToAreaTypeAndAdjust(doc, currentSelection) {
        preprocessPathTextSelection(doc, currentSelection);

        var updatedSelection = doc.selection;
        if (!updatedSelection || updatedSelection.length === 0) { return; }

        /* 選択内容に応じてモードを決定（ポイント文字は常にボタン風）/ Decide mode (point text is always button style) */
        var selectionKinds = classifySelection(updatedSelection);
        if (!selectionKinds.hasPointText) return;

        var createdFrames = selectionKinds.hasPathItem
            ? createAreaTextFromTextAndShape(doc, updatedSelection)
            : createButtonAreaTexts(doc, updatedSelection);
        if (createdFrames.length > 0) {
            doc.selection = createdFrames;
            app.redraw();
            showAdjustDialog(doc, createdFrames[0], createdFrames);
        }
    }

    /**
     * 選択したテキストと閉じたパスから、エリア内文字を1つ作る
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {TextFrame[]} 作成したエリア内文字（作れなければ空）
     */
    function createAreaTextFromTextAndShape(doc, selectedItems) {
        var createdFrames = [];
        var sourceText = null, destPath = null;
        for (var i = 0; i < selectedItems.length; i++) {
            if (!sourceText && selectedItems[i].typename === "TextFrame") { sourceText = selectedItems[i]; }
            else if (!destPath && selectedItems[i].typename === "PathItem" && selectedItems[i].closed) { destPath = selectedItems[i]; }
        }
        if (!sourceText || !destPath) return createdFrames;

        /* 種類によってはエリア内文字にできない / Some shapes cannot become area type */
        try {
            var sourceContents = sourceText.contents;
            var sourceStyle = readFontAndSize(sourceText);
            var framePath = destPath.duplicate();
            framePath.filled = false; framePath.stroked = false;
            var areaTextFrame = doc.textFrames.areaText(framePath);
            areaTextFrame.contents = sourceContents;
            applyFontAndSize(areaTextFrame, sourceStyle);
            sourceText.remove();
            destPath.remove();
            createdFrames.push(areaTextFrame);
        } catch (e) { }
        return createdFrames;
    }

    /**
     * ポイント文字ごとに、幅×BUTTON_WIDTH_RATIO・高さ×BUTTON_HEIGHT_RATIO の枠で中央揃えのエリア内文字を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {TextFrame[]} 作成したエリア内文字
     */
    function createButtonAreaTexts(doc, selectedItems) {
        var createdFrames = [];
        for (var i = selectedItems.length - 1; i >= 0; i--) {
            var pointText = selectedItems[i];
            if (pointText.typename !== "TextFrame" || pointText.kind !== TextType.POINTTEXT) continue;
            /* 1件で失敗しても残りを処理する / One failure must not stop the rest */
            try {
                var textBounds = pointText.geometricBounds;
                var originalWidth = textBounds[2] - textBounds[0], originalHeight = textBounds[1] - textBounds[3];
                var buttonWidth = originalWidth * BUTTON_WIDTH_RATIO, buttonHeight = originalHeight * BUTTON_HEIGHT_RATIO;
                var buttonRect = doc.pathItems.rectangle(
                    textBounds[1] + (buttonHeight - originalHeight) / 2, textBounds[0] - (buttonWidth - originalWidth) / 2, buttonWidth, buttonHeight);
                buttonRect.filled = false; buttonRect.stroked = false;
                var textContents = pointText.contents;
                var sourceStyle = readFontAndSize(pointText);
                var areaTextFrame = doc.textFrames.areaText(buttonRect);
                areaTextFrame.contents = textContents;
                applyFontAndSize(areaTextFrame, sourceStyle);
                /* 行揃えを中央に / Horizontal center */
                try { areaTextFrame.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { }
                /* 縦位置はDOMで不安定なためアクションで中央 / Vertical center via dynamic action (unreliable via DOM) */
                applyAreaTextFrameAlignment(areaTextFrame, 1, false);
                createdFrames.push(areaTextFrame);
                pointText.remove();
            } catch (e) { }
        }
        return createdFrames;
    }

    // =========================================
    // 調整ダイアログ / Adjust dialog
    // =========================================

    /**
     * 幅・高さの入力を検証して有効値を返す（不正なら最終正常値へ戻す）
     * @param {EditText} editText - 対象の入力欄
     * @param {number|null} lastValue - 最終正常値（ルーラー単位）
     * @param {number} pointsPerUnit - ルーラー単位1あたりのポイント数
     * @returns {number|null} 有効値（ルーラー単位。不正なら null）
     */
    function validateSizeField(editText, lastValue, pointsPerUnit) {
        var sizeValue = parseFloat(String(editText.text));
        if (isNaN(sizeValue) || !isFinite(sizeValue) || sizeValue <= 0) {
            if (lastValue !== null) editText.text = lastValue;
            return null;
        }
        /* 表示単位での上限値（極端値防止）/ Max size in ruler units (guards extreme values) */
        var maxSize = 100000 / pointsPerUnit;
        if (sizeValue > maxSize) {
            sizeValue = maxSize;
            editText.text = Math.round(sizeValue * 100) / 100;
        }
        if (sizeValue < 0.01) {
            sizeValue = 0.01;
            editText.text = Math.round(sizeValue * 100) / 100;
        }
        return sizeValue;
    }

    /**
     * 選択からエリア内文字だけを集める
     * @param {Document} doc - 対象ドキュメント
     * @returns {TextFrame[]|null} エリア内文字（無ければ・読めなければ null）
     */
    function collectSelectedAreaTexts(doc) {
        /* undo 後などは選択が読めないことがある / The selection may be unreadable, e.g. after an undo */
        try {
            var currentSelection = doc.selection;
            var areaTextFrames = [];
            if (currentSelection && currentSelection.length) {
                for (var i = 0; i < currentSelection.length; i++) {
                    if (currentSelection[i] && isAreaTextFrame(currentSelection[i])) {
                        areaTextFrames.push(currentSelection[i]);
                    }
                }
            }
            return areaTextFrames.length ? areaTextFrames : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 調整ダイアログを組み立てる（イベントは showAdjustDialog() で結び付ける）
     * @param {string} unitLabel - ルーラー単位の表示
     * @returns {Object} ダイアログと各コントロール
     */
    function buildAdjustDialog(unitLabel) {
        var adjustDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        adjustDialog.alignChildren = "fill";
        adjustDialog.margins = DIALOG_MARGINS;

        var mainColumnGroup = adjustDialog.add("group");
        setupGroup(mainColumnGroup, "column");

        /* フォントサイズ / Font size */
        var fontSizePanel = mainColumnGroup.add("panel", undefined, getLabel("panel.fontSize"));
        setupPanel(fontSizePanel);
        var fontSizeRow = fontSizePanel.add("group");
        fontSizeRow.alignment = "left";
        fontSizeRow.add("statictext", undefined, labelText("fieldLabel.fontSize"));
        var fontSizeInput = fontSizeRow.add("edittext", undefined, "");
        fontSizeInput.characters = 5;
        fontSizeInput.helpTip = getLabel("tooltip.fontSize");
        fontSizeRow.add("statictext", undefined, "pt");
        var oversetButtonRow = fontSizePanel.add("group");
        oversetButtonRow.orientation = "row";
        var btnFixOverset = oversetButtonRow.add("button", undefined, getLabel("button.overset"));
        btnFixOverset.helpTip = getLabel("tooltip.overset");

        /* フレームサイズ / Frame size */
        var frameSizePanel = mainColumnGroup.add("panel", undefined, getLabel("panel.frameSize"));
        setupPanel(frameSizePanel);
        var widthInput = addSizeRow(frameSizePanel, "width", unitLabel);
        var heightInput = addSizeRow(frameSizePanel, "height", unitLabel);

        /* インデント / Indent */
        var indentPanel = mainColumnGroup.add("panel", undefined, getLabel("panel.indent"));
        setupPanel(indentPanel);
        /* 連動チェックを右側に並べるため行方向へ上書き / Override to row so "Link" sits to the right */
        indentPanel.orientation = "row";
        indentPanel.alignChildren = ["left", "top"];
        var indentFieldsColumn = indentPanel.add("group");
        indentFieldsColumn.orientation = "column";
        indentFieldsColumn.alignChildren = "left";
        var leftIndentRow = addIndentRow(indentFieldsColumn, "indentLeft", unitLabel);
        var rightIndentRow = addIndentRow(indentFieldsColumn, "indentRight", unitLabel);
        var syncColumn = indentPanel.add("group");
        syncColumn.orientation = "column";
        syncColumn.alignChildren = "left";
        syncColumn.alignment = ["left", "center"];
        var syncCheckbox = syncColumn.add("checkbox", undefined, getLabel("checkbox.sync"));
        syncCheckbox.helpTip = getLabel("tooltip.sync");

        /* オプション / Options */
        var optionsPanel = mainColumnGroup.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel);
        var marginRow = optionsPanel.add("group");
        var marginCheckbox = marginRow.add("checkbox", undefined, getLabel("checkbox.margin"));
        marginCheckbox.helpTip = getLabel("tooltip.margin");
        var marginInput = marginRow.add("edittext", undefined, "0");
        marginInput.characters = 6;
        marginInput.helpTip = getLabel("tooltip.margin");
        var marginUnitLabel = marginRow.add("statictext", undefined, unitLabel);
        marginInput.enabled = false;
        marginUnitLabel.enabled = false;

        /* ボタンエリア / Button area */
        var btnRowGroup = adjustDialog.add("group");
        setupGroup(btnRowGroup, "row");
        btnRowGroup.alignChildren = ["right", "center"];
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        var btnClose = btnRightGroup.add("button", undefined, getLabel("button.close"), { name: "cancel" });
        var btnRun = btnRightGroup.add("button", undefined, getLabel("button.run"), { name: "ok" });

        return {
            adjustDialog: adjustDialog,
            fontSizeInput: fontSizeInput,
            btnFixOverset: btnFixOverset,
            widthInput: widthInput,
            heightInput: heightInput,
            leftIndentCheckbox: leftIndentRow.checkbox,
            leftIndentInput: leftIndentRow.input,
            rightIndentCheckbox: rightIndentRow.checkbox,
            rightIndentInput: rightIndentRow.input,
            syncCheckbox: syncCheckbox,
            marginCheckbox: marginCheckbox,
            marginInput: marginInput,
            marginUnitLabel: marginUnitLabel,
            btnClose: btnClose,
            btnRun: btnRun
        };
    }

    /**
     * エリア内文字の調整ダイアログを表示する（プレビューは常時ON）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame|null} initialFrame - 初期値を読むエリア内文字
     * @param {TextFrame[]|null} targetFrames - 調整対象（null なら現在の選択）
     * @returns {void}
     */
    function showAdjustDialog(doc, initialFrame, targetFrames) {
        var rulerUnit = getUnitInfo("rulerType");
        var pointsPerUnit = rulerUnit.pointsPerUnit;

        /* 受け取った変換結果を確実に対象にする / Make the passed frames the active target */
        var framesToSelect = (targetFrames && targetFrames.length) ? targetFrames : (initialFrame ? [initialFrame] : null);
        if (framesToSelect) {
            try { doc.selection = framesToSelect; } catch (e) { }
        }
        app.redraw();

        /* モーダル中は selection が変動するため、渡された配列を優先して固定 / Pin targets (selection drifts in modal) */
        var fixedTargets = null;
        if (targetFrames && targetFrames.length) {
            fixedTargets = targetFrames.slice(0);
        } else {
            /* 選択が読めないことがある / The selection may be unreadable */
            try {
                var initialSelection = doc.selection;
                if (initialSelection && initialSelection.length) {
                    fixedTargets = [];
                    for (var i = 0; i < initialSelection.length; i++) { fixedTargets.push(initialSelection[i]); }
                }
            } catch (e4) { }
        }

        var dialogControls = buildAdjustDialog(rulerUnit.label);
        var adjustDialog = dialogControls.adjustDialog;
        var fontSizeInput = dialogControls.fontSizeInput;
        var widthInput = dialogControls.widthInput;
        var heightInput = dialogControls.heightInput;
        var leftIndentCheckbox = dialogControls.leftIndentCheckbox;
        var leftIndentInput = dialogControls.leftIndentInput;
        var rightIndentCheckbox = dialogControls.rightIndentCheckbox;
        var rightIndentInput = dialogControls.rightIndentInput;
        var syncCheckbox = dialogControls.syncCheckbox;
        var marginCheckbox = dialogControls.marginCheckbox;
        var marginInput = dialogControls.marginInput;
        var marginUnitLabel = dialogControls.marginUnitLabel;

        /* 状態変数 / State */
        var isPreviewActive = false;

        /* 入力バリデーション用の最終正常値（ルーラー単位）/ Last valid values for validation (ruler units) */
        var lastValidWidth = null;
        var lastValidHeight = null;

        /**
         * ポイント値をルーラー単位に換算して小数第2位で丸める
         * @param {number} valueInPt - ポイント値
         * @returns {number} ルーラー単位の値
         */
        function toRoundedRulerUnits(valueInPt) {
            return Math.round((valueInPt / pointsPerUnit) * 100) / 100;
        }

        /**
         * 対象フレームから現在値をUIに読み込む
         * @param {TextFrame} sourceFrame - 読み取り元のエリア内文字
         * @returns {void}
         */
        function loadValuesFromFrame(sourceFrame) {
            var hasMultiParagraph;
            try { hasMultiParagraph = (sourceFrame.paragraphs && sourceFrame.paragraphs.length >= 2); } catch (e) { hasMultiParagraph = false; }
            var initialWidth = toRoundedRulerUnits(sourceFrame.textPath.width);
            var initialHeight = toRoundedRulerUnits(sourceFrame.textPath.height);
            var fontSize = 0;
            try { fontSize = sourceFrame.textRange.characterAttributes.size || 0; } catch (e) { }
            if (fontSize > 0) { fontSizeInput.text = Math.round(fontSize * 100) / 100; }
            widthInput.text = initialWidth;
            heightInput.text = initialHeight;
            lastValidWidth = parseFloat(widthInput.text);
            lastValidHeight = parseFloat(heightInput.text);
            try {
                var frameSpacing = sourceFrame.spacing || 0;
                marginInput.text = toRoundedRulerUnits(frameSpacing);
                marginCheckbox.value = (frameSpacing !== 0);
                marginInput.enabled = marginCheckbox.value;
                marginUnitLabel.enabled = marginCheckbox.value;
            } catch (e) { }
            try {
                var leftIndentPt = sourceFrame.paragraphs.length > 0 ? (sourceFrame.paragraphs[0].leftIndent || 0) : 0;
                var rightIndentPt = sourceFrame.paragraphs.length > 0 ? (sourceFrame.paragraphs[0].rightIndent || 0) : 0;
                leftIndentCheckbox.value = (leftIndentPt !== 0);
                leftIndentInput.enabled = leftIndentCheckbox.value;
                leftIndentInput.text = leftIndentCheckbox.value ? toRoundedRulerUnits(leftIndentPt) : "0";
                rightIndentCheckbox.value = (rightIndentPt !== 0);
                rightIndentInput.enabled = rightIndentCheckbox.value;
                rightIndentInput.text = rightIndentCheckbox.value ? toRoundedRulerUnits(rightIndentPt) : "0";
            } catch (e) { }
            dialogControls.btnFixOverset.enabled = !hasMultiParagraph;
        }

        /**
         * UIの値をフレームへ適用する（行揃え・配置は常に中央）
         * @param {boolean} forPreview - プレビューとして適用するか（フレーム整列のアクションは実行しない）
         * @param {boolean} [shrinkToFit] - あふれないところまでフォントサイズを下げるか
         * @returns {void}
         */
        function runAdjust(forPreview, shrinkToFit) {
            var justification = Justification.CENTER;
            var frameAlignment = 1; /* 中央 / center */

            var leftIndentPt = (leftIndentCheckbox.value || syncCheckbox.value)
                ? (parseFloat(leftIndentInput.text) || 0) * pointsPerUnit : 0;
            var rightIndentPt = syncCheckbox.value
                ? leftIndentPt
                : (rightIndentCheckbox.value ? (parseFloat(rightIndentInput.text) || 0) * pointsPerUnit : 0);
            var marginPt = marginCheckbox.value ? (parseFloat(marginInput.text) || 0) * pointsPerUnit : 0;

            var savedSelection = [];
            var originalSelection = doc.selection;
            for (var j = 0; j < originalSelection.length; j++) { savedSelection.push(originalSelection[j]); }

            /* 固定ターゲットを優先 / Prefer pinned targets */
            if (!(fixedTargets && fixedTargets.length)) { fixedTargets = collectSelectedAreaTexts(doc); }
            var adjustTargets = fixedTargets && fixedTargets.length ? fixedTargets : doc.selection;

            for (var i = adjustTargets.length - 1; i >= 0; i--) {
                var targetFrame = adjustTargets[i];
                if (!isAreaTextFrame(targetFrame)) continue;
                try { targetFrame.spacing = marginPt; } catch (e) { }
                /* 幅/高さ：NaN・0以下・極端値をガード / Guard NaN, non-positive, extreme values */
                var widthInRulerUnits = validateSizeField(widthInput, lastValidWidth, pointsPerUnit);
                var heightInRulerUnits = validateSizeField(heightInput, lastValidHeight, pointsPerUnit);
                if (widthInRulerUnits !== null) { lastValidWidth = widthInRulerUnits; try { targetFrame.textPath.width = widthInRulerUnits * pointsPerUnit; } catch (e) { } }
                if (heightInRulerUnits !== null) { lastValidHeight = heightInRulerUnits; try { targetFrame.textPath.height = heightInRulerUnits * pointsPerUnit; } catch (e) { } }
                if (shrinkToFit) { shrinkFont(targetFrame); }
                try {
                    var paragraphAttrs = targetFrame.textRange.paragraphAttributes;
                    paragraphAttrs.justification = justification;
                    paragraphAttrs.leftIndent = leftIndentPt;
                    paragraphAttrs.rightIndent = rightIndentPt;
                } catch (e) { }
                applyAreaTextFrameAlignment(targetFrame, frameAlignment, forPreview);
            }
            if (savedSelection.length > 0) { try { doc.selection = savedSelection; } catch (e) { } }
            app.redraw();
        }

        /**
         * 直前のプレビューを undo で取り消す
         * @returns {void}
         */
        function undoPreview() {
            try { app.undo(); } catch (e) { }
            app.redraw();
            isPreviewActive = false;
        }

        /**
         * プレビューを更新する（常時ON。直前のプレビューは undo で戻す）
         * @returns {void}
         */
        function updatePreview() {
            if (isPreviewActive) {
                undoPreview();
                /* undo 後は参照が無効化されるため対象を取り直す / Refresh targets after undo */
                fixedTargets = collectSelectedAreaTexts(doc);
            }
            if (!(fixedTargets && fixedTargets.length)) { fixedTargets = collectSelectedAreaTexts(doc); }
            runAdjust(true);
            isPreviewActive = true;
        }

        /**
         * フォントサイズ入力を選択中のエリア内文字へ即時反映する
         * @returns {void}
         */
        function applyFontSizeFromField() {
            var newSize = parseFloat(fontSizeInput.text) || 0;
            if (newSize <= 0) return;
            var selectedItems = doc.selection;
            for (var i = 0; i < selectedItems.length; i++) {
                if (isAreaTextFrame(selectedItems[i])) {
                    try { selectedItems[i].textRange.characterAttributes.size = newSize; } catch (e) { }
                }
            }
            app.redraw();
        }

        /**
         * 幅入力を検証してプレビューする
         * @returns {void}
         */
        function onWidthChange() {
            var widthInRulerUnits = validateSizeField(widthInput, lastValidWidth, pointsPerUnit);
            if (widthInRulerUnits !== null) { lastValidWidth = widthInRulerUnits; }
            updatePreview();
        }

        /**
         * 高さ入力を検証してプレビューする
         * @returns {void}
         */
        function onHeightChange() {
            var heightInRulerUnits = validateSizeField(heightInput, lastValidHeight, pointsPerUnit);
            if (heightInRulerUnits !== null) { lastValidHeight = heightInRulerUnits; }
            updatePreview();
        }

        /**
         * 連動時に右インデントを同期してから幅変更扱いにする
         * @returns {void}
         */
        function onAdjustmentChange() {
            if (syncCheckbox.value) { rightIndentInput.text = leftIndentInput.text; }
            onWidthChange();
        }

        /* --- イベントハンドラ / Event handlers --- */
        dialogControls.btnFixOverset.onClick = function () { runAdjust(false, true); };
        marginCheckbox.onClick = function () {
            marginInput.enabled = marginCheckbox.value;
            marginUnitLabel.enabled = marginCheckbox.value;
            if (marginCheckbox.value) { marginInput.text = "1"; }
            onAdjustmentChange();
        };
        syncCheckbox.onClick = function () {
            if (syncCheckbox.value) {
                leftIndentCheckbox.value = true; leftIndentInput.enabled = true;
                rightIndentCheckbox.enabled = false; rightIndentInput.enabled = false;
                rightIndentInput.text = leftIndentInput.text;
            } else {
                rightIndentCheckbox.enabled = true; rightIndentInput.enabled = rightIndentCheckbox.value;
            }
            onAdjustmentChange();
        };
        leftIndentCheckbox.onClick = function () {
            if (!syncCheckbox.value) { leftIndentInput.enabled = leftIndentCheckbox.value; }
            if (!leftIndentCheckbox.value) { leftIndentInput.text = "0"; }
            onAdjustmentChange();
        };
        rightIndentCheckbox.onClick = function () {
            rightIndentInput.enabled = rightIndentCheckbox.value;
            if (!rightIndentCheckbox.value) { rightIndentInput.text = "0"; }
            onAdjustmentChange();
        };

        fontSizeInput.onChange = applyFontSizeFromField;
        marginInput.onChange = onAdjustmentChange;
        widthInput.onChange = onWidthChange;
        heightInput.onChange = onHeightChange;
        leftIndentInput.onChange = onAdjustmentChange;
        rightIndentInput.onChange = onAdjustmentChange;
        changeValueByArrowKey(fontSizeInput, false, applyFontSizeFromField);
        changeValueByArrowKey(marginInput, false, onAdjustmentChange);
        changeValueByArrowKey(widthInput, false, onWidthChange);
        changeValueByArrowKey(heightInput, false, updatePreview);
        changeValueByArrowKey(leftIndentInput, false, onAdjustmentChange);
        changeValueByArrowKey(rightIndentInput, false, onAdjustmentChange);

        dialogControls.btnRun.onClick = function () {
            if (isPreviewActive) { undoPreview(); }
            runAdjust(false);
            adjustDialog.close(1);
        };
        dialogControls.btnClose.onClick = function () {
            if (isPreviewActive) { app.undo(); app.redraw(); }
            adjustDialog.close(0);
        };

        /* 初期値読み込み / Load initial values */
        if (initialFrame) { loadValuesFromFrame(initialFrame); }

        /* 開いたらプレビュー実行（常時ON）/ Run preview on open (always on) */
        updatePreview();

        adjustDialog.show();
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    /**
     * 選択を確かめ、変換または調整ダイアログを実行する（アクションは終了時に必ず破棄）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0) {
            alert(getLabel("alert.selectText"));
            return;
        }

        var selectionKinds = classifySelection(currentSelection);
        if (!selectionKinds.hasPointText && !selectionKinds.hasPathItem && !selectionKinds.hasAreaText) {
            alert(getLabel("alert.selectText"));
            return;
        }

        /* アクションを実行時に読み込み、終了時に破棄 / Load actions at start, unload on exit */
        loadAreaTextActions();
        try {
            if (selectionKinds.hasPointText || selectionKinds.hasPathItem) {
                /* ポイント文字 または 図形 → 変換（常にボタン風）してダイアログ / Point or shape → convert then dialog */
                convertToAreaTypeAndAdjust(doc, currentSelection);
            } else {
                /* エリア内文字のみ → 調整ダイアログ / Area type only → adjust dialog */
                showAdjustDialog(doc, findFirstAreaText(currentSelection), null);
            }
        } finally {
            unloadAreaTextActions();
        }
    }

    main();

})();
