#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した見出しの左右に二重線を引き、まわりに段組みの新聞ダミーを敷いて、新聞の見出しのような画像を作ります。
新聞ダミーはぼかし、全体を［変形］効果で傾けて、欠けない長方形でマスクします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShimbunTitleMaker.md

### Overview

Draws double rules on both sides of the selected headline and lays newspaper-style dummy columns around it.
The dummy text is blurred, the whole piece is tilted with the Transform effect and masked so no corner is missing.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShimbunTitleMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ShimbunTitleMaker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-24";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShimbunTitleMaker.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShimbunTitleMaker.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var DEFAULT_VERTICAL_MARGIN_PT = 5;    /* 上下マージンの既定値（pt）/ Default vertical margin */
    var DEFAULT_HORIZONTAL_MARGIN_PT = 8;  /* 左右マージンの既定値（pt）/ Default horizontal margin */
    var DEFAULT_LINE_GAP_PT = 5;         /* 二重線の間隔の既定値（pt）/ Default gap between the two rules */
    var RULE_STROKE_WIDTH_PT = 0.6;      /* 二重線の線幅（pt）/ Stroke width of the rules */
    var DEFAULT_ROTATION_DEG = 3;        /* 回転の既定値（°）/ Default rotation angle */
    var DEFAULT_CHARS_PER_LINE = 11;     /* 新聞ダミーの1行あたりの文字数 / Dummy characters per line */
    var DEFAULT_TIER_COUNT = 3;          /* 新聞ダミーの段数（0で作らない）/ Dummy tiers (0 = none) */
    var DEFAULT_EXTRA_TIER_COUNT = 1;    /* 新聞ダミーの上下に足す段数 / Dummy tiers added above and below */
    var DEFAULT_DUMMY_WIDTH_PT = 85.04;  /* 新聞ダミーの左右それぞれの幅（pt、30mm）/ Dummy width on each side */
    var DUMMY_LEADING_RATIO = 1.4;       /* 新聞ダミーの行送り（文字サイズ比）/ Dummy leading as a ratio of type size */
    var DUMMY_VERTICAL_SCALE = 85;       /* 新聞ダミーの垂直比率（%）/ Dummy vertical scale */
    /* 新聞ダミーのフォント（PostScript 名、前から順に探す）/ Dummy font PostScript names, tried in order */
    var DUMMY_FONT_NAMES = ["HiraMinProN-W3", "HiraMinPro-W3"];
    var DUMMY_RULE_WIDTH_PT = 0.3;       /* 段の区切り罫の線幅（pt）/ Stroke width of the tier rules */
    var DUMMY_BACKGROUND_CMYK = [0, 0, 10, 25]; /* 新聞ダミーの背景色 [C,M,Y,K]（%）/ Dummy background color */
    var RULE_CMYK = [0, 0, 0, 100];      /* 二重線・区切り罫の色 [C,M,Y,K]（%）/ Color of the rules */
    var DEFAULT_BLUR_RADIUS = 8;         /* 新聞ダミーのぼかし（ガウス）の半径（px、0でかけない）/ Dummy blur radius in px (0 = none) */
    var EFFECT_RESOLUTION = 300;         /* 効果のプレビュー解像度（ppi）/ effect preview resolution */
    /* 新聞ダミーの文字列（約物・小書きかな・長音は行頭禁則で行がずれるので入れない）
       Dummy text; no punctuation, small kana or long vowel marks, which kinsoku would push between lines */
    var DUMMY_TEXT = "吾輩は猫である名前はまだ無いどこで生れたかとんと見当がつかぬ何でも薄暗い所で泣いていた事だけは記憶している";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 15;              /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10]; /* パネルの余白 [左,上,右,下] / panel margins */
    var FIELD_LABEL_WIDTH = 96;           /* 項目名の幅 / width of the field labels */
    var FIELD_CHARS = 5;                  /* 数値入力欄の幅（文字数）/ width of the numeric fields */
    var FIELD_ROW_SPACING = 6;            /* 入力行の間隔 / spacing inside a field row */
    var BUTTON_SPACING = 10;              /* ボタンの間隔 / spacing between buttons */

    /**
     * パネルを縦並び・左揃えにして共通の余白を付ける
     * @param {Panel} targetPanel - 対象のパネル
     * @returns {void}
     */
    function applyPanelLayout(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "新聞見出しメーカー", en: "Newspaper Headline Maker" }
        },
        panel: {
            sideRules: { ja: "左右の二重線", en: "Side Rules" },
            newspaperDummy: { ja: "新聞ダミー", en: "Newspaper Dummy" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            marginVertical: { ja: "上下の伸び", en: "Extend" },
            marginHorizontal: { ja: "見出しとのアキ", en: "Spacing" },
            lineGap: { ja: "線の間隔", en: "Gap" },
            charsPerLine: { ja: "字詰", en: "Chars per line" },
            tierCount: { ja: "段数", en: "Tiers" },
            extraTierCount: { ja: "上下に追加", en: "Extra tiers" },
            dummyWidth: { ja: "左右の幅", en: "Width" },
            blurRadius: { ja: "ぼかし", en: "Blur" },
            rotation: { ja: "回転", en: "Rotation" }
        },
        unit: {
            chars: { ja: "文字", en: "chars" },
            tiers: { ja: "段", en: "tiers" }
        },
        tooltip: {
            marginVertical: { ja: "二重線を対象の上端／下端から伸ばす量", en: "How far the rules extend beyond the top and bottom of the object" },
            marginHorizontal: {
                ja: "対象と内側の線との距離。新聞ダミーの行をそろえるため、少し広がることがあります。",
                en: "Distance between the object and the inner rule. It may widen slightly so the dummy lines align."
            },
            lineGap: { ja: "内側の線と外側の線の間隔", en: "Distance between the inner and outer rule" },
            charsPerLine: { ja: "1段の1行に入れる文字数", en: "Characters in one line of a tier" },
            tierCount: {
                ja: "二重線の天地を分ける段数。0で新聞ダミーを作りません。",
                en: "Number of tiers the rule height is split into. 0 adds no dummy text."
            },
            extraTierCount: {
                ja: "上と下にそれぞれ足す、左右の幅いっぱいの段の数。0で足しません。",
                en: "Full-width tiers added above and below each. 0 adds none."
            },
            dummyWidth: { ja: "左右それぞれの新聞ダミーの幅", en: "Width of the dummy text on each side" },
            blurRadius: {
                ja: "新聞ダミーにかけるぼかし（ガウス）の半径（px）。0でかけません。",
                en: "Radius of the Gaussian blur on the dummy text, in pixels. 0 applies none."
            },
            rotation: {
                ja: "全体のグループに［変形］効果で適用する回転角度。0で効果を付けません。",
                en: "Rotation applied to the whole group with the Transform effect. 0 adds no effect."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTarget: { ja: "テキストオブジェクトまたはグループを1つ選択してください。", en: "Select one text object or one group." }
        },
        itemName: {
            ruleGroup: { ja: "左右二重線", en: "Side rules" },
            dummyGroup: { ja: "新聞ダミー", en: "Newspaper dummy" }
        }
    };

    /**
     * ラベルを表示言語で取得する
     * @param {string} labelPath - "panel.sideRules" のようなラベルのパス
     * @returns {string} 表示用の文字列
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length && labelNode; i++) {
            labelNode = labelNode[pathKeys[i]];
        }
        if (!labelNode) return labelPath;
        return labelNode[uiLang] || labelNode.en || labelNode.ja || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 入力の読み取り / Input handling
    // =========================================

    /**
     * 数値入力欄を↑↓キーで増減できるようにする（Shiftで10、Optionで0.1）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負の値を許すか
     * @param {Function} [onValueChange] - 値を書き換えた後に呼ぶ関数
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onValueChange) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            /* Shiftで10単位、Optionで0.1単位 / Shift steps by 10, Option by 0.1 */
            var delta = keyboard.shiftKey ? 10 : (keyboard.altKey ? 0.1 : 1);

            if (keyboard.shiftKey) {
                /* 10の倍数にスナップ / snap to multiples of ten */
                value = (event.keyName == "Up") ?
                    Math.ceil((value + 1) / delta) * delta :
                    Math.floor((value - 1) / delta) * delta;
            } else {
                value += (event.keyName == "Up") ? delta : -delta;
            }
            event.preventDefault();

            /* Option時だけ小数第1位まで残す / keep one decimal only while Option is held */
            value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);
            if (!allowNegative && value < 0) value = 0;

            editText.text = String(value);
            /* text の代入では onChanging が発火しない / assigning text does not fire onChanging */
            if (onValueChange) onValueChange();
        });
    }

    /**
     * 数値の入力文字列を読む。空白を除き、全角のカンマ・小数点を半角にし、
     * 「1,5」のような小数カンマは小数点に、「1,000」のような桁区切りは取り除く
     * @param {string} inputText - 入力された文字列
     * @returns {number} 読み取った数値（読めなければ NaN）
     */
    function parseLocaleNumber(inputText) {
        var normalizedText = String(inputText);
        normalizedText = normalizedText.replace(/\s+/g, "").replace(/，/g, ",").replace(/．/g, ".");
        if (normalizedText.indexOf(",") >= 0 && normalizedText.indexOf(".") < 0) {
            normalizedText = normalizedText.replace(/,/g, ".");
        } else {
            normalizedText = normalizedText.replace(/,/g, "");
        }
        return Number(normalizedText);
    }

    /**
     * 入力欄の値を数値で読む（読めないときは既定値）
     * @param {EditText} inputField - 対象の入力欄
     * @param {number} fallbackValue - 読めなかったときに使う値
     * @param {boolean} allowNegative - 負の値を許すか（許さないときは0に丸める）
     * @returns {number} 読み取った値
     */
    function readFieldAsNumber(inputField, fallbackValue, allowNegative) {
        var enteredValue = parseLocaleNumber(inputField.text);
        if (isNaN(enteredValue)) return fallbackValue;
        if (!allowNegative && enteredValue < 0) return 0;
        return enteredValue;
    }

    /**
     * 入力欄の値（現在の単位）を pt に換算して読む。負の値は0に丸める
     * @param {EditText} inputField - 対象の入力欄
     * @param {number} fallbackPt - 読めなかったときに使う値（pt）
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @returns {number} 換算した値（pt）
     */
    function readFieldAsPoints(inputField, fallbackPt, pointsPerUnit) {
        var enteredValue = readFieldAsNumber(inputField, NaN, false);
        return isNaN(enteredValue) ? fallbackPt : enteredValue * pointsPerUnit;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確認し、ダイアログでプレビューしながら二重線を描画してグループ化する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var targetItem = getTargetItem(doc.selection);
        if (!targetItem) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        showDialog(doc, targetItem);
    }

    /**
     * 基準にできる選択かを判定して返す（テキスト1つ、またはグループ1つ）
     * @param {Object} currentSelection - ドキュメントの選択
     * @returns {PageItem|null} 基準にするオブジェクト（対象外なら null）
     */
    function getTargetItem(currentSelection) {
        if (!currentSelection || currentSelection.length !== 1) return null;

        var selectedItem = currentSelection[0];
        var isSupported = (
            selectedItem.typename === "TextFrame" ||
            selectedItem.typename === "GroupItem"
        );
        return isSupported ? selectedItem : null;
    }

    /**
     * 二重線の上下端と、内側・外側の線の X 座標を求める
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {{marginVerticalPt: number, marginHorizontalPt: number, lineGapPt: number}} headlineSettings - マージンと間隔（pt）
     * @returns {{topY: number, bottomY: number, innerLeftX: number, innerRightX: number, outerLeftX: number, outerRightX: number}} 二重線の枠
     */
    function getRuleFrame(targetItem, headlineSettings) {
        /* visibleBounds: [左, 上, 右, 下] / visibleBounds: [left, top, right, bottom] */
        var targetBounds = targetItem.visibleBounds;
        var innerLeftX = targetBounds[0] - headlineSettings.marginHorizontalPt;
        var innerRightX = targetBounds[2] + headlineSettings.marginHorizontalPt;
        return {
            topY: targetBounds[1] + headlineSettings.marginVerticalPt,
            bottomY: targetBounds[3] - headlineSettings.marginVerticalPt,
            innerLeftX: innerLeftX,
            innerRightX: innerRightX,
            outerLeftX: innerLeftX - headlineSettings.lineGapPt,
            outerRightX: innerRightX + headlineSettings.lineGapPt
        };
    }

    /**
     * 対象の左右に二重線を描画する
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {Object} ruleFrame - getRuleFrame() で求めた二重線の枠
     * @param {{lineGapPt: number}} headlineSettings - 間隔（pt）
     * @returns {GroupItem} 生成した二重線のグループ
     */
    function drawSideRules(doc, targetItem, ruleFrame, headlineSettings) {
        var ruleGroup = targetItem.layer.groupItems.add();
        ruleGroup.name = getLabel("itemName.ruleGroup");

        var strokeColor = createProcessColor(doc, RULE_CMYK);
        var lineXPositions = [ruleFrame.innerLeftX, ruleFrame.innerRightX];
        /* 間隔が0なら内側の線に重なるので、外側の線は描かない / A zero gap would stack the rules, so skip the outer pair */
        if (headlineSettings.lineGapPt > 0) {
            lineXPositions.push(ruleFrame.outerLeftX);
            lineXPositions.push(ruleFrame.outerRightX);
        }
        for (var i = 0; i < lineXPositions.length; i++) {
            createRuleLine(
                ruleGroup,
                [lineXPositions[i], ruleFrame.topY],
                [lineXPositions[i], ruleFrame.bottomY],
                strokeColor,
                RULE_STROKE_WIDTH_PT
            );
        }

        /* 対象のすぐ背面へ / Move the rules just behind the target */
        ruleGroup.move(targetItem, ElementPlacement.PLACEAFTER);
        return ruleGroup;
    }

    /**
     * 新聞ダミーの文字サイズ・行送り・左右の行数を求める。
     * 二重線の天地（A）を段数で分け、段の間は1文字分空ける。
     * 1文字の送り（文字サイズ × 垂直比率）は A ÷（段数 × 字詰 ＋ 段数 − 1）。
     * 左右の幅は行の整数倍に切り詰める
     * @param {Object} ruleFrame - getRuleFrame() で求めた二重線の枠
     * @param {{charsPerLine: number, tierCount: number, dummyWidthPt: number}} headlineSettings - 新聞ダミーの設定
     * @returns {Object|null} 新聞ダミーの寸法（作らないときは null）
     */
    function computeDummyLayout(ruleFrame, headlineSettings) {
        var tierCount = headlineSettings.tierCount;
        var charsPerLine = headlineSettings.charsPerLine;
        var ruleHeight = ruleFrame.topY - ruleFrame.bottomY;
        if (tierCount < 1 || charsPerLine < 1 || headlineSettings.dummyWidthPt <= 0 || ruleHeight <= 0) return null;

        /* 縦組みでは垂直比率のぶん字送りが詰まる / in vertical text the vertical scale shortens each character's advance */
        var charAdvance = ruleHeight / (tierCount * charsPerLine + tierCount - 1);
        var fontSize = charAdvance / (DUMMY_VERTICAL_SCALE / 100);
        var dummyLayout = {
            tierCount: tierCount,
            charsPerLine: charsPerLine,
            fontSize: fontSize,
            charAdvance: charAdvance,
            linePitch: fontSize * DUMMY_LEADING_RATIO,
            tierPitch: (charsPerLine + 1) * charAdvance
        };
        dummyLayout.sideLineCount = countFittingLines(dummyLayout, headlineSettings.dummyWidthPt);
        dummyLayout.sideTextWidth = getTextWidth(dummyLayout, dummyLayout.sideLineCount);
        return dummyLayout;
    }

    /**
     * 上下の段の行が左右の段の行と揃うように、二重線を左右へ均等に広げる。
     * 左の本文の右端から右の本文の右端までが行送りの整数倍になればよい
     * @param {Object} ruleFrame - getRuleFrame() で求めた二重線の枠（書き換える）
     * @param {Object} dummyLayout - computeDummyLayout() で求めた寸法
     * @returns {void}
     */
    function alignRuleFrameToDummyLines(ruleFrame, dummyLayout) {
        var linePitch = dummyLayout.linePitch;
        /* 外側の線と本文の間は1文字分 / one character of space between the outer rule and the text */
        var lineSpan = (ruleFrame.outerRightX - ruleFrame.outerLeftX) + 2 * dummyLayout.fontSize + dummyLayout.sideTextWidth;
        /* 浮動小数の誤差で1行ぶん余計に広げない / keep float error from adding a whole extra line */
        var alignedSpan = Math.ceil(lineSpan / linePitch - 1e-6) * linePitch;
        var shiftX = (alignedSpan - lineSpan) / 2;

        ruleFrame.innerLeftX -= shiftX;
        ruleFrame.outerLeftX -= shiftX;
        ruleFrame.innerRightX += shiftX;
        ruleFrame.outerRightX += shiftX;
    }

    /**
     * 二重線の外側の左右に段組みの新聞ダミーを描画し、上下には左右の幅いっぱいの段を足す。
     * 段の間は1文字分空けて区切り罫を引く
     * @param {Document} doc - 対象のドキュメント
     * @param {Layer} targetLayer - 描画先のレイヤー
     * @param {Object} ruleFrame - 二重線の枠
     * @param {Object} dummyLayout - computeDummyLayout() で求めた寸法
     * @param {number} extraTierCount - 上下それぞれに足す段数
     * @returns {GroupItem} 生成した新聞ダミーのグループ
     */
    function drawNewspaperDummy(doc, targetLayer, ruleFrame, dummyLayout, extraTierCount) {
        var fontSize = dummyLayout.fontSize;
        var charAdvance = dummyLayout.charAdvance;
        var tierPitch = dummyLayout.tierPitch;
        var sideLineCount = dummyLayout.sideLineCount;
        var sideTextWidth = dummyLayout.sideTextWidth;

        var dummyGroup = targetLayer.groupItems.add();
        dummyGroup.name = getLabel("itemName.dummyGroup");
        var tierContext = {
            targetLayer: targetLayer,
            dummyGroup: dummyGroup,
            charsPerLine: dummyLayout.charsPerLine,
            fontSize: fontSize,
            charAdvance: charAdvance,
            linePitch: dummyLayout.linePitch,
            textFont: findDummyFont(),
            strokeColor: createProcessColor(doc, RULE_CMYK)
        };

        /* 外側の線と本文の間は1文字分 / one character of space between the outer rule and the text */
        var sideTextLefts = [
            ruleFrame.outerLeftX - fontSize - sideTextWidth,
            ruleFrame.outerRightX + fontSize
        ];

        /* 左右の段。2段目以降は上の段との間に罫 / side tiers; a rule above every tier but the first */
        for (var i = 0; i < sideTextLefts.length; i++) {
            for (var k = 0; k < dummyLayout.tierCount; k++) {
                var sideTierTop = ruleFrame.topY - k * tierPitch;
                addDummyTier(tierContext, sideTextLefts[i], sideTierTop, sideLineCount);
                if (k > 0) addTierRule(tierContext, sideTextLefts[i], sideTierTop + charAdvance / 2, sideTextWidth);
            }
        }

        /* 上下の段は左の本文の左端から右の本文の右端まで。alignRuleFrameToDummyLines() で行送りの整数倍になっている
           Tiers above and below span both side regions, already a whole number of lines wide */
        var fullLeft = sideTextLefts[0];
        var fullWidth = sideTextLefts[1] + sideTextWidth - fullLeft;
        var fullLineCount = countFittingLines(tierContext, fullWidth);

        /* どれも内側（ボックス側）の段との間に罫 / each has a rule on the side facing the box */
        for (var j = 1; j <= extraTierCount; j++) {
            var upperTop = ruleFrame.topY + j * tierPitch;
            addDummyTier(tierContext, fullLeft, upperTop, fullLineCount);
            addTierRule(tierContext, fullLeft, upperTop - tierPitch + charAdvance / 2, fullWidth);

            var lowerTop = ruleFrame.bottomY - charAdvance - (j - 1) * tierPitch;
            addDummyTier(tierContext, fullLeft, lowerTop, fullLineCount);
            addTierRule(tierContext, fullLeft, lowerTop + charAdvance / 2, fullWidth);
        }
        return dummyGroup;
    }

    /**
     * 新聞ダミーのフォントを探す（見つからなければ null で既定のフォントのまま）
     * @returns {TextFont|null} フォント
     */
    function findDummyFont() {
        for (var i = 0; i < DUMMY_FONT_NAMES.length; i++) {
            /* getByName は見つからないと例外 / getByName throws when the font is missing */
            try {
                return app.textFonts.getByName(DUMMY_FONT_NAMES[i]);
            } catch (e) {}
        }
        return null;
    }

    /**
     * 指定の幅に収まる行数を返す（最低1行）
     * @param {{fontSize: number, linePitch: number}} tierContext - 段の共通設定（または新聞ダミーの寸法）
     * @param {number} availableWidth - 使える幅（pt）
     * @returns {number} 行数
     */
    function countFittingLines(tierContext, availableWidth) {
        /* ちょうど割り切れる幅が誤差で1行減らないよう、わずかに足す / a tiny epsilon keeps exact fits from losing a line */
        var lineCount = Math.floor((availableWidth - tierContext.fontSize) / tierContext.linePitch + 1e-6) + 1;
        return Math.max(1, lineCount);
    }

    /**
     * 行数ぶんの本文の幅（最初の行の右端から最後の行の左端まで）を返す
     * @param {{fontSize: number, linePitch: number}} tierContext - 段の共通設定
     * @param {number} lineCount - 行数
     * @returns {number} 幅（pt）
     */
    function getTextWidth(tierContext, lineCount) {
        return (lineCount - 1) * tierContext.linePitch + tierContext.fontSize;
    }

    /**
     * 段の区切り罫を1本引く
     * @param {{dummyGroup: GroupItem, strokeColor: Object}} tierContext - 段の共通設定
     * @param {number} ruleLeft - 左端の X 座標
     * @param {number} ruleY - Y 座標
     * @param {number} ruleWidth - 長さ（pt）
     * @returns {void}
     */
    function addTierRule(tierContext, ruleLeft, ruleY, ruleWidth) {
        createRuleLine(
            tierContext.dummyGroup,
            [ruleLeft, ruleY],
            [ruleLeft + ruleWidth, ruleY],
            tierContext.strokeColor,
            DUMMY_RULE_WIDTH_PT
        );
    }

    /**
     * 新聞ダミー1段ぶんの縦組みエリア内文字を作り、新聞ダミーのグループに入れる。
     * 本文がちょうど指定の行数で終わるように、幅と文字数を合わせる
     * @param {{targetLayer: Layer, dummyGroup: GroupItem, charsPerLine: number, fontSize: number, charAdvance: number, linePitch: number, textFont: TextFont}} tierContext - 段の共通設定
     * @param {number} textLeft - 本文の左端の X 座標
     * @param {number} tierTop - 上端の Y 座標
     * @param {number} lineCount - 行数
     * @returns {TextFrame} 生成したテキストフレーム
     */
    function addDummyTier(tierContext, textLeft, tierTop, lineCount) {
        var charsPerLine = tierContext.charsPerLine;
        var charAdvance = tierContext.charAdvance;
        /* 丸めで行や文字がこぼれないよう、幅は左に、高さは下に半端な余裕を足す
           Add slack to the left and bottom so rounding never drops a line or a character */
        var widthSlack = (tierContext.linePitch - tierContext.fontSize) / 2;
        var frameWidth = getTextWidth(tierContext, lineCount) + widthSlack;
        var frameHeight = charsPerLine * charAdvance + charAdvance / 2;

        var targetLayer = tierContext.targetLayer;
        var framePath = targetLayer.pathItems.rectangle(tierTop, textLeft - widthSlack, frameWidth, frameHeight);
        var dummyFrame = targetLayer.textFrames.areaText(framePath, TextOrientation.VERTICAL);
        dummyFrame.contents = buildDummyText(lineCount * charsPerLine);

        var dummyAttributes = dummyFrame.textRange.characterAttributes;
        if (tierContext.textFont) dummyAttributes.textFont = tierContext.textFont;
        dummyAttributes.size = tierContext.fontSize;
        dummyAttributes.verticalScale = DUMMY_VERTICAL_SCALE;
        dummyAttributes.autoLeading = false;
        dummyAttributes.leading = tierContext.linePitch;

        dummyFrame.move(tierContext.dummyGroup, ElementPlacement.PLACEATEND);
        return dummyFrame;
    }

    /**
     * ダミー文字列を指定の文字数まで繰り返す
     * @param {number} charCount - 文字数
     * @returns {string} ダミー文字列
     */
    function buildDummyText(charCount) {
        var dummyText = "";
        while (dummyText.length < charCount) {
            dummyText += DUMMY_TEXT;
        }
        return dummyText.substr(0, charCount);
    }

    /**
     * 対象・二重線・新聞ダミーをグループ化し、新聞ダミーをぼかし、回転があれば［変形］効果で回転する。
     * 最後に、回転後も欠けない長方形で全体をマスクする
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {Object} headlineSettings - readHeadlineSettings() でまとめた設定
     * @returns {GroupItem} 全体をまとめたグループ（マスクできたときはクリップグループ）
     */
    function buildHeadlineGroup(doc, targetItem, headlineSettings) {
        var ruleFrame = getRuleFrame(targetItem, headlineSettings);
        var dummyLayout = computeDummyLayout(ruleFrame, headlineSettings);
        /* 二重線を描く前に、新聞ダミーの行に合わせて左右の間隔を調整 / adjust the side spacing to the dummy lines before drawing the rules */
        if (dummyLayout) alignRuleFrameToDummyLines(ruleFrame, dummyLayout);

        var ruleGroup = drawSideRules(doc, targetItem, ruleFrame, headlineSettings);
        var dummyGroup = dummyLayout ?
            drawNewspaperDummy(doc, targetItem.layer, ruleFrame, dummyLayout, headlineSettings.extraTierCount) :
            null;
        var headlineGroup = groupHeadlineParts(targetItem, ruleGroup, dummyGroup);

        var blursDummy = dummyGroup && headlineSettings.blurRadius > 0;
        if (dummyGroup) {
            /* ぼかす前の範囲で背景を敷く / lay the background over the unblurred dummy bounds */
            addDummyBackground(doc, headlineGroup, dummyGroup.geometricBounds);
            /* ボックス（対象と二重線）以外の新聞ダミーだけをぼかす / blur only the dummy text, not the title box */
            if (blursDummy) applyGaussianBlur(dummyGroup, headlineSettings.blurRadius);
        }

        /* 効果をかける前の範囲。ぼかしでにじむ縁はマスクの外に出す
           Bounds before the effects; keep the blurred fringe outside the mask */
        var contentBounds = headlineGroup.geometricBounds;
        var blurInset = blursDummy ? headlineSettings.blurRadius : 0;

        if (headlineSettings.rotationDeg !== 0) {
            headlineGroup.applyEffect(buildTransformEffectXml(headlineSettings.rotationDeg));
        }
        return maskWithInscribedRect(headlineGroup, contentBounds, blurInset, headlineSettings.rotationDeg);
    }

    /**
     * 対象の位置にグループを置き、前面から対象・二重線・新聞ダミーの順に入れる
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @param {GroupItem} ruleGroup - 二重線のグループ
     * @param {GroupItem|null} dummyGroup - 新聞ダミーのグループ（無ければ null）
     * @returns {GroupItem} まとめたグループ
     */
    function groupHeadlineParts(targetItem, ruleGroup, dummyGroup) {
        var headlineGroup = targetItem.layer.groupItems.add();
        headlineGroup.move(targetItem, ElementPlacement.PLACEBEFORE);
        targetItem.move(headlineGroup, ElementPlacement.PLACEATEND);
        ruleGroup.move(headlineGroup, ElementPlacement.PLACEATEND);
        if (dummyGroup) dummyGroup.move(headlineGroup, ElementPlacement.PLACEATEND);
        return headlineGroup;
    }

    /**
     * 新聞ダミー全体を覆う背景色の長方形を、グループの最背面に置く
     * @param {Document} doc - 対象のドキュメント
     * @param {GroupItem} headlineGroup - 追加先のグループ
     * @param {number[]} dummyBounds - 新聞ダミーの範囲 [左, 上, 右, 下]
     * @returns {PathItem} 生成した長方形
     */
    function addDummyBackground(doc, headlineGroup, dummyBounds) {
        var backgroundRect = headlineGroup.pathItems.rectangle(
            dummyBounds[1],
            dummyBounds[0],
            dummyBounds[2] - dummyBounds[0],
            dummyBounds[1] - dummyBounds[3]
        );
        backgroundRect.filled = true;
        backgroundRect.fillColor = createProcessColor(doc, DUMMY_BACKGROUND_CMYK);
        backgroundRect.stroked = false;
        backgroundRect.move(headlineGroup, ElementPlacement.PLACEATEND);
        return backgroundRect;
    }

    /**
     * 回転後の内容に収まる最大の長方形で、グループ全体をマスクする
     * @param {GroupItem} headlineGroup - マスクするグループ
     * @param {number[]} contentBounds - 回転前の範囲 [左, 上, 右, 下]
     * @param {number} edgeInset - 範囲の四辺から内側に詰める量（pt）
     * @param {number} rotationDeg - 回転角度（°）
     * @returns {GroupItem} クリップグループ（範囲が無いときは headlineGroup のまま）
     */
    function maskWithInscribedRect(headlineGroup, contentBounds, edgeInset, rotationDeg) {
        var contentWidth = contentBounds[2] - contentBounds[0] - edgeInset * 2;
        var contentHeight = contentBounds[1] - contentBounds[3] - edgeInset * 2;
        if (contentWidth <= 0 || contentHeight <= 0) return headlineGroup;

        var maskSize = getInscribedRectSize(contentWidth, contentHeight, rotationDeg);
        /* ［変形］効果は中心を基準に回すので、マスクも中心をそろえる / the Transform effect rotates around the center */
        var centerX = (contentBounds[0] + contentBounds[2]) / 2;
        var centerY = (contentBounds[1] + contentBounds[3]) / 2;

        var clipGroup = headlineGroup.layer.groupItems.add();
        clipGroup.move(headlineGroup, ElementPlacement.PLACEBEFORE);
        headlineGroup.move(clipGroup, ElementPlacement.PLACEATEND);

        var maskRect = clipGroup.pathItems.rectangle(
            centerY + maskSize.height / 2,
            centerX - maskSize.width / 2,
            maskSize.width,
            maskSize.height
        );
        /* クリップグループは最前面のパスでマスクする / a clip group masks with its frontmost path */
        maskRect.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        clipGroup.clipped = true;
        return clipGroup;
    }

    /**
     * 幅 × 高さの長方形を回転させたとき、その内側に収まる面積最大の水平な長方形の寸法を返す
     * @param {number} rectWidth - 回転前の幅（pt）
     * @param {number} rectHeight - 回転前の高さ（pt）
     * @param {number} rotationDeg - 回転角度（°）
     * @returns {{width: number, height: number}} 内接する長方形の寸法（pt）
     */
    function getInscribedRectSize(rectWidth, rectHeight, rotationDeg) {
        var angleRad = rotationDeg * Math.PI / 180;
        var sinA = Math.abs(Math.sin(angleRad));
        var cosA = Math.abs(Math.cos(angleRad));
        if (sinA < 1e-10) return { width: rectWidth, height: rectHeight };

        var widthIsLonger = rectWidth >= rectHeight;
        var longSide = widthIsLonger ? rectWidth : rectHeight;
        var shortSide = widthIsLonger ? rectHeight : rectWidth;

        /* 細長い場合や45°付近は、短辺の中点2つで決まる / thin rectangles and angles near 45° are bound by the short side */
        if (shortSide <= 2 * sinA * cosA * longSide || Math.abs(sinA - cosA) < 1e-10) {
            var halfShort = shortSide / 2;
            return widthIsLonger ?
                { width: halfShort / sinA, height: halfShort / cosA } :
                { width: halfShort / cosA, height: halfShort / sinA };
        }
        /* それ以外は4辺すべてが回転した長方形に接する / otherwise all four corners touch the rotated rectangle */
        var cos2A = cosA * cosA - sinA * sinA;
        return {
            width: (rectWidth * cosA - rectHeight * sinA) / cos2A,
            height: (rectHeight * cosA - rectWidth * sinA) / cos2A
        };
    }

    /**
     * ぼかし（ガウス）をライブエフェクトとして適用する
     * @param {PageItem} targetItem - 対象のページアイテム
     * @param {number} blurRadius - ぼかしの半径（px）
     * @returns {void}
     */
    function applyGaussianBlur(targetItem, blurRadius) {
        var effectXml = '<LiveEffect name="Adobe PSL Gaussian Blur">' +
            '<Dict data="R blur ' + blurRadius + ' R PrevDocScale 1 I PrevDres ' + EFFECT_RESOLUTION + ' "/>' +
            '</LiveEffect>';
        targetItem.applyEffect(effectXml);
    }

    /**
     * 中心を基準に回転だけを行う［変形］効果の XML を作る
     * @param {number} rotateDegrees - 回転角度（°、反時計回りが正）
     * @returns {string} LiveEffect の XML
     */
    function buildTransformEffectXml(rotateDegrees) {
        return '<LiveEffect name="Adobe Transform"><Dict data="' +
            'R scaleH_Percent 100' +
            ' R scaleV_Percent 100' +
            ' R scaleH_Factor 1' +
            ' R scaleV_Factor 1' +
            ' R moveH_Pts 0' +
            ' R moveV_Pts 0' +
            ' R rotate_Degrees ' + rotateDegrees +
            ' R rotate_Radians ' + (rotateDegrees * Math.PI / 180) +
            ' I numCopies 0' +
            ' I pinPoint 4' +
            ' B scaleLines 1' +
            ' B transformPatterns 1' +
            ' B transformObjects 1' +
            ' B reflectX 0' +
            ' B reflectY 0' +
            ' B randomize 0' +
            ' "/></LiveEffect>';
    }

    /**
     * 直線を1本描く
     * @param {GroupItem} parentGroup - 追加先のグループ
     * @param {number[]} startPoint - 始点 [X, Y]
     * @param {number[]} endPoint - 終点 [X, Y]
     * @param {Object} strokeColor - 線の色
     * @param {number} strokeWidth - 線幅（pt）
     * @returns {PathItem} 生成したパス
     */
    function createRuleLine(parentGroup, startPoint, endPoint, strokeColor, strokeWidth) {
        var linePath = parentGroup.pathItems.add();
        linePath.setEntirePath([startPoint, endPoint]);
        linePath.stroked = true;
        linePath.filled = false;
        linePath.strokeWidth = strokeWidth;
        linePath.strokeColor = strokeColor;
        linePath.strokeCap = StrokeCap.BUTTENDCAP;
        return linePath;
    }

    /**
     * ドキュメントのカラーモードに合わせた色を CMYK の値から作る。RGB では単純換算した近似色
     * @param {Document} doc - 対象のドキュメント
     * @param {number[]} cmykValues - [C, M, Y, K]（%）
     * @returns {Object} RGBColor または CMYKColor
     */
    function createProcessColor(doc, cmykValues) {
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            var keyRatio = 1 - cmykValues[3] / 100;
            var rgbColor = new RGBColor();
            rgbColor.red = Math.round(255 * (1 - cmykValues[0] / 100) * keyRatio);
            rgbColor.green = Math.round(255 * (1 - cmykValues[1] / 100) * keyRatio);
            rgbColor.blue = Math.round(255 * (1 - cmykValues[2] / 100) * keyRatio);
            return rgbColor;
        }
        var cmykColor = new CMYKColor();
        cmykColor.cyan = cmykValues[0];
        cmykColor.magenta = cmykValues[1];
        cmykColor.yellow = cmykValues[2];
        cmykColor.black = cmykValues[3];
        return cmykColor;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログと各コントロールを作る
     * @param {string} unitLabel - 入力欄に添える単位の表示名
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @returns {{dialog: Window, fieldInputs: Object}} ダイアログと、キーごとの入力欄
     */
    function buildDialog(unitLabel, pointsPerUnit, onValueChange) {
        var headlineDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        headlineDialog.orientation = "column";
        headlineDialog.alignChildren = ["fill", "top"];
        headlineDialog.margins = DIALOG_MARGINS;

        /* 入力欄はキー名で fieldLabel／tooltip のラベルを引く。pt の既定値は現在の単位に換算して表示
           Each field looks up its fieldLabel/tooltip by key; point defaults are shown in the current unit */
        var fieldInputs = {};

        addNumericFieldPanel(headlineDialog, "panel.sideRules", [
            { key: "marginVertical", initialText: formatUnitValue(DEFAULT_VERTICAL_MARGIN_PT, pointsPerUnit), unitLabel: unitLabel },
            { key: "marginHorizontal", initialText: formatUnitValue(DEFAULT_HORIZONTAL_MARGIN_PT, pointsPerUnit), unitLabel: unitLabel },
            { key: "lineGap", initialText: formatUnitValue(DEFAULT_LINE_GAP_PT, pointsPerUnit), unitLabel: unitLabel }
        ], fieldInputs, onValueChange);

        addNumericFieldPanel(headlineDialog, "panel.newspaperDummy", [
            { key: "charsPerLine", initialText: String(DEFAULT_CHARS_PER_LINE), unitLabel: getLabel("unit.chars") },
            { key: "tierCount", initialText: String(DEFAULT_TIER_COUNT), unitLabel: getLabel("unit.tiers") },
            { key: "extraTierCount", initialText: String(DEFAULT_EXTRA_TIER_COUNT), unitLabel: getLabel("unit.tiers") },
            { key: "dummyWidth", initialText: formatUnitValue(DEFAULT_DUMMY_WIDTH_PT, pointsPerUnit), unitLabel: unitLabel },
            { key: "blurRadius", initialText: String(DEFAULT_BLUR_RADIUS), unitLabel: "px" }
        ], fieldInputs, onValueChange);

        addNumericFieldPanel(headlineDialog, "panel.options", [
            { key: "rotation", initialText: String(DEFAULT_ROTATION_DEG), unitLabel: "°", allowNegative: true }
        ], fieldInputs, onValueChange);

        addButtonRow(headlineDialog);
        return { dialog: headlineDialog, fieldInputs: fieldInputs };
    }

    /**
     * パネルを作り、数値入力の行を並べる
     * @param {Window} parentDialog - 追加先のダイアログ
     * @param {string} panelLabelPath - パネル見出しのラベルパス
     * @param {Object[]} fieldSpecs - 行の定義（key, initialText, unitLabel, allowNegative）
     * @param {Object} fieldInputs - 作った入力欄を key ごとに入れる先
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @returns {Panel} 作ったパネル
     */
    function addNumericFieldPanel(parentDialog, panelLabelPath, fieldSpecs, fieldInputs, onValueChange) {
        var fieldPanel = parentDialog.add("panel", undefined, getLabel(panelLabelPath));
        applyPanelLayout(fieldPanel);
        for (var i = 0; i < fieldSpecs.length; i++) {
            fieldInputs[fieldSpecs[i].key] = addNumericFieldRow(fieldPanel, fieldSpecs[i], onValueChange);
        }
        return fieldPanel;
    }

    /**
     * 「項目名＋数値入力欄＋単位」の行を生成する。項目名と helpTip は fieldLabel.<key> / tooltip.<key> から引く
     * @param {Panel|Group} parentContainer - 追加先
     * @param {{key: string, initialText: string, unitLabel: string, allowNegative: boolean}} fieldSpec - 行の定義
     * @param {Function} onValueChange - 入力値が変わったときに呼ぶ関数
     * @returns {EditText} 生成した入力欄
     */
    function addNumericFieldRow(parentContainer, fieldSpec, onValueChange) {
        var fieldRow = parentContainer.add("group");
        fieldRow.orientation = "row";
        fieldRow.alignChildren = ["left", "center"];
        fieldRow.spacing = FIELD_ROW_SPACING;

        var fieldLabel = fieldRow.add("statictext", undefined, labelText("fieldLabel." + fieldSpec.key));
        fieldLabel.preferredSize = [FIELD_LABEL_WIDTH, -1];
        fieldLabel.justify = "right";

        var inputField = fieldRow.add("edittext", undefined, fieldSpec.initialText);
        inputField.characters = FIELD_CHARS;
        inputField.helpTip = getLabel("tooltip." + fieldSpec.key);
        inputField.onChanging = onValueChange;
        changeValueByArrowKey(inputField, fieldSpec.allowNegative === true, onValueChange);

        fieldRow.add("statictext", undefined, fieldSpec.unitLabel);
        return inputField;
    }

    /**
     * ボタンエリア（キャンセル・OK を左右中央）を作る
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addButtonRow(parentDialog) {
        var btnRowGroup = parentDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.alignment = ["center", "top"];
        btnRowGroup.spacing = BUTTON_SPACING;

        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    }

    /**
     * pt の値を現在の単位の表示文字列にする（小数第2位まで）
     * @param {number} valuePt - 値（pt）
     * @param {number} pointsPerUnit - 1単位あたりのポイント数
     * @returns {string} 入力欄に入れる文字列
     */
    function formatUnitValue(valuePt, pointsPerUnit) {
        var unitValue = valuePt / pointsPerUnit;
        return String(Math.round(unitValue * 100) / 100);
    }

    /**
     * ダイアログを表示し、入力に合わせてプレビューする。
     * プレビューは対象の複製で作り、その間は元の対象を隠しておく。
     * OK で元の対象を使って確定し、キャンセルで元に戻す
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem} targetItem - 基準にするテキストまたはグループ
     * @returns {GroupItem|null} 確定したグループ。キャンセル時は null
     */
    function showDialog(doc, targetItem) {
        var unitInfo = getUnitInfo();
        var pointsPerUnit = unitInfo.pointsPerUnit;
        var previewGroup = null;

        /**
         * 入力欄の値を設定値にまとめる
         * @returns {{marginVerticalPt: number, marginHorizontalPt: number, lineGapPt: number, charsPerLine: number, tierCount: number, extraTierCount: number, dummyWidthPt: number, blurRadius: number, rotationDeg: number}} マージン・間隔・幅（pt）、文字数・段数、ぼかし（px）、回転（°）
         */
        function readHeadlineSettings() {
            var fieldInputs = dialogUI.fieldInputs;
            return {
                marginVerticalPt: readFieldAsPoints(fieldInputs.marginVertical, DEFAULT_VERTICAL_MARGIN_PT, pointsPerUnit),
                marginHorizontalPt: readFieldAsPoints(fieldInputs.marginHorizontal, DEFAULT_HORIZONTAL_MARGIN_PT, pointsPerUnit),
                lineGapPt: readFieldAsPoints(fieldInputs.lineGap, DEFAULT_LINE_GAP_PT, pointsPerUnit),
                /* 文字数・段数は整数に切り捨て / character and tier counts are truncated to integers */
                charsPerLine: Math.floor(readFieldAsNumber(fieldInputs.charsPerLine, DEFAULT_CHARS_PER_LINE, false)),
                tierCount: Math.floor(readFieldAsNumber(fieldInputs.tierCount, DEFAULT_TIER_COUNT, false)),
                extraTierCount: Math.floor(readFieldAsNumber(fieldInputs.extraTierCount, DEFAULT_EXTRA_TIER_COUNT, false)),
                dummyWidthPt: readFieldAsPoints(fieldInputs.dummyWidth, DEFAULT_DUMMY_WIDTH_PT, pointsPerUnit),
                blurRadius: readFieldAsNumber(fieldInputs.blurRadius, DEFAULT_BLUR_RADIUS, false),
                rotationDeg: readFieldAsNumber(fieldInputs.rotation, DEFAULT_ROTATION_DEG, true)
            };
        }

        /**
         * 前回のプレビューを削除する
         * @returns {void}
         */
        function removePreview() {
            if (previewGroup) {
                previewGroup.remove();
                previewGroup = null;
            }
        }

        /**
         * 前回のプレビューを消して、対象の複製と現在の入力値で作り直す
         * @returns {void}
         */
        function updatePreview() {
            removePreview();
            /* 複製は隠した元の hidden を引き継ぐので表示に戻す / the duplicate inherits hidden from the original */
            var previewSource = targetItem.duplicate();
            previewSource.hidden = false;
            previewGroup = buildHeadlineGroup(doc, previewSource, readHeadlineSettings());
            app.redraw();
        }

        var dialogUI = buildDialog(unitInfo.label, pointsPerUnit, updatePreview);
        dialogUI.dialog.onShow = updatePreview;

        targetItem.hidden = true;
        var dialogResult = dialogUI.dialog.show();
        removePreview();
        targetItem.hidden = false;

        if (dialogResult !== 1) {
            targetItem.selected = true;
            app.redraw();
            return null;
        }

        var headlineGroup = buildHeadlineGroup(doc, targetItem, readHeadlineSettings());
        doc.selection = [headlineGroup];
        return headlineGroup;
    }

    main();
})();
