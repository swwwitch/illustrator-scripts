#target illustrator
#targetengine "FontClipboard"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの先頭文字を基準に、文字属性・段落属性を取得して保存します。
保存した内容は ApplyTextAttributesFromClipboard.jsx から読み取って適用できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyTextAttributesToClipboard.md

### Overview

Reads the character and paragraph attributes of the selected text, taking the first character as the reference, and stores them.
The stored values can then be applied with ApplyTextAttributesFromClipboard.jsx.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyTextAttributesToClipboard.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CopyTextAttributesToClipboard"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2021-04-10";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyTextAttributesToClipboard.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyTextAttributesToClipboard.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        errorNoDocument: {
            ja: "ドキュメントを開いてください。",
            en: "Please open a document."
        },
        errorNoTextSelection: {
            ja: "テキストオブジェクトを 1 つ選択するか、テキスト編集モードで文字列を選択してください。",
            en: "Select one text object, or select text in text editing mode."
        },
        errorEmptyTextRange: {
            ja: "文字のある範囲を選択してください。",
            en: "Select a text range that contains characters."
        },
        copiedMessageTitle: {
            ja: "フォント情報を取得しました",
            en: "Font information copied"
        },
        fontSizeLeadingPanelTitle: {
            ja: "フォント・サイズ、行送り",
            en: "Font, Size & Leading"
        },
        kerningPanelTitle: {
            ja: "カーニング関連",
            en: "Kerning"
        },
        paragraphOtherPanelTitle: {
            ja: "段落属性ほか",
            en: "Paragraph & Other"
        },
        fillGraphicStylePanelTitle: {
            ja: "塗りとグラフィックスタイル",
            en: "Fill & Graphic Style"
        },
        graphicStyle: {
            ja: "グラフィックスタイル",
            en: "Graphic Style"
        },
        graphicStyleNotRegistered: {
            ja: "—",
            en: "—"
        },
        postScriptName: {
            ja: "PostScript名",
            en: "PostScript name"
        },
        fontFamilyName: {
            ja: "フォントファミリ名",
            en: "Font family name"
        },
        fontStyle: {
            ja: "スタイル",
            en: "Style"
        },
        fontSize: {
            ja: "フォントサイズ",
            en: "Font size"
        },
        leading: {
            ja: "行送り",
            en: "Leading"
        },
        autoLeading: {
            ja: "自動行送り",
            en: "Auto leading"
        },
        tsume: {
            ja: "文字ツメ",
            en: "Tsume"
        },
        tracking: {
            ja: "トラッキング",
            en: "Tracking"
        },
        kerningMethod: {
            ja: "カーニング",
            en: "Kerning"
        },
        proportionalMetrics: {
            ja: "プロポーショナルメトリクス",
            en: "Proportional metrics"
        },
        orientation: {
            ja: "組み方向",
            en: "Orientation"
        },
        justification: {
            ja: "行揃え",
            en: "Alignment"
        },
        fillColor: {
            ja: "塗り",
            en: "Fill"
        },
        fillColorNone: {
            ja: "なし",
            en: "None"
        },
        valueOn: {
            ja: "オン",
            en: "On"
        },
        valueOff: {
            ja: "オフ",
            en: "Off"
        },
        closeButton: {
            ja: "閉じる",
            en: "Close"
        }
    };

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

    /* 数値を表示用に丸める / Round number for display */
    function formatNumberForDisplay(value, digits) {
        var multiplier = Math.pow(10, digits);
        var rounded = Math.round(value * multiplier) / multiplier;
        return String(rounded);
    }

    /* pt値を現在の文字単位で表示 / Format point value using current text unit */
    function formatPointValueForDisplay(pointValue) {
        var textUnit = getUnitInfo("text/units");
        var displayValue = pointValue / textUnit.pointsPerUnit;
        return formatNumberForDisplay(displayValue, 3) + " " + textUnit.label;
    }

    /* Boolean値をオン/オフラベルに変換 / Convert boolean to on/off label */
    function formatBooleanLabel(boolValue) {
        return boolValue ? getLabel("valueOn") : getLabel("valueOff");
    }

    // =========================================
    // 塗り色ユーティリティ / Fill color utilities
    // =========================================

    /* 色を安全に複製 / Safely clone a color */
    // 参照のまま保持すると Illustrator 側の状態に引きずられるため、新規インスタンスを返す
    function cloneColor(color) {
        if (!color) return null;
        switch (color.typename) {
            case "RGBColor":
                var rgb = new RGBColor();
                rgb.red = color.red;
                rgb.green = color.green;
                rgb.blue = color.blue;
                return rgb;
            case "CMYKColor":
                var cmyk = new CMYKColor();
                cmyk.cyan = color.cyan;
                cmyk.magenta = color.magenta;
                cmyk.yellow = color.yellow;
                cmyk.black = color.black;
                return cmyk;
            case "GrayColor":
                var gray = new GrayColor();
                gray.gray = color.gray;
                return gray;
            case "SpotColor":
                var spot = new SpotColor();
                spot.spot = color.spot;
                spot.tint = color.tint;
                return spot;
            case "GradientColor":
                var grad = new GradientColor();
                grad.gradient = color.gradient;
                grad.angle = color.angle;
                grad.length = color.length;
                grad.origin = color.origin;
                grad.matrix = color.matrix;
                return grad;
            case "NoColor":
                return new NoColor();
            default:
                return null;
        }
    }

    /* テキスト範囲から塗り色を取得 / Get fill color from a text range */
    function getTextRangeFillColor(textRange) {
        try {
            return textRange.characterAttributes.fillColor;
        } catch (e) {
            return null;
        }
    }

    /* 色を表示用文字列へ整形 / Format color for display */
    function formatColorForDisplay(color) {
        if (!color || !color.typename) return getLabel("fillColorNone");
        switch (color.typename) {
            case "RGBColor":
                return "RGB(" + Math.round(color.red) + ", " + Math.round(color.green) + ", " + Math.round(color.blue) + ")";
            case "CMYKColor":
                return "CMYK(" + Math.round(color.cyan) + ", " + Math.round(color.magenta) + ", " + Math.round(color.yellow) + ", " + Math.round(color.black) + ")";
            case "GrayColor":
                return "Gray(" + Math.round(color.gray) + ")";
            case "SpotColor":
                try {
                    return "Spot: " + color.spot.name + " (" + Math.round(color.tint) + "%)";
                } catch (e) {
                    return "Spot";
                }
            case "GradientColor":
                try {
                    return "Gradient: " + color.gradient.name;
                } catch (e) {
                    return "Gradient";
                }
            case "NoColor":
                return getLabel("fillColorNone");
            default:
                return color.typename;
        }
    }

    // =========================================
    // 文字属性ユーティリティ / Text attribute utilities
    // =========================================

    /* テキスト範囲から親TextFrameを取得 / Get parent TextFrame from text range */
    // テキスト編集モードでは textRange.parent が Story を返すことがあるため
    // story.textFrames 経由で TextFrame を取得する
    function getParentTextFrameFromTextRange(textRange) {
        if (!textRange) return null;
        try {
            var story = textRange.story;
            if (story && story.textFrames && story.textFrames.length > 0) {
                return story.textFrames[0];
            }
        } catch (e) { }
        var parentItem = textRange.parent;
        while (parentItem && parentItem.typename !== "TextFrame") {
            parentItem = parentItem.parent;
        }
        return parentItem || null;
    }

    /* カーニング方式の表示名を取得 / Get display label for kerning method */
    function getKerningMethodLabel(kerningMethod) {
        switch (kerningMethod) {
            case AutoKernType.AUTO:
                return (uiLang === "ja") ? "メトリクス" : "Metrics";
            case AutoKernType.METRICSROMANONLY:
                return (uiLang === "ja") ? "和文等幅" : "Metrics - Roman Only";
            case AutoKernType.OPTICAL:
                return (uiLang === "ja") ? "オプティカル" : "Optical";
            default:
                return (uiLang === "ja") ? "なし" : "None";
        }
    }

    /* 組み方向の表示名を取得 / Get display label for text orientation */
    function getOrientationLabel(textFrame) {
        if (!textFrame) return "";
        return (textFrame.orientation === TextOrientation.VERTICAL)
            ? ((uiLang === "ja") ? "縦組み" : "Vertical")
            : ((uiLang === "ja") ? "横組み" : "Horizontal");
    }

    /* 自動行送り値を真偽値として解釈 / Interpret auto leading value as boolean */
    function isAutoLeadingValueOn(autoLeadingValue) {
        return autoLeadingValue === true || autoLeadingValue === 1;
    }

    /* 文字範囲の自動行送りを確認 / Check auto leading in a text range */
    function hasAutoLeadingInTextRange(textRange) {
        if (!textRange) return false;

        try {
            if (textRange.characterAttributes && isAutoLeadingValueOn(textRange.characterAttributes.autoLeading)) {
                return true;
            }
        } catch (e) {
        }

        try {
            if (textRange.characters && textRange.characters.length > 0) {
                for (var characterIndex = 0; characterIndex < textRange.characters.length; characterIndex++) {
                    try {
                        if (isAutoLeadingValueOn(textRange.characters[characterIndex].characterAttributes.autoLeading)) {
                            return true;
                        }
                    } catch (characterError) {
                    }
                }
            }
        } catch (e2) {
        }

        try {
            if (textRange.lines && textRange.lines.length > 0) {
                for (var lineIndex = 0; lineIndex < textRange.lines.length; lineIndex++) {
                    try {
                        if (isAutoLeadingValueOn(textRange.lines[lineIndex].characterAttributes.autoLeading)) {
                            return true;
                        }
                    } catch (lineError) {
                    }
                }
            }
        } catch (e3) {
        }

        return false;
    }

    /* 自動行送りかどうかを安全に取得 / Safely detect whether auto leading is enabled */
    function getAutoLeadingState(textRange, firstCharacter, textFrame) {
        try {
            if (firstCharacter && firstCharacter.characterAttributes) {
                if (isAutoLeadingValueOn(firstCharacter.characterAttributes.autoLeading)) {
                    return true;
                }
            }
        } catch (e) {
        }

        if (hasAutoLeadingInTextRange(textRange)) {
            return true;
        }

        try {
            if (textFrame && textFrame.textRange && hasAutoLeadingInTextRange(textFrame.textRange)) {
                return true;
            }
        } catch (e2) {
        }

        return false;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 一時アクション（再利用パーツ） / Temporary action (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は runTemporaryAction / loadTemporaryActionSet / unloadTemporaryActionSet / toActionHex / buildActionNameLines
    // 2. アクション定義は配列＋join("\n") で組み立てる（''' は ES3 の構文エラー）。
    //    セット名・アクション名は英数字にする。/name [ n 16進 ] は buildActionNameLines で作るとバイト数がずれない
    //      var actionSource = [
    //          "/version 3"
    //      ].concat(buildActionNameLines("", "MySet"), [
    //          "/isOpen 1", "/actionCount 1", "/action-1 {"
    //      ], buildActionNameLines("\t", "myAction"), [ … ]).join("\n");
    // 3. 1回だけ実行するとき:
    //      if (!runTemporaryAction(actionSource, "MySet", "myAction")) alert(getLabel("alert.actionFailed"));
    //    何度も実行するとき（オブジェクトごとなど）は、読み込み・解除を1回ずつにする:
    //      if (!loadTemporaryActionSet(actionSource, "MySet")) { alert(…); return; }
    //      try { for (…) app.doScript("myAction", "MySet"); } finally { unloadTemporaryActionSet("MySet"); }
    // 4. 失敗は例外にせず false で返す（$.writeln に理由を出す）。警告を出すかはコピー先で決める
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // グラフィックスタイル登録 / Graphic style registration
    // =========================================

    /* 登録するスタイル名 / Style name to register
       register-temp-style.jsx と同じ固定名。Apply 側はこの名前を取り出して適用する。
       Same fixed name as register-temp-style.jsx; the Apply script reads this name back. */
    var TEMP_STYLE_NAME = "temp_style";
    var TEMP_STYLE_ACTION_SET = "GraphicStyle";
    var TEMP_STYLE_ACTION_NAME = "AddNewWithoutName";

    /* 強制的に無名グラフィックスタイルを追加するアクション定義 /
       Dynamic action definition that appends an unnamed graphic style */
    function buildForceNewGraphicStyleAction() {
        return '/version 3 /name [ 12 477261706869635374796c65 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382 a4e383ab ] /value 1 } } }';
    }

    /* 現在の選択を配列として退避 / Capture current selection as an array */
    // app.selection が TextRange のときは .length が文字数を返すので、配列扱いせず単体で退避する
    function captureCurrentDocumentSelection() {
        var items = [];
        try {
            var currentSelection = app.selection;
            if (!currentSelection) return items;
            if (currentSelection.typename === "TextRange") {
                items.push(currentSelection);
                return items;
            }
            if (typeof currentSelection.length === "number") {
                for (var selectionIndex = 0; selectionIndex < currentSelection.length; selectionIndex++) {
                    items.push(currentSelection[selectionIndex]);
                }
                return items;
            }
            items.push(currentSelection);
        } catch (e) {
        }
        return items;
    }

    /* 退避した選択を復元 / Restore captured selection */
    // テキスト編集モードへの完全復帰はできないが、TextRange は単体で再選択する
    function restoreDocumentSelection(items) {
        try {
            app.selection = null;
        } catch (clearError) {
        }
        if (!items || items.length === 0) return;

        if (items.length === 1 && items[0] && items[0].typename === "TextRange") {
            try {
                items[0].select();
                return;
            } catch (textRangeSelectError) {
            }
            try {
                app.selection = items[0];
                return;
            } catch (textRangeAssignError) {
            }
            return;
        }

        try {
            app.selection = items;
        } catch (assignError) {
            try {
                app.selection = items[0];
            } catch (singleAssignError) {
            }
        }
    }

    /* TextFrame の見た目を temp_style として登録 / Register the TextFrame's appearance as temp_style */
    // register-temp-style.jsx と同じ手順：既存 temp_style を削除→アクションで末尾に追加→末尾を改名
    function registerTextFrameAsTempGraphicStyle(textFrame) {
        if (!textFrame) return null;

        var activeDoc;
        try {
            activeDoc = app.activeDocument;
        } catch (e) {
            return null;
        }
        if (!activeDoc) return null;

        var graphicStyles;
        try {
            graphicStyles = activeDoc.graphicStyles;
        } catch (graphicStylesError) {
            return null;
        }
        if (!graphicStyles) return null;

        /* 既存の temp_style を削除 / Remove the existing temp_style */
        try {
            graphicStyles.getByName(TEMP_STYLE_NAME).remove();
        } catch (removeExistingError) {
        }

        var savedSelection = captureCurrentDocumentSelection();

        try {
            app.selection = null;
        } catch (clearSelectionError) {
        }

        try {
            textFrame.selected = true;
        } catch (selectError) {
            restoreDocumentSelection(savedSelection);
            return null;
        }

        var beforeCount;
        try {
            beforeCount = graphicStyles.length;
        } catch (countError) {
            restoreDocumentSelection(savedSelection);
            return null;
        }

        var actionSucceeded = runTemporaryAction(buildForceNewGraphicStyleAction(), TEMP_STYLE_ACTION_SET, TEMP_STYLE_ACTION_NAME);

        var registeredName = null;
        if (actionSucceeded) {
            try {
                if (graphicStyles.length > beforeCount) {
                    graphicStyles[graphicStyles.length - 1].name = TEMP_STYLE_NAME;
                    registeredName = TEMP_STYLE_NAME;
                }
            } catch (renameError) {
            }
        }

        restoreDocumentSelection(savedSelection);
        return registeredName;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 結果表示ダイアログ / Show result dialog */
    function showResultDialog(info) {
        var dialog = new Window("dialog", getLabel("copiedMessageTitle"));
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.spacing = 10;
        dialog.margins = 16;

        function createInfoPanel(titleKey) {
            var panel = dialog.add("panel", undefined, getLabel(titleKey));
            panel.orientation = "column";
            panel.alignChildren = "left";
            panel.margins = [15, 20, 15, 10];
            panel.spacing = 8;
            panel.alignment = ["fill", "top"];
            return panel;
        }

        function addRow(parent, key, value) {
            var row = parent.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];
            row.spacing = 8;
            var labelControl = row.add("statictext", undefined, labelText(key));
            labelControl.preferredSize.width = 180;
            labelControl.justify = "right";
            row.add("statictext", undefined, String(value));
        }

        var fontSizeLeadingPanel = createInfoPanel("fontSizeLeadingPanelTitle");
        var kerningPanel = createInfoPanel("kerningPanelTitle");
        var paragraphOtherPanel = createInfoPanel("paragraphOtherPanelTitle");
        var fillGraphicStylePanel = createInfoPanel("fillGraphicStylePanelTitle");

        addRow(fontSizeLeadingPanel, "postScriptName", info.font.name);
        addRow(fontSizeLeadingPanel, "fontFamilyName", info.font.family);
        addRow(fontSizeLeadingPanel, "fontStyle", info.font.style);
        addRow(fontSizeLeadingPanel, "fontSize", formatPointValueForDisplay(info.size));
        addRow(fontSizeLeadingPanel, "leading", info.autoLeading ? "—" : formatPointValueForDisplay(info.leading));
        addRow(fontSizeLeadingPanel, "autoLeading", formatBooleanLabel(info.autoLeading));

        addRow(kerningPanel, "kerningMethod", info.kerningMethodLabel);
        addRow(kerningPanel, "proportionalMetrics", formatBooleanLabel(info.proportionalMetrics));
        addRow(kerningPanel, "tracking", info.tracking);
        addRow(kerningPanel, "tsume", info.tsume);

        addRow(paragraphOtherPanel, "orientation", info.orientationLabel);
        addRow(paragraphOtherPanel, "justification", info.justificationLabel);

        /* 塗りとグラフィックスタイル：取得したものを並べて表示
           Fill & Graphic Style: show captured info side by side */
        addRow(fillGraphicStylePanel, "fillColor", info.fillColorLabel);
        addRow(fillGraphicStylePanel, "graphicStyle", info.graphicStyleName ? info.graphicStyleName : getLabel("graphicStyleNotRegistered"));

        var buttonRow = addButtonRow(dialog);
        var btnClose = buttonRow.rightGroup.add("button", undefined, getLabel("closeButton"), { name: "ok" });
        btnClose.onClick = function () { dialog.close(); };

        prepareDialogWindow(dialog, SCRIPT_NAME);
        dialog.show();
    }

    /* 行揃えの表示名を取得 / Get display label for paragraph justification */
    function getJustificationLabel(justification) {
        try {
            if (justification === Justification.LEFT) {
                return (uiLang === "ja") ? "左揃え" : "Left";
            }
            if (justification === Justification.CENTER) {
                return (uiLang === "ja") ? "中央揃え" : "Center";
            }
            if (justification === Justification.RIGHT) {
                return (uiLang === "ja") ? "右揃え" : "Right";
            }
            if (justification === Justification.FULLJUSTIFYLASTLINELEFT) {
                return (uiLang === "ja") ? "均等配置（最終行左揃え）" : "Justify with last line aligned left";
            }
            if (justification === Justification.FULLJUSTIFYLASTLINECENTER) {
                return (uiLang === "ja") ? "均等配置（最終行中央揃え）" : "Justify with last line centered";
            }
            if (justification === Justification.FULLJUSTIFYLASTLINERIGHT) {
                return (uiLang === "ja") ? "均等配置（最終行右揃え）" : "Justify with last line aligned right";
            }
            if (justification === Justification.FULLJUSTIFY) {
                return (uiLang === "ja") ? "均等配置" : "Full justify";
            }
        } catch (e) {
            return "—";
        }
        return "—";
    }

    (function () {
        if (app.documents.length === 0) {
            alert(getLabel("errorNoDocument"));
            return;
        }

        var currentSelection = app.selection;
        var sourceTextRange = null;

        /* テキスト編集モードでの部分選択 / Partial text selection in text editing mode */
        // app.selection は TextRange オブジェクトが直接返る / app.selection directly returns a TextRange object
        if (currentSelection && currentSelection.typename === "TextRange") {
            sourceTextRange = currentSelection;
        }
        /* オブジェクト選択モード / Object selection mode */
        // 配列で 1 つの TextFrame が返る / A single TextFrame is returned in an array
        else if (currentSelection && currentSelection.length === 1 && currentSelection[0] && currentSelection[0].typename === "TextFrame") {
            sourceTextRange = currentSelection[0].textRange;
        }

        if (!sourceTextRange) {
            alert(getLabel("errorNoTextSelection"));
            return;
        }

        if (sourceTextRange.characters.length === 0) {
            alert(getLabel("errorEmptyTextRange"));
            return;
        }

        /* 先頭文字の属性を参照 / Read attributes from the first character */
        // 混在書式時に値が不定になるのを避ける / Avoid ambiguous values when formatting is mixed
        var sourceFirstCharacter = sourceTextRange.characters[0];
        var sourceCharacterAttributes = sourceFirstCharacter.characterAttributes;
        var sourceParagraphAttributes = sourceFirstCharacter.paragraphAttributes;

        var sourceTextFrame = getParentTextFrameFromTextRange(sourceTextRange);

        var copiedLeading = sourceCharacterAttributes.leading;
        var copiedAutoLeading = getAutoLeadingState(sourceTextRange, sourceFirstCharacter, sourceTextFrame);
        var copiedTsume = Math.round(sourceCharacterAttributes.Tsume);
        var copiedKerningMethod = sourceCharacterAttributes.kerningMethod;
        var copiedProportionalMetrics = sourceCharacterAttributes.proportionalMetrics;
        var copiedTracking = sourceCharacterAttributes.tracking;
        var copiedJustification = sourceParagraphAttributes.justification;
        var copiedJustificationLabel = getJustificationLabel(copiedJustification);

        var copiedOrientation = sourceTextFrame ? sourceTextFrame.orientation : null;
        var copiedOrientationLabel = getOrientationLabel(sourceTextFrame);
        var copiedKerningMethodLabel = getKerningMethodLabel(copiedKerningMethod);

        /* 先頭文字の塗り色を取得 / Get fill color from the first character */
        var copiedFillColor = cloneColor(getTextRangeFillColor(sourceFirstCharacter));
        var copiedFillColorLabel = formatColorForDisplay(copiedFillColor);

        /* 現在の TextFrame の見た目を temp_style として登録（既存があれば上書き）/
           Register the current TextFrame's appearance as temp_style (overwrites if exists) */
        var registeredGraphicStyleName = registerTextFrameAsTempGraphicStyle(sourceTextFrame);

        var hasUsableFillColor = !!(copiedFillColor && copiedFillColor.typename && copiedFillColor.typename !== "NoColor");

        $.global.FontClipboard = {
            font: {
                name: sourceCharacterAttributes.textFont.name,
                family: sourceCharacterAttributes.textFont.family,
                style: sourceCharacterAttributes.textFont.style
            },
            size: sourceCharacterAttributes.size,
            leading: copiedLeading,
            autoLeading: copiedAutoLeading,
            tsume: copiedTsume,
            tracking: copiedTracking,
            kerningMethod: copiedKerningMethod,
            kerningMethodLabel: copiedKerningMethodLabel,
            proportionalMetrics: copiedProportionalMetrics,
            orientation: copiedOrientation,
            orientationLabel: copiedOrientationLabel,
            justification: copiedJustification,
            justificationLabel: copiedJustificationLabel,
            fillColor: copiedFillColor,
            fillColorLabel: copiedFillColorLabel,
            /* 自動登録した temp_style の名前。Apply 側はこの名前でグラフィックスタイルを引く /
               Name of the auto-registered temp_style. The Apply script looks up the graphic style by this name. */
            graphicStyleName: registeredGraphicStyleName,
            /* Apply 側ラジオの初期選択：塗りがあれば fill、グラフィックスタイルだけあれば graphicStyle /
               Default radio in the Apply dialog: fill when a fill is captured, graphicStyle when only a style is, otherwise none */
            fillOrGraphicStyle: hasUsableFillColor ? "fill" : (registeredGraphicStyleName ? "graphicStyle" : "none")
            // 将来拡張
        };

        showResultDialog($.global.FontClipboard);
    })();

})();
