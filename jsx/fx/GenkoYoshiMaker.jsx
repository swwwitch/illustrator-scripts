#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの文字送りをマス目にそろえ、1文字分の間隔で罫線を引き、行の両側に罫線を添えます。縦組み・横組みのどちらにも対応します。
罫線の濃度、前後のマス、行の追加、線種、十字線などをダイアログで指定し、結果はその場でプレビューされます。

詳細は README を参照してください。

### Overview

Matches the character advance of the selected text to a grid of cells, draws a rule at every
character interval, and adds rules along both sides of each line. Vertical and horizontal text
are both supported. A dialog sets the rule density, the extra cells and lines, the rule styles
and the crosshairs, and previews the result live.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GenkoYoshiMaker";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GenkoYoshiMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GenkoYoshiMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n036c34760079"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 言語とラベル
    // Language and labels
    // =========================================

    var uiLang = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextFrame: { ja: "テキストオブジェクトを1つ選択してください。", en: "Please select a single text object." },
            noCharacter: { ja: "テキストに文字が入力されていません。", en: "The selected text is empty." },
            noFontSize: { ja: "文字サイズを取得できませんでした。", en: "Could not read the font size." },
            skipped: { ja: "設定できなかった属性:", en: "Attributes that could not be set:" },
            presetName: { ja: "プリセット名を入力してください。", en: "Enter a preset name." },
            presetSaved: {
                ja: "プリセットを追加しました（このセッションのみ）。\n\nPRESETS に貼り付けると残せます:\n",
                en: "Preset added for this session.\n\nPaste this into PRESETS to keep it:\n"
            }
        },
        itemName: {
            layerName: { ja: "原稿用紙罫線", en: "Manuscript Grid" },
            groupName: { ja: "罫線", en: "Rules" }
        },
        dialog: {
            title: { ja: "原稿用紙メーカー", en: "Manuscript Paper Maker" }
        },
        panel: {
            overall: { ja: "全体", en: "Overall" },
            text: { ja: "文字", en: "Characters" },
            /* 行の両側に添える罫線。縦組みでは縦罫、横組みでは横罫 */
            lineRule: {
                vertical: { ja: "縦罫", en: "Vertical rules" },
                horizontal: { ja: "横罫", en: "Horizontal rules" }
            },
            /* マスの区切りの罫線。縦組みでは横罫、横組みでは縦罫 */
            cellRule: {
                vertical: { ja: "横罫", en: "Horizontal rules" },
                horizontal: { ja: "縦罫", en: "Vertical rules" }
            },
            cross: { ja: "十字線", en: "Crosshairs" }
        },
        label: {
            preset: { ja: "プリセット", en: "Preset" },
            scale: { ja: "比率", en: "Scale" },
            extension: { ja: "伸張", en: "Extension" },
            lineColor: { ja: "罫線の濃度", en: "Rule density" },
            extraCells: { ja: "前後のマス", en: "Extra cells" },
            /* 行が並ぶ向きは書字方向で変わる */
            extraLines: {
                vertical: { ja: "左右の行数", en: "Extra lines" },
                horizontal: { ja: "上下の行数", en: "Extra lines" }
            },
            lineStyle: { ja: "線種", en: "Style" },
            segments: { ja: "分割数", en: "Segments" },
            emphasisPrefix: { ja: "", en: "Thicker every" },
            emphasisSuffix: { ja: "文字ごとに太く", en: "characters" }
        },
        checkbox: {
            adjustSize: { ja: "文字の比率を調整", en: "Adjust the character scale" },
            sideRule: { ja: "外側にもう1本追加", en: "Add an outer rule" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        preset: {
            standard: { ja: "標準", en: "Standard" },
            noSideRule: { ja: "左右の縦罫なし", en: "No side rules" },
            noCross: { ja: "十字線なし", en: "No crosshairs" },
            rulesOnly: { ja: "シンプル", en: "Simple" },
            custom: { ja: "カスタム", en: "Custom" }
        },
        radio: {
            none: { ja: "なし", en: "None" },
            solid: { ja: "実線", en: "Solid" },
            dashed: { ja: "破線", en: "Dashed" }
        },
        tooltip: {
            preset: {
                ja: "設定をまとめて切り替えます。値を手で変えると「カスタム」になります。",
                en: "Fills the whole dialog at once. Editing any value switches to Custom."
            },
            lineColor: {
                ja: "罫線の濃度。shiftを押しながらドラッグすると10%刻みになります。十字線の濃度は別（30%）です。",
                en: "Density of the rules. Hold Shift to snap to 10% steps. The crosshairs keep their own density (30%)."
            },
            extraCells: {
                ja: "文字の前後（上下）に足す空のマスの数。",
                en: "Empty cells added before and after the text."
            },
            extraLines: {
                ja: "テキストの行の外側に足す、空の行の数。マスの1/4だけアキを空けて並べます。",
                en: "Empty lines added outside the text, spaced by a quarter of a cell."
            },
            adjustSize: {
                ja: "文字の比率を変えてマスに合わせます。オフでも、送り（トラッキング・ツメ・アキ・カーニング）は今の比率に合わせてそろえます。",
                en: "Scales the characters to the cells. When off, the advance is still matched to the current scaling."
            },
            scale: {
                ja: "文字の水平比率・垂直比率。下げた分はトラッキングで送りに戻します。",
                en: "Horizontal and vertical scale of the characters; the tracking gives the advance back."
            },
            extension: {
                ja: "縦罫をマスの上下へ伸ばす量。",
                en: "How far the vertical rules run past the cells."
            },
            sideRule: {
                ja: "いちばん外側の行の外へ、マスの1/4だけ離して縦罫を足します。",
                en: "Adds a vertical rule a quarter of a cell outside the outermost line."
            },
            emphasis: {
                ja: "指定した文字数ごとに横罫を太くします。",
                en: "Thickens the horizontal rule at the given interval."
            },
            crossSegments: {
                ja: "十字線1本あたりの線分の本数。0か3以上の奇数で、線分と間隔は同じ長さになります。",
                en: "Dashes per crosshair line: zero or an odd number of 3 or more; the dash and the gap are equal."
            },
            preview: {
                ja: "ダイアログを開いたまま、確定後の状態を表示します。",
                en: "Shows the finished result while the dialog stays open."
            },
            savePreset: {
                ja: "いまの設定に名前を付けてプリセットへ加えます。次回も使うには、表示されるコードを PRESETS に貼り付けてください。",
                en: "Adds the current settings to the preset list. Paste the code it shows into PRESETS to keep it for next time."
            }
        },
        unit: {
            mm: { ja: "mm", en: "mm" },
            percent: { ja: "%", en: "%" }
        },
        button: {
            save: { ja: "保存", en: "Save" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        }
    };

    /**
     * ラベルを現在の言語で取り出す
     * @param {object} labelEntry ja / en を持つラベル定義
     * @return {string} 表示用の文字列
     */
    function getLabel(labelEntry) {
        return (typeof labelEntry[uiLang] === "string") ? labelEntry[uiLang] : labelEntry.en;
    }

    /**
     * 書字方向に合うほうのラベル定義を取り出す
     * @param {object} labelEntry vertical / horizontal を持つラベル定義
     * @param {boolean} isVertical 縦組みかどうか
     * @return {object} ja / en を持つラベル定義
     */
    function orientedEntry(labelEntry, isVertical) {
        return isVertical ? labelEntry.vertical : labelEntry.horizontal;
    }

    /**
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {object} labelEntry ja / en を持つラベル定義
     * @return {string} コロン付きの項目名
     */
    function labelText(labelEntry) {
        return getLabel(labelEntry) + ((uiLang === "ja") ? "：" : ":");
    }

    /**
     * ミリメートルをポイントに変換する
     * @param {number} mm ミリメートル
     * @return {number} ポイント
     */
    function mmToPt(mm) {
        return mm * 72 / 25.4;
    }

    /**
     * ポイントをミリメートルに変換する
     * @param {number} pt ポイント
     * @return {number} ミリメートル
     */
    function ptToMm(pt) {
        return pt * 25.4 / 72;
    }

    /**
     * 文字サイズから伸張の既定値を求める（文字サイズの1/4、小数第1位に丸める）
     * @param {number} fontSize 文字サイズ（pt）
     * @return {number} 伸張（mm）
     */
    function getDefaultExtensionMM(fontSize) {
        return Math.round(ptToMm(fontSize / 4) * 10) / 10;
    }

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* プレビュー用レイヤー名（一時的に作って消す）/ Name of the temporary preview layer */
    var PREVIEW_LAYER_NAME = "__ManuscriptGrid_Preview";

    /* 前回の設定を覚えておくキー。Illustratorを終了するまで有効
       / Key that remembers the last settings until Illustrator quits */
    var SESSION_KEY = "__GenkoYoshiMakerSession";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var LAYOUT = {
        lineStrokeMM: 0.25,         /* 行の両側の罫線の太さ（mm）/ weight of the rules along each line (mm) */
        cellStrokeMM: 0.1,          /* マスの区切りの太さ（mm）/ weight of the cell rules (mm) */
        cellDashMM: [1, 1],         /* マスの区切りを破線にしたときの線分と間隔（mm）/ dash and gap of the dashed cell rules (mm) */
        emphasisStrokeMM: 0.25,     /* 太くするマスの区切りの太さ（mm）/ weight of the emphasized cell rules (mm) */
        crossStrokeMM: 0.1,         /* 十字線の太さ（mm）/ crosshair weight (mm) */
        crossGray: 30,              /* 十字線の濃度（%、100で黒）/ crosshair density (%, 100 is black) */
        sideRuleRatio: 0.25,        /* 外側の罫線の位置と、行と行のアキ（文字サイズに対する比率）/ side rule offset and gutter (ratio of the font size) */
        gridStartOffset: 0          /* 1本目の位置の微調整（pt、プラスで書字方向の手前へ）/ nudge of the first rule (pt) */
    };

    /* ダイアログの初期値 / Dialog defaults */
    var DEFAULTS = {
        strokeGray: 50,          /* 罫線の濃度（%、100で黒）/ rule density (%, 100 is black) */
        extraCells: 0,           /* 文字の前後に足すマスの数 / empty cells added before and after the text */
        extraLines: 0,           /* テキストの行の外側に足す、空の行の数 / empty lines added outside the text */
        adjustSize: true,        /* 文字を比率とトラッキングでマスに合わせるか / fit the characters to the cells */
        scale: 90,               /* 水平・垂直比率（%）/ horizontal and vertical scale (%) */
        extensionMM: null,       /* 罫線を書字方向へ伸ばす量（mm）。null は文字サイズの1/4 / null means a quarter of the font size */
        dashedCellRules: false,  /* マスの区切りを破線にするか / draw the cell rules dashed */
        emphasisEnabled: true,   /* 一定間隔の横罫を太くするか / thicken the horizontal rules at a fixed interval */
        emphasisEvery: 5,        /* 太くする間隔（文字数）/ interval of the thickened rules (characters) */
        showSideRules: true,     /* 左右に縦罫を添えるか / add the side rules */
        crossStyle: "dashed",    /* 十字線の線種（none / solid / dashed）/ crosshair style */
        crossSegments: 9,        /* 十字線1本あたりの線分の本数（線分と間隔は同じ長さ）/ dashes per crosshair line (dash and gap are equal) */
        preview: true            /* プレビューを表示するか / show the live preview */
    };

    // =========================================
    // セッションの記憶 / Session memory
    // =========================================

    /**
     * 前回の設定を取り出す（同じIllustratorのセッション内だけ）
     * @return {object|null} { values, presetIndex }。まだ無いときは null
     */
    function loadSession() {
        return $.global[SESSION_KEY] || null;
    }

    /**
     * 今回の設定を覚えておく（同じIllustratorのセッション内だけ）
     * @param {object} values プリセットと同じ形の設定値
     * @param {number} presetIndex 選んでいたプリセットの位置
     * @return {void}
     */
    function saveSession(values, presetIndex) {
        $.global[SESSION_KEY] = { values: values, presetIndex: presetIndex };
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /* ダイアログの設定をまとめて切り替える定義。settings が null のものは「カスタム」
       / Each preset fills the dialog at once; the one with null settings is "Custom" */
    var PRESETS = [
        { label: LABELS.preset.standard, settings: DEFAULTS },
        {
            label: LABELS.preset.noSideRule,
            settings: {
                strokeGray: 50, extraCells: 0, extraLines: 0, adjustSize: true, scale: 90, extensionMM: null, showSideRules: false,
                dashedCellRules: false, emphasisEnabled: true, emphasisEvery: 5,
                crossStyle: "dashed", crossSegments: 9
            }
        },
        {
            label: LABELS.preset.noCross,
            settings: {
                strokeGray: 50, extraCells: 0, extraLines: 0, adjustSize: true, scale: 90, extensionMM: null, showSideRules: true,
                dashedCellRules: false, emphasisEnabled: true, emphasisEvery: 5,
                crossStyle: "none", crossSegments: 9
            }
        },
        {
            label: LABELS.preset.rulesOnly,
            settings: {
                strokeGray: 50, extraCells: 0, extraLines: 0, adjustSize: true, scale: 90, extensionMM: 0, showSideRules: false,
                dashedCellRules: false, emphasisEnabled: false, emphasisEvery: 5,
                crossStyle: "none", crossSegments: 9
            }
        },
        { label: LABELS.preset.custom, settings: null }
    ];

    /* ダイアログの寸法 / Dialog metrics */
    var UI = {
        labelWidth: (uiLang === "ja") ? 90 : 130,  /* 項目名の幅（px）/ width of the row labels (px) */
        panelMargins: [16, 20, 16, 12],            /* パネルの余白 [左, 上, 右, 下]（px）/ panel margins */
        panelSpacing: 8,                           /* パネル内の行間隔（px）/ spacing between the rows in a panel */
        saveButtonWidth: 70                        /* ［保存］ボタンの幅（px）/ width of the Save button (px) */
    };

    // =========================================
    // 文字属性 / Character attributes
    // =========================================

    /**
     * テキスト全体に文字属性を1つ設定する
     * 取得済みの characterAttributes を使い回すと Error 9503 や片側無視が起きるため、毎回取り直す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {string} attributeName 属性名
     * @param {number|boolean|object} value 設定する値
     * @return {boolean} 設定できたかどうか
     */
    function setCharacterAttribute(textFrame, attributeName, value) {
        try {
            textFrame.textRange.characterAttributes[attributeName] = value;
            return true;
        } catch (attributeError) {
            return false;
        }
    }

    /**
     * 段落ごとに文字属性を1つ設定する
     * 文字ツメ・プロポーショナルメトリクス・自動カーニングは、テキスト全体へ書いても効かない
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {string} attributeName 属性名
     * @param {number|boolean|object} value 設定する値
     * @return {boolean} 設定できたかどうか
     */
    function setParagraphAttribute(textFrame, attributeName, value) {
        try {
            var paragraphs = textFrame.textRange.paragraphs;

            for (var i = 0; i < paragraphs.length; i++) {
                paragraphs[i].characterAttributes[attributeName] = value;
            }

            return true;
        } catch (attributeError) {
            return false;
        }
    }

    /**
     * 手動カーニングを全文字で0に戻す
     * カーニングは文字ごとに持つため、範囲へまとめて書いても残りが消えない
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {boolean} 設定できたかどうか
     */
    function clearManualKerning(textFrame) {
        try {
            var characters = textFrame.textRange.characters;

            for (var i = 0; i < characters.length; i++) {
                characters[i].kerning = 0;
            }

            return true;
        } catch (kerningError) {
            return false;
        }
    }

    /**
     * 1行ぶんの送りを求める（マス1つ＋行と行のアキ）
     * @param {number} cellSize マスの一辺（pt）
     * @return {number} 行送り（pt）
     */
    function getLinePitch(cellSize) {
        return cellSize * (1 + LAYOUT.sideRuleRatio);
    }

    /**
     * 比率を下げた分の文字送りを打ち消すトラッキングを求める
     * @param {number} scale 水平・垂直比率（%）
     * @return {number} トラッキング（1/1000em）
     */
    function getTracking(scale) {
        return 100000 / scale - 1000;
    }

    /**
     * 1文字＝1マスになるよう文字属性をそろえる（1文字目のカーニングは実測後に入れる）
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {number} scale 水平・垂直比率（%）
     * @param {boolean} applyScale 比率そのものを書き換えるか
     * @return {array} 設定できなかった属性名の配列
     */
    function alignCharactersToCells(textFrame, scale, applyScale) {
        var failed = [];
        var fontSize = textFrame.textRange.characters[0].characterAttributes.size;

        /* 文字前後のアキは「自動」（-1）。0は「アキなし」で別物
           / Aki before and after is set to auto (-1); zero would mean "no aki" */
        var rangeSettings = [
            ["tracking", getTracking(scale)],
            ["akiLeft", -1],
            ["akiRight", -1],
            /* 行と行のアキも入れた送りにそろえる。自動行送りのままだと行がマスから外れる
               / The leading carries the gutter, otherwise the lines drift off the cells */
            ["autoLeading", false],
            ["leading", getLinePitch(fontSize)]
        ];

        /* 比率を変えない設定でも、送りは今の比率に合わせてそろえる
           / Even when the scaling stays, the advance is still matched to it */
        if (applyScale) {
            rangeSettings.push(["horizontalScale", scale]);
            rangeSettings.push(["verticalScale", scale]);
        }

        for (var i = 0; i < rangeSettings.length; i++) {
            if (!setCharacterAttribute(textFrame, rangeSettings[i][0], rangeSettings[i][1])) {
                failed.push(rangeSettings[i][0]);
            }
        }

        /* 文字ツメ・プロポーショナルメトリクス・自動カーニングは段落単位で書く
           自動カーニングは「和文等幅」。手動カーニングは1文字目だけ別に入れる
           / Tsume, proportional metrics and auto kerning are written per paragraph */
        var paragraphSettings = [
            ["Tsume", 0],
            ["proportionalMetrics", false],
            ["kerningMethod", AutoKernType.NOAUTOKERN]
        ];

        for (var j = 0; j < paragraphSettings.length; j++) {
            if (!setParagraphAttribute(textFrame, paragraphSettings[j][0], paragraphSettings[j][1])) {
                failed.push(paragraphSettings[j][0]);
            }
        }

        /* 手動カーニングはいったん全文字0に戻す。1文字目だけ実測のあとに入れ直す
           / Clear the manual kerning first; the first character gets its value after the measuring */
        if (!clearManualKerning(textFrame)) {
            failed.push("kerning");
        }

        return failed;
    }

    /**
     * 各行の1文字目にトラッキングの1/4のカーニングを入れて、字面をマスの中央へ落とす
     * 実測より先に入れると罫線ごとずれるため、寸法を測ったあとに呼ぶ
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {number} scale 水平・垂直比率（%）
     * @return {boolean} 設定できたかどうか
     */
    function applyLineStartKerning(textFrame, scale) {
        var kerning = getTracking(scale) / 4;

        try {
            var lines = textFrame.lines;

            for (var i = 0; i < lines.length; i++) {
                var firstCharacter = lines[i].characters[0];

                /* カーニングは characterAttributes ではなく文字そのものが持つ。単位はトラッキングと同じ1/1000em
                   / Kerning lives on the character itself, not on characterAttributes, in the same 1/1000 em unit */
                firstCharacter.characterAttributes.kerningMethod = AutoKernType.NOAUTOKERN;
                firstCharacter.kerning = kerning;
            }

            return true;
        } catch (kerningError) {
            return false;
        }
    }

    /**
     * 罫線に合わせる比率を決める（比率を変えない設定のときは、いまの比率を使う）
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @return {number} 水平・垂直比率（%）
     */
    function getEffectiveScale(textFrame, settings) {
        if (settings.adjustSize) {
            return settings.scale;
        }

        /* 送りは書字方向の比率で決まる / The advance follows the scale along the writing direction */
        var characterAttributes = textFrame.textRange.characters[0].characterAttributes;
        var currentScale = isVerticalText(textFrame) ? characterAttributes.verticalScale : characterAttributes.horizontalScale;

        return (currentScale > 0) ? currentScale : 100;
    }

    /**
     * 行揃えを書字方向の先頭（縦組みは上、横組みは左）にそろえる
     * 行揃えを変えると文字が動くので、書き出し位置を元に戻す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {boolean} 設定できたかどうか
     */
    function alignParagraphsToStart(textFrame) {
        var before = textFrame.geometricBounds;

        try {
            var paragraphs = textFrame.textRange.paragraphs;

            for (var i = 0; i < paragraphs.length; i++) {
                paragraphs[i].paragraphAttributes.justification = Justification.LEFT;
            }
        } catch (justificationError) {
            return false;
        }

        app.redraw();

        var after = textFrame.geometricBounds;

        if (isVerticalText(textFrame)) {
            textFrame.translate(0, before[1] - after[1]);
        } else {
            textFrame.translate(before[0] - after[0], 0);
        }

        return true;
    }

    /**
     * 文字をマス目に合わせた状態にする（組み方向はそのまま）
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @param {number} scale 罫線に合わせる比率（%）
     * @return {array} 設定できなかった属性名の配列
     */
    function applyGridAttributes(textFrame, settings, scale) {
        var failed = alignCharactersToCells(textFrame, scale, settings.adjustSize);

        if (!alignParagraphsToStart(textFrame)) {
            failed.push("justification");
        }

        return failed;
    }

    // =========================================
    // 実測 / Measuring
    // =========================================

    /**
     * テキストが縦組みかどうかを返す
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {boolean} 縦組みかどうか
     */
    function isVerticalText(textFrame) {
        return textFrame.orientation === TextOrientation.VERTICAL;
    }

    /**
     * 行数と、1行あたりの最大文字数を数える
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {object} { lineCount, maxCharacters }（取得できないときは0）
     */
    function countGridCells(textFrame) {
        var lineCount = 0;
        var maxCharacters = 0;

        try {
            var lines = textFrame.lines;
            lineCount = lines.length;

            for (var i = 0; i < lines.length; i++) {
                var characterCount = lines[i].characters.length;

                if (characterCount > maxCharacters) {
                    maxCharacters = characterCount;
                }
            }
        } catch (lineError) {
            return { lineCount: 0, maxCharacters: 0 };
        }

        return { lineCount: lineCount, maxCharacters: maxCharacters };
    }

    /**
     * 罫線を引くのに必要な寸法を実測する（属性をそろえたあとに呼ぶ）
     * originX / originY は、1行目の1マス目の左上。cellCount は書字方向のマス数
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @return {object|null} { isVertical, originX, originY, cellSize, cellCount, lineCount }。文字サイズを取得できないときは null
     */
    function measureGrid(textFrame) {
        /* 先頭文字の文字サイズを基準にする（比率を変えても size は元の値のまま） */
        var fontSize = textFrame.textRange.characters[0].characterAttributes.size;

        if (!fontSize || fontSize <= 0) {
            return null;
        }

        var isVertical = isVerticalText(textFrame);
        var bounds = textFrame.geometricBounds;
        var cellCounts = countGridCells(textFrame);

        var lineCount = (cellCounts.lineCount > 0) ? cellCounts.lineCount : 1;

        /* 最後の文字の先にも罫線を引くため、マス数は1行の文字数から求める */
        var cellCount = cellCounts.maxCharacters;

        if (cellCount <= 0) {
            var runLength = isVertical ? (bounds[1] - bounds[3]) : (bounds[2] - bounds[0]);
            cellCount = Math.ceil(runLength / fontSize);
        }

        /* マスは文字サイズの正方形。字面は比率を下げた分だけ狭いので、
           行が並ぶ向きの寸法は実測せず、行送り×（行数−1）＋マス1つで求めて中央にそろえる
           / Cells are squares of the font size; the glyphs are narrower, so the line axis is centered instead of measured */
        var lineSize = getLinePitch(fontSize) * (lineCount - 1) + fontSize;

        var metrics = {
            isVertical: isVertical,
            cellSize: fontSize, /* マスの一辺。文字サイズそのもので、比率は反映しない */
            cellCount: cellCount,
            lineCount: lineCount
        };

        if (isVertical) {
            metrics.originX = (bounds[0] + bounds[2]) / 2 - lineSize / 2;
            metrics.originY = bounds[1] + LAYOUT.gridStartOffset;
        } else {
            metrics.originX = bounds[0] - LAYOUT.gridStartOffset;
            metrics.originY = (bounds[1] + bounds[3]) / 2 + lineSize / 2;
        }

        return metrics;
    }

    /**
     * テキストをマス目に合わせ、罫線の寸法を測る
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @return {object} { metrics, failed } 実測した寸法と、設定できなかった属性名
     */
    function fitTextToGrid(textFrame, settings) {
        var scale = getEffectiveScale(textFrame, settings);
        var failed = applyGridAttributes(textFrame, settings, scale);

        /* 属性変更後の再組版を反映させてから境界を読む */
        app.redraw();

        var metrics = measureGrid(textFrame);

        /* 行頭のカーニングは実測のあとに入れる（先に入れると罫線ごとずれてしまう） */
        if (metrics && !applyLineStartKerning(textFrame, scale)) {
            failed.push("kerning");
        }

        return { metrics: metrics, failed: failed };
    }

    // =========================================
    // 罫線 / Rules
    // =========================================

    /**
     * 罫線用のレイヤーを用意する（同名があれば再利用）
     * @param {Document} doc 対象のドキュメント
     * @param {string} layerName レイヤー名
     * @return {Layer} 罫線用のレイヤー
     */
    function prepareRuleLayer(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                var existingLayer = doc.layers[i];
                existingLayer.locked = false;
                existingLayer.visible = true;
                return existingLayer;
            }
        }

        var createdLayer = doc.layers.add();
        createdLayer.name = layerName;
        return createdLayer;
    }

    /**
     * その罫線グループが対象テキストのものか（テキストの中心を囲んでいるか）を調べる
     * 罫線はテキストより広いので、中心が入っていれば対象とみなせる
     * 同じレイヤーに別のテキストの罫線が並ぶため、名前だけでは見分けられない
     * @param {GroupItem} group 罫線グループ
     * @param {array} bounds 対象テキストの境界 [left, top, right, bottom]
     * @return {boolean} 対象テキストに引いた罫線かどうか
     */
    function coversTextCenter(group, bounds) {
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        var groupBounds = group.geometricBounds;

        return centerX >= groupBounds[0] && centerX <= groupBounds[2] &&
            centerY <= groupBounds[1] && centerY >= groupBounds[3];
    }

    /**
     * 対象テキストに前回引いた罫線グループだけ削除する（ほかのテキストの罫線は残す）
     * @param {Layer} layer 対象のレイヤー
     * @param {string} groupName グループ名
     * @param {array} bounds 対象テキストの境界 [left, top, right, bottom]
     * @return {void}
     */
    function removeExistingGroup(layer, groupName, bounds) {
        for (var i = layer.groupItems.length - 1; i >= 0; i--) {
            var group = layer.groupItems[i];

            if (group.name === groupName && coversTextCenter(group, bounds)) {
                group.remove();
            }
        }
    }

    /**
     * ドキュメントのカラーモードに合わせたグレーの色をつくる
     * @param {Document} doc 対象のドキュメント
     * @param {number} grayPercent 濃度（%、100で黒）
     * @return {object} GrayColor または RGBColor
     */
    function createGrayColor(doc, grayPercent) {
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            var rgbLevel = Math.round(255 * (100 - grayPercent) / 100);
            var rgbColor = new RGBColor();
            rgbColor.red = rgbLevel;
            rgbColor.green = rgbLevel;
            rgbColor.blue = rgbLevel;
            return rgbColor;
        }

        var grayColor = new GrayColor();
        grayColor.gray = grayPercent;
        return grayColor;
    }

    /**
     * 設定に合わせて罫線と十字線の色をつくる
     * @param {Document} doc 対象のドキュメント
     * @param {object} settings ダイアログで決めた設定
     * @return {object} { rule, cross } 罫線と十字線の色
     */
    function createRuleColors(doc, settings) {
        return {
            rule: createGrayColor(doc, settings.strokeGray),
            cross: createGrayColor(doc, LAYOUT.crossGray)
        };
    }

    /**
     * 罫線を1本引く
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {array} from 始点 [x, y]
     * @param {array} to 終点 [x, y]
     * @param {number} strokeWidth 線幅（pt）
     * @param {array} dashes 破線パターン（実線は空配列）
     * @param {object} color 罫線の色
     * @param {StrokeCap} strokeCap 線端（省略可）
     * @return {void}
     */
    function drawRule(group, from, to, strokeWidth, dashes, color, strokeCap) {
        var rulePath = group.pathItems.add();
        rulePath.setEntirePath([from, to]);

        rulePath.filled = false;
        rulePath.stroked = true;
        rulePath.strokeWidth = strokeWidth;
        rulePath.strokeDashes = dashes;
        rulePath.strokeColor = color;

        if (strokeCap) {
            rulePath.strokeCap = strokeCap;
        }
    }

    /**
     * 十字線の分割数を0か3以上の奇数に丸める
     * 偶数だとマスの中央が間隔になり、十字線が交差しない
     * @param {number} value 入力値
     * @param {boolean} roundUp 大きい側へ丸めるか
     * @return {number} 0、または3以上の奇数
     */
    function normalizeCrossSegments(value, roundUp) {
        var segments = Math.round(value);

        if (isNaN(segments) || segments <= 0) {
            return 0;
        }

        if (segments < 3) {
            return roundUp ? 3 : 0;
        }

        if (segments % 2 === 1) {
            return segments;
        }

        return roundUp ? segments + 1 : segments - 1;
    }

    /**
     * マス1辺の長さに合わせて、十字線の破線パターンを求める
     * 線分と間隔は同じ長さ。分割数＝線分の本数で、両端が線分で終わるように配分する
     * @param {object} settings ダイアログで決めた設定
     * @param {number} cellSize マスの一辺（pt）
     * @return {array} 破線パターン（実線は空配列）
     */
    function getCrossDashes(settings, cellSize) {
        /* 分割数0は破線にしない（実線として引く）/ Zero segments draws a solid line */
        if (settings.crossStyle !== "dashed" || settings.crossSegments < 1) {
            return [];
        }

        /* 線分が n 本、間隔が n−1 回で、合わせてマス1辺 / n dashes and n-1 gaps fill one cell */
        var dash = cellSize / (settings.crossSegments * 2 - 1);

        return [dash, dash];
    }

    /**
     * ブロック1つ分の各マスの中央に十字線を引く
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} rect ブロックの矩形 { left, right, top, bottom }
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} color 十字線の色
     * @return {void}
     */
    function drawCrosshairs(group, rect, grid, settings, color) {
        var cellSize = grid.cellSize;
        var strokeWidth = mmToPt(LAYOUT.crossStrokeMM);
        var dashes = getCrossDashes(settings, cellSize);

        var columnCount = Math.round((rect.right - rect.left) / cellSize);
        var rowCount = Math.round((rect.top - rect.bottom) / cellSize);

        for (var column = 0; column < columnCount; column++) {
            var cellLeft = rect.left + cellSize * column;
            var cellRight = cellLeft + cellSize;
            var centerX = cellLeft + cellSize / 2;

            for (var i = 0; i < rowCount; i++) {
                var cellTop = rect.top - cellSize * i;
                var cellBottom = cellTop - cellSize;
                var centerY = cellTop - cellSize / 2;

                drawRule(group, [centerX, cellTop], [centerX, cellBottom], strokeWidth, dashes, color);
                drawRule(group, [cellLeft, centerY], [cellRight, centerY], strokeWidth, dashes, color);
            }
        }
    }

    /**
     * 文字の前後に足すマスを反映した寸法を返す（書字方向の手前へ原点をずらす）
     * @param {object} metrics 実測した寸法
     * @param {object} settings ダイアログで決めた設定
     * @return {object} 前後にマスを足した寸法
     */
    function expandGrid(metrics, settings) {
        var startShift = metrics.cellSize * settings.extraCells;

        return {
            isVertical: metrics.isVertical,
            originX: metrics.isVertical ? metrics.originX : metrics.originX - startShift,
            originY: metrics.isVertical ? metrics.originY + startShift : metrics.originY,
            cellSize: metrics.cellSize,
            cellCount: metrics.cellCount + settings.extraCells * 2,
            lineCount: metrics.lineCount
        };
    }

    /**
     * 罫線を引く行の位置を並べる
     * テキストの行と、その外側に足す行を、1行ぶんの送りで等間隔に置く
     * 値は、行が並ぶ向きに測った原点からの距離（pt）
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @return {array} 行の位置（pt）の配列
     */
    function buildGridBlocks(grid, settings) {
        var linePositions = [];
        var linePitch = getLinePitch(grid.cellSize);

        for (var i = 0; i < grid.lineCount; i++) {
            linePositions.push(linePitch * i);
        }

        for (var j = 1; j <= settings.extraLines; j++) {
            linePositions.push(-linePitch * j);
            linePositions.push(linePitch * (grid.lineCount - 1 + j));
        }

        return linePositions;
    }

    /**
     * 1行ぶんの矩形を求める
     * @param {object} grid 前後のマスを足した寸法
     * @param {number} linePos 行が並ぶ向きに測った原点からの距離（pt）
     * @return {object} { left, right, top, bottom }
     */
    function getBlockRect(grid, linePos) {
        var runSize = grid.cellSize * grid.cellCount;

        if (grid.isVertical) {
            var left = grid.originX + linePos;

            return { left: left, right: left + grid.cellSize, top: grid.originY, bottom: grid.originY - runSize };
        }

        var top = grid.originY - linePos;

        return { left: grid.originX, right: grid.originX + runSize, top: top, bottom: top - grid.cellSize };
    }

    /**
     * すべての行を囲む矩形を求める
     * @param {object} grid 前後のマスを足した寸法
     * @param {array} linePositions 行の位置の配列
     * @return {object} { left, right, top, bottom }
     */
    function getBlocksRect(grid, linePositions) {
        var rect = getBlockRect(grid, linePositions[0]);

        for (var i = 1; i < linePositions.length; i++) {
            var blockRect = getBlockRect(grid, linePositions[i]);

            if (blockRect.left < rect.left) {
                rect.left = blockRect.left;
            }

            if (blockRect.right > rect.right) {
                rect.right = blockRect.right;
            }

            if (blockRect.top > rect.top) {
                rect.top = blockRect.top;
            }

            if (blockRect.bottom < rect.bottom) {
                rect.bottom = blockRect.bottom;
            }
        }

        return rect;
    }

    /**
     * マスの区切りを引く（縦組みでは横罫、横組みでは縦罫）
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} rect ブロックの矩形
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} color 罫線の色
     * @return {void}
     */
    function drawCellRules(group, rect, grid, settings, color) {
        var cellWidth = mmToPt(LAYOUT.cellStrokeMM);
        var emphasisWidth = mmToPt(LAYOUT.emphasisStrokeMM);

        var dashes = settings.dashedCellRules ?
            [mmToPt(LAYOUT.cellDashMM[0]), mmToPt(LAYOUT.cellDashMM[1])] : [];

        /* 誤差を溜めないよう、都度 原点から引いた位置を使う。
           太くする位置は、前後に足したマスではなく1文字目から数える
           / The emphasis is counted from the first character, not from the added cells */
        for (var i = 0; i <= grid.cellCount; i++) {
            var cellIndex = i - settings.extraCells;
            var isEmphasized = (settings.emphasisEvery > 0) && (cellIndex % settings.emphasisEvery === 0);
            var strokeWidth = isEmphasized ? emphasisWidth : cellWidth;

            if (grid.isVertical) {
                var y = rect.top - grid.cellSize * i;

                drawRule(group, [rect.left, y], [rect.right, y], strokeWidth, dashes, color);
            } else {
                var x = rect.left + grid.cellSize * i;

                drawRule(group, [x, rect.top], [x, rect.bottom], strokeWidth, dashes, color);
            }
        }
    }

    /**
     * 行の両側の罫線を引く（縦組みでは縦罫、横組みでは横罫）。書字方向の前後へ伸ばす
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} rect ブロックの矩形
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} color 罫線の色
     * @return {void}
     */
    function drawLineRules(group, rect, grid, settings, color) {
        var strokeWidth = mmToPt(LAYOUT.lineStrokeMM);

        /* 突出線端にして、マスの角までしっかり届かせる / Projecting caps reach the corners of the cells */
        var strokeCap = StrokeCap.PROJECTINGENDCAP;

        if (grid.isVertical) {
            var top = rect.top + settings.extension;
            var bottom = rect.bottom - settings.extension;

            drawRule(group, [rect.left, top], [rect.left, bottom], strokeWidth, [], color, strokeCap);
            drawRule(group, [rect.right, top], [rect.right, bottom], strokeWidth, [], color, strokeCap);
            return;
        }

        var left = rect.left - settings.extension;
        var right = rect.right + settings.extension;

        drawRule(group, [left, rect.top], [right, rect.top], strokeWidth, [], color, strokeCap);
        drawRule(group, [left, rect.bottom], [right, rect.bottom], strokeWidth, [], color, strokeCap);
    }

    /**
     * いちばん外側の行の外へ、マスの1/4だけ離して罫線を添える
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} rect すべてのブロックを囲む矩形
     * @param {object} grid 前後のマスを足した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} color 罫線の色
     * @return {void}
     */
    function drawSideRules(group, rect, grid, settings, color) {
        var sideOffset = grid.cellSize * LAYOUT.sideRuleRatio;
        var strokeWidth = mmToPt(LAYOUT.lineStrokeMM);
        var strokeCap = StrokeCap.PROJECTINGENDCAP;

        if (grid.isVertical) {
            var top = rect.top + settings.extension;
            var bottom = rect.bottom - settings.extension;

            drawRule(group, [rect.left - sideOffset, top], [rect.left - sideOffset, bottom], strokeWidth, [], color, strokeCap);
            drawRule(group, [rect.right + sideOffset, top], [rect.right + sideOffset, bottom], strokeWidth, [], color, strokeCap);
            return;
        }

        var left = rect.left - settings.extension;
        var right = rect.right + settings.extension;

        drawRule(group, [left, rect.top + sideOffset], [right, rect.top + sideOffset], strokeWidth, [], color, strokeCap);
        drawRule(group, [left, rect.bottom - sideOffset], [right, rect.bottom - sideOffset], strokeWidth, [], color, strokeCap);
    }

    /**
     * 罫線をひとそろい引く
     * @param {GroupItem} group 罫線を入れるグループ
     * @param {object} metrics 実測した寸法
     * @param {object} settings ダイアログで決めた設定
     * @param {object} colors 罫線と十字線の色 { rule, cross }
     * @return {void}
     */
    function drawGrid(group, metrics, settings, colors) {
        var grid = expandGrid(metrics, settings);
        var linePositions = buildGridBlocks(grid, settings);

        for (var i = 0; i < linePositions.length; i++) {
            var rect = getBlockRect(grid, linePositions[i]);

            /* 十字線は罫線より先に引いて背面へ回す */
            if (settings.crossStyle !== "none") {
                drawCrosshairs(group, rect, grid, settings, colors.cross);
            }

            drawCellRules(group, rect, grid, settings, colors.rule);
            drawLineRules(group, rect, grid, settings, colors.rule);
        }

        if (settings.showSideRules) {
            drawSideRules(group, getBlocksRect(grid, linePositions), grid, settings, colors.rule);
        }
    }

    /**
     * レイヤーごと最背面へ送る
     * @param {Layer} layer 対象のレイヤー
     * @return {void}
     */
    function sendLayerToBack(layer) {
        try {
            layer.zOrder(ZOrderMethod.SENDTOBACK);
        } catch (zOrderError) {
            layer.move(layer.parent, ElementPlacement.PLACEATEND);
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * 対象テキストに前回引いた罫線グループを集める（プレビュー中は隠して重ねないため）
     * @param {Document} doc 対象のドキュメント
     * @param {string} layerName レイヤー名
     * @param {string} groupName グループ名
     * @param {array} bounds 対象テキストの境界 [left, top, right, bottom]
     * @return {array} 表示中の罫線グループの配列
     */
    function findExistingGroups(doc, layerName, groupName, bounds) {
        var existingGroups = [];

        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];

            if (layer.name !== layerName || layer.locked) {
                continue;
            }

            for (var j = 0; j < layer.groupItems.length; j++) {
                var group = layer.groupItems[j];

                if (group.name === groupName && !group.hidden && coversTextCenter(group, bounds)) {
                    existingGroups.push(group);
                }
            }
        }

        return existingGroups;
    }

    /**
     * プレビュー用のレイヤーが残っていれば消す
     * @param {Document} doc 対象のドキュメント
     * @return {void}
     */
    function removePreviewLayer(doc) {
        for (var i = doc.layers.length - 1; i >= 0; i--) {
            if (doc.layers[i].name === PREVIEW_LAYER_NAME) {
                doc.layers[i].remove();
            }
        }
    }

    /**
     * 原本と複製の表示を入れ替える（プレビュー中は前回の罫線も隠す）
     * @param {object} previewState beginPreview が返した状態
     * @param {boolean} visible プレビューを見せるかどうか
     * @return {void}
     */
    function togglePreviewItems(previewState, visible) {
        previewState.textFrame.hidden = visible;
        previewState.previewFrame.hidden = !visible;

        for (var i = 0; i < previewState.existingGroups.length; i++) {
            previewState.existingGroups[i].hidden = visible;
        }
    }

    /**
     * プレビューの下ごしらえ（原本は隠し、確定後と同じ状態にした複製を代わりに見せる）
     * @param {Document} doc 対象のドキュメント
     * @param {TextFrame} textFrame 原本のテキストフレーム
     * @return {object} 片付けに使う状態
     */
    function beginPreview(doc, textFrame) {
        var previewState = {
            doc: doc,
            textFrame: textFrame,
            activeLayer: doc.activeLayer,
            existingGroups: findExistingGroups(doc, getLabel(LABELS.itemName.layerName), getLabel(LABELS.itemName.groupName), textFrame.geometricBounds),
            isVertical: isVerticalText(textFrame),
            previewFrame: textFrame.duplicate()
        };

        togglePreviewItems(previewState, true);

        return previewState;
    }

    /**
     * プレビュー用の複製を今の設定に合わせ直し、寸法を測り直す
     * @param {object} previewContext プレビューに使う { previewState, appliedShape, metrics }
     * @param {object} settings ダイアログで決めた設定
     * @return {void}
     */
    function reshapePreviewFrame(previewContext, settings) {
        var previewState = previewContext.previewState;

        if (previewContext.appliedShape.adjustSize === settings.adjustSize && previewContext.appliedShape.scale === settings.scale) {
            return;
        }

        /* 文字属性は元へ戻せないので、比率を変えるたびに複製を作り直す
           / Character attributes cannot be rolled back, so rebuild the copy whenever the scaling changes */
        previewState.previewFrame.remove();
        previewState.previewFrame = previewState.textFrame.duplicate();
        previewState.previewFrame.hidden = false;

        var fitResult = fitTextToGrid(previewState.previewFrame, settings);

        if (fitResult.metrics) {
            previewContext.metrics = fitResult.metrics;
        }

        previewContext.appliedShape = { adjustSize: settings.adjustSize, scale: settings.scale };
    }

    /**
     * プレビューを片付けて元の状態に戻す（画面の更新はスクリプト終了時に任せる）
     * @param {object} previewState beginPreview が返した状態
     * @return {void}
     */
    function endPreview(previewState) {
        removePreviewLayer(previewState.doc);
        togglePreviewItems(previewState, false);
        previewState.previewFrame.remove();

        /* プレビューでレイヤーと選択が変わるので、控えておいた状態に戻す */
        previewState.doc.activeLayer = previewState.activeLayer;
        previewState.doc.selection = [previewState.textFrame];
    }

    /**
     * 現在の設定で罫線をプレビューする
     * @param {object} previewContext プレビューに使う { previewState, appliedShape, metrics }
     * @param {object} settings ダイアログで決めた設定
     * @return {void}
     */
    function drawPreviewGrid(previewContext, settings) {
        var doc = previewContext.previewState.doc;

        togglePreviewItems(previewContext.previewState, true);
        reshapePreviewFrame(previewContext, settings);
        removePreviewLayer(doc);

        if (!previewContext.metrics) {
            app.redraw();
            return;
        }

        var previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;

        drawGrid(previewLayer.groupItems.add(), previewContext.metrics, settings, createRuleColors(doc, settings));

        /* 本処理と同じく、テキストの下に回す */
        sendLayerToBack(previewLayer);

        app.redraw();
    }

    /**
     * プレビューを消して、実行前の見た目に戻す
     * @param {object} previewContext プレビューに使う { previewState, appliedShape, metrics }
     * @return {void}
     */
    function clearPreviewGrid(previewContext) {
        removePreviewLayer(previewContext.previewState.doc);
        togglePreviewItems(previewContext.previewState, false);

        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 数値入力欄を↑↓キーで増減できるようにする（Shiftで10の倍数にスナップ）
     * @param {object} editText 対象の edittext
     * @param {number} minimum 下限値
     * @param {function} onUpdate 値を変えたあとに呼ぶ処理（省略可）
     * @param {function} normalize 増減後の値を丸める処理（省略可。値と増加方向を受け取る）
     * @return {void}
     */
    function changeValueByArrowKey(editText, minimum, onUpdate, normalize) {
        editText.addEventListener("keydown", function (event) {
            var fieldValue = Number(editText.text);
            if (isNaN(fieldValue)) return;

            var keyboardState = ScriptUI.environment.keyboardState;

            if (event.keyName == "Up" || event.keyName == "Down") {
                var isUp = event.keyName == "Up";
                var delta = 1;

                if (keyboardState.shiftKey) {
                    fieldValue = Math.floor(fieldValue / 10) * 10;
                    delta = 10;
                }

                fieldValue += isUp ? delta : -delta;
                if (fieldValue < minimum) fieldValue = minimum;

                if (normalize) {
                    fieldValue = normalize(fieldValue, isUp);
                }

                event.preventDefault();
                editText.text = fieldValue;

                /* 値の代入では onChanging が呼ばれないため、ここでプレビューを更新する */
                if (onUpdate) {
                    onUpdate();
                }
            }
        });
    }

    /**
     * 濃度をK表記にする
     * @param {number} grayPercent 濃度（%、100で黒）
     * @return {string} 表示用の文字列
     */
    function formatGray(grayPercent) {
        return "K" + Math.round(grayPercent);
    }

    /**
     * 項目名のラベルを1つ追加する
     * @param {object} parent 追加先のグループ
     * @param {object} labelEntry ja / en を持つラベル定義
     * @return {object} 追加した statictext
     */
    function addRowLabel(parent, labelEntry) {
        var label = parent.add("statictext", undefined, labelText(labelEntry));
        label.justify = "right";
        label.preferredSize.width = UI.labelWidth;
        return label;
    }

    /**
     * ダイアログのパネルを1つ追加する
     * @param {object} parent 追加先のウィンドウ
     * @param {object} labelEntry ja / en を持つラベル定義
     * @return {object} 追加したパネル
     */
    function addPanel(parent, labelEntry) {
        var panel = parent.add("panel", undefined, getLabel(labelEntry));
        panel.orientation = "column";
        panel.alignment = ["fill", "top"];
        panel.alignChildren = ["fill", "top"];
        panel.margins = UI.panelMargins;
        panel.spacing = UI.panelSpacing;
        return panel;
    }

    /**
     * ダイアログの1行分のグループを追加する
     * @param {object} parent 追加先のパネルまたはウィンドウ
     * @return {object} 追加したグループ
     */
    function addRow(parent) {
        var row = parent.add("group");
        row.orientation = "row";
        row.alignment = ["fill", "center"];
        row.alignChildren = ["left", "center"];
        return row;
    }

    /**
     * 複数のコントロールに同じツールチップを付ける
     * @param {array} targets 対象のコントロール
     * @param {object} tooltipEntry ツールチップのラベル定義（不要なときは null）
     * @return {void}
     */
    function setTooltip(targets, tooltipEntry) {
        if (!tooltipEntry) {
            return;
        }

        var tooltip = getLabel(tooltipEntry);

        for (var i = 0; i < targets.length; i++) {
            targets[i].helpTip = tooltip;
        }
    }

    /**
     * 項目名・数値入力欄・単位をまとめた行を追加する
     * @param {object} parent 追加先のパネル
     * @param {object} labelEntry ja / en を持つラベル定義
     * @param {number} initialValue 入力欄の初期値
     * @param {object} unitEntry 単位のラベル定義（単位がないときは null）
     * @param {object} tooltipEntry ツールチップのラベル定義（不要なときは null）
     * @return {object} { row, input } 追加した行と入力欄
     */
    function addNumberRow(parent, labelEntry, initialValue, unitEntry, tooltipEntry) {
        var row = addRow(parent);
        var label = addRowLabel(row, labelEntry);

        var input = row.add("edittext", undefined, String(initialValue));
        input.characters = 5;

        if (unitEntry) {
            row.add("statictext", undefined, getLabel(unitEntry));
        }

        setTooltip([label, input], tooltipEntry);

        return { row: row, input: input };
    }

    /**
     * 項目名の幅だけ字下げしたチェックボックスの行を追加する
     * @param {object} parent 追加先のパネル
     * @param {object} labelEntry ja / en を持つラベル定義
     * @param {boolean} initialValue 初期値
     * @param {object} tooltipEntry ツールチップのラベル定義（不要なときは null）
     * @return {object} 追加した checkbox
     */
    function addIndentedCheckbox(parent, labelEntry, initialValue, tooltipEntry) {
        var row = addRow(parent);
        row.add("statictext", undefined, "").preferredSize.width = UI.labelWidth;

        var checkbox = row.add("checkbox", undefined, getLabel(labelEntry));
        checkbox.value = initialValue;

        setTooltip([checkbox], tooltipEntry);

        return checkbox;
    }

    /**
     * 項目名と排他ラジオをまとめた行を追加する
     * ラジオは同じ親の中でしか排他にならないため、1つのグループへまとめて入れる
     * @param {object} parent 追加先のパネル
     * @param {object} labelEntry ja / en を持つラベル定義
     * @param {array} radioEntries ラジオのラベル定義の配列
     * @param {number} selectedIndex 最初に選ぶラジオの位置
     * @return {object} { row, radios } 追加した行とラジオの配列
     */
    function addRadioRow(parent, labelEntry, radioEntries, selectedIndex) {
        var row = addRow(parent);
        addRowLabel(row, labelEntry);

        var radioGroup = row.add("group");
        radioGroup.orientation = "row";
        radioGroup.alignChildren = ["left", "center"];

        var radios = [];

        for (var i = 0; i < radioEntries.length; i++) {
            var radio = radioGroup.add("radiobutton", undefined, getLabel(radioEntries[i]));
            radio.value = (i === selectedIndex);
            radios.push(radio);
        }

        return { row: row, radios: radios };
    }

    /**
     * 入力欄の数値を読む（読めないときや下限を下回るときは既定値に寄せる）
     * @param {object} editText 対象の edittext
     * @param {number} minimum 下限値
     * @param {number} fallback 読めないときの値
     * @param {boolean} asInteger 整数に丸めるか
     * @return {number} 読み取った値
     */
    function readNumberField(editText, minimum, fallback, asInteger) {
        var value = Number(editText.text);

        if (isNaN(value)) {
            return fallback;
        }

        if (asInteger) {
            value = Math.round(value);
        }

        return (value < minimum) ? fallback : value;
    }

    /**
     * プリセット一覧をドロップダウンへ流し込む
     * @param {object} presetDropdown 対象のドロップダウン
     * @return {void}
     */
    function fillPresetDropdown(presetDropdown) {
        presetDropdown.removeAll();

        for (var i = 0; i < PRESETS.length; i++) {
            presetDropdown.add("item", getLabel(PRESETS[i].label))._preset = PRESETS[i];
        }
    }

    /**
     * 値をソース表記にする（プリセットのコード出力用）
     * @param {number|boolean|string} value 設定値
     * @return {string} ソース表記
     */
    function formatCodeValue(value) {
        if (typeof value === "string") {
            return '"' + value + '"';
        }

        if (typeof value === "boolean") {
            return value ? "true" : "false";
        }

        return String(value);
    }

    /**
     * プリセットを PRESETS へ貼り付けられるコードに変換する
     * @param {string} presetName プリセット名
     * @param {object} values プリセットの設定値
     * @return {string} 組み込み用のコード
     */
    function presetToCode(presetName, values) {
        var valueLines = [];

        for (var key in values) {
            if (values.hasOwnProperty(key)) {
                valueLines.push("                " + key + ": " + formatCodeValue(values[key]));
            }
        }

        return ",\n        {\n" +
            '            label: { ja: "' + presetName + '", en: "' + presetName + '" },\n' +
            "            settings: {\n" + valueLines.join(",\n") + "\n            }\n        }";
    }

    /**
     * プリセットを選ぶ行（ドロップダウンと［保存］ボタン）を追加する
     * @param {object} dialog 対象のウィンドウ
     * @return {object} { presetDropdown, saveButton }
     */
    function addPresetRow(dialog) {
        var presetRow = addRow(dialog);
        var presetLabel = addRowLabel(presetRow, LABELS.label.preset);

        var presetDropdown = presetRow.add("dropdownlist");
        presetDropdown.alignment = ["fill", "center"];

        fillPresetDropdown(presetDropdown);
        presetDropdown.selection = 0;

        var saveButton = presetRow.add("button", undefined, getLabel(LABELS.button.save));
        saveButton.preferredSize.width = UI.saveButtonWidth;

        setTooltip([presetLabel, presetDropdown], LABELS.tooltip.preset);
        setTooltip([saveButton], LABELS.tooltip.savePreset);

        return { presetDropdown: presetDropdown, saveButton: saveButton };
    }

    /**
     * ［全体］パネルを追加する
     * @param {object} dialog 対象のウィンドウ
     * @param {boolean} isVertical 縦組みかどうか
     * @return {object} { colorSlider, colorValueLabel, extraCellsInput, extraLinesInput }
     */
    function addOverallPanel(dialog, isVertical) {
        var overallPanel = addPanel(dialog, LABELS.panel.overall);

        var colorRow = addRow(overallPanel);
        var colorLabel = addRowLabel(colorRow, LABELS.label.lineColor);

        var colorSlider = colorRow.add("slider", undefined, DEFAULTS.strokeGray, 0, 100);
        colorSlider.alignment = ["fill", "center"];

        /* 幅は作成時の文字列で確保しておく / Reserve the width with the longest text */
        var colorValueLabel = colorRow.add("statictext", undefined, formatGray(100));
        colorValueLabel.text = formatGray(DEFAULTS.strokeGray);

        setTooltip([colorLabel, colorSlider], LABELS.tooltip.lineColor);

        var extraCellsRow = addNumberRow(overallPanel, LABELS.label.extraCells, DEFAULTS.extraCells, null, LABELS.tooltip.extraCells);
        var extraLinesRow = addNumberRow(overallPanel, orientedEntry(LABELS.label.extraLines, isVertical), DEFAULTS.extraLines, null, LABELS.tooltip.extraLines);

        return {
            colorSlider: colorSlider,
            colorValueLabel: colorValueLabel,
            extraCellsInput: extraCellsRow.input,
            extraLinesInput: extraLinesRow.input
        };
    }

    /**
     * ［文字］パネルを追加する
     * @param {object} dialog 対象のウィンドウ
     * @return {object} { adjustSizeCheckbox, scaleRow, scaleInput }
     */
    function addTextPanel(dialog) {
        var textPanel = addPanel(dialog, LABELS.panel.text);

        var adjustSizeCheckbox = addIndentedCheckbox(textPanel, LABELS.checkbox.adjustSize, DEFAULTS.adjustSize, LABELS.tooltip.adjustSize);
        var scaleRow = addNumberRow(textPanel, LABELS.label.scale, DEFAULTS.scale, LABELS.unit.percent, LABELS.tooltip.scale);

        return {
            adjustSizeCheckbox: adjustSizeCheckbox,
            scaleRow: scaleRow.row,
            scaleInput: scaleRow.input
        };
    }

    /**
     * 行の両側の罫線のパネルを追加する（縦組みでは［縦罫］、横組みでは［横罫］）
     * @param {object} dialog 対象のウィンドウ
     * @param {boolean} isVertical 縦組みかどうか
     * @param {number} defaultExtensionMM 伸張の既定値（mm）
     * @return {object} { extensionInput, sideRuleCheckbox }
     */
    function addLineRulePanel(dialog, isVertical, defaultExtensionMM) {
        var linePanel = addPanel(dialog, orientedEntry(LABELS.panel.lineRule, isVertical));

        var extensionRow = addNumberRow(linePanel, LABELS.label.extension, defaultExtensionMM, LABELS.unit.mm, LABELS.tooltip.extension);
        var sideRuleCheckbox = addIndentedCheckbox(linePanel, LABELS.checkbox.sideRule, DEFAULTS.showSideRules, LABELS.tooltip.sideRule);

        return {
            extensionInput: extensionRow.input,
            sideRuleCheckbox: sideRuleCheckbox
        };
    }

    /**
     * マスの区切りのパネルを追加する（縦組みでは［横罫］、横組みでは［縦罫］）
     * @param {object} dialog 対象のウィンドウ
     * @param {boolean} isVertical 縦組みかどうか
     * @return {object} { cellSolidRadio, cellDashedRadio, emphasisCheckbox, emphasisInput }
     */
    function addCellRulePanel(dialog, isVertical) {
        var cellPanel = addPanel(dialog, orientedEntry(LABELS.panel.cellRule, isVertical));

        var styleRow = addRadioRow(cellPanel, LABELS.label.lineStyle,
            [LABELS.radio.solid, LABELS.radio.dashed], DEFAULTS.dashedCellRules ? 1 : 0);

        var emphasisRow = addRow(cellPanel);
        emphasisRow.add("statictext", undefined, "").preferredSize.width = UI.labelWidth;

        var emphasisCheckbox = emphasisRow.add("checkbox", undefined, "");
        emphasisCheckbox.value = DEFAULTS.emphasisEnabled;

        var emphasisPrefix = getLabel(LABELS.label.emphasisPrefix);

        if (emphasisPrefix.length > 0) {
            emphasisRow.add("statictext", undefined, emphasisPrefix);
        }

        var emphasisInput = emphasisRow.add("edittext", undefined, String(DEFAULTS.emphasisEvery));
        emphasisInput.characters = 3;

        emphasisRow.add("statictext", undefined, getLabel(LABELS.label.emphasisSuffix));
        setTooltip([emphasisCheckbox, emphasisInput], LABELS.tooltip.emphasis);

        return {
            cellSolidRadio: styleRow.radios[0],
            cellDashedRadio: styleRow.radios[1],
            emphasisCheckbox: emphasisCheckbox,
            emphasisInput: emphasisInput
        };
    }

    /**
     * ［十字線］パネルを追加する
     * @param {object} dialog 対象のウィンドウ
     * @return {object} { crossNoneRadio, crossSolidRadio, crossDashedRadio, crossSegmentsRow, crossSegmentsInput }
     */
    function addCrossPanel(dialog) {
        var crossPanel = addPanel(dialog, LABELS.panel.cross);

        var crossStyleRow = addRadioRow(crossPanel, LABELS.label.lineStyle,
            [LABELS.radio.none, LABELS.radio.solid, LABELS.radio.dashed], crossStyleIndex(DEFAULTS.crossStyle));

        var crossSegmentsRow = addNumberRow(crossPanel, LABELS.label.segments, DEFAULTS.crossSegments, null, LABELS.tooltip.crossSegments);

        return {
            crossNoneRadio: crossStyleRow.radios[0],
            crossSolidRadio: crossStyleRow.radios[1],
            crossDashedRadio: crossStyleRow.radios[2],
            crossSegmentsRow: crossSegmentsRow.row,
            crossSegmentsInput: crossSegmentsRow.input
        };
    }

    /**
     * ボタンエリア（左にプレビュー、右にキャンセル・OK）を追加する
     * @param {object} dialog 対象のウィンドウ
     * @return {object} 追加したプレビューの checkbox
     */
    function addButtonRow(dialog) {
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["left", "center"];
        btnRowGroup.alignment = ["fill", "center"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignment = ["left", "center"];
        btnLeftGroup.alignChildren = ["left", "center"];

        var previewCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.preview));
        previewCheckbox.value = DEFAULTS.preview;
        setTooltip([previewCheckbox], LABELS.tooltip.preview);

        /* スペーサー（右側のボタンを押し出す） */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });

        return previewCheckbox;
    }

    /**
     * ダイアログのコントロールを組み立てる
     * @param {object} dialog 対象のウィンドウ
     * @param {boolean} isVertical 縦組みかどうか
     * @param {number} defaultExtensionMM 伸張の既定値（mm）
     * @return {object} コントロール一式
     */
    function buildDialogControls(dialog, isVertical, defaultExtensionMM) {
        var preset = addPresetRow(dialog);
        var overall = addOverallPanel(dialog, isVertical);
        var text = addTextPanel(dialog);
        var lineRule = addLineRulePanel(dialog, isVertical, defaultExtensionMM);
        var cellRule = addCellRulePanel(dialog, isVertical);
        var cross = addCrossPanel(dialog);
        var previewCheckbox = addButtonRow(dialog);

        return {
            presetDropdown: preset.presetDropdown,
            saveButton: preset.saveButton,
            colorSlider: overall.colorSlider,
            colorValueLabel: overall.colorValueLabel,
            extraCellsInput: overall.extraCellsInput,
            extraLinesInput: overall.extraLinesInput,
            adjustSizeCheckbox: text.adjustSizeCheckbox,
            scaleRow: text.scaleRow,
            scaleInput: text.scaleInput,
            extensionInput: lineRule.extensionInput,
            sideRuleCheckbox: lineRule.sideRuleCheckbox,
            cellSolidRadio: cellRule.cellSolidRadio,
            cellDashedRadio: cellRule.cellDashedRadio,
            emphasisCheckbox: cellRule.emphasisCheckbox,
            emphasisInput: cellRule.emphasisInput,
            crossNoneRadio: cross.crossNoneRadio,
            crossSolidRadio: cross.crossSolidRadio,
            crossDashedRadio: cross.crossDashedRadio,
            crossSegmentsRow: cross.crossSegmentsRow,
            crossSegmentsInput: cross.crossSegmentsInput,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * 十字線の線種に対応するラジオの位置を返す
     * @param {string} crossStyle 線種（none / solid / dashed）
     * @return {number} ラジオの位置
     */
    function crossStyleIndex(crossStyle) {
        if (crossStyle === "none") {
            return 0;
        }

        return (crossStyle === "solid") ? 1 : 2;
    }

    /**
     * チェックやラジオの状態に合わせて、入力欄の有効・無効を切り替える
     * @param {object} controls コントロール一式
     * @return {void}
     */
    function updateDialogState(controls) {
        controls.scaleRow.enabled = controls.adjustSizeCheckbox.value;
        controls.emphasisInput.enabled = controls.emphasisCheckbox.value;
        controls.crossSegmentsRow.enabled = controls.crossDashedRadio.value;
    }

    /**
     * ダイアログの入力内容を、プリセットと同じ形で読む
     * @param {object} controls コントロール一式
     * @return {object} プリセットの設定値
     */
    function readPresetValues(controls) {
        return {
            strokeGray: Math.round(controls.colorSlider.value),
            extraCells: readNumberField(controls.extraCellsInput, 0, 0, true),
            extraLines: readNumberField(controls.extraLinesInput, 0, 0, true),
            adjustSize: controls.adjustSizeCheckbox.value,
            scale: readNumberField(controls.scaleInput, 1, DEFAULTS.scale, false),
            extensionMM: readNumberField(controls.extensionInput, 0, 0, false),
            dashedCellRules: controls.cellDashedRadio.value,
            emphasisEnabled: controls.emphasisCheckbox.value,
            emphasisEvery: readNumberField(controls.emphasisInput, 1, 0, true),
            showSideRules: controls.sideRuleCheckbox.value,
            crossStyle: controls.crossNoneRadio.value ? "none" : (controls.crossSolidRadio.value ? "solid" : "dashed"),
            crossSegments: normalizeCrossSegments(Number(controls.crossSegmentsInput.text), true)
        };
    }

    /**
     * ダイアログの入力内容を、罫線を引くための設定にまとめる
     * @param {object} controls コントロール一式
     * @return {object} 罫線を引くための設定
     */
    function readDialogSettings(controls) {
        var values = readPresetValues(controls);

        return {
            strokeGray: values.strokeGray,
            extraCells: values.extraCells,
            extraLines: values.extraLines,
            adjustSize: values.adjustSize,
            scale: values.scale,
            extension: mmToPt(values.extensionMM),
            dashedCellRules: values.dashedCellRules,
            emphasisEvery: values.emphasisEnabled ? values.emphasisEvery : 0,
            showSideRules: values.showSideRules,
            crossStyle: values.crossStyle,
            crossSegments: values.crossSegments
        };
    }

    /**
     * プリセットの値をダイアログへ流し込む
     * @param {object} controls コントロール一式
     * @param {object} values プリセットの設定値
     * @param {number} defaultExtensionMM 伸張の既定値（mm）
     * @return {void}
     */
    function applyPresetValues(controls, values, defaultExtensionMM) {
        controls.colorSlider.value = values.strokeGray;
        controls.colorValueLabel.text = formatGray(values.strokeGray);
        controls.extraCellsInput.text = String(values.extraCells);
        controls.extraLinesInput.text = String(values.extraLines);

        controls.adjustSizeCheckbox.value = values.adjustSize;
        controls.scaleInput.text = String(values.scale);

        /* null は「文字サイズの1/4」/ null means a quarter of the font size */
        controls.extensionInput.text = String((values.extensionMM === null) ? defaultExtensionMM : values.extensionMM);
        controls.sideRuleCheckbox.value = values.showSideRules;

        controls.cellSolidRadio.value = !values.dashedCellRules;
        controls.cellDashedRadio.value = values.dashedCellRules;

        controls.emphasisCheckbox.value = values.emphasisEnabled;
        controls.emphasisInput.text = String(values.emphasisEvery);

        controls.crossNoneRadio.value = (values.crossStyle === "none");
        controls.crossSolidRadio.value = (values.crossStyle === "solid");
        controls.crossDashedRadio.value = (values.crossStyle === "dashed");
        controls.crossSegmentsInput.text = String(values.crossSegments);

        updateDialogState(controls);
    }

    /**
     * ダイアログのイベントをつなぐ
     * @param {object} controls コントロール一式
     * @param {object} callbacks { onPresetSelected, onPresetSaved, onSettingChanged, onPreviewToggled }
     * @return {void}
     */
    function bindDialogHandlers(controls, callbacks) {
        /* チェックやラジオは、ディム状態を直してからプレビューを更新する */
        function onToggle() {
            updateDialogState(controls);
            callbacks.onSettingChanged();
        }

        controls.presetDropdown.onChange = callbacks.onPresetSelected;
        controls.saveButton.onClick = callbacks.onPresetSaved;

        controls.colorSlider.onChanging = function () {
            var grayPercent = Math.round(controls.colorSlider.value);

            /* shiftを押している間は10%刻みにする / Shift snaps the slider to steps of 10% */
            if (ScriptUI.environment.keyboardState.shiftKey) {
                grayPercent = Math.round(grayPercent / 10) * 10;
                controls.colorSlider.value = grayPercent;
            }

            controls.colorValueLabel.text = formatGray(grayPercent);
        };

        /* ドラッグ中に引き直すと重いので、離したタイミングで更新する / Redraw once the drag ends */
        controls.colorSlider.onChange = callbacks.onSettingChanged;

        var numberFields = [
            [controls.extraCellsInput, 0],
            [controls.extraLinesInput, 0],
            [controls.scaleInput, 1],
            [controls.extensionInput, 0],
            [controls.emphasisInput, 1]
        ];

        for (var i = 0; i < numberFields.length; i++) {
            numberFields[i][0].onChanging = callbacks.onSettingChanged;
            changeValueByArrowKey(numberFields[i][0], numberFields[i][1], callbacks.onSettingChanged);
        }

        var toggles = [
            controls.adjustSizeCheckbox,
            controls.sideRuleCheckbox,
            controls.emphasisCheckbox,
            controls.cellSolidRadio,
            controls.cellDashedRadio,
            controls.crossNoneRadio,
            controls.crossSolidRadio,
            controls.crossDashedRadio
        ];

        for (var j = 0; j < toggles.length; j++) {
            toggles[j].onClick = onToggle;
        }

        controls.crossSegmentsInput.onChanging = callbacks.onSettingChanged;
        changeValueByArrowKey(controls.crossSegmentsInput, 0, callbacks.onSettingChanged, normalizeCrossSegments);

        /* 入力欄を離れたら、実際に使う値（0か奇数）に揃える / Snap the field to the value actually used */
        controls.crossSegmentsInput.onChange = function () {
            controls.crossSegmentsInput.text = String(normalizeCrossSegments(Number(controls.crossSegmentsInput.text), true));
            callbacks.onSettingChanged();
        };

        controls.previewCheckbox.onClick = callbacks.onPreviewToggled;
    }

    /**
     * 設定ダイアログを表示する
     * @param {object} previewContext プレビューに使う { previewState, appliedShape, metrics }
     * @return {object|null} 設定値。キャンセルしたときは null
     */
    function showSettingsDialog(previewContext) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = 15;
        dialog.spacing = 10;

        var defaultExtensionMM = getDefaultExtensionMM(previewContext.fontSize);
        var controls = buildDialogControls(dialog, previewContext.previewState.isVertical, defaultExtensionMM);

        /* 前回の設定があれば、ハンドラをつなぐ前に流し込む（onChangeを呼ばせない）
           / Restore the last settings before the handlers are wired, so no events fire */
        var session = loadSession();

        if (session) {
            applyPresetValues(controls, session.values, defaultExtensionMM);

            if (session.presetIndex < controls.presetDropdown.items.length) {
                controls.presetDropdown.selection = session.presetIndex;
            }
        }

        /**
         * プレビューの表示・非表示を現在の設定に合わせる
         * @return {void}
         */
        function refreshPreview() {
            if (controls.previewCheckbox.value) {
                drawPreviewGrid(previewContext, readDialogSettings(controls));
            } else {
                clearPreviewGrid(previewContext);
            }
        }

        /**
         * 手で値を変えたら、プリセットの選択を「カスタム」へ戻してプレビューを更新する
         * @return {void}
         */
        function onSettingChanged() {
            /* カスタムは末尾に置いてある / Custom is the last item */
            controls.presetDropdown.selection = controls.presetDropdown.items.length - 1;
            refreshPreview();
        }

        /**
         * 選んだプリセットをダイアログへ流し込む
         * @return {void}
         */
        function onPresetSelected() {
            var preset = controls.presetDropdown.selection ? controls.presetDropdown.selection._preset : null;

            if (!preset || !preset.settings) {
                return;
            }

            applyPresetValues(controls, preset.settings, defaultExtensionMM);
            refreshPreview();
        }

        /**
         * いまの設定に名前を付けてプリセットへ加える（残すためのコードは通知で出す）
         * @return {void}
         */
        function onPresetSaved() {
            var currentName = controls.presetDropdown.selection ? controls.presetDropdown.selection.text : "";
            var presetName = prompt(getLabel(LABELS.alert.presetName), currentName);

            if (presetName === null) {
                return;
            }

            presetName = String(presetName).replace(/^\s+|\s+$/g, "");

            if (!presetName) {
                return;
            }

            var values = readPresetValues(controls);

            /* 「カスタム」は末尾に置いたままにする / Keep Custom at the end of the list */
            PRESETS.splice(PRESETS.length - 1, 0, { label: { ja: presetName, en: presetName }, settings: values });

            fillPresetDropdown(controls.presetDropdown);
            controls.presetDropdown.selection = PRESETS.length - 2;

            alert(getLabel(LABELS.alert.presetSaved) + presetToCode(presetName, values));
        }

        bindDialogHandlers(controls, {
            onPresetSelected: onPresetSelected,
            onPresetSaved: onPresetSaved,
            onSettingChanged: onSettingChanged,
            onPreviewToggled: refreshPreview
        });

        /* プレビューの初期状態をダイアログを開く前に反映しておく */
        updateDialogState(controls);
        refreshPreview();

        var dialogResult = dialog.show();

        if (dialogResult !== 1) {
            return null;
        }

        var selectedPreset = controls.presetDropdown.selection;
        saveSession(readPresetValues(controls), selectedPreset ? selectedPreset.index : 0);

        return readDialogSettings(controls);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 決まった設定でテキストを整え、罫線を引く
     * @param {Document} doc 対象のドキュメント
     * @param {TextFrame} textFrame 対象のテキストフレーム
     * @param {object} settings ダイアログで決めた設定
     * @return {void}
     */
    function applyGridToText(doc, textFrame, settings) {
        var fitResult = fitTextToGrid(textFrame, settings);

        if (!fitResult.metrics) {
            alert(getLabel(LABELS.alert.noFontSize));
            return;
        }

        var layerName = getLabel(LABELS.itemName.layerName);
        var groupName = getLabel(LABELS.itemName.groupName);

        var ruleLayer = prepareRuleLayer(doc, layerName);
        removeExistingGroup(ruleLayer, groupName, textFrame.geometricBounds);

        var ruleGroup = ruleLayer.groupItems.add();
        ruleGroup.name = groupName;

        drawGrid(ruleGroup, fitResult.metrics, settings, createRuleColors(doc, settings));

        /* グループの zOrder だけではテキストの下に回らないため、レイヤーごと最背面へ送る */
        sendLayerToBack(ruleLayer);

        /* 成功時は通知しない。設定できなかった属性があるときだけ知らせる */
        if (fitResult.failed.length > 0) {
            alert(getLabel(LABELS.alert.skipped) + " " + fitResult.failed.join(", "));
        }
    }

    /**
     * 選択のエッジ表示を切り替える（プレビューを見やすくする）
     * @return {void}
     */
    function toggleSelectionEdges() {
        if (app.documents.length === 0) {
            return;
        }

        app.executeMenuCommand("edge");
    }

    /**
     * 選択したテキストに罫線を引く
     * @return {void}
     */
    function runGridMaker() {
        if (app.documents.length === 0) {
            alert(getLabel(LABELS.alert.noDocument));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (!selectedItems || selectedItems.length !== 1 || selectedItems[0].typename !== "TextFrame") {
            alert(getLabel(LABELS.alert.noTextFrame));
            return;
        }

        var textFrame = selectedItems[0];

        if (textFrame.textRange.characters.length === 0) {
            alert(getLabel(LABELS.alert.noCharacter));
            return;
        }

        /* 先頭文字の文字サイズを基準にするので、先に読めることを確かめる */
        var fontSize = textFrame.textRange.characters[0].characterAttributes.size;

        if (!fontSize || fontSize <= 0) {
            alert(getLabel(LABELS.alert.noFontSize));
            return;
        }

        /* 原本は触らず、確定後と同じ状態にした複製でプレビューする / Preview on a copy, leaving the original untouched */
        var previewState = beginPreview(doc, textFrame);

        var previewContext = {
            previewState: previewState,
            fontSize: fontSize,
            appliedShape: { adjustSize: null, scale: 0 }, /* 最初の描画で必ず作り直す / forces the first reshape */
            metrics: null
        };

        var settings = null;

        /* 途中で失敗しても原本を隠したままにしない / Never leave the original hidden if something fails */
        try {
            settings = showSettingsDialog(previewContext);
        } finally {
            endPreview(previewState);
        }

        if (!settings) {
            return;
        }

        applyGridToText(doc, textFrame, settings);
    }

    /**
     * エッジ表示を戻せるよう、本体を挟んで実行する
     * @return {void}
     */
    function main() {
        toggleSelectionEdges();

        try {
            runGridMaker();
        } finally {
            toggleSelectionEdges();
        }
    }

    main();
})();
