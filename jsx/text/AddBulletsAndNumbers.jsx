#target illustrator
#targetengine "addBulletsAndNumbers"
#include "ColorPicker.jsx"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要

- 選択したテキストフレームの各行頭に、箇条書き記号または連番を付与します。
- ダイアログを開くと現在の行頭マーカーから設定を推定し、プレビューを見ながら「現状の続き」として編集できます。
- 詳細な機能・オプションはREADMEを参照してください。

### Overview

- Adds a bullet symbol or sequential numbers to the head of each line in the selected text frames.
- On open, the existing leading markers are detected so you can keep editing from the current state while watching a live preview.
- See the README for the full feature and option list.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddBulletsAndNumbers";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-18";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBulletsAndNumbers.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddBulletsAndNumbers.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd738e3258989"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var BULLET_MARK = "・"; // 箇条書き記号の既定 / Default bullet symbol
    var START_NUMBER = 1;   // 開始番号 / Start number

    // 箇条書き記号の候補と、選択時に適用する比率(%)・本文タブストップ倍率 / Bullet symbol candidates, the scale (%) and the body tab-stop ratio applied when selected
    // bodyRatio 省略時は TAB_STOP_RATIO_1 を使用 / when bodyRatio is omitted, TAB_STOP_RATIO_1 is used
    var BULLET_SYMBOLS = [
        { mark: "•", scale: 100, bodyRatio: 0.8 },
        { mark: "-", scale: 100, bodyRatio: 0.8 },
        { mark: "●", scale: 50 },
        { mark: "○", scale: 50 },
        { mark: "◎", scale: 90 },
        // { mark: "★", scale: 90 },
        // { mark: "☆", scale: 90 },
        { mark: "■", scale: 90 },
        { mark: "□", scale: 90 },
        { mark: "◆", scale: 90 },
        { mark: "◇", scale: 90 }
    ];

    // 区切り文字の候補（空文字＝なし。コロンは全角「：」）/ Delimiter candidates (empty = none; colon is full-width "：")
    var DELIMITERS = ["", ".", "：", "|"];

    // tabストップ既定倍率（文字サイズ基準, 種類ごと）/ Default tab stop ratios (relative to font size, per type)
    var TAB_STOP_RATIO_1 = 1.2;          // 箇条書き・丸数字の1つ目／本文（1ストップ）/ bullet & circled: 1st stop / body (single stop)
    var NUMBERED_TAB_STOP_RATIO_1 = 1.5; // 数字の1つ目（番号列）/ numbers 1st stop (number column)
    var ALPHA_TAB_STOP_RATIO_1 = 1.0;    // ABC/abc の1つ目（マーカー列）/ ABC/abc 1st stop (marker column)
    var NUMBERED_TAB_STOP_RATIO_2 = 2.0; // 番号リストの2つ目（本文）/ numbered 2nd stop (text)

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margins */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / column gutter */
    var DENSE_SPACING  = 6;                  /* ラジオ・入力行を詰めるパネルの間隔 / spacing for dense panels */
    var FIELD_LABEL_WIDTH = 100;             /* 書式・位置ラベルの幅 / label width for format & position rows */
    var PARAGRAPH_LABEL_WIDTH = 110;         /* 段落設定ラベルの幅 / label width for paragraph rows */
    var REORDER_BUTTON_HEIGHT = 22;          /* 並べ替えボタンの高さ / reorder button height */
    var SWATCH_SIZE = 20;                    /* カラースウォッチの一辺 / color swatch size */
    var PREVIEW_LIST_SIZE = [250, 350];      /* 行一覧の大きさ [幅,高さ] / preview list size */
    var PREVIEW_LIST_NUMBER_COL = 40;        /* 行一覧の行番号カラム幅 / row-number column width */
    var PREVIEW_LIST_FONT_SIZE = 14;         /* 行一覧の文字サイズ / preview list font size */

    /**
     * ウィンドウの共通レイアウトを設定する
     * @param {Window} win - 対象のダイアログ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = ["fill", "top"];
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通レイアウトを設定する
     * @param {Panel} panel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 行グループの共通レイアウトを設定する（alignment と alignChildren は対で指定する）
     * @param {Group} targetGroup - 対象の行グループ
     * @param {string} [alignment] - 横方向の揃え（既定は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, alignment, spacing) {
        targetGroup.orientation = "row";
        targetGroup.alignment = [alignment || "left", "top"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 縦積みカラムの共通レイアウトを設定する
     * @param {Group} targetGroup - 対象のグループ
     * @param {number} [spacing] - 要素間隔（省略時は COLUMN_SPACING）
     * @returns {void}
     */
    function setupColumn(targetGroup, spacing) {
        targetGroup.orientation = "column";
        targetGroup.alignChildren = ["fill", "top"];
        targetGroup.alignment = ["fill", "top"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : COLUMN_SPACING;
    }

    /**
     * 共通レイアウト済みのパネルを追加する
     * @param {Group|Panel|Window} parent - 追加先
     * @param {string} titleText - パネルのタイトル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parent, titleText, spacing) {
        var panel = parent.add("panel", undefined, titleText);
        setupPanel(panel, spacing);
        return panel;
    }

    /**
     * ラベル＋入力欄＋単位の1行を追加する
     * @param {Group|Panel} parent - 追加先
     * @param {string} labelString - ラベル文字列
     * @param {string} initialText - 入力欄の初期値
     * @param {number} [labelWidth] - ラベルの幅（省略時は内容にフィット）
     * @param {string} [unitLabel] - 入力欄の後ろに置く単位（省略可）
     * @returns {{row: Group, label: StaticText, input: EditText}} 生成した行と部品
     */
    function addFieldRow(parent, labelString, initialText, labelWidth, unitLabel) {
        var row = parent.add("group");
        setupRow(row);
        var label = row.add("statictext", undefined, labelString);
        if (typeof labelWidth === "number") label.preferredSize.width = labelWidth;
        var input = row.add("edittext", undefined, initialText);
        input.characters = 4;
        if (unitLabel) row.add("statictext", undefined, unitLabel);
        return { row: row, label: label, input: input };
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語コードを取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "箇条書き・番号リスト", en: "Add Bullets and Numbers" },
            colorPicker: { ja: "カラーピッカー", en: "Color Picker" }
        },
        panel: {
            listType: { ja: "リストの種類", en: "List Type" },
            bulletStyle: { ja: "箇条書き記号", en: "Bullet Symbol" },
            numberStyle: { ja: "番号スタイル", en: "Number Style" },
            position: { ja: "位置（タブストップ）", en: "Position (Tab Stops)" },
            format: { ja: "マーカーの書式", en: "Marker Format" },
            sortElements: { ja: "行の並べ替え", en: "Reorder Lines" },
            paragraph: { ja: "段落設定", en: "Paragraph Settings" }
        },
        checkbox: {
            jpOnly: { ja: "和文フォントのみ", en: "Japanese fonts only" },
            zeroPad: { ja: "ゼロ埋め", en: "Zero padding" },
            hanging: { ja: "折り返し位置を揃える", en: "Align wrapped lines" },
            resetPerFrame: { ja: "フレームごとにリセット", en: "Restart each frame" }
        },
        delimiter: {
            label: { ja: "区切り文字", en: "Delimiter" },
            none: { ja: "なし", en: "None" }
        },
        header: {
            number: { ja: "#", en: "#" },
            text: { ja: "テキスト", en: "Text" }
        },
        reorder: {
            top: { ja: "先頭へ", en: "Top" },
            up: { ja: "上へ", en: "Up" },
            down: { ja: "下へ", en: "Down" },
            bottom: { ja: "末尾へ", en: "Bottom" },
            sortByName: { ja: "名前順", en: "By Name" },
            sortByValue: { ja: "数値順", en: "By Number" },
            sortByValueDesc: { ja: "数値順（逆）", en: "By Number ↓" }
        },
        numberStyle: {
            number: { ja: "数字", en: "Numbers" },
            circledWhite: { ja: "白丸数字", en: "Circled (white)" },
            circledBlack: { ja: "黒丸数字", en: "Circled (black)" }
        },
        fieldLabel: {
            tabStop1: { ja: "マーカー", en: "Marker" },
            tabStop2: { ja: "本文", en: "Body" },
            font: { ja: "フォント", en: "Font" },
            fontStyle: { ja: "スタイル", en: "Style" },
            scale: { ja: "比率", en: "Scale" },
            baselineShift: { ja: "ベースライン", en: "Baseline" },
            leading: { ja: "行送り", en: "Leading" },
            spaceAfter: { ja: "段落後のアキ", en: "Space After" },
            leftIndent: { ja: "インデント", en: "Indent" },
            markerColor: { ja: "記号／番号", en: "Marker/Number" },
            startNumber: { ja: "開始番号", en: "Start No." }
        },
        radio: {
            bullet: { ja: "箇条書き", en: "Bullets" },
            numbered: { ja: "番号リスト", en: "Numbered List" },
            none: { ja: "なし", en: "None" }
        },
        tabAlign: {
            left: { ja: "左", en: "Left" },
            center: { ja: "中央", en: "Center" },
            right: { ja: "右", en: "Right" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            showHidden: { ja: "制御文字を表示", en: "Show Hidden Characters" },
            reset: { ja: "リセット", en: "Reset" }
        },
        tooltip: {
            tabStop1: {
                ja: "箇条書き記号・丸数字では本文開始位置、数字／ABC／abcでは番号を置く位置です。",
                en: "For bullets and circled numbers, this is where the body starts; for numbers / ABC / abc, this is where the marker is placed."
            },
            tabStop2: {
                ja: "本文の開始位置です。数字／ABC／abc の番号リストで使用します。",
                en: "Where the body text starts. Used for number / ABC / abc lists."
            },
            tabAlign: {
                ja: "数字／ABC／abc のマーカー位置での揃え方です。箇条書きと丸数字では使用しません。",
                en: "Alignment at the marker position for number / ABC / abc lists. Not used for bullets or circled numbers."
            },
            scale: {
                ja: "記号・番号の拡大率（水平・垂直を同率で適用）",
                en: "Scale of the marker/number (applied equally to horizontal and vertical)"
            },
            baseline: {
                ja: "記号・番号の上下位置（ベースラインシフト）",
                en: "Vertical offset of the marker/number (baseline shift)"
            },
            zeroPad: {
                ja: "最大桁数に合わせて先頭を 0 で埋めます（数字のみ）",
                en: "Pad numbers with leading zeros to the largest width (numbers style only)"
            },
            font: {
                ja: "記号・番号に適用するフォントファミリーです。",
                en: "Font family applied to the marker/number."
            },
            fontStyle: {
                ja: "選択したフォントファミリー内のスタイルです。",
                en: "Style within the selected font family."
            },
            jpOnly: {
                ja: "フォント一覧を和文フォントだけに絞り込みます",
                en: "Limit the font list to Japanese fonts"
            },
            leading: {
                ja: "行送り（空欄のままなら変更しません）",
                en: "Line leading (left blank = leave unchanged)"
            },
            spaceAfter: {
                ja: "段落後のアキ（空欄のままなら変更しません）",
                en: "Space after the paragraph (left blank = leave unchanged)"
            },
            hanging: {
                ja: "2行以上に折り返したとき、2行目以降の開始位置を本文位置にそろえます。",
                en: "When a line wraps, align the second and later lines with the body text."
            },
            leftIndent: {
                ja: "折り返し行の開始位置です。通常は本文位置と同じ値にします。",
                en: "Start position for wrapped lines. Usually set this to the same value as the body position."
            },
            showHidden: {
                ja: "タブや改行などの制御文字の表示／非表示を切り替えます",
                en: "Toggle display of hidden characters such as tabs and line breaks"
            },
            color: {
                ja: "記号・番号のカラー。四角をクリックするとカラーピッカーを開きます。",
                en: "Color of the marker/number. Click the swatch to open the color picker."
            },
            delimiterColor: {
                ja: "区切り文字のカラー。四角をクリックするとカラーピッカーを開きます。",
                en: "Color of the delimiter. Click the swatch to open the color picker."
            },
            reset: {
                ja: "比率・ベースライン・カラー・インデント・タブストップをまとめてリセット",
                en: "Reset scale, baseline, color, indent, and tab stops"
            },
            startNumber: {
                ja: "連番の開始番号",
                en: "The number to start counting from"
            },
            resetPerFrame: {
                ja: "複数フレーム選択時、フレームごとに開始番号から振り直します",
                en: "With multiple frames selected, restart from the start number in each frame"
            }
        },
        alert: {
            noDoc: {
                ja: "ドキュメントが開かれていません。",
                en: "No document open."
            },
            noSelection: {
                ja: "テキストオブジェクトを選択してください。",
                en: "Please select text objects."
            },
            noTextFrame: {
                ja: "テキストフレームを選択してください。",
                en: "Please select text frames."
            },
            circledLimit: {
                ja: "丸数字は20までです。21番目以降は素の数字になります。",
                en: "Circled numbers support up to 20; the 21st and later become plain numbers."
            }
        }
    };

    /**
     * ドット区切りキーで現在の言語のラベルを取得する
     * @param {string} key - LABELS のキー（例: "panel.listType"）
     * @returns {string} ラベル文字列
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            labelNode = labelNode[keyParts[i]];
        }
        return labelNode[currentLanguage] || labelNode["en"];
    }

    /**
     * コロン付きのラベルを取得する（日本語は全角、英語は半角）
     * @param {string} key - LABELS のキー
     * @returns {string} コロンを付けたラベル文字列
     */
    function getLabelWithColon(key) {
        return getLabel(key) + (currentLanguage === 'ja' ? '：' : ':');
    }

    // =========================================
    // テキスト行の共通処理 / Line helpers
    // =========================================

    /**
     * 改行コードを統一したうえでテキストを行配列へ分割する
     * @param {string} text - 対象のテキスト
     * @returns {string[]} 行の配列
     */
    function splitIntoLines(text) {
        return String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトからテキストフレームを集める（グループ内も再帰）
     * @param {Array<PageItem>} items - 走査対象のオブジェクト配列
     * @param {Array<TextFrame>} out - 収集先の配列
     * @returns {Array<TextFrame>} 収集したテキストフレーム
     */
    function collectTextFrames(items, out) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            if (item.typename === "TextFrame") {
                out.push(item);
            } else if (item.typename === "GroupItem") {
                // グループ内のテキストフレームも対象 / include text frames inside groups
                collectTextFrames(item.pageItems, out);
            }
        }
        return out;
    }

    main();

    /**
     * 前提チェックののち、ダイアログ（プレビュー＆確定）を表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) { alert(getLabel('alert.noDoc')); return; }
        if (app.selection.length === 0) { alert(getLabel('alert.noSelection')); return; }

        // 選択（グループ内含む）からテキストフレームを収集 / Gather text frames from the selection (groups included)
        var targetSelection = collectTextFrames(app.selection, []);
        if (targetSelection.length === 0) {
            alert(getLabel('alert.noTextFrame'));
            return;
        }

        // 複数選択時は「上から下、同じ高さなら左から右」に並べ替えて連番が見た目順になるように
        // For multiple frames, order top→bottom (then left→right) so numbering matches visual order
        targetSelection.sort(function (a, b) {
            var aTop = 0, aLeft = 0, bTop = 0, bLeft = 0;
            try { aTop = a.top; aLeft = a.left; } catch (e1) { }
            try { bTop = b.top; bLeft = b.left; } catch (e2) { }
            if (bTop !== aTop) return bTop - aTop; // Illustrator は上ほど top が大きい / higher = larger top
            return aLeft - bLeft;
        });

        showDialog(targetSelection);
    }

    // =========================================
    // カラー / Color (ColorPicker 連携) / Color (ColorPicker integration)
    // =========================================

    /**
     * ドキュメントのカラースペースに応じた黒を生成する
     * @returns {CMYKColor|RGBColor} 生成した黒
     */
    function makeDefaultColor() {
        try {
            if (app.documents.length && app.activeDocument.documentColorSpace === DocumentColorSpace.CMYK) {
                var cmykBlack = new CMYKColor(); cmykBlack.cyan = 0; cmykBlack.magenta = 0; cmykBlack.yellow = 0; cmykBlack.black = 100; return cmykBlack;
            }
        } catch (e) { }
        var rgbBlack = new RGBColor(); rgbBlack.red = 0; rgbBlack.green = 0; rgbBlack.blue = 0; return rgbBlack;
    }

    /**
     * 先頭テキストフレームの塗り色を取得する（取得できなければ黒）
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {CMYKColor|RGBColor|GrayColor} 基準となる塗り色
     */
    function getBaseFillColor(selection) {
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") {
                try {
                    var col = selection[i].textRange.characterAttributes.fillColor;
                    if (col && col.typename && col.typename !== "NoColor") return col;
                } catch (e) { }
                break;
            }
        }
        return makeDefaultColor();
    }

    /**
     * Illustratorのカラーを ColorPicker 用の文字列へ変換する
     * @param {CMYKColor|RGBColor|GrayColor} aiColor - 変換元のカラー
     * @returns {string} 16進数文字列または "cmyk:c,m,y,k" 形式の文字列
     */
    function aiColorToPickerString(aiColor) {
        try {
            if (aiColor.typename === "RGBColor") {
                return ColorPicker.rgbToHex(aiColor.red, aiColor.green, aiColor.blue);
            } else if (aiColor.typename === "CMYKColor") {
                return "cmyk:" + Math.round(aiColor.cyan) + "," + Math.round(aiColor.magenta) + "," + Math.round(aiColor.yellow) + "," + Math.round(aiColor.black);
            } else if (aiColor.typename === "GrayColor") {
                return "cmyk:0,0,0," + Math.round(aiColor.gray);
            }
        } catch (e) { }
        return "000000";
    }

    /**
     * ColorPicker の文字列を Illustrator のカラーへ変換する
     * @param {string} pickerString - 16進数文字列または "cmyk:c,m,y,k" 形式の文字列
     * @returns {CMYKColor|RGBColor} 変換したカラー
     */
    function pickerStringToAiColor(pickerString) {
        if (ColorPicker.isCmykString(pickerString)) {
            var cmykParts = ColorPicker.parseCmykString(pickerString);
            var cmykColor = new CMYKColor();
            cmykColor.cyan = cmykParts.c; cmykColor.magenta = cmykParts.m; cmykColor.yellow = cmykParts.y; cmykColor.black = cmykParts.k;
            return cmykColor;
        }
        var rgbParts = ColorPicker.hexToRGB(pickerString);
        var rgbColor = new RGBColor();
        rgbColor.red = rgbParts.r; rgbColor.green = rgbParts.g; rgbColor.blue = rgbParts.b;
        return rgbColor;
    }

    /**
     * Illustratorのカラーをスウォッチ描画用のScriptUIブラシへ変換する
     * @param {object} graphics - 描画対象のScriptUI graphics オブジェクト
     * @param {CMYKColor|RGBColor|GrayColor} aiColor - 変換元のカラー
     * @returns {object} 生成したブラシ
     */
    function aiColorToScriptUIBrush(graphics, aiColor) {
        try {
            if (aiColor.typename === "RGBColor") {
                return graphics.newBrush(graphics.BrushType.SOLID_COLOR, [aiColor.red / 255, aiColor.green / 255, aiColor.blue / 255, 1]);
            } else if (aiColor.typename === "CMYKColor") {
                // CMYK→RGB の簡易近似（表示用）/ rough CMYK→RGB approximation for display
                var red = 1 - Math.min(1, aiColor.cyan / 100 + aiColor.black / 100);
                var green = 1 - Math.min(1, aiColor.magenta / 100 + aiColor.black / 100);
                var blue = 1 - Math.min(1, aiColor.yellow / 100 + aiColor.black / 100);
                return graphics.newBrush(graphics.BrushType.SOLID_COLOR, [red, green, blue, 1]);
            } else if (aiColor.typename === "GrayColor") {
                var grayLevel = 1 - (aiColor.gray / 100);
                return graphics.newBrush(graphics.BrushType.SOLID_COLOR, [grayLevel, grayLevel, grayLevel, 1]);
            }
        } catch (e) { }
        return graphics.newBrush(graphics.BrushType.SOLID_COLOR, [1, 1, 1, 1]);
    }

    /**
     * カラースウォッチ（クリックでカラーピッカーを開く四角）を生成する
     * @param {Group|Panel} parent - 追加先
     * @param {CMYKColor|RGBColor|GrayColor} aiColor - 初期カラー
     * @param {number} swatchSize - 一辺の長さ
     * @returns {Group} 生成したスウォッチ（_aiColor に現在のカラーを保持）
     */
    function createColorSwatch(parent, aiColor, swatchSize) {
        var swatch = parent.add("group");
        swatch.preferredSize = [swatchSize, swatchSize];
        swatch.minimumSize = [swatchSize, swatchSize];
        swatch._aiColor = aiColor;
        swatch.onDraw = function () {
            var graphics = this.graphics;
            var brush = aiColorToScriptUIBrush(graphics, this._aiColor);
            if (brush) { graphics.rectPath(0, 0, this.size[0], this.size[1]); graphics.fillPath(brush); }
            var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1);
            graphics.rectPath(0, 0, this.size[0], this.size[1]);
            graphics.strokePath(pen);
        };
        return swatch;
    }

    /**
     * カラースウォッチを再描画する（_aiColor の変更を画面へ反映）
     * @param {Group} swatch - createColorSwatch() が返したスウォッチ
     * @returns {void}
     */
    function redrawSwatch(swatch) {
        try { swatch.hide(); swatch.show(); } catch (e) { }
    }

    // =========================================
    // 現在状態の推定 / Detect current state
    // =========================================

    /**
     * 選択先頭のテキストフレームから最初の非空行を取得する
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {string|null} 最初の非空行（見つからなければ null）
     */
    function firstMarkedLine(selection) {
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename !== "TextFrame") continue;
            var lines = splitIntoLines(selection[i].contents);
            for (var j = 0; j < lines.length; j++) {
                if (/\S/.test(lines[j])) return lines[j];
            }
            return null;
        }
        return null;
    }

    /**
     * 区切り文字の DELIMITERS 内でのインデックスを取得する
     * @param {string} ch - 区切り文字
     * @returns {number} インデックス（未一致は 0＝なし）
     */
    function delimiterIndexOf(ch) {
        for (var i = 0; i < DELIMITERS.length; i++) {
            if (DELIMITERS[i] === ch) return i;
        }
        return 0;
    }

    /**
     * 現在の行頭マーカーからダイアログの初期状態を推定する（このスクリプトが付与した形式を想定）
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {{type: string, bulletIndex: number, numberStyle: string, delimiterIndex: number}} 推定した初期状態
     */
    function detectInitialState(selection) {
        var state = { type: "bullet", bulletIndex: 0, numberStyle: "number", delimiterIndex: 1 };
        var line = firstMarkedLine(selection);
        if (line == null) return state;

        // 丸数字（先頭タブなし／あり、区切りはタブまたはスペース）/ circled (with/without leading tab, separated by tab or space)
        if (/^\t?[①-⑳][\t 　]/.test(line)) { state.type = "numbered"; state.numberStyle = "circledWhite"; state.delimiterIndex = 0; return state; }
        if (/^\t?(?:[❶-❿]|[⓫-⓴])[\t 　]/.test(line)) { state.type = "numbered"; state.numberStyle = "circledBlack"; state.delimiterIndex = 0; return state; }

        // 数字 / ABC / abc / number / ABC / abc
        // 先頭タブあり: 区切りは任意・タブまたはスペース / leading tab: delimiter optional, separated by tab or space
        // 先頭タブなし（手打ち）: 区切り必須・スペース/タブ（本文の誤検出を抑える）/ no leading tab (hand-typed): delimiter required + space/tab (avoids false positives in body text)
        var numberedMatch = line.match(/^\t([A-Z]+|[a-z]+|\d+)([.：:|]?)[\t 　]/) ||
                            line.match(/^([A-Z]+|[a-z]+|\d+)([.：:|])[\t 　]/);
        if (numberedMatch) {
            var markerGlyph = numberedMatch[1];
            if (/^\d+$/.test(markerGlyph)) state.numberStyle = "number";
            else if (/^[A-Z]+$/.test(markerGlyph)) state.numberStyle = "upperAlpha";
            else state.numberStyle = "lowerAlpha";
            state.type = "numbered";
            state.delimiterIndex = delimiterIndexOf(numberedMatch[2]);
            return state;
        }

        // 箇条書き記号 + タブ / bullet glyph + tab
        if (line.charAt(1) === "\t") {
            var leadingGlyph = line.charAt(0);
            for (var bi = 0; bi < BULLET_SYMBOLS.length; bi++) {
                if (leadingGlyph === BULLET_SYMBOLS[bi].mark) {
                    state.type = "bullet";
                    state.bulletIndex = bi;
                    return state;
                }
            }
        }

        return state; // 未検出は既定（箇条書き・先頭記号）/ undetected → default (bullet, first glyph)
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを表示し、プレビューと確定適用を行う
     * @param {Array<TextFrame>} targetSelection - 対象のテキストフレーム配列（上→下、左→右の順）
     * @returns {void}
     */
    function showDialog(targetSelection) {
        // プレビュー前の状態を保存（キャンセル時に復元・並べ替えの作業領域）/ Snapshot the pre-preview state (restored on cancel; also the reorder work area)
        var baseline = captureBaseline(targetSelection);

        // プレビューの undo 状態（直前のプレビューを app.undo() で巻き戻すため）/ Preview undo state (so the previous preview can be rolled back via app.undo())
        var previewState = { isUndo: false };

        // 表示単位（text/units）と基準文字サイズ / Display unit (text/units) and base font size
        var unitInfo = getTextUnitInfo();
        var baseFontSize = getBaseFontSize(targetSelection);

        /**
         * ポイント値を表示単位へ換算する（小数第1位まで）
         * @param {number} valuePt - ポイント値
         * @returns {number} 表示単位での値
         */
        function toDisplayUnit(valuePt) {
            return Math.round(valuePt / unitInfo.factor * 10) / 10;
        }

        /**
         * 選択中の箇条書き記号の定義を取得する（初期化中でラジオ未生成なら先頭の記号）
         * @returns {{mark: string, scale: number, bodyRatio: number}} BULLET_SYMBOLS の要素
         */
        function selectedBulletSymbol() {
            if (typeof bulletRadios !== "undefined" && bulletRadios) {
                for (var bi = 0; bi < bulletRadios.length; bi++) {
                    if (bulletRadios[bi].value) return BULLET_SYMBOLS[bi];
                }
            }
            return BULLET_SYMBOLS[0];
        }

        /**
         * 選択中の箇条書き記号の本文タブストップ倍率を取得する（•/- は0.8など）
         * @returns {number} 本文タブストップの倍率（bodyRatio 未指定は TAB_STOP_RATIO_1）
         */
        function selectedBulletBodyRatio() {
            var symbol = selectedBulletSymbol();
            return (symbol.bodyRatio != null) ? symbol.bodyRatio : TAB_STOP_RATIO_1;
        }

        /**
         * 種類ごとの既定タブストップ（pt）を求める
         * 数字はマーカー×1.5・本文×2.0、ABC/abc はマーカー×1.0・本文×2.0、
         * 箇条書きの本文は記号ごとの bodyRatio（既定×1.2、•/-=0.8）、丸数字・なしは×1.2
         * @param {string} type - "numbered" / "alpha" / "bullet" / "circled" / "none"
         * @returns {{stop1Pt: number, stop2Pt: number}} マーカー位置と本文位置（pt）
         */
        function defaultStopsForType(type) {
            if (type === "numbered") return { stop1Pt: baseFontSize * NUMBERED_TAB_STOP_RATIO_1, stop2Pt: baseFontSize * NUMBERED_TAB_STOP_RATIO_2 };
            if (type === "alpha") return { stop1Pt: baseFontSize * ALPHA_TAB_STOP_RATIO_1, stop2Pt: baseFontSize * NUMBERED_TAB_STOP_RATIO_2 };
            if (type === "bullet") return { stop1Pt: baseFontSize * TAB_STOP_RATIO_1, stop2Pt: baseFontSize * selectedBulletBodyRatio() };
            // 丸数字・なし（1ストップ）/ circled, none (single stop)
            return { stop1Pt: baseFontSize * TAB_STOP_RATIO_1, stop2Pt: baseFontSize * TAB_STOP_RATIO_1 };
        }

        // 初期は箇条書き / Initial type is bullet
        var initialStops = defaultStopsForType("bullet");

        var dlg = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(dlg);

        /* 上段: 2カラム（左: 設定 / 右: テキストプレビュー）/ Top: two columns (left: settings / right: text preview) */
        var topRow = dlg.add("group");
        topRow.orientation = "row";
        topRow.alignChildren = ["fill", "fill"];
        topRow.spacing = COLUMN_SPACING;

        /* 左カラム / Left column */
        var topLeftCol = topRow.add("group");
        setupColumn(topLeftCol);

        /* リストの種類パネル（ラジオは横並び）/ List type panel (radios in a row) */
        var typePanel = addPanel(topLeftCol, getLabel('panel.listType'), DENSE_SPACING);

        // ラジオを横並びにし、パネル内で左右中央へ / Radios in a row, centered horizontally within the panel
        var typeRow = typePanel.add("group");
        setupRow(typeRow, "center", COLUMN_SPACING);

        var rbBullet = typeRow.add("radiobutton", undefined, getLabel('radio.bullet'));
        var rbNumbered = typeRow.add("radiobutton", undefined, getLabel('radio.numbered'));
        var rbNone = typeRow.add("radiobutton", undefined, getLabel('radio.none'));
        rbBullet.value = true;

        /* 記号／番号の種類／区切り文字の3カラム / Symbol / number style / delimiter (three columns) */
        var styleRow = topLeftCol.add("group");
        styleRow.orientation = "row";
        styleRow.alignChildren = ["fill", "top"];
        styleRow.spacing = COLUMN_SPACING;

        /* 箇条書き記号パネル（箇条書きのみ有効）/ Bullet symbol panel (bullet only) */
        var bulletStylePanel = addPanel(styleRow, getLabel('panel.bulletStyle'), DENSE_SPACING);

        // 記号候補をラジオで列挙 / List the symbol candidates as radios
        var bulletRadios = [];
        for (var symbolIndex = 0; symbolIndex < BULLET_SYMBOLS.length; symbolIndex++) {
            bulletRadios.push(bulletStylePanel.add("radiobutton", undefined, BULLET_SYMBOLS[symbolIndex].mark));
        }
        bulletRadios[0].value = true;

        /* 番号スタイル＋区切り文字を縦に並べるカラム / Column stacking number style + delimiter vertically */
        var numberStyleColumn = styleRow.add("group");
        setupColumn(numberStyleColumn, DENSE_SPACING);

        /* 番号スタイルパネル（番号リストのみ有効）/ Number style panel (numbered only) */
        var numberStylePanel = addPanel(numberStyleColumn, getLabel('panel.numberStyle'), DENSE_SPACING);
        numberStylePanel.alignChildren = ["left", "top"];

        // 「数字」とゼロ埋めを横並び（ゼロ埋めは数字スタイルのみ有効）/ "Numbers" + zero padding in a row (zero pad only for the numbers style)
        var numberStyleNumberRow = numberStylePanel.add("group");
        setupRow(numberStyleNumberRow, "left", DENSE_SPACING);
        var rbStyleNumber = numberStyleNumberRow.add("radiobutton", undefined, getLabel('numberStyle.number'));
        var chkZeroPad = numberStyleNumberRow.add("checkbox", undefined, getLabel('checkbox.zeroPad'));

        var rbStyleCircledWhite = numberStylePanel.add("radiobutton", undefined, getLabel('numberStyle.circledWhite'));
        var rbStyleCircledBlack = numberStylePanel.add("radiobutton", undefined, getLabel('numberStyle.circledBlack'));
        // アルファベット大文字 / 小文字（ラベルはそのまま表示）/ Uppercase / lowercase alphabet (self-explanatory labels)
        var rbStyleUpperAlpha = numberStylePanel.add("radiobutton", undefined, "ABC");
        var rbStyleLowerAlpha = numberStylePanel.add("radiobutton", undefined, "abc");
        rbStyleNumber.value = true;

        // 「数字」を別グループ（行）へ移したため、ラジオの排他は手動で管理 / "Numbers" is in its own row, so manage radio exclusivity manually
        var numberStyleRadios = [rbStyleNumber, rbStyleCircledWhite, rbStyleCircledBlack, rbStyleUpperAlpha, rbStyleLowerAlpha];

        /**
         * 番号スタイルのラジオを排他選択にする
         * @param {RadioButton} selected - 選択状態にするラジオボタン
         * @returns {void}
         */
        function selectNumberStyleExclusive(selected) {
            for (var nsi = 0; nsi < numberStyleRadios.length; nsi++) {
                numberStyleRadios[nsi].value = (numberStyleRadios[nsi] === selected);
            }
        }

        // 開始番号 / Start number
        var startNumberRow = numberStylePanel.add("group");
        setupRow(startNumberRow, "left", DENSE_SPACING);
        startNumberRow.add("statictext", undefined, getLabelWithColon('fieldLabel.startNumber'));
        var inputStartNumber = startNumberRow.add("edittext", undefined, String(START_NUMBER));
        inputStartNumber.characters = 3;

        // フレームごとにリセット（複数フレーム時、各フレームで開始番号から振り直す）/ Restart numbering at each frame (multi-frame)
        var chkResetPerFrame = numberStylePanel.add("checkbox", undefined, getLabel('checkbox.resetPerFrame'));

        /* 区切り文字パネル（番号リストのみ有効・縦並び）番号スタイルの下に配置 / Delimiter panel (numbered only); placed under the number style panel */
        var delimiterPanel = addPanel(numberStyleColumn, getLabel('delimiter.label'), DENSE_SPACING);
        delimiterPanel.alignChildren = ["left", "top"];

        // 区切り文字（なし/./：/|）は横並び / Delimiter radios in a row
        var delimiterRow = delimiterPanel.add("group");
        setupRow(delimiterRow, "left", DENSE_SPACING);
        var delimiterRadios = [];
        for (var delimIndex = 0; delimIndex < DELIMITERS.length; delimIndex++) {
            var delimLabel = (DELIMITERS[delimIndex] === "") ? getLabel('delimiter.none') : DELIMITERS[delimIndex];
            delimiterRadios.push(delimiterRow.add("radiobutton", undefined, delimLabel));
        }
        delimiterRadios[1].value = true; // 既定は「.」 / default "."

        /* 位置調整パネル / Position panel */
        var positionPanel = addPanel(topLeftCol, getLabel('panel.position'), COLUMN_SPACING);

        // 1行目: 1つ目／2つ目を左右2カラムに（それぞれ上揃え）/ Row 1: 1st / 2nd in two columns (each top-aligned)
        var tabStopColumns = positionPanel.add("group");
        setupRow(tabStopColumns, "left", COLUMN_SPACING * 2); // 1つ目／2つ目カラムのガターを広めに / wider gutter between the 1st / 2nd columns
        tabStopColumns.alignChildren = ["left", "top"];

        // 左: 1つ目のtabストップ（数字／箇条書きの位置）＋揃え種類 / Left column: 1st tab stop (number / bullet) + alignment type
        var tabStopLeftCol = tabStopColumns.add("group");
        setupColumn(tabStopLeftCol, DENSE_SPACING);
        tabStopLeftCol.alignChildren = ["left", "top"];
        tabStopLeftCol.alignment = ["left", "top"];

        var tabStop1UI = addFieldRow(tabStopLeftCol, getLabelWithColon('fieldLabel.tabStop1'), String(toDisplayUnit(initialStops.stop1Pt)), null, unitInfo.label);
        var tabStopRow1 = tabStop1UI.row;
        var tabStopLabel1 = tabStop1UI.label;
        var inputTabStop1 = tabStop1UI.input;

        // 1つ目のタブストップの揃え種類（左／中央／右）/ Alignment type of the 1st tab stop (left/center/right)
        var tabAlignRow = tabStopLeftCol.add("group");
        setupRow(tabAlignRow, "left", DENSE_SPACING);
        tabAlignRow.margins = [0, 5, 0, 0]; // 1つ目入力との間に上マージン / top margin from the 1st-stop input
        var rbAlignLeft = tabAlignRow.add("radiobutton", undefined, getLabel('tabAlign.left'));
        var rbAlignCenter = tabAlignRow.add("radiobutton", undefined, getLabel('tabAlign.center'));
        var rbAlignRight = tabAlignRow.add("radiobutton", undefined, getLabel('tabAlign.right'));
        rbAlignRight.value = true; // 既定は右揃え（番号リストの番号位置）/ default right (number column for numbered)

        // 右: 2つ目のtabストップ（本文の位置, 番号リストのみ）/ Right: 2nd tab stop (text column, numbered only)
        var tabStop2UI = addFieldRow(tabStopColumns, getLabelWithColon('fieldLabel.tabStop2'), String(toDisplayUnit(initialStops.stop2Pt)), null, unitInfo.label);
        var tabStopRow2 = tabStop2UI.row;
        var tabStopLabel2 = tabStop2UI.label;
        var inputTabStop2 = tabStop2UI.input;

        /* 上段右カラム: 行の並べ替え / Top-right column: reorder lines */
        var topRightCol = topRow.add("group");
        setupColumn(topRightCol, 0);
        topRightCol.margins = 0;

        /* 行の並べ替えパネル（一覧＋並べ替え操作をまとめる）/ Reorder-lines wrapper panel (list + reorder controls) */
        // 並べ替えボタンを listbox の直下に置くため間隔は0 / spacing 0 so the reorder buttons sit flush under the listbox
        var sortPanel = addPanel(topRightCol, getLabel('panel.sortElements'), 0);

        // 左=行番号 / 右=本文（マーカーは含めない）/ Left = row number, Right = body text (markers excluded)
        // 列幅の合計をリストボックス幅に合わせて2列目を端まで使う / column widths sum to the listbox width so column 2 fills to the edge
        var previewList = sortPanel.add("listbox", undefined, [], {
            numberOfColumns: 2,
            showHeaders: true,
            multiselect: true, // 複数行を選択して一括移動できるように / allow selecting multiple rows to move them together
            columnTitles: [getLabel('header.number'), getLabel('header.text')],
            columnWidths: [PREVIEW_LIST_NUMBER_COL, PREVIEW_LIST_SIZE[0] - PREVIEW_LIST_NUMBER_COL]
        });
        previewList.preferredSize = PREVIEW_LIST_SIZE;
        // リストの文字サイズを大きめに / enlarge the list font
        var previewListFont = previewList.graphics.font;
        previewList.graphics.font = ScriptUI.newFont(previewListFont.name, previewListFont.style, PREVIEW_LIST_FONT_SIZE);

        /**
         * 並べ替えボタンを1つ追加する
         * @param {Group} parent - 追加先の行グループ
         * @param {string} labelString - ボタンのラベル
         * @param {number} width - ボタンの幅
         * @returns {Button} 追加したボタン
         */
        function addReorderButton(parent, labelString, width) {
            var button = parent.add("button", undefined, labelString);
            button.preferredSize = [width, REORDER_BUTTON_HEIGHT];
            return button;
        }

        /* 並べ替え行（先頭へ/上へ/下へ/末尾へ）/ Reorder row (top/up/down/bottom) */
        var reorderRow = sortPanel.add("group");
        setupRow(reorderRow, "fill", 4);
        reorderRow.margins = [0, 10, 0, 0]; // listbox との間に上マージン / top margin from the listbox

        var btnMoveTop = addReorderButton(reorderRow, getLabel('reorder.top'), 52);
        var btnMoveUp = addReorderButton(reorderRow, getLabel('reorder.up'), 46);
        var btnMoveDown = addReorderButton(reorderRow, getLabel('reorder.down'), 46);
        var btnMoveBottom = addReorderButton(reorderRow, getLabel('reorder.bottom'), 52);

        // 名前順・数値順は次の行に / "By name" / "By number" go on the next row
        var sortByNameRow = sortPanel.add("group");
        setupRow(sortByNameRow, "fill", 4);
        sortByNameRow.margins = [0, 5, 0, 0]; // 上にマージン / top margin

        var btnSortByName = addReorderButton(sortByNameRow, getLabel('reorder.sortByName'), 70);
        var btnSortByValue = addReorderButton(sortByNameRow, getLabel('reorder.sortByValue'), 70);
        var btnSortByValueDesc = addReorderButton(sortByNameRow, getLabel('reorder.sortByValueDesc'), 100);

        /* ===== 2行目: 左=記号や番号の書式 / 右=段落の書式 / Row 2: left = symbol/number format, right = paragraph format ===== */
        var bottomRow = dlg.add("group");
        bottomRow.orientation = "row";
        bottomRow.alignChildren = ["fill", "top"];
        bottomRow.spacing = COLUMN_SPACING;

        /* 下段左カラム / Bottom-left column */
        var bottomLeft = bottomRow.add("group");
        setupColumn(bottomLeft);

        /* マーカーの書式パネル / Marker format panel */
        var formatPanel = addPanel(bottomLeft, getLabel('panel.format'), DENSE_SPACING);

        // フォントファミリー（ポップアップ）/ Font family (popup)
        var baseFontName = getBaseFontName(targetSelection);
        var fontRow = formatPanel.add("group");
        setupRow(fontRow);
        var fontLabel = fontRow.add("statictext", undefined, getLabelWithColon('fieldLabel.font'));
        fontLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        var fontFamilyDropdown = fontRow.add("dropdownlist", undefined, []);
        fontFamilyDropdown.preferredSize.width = 200;

        // フォントスタイル（ポップアップ）/ Font style (popup)
        var fontStyleRow = formatPanel.add("group");
        setupRow(fontStyleRow);
        var fontStyleLabel = fontStyleRow.add("statictext", undefined, getLabelWithColon('fieldLabel.fontStyle'));
        fontStyleLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        var fontStyleDropdown = fontStyleRow.add("dropdownlist", undefined, []);
        fontStyleDropdown.preferredSize.width = 200;

        // 和文フォントのみ / Japanese fonts only
        var jpOnlyRow = formatPanel.add("group");
        setupRow(jpOnlyRow);
        jpOnlyRow.add("statictext", undefined, "").preferredSize.width = FIELD_LABEL_WIDTH; // ラベル列に合わせるスペーサー / spacer to align with the label column
        var chkJPOnly = jpOnlyRow.add("checkbox", undefined, getLabel('checkbox.jpOnly'));

        populateFontFamilyDropdown(fontFamilyDropdown, baseFontName, chkJPOnly.value);
        populateFontStyleDropdown(fontStyleDropdown, fontFamilyDropdown.selection, baseFontName);

        // 2カラム: 左=比率・ベースライン / 右=カラー / Two columns: left = scale & baseline, right = color
        var formatColumns = formatPanel.add("group");
        setupRow(formatColumns, "left", COLUMN_SPACING);
        formatColumns.alignChildren = ["left", "top"];

        // 左カラム / Left column: scale & baseline
        var formatLeft = formatColumns.add("group");
        setupColumn(formatLeft, DENSE_SPACING);
        formatLeft.alignChildren = ["left", "top"];
        formatLeft.alignment = ["left", "top"];

        // 比率（水平・垂直を同率で適用, %）。初期は箇条書きの記号に合わせた比率 / Scale (applied equally to H and V, %), initialized for the default bullet glyph
        var scaleUI = addFieldRow(formatLeft, getLabelWithColon('fieldLabel.scale'), "120", FIELD_LABEL_WIDTH, "%");
        var scaleLabel = scaleUI.label;
        var inputScale = scaleUI.input;

        // ベースラインシフト / Baseline shift
        var baselineUI = addFieldRow(formatLeft, getLabelWithColon('fieldLabel.baselineShift'), "0", FIELD_LABEL_WIDTH, unitInfo.label);
        var baselineLabel = baselineUI.label;
        var inputBaseline = baselineUI.input;

        // 右カラム / Right column: color
        var formatRight = formatColumns.add("group");
        setupColumn(formatRight, DENSE_SPACING);
        formatRight.alignChildren = ["left", "top"];
        formatRight.alignment = ["left", "top"];

        // カラー（■をクリックでカラーピッカー）。記号／番号と区切り文字を別々に指定 / Color (click a swatch to open the picker). Marker/number and delimiter are set separately
        var baseFillColor = getBaseFillColor(targetSelection);

        // カラー行のラベル幅（比率・ベースラインの FIELD_LABEL_WIDTH とは独立）/ Color-row label width (independent of FIELD_LABEL_WIDTH used by scale/baseline)
        var colorLabelWidth = (currentLanguage === 'ja') ? 80 : 110;

        /**
         * ラベル＋カラースウォッチの1行を追加する（スウォッチのクリックでピッカー→プレビュー）
         * @param {Group|Panel} parent - 追加先
         * @param {string} labelString - ラベル文字列
         * @param {CMYKColor|RGBColor|GrayColor} initColor - 初期カラー
         * @returns {{row: Group, label: StaticText, swatch: Group}} 生成した行と部品
         */
        function addColorRow(parent, labelString, initColor) {
            var row = parent.add("group");
            setupRow(row, "left", DENSE_SPACING);
            var label = row.add("statictext", undefined, labelString);
            label.preferredSize.width = colorLabelWidth;
            var swatch = createColorSwatch(row, initColor, SWATCH_SIZE);
            swatch.addEventListener("click", function () {
                var result = ColorPicker.show({
                    value: aiColorToPickerString(swatch._aiColor),
                    title: getLabel('dialog.colorPicker'),
                    lang: currentLanguage
                });
                if (result !== null) {
                    swatch._aiColor = pickerStringToAiColor(result);
                    redrawSwatch(swatch);
                    updatePreview();
                }
            });
            return { row: row, label: label, swatch: swatch };
        }

        var markerColorUI = addColorRow(formatRight, getLabelWithColon('fieldLabel.markerColor'), baseFillColor);
        var delimiterColorUI = addColorRow(formatRight, getLabelWithColon('delimiter.label'), baseFillColor);
        var colorSwatch = markerColorUI.swatch;             // 記号／番号のカラー / marker-number color
        var delimiterColorSwatch = delimiterColorUI.swatch; // 区切り文字のカラー / delimiter color

        /* 下段右カラム / Bottom-right column */
        var bottomRight = bottomRow.add("group");
        setupColumn(bottomRight);

        /* 段落設定パネル / Paragraph settings panel */
        var paragraphPanel = addPanel(bottomRight, getLabel('panel.paragraph'), DENSE_SPACING);

        // 行送り（初期値は選択テキストの現在値）/ Leading (initialized to the selection's current value)
        var baseLeadingPt = getBaseLeading(targetSelection);
        var leadingDefaultText = (baseLeadingPt != null) ? String(toDisplayUnit(baseLeadingPt)) : "";
        var leadingUI = addFieldRow(paragraphPanel, getLabelWithColon('fieldLabel.leading'), leadingDefaultText, PARAGRAPH_LABEL_WIDTH, unitInfo.label);
        var leadingLabel = leadingUI.label;
        var inputLeading = leadingUI.input;

        // 段落後のアキ / Space after
        var spaceAfterUI = addFieldRow(paragraphPanel, getLabelWithColon('fieldLabel.spaceAfter'), "0", PARAGRAPH_LABEL_WIDTH, unitInfo.label);
        var spaceAfterLabel = spaceAfterUI.label;
        var inputSpaceAfter = spaceAfterUI.input;

        // ハンギング対応（チェック時のみ左/1行目インデントを適用）/ Hanging indent (apply left/first-line indent only when checked)
        // チェックボックスは他のラベルと同じ左位置に / checkbox aligned at the label column like the other rows
        var hangingRow = paragraphPanel.add("group");
        setupRow(hangingRow, "left", DENSE_SPACING);
        hangingRow.margins = [0, 12, 0, 0]; // 上に余白 / top margin
        var chkHanging = hangingRow.add("checkbox", undefined, getLabel('checkbox.hanging'));

        // 左インデント（1行目インデントはUIを持たず、内部で左インデントの正負を反転して適用）
        // Left indent (first-line indent has no UI; it's the negated left indent, computed internally)
        var leftIndentUI = addFieldRow(paragraphPanel, getLabelWithColon('fieldLabel.leftIndent'), "0", PARAGRAPH_LABEL_WIDTH, unitInfo.label);
        var leftIndentLabel = leftIndentUI.label;
        var inputLeftIndent = leftIndentUI.input;

        /* ボタンエリア（左: 制御文字表示 / 中央: spacer / 右: Cancel → OK）/ Button area (left: show hidden / center: spacer / right: Cancel → OK) */
        var btnGroup = dlg.add("group");
        btnGroup.orientation = "row";
        btnGroup.alignment = "fill";

        // 左カラム / Left column
        var btnLeft = btnGroup.add("group");
        var btnShowHidden = btnLeft.add("button", undefined, getLabel('button.showHidden'));
        var btnReset = btnLeft.add("button", undefined, getLabel('button.reset'));

        // 中央カラム（伸縮スペーサー）/ Center column (flexible spacer)
        var btnSpacer = btnGroup.add("group");
        btnSpacer.alignment = ["fill", "fill"];
        btnSpacer.minimumSize.width = 0;

        // 右カラム / Right column
        var btnRight = btnGroup.add("group");
        btnRight.alignChildren = ["right", "center"];
        var btnCancel = btnRight.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnOK = btnRight.add("button", undefined, "OK", { name: "ok" });

        /* 制御文字（隠し文字）の表示切り替え / Toggle hidden characters */
        btnShowHidden.onClick = function () {
            app.executeMenuCommand('showHiddenChar');
        };

        /* まとめてリセット / Reset everything listed */
        btnReset.onClick = function () { resetAllValues(); };

        /**
         * 現在選択されているリストの種類を取得する
         * @returns {string} "bullet" / "numbered" / "none"
         */
        function currentType() {
            if (rbNumbered.value) return "numbered";
            if (rbNone.value) return "none";
            return "bullet";
        }

        /**
         * 現在選択されている箇条書き記号を取得する
         * @returns {string} 箇条書き記号
         */
        function currentBulletMark() {
            return selectedBulletSymbol().mark;
        }

        /**
         * 現在選択されている区切り文字を取得する
         * @returns {string} 区切り文字（「なし」は空文字）
         */
        function currentDelimiter() {
            for (var di = 0; di < delimiterRadios.length; di++) {
                if (delimiterRadios[di].value) return DELIMITERS[di];
            }
            return "";
        }

        /**
         * 現在選択されている番号スタイルを取得する
         * @returns {string} "number" / "circledWhite" / "circledBlack" / "upperAlpha" / "lowerAlpha"
         */
        function currentStyle() {
            if (rbStyleCircledWhite.value) return "circledWhite";
            if (rbStyleCircledBlack.value) return "circledBlack";
            if (rbStyleUpperAlpha.value) return "upperAlpha";
            if (rbStyleLowerAlpha.value) return "lowerAlpha";
            return "number";
        }

        /**
         * 開始番号を取得する（不正値・負値は START_NUMBER へフォールバック）
         * @returns {number} 開始番号
         */
        function currentStartNumber() {
            var value = parseInt(inputStartNumber.text, 10);
            return (isNaN(value) || value < 0) ? START_NUMBER : value;
        }

        /**
         * 付与される最大の番号を求める（丸数字の警告・ゼロ埋めの桁そろえ用）
         * @returns {number} 最大の番号
         */
        function largestNumberValue() {
            var count = chkResetPerFrame.value
                ? maxNumberedItemsPerFrame(targetSelection)   // フレームごとにリセット時は最大フレームの件数 / per-frame: the biggest frame's count
                : countNumberedItems(targetSelection);        // 継続時は総数 / continuous: total
            return currentStartNumber() + count - 1;
        }

        /**
         * パーセント入力欄の値を取得する（不正値・0以下は100へフォールバック）
         * @param {EditText} input - 対象の入力欄
         * @returns {number} パーセント値
         */
        function readPercent(input) {
            var value = parseFloat(input.text);
            return (isNaN(value) || value <= 0) ? 100 : value;
        }

        /**
         * マーカーの書式（フォント・比率・ベースラインシフト・カラー）を取得する
         * @returns {{fontName: string, horizontalScale: number, verticalScale: number, baselineShiftPt: number, fillColor: object, delimiterFillColor: object}} 書式
         */
        function currentFormat() {
            var fontName = getSelectedFontName(fontFamilyDropdown, fontStyleDropdown);

            var scale = readPercent(inputScale); // 水平・垂直に同率で適用 / applied equally to H and V

            var baseline = parseFloat(inputBaseline.text);
            if (isNaN(baseline)) baseline = 0;

            return {
                fontName: fontName,
                horizontalScale: scale,
                verticalScale: scale,
                baselineShiftPt: baseline * unitInfo.factor,
                fillColor: colorSwatch._aiColor,          // 記号／番号の塗り色 / marker-number fill color
                delimiterFillColor: delimiterColorSwatch._aiColor // 区切り文字の塗り色 / delimiter fill color
            };
        }

        /**
         * 段落の書式（行送り・段落後のアキ・インデント）を取得する。空欄は null（＝変更しない）
         * @returns {{leadingPt: (number|null), spaceAfterPt: (number|null), leftIndentPt: (number|null), firstLineIndentPt: (number|null)}} 段落の書式
         */
        function currentParagraphFormat() {
            /**
             * 入力欄をポイント値として読む（空欄・不正値・負値は null）
             * @param {EditText} input - 対象の入力欄
             * @returns {number|null} ポイント値
             */
            function readOptionalPt(input) {
                if (!input.text || !/\S/.test(input.text)) return null; // 空欄は変更なし / blank = no change
                var value = parseFloat(input.text);
                if (isNaN(value) || value < 0) return null;
                return value * unitInfo.factor; // 表示単位→pt / display unit → pt
            }
            // ハンギングON時のみインデント適用。1行目インデントは左インデントの正負反転（UIなし・内部計算）
            // Indents apply only when hanging is on; first-line indent = negated left indent (no UI, computed internally)
            var leftIndentPt = chkHanging.value ? readOptionalPt(inputLeftIndent) : null;
            var firstLineIndentPt = (leftIndentPt != null) ? -leftIndentPt : null;
            return {
                leadingPt: readOptionalPt(inputLeading),
                spaceAfterPt: readOptionalPt(inputSpaceAfter),
                leftIndentPt: leftIndentPt,
                firstLineIndentPt: firstLineIndentPt
            };
        }

        /**
         * 「折り返し位置を揃える」のON/OFFに応じてインデント欄をディムする
         * @returns {void}
         */
        function updateHangingEnabled() {
            var hangingOn = chkHanging.value;
            leftIndentLabel.enabled = hangingOn;
            inputLeftIndent.enabled = hangingOn;
        }

        /**
         * 「折り返し位置を揃える」がONのとき、本文位置に左インデントの値を反映する
         * @returns {void}
         */
        function syncHangingTabStop() {
            if (!chkHanging.value || rbNone.value) return;
            inputTabStop2.text = inputLeftIndent.text; // 本文位置＝左インデント（全種類共通）/ body position = left indent (all types)
        }

        /**
         * タブストップ入力欄の値を表示単位で取得する
         * @param {EditText} input - 対象の入力欄
         * @param {number} fallbackPt - 不正値のときに使う既定値（pt）
         * @returns {number} 表示単位での値
         */
        function readTabStop(input, fallbackPt) {
            var value = parseFloat(input.text);
            return (isNaN(value) || value < 0) ? toDisplayUnit(fallbackPt) : value;
        }

        /**
         * タブストップ既定値を決めるための種類を取得する（丸数字は1ストップ、ABC/abc はマーカー×1.0）
         * @returns {string} "bullet" / "none" / "circled" / "alpha" / "numbered"
         */
        function currentStopType() {
            if (rbBullet.value) return "bullet";
            if (rbNone.value) return "none";
            if (rbStyleCircledWhite.value || rbStyleCircledBlack.value) return "circled";
            if (rbStyleUpperAlpha.value || rbStyleLowerAlpha.value) return "alpha";
            return "numbered";
        }

        /**
         * マーカー位置のタブストップをポイントで取得する
         * @returns {number} マーカー位置（pt）
         */
        function currentTabStop1Pt() {
            return readTabStop(inputTabStop1, defaultStopsForType(currentStopType()).stop1Pt) * unitInfo.factor;
        }
        function currentTabStop2Pt() {
            return readTabStop(inputTabStop2, defaultStopsForType(currentStopType()).stop2Pt) * unitInfo.factor;
        }

        /* 1つ目のタブストップの揃え種類を取得 / Get the chosen alignment for the 1st tab stop */
        function currentTab1Alignment() {
            if (rbAlignLeft.value) return TabStopAlignment.Left;
            if (rbAlignCenter.value) return TabStopAlignment.Center;
            return TabStopAlignment.Right;
        }

        /* 種類に応じて既定のtabストップ値を入力欄へ反映 / Reset the tab stop fields to the type's defaults */
        function applyTypeDefaults() {
            if (rbNone.value) return; // 「なし」では位置調整は未使用 / Position is unused for "none"
            var stops = defaultStopsForType(currentStopType());
            inputTabStop1.text = String(toDisplayUnit(stops.stop1Pt));
            inputTabStop2.text = String(toDisplayUnit(stops.stop2Pt));
        }

        /* 現在状態の推定結果をUIへ反映（「現状の続き」として編集できるように）/ Apply the detected current state to the UI (so editing continues from it) */
        function applyDetectedState(detectedState) {
            // 種類 / type
            rbBullet.value = (detectedState.type === "bullet");
            rbNumbered.value = (detectedState.type === "numbered");
            rbNone.value = (detectedState.type === "none");
            // 箇条書き記号 / bullet glyph
            for (var i = 0; i < bulletRadios.length; i++) bulletRadios[i].value = (i === detectedState.bulletIndex);
            // 番号スタイル / number style
            rbStyleNumber.value = (detectedState.numberStyle === "number");
            rbStyleCircledWhite.value = (detectedState.numberStyle === "circledWhite");
            rbStyleCircledBlack.value = (detectedState.numberStyle === "circledBlack");
            rbStyleUpperAlpha.value = (detectedState.numberStyle === "upperAlpha");
            rbStyleLowerAlpha.value = (detectedState.numberStyle === "lowerAlpha");
            // 区切り文字 / delimiter
            for (var d = 0; d < delimiterRadios.length; d++) delimiterRadios[d].value = (d === detectedState.delimiterIndex);
            // 比率 / scale
            if (detectedState.type === "bullet") inputScale.text = String(BULLET_SYMBOLS[detectedState.bulletIndex].scale);
            else if (detectedState.type === "numbered") inputScale.text = "100";
            // 1つ目タブストップの揃え / 1st-stop alignment
            if (detectedState.type === "numbered" && (detectedState.numberStyle === "upperAlpha" || detectedState.numberStyle === "lowerAlpha")) rbAlignCenter.value = true;
            else if (detectedState.type === "numbered" && detectedState.numberStyle === "number") rbAlignRight.value = true;
            else rbAlignLeft.value = true; // 箇条書き・丸数字 / bullet & circled
            // 種類に応じたタブストップ既定値 / type-based default tab stops
            applyTypeDefaults();
        }

        /* 一覧を更新（左=行番号, 右=マーカーを除いた本文）/ Refresh the list (left = row number, right = body without markers) */
        function refreshList() {
            // removeAll で選択も消えるため、先に選択行を退避して再構築後に復元
            // removeAll() drops the selection too, so capture it first and restore after rebuilding
            // （並べ替え時は applyReorder が後から正しい選択で上書きするので競合しない / on reorder, applyReorder overrides this afterward)
            var prevSelected = selectedRowIndices();

            previewList.removeAll();
            var bodyEntries = getFlatBodyEntries(baseline);
            for (var k = 0; k < bodyEntries.length; k++) {
                var row = previewList.add("item", String(k + 1)); // 左カラム: 連番 / Left column: sequential number
                row.subItems[0].text = bodyEntries[k].text;        // 右カラム: 本文 / Right column: body text
            }

            // 退避した選択を復元（行数が減っている場合は範囲内のみ）/ Restore selection (only indices still in range)
            if (prevSelected.length) {
                var restore = [];
                for (var s = 0; s < prevSelected.length; s++) {
                    if (prevSelected[s] < previewList.items.length) restore.push(prevSelected[s]);
                }
                if (restore.length) previewList.selection = restore;
            }

            updateMoveButtonEnabled();
        }

        /* 選択中の行インデックスを昇順配列で取得（複数選択対応, 未選択は空配列）/ Selected row indices as a sorted array (multi-select; empty if none) */
        function selectedRowIndices() {
            var sel = previewList.selection;
            if (!sel) return [];
            // multiselect では配列、単一選択でも item が返るため両対応 / multiselect returns an array; a single item is also handled
            var items = (sel instanceof Array) ? sel : [sel];
            var indices = [];
            for (var i = 0; i < items.length; i++) {
                if (items[i]) indices.push(items[i].index);
            }
            indices.sort(function (a, b) { return a - b; });
            return indices;
        }

        function updateMoveButtonEnabled() {
            var hasSelection = selectedRowIndices().length > 0;
            btnMoveTop.enabled = hasSelection;
            btnMoveUp.enabled = hasSelection;
            btnMoveDown.enabled = hasSelection;
            btnMoveBottom.enabled = hasSelection;
        }

        /* 並べ替え後に baseline へ反映し、プレビュー・選択を更新（複数選択を復元）/ Write the reorder back, refresh preview, restore the (multi) selection */
        function applyReorder(flatEntries, newIndices) {
            writeFlatBodyEntries(baseline, flatEntries);
            updatePreview(); // 一覧の再構築と再適用 / rebuilds the list and reapplies
            var valid = [];
            for (var i = 0; i < newIndices.length; i++) {
                if (newIndices[i] >= 0 && newIndices[i] < previewList.items.length) valid.push(newIndices[i]);
            }
            if (valid.length) previewList.selection = valid; // 配列で複数選択を復元 / restore multi-selection via an array
            updateMoveButtonEnabled();
        }

        /* 選択行（複数可）を delta（±1）だけ移動。相対順を保ち、端や他の選択行で詰まる分は動かさない / Move selected rows by delta (±1), preserving order; blocked by the edge or another selected row */
        function moveSelectedRows(delta) {
            var indices = selectedRowIndices();
            if (!indices.length) return;
            var flatEntries = getFlatBodyEntries(baseline);
            var n = flatEntries.length;
            var selected = {};
            for (var s = 0; s < indices.length; s++) selected[indices[s]] = true;

            if (delta < 0) {
                // 上へ: 上から順に、直上が未選択なら入れ替え / up: top-down, swap with the slot above when it's free
                for (var up = 1; up < n; up++) {
                    if (selected[up] && !selected[up - 1]) {
                        var t1 = flatEntries[up]; flatEntries[up] = flatEntries[up - 1]; flatEntries[up - 1] = t1;
                        delete selected[up]; selected[up - 1] = true;
                    }
                }
            } else if (delta > 0) {
                // 下へ: 下から順に、直下が未選択なら入れ替え / down: bottom-up, swap with the slot below when it's free
                for (var dn = n - 2; dn >= 0; dn--) {
                    if (selected[dn] && !selected[dn + 1]) {
                        var t2 = flatEntries[dn]; flatEntries[dn] = flatEntries[dn + 1]; flatEntries[dn + 1] = t2;
                        delete selected[dn]; selected[dn + 1] = true;
                    }
                }
            }

            var newIndices = [];
            for (var k = 0; k < n; k++) { if (selected[k]) newIndices.push(k); }
            applyReorder(flatEntries, newIndices);
        }

        /* 選択行（複数可）を先頭/末尾へ移動（選択順＝元の並び順を維持）/ Move selected rows to top/bottom (keeping their original relative order) */
        function moveSelectedRowsTo(position) {
            var indices = selectedRowIndices();
            if (!indices.length) return;
            var flatEntries = getFlatBodyEntries(baseline);
            // 後ろから抜き出してインデックスのズレを防ぎ、元の順序で moved に並べる / splice from the end to avoid index drift; keep original order in `moved`
            var moved = [];
            for (var i = indices.length - 1; i >= 0; i--) {
                moved.unshift(flatEntries.splice(indices[i], 1)[0]);
            }
            var newIndices = [];
            if (position === "top") {
                for (var j = 0; j < moved.length; j++) { flatEntries.splice(j, 0, moved[j]); newIndices.push(j); }
            } else {
                var base = flatEntries.length;
                for (var j2 = 0; j2 < moved.length; j2++) { flatEntries.push(moved[j2]); newIndices.push(base + j2); }
            }
            applyReorder(flatEntries, newIndices);
        }

        /* 全行を名前順（昇順）に並べ替え / Sort all rows by name (ascending) */
        function sortRowsByName() {
            var flatEntries = getFlatBodyEntries(baseline);
            flatEntries.sort(function (a, b) { return (a.text < b.text) ? -1 : (a.text > b.text ? 1 : 0); });
            applyReorder(flatEntries, []);
        }

        /* 本文中の最初の数値を取り出す（小数・マイナス可、なければ null）/ Extract the first number from the text (decimals/negatives ok; null if none) */
        function extractFirstNumber(text) {
            var match = String(text).match(/-?\d+(?:\.\d+)?/);
            return match ? parseFloat(match[0]) : null;
        }

        /* 全行を「項目内の数値」で並べ替え（数値なしは末尾）/ Sort all rows by the number found in each item (no-number rows go last) */
        function sortRowsByValue(descending) {
            var flatEntries = getFlatBodyEntries(baseline);
            flatEntries.sort(function (a, b) {
                var av = extractFirstNumber(a.text), bv = extractFirstNumber(b.text);
                if (av === null && bv === null) return 0;
                if (av === null) return 1;   // 数値なしは末尾へ / push no-number rows to the end
                if (bv === null) return -1;
                return descending ? (bv - av) : (av - bv);
            });
            applyReorder(flatEntries, []);
        }

        /* パネルの有効/無効を切り替え（種類に応じてディム）/ Enable/disable panels based on the list type */
        function updateEnabledStates() {
            // マーカー位置（番号列）は数字/ABC/abc のみ。本文位置は「なし」以外の全種類で使う。
            // Marker position (number column) only for number/ABC/abc; body position is used by every type except "none".
            var twoStopNumbered = rbNumbered.value && !(rbStyleCircledWhite.value || rbStyleCircledBlack.value);
            positionPanel.enabled = !rbNone.value;
            tabAlignRow.enabled = twoStopNumbered;        // 揃え種類はマーカー位置に対応（箇条書き・丸数字は左固定）/ alignment applies to the marker column
            bulletStylePanel.enabled = rbBullet.value;    // 記号の種類は箇条書きのみ有効 / Bullet symbol only for bullet
            numberStylePanel.enabled = rbNumbered.value;  // 番号の種類は番号リストのみ有効 / Number style only for numbered
            delimiterPanel.enabled = rbNumbered.value;    // 区切り文字は番号リストのみ有効 / Delimiter only for numbered
            chkZeroPad.enabled = rbStyleNumber.value;     // ゼロ埋めは「数字」のみ有効 / Zero padding only for the numbers style
            chkResetPerFrame.enabled = targetSelection.length > 1; // フレームごとにリセットは複数選択時のみ / per-frame reset only when 2+ frames selected
            formatPanel.enabled = !rbNone.value;          // 書式は「なし」では未使用 / Format unused for "none"

            // 区切り文字カラーは「数字／ABC／abc かつ区切り文字あり」のときだけ有効 / Delimiter color only when number/ABC/abc with a delimiter
            var delimiterColorApplicable = twoStopNumbered && currentDelimiter() !== "";
            delimiterColorUI.label.enabled = delimiterColorApplicable;
            delimiterColorUI.swatch.enabled = delimiterColorApplicable;

            // マーカー位置は数字/ABC/abc のみ有効、本文位置は全種類で有効（ハンギングON時のみ左インデント参照でディム）
            // Marker position only for number/ABC/abc; body position for all types (dimmed under hanging since it follows the left indent)
            var hanging = chkHanging.value && !rbNone.value;
            tabStopRow1.enabled = twoStopNumbered;           // マーカー位置 / marker position
            inputTabStop2.enabled = !rbNone.value && !hanging; // 本文位置（ハンギングでディム）/ body position (dimmed under hanging)
            syncHangingTabStop();                              // 本文位置に左インデントを反映 / mirror left indent into the body position
        }

        /* プレビューを更新（直前のプレビューを app.undo() で巻き戻してから再適用）/ Refresh preview (undo the previous preview, then reapply) */
        function updatePreview() {
            refreshList();
            // 直前のプレビュー1回ぶんを巻き戻す（実機検証: 1プレビュー = app.undo() 1回で完全復元）/ Roll back the previous preview pass (verified: one preview = one app.undo())
            if (previewState.isUndo) { try { app.undo(); app.redraw(); } catch (eUndo) {} previewState.isUndo = false; }
            // 並べ替えはJS側 baseline で管理しているため、undo 後に並べ替え済みの素テキストを書き戻してから付与
            // Reordering lives in the JS-side baseline, so after undo write back the (reordered) plain text, then apply
            restoreBaseline(baseline);
            applyListMarkers(targetSelection, {
                listType: currentType(),
                tabStop1Pt: currentTabStop1Pt(),
                tabStop2Pt: currentTabStop2Pt(),
                tab1Alignment: currentTab1Alignment(),
                bulletMark: currentBulletMark(),
                numberStyle: currentStyle(),
                zeroPad: chkZeroPad.value,
                startNumber: currentStartNumber(),
                resetPerFrame: chkResetPerFrame.value,
                delimiter: currentDelimiter(),
                baseline: baseline,
                format: currentFormat(),
                paragraphFormat: currentParagraphFormat()
            });
            previewState.isUndo = true; // 次回はこのプレビューを巻き戻す / next call rolls this one back
            app.redraw();
        }

        /* 種類変更時は既定値・ディム状態を更新してからプレビュー / On type change, reset defaults, refresh dim state, then preview */
        function onTypeChange() {
            applyTypeDefaults();
            // 比率: 番号リストは100へ戻す / 箇条書きは選択中の記号の比率へ / Scale: reset to 100 for numbered; use the selected symbol's scale for bullet
            if (rbNumbered.value) {
                inputScale.text = "100";
                rbAlignRight.value = true; // 番号は右揃えが既定 / numbered defaults to right
            } else if (rbBullet.value) {
                for (var bi = 0; bi < bulletRadios.length; bi++) {
                    if (bulletRadios[bi].value) { inputScale.text = String(BULLET_SYMBOLS[bi].scale); break; }
                }
                rbAlignLeft.value = true; // 箇条書きは左揃えが既定 / bullet defaults to left
            }
            updateEnabledStates();
            updatePreview();
        }

        /* 番号スタイル変更時はゼロ埋めのディムを更新してからプレビュー / On style change, refresh zero-pad dim state then preview */
        function onStyleChange() {
            chkZeroPad.enabled = rbStyleNumber.value; // ゼロ埋めは「数字」のみ / zero padding only for the numbers style
            // 区切り文字: 白丸・黒丸数字は「なし」, それ以外（数字・ABC・abc）は「.」/ Delimiter: "none" for circled (white/black), "." otherwise
            var isCircled = rbStyleCircledWhite.value || rbStyleCircledBlack.value;
            var autoDelimIndex = isCircled ? 0 : 1; // 0="none" / 1="."
            delimiterRadios[autoDelimIndex].value = true;
            // 揃え種類の既定: 数字=右 / ABC・abc=中央（丸数字は左固定でディム）/ Alignment default: number=right, ABC/abc=center (circled is fixed left & dimmed)
            if (rbStyleUpperAlpha.value || rbStyleLowerAlpha.value) {
                rbAlignCenter.value = true;
            } else if (rbStyleNumber.value) {
                rbAlignRight.value = true;
            }
            // タブストップ既定を種類に合わせて更新（数字/ABC/abc=1.5/2.0、白丸・黒丸数字=1.2の1ストップ）/ Reset tab stops to the style's defaults (number/ABC/abc = 1.5/2.0; circled = single ×1.2 stop)
            applyTypeDefaults();
            updateEnabledStates(); // 丸数字は2つ目タブストップ・揃え種類をディム / circled dims the 2nd stop & alignment
            updatePreview();
            // 丸数字で21以上になる場合は範囲外（素の数字）になるため警告 / Warn when circled numbering exceeds 20 (the rest become plain numbers)
            if (isCircled && largestNumberValue() > 20) {
                alert(getLabel('alert.circledLimit'));
            }
        }

        /* 箇条書き記号変更時は比率・本文タブストップを記号ごとの既定へ / On bullet change, apply each symbol's scale and body tab stop */
        function onBulletChange() {
            for (var bi = 0; bi < bulletRadios.length; bi++) {
                if (bulletRadios[bi].value) { inputScale.text = String(BULLET_SYMBOLS[bi].scale); break; }
            }
            applyTypeDefaults();   // 本文位置を記号の bodyRatio（•/-=0.8 など）で更新 / refresh body position with the glyph's bodyRatio (e.g. 0.8 for •/-)
            updateEnabledStates(); 
            updatePreview();
        }

        /* まとめてリセット（比率・ベースライン・カラー・インデント・タブストップ）/ Reset scale, baseline, color, indent, and tab stops together */
        function resetAllValues() {
            inputScale.text = "100";                                   // 比率 / scale
            inputBaseline.text = "0";                                  // ベースラインシフト / baseline shift
            inputStartNumber.text = String(START_NUMBER);              // 開始番号を1へ / start number back to 1
            // カラー（記号／番号・区切り文字とも基準色へ）/ color (marker-number & delimiter back to base)
            var resetColor = getBaseFillColor(targetSelection);
            colorSwatch._aiColor = resetColor;
            delimiterColorSwatch._aiColor = resetColor;
            try { colorSwatch.hide(); colorSwatch.show(); } catch (e) { }
            try { delimiterColorSwatch.hide(); delimiterColorSwatch.show(); } catch (eD) { }
            inputSpaceAfter.text = "0";                                // 段落後のアキ / space after
            inputLeftIndent.text = "0";                                // 左インデント（1行目は内部で正負反転）/ left indent (first-line derived internally)
            // 左・1行目インデントを実際に0へ（ハンギングOFFでも確実にクリア）/ Force left & first-line indent to 0 (even when hanging is off)
            // すべてのタブストップを削除（baselineの控えを空にし、入力欄は種類の既定へ）/ Delete all tab stops (clear snapshot; reset inputs to type defaults)
            for (var i = 0; i < baseline.length; i++) {
                baseline[i].tabStops = [];
                if (baseline[i].paraFormat) {
                    baseline[i].paraFormat.spaceAfter = 0;     // 段落後のアキも控えごと0へ（restoreBaselineで元値に戻るのを防ぐ）/ clear snapshot space-after too (else restoreBaseline brings the old value back)
                    baseline[i].paraFormat.leftIndent = 0;
                    baseline[i].paraFormat.firstLineIndent = 0;
                }
            }
            applyTypeDefaults();
            updateEnabledStates();
            updatePreview();
        }

        /* 操作のたびにプレビュー反映 / Reflect preview on every change */
        rbBullet.onClick = rbNumbered.onClick = rbNone.onClick = onTypeChange;
        // 番号スタイルのラジオ: クリックで排他選択してからスタイル変更処理 / number-style radios: enforce exclusivity, then handle the change
        for (var nsWireIndex = 0; nsWireIndex < numberStyleRadios.length; nsWireIndex++) {
            numberStyleRadios[nsWireIndex].onClick = function () {
                selectNumberStyleExclusive(this);
                onStyleChange();
            };
        }
        chkZeroPad.onClick = updatePreview;
        // 開始番号・フレームごとにリセット / start number & per-frame reset
        inputStartNumber.onChanging = updatePreview;
        changeValueByArrowKey(inputStartNumber, false, updatePreview);
        chkResetPerFrame.onClick = updatePreview;
        for (var bulletWireIndex = 0; bulletWireIndex < bulletRadios.length; bulletWireIndex++) {
            bulletRadios[bulletWireIndex].onClick = onBulletChange;
        }
        rbAlignLeft.onClick = rbAlignCenter.onClick = rbAlignRight.onClick = updatePreview; // 1つ目の揃え種類 / 1st-stop alignment
        for (var delimWireIndex = 0; delimWireIndex < delimiterRadios.length; delimWireIndex++) {
            delimiterRadios[delimWireIndex].onClick = function () {
                updateEnabledStates(); // 区切り文字なし→カラー行をディム / dim the color row when delimiter is "none"
                updatePreview();
            };
        }

        /* 並べ替えボタン / Reorder buttons */
        btnMoveTop.onClick = function () { moveSelectedRowsTo("top"); };
        btnMoveUp.onClick = function () { moveSelectedRows(-1); };
        btnMoveDown.onClick = function () { moveSelectedRows(1); };
        btnMoveBottom.onClick = function () { moveSelectedRowsTo("bottom"); };
        previewList.onChange = updateMoveButtonEnabled;
        btnSortByName.onClick = sortRowsByName;
        btnSortByValue.onClick = function () { sortRowsByValue(false); };
        btnSortByValueDesc.onClick = function () { sortRowsByValue(true); };
        inputTabStop1.onChanging = updatePreview;
        inputTabStop2.onChanging = updatePreview;

        fontFamilyDropdown.onChange = function () {
            populateFontStyleDropdown(fontStyleDropdown, fontFamilyDropdown.selection, null);
            updatePreview();
        };
        fontStyleDropdown.onChange = updatePreview;

        inputScale.onChanging = updatePreview;
        inputBaseline.onChanging = updatePreview;
        inputLeading.onChanging = updatePreview;
        inputSpaceAfter.onChanging = updatePreview;
        // ハンギング対応: ディム更新（インデント＋本文タブストップ）→プレビュー / Hanging: refresh dim (indent + body tab stop) then preview
        chkHanging.onClick = function () {
            // ONにしたらインデントを設定: 箇条書き=文字サイズ / 番号リスト=文字サイズの150%
            // On enable, set the indent: bullet = font size, numbered = font size × 150%
            if (chkHanging.value) {
                var indentPt = rbNumbered.value ? (baseFontSize * 1.5) : baseFontSize;
                inputLeftIndent.text = String(toDisplayUnit(indentPt));
            }
            updateHangingEnabled();
            updateEnabledStates(); // 本文タブストップにインデント値を反映 / mirror indent into the body tab stop
            updatePreview();
        };
        // インデント変更で本文タブストップを自動更新→プレビュー（1行目インデントは内部で正負反転）/ Indent change syncs the body tab stop, then preview (first-line indent is negated internally)
        inputLeftIndent.onChanging = function () { syncHangingTabStop(); updatePreview(); };

        /* 和文フォントのみ切り替え時はドロップダウンを再構築 / Rebuild the dropdown when toggling Japanese-only */
        chkJPOnly.onClick = function () {
            var currentFontName = getSelectedFontName(fontFamilyDropdown, fontStyleDropdown) || getBaseFontName(targetSelection);
            populateFontFamilyDropdown(fontFamilyDropdown, currentFontName, chkJPOnly.value);
            populateFontStyleDropdown(fontStyleDropdown, fontFamilyDropdown.selection, currentFontName);
            updatePreview();
        };

        // ↑↓キー（Shift=±10, Option=±0.1）で増減し、プレビュー更新 / Arrow keys (Shift=±10, Option=±0.1) adjust and refresh preview
        changeValueByArrowKey(inputTabStop1, false, updatePreview);
        changeValueByArrowKey(inputTabStop2, false, updatePreview);
        changeValueByArrowKey(inputScale, false, updatePreview);    // 比率% は負値不可 / scale % cannot be negative
        changeValueByArrowKey(inputBaseline, true, updatePreview);  // ベースラインシフトは負値可 / baseline shift allows negatives
        changeValueByArrowKey(inputLeading, false, updatePreview);      // 行送りは負値不可 / leading cannot be negative
        changeValueByArrowKey(inputSpaceAfter, false, updatePreview);   // 段落後のアキは負値不可 / space after cannot be negative
        changeValueByArrowKey(inputLeftIndent, false, function () { syncHangingTabStop(); updatePreview(); }); // インデント→本文タブストップ同期 / indent syncs body tab stop

        // ツールチップ（意味が伝わりにくいラベル・入力に補足）/ Tooltips for labels and inputs whose meaning isn't obvious
        tabStopLabel1.helpTip = inputTabStop1.helpTip = getLabel('tooltip.tabStop1');
        tabStopLabel2.helpTip = inputTabStop2.helpTip = getLabel('tooltip.tabStop2');
        rbAlignLeft.helpTip = rbAlignCenter.helpTip = rbAlignRight.helpTip = getLabel('tooltip.tabAlign');
        scaleLabel.helpTip = inputScale.helpTip = getLabel('tooltip.scale');
        baselineLabel.helpTip = inputBaseline.helpTip = getLabel('tooltip.baseline');
        chkZeroPad.helpTip = getLabel('tooltip.zeroPad');
        inputStartNumber.helpTip = getLabel('tooltip.startNumber');
        chkResetPerFrame.helpTip = getLabel('tooltip.resetPerFrame');

        fontLabel.helpTip = fontFamilyDropdown.helpTip = getLabel('tooltip.font');
        fontStyleLabel.helpTip = fontStyleDropdown.helpTip = getLabel('tooltip.fontStyle');

        chkJPOnly.helpTip = getLabel('tooltip.jpOnly');
        markerColorUI.label.helpTip = markerColorUI.swatch.helpTip = getLabel('tooltip.color');
        delimiterColorUI.label.helpTip = delimiterColorUI.swatch.helpTip = getLabel('tooltip.delimiterColor');
        leadingLabel.helpTip = inputLeading.helpTip = getLabel('tooltip.leading');
        spaceAfterLabel.helpTip = inputSpaceAfter.helpTip = getLabel('tooltip.spaceAfter');
        chkHanging.helpTip = getLabel('tooltip.hanging');
        leftIndentLabel.helpTip = inputLeftIndent.helpTip = getLabel('tooltip.leftIndent');
        btnShowHidden.helpTip = getLabel('tooltip.showHidden');
        btnReset.helpTip = getLabel('tooltip.reset');

        var committed = false;

        btnOK.onClick = function () {
            // プレビュー分を app.undo() で巻き戻してから、本番設定で1回だけ適用（履歴を最小化）
            // Roll back the preview via app.undo(), then apply once with the final settings (minimal history)
            if (previewState.isUndo) { try { app.undo(); app.redraw(); } catch (eUndo) {} previewState.isUndo = false; }
            restoreBaseline(baseline);
            applyListMarkers(targetSelection, {
                listType: currentType(),
                tabStop1Pt: currentTabStop1Pt(),
                tabStop2Pt: currentTabStop2Pt(),
                tab1Alignment: currentTab1Alignment(),
                bulletMark: currentBulletMark(),
                numberStyle: currentStyle(),
                zeroPad: chkZeroPad.value,
                startNumber: currentStartNumber(),
                resetPerFrame: chkResetPerFrame.value,
                delimiter: currentDelimiter(),
                baseline: baseline,
                format: currentFormat(),
                paragraphFormat: currentParagraphFormat()
            });
            app.redraw();
            committed = true;
            dlg.close(1);
        };
        btnCancel.onClick = function () { dlg.close(0); };

        /* OK以外で閉じた場合はプレビューを取り消す / Roll back preview unless committed via OK */
        dlg.onClose = function () {
            if (!committed) {
                // プレビュー分を巻き戻し、原状（並べ替えは baseline 準拠）へ / Roll back the preview, then restore (reorder follows baseline)
                if (previewState.isUndo) { try { app.undo(); app.redraw(); } catch (eUndo) {} previewState.isUndo = false; }
                restoreBaseline(baseline);
                app.redraw();
            }
            return true;
        };

        // 現在の状態から初期選択を復元（「現状の続き」として編集）/ Initialize from the current state so editing continues from it
        applyDetectedState(detectInitialState(targetSelection));

        // 表示前に初期状態（ディム）だけ反映（プレビューは onShow から起動）/ Apply initial dim state before showing (preview is kicked off from onShow)
        updateEnabledStates();
        updateHangingEnabled();       // ハンギング対応の初期ディム / initial dim state for hanging

        // 初期プレビューは onShow から起動（同期実行だと undo トランザクションが閉じず履歴が残るため）
        // Kick off the first preview from onShow (a synchronous pre-show preview leaves the undo transaction open → history piles up)
        dlg.onShow = function () { updatePreview(); };
        dlg.show();
    }

    // =========================================
    // 入力補助 / Input helpers
    // =========================================

    /* ↑↓キーで値を増減（Shift=±10, Option=±0.1）。プレビューは keyup まで遅延し、値が変わった時だけ1回実行
       Arrow keys adjust the value (Shift=±10, Option=±0.1). Preview is deferred to keyup and runs once only when the value actually changed
       - keydown: 値の更新のみ（押しっぱなしのオートリピートでもプレビューしない）/ keydown: update the value only (no preview, even on auto-repeat)
       - keyup:   keydown 前の値と比較し、変化があればプレビュー1回 / keyup: compare with the value before keydown and preview once if it changed */
    function changeValueByArrowKey(editText, allowNegative, onUpdate) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return; // ↑↓以外は通常入力に任せる / leave non-arrow keys to normal input

            var value = Number(editText.text);
            if (isNaN(value)) return;

            // 一連の ↑↓ 操作の開始値を控える（keyup での変化判定用。連続押下では最初の1回だけ記録）
            // Remember the value at the start of this arrow burst (for the keyup change check; recorded only on the first press of a burst)
            if (editText._arrowBaseValue == null) editText._arrowBaseValue = value;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 with Shift
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                // Optionキー押下時は0.1単位で増減 / Step by 0.1 with Option
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

            if (!allowNegative && value < 0) value = 0;

            editText.text = value; /* 値は即時反映、プレビューは keyup まで遅延 / reflect the value now; defer preview to keyup */
        });

        editText.addEventListener("keyup", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var baseValue = editText._arrowBaseValue;
            editText._arrowBaseValue = null; // バースト終了 / end of burst
            // 値が変わっていない（下限張り付きなど）ならプレビューしない / skip preview when the value did not change (e.g. clamped at the lower bound)
            if (baseValue != null && Number(editText.text) === baseValue) return;
            if (typeof onUpdate === "function") onUpdate();
        });
    }

    // =========================================
    // リスト付与 / Apply list
    // =========================================

    /* 丸数字スタイルか（箇条書きと同じく先頭タブなし・タブストップ1つで扱う）/ Whether the style is circled (laid out like a bullet: no leading tab, single tab stop) */
    function isCircledStyle(style) {
        return style === "circledWhite" || style === "circledBlack";
    }

    /* 行頭の既存マーカー（箇条書き記号・各種番号）を除去 / Remove existing leading marker (bullet or any number style) */
    function stripListMarker(line) {
        // 生成済みマーカー（先頭タブ + 番号記号 + 区切り文字(任意) + タブ）を優先的に除去 / Strip our generated marker (leading tab + glyph + optional delimiter + tab)
        var generated = /^\t(?:[①-⑳]|[❶-❿]|[⓫-⓴]|[A-Za-z]+|[〇一二三四五六七八九十百千]+|\d+)[.:：|]?\t/;
        if (generated.test(line)) return line.replace(generated, "");

        // 丸数字（先頭タブなしで「丸数字 + タブ/スペース」）を除去。丸数字は一意なので区切りなしでも対象 / Strip circled marker (no leading tab: circled glyph + tab/space). Circled glyphs are unambiguous, so no delimiter is required
        var circledLed = /^(?:[①-⑳]|[❶-❿]|[⓫-⓴])[\t 　]*/;
        if (circledLed.test(line)) return line.replace(circledLed, "");

        // 手打ちの番号リスト（先頭タブなし・「数字/ABC/abc + 区切り + スペース/タブ」）を除去
        // 区切り文字を必須にして本文（例: "Apple is" / "Mr. Smith"）の誤除去を抑える。"12.5" は区切り直後が数字で [\t 　]+ に一致せず対象外
        // Strip hand-typed lists (no leading tab: "number/ABC/abc + delimiter + space/tab"). A delimiter is required so body text is mostly preserved; "12.5" is excluded because a digit (not whitespace) follows the delimiter
        var spacedLed = /^(?:[A-Za-z]+|\d+)[.:：|][\t 　]+/;
        if (spacedLed.test(line)) return line.replace(spacedLed, "");

        // 元テキストのマーカー（箇条書き記号 / 数字.）を除去。数字.の直後が数字の場合（例: 12.5）は対象外
        // 中黒は異体字（・=U+30FB / ･=U+FF65 / ·=U+00B7）も対象。行頭の空白・タブも一緒に除去
        // 文字クラスは BULLET_SYMBOLS の記号をすべて含めること（記号を追加したらここにも追加）
        // 「-」だけは直後がタブ/空白のときのみ除去（本文の行頭ハイフン例: -5℃ / -10% を守る）
        // Strip original markers (bullet symbols / number.); "number." followed by a digit (e.g. 12.5) is left intact.
        // Middle-dot variants (・ U+30FB / ･ U+FF65 / · U+00B7) are included; leading whitespace/tab is removed too.
        // The class must cover every BULLET_SYMBOLS glyph (add new glyphs here too).
        // "-" is stripped only when immediately followed by a tab/space, so body text like "-5℃" / "-10%" is preserved.
        return line.replace(/^[\t 　]*(?:[・･·•◦●○◎□■◆◇✓]|-(?=[\t 　])|\d+\.(?!\d))[\t 　]*/, "");
    }

    /* 1フレーム分の行へマーカーを付与（連番カウンタを共有）/ Build marked lines for one frame (shares the number counter) */
    function buildMarkedLines(frameText, listType, options, numberCounter, padWidth) {
        // 改行コードを統一 / Normalize line breaks
        var lines = splitIntoLines(frameText);

        var bulletMark = options.bulletMark || BULLET_MARK;

        for (var j = 0; j < lines.length; j++) {
            // 箇条書き↔番号リストの切り替えに備え、まず既存マーカーを除去（=「なし」を経由）/ Strip existing marker first so bullet↔numbered switches cleanly
            var line = stripListMarker(lines[j]);

            // 空行（空白のみ含む）は常にスキップ：マーカーも番号も付けない / Always skip empty (whitespace-only) lines: no marker, no number
            if (!/\S/.test(line)) {
                lines[j] = line;
                continue;
            }

            if (listType === "numbered") {
                var numStr = formatNumber(numberCounter.value, options.numberStyle, padWidth, options.delimiter);
                if (isCircledStyle(options.numberStyle)) {
                    // 丸数字は箇条書きと同じ構造（丸数字 + tab + 本文, タブストップ1つ）/ Circled: bullet-like (glyph + tab + text, one tab stop)
                    line = numStr + "\t" + line;
                } else {
                    // tab + 番号記号 + tab + 本文 / tab + number glyph + tab + text
                    line = "\t" + numStr + "\t" + line;
                }
                numberCounter.value++;
            } else if (listType === "bullet") {
                line = bulletMark + "\t" + line;
            }
            // listType === "none" は除去のみ / "none" only strips

            lines[j] = line;
        }
        return lines;
    }

    /* 各テキストフレームの行頭を、いったん「なし」に戻してから指定種類を付与 / Reset each line to "none" first, then apply the chosen type */
    /* options: { listType, tabStop1Pt, tabStop2Pt, numberStyle, format } */
    function applyListMarkers(selection, options) {
        var listType = options.listType;
        var startNumber = (options.startNumber != null) ? options.startNumber : START_NUMBER;
        var numberCounter = { value: startNumber };

        // ゼロ埋めの桁数（数字スタイル時のみ。最大の番号の桁数に合わせる）/ Zero-pad width (numbers style only; based on the largest number)
        // フレームごとにリセット時は最大フレームの件数で桁を決める / per-frame reset sizes padding by the biggest frame
        var padWidth = 0;
        if (listType === "numbered" && options.zeroPad && options.numberStyle === "number") {
            var itemCount = options.resetPerFrame ? maxNumberedItemsPerFrame(selection) : countNumberedItems(selection);
            var maxNumber = startNumber + itemCount - 1;
            padWidth = String(maxNumber).length;
        }

        for (var i = 0; i < selection.length; i++) {
            var selectedObject = selection[i];

            // テキストフレームのみ対象 / TextFrame only
            if (selectedObject.typename !== "TextFrame") {
                continue;
            }

            // フレームごとにリセット時は各フレームの先頭で開始番号へ戻す / restart at each frame when requested
            if (options.resetPerFrame) numberCounter.value = startNumber;

            var lines = buildMarkedLines(selectedObject.contents, listType, options, numberCounter, padWidth);
            selectedObject.contents = lines.join("\r");

            // contents 再設定で失われた本文の文字属性を復元 / Restore body attributes lost by resetting contents
            var baselineEntry = findBaselineEntry(options.baseline, selectedObject);
            if (baselineEntry) restoreBodyAttributes(selectedObject, baselineEntry.bodyAttrsPerLine);

            // 行頭マーカーの送り位置をtabストップで揃える（「なし」では設定しない）/ Align the marker via tab stops (skip for "none")
            // マーカー位置の揃えはラジオ指定（既定は右）/ marker-column alignment comes from the radios (default right)
            // 本文位置（tabStop2Pt）は全種類で本文の開始位置。マーカー位置（tabStop1Pt）は数字/ABC/abc の番号列のみ。
            // Body position (tabStop2Pt) starts the text for every type; marker position (tabStop1Pt) is only the number column for number/ABC/abc.
            var tab1Alignment = options.tab1Alignment || TabStopAlignment.Right;
            if (listType === "numbered" && !isCircledStyle(options.numberStyle)) {
                // 数字/ABC/abc: マーカー位置=番号（指定の揃え）, 本文位置=本文（左揃え）/ marker = number (chosen alignment), body = text (left)
                setTabStops(selectedObject, [
                    { position: options.tabStop1Pt, alignment: tab1Alignment },
                    { position: options.tabStop2Pt, alignment: TabStopAlignment.Left }
                ]);
            } else if (listType === "numbered" || listType === "bullet") {
                // 箇条書き・丸数字: 本文位置の1か所のみ（左揃え固定）/ Bullet & circled: a single left-aligned stop at the body position
                setTabStops(selectedObject, [
                    { position: options.tabStop2Pt, alignment: TabStopAlignment.Left }
                ]);
            } else {
                // 「なし」はタブストップをすべて削除 / "none" clears all tab stops
                setTabStops(selectedObject, []);
            }

            // 行頭マーカーに書式（フォント・サイズ・ベースラインシフト）を適用 / Apply marker format
            applyMarkerFormat(selectedObject, listType, options.format, options.numberStyle, options.delimiter);

            // 段落の書式（行送り・段落前/後のアキ）を適用 / Apply paragraph format (leading, space before/after)
            applyParagraphFormat(selectedObject, options.paragraphFormat);
        }
    }

    /* 段落の書式（行送り・段落前/後のアキ）をフレーム全体へ適用。null は変更しない / Apply paragraph format to the whole frame; null leaves a value unchanged */
    function applyParagraphFormat(frame, paragraphFormat) {
        if (!paragraphFormat) return;
        // 各設定は個別に try で囲む。1項目（例: 行送り）の例外で後続（spaceAfter・インデント）が
        // 巻き込まれてスキップされるのを防ぐ / Guard each setter separately so a failure in one
        // (e.g. leading) does not skip the rest (spaceAfter, indents)
        if (paragraphFormat.leadingPt != null) {
            // 行送りは文字属性（自動行送りを解除して固定値に）/ Leading is a character attribute (disable auto-leading)
            try { frame.textRange.characterAttributes.autoLeading = false; } catch (eAuto) { }
            try { frame.textRange.characterAttributes.leading = paragraphFormat.leadingPt; } catch (eLead) { }
        }
        if (paragraphFormat.spaceAfterPt != null) {
            try { frame.textRange.paragraphAttributes.spaceAfter = paragraphFormat.spaceAfterPt; } catch (eSA) { }
        }
        // インデント（ハンギング対応OFFのときは null で渡され、変更しない）/ Indents (null when hanging is off → unchanged)
        if (paragraphFormat.leftIndentPt != null) {
            try { frame.textRange.paragraphAttributes.leftIndent = paragraphFormat.leftIndentPt; } catch (eLI) { }
        }
        if (paragraphFormat.firstLineIndentPt != null) {
            try { frame.textRange.paragraphAttributes.firstLineIndent = paragraphFormat.firstLineIndentPt; } catch (eFLI) { }
        }
    }

    // =========================================
    // 並べ替え / Reorder
    // =========================================

    /**
     * baseline の各フレームの行数を取得する
     * @param {Array<object>} baseline - captureBaseline() が返した控え
     * @returns {number[]} フレームごとの行数
     */
    function getFrameLineCounts(baseline) {
        var counts = [];
        for (var i = 0; i < baseline.length; i++) {
            counts.push(splitIntoLines(baseline[i].contents).length);
        }
        return counts;
    }

    /**
     * 全フレームの本文（マーカー除去後）と文字属性を1つのフラット配列で取得する
     * @param {Array<object>} baseline - captureBaseline() が返した控え
     * @returns {Array<{text: string, attrs: Array<object>}>} 行ごとの本文と文字属性
     */
    function getFlatBodyEntries(baseline) {
        var flat = [];
        for (var i = 0; i < baseline.length; i++) {
            var lines = splitIntoLines(baseline[i].contents);
            var attrsPerLine = baseline[i].bodyAttrsPerLine || [];
            for (var j = 0; j < lines.length; j++) {
                flat.push({
                    text: stripListMarker(lines[j]),
                    attrs: attrsPerLine[j] || []
                });
            }
        }
        return flat;
    }
    /**
     * フラット配列の本文と文字属性を、元のフレーム行数どおりに baseline へ書き戻す
     * @param {Array<object>} baseline - captureBaseline() が返した控え
     * @param {Array<{text: string, attrs: Array<object>}>} flatEntries - 書き戻す行データ
     * @returns {void}
     */
    function writeFlatBodyEntries(baseline, flatEntries) {
        var counts = getFrameLineCounts(baseline);
        var index = 0;
        for (var i = 0; i < baseline.length; i++) {
            var lineCount = counts[i];
            var lines = [];
            var attrsPerLine = [];
            for (var j = 0; j < lineCount; j++) {
                var entry = flatEntries[index + j];
                lines.push(entry ? entry.text : "");
                attrsPerLine.push(entry ? entry.attrs : []);
            }
            baseline[i].contents = lines.join("\r");
            baseline[i].bodyAttrsPerLine = attrsPerLine;
            index += lineCount;
        }
    }

    // =========================================
    // 番号スタイル / Number style
    // =========================================

    /* 番号を選択スタイルの文字列へ変換（末尾に区切り文字を付与）/ Format a number into the chosen style (append the delimiter) */
    function formatNumber(n, style, padWidth, delimiter) {
        var core;
        if (style === "circledWhite") {
            core = toCircledWhite(n);
        } else if (style === "circledBlack") {
            core = toCircledBlack(n);
        } else if (style === "upperAlpha") {
            core = toAlphabet(n, true);
        } else if (style === "lowerAlpha") {
            core = toAlphabet(n, false);
        } else {
            // 数字（ゼロ埋め指定があれば桁をそろえる）/ Plain numbers (zero-padded when requested)
            core = (padWidth && padWidth > 0) ? zeroPadNumber(n, padWidth) : String(n);
        }
        return core + (delimiter || ""); // 区切り文字（なし=空文字）/ delimiter (none = "")
    }

    /* 指定桁数までゼロ埋め / Left-pad a number with zeros to the given width */
    function zeroPadNumber(n, width) {
        var text = String(n);
        while (text.length < width) text = "0" + text;
        return text;
    }

    /* 1フレームの番号付け対象行数（空行は除外）/ Numbered (non-empty) line count of one frame */
    function numberedItemsInFrame(frame) {
        var lines = splitIntoLines(frame.contents);
        var count = 0;
        for (var j = 0; j < lines.length; j++) {
            if (/\S/.test(stripListMarker(lines[j]))) count++; // 空行はスキップ / skip empty lines
        }
        return count;
    }

    /* 番号付け対象の総行数（継続時の桁そろえ・件数チェック用）/ Total numbered lines (for continuous padding & count checks) */
    function countNumberedItems(selection) {
        var total = 0;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") total += numberedItemsInFrame(selection[i]);
        }
        return total;
    }

    /* 1フレームあたりの最大行数（フレームごとにリセット時の桁そろえ・件数チェック用）/ Largest per-frame count (for per-frame-reset padding & checks) */
    function maxNumberedItemsPerFrame(selection) {
        var maxCount = 0;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename !== "TextFrame") continue;
            var count = numberedItemsInFrame(selection[i]);
            if (count > maxCount) maxCount = count;
        }
        return maxCount;
    }

    /* 白丸数字 ①〜⑳（範囲外は数字.）/ White circled 1-20 (fallback to "n.") */
    function toCircledWhite(n) {
        if (n >= 1 && n <= 20) return String.fromCharCode(0x2460 + (n - 1));
        return String(n); // 範囲外は素の数字（区切り文字は呼び出し側で付与）/ out of range: plain number
    }

    /* 黒丸数字 ❶〜⓴（範囲外は数字.）/ Black circled 1-20 (fallback to "n.") */
    function toCircledBlack(n) {
        if (n >= 1 && n <= 10) return String.fromCharCode(0x2776 + (n - 1)); // ❶〜❿
        if (n >= 11 && n <= 20) return String.fromCharCode(0x24EB + (n - 11)); // ⓫〜⓴
        return String(n); // 範囲外は素の数字（区切り文字は呼び出し側で付与）/ out of range: plain number
    }

    /* アルファベット a, b, c... / A, B, C...（範囲外は数字）/ Alphabet a-z or A-Z (fallback to number) */
    function toAlphabet(n, upper) {
        if (n >= 1 && n <= 26) return String.fromCharCode((upper ? 0x41 : 0x61) + (n - 1));
        return String(n);
    }

    // =========================================
    // 書式（マーカーへのフォント・サイズ・ベースラインシフト）/ Marker format
    // =========================================

    /* 行頭マーカーの文字にフォント・サイズ・ベースラインシフト・カラーを適用 / Apply font, size, baseline shift, and color to the marker characters */
    function applyMarkerFormat(frame, listType, format, numberStyle, delimiter) {
        if (!format || listType === "none") return;

        // 数字／ABC／abc では番号の末尾に区切り文字が付く（その分だけ別カラーにできる）/ For number/ABC/abc, the delimiter trails the number, so it can take a separate color
        var delimiterText = delimiter || "";
        var delimiterLen = (listType === "numbered" && !isCircledStyle(numberStyle)) ? delimiterText.length : 0;

        var font = null;
        if (format.fontName) {
            try { font = app.textFonts.getByName(format.fontName); } catch (e) { }
        }

        try {
            var paragraphs = frame.paragraphs;
            for (var p = 0; p < paragraphs.length; p++) {
                var paragraph = paragraphs[p];
                var paragraphText = paragraph.contents;
                if (!paragraphText || paragraphText.length === 0) continue;

                // マーカー文字の範囲 [markerStart, markerEnd) を求める / Determine the marker character range
                var markerStart, markerEnd;
                if (listType === "bullet" || (listType === "numbered" && isCircledStyle(numberStyle))) {
                    // 記号1字（箇条書き／丸数字）。「記号 + タブ」でない行（=空行など未付与）はスキップ
                    // single glyph (bullet / circled); skip lines that aren't "glyph + tab" (e.g. skipped empty lines)
                    if (paragraphText.charAt(1) !== "\t") continue;
                    markerStart = 0; markerEnd = 1;
                } else { // numbered (数字/ABC/abc): 先頭タブを飛ばし、次のタブまで / skip leading tab, up to the next tab
                    if (paragraphText.charAt(0) !== "\t") continue; // 先頭タブが無い行（未付与）はスキップ / skip lines without a leading tab
                    markerStart = 1; markerEnd = 1;
                    while (markerEnd < paragraphText.length && paragraphText.charAt(markerEnd) !== "\t") markerEnd++;
                }

                // 区切り文字部分の開始位置（マーカー範囲の末尾から delimiterLen 文字）/ Where the delimiter chars start (last delimiterLen chars of the marker range)
                var delimiterStart = markerEnd - delimiterLen;
                for (var c = markerStart; c < markerEnd; c++) {
                    if (c >= paragraph.characters.length) break;
                    var charAttr = paragraph.characters[c].characterAttributes;
                    if (font) { try { charAttr.textFont = font; } catch (e1) { } }
                    if (format.horizontalScale) { try { charAttr.horizontalScale = format.horizontalScale; } catch (e3) { } }
                    if (format.verticalScale) { try { charAttr.verticalScale = format.verticalScale; } catch (e4) { } }
                    try { charAttr.baselineShift = format.baselineShiftPt; } catch (e5) { }
                    // カラー: 区切り文字は専用色、それ以外（記号・番号）はマーカー色 / Color: delimiter chars use their own color, the rest use the marker color
                    var isDelimiterChar = (delimiterLen > 0 && c >= delimiterStart);
                    var fillColorToUse = isDelimiterChar ? format.delimiterFillColor : format.fillColor;
                    if (fillColorToUse) { try { charAttr.fillColor = fillColorToUse; } catch (e6) { } }
                }
            }
        } catch (eOuter) { }
    }

    /* 先頭テキストフレームのフォント名（PostScript名）を取得 / Get the first text frame's font name (PostScript name) */
    function getBaseFontName(selection) {
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") {
                try {
                    var font = selection[i].textRange.characterAttributes.textFont;
                    if (font) return font.name;
                } catch (e) { }
                break;
            }
        }
        return null;
    }

    /* フォント一覧をファミリー別にまとめる / Group Illustrator fonts by family */
    function buildFontFamilyMap(japaneseOnly) {
        var familyMap = {};
        var familyNames = [];

        for (var i = 0; i < app.textFonts.length; i++) {
            var font = app.textFonts[i];
            var familyName = font.family || font.name;
            var styleName = font.style || font.name;

            if (japaneseOnly && !isJapaneseFont(font)) continue;

            if (!familyMap[familyName]) {
                familyMap[familyName] = [];
                familyNames.push(familyName);
            }
            familyMap[familyName].push({ name: font.name, family: familyName, style: styleName });
        }

        familyNames.sort(function (a, b) { return (a < b) ? -1 : (a > b ? 1 : 0); });
        for (var f = 0; f < familyNames.length; f++) {
            familyMap[familyNames[f]].sort(function (a, b) { return (a.style < b.style) ? -1 : (a.style > b.style ? 1 : 0); });
        }

        return { names: familyNames, map: familyMap };
    }

    /* フォント名からTextFontを取得 / Get a TextFont by its internal name */
    function findTextFontByName(fontName) {
        if (!fontName) return null;
        for (var i = 0; i < app.textFonts.length; i++) {
            if (app.textFonts[i].name === fontName) return app.textFonts[i];
        }
        return null;
    }

    /* フォントファミリーのドロップダウンを作成 / Populate the font-family dropdown */
    function populateFontFamilyDropdown(dropdown, selectedFontName, japaneseOnly) {
        dropdown.removeAll();

        var data = buildFontFamilyMap(japaneseOnly);
        dropdown._fontFamilyMap = data.map;

        var selectedFont = findTextFontByName(selectedFontName);
        var selectedFamily = selectedFont ? (selectedFont.family || selectedFont.name) : null;
        var selectedIndex = 0;

        for (var i = 0; i < data.names.length; i++) {
            dropdown.add("item", data.names[i]);
            if (selectedFamily && data.names[i] === selectedFamily) selectedIndex = i;
        }

        if (dropdown.items.length > 0) dropdown.selection = Math.min(selectedIndex, dropdown.items.length - 1);
    }

    /* 選択中ファミリーに対応するスタイルのドロップダウンを作成 / Populate styles for the selected family */
    function populateFontStyleDropdown(dropdown, familySelection, selectedFontName) {
        dropdown.removeAll();
        dropdown._fontStyleNames = [];

        if (!familySelection || !familySelection.parent || !familySelection.parent._fontFamilyMap) return;

        var familyName = familySelection.text;
        var styles = familySelection.parent._fontFamilyMap[familyName] || [];
        var selectedIndex = 0;

        for (var i = 0; i < styles.length; i++) {
            dropdown.add("item", styles[i].style || styles[i].name);
            dropdown._fontStyleNames.push(styles[i].name);
            if (selectedFontName && styles[i].name === selectedFontName) selectedIndex = i;
        }

        if (dropdown.items.length > 0) dropdown.selection = Math.min(selectedIndex, dropdown.items.length - 1);
    }

    /* UIで選択されたファミリー＋スタイルからIllustrator内部フォント名を取得 / Resolve the selected family + style to Illustrator's internal font name */
    function getSelectedFontName(familyDropdown, styleDropdown) {
        if (!familyDropdown || !familyDropdown.selection || !styleDropdown || !styleDropdown.selection) return null;
        if (!styleDropdown._fontStyleNames) return null;
        return styleDropdown._fontStyleNames[styleDropdown.selection.index] || null;
    }

    /* 和文文字（ひらがな・カタカナ・漢字）を含むか / Whether the string contains Japanese characters */
    function hasJapaneseCharacters(str) {
        try {
            if (!str) return false;
            return /[぀-ゟ゠-ヿ一-鿿]/.test(String(str));
        } catch (e) {
            return false;
        }
    }

    /* 和文フォントかどうかを判定（AutoTouchType.jsx のロジックを流用）/ Detect a Japanese font (logic ported from AutoTouchType.jsx) */
    function isJapaneseFont(font) {
        if (!font) return false;

        var fontName = "", fontFamily = "", fontFull = "", fontPS = "";
        try { fontName = font.name ? String(font.name) : ""; } catch (e1) { }
        try { fontFamily = font.family ? String(font.family) : ""; } catch (e2) { }
        try { fontFull = font.fullName ? String(font.fullName) : ""; } catch (e3) { }
        try { fontPS = font.postScriptName ? String(font.postScriptName) : ""; } catch (e4) { }

        // 明示的に除外（ラテン／意図しないゴシック系）/ Explicit exclusions (Latin / unintended Gothic variants)
        var denyList = ["Apple LiGothic", "RyoGothicStd", "-KO", "-KL", "LogoArl", "Kana"];
        for (var di = 0; di < denyList.length; di++) {
            var denyName = denyList[di];
            if (!denyName) continue;
            if (fontName.indexOf(denyName) !== -1 || fontFamily.indexOf(denyName) !== -1 ||
                fontFull.indexOf(denyName) !== -1 || fontPS.indexOf(denyName) !== -1) return false;
        }

        // 識別子に和文文字が含まれていれば和文 / Direct Japanese-character hit in any identifier
        if (hasJapaneseCharacters(fontName) || hasJapaneseCharacters(fontFamily) || hasJapaneseCharacters(fontFull)) {
            return true;
        }

        // キーワード判定（和文ファミリー・主要ブランド）/ Keyword-based heuristic (JP families and common brands)
        var jpKeywords = [
            "ゴシック", "明朝", "丸ゴ", "教科書", "楷書",
            "Mincho", "Maru",
            "Hiragino", "ヒラギノ",
            "Yu Gothic", "Yu Mincho", "游ゴシック", "游明朝",
            "Meiryo", "メイリオ",
            "MS Gothic", "MS Mincho", "MS ゴシック", "MS 明朝",
            "Kozuka", "小塚",
            "Morisawa", "モリサワ",
            "Ryumin", "Shin Go", "新ゴ",
            "Heisei", "平成",
            "Klee", "クレー",
            "Tsukushi", "筑紫",
            "A-OTF", "AP-OTF ", "-OTF",
            "FOT", "Pr6N", "Pr6",
            "Noto Sans JP", "Noto Serif JP",
            "Source Han", "源ノ角", "源ノ明",
            "Min2"
        ];
        for (var i = 0; i < jpKeywords.length; i++) {
            var keyword = jpKeywords[i];
            if (!keyword) continue;
            if (fontName.indexOf(keyword) !== -1 || fontFamily.indexOf(keyword) !== -1 ||
                fontFull.indexOf(keyword) !== -1 || fontPS.indexOf(keyword) !== -1) return true;
        }

        return false;
    }

    // =========================================
    // tabストップ / Tab stop
    // =========================================

    /* テキストの表示単位を取得（"text/units" を参照）/ Get the text display unit (reads "text/units") */
    function getTextUnitInfo() {
        var rulerUnit = app.preferences.getIntegerPreference("text/units");
        var label = "pt";
        var factor = 1.0; // 1単位あたりのポイント数 / Points per unit

        switch (rulerUnit) {
            case 0: label = "inch"; factor = 72.0; break;            // インチ / inch
            case 1: label = "mm"; factor = 72.0 / 25.4; break;     // ミリ / mm
            case 2: label = "pt"; factor = 1.0; break;             // ポイント / pt
            case 3: label = "pica"; factor = 12.0; break;            // パイカ / pica
            case 4: label = "cm"; factor = 72.0 / 2.54; break;     // センチ / cm
            case 5: label = "Q"; factor = 72.0 / 25.4 * 0.25; break; // 級 / Q
            case 6: label = "px"; factor = 1.0; break;             // ピクセル / px
            default: label = "pt"; factor = 1.0;
        }
        return { label: label, factor: factor };
    }

    /* 先頭テキストフレームの文字サイズ（pt）を取得 / Get the first text frame's font size (pt) */
    function getBaseFontSize(selection) {
        var fontSize = 12; // フォールバック / Fallback
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") {
                try {
                    var size = selection[i].textRange.characterAttributes.size;
                    if (size && !isNaN(size)) { fontSize = size; }
                } catch (e) { }
                break;
            }
        }
        return fontSize;
    }

    /* 先頭テキストフレームの行送り（pt）を取得。取得不可なら null / Get the first text frame's leading (pt); null if unavailable */
    function getBaseLeading(selection) {
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") {
                try {
                    var leading = selection[i].textRange.characterAttributes.leading;
                    if (leading != null && !isNaN(leading)) return leading;
                } catch (e) { }
                break;
            }
        }
        return null;
    }

    /* 指定の揃えでタブストップを生成 / Create a tab stop with the given alignment */
    function makeTab(positionPt, alignment) {
        var tab = new TabStopInfo();
        tab.alignment = alignment;
        tab.position = positionPt;
        return tab;
    }

    /* タブストップを指定内容だけに置き換える（前タイプの残存タブを消すため全消去してから設定）/ Replace all tab stops with the given specs (clears leftovers from the previous type) */
    /* tabSpecs: [{ position, alignment }, ...] */
    function setTabStops(frame, tabSpecs) {
        try {
            var newTabs = [];
            for (var s = 0; s < tabSpecs.length; s++) {
                newTabs.push(makeTab(tabSpecs[s].position, tabSpecs[s].alignment));
            }

            // 位置順にソート / Sort by position
            newTabs.sort(function (a, b) { return a.position - b.position; });

            frame.textRange.paragraphAttributes.tabStops = newTabs;
        } catch (e) { }
    }

    // =========================================
    // プレビュー状態の保存・復元 / Preview snapshot & restore
    // =========================================

    /* テキストとタブストップの現状を保存 / Snapshot contents and tab stops */
    function captureBaseline(selection) {
        var baseline = [];
        for (var i = 0; i < selection.length; i++) {
            var frame = selection[i];
            if (frame.typename !== "TextFrame") continue;

            var entry = { frame: frame, contents: frame.contents, tabStops: null, bodyAttrsPerLine: null, paraFormat: null };
            try {
                var existingTabs = frame.textRange.paragraphAttributes.tabStops;
                entry.tabStops = [];
                for (var t = 0; t < existingTabs.length; t++) {
                    // 復元用に位置と揃えだけ控える / Keep position and alignment for restore
                    entry.tabStops.push({ position: existingTabs[t].position, alignment: existingTabs[t].alignment });
                }
            } catch (e) { }
            // 段落の書式（行送り・自動行送り・段落前/後のアキ）を退避 / Snapshot paragraph format (leading, auto-leading, space before/after)
            entry.paraFormat = {};
            try { entry.paraFormat.leading = frame.textRange.characterAttributes.leading; } catch (eL) { }
            try { entry.paraFormat.autoLeading = frame.textRange.characterAttributes.autoLeading; } catch (eA) { }
            try { entry.paraFormat.spaceAfter = frame.textRange.paragraphAttributes.spaceAfter; } catch (eF) { }
            try { entry.paraFormat.leftIndent = frame.textRange.paragraphAttributes.leftIndent; } catch (eLI) { }
            try { entry.paraFormat.firstLineIndent = frame.textRange.paragraphAttributes.firstLineIndent; } catch (eFI) { }
            // 本文（マーカー以降）の文字属性を退避 / Snapshot the body character attributes (after the marker)
            entry.bodyAttrsPerLine = captureBodyAttributes(frame);
            baseline.push(entry);
        }
        return baseline;
    }

    /* 保存しておいた状態へ復元 / Restore the snapshot */
    function restoreBaseline(baseline) {
        for (var i = 0; i < baseline.length; i++) {
            var entry = baseline[i];
            try { entry.frame.contents = entry.contents; } catch (e) { }
            if (entry.tabStops !== null) {
                try {
                    var rebuiltTabs = [];
                    for (var t = 0; t < entry.tabStops.length; t++) {
                        var tab = new TabStopInfo();
                        tab.alignment = entry.tabStops[t].alignment;
                        tab.position = entry.tabStops[t].position;
                        rebuiltTabs.push(tab);
                    }
                    entry.frame.textRange.paragraphAttributes.tabStops = rebuiltTabs;
                } catch (e2) { }
            }
            // 段落の書式を復元 / Restore paragraph format
            if (entry.paraFormat) {
                if (entry.paraFormat.autoLeading != null) { try { entry.frame.textRange.characterAttributes.autoLeading = entry.paraFormat.autoLeading; } catch (eRA) { } }
                if (entry.paraFormat.leading != null) { try { entry.frame.textRange.characterAttributes.leading = entry.paraFormat.leading; } catch (eRL) { } }
                if (entry.paraFormat.spaceAfter != null) { try { entry.frame.textRange.paragraphAttributes.spaceAfter = entry.paraFormat.spaceAfter; } catch (eRF) { } }
                if (entry.paraFormat.leftIndent != null) { try { entry.frame.textRange.paragraphAttributes.leftIndent = entry.paraFormat.leftIndent; } catch (eRLI) { } }
                if (entry.paraFormat.firstLineIndent != null) { try { entry.frame.textRange.paragraphAttributes.firstLineIndent = entry.paraFormat.firstLineIndent; } catch (eRFI) { } }
            }
            restoreBodyAttributes(entry.frame, entry.bodyAttrsPerLine);
        }
    }

    // =========================================
    // 本文の文字属性の退避・復元 / Body character attributes snapshot & restore
    // =========================================
    // contents の再設定でフレーム全体の文字書式が初期化されるため、本文（マーカー以降）の属性を退避し復元する
    // Setting .contents resets the frame's character formatting, so the body (after the marker) is snapshotted and restored.

    /* 1文字分の主要な文字属性を控える / Snapshot the key character attributes of one character */
    function snapshotCharAttributes(characterAttr) {
        var snap = {};
        try { snap.textFont = characterAttr.textFont; } catch (e1) { }
        try { snap.size = characterAttr.size; } catch (e2) { }
        try { snap.horizontalScale = characterAttr.horizontalScale; } catch (e3) { }
        try { snap.verticalScale = characterAttr.verticalScale; } catch (e4) { }
        try { snap.baselineShift = characterAttr.baselineShift; } catch (e5) { }
        try { snap.tracking = characterAttr.tracking; } catch (e6) { }
        try { snap.fillColor = characterAttr.fillColor; } catch (e7) { }
        return snap;
    }

    /* 控えた文字属性を1文字へ復元 / Restore snapshotted attributes onto one character */
    function restoreCharAttributes(characterAttr, snap) {
        if (!snap) return;
        if (snap.textFont) { try { characterAttr.textFont = snap.textFont; } catch (e1) { } }
        if (snap.size != null) { try { characterAttr.size = snap.size; } catch (e2) { } }
        if (snap.horizontalScale != null) { try { characterAttr.horizontalScale = snap.horizontalScale; } catch (e3) { } }
        if (snap.verticalScale != null) { try { characterAttr.verticalScale = snap.verticalScale; } catch (e4) { } }
        if (snap.baselineShift != null) { try { characterAttr.baselineShift = snap.baselineShift; } catch (e5) { } }
        if (snap.tracking != null) { try { characterAttr.tracking = snap.tracking; } catch (e6) { } }
        if (snap.fillColor) { try { characterAttr.fillColor = snap.fillColor; } catch (e7) { } }
    }

    /* 各段落の本文（マーカー以降）の文字属性を退避 / Snapshot the body character attributes per paragraph */
    function captureBodyAttributes(frame) {
        var perLine = [];
        try {
            var paragraphs = frame.paragraphs;
            for (var p = 0; p < paragraphs.length; p++) {
                var text = paragraphs[p].contents;
                var markerLength = text.length - stripListMarker(text).length; // 既存マーカー長 / existing marker length
                var lineAttrs = [];
                for (var k = markerLength; k < text.length; k++) {
                    lineAttrs.push(snapshotCharAttributes(paragraphs[p].characters[k].characterAttributes));
                }
                perLine.push(lineAttrs);
            }
        } catch (e) { }
        return perLine;
    }

    /* 退避した本文属性を、再構築後の本文文字へ復元（本文文字数は不変なので末尾から数えて対応）/ Restore body attributes (body length is unchanged, so map from the tail) */
    function restoreBodyAttributes(frame, perLine) {
        if (!perLine) return;
        try {
            var paragraphs = frame.paragraphs;
            for (var p = 0; p < paragraphs.length && p < perLine.length; p++) {
                var lineAttrs = perLine[p];
                if (!lineAttrs || lineAttrs.length === 0) continue;

                var text = paragraphs[p].contents;
                var bodyStart = text.length - lineAttrs.length; // 本文は末尾 lineAttrs.length 文字 / body = last N characters
                if (bodyStart < 0) continue;

                for (var k = 0; k < lineAttrs.length; k++) {
                    var charIndex = bodyStart + k;
                    if (charIndex >= paragraphs[p].characters.length) break;
                    restoreCharAttributes(paragraphs[p].characters[charIndex].characterAttributes, lineAttrs[k]);
                }
            }
        } catch (e) { }
    }

    /* frame に対応する baseline エントリを取得 / Find the baseline entry for a frame */
    function findBaselineEntry(baseline, frame) {
        if (!baseline) return null;
        for (var i = 0; i < baseline.length; i++) {
            if (baseline[i].frame === frame) return baseline[i];
        }
        return null;
    }
})();