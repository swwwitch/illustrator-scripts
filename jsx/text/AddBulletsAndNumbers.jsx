#target illustrator
#targetengine "addBulletsAndNumbers"
#include "ColorPicker.jsx"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレームの各行頭に、箇条書き記号または連番を付与します。
ダイアログを開くと現在の行頭マーカーから設定を推定し、プレビューを見ながら「現状の続き」として編集できます。

詳細は README を参照してください。

### Overview

Adds a bullet or a sequence number to the start of each line in the selected text frames.
The dialog infers its settings from the markers already present, so you can carry on from the current state while watching a preview.

See the README for details.

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
    var DEFAULT_FONT_SIZE = 12; // 文字サイズを取得できなかったときの既定値(pt) / fallback font size (pt)
    var PREVIEW_DELAY_MS = 150; // 次の操作を待つ時間(ms)。この間に次の操作が来たら前の分は計算しない / how long to wait for the next input before previewing

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

    // 行揃えボタンの選択肢（左から並ぶ順）/ Justification buttons, in the order they appear
    var JUSTIFY_OPTIONS = [
        { id: "left", labelKey: "justify.left" },
        { id: "center", labelKey: "justify.center" },
        { id: "right", labelKey: "justify.right" },
        { id: "justifyLeft", labelKey: "justify.justifyLeft" },
        { id: "justifyAll", labelKey: "justify.justifyAll" }
    ];

    // 和文フォント判定のキーワード（フォント数ぶん走査するため、呼び出しごとに作り直さない）
    // Keywords for the Japanese-font check (kept out of the function; it runs once per font)
    var JP_FONT_DENY_KEYWORDS = ["Apple LiGothic", "RyoGothicStd", "-KO", "-KL", "LogoArl", "Kana"];
    var JP_FONT_KEYWORDS = [
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

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margins */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / column gutter */
    var WIDE_COLUMN_SPACING = 20;            /* 広めに空ける2カラムの間隔 / wide column gutter */
    var DENSE_SPACING  = 6;                  /* ラジオ・入力行を詰めるパネルの間隔 / spacing for dense panels */
    var FIELD_LABEL_WIDTH = 100;             /* 書式・位置ラベルの幅 / label width for format & position rows */
    var PARAGRAPH_LABEL_WIDTH = 110;         /* 段落設定ラベルの幅 / label width for paragraph rows */
    var REORDER_BUTTON_HEIGHT = 22;          /* 並べ替えボタンの高さ / reorder button height */
    var SWATCH_SIZE = 20;                    /* カラースウォッチの一辺 / color swatch size */
    var JUSTIFY_BUTTON_SIZE = 26;            /* 行揃えボタンの一辺 / justification button size */
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
     * 行の先頭に置くラベルを追加する
     * 幅を指定した場合は入力欄の位置がそろうよう右揃えにする（幅がないと右揃えは効かない）
     * @param {Group} row - 追加先の行グループ
     * @param {string} labelString - ラベル文字列
     * @param {number} [labelWidth] - ラベルの幅（省略時は内容にフィット）
     * @returns {StaticText} 追加したラベル
     */
    function addFieldLabel(row, labelString, labelWidth) {
        if (typeof labelWidth !== "number") return row.add("statictext", undefined, labelString);

        // 右揃えは生成時プロパティで渡す（後付けの代入だけでは効かない環境があるため、両方指定する）
        // Pass the justification as a creation property (a post-hoc assignment alone is ignored on some platforms)
        var label = row.add("statictext", undefined, labelString, { justify: "right" });
        label.preferredSize.width = labelWidth;
        label.justify = "right";
        return label;
    }

    /**
     * ラベル＋入力欄＋単位の1行を追加する
     * @param {Group|Panel} parent - 追加先
     * @param {string} labelString - ラベル文字列
     * @param {string} initialText - 入力欄の初期値
     * @param {number} [labelWidth] - ラベルの幅（指定すると右揃え。省略時は内容にフィット）
     * @param {string} [unitLabel] - 入力欄の後ろに置く単位（省略可）
     * @returns {{row: Group, label: StaticText, input: EditText}} 生成した行と部品
     */
    function addFieldRow(parent, labelString, initialText, labelWidth, unitLabel) {
        var row = parent.add("group");
        setupRow(row);
        var label = addFieldLabel(row, labelString, labelWidth);
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
            justify: { ja: "行揃え", en: "Justification" },
            markerColor: { ja: "記号／番号", en: "Marker/Number" },
            startNumber: { ja: "開始番号", en: "Start No." }
        },
        radio: {
            bullet: { ja: "箇条書き", en: "Bullets" },
            numbered: { ja: "番号リスト", en: "Numbered List" },
            none: { ja: "なし", en: "None" }
        },
        justify: {
            left: { ja: "左揃え", en: "Left" },
            center: { ja: "中央揃え", en: "Center" },
            right: { ja: "右揃え", en: "Right" },
            justifyLeft: { ja: "均等配置（最終行左揃え）", en: "Justify (last line left)" },
            justifyAll: { ja: "両端揃え", en: "Justify all" }
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
            typeNone: {
                ja: "行頭のマーカーを削除します。タブストップも解除します。",
                en: "Removes the leading markers, and clears the tab stops as well."
            },
            bulletStyle: {
                ja: "記号ごとに拡大率と本文位置の既定値が変わります。",
                en: "Each glyph brings its own default scale and body position."
            },
            circledNumber: {
                ja: "①〜⑳の範囲外は素の数字になります。",
                en: "Numbers outside 1-20 fall back to plain digits."
            },
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
            baselineShift: {
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
            justify: {
                ja: "段落の行揃えです。既定はテキストの種類に合わせて、ポイント文字＝左揃え、エリア内文字＝均等配置（最終行左揃え）になります。",
                en: "Paragraph justification. The default follows the text kind: left for point text, justified with the last line left for area text."
            },
            leading: {
                ja: "行送り（空欄のままなら変更しません）。固定値ではなく、文字サイズに対する割合として自動行送りに設定します。",
                en: "Line leading (left blank = leave unchanged). Applied as an auto-leading amount (% of the font size), not as a fixed value."
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
                ja: "比率・ベースライン・カラー・インデント・タブストップをまとめてリセットします。行の並べ替えは元に戻りません。",
                en: "Resets scale, baseline shift, color, indent, and tab stops at once. Line reordering is not undone."
            },
            startNumber: {
                ja: "連番の開始番号です。複数フレームを選択したときは、先頭のフレームから通しで番号を振ります。",
                en: "The number to start counting from. With multiple frames selected, numbering runs continuously from the first frame."
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
    // セッション状態 / Session state
    // =========================================
    // Illustratorの起動中だけダイアログの値を保持する（ファイルには保存しない）
    // Dialog values are kept only while Illustrator is running (nothing is written to disk)

    var ENGINE_STATE_KEY = "__AddBulletsAndNumbers__";
    $.global[ENGINE_STATE_KEY] = $.global[ENGINE_STATE_KEY] || {};

    /**
     * 前回のダイアログ設定を読み込む
     * @returns {object} 保存されている設定（初回は空のオブジェクト）
     */
    function loadSessionState() {
        return $.global[ENGINE_STATE_KEY] || {};
    }

    /**
     * 現在のダイアログ設定を保存する
     * @param {object} state - 保存する設定
     * @returns {void}
     */
    function saveSessionState(state) {
        $.global[ENGINE_STATE_KEY] = state || {};
    }

    // =========================================
    // テキスト行の共通処理 / Line helpers
    // =========================================

    /**
     * 選択の中で最初に見つかったテキストフレームを返す
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {TextFrame|null} 先頭のテキストフレーム（なければ null）
     */
    function firstTextFrame(selection) {
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") return selection[i];
        }
        return null;
    }

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
     * @param {Array<TextFrame>} collected - 収集先の配列
     * @returns {Array<TextFrame>} 収集したテキストフレーム
     */
    function collectTextFrames(items, collected) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!item) continue;
            if (item.typename === "TextFrame") {
                collected.push(item);
            } else if (item.typename === "GroupItem") {
                // グループ内のテキストフレームも対象 / include text frames inside groups
                collectTextFrames(item.pageItems, collected);
            }
        }
        return collected;
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
            // 片方の読み取り失敗でもう片方まで0にすると比較が非対称になるため、オブジェクトごとに囲む
            // Guard each object separately: zeroing the other one too would make the comparator inconsistent
            var aTop = 0, aLeft = 0, bTop = 0, bLeft = 0;
            try { aTop = a.top; aLeft = a.left; } catch (eA) { }
            try { bTop = b.top; bLeft = b.left; } catch (eB) { }
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
    function makeDocumentBlack() {
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
        var frame = firstTextFrame(selection);
        if (frame) {
            try {
                var fillColor = frame.textRange.characterAttributes.fillColor;
                if (fillColor && fillColor.typename && fillColor.typename !== "NoColor") return fillColor;
            } catch (e) { }
        }
        return makeDocumentBlack();
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
    // 行揃えボタンの描画 / Justification button icons
    // =========================================
    // UnifiedTypePanel.jsx の実装を流用 / Ported from UnifiedTypePanel.jsx

    /**
     * IllustratorのUIが明るいテーマかどうかを判定する
     * @returns {boolean} 明るいテーマなら true
     */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /**
     * テーマと選択状態に応じたボタンの配色を返す
     * @param {boolean} isLight - 明るいテーマかどうか
     * @param {boolean} isActive - 選択中かどうか
     * @returns {{bg: number[], border: (number[]|null), line: number[]}} 背景・枠線・罫線の色
     */
    function getJustifyColors(isLight, isActive) {
        if (isLight) {
            return {
                bg: isActive ? [0.40, 0.40, 0.40, 1] : [1, 1, 1, 1],
                border: isActive ? [0.30, 0.30, 0.30, 1] : [0.62, 0.62, 0.62, 1],
                line: isActive ? [1, 1, 1, 1] : [0.25, 0.25, 0.25, 1]
            };
        }
        return {
            bg: isActive ? [0.92, 0.92, 0.92, 1] : [0.30, 0.30, 0.30, 1],
            border: null,
            line: isActive ? [0.16, 0.16, 0.16, 1] : [0.82, 0.82, 0.82, 1]
        };
    }

    /**
     * アイコンの各行の長さを返す
     * @param {string} justifyId - 行揃えのID
     * @param {number} longWidth - 長い行の幅
     * @param {number} shortWidth - 短い行の幅
     * @returns {number[]} 上から4行ぶんの幅
     */
    function getJustifyLineWidths(justifyId, longWidth, shortWidth) {
        if (justifyId === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
        if (justifyId === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
        return [longWidth, shortWidth, longWidth, shortWidth];
    }

    /**
     * アイコンの各行の開始位置を返す
     * @param {string} justifyId - 行揃えのID
     * @param {number} buttonWidth - ボタンの幅
     * @param {number} lineWidth - その行の幅
     * @returns {number} 行の開始X座標
     */
    function getJustifyLineX(justifyId, buttonWidth, lineWidth) {
        var margin = 5;
        if (justifyId === "right") return buttonWidth - margin - lineWidth;
        if (justifyId === "center") return Math.round((buttonWidth - lineWidth) / 2);
        return margin;
    }

    /**
     * 行揃えボタンの背景とアイコンを描く
     * @param {Button} button - 対象のボタン（justifyId を持つ）
     * @param {boolean} isActive - 選択中かどうか
     * @param {boolean} isLight - 明るいテーマかどうか
     * @returns {void}
     */
    function drawJustifyIcon(button, isActive, isLight) {
        var graphics = button.graphics;
        var buttonWidth = button.size[0];
        var buttonHeight = button.size[1];
        var colors = getJustifyColors(isLight, isActive);

        try {
            graphics.rectPath(0, 0, buttonWidth, buttonHeight);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.bg));
            if (colors.border) {
                graphics.rectPath(0, 0, buttonWidth, buttonHeight);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, colors.border, 1));
            }
        } catch (eFill) {
            try { graphics.drawOSControl(); } catch (eOS) { }
        }

        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, colors.line, 1.2);
        var rowYs = [7, 11, 15, 19];
        var lineWidths = getJustifyLineWidths(button.justifyId, 15, 10);
        for (var i = 0; i < rowYs.length; i++) {
            var lineStartX = getJustifyLineX(button.justifyId, buttonWidth, lineWidths[i]);
            graphics.newPath();
            graphics.moveTo(lineStartX, rowYs[i]);
            graphics.lineTo(lineStartX + lineWidths[i], rowYs[i]);
            graphics.strokePath(pen);
        }
    }

    /**
     * 行揃えのIDを Justification へ変換する
     * @param {string} justifyId - 行揃えのID
     * @returns {Justification} 対応する Justification
     */
    function resolveJustification(justifyId) {
        if (justifyId === "center") return Justification.CENTER;
        if (justifyId === "right") return Justification.RIGHT;
        if (justifyId === "justifyLeft") return Justification.FULLJUSTIFYLASTLINELEFT;
        if (justifyId === "justifyAll") return Justification.FULLJUSTIFY;
        return Justification.LEFT;
    }

    /**
     * テキストの種類に応じた既定の行揃えIDを求める
     * ポイント文字は左揃え、エリア内文字は均等配置（最終行左揃え）
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {string} 行揃えのID
     */
    function defaultJustifyId(selection) {
        var frame = firstTextFrame(selection);
        try {
            if (frame && frame.kind === TextType.AREATEXT) return "justifyLeft";
        } catch (e) { }
        return "left";
    }

    // =========================================
    // 現在状態の推定 / Detect current state
    // =========================================

    /**
     * 選択先頭のテキストフレームから最初の非空行を取得する
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {string|null} 最初の非空行（見つからなければ null）
     */
    function firstNonEmptyLine(selection) {
        var frame = firstTextFrame(selection);
        if (!frame) return null;
        var lines = splitIntoLines(frame.contents);
        for (var i = 0; i < lines.length; i++) {
            if (/\S/.test(lines[i])) return lines[i];
        }
        return null;
    }

    /**
     * 区切り文字の DELIMITERS 内でのインデックスを取得する
     * @param {string} delimiterChar - 区切り文字
     * @returns {number} インデックス（未一致は 0＝なし）
     */
    function findDelimiterIndex(delimiterChar) {
        for (var i = 0; i < DELIMITERS.length; i++) {
            if (DELIMITERS[i] === delimiterChar) return i;
        }
        return 0;
    }

    /**
     * 現在の行頭マーカーからダイアログの初期状態を推定する（このスクリプトが付与した形式を想定）
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {{type: string, bulletIndex: number, numberStyle: string, delimiterIndex: number}} 推定した初期状態
     */
    function detectCurrentListState(selection) {
        var state = { type: "bullet", bulletIndex: 0, numberStyle: "number", delimiterIndex: 1 };
        var line = firstNonEmptyLine(selection);
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
            state.delimiterIndex = findDelimiterIndex(numberedMatch[2]);
            return state;
        }

        // 箇条書き記号 + タブ / bullet glyph + tab
        if (line.charAt(1) === "\t") {
            var leadingGlyph = line.charAt(0);
            for (var i = 0; i < BULLET_SYMBOLS.length; i++) {
                if (leadingGlyph === BULLET_SYMBOLS[i].mark) {
                    state.type = "bullet";
                    state.bulletIndex = i;
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
        var frameSnapshots = captureFrameSnapshots(targetSelection);

        // リセット済みか（ハンギングOFFでもインデントを0として適用する）/ Whether Reset was pressed (apply zero indents even with hanging off)
        var indentsCleared = false;

        // セッション設定の復元中か（値の代入で走るハンドラーに邪魔をさせない）
        // Whether a saved session is being restored (assigning values fires handlers that would interfere)
        var restoringSession = false;

        // プレビューの遅延実行 / Deferred preview
        // 入力のたびに再適用すると画面がちらつくため、入力が落ち着くまでまとめる
        // Reapplying on every keystroke makes the canvas flicker, so updates are coalesced until input settles
        var PREVIEW_TASK_KEY = "__addBulletsAndNumbersPreview";
        var PREVIEW_TASK_CALL = "if ($.global." + PREVIEW_TASK_KEY + ") $.global." + PREVIEW_TASK_KEY + "();";
        var previewTaskId = null;

        // 予約したタスクはグローバルスコープで評価されるため、呼び出し口を $.global に置く
        // A scheduled task is evaluated in the global scope, so the entry point lives on $.global
        $.global[PREVIEW_TASK_KEY] = function () {
            previewTaskId = null;
            updatePreview();
        };

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
                for (var symbolIndex = 0; symbolIndex < bulletRadios.length; symbolIndex++) {
                    if (bulletRadios[symbolIndex].value) return BULLET_SYMBOLS[symbolIndex];
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

        var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 上段: 2カラム（左: 設定 / 右: テキストプレビュー）/ Top: two columns (left: settings / right: text preview) */
        var topRow = dialog.add("group");
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
        var markerStyleRow = topLeftCol.add("group");
        markerStyleRow.orientation = "row";
        markerStyleRow.alignChildren = ["fill", "top"];
        markerStyleRow.spacing = COLUMN_SPACING;

        /* 箇条書き記号パネル（箇条書きのみ有効）/ Bullet symbol panel (bullet only) */
        var bulletStylePanel = addPanel(markerStyleRow, getLabel('panel.bulletStyle'), DENSE_SPACING);

        // 記号候補をラジオで列挙 / List the symbol candidates as radios
        var bulletRadios = [];
        for (var symbolIndex = 0; symbolIndex < BULLET_SYMBOLS.length; symbolIndex++) {
            bulletRadios.push(bulletStylePanel.add("radiobutton", undefined, BULLET_SYMBOLS[symbolIndex].mark));
        }
        bulletRadios[0].value = true;

        /* 番号スタイル＋区切り文字を縦に並べるカラム / Column stacking number style + delimiter vertically */
        var numberStyleColumn = markerStyleRow.add("group");
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
        var numberStyleIds = ["number", "circledWhite", "circledBlack", "upperAlpha", "lowerAlpha"]; // 上の並びと対応 / matches the order above

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
        setupRow(tabStopColumns, "left", WIDE_COLUMN_SPACING); // 1つ目／2つ目カラムのガターを広めに / wider gutter between the 1st / 2nd columns
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
        var tabAlignRadios = [rbAlignLeft, rbAlignCenter, rbAlignRight];
        var tabAlignIds = ["left", "center", "right"]; // 上の並びと対応 / matches the order above

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
        var reorderPanel = addPanel(topRightCol, getLabel('panel.sortElements'), 0);

        // 左=行番号 / 右=本文（マーカーは含めない）/ Left = row number, Right = body text (markers excluded)
        // 列幅の合計をリストボックス幅に合わせて2列目を端まで使う / column widths sum to the listbox width so column 2 fills to the edge
        var previewList = reorderPanel.add("listbox", undefined, [], {
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
        var reorderRow = reorderPanel.add("group");
        setupRow(reorderRow, "fill", 4);
        reorderRow.margins = [0, 10, 0, 0]; // listbox との間に上マージン / top margin from the listbox

        var btnMoveTop = addReorderButton(reorderRow, getLabel('reorder.top'), 52);
        var btnMoveUp = addReorderButton(reorderRow, getLabel('reorder.up'), 46);
        var btnMoveDown = addReorderButton(reorderRow, getLabel('reorder.down'), 46);
        var btnMoveBottom = addReorderButton(reorderRow, getLabel('reorder.bottom'), 52);

        // 名前順・数値順は次の行に / "By name" / "By number" go on the next row
        var sortButtonRow = reorderPanel.add("group");
        setupRow(sortButtonRow, "fill", 4);
        sortButtonRow.margins = [0, 5, 0, 0]; // 上にマージン / top margin

        var btnSortByName = addReorderButton(sortButtonRow, getLabel('reorder.sortByName'), 70);
        var btnSortByValue = addReorderButton(sortButtonRow, getLabel('reorder.sortByValue'), 70);
        var btnSortByValueDesc = addReorderButton(sortButtonRow, getLabel('reorder.sortByValueDesc'), 100);

        /* ===== 2行目: 左=記号や番号の書式 / 右=段落の書式 / Row 2: left = symbol/number format, right = paragraph format ===== */
        var bottomRow = dialog.add("group");
        bottomRow.orientation = "row";
        bottomRow.alignChildren = ["fill", "top"];
        bottomRow.spacing = COLUMN_SPACING;

        /* 下段左カラム / Bottom-left column */
        var bottomLeftCol = bottomRow.add("group");
        setupColumn(bottomLeftCol);

        /* マーカーの書式パネル / Marker format panel */
        var formatPanel = addPanel(bottomLeftCol, getLabel('panel.format'), DENSE_SPACING);

        // フォントファミリー（ポップアップ）/ Font family (popup)
        var baseFontName = getBaseFontName(targetSelection);
        var fontRow = formatPanel.add("group");
        setupRow(fontRow);
        var fontLabel = addFieldLabel(fontRow, getLabelWithColon('fieldLabel.font'), FIELD_LABEL_WIDTH);
        var fontFamilyDropdown = fontRow.add("dropdownlist", undefined, []);
        fontFamilyDropdown.preferredSize.width = 200;

        // フォントスタイル（ポップアップ）/ Font style (popup)
        var fontStyleRow = formatPanel.add("group");
        setupRow(fontStyleRow);
        var fontStyleLabel = addFieldLabel(fontStyleRow, getLabelWithColon('fieldLabel.fontStyle'), FIELD_LABEL_WIDTH);
        var fontStyleDropdown = fontStyleRow.add("dropdownlist", undefined, []);
        fontStyleDropdown.preferredSize.width = 200;

        // 和文フォントのみ / Japanese fonts only
        var jpOnlyRow = formatPanel.add("group");
        setupRow(jpOnlyRow);
        jpOnlyRow.add("statictext", undefined, "").preferredSize.width = FIELD_LABEL_WIDTH; // ラベル列に合わせるスペーサー / spacer to align with the label column
        var chkJPOnly = jpOnlyRow.add("checkbox", undefined, getLabel('checkbox.jpOnly'));

        populateFontFamilyDropdown(fontFamilyDropdown, baseFontName, chkJPOnly.value);
        populateFontStyleDropdown(fontStyleDropdown, fontFamilyDropdown.selection, baseFontName);

        // 2カラム: 左=比率・ベースライン / 右=カラー / Two columns: left = scale & baseline shift, right = color
        var formatColumns = formatPanel.add("group");
        setupRow(formatColumns, "left", WIDE_COLUMN_SPACING);
        formatColumns.alignChildren = ["left", "top"];

        // 左カラム / Left column: scale & baseline shift
        var formatLeftCol = formatColumns.add("group");
        setupColumn(formatLeftCol, DENSE_SPACING);
        formatLeftCol.alignChildren = ["left", "top"];
        formatLeftCol.alignment = ["left", "top"];

        // 比率（水平・垂直を同率で適用, %）。初期は箇条書きの記号に合わせた比率 / Scale (applied equally to H and V, %), initialized for the default bullet glyph
        var scaleUI = addFieldRow(formatLeftCol, getLabelWithColon('fieldLabel.scale'), "120", FIELD_LABEL_WIDTH, "%");
        var scaleLabel = scaleUI.label;
        var inputScale = scaleUI.input;

        // ベースラインシフト / Baseline shift
        var baselineUI = addFieldRow(formatLeftCol, getLabelWithColon('fieldLabel.baselineShift'), "0", FIELD_LABEL_WIDTH, unitInfo.label);
        var baselineLabel = baselineUI.label;
        var inputBaseline = baselineUI.input;

        // 右カラム / Right column: color
        var formatRightCol = formatColumns.add("group");
        setupColumn(formatRightCol, DENSE_SPACING);
        formatRightCol.alignChildren = ["left", "top"];
        formatRightCol.alignment = ["left", "top"];

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
            var label = addFieldLabel(row, labelString, colorLabelWidth);
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

        var markerColorUI = addColorRow(formatRightCol, getLabelWithColon('fieldLabel.markerColor'), baseFillColor);
        var delimiterColorUI = addColorRow(formatRightCol, getLabelWithColon('delimiter.label'), baseFillColor);
        var colorSwatch = markerColorUI.swatch;             // 記号／番号のカラー / marker-number color
        var delimiterColorSwatch = delimiterColorUI.swatch; // 区切り文字のカラー / delimiter color

        /* 下段右カラム / Bottom-right column */
        var bottomRightCol = bottomRow.add("group");
        setupColumn(bottomRightCol);

        /* 段落設定パネル / Paragraph settings panel */
        var paragraphPanel = addPanel(bottomRightCol, getLabel('panel.paragraph'), DENSE_SPACING);

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

        // 行揃え（アイコンボタン）。既定はテキストの種類から決める / Justification (icon buttons); the default comes from the text kind
        var justifyRow = paragraphPanel.add("group");
        setupRow(justifyRow, "left", DENSE_SPACING);
        var justifyLabel = addFieldLabel(justifyRow, getLabelWithColon('fieldLabel.justify'), PARAGRAPH_LABEL_WIDTH);

        var justifyButtonRow = justifyRow.add("group");
        setupRow(justifyButtonRow, "left", 2);

        // 選択中のIDとUIの明暗を共有する（onDraw のクロージャから参照）/ Shared active id + theme, read by the onDraw closures
        var justifyState = { activeId: defaultJustifyId(targetSelection), isLight: isLightUI() };
        var justifyButtons = [];
        for (var justifyIndex = 0; justifyIndex < JUSTIFY_OPTIONS.length; justifyIndex++) {
            var justifyButton = justifyButtonRow.add("button", undefined, "");
            justifyButton.justifyId = JUSTIFY_OPTIONS[justifyIndex].id;
            justifyButton.helpTip = getLabel(JUSTIFY_OPTIONS[justifyIndex].labelKey);
            justifyButton.preferredSize = [JUSTIFY_BUTTON_SIZE, JUSTIFY_BUTTON_SIZE];
            justifyButton.minimumSize = [JUSTIFY_BUTTON_SIZE, JUSTIFY_BUTTON_SIZE];
            justifyButton.maximumSize = [JUSTIFY_BUTTON_SIZE, JUSTIFY_BUTTON_SIZE];
            justifyButton.onDraw = function () { drawJustifyIcon(this, this.justifyId === justifyState.activeId, justifyState.isLight); };
            justifyButtons.push(justifyButton);
        }

        // ハンギング対応（チェック時のみ左/1行目インデントを適用）/ Hanging indent (apply left/first-line indent only when checked)
        // チェックボックスは他のラベルと同じ左位置に / checkbox aligned at the label column like the other rows
        var hangingRow = paragraphPanel.add("group");
        setupRow(hangingRow, "left", DENSE_SPACING);
        hangingRow.margins = [0, 12, 0, 0]; // 上に余白 / top margin
        var chkHanging = hangingRow.add("checkbox", undefined, getLabel('checkbox.hanging'));
        chkHanging.value = true; // 折り返し位置をそろえた状態を既定にする / aligned wrapped lines by default

        // 左インデント（1行目インデントはUIを持たず、内部で左インデントの正負を反転して適用）
        // Left indent (first-line indent has no UI; it's the negated left indent, computed internally)
        var leftIndentUI = addFieldRow(paragraphPanel, getLabelWithColon('fieldLabel.leftIndent'), "0", PARAGRAPH_LABEL_WIDTH, unitInfo.label);
        var leftIndentLabel = leftIndentUI.label;
        var inputLeftIndent = leftIndentUI.input;

        /* ボタンエリア（左: 制御文字表示 / 中央: spacer / 右: Cancel → OK）/ Button area (left: show hidden / center: spacer / right: Cancel → OK) */
        var buttonRow = dialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = "fill";

        // 左カラム / Left column
        var buttonLeftGroup = buttonRow.add("group");
        var btnShowHidden = buttonLeftGroup.add("button", undefined, getLabel('button.showHidden'));
        var btnReset = buttonLeftGroup.add("button", undefined, getLabel('button.reset'));

        // 中央カラム（伸縮スペーサー）/ Center column (flexible spacer)
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;

        // 右カラム / Right column
        var buttonRightGroup = buttonRow.add("group");
        buttonRightGroup.alignChildren = ["right", "center"];
        var btnCancel = buttonRightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnOK = buttonRightGroup.add("button", undefined, "OK", { name: "ok" });

        /* 制御文字（隠し文字）の表示切り替え / Toggle hidden characters */
        btnShowHidden.onClick = function () {
            app.executeMenuCommand('showHiddenChar');
        };

        /* まとめてリセット / Reset everything listed */
        btnReset.onClick = function () { resetAllValues(); };

        /**
         * ラジオボタンの配列から、選択されているものの位置を返す
         * @param {Array<RadioButton>} radios - 対象のラジオボタン
         * @returns {number} 選択位置（未選択なら0）
         */
        function selectedIndexOf(radios) {
            for (var i = 0; i < radios.length; i++) {
                if (radios[i].value) return i;
            }
            return 0;
        }

        /**
         * ラジオボタンの配列から、指定位置だけを選択状態にする
         * @param {Array<RadioButton>} radios - 対象のラジオボタン
         * @param {number} index - 選択する位置
         * @returns {void}
         */
        function selectRadioAt(radios, index) {
            if (index == null || index < 0 || index >= radios.length) return;
            for (var i = 0; i < radios.length; i++) radios[i].value = (i === index);
        }

        /**
         * 配列の中から値の位置を探す
         * @param {Array} list - 探索対象の配列
         * @param {*} value - 探す値
         * @returns {number} 見つかった位置（無ければ -1）
         */
        function indexOfValue(list, value) {
            for (var i = 0; i < list.length; i++) {
                if (list[i] === value) return i;
            }
            return -1;
        }

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
            for (var i = 0; i < delimiterRadios.length; i++) {
                if (delimiterRadios[i].value) return DELIMITERS[i];
            }
            return "";
        }

        /**
         * 現在選択されている番号スタイルを取得する
         * @returns {string} "number" / "circledWhite" / "circledBlack" / "upperAlpha" / "lowerAlpha"
         */
        function currentStyle() {
            return numberStyleIds[selectedIndexOf(numberStyleRadios)];
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

            var baselineShift = parseFloat(inputBaseline.text);
            if (isNaN(baselineShift)) baselineShift = 0;

            return {
                fontName: fontName,
                horizontalScale: scale,
                verticalScale: scale,
                baselineShiftPt: baselineShift * unitInfo.factor,
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
            // リセット後はハンギングOFFでも0を適用してインデントを消す
            // Indents apply only when hanging is on; first-line indent = negated left indent (no UI, computed internally).
            // After Reset, zero is applied even with hanging off so the indents are actually cleared.
            var leftIndentPt = chkHanging.value ? readOptionalPt(inputLeftIndent) : (indentsCleared ? 0 : null);
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
         * 「折り返し位置を揃える」の既定インデントを入力欄へ入れる
         * 箇条書きは文字サイズ、番号リストは文字サイズの150%
         * @returns {void}
         */
        function applyDefaultHangingIndent() {
            var indentPt = rbNumbered.value ? (baseFontSize * 1.5) : baseFontSize;
            inputLeftIndent.text = String(toDisplayUnit(indentPt));
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
        /**
         * 本文位置のタブストップをポイントで取得する
         * @returns {number} 本文位置（pt）
         */
        function currentTabStop2Pt() {
            return readTabStop(inputTabStop2, defaultStopsForType(currentStopType()).stop2Pt) * unitInfo.factor;
        }

        /**
         * マーカー位置のタブストップの揃え種類を取得する
         * @returns {TabStopAlignment} 揃え種類
         */
        function currentTab1Alignment() {
            var alignId = tabAlignIds[selectedIndexOf(tabAlignRadios)];
            if (alignId === "left") return TabStopAlignment.Left;
            if (alignId === "center") return TabStopAlignment.Center;
            return TabStopAlignment.Right;
        }

        /**
         * 種類に応じた既定のタブストップ値を入力欄へ反映する
         * @returns {void}
         */
        function applyTypeDefaults() {
            if (rbNone.value) return; // 「なし」では位置調整は未使用 / Position is unused for "none"
            var stops = defaultStopsForType(currentStopType());
            inputTabStop1.text = String(toDisplayUnit(stops.stop1Pt));
            inputTabStop2.text = String(toDisplayUnit(stops.stop2Pt));
        }

        /**
         * 推定した現在状態をUIへ反映する（「現状の続き」として編集できるように）
         * @param {{type: string, bulletIndex: number, numberStyle: string, delimiterIndex: number}} detectedState - detectCurrentListState() の結果
         * @returns {void}
         */
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
            for (var i = 0; i < delimiterRadios.length; i++) delimiterRadios[i].value = (i === detectedState.delimiterIndex);
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

        /**
         * 現在のダイアログ設定を、セッション保存用のオブジェクトにまとめる
         * @returns {object} 保存する設定
         */
        function collectSessionState() {
            return {
                listType: currentType(),
                bulletIndex: selectedIndexOf(bulletRadios),
                numberStyleIndex: selectedIndexOf(numberStyleRadios),
                delimiterIndex: selectedIndexOf(delimiterRadios),
                tabAlignIndex: selectedIndexOf(tabAlignRadios),
                startNumber: inputStartNumber.text,
                zeroPad: chkZeroPad.value,
                resetPerFrame: chkResetPerFrame.value,
                tabStop1: inputTabStop1.text,
                tabStop2: inputTabStop2.text,
                fontName: getSelectedFontName(fontFamilyDropdown, fontStyleDropdown),
                japaneseOnly: chkJPOnly.value,
                scale: inputScale.text,
                baselineShift: inputBaseline.text,
                // カラーは文字列で持つ（ドキュメントが変わってもそのまま復元できる）/ keep colors as strings so they survive a document change
                markerColor: aiColorToPickerString(colorSwatch._aiColor),
                delimiterColor: aiColorToPickerString(delimiterColorSwatch._aiColor),
                leading: inputLeading.text,
                spaceAfter: inputSpaceAfter.text,
                justifyId: justifyState.activeId,
                hanging: chkHanging.value,
                leftIndent: inputLeftIndent.text
            };
        }

        /**
         * 保存しておいた設定をダイアログへ反映する（同じセッションの2回目以降）
         * @param {object} state - loadSessionState() が返した設定
         * @returns {void}
         */
        function applySessionState(state) {
            if (!state || !state.listType) return; // 初回は推定した状態のまま / first run keeps the detected state

            restoringSession = true;
            rbBullet.value = (state.listType === "bullet");
            rbNumbered.value = (state.listType === "numbered");
            rbNone.value = (state.listType === "none");
            selectRadioAt(bulletRadios, state.bulletIndex);
            selectRadioAt(numberStyleRadios, state.numberStyleIndex);
            selectRadioAt(delimiterRadios, state.delimiterIndex);
            selectRadioAt(tabAlignRadios, state.tabAlignIndex);

            if (state.startNumber != null) inputStartNumber.text = state.startNumber;
            if (state.zeroPad != null) chkZeroPad.value = state.zeroPad;
            if (state.resetPerFrame != null) chkResetPerFrame.value = state.resetPerFrame;
            if (state.tabStop1 != null) inputTabStop1.text = state.tabStop1;
            if (state.tabStop2 != null) inputTabStop2.text = state.tabStop2;
            if (state.scale != null) inputScale.text = state.scale;
            if (state.baselineShift != null) inputBaseline.text = state.baselineShift;
            if (state.leading != null) inputLeading.text = state.leading;
            if (state.spaceAfter != null) inputSpaceAfter.text = state.spaceAfter;
            if (state.hanging != null) chkHanging.value = state.hanging;
            if (state.leftIndent != null) inputLeftIndent.text = state.leftIndent;
            if (state.justifyId) selectJustify(state.justifyId);

            // フォントは一覧を作り直してから選択し直す
            // ファミリーの選択で onChange が走りスタイルが初期化されるため、スタイルの復元は必ずそのあとに行う
            // Rebuild the dropdowns, then reselect: choosing a family fires onChange and resets the style,
            // so the style has to be restored afterwards
            if (state.japaneseOnly != null) chkJPOnly.value = state.japaneseOnly;
            if (state.fontName) {
                populateFontFamilyDropdown(fontFamilyDropdown, state.fontName, chkJPOnly.value);
                populateFontStyleDropdown(fontStyleDropdown, fontFamilyDropdown.selection, state.fontName);
            }

            // カラーは表示前なので _aiColor を差し替えるだけでよい / the dialog is not shown yet, so swapping _aiColor is enough
            if (state.markerColor) colorSwatch._aiColor = pickerStringToAiColor(state.markerColor);
            if (state.delimiterColor) delimiterColorSwatch._aiColor = pickerStringToAiColor(state.delimiterColor);

            restoringSession = false;
        }

        /**
         * 行一覧を更新する（左=行番号、右=マーカーを除いた本文）
         * @returns {void}
         */
        function refreshList() {
            // removeAll で選択も消えるため、先に選択行を退避して再構築後に復元
            // removeAll() drops the selection too, so capture it first and restore after rebuilding
            // （並べ替え時は applyReorder が後から正しい選択で上書きするので競合しない / on reorder, applyReorder overrides this afterward)
            var previousSelection = selectedRowIndices();

            previewList.removeAll();
            var bodyEntries = getFlatBodyEntries(frameSnapshots);
            for (var k = 0; k < bodyEntries.length; k++) {
                var row = previewList.add("item", String(k + 1)); // 左カラム: 連番 / Left column: sequential number
                row.subItems[0].text = bodyEntries[k].text;        // 右カラム: 本文 / Right column: body text
            }

            // 退避した選択を復元（行数が減っている場合は範囲内のみ）/ Restore selection (only indices still in range)
            selectRows(previousSelection);
            updateMoveButtonEnabled();
        }

        /**
         * 選択中の行インデックスを昇順の配列で取得する（複数選択対応）
         * @returns {number[]} 選択行のインデックス（未選択は空配列）
         */
        function selectedRowIndices() {
            var selection = previewList.selection;
            if (!selection) return [];
            // multiselect では配列、単一選択でも item が返るため両対応 / multiselect returns an array; a single item is also handled
            var items = (selection instanceof Array) ? selection : [selection];
            var indices = [];
            for (var i = 0; i < items.length; i++) {
                if (items[i]) indices.push(items[i].index);
            }
            indices.sort(function (a, b) { return a - b; });
            return indices;
        }

        /**
         * 指定した行を選択状態にする（範囲外のインデックスは捨てる）
         * @param {number[]} indices - 選択したい行のインデックス
         * @returns {void}
         */
        function selectRows(indices) {
            if (!indices || !indices.length) return;
            var valid = [];
            for (var i = 0; i < indices.length; i++) {
                if (indices[i] >= 0 && indices[i] < previewList.items.length) valid.push(indices[i]);
            }
            if (valid.length) previewList.selection = valid; // 配列で複数選択を復元 / restore multi-selection via an array
        }

        /**
         * 行揃えボタンの選択状態を切り替えて描き直す
         * @param {string} justifyId - 選択する行揃えのID
         * @returns {void}
         */
        function selectJustify(justifyId) {
            justifyState.activeId = justifyId;
            for (var i = 0; i < justifyButtons.length; i++) {
                try { justifyButtons[i].notify("onDraw"); } catch (e) { }
            }
        }

        /**
         * 選択の有無に応じて移動ボタンの有効／無効を切り替える
         * @returns {void}
         */
        function updateMoveButtonEnabled() {
            var hasSelection = selectedRowIndices().length > 0;
            btnMoveTop.enabled = hasSelection;
            btnMoveUp.enabled = hasSelection;
            btnMoveDown.enabled = hasSelection;
            btnMoveBottom.enabled = hasSelection;
        }

        /**
         * 並べ替え結果を控えへ書き戻し、プレビューと選択を更新する
         * @param {Array<{text: string, attrs: Array<object>}>} flatEntries - 並べ替え後の行データ
         * @param {number[]} newIndices - 並べ替え後に選択しておく行インデックス
         * @returns {void}
         */
        function applyReorder(flatEntries, newIndices) {
            writeFlatBodyEntries(frameSnapshots, flatEntries);
            refreshList();   // 一覧の内容が変わるのは並べ替えのときだけ / the list contents change on reorder only
            updatePreview();
            selectRows(newIndices);
            updateMoveButtonEnabled();
        }

        /**
         * 選択行を1行ぶん上下へ移動する（相対順を保ち、端や他の選択行で詰まる分は動かさない）
         * @param {number} delta - -1 で上へ、+1 で下へ
         * @returns {void}
         */
        function moveSelectedRows(delta) {
            var indices = selectedRowIndices();
            if (!indices.length) return;
            var flatEntries = getFlatBodyEntries(frameSnapshots);
            var rowCount = flatEntries.length;
            var selected = {};
            for (var i = 0; i < indices.length; i++) selected[indices[i]] = true;

            if (delta < 0) {
                // 上へ: 上から順に、直上が未選択なら入れ替え / up: top-down, swap with the slot above when it's free
                for (var rowIndex = 1; rowIndex < rowCount; rowIndex++) {
                    if (selected[rowIndex] && !selected[rowIndex - 1]) {
                        var swapped = flatEntries[rowIndex]; flatEntries[rowIndex] = flatEntries[rowIndex - 1]; flatEntries[rowIndex - 1] = swapped;
                        delete selected[rowIndex]; selected[rowIndex - 1] = true;
                    }
                }
            } else if (delta > 0) {
                // 下へ: 下から順に、直下が未選択なら入れ替え / down: bottom-up, swap with the slot below when it's free
                for (var rowIndex = rowCount - 2; rowIndex >= 0; rowIndex--) {
                    if (selected[rowIndex] && !selected[rowIndex + 1]) {
                        var swapped = flatEntries[rowIndex]; flatEntries[rowIndex] = flatEntries[rowIndex + 1]; flatEntries[rowIndex + 1] = swapped;
                        delete selected[rowIndex]; selected[rowIndex + 1] = true;
                    }
                }
            }

            var newIndices = [];
            for (var k = 0; k < rowCount; k++) { if (selected[k]) newIndices.push(k); }
            applyReorder(flatEntries, newIndices);
        }

        /**
         * 選択行を先頭または末尾へ移動する（選択行どうしの元の並び順は維持）
         * @param {string} position - "top" または "bottom"
         * @returns {void}
         */
        function moveSelectedRowsTo(position) {
            var indices = selectedRowIndices();
            if (!indices.length) return;
            var flatEntries = getFlatBodyEntries(frameSnapshots);
            // 後ろから抜き出してインデックスのズレを防ぎ、元の順序で moved に並べる / splice from the end to avoid index drift; keep original order in `moved`
            var movedEntries = [];
            for (var i = indices.length - 1; i >= 0; i--) {
                movedEntries.unshift(flatEntries.splice(indices[i], 1)[0]);
            }
            var newIndices = [];
            if (position === "top") {
                for (var j = 0; j < movedEntries.length; j++) { flatEntries.splice(j, 0, movedEntries[j]); newIndices.push(j); }
            } else {
                var tailStart = flatEntries.length;
                for (var k = 0; k < movedEntries.length; k++) { flatEntries.push(movedEntries[k]); newIndices.push(tailStart + k); }
            }
            applyReorder(flatEntries, newIndices);
        }

        /**
         * 全行を名前順（昇順）に並べ替える
         * @returns {void}
         */
        function sortRowsByName() {
            var flatEntries = getFlatBodyEntries(frameSnapshots);
            flatEntries.sort(function (a, b) { return (a.text < b.text) ? -1 : (a.text > b.text ? 1 : 0); });
            applyReorder(flatEntries, []);
        }

        /**
         * 本文中の最初の数値を取り出す（小数・マイナス可）
         * @param {string} text - 対象の本文
         * @returns {number|null} 見つかった数値（なければ null）
         */
        function extractFirstNumber(text) {
            var match = String(text).match(/-?\d+(?:\.\d+)?/);
            return match ? parseFloat(match[0]) : null;
        }

        /**
         * 全行を項目内の数値で並べ替える（数値を含まない行は末尾）
         * @param {boolean} descending - true で降順
         * @returns {void}
         */
        function sortRowsByValue(descending) {
            var flatEntries = getFlatBodyEntries(frameSnapshots);
            flatEntries.sort(function (a, b) {
                var aValue = extractFirstNumber(a.text), bValue = extractFirstNumber(b.text);
                if (aValue === null && bValue === null) return 0;
                if (aValue === null) return 1;   // 数値なしは末尾へ / push no-number rows to the end
                if (bValue === null) return -1;
                return descending ? (bValue - aValue) : (aValue - bValue);
            });
            applyReorder(flatEntries, []);
        }

        /**
         * リストの種類・番号スタイルに応じてパネルや入力欄をディムする
         * @returns {void}
         */
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
            justifyRow.enabled = !rbNone.value;           // 行揃えは「なし」では変更しない / justification is left alone for "none"

            // 区切り文字カラーは「数字／ABC／abc かつ区切り文字あり」のときだけ有効 / Delimiter color only when number/ABC/abc with a delimiter
            var delimiterColorApplicable = twoStopNumbered && currentDelimiter() !== "";
            delimiterColorUI.label.enabled = delimiterColorApplicable;
            delimiterColorUI.swatch.enabled = delimiterColorApplicable;

            // マーカー位置は数字/ABC/abc のみ有効、本文位置は全種類で有効（ハンギングON時のみ左インデント参照でディム）
            // Marker position only for number/ABC/abc; body position for all types (dimmed under hanging since it follows the left indent)
            var hangingOn = chkHanging.value && !rbNone.value;
            tabStopRow1.enabled = twoStopNumbered;                 // マーカー位置 / marker position
            tabStopRow2.enabled = !rbNone.value && !hangingOn;     // 本文位置（ラベル・単位ごとディム）/ body position (label and unit dim together)
            syncHangingTabStop();                              // 本文位置に左インデントを反映 / mirror left indent into the body position
        }

        /**
         * 現在のUIの状態を applyListMarkers() へ渡すオプションにまとめる
         * @returns {object} applyListMarkers() のオプション
         */
        function currentApplyOptions() {
            return {
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
                justifyId: justifyState.activeId,
                frameSnapshots: frameSnapshots,
                format: currentFormat(),
                paragraphFormat: currentParagraphFormat()
            };
        }

        /**
         * プレビューを更新する
         * 行一覧は並べ替えたときだけ内容が変わるため、ここでは作り直さない
         * @returns {void}
         */
        function updatePreview() {
            if (restoringSession) return; // 復元が終わるまでは描き直さない / hold off until the restore finishes
            applyCurrentSettings();
        }

        /**
         * 直前のプレビューを巻き戻し、控えを書き戻したうえで現在の設定を適用する
         * 並べ替えはJS側の控えで管理しているため、undo のあとに並べ替え済みの素テキストを戻してから付与する
         * @returns {void}
         */
        function applyCurrentSettings() {
            // 本文は applyListMarkers() が控えから組み直すため、ここでは書式だけ元に戻せばよい
            // applyListMarkers() rebuilds the text from the snapshot, so only the formats need resetting here
            restoreFrameFormats(frameSnapshots);
            applyListMarkers(targetSelection, currentApplyOptions());
            app.redraw();
        }

        /**
         * 予約済みのプレビューを取り消す
         * @returns {void}
         */
        function cancelScheduledPreview() {
            if (previewTaskId == null) return;
            try { app.cancelTask(previewTaskId); } catch (e) { }
            previewTaskId = null;
        }

        /**
         * プレビューの更新を予約する
         * 予約中に呼ばれたら前の予約を捨てて取り直すので、操作が続く間は更新されず、
         * 止まってから1回だけ反映される（20→15と押したら15だけを計算する）
         * @returns {void}
         */
        function requestPreview() {
            if (restoringSession) return; // 復元が終わるまでは予約もしない / do not even schedule while restoring
            cancelScheduledPreview();
            try {
                previewTaskId = app.scheduleTask(PREVIEW_TASK_CALL, PREVIEW_DELAY_MS, false);
            } catch (e) {
                previewTaskId = null;
            }
            // 予約できない環境ではまとめられないので、そのまま更新する / no scheduling available: update as before
            if (previewTaskId == null) updatePreview();
        }

        /**
         * 入力欄をプレビュー更新に配線する（入力中は遅延、確定時と↑↓キー操作後にまとめて反映）
         * @param {EditText} input - 対象の入力欄
         * @param {boolean} allowNegative - ↑↓キーで負値を許可するか
         * @param {function} [beforeUpdate] - プレビューの前に実行する処理（省略可）
         * @returns {void}
         */
        function wirePreviewInput(input, allowNegative, beforeUpdate) {
            // 入力・↑↓キーとも予約だけを行う。onChange（確定）で即時更新すると、
            // 1回の操作で予約と即時更新が二重に走るため配線しない
            // Both typing and arrow keys only schedule. Wiring onChange to an immediate update
            // would make a single edit run the preview twice
            input.onChanging = function () {
                if (beforeUpdate) beforeUpdate();
                requestPreview();
            };
            changeValueByArrowKey(input, allowNegative, function () {
                if (beforeUpdate) beforeUpdate();
                requestPreview();
            });
        }

        /**
         * リストの種類が変わったときに既定値とディム状態を更新してプレビューする
         * @returns {void}
         */
        function onTypeChange() {
            applyTypeDefaults();
            // 比率: 番号リストは100へ戻す / 箇条書きは選択中の記号の比率へ / Scale: reset to 100 for numbered; use the selected symbol's scale for bullet
            if (rbNumbered.value) {
                inputScale.text = "100";
                rbAlignRight.value = true; // 番号は右揃えが既定 / numbered defaults to right
            } else if (rbBullet.value) {
                inputScale.text = String(selectedBulletSymbol().scale);
                rbAlignLeft.value = true; // 箇条書きは左揃えが既定 / bullet defaults to left
            }
            updateEnabledStates();
            updatePreview();
        }

        /**
         * 番号スタイルが変わったときに区切り文字・揃え・タブストップの既定を更新してプレビューする
         * @returns {void}
         */
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

        /**
         * 箇条書き記号が変わったときに比率と本文タブストップを記号ごとの既定へ更新してプレビューする
         * @returns {void}
         */
        function onBulletChange() {
            inputScale.text = String(selectedBulletSymbol().scale);
            applyTypeDefaults();   // 本文位置を記号の bodyRatio（•/-=0.8 など）で更新 / refresh body position with the glyph's bodyRatio (e.g. 0.8 for •/-)
            updateEnabledStates();
            updatePreview();
        }

        /**
         * 比率・ベースライン・カラー・インデント・タブストップをまとめてリセットする
         * @returns {void}
         */
        function resetAllValues() {
            inputScale.text = "100";                                   // 比率 / scale
            inputBaseline.text = "0";                                  // ベースラインシフト / baseline shift
            inputStartNumber.text = String(START_NUMBER);              // 開始番号を1へ / start number back to 1
            // カラー（記号／番号・区切り文字とも基準色へ）/ color (marker-number & delimiter back to base)
            var resetColor = getBaseFillColor(targetSelection);
            colorSwatch._aiColor = resetColor;
            delimiterColorSwatch._aiColor = resetColor;
            redrawSwatch(colorSwatch);
            redrawSwatch(delimiterColorSwatch);
            inputSpaceAfter.text = "0";                                // 段落後のアキ / space after
            inputLeftIndent.text = "0";                                // 左インデント（1行目は内部で正負反転）/ left indent (first-line derived internally)
            // ハンギングOFFでもインデントを0として適用する（控えは書き換えないので、キャンセルで原状復帰できる）
            // Apply zero indents even while hanging is off; the snapshot stays intact so Cancel can still restore
            indentsCleared = true;
            applyTypeDefaults();
            updateEnabledStates();
            updatePreview();
        }

        /**
         * 各コントロールにイベントハンドラーを割り当てる
         * @returns {void}
         */
        function wireEventHandlers() {
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
            chkResetPerFrame.onClick = updatePreview;
            for (var justifyWireIndex = 0; justifyWireIndex < justifyButtons.length; justifyWireIndex++) {
                justifyButtons[justifyWireIndex].onClick = function () {
                    selectJustify(this.justifyId);
                    updatePreview();
                };
            }
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

            fontFamilyDropdown.onChange = function () {
                populateFontStyleDropdown(fontStyleDropdown, fontFamilyDropdown.selection, null);
                updatePreview();
            };
            fontStyleDropdown.onChange = updatePreview;

            // ハンギング対応: ディム更新（インデント＋本文タブストップ）→プレビュー / Hanging: refresh dim (indent + body tab stop) then preview
            chkHanging.onClick = function () {
                if (chkHanging.value) applyDefaultHangingIndent();
                updateHangingEnabled();
                updateEnabledStates(); // 本文タブストップにインデント値を反映 / mirror indent into the body tab stop
                updatePreview();
            };
            /* 和文フォントのみ切り替え時はドロップダウンを再構築 / Rebuild the dropdown when toggling Japanese-only */
            chkJPOnly.onClick = function () {
                var currentFontName = getSelectedFontName(fontFamilyDropdown, fontStyleDropdown) || getBaseFontName(targetSelection);
                populateFontFamilyDropdown(fontFamilyDropdown, currentFontName, chkJPOnly.value);
                populateFontStyleDropdown(fontStyleDropdown, fontFamilyDropdown.selection, currentFontName);
                updatePreview();
            };

            // 数値入力欄: 入力中とキー操作は遅延、確定時は即時にプレビュー
            // Numeric fields: coalesce while typing and stepping, apply immediately on commit
            wirePreviewInput(inputStartNumber, false);
            wirePreviewInput(inputTabStop1, false);
            wirePreviewInput(inputTabStop2, false);
            wirePreviewInput(inputScale, false);                          // 比率% は負値不可 / scale % cannot be negative
            wirePreviewInput(inputBaseline, true);                        // ベースラインシフトは負値可 / baseline shift allows negatives
            wirePreviewInput(inputLeading, false);                        // 行送りは負値不可 / leading cannot be negative
            wirePreviewInput(inputSpaceAfter, false);                     // 段落後のアキは負値不可 / space after cannot be negative
            wirePreviewInput(inputLeftIndent, false, syncHangingTabStop); // インデント→本文タブストップ同期 / indent syncs the body tab stop
        }

        /**
         * 意味が伝わりにくいラベル・入力にツールチップを設定する
         * @returns {void}
         */
        function applyTooltips() {
            // ツールチップ（意味が伝わりにくいラベル・入力に補足）/ Tooltips for labels and inputs whose meaning isn't obvious
            tabStopLabel1.helpTip = inputTabStop1.helpTip = getLabel('tooltip.tabStop1');
            tabStopLabel2.helpTip = inputTabStop2.helpTip = getLabel('tooltip.tabStop2');
            rbAlignLeft.helpTip = rbAlignCenter.helpTip = rbAlignRight.helpTip = getLabel('tooltip.tabAlign');
            rbNone.helpTip = getLabel('tooltip.typeNone');
            bulletStylePanel.helpTip = getLabel('tooltip.bulletStyle');
            rbStyleCircledWhite.helpTip = rbStyleCircledBlack.helpTip = getLabel('tooltip.circledNumber');
            scaleLabel.helpTip = inputScale.helpTip = getLabel('tooltip.scale');
            baselineLabel.helpTip = inputBaseline.helpTip = getLabel('tooltip.baselineShift');
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
            justifyLabel.helpTip = getLabel('tooltip.justify');
            chkHanging.helpTip = getLabel('tooltip.hanging');
            leftIndentLabel.helpTip = inputLeftIndent.helpTip = getLabel('tooltip.leftIndent');
            btnShowHidden.helpTip = getLabel('tooltip.showHidden');
            btnReset.helpTip = getLabel('tooltip.reset');
        }

        wireEventHandlers();
        applyTooltips();

        var committed = false;

        btnOK.onClick = function () {
            // 画面に出ている状態がそのまま確定結果になる / what is on screen is what gets committed
            applyCurrentSettings();
            committed = true;
            dialog.close(1);
        };
        btnCancel.onClick = function () { dialog.close(0); };

        /* OK以外で閉じた場合はプレビューを取り消す / Roll back preview unless committed via OK */
        dialog.onClose = function () {
            // 予約が残っているとダイアログを閉じたあとに実行されるため、必ず取り消す
            // A leftover task would run after the dialog is gone, so always cancel it
            cancelScheduledPreview();
            $.global[PREVIEW_TASK_KEY] = null;

            // OK・キャンセルのどちらでも、閉じた時点の設定を覚えておく / Remember the settings as they were, on OK and Cancel alike
            saveSessionState(collectSessionState());
            if (!committed) {
                // 控えから原状へ戻す（本文の並びは控えに従う）/ Restore from the snapshot (the line order follows it)
                restoreFrameSnapshots(frameSnapshots);
                app.redraw();
            }
            return true;
        };

        // 現在の状態から初期選択を復元（「現状の続き」として編集）/ Initialize from the current state so editing continues from it
        applyDetectedState(detectCurrentListState(targetSelection));

        // 同じセッションで前回使った設定があれば、そちらで上書きする / Overwrite with the settings used earlier in this session
        var sessionState = loadSessionState();
        applySessionState(sessionState);

        // 初回は「折り返し位置を揃える」がONなので、インデントに既定値を入れておく
        // On the first run the hanging option is on, so seed the indent with its default
        if (!sessionState.listType && chkHanging.value) applyDefaultHangingIndent();

        // 表示前に初期状態（ディム）だけ反映（プレビューは onShow から起動）/ Apply initial dim state before showing (preview is kicked off from onShow)
        updateEnabledStates();
        updateHangingEnabled();       // ハンギング対応の初期ディム / initial dim state for hanging

        // 初期プレビューは onShow から起動（同期実行だと undo トランザクションが閉じず履歴が残るため）
        // Kick off the first preview from onShow (a synchronous pre-show preview leaves the undo transaction open → history piles up)
        dialog.onShow = function () {
            refreshList();
            updatePreview();
        };
        dialog.show();
    }

    // =========================================
    // 入力補助 / Input helpers
    // =========================================

    /**
     * ↑↓キーで入力欄の値を増減させる（Shift=±10、Option=±0.1）
     * keydown では値の更新だけを行い、keyup で開始値と比較して変化があるときだけ1回コールバックする
     * （押しっぱなしのオートリピートでプレビューが連続実行されるのを防ぐ）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - 負値を許可するか
     * @param {function} onUpdate - 値が変わったときに呼ぶコールバック
     * @returns {void}
     */
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

    /**
     * 丸数字スタイルかどうかを判定する（箇条書きと同じく先頭タブなし・タブストップ1つで扱う）
     * @param {string} style - 番号スタイル
     * @returns {boolean} 丸数字なら true
     */
    function isCircledStyle(style) {
        return style === "circledWhite" || style === "circledBlack";
    }

    /**
     * 行頭の既存マーカー（箇条書き記号・各種番号）を除去する
     * @param {string} line - 対象の1行
     * @returns {string} マーカーを除去した行
     */
    function stripListMarker(line) {
        // 生成済みマーカー（先頭タブ + 番号記号 + 区切り文字(任意) + タブ）を優先的に除去 / Strip our generated marker (leading tab + glyph + optional delimiter + tab)
        var generatedMarkerPattern = /^\t(?:[①-⑳]|[❶-❿]|[⓫-⓴]|[A-Za-z]+|[〇一二三四五六七八九十百千]+|\d+)[.:：|]?\t/;
        if (generatedMarkerPattern.test(line)) return line.replace(generatedMarkerPattern, "");

        // 丸数字（先頭タブなしで「丸数字 + タブ/スペース」）を除去。丸数字は一意なので区切りなしでも対象 / Strip circled marker (no leading tab: circled glyph + tab/space). Circled glyphs are unambiguous, so no delimiter is required
        var circledMarkerPattern = /^(?:[①-⑳]|[❶-❿]|[⓫-⓴])[\t 　]*/;
        if (circledMarkerPattern.test(line)) return line.replace(circledMarkerPattern, "");

        // 手打ちの番号リスト（先頭タブなし・「数字/ABC/abc + 区切り + スペース/タブ」）を除去
        // 区切り文字を必須にして本文（例: "Apple is" / "Mr. Smith"）の誤除去を抑える。"12.5" は区切り直後が数字で [\t 　]+ に一致せず対象外
        // Strip hand-typed lists (no leading tab: "number/ABC/abc + delimiter + space/tab"). A delimiter is required so body text is mostly preserved; "12.5" is excluded because a digit (not whitespace) follows the delimiter
        var handTypedMarkerPattern = /^(?:[A-Za-z]+|\d+)[.:：|][\t 　]+/;
        if (handTypedMarkerPattern.test(line)) return line.replace(handTypedMarkerPattern, "");

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

    /**
     * 1フレーム分のテキストにマーカーを付与した行配列を作る（連番カウンタは呼び出し側と共有）
     * @param {string} frameText - フレームのテキスト
     * @param {string} listType - "bullet" / "numbered" / "none"
     * @param {object} options - 付与オプション（bulletMark / numberStyle / delimiter など）
     * @param {{value: number}} numberCounter - 連番カウンタ
     * @param {number} padWidth - ゼロ埋めの桁数（0で無効）
     * @returns {string[]} マーカーを付与した行の配列
     */
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
                var markerText = formatMarkerText(numberCounter.value, options.numberStyle, padWidth, options.delimiter);
                if (isCircledStyle(options.numberStyle)) {
                    // 丸数字は箇条書きと同じ構造（丸数字 + tab + 本文, タブストップ1つ）/ Circled: bullet-like (glyph + tab + text, one tab stop)
                    line = markerText + "\t" + line;
                } else {
                    // tab + 番号記号 + tab + 本文 / tab + number glyph + tab + text
                    line = "\t" + markerText + "\t" + line;
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

    /**
     * 各テキストフレームの行頭をいったん「なし」に戻したうえで、指定の種類のマーカーを付与する
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @param {object} options - currentApplyOptions() が返す付与オプション
     * @returns {void}
     */
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
            var targetFrame = selection[i];

            // テキストフレームのみ対象 / TextFrame only
            if (targetFrame.typename !== "TextFrame") {
                continue;
            }

            // フレームごとにリセット時は各フレームの先頭で開始番号へ戻す / restart at each frame when requested
            if (options.resetPerFrame) numberCounter.value = startNumber;

            // マーカーは控えの本文から組み立てる（ドキュメントを一度素に戻す必要がない）
            // 組み上がりが今の本文と同じなら書き戻さない。比率・行送り・タブストップだけを変えた場合は
            // 本文に一切触れないので、全文字の書式復元も走らない
            // Build the marked text from the snapshot, so the document never has to be reset first.
            // When the result matches the current text nothing is written, which means a change to
            // scale / leading / tab stops alone never touches the body or its character attributes
            var frameSnapshot = findFrameSnapshot(options.frameSnapshots, targetFrame);
            var sourceContents = frameSnapshot ? frameSnapshot.contents : targetFrame.contents;
            var markedContents = buildMarkedLines(sourceContents, listType, options, numberCounter, padWidth).join("\r");
            if (targetFrame.contents !== markedContents) {
                targetFrame.contents = markedContents;

                // contents 再設定で失われた本文の文字属性を復元 / Restore body attributes lost by resetting contents
                if (frameSnapshot) restoreBodyAttributes(targetFrame, frameSnapshot.bodyAttrsPerLine);
            }

            // 行頭マーカーの送り位置をtabストップで揃える / Align the marker via tab stops
            setTabStops(targetFrame, tabStopSpecsFor(listType, options));

            // 行頭マーカーに書式（フォント・サイズ・ベースラインシフト）を適用 / Apply marker format
            applyMarkerFormat(targetFrame, listType, options.format, options.numberStyle, options.delimiter);

            // 段落の書式（行送り・段落後のアキ・インデント）を適用 / Apply paragraph format (leading, space after, indents)
            applyParagraphFormat(targetFrame, options.paragraphFormat);

            // 段落の行揃えを適用（「なし」では変更しない）/ Apply the paragraph justification (skipped for "none")
            if (listType !== "none") applyJustification(targetFrame, options.justifyId);
        }
    }

    /**
     * 種類に応じたタブストップの指定内容を求める
     * 本文位置は「なし」以外の全種類で使い、マーカー位置は数字／ABC／abc のときだけ加える
     * @param {string} listType - "bullet" / "numbered" / "none"
     * @param {object} options - currentApplyOptions() が返す付与オプション
     * @returns {Array<{position: number, alignment: TabStopAlignment}>} タブストップの指定
     */
    function tabStopSpecsFor(listType, options) {
        if (listType === "none") return []; // 「なし」はタブストップをすべて削除 / "none" clears all tab stops

        // 本文位置は左揃え固定 / the body position is always left-aligned
        var bodyStop = { position: options.tabStop2Pt, alignment: TabStopAlignment.Left };
        if (listType !== "numbered" || isCircledStyle(options.numberStyle)) {
            return [bodyStop]; // 箇条書き・丸数字は1ストップ / bullets and circled numbers use a single stop
        }

        // 数字/ABC/abc はマーカー位置（指定の揃え）を加えた2ストップ / number/ABC/abc add the marker column
        var markerAlignment = options.tab1Alignment || TabStopAlignment.Right;
        return [{ position: options.tabStop1Pt, alignment: markerAlignment }, bodyStop];
    }

    /**
     * 段落の書式（行送り・段落後のアキ・インデント）をフレーム全体へ適用する（null の項目は変更しない）
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {{leadingPt: (number|null), spaceAfterPt: (number|null), leftIndentPt: (number|null), firstLineIndentPt: (number|null)}} paragraphFormat - 段落の書式
     * @returns {void}
     */
    function applyParagraphFormat(frame, paragraphFormat) {
        if (!paragraphFormat) return;
        // 各設定は個別に try で囲む。1項目（例: 行送り）の例外で後続（spaceAfter・インデント）が
        // 巻き込まれてスキップされるのを防ぐ / Guard each setter separately so a failure in one
        // (e.g. leading) does not skip the rest (spaceAfter, indents)
        if (paragraphFormat.leadingPt != null) applyAutoLeading(frame, paragraphFormat.leadingPt);
        if (paragraphFormat.spaceAfterPt != null) {
            try { frame.textRange.paragraphAttributes.spaceAfter = paragraphFormat.spaceAfterPt; } catch (eSpaceAfter) { }
        }
        // インデント（ハンギング対応OFFのときは null で渡され、変更しない）/ Indents (null when hanging is off → unchanged)
        if (paragraphFormat.leftIndentPt != null) {
            try { frame.textRange.paragraphAttributes.leftIndent = paragraphFormat.leftIndentPt; } catch (eLeftIndent) { }
        }
        if (paragraphFormat.firstLineIndentPt != null) {
            try { frame.textRange.paragraphAttributes.firstLineIndent = paragraphFormat.firstLineIndentPt; } catch (eFirstIndent) { }
        }
    }

    /**
     * 段落の行揃えを設定する
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {string} justifyId - 行揃えのID
     * @returns {void}
     */
    function applyJustification(frame, justifyId) {
        try { frame.textRange.paragraphAttributes.justification = resolveJustification(justifyId); } catch (e) { }
    }

    /**
     * 行送りを「自動行送りの値（％）」として適用する
     * 固定値を書き込まず、段落ごとに文字サイズに対する割合を求めて autoLeadingAmount へ入れる
     * 行送りは自動のままなので、あとから文字サイズを変えても行送りが比率で追従する
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {number} leadingPt - 目標の行送り（pt）
     * @returns {void}
     */
    function applyAutoLeading(frame, leadingPt) {
        try {
            var paragraphs = frame.paragraphs;
            for (var paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
                var paragraph = paragraphs[paragraphIndex];
                var fontSize = NaN;
                try { fontSize = paragraph.characterAttributes.size; } catch (eSize) { }
                if (isNaN(fontSize) || fontSize <= 0) continue; // 文字サイズが取れない段落は変更しない / leave paragraphs whose size is unreadable

                try {
                    paragraph.characterAttributes.autoLeading = true;
                    paragraph.paragraphAttributes.autoLeadingAmount = (leadingPt / fontSize) * 100;
                } catch (eLead) { }
            }
        } catch (e) { }
    }

    // =========================================
    // 並べ替え / Reorder
    // =========================================

    /**
     * 控えた各フレームの行数を取得する
     * @param {Array<object>} frameSnapshots - captureFrameSnapshots() が返した控え
     * @returns {number[]} フレームごとの行数
     */
    function getFrameLineCounts(frameSnapshots) {
        var counts = [];
        for (var i = 0; i < frameSnapshots.length; i++) {
            counts.push(splitIntoLines(frameSnapshots[i].contents).length);
        }
        return counts;
    }

    /**
     * 全フレームの本文（マーカー除去後）と文字属性を1つのフラット配列で取得する
     * @param {Array<object>} frameSnapshots - captureFrameSnapshots() が返した控え
     * @returns {Array<{text: string, attrs: Array<object>}>} 行ごとの本文と文字属性
     */
    function getFlatBodyEntries(frameSnapshots) {
        var flat = [];
        for (var i = 0; i < frameSnapshots.length; i++) {
            var lines = splitIntoLines(frameSnapshots[i].contents);
            var attrsPerLine = frameSnapshots[i].bodyAttrsPerLine || [];
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
     * フラット配列の本文と文字属性を、元のフレーム行数どおりに控えへ書き戻す
     * @param {Array<object>} frameSnapshots - captureFrameSnapshots() が返した控え
     * @param {Array<{text: string, attrs: Array<object>}>} flatEntries - 書き戻す行データ
     * @returns {void}
     */
    function writeFlatBodyEntries(frameSnapshots, flatEntries) {
        var counts = getFrameLineCounts(frameSnapshots);
        var index = 0;
        for (var i = 0; i < frameSnapshots.length; i++) {
            var lineCount = counts[i];
            var lines = [];
            var attrsPerLine = [];
            for (var j = 0; j < lineCount; j++) {
                var entry = flatEntries[index + j];
                lines.push(entry ? entry.text : "");
                attrsPerLine.push(entry ? entry.attrs : []);
            }
            frameSnapshots[i].contents = lines.join("\r");
            frameSnapshots[i].bodyAttrsPerLine = attrsPerLine;
            index += lineCount;
        }
    }

    // =========================================
    // 番号スタイル / Number style
    // =========================================

    /**
     * 番号を選択スタイルの文字列へ変換する（末尾に区切り文字を付与）
     * @param {number} numberValue - 変換する番号
     * @param {string} style - 番号スタイル
     * @param {number} padWidth - ゼロ埋めの桁数（0で無効）
     * @param {string} delimiter - 区切り文字（なしは空文字）
     * @returns {string} 変換した文字列
     */
    function formatMarkerText(numberValue, style, padWidth, delimiter) {
        var markerText;
        if (style === "circledWhite") {
            markerText = toCircledWhite(numberValue);
        } else if (style === "circledBlack") {
            markerText = toCircledBlack(numberValue);
        } else if (style === "upperAlpha") {
            markerText = toAlphabet(numberValue, true);
        } else if (style === "lowerAlpha") {
            markerText = toAlphabet(numberValue, false);
        } else {
            // 数字（ゼロ埋め指定があれば桁をそろえる）/ Plain numbers (zero-padded when requested)
            markerText = (padWidth && padWidth > 0) ? zeroPadNumber(numberValue, padWidth) : String(numberValue);
        }
        return markerText + (delimiter || ""); // 区切り文字（なし=空文字）/ delimiter (none = "")
    }

    /**
     * 数値を指定桁数までゼロ埋めする
     * @param {number} numberValue - 対象の数値
     * @param {number} width - 桁数
     * @returns {string} ゼロ埋めした文字列
     */
    function zeroPadNumber(numberValue, width) {
        var text = String(numberValue);
        while (text.length < width) text = "0" + text;
        return text;
    }

    /**
     * 1フレームの番号付け対象行数を数える（空行は除外）
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @returns {number} 番号付け対象の行数
     */
    function numberedItemsInFrame(frame) {
        var lines = splitIntoLines(frame.contents);
        var count = 0;
        for (var j = 0; j < lines.length; j++) {
            if (/\S/.test(stripListMarker(lines[j]))) count++; // 空行はスキップ / skip empty lines
        }
        return count;
    }

    /**
     * 選択全体の番号付け対象行数を数える（通し番号時の桁そろえ・件数チェック用）
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {number} 番号付け対象の総行数
     */
    function countNumberedItems(selection) {
        var total = 0;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename === "TextFrame") total += numberedItemsInFrame(selection[i]);
        }
        return total;
    }

    /**
     * フレーム1つあたりの番号付け対象行数の最大値を求める（フレームごとにリセット時の桁そろえ用）
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {number} 最大の行数
     */
    function maxNumberedItemsPerFrame(selection) {
        var maxCount = 0;
        for (var i = 0; i < selection.length; i++) {
            if (selection[i].typename !== "TextFrame") continue;
            var count = numberedItemsInFrame(selection[i]);
            if (count > maxCount) maxCount = count;
        }
        return maxCount;
    }

    /**
     * 数値を白丸数字（①〜⑳）へ変換する
     * @param {number} numberValue - 変換する番号
     * @returns {string} 白丸数字（範囲外は素の数字）
     */
    function toCircledWhite(numberValue) {
        if (numberValue >= 1 && numberValue <= 20) return String.fromCharCode(0x2460 + (numberValue - 1));
        return String(numberValue); // 範囲外は素の数字（区切り文字は呼び出し側で付与）/ out of range: plain number
    }

    /**
     * 数値を黒丸数字（❶〜⓴）へ変換する
     * @param {number} numberValue - 変換する番号
     * @returns {string} 黒丸数字（範囲外は素の数字）
     */
    function toCircledBlack(numberValue) {
        if (numberValue >= 1 && numberValue <= 10) return String.fromCharCode(0x2776 + (numberValue - 1)); // ❶〜❿
        if (numberValue >= 11 && numberValue <= 20) return String.fromCharCode(0x24EB + (numberValue - 11)); // ⓫〜⓴
        return String(numberValue); // 範囲外は素の数字（区切り文字は呼び出し側で付与）/ out of range: plain number
    }

    /**
     * 数値をアルファベット（A〜Z / a〜z）へ変換する
     * @param {number} numberValue - 変換する番号
     * @param {boolean} useUpperCase - true で大文字
     * @returns {string} アルファベット（範囲外は素の数字）
     */
    function toAlphabet(numberValue, useUpperCase) {
        if (numberValue >= 1 && numberValue <= 26) return String.fromCharCode((useUpperCase ? 0x41 : 0x61) + (numberValue - 1));
        return String(numberValue);
    }

    // =========================================
    // 書式（マーカーへのフォント・サイズ・ベースラインシフト）/ Marker format
    // =========================================

    /**
     * 段落テキストの中で、行頭マーカーが占める文字範囲を求める
     * @param {string} paragraphText - 段落のテキスト
     * @param {string} listType - "bullet" / "numbered" / "none"
     * @param {string} numberStyle - 番号スタイル
     * @returns {{start: number, end: number}|null} マーカーの範囲（マーカーが無ければ null）
     */
    function markerRangeOf(paragraphText, listType, numberStyle) {
        if (listType === "bullet" || (listType === "numbered" && isCircledStyle(numberStyle))) {
            // 記号1字（箇条書き／丸数字）。「記号 + タブ」でない行（空行など未付与）は対象外
            // Single glyph (bullet / circled); lines that aren't "glyph + tab" have no marker
            if (paragraphText.charAt(1) !== "\t") return null;
            return { start: 0, end: 1 };
        }

        // 数字/ABC/abc: 先頭タブを飛ばし、次のタブまで / number/ABC/abc: skip the leading tab, up to the next tab
        if (paragraphText.charAt(0) !== "\t") return null; // 先頭タブが無い行（未付与）は対象外 / no leading tab = no marker
        var markerEnd = 1;
        while (markerEnd < paragraphText.length && paragraphText.charAt(markerEnd) !== "\t") markerEnd++;
        return { start: 1, end: markerEnd };
    }

    /**
     * 行頭マーカーの文字にフォント・比率・ベースラインシフト・カラーを適用する
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {string} listType - "bullet" / "numbered" / "none"
     * @param {object} format - currentFormat() が返す書式
     * @param {string} numberStyle - 番号スタイル
     * @param {string} delimiter - 区切り文字（なしは空文字）
     * @returns {void}
     */
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
            for (var paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
                var paragraph = paragraphs[paragraphIndex];
                var paragraphText = paragraph.contents;
                if (!paragraphText || paragraphText.length === 0) continue;

                // マーカー文字の範囲を求める（マーカー未付与の行は null）/ Locate the marker range (null when the line has no marker)
                var markerRange = markerRangeOf(paragraphText, listType, numberStyle);
                if (!markerRange) continue;
                var markerStart = markerRange.start, markerEnd = markerRange.end;

                // 区切り文字部分の開始位置（マーカー範囲の末尾から delimiterLen 文字）/ Where the delimiter chars start (last delimiterLen chars of the marker range)
                var delimiterStart = markerEnd - delimiterLen;
                for (var charIndex = markerStart; charIndex < markerEnd; charIndex++) {
                    if (charIndex >= paragraph.characters.length) break;
                    var charAttr = paragraph.characters[charIndex].characterAttributes;
                    if (font) { try { charAttr.textFont = font; } catch (eFont) { } }
                    if (format.horizontalScale) { try { charAttr.horizontalScale = format.horizontalScale; } catch (eHScale) { } }
                    if (format.verticalScale) { try { charAttr.verticalScale = format.verticalScale; } catch (eVScale) { } }
                    try { charAttr.baselineShift = format.baselineShiftPt; } catch (eBaseline) { }
                    // カラー: 区切り文字は専用色、それ以外（記号・番号）はマーカー色 / Color: delimiter chars use their own color, the rest use the marker color
                    var isDelimiterChar = (delimiterLen > 0 && charIndex >= delimiterStart);
                    var fillColorToUse = isDelimiterChar ? format.delimiterFillColor : format.fillColor;
                    if (fillColorToUse) { try { charAttr.fillColor = fillColorToUse; } catch (eFill) { } }
                }
            }
        } catch (eOuter) { }
    }

    /**
     * 先頭テキストフレームのフォント名（Illustrator内部名）を取得する
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {string|null} フォント名（取得できなければ null）
     */
    function getBaseFontName(selection) {
        var frame = firstTextFrame(selection);
        if (!frame) return null;
        try {
            var font = frame.textRange.characterAttributes.textFont;
            if (font) return font.name;
        } catch (e) { }
        return null;
    }

    /**
     * Illustratorのフォント一覧をファミリー別にまとめる
     * @param {boolean} japaneseOnly - true なら和文フォントだけに絞り込む
     * @returns {{names: string[], map: object}} ファミリー名の配列と、ファミリー名→スタイル配列のマップ
     */
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
        for (var i = 0; i < familyNames.length; i++) {
            familyMap[familyNames[i]].sort(function (a, b) { return (a.style < b.style) ? -1 : (a.style > b.style ? 1 : 0); });
        }

        return { names: familyNames, map: familyMap };
    }

    /**
     * フォント名から TextFont を取得する
     * @param {string} fontName - Illustrator内部のフォント名
     * @returns {TextFont|null} 見つかったフォント（なければ null）
     */
    function findTextFontByName(fontName) {
        if (!fontName) return null;
        for (var i = 0; i < app.textFonts.length; i++) {
            if (app.textFonts[i].name === fontName) return app.textFonts[i];
        }
        return null;
    }

    /**
     * フォントファミリーのドロップダウンを作り直す
     * @param {DropDownList} dropdown - 対象のドロップダウン
     * @param {string} selectedFontName - 選択状態にしたいフォント名
     * @param {boolean} japaneseOnly - true なら和文フォントだけに絞り込む
     * @returns {void}
     */
    function populateFontFamilyDropdown(dropdown, selectedFontName, japaneseOnly) {
        dropdown.removeAll();

        var familyData = buildFontFamilyMap(japaneseOnly);
        dropdown._fontFamilyMap = familyData.map;

        var selectedFont = findTextFontByName(selectedFontName);
        var selectedFamily = selectedFont ? (selectedFont.family || selectedFont.name) : null;
        var selectedIndex = 0;

        for (var i = 0; i < familyData.names.length; i++) {
            dropdown.add("item", familyData.names[i]);
            if (selectedFamily && familyData.names[i] === selectedFamily) selectedIndex = i;
        }

        if (dropdown.items.length > 0) dropdown.selection = Math.min(selectedIndex, dropdown.items.length - 1);
    }

    /**
     * 選択中のファミリーに対応するスタイルのドロップダウンを作り直す
     * @param {DropDownList} dropdown - 対象のドロップダウン
     * @param {ListItem} familySelection - ファミリードロップダウンの選択項目
     * @param {string} selectedFontName - 選択状態にしたいフォント名
     * @returns {void}
     */
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

    /**
     * UIで選択されたファミリー＋スタイルから Illustrator 内部のフォント名を求める
     * @param {DropDownList} familyDropdown - ファミリーのドロップダウン
     * @param {DropDownList} styleDropdown - スタイルのドロップダウン
     * @returns {string|null} フォント名（決まらなければ null）
     */
    function getSelectedFontName(familyDropdown, styleDropdown) {
        if (!familyDropdown || !familyDropdown.selection || !styleDropdown || !styleDropdown.selection) return null;
        if (!styleDropdown._fontStyleNames) return null;
        return styleDropdown._fontStyleNames[styleDropdown.selection.index] || null;
    }

    /**
     * 文字列が和文文字（ひらがな・カタカナ・漢字）を含むかを判定する
     * @param {string} text - 対象の文字列
     * @returns {boolean} 含むなら true
     */
    function hasJapaneseCharacters(text) {
        if (!text) return false;
        return /[぀-ゟ゠-ヿ一-鿿]/.test(String(text));
    }

    /**
     * 和文フォントかどうかを判定する（AutoTouchType.jsx のロジックを流用）
     * @param {TextFont} font - 対象のフォント
     * @returns {boolean} 和文フォントなら true
     */
    function isJapaneseFont(font) {
        if (!font) return false;

        // メタデータが壊れたフォントでは識別子ごとに読み取りが失敗しうるため、個別に囲む
        // A broken font can fail per identifier, so guard each read separately
        var identifiers = [];
        try { identifiers.push(font.name ? String(font.name) : ""); } catch (eName) { }
        try { identifiers.push(font.family ? String(font.family) : ""); } catch (eFamily) { }
        try { identifiers.push(font.fullName ? String(font.fullName) : ""); } catch (eFull) { }
        try { identifiers.push(font.postScriptName ? String(font.postScriptName) : ""); } catch (ePS) { }

        // 明示的に除外（ラテン／意図しないゴシック系）/ Explicit exclusions (Latin / unintended Gothic variants)
        if (containsAnyKeyword(identifiers, JP_FONT_DENY_KEYWORDS)) return false;

        // 識別子に和文文字が含まれていれば和文 / A Japanese character in any identifier means Japanese
        for (var i = 0; i < identifiers.length; i++) {
            if (hasJapaneseCharacters(identifiers[i])) return true;
        }

        // キーワード判定（和文ファミリー・主要ブランド）/ Keyword-based heuristic (JP families and common brands)
        return containsAnyKeyword(identifiers, JP_FONT_KEYWORDS);
    }

    /**
     * 識別子のいずれかがキーワードのいずれかを含むかを判定する
     * @param {string[]} identifiers - 判定対象の文字列
     * @param {string[]} keywords - 探すキーワード
     * @returns {boolean} 1つでも一致すれば true
     */
    function containsAnyKeyword(identifiers, keywords) {
        for (var i = 0; i < keywords.length; i++) {
            if (!keywords[i]) continue;
            for (var j = 0; j < identifiers.length; j++) {
                if (identifiers[j].indexOf(keywords[i]) !== -1) return true;
            }
        }
        return false;
    }

    // =========================================
    // tabストップ / Tab stop
    // =========================================

    /**
     * テキストの表示単位を取得する（環境設定の "text/units" を参照）
     * @returns {{label: string, factor: number}} 単位ラベルと1単位あたりのポイント数
     */
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

    /**
     * 先頭テキストフレームの文字サイズを取得する
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {number} 文字サイズ（pt、取得できなければ DEFAULT_FONT_SIZE）
     */
    function getBaseFontSize(selection) {
        var frame = firstTextFrame(selection);
        if (frame) {
            try {
                var fontSize = frame.textRange.characterAttributes.size;
                if (fontSize && !isNaN(fontSize)) return fontSize;
            } catch (e) { }
        }
        return DEFAULT_FONT_SIZE;
    }

    /**
     * 先頭テキストフレームの行送りを取得する
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {number|null} 行送り（pt、取得できなければ null）
     */
    function getBaseLeading(selection) {
        var frame = firstTextFrame(selection);
        if (frame) {
            try {
                var leading = frame.textRange.characterAttributes.leading;
                if (leading != null && !isNaN(leading)) return leading;
            } catch (e) { }
        }
        return null;
    }

    /**
     * 指定の位置と揃えでタブストップを生成する
     * @param {number} positionPt - タブストップの位置（pt）
     * @param {TabStopAlignment} alignment - 揃え種類
     * @returns {TabStopInfo} 生成したタブストップ
     */
    function makeTabStop(positionPt, alignment) {
        var tab = new TabStopInfo();
        tab.alignment = alignment;
        tab.position = positionPt;
        return tab;
    }

    /**
     * タブストップを指定内容だけに置き換える（前の種類の残存タブを消すため全消去してから設定）
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {Array<{position: number, alignment: TabStopAlignment}>} tabSpecs - 設定するタブストップ
     * @returns {void}
     */
    function setTabStops(frame, tabSpecs) {
        try {
            var newTabs = [];
            for (var i = 0; i < tabSpecs.length; i++) {
                newTabs.push(makeTabStop(tabSpecs[i].position, tabSpecs[i].alignment));
            }

            // 位置順にソート / Sort by position
            newTabs.sort(function (a, b) { return a.position - b.position; });

            frame.textRange.paragraphAttributes.tabStops = newTabs;
        } catch (e) { }
    }

    // =========================================
    // プレビュー状態の保存・復元 / Preview snapshot & restore
    // =========================================

    /**
     * テキスト・タブストップ・段落書式・本文の文字属性の現状を控える
     * @param {Array<TextFrame>} selection - 対象のテキストフレーム配列
     * @returns {Array<object>} フレームごとの控え
     */
    function captureFrameSnapshots(selection) {
        var frameSnapshots = [];
        for (var i = 0; i < selection.length; i++) {
            var frame = selection[i];
            if (frame.typename !== "TextFrame") continue;

            var entry = { frame: frame, contents: frame.contents, tabStops: null, bodyAttrsPerLine: null, paraFormat: null };
            try {
                var existingTabs = frame.textRange.paragraphAttributes.tabStops;
                entry.tabStops = [];
                for (var tabIndex = 0; tabIndex < existingTabs.length; tabIndex++) {
                    // 復元用に位置と揃えだけ控える / Keep position and alignment for restore
                    entry.tabStops.push({ position: existingTabs[tabIndex].position, alignment: existingTabs[tabIndex].alignment });
                }
            } catch (e) { }
            // 段落の書式（行送り・自動行送りとその値・揃え・段落後のアキ・インデント）を退避 / Snapshot paragraph format (leading, auto-leading and its amount, justification, space after, indents)
            // 混在した値の読み取りは項目ごとに失敗しうるため、1項目の失敗で他を落とさないよう個別に囲む
            // Reading a mixed value can fail per item, so guard each one to keep the others
            entry.paraFormat = {};
            try { entry.paraFormat.leading = frame.textRange.characterAttributes.leading; } catch (eLead) { }
            try { entry.paraFormat.autoLeading = frame.textRange.characterAttributes.autoLeading; } catch (eAuto) { }
            try { entry.paraFormat.autoLeadingAmount = frame.textRange.paragraphAttributes.autoLeadingAmount; } catch (eAmount) { }
            try { entry.paraFormat.justification = frame.textRange.paragraphAttributes.justification; } catch (eJustify) { }
            try { entry.paraFormat.spaceAfter = frame.textRange.paragraphAttributes.spaceAfter; } catch (eSpace) { }
            try { entry.paraFormat.leftIndent = frame.textRange.paragraphAttributes.leftIndent; } catch (eLeft) { }
            try { entry.paraFormat.firstLineIndent = frame.textRange.paragraphAttributes.firstLineIndent; } catch (eFirst) { }
            // 本文（マーカー以降）の文字属性を退避 / Snapshot the body character attributes (after the marker)
            entry.bodyAttrsPerLine = captureBodyAttributes(frame);
            frameSnapshots.push(entry);
        }
        return frameSnapshots;
    }

    /**
     * 控えておいた状態へ復元する
     * @param {Array<object>} frameSnapshots - captureFrameSnapshots() が返した控え
     * @returns {void}
     */
    function restoreFrameSnapshots(frameSnapshots) {
        for (var i = 0; i < frameSnapshots.length; i++) {
            var entry = frameSnapshots[i];

            // contents の入れ直しは文字書式を初期化するため、変わっていなければ触らない
            // Reassigning contents resets the character formatting, so skip it when the text already matches
            var contentsChanged = true;
            try { contentsChanged = (entry.frame.contents !== entry.contents); } catch (eRead) { }
            if (contentsChanged) {
                try { entry.frame.contents = entry.contents; } catch (e) { }
                restoreBodyAttributes(entry.frame, entry.bodyAttrsPerLine);
            }
        }
        restoreFrameFormats(frameSnapshots);
    }

    /**
     * 控えから段落の書式とタブストップだけを戻す（本文と文字属性には触らない）
     * プレビューのたびに呼ぶため、フレーム単位の軽い書き込みだけで済ませる
     * @param {Array<object>} frameSnapshots - captureFrameSnapshots() が返した控え
     * @returns {void}
     */
    function restoreFrameFormats(frameSnapshots) {
        for (var i = 0; i < frameSnapshots.length; i++) {
            var entry = frameSnapshots[i];

            if (entry.tabStops !== null) {
                try {
                    var rebuiltTabs = [];
                    for (var tabIndex = 0; tabIndex < entry.tabStops.length; tabIndex++) {
                        var tab = new TabStopInfo();
                        tab.alignment = entry.tabStops[tabIndex].alignment;
                        tab.position = entry.tabStops[tabIndex].position;
                        rebuiltTabs.push(tab);
                    }
                    entry.frame.textRange.paragraphAttributes.tabStops = rebuiltTabs;
                } catch (e) { }
            }

            // 段落の書式を復元 / Restore paragraph format
            // キャンセル時の原状復帰が途中で止まらないよう、項目ごとに try を分ける
            // Guard each item so a single failure cannot abort the rest of the rollback
            if (!entry.paraFormat) continue;
            // 行送りの代入は自動行送りを解除するため、autoLeading は最後に戻す
            // Assigning leading turns auto-leading off, so restore autoLeading last
            if (entry.paraFormat.leading != null) { try { entry.frame.textRange.characterAttributes.leading = entry.paraFormat.leading; } catch (eLead) { } }
            if (entry.paraFormat.autoLeadingAmount != null) { try { entry.frame.textRange.paragraphAttributes.autoLeadingAmount = entry.paraFormat.autoLeadingAmount; } catch (eAmount) { } }
            if (entry.paraFormat.autoLeading != null) { try { entry.frame.textRange.characterAttributes.autoLeading = entry.paraFormat.autoLeading; } catch (eAuto) { } }
            if (entry.paraFormat.justification != null) { try { entry.frame.textRange.paragraphAttributes.justification = entry.paraFormat.justification; } catch (eJustify) { } }
            if (entry.paraFormat.spaceAfter != null) { try { entry.frame.textRange.paragraphAttributes.spaceAfter = entry.paraFormat.spaceAfter; } catch (eSpace) { } }
            if (entry.paraFormat.leftIndent != null) { try { entry.frame.textRange.paragraphAttributes.leftIndent = entry.paraFormat.leftIndent; } catch (eLeft) { } }
            if (entry.paraFormat.firstLineIndent != null) { try { entry.frame.textRange.paragraphAttributes.firstLineIndent = entry.paraFormat.firstLineIndent; } catch (eFirst) { } }
        }
    }

    // =========================================
    // 本文の文字属性の退避・復元 / Body character attributes snapshot & restore
    // =========================================
    // contents の再設定でフレーム全体の文字書式が初期化されるため、本文（マーカー以降）の属性を退避し復元する
    // Setting .contents resets the frame's character formatting, so the body (after the marker) is snapshotted and restored.

    /**
     * 1文字分の主要な文字属性を控える
     * @param {CharacterAttributes} characterAttr - 対象の文字属性
     * @returns {object} 控えた属性
     */
    function snapshotCharAttributes(characterAttr) {
        var snap = {};

        // 文字数ぶん繰り返す処理なので、まずは1つの try でまとめて読む（速い経路）
        // This runs once per character, so read everything in a single try first (fast path)
        try {
            snap.textFont = characterAttr.textFont;
            snap.size = characterAttr.size;
            snap.horizontalScale = characterAttr.horizontalScale;
            snap.verticalScale = characterAttr.verticalScale;
            snap.baselineShift = characterAttr.baselineShift;
            snap.tracking = characterAttr.tracking;
            snap.fillColor = characterAttr.fillColor;
            return snap;
        } catch (eBatch) { }

        // まとめて読めなかったときだけ属性ごとに読み直す（1つの失敗で控え全体を失わないように）
        // Fall back to one try per attribute so a single failure cannot discard the whole snapshot
        try { snap.textFont = characterAttr.textFont; } catch (eFont) { }
        try { snap.size = characterAttr.size; } catch (eSize) { }
        try { snap.horizontalScale = characterAttr.horizontalScale; } catch (eHScale) { }
        try { snap.verticalScale = characterAttr.verticalScale; } catch (eVScale) { }
        try { snap.baselineShift = characterAttr.baselineShift; } catch (eBaseline) { }
        try { snap.tracking = characterAttr.tracking; } catch (eTracking) { }
        try { snap.fillColor = characterAttr.fillColor; } catch (eFill) { }
        return snap;
    }

    /**
     * 控えた文字属性を1文字へ復元する
     * @param {CharacterAttributes} characterAttr - 復元先の文字属性
     * @param {object} snapshot - snapshotCharAttributes() が返した控え
     * @returns {void}
     */
    function restoreCharAttributes(characterAttr, snapshot) {
        if (!snapshot) return;

        // 文字数ぶん繰り返す処理なので、まずは1つの try でまとめて書く（速い経路）
        // This runs once per character, so write everything in a single try first (fast path)
        try {
            if (snapshot.textFont) characterAttr.textFont = snapshot.textFont;
            if (snapshot.size != null) characterAttr.size = snapshot.size;
            if (snapshot.horizontalScale != null) characterAttr.horizontalScale = snapshot.horizontalScale;
            if (snapshot.verticalScale != null) characterAttr.verticalScale = snapshot.verticalScale;
            if (snapshot.baselineShift != null) characterAttr.baselineShift = snapshot.baselineShift;
            if (snapshot.tracking != null) characterAttr.tracking = snapshot.tracking;
            if (snapshot.fillColor) characterAttr.fillColor = snapshot.fillColor;
            return;
        } catch (eBatch) { }

        // まとめて書けなかったときだけ属性ごとに書き直す（1つの失敗で残りを巻き込まないように）
        // Fall back to one try per attribute so a single failure cannot skip the rest
        if (snapshot.textFont) { try { characterAttr.textFont = snapshot.textFont; } catch (eFont) { } }
        if (snapshot.size != null) { try { characterAttr.size = snapshot.size; } catch (eSize) { } }
        if (snapshot.horizontalScale != null) { try { characterAttr.horizontalScale = snapshot.horizontalScale; } catch (eHScale) { } }
        if (snapshot.verticalScale != null) { try { characterAttr.verticalScale = snapshot.verticalScale; } catch (eVScale) { } }
        if (snapshot.baselineShift != null) { try { characterAttr.baselineShift = snapshot.baselineShift; } catch (eBaseline) { } }
        if (snapshot.tracking != null) { try { characterAttr.tracking = snapshot.tracking; } catch (eTracking) { } }
        if (snapshot.fillColor) { try { characterAttr.fillColor = snapshot.fillColor; } catch (eFill) { } }
    }

    /**
     * 各段落の本文（マーカー以降）の文字属性を控える
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @returns {Array<Array<object>>} 段落ごと・文字ごとの属性
     */
    function captureBodyAttributes(frame) {
        var attrsPerParagraph = [];
        try {
            var paragraphs = frame.paragraphs;
            for (var paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex++) {
                var text = paragraphs[paragraphIndex].contents;
                var markerLength = text.length - stripListMarker(text).length; // 既存マーカー長 / existing marker length
                var paragraphAttrs = [];
                var characters = paragraphs[paragraphIndex].characters; // ループの外で1度だけ取る / resolve once per paragraph
                for (var k = markerLength; k < text.length; k++) {
                    paragraphAttrs.push(snapshotCharAttributes(characters[k].characterAttributes));
                }
                attrsPerParagraph.push(paragraphAttrs);
            }
        } catch (e) { }
        return attrsPerParagraph;
    }

    /**
     * 控えた本文の文字属性を、再構築後の本文へ復元する（本文の文字数は不変なので末尾から数えて対応づける）
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {Array<Array<object>>} attrsPerParagraph - captureBodyAttributes() が返した控え
     * @returns {void}
     */
    function restoreBodyAttributes(frame, attrsPerParagraph) {
        if (!attrsPerParagraph) return;
        try {
            var paragraphs = frame.paragraphs;
            for (var paragraphIndex = 0; paragraphIndex < paragraphs.length && paragraphIndex < attrsPerParagraph.length; paragraphIndex++) {
                var paragraphAttrs = attrsPerParagraph[paragraphIndex];
                if (!paragraphAttrs || paragraphAttrs.length === 0) continue;

                var text = paragraphs[paragraphIndex].contents;
                var bodyStart = text.length - paragraphAttrs.length; // 本文は末尾 paragraphAttrs.length 文字 / body = last N characters
                if (bodyStart < 0) continue;

                // 文字コレクションはループの外で1度だけ取る / resolve the character collection once per paragraph
                var characters = paragraphs[paragraphIndex].characters;
                var characterCount = characters.length;
                for (var k = 0; k < paragraphAttrs.length; k++) {
                    var charIndex = bodyStart + k;
                    if (charIndex >= characterCount) break;
                    restoreCharAttributes(characters[charIndex].characterAttributes, paragraphAttrs[k]);
                }
            }
        } catch (e) { }
    }

    /**
     * テキストフレームに対応する控えを探す
     * @param {Array<object>} frameSnapshots - captureFrameSnapshots() が返した控え
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @returns {object|null} 対応する控え（なければ null）
     */
    function findFrameSnapshot(frameSnapshots, frame) {
        if (!frameSnapshots) return null;
        for (var i = 0; i < frameSnapshots.length; i++) {
            if (frameSnapshots[i].frame === frame) return frameSnapshots[i];
        }
        return null;
    }
})();
