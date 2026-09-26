#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形からエリア内文字をつくり、そのまま体裁（サイズ・行送り・行揃え・日本語の組版など）を調整します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AreaTypeToolkit.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nfd6cc5e13654

### Overview

Builds area text from point text, text on a path, or a shape, and adjusts its formatting — size, leading, justification and Japanese composition — in the same pass.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AreaTypeToolkit.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AreaTypeToolkit";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AreaTypeToolkit.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AreaTypeToolkit.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nfd6cc5e13654"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 「選択した図形にダミーテキスト」で流し込むダミー文字 / Dummy text for "Dummy text in selected shape" */
    var DUMMY_TEXT_JA = "今日は15:00から予定外のミーティングがありました。疲れを癒すため、夕方に近くのカフェWAVEで、お気に入りの抹茶ラテを楽しみました。短い休憩でしたが、心が“ほっと”しました。";
    var DUMMY_TEXT_EN = "An unplanned meeting popped up at 3:00 PM today. To reset, I stopped by a nearby café for a matcha latte. It was a short break, but it helped me breathe and refocus.";

    /* ダミーテキストの優先フォント候補とサイズ / Preferred fonts and size for dummy text */
    var DUMMY_FONT_JA = ["HiraginoSans-W3", "Hiragino Sans W3"];
    var DUMMY_FONT_EN = ["MyriadPro-Regular", "Myriad Pro Regular", "MyriadPro", "Myriad"];
    var DUMMY_FONT_SIZE = 10;

    /* 日本語の組版の初期値（テキスト側で未設定のときに使う）/ Defaults for Japanese composition (used when the text has none set) */
    var DEFAULT_KINSOKU = "Soft_v2";   /* 弱い禁則 v2 / Loose v2 */
    var DEFAULT_MOJIKUMI_INDEX = 6;    /* ベタ組み / Solid */

    /* 種別プリセット（本文／見出し／メニュー）。ラジオを押すと行送り・行揃え・禁則・文字組み・タブをまとめて設定する
       Role presets (Body / Heading / Menu): one click sets leading, justification, kinsoku, mojikumi and tabs */
    var ROLE_PRESETS = {
        body: { leadingPercent: 160, justifyId: "lastLineLeft", kinsoku: "Soft_v2", mojikumiIndex: 6, alignId: "top", tabMode: "clear" },
        heading: { leadingPercent: 120, justifyId: "left", kinsoku: "Soft_v2", mojikumiIndex: 5, alignId: "top", tabMode: "clear" },
        menu: { leadingPercent: 150, justifyId: "right", kinsoku: "Soft_v2", mojikumiIndex: 5, alignId: "top", tabMode: "leader" }
    };

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

    /* 日英ラベル定義（カテゴリ別）/ Japanese-English labels grouped by category */
    var LABELS = {
        dialog: {
            title: { ja: "エリア内文字を調整", en: "Adjust Area Type" },
            convertTitle: { ja: "エリア内文字に変換", en: "Convert to Area Type" },
            separateTitle: { ja: "テキストを分離", en: "Separate text" }
        },
        panel: {
            createMethod: { ja: "作成方法", en: "Creation method" },
            role: { ja: "種別", en: "Role" },
            frameHandling: { ja: "枠の処理", en: "Frame handling" },
            textHandling: { ja: "テキストの処理", en: "Text handling" },
            leading: { ja: "行送り", en: "Leading" },
            justification: { ja: "行揃え", en: "Justification" },
            textAlign: { ja: "垂直方向の配置", en: "Vertical alignment" },
            fontSize: { ja: "フォントサイズ", en: "Font size" },
            frameSize: { ja: "フレームサイズ", en: "Frame size" },
            indent: { ja: "インデント", en: "Indent" },
            jpComposition: { ja: "日本語の組版", en: "Japanese composition" },
            offset: { ja: "オフセット", en: "Offset" }
        },
        radio: {
            strokeBlack: { ja: "枠を1pt黒に", en: "Frame: 1pt black" },
            hidePath: { ja: "枠を塗り・線なしに", en: "Frame: unpainted" },
            removePath: { ja: "枠を削除", en: "Delete frame" },
            roleBody: { ja: "本文", en: "Body" },
            roleHeading: { ja: "見出し", en: "Heading" },
            roleMenu: { ja: "メニュー", en: "Menu" },
            styleSimple: { ja: "シンプル", en: "Simple" },
            styleButton: { ja: "ボタン風", en: "Button style" },
            useShape: { ja: "選択した図形に流し込む", en: "Pour into selected shape" },
            useShapeDummy: { ja: "選択した図形にダミーテキスト", en: "Dummy text in selected shape" }
        },
        iconButton: {
            justifyLeft: { ja: "左揃え", en: "Left" },
            justifyCenter: { ja: "中央揃え", en: "Center" },
            justifyRight: { ja: "右揃え", en: "Right" },
            justifyLastLineLeft: { ja: "均等配置（最終行左揃え）", en: "Justify (last line left)" },
            justifyAllLines: { ja: "両端揃え", en: "Justify all lines" },
            alignTop: { ja: "上揃え", en: "Top" },
            alignCenter: { ja: "中央揃え", en: "Center" },
            alignBottom: { ja: "下揃え", en: "Bottom" },
            alignJustify: { ja: "均等配置", en: "Justify" }
        },
        dropdown: {
            kinsokuNone: { ja: "なし", en: "None" },
            kinsokuHard: { ja: "強い禁則", en: "Strict" },
            kinsokuSoft: { ja: "弱い禁則", en: "Loose" },
            kinsokuSoftV2: { ja: "弱い禁則 v2", en: "Loose v2" },
            mojikumiNone: { ja: "なし", en: "None" },
            mojikumiLineEndFullHalf: { ja: "行末約物全角/半角", en: "Line-end punct full/half" },
            mojikumiPunctHalf: { ja: "約物半角", en: "Half-width punctuation" },
            mojikumiLineEndHalf: { ja: "行末約物半角", en: "Line-end punct half" },
            mojikumiLineEndFull: { ja: "行末約物全角", en: "Line-end punct full" },
            mojikumiPunctFull: { ja: "約物全角", en: "Full-width punctuation" },
            mojikumiTight: { ja: "ツメ組み", en: "Tight" },
            mojikumiSolid: { ja: "ベタ組み", en: "Solid" }
        },
        checkbox: {
            linkIndents: { ja: "連動", en: "Link" },
            autoSize: { ja: "自動サイズ調整", en: "Auto-size" },
            resolveOverset: { ja: "オーバーセットを解決する", en: "Resolve overset" },
            forceLineBreaks: { ja: "見かけの改行を強制改行に変換", en: "Convert visual line breaks to hard returns" }
        },
        button: {
            convert: { ja: "変換", en: "Convert" },
            shrinkToFit: { ja: "文字あふれ解消", en: "Clear overset" },
            fitFontSize: { ja: "枠にフィット", en: "Fit to frame" },
            separateText: { ja: "テキストを分離...", en: "Separate text..." },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        fieldLabel: {
            leadingPercent: { ja: "行送り", en: "Leading" },
            leadingEffective: { ja: "実寸", en: "Actual" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            charsPerLine: { ja: "文字", en: "chars" },
            kinsoku: { ja: "禁則", en: "Kinsoku" },
            mojikumi: { ja: "文字組み", en: "Mojikumi" },
            indentLeft: { ja: "左", en: "Left" },
            indentRight: { ja: "右", en: "Right" }
        },
        tooltip: {
            styleSimple: {
                ja: "ポイント文字の実寸＋1ptの長方形をフレームにします。",
                en: "Uses a rectangle the size of the point text plus 1pt as the frame."
            },
            styleButton: {
                ja: "幅1.2倍・高さ1.8倍の長方形をフレームにし、行揃えとテキストの配置を中央にします。",
                en: "Uses a rectangle 1.2x wide and 1.8x tall, with the text centered both ways."
            },
            useShape: {
                ja: "選択した図形を複製してフレームにし、選択したテキストを流し込みます。複数組を選べます。",
                en: "Duplicates each selected shape as a frame and pours the selected text into it. Several pairs at once."
            },
            useShapeDummy: {
                ja: "選択した図形自体をエリア内文字にして、ダミー文字を流し込みます。",
                en: "Turns each selected shape itself into Area Type and fills it with dummy text."
            },
            shrinkToFit: {
                ja: "あふれが消えるまでフォントサイズを縮小します。改行を含むテキストには使えません。サイズが混在していると1つのサイズに揃います。",
                en: "Shrinks the font until the overset clears. Not available for text with line breaks. Mixed sizes are flattened to a single size."
            },
            fitFontSize: {
                ja: "あふれるまで拡大してから縮小し、枠いっぱいに収めます。改行を含むテキストには使えません。",
                en: "Grows until it oversets, then shrinks to fill the frame. Not available for text with line breaks."
            },
            autoSize: {
                ja: "エリア内文字の自動サイズ調整。クリックするとその場で反映されます。ONのまま［文字あふれ解消］［枠にフィット］を押すと、いったんOFFにして実行し、またONに戻します。",
                en: "Area Type auto-sizing. Applied as soon as you click. Clear overset and Fit to frame switch it off for the pass and back on afterwards."
            },
            leading: {
                ja: "自動行送り量（%）。実寸＝フォントサイズ×%。実寸に入力すると%を逆算します。固定行送りのテキストでは空で開きます。",
                en: "Auto-leading amount in %. Actual = font size x %. Enter an actual value to back-calculate the %. Empty for text with a fixed leading."
            },
            textAlign: {
                ja: "フレーム内でのテキストの縦位置。クリックするとその場で反映されます。",
                en: "Vertical placement inside the frame. Applied as soon as you click."
            },
            charsPerLine: {
                ja: "1行の文字数から幅を逆算します。字幅が一定でないため日本語UIのみ。",
                en: "Works the width back out from the characters per line. Japanese UI only, since Roman glyph widths vary."
            },
            roleBody: {
                ja: "本文向けの設定をまとめて適用します（行送り160%・均等配置（最終行左揃え）・弱い禁則 v2・ベタ組み・上揃え）。タブ設定は削除します。",
                en: "Applies the body-text preset (160% leading, justify with last line left, Loose v2 kinsoku, Solid mojikumi, top alignment). Tab stops are cleared."
            },
            roleHeading: {
                ja: "見出し向けの設定をまとめて適用します（行送り120%・左揃え・弱い禁則 v2・ツメ組み・上揃え）。タブ設定は削除します。",
                en: "Applies the heading preset (120% leading, left, Loose v2 kinsoku, Tight mojikumi, top alignment). Tab stops are cleared."
            },
            roleMenu: {
                ja: "メニュー向けの設定をまとめて適用します（行送り150%・右揃え・弱い禁則 v2・ツメ組み・上揃え）。各段落に右揃えタブ（リーダー「…」／400pt）を設定します。行揃えを変えると指定が解除され、タブ設定も削除されます。",
                en: "Applies the menu preset (150% leading, right, Loose v2 kinsoku, Tight mojikumi, top alignment) and sets a right-aligned tab with a leader at 400pt on every paragraph. Changing the justification drops this role and clears the tab stops."
            },
            kinsoku: {
                ja: "段落の禁則処理（なし／強い禁則／弱い禁則など）をまとめて適用します。",
                en: "Applies a kinsoku (line-break) set to the paragraphs (None / Strict / Loose and so on)."
            },
            mojikumi: {
                ja: "文字組みアキ量設定を全段落に適用します。設定が混在しているときは選択が空になり、そのままなら変更しません。",
                en: "Applies a mojikumi spacing set to every paragraph. With mixed settings the menu opens empty and nothing is changed."
            },
            linkIndents: { ja: "左インデントの値を右にも適用します。", en: "Applies the left indent value to the right as well." },
            offset: {
                ja: "エリア内文字オプションの「間隔」（テキストと枠のあいだ）を設定します。［テキストを分離...］では、この分だけ囲み罫を外側に広げます。",
                en: "Sets the inset spacing between the text and the frame (Area Type Options). Separate text... grows the rectangle outward by the same amount."
            },
            separateText: {
                ja: "エリア内文字を、囲み罫（長方形）とポイント文字に分解します。別ダイアログで枠の処理を選びます。選択したエリア内文字をまとめて処理します。",
                en: "Breaks Area Type apart into a rectangle plus point text. A separate dialog picks how the frame is handled. Every selected Area Type frame is processed at once."
            },
            resolveOverset: {
                ja: "あふれている文字が隠れたまま分離されないよう、枠の高さを全文が収まる高さまで広げてから分解します。あふれていないフレームには何もしません。",
                en: "Grows the frame until the whole text fits before separating, so no overset character is left hidden. Frames without overset are untouched."
            },
            forceLineBreaks: {
                ja: "エリア内文字で折り返していた位置に改行を入れ、見た目の行分けをポイント文字でも保ちます。OFFにすると段落ごとに1行につながります。",
                en: "Inserts a return where the Area Type wrapped, so the point text keeps the same line breaks. When off, each paragraph becomes a single long line."
            }
        },
        alert: {
            selectText: {
                ja: "ポイント文字・パス上文字・エリア内文字・図形を選択してください。",
                en: "Please select point text, path text, area text, or a shape."
            },
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            lineBreakNotSupported: {
                ja: "改行を含むテキストには対応していません。改行のないテキストを選択してください。",
                en: "Text containing line breaks is not supported. Select text without line breaks."
            },
            convertFailed: {
                ja: "変換できる対象がありませんでした。ポイント文字や閉じたパスを選択してください。",
                en: "Nothing could be converted. Select point text or a closed path."
            },
            separateFailed: {
                ja: "分離できるエリア内文字がありませんでした。ロックや非表示になっていないか確認してください。",
                en: "No Area Type could be separated. Check whether the frames are locked or hidden."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す（{slash} は / に展開）
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        var resolvedText = labelNode[uiLang] || labelNode.en || labelPath;
        return resolvedText.replace(/\{slash\}/g, "/");
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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
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
    // パネルレイアウト / Panel layout
    // =========================================

    /* 禁則・文字組みポップアップの幅（英語は語が長いので広め）/ Width of the kinsoku and mojikumi popups (wider in English) */
    var JP_DROPDOWN_WIDTH = (uiLang === "ja") ? 140 : 190;

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    var DIALOG_MARGINS = 20;               /* ダイアログの余白 / dialog margins */
    var ICON_BUTTON_SIZE = 26;             /* 行揃え・配置のアイコンボタンの一辺 / side of the icon buttons */
    var LEADING_LABEL_WIDTH = 48;          /* 行送りの行ラベルの幅 / width of the leading row labels */
    var FRAME_LABEL_WIDTH = 28;            /* 幅・高さの行ラベルの幅 / width of the width and height labels */
    var SMALL_FIELD_CHARACTERS = 4;        /* 数値欄の文字数 / width of a number field */
    var SIZE_FIELD_CHARACTERS = 5;         /* 幅・高さ欄の文字数 / width of the width and height fields */
    var ROW_GAP_HEIGHT = 5;                /* 行のあいだの空き / gap between rows */

    /**
     * パネルの共通設定を当てる
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 子どうしの間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function applyPanelLayout(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）
     * size を変えても location は動かないので、上下から均等に詰めるには位置も半分ずらす
     * Changing size keeps the location, so the position shifts by half to trim top and bottom evenly
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPx - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPx) {
        /* レイアウト前など、size・location を変えられないときは何もしない / Leave it as is when size or location cannot be changed yet */
        try {
            var buttonLocation = [targetButton.location[0], targetButton.location[1]];
            targetButton.size = [targetButton.size.width, targetButton.size.height - trimPx];
            targetButton.location = [buttonLocation[0], buttonLocation[1] + trimPx / 2];
        } catch (e) { }
    }

    /**
     * ラジオ群から1つだけを選択状態にする（null ならすべて外す）
     * @param {RadioButton[]} radioButtons - ラジオボタンの配列
     * @param {RadioButton|null} targetRadio - 選択するラジオボタン
     * @returns {void}
     */
    function selectRadio(radioButtons, targetRadio) {
        for (var i = 0; i < radioButtons.length; i++) {
            radioButtons[i].value = (radioButtons[i] === targetRadio);
        }
    }

    /**
     * ラジオ群にクリック時の排他制御を割り当てる（パネルをまたいでも1つだけ選ばれるように）
     * @param {RadioButton[]} radioButtons - ラジオボタンの配列
     * @param {Function} [onSelect] - 選んだあとに呼ぶ処理（引数は選んだラジオボタン）
     * @returns {void}
     */
    function bindExclusiveRadios(radioButtons, onSelect) {
        for (var i = 0; i < radioButtons.length; i++) {
            (function (clickedRadio) {
                clickedRadio.onClick = function () {
                    selectRadio(radioButtons, clickedRadio);
                    if (typeof onSelect === "function") { onSelect(clickedRadio); }
                };
            })(radioButtons[i]);
        }
    }

    // =========================================
    // アイコンボタン（行揃え／テキストの配置）/ Icon buttons (justification and text alignment)
    //   UnifiedTypePalette.jsx（原典は Keep_TextPosition.jsx）から移植
    //   Ported from UnifiedTypePalette.jsx (originally Keep_TextPosition.jsx)
    // =========================================

    /* 行揃えの選択肢（id・ラベル・Justification 値・アイコン種別・ショートカット）
       Justification options (id, label, Justification value, icon type, shortcut key) */
    var JUSTIFY_OPTIONS = [
        { id: "left", labelKey: "iconButton.justifyLeft", value: Justification.LEFT, icon: "left", shortcut: "L" },
        { id: "center", labelKey: "iconButton.justifyCenter", value: Justification.CENTER, icon: "center", shortcut: "C" },
        { id: "right", labelKey: "iconButton.justifyRight", value: Justification.RIGHT, icon: "right", shortcut: "R" },
        { id: "lastLineLeft", labelKey: "iconButton.justifyLastLineLeft", value: Justification.FULLJUSTIFYLASTLINELEFT, icon: "justifyLeft", shortcut: "J" },
        { id: "allLines", labelKey: "iconButton.justifyAllLines", value: Justification.FULLJUSTIFY, icon: "justifyAll", shortcut: "F" }
    ];

    /**
     * 行揃え id から Justification 値を引く
     * @param {string} justifyId - 行揃え id
     * @returns {Justification} Justification 値（見つからなければ LEFT）
     */
    function getJustificationValue(justifyId) {
        for (var i = 0; i < JUSTIFY_OPTIONS.length; i++) {
            if (JUSTIFY_OPTIONS[i].id === justifyId) return JUSTIFY_OPTIONS[i].value;
        }
        return Justification.LEFT;
    }

    /**
     * Justification 値から行揃え id を引く
     * @param {Justification} justificationValue - Justification 値
     * @returns {string} 行揃え id（見つからなければ "left"）
     */
    function getJustificationId(justificationValue) {
        for (var i = 0; i < JUSTIFY_OPTIONS.length; i++) {
            if (JUSTIFY_OPTIONS[i].value === justificationValue) return JUSTIFY_OPTIONS[i].id;
        }
        return "left";
    }

    /* テキストの配置の選択肢（id・ラベル・ダイナミックアクションの値・アイコン種別）
       Text-alignment options (id, label, dynamic-action value, icon type) */
    var ALIGN_OPTIONS = [
        { id: "top", labelKey: "iconButton.alignTop", value: 0, icon: "top" },
        { id: "center", labelKey: "iconButton.alignCenter", value: 1, icon: "center" },
        { id: "bottom", labelKey: "iconButton.alignBottom", value: 2, icon: "bottom" },
        { id: "justify", labelKey: "iconButton.alignJustify", value: 3, icon: "justify" }
    ];

    /**
     * 配置 id からアクションの値を引く
     * @param {string} alignId - 配置 id
     * @returns {number} アクションの値（見つからなければ 0＝上揃え）
     */
    function getAlignmentValue(alignId) {
        for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
            if (ALIGN_OPTIONS[i].id === alignId) return ALIGN_OPTIONS[i].value;
        }
        return 0;
    }

    /**
     * アクションの値から配置 id を引く
     * @param {number|null} alignmentValue - アクションの値
     * @returns {string} 配置 id（見つからなければ "top"）
     */
    function getAlignmentId(alignmentValue) {
        for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
            if (ALIGN_OPTIONS[i].value === alignmentValue) return ALIGN_OPTIONS[i].id;
        }
        return "top";
    }

    /**
     * 環境設定の UI の明るさが明るい側か（0=最暗〜1=最明。0.5 は暗い側に含める）
     * @returns {boolean} 明るい側なら true
     */
    function isLightUI() {
        /* 読めない環境では暗い側として扱う / Treat an unreadable preference as dark */
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /**
     * テーマと選択状態に応じたボタンの配色を返す
     * @param {boolean} isLight - 明るい UI か
     * @param {boolean} isActive - 選択中か
     * @returns {{bg: number[], border: number[]|null, line: number[]}} 背景・枠・罫線の色
     */
    function getIconButtonColors(isLight, isActive) {
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
     * 行揃えアイコンの行ごとの線の長さを返す
     * @param {string} iconType - アイコン種別
     * @param {number} longWidth - 長い線の長さ
     * @param {number} shortWidth - 短い線の長さ
     * @returns {number[]} 4行ぶんの線の長さ
     */
    function getJustifyLineWidths(iconType, longWidth, shortWidth) {
        if (iconType === "justifyAll") return [longWidth, longWidth, longWidth, longWidth];
        if (iconType === "justifyLeft") return [longWidth, longWidth, longWidth, shortWidth];
        return [longWidth, shortWidth, longWidth, shortWidth];
    }

    /**
     * 行揃えアイコンの線の開始 X（アイコン種別ごとに左／中央／右）を返す
     * @param {string} iconType - アイコン種別
     * @param {number} buttonWidth - ボタンの幅
     * @param {number} lineWidth - 線の長さ
     * @returns {number} 開始 X
     */
    function getJustifyLineX(iconType, buttonWidth, lineWidth) {
        var sideMargin = 5;
        if (iconType === "right") return buttonWidth - sideMargin - lineWidth;
        if (iconType === "center") return Math.round((buttonWidth - lineWidth) / 2);
        return sideMargin;
    }

    /**
     * 行揃えアイコンの罫線を描く
     * @param {ScriptUIGraphics} graphics - ボタンの graphics
     * @param {string} iconType - アイコン種別
     * @param {number} buttonWidth - ボタンの幅
     * @param {number[]} lineColor - 線の色
     * @returns {void}
     */
    function drawJustifyIconLines(graphics, iconType, buttonWidth, lineColor) {
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
        var rowYs = [7, 11, 15, 19];
        var lineWidths = getJustifyLineWidths(iconType, 15, 10);
        for (var i = 0; i < rowYs.length; i++) {
            var lineWidth = lineWidths[i];
            var lineStartX = getJustifyLineX(iconType, buttonWidth, lineWidth);
            graphics.newPath();
            graphics.moveTo(lineStartX, rowYs[i]);
            graphics.lineTo(lineStartX + lineWidth, rowYs[i]);
            graphics.strokePath(linePen);
        }
    }

    /**
     * テキストの配置アイコンの行の Y 位置（上寄せ／中央／下寄せ／均等）を返す
     * @param {string} iconType - アイコン種別
     * @returns {number[]} 3行ぶんの Y 位置
     */
    function getAlignRowYs(iconType) {
        if (iconType === "center") return [9, 13, 17];
        if (iconType === "bottom") return [11, 15, 19];
        if (iconType === "justify") return [7, 13, 19];
        return [7, 11, 15];
    }

    /**
     * テキストの配置アイコンの罫線を描く（幅は3本とも同じで左右中央）
     * @param {ScriptUIGraphics} graphics - ボタンの graphics
     * @param {string} iconType - アイコン種別
     * @param {number} buttonWidth - ボタンの幅
     * @param {number[]} lineColor - 線の色
     * @returns {void}
     */
    function drawAlignIconLines(graphics, iconType, buttonWidth, lineColor) {
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1.2);
        var lineWidth = 14;
        var lineStartX = Math.round((buttonWidth - lineWidth) / 2);
        var rowYs = getAlignRowYs(iconType);
        for (var i = 0; i < rowYs.length; i++) {
            graphics.newPath();
            graphics.moveTo(lineStartX, rowYs[i]);
            graphics.lineTo(lineStartX + lineWidth, rowYs[i]);
            graphics.strokePath(linePen);
        }
    }

    /**
     * アイコンボタンの背景を描く（描けなければ OS 標準のボタンを描く）
     * @param {ScriptUIGraphics} graphics - ボタンの graphics
     * @param {number} buttonWidth - ボタンの幅
     * @param {number} buttonHeight - ボタンの高さ
     * @param {Object} colors - getIconButtonColors() の戻り値
     * @returns {void}
     */
    function drawIconButtonBackground(graphics, buttonWidth, buttonHeight, colors) {
        try {
            graphics.rectPath(0, 0, buttonWidth, buttonHeight);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, colors.bg));
            if (colors.border) {
                graphics.rectPath(0, 0, buttonWidth, buttonHeight);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, colors.border, 1));
            }
        } catch (e) {
            try { graphics.drawOSControl(); } catch (e0) { }
        }
    }

    /**
     * 行揃えボタンを描く（テーマ・選択状態に対応）
     * @param {Button} iconButton - 描くボタン
     * @param {boolean} isActive - 選択中か
     * @param {boolean} isLight - 明るい UI か
     * @returns {void}
     */
    function drawJustifyIcon(iconButton, isActive, isLight) {
        var colors = getIconButtonColors(isLight, isActive);
        drawIconButtonBackground(iconButton.graphics, iconButton.size[0], iconButton.size[1], colors);
        drawJustifyIconLines(iconButton.graphics, iconButton.iconType, iconButton.size[0], colors.line);
    }

    /**
     * テキストの配置ボタンを描く（テーマ・選択状態に対応）
     * @param {Button} iconButton - 描くボタン
     * @param {boolean} isActive - 選択中か
     * @param {boolean} isLight - 明るい UI か
     * @returns {void}
     */
    function drawAlignIcon(iconButton, isActive, isLight) {
        var colors = getIconButtonColors(isLight, isActive);
        drawIconButtonBackground(iconButton.graphics, iconButton.size[0], iconButton.size[1], colors);
        drawAlignIconLines(iconButton.graphics, iconButton.iconType, iconButton.size[0], colors.line);
    }

    // =========================================
    // 日本語の組版（禁則・文字組みアキ量設定）/ Japanese composition (kinsoku and mojikumi)
    //   UnifiedTypePalette.jsx から移植 / Ported from UnifiedTypePalette.jsx
    // =========================================

    /* 禁則の選択肢（id は paragraphAttributes.kinsoku に渡す値）
       Kinsoku choices (id is the value passed to paragraphAttributes.kinsoku) */
    var KINSOKU_CHOICES = [
        { id: "None", labelKey: "dropdown.kinsokuNone" },
        { id: "Hard", labelKey: "dropdown.kinsokuHard" },
        { id: "Soft", labelKey: "dropdown.kinsokuSoft" },
        { id: "Soft_v2", labelKey: "dropdown.kinsokuSoftV2" }
    ];

    /* 文字組みアキ量設定の選択肢（index は mojikumiSet の添字。-1 は「なし」）
       Mojikumi choices (index is the mojikumiSet index; -1 means "None") */
    var MOJIKUMI_CHOICES = [
        { index: -1, labelKey: "dropdown.mojikumiNone" },
        { index: 0, labelKey: "dropdown.mojikumiLineEndFullHalf" },
        { index: 1, labelKey: "dropdown.mojikumiPunctHalf" },
        { index: 2, labelKey: "dropdown.mojikumiLineEndHalf" },
        { index: 3, labelKey: "dropdown.mojikumiLineEndFull" },
        { index: 4, labelKey: "dropdown.mojikumiPunctFull" },
        { index: 5, labelKey: "dropdown.mojikumiTight" },
        { index: 6, labelKey: "dropdown.mojikumiSolid" }
    ];

    /**
     * 選択肢テーブルからドロップダウン用のラベル配列を作る
     * @param {Object[]} choices - 選択肢テーブル（labelKey を持つ）
     * @returns {string[]} 表示言語のラベル
     */
    function buildChoiceLabels(choices) {
        var choiceLabels = [];
        for (var i = 0; i < choices.length; i++) { choiceLabels.push(getLabel(choices[i].labelKey)); }
        return choiceLabels;
    }

    /**
     * 選択肢テーブルの値に一致する項目をドロップダウンで選ぶ（無ければ未選択のままにする）
     * @param {DropDownList} targetDropdown - 対象のドロップダウン
     * @param {Object[]} choices - 選択肢テーブル
     * @param {string} keyName - 比べるプロパティ名（"id" / "index"）
     * @param {*} matchValue - 選びたい値
     * @returns {void}
     */
    function selectChoiceByValue(targetDropdown, choices, keyName, matchValue) {
        for (var i = 0; i < choices.length; i++) {
            if (choices[i][keyName] === matchValue) { targetDropdown.selection = i; return; }
        }
        targetDropdown.selection = null;
    }

    /* 「なし」を表す文字組み設定名の候補（UI言語で変わるため）/ Names that mean "none" for mojikumi (they vary by UI language) */
    var MOJIKUMI_NONE_NAMES = ["なし", "None"];

    /**
     * 文字組み設定名が「なし」を指すか
     * @param {string} mojikumiName - 文字組み設定名
     * @returns {boolean} 「なし」なら true
     */
    function isMojikumiNoneName(mojikumiName) {
        for (var i = 0; i < MOJIKUMI_NONE_NAMES.length; i++) {
            if (mojikumiName === MOJIKUMI_NONE_NAMES[i]) return true;
        }
        return false;
    }

    /**
     * 段落の文字組みを「なし」にする（設定名の候補を順に試す）
     * @param {TextRange} paragraph - 対象の段落
     * @returns {void}
     */
    function clearParagraphMojikumi(paragraph) {
        for (var i = 0; i < MOJIKUMI_NONE_NAMES.length; i++) {
            /* UI 言語に合わない名前は例外になるので次の候補へ / A name that does not match the UI language throws, so try the next */
            try {
                paragraph.paragraphAttributes.mojikumi = MOJIKUMI_NONE_NAMES[i];
                return;
            } catch (e) { }
        }
    }

    /**
     * 文字組み設定名から mojikumiSet の添字を引く
     * @param {string} mojikumiName - 文字組み設定名
     * @returns {number} mojikumiSet の添字（見つからなければ -2）
     */
    function findMojikumiIndexByName(mojikumiName) {
        try {
            var mojikumiSets = app.activeDocument.mojikumiSet;
            for (var i = 0; i < mojikumiSets.length; i++) {
                if (mojikumiSets[i].name === mojikumiName) return i;
            }
        } catch (e) { }
        return -2;
    }

    /**
     * 先頭段落の文字組み設定を添字で返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number} mojikumiSet の添字（-1＝なし、-2＝不明・混在）
     */
    function getMojikumiIndex(textFrame) {
        var mojikumiAttr;
        try {
            if (textFrame.paragraphs.length === 0) return -2;
            mojikumiAttr = textFrame.paragraphs[0].paragraphAttributes.mojikumi;
        } catch (e) { return -2; }
        if (mojikumiAttr === undefined || mojikumiAttr === null) return -1;
        if (typeof mojikumiAttr === "string") {
            return (mojikumiAttr === "" || isMojikumiNoneName(mojikumiAttr)) ? -1 : findMojikumiIndexByName(mojikumiAttr);
        }
        var mojikumiName;
        /* 文字組みアキ量設定の name は読めないことがある / The name of a mojikumi set can be unreadable */
        try { mojikumiName = (mojikumiAttr.name === undefined || mojikumiAttr.name === null) ? "" : mojikumiAttr.name; } catch (e0) { return -2; }
        if (mojikumiName === "") return -2;
        if (isMojikumiNoneName(mojikumiName)) return -1;
        return findMojikumiIndexByName(mojikumiName);
    }

    /**
     * 先頭段落の禁則処理を id で返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {string} 禁則の id（未設定・読めないときは "None"）
     */
    function getKinsokuId(textFrame) {
        try {
            if (textFrame.paragraphs.length === 0) return "None";
            var kinsokuValue = textFrame.paragraphs[0].paragraphAttributes.kinsoku;
            if (kinsokuValue === undefined || kinsokuValue === null || kinsokuValue === "") return "None";
            return String(kinsokuValue);
        } catch (e) {
            return "None";
        }
    }

    /**
     * 全段落に文字組みアキ量設定を適用する（-2＝不明のときは触らない）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} mojikumiIndex - mojikumiSet の添字（-1＝なし）
     * @returns {void}
     */
    function applyMojikumiToTextFrame(textFrame, mojikumiIndex) {
        if (mojikumiIndex === undefined || mojikumiIndex === null || mojikumiIndex === -2) return;
        if (mojikumiIndex === -1) {
            for (var i = 0; i < textFrame.paragraphs.length; i++) { clearParagraphMojikumi(textFrame.paragraphs[i]); }
            return;
        }
        var mojikumiValue;
        try { mojikumiValue = app.activeDocument.mojikumiSet[mojikumiIndex]; } catch (e) { return; }
        for (var j = 0; j < textFrame.paragraphs.length; j++) {
            try { textFrame.paragraphs[j].paragraphAttributes.mojikumi = mojikumiValue; } catch (e0) { }
        }
    }

    /**
     * 禁則処理を適用する（「なし」も含めて設定できる）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {string|null} kinsokuId - 禁則の id（null なら触らない）
     * @returns {void}
     */
    function applyKinsokuToTextFrame(textFrame, kinsokuId) {
        if (!kinsokuId) return;
        try { textFrame.textRange.paragraphAttributes.kinsoku = kinsokuId; } catch (e) { }
    }

    // =========================================
    // パス上文字の前処理 / Path text preprocessing
    // =========================================

    /**
     * パス上文字をポイント文字に置き換える（変換ダイアログの前処理）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame[]} pathTextFrames - パス上文字
     * @returns {TextFrame[]} 作成したポイント文字
     */
    function detachPathTextToPointText(doc, pathTextFrames) {
        var createdPointTexts = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) return createdPointTexts;

        /* 新しく作ったテキストだけを選べるように、いったん選択を解除する / Deselect so that only the new text ends up selected */
        doc.selection = null;

        for (var i = pathTextFrames.length - 1; i >= 0; i--) {
            var pathText = pathTextFrames[i];
            if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) continue;
            /* 1件で失敗しても残りを処理する / One failure must not stop the rest */
            try {
                var pointText = replacePathTextWithPointText(doc, pathText);
                if (pointText) { createdPointTexts.push(pointText); }
            } catch (e) { }
        }

        return createdPointTexts;
    }

    /**
     * パス上文字1件をポイント文字に置き換える
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} pathText - パス上文字
     * @returns {TextFrame|null} 作成したポイント文字（パスが無ければ null）
     */
    function replacePathTextWithPointText(doc, pathText) {
        var originalPath = pathText.textPath;
        if (!originalPath) return null;

        var charAttrSnapshots = snapshotCharacterAttributes(pathText);
        var textContents = pathText.contents;
        var justification = pathText.paragraphs.length > 0
            ? pathText.paragraphs[0].paragraphAttributes.justification
            : null;

        var pointText = doc.textFrames.add();

        /* パスの始点にそろえる / Align with the start point of the path */
        if (originalPath.pathPoints.length > 0) {
            var anchorPoint = originalPath.pathPoints[0].anchor;
            pointText.position = [anchorPoint[0], anchorPoint[1]];
        }

        pointText.contents = textContents;
        if (justification !== null && pointText.paragraphs.length > 0) {
            pointText.paragraphs[0].paragraphAttributes.justification = justification;
        }

        /* 既定の線をいったん外し、このあと文字単位で復元する / Clear the default stroke, then restore per character */
        pointText.textRange.characterAttributes.strokeColor = new NoColor();
        pointText.textRange.characterAttributes.strokeWeight = 0;
        restoreCharacterAttributes(pointText, charAttrSnapshots, true);

        pathText.remove(); /* パスも一緒に削除される / the path is removed along with it */
        pointText.selected = true;
        return pointText;
    }

    /**
     * 文字単位の属性を控える
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {Object[]} 1文字ごとの属性
     */
    function snapshotCharacterAttributes(textFrame) {
        var snapshots = [];
        for (var i = 0; i < textFrame.characters.length; i++) {
            var charAttr = textFrame.characters[i].characterAttributes;
            snapshots.push({
                font: charAttr.textFont,
                size: charAttr.size,
                fillColor: charAttr.fillColor,
                strokeColor: charAttr.strokeColor,
                strokeWeight: charAttr.strokeWeight,
                autoLeading: charAttr.autoLeading,
                leading: charAttr.leading,
                baselineShift: charAttr.baselineShift,
                horizontalScale: charAttr.horizontalScale,
                verticalScale: charAttr.verticalScale
            });
        }
        return snapshots;
    }

    /**
     * 控えた文字属性を書き戻す
     * dropTransforms が true のときは、パス上文字で付いたベースライン移動・拡大縮小を捨てる
     * When dropTransforms is true, the baseline shift and scaling baked in by path text are dropped
     * @param {TextFrame} textFrame - 書き戻す先
     * @param {Object[]} snapshots - snapshotCharacterAttributes() の戻り値
     * @param {boolean} dropTransforms - ベースライン移動・拡大縮小を捨てるか
     * @returns {void}
     */
    function restoreCharacterAttributes(textFrame, snapshots, dropTransforms) {
        var copyCount = Math.min(textFrame.characters.length, snapshots.length);
        for (var i = 0; i < copyCount; i++) {
            var destAttrs = textFrame.characters[i].characterAttributes;
            var srcAttrs = snapshots[i];

            /* 未インストールのフォントは失敗しうるので、他の属性と分けて適用する / An uninstalled font can fail, so it is applied apart from the rest */
            try { destAttrs.textFont = srcAttrs.font; } catch (e0) { }
            try {
                destAttrs.size = srcAttrs.size;
                destAttrs.fillColor = srcAttrs.fillColor;
                destAttrs.strokeColor = srcAttrs.strokeColor;
                destAttrs.strokeWeight = (srcAttrs.strokeColor && srcAttrs.strokeColor.typename === "NoColor") ? 0 : srcAttrs.strokeWeight;
                destAttrs.baselineShift = dropTransforms ? 0 : srcAttrs.baselineShift;
                destAttrs.horizontalScale = dropTransforms ? 100 : srcAttrs.horizontalScale;
                destAttrs.verticalScale = dropTransforms ? 100 : srcAttrs.verticalScale;
                destAttrs.autoLeading = srcAttrs.autoLeading;
                if (!srcAttrs.autoLeading) { destAttrs.leading = srcAttrs.leading; }
            } catch (e1) { }
        }
    }

    /**
     * 選択中のパス上文字をポイント文字に置き換え、新しい選択を返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 現在の選択
     * @returns {PageItem[]} 置き換え後の選択（パス上文字が無ければ元の選択）
     */
    function preprocessSelectionForConvertDialog(doc, selectedItems) {
        if (!doc || !selectedItems || !selectedItems.length) return selectedItems;

        /* 削除で参照が無効になる前に振り分ける / Sort before any removal invalidates references */
        var pathTextItems = [], otherItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (!selectedItem) continue;
            if (selectedItem.typename === "TextFrame" && selectedItem.kind === TextType.PATHTEXT) { pathTextItems.push(selectedItem); }
            else { otherItems.push(selectedItem); }
        }
        if (!pathTextItems.length) return selectedItems;

        var pointTextItems = detachPathTextToPointText(doc, pathTextItems);
        if (!pointTextItems.length) return selectedItems;

        var nextSelection = otherItems.concat(pointTextItems);
        doc.selection = nextSelection;
        app.redraw();
        return nextSelection;
    }

    // =========================================
    // 文字あふれとフォントサイズ / Overset and font size
    //   AutoFitTextFrame.jsx 参考 / Based on AutoFitTextFrame.jsx
    // =========================================

    /**
     * overflows プロパティを安全に読む
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean|null} あふれていれば true（読めなければ null）
     */
    function getOverflowState(textFrame) {
        try {
            if (textFrame && typeof textFrame.overflows !== "undefined") return !!textFrame.overflows;
        } catch (e) { }
        return null;
    }

    /**
     * 指定行数までに全文字が収まっているかで文字あふれを判定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} [lineLimit] - 見る行数（省略時は1）
     * @returns {boolean} あふれていれば true
     */
    function isTextOverset(textFrame, lineLimit) {
        if (textFrame.lines.length > 0) {
            var charCount = 0;
            lineLimit = (typeof lineLimit === "undefined" || lineLimit === null) ? 1 : Math.floor(lineLimit);
            if (lineLimit < 1) lineLimit = 1;
            if (lineLimit > textFrame.lines.length) lineLimit = textFrame.lines.length;
            for (var i = 0; i < lineLimit; i++) { charCount += textFrame.lines[i].characters.length; }
            return charCount < textFrame.characters.length;
        }
        return textFrame.characters.length > 0;
    }

    /**
     * テキストフレームが文字あふれしているか（エリア内文字は overflows を優先）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} あふれていれば true
     */
    function isFrameOverset(textFrame) {
        if (textFrame && textFrame.kind === TextType.AREATEXT) {
            var overflowState = getOverflowState(textFrame);
            if (overflowState !== null) return overflowState;
        }
        /* 行・文字を数えられないときはあふれていないものとして扱う / Treat it as not overset when lines and characters cannot be counted */
        try { return isTextOverset(textFrame, textFrame.lines.length || 1); } catch (e) { return false; }
    }

    /**
     * エリア内文字の枠の高さを返す
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {number} 高さ（pt。取れなければ 0）
     */
    function getFramePathHeight(textFrame) {
        try { return textFrame.textPath.height || 0; } catch (e) { return 0; }
    }

    /**
     * 改行を含むか
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {boolean} 改行を含めば true
     */
    function hasLineBreak(textFrame) {
        try { return /[\r\n\u2028\u2029]/.test(textFrame.contents); } catch (e) { return false; }
    }

    /**
     * 代表となるフォントサイズを返す（混在しているときは先頭文字の値）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number} フォントサイズ（pt。取れなければ 0）
     */
    function getRepresentativeFontSize(textFrame) {
        /* サイズが混在していると textRange のサイズは読めない / A mixed size cannot be read from the whole range */
        try {
            var rangeFontSize = textFrame.textRange.characterAttributes.size;
            if (rangeFontSize > 0) return rangeFontSize;
        } catch (e) { }
        try { return textFrame.characters[0].characterAttributes.size || 0; } catch (e0) { }
        return 0;
    }

    /**
     * 固定行送りのとき、行送り÷フォントサイズの比率を返す
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {{ratio: number}|null} 比率（自動行送り・読めないときは null）
     */
    function getLeadingRatioInfo(textFrame) {
        try {
            var charAttrs = textFrame.textRange.characterAttributes;
            if (charAttrs.autoLeading) return null;
            var fontSize = charAttrs.size, leadingPt = charAttrs.leading;
            if (fontSize > 0 && leadingPt > 0) return { ratio: leadingPt / fontSize };
        } catch (e) { }
        return null;
    }

    /**
     * 比率を保ったまま、新しいフォントサイズに行送りを合わせる
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} newSize - 新しいフォントサイズ（pt）
     * @param {{ratio: number}|null} leadingRatioInfo - getLeadingRatioInfo() の戻り値
     * @returns {void}
     */
    function applyProportionalLeading(textFrame, newSize, leadingRatioInfo) {
        if (!leadingRatioInfo) return;
        try { textFrame.textRange.characterAttributes.leading = newSize * leadingRatioInfo.ratio; } catch (e) { }
    }

    // =========================================
    // ダイナミックアクション / Dynamic actions
    //   スクリプト実行時に読み込み、終了時にアンロードする
    //   Loaded at startup, unloaded on exit
    // =========================================

    var ACTION_SET_AUTO_SIZE = "AreaTypeToolkit_AutoSize";
    var ACTION_SET_ALIGNMENT = "AreaTypeToolkit_Alignment";

    /* 自動サイズ調整のアクション名（1=ON, 2=OFF）/ Auto-size action names (1=ON, 2=OFF) */
    var AUTO_SIZE_ACTIONS = [
        { name: "AutoSizeOn", value: 1 },
        { name: "AutoSizeOff", value: 2 }
    ];

    /* テキストの配置のアクション名（添字＝アクションの値：0=上, 1=中央, 2=下, 3=均等）
       Frame-alignment action names (the index is the action value: 0=top, 1=center, 2=bottom, 3=justify) */
    var ALIGNMENT_ACTIONS = [
        { name: "AlignTop", value: 0 },
        { name: "AlignCenter", value: 1 },
        { name: "AlignBottom", value: 2 },
        { name: "AlignJustify", value: 3 }
    ];

    /**
     * 文字列を ASCII 16進に変換する
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
     * アクション名ブロック /name [ <len> <hex> ] を生成する
     * @param {string} actionName - アクション名またはセット名
     * @returns {string} 名前ブロックの文字列
     */
    function buildActionNameBlock(actionName) {
        return "/name [ " + actionName.length + " " + asciiToHex(actionName).toUpperCase() + " ]";
    }

    /**
     * アクションセット定義（.aia 文字列）を組み立てる
     * @param {string} setName - アクションセット名
     * @param {string} internalName - イベントの内部名
     * @param {string} localizedNameHex - イベントの表示名（"<長さ> <16進>"。空なら省く）
     * @param {number} parameterKey - パラメーターのキー
     * @param {Object[]} actionDefinitions - { name, value } の配列（1件が1アクション）
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

    /**
     * .aia 文字列を一時ファイル経由で読み込む（既存があれば先に外す）
     * @param {string} setName - アクションセット名
     * @param {string} aiaString - .aia 形式のアクションセット定義
     * @returns {void}
     */
    function loadActionSet(setName, aiaString) {
        /* 読み込まれていなければ例外になる / Throws when the set is not loaded */
        try { app.unloadAction(setName, ""); } catch (e0) { }
        var tempFile = new File(Folder.temp + "/AreaTypeToolkit_action_" + setName + ".aia");
        tempFile.open("w");
        tempFile.write(aiaString);
        tempFile.close();
        app.loadAction(tempFile);
        try { tempFile.remove(); } catch (e1) { }
    }

    /**
     * 使用するアクションセットを読み込む（スクリプト開始時に1回）
     * @returns {void}
     */
    function loadDynamicActions() {
        loadActionSet(ACTION_SET_AUTO_SIZE, buildActionSetAia(
            ACTION_SET_AUTO_SIZE,
            "adobe_SLOAreaTextDialog",
            "33 e382a8e383aae382a2e58685e69687e5ad97e382aae38397e382b7e383a7e383b3",
            1952539754,
            AUTO_SIZE_ACTIONS
        ));
        loadActionSet(ACTION_SET_ALIGNMENT, buildActionSetAia(
            ACTION_SET_ALIGNMENT,
            "adobe_frameAlignment",
            "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
            1717660782,
            ALIGNMENT_ACTIONS
        ));
    }

    /**
     * 読み込んだアクションセットを破棄する（スクリプト終了時）
     * @returns {void}
     */
    function unloadDynamicActions() {
        try { app.unloadAction(ACTION_SET_AUTO_SIZE, ""); } catch (e) { }
        try { app.unloadAction(ACTION_SET_ALIGNMENT, ""); } catch (e) { }
    }

    /**
     * フレームサイズ：自動サイズ調整を ON にして文字に合わせて広げる
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {void}
     */
    function enableFrameAutoSize(textFrame) {
        app.activeDocument.selection = [textFrame];
        runAutoSizeAction(1);
    }

    /**
     * フレームサイズ：自動サイズ調整を OFF にする
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @returns {void}
     */
    function disableFrameAutoSize(textFrame) {
        app.activeDocument.selection = [textFrame];
        runAutoSizeAction(2);
    }

    /**
     * アクションの値に対応するアクション名を返す
     * @param {Object[]} actionDefinitions - { name, value } の配列
     * @param {number} actionValue - アクションの値
     * @returns {string|null} アクション名（無ければ null）
     */
    function findActionName(actionDefinitions, actionValue) {
        for (var i = 0; i < actionDefinitions.length; i++) {
            if (actionDefinitions[i].value === actionValue) return actionDefinitions[i].name;
        }
        return null;
    }

    /**
     * 自動サイズ調整アクションを実行する（1=ON, 2=OFF）
     * @param {number} autoSizeValue - アクションの値
     * @returns {void}
     */
    function runAutoSizeAction(autoSizeValue) {
        var actionName = findActionName(AUTO_SIZE_ACTIONS, autoSizeValue);
        if (!actionName) return;
        try { app.doScript(actionName, ACTION_SET_AUTO_SIZE, false); } catch (e) { }
    }

    /**
     * テキストの配置アクションを実行する（0=上, 1=中央, 2=下, 3=均等）
     * @param {number} alignmentValue - アクションの値
     * @returns {void}
     */
    function runFrameAlignmentAction(alignmentValue) {
        var actionName = findActionName(ALIGNMENT_ACTIONS, alignmentValue);
        if (!actionName) return;
        try { app.doScript(actionName, ACTION_SET_ALIGNMENT, false); } catch (e) { }
    }

    /**
     * 対象を選択してテキストの配置アクションを実行する（プレビュー中は実行しない）
     * @param {TextFrame} textFrame - 対象のエリア内文字
     * @param {number|null} alignmentValue - アクションの値（null なら触らない）
     * @param {boolean} forPreview - プレビュー中か
     * @returns {void}
     */
    function applyAreaTextFrameAlignment(textFrame, alignmentValue, forPreview) {
        /* app.doScript はプレビュー中にダイアログから呼ぶと不安定なためスキップ / app.doScript is unstable from a dialog during preview */
        if (forPreview) return;
        /* 配置未指定（ユーザーが触っていない）なら、いまの配置をそのままにする
           No alignment requested (the user never touched it), so the current placement is left as is */
        if (alignmentValue === null || alignmentValue === undefined) return;
        /* 選択を付け替えられないときは何もしない / Do nothing when the selection cannot be switched */
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [textFrame];
            app.redraw(); /* Illustrator に選択状態を確定させる / Let Illustrator commit the selection */
            runFrameAlignmentAction(alignmentValue);
        } catch (e) { }
    }

    // =========================================
    // フォントサイズの自動調整 / Font-size fitting
    // =========================================

    /**
     * フォントサイズとそれに比例した行送りを設定する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} fontSize - フォントサイズ（pt）
     * @param {{ratio: number}|null} leadingRatioInfo - getLeadingRatioInfo() の戻り値
     * @returns {void}
     */
    function setFontSizeWithLeading(textFrame, fontSize, leadingRatioInfo) {
        textFrame.textRange.characterAttributes.size = fontSize;
        applyProportionalLeading(textFrame, fontSize, leadingRatioInfo);
    }

    /**
     * あふれが消えるまで縮小する（二分探索：最大 40 回で、あふれない最大のサイズを探す）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function shrinkFontToFit(textFrame) {
        if (textFrame.characters.length <= 0 || !isFrameOverset(textFrame)) return;
        var upperSize = getRepresentativeFontSize(textFrame);
        if (upperSize <= 0) return;
        var leadingRatioInfo = getLeadingRatioInfo(textFrame);
        var lowerSize = 0.1;

        /* 最小サイズでも収まらないなら、極小のまま残さず元のサイズに戻す / If even the minimum oversets, put the original size back */
        setFontSizeWithLeading(textFrame, lowerSize, leadingRatioInfo);
        if (isFrameOverset(textFrame)) {
            setFontSizeWithLeading(textFrame, upperSize, leadingRatioInfo);
            return;
        }

        for (var attempt = 0; attempt < 40; attempt++) {
            var midSize = (lowerSize + upperSize) / 2;
            setFontSizeWithLeading(textFrame, midSize, leadingRatioInfo);
            if (isFrameOverset(textFrame)) {
                upperSize = midSize;
            } else {
                lowerSize = midSize;
            }
            if (upperSize - lowerSize < 0.1) break;
        }

        /* あふれない側（lowerSize）に確定 / Settle on the side that does not overset */
        setFontSizeWithLeading(textFrame, lowerSize, leadingRatioInfo);
    }

    /**
     * 枠にフィットさせる：あふれるまで拡大してから縮小する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function fitFontSizeToFrame(textFrame) {
        if (textFrame.characters.length <= 0) return;
        var originalSize = getRepresentativeFontSize(textFrame);
        if (originalSize <= 0) return;
        var leadingRatioInfo = getLeadingRatioInfo(textFrame);

        /**
         * 元のフォントサイズに戻す
         * @returns {void}
         */
        function restoreOriginalSize() {
            try { setFontSizeWithLeading(textFrame, originalSize, leadingRatioInfo); } catch (e) { }
        }

        /* 自動サイズ調整がONだと枠が文字に追従してあふれないため、拡大が止まらない。
           枠の高さが変わったらそれと判断して中止する。
           With auto-size on the frame follows the text and never oversets, so the growth never stops.
           A change in the frame height means that is happening, so bail out. */
        var initialFrameHeight = getFramePathHeight(textFrame);

        /* あふれが出るまで 2 倍ずつ拡大 / Double the size until it oversets */
        if (!isFrameOverset(textFrame)) {
            var trialSize = originalSize, guardCount = 0;
            while (!isFrameOverset(textFrame) && guardCount < 25) {
                guardCount++;
                trialSize = trialSize * 2;
                if (trialSize > 100000) break;
                try { setFontSizeWithLeading(textFrame, trialSize, leadingRatioInfo); } catch (e) { break; }
                if (Math.abs(getFramePathHeight(textFrame) - initialFrameHeight) > 0.01) {
                    restoreOriginalSize();
                    return;
                }
            }
        }

        /* あふれが出なければ元に戻して終了 / Put it back and stop when it never oversets */
        if (!isFrameOverset(textFrame)) {
            restoreOriginalSize();
            return;
        }

        shrinkFontToFit(textFrame);
    }

    /**
     * ↑↓キーで値を増減する（Shift：10 の倍数へ、Option：±0.1）
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

    // =========================================
    // タブ設定 / Tab stops
    // =========================================

    /**
     * メニュー用のリーダー罫：各段落に右揃えタブ（リーダー「…」400pt）を設定し、右揃えにする
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function applyLeaderTabStops(textFrame) {
        try {
            var tabStop = new TabStopInfo();
            tabStop.position = 400;
            tabStop.alignment = TabStopAlignment.Right;
            tabStop.leader = "…";
            for (var i = 0; i < textFrame.paragraphs.length; i++) {
                var paraAttrs = textFrame.paragraphs[i].paragraphAttributes;
                paraAttrs.tabStops = [tabStop];
                paraAttrs.justification = Justification.RIGHT;
            }
        } catch (e) { }
    }

    /**
     * 段落のタブ設定を削除する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {void}
     */
    function clearTabStops(textFrame) {
        try {
            for (var i = 0; i < textFrame.paragraphs.length; i++) {
                textFrame.paragraphs[i].paragraphAttributes.tabStops = [];
            }
        } catch (e) { }
    }

    // ============================================================
    // 変換ダイアログ：ポイント文字 → エリア内文字
    // Convert dialog: point text → Area Type
    // ============================================================

    /**
     * 選択にどの種類のオブジェクトが含まれるかを調べる
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {{hasPointText: boolean, hasAreaText: boolean, hasPath: boolean}} ポイント文字（パス上文字を含む）・エリア内文字・パスの有無
     */
    function detectSelectionKinds(selectedItems) {
        var selectionKinds = { hasPointText: false, hasAreaText: false, hasPath: false };
        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (selectedItem.typename === "TextFrame") {
                if (selectedItem.kind === TextType.POINTTEXT || selectedItem.kind === TextType.PATHTEXT) selectionKinds.hasPointText = true;
                if (selectedItem.kind === TextType.AREATEXT) selectionKinds.hasAreaText = true;
            }
            if (selectedItem.typename === "PathItem" || selectedItem.typename === "CompoundPathItem") {
                selectionKinds.hasPath = true;
            }
        }
        return selectionKinds;
    }

    /**
     * 選択内容に応じて作成方法のラジオを誘導する（使えない方法をディムし、既定を選ぶ）
     * @param {Object} methodRadios - simple / button / useShape / useShapeDummy のラジオボタン
     * @param {Object} selectionKinds - detectSelectionKinds() の戻り値
     * @returns {void}
     */
    function applyConvertMethodGuidance(methodRadios, selectionKinds) {
        var hasPointText = selectionKinds.hasPointText, hasPath = selectionKinds.hasPath;
        if (hasPointText && hasPath) {
            /* テキスト＋図形 → シンプル・ボタン風をディムし、「選択した図形に流し込む」を選ぶ（ダミーテキストも選べる）
               Text plus shape: dim Simple and Button style and pick "Pour into selected shape" (dummy text stays available) */
            methodRadios.simple.enabled = false;
            methodRadios.button.enabled = false;
            methodRadios.useShape.value = true;
        } else if (hasPointText && !hasPath) {
            /* ポイント文字のみ → 図形を使う方法をディム / Point text only: dim the shape-based methods */
            methodRadios.useShape.enabled = false;
            methodRadios.useShapeDummy.enabled = false;
        } else if (!hasPointText && hasPath) {
            /* 図形のみ → 「選択した図形にダミーテキスト」を選び、残りをディム / Shapes only: pick the dummy-text method and dim the rest */
            methodRadios.simple.enabled = false;
            methodRadios.button.enabled = false;
            methodRadios.useShape.enabled = false;
            methodRadios.useShapeDummy.enabled = true;
            methodRadios.useShapeDummy.value = true;
        }
    }

    /**
     * 変換ダイアログを表示し、変換できたら続けて調整ダイアログを開く
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 起動時の選択
     * @returns {void}
     */
    function showConvertDialog(doc, selectedItems) {
        var convertDialog = new Window("dialog", getLabel("dialog.convertTitle") + " " + SCRIPT_VERSION);
        convertDialog.alignChildren = "fill";
        convertDialog.margins = DIALOG_MARGINS;

        /* 作成方法のラジオボタン / Creation-method radios */
        var createMethodPanel = convertDialog.add("panel", undefined, getLabel("panel.createMethod"));
        applyPanelLayout(createMethodPanel);

        var radStyleSimple = createMethodPanel.add("radiobutton", undefined, getLabel("radio.styleSimple"));
        var radStyleButton = createMethodPanel.add("radiobutton", undefined, getLabel("radio.styleButton"));
        var radUseShape = createMethodPanel.add("radiobutton", undefined, getLabel("radio.useShape"));
        var radUseShapeDummy = createMethodPanel.add("radiobutton", undefined, getLabel("radio.useShapeDummy"));
        radStyleSimple.value = true;
        radStyleSimple.helpTip = getLabel("tooltip.styleSimple");
        radStyleButton.helpTip = getLabel("tooltip.styleButton");
        radUseShape.helpTip = getLabel("tooltip.useShape");
        radUseShapeDummy.helpTip = getLabel("tooltip.useShapeDummy");

        /* ボタンエリア：中央に［キャンセル］［変換］/ Button area: Cancel and Convert, centered */
        var btnRowGroup = convertDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["center", "center"];
        btnRowGroup.spacing = 10;
        var btnCancelConvert = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnConvert = btnRowGroup.add("button", undefined, getLabel("button.convert"), { name: "ok" });

        applyConvertMethodGuidance({
            simple: radStyleSimple,
            button: radStyleButton,
            useShape: radUseShape,
            useShapeDummy: radUseShapeDummy
        }, detectSelectionKinds(selectedItems));

        bindExclusiveRadios([radStyleSimple, radStyleButton, radUseShape, radUseShapeDummy]);

        /* 変換後に調整ダイアログを開くための控え / What the adjust dialog needs after the conversion */
        var convertedFrames = [];
        var shouldOpenAdjustDialog = false;
        var initialAlignmentValue = null; /* 0=上, 1=中央, 2=下, 3=均等 / 0=top, 1=center, 2=bottom, 3=justify */

        btnConvert.onClick = function () {
            var currentSelection = app.activeDocument.selection;
            if (!currentSelection || currentSelection.length === 0) { return; }
            /* パス上文字の置き換えは、変換を確定してから行う（キャンセル時に元のパス上文字を壊さないため）
               Path text is swapped out only once the conversion is confirmed, so Cancel leaves it intact */
            currentSelection = preprocessSelectionForConvertDialog(doc, currentSelection);
            var createdFrames = convertSelectionToAreaText(doc, currentSelection,
                radStyleSimple.value, radStyleButton.value, !!radUseShape.value, !!radUseShapeDummy.value);
            if (createdFrames.length > 0) {
                convertedFrames = createdFrames;
                /* ボタン風で変換したときは、調整ダイアログのテキストの配置を中央で開く
                   After a Button-style conversion the adjust dialog opens with the placement centered */
                initialAlignmentValue = radStyleButton.value ? 1 : null;
                shouldOpenAdjustDialog = true;
                convertDialog.close(1);
            } else {
                alert(getLabel("alert.convertFailed"));
            }
        };

        btnCancelConvert.onClick = function () { convertDialog.close(0); };

        convertDialog.show();

        /* show() はブロッキング。閉じた後に調整ダイアログを開く / show() blocks, so the adjust dialog opens once this one is closed */
        if (shouldOpenAdjustDialog && convertedFrames.length > 0) {
            showAdjustDialog(doc, convertedFrames[0], convertedFrames, initialAlignmentValue);
        }
    }

    // ============================================================
    // エリア内文字への変換 / Conversion to Area Type
    // ============================================================

    /* フレームの起こし方（シンプル＝実寸＋1pt、ボタン風＝幅1.2倍・高さ1.8倍で中央）/ Frame styles (simple: +1pt; button: 1.2x / 1.8x, centered) */
    var AREA_TEXT_STYLE_SIMPLE = { widthScale: 1, heightScale: 1, widthPadPt: 1, heightPadPt: 1, centerOnText: false, centerText: false };
    var AREA_TEXT_STYLE_BUTTON = { widthScale: 1.2, heightScale: 1.8, widthPadPt: 0, heightPadPt: 0, centerOnText: true, centerText: true };

    /**
     * 候補名から使用可能なフォントを返す（無ければ先頭のフォント）
     * @param {string[]} candidateFontNames - フォント名の候補
     * @returns {TextFont|null} 見つかったフォント
     */
    function findAvailableTextFont(candidateFontNames) {
        for (var i = 0; i < candidateFontNames.length; i++) {
            try {
                var candidateFont = app.textFonts.getByName(candidateFontNames[i]);
                if (candidateFont) return candidateFont;
            } catch (e) { /* 無いフォント名は例外になる / getByName throws for a missing font */ }
        }
        return (app.textFonts.length > 0) ? app.textFonts[0] : null;
    }

    /**
     * エリア内文字か
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} エリア内文字なら true
     */
    function isAreaTextFrame(pageItem) {
        return !!pageItem && pageItem.typename === "TextFrame" && pageItem.kind === TextType.AREATEXT;
    }

    /**
     * 閉じたパスを取り出す（複合パスは先頭のパスを見る）
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {PathItem|null} 閉じたパス（無ければ null）
     */
    function getClosedPathItem(pageItem) {
        if (!pageItem) return null;
        if (pageItem.typename === "PathItem") {
            return pageItem.closed ? pageItem : null;
        }
        if (pageItem.typename === "CompoundPathItem" && pageItem.pathItems.length > 0) {
            var firstPath = pageItem.pathItems[0];
            return (firstPath && firstPath.closed) ? firstPath : null;
        }
        return null;
    }

    /**
     * 閉じたパスを塗り・線なしのエリア内文字フレームにする
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} pathItem - 閉じたパス
     * @returns {TextFrame} 作成したエリア内文字
     */
    function createAreaTextFromPath(doc, pathItem) {
        pathItem.filled = false;
        pathItem.stroked = false;
        return doc.textFrames.areaText(pathItem);
    }

    /**
     * 中身が無くなった複合パスの殻を片付ける
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {void}
     */
    function removeEmptyCompoundPath(pageItem) {
        /* エリア内文字にしたあとの複合パスは参照が無効になっていることがある / The compound path may already be invalid after the conversion */
        try {
            if (pageItem && pageItem.typename === "CompoundPathItem" && pageItem.pathItems.length === 0) { pageItem.remove(); }
        } catch (e) { }
    }

    /**
     * 先頭文字のフォントとサイズを読み取る
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {{font: TextFont|null, size: number}} フォントとサイズ（読めなければ null と 0）
     */
    function getFontAndSize(textFrame) {
        try {
            var charAttrs = textFrame.textRange.characterAttributes;
            return { font: charAttrs.textFont, size: charAttrs.size };
        } catch (e) {
            return { font: null, size: 0 };
        }
    }

    /**
     * フォントとサイズを適用する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {{font: TextFont|null, size: number}} fontAndSize - 適用するフォントとサイズ
     * @returns {void}
     */
    function applyFontAndSize(textFrame, fontAndSize) {
        if (!fontAndSize) return;
        var charAttrs = textFrame.textRange.characterAttributes;
        if (fontAndSize.size > 0) { charAttrs.size = fontAndSize.size; }
        /* 未インストールのフォントは適用に失敗しうる / An uninstalled font can fail to apply */
        if (fontAndSize.font) { try { charAttrs.textFont = fontAndSize.font; } catch (e) { } }
    }

    /**
     * 段落に自動行送りを設定する（Illustrator 上は「自動」表示になり、行送りがフォントサイズに追従する）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {number} leadingPercent - 自動行送り量（％。0 以下・NaN なら何もしない）
     * @returns {void}
     */
    function applyAutoLeading(textFrame, leadingPercent) {
        if (!(leadingPercent > 0)) return;
        for (var i = 0; i < textFrame.paragraphs.length; i++) {
            try {
                textFrame.paragraphs[i].paragraphAttributes.autoLeadingAmount = leadingPercent;
                textFrame.paragraphs[i].characterAttributes.autoLeading = true;
            } catch (e) { }
        }
    }

    /**
     * 自動行送りが有効ならその量（％）を返す。固定の行送りなら 0 を返し、欄を空のままにする
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {number} 自動行送り量（％）
     */
    function getAutoLeadingPercent(textFrame) {
        try {
            if (!textFrame.textRange.characterAttributes.autoLeading) return 0;
            if (textFrame.paragraphs.length > 0) {
                return textFrame.paragraphs[0].paragraphAttributes.autoLeadingAmount || 0;
            }
        } catch (e) { }
        return 0;
    }

    /**
     * geometricBounds の中心座標を返す
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {number[]} [x, y]
     */
    function getBoundsCenter(pageItem) {
        var bounds = pageItem.geometricBounds; /* [left, top, right, bottom] */
        return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
    }

    /**
     * テキストの中心を含む図形、無ければ中心が最も近い図形を返す（未使用のものだけ）
     * @param {TextFrame} sourceText - 流し込むテキスト
     * @param {PageItem[]} shapeItems - 候補の図形
     * @param {PageItem[]} usedShapes - 使用済みの図形
     * @returns {PageItem|null} 流し込み先の図形
     */
    function findShapeForText(sourceText, shapeItems, usedShapes) {
        var textCenter = getBoundsCenter(sourceText);
        var nearestShape = null, nearestDistance = -1;

        for (var i = 0; i < shapeItems.length; i++) {
            var shapeItem = shapeItems[i];

            var isUsed = false;
            for (var j = 0; j < usedShapes.length; j++) {
                if (usedShapes[j] === shapeItem) { isUsed = true; break; }
            }
            if (isUsed) continue;

            var shapeBounds = shapeItem.geometricBounds;
            if (textCenter[0] >= shapeBounds[0] && textCenter[0] <= shapeBounds[2] &&
                textCenter[1] <= shapeBounds[1] && textCenter[1] >= shapeBounds[3]) {
                return shapeItem;
            }

            var shapeCenter = getBoundsCenter(shapeItem);
            var dx = shapeCenter[0] - textCenter[0], dy = shapeCenter[1] - textCenter[1];
            var distance = dx * dx + dy * dy;
            if (nearestDistance < 0 || distance < nearestDistance) {
                nearestDistance = distance;
                nearestShape = shapeItem;
            }
        }
        return nearestShape;
    }

    /**
     * 選択内容を指定の方法でエリア内文字に変換し、作成したフレームを選択して返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 変換する選択
     * @param {boolean} useSimple - シンプル
     * @param {boolean} useButton - ボタン風
     * @param {boolean} useShape - 選択した図形に流し込む
     * @param {boolean} useShapeDummy - 選択した図形にダミーテキスト
     * @returns {TextFrame[]} 作成したエリア内文字
     */
    function convertSelectionToAreaText(doc, selectedItems, useSimple, useButton, useShape, useShapeDummy) {
        var createdFrames = [];

        if (useShapeDummy) { createdFrames = fillShapesWithDummyText(doc, selectedItems); }
        else if (useShape) { createdFrames = pourTextsIntoShapes(doc, selectedItems); }
        else if (useSimple) { createdFrames = convertPointTexts(doc, selectedItems, AREA_TEXT_STYLE_SIMPLE); }
        else if (useButton) { createdFrames = convertPointTexts(doc, selectedItems, AREA_TEXT_STYLE_BUTTON); }

        if (createdFrames.length > 0) {
            app.activeDocument.selection = createdFrames;
            app.redraw();
        }
        return createdFrames;
    }

    /**
     * 選択した閉じたパスをエリア内文字にしてダミー文字を流し込む
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @returns {TextFrame[]} 作成したエリア内文字
     */
    function fillShapesWithDummyText(doc, selectedItems) {
        var createdFrames = [];
        var dummyText = (uiLang === "ja") ? DUMMY_TEXT_JA : DUMMY_TEXT_EN;
        var dummyFont = findAvailableTextFont((uiLang === "ja") ? DUMMY_FONT_JA : DUMMY_FONT_EN);

        for (var i = 0; i < selectedItems.length; i++) {
            var sourceShape = selectedItems[i];
            var closedPath = getClosedPathItem(sourceShape);
            if (!closedPath) continue;
            /* 種類によってはエリア内文字にできないので、失敗したものは飛ばす / Some shapes cannot become Area Type, so failures are skipped */
            try {
                var dummyFrame = createAreaTextFromPath(doc, closedPath);
                dummyFrame.contents = dummyText;
                applyFontAndSize(dummyFrame, { font: dummyFont, size: DUMMY_FONT_SIZE });
                createdFrames.push(dummyFrame);
                /* 複合パスは先頭のパスだけを枠にするので、空になった殻を残さない / Drop the compound shell left empty */
                removeEmptyCompoundPath(sourceShape);
            } catch (e) { }
        }
        return createdFrames;
    }

    /**
     * 選択テキストを、対応する閉じたパスのエリア内文字に流し込む
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト（テキストと図形）
     * @returns {TextFrame[]} 作成したエリア内文字
     */
    function pourTextsIntoShapes(doc, selectedItems) {
        var createdFrames = [];
        var sourceTexts = [], shapeItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            /* 複合パスも図形として受け付ける（変換ダイアログのラジオ判定と合わせる）
               Compound paths count as shapes too, matching how the convert dialog enables the radios */
            if (selectedItems[i].typename === "TextFrame") { sourceTexts.push(selectedItems[i]); }
            else if (getClosedPathItem(selectedItems[i])) { shapeItems.push(selectedItems[i]); }
        }

        var usedShapes = [];
        for (var textIndex = 0; textIndex < sourceTexts.length; textIndex++) {
            var sourceText = sourceTexts[textIndex];
            var targetShape = findShapeForText(sourceText, shapeItems, usedShapes);
            if (!targetShape) continue;
            usedShapes.push(targetShape);
            /* 1組で失敗しても残りを処理する / One failing pair must not stop the rest */
            try {
                var sourceContents = sourceText.contents;
                var fontAndSize = getFontAndSize(sourceText);
                /* 複合パスごと複製してから、その中の閉じたパスを枠にする
                   Duplicate the whole item, then use the closed path inside the copy as the frame */
                var duplicatedShape = targetShape.duplicate();
                var areaFrame = createAreaTextFromPath(doc, getClosedPathItem(duplicatedShape));
                areaFrame.contents = sourceContents;
                applyFontAndSize(areaFrame, fontAndSize);
                removeEmptyCompoundPath(duplicatedShape);
                sourceText.remove();
                targetShape.remove();
                createdFrames.push(areaFrame);
            } catch (e) { }
        }
        return createdFrames;
    }

    /**
     * ポイント文字を、その大きさから起こした長方形のエリア内文字に置き換える
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択オブジェクト
     * @param {Object} frameStyle - AREA_TEXT_STYLE_SIMPLE / AREA_TEXT_STYLE_BUTTON
     * @returns {TextFrame[]} 作成したエリア内文字
     */
    function convertPointTexts(doc, selectedItems, frameStyle) {
        var createdFrames = [];
        for (var i = selectedItems.length - 1; i >= 0; i--) {
            var pointText = selectedItems[i];
            if (pointText.typename !== "TextFrame" || pointText.kind !== TextType.POINTTEXT) continue;
            /* 1件で失敗しても残りを処理する / One failure must not stop the rest */
            try {
                var frameBox = buildFrameBox(pointText.geometricBounds, frameStyle);
                var textContents = pointText.contents;
                var fontAndSize = getFontAndSize(pointText);

                var frameRect = doc.pathItems.rectangle(frameBox.top, frameBox.left, frameBox.width, frameBox.height);
                var areaFrame = createAreaTextFromPath(doc, frameRect);
                areaFrame.contents = textContents;
                applyFontAndSize(areaFrame, fontAndSize);

                if (frameStyle.centerText) {
                    areaFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
                    /* DOM の verticalAlignment はエリア内文字で効かないことがあるため、
                       調整ダイアログと同じダイナミックアクション（adobe_frameAlignment）で中央に寄せる
                       The DOM's verticalAlignment can fail on Area Type, so the same dynamic action as the adjust dialog centers it */
                    applyAreaTextFrameAlignment(areaFrame, 1, false);
                }

                createdFrames.push(areaFrame);
                pointText.remove();
            } catch (e) { }
        }
        return createdFrames;
    }

    /**
     * スタイル定義からフレーム矩形の位置とサイズを求める
     * @param {number[]} bounds - 元テキストの geometricBounds
     * @param {Object} frameStyle - AREA_TEXT_STYLE_SIMPLE / AREA_TEXT_STYLE_BUTTON
     * @returns {{top: number, left: number, width: number, height: number}} 矩形
     */
    function buildFrameBox(bounds, frameStyle) {
        var origWidth = bounds[2] - bounds[0], origHeight = bounds[1] - bounds[3];
        var width = origWidth * frameStyle.widthScale + frameStyle.widthPadPt;
        var height = origHeight * frameStyle.heightScale + frameStyle.heightPadPt;
        return {
            top: frameStyle.centerOnText ? bounds[1] + (height - origHeight) / 2 : bounds[1],
            left: frameStyle.centerOnText ? bounds[0] - (width - origWidth) / 2 : bounds[0],
            width: width,
            height: height
        };
    }

    /**
     * 段落の行揃え・インデント・行送り・禁則・文字組み・タブを適用する
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object} adjustSettings - readAdjustmentSettings() の戻り値
     * @returns {void}
     */
    function applyParagraphSettings(textFrame, adjustSettings) {
        try {
            var paraAttrs = textFrame.textRange.paragraphAttributes;
            paraAttrs.justification = adjustSettings.justification;
            paraAttrs.leftIndent = adjustSettings.leftIndentPt;
            paraAttrs.rightIndent = adjustSettings.rightIndentPt;
        } catch (e) { }
        applyAutoLeading(textFrame, adjustSettings.leadingPercent);
        applyKinsokuToTextFrame(textFrame, adjustSettings.kinsoku);
        applyMojikumiToTextFrame(textFrame, adjustSettings.mojikumiIndex);
        if (adjustSettings.tabMode === "leader") { applyLeaderTabStops(textFrame); }
        else if (adjustSettings.tabMode === "clear") { clearTabStops(textFrame); }
    }

    // ============================================================
    // 分離ダイアログ：エリア内文字 → 囲み罫＋ポイント文字
    // Separate dialog: Area Type → rectangle plus point text
    // ============================================================

    /**
     * 分離できるエリア内文字か（ロック・非表示のものは対象外）
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} 分離できれば true
     */
    function isSeparableAreaTextFrame(pageItem) {
        /* 削除済みの参照などはプロパティを読むと例外になる / A stale reference throws on any property read */
        try {
            if (!isAreaTextFrame(pageItem)) return false;
            if (pageItem.locked || pageItem.hidden) return false;
            var ownerLayer = pageItem.layer;
            if (ownerLayer && (ownerLayer.locked || !ownerLayer.visible)) return false;
        } catch (e) { return false; }
        return true;
    }

    /**
     * あふれた文字が隠れないよう、枠の高さを収まる最小の高さまで広げる
     * 入れ子のダイアログから app.doScript を呼ばずに済むよう、自動サイズ調整アクションは使わない
     * The auto-size action is avoided here, so nothing calls app.doScript from a nested dialog
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @returns {void}
     */
    function resolveFrameOverset(areaTextFrame) {
        if (!isFrameOverset(areaTextFrame)) return;

        var textPath;
        try { textPath = areaTextFrame.textPath; } catch (e) { return; }

        var startHeight = getFramePathHeight(areaTextFrame);
        if (!(startHeight > 0)) return;

        var growStep = getRepresentativeFontSize(areaTextFrame) * 2;
        if (!(growStep > 0)) growStep = 24;

        /**
         * その高さにしてもあふれが残るか
         * @param {number} frameHeight - 試す高さ（pt）
         * @returns {boolean} あふれが残れば true（高さを変えられなければ true）
         */
        function isOversetAtHeight(frameHeight) {
            try { textPath.height = frameHeight; } catch (e0) { return true; }
            return isFrameOverset(areaTextFrame);
        }

        /* 収まる高さが見つかるまで、伸ばす量を倍にしながら広げる / Grow with a doubling step until a height that fits turns up */
        var oversetHeight = startHeight, fittedHeight = 0;
        for (var attempt = 0; attempt < 24; attempt++) {
            var candidateHeight = oversetHeight + growStep;
            if (!isOversetAtHeight(candidateHeight)) { fittedHeight = candidateHeight; break; }
            oversetHeight = candidateHeight;
            growStep *= 2;
        }
        if (!fittedHeight) {
            /* どこまで広げても収まらないときは元の高さに戻す / Put the original height back when nothing ever fits */
            try { textPath.height = startHeight; } catch (e1) { }
            return;
        }

        /* 収まる範囲で最小の高さに詰める（1pt刻みまで）/ Narrow down to the shortest height that still fits (within 1pt) */
        while (fittedHeight - oversetHeight > 1) {
            var midHeight = (fittedHeight + oversetHeight) / 2;
            if (isOversetAtHeight(midHeight)) { oversetHeight = midHeight; }
            else { fittedHeight = midHeight; }
        }
        try { textPath.height = fittedHeight; } catch (e2) { }
    }

    /**
     * 改行として扱う文字か
     * @param {string} character - 1文字
     * @returns {boolean} 改行なら true
     */
    function isLineBreakChar(character) {
        return character === "\r" || character === "\n" || character === "\u2028" || character === "\u2029";
    }

    /**
     * エリア内文字の親（レイヤーまたはグループ）を返す。取れなければドキュメント
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {Layer|GroupItem|Document} パスとテキストを追加できる親
     */
    function getOwnerContainer(doc, pageItem) {
        try {
            var parentContainer = pageItem.parent;
            if (parentContainer && parentContainer.pathItems && parentContainer.textFrames) return parentContainer;
        } catch (e) { }
        return doc;
    }

    /**
     * 折り返し位置に強制改行を入れた文字列と、1文字ごとの元インデックスを返す
     * @param {TextFrame} areaTextFrame - 対象のエリア内文字
     * @returns {{contents: string, sourceIndices: number[]}|null} 変換後の文字列と元インデックス（折り返しが無ければ null）
     */
    function buildForcedLineBreakText(areaTextFrame) {
        var sourceText, frameLines;
        try {
            sourceText = areaTextFrame.contents;
            frameLines = areaTextFrame.lines;
        } catch (e) { return null; }
        if (!sourceText || !frameLines || frameLines.length < 2) return null;

        var brokenText = "", sourceIndices = [], position = 0;
        for (var i = 0; i < frameLines.length && position < sourceText.length; i++) {
            var lineLength = 0;
            try { lineLength = frameLines[i].characters.length; } catch (e0) { lineLength = 0; }
            for (var j = 0; j < lineLength && position < sourceText.length; j++) {
                brokenText += sourceText.charAt(position);
                sourceIndices.push(position);
                position++;
            }
            if (position >= sourceText.length) break;

            /* すでに改行で終わっている行はそのまま。折り返しだけを強制改行に置き換える
               A line that already ends in a return is left alone; only the wraps become hard returns */
            if (isLineBreakChar(sourceText.charAt(position - 1))) continue;
            if (isLineBreakChar(sourceText.charAt(position))) {
                brokenText += sourceText.charAt(position);
                sourceIndices.push(position);
                position++;
            } else {
                brokenText += "\r";
                sourceIndices.push(position > 0 ? position - 1 : 0);
            }
        }

        /* 文字あふれで lines に出てこない分を、そのまま末尾に足す
           Overset characters never show up in lines, so they are appended as they are */
        while (position < sourceText.length) {
            brokenText += sourceText.charAt(position);
            sourceIndices.push(position);
            position++;
        }
        return { contents: brokenText, sourceIndices: sourceIndices };
    }

    /**
     * 強制改行を入れた文字列に合わせて、控えた文字属性を並べ直す
     * @param {Object[]} snapshots - snapshotCharacterAttributes() の戻り値
     * @param {number[]} sourceIndices - 1文字ごとの元インデックス
     * @returns {Object[]} 並べ直した文字属性
     */
    function remapCharacterSnapshots(snapshots, sourceIndices) {
        if (!snapshots.length) return snapshots;
        var remappedSnapshots = [];
        for (var i = 0; i < sourceIndices.length; i++) {
            var sourceIndex = sourceIndices[i];
            if (sourceIndex < 0) sourceIndex = 0;
            if (sourceIndex >= snapshots.length) sourceIndex = snapshots.length - 1;
            remappedSnapshots.push(snapshots[sourceIndex]);
        }
        return remappedSnapshots;
    }

    /**
     * 囲み罫の見た目を「枠の処理」に合わせる（削除したときは null を返す）
     * @param {PathItem} frameRect - 囲み罫
     * @param {Object} adjustSettings - strokeBlack / hidePath / removePath を持つ設定
     * @returns {PathItem|null} 残した囲み罫
     */
    function applyFrameHandling(frameRect, adjustSettings) {
        if (adjustSettings.strokeBlack) {
            var blackColor = new CMYKColor();
            blackColor.cyan = 0; blackColor.magenta = 0; blackColor.yellow = 0; blackColor.black = 100;
            frameRect.strokeColor = blackColor;
            frameRect.strokeWidth = 1;
        } else if (adjustSettings.hidePath) {
            frameRect.stroked = false;
        } else if (adjustSettings.removePath) {
            frameRect.remove();
            return null;
        }
        return frameRect;
    }

    /**
     * エリア内文字を、囲み罫とポイント文字に分解し、できたオブジェクトを返す
     * できたオブジェクトは元のレイヤー・グループの、元の重ね順の位置に置く
     * The new objects stay in the original layer or group, at the original stacking position
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} areaTextFrame - 分解するエリア内文字
     * @param {Object} adjustSettings - 調整ダイアログの設定に、枠とテキストの処理を足したもの
     * @returns {PageItem[]} 作成したポイント文字と囲み罫
     */
    function separateAreaTextFrame(doc, areaTextFrame, adjustSettings) {
        /* 枠の大きさを読む前に解決しておく（囲み罫と行の分かれ方を実際の文字量に合わせるため）
           Resolved before the size is read, so the rectangle and the line breaks match the real amount of text */
        if (adjustSettings.resolveOverset) { resolveFrameOverset(areaTextFrame); }

        var frameBounds = areaTextFrame.geometricBounds;
        var spacingPt = adjustSettings.spacingPt;
        var left = frameBounds[0] - spacingPt, top = frameBounds[1] + spacingPt;
        var right = frameBounds[2] + spacingPt, bottom = frameBounds[3] - spacingPt;

        /* 削除する前に、文字ごとの書式と1行目のベースライン位置に使うサイズを控える
           Snapshot the per-character formatting and the first-baseline size before the frame goes away */
        var charAttrSnapshots = snapshotCharacterAttributes(areaTextFrame);
        var firstLineOffset = getRepresentativeFontSize(areaTextFrame);
        var pointTextContents = areaTextFrame.contents;

        /* 見かけの改行を強制改行に変換する（文字属性も同じ並びに合わせる）
           Turn the visual wraps into hard returns, keeping the attributes lined up with them */
        if (adjustSettings.forceLineBreaks) {
            var brokenText = buildForcedLineBreakText(areaTextFrame);
            if (brokenText) {
                pointTextContents = brokenText.contents;
                charAttrSnapshots = remapCharacterSnapshots(charAttrSnapshots, brokenText.sourceIndices);
            }
        }

        var ownerContainer = getOwnerContainer(doc, areaTextFrame);
        var frameRect = ownerContainer.pathItems.rectangle(top, left, right - left, top - bottom);
        frameRect.filled = false;
        frameRect.stroked = true;

        var pointTextFrame = ownerContainer.textFrames.add();
        pointTextFrame.contents = pointTextContents;
        pointTextFrame.position = [left, top - firstLineOffset];

        /* 元の重ね順に差し込む（囲み罫の上にポイント文字）/ Slot both in at the original depth, text above the rectangle */
        try {
            frameRect.moveBefore(areaTextFrame);
            pointTextFrame.moveBefore(frameRect);
        } catch (eOrder) { }

        /* 既定の線をいったん外し、このあと文字単位で復元する / Clear the default stroke, then restore per character */
        try {
            pointTextFrame.textRange.characterAttributes.strokeColor = new NoColor();
            pointTextFrame.textRange.characterAttributes.strokeWeight = 0;
        } catch (e0) { }
        restoreCharacterAttributes(pointTextFrame, charAttrSnapshots, false);

        areaTextFrame.remove();

        var createdItems = [pointTextFrame];
        var keptFrameRect = applyFrameHandling(frameRect, adjustSettings);
        if (keptFrameRect) { createdItems.push(keptFrameRect); }

        applyParagraphSettings(pointTextFrame, adjustSettings);
        return createdItems;
    }

    /**
     * 「テキストを分離」ダイアログ。1つでも分離できたら true を返す
     * adjustSettings は調整ダイアログから受け取り、枠とテキストの処理だけを足して使う
     * adjustSettings comes from the adjust dialog; only the frame and text handling is added here
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame[]} areaTextFrames - 分離するエリア内文字
     * @param {Object} adjustSettings - readAdjustmentSettings() の戻り値
     * @returns {boolean} 1つでも分離できたら true
     */
    function showSeparateTextDialog(doc, areaTextFrames, adjustSettings) {
        var separateDialog = new Window("dialog", getLabel("dialog.separateTitle"));
        separateDialog.alignChildren = "fill";
        separateDialog.margins = DIALOG_MARGINS;

        var frameHandlingPanel = separateDialog.add("panel", undefined, getLabel("panel.frameHandling"));
        applyPanelLayout(frameHandlingPanel);
        var radStrokeBlack = frameHandlingPanel.add("radiobutton", undefined, getLabel("radio.strokeBlack"));
        var radHidePath = frameHandlingPanel.add("radiobutton", undefined, getLabel("radio.hidePath"));
        var radRemovePath = frameHandlingPanel.add("radiobutton", undefined, getLabel("radio.removePath"));
        radStrokeBlack.value = true;
        bindExclusiveRadios([radStrokeBlack, radHidePath, radRemovePath]);

        /* テキストの処理：あふれの解決と、折り返しの強制改行化
           Text handling: resolving the overset and turning the wraps into hard returns */
        var textHandlingPanel = separateDialog.add("panel", undefined, getLabel("panel.textHandling"));
        applyPanelLayout(textHandlingPanel);
        var chkResolveOverset = textHandlingPanel.add("checkbox", undefined, getLabel("checkbox.resolveOverset"));
        var chkForceLineBreaks = textHandlingPanel.add("checkbox", undefined, getLabel("checkbox.forceLineBreaks"));
        chkResolveOverset.helpTip = getLabel("tooltip.resolveOverset");
        chkForceLineBreaks.helpTip = getLabel("tooltip.forceLineBreaks");
        chkResolveOverset.value = true;
        chkForceLineBreaks.value = true;

        /* ボタンエリア：右寄せで［キャンセル］［OK］/ Button area: Cancel and OK, right-aligned */
        var btnRowGroup = separateDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["right", "center"];
        btnRowGroup.alignChildren = ["right", "center"];
        var btnCancelSeparate = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnRunSeparate = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        var didSeparate = false;

        btnRunSeparate.onClick = function () {
            adjustSettings.strokeBlack = radStrokeBlack.value;
            adjustSettings.hidePath = radHidePath.value;
            adjustSettings.removePath = radRemovePath.value;
            adjustSettings.resolveOverset = chkResolveOverset.value;
            adjustSettings.forceLineBreaks = chkForceLineBreaks.value;

            /* 分離するとオブジェクトが置き換わるので後ろから処理する
               Each frame is replaced as it goes, so the list is walked from the back */
            var createdItems = [];
            for (var i = areaTextFrames.length - 1; i >= 0; i--) {
                var areaTextFrame = areaTextFrames[i];
                if (!isSeparableAreaTextFrame(areaTextFrame)) continue;
                /* 1フレームで失敗しても残りを処理できるようにする / One failing frame must not stop the rest */
                try {
                    createdItems = createdItems.concat(separateAreaTextFrame(doc, areaTextFrame, adjustSettings));
                    didSeparate = true;
                } catch (e) { }
            }

            /* 何も分離できなかったときは、選択も触らずダイアログを開いたままにする
               When nothing could be separated, the selection is left alone and the dialog stays open */
            if (!didSeparate) { alert(getLabel("alert.separateFailed")); return; }

            /* 削除済みの参照が選択に残らないよう、できたオブジェクトを選び直す
               Re-select what was created, so the deleted frames do not linger in the selection */
            try { doc.selection = createdItems.length ? createdItems : null; } catch (e2) { }
            app.redraw();
            separateDialog.close(1);
        };
        btnCancelSeparate.onClick = function () { separateDialog.close(0); };

        separateDialog.show();
        return didSeparate;
    }

    // ============================================================
    // 調整ダイアログの組み立て / Building the adjust dialog
    // ============================================================

    /**
     * 行揃え・テキストの配置のアイコンボタンを1つ追加する（描画は呼び出し側で onDraw に付ける）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} helpTipText - ツールチップ
     * @param {string} iconType - アイコン種別
     * @returns {Button} 追加したボタン
     */
    function addIconButton(parentPanel, helpTipText, iconType) {
        var iconButton = parentPanel.add("button", undefined, "");
        iconButton.helpTip = helpTipText;
        iconButton.preferredSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
        iconButton.minimumSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
        iconButton.maximumSize = [ICON_BUTTON_SIZE, ICON_BUTTON_SIZE];
        iconButton.iconType = iconType;
        return iconButton;
    }

    /**
     * ラベル・数値欄・単位を並べた行を追加する
     * @param {Group} parentGroup - 追加先
     * @param {string} labelString - 行ラベル（空文字なら字下げ用の空きラベル）
     * @param {number|null} labelWidth - 行ラベルの幅（null なら指定しない）
     * @param {string} initialText - 欄の初期値
     * @param {number} fieldCharacters - 欄の文字数
     * @param {string} unitText - 単位
     * @returns {Object} { row, label, field, unitLabel }
     */
    function addNumberRow(parentGroup, labelString, labelWidth, initialText, fieldCharacters, unitText) {
        var fieldRow = parentGroup.add("group");
        var rowLabel = fieldRow.add("statictext", undefined, labelString);
        if (labelWidth !== null) { rowLabel.preferredSize.width = labelWidth; }
        var numberField = fieldRow.add("edittext", undefined, initialText);
        numberField.characters = fieldCharacters;
        var unitLabel = fieldRow.add("statictext", undefined, unitText);
        return { row: fieldRow, label: rowLabel, field: numberField, unitLabel: unitLabel };
    }

    /**
     * 右カラム：種別（本文／見出し／メニュー）パネルを追加する
     * @param {Object} dialogControls - コントロールの格納先
     * @param {Group} parentColumn - 追加先の列
     * @returns {void}
     */
    function addRolePanel(dialogControls, parentColumn) {
        var rolePanel = parentColumn.add("panel", undefined, getLabel("panel.role"));
        applyPanelLayout(rolePanel, 4);
        rolePanel.orientation = "row";
        rolePanel.alignChildren = ["left", "center"];
        dialogControls.radRoleBody = rolePanel.add("radiobutton", undefined, getLabel("radio.roleBody"));
        dialogControls.radRoleHeading = rolePanel.add("radiobutton", undefined, getLabel("radio.roleHeading"));
        dialogControls.radRoleMenu = rolePanel.add("radiobutton", undefined, getLabel("radio.roleMenu"));
        dialogControls.radRoleBody.helpTip = getLabel("tooltip.roleBody");
        dialogControls.radRoleHeading.helpTip = getLabel("tooltip.roleHeading");
        dialogControls.radRoleMenu.helpTip = getLabel("tooltip.roleMenu");
        dialogControls.roleRadios = [dialogControls.radRoleBody, dialogControls.radRoleHeading, dialogControls.radRoleMenu];
    }

    /**
     * 右カラム：行送り（実寸と自動行送り量％）パネルを追加する
     * @param {Object} dialogControls - コントロールの格納先
     * @param {Group} parentColumn - 追加先の列
     * @returns {void}
     */
    function addLeadingPanel(dialogControls, parentColumn) {
        var leadingPanel = parentColumn.add("panel", undefined, getLabel("panel.leading"));
        applyPanelLayout(leadingPanel);
        dialogControls.etLeadingEffective = addNumberRow(leadingPanel, labelText("fieldLabel.leadingEffective"), LEADING_LABEL_WIDTH, "", SMALL_FIELD_CHARACTERS, "pt").field;
        dialogControls.etLeadingPercent = addNumberRow(leadingPanel, labelText("fieldLabel.leadingPercent"), LEADING_LABEL_WIDTH, "", SMALL_FIELD_CHARACTERS, "%").field;
        leadingPanel.helpTip = getLabel("tooltip.leading");
        dialogControls.etLeadingPercent.helpTip = leadingPanel.helpTip;
        dialogControls.etLeadingEffective.helpTip = leadingPanel.helpTip;
    }

    /**
     * 右カラム：行揃え（アイコンボタン。ラベルはツールチップで見せる）パネルを追加する
     * @param {Object} dialogControls - コントロールの格納先
     * @param {Group} parentColumn - 追加先の列
     * @param {Object} justifyState - 選択中の id と UI 明暗（onDraw から参照）
     * @returns {void}
     */
    function addJustificationPanel(dialogControls, parentColumn, justifyState) {
        var justificationPanel = parentColumn.add("panel", undefined, getLabel("panel.justification"));
        applyPanelLayout(justificationPanel, 4);
        justificationPanel.orientation = "row";
        justificationPanel.alignChildren = ["center", "center"];

        dialogControls.justifyButtons = [];
        for (var i = 0; i < JUSTIFY_OPTIONS.length; i++) {
            var justifyOption = JUSTIFY_OPTIONS[i];
            var justifyButton = addIconButton(justificationPanel, getLabel(justifyOption.labelKey) + " (" + justifyOption.shortcut + ")", justifyOption.icon);
            justifyButton.justifyId = justifyOption.id;
            justifyButton.onDraw = function () {
                drawJustifyIcon(this, this.justifyId === justifyState.activeId, justifyState.isLight);
            };
            dialogControls.justifyButtons.push(justifyButton);
        }
    }

    /**
     * 右カラム：インデント（左右と連動）パネルを追加する
     * @param {Object} dialogControls - コントロールの格納先
     * @param {Group} parentColumn - 追加先の列
     * @param {string} rulerLabel - 定規の単位の表示名
     * @returns {void}
     */
    function addIndentPanel(dialogControls, parentColumn, rulerLabel) {
        var indentPanel = parentColumn.add("panel", undefined, getLabel("panel.indent"));
        applyPanelLayout(indentPanel);
        indentPanel.orientation = "row";
        indentPanel.alignChildren = ["left", "top"];
        var indentFieldsColumn = indentPanel.add("group");
        indentFieldsColumn.orientation = "column";
        indentFieldsColumn.alignChildren = "left";
        dialogControls.etLeftIndent = addNumberRow(indentFieldsColumn, labelText("fieldLabel.indentLeft"), null, "0", SMALL_FIELD_CHARACTERS, rulerLabel).field;
        dialogControls.etRightIndent = addNumberRow(indentFieldsColumn, labelText("fieldLabel.indentRight"), null, "0", SMALL_FIELD_CHARACTERS, rulerLabel).field;
        var linkColumn = indentPanel.add("group");
        linkColumn.orientation = "column";
        linkColumn.alignChildren = "left";
        linkColumn.alignment = ["left", "center"];
        dialogControls.chkLinkIndents = linkColumn.add("checkbox", undefined, getLabel("checkbox.linkIndents"));
        dialogControls.chkLinkIndents.helpTip = getLabel("tooltip.linkIndents");
        dialogControls.chkLinkIndents.value = true;
    }

    /**
     * ラベルと、幅を指定したドロップダウンを縦に並べて追加する（既定では fill でパネル幅いっぱいに広がるため幅を指定する）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - ラベルのパス
     * @param {Object[]} choices - 選択肢テーブル
     * @param {string} tooltipPath - ツールチップのラベルパス
     * @returns {DropDownList} 追加したドロップダウン
     */
    function addLabeledDropdown(parentPanel, labelPath, choices, tooltipPath) {
        var dropdownLabel = parentPanel.add("statictext", undefined, labelText(labelPath));
        var choiceDropdown = parentPanel.add("dropdownlist", undefined, buildChoiceLabels(choices));
        dropdownLabel.helpTip = getLabel(tooltipPath);
        choiceDropdown.helpTip = dropdownLabel.helpTip;
        return choiceDropdown;
    }

    /**
     * 右カラム：日本語の組版（禁則・文字組みアキ量設定）パネルを追加する
     * @param {Object} dialogControls - コントロールの格納先
     * @param {Group} parentColumn - 追加先の列
     * @returns {void}
     */
    function addJpCompositionPanel(dialogControls, parentColumn) {
        var jpCompositionPanel = parentColumn.add("panel", undefined, getLabel("panel.jpComposition"));
        applyPanelLayout(jpCompositionPanel, 4);
        dialogControls.kinsokuDropdown = addLabeledDropdown(jpCompositionPanel, "fieldLabel.kinsoku", KINSOKU_CHOICES, "tooltip.kinsoku");
        selectChoiceByValue(dialogControls.kinsokuDropdown, KINSOKU_CHOICES, "id", DEFAULT_KINSOKU);
        /* fill を打ち消して幅を指定する / Cancel fill and set the width */
        dialogControls.kinsokuDropdown.alignment = "left";
        dialogControls.kinsokuDropdown.preferredSize.width = JP_DROPDOWN_WIDTH;
        /* 禁則との間を少し空ける / A little breathing room after the kinsoku row */
        var mojikumiGap = jpCompositionPanel.add("group");
        mojikumiGap.preferredSize.height = ROW_GAP_HEIGHT;
        dialogControls.mojikumiDropdown = addLabeledDropdown(jpCompositionPanel, "fieldLabel.mojikumi", MOJIKUMI_CHOICES, "tooltip.mojikumi");
        selectChoiceByValue(dialogControls.mojikumiDropdown, MOJIKUMI_CHOICES, "index", DEFAULT_MOJIKUMI_INDEX);
        dialogControls.mojikumiDropdown.alignment = "left";
        dialogControls.mojikumiDropdown.preferredSize.width = JP_DROPDOWN_WIDTH;
    }

    /**
     * 左カラム：フォントサイズ（欄と［文字あふれ解消］［枠にフィット］）パネルを追加する
     * @param {Object} dialogControls - コントロールの格納先
     * @param {Group} parentColumn - 追加先の列
     * @returns {void}
     */
    function addFontSizePanel(dialogControls, parentColumn) {
        var fontSizePanel = parentColumn.add("panel", undefined, getLabel("panel.fontSize"));
        applyPanelLayout(fontSizePanel);
        /* パネル名が「フォントサイズ」なので、行のラベルは省く / The panel title already says it, so the row label is dropped */
        var fontSizeRow = fontSizePanel.add("group");
        fontSizeRow.alignment = "left";
        dialogControls.etFontSize = fontSizeRow.add("edittext", undefined, "");
        dialogControls.etFontSize.characters = SMALL_FIELD_CHARACTERS;
        fontSizeRow.add("statictext", undefined, "pt");
        /* フォントサイズ欄との間を少し空ける / A little breathing room after the font-size field */
        var fontSizeButtonGap = fontSizePanel.add("group");
        fontSizeButtonGap.preferredSize.height = ROW_GAP_HEIGHT;

        /* 縦並び。ボタンはラベル幅のまま（パネル幅いっぱいに伸ばさない）
           Stacked vertically, each button keeping its label width (not stretched to the panel) */
        var fontSizeButtonRow = fontSizePanel.add("group");
        fontSizeButtonRow.orientation = "column";
        fontSizeButtonRow.alignChildren = ["left", "top"];
        dialogControls.btnShrinkToFit = fontSizeButtonRow.add("button", undefined, getLabel("button.shrinkToFit"));
        dialogControls.btnFitFontSize = fontSizeButtonRow.add("button", undefined, getLabel("button.fitFontSize"));
        dialogControls.btnShrinkToFit.helpTip = getLabel("tooltip.shrinkToFit");
        dialogControls.btnFitFontSize.helpTip = getLabel("tooltip.fitFontSize");
    }

    /**
     * 左カラム：フレームサイズ（幅・1行の文字数・高さ・自動サイズ調整）パネルを追加する
     * @param {Object} dialogControls - コントロールの格納先
     * @param {Group} parentColumn - 追加先の列
     * @param {string} rulerLabel - 定規の単位の表示名
     * @returns {void}
     */
    function addFrameSizePanel(dialogControls, parentColumn, rulerLabel) {
        var frameSizePanel = parentColumn.add("panel", undefined, getLabel("panel.frameSize"));
        applyPanelLayout(frameSizePanel);
        dialogControls.etWidth = addNumberRow(frameSizePanel, labelText("fieldLabel.width"), FRAME_LABEL_WIDTH, "", SIZE_FIELD_CHARACTERS, rulerLabel).field;
        /* 幅の下に字詰め欄。空きラベルで幅の入力欄と左端をそろえる
           Chars per line goes under the width, lined up with the width field via an empty label */
        var charsPerLineRow = addNumberRow(frameSizePanel, "", FRAME_LABEL_WIDTH, "", SMALL_FIELD_CHARACTERS, getLabel("fieldLabel.charsPerLine"));
        dialogControls.etCharsPerLine = charsPerLineRow.field;
        var lblCharsPerLine = charsPerLineRow.unitLabel;
        lblCharsPerLine.helpTip = getLabel("tooltip.charsPerLine");
        dialogControls.etCharsPerLine.helpTip = lblCharsPerLine.helpTip;
        var heightRow = addNumberRow(frameSizePanel, labelText("fieldLabel.height"), FRAME_LABEL_WIDTH, "", SIZE_FIELD_CHARACTERS, rulerLabel);
        dialogControls.heightRow = heightRow.row;
        dialogControls.etHeight = heightRow.field;
        /* 高さの下に自動サイズ調整（heightRow の外に置く。ONのあいだ heightRow はディムするため）
           Auto-size sits under the height (outside heightRow, which gets dimmed while it is on) */
        dialogControls.chkAutoSize = frameSizePanel.add("checkbox", undefined, getLabel("checkbox.autoSize"));
        dialogControls.chkAutoSize.helpTip = getLabel("tooltip.autoSize");
        /* 英語UIでは字詰めの計算が不正確なため使用不可にする / Chars per line is disabled in the English UI, where it is inaccurate */
        if (uiLang !== "ja") {
            dialogControls.etCharsPerLine.enabled = false;
            lblCharsPerLine.enabled = false;
        }
    }

    /**
     * 左カラム：オフセット（パネル名で足りるので、チェックボックスのラベルは省く）パネルを追加する
     * @param {Object} dialogControls - コントロールの格納先
     * @param {Group} parentColumn - 追加先の列
     * @param {string} rulerLabel - 定規の単位の表示名
     * @returns {void}
     */
    function addOffsetPanel(dialogControls, parentColumn, rulerLabel) {
        var offsetPanel = parentColumn.add("panel", undefined, getLabel("panel.offset"));
        applyPanelLayout(offsetPanel);
        var spacingRow = offsetPanel.add("group");
        dialogControls.chkSpacing = spacingRow.add("checkbox", undefined, "");
        dialogControls.etSpacing = spacingRow.add("edittext", undefined, "0");
        dialogControls.etSpacing.characters = SMALL_FIELD_CHARACTERS;
        dialogControls.lblSpacingUnit = spacingRow.add("statictext", undefined, rulerLabel);
        dialogControls.chkSpacing.helpTip = getLabel("tooltip.offset");
        dialogControls.etSpacing.helpTip = dialogControls.chkSpacing.helpTip;
        dialogControls.etSpacing.enabled = false;
        dialogControls.lblSpacingUnit.enabled = false;
    }

    /**
     * 左カラム：テキストの配置（アイコンボタン。ラベルはツールチップで見せる）パネルを追加する
     * @param {Object} dialogControls - コントロールの格納先
     * @param {Group} parentColumn - 追加先の列
     * @param {Object} alignState - 選択中の id と UI 明暗（onDraw から参照）
     * @returns {void}
     */
    function addTextAlignPanel(dialogControls, parentColumn, alignState) {
        var textAlignPanel = parentColumn.add("panel", undefined, getLabel("panel.textAlign"));
        applyPanelLayout(textAlignPanel, 4);
        textAlignPanel.orientation = "row";
        textAlignPanel.alignChildren = ["center", "center"];
        textAlignPanel.helpTip = getLabel("tooltip.textAlign");

        dialogControls.alignButtons = [];
        for (var i = 0; i < ALIGN_OPTIONS.length; i++) {
            var alignOption = ALIGN_OPTIONS[i];
            var alignButton = addIconButton(textAlignPanel, getLabel(alignOption.labelKey), alignOption.icon);
            alignButton.alignId = alignOption.id;
            alignButton.onDraw = function () {
                drawAlignIcon(this, this.alignId === alignState.activeId, alignState.isLight);
            };
            dialogControls.alignButtons.push(alignButton);
        }
    }

    /**
     * 調整ダイアログを組み立てる（イベントは showAdjustDialog() で付ける）
     * @param {Object} rulerInfo - 定規の単位（getUnitInfo() の結果）
     * @param {Object} justifyState - 行揃えの選択中 id と UI 明暗
     * @param {Object} alignState - テキストの配置の選択中 id と UI 明暗
     * @returns {Object} ダイアログ本体（window）と各コントロール
     */
    function buildAdjustDialog(rulerInfo, justifyState, alignState) {
        var dialogControls = {};
        dialogControls.window = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialogControls.window.alignChildren = "fill";
        dialogControls.window.margins = DIALOG_MARGINS;

        /* 2カラムレイアウト（左右カラムは上揃えで横いっぱいに）/ Two-column layout (columns fill width, top-aligned) */
        var columnsGroup = dialogControls.window.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = 10;

        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";

        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";

        /* 右カラム：種別・行送り・行揃え・インデント・日本語の組版 / Right column */
        addRolePanel(dialogControls, rightColumn);
        addLeadingPanel(dialogControls, rightColumn);
        addJustificationPanel(dialogControls, rightColumn, justifyState);
        addIndentPanel(dialogControls, rightColumn, rulerInfo.label);
        addJpCompositionPanel(dialogControls, rightColumn);

        /* 左カラム：フォントサイズ・フレームサイズ・オフセット・テキストの配置 / Left column */
        addFontSizePanel(dialogControls, leftColumn);
        addFrameSizePanel(dialogControls, leftColumn, rulerInfo.label);
        addOffsetPanel(dialogControls, leftColumn, rulerInfo.label);
        addTextAlignPanel(dialogControls, leftColumn, alignState);

        /* ボタンエリア：左に［テキストを分離...］、右に［キャンセル］［OK］
           Button area: "Separate text..." on the left, Cancel and OK on the right */
        var btnRowGroup = dialogControls.window.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "fill";
        btnRowGroup.alignChildren = ["fill", "center"];
        dialogControls.btnSeparateText = btnRowGroup.add("button", undefined, getLabel("button.separateText"));
        dialogControls.btnSeparateText.alignment = ["left", "center"];
        dialogControls.btnSeparateText.helpTip = getLabel("tooltip.separateText");
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        dialogControls.btnCancelAdjust = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        dialogControls.btnRun = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return dialogControls;
    }

    // ============================================================
    // 調整ダイアログ：エリア内文字の調整 / Adjust dialog: adjusting Area Type
    // ============================================================

    /* 幅・高さの上限（pt）。極端な値で Illustrator が不安定になるのを防ぐ / Upper limit for width and height (pt), keeping Illustrator stable */
    var MAX_FRAME_SIZE_PT = 100000;

    /**
     * 配列やコレクションを、普通の配列に写す
     * @param {Object} itemList - 配列、または length と添字を持つコレクション
     * @returns {Object[]} 写した配列
     */
    function copyItemList(itemList) {
        var copiedItems = [];
        for (var i = 0; i < itemList.length; i++) { copiedItems.push(itemList[i]); }
        return copiedItems;
    }

    /**
     * pt の値を、定規の単位で小数第2位までの数値にする（入力欄の表示用）
     * @param {number} ptValue - pt の値
     * @param {Object} rulerInfo - 定規の単位（getUnitInfo() の結果）
     * @returns {number} 定規の単位の値
     */
    function toRulerFieldValue(ptValue, rulerInfo) {
        return Math.round((ptValue / rulerInfo.pointsPerUnit) * 100) / 100;
    }

    /**
     * 幅・高さ欄の値を検証して返す（0以下・NaN・極端値を弾いて Illustrator の不安定化を避ける）
     * 不正なときは直前の有効な値に戻す。上限・下限を超えたら丸めて欄も書き換える
     * @param {EditText} sizeField - 幅または高さの欄
     * @param {number|null} lastValue - 直前の有効な値（定規の単位）
     * @param {Object} rulerInfo - 定規の単位（getUnitInfo() の結果）
     * @returns {number|null} 有効な値（定規の単位。不正なら null）
     */
    function validateSizeField(sizeField, lastValue, rulerInfo) {
        var value = parseFloat(String(sizeField.text));
        if (isNaN(value) || !isFinite(value) || value <= 0) {
            if (lastValue !== null) sizeField.text = lastValue;
            return null;
        }
        var maxValue = MAX_FRAME_SIZE_PT / rulerInfo.pointsPerUnit;
        if (value > maxValue) {
            value = maxValue;
            sizeField.text = Math.round(value * 100) / 100;
        }
        /* 小さすぎる値も事故の元なので下限を設ける / Too small a value is just as risky, so there is a floor too */
        if (value < 0.01) {
            value = 0.01;
            sizeField.text = Math.round(value * 100) / 100;
        }
        return value;
    }

    /**
     * 調整対象を選択し直し、調整するフレームの一覧を返す
     * 変換ダイアログから受け取った変換結果を確実に対象にする（選択が変わっても崩れないように）。
     * モーダルダイアログ中は selection が変動・取得不能になることがあるため、ここで控えておく
     * The converted frames are pinned as targets, since the selection can shift or become unreadable in a modal dialog
     * @param {TextFrame|null} initialFrame - 値を読み込むフレーム
     * @param {TextFrame[]|null} targetFrames - 変換ダイアログで作ったフレーム
     * @returns {PageItem[]|null} 調整するフレーム（決まらなければ null）
     */
    function collectAdjustTargets(initialFrame, targetFrames) {
        var hasTargetFrames = !!(targetFrames && targetFrames.length);
        if (hasTargetFrames) {
            try { app.activeDocument.selection = targetFrames; } catch (e) { }
        } else if (initialFrame) {
            try { app.activeDocument.selection = [initialFrame]; } catch (e2) { }
        }
        app.redraw();

        if (hasTargetFrames) return targetFrames.slice(0);
        try {
            var currentSelection = app.activeDocument.selection;
            if (currentSelection && currentSelection.length) return copyItemList(currentSelection);
        } catch (e4) { }
        return null;
    }

    /**
     * 調整ダイアログを表示する（プレビューを貼りながら調整し、OK で確定する）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame|null} initialFrame - 値を読み込むフレーム
     * @param {TextFrame[]|null} targetFrames - 変換ダイアログで作ったフレーム（既存のエリア内文字なら null）
     * @param {number|null} initialAlignmentValue - テキストの配置の初期値（null なら上揃えで開く）
     * @returns {void}
     */
    function showAdjustDialog(doc, initialFrame, targetFrames, initialAlignmentValue) {
        var rulerInfo = getUnitInfo("rulerType");

        /* 変換したてのフレームかどうか。既存のエリア内文字では、DOM から読み取れない項目
           （テキストの配置・自動サイズ調整）や未設定の項目を、触っていないのに書き換えない
           Whether these frames were just converted. For existing Area Type, settings that cannot be
           read back from the DOM (placement, auto-size) are left alone unless the user touches them */
        var isNewlyConverted = !!(targetFrames && targetFrames.length);

        /* ユーザーが操作した項目 / Which settings the user has touched */
        var userTouched = {
            alignment: isNewlyConverted,
            height: isNewlyConverted,
            kinsoku: isNewlyConverted,
            mojikumi: isNewlyConverted
        };

        /* 調整するフレーム（undo 後などは選択から拾い直す）/ Frames to adjust (re-picked from the selection after an undo, etc.) */
        var targetAreaFrames = collectAdjustTargets(initialFrame, targetFrames);

        /* 行揃え・テキストの配置の選択中 id と UI 明暗（onDraw のクロージャから参照）
           Active ids and UI brightness, shared with the onDraw closures */
        var isLightTheme = isLightUI();
        var justifyState = { activeId: "left", isLight: isLightTheme };
        /* ボタン風で変換したときは中央、それ以外は上揃えで開く / Opens centered after a Button-style conversion, otherwise top */
        var alignState = { activeId: getAlignmentId(initialAlignmentValue), isLight: isLightTheme };
        /* 選択中の種別と、それが決めたタブ設定。tabMode が "none" のあいだはタブに触らない
           The active role and the tab setting it chose; tabs are left alone while tabMode is "none" */
        var roleState = { activeId: "", tabMode: "none" };

        var dialogControls = buildAdjustDialog(rulerInfo, justifyState, alignState);

        /* 状態 / State */
        var isPreviewActive = false;
        var fontFitMode = "none";
        var currentFontSize = 0;
        var lastValidWidth = null;  /* 定規の単位 / ruler units */
        var lastValidHeight = null; /* 定規の単位 / ruler units */

        /**
         * 現在の選択からエリア内文字だけを拾い直す
         * @returns {void}
         */
        function refreshTargetAreaFrames() {
            /* undo 直後などは選択を読めないことがある / The selection can be unreadable right after an undo */
            try {
                var selectedItems = app.activeDocument.selection;
                var areaFrames = [];
                if (selectedItems && selectedItems.length) {
                    for (var i = 0; i < selectedItems.length; i++) {
                        if (isAreaTextFrame(selectedItems[i])) { areaFrames.push(selectedItems[i]); }
                    }
                }
                targetAreaFrames = areaFrames.length ? areaFrames : null;
            } catch (e) {
                targetAreaFrames = null;
            }
        }

        /**
         * 調整対象を解決する（固定ターゲット優先、無ければ選択）
         * @returns {PageItem[]} 調整するフレーム
         */
        function getTargetAreaFrames() {
            if (!(targetAreaFrames && targetAreaFrames.length)) { refreshTargetAreaFrames(); }
            return (targetAreaFrames && targetAreaFrames.length) ? targetAreaFrames : app.activeDocument.selection;
        }

        /**
         * 幅の欄を検証し、有効なら直前の値として控える
         * @returns {number|null} 有効な幅（定規の単位。不正なら null）
         */
        function validateWidthField() {
            var widthValue = validateSizeField(dialogControls.etWidth, lastValidWidth, rulerInfo);
            if (widthValue !== null) { lastValidWidth = widthValue; }
            return widthValue;
        }

        /**
         * 高さの欄を検証し、有効なら直前の値として控える
         * @returns {number|null} 有効な高さ（定規の単位。不正なら null）
         */
        function validateHeightField() {
            var heightValue = validateSizeField(dialogControls.etHeight, lastValidHeight, rulerInfo);
            if (heightValue !== null) { lastValidHeight = heightValue; }
            return heightValue;
        }

        /**
         * フレームの幅・高さを欄に表示し、直前の有効な値として控える
         * @param {TextFrame} areaFrame - 対象のエリア内文字
         * @returns {void}
         */
        function showFrameSizeInFields(areaFrame) {
            dialogControls.etWidth.text = toRulerFieldValue(areaFrame.textPath.width, rulerInfo);
            dialogControls.etHeight.text = toRulerFieldValue(areaFrame.textPath.height, rulerInfo);
            lastValidWidth = parseFloat(dialogControls.etWidth.text);
            lastValidHeight = parseFloat(dialogControls.etHeight.text);
        }

        /**
         * アイコンボタンを再描画する
         * @param {Button[]} iconButtons - 行揃えまたはテキストの配置のボタン
         * @returns {void}
         */
        function redrawIconButtons(iconButtons) {
            /* 閉じかけのダイアログなどでは再描画できないことがある / Repainting can fail, e.g. while the dialog is closing */
            for (var i = 0; i < iconButtons.length; i++) {
                try { iconButtons[i].notify("onDraw"); } catch (e) { }
            }
            try { dialogControls.window.update(); } catch (e2) { }
        }

        /**
         * メニュー指定を解除し、設定済みのタブも削除する
         * @returns {void}
         */
        function cancelMenuRole() {
            if (roleState.activeId !== "menu") return;
            roleState.activeId = "";
            roleState.tabMode = "clear";
            selectRadio(dialogControls.roleRadios, null);
        }

        /**
         * 行揃えを選ぶ（メニューの右揃え以外にしたら、メニュー指定を解除する）
         * @param {string} justifyId - 行揃え id
         * @returns {void}
         */
        function setJustification(justifyId) {
            justifyState.activeId = justifyId;
            if (justifyId !== ROLE_PRESETS.menu.justifyId) { cancelMenuRole(); }
            redrawIconButtons(dialogControls.justifyButtons);
        }

        /**
         * 定規の単位の欄を pt にする（数値でなければ 0）
         * @param {EditText} numberField - 対象の欄
         * @returns {number} pt の値
         */
        function fieldToPt(numberField) {
            return (parseFloat(numberField.text) || 0) * rulerInfo.pointsPerUnit;
        }

        /**
         * 左右インデントとオフセット（間隔）を pt で読む（連動中の右は左と同じ、オフセット OFF なら 0）
         * @returns {{leftIndentPt: number, rightIndentPt: number, spacingPt: number}} インデントと間隔
         */
        function readIndentAndSpacingPt() {
            var leftIndentPt = fieldToPt(dialogControls.etLeftIndent);
            return {
                leftIndentPt: leftIndentPt,
                rightIndentPt: dialogControls.chkLinkIndents.value ? leftIndentPt : fieldToPt(dialogControls.etRightIndent),
                spacingPt: dialogControls.chkSpacing.value ? fieldToPt(dialogControls.etSpacing) : 0
            };
        }

        /**
         * 幅から差し引く余白（間隔×2＋左右インデント）を返す
         * @returns {number} 余白（pt）
         */
        function getWidthAdjustmentPt() {
            var indentAndSpacing = readIndentAndSpacingPt();
            return 2 * indentAndSpacing.spacingPt + indentAndSpacing.leftIndentPt + indentAndSpacing.rightIndentPt;
        }

        /**
         * 幅とフォントサイズから字詰め欄を更新する
         * @returns {void}
         */
        function updateCharsPerLineField() {
            if (currentFontSize <= 0) return;
            var widthPt = fieldToPt(dialogControls.etWidth);
            dialogControls.etCharsPerLine.text = Math.round(((widthPt - getWidthAdjustmentPt()) / currentFontSize) * 100) / 100;
        }

        /**
         * 禁則・文字組みの現在値をダイアログに読み込む
         * 設定済みならその値を表示して適用対象にする（表示＝実際の設定なので、当ててもフレームは変わらない）。
         * 未設定のときは、変換直後だけ初期値（DEFAULT_KINSOKU / DEFAULT_MOJIKUMI_INDEX）を当て、
         * 既存フレームでは「なし」を表示して勝手に付けない。混在（-2）のときは未選択にして触らない
         * A value that is already set is shown and applied back (a no-op, but it carries over when separating).
         * When nothing is set, the defaults apply to freshly converted frames only; existing frames show "None".
         * @param {TextFrame} sourceFrame - 読み込むフレーム
         * @returns {void}
         */
        function loadCompositionFromFrame(sourceFrame) {
            var frameKinsoku = getKinsokuId(sourceFrame);
            if (frameKinsoku !== "None") {
                selectChoiceByValue(dialogControls.kinsokuDropdown, KINSOKU_CHOICES, "id", frameKinsoku);
                userTouched.kinsoku = true;
            } else if (!isNewlyConverted) {
                selectChoiceByValue(dialogControls.kinsokuDropdown, KINSOKU_CHOICES, "id", "None");
                userTouched.kinsoku = true;
            }
            var frameMojikumi = getMojikumiIndex(sourceFrame);
            if (frameMojikumi >= 0) {
                selectChoiceByValue(dialogControls.mojikumiDropdown, MOJIKUMI_CHOICES, "index", frameMojikumi);
                userTouched.mojikumi = true;
            } else if (frameMojikumi === -2) {
                dialogControls.mojikumiDropdown.selection = null;
            } else if (!isNewlyConverted) {
                selectChoiceByValue(dialogControls.mojikumiDropdown, MOJIKUMI_CHOICES, "index", -1);
                userTouched.mojikumi = true;
            }
        }

        /**
         * 選択フレームの現在値をダイアログに読み込む
         * @param {TextFrame} sourceFrame - 読み込むフレーム
         * @returns {void}
         */
        function loadValuesFromFrame(sourceFrame) {
            showFrameSizeInFields(sourceFrame);
            currentFontSize = 0;
            try { currentFontSize = sourceFrame.textRange.characterAttributes.size || 0; } catch (e) { }
            if (currentFontSize > 0) { dialogControls.etFontSize.text = Math.round(currentFontSize * 100) / 100; }
            try {
                var justification = sourceFrame.paragraphs.length > 0
                    ? sourceFrame.paragraphs[0].paragraphAttributes.justification
                    : Justification.LEFT;
                justifyState.activeId = getJustificationId(justification);
            } catch (e) { justifyState.activeId = "left"; }
            try {
                var spacingPt = sourceFrame.spacing || 0;
                dialogControls.etSpacing.text = toRulerFieldValue(spacingPt, rulerInfo);
                dialogControls.chkSpacing.value = (spacingPt !== 0);
                updateSpacingEnabled();
            } catch (e) { }
            try {
                var firstParaAttrs = sourceFrame.paragraphs.length > 0 ? sourceFrame.paragraphs[0].paragraphAttributes : null;
                var leftIndentPt = firstParaAttrs ? (firstParaAttrs.leftIndent || 0) : 0;
                var rightIndentPt = firstParaAttrs ? (firstParaAttrs.rightIndent || 0) : 0;
                dialogControls.etLeftIndent.text = toRulerFieldValue(leftIndentPt, rulerInfo);
                dialogControls.etRightIndent.text = toRulerFieldValue(rightIndentPt, rulerInfo);
                /* 左右が違うテキストは連動を外して開く（右の値を潰さないため）
                   Open with the link off when the two differ, so the right value is not overwritten */
                if (Math.abs(leftIndentPt - rightIndentPt) > 0.01) { dialogControls.chkLinkIndents.value = false; }
            } catch (e) { }
            loadCompositionFromFrame(sourceFrame);

            var leadingPercent = getAutoLeadingPercent(sourceFrame);
            dialogControls.etLeadingPercent.text = (leadingPercent > 0) ? Math.round(leadingPercent * 10) / 10 : "";
            updateLeadingEffective();

            updateCharsPerLineField();
        }

        /**
         * ダイアログの入力を1つの設定オブジェクトにまとめる
         * 幅・高さの検証はここで1回だけ行う（不正なら null にして書き込まない）
         * Width and height are validated once here; null means the value is not written back
         * @returns {Object} 調整の設定
         */
        function readAdjustmentSettings() {
            var widthValue = validateWidthField();
            var heightValue = validateHeightField();
            var indentAndSpacing = readIndentAndSpacingPt();

            return {
                shrinkFont: (fontFitMode === "shrink"),
                fitFont: (fontFitMode === "fit"),
                autoSize: dialogControls.chkAutoSize.value,
                tabMode: roleState.tabMode,
                justification: getJustificationValue(justifyState.activeId),
                alignment: userTouched.alignment ? getAlignmentValue(alignState.activeId) : null,
                widthPt: (widthValue !== null) ? widthValue * rulerInfo.pointsPerUnit : null,
                heightPt: (heightValue !== null && userTouched.height) ? heightValue * rulerInfo.pointsPerUnit : null,
                leadingPercent: parseFloat(dialogControls.etLeadingPercent.text),
                kinsoku: (userTouched.kinsoku && dialogControls.kinsokuDropdown.selection) ? KINSOKU_CHOICES[dialogControls.kinsokuDropdown.selection.index].id : null,
                mojikumiIndex: (userTouched.mojikumi && dialogControls.mojikumiDropdown.selection) ? MOJIKUMI_CHOICES[dialogControls.mojikumiDropdown.selection.index].index : -2,
                leftIndentPt: indentAndSpacing.leftIndentPt,
                rightIndentPt: indentAndSpacing.rightIndentPt,
                spacingPt: indentAndSpacing.spacingPt
            };
        }

        /**
         * エリア内文字のフレームサイズ・行揃え・配置などを適用する
         * @param {TextFrame} areaTextFrame - 対象のエリア内文字
         * @param {Object} adjustSettings - readAdjustmentSettings() の戻り値
         * @param {boolean} forPreview - プレビュー中か
         * @returns {void}
         */
        function adjustAreaTextFrame(areaTextFrame, adjustSettings, forPreview) {
            /* 枠が変形されているなどで大きさを変えられないことがある / The frame size can refuse to change, e.g. on a transformed frame */
            try {
                areaTextFrame.spacing = adjustSettings.spacingPt;
                if (adjustSettings.widthPt !== null) { areaTextFrame.textPath.width = adjustSettings.widthPt; }
                /* 自動サイズ調整ONのあいだは枠が文字に追従するので、高さには触らない
                   While auto-size is on the frame follows the text, so the height is left alone */
                if (adjustSettings.heightPt !== null && !adjustSettings.autoSize) { areaTextFrame.textPath.height = adjustSettings.heightPt; }
            } catch (e) { }

            if (adjustSettings.shrinkFont) { shrinkFontToFit(areaTextFrame); }
            else if (adjustSettings.fitFont) { fitFontSizeToFrame(areaTextFrame); }

            /* プレビュー中は app.doScript 経由の処理を走らせない（不安定化・クラッシュ回避）
               No app.doScript during preview (it destabilizes or crashes Illustrator) */
            if (adjustSettings.autoSize && !forPreview) { enableFrameAutoSize(areaTextFrame); }

            applyParagraphSettings(areaTextFrame, adjustSettings);
            applyAreaTextFrameAlignment(areaTextFrame, adjustSettings.alignment, forPreview);
        }

        /**
         * 対象のエリア内文字すべてに現在の設定を適用する（固定ターゲットを優先して、プレビューが確実に反映されるようにする）
         * @param {boolean} forPreview - プレビュー中か
         * @returns {void}
         */
        function applyAdjustments(forPreview) {
            var adjustSettings = readAdjustmentSettings();

            var savedSelection = copyItemList(app.activeDocument.selection);
            var framesToAdjust = getTargetAreaFrames();

            for (var i = framesToAdjust.length - 1; i >= 0; i--) {
                var areaTextFrame = framesToAdjust[i];
                if (!isAreaTextFrame(areaTextFrame)) continue;
                /* 1フレームで失敗しても残りを処理できるようにする / One failing frame must not stop the rest */
                try { adjustAreaTextFrame(areaTextFrame, adjustSettings, forPreview); } catch (e) { }
            }

            if (savedSelection.length > 0) { app.activeDocument.selection = savedSelection; }
            app.redraw();
        }

        /**
         * 適用中のプレビューを取り消して未適用の状態に戻す
         * @returns {void}
         */
        function revertPreview() {
            if (!isPreviewActive) return;
            /* undo に失敗したときはプレビューが残るので、フラグを倒さず次の機会にやり直す
               A failed undo leaves the preview in place, so the flag stays up for the next attempt */
            try { app.undo(); } catch (e) { return; }
            app.redraw();
            isPreviewActive = false;

            /* undo 後は参照が無効化されるため、対象を取り直す / References go stale after an undo, so re-pick the targets */
            refreshTargetAreaFrames();
        }

        /**
         * プレビューを貼り直す（プレビューは常にON）
         * @returns {void}
         */
        function updatePreview() {
            revertPreview();
            applyAdjustments(true);
            isPreviewActive = true;
        }

        /**
         * 自動サイズ調整ONのあいだは高さ欄を使えなくする（枠が文字に追従して指定できないため）
         * @returns {void}
         */
        function updateHeightEnabled() {
            dialogControls.heightRow.enabled = !dialogControls.chkAutoSize.value;
        }

        /**
         * 連動中は右インデント欄を使えなくする（左の値をそのまま使うため）
         * @returns {void}
         */
        function updateRightIndentEnabled() {
            dialogControls.etRightIndent.enabled = !dialogControls.chkLinkIndents.value;
        }

        /**
         * オフセット欄の使用可否を更新する
         * @returns {void}
         */
        function updateSpacingEnabled() {
            dialogControls.etSpacing.enabled = dialogControls.chkSpacing.value;
            dialogControls.lblSpacingUnit.enabled = dialogControls.chkSpacing.value;
        }

        /**
         * 実寸の行送り（フォントサイズ×％）の表示を更新する
         * @returns {void}
         */
        function updateLeadingEffective() {
            var fontSize = parseFloat(dialogControls.etFontSize.text);
            var leadingPercent = parseFloat(dialogControls.etLeadingPercent.text);
            dialogControls.etLeadingEffective.text = (isNaN(fontSize) || isNaN(leadingPercent))
                ? "" : Math.round(fontSize * leadingPercent / 100 * 10) / 10;
        }

        /**
         * 行送り％の変更を反映する
         * @returns {void}
         */
        function onLeadingPercentChange() {
            updateLeadingEffective();
            updatePreview();
        }

        /**
         * 実寸の行送りの入力から行送り％を逆算する
         * @returns {void}
         */
        function onLeadingEffectiveChange() {
            var effectiveLeading = parseFloat(dialogControls.etLeadingEffective.text);
            var fontSize = parseFloat(dialogControls.etFontSize.text);
            if (isNaN(effectiveLeading) || isNaN(fontSize) || fontSize <= 0) return;
            dialogControls.etLeadingPercent.text = Math.round((effectiveLeading / fontSize) * 100 * 10) / 10;
            updatePreview();
        }

        /**
         * 対象のエリア内文字ごとにダイナミックアクションを流す
         * アクションは選択に効くので選択を付け替えながら実行し、最後に元の対象へ戻す
         * The action works on the selection, so it is swapped per frame and restored at the end
         * @param {Function} runOnFrame - フレームごとに呼ぶ処理（引数はエリア内文字）
         * @returns {void}
         */
        function runActionOnTargetFrames(runOnFrame) {
            var framesToAdjust = getTargetAreaFrames();
            var restoreSelection = copyItemList(framesToAdjust);

            for (var i = 0; i < framesToAdjust.length; i++) {
                if (isAreaTextFrame(framesToAdjust[i])) { runOnFrame(framesToAdjust[i]); }
            }

            /* 削除済みの参照が混じると選択し直せない / Re-selecting fails when a stale reference is mixed in */
            try { app.activeDocument.selection = restoreSelection; } catch (e) { }
            app.redraw();
        }

        /**
         * 現在のテキストの配置を対象フレームへ本適用する
         * @returns {void}
         */
        function applyTextAlignment() {
            var alignmentValue = getAlignmentValue(alignState.activeId);
            runActionOnTargetFrames(function (areaFrame) {
                applyAreaTextFrameAlignment(areaFrame, alignmentValue, false);
            });
        }

        /**
         * 対象フレームの自動サイズ調整を切り替える
         * @param {boolean} turnOn - ON にするか
         * @returns {void}
         */
        function setAutoSizeOnTargets(turnOn) {
            runActionOnTargetFrames(function (areaFrame) {
                if (turnOn) { enableFrameAutoSize(areaFrame); } else { disableFrameAutoSize(areaFrame); }
            });
        }

        /**
         * フォントサイズ欄の値を対象フレームに適用する
         * プレビュー分を取り消してから本適用する（そうしないと後の undo がサイズ変更を巻き戻す）
         * The preview is reverted first; otherwise a later undo would roll the size change back
         * @returns {void}
         */
        function applyFontSizeFromField() {
            var newSize = parseFloat(dialogControls.etFontSize.text) || 0;
            if (newSize <= 0) return;
            revertPreview();
            currentFontSize = newSize;
            var framesToAdjust = getTargetAreaFrames();
            for (var i = 0; i < framesToAdjust.length; i++) {
                if (!isAreaTextFrame(framesToAdjust[i])) continue;
                /* 1フレームで失敗しても残りに適用する / One failing frame must not stop the rest */
                try { framesToAdjust[i].textRange.characterAttributes.size = newSize; } catch (e) { }
            }
            updateCharsPerLineField();
            updateLeadingEffective();
            app.redraw();
            updatePreview();
        }

        /**
         * 幅の変更を検証し、文字数表示とプレビューを更新する
         * @returns {void}
         */
        function onWidthChange() {
            validateWidthField();
            updateCharsPerLineField();
            updatePreview();
        }

        /**
         * 高さの変更を検証してプレビューを更新する（以後、高さを枠へ書き戻す対象にする）
         * @returns {void}
         */
        function onHeightChange() {
            validateHeightField();
            userTouched.height = true;
            updatePreview();
        }

        /**
         * 1行の文字数から幅を逆算する
         * @returns {void}
         */
        function onCharsPerLineChange() {
            if (currentFontSize > 0) {
                var nextWidth = (((parseFloat(dialogControls.etCharsPerLine.text) || 0) * currentFontSize + getWidthAdjustmentPt()) / rulerInfo.pointsPerUnit);
                if (!isNaN(nextWidth) && isFinite(nextWidth) && nextWidth > 0) {
                    dialogControls.etWidth.text = Math.round(nextWidth * 100) / 100;
                    validateWidthField();
                }
            }
            updatePreview();
        }

        /**
         * インデント・間隔の変更を幅の再計算に回す
         * @returns {void}
         */
        function onIndentOrSpacingChange() {
            if (dialogControls.chkLinkIndents.value) { dialogControls.etRightIndent.text = dialogControls.etLeftIndent.text; }
            onWidthChange();
        }

        /**
         * L/C/R/J/F キーで行揃えを切り替える
         * @param {Window} targetDialog - キー入力を受けるダイアログ
         * @returns {void}
         */
        function addJustificationShortcuts(targetDialog) {
            targetDialog.addEventListener("keydown", function (event) {
                if (!event || !event.keyName) return;
                /* 入力欄にフォーカスがあるときは通常の文字入力を優先する / Typing into a field takes priority */
                try { if (event.target && event.target.type === "edittext") return; } catch (e0) { }
                var keyName = String(event.keyName).toUpperCase();
                for (var i = 0; i < JUSTIFY_OPTIONS.length; i++) {
                    if (JUSTIFY_OPTIONS[i].shortcut !== keyName) continue;
                    setJustification(JUSTIFY_OPTIONS[i].id);
                    event.preventDefault();
                    updatePreview();
                    return;
                }
            });
        }

        /**
         * 種別プリセットをダイアログに反映して適用する
         * @param {string} roleId - "body" / "heading" / "menu"
         * @returns {void}
         */
        function applyRolePreset(roleId) {
            var rolePreset = ROLE_PRESETS[roleId];
            if (!rolePreset) return;
            roleState.activeId = roleId;
            roleState.tabMode = rolePreset.tabMode;
            dialogControls.etLeadingPercent.text = rolePreset.leadingPercent;
            updateLeadingEffective();
            setJustification(rolePreset.justifyId);
            selectChoiceByValue(dialogControls.kinsokuDropdown, KINSOKU_CHOICES, "id", rolePreset.kinsoku);
            selectChoiceByValue(dialogControls.mojikumiDropdown, MOJIKUMI_CHOICES, "index", rolePreset.mojikumiIndex);
            alignState.activeId = rolePreset.alignId;
            /* 種別はまとめて指定するものなので、関係する項目をすべて適用対象にする / A role sets everything at once, so all of it is applied */
            userTouched.alignment = true;
            userTouched.kinsoku = true;
            userTouched.mojikumi = true;
            redrawIconButtons(dialogControls.alignButtons);

            /* テキストの配置はアクションで適用するため、プレビューを外した状態で流す / Placement runs as an action, so it goes in with the preview off */
            revertPreview();
            applyTextAlignment();
            updatePreview();
        }

        /**
         * 対象フレームの現在の幅・高さを欄に取り込む（自動サイズ調整やフィットで変わるため）
         * @returns {void}
         */
        function refreshFrameSizeFields() {
            var framesToAdjust = getTargetAreaFrames();
            for (var i = 0; i < framesToAdjust.length; i++) {
                if (!isAreaTextFrame(framesToAdjust[i])) continue;
                /* 大きさを読めないときは欄をそのままにする / Leave the fields as they are when the size cannot be read */
                try { showFrameSizeInFields(framesToAdjust[i]); } catch (e) { }
                return;
            }
        }

        /**
         * 対象フレームの現在のフォントサイズを欄に取り込む
         * @returns {void}
         */
        function refreshFontSizeField() {
            var framesToAdjust = getTargetAreaFrames();
            var newSize = 0;
            for (var i = 0; i < framesToAdjust.length; i++) {
                if (isAreaTextFrame(framesToAdjust[i])) {
                    newSize = getRepresentativeFontSize(framesToAdjust[i]);
                    break;
                }
            }
            if (newSize <= 0) return;
            currentFontSize = newSize;
            dialogControls.etFontSize.text = Math.round(newSize * 100) / 100;
            updateCharsPerLineField();
            updateLeadingEffective();
        }

        /**
         * フォントサイズの自動調整を本適用する（プレビュー分は先に取り消す）
         * @param {string} fitMode - "shrink"（文字あふれ解消）/ "fit"（枠にフィット）
         * @returns {void}
         */
        function runFontFit(fitMode) {
            /* 改行があると行数を保ったまま収められず、極端に小さいサイズになるため中止する
               With line breaks the line count cannot be kept, so the size would collapse; stop instead */
            var framesToAdjust = getTargetAreaFrames();
            for (var i = 0; i < framesToAdjust.length; i++) {
                if (isAreaTextFrame(framesToAdjust[i]) && hasLineBreak(framesToAdjust[i])) {
                    alert(getLabel("alert.lineBreakNotSupported"));
                    return;
                }
            }
            revertPreview();

            /* 自動サイズ調整がONだと枠が文字に追従してあふれないため、いったんOFFにしてから実行する。
               ONへ戻すのは applyAdjustments 内（adjustSettings.autoSize）。新しい文字サイズで枠が引き直される。
               With auto-size on the frame follows the text and never oversets, so it is switched off first.
               applyAdjustments turns it back on (adjustSettings.autoSize), redrawing the frame around the new size. */
            if (dialogControls.chkAutoSize.value) { setAutoSizeOnTargets(false); }

            fontFitMode = fitMode;
            applyAdjustments(false);
            fontFitMode = "none";

            /* 縮小・拡大した結果をダイアログにも戻す / Reflect the new size back into the dialog */
            refreshFontSizeField();
            refreshFrameSizeFields();
        }

        /**
         * ［テキストを分離...］：いま見えている調整を本適用してから、分離ダイアログを開く
         * 囲み罫はフレームの現在の大きさから作るため、プレビューを外して確定させておく
         * The rectangle is built from the frame's current size, so the preview is committed first
         * @returns {void}
         */
        function separateTargetFrames() {
            revertPreview();
            applyAdjustments(false);

            /* 選択中のエリア内文字をまとめて渡す（ロック・非表示のものは外す）
               Hand over every selected Area Type frame at once, minus the locked and hidden ones */
            var currentTargets = getTargetAreaFrames();
            var framesToSeparate = [];
            for (var i = 0; i < currentTargets.length; i++) {
                if (isSeparableAreaTextFrame(currentTargets[i])) { framesToSeparate.push(currentTargets[i]); }
            }
            if (!framesToSeparate.length) {
                alert(getLabel("alert.separateFailed"));
                updatePreview();
                return;
            }

            if (showSeparateTextDialog(doc, framesToSeparate, readAdjustmentSettings())) {
                /* 分離するとエリア内文字が無くなるので、調整ダイアログも閉じる / Separating removes the Area Type, so the adjust dialog closes too */
                dialogControls.window.close(1);
            } else {
                /* 分離しなかったときはプレビューを貼り直して調整を続ける / Nothing separated: re-apply the preview and keep adjusting */
                updatePreview();
            }
        }

        /**
         * 各コントロールにイベントを付ける
         * @returns {void}
         */
        function bindAdjustDialogEvents() {
            bindExclusiveRadios(dialogControls.roleRadios, function (clickedRadio) {
                if (clickedRadio === dialogControls.radRoleBody) { applyRolePreset("body"); }
                else if (clickedRadio === dialogControls.radRoleHeading) { applyRolePreset("heading"); }
                else { applyRolePreset("menu"); }
            });
            for (var buttonIndex = 0; buttonIndex < dialogControls.justifyButtons.length; buttonIndex++) {
                dialogControls.justifyButtons[buttonIndex].onClick = function () {
                    setJustification(this.justifyId);
                    updatePreview();
                };
            }
            for (var alignButtonIndex = 0; alignButtonIndex < dialogControls.alignButtons.length; alignButtonIndex++) {
                dialogControls.alignButtons[alignButtonIndex].onClick = function () {
                    alignState.activeId = this.alignId;
                    userTouched.alignment = true;
                    redrawIconButtons(dialogControls.alignButtons);
                    /* プレビューを外してから本適用し、あらためてプレビューを貼り直す / Commit with the preview off, then re-apply the preview */
                    revertPreview();
                    applyTextAlignment();
                    updatePreview();
                };
            }
            addJustificationShortcuts(dialogControls.window);

            dialogControls.btnShrinkToFit.onClick = function () { runFontFit("shrink"); };
            dialogControls.btnFitFontSize.onClick = function () { runFontFit("fit"); };
            dialogControls.kinsokuDropdown.onChange = function () {
                userTouched.kinsoku = true;
                updatePreview();
            };
            dialogControls.mojikumiDropdown.onChange = function () {
                userTouched.mojikumi = true;
                updatePreview();
            };
            dialogControls.chkAutoSize.onClick = function () {
                updateHeightEnabled();
                /* プレビューを外してから ON/OFF を本適用し、あらためてプレビューを貼り直す / Commit ON/OFF with the preview off, then re-apply it */
                revertPreview();
                setAutoSizeOnTargets(dialogControls.chkAutoSize.value);
                /* 枠の高さが変わるので、欄の値も追従させる（古い値でプレビューが枠を戻さないように）
                   The frame height changes, so the fields follow (the preview must not push the old size back) */
                refreshFrameSizeFields();
                updatePreview();
            };
            dialogControls.chkSpacing.onClick = function () {
                if (dialogControls.chkSpacing.value) { dialogControls.etSpacing.text = "1"; }
                updateSpacingEnabled();
                onIndentOrSpacingChange();
            };
            dialogControls.chkLinkIndents.onClick = function () {
                /* 連動中は右を左に追従させるので、右の欄は触れないようにする / While linked the right follows the left, so its field is locked */
                if (dialogControls.chkLinkIndents.value) { dialogControls.etRightIndent.text = dialogControls.etLeftIndent.text; }
                updateRightIndentEnabled();
                onIndentOrSpacingChange();
            };

            dialogControls.etFontSize.onChange = applyFontSizeFromField;
            dialogControls.etLeadingPercent.onChange = onLeadingPercentChange;
            dialogControls.etLeadingEffective.onChange = onLeadingEffectiveChange;
            dialogControls.etSpacing.onChange = onIndentOrSpacingChange;
            dialogControls.etWidth.onChange = onWidthChange;
            dialogControls.etHeight.onChange = onHeightChange;
            dialogControls.etCharsPerLine.onChange = onCharsPerLineChange;
            dialogControls.etLeftIndent.onChange = onIndentOrSpacingChange;
            dialogControls.etRightIndent.onChange = onIndentOrSpacingChange;
            changeValueByArrowKey(dialogControls.etFontSize, false, applyFontSizeFromField);
            changeValueByArrowKey(dialogControls.etLeadingPercent, false, onLeadingPercentChange);
            changeValueByArrowKey(dialogControls.etLeadingEffective, false, onLeadingEffectiveChange);
            changeValueByArrowKey(dialogControls.etSpacing, false, onIndentOrSpacingChange);
            changeValueByArrowKey(dialogControls.etWidth, false, onWidthChange);
            changeValueByArrowKey(dialogControls.etHeight, false, onHeightChange);
            changeValueByArrowKey(dialogControls.etCharsPerLine, false, onCharsPerLineChange);
            changeValueByArrowKey(dialogControls.etLeftIndent, false, onIndentOrSpacingChange);
            changeValueByArrowKey(dialogControls.etRightIndent, false, onIndentOrSpacingChange);

            dialogControls.btnSeparateText.onClick = separateTargetFrames;
            dialogControls.btnRun.onClick = function () {
                revertPreview();
                applyAdjustments(false);
                dialogControls.window.close(1);
            };
            dialogControls.btnCancelAdjust.onClick = function () {
                revertPreview();
                dialogControls.window.close(0);
            };
        }

        bindAdjustDialogEvents();

        /* 初期値の読み込み / Load the initial values */
        if (initialFrame) { loadValuesFromFrame(initialFrame); }

        updateSpacingEnabled();
        updateHeightEnabled();
        updateRightIndentEnabled();

        /* 調整ダイアログを開いた時点からプレビューを反映する / The preview is on from the moment the dialog opens */
        updatePreview();

        /* ［文字あふれ解消］［枠にフィット］だけ、上下2pxずつ詰めて小ぶりにする
           Make the two font-fit buttons a little shorter (2px off the top and bottom) */
        dialogControls.window.layout.layout(true);
        trimButtonHeight(dialogControls.btnShrinkToFit, 4);
        trimButtonHeight(dialogControls.btnFitFontSize, 4);

        dialogControls.window.show();
    }

    // ============================================================
    // エントリポイント / Entry point
    //   ダイナミックアクションを読み込み、終了時に必ずアンロードする
    //   Load dynamic actions and always unload them on exit
    // ============================================================
    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }

    var doc = app.activeDocument;
    var docSelection = doc.selection;

    if (!docSelection || docSelection.length === 0) {
        alert(getLabel("alert.selectText"));
        return;
    }

    loadDynamicActions();
    try {
        var selectionKinds = detectSelectionKinds(docSelection);

        if (selectionKinds.hasPointText || selectionKinds.hasPath) {
            /* ポイント文字 または 図形 → 変換ダイアログ / Point text or shape → convert dialog */
            showConvertDialog(doc, docSelection);
        } else if (selectionKinds.hasAreaText) {
            /* エリア内文字のみ → 調整ダイアログ / Area text only → adjust dialog */
            var firstAreaFrame = null;
            for (var i = 0; i < docSelection.length; i++) {
                if (isAreaTextFrame(docSelection[i])) {
                    firstAreaFrame = docSelection[i];
                    break;
                }
            }
            showAdjustDialog(doc, firstAreaFrame, null, null);
        } else {
            alert(getLabel("alert.selectText"));
        }
    } finally {
        unloadDynamicActions();
    }

})();
