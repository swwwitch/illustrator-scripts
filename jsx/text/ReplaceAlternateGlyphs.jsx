#target illustrator
#targetengine "ReplaceAlternateGlyphsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ReplaceAlternateGlyphs
Name=選択した文字の字形と異体字を切り換える
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.2.0
Release-Date=2026-10-17
Description-BEGIN
選択している文字の字形（JIS78・JIS90・旧字体など）を切り換えたり、異体字に置き換えたりします。

Switches the selected characters to another glyph form (JIS78, JIS90, Traditional, etc.) or replaces them with a variant character.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

選択している文字の字形（JIS78・JIS90・旧字体など）を切り換えたり、異体字に置き換えたりします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceAlternateGlyphs.md

### Overview

Switches the selected characters to another glyph form (JIS78, JIS90, Traditional, etc.) or replaces them with a variant character.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceAlternateGlyphs.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ReplaceAlternateGlyphs";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-17";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "選択した文字の字形と異体字を切り換える"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceAlternateGlyphs.md"; /* README（日本語） */
var SCRIPT_README_EN  = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceAlternateGlyphs.md"; /* README (English) */
var SCRIPT_SOURCE_URL = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/ReplaceAlternateGlyphs.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 異体字の組（1つの文字列が互いに異体字の文字の組。先頭は常用の字形）。
       同じ文字を含む組は実行時に1つにまとめるので、行を足すだけで増やせる。
       互換漢字（見た目は同じで別の符号位置）は \uXXXX で書く
       Variant groups: each string lists characters that are variants of one another (the common form first).
       Groups sharing a character are merged at run time. Compatibility ideographs are written as \uXXXX */
    var ITAIJI_GROUPS = [
        /* 人名・地名に多いもの / common in names */
        "高髙", "崎﨑嵜碕", "辺邊邉", "島嶋嶌", "富冨", "沢澤", "浜濱濵", "斉齊", "斎齋",
        "桜櫻", "国國圀", "竜龍", "徳德", "瀬瀨", "曽曾", "槙槇", "栄榮", "円圓", "広廣",
        "恵惠", "条條", "黒黑", "館舘館", "顔顏", "真眞", "関關", "峰峯", "鉄鐵銕", "万萬",
        "吉𠮷", "塚塚", "隆隆", "晴晴", "猪猪", "朗朗",
        "神神", "福福", "祥祥", "靖靖", "精精", "諸諸", "都都",
        /* 新字体と旧字体 / shinjitai and kyujitai */
        "亜亞", "悪惡", "圧壓", "囲圍", "医醫", "営營", "駅驛", "応應", "仮假", "価價",
        "画畫", "絵繪", "会會", "拡擴", "覚覺", "学學", "楽樂", "勧勸", "巻卷", "観觀",
        "気氣", "帰歸", "旧舊", "拠據", "挙擧", "峡峽", "狭狹", "挟挾", "暁曉", "区區",
        "駆驅", "勲勳", "薫薰", "径徑", "茎莖", "経經", "渓溪", "継繼", "軽輕", "鶏鷄",
        "芸藝", "撃擊", "剣劍", "県縣", "険險", "検檢", "圏圈", "権權", "献獻", "研硏",
        "顕顯", "験驗", "厳嚴", "巌巖", "効效", "恒恆", "鉱鑛", "号號", "済濟", "剤劑",
        "雑雜", "蚕蠶", "惨慘", "賛贊", "残殘", "糸絲", "歯齒", "児兒", "辞辭", "湿濕",
        "実實", "写寫", "舎舍", "釈釋", "寿壽", "収收", "従從", "渋澁", "獣獸", "縦縱",
        "粛肅", "処處", "将將", "称稱", "焼燒", "証證", "乗乘", "浄淨", "状狀", "畳疊",
        "譲讓", "嬢孃", "醸釀", "触觸", "図圖", "粋粹", "酔醉", "穂穗", "随隨", "枢樞",
        "数數", "声聲", "静靜", "窃竊", "摂攝", "戦戰", "専專", "銭錢", "浅淺", "践踐",
        "潜潛", "繊纖", "禅禪", "双雙", "壮壯", "争爭", "荘莊", "捜搜", "挿插", "巣巢",
        "総總", "装裝", "騒騷", "蔵藏", "臓臟", "属屬", "続續", "体體", "台臺", "堕墮",
        "対對", "滝瀧", "択擇", "担擔", "胆膽", "団團", "断斷", "弾彈", "遅遲", "痴癡",
        "昼晝", "虫蟲", "鋳鑄", "庁廳", "聴聽", "鎮鎭", "逓遞", "転轉", "点點", "伝傳",
        "灯燈", "当當", "党黨", "盗盜", "闘鬭", "独獨", "読讀", "届屆", "弐貳", "悩惱",
        "脳腦", "覇霸", "廃廢", "売賣", "髪髮", "発發", "抜拔", "蛮蠻", "秘祕", "瓶甁",
        "払拂", "仏佛", "併倂", "並竝", "変變", "弁辨瓣辯", "舗舖", "歩步", "宝寶", "豊豐",
        "麦麥", "黙默", "訳譯", "薬藥", "与與", "予豫", "余餘", "誉譽", "揺搖", "様樣",
        "謡謠", "来來", "乱亂", "覧覽", "両兩", "猟獵", "緑綠", "塁壘", "礼禮", "励勵",
        "霊靈", "齢齡", "暦曆", "歴歷", "恋戀", "炉爐", "労勞", "楼樓", "湾灣",
        /* 互換漢字（JIS X 0213） / compatibility ideographs */
        "侮侮", "僧僧", "免免", "勉勉", "勤勤", "卑卑", "喝喝",
        "嘆嘆", "器器", "塀塀", "墨墨", "層層", "悔悔", "慨慨",
        "憎憎", "懲懲", "敏敏", "既既", "暑暑", "梅梅", "海海",
        "渚渚", "漢漢", "煮煮", "琢琢", "碑碑", "社社", "祉祉",
        "祈祈", "祐祐", "祖祖", "祝祝", "禍禍", "禎禎", "穀穀",
        "突突", "節節", "練練", "繁繁", "署署", "者者", "臭臭",
        "著著", "褐褐", "視視", "謁謁", "謹謹", "賓賓", "贈贈",
        "逸逸", "難難", "響響", "頻頻",
        /* 2010年の常用漢字の追加分ほか / others */
        "ー\uFF70", "籠篭", "蘆芦", "蝋蠟", "繋繫", "撹攪", "嘘噓", "叱𠮟", "剥剝", "頬頰", "填塡"
    ];

    var GLYPH_FORM_CHECK_MAX      = 10;       /* 字形の対応を調べる文字の数（先頭から。多いと開くのが遅くなる） / characters checked for supported forms, from the start */

    /* ［一覧を作成］の書き出し先と大きさ / Output of Create List */
    var GLYPH_LIST_LAYER_NAME     = "// glyphs"; /* 一覧を作るレイヤー（無ければ最前面に作る） / layer for the list (created on top if missing) */
    var DELETE_LIST_LAYER_DEFAULT = true;     /* ［レイヤーを削除してから作成］を選んだ状態で開く / open with "delete the layer, then create" chosen */
    var GLYPH_LIST_CODE_RATIO     = 0.25;     /* 符号位置の大きさ（見本の字に対する比率。字形の名前はその3/4） / code point size relative to the sample (form names are 3/4 of it) */
    var GLYPH_LIST_LABEL_MIN_SIZE = 6;        /* 符号位置・字形の名前の最小の大きさ（pt） / minimum label size in points */
    var GLYPH_LIST_IVS_MAX        = 64;       /* 試す異体字セレクタの数（U+E0100 から） / number of variation selectors tried, from U+E0100 */
    var GLYPH_LIST_IVS_MISS_LIMIT = 3;        /* 新しい形が続けて出なければ打ち切る回数 / stop after this many selectors in a row add no new shape */
    var GLYPH_LIST_BACK_OPACITY   = 30;       /* 背面の黒い長方形の不透明度（%） / opacity of the black backdrop, in percent */
    var GLYPH_LIST_BACK_PADDING   = 1;        /* 背面の長方形の余白（1行目の見本の字に対する比率） / backdrop padding relative to the first sample size */
    var GLYPH_LIST_COLUMNS        = 5;        /* 一覧の1行に並べる数（超えたら折り返す） / samples per line before wrapping */
    var GLYPH_LIST_FIT_PERCENT    = 65;       /* 作った一覧がウインドウに占める割合（%） / share of the window the list fills, in percent */
    /* 符号位置・字形の名前のフォント（PostScript 名。前から順に探し、無ければ元の文字のフォント）
       Font for code points and form names (PostScript names, tried in order; falls back to the source font) */
    var GLYPH_LIST_LABEL_FONT_NAMES = ["HiraginoSans-W3", "HiraKakuProN-W3", "HiraKakuPro-W3"];

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
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
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    var BUTTON_WIDTH              = 90;          /* ボタンの幅 / button width */
    var VARIANT_LIST_SIZE         = [170, 240];  /* 異体字の一覧の大きさ [幅, 高さ] / variant list size */
    var LIST_FONT_SIZE            = 18;          /* 一覧の文字の大きさ（細部の違いを見分けやすく） / list font size, large enough to tell variants apart */
    var SELECTION_CODES_MARGIN    = 5;           /* 上段の符号位置の下に足す余白 / extra space below the code points at the top */
    var SELECTION_CODES_MAX       = 5;           /* ダイアログボックスの上に符号位置を出す文字の数 / characters whose code points are shown at the top */
    var ROW_LABEL_WIDTH_EM        = 6;           /* 一覧の行の見出しの幅（符号位置の大きさの何倍か） / width of the list's row headings, in code point sizes */
    var DIALOG_RESULT_CREATE_LIST = 3;           /* ［一覧を作成］で閉じたときの show() の戻り値 / show() result when closed by Create List */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        dialog: {
            title: { ja: "字形・異体字の切り換え", en: "Switch Glyph Forms and Variants" }
        },
        panel: {
            glyph: { ja: "字形", en: "Glyph Form" },
            variants: { ja: "異体字", en: "Variant Characters" },
            list: { ja: "一覧", en: "List" }
        },
        radio: {
            glyphDefault: { ja: "標準字形", en: "Default Form" },
            glyphJis78: { ja: "JIS78字形", en: "JIS78 Forms" },
            glyphJis83: { ja: "JIS83字形", en: "JIS83 Forms" },
            glyphJis90: { ja: "JIS90字形", en: "JIS90 Forms" },
            glyphJis04: { ja: "JIS04字形", en: "JIS04 Forms" },
            glyphTraditional: { ja: "旧字体", en: "Traditional Forms" },
            glyphExpert: { ja: "エキスパート字形", en: "Expert Forms" },
            glyphFull: { ja: "等幅全角字形", en: "Full Widths" },
            glyphHalf: { ja: "等幅半角字形", en: "Half Widths" },
            glyphThird: { ja: "等幅三分字形", en: "Third Widths" },
            glyphQuarter: { ja: "等幅四分字形", en: "Quarter Widths" },
            glyphProportional: { ja: "プロポーショナル字形", en: "Proportional Widths" },
            listAddToLayer: { ja: "［{layer}］レイヤーに作成", en: "Create on the [{layer}] layer" },
            listReplaceLayer: { ja: "［{layer}］レイヤーを削除してから作成", en: "Delete the [{layer}] layer, then create" }
        },
        status: {
            selectedCharacters: { ja: "選択した文字", en: "Selected" },
            moreCharacters: { ja: "ほか{count}字", en: "+{count} more" },
            noVariants: { ja: "登録された異体字はありません", en: "No registered variants" }
        },
        listHeading: {
            horizontal: { ja: "横組み用", en: "Horizontal" },
            vertical: { ja: "縦組み用", en: "Vertical" }
        },
        button: {
            createList: { ja: "一覧を作成", en: "Create List" },
            showIvs: { ja: "IVSを表示", en: "Show IVS" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            selectionCodes: {
                ja: "選択している文字と Unicode の符号位置です（異体字に置き換えると更新します）",
                en: "The selected characters and their Unicode code points (updated after a variant replacement)"
            },
            glyph: {
                ja: "選択している文字の字形をすぐに変えます。フォントにその字形が無いものはディム表示になります（キャンセルで元に戻ります）",
                en: "Changes the glyph form of the selected characters right away. Forms the font lacks are dimmed (Cancel restores them)"
            },
            listAddToLayer: { ja: "前回までの一覧を残したまま、同じレイヤーに作ります", en: "Adds the list to the same layer, keeping earlier lists" },
            listReplaceLayer: {
                ja: "前回の一覧のレイヤーを消してから作り直します（選択中の文字がそのレイヤーにあるときは消しません）",
                en: "Removes the previous list layer before making a new one (skipped when the selected text is on that layer)"
            },
            createList: {
                ja: "ダイアログボックスを閉じて、選択した文字の字形と異体字のうち形の異なるものを、符号位置・字形の名前とともにレイヤー「{layer}」に並べ、選択して表示します（字形・異体字の切り換えは残ります）",
                en: "Closes the dialog box, lays out the distinct glyph forms and variants of the selected characters with code points and form names on the layer \"{layer}\", then selects and shows them (glyph and variant changes are kept)"
            },
            showIvs: {
                ja: "選択した文字の IVS（異体字セレクタで呼び出す異体字）を調べて、上の一覧に足します。一覧の字は UI のフォントで表示されるので、ドキュメントのフォントの形とは異なることがあります",
                en: "Looks up the IVS glyphs (variants called with variation selectors) of the selected characters and adds them to the list above. The list uses the UI font, so shapes may differ from the document's font"
            },
            variants: {
                ja: "クリックした異体字（別の文字）に、選択している同じ文字をすべてすぐに置き換えます（キャンセルで元に戻ります）",
                en: "Replaces every selected occurrence of that character with the clicked variant right away (Cancel restores them)"
            }
        },
        alert: {
            noSelection: { ja: "テキストを選択してください。", en: "Please select text." },
            notFound: { ja: "選択に文字がありません。", en: "The selection contains no characters." },
            noIvs: { ja: "このフォントでは IVS の異体字が見つかりませんでした。", en: "No IVS variants were found in this font." }
        }
    };

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    // =========================================
    // 文字の収集 / Collecting characters
    // =========================================

    /**
     * 選択から調べる対象を集める。文字選択中はその範囲、オブジェクト選択ならテキストフレームごと（連結テキストはストーリー全体を1つ）
     * 書き換えで範囲の参照がずれないよう、文字はフレームかストーリーの characters を番号で読む
     * @param {Document} doc - 対象のドキュメント
     * @returns {Array<{owner: (TextFrame|Story), start: number, length: number, text: string}>} 文字を持つオブジェクト・調べる番号の範囲・その文字列
     */
    function collectTextTargets(doc) {
        var selection = doc.selection;
        var textTargets = [];
        /* 文字選択中は selection が TextRange（[0] が無い） / While text is selected, selection is a TextRange without [0] */
        if (selection && selection.typename === "TextRange") {
            if (selection.length > 0) {
                textTargets.push({ owner: selection.story, start: selection.characterOffset - 1, length: selection.length, text: selection.contents });
            }
            return textTargets;
        }
        var textFrames = collectSelectionTextFrames(selection, { skipLocked: true, skipHidden: true });
        var storyHeadFrames = [];
        for (var i = 0; i < textFrames.length; i++) {
            var textStory = textFrames[i].story;
            if (textStory.textFrames.length <= 1) {
                textTargets.push({ owner: textFrames[i], start: 0, length: textFrames[i].characters.length, text: textFrames[i].contents });
                continue;
            }
            /* 連結テキストは同じストーリーを2度数えない / count a threaded story only once */
            var headFrame = textStory.textFrames[0];
            if (indexOfItem(storyHeadFrames, headFrame) >= 0) continue;
            storyHeadFrames.push(headFrame);
            textTargets.push({ owner: textStory, start: 0, length: textStory.characters.length, text: readStoryText(textStory) });
        }
        return textTargets;
    }

    /**
     * ストーリー全体の文字列を返す（Story に contents は無いので textRange から読み、読めなければ1文字ずつつなぐ）
     * @param {Story} textStory - 対象のストーリー
     * @returns {string} 文字列
     */
    function readStoryText(textStory) {
        try {
            return textStory.textRange.contents;
        } catch (e) {
            var storyCharacters = textStory.characters;
            var storyText = "";
            for (var i = 0; i < storyCharacters.length; i++) storyText += storyCharacters[i].contents;
            return storyText;
        }
    }

    /**
     * 配列の中の位置を返す（DOM の参照は === で比べられる）
     * @param {Array} items - 配列
     * @param {*} targetItem - 探す値
     * @returns {number} 位置（無ければ -1）
     */
    function indexOfItem(items, targetItem) {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === targetItem) return i;
        }
        return -1;
    }

    /**
     * 文字列を1文字（サロゲートペアは2つで1文字）ずつの配列に分ける
     * @param {string} text - 文字列
     * @returns {string[]} 1文字ずつの配列
     */
    function splitCodePoints(text) {
        var codePoints = [];
        for (var i = 0; i < text.length; i++) {
            var code = text.charCodeAt(i);
            if (code >= 0xD800 && code <= 0xDBFF && i + 1 < text.length) {
                codePoints.push(text.substr(i, 2));
                i++;
                continue;
            }
            codePoints.push(text.charAt(i));
        }
        return codePoints;
    }

    /**
     * 1文字の符号位置を「U+XXXX」の形で返す
     * @param {string} codePoint - 1文字（サロゲートペア可）
     * @returns {string} 符号位置の表記
     */
    function formatCodePoint(codePoint) {
        var code = codePoint.charCodeAt(0);
        if (codePoint.length === 2) code = (code - 0xD800) * 0x400 + (codePoint.charCodeAt(1) - 0xDC00) + 0x10000;
        var hex = code.toString(16).toUpperCase();
        while (hex.length < 4) hex = "0" + hex;
        return "U+" + hex;
    }

    /**
     * 文字列の符号位置を「U+9089 U+E0102」のように1文字ずつ並べて返す（IVS は2つになる）
     * @param {string} text - 文字列
     * @returns {string} 符号位置の表記
     */
    function formatCodePoints(text) {
        var codePoints = splitCodePoints(text);
        for (var i = 0; i < codePoints.length; i++) codePoints[i] = formatCodePoint(codePoints[i]);
        return codePoints.join(" ");
    }

    /**
     * 文字列から異体字セレクタ（U+FE00〜U+FE0F・U+E0100〜U+E01EF）を取り除く
     * @param {string} text - 文字列
     * @returns {string} セレクタを除いた文字列
     */
    function stripVariationSelectors(text) {
        return text.replace(/\uDB40[\uDD00-\uDDEF]|[\uFE00-\uFE0F]/g, "");
    }

    /**
     * 対象にする文字か（空白・改行・タブ以外）
     * @param {string} codePoint - 1文字（サロゲートペア可）
     * @returns {boolean} 対象なら true
     */
    function isTargetCharacter(codePoint) {
        return !/^[\s　\u0003]$/.test(codePoint);
    }

    /**
     * ITAIJI_GROUPS から「文字 → 異体字の組」の表を作る。同じ文字を含む組は1つにまとめる
     * @param {string[]} variantGroups - 異体字の組
     * @returns {Object} 文字をキー、組の文字の配列を値にした表
     */
    function buildVariantMap(variantGroups) {
        var variantMap = {};
        for (var i = 0; i < variantGroups.length; i++) {
            var members = splitCodePoints(variantGroups[i]);
            var mergedGroup = [];
            for (var j = 0; j < members.length; j++) {
                var existingGroup = variantMap.hasOwnProperty(members[j]) ? variantMap[members[j]] : [members[j]];
                for (var k = 0; k < existingGroup.length; k++) {
                    if (indexOfItem(mergedGroup, existingGroup[k]) < 0) mergedGroup.push(existingGroup[k]);
                }
            }
            for (var m = 0; m < mergedGroup.length; m++) variantMap[mergedGroup[m]] = mergedGroup;
        }
        return variantMap;
    }

    /**
     * 対象の文字を前から順に調べ、元の文字ごとにまとめる
     * @param {{owner: (TextFrame|Story), start: number, length: number, text: string}} textTarget - collectTextTargets() の要素
     * @param {Object} variantMap - buildVariantMap() の表
     * @param {Object} collectedCharacters - { entries: 文字ごとのまとまり, entryByCharacter: 文字 → まとまり, places: すべての箇所 }
     * @returns {void}
     */
    function collectCharacterPlaces(textTarget, variantMap, collectedCharacters) {
        var targetText = textTarget.text;
        /* contents と characters の数が合わなければ1文字ずつ読み直す / re-read character by character if the counts disagree */
        if (targetText.length !== textTarget.length) {
            targetText = "";
            for (var k = 0; k < textTarget.length; k++) targetText += textTarget.owner.characters[textTarget.start + k].contents;
        }
        var codePoints = splitCodePoints(targetText);
        var offset = 0;
        for (var i = 0; i < codePoints.length; i++) {
            var codePoint = codePoints[i];
            var placeStart = textTarget.start + offset;
            offset += codePoint.length;
            if (!isTargetCharacter(codePoint)) continue;
            var characterEntry = collectedCharacters.entryByCharacter.hasOwnProperty(codePoint) ? collectedCharacters.entryByCharacter[codePoint] : null;
            if (!characterEntry) {
                characterEntry = {
                    original: codePoint,
                    current: codePoint,
                    variants: variantMap.hasOwnProperty(codePoint) ? variantMap[codePoint] : [],
                    places: []
                };
                collectedCharacters.entryByCharacter[codePoint] = characterEntry;
                collectedCharacters.entries.push(characterEntry);
            }
            var characterPlace = { owner: textTarget.owner, start: placeStart, length: codePoint.length };
            characterEntry.places.push(characterPlace);
            collectedCharacters.places.push(characterPlace);
        }
    }

    // =========================================
    // 文字の範囲と選択の控え / Character spans and saved selection
    // =========================================

    /**
     * 番号と文字数から文字の範囲を作る（1文字目の length を伸ばす。characterOffset への代入は使えない）
     * @param {(TextFrame|Story)} owner - 文字を持つオブジェクト
     * @param {number} start - 先頭の文字の番号（0始まり）
     * @param {number} length - 文字数
     * @returns {TextRange} 文字の範囲
     */
    function getCharacterSpan(owner, start, length) {
        var characterSpan = owner.characters[start];
        characterSpan.length = length;
        return characterSpan;
    }

    /**
     * 箇所の文字の範囲を返す
     * @param {{owner: (TextFrame|Story), start: number, length: number}} characterPlace - 文字の箇所
     * @returns {TextRange} 箇所の範囲
     */
    function getPlaceSpan(characterPlace) {
        return getCharacterSpan(characterPlace.owner, characterPlace.start, characterPlace.length);
    }

    /**
     * 今の選択を控える（文字選択はストーリーと番号、オブジェクト選択は配列）
     * @param {Document} doc - 対象のドキュメント
     * @returns {Object} restoreSelection() に渡す控え
     */
    function saveSelection(doc) {
        var selection = doc.selection;
        if (selection && selection.typename === "TextRange") {
            return { story: selection.story, start: selection.characterOffset - 1, length: selection.length };
        }
        var selectedItems = [];
        for (var i = 0; selection && i < selection.length; i++) selectedItems.push(selection[i]);
        return { items: selectedItems };
    }

    /**
     * 控えた選択に戻す
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} savedSelection - saveSelection() の控え
     * @returns {void}
     */
    function restoreSelection(doc, savedSelection) {
        doc.selection = null;
        if (!savedSelection.story) {
            doc.selection = savedSelection.items;
            return;
        }
        if (savedSelection.length > 0) getCharacterSpan(savedSelection.story, savedSelection.start, savedSelection.length).select();
    }

    // =========================================
    // 字形と異体字の書き換え / Switching glyph forms and variants
    // =========================================

    /* 選べる字形（labelKey は LABELS のパス） / Selectable glyph forms (labelKey is a LABELS path) */
    var GLYPH_FORMS = [
        { labelKey: "radio.glyphDefault", form: AlternateGlyphsForm.DEFAULTFORM },
        { labelKey: "radio.glyphJis78", form: AlternateGlyphsForm.JIS78FORM },
        { labelKey: "radio.glyphJis83", form: AlternateGlyphsForm.JIS83FORM },
        { labelKey: "radio.glyphJis90", form: AlternateGlyphsForm.JIS90FORM },
        { labelKey: "radio.glyphJis04", form: AlternateGlyphsForm.JIS04FORM },
        { labelKey: "radio.glyphTraditional", form: AlternateGlyphsForm.TRADITIONAL },
        { labelKey: "radio.glyphExpert", form: AlternateGlyphsForm.EXPERT },
        { labelKey: "radio.glyphFull", form: AlternateGlyphsForm.FULLWIDTH },
        { labelKey: "radio.glyphHalf", form: AlternateGlyphsForm.HALFWIDTH },
        { labelKey: "radio.glyphThird", form: AlternateGlyphsForm.THIRDWIDTH },
        { labelKey: "radio.glyphQuarter", form: AlternateGlyphsForm.QUARTERWIDTH }
    ];

    /**
     * ［一覧を作成］で試す字形を返す（GLYPH_FORMS に、ラジオボタンには出さないプロポーショナル字形を足す）
     * @returns {Array<{labelKey: string, form: *}>} 試す字形（先頭は標準字形）
     */
    function getListGlyphForms() {
        var listForms = GLYPH_FORMS.slice(0);
        /* 古い環境で列挙に無ければ足さない（無い値は undefined になる） / skip when the enumeration lacks it (it reads as undefined) */
        if (AlternateGlyphsForm.PROPORTIONALWIDTH !== undefined) {
            listForms.push({ labelKey: "radio.glyphProportional", form: AlternateGlyphsForm.PROPORTIONALWIDTH });
        }
        return listForms;
    }

    /**
     * 箇所の字形を読む（フォントによっては読めないので、そのときは null）
     * @param {Object} characterPlace - 文字の箇所
     * @returns {*} AlternateGlyphsForm の値または null
     */
    function readPlaceGlyphForm(characterPlace) {
        try {
            return characterPlace.owner.characters[characterPlace.start].characterAttributes.alternateGlyphs;
        } catch (e) {
            return null;
        }
    }

    /**
     * まとまりの全箇所の元の字形を控え、共通の字形を entry.form に入れる（2回目からは何もしない）
     * @param {Object} characterEntry - 文字のまとまり
     * @returns {void}
     */
    function readEntryGlyphForms(characterEntry) {
        if (characterEntry.hasOwnProperty("form")) return;
        var commonForm;
        for (var i = 0; i < characterEntry.places.length; i++) {
            var characterPlace = characterEntry.places[i];
            characterPlace.originalForm = readPlaceGlyphForm(characterPlace);
            if (i === 0) commonForm = characterPlace.originalForm;
            else if (commonForm !== characterPlace.originalForm) commonForm = null; /* 箇所ごとに異なる / mixed */
        }
        characterEntry.form = commonForm;
        characterEntry.originalCommonForm = commonForm;
    }

    /**
     * まとまりの全箇所の字形を変える
     * @param {Object} characterEntry - 文字のまとまり
     * @param {*} glyphForm - AlternateGlyphsForm の値
     * @returns {void}
     */
    function applyEntryGlyphForm(characterEntry, glyphForm) {
        readEntryGlyphForms(characterEntry);
        if (characterEntry.form === glyphForm) return;
        for (var i = 0; i < characterEntry.places.length; i++) {
            getPlaceSpan(characterEntry.places[i]).characterAttributes.alternateGlyphs = glyphForm;
        }
        characterEntry.form = glyphForm;
    }

    /**
     * 1箇所の文字を書き換える（書式はその文字のものが残る）。文字数が変わったら、同じオブジェクトの後ろの箇所の番号をずらす
     * @param {Object} characterPlace - 書き換える箇所
     * @param {string} newCharacter - 書き換え後の1文字（サロゲートペア可）
     * @param {Array<Object>} allPlaces - すべての箇所
     * @returns {void}
     */
    function rewritePlace(characterPlace, newCharacter, allPlaces) {
        getPlaceSpan(characterPlace).contents = newCharacter;
        var lengthDelta = newCharacter.length - characterPlace.length;
        characterPlace.length = newCharacter.length;
        if (lengthDelta === 0) return;
        for (var i = 0; i < allPlaces.length; i++) {
            if (allPlaces[i].owner === characterPlace.owner && allPlaces[i].start > characterPlace.start) allPlaces[i].start += lengthDelta;
        }
    }

    /**
     * まとまりの全箇所を別の文字（異体字）に置き換える
     * @param {Object} characterEntry - 文字のまとまり
     * @param {string} newCharacter - 置き換え後の1文字
     * @param {Array<Object>} allPlaces - すべての箇所
     * @returns {void}
     */
    function applyEntryVariant(characterEntry, newCharacter, allPlaces) {
        if (characterEntry.current === newCharacter) return;
        for (var i = characterEntry.places.length - 1; i >= 0; i--) rewritePlace(characterEntry.places[i], newCharacter, allPlaces);
        characterEntry.current = newCharacter;
    }

    /**
     * すべてのまとまりの字形を変える
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @param {*} glyphForm - AlternateGlyphsForm の値
     * @returns {void}
     */
    function applyGlyphFormToAll(characterEntries, glyphForm) {
        for (var i = 0; i < characterEntries.length; i++) applyEntryGlyphForm(characterEntries[i], glyphForm);
    }

    /**
     * すべてのまとまりに共通の字形を返す
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @returns {*} AlternateGlyphsForm の値（まとまりごとに異なるか読めなければ null）
     */
    function getCommonGlyphForm(characterEntries) {
        var commonForm = null;
        for (var i = 0; i < characterEntries.length; i++) {
            readEntryGlyphForms(characterEntries[i]);
            if (i === 0) commonForm = characterEntries[i].form;
            else if (commonForm !== characterEntries[i].form) return null;
        }
        return commonForm;
    }

    /**
     * キャンセルしたとき、すべてのまとまりを元の文字と字形に戻す
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @param {Array<Object>} allPlaces - すべての箇所
     * @returns {void}
     */
    function revertAllEntries(characterEntries, allPlaces) {
        for (var i = 0; i < characterEntries.length; i++) {
            var characterEntry = characterEntries[i];
            applyEntryVariant(characterEntry, characterEntry.original, allPlaces);
            if (!characterEntry.hasOwnProperty("form") || characterEntry.form === characterEntry.originalCommonForm) continue;
            for (var j = 0; j < characterEntry.places.length; j++) {
                var characterPlace = characterEntry.places[j];
                if (characterPlace.originalForm !== null) getPlaceSpan(characterPlace).characterAttributes.alternateGlyphs = characterPlace.originalForm;
            }
        }
    }

    // =========================================
    // 異体字の一覧 / Variant list
    // =========================================

    /**
     * 一覧のレイヤーを探す
     * @param {Document} doc - 対象のドキュメント
     * @returns {Layer|null} レイヤー（無ければ null）
     */
    function findGlyphListLayer(doc) {
        /* 無い名前は例外になる / getByName throws for a missing name */
        try {
            return doc.layers.getByName(GLYPH_LIST_LAYER_NAME);
        } catch (e) {
            return null;
        }
    }

    /**
     * 元の文字のフォントとサイズを読む（見本の字に引き継ぐ）
     * @param {Object} characterEntry - 文字のまとまり
     * @returns {{textFont: (TextFont|null), size: number}} フォント（読めなければ null）とサイズ
     */
    function readSourceTextStyle(characterEntry) {
        var sourceAttributes = getPlaceSpan(characterEntry.places[0]).characterAttributes;
        var textFont = null;
        /* 環境に無いフォントなどでは読めないことがある / unreadable for some fonts, such as missing ones */
        try { textFont = sourceAttributes.textFont; } catch (e) { }
        return { textFont: textFont, size: sourceAttributes.size };
    }

    /**
     * テキストの字形を変える（フォントに無い字形は標準のまま）
     * @param {TextFrame} textFrame - 1文字のテキスト
     * @param {*} glyphForm - AlternateGlyphsForm の値
     * @returns {void}
     */
    function setFrameGlyphForm(textFrame, glyphForm) {
        /* 字形を受け付けないフォントでは例外になることがある / some fonts reject the form */
        try { textFrame.textRange.characterAttributes.alternateGlyphs = glyphForm; } catch (e) { }
    }

    /**
     * 一覧を作るレイヤーを返す（無ければ最前面に作る。ロック・非表示は解除する）
     * @param {Document} doc - 対象のドキュメント
     * @returns {Layer} レイヤー
     */
    function getGlyphListLayer(doc) {
        var listLayer = findGlyphListLayer(doc);
        if (!listLayer) {
            listLayer = doc.layers.add();
            listLayer.name = GLYPH_LIST_LAYER_NAME;
        }
        listLayer.locked = false;
        listLayer.visible = true;
        return listLayer;
    }

    /**
     * 一覧のレイヤーを削除する。選択中の文字がそのレイヤーにあるときは消さない
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @returns {void}
     */
    function deleteGlyphListLayer(doc, characterEntries) {
        var listLayer = findGlyphListLayer(doc);
        if (!listLayer) return;
        for (var i = 0; i < characterEntries.length; i++) {
            var sourceFrame = resolveTextRangeFrame(getPlaceSpan(characterEntries[i].places[0]));
            if (sourceFrame && sourceFrame.layer === listLayer) return;
        }
        listLayer.locked = false;
        listLayer.visible = true;
        listLayer.remove();
    }

    /**
     * ダイアログボックスの［IVS を表示］用に、選択した文字（先頭の GLYPH_FORM_CHECK_MAX 文字）の IVS を一時レイヤーで調べる。
     * 異体字の組の標準字形と、すでに見つけた IVS と形が異なるものだけを返す
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @returns {Array<{entry: Object, variant: string, selector: string}>} まとまり・IVS の文字列・セレクタ
     */
    function discoverIvsVariants(doc, characterEntries) {
        var ivsChoices = [];
        var savedActiveLayer = doc.activeLayer;
        var checkLayer = doc.layers.add();
        /* 途中で失敗しても一時レイヤーは必ず消す / Always remove the temporary layer */
        try {
            var checkCount = Math.min(characterEntries.length, GLYPH_FORM_CHECK_MAX);
            for (var i = 0; i < checkCount; i++) {
                var characterEntry = characterEntries[i];
                var sourceStyle = readSourceTextStyle(characterEntry);
                var groupMembers = listGroupMembers(characterEntry);
                var keptShapes = [];

                /**
                 * 1文字を一時レイヤーに描いて形を測り、すぐ消す
                 * @param {string} character - 文字（IVS 可）
                 * @returns {{shape: string, glyphCount: number}|null} measureGlyphShape() の結果
                 */
                var measureCharacter = function (character) {
                    var checkFrame = addPointLabel(checkLayer, character, [0, 0], sourceStyle.textFont, sourceStyle.size);
                    var measuredShape = measureGlyphShape(checkFrame);
                    checkFrame.remove();
                    return measuredShape;
                };

                for (var j = 0; j < groupMembers.length; j++) {
                    var memberShape = measureCharacter(groupMembers[j]);
                    if (memberShape) keptShapes.push(memberShape.shape);
                }
                placeIvsCells(groupMembers, function (ivsCell) {
                    var measuredShape = measureCharacter(ivsCell.character);
                    if (measuredShape === null || measuredShape.glyphCount !== 1) return "invalid";
                    if (indexOfItem(keptShapes, measuredShape.shape) >= 0) return "duplicate";
                    keptShapes.push(measuredShape.shape);
                    ivsChoices.push({ entry: characterEntry, variant: ivsCell.character, selector: ivsCell.selector });
                    return "kept";
                });
            }
        } finally {
            checkLayer.remove();
            doc.activeLayer = savedActiveLayer;
        }
        return ivsChoices;
    }

    /**
     * フォントが持っている字形を調べる。先頭の GLYPH_FORM_CHECK_MAX 文字について、各字形の形を一時レイヤーで描いて
     * 標準字形と比べ、どれか1文字でも形が変わる字形を「対応している」とする（測れなかったときも対応扱い）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @returns {boolean[]} GLYPH_FORMS と同じ並びの、対応していれば true の配列（標準字形は常に true）
     */
    function detectSupportedGlyphForms(doc, characterEntries) {
        var supportedForms = [];
        for (var k = 0; k < GLYPH_FORMS.length; k++) supportedForms.push(k === 0);
        var savedActiveLayer = doc.activeLayer;
        var checkLayer = null;
        /* 途中で失敗しても一時レイヤーは必ず消す / Always remove the temporary layer */
        try {
            checkLayer = doc.layers.add();
            var checkCount = Math.min(characterEntries.length, GLYPH_FORM_CHECK_MAX);
            for (var i = 0; i < checkCount; i++) {
                var sourceStyle = readSourceTextStyle(characterEntries[i]);
                var defaultShape = null;
                for (var j = 0; j < GLYPH_FORMS.length; j++) {
                    if (j > 0 && supportedForms[j]) continue; /* もう分かった字形は測らない / skip forms already known */
                    var checkFrame = addPointLabel(checkLayer, characterEntries[i].current, [0, 0], sourceStyle.textFont, sourceStyle.size);
                    setFrameGlyphForm(checkFrame, GLYPH_FORMS[j].form);
                    var measuredShape = measureGlyphShape(checkFrame);
                    checkFrame.remove();
                    if (j === 0) {
                        defaultShape = measuredShape ? measuredShape.shape : null;
                        continue;
                    }
                    if (measuredShape === null || defaultShape === null || measuredShape.shape !== defaultShape) supportedForms[j] = true;
                }
            }
        } catch (e) {
            /* 調べられなければ、すべて対応扱い（ディムにしない） / when it cannot be checked, dim nothing */
            for (k = 0; k < supportedForms.length; k++) supportedForms[k] = true;
        } finally {
            if (checkLayer) checkLayer.remove();
            doc.activeLayer = savedActiveLayer;
        }
        return supportedForms;
    }

    /**
     * ポイント文字を置く（既定は中央揃え）
     * @param {Layer} listLayer - 置き先のレイヤー
     * @param {string} text - 文字列
     * @param {number[]} anchor - ベースラインの位置 [x, y]（ドキュメント座標）
     * @param {TextFont|null} textFont - フォント（null なら既定のまま）
     * @param {number} fontSize - 文字サイズ（pt）
     * @param {Object} [labelOptions] - isVertical: 縦組み（anchor は字の中心になる）／alignLeft: 左揃え
     * @returns {TextFrame} 置いたテキスト
     */
    function addPointLabel(listLayer, text, anchor, textFont, fontSize, labelOptions) {
        var isVertical = !!(labelOptions && labelOptions.isVertical);
        var alignLeft = !!(labelOptions && labelOptions.alignLeft);
        /* 向きは作るときに決める（あとから変えるとテキストが作り直され、参照が無効になる）
           Set the orientation on creation; changing it later recreates the frame and invalidates the reference */
        var labelFrame = isVertical ? listLayer.textFrames.pointText(anchor, TextOrientation.VERTICAL) : listLayer.textFrames.pointText(anchor);
        labelFrame.contents = text;
        if (!alignLeft) labelFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
        var labelAttributes = labelFrame.textRange.characterAttributes;
        labelAttributes.size = fontSize;
        if (textFont) {
            /* 環境に無いフォントなどで失敗したら既定のまま / keep the default if the font cannot be set */
            try { labelAttributes.textFont = textFont; } catch (e) { }
        }
        return labelFrame;
    }

    /**
     * ドキュメントのカラーモードに合わせた黒を返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {CMYKColor|RGBColor} 黒
     */
    function makeBlackColor(doc) {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykBlack = new CMYKColor();
            cmykBlack.black = 100;
            return cmykBlack;
        }
        return new RGBColor(); /* 0, 0, 0 */
    }

    /**
     * 一覧の背面に、余白を付けた黒い長方形（不透明度 GLYPH_LIST_BACK_OPACITY ％）を描く
     * @param {Document} doc - 対象のドキュメント
     * @param {Layer} listLayer - 一覧のレイヤー
     * @param {PageItem[]} listFrames - 一覧のテキスト（先頭が最背面）
     * @param {number} padding - 余白（pt）
     * @returns {PathItem} 描いた長方形
     */
    function addListBackdrop(doc, listLayer, listFrames, padding) {
        var listBounds = FitViewToItems.getItemsBounds(listFrames);
        var backdrop = listLayer.pathItems.rectangle(
            listBounds[1] + padding,
            listBounds[0] - padding,
            listBounds[2] - listBounds[0] + padding * 2,
            listBounds[1] - listBounds[3] + padding * 2
        );
        backdrop.filled = true;
        backdrop.fillColor = makeBlackColor(doc);
        backdrop.stroked = false;
        backdrop.opacity = GLYPH_LIST_BACK_OPACITY;
        /* 今回の一覧のすぐ背面に置く（前に作った一覧は隠さない） / just behind this list, without covering earlier lists */
        backdrop.move(listFrames[0], ElementPlacement.PLACEAFTER);
        return backdrop;
    }

    /**
     * 符号位置・字形の名前に使うフォントを返す（GLYPH_LIST_LABEL_FONT_NAMES を前から探す）
     * @returns {TextFont|null} フォント（どれも無ければ null）
     */
    function findLabelFont() {
        for (var i = 0; i < GLYPH_LIST_LABEL_FONT_NAMES.length; i++) {
            /* 無いフォント名は例外になる / getByName throws for a missing name */
            try { return app.textFonts.getByName(GLYPH_LIST_LABEL_FONT_NAMES[i]); } catch (e) { }
        }
        return null;
    }

    /**
     * 一覧の1行で試す文字を返す（今の文字を先頭に、登録された異体字を続ける。異体字セレクタは除く）
     * @param {Object} characterEntry - 文字のまとまり
     * @returns {string[]} 試す文字
     */
    function listGroupMembers(characterEntry) {
        var baseCharacter = stripVariationSelectors(characterEntry.current); /* IVS に置き換えたあとは元の文字で試す / use the base after an IVS replacement */
        var groupMembers = [baseCharacter];
        for (var i = 0; i < characterEntry.variants.length; i++) {
            if (characterEntry.variants[i] !== baseCharacter) groupMembers.push(characterEntry.variants[i]);
        }
        return groupMembers;
    }

    /**
     * 一覧の1行に並べる字を返す。先に異体字の組の文字をそれぞれ標準字形で並べ（Unicode の異体字が先に見えるように）、
     * 続けて文字ごとにほかの字形、最後に縦組みでのすべての字形を並べる（音引きの縦組み用の字形など。一覧では別の行に置く）。
     * 異体字の標準字形は異体字そのものなので字形の名前を添えない
     * @param {string[]} groupMembers - listGroupMembers() の文字
     * @returns {Array<{character: string, form: *, isVertical: boolean, codeText: string, noteText: string}>} 並べる字・字形・縦組みか・符号位置・添え書き
     */
    function listGlyphCells(groupMembers) {
        var listForms = getListGlyphForms();
        var glyphCells = [];
        var i, j;

        /**
         * 1つ足す
         * @param {string} character - 文字
         * @param {Object} glyphForm - getListGlyphForms() の要素
         * @param {boolean} isVertical - 縦組みなら true
         * @param {boolean} showFormName - 字形の名前を添えるなら true
         * @returns {void}
         */
        function pushCell(character, glyphForm, isVertical, showFormName) {
            var noteText = showFormName ? getLabel(glyphForm.labelKey) : "";
            glyphCells.push({ character: character, form: glyphForm.form, isVertical: isVertical, codeText: formatCodePoint(character), noteText: noteText });
        }

        for (i = 0; i < groupMembers.length; i++) pushCell(groupMembers[i], listForms[0], false, i === 0);
        for (i = 0; i < groupMembers.length; i++) {
            for (j = 1; j < listForms.length; j++) pushCell(groupMembers[i], listForms[j], false, true); /* 0 は標準字形 / index 0 is the default form */
        }
        for (i = 0; i < groupMembers.length; i++) {
            for (j = 0; j < listForms.length; j++) pushCell(groupMembers[i], listForms[j], true, true);
        }
        return glyphCells;
    }

    /**
     * 文字に異体字セレクタ（U+E0100 から数えて selectorIndex 番目）を付けた IVS の見本を返す
     * @param {string} baseCharacter - 元の文字
     * @param {number} selectorIndex - 0 始まりの番号（0 で U+E0100）
     * @returns {{character: string, selector: string, form: null, isVertical: boolean, codeText: string, noteText: string}} 並べる字・セレクタ・符号位置・添え書き（セレクタの符号位置）
     */
    function makeIvsCell(baseCharacter, selectorIndex) {
        var selector = String.fromCharCode(0xDB40, 0xDD00 + selectorIndex); /* U+E0100 + selectorIndex */
        return { character: baseCharacter + selector, selector: selector, form: null, isVertical: false, codeText: formatCodePoint(baseCharacter), noteText: formatCodePoint(selector) };
    }

    /**
     * 字のアウトラインの形を、比べられる文字列にする（位置によらないよう、外接範囲の左上からの座標で測る）。
     * 複製をアウトライン化して測り、測り終えたら消す
     * @param {TextFrame} glyphFrame - 1文字のテキスト
     * @returns {{shape: string, glyphCount: number}|null} 形の文字列と字形の数（測れなければ null）
     */
    function measureGlyphShape(glyphFrame) {
        var outlineGroup = null;
        var shapeParts = [];

        /**
         * パスの頂点を外接範囲の左上からの座標で控える
         * @param {PathItem} pathItem - パス
         * @param {number[]} originPoint - 外接範囲の左上 [x, y]
         * @returns {void}
         */
        function addPathShape(pathItem, originPoint) {
            var pathPoints = pathItem.pathPoints;
            for (var k = 0; k < pathPoints.length; k++) {
                var anchor = pathPoints[k].anchor;
                shapeParts.push(Math.round((anchor[0] - originPoint[0]) * 100) + "," + Math.round((originPoint[1] - anchor[1]) * 100));
            }
            shapeParts.push("|");
        }

        /**
         * グループ・複合パスをたどってパスを控える
         * @param {PageItem} item - 対象
         * @param {number[]} originPoint - 外接範囲の左上 [x, y]
         * @returns {void}
         */
        function walkShape(item, originPoint) {
            var k;
            if (item.typename === "PathItem") {
                addPathShape(item, originPoint);
            } else if (item.typename === "CompoundPathItem") {
                for (k = 0; k < item.pathItems.length; k++) addPathShape(item.pathItems[k], originPoint);
            } else if (item.typename === "GroupItem") {
                for (k = 0; k < item.pageItems.length; k++) walkShape(item.pageItems[k], originPoint);
            }
        }

        /* createOutline() は複製を消費し、字形の無い文字では中身の無いグループになるので、失敗は「測れない」とする
           createOutline() consumes the duplicate and may yield an empty group, so any failure means "unmeasurable" */
        try {
            outlineGroup = glyphFrame.duplicate().createOutline();
            var outlineBounds = outlineGroup.geometricBounds;
            walkShape(outlineGroup, [outlineBounds[0], outlineBounds[1]]);
            /* アウトラインは字形ごとに1つずつ並ぶ / the outline holds one item per glyph */
            return { shape: shapeParts.join(";"), glyphCount: outlineGroup.pageItems.length };
        } catch (e) {
            return null;
        } finally {
            if (outlineGroup) {
                try { outlineGroup.remove(); } catch (e2) { }
            }
        }
    }

    /**
     * 見本の字の大きさから、符号位置・字形の名前の大きさと字の間隔を求める
     * @param {number} glyphSize - 見本の字の大きさ（pt）
     * @returns {{glyph: number, code: number, formName: number, columnPitch: number}} 大きさと字の間隔（pt）
     */
    function computeListSizes(glyphSize) {
        var codeSize = Math.max(glyphSize * GLYPH_LIST_CODE_RATIO, GLYPH_LIST_LABEL_MIN_SIZE);
        var formNameSize = Math.max(codeSize * 0.75, GLYPH_LIST_LABEL_MIN_SIZE);
        /* 字の間隔は、字と最も長い字形の名前の広いほう / wide enough for the glyph and the longest form name */
        return { glyph: glyphSize, code: codeSize, formName: formNameSize, columnPitch: Math.max(glyphSize * 1.6, formNameSize * 11) };
    }

    /**
     * 行の上端から、字・符号位置・字形の名前のベースラインと、次の行の上端を求める
     * @param {number} lineTop - 行の上端
     * @param {Object} listSizes - computeListSizes() の結果
     * @returns {{glyph: number, code: number, formName: number, nextTop: number}} ベースラインと次の行の上端
     */
    function computeLineBaselines(lineTop, listSizes) {
        var glyphLine = lineTop - listSizes.glyph * 0.88;
        var codeLine = glyphLine - listSizes.glyph * 0.3 - listSizes.code;
        var formNameLine = codeLine - listSizes.formName * 1.4;
        return { glyph: glyphLine, code: codeLine, formName: formNameLine, nextTop: formNameLine - Math.max(listSizes.glyph * 0.5, listSizes.formName * 2) };
    }

    /**
     * 一覧全体の置き場所を決める。行の見出しは元のテキストの右、見本はその右（1文字目の大きさで決める）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @returns {Object} listLayer / labelFont / createdItems / rowTop / headingX / originX / backdropPadding
     */
    function createListContext(doc, characterEntries) {
        var sourceFrame = resolveTextRangeFrame(getPlaceSpan(characterEntries[0].places[0]));
        var sourceBounds = sourceFrame ? sourceFrame.geometricBounds : [0, 0, 0, 0];
        var firstSizes = computeListSizes(readSourceTextStyle(characterEntries[0]).size);
        var headingX = sourceBounds[2] + firstSizes.glyph; /* 元のテキストの右 / right of the source text */
        return {
            listLayer: getGlyphListLayer(doc),
            labelFont: findLabelFont(),
            createdItems: [],
            rowTop: sourceBounds[1],
            headingX: headingX,
            originX: headingX + firstSizes.code * ROW_LABEL_WIDTH_EM, /* 行の見出しの右 / right of the row heading */
            backdropPadding: firstSizes.glyph * GLYPH_LIST_BACK_PADDING
        };
    }

    /**
     * IVS の見本を、異体字の組の文字ごとに新しい形が続けて出なくなるまで試す
     * （IVS は符号位置ごとに登録されるので、辺には無く邉に多いなど、組の全員を試す）
     * @param {string[]} groupMembers - listGroupMembers() の文字
     * @param {Function} placeGlyphCell - 見本を置く関数（"kept" / "duplicate" / "invalid" を返す）
     * @returns {void}
     */
    function placeIvsCells(groupMembers, placeGlyphCell) {
        for (var memberIndex = 0; memberIndex < groupMembers.length; memberIndex++) {
            var missCount = 0;
            for (var selectorIndex = 0; selectorIndex < GLYPH_LIST_IVS_MAX && missCount < GLYPH_LIST_IVS_MISS_LIMIT; selectorIndex++) {
                var placeResult = placeGlyphCell(makeIvsCell(groupMembers[memberIndex], selectorIndex), true);
                if (placeResult === "invalid") break;
                missCount = (placeResult === "kept") ? 0 : missCount + 1;
            }
        }
    }

    /**
     * 1文字分の「横組み用」「縦組み用」の行を置く。形が同じ字は最初の1つだけを残し、GLYPH_LIST_COLUMNS で折り返す
     * @param {Object} listContext - createListContext() の結果（rowTop を進め、createdItems に足す）
     * @param {Object} characterEntry - 文字のまとまり
     * @returns {void}
     */
    function layoutEntryRows(listContext, characterEntry) {
        var listLayer = listContext.listLayer;
        var sourceStyle = readSourceTextStyle(characterEntry); /* 見本の字は元の文字のフォントとサイズを引き継ぐ / samples inherit the source font and size */
        var listSizes = computeListSizes(sourceStyle.size);
        var labelFont = listContext.labelFont || sourceStyle.textFont;
        var lineBaselines;      /* 今の行のベースライン / baselines of the current line */
        var headingBaseline;    /* 見出しを置く1行目の字のベースライン / glyph baseline of the first line, for the heading */
        var keptShapes = [];    /* 横組み用・縦組み用の両方で同じ形は1つだけ / one per shape across both rows */
        var columnIndex = 0;    /* 今の行に置いた数 / samples on the current line */
        var placedInRow = 0;    /* 見出しの行（折り返しを含む）に置いた数 / samples in the heading's row, wraps included */

        /**
         * 見出しの付いた行（横組み用・縦組み用）を始める
         * @returns {void}
         */
        function startRow() {
            lineBaselines = computeLineBaselines(listContext.rowTop, listSizes);
            headingBaseline = lineBaselines.glyph;
            columnIndex = 0;
            placedInRow = 0;
        }

        /**
         * 行を閉じる。見本を1つでも置いたら1行目の左に見出しを付けて次の行へ進み、置かなければ何もしない
         * @param {string} headingKey - 見出しの LABELS のパス
         * @returns {void}
         */
        function finishRow(headingKey) {
            if (placedInRow === 0) return;
            var headingAnchor = [listContext.headingX, headingBaseline + listSizes.glyph * 0.38 - listSizes.code * 0.35];
            listContext.createdItems.push(addPointLabel(listLayer, getLabel(headingKey), headingAnchor, labelFont, listSizes.code, { alignLeft: true }));
            listContext.rowTop = lineBaselines.nextTop; /* 次の行の上端 / top of the next row */
        }

        /**
         * 見本を1つ置き、すでに並べた字と同じ形なら消す
         * @param {Object} glyphCell - listGlyphCells() / makeIvsCell() の要素
         * @param {boolean} isIvsCell - IVS の見本なら true（測れない・2字形に分かれたものは無効とする）
         * @returns {string} 残したら "kept"、同じ形で消したら "duplicate"、IVS として無効なら "invalid"
         */
        function placeGlyphCell(glyphCell, isIvsCell) {
            /* 行がいっぱいなら次の行の先頭に置く（残すと決まってから折り返す） / a full line puts it at the head of the next line (wrap only once kept) */
            var willWrap = (columnIndex >= GLYPH_LIST_COLUMNS);
            var slotBaselines = willWrap ? computeLineBaselines(lineBaselines.nextTop, listSizes) : lineBaselines;
            var centerX = listContext.originX + listSizes.columnPitch * (willWrap ? 0 : columnIndex) + listSizes.columnPitch / 2;
            /* 縦組みは字の中心が anchor になるので、横組みの字面の中心（ベースラインの約0.38em上）に合わせる
               Vertical text centers on the anchor; match the center of a horizontal glyph (about 0.38 em above the baseline) */
            var glyphAnchor = glyphCell.isVertical ? [centerX, slotBaselines.glyph + listSizes.glyph * 0.38] : [centerX, slotBaselines.glyph];
            var glyphFrame = addPointLabel(listLayer, glyphCell.character, glyphAnchor, sourceStyle.textFont, listSizes.glyph, { isVertical: glyphCell.isVertical });
            if (glyphCell.form !== null) setFrameGlyphForm(glyphFrame, glyphCell.form);
            var measuredShape = measureGlyphShape(glyphFrame);
            /* IVS を解釈しないとセレクタが別の字形（豆腐など）になる / unsupported IVS renders the selector as a separate glyph */
            if (isIvsCell && (measuredShape === null || measuredShape.glyphCount !== 1)) {
                glyphFrame.remove();
                return "invalid";
            }
            /* すでに並べた字と同じ形なら出さない（測れなかった字は残す） / skip shapes already listed; keep unmeasurable ones */
            if (measuredShape !== null && indexOfItem(keptShapes, measuredShape.shape) >= 0) {
                glyphFrame.remove();
                return "duplicate";
            }
            if (measuredShape !== null) keptShapes.push(measuredShape.shape);
            if (willWrap) {
                lineBaselines = slotBaselines;
                columnIndex = 0;
            }
            columnIndex++;
            placedInRow++;
            listContext.createdItems.push(glyphFrame);
            listContext.createdItems.push(addPointLabel(listLayer, glyphCell.codeText, [centerX, slotBaselines.code], labelFont, listSizes.code));
            if (glyphCell.noteText) {
                listContext.createdItems.push(addPointLabel(listLayer, glyphCell.noteText, [centerX, slotBaselines.formName], labelFont, listSizes.formName));
            }
            return "kept";
        }

        var groupMembers = listGroupMembers(characterEntry);
        var glyphCells = listGlyphCells(groupMembers);
        var i;

        /* 横組み用の行：字形・異体字・IVS / Horizontal row: forms, variants and IVS */
        startRow();
        for (i = 0; i < glyphCells.length; i++) {
            if (!glyphCells[i].isVertical) placeGlyphCell(glyphCells[i], false);
        }
        placeIvsCells(groupMembers, placeGlyphCell);
        finishRow("listHeading.horizontal");

        /* 縦組み用の行：横組みと形が異なるものだけ（無ければ行ごと出さない） / Vertical row: only shapes that differ (omitted when none) */
        startRow();
        for (i = 0; i < glyphCells.length; i++) {
            if (glyphCells[i].isVertical) placeGlyphCell(glyphCells[i], false);
        }
        finishRow("listHeading.vertical");
    }

    /**
     * 選択した文字ごとに「横組み用」「縦組み用」の2行で、字形と異体字を、大きな字・符号位置・字形の名前で並べた一覧を作り、
     * 背面に黒い長方形を敷く。置き場所は元のテキストフレームの右
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @returns {PageItem[]} 作ったテキストと背面の長方形
     */
    function createGlyphList(doc, characterEntries) {
        var savedActiveLayer = doc.activeLayer;
        var savedCoordinateSystem = app.coordinateSystem;
        var listContext;
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        /* 途中で失敗しても座標系と作業レイヤーは必ず戻す / Always restore the coordinate system and the active layer */
        try {
            listContext = createListContext(doc, characterEntries);
            for (var i = 0; i < characterEntries.length; i++) layoutEntryRows(listContext, characterEntries[i]);
            if (listContext.createdItems.length > 0) {
                listContext.createdItems.push(addListBackdrop(doc, listContext.listLayer, listContext.createdItems, listContext.backdropPadding));
            }
        } finally {
            app.coordinateSystem = savedCoordinateSystem;
            doc.activeLayer = savedActiveLayer;
        }
        return listContext.createdItems;
    }

    /**
     * 字形・異体字の一覧を作り、作ったものを選択して、ウインドウに GLYPH_LIST_FIT_PERCENT ％ほどで収まるよう表示する
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @returns {void}
     */
    function createAndShowGlyphList(doc, characterEntries) {
        var createdItems = createGlyphList(doc, characterEntries);
        doc.selection = null;
        doc.selection = createdItems;
        FitViewToItems.fit(createdItems, { doc: doc, fillRatio: GLYPH_LIST_FIT_PERCENT / 100 });
        app.redraw();
    }

    // 画面にフィット（再利用パーツ、_templates/FitViewToItems.jsx） / Fit view to items (reusable)

    var FitViewToItems = (function () {

        // =========================================
        // ユーザー設定 / User Settings
        // =========================================

        /* ウィンドウに対して対象が占める割合の既定値（%）と範囲。呼び出し側で上書きできる
           Default share of the window the items fill, in percent, and its range; callers can override the default */
        var DEFAULT_FIT_PERCENT = 65;
        var FIT_PERCENT_RANGE = [10, 100];

        /* Illustratorが受け付ける表示倍率の範囲（3.125%〜6400%） / Zoom range Illustrator accepts */
        var VIEW_ZOOM_RANGE = [0.03125, 64];

        // =========================================
        // ローカライズ / Localization
        // =========================================
        var LABELS = {
            checkbox: {
                fitView: { ja: "画面にフィット", en: "Fit to Window" }
            },
            tooltip: {
                fitView: {
                    ja: "作成するオブジェクトが収まるよう表示倍率を合わせます。",
                    en: "Refits the view to the objects being created."
                },
                fitViewPercent: {
                    ja: "ウィンドウに対するオブジェクトの大きさ（100%でいっぱい）",
                    en: "Size of the objects relative to the window; 100% fills it"
                }
            }
        };

        /**
         * UI言語を返す
         * @returns {string} "ja" または "en"
         */
        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }

        /**
         * LABELS の組から指定言語の文言を返す
         * @param {object} labelSet - { ja: string, en: string } の組
         * @param {string} uiLang - "ja" または "en"
         * @returns {string} 文言（無ければ英語）
         */
        function getLabel(labelSet, uiLang) {
            return (labelSet[uiLang] != null) ? labelSet[uiLang] : labelSet.en;
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        /**
         * 数値を範囲に収める。数値として読めないときは既定値を返す
         * @param {string|number} value - 入力値
         * @param {number[]} range - [下限, 上限]
         * @param {number} fallbackValue - 読めないときの既定値
         * @returns {number} 範囲内の数値
         */
        function clampNumber(value, range, fallbackValue) {
            var numberValue = Number(value);
            if (isNaN(numberValue) || (typeof value === "string" && !/\S/.test(value))) numberValue = fallbackValue;
            return Math.min(range[1], Math.max(range[0], numberValue));
        }

        /**
         * 「□画面にフィット［65］%」の行を追加する（ラベルとツールチップは内蔵）
         * @param {Group|Panel|Window} parentContainer - 追加先のコンテナ
         * @param {object} [rowOptions] - value: チェックの初期値（既定 false）／percent: 割合の初期値／lang: 表示言語／
         *     stepper: true で入力欄の左に∧∨を付け、↑↓キーでも増減（StepperButtons 部品の addStepper / bindSteppedArrowKeys /
         *     redrawSteppersIn を使う。増減したら percentInput.onChange を呼ぶ）
         * @returns {{row: Group, checkbox: Checkbox, percentInput: EditText, getFillRatio: function, updateEnabled: function}} 作成したコントロール一式
         */
        function addControls(parentContainer, rowOptions) {
            if (!rowOptions) rowOptions = {};
            var uiLang = rowOptions.lang || getCurrentLang();

            var fitViewRow = parentContainer.add("group");
            fitViewRow.orientation = "row";
            fitViewRow.alignChildren = ["left", "center"];
            fitViewRow.spacing = 6;

            var fitViewCheck = fitViewRow.add("checkbox", undefined, getLabel(LABELS.checkbox.fitView, uiLang));
            fitViewCheck.helpTip = getLabel(LABELS.tooltip.fitView, uiLang);
            /* 明示的に true を渡したときだけONで始める / only an explicit true starts it checked */
            fitViewCheck.value = (rowOptions.value === true);

            var fitPercent = (rowOptions.percent > 0) ? rowOptions.percent : DEFAULT_FIT_PERCENT;
            var percentInput;
            var percentParent = fitViewRow;
            var stepperGroup = null;
            if (rowOptions.stepper === true) {
                /* ∧∨と入力欄は隙間0で突き合わせる / Butt the stepper against the field */
                percentParent = fitViewRow.add("group");
                percentParent.orientation = "row";
                percentParent.alignChildren = ["left", "center"];
                percentParent.spacing = 0;
                percentParent.margins = 0;
                stepperGroup = addStepper(percentParent, function () { return percentInput; }, {
                    integer: true,
                    min: FIT_PERCENT_RANGE[0],
                    max: FIT_PERCENT_RANGE[1],
                    /* text の代入ではイベントが来ないので、増減後に onChange を呼ぶ / assigning text fires no event */
                    onStep: function () { if (percentInput.onChange) percentInput.onChange(); }
                });
            }
            percentInput = percentParent.add("edittext", undefined, String(fitPercent));
            percentInput.characters = 3;
            percentInput.helpTip = getLabel(LABELS.tooltip.fitViewPercent, uiLang);
            if (stepperGroup) bindSteppedArrowKeys(percentInput, stepperGroup); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */
            var percentUnitLabel = fitViewRow.add("statictext", undefined, "%");

            var controls = {
                row: fitViewRow,
                checkbox: fitViewCheck,
                percentInput: percentInput,

                /**
                 * 入力欄の割合を 0〜1 の比率で返す
                 * @returns {number} ウィンドウに対して占める割合（1でいっぱい）
                 */
                getFillRatio: function () {
                    return clampNumber(percentInput.text, FIT_PERCENT_RANGE, fitPercent) / 100;
                },

                /**
                 * 割合の入力欄をチェックの状態に合わせて有効・無効にする
                 * @returns {void}
                 */
                updateEnabled: function () {
                    percentInput.enabled = fitViewCheck.value;
                    percentUnitLabel.enabled = fitViewCheck.value;
                    if (!stepperGroup) return;
                    stepperGroup.enabled = fitViewCheck.value;
                    /* ∧∨は自作描画なので、表示中なら描き直してディム表示を切り替える / redraw the custom-drawn stepper once shown */
                    if (fitViewRow.window.visible) redrawSteppersIn(stepperGroup);
                }
            };
            controls.updateEnabled();
            return controls;
        }

        /**
         * 複数アイテムを囲む外接範囲を求める（効果を含まない geometricBounds）
         * @param {PageItem[]} targetItems - 対象アイテム
         * @returns {number[]|null} [left, top, right, bottom]（求められない場合は null）
         */
        function getItemsBounds(targetItems) {
            var unionBounds = null;
            for (var i = 0; i < targetItems.length; i++) {
                var itemBounds = targetItems[i].geometricBounds;
                if (unionBounds === null) {
                    unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
                    continue;
                }
                if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
                if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
                if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
                if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
            }
            return unionBounds;
        }

        /**
         * 対象が指定の割合でウィンドウに収まるよう、中心を合わせて表示倍率を変える
         * 拡大・縮小のどちらも行う（KeepInView と違い、常に同じ大きさに見せる）
         * @param {PageItem[]} targetItems - 対象アイテム
         * @param {object} [fitOptions] - doc: 対象ドキュメント（省略時は最前面）／fillRatio: 占める割合（1でいっぱい）
         * @returns {boolean} 表示を動かしたら true
         */
        function fit(targetItems, fitOptions) {
            if (!targetItems || targetItems.length === 0) return false;
            if (!fitOptions) fitOptions = {};

            var targetDoc = fitOptions.doc || app.activeDocument;
            var bounds = getItemsBounds(targetItems);
            if (bounds === null) return false;

            var itemWidth = bounds[2] - bounds[0];
            var itemHeight = bounds[1] - bounds[3];
            var activeView = targetDoc.activeView;
            activeView.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
            if (itemWidth <= 0 || itemHeight <= 0) return true;

            /* 中心をそろえたあとの表示範囲を基準に倍率を求める / Scale from the view bounds after the center has moved */
            var fillRatio = (fitOptions.fillRatio > 0) ? fitOptions.fillRatio : DEFAULT_FIT_PERCENT / 100;
            var viewBounds = activeView.bounds;
            var scale = Math.min(
                (viewBounds[2] - viewBounds[0]) / itemWidth,
                (viewBounds[1] - viewBounds[3]) / itemHeight
            ) * fillRatio;
            activeView.zoom = clampNumber(activeView.zoom * scale, VIEW_ZOOM_RANGE, 1);
            return true;
        }

        /**
         * 現在の表示位置と倍率を控える（キャンセル時に restoreView で戻す）
         * @param {Document} [targetDoc] - 対象ドキュメント（省略時は最前面）
         * @returns {{centerPoint: number[], zoom: number}} 控えた表示状態
         */
        function captureView(targetDoc) {
            var activeView = (targetDoc || app.activeDocument).activeView;
            return { centerPoint: activeView.centerPoint, zoom: activeView.zoom };
        }

        /**
         * captureView で控えた表示位置と倍率に戻す
         * @param {{centerPoint: number[], zoom: number}} viewState - 控えた表示状態
         * @param {Document} [targetDoc] - 対象ドキュメント（省略時は最前面）
         * @returns {void}
         */
        function restoreView(viewState, targetDoc) {
            if (!viewState) return;
            var activeView = (targetDoc || app.activeDocument).activeView;
            activeView.centerPoint = viewState.centerPoint;
            activeView.zoom = viewState.zoom;
        }

        return {
            addControls: addControls,
            fit: fit,
            captureView: captureView,
            restoreView: restoreView,
            getItemsBounds: getItemsBounds
        };

    })();

    // 画面にフィット（再利用パーツ）ここまで / End of the reusable fit view


    // =========================================
    // ダイアログ / Dialog
    // =========================================

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    /**
     * 一覧の文字を大きくする（細部の違いを見分けやすく。フォントを作れない環境ではそのまま）
     * @param {ListBox} targetList - 対象の一覧
     * @returns {void}
     */
    function enlargeListFont(targetList) {
        try {
            targetList.graphics.font = ScriptUI.newFont(targetList.graphics.font.name, "REGULAR", LIST_FONT_SIZE);
        } catch (e) { /* 既定の大きさのまま / keep the default size */ }
    }

    /**
     * 異体字のあるまとまりだけを返す
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @returns {Array<Object>} 異体字のあるまとまり
     */
    function filterVariantEntries(characterEntries) {
        var variantEntries = [];
        for (var i = 0; i < characterEntries.length; i++) {
            if (characterEntries[i].variants.length > 0) variantEntries.push(characterEntries[i]);
        }
        return variantEntries;
    }

    /**
     * 選択した文字と符号位置を「葛 U+845B　鴎 U+9D0E」の形で返す（SELECTION_CODES_MAX を超えたら「ほか N 字」）
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @returns {string} 表示する文字列
     */
    function formatSelectionCodes(characterEntries) {
        var codeParts = [];
        var shownCount = Math.min(characterEntries.length, SELECTION_CODES_MAX);
        for (var i = 0; i < shownCount; i++) {
            codeParts.push(characterEntries[i].current + " " + formatCodePoints(characterEntries[i].current));
        }
        var codesText = codeParts.join("   ");
        if (characterEntries.length > shownCount) {
            codesText += "   " + getLabel("status.moreCharacters", { count: characterEntries.length - shownCount });
        }
        return labelValueText("status.selectedCharacters", codesText);
    }

    /**
     * 区切り線を足す（区切り線の panel は setupPanel を通さない）
     * @param {Panel|Group} parentContainer - 追加先
     * @returns {Panel} 区切り線
     */
    function addDivider(parentContainer) {
        var divider = parentContainer.add("panel");
        divider.alignment = ["fill", "top"];
        divider.preferredSize.height = 1;
        divider.margins = 0;
        return divider;
    }

    /**
     * 左カラム：字形のラジオボタン（選択している文字に共通の字形を選んだ状態にする）
     * @param {Group} parentGroup - 追加先
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @param {boolean[]} supportedForms - detectSupportedGlyphForms() の結果（対応していない字形はディム）
     * @param {Object} dialogControls - 作ったコントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildGlyphPanel(parentGroup, characterEntries, supportedForms, dialogControls) {
        var glyphPanel = parentGroup.add("panel", undefined, getLabel("panel.glyph"));
        setupPanel(glyphPanel, 6);
        glyphPanel.alignChildren = ["left", "top"];
        var commonForm = getCommonGlyphForm(characterEntries);
        dialogControls.glyphRadios = [];
        for (var i = 0; i < GLYPH_FORMS.length; i++) {
            /* 字形（JIS78 など）と幅（等幅全角など）の間に区切り線 / divider between the forms and the widths */
            if (GLYPH_FORMS[i].form === AlternateGlyphsForm.FULLWIDTH) addDivider(glyphPanel);
            var glyphRadio = glyphPanel.add("radiobutton", undefined, getLabel(GLYPH_FORMS[i].labelKey));
            glyphRadio.helpTip = getLabel("tooltip.glyph");
            glyphRadio.glyphForm = GLYPH_FORMS[i].form;
            glyphRadio.value = (GLYPH_FORMS[i].form === commonForm);
            /* 今掛かっている字形は、対応していなくても選べるままにする / keep the current form selectable */
            glyphRadio.enabled = supportedForms[i] || glyphRadio.value;
            dialogControls.glyphRadios.push(glyphRadio);
        }
    }

    /**
     * 異体字の一覧の1行を作る（異体字のある文字が複数なら「元の文字 → 」を前に付ける）
     * @param {Object} characterEntry - 文字のまとまり
     * @param {string} variant - 異体字（IVS 可）
     * @param {boolean} hasPrefix - 元の文字を前に付けるなら true
     * @returns {string} 一覧の行
     */
    function formatVariantLine(characterEntry, variant, hasPrefix) {
        var linePrefix = hasPrefix ? characterEntry.original + " → " : "";
        return linePrefix + variant + "  " + formatCodePoints(variant);
    }

    /**
     * 右カラム：異体字（別の文字）の一覧と［IVS を表示］。異体字のある文字が1つなら異体字だけ、
     * 複数なら「元の文字 → 異体字」の形で並べる
     * @param {Group} parentGroup - 追加先
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @param {Object} dialogControls - 作ったコントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildVariantPanel(parentGroup, characterEntries, dialogControls) {
        var variantPanel = parentGroup.add("panel", undefined, getLabel("panel.variants"));
        setupPanel(variantPanel);

        var variantEntries = filterVariantEntries(characterEntries);
        var variantChoices = []; /* 一覧の行ごとの { entry, variant } / { entry, variant } per list line */
        var variantLines = [];
        for (var j = 0; j < variantEntries.length; j++) {
            for (var k = 0; k < variantEntries[j].variants.length; k++) {
                variantChoices.push({ entry: variantEntries[j], variant: variantEntries[j].variants[k] });
                variantLines.push(formatVariantLine(variantEntries[j], variantEntries[j].variants[k], variantEntries.length > 1));
            }
        }
        /* IVS を足せるよう、異体字が無くても一覧は作る / the list exists even without variants, so IVS can be added */
        if (variantEntries.length === 0) dialogControls.noVariantsText = variantPanel.add("statictext", undefined, getLabel("status.noVariants"));
        var variantList = variantPanel.add("listbox", undefined, variantLines);
        variantList.preferredSize = VARIANT_LIST_SIZE;
        variantList.alignment = ["fill", "fill"];
        variantList.helpTip = getLabel("tooltip.variants");
        enlargeListFont(variantList);
        /* 異体字のある文字が1つなら、今の文字を選んだ状態にする / with a single character, select its current form */
        if (variantEntries.length === 1) variantList.selection = indexOfItem(variantEntries[0].variants, variantEntries[0].current);

        var btnShowIvs = variantPanel.add("button", undefined, getLabel("button.showIvs"));
        btnShowIvs.alignment = "left";
        btnShowIvs.helpTip = getLabel("tooltip.showIvs");

        dialogControls.variantList = variantList;
        dialogControls.variantChoices = variantChoices;
        dialogControls.btnShowIvs = btnShowIvs;
    }

    /**
     * ［IVS を表示］：IVS を調べて異体字の一覧に足す（1回だけ。見つからなければ知らせる）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {void}
     */
    function appendIvsToVariantList(doc, characterEntries, dialogControls) {
        dialogControls.btnShowIvs.enabled = false;
        var ivsChoices = discoverIvsVariants(doc, characterEntries);
        if (ivsChoices.length === 0) {
            alert(getLabel("alert.noIvs"));
            return;
        }
        var hasPrefix = (characterEntries.length > 1);
        for (var i = 0; i < ivsChoices.length; i++) {
            dialogControls.variantList.add("item", formatVariantLine(ivsChoices[i].entry, ivsChoices[i].variant, hasPrefix));
            dialogControls.variantChoices.push({ entry: ivsChoices[i].entry, variant: ivsChoices[i].variant });
        }
        if (dialogControls.noVariantsText) dialogControls.noVariantsText.visible = false;
    }

    /**
     * 2カラムの下（ボタン行の上）：一覧の作り方のラジオボタンと［一覧を作成］
     * @param {Window} dialog - ダイアログ（閉じたあと main で読めるよう、作り方のラジオを .replaceLayerRadio に入れる）
     * @param {Object} dialogControls - 作ったコントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildListPanel(dialog, dialogControls) {
        var listPanel = dialog.add("panel", undefined, getLabel("panel.list"));
        setupPanel(listPanel, 6);
        listPanel.alignChildren = ["left", "top"];
        var addToLayerRadio = listPanel.add("radiobutton", undefined, getLabel("radio.listAddToLayer", { layer: GLYPH_LIST_LAYER_NAME }));
        addToLayerRadio.helpTip = getLabel("tooltip.listAddToLayer");
        var replaceLayerRadio = listPanel.add("radiobutton", undefined, getLabel("radio.listReplaceLayer", { layer: GLYPH_LIST_LAYER_NAME }));
        replaceLayerRadio.helpTip = getLabel("tooltip.listReplaceLayer");
        replaceLayerRadio.value = DELETE_LIST_LAYER_DEFAULT;
        addToLayerRadio.value = !DELETE_LIST_LAYER_DEFAULT;
        dialog.replaceLayerRadio = replaceLayerRadio;

        var btnCreateList = listPanel.add("button", undefined, getLabel("button.createList"));
        btnCreateList.helpTip = getLabel("tooltip.createList", { layer: GLYPH_LIST_LAYER_NAME });
        dialogControls.btnCreateList = btnCreateList;
    }

    /**
     * ラジオボタン・異体字の一覧・［IVS を表示］・［一覧を作成］の動きを付ける
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectedCharacters - 文字のまとまりと箇所
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {void}
     */
    function bindDialogEvents(doc, collectedCharacters, dialogControls) {
        /* IVS は押したときだけ調べる（開くのが遅くならないように） / look up IVS only on demand, to keep opening fast */
        dialogControls.btnShowIvs.onClick = function () {
            appendIvsToVariantList(doc, collectedCharacters.entries, dialogControls);
        };
        /* 閉じてから一覧を作る（作るのは main） / Close first; main creates the list */
        dialogControls.btnCreateList.onClick = function () {
            this.window.close(DIALOG_RESULT_CREATE_LIST);
        };
        /* 字形はクリックしたらすぐに変える / Change the glyph form as soon as a radio is clicked */
        for (var i = 0; i < dialogControls.glyphRadios.length; i++) {
            dialogControls.glyphRadios[i].onClick = function () {
                applyGlyphFormToAll(collectedCharacters.entries, this.glyphForm);
                app.redraw();
            };
        }
        /* 異体字はクリックしたらすぐに置き換える / Replace with the variant as soon as it is clicked */
        var variantList = dialogControls.variantList;
        variantList.onChange = function () {
            if (!variantList.selection) return;
            var variantChoice = dialogControls.variantChoices[variantList.selection.index];
            applyEntryVariant(variantChoice.entry, variantChoice.variant, collectedCharacters.places);
            dialogControls.selectionCodesText.text = formatSelectionCodes(collectedCharacters.entries);
            app.redraw();
        };
    }

    /**
     * 字形（左）と異体字（右）の2カラムのダイアログを作る
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collectedCharacters - 文字のまとまりと箇所
     * @param {boolean[]} supportedForms - detectSupportedGlyphForms() の結果
     * @returns {Window} ダイアログ
     */
    function buildDialog(doc, collectedCharacters, supportedForms) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var dialogControls = {};
        /* 選択した文字の符号位置（異体字に置き換えると書き直す） / code points of the selection, rewritten after a variant replacement */
        /* 下に余白を足すため group で包む / wrapped in a group to add space below */
        var selectionCodesGroup = dialog.add("group");
        selectionCodesGroup.alignment = ["fill", "top"];
        selectionCodesGroup.margins = [0, 0, 0, SELECTION_CODES_MARGIN];
        var selectionCodesText = selectionCodesGroup.add("statictext", undefined, formatSelectionCodes(collectedCharacters.entries));
        selectionCodesText.alignment = ["fill", "top"];
        selectionCodesText.justify = "center"; /* 幅いっぱいの中で左右中央 / centered across the full width */
        selectionCodesText.helpTip = getLabel("tooltip.selectionCodes");
        dialogControls.selectionCodesText = selectionCodesText;

        var columnGroup = dialog.add("group");
        columnGroup.orientation = "row";
        columnGroup.alignChildren = ["fill", "fill"];
        columnGroup.spacing = COLUMN_SPACING;

        buildGlyphPanel(columnGroup, collectedCharacters.entries, supportedForms, dialogControls);
        buildVariantPanel(columnGroup, collectedCharacters.entries, dialogControls);
        buildListPanel(dialog, dialogControls);
        bindDialogEvents(doc, collectedCharacters, dialogControls);

        var buttonRow = addButtonRow(dialog);

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        btnCancel.preferredSize.width = BUTTON_WIDTH;
        btnOK.preferredSize.width = BUTTON_WIDTH;
        alignRightOnlyButtonRow(buttonRow);
        return dialog;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択している文字の字形・異体字をダイアログで切り換える（キャンセルなら元に戻す）
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        var doc = app.activeDocument;
        var textTargets = collectTextTargets(doc);
        if (textTargets.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        var collectedCharacters = { entries: [], entryByCharacter: {}, places: [] };
        var variantMap = buildVariantMap(ITAIJI_GROUPS);
        for (var i = 0; i < textTargets.length; i++) collectCharacterPlaces(textTargets[i], variantMap, collectedCharacters);
        if (collectedCharacters.entries.length === 0) {
            alert(getLabel("alert.notFound"));
            return;
        }
        var savedSelection = saveSelection(doc);
        var supportedForms = detectSupportedGlyphForms(doc, collectedCharacters.entries);
        var dialog = buildDialog(doc, collectedCharacters, supportedForms);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        var dialogResult = dialog.show();
        var shouldDeleteListLayer = dialog.replaceLayerRadio.value;
        /* ［一覧を作成］なら切り換えは残し、一覧を作って選択・表示する / Create List keeps the changes, then makes, selects and shows the list */
        if (dialogResult === DIALOG_RESULT_CREATE_LIST) {
            if (shouldDeleteListLayer) deleteGlyphListLayer(doc, collectedCharacters.entries);
            createAndShowGlyphList(doc, collectedCharacters.entries);
            return;
        }
        /* キャンセルなら、すぐに反映した字形・異体字を元に戻す / On Cancel, undo the forms and variants applied along the way */
        if (dialogResult !== 1) revertAllEntries(collectedCharacters.entries, collectedCharacters.places);
        /* 実行前の選択に戻す。異体字で文字数が変わり範囲がずれていることがある
           Put back the original selection; variants may have shifted the range */
        try {
            restoreSelection(doc, savedSelection);
        } catch (e) {
            doc.selection = null;
        }
    }

    main();

})();
