#target illustrator
#targetengine "ReplaceAlternateGlyphsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

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
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-17";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ReplaceAlternateGlyphs.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ReplaceAlternateGlyphs.md"; /* README (English) */

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
        "籠篭", "蘆芦", "蝋蠟", "繋繫", "撹攪", "嘘噓", "叱𠮟", "剥剝", "頬頰", "填塡"
    ];

    /* ［一覧を作成］の書き出し先と大きさ / Output of Create List */
    var GLYPH_LIST_LAYER_NAME = "_glyph"; /* 一覧を作るレイヤー（無ければ最前面に作る） / layer for the list (created on top if missing) */
    var GLYPH_LIST_GLYPH_SIZE = 72;       /* 字の大きさ（pt） / glyph size in points */
    var GLYPH_LIST_CODE_SIZE  = 18;       /* 符号位置の大きさ（pt） / code point size in points */

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

    var BUTTON_WIDTH       = 90;          /* ボタンの幅 / button width */
    var VARIANT_LIST_SIZE  = [170, 240];  /* 異体字の一覧の大きさ [幅, 高さ] / variant list size */
    var LIST_FONT_SIZE     = 18;          /* 一覧の文字の大きさ（細部の違いを見分けやすく） / list font size, large enough to tell variants apart */

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
            variants: { ja: "異体字", en: "Variant Characters" }
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
            glyphQuarter: { ja: "等幅四分字形", en: "Quarter Widths" }
        },
        status: {
            noVariants: { ja: "登録された異体字はありません", en: "No registered variants" }
        },
        button: {
            createList: { ja: "一覧を作成", en: "Create List" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            glyph: {
                ja: "選択している文字の字形をすぐに変えます。フォントにその字形が無ければ見た目は変わりません（キャンセルで元に戻ります）",
                en: "Changes the glyph form of the selected characters right away. Nothing changes if the font lacks that form (Cancel restores them)"
            },
            createList: {
                ja: "異体字と符号位置を並べた一覧を、レイヤー「{layer}」に作ります（キャンセルしても残ります）",
                en: "Lays out the variants and their code points on the layer \"{layer}\" (kept even if you cancel)"
            },
            variants: {
                ja: "クリックした異体字（別の文字）に、選択している同じ文字をすべてすぐに置き換えます（キャンセルで元に戻ります）",
                en: "Replaces every selected occurrence of that character with the clicked variant right away (Cancel restores them)"
            }
        },
        alert: {
            noSelection: { ja: "テキストを選択してください。", en: "Please select text." },
            notFound: { ja: "選択に文字がありません。", en: "The selection contains no characters." }
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
     * @param {Object} collected - { entries: 文字ごとのまとまり, entryByCharacter: 文字 → まとまり, places: すべての箇所 }
     * @returns {void}
     */
    function collectCharacterPlaces(textTarget, variantMap, collected) {
        var targetText = textTarget.text;
        /* contents と characters の数が合わなければ1文字ずつ読み直す / re-read character by character if the counts disagree */
        if (targetText.length !== textTarget.length) {
            targetText = "";
            for (var r = 0; r < textTarget.length; r++) targetText += textTarget.owner.characters[textTarget.start + r].contents;
        }
        var codePoints = splitCodePoints(targetText);
        var offset = 0;
        for (var i = 0; i < codePoints.length; i++) {
            var codePoint = codePoints[i];
            var placeStart = textTarget.start + offset;
            offset += codePoint.length;
            if (!isTargetCharacter(codePoint)) continue;
            var characterEntry = collected.entryByCharacter.hasOwnProperty(codePoint) ? collected.entryByCharacter[codePoint] : null;
            if (!characterEntry) {
                characterEntry = {
                    original: codePoint,
                    current: codePoint,
                    variants: variantMap.hasOwnProperty(codePoint) ? variantMap[codePoint] : [],
                    places: []
                };
                collected.entryByCharacter[codePoint] = characterEntry;
                collected.entries.push(characterEntry);
            }
            var characterPlace = { owner: textTarget.owner, start: placeStart, length: codePoint.length };
            characterEntry.places.push(characterPlace);
            collected.places.push(characterPlace);
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
     * 一覧を作るレイヤーを返す（無ければ最前面に作る。ロック・非表示は解除する）
     * @param {Document} doc - 対象のドキュメント
     * @returns {Layer} レイヤー
     */
    function getGlyphListLayer(doc) {
        var listLayer;
        try {
            listLayer = doc.layers.getByName(GLYPH_LIST_LAYER_NAME);
        } catch (e) {
            listLayer = doc.layers.add();
            listLayer.name = GLYPH_LIST_LAYER_NAME;
        }
        listLayer.locked = false;
        listLayer.visible = true;
        return listLayer;
    }

    /**
     * 中央揃えのポイント文字を置く
     * @param {Layer} listLayer - 置き先のレイヤー
     * @param {string} text - 文字列
     * @param {number[]} anchor - ベースライン中央の位置 [x, y]（ドキュメント座標）
     * @param {TextFont|null} textFont - フォント（null なら既定のまま）
     * @param {number} fontSize - 文字サイズ（pt）
     * @returns {TextFrame} 置いたテキスト
     */
    function addCenteredLabel(listLayer, text, anchor, textFont, fontSize) {
        var labelFrame = listLayer.textFrames.pointText(anchor);
        labelFrame.contents = text;
        labelFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
        var labelAttributes = labelFrame.textRange.characterAttributes;
        labelAttributes.size = fontSize;
        if (textFont) {
            /* 環境に無いフォントなどで失敗したら既定のまま / keep the default if the font cannot be set */
            try { labelAttributes.textFont = textFont; } catch (e) { }
        }
        return labelFrame;
    }

    /**
     * 異体字のある文字ごとに1行、異体字の大きな字と符号位置を並べた一覧を作る。
     * 置き場所は元のテキストフレームの右。フォントは元の文字のもの
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} variantEntries - 異体字のある文字のまとまり
     * @returns {void}
     */
    function createGlyphList(doc, variantEntries) {
        var glyphSize = GLYPH_LIST_GLYPH_SIZE;
        var columnPitch = glyphSize * 1.6;                 /* 字の間隔 / distance between glyphs */
        var rowPitch = glyphSize * 1.4 + GLYPH_LIST_CODE_SIZE * 2; /* 行の間隔 / distance between rows */
        var savedActiveLayer = doc.activeLayer;
        var savedCoordinateSystem = app.coordinateSystem;
        app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        /* 途中で失敗しても座標系と作業レイヤーは必ず戻す / Always restore the coordinate system and the active layer */
        try {
            var listLayer = getGlyphListLayer(doc);
            var firstPlace = variantEntries[0].places[0];
            var sourceFrame = resolveTextRangeFrame(getPlaceSpan(firstPlace));
            var sourceBounds = sourceFrame ? sourceFrame.geometricBounds : [0, 0, 0, 0];
            var originX = sourceBounds[2] + glyphSize; /* 元のテキストの右 / right of the source text */
            var originTop = sourceBounds[1];
            for (var i = 0; i < variantEntries.length; i++) {
                var variantEntry = variantEntries[i];
                var textFont = null;
                try { textFont = getPlaceSpan(variantEntry.places[0]).characterAttributes.textFont; } catch (e) { }
                var glyphBaseline = originTop - rowPitch * i - glyphSize * 0.88;
                var codeBaseline = glyphBaseline - glyphSize * 0.3 - GLYPH_LIST_CODE_SIZE;
                for (var j = 0; j < variantEntry.variants.length; j++) {
                    var centerX = originX + columnPitch * j + columnPitch / 2;
                    addCenteredLabel(listLayer, variantEntry.variants[j], [centerX, glyphBaseline], textFont, glyphSize);
                    addCenteredLabel(listLayer, formatCodePoint(variantEntry.variants[j]), [centerX, codeBaseline], textFont, GLYPH_LIST_CODE_SIZE);
                }
            }
        } finally {
            app.coordinateSystem = savedCoordinateSystem;
            try { doc.activeLayer = savedActiveLayer; } catch (e2) { }
        }
    }

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
     * 左カラム：字形のラジオボタン（選択している文字に共通の字形を選んだ状態にする）
     * @param {Group} parentGroup - 追加先
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @param {Object} dialogControls - 作ったコントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildGlyphPanel(parentGroup, characterEntries, dialogControls) {
        var glyphPanel = parentGroup.add("panel", undefined, getLabel("panel.glyph"));
        setupPanel(glyphPanel, 6);
        glyphPanel.alignChildren = ["left", "top"];
        var commonForm = getCommonGlyphForm(characterEntries);
        dialogControls.glyphRadios = [];
        for (var i = 0; i < GLYPH_FORMS.length; i++) {
            var glyphRadio = glyphPanel.add("radiobutton", undefined, getLabel(GLYPH_FORMS[i].labelKey));
            glyphRadio.helpTip = getLabel("tooltip.glyph");
            glyphRadio.glyphForm = GLYPH_FORMS[i].form;
            glyphRadio.value = (GLYPH_FORMS[i].form === commonForm);
            dialogControls.glyphRadios.push(glyphRadio);
        }
    }

    /**
     * 右カラム：異体字（別の文字）の一覧。異体字のある文字が1つなら異体字だけ、
     * 複数なら「元の文字 → 異体字」の形で並べる
     * @param {Group} parentGroup - 追加先
     * @param {Array<Object>} characterEntries - 文字のまとまり
     * @param {Object} dialogControls - 作ったコントロールを入れるオブジェクト
     * @returns {void}
     */
    function buildVariantPanel(parentGroup, characterEntries, dialogControls) {
        var variantPanel = parentGroup.add("panel", undefined, getLabel("panel.variants"));
        setupPanel(variantPanel);

        var variantEntries = [];
        for (var i = 0; i < characterEntries.length; i++) {
            if (characterEntries[i].variants.length > 0) variantEntries.push(characterEntries[i]);
        }
        dialogControls.variantEntries = variantEntries;
        if (variantEntries.length === 0) {
            variantPanel.add("statictext", undefined, getLabel("status.noVariants"));
            dialogControls.variantList = null;
            return;
        }

        var variantChoices = []; /* 一覧の行ごとの { entry, variant } / { entry, variant } per list line */
        var variantLines = [];
        for (var j = 0; j < variantEntries.length; j++) {
            for (var k = 0; k < variantEntries[j].variants.length; k++) {
                var variant = variantEntries[j].variants[k];
                var linePrefix = (variantEntries.length > 1) ? variantEntries[j].original + " → " : "";
                variantChoices.push({ entry: variantEntries[j], variant: variant });
                variantLines.push(linePrefix + variant + "  " + formatCodePoint(variant));
            }
        }
        var variantList = variantPanel.add("listbox", undefined, variantLines);
        variantList.preferredSize = VARIANT_LIST_SIZE;
        variantList.alignment = ["fill", "fill"];
        variantList.helpTip = getLabel("tooltip.variants");
        enlargeListFont(variantList);
        /* 異体字のある文字が1つなら、今の文字を選んだ状態にする / with a single character, select its current form */
        if (variantEntries.length === 1) variantList.selection = indexOfItem(variantEntries[0].variants, variantEntries[0].current);
        dialogControls.variantList = variantList;
        dialogControls.variantChoices = variantChoices;
    }

    /**
     * ラジオボタン・異体字の一覧・［一覧を作成］の動きを付ける
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collected - 文字のまとまりと箇所
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {void}
     */
    function bindDialogEvents(doc, collected, dialogControls) {
        /* 異体字の一覧をレイヤーに作る / Lay out the variant list on its layer */
        dialogControls.btnCreateList.onClick = function () {
            createGlyphList(doc, dialogControls.variantEntries);
            app.redraw();
        };
        /* 字形はクリックしたらすぐに変える / Change the glyph form as soon as a radio is clicked */
        for (var i = 0; i < dialogControls.glyphRadios.length; i++) {
            dialogControls.glyphRadios[i].onClick = function () {
                applyGlyphFormToAll(collected.entries, this.glyphForm);
                app.redraw();
            };
        }
        /* 異体字はクリックしたらすぐに置き換える / Replace with the variant as soon as it is clicked */
        var variantList = dialogControls.variantList;
        if (!variantList) return;
        variantList.onChange = function () {
            if (!variantList.selection) return;
            var variantChoice = dialogControls.variantChoices[variantList.selection.index];
            applyEntryVariant(variantChoice.entry, variantChoice.variant, collected.places);
            app.redraw();
        };
    }

    /**
     * 字形（左）と異体字（右）の2カラムのダイアログを作る
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} collected - 文字のまとまりと箇所
     * @returns {Window} ダイアログ
     */
    function buildDialog(doc, collected) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        var columnGroup = dialog.add("group");
        columnGroup.orientation = "row";
        columnGroup.alignChildren = ["fill", "fill"];
        columnGroup.spacing = COLUMN_SPACING;

        var dialogControls = {};
        buildGlyphPanel(columnGroup, collected.entries, dialogControls);
        buildVariantPanel(columnGroup, collected.entries, dialogControls);
        var buttonRow = addButtonRow(dialog);
        var btnCreateList = buttonRow.leftGroup.add("button", undefined, getLabel("button.createList"));
        btnCreateList.helpTip = getLabel("tooltip.createList", { layer: GLYPH_LIST_LAYER_NAME });
        btnCreateList.enabled = (dialogControls.variantEntries.length > 0); /* 異体字が無ければ作れない / nothing to list without variants */
        dialogControls.btnCreateList = btnCreateList;
        bindDialogEvents(doc, collected, dialogControls);

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
        var collected = { entries: [], entryByCharacter: {}, places: [] };
        var variantMap = buildVariantMap(ITAIJI_GROUPS);
        for (var i = 0; i < textTargets.length; i++) collectCharacterPlaces(textTargets[i], variantMap, collected);
        if (collected.entries.length === 0) {
            alert(getLabel("alert.notFound"));
            return;
        }
        var savedSelection = saveSelection(doc);
        var dialog = buildDialog(doc, collected);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        var isConfirmed = (dialog.show() === 1);
        /* キャンセルなら、すぐに反映した字形・異体字を元に戻す / On Cancel, undo the forms and variants applied along the way */
        if (!isConfirmed) revertAllEntries(collected.entries, collected.places);
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
