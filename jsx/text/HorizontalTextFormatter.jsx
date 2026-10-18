#target illustrator
#targetengine "HorizontalTextFormatterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.HorizontalTextFormatter
Name=HorizontalTextFormatter
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.0
Release-Date=2026-10-16
Meta-URL=https://github.com/swwwitch/illustrator-scripts
Description-BEGIN
縦組み用に整えたテキストを横組み用に戻します。縦組みを横組みにし、欧文回転・縦中横を外して、
全角英数字・読点の桁区切り・中黒の小数点・組文字の単位などを横組みの書き方に書き換えます（VerticalTextFormatter の逆）。

Turns text formatted for vertical setting back into horizontal text: switches the orientation, removes Roman rotation and
Tate-Chu-Yoko, and rewrites full-width alphanumerics, 、 separators, ・ decimals, unit squares and more (the reverse of VerticalTextFormatter).
Description-END
SCRIPTMETA-END
*/

/*

### 概要

縦組み用に整えたテキストを横組み用に戻します。縦組みを横組みにし、欧文回転・縦中横を外して、
全角英数字・読点の桁区切り・中黒の小数点・組文字の単位などを横組みの書き方に書き換えます（VerticalTextFormatter の逆）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/HorizontalTextFormatter.md

### Overview

Turns text formatted for vertical setting back into horizontal text: switches the orientation, removes Roman rotation and
Tate-Chu-Yoko, and rewrites full-width alphanumerics, 、 separators, ・ decimals, unit squares and more (the reverse of VerticalTextFormatter).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/HorizontalTextFormatter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "HorizontalTextFormatter";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-16";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/HorizontalTextFormatter.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/HorizontalTextFormatter.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 基本設定 / Settings
    // =========================================

    /* ダイアログの初期値（前回の値は SettingsStore に残る） / Dialog defaults (the last values are kept by SettingsStore) */
    var DEFAULT_SETTINGS = {
        toHorizontal: true,
        rotateOff: true,
        tcyOff: true,
        glyphReset: true,
        fullwidthToHalf: true,
        toutenToComma: true,
        nakaguroToPeriod: true,
        symbolsToHalf: true,
        chonchonToQuotes: true,
        phoneHyphenToAscii: true,
        kumimojiToLetters: true,
        waveToHyphen: false,
        eraToWestern: false,
        kanjiToArabic: false
    };

    /* パネルごとの項目（ダイアログの並び順） / Options per panel, in dialog order */
    var ORIENTATION_KEYS = ["toHorizontal", "rotateOff", "tcyOff", "glyphReset"];
    var REPLACE_KEYS = ["fullwidthToHalf", "toutenToComma", "nakaguroToPeriod", "symbolsToHalf", "chonchonToQuotes",
        "phoneHyphenToAscii", "kumimojiToLetters", "waveToHyphen", "eraToWestern", "kanjiToArabic"];

    var PREVIEW_MAX_LINES = 300;   /* 確認の一覧に出す最大の行数 / max lines in the check list */

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

    var BUTTON_WIDTH        = 90;    /* OK／キャンセルの幅 / OK and Cancel width */
    var PREVIEW_LIST_HEIGHT = 142;   /* 確認の一覧の高さ / height of the check list */
    var PROGRESS_WIDTH      = 300;   /* プログレスバーの幅 / progress bar width */

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
            title: { ja: "横組みの整形", en: "Horizontal Text Formatter" }
        },
        panel: {
            orientation: { ja: "組み方向・回転", en: "Orientation and Rotation" },
            replace: { ja: "文字の変換", en: "Character Conversion" },
            preview: { ja: "変換の確認", en: "Conversion Check" }
        },
        checkbox: {
            toHorizontal: { ja: "縦組みを横組みに", en: "Make vertical text horizontal" },
            rotateOff: { ja: "欧文回転を外す", en: "Remove Roman rotation" },
            tcyOff: { ja: "縦中横を外す", en: "Remove Tate-Chu-Yoko" },
            glyphReset: { ja: "等幅字形を標準に戻す", en: "Reset fixed-width glyphs" },
            fullwidthToHalf: { ja: "全角英数字を半角に", en: "Full-width alnum to half-width" },
            toutenToComma: { ja: "桁区切り「、」→「,」", en: "Separator 、 → ," },
            nakaguroToPeriod: { ja: "小数点「・」→「.」", en: "Decimal ・ → ." },
            symbolsToHalf: { ja: "記号を半角に（％ ／ ＆ など）", en: "Symbols to half-width (％ ／ ＆ …)" },
            chonchonToQuotes: { ja: "引用符〝〟→“”", en: "Quotes 〝〟 → “”" },
            phoneHyphenToAscii: { ja: "番号のハイフンを「-」に", en: "Number hyphens to -" },
            kumimojiToLetters: { ja: "組文字の単位を英字に（㌢ → cm）", en: "Unit squares to letters (㌢ → cm)" },
            waveToHyphen: { ja: "範囲の「〜」→「-」", en: "Range 〜 → -" },
            eraToWestern: { ja: "和暦を西暦に", en: "Japanese era to Western year" },
            kanjiToArabic: { ja: "漢数字を算用数字に", en: "Kanji numerals to Arabic" }
        },
        preview: {
            none: { ja: "（変換する文字はありません）", en: "(Nothing to convert)" }
        },
        progress: {
            title: { ja: "横組みの整形", en: "Horizontal Text Formatter" },
            attributes: { ja: "欧文回転・縦中横を外しています...", en: "Removing rotation and Tate-Chu-Yoko..." },
            edit: { ja: "文字を変換中... {current}/{total}", en: "Converting text... {current}/{total}" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noSelection: { ja: "テキストを選択してください。", en: "Select text first." },
            actionFailed: { ja: "横組みの整形を適用できませんでした。", en: "Could not apply the horizontal text settings." }
        },
        tooltip: {
            toHorizontal: { ja: "縦組みのテキストを横組みにします（連結したテキストはすべてのフレーム）", en: "Turns vertical text horizontal (every frame of threaded text)" },
            rotateOff: { ja: "文字パネルの［縦組み中の欧文回転］をテキスト全体で外します", en: "Turns off Rotate Standard Roman Characters for the whole text" },
            tcyOff: { ja: "文字パネルの［縦中横］をテキスト全体で外します", en: "Turns off Tate-Chu-Yoko for the whole text" },
            glyphReset: {
                ja: "縦中横の等幅半角・三分・四分字形など、字形をテキスト全体で標準に戻します",
                en: "Resets glyph forms (such as the half-, third- and quarter-width forms used in Tate-Chu-Yoko) for the whole text"
            },
            fullwidthToHalf: { ja: "全角の英数字と「！」「？」を半角にします（ＰＤＦ → PDF）", en: "Makes full-width alphanumerics and ！？ half-width (ＰＤＦ → PDF)" },
            toutenToComma: { ja: "数字の間で、後ろがちょうど3桁の「、」を「,」にします（3、000 → 3,000）", en: "Turns 、 followed by exactly three digits into , (3、000 → 3,000)" },
            nakaguroToPeriod: { ja: "数字の間の「・」を「.」にします（12・5 → 12.5）", en: "Turns ・ between digits into . (12・5 → 12.5)" },
            symbolsToHalf: {
                ja: "全角の ％ ＆ ＃ ＠ ＊ ＋ ＝ ＜ ＞ ／ を半角にします。括弧と「：」「；」は全角のまま",
                en: "Makes ％ ＆ ＃ ＠ ＊ ＋ ＝ ＜ ＞ ／ half-width; brackets, ： and ； stay full-width"
            },
            chonchonToQuotes: { ja: "チョンチョン（〝 〟）を “ ” にします", en: "Turns 〝 〟 into “ ”" },
            phoneHyphenToAscii: {
                ja: "電話番号・郵便番号（数字-数字-数字、3桁-4桁）の「‐」「ー」「－」を「-」にします",
                en: "Turns ‐ ー － in phone and postal numbers into -"
            },
            kumimojiToLetters: { ja: "㍉ ㌢ ㌖ ㍍ ㍉㌘ ㌘ ㌕ ㌧ ㍉㍑ ㌥㍑ ㌔㍑ ㍑ ㌍ ㌔㌍ ㏄ ㎡ ㎤ などを mm cm km m mg g kg t ml dL kL L cal kcal cc m² cm³ などにします", en: "Turns unit squares back into letters (mm, cm, kg, ml, kcal, cc, m², cm³ and so on)" },
            waveToHyphen: { ja: "数字の間の「〜」「～」を「-」にします（10〜20 → 10-20）", en: "Turns 〜 between digits into - (10〜20 → 10-20)" },
            eraToWestern: { ja: "「令和8年」を「2026年」にします（元年・漢数字の年も）", en: "Turns 令和8年 into 2026年 (元年 and kanji years too)" },
            kanjiToArabic: {
                ja: "2字以上の漢数字の並びを算用数字にします（二〇二六 → 2026、二千二十六 → 2026、三、〇〇〇 → 3,000）。1字の「一」「十」などは変えません",
                en: "Turns runs of two or more kanji numerals into Arabic (二千二十六 → 2026, 三、〇〇〇 → 3,000); single characters are left"
            }
        }
    };

    // 設定の保存（再利用パーツ） / Settings store (reusable)

    var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
    var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

    /**
     * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
     * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
     * @param {string} lifetime - "session" または "persistent"
     * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
     * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
     */
    function createSettingsStore(storeName, lifetime, storeOptions) {
        var isPersistent = (lifetime === "persistent");
        var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
        var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
        var sessionKey = "__" + safeStoreName + "_Settings";
        var settingsFile = isPersistent
            ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
            : null;

        /**
         * 保存してある文字列を返す
         * @returns {string|null} 保存文字列。1度も保存していなければ null
         */
        function readStoredText() {
            if (!isPersistent) {
                return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
            }
            return settingsStoreReadTextFile(settingsFile);
        }

        /**
         * 文字列を保存する
         * @param {string} storedText - 保存する文字列
         * @returns {boolean} 保存できたら true
         */
        function writeStoredText(storedText) {
            if (!isPersistent) {
                $.global[sessionKey] = storedText;
                return true;
            }
            return settingsStoreWriteTextFile(settingsFile, storedText);
        }

        /**
         * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
         * @param {Object} defaultSettings - 既定値
         * @returns {Object} 設定（毎回新しいオブジェクト）
         */
        function load(defaultSettings) {
            var savedSettings = null;
            try {
                var storedText = readStoredText();
                if (storedText !== null) {
                    savedSettings = settingsStoreParse(storedText);
                } else if (legacyReader) {
                    savedSettings = legacyReader();
                }
            } catch (e) {
                $.writeln("SettingsStore.load(" + storeName + "): " + e);
                savedSettings = null;
            }
            return settingsStoreMerge(defaultSettings, savedSettings);
        }

        /**
         * 設定を保存する
         * @param {Object} settingValues - 保存する値
         * @returns {boolean} 保存できたら true
         */
        function save(settingValues) {
            try {
                return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
            } catch (e) {
                $.writeln("SettingsStore.save(" + storeName + "): " + e);
                return false;
            }
        }

        /**
         * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
         * @returns {boolean} 消せたら true
         */
        function clear() {
            if (legacyReader) return writeStoredText("{}");
            if (!isPersistent) {
                try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                return true;
            }
            try {
                return settingsFile.exists ? settingsFile.remove() : true;
            } catch (e) {
                $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                return false;
            }
        }

        return { load: load, save: save, clear: clear };
    }

    /**
     * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
     * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
     * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
     */
    function readSettingsLegacyFile(legacyFileOrPath) {
        try {
            var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
            var legacyText = settingsStoreReadTextFile(legacyFile);
            return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
        } catch (e) {
            $.writeln("readSettingsLegacyFile: " + e);
            return null;
        }
    }

    /**
     * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
     * @param {string} preferenceKey - 環境設定のキー
     * @returns {Object|null} 読み込んだ値。無い・読めないときは null
     */
    function readSettingsLegacyPreference(preferenceKey) {
        try {
            var legacyText = app.preferences.getStringPreference(preferenceKey);
            if (!legacyText) return null;
            return settingsStoreParseLegacyText(String(legacyText));
        } catch (e) {
            $.writeln("readSettingsLegacyPreference: " + e);
            return null;
        }
    }

    /**
     * テキストファイルを UTF-8 で読む
     * @param {File} textFile - 読むファイル
     * @returns {string|null} 中身。ファイルが無ければ null
     */
    function settingsStoreReadTextFile(textFile) {
        if (!textFile.exists) return null;
        textFile.encoding = "UTF-8";
        if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
        try {
            return textFile.read().replace(/^\uFEFF/, "");
        } finally {
            textFile.close();
        }
    }

    /**
     * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
     * @param {File} textFile - 書くファイル
     * @param {string} fileText - 中身
     * @returns {boolean} 書けたら true
     */
    function settingsStoreWriteTextFile(textFile, fileText) {
        try {
            var parentFolder = textFile.parent;
            if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
            try {
                textFile.write(fileText);
            } finally {
                textFile.close();
            }
            return true;
        } catch (e) {
            $.writeln("SettingsStore write: " + e);
            return false;
        }
    }

    /**
     * 値が配列か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 配列なら true
     */
    function settingsStoreIsArray(checkedValue) {
        return Object.prototype.toString.call(checkedValue) === "[object Array]";
    }

    /**
     * 値が素のオブジェクト（{ } で作ったもの）か
     * @param {*} checkedValue - 調べる値
     * @returns {boolean} 素のオブジェクトなら true
     */
    function settingsStoreIsPlainObject(checkedValue) {
        return checkedValue !== null && typeof checkedValue === "object"
            && Object.prototype.toString.call(checkedValue) === "[object Object]"
            && checkedValue.constructor === Object;
    }

    /**
     * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
     * @param {string} sourceText - 文字列
     * @returns {string} 引用符つきの文字列
     */
    function settingsStoreQuote(sourceText) {
        var quotedText = "\"";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            var oneChar = sourceText.charAt(i);
            if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
            else if (oneChar === "\n") quotedText += "\\n";
            else if (oneChar === "\r") quotedText += "\\r";
            else if (oneChar === "\t") quotedText += "\\t";
            else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
            else quotedText += oneChar;
        }
        return quotedText + "\"";
    }

    /**
     * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
     * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
     * @param {*} sourceValue - 値
     * @param {string} indentText - 今の字下げ
     * @param {number} depth - 入れ子の深さ
     * @returns {string|undefined} JSON の文字列。書けない値は undefined
     */
    function settingsStoreSerialize(sourceValue, indentText, depth) {
        if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
        if (sourceValue === null) return "null";
        var valueType = typeof sourceValue;
        if (valueType === "boolean") return sourceValue ? "true" : "false";
        if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
        if (valueType === "string") return settingsStoreQuote(sourceValue);
        var innerIndent = indentText + "  ";
        var itemTexts = [];
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var hasNested = false;
            for (i = 0; i < sourceValue.length; i++) {
                var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                itemTexts.push(itemText === undefined ? "null" : itemText);
                if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
            }
            if (!itemTexts.length) return "[]";
            if (!hasNested) return "[" + itemTexts.join(", ") + "]";
            return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
        }
        if (settingsStoreIsPlainObject(sourceValue)) {
            for (var key in sourceValue) {
                if (!sourceValue.hasOwnProperty(key)) continue;
                var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
            }
            if (!itemTexts.length) return "{}";
            return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
        }
        return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
    }

    /**
     * JSON（と toSource の出力）を読む。eval は使わない。
     * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
     * @param {string} sourceText - 読む文字列
     * @returns {*} 読み込んだ値
     */
    function settingsStoreParse(sourceText) {
        var readPos = 0;
        var textLength = sourceText.length;

        /**
         * 読み取り位置で失敗を知らせる
         * @param {string} reasonText - 理由
         * @returns {void}
         */
        function fail(reasonText) {
            throw new Error("settings parse error at " + readPos + ": " + reasonText);
        }

        /**
         * 空白を読み飛ばす
         * @returns {void}
         */
        function skipSpaces() {
            while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
        }

        /**
         * 識別子（英数字・_・$）を読む
         * @returns {string} 識別子。無ければ空文字
         */
        function readWord() {
            var startPos = readPos;
            while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
            return sourceText.substring(startPos, readPos);
        }

        /**
         * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
         * @returns {string} 文字列
         */
        function readString() {
            var quoteChar = sourceText.charAt(readPos++);
            var resultText = "";
            while (readPos < textLength) {
                var oneChar = sourceText.charAt(readPos++);
                if (oneChar === quoteChar) return resultText;
                if (oneChar !== "\\") { resultText += oneChar; continue; }
                var escapeChar = sourceText.charAt(readPos++);
                if (escapeChar === "n") resultText += "\n";
                else if (escapeChar === "r") resultText += "\r";
                else if (escapeChar === "t") resultText += "\t";
                else if (escapeChar === "b") resultText += "\b";
                else if (escapeChar === "f") resultText += "\f";
                else if (escapeChar === "v") resultText += "\v";
                else if (escapeChar === "0") resultText += "\0";
                else if (escapeChar === "u" || escapeChar === "x") {
                    var hexLength = (escapeChar === "u") ? 4 : 2;
                    var hexText = sourceText.substr(readPos, hexLength);
                    if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                    resultText += String.fromCharCode(parseInt(hexText, 16));
                    readPos += hexLength;
                } else resultText += escapeChar;
            }
            fail("unterminated string");
        }

        /**
         * 値を1つ読む
         * @param {number} depth - 入れ子の深さ
         * @returns {*} 値
         */
        function readValue(depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
            skipSpaces();
            var oneChar = sourceText.charAt(readPos);
            if (oneChar === "{") return readObject(depth);
            if (oneChar === "[") return readArray(depth);
            if (oneChar === "\"" || oneChar === "'") return readString();
            if (oneChar === "(") {
                readPos++;
                var innerValue = readValue(depth + 1);
                skipSpaces();
                if (sourceText.charAt(readPos) !== ")") fail("expected )");
                readPos++;
                return innerValue;
            }
            var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
            if (numberMatch) {
                readPos += numberMatch[0].length;
                return Number(numberMatch[0]);
            }
            var wordText = readWord();
            if (wordText === "true") return true;
            if (wordText === "false") return false;
            if (wordText === "null") return null;
            if (wordText === "NaN") return NaN;
            if (wordText === "Infinity") return Infinity;
            if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
            fail("unexpected " + (wordText || oneChar || "end of text"));
        }

        /**
         * 配列を読む
         * @param {number} depth - 入れ子の深さ
         * @returns {Array} 配列
         */
        function readArray(depth) {
            var resultArray = [];
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "]") {
                resultArray.push(readValue(depth + 1));
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
            }
            readPos++;
            return resultArray;
        }

        /**
         * オブジェクトを読む（__proto__ のキーは捨てる）
         * @param {number} depth - 入れ子の深さ
         * @returns {Object} オブジェクト
         */
        function readObject(depth) {
            var resultObject = {};
            readPos++;
            skipSpaces();
            while (sourceText.charAt(readPos) !== "}") {
                var keyChar = sourceText.charAt(readPos);
                var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                if (memberKey === "") fail("expected a key");
                skipSpaces();
                if (sourceText.charAt(readPos) !== ":") fail("expected :");
                readPos++;
                var memberValue = readValue(depth + 1);
                if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                skipSpaces();
                if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
            }
            readPos++;
            return resultObject;
        }

        var parsedValue = readValue(0);
        skipSpaces();
        if (readPos < textLength) fail("unexpected text after the value");
        return parsedValue;
    }

    /**
     * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
     * @param {string} legacyText - 旧形式の文字列
     * @returns {Object|null} 読み込んだ値
     */
    function settingsStoreParseLegacyText(legacyText) {
        var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return null;
        if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
        var keyValues = {};
        var textLines = trimmedText.split(/\r\n|\r|\n/);
        for (var i = 0; i < textLines.length; i++) {
            var separatorIndex = textLines[i].indexOf("=");
            if (separatorIndex < 1) continue;
            var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
            if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
        }
        return keyValues;
    }

    /**
     * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
     * @param {*} sourceValue - コピー元
     * @returns {*} コピー
     */
    function settingsStoreClone(sourceValue) {
        if (sourceValue === null || typeof sourceValue !== "object") {
            return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
        }
        var i;
        if (settingsStoreIsArray(sourceValue)) {
            var arrayCopy = [];
            for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
            return arrayCopy;
        }
        if (!settingsStoreIsPlainObject(sourceValue)) return null;
        var objectCopy = {};
        for (var key in sourceValue) {
            if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
        }
        return objectCopy;
    }

    /**
     * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
     * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
     * @param {*} defaultValue - 既定値
     * @param {*} savedValue - 保存値
     * @returns {*} 突き合わせた値（新しいオブジェクト）
     */
    function settingsStoreMerge(defaultValue, savedValue) {
        if (defaultValue === null || defaultValue === undefined) {
            return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
        }
        var defaultType = typeof defaultValue;
        var savedType = typeof savedValue;
        if (defaultType === "boolean") {
            if (savedType === "boolean") return savedValue;
            if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
            if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
            return defaultValue;
        }
        if (defaultType === "number") {
            if (savedType === "number" && isFinite(savedValue)) return savedValue;
            if (savedType === "string" && /\S/.test(savedValue)) {
                var parsedNumber = Number(savedValue);
                if (isFinite(parsedNumber)) return parsedNumber;
            }
            return defaultValue;
        }
        if (defaultType === "string") {
            if (savedType === "string") return savedValue;
            if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
            if (savedType === "boolean") return String(savedValue);
            return defaultValue;
        }
        if (settingsStoreIsArray(defaultValue)) {
            return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
        }
        if (defaultType === "object") {
            var savedIsObject = settingsStoreIsPlainObject(savedValue);
            var hasDefaultKeys = false;
            var mergedObject = {};
            for (var key in defaultValue) {
                if (!defaultValue.hasOwnProperty(key)) continue;
                hasDefaultKeys = true;
                mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
            }
            /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
            if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
            return mergedObject;
        }
        return defaultValue;
    }

    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

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

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

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
    // 変換の計画 / Conversion plan
    // =========================================

    /* 組文字の単位を英字に戻す対応（2字の組み合わせを先に） / Unit squares back to letters (two-character pairs first) */
    var KUMIMOJI_UNIT_PAIRS = [["㍉㌘", "mg"], ["㍉㍑", "ml"], ["㌥㍑", "dL"], ["㌔㍑", "kL"], ["㌔㌍", "kcal"]];
    var KUMIMOJI_UNITS = {
        "㍉": "mm", "㌢": "cm", "㌖": "km", "㍍": "m", "㌘": "g", "㌕": "kg", "㌧": "t", "㍑": "L", "㌍": "cal", "㏄": "cc",
        "㎟": "mm²", "㎠": "cm²", "㎡": "m²", "㎢": "km²", "㎣": "mm³", "㎤": "cm³", "㎥": "m³", "㎦": "km³"
    };

    /* 半角にする全角の記号（括弧・「：」「；」は横組みでも全角が普通なので除く）
       Full-width symbols made half-width (brackets, ： and ； usually stay full-width in horizontal text too) */
    var HALFWIDTH_SYMBOLS = "％＆＃＠＊＋＝＜＞／";

    /* 番号のハイフンとみなす文字 / Characters treated as hyphens in phone and postal numbers */
    var PHONE_HYPHEN_SOURCE = "[-‐ー－−]";

    var KANJI_NUMERAL_DIGITS = "〇一二三四五六七八九";
    var KANJI_SMALL_UNIT_VALUES = { "十": 10, "百": 100, "千": 1000 };
    var KANJI_LARGE_UNIT_VALUES = { "万": 10000, "億": 100000000, "兆": 1000000000000 };

    /* 元号の開始年（西暦）。元年は開始年 / First Western year of each era (元年 is that year) */
    var ERA_START_YEARS = { "令和": 2019, "平成": 1989, "昭和": 1926, "大正": 1912, "明治": 1868 };

    /**
     * 半角・全角の数字か
     * @param {string} oneChar - 1文字
     * @returns {boolean} 0〜9・０〜９ なら true
     */
    function isAnyDigit(oneChar) {
        return (oneChar >= "0" && oneChar <= "9") || (oneChar >= "０" && oneChar <= "９");
    }

    /**
     * 全角の英数字と「！」「？」なら半角を、そうでなければ null を返す
     * @param {string} oneChar - 1文字
     * @returns {string|null} 半角の文字
     */
    function toHalfwidthAlnumChar(oneChar) {
        var charCode = oneChar.charCodeAt(0);
        var isFullAlnum = (charCode >= 0xFF10 && charCode <= 0xFF19) || (charCode >= 0xFF21 && charCode <= 0xFF3A) ||
            (charCode >= 0xFF41 && charCode <= 0xFF5A) || oneChar === "！" || oneChar === "？";
        return isFullAlnum ? String.fromCharCode(charCode - 0xFEE0) : null;
    }

    /**
     * 漢数字の並びを算用数字の文字列にする（二〇二六 → "2026"、〇〇五 → "005"、二千二十六 → "2026"、二万五〇〇〇 → "25000"）
     * 単位語（十・百・千・万…）が無ければ1字ずつ置き換えて桁を残す
     * @param {string} kanjiText - 漢数字の並び
     * @returns {string|null} 算用数字（読めなければ null）
     */
    function parseKanjiNumber(kanjiText) {
        var hasUnit = /[十百千万億兆]/.test(kanjiText);
        if (!hasUnit) {
            var digitText = "";
            for (var i = 0; i < kanjiText.length; i++) digitText += KANJI_NUMERAL_DIGITS.indexOf(kanjiText.charAt(i));
            return digitText;
        }
        var total = 0;
        var section = 0;
        var pendingDigits = "";
        for (var j = 0; j < kanjiText.length; j++) {
            var kanjiChar = kanjiText.charAt(j);
            var digitIndex = KANJI_NUMERAL_DIGITS.indexOf(kanjiChar);
            if (digitIndex >= 0) {
                pendingDigits += digitIndex;
            } else if (KANJI_SMALL_UNIT_VALUES[kanjiChar]) {
                section += (pendingDigits === "" ? 1 : parseInt(pendingDigits, 10)) * KANJI_SMALL_UNIT_VALUES[kanjiChar];
                pendingDigits = "";
            } else if (KANJI_LARGE_UNIT_VALUES[kanjiChar]) {
                section += (pendingDigits === "" ? 0 : parseInt(pendingDigits, 10));
                if (section === 0) section = 1;
                total += section * KANJI_LARGE_UNIT_VALUES[kanjiChar];
                section = 0;
                pendingDigits = "";
            } else {
                return null;
            }
        }
        return String(total + section + (pendingDigits === "" ? 0 : parseInt(pendingDigits, 10)));
    }

    /**
     * テキストの変換内容を決める。書き換え（後ろから順）と、確認の一覧用の変化（前から順）を返す
     * @param {string} sourceText - 元の文字列
     * @param {Object} formatOptions - 変換の設定
     * @returns {{edits: Array<{start: number, length: number, text: string}>, changes: Array<{before: string, after: string}>, outputLength: number}}
     */
    function planHorizontalText(sourceText, formatOptions) {
        var edits = [];
        var changes = [];
        var outputLength = 0;
        var i = 0;

        /**
         * 元の [start, start+length) を replacementText にしたことを控える
         * @param {number} start - 元の位置
         * @param {number} length - 元の文字数
         * @param {string} replacementText - 置き換える文字列
         * @returns {void}
         */
        function addSegment(start, length, replacementText) {
            var originalText = sourceText.substr(start, length);
            if (replacementText !== originalText) {
                edits.push({ start: start, length: length, text: replacementText });
                changes.push({ before: originalText, after: replacementText });
            }
            outputLength += replacementText.length;
        }

        while (i < sourceText.length) {
            var restText = sourceText.substr(i, 40);
            var currentChar = sourceText.charAt(i);
            var previousChar = sourceText.charAt(i - 1);
            var nextChar = sourceText.charAt(i + 1);
            var matched = null;

            /* 和暦 → 西暦（令和8年 → 2026年） / Japanese era → Western year */
            if (formatOptions.eraToWestern && (matched = /^(令和|平成|昭和|大正|明治)(元|[0-9０-９]+|[〇一二三四五六七八九十]+)年/.exec(restText))) {
                var eraYearText = matched[2];
                var eraYear = 1;
                if (eraYearText !== "元") {
                    eraYear = parseInt(/[〇一二三四五六七八九十]/.test(eraYearText) ? parseKanjiNumber(eraYearText) : toHalfwidthDigits(eraYearText), 10);
                }
                addSegment(i, matched[0].length, (ERA_START_YEARS[matched[1]] + eraYear - 1) + "年");
                i += matched[0].length;
                continue;
            }
            /* 漢数字 → 算用数字。読点で3桁ずつ区切った並び（三、〇〇〇）は1つの数として扱う
               Kanji numerals → Arabic; digits grouped by 、 every three (三、〇〇〇) are one number */
            if (formatOptions.kanjiToArabic && (matched = /^[〇一二三四五六七八九]+(?:、[〇一二三四五六七八九]{3})+(?![〇一二三四五六七八九])/.exec(restText))) {
                var groupedDigits = matched[0].split("、");
                for (var g = 0; g < groupedDigits.length; g++) groupedDigits[g] = parseKanjiNumber(groupedDigits[g]);
                addSegment(i, matched[0].length, groupedDigits.join(formatOptions.toutenToComma ? "," : "、"));
                i += matched[0].length;
                continue;
            }
            /* 漢数字の小数（一二・五）も1つの数として扱う / Kanji decimals (一二・五) are one number */
            if (formatOptions.kanjiToArabic && (matched = /^[〇一二三四五六七八九]+・[〇一二三四五六七八九]+/.exec(restText))) {
                var decimalParts = matched[0].split("・");
                addSegment(i, matched[0].length, parseKanjiNumber(decimalParts[0]) + (formatOptions.nakaguroToPeriod ? "." : "・") + parseKanjiNumber(decimalParts[1]));
                i += matched[0].length;
                continue;
            }
            /* 2字以上の漢数字の並びだけ（「一緒」「十」のような1字は変えない） / Runs of two or more only (single characters such as 一 are left) */
            if (formatOptions.kanjiToArabic && (matched = /^[〇一二三四五六七八九十百千万億兆]{2,}/.exec(restText))) {
                var kanjiValue = parseKanjiNumber(matched[0]);
                if (kanjiValue !== null && kanjiValue !== "NaN") {
                    addSegment(i, matched[0].length, kanjiValue);
                    i += matched[0].length;
                    continue;
                }
            }
            /* 電話番号・郵便番号のハイフン → 「-」 / Hyphens in phone and postal numbers → "-" */
            if (formatOptions.phoneHyphenToAscii && isAnyDigit(currentChar) && !isAnyDigit(previousChar)) {
                matched = new RegExp("^[0-9０-９]+(?:" + PHONE_HYPHEN_SOURCE + "[0-9０-９]+){2,}").exec(restText) ||
                    new RegExp("^[0-9０-９]{3}" + PHONE_HYPHEN_SOURCE + "[0-9０-９]{4}(?![0-9０-９])").exec(restText);
                if (matched) {
                    var phoneText = matched[0].replace(new RegExp(PHONE_HYPHEN_SOURCE, "g"), "-");
                    if (formatOptions.fullwidthToHalf) phoneText = toHalfwidthDigits(phoneText);
                    addSegment(i, matched[0].length, phoneText);
                    i += matched[0].length;
                    continue;
                }
            }
            /* 組文字の単位 → 英字 / Unit squares → letters */
            if (formatOptions.kumimojiToLetters) {
                var pairText = sourceText.substr(i, 2);
                var pairReplacement = null;
                for (var p = 0; p < KUMIMOJI_UNIT_PAIRS.length; p++) {
                    if (KUMIMOJI_UNIT_PAIRS[p][0] === pairText) pairReplacement = KUMIMOJI_UNIT_PAIRS[p][1];
                }
                if (pairReplacement) {
                    addSegment(i, 2, pairReplacement);
                    i += 2;
                    continue;
                }
                if (KUMIMOJI_UNITS.hasOwnProperty(currentChar)) {
                    addSegment(i, 1, KUMIMOJI_UNITS[currentChar]);
                    i++;
                    continue;
                }
            }

            var replacementChar = currentChar;
            var halfChar = toHalfwidthAlnumChar(currentChar);
            if (formatOptions.fullwidthToHalf && halfChar !== null) {
                replacementChar = halfChar;
            } else if (formatOptions.toutenToComma && currentChar === "、" && isAnyDigit(previousChar) &&
                    /^[0-9０-９]{3}(?![0-9０-９])/.test(sourceText.substr(i + 1, 4))) {
                /* 数字の間で、後ろがちょうど3桁なら桁区切り / A separator when exactly three digits follow */
                replacementChar = ",";
            } else if (formatOptions.nakaguroToPeriod && currentChar === "・" && isAnyDigit(previousChar) && isAnyDigit(nextChar)) {
                replacementChar = ".";
            } else if (formatOptions.waveToHyphen && (currentChar === "〜" || currentChar === "～") && isAnyDigit(previousChar) && isAnyDigit(nextChar)) {
                replacementChar = "-";
            } else if (formatOptions.symbolsToHalf && HALFWIDTH_SYMBOLS.indexOf(currentChar) >= 0) {
                replacementChar = String.fromCharCode(currentChar.charCodeAt(0) - 0xFEE0);
            } else if (formatOptions.chonchonToQuotes && (currentChar === "〝" || currentChar === "〟")) {
                replacementChar = (currentChar === "〝") ? "“" : "”";
            }
            addSegment(i, 1, replacementChar);
            i++;
        }
        edits.reverse();
        return { edits: edits, changes: changes, outputLength: outputLength };
    }

    /**
     * 全角数字を半角にする
     * @param {string} digitText - 文字列
     * @returns {string} 全角数字を半角にした文字列
     */
    function toHalfwidthDigits(digitText) {
        return digitText.replace(/[０-９]/g, function (fullDigit) { return String.fromCharCode(fullDigit.charCodeAt(0) - 0xFEE0); });
    }

    // =========================================
    // 一時アクション定義 / Action definition
    // =========================================
    var ACTION_SET_NAME = "HorizontalTextFormatter";   /* アクションセット名 / action set name */
    var ROTATE_OFF_ACTION_NAME = "rotateRomanOff";     /* 欧文回転を外す / turn Roman rotation off */
    var TCY_OFF_ACTION_NAME = "tateChuYokoOff";        /* 縦中横を外す / turn Tate-Chu-Yoko off */

    /**
     * 文字パネルの真偽値の項目を1つだけ設定するアクション1件ぶんの行を返す
     * （文字パネルのイベントは、書いたパラメーターだけを適用する）
     * @param {number} actionIndex - アクションの番号（1始まり）
     * @param {string} actionName - アクション名
     * @param {number} paramKey - パラメーターのキー
     * @param {number} paramValue - 値（1 でオン、0 でオフ）
     * @returns {string[]} アクション定義の行
     */
    function buildCharacterToggleActionLines(actionIndex, actionName, paramKey, paramValue) {
        return ["/action-" + actionIndex + " {"].concat(
            buildActionNameLines("\t", actionName),
            [
                "\t/keyIndex 0",
                "\t/colorIndex 0",
                "\t/isOpen 1",
                "\t/eventCount 1",
                "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0",
                "\t\t/internalName (adobe_SLOCharacterPalette)"
            ],
            buildActionNameLines("\t\t", "文字", "localizedName"),
            [
                "\t\t/isOpen 1",
                "\t\t/isOn 1",
                "\t\t/hasDialog 0",
                "\t\t/parameterCount 1",
                "\t\t/parameter-1 {",
                "\t\t\t/key " + paramKey,
                "\t\t\t/showInPalette 4294967295",
                "\t\t\t/type (boolean)",
                "\t\t\t/value " + paramValue,
                "\t\t}",
                "\t}",
                "}"
            ]
        );
    }

    /**
     * 欧文回転（'rriv'）と縦中横（'tyen'）を外す2件を持つアクション定義を返す
     * @returns {string} アクション定義のテキスト
     */
    function buildHorizontalTextActions() {
        return ["/version 3"].concat(
            buildActionNameLines("", ACTION_SET_NAME),
            ["/isOpen 1", "/actionCount 2"],
            buildCharacterToggleActionLines(1, ROTATE_OFF_ACTION_NAME, 1920100726, 0),
            buildCharacterToggleActionLines(2, TCY_OFF_ACTION_NAME, 1954112878, 0),
            [""]
        ).join("\n");
    }

    // =========================================
    // 対象のテキスト / Text targets
    // =========================================

    /**
     * テキストフレームが縦組みか
     * @param {TextFrame} textFrame - 判定するフレーム
     * @returns {boolean} 縦組みなら true
     */
    function isVerticalFrame(textFrame) {
        return !!textFrame && textFrame.orientation === TextOrientation.VERTICAL;
    }

    /**
     * ストーリーのテキストフレームを配列で返す
     * @param {Story} textStory - ストーリー
     * @returns {TextFrame[]} 連結しているテキストフレーム（連結していなければ1つ）
     */
    function getStoryFrames(textStory) {
        var storyFrames = [];
        for (var i = 0; i < textStory.textFrames.length; i++) storyFrames.push(textStory.textFrames[i]);
        return storyFrames;
    }

    /**
     * 選択から処理の対象を集める。文字選択中はその範囲、オブジェクト選択ならテキストフレーム（連結はストーリーごと）
     * @param {Document} doc - 対象のドキュメント
     * @returns {Array<Object>} { frame（連結していないフレーム）, story, ownerFrames, offset（ストーリーの文字の番号のずれ）, text, isTextSelection }
     */
    function collectTextTargets(doc) {
        var selection = doc.selection;
        var textTargets = [];
        /* 文字選択中は selection が TextRange（[0] が無い） / While text is selected, selection is a TextRange without [0] */
        if (selection && selection.typename === "TextRange") {
            if (selection.length > 0) {
                textTargets.push({
                    story: selection.story,
                    offset: selection.characterOffset - 1,
                    text: selection.contents,
                    ownerFrames: getStoryFrames(selection.story),
                    isTextSelection: true
                });
            }
            return textTargets;
        }
        var textFrames = collectSelectionTextFrames(selection, { skipLocked: true, skipHidden: true });
        var storyHeadFrames = [];
        for (var i = 0; i < textFrames.length; i++) {
            var storyFrames = getStoryFrames(textFrames[i].story);
            if (storyFrames.length <= 1) {
                textTargets.push({ frame: textFrames[i], story: textFrames[i].story, offset: 0, text: textFrames[i].contents, ownerFrames: [textFrames[i]], isTextSelection: false });
                continue;
            }
            /* 連結テキストはストーリー全体を1つの対象にする / Threaded text is one target for the whole story */
            if (indexOfValue(storyHeadFrames, storyFrames[0]) >= 0) continue;
            storyHeadFrames.push(storyFrames[0]);
            textTargets.push({ story: textFrames[i].story, offset: 0, text: textFrames[i].story.textRange.contents, ownerFrames: storyFrames, isTextSelection: false });
        }
        return textTargets;
    }

    /**
     * 配列の中で値が一致する位置を返す（ES3 に Array.indexOf は無い）
     * @param {Array} values - 配列
     * @param {*} targetValue - 探す値
     * @returns {number} 位置（無ければ -1）
     */
    function indexOfValue(values, targetValue) {
        for (var i = 0; i < values.length; i++) {
            if (values[i] === targetValue) return i;
        }
        return -1;
    }

    /**
     * 縦組みのテキストフレームを横組みにし、選択から対象を集め直す。
     * 向きを変えるとテキストフレームが作り直され、変える前の参照は使えなくなる（there is no document）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - collectTextTargets() の戻り値
     * @returns {Array<Object>} 集め直した対象（縦組みが無ければ渡したまま）
     */
    function convertTargetsToHorizontal(doc, textTargets) {
        var hasVertical = false;
        for (var i = 0; i < textTargets.length && !hasVertical; i++) {
            for (var j = 0; j < textTargets[i].ownerFrames.length; j++) {
                if (isVerticalFrame(textTargets[i].ownerFrames[j])) hasVertical = true;
            }
        }
        if (!hasVertical) return textTargets;

        var textSelectionTarget = textTargets[0].isTextSelection ? textTargets[0] : null;
        doc.selection = null;
        for (var k = 0; k < textTargets.length; k++) {
            for (var m = 0; m < textTargets[k].ownerFrames.length; m++) textTargets[k].ownerFrames[m].selected = true;
        }
        /* 向きを変えるたびに選択を読み直し、縦組みが残っていれば1つずつ変える / Re-read the selection after each change, one frame at a time */
        var frameCount = collectSelectionTextFrames(doc.selection, { skipLocked: true, skipHidden: true }).length;
        for (var guard = 0; guard < frameCount; guard++) {
            var currentFrames = collectSelectionTextFrames(doc.selection, { skipLocked: true, skipHidden: true });
            var verticalFrame = null;
            for (var n = 0; n < currentFrames.length && !verticalFrame; n++) {
                if (isVerticalFrame(currentFrames[n])) verticalFrame = currentFrames[n];
            }
            if (!verticalFrame) break;
            verticalFrame.orientation = TextOrientation.HORIZONTAL;
        }

        if (!textSelectionTarget) return collectTextTargets(doc);
        /* 文字選択は、作り直したフレームのストーリーで同じ位置を指す / A text selection points at the same offset in the new story */
        var newStory = collectSelectionTextFrames(doc.selection, { skipLocked: true, skipHidden: true })[0].story;
        return [{
            story: newStory,
            offset: textSelectionTarget.offset,
            text: textSelectionTarget.text,
            ownerFrames: getStoryFrames(newStory),
            isTextSelection: true
        }];
    }

    /**
     * 対象の文字のコレクションを返す（文字選択・連結はストーリー全体の番号で扱う）
     * @param {Object} textTarget - 対象
     * @returns {Characters} 文字のコレクション
     */
    function getTargetCharacters(textTarget) {
        return textTarget.frame ? textTarget.frame.characters : textTarget.story.characters;
    }

    /**
     * 対象の中の範囲を返す
     * @param {Object} textTarget - 対象
     * @param {number} start - 範囲の先頭（対象の中の位置）
     * @param {number} length - 文字数
     * @returns {TextRange} 範囲
     */
    function getTargetSpan(textTarget, start, length) {
        var spanRange = getTargetCharacters(textTarget)[textTarget.offset + start];
        spanRange.length = length;
        return spanRange;
    }

    /**
     * 書き換えを適用する。文字数が同じなら1文字ずつ、異なるなら余分を消して先頭の文字に入れる（書式は先頭の文字のもの）
     * @param {Object} textTarget - 対象
     * @param {Array<Object>} edits - planHorizontalText() の edits（後ろから順）
     * @returns {void}
     */
    function applyTextEdits(textTarget, edits) {
        for (var i = 0; i < edits.length; i++) {
            var characters = getTargetCharacters(textTarget);
            var baseIndex = textTarget.offset + edits[i].start;
            var editText = edits[i].text;
            if (editText.length === edits[i].length) {
                for (var k = 0; k < editText.length; k++) {
                    if (characters[baseIndex + k].contents !== editText.charAt(k)) characters[baseIndex + k].contents = editText.charAt(k);
                }
            } else {
                for (var r = edits[i].length - 1; r >= 1; r--) characters[baseIndex + r].remove();
                characters[baseIndex].contents = editText;
            }
        }
    }

    /**
     * 対象をまとめて選択する（文字選択はその範囲、オブジェクト選択はテキストフレーム）
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - 対象
     * @param {number} rangeLength - 文字選択のときの範囲の文字数
     * @returns {void}
     */
    function selectTextTargets(doc, textTargets, rangeLength) {
        doc.selection = null;
        if (textTargets[0].isTextSelection) {
            getTargetSpan(textTargets[0], 0, rangeLength).select();
            return;
        }
        for (var i = 0; i < textTargets.length; i++) {
            for (var j = 0; j < textTargets[i].ownerFrames.length; j++) textTargets[i].ownerFrames[j].selected = true;
        }
    }

    /**
     * 対象の全体の等幅字形を標準の字形に戻す（縦中横の等幅半角・三分・四分字形など）
     * @param {Object} textTarget - 対象
     * @param {number} outputLength - 書き換え後の文字数
     * @returns {void}
     */
    function resetGlyphForms(textTarget, outputLength) {
        if (outputLength === 0) return;
        getTargetSpan(textTarget, 0, outputLength).characterAttributes.alternateGlyphs = AlternateGlyphsForm.DEFAULTFORM;
    }

    /**
     * 横組みへの変換・欧文回転と縦中横を外す・文字の書き換え・字形を戻す、を順に行う
     * @param {Document} doc - 対象のドキュメント
     * @param {Array<Object>} textTargets - 対象
     * @param {Object} formatOptions - 変換の設定
     * @returns {string} 失敗したときの理由（成功は空文字）
     */
    function runHorizontalTextConversion(doc, textTargets, formatOptions) {
        /* 横組みにするとフレームが作り直されるので、最初に変えて対象を集め直す / Converting recreates frames, so do it first and re-collect */
        if (formatOptions.toHorizontal) {
            try {
                textTargets = convertTargetsToHorizontal(doc, textTargets);
            } catch (convertError) {
                return convertError + " (line " + convertError.line + ")";
            }
        }
        var plans = [];
        for (var i = 0; i < textTargets.length; i++) plans.push(planHorizontalText(textTargets[i].text, formatOptions));
        var needsAction = formatOptions.rotateOff || formatOptions.tcyOff;
        var stepTotal = (needsAction ? 1 : 0) + textTargets.length;
        var stepDone = 0;
        var progress = createProgressWindow(getLabel("progress.title"));

        /**
         * 進み具合を1つ進めて表示する
         * @param {string} message - 表示する文言
         * @returns {void}
         */
        function advanceProgress(message) {
            stepDone++;
            progress.set(stepDone / stepTotal, message);
        }

        if (needsAction && !loadTemporaryActionSet(buildHorizontalTextActions(), ACTION_SET_NAME)) {
            progress.close();
            return "loadTemporaryActionSet";
        }
        try {
            /* 欧文回転・縦中横を、書き換え前のテキスト全体から外す / Turn rotation and Tate-Chu-Yoko off over the whole text first */
            if (needsAction) {
                progress.set(0, getLabel("progress.attributes"));
                selectTextTargets(doc, textTargets, textTargets[0].text.length);
                if (formatOptions.rotateOff) app.doScript(ROTATE_OFF_ACTION_NAME, ACTION_SET_NAME, false);
                if (formatOptions.tcyOff) app.doScript(TCY_OFF_ACTION_NAME, ACTION_SET_NAME, false);
                advanceProgress(getLabel("progress.attributes"));
            }
            for (var j = 0; j < textTargets.length; j++) {
                applyTextEdits(textTargets[j], plans[j].edits);
                if (formatOptions.glyphReset) resetGlyphForms(textTargets[j], plans[j].outputLength);
                advanceProgress(getLabel("progress.edit", { current: j + 1, total: textTargets.length }));
            }
            return "";
        } catch (e) {
            $.writeln("runHorizontalTextConversion: " + e);
            return e + " (line " + e.line + ")";
        } finally {
            if (needsAction) unloadTemporaryActionSet(ACTION_SET_NAME);
            progress.close();
            /* 元の選択に戻す / Restore the original selection */
            try {
                selectTextTargets(doc, textTargets, plans.length ? plans[0].outputLength : 0);
            } catch (restoreError) { /* 戻せなくても続ける / keep going */ }
        }
    }

    /**
     * 確認の一覧に出す行を作る（同じ変換はまとめて件数を付ける）
     * @param {Array<Object>} textTargets - 対象
     * @param {Object} formatOptions - 変換の設定
     * @returns {string[]} 一覧の行
     */
    function buildPreviewLines(textTargets, formatOptions) {
        var lineKeys = [];
        var lineCounts = {};
        for (var i = 0; i < textTargets.length; i++) {
            var changes = planHorizontalText(textTargets[i].text, formatOptions).changes;
            for (var c = 0; c < changes.length; c++) {
                var lineText = changes[c].before + " → " + changes[c].after;
                if (lineCounts.hasOwnProperty(lineText)) {
                    lineCounts[lineText]++;
                } else {
                    lineCounts[lineText] = 1;
                    lineKeys.push(lineText);
                }
            }
        }
        var previewLines = [];
        for (var k = 0; k < lineKeys.length && k < PREVIEW_MAX_LINES; k++) {
            previewLines.push(lineKeys[k] + (lineCounts[lineKeys[k]] > 1 ? "  ×" + lineCounts[lineKeys[k]] : ""));
        }
        return previewLines;
    }

    // =========================================
    // プログレスバー / Progress bar
    // =========================================

    /**
     * 進捗表示用のパレットウィンドウを生成して表示する（非モーダルなので処理中に更新できる）
     * 戻り値の set(fraction 0..1, message) で進捗を更新、close() で閉じる
     * @param {string} title - ウィンドウのタイトル
     * @returns {{set: function, close: function}} 進捗の更新と終了の操作
     */
    function createProgressWindow(title) {
        var progressWindow = new Window("palette", title);
        setupWindow(progressWindow, 8);
        var progressBar = progressWindow.add("progressbar", undefined, 0, 100);
        progressBar.preferredSize = [PROGRESS_WIDTH, 8];
        var progressLabel = progressWindow.add("statictext", undefined, "");
        progressLabel.preferredSize.width = PROGRESS_WIDTH;
        progressWindow.show();
        progressWindow.update();
        return {
            set: function (fraction, message) {
                var percent = Math.round(fraction * 100);
                progressBar.value = (percent < 0) ? 0 : (percent > 100 ? 100 : percent);
                if (message != null) { progressLabel.text = message; }
                progressWindow.update(); /* 同期ループ中でも再描画させる / force a repaint even inside a synchronous loop */
            },
            close: function () { progressWindow.close(); }
        };
    }

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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * チェックボックスを追加する（説明があればツールチップに）
     * @param {Group|Panel} parent - 追加先
     * @param {string} settingKey - 設定のキー（LABELS.checkbox / LABELS.tooltip のキーも同じ）
     * @param {boolean} value - 初期値
     * @returns {Checkbox} チェックボックス
     */
    function addOptionCheckbox(parent, settingKey, value) {
        var optionCheckbox = parent.add("checkbox", undefined, getLabel("checkbox." + settingKey));
        optionCheckbox.value = value;
        if (LABELS.tooltip[settingKey]) optionCheckbox.helpTip = getLabel(LABELS.tooltip[settingKey]);
        return optionCheckbox;
    }

    /**
     * ダイアログを組み立てる
     * @param {Object} settings - 初期値（DEFAULT_SETTINGS の形）
     * @returns {Object} dialog と各コントロール
     */
    function buildDialog(settings) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);
        var dialogControls = { dialog: dialog };

        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        /* 組み方向・回転 / Orientation and rotation */
        var orientationPanel = columnsGroup.add("panel", undefined, getLabel("panel.orientation"));
        setupPanel(orientationPanel, 6);
        for (var i = 0; i < ORIENTATION_KEYS.length; i++) {
            dialogControls[ORIENTATION_KEYS[i]] = addOptionCheckbox(orientationPanel, ORIENTATION_KEYS[i], settings[ORIENTATION_KEYS[i]]);
        }

        /* 文字の変換 / Character conversion */
        var replacePanel = columnsGroup.add("panel", undefined, getLabel("panel.replace"));
        setupPanel(replacePanel, 6);
        for (var j = 0; j < REPLACE_KEYS.length; j++) {
            dialogControls[REPLACE_KEYS[j]] = addOptionCheckbox(replacePanel, REPLACE_KEYS[j], settings[REPLACE_KEYS[j]]);
        }

        /* 変換の確認 / Conversion check */
        var previewPanel = dialog.add("panel", undefined, getLabel("panel.preview"));
        setupPanel(previewPanel);
        dialogControls.previewList = previewPanel.add("listbox", undefined, []);
        dialogControls.previewList.preferredSize.height = PREVIEW_LIST_HEIGHT;

        var buttonRow = addButtonRow(dialog);
        dialogControls.btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        dialogControls.btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);
        dialogControls.btnCancel.preferredSize.width = BUTTON_WIDTH;
        dialogControls.btnOK.preferredSize.width = BUTTON_WIDTH;
        return dialogControls;
    }

    /**
     * ダイアログの値を設定のオブジェクトにする
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @returns {Object} DEFAULT_SETTINGS と同じ形
     */
    function readDialogOptions(dialogControls) {
        var formatOptions = {};
        var settingKeys = ORIENTATION_KEYS.concat(REPLACE_KEYS);
        for (var i = 0; i < settingKeys.length; i++) formatOptions[settingKeys[i]] = dialogControls[settingKeys[i]].value;
        return formatOptions;
    }

    /**
     * ダイアログのイベントをつなぐ（項目を変えたら確認の一覧を作り直す）
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {Array<Object>} textTargets - collectTextTargets() の戻り値
     * @returns {void}
     */
    function bindDialogEvents(dialogControls, textTargets) {
        /**
         * 確認の一覧を今の設定で作り直す
         * @returns {void}
         */
        function updatePreview() {
            var previewLines = buildPreviewLines(textTargets, readDialogOptions(dialogControls));
            dialogControls.previewList.removeAll();
            if (previewLines.length === 0) previewLines.push(getLabel("preview.none"));
            for (var i = 0; i < previewLines.length; i++) dialogControls.previewList.add("item", previewLines[i]);
        }

        var settingKeys = ORIENTATION_KEYS.concat(REPLACE_KEYS);
        for (var j = 0; j < settingKeys.length; j++) dialogControls[settingKeys[j]].onClick = updatePreview;
        updatePreview();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめ、ダイアログで設定を選んで実行する
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

        var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent");
        var dialogControls = buildDialog(settingsStore.load(DEFAULT_SETTINGS));
        bindDialogEvents(dialogControls, textTargets);
        prepareDialogWindow(dialogControls.dialog, SCRIPT_NAME);
        if (dialogControls.dialog.show() !== 1) return;

        var formatOptions = readDialogOptions(dialogControls);
        settingsStore.save(formatOptions);
        var failureReason = runHorizontalTextConversion(doc, textTargets, formatOptions);
        if (failureReason !== "") alert(getLabel("alert.actionFailed") + "\n\n" + failureReason);
    }

    main();

})();
