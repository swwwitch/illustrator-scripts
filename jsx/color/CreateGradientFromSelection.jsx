#target illustrator
#targetengine "CreateGradientFromSelectionEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの塗り／線カラーを配置順（左→右、上→下）で抽出し、スウォッチグループに登録してグラデーションを自動生成します。
複製でブレンドを作ることもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGradientFromSelection.md

### Overview

Extracts the fill and stroke colors of the selection in layout order (left to right, top to bottom), registers them as a swatch group, and builds a gradient from them.
It can also blend duplicates of the selection.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGradientFromSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CreateGradientFromSelection";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.10.2";                      /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGradientFromSelection.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGradientFromSelection.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* セパレートグラデーションで許可する最大色数 / Max colors allowed for Separate gradients */
    var SEPARATE_MAX_COLORS = 6;

    /* 長方形サイズの既定値（pt）/ Default rectangle size in points */
    var DEFAULT_RECT_SIZE = 100;

    /* スウォッチ由来時の固定サイズ（pt）/ Fixed size when invoked from swatches */
    var SWATCH_RECT_WIDTH = 200;
    var SWATCH_RECT_HEIGHT = 100;

    /* 新規スウォッチグループの既定名 / Default name for the new swatch group */
    var SWATCH_GROUP_BASE_NAME = "AutoGradient";
    var SWATCH_BASE_NAME = "AutoColor";
    var GRADIENT_BASE_NAME = "New Gradient";

    /* ダイアログのデフォルトオプション / Default dialog values */
    var DEFAULT_OPTIONS = {
        makeGlobal: true,
        makeGradient: true,
        makeRect: true,
        useSelectionSize: true,
        registerGraphicStyle: true,
        separateGradient: false,
        makeBlend: false
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネルの余白 [左, 上, 右, 下] / Panel margins [left, top, right, bottom] */
    var PANEL_MARGINS = [15, 20, 15, 10];

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        dialog: {
            title: { ja: "グラデーション作成", en: "Create Gradient" }
        },
        panel: {
            color: { ja: "カラー", en: "Colors" },
            rect: { ja: "長方形", en: "Rectangle" }
        },
        checkbox: {
            globalColor: { ja: "グローバルカラー化", en: "Make Global colors" },
            createGradient: { ja: "グラデーションを作成", en: "Create gradient" },
            createRect: { ja: "長方形を作成してグラデーションを適用", en: "Create rectangle and apply gradient" },
            useSelectionSize: { ja: "選択オブジェクトのサイズに合わせる", en: "Match selection size" },
            registerGraphicStyle: { ja: "グラフィックスタイルとして登録", en: "Save as Graphic Style" },
            makeBlend: { ja: "複製でブレンドを作成", en: "Blend duplicates" }
        },
        radio: {
            normal: { ja: "通常", en: "Smooth" },
            separate: { ja: "セパレート", en: "Segmented" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            globalColor: {
                ja: "スウォッチをグローバルカラー（プロセス）として登録し、後から一括で色を変更可能にします。",
                en: "Register swatches as Global Process colors so they can be edited together later."
            },
            createGradient: {
                ja: "抽出した色を並べた線形グラデーションを作ります。OFF のときはスウォッチの登録だけを行います。",
                en: "Builds a linear gradient from the extracted colors. When off, only the swatches are registered."
            },
            normal: {
                ja: "色を等間隔に置き、なめらかにつなぐグラデーションにします。",
                en: "Places the colors evenly and blends them smoothly."
            },
            separate: {
                ja: "色の境界をくっきり分割するグラデーション（最大 6 色）。選択が 7 つ以上のときは使えません。",
                en: "Hard-edged segmented gradient (up to 6 colors). Disabled when 7 or more items are selected."
            },
            createRect: {
                ja: "作ったグラデーションで塗った長方形を描きます。横並びの選択なら下、縦並びなら右に、長方形1個分の間隔をあけて置きます。",
                en: "Draws a rectangle filled with the new gradient, leaving a one-rectangle gap below a horizontal selection or to the right of a vertical one."
            },
            useSelectionSize: {
                ja: "選択オブジェクトの外接サイズに合わせて長方形を作成します。",
                en: "Use the bounding size of the selection for the rectangle."
            },
            registerGraphicStyle: {
                ja: "作成した長方形の見た目をグラフィックスタイルに登録します。長方形作成 OFF のときは一時長方形で登録します。",
                en: "Register the rectangle's appearance as a Graphic Style. When rectangle output is off, a temporary rectangle is used."
            },
            makeBlend: {
                ja: "選択オブジェクトを複製し、複製でブレンドを作ります（元のオブジェクトはそのまま）。スウォッチから色を集めたときと、選択が1つのときは使えません。",
                en: "Duplicates the selected objects and blends the duplicates, leaving the originals as they are. Unavailable when colors come from swatches or only one object is selected."
            }
        }
    };

    // =========================================
    // セッション設定 / Session Settings
    // =========================================

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

    var settingsStore = createSettingsStore("CreateGradientFromSelection", "session");

    /**
     * ダイアログの値を保持するオブジェクトを返す（targetengine 内だけで保持し、Illustrator の再起動で消える）
     * 既定値は {}（項目ごとに loadBool で既定値を補う）
     * @returns {Object} 保存してある設定（毎回新しいオブジェクト）
     */
    function getSessionSettings() {
        return settingsStore.load({});
    }

    /**
     * 保持している真偽値を読み出す
     * @param {string} settingKey - 設定のキー
     * @param {boolean} defaultValue - 保持していないときの値
     * @returns {boolean} 保持している値、または defaultValue
     */
    function loadBool(settingKey, defaultValue) {
        var sessionSettings = getSessionSettings();
        if (typeof sessionSettings[settingKey] === 'boolean') return sessionSettings[settingKey];
        return defaultValue;
    }

    /**
     * 真偽値を保持する
     * @param {string} settingKey - 設定のキー
     * @param {boolean} value - 保持する値
     * @returns {void}
     */
    function saveBool(settingKey, value) {
        /* ほかの項目を消さないよう、読み込んで書き足す / Read, add and write back so the other keys survive */
        var sessionSettings = getSessionSettings();
        sessionSettings[settingKey] = !!value;
        settingsStore.save(sessionSettings);
    }

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
    // 選択範囲の解析 / Selection Analysis
    // =========================================

    /**
     * 現在の選択を配列に写し取る
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択中のオブジェクト（選択が無ければ空配列）
     */
    function snapshotSelection(doc) {
        var selectedItems = [];
        var currentSelection = doc.selection;
        if (currentSelection && currentSelection.length) {
            for (var i = 0; i < currentSelection.length; i++) selectedItems.push(currentSelection[i]);
        }
        return selectedItems;
    }

    /**
     * アイテムの左上座標を取得する
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {{left: number, top: number}} 左上座標（取得できないときは 0, 0）
     */
    function getItemTopLeft(item) {
        try {
            var bounds = item.geometricBounds; /* [left, top, right, bottom] */
            return { left: bounds[0], top: bounds[1] };
        } catch (e) {
            return { left: 0, top: 0 };
        }
    }

    /**
     * 選択全体の外接バウンディングを取得する
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {{left: number, top: number, right: number, bottom: number}|null} 外接矩形（求められないときは null）
     */
    function getSelectionBounds(selectedItems) {
        /* クリップグループはマスクで測る。空白だけのテキストなど境界が取れないものは飛ばす
           Clip groups are measured by their mask; items without bounds (whitespace-only text etc.) are skipped */
        var unionBounds = getClipAwareUnionBounds(selectedItems, false);
        if (!unionBounds) return null;
        var left = unionBounds[0], top = unionBounds[1], right = unionBounds[2], bottom = unionBounds[3];
        if (left > right || bottom > top) return null;
        return { left: left, top: top, right: right, bottom: bottom };
    }

    /**
     * 選択オブジェクトが横並びか縦並びかを、各オブジェクトの中心の散らばりから判定する
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {string} "horizontal" / "vertical" / "mixed" / "unknown"
     */
    function detectSelectionOrientation(selectedItems) {
        if (!selectedItems || selectedItems.length < 2) return "unknown";

        var minX = 1e12, maxX = -1e12;
        var minY = 1e12, maxY = -1e12;

        for (var i = 0; i < selectedItems.length; i++) {
            try {
                var bounds = selectedItems[i].geometricBounds;
                var centerX = (bounds[0] + bounds[2]) / 2;
                var centerY = (bounds[1] + bounds[3]) / 2;
                if (centerX < minX) minX = centerX;
                if (centerX > maxX) maxX = centerX;
                if (centerY < minY) minY = centerY;
                if (centerY > maxY) maxY = centerY;
            } catch (e) { /* bounds を取得できないものは無視 / skip items without bounds */ }
        }

        if (minX > maxX || minY > maxY) return "unknown";

        var dx = Math.abs(maxX - minX);
        var dy = Math.abs(maxY - minY);
        if (dx > dy) return "horizontal";
        if (dy > dx) return "vertical";
        return "mixed";
    }

    // =========================================
    // ブレンド / Blend
    // =========================================

    /**
     * 選択を指定したオブジェクトに戻す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択し直すオブジェクト
     * @returns {void}
     */
    function restoreSelection(doc, selectedItems) {
        try {
            doc.selection = null;
            if (selectedItems && selectedItems.length) doc.selection = selectedItems;
        } catch (e) { /* 削除済み・ロック中のオブジェクトは選択できない / Removed or locked items cannot be selected */ }
    }

    /**
     * 選択オブジェクトを複製し、複製側でブレンドを作る（元の選択はそのまま残す）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function duplicateSelectionAndBlend(doc) {
        try {
            if (!doc.selection || doc.selection.length < 2) return;

            var originalItems = snapshotSelection(doc);

            /* 複製する / Duplicate the items */
            var duplicatedItems = [];
            for (var i = 0; i < originalItems.length; i++) {
                try { duplicatedItems.push(originalItems[i].duplicate()); } catch (eD) { }
            }
            if (duplicatedItems.length < 2) {
                restoreSelection(doc, originalItems);
                return;
            }

            /* 複製だけを選択する / Select only the duplicates */
            doc.selection = null;
            for (var j = 0; j < duplicatedItems.length; j++) {
                /* 非表示・ロックを引き継いだ複製は選択できない / Duplicates that inherit hidden or locked cannot be selected */
                try { duplicatedItems[j].selected = true; } catch (eS) { }
            }

            /* ブレンドのメニューコマンド名は環境で異なるため順に試す / Try known Blend menu commands (varies by locale/version) */
            try {
                app.executeMenuCommand('Make Blend');
            } catch (e1) {
                try {
                    app.executeMenuCommand('Blend Make');
                } catch (e2) {
                    try { app.executeMenuCommand('blend'); } catch (e3) { }
                }
            }

            restoreSelection(doc, originalItems);
        } catch (e) {
            /* ブレンドを作れなくても無言で続行 / Continue silently when the blend fails */
        }
    }

    // =========================================
    // カラーユーティリティ / Color Utilities
    // =========================================

    /**
     * 「なし」の色かどうかを判定する
     * @param {Color} color - 判定する色
     * @returns {boolean} null または NoColor なら true
     */
    function isNoColor(color) {
        try {
            return (color == null) || (color.typename === "NoColor");
        } catch (e) {
            return true;
        }
    }

    /**
     * 重複除去用のカラーキーを作る
     * @param {Color} color - 対象の色
     * @returns {string} 同じ色なら同じになるキー
     */
    function colorKey(color) {
        if (!color) return "null";
        var typeName = color.typename;
        try {
            if (typeName === "RGBColor") return "RGB:" + [color.red, color.green, color.blue].join(",");
            if (typeName === "CMYKColor") return "CMYK:" + [color.cyan, color.magenta, color.yellow, color.black].join(",");
            if (typeName === "GrayColor") return "Gray:" + color.gray;
            if (typeName === "SpotColor") {
                var spotName = (color.spot && color.spot.name) ? color.spot.name : "(spot)";
                return "Spot:" + spotName + ":" + color.tint;
            }
            if (typeName === "PatternColor") {
                var patternName = (color.pattern && color.pattern.name) ? color.pattern.name : "(pattern)";
                return "Pattern:" + patternName;
            }
            if (typeName === "GradientColor") {
                var gradientName = (color.gradient && color.gradient.name) ? color.gradient.name : "(gradient)";
                return "Gradient:" + gradientName;
            }
        } catch (e) { /* 名前を読めない色は種類だけのキーにする / fall back to the type name */ }
        return "Other:" + typeName;
    }

    /**
     * 選択から色を読むオブジェクトを集める（グループ・複合パスの中まで降りる）
     * @param {*} selectedItems - 選択中のオブジェクト
     * @returns {PageItem[]} グループ・複合パス以外のオブジェクト（前面→背面の順）
     */
    function collectColorLeafItems(selectedItems) {
        return collectSelectionItems(selectedItems, {
            enterCompoundPaths: true,
            textRangeToFrame: false,
            accept: function (item) {
                return item.typename !== "GroupItem" && item.typename !== "CompoundPathItem";
            }
        });
    }

    /**
     * オブジェクト1件から塗り・線の色を位置付きで集める（グループ・複合パスの中は collectColorLeafItems で展開済み）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {Object[]} outEntries - {left, top, color} を追加する配列
     * @returns {void}
     */
    function collectColorEntries(item, outEntries) {
        if (!item) return;

        try {
            if (item.typename === "TextFrame") {
                var textPos = getItemTopLeft(item);
                var textFill = item.textRange.characterAttributes.fillColor;
                if (!isNoColor(textFill)) outEntries.push({ left: textPos.left, top: textPos.top, color: textFill });
                try {
                    var textStroke = item.textRange.characterAttributes.strokeColor;
                    if (!isNoColor(textStroke)) outEntries.push({ left: textPos.left, top: textPos.top, color: textStroke });
                } catch (e) { /* 線の色を読めないテキストは塗りだけ / fill only when the stroke is unreadable */ }
                return;
            }

            if (typeof item.filled !== "undefined" || typeof item.stroked !== "undefined") {
                var pathPos = getItemTopLeft(item);
                if (item.filled) {
                    var pathFill = item.fillColor;
                    if (!isNoColor(pathFill)) outEntries.push({ left: pathPos.left, top: pathPos.top, color: pathFill });
                }
                if (item.stroked) {
                    var pathStroke = item.strokeColor;
                    if (!isNoColor(pathStroke)) outEntries.push({ left: pathPos.left, top: pathPos.top, color: pathStroke });
                }
            }
        } catch (e) { /* 取得できないアイテムは無視 / skip unreadable items */ }
    }

    /**
     * 選択から重複を除いた色を、長辺方向（横なら左→右／縦なら上→下）の順で返す
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @param {string} selectionOrientation - detectSelectionOrientation() の結果
     * @returns {Color[]} 重複を除いた色
     */
    function collectColorsFromSelection(selectedItems, selectionOrientation) {
        var entries = [];
        var leafItems = collectColorLeafItems(selectedItems);
        for (var i = 0; i < leafItems.length; i++) {
            collectColorEntries(leafItems[i], entries);
        }

        /* 縦並び（dy > dx）なら上→下を優先キーに / Use top→bottom as primary key when vertical */
        var verticalPrimary = (selectionOrientation === "vertical");

        entries.sort(function (a, b) {
            if (verticalPrimary) {
                if (a.top > b.top) return -1;
                if (a.top < b.top) return 1;
                if (a.left < b.left) return -1;
                if (a.left > b.left) return 1;
                return 0;
            }
            if (a.left < b.left) return -1;
            if (a.left > b.left) return 1;
            if (a.top > b.top) return -1;
            if (a.top < b.top) return 1;
            return 0;
        });

        var colors = [];
        var seenKeys = {};
        for (var k = 0; k < entries.length; k++) {
            var entryColor = entries[k].color;
            if (isNoColor(entryColor)) continue;
            var entryKey = colorKey(entryColor);
            if (seenKeys[entryKey]) continue;
            seenKeys[entryKey] = true;
            colors.push(entryColor);
        }
        return colors;
    }

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

    // =========================================
    // アクション定義 / Action Definitions
    // =========================================

    /**
     * 選択中のオブジェクトのグラデーション角度を 90° にするアクションを実行する
     * @returns {void}
     */
    function runGradientAngle90Action() {
        var CR = String.fromCharCode(13);
        var actionCode = [
            " /version 3",
            "/name [ 8",
            "\t6772616469656e74",
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "\t/name [ 8",
            "\t\t3930646567726565",
            "\t]",
            "\t/keyIndex 0",
            "\t/colorIndex 0",
            "\t/isOpen 1",
            "\t/eventCount 1",
            "\t/event-1 {",
            "\t\t/useRulersIn1stQuadrant 0",
            "\t\t/internalName (ai_plugin_setGradient)",
            "\t\t/localizedName [ 30",
            "\t\t\te382b0e383a9e38387e383bce382b7e383a7e383b3e38292e8a8ade5ae9a",
            "\t\t]",
            "\t\t/isOpen 1",
            "\t\t/isOn 1",
            "\t\t/hasDialog 0",
            "\t\t/parameterCount 1",
            "\t\t/parameter-1 {",
            "\t\t\t/key 1634625388",
            "\t\t\t/showInPalette 4294967295",
            "\t\t\t/type (unit real)",
            "\t\t\t/value -90.0",
            "\t\t\t/unit 591490663",
            "\t\t}",
            "\t}",
            "}",
            ""
        ].join(CR);
        runTemporaryAction(actionCode, "gradient", "90degree");
    }

    /**
     * 「新規グラフィックスタイル」をアクションで実行する
     * @returns {void}
     */
    function runGraphicStyleAction() {
        var CR = String.fromCharCode(13);
        var actionCode = [
            " /version 3",
            "/name [ 12",
            "\t477261706869635374796c65",
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "\t/name [ 3",
            "\t\t6e6577",
            "\t]",
            "\t/keyIndex 0",
            "\t/colorIndex 0",
            "\t/isOpen 1",
            "\t/eventCount 1",
            "\t/event-1 {",
            "\t\t/useRulersIn1stQuadrant 0",
            "\t\t/internalName (ai_plugin_styles)",
            "\t\t/localizedName [ 30",
            "\t\t\te382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab",
            "\t\t]",
            "\t\t/isOpen 1",
            "\t\t/isOn 1",
            "\t\t/hasDialog 1",
            "\t\t/showDialog 0",
            "\t\t/parameterCount 1",
            "\t\t/parameter-1 {",
            "\t\t\t/key 1835363957",
            "\t\t\t/showInPalette 4294967295",
            "\t\t\t/type (enumerated)",
            "\t\t\t/name [ 36",
            "\t\t\t\te696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382",
            "\t\t\t\ta4e383ab",
            "\t\t\t]",
            "\t\t\t/value 1",
            "\t\t}",
            "\t}",
            "}",
            ""
        ].join(CR);
        runTemporaryAction(actionCode, "GraphicStyle", "new");
    }

    // =========================================
    // スウォッチ操作 / Swatch Operations
    // =========================================

    /**
     * コレクションに同じ名前の項目があるかを調べる
     * @param {Object} namedCollection - doc.swatches / doc.swatchGroups / doc.gradients など getByName を持つコレクション
     * @param {string} itemName - 調べる名前
     * @returns {boolean} あれば true
     */
    function hasItemNamed(namedCollection, itemName) {
        try {
            namedCollection.getByName(itemName); /* 見つからないと例外 / throws when not found */
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 名前の衝突を避けた一意な名前を作る（"名前", "名前 1", "名前 2" …）
     * @param {string} baseName - 元の名前
     * @param {Object} namedCollection - 名前の重複を調べるコレクション
     * @returns {string} 使われていない名前
     */
    function uniqueName(baseName, namedCollection) {
        var candidateName = baseName;
        var suffix = 1;
        while (hasItemNamed(namedCollection, candidateName)) {
            candidateName = baseName + " " + suffix;
            suffix++;
        }
        return candidateName;
    }

    /**
     * 色をグローバルカラー（プロセス）に変換する
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} baseColor - 元の色
     * @param {string} spotName - 作るスポットの名前
     * @returns {Color} グローバルカラー（失敗したときは元の色）
     */
    function toGlobalProcessColor(doc, baseColor, spotName) {
        try {
            var globalSpot = doc.spots.add();
            globalSpot.name = spotName;
            globalSpot.colorType = ColorModel.PROCESS;
            globalSpot.color = baseColor;

            var spotColor = new SpotColor();
            spotColor.spot = globalSpot;
            spotColor.tint = 100;
            return spotColor;
        } catch (e) {
            return baseColor;
        }
    }

    /**
     * 1色をスウォッチに登録する（必要ならグローバルカラーにする）
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} sourceColor - 登録する色
     * @param {string} baseName - スウォッチ名の元
     * @param {boolean} makeGlobal - グローバルカラーにするか
     * @returns {Swatch} 作ったスウォッチ
     */
    function addSwatchForColor(doc, sourceColor, baseName, makeGlobal) {
        var swatch = doc.swatches.add();
        var swatchName = uniqueName(baseName, doc.swatches);
        swatch.name = swatchName;
        swatch.color = makeGlobal ? toGlobalProcessColor(doc, sourceColor, swatchName) : sourceColor;
        try { swatch.selected = false; } catch (e) { /* 無視 / ignore */ }
        return swatch;
    }

    /**
     * 新しいスウォッチグループを作り、色をスウォッチとして登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 登録する色
     * @param {boolean} makeGlobal - グローバルカラーにするか
     * @returns {Swatch[]} 作ったスウォッチ（colors と同じ順）
     */
    function registerColorSwatches(doc, colors, makeGlobal) {
        var groupName = uniqueName(SWATCH_GROUP_BASE_NAME, doc.swatchGroups);
        var swatchGroup = doc.swatchGroups.add();
        swatchGroup.name = groupName;

        var createdSwatches = [];
        for (var i = 0; i < colors.length; i++) {
            var createdSwatch = addSwatchForColor(doc, colors[i], SWATCH_BASE_NAME, makeGlobal);
            createdSwatches.push(createdSwatch);
            try { swatchGroup.addSwatch(createdSwatch); } catch (e) { /* 無視 / ignore */ }
        }
        return createdSwatches;
    }

    // =========================================
    // レイヤー操作 / Layer Helpers
    // =========================================

    /**
     * 描画できるレイヤー（ロックも非表示もされていないもの）を返す。アクティブレイヤーを優先する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer|null} 描画できるレイヤー（無ければ null）
     */
    function getUnlockedVisibleLayer(doc) {
        var activeLayer = doc.activeLayer;
        if (activeLayer && !activeLayer.locked && activeLayer.visible) return activeLayer;
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (!layer.locked && layer.visible) return layer;
        }
        return null;
    }

    // =========================================
    // グラフィックスタイル登録 / Graphic Style Registration
    // =========================================

    /**
     * 選択中のオブジェクトの見た目を、1つずつ新規グラフィックスタイルとして登録する（名前は既定のまま）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function registerGraphicStyleFromSelected(doc) {
        if (!doc.graphicStyles) return;

        var selectedItems = snapshotSelection(doc);
        for (var i = 0; i < selectedItems.length; i++) {
            /* 1つだけ選択してアクションで登録 / Select just this item and register it via the action */
            doc.selection = null;
            try { doc.selection = [selectedItems[i]]; }
            catch (e) { try { selectedItems[i].selected = true; } catch (e2) { /* 無視 / ignore */ } }

            runGraphicStyleAction();
        }
    }

    // =========================================
    // 入力収集 / Input Collection
    // =========================================

    /**
     * 選択オブジェクト、または（選択が無ければ）選択中のスウォッチから、色と関連情報を集める
     * @param {Document} doc - 対象ドキュメント
     * @returns {{colors: Color[], fromSwatches: boolean, selectionBounds: Object, selectionOrientation: string, itemCount: number}} 集めた色と関連情報
     */
    function collectInputColors(doc) {
        var colorInput = {
            colors: [],
            fromSwatches: false,
            selectionBounds: null,
            selectionOrientation: "unknown",
            itemCount: 0
        };

        if (doc.selection && doc.selection.length > 0) {
            colorInput.selectionOrientation = detectSelectionOrientation(doc.selection);
            colorInput.colors = collectColorsFromSelection(doc.selection, colorInput.selectionOrientation);
            colorInput.selectionBounds = getSelectionBounds(doc.selection);
            colorInput.itemCount = doc.selection.length;
            return colorInput;
        }

        var selectedSwatches = null;
        try { selectedSwatches = doc.swatches.getSelected(); } catch (e) { selectedSwatches = null; }
        if (!selectedSwatches || selectedSwatches.length < 2) return colorInput;

        colorInput.fromSwatches = true;
        for (var i = 0; i < selectedSwatches.length; i++) {
            try {
                var swatchColor = selectedSwatches[i].color;
                if (swatchColor && swatchColor.typename !== "NoColor") colorInput.colors.push(swatchColor);
            } catch (e) { /* 無視 / ignore */ }
        }
        colorInput.itemCount = colorInput.colors.length;
        return colorInput;
    }

    // =========================================
    // ダイアログ / Options Dialog
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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）とき、行を左右中央に並べ直す。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function centerButtonRowIfRightOnly(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var btnRowGroup = buttonRow.rowGroup;
        /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
        btnRowGroup.remove(buttonRow.leftGroup);
        btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        buttonRow.leftGroup = null;
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    /**
     * 前回の値（無ければ既定値）からダイアログの初期値を作る
     * @param {boolean} disallowSeparate - セパレートを選べないとき true（常に OFF にする）
     * @returns {Object} makeGlobal / makeGradient / makeRect / useSelectionSize / registerGraphicStyle / separateGradient / makeBlend
     */
    function loadDialogOptions(disallowSeparate) {
        return {
            makeGlobal: loadBool('makeGlobal', DEFAULT_OPTIONS.makeGlobal),
            makeGradient: loadBool('makeGradient', DEFAULT_OPTIONS.makeGradient),
            makeRect: loadBool('makeRect', DEFAULT_OPTIONS.makeRect),
            useSelectionSize: loadBool('useSelectionSize', DEFAULT_OPTIONS.useSelectionSize),
            registerGraphicStyle: loadBool('registerGraphicStyle', DEFAULT_OPTIONS.registerGraphicStyle),
            separateGradient: (disallowSeparate ? false : loadBool('separateGradient', DEFAULT_OPTIONS.separateGradient)),
            makeBlend: loadBool('makeBlend', DEFAULT_OPTIONS.makeBlend)
        };
    }

    /**
     * 縦並びのパネルを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} titlePath - パネル名のラベルのパス
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentWindow, titlePath) {
        var optionPanel = parentWindow.add('panel', undefined, getLabel(titlePath));
        optionPanel.orientation = 'column';
        optionPanel.alignChildren = ['fill', 'top'];
        optionPanel.margins = PANEL_MARGINS;
        return optionPanel;
    }

    /**
     * オプションダイアログを表示し、確定値を返す
     * @param {boolean} disallowSeparate - セパレートを選べないとき true
     * @param {boolean} fromSwatches - 色をスウォッチから集めたとき true（［選択オブジェクトのサイズに合わせる］を使えない）
     * @param {boolean} canBlend - ブレンドを作れるとき true（選択オブジェクトが2つ以上）
     * @returns {Object|null} 確定したオプション（キャンセル時は null）
     */
    function showOptionsDialog(disallowSeparate, fromSwatches, canBlend) {
        var initialOptions = loadDialogOptions(disallowSeparate);

        var optionsDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        optionsDialog.orientation = 'column';
        optionsDialog.alignChildren = ['fill', 'top'];

        /* カラー関連パネル / Color-related panel */
        var colorPanel = addOptionPanel(optionsDialog, 'panel.color');

        var globalColorCheckbox = colorPanel.add('checkbox', undefined, getLabel('checkbox.globalColor'));
        globalColorCheckbox.value = initialOptions.makeGlobal;
        globalColorCheckbox.helpTip = getLabel('tooltip.globalColor');

        var gradientCheckbox = colorPanel.add('checkbox', undefined, getLabel('checkbox.createGradient'));
        gradientCheckbox.value = initialOptions.makeGradient;
        gradientCheckbox.helpTip = getLabel('tooltip.createGradient');

        var gradientTypeGroup = colorPanel.add('group');
        gradientTypeGroup.orientation = 'row';
        gradientTypeGroup.alignChildren = ['left', 'center'];

        var normalRadio = gradientTypeGroup.add('radiobutton', undefined, getLabel('radio.normal'));
        normalRadio.helpTip = getLabel('tooltip.normal');
        var separateRadio = gradientTypeGroup.add('radiobutton', undefined, getLabel('radio.separate'));
        separateRadio.helpTip = getLabel('tooltip.separate');
        separateRadio.value = !!initialOptions.separateGradient;
        normalRadio.value = !separateRadio.value;

        var blendCheckbox = colorPanel.add('checkbox', undefined, getLabel('checkbox.makeBlend'));
        blendCheckbox.value = canBlend && initialOptions.makeBlend;
        blendCheckbox.enabled = canBlend;
        blendCheckbox.helpTip = getLabel('tooltip.makeBlend');

        /* 長方形パネル / Rectangle panel */
        var rectPanel = addOptionPanel(optionsDialog, 'panel.rect');

        var rectCheckbox = rectPanel.add('checkbox', undefined, getLabel('checkbox.createRect'));
        rectCheckbox.value = initialOptions.makeRect;
        rectCheckbox.helpTip = getLabel('tooltip.createRect');

        var selectionSizeCheckbox = rectPanel.add('checkbox', undefined, getLabel('checkbox.useSelectionSize'));
        selectionSizeCheckbox.value = initialOptions.useSelectionSize;
        selectionSizeCheckbox.helpTip = getLabel('tooltip.useSelectionSize');

        var graphicStyleCheckbox = rectPanel.add('checkbox', undefined, getLabel('checkbox.registerGraphicStyle'));
        graphicStyleCheckbox.value = initialOptions.registerGraphicStyle;
        graphicStyleCheckbox.helpTip = getLabel('tooltip.registerGraphicStyle');

        /**
         * チェック状態の連動（グラデーションを作らないときは長方形・スタイル・種類を OFF にしてディム）
         * @returns {void}
         */
        function syncEnable() {
            rectCheckbox.enabled = gradientCheckbox.value;
            selectionSizeCheckbox.enabled = gradientCheckbox.value && rectCheckbox.value && !fromSwatches;
            graphicStyleCheckbox.enabled = gradientCheckbox.value;

            normalRadio.enabled = gradientCheckbox.value;
            separateRadio.enabled = gradientCheckbox.value && !disallowSeparate;
            gradientTypeGroup.enabled = gradientCheckbox.value;

            if (!gradientCheckbox.value) {
                rectCheckbox.value = false;
                selectionSizeCheckbox.value = false;
                graphicStyleCheckbox.value = false;
            }
            if (fromSwatches) selectionSizeCheckbox.value = false;
            if (!gradientCheckbox.value || disallowSeparate) {
                separateRadio.value = false;
                normalRadio.value = true;
            }
        }
        gradientCheckbox.onClick = syncEnable;
        rectCheckbox.onClick = syncEnable;
        syncEnable();

        /* OK／キャンセル / OK and Cancel */
        var buttonRow = addButtonRow(optionsDialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        /**
         * チェック状態をセッション設定に保存する
         * @returns {void}
         */
        function persistFromUI() {
            saveBool('makeGlobal', globalColorCheckbox.value);
            saveBool('makeGradient', gradientCheckbox.value);
            saveBool('makeRect', rectCheckbox.value);
            saveBool('useSelectionSize', selectionSizeCheckbox.value);
            saveBool('registerGraphicStyle', graphicStyleCheckbox.value);
            saveBool('separateGradient', (disallowSeparate ? false : separateRadio.value));
            /* 使えないときの OFF は記録しない / Do not record the forced OFF when blending is unavailable */
            if (canBlend) saveBool('makeBlend', blendCheckbox.value);
        }
        optionsDialog.onClose = persistFromUI;

        centerButtonRowIfRightOnly(buttonRow);
        prepareDialogWindow(optionsDialog, SCRIPT_NAME);
        if (optionsDialog.show() !== 1) return null;

        return {
            makeGlobal: !!globalColorCheckbox.value,
            makeGradient: !!gradientCheckbox.value,
            makeRect: !!rectCheckbox.value,
            useSelectionSize: !!selectionSizeCheckbox.value,
            registerGraphicStyle: !!graphicStyleCheckbox.value,
            separateGradient: (disallowSeparate ? false : !!separateRadio.value),
            makeBlend: canBlend && !!blendCheckbox.value
        };
    }

    // =========================================
    // グラデーション生成 / Gradient Construction
    // =========================================

    /**
     * ストップに使う色を返す（作成済みスウォッチの色を優先し、無ければ元の色）
     * @param {Swatch[]} createdSwatches - 登録したスウォッチ
     * @param {Color[]} colors - 元の色
     * @param {number} index - 色の番号
     * @returns {Color} ストップに使う色
     */
    function pickStopColor(createdSwatches, colors, index) {
        try {
            if (createdSwatches && createdSwatches[index] && createdSwatches[index].color) {
                return createdSwatches[index].color;
            }
        } catch (e) { /* 無視 / ignore */ }
        return colors[index];
    }

    /**
     * グラデーションのストップ数を targetCount に合わせる
     * @param {Gradient} gradient - 対象のグラデーション
     * @param {number} targetCount - ストップ数
     * @returns {void}
     */
    function resizeGradientStops(gradient, targetCount) {
        while (gradient.gradientStops.length < targetCount) gradient.gradientStops.add();
        while (gradient.gradientStops.length > targetCount) {
            gradient.gradientStops[gradient.gradientStops.length - 1].remove();
        }
    }

    /**
     * 色を等間隔に並べた通常（スムーズ）グラデーションを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 並べる色
     * @param {Swatch[]} createdSwatches - colors と同じ順に登録したスウォッチ
     * @returns {Gradient} 作ったグラデーション
     */
    function buildNormalGradient(doc, colors, createdSwatches) {
        var gradient = doc.gradients.add();
        gradient.type = GradientType.LINEAR;
        resizeGradientStops(gradient, colors.length);

        for (var i = 0; i < colors.length; i++) {
            var stop = gradient.gradientStops[i];
            stop.rampPoint = (i / (colors.length - 1)) * 100;
            try { stop.color = pickStopColor(createdSwatches, colors, i); }
            catch (e) { try { stop.color = colors[i]; } catch (e2) { /* 無視 / ignore */ } }
            stop.midPoint = 50;
            stop.opacity = 100;
        }

        gradient.name = uniqueName(GRADIENT_BASE_NAME, doc.gradients);
        return gradient;
    }

    /**
     * セパレート（境界がくっきり）グラデーションを作る（2〜SEPARATE_MAX_COLORS 色）
     * 境界ごとに左右 0.01% の位置へ同じ色の組を置き、色の帯を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 並べる色
     * @param {Swatch[]} createdSwatches - colors と同じ順に登録したスウォッチ
     * @returns {Gradient} 作ったグラデーション
     */
    function buildSeparateGradient(doc, colors, createdSwatches) {
        var colorCount = colors.length;
        var epsilon = 0.01;
        var bandWidth = 100 / colorCount;
        if (colorCount === 3 || colorCount === 6) bandWidth = Math.round(bandWidth * 10) / 10;

        var stopPoints = [0];
        var stopColors = [pickStopColor(createdSwatches, colors, 0)];

        for (var k = 1; k <= colorCount - 1; k++) {
            var boundary = bandWidth * k;
            if (colorCount === 3 || colorCount === 6) boundary = Math.round(boundary * 10) / 10;
            var leftPoint = Math.max(0, boundary - epsilon);
            var rightPoint = Math.min(100, boundary + epsilon);

            stopPoints.push(leftPoint);
            stopColors.push(pickStopColor(createdSwatches, colors, k - 1));
            stopPoints.push(rightPoint);
            stopColors.push(pickStopColor(createdSwatches, colors, k));
        }

        stopPoints.push(100);
        stopColors.push(pickStopColor(createdSwatches, colors, colorCount - 1));

        var gradient = doc.gradients.add();
        gradient.type = GradientType.LINEAR;
        resizeGradientStops(gradient, stopPoints.length);

        for (var i = 0; i < stopPoints.length; i++) {
            var stop = gradient.gradientStops[i];
            stop.rampPoint = stopPoints[i];
            try { stop.color = stopColors[i]; }
            catch (e) {
                try { stop.color = colors[Math.min(colors.length - 1, Math.max(0, Math.floor(i / 2)))]; } catch (e2) { /* 無視 / ignore */ }
            }
            stop.midPoint = 50;
            stop.opacity = 100;
        }
        return gradient;
    }

    // =========================================
    // 長方形配置 / Rectangle Placement
    // =========================================

    /**
     * 長方形のサイズを決める
     * @param {Object} gradientOptions - ダイアログのオプション
     * @param {Object} colorInput - collectInputColors() の結果
     * @returns {{width: number, height: number}} 長方形のサイズ
     */
    function computeRectSize(gradientOptions, colorInput) {
        if (colorInput.fromSwatches) return { width: SWATCH_RECT_WIDTH, height: SWATCH_RECT_HEIGHT };
        if (gradientOptions.useSelectionSize && colorInput.selectionBounds) {
            var boundsWidth = Math.abs(colorInput.selectionBounds.right - colorInput.selectionBounds.left);
            var boundsHeight = Math.abs(colorInput.selectionBounds.top - colorInput.selectionBounds.bottom);
            if (boundsWidth > 0 && boundsHeight > 0) return { width: boundsWidth, height: boundsHeight };
        }
        return { width: DEFAULT_RECT_SIZE, height: DEFAULT_RECT_SIZE };
    }

    /**
     * 長方形の配置（左上座標）を決める
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} colorInput - collectInputColors() の結果
     * @param {{width: number, height: number}} rectSize - 長方形のサイズ
     * @returns {{left: number, top: number}} 長方形の左上
     */
    function computeRectPosition(doc, colorInput, rectSize) {
        var viewCenter = doc.activeView.centerPoint;
        var left = viewCenter[0] - rectSize.width / 2;
        var top = viewCenter[1] + rectSize.height / 2;

        if (!colorInput.fromSwatches && colorInput.selectionBounds) {
            if (colorInput.selectionOrientation === "horizontal") {
                /* 横並び: 選択の左端揃え／真下に 1 個分離す / Horizontal: align to left edge, offset below */
                left = colorInput.selectionBounds.left;
                top = colorInput.selectionBounds.bottom - rectSize.height;
            } else if (colorInput.selectionOrientation === "vertical") {
                /* 縦並び: 選択の上端揃え／右に 1 個分離す / Vertical: align to top edge, offset to right */
                left = colorInput.selectionBounds.right + rectSize.width;
                top = colorInput.selectionBounds.top;
            }
        }
        return { left: left, top: top };
    }

    /**
     * 長方形を作ってグラデーションを適用し、必要ならグラフィックスタイルに登録する
     * 長方形を出力しない設定では、一時レイヤーの一時長方形でスタイルを登録してから片付ける
     * @param {Document} doc - 対象ドキュメント
     * @param {Gradient} gradient - 適用するグラデーション
     * @param {Object} gradientOptions - ダイアログのオプション
     * @param {Object} colorInput - collectInputColors() の結果
     * @returns {void}
     */
    function createGradientRect(doc, gradient, gradientOptions, colorInput) {
        var targetLayer = getUnlockedVisibleLayer(doc);
        if (!targetLayer) return;

        var tempRectForStyle = (!gradientOptions.makeRect && gradientOptions.registerGraphicStyle);
        var previousActiveLayer = null;
        var tempLayer = null;

        try { previousActiveLayer = doc.activeLayer; } catch (e) { /* 無視 / ignore */ }
        var previousSelection = snapshotSelection(doc);

        if (tempRectForStyle) {
            try {
                tempLayer = doc.layers.add();
                tempLayer.name = "__TempGraphicStyle";
                doc.activeLayer = tempLayer;
            } catch (e) { tempLayer = null; }
        }

        var drawLayer = (tempRectForStyle && tempLayer) ? tempLayer : targetLayer;
        var rectSize = computeRectSize(gradientOptions, colorInput);
        var rectPosition = computeRectPosition(doc, colorInput, rectSize);

        var gradientRect = drawLayer.pathItems.rectangle(rectPosition.top, rectPosition.left, rectSize.width, rectSize.height);
        doc.selection = null;
        gradientRect.selected = true;

        /* 縦並びならグラデーション角度を 90° に / If vertical, rotate gradient by action */
        if (colorInput.selectionOrientation === "vertical") runGradientAngle90Action();

        gradientRect.stroked = false;
        gradientRect.filled = true;
        var gradientFill = new GradientColor();
        gradientFill.gradient = gradient;
        gradientRect.fillColor = gradientFill;

        if (gradientOptions.registerGraphicStyle) {
            doc.selection = null;
            gradientRect.selected = true;
            try { registerGraphicStyleFromSelected(doc); } catch (e) { /* 無視 / ignore */ }
        }

        /* 一時長方形だった場合の後始末 / Clean up the temporary rectangle */
        if (tempRectForStyle) {
            try { gradientRect.remove(); } catch (e) { /* 無視 / ignore */ }
            if (tempLayer) { try { tempLayer.remove(); } catch (e) { /* 無視 / ignore */ } }
            try { if (previousActiveLayer) doc.activeLayer = previousActiveLayer; } catch (e) { /* 無視 / ignore */ }
            try {
                doc.selection = null;
                if (previousSelection.length) doc.selection = previousSelection;
            } catch (e) { /* 削除済みのオブジェクトは選択できない / removed items cannot be selected */ }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 全体フロー: 入力 → ダイアログ → ブレンド → スウォッチ登録 → グラデーション → 長方形・スタイル
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        var colorInput = collectInputColors(doc);
        if (colorInput.colors.length < 2) return;

        var disallowSeparate = (colorInput.itemCount >= 7);
        var canBlend = !colorInput.fromSwatches && colorInput.itemCount >= 2;
        var gradientOptions = showOptionsDialog(disallowSeparate, colorInput.fromSwatches, canBlend);
        if (!gradientOptions) return;

        try {
            /* 複製したオブジェクトでブレンドを作る（元の選択は維持） / Blend the duplicates, keeping the original selection */
            if (gradientOptions.makeBlend) duplicateSelectionAndBlend(doc);

            /* 新規スウォッチグループに抽出色を登録 / Register extracted colors in a new swatch group */
            var createdSwatches = registerColorSwatches(doc, colorInput.colors, gradientOptions.makeGlobal);

            doc.selection = null;

            /* グラデーション作成 / Build the gradient */
            var gradient = null;
            if (gradientOptions.makeGradient) {
                var canSeparate = gradientOptions.separateGradient
                    && colorInput.colors.length >= 2
                    && colorInput.colors.length <= SEPARATE_MAX_COLORS;
                gradient = canSeparate
                    ? buildSeparateGradient(doc, colorInput.colors, createdSwatches)
                    : buildNormalGradient(doc, colorInput.colors, createdSwatches);
            }

            /* 長方形・グラフィックスタイル / Rectangle and Graphic Style */
            if ((gradientOptions.makeRect || gradientOptions.registerGraphicStyle) && gradient) {
                try { createGradientRect(doc, gradient, gradientOptions, colorInput); } catch (e) { /* 無視 / ignore */ }
            }

            /* 作成した最後のスウォッチ（= グラデーション）を選択 / Select the last created swatch */
            if (gradient) {
                var lastIndex = doc.swatches.length - 1;
                if (lastIndex >= 0) {
                    try { doc.swatches[lastIndex].selected = true; } catch (e) { /* 無視 / ignore */ }
                }
            }
        } catch (e) {
            /* 無言で終了 / silent */
        }
    }

    main();

})();
