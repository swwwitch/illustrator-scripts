#target illustrator
#targetengine "CharacterWidthConverterEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.CharacterWidthConverter
Name=文字の種類ごとに全角・半角を変換
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.3.0
Release-Date=2026-10-30
Description-BEGIN
カタカナ・英数字・スペース・句読点・括弧・記号を、文字の種類ごとに「そのまま／半角／全角」を選んで変換します。
設定は［横組み］［縦組み］のタブごとに持ち、開いているタブの設定で変換します。縦組み用・欧文組版用などのプリセットで設定を切り替えられます。

Converts katakana, letters, digits, spaces, punctuation, brackets and symbols between half-width and full-width, choosing "Keep", "Half" or "Full" for each kind of character.
The Horizontal and Vertical tabs each keep their own settings, and the open tab's settings are used. Presets such as vertical and Western typesetting switch the settings.
Description-END
SCRIPTMETA-END
*/

/*

### 概要

カタカナ・英数字・スペース・句読点・括弧・記号を、文字の種類ごとに「そのまま／半角／全角」を選んで変換します。
設定は［横組み］［縦組み］のタブごとに持ち、開いているタブの設定で変換します。縦組み用・欧文組版用などのプリセットで設定を切り替えられます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CharacterWidthConverter.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nf9d7245b0653

### Overview

Converts katakana, letters, digits, spaces, punctuation, brackets and symbols between half-width and full-width, choosing "Keep", "Half" or "Full" for each kind of character.
The Horizontal and Vertical tabs each keep their own settings, and the open tab's settings are used. Presets such as vertical and Western typesetting switch the settings.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CharacterWidthConverter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CharacterWidthConverter";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-30";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-11-01";                   /* 更新日 / last updated */

var SCRIPT_DESCRIPTION_JA = "文字の種類ごとに全角・半角を変換"; /* README の見出し / README title (Japanese) */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CharacterWidthConverter.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CharacterWidthConverter.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf9d7245b0653"; /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL  = "https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/CharacterWidthConverter.jsx"; /* 最新版のソース / latest source */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 選択肢の番号（ダイアログの列の順）/ Option indexes, in the order of the dialog columns */
    var OPTION_KEEP = 0;
    var OPTION_HALF = 1;
    var OPTION_FULL = 2;

    /* ［デフォルトに戻す］の値。ここに無い行は「そのまま」/ Values for "Restore Defaults"; rows not listed here are "Keep" */
    var DEFAULT_ROW_OPTIONS = {
        katakana: OPTION_FULL,
        alphabet: OPTION_HALF,
        digits: OPTION_HALF,
        kanaPunctuation: OPTION_FULL,
        middleDot: OPTION_FULL,
        kanaLeftBracket: OPTION_FULL,
        kanaRightBracket: OPTION_FULL
    };

    /* ［変換した文字に色を付ける］の色 / Color for "Color converted characters" */
    var HIGHLIGHT_RGB  = [255, 0, 0];      /* RGB のドキュメント / RGB documents */
    var HIGHLIGHT_CMYK = [0, 100, 100, 0]; /* CMYK のドキュメント / CMYK documents */

    // =========================================
    // 設定の保存 / Settings store
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

    /* 前回の設定は Illustrator を終了するまで（#targetengine が必要）、プリセットはファイルに保存 / The last settings last until Illustrator quits (needs #targetengine); presets are saved to a file */
    var sessionStore = createSettingsStore(SCRIPT_NAME, "session");
    var presetStore = createSettingsStore(SCRIPT_NAME, "persistent");

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* カラムの間隔 / gap between columns */
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

    var TABLE_ROW_SPACING     = 2;   /* 表の行どうしの間隔 / spacing between table rows */
    var TABLE_CELL_SPACING    = 4;   /* 項目名と選択肢の列の間隔 / spacing between the row label and the option columns */
    var OPTION_COLUMN_WIDTH   = 52;  /* 選択肢の列の幅 / width of each option column */
    var PRESET_DROPDOWN_WIDTH = 220; /* プリセットのドロップダウンの幅 / width of the preset dropdown */

    // =========================================
    // ラベル定義 / Labels
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
            title: { ja: "全角・半角の変換と整形", en: "Convert Width & Clean Up Text" }
        },
        tab: {
            horizontal: { ja: "横組み", en: "Horizontal" },
            vertical: { ja: "縦組み", en: "Vertical" }
        },
        panel: {
            target: { ja: "対象", en: "Scope" },
            characterType: { ja: "文字種", en: "Character Types" },
            symbol: { ja: "記号", en: "Symbols" },
            bracket: { ja: "括弧", en: "Brackets" },
            punctuation: { ja: "句読点", en: "Punctuation" },
            vertical: { ja: "縦組み用", en: "Vertical Text" },
            highlight: { ja: "確認用の色分け", en: "Check Coloring" }
        },
        radio: {
            targetSelection: { ja: "選択範囲", en: "Selection" },
            targetDocument: { ja: "ドキュメント", en: "Document" }
        },
        column: {
            keep: { ja: "そのまま", en: "Keep" },
            half: { ja: "半角", en: "Half" },
            full: { ja: "全角", en: "Full" }
        },
        row: {
            katakana: { ja: "カタカナ", en: "Katakana" },
            alphabet: { ja: "アルファベット", en: "Letters" },
            digits: { ja: "数字", en: "Digits" },
            space: { ja: "スペース", en: "Space" },
            exclamation: { ja: "感嘆符", en: "Exclamation mark" },
            doubleQuote: { ja: "二重引用符", en: "Double quote" },
            numberSign: { ja: "シャープ", en: "Number sign" },
            dollar: { ja: "ドル", en: "Dollar sign" },
            percent: { ja: "パーセント", en: "Percent sign" },
            ampersand: { ja: "アンパサンド", en: "Ampersand" },
            apostrophe: { ja: "引用符", en: "Apostrophe" },
            leftParen: { ja: "左かっこ", en: "Left parenthesis" },
            rightParen: { ja: "右かっこ", en: "Right parenthesis" },
            asterisk: { ja: "アスタリスク", en: "Asterisk" },
            plus: { ja: "プラス", en: "Plus sign" },
            comma: { ja: "コンマ", en: "Comma" },
            hyphen: { ja: "ハイフン", en: "Hyphen" },
            period: { ja: "ピリオド", en: "Period" },
            slash: { ja: "スラッシュ", en: "Slash" },
            colon: { ja: "コロン", en: "Colon" },
            semicolon: { ja: "セミコロン", en: "Semicolon" },
            lessThan: { ja: "小なり", en: "Less-than sign" },
            equals: { ja: "イコール", en: "Equals sign" },
            greaterThan: { ja: "大なり", en: "Greater-than sign" },
            question: { ja: "疑問符", en: "Question mark" },
            exclamationQuestion: { ja: "感嘆符疑問符", en: "Exclamation-question" },
            exclamationRun: { ja: "2個以上の感嘆符", en: "Two or more exclamation marks" },
            atSign: { ja: "アットマーク", en: "At sign" },
            leftBracket: { ja: "左大かっこ", en: "Left bracket" },
            backslash: { ja: "バックスラッシュ", en: "Backslash" },
            rightBracket: { ja: "右大かっこ", en: "Right bracket" },
            caret: { ja: "ハット", en: "Caret" },
            underscore: { ja: "アンダースコア", en: "Underscore" },
            backtick: { ja: "バッククォート", en: "Backtick" },
            leftBrace: { ja: "左中かっこ", en: "Left brace" },
            verticalBar: { ja: "パイプ", en: "Vertical bar" },
            rightBrace: { ja: "右中かっこ", en: "Right brace" },
            tilde: { ja: "チルダ", en: "Tilde" },
            yen: { ja: "円マーク", en: "Yen sign" },
            kanaPunctuation: { ja: "句読点", en: "Japanese punctuation" },
            middleDot: { ja: "中黒", en: "Middle dot" },
            kanaLeftBracket: { ja: "左かぎかっこ", en: "Left corner bracket" },
            kanaRightBracket: { ja: "右かぎかっこ", en: "Right corner bracket" }
        },
        fieldLabel: {
            preset: { ja: "プリセット", en: "Preset" },
            /* 「…」は ScriptUI で天地中央に出るので「...」で書く / Written as "..." since "…" sits vertically centered in ScriptUI */
            toEllipsis: { ja: "...に変換", en: "Convert to ..." },
            punctuationStyle: { ja: "句読点の種類", en: "Punctuation style" },
            quoteForm: { ja: "全角の引用符", en: "Full-width quotes" },
            hyphenForm: { ja: "全角のハイフン", en: "Full-width hyphen" },
            tildeForm: { ja: "全角のチルダ", en: "Full-width tilde" },
            asteriskForm: { ja: "全角のアスタリスク", en: "Full-width asterisk" }
        },
        checkbox: {
            digitsByLength: { ja: "数字は1桁を全角、2桁以上を半角にする", en: "Digits: one digit full-width, two or more half-width" },
            spaceBetweenJapaneseOnly: { ja: "和文にはさまれたスペースだけを変換", en: "Convert spaces only between Japanese characters" },
            japaneseLatinSpace: { ja: "和欧間のスペースを削除", en: "Remove spaces between Japanese and Western text" },
            nakaguroToEllipsis: { ja: "・・・", en: "・・・" },
            periodsToEllipsis: { ja: "...", en: "..." },
            kutenToEllipsis: { ja: "。。。", en: "。。。" },
            westernPunctuationSpace: { ja: "欧文の,.:のあとにスペースを入れる", en: "Add a space after , . : in Western text" },
            bracketInnerSpace: { ja: "括弧の内側のスペースを削除", en: "Remove spaces inside brackets" },
            bracketOuterSpace: { ja: "括弧の外側のスペースを削除", en: "Remove spaces outside brackets" },
            numberCommaToTouten: { ja: "数字の中の「,」を「、」にする", en: "\",\" in numbers to \"、\"" },
            decimalToNakaguro: { ja: "数字の中の「.」を「・」にする", en: "\".\" in numbers to \"・\"" },
            longDigitsFullwidth: { ja: "3桁以上の数字は全角にする", en: "Numbers of 3+ digits to full-width" },
            shortCapsFullwidth: { ja: "3文字以内の大文字で構成する英単語は全角にする", en: "All-capital words of up to 3 letters to full-width" },
            highlightChanges: { ja: "変換した文字に色を付ける", en: "Color converted characters" }
        },
        preset: {
            standard: { ja: "標準", en: "Standard" },
            asciiHalf: { ja: "英数字・記号を半角", en: "Letters, digits & symbols to half-width" },
            asciiFull: { ja: "英数字・記号を全角", en: "Letters, digits & symbols to full-width" },
            katakanaOnly: { ja: "カタカナだけ全角", en: "Katakana to full-width only" },
            vertical: { ja: "縦組み用", en: "For vertical typesetting" },
            western: { ja: "欧文組版用", en: "For Western typesetting" }
        },
        dropdown: {
            customPreset: { ja: "カスタム", en: "Custom" },
            punctuationStyleKeep: { ja: "そのまま", en: "Keep" },
            punctuationStyleJapanese: { ja: "、。にそろえる", en: "Unify to 、。" },
            punctuationStyleWestern: { ja: "，．にそろえる", en: "Unify to ，．" },
            quoteFormPlain: { ja: "＂＇（全角）", en: "＂＇ (full-width)" },
            quoteFormCurly: { ja: "“ ” ‘ ’（向きあり）", en: "“ ” ‘ ’ (curly)" },
            quoteFormMinute: { ja: "〝 〟（ミニュート）", en: "〝 〟 (minutes)" },
            hyphenFormPlain: { ja: "－（全角ハイフンマイナス）", en: "－ (full-width hyphen-minus)" },
            hyphenFormHyphen: { ja: "‐（ハイフン）", en: "‐ (hyphen)" },
            hyphenFormMinus: { ja: "−（マイナス）", en: "− (minus sign)" },
            tildeFormPlain: { ja: "～（全角チルダ）", en: "～ (full-width tilde)" },
            tildeFormWave: { ja: "〜（波ダッシュ）", en: "〜 (wave dash)" },
            asteriskFormPlain: { ja: "＊（全角アスタリスク）", en: "＊ (full-width asterisk)" },
            asteriskFormReference: { ja: "※（米印）", en: "※ (reference mark)" }
        },
        button: {
            restoreDefaults: { ja: "デフォルトに戻す", en: "Restore Defaults" },
            savePreset: { ja: "保存...", en: "Save..." },
            deletePreset: { ja: "削除", en: "Delete" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            convert: { ja: "変換する", en: "Convert" }
        },
        tooltip: {
            targetSelection: { ja: "選択したテキストオブジェクト、または文字ツールで選択した文字を変換します。", en: "Converts the selected text objects, or the text selected with the Type tool." },
            targetDocument: { ja: "ドキュメント内のすべてのテキストを変換します（ロック・非表示のものを除く）。", en: "Converts all text in the document (locked or hidden text excluded)." },
            keep: { ja: "変換しません。", en: "Leaves the characters as they are." },
            half: { ja: "全角の文字を半角にします。", en: "Converts full-width characters to half-width." },
            full: { ja: "半角の文字を全角にします。", en: "Converts half-width characters to full-width." },
            digitsByLength: {
                ja: "数字が1桁なら全角、2桁以上続くなら半角にします。3.5 や 1,000 のように区切りをはさむ数字はひとつながりとして数え、英字に続く数字（A4 など）は半角にします。オンのあいだ、数字の行の選択肢は使いません。",
                en: "One digit becomes full-width, two or more become half-width. Numbers such as 3.5 or 1,000 count as one number, and digits next to letters (A4) become half-width. The digits row is ignored while this is on."
            },
            spaceBetweenJapaneseOnly: {
                ja: "スペースの行で変換するのを、前と後ろがどちらも和文の文字（かな・漢字・全角文字）のスペースだけにします。",
                en: "The space row converts only spaces with Japanese characters (kana, kanji, full-width characters) on both sides."
            },
            punctuationStyle: { ja: "，．（全角コンマ・ピリオド）と、。のどちらかにそろえます。", en: "Unifies Japanese punctuation to either 、。 or ，．." },
            japaneseLatinSpace: {
                ja: "和文の文字（かな・漢字・全角文字）と英数字のあいだのスペース（半角・全角）を削除します（第 3 章 → 第3章）。",
                en: "Removes spaces (half- or full-width) between Japanese characters (kana, kanji, full-width characters) and Latin letters or digits (第 3 章 → 第3章)."
            },
            westernPunctuationSpace: {
                ja: "英数字と英字にはさまれた半角の「,」、英字にはさまれた「:」、小文字と大文字にはさまれた「.」のあとに半角スペースを入れます（apples,oranges → apples, oranges、end.Next → end. Next）。1,000 や example.com、10:30 は変えません。",
                en: "Adds a space after a half-width \",\" between a letter or digit and a letter, a \":\" between letters, and a \".\" between a lowercase and an uppercase letter (apples,oranges → apples, oranges; end.Next → end. Next). 1,000, example.com and 10:30 are left alone."
            },
            nakaguroToEllipsis: {
                ja: "3つ続く中黒（・ ･）を三点リーダー（…）にします。6つなら……になります。",
                en: "Turns three middle dots in a row (・ or ･) into an ellipsis (…); six become ……."
            },
            periodsToEllipsis: {
                ja: "3つ続くピリオド（. ．）を三点リーダー（…）にします。6つなら……になります。",
                en: "Turns three periods in a row (. or ．) into an ellipsis (…); six become ……."
            },
            kutenToEllipsis: {
                ja: "3つ続く句点（。 ｡）を三点リーダー（…）にします。6つなら……になります。",
                en: "Turns three Japanese full stops in a row (。 or ｡) into an ellipsis (…); six become ……."
            },
            bracketInnerSpace: {
                ja: "開き括弧の直後と閉じ括弧の直前のスペース（半角・全角）を削除します。向きの無い \" ' と < > は対象外です。",
                en: "Removes spaces (half- or full-width) right after opening brackets and right before closing brackets. Straight quotes and < > are not affected."
            },
            bracketOuterSpace: {
                ja: "全角の括弧・かぎかっこの直前と直後のスペースを削除します。半角の括弧の外側のスペースは残します。",
                en: "Removes spaces right before and after full-width brackets. Spaces outside half-width brackets are kept."
            },
            quoteForm: {
                ja: "二重引用符・引用符を全角にしたときの字形です。“ ” ‘ ’ と 〝 〟 は段落の中の順番で開き・閉じを決め、英数字にはさまれた ' はアポストロフィ（’）にします。〝 〟（縦組み用）のときの引用符は ＇ にします。半角にするときは、選んだ字形だけを半角に戻します。",
                en: "The form used when quotes become full-width. Curly quotes and minutes (〝 〟) open and close by their order in the paragraph; a ' between letters becomes an apostrophe (’). With minutes (for vertical text), single quotes become ＇. Converting to half-width only reverts the chosen form."
            },
            numberCommaToTouten: {
                ja: "数字にはさまれ、後ろがちょうど3桁の「,」「，」を「、」にします（1,000 → 1、000）。",
                en: "Turns \",\" and \"，\" between digits, followed by exactly three digits, into \"、\" (1,000 → 1、000)."
            },
            decimalToNakaguro: {
                ja: "数字にはさまれた「.」「．」を「・」にします（12.5 → 12・5）。",
                en: "Turns \".\" and \"．\" between digits into \"・\" (12.5 → 12・5)."
            },
            longDigitsFullwidth: {
                ja: "3桁以上の数字を全角にします（2026 → ２０２６）。3.5 や 1,000 のように区切りをはさむ数字はひとつながりとして数え、英字に続く数字は変えません。1〜2桁の数字は数字の行（または［数字は1桁を全角、2桁以上を半角にする］）に従います。",
                en: "Makes numbers of three or more digits full-width (2026 → ２０２６). Numbers such as 3.5 or 1,000 count as one number, and digits next to letters are left alone. Numbers of one or two digits follow the digits row (or \"Digits: one digit full-width…\")."
            },
            shortCapsFullwidth: {
                ja: "大文字だけでできた3文字以内の英単語を全角にします（PDF → ＰＤＦ、A → Ａ）。数字や記号を含む語（A4、U.S.）と4文字以上の語は変えません。",
                en: "Makes words of up to three capital letters full-width (PDF → ＰＤＦ, A → Ａ). Words with digits or marks (A4, U.S.) and longer words are left alone."
            },
            highlightChanges: {
                ja: "変換した文字の塗りを赤にします（仕上がりの確認用。元の色には戻りません）。削除したスペースには付きません。",
                en: "Fills converted characters with red (for checking; the original colors are not restored). Removed spaces are not marked."
            },
            hyphenForm: { ja: "ハイフンを全角にしたときの字形です。半角にするときは、この字形だけを半角に戻します。", en: "The form used when hyphens become full-width. Converting to half-width only reverts this form." },
            tildeForm: { ja: "チルダを全角にしたときの字形です。半角にするときは、この字形だけを半角に戻します。", en: "The form used when tildes become full-width. Converting to half-width only reverts this form." },
            asteriskForm: { ja: "アスタリスクを全角にしたときの字形です。※（米印）を選ぶと、* と ※ を変換します。", en: "The form used when asterisks become full-width. Choose ※ (reference mark) to convert between * and ※." },
            optionClick: { ja: "Option キーを押しながらクリックすると、このパネルの行をすべて同じ選択肢にします。", en: "Option-click to set every row in this panel to the same option." },
            restoreDefaults: { ja: "選択肢を最初の状態に戻します。", en: "Resets the options to their defaults." },
            preset: { ja: "選んだプリセットの設定を各行に反映します。今の設定に合うプリセットが無いときは「カスタム」になります。", en: "Applies the chosen preset to the rows. Shows \"Custom\" when no preset matches the current settings." },
            savePreset: { ja: "今の設定に名前を付けてプリセットとして保存します。", en: "Saves the current settings as a named preset." },
            deletePreset: { ja: "選んでいるプリセットを削除します（組み込みのプリセットは削除できません）。", en: "Deletes the chosen preset (built-in presets cannot be deleted)." }
        },
        prompt: {
            presetName: { ja: "プリセット名を入力してください。", en: "Enter a name for the preset." }
        },
        confirm: {
            overwritePreset: { ja: "プリセット「{name}」を上書きしますか？", en: "Overwrite the preset \"{name}\"?" },
            deletePreset: { ja: "プリセット「{name}」を削除しますか？", en: "Delete the preset \"{name}\"?" }
        },
        alert: {
            builtInPresetName: { ja: "「{name}」は組み込みのプリセットと同じ名前のため使えません。", en: "\"{name}\" is the name of a built-in preset and cannot be used." },
            presetSaveFailed: { ja: "プリセットを保存できませんでした。", en: "Could not save the presets." },
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            allKeep: { ja: "すべて「そのまま」のため、変換する文字がありません。", en: "Every row is set to \"Keep\", so there is nothing to convert." },
            notFound: { ja: "変換する文字は見つかりませんでした。", en: "No characters to convert were found." },
            converted: { ja: "{count} か所を変換しました。", en: "Converted {count} places." },
            breakdownLine: { ja: "{label} : {count}", en: "{label}: {count}" },
            failed: { ja: "{count} か所は変換できませんでした（ロックされたテキストなど）。", en: "{count} places could not be converted (locked text, etc.)." }
        },
        resultLabel: {
            punctuationStyle: { ja: "句読点の種類", en: "Punctuation style" },
            westernPunctuationSpace: { ja: "欧文の,.:のあとのスペース", en: "Spaces after , . : in Western text" },
            japaneseLatinSpace: { ja: "和欧間のスペース", en: "Spaces between Japanese and Western text" },
            nakaguroToEllipsis: { ja: "・・・ → …", en: "・・・ → …" },
            periodsToEllipsis: { ja: "... → …", en: "... → …" },
            kutenToEllipsis: { ja: "。。。 → …", en: "。。。 → …" },
            bracketInnerSpace: { ja: "括弧の内側のスペース", en: "Spaces inside brackets" },
            bracketOuterSpace: { ja: "括弧の外側のスペース", en: "Spaces outside brackets" },
            numberCommaToTouten: { ja: "数字の中の「,」", en: "\",\" in numbers" },
            decimalToNakaguro: { ja: "数字の中の「.」", en: "\".\" in numbers" },
            longDigitsFullwidth: { ja: "3桁以上の数字", en: "Numbers of 3+ digits" },
            shortCapsFullwidth: { ja: "3文字以内の大文字の英単語", en: "Capital words of up to 3 letters" }
        }
    };

    // =========================================
    // 変換表 / Conversion table
    // =========================================

    /* 全角と半角のコード差（ASCII の ! 〜 ~ と全角英数・記号）/ Code offset between ASCII ! to ~ and their full-width forms */
    var FULL_WIDTH_OFFSET = 0xFEE0;

    /* 半角カタカナと全角の対応（同じ位置どうし）/ Half-width katakana and their full-width forms, matched by position */
    var HALF_KANA = "ｦｧｨｩｪｫｬｭｮｯｰｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜﾝﾞﾟ";
    var FULL_KANA = "ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン゛゜";

    /* 濁音・半濁音は半角では2文字になる / Voiced and semi-voiced kana take two characters in half-width */
    var HALF_VOICED_BASES      = "ｳｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾊﾋﾌﾍﾎﾜｦ";
    var FULL_VOICED_KANA       = "ヴガギグゲゴザジズゼゾダヂヅデドバビブベボヷヺ";
    var HALF_SEMI_VOICED_BASES = "ﾊﾋﾌﾍﾎ";
    var FULL_SEMI_VOICED_KANA  = "パピプペポ";

    /* 記号の行（キー・半角の文字[・全角の文字]）/ Symbol rows (key, half-width character[, full-width character]) */
    var SYMBOL_ROW_DEFS = [
        ["hyphen", "-"], ["underscore", "_"], ["ampersand", "&"], ["numberSign", "#"],
        ["dollar", "$"], ["yen", "¥", "￥"], ["percent", "%"], ["atSign", "@"],
        ["asterisk", "*"], ["equals", "="], ["plus", "+"], ["slash", "/"], ["backslash", "\\"],
        ["verticalBar", "|"], ["caret", "^"], ["backtick", "`"], ["tilde", "~"]
    ];

    /* 句読点の行（ASCII の句読点。和文の句読点は PUNCTUATION_ROWS で足す）/ Punctuation rows (ASCII punctuation; the Japanese punctuation is added in PUNCTUATION_ROWS) */
    var ASCII_PUNCTUATION_ROW_DEFS = [
        ["comma", ","], ["period", "."], ["colon", ":"],
        ["semicolon", ";"], ["exclamation", "!"], ["question", "?"]
    ];

    /* 括弧の行（ASCII の引用符・括弧・不等号。かぎかっこは BRACKET_ROWS で足す）/ Bracket rows (ASCII quotes, brackets and less/greater-than signs; the corner brackets are added in BRACKET_ROWS) */
    var ASCII_BRACKET_ROW_DEFS = [
        ["doubleQuote", "\""], ["apostrophe", "'"],
        ["leftParen", "("], ["rightParen", ")"], ["leftBracket", "["],
        ["rightBracket", "]"], ["leftBrace", "{"], ["rightBrace", "}"],
        ["lessThan", "<"], ["greaterThan", ">"]
    ];

    /**
     * 1文字の半角とその全角（コード差 FULL_WIDTH_OFFSET）の組を返す
     * @param {number} halfCode - 半角の文字コード
     * @returns {string[]} [半角, 全角]
     */
    function offsetPair(halfCode) {
        return [String.fromCharCode(halfCode), String.fromCharCode(halfCode + FULL_WIDTH_OFFSET)];
    }

    /**
     * 文字コードの範囲の半角・全角の組を並べる
     * @param {number} firstCode - 最初の半角の文字コード
     * @param {number} lastCode - 最後の半角の文字コード
     * @returns {Array<string[]>} [半角, 全角] の配列
     */
    function offsetPairsInRange(firstCode, lastCode) {
        var pairs = [];
        for (var code = firstCode; code <= lastCode; code++) pairs.push(offsetPair(code));
        return pairs;
    }

    /**
     * 同じ長さの2つの文字列から、同じ位置の文字どうしの組を作る
     * @param {string} halfChars - 半角側の文字列
     * @param {string} fullChars - 全角側の文字列
     * @param {string} [halfSuffix] - 半角側の各文字に付ける文字（濁点・半濁点）
     * @returns {Array<string[]>} [半角, 全角] の配列
     */
    function zipPairs(halfChars, fullChars, halfSuffix) {
        var pairs = [];
        for (var i = 0; i < halfChars.length; i++) {
            pairs.push([halfChars.charAt(i) + (halfSuffix || ""), fullChars.charAt(i)]);
        }
        return pairs;
    }

    /**
     * カタカナの半角・全角の組を返す（濁音・半濁音を含む）
     * @returns {Array<string[]>} [半角, 全角] の配列
     */
    function buildKatakanaPairs() {
        return zipPairs(HALF_KANA, FULL_KANA)
            .concat(zipPairs(HALF_VOICED_BASES, FULL_VOICED_KANA, "ﾞ"))
            .concat(zipPairs(HALF_SEMI_VOICED_BASES, FULL_SEMI_VOICED_KANA, "ﾟ"));
    }

    /**
     * 行の定義を作る
     * @param {string} key - 行のキー（LABELS.row のキー、設定の保存にも使う）
     * @param {string} glyph - 項目名の後ろに添える文字（無ければ空文字）
     * @param {Array<string[]>} pairs - [半角, 全角] の配列
     * @param {Object} [rowFlags] - { keepsUnit: true } なら「そのまま」でも組をひとまとまりとして拾い、1文字ずつの行の変換から守る。
     *     { runChar: [半角, 全角] } なら、その文字が2個以上続く並びを1組として扱う（pairs は使わない）
     * @returns {{key: string, glyph: string, pairs: Array<string[]>, keepsUnit: boolean, runChar: string[]|null}} 行の定義
     */
    function rowDef(key, glyph, pairs, rowFlags) {
        return {
            key: key,
            glyph: glyph,
            pairs: pairs,
            keepsUnit: !!(rowFlags && rowFlags.keepsUnit),
            runChar: (rowFlags && rowFlags.runChar) || null
        };
    }

    var CHARACTER_TYPE_ROWS = [
        rowDef("katakana", "", buildKatakanaPairs()),
        rowDef("alphabet", "", offsetPairsInRange(0x41, 0x5A).concat(offsetPairsInRange(0x61, 0x7A))),
        rowDef("digits", "", offsetPairsInRange(0x30, 0x39)),
        rowDef("space", "", [[" ", "　"]])
    ];

    /**
     * ASCII の記号1文字ずつの行を作る（全角は3つ目の文字、無ければコード差 FULL_WIDTH_OFFSET）
     * @param {Array<string[]>} symbolRowDefs - [キー, 半角の文字, 全角の文字（省略可）] の配列
     * @returns {Array} 行の定義の配列
     */
    function buildSymbolRows(symbolRowDefs) {
        var symbolRows = [];
        for (var i = 0; i < symbolRowDefs.length; i++) {
            var symbolChar = symbolRowDefs[i][1];
            var symbolPair = symbolRowDefs[i][2] ? [symbolChar, symbolRowDefs[i][2]] : offsetPair(symbolChar.charCodeAt(0));
            symbolRows.push(rowDef(symbolRowDefs[i][0], symbolChar, [symbolPair]));
        }
        return symbolRows;
    }

    /* 中黒は記号パネルに置くが、プリセットでは和文の約物として扱う / The middle dot sits in the symbol panel but presets treat it as Japanese punctuation */
    var ASCII_SYMBOL_ROWS = buildSymbolRows(SYMBOL_ROW_DEFS);
    var SYMBOL_ROWS = ASCII_SYMBOL_ROWS.concat([
        rowDef("middleDot", "･", [["･", "・"]])
    ]);

    var ASCII_BRACKET_ROWS = buildSymbolRows(ASCII_BRACKET_ROW_DEFS);
    var BRACKET_ROWS = ASCII_BRACKET_ROWS.concat([
        rowDef("kanaLeftBracket", "｢", [["｢", "「"]]),
        rowDef("kanaRightBracket", "｣", [["｣", "」"]])
    ]);

    /* !? ?! と2個以上続く ! は1組。「そのまま」なら、感嘆符・疑問符の行の設定に関係なく元のまま残す / "!?", "?!" and runs of "!" are handled as units; when kept, they stay as they are whatever the exclamation and question rows say */
    var ASCII_PUNCTUATION_ROWS = buildSymbolRows(ASCII_PUNCTUATION_ROW_DEFS).concat([
        rowDef("exclamationQuestion", "!?", [["!?", "！？"], ["?!", "？！"]], { keepsUnit: true }),
        rowDef("exclamationRun", "!!", [], { runChar: ["!", "！"] })
    ]);
    var PUNCTUATION_ROWS = [
        rowDef("kanaPunctuation", "｡､", [["｡", "。"], ["､", "、"]])
    ].concat(ASCII_PUNCTUATION_ROWS);

    var ALL_ROWS = CHARACTER_TYPE_ROWS.concat(PUNCTUATION_ROWS).concat(SYMBOL_ROWS).concat(BRACKET_ROWS);

    /* 全角にしたときの字形を選べる行（行のキー → 設定のキーと字形の候補）。2文字の字形は、前後で決まる開き・閉じ / Rows whose full-width form can be chosen (row key → setting key and candidate forms); a two-character form is an opening/closing pair chosen by context */
    var FULL_FORM_CHOICES = {
        hyphen: { settingKey: "hyphenForm", forms: [["\uFF0D"], ["\u2010"], ["\u2212"]] },
        tilde: { settingKey: "tildeForm", forms: [["\uFF5E"], ["\u301C"]] },
        asterisk: { settingKey: "asteriskForm", forms: [["\uFF0A"], ["\u203B"]] },
        doubleQuote: { settingKey: "quoteForm", forms: [["\uFF02"], ["\u201C", "\u201D"], ["\u301D", "\u301F"]] },
        apostrophe: { settingKey: "quoteForm", forms: [["\uFF07"], ["\u2018", "\u2019"], ["\uFF07"]] }  /* ミニュートに一重は無いので ＇ / minutes have no single form, so ＇ */
    };

    /* 行のほかの設定（キーと選べる値の数。値は 0 から数え、0 が既定）。プリセットにも入る。
       tab があるものはそのタブにだけ置き、ほかのタブでは 0 のまま /
       Settings besides the rows (key and number of values, counted from 0, with 0 as the default); saved in presets too.
       Those with a tab appear only on that tab and stay 0 on the other */
    var EXTRA_SETTING_DEFS = [
        { key: "digitsByLength", valueCount: 2 },
        { key: "spaceBetweenJapaneseOnly", valueCount: 2 },
        { key: "japaneseLatinSpace", valueCount: 2 },
        { key: "punctuationStyle", valueCount: 3 },
        { key: "nakaguroToEllipsis", valueCount: 2 },
        { key: "periodsToEllipsis", valueCount: 2 },
        { key: "kutenToEllipsis", valueCount: 2 },
        { key: "westernPunctuationSpace", valueCount: 2, tab: "horizontal" },
        { key: "bracketInnerSpace", valueCount: 2 },
        { key: "bracketOuterSpace", valueCount: 2 },
        { key: "quoteForm", valueCount: 3 },
        { key: "hyphenForm", valueCount: 3 },
        { key: "tildeForm", valueCount: 2 },
        { key: "asteriskForm", valueCount: 2 },
        { key: "numberCommaToTouten", valueCount: 2, tab: "vertical" },
        { key: "decimalToNakaguro", valueCount: 2, tab: "vertical" },
        { key: "longDigitsFullwidth", valueCount: 2, tab: "vertical" },
        { key: "shortCapsFullwidth", valueCount: 2, tab: "vertical" }
    ];

    /* 設定を別々に持つタブ（組み方向）の並び / Tabs (text directions), each with its own settings */
    var TAB_KEYS = ["horizontal", "vertical"];

    /* 行と行のほかの設定をまとめた一覧と、キーごとの値の数 / Every setting (rows and extras) and the number of values per key */
    var SETTING_DEFS = [];
    var SETTING_VALUE_COUNTS = {};
    for (var rowIndex = 0; rowIndex < ALL_ROWS.length; rowIndex++) SETTING_DEFS.push({ key: ALL_ROWS[rowIndex].key, valueCount: 3 });
    SETTING_DEFS = SETTING_DEFS.concat(EXTRA_SETTING_DEFS);
    for (var settingIndex = 0; settingIndex < SETTING_DEFS.length; settingIndex++) {
        SETTING_VALUE_COUNTS[SETTING_DEFS[settingIndex].key] = SETTING_DEFS[settingIndex].valueCount;
    }

    /* 句読点の種類（punctuationStyle の値 → 置き換える組）/ Punctuation style (punctuationStyle value → replacements) */
    var PUNCTUATION_STYLE_MAPS = [
        null,
        { "\uFF0C": "\u3001", "\uFF0E": "\u3002" },  /* ，． → 、。 */
        { "\u3001": "\uFF0C", "\u3002": "\uFF0E" }   /* 、。 → ，． */
    ];

    /* 3つ続くと三点リーダーにする文字（設定のキー → 半角・全角の文字）/ Characters that become an ellipsis when three run together (setting key → half- and full-width characters) */
    var ELLIPSIS_SOURCES = [
        { key: "nakaguroToEllipsis", chars: "\u30FB\uFF65" },  /* ・ ･ */
        { key: "periodsToEllipsis", chars: ".\uFF0E" },         /* . ． */
        { key: "kutenToEllipsis", chars: "\u3002\uFF61" }      /* 。 ｡ */
    ];

    /* 括弧の内側のスペースを削除する括弧（向きの無い " ' と不等号、アポストロフィと紛れる ‘ ’ は含めない）/ Brackets for removing inner spaces (straight quotes, < >, and ‘ ’ which look like apostrophes are excluded) */
    var OPENING_BRACKETS = "([{\uFF08\uFF3B\uFF5B\uFF62\u300C\u300E\u3010\u3014\u3008\u300A\u201C\u301D";
    var CLOSING_BRACKETS = ")]}\uFF09\uFF3D\uFF5D\uFF63\u300D\u300F\u3011\u3015\u3009\u300B\u201D\u301F";
    /* 括弧の外側のスペースを削除する括弧（全角だけ。欧文の「word (note)」の空きは残す）/ Brackets for removing outer spaces (full-width only, so "word (note)" keeps its space) */
    var FULL_WIDTH_OPENING_BRACKETS = "\uFF08\uFF3B\uFF5B\uFF62\u300C\u300E\u3010\u3014\u3008\u300A\u301D";
    var FULL_WIDTH_CLOSING_BRACKETS = "\uFF09\uFF3D\uFF5D\uFF63\u300D\u300F\u3011\u3015\u3009\u300B\u301F";

    /* 和文の文字（スペースの前後の判定に使う。全角スペースは含めない）/ Japanese characters, for the space check (the ideographic space itself is excluded) */
    var JAPANESE_CHAR_PATTERN = /[\u3001-\u303F\u3040-\u30FF\u31F0-\u31FF\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF\uFF01-\uFF60\uFF61-\uFF9F\uFFE0-\uFFEF]/;
    /* 欧文の英字（1桁の数字に続く文字の判定に使う）/ Latin letters, for digits next to letters */
    var LATIN_LETTER_PATTERN = /[A-Za-z\u00C0-\u024F\uFF21-\uFF3A\uFF41-\uFF5A]/;
    /* 欧文の英数字（アポストロフィの判定に使う）/ Latin letters and digits, for apostrophes */
    var LATIN_ALNUM_PATTERN = /[A-Za-z0-9\u00C0-\u024F]/;
    /* 英単語をつくる文字（半角・全角の英数字）と、前後が英数字のときだけ語の一部になるつなぎ（U.S. や AT&T）/ Characters forming a Latin word, and joiners that belong to the word only between two of them (U.S., AT&T) */
    var WORD_CHAR_PATTERN = /[A-Za-z0-9\u00C0-\u024F\uFF10-\uFF19\uFF21-\uFF3A\uFF41-\uFF5A]/;
    var WORD_JOINERS = "'\u2019-.&";
    /* ［3文字以内の大文字で構成する英単語は全角にする］の対象（半角・全角の大文字だけで1〜3文字）/ Words for "all-capital words of up to 3 letters" (one to three capitals, half- or full-width) */
    var SHORT_CAPS_WORD_PATTERN = /^[A-Z\uFF21-\uFF3A]{1,3}$/;
    /* ［3桁以上の数字は全角にする］の桁数 / Digit count for "numbers of 3+ digits to full-width" */
    var LONG_DIGIT_COUNT = 3;

    /* 数字と、数字のあいだにはさまる区切り（3.5 や 1,000 をひとつながりにする）/ Digits, and the separators kept inside a number such as 3.5 or 1,000 */
    var DIGIT_PATTERN = /[0-9\uFF10-\uFF19]/;
    var DIGIT_SEPARATORS = ",.\uFF0C\uFF0E";
    /* 括弧のまわりで削除するスペース（半角・全角）/ Spaces removed around brackets (half- and full-width) */
    var BRACKET_SPACES = " \u3000";
    /* スペースを入れる欧文の「,」「.」「:」の前後（「,」は英数字と英字、「.」は小文字と大文字、「:」は英字どうし）/ Neighbors of Western ",", "." and ":" that get a space ("," letter or digit then letter, "." lowercase then uppercase, ":" letters) */
    var WESTERN_PUNCTUATION_SPACE_RULES = {
        ",": { previous: /[A-Za-z0-9]/, next: /[A-Za-z]/ },
        ".": { previous: /[a-z]/, next: /[A-Z]/ },
        ":": { previous: /[A-Za-z]/, next: /[A-Za-z]/ }
    };

    // =========================================
    // 前提チェック / Preconditions
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var activeDocument = app.activeDocument;

    // =========================================
    // 設定 / Settings
    // =========================================

    /**
     * 各行と行のほかの設定の既定値を返す
     * @returns {Object} 設定のキー → 値
     */
    function getDefaultRowOptions() {
        var rowOptions = {};
        for (var i = 0; i < SETTING_DEFS.length; i++) {
            var settingKey = SETTING_DEFS[i].key;
            rowOptions[settingKey] = DEFAULT_ROW_OPTIONS.hasOwnProperty(settingKey) ? DEFAULT_ROW_OPTIONS[settingKey] : 0;
        }
        return rowOptions;
    }

    /**
     * すべての行を同じ選択肢にし（行のほかの設定は既定値）、一部だけ別の値にした設定を返す
     * @param {number} baseOption - すべての行の選択肢
     * @param {Object} [overrides] - 設定のキー → 値（baseOption より優先）
     * @returns {Object} 設定のキー → 値
     */
    function rowOptionsWith(baseOption, overrides) {
        var rowOptions = {};
        for (var i = 0; i < EXTRA_SETTING_DEFS.length; i++) rowOptions[EXTRA_SETTING_DEFS[i].key] = 0;
        for (var j = 0; j < ALL_ROWS.length; j++) rowOptions[ALL_ROWS[j].key] = baseOption;
        return copyRowOptions(overrides, rowOptions);
    }

    /**
     * 保存してあった設定のうち、知っているキーの正しい値だけを写す
     * @param {Object} sourceOptions - 写す元（読み込んだ値）
     * @param {Object} rowOptions - 写す先（すべてのキーの値が入ったもの）
     * @returns {Object} 写す先と同じオブジェクト
     */
    function copyRowOptions(sourceOptions, rowOptions) {
        if (!sourceOptions) return rowOptions;
        for (var settingKey in rowOptions) {
            if (!rowOptions.hasOwnProperty(settingKey)) continue;
            var value = sourceOptions[settingKey];
            if (typeof value === "number" && value % 1 === 0 && value >= 0 && value < SETTING_VALUE_COUNTS[settingKey]) {
                rowOptions[settingKey] = value;
            }
        }
        return rowOptions;
    }

    /**
     * 2つの設定がすべてのキーで同じか
     * @param {Object} firstOptions - 設定のキー → 値
     * @param {Object} secondOptions - 設定のキー → 値
     * @returns {boolean} 同じなら true
     */
    function isSameRowOptions(firstOptions, secondOptions) {
        for (var i = 0; i < SETTING_DEFS.length; i++) {
            var settingKey = SETTING_DEFS[i].key;
            if (firstOptions[settingKey] !== secondOptions[settingKey]) return false;
        }
        return true;
    }

    /**
     * そのタブに置かない設定を 0 にした写しを返す
     * @param {Object} rowOptions - 設定のキー → 値（すべてのキーの値が入ったもの）
     * @param {string} tabKey - "horizontal" または "vertical"
     * @returns {Object} 設定のキー → 値（新しいオブジェクト）
     */
    function rowOptionsForTab(rowOptions, tabKey) {
        var tabOptions = {};
        for (var i = 0; i < SETTING_DEFS.length; i++) {
            var settingDef = SETTING_DEFS[i];
            tabOptions[settingDef.key] = (settingDef.tab && settingDef.tab !== tabKey) ? 0 : rowOptions[settingDef.key];
        }
        return tabOptions;
    }

    /**
     * 前回の設定を読む（無ければ既定値）
     * @returns {{rowOptionsByTab: Object, activeTab: string, scope: string|null, highlightChanges: boolean}} タブごとの設定・開いていたタブ・対象・色を付けるか
     */
    function loadSettings() {
        var saved = sessionStore.load({ rowOptionsByTab: { horizontal: {}, vertical: {} }, activeTab: "horizontal", scope: "", highlightChanges: false });
        var rowOptionsByTab = {};
        for (var i = 0; i < TAB_KEYS.length; i++) {
            var tabKey = TAB_KEYS[i];
            rowOptionsByTab[tabKey] = rowOptionsForTab(copyRowOptions(saved.rowOptionsByTab[tabKey], getDefaultRowOptions()), tabKey);
        }
        return {
            rowOptionsByTab: rowOptionsByTab,
            activeTab: (saved.activeTab === "vertical") ? "vertical" : "horizontal",
            scope: saved.scope || null,
            highlightChanges: saved.highlightChanges
        };
    }

    /**
     * 設定を覚える（Illustrator を終了するまで）
     * @param {{rowOptionsByTab: Object, activeTab: string, scope: string, highlightChanges: boolean}} settings - タブごとの設定・開いているタブ・対象・色を付けるか
     * @returns {void}
     */
    function saveSettings(settings) {
        sessionStore.save({
            rowOptionsByTab: settings.rowOptionsByTab,
            activeTab: settings.activeTab,
            scope: settings.scope,
            highlightChanges: settings.highlightChanges
        });
    }

    // =========================================
    // プリセット / Presets
    // =========================================

    /* 組み込みのプリセット（削除・上書きはできない）。tab があるものは、選ぶとそのタブに切り替えて反映する / Built-in presets (cannot be deleted or overwritten); choosing one with a tab switches to that tab */
    var BUILTIN_PRESETS = [
        { name: getLabel(LABELS.preset.standard), builtIn: true, rowOptions: getDefaultRowOptions() },
        {
            name: getLabel(LABELS.preset.asciiHalf), builtIn: true,
            rowOptions: rowOptionsWith(OPTION_KEEP, presetOverrides(OPTION_HALF, OPTION_HALF))
        },
        {
            name: getLabel(LABELS.preset.asciiFull), builtIn: true,
            rowOptions: rowOptionsWith(OPTION_KEEP, presetOverrides(OPTION_FULL, OPTION_FULL))
        },
        {
            name: getLabel(LABELS.preset.katakanaOnly), builtIn: true,
            rowOptions: rowOptionsWith(OPTION_KEEP, presetOverrides(OPTION_KEEP, OPTION_KEEP))
        },
        {
            /* 縦組み：数字は1桁を全角・2桁以上を半角（縦中横向け）、!? !! は半角、和文の間のスペースは全角、チルダは波ダッシュ。
               英字・コンマ・ピリオド・引用符の行はそのまま（欧文の単語を崩さない）、数字の中の「,」「.」は「、」「・」。
               和欧間のスペースは削除し、・・・ ... 。。。 は … にする /
               Vertical: one digit full-width, two or more half-width (for tate-chu-yoko); !? and !! half-width; spaces between Japanese full-width;
               tildes become wave dashes. Letter, comma, period and quote rows are kept so English words stay intact; "," and "." in numbers become 、 and ・.
               Spaces between Japanese and Western text are removed, and ・・・ ... 。。。 become … */
            name: getLabel(LABELS.preset.vertical), builtIn: true, tab: "vertical",
            rowOptions: rowOptionsWith(OPTION_FULL, {
                alphabet: OPTION_KEEP,
                digits: OPTION_KEEP,
                comma: OPTION_KEEP,
                period: OPTION_KEEP,
                doubleQuote: OPTION_KEEP,
                apostrophe: OPTION_KEEP,
                exclamationQuestion: OPTION_HALF,
                exclamationRun: OPTION_HALF,
                digitsByLength: 1,
                spaceBetweenJapaneseOnly: 1,
                tildeForm: 1,
                numberCommaToTouten: 1,
                decimalToNakaguro: 1,
                japaneseLatinSpace: 1,
                nakaguroToEllipsis: 1,
                periodsToEllipsis: 1,
                kutenToEllipsis: 1
            })
        },
        {
            /* 欧文組版：ASCII にある文字（全角スペース・円マークも）はすべて半角、和文のカタカナ・約物はそのまま、... は … にし、欧文の,.:のあとにスペース /
               Western typesetting: everything with an ASCII form (ideographic space and yen too) to half-width; Japanese kana and punctuation kept; ... becomes …; a space after , . : */
            name: getLabel(LABELS.preset.western), builtIn: true, tab: "horizontal",
            rowOptions: rowOptionsWith(OPTION_HALF, {
                katakana: OPTION_KEEP,
                kanaPunctuation: OPTION_KEEP,
                middleDot: OPTION_KEEP,
                kanaLeftBracket: OPTION_KEEP,
                kanaRightBracket: OPTION_KEEP,
                periodsToEllipsis: 1,
                westernPunctuationSpace: 1
            })
        }
    ];

    /**
     * 組み込みのプリセットの選択肢を作る（カタカナ・句読点・かぎかっこは全角、スペースはそのまま）
     * @param {number} asciiOption - アルファベット・数字の選択肢
     * @param {number} symbolOption - 記号（円マークを含む）と ASCII の句読点・括弧の行の選択肢
     * @returns {Object} 行のキー → 選択肢の番号
     */
    function presetOverrides(asciiOption, symbolOption) {
        var overrides = {
            katakana: OPTION_FULL,
            alphabet: asciiOption,
            digits: asciiOption,
            kanaPunctuation: OPTION_FULL,
            middleDot: OPTION_FULL,
            kanaLeftBracket: OPTION_FULL,
            kanaRightBracket: OPTION_FULL
        };
        var symbolRows = ASCII_SYMBOL_ROWS.concat(ASCII_PUNCTUATION_ROWS).concat(ASCII_BRACKET_ROWS);
        for (var i = 0; i < symbolRows.length; i++) overrides[symbolRows[i].key] = symbolOption;
        return overrides;
    }

    /**
     * プリセットがそのタブの設定か（tab の無いプリセットはどちらのタブの設定にもなる）。今の設定に合うプリセットを探すときに使う
     * @param {{tab: string|null}} preset - プリセット
     * @param {string} tabKey - "horizontal" または "vertical"
     * @returns {boolean} そのタブの設定なら true
     */
    function isPresetForTab(preset, tabKey) {
        return !preset.tab || preset.tab === tabKey;
    }

    /**
     * 保存してあるユーザーのプリセットを読む
     * @returns {Array<{name: string, builtIn: boolean, tab: string|null, rowOptions: Object}>} プリセットの配列
     */
    function loadUserPresets() {
        var storedPresets = presetStore.load({ presets: [] }).presets;
        var userPresets = [];
        for (var i = 0; i < storedPresets.length; i++) {
            var storedPreset = storedPresets[i];
            if (!storedPreset || typeof storedPreset.name !== "string" || storedPreset.name === "") continue;
            /* タブを持たない（v1.2.0 までに保存した）プリセットは、開いているタブに反映する / Presets without a tab (saved up to v1.2.0) apply to the open tab */
            var presetTab = (storedPreset.tab === "horizontal" || storedPreset.tab === "vertical") ? storedPreset.tab : null;
            /* 保存に無い行は「そのまま」/ Rows missing from the file are "Keep" */
            var presetOptions = copyRowOptions(storedPreset.rowOptions, rowOptionsWith(OPTION_KEEP));
            userPresets.push({
                name: storedPreset.name,
                builtIn: false,
                tab: presetTab,
                rowOptions: presetTab ? rowOptionsForTab(presetOptions, presetTab) : presetOptions
            });
        }
        return userPresets;
    }

    /**
     * ユーザーのプリセットをファイルに保存する
     * @param {Array<{name: string, tab: string|null, rowOptions: Object}>} userPresets - プリセットの配列
     * @returns {boolean} 保存できたら true
     */
    function saveUserPresets(userPresets) {
        var storedPresets = [];
        for (var i = 0; i < userPresets.length; i++) {
            var storedPreset = { name: userPresets[i].name, rowOptions: userPresets[i].rowOptions };
            if (userPresets[i].tab) storedPreset.tab = userPresets[i].tab;
            storedPresets.push(storedPreset);
        }
        return presetStore.save({ presets: storedPresets });
    }

    // =========================================
    // 対象の収集 / Target collection
    // =========================================

    /*
     * 変換の対象は { base, start, length } で持つ。base は文字を数える TextRange、start はその中の先頭位置、
     * length は文字数（-1 なら base 全体）。文字を消したり足したりしたら length を合わせる
     * A target is { base, start, length }: base is the TextRange whose characters are indexed, start is the offset in it,
     * and length is the number of characters (-1 for the whole base). length follows insertions and deletions
     */

    /**
     * テキストフレーム全体を対象にする
     * @param {TextRange} baseRange - テキストフレームまたはストーリーの textRange
     * @returns {{base: TextRange, start: number, length: number}} 対象
     */
    function wholeRangeTarget(baseRange) {
        return { base: baseRange, start: 0, length: -1 };
    }

    /**
     * 文字ツールで選択した文字を対象にする。
     * ストーリーの中の位置で持ち、位置が合わなければ選択した TextRange そのものを使う
     * @param {TextRange} selectedRange - 選択した文字
     * @returns {{base: TextRange, start: number, length: number}} 対象
     */
    function selectedRangeTarget(selectedRange) {
        try {
            var storyRange = selectedRange.story.textRange;
            /* characterOffset は1始まり / characterOffset is 1-based */
            var rangeStart = selectedRange.characterOffset - 1;
            if (String(storyRange.contents).substr(rangeStart, selectedRange.length) === selectedRange.contents) {
                return { base: storyRange, start: rangeStart, length: selectedRange.length };
            }
        } catch (e) { }
        return wholeRangeTarget(selectedRange);
    }

    /**
     * 選択（グループの中も）からテキストフレームを集める
     * @param {Array} items - 選択、またはグループ内のページアイテム
     * @param {TextFrame[]} textFrames - 集めたテキストフレームの格納先
     * @returns {void}
     */
    function collectTextFrames(items, textFrames) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "TextFrame") {
                textFrames.push(item);
            } else if (item.typename === "GroupItem") {
                collectTextFrames(item.pageItems, textFrames);
            }
        }
    }

    /**
     * 選択から変換の対象を集める（文字を選択中ならその文字、オブジェクトを選択中なら各テキストフレーム全体）
     * @returns {Array} 対象の配列
     */
    function getSelectionTargets() {
        var currentSelection = activeDocument.selection;
        var targets = [];
        if (!currentSelection) return targets;
        /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
        if (currentSelection.typename === "TextRange") {
            if (currentSelection.length > 0) targets.push(selectedRangeTarget(currentSelection));
            return targets;
        }
        var textFrames = [];
        collectTextFrames(currentSelection, textFrames);
        for (var i = 0; i < textFrames.length; i++) targets.push(wholeRangeTarget(textFrames[i].textRange));
        return targets;
    }

    /**
     * テキストフレームが書き換えられるか（本体・親のグループ・レイヤーがロック・非表示でない）
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {boolean} 書き換えられるなら true
     */
    function isEditableTextFrame(textFrame) {
        try {
            if (textFrame.locked || textFrame.hidden) return false;
            var parentItem = textFrame.parent;
            while (parentItem && parentItem.typename !== "Document") {
                if (parentItem.typename === "Layer") {
                    if (parentItem.locked || !parentItem.visible) return false;
                } else if (parentItem.locked || parentItem.hidden) {
                    return false;
                }
                parentItem = parentItem.parent;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * ドキュメントのすべてのストーリーを対象にする（ロック・非表示のフレームを含むストーリーは除く）。
     * 連結したテキストはストーリー単位で1回だけ処理する
     * @returns {Array} 対象の配列
     */
    function getDocumentTargets() {
        var targets = [];
        var stories = activeDocument.stories;
        for (var i = 0; i < stories.length; i++) {
            var storyFrames = stories[i].textFrames;
            var isEditable = storyFrames.length > 0;
            for (var j = 0; j < storyFrames.length && isEditable; j++) {
                if (!isEditableTextFrame(storyFrames[j])) isEditable = false;
            }
            if (isEditable) targets.push(wholeRangeTarget(stories[i].textRange));
        }
        return targets;
    }

    // =========================================
    // 変換 / Conversion
    // =========================================

    /**
     * 行の [半角, 全角の字形の配列] の組を返す（字形を選べる行は、選んだ字形にする）
     * @param {Object} row - 行の定義
     * @param {Object} rowOptions - 設定のキー → 値
     * @returns {Array} [半角, 全角の字形の配列] の配列。字形が2つなら開き・閉じ
     */
    function getRowPairs(row, rowOptions) {
        var formChoice = FULL_FORM_CHOICES[row.key];
        if (formChoice) return [[row.pairs[0][0], formChoice.forms[rowOptions[formChoice.settingKey] || 0]]];
        var rowPairs = [];
        for (var i = 0; i < row.pairs.length; i++) rowPairs.push([row.pairs[i][0], [row.pairs[i][1]]]);
        return rowPairs;
    }

    /**
     * 設定から、変換元の文字 → 変換後の文字の表と、ほかの処理の有無をまとめる
     * @param {Object} rowOptions - 設定のキー → 値
     * @returns {{map: Object, rowKeyBySource: Object, runUnits: Array, size: number, digitsByLength: boolean,
     *     spaceBetweenJapaneseOnly: boolean, japaneseLatinSpace: boolean, punctuationStyle: number, nakaguroToEllipsis: boolean, periodsToEllipsis: boolean, kutenToEllipsis: boolean, westernPunctuationSpace: boolean,
     *     bracketInnerSpace: boolean, bracketOuterSpace: boolean,
     *     numberCommaToTouten: boolean, decimalToNakaguro: boolean, longDigitsFullwidth: boolean, shortCapsFullwidth: boolean}}
     *     変換表（値は文字列、向きで決まる引用符は [開き, 閉じ]）・変換元の行のキー・連続する文字の組・実際に変換する項目の数（守るだけの組は数えない）とほかの処理
     */
    function buildConversionMap(rowOptions) {
        var conversionMap = {};
        var rowKeyBySource = {};
        var runUnits = [];
        var mapSize = 0;
        var digitsByLength = (rowOptions.digitsByLength === 1);
        for (var i = 0; i < ALL_ROWS.length; i++) {
            var row = ALL_ROWS[i];
            var option = rowOptions[row.key];
            var isConverting = (option === OPTION_HALF || option === OPTION_FULL);
            /* 桁数で決めるときは、数字の行の選択肢は使わない / The digits row is ignored when digits go by length */
            if (row.key === "digits" && digitsByLength) continue;
            if (row.runChar) {
                /* 「そのまま」でも拾って、1文字ずつの行の変換から守る / Picked up even when kept, so the single-character rows leave it alone */
                runUnits.push({ key: row.key, halfChar: row.runChar[0], fullChar: row.runChar[1], option: option });
                if (isConverting) mapSize++;
                continue;
            }
            var rowPairs = getRowPairs(row, rowOptions);
            for (var j = 0; j < rowPairs.length; j++) {
                var halfText = rowPairs[j][0];
                var fullForms = rowPairs[j][1];
                if (row.keepsUnit) {
                    /* ひとまとまりの組は、まず元と同じ文字に置き換える項目を入れて守る（数には入れない）/ Map a unit to itself first to protect it (not counted) */
                    conversionMap[halfText] = halfText;
                    for (var k = 0; k < fullForms.length; k++) conversionMap[fullForms[k]] = fullForms[k];
                }
                if (!isConverting) continue;
                /* 半角にするなら全角が元、全角にするなら半角が元 / The source is the full-width form when converting to half-width, and vice versa */
                if (option === OPTION_HALF) {
                    for (var m = 0; m < fullForms.length; m++) {
                        conversionMap[fullForms[m]] = halfText;
                        rowKeyBySource[fullForms[m]] = row.key;
                    }
                } else {
                    conversionMap[halfText] = (fullForms.length === 1) ? fullForms[0] : fullForms;
                    rowKeyBySource[halfText] = row.key;
                }
                mapSize++;
            }
        }
        return {
            map: conversionMap,
            rowKeyBySource: rowKeyBySource,
            runUnits: runUnits,
            size: mapSize,
            digitsByLength: digitsByLength,
            spaceBetweenJapaneseOnly: (rowOptions.spaceBetweenJapaneseOnly === 1),
            japaneseLatinSpace: (rowOptions.japaneseLatinSpace === 1),
            punctuationStyle: rowOptions.punctuationStyle || 0,
            nakaguroToEllipsis: (rowOptions.nakaguroToEllipsis === 1),
            periodsToEllipsis: (rowOptions.periodsToEllipsis === 1),
            kutenToEllipsis: (rowOptions.kutenToEllipsis === 1),
            westernPunctuationSpace: (rowOptions.westernPunctuationSpace === 1),
            bracketInnerSpace: (rowOptions.bracketInnerSpace === 1),
            bracketOuterSpace: (rowOptions.bracketOuterSpace === 1),
            numberCommaToTouten: (rowOptions.numberCommaToTouten === 1),
            decimalToNakaguro: (rowOptions.decimalToNakaguro === 1),
            longDigitsFullwidth: (rowOptions.longDigitsFullwidth === 1),
            shortCapsFullwidth: (rowOptions.shortCapsFullwidth === 1)
        };
    }

    /**
     * 変換する処理が1つでもあるか
     * @param {Object} conversion - buildConversionMap() の戻り値
     * @returns {boolean} あれば true
     */
    function hasConversionWork(conversion) {
        return conversion.size > 0 || conversion.digitsByLength || conversion.punctuationStyle > 0 ||
            conversion.japaneseLatinSpace || conversion.nakaguroToEllipsis || conversion.periodsToEllipsis || conversion.kutenToEllipsis || conversion.westernPunctuationSpace ||
            conversion.bracketInnerSpace || conversion.bracketOuterSpace || conversion.numberCommaToTouten || conversion.decimalToNakaguro ||
            conversion.longDigitsFullwidth || conversion.shortCapsFullwidth;
    }

    /**
     * 段落の文字列の位置にある ' が、英数字にはさまれたアポストロフィか
     * @param {string} paragraphText - 段落の文字列
     * @param {number} offset - ' の位置
     * @returns {boolean} アポストロフィなら true
     */
    function isApostropheAt(paragraphText, offset) {
        return LATIN_ALNUM_PATTERN.test(paragraphText.charAt(offset - 1)) && LATIN_ALNUM_PATTERN.test(paragraphText.charAt(offset + 1));
    }

    /**
     * 文字列が同じ文字の2個以上の並びか
     * @param {string} sourceText - 調べる文字列
     * @param {string} repeatedChar - 文字
     * @returns {boolean} 並びなら true
     */
    function isRunOf(sourceText, repeatedChar) {
        if (sourceText.length < 2) return false;
        for (var i = 0; i < sourceText.length; i++) {
            if (sourceText.charAt(i) !== repeatedChar) return false;
        }
        return true;
    }

    /**
     * 文字列の中の文字をすべて別の文字に置き換える
     * @param {string} sourceText - 元の文字列
     * @param {string} fromChar - 置き換える文字
     * @param {string} toChar - 置き換え後の文字
     * @returns {string} 置き換えた文字列
     */
    function replaceAllChars(sourceText, fromChar, toChar) {
        return sourceText.split(fromChar).join(toChar);
    }

    /**
     * 文字列の位置にある文字の直前と直後の1文字を返す（段落の外は空文字。直後の段落の終わりは "\r"）
     * @param {string} fullText - 対象の文字列
     * @param {number} offset - 位置
     * @param {number} sourceLength - 文字の長さ
     * @returns {{previous: string, next: string}} 直前と直後の文字
     */
    function getNeighborChars(fullText, offset, sourceLength) {
        var previousChar = (offset > 0) ? fullText.charAt(offset - 1) : "";
        if (previousChar === "\r") previousChar = "";
        return { previous: previousChar, next: fullText.charAt(offset + sourceLength) };
    }

    /**
     * 向きのある引用符の開き・閉じを選ぶ。
     * 段落の中で前にある同じ引用符の数が偶数なら開き、奇数なら閉じ。英数字にはさまれた ' はアポストロフィ（閉じ）
     * @param {string} fullText - 対象の文字列
     * @param {number} offset - 引用符の位置
     * @param {string} sourceText - 見つかった文字（" または '）
     * @param {string[]} quoteForms - [開き, 閉じ]
     * @returns {string} 置き換える引用符
     */
    function chooseQuoteForm(fullText, offset, sourceText, quoteForms) {
        if (sourceText === "'" && isApostropheAt(fullText, offset)) return quoteForms[1];
        var paragraphStart = (offset > 0) ? fullText.lastIndexOf("\r", offset - 1) + 1 : 0;
        var quoteCount = 0;
        for (var i = paragraphStart; i < offset; i++) {
            if (fullText.charAt(i) !== sourceText) continue;
            if (sourceText === "'" && isApostropheAt(fullText, i)) continue;
            quoteCount++;
        }
        return (quoteCount % 2 === 0) ? quoteForms[0] : quoteForms[1];
    }

    /**
     * 位置から始まる変換表の単位の長さを返す。
     * 2個以上続く文字、2文字の組（半角の濁音・半濁音、!? など）、1文字の順に調べる（正規表現は使わない）
     * @param {string} fullText - 対象の文字列
     * @param {number} offset - 位置
     * @param {Object} conversion - buildConversionMap() の戻り値
     * @returns {number} 単位の長さ。当てはまらなければ 0
     */
    function matchTableUnitLength(fullText, offset, conversion) {
        var currentChar = fullText.charAt(offset);
        for (var i = 0; i < conversion.runUnits.length; i++) {
            var runUnit = conversion.runUnits[i];
            if (currentChar !== runUnit.halfChar && currentChar !== runUnit.fullChar) continue;
            var runEnd = offset + 1;
            while (fullText.charAt(runEnd) === currentChar) runEnd++;
            if (runEnd - offset >= 2) return runEnd - offset;
        }
        if (offset + 2 <= fullText.length && conversion.map.hasOwnProperty(fullText.substr(offset, 2))) return 2;
        if (conversion.map.hasOwnProperty(currentChar)) return 1;
        return 0;
    }

    /**
     * 変換表で置き換える文字を決める
     * @param {string} fullText - 対象の文字列
     * @param {number} offset - 位置
     * @param {string} sourceText - 見つかった文字列
     * @param {Object} conversion - buildConversionMap() の戻り値
     * @returns {{key: string, text: string}|null} 置き換え。変えなければ null
     */
    function decideTableReplacement(fullText, offset, sourceText, conversion) {
        for (var i = 0; i < conversion.runUnits.length; i++) {
            var runUnit = conversion.runUnits[i];
            if (!isRunOf(sourceText, runUnit.halfChar) && !isRunOf(sourceText, runUnit.fullChar)) continue;
            /* 2個以上続く文字。「そのまま」なら守るだけ / A run of characters; only protected when kept */
            var runText = sourceText;
            if (runUnit.option === OPTION_HALF) runText = replaceAllChars(sourceText, runUnit.fullChar, runUnit.halfChar);
            else if (runUnit.option === OPTION_FULL) runText = replaceAllChars(sourceText, runUnit.halfChar, runUnit.fullChar);
            return (runText === sourceText) ? null : { key: runUnit.key, text: runText };
        }
        if (!conversion.map.hasOwnProperty(sourceText)) return null;
        var replacement = conversion.map[sourceText];
        /* 守るだけの組（元と同じ）は書き換えない / Units that are only protected (mapped to themselves) are left untouched */
        if (replacement === sourceText) return null;
        if (isLeftToLaterPass(fullText, offset, conversion.rowKeyBySource[sourceText], conversion)) return null;
        if (conversion.spaceBetweenJapaneseOnly && (sourceText === " " || sourceText === "　")) {
            var neighbors = getNeighborChars(fullText, offset, sourceText.length);
            if (!JAPANESE_CHAR_PATTERN.test(neighbors.previous) || !JAPANESE_CHAR_PATTERN.test(neighbors.next)) return null;
        }
        if (replacement instanceof Array) replacement = chooseQuoteForm(fullText, offset, sourceText, replacement);
        return { key: conversion.rowKeyBySource[sourceText], text: replacement };
    }

    /**
     * 変換表に従って置き換える箇所を、前から順に集める
     * @param {string} fullText - 対象の文字列
     * @param {Object} conversion - buildConversionMap() の戻り値
     * @returns {Array<{offset: number, sourceText: string, text: string, key: string}>} 置き換える箇所
     */
    function findTableReplacements(fullText, conversion) {
        var replacements = [];
        var offset = 0;
        while (offset < fullText.length) {
            var unitLength = matchTableUnitLength(fullText, offset, conversion);
            if (unitLength === 0) {
                offset++;
                continue;
            }
            var sourceText = fullText.substr(offset, unitLength);
            var replacement = decideTableReplacement(fullText, offset, sourceText, conversion);
            if (replacement) replacements.push({ offset: offset, sourceText: sourceText, text: replacement.text, key: replacement.key });
            offset += unitLength;
        }
        return replacements;
    }

    /**
     * 3つ続く中黒・ピリオド・句点を三点リーダーにする箇所を集める（半角・全角が混ざっていてもよい）
     * @param {string} fullText - 対象の文字列
     * @param {Object} conversion - buildConversionMap() の戻り値（nakaguroToEllipsis などでどれを置き換えるかを見る）
     * @returns {Array<{offset: number, sourceText: string, text: string, key: string}>} 置き換える箇所
     */
    function findEllipsisReplacements(fullText, conversion) {
        var replacements = [];
        var offset = 0;
        while (offset <= fullText.length - 3) {
            var matchedKey = null;
            for (var i = 0; i < ELLIPSIS_SOURCES.length; i++) {
                var source = ELLIPSIS_SOURCES[i];
                if (!conversion[source.key]) continue;
                if (source.chars.indexOf(fullText.charAt(offset)) >= 0 && source.chars.indexOf(fullText.charAt(offset + 1)) >= 0 &&
                    source.chars.indexOf(fullText.charAt(offset + 2)) >= 0) {
                    matchedKey = source.key;
                    break;
                }
            }
            if (!matchedKey) {
                offset++;
                continue;
            }
            replacements.push({ offset: offset, sourceText: fullText.substring(offset, offset + 3), text: "\u2026", key: matchedKey });
            offset += 3;
        }
        return replacements;
    }

    /**
     * 句読点の種類をそろえる箇所を集める
     * @param {string} fullText - 対象の文字列
     * @param {Object} styleMap - 元の文字 → 置き換え後の文字（PUNCTUATION_STYLE_MAPS の要素）
     * @returns {Array<{offset: number, sourceText: string, text: string, key: string}>} 置き換える箇所
     */
    function findPunctuationStyleReplacements(fullText, styleMap) {
        var replacements = [];
        for (var i = 0; i < fullText.length; i++) {
            var currentChar = fullText.charAt(i);
            if (styleMap.hasOwnProperty(currentChar)) replacements.push({ offset: i, sourceText: currentChar, text: styleMap[currentChar], key: "punctuationStyle" });
        }
        return replacements;
    }

    /**
     * 位置から始まる数字のつながりを読む（3.5 や 1,000 のように、数字にはさまれた区切りもつながりに含める）
     * @param {string} text - 文字列
     * @param {number} startOffset - つながりの先頭（数字）の位置
     * @returns {{end: number, digitCount: number}} 終わり（含まない）と数字の数
     */
    function readDigitRun(text, startOffset) {
        var runEnd = startOffset;
        var digitCount = 0;
        while (runEnd < text.length) {
            if (DIGIT_PATTERN.test(text.charAt(runEnd))) {
                digitCount++;
                runEnd++;
            } else if (DIGIT_SEPARATORS.indexOf(text.charAt(runEnd)) >= 0 && DIGIT_PATTERN.test(text.charAt(runEnd + 1))) {
                runEnd++;
            } else {
                break;
            }
        }
        return { end: runEnd, digitCount: digitCount };
    }

    /**
     * 数字のつながりを［3桁以上の数字は全角にする］で全角にするか（英字に続くものは除く）
     * @param {string} text - 文字列
     * @param {number} runStart - つながりの先頭
     * @param {{end: number, digitCount: number}} digitRun - readDigitRun() の戻り値
     * @returns {boolean} 全角にするなら true
     */
    function isLongDigitRun(text, runStart, digitRun) {
        return digitRun.digitCount >= LONG_DIGIT_COUNT &&
            !LATIN_LETTER_PATTERN.test(text.charAt(runStart - 1)) && !LATIN_LETTER_PATTERN.test(text.charAt(digitRun.end));
    }

    /**
     * 位置の数字が、［3桁以上の数字は全角にする］で全角にするつながりの中にあるか
     * @param {string} text - 文字列
     * @param {number} offset - 数字の位置
     * @returns {boolean} 中にあれば true
     */
    function isInLongDigitRun(text, offset) {
        var runStart = offset;
        while (runStart > 0) {
            if (DIGIT_PATTERN.test(text.charAt(runStart - 1))) runStart--;
            else if (DIGIT_SEPARATORS.indexOf(text.charAt(runStart - 1)) >= 0 && DIGIT_PATTERN.test(text.charAt(runStart - 2))) runStart -= 2;
            else break;
        }
        return isLongDigitRun(text, runStart, readDigitRun(text, runStart));
    }

    /**
     * 位置の文字が英単語の一部か（つなぎは前後が英数字のときだけ）
     * @param {string} text - 文字列
     * @param {number} offset - 位置
     * @returns {boolean} 英単語の一部なら true
     */
    function isWordCharAt(text, offset) {
        var currentChar = text.charAt(offset);
        if (currentChar === "") return false;
        if (WORD_CHAR_PATTERN.test(currentChar)) return true;
        return WORD_JOINERS.indexOf(currentChar) >= 0 && WORD_CHAR_PATTERN.test(text.charAt(offset - 1)) && WORD_CHAR_PATTERN.test(text.charAt(offset + 1));
    }

    /**
     * 位置の文字を含む英単語の範囲を返す
     * @param {string} text - 文字列
     * @param {number} offset - 英単語の中の位置
     * @returns {{start: number, end: number}} 先頭と終わり（含まない）
     */
    function findWordBounds(text, offset) {
        var wordStart = offset;
        while (wordStart > 0 && isWordCharAt(text, wordStart - 1)) wordStart--;
        var wordEnd = offset + 1;
        while (wordEnd < text.length && isWordCharAt(text, wordEnd)) wordEnd++;
        return { start: wordStart, end: wordEnd };
    }

    /**
     * 位置の文字が、［3文字以内の大文字で構成する英単語は全角にする］の対象の語の中にあるか
     * @param {string} text - 文字列
     * @param {number} offset - 位置
     * @returns {boolean} 中にあれば true
     */
    function isInShortCapsWord(text, offset) {
        var wordBounds = findWordBounds(text, offset);
        return SHORT_CAPS_WORD_PATTERN.test(text.substring(wordBounds.start, wordBounds.end));
    }

    /**
     * 数字の幅をそろえる（区切りはそのまま）
     * @param {string} sourceText - 数字のつながり
     * @param {boolean} toFullWidth - true なら全角、false なら半角
     * @returns {string} そろえた文字列
     */
    function setDigitWidth(sourceText, toFullWidth) {
        var digitText = "";
        for (var i = 0; i < sourceText.length; i++) {
            var charCode = sourceText.charCodeAt(i);
            if (toFullWidth && charCode >= 0x30 && charCode <= 0x39) charCode += FULL_WIDTH_OFFSET;
            else if (!toFullWidth && charCode >= 0xFF10 && charCode <= 0xFF19) charCode -= FULL_WIDTH_OFFSET;
            digitText += String.fromCharCode(charCode);
        }
        return digitText;
    }

    /**
     * 数字のつながりの置き換えを決める。［3桁以上の数字は全角にする］の対象なら全角、
     * ［数字は1桁を全角、2桁以上を半角にする］なら1桁を全角・2桁以上を半角（英字に続く数字は半角）、どちらでもなければ変えない
     * @param {string} text - 文字列
     * @param {number} runStart - つながりの先頭
     * @param {{end: number, digitCount: number}} digitRun - readDigitRun() の戻り値
     * @param {{digitsByLength: boolean, longDigitsFullwidth: boolean}} conversion - どちらの処理をするか
     * @returns {{key: string, text: string}|null} 置き換え。変えなければ null
     */
    function decideDigitRunReplacement(text, runStart, digitRun, conversion) {
        var sourceText = text.substring(runStart, digitRun.end);
        var replacementKey;
        var toFullWidth;
        if (conversion.longDigitsFullwidth && isLongDigitRun(text, runStart, digitRun)) {
            replacementKey = "longDigitsFullwidth";
            toFullWidth = true;
        } else if (conversion.digitsByLength) {
            replacementKey = "digits";
            toFullWidth = (digitRun.digitCount === 1 && !LATIN_LETTER_PATTERN.test(text.charAt(runStart - 1)) && !LATIN_LETTER_PATTERN.test(text.charAt(digitRun.end)));
        } else {
            return null;
        }
        var digitText = setDigitWidth(sourceText, toFullWidth);
        return (digitText === sourceText) ? null : { key: replacementKey, text: digitText };
    }

    /**
     * 英単語の置き換えを決める（大文字だけの3文字以内の語なら、半角の大文字を全角にする）
     * @param {string} text - 文字列
     * @param {{start: number, end: number}} wordBounds - 英単語の範囲
     * @returns {{key: string, text: string}|null} 置き換え。変えなければ null
     */
    function decideShortCapsReplacement(text, wordBounds) {
        var wordText = text.substring(wordBounds.start, wordBounds.end);
        if (!SHORT_CAPS_WORD_PATTERN.test(wordText)) return null;
        var fullText = "";
        for (var i = 0; i < wordText.length; i++) {
            var charCode = wordText.charCodeAt(i);
            fullText += String.fromCharCode((charCode >= 0x41 && charCode <= 0x5A) ? charCode + FULL_WIDTH_OFFSET : charCode);
        }
        return (fullText === wordText) ? null : { key: "shortCapsFullwidth", text: fullText };
    }

    /**
     * 変換表の置き換えを、［3桁以上の数字は全角にする］［3文字以内の大文字で構成する英単語は全角にする］に任せるか
     * （その処理で全角にする文字は、数字・アルファベットの行で変えない）
     * @param {string} text - 文字列
     * @param {number} offset - 置き換える文字の位置
     * @param {string} rowKey - 置き換える文字の行のキー
     * @param {{longDigitsFullwidth: boolean, shortCapsFullwidth: boolean}} conversion - どちらの処理をするか
     * @returns {boolean} 任せるなら true
     */
    function isLeftToLaterPass(text, offset, rowKey, conversion) {
        if (rowKey === "digits") return conversion.longDigitsFullwidth && isInLongDigitRun(text, offset);
        if (rowKey === "alphabet") return conversion.shortCapsFullwidth && isInShortCapsWord(text, offset);
        return false;
    }

    /**
     * 数字のつながりを置き換える箇所を集める（置き換え方は decideDigitRunReplacement()）。
     * 3.5 や 1,000 のように、数字にはさまれた区切りはひとつながりとして数える
     * @param {string} fullText - 対象の文字列
     * @param {Object} conversion - buildConversionMap() の戻り値
     * @returns {Array<{offset: number, sourceText: string, text: string, key: string}>} 置き換える箇所
     */
    function findDigitRunReplacements(fullText, conversion) {
        var replacements = [];
        var offset = 0;
        while (offset < fullText.length) {
            if (!DIGIT_PATTERN.test(fullText.charAt(offset))) {
                offset++;
                continue;
            }
            var digitRun = readDigitRun(fullText, offset);
            var replacement = decideDigitRunReplacement(fullText, offset, digitRun, conversion);
            if (replacement) replacements.push({ offset: offset, sourceText: fullText.substring(offset, digitRun.end), text: replacement.text, key: replacement.key });
            offset = digitRun.end;
        }
        return replacements;
    }

    /**
     * 大文字だけでできた3文字以内の英単語を全角にする箇所を集める
     * @param {string} fullText - 対象の文字列
     * @returns {Array<{offset: number, sourceText: string, text: string, key: string}>} 置き換える箇所
     */
    function findShortCapsReplacements(fullText) {
        var replacements = [];
        var offset = 0;
        while (offset < fullText.length) {
            if (!isWordCharAt(fullText, offset)) {
                offset++;
                continue;
            }
            var wordBounds = findWordBounds(fullText, offset);
            var replacement = decideShortCapsReplacement(fullText, wordBounds);
            if (replacement) {
                replacements.push({ offset: wordBounds.start, sourceText: fullText.substring(wordBounds.start, wordBounds.end), text: replacement.text, key: replacement.key });
            }
            offset = wordBounds.end;
        }
        return replacements;
    }

    /**
     * 数字の中の「,」「.」を縦組み用に置き換える箇所を集める（後ろがちょうど3桁の「,」は「、」、「.」は「・」）
     * @param {string} fullText - 対象の文字列
     * @param {{numberCommaToTouten: boolean, decimalToNakaguro: boolean}} conversion - どちらを置き換えるか
     * @returns {Array<{offset: number, sourceText: string, text: string, key: string}>} 置き換える箇所
     */
    function findNumberMarkReplacements(fullText, conversion) {
        var replacements = [];
        for (var i = 1; i < fullText.length - 1; i++) {
            if (!DIGIT_PATTERN.test(fullText.charAt(i - 1))) continue;
            var currentChar = fullText.charAt(i);
            if (conversion.numberCommaToTouten && (currentChar === "," || currentChar === "\uFF0C") && isThreeDigitGroupAt(fullText, i + 1)) {
                replacements.push({ offset: i, sourceText: currentChar, text: "\u3001", key: "numberCommaToTouten" });
            } else if (conversion.decimalToNakaguro && (currentChar === "." || currentChar === "\uFF0E") && DIGIT_PATTERN.test(fullText.charAt(i + 1))) {
                replacements.push({ offset: i, sourceText: currentChar, text: "\u30FB", key: "decimalToNakaguro" });
            }
        }
        return replacements;
    }

    /**
     * 位置から数字がちょうど3桁続くか（4桁目は数字でない）
     * @param {string} fullText - 対象の文字列
     * @param {number} offset - 調べ始める位置
     * @returns {boolean} 3桁ちょうどなら true
     */
    function isThreeDigitGroupAt(fullText, offset) {
        for (var i = 0; i < 3; i++) {
            if (!DIGIT_PATTERN.test(fullText.charAt(offset + i))) return false;
        }
        return !DIGIT_PATTERN.test(fullText.charAt(offset + 3));
    }

    /**
     * 欧文の「,」「.」「:」のあとにスペースを入れる箇所を集める（前後の文字は WESTERN_PUNCTUATION_SPACE_RULES）
     * @param {string} fullText - 対象の文字列
     * @returns {Array<{offset: number, sourceText: string, text: string, key: string}>} 置き換える箇所（text は記号＋スペース）
     */
    function findWesternPunctuationSpaceReplacements(fullText) {
        var replacements = [];
        for (var i = 1; i < fullText.length - 1; i++) {
            var currentChar = fullText.charAt(i);
            if (!WESTERN_PUNCTUATION_SPACE_RULES.hasOwnProperty(currentChar)) continue;
            var rule = WESTERN_PUNCTUATION_SPACE_RULES[currentChar];
            if (rule.previous.test(fullText.charAt(i - 1)) && rule.next.test(fullText.charAt(i + 1))) {
                replacements.push({ offset: i, sourceText: currentChar, text: currentChar + " ", key: "westernPunctuationSpace" });
            }
        }
        return replacements;
    }

    /**
     * 和文の文字と英数字のあいだのスペース（半角・全角が続いていてもまとめて）を削除する箇所を集める
     * @param {string} fullText - 対象の文字列
     * @returns {Array<{offset: number, sourceText: string, text: string, key: string}>} 削除する箇所（text は空文字）
     */
    function findJapaneseLatinSpaceReplacements(fullText) {
        var replacements = [];
        var offset = 0;
        while (offset < fullText.length) {
            if (BRACKET_SPACES.indexOf(fullText.charAt(offset)) < 0) {
                offset++;
                continue;
            }
            var runEnd = offset;
            while (runEnd < fullText.length && BRACKET_SPACES.indexOf(fullText.charAt(runEnd)) >= 0) runEnd++;
            var previousChar = (offset > 0) ? fullText.charAt(offset - 1) : "";
            var nextChar = fullText.charAt(runEnd);
            var isTarget = (JAPANESE_CHAR_PATTERN.test(previousChar) && LATIN_ALNUM_PATTERN.test(nextChar)) ||
                (LATIN_ALNUM_PATTERN.test(previousChar) && JAPANESE_CHAR_PATTERN.test(nextChar));
            if (isTarget) replacements.push({ offset: offset, sourceText: fullText.substring(offset, runEnd), text: "", key: "japaneseLatinSpace" });
            offset = runEnd;
        }
        return replacements;
    }

    /**
     * 括弧のまわりのスペースを削除する箇所を集める（後読みが使えないので、前後の文字を自前で調べる）
     * @param {string} fullText - 対象の文字列
     * @param {boolean} isInner - true なら内側（開き括弧の直後・閉じ括弧の直前）、false なら外側（全角の括弧の直前・直後）
     * @returns {Array<{offset: number, sourceText: string, text: string, key: string}>} 削除する箇所（text は空文字）
     */
    function findBracketSpaceReplacements(fullText, isInner) {
        var replacements = [];
        var openingBrackets = isInner ? OPENING_BRACKETS : FULL_WIDTH_OPENING_BRACKETS;
        var closingBrackets = isInner ? CLOSING_BRACKETS : FULL_WIDTH_CLOSING_BRACKETS;
        var offset = 0;
        while (offset < fullText.length) {
            if (BRACKET_SPACES.indexOf(fullText.charAt(offset)) < 0) {
                offset++;
                continue;
            }
            var runEnd = offset;
            while (runEnd < fullText.length && BRACKET_SPACES.indexOf(fullText.charAt(runEnd)) >= 0) runEnd++;
            var previousChar = (offset > 0) ? fullText.charAt(offset - 1) : "";
            var nextChar = fullText.charAt(runEnd);
            /* 内側は開き括弧のあと・閉じ括弧の前、外側は開き括弧の前・閉じ括弧のあと / Inside: after opening or before closing; outside: before opening or after closing */
            var isTarget = isInner
                ? ((previousChar !== "" && openingBrackets.indexOf(previousChar) >= 0) || (nextChar !== "" && closingBrackets.indexOf(nextChar) >= 0))
                : ((nextChar !== "" && openingBrackets.indexOf(nextChar) >= 0) || (previousChar !== "" && closingBrackets.indexOf(previousChar) >= 0));
            if (isTarget) {
                replacements.push({
                    offset: offset,
                    sourceText: fullText.substring(offset, runEnd),
                    text: "",
                    key: isInner ? "bracketInnerSpace" : "bracketOuterSpace"
                });
            }
            offset = runEnd;
        }
        return replacements;
    }

    /**
     * ［変換した文字に色を付ける］の色を、ドキュメントのカラーモードに合わせて作る
     * @returns {RGBColor|CMYKColor} 色
     */
    function createHighlightColor() {
        if (activeDocument.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = HIGHLIGHT_CMYK[0];
            cmykColor.magenta = HIGHLIGHT_CMYK[1];
            cmykColor.yellow = HIGHLIGHT_CMYK[2];
            cmykColor.black = HIGHLIGHT_CMYK[3];
            return cmykColor;
        }
        var rgbColor = new RGBColor();
        rgbColor.red = HIGHLIGHT_RGB[0];
        rgbColor.green = HIGHLIGHT_RGB[1];
        rgbColor.blue = HIGHLIGHT_RGB[2];
        return rgbColor;
    }

    /**
     * 対象の今の文字列を返す
     * @param {{base: TextRange, start: number, length: number}} target - 対象
     * @returns {string|null} 文字列。読めなければ null
     */
    function getTargetText(target) {
        var baseText = target.base.contents;
        if (typeof baseText !== "string") return null;
        return (target.length < 0) ? baseText : baseText.substr(target.start, target.length);
    }

    /**
     * 1か所を置き換える。書式を残すため characters 単位で書き換える：
     * 長さが同じなら変わる文字だけを差し替え、長さが変わるときは先頭の1文字に新しい文字列を入れて残りを消す
     * @param {{base: TextRange, start: number, length: number}} target - 対象
     * @param {{offset: number, sourceText: string, text: string}} replacement - 置き換える箇所
     * @param {RGBColor|CMYKColor|null} highlightColor - 変換した文字に付ける色（付けないなら null）。書き換える前の文字に付け、書き換えで引き継がせる
     * @returns {void}
     */
    function applyReplacement(target, replacement, highlightColor) {
        var firstIndex = target.start + replacement.offset;
        var sourceText = replacement.sourceText;
        var replacementText = replacement.text;
        if (replacementText.length === sourceText.length) {
            for (var i = sourceText.length - 1; i >= 0; i--) {
                if (sourceText.charAt(i) === replacementText.charAt(i)) continue;
                var changedChar = target.base.characters[firstIndex + i];
                if (highlightColor) changedChar.characterAttributes.fillColor = highlightColor;
                changedChar.contents = replacementText.charAt(i);
            }
            return;
        }
        for (var j = sourceText.length - 1; j >= 1; j--) target.base.characters[firstIndex + j].remove();
        var firstChar = target.base.characters[firstIndex];
        if (replacementText === "") {
            firstChar.remove();
            return;
        }
        if (highlightColor) firstChar.characterAttributes.fillColor = highlightColor;
        firstChar.contents = replacementText;
    }

    /**
     * 対象の文字列から置き換える箇所を集め、後ろから置き換える（前の箇所の位置がずれないように）
     * @param {{base: TextRange, start: number, length: number}} target - 対象
     * @param {Function} findReplacements - 文字列 → 置き換える箇所の配列
     * @param {{converted: number, failed: number, countsByKey: Object}} result - 数を足していく結果
     * @param {RGBColor|CMYKColor|null} highlightColor - 変換した文字に付ける色（付けないなら null）
     * @returns {void}
     */
    function runReplacePass(target, findReplacements, result, highlightColor) {
        var targetText;
        try {
            targetText = getTargetText(target);
        } catch (e) {
            return;
        }
        if (targetText === null) return;
        var replacements = findReplacements(targetText);
        for (var i = replacements.length - 1; i >= 0; i--) {
            var replacement = replacements[i];
            try {
                applyReplacement(target, replacement, highlightColor);
                /* 選択した文字の範囲は、増えた・減った文字数に合わせる / Keep a selected range in step with added or removed characters */
                if (target.length >= 0) target.length += replacement.text.length - replacement.sourceText.length;
                result.converted++;
                result.countsByKey[replacement.key] = (result.countsByKey[replacement.key] || 0) + 1;
            } catch (e2) {
                result.failed++;
            }
        }
    }

    /**
     * 対象の文字を設定に従って変換する。
     * 三点リーダー → 和欧間のスペース → 変換表 → 句読点の種類 → 数字の桁数 → 3文字以内の大文字の英単語 → 数字の中の「,」「.」 → 欧文の,.:のあとのスペース → 括弧のまわりのスペースの順に処理する。characters 単位の書き換えなので、文字の書式はそのまま残る
     * @param {Array} targets - 対象の配列
     * @param {Object} conversion - buildConversionMap() の戻り値に highlightColor（変換した文字に付ける色、付けないなら null）を足したもの
     * @returns {{converted: number, failed: number, countsByKey: Object}} 変換できた数・できなかった数・キーごとの数
     */
    function convertTargets(targets, conversion) {
        var result = { converted: 0, failed: 0, countsByKey: {} };
        var styleMap = PUNCTUATION_STYLE_MAPS[conversion.punctuationStyle];
        var highlightColor = conversion.highlightColor || null;
        for (var i = 0; i < targets.length; i++) {
            var target = targets[i];
            /* 行の変換より先に置き換える（中黒・ピリオド・句点が1つずつ変換されて数が重ならないように）/ Before the rows, so the dots are not converted one by one first */
            if (conversion.nakaguroToEllipsis || conversion.periodsToEllipsis || conversion.kutenToEllipsis) {
                runReplacePass(target, function (fullText) { return findEllipsisReplacements(fullText, conversion); }, result, highlightColor);
            }
            /* 英数字を全角にする前に、元の文字で和欧間を判定する / Judged on the original characters, before letters and digits turn full-width */
            if (conversion.japaneseLatinSpace) runReplacePass(target, findJapaneseLatinSpaceReplacements, result, highlightColor);
            if (conversion.size > 0) {
                runReplacePass(target, function (fullText) { return findTableReplacements(fullText, conversion); }, result, highlightColor);
            }
            if (styleMap) {
                runReplacePass(target, function (fullText) { return findPunctuationStyleReplacements(fullText, styleMap); }, result, highlightColor);
            }
            if (conversion.digitsByLength || conversion.longDigitsFullwidth) {
                runReplacePass(target, function (fullText) { return findDigitRunReplacements(fullText, conversion); }, result, highlightColor);
            }
            if (conversion.shortCapsFullwidth) runReplacePass(target, findShortCapsReplacements, result, highlightColor);
            /* 数字の桁数を決めたあとに区切りを置き換える（「、」「・」で数字のつながりが切れないように）/ After the digit lengths, so 、 and ・ do not split the runs */
            if (conversion.numberCommaToTouten || conversion.decimalToNakaguro) {
                runReplacePass(target, function (fullText) { return findNumberMarkReplacements(fullText, conversion); }, result, highlightColor);
            }
            /* 行の変換のあと（半角のまま残った「,」「.」「:」だけが対象）/ After the rows, so only "," "." ":" left half-width are touched */
            if (conversion.westernPunctuationSpace) runReplacePass(target, findWesternPunctuationSpaceReplacements, result, highlightColor);
            if (conversion.bracketInnerSpace) {
                runReplacePass(target, function (fullText) { return findBracketSpaceReplacements(fullText, true); }, result, highlightColor);
            }
            if (conversion.bracketOuterSpace) {
                runReplacePass(target, function (fullText) { return findBracketSpaceReplacements(fullText, false); }, result, highlightColor);
            }
        }
        return result;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

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

    /* 列の見出し（選択肢の順）/ Column headings, in option order */
    var OPTION_COLUMNS = [
        { label: LABELS.column.keep, tooltip: LABELS.tooltip.keep },
        { label: LABELS.column.half, tooltip: LABELS.tooltip.half },
        { label: LABELS.column.full, tooltip: LABELS.tooltip.full }
    ];

    /**
     * 行の選択肢を1つだけ選んだ状態にする（ラジオはセルごとに別グループなので排他を自前で行う）。
     * 表示前の value は読み戻せないので、選んだ値は rowControl.option に持つ
     * @param {{radios: RadioButton[], option: number}} rowControl - 行のコントロール
     * @param {number} option - 選ぶ選択肢の番号
     * @returns {void}
     */
    function setRowOption(rowControl, option) {
        rowControl.option = option;
        for (var i = 0; i < rowControl.radios.length; i++) {
            rowControl.radios[i].value = (i === option);
        }
    }

    /**
     * 行で選ばれている選択肢の番号を返す
     * @param {{radios: RadioButton[], option: number}} rowControl - 行のコントロール
     * @returns {number} 選択肢の番号
     */
    function getRowOption(rowControl) {
        return rowControl.option;
    }

    /**
     * 行のほかの設定のチェックボックスを足す（値は 0／1。表示前の value は読み戻せないので値は自前で持つ）
     * @param {Panel|Group} parent - 足す先
     * @param {string} settingKey - 設定のキー
     * @param {Object} labelEntry - 文言の { ja, en }
     * @param {Object} tooltipEntry - ツールチップの { ja, en }
     * @param {{extraControls: Object, onRowChange: Function}} tableContext - 作ったコントロールを足す先と、値を変えたときの処理
     * @returns {void}
     */
    function addExtraCheckbox(parent, settingKey, labelEntry, tooltipEntry, tableContext) {
        var extraCheckbox = parent.add("checkbox", undefined, getLabel(labelEntry));
        extraCheckbox.helpTip = getLabel(tooltipEntry);
        var extraControl = {
            value: 0,
            setValue: function (value) {
                extraControl.value = value;
                extraCheckbox.value = (value === 1);
            }
        };
        extraCheckbox.onClick = function () {
            extraControl.value = extraCheckbox.value ? 1 : 0;
            tableContext.onRowChange();
        };
        tableContext.extraControls[settingKey] = extraControl;
    }

    /**
     * 行のほかの設定の「項目名 : ドロップダウン」を足す（値は項目の番号）
     * @param {Panel|Group} parent - 足す先
     * @param {string} settingKey - 設定のキー
     * @param {Object} fieldLabelEntry - 項目名の { ja, en }
     * @param {Object[]} itemEntries - 項目の { ja, en } の配列
     * @param {Object} tooltipEntry - ツールチップの { ja, en }
     * @param {{extraControls: Object, onRowChange: Function}} tableContext - 作ったコントロールを足す先と、値を変えたときの処理
     * @returns {void}
     */
    function addExtraDropdown(parent, settingKey, fieldLabelEntry, itemEntries, tooltipEntry, tableContext) {
        var dropdownRowGroup = parent.add("group");
        setupRow(dropdownRowGroup, "left", 6);
        var fieldLabel = dropdownRowGroup.add("statictext", undefined, labelText(fieldLabelEntry));
        var itemLabels = [];
        for (var i = 0; i < itemEntries.length; i++) itemLabels.push(getLabel(itemEntries[i]));
        var extraDropdown = dropdownRowGroup.add("dropdownlist", undefined, itemLabels);
        extraDropdown.helpTip = getLabel(tooltipEntry);
        var extraControl = {
            value: 0,
            isSetting: false,
            setValue: function (value) {
                extraControl.value = value;
                extraControl.isSetting = true;
                extraDropdown.selection = value;
                extraControl.isSetting = false;
            },
            setEnabled: function (isEnabled) {
                fieldLabel.enabled = isEnabled;
                extraDropdown.enabled = isEnabled;
            }
        };
        extraDropdown.onChange = function () {
            if (extraControl.isSetting || !extraDropdown.selection) return;
            extraControl.value = extraDropdown.selection.index;
            tableContext.onRowChange();
        };
        extraControl.setValue(0);
        tableContext.extraControls[settingKey] = extraControl;
    }

    /**
     * 複数のラベル（statictext）の幅を最長のものへ揃え、指定方向に揃えます。
     * 実際の描画幅（preferredSize.width）を測るため、文言長やロケールに依存しません。
     * @param {StaticText[]} labelControls - 幅を揃える statictext の配列。
     * @param {string} [justify] - 揃え方向 "left" | "center" | "right"（省略時は "right"）。
     * @returns {number} 揃えた後の共通ラベル幅（px）。
     */
    function alignLabelWidths(labelControls, justify) {
        if (!labelControls || !labelControls.length) return 0;
        justify = justify || 'right';

        /* 1パス目：各ラベルの自然幅を測り、最大値を求める / Pass 1: find the widest natural width */
        var maxLabelWidth = 0;
        for (var i = 0; i < labelControls.length; i++) {
            /* preferredSize を返さない環境では実サイズで代用する / fall back to the laid-out size */
            var naturalWidth = labelControls[i].preferredSize.width || labelControls[i].size.width;
            if (naturalWidth > maxLabelWidth) maxLabelWidth = naturalWidth;
        }

        /* 2パス目：全ラベルを最大幅に固定し、揃え方向を適用 / Pass 2: apply common width and justification */
        for (var j = 0; j < labelControls.length; j++) {
            labelControls[j].preferredSize.width = maxLabelWidth;
            labelControls[j].justify = justify;
        }

        return maxLabelWidth;
    }

    /**
     * 表の見出し行（列名）を足す
     * @param {Group} tableGroup - 表のグループ
     * @returns {StaticText} 項目名の列の上の空きラベル（幅はあとで項目名に揃える）
     */
    function addTableHeader(tableGroup) {
        var headerRowGroup = tableGroup.add("group");
        setupRow(headerRowGroup, "left", TABLE_CELL_SPACING);
        var labelSpacer = headerRowGroup.add("statictext", undefined, "");
        for (var i = 0; i < OPTION_COLUMNS.length; i++) {
            var headerText = headerRowGroup.add("statictext", undefined, getLabel(OPTION_COLUMNS[i].label));
            headerText.preferredSize.width = OPTION_COLUMN_WIDTH;
            headerText.justify = "center";
            headerText.helpTip = getLabel(OPTION_COLUMNS[i].tooltip);
        }
        return labelSpacer;
    }

    /**
     * ラジオをクリックしたときの処理を作る。Option キーを押していれば、同じパネルの行をすべて同じ選択肢にする
     * @param {{radios: RadioButton[]}} rowControl - 行のコントロール
     * @param {number} option - このラジオの選択肢の番号
     * @param {Array} panelRows - 同じパネルの行の定義の配列
     * @param {{rowControls: Object, onRowChange: Function}} tableContext - 行のコントロールと、選択肢を変えたあとに呼ぶ関数
     * @returns {Function} onClick に入れる関数
     */
    function makeRadioClickHandler(rowControl, option, panelRows, tableContext) {
        return function () {
            if (ScriptUI.environment.keyboardState.altKey) {
                for (var i = 0; i < panelRows.length; i++) {
                    var panelRowControl = tableContext.rowControls[panelRows[i].key];
                    if (panelRowControl) setRowOption(panelRowControl, option);
                }
            } else {
                setRowOption(rowControl, option);
            }
            tableContext.onRowChange();
        };
    }

    /**
     * 「項目名＋3つの選択肢」の表を足す
     * @param {Panel|Group} parent - 表を足す先
     * @param {Array} rows - 行の定義の配列
     * @param {{rowControls: Object, rowLabels: StaticText[], headerSpacers: StaticText[], onRowChange: Function}} tableContext -
     *     行のコントロール（キー → コントロール）、幅を揃える項目名・見出しの空きラベル（作ったものを足す）、選択肢を変えたときの処理
     * @returns {void}
     */
    function addOptionTable(parent, rows, tableContext) {
        var tableGroup = parent.add("group");
        tableGroup.orientation = "column";
        tableGroup.alignChildren = ["left", "top"];
        tableGroup.alignment = ["left", "top"];
        tableGroup.spacing = TABLE_ROW_SPACING;
        tableContext.headerSpacers.push(addTableHeader(tableGroup));

        for (var i = 0; i < rows.length; i++) {
            var rowGroup = tableGroup.add("group");
            setupRow(rowGroup, "left", TABLE_CELL_SPACING);
            var rowLabel = rowGroup.add("statictext", undefined, getLabel(LABELS.row[rows[i].key]) + (rows[i].glyph ? "  " + rows[i].glyph : ""));
            tableContext.rowLabels.push(rowLabel);

            var rowControl = { radios: [], option: OPTION_KEEP };
            for (var option = 0; option < OPTION_COLUMNS.length; option++) {
                /* ラジオを列の中央に置くため、セルごとにグループで包む / Wrap each radio in a cell group to center it under the heading */
                var cellGroup = rowGroup.add("group");
                /* preferredSize はレイアウトの計算し直しで子の幅に戻るので、最小幅で取る / preferredSize is reset when the layout is recalculated, so reserve a minimum width */
                cellGroup.minimumSize.width = OPTION_COLUMN_WIDTH;
                cellGroup.alignChildren = ["center", "center"];
                var optionRadio = cellGroup.add("radiobutton", undefined, "");
                optionRadio.helpTip = getLabel(OPTION_COLUMNS[option].tooltip) + "\n" + getLabel(LABELS.tooltip.optionClick);
                optionRadio.onClick = makeRadioClickHandler(rowControl, option, rows, tableContext);
                rowControl.radios.push(optionRadio);
            }
            tableContext.rowControls[rows[i].key] = rowControl;
        }
    }

    /**
     * 組み方向のタブを足し、文字種・句読点・括弧・記号のパネルを3カラムで並べる。
     * 縦組みのタブには［縦組み用］、横組みのタブには［欧文の,.:のあとにスペースを入れる］を置く
     * @param {TabbedPanel} tabbedPanel - タブを足す先
     * @param {string} tabKey - "horizontal" または "vertical"
     * @param {Function} onRowChange - 選択肢を変えたときの処理（引数はこのタブのコントロール）
     * @returns {{tab: Tab, tabKey: string, rowControls: Object, extraControls: Object, rowLabels: StaticText[], headerSpacers: StaticText[], onRowChange: Function}}
     *     タブと、そのタブの行のコントロール・行のほかの設定のコントロール・幅を揃える項目名と見出しの空きラベル
     */
    function addSettingsTab(tabbedPanel, tabKey, onRowChange) {
        var settingsTab = tabbedPanel.add("tab", undefined, getLabel(LABELS.tab[tabKey]));
        setupTab(settingsTab);
        var isVerticalTab = (tabKey === "vertical");
        var tabControls = {
            tab: settingsTab,
            tabKey: tabKey,
            rowControls: {},
            extraControls: {},
            rowLabels: [],
            headerSpacers: [],
            onRowChange: function () { onRowChange(tabControls); }
        };

        /* 3カラム（左：文字種・句読点、中：括弧、右：記号）。縦に積むと画面に収まらない / Three columns (left: character types and punctuation; middle: brackets; right: symbols); stacked, the dialog does not fit the screen */
        var columnsGroup = settingsTab.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["left", "fill"];
        columnsGroup.spacing = COLUMN_SPACING;

        var leftColumnGroup = columnsGroup.add("group");
        leftColumnGroup.orientation = "column";
        leftColumnGroup.alignChildren = ["fill", "top"];
        leftColumnGroup.spacing = PANEL_SPACING;

        var characterTypePanel = leftColumnGroup.add("panel", undefined, getLabel(LABELS.panel.characterType));
        setupPanel(characterTypePanel);
        addOptionTable(characterTypePanel, CHARACTER_TYPE_ROWS, tabControls);
        var characterTypeExtrasGroup = characterTypePanel.add("group");
        characterTypeExtrasGroup.orientation = "column";
        characterTypeExtrasGroup.alignChildren = ["left", "top"];
        characterTypeExtrasGroup.spacing = 6;
        addExtraCheckbox(characterTypeExtrasGroup, "digitsByLength", LABELS.checkbox.digitsByLength, LABELS.tooltip.digitsByLength, tabControls);
        addExtraCheckbox(characterTypeExtrasGroup, "spaceBetweenJapaneseOnly", LABELS.checkbox.spaceBetweenJapaneseOnly,
            LABELS.tooltip.spaceBetweenJapaneseOnly, tabControls);
        addExtraCheckbox(characterTypeExtrasGroup, "japaneseLatinSpace", LABELS.checkbox.japaneseLatinSpace, LABELS.tooltip.japaneseLatinSpace, tabControls);
        if (!isVerticalTab) {
            addExtraCheckbox(characterTypeExtrasGroup, "westernPunctuationSpace", LABELS.checkbox.westernPunctuationSpace, LABELS.tooltip.westernPunctuationSpace, tabControls);
        }

        var punctuationPanel = leftColumnGroup.add("panel", undefined, getLabel(LABELS.panel.punctuation));
        setupPanel(punctuationPanel);
        addOptionTable(punctuationPanel, PUNCTUATION_ROWS, tabControls);
        addExtraDropdown(punctuationPanel, "punctuationStyle", LABELS.fieldLabel.punctuationStyle,
            [LABELS.dropdown.punctuationStyleKeep, LABELS.dropdown.punctuationStyleJapanese, LABELS.dropdown.punctuationStyleWestern],
            LABELS.tooltip.punctuationStyle, tabControls);
        /* 三点リーダーにする文字（「...に変換 : ☐・・・ ☐... ☐。。。」の1行）/ Characters turned into an ellipsis, on one row */
        var ellipsisRowGroup = punctuationPanel.add("group");
        setupRow(ellipsisRowGroup, "left", 6);
        ellipsisRowGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.toEllipsis));
        addExtraCheckbox(ellipsisRowGroup, "nakaguroToEllipsis", LABELS.checkbox.nakaguroToEllipsis, LABELS.tooltip.nakaguroToEllipsis, tabControls);
        addExtraCheckbox(ellipsisRowGroup, "periodsToEllipsis", LABELS.checkbox.periodsToEllipsis, LABELS.tooltip.periodsToEllipsis, tabControls);
        addExtraCheckbox(ellipsisRowGroup, "kutenToEllipsis", LABELS.checkbox.kutenToEllipsis, LABELS.tooltip.kutenToEllipsis, tabControls);

        var middleColumnGroup = columnsGroup.add("group");
        middleColumnGroup.orientation = "column";
        middleColumnGroup.alignChildren = ["fill", "top"];
        middleColumnGroup.spacing = PANEL_SPACING;

        var bracketPanel = middleColumnGroup.add("panel", undefined, getLabel(LABELS.panel.bracket));
        setupPanel(bracketPanel);
        addOptionTable(bracketPanel, BRACKET_ROWS, tabControls);
        var bracketExtrasGroup = bracketPanel.add("group");
        bracketExtrasGroup.orientation = "column";
        bracketExtrasGroup.alignChildren = ["left", "top"];
        bracketExtrasGroup.spacing = 6;
        addExtraCheckbox(bracketExtrasGroup, "bracketInnerSpace", LABELS.checkbox.bracketInnerSpace, LABELS.tooltip.bracketInnerSpace, tabControls);
        addExtraCheckbox(bracketExtrasGroup, "bracketOuterSpace", LABELS.checkbox.bracketOuterSpace, LABELS.tooltip.bracketOuterSpace, tabControls);
        addExtraDropdown(bracketPanel, "quoteForm", LABELS.fieldLabel.quoteForm,
            [LABELS.dropdown.quoteFormPlain, LABELS.dropdown.quoteFormCurly, LABELS.dropdown.quoteFormMinute], LABELS.tooltip.quoteForm, tabControls);

        /* 縦組み用（縦組みのタブだけ）/ Vertical text (vertical tab only) */
        if (isVerticalTab) {
            var verticalPanel = middleColumnGroup.add("panel", undefined, getLabel(LABELS.panel.vertical));
            setupPanel(verticalPanel, 6);
            addExtraCheckbox(verticalPanel, "longDigitsFullwidth", LABELS.checkbox.longDigitsFullwidth, LABELS.tooltip.longDigitsFullwidth, tabControls);
            addExtraCheckbox(verticalPanel, "numberCommaToTouten", LABELS.checkbox.numberCommaToTouten, LABELS.tooltip.numberCommaToTouten, tabControls);
            addExtraCheckbox(verticalPanel, "decimalToNakaguro", LABELS.checkbox.decimalToNakaguro, LABELS.tooltip.decimalToNakaguro, tabControls);
            addExtraCheckbox(verticalPanel, "shortCapsFullwidth", LABELS.checkbox.shortCapsFullwidth, LABELS.tooltip.shortCapsFullwidth, tabControls);
        }

        var rightColumnGroup = columnsGroup.add("group");
        rightColumnGroup.orientation = "column";
        rightColumnGroup.alignChildren = ["fill", "top"];
        rightColumnGroup.spacing = PANEL_SPACING;

        var symbolPanel = rightColumnGroup.add("panel", undefined, getLabel(LABELS.panel.symbol));
        setupPanel(symbolPanel);
        addOptionTable(symbolPanel, SYMBOL_ROWS, tabControls);
        addExtraDropdown(symbolPanel, "hyphenForm", LABELS.fieldLabel.hyphenForm,
            [LABELS.dropdown.hyphenFormPlain, LABELS.dropdown.hyphenFormHyphen, LABELS.dropdown.hyphenFormMinus], LABELS.tooltip.hyphenForm, tabControls);
        addExtraDropdown(symbolPanel, "asteriskForm", LABELS.fieldLabel.asteriskForm,
            [LABELS.dropdown.asteriskFormPlain, LABELS.dropdown.asteriskFormReference], LABELS.tooltip.asteriskForm, tabControls);
        addExtraDropdown(symbolPanel, "tildeForm", LABELS.fieldLabel.tildeForm,
            [LABELS.dropdown.tildeFormPlain, LABELS.dropdown.tildeFormWave], LABELS.tooltip.tildeForm, tabControls);

        return tabControls;
    }

    /**
     * 設定（行の選択肢と行のほかの設定）をタブにまとめて反映する（タブに無い設定は飛ばす）
     * @param {{rowControls: Object, extraControls: Object}} tabControls - タブのコントロール
     * @param {Object} rowOptions - 設定のキー → 値
     * @returns {void}
     */
    function applyRowOptions(tabControls, rowOptions) {
        for (var settingKey in rowOptions) {
            if (!rowOptions.hasOwnProperty(settingKey)) continue;
            if (tabControls.rowControls[settingKey]) setRowOption(tabControls.rowControls[settingKey], rowOptions[settingKey]);
            else if (tabControls.extraControls[settingKey]) tabControls.extraControls[settingKey].setValue(rowOptions[settingKey]);
        }
        updateDependentControls(tabControls);
    }

    /**
     * タブの今の設定を返す（タブに無い設定は 0）
     * @param {{rowControls: Object, extraControls: Object}} tabControls - タブのコントロール
     * @returns {Object} 設定のキー → 値（すべてのキー）
     */
    function readRowOptions(tabControls) {
        var currentOptions = rowOptionsWith(OPTION_KEEP);
        for (var rowKey in tabControls.rowControls) {
            if (tabControls.rowControls.hasOwnProperty(rowKey)) currentOptions[rowKey] = getRowOption(tabControls.rowControls[rowKey]);
        }
        for (var extraKey in tabControls.extraControls) {
            if (tabControls.extraControls.hasOwnProperty(extraKey)) currentOptions[extraKey] = tabControls.extraControls[extraKey].value;
        }
        return currentOptions;
    }

    /**
     * ほかの設定で使えるかが決まるコントロールを更新する。
     * 数字を桁数で決めるときは数字の行のラジオを、行が「そのまま」のときは全角の字形のドロップダウンを使えなくする
     * （半角にするときも、選んだ字形だけを戻すので使う。引用符は二重引用符・引用符の両方が「そのまま」なら）
     * @param {{rowControls: Object, extraControls: Object}} tabControls - タブのコントロール
     * @returns {void}
     */
    function updateDependentControls(tabControls) {
        var rowControls = tabControls.rowControls;
        var digitsByLengthControl = tabControls.extraControls.digitsByLength;
        var isByLength = !!digitsByLengthControl && (digitsByLengthControl.value === 1);
        for (var i = 0; i < rowControls.digits.radios.length; i++) rowControls.digits.radios[i].enabled = !isByLength;

        var isUsedBySetting = {};
        for (var rowKey in FULL_FORM_CHOICES) {
            if (!FULL_FORM_CHOICES.hasOwnProperty(rowKey)) continue;
            var formSettingKey = FULL_FORM_CHOICES[rowKey].settingKey;
            isUsedBySetting[formSettingKey] = isUsedBySetting[formSettingKey] || (getRowOption(rowControls[rowKey]) !== OPTION_KEEP);
        }
        for (var settingKey in isUsedBySetting) {
            if (isUsedBySetting.hasOwnProperty(settingKey)) tabControls.extraControls[settingKey].setEnabled(isUsedBySetting[settingKey]);
        }
    }

    /**
     * 変換の設定ダイアログを表示する
     * @param {{rowOptionsByTab: Object, activeTab: string, scope: string|null, highlightChanges: boolean}} settings - 前回の設定
     * @param {{selection: boolean}} scopeAvailability - 選べる対象
     * @returns {{rowOptionsByTab: Object, activeTab: string, scope: string, highlightChanges: boolean}|null} 設定。キャンセルなら null
     */
    function showConvertDialog(settings, scopeAvailability) {
        var mainDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(mainDialog);

        /* 上の行：左に対象と確認用の色分け、右にプリセット（間はスペーサーで空ける）。パネルの高さは揃える /
           Top row: scope and check coloring on the left, presets on the right with a spacer between; the panels share the height */
        var topRowGroup = mainDialog.add("group");
        setupRow(topRowGroup, "fill", COLUMN_SPACING);
        topRowGroup.alignChildren = ["left", "fill"];

        /* 対象（ラジオは横並び）/ Scope (radios in a row) */
        var targetPanel = topRowGroup.add("panel", undefined, getLabel(LABELS.panel.target));
        setupPanel(targetPanel);
        targetPanel.orientation = "row";
        targetPanel.alignment = ["left", "fill"];
        targetPanel.alignChildren = ["left", "center"];

        var scopeRadios = {
            selection: targetPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetSelection)),
            document: targetPanel.add("radiobutton", undefined, getLabel(LABELS.radio.targetDocument))
        };
        scopeRadios.selection.helpTip = getLabel(LABELS.tooltip.targetSelection);
        scopeRadios.document.helpTip = getLabel(LABELS.tooltip.targetDocument);
        scopeRadios.selection.enabled = scopeAvailability.selection;

        /* 前回の対象が選べなければ、選択範囲 → ドキュメントの順に選ぶ / Fall back to selection, then document when the last scope is unavailable */
        var initialScope = settings.scope;
        if (!initialScope || !scopeRadios[initialScope] || !scopeRadios[initialScope].enabled) {
            initialScope = scopeAvailability.selection ? "selection" : "document";
        }
        scopeRadios[initialScope].value = true;

        /* 確認用の色分け（どちらのタブでも使う。表示前の value は読み戻せないので値は自前で持つ）/
           Check coloring (used with either tab; the value is kept here, since value cannot be read back before show) */
        var highlightPanel = topRowGroup.add("panel", undefined, getLabel(LABELS.panel.highlight));
        setupPanel(highlightPanel);
        highlightPanel.orientation = "row";
        highlightPanel.alignment = ["left", "fill"];
        highlightPanel.alignChildren = ["left", "center"];
        var highlightCheckbox = highlightPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.highlightChanges));
        highlightCheckbox.helpTip = getLabel(LABELS.tooltip.highlightChanges);
        var highlightChanges = (settings.highlightChanges === true);
        highlightCheckbox.value = highlightChanges;
        highlightCheckbox.onClick = function () { highlightChanges = highlightCheckbox.value; };

        /* 余りの幅を吸って、プリセットを右端に寄せる / Absorbs the extra width so the presets sit at the right edge */
        var topRowSpacer = topRowGroup.add("group");
        topRowSpacer.alignment = ["fill", "fill"];
        topRowSpacer.minimumSize.width = 0;

        /* プリセット（タイトルなしのパネル。開いているタブに反映する）/ Presets (untitled panel; applied to the open tab) */
        var presetPanel = topRowGroup.add("panel");
        setupPanel(presetPanel);
        presetPanel.orientation = "row";
        presetPanel.alignment = ["right", "fill"];
        presetPanel.alignChildren = ["left", "center"];
        var presetRowGroup = presetPanel.add("group");
        setupRow(presetRowGroup, "left", 6);
        presetRowGroup.add("statictext", undefined, labelText(LABELS.fieldLabel.preset));
        var presetDropdown = presetRowGroup.add("dropdownlist");
        presetDropdown.preferredSize.width = PRESET_DROPDOWN_WIDTH;
        presetDropdown.helpTip = getLabel(LABELS.tooltip.preset);
        var btnSavePreset = presetRowGroup.add("button", undefined, getLabel(LABELS.button.savePreset));
        btnSavePreset.helpTip = getLabel(LABELS.tooltip.savePreset);
        var btnDeletePreset = presetRowGroup.add("button", undefined, getLabel(LABELS.button.deletePreset));
        btnDeletePreset.helpTip = getLabel(LABELS.tooltip.deletePreset);

        /* 横組み・縦組みのタブ（設定はタブごと。変換には開いているタブの設定を使う）/ Horizontal and vertical tabs (each keeps its own settings; the open tab's are used) */
        var tabbedPanel = mainDialog.add("tabbedpanel");
        tabbedPanel.alignChildren = ["fill", "top"];
        var tabControlsByKey = {};
        var rowLabels = [];
        var headerSpacers = [];
        for (var tabIndex = 0; tabIndex < TAB_KEYS.length; tabIndex++) {
            var tabControls = addSettingsTab(tabbedPanel, TAB_KEYS[tabIndex], function (changedTabControls) {
                updateDependentControls(changedTabControls);
                syncPresetSelection();
            });
            tabControlsByKey[TAB_KEYS[tabIndex]] = tabControls;
            rowLabels = rowLabels.concat(tabControls.rowLabels);
            headerSpacers = headerSpacers.concat(tabControls.headerSpacers);
        }

        /* 項目名の幅をすべての表で同じにし、見出しの空きも同じ幅にする / Give every table the same label width, and the header gap too */
        var rowLabelWidth = alignLabelWidths(rowLabels, "right");
        for (var spacerIndex = 0; spacerIndex < headerSpacers.length; spacerIndex++) {
            headerSpacers[spacerIndex].preferredSize.width = rowLabelWidth;
        }

        /**
         * 開いているタブのコントロールを返す
         * @returns {Object} タブのコントロール（addSettingsTab() の戻り値）
         */
        function getActiveTabControls() {
            return (tabbedPanel.selection === tabControlsByKey.vertical.tab) ? tabControlsByKey.vertical : tabControlsByKey.horizontal;
        }

        /* ドロップダウンの項目番号 → プリセット（「カスタム」と区切り線は null）/ Dropdown item index → preset ("Custom" and separators are null) */
        var userPresets = loadUserPresets();
        var presetItems = [];
        var isUpdatingPresetDropdown = false;

        /**
         * プリセットのドロップダウンを作り直す（組み込み → 区切り線 → ユーザーのプリセット。どちらのタブでもすべて並べる）
         * @returns {void}
         */
        function rebuildPresetDropdown() {
            isUpdatingPresetDropdown = true;
            presetDropdown.removeAll();
            presetItems = [];
            presetDropdown.add("item", getLabel(LABELS.dropdown.customPreset));
            presetItems.push(null);
            for (var i = 0; i < BUILTIN_PRESETS.length; i++) {
                presetDropdown.add("item", BUILTIN_PRESETS[i].name);
                presetItems.push(BUILTIN_PRESETS[i]);
            }
            if (userPresets.length > 0) {
                presetDropdown.add("separator");
                presetItems.push(null);
            }
            for (var j = 0; j < userPresets.length; j++) {
                presetDropdown.add("item", userPresets[j].name);
                presetItems.push(userPresets[j]);
            }
            isUpdatingPresetDropdown = false;
            syncPresetSelection();
        }

        /**
         * 開いているタブの設定に合うプリセットをドロップダウンで選ぶ（合うものが無ければ「カスタム」）。ほかのタブのプリセットは合わせない
         * @returns {void}
         */
        function syncPresetSelection() {
            var activeTabControls = getActiveTabControls();
            var currentOptions = readRowOptions(activeTabControls);
            var matchedIndex = 0;
            for (var i = 0; i < presetItems.length; i++) {
                if (presetItems[i] && isPresetForTab(presetItems[i], activeTabControls.tabKey) && isSameRowOptions(rowOptionsForTab(presetItems[i].rowOptions, activeTabControls.tabKey), currentOptions)) {
                    matchedIndex = i;
                    break;
                }
            }
            isUpdatingPresetDropdown = true;
            presetDropdown.selection = matchedIndex;
            isUpdatingPresetDropdown = false;
            btnDeletePreset.enabled = !!(presetItems[matchedIndex] && !presetItems[matchedIndex].builtIn);
        }

        /**
         * 名前が同じユーザーのプリセットの番号を返す
         * @param {string} presetName - プリセット名
         * @returns {number} 番号。無ければ -1
         */
        function findUserPresetIndex(presetName) {
            for (var i = 0; i < userPresets.length; i++) {
                if (userPresets[i].name === presetName) return i;
            }
            return -1;
        }

        /**
         * ユーザーのプリセットをファイルに書き、ドロップダウンを作り直す
         * @returns {void}
         */
        function storeUserPresets() {
            if (!saveUserPresets(userPresets)) alert(getLabel(LABELS.alert.presetSaveFailed));
            rebuildPresetDropdown();
        }

        presetDropdown.onChange = function () {
            if (isUpdatingPresetDropdown || !presetDropdown.selection) return;
            var chosenPreset = presetItems[presetDropdown.selection.index];
            btnDeletePreset.enabled = !!(chosenPreset && !chosenPreset.builtIn);
            if (!chosenPreset) return;
            /* タブを持つプリセット（「縦組み用」など）は、そのタブに切り替えて反映する / A preset with a tab (such as vertical) switches to that tab */
            var targetTabControls = chosenPreset.tab ? tabControlsByKey[chosenPreset.tab] : getActiveTabControls();
            applyRowOptions(targetTabControls, chosenPreset.rowOptions);
            if (tabbedPanel.selection !== targetTabControls.tab) {
                tabbedPanel.selection = targetTabControls.tab;
                syncPresetSelection();
            }
        };

        btnSavePreset.onClick = function () {
            var chosenPreset = presetDropdown.selection ? presetItems[presetDropdown.selection.index] : null;
            var suggestedName = (chosenPreset && !chosenPreset.builtIn) ? chosenPreset.name : "";
            var presetName = prompt(getLabel(LABELS.prompt.presetName), suggestedName);
            if (presetName === null) return;
            presetName = String(presetName).replace(/^\s+|\s+$/g, "");
            if (presetName === "") return;
            for (var i = 0; i < BUILTIN_PRESETS.length; i++) {
                if (BUILTIN_PRESETS[i].name === presetName) {
                    alert(getLabel(LABELS.alert.builtInPresetName, { name: presetName }));
                    return;
                }
            }
            /* 選ぶと保存したタブに切り替える / Choosing it later switches to the tab it was saved from */
            var activeTabControls = getActiveTabControls();
            var newPreset = { name: presetName, builtIn: false, tab: activeTabControls.tabKey, rowOptions: readRowOptions(activeTabControls) };
            var existingIndex = findUserPresetIndex(presetName);
            if (existingIndex >= 0) {
                if (!confirm(getLabel(LABELS.confirm.overwritePreset, { name: presetName }), true)) return;
                userPresets[existingIndex] = newPreset;
            } else {
                userPresets.push(newPreset);
            }
            storeUserPresets();
        };

        btnDeletePreset.onClick = function () {
            var chosenPreset = presetDropdown.selection ? presetItems[presetDropdown.selection.index] : null;
            if (!chosenPreset || chosenPreset.builtIn) return;
            if (!confirm(getLabel(LABELS.confirm.deletePreset, { name: chosenPreset.name }), true)) return;
            userPresets.splice(findUserPresetIndex(chosenPreset.name), 1);
            storeUserPresets();
        };

        for (var applyIndex = 0; applyIndex < TAB_KEYS.length; applyIndex++) {
            applyRowOptions(tabControlsByKey[TAB_KEYS[applyIndex]], settings.rowOptionsByTab[TAB_KEYS[applyIndex]]);
        }
        tabbedPanel.selection = tabControlsByKey[settings.activeTab].tab;
        rebuildPresetDropdown();
        /* タブを切り替えたら、そのタブの設定に合うプリセットを選び直す / Switching tabs re-selects the preset matching that tab */
        tabbedPanel.onChange = function () { syncPresetSelection(); };

        var buttonRow = addButtonRow(mainDialog);
        var btnRestoreDefaults = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.restoreDefaults));
        btnRestoreDefaults.helpTip = getLabel(LABELS.tooltip.restoreDefaults);
        btnRestoreDefaults.onClick = function () {
            applyRowOptions(getActiveTabControls(), getDefaultRowOptions());
            syncPresetSelection();
        };

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnConvert = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.convert), { name: "ok" });
        btnCancel.onClick = function () { mainDialog.close(0); };
        btnConvert.onClick = function () { mainDialog.close(1); };
        alignRightOnlyButtonRow(buttonRow);

        prepareDialogWindow(mainDialog, SCRIPT_NAME);
        if (mainDialog.show() !== 1) return null;

        var chosenRowOptionsByTab = {};
        for (var readIndex = 0; readIndex < TAB_KEYS.length; readIndex++) {
            chosenRowOptionsByTab[TAB_KEYS[readIndex]] = readRowOptions(tabControlsByKey[TAB_KEYS[readIndex]]);
        }
        var chosenScope = "document";
        for (var scopeKey in scopeRadios) {
            if (scopeRadios.hasOwnProperty(scopeKey) && scopeRadios[scopeKey].value) chosenScope = scopeKey;
        }
        return {
            rowOptionsByTab: chosenRowOptionsByTab,
            activeTab: getActiveTabControls().tabKey,
            scope: chosenScope,
            highlightChanges: highlightChanges
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで設定を選び、対象の文字の全角・半角を変換する
     * @returns {void}
     */
    function main() {
        var selectionTargets = getSelectionTargets();

        var settings = showConvertDialog(loadSettings(), { selection: selectionTargets.length > 0 });
        if (!settings) return;
        saveSettings(settings);

        /* 開いていたタブの設定で変換する（タブに無い設定は 0）/ Convert with the open tab's settings (settings not on that tab are 0) */
        var conversion = buildConversionMap(settings.rowOptionsByTab[settings.activeTab]);
        if (!hasConversionWork(conversion)) {
            alert(getLabel(LABELS.alert.allKeep));
            return;
        }

        conversion.highlightColor = settings.highlightChanges ? createHighlightColor() : null;
        var targets = (settings.scope === "selection") ? selectionTargets : getDocumentTargets();
        var result = convertTargets(targets, conversion);

        if (result.converted === 0 && result.failed === 0) {
            alert(getLabel(LABELS.alert.notFound));
            return;
        }
        var message = getLabel(LABELS.alert.converted, { count: result.converted });
        if (result.failed > 0) message += "\n" + getLabel(LABELS.alert.failed, { count: result.failed });
        var breakdownLines = formatResultBreakdown(result.countsByKey);
        if (breakdownLines.length > 0) message += "\n\n" + breakdownLines.join("\n");
        alert(message);
    }

    /* 内訳に出す、行以外のキー（行のあとに並べる）/ Keys besides the rows shown in the breakdown, listed after the rows */
    var RESULT_EXTRA_KEYS = ["nakaguroToEllipsis", "periodsToEllipsis", "kutenToEllipsis", "japaneseLatinSpace", "punctuationStyle", "longDigitsFullwidth", "shortCapsFullwidth",
        "numberCommaToTouten", "decimalToNakaguro", "westernPunctuationSpace", "bracketInnerSpace", "bracketOuterSpace"];

    /**
     * 変換した数の内訳を、行の並び順で「項目名 : 数」の行にする
     * @param {Object} countsByKey - キー → 変換した数
     * @returns {string[]} 内訳の行
     */
    function formatResultBreakdown(countsByKey) {
        var breakdownLines = [];
        for (var i = 0; i < ALL_ROWS.length; i++) {
            var row = ALL_ROWS[i];
            if (!countsByKey[row.key]) continue;
            var rowLabel = getLabel(LABELS.row[row.key]) + (row.glyph ? "  " + row.glyph : "");
            breakdownLines.push(getLabel(LABELS.alert.breakdownLine, { label: rowLabel, count: countsByKey[row.key] }));
        }
        for (var j = 0; j < RESULT_EXTRA_KEYS.length; j++) {
            var extraKey = RESULT_EXTRA_KEYS[j];
            if (!countsByKey[extraKey]) continue;
            breakdownLines.push(getLabel(LABELS.alert.breakdownLine, { label: getLabel(LABELS.resultLabel[extraKey]), count: countsByKey[extraKey] }));
        }
        return breakdownLines;
    }

    main();

})();
