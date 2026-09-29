#targetengine "TabularizeEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを「表」として解釈し、表組み用の塗りと線（横ケイ／縦ケイ）を生成します。
計算のためにテキストを複製・アウトライン化しますが、元のテキストは編集可能なまま残ります。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/tabularize.md

### Overview

Interprets the selection as a table grid and generates fills and rules, both horizontal and vertical.
Text is duplicated and outlined for measurement only, so the original stays editable.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/tabularize.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "tabularize";                   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/tabularize.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/tabularize.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 設定の保存（再利用パーツ） / Settings store (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
            return textFile.read().replace(/^﻿/, "");
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
        var trimmedText = legacyText.replace(/^﻿/, "").replace(/^\s+|\s+$/g, "");
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // セッション記憶 / Session memory
    // =========================================

    /* Illustrator を終了するまでダイアログの値を覚える / Remember dialog values until Illustrator quits */
    var settingsStore = createSettingsStore(SCRIPT_NAME, "session");

    /* 保存する値の既定値 / Default values of the saved settings */
    var DEFAULT_SETTINGS = {
        presetIndex: 0,
        presetKey: "",
        // Options
        useGutter: true,
        gutterText: "",
        headerRow: true,
        // Fill
        doFill: false,
        zebra: false,
        fillJoinRow: false,
        fillHeaderOnly: false,
        // Lines
        doRule: true,
        vRuleMode: "gapsOnly" /* 'none' | 'gapsOnly' | 'all' */
    };

    (function () {
        // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        // UI の明暗（再利用パーツ） / UI theme (reusable)
        // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

        /**
         * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
         * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
         */
        function isDarkUI() {
            try {
                if (app.preferences && app.preferences.getRealPreference) {
                    return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
                }
                return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
            } catch (e) {
                return false;
            }
        }

        // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
        // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
        // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

        // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
        // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

        // -----------------------------------------
        // ステップボタンの寸法・増減量 / Stepper metrics and steps
        // -----------------------------------------
        var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
        var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
        var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
        var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
        var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
        var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
        var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

        // -----------------------------------------
        // ステップボタンの配色 / Stepper colors
        // -----------------------------------------
        var STEPPER_UI_DARK           = isDarkUI();
        /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
           ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
           UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
           dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
        var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
        var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
        var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
        var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
        var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
        var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
        var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

        // -----------------------------------------
        // 数値欄を作る（外から呼ぶ関数） / Public API
        // -----------------------------------------
        /**
         * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
         * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
         * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
         * @param {Group|Panel} parent - 追加先
         * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
         *     step / min / max / integer（true で整数のみ）/ unit / onStep
         * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
         */
        function addSteppedField(parent, fieldOptions) {
            var fieldRowGroup = parent.add("group");
            fieldRowGroup.orientation = "row";
            fieldRowGroup.alignChildren = ["left", "center"];
            fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

            var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
            if (fieldOptions.labelWidth) {
                fieldLabel.preferredSize.width = fieldOptions.labelWidth;
                fieldLabel.justify = "right";
            }

            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            var stepperInputGroup = fieldRowGroup.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;

            var numberInput;
            var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
            numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
            numberInput.characters = fieldOptions.characters || 6;
            numberInput.fieldLabel = fieldLabel;
            numberInput.stepperGroup = stepperGroup;

            /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
            bindSteppedArrowKeys(numberInput, stepperGroup);

            /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
            numberInput.lastValidText = numberInput.text;
            numberInput.onChange = function () {
                var value = parseFloat(numberInput.text);
                if (isNaN(value)) {
                    numberInput.text = numberInput.lastValidText;
                    return;
                }
                writeSteppedValue(numberInput, value, fieldOptions);
            };
            return numberInput;
        }

        /**
         * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
         * @param {EditText} numberInput - addSteppedField() で作った入力欄
         * @param {boolean} isEnabled - 有効にするなら true
         * @returns {void}
         */
        function setSteppedFieldEnabled(numberInput, isEnabled) {
            numberInput.enabled = isEnabled;
            numberInput.fieldLabel.enabled = isEnabled;
            numberInput.stepperGroup.enabled = isEnabled;
            /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
            for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
                redrawStepperGroup(numberInput.stepperGroup.children[i]);
            }
        }

        /**
         * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
         * @param {Group|Panel} parent - 追加先
         * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
         * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
         * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
         */
        function addStepper(parent, getNumberInput, stepOptions) {
            var stepperGroup = parent.add("group");
            stepperGroup.orientation = "column";
            stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
            stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
            stepperGroup.alignment = ["left", "center"];

            /**
             * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
             * @param {number} direction - 増やすなら 1、減らすなら -1
             * @returns {void}
             */
            function stepBy(direction) {
                var numberInput = getNumberInput();
                if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
                var value = parseFloat(numberInput.text);
                if (isNaN(value)) value = 0;
                writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
                if (stepOptions.onStep) stepOptions.onStep(numberInput);
            }

            /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
            var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
            var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
            makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
            makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
            stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
            return stepperGroup;
        }

        /**
         * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
         * @param {EditText} numberInput - 対象の入力欄
         * @param {Group} stepperGroup - addStepper() で作った∧∨
         * @returns {void}
         */
        function bindSteppedArrowKeys(numberInput, stepperGroup) {
            numberInput.addEventListener("keydown", function (event) {
                if (event.keyName !== "Up" && event.keyName !== "Down") return;
                stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
                event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
            });
        }

        // -----------------------------------------
        // 値の計算 / Value helpers
        // -----------------------------------------
        /**
         * 押された修飾キーに応じて、1回分増減した値を返す
         * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
         * 整数の欄では option を無視して step の倍数へ）
         * @param {number} value - 元の値
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
         * @returns {number} 増減した値（下限・上限は未適用）
         */
        function computeSteppedValue(value, direction, stepOptions) {
            var keyState = ScriptUI.environment.keyboardState;
            if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
            if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
            return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
        }

        /**
         * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
         * @param {number} value - 元の値
         * @param {number} multiple - 倍数の単位（例 10）
         * @param {number} direction - 上げるなら 1、下げるなら -1
         * @returns {number} 移した値
         */
        function snapStepperToNextMultiple(value, multiple, direction) {
            /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
               round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
            var quotient = Math.round(value / multiple * 1e6) / 1e6;
            if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
            return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
        }

        /**
         * 値を下限・上限の範囲に収める
         * @param {number} value - 数値
         * @param {Object} rangeOptions - min / max（どちらも省略可）
         * @returns {number} 範囲に収めた値
         */
        function clampSteppedValue(value, rangeOptions) {
            if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
            if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
            return value;
        }

        /**
         * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
         * @param {EditText} numberInput - 書き込む入力欄
         * @param {number} value - 数値
         * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
         * @returns {void}
         */
        function writeSteppedValue(numberInput, value, valueOptions) {
            numberInput.text = formatSteppedValue(value, valueOptions);
            numberInput.lastValidText = numberInput.text;
        }

        /**
         * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
         * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
         * @param {number} value - 数値
         * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
         * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
         */
        function formatSteppedValue(value, valueOptions) {
            if (valueOptions.integer) value = Math.round(value);
            return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
        }

        /**
         * 小数第2位で丸めた数値を文字列で返す
         * @param {number} value - 数値
         * @returns {string} 表示用の数値文字列
         */
        function formatStepperNumber(value) {
            return String(Math.round(value * 100) / 100);
        }

        // -----------------------------------------
        // ∧∨ボタンの描画 / Drawing
        // -----------------------------------------
        /**
         * 山形（∧／∨）の極小ボタンを作成する。
         * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
         * 継ぎ目に線は引かない
         * @param {Group|Panel} parent - 追加先
         * @param {string} direction - "up" または "down"
         * @param {Function} onClickFn - クリック時の処理
         * @returns {Group} ボタンとして使う group
         */
        function makeStepperChevronButton(parent, direction, onClickFn) {
            var buttonWidth = STEPPER_BUTTON_WIDTH;
            var buttonHeight = STEPPER_BUTTON_HEIGHT;
            var isUp = (direction === "up");
            var chevronBox = parent.add("group");
            chevronBox.margins = 0;
            chevronBox.spacing = 0;
            chevronBox.preferredSize = [buttonWidth, buttonHeight];
            chevronBox.minimumSize = [buttonWidth, buttonHeight];
            chevronBox.maximumSize = [buttonWidth, buttonHeight];
            chevronBox.isPressed = false;
            chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

            chevronBox.onDraw = function () {
                var boxGraphics = chevronBox.graphics;
                /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
                   Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
                var isDimmed = !isStepperEnabledInTree(chevronBox);

                /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
                var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
                boxGraphics.newPath();
                boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
                boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

                drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
                drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
            };

            /**
             * 押下状態を変えて描き直す
             * @param {boolean} isPressed - 押下中なら true
             * @returns {void}
             */
            function repaint(isPressed) {
                if (chevronBox.isPressed === isPressed) return;
                chevronBox.isPressed = isPressed;
                redrawStepperGroup(chevronBox);
            }
            chevronBox.addEventListener("mousedown", function () {
                if (!isStepperEnabledInTree(chevronBox)) return;
                repaint(true);
                if (onClickFn) onClickFn();
            });
            chevronBox.addEventListener("mouseup", function () { repaint(false); });
            /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
            chevronBox.addEventListener("mouseout", function () { repaint(false); });
            return chevronBox;
        }

        /**
         * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
         * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
         * @param {ScriptUIGraphics} boxGraphics - 描画先
         * @param {number} boxWidth - ボタンの幅
         * @param {number} boxHeight - ボタンの高さ
         * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
         * @param {number[]} frameColor - [r, g, b, a]
         * @returns {void}
         */
        function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
            var frameLeft = 0.5;
            var frameRight = boxWidth - 0.5;
            var outerY = isUp ? 0.5 : boxHeight - 0.5;
            var seamY = isUp ? boxHeight : 0;
            var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
            var radius = STEPPER_CORNER_RADIUS;
            var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
            var angle, k;

            boxGraphics.newPath();
            boxGraphics.moveTo(frameLeft, seamY);
            /* 左の角丸 / left corner */
            for (k = 0; k <= arcSteps; k++) {
                angle = (Math.PI / 2) * k / arcSteps;
                boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
            }
            /* 右の角丸 / right corner */
            for (k = 0; k <= arcSteps; k++) {
                angle = (Math.PI / 2) * k / arcSteps;
                boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
            }
            boxGraphics.lineTo(frameRight, seamY);
            boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
        }

        /**
         * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
         * @param {ScriptUIGraphics} boxGraphics - 描画先
         * @param {number} boxWidth - ボタンの幅
         * @param {number} boxHeight - ボタンの高さ
         * @param {boolean} isUp - ∧なら true、∨なら false
         * @param {number[]} chevronColor - [r, g, b, a]
         * @returns {void}
         */
        function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
            var centerX = boxWidth / 2;
            var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
            var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
            var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
            boxGraphics.newPath();
            boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
            boxGraphics.lineTo(centerX, centerY + tipOffsetY);
            boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
            boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
        }

        /**
         * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
         * @param {Object} control - 対象のコントロール
         * @returns {boolean} すべて有効なら true
         */
        function isStepperEnabledInTree(control) {
            for (var node = control; node; node = node.parent) {
                if (!node.enabled) return false;
            }
            return true;
        }

        /**
         * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
         * @param {Object} container - 行・グループ・パネルなど
         * @returns {void}
         */
        function redrawSteppersIn(container) {
            if (!container.children) return;
            for (var i = 0; i < container.children.length; i++) {
                var child = container.children[i];
                if (child.isStepperButton) redrawStepperGroup(child);
                else redrawSteppersIn(child);
            }
        }

        /**
         * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
         * @param {Group} targetGroup - 描き直す group
         * @returns {void}
         */
        function redrawStepperGroup(targetGroup) {
            targetGroup.hide();
            targetGroup.show();
        }

        // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
        // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
        // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

        // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        // ローカライズ（再利用パーツ） / Localization (reusable)
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
            tipPreset: { ja: "保存した組み方を読み込みます。", en: "Loads a saved table style." },
            tipUseGutter: { ja: "列のあいだにアキを入れます。幅は右の欄で指定します。", en: "Adds a gutter between columns. The field on the right sets its width." },
            tipGutter: { ja: "列のあいだのアキの幅です。", en: "Width of the gutter between columns." },
            tipHeaderRow: { ja: "1行目を見出し行として扱います。", en: "Treats the first row as a header." },
            tipFill: { ja: "セルに背景色を敷きます。", en: "Fills the cells with a background color." },
            tipFillJoinRow: { ja: "行内のセルをつなげて、1本の帯として塗ります。", en: "Merges the cells in a row and fills them as one band." },
            tipZebra: { ja: "1行おきに色を変えて縞模様にします。", en: "Alternates the fill row by row." },
            tipFillHeaderOnly: { ja: "見出し行だけを塗ります。", en: "Fills the header row only." },
            tipRule: { ja: "セルの境にケイ線を引きます。", en: "Draws rules between the cells." },
            tipVruleNone: { ja: "縦のケイ線は引きません。", en: "Draws no vertical rules." },
            tipVruleGapsOnly: { ja: "アキのある位置だけに縦のケイ線を引きます。", en: "Draws vertical rules only where there is a gutter." },
            tipVruleAll: { ja: "すべての列の境に縦のケイ線を引きます。", en: "Draws a vertical rule between every column." },
            tipPreview: { ja: "結果を画面で確認します。キャンセルすると元に戻ります。", en: "Shows the result on the canvas. Cancel restores the original state." },
            previewLabel: { ja: "プレビュー", en: "Preview" },
            dialogTitle: {
                ja: "表組み化" + ' ' + SCRIPT_VERSION,
                en: "Tabularize" + ' ' + SCRIPT_VERSION
            },
            vRulePanel: {
                ja: "線",
                en: "Strokes"
            },
            fillPanel: {
                ja: "塗り",
                en: "Fill"
            },
            fillCheck: {
                ja: "塗り",
                en: "Fill"
            },
            fillOptionPanel: {
                ja: "オプション",
                en: "Options"
            },
            zebra: {
                ja: "ゼブラ",
                en: "Zebra"
            },
            fillJoinRow: {
                ja: "行方向に連結",
                en: "Join by row"
            },
            fillHeaderOnly: {
                ja: "ヘッダー行のみ",
                en: "Header only"
            },
            ruleCheck: {
                ja: "線",
                en: "Rules"
            },
            fillNone: {
                ja: "塗りなし",
                en: "No fill"
            },
            fillOnly: {
                ja: "塗りのみ",
                en: "Fill only"
            },
            fillAndRule: {
                ja: "塗りとケイ",
                en: "Fill + rules"
            },
            gutter: {
                ja: "ガター",
                en: "Gutter"
            },
            useGutter: {
                ja: "ガター",
                en: "Gutter"
            },
            headerRow: {
                ja: "1行目をヘッダー行にする",
                en: "Treat first row as header"
            },
            gapsOnly: {
                ja: "列間のみ",
                en: "Gaps only"
            },
            all: {
                ja: "すべて",
                en: "All"
            },
            none: {
                ja: "なし",
                en: "None"
            },
            vRuleLabel: {
                ja: "縦ケイ",
                en: "Vertical"
            },
            cancel: {
                ja: "キャンセル",
                en: "Cancel"
            },
            ok: {
                ja: "OK",
                en: "OK"
            },
            preset: {
                ja: "プリセット",
                en: "Preset"
            },
            alertOpenDoc: {
                ja: "ドキュメントを開いてください。",
                en: "Please open a document."
            },
            alertSelectObj: {
                ja: "オブジェクトを選択してください。",
                en: "Please select objects."
            },
            optionPanel: {
                ja: "オプション",
                en: "Options"
            },
            /* ステップボタン / Stepper buttons */
            tooltip: {
                stepUp: {
                    ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                    en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
                },
                stepDown: {
                    ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                    en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
                },
                stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
                stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
            }
        };

        /* 単位ユーティリティ / Unit utilities */

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

        /* 線幅・余白・既定の間隔は常に mm で持つ / Line weights, padding and default gutter are always in mm */
        var POINTS_PER_MM = UNITS[1].pointsPerUnit;

        /* 設定項目 / Settings */
        var lineWeightMM = 0.1; // 線の太さ (mm) ※細めが良い場合は0.1など
        var paddingMM = 0.0;    // テキストの左右に少し余白を持たせるか (mm)
        // ----------------

        var lineWeightPt = lineWeightMM * POINTS_PER_MM;
        var paddingPt = paddingMM * POINTS_PER_MM;
        var HEADER_LINE_WEIGHT_MM = 0.3; // チェックON時、上から1本目と2本目だけ太くする(mm)
        var headerLineWeightPt = HEADER_LINE_WEIGHT_MM * POINTS_PER_MM;

        // --- Shared document/layer/color references (must be declared to avoid ReferenceError in ExtendScript) ---
        var doc = null;
        var currentSelection = null;
        var baseLayer = null;
        var lineLayer = null;
        var fillLayer = null;

        var blackColor = null;
        var fillGray = null;
        var fillGrayHeader = null;
        var fillGrayHeaderZebra = null;
        var fillGrayZebra = null;

        // ドキュメントチェックと選択チェックをダイアログより前に移動
        // 横罫ガターなど単位周辺の初期値計算はこのまま

        /* ダイアログ位置 / Dialog position */
        var offsetX = 300;
        var offsetY = 0;

        function shiftDialogPosition(dlg, offsetX, offsetY) {
            dlg.onShow = function () {
                try { requestPreviewUpdate(); } catch (e) { }

                try {
                    var currentX = dlg.location[0];
                    var currentY = dlg.location[1];
                    dlg.location = [currentX + offsetX, currentY + offsetY];
                } catch (e) { }
            };
        }

        // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
        // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

        // --- Preview state ---
        var cbPreview = null;
        var isPreviewing = false;
        var previewItems = []; // created items for preview (lines + fills)
        var previewIsCurrent = false;

        function clearPreview() {
            try {
                for (var i = 0; i < previewItems.length; i++) {
                    try { previewItems[i].remove(); } catch (e) { }
                }
            } catch (e) { }
            previewItems = [];
            previewIsCurrent = false;
            app.redraw();
        }

        /* ダイアログ / Dialog */
        var dlg = new Window('dialog', getLabel('dialogTitle'));
        dlg.orientation = 'column';
        dlg.alignChildren = 'left';

        shiftDialogPosition(dlg, offsetX, offsetY);

        /* ダイアログ状態の復元/保存 / Restore & save dialog state */
        function restoreDialogState(controls) {
            var st = settingsStore.load(DEFAULT_SETTINGS);
            // ガード：プリセット変更で手動へ戻す処理を抑制
            isApplyingPreset = true;
            try {
                // プリセット復元：まず安定ID（presetKey）で復元、次に旧方式（表示テキスト）で復元
                var restored = false;
                try {
                    if (st.presetKey && controls.ddPreset.items && controls.ddPreset.items.length) {
                        // 1) presetKey is a stable ID
                        if (PRESET_IDS && PRESET_IDS.length) {
                            for (var pi = 0; pi < controls.ddPreset.items.length; pi++) {
                                if (String(PRESET_IDS[pi]) === String(st.presetKey)) {
                                    controls.ddPreset.selection = pi;
                                    restored = true;
                                    break;
                                }
                            }
                        }
                        // 2) Backward compatibility: presetKey might be the old display text
                        if (!restored) {
                            for (var pj = 0; pj < controls.ddPreset.items.length; pj++) {
                                if (String(controls.ddPreset.items[pj].text) === String(st.presetKey)) {
                                    controls.ddPreset.selection = pj;
                                    restored = true;
                                    break;
                                }
                            }
                        }
                    }
                } catch (e) { restored = false; }
                if (!restored) {
                    if (typeof st.presetIndex === 'number' && controls.ddPreset.items && controls.ddPreset.items.length > st.presetIndex) {
                        controls.ddPreset.selection = st.presetIndex;
                    }
                }

                // Options
                if (typeof st.useGutter === 'boolean') controls.cbUseGutter.value = st.useGutter;
                if (typeof st.gutterText === 'string' && st.gutterText !== "") controls.etHGutter.text = st.gutterText;
                if (typeof st.headerRow === 'boolean') controls.cbHeader.value = st.headerRow;

                // Fill
                if (typeof st.doFill === 'boolean') controls.cbFill.value = st.doFill;
                if (typeof st.zebra === 'boolean') controls.cbZebra.value = st.zebra;
                if (typeof st.fillJoinRow === 'boolean') controls.cbFillJoinRow.value = st.fillJoinRow;
                if (typeof st.fillHeaderOnly === 'boolean') controls.cbFillHeaderOnly.value = st.fillHeaderOnly;

                // Lines
                if (typeof st.doRule === 'boolean') controls.cbRule.value = st.doRule;
                if (st.vRuleMode === 'none') {
                    controls.rbVruleNone.value = true;
                } else if (st.vRuleMode === 'all') {
                    controls.rbVruleAll.value = true;
                } else {
                    controls.rbVruleGapsOnly.value = true;
                }

            } catch (e) {
            } finally {
                isApplyingPreset = false;
            }

            // 依存UIの反映
            try { controls.applyGutterEnabled(); } catch (e) { }
            try { controls.applyFillEnabled(); } catch (e) { }
            try { controls.applyRuleEnabled(); } catch (e) { }
            try { controls.applyFillJoinRow(); } catch (e) { }
            try { controls.applyFillHeaderOnly(); } catch (e) { }
        }

        function saveDialogState(controls) {
            var st = {};

            // プリセット
            st.presetIndex = (controls.ddPreset.selection) ? controls.ddPreset.selection.index : 0;
            // Store stable preset ID instead of display text (safer across renames)
            try {
                var _pi = (controls.ddPreset.selection) ? controls.ddPreset.selection.index : 0;
                st.presetKey = (PRESET_IDS && PRESET_IDS[_pi]) ? String(PRESET_IDS[_pi]) : '';
            } catch (e) {
                st.presetKey = '';
            }
            // Options
            st.useGutter = !!controls.cbUseGutter.value;
            st.gutterText = String(controls.etHGutter.text || "");
            st.headerRow = !!controls.cbHeader.value;

            // Fill
            st.doFill = !!controls.cbFill.value;
            st.zebra = !!controls.cbZebra.value;
            st.fillJoinRow = !!controls.cbFillJoinRow.value;
            st.fillHeaderOnly = !!controls.cbFillHeaderOnly.value;

            // Lines
            st.doRule = !!controls.cbRule.value;
            st.vRuleMode = controls.rbVruleNone.value ? 'none' : (controls.rbVruleAll.value ? 'all' : 'gapsOnly');
            settingsStore.save(st);
        }
        var gPreset = dlg.add('group');
        gPreset.orientation = 'row';
        gPreset.alignChildren = ['left', 'center'];

        // Preset label
        var stPreset = gPreset.add('statictext', undefined, (uiLang === 'ja') ? 'プリセット：' : 'Preset:');

        var ddPreset = gPreset.add('dropdownlist', undefined, [
            (uiLang === 'ja') ? '（手動）' : '(Manual)',
            (uiLang === 'ja') ? '塗りA' : 'Fill A',
            (uiLang === 'ja') ? '塗りB' : 'Fill B',
            (uiLang === 'ja') ? '塗りC' : 'Fill C',
            (uiLang === 'ja') ? '塗り＋線A' : 'Fill+Stroke A',
            (uiLang === 'ja') ? '塗り＋線B' : 'Fill+Stroke B',
            (uiLang === 'ja') ? '線A-1' : 'Stroke A-1',
            (uiLang === 'ja') ? '線A-2' : 'Stroke A-2',
            (uiLang === 'ja') ? '線B-1' : 'Stroke B-1',
            (uiLang === 'ja') ? '線B-2' : 'Stroke B-2',
            (uiLang === 'ja') ? '線C-1' : 'Stroke C-1',
            (uiLang === 'ja') ? '線C-2' : 'Stroke C-2'
        ]);
        ddPreset.helpTip = getLabel('tipPreset');
        ddPreset.selection = 0;
        // 初期状態は手動（＝何もしない）

        // Stable preset IDs aligned with ddPreset items order
        var PRESET_IDS = [
            "manual",
            "fillA",
            "fillB",
            "fillC",
            "fillStrokeA",
            "fillStrokeB",
            "strokeA1",
            "strokeA2",
            "strokeB1",
            "strokeB2",
            "strokeC1",
            "strokeC2"
        ];

        function getSelectedPresetId() {
            try {
                var i = (ddPreset && ddPreset.selection) ? ddPreset.selection.index : 0;
                if (i < 0 || i >= PRESET_IDS.length) return "manual";
                return PRESET_IDS[i] || "manual";
            } catch (e) {
                return "manual";
            }
        }

        // 手動操作が入ったらプリセットを「手動」に戻す / Switch preset to Manual on any manual change
        var isApplyingPreset = false;
        function setPresetManual() {
            if (isApplyingPreset) return;
            if (ddPreset.selection && ddPreset.selection.index !== 0) {
                ddPreset.selection = 0;
            }
        }

        /* オプション / Options */
        var pOpt = dlg.add('panel', undefined, getLabel('optionPanel'));
        pOpt.orientation = 'column';
        pOpt.alignChildren = 'left';
        pOpt.margins = [15, 20, 15, 10];

        /* ガター設定 / Gutter */
        var gGutter = pOpt.add('group');
        gGutter.orientation = 'row';
        gGutter.alignChildren = ['left', 'center'];

        // チェックOFF時はガター=0＆ディム表示 / When OFF, set gutter=0 and dim
        var cbUseGutter = gGutter.add('checkbox', undefined, getLabel('useGutter'));
        cbUseGutter.helpTip = getLabel('tipUseGutter');
        cbUseGutter.value = true;

        var rulerUnit = getUnitInfo('rulerType');
        var rulerFactorPt = rulerUnit.pointsPerUnit;
        var rulerUnitLabel = rulerUnit.label;
        // デフォルトガター：1mm 相当を rulerType に変換 / Default gutter ≈ 1mm in rulerType
        var defaultGutterMm = 1;
        var defaultGutterPt = defaultGutterMm * POINTS_PER_MM;
        var defaultGutterVal = defaultGutterPt / rulerFactorPt;

        // 表示用に丸め（pt/pxは整数、その他は小数1桁）
        if (rulerUnitLabel === 'pt' || rulerUnitLabel === 'px') {
            defaultGutterVal = Math.round(defaultGutterVal);
        } else {
            defaultGutterVal = Math.round(defaultGutterVal * 10) / 10;
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var gGutterField = gGutter.add('group');
        gGutterField.orientation = 'row';
        gGutterField.alignChildren = ['left', 'center'];
        gGutterField.spacing = 0;
        gGutterField.margins = 0;
        var etHGutter;
        var gutterStepper = addStepper(gGutterField, function () { return etHGutter; }, {
            min: 0,
            /* text の代入ではイベントが出ないので、プレビューが追従するよう呼び出す / fire the handlers so the preview follows */
            onStep: function (numberInput) {
                try {
                    if (typeof numberInput.onChanging === 'function') numberInput.onChanging();
                } catch (e) { }
                try {
                    if (typeof numberInput.onChange === 'function') numberInput.onChange();
                } catch (e) { }
            }
        });
        etHGutter = gGutterField.add('edittext', undefined, String(defaultGutterVal)); // rulerType
        etHGutter.characters = 3;
        etHGutter.helpTip = getLabel('tipGutter');
        bindSteppedArrowKeys(etHGutter, gutterStepper); /* ↑↓キーも∧∨と同じ処理で増減 / arrow keys share the stepper's logic */

        var stGutterUnit = gGutter.add('statictext', undefined, rulerUnitLabel);

        // OFF→0固定 & ディム / ON→復帰
        var lastGutterText = etHGutter.text;
        function applyGutterEnabled() {
            if (!cbUseGutter.value) {
                lastGutterText = etHGutter.text;
                etHGutter.text = '0';
                etHGutter.enabled = false;
                gutterStepper.enabled = false;
                stGutterUnit.enabled = false;
                redrawSteppersIn(gutterStepper);
            } else {
                etHGutter.enabled = true;
                gutterStepper.enabled = true;
                stGutterUnit.enabled = true;
                redrawSteppersIn(gutterStepper);
                // 0のまま戻したくない場合は直前値に戻す
                if (etHGutter.text === '0' && lastGutterText && lastGutterText !== '0') {
                    etHGutter.text = lastGutterText;
                }
            }
        }
        cbUseGutter.onClick = applyGutterEnabled;
        applyGutterEnabled();

        // ガター値をmm指定でセット（rulerTypeに変換して入力欄へ）
        function setGutterByMm(mmVal) {
            var val = (mmVal * POINTS_PER_MM) / rulerFactorPt;
            // 表示用に丸め（pt/pxは整数、その他は小数1桁）
            if (rulerUnitLabel === 'pt' || rulerUnitLabel === 'px') {
                val = Math.round(val);
            } else {
                val = Math.round(val * 10) / 10;
            }
            etHGutter.text = String(val);
        }

        /* 1行目をヘッダー行にする / Treat first row as header */
        var cbHeader = pOpt.add('checkbox', undefined, getLabel('headerRow'));
        cbHeader.helpTip = getLabel('tipHeaderRow');
        cbHeader.value = true;

        /* 2カラムレイアウト / Two-column layout */
        var gCols = dlg.add('group');
        gCols.orientation = 'row';
        gCols.alignChildren = ['fill', 'top'];

        var gLeft = gCols.add('group');
        gLeft.orientation = 'column';
        gLeft.alignChildren = 'fill';

        var gRight = gCols.add('group');
        gRight.orientation = 'column';
        gRight.alignChildren = 'fill';

        /* 塗り / Fill */
        var pFill = gLeft.add('panel', undefined, getLabel('fillPanel'));
        pFill.orientation = 'column';
        pFill.alignChildren = 'left';
        pFill.margins = [15, 20, 15, 10];

        var gFill = pFill.add('group');
        gFill.orientation = 'row';
        gFill.alignChildren = ['left', 'center'];

        var cbFill = gFill.add('checkbox', undefined, getLabel('fillCheck'));
        cbFill.helpTip = getLabel('tipFill');

        // デフォルト：塗りOFF
        cbFill.value = false;

        /* 塗りオプション / Fill options */
        var pFillOpt = pFill.add('panel', undefined, getLabel('fillOptionPanel'));
        pFillOpt.orientation = 'column';
        pFillOpt.alignChildren = 'left';
        pFillOpt.margins = [15, 20, 15, 10];

        // 行方向に連結（UI）
        var cbFillJoinRow = pFillOpt.add('checkbox', undefined, getLabel('fillJoinRow'));
        cbFillJoinRow.helpTip = getLabel('tipFillJoinRow');
        cbFillJoinRow.value = false;

        // ゼブラ（UI）
        var cbZebra = pFillOpt.add('checkbox', undefined, getLabel('zebra'));
        cbZebra.helpTip = getLabel('tipZebra');
        cbZebra.value = false;

        // ヘッダー行のみ（UI）
        var cbFillHeaderOnly = pFillOpt.add('checkbox', undefined, getLabel('fillHeaderOnly'));
        cbFillHeaderOnly.helpTip = getLabel('tipFillHeaderOnly');
        cbFillHeaderOnly.value = false;

        // 塗りOFFならゼブラ/行方向に連結/ヘッダー行のみ はディム表示
        function applyFillEnabled() {
            var on = !!cbFill.value;

            pFillOpt.enabled = on;
            cbZebra.enabled = on;
            cbFillJoinRow.enabled = on;
            cbFillHeaderOnly.enabled = on;

            if (!on) {
                cbZebra.value = false;
                cbFillJoinRow.value = false;
                cbFillHeaderOnly.value = false;
            }
        }

        // ON時：ガターを0にして横方向に連結 / When ON: force gutter=0 and join horizontally
        function applyFillJoinRow() {
            if (cbFillJoinRow.value) {
                cbUseGutter.value = false;
                applyGutterEnabled();
            }
        }
        cbFillJoinRow.onClick = applyFillJoinRow;

        // ON時：ガター0 + ヘッダーON + 塗りは1行目のみ
        function applyFillHeaderOnly() {
            if (cbFillHeaderOnly.value) {
                // ガターを0に
                cbUseGutter.value = false;
                applyGutterEnabled();

                // 1行目をヘッダーに
                cbHeader.value = true;

                // 競合回避：ゼブラはOFF（排他）
                cbZebra.value = false;

                // 競合回避：行方向に連結はOFF
                cbFillJoinRow.value = false;
            }
        }
        cbFillHeaderOnly.onClick = applyFillHeaderOnly;

        // 排他制御：ゼブラON時は「ヘッダー行のみ」をOFF
        cbZebra.onClick = function () {
            if (cbZebra.value) {
                cbFillHeaderOnly.value = false;
            }
        };

        cbFill.onClick = function () {
            applyFillEnabled();
            // 塗りOFFになったら状態をリセット
            if (!cbFill.value) {
                cbZebra.value = false;
                cbFillJoinRow.value = false;
                cbFillHeaderOnly.value = false;
            }
        };

        // 初期反映
        applyFillEnabled();

        /* 縦罫 / Vertical rules */
        var pVrule = gRight.add('panel', undefined, getLabel('vRulePanel'));
        pVrule.orientation = 'column';
        pVrule.alignChildren = 'left';
        pVrule.margins = [15, 20, 15, 10];

        // 線（横ケイ＋縦ケイの有効/無効）
        var cbRule = pVrule.add('checkbox', undefined, getLabel('ruleCheck'));
        cbRule.helpTip = getLabel('tipRule');
        cbRule.value = true;

        // 縦ケイ
        var pVkei = pVrule.add('panel', undefined, getLabel('vRuleLabel'));
        pVkei.orientation = 'column';
        pVkei.alignChildren = 'left';
        pVkei.margins = [15, 20, 15, 10];

        var gVrule = pVkei.add('group');
        gVrule.orientation = 'column';
        gVrule.alignChildren = ['left', 'top'];

        var rbVruleNone = gVrule.add('radiobutton', undefined, getLabel('none'));
        rbVruleNone.helpTip = getLabel('tipVruleNone');
        var rbVruleGapsOnly = gVrule.add('radiobutton', undefined, getLabel('gapsOnly'));
        rbVruleGapsOnly.helpTip = getLabel('tipVruleGapsOnly');
        var rbVruleAll = gVrule.add('radiobutton', undefined, getLabel('all'));
        rbVruleAll.helpTip = getLabel('tipVruleAll');

        // デフォルト：列間のみ
        rbVruleGapsOnly.value = true;

        // 「線」OFF時は縦ケイを「なし」にしてディム表示 / If Rules OFF, force vertical rules to None and dim
        function applyRuleEnabled() {
            if (!cbRule.value) {
                rbVruleNone.value = true;
                pVkei.enabled = false;
            } else {
                pVkei.enabled = true;
            }
        }
        cbRule.onClick = applyRuleEnabled;
        applyRuleEnabled();

        // プリセット適用（UIのみ。設定ロジックは後で追加可能）
        function applyPreset() {
            isApplyingPreset = true;
            try {
                var presetId = getSelectedPresetId();
                if (presetId === "manual") return; // 手動

                // 共通：1行目ON
                cbHeader.value = true;
                cbFillJoinRow.value = false;

                // プリセット用：ヘッダー行のみ（現状のプリセットはすべてOFF）
                var presetHeaderOnly = false;
                cbFillHeaderOnly.value = presetHeaderOnly;

                // 1) 塗りA：アイテムをセルとして扱い、各セルに塗りを設定
                if (presetId === "fillA") {
                    cbFill.value = true;
                    applyFillEnabled();

                    // 塗りオプション（明示）
                    cbFillJoinRow.value = false;
                    cbFillHeaderOnly.value = false;
                    cbZebra.value = false; // ←重要：ゼブラは必ずOFF

                    cbRule.value = false;
                    applyRuleEnabled();

                    cbUseGutter.value = true;
                    applyGutterEnabled();
                    setGutterByMm(1);

                    // 線OFFのため縦ケイは「なし」
                    rbVruleNone.value = true;
                    applyRuleEnabled();
                    return;
                }

                // 2) 塗りB (idx==2): 塗り＋線（ガターOFF）＋見出し行のみ塗り（縦ケイ：列間のみ）
                if (presetId === "fillB") {
                    cbFill.value = true;
                    applyFillEnabled();

                    // 塗りオプション（明示）
                    cbFillJoinRow.value = false;
                    cbFillHeaderOnly.value = false;
                    cbZebra.value = true;

                    cbRule.value = false;
                    applyRuleEnabled();

                    cbUseGutter.value = true;
                    applyGutterEnabled();
                    setGutterByMm(1);

                    // 線OFFのため縦ケイは「なし」
                    rbVruleNone.value = true;
                    applyRuleEnabled();
                    return;
                }

                // 3) 塗りC (idx==3): 塗りBの複製（ガターOFF）
                if (presetId === "fillC") {
                    cbFill.value = true;
                    applyFillEnabled();

                    // 塗りオプション（明示）
                    cbFillJoinRow.value = true;
                    cbFillHeaderOnly.value = false;
                    cbZebra.value = true;

                    cbRule.value = false;
                    applyRuleEnabled();

                    cbUseGutter.value = false; // 0扱い（連結）
                    applyGutterEnabled();
                    setGutterByMm(1);

                    // 線OFFのため縦ケイは「なし」
                    rbVruleNone.value = true;
                    applyRuleEnabled();
                    return;
                }

                // 4) 塗り＋線A (idx==4): 塗り＋線（ガターOFF）＋見出し行のみ塗り（縦ケイ：すべて）
                if (presetId === "fillStrokeA") {
                    cbFill.value = true;
                    applyFillEnabled();

                    // 塗りオプション（明示）
                    cbFillJoinRow.value = false;
                    cbZebra.value = false;

                    cbRule.value = true;
                    applyRuleEnabled();

                    cbUseGutter.value = false; // 0扱い
                    applyGutterEnabled();

                    cbHeader.value = true;

                    // 縦ケイ：すべて
                    rbVruleAll.value = true;
                    applyRuleEnabled();

                    // 見出し行のみ
                    cbFillHeaderOnly.value = true;
                    applyFillHeaderOnly();
                    return;
                }

                // 5) 塗り＋線B (idx==5): 塗り＋線（ガターOFF）＋見出し行のみ塗り（縦ケイ：列間のみ）
                if (presetId === "fillStrokeB") {
                    cbFill.value = true;
                    applyFillEnabled();

                    // 塗りオプション（明示）
                    cbFillJoinRow.value = false;
                    cbZebra.value = false;

                    cbRule.value = true;
                    applyRuleEnabled();

                    cbUseGutter.value = false; // 0扱い
                    applyGutterEnabled();

                    cbHeader.value = true;

                    // 縦ケイ：列間のみ
                    rbVruleGapsOnly.value = true;
                    applyRuleEnabled();

                    // 見出し行のみ
                    cbFillHeaderOnly.value = true;
                    applyFillHeaderOnly();
                    return;
                }

                // 6) 線A-1 (idx==6): セルごとにすべての罫線（ガターあり）
                if (presetId === "strokeA1") {
                    cbFill.value = false;
                    applyFillEnabled();
                    cbFillHeaderOnly.value = presetHeaderOnly;
                    cbRule.value = true;

                    cbUseGutter.value = true;
                    applyGutterEnabled();
                    setGutterByMm(2);

                    rbVruleAll.value = true;
                    applyRuleEnabled();
                    return;
                }

                // 7) 線A-2 (idx==7): セルごとにすべての罫線（ガターあり）左右の縦ケイなし
                if (presetId === "strokeA2") {
                    cbFill.value = false;
                    applyFillEnabled();
                    cbFillHeaderOnly.value = presetHeaderOnly;
                    cbRule.value = true;

                    cbUseGutter.value = true;
                    applyGutterEnabled();
                    setGutterByMm(1);

                    rbVruleGapsOnly.value = true;
                    applyRuleEnabled();
                    return;
                }

                // 8) 線B-1 (idx==8): 罫線（ガターあり）左右の縦ケイなし
                if (presetId === "strokeB1") {
                    cbFill.value = false;
                    applyFillEnabled();
                    cbFillHeaderOnly.value = presetHeaderOnly;
                    cbRule.value = true;

                    cbUseGutter.value = false; // 0扱い（連結）
                    applyGutterEnabled();

                    rbVruleGapsOnly.value = true;
                    applyRuleEnabled();
                    return;
                }

                // 9) 線B-2 (idx==9): 罫線（ガターあり）ケイなし
                if (presetId === "strokeB2") {
                    cbFill.value = false;
                    applyFillEnabled();
                    cbFillHeaderOnly.value = presetHeaderOnly;
                    cbRule.value = true;

                    cbUseGutter.value = false; // 0扱い（連結）
                    applyGutterEnabled();

                    rbVruleNone.value = true;
                    applyRuleEnabled();
                    return;
                }

                // 10) 線C-1 (idx==10): すべてのセルに罫線（ガターなし）
                if (presetId === "strokeC1") {
                    cbFill.value = false;
                    applyFillEnabled();
                    cbFillHeaderOnly.value = presetHeaderOnly;

                    cbRule.value = true;

                    cbUseGutter.value = false; // 0扱い（連結）
                    applyGutterEnabled();

                    // ヘッダーOFF
                    cbHeader.value = false;

                    rbVruleAll.value = true;
                    applyRuleEnabled();
                    return;
                }

                // 11) 線C-2 (idx==11): すべてのセルに罫線（ガターなし）＋見出し行
                if (presetId === "strokeC2") {
                    cbFill.value = false;
                    applyFillEnabled();
                    cbFillHeaderOnly.value = presetHeaderOnly;

                    cbRule.value = true;

                    cbUseGutter.value = false; // 0扱い（連結）
                    applyGutterEnabled();

                    // ヘッダーON
                    cbHeader.value = true;

                    rbVruleAll.value = true;
                    applyRuleEnabled();
                    return;
                }
                // NOTE: 今後プリセットでヘッダー行のみをONにする場合は presetHeaderOnly=true にしてから
                // cbFillHeaderOnly.value を反映し、必要なら applyFillHeaderOnly() を呼ぶ。
            } finally {
                isApplyingPreset = false;
                // プリセット適用後にプレビュー更新（return で抜けても finally は必ず通る）
                try { requestPreviewUpdate(); } catch (e) { }
            }
        }

        ddPreset.onChange = function () {
            applyPreset();
        };

        // --- 手動変更検知 / Manual change detection ---
        function hookManual(control) {
            var prev = control.onClick;
            control.onClick = function () {
                if (prev) prev();
                setPresetManual();
                try { requestPreviewUpdate(); } catch (e) { }
            };
        }

        // チェックボックス類
        hookManual(cbFill);
        hookManual(cbZebra);
        hookManual(cbFillJoinRow);
        hookManual(cbFillHeaderOnly);
        hookManual(cbUseGutter);
        hookManual(cbHeader);
        hookManual(cbRule);

        // ラジオボタン
        hookManual(rbVruleNone);
        hookManual(rbVruleGapsOnly);
        hookManual(rbVruleAll);

        // ガター数値の手動変更
        etHGutter.onChanging = function () {
            setPresetManual();
            try { requestPreviewUpdate(); } catch (e) { }
        };

        // セッション状態を復元
        restoreDialogState({
            ddPreset: ddPreset,
            cbUseGutter: cbUseGutter,
            etHGutter: etHGutter,
            cbHeader: cbHeader,
            cbFill: cbFill,
            cbZebra: cbZebra,
            cbFillJoinRow: cbFillJoinRow,
            cbFillHeaderOnly: cbFillHeaderOnly,
            cbRule: cbRule,
            rbVruleNone: rbVruleNone,
            rbVruleGapsOnly: rbVruleGapsOnly,
            rbVruleAll: rbVruleAll,
            applyGutterEnabled: applyGutterEnabled,
            applyFillEnabled: applyFillEnabled,
            applyRuleEnabled: applyRuleEnabled,
            applyFillJoinRow: applyFillJoinRow,
            applyFillHeaderOnly: applyFillHeaderOnly
        });

        /* ボタン行（左：プレビュー、右：キャンセル・OK） / Button row (left: Preview, right: Cancel and OK) */
        var buttonRow = addButtonRow(dlg);

        cbPreview = buttonRow.leftGroup.add('checkbox', undefined, getLabel('previewLabel'));
        cbPreview.helpTip = getLabel('tipPreview');
        cbPreview.value = true;

        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('ok'), { name: 'ok' });

        btnCancel.onClick = function () {
            saveDialogState({
                ddPreset: ddPreset,
                cbUseGutter: cbUseGutter,
                etHGutter: etHGutter,
                cbHeader: cbHeader,
                cbFill: cbFill,
                cbZebra: cbZebra,
                cbFillJoinRow: cbFillJoinRow,
                cbFillHeaderOnly: cbFillHeaderOnly,
                cbRule: cbRule,
                rbVruleNone: rbVruleNone,
                rbVruleGapsOnly: rbVruleGapsOnly,
                rbVruleAll: rbVruleAll
            });
            try { clearPreview(); } catch (e) { }
            dlg.close(0);
        };
        btnOK.onClick = function () {
            saveDialogState({
                ddPreset: ddPreset,
                cbUseGutter: cbUseGutter,
                etHGutter: etHGutter,
                cbHeader: cbHeader,
                cbFill: cbFill,
                cbZebra: cbZebra,
                cbFillJoinRow: cbFillJoinRow,
                cbFillHeaderOnly: cbFillHeaderOnly,
                cbRule: cbRule,
                rbVruleNone: rbVruleNone,
                rbVruleGapsOnly: rbVruleGapsOnly,
                rbVruleAll: rbVruleAll
            });
            dlg.close(1);
        };
        dlg.onClose = function () {
            try { clearPreview(); } catch (e) { }
            try {
                saveDialogState({
                    ddPreset: ddPreset,
                    cbUseGutter: cbUseGutter,
                    etHGutter: etHGutter,
                    cbHeader: cbHeader,
                    cbFill: cbFill,
                    cbZebra: cbZebra,
                    cbFillJoinRow: cbFillJoinRow,
                    cbFillHeaderOnly: cbFillHeaderOnly,
                    cbRule: cbRule,
                    rbVruleNone: rbVruleNone,
                    rbVruleGapsOnly: rbVruleGapsOnly,
                    rbVruleAll: rbVruleAll
                });
            } catch (e) { }
        };

        function requestPreviewUpdate() {
            try {
                if (cbPreview && cbPreview.value) {
                    applyPreviewNow();
                } else {
                    clearPreview();
                }
            } catch (e) { }
        }

        // Helper to ensure doc/currentSelection/layers/colors for both preview and final
        function ensureDocSelAndLayers() {
            try {
                if (app.documents.length === 0) return false;
                doc = app.activeDocument;
                currentSelection = doc.selection;
                baseLayer = doc.activeLayer;
                if (!currentSelection || currentSelection.length === 0) return false;

                // Layers
                try {
                    if (!lineLayer) lineLayer = doc.layers.getByName('罫線レイヤー');
                } catch (e1) {
                    try {
                        lineLayer = doc.layers.add();
                        lineLayer.name = '罫線レイヤー';
                    } catch (e) { }
                }
                try {
                    if (!fillLayer) fillLayer = doc.layers.getByName('塗りレイヤー');
                } catch (e2) {
                    try {
                        fillLayer = doc.layers.add();
                        fillLayer.name = '塗りレイヤー';
                    } catch (e) { }
                }
                try { if (lineLayer && baseLayer && lineLayer !== baseLayer) lineLayer.move(baseLayer, ElementPlacement.PLACEAFTER); } catch (e) { }
                try { if (fillLayer && lineLayer && fillLayer !== lineLayer) fillLayer.move(lineLayer, ElementPlacement.PLACEAFTER); } catch (e) { }

                // Colors
                try {
                    if (!blackColor) {
                        blackColor = new CMYKColor();
                        blackColor.cyan = 0; blackColor.magenta = 0; blackColor.yellow = 0; blackColor.black = 100;
                    }
                    if (!fillGray) {
                        fillGray = new CMYKColor();
                        fillGray.cyan = 0; fillGray.magenta = 0; fillGray.yellow = 0; fillGray.black = 15;
                    }
                    if (!fillGrayHeader) {
                        fillGrayHeader = new CMYKColor();
                        fillGrayHeader.cyan = 0; fillGrayHeader.magenta = 0; fillGrayHeader.yellow = 0; fillGrayHeader.black = 40;
                    }
                    if (!fillGrayHeaderZebra) {
                        fillGrayHeaderZebra = new CMYKColor();
                        fillGrayHeaderZebra.cyan = 0; fillGrayHeaderZebra.magenta = 0; fillGrayHeaderZebra.yellow = 0; fillGrayHeaderZebra.black = 50;
                    }
                    if (!fillGrayZebra) {
                        fillGrayZebra = new CMYKColor();
                        fillGrayZebra.cyan = 0; fillGrayZebra.magenta = 0; fillGrayZebra.yellow = 0; fillGrayZebra.black = 30;
                    }
                } catch (e) { }

                return true;
            } catch (e) {
                return false;
            }
        }

        function applyPreviewNow() {
            if (!cbPreview || !cbPreview.value) {
                clearPreview();
                return;
            }

            // Rebuild preview from current dialog state
            clearPreview();

            if (!ensureDocSelAndLayers()) return;

            try {
                isPreviewing = true;
                previewIsCurrent = false;

                // Columns + params from UI
                buildCalcColumnsAndParams();

                // Draw preview
                generateMain();
                app.redraw();

                // Cleanup calc proxies only (outlined duplicates)
                try {
                    var _pc = 0;
                    try { _pc = (calcCleanups && typeof calcCleanups.length === 'number') ? calcCleanups.length : 0; } catch (eLen2) { _pc = 0; }
                    for (var ii = 0; ii < _pc; ii++) {
                        try {
                            var fn2 = null;
                            try { fn2 = calcCleanups[ii]; } catch (eGet2) { fn2 = null; }
                            if (fn2) { try { fn2(); } catch (eRun2) { } }
                        } catch (e) { }
                    }
                } catch (e) { }

                previewIsCurrent = true;
                app.redraw();
            } catch (e) {
                try { clearPreview(); } catch (__) { }
                previewIsCurrent = false;
            } finally {
                isPreviewing = false;
            }
        }

        // ドキュメントチェック（ダイアログより前に移動）
        // Ensure doc/currentSelection/layers/colors
        // Declare variables as globals (remove var to avoid shadowing)
        if (!ensureDocSelAndLayers()) {
            if (app.documents.length === 0) { alert(getLabel('alertOpenDoc')); } else { alert(getLabel('alertSelectObj')); }
            return;
        }
        // --- Calculation proxy layer and outline proxies for geometricBounds ---
        // Build calculation proxies for selection after dialog confirmation
        // After dlg.show() confirmed:
        // 1. Create temp layer
        // 2. For each selection, duplicate & outline as needed
        // 3. Build srcItems, calcItems, columns based on outlined proxies
        // 4. Clean up all proxies/layer at the end

        // Proxy/calc data
        var srcItems = [];
        var calcProxyList = [];
        var calcCleanups = [];
        var calcLayer = null;
        var columns = [];
        var calcItems = [];

        prepareDialogWindow(dlg, SCRIPT_NAME);
        if (dlg.show() !== 1) return;

        // OK: if preview is on and current, keep it as final (no regeneration)
        if (cbPreview && cbPreview.value && previewIsCurrent) {
            previewItems = [];      // keep objects (do not remove on exit)
            previewIsCurrent = false;
            return;
        }

        function buildCalcColumnsAndParams() {
            // Reset containers
            srcItems = [];
            calcProxyList = [];
            calcCleanups = [];
            calcLayer = null;
            columns = [];
            calcItems = [];

            // 1) temp layer
            var calcLayerName = "__TabularizeCalc__";
            try { calcLayer = doc.layers.getByName(calcLayerName); }
            catch (e) { calcLayer = doc.layers.add(); calcLayer.name = calcLayerName; }
            try { calcLayer.zOrder(ZOrderMethod.SENDTOFRONT); } catch (e) { }
            try {
                calcLayer.visible = true;
            } catch (e) {}
            try { calcLayer.locked = false; } catch (e) { }

            // 2) proxies
            var seenSrc = [];
            for (var i = 0; i < currentSelection.length; i++) {
                var it = currentSelection[i];
                var g = getSelectedAncestorGroup(it);
                if (g) it = g;
                if (arrayHasRef(seenSrc, it)) continue;
                seenSrc.push(it);
                srcItems.push(it);

                var proxy = { src: it, calc: it, cleanup: null };

                if (it.typename === "TextFrame") {
                    var dup = it.duplicate(calcLayer, ElementPlacement.PLACEATBEGINNING);
                    var outlined = null;
                    try { outlined = dup.createOutline(); } catch (eO) { outlined = null; }
                    try { if (dup && dup.parent) dup.remove(); } catch (e) { }
                    if (outlined) {
                        try { outlined.opacity = 0; } catch (e) { }
                        proxy.calc = outlined;
                        proxy.cleanup = (function (gg) { return function () { try { gg.remove(); } catch (e) { } }; })(outlined);
                    }
                } else if (it.typename === "GroupItem" && hasAnyTextFrame(it)) {
                    var dupg = it.duplicate(calcLayer, ElementPlacement.PLACEATBEGINNING);
                    try { while (dupg.textFrames.length > 0) { dupg.textFrames[0].createOutline(); } } catch (e) { }
                    try { dupg.opacity = 0; } catch (e) { }
                    proxy.calc = dupg;
                    proxy.cleanup = (function (gg2) { return function () { try { gg2.remove(); } catch (e) { } }; })(dupg);
                }

                calcProxyList.push(proxy);
                if (proxy.cleanup) calcCleanups.push(proxy.cleanup);
            }

            // 3) columns
            calcItems = [];
            for (var ci = 0; ci < calcProxyList.length; ci++) calcItems.push(calcProxyList[ci].calc);
            calcItems.sort(function (a, b) { return a.geometricBounds[0] - b.geometricBounds[0]; });

            columns = [];
            if (calcItems.length > 0) {
                var currentColumn = [calcItems[0]];
                columns.push(currentColumn);
                for (var j = 1; j < calcItems.length; j++) {
                    var item = calcItems[j];
                    var itemLeft = item.geometricBounds[0];
                    var prevColMaxRight = getMaxRightInColumn(currentColumn);
                    if (itemLeft < prevColMaxRight) currentColumn.push(item);
                    else { currentColumn = [item]; columns.push(currentColumn); }
                }
            }

            // 4) params from UI
            isZebra = cbZebra.value;
            isFillJoinRow = cbFillJoinRow.value;
            isFillHeaderOnly = cbFillHeaderOnly.value;
            isHeaderRow = cbHeader.value;

            hGutterVal = cbUseGutter.value ? parseFloat(etHGutter.text) : 0;
            if (isNaN(hGutterVal) || hGutterVal < 0) hGutterVal = 0;
            hGutterPt = hGutterVal * rulerFactorPt;

            vRuleMode = rbVruleNone.value ? 'none' : (rbVruleAll.value ? 'all' : 'gapsOnly');

            doFill = cbFill.value;
            doRule = cbRule.value;

            fillMode = (doFill && doRule) ? 'fillAndRule' : (doFill ? 'fillOnly' : 'none');
            KEEP_GAP_PT = hGutterPt;
        }

        try {
            buildCalcColumnsAndParams();
            generateMain();
        } finally {
            // --- Cleanup calc proxies (outlined duplicates only) ---
            try {
                var _cleanupCount = 0;
                try { _cleanupCount = (calcCleanups && typeof calcCleanups.length === 'number') ? calcCleanups.length : 0; } catch (e) { _cleanupCount = 0; }
                for (var cl = 0; cl < _cleanupCount; cl++) {
                    try {
                        var fn = null;
                        try { fn = calcCleanups[cl]; } catch (e) { fn = null; }
                        if (fn) { try { fn(); } catch (e) { } }
                    } catch (e) { }
                }
            } catch (e) { }
        }

        // 生成処理本体 / Main generation
        function generateMain() {

            // 列ごとの左右端を先に計算しておく（列間Aを求めるため）
            var colBounds = []; // {minX, maxX}
            for (var c0 = 0; c0 < columns.length; c0++) {
                var colItems0 = columns[c0];
                var minX0 = 999999;
                var maxX0 = -999999;
                for (var k0 = 0; k0 < colItems0.length; k0++) {
                    var bb0 = colItems0[k0].geometricBounds; // [left, top, right, bottom]
                    if (bb0[0] < minX0) minX0 = bb0[0];
                    if (bb0[2] > maxX0) maxX0 = bb0[2];
                }
                colBounds.push({ minX: minX0, maxX: maxX0 });
            }

            // --- 行（ロウ）を全体で共通化 ---
            // 高さが異なるテキストが混在しても、同じ行として扱い、全列で同じYに横罫を引く
            var rowTol = 2.0; // 同一行とみなす中心Yの許容値(pt) ※必要に応じて調整

            // 行（ロウ）定義は「最も行数が多い列」を基準にする
            // 例：別列に複数行をまたぐグループ（背の高い要素）があっても、行グリッドが歪まない
            var baseColIdx = 0;
            var maxCount = -1;
            for (var bc = 0; bc < columns.length; bc++) {
                if (columns[bc].length > maxCount) {
                    maxCount = columns[bc].length;
                    baseColIdx = bc;
                }
            }

            // 基準列のアイテムだけで行をクラスタリング
            // 行の高さは「その行で一番高さのあるアイテム」を基準にする（基準列内で）
            var allRows = []; // {centerY, maxH}
            var baseItemsSorted = columns[baseColIdx].slice(0);
            baseItemsSorted.sort(function (a, b) { return b.geometricBounds[1] - a.geometricBounds[1]; });

            for (var ai = 0; ai < baseItemsSorted.length; ai++) {
                var it = baseItemsSorted[ai];
                var bb = it.geometricBounds; // [left, top, right, bottom]
                var top = bb[1];
                var bottom = bb[3];
                var h = top - bottom;
                var center = (top + bottom) / 2;

                var ridx = findRowIndex(allRows, center, rowTol);
                if (ridx === -1) {
                    allRows.push({ centerY: center, maxH: h });
                } else {
                    if (h > allRows[ridx].maxH) allRows[ridx].maxH = h;
                    // centerは軽く追従（安定化）
                    allRows[ridx].centerY = (allRows[ridx].centerY + center) / 2;
                }
            }

            // 上→下に並び替え（centerYで）
            allRows.sort(function (a, b) { return b.centerY - a.centerY; });

            // 罫線Yを確定：各行間の中間（全列共通）
            // 行ボックスは centerY ± (maxH/2) で定義（行内で一番高い要素に合わせる）
            var yListGlobal = [];

            function rowTop(i) {
                return allRows[i].centerY + (allRows[i].maxH / 2);
            }
            function rowBottom(i) {
                return allRows[i].centerY - (allRows[i].maxH / 2);
            }

            if (allRows.length >= 2) {
                // 行間の中間（行区切り）
                for (var r = 0; r < allRows.length - 1; r++) {
                    var upperBottom = rowBottom(r);
                    var lowerTop = rowTop(r + 1);
                    var gap = upperBottom - lowerTop;
                    var mid = upperBottom - (gap / 2);
                    yListGlobal.push(mid);
                }

                // 最上段：次の行間から算出
                var firstGap = rowBottom(0) - rowTop(1);
                yListGlobal.push(rowTop(0) + (firstGap / 2));

                // 最下段：直前の行間から算出
                var last = allRows.length - 1;
                var lastGap = rowBottom(last - 1) - rowTop(last);
                yListGlobal.push(rowBottom(last) - (lastGap / 2));

            } else if (allRows.length === 1) {
                var t0 = rowTop(0);
                var b0 = rowBottom(0);
                var rowH = t0 - b0;
                yListGlobal.push(t0 + (rowH / 2));
                yListGlobal.push(b0 - (rowH / 2));
            }

            // 近いYを統合してから描画（最後の保険）
            var yTol = 0.4;
            var yUniqGlobal = [];
            for (var yiG = 0; yiG < yListGlobal.length; yiG++) {
                addUniqueY(yUniqGlobal, yListGlobal[yiG], yTol);
            }
            yUniqGlobal.sort(function (a, b) { return b - a; });

            var yTopBorder = (yUniqGlobal.length > 0) ? yUniqGlobal[0] : 0;
            var yBottomBorder = (yUniqGlobal.length > 0) ? yUniqGlobal[yUniqGlobal.length - 1] : 0;

            // 列ごとのセル領域（塗り）の左右境界を作る / Build fill boundaries per column
            // 列境界は中央（xMid）を共有し、塗り側のinsetでガター見かけを作る
            var colFillLeft = [];
            var colFillRight = [];
            var halfGap = 0;

            for (var cc2 = 0; cc2 < columns.length; cc2++) {
                // 左境界
                if (cc2 === 0) {
                    // 外側は現行の延長ロジックと整合（padding + 端の伸ばし）
                    var minX_0 = colBounds[cc2].minX;
                    var maxX_0 = colBounds[cc2].maxX;
                    var extL0 = 0;
                    if (columns.length >= 2) {
                        var A0_ = colBounds[1].minX - maxX_0;
                        extL0 = (A0_ - KEEP_GAP_PT) / 2;
                        if (extL0 < 0) extL0 = 0;
                    }
                    colFillLeft[cc2] = minX_0 - paddingPt - extL0;
                } else {
                    var xMidL = (colBounds[cc2 - 1].maxX + colBounds[cc2].minX) / 2;
                    colFillLeft[cc2] = xMidL;
                }

                // 右境界
                if (cc2 === columns.length - 1) {
                    var minX_n = colBounds[cc2].minX;
                    var maxX_n = colBounds[cc2].maxX;
                    var extRn = 0;
                    if (columns.length >= 2) {
                        var An_ = minX_n - colBounds[cc2 - 1].maxX;
                        extRn = (An_ - KEEP_GAP_PT) / 2;
                        if (extRn < 0) extRn = 0;
                    }
                    colFillRight[cc2] = maxX_n + paddingPt + extRn;
                } else {
                    var xMidR = (colBounds[cc2].maxX + colBounds[cc2 + 1].minX) / 2;
                    colFillRight[cc2] = xMidR;
                }
            }

            // --- 塗り / Fill ---
            if (doFill && yUniqGlobal.length >= 2) {
                if (isFillHeaderOnly) {
                    // ヘッダー行のみ：1行目だけ塗る（横方向に連結）
                    var xL0 = colFillLeft[0];
                    var xR0 = colFillRight[columns.length - 1];

                    var yTopH = yUniqGlobal[0];
                    var yBotH = yUniqGlobal[1];

                    var wH = xR0 - xL0;
                    var hH = yTopH - yBotH;
                    if (wH > 0 && hH > 0) {
                        var rectH = fillLayer.pathItems.rectangle(yTopH, xL0, wH, hH);
                        rectH.stroked = false;
                        rectH.filled = true;
                        // 塗り色：ヘッダーを優先し、ゼブラONならK50
                        rectH.fillColor = isZebra ? fillGrayHeaderZebra : fillGrayHeader;
                        if (isPreviewing) { try { previewItems.push(rectH); } catch (e) { } }
                    }

                } else if (isFillJoinRow) {
                    // 行方向に連結：各行1つの矩形（横方向に連結）
                    var xL0 = colFillLeft[0];
                    var xR0 = colFillRight[columns.length - 1];

                    for (var fr = 0; fr < yUniqGlobal.length - 1; fr++) {
                        var yTop = yUniqGlobal[fr];
                        var yBot = yUniqGlobal[fr + 1];

                        var left = xL0;
                        var right = xR0;
                        var top = yTop;
                        var bottom = yBot;

                        var w = right - left;
                        var h = top - bottom;
                        if (w <= 0 || h <= 0) continue;

                        var rect = fillLayer.pathItems.rectangle(top, left, w, h);
                        rect.stroked = false;
                        rect.filled = true;

                        // 塗り色：ヘッダーを優先し、ゼブラONなら奇数行をK30
                        if (isHeaderRow && fr === 0) {
                            rect.fillColor = isZebra ? fillGrayHeaderZebra : fillGrayHeader;
                        } else if (isZebra && ((fr + 1) % 2 === 1)) {
                            rect.fillColor = fillGrayZebra;
                        } else {
                            rect.fillColor = fillGray;
                        }
                        if (isPreviewing) { try { previewItems.push(rect); } catch (e) { } }
                    }

                } else {
                    // 通常：セルごと
                    // 見かけのガターを統一：列間・段間ともに KEEP_GAP_PT/2 にする
                    // 各方向をinsetずつ詰めるので、間隔は 2*inset になる
                    var insetX = KEEP_GAP_PT / 2;
                    var insetY = KEEP_GAP_PT / 2;

                    for (var fc = 0; fc < columns.length; fc++) {
                        var xL = colFillLeft[fc];
                        var xR = colFillRight[fc];

                        for (var fr2 = 0; fr2 < yUniqGlobal.length - 1; fr2++) {
                            var yTop2 = yUniqGlobal[fr2];
                            var yBot2 = yUniqGlobal[fr2 + 1];

                            // 内側へ詰める（負のオフセット相当）
                            var left2 = xL + insetX;
                            var right2 = xR - insetX;
                            var top2 = yTop2 - insetY;
                            var bottom2 = yBot2 + insetY;

                            var w2 = right2 - left2;
                            var h2 = top2 - bottom2;
                            if (w2 <= 0 || h2 <= 0) continue;

                            var rect2 = fillLayer.pathItems.rectangle(top2, left2, w2, h2);
                            rect2.stroked = false;
                            rect2.filled = true;

                            // 塗り色：ヘッダーを優先し、ゼブラONなら奇数行をK30
                            if (isHeaderRow && fr2 === 0) {
                                rect2.fillColor = isZebra ? fillGrayHeaderZebra : fillGrayHeader;
                            } else if (isZebra && ((fr2 + 1) % 2 === 1)) {
                                rect2.fillColor = fillGrayZebra;
                            } else {
                                rect2.fillColor = fillGray;
                            }
                            if (isPreviewing) { try { previewItems.push(rect2); } catch (e) { } }
                        }
                    }
                }
            }

            // --- 横罫（行ごと） ---
            if (doRule) {
                // ガターOFF（=0）なら、横ケイも表全体で連結して1本にする
                if (!cbUseGutter.value || KEEP_GAP_PT === 0) {
                    // 表全体の左右端（外側の伸ばしも考慮）
                    var minXLeft = colBounds[0].minX;
                    var maxXLeft = colBounds[0].maxX;
                    var minXRight = colBounds[columns.length - 1].minX;
                    var maxXRight = colBounds[columns.length - 1].maxX;

                    var extLeft = 0;
                    var extRight = 0;

                    if (columns.length >= 2) {
                        // 左端：1列目と2列目の間隔Aから B=(A-KEEP)/2
                        var A0 = colBounds[1].minX - maxXLeft;
                        extLeft = (A0 - KEEP_GAP_PT) / 2;
                        if (extLeft < 0) extLeft = 0;

                        // 右端：最終-1列目と最終列目の間隔Aから B=(A-KEEP)/2
                        var lastIdx = columns.length - 1;
                        var An = minXRight - colBounds[lastIdx - 1].maxX;
                        extRight = (An - KEEP_GAP_PT) / 2;
                        if (extRight < 0) extRight = 0;
                    }

                    var globalStartX = minXLeft - paddingPt - extLeft;
                    var globalEndX = maxXRight + paddingPt + extRight;

                    // When gutter is 0 and vertical rules are "all",
                    // draw the outer frame as a single rectangle (avoid 4 separate border lines)
                    var useRectBorder = (vRuleMode === 'all');
                    if (useRectBorder) {
                        try {
                            var bw = globalEndX - globalStartX;
                            var bh = yTopBorder - yBottomBorder;
                            if (bw > 0 && bh > 0) {
                                var borderRect = lineLayer.pathItems.rectangle(yTopBorder, globalStartX, bw, bh);
                                borderRect.filled = false;
                                borderRect.stroked = true;
                                borderRect.strokeColor = blackColor;
                                // If header row is ON, use the header line weight for the outer frame
                                borderRect.strokeWidth = isHeaderRow ? headerLineWeightPt : lineWeightPt;
                                try { borderRect.strokeJoin = StrokeJoin.MITERENDJOIN; } catch (e) { }
                                if (isPreviewing) { try { previewItems.push(borderRect); } catch (e) { } }
                            }
                        } catch (e) { }
                    }

                    var hStart = useRectBorder ? 1 : 0;
                    var hEnd = useRectBorder ? (yUniqGlobal.length - 1) : yUniqGlobal.length;
                    for (var iLine = hStart; iLine < hEnd; iLine++) {
                        // ヘッダーON時は「上から1本目と2本目」だけ 0.25mm、それ以外は基本 0.1mm
                        var strokeW = lineWeightPt;
                        if (isHeaderRow && iLine < 2) strokeW = headerLineWeightPt;
                        drawLine(globalStartX, yUniqGlobal[iLine], globalEndX, yUniqGlobal[iLine], strokeW);
                    }

                } else {
                    // ガターON：従来どおり、列ごとに線長を計算して描画
                    for (var c = 0; c < columns.length; c++) {
                        var colItems = columns[c];

                        var minX = colBounds[c].minX;
                        var maxX = colBounds[c].maxX;

                        // A：列間を計算（左/右）
                        // B = (A - 1mm) / 2 （マイナスなら0）
                        var extendLeft = 0;
                        var extendRight = 0;

                        // 内側（隣接列に向かう側）の伸ばし：列間Aから B=(A-1mm)/2 を算出
                        if (c > 0) {
                            var prevMaxX = colBounds[c - 1].maxX;
                            var A_left = minX - prevMaxX;
                            extendLeft = (A_left - KEEP_GAP_PT) / 2;
                            if (extendLeft < 0) extendLeft = 0;
                        }
                        if (c < columns.length - 1) {
                            var nextMinX = colBounds[c + 1].minX;
                            var A_right = nextMinX - maxX;
                            extendRight = (A_right - KEEP_GAP_PT) / 2;
                            if (extendRight < 0) extendRight = 0;
                        }

                        // 外側（端）の伸ばし：
                        // 1列目の左は「1列目と2列目の間隔A」から B=(A-1mm)/2
                        if (c === 0 && columns.length >= 2) {
                            var nextMinX0 = colBounds[1].minX;
                            var A0a = nextMinX0 - maxX;
                            var ext0 = (A0a - KEEP_GAP_PT) / 2;
                            if (ext0 < 0) ext0 = 0;
                            extendLeft = ext0;
                        }

                        // 最終列の右は「最終-1列目と最終列目の間隔A」から B=(A-1mm)/2
                        if (c === columns.length - 1 && columns.length >= 2) {
                            var prevMaxXn = colBounds[columns.length - 2].maxX;
                            var An2 = minX - prevMaxXn;
                            var extn = (An2 - KEEP_GAP_PT) / 2;
                            if (extn < 0) extn = 0;
                            extendRight = extn;
                        }

                        // パディング＋延長を適用
                        var lineStartX = minX - paddingPt - extendLeft;
                        var lineEndX = maxX + paddingPt + extendRight;

                        // 横罫線は全列共通のY（yUniqGlobal）で描画する
                        for (var iLine2 = 0; iLine2 < yUniqGlobal.length; iLine2++) {
                            var strokeW2 = lineWeightPt;
                            if (isHeaderRow && iLine2 < 2) strokeW2 = headerLineWeightPt;
                            drawLine(lineStartX, yUniqGlobal[iLine2], lineEndX, yUniqGlobal[iLine2], strokeW2);
                        }
                    }
                }
            }

            // --- 縦罫（列間のみ / すべて） ---
            if (doRule && vRuleMode !== 'none' && yUniqGlobal && yUniqGlobal.length > 0 && columns.length >= 2) {
                // 縦罫のX位置（列境界）を作る
                var vXs = [];

                // 列間のみ：列間（ガターの中央）
                for (var cc = 0; cc < columns.length - 1; cc++) {
                    var xMid = (colBounds[cc].maxX + colBounds[cc + 1].minX) / 2;
                    vXs.push(xMid);
                }

                // すべて：外枠（左/右）も追加
                if (vRuleMode === 'all') {
                    // 左端は「1列目と2列目の間隔A」から B=(A-KEEP)/2 を算出して外側に広げる
                    var leftMinX = colBounds[0].minX;
                    var leftExtend = 0;
                    {
                        var A0 = colBounds[1].minX - colBounds[0].maxX;
                        leftExtend = (A0 - KEEP_GAP_PT) / 2;
                        if (leftExtend < 0) leftExtend = 0;
                    }
                    var xLeftBorder = leftMinX - paddingPt - leftExtend;

                    // 右端は「最終-1列目と最終列目の間隔A」から B=(A-KEEP)/2
                    var lastIdx = columns.length - 1;
                    var rightMaxX = colBounds[lastIdx].maxX;
                    var rightExtend = 0;
                    {
                        var An = colBounds[lastIdx].minX - colBounds[lastIdx - 1].maxX;
                        rightExtend = (An - KEEP_GAP_PT) / 2;
                        if (rightExtend < 0) rightExtend = 0;
                    }
                    var xRightBorder = rightMaxX + paddingPt + rightExtend;

                    // 先頭/末尾に追加（重複しない順番で）
                    vXs.unshift(xLeftBorder);
                    vXs.push(xRightBorder);
                }

                // If outer border is drawn as a rectangle, remove the left/right border lines from vXs
                if ((!cbUseGutter.value || KEEP_GAP_PT === 0) && vRuleMode === 'all' && vXs.length >= 2) {
                    try { vXs = vXs.slice(1, vXs.length - 1); } catch (e) { }
                }

                // 縦罫を描画
                if (!cbUseGutter.value) {
                    // ガターOFF：縦罫は連結して1本で描画
                    for (var vx = 0; vx < vXs.length; vx++) {
                        drawLine(vXs[vx], yTopBorder, vXs[vx], yBottomBorder, lineWeightPt);
                    }
                } else {
                    // ガターON：セルごとに分割して描画（既存挙動）
                    // 端をvTrimずつ詰める（段間=2*vTrim）
                    var vTrim = KEEP_GAP_PT / 2;
                    for (var vx = 0; vx < vXs.length; vx++) {
                        for (var ry = 0; ry < yUniqGlobal.length - 1; ry++) {
                            var y1 = yUniqGlobal[ry];
                            var y2 = yUniqGlobal[ry + 1];

                            // y1 は上、y2 は下（y1 > y2）の想定
                            var segH = y1 - y2;
                            if (segH <= vTrim * 2) {
                                // 短すぎる場合は無理に詰めない（0長さや逆転防止）
                                drawLine(vXs[vx], y1, vXs[vx], y2, lineWeightPt);
                                continue;
                            }

                            var ys = y1 - vTrim;
                            var ye = y2 + vTrim;
                            drawLine(vXs[vx], ys, vXs[vx], ye, lineWeightPt);
                        }
                    }
                }
            }

        } // end generateMain

        /* ユーティリティ関数 / Utilities */

        /* 選択の正規化 / Normalize selection */

        /* グループを1アイテム扱い / Treat group as a single item */

        // 選択オブジェクトがグループ内にある場合、最上位の親グループを返す（グループ選択の有無は問わない）
        // If the item is inside a group, return the topmost parent group (regardless of group selection state).
        function getSelectedAncestorGroup(item) {
            var p = item;
            var found = null;
            while (p && p.parent && p.parent.typename === 'GroupItem') {
                found = p.parent;
                p = p.parent;
            }
            return found;
        }

        // ExtendScript互換：参照配列にobjが含まれるか（indexOfが無い環境向け）
        function arrayHasRef(arr, obj) {
            for (var i = 0; i < arr.length; i++) {
                if (arr[i] === obj) return true;
            }
            return false;
        }

        // 近いYを同一とみなしてユニークに追加する
        function addUniqueY(list, y, tol) {
            for (var i = 0; i < list.length; i++) {
                if (Math.abs(list[i] - y) <= tol) return;
            }
            list.push(y);
        }

        // rows配列から、centerYが近い行のインデックスを返す（見つからなければ-1）
        function findRowIndex(rows, centerY, tol) {
            for (var i = 0; i < rows.length; i++) {
                if (Math.abs(rows[i].centerY - centerY) <= tol) return i;
            }
            return -1;
        }

        // 列内の最大右端座標を取得する関数
        function getMaxRightInColumn(colAry) {
            var maxR = -999999;
            for (var i = 0; i < colAry.length; i++) {
                var r = colAry[i].geometricBounds[2];
                if (r > maxR) maxR = r;
            }
            return maxR;
        }

        // グループまたは子孫にTextFrameが含まれるか
        function hasAnyTextFrame(item) {
            if (!item) return false;
            if (item.typename === "TextFrame") return true;
            if (item.typename === "GroupItem") {
                if (item.textFrames && item.textFrames.length > 0) return true;
                // Recursively check subgroups
                for (var i = 0; i < item.groupItems.length; i++) {
                    if (hasAnyTextFrame(item.groupItems[i])) return true;
                }
            }
            return false;
        }

        // 線描画関数
        function drawLine(x1, y1, x2, y2, strokeW) {
            var pathItem = lineLayer.pathItems.add();
            pathItem.setEntirePath([[x1, y1], [x2, y2]]);
            pathItem.stroked = true;

            // 端部が重なってガターが黒く見えるのを防ぐ（端部をバットに固定）
            try { pathItem.strokeCap = StrokeCap.BUTTENDCAP; } catch (e) { }
            try { pathItem.strokeJoin = StrokeJoin.MITERENDJOIN; } catch (e) { }

            pathItem.strokeWidth = (typeof strokeW === 'number') ? strokeW : lineWeightPt;
            pathItem.strokeColor = blackColor;
            pathItem.filled = false;
            if (isPreviewing) {
                try { previewItems.push(pathItem); } catch (e) { }
            }
        }

    })();

})();
