#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

指定したフォルダー内の .jsx / .js / .jsxbin をキーワードで絞り込み、選んだスクリプトをその場で実行するランチャーです。
フォルダーとファイル名を左右のリストに分けて表示し、絞り込んだ結果によく出てくる語をワンクリックのボタンとして自動で並べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiScriptLauncher.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n86fe7e6251ec

### Overview

A launcher that filters .jsx / .js / .jsxbin files in a chosen folder by keyword and runs the selected script on the spot.
Folders and file names are shown in two side-by-side lists, and words that appear often in the filtered results become one-click buttons automatically.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiScriptLauncher.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiScriptLauncher";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-11-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiScriptLauncher.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiScriptLauncher.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n86fe7e6251ec"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考 / Reference
 * 「Finderで表示」の仕組み（Automatorアプリとの連携）
 * 自分用メモ (@mute_racoon3631)「真の「Finderで表示」をイラレでも」
 * https://note.com/mute_racoon3631/n/n9e0e08f5d5f7
 */

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 一覧に載せる拡張子 / Script file extensions to list */
    var SCRIPT_EXT_RE = /\.(jsx|js|jsxbin)$/i;

    /* スクリプト名から自動で作るキーワードボタンの条件 / Rules for the auto keyword buttons */
    var KEYWORD_MIN_WORD_LENGTH    = 3;      /* ボタンにする語の最小文字数 / minimum word length */
    var KEYWORD_PRESET_MIN_COUNT   = 4;      /* ボタンにする最小出現ファイル数 / minimum number of files */
    var KEYWORD_PRESET_MAX_BUTTONS = 10;     /* ボタンの最大個数 / maximum number of buttons */
    var KEYWORD_PRESET_LIMIT       = 40;     /* 環境設定で指定できる上限 / ceiling for the preferences dialog */
    var KEYWORD_PRESETS_PER_ROW    = 5;      /* 1行に並べる個数 / buttons per row */

    /* ボタンにしない語。件数は多いが検索の役に立たない接続語 / Words never turned into buttons */
    var KEYWORD_STOP_WORDS = ["and", "the", "for", "with", "from", "into", "その他"];

    /* 出現数に関係なく先頭に並べる語（小文字）。該当するファイルが1件もなければ出さない / Words always shown first */
    var KEYWORD_PINNED_WORDS = ["palette"];

    /* Finder表示に使うAutomatorアプリと、パスを受け渡す一時ファイル / Automator app used to reveal a file */
    /* 一時ファイル名はアプリ内のAppleScriptが読む固定名。アプリを旧名 IllustratorRevealLink.app から
       改名した名残で綴りが揃っていないが、アプリ側を直すまでここは変えない
       / The temp file name is hard-coded in the app; leave it until the app itself is updated */
    var REVEAL_APP_PATH  = "/Applications/RevealInFinder.app";
    var REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";

    /* 「サブディレクトリを含む」の初期状態 / Initial state of the subfolder checkbox */
    var INCLUDE_SUBFOLDERS_DEFAULT = false;

    /* 「フルパス」の初期状態。OFFではホームフォルダーを ~ に略す / Initial state of the full path checkbox */
    var SHOW_FULL_PATH_DEFAULT = false;

    /* 「検索条件を記憶」の初期状態 / Initial state of the remember checkbox */
    var REMEMBER_SETTINGS_DEFAULT = true;

    /* サブディレクトリOFF時に残す階層の深さ（対象フォルダー直下のフォルダーまで） / Folder depth kept when subfolders are excluded */
    var NESTED_FOLDER_DEPTH_LIMIT = 1;

    // =========================================
    // 設定の保存 / Stored settings
    // =========================================

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

    /* 再起動しても残す（Folder.userData/illustrator-scripts/AiScriptLauncher.json）。
       以前は Illustrator の環境設定に保存していたので、新しい保存が無いときだけそこから読み継ぐ
       Kept across restarts; the values formerly stored in Illustrator's preferences are read only until the first save */
    var LEGACY_PREF_KEY_FOLDER      = "AiScriptLauncher.targetFolder";
    var LEGACY_PREF_KEY_MIN_COUNT   = "AiScriptLauncher.keywordMinCount";
    var LEGACY_PREF_KEY_MAX_BUTTONS = "AiScriptLauncher.keywordMaxButtons";

    var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent", {
        legacy: function () {
            var legacyFolder = trimWhitespace(app.preferences.getStringPreference(LEGACY_PREF_KEY_FOLDER));
            /* 未登録のキーは空文字・0 が返る / missing keys come back as "" or 0 */
            var legacyMinCount = app.preferences.getIntegerPreference(LEGACY_PREF_KEY_MIN_COUNT);
            var legacyMaxButtons = app.preferences.getIntegerPreference(LEGACY_PREF_KEY_MAX_BUTTONS);
            if (!legacyFolder && !(legacyMinCount > 0) && !(legacyMaxButtons > 0)) return null;
            return { targetFolder: legacyFolder, keywordMinCount: legacyMinCount, keywordMaxButtons: legacyMaxButtons };
        }
    });
    /* 既定値。0 は「記録なし」（ユーザー設定の初期値を使う）/ Defaults; 0 means "not recorded" */
    var DEFAULT_SETTINGS = { targetFolder: "", keywordMinCount: 0, keywordMaxButtons: 0 };

    /**
     * 前回使った対象フォルダーを読み出す
     * @returns {Folder|null} 実在するフォルダー。記録がなければ null
     */
    function readSavedScriptFolder() {
        var savedPath = trimWhitespace(settingsStore.load(DEFAULT_SETTINGS).targetFolder);
        if (!savedPath) return null;

        var savedFolder = new Folder(savedPath);
        return savedFolder.exists ? savedFolder : null;
    }

    /**
     * 対象フォルダーを次回起動用に記録する（キーワード設定はそのまま）
     * @param {Folder} targetFolder - 記録するフォルダー
     * @returns {void}
     */
    function saveScriptFolder(targetFolder) {
        var storedSettings = settingsStore.load(DEFAULT_SETTINGS);
        storedSettings.targetFolder = targetFolder.fsName;
        settingsStore.save(storedSettings);
    }

    /**
     * 記録済みのキーワード設定を読み出す
     * @returns {{minCount: number, maxButtons: number}} 記録がなければユーザー設定の初期値
     */
    function readKeywordSettings() {
        var storedSettings = settingsStore.load(DEFAULT_SETTINGS);
        var minCount = storedSettings.keywordMinCount;
        var maxButtons = storedSettings.keywordMaxButtons;

        return {
            minCount: (minCount > 0) ? minCount : KEYWORD_PRESET_MIN_COUNT,
            maxButtons: (maxButtons > 0) ? maxButtons : KEYWORD_PRESET_MAX_BUTTONS
        };
    }

    /**
     * キーワード設定を記録する（対象フォルダーはそのまま）
     * @param {number} minCount - 出現数
     * @param {number} maxButtons - キーワード数
     * @returns {void}
     */
    function saveKeywordSettings(minCount, maxButtons) {
        var storedSettings = settingsStore.load(DEFAULT_SETTINGS);
        storedSettings.keywordMinCount = minCount;
        storedSettings.keywordMaxButtons = maxButtons;
        settingsStore.save(storedSettings);
    }

    // =========================================
    // セッションの記憶 / Session state
    // =========================================

    /* メインエンジンはスクリプトを実行するたびに初期化され、グローバル変数は次の実行まで残らない。
       そのため一時ファイルに置き、このセッションを表すidを一緒に書いておく。Illustratorを起動し直すと
       idが変わり、前回の記録は使われない
       / Globals do not survive between runs, so the state lives in a temp file tagged with a session id */
    var SESSION_FILE_NAME = "AiScriptLauncher.session.txt";

    /* Illustratorが起動時に作り、終了時に消す空のファイル。作成時刻をセッションのidに使う
       / Illustrator creates this empty file at launch and removes it on quit */
    var SESSION_LOCK_FILE_MASK = "IllustratorSession.lck*";

    /* 1回の実行の中でidは変わらないので、調べるのは最初の1回だけにする / Look it up once per run */
    var isSessionIdChecked = false;
    var cachedSessionId = null;

    /**
     * 一時フォルダー内のファイルを返す
     * @param {string} fileName - ファイル名
     * @returns {File} 一時フォルダー内のファイル
     */
    function temporaryFile(fileName) {
        return new File(Folder.temp.fsName + "/" + fileName);
    }

    /**
     * フォルダーの中身を取り出す
     * @param {Folder} sourceFolder - 対象のフォルダー
     * @param {string} [nameMask] - 名前のマスク。省略時はすべて
     * @returns {Array} 見つかった項目。読めなければ空の配列
     */
    function listFolderEntries(sourceFolder, nameMask) {
        try {
            /* getFiles は読めないと null を返す / getFiles returns null when it fails */
            var folderItems = (nameMask === undefined) ? sourceFolder.getFiles() : sourceFolder.getFiles(nameMask);
            return folderItems ? folderItems : [];
        } catch (e) {
            return [];
        }
    }

    /**
     * Illustratorのセッションのロックファイルを探す
     * @returns {File|null} いちばん新しいロックファイル。無ければ null
     */
    function findSessionLockFile() {
        var adobeFolder = new Folder(Folder.userData.fsName + "/Adobe");
        if (!adobeFolder.exists) return null;

        /* 設定フォルダーの名前はビルドによって変わる（Adobe Illustrator 30 / Adobe Illustrator 30.8.0 Beta Settings）
           ので、メジャーバージョンで始まるものをまとめて見る / The folder name differs between builds */
        var settingsFolders = listFolderEntries(adobeFolder, "Adobe Illustrator " + parseInt(app.version, 10) + "*");
        var newestLockFile = null;

        for (var i = 0; i < settingsFolders.length; i++) {
            if (!(settingsFolders[i] instanceof Folder)) continue;

            /* ロックファイルは言語名のフォルダーの中にある / The lock file sits in the folder named after the language */
            var localeFolders = listFolderEntries(settingsFolders[i]);
            for (var j = 0; j < localeFolders.length; j++) {
                if (!(localeFolders[j] instanceof Folder)) continue;

                var lockFiles = listFolderEntries(localeFolders[j], SESSION_LOCK_FILE_MASK);
                for (var k = 0; k < lockFiles.length; k++) {
                    /* 落ちたセッションのロックが残っていることがあるので、いちばん新しいものを選ぶ
                       / A crashed session can leave one behind, so take the newest */
                    if (!newestLockFile || lockFiles[k].modified > newestLockFile.modified) newestLockFile = lockFiles[k];
                }
            }
        }
        return newestLockFile;
    }

    /**
     * このセッションのidを調べる
     * @returns {string|null} ロックファイルの時刻から作ったid。見つからなければ null
     */
    function lookUpSessionId() {
        var lockFile = findSessionLockFile();
        if (!lockFile) return null;

        /* 作られたあとは書き換えられないファイルなので、更新時刻が起動時刻になる / Never written after it is created */
        var lockTime = lockFile.modified || lockFile.created;
        return lockTime ? String(lockTime.getTime()) : null;
    }

    /**
     * このセッションのidを返す
     * @returns {string|null} セッションのid。調べられなかったときは null
     */
    function readSessionId() {
        if (!isSessionIdChecked) {
            isSessionIdChecked = true;
            cachedSessionId = lookUpSessionId();
        }
        return cachedSessionId;
    }

    /**
     * 記録ファイルを「項目名 → 値」の形で読み出す
     * @returns {Object|null} 読み出した項目。ファイルが無ければ null
     */
    function readSessionFields() {
        var savedText = readTextFile(temporaryFile(SESSION_FILE_NAME));
        if (!savedText) return null;

        var savedLines = savedText.split(/\r\n|\r|\n/);
        var savedFields = {};
        for (var i = 0; i < savedLines.length; i++) {
            var separatorIndex = savedLines[i].indexOf("\t");
            if (separatorIndex === -1) continue;
            savedFields[savedLines[i].substring(0, separatorIndex)] = savedLines[i].substring(separatorIndex + 1);
        }
        return savedFields;
    }

    /**
     * 記録の項目を文字列として読む
     * @param {Object} savedFields - readSessionFields() の結果
     * @param {string} fieldName - 項目名
     * @param {string|null} defaultValue - 記録がないときの値
     * @returns {string|null} 記録された文字列、または defaultValue
     */
    function readTextField(savedFields, fieldName, defaultValue) {
        return (savedFields[fieldName] === undefined) ? defaultValue : savedFields[fieldName];
    }

    /**
     * 記録の項目を真偽値として読む（"1" が true）
     * @param {Object} savedFields - readSessionFields() の結果
     * @param {string} fieldName - 項目名
     * @param {boolean} defaultValue - 記録がないときの値
     * @returns {boolean} 記録された値、または defaultValue
     */
    function readFlagField(savedFields, fieldName, defaultValue) {
        return (savedFields[fieldName] === undefined) ? defaultValue : savedFields[fieldName] === "1";
    }

    /**
     * このセッションで前回閉じたときの状態を読み出す
     * @returns {{remember: boolean, targetFolder: string|null, keyword: string, includeSubfolders: boolean, showFullPath: boolean, folderPath: string|null, fileName: string|null}} 記録がなければユーザー設定の初期値
     */
    function readSessionState() {
        var defaultState = {
            remember: REMEMBER_SETTINGS_DEFAULT,
            targetFolder: null,
            keyword: "",
            includeSubfolders: INCLUDE_SUBFOLDERS_DEFAULT,
            showFullPath: SHOW_FULL_PATH_DEFAULT,
            folderPath: null,
            fileName: null
        };

        var sessionId = readSessionId();
        if (!sessionId) return defaultState;

        /* 前回のIllustratorが残した記録は使わない / Ignore what an earlier Illustrator left behind */
        var savedFields = readSessionFields();
        if (!savedFields || savedFields.session !== sessionId) return defaultState;

        return {
            /* 「記憶しない」で閉じたあとは初期値だけが書いてあるので、そのまま読めばよい / Defaults were stored */
            remember: readFlagField(savedFields, "remember", defaultState.remember),
            targetFolder: readTextField(savedFields, "targetFolder", defaultState.targetFolder),
            keyword: readTextField(savedFields, "keyword", defaultState.keyword),
            includeSubfolders: readFlagField(savedFields, "includeSubfolders", defaultState.includeSubfolders),
            showFullPath: readFlagField(savedFields, "showFullPath", defaultState.showFullPath),

            /* 「すべて」を選んでいたときは記録しない。空文字はルート直下を指すので区別する / An empty value means the root */
            folderPath: readTextField(savedFields, "folderPath", defaultState.folderPath),
            fileName: readTextField(savedFields, "fileName", defaultState.fileName)
        };
    }

    /**
     * このセッションの記憶を書き換える
     * @param {Object} changedFields - 書き換える項目だけを持つオブジェクト
     * @returns {void}
     */
    function saveSessionState(changedFields) {
        var sessionId = readSessionId();
        if (!sessionId) return;

        var sessionState = readSessionState();
        for (var stateKey in changedFields) {
            if (changedFields.hasOwnProperty(stateKey)) sessionState[stateKey] = changedFields[stateKey];
        }

        var savedFields = {
            session: sessionId,
            remember: sessionState.remember ? "1" : "0",
            targetFolder: sessionState.targetFolder,
            keyword: sessionState.keyword,
            includeSubfolders: sessionState.includeSubfolders ? "1" : "0",
            showFullPath: sessionState.showFullPath ? "1" : "0",
            folderPath: sessionState.folderPath,
            fileName: sessionState.fileName
        };

        var savedLines = [];
        for (var fieldName in savedFields) {
            /* null の項目は行ごと書かない。読み出すときに「記録なし」と区別できるようにする / Skip empty fields */
            if (!savedFields.hasOwnProperty(fieldName)) continue;
            if (savedFields[fieldName] === null) continue;

            /* 改行やタブが混じると行の形が壊れるので、その項目はあきらめる / Would break the line format */
            var fieldValue = String(savedFields[fieldName]);
            if (/[\t\r\n]/.test(fieldValue)) continue;

            savedLines.push(fieldName + "\t" + fieldValue);
        }
        writeTextFile(temporaryFile(SESSION_FILE_NAME), savedLines.join("\n"));
    }

    // =========================================
    // テキストファイルの読み書き / Text files
    // =========================================

    /**
     * テキストファイルを読み込む
     * @param {File} textFile - 読み込むファイル
     * @returns {string} 中身。読めなければ空文字
     */
    function readTextFile(textFile) {
        if (!textFile.exists) return "";
        try {
            textFile.encoding = "UTF-8";
            if (!textFile.open("r")) return "";
            return textFile.read();
        } catch (e) {
            return "";
        } finally {
            try { textFile.close(); } catch (e) {}
        }
    }

    /**
     * テキストファイルへ書き出す
     * @param {File} textFile - 書き出すファイル
     * @param {string} fileText - 書き出す内容
     * @returns {boolean} 書き出せたら true
     */
    function writeTextFile(textFile, fileText) {
        try {
            textFile.encoding = "UTF-8";
            textFile.lineFeed = "Unix";
            if (!textFile.open("w")) return false;

            textFile.write(fileText);
            return true;
        } catch (e) {
            return false;
        } finally {
            try { textFile.close(); } catch (e) {}
        }
    }

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* 密なパネル・行の間隔 / Spacing for dense panels and rows */
    var DENSE_SPACING      = 8;              /* パネル内を詰めるときの間隔 / dense panel spacing */
    var LIST_LABEL_SPACING = 4;              /* 見出しとリストの間隔 / gap under a list caption */

    /* リストと入力欄の寸法 / List and field sizes */
    var FOLDER_LIST_SIZE  = [165, 360];      /* フォルダーリストの寸法 [幅,高さ] / folder list size */
    var SCRIPT_LIST_SIZE  = [280, 360];      /* ファイル名リストの寸法 [幅,高さ]。ダイアログ幅＝キーワード欄の長さを決める / script list size */

    /* ボタンの寸法と余白 / Button sizes and margins */
    var BUTTON_HEIGHT           = 28;        /* ボタンの高さ / button height */
    var FOLDER_BUTTON_WIDTH     = 120;       /* フォルダー変更ボタンの幅 / change folder button width */
    var DIALOG_BUTTON_WIDTH     = 92;        /* 実行・キャンセルの幅 / dialog button width */
    var SETTINGS_BUTTON_WIDTH   = 100;       /* 環境設定ボタンの幅 / preferences button width */

    /* 環境設定ダイアログ / Preferences dialog */
    var SETTINGS_LABEL_WIDTH = 100;          /* 項目名の幅 / field label width */
    var SETTINGS_INPUT_WIDTH = 60;           /* 数値入力欄の幅 / number field width */
    var FOLDER_PATH_WIDTH    = 360;          /* 対象フォルダーのパス表示の幅 / target folder path width */
    var CHECKBOX_TOP_MARGIN     = 10;        /* リスト下のチェックボックスの上余白 / top margin above a checkbox under a list */
    var REMEMBER_ROW_TOP_MARGIN = 5;         /* 「検索条件を記憶」の上余白 / top margin above the remember checkbox */

    /* キーワードボタンは小ぶりにする / Keyword preset buttons are smaller */
    var PRESET_BUTTON_HEIGHT  = 22;          /* キーワードボタンの高さ / preset button height */
    var PRESET_BUTTON_PADDING = 16;          /* 文字幅に足す左右の余白 / horizontal padding */
    var PRESET_CHAR_WIDTH     = 7;           /* 実測できない環境用の1文字概算幅 / fallback char width */

    /* キーワード欄のクリアボタン（自前描画） / Clear button next to the keyword field (custom drawn) */
    var CLEAR_BUTTON_SIZE    = 20;           /* クリアボタンの一辺 / clear button size */
    var CLEAR_CIRCLE_INSET   = 2;            /* 円と外周の間隔 / inset of the circle */
    var CLEAR_GLYPH_INSET    = 6;            /* ×と外周の間隔 / inset of the × glyph */
    var CLEAR_STROKE_WIDTH   = 1.5;          /* 円と×の線幅 / stroke width */

    /**
     * ウィンドウの共通設定を適用する
     * @param {Window} targetWindow - 対象ウィンドウ
     * @param {number} [spacing] - 要素間隔。省略時は WINDOW_SPACING
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定を適用する
     * @param {Panel} targetPanel - 対象パネル
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING
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
     * 行グループの共通設定を適用する（ボタン列など）
     * @param {Group} rowGroup - 対象グループ
     * @param {string} [alignment] - 配置。省略時は "left"
     * @param {number} [spacing] - 要素間隔。省略時は PANEL_SPACING
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = [alignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 見出し付きのリスト用カラムを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} captionText - リストの見出し
     * @returns {Group} 見出しを追加済みのカラムグループ
     */
    function addListColumn(parentGroup, captionText) {
        var listColumn = parentGroup.add("group");
        listColumn.orientation = "column";
        listColumn.alignChildren = ["fill", "top"];
        listColumn.spacing = LIST_LABEL_SPACING;
        listColumn.add("statictext", undefined, captionText);
        return listColumn;
    }

    /**
     * ボタンの寸法をそろえる
     * @param {Button} targetButton - 対象ボタン
     * @param {number} buttonWidth - ボタンの幅
     * @returns {void}
     */
    function applyButtonSize(targetButton, buttonWidth) {
        targetButton.preferredSize = [buttonWidth, BUTTON_HEIGHT];
        targetButton.minimumSize = [buttonWidth, BUTTON_HEIGHT];
    }

    /**
     * キーワードボタンの寸法を文字幅に合わせて詰める
     * @param {Button} presetButton - 対象ボタン
     * @returns {void}
     */
    function applyPresetButtonSize(presetButton) {
        /* 実測できればそれを使い、駄目なら文字数から概算する / Measure if possible, else estimate */
        var textWidth = String(presetButton.text).length * PRESET_CHAR_WIDTH;
        try {
            var measuredSize = presetButton.graphics.measureString(presetButton.text);
            var measuredWidth = (measuredSize.width !== undefined) ? measuredSize.width : measuredSize[0];
            /* 環境によっては値が取れずNaNになる。その場合は概算のままにする / Keep the estimate if unusable */
            if (!isNaN(measuredWidth) && measuredWidth > 0) textWidth = measuredWidth;
        } catch (e) {}

        var buttonWidth = Math.ceil(textWidth) + PRESET_BUTTON_PADDING;
        presetButton.preferredSize = [buttonWidth, PRESET_BUTTON_HEIGHT];
        presetButton.minimumSize = [buttonWidth, PRESET_BUTTON_HEIGHT];
    }

    /**
     * IllustratorのUIが明るい設定かどうかを返す
     * @returns {boolean} 明るいUIなら true
     */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /**
     * クリアボタンの丸と×を描く
     * @param {Button} clearButton - 対象ボタン
     * @returns {void}
     */
    function drawClearButton(clearButton) {
        var buttonGraphics = clearButton.graphics;
        var buttonSize = clearButton.size[0];
        var useLightUI = isLightUI();

        /* 主役ではないので、通常時も少し薄めに描く。押している間は濃くして手応えを出す / Slightly muted, darker while pressed */
        var glyphColor = useLightUI ? [0.45, 0.45, 0.45, 1] : [0.72, 0.72, 0.72, 1];
        if (clearButton.pressed) glyphColor = useLightUI ? [0.20, 0.20, 0.20, 1] : [0.95, 0.95, 0.95, 1];

        /* キーワードが空のときは無効にしてあるので、さらに淡くする / Dim further while disabled */
        if (clearButton.enabled === false) glyphColor = useLightUI ? [0.78, 0.78, 0.78, 1] : [0.42, 0.42, 0.42, 1];

        /* 前回の描画を消すため、まず親と同じ色で塗りつぶす。背景色のブラシを持たない環境では省く
           / Repaint the background first; skipped where no background brush is available */
        try {
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.backgroundColor);
        } catch (e) {}

        var glyphPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, glyphColor, CLEAR_STROKE_WIDTH);
        var circleSize = buttonSize - CLEAR_CIRCLE_INSET * 2;
        buttonGraphics.newPath();
        buttonGraphics.ellipsePath(CLEAR_CIRCLE_INSET, CLEAR_CIRCLE_INSET, circleSize, circleSize);
        buttonGraphics.strokePath(glyphPen);

        var glyphStart = CLEAR_GLYPH_INSET;
        var glyphEnd = buttonSize - CLEAR_GLYPH_INSET;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(glyphStart, glyphStart);
        buttonGraphics.lineTo(glyphEnd, glyphEnd);
        buttonGraphics.strokePath(glyphPen);
        buttonGraphics.newPath();
        buttonGraphics.moveTo(glyphEnd, glyphStart);
        buttonGraphics.lineTo(glyphStart, glyphEnd);
        buttonGraphics.strokePath(glyphPen);
    }

    /**
     * キーワードを消すクリアボタンを作る
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {Button} 丸に×を自前描画したボタン
     */
    function addClearButton(parentGroup) {
        var clearButton = parentGroup.add("button", undefined, "");
        clearButton.helpTip = getLabel(LABELS.tooltip.clearKeyword);
        clearButton.alignment = ["right", "center"];
        clearButton.preferredSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        clearButton.minimumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        clearButton.maximumSize = [CLEAR_BUTTON_SIZE, CLEAR_BUTTON_SIZE];
        clearButton.pressed = false;

        /* 実際の状態は呼び出し側が入力欄に合わせる / The caller syncs this with the field */
        clearButton.enabled = false;

        clearButton.onDraw = function () {
            drawClearButton(this);
        };
        clearButton.addEventListener("mousedown", function () {
            this.pressed = true;
            this.notify("onDraw");
        });
        clearButton.addEventListener("mouseup", function () {
            this.pressed = false;
            this.notify("onDraw");
        });
        /* ボタンの外でマウスを離したときも押下表示を戻す / Reset the pressed look on mouseout */
        clearButton.addEventListener("mouseout", function () {
            if (!this.pressed) return;
            this.pressed = false;
            this.notify("onDraw");
        });

        return clearButton;
    }

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

    // =========================================
    // ラベル定義 / Labels
    // =========================================

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

    var LABELS = {
        dialog: {
            title: { ja: "スクリプトランチャー", en: "Script Launcher" },
            selectFolder: { ja: "検索対象のスクリプトフォルダーを選択してください", en: "Select the script folder to search" },
            preferences: { ja: "環境設定", en: "Preferences" }
        },
        panel: {
            keyword: { ja: "絞り込み（%1件）", en: "Filter (%1)" },
            scriptFolder: { ja: "対象フォルダー", en: "Target Folder" },
            keywordButtons: { ja: "キーワードボタン", en: "Keyword Buttons" }
        },
        listCaption: {
            folder: { ja: "フォルダー", en: "Folder" },
            fileName: { ja: "ファイル名", en: "File Name" }
        },
        listItem: {
            allFolders: { ja: "（すべて）", en: "(All)" },
            rootFolder: { ja: "（ルート）", en: "(Root)" }
        },
        fieldLabel: {
            keyword: { ja: "キーワード", en: "Keyword" },
            minCount: { ja: "出現数", en: "Occurrences" },
            maxButtons: { ja: "キーワード数", en: "Keywords" }
        },
        checkbox: {
            includeSubfolders: { ja: "サブディレクトリを含む", en: "Include subdirectories" },
            showFullPath: { ja: "フルパス", en: "Full path" },
            rememberSettings: { ja: "検索条件を記憶", en: "Remember the search" }
        },
        tooltip: {
            rememberSettings: { ja: "Illustratorを終了するまでのあいだ引き継ぎます", en: "Carried over until Illustrator quits" },
            clearKeyword: { ja: "キーワードをクリア", en: "Clear keyword" },
            keyword: {
                ja: "空白で区切ると、すべての語を含むものに絞り込みます。大文字小文字・全角半角・ひらがなとカタカナの違いは区別しません",
                en: "Separate words with spaces to match all of them. Case, full-width/half-width, and hiragana/katakana differences are ignored"
            },
            keywordPreset: {
                ja: "クリックでキーワードを置き換えます。option＋クリックでキーワードに追加します",
                en: "Click to replace the keyword. Option-click adds it to the keyword"
            },
            folderList: {
                ja: "ダブルクリックでフォルダーを Finder で開きます。（すべて）と（ルート）は対象フォルダーを開きます",
                en: "Double-click to open the folder in Finder. (All) and (Root) open the target folder"
            },
            scriptList: {
                ja: "ダブルクリックで実行します。option＋ダブルクリックでは実行せず Finder で表示します",
                en: "Double-click to run. Option-double-click reveals the file in Finder instead"
            },
            includeSubfolders: {
                ja: "OFF のときは、対象フォルダー直下と、その直下のフォルダーにあるスクリプトだけを表示します",
                en: "When off, only scripts in the target folder and its immediate subfolders are listed"
            },
            showFullPath: { ja: "OFF のときはホームフォルダーを ~ で表示します", en: "When off, the home folder is shown as ~" },
            minCount: {
                ja: "この数以上のファイル名に含まれる語をキーワードボタンにします",
                en: "Words found in at least this many file names become keyword buttons"
            },
            maxButtons: { ja: "キーワードボタンを並べる最大数です（%1まで）", en: "Maximum number of keyword buttons (up to %1)" },
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
        },
        button: {
            changeFolder: { ja: "フォルダー変更", en: "Change Folder" },
            preferences: { ja: "環境設定", en: "Preferences" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            run: { ja: "実行", en: "Run" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noScripts: {
                ja: "対象フォルダー内に .jsx / .js / .jsxbin ファイルが見つかりませんでした。",
                en: "No .jsx / .js / .jsxbin files were found in the target folder."
            },
            missingScript: {
                ja: "選択したスクリプトが見つかりません。",
                en: "The selected script could not be found."
            },
            runFailed: {
                ja: "スクリプトの実行中にエラーが発生しました。\n\n%1\n\nエラー: %2",
                en: "An error occurred while running the script.\n\n%1\n\nError: %2"
            },
            errorLine: {
                ja: "\n行: %1",
                en: "\nLine: %1"
            }
        }
    };

    // =========================================
    // スクリプトの収集と検索キー / Script collection and search keys
    // =========================================

    /**
     * スクリプト1件分の情報
     * @typedef {object} ScriptEntry
     * @property {File} file - スクリプトファイル
     * @property {string} fileName - デコード済みのファイル名
     * @property {string} localizedFileName - OSの表示名
     * @property {string} relativePath - 対象フォルダーからの相対パス
     * @property {string} folderPath - 相対パスのフォルダー部分
     * @property {number} folderDepth - 対象フォルダーからの階層の深さ
     * @property {string} searchText - 検索対象を連結した文字列
     * @property {string} normalizedSearchText - 正規化済みの検索キー
     * @property {Array<string>} nameWords - ファイル名から取り出した語（キーワードボタン用）
     */

    /**
     * 前後の空白を取り除く
     * @param {string} sourceText - 対象の文字列
     * @returns {string} 前後の空白を除いた文字列
     */
    function trimWhitespace(sourceText) {
        return String(sourceText).replace(/^\s+|\s+$/g, "");
    }

    /* 半角カナを全角カタカナへ置き換える並び。U+FF61 から順に対応する / Half-width kana in code point order */
    var KANA_HALFWIDTH_START = 0xFF61;
    var KANA_HALFWIDTH_TABLE = "。「」、・ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン";
    var KANA_HALFWIDTH_END   = KANA_HALFWIDTH_START + KANA_HALFWIDTH_TABLE.length - 1;

    /**
     * カタカナを濁点・半濁点・小書きのない基音へ寄せる
     * 変換表は持たず、Unicodeの並び方から計算する
     * @param {number} charCode - カタカナのコードポイント
     * @returns {number} 基音のコードポイント。カタカナ以外はそのまま返す
     */
    function foldKatakana(charCode) {
        /* ヴ と ヷヸヹヺ は並びから外れるので個別に扱う / These sit outside the regular rows */
        if (charCode === 0x30F4) return 0x30A6;
        if (charCode >= 0x30F7 && charCode <= 0x30FA) return charCode - 8;

        /* カ〜ヂ、ツ〜ド は「清音・濁音」の2つ並び。ッ を挟んで並びが一度切れる / Pairs, broken by ッ */
        if (charCode >= 0x30AB && charCode <= 0x30C2 && (charCode - 0x30AB) % 2 === 1) charCode -= 1;
        else if (charCode >= 0x30C4 && charCode <= 0x30C9 && (charCode - 0x30C4) % 2 === 1) charCode -= 1;

        /* ハ〜ポ は「清音・濁音・半濁音」の3つ並び / The ha row runs in threes */
        else if (charCode >= 0x30CF && charCode <= 0x30DD) charCode -= (charCode - 0x30CF) % 3;

        /* 小書きは対応する大書きの1つ手前にある / Each small kana sits right before its large form */
        if (charCode === 0x30A1 || charCode === 0x30A3 || charCode === 0x30A5 || charCode === 0x30A7 || charCode === 0x30A9 ||
            charCode === 0x30C3 || charCode === 0x30E3 || charCode === 0x30E5 || charCode === 0x30E7 || charCode === 0x30EE) {
            charCode += 1;
        }
        return charCode;
    }

    /**
     * 比較用の検索キーへ変換する
     * 大文字小文字・全角半角・かなの種類・濁点や小書きの違いを無視して一致させる
     * @param {string} sourceText - 変換前の文字列
     * @returns {string} 正規化した文字列
     */
    function normalizeSearchKey(sourceText) {
        var lowerText = String(sourceText).toLowerCase();
        var normalizedText = "";

        for (var i = 0; i < lowerText.length; i++) {
            var currentChar = lowerText.charAt(i);
            if (/\s/.test(currentChar)) continue;

            var charCode = lowerText.charCodeAt(i);

            /* 単独の濁点・半濁点は落とす。濁りは基音へ寄せるので不要 / Standalone voiced marks are dropped */
            if (charCode === 0x3099 || charCode === 0x309A || charCode === 0x309B || charCode === 0x309C ||
                charCode === 0xFF9E || charCode === 0xFF9F) continue;

            /* 全角の英数記号は半角へ / Full-width ASCII to half-width */
            if (charCode >= 0xFF01 && charCode <= 0xFF5E) {
                normalizedText += String.fromCharCode(charCode - 0xFEE0).toLowerCase();
                continue;
            }

            /* 半角カナは全角カタカナへ / Half-width kana to full-width katakana */
            if (charCode >= KANA_HALFWIDTH_START && charCode <= KANA_HALFWIDTH_END) {
                charCode = KANA_HALFWIDTH_TABLE.charCodeAt(charCode - KANA_HALFWIDTH_START);
            }

            /* ひらがなはカタカナへ寄せる / Hiragana to katakana */
            if (charCode >= 0x3041 && charCode <= 0x3096) charCode += 0x60;

            normalizedText += String.fromCharCode(foldKatakana(charCode));
        }
        return normalizedText;
    }

    /**
     * 入力文字列を空白区切りの検索語に分解する
     * @param {string} keywordText - キーワード欄の文字列
     * @returns {Array<string>} 正規化した検索語。空の語は含まない
     */
    function splitSearchTerms(keywordText) {
        var rawTerms = String(keywordText).split(/\s+/);
        var searchTerms = [];
        for (var i = 0; i < rawTerms.length; i++) {
            var searchTerm = normalizeSearchKey(rawTerms[i]);
            if (searchTerm !== "") searchTerms.push(searchTerm);
        }
        return searchTerms;
    }

    /**
     * 検索語をすべて含むかどうかを判定する（AND検索）
     * @param {string} searchTarget - 正規化済みの検索対象
     * @param {Array<string>} searchTerms - 正規化済みの検索語
     * @returns {boolean} すべて含むなら true
     */
    function matchesSearchTerms(searchTarget, searchTerms) {
        for (var i = 0; i < searchTerms.length; i++) {
            if (searchTarget.indexOf(searchTerms[i]) === -1) return false;
        }
        return true;
    }

    /**
     * 対象フォルダーを選ばせる
     * @param {Folder|null} currentFolder - 初期表示に使うフォルダー
     * @returns {Folder|null} 選ばれたフォルダー。取り消し時は null
     */
    function chooseScriptFolder(currentFolder) {
        var startFolder = currentFolder && currentFolder.exists ? currentFolder : Folder.myDocuments;
        return startFolder.selectDlg(getLabel(LABELS.dialog.selectFolder));
    }

    /**
     * 対象フォルダーのパスを表示用に整える
     * @param {Folder} targetFolder - 表示するフォルダー
     * @param {boolean} showFullPath - true でフルパス、false でホームフォルダーを ~ に略す
     * @returns {string} 表示用のパス文字列
     */
    function formatFolderPath(targetFolder, showFullPath) {
        var fullPath = targetFolder.fsName;
        if (showFullPath) return fullPath;

        var homePath = Folder("~").fsName;
        if (homePath && fullPath.indexOf(homePath) === 0) {
            /* 区切りの手前で切れているか確かめる。/Users/tak が /Users/takano に一致しないように */
            var pathAfterHome = fullPath.substring(homePath.length);
            if (pathAfterHome === "" || pathAfterHome.charAt(0) === "/" || pathAfterHome.charAt(0) === "\\") return "~" + pathAfterHome;
        }
        return fullPath;
    }

    /**
     * decodeURI に失敗しても元の文字列を返す
     * @param {string} encodedText - デコードする文字列
     * @returns {string} デコード結果。失敗時は元の文字列
     */
    function decodeSafely(encodedText) {
        try {
            return decodeURI(encodedText);
        } catch (e) {
            return encodedText;
        }
    }

    /**
     * フォルダーを再帰的にたどってスクリプトを集める
     * @param {Folder} targetFolder - 走査するフォルダー
     * @param {Array<ScriptEntry>} collectedEntries - 収集結果の追加先
     * @param {Folder} rootFolder - 相対パスの基準になる対象フォルダー
     * @returns {void}
     */
    function collectScriptFiles(targetFolder, collectedEntries, rootFolder) {
        var folderItems;
        try {
            folderItems = targetFolder.getFiles();
        } catch (e) {
            return;
        }

        /* ファイルごとに変わらない値は先に1回だけ求める / Compute the per-folder constants once */
        var launcherPath = File($.fileName).fsName;
        var rootPath = rootFolder.fsName;

        for (var i = 0; i < folderItems.length; i++) {
            var folderItem = folderItems[i];

            if (folderItem instanceof Folder) {
                /* エイリアスは循環の元になるのでたどらない / Aliases can loop back into an ancestor */
                if (!folderItem.alias && !/^\./.test(folderItem.name)) collectScriptFiles(folderItem, collectedEntries, rootFolder);
                continue;
            }

            if (!(folderItem instanceof File) || !SCRIPT_EXT_RE.test(folderItem.name)) continue;

            /* 対象フォルダー内にこのランチャー自身がある場合は一覧に載せない / Skip this launcher itself */
            if (launcherPath === folderItem.fsName) continue;

            var relativePath = folderItem.fsName;
            if (relativePath.indexOf(rootPath) === 0) {
                relativePath = relativePath.substring(rootPath.length).replace(/^[\\\/]+/, "");
            }

            var decodedFileName = decodeSafely(folderItem.name);
            var localizedFileName = folderItem.displayName || decodedFileName;
            var decodedRelativePath = decodeSafely(relativePath);

            /* リストは左右2本なので、フォルダー部分とファイル名を分けて持つ / Split folder and file name for the two lists */
            var decodedFolderPath = decodedRelativePath.replace(/[\\\/][^\\\/]*$/, "");
            if (decodedFolderPath === decodedRelativePath) decodedFolderPath = "";

            /* 対象フォルダーからの階層の深さ。直下は0、artboard/backup は2 / Depth below the root folder */
            var folderDepth = (decodedFolderPath === "") ? 0 : decodedFolderPath.split(/[\\\/]/).length;

            collectedEntries.push({
                file: folderItem,
                fileName: decodedFileName,
                localizedFileName: localizedFileName,
                relativePath: decodedRelativePath,
                folderPath: decodedFolderPath,
                folderDepth: folderDepth,
                searchText: decodedRelativePath + " " + localizedFileName
            });
        }
    }

    /**
     * 相対パス順にスクリプトを並べ替える
     * @param {Array<ScriptEntry>} scriptEntries - 並べ替える配列（破壊的に変更する）
     * @returns {void}
     */
    function sortScriptEntries(scriptEntries) {
        scriptEntries.sort(function (entryA, entryB) {
            var pathA = entryA.relativePath.toLowerCase();
            var pathB = entryB.relativePath.toLowerCase();
            return pathA < pathB ? -1 : (pathA > pathB ? 1 : 0);
        });
    }

    /**
     * 入力文字列を1以上の整数として読み取る
     * @param {string} inputText - 入力文字列
     * @param {number} fallbackValue - 数値として読めないときに返す値
     * @param {number} [maxValue] - 上限。超えた場合はこの値に丸める
     * @returns {number} 1以上の整数
     */
    function parsePositiveInt(inputText, fallbackValue, maxValue) {
        var parsedValue = parseInt(trimWhitespace(inputText), 10);
        if (isNaN(parsedValue) || parsedValue < 1) return fallbackValue;
        if (typeof maxValue === "number" && parsedValue > maxValue) return maxValue;
        return parsedValue;
    }

    /**
     * ボタンにしない語かどうかを判定する
     * @param {string} word - 小文字化した語
     * @returns {boolean} 除外語なら true
     */
    function isStopWord(word) {
        for (var i = 0; i < KEYWORD_STOP_WORDS.length; i++) {
            if (KEYWORD_STOP_WORDS[i] === word) return true;
        }
        return false;
    }

    /**
     * ファイル名を単語に分解する（キャメルケース・ハイフン・アンダースコア区切り）
     * @param {string} fileName - 拡張子付きのファイル名
     * @returns {Array<string>} 小文字化した単語の配列
     */
    function extractNameWords(fileName) {
        var baseName = String(fileName).replace(/\.[^.]+$/, "");

        /* 区切り文字と数字を空白にし、続けてキャメルケースの境目を空ける / Split separators, then camel case */
        var spacedName = baseName.replace(/[-_\s\d]+/g, " ");
        spacedName = spacedName.replace(/([a-z])([A-Z])/g, "$1 $2");
        spacedName = spacedName.replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2");

        var nameWords = [];
        var nameParts = spacedName.split(" ");
        for (var i = 0; i < nameParts.length; i++) {
            var word = nameParts[i].toLowerCase();
            /* 英字だけの語に限る。日本語名や記号混じりはボタンにしない / ASCII words only */
            if (!/^[a-z]+$/.test(word)) continue;
            if (word.length < KEYWORD_MIN_WORD_LENGTH) continue;
            if (isStopWord(word)) continue;
            nameWords.push(word);
        }
        return nameWords;
    }

    /**
     * すでに入力済みの検索語で覆われている語かどうかを判定する
     * @param {string} word - 判定する語
     * @param {Array<string>} searchTerms - 正規化済みの検索語
     * @returns {boolean} いずれかの検索語を含むなら true
     */
    function isCoveredByTerms(word, searchTerms) {
        for (var i = 0; i < searchTerms.length; i++) {
            if (word.indexOf(searchTerms[i]) !== -1) return true;
        }
        return false;
    }

    /**
     * スクリプト名によく出てくる語を求める
     * @param {Array<ScriptEntry>} scriptEntries - 集計対象のスクリプト
     * @param {number} minCount - ボタンにする最小出現ファイル数
     * @param {number} maxButtons - 返す語の最大個数
     * @param {Array<string>} searchTerms - 入力済みの検索語。これを含む語は除く
     * @returns {Array<string>} 固定の語に続けて、出現ファイル数の多い順に並べた語
     */
    function collectFrequentWords(scriptEntries, minCount, maxButtons, searchTerms) {
        var wordCounts = {};
        for (var i = 0; i < scriptEntries.length; i++) {
            var nameWords = scriptEntries[i].nameWords;

            /* 同じ語が1つのファイル名に複数あっても1件と数える / Count each file once per word */
            var seenWords = {};
            for (var j = 0; j < nameWords.length; j++) {
                var wordKey = "#" + nameWords[j];
                if (seenWords[wordKey]) continue;
                seenWords[wordKey] = true;
                wordCounts[wordKey] = (wordCounts[wordKey] || 0) + 1;
            }
        }

        var presetWords = [];
        var pinnedKeys = {};
        for (var p = 0; p < KEYWORD_PINNED_WORDS.length; p++) {
            var pinnedWord = KEYWORD_PINNED_WORDS[p];
            pinnedKeys["#" + pinnedWord] = true;
            if (presetWords.length >= maxButtons) break;
            if (!wordCounts["#" + pinnedWord]) continue;
            if (isCoveredByTerms(pinnedWord, searchTerms)) continue;
            presetWords.push(pinnedWord);
        }

        var frequentWords = [];
        for (var countKey in wordCounts) {
            if (!wordCounts.hasOwnProperty(countKey)) continue;
            if (pinnedKeys[countKey]) continue;
            if (wordCounts[countKey] < minCount) continue;

            /* 入力済みの語をボタンにしても絞り込めないので外す / A word the query already covers is a no-op */
            var word = countKey.substring(1);
            if (isCoveredByTerms(word, searchTerms)) continue;
            frequentWords.push({ word: word, count: wordCounts[countKey] });
        }

        frequentWords.sort(function (wordA, wordB) {
            if (wordA.count !== wordB.count) return wordB.count - wordA.count;
            return wordA.word < wordB.word ? -1 : (wordA.word > wordB.word ? 1 : 0);
        });

        for (var k = 0; k < frequentWords.length && presetWords.length < maxButtons; k++) {
            presetWords.push(frequentWords[k].word);
        }
        return presetWords;
    }

    /**
     * 語の先頭を大文字にする（ボタン表示用）
     * @param {string} word - 小文字の語
     * @returns {string} 先頭だけ大文字にした語
     */
    function capitalizeWord(word) {
        return word.charAt(0).toUpperCase() + word.substring(1);
    }

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

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * キーワードボタン置き場の高さを固定する（絞り込みのたびにダイアログが伸縮しないようにする）
     * @param {Group} presetContainer - ボタン行を入れるグループ
     * @param {number} maxButtons - 並べうるボタンの最大個数
     * @returns {void}
     */
    function reserveKeywordPresetHeight(presetContainer, maxButtons) {
        var rowCount = Math.ceil(maxButtons / KEYWORD_PRESETS_PER_ROW);
        var reservedHeight = rowCount * PRESET_BUTTON_HEIGHT + (rowCount - 1) * DENSE_SPACING;
        presetContainer.preferredSize.height = reservedHeight;
        presetContainer.minimumSize.height = reservedHeight;
    }

    /**
     * 項目名付きの数値入力欄を1行追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {{ja: string, en: string}} labelSet - 項目名のラベル定義
     * @param {number} initialValue - 初期値
     * @param {string} tooltipText - 入力欄の helpTip
     * @param {Object} stepOptions - ∧∨の設定（min / max / integer。addStepper() に渡す）
     * @returns {EditText} 追加した数値入力欄
     */
    function addNumberField(parentPanel, labelSet, initialValue, tooltipText, stepOptions) {
        var fieldRow = parentPanel.add("group");
        setupRow(fieldRow, "left", DENSE_SPACING);

        var fieldLabel = fieldRow.add("statictext", undefined, labelText(labelSet));
        fieldLabel.preferredSize.width = SETTINGS_LABEL_WIDTH;
        fieldLabel.justify = "right";

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRow.add("group");
        setupRow(stepperInputGroup, "left", 0);
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, String(initialValue));
        numberInput.preferredSize.width = SETTINGS_INPUT_WIDTH;
        numberInput.helpTip = tooltipText;
        bindSteppedArrowKeys(numberInput, stepperGroup);
        return numberInput;
    }

    /**
     * 環境設定（キーワードボタンの出現条件）を編集するダイアログを表示する
     * @param {Folder} targetFolder - 現在の対象フォルダー
     * @param {number} minCount - 現在の出現数
     * @param {number} maxButtons - 現在のキーワード数
     * @param {boolean} remembersSettings - 「フルパス」の状態を覚えるかどうか
     * @returns {{folder: Folder, minCount: number, maxButtons: number}|null} 入力された値。取り消し時は null
     */
    function showPreferencesDialog(targetFolder, minCount, maxButtons, remembersSettings) {
        var settingsDialog = new Window("dialog", getLabel(LABELS.dialog.preferences));
        setupWindow(settingsDialog, DENSE_SPACING);

        var selectedFolder = targetFolder;
        var folderPanelParts = buildScriptFolderPanel(settingsDialog, selectedFolder, remembersSettings);
        folderPanelParts.btnChangeFolder.onClick = function () {
            var pickedFolder = chooseScriptFolder(selectedFolder);
            if (!pickedFolder) return;
            selectedFolder = pickedFolder;
            folderPanelParts.refreshPath(selectedFolder);
        };

        var keywordButtonPanel = settingsDialog.add("panel", undefined, getLabel(LABELS.panel.keywordButtons));
        setupPanel(keywordButtonPanel, DENSE_SPACING);
        var minCountInput = addNumberField(keywordButtonPanel, LABELS.fieldLabel.minCount, minCount,
            getLabel(LABELS.tooltip.minCount), { min: 1, integer: true });
        var maxButtonsInput = addNumberField(keywordButtonPanel, LABELS.fieldLabel.maxButtons, maxButtons,
            getLabel(LABELS.tooltip.maxButtons, [KEYWORD_PRESET_LIMIT]),
            { min: 1, max: KEYWORD_PRESET_LIMIT, integer: true });

        var buttonRow = addButtonRow(settingsDialog);
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        applyButtonSize(btnCancel, DIALOG_BUTTON_WIDTH);
        applyButtonSize(btnOK, DIALOG_BUTTON_WIDTH);

        settingsDialog.center();
        minCountInput.active = true;
        if (settingsDialog.show() !== 1) return null;

        return {
            folder: selectedFolder,
            minCount: parsePositiveInt(minCountInput.text, minCount),
            maxButtons: parsePositiveInt(maxButtonsInput.text, maxButtons, KEYWORD_PRESET_LIMIT)
        };
    }

    /**
     * 対象フォルダーのスクリプトを集め、検索キーまで用意する
     * @param {Folder} targetFolder - 検索対象のフォルダー
     * @returns {Array<ScriptEntry>} 相対パス順に並べ、検索キーを付けたスクリプト
     */
    function loadScriptEntries(targetFolder) {
        var scriptEntries = [];
        collectScriptFiles(targetFolder, scriptEntries, targetFolder);
        sortScriptEntries(scriptEntries);

        /* 検索キーは起動時に一度だけ作ってキャッシュし、入力のたびの再計算を避ける / Cache search keys once */
        for (var i = 0; i < scriptEntries.length; i++) {
            scriptEntries[i].normalizedSearchText = normalizeSearchKey(scriptEntries[i].searchText);
            scriptEntries[i].nameWords = extractNameWords(scriptEntries[i].fileName);
        }
        return scriptEntries;
    }

    /**
     * 絞り込みパネルを組み立てる
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @param {string} initialKeyword - 最初に表示するキーワード
     * @param {boolean} remembersSettings - 「検索条件を記憶」の初期状態
     * @returns {{keywordPanel: Panel, keywordInput: EditText, clearButton: Button, presetContainer: Group, rememberCheckbox: Checkbox}} パネルの部品
     */
    function buildKeywordPanel(parentWindow, initialKeyword, remembersSettings) {
        var keywordPanel = parentWindow.add("panel", undefined, getLabel(LABELS.panel.keyword, [0]));
        setupPanel(keywordPanel, DENSE_SPACING);

        var keywordRow = keywordPanel.add("group");
        setupRow(keywordRow, "fill", DENSE_SPACING);
        keywordRow.add("statictext", undefined, labelText(LABELS.fieldLabel.keyword));

        var keywordInput = keywordRow.add("edittext", undefined, initialKeyword);
        keywordInput.alignment = ["fill", "center"];
        keywordInput.helpTip = getLabel(LABELS.tooltip.keyword);

        var clearButton = addClearButton(keywordRow);

        /* ボタンは絞り込みのたびに作り直すので、置き場だけ先に用意する / Reserve the area for the preset buttons */
        var presetContainer = keywordPanel.add("group");
        presetContainer.orientation = "column";
        presetContainer.alignChildren = ["left", "top"];
        presetContainer.alignment = ["fill", "top"];
        presetContainer.spacing = DENSE_SPACING;

        var rememberRow = keywordPanel.add("group");
        setupRow(rememberRow, "left", 0);
        rememberRow.margins = [0, REMEMBER_ROW_TOP_MARGIN, 0, 0];

        var rememberCheckbox = rememberRow.add("checkbox", undefined, getLabel(LABELS.checkbox.rememberSettings));
        rememberCheckbox.helpTip = getLabel(LABELS.tooltip.rememberSettings);
        rememberCheckbox.value = remembersSettings;

        return {
            keywordPanel: keywordPanel,
            keywordInput: keywordInput,
            clearButton: clearButton,
            presetContainer: presetContainer,
            rememberCheckbox: rememberCheckbox
        };
    }

    /**
     * 対象フォルダーパネルを組み立てる
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @param {Folder} targetFolder - 最初に表示するフォルダー
     * @param {boolean} remembersSettings - 「フルパス」の状態を覚えるかどうか
     * @returns {{btnChangeFolder: Button, refreshPath: function(Folder): void}} 変更ボタンと表示更新関数
     */
    function buildScriptFolderPanel(parentWindow, targetFolder, remembersSettings) {
        var currentFolder = targetFolder;
        var showFullPath = readSessionState().showFullPath;

        var folderPanel = parentWindow.add("panel", undefined, getLabel(LABELS.panel.scriptFolder));
        setupPanel(folderPanel, DENSE_SPACING);

        var pathRow = folderPanel.add("group");
        setupRow(pathRow, "fill", DENSE_SPACING);

        /* 入力欄に見せないよう statictext で表示し、長いパスは中央を省略する / Plain text, truncated in the middle */
        var pathText = pathRow.add("statictext", undefined, formatFolderPath(currentFolder, showFullPath), { truncate: "middle" });
        pathText.preferredSize.width = FOLDER_PATH_WIDTH;

        var btnChangeFolder = pathRow.add("button", undefined, getLabel(LABELS.button.changeFolder));
        btnChangeFolder.alignment = ["right", "center"];
        applyButtonSize(btnChangeFolder, FOLDER_BUTTON_WIDTH);

        var fullPathCheckbox = folderPanel.add("checkbox", undefined, getLabel(LABELS.checkbox.showFullPath));
        fullPathCheckbox.alignment = "left";
        fullPathCheckbox.value = showFullPath;
        fullPathCheckbox.helpTip = getLabel(LABELS.tooltip.showFullPath);

        /**
         * パス表示を今のフォルダーと「フルパス」の状態に合わせる
         * @param {Folder} [newFolder] - 新しいフォルダー。省略時は表示だけ更新する
         * @returns {void}
         */
        function refreshPath(newFolder) {
            if (newFolder) currentFolder = newFolder;
            pathText.text = formatFolderPath(currentFolder, fullPathCheckbox.value);
        }

        fullPathCheckbox.onClick = function () {
            /* 表示の切り替えはセッション中だけ覚えておく / Remember the toggle for this session */
            if (remembersSettings) saveSessionState({ showFullPath: fullPathCheckbox.value });
            refreshPath();
        };
        return { btnChangeFolder: btnChangeFolder, refreshPath: refreshPath };
    }

    /**
     * 左右2本のリストを組み立てる
     * @param {Window} parentWindow - 追加先のウィンドウ
     * @param {boolean} includeSubfolders - 「サブディレクトリを含む」の初期状態
     * @returns {{folderListBox: ListBox, subfoldersCheckbox: Checkbox, scriptListBox: ListBox}} リスト部品
     */
    function buildListColumns(parentWindow, includeSubfolders) {
        /* 左でフォルダーを選び、右にそのフォルダー内のファイル名だけを並べる / Folder on the left, file names on the right */
        var listRow = parentWindow.add("group");
        setupRow(listRow, "fill", COLUMN_SPACING);
        listRow.alignChildren = ["fill", "fill"];

        var folderColumn = addListColumn(listRow, getLabel(LABELS.listCaption.folder));
        var folderListBox = folderColumn.add("listbox", undefined, [], { multiselect: false });
        folderListBox.preferredSize = FOLDER_LIST_SIZE;
        folderListBox.helpTip = getLabel(LABELS.tooltip.folderList);

        var checkboxRow = folderColumn.add("group");
        setupRow(checkboxRow, "left", 0);
        checkboxRow.margins = [0, CHECKBOX_TOP_MARGIN, 0, 0];
        var subfoldersCheckbox = checkboxRow.add("checkbox", undefined, getLabel(LABELS.checkbox.includeSubfolders));
        subfoldersCheckbox.value = includeSubfolders;
        subfoldersCheckbox.helpTip = getLabel(LABELS.tooltip.includeSubfolders);

        var scriptColumn = addListColumn(listRow, getLabel(LABELS.listCaption.fileName));
        var scriptListBox = scriptColumn.add("listbox", undefined, [], { multiselect: false });
        scriptListBox.preferredSize = SCRIPT_LIST_SIZE;
        scriptListBox.helpTip = getLabel(LABELS.tooltip.scriptList);

        return { folderListBox: folderListBox, subfoldersCheckbox: subfoldersCheckbox, scriptListBox: scriptListBox };
    }

    /**
     * Enter / Return でスクリプトを実行し、↓でキーワード欄からファイル名リストへ移るキー操作を組み込む
     * @param {Window} launcherDialog - ランチャーのダイアログ
     * @param {EditText} keywordInput - キーワード欄
     * @param {ListBox} scriptListBox - ファイル名リスト
     * @param {Function} runSelectedScript - 選択中のスクリプトを実行する関数
     * @returns {void}
     */
    function bindLauncherKeys(launcherDialog, keywordInput, scriptListBox, runSelectedScript) {
        keywordInput.addEventListener("keydown", function (event) {
            if (event.keyName === "Down" && scriptListBox.items.length > 0) {
                scriptListBox.active = true;
                if (!scriptListBox.selection) scriptListBox.selection = 0;
                event.preventDefault();
            } else if (event.keyName === "Enter" || event.keyName === "Return") {
                /* 絞り込みは onChanging / onChange で済んでいる。ここで組み直すと選択が先頭へ戻る */
                runSelectedScript();
                event.preventDefault();
                if (event.stopPropagation) event.stopPropagation();
            }
        });

        scriptListBox.addEventListener("keydown", function (event) {
            if (event.keyName === "Enter" || event.keyName === "Return") {
                runSelectedScript();
                event.preventDefault();
                if (event.stopPropagation) event.stopPropagation();
            }
        });

        /* テンキーEnterを含め、ダイアログ内のどこにフォーカスがあっても実行する / Enter runs from anywhere */
        launcherDialog.addEventListener("keydown", function (event) {
            /* ボタンにフォーカスがあるときは、そのボタン自身の動作に任せる */
            if (event.target && event.target.type === "button") return;
            if (event.keyName === "Enter" || event.keyName === "Return") {
                runSelectedScript();
                event.preventDefault();
            }
        });
    }

    /**
     * ランチャーのダイアログを表示する
     * @param {Folder} targetFolder - 検索対象のフォルダー
     * @returns {{action: string, file: File|null, folder: Folder|null}} 操作結果（"run" / "changeFolder" / "cancel"）
     */
    function showLauncherDialog(targetFolder) {
        var scriptEntries = loadScriptEntries(targetFolder);
        if (scriptEntries.length === 0) {
            alert(getLabel(LABELS.alert.noScripts), getLabel(LABELS.dialog.title));
            return { action: "changeFolder", file: null, folder: null };
        }

        /* 同じセッション中に開き直したときは、前回の絞り込みから再開する / Resume the previous search within the session */
        var sessionState = readSessionState();

        /* 対象フォルダーが変わっていれば、フォルダーとファイルの選択は引き継がない
           / Those two are relative to the folder they were stored for */
        if (sessionState.targetFolder !== targetFolder.fsName) {
            sessionState.folderPath = null;
            sessionState.fileName = null;
        }

        var launcherDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        setupWindow(launcherDialog);

        /* 検索欄を最初の操作部品にして、起動時のフォーカスを安定させる / Keep the keyword field first */
        var keywordParts = buildKeywordPanel(launcherDialog, sessionState.keyword, sessionState.remember);
        var keywordPanel = keywordParts.keywordPanel;
        var keywordInput = keywordParts.keywordInput;
        var keywordClearButton = keywordParts.clearButton;
        var keywordPresetContainer = keywordParts.presetContainer;
        var rememberCheckbox = keywordParts.rememberCheckbox;

        /* 絞り込み結果によく出る語をワンクリックで入れる / One-click presets from the filtered results */
        var storedSettings = readKeywordSettings();
        var presetMinCount = storedSettings.minCount;
        var presetMaxButtons = storedSettings.maxButtons;
        var keywordPresetRows = [];
        var currentPresetWords = null;
        var isDialogShown = false;
        reserveKeywordPresetHeight(keywordPresetContainer, presetMaxButtons);

        var listParts = buildListColumns(launcherDialog, sessionState.includeSubfolders);
        var folderListBox = listParts.folderListBox;
        var includeSubfoldersCheckbox = listParts.subfoldersCheckbox;
        var scriptListBox = listParts.scriptListBox;

        /* ボタンエリア：左に［環境設定］、右にキャンセル・実行 / Button row: Preferences on the left, Cancel / Run on the right */
        var buttonRow = addButtonRow(launcherDialog);
        var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel(LABELS.button.preferences));
        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnRun = buttonRow.rightGroup.add("button", undefined, getLabel(LABELS.button.run), { name: "ok" });
        btnRun.enabled = false;
        applyButtonSize(btnPreferences, SETTINGS_BUTTON_WIDTH);
        applyButtonSize(btnCancel, DIALOG_BUTTON_WIDTH);
        applyButtonSize(btnRun, DIALOG_BUTTON_WIDTH);

        var filteredScripts = [];
        var listedScripts = [];
        var dialogResult = { action: "cancel", file: null, folder: null };

        /* 組み直し中の選択変更でファイル名リストが何度も再構築されるのを防ぐ / Suppress cascaded rebuilds */
        var isRebuildingFolderList = false;

        /* 前回の選択は最初の組み立てでだけ使う。以降は通常どおり先頭を選ぶ / Used only while building the lists for the first time */
        var restoringFolderPath = sessionState.folderPath;
        var restoringFileName = sessionState.fileName;

        /* 記録は1回だけ。閉じ方によっては二重に呼ばれる / Guards against being called twice */
        var isSessionSaved = false;

        /**
         * 絞り込み結果に合わせてキーワードボタンを作り直す
         * @param {Array<ScriptEntry>} filteredEntries - 集計対象のスクリプト（現在の絞り込み結果）
         * @param {Array<string>} searchTerms - 入力済みの検索語。これを含む語はボタンにしない
         * @returns {void}
         */
        function refreshKeywordPresetButtons(filteredEntries, searchTerms) {
            var presetWords = collectFrequentWords(filteredEntries, presetMinCount, presetMaxButtons, searchTerms);

            /* 顔ぶれが変わらないなら作り直さない。1打鍵ごとの再構築を避ける / Skip the rebuild when nothing changed */
            var presetWordsKey = presetWords.join(" ");
            if (presetWordsKey === currentPresetWords) return;
            currentPresetWords = presetWordsKey;

            for (var i = 0; i < keywordPresetRows.length; i++) {
                keywordPresetContainer.remove(keywordPresetRows[i]);
            }
            keywordPresetRows = [];

            var presetRow = null;
            for (var j = 0; j < presetWords.length; j++) {
                if (j % KEYWORD_PRESETS_PER_ROW === 0) {
                    presetRow = keywordPresetContainer.add("group");
                    setupRow(presetRow, "left", DENSE_SPACING);
                    keywordPresetRows.push(presetRow);
                }
                var presetButton = presetRow.add("button", undefined, capitalizeWord(presetWords[j]));
                applyPresetButtonSize(presetButton);
                presetButton.helpTip = getLabel(LABELS.tooltip.keywordPreset);
                presetButton.onClick = makeKeywordPresetHandler(presetWords[j]);
            }

            /* 置き場の高さは固定してあるので、ここだけ組み直せばダイアログは動かない / Fixed height keeps the dialog still */
            if (isDialogShown) keywordPresetContainer.layout.layout(true);
        }

        /**
         * プリセットボタン用の onClick ハンドラーを作る
         * @param {string} presetKeyword - ボタンに割り当てるキーワード
         * @returns {function(): void} 通常クリックで置き換え、option+クリックで追加するハンドラー
         */
        function makeKeywordPresetHandler(presetKeyword) {
            return function () {
                /* option+クリックは空白を挟んで語を足し、AND検索で絞り込む / Option-click appends the word */
                var currentText = trimWhitespace(keywordInput.text);
                var isAppend = ScriptUI.environment.keyboardState.altKey && currentText !== "";
                keywordInput.text = isAppend ? currentText + " " + presetKeyword : presetKeyword;
                lastNormalizedQuery = splitSearchTerms(keywordInput.text).join(" ");

                /* 押されたボタンは組み直しで消える。先にフォーカスを入力欄へ移しておく */
                keywordInput.active = true;
                updateClearButtonState();
                refreshFolderList();
            };
        }

        /**
         * 選択中フォルダーの相対パスを返す
         * @returns {string|null} 相対パス。「すべて」選択時と未選択時は null
         */
        function selectedFolderPath() {
            if (!folderListBox.selection) return null;
            var folderPath = folderListBox.selection.folderPath;
            return (folderPath === undefined) ? null : folderPath;
        }

        /**
         * キーワードとサブディレクトリ設定で絞り込み、左のフォルダーリストを組み直す
         * @returns {void}
         */
        function refreshFolderList() {
            var searchTerms = splitSearchTerms(keywordInput.text);

            /* 最初の組み立てだけは前回のセッションで選んでいたフォルダーを狙う / Aim at the remembered folder on the first pass */
            var previousFolderPath = selectedFolderPath();
            if (previousFolderPath === null) previousFolderPath = restoringFolderPath;
            restoringFolderPath = null;

            var depthLimit = includeSubfoldersCheckbox.value ? null : NESTED_FOLDER_DEPTH_LIMIT;
            isRebuildingFolderList = true;
            filteredScripts = [];

            var seenFolderPaths = {};
            var folderPaths = [];
            for (var i = 0; i < scriptEntries.length; i++) {
                var searchTarget = scriptEntries[i].normalizedSearchText;
                if (!matchesSearchTerms(searchTarget, searchTerms)) continue;
                if (depthLimit !== null && scriptEntries[i].folderDepth > depthLimit) continue;

                filteredScripts.push(scriptEntries[i]);

                /* ハッシュのキー衝突を避けるため接頭辞を付けて既出判定する / Prefix the key to avoid collisions */
                var folderKey = "#" + scriptEntries[i].folderPath;
                if (!seenFolderPaths[folderKey]) {
                    seenFolderPaths[folderKey] = true;
                    folderPaths.push(scriptEntries[i].folderPath);
                }
            }
            folderPaths.sort();
            refreshKeywordPresetButtons(filteredScripts, searchTerms);

            folderListBox.removeAll();
            folderListBox.add("item", getLabel(LABELS.listItem.allFolders));
            for (var j = 0; j < folderPaths.length; j++) {
                var folderItem = folderListBox.add("item", folderPaths[j] === "" ? getLabel(LABELS.listItem.rootFolder) : folderPaths[j]);
                folderItem.folderPath = folderPaths[j];
            }

            /* 絞り込み前に選んでいたフォルダーが残っていれば選択を引き継ぐ / Keep the previous folder selection */
            folderListBox.selection = 0;
            if (previousFolderPath !== null) {
                for (var k = 1; k < folderListBox.items.length; k++) {
                    if (folderListBox.items[k].folderPath === previousFolderPath) {
                        folderListBox.selection = k;
                        break;
                    }
                }
            }

            isRebuildingFolderList = false;
            refreshScriptList();
        }

        /**
         * ファイル名リストから指定した名前の位置を探す
         * @param {string|null} fileName - 探すファイル名
         * @returns {number} 見つかった位置。見つからなければ 0（先頭）
         */
        function findScriptIndex(fileName) {
            if (!fileName) return 0;
            for (var i = 0; i < scriptListBox.items.length; i++) {
                if (scriptListBox.items[i].text === fileName) return i;
            }
            return 0;
        }

        /**
         * 選択中フォルダーに合わせて右のファイル名リストを組み直す
         * @returns {void}
         */
        function refreshScriptList() {
            var folderPath = selectedFolderPath();
            scriptListBox.removeAll();
            listedScripts = [];

            for (var i = 0; i < filteredScripts.length; i++) {
                if (folderPath !== null && filteredScripts[i].folderPath !== folderPath) continue;
                listedScripts.push(filteredScripts[i]);
                var scriptItem = scriptListBox.add("item", filteredScripts[i].fileName);
                scriptItem.scriptIndex = listedScripts.length - 1;
            }

            /* 件数はパネルのタイトルに出す / Show the match count in the panel title */
            keywordPanel.text = getLabel(LABELS.panel.keyword, [listedScripts.length]);

            /* 最初の組み立てだけは前回のセッションで選んでいたファイルを狙う / Aim at the remembered file on the first pass */
            var selectionIndex = findScriptIndex(restoringFileName);
            restoringFileName = null;

            if (scriptListBox.items.length > 0) {
                scriptListBox.selection = selectionIndex;
                btnRun.enabled = true;
            } else {
                btnRun.enabled = false;
            }
        }

        /**
         * 右のリストで選択中のスクリプトを返す
         * @returns {ScriptEntry|null} 選択中のスクリプト。未選択なら null
         */
        function selectedScriptEntry() {
            if (!scriptListBox.selection) return null;
            var selectedIndex = scriptListBox.selection.scriptIndex;
            if (selectedIndex === undefined) return null;
            return listedScripts[selectedIndex] || null;
        }

        /**
         * 今の絞り込みと選択を控える
         * 閉じたあとでは部品の値を読めないことがあるので、記録はここでまとめて行う
         * @returns {void}
         */
        function saveLauncherSession() {
            if (isSessionSaved) return;
            isSessionSaved = true;

            var savedState = { remember: rememberCheckbox.value };
            if (rememberCheckbox.value) {
                var chosenScript = selectedScriptEntry();

                /* 対象フォルダーも控える。別のフォルダーで開いたときに選択を引き継がないため
                   / The selections below only make sense against this folder */
                savedState.targetFolder = targetFolder.fsName;
                savedState.keyword = keywordInput.text;
                savedState.includeSubfolders = includeSubfoldersCheckbox.value;
                savedState.folderPath = selectedFolderPath();
                savedState.fileName = chosenScript ? chosenScript.fileName : null;
            } else {
                /* OFFのときは初期値を書き、次に開いたときは何も引き継がない / Store the defaults instead */
                savedState.targetFolder = null;
                savedState.keyword = "";
                savedState.includeSubfolders = INCLUDE_SUBFOLDERS_DEFAULT;
                savedState.showFullPath = SHOW_FULL_PATH_DEFAULT;
                savedState.folderPath = null;
                savedState.fileName = null;
            }
            saveSessionState(savedState);
        }

        /**
         * 今の絞り込みと選択を控えてからダイアログを閉じる
         * @param {number} [closeCode] - close() に渡す値
         * @returns {void}
         */
        function closeLauncher(closeCode) {
            saveLauncherSession();

            if (closeCode === undefined) launcherDialog.close();
            else launcherDialog.close(closeCode);
        }

        /**
         * 選択中のスクリプトを実行対象に確定してダイアログを閉じる
         * @returns {void}
         */
        function runSelectedScript() {
            var selectedScript = selectedScriptEntry();
            if (!selectedScript) return;
            dialogResult.action = "run";
            dialogResult.file = selectedScript.file;
            closeLauncher(1);
        }

        /**
         * 左のリストで選択中のフォルダーをFinderで開く（ダイアログは開いたまま）
         * 「すべて」と「ルート」は対象フォルダーそのものを開く
         * @returns {void}
         */
        function openSelectedFolder() {
            var folderPath = selectedFolderPath();
            var folderToOpen = targetFolder;

            /* パス文字列を組み立てず、そのフォルダーにあるスクリプトの親を使う / Use a real Folder object */
            if (folderPath) {
                for (var i = 0; i < filteredScripts.length; i++) {
                    if (filteredScripts[i].folderPath === folderPath) {
                        folderToOpen = filteredScripts[i].file.parent;
                        break;
                    }
                }
            }
            folderToOpen.execute();
        }

        /**
         * 選択中のスクリプトをFinderで表示する（ダイアログは開いたまま）
         * @returns {void}
         */
        function revealSelectedScript() {
            var selectedScript = selectedScriptEntry();
            if (selectedScript) revealScriptFile(selectedScript.file);
        }

        /* 検索キーは作成済みなので入力時は比較だけを行う。onChangingで即時反映し、onChangeでIME確定も拾う */
        var lastNormalizedQuery = null;

        /**
         * キーワードが変わったときだけリストを組み直す
         * @returns {void}
         */
        function refreshListIfQueryChanged() {
            var normalizedQuery = splitSearchTerms(keywordInput.text).join(" ");
            if (normalizedQuery === lastNormalizedQuery) return;
            lastNormalizedQuery = normalizedQuery;
            refreshFolderList();
        }

        /**
         * キーワードの有無に合わせてクリアボタンのディム表示を切り替える
         * @returns {void}
         */
        function updateClearButtonState() {
            var hasKeyword = trimWhitespace(keywordInput.text) !== "";
            if (keywordClearButton.enabled === hasKeyword) return;
            keywordClearButton.enabled = hasKeyword;
            keywordClearButton.notify("onDraw");
        }

        /**
         * キーワードが変わったときの処理をまとめる
         * @returns {void}
         */
        function handleKeywordChanged() {
            updateClearButtonState();
            refreshListIfQueryChanged();
        }

        keywordInput.onChanging = handleKeywordChanged;
        keywordInput.onChange = handleKeywordChanged;
        keywordClearButton.onClick = function () {
            keywordInput.text = "";
            /* 続けて打ち直せるよう、フォーカスは入力欄へ戻す / Put the focus back in the field */
            keywordInput.active = true;
            handleKeywordChanged();
        };
        btnPreferences.onClick = function () {
            var preferences = showPreferencesDialog(targetFolder, presetMinCount, presetMaxButtons, rememberCheckbox.value);
            if (!preferences) return;
            presetMinCount = preferences.minCount;
            presetMaxButtons = preferences.maxButtons;
            saveKeywordSettings(presetMinCount, presetMaxButtons);

            /* フォルダーが変わったら一覧を作り直すため、ダイアログを開き直す / Reopen the dialog to rescan */
            if (preferences.folder.fsName !== targetFolder.fsName) {
                dialogResult.action = "changeFolder";
                dialogResult.folder = preferences.folder;
                closeLauncher(2);
                return;
            }

            /* 置き場の高さが変わるので、ダイアログ全体を組み直す / The reserved height changes, so re-layout the dialog */
            reserveKeywordPresetHeight(keywordPresetContainer, presetMaxButtons);
            currentPresetWords = null;
            refreshFolderList();
            launcherDialog.layout.layout(true);
        };
        includeSubfoldersCheckbox.onClick = refreshFolderList;
        folderListBox.onChange = function () {
            if (isRebuildingFolderList) return;
            refreshScriptList();
        };
        scriptListBox.onChange = function () {
            btnRun.enabled = !!scriptListBox.selection;
        };
        folderListBox.onDoubleClick = openSelectedFolder;
        scriptListBox.onDoubleClick = function () {
            /* option+ダブルクリックは実行せずFinderで場所を開く / Option-double-click reveals instead of running */
            if (ScriptUI.environment.keyboardState.altKey) {
                revealSelectedScript();
                return;
            }
            runSelectedScript();
        };
        btnRun.onClick = runSelectedScript;
        btnCancel.onClick = function () {
            dialogResult.action = "cancel";
            closeLauncher();
        };
        bindLauncherKeys(launcherDialog, keywordInput, scriptListBox, runSelectedScript);

        /* セッションを見分けられない環境では覚えようがないので、操作させない / Nothing to remember without a session id */
        rememberCheckbox.enabled = !!readSessionId();

        /* ボタン以外の方法で閉じられても記録を残す / Cover any other way of dismissing the dialog */
        launcherDialog.onClose = function () {
            saveLauncherSession();
            return true;
        };

        refreshFolderList();
        lastNormalizedQuery = splitSearchTerms(keywordInput.text).join(" ");
        updateClearButtonState();
        launcherDialog.center();

        /* 検索欄を先頭の操作部品にしたうえで、表示時にも明示的にフォーカスする / Focus the keyword field on show */
        keywordInput.active = true;
        launcherDialog.onShow = function () {
            keywordInput.active = true;
            keywordInput.selection = [0, 0];
        };

        isDialogShown = true;
        launcherDialog.show();
        return dialogResult;
    }

    // =========================================
    // 実行 / Run
    // =========================================

    /**
     * Finder表示用のAutomatorアプリを探す
     * .app は実体がフォルダーなので File.exists が false を返す環境がある。Folder でも確認する
     * @returns {File|null} 実在するアプリ。無ければ null
     */
    function findRevealApp() {
        if ($.os.indexOf("Macintosh") === -1) return null;
        if (!new Folder(REVEAL_APP_PATH).exists && !new File(REVEAL_APP_PATH).exists) return null;
        return new File(REVEAL_APP_PATH);
    }

    /**
     * Finder表示用アプリに渡すパスを一時ファイルへ書き出す
     * @param {File} scriptFile - 対象のスクリプトファイル
     * @returns {boolean} 書き出せたら true
     */
    function writeRevealPath(scriptFile) {
        /* fsName で ~ ではなく絶対パスを渡す / fsName gives the absolute POSIX path */
        return writeTextFile(new File(REVEAL_PATH_FILE), scriptFile.fsName);
    }

    /**
     * スクリプトの場所をFinderで開く
     * @param {File} scriptFile - 対象のスクリプトファイル
     * @returns {void}
     */
    function revealScriptFile(scriptFile) {
        /* アプリが無い環境では囲みフォルダーを開くだけにとどめる / Fall back to the enclosing folder */
        var revealApp = findRevealApp();
        if (!revealApp) {
            scriptFile.parent.execute();
            return;
        }

        /* アプリは一時ファイルからパスを読むので、書き出してから起動する / The app reads the path from a temp file */
        if (writeRevealPath(scriptFile) && revealApp.execute()) return;

        /* 起動できなかった場合も何も起きないままにはしない / Never end up doing nothing */
        scriptFile.parent.execute();
    }

    /**
     * 選ばれたスクリプトを実行する
     * @param {File} scriptFile - 実行するスクリプトファイル
     * @returns {void}
     */
    function runScriptFile(scriptFile) {
        if (!scriptFile || !scriptFile.exists) {
            alert(getLabel(LABELS.alert.missingScript), getLabel(LABELS.dialog.title));
            return;
        }

        try {
            $.evalFile(scriptFile);
        } catch (e) {
            var errorMessage = getLabel(LABELS.alert.runFailed, [scriptFile.fsName, e.message]);
            if (e.line) errorMessage += getLabel(LABELS.alert.errorLine, [e.line]);
            alert(errorMessage, getLabel(LABELS.dialog.title));
        }
    }

    /**
     * 対象フォルダーの確認からスクリプト実行までを進める
     * @returns {void}
     */
    function main() {
        var scriptFolder = readSavedScriptFolder();

        while (true) {
            if (!scriptFolder || !scriptFolder.exists) {
                scriptFolder = chooseScriptFolder(scriptFolder);
                if (!scriptFolder) return;
                saveScriptFolder(scriptFolder);
            }

            var launcherResult = showLauncherDialog(scriptFolder);

            if (launcherResult.action === "changeFolder") {
                /* 環境設定で選び直した場合はそのフォルダーを使い、見つからなかった場合だけ選ばせる */
                var changedFolder = launcherResult.folder || chooseScriptFolder(scriptFolder);

                /* 選び直しを取り消したら終了する。対象が空のまま繰り返すと抜けられなくなる */
                if (!changedFolder) return;
                scriptFolder = changedFolder;
                saveScriptFolder(scriptFolder);
                continue;
            }

            if (launcherResult.action === "run") runScriptFile(launcherResult.file);
            break;
        }
    }

    main();

})();
