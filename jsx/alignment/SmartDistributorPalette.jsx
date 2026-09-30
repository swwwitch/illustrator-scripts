#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "smartDistributorPalette"

/*

### 概要

DistributeDownFromTop.jsx / DistributeUpFromTop.jsx を統合した常駐パレットです。
十字ボタン（↑ / ← 0 → / ↓）を押すたびに、その時点の選択へ1ステップぶん適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDistributorPalette.md

### Overview

A persistent palette that combines DistributeDownFromTop.jsx and DistributeUpFromTop.jsx.
Each press of the cross buttons (↑ / ← 0 → / ↓) applies one step to whatever is selected at that moment.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDistributorPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartDistributorPalette";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDistributorPalette.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDistributorPalette.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 概要 / Overview
    // =========================================
    /*
    SmartDistributorPalette.jsx

    DistributeDownFromTop / DistributeUpFromTop の統合パレット。
    十字ボタン（↑ / ← 0 → / ↓）を押すたびに、その時点の選択へ 1 ステップ適用する。
    実際のドキュメント操作は BridgeTalk でメインエンジンへ送って実行する（1 クリック = 取り消し 1 回）。

    複数オブジェクトを選択しているとき:
      縦（↑/↓）… 基準点で決めた端（既定=上端）を固定し、以降を「移動距離」ぶんずつ上下に等間隔調整
      横（←/→）… 基準点で決めた端（既定=左端）を固定し、以降を「移動距離」ぶんずつ左右に等間隔調整
      0       … オブジェクト間の隙間を 0 にする（基準点で固定端を決定。中央は並びの向きで自動判定）

    テキストを 1 つだけ選択しているとき:
      縦（↑/↓）… 行送りを「移動距離」ぶん、↓ で加算／↑ で減算
      横（←/→）・0 … 無効（ディム表示）

    Shift を押しながらボタンで 10 倍。
    Option + 矢印キーでも十字ボタンと同じ操作ができる。

    「基準点」パネル（5点の十字ラジオ）で固定する位置を指定する:
      ・上端/下端 … 縦方向の固定端（中段3点の縦成分は中央＝対称）
      ・左端/右端 … 横方向の固定端（上下点の横成分は中央）
      ・中央      … 縦横とも中央固定（対称分配・0 詰めは広がりの大きい軸）

    移動距離は「移動距離」パネルのラジオで選ぶ:
      ・環境設定のテキスト/行送り（text/sizeIncrement × 表示単位 text/units）
      ・環境設定のキー増加（cursorKeyLength）
      ・カスタム（pt 指定）
    いずれも内部では pt 換算して処理する。
    */

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 保存した設定が無いときの移動距離 / Distance source when nothing is saved
       "leading" = サイズ／行送り、"keyinput" = キー増加、"custom" = カスタム */
    var DEFAULT_DISTANCE_MODE = "leading";

    /* カスタムの移動距離の初期値（pt）。入力が数値でないときもこの値を使う
       Initial custom step in points; also used when the field is not a number */
    var DEFAULT_CUSTOM_STEP_PT = 0.1;

    /* 保存した設定が無いときの基準点 / Reference point when nothing is saved
       "top" / "bottom" / "left" / "right" / "center" */
    var DEFAULT_ANCHOR = "top";

    // =========================================
    // 保存ファイル / Saved files
    // =========================================

    /* パレット位置の保存先 / Where the palette position is stored */
    var POSITION_FILE = new File(Folder.userData + "/SmartDistributor/palette-position.txt");

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

    /* 設定の保存先は Folder.userData/illustrator-scripts/SmartDistributorPalette.json。
       旧版の SmartDistributor/settings.txt は最初の1回だけ読み継ぐ
       Settings live in illustrator-scripts/SmartDistributorPalette.json; the old settings.txt is read once */
    var settingsStore = createSettingsStore(SCRIPT_NAME, "persistent", {
        legacy: function () {
            return readSettingsLegacyFile(Folder.userData + "/SmartDistributor/settings.txt");
        }
    });

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

    var PALETTE_OPACITY = 0.97;             /* パレットの不透明度 / palette opacity */
    var ANCHOR_PANEL_SPACING = 4;           /* 基準点の行間 / spacing between the reference point rows */
    var ANCHOR_RADIO_GAP = 18;              /* 基準点の中段ラジオの間隔 / gap between the middle-row radios */
    var NUDGE_BUTTON_SIZE = [32, 24];       /* 十字ボタンの大きさ / size of the nudge buttons */
    var NUDGE_BUTTON_GAP = 4;               /* 十字ボタンの間隔 / gap between the nudge buttons */
    var CUSTOM_ROW_SPACING = 6;             /* カスタム行の間隔 / spacing inside the custom row */
    var CUSTOM_FIELD_CHARS = 5;             /* カスタム値の入力欄の幅（文字数）/ width of the custom field */

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
            title: { ja: "スマート均等配置", en: "Smart Distributor" }
        },
        panel: {
            anchor: { ja: "基準点", en: "Reference point" },
            distance: { ja: "移動距離", en: "Distance" }
        },
        radio: {
            sourceTextLeading: { ja: "サイズ/行送り（環境設定）", en: "Size/Leading (Pref)" },
            sourceKeyInput: { ja: "キー増加（環境設定）", en: "Keyboard Increment (Pref)" },
            sourceCustom: { ja: "カスタム", en: "Custom" }
        },
        tooltip: {
            nudgeUp: {
                ja: "基準点で決めた端を固定し、上方向へ「移動距離」ぶん広げます（Shiftで10倍）。テキスト1つの選択では行送りを詰めます。",
                en: "Spreads the objects upward by the distance, keeping the anchored edge fixed (Shift for x10). With a single text object it tightens the leading."
            },
            nudgeDown: {
                ja: "基準点で決めた端を固定し、下方向へ「移動距離」ぶん広げます（Shiftで10倍）。テキスト1つの選択では行送りを広げます。",
                en: "Spreads the objects downward by the distance, keeping the anchored edge fixed (Shift for x10). With a single text object it opens up the leading."
            },
            nudgeLeft: {
                ja: "基準点で決めた端を固定し、左方向へ「移動距離」ぶん広げます（Shiftで10倍）。",
                en: "Spreads the objects to the left by the distance, keeping the anchored edge fixed (Shift for x10)."
            },
            nudgeRight: {
                ja: "基準点で決めた端を固定し、右方向へ「移動距離」ぶん広げます（Shiftで10倍）。",
                en: "Spreads the objects to the right by the distance, keeping the anchored edge fixed (Shift for x10)."
            },
            collapse: {
                ja: "オブジェクト間の隙間を 0 にして密着させます。基準点で固定する端を決めます。",
                en: "Closes the gaps between the objects. The reference point decides which edge stays put."
            },
            sourceTextLeading: {
                ja: "環境設定［テキスト］の「サイズ／行送り」の値を1回ぶんの移動距離に使います。",
                en: "Uses the Size/Leading increment from the Type preferences as one step."
            },
            sourceKeyInput: {
                ja: "環境設定［一般］の「キー入力」の値を1回ぶんの移動距離に使います。",
                en: "Uses the Keyboard Increment from the General preferences as one step."
            },
            sourceCustom: { ja: "1回ぶんの移動距離を pt で直接指定します。", en: "Sets one step directly, in points." },
            anchorTop: { ja: "上端を固定", en: "Fix top" },
            anchorBottom: { ja: "下端を固定", en: "Fix bottom" },
            anchorLeft: { ja: "左端を固定", en: "Fix left" },
            anchorRight: { ja: "右端を固定", en: "Fix right" },
            anchorCenter: { ja: "中央を固定", en: "Fix center" },
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

    /**
     * 項目名と値の間に入れるコロンを返す（日本語は全角、英語は半角＋空白）
     * @returns {string} コロン
     */
    function labelColon() {
        return (uiLang === "ja") ? "：" : ": ";
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
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
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // UI の明暗（再利用パーツ） / UI theme (reusable)

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

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

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

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // =========================================
    // パレットの状態 / Palette state
    // =========================================

    /* パレットとコントロール（showPalette() の中で代入）/ Palette and controls, assigned in showPalette() */
    var paletteWindow = null;
    var anchorPanel = null;
    var anchorRadios = {};              /* 基準点キー → ラジオ / reference point key -> radio */
    var upButton = null;
    var downButton = null;
    var leftButton = null;
    var rightButton = null;
    var zeroButton = null;
    var sourceTextLeadingRadio = null;
    var sourceKeyInputRadio = null;
    var sourceCustomRadio = null;
    var customField = null;
    var customStepper = null;           /* カスタム値の∧∨ / stepper for the custom value */

    var currentAnchor = DEFAULT_ANCHOR; /* "top" / "bottom" / "left" / "right" / "center" */
    var workerInstalled = false;        /* BridgeTalk ワーカーを登録済みか / whether the worker is installed */

    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)

    /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
    var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

    /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
    var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

    /* 修飾キーの別名 / Aliases accepted for the modifiers */
    var KEY_SHORTCUT_MODIFIER_ALIASES = {
        SHIFT: "SHIFT",
        ALT: "ALT", OPTION: "ALT", OPT: "ALT",
        CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
    };

    /**
     * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
     * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
     * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
     */
    function normalizeKeyShortcutSpec(keySpec) {
        var specParts = String(keySpec).split("+");
        var baseKey = specParts.pop().toUpperCase();
        var modifierFlags = {};
        for (var i = 0; i < specParts.length; i++) {
            var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
            if (modifierName) modifierFlags[modifierName] = true;
        }
        return buildKeyShortcutSpec(modifierFlags, baseKey);
    }

    /**
     * 修飾キーの状態とキー名から照合用の表記を組み立てる
     * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
     * @param {string} baseKey - 大文字のキー名
     * @returns {string} 照合用の表記
     */
    function buildKeyShortcutSpec(modifierFlags, baseKey) {
        var specText = "";
        for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
            if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
        }
        return specText + baseKey;
    }

    /**
     * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
     * @param {Object} keyEvent - keydown イベント
     * @returns {string} 照合用の表記。キー名が無いときは空文字
     */
    function readKeyShortcutSpec(keyEvent) {
        if (!keyEvent || !keyEvent.keyName) return "";
        var keyboardState = {};
        try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
        var modifierFlags = {
            SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
            ALT: !!(keyEvent.altKey || keyboardState.altKey),
            CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
        };
        return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
    }

    /**
     * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
     * @param {Object} control - コントロール
     * @returns {boolean} 押せるなら true
     */
    function isKeyShortcutControlUsable(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false || node.visible === false) return false;
        }
        return true;
    }

    /**
     * キーを受けたコントロールが、文字を入力する欄か
     * @param {Object} focusedControl - イベントの発生元
     * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
     * @returns {boolean} 入力中としてショートカットを止めるなら true
     */
    function isKeyShortcutTypingTarget(focusedControl, numericFields) {
        if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
        for (var i = 0; i < numericFields.length; i++) {
            if (numericFields[i] === focusedControl) return false;
        }
        return true;
    }

    /**
     * コントロールをクリックしたときと同じ動作をする
     * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
     * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
     * @returns {void}
     */
    function pressKeyShortcutControl(control) {
        if (control.type === "radiobutton") {
            /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
            var siblings = control.parent ? control.parent.children : [];
            for (var i = 0; i < siblings.length; i++) {
                if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
            }
            control.value = true;
        } else if (control.type === "checkbox") {
            control.value = !control.value;
        }
        if (typeof control.onClick === "function") {
            control.onClick.call(control);
        } else if (control.type === "button" && typeof control.notify === "function") {
            /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
            control.notify("onClick");
        }
    }

    /**
     * 1つのショートカットを実行する
     * @param {Object|Function} shortcutTarget - コントロール、または関数
     * @param {Object} keyEvent - keydown イベント
     * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
     */
    function runKeyShortcutTarget(shortcutTarget, keyEvent) {
        var targetControl = shortcutTarget;
        if (typeof shortcutTarget === "function") {
            var runResult = shortcutTarget(keyEvent);
            if (runResult === false || runResult === null) return false;
            if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
            targetControl = runResult;
        }
        /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
        if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
        return true;
    }

    /**
     * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
     * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
     * @returns {string} 表示用の表記（"Shift+R" など）
     */
    function formatKeyShortcutLabel(normalizedSpec) {
        var isMac = ($.os.indexOf("Mac") === 0);
        var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
        var specParts = normalizedSpec.split("+");
        var baseKey = specParts.pop();
        var labelText = "";
        for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
        if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
        return labelText + baseKey;
    }

    /**
     * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
     * @param {Object} control - コントロール
     * @param {string} normalizedSpec - 照合用の表記
     * @returns {void}
     */
    function appendKeyShortcutToTip(control, normalizedSpec) {
        var keyLabel = formatKeyShortcutLabel(normalizedSpec);
        var currentTip = control.helpTip ? String(control.helpTip) : "";
        if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
        var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
        control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
    }

    /**
     * ダイアログ・パレットに文字キーのショートカットを付ける
     * @param {Window} targetWindow - キーを受けるダイアログ・パレット
     * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
     * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
     * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
     */
    function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
        var shortcutSettings = shortcutOptions || {};
        var numericFields = shortcutSettings.numericFields || [];
        var bindingTable = {};

        for (var keySpec in shortcutMap) {
            if (!shortcutMap.hasOwnProperty(keySpec)) continue;
            var mapEntry = shortcutMap[keySpec];
            if (!mapEntry) continue;
            var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
            var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
            bindingTable[normalizedSpec] = {
                target: isWrapped ? mapEntry.target : mapEntry,
                inFields: !!(isWrapped && mapEntry.inFields)
            };
            var tipControl = bindingTable[normalizedSpec].target;
            if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                appendKeyShortcutToTip(tipControl, normalizedSpec);
            }
        }

        /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
        targetWindow.addEventListener("keydown", function (keyEvent) {
            var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
            if (!binding) return;
            if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
            if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
            if (keyEvent.preventDefault) keyEvent.preventDefault();
            if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
        }, true);

        return bindingTable;
    }

    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts

    // =========================================
    // パレットの構築 / Palette construction
    // =========================================

    /**
     * 前回のパレットが開いていれば閉じる（多重起動ガード）
     * @returns {void}
     */
    function closeExistingPalette() {
        /* 前回のウィンドウが破棄済みだと close() が失敗することがある / close() may fail on a disposed window */
        try {
            if ($.global.smartDistributorWindow) {
                $.global.smartDistributorWindow.close();
                $.global.smartDistributorWindow = null;
            }
        } catch (e) {
        }
    }

    /**
     * パレットのウィンドウを作る
     * @returns {Window} 作ったパレット
     */
    function createPaletteWindow() {
        var newPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(newPalette);
        newPalette.opacity = PALETTE_OPACITY;
        return newPalette;
    }

    /**
     * 基準点のラジオを1つ追加し、anchorRadios に登録する
     * @param {Group} anchorRow - 追加先の行
     * @param {string} anchorKey - 基準点キー（"top" など）
     * @param {string} tooltipText - ツールチップ
     * @returns {void}
     */
    function addAnchorRadio(anchorRow, anchorKey, tooltipText) {
        var anchorRadio = anchorRow.add("radiobutton", undefined, "");
        anchorRadio.helpTip = tooltipText;
        anchorRadio.onClick = function () { selectAnchor(anchorKey); };
        anchorRadios[anchorKey] = anchorRadio;
    }

    /**
     * 基準点パネル（固定する位置を5点の十字から選ぶ）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} initialAnchor - 最初に選ぶ基準点キー
     * @returns {void}
     */
    function addAnchorPanel(parentGroup, initialAnchor) {
        anchorPanel = parentGroup.add("panel", undefined, getLabel("panel.anchor"));
        setupPanel(anchorPanel, ANCHOR_PANEL_SPACING);
        anchorPanel.alignChildren = ["center", "center"];  /* 十字を中央寄せ / center the cross */
        anchorPanel.alignment = ["left", "fill"];

        var anchorTopRow = anchorPanel.add("group");
        addAnchorRadio(anchorTopRow, "top", getLabel("tooltip.anchorTop"));

        var anchorMidRow = anchorPanel.add("group");
        anchorMidRow.spacing = ANCHOR_RADIO_GAP;
        addAnchorRadio(anchorMidRow, "left", getLabel("tooltip.anchorLeft"));
        addAnchorRadio(anchorMidRow, "center", getLabel("tooltip.anchorCenter"));
        addAnchorRadio(anchorMidRow, "right", getLabel("tooltip.anchorRight"));

        var anchorBottomRow = anchorPanel.add("group");
        addAnchorRadio(anchorBottomRow, "bottom", getLabel("tooltip.anchorBottom"));

        applyAnchorSelection(initialAnchor);  /* 初期選択（保存はしない）/ initial selection, not saved */
    }

    /**
     * 5点のうち1つだけを選択状態にする（別コンテナのため自動排他が効かない）
     * @param {string} anchorKey - 選ぶ基準点キー
     * @returns {void}
     */
    function applyAnchorSelection(anchorKey) {
        for (var radioKey in anchorRadios) {
            anchorRadios[radioKey].value = (radioKey === anchorKey);
        }
        currentAnchor = anchorKey;
    }

    /**
     * 基準点を選び、設定を保存する
     * @param {string} anchorKey - 選ぶ基準点キー
     * @returns {void}
     */
    function selectAnchor(anchorKey) {
        applyAnchorSelection(anchorKey);
        saveSettings();
    }

    /**
     * 十字ボタンを1つ追加する
     * @param {Group} buttonRow - 追加先の行
     * @param {string} buttonText - ボタンの表示文字（矢印・0）
     * @param {string} tooltipText - ツールチップ
     * @param {function} clickHandler - クリック時の処理
     * @returns {Button} 追加したボタン
     */
    function addNudgeButton(buttonRow, buttonText, tooltipText, clickHandler) {
        var nudgeButton = buttonRow.add("button", undefined, buttonText);
        nudgeButton.helpTip = tooltipText;
        nudgeButton.preferredSize = NUDGE_BUTTON_SIZE;
        nudgeButton.onClick = clickHandler;
        return nudgeButton;
    }

    /**
     * 十字ボタン（↑ / ← 0 → / ↓）を追加する
     * dx / dy は座標差分（→ / ↓ を正）
     * @param {Group} parentGroup - 追加先のグループ
     * @returns {void}
     */
    function addNudgePad(parentGroup) {
        var nudgePadGroup = parentGroup.add("group");
        nudgePadGroup.orientation = "column";
        nudgePadGroup.alignChildren = ["center", "center"];
        nudgePadGroup.alignment = ["fill", "center"];
        nudgePadGroup.spacing = NUDGE_BUTTON_GAP;

        var nudgeTopRow = nudgePadGroup.add("group");
        upButton = addNudgeButton(nudgeTopRow, "↑", getLabel("tooltip.nudgeUp"), function () { applyNudge(0, -1); });

        var nudgeMidRow = nudgePadGroup.add("group");
        nudgeMidRow.spacing = NUDGE_BUTTON_GAP;
        leftButton = addNudgeButton(nudgeMidRow, "←", getLabel("tooltip.nudgeLeft"), function () { applyNudge(-1, 0); });
        zeroButton = addNudgeButton(nudgeMidRow, "0", getLabel("tooltip.collapse"), function () { collapseSpacing(); });
        rightButton = addNudgeButton(nudgeMidRow, "→", getLabel("tooltip.nudgeRight"), function () { applyNudge(1, 0); });

        var nudgeBottomRow = nudgePadGroup.add("group");
        downButton = addNudgeButton(nudgeBottomRow, "↓", getLabel("tooltip.nudgeDown"), function () { applyNudge(0, 1); });
    }

    /**
     * 移動距離パネル（ラジオ3択とカスタム値）を十字ボタンの下に追加する
     * @param {Window} parentWindow - 追加先のパレット
     * @param {string} initialMode - 最初に選ぶ移動距離（"leading" / "keyinput" / "custom"）
     * @param {string} initialCustom - カスタム値の初期表示
     * @returns {void}
     */
    function addDistancePanel(parentWindow, initialMode, initialCustom) {
        var distancePanel = parentWindow.add("panel", undefined, getLabel("panel.distance"));
        setupPanel(distancePanel, 6);
        distancePanel.alignChildren = ["left", "top"];  /* ラジオは左寄せ / left-align radios */

        sourceTextLeadingRadio = distancePanel.add("radiobutton", undefined, getLabel("radio.sourceTextLeading"));
        sourceTextLeadingRadio.helpTip = getLabel("tooltip.sourceTextLeading");
        sourceKeyInputRadio = distancePanel.add("radiobutton", undefined, getLabel("radio.sourceKeyInput"));
        sourceKeyInputRadio.helpTip = getLabel("tooltip.sourceKeyInput");
        refreshSourceLabels();  /* 環境設定の現在値をラベルに反映 / show the current preference values */

        var customRow = distancePanel.add("group");
        customRow.spacing = CUSTOM_ROW_SPACING;
        sourceCustomRadio = customRow.add("radiobutton", undefined, getLabel("radio.sourceCustom"));
        sourceCustomRadio.helpTip = getLabel("tooltip.sourceCustom");
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var customStepperGroup = customRow.add("group");
        customStepperGroup.orientation = "row";
        customStepperGroup.alignChildren = ["left", "center"];
        customStepperGroup.spacing = 0;
        customStepperGroup.margins = 0;
        customStepper = addStepper(customStepperGroup, function () { return customField; }, {
            step: 1, min: 0,
            onStep: function () { saveSettings(); }
        });
        customField = customStepperGroup.add("edittext", undefined, initialCustom);
        customField.helpTip = getLabel("tooltip.sourceCustom");
        customField.characters = CUSTOM_FIELD_CHARS;
        bindSteppedArrowKeys(customField, customStepper);
        customRow.add("statictext", undefined, "pt");

        /* 保存値からラジオの初期状態を復元 / Restore the radios from the saved mode */
        sourceTextLeadingRadio.value = (initialMode === "leading");
        sourceKeyInputRadio.value = (initialMode === "keyinput");
        sourceCustomRadio.value = (initialMode === "custom");
        if (!sourceTextLeadingRadio.value && !sourceKeyInputRadio.value && !sourceCustomRadio.value) {
            sourceTextLeadingRadio.value = true;
        }

        /* カスタムだけ別コンテナ（customRow）にあり ScriptUI の自動排他が効かないので手動で排他にする
           The custom radio lives in its own row, so exclusivity is handled by hand */
        sourceTextLeadingRadio.onClick = function () { selectDistanceMode(sourceTextLeadingRadio); };
        sourceKeyInputRadio.onClick = function () { selectDistanceMode(sourceKeyInputRadio); };
        sourceCustomRadio.onClick = function () { selectDistanceMode(sourceCustomRadio); };
        customField.onChange = function () { saveSettings(); };
        setCustomFieldEnabled(sourceCustomRadio.value);
    }

    /**
     * カスタム値の入力欄と∧∨の有効／無効をまとめて切り替える
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setCustomFieldEnabled(isEnabled) {
        customField.enabled = isEnabled;
        customStepper.enabled = isEnabled;
        redrawSteppersIn(customStepper);
    }

    /**
     * 移動距離のラジオを1つだけ選び、カスタム値の入力欄を切り替えて設定を保存する
     * @param {RadioButton} selectedRadio - 選んだラジオ
     * @returns {void}
     */
    function selectDistanceMode(selectedRadio) {
        sourceTextLeadingRadio.value = (selectedRadio === sourceTextLeadingRadio);
        sourceKeyInputRadio.value = (selectedRadio === sourceKeyInputRadio);
        sourceCustomRadio.value = (selectedRadio === sourceCustomRadio);
        setCustomFieldEnabled(sourceCustomRadio.value);  /* カスタム以外はディム / dim unless custom */
        saveSettings();
    }

    /**
     * 移動距離ラジオのラベルに環境設定の現在値を表示する（表示単位込み）
     * @returns {void}
     */
    function refreshSourceLabels() {
        /* 行送りは text/units 単位の値そのまま / the leading increment is already in text/units */
        var leadingValue = app.preferences.getRealPreference("text/sizeIncrement");
        sourceTextLeadingRadio.text = getLabel("radio.sourceTextLeading") + labelColon() + leadingValue + getUnitInfo("text/units").label;

        /* cursorKeyLength は pt で返るので一般単位（rulerType）へ換算して表示 / cursorKeyLength is in points */
        var rulerUnit = getUnitInfo("rulerType");
        var keyValue = app.preferences.getRealPreference("cursorKeyLength") / rulerUnit.pointsPerUnit;
        keyValue = Math.round(keyValue * 1000) / 1000;
        sourceKeyInputRadio.text = getLabel("radio.sourceKeyInput") + labelColon() + keyValue + rulerUnit.label;
    }

    /**
     * Option + 矢印キーで十字ボタンと同じ操作をする（カスタム値の編集中は入力欄に任せる）
     * @param {KeyboardEvent} keyEvent - キーイベント
     * @returns {void}
     */
    function onPaletteKeyDown(keyEvent) {
        if (!keyEvent.altKey) return;
        if (keyEvent.target === customField) return;
        if (keyEvent.keyName === "Up" && upButton.enabled) {
            applyNudge(0, -1);
            keyEvent.preventDefault();
        } else if (keyEvent.keyName === "Down" && downButton.enabled) {
            applyNudge(0, 1);
            keyEvent.preventDefault();
        } else if (keyEvent.keyName === "Left" && leftButton.enabled) {
            applyNudge(-1, 0);
            keyEvent.preventDefault();
        } else if (keyEvent.keyName === "Right" && rightButton.enabled) {
            applyNudge(1, 0);
            keyEvent.preventDefault();
        }
    }

    /**
     * パレットの移動・終了・アクティブ化・キー操作のイベントを登録する
     * @returns {void}
     */
    function attachPaletteEvents() {
        paletteWindow.onMove = function () {
            saveWindowPosition(paletteWindow);
        };

        paletteWindow.onClose = function () {
            saveWindowPosition(paletteWindow);
            saveSettings();
            $.global.smartDistributorWindow = null;
        };

        /* 選択が変わるたびにボタンの有効状態とラベルを更新 / Refresh buttons and labels on activation */
        paletteWindow.onActivate = function () {
            updateButtonStates();
            refreshSourceLabels();
        };

        paletteWindow.addEventListener("keydown", onPaletteKeyDown);

        /* パレットは Esc で閉じないので、閉じる処理を割り当てる（入力中も効かせる）
           Palettes do not close on Esc by themselves; map it to close, also while typing */
        addKeyShortcuts(paletteWindow, {
            "Escape": { target: function () { paletteWindow.close(); }, inFields: true }
        });
    }

    // =========================================
    // ボタン状態・呼び出し / Button state and actions
    // =========================================

    /**
     * 現在の選択を返す
     * @returns {Array|null} 選択（ドキュメントが無ければ null）
     */
    function getCurrentSelection() {
        if (app.documents.length < 1) return null;
        return app.activeDocument.selection;
    }

    /**
     * テキストを1つだけ選択しているかを返す
     * @param {Array|null} selectedObjects - 選択
     * @returns {boolean} テキスト1つなら true
     */
    function isSingleTextSelection(selectedObjects) {
        return selectedObjects && selectedObjects.length === 1 && selectedObjects[0].typename === "TextFrame";
    }

    /**
     * 選択内容に応じてボタンの有効・無効（ディム表示）を更新する
     * @returns {void}
     */
    function updateButtonStates() {
        var selectedObjects = getCurrentSelection();
        var hasSingleText = isSingleTextSelection(selectedObjects);
        var hasMultiple = selectedObjects && selectedObjects.length >= 2;

        /* 縦：テキスト1つ または 複数選択で有効 / vertical: single text or multiple objects */
        upButton.enabled = hasSingleText || hasMultiple;
        downButton.enabled = hasSingleText || hasMultiple;
        /* 横・0：複数選択のときだけ有効 / horizontal and 0: multiple objects only */
        leftButton.enabled = hasMultiple;
        rightButton.enabled = hasMultiple;
        zeroButton.enabled = hasMultiple;
        /* 基準点は複数選択時のみ／テキスト1つではキー増加も使えないのでディム
           Reference point needs multiple objects; the key increment is dimmed for a single text */
        anchorPanel.enabled = hasMultiple;
        sourceKeyInputRadio.enabled = !hasSingleText;
    }

    /**
     * 選択したラジオに応じた1回ぶんの移動距離を返す
     * @returns {number} 移動距離（pt）
     */
    function getStepPoints() {
        if (sourceKeyInputRadio.value) {
            return app.preferences.getRealPreference("cursorKeyLength");
        }
        if (sourceCustomRadio.value) {
            /* カスタムは pt 指定 / the custom value is in points */
            var customStepPt = parseFloat(customField.text);
            if (isNaN(customStepPt)) customStepPt = DEFAULT_CUSTOM_STEP_PT;
            return customStepPt;
        }
        /* 既定：環境設定のテキスト/行送り（表示単位込みで pt 換算）/ default: the Size/Leading increment in points */
        return app.preferences.getRealPreference("text/sizeIncrement") * getUnitInfo("text/units").pointsPerUnit;
    }

    /**
     * 矢印1回ぶんをワーカーへ送って適用する（Shift で 10 倍）
     * 選択の検証はワーカー側に任せる（collapse と同じ流れ）
     * @param {number} dx - 横方向（→ が 1、← が -1）
     * @param {number} dy - 縦方向（↓ が 1、↑ が -1）
     * @returns {void}
     */
    function applyNudge(dx, dy) {
        var multiplier = 1;
        /* keyboardState は環境によって例外になるため保護 / keyboardState may throw on some setups */
        try {
            if (ScriptUI.environment.keyboardState.shiftKey) multiplier = 10;
        } catch (e) {
        }
        callWorker("nudge", dx, dy, getStepPoints() * multiplier, currentAnchor);
    }

    /**
     * オブジェクト間の隙間を 0 にする
     * @returns {void}
     */
    function collapseSpacing() {
        callWorker("collapse", 0, 0, 0, currentAnchor);
    }

    // =========================================
    // BridgeTalk でメインエンジンへ送って実行 / Dispatch to the main engine via BridgeTalk
    // パレットエンジンからの直接編集は不安定なため、実処理はメインエンジン側で行う
    // =========================================

    var MAX_PENDING_JOBS = 12;  /* 保持する BridgeTalk ジョブの上限 / cap on pending BridgeTalk jobs */

    /**
     * ワーカーを（必要なら登録してから）短い呼び出しで実行する
     * @param {string} workerAction - "nudge" / "collapse"
     * @param {number} dx - 横方向
     * @param {number} dy - 縦方向
     * @param {number} step - 1回ぶんの移動距離（pt）
     * @param {string} anchor - 基準点キー
     * @returns {void}
     */
    function callWorker(workerAction, dx, dy, step, anchor) {
        if (typeof BridgeTalk === "undefined") return;
        if (!workerInstalled) installWorker();
        /* 不正な基準点は top にフォールバック / Fall back to top for invalid anchors */
        var safeAnchor = (anchor === "top" || anchor === "bottom" || anchor === "left"
            || anchor === "right" || anchor === "center") ? anchor : "top";
        sendWorkerCall('$.global.smartDistributorWorker("' + workerAction + '",' + dx + ',' + dy + ',' + step + ',"' + safeAnchor + '");', true);
    }

    /**
     * ワーカー本体をメインエンジンへ一度だけ登録する
     * @returns {void}
     */
    function installWorker() {
        var bridgeTalk = createIllustratorBridgeTalk();
        bridgeTalk.body = "$.global.smartDistributorWorker = (" + getWorkerSource() + ");";
        bridgeTalk.onResult = function () { removePendingJob(bridgeTalk); };
        bridgeTalk.onError = function () { removePendingJob(bridgeTalk); };
        $.global.smartDistributorJobs.push(bridgeTalk);
        trimPendingJobs();
        bridgeTalk.send();
        workerInstalled = true;
    }

    /**
     * ワーカー関数のソースを返す（toString() が関数の後ろのコメントまで取り込むので、終わりの目印で切り詰める）
     * @returns {string} 送信する関数のソース
     */
    function getWorkerSource() {
        var workerSource = adjustSelectionWorker.toString();
        var endMarkerIndex = workerSource.indexOf("/* @@smart-distributor-worker-end@@ */");
        if (endMarkerIndex < 0) {
            return workerSource;
        }
        return workerSource.substring(0, endMarkerIndex) + "}";
    }

    /**
     * 登録済みのワーカーを短い呼び出しで実行する（失敗時は1回だけ再登録して送り直す）
     * @param {string} callBody - メインエンジンで評価する呼び出し
     * @param {boolean} allowReinstall - 失敗時に再登録するか
     * @returns {void}
     */
    function sendWorkerCall(callBody, allowReinstall) {
        var bridgeTalk = createIllustratorBridgeTalk();
        bridgeTalk.body = callBody;
        bridgeTalk.onResult = function () { removePendingJob(bridgeTalk); };
        bridgeTalk.onError = function () {
            removePendingJob(bridgeTalk);
            if (allowReinstall) {
                workerInstalled = false;
                installWorker();
                sendWorkerCall(callBody, false);
            }
        };
        $.global.smartDistributorJobs.push(bridgeTalk);
        trimPendingJobs();
        bridgeTalk.send();
    }

    /**
     * Illustrator 宛ての BridgeTalk を作る
     * @returns {BridgeTalk} 宛先を設定した BridgeTalk
     */
    function createIllustratorBridgeTalk() {
        var bridgeTalk = new BridgeTalk();
        /* 指定子を引けない環境では素の名前で送る / fall back to the bare name when no specifier is found */
        try {
            bridgeTalk.target = BridgeTalk.getSpecifier("illustrator");
        } catch (e) {
            bridgeTalk.target = "illustrator";
        }
        return bridgeTalk;
    }

    /**
     * 保留中のジョブを MAX_PENDING_JOBS 件までに保つ
     * @returns {void}
     */
    function trimPendingJobs() {
        var pendingJobs = $.global.smartDistributorJobs;
        while (pendingJobs.length > MAX_PENDING_JOBS) pendingJobs.shift();
    }

    /**
     * 完了したジョブを保留リストから除く
     * @param {BridgeTalk} bridgeTalk - 完了したジョブ
     * @returns {void}
     */
    function removePendingJob(bridgeTalk) {
        var pendingJobs = $.global.smartDistributorJobs;
        for (var i = pendingJobs.length - 1; i >= 0; i--) {
            if (pendingJobs[i] === bridgeTalk) pendingJobs.splice(i, 1);
        }
    }

    // =========================================
    // メインエンジンで実行される本体（文字列化して登録）/ Worker run in the main engine (registered as a string)
    // ※ 外側の変数を参照せず、引数とアプリ DOM だけで完結させる
    // ※ toString() で送るため JSDoc を付けず、関数内のコメントは /* */ だけにする
    //    Sent through toString(): no JSDoc, and only block comments inside
    // ※ toString() は関数の後ろのコメントまで取り込むので、末尾の目印 @@smart-distributor-worker-end@@ で切り詰めて送る（消さないこと）
    // =========================================

    /* 選択の間隔・行送りを1ステップ調整する / Adjust the spacing or leading of the selection by one step
       workerAction: "nudge" / "collapse"、anchor: "top" / "bottom" / "left" / "right" / "center"（固定する基準点） */
    function adjustSelectionWorker(workerAction, dx, dy, step, anchor) {
        var selectedItems;

        /* 配列へ写してから並べ替える / Copy into an array, then sort */
        function sortItemsBy(targetItems, compareItems) {
            var sortedItems = [];
            for (var i = 0; i < targetItems.length; i++) sortedItems.push(targetItems[i]);
            sortedItems.sort(compareItems);
            return sortedItems;
        }

        /* 並べ替えは見た目のボックス（geometricBounds）で判定する。TextFrame の position はベースライン（≒下端）を
           指すため、position で並べると上端側の固定要素が見た目の最上段とずれる（top / left の分配が崩れる原因） */
        /* 上→下（geometricBounds[1] = 上端 y）/ top to bottom */
        function sortTopToBottom(targetItems) {
            return sortItemsBy(targetItems, function (itemA, itemB) { return itemB.geometricBounds[1] - itemA.geometricBounds[1]; });
        }

        /* 左→右（geometricBounds[0] = 左端 x）/ left to right */
        function sortLeftToRight(targetItems) {
            return sortItemsBy(targetItems, function (itemA, itemB) { return itemA.geometricBounds[0] - itemB.geometricBounds[0]; });
        }

        /* 詰める向きを決める（中央は position の広がりが大きい軸）/ Pick the axis; center uses the wider spread */
        function resolveCollapseAxis() {
            if (anchor === "top" || anchor === "bottom") return "v";
            if (anchor === "left" || anchor === "right") return "h";
            var minX = null, maxX = null, minY = null, maxY = null;
            for (var i = 0; i < selectedItems.length; i++) {
                var itemPosition = selectedItems[i].position;
                if (minX === null || itemPosition[0] < minX) minX = itemPosition[0];
                if (maxX === null || itemPosition[0] > maxX) maxX = itemPosition[0];
                if (minY === null || itemPosition[1] < minY) minY = itemPosition[1];
                if (maxY === null || itemPosition[1] > maxY) maxY = itemPosition[1];
            }
            return ((maxX - minX) > (maxY - minY)) ? "h" : "v";
        }

        /* 縦の隙間を 0 に詰める / Close the vertical gaps */
        function collapseVertical() {
            var itemsTopToBottom = sortTopToBottom(selectedItems);
            var heights = [], totalHeight = 0, maxTop = null, minBottom = null;
            for (var i = 0; i < itemsTopToBottom.length; i++) {
                var bounds = itemsTopToBottom[i].geometricBounds;
                heights.push(bounds[1] - bounds[3]);
                totalHeight += bounds[1] - bounds[3];
                if (maxTop === null || bounds[1] > maxTop) maxTop = bounds[1];
                if (minBottom === null || bounds[3] < minBottom) minBottom = bounds[3];
            }
            /* 入れ子三項は必ず括弧で右結合を明示する（ExtendScript は括弧なしだと左結合に誤評価する） */
            var nextTop = (anchor === "bottom") ? (minBottom + totalHeight)
                : ((anchor === "top") ? maxTop : ((maxTop + minBottom) / 2 + totalHeight / 2));
            for (var j = 0; j < itemsTopToBottom.length; j++) {
                itemsTopToBottom[j].translate(0, nextTop - itemsTopToBottom[j].geometricBounds[1]);
                nextTop = nextTop - heights[j];
            }
        }

        /* 横の隙間を 0 に詰める / Close the horizontal gaps */
        function collapseHorizontal() {
            var itemsLeftToRight = sortLeftToRight(selectedItems);
            var widths = [], totalWidth = 0, minLeft = null, maxRight = null;
            for (var i = 0; i < itemsLeftToRight.length; i++) {
                var bounds = itemsLeftToRight[i].geometricBounds;
                widths.push(bounds[2] - bounds[0]);
                totalWidth += bounds[2] - bounds[0];
                if (minLeft === null || bounds[0] < minLeft) minLeft = bounds[0];
                if (maxRight === null || bounds[2] > maxRight) maxRight = bounds[2];
            }
            var nextLeft = (anchor === "right") ? (maxRight - totalWidth)
                : ((anchor === "left") ? minLeft : ((minLeft + maxRight) / 2 - totalWidth / 2));
            for (var j = 0; j < itemsLeftToRight.length; j++) {
                itemsLeftToRight[j].translate(nextLeft - itemsLeftToRight[j].geometricBounds[0], 0);
                nextLeft = nextLeft + widths[j];
            }
        }

        /* 隙間を 0 に詰める（基準点で固定端を決定）/ Close the gaps; the anchor decides the fixed edge */
        function collapseGaps() {
            if (selectedItems.length < 2) return "need 2+";
            if (resolveCollapseAxis() === "v") {
                collapseVertical();
            } else {
                collapseHorizontal();
            }
            return "ok";
        }

        /* テキスト1つの行送り増減（横は無効）/ Change the leading of a single text (vertical only) */
        function shiftTextLeading() {
            if (dy === 0) return "h disabled for text";
            var charAttributes = selectedItems[0].textRange.characterAttributes;
            charAttributes.leading = charAttributes.leading + dy * step;
            return "ok";
        }

        /* 1つの軸に沿って分配する（端固定＝矢印方向／中央＝対称）/ Distribute along one axis */
        function distributeAlongAxis(isVertical) {
            var sortedItems = isVertical ? sortTopToBottom(selectedItems) : sortLeftToRight(selectedItems);
            var startAnchor = isVertical ? "top" : "left";
            var endAnchor = isVertical ? "bottom" : "right";
            var isEdgeAnchor = (anchor === startAnchor || anchor === endAnchor);
            /* 入れ子三項は必ず括弧で右結合を明示する（括弧なしだと左結合に誤評価し、端が中央扱いになる） */
            var pivotIndex = (anchor === startAnchor) ? 0 : ((anchor === endAnchor) ? (sortedItems.length - 1) : (sortedItems.length - 1) / 2);
            for (var i = 0; i < sortedItems.length; i++) {
                var stepCount = i - pivotIndex;
                if (isEdgeAnchor) stepCount = Math.abs(stepCount);
                if (isVertical) {
                    sortedItems[i].translate(0, -stepCount * dy * step);
                } else {
                    sortedItems[i].translate(stepCount * dx * step, 0);
                }
            }
        }

        /* 例外はメインエンジンへ持ち込まず、文字列で返す / Report errors as strings instead of throwing */
        try {
            if (!app.documents || app.documents.length < 1) return "no document";
            selectedItems = app.activeDocument.selection;
            if (!selectedItems || selectedItems.length < 1) return "no selection";

            var workerResult;
            if (workerAction === "collapse") {
                workerResult = collapseGaps();
            } else if (selectedItems.length === 1 && selectedItems[0].typename === "TextFrame") {
                workerResult = shiftTextLeading();
            } else if (selectedItems.length < 2) {
                return "need 2+";
            } else {
                if (dy !== 0) distributeAlongAxis(true);
                if (dx !== 0) distributeAlongAxis(false);
                workerResult = "ok";
            }
            app.redraw();
            return workerResult;
        } catch (e) {
            return "error: " + e.message;
        }
        /* @@smart-distributor-worker-end@@ */
    }

    // =========================================
    // ウィンドウ位置の保存・復元 / Window position
    // =========================================

    /**
     * 保存したウィンドウ位置を復元する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function restoreWindowPosition(targetWindow) {
        if (!POSITION_FILE.exists) return;
        try {
            if (!POSITION_FILE.open("r")) return;
            var savedPositionText = POSITION_FILE.read();
            POSITION_FILE.close();

            var positionParts = savedPositionText.split(",");
            if (positionParts.length !== 2) return;

            var savedX = parseInt(positionParts[0], 10);
            var savedY = parseInt(positionParts[1], 10);
            if (isNaN(savedX) || isNaN(savedY)) return;

            targetWindow.location = [savedX, savedY];
        } catch (e) {
        }
    }

    /**
     * 現在のウィンドウ位置を保存する
     * @param {Window} targetWindow - 対象のウィンドウ
     * @returns {void}
     */
    function saveWindowPosition(targetWindow) {
        try {
            var positionFolder = POSITION_FILE.parent;
            if (!positionFolder.exists) positionFolder.create();
            if (!POSITION_FILE.open("w")) return;
            POSITION_FILE.write(targetWindow.location[0] + "," + targetWindow.location[1]);
            POSITION_FILE.close();
        } catch (e) {
        }
    }

    // =========================================
    // 設定の保存・復元 / Settings
    // =========================================

    /**
     * 現在の設定を保存する
     * @returns {void}
     */
    function saveSettings() {
        var distanceMode = sourceKeyInputRadio.value ? "keyinput"
            : (sourceCustomRadio.value ? "custom" : "leading");
        settingsStore.save({ mode: distanceMode, custom: customField.text, anchor: currentAnchor });
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 保存した設定を読み込み、前回のパレットを閉じてから新しいパレットを表示する
     * @returns {void}
     */
    function showPalette() {
        var savedSettings = settingsStore.load({
            mode: DEFAULT_DISTANCE_MODE,
            custom: String(DEFAULT_CUSTOM_STEP_PT),
            anchor: DEFAULT_ANCHOR
        });
        var initialMode = savedSettings.mode;
        var initialCustom = savedSettings.custom;
        var initialAnchor = savedSettings.anchor;
        currentAnchor = initialAnchor;

        /* BridgeTalk ジョブの保持先 / Holder for pending BridgeTalk jobs */
        if (!$.global.smartDistributorJobs) {
            $.global.smartDistributorJobs = [];
        }

        closeExistingPalette();

        paletteWindow = createPaletteWindow();
        $.global.smartDistributorWindow = paletteWindow;

        /* 2カラム：左＝基準点 / 右＝十字ボタン / Two columns: reference point | nudge buttons */
        var mainRowGroup = paletteWindow.add("group");
        mainRowGroup.orientation = "row";
        mainRowGroup.alignChildren = ["fill", "center"];
        mainRowGroup.spacing = COLUMN_SPACING;

        addAnchorPanel(mainRowGroup, initialAnchor);
        addNudgePad(mainRowGroup);
        addDistancePanel(paletteWindow, initialMode, initialCustom);

        restoreWindowPosition(paletteWindow);
        attachPaletteEvents();

        updateButtonStates();
        paletteWindow.show();

        /* 起動直後にワーカーを登録（最初のクリックも短い呼び出しだけで済む）/ Install the worker at startup */
        if (typeof BridgeTalk !== "undefined") {
            installWorker();
        }
    }

    showPalette();

})();
