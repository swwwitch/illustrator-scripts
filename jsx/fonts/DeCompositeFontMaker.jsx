#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの合成フォントを、文字ごとに構成フォントへ置き換えます。サイズ・比率・ベースラインも写すので、見た目はそのままです。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeCompositeFontMaker.md

### 注意

- 合成フォントのファイルがこの環境に無いと置き換えられません
- 特例文字セットは無視します

### Overview

Replaces the composite fonts in the selected text with their component fonts, character by character. Size, scale and baseline carry over, so the text keeps its look.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeCompositeFontMaker.md

### Notes

- Composite fonts whose file is not on this machine cannot be replaced
- Custom character sets are ignored

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DeCompositeFontMaker";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeCompositeFontMaker.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeCompositeFontMaker.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 自動行送りの文字は、置き換え前の行送りを固定値で残す（サイズが変わると自動行送りの値も変わるため）
       Pin auto-leading characters to their current leading, since resized characters would change it */
    var KEEP_LEADING = true;

    /* トラッキング・カーニングを、置き換え後の文字サイズに合わせて換算する（1/1000em 単位のため）
       Rescale tracking and kerning to the new size, since both are in 1/1000 em */
    var KEEP_SPACING = true;

    // =========================================
    // 合成フォントのファイル形式 / Composite font file format
    // =========================================
    /* 構成は CompositeFontMaker.jsx を参照。CFMA にサイズ・水平比率・垂直比率（16.16固定小数点）、
       PS の usematrix にベースライン、bfrange に文字セットの範囲が入る
       See CompositeFontMaker.jsx. CFMA holds size / h-scale / v-scale (16.16 fixed),
       the PS usematrix holds the baseline and bfrange holds each set's ranges */
    var STANDARD_CHARSET_COUNT = 6; /* 漢字・かな・全角約物・全角記号・半角欧文・半角数字。以降は特例文字 / standard sets; the rest are custom */

    // =========================================
    // ローカライズ / Localization
    // =========================================
    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "テキストを選択してください。", en: "Select some text." },
            noComposite: { ja: "選択したテキストに合成フォントは使われていません。", en: "The selected text uses no composite fonts." },
            noFolder: { ja: "合成フォントのフォルダーが見つかりません。", en: "The composite font folder was not found." },
            done: { ja: "%1文字を構成フォントに置き換えました。", en: "Replaced %1 characters with component fonts." },
            fileNotFound: {
                ja: "合成フォント「%1」のファイルが見つからないため、%2文字をそのままにしました。",
                en: "The file for the composite font \"%1\" was not found; %2 characters were left as they are."
            },
            xmpMembers: { ja: "ドキュメントに記録された構成フォント：", en: "Component fonts recorded in the document:" },
            missingFont: {
                ja: "構成フォント「%1」がこの環境に無いため、%2文字をそのままにしました。",
                en: "The component font \"%1\" is not installed; %2 characters were left as they are."
            }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "alert.done" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    // =========================================
    // 合成フォントの読み込み / Reading composite fonts
    // =========================================
    /**
     * バイナリ文字列の位置から16ビット符号なし整数を読む（ビッグエンディアン）
     * @param {string} fontBinary - バイナリ文字列
     * @param {number} offset - 位置
     * @returns {number} 値
     */
    function readUint16(fontBinary, offset) {
        return (fontBinary.charCodeAt(offset) << 8) | fontBinary.charCodeAt(offset + 1);
    }

    /**
     * バイナリ文字列の位置から32ビット符号なし整数を読む（ビッグエンディアン）
     * @param {string} fontBinary - バイナリ文字列
     * @param {number} offset - 位置
     * @returns {number} 値
     */
    function readUint32(fontBinary, offset) {
        return readUint16(fontBinary, offset) * 65536 + readUint16(fontBinary, offset + 2);
    }

    /**
     * ファイルをバイナリ文字列として読む
     * @param {File} fontFile - 読むファイル
     * @returns {string} 中身。読めなければ空文字
     */
    function readBinaryFile(fontFile) {
        fontFile.encoding = "BINARY";
        if (!fontFile.open("r")) return "";
        var fontBinary = fontFile.read();
        fontFile.close();
        return fontBinary;
    }

    /**
     * Illustrator の合成フォントフォルダーを探す
     * @returns {Folder|null} 見つかったフォルダー。なければ null
     */
    function findCompositeFontFolder() {
        var majorVersion = parseInt(app.version, 10);
        var localeFolderPath = Folder.userData.fsName + "/Adobe/Adobe Illustrator " + majorVersion + "/" + app.locale;
        var folderNames = ["合成フォント", "Composite Fonts"];
        for (var i = 0; i < folderNames.length; i++) {
            var candidate = new Folder(localeFolderPath + "/" + folderNames[i]);
            if (candidate.exists) return candidate;
        }
        return null;
    }

    /**
     * バイナリが指定の内部名の合成フォントかを判定する
     * @param {string} fontBinary - ファイルの中身
     * @param {string} internalName - 内部名（ATC-…）
     * @returns {boolean} 一致すれば true
     */
    function definesCompositeFont(fontBinary, internalName) {
        return fontBinary.substr(0, 4) === "typ1" &&
            fontBinary.indexOf("%ADOStartRearrangedFont\r/" + internalName + "\r") >= 0;
    }

    /**
     * 合成フォントのファイルを探して中身を返す
     * ファイル名は合成フォント名と同じなので先にそれを見て、外れたらフォルダー内を内部名で照合する
     * （日本語の合成フォント名は、内部名とファイル名の対応が単純でないため）
     * The file is named after the composite font, so try that first, then match internal names across the folder
     * @param {Folder} fontFolder - 合成フォントのフォルダー
     * @param {string} familyName - 合成フォント名
     * @param {string} internalName - 内部名（ATC-…）
     * @returns {string} ファイルの中身。見つからなければ空文字
     */
    function readCompositeFontBinary(fontFolder, familyName, internalName) {
        var namedFile = new File(fontFolder.fsName + "/" + familyName);
        if (namedFile.exists) {
            var namedBinary = readBinaryFile(namedFile);
            if (definesCompositeFont(namedBinary, internalName)) return namedBinary;
        }
        var fontFiles = fontFolder.getFiles(function (item) {
            return (item instanceof File) && item.name.charAt(0) !== ".";
        });
        for (var i = 0; i < fontFiles.length; i++) {
            var fontBinary = readBinaryFile(fontFiles[i]);
            if (definesCompositeFont(fontBinary, internalName)) return fontBinary;
        }
        return "";
    }

    /**
     * テーブルの位置を返す
     * @param {string} fontBinary - ファイルの中身
     * @param {string} tableTag - テーブルのタグ（"CFMA" など）
     * @returns {number} 位置。無ければ -1
     */
    function findTableOffset(fontBinary, tableTag) {
        var tableCount = readUint16(fontBinary, 4);
        for (var t = 0; t < tableCount; t++) {
            var entryOffset = 12 + 16 * t;
            if (fontBinary.substr(entryOffset, 4) === tableTag) return readUint32(fontBinary, entryOffset + 8);
        }
        return -1;
    }

    /**
     * CFMA からサイズ・水平比率・垂直比率（％）を文字セット順に読む
     * @param {string} fontBinary - ファイルの中身
     * @returns {{size: number[], hScale: number[], vScale: number[]}|null} 値。CFMA が無ければ null
     */
    function readScaleTable(fontBinary) {
        var cfmaOffset = findTableOffset(fontBinary, "CFMA");
        if (cfmaOffset < 0) return null;
        /* 見出し5語のあと、フラグが文字セット数×4語、続いてサイズ・水平・垂直が各セット数ぶん
           5-word header, then 4 flag words per set, then size / h-scale / v-scale per set */
        var setCount = readUint16(fontBinary, cfmaOffset + 8);
        var valuePos = cfmaOffset + 10 + setCount * 8;
        var scaleLists = { size: [], hScale: [], vScale: [] };
        var listKeys = ["size", "hScale", "vScale"];
        for (var k = 0; k < listKeys.length; k++) {
            for (var i = 0; i < setCount; i++) {
                scaleLists[listKeys[k]].push(readUint32(fontBinary, valuePos) / 65536);
                valuePos += 4;
            }
        }
        return scaleLists;
    }

    /**
     * bfrange の行から範囲を読む
     * @param {string} rangeText - beginbfrange と endbfrange の間の文字列
     * @returns {number[][]} [下限, 上限] の配列
     */
    function parseRangeLines(rangeText) {
        var ranges = [];
        var rangeLines = rangeText.split(/[\r\n]+/);
        for (var i = 0; i < rangeLines.length; i++) {
            /* 1行に「<下限> <上限> <割り当て先>」。割り当て先は使わない / "<low> <high> <dest>" per line; dest is unused */
            var rangeMatch = rangeLines[i].match(/^\s*<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/);
            if (rangeMatch) ranges.push([parseInt(rangeMatch[1], 16), parseInt(rangeMatch[2], 16)]);
        }
        return ranges;
    }

    /**
     * 合成フォント1文字セット分の設定
     * @typedef {Object} CharsetSpec
     * @property {string} psName - 構成フォントの PostScript名
     * @property {number} size - サイズ（％）
     * @property {number} hScale - 水平比率（％）
     * @property {number} vScale - 垂直比率（％）
     * @property {number} baseline - ベースライン（％、正の値で上）
     * @property {number[][]} ranges - 文字の範囲。漢字（0）は既定なので空
     */

    /**
     * 合成フォントのファイルから標準の文字セットの設定を読む（特例文字は読まない）
     * @param {string} fontBinary - ファイルの中身
     * @returns {CharsetSpec[]|null} 文字セット順の設定。読めなければ null
     */
    function parseCompositeFont(fontBinary) {
        var psStart = fontBinary.indexOf("%ADOStartRearrangedFont");
        if (psStart < 0) return null;
        var psText = fontBinary.substring(psStart);

        /* 構成フォントの並び（PS名＋CMap名）/ component fonts as PS name + CMap name */
        var fontListMatch = psText.match(/\[([\s\S]*?)\]\s*beginrearrangedfont/);
        if (!fontListMatch) return null;
        var cmapNames = fontListMatch[1].match(/\/\S+/g);
        if (!cmapNames || cmapNames.length < STANDARD_CHARSET_COUNT) return null;

        var scaleLists = readScaleTable(fontBinary);
        var charsetSpecs = [];
        var i;
        for (i = 0; i < STANDARD_CHARSET_COUNT; i++) {
            charsetSpecs.push({
                psName: cmapNames[i].substring(1).replace(/-UniJIS[\w-]*-[HV]$/, ""),
                size: scaleLists ? scaleLists.size[i] : 100,
                hScale: scaleLists ? scaleLists.hScale[i] : 100,
                vScale: scaleLists ? scaleLists.vScale[i] : 100,
                baseline: 0,
                ranges: []
            });
        }

        /* ベースラインは usematrix の最後の値（em に対する比率）/ the baseline is the last usematrix value, as a fraction of the em */
        var matrixPattern = /(\d+)\s+beginusematrix\s*\[([^\]]*)\]\s*endusematrix/g;
        var matrixMatch;
        while ((matrixMatch = matrixPattern.exec(psText)) !== null) {
            var matrixSet = parseInt(matrixMatch[1], 10);
            var matrixValues = matrixMatch[2].split(/\s+/);
            if (matrixSet < STANDARD_CHARSET_COUNT && matrixValues.length >= 6) {
                charsetSpecs[matrixSet].baseline = parseFloat(matrixValues[5]) * 100;
            }
        }

        var rangePattern = /(\d+)\s+usefont\s+\d+\s+beginbfrange([\s\S]*?)endbfrange/g;
        var rangeMatch;
        while ((rangeMatch = rangePattern.exec(psText)) !== null) {
            var rangeSet = parseInt(rangeMatch[1], 10);
            if (rangeSet < STANDARD_CHARSET_COUNT) {
                charsetSpecs[rangeSet].ranges = charsetSpecs[rangeSet].ranges.concat(parseRangeLines(rangeMatch[2]));
            }
        }
        return charsetSpecs;
    }

    /**
     * 文字コードが属する文字セットの番号を返す
     * 範囲は重なることがあり（半角欧文と半角数字）、後の文字セットが優先されるので後ろから見る
     * Ranges overlap (Roman and numerals) and later sets win, so search from the end
     * @param {CharsetSpec[]} charsetSpecs - 文字セット順の設定
     * @param {number} code - 文字コード
     * @returns {number} 文字セットの番号。どれにも入らなければ 0（漢字）
     */
    function findCharsetIndex(charsetSpecs, code) {
        for (var i = charsetSpecs.length - 1; i >= 1; i--) {
            var ranges = charsetSpecs[i].ranges;
            for (var k = 0; k < ranges.length; k++) {
                if (code >= ranges[k][0] && code <= ranges[k][1]) return i;
            }
        }
        return 0;
    }

    // =========================================
    // XMP / XMP
    // =========================================
    /**
     * ドキュメントの XMP から、合成フォントの構成フォント名（ファイル名）を読む
     * 文字セットとの対応は XMP に無いので、ファイルが見つからないときの案内にだけ使う
     * XMP lacks the set assignment, so this only feeds the message when the file is missing
     * @param {Document} doc - 対象のドキュメント
     * @param {string} internalName - 合成フォントの内部名（ATC-…）
     * @returns {string[]} 構成フォント名の配列。無ければ空配列
     */
    function readXmpCompositeMembers(doc, internalName) {
        var xmpString = "";
        /* XMP を持たないドキュメントでは参照が例外になることがある / may throw on a document without XMP */
        try {
            xmpString = doc.XMPString || "";
        } catch (e) {
            return [];
        }
        var fontsMatch = xmpString.match(/<xmpTPg:Fonts>[\s\S]*?<\/xmpTPg:Fonts>/);
        if (!fontsMatch) return [];

        /* 構成フォントは主フォントの rdf:li に入れ子で並ぶので、主フォントの開始タグで分ける
           Members nest inside the primary rdf:li, so split on the primary's start tag */
        var fontChunks = fontsMatch[0].split(/<rdf:li[^>]*rdf:parseType="Resource"[^>]*>/);
        for (var i = 1; i < fontChunks.length; i++) {
            if (fontChunks[i].indexOf("<stFnt:fontName>" + internalName + "</stFnt:fontName>") < 0) continue;
            var memberEntries = fontChunks[i].match(/<rdf:li[^>]*>[\s\S]*?<\/rdf:li>/g) || [];
            var members = [];
            for (var k = 0; k < memberEntries.length; k++) members.push(memberEntries[k].replace(/<[^>]+>/g, ""));
            return members;
        }
        return [];
    }

    // =========================================
    // 置き換え / Replacing
    // =========================================
    /**
     * インストール済みのフォントを PostScript名で取る
     * 環境にないフォントの仮エントリ（style が空でファミリー名が PS名のまま）は無いものとする
     * @param {string} psName - PostScript名
     * @returns {TextFont|null} フォント。無ければ null
     */
    function getInstalledFont(psName) {
        /* getByName は見つからないと例外になる / getByName throws when the font is absent */
        try {
            var textFont = app.textFonts.getByName(psName);
            if (textFont.style === "" && textFont.family === textFont.name) return null;
            return textFont;
        } catch (e) {
            return null;
        }
    }

    /**
     * 選択からテキストの範囲を集める（グループの中も含む）
     * @param {Object} docSelection - doc.selection
     * @returns {TextRange[]} テキストの範囲の配列
     */
    function collectSelectedTextRanges(docSelection) {
        var textRanges = [];
        if (!docSelection) return textRanges;
        if (docSelection.typename === "TextRange") {
            textRanges.push(docSelection);
            return textRanges;
        }
        var collectFromPageItem = function (pageItem) {
            if (pageItem.typename === "TextFrame") {
                textRanges.push(pageItem.textRange);
            } else if (pageItem.typename === "GroupItem") {
                for (var k = 0; k < pageItem.pageItems.length; k++) collectFromPageItem(pageItem.pageItems[k]);
            }
        };
        for (var i = 0; i < docSelection.length; i++) collectFromPageItem(docSelection[i]);
        return textRanges;
    }

    /**
     * 1文字を構成フォントに置き換える
     * @param {TextRange} textCharacter - 対象の文字
     * @param {TextFont} componentFont - 構成フォント
     * @param {CharsetSpec} charsetSpec - 文字セットの設定
     * @param {boolean} pinLeading - 自動行送りを固定値にするなら true
     * @returns {void}
     */
    function replaceCharacterFont(textCharacter, componentFont, charsetSpec, pinLeading) {
        var charAttributes = textCharacter.characterAttributes;
        var baseSize = charAttributes.size;
        var sizeRatio = charsetSpec.size / 100;

        /* 行送りはサイズを変える前の値で固定する / pin the leading before the size changes */
        if (pinLeading && charAttributes.autoLeading) {
            var autoLeadingAmount = textCharacter.paragraphAttributes.autoLeadingAmount;
            charAttributes.autoLeading = false;
            charAttributes.leading = baseSize * autoLeadingAmount / 100;
        }

        charAttributes.textFont = componentFont;
        if (sizeRatio !== 1) {
            charAttributes.size = baseSize * sizeRatio;
            if (KEEP_SPACING) {
                if (charAttributes.tracking !== 0) charAttributes.tracking = Math.round(charAttributes.tracking / sizeRatio);
                /* 手動カーニングが無い位置（自動・オプティカル）では読むと例外になる / reading throws where no manual kerning is set */
                try {
                    var manualKerning = textCharacter.kerning;
                    if (manualKerning !== 0) textCharacter.kerning = Math.round(manualKerning / sizeRatio);
                } catch (e) {}
            }
        }
        if (charsetSpec.hScale !== 100) charAttributes.horizontalScale = charAttributes.horizontalScale * charsetSpec.hScale / 100;
        if (charsetSpec.vScale !== 100) charAttributes.verticalScale = charAttributes.verticalScale * charsetSpec.vScale / 100;
        /* ベースラインは元のサイズに対する比率 / the baseline is relative to the original size */
        if (charsetSpec.baseline !== 0) charAttributes.baselineShift = charAttributes.baselineShift + baseSize * charsetSpec.baseline / 100;
    }

    /**
     * 合成フォント1つ分の置き換え情報
     * @typedef {Object} CompositeEntry
     * @property {string} family - 合成フォント名
     * @property {CharsetSpec[]|null} charsetSpecs - 文字セット順の設定。ファイルが無ければ null
     * @property {Array} componentFonts - 文字セット順の構成フォント（TextFont）。環境に無いものは null
     * @property {boolean} pinLeading - 自動行送りを固定値にするなら true
     */

    /**
     * 合成フォント1つ分の置き換え情報を用意する（合成フォントごとに1回だけ読む）
     * @param {TextFont} compositeFont - 合成フォント
     * @param {Folder|null} fontFolder - 合成フォントのフォルダー。無ければ null
     * @returns {CompositeEntry} 置き換え情報
     */
    function prepareCompositeEntry(compositeFont, fontFolder) {
        var fontBinary = fontFolder ? readCompositeFontBinary(fontFolder, compositeFont.family, compositeFont.name) : "";
        var charsetSpecs = fontBinary ? parseCompositeFont(fontBinary) : null;
        var componentFonts = [];
        var hasResizedSet = false;
        if (charsetSpecs) {
            for (var i = 0; i < charsetSpecs.length; i++) {
                componentFonts.push(getInstalledFont(charsetSpecs[i].psName));
                if (charsetSpecs[i].size !== 100) hasResizedSet = true;
            }
        }
        return {
            family: compositeFont.family,
            charsetSpecs: charsetSpecs,
            componentFonts: componentFonts,
            /* サイズの違う文字セットがあるときだけ行送りを固定する / pin leading only when some set is resized */
            pinLeading: KEEP_LEADING && hasResizedSet
        };
    }

    /**
     * 置き換えの集計
     * @typedef {Object} ReplaceTally
     * @property {Object} compositeEntries - 内部名ごとの置き換え情報（CompositeEntry）
     * @property {Object} notFoundCounts - ファイルが無かった合成フォントの内部名ごとの文字数
     * @property {Object} missingFontCounts - 環境に無い構成フォントの PS名ごとの文字数
     * @property {number} replacedCount - 置き換えた文字数
     * @property {number} compositeCount - 合成フォントの文字数
     */

    /**
     * キーごとの件数を1つ増やす
     * @param {Object} countsByKey - キーごとの件数
     * @param {string} countKey - キー
     * @returns {void}
     */
    function incrementCount(countsByKey, countKey) {
        countsByKey[countKey] = (countsByKey[countKey] || 0) + 1;
    }

    /**
     * テキストの範囲にある合成フォントの文字を、構成フォントに置き換える
     * @param {TextRange} textRange - 対象のテキスト
     * @param {Folder|null} fontFolder - 合成フォントのフォルダー
     * @param {ReplaceTally} replaceTally - 集計（書き換える）
     * @returns {void}
     */
    function replaceCompositeFontsInRange(textRange, fontFolder, replaceTally) {
        /* 文字の番号は contents と1対1なので、文字コードは contents から読む / character indexes match contents one to one */
        var rangeContents = textRange.contents;
        var textCharacters = textRange.characters;
        var charCount = textCharacters.length;
        for (var j = 0; j < charCount; j++) {
            var textCharacter = textCharacters[j];
            var currentFont;
            /* 壊れたストーリーやプラグイン所有のテキストでは参照自体が失敗するので、その文字だけ飛ばす
               A broken story or plugin-owned text throws on access, so skip just that character */
            try {
                currentFont = textCharacter.characterAttributes.textFont;
            } catch (e) {
                continue;
            }
            if (!currentFont || currentFont.name.indexOf("ATC-") !== 0) continue;
            replaceTally.compositeCount++;

            var internalName = currentFont.name;
            if (!replaceTally.compositeEntries[internalName]) {
                replaceTally.compositeEntries[internalName] = prepareCompositeEntry(currentFont, fontFolder);
            }
            var compositeEntry = replaceTally.compositeEntries[internalName];
            if (!compositeEntry.charsetSpecs) {
                incrementCount(replaceTally.notFoundCounts, internalName);
                continue;
            }

            var setIndex = findCharsetIndex(compositeEntry.charsetSpecs, rangeContents.charCodeAt(j));
            var charsetSpec = compositeEntry.charsetSpecs[setIndex];
            var componentFont = compositeEntry.componentFonts[setIndex];
            if (!componentFont) {
                incrementCount(replaceTally.missingFontCounts, charsetSpec.psName);
                continue;
            }
            replaceCharacterFont(textCharacter, componentFont, charsetSpec, compositeEntry.pinLeading);
            replaceTally.replacedCount++;
        }
    }

    /**
     * 結果の文面を作る
     * @param {Document} doc - 対象のドキュメント
     * @param {ReplaceTally} replaceTally - 集計
     * @returns {string} アラートの文面
     */
    function buildResultMessage(doc, replaceTally) {
        var messageLines = [getLabel("alert.done").replace("%1", replaceTally.replacedCount)];
        var countKey;
        for (countKey in replaceTally.notFoundCounts) {
            messageLines.push("");
            messageLines.push(getLabel("alert.fileNotFound")
                .replace("%1", replaceTally.compositeEntries[countKey].family).replace("%2", replaceTally.notFoundCounts[countKey]));
            var xmpMembers = readXmpCompositeMembers(doc, countKey);
            if (xmpMembers.length > 0) messageLines.push(getLabel("alert.xmpMembers") + "\n" + xmpMembers.join("\n"));
        }
        for (countKey in replaceTally.missingFontCounts) {
            messageLines.push("");
            messageLines.push(getLabel("alert.missingFont").replace("%1", countKey).replace("%2", replaceTally.missingFontCounts[countKey]));
        }
        return messageLines.join("\n");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================
    /**
     * 選択したテキストの合成フォントを構成フォントに置き換える
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var textRanges = collectSelectedTextRanges(doc.selection);
        if (textRanges.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }
        var fontFolder = findCompositeFontFolder();
        var replaceTally = { compositeEntries: {}, notFoundCounts: {}, missingFontCounts: {}, replacedCount: 0, compositeCount: 0 };
        for (var i = 0; i < textRanges.length; i++) replaceCompositeFontsInRange(textRanges[i], fontFolder, replaceTally);

        if (replaceTally.compositeCount === 0) {
            alert(getLabel("alert.noComposite"));
        } else if (!fontFolder) {
            alert(getLabel("alert.noFolder"));
        } else {
            alert(buildResultMessage(doc, replaceTally));
        }
    }

    main();

})();
