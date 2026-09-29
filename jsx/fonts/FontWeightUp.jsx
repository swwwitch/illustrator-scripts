#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択しているテキストのフォントを、同じファミリー・同じ系列（イタリックや字幅）の一つ上のウェイトに切り替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontWeightUp.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n255437cfdba0

### 注意

合成フォントと、ファミリー内で最も太いウェイトは変更しません。ウェイトの判定は TypefaceSampler.jsx と同じです。

### Overview

Switches the selected text to the next heavier weight in the same family and the same series (italic, width).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontWeightUp.md

### Notes

Composite fonts and the heaviest weight of a family are left unchanged. Weights are judged the same way as TypefaceSampler.jsx.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontWeightUp";                 /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontWeightUp.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontWeightUp.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n255437cfdba0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 別々のファミリーとして並ぶフォントを一つのファミリーとして扱う。fonts は細い順に並べ、
       ファミリー名・PostScript 名のどちらかと一致すれば該当する（大文字小文字は区別しない）
       Fonts listed as separate families but treated as one; list fonts from thin to heavy.
       A font matches by family name or PostScript name, case-insensitively */
    var FAMILY_GROUPS = [
        { family: "sw", fonts: ["sw-L", "sw-R", "sw-B", "sw-H"] }
    ];

    // =========================================
    // ウェイト語句の定義 / Weight term definitions
    // （TypefaceSampler.jsx と同じ定義 / Shared with TypefaceSampler.jsx）
    // =========================================

    /* ウェイト語句の並び（インデックスが大きいほど太い）/ Weight terms ordered from thin to bold */
    var WEIGHT_GROUPS = [
        ["hairline", "hair"], // +0
        ["ultra thin", "ultrathin", "ut"], // +1
        ["thin", "th"], // +2
        ["default"], // +3
        ["ultralight", "ultra light", "ultlt", "ul"], // +4
        ["extralight", "extra light", "el", "xlight", "xl"], // +5
        ["lightsemi"], // +6
        ["light", "lt", "lite", "l"], // +7
        ["lb"], // +8
        ["book", "bk"], // +9
        ["n", "normal"], // +10
        ["middle"], // +11
        ["regular", "roman", "レギュラー", "r"], // +12
        ["rb"], // +13
        ["medium", "md", "ミディアム", "m"], // +14
        ["semibold", "semi bold", "sb"], // +15
        ["demibold", "demi bold", "db", "デミボールド", "demi", "d", "demixtra"], // +16
        ["bold", "bd", "ボールド", "b"], // +17
        ["extrabold", "extra bold", "xbold", "エクストラボールド", "e", "eb", "xb"], // +18
        ["heavy", "h"], // +19
        ["black"], // +20
        ["xblack", "extra black", "extrablack"], // +21
        ["ultra", "u", "ub", "ultra black", "ultrablack"] // +22
    ];

    /* 単独で使われたら Regular 扱いする装飾語句 / Decoration-only styles treated as Regular */
    var DECORATION_ONLY_STYLES = [
        "display", "compressed", "comp", "compact", "expanded", "extended", "semiextended",
        "ultracondensed", "extracondensed", "semicondensed", "cond", "condensed", "wide",
        "headline", "text", "low", "micro", "extra compressed",
        "semi expanded", "semiexpanded"
    ];

    /* 幅を表す複合語（Ultra Condensed など）。ultra / extra をウェイト語と取り違えないよう照合前に除く
       Width compounds such as "ultra condensed"; removed first so "ultra" is not read as a weight */
    var WIDTH_COMPOUND_PATTERN = /(^|\s)(ultra|extra|semi)\s+(condensed|cond|compressed|comp|expanded|extended)(?=\s|$)/g;

    /* WEIGHT_GROUPS における Regular のインデックス / Index of "regular" in WEIGHT_GROUPS */
    var REGULAR_GROUP_INDEX = (function() {
        for (var i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (var j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                if (WEIGHT_GROUPS[i][j] === "regular") return i;
            }
        }
        return 12; /* fallback */
    })();

    /* 複合語一致用に、長い語から順に並べた照合テーブル / Match table sorted by term length */
    var WEIGHT_TERM_PATTERNS = (function() {
        var termPatterns = [];
        for (var i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (var j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                var weightTerm = WEIGHT_GROUPS[i][j];
                /* \b は和文の前後で効かないため、英数字以外を境界とみなす / \b fails next to Japanese, so treat any non-alphanumeric as a boundary */
                termPatterns.push({
                    term: weightTerm,
                    groupIndex: i,
                    pattern: new RegExp("(?:^|[^a-z0-9])" + weightTerm.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&') + "(?=[^a-z0-9]|$)")
                });
            }
        }
        /* 同じ長さなら細い方を先に（並べ替えの結果を環境で変えない）/ Break ties by group so the order is deterministic */
        termPatterns.sort(function(a, b) {
            return (b.term.length - a.term.length) || (a.groupIndex - b.groupIndex);
        });
        return termPatterns;
    })();

    /**
     * スタイル文字列を照合用に正規化する
     * @param {string} rawStyle - font.style の値
     * @returns {string} 小文字化し、区切り記号を空白に置き換えた文字列
     */
    function normalizeStyle(rawStyle) {
        return (rawStyle || "").toLowerCase().replace(/[_\-]+/g, " ").replace(/^\s+|\s+$/g, "");
    }

    /**
     * スタイル文字列に一致する WEIGHT_GROUPS のインデックスを返す
     * 幅の複合語を除いてから、完全一致を優先し、なければ長い語から順に語の境界つきで照合する
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @returns {number} 一致したインデックス。見つからない場合は -1
     */
    function getWeightGroupIndex(normalizedStyle) {
        var weightStyle = normalizedStyle.replace(WIDTH_COMPOUND_PATTERN, " ").replace(/\s+/g, " ").replace(/^\s+|\s+$/g, "");
        if (weightStyle === "") return -1;

        var i, j;
        for (i = 0; i < WEIGHT_GROUPS.length; i++) {
            for (j = 0; j < WEIGHT_GROUPS[i].length; j++) {
                if (weightStyle === WEIGHT_GROUPS[i][j]) return i;
            }
        }

        for (i = 0; i < WEIGHT_TERM_PATTERNS.length; i++) {
            if (WEIGHT_TERM_PATTERNS[i].pattern.test(weightStyle)) return WEIGHT_TERM_PATTERNS[i].groupIndex;
        }

        return -1;
    }

    /**
     * W3・W600・25 Ultra Light のような数値スタイルを読み取る
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @returns {object|null} value（数値）と digitCount（桁数）。数値スタイルでなければ null
     */
    function getNumericWeight(normalizedStyle) {
        var numericMatch = normalizedStyle.match(/^w?(\d{1,3})(?=\D|$)/);
        if (!numericMatch) return null;
        return { value: parseInt(numericMatch[1], 10), digitCount: numericMatch[1].length };
    }

    // =========================================
    // ウェイト評価 / Weight scoring
    // （TypefaceSampler.jsx の並べ替え評価を、ウェイトと装飾語の加点に分けたもの
    //   TypefaceSampler.jsx's sort score, split into weight and decoration offset）
    // =========================================

    /**
     * スタイル文字列に対する基本ウェイトスコアを取得する
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @param {string} postscriptName - 小文字化した PostScript 名
     * @param {string} familyName - 小文字化したファミリー名
     * @returns {number} ウェイトの評価値（小さいほど細い）
     */
    function getBaseWeightScore(normalizedStyle, postscriptName, familyName) {
        var styleWords = normalizedStyle.split(/\s+/);
        var i;

        var applyFrutigerCorrection = (/frutiger/i.test(familyName) && /ultralight/.test(normalizedStyle));

        /* W0〜W9、W000〜W999、先頭数値（例：25 Ultra Light）/ Numeric styles */
        var numericWeight = getNumericWeight(normalizedStyle);
        if (numericWeight) return numericWeight.value;

        /* 特例：HelveticaNeue, Tazugane, UniversNextPro + Ultra Light → 999 */
        if (
            (
                /helveticaneue/i.test(postscriptName) ||
                /tazugane/i.test(postscriptName) ||
                /universnextpro/i.test(postscriptName)
            ) &&
            /ultralight|ultra light|ultlt/i.test(normalizedStyle)
        ) {
            return 999;
        }

        /* 単独語が italic / oblique / wide → Regular 扱い */
        if (styleWords.length === 1 && /^(italic|oblique|it|wide)$/.test(styleWords[0])) {
            return 1000 + REGULAR_GROUP_INDEX;
        }

        /* 装飾語だけなら Regular 扱い */
        if (styleWords.length === 1) {
            for (i = 0; i < DECORATION_ONLY_STYLES.length; i++) {
                if (styleWords[0] === DECORATION_ONLY_STYLES[i]) return 1000 + REGULAR_GROUP_INDEX;
            }
        }

        /* 完全一致・複合語一致（長い語優先）/ Exact match, then longest-term match */
        var groupIndex = getWeightGroupIndex(normalizedStyle);
        if (groupIndex !== -1) {
            var weightScore = 1000 + groupIndex;
            if (applyFrutigerCorrection && groupIndex === 4) weightScore -= 5;
            return weightScore;
        }

        /* fallbackScore：Regular 扱い */
        var fallbackScore = 1000 + REGULAR_GROUP_INDEX;
        if (applyFrutigerCorrection) fallbackScore -= 5;
        return fallbackScore;
    }

    /**
     * 字幅・用途・イタリックなど、ウェイト以外の装飾語に対する加点を取得する
     * 同じ加点のフォント同士を「同じ系列」とみなす
     * @param {string} normalizedStyle - 正規化済みのスタイル文字列
     * @returns {number} 装飾語の加点（装飾なしは 0）
     */
    function getDecorationOffset(normalizedStyle) {
        var decorationOffset = 0;
        var styleWords = normalizedStyle.split(/\s+/);

        /* 装飾フラグ初期化 / Initialize decoration flags */
        var decorationFlags = {
            hasText: false,
            hasHeadline: false,
            hasCondensed: false,
            hasCn: false,
            hasExpanded: false,
            hasExtended: false,
            hasUltraCondensed: false,
            hasExtraCondensed: false,
            hasSemiCondensed: false,
            hasCompressed: false,
            hasExtraCompressed: false,
            hasCompact: false,
            hasDisplay: false,
            hasMicro: false,
            hasLow: false,
            hasWide: false
        };

        /* 装飾キーワードに応じたフラグ設定 / Set a flag for each decoration keyword */
        for (var i = 0; i < styleWords.length; i++) {
            var styleWord = styleWords[i];
            if (styleWord === "text") decorationFlags.hasText = true;
            if (styleWord === "headline") decorationFlags.hasHeadline = true;
            if (styleWord === "cond" || styleWord === "condensed") decorationFlags.hasCondensed = true;
            if (styleWord === "cn") decorationFlags.hasCn = true;
            if (styleWord === "expanded") decorationFlags.hasExpanded = true;
            if (styleWord === "extended") decorationFlags.hasExtended = true;
            if (styleWord === "semiextended" || (styleWord === "semi" && styleWords[i + 1] === "extended")) decorationFlags.hasExtended = true;
            if (styleWord === "semiexpanded" || (styleWord === "semi" && styleWords[i + 1] === "expanded")) decorationFlags.hasExpanded = true;
            if (styleWord === "ultracondensed" || (styleWord === "ultra" && styleWords[i + 1] === "condensed")) decorationFlags.hasUltraCondensed = true;
            if (styleWord === "extracondensed" || (styleWord === "extra" && styleWords[i + 1] === "condensed")) decorationFlags.hasExtraCondensed = true;
            if (styleWord === "semicondensed" || (styleWord === "semi" && styleWords[i + 1] === "condensed")) decorationFlags.hasSemiCondensed = true;
            if (styleWord === "compressed" || styleWord === "comp") decorationFlags.hasCompressed = true;
            if (styleWord === "extra" && styleWords[i + 1] === "compressed") decorationFlags.hasExtraCompressed = true;
            if (styleWord === "compact") decorationFlags.hasCompact = true;
            if (styleWord === "display") decorationFlags.hasDisplay = true;
            if (styleWord === "micro") decorationFlags.hasMicro = true;
            if (styleWord === "low") decorationFlags.hasLow = true;
            if (styleWord === "wide") decorationFlags.hasWide = true;
        }

        /* Italic 判定（全体 styleName に対して）/ Detect italic across the whole styleName */
        var isItalic = /italic|oblique|slanted|inclined|kursiv|\bit\b/.test(normalizedStyle);

        /* 加点処理（100刻み + 特例あり）/ Offsets in steps of 100, with exceptions */
        if (decorationFlags.hasDisplay) decorationOffset += 100;
        if (decorationFlags.hasCompressed) decorationOffset += 200;
        if (decorationFlags.hasCompact) decorationOffset += 300;
        if (decorationFlags.hasExpanded) decorationOffset += 400;
        if (decorationFlags.hasExtended) decorationOffset += 500;
        if (decorationFlags.hasUltraCondensed) decorationOffset += 600;
        if (decorationFlags.hasExtraCondensed) decorationOffset += 700;
        if (decorationFlags.hasSemiCondensed) decorationOffset += 850;

        /* Condensed系代表加点（複数条件一致でも一度のみ）/ Applied once even on multiple matches */
        if (
            decorationFlags.hasCondensed ||
            decorationFlags.hasCn ||
            decorationFlags.hasWide ||
            decorationFlags.hasSemiCondensed ||
            decorationFlags.hasExtraCompressed
        ) {
            decorationOffset += 900;
        }

        if (decorationFlags.hasHeadline) decorationOffset += 1000;
        if (decorationFlags.hasText) decorationOffset += 1100;
        if (decorationFlags.hasLow) decorationOffset += 1200;
        if (decorationFlags.hasMicro) decorationOffset += 1250;
        if (decorationFlags.hasWide) decorationOffset += 1275;
        if (decorationFlags.hasExtraCompressed) decorationOffset += 150; /* 特別加点 / Extra offset */
        if (isItalic) decorationOffset += 1300;

        return decorationOffset;
    }

    /**
     * FAMILY_GROUPS に登録されたフォントなら、ファミリーと並び順を返す
     * @param {TextFont} textFont - 対象フォント
     * @returns {{family: string, order: number}|null} 登録されたファミリー名と細い順の番号。未登録は null
     */
    function findFamilyGroupMember(textFont) {
        var familyName = (textFont.family || "").toLowerCase();
        var postscriptName = (textFont.name || "").toLowerCase();
        for (var i = 0; i < FAMILY_GROUPS.length; i++) {
            var groupFonts = FAMILY_GROUPS[i].fonts;
            for (var j = 0; j < groupFonts.length; j++) {
                var memberName = groupFonts[j].toLowerCase();
                if (memberName === familyName || memberName === postscriptName) {
                    return { family: FAMILY_GROUPS[i].family, order: j };
                }
            }
        }
        return null;
    }

    /**
     * ウェイトをそろえる単位となるファミリー名を返す（FAMILY_GROUPS に登録されたフォントはまとめた名前）
     * @param {TextFont} textFont - 対象フォント
     * @returns {string} ファミリー名
     */
    function getFamilyKey(textFont) {
        var groupMember = findFamilyGroupMember(textFont);
        return groupMember ? groupMember.family : textFont.family;
    }

    /**
     * フォントのウェイトと系列（装飾語の加点）を返す
     * @param {TextFont} textFont - 対象フォント
     * @returns {{weight: number, variant: number}|null} 判定結果。style が空（合成フォント・置換用の仮エントリ）は null
     */
    function getWeightInfo(textFont) {
        /* FAMILY_GROUPS に登録されたフォントは並び順をウェイトとする / Registered fonts use their listed order */
        var groupMember = findFamilyGroupMember(textFont);
        if (groupMember) return { weight: groupMember.order, variant: 0 };

        if (!textFont.style) return null;

        var styleName = normalizeStyle(textFont.style);
        var postscriptName = textFont.name.toLowerCase();

        /* 特例：PostScript名が「FuturaPT-Heavy」なら 1015 固定（加点処理なし）/ Fixed rank, no offsets */
        if (postscriptName === "futurapt-heavy") return { weight: 1015, variant: 0 };

        return {
            weight: getBaseWeightScore(styleName, postscriptName, textFont.family.toLowerCase()),
            variant: getDecorationOffset(styleName)
        };
    }

    // =========================================
    // 対象テキストの収集 / Collect target text
    // =========================================

    /**
     * 選択から対象のテキスト範囲を集める（文字選択中はその範囲、オブジェクト選択ではグループ内も含む）
     * @param {Object} docSelection - doc.selection
     * @returns {TextRange[]} テキスト範囲の配列
     */
    function collectTextRanges(docSelection) {
        var textRanges = [];
        if (!docSelection) return textRanges;

        /* 文字選択中は TextRange が単独で返る / A TextRange is returned while editing text */
        if (docSelection.typename === "TextRange") {
            textRanges.push(docSelection);
            return textRanges;
        }

        for (var i = 0; i < docSelection.length; i++) {
            collectTextRangesFromItem(docSelection[i], textRanges);
        }
        return textRanges;
    }

    /**
     * ページアイテムからテキスト範囲を集める（グループは再帰）
     * @param {PageItem} pageItem - 対象アイテム
     * @param {TextRange[]} textRanges - 追加先
     * @returns {void}
     */
    function collectTextRangesFromItem(pageItem, textRanges) {
        if (pageItem.typename === "TextFrame") {
            textRanges.push(pageItem.textRange);
        } else if (pageItem.typename === "GroupItem") {
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                collectTextRangesFromItem(pageItem.pageItems[i], textRanges);
            }
        }
    }

    /**
     * テキスト範囲の文字ごとに現在のフォントを控え、使われているファミリーを集める
     * @param {TextRange[]} textRanges - 対象のテキスト範囲
     * @returns {{characterEntries: Object[], familyNames: Object}} 文字と元フォントの組、ファミリー名の集合
     */
    function collectCharacterFonts(textRanges) {
        var characterEntries = [];
        var familyNames = {};
        for (var i = 0; i < textRanges.length; i++) {
            var characters = textRanges[i].characters;
            for (var j = 0; j < characters.length; j++) {
                var textFont = characters[j].characterAttributes.textFont;
                characterEntries.push({ character: characters[j], font: textFont });
                familyNames[getFamilyKey(textFont)] = true;
            }
        }
        return { characterEntries: characterEntries, familyNames: familyNames };
    }

    // =========================================
    // ファミリーの索引 / Family index
    // =========================================

    /**
     * 指定ファミリーに属するフォントを app.textFonts から一度の走査で集める
     * @param {Object} familyNames - ファミリー名をキーにした集合
     * @returns {Object} ファミリー名 → [{font: TextFont, weight: number, variant: number}]
     */
    function buildFamilyIndex(familyNames) {
        var familyIndex = {};
        var textFonts = app.textFonts;
        for (var i = 0; i < textFonts.length; i++) {
            var textFont = textFonts[i];
            var family = getFamilyKey(textFont);
            if (!familyNames.hasOwnProperty(family)) continue;

            var weightInfo = getWeightInfo(textFont);
            if (!weightInfo) continue;

            if (!familyIndex.hasOwnProperty(family)) familyIndex[family] = [];
            familyIndex[family].push({ font: textFont, weight: weightInfo.weight, variant: weightInfo.variant });
        }
        return familyIndex;
    }

    /**
     * 同じファミリー・同じ系列で、一つ上のウェイトのフォントを返す
     * @param {TextFont} currentFont - 現在のフォント
     * @param {Object} familyIndex - buildFamilyIndex() の結果
     * @returns {TextFont|null} 一つ上のフォント。無ければ null
     */
    function findNextHeavierFont(currentFont, familyIndex) {
        var currentWeightInfo = getWeightInfo(currentFont);
        var familyMembers = familyIndex[getFamilyKey(currentFont)];
        if (!currentWeightInfo || !familyMembers) return null;

        var nextMember = null;
        for (var i = 0; i < familyMembers.length; i++) {
            var familyMember = familyMembers[i];
            if (familyMember.variant !== currentWeightInfo.variant) continue;
            if (familyMember.weight <= currentWeightInfo.weight) continue;
            /* 同じウェイトは PostScript 名順で先のもの（TypefaceSampler.jsx の並びに合わせる）/ Ties: PostScript name order, as in TypefaceSampler.jsx */
            if (!nextMember || familyMember.weight < nextMember.weight ||
                (familyMember.weight === nextMember.weight && familyMember.font.name < nextMember.font.name)) {
                nextMember = familyMember;
            }
        }
        return nextMember ? nextMember.font : null;
    }

    // =========================================
    // ウェイトの適用 / Apply weights
    // =========================================

    /**
     * 各文字に一つ上のウェイトを適用する（行き先はフォント名ごとに控えて使い回す）
     * @param {Object[]} characterEntries - collectCharacterFonts() の文字と元フォントの組
     * @param {Object} familyIndex - buildFamilyIndex() の結果
     * @returns {string[]} 一つ上が見つからず変更しなかったフォントの表示名
     */
    function applyNextHeavierFonts(characterEntries, familyIndex) {
        var nextFontByName = {};
        var unchangedFontNames = [];
        for (var i = 0; i < characterEntries.length; i++) {
            var sourceFont = characterEntries[i].font;
            var fontName = sourceFont.name;
            if (!nextFontByName.hasOwnProperty(fontName)) {
                nextFontByName[fontName] = findNextHeavierFont(sourceFont, familyIndex);
                if (!nextFontByName[fontName]) {
                    unchangedFontNames.push(sourceFont.style ? (sourceFont.family + " " + sourceFont.style) : sourceFont.family);
                }
            }
            if (nextFontByName[fontName]) {
                characterEntries[i].character.characterAttributes.textFont = nextFontByName[fontName];
            }
        }
        return unchangedFontNames;
    }

    // =========================================
    // ローカライズ / Localization
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
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noText:     { ja: "テキストが選択されていません。", en: "No text is selected." },
            unchangedFonts: {
                ja: "次のフォントは一つ上のウェイトが見つからないため、変更しませんでした。",
                en: "The following fonts were left unchanged because no heavier weight was found."
            }
        }
    };

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var characterFonts = collectCharacterFonts(collectTextRanges(app.activeDocument.selection));
    if (characterFonts.characterEntries.length === 0) {
        alert(getLabel(LABELS.alert.noText));
        return;
    }

    var familyIndex = buildFamilyIndex(characterFonts.familyNames);
    var unchangedFontNames = applyNextHeavierFonts(characterFonts.characterEntries, familyIndex);
    if (unchangedFontNames.length > 0) {
        alert(getLabel(LABELS.alert.unchangedFonts) + "\n\n" + unchangedFontNames.join("\n"));
    }

})();
