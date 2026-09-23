#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択テキスト内で文字サイズが混在しているとき、各テキストの先頭文字のサイズへ統一します。
見た目の大きさは変えず、サイズ差を水平比率・垂直比率に変換して補正します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontSizeToScaleConverter.md

### Overview

Unifies mixed font sizes in the selected text to the size of each text's first character.
The apparent size is preserved by converting the difference into horizontal and vertical scale.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontSizeToScaleConverter.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontSizeToScaleConverter";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontSizeToScaleConverter.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontSizeToScaleConverter.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var FONT_SIZE_TOLERANCE = 0.001;  /* 同一サイズとみなす許容誤差（pt） / tolerance for treating sizes as equal (pt) */
    var MIN_CHARACTER_SCALE = 1;      /* 文字比率の下限（%） / minimum character scale (%) */
    var MAX_CHARACTER_SCALE = 10000;  /* 文字比率の上限（%） / maximum character scale (%) */
    var PRESERVE_LEADING    = true;   /* 変換前の行送りを実値で固定する / pin the current leading as a fixed value */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールからUIの表示言語を判定する
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "テキストを選択してください。", en: "Please select text." },
            noText: { ja: "処理できるテキストが選択されていません。", en: "No editable text is selected." }
        }
    };

    /**
     * キーからラベルを現在の言語で取得する（"alert.noDocument" のようにドット区切り）
     * @param {string} labelPath - カテゴリ名とキー名をドットでつないだラベルキー
     * @returns {string} 現在の言語のラベル文字列（未定義の場合は英語にフォールバック）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelEntry = LABELS[pathKeys[0]][pathKeys[1]];
        return labelEntry[uiLang] || labelEntry.en;
    }

    // =========================================
    // 文字属性ユーティリティ / Character attribute helpers
    // =========================================

    /**
     * 文字比率が Illustrator で設定できる範囲に収まっているか判定する
     * @param {number} scale - 判定する文字比率（%）
     * @returns {boolean} 範囲内なら true
     */
    function isScaleInRange(scale) {
        return scale >= MIN_CHARACTER_SCALE && scale <= MAX_CHARACTER_SCALE;
    }

    /**
     * 自動行送りの文字について、現在の行送りを実値へ固定する
     * @param {CharacterAttributes} charAttributes - 対象文字の文字属性
     * @returns {void}
     */
    function pinCurrentLeading(charAttributes) {
        try {
            /* 自動行送りでなければ、サイズを変えても行送りは動かない
               A fixed leading is unaffected by the size change */
            if (charAttributes.autoLeading !== true) return;

            var currentLeading = charAttributes.leading;
            if (!currentLeading || currentLeading <= 0) return;

            charAttributes.autoLeading = false;
            charAttributes.leading = currentLeading;
        } catch (e) {
            /* 行送りを扱えない環境では固定をあきらめる / Give up pinning when leading is unavailable */
        }
    }

    /**
     * テキスト内の各文字サイズを先頭文字のサイズへ揃え、差分を文字比率へ変換する
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @returns {boolean} 処理対象として扱えたら true、空テキストや基準サイズが不正なら false
     */
    function convertTextRangeSizes(textRange) {
        /* コレクションと文字数をキャッシュ（毎回参照すると長文で著しく遅くなる）
           Cache the collection and count (re-reading them per iteration is very slow for long text) */
        var characters = textRange.characters;
        var characterCount = characters.length;

        /* 空テキストは処理対象に数えない / Do not count empty text as processed */
        if (characterCount === 0) return false;

        /* 先頭文字のサイズを基準にする / Use the first character's size as the base */
        var baseFontSize = characters[0].characterAttributes.size;

        /* 基準サイズが 0 や不正値なら除算できないのでスキップ / Skip when the base size is zero or invalid (cannot divide) */
        if (!baseFontSize || baseFontSize <= 0) return false;

        for (var i = 0; i < characterCount; i++) {
            var charAttributes = characters[i].characterAttributes;
            var originalFontSize = charAttributes.size;

            /* 許容誤差内は同一サイズとみなす（浮動小数の誤差で無意味な比率が入るのを防ぐ）
               Treat sizes within the tolerance as equal (avoids bogus scales caused by float error) */
            if (Math.abs(originalFontSize - baseFontSize) <= FONT_SIZE_TOLERANCE) continue;

            /* 見た目の大きさを保つための倍率 / Ratio that preserves the visual size */
            var sizeRatio = originalFontSize / baseFontSize;
            var newHorizontalScale = charAttributes.horizontalScale * sizeRatio;
            var newVerticalScale = charAttributes.verticalScale * sizeRatio;

            /* 設定できる範囲を超える文字は見た目を保てないので変換しない
               Skip characters whose scale would fall outside the settable range */
            if (!isScaleInRange(newHorizontalScale) || !isScaleInRange(newVerticalScale)) continue;

            /* 自動行送りはサイズから算出されるため、サイズ変更前に行送りを固定する
               Auto leading is derived from the size, so pin it before changing the size */
            if (PRESERVE_LEADING) pinCurrentLeading(charAttributes);

            charAttributes.size = baseFontSize;
            charAttributes.horizontalScale = newHorizontalScale;
            charAttributes.verticalScale = newVerticalScale;
        }
        return true;
    }

    /**
     * オブジェクト種別ごとに処理を振り分ける（テキストは変換、グループは再帰）
     * @param {PageItem|TextRange} pageItem - 選択オブジェクト、またはグループ内のオブジェクト
     * @returns {number} 処理対象として扱えたテキストの数
     */
    function convertPageItem(pageItem) {
        /* テキスト編集中の選択範囲。ロック・非表示の概念が無いので種別判定を先に行う
           A range selected while editing text has no locked/hidden state, so check the type first */
        if (pageItem.typename === "TextRange") {
            return convertTextRangeSizes(pageItem) ? 1 : 0;
        }

        /* ロック中・非表示のオブジェクトはスキップ / Skip locked or hidden objects */
        if (pageItem.locked || pageItem.hidden) return 0;

        if (pageItem.typename === "TextFrame") {
            /* ポイント文字・エリア内文字・パス上文字はすべて TextFrame として扱える
               Point, area, and path text are all TextFrame objects */
            return convertTextRangeSizes(pageItem.textRange) ? 1 : 0;
        }

        var convertedCount = 0;
        if (pageItem.typename === "GroupItem") {
            /* グループ内のテキストを再帰的にたどる / Recurse into group contents */
            for (var i = 0; i < pageItem.pageItems.length; i++) {
                convertedCount += convertPageItem(pageItem.pageItems[i]);
            }
        }
        return convertedCount;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 前提チェックののち、選択したテキストの文字サイズを文字比率へ変換する
     * @returns {void}
     */
    function main() {
        /* ドキュメントの有無を確認 / Check that a document is open */
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        /* 選択の有無を確認 / Check that something is selected */
        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 処理中に選択が変化しても影響を受けないよう控えを取る / Snapshot the selection before modifying anything */
        var selectedItems = [];
        for (var i = 0; i < doc.selection.length; i++) {
            selectedItems.push(doc.selection[i]);
        }

        /* 選択した各オブジェクトを処理 / Process each selected object */
        var targetTextCount = 0;
        for (var j = 0; j < selectedItems.length; j++) {
            targetTextCount += convertPageItem(selectedItems[j]);
        }

        /* 処理対象が無かった場合の通知 / Notify when nothing could be processed */
        if (targetTextCount === 0) {
            alert(getLabel("alert.noText"));
        }
    }

    main();

})();
