#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの行送り（表示値）が 整数で10ステップ小さく するように、自動行送り量（％）を逆算して設定します。
ダイアログは表示せず、実行するとその場で反映されます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoLeadingStep-10.md

### Overview

Works out the auto-leading percentage that decreases by ten whole steps the displayed leading of the selected text, and applies it.
There is no dialog; running it applies the change straight away.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoLeadingStep-10.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoLeadingStep-10";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoLeadingStep-10.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoLeadingStep-10.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // ステップ数はファイル名末尾の符号付き整数から取得する（例: AutoLeadingStep+10 → 10 / AutoLeadingStep-1 → -1）
    // 下記はファイル名から数値を読めなかったときの既定値
    // The step is read from the trailing signed integer of the file name (e.g. AutoLeadingStep+10 → 10);
    // this is the fallback used only when no number can be read.
    var DEFAULT_LEADING_STEP = 1;

    (function () {

        // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
        // ローカライズ（再利用パーツ） / Localization (reusable)
        //
        // 【移植手順 / How to port】
        // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
        //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
        // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
        //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
        // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
        //      getLabel("dialog.title")        … パス
        //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
        //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
        //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
        //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
        //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
        // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
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
                noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
                noSelection: { ja: "テキストが選択されていません。", en: "No text is selected." }
            }
        };

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
            return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
        }

        /* 実行中スクリプトのファイル名末尾の符号付き整数をステップ数として取得（読めなければ既定値）
           Read the trailing signed integer of the running script's file name as the step (fallback to default)
           例: AutoLeadingStep+10.jsx → 10 / AutoLeadingStep-1.jsx → -1 / AutoLeadingStep+1.jsx → 1
           @param {number} defaultStep ファイル名から数値を読めないときの既定値
           @returns {number} ステップ数（整数） */
        function getScriptStep(defaultStep) {
            try {
                var baseName = new File($.fileName).name.replace(/\.[^.]*$/, "");
                var matched = baseName.match(/([+\-]?\d+)\s*$/);
                if (matched) {
                    var parsed = parseInt(matched[1], 10);
                    if (!isNaN(parsed)) return parsed;
                }
            } catch (e) { }
            return defaultStep;
        }

        /* 型名を安全に取得 / Safely resolve a type name */
        function getTypeName(obj) {
            if (obj === null || obj === undefined) return "";
            if (obj.typename) return obj.typename;
            try { return obj.constructor ? obj.constructor.name : ""; } catch (e) { return ""; }
        }

        /* 親をたどって TextFrame を返す / Walk up parents to the enclosing TextFrame */
        function findParentTextFrame(item) {
            for (var i = 0; i < 20 && item; i++) {
                if (getTypeName(item) === "TextFrame") return item;
                try { item = item.parent; } catch (e) { return null; }
            }
            return null;
        }

        /* 選択から処理対象の TextFrame を収集(グループは再帰)
           Collect processable TextFrames from the selection (recurse into groups) */
        function collectTextFrames(item, frames) {
            if (!item) return;
            var typeName = getTypeName(item);
            if (typeName === "TextFrame") {
                if (item.contents && item.lines && item.lines.length > 0) frames.push(item);
            } else if (typeName === "GroupItem" && item.pageItems) {
                for (var i = 0; i < item.pageItems.length; i++) collectTextFrames(item.pageItems[i], frames);
            }
        }

        /* 範囲が触れている段落(段落全体)を対象配列へ追加 / Add the full paragraphs the range touches */
        function collectParagraphs(range, paragraphTargets) {
            try {
                var paragraphs = range.paragraphs;
                for (var i = 0; i < paragraphs.length; i++) paragraphTargets.push(paragraphs[i]);
            } catch (e) { }
        }

        /* 1段落の行送り(表示値)を整数で step だけ動かす自動行送り量(%)を逆算して適用
           Apply the auto-leading amount (%) that steps the paragraph's displayed leading by `step` integers
           @param {object} paragraph 段落範囲(Paragraph)
           @param {number} unitFactor 表示単位の pt 換算係数
           @param {number} step 行送り(表示単位・整数)のステップ数 */
        function applyLeadingStep(paragraph, unitFactor, step) {
            if (!paragraph.characters || paragraph.characters.length === 0) return;
            try {
                var charAttr = paragraph.characters[0].characterAttributes;
                var sizePt = charAttr.size;
                var leadingPt = charAttr.leading;
                if (isNaN(sizePt) || sizePt <= 0 || isNaN(leadingPt)) return;

                // 現在の行送りを表示単位に換算し、次の整数を目標にする(小さな誤差は吸収)
                // Convert the current leading to display units and target the next integer (absorb tiny float error)
                var currentInUnit = leadingPt / unitFactor;
                var base = (step >= 0) ? Math.floor(currentInUnit + 1e-4) : Math.ceil(currentInUnit - 1e-4);
                var targetInUnit = base + step;
                if (targetInUnit <= 0) return;

                // 目標の整数行送りになる自動行送り量(%)を逆算 / Back-calculate the auto-leading amount (%) for the target integer leading
                var targetLeadingPt = targetInUnit * unitFactor;
                paragraph.paragraphAttributes.autoLeadingAmount = (targetLeadingPt / sizePt) * 100;
                paragraph.characterAttributes.autoLeading = true;
            } catch (e) { }
        }

        function main() {
            if (app.documents.length === 0) {
                alert(getLabel("alert.noDocument"));
                return;
            }

            var selection = app.activeDocument.selection;
            var paragraphTargets = []; // 対象の段落範囲 / Target paragraph ranges
            var typeFrames = [];       // leadingType を設定するフレーム / Frames to set leadingType on

            if (getTypeName(selection) === "TextRange") {
                // テキスト編集モード：選択が触れている段落だけを対象(一部の文字選択でも段落全体に適用)
                // Text-edit mode: target only the paragraphs the selection touches (partial char selection → whole paragraph)
                collectParagraphs(selection, paragraphTargets);
                var editFrame = findParentTextFrame(selection);
                if (editFrame) typeFrames.push(editFrame);
            } else {
                // 選択ツール：選択したフレーム(グループ内含む)の全段落を対象
                // Selection tool: target every paragraph of the selected frames (including those inside groups)
                var frames = [];
                var items = selection || [];
                for (var i = 0; i < items.length; i++) collectTextFrames(items[i], frames);
                for (var f = 0; f < frames.length; f++) {
                    collectParagraphs(frames[f].textRange, paragraphTargets);
                    typeFrames.push(frames[f]);
                }
            }

            if (paragraphTargets.length === 0) {
                alert(getLabel("alert.noSelection"));
                return;
            }

            // ステップ数はファイル名から取得（例: AutoLeadingStep+10 → 10）/ The step comes from the file name (e.g. +10)
            var step = getScriptStep(DEFAULT_LEADING_STEP);
            // 段落ごとに行送りを整数で step 動かす / Step each paragraph's leading by `step` integers
            var unitFactor = getUnitInfo("text/units").pointsPerUnit;
            for (var p = 0; p < paragraphTargets.length; p++) applyLeadingStep(paragraphTargets[p], unitFactor, step);
            // 基準は仮想ボディの上に固定(フレーム単位) / Fix the leading basis to the top of the virtual body (per frame)
            for (var g = 0; g < typeFrames.length; g++) {
                try { typeFrames[g].textRange.leadingType = AutoLeadingType.TOPTOTOP; } catch (e) { }
            }
            app.redraw();
        }

        main();

    })();

})();
