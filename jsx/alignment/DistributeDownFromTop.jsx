#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択内容に応じて、行送り（leading）と配置を調整します。
最も上のオブジェクトを固定し、以降を下方向へ広げます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeDownFromTop.md

### Overview

Adjusts the leading and the placement according to what is selected.
The topmost object stays fixed and the rest spread downwards.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeDownFromTop.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DistributeDownFromTop";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeDownFromTop.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DistributeDownFromTop.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 「Y座標がほぼ同じ（横並び）」とみなす上端Yの許容差（pt）/ tolerance for treating tops as one row */
    var SAME_Y_TOLERANCE_PT = 2.0;

    /* 行送りがこの差以内なら「すでに統一済み」とみなし、再実行で行送りを増やす（pt）
       once leadings are this close, a re-run increases them instead of unifying */
    var LEADING_UNIFORM_TOLERANCE_PT = 0.01;

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
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 判定 / Checks
    // =========================================

    /**
     * 選択がすべて TextFrame かどうかを返す
     * @param {PageItem[]} targetItems - 判定するオブジェクト
     * @returns {boolean} すべて TextFrame なら true
     */
    function areAllTextFrames(targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            if (targetItems[i].typename !== "TextFrame") return false;
        }
        return true;
    }

    /**
     * 上端Y（geometricBounds[1]）の最大差が許容内なら「横並び（同じ行）」とみなす
     * position はベースライン基準でズレるため geometricBounds を使う。
     * @param {PageItem[]} targetItems - 判定するオブジェクト
     * @param {number} tolerancePt - 同じ行とみなす許容差（pt）
     * @returns {boolean} 横並びとみなせるなら true
     */
    function areTopEdgesAligned(targetItems, tolerancePt) {
        var maxTopY = targetItems[0].geometricBounds[1];
        var minTopY = maxTopY;
        for (var i = 1; i < targetItems.length; i++) {
            var topY = targetItems[i].geometricBounds[1];
            if (topY > maxTopY) maxTopY = topY;
            if (topY < minTopY) minTopY = topY;
        }
        return (maxTopY - minTopY) <= tolerancePt;
    }

    /**
     * 全テキストの行送りが許容内で揃っているか（揃っていれば再実行とみなす）
     * @param {TextFrame[]} textFrames - 判定するテキストフレーム
     * @param {number} tolerancePt - 揃っているとみなす許容差（pt）
     * @returns {boolean} 揃っていれば true
     */
    function areLeadingsUniform(textFrames, tolerancePt) {
        var maxLeading = textFrames[0].textRange.characterAttributes.leading;
        var minLeading = maxLeading;
        for (var i = 1; i < textFrames.length; i++) {
            var leading = textFrames[i].textRange.characterAttributes.leading;
            if (leading > maxLeading) maxLeading = leading;
            if (leading < minLeading) minLeading = leading;
        }
        return (maxLeading - minLeading) <= tolerancePt;
    }

    // =========================================
    // 行送り / Leading
    // =========================================

    /**
     * 全テキストの行送りを deltaPt 分だけずらす（位置は動かさない）／自動行送り量（％）で適用
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {number} deltaPt - 行送りの増減量（pt。負値で詰まる）
     * @returns {void}
     */
    function shiftLeadingBy(textFrames, deltaPt) {
        for (var i = 0; i < textFrames.length; i++) {
            applyLeadingAsAutoLeading(textFrames[i], function (currentLeadingPt) {
                return currentLeadingPt + deltaPt;
            });
        }
    }

    /**
     * 各テキストの行送りを平均値に統一する（位置は動かさない）／自動行送り量（％）で適用
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {void}
     */
    function unifyLeadingToAverage(textFrames) {
        var totalLeading = 0;
        for (var i = 0; i < textFrames.length; i++) {
            /* 自動行送りのときは算出値が返る / Auto leading returns the computed value */
            totalLeading += textFrames[i].textRange.characterAttributes.leading;
        }
        var averageLeading = totalLeading / textFrames.length;
        for (var j = 0; j < textFrames.length; j++) {
            applyLeadingAsAutoLeading(textFrames[j], function () {
                return averageLeading;
            });
        }
    }

    /**
     * 目標行送り（pt）を自動行送り量（％）に逆算して各段落へ設定する（手動行送りは使わない）
     * 各段落の先頭文字のフォントサイズを基準に「目標行送り ÷ サイズ × 100」で autoLeadingAmount を求め、
     * autoLeading=true・基準を TOPTOTOP に固定する。
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {function} resolveTargetLeadingPt - 現在の行送り(pt)を受け取り目標行送り(pt)を返す関数
     * @returns {void}
     */
    function applyLeadingAsAutoLeading(textFrame, resolveTargetLeadingPt) {
        var paragraphs = textFrame.textRange.paragraphs;
        for (var i = 0; i < paragraphs.length; i++) {
            var paragraph = paragraphs[i];
            if (!paragraph.characters || paragraph.characters.length === 0) continue;

            var charAttributes = paragraph.characters[0].characterAttributes;
            var fontSizePt = charAttributes.size;
            var currentLeadingPt = charAttributes.leading; /* 自動行送りのときは算出値が返る / computed value under auto leading */
            if (isNaN(fontSizePt) || fontSizePt <= 0 || isNaN(currentLeadingPt)) continue;

            var targetLeadingPt = resolveTargetLeadingPt(currentLeadingPt);
            if (isNaN(targetLeadingPt) || targetLeadingPt <= 0) continue;

            paragraph.paragraphAttributes.autoLeadingAmount = (targetLeadingPt / fontSizePt) * 100;
            paragraph.characterAttributes.autoLeading = true;
        }
        /* leadingType は環境によって TextRange 側に無いことがあるため握る / not exposed on every build */
        try {
            textFrame.textRange.leadingType = AutoLeadingType.TOPTOTOP;
        } catch (e) {}
    }

    // =========================================
    // 選択と並べ替え / Selection and sorting
    // =========================================

    /**
     * 上端Yの降順（上から下）に並べ替えた新しい配列を返す
     * @param {PageItem[]} targetObjects - 並べ替える対象のオブジェクト
     * @returns {PageItem[]} 上から下の順に並べ替えた新しい配列
     */
    function sortTopToBottom(targetObjects) {
        var sortedObjects = [];
        for (var i = 0; i < targetObjects.length; i++) sortedObjects.push(targetObjects[i]);
        sortedObjects.sort(function (itemA, itemB) {
            return itemB.position[1] - itemA.position[1];
        });
        return sortedObjects;
    }

    /**
     * 選択中の TextRange を含む単一の TextFrame を選択し直す
     * @returns {void}
     */
    function selectSingleTextFrameFromTextRange() {
        if (app.selection.constructor.name !== "TextRange") return;

        var textFramesInStory = app.selection.story.textFrames;
        if (textFramesInStory.length !== 1) return;

        app.executeMenuCommand("deselectall");  /* 現在の選択を解除 / clear the caret selection */
        app.selection = [textFramesInStory[0]]; /* 該当の TextFrame を選択 / select that frame */
        app.selectTool("Adobe Select Tool");    /* 選択ツールに戻す / back to the selection tool */
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択内容に応じて、行送りの調整または縦方向の等間隔配置を行う
     * @returns {void}
     */
    function main() {
        if (app.documents.length < 1) return;

        /* テキスト範囲（カーソル）を選択しているときは、その story の単一 TextFrame を選択し直す / Reselect the single frame of a caret selection */
        selectSingleTextFrameFromTextRange();

        var selectedObjects = app.activeDocument.selection;
        if (selectedObjects.length < 1) return;

        /* 「サイズ／行送り」キー増加（text/sizeIncrement）を表示単位（text/units）込みで pt 換算 / Size/Leading increment in points */
        var leadingStepPt = app.preferences.getRealPreference("text/sizeIncrement") * getUnitInfo("text/units").pointsPerUnit;

        /* テキストを1つだけ選択 → 行送りを「サイズ／行送り」分増やす / One text frame: increase its leading by the increment */
        if (selectedObjects.length === 1 && selectedObjects[0].typename === "TextFrame") {
            shiftLeadingBy([selectedObjects[0]], leadingStepPt);
            return;
        }

        if (selectedObjects.length < 2) return;

        /* 全てテキストで、上端Yがほぼ同じ（横並び）→ 位置は動かさず行送りを調整 / Text frames side by side: adjust leading, keep positions */
        if (areAllTextFrames(selectedObjects) && areTopEdgesAligned(selectedObjects, SAME_Y_TOLERANCE_PT)) {
            if (areLeadingsUniform(selectedObjects, LEADING_UNIFORM_TOLERANCE_PT)) {
                /* 再実行 → 「複数行の1テキスト」のように全体の行送りを増やす / Re-run: increase every leading */
                shiftLeadingBy(selectedObjects, leadingStepPt);
            } else {
                /* 初回 → 行送りを平均値に統一 / First run: unify to the average leading */
                unifyLeadingToAverage(selectedObjects);
            }
            return;
        }

        /* それ以外（縦積み）→ 最上部を固定し、以降を leadingStepPt ずつ下へ等間隔配置 / Stacked: keep the top one, move the rest down step by step */
        var objectsTopToBottom = sortTopToBottom(selectedObjects);
        for (var i = 1; i < objectsTopToBottom.length; i++) {
            objectsTopToBottom[i].translate(0, -i * leadingStepPt);
        }
    }

    main();

})();
