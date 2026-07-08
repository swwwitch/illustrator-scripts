#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// 概要 / Overview
// =========================================
/*
DistributeUpFromTop.jsx

選択内容に応じて、行送り（leading）と配置を調整する。判定は以下の順で行う。

(1) テキストを 1 つだけ選択しているとき:
    そのテキストの行送りから「サイズ／行送り」の値を引く。

(2) 複数のテキストを選択していて、上端 Y がほぼ同じ（横並び）のとき:
    位置は動かさず、行送りだけを調整する。
      ・初回（行送りがバラバラ） → 全テキストの行送りを平均値に統一する。
      ・再実行（行送りが揃っている）→「複数行の 1 テキスト」のように、
        全テキストの行送りを「サイズ／行送り」分ずつ減らす。

(3) それ以外の複数選択（縦積み）:
    最上部のオブジェクトを固定し、以降を「サイズ／行送り」の値ぶんずつ
    上方向へ等間隔に再配置する。

移動・減算に使う値は、環境設定［テキスト］の「サイズ／行送り」キー入力増分
（text/sizeIncrement）を、表示単位（text/units）込みで pt 換算したもの。

行送りの適用は手動行送りではなく、目標行送りから自動行送り量（autoLeadingAmount ％）を
逆算して設定する（autoLeading=true、基準は TOPTOTOP に固定）。段落ごとに先頭文字の
フォントサイズを基準に算出する。
*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.3.0";

// 「Y座標がほぼ同じ（横並び）」とみなす上端 Y の許容差（pt）
var SAME_Y_TOLERANCE_PT = 2.0;

// 行送りがこの差以内なら「すでに統一済み」とみなし、再実行で行送りを減らす（pt）
var LEADING_UNIFORM_TOLERANCE_PT = 0.01;

(function () {
    if (app.documents.length < 1) return

    // テキスト範囲（カーソル）を選択しているときは、その story の単一 TextFrame を選択し直す
    selectSingleTextFrameFromTextRange()

    var selection = app.activeDocument.selection
    if (selection.length < 1) return

    // 「サイズ／行送り」キー入力（text/sizeIncrement）を表示単位（text/units）込みで pt 換算
    var textUnit = app.preferences.getIntegerPreference("text/units")
    var sizeLeadingStep = app.preferences.getRealPreference("text/sizeIncrement") * pointsPerTextUnit(textUnit)

    // テキストを1つだけ選択 → 行送りから「サイズ／行送り」分を引いた値を自動行送り量（％）で適用
    if (selection.length === 1 && selection[0].typename === "TextFrame") {
        applyLeadingAsAutoLeading(selection[0], function (currentLeadingPt) {
            return currentLeadingPt - sizeLeadingStep
        })
        return
    }

    if (selection.length < 2) return

    // 全てテキストで、上端 Y がほぼ同じ（横並び）→ 位置は動かさず行送りを調整
    if (allTextFrames(selection) && topYNearlySame(selection, SAME_Y_TOLERANCE_PT)) {
        if (leadingsAreUniform(selection, LEADING_UNIFORM_TOLERANCE_PT)) {
            // 再実行 → 「複数行の1テキスト」のように全体の行送りを減らす
            decreaseLeading(selection, sizeLeadingStep)
        } else {
            // 初回 → 行送りを平均値に統一
            unifyLeadingToAverage(selection)
        }
        return
    }

    // それ以外（縦積み）→ 最上部を固定し、以降を sizeLeadingStep ずつ上へ等間隔配置
    var selectedItems = sortByVerticalPosition(selection)
    for (var i = 1; i < selectedItems.length; i++) {
        selectedItems[i].translate(0, i * sizeLeadingStep)
    }

    // 環境設定［テキスト］の単位（text/units）を pt へ換算する係数
    // 0=inch, 1=mm, 2=pt, 3=pica, 4=cm, 5=Q, 6=px
    function pointsPerTextUnit(unitType) {
        if (unitType === 0) return 72                // inch
        if (unitType === 1) return 72 / 25.4         // mm
        if (unitType === 3) return 12                // pica
        if (unitType === 4) return 72 / 2.54         // cm
        if (unitType === 5) return 72 / 25.4 * 0.25  // Q（1Q = 0.25mm）
        return 1                                     // pt / px / 既定
    }

    // 選択がすべて TextFrame かどうか
    function allTextFrames(items) {
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename !== "TextFrame") return false
        }
        return true
    }

    // 上端 Y（geometricBounds[1]）の最大差が許容内なら「横並び（同じ行）」とみなす
    // ※ position はベースライン基準でズレるため geometricBounds を使う
    function topYNearlySame(items, tolerancePt) {
        var maxTop = items[0].geometricBounds[1]
        var minTop = maxTop
        for (var i = 1; i < items.length; i++) {
            var top = items[i].geometricBounds[1]
            if (top > maxTop) maxTop = top
            if (top < minTop) minTop = top
        }
        return (maxTop - minTop) <= tolerancePt
    }

    // 全テキストの行送りが許容内で揃っているか（揃っていれば再実行とみなす）
    function leadingsAreUniform(textFrames, tolerancePt) {
        var maxLeading = textFrames[0].textRange.characterAttributes.leading
        var minLeading = maxLeading
        for (var i = 1; i < textFrames.length; i++) {
            var leading = textFrames[i].textRange.characterAttributes.leading
            if (leading > maxLeading) maxLeading = leading
            if (leading < minLeading) minLeading = leading
        }
        return (maxLeading - minLeading) <= tolerancePt
    }

    // 全テキストの行送りを stepPt 分だけ減らす（位置は動かさない）／自動行送り量（％）で適用
    function decreaseLeading(textFrames, stepPt) {
        for (var i = 0; i < textFrames.length; i++) {
            applyLeadingAsAutoLeading(textFrames[i], function (currentLeadingPt) {
                return currentLeadingPt - stepPt
            })
        }
    }

    // 各テキストの行送りを平均値に統一する（位置は動かさない）／自動行送り量（％）で適用
    function unifyLeadingToAverage(textFrames) {
        var totalLeading = 0
        for (var i = 0; i < textFrames.length; i++) {
            // 自動行送りのときは算出値が返る
            totalLeading += textFrames[i].textRange.characterAttributes.leading
        }
        var averageLeading = totalLeading / textFrames.length
        for (var j = 0; j < textFrames.length; j++) {
            applyLeadingAsAutoLeading(textFrames[j], function () {
                return averageLeading
            })
        }
    }

    /* 目標行送り（pt）を自動行送り量（％）に逆算して各段落へ設定する（手動行送りは使わない）
       各段落の先頭文字のフォントサイズを基準に「目標行送り ÷ サイズ × 100」で autoLeadingAmount を求め、
       autoLeading=true・基準を TOPTOTOP に固定する。目標行送りは段落ごとに resolveTargetLeadingPt で算出。
       @param {TextFrame} textFrame 対象のテキストフレーム
       @param {function} resolveTargetLeadingPt 現在の行送り(pt)を受け取り目標行送り(pt)を返す関数 */
    function applyLeadingAsAutoLeading(textFrame, resolveTargetLeadingPt) {
        var paragraphs = textFrame.textRange.paragraphs
        for (var i = 0; i < paragraphs.length; i++) {
            var paragraph = paragraphs[i]
            if (!paragraph.characters || paragraph.characters.length === 0) { continue; }
            var charAttr = paragraph.characters[0].characterAttributes
            var fontSizePt = charAttr.size
            var currentLeadingPt = charAttr.leading // 自動行送りのときは算出値が返る
            if (isNaN(fontSizePt) || fontSizePt <= 0 || isNaN(currentLeadingPt)) { continue; }
            var targetLeadingPt = resolveTargetLeadingPt(currentLeadingPt)
            if (isNaN(targetLeadingPt) || targetLeadingPt <= 0) { continue; }
            paragraph.paragraphAttributes.autoLeadingAmount = (targetLeadingPt / fontSizePt) * 100
            paragraph.characterAttributes.autoLeading = true
        }
        try { textFrame.textRange.leadingType = AutoLeadingType.TOPTOTOP } catch (e) {}
    }

    function sortByVerticalPosition(selection) {
        var sorted = []
        for (var i = 0; i < selection.length; i++) sorted.push(selection[i])
        sorted.sort(function (a, b) {
            return b.position[1] - a.position[1]
        })
        return sorted
    }

    // 選択中の TextRange を含む単一の TextFrame を選択し直す
    function selectSingleTextFrameFromTextRange() {
        if (app.selection.constructor.name === "TextRange") {
            var textFramesInStory = app.selection.story.textFrames
            if (textFramesInStory.length === 1) {
                app.executeMenuCommand("deselectall") // 現在の選択を解除
                app.selection = [textFramesInStory[0]] // 該当の TextFrame を選択
                try {
                    app.selectTool("Adobe Select Tool") // 選択ツールに戻す
                } catch (e) {}
            }
        }
    }
})()
