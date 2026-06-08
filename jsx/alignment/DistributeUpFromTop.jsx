#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// =========================================
// 概要 / Overview
// =========================================
/*
DistributeUpFromTop.jsx

複数オブジェクトを選択しているとき:
  最上部のオブジェクトを固定し、以降を「サイズ／行送り」の値ぶんずつ
  上方向へ等間隔に再配置する。

テキストを 1 つだけ選択しているとき:
  そのテキストの行送りから「サイズ／行送り」の値を引く。

移動・減算に使う値は、環境設定［テキスト］の「サイズ／行送り」キー入力増分
（text/sizeIncrement）を、表示単位（text/units）込みで pt 換算したもの。
*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.0.0";

(function () {
    if (app.documents.length < 1) return
    var selection = app.activeDocument.selection
    if (selection.length < 1) return

    // 「サイズ／行送り」キー入力（text/sizeIncrement）を表示単位（text/units）込みで pt 換算
    var textUnit = app.preferences.getIntegerPreference("text/units")
    var sizeLeadingStep = app.preferences.getRealPreference("text/sizeIncrement") * pointsPerTextUnit(textUnit)

    // テキストを1つだけ選択 → 行送りから「サイズ／行送り」分を引く
    if (selection.length === 1 && selection[0].typename === "TextFrame") {
        var attributes = selection[0].textRange.characterAttributes
        attributes.leading = attributes.leading - sizeLeadingStep
        return
    }

    // 複数選択 → 最上部を固定し、以降を sizeLeadingStep ずつ上へ等間隔配置
    if (selection.length < 2) return
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

    function sortByVerticalPosition(selection) {
        var sorted = []
        for (var i = 0; i < selection.length; i++) sorted.push(selection[i])
        sorted.sort(function (a, b) {
            return b.position[1] - a.position[1]
        })
        return sorted
    }
})()
