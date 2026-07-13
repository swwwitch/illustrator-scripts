#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの文字ツメを30%に設定する。
カーニング方式・プロポーショナルメトリクスは変更しない。
選択範囲全体に即適用し、未選択時は何もしない。

*/

(function () {

    /* 型名を安全に取得（host オブジェクトは typename を優先、JS オブジェクトは constructor.name） */
    function getTypeName(obj) {
        if (obj === null || obj === undefined) return "";
        if (obj.typename) return obj.typename;
        try {
            return obj.constructor ? obj.constructor.name : "";
        } catch (e) {
            return "";
        }
    }

    /* 選択中のテキスト範囲を取得 */
    function getSelectedTextRanges() {
        var activeDoc = app.activeDocument;
        var currentSelection = activeDoc.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる */
        if (getTypeName(currentSelection) === "TextRange") {
            selectedRanges.push(currentSelection);
            return selectedRanges;
        }
        if (currentSelection.length === 0) return selectedRanges;
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            var itemType = getTypeName(selectedItem);
            if (itemType === "TextFrame") {
                selectedRanges.push(selectedItem.textRange);
            } else if (itemType === "TextRange") {
                selectedRanges.push(selectedItem);
            }
        }
        return selectedRanges;
    }

    /* 選択範囲に文字ツメを適用（0〜100 の百分率。30 = 30%） */
    function applyTsumeToRanges(ranges, tsumePercent) {
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.Tsume = tsumePercent;
            } catch (e) {
                // 適用できない範囲はスキップ
            }
        }
    }

    function main() {
        if (app.documents.length <= 0) return;
        var targetRanges = getSelectedTextRanges();
        if (targetRanges.length === 0) return;
        applyTsumeToRanges(targetRanges, 30);
    }

    main();

})();
