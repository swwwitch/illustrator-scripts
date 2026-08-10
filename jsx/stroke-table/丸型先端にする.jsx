// スクリプトの説明
// 作成日：2024年8月22日
// 更新日：2026年4月12日
// このスクリプトは、選択されているパスアイテム（グループ内も含む）の
// 線端を突出先端（プロジェクティングキャップ）に設定します。

// パスアイテムに突出先端を適用する再帰関数
function applyProjectingCap(item) {
    if (item.typename === "GroupItem") {
        // グループの場合、子アイテムを再帰的に処理
        for (var i = 0; i < item.pageItems.length; i++) {
            applyProjectingCap(item.pageItems[i]);
        }
    } else if (item.typename === "PathItem" || item.typename === "CompoundPathItem") {
        if (item.stroked && item.strokeCap !== StrokeCap.ROUNDENDCAP) {
            item.strokeCap = StrokeCap.ROUNDENDCAP;
        }
    }
}

if (app.documents.length > 0) {
    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length > 0) {
        for (var i = 0; i < sel.length; i++) {
            applyProjectingCap(sel[i]);
        }
        // alert("選択されたパスアイテムの線端を突出先端に設定しました。");
    } else {
        alert("パスアイテムが選択されていません。パスアイテムを選択してください。");
    }
} else {
    alert("ドキュメントが開かれていません。ドキュメントを開いてください。");
}
