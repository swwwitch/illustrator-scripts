#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SimplifyGroups.jsx

### 概要

- 選択したグループ内のサブグループを再帰的に解除します。
- 最外層のグループは解除せず残します。
- 選択にグループが1つだけあり、他に非グループがある場合は非グループをそのグループに追加します。
- 複数のグループと非グループが混在している場合はまとめて1つのグループにしてから処理します。

### 主な機能

- サブグループの再帰的解除
- 混在選択時の自動グループ化
- 既存グループへの非グループ追加
- Illustratorメニュー「グループ解除」コマンドの利用

### 処理の流れ

1. ドキュメントと選択を確認
2. 選択にグループが1つだけの場合は非グループをそのグループに追加
3. 複数のグループと非グループが混在していればまとめてグループ化
4. グループ内のサブグループを再帰的に探索し解除

### note

https://note.com/dtp_tranist/n/n45797beb72bb

### 更新履歴

- v1.0 (20250707) : 初版公開
- v1.1 (20250707) : 実行後の選択状態を調整
- v1.2 (20250707) : 非グループを既存グループに追加する際の挙動を変更
- v1.3 (20250707) : クリップグループに選択オブジェクトを追加する機能を追加

---

### Script Name:

SimplifyGroups.jsx

### Overview

- Recursively ungroups subgroups inside the selected group.
- The outermost group remains intact.
- If there is only one group in the selection and other non-group objects, the non-groups are added to that group.
- If groups and non-groups are mixed, they are grouped together before processing.

### Main Features

- Recursive ungrouping of subgroups
- Auto-grouping when mixed selection
- Adding non-groups to existing single group
- Uses Illustrator's "ungroup" menu command internally

### Process Flow

1. Check document and selection
2. If only one group selected with other non-groups, add non-groups to that group
3. Group mixed selection if needed
4. Recursively find and ungroup subgroups

### Change Log

- v1.0 (20250707): Initial release
- v1.１ (20250707): Added feature to add non-groups to existing single group
- v1.2 (20250707): Improved behavior when adding non-groups to existing group
- v1.3 (20250707): Added functionality to add selected objects to clip groups

*/

function main() {
    // ドキュメントと選択チェック
    if (app.documents.length === 0) return;
    var sel = app.activeDocument.selection;
    if (!sel || sel.length === 0) return;

    // グループと非グループの数を数える
    var groupCount = 0, nonGroupCount = 0;
    var firstGroup = null;
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "GroupItem") {
            groupCount++;
            if (!firstGroup) firstGroup = sel[i];
        } else {
            nonGroupCount++;
        }
    }

    // クリップグループとオブジェクトが選択されている場合
    if (groupCount === 1 && nonGroupCount > 0 && firstGroup.clipped) {
        for (var i = sel.length - 1; i >= 0; i--) {
            var item = sel[i];
            if (item !== firstGroup && !item.locked && !item.hidden) {
                // クリップグループの場合は無条件で先頭に配置
                item.move(firstGroup, ElementPlacement.PLACEATBEGINNING);
            }
        }
        bringMaskPathToFront(firstGroup);
        ungroupSubGroups(firstGroup);
        app.selection = [firstGroup]; // 最終的に残ったグループを選択
        return;
    }

    // 1つだけグループがあり、他が非グループの場合
    if (groupCount === 1 && nonGroupCount > 0) {
        for (var i = sel.length - 1; i >= 0; i--) {
            var item = sel[i];
            if (item !== firstGroup && !item.locked && !item.hidden) {
                var posA = firstGroup.zOrderPosition;
                var posB;
                try {
                    posB = item.zOrderPosition;
                } catch (e) {
                    posB = posA + 1; // デフォルトで A より背面扱い
                }

                if (posB < posA) {
                    item.move(firstGroup, ElementPlacement.PLACEATBEGINNING);
                } else {
                    item.move(firstGroup, ElementPlacement.PLACEATEND);
                }
            }
        }
        ungroupSubGroups(firstGroup);
        app.selection = [firstGroup]; // 最終的に残ったグループを選択
        return;
    }

    // 複数のグループと非グループが混在している場合
    if (groupCount > 0 && nonGroupCount > 0) {
        app.executeMenuCommand("group");
        var newSel = app.activeDocument.selection;
        if (newSel.length === 1 && newSel[0].typename === "GroupItem") {
            ungroupSubGroups(newSel[0]);
            app.selection = [newSel[0]];
        }
        return;
    }

    // 選択内の各グループに対して処理
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename === "GroupItem") {
            ungroupSubGroups(item);
            app.selection = [item];
        }
    }
}

// サブグループを再帰的に解除
function ungroupSubGroups(group) {
    for (var i = group.pageItems.length - 1; i >= 0; i--) {
        var item = group.pageItems[i];
        if (item.typename === "GroupItem") {
            ungroupSubGroups(item);
            app.selection = [item];
            app.executeMenuCommand("ungroup");
        }
    }
}

function bringMaskPathToFront(clipGroup) {
  if (!(clipGroup instanceof GroupItem) || !clipGroup.clipped) {
    alert("有効なクリップグループを指定してください。");
    return;
  }

  var maskPath = null;

  // クリップグループ内のマスクパスを探す
  for (var i = 0; i < clipGroup.pageItems.length; i++) {
    var item = clipGroup.pageItems[i];
    if (item.clipping) {
      maskPath = item;
      break;
    }
  }

  if (!maskPath) {
    alert("マスクパスが見つかりません。");
    return;
  }

  // マスクパスをグループ内の最上位に移動
  maskPath.zOrder(ZOrderMethod.BRINGTOFRONT);
  // alert("マスクパスを最上位に移動しました。");
}

main();