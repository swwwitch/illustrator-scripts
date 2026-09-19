#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したグループの中にあるサブグループを再帰的に解除し、最外層のグループだけを残します。
グループと非グループが混在している場合は、まとめて1つのグループにしてから処理します。

詳細は README を参照してください。

### Overview

Recursively ungroups the subgroups inside the selection, leaving only the outermost group.
When groups and non-groups are mixed, they are grouped together first.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SimplifyGroups";               /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-07";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-07-07";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SimplifyGroups.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SimplifyGroups.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n45797beb72bb"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// zOrderPosition の増減方向を実ドキュメント上で判定
// true: 数値が大きいほど前面 / false: 数値が小さいほど前面
function detectLargerIsFront(parent) {
    var doc = app.activeDocument;
    var t = null;
    try {
        // できるだけ影響の少ない極小パスを作って即削除
        t = doc.pathItems.rectangle(0, 0, 1, 1);
        t.stroked = false;
        t.filled = false;
        t.opacity = 0;
        t.move(parent, ElementPlacement.PLACEATEND);
        t.zOrder(ZOrderMethod.SENDTOBACK);
        var backPos = t.zOrderPosition;
        t.zOrder(ZOrderMethod.BRINGTOFRONT);
        var frontPos = t.zOrderPosition;
        return frontPos > backPos;
    } catch (e) {
        // 失敗時は「大きいほど前面」寄りに倒す（現場の実挙動に合わせやすい）
        return true;
    } finally {
        try { if (t) t.remove(); } catch (e2) {}
    }
}

function isBInFrontOfA(posB, posA, largerIsFront) {
    return largerIsFront ? (posB > posA) : (posB < posA);
}

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
    var largerIsFront = detectLargerIsFront(firstGroup.parent);

    // クリップグループとオブジェクトが選択されている場合
    if (groupCount === 1 && nonGroupCount > 0 && firstGroup.clipped) {
        for (var i = sel.length - 1; i >= 0; i--) {
            var item = sel[i];
            if (item !== firstGroup && !item.locked && !item.hidden) {
                var posA = firstGroup.zOrderPosition;
                var posB;
                try {
                    posB = item.zOrderPosition;
                } catch (e) {
                    posB = posA + 1; // 取得できない場合は背面寄り扱い
                }

                var bIsFront = isBInFrontOfA(posB, posA, largerIsFront);

                // B が A より前面なら → グループの最前面（END）
                // B が A より背面なら → グループの最背面（BEGINNING）
                item.move(firstGroup, bIsFront ? ElementPlacement.PLACEATEND : ElementPlacement.PLACEATBEGINNING);
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

                var bIsFront = isBInFrontOfA(posB, posA, largerIsFront);

                // B が A より前面なら → グループの最前面（END）
                // B が A より背面なら → グループの最背面（BEGINNING）
                item.move(firstGroup, bIsFront ? ElementPlacement.PLACEATEND : ElementPlacement.PLACEATBEGINNING);
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