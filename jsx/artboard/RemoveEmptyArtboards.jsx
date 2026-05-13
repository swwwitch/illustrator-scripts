#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
    概要 / Overview
    -----------------------------------------------------------------
    アートボード上に表示オブジェクトが存在しないアートボードを検出し、
    まとめて削除するスクリプトです。オプションで非表示のレイヤー／
    オブジェクトを「占有あり」とみなすかを切り替えられます（既定 ON：
    無視）。トグルに応じて対象数表示も即時更新します。
    非表示判定は祖先を遡って行うため、グループ内の非表示アイテムや
    非表示グループ／レイヤー配下のアイテムも同じ扱いになります。

    Detects artboards that contain no occupants and removes them in a
    single pass. A checkbox toggles whether hidden layers/items are
    counted as occupants (ON by default = ignore hidden). The count
    summary updates live with the checkbox state. Visibility is checked
    up the ancestor chain, so items nested inside hidden groups or
    layers (and individually hidden items inside visible groups) are
    handled the same way.
*/

// =========================================
// バージョンとローカライズ / Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.0.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "空アートボードの削除", en: "Remove Empty Artboards" },
    ignoreHidden: { ja: "非表示レイヤー、非表示オブジェクトは無視", en: "Ignore hidden layers and hidden objects" },
    removalTarget: { ja: "削除対象アートボード数", en: "Artboards to remove" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
    onlyOneArtboard: { ja: "アートボードが1枚しかないため削除できません。", en: "Only one artboard exists; cannot remove." }
};

/* ラベル取得 / Get localized label by key */
function L(key) {
    return LABELS[key][lang];
}

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return L(key) + (lang === 'ja' ? '：' : ':');
}

/* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
function labelWithCount(key, count) {
    if (lang === 'ja') {
        return L(key) + '（' + count + '）';
    }
    return L(key) + ' (' + count + ')';
}

// =========================================
// アートボード判定 / Artboard detection
// =========================================

/* 2 矩形の交差判定 / Test whether two rectangles overlap.
   矩形は [left, top, right, bottom] / Rectangles are [left, top, right, bottom] */
function rectsOverlap(artboardRect, itemBounds) {
    return (
        itemBounds[0] < artboardRect[2] &&
        itemBounds[2] > artboardRect[0] &&
        itemBounds[1] > artboardRect[3] &&
        itemBounds[3] < artboardRect[1]
    );
}

/* オブジェクトと祖先がすべて表示状態か / Walk ancestry: every ancestor must be visible */
function isItemVisible(pageItem) {
    try {
        var node = pageItem;
        while (node && node.typename !== "Document") {
            if (node.typename === "Layer") {
                if (node.visible === false) return false;
            } else if (node.hidden) {
                return false;
            }
            node = node.parent;
        }
    } catch (err) {
        return false;
    }
    return true;
}

/* アートボード上に占有オブジェクトが存在するか / Any occupant touching this artboard?
   ignoreHidden=true なら非表示は無視（visibleBounds）/ When ignoreHidden, skip hidden items.
   ignoreHidden=false なら非表示も占有とみなす（geometricBounds）/ Otherwise hidden items count too. */
function hasOccupantOnArtboard(doc, artboardRect, ignoreHidden) {
    var pageItems = doc.pageItems;
    for (var i = 0; i < pageItems.length; i++) {
        var pageItem = pageItems[i];
        var visible = isItemVisible(pageItem);
        if (ignoreHidden && !visible) continue;
        try {
            var bounds = visible ? pageItem.visibleBounds : pageItem.geometricBounds;
            if (rectsOverlap(artboardRect, bounds)) return true;
        } catch (err) {
            // bounds が取得できない場合は無視 / Skip items without bounds
        }
    }
    return false;
}

/* 空アートボードのインデックス一覧 / Collect indices of empty artboards */
function findEmptyArtboardIndices(doc, ignoreHidden) {
    var indices = [];
    var artboards = doc.artboards;
    for (var i = 0; i < artboards.length; i++) {
        if (!hasOccupantOnArtboard(doc, artboards[i].artboardRect, ignoreHidden)) {
            indices.push(i);
        }
    }
    return indices;
}

// =========================================
// アートボード削除 / Artboard removal
// =========================================

/* 指定インデックスのアートボードを削除（最後の1枚は残す）
   Remove artboards at the given indices (must leave at least one) */
function removeArtboardsByIndices(doc, indices) {
    var artboards = doc.artboards;
    var removedCount = 0;
    for (var i = indices.length - 1; i >= 0; i--) {
        if (artboards.length <= 1) break;
        try {
            artboards[indices[i]].remove();
            removedCount++;
        } catch (err) { }
    }
    return removedCount;
}

// =========================================
// ダイアログ / Dialog
// =========================================

/* オプションダイアログ / Show options dialog, return user choice or null on cancel.
   対象数が 0 のときは OK をディム表示 / OK is disabled when count is 0.
   countIgnore: 非表示を無視した場合の対象数 / count when ignoring hidden
   countAll:    非表示も占有とみなした場合の対象数 / count when hidden counts as occupant */
function showOptionsDialog(countIgnore, countAll) {
    var dialog = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dialog.orientation = 'column';
    dialog.alignChildren = 'fill';
    dialog.margins = 16;
    dialog.spacing = 12;

    /* 削除対象数の表示 / Count summary */
    var countText = dialog.add('statictext', undefined, labelText('removalTarget') + countIgnore);

    /* オプション: 非表示要素を無視 / Option: ignore hidden elements */
    var hiddenCheckbox = dialog.add('checkbox', undefined, L('ignoreHidden'));
    hiddenCheckbox.value = true;

    /* ボタン行 (Mac 規約: Cancel → OK) / Button row (Mac convention) */
    var buttonGroup = dialog.add('group');
    buttonGroup.alignment = 'right';
    buttonGroup.add('button', undefined, L('cancel'), { name: 'cancel' });
    var okButton = buttonGroup.add('button', undefined, 'OK', { name: 'ok' });
    okButton.enabled = (countIgnore > 0);

    /* チェックボックス連動で件数と OK 活性を更新 / Sync count and OK state on toggle */
    hiddenCheckbox.onClick = function () {
        var c = hiddenCheckbox.value ? countIgnore : countAll;
        countText.text = labelText('removalTarget') + c;
        okButton.enabled = (c > 0);
    };

    if (dialog.show() !== 1) return null;
    return { ignoreHidden: hiddenCheckbox.value };
}

// =========================================
// メイン / Main
// =========================================

(function main() {
    if (app.documents.length === 0) {
        alert(L('noDocument'));
        return;
    }

    var doc = app.activeDocument;
    if (doc.artboards.length <= 1) {
        alert(L('onlyOneArtboard'));
        return;
    }

    var indicesIgnore = findEmptyArtboardIndices(doc, true);
    var indicesAll = findEmptyArtboardIndices(doc, false);

    var options = showOptionsDialog(indicesIgnore.length, indicesAll.length);
    if (!options) return;

    var emptyIndices = options.ignoreHidden ? indicesIgnore : indicesAll;
    removeArtboardsByIndices(doc, emptyIndices);
})();
