#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

概要

選択した複数オブジェクトを、それぞれ個別のシンボルとして登録し、
元オブジェクトをそのシンボルインスタンスで置き換える Illustrator 用 JSX スクリプト。

- 実行時にダイアログで「指定方法（個別 / 一括）」「シンボル名（接頭辞・連番桁数・テキスト流用）」「基準点」「リンク画像の扱い」を設定する
- ［指定方法］
  - ［一括］：接頭辞・連番・テキスト流用・基準点をスクリプト側で適用して自動的にシンボル化する
  - ［個別］：オブジェクトごとに Illustrator 標準の「新規シンボル…」ダイアログ（Adobe New Symbol Shortcut）を開き、名前と基準点を都度指定する
- 個別モードでは対象オブジェクト以外の選択を解除し、ビューを対象へスクロール＋ズーム調整（約 50% 占有）してからダイアログを開く
- 個別モードの開始時のビュー状態（ズーム・中心点）を退避し、終了時に復元する
- 個別モード中は、ダイアログ上で「接頭辞・連番・テキスト流用・基準点」のコントロールをディム表示する
- 一括モード時のシンボル名の優先順位は「テキスト内容（流用 ON 時）→ レイヤー名（既定名は除外）→ item.note → 接頭辞＋連番」
- ［テキスト流用］が ON のときは TextFrame の内容（または GroupItem 内の最初の TextFrame）をシンボル名に使う
- レイヤー名は "Layer N" / "レイヤー N" のような既定名はスキップし、明示的にリネームされた場合のみ採用する
- メモは Attributes パネルの note 欄を参照する（空白だけは無効扱い）
- どれも取れない場合は接頭辞＋指定桁数のゼロ埋め連番（0 / 00 / 000）で命名する
- 既存シンボルと同名になる場合は末尾に "_2", "_3"... を付けて重複を回避する
- リンク画像（PlacedItem）は［強制埋め込み］でシンボル化前に embed、［無視］でスキップする
- すでに SymbolItem の選択はそのまま残し、新規登録の対象から除外する
- 一括モードでは元オブジェクトの geometricBounds を記録し、生成したシンボルインスタンスを同位置・近い重なり順に配置する
- 処理後は新規シンボルインスタンスとスルーした既存 SymbolItem / 無視した PlacedItem を選択状態にする
- 完了時に「新規作成数 / 既存シンボル数 / 無視リンク画像数 / 失敗数」をまとめて通知する

Overview

Illustrator JSX script that converts each selected object into its own symbol
and replaces the original with an instance of that newly created symbol.

- Shows a dialog at launch to configure the input method (Individual / Batch), symbol-name settings, registration point, and linked-image policy
- Method:
  - "Batch": applies prefix / sequence / text reuse / registration point automatically and symbolizes each item without further interaction
  - "Individual": opens Illustrator's native "New Symbol..." dialog (Adobe New Symbol Shortcut) for each item so name and registration point can be entered per item
- In Individual mode, deselects other items and scrolls/zooms the view onto the target (~50% of the view) before opening the dialog
- Saves the view state (zoom and center) at the start of Individual mode and restores it on exit
- During Individual mode, dims prefix / sequence / text-reuse / registration-point controls in the dialog
- In Batch mode, symbol-name priority: text contents (when "Use text" is on) → parent layer name (excluding default "Layer N" / "レイヤー N") → item note → prefix + zero-padded sequence
- When "Use text contents" is on, uses the TextFrame contents (or first TextFrame inside a GroupItem)
- Layer names matching the default "Layer N" / "レイヤー N" pattern are treated as unset
- Item notes come from the Attributes panel's note field (whitespace-only values are treated as unset)
- Falls back to the prefix plus a zero-padded sequence number (0 / 00 / 000) when none of the above is available
- Appends "_2", "_3"... to avoid collisions with existing symbol names
- Linked images (PlacedItem) are embedded before symbolizing under "Force embed" or skipped under "Ignore"
- Leaves existing SymbolItems in the selection untouched and excludes them from registration
- In Batch mode, records the original geometricBounds and places the new instance at the same position and near the same z-order
- Leaves the newly created instances and any skipped SymbolItems / ignored PlacedItems selected after the run
- Reports created / existing-symbol / ignored-linked-image / failed counts in a summary alert

*/

// =========================================
// バージョンとローカライズ / Version and localization
// =========================================

var SCRIPT_VERSION = "v1.0.0";

/* 現在の UI 言語 / Current UI language */
var currentLang = ($.locale.indexOf('ja') === 0) ? 'ja' : 'en';

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: '個別にシンボル化', en: 'Symbolize Individually' },
    panelMode: { ja: '指定方法', en: 'Method' },
    radioModeIndividual: { ja: '個別', en: 'Individual' },
    radioModeBatch: { ja: '一括', en: 'Batch' },
    panelSymbolName: { ja: 'シンボル名', en: 'Symbol name' },
    labelPrefix: { ja: '接頭辞', en: 'Prefix' },
    labelSequence: { ja: '連番', en: 'Sequence' },
    checkboxUseText: { ja: 'テキストの文字列を参照', en: 'Use text contents' },
    helpUseText: {
        ja: 'TextFrame（またはグループ内の最初の TextFrame）の内容をシンボル名に流用します。',
        en: 'Reuse the TextFrame contents (or the first TextFrame in a group) as the symbol name.'
    },
    helpPrefix: {
        ja: 'テキスト以外、または流用 OFF のときに使う接頭辞。',
        en: 'Prefix used for non-text objects or when text reuse is off.'
    },
    helpSequence: { ja: '連番のゼロ埋め桁数。', en: 'Zero-pad width for the sequence number.' },
    panelReferencePoint: { ja: '基準点', en: 'Registration point' },
    helpReferencePoint: { ja: 'シンボル登録時の基準点。', en: 'Registration point used when registering each symbol.' },
    panelLinkedImage: { ja: 'リンク画像', en: 'Linked image' },
    radioIgnore: { ja: '無視', en: 'Ignore' },
    radioForceEmbed: { ja: '強制埋め込み', en: 'Force embed' },
    helpLinkedImage: {
        ja: 'リンク画像（PlacedItem）の扱い。［無視］は選択に残しシンボル化対象から外します。［強制埋め込み］はシンボル化前に embed します。',
        en: 'Policy for linked images (PlacedItem). "Ignore" leaves them selected but skips them; "Force embed" embeds before symbolizing.'
    },
    buttonCancel: { ja: 'キャンセル', en: 'Cancel' },

    errorNoDocument: {
        ja: 'ドキュメントが開かれていません。',
        en: 'No document is open.'
    },

    errorNoSelection: {
        ja: 'シンボル化したいオブジェクトを選択してください。',
        en: 'Select the objects you want to symbolize.'
    },

    resultCompleted: {
        ja: '完了しました。',
        en: 'Completed.'
    },

    resultCreated: {
        ja: '新規作成したシンボル数',
        en: 'Created symbols'
    },

    resultExisting: {
        ja: 'スルーした既存シンボル数',
        en: 'Skipped existing symbols'
    },

    resultIgnoredLinked: {
        ja: '無視したリンク画像数',
        en: 'Ignored linked images'
    },

    resultFailed: {
        ja: '失敗',
        en: 'Failed'
    },

    resultFailureDetails: {
        ja: '失敗の詳細',
        en: 'Failure details'
    }
};

/* ラベルから現在の UI 言語の文字列を取り出す / Resolve a label for the current UI language */
function L(key) { return LABELS[key][currentLang]; }

// =========================================
// 定数とプリセット / Constants and presets
// =========================================

/* 既定の接頭辞 / Default prefix string */
var DEFAULT_PREFIX = "Symbol_";

/* 連番の桁数候補（"0" / "00" / "000" に対応） / Sequence-padding presets */
var SEQUENCE_PADDINGS = [1, 2, 3];
var DEFAULT_SEQUENCE_INDEX = 2; /* 000 */

/* 3×3 基準点と SymbolRegistrationPoint の対応表（行優先：上 → 中 → 下、列：左 → 中 → 右）
   Row-major 3×3 reference points mapped to SymbolRegistrationPoint values */
var REFERENCE_POINTS = [
    SymbolRegistrationPoint.SYMBOLTOPLEFTPOINT,
    SymbolRegistrationPoint.SYMBOLTOPMIDDLEPOINT,
    SymbolRegistrationPoint.SYMBOLTOPRIGHTPOINT,
    SymbolRegistrationPoint.SYMBOLMIDDLELEFTPOINT,
    SymbolRegistrationPoint.SYMBOLCENTERPOINT,
    SymbolRegistrationPoint.SYMBOLMIDDLERIGHTPOINT,
    SymbolRegistrationPoint.SYMBOLBOTTOMLEFTPOINT,
    SymbolRegistrationPoint.SYMBOLBOTTOMMIDDLEPOINT,
    SymbolRegistrationPoint.SYMBOLBOTTOMRIGHTPOINT
];
var DEFAULT_REFERENCE_POINT_INDEX = 4; /* CENTER */

/* リンク画像の扱い / Linked-image policy values */
var LINKED_IMAGE_POLICY = { IGNORE: 'ignore', EMBED: 'embed' };
var DEFAULT_LINKED_IMAGE_POLICY_INDEX = 1; /* EMBED */

/* シンボル化モード（個別 / 一括）。ロジックは追って実装 / Symbolization mode (Individual / Batch); logic to be wired up later */
var SYMBOLIZE_MODE = { INDIVIDUAL: 'individual', BATCH: 'batch' };
var DEFAULT_SYMBOLIZE_MODE_INDEX = 1; /* BATCH */

// =========================================
// メイン処理 / Main entry
// =========================================

(function () {

    if (app.documents.length === 0) {
        alert(L('errorNoDocument'));
        return;
    }

    var doc = app.activeDocument;

    if (!doc.selection || doc.selection.length === 0) {
        alert(L('errorNoSelection'));
        return;
    }

    var originalSelection = snapshotSelection(doc);

    var settings = showSettingsDialog();
    if (!settings) return;

    doc.selection = null;

    var stats = settings.mode === SYMBOLIZE_MODE.INDIVIDUAL
        ? processSelectionIndividual(doc, originalSelection, settings)
        : processSelection(doc, originalSelection, settings);

    applySelection(doc, stats.finalSelection);
    showResultSummary(stats);

})();

/* 現在の選択を配列としてコピーする / Copy the current selection into a plain array */
function snapshotSelection(doc) {
    var snapshot = [];
    for (var i = 0; i < doc.selection.length; i++) {
        snapshot.push(doc.selection[i]);
    }
    return snapshot;
}

/* 選択配列を順に処理してシンボル化結果を集計する
   Iterate the selection, symbolize each item, and return the aggregated result */
function processSelection(doc, items, settings) {
    var finalSelection = [];
    var createdCount = 0;
    var existingSymbolCount = 0;
    var ignoredLinkedImageCount = 0;
    var failedCount = 0;
    var failureMessages = [];

    for (var j = 0; j < items.length; j++) {
        var originalItem = items[j];

        /* 既存シンボルインスタンスはそのまま残す / Leave existing symbol instances untouched */
        if (originalItem.typename === "SymbolItem") {
            existingSymbolCount++;
            finalSelection.push(originalItem);
            continue;
        }

        /* リンク画像で［無視］が指定されている場合はスキップ / Skip linked images when policy is "ignore" */
        if (originalItem.typename === "PlacedItem" && settings.linkedImagePolicy === LINKED_IMAGE_POLICY.IGNORE) {
            ignoredLinkedImageCount++;
            finalSelection.push(originalItem);
            continue;
        }

        try {
            var newInstance = symbolizeOneItem(doc, originalItem, j, settings);
            finalSelection.push(newInstance);
            createdCount++;
        } catch (err) {
            failedCount++;
            failureMessages.push(formatFailureMessage(originalItem, err));
        }
    }

    return {
        finalSelection: finalSelection,
        createdCount: createdCount,
        existingSymbolCount: existingSymbolCount,
        ignoredLinkedImageCount: ignoredLinkedImageCount,
        failedCount: failedCount,
        failureMessages: failureMessages
    };
}

/* 個別モード：選択中の各アイテムをひとつずつ選択し直し、Illustrator 標準の「新規シンボル…」ダイアログを呼び出す
   Individual mode: for each selected item, re-select it alone and invoke Illustrator's native "New Symbol..." dialog */
function processSelectionIndividual(doc, items, settings) {
    var finalSelection = [];
    var createdCount = 0;
    var existingSymbolCount = 0;
    var ignoredLinkedImageCount = 0;
    var failedCount = 0;
    var failureMessages = [];

    /* 開始時のビューを退避し、終了時に必ず元に戻す / Save the current view so it can be restored at the end */
    var savedView = captureViewState(doc);

    try {
        for (var j = 0; j < items.length; j++) {
            var originalItem = items[j];

            /* 既存シンボルインスタンスはそのまま残す / Leave existing symbol instances untouched */
            if (originalItem.typename === "SymbolItem") {
                existingSymbolCount++;
                finalSelection.push(originalItem);
                continue;
            }

            /* リンク画像で［無視］が指定されている場合はスキップ / Skip linked images when policy is "ignore" */
            if (originalItem.typename === "PlacedItem" && settings.linkedImagePolicy === LINKED_IMAGE_POLICY.IGNORE) {
                ignoredLinkedImageCount++;
                finalSelection.push(originalItem);
                continue;
            }

            try {
                /* リンク画像（強制埋め込み時）はここで埋め込み、置換後のアイテムを以後の処理対象にする
                   Embed linked images (under "force embed") and continue with the replacement item */
                var workingItem = ensureEmbedded(doc, originalItem);

                /* 対象だけを選択し、ビューを対象へフォーカスしてから redraw でハイライトを描画
                   Isolate the selection, focus the view on the item, and force a redraw so highlight appears before the modal */
                doc.selection = null;
                workingItem.selected = true;
                focusViewOnItem(doc, workingItem);
                app.redraw();

                var beforeSymbolCount = doc.symbols.length;

                /* 標準の「新規シンボル…」ダイアログを開く。ユーザーが名前と基準点を指定する
                   Open the native "New Symbol..." dialog so the user can enter name and registration point */
                app.executeMenuCommand('Adobe New Symbol Shortcut');

                if (doc.symbols.length > beforeSymbolCount) {
                    createdCount++;
                }

                /* メニュー実行後の選択をそのまま結果へ取り込む（キャンセル時は元アイテムが残る）
                   Forward whatever is selected after the menu command (cancel leaves the original selected) */
                if (doc.selection && doc.selection.length > 0) {
                    for (var k = 0; k < doc.selection.length; k++) {
                        finalSelection.push(doc.selection[k]);
                    }
                }
            } catch (err) {
                failedCount++;
                failureMessages.push(formatFailureMessage(originalItem, err));
            }
        }
    } finally {
        restoreViewState(doc, savedView);
    }

    return {
        finalSelection: finalSelection,
        createdCount: createdCount,
        existingSymbolCount: existingSymbolCount,
        ignoredLinkedImageCount: ignoredLinkedImageCount,
        failedCount: failedCount,
        failureMessages: failureMessages
    };
}

/* 失敗したアイテムの種類・名前・エラー内容を短く整形する
   Format the failed item's type, name, and error message for diagnostics */
function formatFailureMessage(item, err) {
    var itemType = 'Unknown';
    var itemName = '';
    var errorMessage = '';

    try {
        if (item && item.typename) itemType = item.typename;
    } catch (typeError) { }

    try {
        if (item && item.name) itemName = item.name;
    } catch (nameError) { }

    try {
        if (err && err.message) errorMessage = err.message;
        else errorMessage = String(err);
    } catch (messageError) {
        errorMessage = 'Unknown error';
    }

    if (itemName !== '') {
        return itemType + ' "' + itemName + '": ' + errorMessage;
    }
    return itemType + ': ' + errorMessage;
}

/* 指定アイテム群を選択状態にする（既存選択はクリア）
   Replace the current selection with the given items */
function applySelection(doc, items) {
    doc.selection = null;
    for (var k = 0; k < items.length; k++) {
        try {
            items[k].selected = true;
        } catch (e) { }
    }
}

/* 処理結果のサマリーを通知する / Show a summary alert with the processing counts */
function showResultSummary(stats) {
    var message =
        L('resultCompleted') + "\n\n" +
        L('resultCreated') + ": " + stats.createdCount + "\n" +
        L('resultExisting') + ": " + stats.existingSymbolCount + "\n" +
        L('resultIgnoredLinked') + ": " + stats.ignoredLinkedImageCount + "\n" +
        L('resultFailed') + ": " + stats.failedCount;

    if (stats.failureMessages && stats.failureMessages.length > 0) {
        message += "\n\n" + L('resultFailureDetails') + "\n" + stats.failureMessages.join("\n");
    }

    alert(message);
}

// =========================================
// シンボル化処理 / Symbolize one item
// =========================================

/* 1 オブジェクトを新規シンボル化し、元オブジェクトを生成したシンボルインスタンスで置き換える
   Convert one item into a new symbol and replace the original with the resulting instance */
function symbolizeOneItem(doc, originalItem, index, settings) {
    /* 埋め込みで bounds が変わる可能性があるため、元オブジェクトの位置は先に記録する
       Capture the original bounds before embedding because embed() may alter the replacement bounds */
    var originalBounds = originalItem.geometricBounds;

    /* リンク画像（強制埋め込み時）はここで埋め込み、置換後のアイテムを以後の処理対象にする
       Embed linked images (under "force embed") and continue with the replacement item */
    var workingItem = ensureEmbedded(doc, originalItem);

    var symbolName = resolveSymbolName(doc, workingItem, index, settings);

    var newSymbol = createSymbolFromItem(doc, workingItem, symbolName, settings.referencePoint);
    var newInstance = doc.symbolItems.add(newSymbol);

    /* 元オブジェクト直前へ移動して重なり順を近づける / Move just before the original to keep the z-order close */
    try {
        newInstance.move(workingItem, ElementPlacement.PLACEBEFORE);
    } catch (e) { }

    alignToBounds(newInstance, originalBounds);
    workingItem.remove();

    return newInstance;
}

/* PlacedItem（リンク画像）であれば埋め込み、置き換え後のアイテムを返す。それ以外はそのまま返す
   Embed a PlacedItem in place and return the replacement; pass through other item types */
function ensureEmbedded(doc, item) {
    if (item.typename !== "PlacedItem") return item;

    /* embed() 前に同じ親内の pageItems を記録し、embed() 後に新規追加されたアイテムを探す
       Record the parent's pageItems before embed(), then find the newly added item afterward */
    var parent = item.parent;
    var beforeItems = snapshotParentPageItems(parent);

    item.embed();

    var embeddedItem = findNewPageItem(parent, beforeItems);
    if (embeddedItem) return embeddedItem;

    /* Illustrator の状態によっては埋め込み後の選択に置換アイテムが入るため、最後の保険として使う
       As a last resort, use the active selection because Illustrator may select the embedded replacement */
    if (doc.selection && doc.selection.length > 0) {
        return doc.selection[0];
    }

    throw new Error("Embedded replacement item could not be located.");
}

/* 親コンテナ直下の pageItems を配列として記録する
   Snapshot direct pageItems under the parent container */
function snapshotParentPageItems(parent) {
    var items = [];
    try {
        for (var i = 0; i < parent.pageItems.length; i++) {
            items.push(parent.pageItems[i]);
        }
    } catch (e) { }
    return items;
}

/* embed() 前の一覧に存在しない pageItem を探す
   Find a pageItem that did not exist in the pre-embed snapshot */
function findNewPageItem(parent, beforeItems) {
    try {
        for (var i = 0; i < parent.pageItems.length; i++) {
            if (!containsPageItem(beforeItems, parent.pageItems[i])) {
                return parent.pageItems[i];
            }
        }
    } catch (e) { }
    return null;
}

/* pageItem 配列に同一参照が含まれるかを確認する
   Return true when the array contains the same pageItem reference */
function containsPageItem(items, targetItem) {
    for (var i = 0; i < items.length; i++) {
        if (items[i] === targetItem) return true;
    }
    return false;
}

/* シンボル名を決定する。優先順位は次の通り
   1. テキスト内容（テキスト流用 ON のとき）
   2. レイヤー名（既定の "Layer N" / "レイヤー N" は除外）
   3. メモ（item.note）
   4. 接頭辞＋連番
   Resolve a unique symbol name following this priority chain:
   1. Text contents (when the "use text" checkbox is on)
   2. Layer name (default names like "Layer N" / "レイヤー N" are skipped)
   3. Item note
   4. Prefix + zero-padded sequence number */
function resolveSymbolName(doc, item, index, settings) {
    var baseName = "";

    if (settings.useTextAsName) {
        baseName = sanitizeSymbolName(getTextNameFromItem(item));
    }
    if (baseName === "") {
        baseName = sanitizeSymbolName(getLayerNameFromItem(item));
    }
    if (baseName === "") {
        baseName = sanitizeSymbolName(getNoteFromItem(item));
    }
    if (baseName === "") {
        baseName = settings.defaultPrefix + zeroPadding(index + 1, settings.paddingLength);
    }

    return getUniqueSymbolName(doc, baseName);
}

/* 元オブジェクトを複製してシンボル登録し、複製は可能なら破棄する
   Duplicate the source item, register it as a symbol, then remove the duplicate when possible */
function createSymbolFromItem(doc, sourceItem, symbolName, referencePoint) {
    var duplicated = sourceItem.duplicate();
    var newSymbol = doc.symbols.add(duplicated, referencePoint);
    newSymbol.name = symbolName;

    try {
        duplicated.remove();
    } catch (removeError) { }

    return newSymbol;
}

// =========================================
// 配置ユーティリティ / Placement helpers
// =========================================

/* オブジェクトを指定 bounds の左上に揃うよう平行移動する
   Translate an item so its top-left aligns with the target bounds */
function alignToBounds(item, targetBounds) {
    var currentBounds = item.geometricBounds;
    var dx = targetBounds[0] - currentBounds[0];
    var dy = targetBounds[1] - currentBounds[1];
    item.translate(dx, dy);
}

/* 個別モード用：現在のビュー状態（ズーム倍率と中心点）を退避する
   Capture current view state (zoom and center) for restoration after individual mode */
function captureViewState(doc) {
    try {
        var view = doc.views[0];
        return {
            zoom: view.zoom,
            centerPoint: [view.centerPoint[0], view.centerPoint[1]]
        };
    } catch (e) {
        return null;
    }
}

/* 退避したビュー状態を復元する / Restore the saved view state */
function restoreViewState(doc, state) {
    if (!state) return;
    try {
        var view = doc.views[0];
        view.zoom = state.zoom;
        view.centerPoint = state.centerPoint;
    } catch (e) { }
}

/* 指定アイテムをビューの中央にスクロールし、必要に応じてズームを調整する
   Scroll the view to center on the item, adjusting zoom so it occupies ~50% of the view */
function focusViewOnItem(doc, item) {
    try {
        var view = doc.views[0];
        var b = item.geometricBounds; /* [left, top, right, bottom] / top > bottom */
        var itemW = b[2] - b[0];
        var itemH = b[1] - b[3];
        if (itemW <= 0 || itemH <= 0) return;

        var centerX = (b[0] + b[2]) / 2;
        var centerY = (b[1] + b[3]) / 2;

        var vb = view.bounds;
        var viewW = vb[2] - vb[0];
        var viewH = vb[1] - vb[3];
        var currentZoom = view.zoom;

        var targetZoom = Math.min(
            0.5 * viewW * currentZoom / itemW,
            0.5 * viewH * currentZoom / itemH
        );
        if (targetZoom < 0.05) targetZoom = 0.05;
        if (targetZoom > 64) targetZoom = 64;

        view.zoom = targetZoom;
        view.centerPoint = [centerX, centerY];
    } catch (e) { }
}

// =========================================
// 名前候補の抽出 / Name candidate extraction
// =========================================

/* TextFrame そのもの、またはグループ内に含まれる最初の TextFrame の内容を返す
   Return the contents of a TextFrame, or the first TextFrame found inside a GroupItem */
function getTextNameFromItem(item) {
    if (item.typename === "TextFrame") {
        return item.contents;
    }
    if (item.typename === "GroupItem") {
        return getTextNameFromGroup(item);
    }
    return null;
}

/* グループ内を再帰的に走査し、最初の TextFrame の内容を返す
   Recursively walk a group and return the first TextFrame's contents */
function getTextNameFromGroup(groupItem) {
    for (var i = 0; i < groupItem.pageItems.length; i++) {
        var child = groupItem.pageItems[i];
        if (child.typename === "TextFrame") {
            return child.contents;
        }
        if (child.typename === "GroupItem") {
            var nestedName = getTextNameFromGroup(child);
            if (nestedName) return nestedName;
        }
    }
    return null;
}

/* アイテムが属するレイヤーの名前を返す。既定名（"Layer N" / "レイヤー N"）は無視する
   Return the parent layer's name; default names like "Layer N" / "レイヤー N" are treated as unset */
function getLayerNameFromItem(item) {
    try {
        var layer = item.layer;
        if (!layer) return null;
        var name = layer.name;
        if (!name) return null;
        if (/^(Layer|レイヤー)\s*\d+$/.test(name)) return null;
        return name;
    } catch (e) {
        return null;
    }
}

/* アイテムに設定されたメモ（item.note）を返す。空文字は無効扱い
   Return the item.note value; empty strings are treated as unset */
function getNoteFromItem(item) {
    try {
        var note = item.note;
        if (!note || trimString(String(note)) === "") return null;
        return note;
    } catch (e) {
        return null;
    }
}


// =========================================
// 文字列・名前ユーティリティ / String and name helpers
// =========================================

/* シンボル名向けに整形する（改行・タブを空白化、前後空白除去、80 文字に切り詰め）
   Normalize a string for use as a symbol name */
function sanitizeSymbolName(name) {
    if (name === null || name === undefined) return "";
    name = trimString(String(name).replace(/[\r\n\t]+/g, " "));
    if (name.length > 80) name = name.substring(0, 80);
    return name;
}

/* 数値を指定桁数までゼロ埋めする / Left-pad a number with zeros to the given length */
function zeroPadding(num, length) {
    var str = String(num);
    while (str.length < length) str = "0" + str;
    return str;
}

/* 既存シンボル名と重複しない名前を返す（必要に応じて "_2", "_3" ... を付与）
   Build a symbol name unique within the document, suffixing "_2", "_3" ... as needed */
function getUniqueSymbolName(doc, baseName) {
    var name = baseName;
    var count = 2;
    while (symbolExists(doc, name)) {
        name = baseName + "_" + count;
        count++;
    }
    return name;
}

/* 同名のシンボルが既に存在するかを確認 / Return true if a symbol with this name already exists */
function symbolExists(doc, name) {
    try {
        doc.symbols.getByName(name);
        return true;
    } catch (e) {
        return false;
    }
}

/* String.trim 互換（前後の空白を除去） / String.trim equivalent for ES3 */
function trimString(str) {
    return str.replace(/^\s+|\s+$/g, "");
}

// =========================================
// ダイアログ / Dialog
// =========================================

/* ダイアログを表示し、設定オブジェクトを返す。Cancel 時は null を返す
   Show the settings dialog and return a settings object; null on cancel */
function showSettingsDialog() {
    var dialog = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dialog.alignChildren = ['fill', 'top'];
    dialog.margins = 16;
    dialog.spacing = 12;

    var modeButtons = buildModePanel(dialog);
    var symbolNameControls = buildSymbolNamePanel(dialog);

    /* 基準点とリンク画像を横並び 2 カラムに配置 / Lay out registration point and linked image side-by-side */
    var columnsGroup = dialog.add('group');
    columnsGroup.orientation = 'row';
    columnsGroup.alignChildren = ['fill', 'fill'];
    columnsGroup.spacing = 12;

    var referencePointButtons = buildReferencePointPanel(columnsGroup);
    var linkedImageButtons = buildLinkedImagePanel(columnsGroup);

    /* モードに応じて 一括専用コントロール（接頭辞・連番・テキスト流用・基準点）の有効状態を切り替える
       Toggle batch-only controls (prefix / sequence / use-text / registration point) based on the selected mode */
    function refreshModeUI() {
        var batchEnabled = (modeButtons.selectedIndex !== 0);
        symbolNameControls.prefixRow.enabled = batchEnabled;
        symbolNameControls.sequenceRow.enabled = batchEnabled;
        symbolNameControls.useTextCheckbox.enabled = batchEnabled;
        referencePointButtons.panel.enabled = batchEnabled;
    }

    for (var mi = 0; mi < modeButtons.length; mi++) {
        (function (idx) {
            modeButtons[idx].onClick = function () {
                selectExclusive(modeButtons, idx);
                refreshModeUI();
            };
        })(mi);
    }
    refreshModeUI();

    /* OK / Cancel（Mac 規約：Cancel → OK）/ OK / Cancel (Mac order: Cancel → OK) */
    var buttonGroup = dialog.add('group');
    buttonGroup.alignment = 'right';
    buttonGroup.add('button', undefined, L('buttonCancel'), { name: 'cancel' });
    var okButton = buttonGroup.add('button', undefined, 'OK', { name: 'ok' });
    dialog.defaultElement = okButton;

    if (dialog.show() !== 1) return null;

    var prefixValue = trimString(symbolNameControls.prefixInput.text);
    if (prefixValue === "") prefixValue = DEFAULT_PREFIX;

    return {
        mode: modeButtons.selectedIndex === 0
            ? SYMBOLIZE_MODE.INDIVIDUAL
            : SYMBOLIZE_MODE.BATCH,
        defaultPrefix: prefixValue,
        paddingLength: SEQUENCE_PADDINGS[symbolNameControls.sequenceButtons.selectedIndex],
        useTextAsName: symbolNameControls.useTextCheckbox.value,
        referencePoint: REFERENCE_POINTS[referencePointButtons.selectedIndex],
        linkedImagePolicy: linkedImageButtons.selectedIndex === 0
            ? LINKED_IMAGE_POLICY.IGNORE
            : LINKED_IMAGE_POLICY.EMBED
    };
}

/* モードパネル（個別 / 一括）を構築。ロジックは未接続 / Build the mode panel (Individual / Batch); not wired to logic yet */
function buildModePanel(parent) {
    var panel = parent.add('panel', undefined, L('panelMode'));
    panel.orientation = 'row';
    panel.alignChildren = ['left', 'center'];
    panel.margins = [12, 16, 12, 12];
    panel.spacing = 12;

    var radioButtons = [];
    radioButtons.push(panel.add('radiobutton', undefined, L('radioModeIndividual')));
    radioButtons.push(panel.add('radiobutton', undefined, L('radioModeBatch')));
    selectExclusive(radioButtons, DEFAULT_SYMBOLIZE_MODE_INDEX);
    bindExclusiveRadios(radioButtons);
    return radioButtons;
}

/* シンボル名パネル（接頭辞・連番桁数・テキスト流用チェック）を構築
   Build the symbol-name panel (prefix, sequence padding, text-reuse checkbox) */
function buildSymbolNamePanel(dialog) {
    var panel = dialog.add('panel', undefined, L('panelSymbolName'));
    panel.alignChildren = ['fill', 'top'];
    panel.margins = [12, 16, 12, 12];
    panel.spacing = 8;

    /* 接頭辞 / Prefix */
    var prefixRow = panel.add('group');
    prefixRow.orientation = 'row';
    prefixRow.alignChildren = ['left', 'center'];
    prefixRow.add('statictext', undefined, L('labelPrefix'));
    var prefixInput = prefixRow.add('edittext', undefined, DEFAULT_PREFIX);
    prefixInput.characters = 16;
    prefixInput.preferredSize.width = 180;
    prefixInput.helpTip = L('helpPrefix');
    prefixInput.active = true;

    /* 連番桁数 / Sequence padding */
    var sequenceRow = panel.add('group');
    sequenceRow.orientation = 'row';
    sequenceRow.alignChildren = ['left', 'center'];
    sequenceRow.spacing = 8;
    sequenceRow.add('statictext', undefined, L('labelSequence'));
    var sequenceButtons = createPaddingRadios(sequenceRow, DEFAULT_SEQUENCE_INDEX);

    /* テキスト流用 / Use text as name */
    var useTextCheckbox = panel.add('checkbox', undefined, L('checkboxUseText'));
    useTextCheckbox.value = true;
    useTextCheckbox.helpTip = L('helpUseText');

    return {
        prefixRow: prefixRow,
        prefixInput: prefixInput,
        sequenceRow: sequenceRow,
        sequenceButtons: sequenceButtons,
        useTextCheckbox: useTextCheckbox
    };
}

/* 連番桁数のラジオボタンを生成（"0" / "00" / "000"）
   Build the sequence-padding radio buttons ("0" / "00" / "000") */
function createPaddingRadios(parentGroup, defaultIndex) {
    var radioButtons = [];
    for (var i = 0; i < SEQUENCE_PADDINGS.length; i++) {
        var paddingLabel = zeroPadding(0, SEQUENCE_PADDINGS[i]);
        var button = parentGroup.add('radiobutton', undefined, paddingLabel);
        button.helpTip = L('helpSequence');
        radioButtons.push(button);
    }
    selectExclusive(radioButtons, defaultIndex);
    bindExclusiveRadios(radioButtons);
    return radioButtons;
}

/* 基準点パネル（3×3 グリッド）を構築 / Build the registration-point panel (3×3 grid) */
function buildReferencePointPanel(parent) {
    var panel = parent.add('panel', undefined, L('panelReferencePoint'));
    panel.alignChildren = 'center';
    panel.margins = [12, 16, 12, 12];
    panel.spacing = 6;
    panel.helpTip = L('helpReferencePoint');
    var grid = createReferencePointGrid(panel, DEFAULT_REFERENCE_POINT_INDEX);
    grid.panel = panel;
    return grid;
}

/* リンク画像パネル（無視 / 強制埋め込み）を構築 / Build the linked-image panel (Ignore / Force embed) */
function buildLinkedImagePanel(parent) {
    var panel = parent.add('panel', undefined, L('panelLinkedImage'));
    panel.alignChildren = ['left', 'top'];
    panel.margins = [12, 16, 12, 12];
    panel.spacing = 6;
    panel.helpTip = L('helpLinkedImage');

    var radioButtons = [];
    radioButtons.push(panel.add('radiobutton', undefined, L('radioIgnore')));
    radioButtons.push(panel.add('radiobutton', undefined, L('radioForceEmbed')));
    selectExclusive(radioButtons, DEFAULT_LINKED_IMAGE_POLICY_INDEX);
    bindExclusiveRadios(radioButtons);
    return radioButtons;
}

/* 3×3 基準点ラジオボタンを生成（手動排他、選択値は radioButtons.selectedIndex に保持）
   Build a 3x3 radio grid with manual mutual exclusion; the selected index is exposed as radioButtons.selectedIndex */
function createReferencePointGrid(parentPanel, defaultIndex) {
    var radioButtons = [];
    for (var row = 0; row < 3; row++) {
        var rowGroup = parentPanel.add('group');
        rowGroup.orientation = 'row';
        rowGroup.spacing = 4;
        for (var col = 0; col < 3; col++) {
            var button = rowGroup.add('radiobutton', undefined, '');
            radioButtons.push(button);
        }
    }
    selectExclusive(radioButtons, defaultIndex);
    bindExclusiveRadios(radioButtons);
    return radioButtons;
}

/* ラジオボタン群を排他的に選択し、selectedIndex に選択値を保持する
   Set radios to be mutually exclusive and expose the selection as selectedIndex */
function selectExclusive(radioButtons, selectedIndex) {
    for (var i = 0; i < radioButtons.length; i++) {
        radioButtons[i].value = (i === selectedIndex);
    }
    radioButtons.selectedIndex = selectedIndex;
}

/* ラジオボタン群のクリック時に排他選択を反映するイベントを設定する
   Wire each radio button to refresh the group's exclusive selection on click */
function bindExclusiveRadios(radioButtons) {
    for (var i = 0; i < radioButtons.length; i++) {
        (function (buttonIndex) {
            radioButtons[buttonIndex].onClick = function () {
                selectExclusive(radioButtons, buttonIndex);
            };
        })(i);
    }
}