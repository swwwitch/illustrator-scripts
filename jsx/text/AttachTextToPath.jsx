/*
### スクリプト名：
ポイント文字をパス上文字に変換

### 更新日：
20260302

### 概要：
選択中の「ポイント文字」を、同時に選択しているパス（PathItem）上の文字に変換します。
- パスが一緒に選択されている場合：そのパスを複製して使用（複数テキストにも対応）
- パスが選択されていない場合：処理を実行せず終了します（パス必須）
*/

(function () {

    // Constant
    const SCRIPT_TITLE = 'ポイント文字とパスをパス上文字に変換';
    const SCRIPT_VERSION = '0.';

    // Get items
    if (app.documents.length === 0) {
        alert('ドキュメントが開かれていません');
        return false;
    }
    var doc = app.activeDocument;
    var sel = doc.selection;

    var targetItems = getTargetTextItems(sel);
    var selectedPaths = getSelectedPathItems(sel);

    // Validation
    if (targetItems.length === 0) {
        alert('対象のポイント文字が見つかりません');
        return false;
    }

    if (!selectedPaths || selectedPaths.length === 0) {
        alert('パスを一緒に選択してください');
        return false;
    }

    mainProcess();

    // Main process
    function mainProcess() {

        for (var j = 0; j < targetItems.length; j++) {
            // Get original text frame item & current layer
            var originalText = targetItems[j];
            var currentlayer = originalText.layer;

            // Resolve base path + duplicate path for this text
            var basePath = resolveBasePathForText(selectedPaths, targetItems.length, j);
            if (!basePath) {
                // Should not happen because of Validation, but keep defensive
                alert('パスを一緒に選択してください');
                return;
            }

            makeOriginalPathInvisible(basePath);
            var textPath = duplicatePathForText(basePath, currentlayer);
            if (!textPath) {
                alert('パスの複製に失敗しました');
                continue;
            }

            // Create Text on a path
            var textOnAPath = currentlayer.textFrames.pathText(textPath);

            // Duplicate textrange from original text frame item
            for (var i = 0; i < originalText.textRanges.length; i++) {
                originalText.textRanges[i].duplicate(textOnAPath);
            }

            // Remove original text frame item
            originalText.remove();

            // Select text on a path
            textOnAPath.selected = true;
        }
    }

    // Resolve which selected path to use for each text
    function resolveBasePathForText(selectedPaths, textCount, index) {
        if (!selectedPaths || selectedPaths.length === 0) return null;
        if (selectedPaths.length === textCount) return selectedPaths[index];
        return selectedPaths[0];
    }

    // Make ORIGINAL selected path invisible (no fill / no stroke)
    function makeOriginalPathInvisible(basePath) {
        try {
            if (basePath.typename === 'CompoundPathItem') {
                for (var cp = 0; cp < basePath.pathItems.length; cp++) {
                    basePath.pathItems[cp].stroked = false;
                    basePath.pathItems[cp].filled = false;
                }
            } else {
                basePath.stroked = false;
                basePath.filled = false;
            }
        } catch (_) { }
    }

    // Resolve actual PathItem to duplicate (CompoundPathItem -> first pathItem)
    function resolveSourcePath(basePath) {
        try {
            return (basePath.typename === 'CompoundPathItem') ? basePath.pathItems[0] : basePath;
        } catch (_) {
            return null;
        }
    }

    // Duplicate a path for the text and place it on the same layer
    function duplicatePathForText(basePath, currentLayer) {
        var srcPath = resolveSourcePath(basePath);
        if (!srcPath) return null;
        try {
            var dup = srcPath.duplicate();
            try { dup.move(currentLayer, ElementPlacement.PLACEATBEGINNING); } catch (_) { }
            return dup;
        } catch (eDup) {
            return null;
        }
    }

    // Get target point-text items (recursive)
    function getTargetTextItems(items) {
        var ti = [];
        for (var i = 0; i < items.length; i++) {
            if (items[i].typename === 'TextFrame' && items[i].kind === TextType.POINTTEXT) {
                ti.push(items[i]);
            } else if (items[i].typename === 'GroupItem') {
                ti = ti.concat(getTargetTextItems(items[i].pageItems));
            }
        }
        return ti;
    }

    // Get selected path items (PathItem). If GroupItem contains paths, collect them too.
    function getSelectedPathItems(items) {
        var pi = [];
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (it.typename === 'PathItem') {
                pi.push(it);
            } else if (it.typename === 'CompoundPathItem') {
                pi.push(it);
            } else if (it.typename === 'GroupItem') {
                pi = pi.concat(getSelectedPathItems(it.pageItems));
            }
        }
        return pi;
    }

}());1