#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

AddPageNumberFromTextSelection.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AddPageNumberFromTextSelection.jsx

### 概要：

選択中のポイントテキストを雛形に、すべてのアートボードへ連番テキストを配置します。
値を変更すると即時プレビューし、OKで確定、キャンセルでプレビューを破棄します。
_pagenumber レイヤーが無い場合は自動作成し、OK時は残します。キャンセル時は、元々存在しなかった場合のみ削除します。

### 主な機能：

- 開始番号／接頭辞／接尾辞／ゼロパディング／総ページ数表示
- ライブプレビューとキャンセル時の復元
- _pagenumber レイヤーの自動作成・一時解除・状態復元
- ダイアログタイトルにバージョン番号を表示

### note：

- 対象はポイントテキストです。段落揃えなどの体裁変更は行いません。

### 更新履歴：

- v2.0.1 (20260516) : 内部整理。OK確定時の雛形判定、プレビューUndo管理、_pagenumber 復元処理を改善。
- v2.0 (20260108) : PreviewManager によるロールバック型プレビュー管理を追加。
- v1.9 (20250810) : _pagenumber レイヤーの自動作成・一時解除・復元に対応。
- v1.8 (20250625) : ライブプレビューと全アートボード複製に対応。
- v1.0 (20250625) : 初期バージョン

----

### Script Name：

AddPageNumberFromTextSelection.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AddPageNumberFromTextSelection.jsx

### Overview：

Places sequential page-number text on all artboards using the selected point text as a template.
Value changes update the live preview immediately. OK commits the result; Cancel discards the preview.
If the _pagenumber layer does not exist, it is auto-created and kept on OK. On Cancel, it is removed only if it did not originally exist.

### Main Features：

- Start number / prefix / suffix / zero padding / total page display
- Live preview with cancel restoration
- Auto-create, temporarily unlock, and restore the _pagenumber layer
- Version number shown in the dialog title

### Notes：

- Target must be a point text. This script does not change paragraph alignment or styling.

### Update History：

- v2.0.1 (2026-05-16): Internal cleanup. Improved final template detection, preview undo tracking, and _pagenumber restoration.
- v2.0 (2026-01-08): Added rollback-based preview management with PreviewManager.
- v1.9 (2025-08-10): Added auto-create, temporary unlock, and restoration for the _pagenumber layer.
- v1.8 (2025-06-25): Added live preview and all-artboards duplication.
- v1.0 (2025-06-25): Initial release
*/

(function () {

    // =========================================
    // バージョンとローカライズ / Version & Localization
    // =========================================

    var SCRIPT_VERSION = "v2.0.2";

    // 連番テキストを配置する対象レイヤー名 / Layer that receives the page-number text
    var PAGENUMBER_LAYER_NAME = "_pagenumber";
    // プレビュー中の雛形を退避する一時レイヤー名 / Temp layer used to back up the template during preview
    var TMP_LAYER_NAME = "_pagenumber_preview";

    /* 実行環境のUI言語を判定（日本語環境は "ja"、その他は "en"）/ Detect the environment's UI language ("ja" for Japanese, otherwise "en") */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var lang = getCurrentLang();

    // UI 文字列（OK ボタンは非ローカライズのため LABELS に含めない）/ UI strings (the OK button is not localized, so it is not in LABELS)
    var LABELS = {
        dialogTitle: { ja: "ページ番号を追加 " + SCRIPT_VERSION, en: "Add Page Numbers " + SCRIPT_VERSION },
        prefixLabel: { ja: "接頭辞", en: "Prefix" },
        startLabel: { ja: "開始番号", en: "Starting number" },
        zeroPadLabel: { ja: "ゼロパディング", en: "Zero pad" },
        suffixLabel: { ja: "接尾辞", en: "Suffix" },
        totalPageLabel: { ja: "総ページ数を表示", en: "Show total pages" },
        cancelLabel: { ja: "キャンセル", en: "Cancel" },
        errorNotNumber: { ja: "有効な数字を入力してください", en: "Please enter a valid number" },
        errorInvalidSelection: {
            ja: "各アートボードのノンブルのテンプレとなるテキストを入力し、選択してください。",
            en: "Please create and select the template text used for page numbers on each artboard."
        }
    };

    // =========================================
    // 安全実行ヘルパー / Safe Execution Helpers
    // =========================================

    /* 関数を try/catch 内で実行し、fn の戻り値を返す。例外時は onError(e) を呼ぶ（省略時は無視）。
       例外を握りつぶしてよい処理の共通ヘルパー / Shared helper for operations where ignored exceptions are acceptable */
    function tryCall(fn, onError) {
        try {
            return fn ? fn() : undefined;
        } catch (e) {
            if (onError) onError(e);
        }
    }

    /* プロパティ代入を例外無視で実行 / Assign a property, ignoring any error */
    function trySet(obj, prop, value) {
        tryCall(function () { if (obj) obj[prop] = value; });
    }

    /* 関数を実行し、例外時は label 付きでアラート表示 / Run a function; on error show an alert prefixed with label */
    function runOrAlert(label, fn) {
        tryCall(fn, function (e) { alert(label + ": " + e); });
    }

    /* 画面を安全に再描画 / Redraw the screen safely */
    function safeRedraw() {
        tryCall(function () { app.redraw(); });
    }

    /* 指定名のレイヤーを返す（無ければ undefined）/ Return a layer by name, or undefined if it does not exist */
    function findLayerByName(doc, name) {
        return tryCall(function () { return doc.layers.getByName(name); });
    }

    // =========================================
    // レイヤー操作 / Layer Operations
    // =========================================

    /* 指定名のレイヤーを取得、無ければ新規作成して返す / Get a layer by name, creating it if it does not exist */
    function getOrCreateLayer(doc, name) {
        var layer = findLayerByName(doc, name);
        if (!layer) {
            layer = doc.layers.add();
            layer.name = name;
        }
        return layer;
    }

    /* _pagenumber レイヤーを取得、無ければ新規作成して返す / Get the _pagenumber layer, creating it if needed */
    function getOrCreatePagenumberLayer(doc) {
        return getOrCreateLayer(doc, PAGENUMBER_LAYER_NAME);
    }

    /* 指定名のレイヤーを確実に削除（ロック解除→中身を全削除→レイヤー削除）/ Force-remove a layer by name (unlock, purge contents, then remove) */
    function forceRemoveLayerByName(doc, name) {
        var layer = findLayerByName(doc, name);
        if (!layer) return;

        trySet(layer, 'locked', false);
        trySet(layer, 'visible', true);

        for (var i = layer.pageItems.length - 1; i >= 0; i--) {
            var item = layer.pageItems[i];
            trySet(item, 'locked', false);
            tryCall(function () { item.remove(); });
        }
        for (var j = layer.layers.length - 1; j >= 0; j--) {
            var subLayer = layer.layers[j];
            tryCall(function () { subLayer.remove(); });
        }
        tryCall(function () { layer.remove(); });
    }

    // =========================================
    // Undo / プレビュー管理 / Undo & Preview Manager
    // =========================================

    /* プレビュー編集をUndoステップとして積み、巻き戻し・確定を一括管理するクラス / Manages preview edits as undo steps for batch rollback or commit */
    function PreviewManager() {
        this.undoDepth = 0; // プレビュー中に実行したアクション数 / number of preview actions executed

        // 変更操作を実行し、実際に変更があった場合だけ1ステップとしてカウント / Run an action and count it only when it actually changes the document
        this.addStep = function (func) {
            var self = this;
            runOrAlert("Preview Error", function () {
                var changed = (typeof func === "function") ? func() : false;
                if (changed) {
                    self.undoDepth++;
                    safeRedraw();
                }
            });
        };

        // プレビュー分の変更をすべて取り消す / Roll back all preview steps
        this.rollback = function () {
            while (this.undoDepth > 0) {
                try {
                    app.undo();
                } catch (e) {
                    break;
                }
                this.undoDepth--;
            }
            safeRedraw();
        };

        // 確定：プレビュー分を全て取り消してから本処理を1回実行 / Commit: undo all preview steps, then run the final action once
        this.confirm = function (finalAction) {
            this.rollback();
            if (typeof finalAction === "function") {
                runOrAlert("Final Error", finalAction);
            }
        };
    }

    // =========================================
    // 選択・型判定 / Selection & Type Guards
    // =========================================

    /* オブジェクトが TextFrame かどうかを判定 / Return true if the object is a TextFrame */
    function isTextFrame(obj) {
        return tryCall(function () {
            return !!obj && obj.typename === "TextFrame";
        }) === true;
    }

    /* 選択先頭が TextFrame ならそれを返す（無ければ null）/ Return the selected TextFrame, or null if none is selected */
    function getSelectedTextFrame() {
        var selection = tryCall(function () { return app.selection; });
        if (selection && selection.length > 0 && isTextFrame(selection[0])) {
            return selection[0];
        }
        return null;
    }

    // =========================================
    // _pagenumber レイヤーの状態管理 / Pagenumber Layer State
    // =========================================

    /* ドキュメント内でのレイヤーの並び順インデックスを返す（無ければ -1）/ Return the stacking-order index of a layer (or -1) */
    function getLayerIndex(doc, layer) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i] === layer) return i;
        }
        return -1;
    }

    /* 指定インデックスのレイヤー名を返す（取得不可なら undefined）/ Return the layer name at the given index (or undefined) */
    function getLayerNameAtIndex(doc, index) {
        return tryCall(function () { return doc.layers[index].name; });
    }

    /* _pagenumber レイヤーの現在状態（ロック・表示・並び順）を記録 / Capture the current state (lock, visibility, stacking order) of the _pagenumber layer */
    function capturePagenumberState(doc, layer, existed) {
        var index = getLayerIndex(doc, layer);
        return {
            existed: !!existed,
            locked: layer.locked,
            visible: layer.visible,
            // ひとつ上（前面側）のレイヤー名を復元の基準として記録 / remember the neighbor layer above as a restore anchor
            neighborAboveName: (index > 0) ? getLayerNameAtIndex(doc, index - 1) : null
        };
    }

    /* _pagenumber レイヤーを作業用に準備（ロック解除・表示・最前面へ）/ Prepare the _pagenumber layer for work (unlock, show, move to top) */
    function preparePagenumberForWork(doc, layer) {
        trySet(layer, 'locked', false);
        trySet(layer, 'visible', true);
        tryCall(function () { layer.move(doc, ElementPlacement.PLACEATBEGINNING); });
    }

    /* capturePagenumberState で記録した状態へ _pagenumber レイヤーを復元 / Restore the _pagenumber layer to the captured state */
    function restorePagenumberState(doc, layer, layerState, removeIfNew) {
        if (!layer || !layerState) return;

        // キャンセル時のみ、元々存在しなかった _pagenumber を削除 / remove an auto-created _pagenumber only on Cancel
        if (!layerState.existed && removeIfNew) {
            forceRemoveLayerByName(doc, PAGENUMBER_LAYER_NAME);
            return;
        }

        // 並び順を復元 / restore stacking order
        if (layerState.neighborAboveName) {
            var neighborLayer = findLayerByName(doc, layerState.neighborAboveName);
            if (neighborLayer) {
                tryCall(function () { layer.move(neighborLayer, ElementPlacement.PLACEAFTER); });
            }
        } else {
            tryCall(function () { layer.move(doc, ElementPlacement.PLACEATBEGINNING); });
        }

        // 表示・ロック状態を復元 / restore visibility & lock
        trySet(layer, 'visible', layerState.visible);
        trySet(layer, 'locked', layerState.locked);
    }

    // =========================================
    // アートボードとフレームの探索 / Artboard & Frame Lookup
    // =========================================

    /* 座標 pos が矩形 rect 内にあるか判定 / Return true if point pos is inside rectangle rect */
    function isPointInRect(pos, rect) {
        return pos[0] >= rect[0] && pos[0] <= rect[2] && pos[1] <= rect[1] && pos[1] >= rect[3];
    }

    /* 指定座標が含まれるアートボードのインデックスを返す（無ければ -1）/ Return the index of the artboard containing the given point (or -1) */
    function getArtboardIndexByPosition(doc, pos) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isPointInRect(pos, doc.artboards[i].artboardRect)) return i;
        }
        return -1;
    }

    /* 指定アートボード上にある最初の TextFrame を返す（無ければ null）/ Return the first TextFrame located on the given artboard (or null) */
    function getFirstTextFrameOnArtboard(layer, doc, artboardIndex) {
        var artboardRect = doc.artboards[artboardIndex].artboardRect;
        for (var i = 0; i < layer.textFrames.length; i++) {
            var frame = layer.textFrames[i];
            if (isPointInRect(frame.position, artboardRect)) return frame;
        }
        return null;
    }

    /* いずれかのアートボード上で最初に見つかった TextFrame を返す / Return the first TextFrame found on any artboard */
    function findTextFrameOnAnyArtboard(layer, doc) {
        for (var i = 0; i < doc.artboards.length; i++) {
            var frame = getFirstTextFrameOnArtboard(layer, doc, i);
            if (frame) return frame;
        }
        return null;
    }

    /* TextFrame 群をアートボード順に並べた配列を返す（exclude は除外）/ Return the TextFrames sorted by artboard order (exclude is skipped) */
    function sortFramesByArtboard(doc, textFrames, exclude) {
        var entries = [];
        for (var i = 0; i < textFrames.length; i++) {
            var frame = textFrames[i];
            if (exclude && frame === exclude) continue;
            entries.push({ frame: frame, artboardIndex: getArtboardIndexByPosition(doc, frame.position) });
        }
        entries.sort(function (a, b) { return a.artboardIndex - b.artboardIndex; });

        var sorted = [];
        for (var j = 0; j < entries.length; j++) sorted.push(entries[j].frame);
        return sorted;
    }

    // =========================================
    // ページ番号テキストの生成・配置 / Page Number Generation & Placement
    // =========================================

    /* 番号・接頭辞/接尾辞・ゼロ埋め・総ページ表示からページ番号文字列を生成 / Build the page-number string from number, prefix/suffix, zero padding, and the optional total */
    function buildPageNumberString(num, maxDigits, prefix, suffix, zeroPad, totalPageNum, showTotal) {
        var numStr = String(num);
        if (zeroPad && numStr.length < maxDigits) {
            numStr = Array(maxDigits - numStr.length + 1).join("0") + numStr; // ES3対応ゼロ埋め / ES3-safe zero pad
        }
        var result = (prefix || "") + numStr + (suffix || "");
        if (showTotal) result += "/" + totalPageNum;
        return result;
    }

    /* アートボード順に並んだフレーム配列へ連番テキストを流し込む / Write sequential page-number text into the sorted frame list */
    function applyNumbering(frames, startNum, maxDigits, prefix, suffix, zeroPad, maxNum, showTotal) {
        for (var i = 0; i < frames.length; i++) {
            var text = buildPageNumberString(startNum + i, maxDigits, prefix, suffix, zeroPad, maxNum, showTotal);
            trySet(frames[i], 'contents', text);
        }
    }

    /* 指定レイヤー上の TextFrame を except 以外すべて削除 / Remove every TextFrame on the layer except the given one */
    function removeOtherTextFrames(layer, except) {
        var frames = layer.textFrames;
        for (var i = frames.length - 1; i >= 0; i--) {
            var frame = frames[i];
            if (frame === except) continue;
            trySet(frame, 'locked', false);
            tryCall(function () { frame.remove(); });
        }
    }

    /* 対象テキストを開始番号で初期化し、カット→全アートボードへ貼り付け / Seed the target text with the start number, then cut and paste it onto every artboard */
    function seedAndPasteToAllArtboards(doc, targetText, startNum) {
        trySet(targetText, 'contents', String(startNum));

        // 所属レイヤーの一時状態を退避 / back up the source layer's state
        var sourceLayer = tryCall(function () { return targetText.layer; });
        var prevLocked = sourceLayer ? sourceLayer.locked : null;
        var prevVisible = sourceLayer ? sourceLayer.visible : null;

        // 対象と所属レイヤーを一時的にロック解除＆可視化 / temporarily unlock & show the target and its layer
        trySet(targetText, 'locked', false);
        trySet(sourceLayer, 'locked', false);
        trySet(sourceLayer, 'visible', true);

        // 対象が乗るアートボードをアクティブ化 / activate the artboard the target sits on
        var sourceArtboardIndex = getArtboardIndexByPosition(doc, targetText.position);
        if (sourceArtboardIndex >= 0) {
            tryCall(function () { doc.artboards.setActiveArtboardIndex(sourceArtboardIndex); });
        }

        // 選択→カット→全アートボードへ貼り付け / select -> cut -> paste onto all artboards
        tryCall(function () { app.selection = null; });
        trySet(targetText, 'selected', true);
        tryCall(function () { app.cut(); });
        if (sourceLayer) trySet(doc, 'activeLayer', sourceLayer);
        tryCall(function () { app.executeMenuCommand('pasteInAllArtboard'); });

        // レイヤーの一時状態を元へ戻す / restore the layer's temporary state
        if (sourceLayer) {
            trySet(sourceLayer, 'locked', prevLocked);
            trySet(sourceLayer, 'visible', prevVisible);
        }
    }

    // =========================================
    // ライブプレビュー / Live Preview
    // =========================================

    /* 雛形を一時レイヤーへ退避しつつ、全アートボードへクリーンに複製し直す / Back up the template, then cleanly re-duplicate it across every artboard */
    function rebuildFramesAcrossArtboards(doc, layer, template) {
        // 退避用一時レイヤーへ非表示バックアップを作成 / make a hidden backup on a temp layer
        var tempLayer = getOrCreateLayer(doc, TMP_LAYER_NAME);
        trySet(tempLayer, 'visible', false);
        trySet(tempLayer, 'locked', false);

        var backup = null;
        tryCall(function () { backup = template.duplicate(tempLayer, ElementPlacement.PLACEATBEGINNING); });
        trySet(backup, 'visible', false);
        trySet(backup, 'locked', true);

        // 雛形と所属レイヤーを一時的にロック解除＆可視化 / temporarily unlock & show the template and its layer
        trySet(template, 'locked', false);
        var sourceLayer = tryCall(function () { return template.layer; });
        trySet(sourceLayer, 'locked', false);
        trySet(sourceLayer, 'visible', true);

        // 雛形が乗るアートボードをアクティブ化 / activate the artboard the template sits on
        var sourceArtboardIndex = getArtboardIndexByPosition(doc, template.position);
        if (sourceArtboardIndex >= 0) {
            tryCall(function () { doc.artboards.setActiveArtboardIndex(sourceArtboardIndex); });
        }

        // 選択→カット→既存を一掃→全アートボードへ貼り付け / select -> cut -> clear -> paste onto all artboards
        tryCall(function () { app.selection = null; });
        trySet(template, 'selected', true);
        tryCall(function () { app.cut(); });
        removeOtherTextFrames(layer, null);
        trySet(doc, 'activeLayer', layer);
        tryCall(function () { app.executeMenuCommand('pasteInAllArtboard'); });
    }

    /* 選択テキスト（無ければレイヤー上の先頭テキスト）を雛形に、全アートボードへ連番をプレビュー / Render a sequential-numbering preview on every artboard, using the selected text (or the first text on the layer) as a template */
    function updatePreview(doc, layerName, start, prefix, suffix, zeroPad, showTotal) {
        if (!doc || isNaN(start)) return false;
        var layer = findLayerByName(doc, layerName);
        if (!layer) return false;

        // 雛形を決定：選択を優先、無ければアートボード順の先頭フレーム / pick the template: prefer the selection
        var template = getSelectedTextFrame() || sortFramesByArtboard(doc, layer.textFrames, null)[0];
        if (!template) return false;

        rebuildFramesAcrossArtboards(doc, layer, template);

        // アートボード順に連番を流し込む / write sequential numbers in artboard order
        var frames = sortFramesByArtboard(doc, layer.textFrames, null);
        var maxNum = start + doc.artboards.length - 1;
        applyNumbering(frames, start, String(maxNum).length,
            prefix || "", suffix || "", !!zeroPad, maxNum, !!showTotal);

        safeRedraw();
        return true;
    }

    /* 退避レイヤーに残ったテキストを _pagenumber へ戻し、退避レイヤーを削除 / Move any text left on the backup layer back to _pagenumber, then remove the backup layer */
    function restorePreviewBackupOnCancel(doc) {
        var backupLayer = findLayerByName(doc, TMP_LAYER_NAME);
        if (!backupLayer) return;

        var pagenumberLayer = getOrCreatePagenumberLayer(doc);
        removeOtherTextFrames(pagenumberLayer, null);
        trySet(backupLayer, 'locked', false);
        trySet(backupLayer, 'visible', true);

        for (var i = backupLayer.pageItems.length - 1; i >= 0; i--) {
            var item = backupLayer.pageItems[i];
            trySet(item, 'locked', false);
            tryCall(function () { item.move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING); });
            trySet(item, 'visible', true);
        }
        forceRemoveLayerByName(doc, TMP_LAYER_NAME);
        safeRedraw();
    }

    // =========================================
    // キーボード操作 / Keyboard Handlers
    // =========================================

    /* 入力欄で↑↓キーによる増減を有効化（Shiftで10の倍数へスナップ）/ Enable Up/Down arrow increment-decrement on an edittext (Shift snaps to multiples of 10) */
    function changeValueByArrowKey(editText, onChanged) {
        if (!editText || !editText.addEventListener) return;
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyName = event.keyName;
            var isUp = (keyName === "Up" || keyName === "UpArrow");
            var isDown = (keyName === "Down" || keyName === "DownArrow");
            if (!isUp && !isDown) return;

            if (ScriptUI.environment.keyboardState.shiftKey) {
                // 10の倍数へスナップ / snap to a multiple of 10
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else {
                value += isUp ? 1 : -1;
            }
            if (value < 0) value = 0;

            editText.text = Math.round(value);
            event.preventDefault();
            if (typeof onChanged === "function") tryCall(onChanged);
        });
    }

    /* 指定キー押下でチェックボックスをトグルするハンドラを登録 / Register a handler that toggles a checkbox when the given key is pressed */
    function addToggleKeyHandler(dialog, keyChar, checkbox, onChanged, skipWhenEditTextFocus) {
        if (!dialog || !checkbox || !dialog.addEventListener) return;
        dialog.addEventListener("keydown", function (event) {
            // 入力欄フォーカス中はスキップしたい場合のみスキップ / skip while an edittext is focused, only when requested
            if (skipWhenEditTextFocus && event.target && event.target.type === "edittext") return;
            if ((event.keyName || "").toUpperCase() !== String(keyChar).toUpperCase()) return;

            checkbox.value = !checkbox.value;
            if (typeof onChanged === "function") tryCall(onChanged);
            event.preventDefault();
        });
    }

    // =========================================
    // UI 構築 / UI Construction
    // =========================================

    /* 親に縦並びカラム（group）を追加して返す / Add a vertical column group to the parent and return it */
    function addColumn(parent, align) {
        var column = parent.add("group");
        column.orientation = "column";
        column.alignChildren = align || "left";
        return column;
    }

    /* ラベル付き入力欄を追加し、入力欄（edittext）を返す / Add a labeled edittext and return the edittext */
    function addLabeledEditText(parent, label, initValue, width) {
        parent.add("statictext", undefined, label);
        var field = parent.add("edittext", undefined, initValue);
        field.characters = width;
        return field;
    }

    /* チェックボックスを追加して返す / Add a checkbox and return it */
    function addCheckbox(parent, label) {
        return parent.add("checkbox", undefined, label);
    }

    /* ダイアログと各UIコントロールを生成し、参照をまとめて返す / Build the dialog and its controls, returning all references */
    function buildDialog() {
        var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
        dialog.orientation = "column";
        dialog.alignChildren = "left";

        // 3カラムレイアウト / 3-column layout
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";

        // 左カラム: 接頭辞 / left column: prefix
        var leftGroup = addColumn(columnsGroup);
        var prefixField = addLabeledEditText(leftGroup, LABELS.prefixLabel[lang], "", 10);

        // 中央カラム: 開始番号 + ゼロ埋め / center column: start number + zero pad
        var centerGroup = addColumn(columnsGroup);
        var startNumberField = addLabeledEditText(centerGroup, LABELS.startLabel[lang], "1", 6);
        var zeroPadCheckbox = addCheckbox(centerGroup, LABELS.zeroPadLabel[lang]);

        // 右カラム: 接尾辞 + 総ページ表示 / right column: suffix + show-total
        var rightGroup = addColumn(columnsGroup);
        var suffixField = addLabeledEditText(rightGroup, LABELS.suffixLabel[lang], "", 10);
        var totalPageCheckbox = addCheckbox(rightGroup, LABELS.totalPageLabel[lang]);

        // ボタン行（右寄せ、キャンセル → OK の順）/ button row (right-aligned, Cancel then OK)
        var buttonRow = dialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignChildren = ["right", "center"];
        buttonRow.margins = [10, 10, 10, 0];
        buttonRow.alignment = ["right", "bottom"];
        var cancelButton = buttonRow.add("button", undefined, LABELS.cancelLabel[lang], { name: "cancel" });
        var okButton = buttonRow.add("button", undefined, "OK", { name: "ok" });

        // 透明度と表示位置の調整 / adjust opacity and position
        dialog.opacity = 0.97;
        dialog.onShow = function () {
            dialog.location = [dialog.location[0] + 300, dialog.location[1]];
        };

        return {
            dialog: dialog,
            prefixField: prefixField,
            startNumberField: startNumberField,
            zeroPadCheckbox: zeroPadCheckbox,
            suffixField: suffixField,
            totalPageCheckbox: totalPageCheckbox,
            cancelButton: cancelButton,
            okButton: okButton
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* ダイアログを表示し、選択テキストを雛形に全アートボードへページ番号を配置 / Show the dialog and place page numbers on every artboard using the selected text as a template */
    function main() {
        // テキスト未選択なら終了 / Exit if no text is selected
        var originalTemplateText = getSelectedTextFrame();
        if (app.documents.length === 0 || !originalTemplateText) {
            alert(LABELS.errorInvalidSelection[lang]);
            return;
        }

        var doc = app.activeDocument;

        // _pagenumber レイヤーの存在有無を記録
        var pagenumberLayerExisted = !!findLayerByName(doc, PAGENUMBER_LAYER_NAME);

        // _pagenumber レイヤーを事前作成し、元状態を保存してから作業用に準備 / pre-create _pagenumber, capture its state, then prepare it for work
        var pagenumberLayer = getOrCreatePagenumberLayer(doc);
        var pagenumberLayerState = capturePagenumberState(doc, pagenumberLayer, pagenumberLayerExisted);
        preparePagenumberForWork(doc, pagenumberLayer);

        var ui = buildDialog();
        var previewManager = new PreviewManager();

        /* 現在の入力値でライブプレビューを更新（前回分を巻き戻し、1ステップとして再実行）/ Refresh the live preview with current input values (roll back the previous one, run as a single step) */
        function triggerPreview() {
            var start = parseInt(ui.startNumberField.text, 10);
            if (isNaN(start)) return;

            previewManager.rollback();
            previewManager.addStep(function () {
                return updatePreview(doc, PAGENUMBER_LAYER_NAME, start,
                    ui.prefixField.text || "", ui.suffixField.text || "",
                    !!ui.zeroPadCheckbox.value, !!ui.totalPageCheckbox.value);
            });
        }

        /* 捕捉済みの _pagenumber レイヤー状態を復元 / Restore the captured _pagenumber layer state */
        function restoreCapturedState(removeIfNew) {
            tryCall(function () {
                restorePagenumberState(doc, pagenumberLayer, pagenumberLayerState, !!removeIfNew);
            });
        }

        /* OK確定時の雛形テキストを取得 / Resolve the template text used for the final commit */
        function resolveFinalTemplateText(layer) {
            var targetText = isTextFrame(originalTemplateText) ? originalTemplateText : getSelectedTextFrame();
            if (!isTextFrame(targetText)) {
                targetText = findTextFrameOnAnyArtboard(layer, doc);
            }
            return isTextFrame(targetText) ? targetText : null;
        }

        /* 雛形テキストを _pagenumber 上へ移し、他のテキストを除去 / Move the template to _pagenumber and remove the other text frames */
        function prepareFinalTemplateText(layer, targetText) {
            if (!isTextFrame(targetText)) return false;
            if (targetText.layer.name !== PAGENUMBER_LAYER_NAME) {
                tryCall(function () { targetText.move(layer, ElementPlacement.PLACEATBEGINNING); });
            }
            if (targetText.layer.name !== PAGENUMBER_LAYER_NAME) return false;
            removeOtherTextFrames(layer, targetText);
            return true;
        }

        /* 確定用の連番を全アートボードへ適用 / Apply final sequential page numbers to every artboard */
        function applyFinalNumbering(layer, targetText, startNum) {
            seedAndPasteToAllArtboards(doc, targetText, startNum);

            var frames = sortFramesByArtboard(doc, layer.textFrames, targetText);
            var maxNum = startNum + doc.artboards.length - 1;
            applyNumbering(frames, startNum, String(maxNum).length,
                ui.prefixField.text || "", ui.suffixField.text || "",
                !!ui.zeroPadCheckbox.value, maxNum, !!ui.totalPageCheckbox.value);
        }

        /* OK確定時の本処理：雛形テキストを _pagenumber へ移し、全アートボードへ連番を確定配置 / Final commit: move the template text to _pagenumber and place sequential numbers on every artboard */
        function runFinalAction() {
            var startNum = parseInt(ui.startNumberField.text, 10);
            if (isNaN(startNum)) {
                alert(LABELS.errorNotNumber[lang]);
                return;
            }

            // プレビュー用バックアップを破棄 / discard the preview backup layer
            forceRemoveLayerByName(doc, TMP_LAYER_NAME);

            var layer = getOrCreatePagenumberLayer(doc);
            var targetText = resolveFinalTemplateText(layer);
            if (!targetText || !prepareFinalTemplateText(layer, targetText)) {
                alert(LABELS.errorInvalidSelection[lang]);
                return;
            }

            applyFinalNumbering(layer, targetText, startNum);
            safeRedraw();
            restoreCapturedState(false);
        }

        changeValueByArrowKey(ui.startNumberField, triggerPreview);

        // Zキーでゼロ埋め、Aキーで総ページ表示をトグル / Z toggles zero-pad, A toggles show-total
        addToggleKeyHandler(ui.dialog, "Z", ui.zeroPadCheckbox, triggerPreview, false);
        addToggleKeyHandler(ui.dialog, "A", ui.totalPageCheckbox, triggerPreview, true);

        // OK：プレビュー分を全Undoしてから本処理を1回だけ実行 / OK: undo all preview steps, then run the final action once
        ui.okButton.onClick = function () {
            previewManager.confirm(runFinalAction);
            tryCall(function () { forceRemoveLayerByName(doc, TMP_LAYER_NAME); });
            ui.dialog.close(1);
        };

        // キャンセル：プレビューを巻き戻し、退避テキストと _pagenumber 状態を復元 / Cancel: roll back the preview, restore the backed-up text and the _pagenumber state
        ui.cancelButton.onClick = function () {
            previewManager.rollback();
            restorePreviewBackupOnCancel(doc);
            restoreCapturedState(true);
            ui.dialog.close(0);
        };

        ui.startNumberField.active = true;

        // 初回プレビュー / first preview pass
        tryCall(triggerPreview);

        ui.dialog.show();
    }

    main();

})();