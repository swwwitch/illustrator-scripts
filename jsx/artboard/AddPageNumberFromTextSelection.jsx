#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

AddPageNumberFromTextSelection.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AddPageNumberFromTextSelection.jsx

### 概要：

- 選択中のポイントテキスト（_pagenumber レイヤー推奨）を雛形に、すべてのアートボードへ連番テキストを配置します。
- 更新日：20260516
- 値の変更（接頭辞・開始番号・接尾辞・ゼロパディング・総ページ）や ↑↓/Shift+↑↓ による数値変更で、ドキュメント上のプレビューが即時更新されます（OK は確定のみ、キャンセルでプレビューを破棄）。
- _pagenumber レイヤーが無い場合は自動作成し、プレビュー中は一時的にロック解除・可視化・最前面移動して安全に処理します。OK/キャンセル時に元のロック／可視状態と重ね順を確実に復元します。

### 主な機能：

- 開始番号／接頭辞／接尾辞／ゼロパディング／総ページの指定
- 3カラムUI、ダイアログの透明度・表示位置の調整
- テンプレートが1件のみでも、プレビュー中に全アートボードへ仮増殖して確認可能（内部: pasteInAllArtboard）
- _pagenumber レイヤーの自動作成・一時解除・最前面化・確定/取消時の復元
- ダイアログタイトルにバージョン番号を表示（SCRIPT_VERSION）

### 処理の流れ：

1) 適当なポイントテキストを選択（_pagenumber レイヤー推奨）
2) ダイアログで各値を設定（↑↓/Shift+↑↓で数値調整、Zでゼロ埋め、Aで総ページ表示を切替）
3) 値変更ごとにライブプレビューを確認（プレビュー中は _pagenumber を一時解除＋最前面化）
4) OK で確定（プレビュー用バックアップ破棄＆状態復元）／キャンセルでプレビューを破棄して完全復元

### note：

- 対象はポイントテキストです。段落揃えなどの体裁変更は行いません。

### 更新履歴：

- v2.0.1 (20260516) : 内部リファクタリング（IIFE化、関数分割、命名整理、try/catch を共通ヘルパーへ集約）。動作の変更なし。
- v2.0 (20260108) : プレビュー時にUndo履歴を汚しにくいよう、app.undo() を用いたロールバック型プレビュー管理（PreviewManager）を追加。OK時は一括Undo後に本処理を1回で再実行し、取り消しを1回で戻せるように改善。
- v1.9 (20250810) : _pagenumber の自動作成、ロック/可視の一時解除、最前面化、OK/キャンセル時の重ね順・可視・ロックの完全復元を追加。プレビュー耐性を強化。
- v1.8 (20250625) : リファクタリング、ライブプレビューと全ABプレビュー複製に対応、UI整理
- v1.0 (20250625) : 初期バージョン

---

### Script Name：

AddPageNumberFromTextSelection.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/artboard/AddPageNumberFromTextSelection.jsx

### Overview：

- Places sequential page-number text on all artboards using the currently selected point text as a template (the _pagenumber layer is recommended).
- Updated: 2026-05-16
- Live preview updates instantly when values change (prefix, start number, suffix, zero padding, total pages) and on Up/Down or Shift+Up/Down. OK commits; Cancel discards the preview.
- If the _pagenumber layer doesn’t exist, it is auto-created. During preview, the layer is temporarily unlocked, shown, and moved to the top; on OK/Cancel, its lock/visibility and stacking order are fully restored.

### Main Features：

- Configure start number / prefix / suffix / zero padding / total pages
- 3-column dialog UI; adjustable dialog opacity and position
- Preview can temporarily duplicate a single template across all artboards (internally: pasteInAllArtboard)
- Auto-create, temporarily unlock/show/move-to-top for _pagenumber; fully restore state on commit/cancel
- Version number shown in the dialog title (SCRIPT_VERSION)

### Flow：

1) Select a point text object (the _pagenumber layer is recommended)
2) Set values in the dialog (Up/Down or Shift+Up/Down; Z toggles zero-pad; A toggles total pages)
3) Live preview updates every change (_pagenumber is temporarily prepared and brought to front)
4) Click OK to commit (discard preview backup & restore state) or Cancel to discard and fully restore

### Notes：

- Target must be a point text. This script does not change paragraph alignment or styling.

### Update History：

- v2.0.1 (2026-05-16): Internal refactor (IIFE wrapper, function splitting, clearer naming, try/catch consolidated into shared helpers). No behavior change.
- v2.0 (2026-01-08): Added rollback-based preview history management using app.undo() (PreviewManager). On OK, undo preview steps and re-run the final action once so users can undo the whole operation with a single Ctrl/Cmd+Z.
- v1.9 (2025-08-10): Added auto-create for _pagenumber, temporary unlock/visibility, move-to-top during preview, and full restoration of stacking order & visibility/lock on OK/Cancel; improved preview robustness.
- v1.8 (2025-06-25): Refactor; added live preview and all-artboards preview duplication; UI cleanup
- v1.0 (2025-06-25): Initial release
*/

(function () {

    // =========================================
    // バージョンとローカライズ / Version & Localization
    // =========================================

    var SCRIPT_VERSION = "v2.0.1";

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
            ja: "ポイントテキストを選択してください（_pagenumber レイヤー推奨）",
            en: "Please select a point text (the _pagenumber layer is recommended)."
        }
    };

    // =========================================
    // 安全実行ヘルパー / Safe Execution Helpers
    // =========================================

    /* 関数を try/catch 内で実行し、fn の戻り値を返す。例外時は onError(e) を呼ぶ（省略時は無視）。
       本スクリプトで唯一 try/catch を持つ基盤関数 / The single try/catch primitive: run fn and return its value, calling onError(e) on failure */
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

        // 変更操作を実行し、1ステップとしてカウント / Run an action and count it as one preview step
        this.addStep = function (func) {
            var self = this;
            runOrAlert("Preview Error", function () {
                if (typeof func === "function") func();
                self.undoDepth++;
                safeRedraw();
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
        return !!obj && obj.typename === "TextFrame";
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
    function capturePagenumberState(doc, layer) {
        var index = getLayerIndex(doc, layer);
        return {
            locked: layer.locked,
            visible: layer.visible,
            index: index,
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
    function restorePagenumberState(doc, layer, layerState) {
        if (!layer || !layerState) return;

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

    /* 選択中の TextFrame を指定レイヤーの先頭へ移動して返す（無ければ null）/ Move the selected TextFrame to the front of the layer and return it (or null) */
    function moveSelectionToLayer(doc, layer) {
        var textFrame = getSelectedTextFrame();
        if (!textFrame) return null;
        tryCall(function () { textFrame.move(layer, ElementPlacement.PLACEATBEGINNING); });
        return textFrame;
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
        if (!doc || isNaN(start)) return;
        var layer = findLayerByName(doc, layerName);
        if (!layer) return;

        // 雛形を決定：選択を優先、無ければアートボード順の先頭フレーム / pick the template: prefer the selection
        var template = getSelectedTextFrame() || sortFramesByArtboard(doc, layer.textFrames, null)[0];
        if (!template) return;

        rebuildFramesAcrossArtboards(doc, layer, template);

        // アートボード順に連番を流し込む / write sequential numbers in artboard order
        var frames = sortFramesByArtboard(doc, layer.textFrames, null);
        var maxNum = start + doc.artboards.length - 1;
        applyNumbering(frames, start, String(maxNum).length,
            prefix || "", suffix || "", !!zeroPad, maxNum, !!showTotal);

        safeRedraw();
    }

    /* 退避レイヤーに残ったテキストを _pagenumber へ戻し、退避レイヤーを削除 / Move any text left on the backup layer back to _pagenumber, then remove the backup layer */
    function restoreTextFromBackupLayer(doc) {
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
        if (app.documents.length === 0 || !getSelectedTextFrame()) {
            alert(LABELS.errorInvalidSelection[lang]);
            return;
        }

        var doc = app.activeDocument;

        // _pagenumber レイヤーを事前作成し、元状態を保存してから作業用に準備 / pre-create _pagenumber, capture its state, then prepare it for work
        var pagenumberLayer = getOrCreatePagenumberLayer(doc);
        var pagenumberLayerState = capturePagenumberState(doc, pagenumberLayer);
        preparePagenumberForWork(doc, pagenumberLayer);

        var ui = buildDialog();
        var previewManager = new PreviewManager();

        /* 現在の入力値でライブプレビューを更新（前回分を巻き戻し、1ステップとして再実行）/ Refresh the live preview with current input values (roll back the previous one, run as a single step) */
        function triggerPreview() {
            var start = parseInt(ui.startNumberField.text, 10);
            if (isNaN(start)) return;

            previewManager.rollback();
            previewManager.addStep(function () {
                updatePreview(doc, PAGENUMBER_LAYER_NAME, start,
                    ui.prefixField.text || "", ui.suffixField.text || "",
                    !!ui.zeroPadCheckbox.value, !!ui.totalPageCheckbox.value);
            });
        }

        /* 捕捉済みの _pagenumber レイヤー状態を復元 / Restore the captured _pagenumber layer state */
        function restoreCapturedState() {
            tryCall(function () {
                restorePagenumberState(doc, pagenumberLayer, pagenumberLayerState);
            });
        }

        /* OK確定時の本処理：選択テキストを _pagenumber へ移し、全アートボードへ連番を確定配置 / Final commit: move the selected text to _pagenumber and place sequential numbers on every artboard */
        function runFinalAction() {
            var startNum = parseInt(ui.startNumberField.text, 10);
            if (isNaN(startNum)) {
                alert(LABELS.errorNotNumber[lang]);
                return;
            }

            // プレビュー用バックアップを破棄 / discard the preview backup layer
            forceRemoveLayerByName(doc, TMP_LAYER_NAME);

            var layer = getOrCreatePagenumberLayer(doc);

            // 選択テキストを _pagenumber へ移し、他のテキストを除去 / move the selection onto _pagenumber and clear the rest
            var movedText = moveSelectionToLayer(doc, layer);
            removeOtherTextFrames(layer, movedText);

            // 雛形テキストを決定（選択優先、無ければ _pagenumber 内を探索）/ resolve the template text (prefer the selection)
            var targetText = getSelectedTextFrame() || findTextFrameOnAnyArtboard(layer, doc);
            if (!isTextFrame(targetText)) {
                alert(LABELS.errorInvalidSelection[lang]);
                return;
            }

            // 対象を _pagenumber 上へ確実に配置 / make sure the target sits on _pagenumber
            if (targetText.layer.name !== PAGENUMBER_LAYER_NAME) {
                tryCall(function () { targetText.move(layer, ElementPlacement.PLACEATBEGINNING); });
            }
            removeOtherTextFrames(layer, targetText);
            if (targetText.layer.name !== PAGENUMBER_LAYER_NAME) {
                alert(LABELS.errorInvalidSelection[lang]);
                return;
            }

            // 開始番号で初期化し、全アートボードへ複製 / seed with the start number and duplicate across all artboards
            seedAndPasteToAllArtboards(doc, targetText, startNum);

            // アートボード順に連番を流し込む / write sequential numbers in artboard order
            var frames = sortFramesByArtboard(doc, layer.textFrames, targetText);
            var maxNum = startNum + doc.artboards.length - 1;
            applyNumbering(frames, startNum, String(maxNum).length,
                ui.prefixField.text || "", ui.suffixField.text || "",
                !!ui.zeroPadCheckbox.value, maxNum, !!ui.totalPageCheckbox.value);

            safeRedraw();
            restoreCapturedState();
        }

        // 値の変更でプレビューを更新 / refresh the preview whenever a value changes
        ui.prefixField.addEventListener("changing", triggerPreview);
        ui.startNumberField.addEventListener("changing", triggerPreview);
        ui.suffixField.addEventListener("changing", triggerPreview);
        ui.zeroPadCheckbox.onClick = triggerPreview;
        ui.totalPageCheckbox.onClick = triggerPreview;
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
            restoreTextFromBackupLayer(doc);
            restoreCapturedState();
            ui.dialog.close(0);
        };

        ui.startNumberField.active = true;

        // 初回プレビュー / first preview pass
        tryCall(triggerPreview);

        ui.dialog.show();
    }

    main();

})();