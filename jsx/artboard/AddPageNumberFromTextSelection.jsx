#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択中のポイントテキストを雛形に、すべてのアートボードへページ番号を配置
- 接頭辞／接尾辞／ゼロ埋め／総ページ数表示に対応し、変更は即時プレビュー
- 配置先は _pagenumber レイヤー（無ければ自動作成）
- 対象はポイントテキストのみ。体裁の変更は行わない

### Overview

- Places page numbers on every artboard, using the selected point text as a template
- Supports prefix / suffix / zero padding / total page display, with live preview
- Text is placed on the _pagenumber layer (auto-created when missing)
- Point text only; paragraph alignment and styling are left untouched

*/

(function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================
    var SCRIPT_NAME     = "AddPageNumberFromTextSelection"; /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v2.1.0";                       /* バージョン / version */
    var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
    var SCRIPT_RELEASED = "2025-06-25";                   /* 最初のリリース日 / first release date */
    var SCRIPT_UPDATED  = "2026-08-19";                   /* 更新日 / last updated */

    // README (Japanese)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddPageNumberFromTextSelection.md
    // README (English)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddPageNumberFromTextSelection.md

    // Released under the MIT license
    // http://opensource.org/licenses/mit-license.php

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    // 連番テキストを配置する対象レイヤー名 / Layer that receives the page-number text
    var PAGENUMBER_LAYER_NAME = "_pagenumber";
    // プレビュー中の雛形を退避する一時レイヤー名 / Temp layer used to back up the template during preview
    var BACKUP_LAYER_NAME = "_pagenumber_preview";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 実行環境のUI言語を判定（日本語環境は "ja"、その他は "en"）/ Detect the environment's UI language ("ja" for Japanese, otherwise "en") */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var currentLanguage = getCurrentLang();

    // UI 文字列（OK ボタンのラベルは非ローカライズ）/ UI strings (the OK button label is not localized)
    var LABELS = {
        dialog: {
            title: { ja: "ページ番号を一括配置", en: "Place Page Numbers" }
        },
        field: {
            prefix: { ja: "接頭辞", en: "Prefix" },
            start: { ja: "開始番号", en: "Start number" },
            suffix: { ja: "接尾辞", en: "Suffix" }
        },
        checkbox: {
            zeroPad: { ja: "ゼロ埋め", en: "Zero padding" },
            showTotal: { ja: "総ページ数を表示", en: "Show total pages" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            prefix: {
                ja: "番号の前に付ける文字列（例：P.）",
                en: "Text placed before the number (e.g. P.)"
            },
            start: {
                ja: "先頭のアートボードに付ける番号。↑↓キーで増減、Shift+↑↓で10単位。",
                en: "Number for the first artboard. Up/Down to change, Shift+Up/Down for steps of 10."
            },
            suffix: {
                ja: "番号の後ろに付ける文字列（例：ページ）",
                en: "Text placed after the number (e.g. page)"
            },
            zeroPad: {
                ja: "総ページ数の桁数に合わせて0を補います（例：1 → 01）。Zキーで切り替え。",
                en: "Pads numbers with zeros to match the total (e.g. 1 → 01). Press Z to toggle."
            },
            showTotal: {
                ja: "「番号/総ページ数」の形式で表示します（例：3/12）。Aキーで切り替え。",
                en: "Shows the number as \"current/total\" (e.g. 3/12). Press A to toggle."
            },
            cancel: {
                ja: "プレビューを破棄して閉じます。",
                en: "Discard the preview and close."
            },
            ok: {
                ja: "プレビューの内容で確定し、" + PAGENUMBER_LAYER_NAME + " レイヤーに配置します。",
                en: "Commit the preview and place the text on the " + PAGENUMBER_LAYER_NAME + " layer."
            }
        },
        alert: {
            notNumber: {
                ja: "開始番号には数値を入力してください。",
                en: "Enter a number for the start number."
            },
            invalidSelection: {
                ja: "ページ番号の雛形となるポイントテキストを1つ選択してから実行してください。",
                en: "Select a single point text object to use as the page-number template, then run the script again."
            },
            commitFailed: {
                ja: "ページ番号を配置できませんでした。テキストや配置先レイヤーのロック状態を確認してください。",
                en: "Could not place the page numbers. Check whether the text or the destination layer is locked."
            }
        }
    };

    // =========================================
    // 安全実行ヘルパー / Safe Execution Helpers
    // =========================================

    /* 関数を try/catch 内で実行し、action の戻り値を返す。例外時は onError(e) を呼ぶ（省略時は無視）。
       例外を握りつぶしてよい処理の共通ヘルパー / Shared helper for operations where ignored exceptions are acceptable */
    function tryCall(action, onError) {
        try {
            return action ? action() : undefined;
        } catch (e) {
            if (onError) onError(e);
        }
    }

    /* プロパティ代入を例外無視で実行（ロック中・削除済みオブジェクトは代入で例外を出すため）/ Assign a property, ignoring any error (locked or deleted objects throw on assignment) */
    function trySetProperty(target, propertyName, value) {
        tryCall(function () { if (target) target[propertyName] = value; });
    }

    /* 関数を実行し、例外時は errorLabel 付きでアラート表示 / Run a function; on error show an alert prefixed with errorLabel */
    function runOrAlert(errorLabel, action) {
        tryCall(action, function (e) { alert(errorLabel + ": " + e); });
    }

    /* 画面を安全に再描画 / Redraw the screen safely */
    function safeRedraw() {
        tryCall(function () { app.redraw(); });
    }

    // =========================================
    // レイヤー操作 / Layer Operations
    // =========================================

    /* 指定名のレイヤーをサブレイヤーまで含めて探す（無ければ null）/ Find a layer by name, including sub-layers (null when it does not exist) */
    function findLayerByName(container, layerName) {
        for (var i = 0; i < container.layers.length; i++) {
            var layer = container.layers[i];
            if (layer.name === layerName) return layer;
            // サブレイヤーに同名があっても取りこぼさない / do not miss a nested layer with the same name
            var nestedLayer = findLayerByName(layer, layerName);
            if (nestedLayer) return nestedLayer;
        }
        return null;
    }

    /* 指定名のレイヤーを取得、無ければ新規作成して返す / Get a layer by name, creating it if it does not exist */
    function getOrCreateLayer(doc, layerName) {
        var layer = findLayerByName(doc, layerName);
        if (!layer) {
            layer = doc.layers.add();
            layer.name = layerName;
        }
        return layer;
    }

    /* アイテムが属するレイヤーを返す（削除済みなら null）/ Return the layer owning the item, or null if the item is gone */
    function getOwnerLayer(pageItem) {
        return tryCall(function () { return pageItem.layer; }) || null;
    }

    /* 指定名のレイヤーを確実に削除（中身のロックを解除してから削除）/ Force-remove a layer by name (unlock its contents first, then remove) */
    function forceRemoveLayerByName(doc, layerName) {
        var layer = findLayerByName(doc, layerName);
        if (!layer) return;

        trySetProperty(layer, 'locked', false);
        trySetProperty(layer, 'visible', true);

        // ロックされた中身が削除を妨げるため、先にすべて解除 / locked contents block removal, so unlock them first
        for (var i = 0; i < layer.pageItems.length; i++) trySetProperty(layer.pageItems[i], 'locked', false);
        for (var j = 0; j < layer.layers.length; j++) trySetProperty(layer.layers[j], 'locked', false);

        tryCall(function () { layer.remove(); });
    }

    // =========================================
    // Undo / プレビュー管理 / Undo & Preview Manager
    // =========================================

    /* プレビュー編集をUndoステップとして積み、巻き戻し・確定を一括管理するクラス / Manages preview edits as undo steps for batch rollback or commit */
    function PreviewManager() {
        this.undoDepth = 0; // プレビュー中に実行したアクション数 / number of preview actions executed

        // 変更操作を実行し、実際に変更があった場合だけ1ステップとしてカウント / Run an action and count it only when it actually changes the document
        this.runAsStep = function (previewAction) {
            var self = this;
            runOrAlert("Preview Error", function () {
                var changed = (typeof previewAction === "function") ? previewAction() : false;
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

        // 確定：プレビュー分を全て取り消してから確定処理を1回実行 / Commit: undo all preview steps, then run the commit action once
        this.commit = function (commitAction) {
            this.rollback();
            if (typeof commitAction === "function") {
                runOrAlert("Commit Error", commitAction);
            }
        };
    }

    // =========================================
    // 選択・型判定 / Selection & Type Guards
    // =========================================

    /* オブジェクトが TextFrame かどうかを判定（削除済み参照は false）/ Return true if the object is a TextFrame (a deleted reference yields false) */
    function isTextFrame(target) {
        try {
            return !!target && target.typename === "TextFrame";
        } catch (e) {
            return false;
        }
    }

    /* 選択先頭が TextFrame ならそれを返す（無ければ null）/ Return the selected TextFrame, or null if none is selected */
    function getSelectedTextFrame() {
        if (app.documents.length === 0) return null;
        var currentSelection = app.selection;
        if (currentSelection && currentSelection.length > 0 && isTextFrame(currentSelection[0])) {
            return currentSelection[0];
        }
        return null;
    }

    // =========================================
    // _pagenumber レイヤーの状態管理 / Pagenumber Layer State
    // =========================================

    /* 同じ親の中でのレイヤーの重ね順インデックスを返す（無ければ -1）/ Return the stacking-order index of a layer among its siblings (or -1) */
    function getLayerStackIndex(layer) {
        var siblingLayers = layer.parent.layers;
        for (var i = 0; i < siblingLayers.length; i++) {
            if (siblingLayers[i] === layer) return i;
        }
        return -1;
    }

    /* _pagenumber レイヤーの現在状態（ロック・表示・所属・重ね順）を記録 / Capture the current state (lock, visibility, parent, stacking order) of the _pagenumber layer */
    function capturePagenumberState(pagenumberLayer, layerExisted) {
        // 親コンテナ（ドキュメントまたは親レイヤー）ごと覚えておく / remember the parent container (document or parent layer) as well
        var parentContainer = pagenumberLayer.parent;
        var stackIndex = getLayerStackIndex(pagenumberLayer);
        return {
            existed: !!layerExisted,
            locked: pagenumberLayer.locked,
            visible: pagenumberLayer.visible,
            parentContainer: parentContainer,
            // ひとつ上（前面側）のレイヤーそのものを復元の基準として記録（同名レイヤーがあっても取り違えない）
            // remember the neighbor layer above itself as a restore anchor, so duplicate layer names cannot confuse it
            neighborAbove: (stackIndex > 0) ? parentContainer.layers[stackIndex - 1] : null
        };
    }

    /* _pagenumber レイヤーを用意し、元状態を記録したうえで作業用に整える / Prepare the _pagenumber layer for work and capture its original state */
    function setupPagenumberLayer(doc) {
        var layerExisted = !!findLayerByName(doc, PAGENUMBER_LAYER_NAME);
        var pagenumberLayer = getOrCreateLayer(doc, PAGENUMBER_LAYER_NAME);
        var originalState = capturePagenumberState(pagenumberLayer, layerExisted);

        // ロック解除・表示・最前面へ / unlock, show, and move to the top
        trySetProperty(pagenumberLayer, 'locked', false);
        trySetProperty(pagenumberLayer, 'visible', true);
        tryCall(function () { pagenumberLayer.move(doc, ElementPlacement.PLACEATBEGINNING); });

        return { layer: pagenumberLayer, originalState: originalState };
    }

    /* capturePagenumberState で記録した状態へ _pagenumber レイヤーを復元 / Restore the _pagenumber layer to the captured state */
    function restorePagenumberState(doc, pagenumberLayer, originalState, removeWhenAutoCreated) {
        if (!pagenumberLayer || !originalState) return;

        // キャンセル時のみ、元々存在しなかった _pagenumber を削除 / remove an auto-created _pagenumber only on Cancel
        if (!originalState.existed && removeWhenAutoCreated) {
            forceRemoveLayerByName(doc, PAGENUMBER_LAYER_NAME);
            return;
        }

        // 所属と重ね順を復元 / restore the parent container and the stacking order
        if (originalState.neighborAbove) {
            tryCall(function () { pagenumberLayer.move(originalState.neighborAbove, ElementPlacement.PLACEAFTER); });
        } else {
            tryCall(function () {
                pagenumberLayer.move(originalState.parentContainer || doc, ElementPlacement.PLACEATBEGINNING);
            });
        }

        // 表示・ロック状態を復元 / restore visibility & lock
        trySetProperty(pagenumberLayer, 'visible', originalState.visible);
        trySetProperty(pagenumberLayer, 'locked', originalState.locked);
    }

    // =========================================
    // アートボードとフレームの探索 / Artboard & Frame Lookup
    // =========================================

    /* 座標 point が矩形 rect 内にあるか判定 / Return true if the point is inside the rectangle */
    function isPointInRect(point, rect) {
        return point[0] >= rect[0] && point[0] <= rect[2] && point[1] <= rect[1] && point[1] >= rect[3];
    }

    /* 指定座標が含まれるアートボードのインデックスを返す（無ければ -1）/ Return the index of the artboard containing the given point (or -1) */
    function getArtboardIndexByPosition(doc, point) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isPointInRect(point, doc.artboards[i].artboardRect)) return i;
        }
        return -1;
    }

    /* いずれかのアートボード上で最初に見つかった TextFrame を返す / Return the first TextFrame found on any artboard */
    function findTextFrameOnAnyArtboard(doc, targetLayer) {
        for (var i = 0; i < targetLayer.textFrames.length; i++) {
            var textFrame = targetLayer.textFrames[i];
            if (getArtboardIndexByPosition(doc, textFrame.position) >= 0) return textFrame;
        }
        return null;
    }

    /* TextFrame 群をアートボード順に並べた配列を返す（excludedFrame とアートボード外は除外）/ Return the TextFrames sorted by artboard order (excludedFrame and off-artboard frames are skipped) */
    function sortFramesByArtboard(doc, textFrames, excludedFrame) {
        var frameEntries = [];
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (excludedFrame && textFrame === excludedFrame) continue;
            var artboardIndex = getArtboardIndexByPosition(doc, textFrame.position);
            // どのアートボードにも乗らないテキストは採番対象外 / text that sits on no artboard is not numbered
            if (artboardIndex < 0) continue;
            frameEntries.push({ frame: textFrame, artboardIndex: artboardIndex });
        }
        frameEntries.sort(function (a, b) { return a.artboardIndex - b.artboardIndex; });

        var sortedFrames = [];
        for (var j = 0; j < frameEntries.length; j++) sortedFrames.push(frameEntries[j].frame);
        return sortedFrames;
    }

    // =========================================
    // ページ番号テキストの生成・配置 / Page Number Generation & Placement
    // =========================================

    /* 番号・接頭辞/接尾辞・ゼロ埋め・総ページ表示からページ番号文字列を生成 / Build the page-number string from the number, prefix/suffix, zero padding, and the optional total */
    function buildPageNumberText(pageNumber, digitCount, formatOptions, totalPages) {
        var numberText = String(pageNumber);
        if (formatOptions.zeroPad && numberText.length < digitCount) {
            numberText = Array(digitCount - numberText.length + 1).join("0") + numberText; // ES3対応ゼロ埋め / ES3-safe zero pad
        }
        var pageNumberText = formatOptions.prefix + numberText + formatOptions.suffix;
        if (formatOptions.showTotal) pageNumberText += "/" + totalPages;
        return pageNumberText;
    }

    /* レイヤー上のテキストをアートボード順に並べ、連番を流し込む / Sort the layer's text frames by artboard and write sequential numbers into them */
    function numberFramesInOrder(doc, targetLayer, excludedFrame, startNumber, formatOptions) {
        var sortedFrames = sortFramesByArtboard(doc, targetLayer.textFrames, excludedFrame);
        var lastPageNumber = startNumber + doc.artboards.length - 1;
        var digitCount = String(lastPageNumber).length;
        for (var i = 0; i < sortedFrames.length; i++) {
            trySetProperty(sortedFrames[i], 'contents',
                buildPageNumberText(startNumber + i, digitCount, formatOptions, lastPageNumber));
        }
    }

    /* 指定レイヤー上の TextFrame を keptFrame 以外すべて削除 / Remove every TextFrame on the layer except keptFrame */
    function removeOtherTextFrames(targetLayer, keptFrame) {
        var textFrames = targetLayer.textFrames;
        for (var i = textFrames.length - 1; i >= 0; i--) {
            var textFrame = textFrames[i];
            if (textFrame === keptFrame) continue;
            trySetProperty(textFrame, 'locked', false);
            tryCall(function () { textFrame.remove(); });
        }
    }

    /* 雛形テキストをカットし、全アートボードへ貼り付ける（プレビューと確定で共通）。成功したら true / Cut the given text and paste it onto every artboard (shared by preview and commit); returns true on success */
    function cutAndPasteToAllArtboards(doc, textFrame, pasteLayer, beforePaste) {
        // 対象と所属レイヤーを一時的にロック解除＆可視化 / temporarily unlock & show the target and its layer
        var sourceLayer = getOwnerLayer(textFrame);
        trySetProperty(textFrame, 'locked', false);
        trySetProperty(sourceLayer, 'locked', false);
        trySetProperty(sourceLayer, 'visible', true);

        // 対象が乗るアートボードをアクティブ化 / activate the artboard the target sits on
        var sourceArtboardIndex = getArtboardIndexByPosition(doc, textFrame.position);
        if (sourceArtboardIndex >= 0) doc.artboards.setActiveArtboardIndex(sourceArtboardIndex);

        // 選択→カット / select -> cut
        app.selection = null;
        trySetProperty(textFrame, 'selected', true);
        var cutSucceeded = tryCall(function () {
            app.cut();
            return true;
        }) === true;

        // カットできていない場合、この先へ進むと既存テキストを消すだけになるため中断
        // Bail out when the cut failed: continuing would only delete the existing text without pasting anything back
        if (!cutSucceeded) return false;

        // 貼り付け直前の後始末（既存テキストの一掃など）/ cleanup right before pasting (e.g. clearing existing text)
        if (beforePaste) beforePaste();

        // 貼り付け先レイヤーをアクティブにして全アートボードへ貼り付け / activate the destination layer, then paste onto all artboards
        var destinationLayer = pasteLayer || sourceLayer;
        if (destinationLayer) trySetProperty(doc, 'activeLayer', destinationLayer);
        return tryCall(function () {
            app.executeMenuCommand('pasteInAllArtboard');
            return true;
        }) === true;
    }

    /* 雛形テキストを開始番号で初期化し、全アートボードへ複製（所属レイヤーの状態は元へ戻す）/ Seed the template text with the start number and duplicate it across all artboards, restoring its layer state afterwards */
    function seedAndPasteToAllArtboards(doc, templateText, startNumber) {
        trySetProperty(templateText, 'contents', String(startNumber));

        // 所属レイヤーの一時状態を退避 / back up the source layer's state
        var sourceLayer = getOwnerLayer(templateText);
        var originalLocked = sourceLayer ? sourceLayer.locked : null;
        var originalVisible = sourceLayer ? sourceLayer.visible : null;

        var pasted = cutAndPasteToAllArtboards(doc, templateText, sourceLayer);

        // レイヤーの一時状態を元へ戻す / restore the layer's temporary state
        if (sourceLayer) {
            trySetProperty(sourceLayer, 'locked', originalLocked);
            trySetProperty(sourceLayer, 'visible', originalVisible);
        }
        return pasted;
    }

    // =========================================
    // ライブプレビュー / Live Preview
    // =========================================

    /* 雛形を退避レイヤーへ非表示コピーする（キャンセル時の復元用。常に最新の1つだけ保持）/ Copy the template onto a hidden backup layer for restoring on Cancel, keeping only the latest copy */
    function backupTemplateText(doc, templateText) {
        var backupLayer = getOrCreateLayer(doc, BACKUP_LAYER_NAME);
        backupLayer.visible = false;
        backupLayer.locked = false;

        // 前回の退避が残っていると、キャンセル時に雛形が重複して復元されるため先に破棄
        // Leftover backups would be restored on top of each other on Cancel, so discard them first
        for (var i = backupLayer.pageItems.length - 1; i >= 0; i--) {
            var staleItem = backupLayer.pageItems[i];
            trySetProperty(staleItem, 'locked', false);
            tryCall(function () { staleItem.remove(); });
        }

        var backupText = tryCall(function () {
            return templateText.duplicate(backupLayer, ElementPlacement.PLACEATBEGINNING);
        });
        trySetProperty(backupText, 'visible', false);
        trySetProperty(backupText, 'locked', true);
    }

    /* 雛形を退避しつつ、全アートボードへクリーンに複製し直す / Back up the template, then cleanly re-duplicate it across every artboard */
    function rebuildFramesAcrossArtboards(doc, pagenumberLayer, templateText) {
        backupTemplateText(doc, templateText);
        var pasted = cutAndPasteToAllArtboards(doc, templateText, pagenumberLayer, function () {
            // 貼り付け前に既存のページ番号を一掃 / clear the existing page numbers before pasting
            removeOtherTextFrames(pagenumberLayer, null);
        });

        // 失敗時は退避レイヤーごと破棄して、中途半端な状態を残さない / on failure, drop the backup layer so no half-finished state remains
        if (!pasted) forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);
        return pasted;
    }

    /* 選択テキスト（無ければレイヤー上の先頭テキスト）を雛形に、全アートボードへ連番をプレビュー / Render a sequential-numbering preview on every artboard, using the selected text (or the first text on the layer) as a template */
    function updatePreview(doc, layerName, startNumber, formatOptions) {
        if (!doc || isNaN(startNumber)) return false;
        var pagenumberLayer = findLayerByName(doc, layerName);
        if (!pagenumberLayer) return false;

        // 雛形を決定：選択 → アートボード上の先頭テキスト → レイヤー上の先頭テキスト
        // pick the template: selection -> first text on an artboard -> first text on the layer
        var templateText = getSelectedTextFrame() ||
            sortFramesByArtboard(doc, pagenumberLayer.textFrames, null)[0] ||
            pagenumberLayer.textFrames[0];
        if (!templateText) return false;

        if (!rebuildFramesAcrossArtboards(doc, pagenumberLayer, templateText)) return false;
        numberFramesInOrder(doc, pagenumberLayer, null, startNumber, formatOptions);

        // 再描画は呼び出し元（PreviewManager）で1回だけ行い、Undoの区切りを1プレビュー＝1ステップに保つ
        // The caller (PreviewManager) redraws once, keeping the undo boundary at one step per preview pass
        return true;
    }

    /* 退避レイヤーに残ったテキストを _pagenumber へ戻し、退避レイヤーを削除 / Move any text left on the backup layer back to _pagenumber, then remove the backup layer */
    function restorePreviewBackupOnCancel(doc) {
        var backupLayer = findLayerByName(doc, BACKUP_LAYER_NAME);
        if (!backupLayer) return;

        var pagenumberLayer = getOrCreateLayer(doc, PAGENUMBER_LAYER_NAME);
        removeOtherTextFrames(pagenumberLayer, null);
        backupLayer.locked = false;
        backupLayer.visible = true;

        for (var i = backupLayer.pageItems.length - 1; i >= 0; i--) {
            var backupItem = backupLayer.pageItems[i];
            trySetProperty(backupItem, 'locked', false);
            tryCall(function () { backupItem.move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING); });
            trySetProperty(backupItem, 'visible', true);
        }
        forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);
        safeRedraw();
    }

    // =========================================
    // キーボード操作 / Keyboard Handlers
    // =========================================

    /* 数値入力欄で↑↓キーによる増減を有効化（Shiftで10の倍数へスナップ）/ Enable Up/Down arrow increment-decrement on a number field (Shift snaps to multiples of 10) */
    function changeValueByArrowKey(numberField, onChanged) {
        if (!numberField || !numberField.addEventListener) return;
        numberField.addEventListener("keydown", function (event) {
            var currentValue = Number(numberField.text);
            if (isNaN(currentValue)) return;

            var keyName = event.keyName;
            var isUp = (keyName === "Up" || keyName === "UpArrow");
            var isDown = (keyName === "Down" || keyName === "DownArrow");
            if (!isUp && !isDown) return;

            if (ScriptUI.environment.keyboardState.shiftKey) {
                // 10の倍数へスナップ / snap to a multiple of 10
                currentValue = isUp ? Math.ceil((currentValue + 1) / 10) * 10 : Math.floor((currentValue - 1) / 10) * 10;
            } else {
                currentValue += isUp ? 1 : -1;
            }
            if (currentValue < 0) currentValue = 0;

            numberField.text = Math.round(currentValue);
            event.preventDefault();
            if (onChanged) onChanged();
        });
    }

    /* 指定キー押下でチェックボックスをトグルするハンドラを登録 / Register a handler that toggles a checkbox when the given key is pressed */
    function addToggleKeyHandler(targetDialog, toggleKey, checkbox, onChanged, skipWhenEditTextFocus) {
        if (!targetDialog || !checkbox || !targetDialog.addEventListener) return;
        targetDialog.addEventListener("keydown", function (event) {
            // 入力欄フォーカス中はスキップしたい場合のみスキップ / skip while an edittext is focused, only when requested
            if (skipWhenEditTextFocus && event.target && event.target.type === "edittext") return;
            if ((event.keyName || "").toUpperCase() !== String(toggleKey).toUpperCase()) return;

            checkbox.value = !checkbox.value;
            if (onChanged) onChanged();
            event.preventDefault();
        });
    }

    // =========================================
    // UI 構築 / UI Construction
    // =========================================

    /* 親に縦並びカラム（group）を追加して返す / Add a vertical column group to the parent and return it */
    function addColumnGroup(parentGroup, childAlignment) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = childAlignment || "left";
        return columnGroup;
    }

    /* ラベル付き入力欄を追加し、入力欄（edittext）を返す / Add a labeled edittext and return the edittext */
    function addLabeledEditText(parentGroup, labelText, initialValue, characterWidth, tooltipText) {
        var captionLabel = parentGroup.add("statictext", undefined, labelText);
        var inputField = parentGroup.add("edittext", undefined, initialValue);
        inputField.characters = characterWidth;
        // ラベル・入力欄のどちらにマウスを乗せても説明が出るようにする / show the hint from both the caption and the field
        captionLabel.helpTip = tooltipText;
        inputField.helpTip = tooltipText;
        return inputField;
    }

    /* チェックボックスを追加して返す / Add a checkbox and return it */
    function addCheckbox(parentGroup, labelText, tooltipText) {
        var checkbox = parentGroup.add("checkbox", undefined, labelText);
        checkbox.helpTip = tooltipText;
        return checkbox;
    }

    /* 右寄せのボタン行（キャンセル → OK）を追加して返す / Add the right-aligned button row (Cancel then OK) and return both buttons */
    function addButtonRow(targetDialog) {
        var buttonRow = targetDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignChildren = ["right", "center"];
        // 右マージンは0（ダイアログ端にボタンを寄せる）/ no right margin, so the buttons sit flush with the dialog edge
        buttonRow.margins = [10, 10, 0, 0];
        buttonRow.alignment = ["right", "bottom"];

        var cancelButton = buttonRow.add("button", undefined, LABELS.button.cancel[currentLanguage], { name: "cancel" });
        cancelButton.helpTip = LABELS.tooltip.cancel[currentLanguage];

        var okButton = buttonRow.add("button", undefined, "OK", { name: "ok" });
        okButton.helpTip = LABELS.tooltip.ok[currentLanguage];

        return { cancelButton: cancelButton, okButton: okButton };
    }

    /* ダイアログと各UIコントロールを生成し、参照をまとめて返す / Build the dialog and its controls, returning all references */
    function buildDialog() {
        // タイトルバーにはバージョンを併記 / show the version in the title bar
        var dialog = new Window("dialog", LABELS.dialog.title[currentLanguage] + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "left";

        // 3カラムレイアウト / 3-column layout
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";

        // 左カラム: 接頭辞 / left column: prefix
        var prefixColumn = addColumnGroup(columnsGroup);
        var prefixField = addLabeledEditText(prefixColumn, LABELS.field.prefix[currentLanguage], "", 10,
            LABELS.tooltip.prefix[currentLanguage]);

        // 中央カラム: 開始番号 + ゼロ埋め / center column: start number + zero pad
        var startNumberColumn = addColumnGroup(columnsGroup);
        var startNumberField = addLabeledEditText(startNumberColumn, LABELS.field.start[currentLanguage], "1", 6,
            LABELS.tooltip.start[currentLanguage]);
        var zeroPadCheckbox = addCheckbox(startNumberColumn, LABELS.checkbox.zeroPad[currentLanguage],
            LABELS.tooltip.zeroPad[currentLanguage]);

        // 右カラム: 接尾辞 + 総ページ表示 / right column: suffix + show-total
        var suffixColumn = addColumnGroup(columnsGroup);
        var suffixField = addLabeledEditText(suffixColumn, LABELS.field.suffix[currentLanguage], "", 10,
            LABELS.tooltip.suffix[currentLanguage]);
        var totalPageCheckbox = addCheckbox(suffixColumn, LABELS.checkbox.showTotal[currentLanguage],
            LABELS.tooltip.showTotal[currentLanguage]);

        var dialogButtons = addButtonRow(dialog);

        // 透明度と表示位置の調整 / adjust opacity and position
        dialog.opacity = 0.98;
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
            cancelButton: dialogButtons.cancelButton,
            okButton: dialogButtons.okButton
        };
    }

    // =========================================
    // 確定処理 / Commit
    // =========================================

    /* OK確定時の雛形テキストを取得（優先候補 → 現在の選択 → レイヤー上の既存テキスト）/ Resolve the template text for the commit (preferred -> current selection -> existing text on the layer) */
    function resolveTemplateTextForCommit(doc, pagenumberLayer, preferredText) {
        var templateText = isTextFrame(preferredText) ? preferredText : getSelectedTextFrame();
        if (!isTextFrame(templateText)) {
            templateText = findTextFrameOnAnyArtboard(doc, pagenumberLayer);
        }
        return isTextFrame(templateText) ? templateText : null;
    }

    /* 雛形テキストを _pagenumber 上へ移し、他のテキストを除去 / Move the template to _pagenumber and remove the other text frames */
    function moveTemplateTextToPagenumberLayer(pagenumberLayer, templateText) {
        if (!isTextFrame(templateText)) return false;
        if (templateText.layer.name !== PAGENUMBER_LAYER_NAME) {
            tryCall(function () { templateText.move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING); });
        }
        if (templateText.layer.name !== PAGENUMBER_LAYER_NAME) return false;
        removeOtherTextFrames(pagenumberLayer, templateText);
        return true;
    }

    /* 確定用の連番を全アートボードへ適用 / Apply the committed sequential page numbers to every artboard */
    function applyNumberingToAllArtboards(doc, pagenumberLayer, templateText, startNumber, formatOptions) {
        // 複製に失敗した場合は採番せず、雛形をそのまま残す / when the duplication fails, leave the template as it is
        if (!seedAndPasteToAllArtboards(doc, templateText, startNumber)) return false;
        numberFramesInOrder(doc, pagenumberLayer, templateText, startNumber, formatOptions);
        return true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* ダイアログを表示し、選択テキストを雛形に全アートボードへページ番号を配置 / Show the dialog and place page numbers on every artboard using the selected text as a template */
    function main() {
        // テキスト未選択なら終了 / Exit if no text is selected
        var originalTemplateText = getSelectedTextFrame();
        if (!originalTemplateText) {
            alert(LABELS.alert.invalidSelection[currentLanguage]);
            return;
        }

        var doc = app.activeDocument;
        var pagenumberSetup = setupPagenumberLayer(doc);
        var dialogUI = buildDialog();
        var previewManager = new PreviewManager();

        /* 現在の入力値を書式オプションとしてまとめる / Collect the current input values as formatting options */
        function getFormatOptions() {
            return {
                prefix: dialogUI.prefixField.text || "",
                suffix: dialogUI.suffixField.text || "",
                zeroPad: !!dialogUI.zeroPadCheckbox.value,
                showTotal: !!dialogUI.totalPageCheckbox.value
            };
        }

        /* 現在の入力値でライブプレビューを更新（前回分を巻き戻し、1ステップとして再実行）/ Refresh the live preview with current input values (roll back the previous one, run as a single step) */
        function refreshPreview() {
            var startNumber = parseInt(dialogUI.startNumberField.text, 10);
            if (isNaN(startNumber)) return;

            var formatOptions = getFormatOptions();
            previewManager.rollback();
            previewManager.runAsStep(function () {
                return updatePreview(doc, PAGENUMBER_LAYER_NAME, startNumber, formatOptions);
            });
        }

        /* OK確定時の本処理：雛形テキストを _pagenumber へ移し、全アートボードへ連番を確定配置 / Commit: move the template text to _pagenumber and place sequential numbers on every artboard */
        function commitPageNumbers() {
            var startNumber = parseInt(dialogUI.startNumberField.text, 10);
            if (isNaN(startNumber)) {
                alert(LABELS.alert.notNumber[currentLanguage]);
                return;
            }

            // プレビュー用の退避レイヤーを破棄 / discard the preview backup layer
            forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);

            var pagenumberLayer = getOrCreateLayer(doc, PAGENUMBER_LAYER_NAME);
            var templateText = resolveTemplateTextForCommit(doc, pagenumberLayer, originalTemplateText);
            if (!templateText || !moveTemplateTextToPagenumberLayer(pagenumberLayer, templateText)) {
                alert(LABELS.alert.invalidSelection[currentLanguage]);
                return;
            }

            var placed = applyNumberingToAllArtboards(doc, pagenumberLayer, templateText, startNumber, getFormatOptions());
            safeRedraw();
            restorePagenumberState(doc, pagenumberSetup.layer, pagenumberSetup.originalState, false);
            if (!placed) alert(LABELS.alert.commitFailed[currentLanguage]);
        }

        changeValueByArrowKey(dialogUI.startNumberField, refreshPreview);

        // 入力確定（Tabやフォーカス移動）でプレビューを更新。onChanging は1文字ごとに全アートボードを組み直すため使わない
        // Refresh on commit of the field (Tab or focus change); onChanging would rebuild every artboard on each keystroke
        dialogUI.prefixField.onChange = refreshPreview;
        dialogUI.suffixField.onChange = refreshPreview;
        dialogUI.startNumberField.onChange = refreshPreview;

        // チェックボックスのON/OFFでプレビューを更新（キー操作での切り替えは値を直接書き換えるため onClick は発火しない）
        // Update the preview when a checkbox is toggled (key shortcuts set .value directly, so onClick does not fire for them)
        dialogUI.zeroPadCheckbox.onClick = refreshPreview;
        dialogUI.totalPageCheckbox.onClick = refreshPreview;

        // Zキーでゼロ埋め、Aキーで総ページ表示をトグル（入力欄では文字入力を優先）
        // Z toggles zero-pad, A toggles show-total; typing in a text field takes precedence
        addToggleKeyHandler(dialogUI.dialog, "Z", dialogUI.zeroPadCheckbox, refreshPreview, true);
        addToggleKeyHandler(dialogUI.dialog, "A", dialogUI.totalPageCheckbox, refreshPreview, true);

        // OK：プレビュー分を全Undoしてから確定処理を1回だけ実行 / OK: undo all preview steps, then run the commit action once
        dialogUI.okButton.onClick = function () {
            previewManager.commit(commitPageNumbers);
            forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);
            dialogUI.dialog.close(1);
        };

        // キャンセル：プレビューを巻き戻し、退避テキストと _pagenumber 状態を復元 / Cancel: roll back the preview, restore the backed-up text and the _pagenumber state
        dialogUI.cancelButton.onClick = function () {
            previewManager.rollback();
            restorePreviewBackupOnCancel(doc);
            restorePagenumberState(doc, pagenumberSetup.layer, pagenumberSetup.originalState, true);
            dialogUI.dialog.close(0);
        };

        dialogUI.startNumberField.active = true;

        // 初回プレビュー / first preview pass
        refreshPreview();

        dialogUI.dialog.show();
    }

    main();

})();
