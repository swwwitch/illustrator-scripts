#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のポイントテキストを雛形に、すべてのアートボードへページ番号を配置します。
接頭辞・接尾辞・ゼロ埋め・総ページ数表示に対応し、変更は即時プレビューされます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddPageNumberFromTextSelection.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ndc3d96ffc335

### Overview

Uses the selected point text as a template and places a page number on every artboard.
Prefix, suffix, zero padding and a total-pages display are supported, with an immediate preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddPageNumberFromTextSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddPageNumberFromTextSelection"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-06-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddPageNumberFromTextSelection.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddPageNumberFromTextSelection.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ndc3d96ffc335"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    // 連番テキストを配置する対象レイヤー名 / Layer that receives the page-number text
    var PAGENUMBER_LAYER_NAME = "_pagenumber";
    // プレビュー中の雛形を退避する一時レイヤー名 / Temp layer used to back up the template during preview
    var BACKUP_LAYER_NAME = "_pagenumber_preview";

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OFFSET_X = 300;              /* ダイアログの表示位置：右(+)／左(-) / dialog offset to the right */
    var DIALOG_OPACITY = 0.98;              /* ダイアログの不透明度 / dialog opacity */
    var AFFIX_FIELD_CHARS = 10;             /* 接頭辞・接尾辞欄の文字数 / width of the prefix and suffix fields */
    var START_NUMBER_FIELD_CHARS = 6;       /* 開始番号欄の文字数 / width of the start number field */
    var BUTTON_ROW_MARGINS = [10, 10, 0, 0];  /* ボタン行の余白（右0でダイアログ端に寄せる）/ button row margins, flush right */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のUI言語を判定（日本語環境は "ja"、その他は "en"）/ Detect the environment's UI language ("ja" for Japanese, otherwise "en")
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    // UI 文字列（OK ボタンのラベルは非ローカライズ）/ UI strings (the OK button label is not localized)
    var LABELS = {
        dialog: {
            title: { ja: "ページ番号を一括配置", en: "Place Page Numbers" }
        },
        fieldLabel: {
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

    /**
     * LABELS からドット区切りのパスで現在の言語の文字列を取り出す
     * @param {string} labelPath - "fieldLabel.prefix" のようなパス
     * @returns {string} 現在の言語の文字列
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
        }
        return labelNode[uiLang];
    }

    // =========================================
    // 安全実行ヘルパー / Safe Execution Helpers
    // =========================================

    /**
     * 関数を try/catch 内で実行し、action の戻り値を返す。例外時は onError(e) を呼ぶ（省略時は無視）。
     * 例外を握りつぶしてよい処理の共通ヘルパー / Shared helper for operations where ignored exceptions are acceptable
     * @param {Function} action - 実行する関数
     * @param {Function} [onError] - 例外時に呼ぶ関数
     * @returns {*} action の戻り値（例外時は undefined）
     */
    function tryCall(action, onError) {
        try {
            return action ? action() : undefined;
        } catch (e) {
            if (onError) onError(e);
        }
    }

    /**
     * プロパティ代入を例外無視で実行（ロック中・削除済みオブジェクトは代入で例外を出すため）/ Assign a property, ignoring any error (locked or deleted objects throw on assignment)
     * @param {object} targetObject - 代入先（null なら何もしない）
     * @param {string} propertyName - プロパティ名
     * @param {*} propertyValue - 代入する値
     * @returns {void}
     */
    function trySetProperty(targetObject, propertyName, propertyValue) {
        tryCall(function () { if (targetObject) targetObject[propertyName] = propertyValue; });
    }

    /**
     * 関数を実行し、例外時は errorLabel 付きでアラート表示
     * Run a function; on error show an alert prefixed with errorLabel
     * @param {string} errorLabel - アラートの先頭に付ける文字列
     * @param {Function} action - 実行する関数
     * @returns {void}
     */
    function runOrAlert(errorLabel, action) {
        tryCall(action, function (e) { alert(errorLabel + ": " + e); });
    }

    /**
     * 画面を安全に再描画
     * Redraw the screen safely
     * @returns {void}
     */
    function safeRedraw() {
        tryCall(function () { app.redraw(); });
    }

    // =========================================
    // レイヤー操作 / Layer Operations
    // =========================================

    /**
     * 指定名のレイヤーをサブレイヤーまで含めて探す（無ければ null）/ Find a layer by name, including sub-layers (null when it does not exist)
     * @param {Document|Layer} parentContainer - 探す範囲
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー
     */
    function findLayerByName(parentContainer, layerName) {
        for (var i = 0; i < parentContainer.layers.length; i++) {
            var childLayer = parentContainer.layers[i];
            if (childLayer.name === layerName) return childLayer;
            // サブレイヤーに同名があっても取りこぼさない / do not miss a nested layer with the same name
            var nestedLayer = findLayerByName(childLayer, layerName);
            if (nestedLayer) return nestedLayer;
        }
        return null;
    }

    /**
     * 指定名のレイヤーを取得、無ければ新規作成して返す
     * Get a layer by name, creating it if it does not exist
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 既存または新規のレイヤー
     */
    function getOrCreateLayer(doc, layerName) {
        var targetLayer = findLayerByName(doc, layerName);
        if (!targetLayer) {
            targetLayer = doc.layers.add();
            targetLayer.name = layerName;
        }
        return targetLayer;
    }

    /**
     * アイテムが属するレイヤーを返す（削除済みなら null）/ Return the layer owning the item, or null if the item is gone
     * @param {PageItem} pageItem - 対象のアイテム
     * @returns {Layer|null} 所属レイヤー
     */
    function getOwnerLayer(pageItem) {
        return tryCall(function () { return pageItem.layer; }) || null;
    }

    /**
     * 指定名のレイヤーを確実に削除（中身のロックを解除してから削除）/ Force-remove a layer by name (unlock its contents first, then remove)
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function forceRemoveLayerByName(doc, layerName) {
        var targetLayer = findLayerByName(doc, layerName);
        if (!targetLayer) return;

        trySetProperty(targetLayer, 'locked', false);
        trySetProperty(targetLayer, 'visible', true);

        // ロックされた中身が削除を妨げるため、先にすべて解除 / locked contents block removal, so unlock them first
        for (var i = 0; i < targetLayer.pageItems.length; i++) trySetProperty(targetLayer.pageItems[i], 'locked', false);
        for (var j = 0; j < targetLayer.layers.length; j++) trySetProperty(targetLayer.layers[j], 'locked', false);

        tryCall(function () { targetLayer.remove(); });
    }

    // =========================================
    // Undo / プレビュー管理 / Undo & Preview Manager
    // =========================================

    /**
     * プレビュー編集をUndoステップとして積み、巻き戻し・確定を一括管理するクラス
     * Manages preview edits as undo steps for batch rollback or commit
     * @constructor
     */
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

    /**
     * オブジェクトが TextFrame かどうかを判定（削除済み参照は false）/ Return true if the object is a TextFrame (a deleted reference yields false)
     * @param {*} candidate - 判定する値
     * @returns {boolean} TextFrame なら true
     */
    function isTextFrame(candidate) {
        try {
            return !!candidate && candidate.typename === "TextFrame";
        } catch (e) {
            return false;
        }
    }

    /**
     * 選択先頭が TextFrame ならそれを返す（無ければ null）/ Return the selected TextFrame, or null if none is selected
     * @returns {TextFrame|null} 選択中のテキスト
     */
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

    /**
     * 同じ親の中でのレイヤーの重ね順インデックスを返す（無ければ -1）/ Return the stacking-order index of a layer among its siblings (or -1)
     * @param {Layer} targetLayer - 対象のレイヤー
     * @returns {number} 重ね順のインデックス
     */
    function getLayerStackIndex(targetLayer) {
        var siblingLayers = targetLayer.parent.layers;
        for (var i = 0; i < siblingLayers.length; i++) {
            if (siblingLayers[i] === targetLayer) return i;
        }
        return -1;
    }

    /**
     * _pagenumber レイヤーの現在状態（ロック・表示・所属・重ね順）を記録
     * Capture the current state (lock, visibility, parent, stacking order) of the _pagenumber layer
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {boolean} layerExisted - 実行前から存在したか
     * @returns {object} existed / locked / visible / parentContainer / neighborAbove
     */
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

    /**
     * _pagenumber レイヤーを用意し、元状態を記録したうえで作業用に整える
     * Prepare the _pagenumber layer for work and capture its original state
     * @param {Document} doc - 対象ドキュメント
     * @returns {{layer: Layer, originalState: object}} レイヤーと元の状態
     */
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

    /**
     * capturePagenumberState で記録した状態へ _pagenumber レイヤーを復元
     * Restore the _pagenumber layer to the captured state
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {object} originalState - capturePagenumberState() の戻り値
     * @param {boolean} removeWhenAutoCreated - 自動作成したレイヤーなら削除する（キャンセル時）
     * @returns {void}
     */
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

    /**
     * 座標 point が矩形 rect 内にあるか判定
     * Return true if the point is inside the rectangle
     * @param {number[]} point - 座標 [x, y]
     * @param {number[]} rect - 矩形 [左, 上, 右, 下]
     * @returns {boolean} 内側なら true
     */
    function isPointInRect(point, rect) {
        return point[0] >= rect[0] && point[0] <= rect[2] && point[1] <= rect[1] && point[1] >= rect[3];
    }

    /**
     * 指定座標が含まれるアートボードのインデックスを返す（無ければ -1）/ Return the index of the artboard containing the given point (or -1)
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} point - 座標 [x, y]
     * @returns {number} アートボードのインデックス
     */
    function getArtboardIndexByPosition(doc, point) {
        for (var i = 0; i < doc.artboards.length; i++) {
            if (isPointInRect(point, doc.artboards[i].artboardRect)) return i;
        }
        return -1;
    }

    /**
     * いずれかのアートボード上で最初に見つかった TextFrame を返す
     * Return the first TextFrame found on any artboard
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} targetLayer - 探すレイヤー
     * @returns {TextFrame|null} 見つかったテキスト
     */
    function findTextFrameOnAnyArtboard(doc, targetLayer) {
        for (var i = 0; i < targetLayer.textFrames.length; i++) {
            var textFrame = targetLayer.textFrames[i];
            if (getArtboardIndexByPosition(doc, textFrame.position) >= 0) return textFrame;
        }
        return null;
    }

    /**
     * TextFrame 群をアートボード順に並べた配列を返す（excludedFrame とアートボード外は除外）/ Return the TextFrames sorted by artboard order (excludedFrame and off-artboard frames are skipped)
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrames|TextFrame[]} textFrames - 並べるテキスト
     * @param {TextFrame|null} excludedFrame - 除外するテキスト
     * @returns {TextFrame[]} アートボード順のテキスト
     */
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
        frameEntries.sort(function (entryA, entryB) { return entryA.artboardIndex - entryB.artboardIndex; });

        var sortedFrames = [];
        for (var j = 0; j < frameEntries.length; j++) sortedFrames.push(frameEntries[j].frame);
        return sortedFrames;
    }

    // =========================================
    // ページ番号テキストの生成・配置 / Page Number Generation & Placement
    // =========================================

    /**
     * 番号・接頭辞/接尾辞・ゼロ埋め・総ページ表示からページ番号文字列を生成
     * Build the page-number string from the number, prefix/suffix, zero padding, and the optional total
     * @param {number} pageNumber - 番号
     * @param {number} digitCount - ゼロ埋めの桁数
     * @param {object} formatOptions - prefix / suffix / zeroPad / showTotal
     * @param {number} totalPages - 総ページ数として表示する値
     * @returns {string} ページ番号の文字列
     */
    function buildPageNumberText(pageNumber, digitCount, formatOptions, totalPages) {
        var numberText = String(pageNumber);
        if (formatOptions.zeroPad && numberText.length < digitCount) {
            numberText = Array(digitCount - numberText.length + 1).join("0") + numberText; // ES3対応ゼロ埋め / ES3-safe zero pad
        }
        var pageNumberText = formatOptions.prefix + numberText + formatOptions.suffix;
        if (formatOptions.showTotal) pageNumberText += "/" + totalPages;
        return pageNumberText;
    }

    /**
     * レイヤー上のテキストをアートボード順に並べ、連番を流し込む
     * Sort the layer's text frames by artboard and write sequential numbers into them
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} targetLayer - テキストのあるレイヤー
     * @param {TextFrame|null} excludedFrame - 採番しないテキスト
     * @param {number} startNumber - 開始番号
     * @param {object} formatOptions - prefix / suffix / zeroPad / showTotal
     * @returns {void}
     */
    function numberFramesInOrder(doc, targetLayer, excludedFrame, startNumber, formatOptions) {
        var sortedFrames = sortFramesByArtboard(doc, targetLayer.textFrames, excludedFrame);
        var lastPageNumber = startNumber + doc.artboards.length - 1;
        var digitCount = String(lastPageNumber).length;
        for (var i = 0; i < sortedFrames.length; i++) {
            trySetProperty(sortedFrames[i], 'contents',
                buildPageNumberText(startNumber + i, digitCount, formatOptions, lastPageNumber));
        }
    }

    /**
     * 指定レイヤー上の TextFrame を keptFrame 以外すべて削除
     * Remove every TextFrame on the layer except keptFrame
     * @param {Layer} targetLayer - 対象のレイヤー
     * @param {TextFrame|null} keptFrame - 残すテキスト
     * @returns {void}
     */
    function removeOtherTextFrames(targetLayer, keptFrame) {
        var textFrames = targetLayer.textFrames;
        for (var i = textFrames.length - 1; i >= 0; i--) {
            var textFrame = textFrames[i];
            if (textFrame === keptFrame) continue;
            trySetProperty(textFrame, 'locked', false);
            tryCall(function () { textFrame.remove(); });
        }
    }

    /**
     * 雛形テキストをカットし、全アートボードへ貼り付ける（プレビューと確定で共通）。成功したら true
     * Cut the given text and paste it onto every artboard (shared by preview and commit); returns true on success
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} textFrame - 雛形テキスト
     * @param {Layer} pasteLayer - 貼り付け先レイヤー（null なら元のレイヤー）
     * @param {Function} [beforePaste] - 貼り付け直前に呼ぶ関数
     * @returns {boolean} 成功したら true
     */
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

    /**
     * 雛形テキストを開始番号で初期化し、全アートボードへ複製（所属レイヤーの状態は元へ戻す）/ Seed the template text with the start number and duplicate it across all artboards, restoring its layer state afterwards
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} templateText - 雛形テキスト
     * @param {number} startNumber - 開始番号
     * @returns {boolean} 成功したら true
     */
    function seedAndPasteToAllArtboards(doc, templateText, startNumber) {
        trySetProperty(templateText, 'contents', String(startNumber));

        // 所属レイヤーの一時状態を退避 / back up the source layer's state
        var sourceLayer = getOwnerLayer(templateText);
        var originalLocked = sourceLayer ? sourceLayer.locked : null;
        var originalVisible = sourceLayer ? sourceLayer.visible : null;

        var pasteSucceeded = cutAndPasteToAllArtboards(doc, templateText, sourceLayer);

        // レイヤーの一時状態を元へ戻す / restore the layer's temporary state
        if (sourceLayer) {
            trySetProperty(sourceLayer, 'locked', originalLocked);
            trySetProperty(sourceLayer, 'visible', originalVisible);
        }
        return pasteSucceeded;
    }

    // =========================================
    // ライブプレビュー / Live Preview
    // =========================================

    /**
     * 雛形を退避レイヤーへ非表示コピーする（キャンセル時の復元用。常に最新の1つだけ保持）/ Copy the template onto a hidden backup layer for restoring on Cancel, keeping only the latest copy
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame} templateText - 雛形テキスト
     * @returns {void}
     */
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

    /**
     * 雛形を退避しつつ、全アートボードへクリーンに複製し直す
     * Back up the template, then cleanly re-duplicate it across every artboard
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {TextFrame} templateText - 雛形テキスト
     * @returns {boolean} 成功したら true
     */
    function rebuildFramesAcrossArtboards(doc, pagenumberLayer, templateText) {
        backupTemplateText(doc, templateText);
        var pasteSucceeded = cutAndPasteToAllArtboards(doc, templateText, pagenumberLayer, function () {
            // 貼り付け前に既存のページ番号を一掃 / clear the existing page numbers before pasting
            removeOtherTextFrames(pagenumberLayer, null);
        });

        // 失敗時は退避レイヤーごと破棄して、中途半端な状態を残さない / on failure, drop the backup layer so no half-finished state remains
        if (!pasteSucceeded) forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);
        return pasteSucceeded;
    }

    /**
     * 選択テキスト（無ければレイヤー上の先頭テキスト）を雛形に、全アートボードへ連番をプレビュー
     * Render a sequential-numbering preview on every artboard, using the selected text (or the first text on the layer) as a template
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - ページ番号のレイヤー名
     * @param {number} startNumber - 開始番号
     * @param {object} formatOptions - prefix / suffix / zeroPad / showTotal
     * @returns {boolean} ドキュメントを変更したら true
     */
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

    /**
     * 退避レイヤーに残ったテキストを _pagenumber へ戻し、退避レイヤーを削除
     * Move any text left on the backup layer back to _pagenumber, then remove the backup layer
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
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

    /**
     * 数値入力欄で↑↓キーによる増減を有効化（Shiftで10の倍数へスナップ）/ Enable Up/Down arrow increment-decrement on a number field (Shift snaps to multiples of 10)
     * @param {EditText} numberField - 数値入力欄
     * @param {Function} [onChanged] - 値を変えたあとに呼ぶ関数
     * @returns {void}
     */
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

    /**
     * 指定キー押下でチェックボックスをトグルするハンドラを登録
     * Register a handler that toggles a checkbox when the given key is pressed
     * @param {Window} targetDialog - キー入力を受けるダイアログ
     * @param {string} toggleKey - 切り替えキー
     * @param {Checkbox} targetCheckbox - 切り替えるチェックボックス
     * @param {Function} onChanged - 切り替えたあとに呼ぶ関数
     * @param {boolean} skipWhenEditTextFocus - 入力欄にフォーカスがあるときは無視する
     * @returns {void}
     */
    function addToggleKeyHandler(targetDialog, toggleKey, targetCheckbox, onChanged, skipWhenEditTextFocus) {
        if (!targetDialog || !targetCheckbox || !targetDialog.addEventListener) return;
        targetDialog.addEventListener("keydown", function (event) {
            // 入力欄フォーカス中はスキップしたい場合のみスキップ / skip while an edittext is focused, only when requested
            if (skipWhenEditTextFocus && event.target && event.target.type === "edittext") return;
            if ((event.keyName || "").toUpperCase() !== String(toggleKey).toUpperCase()) return;

            targetCheckbox.value = !targetCheckbox.value;
            if (onChanged) onChanged();
            event.preventDefault();
        });
    }

    // =========================================
    // UI 構築 / UI Construction
    // =========================================

    /**
     * 親に縦並びカラム（group）を追加して返す
     * @param {Group} parentGroup - 追加先
     * @param {string} [childAlignment] - 子の揃え（省略時は "left"）
     * @returns {Group} 追加したカラム
     */
    function addColumnGroup(parentGroup, childAlignment) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = childAlignment || "left";
        return columnGroup;
    }

    /**
     * ラベル付き入力欄を追加し、入力欄（edittext）を返す
     * @param {Group} parentGroup - 追加先
     * @param {string} captionText - 入力欄の上に出す項目名
     * @param {string} initialValue - 入力欄の初期値
     * @param {number} characterWidth - 入力欄の文字数
     * @param {string} tooltipText - 項目名と入力欄に付ける tooltip
     * @returns {EditText} 追加した入力欄
     */
    function addLabeledEditText(parentGroup, captionText, initialValue, characterWidth, tooltipText) {
        var captionLabel = parentGroup.add("statictext", undefined, captionText);
        var inputField = parentGroup.add("edittext", undefined, initialValue);
        inputField.characters = characterWidth;
        // ラベル・入力欄のどちらにマウスを乗せても説明が出るようにする / show the hint from both the caption and the field
        captionLabel.helpTip = tooltipText;
        inputField.helpTip = tooltipText;
        return inputField;
    }

    /**
     * tooltip 付きのチェックボックスを追加して返す
     * @param {Group} parentGroup - 追加先
     * @param {string} checkboxText - チェックボックスの文言
     * @param {string} tooltipText - tooltip
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentGroup, checkboxText, tooltipText) {
        var optionCheckbox = parentGroup.add("checkbox", undefined, checkboxText);
        optionCheckbox.helpTip = tooltipText;
        return optionCheckbox;
    }

    /**
     * 右寄せのボタン行（キャンセル → OK）を追加して返す
     * @param {Window} targetDialog - 追加先のダイアログ
     * @returns {{btnCancel: Button, btnOK: Button}} 追加したボタン
     */
    function addButtonRow(targetDialog) {
        var btnRowGroup = targetDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["right", "center"];
        btnRowGroup.margins = BUTTON_ROW_MARGINS;
        btnRowGroup.alignment = ["right", "bottom"];

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnCancel.helpTip = getLabel("tooltip.cancel");

        var btnOK = btnRowGroup.add("button", undefined, "OK", { name: "ok" });
        btnOK.helpTip = getLabel("tooltip.ok");

        return { btnCancel: btnCancel, btnOK: btnOK };
    }

    /**
     * ダイアログと各UIコントロールを生成し、参照をまとめて返す
     * @returns {object} pageNumberDialog と各入力欄・チェックボックス・ボタン
     */
    function buildPageNumberDialog() {
        // タイトルバーにはバージョンを併記 / show the version in the title bar
        var pageNumberDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        pageNumberDialog.orientation = "column";
        pageNumberDialog.alignChildren = "left";

        // 3カラムレイアウト / 3-column layout
        var columnsGroup = pageNumberDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";

        // 左カラム: 接頭辞 / left column: prefix
        var prefixColumn = addColumnGroup(columnsGroup);
        var prefixField = addLabeledEditText(prefixColumn, getLabel("fieldLabel.prefix"), "", AFFIX_FIELD_CHARS,
            getLabel("tooltip.prefix"));

        // 中央カラム: 開始番号 + ゼロ埋め / center column: start number + zero pad
        var startNumberColumn = addColumnGroup(columnsGroup);
        var startNumberField = addLabeledEditText(startNumberColumn, getLabel("fieldLabel.start"), "1", START_NUMBER_FIELD_CHARS,
            getLabel("tooltip.start"));
        var zeroPadCheckbox = addCheckbox(startNumberColumn, getLabel("checkbox.zeroPad"),
            getLabel("tooltip.zeroPad"));

        // 右カラム: 接尾辞 + 総ページ表示 / right column: suffix + show-total
        var suffixColumn = addColumnGroup(columnsGroup);
        var suffixField = addLabeledEditText(suffixColumn, getLabel("fieldLabel.suffix"), "", AFFIX_FIELD_CHARS,
            getLabel("tooltip.suffix"));
        var totalPageCheckbox = addCheckbox(suffixColumn, getLabel("checkbox.showTotal"),
            getLabel("tooltip.showTotal"));

        var dialogButtons = addButtonRow(pageNumberDialog);

        // 透明度と表示位置の調整 / adjust opacity and position
        pageNumberDialog.opacity = DIALOG_OPACITY;
        pageNumberDialog.onShow = function () {
            pageNumberDialog.location = [pageNumberDialog.location[0] + DIALOG_OFFSET_X, pageNumberDialog.location[1]];
        };

        return {
            pageNumberDialog: pageNumberDialog,
            prefixField: prefixField,
            startNumberField: startNumberField,
            zeroPadCheckbox: zeroPadCheckbox,
            suffixField: suffixField,
            totalPageCheckbox: totalPageCheckbox,
            btnCancel: dialogButtons.btnCancel,
            btnOK: dialogButtons.btnOK
        };
    }

    /**
     * 開始番号欄を整数として読む
     * @param {object} dialogControls - buildPageNumberDialog() の戻り値
     * @returns {number} 開始番号（数値でなければ NaN）
     */
    function readStartNumber(dialogControls) {
        return parseInt(dialogControls.startNumberField.text, 10);
    }

    /**
     * 現在の入力値を書式オプションとしてまとめる
     * @param {object} dialogControls - buildPageNumberDialog() の戻り値
     * @returns {{prefix: string, suffix: string, zeroPad: boolean, showTotal: boolean}} 書式オプション
     */
    function readFormatOptions(dialogControls) {
        return {
            prefix: dialogControls.prefixField.text || "",
            suffix: dialogControls.suffixField.text || "",
            zeroPad: !!dialogControls.zeroPadCheckbox.value,
            showTotal: !!dialogControls.totalPageCheckbox.value
        };
    }

    // =========================================
    // 確定処理 / Commit
    // =========================================

    /**
     * OK確定時の雛形テキストを取得（優先候補 → 現在の選択 → レイヤー上の既存テキスト）
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {TextFrame} preferredText - 最優先の候補（実行時に選択していたテキスト）
     * @returns {TextFrame|null} 雛形テキスト
     */
    function resolveTemplateTextForCommit(doc, pagenumberLayer, preferredText) {
        var templateText = isTextFrame(preferredText) ? preferredText : getSelectedTextFrame();
        if (!isTextFrame(templateText)) {
            templateText = findTextFrameOnAnyArtboard(doc, pagenumberLayer);
        }
        return isTextFrame(templateText) ? templateText : null;
    }

    /**
     * 雛形テキストを _pagenumber 上へ移し、他のテキストを除去
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {TextFrame} templateText - 雛形テキスト
     * @returns {boolean} 移せたら true
     */
    function moveTemplateTextToPagenumberLayer(pagenumberLayer, templateText) {
        if (!isTextFrame(templateText)) return false;
        if (templateText.layer.name !== PAGENUMBER_LAYER_NAME) {
            tryCall(function () { templateText.move(pagenumberLayer, ElementPlacement.PLACEATBEGINNING); });
        }
        if (templateText.layer.name !== PAGENUMBER_LAYER_NAME) return false;
        removeOtherTextFrames(pagenumberLayer, templateText);
        return true;
    }

    /**
     * 確定用の連番を全アートボードへ適用
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} pagenumberLayer - _pagenumber レイヤー
     * @param {TextFrame} templateText - 雛形テキスト
     * @param {number} startNumber - 開始番号
     * @param {object} formatOptions - readFormatOptions() の戻り値
     * @returns {boolean} 配置できたら true
     */
    function applyNumberingToAllArtboards(doc, pagenumberLayer, templateText, startNumber, formatOptions) {
        // 複製に失敗した場合は採番せず、雛形をそのまま残す / when the duplication fails, leave the template as it is
        if (!seedAndPasteToAllArtboards(doc, templateText, startNumber)) return false;
        numberFramesInOrder(doc, pagenumberLayer, templateText, startNumber, formatOptions);
        return true;
    }

    /**
     * OK確定時の本処理：雛形テキストを _pagenumber へ移し、全アートボードへ連番を確定配置
     * @param {Document} doc - 対象ドキュメント
     * @param {object} dialogControls - buildPageNumberDialog() の戻り値
     * @param {TextFrame} originalTemplateText - 実行時に選択していたテキスト
     * @param {object} pagenumberSetup - setupPagenumberLayer() の戻り値（元の状態へ戻すため）
     * @returns {void}
     */
    function commitPageNumbers(doc, dialogControls, originalTemplateText, pagenumberSetup) {
        var startNumber = readStartNumber(dialogControls);
        if (isNaN(startNumber)) {
            alert(getLabel("alert.notNumber"));
            return;
        }

        // プレビュー用の退避レイヤーを破棄 / discard the preview backup layer
        forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);

        var pagenumberLayer = getOrCreateLayer(doc, PAGENUMBER_LAYER_NAME);
        var templateText = resolveTemplateTextForCommit(doc, pagenumberLayer, originalTemplateText);
        if (!templateText || !moveTemplateTextToPagenumberLayer(pagenumberLayer, templateText)) {
            alert(getLabel("alert.invalidSelection"));
            return;
        }

        var placed = applyNumberingToAllArtboards(doc, pagenumberLayer, templateText, startNumber, readFormatOptions(dialogControls));
        safeRedraw();
        restorePagenumberState(doc, pagenumberSetup.layer, pagenumberSetup.originalState, false);
        if (!placed) alert(getLabel("alert.commitFailed"));
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示し、選択テキストを雛形に全アートボードへページ番号を配置
     * @returns {void}
     */
    function main() {
        // テキスト未選択なら終了 / Exit if no text is selected
        var originalTemplateText = getSelectedTextFrame();
        if (!originalTemplateText) {
            alert(getLabel("alert.invalidSelection"));
            return;
        }

        var doc = app.activeDocument;
        var pagenumberSetup = setupPagenumberLayer(doc);
        var dialogControls = buildPageNumberDialog();
        var previewManager = new PreviewManager();

        /* 現在の入力値でライブプレビューを更新（前回分を巻き戻し、1ステップとして再実行）
           Refresh the live preview with current input values (roll back the previous one, run as a single step) */
        function refreshPreview() {
            var startNumber = readStartNumber(dialogControls);
            if (isNaN(startNumber)) return;

            var formatOptions = readFormatOptions(dialogControls);
            previewManager.rollback();
            previewManager.runAsStep(function () {
                return updatePreview(doc, PAGENUMBER_LAYER_NAME, startNumber, formatOptions);
            });
        }

        changeValueByArrowKey(dialogControls.startNumberField, refreshPreview);

        // 入力確定（Tabやフォーカス移動）でプレビューを更新。onChanging は1文字ごとに全アートボードを組み直すため使わない
        // Refresh on commit of the field (Tab or focus change); onChanging would rebuild every artboard on each keystroke
        dialogControls.prefixField.onChange = refreshPreview;
        dialogControls.suffixField.onChange = refreshPreview;
        dialogControls.startNumberField.onChange = refreshPreview;

        // チェックボックスのON/OFFでプレビューを更新（キー操作での切り替えは値を直接書き換えるため onClick は発火しない）
        // Update the preview when a checkbox is toggled (key shortcuts set .value directly, so onClick does not fire for them)
        dialogControls.zeroPadCheckbox.onClick = refreshPreview;
        dialogControls.totalPageCheckbox.onClick = refreshPreview;

        // Zキーでゼロ埋め、Aキーで総ページ表示をトグル（入力欄では文字入力を優先）
        // Z toggles zero-pad, A toggles show-total; typing in a text field takes precedence
        addToggleKeyHandler(dialogControls.pageNumberDialog, "Z", dialogControls.zeroPadCheckbox, refreshPreview, true);
        addToggleKeyHandler(dialogControls.pageNumberDialog, "A", dialogControls.totalPageCheckbox, refreshPreview, true);

        // OK：プレビュー分を全Undoしてから確定処理を1回だけ実行 / OK: undo all preview steps, then run the commit action once
        dialogControls.btnOK.onClick = function () {
            previewManager.commit(function () {
                commitPageNumbers(doc, dialogControls, originalTemplateText, pagenumberSetup);
            });
            forceRemoveLayerByName(doc, BACKUP_LAYER_NAME);
            dialogControls.pageNumberDialog.close(1);
        };

        // キャンセル：プレビューを巻き戻し、退避テキストと _pagenumber 状態を復元 / Cancel: roll back the preview, restore the backed-up text and the _pagenumber state
        dialogControls.btnCancel.onClick = function () {
            previewManager.rollback();
            restorePreviewBackupOnCancel(doc);
            restorePagenumberState(doc, pagenumberSetup.layer, pagenumberSetup.originalState, true);
            dialogControls.pageNumberDialog.close(0);
        };

        dialogControls.startNumberField.active = true;

        // 初回プレビュー / first preview pass
        refreshPreview();

        dialogControls.pageNumberDialog.show();
    }

    main();

})();
