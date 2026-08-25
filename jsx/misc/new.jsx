#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトから、新規レイヤー・新規アートボード・新規ドキュメントを作成します。
何を作成するかは、ダイアログのラジオボタンで選択します。

詳細は README を参照してください。

### Overview

Creates a new layer, a new artboard, or a new document from the selected objects.
Which one is produced is chosen with radio buttons in the dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "new";                          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-07-29";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-29";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/new.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/new.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

// アートボードの配置ロジック / Artboard layout logic
// AddArtboardPlus.jsx（jsx/artboard/AddArtboardPlus.jsx）から移植
// Ported from AddArtboardPlus.jsx
// Original: Copyright (c) 2018 Takeshi Umeda (noellabo)
// https://dtp-discourse.jp/t/illustrator/99

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 新規アートボードの挿入位置 / Insert position of the new artboard */
    /* true = 現在のアートボードの次 / false = 末尾 */
    /* true = after the current artboard, false = at the end */
    var ARTBOARD_INSERT_AFTER_CURRENT = true;

    /* 新規アートボードを並べる方向 / Direction the new artboard runs along */
    /* 0 = 右（横並び） / 1 = 下（縦並び） */
    /* 0 = right (horizontal), 1 = down (vertical) */
    var ARTBOARD_DIRECTION_AXIS = 0;

    /* 新規アートボード名に付ける接尾辞 / Suffix appended to the new artboard's name */
    var ARTBOARD_NAME_SUFFIX = "_new";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境の言語を判定します。
     *
     * @returns {string} 日本語環境なら "ja"、それ以外は "en"。
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var lang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle:    { ja: "選択オブジェクトから作成", en: "New from Selection" },
        panelTarget:    { ja: "作成するもの", en: "Create" },
        targetLayer:    { ja: "レイヤー", en: "Layer" },
        targetArtboard: { ja: "アートボード", en: "Artboard" },
        targetDocument: { ja: "ドキュメント", en: "Document" },
        duplicate:      { ja: "複製", en: "Duplicate" },
        ok:             { ja: "OK", en: "OK" },
        cancel:         { ja: "キャンセル", en: "Cancel" },
        layerName:      { ja: "新規レイヤー", en: "New Layer" },
        noDocument:     { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection:    { ja: "オブジェクトが選択されていません。", en: "No objects are selected." },
        artboardLimit:  { ja: "アートボードの最大作成可能数を超えています。", en: "The maximum number of artboards would be exceeded." },
        noSpace:        { ja: "アートボードを作成する十分なスペースがありません。", en: "There is not enough space to create the artboard." },
        needsSave:      { ja: "ドキュメントを複製するため、先に保存してください。", en: "Save the document first so it can be duplicated." },
        copyFailed:     { ja: "ドキュメントの複製に失敗しました。", en: "Failed to duplicate the document." },
        structureError: { ja: "複製したドキュメントの構造が一致しないため、中止しました。", en: "Aborted: the duplicated document's structure does not match." }
    };

    /**
     * ラベル定義から現在の言語の文言を取得します。
     *
     * @param {string} key - LABELS のキー。
     * @returns {string} 対応する文言。見つからない場合はキーをそのまま返す。
     */
    function L(key) {
        if (LABELS[key] && LABELS[key][lang]) return LABELS[key][lang];
        if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
        return key;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログの設定値。
     *
     * @typedef {object} DialogOptions
     * @property {string} target - 作成対象（"layer" / "artboard" / "document"）。
     * @property {boolean} duplicate - 元のオブジェクトを残して複製する場合は true。
     */

    /**
     * 何を作成するかを選択するダイアログを表示します。
     *
     * @returns {DialogOptions|null} 選択された設定。キャンセルされた場合は null。
     */
    function showTargetDialog() {
        var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.spacing = 12;
        dialog.margins = 16;

        /* 作成対象パネル / Target panel */
        var targetPanel = dialog.add("panel", undefined, L("panelTarget"));
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.spacing = 8;
        targetPanel.margins = [15, 20, 15, 15];

        var layerRadio = targetPanel.add("radiobutton", undefined, L("targetLayer"));
        var artboardRadio = targetPanel.add("radiobutton", undefined, L("targetArtboard"));
        var documentRadio = targetPanel.add("radiobutton", undefined, L("targetDocument"));
        layerRadio.value = true;

        /* 複製オプション / Duplicate option */
        var duplicateCheck = dialog.add("checkbox", undefined, L("duplicate"));
        duplicateCheck.value = false;

        /**
         * ［複製］の有効・無効を作成対象に合わせて切り替えます。
         *
         * ドキュメント作成は元のドキュメントを残す＝常に複製になるため、
         * チェックボックスをディム表示にします。
         *
         * @returns {void}
         */
        function updateDuplicateState() {
            duplicateCheck.enabled = !documentRadio.value;
        }

        layerRadio.onClick = updateDuplicateState;
        artboardRadio.onClick = updateDuplicateState;
        documentRadio.onClick = updateDuplicateState;
        updateDuplicateState();

        /* ボタン行 / Button row */
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = ["right", "top"];
        buttonGroup.add("button", undefined, L("cancel"), { name: "cancel" });
        buttonGroup.add("button", undefined, L("ok"), { name: "ok" });

        if (dialog.show() !== 1) return null;

        var target = "layer";
        if (artboardRadio.value) target = "artboard";
        else if (documentRadio.value) target = "document";

        return {
            target: target,
            duplicate: (duplicateCheck.enabled && duplicateCheck.value)
        };
    }

    // =========================================
    // 選択オブジェクトの取得 / Collecting the selection
    // =========================================

    /**
     * 選択オブジェクトを配列にスナップショットします。
     *
     * `doc.selection` は参照するたびに現在の選択状態から作り直されるため、
     * move() で選択が変化するループ内で直接参照すると対象を取りこぼします。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @returns {Array<PageItem>} 選択オブジェクトの配列。
     */
    function snapshotSelection(doc) {
        var items = [];
        var selection = doc.selection;

        for (var i = 0; i < selection.length; i++) {
            items.push(selection[i]);
        }
        return items;
    }

    /**
     * 重ね順の比較に使うソートキーを取得します。
     *
     * レイヤーとコンテナそれぞれの zOrderPosition を組み合わせ、
     * 値が大きいほど前面（上）になるようにします。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {{layerOrder: number, itemOrder: number}} 並べ替え用のキー。
     */
    function getStackKey(item) {
        var layerOrder = 0;
        var itemOrder = 0;

        /* 削除済みや取得できない参照に触ると例外になるため保護
           A stale or unsupported reference throws, so guard the lookup */
        try {
            layerOrder = item.layer.zOrderPosition;
        } catch (_) {}
        try {
            itemOrder = item.zOrderPosition;
        } catch (_) {}

        return { layerOrder: layerOrder, itemOrder: itemOrder };
    }

    /**
     * オブジェクトを前面から背面の順（元の重ね順）に並べ替えます。
     *
     * @param {Array<PageItem>} items - 並べ替えるオブジェクトの配列。
     * @returns {Array<PageItem>} 前面が先頭になるよう並べ替えた配列。
     */
    function sortByStackOrder(items) {
        var entries = [];

        for (var i = 0; i < items.length; i++) {
            var key = getStackKey(items[i]);
            entries.push({ item: items[i], key: key, index: i });
        }

        entries.sort(function (a, b) {
            if (a.key.layerOrder !== b.key.layerOrder) {
                return b.key.layerOrder - a.key.layerOrder;
            }
            if (a.key.itemOrder !== b.key.itemOrder) {
                return b.key.itemOrder - a.key.itemOrder;
            }
            /* キーが同じときは元の並び順を保つ / Keep the original order for ties */
            return a.index - b.index;
        });

        var sorted = [];
        for (var j = 0; j < entries.length; j++) {
            sorted.push(entries[j].item);
        }
        return sorted;
    }

    // =========================================
    // レイヤーの作成 / Creating the layer
    // =========================================

    /**
     * 同じ名前のレイヤーがすでに存在するかを判定します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {string} name - 探すレイヤー名。
     * @returns {boolean} 存在する場合は true。
     */
    function layerNameExists(doc, name) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === name) return true;
        }
        return false;
    }

    /**
     * 重複しないレイヤー名を作ります（「新規レイヤー 2」のように連番を付与）。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {string} baseName - 基準となるレイヤー名。
     * @returns {string} 重複しないレイヤー名。
     */
    function makeUniqueLayerName(doc, baseName) {
        if (!layerNameExists(doc, baseName)) return baseName;

        var suffix = 2;
        while (layerNameExists(doc, baseName + " " + suffix)) {
            suffix++;
        }
        return baseName + " " + suffix;
    }

    // =========================================
    // レイヤーを作成 / Create a layer
    // =========================================

    /**
     * 選択オブジェクトを収めた新規レイヤーを作成します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {boolean} useDuplicate - true なら元のオブジェクトを残して複製する。
     * @returns {void}
     */
    function createLayerFromSelection(doc, useDuplicate) {
        /* ループ前にスナップショットを取る（move で選択が変化するため）
           Snapshot before the loop: move() changes the live selection */
        var items = snapshotSelection(doc);
        if (items.length === 0) {
            alert(L("noSelection"));
            return;
        }

        items = sortByStackOrder(items);

        var newLayer = doc.layers.add();
        newLayer.name = makeUniqueLayerName(doc, L("layerName"));

        /* 前面のものから PLACEATEND で送ると元の重ね順が保たれる
           Sending front-to-back with PLACEATEND preserves the original order */
        var placed = [];
        for (var i = 0; i < items.length; i++) {
            try {
                if (useDuplicate) {
                    placed.push(items[i].duplicate(newLayer, ElementPlacement.PLACEATEND));
                } else {
                    items[i].move(newLayer, ElementPlacement.PLACEATEND);
                    placed.push(items[i]);
                }
            } catch (_) {}
        }

        /* 1つも配置できなかったときは空のレイヤーを残さない
           Do not leave an empty layer behind when nothing was placed */
        if (placed.length === 0) {
            newLayer.remove();
            return;
        }

        /* 新規レイヤー側のオブジェクトを選択状態にする
           Select the items that ended up on the new layer */
        try {
            doc.selection = placed;
        } catch (_) {}
    }

    // =========================================
    // 選択範囲とアートボードの対応 / Selection and artboards
    // =========================================

    /**
     * 複数オブジェクトを囲む矩形を求めます。
     *
     * @param {Array<PageItem>} items - 対象のオブジェクト。
     * @returns {Array<number>|null} [左, 上, 右, 下]。求められない場合は null。
     */
    function getUnionBounds(items) {
        var bounds = null;

        for (var i = 0; i < items.length; i++) {
            var itemBounds;
            /* 削除済みの参照に触ると例外になるため保護 / A stale reference throws */
            try {
                itemBounds = items[i].geometricBounds;
            } catch (_) {
                continue;
            }
            if (bounds === null) {
                bounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
                continue;
            }
            if (itemBounds[0] < bounds[0]) bounds[0] = itemBounds[0];
            if (itemBounds[1] > bounds[1]) bounds[1] = itemBounds[1];
            if (itemBounds[2] > bounds[2]) bounds[2] = itemBounds[2];
            if (itemBounds[3] < bounds[3]) bounds[3] = itemBounds[3];
        }
        return bounds;
    }

    /**
     * オブジェクト群の中心を含むアートボードのインデックスを求めます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {Array<PageItem>} items - 対象のオブジェクト。
     * @param {number} fallbackIndex - 見つからないときに返すインデックス。
     * @returns {number} アートボードのインデックス。
     */
    function findArtboardIndexForItems(doc, items, fallbackIndex) {
        var bounds = getUnionBounds(items);
        if (bounds === null) return fallbackIndex;

        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;

        for (var i = 0; i < doc.artboards.length; i++) {
            var rect = doc.artboards[i].artboardRect;
            if (centerX >= rect[0] && centerX <= rect[2] &&
                centerY <= rect[1] && centerY >= rect[3]) {
                return i;
            }
        }
        return fallbackIndex;
    }

    // =========================================
    // アートボードの配置計算 / Artboard layout math
    // =========================================

    /**
     * 軸方向のアートボードサイズを取得します。
     *
     * @param {Array<number>} rect - artboardRect（[左, 上, 右, 下]）。
     * @param {number} axisIndex - 0 なら幅、1 なら高さ。
     * @returns {number} 指定軸のサイズ。
     */
    function getAxisSize(rect, axisIndex) {
        return (axisIndex === 0)
            ? (rect[2] - rect[0])           /* 幅 / width */
            : Math.abs(rect[3] - rect[1]);  /* 高さ / height */
    }

    /**
     * 隣り合う2枚のアートボードから主軸を判定します。
     *
     * @param {Array<number>} rectA - 1枚目の artboardRect。
     * @param {Array<number>} rectB - 2枚目の artboardRect。
     * @returns {number} 左端が同じなら 1（縦並び）、違えば 0（横並び）。
     */
    function getPrimaryAxisIndex(rectA, rectB) {
        return (rectA[0] === rectB[0]) ? 1 : 0;
    }

    /**
     * 既存アートボードの並びから現在の間隔（pt）を推定します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {number} fallbackPt - 2枚未満のときに使う値（pt）。
     * @returns {number} 推定した間隔（pt）。
     */
    function computeAutoSpacingPt(doc, fallbackPt) {
        var artboardList = doc.artboards;
        if (artboardList.length < 2) return fallbackPt;

        var baseRect = artboardList[0].artboardRect;
        var adjacentRect = artboardList[1].artboardRect;
        var primaryAxisIndex = getPrimaryAxisIndex(baseRect, adjacentRect);
        var pitch = Math.abs(adjacentRect[primaryAxisIndex] - baseRect[primaryAxisIndex]);
        var spacing = pitch - getAxisSize(baseRect, primaryAxisIndex);
        return (spacing < 0) ? 0 : spacing;
    }

    /**
     * 最大キャンバス範囲を取得します。
     *
     * Original idea by OMOTI
     * https://forums.adobe.com/thread/2459293
     *
     * @param {Document} doc - 対象のドキュメント。
     * @returns {Array<number>} [左, 上, 右, 下]。
     */
    function getLargestCanvasBounds(doc) {
        var LARGEST_SIZE = 16383;

        var tempLayer = doc.layers.add();
        var tempText = tempLayer.textFrames.add();
        var left = tempText.matrix.mValueTX;
        var top = tempText.matrix.mValueTY;
        tempLayer.remove();

        return [left, top, left + LARGEST_SIZE, top - LARGEST_SIZE];
    }

    /**
     * アートボードの配置計画。
     *
     * @typedef {object} ArtboardLayout
     * @property {boolean} canInherit - 既存の並びのグリッドを引き継げる場合は true。
     * @property {Array<number>} gridStep - グリッド1セルあたりの移動量 [X, Y]。
     * @property {number} columns - グリッドの列数。
     * @property {number} primaryAxisIndex - 主軸（0=横 / 1=縦）。
     * @property {number} secondaryAxisIndex - 副軸。
     * @property {number} primarySign - 主軸方向の符号（+1 / -1）。
     * @property {number} spacing - アートボード間の間隔（pt）。
     * @property {Array<number>} baseRect - 先頭アートボードの artboardRect。
     * @property {Array<number>} referenceRect - サイズの基準にする artboardRect。
     */

    /**
     * 既存の並びを解析し、新規アートボードの配置計画を作ります。
     *
     * 既存の並びが ARTBOARD_DIRECTION_AXIS と一致する場合は、そのピッチと列数を
     * 引き継ぎます。一致しない場合や1枚のみの場合はグリッドを使わず、挿入位置の
     * 直前のアートボードを基準に配置します。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {number} referenceIndex - サイズの基準にするアートボードのインデックス。
     * @returns {ArtboardLayout|null} 配置計画。スペースが足りない場合は null。
     */
    function planArtboardLayout(doc, referenceIndex) {
        var artboards = doc.artboards;
        var artboardCount = artboards.length;

        var spacingPreferencePt = app.preferences.getRealPreference('plugin/ArtboardRearrange/ArtboardSpacing');
        var spacing = computeAutoSpacingPt(doc, spacingPreferencePt);

        var baseRect = artboards[0].artboardRect;
        var referenceRect = artboards[referenceIndex].artboardRect;

        var primaryAxisIndex = ARTBOARD_DIRECTION_AXIS;
        var secondaryAxisIndex = 1 - primaryAxisIndex;

        /* グリッド1セルあたりの移動量（Y軸は上→下なので間隔を引く）
           Grid step per cell (the Y axis is top-down, so the spacing is subtracted) */
        var gridStep = [
            (baseRect[2] - baseRect[0]) + spacing,
            (baseRect[3] - baseRect[1]) - spacing
        ];
        var columns = 0;
        var canInherit = false;

        /* 既存の並びが指定方向と一致するときだけ、ピッチと列数を引き継ぐ
           Inherit pitch and columns only when the existing layout matches the direction */
        if (artboardCount >= 2) {
            var adjacentRect = artboards[1].artboardRect;

            if (getPrimaryAxisIndex(baseRect, adjacentRect) === primaryAxisIndex) {
                canInherit = true;
                gridStep[primaryAxisIndex] = adjacentRect[primaryAxisIndex] - baseRect[primaryAxisIndex];

                for (var i = 2; i < artboardCount; i++) {
                    var scannedRect = artboards[i].artboardRect;
                    if (baseRect[secondaryAxisIndex] !== scannedRect[secondaryAxisIndex]) {
                        gridStep[secondaryAxisIndex] = scannedRect[secondaryAxisIndex] - baseRect[secondaryAxisIndex];
                        columns = i;
                        break;
                    }
                }
            }
        }

        var canvasRect = getLargestCanvasBounds(doc);
        var gridUnitRect = [
            baseRect[0] + Math.abs(gridStep[0]),
            baseRect[1] - Math.abs(gridStep[1]),
            baseRect[0],
            baseRect[1]
        ];

        var primaryEdgeIndex =
            (primaryAxisIndex ^ +(gridStep[primaryAxisIndex] < 0)) ? primaryAxisIndex : primaryAxisIndex + 2;
        var secondaryEdgeIndex =
            (secondaryAxisIndex ^ +(gridStep[secondaryAxisIndex] < 0)) ? secondaryAxisIndex : secondaryAxisIndex + 2;

        columns = columns ||
            Math.abs(Math.floor((canvasRect[primaryEdgeIndex] - gridUnitRect[primaryEdgeIndex]) / gridStep[primaryAxisIndex]));
        var rows =
            Math.abs(Math.floor((canvasRect[secondaryEdgeIndex] - gridUnitRect[secondaryEdgeIndex]) / gridStep[secondaryAxisIndex]));

        if (artboardCount + 1 > columns * rows) return null;

        return {
            canInherit: canInherit,
            gridStep: gridStep,
            columns: columns,
            primaryAxisIndex: primaryAxisIndex,
            secondaryAxisIndex: secondaryAxisIndex,
            primarySign: (gridStep[primaryAxisIndex] < 0) ? -1 : 1,
            spacing: spacing,
            baseRect: baseRect,
            referenceRect: referenceRect
        };
    }

    /**
     * グリッド上のインデックスに対応する位置を求めます。
     *
     * @param {ArtboardLayout} layout - 配置計画。
     * @param {number} gridIndex - グリッド上のインデックス。
     * @returns {Array<number>} [左, 上] の座標。
     */
    function getArtboardGridPosition(layout, gridIndex) {
        var offset = [];
        offset[layout.primaryAxisIndex] = (gridIndex % layout.columns) * layout.gridStep[layout.primaryAxisIndex];
        offset[layout.secondaryAxisIndex] = Math.floor(gridIndex / layout.columns) * layout.gridStep[layout.secondaryAxisIndex];
        return [layout.baseRect[0] + offset[0], layout.baseRect[1] + offset[1]];
    }

    /**
     * グリッドを引き継げないときの配置位置を求めます。
     *
     * 先頭ではなく、挿入位置の直前のアートボードを基準に指定方向へ1枚分進めます。
     * グリッドのインデックスを歩数に使うと、既存の並びと軸が違う場合に
     * 枚数分だけ離れた位置へ飛んでしまうため、実位置から積み上げます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {ArtboardLayout} layout - 配置計画。
     * @param {number} anchorIndex - 基準にするアートボードのインデックス。
     * @returns {Array<number>} [左, 上] の座標。
     */
    function getAnchoredArtboardPosition(doc, layout, anchorIndex) {
        var anchorRect = doc.artboards[anchorIndex].artboardRect;
        var advance = getAxisSize(anchorRect, layout.primaryAxisIndex) + layout.spacing;
        var position = [anchorRect[0], anchorRect[1]];
        position[layout.primaryAxisIndex] += layout.primarySign * advance;
        return position;
    }

    // =========================================
    // アートボードの移動 / Moving artboards
    // =========================================

    /**
     * アートボード上のアイテムを収集します。
     *
     * doc.pageItems はグループ・複合パスの子まで再帰的に含むため、最上位
     * （親がレイヤー）のアイテムのみを対象にします。子は親と一緒に動くので、
     * ここで拾うと二重に処理されてしまいます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {Array<number>} artboardRect - 対象アートボードの矩形。
     * @returns {Array<PageItem>} アートボードに属するアイテム。
     */
    function getItemsAssignedToArtboard(doc, artboardRect) {
        var items = [];

        for (var k = 0; k < doc.pageItems.length; k++) {
            var item = doc.pageItems[k];
            if (item.parent.typename !== 'Layer') continue;

            var itemBounds = item.geometricBounds;
            var centerX = (itemBounds[0] + itemBounds[2]) / 2;
            var centerY = (itemBounds[1] + itemBounds[3]) / 2;
            if (centerX >= artboardRect[0] && centerX <= artboardRect[2] &&
                centerY <= artboardRect[1] && centerY >= artboardRect[3]) {
                items.push(item);
            }
        }
        return items;
    }

    /**
     * アイテムと祖先レイヤーのロック／表示状態を一時的に解除します。
     *
     * ロック／非表示のアイテムは translate() が例外になります。途中で例外になると
     * そこまで動かしたアートボードだけが残って崩れるため、一時解除してから処理します。
     *
     * @param {PageItem} item - 対象のオブジェクト。
     * @returns {Array<object>} 復元用の情報リスト。
     */
    function unlockItemTemporarily(item) {
        var lockRestoreList = [];
        var ancestorLayer = item.parent;

        while (ancestorLayer && ancestorLayer.typename === 'Layer') {
            if (ancestorLayer.locked) {
                ancestorLayer.locked = false;
                lockRestoreList.push({ target: ancestorLayer, property: 'locked', value: true });
            }
            if (!ancestorLayer.visible) {
                ancestorLayer.visible = true;
                lockRestoreList.push({ target: ancestorLayer, property: 'visible', value: false });
            }
            ancestorLayer = ancestorLayer.parent;
        }
        if (item.locked) {
            item.locked = false;
            lockRestoreList.push({ target: item, property: 'locked', value: true });
        }
        if (item.hidden) {
            item.hidden = false;
            lockRestoreList.push({ target: item, property: 'hidden', value: true });
        }
        return lockRestoreList;
    }

    /**
     * 一時解除したロック／表示状態を元に戻します。
     *
     * @param {Array<object>} lockRestoreList - unlockItemTemporarily の戻り値。
     * @returns {void}
     */
    function restoreLockedState(lockRestoreList) {
        /* 解除と逆順に戻す（内側→外側）/ Restore in reverse order (inner → outer) */
        for (var i = lockRestoreList.length - 1; i >= 0; i--) {
            lockRestoreList[i].target[lockRestoreList[i].property] = lockRestoreList[i].value;
        }
    }

    /**
     * 挿入位置以降のアートボードを1枚分だけ後ろへずらし、アートワークも一緒に運びます。
     *
     * アイテムの帰属は移動前の位置でまとめて取得（スナップショット）してから動かすため、
     * 移動順による取り違えが起きません。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {ArtboardLayout} layout - 配置計画。
     * @param {number} fromIndex - ずらし始めるインデックス。
     * @returns {void}
     */
    function relayoutExistingArtboards(doc, layout, fromIndex) {
        var artboards = doc.artboards;
        var artboardCount = artboards.length;
        var plannedMoves = [];

        for (var i = fromIndex; i < artboardCount; i++) {
            var currentRect = artboards[i].artboardRect;
            var targetPosition = getArtboardGridPosition(layout, i + 1);
            plannedMoves.push({
                index: i,
                dx: targetPosition[0] - currentRect[0],
                dy: targetPosition[1] - currentRect[1],
                items: getItemsAssignedToArtboard(doc, currentRect)
            });
        }

        for (var j = 0; j < plannedMoves.length; j++) {
            var plannedMove = plannedMoves[j];

            for (var k = 0; k < plannedMove.items.length; k++) {
                var itemToMove = plannedMove.items[k];
                var lockRestoreList = unlockItemTemporarily(itemToMove);
                try {
                    itemToMove.translate(plannedMove.dx, plannedMove.dy);
                } finally {
                    restoreLockedState(lockRestoreList);
                }
            }

            var movedRect = artboards[plannedMove.index].artboardRect;
            artboards[plannedMove.index].artboardRect = [
                movedRect[0] + plannedMove.dx, movedRect[1] + plannedMove.dy,
                movedRect[2] + plannedMove.dx, movedRect[3] + plannedMove.dy
            ];
        }
    }

    /**
     * 末尾に追加されたアートボードを、パネル上の挿入位置へ並べ替えます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {number} insertIndex - 挿入位置のインデックス。
     * @param {number} originalCount - 追加前のアートボード数。
     * @returns {void}
     */
    function reorderAppendedArtboard(doc, insertIndex, originalCount) {
        var artboards = doc.artboards;

        /* 追加分の枠と名前を退避（後続シフトで上書きされる前に）
           Save the appended rect and name before the shift clobbers them */
        var appendedRect = artboards[originalCount].artboardRect;
        var appendedName = artboards[originalCount].name;

        /* 後続を1枚分だけ後ろへ（高位から処理して上書き衝突を回避）
           Shift trailing artboards back by one (high → low to avoid clobbering) */
        for (var j = originalCount - 1; j >= insertIndex; j--) {
            artboards[j + 1].artboardRect = artboards[j].artboardRect;
            artboards[j + 1].name = artboards[j].name;
        }

        artboards[insertIndex].artboardRect = appendedRect;
        artboards[insertIndex].name = appendedName;
    }

    // =========================================
    // アートボードを作成 / Create an artboard
    // =========================================

    /**
     * 選択オブジェクトから新規アートボードを作成します。
     *
     * 新規アートボードのサイズはアクティブアートボードと同じで、選択オブジェクトは
     * 元のアートボード内での相対位置を保ったまま移動（または複製）されます。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @param {boolean} useDuplicate - true なら元のオブジェクトを残して複製する。
     * @returns {void}
     */
    function createArtboardFromSelection(doc, useDuplicate) {
        var items = snapshotSelection(doc);
        if (items.length === 0) {
            alert(L("noSelection"));
            return;
        }

        var artboards = doc.artboards;
        var artboardCount = artboards.length;
        var artboardLimit = (parseFloat(app.version) >= 22) ? 1000 : 100;

        if (artboardCount + 1 > artboardLimit) {
            alert(L("artboardLimit"));
            return;
        }

        var activeIndex = artboards.getActiveArtboardIndex();
        var insertIndex = ARTBOARD_INSERT_AFTER_CURRENT ? (activeIndex + 1) : artboardCount;

        var layout = planArtboardLayout(doc, activeIndex);
        if (layout === null) {
            alert(L("noSpace"));
            return;
        }

        /* 選択オブジェクトが載っているアートボードは、ずらす前に特定しておく
           Resolve the source artboard before anything shifts */
        var sourceIndex = findArtboardIndexForItems(doc, items, activeIndex);

        /* 既存アートボードを後ろへずらして挿入スペースを空ける。
           グリッドを引き継げないときは既存の並びを動かさない（別軸のグリッドへ
           流し込むと並びが崩れるため）
           Free the insert space. When the grid can't be inherited, leave the existing
           artboards alone — re-flowing onto the other axis would break the arrangement */
        if (layout.canInherit) {
            relayoutExistingArtboards(doc, layout, insertIndex);
        }

        /* 基準の矩形はずらしたあとに取得する。選択オブジェクトも一緒にずれているので、
           この矩形との差分をとれば二重移動にならない
           Read the source rect after the shift: the selection moved with it, so the
           delta against this rect never double-counts */
        var sourceRect = artboards[sourceIndex].artboardRect;

        var newPosition = layout.canInherit
            ? getArtboardGridPosition(layout, insertIndex)
            : getAnchoredArtboardPosition(doc, layout, insertIndex - 1);

        var newWidth = layout.referenceRect[2] - layout.referenceRect[0];
        var newHeight = layout.referenceRect[3] - layout.referenceRect[1];

        artboards.add([
            newPosition[0], newPosition[1],
            newPosition[0] + newWidth, newPosition[1] + newHeight
        ]);

        /* パネル上の順序を挿入位置へ並べ替え / Reorder in the panel */
        reorderAppendedArtboard(doc, insertIndex, artboardCount);
        artboards[insertIndex].name = artboards[activeIndex].name + ARTBOARD_NAME_SUFFIX;

        /* 相対位置を保ったまま新規アートボードへ / Carry the relative position over */
        var offsetX = newPosition[0] - sourceRect[0];
        var offsetY = newPosition[1] - sourceRect[1];

        var placed = [];
        for (var i = 0; i < items.length; i++) {
            try {
                if (useDuplicate) {
                    var copiedItem = items[i].duplicate();
                    copiedItem.translate(offsetX, offsetY);
                    placed.push(copiedItem);
                } else {
                    items[i].translate(offsetX, offsetY);
                    placed.push(items[i]);
                }
            } catch (_) {}
        }

        artboards.setActiveArtboardIndex(insertIndex);
        try {
            doc.selection = placed;
        } catch (_) {}
        app.redraw();
    }

    // =========================================
    // ドキュメントを作成 / Create a document
    // =========================================

    /**
     * 重複しない一時ファイルを作ります（「temp-元のファイル名」、重複時は連番）。
     *
     * @param {Folder} folder - 保存先フォルダー。
     * @param {string} fileName - 元のファイル名（拡張子込み）。
     * @returns {File} 重複しない一時ファイル。
     */
    function makeTempFile(folder, fileName) {
        /* 「my.logo.ai」のような名前でも拡張子だけを切り出す
           Split off the extension only, even for names like "my.logo.ai" */
        var dotIndex = fileName.lastIndexOf('.');
        var baseName = (dotIndex > 0) ? fileName.substring(0, dotIndex) : fileName;
        var extension = (dotIndex > 0) ? fileName.substring(dotIndex) : '';
        var tempBaseName = "temp-" + baseName;

        var candidate = new File(folder.fsName + "/" + tempBaseName + extension);
        var counter = 1;
        while (candidate.exists) {
            candidate = new File(folder.fsName + "/" + tempBaseName + "-" + counter + extension);
            counter++;
        }
        return candidate;
    }

    /**
     * コンテナの直接の子アイテムをz順で取得します。
     *
     * layer.pageItems / group.pageItems は子孫まで含むことがあるため、
     * parent が一致するものだけを拾います。
     *
     * @param {Layer|GroupItem|CompoundPathItem} container - 対象のコンテナ。
     * @returns {Array<PageItem>} 直接の子アイテム。
     */
    function getDirectChildItems(container) {
        var children = [];
        var collection = container.pageItems;

        for (var i = 0; i < collection.length; i++) {
            if (collection[i].parent === container) children.push(collection[i]);
        }
        return children;
    }

    /**
     * コンテナ配下のアイテムを、決まった順序で再帰的に集めます。
     *
     * @param {Layer|GroupItem|CompoundPathItem} container - 対象のコンテナ。
     * @param {Array<PageItem>} out - 集めたアイテムを追加する配列。
     * @returns {void}
     */
    function collectItemsInOrder(container, out) {
        var children = getDirectChildItems(container);

        for (var i = 0; i < children.length; i++) {
            out.push(children[i]);
            var typeName = children[i].typename;
            if (typeName === 'GroupItem' || typeName === 'CompoundPathItem') {
                collectItemsInOrder(children[i], out);
            }
        }
    }

    /**
     * レイヤー（サブレイヤーを含む）配下のアイテムを順に集めます。
     *
     * @param {Layers} layers - 対象のレイヤーコレクション。
     * @param {Array<PageItem>} out - 集めたアイテムを追加する配列。
     * @returns {void}
     */
    function collectLayerItemsInOrder(layers, out) {
        for (var i = 0; i < layers.length; i++) {
            collectItemsInOrder(layers[i], out);
            collectLayerItemsInOrder(layers[i].layers, out);
        }
    }

    /**
     * ドキュメント内の全アイテムを、決まった順序で集めます。
     *
     * 同一内容のドキュメント同士なら同じ順序で並ぶため、この配列の
     * インデックスをドキュメント間の対応付けに使えます。
     *
     * @param {Document} targetDoc - 対象のドキュメント。
     * @returns {Array<PageItem>} 全アイテム。
     */
    function collectDocumentItems(targetDoc) {
        var out = [];
        collectLayerItemsInOrder(targetDoc.layers, out);
        return out;
    }

    /**
     * すべてのレイヤーのロック・非表示を解除します。
     *
     * ロックまたは非表示のアイテムは remove() が例外になるため、削除の前に解除します。
     * 複製側でのみ使うので、元に戻す必要はありません。
     *
     * @param {Layers} layers - 対象のレイヤーコレクション。
     * @returns {void}
     */
    function unlockAllLayers(layers) {
        for (var i = 0; i < layers.length; i++) {
            layers[i].locked = false;
            layers[i].visible = true;
            unlockAllLayers(layers[i].layers);
        }
    }

    /**
     * 配列の中に指定のオブジェクトが含まれるかを判定します。
     *
     * @param {Array<PageItem>} items - 探す対象の配列。
     * @param {PageItem} target - 探すオブジェクト。
     * @returns {boolean} 含まれる場合は true。
     */
    function containsItem(items, target) {
        for (var i = 0; i < items.length; i++) {
            if (items[i] === target) return true;
        }
        return false;
    }

    /**
     * 選択オブジェクトから新規ドキュメントを作成します。
     *
     * 元ドキュメントのファイルをコピーして複製ドキュメントを作り、複製側で
     * 選択オブジェクト以外と、現在のアートボード以外を削除します。
     * スウォッチ・シンボル・ドキュメント設定などがそのまま引き継がれます。
     *
     * `saveAs()` は別名保存であって複製ではなく（元ドキュメント自体が保存先に
     * 紐づき直してしまう）、`File.copy()` はディスク上の保存済みの状態を写すため、
     * 保存済みのドキュメントだけを対象にします。
     *
     * @param {Document} doc - 対象のドキュメント。
     * @returns {void}
     */
    function createDocumentFromSelection(doc) {
        var items = snapshotSelection(doc);
        if (items.length === 0) {
            alert(L("noSelection"));
            return;
        }

        /* 未保存・未変更でないとディスク上の状態と画面上の状態がずれる
           Unsaved changes would make the copy differ from what is on screen */
        var originalFile = null;
        try {
            originalFile = doc.fullName;
        } catch (_) {}
        if (!doc.saved || originalFile === null || !originalFile.exists) {
            alert(L("needsSave"));
            return;
        }

        /* 元ドキュメント側で「残すアイテム」の位置を記録する。
           複製は同一ファイルのコピーなので、同じ順序・同じ個数で並ぶ
           Record which positions to keep; the copy enumerates in the same order */
        var allItems = collectDocumentItems(doc);
        var selectedFlags = [];
        var keepFlags = [];
        var i;

        for (i = 0; i < allItems.length; i++) {
            var isSelected = containsItem(items, allItems[i]);
            selectedFlags.push(isSelected);
            keepFlags.push(isSelected);
        }

        /* 選択オブジェクトの祖先（グループ・複合パス）も残す。
           祖先を消すと中の選択オブジェクトごと消えてしまうため
           Keep the ancestors too: removing a group takes its kept children with it */
        for (i = 0; i < allItems.length; i++) {
            if (!selectedFlags[i]) continue;

            var ancestor = allItems[i].parent;
            while (ancestor && ancestor.typename !== 'Layer') {
                for (var a = 0; a < allItems.length; a++) {
                    if (allItems[a] === ancestor) {
                        keepFlags[a] = true;
                        break;
                    }
                }
                ancestor = ancestor.parent;
            }
        }

        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();

        /* ドキュメントを複製 / Duplicate the document */
        var tempFile = makeTempFile(originalFile.parent, originalFile.name);
        if (!originalFile.copy(tempFile)) {
            alert(L("copyFailed"));
            return;
        }

        var duplicateDoc = app.open(tempFile);
        var duplicateItems = collectDocumentItems(duplicateDoc);

        /* 個数が違う＝対応付けが崩れているので、削除せずに中止する
           A different count means the mapping is broken; abort instead of deleting */
        if (duplicateItems.length !== allItems.length) {
            alert(L("structureError"));
            duplicateDoc.close(SaveOptions.DONOTSAVECHANGES);
            tempFile.remove();
            return;
        }

        unlockAllLayers(duplicateDoc.layers);

        /* 選択オブジェクト以外を削除。後ろ（深い側）から処理して、
           親を消したあとの無効な参照に触らないようにする
           Remove everything else, deepest first so removed parents are never revisited */
        var placed = [];
        for (i = duplicateItems.length - 1; i >= 0; i--) {
            if (keepFlags[i]) {
                if (selectedFlags[i]) placed.push(duplicateItems[i]);
                continue;
            }
            try {
                duplicateItems[i].locked = false;
                duplicateItems[i].hidden = false;
                duplicateItems[i].remove();
            } catch (_) {}
        }

        /* 現在のアートボード以外を削除。降順に処理するので、
           削除しても未処理側のインデックスはずれない
           Remove the other artboards; a descending loop keeps pending indexes valid */
        var duplicateArtboards = duplicateDoc.artboards;
        if (duplicateArtboards.length > 1) {
            for (i = duplicateArtboards.length - 1; i >= 0; i--) {
                if (i !== activeArtboardIndex) duplicateArtboards.remove(i);
            }
        }

        app.activeDocument = duplicateDoc;
        try {
            duplicateDoc.selection = placed;
        } catch (_) {}
        app.redraw();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログで作成対象を選び、対応する処理を実行します。
     *
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(L("noDocument"));
            return;
        }

        var doc = app.activeDocument;

        var options = showTargetDialog();
        if (options === null) return;

        if (options.target === "artboard") {
            createArtboardFromSelection(doc, options.duplicate);
        } else if (options.target === "document") {
            createDocumentFromSelection(doc);
        } else {
            createLayerFromSelection(doc, options.duplicate);
        }
    }

    main();

})();
