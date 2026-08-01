#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要

- 複数アートボードのドキュメントで、選択したオブジェクトを各アートボード上の指定位置へ整列します。
- 整列先は3×3の9点から選択でき、辺からのマージンを指定できます。
- ダイアログ表示中はライブプレビューで結果を確認でき、キャンセルすると元の位置に戻ります。
- 詳細な機能・オプションはREADMEを参照してください。

### Overview

- Aligns the selected objects to a chosen position on each artboard in a multi-artboard document.
- Pick one of nine anchor points in a 3x3 grid and set a margin from the artboard edges.
- Live preview while the dialog is open; Cancel restores the original positions.
- See the README for the full feature and option list.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AlignToArtboards";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-17";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-01";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AlignToArtboards.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AlignToArtboards.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n50aacdeb4908"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* パネル共通レイアウト / Common panel layout */
    var PANEL_MARGINS = [15, 20, 15, 10];
    /* 整列先パネルは左右・上下とも余白を詰める（9軸ウィジェットの余白調整）
       The anchor panel uses tighter padding on all sides to fit the 9-axis widget */
    var ANCHOR_PANEL_MARGINS = [9, 13, 9, 4];
    var MARGIN_FIELD_CHARACTERS = 4;

    /* 9軸ウィジェットの寸法（onDrawで描画） / 9-axis widget metrics (drawn via onDraw) */
    var ANCHOR_WIDGET_SIZE = 66;
    var ANCHOR_CELL_SIZE = 9;
    var ANCHOR_CELL_GAP = 7.5;

    /* ダイアログの初期値。必要に応じて編集 / Dialog defaults; edit as needed */
    var DEFAULT_SETTINGS = {
        useActiveArtboardAsReference: true, /* 「アクティブを基準」を初期選択 / Start in active-artboard reference mode */
        anchorCode: "LT",                   /* 初期の整列先 / Initial anchor */
        marginText: "0",                    /* マージン入力欄の初期値 / Initial margin input */
        linkMargins: true,                  /* 左右の値を上下にも反映 / Mirror horizontal margin to vertical */
        useVisibleBounds: true              /* プレビュー境界で整列 / Align by visible bounds */
    };

    /* 9点アンカーの定義。ratioX/ratioY は矩形内の相対位置（0=左/上, 0.5=中央, 1=右/下）
       Nine anchor points; ratioX/ratioY are relative positions inside a rectangle */
    var ANCHOR_DEFINITIONS = [
        { code: "LT", labelKey: "topLeft",      shortcutKey: "q", ratioX: 0,   ratioY: 0 },
        { code: "TC", labelKey: "topCenter",    shortcutKey: "w", ratioX: 0.5, ratioY: 0 },
        { code: "RT", labelKey: "topRight",     shortcutKey: "e", ratioX: 1,   ratioY: 0 },
        { code: "LM", labelKey: "middleLeft",   shortcutKey: "a", ratioX: 0,   ratioY: 0.5 },
        { code: "C",  labelKey: "center",       shortcutKey: "s", ratioX: 0.5, ratioY: 0.5 },
        { code: "RM", labelKey: "middleRight",  shortcutKey: "d", ratioX: 1,   ratioY: 0.5 },
        { code: "LB", labelKey: "bottomLeft",   shortcutKey: "z", ratioX: 0,   ratioY: 1 },
        { code: "BC", labelKey: "bottomCenter", shortcutKey: "x", ratioX: 0.5, ratioY: 1 },
        { code: "RB", labelKey: "bottomRight",  shortcutKey: "c", ratioX: 1,   ratioY: 1 }
    ];

    /* 中央アンカー（マージンを持たない） / Center anchor (no margin) */
    var CENTER_ANCHOR_CODE = "C";

    /* ルーラー単位コード → 単位ラベル / Ruler unit code to unit label */
    var RULER_UNIT_LABELS = {
        0: "in",
        1: "mm",
        2: "pt",
        3: "pica",
        4: "cm",
        5: "Q/H",
        6: "px",
        7: "ft/in",
        8: "m",
        9: "yd",
        10: "ft"
    };

    /* 単位ラベル → UnitValue に渡す単位名 / Ruler unit label to the unit name passed to UnitValue */
    var UNIT_VALUE_NAMES = {
        "in": "in",
        "mm": "mm",
        "pt": "pt",
        "pica": "pc",
        "cm": "cm",
        "px": "px",
        "m": "m",
        "yd": "yd",
        "ft": "ft",
        "ft/in": "in" /* フィートインチ表示のときは入力値をインチとして扱う / Treat input as inches when the ruler shows feet/inches */
    };

    /* 級（Q）・歯（H）は 1単位 = 0.25mm。UnitValue が扱えないので mm 経由で換算
       One Q or H equals 0.25 mm; UnitValue cannot parse it, so convert through mm */
    var MILLIMETERS_PER_Q = 0.25;
    var Q_UNIT_LABEL = "Q/H";

    /* マージンの最小値（内側へのオフセットのみ許可） / Minimum margin (inward offset only) */
    var MINIMUM_MARGIN_VALUE = 0;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

    var LABELS = {
        dialog: {
            title: { ja: "各アートボードに整列", en: "Align to Artboards" }
        },
        panel: {
            alignmentBase: { ja: "整列の基準", en: "Align based on" },
            anchor: { ja: "整列先", en: "Target" },
            margin: { ja: "マージン", en: "Margin" }
        },
        radio: {
            eachArtboard: { ja: "すべてのアートボード", en: "All Artboards" },
            activeArtboard: { ja: "アクティブなアートボード", en: "Based on Active Artboard" }
        },
        anchor: {
            topLeft: { ja: "左上", en: "Top-Left" },
            topCenter: { ja: "上中央", en: "Top-Center" },
            topRight: { ja: "右上", en: "Top-Right" },
            middleLeft: { ja: "左中央", en: "Middle-Left" },
            center: { ja: "中央", en: "Center" },
            middleRight: { ja: "右中央", en: "Middle-Right" },
            bottomLeft: { ja: "左下", en: "Bottom-Left" },
            bottomCenter: { ja: "下中央", en: "Bottom-Center" },
            bottomRight: { ja: "右下", en: "Bottom-Right" }
        },
        fieldLabel: {
            marginHorizontal: { ja: "左右", en: "Horizontal" },
            marginVertical: { ja: "上下", en: "Vertical" }
        },
        checkbox: {
            linkMargins: { ja: "連動", en: "Linked" },
            useVisibleBounds: { ja: "プレビュー境界を使用", en: "Use Preview Bounds" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            eachArtboard: {
                ja: "選択オブジェクトを中心点が属するアートボードごとに振り分け、各アートボード内の指定位置へ整列します。",
                en: "Groups selected objects by the artboard containing their center point, then aligns them to the selected position on each artboard."
            },
            activeArtboard: {
                ja: "アクティブアートボード上の選択を基準に、他のアートボード上の選択を同じ相対位置へ整列します。アクティブ側は動かしません。",
                en: "Uses the selection on the active artboard as the reference and aligns selections on other artboards to the same relative position. Objects on the active artboard are not moved."
            },
            anchor: {
                ja: "整列先の9点を選択します。アクティブなアートボードを基準にする場合は、その時点の選択が維持されます。",
                en: "Choose one of the 9 alignment positions. When using Based on Active Artboard, the current choice is preserved."
            },
            margin: {
                ja: "対応する辺から内側へオフセットします。中央、またはアクティブなアートボードを基準にする場合は無効です。",
                en: "Offsets objects inward from the corresponding edge. Disabled for Center and Based on Active Artboard."
            },
            linkMargins: {
                ja: "ONのときは左右の値を上下にも連動します。",
                en: "When enabled, the horizontal value is also used for the vertical margin."
            },
            useVisibleBounds: {
                ja: "ON：線幅や効果を含む見た目の境界で整列。OFF：図形本体の幾何境界で整列。",
                en: "On: align by visual bounds including strokes and effects. Off: align by geometric bounds of the object shape."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.anchor" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    /**
     * 表示言語に合わせたコロンを返す（日本語＝全角、英語＝半角）
     * @returns {string} コロン
     */
    function getLocalizedColon() {
        return (uiLang === "ja") ? "：" : ":";
    }

    // =========================================
    // 境界取得（プレビュー境界 or 幾何境界） / Bounds resolver
    // プレビュー境界＝ストローク・効果を含む、幾何境界＝図形のみ
    // Preview bounds include strokes/effects; geometric bounds use the object shape only
    // =========================================

    /* ダイアログのチェックボックスで切り替わる整列基準の境界 / Bounds type toggled from the dialog */
    var useVisibleBounds = DEFAULT_SETTINGS.useVisibleBounds;

    /**
     * アイテムの整列用バウンディングボックスを返す
     * @param {PageItem} item - 対象アイテム
     * @returns {number[]} [左, 上, 右, 下]。取得できない場合は null
     */
    function getItemBounds(item) {
        /* 文字編集中の TextRange など、境界を持たない選択は対象外 / Skip selections without bounds (e.g. TextRange) */
        try {
            return useVisibleBounds ? item.visibleBounds : item.geometricBounds;
        } catch (e) {
            return null;
        }
    }

    /**
     * 複数アイテムを包含する最小バウンディングを返す
     * @param {Array} items - 対象アイテムの配列
     * @returns {number[]} [左, 上, 右, 下]。取得できない場合は null
     */
    function getUnionBounds(items) {
        var unionBounds = null;
        for (var i = 0; i < items.length; i++) {
            var itemBounds = getItemBounds(items[i]);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    // =========================================
    // 単位（rulerType） / Units (rulerType)
    // =========================================

    /**
     * 現在のルーラー単位ラベルを取得する
     * @returns {string} 単位ラベル（mm, pt, px など）
     */
    function getCurrentUnitLabel() {
        var unitCode = app.preferences.getIntegerPreference("rulerType");
        return RULER_UNIT_LABELS[unitCode] || "pt";
    }

    /**
     * 入力値（現在のルーラー単位）を pt に変換する
     * @param {number} inputValue - 入力値
     * @param {string} unitLabel - 現在の単位ラベル
     * @returns {number} pt に換算した値
     */
    function convertToPoints(inputValue, unitLabel) {
        if (unitLabel === Q_UNIT_LABEL) {
            return new UnitValue(inputValue * MILLIMETERS_PER_Q, "mm").as("pt");
        }
        /* 対応表にない単位は pt として処理 / Units missing from the table are treated as pt */
        return new UnitValue(inputValue, UNIT_VALUE_NAMES[unitLabel] || "pt").as("pt");
    }

    // =========================================
    // アートボード振り分け / Artboard mapping
    // =========================================

    /**
     * アイテムまたは親階層がロック／非表示か判定する
     * @param {PageItem} item - 対象アイテム
     * @returns {boolean} ロックまたは非表示なら true
     */
    function isLockedOrHidden(item) {
        if (item.locked || item.hidden) return true;
        var container = item.parent;
        while (container && container.typename !== "Document") {
            if (container.typename === "Layer") {
                if (container.locked || !container.visible) return true;
            } else if (container.locked || container.hidden) {
                return true;
            }
            container = container.parent;
        }
        return false;
    }

    /**
     * 座標を含むアートボードのインデックスを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {number} pointX - X座標（pt）
     * @param {number} pointY - Y座標（pt）
     * @returns {number} アートボードのインデックス。該当なしは -1
     */
    function findArtboardIndexByPoint(doc, pointX, pointY) {
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardRect = doc.artboards[i].artboardRect; /* [左, 上, 右, 下] / [L, T, R, B] */
            if (pointX >= artboardRect[0] && pointX <= artboardRect[2] &&
                pointY <= artboardRect[1] && pointY >= artboardRect[3]) {
                return i;
            }
        }
        return -1;
    }

    /**
     * アイテムを中心点が属するアートボードごとに振り分ける
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - 振り分けるアイテムの配列
     * @returns {Object} アートボードインデックスをキーにしたアイテム配列
     */
    function groupItemsByArtboard(doc, items) {
        var itemsByArtboard = {};
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            var itemBounds = getItemBounds(item);
            if (!itemBounds) continue;
            if (isLockedOrHidden(item)) continue;

            var centerX = (itemBounds[0] + itemBounds[2]) / 2;
            var centerY = (itemBounds[1] + itemBounds[3]) / 2;
            var artboardIndex = findArtboardIndexByPoint(doc, centerX, centerY);
            if (artboardIndex === -1) continue;

            if (!itemsByArtboard[artboardIndex]) itemsByArtboard[artboardIndex] = [];
            itemsByArtboard[artboardIndex].push(item);
        }
        return itemsByArtboard;
    }

    // =========================================
    // プレビュー状態 / Preview state
    // =========================================

    /**
     * プレビュー状態を作成する（移動量を記録して巻き戻せるようにする）
     * @param {Array} items - プレビュー対象のアイテム
     * @returns {Object} プレビュー状態
     */
    function createPreviewState(items) {
        var previewState = { items: [], offsetsX: [], offsetsY: [] };
        for (var i = 0; i < items.length; i++) {
            previewState.items.push(items[i]);
            previewState.offsetsX.push(0);
            previewState.offsetsY.push(0);
        }
        return previewState;
    }

    /**
     * プレビュー状態にアイテムの移動量を積算する
     * @param {Object} previewState - プレビュー状態。null のときは何もしない
     * @param {PageItem} item - 移動したアイテム
     * @param {number} dx - X方向の移動量
     * @param {number} dy - Y方向の移動量
     * @returns {void}
     */
    function recordPreviewTranslation(previewState, item, dx, dy) {
        if (!previewState) return;
        for (var i = 0; i < previewState.items.length; i++) {
            if (previewState.items[i] !== item) continue;
            previewState.offsetsX[i] += dx;
            previewState.offsetsY[i] += dy;
            return;
        }
    }

    /**
     * プレビューの移動を巻き戻して元の位置に戻す
     * @param {Object} previewState - プレビュー状態
     * @returns {void}
     */
    function revertPreview(previewState) {
        for (var i = 0; i < previewState.items.length; i++) {
            var dx = previewState.offsetsX[i];
            var dy = previewState.offsetsY[i];
            if (dx === 0 && dy === 0) continue;
            translateItem(previewState.items[i], -dx, -dy, null);
            previewState.offsetsX[i] = 0;
            previewState.offsetsY[i] = 0;
        }
    }

    // =========================================
    // 整列処理 / Alignment
    // =========================================

    /* アンカーコードから定義を引くためのマップ / Lookup map from anchor code to definition */
    var anchorDefinitionByCode = {};
    for (var anchorIndex = 0; anchorIndex < ANCHOR_DEFINITIONS.length; anchorIndex++) {
        anchorDefinitionByCode[ANCHOR_DEFINITIONS[anchorIndex].code] = ANCHOR_DEFINITIONS[anchorIndex];
    }

    /**
     * 矩形上の9点アンカー座標を返す
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {string} anchorCode - アンカーコード（LT, TC, RT, LM, C, RM, LB, BC, RB）
     * @returns {number[]} [X座標, Y座標]
     */
    function getAnchorPoint(bounds, anchorCode) {
        var anchor = anchorDefinitionByCode[anchorCode] || anchorDefinitionByCode[DEFAULT_SETTINGS.anchorCode];
        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        return [left + (right - left) * anchor.ratioX, top - (top - bottom) * anchor.ratioY];
    }

    /**
     * 矩形を四辺から内側へ縮める（マージンの適用）
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {number} marginX - 左右のマージン（pt）
     * @param {number} marginY - 上下のマージン（pt）
     * @returns {number[]} 縮めた矩形 [左, 上, 右, 下]
     */
    function insetBounds(bounds, marginX, marginY) {
        return [bounds[0] + marginX, bounds[1] - marginY, bounds[2] - marginX, bounds[3] + marginY];
    }

    /**
     * アイテムを移動し、プレビュー中は移動量を記録する
     * @param {PageItem} item - 対象アイテム
     * @param {number} dx - X方向の移動量
     * @param {number} dy - Y方向の移動量
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function translateItem(item, dx, dy, previewState) {
        if (dx === 0 && dy === 0) return;
        /* クリッピングマスクなど、移動できないアイテムは無視 / Ignore items that cannot be moved */
        try {
            item.translate(dx, dy);
        } catch (e) {
            return;
        }
        recordPreviewTranslation(previewState, item, dx, dy);
    }

    /**
     * アイテム群をまとめて同じ量だけ移動する
     * @param {Array} items - 対象アイテムの配列
     * @param {number} dx - X方向の移動量
     * @param {number} dy - Y方向の移動量
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function translateItems(items, dx, dy, previewState) {
        for (var i = 0; i < items.length; i++) {
            translateItem(items[i], dx, dy, previewState);
        }
    }

    /**
     * 各アイテムを個別に、アートボード上のアンカーへ整列する
     * @param {Array} items - 対象アイテムの配列
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {string} anchorCode - アンカーコード
     * @param {number} marginX - 左右のマージン（pt）
     * @param {number} marginY - 上下のマージン（pt）
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function alignItemsToArtboardAnchor(items, artboardRect, anchorCode, marginX, marginY, previewState) {
        var targetPoint = getAnchorPoint(insetBounds(artboardRect, marginX, marginY), anchorCode);
        for (var i = 0; i < items.length; i++) {
            var itemBounds = getItemBounds(items[i]);
            if (!itemBounds) continue;
            var itemAnchor = getAnchorPoint(itemBounds, anchorCode);
            translateItem(items[i], targetPoint[0] - itemAnchor[0], targetPoint[1] - itemAnchor[1], previewState);
        }
    }

    /**
     * アイテム群の相対位置を保ったまま、アートボード上のアンカー＋オフセット位置へ整列する
     * @param {Array} items - 対象アイテムの配列
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @param {string} anchorCode - アンカーコード
     * @param {Object} offset - アンカーからの相対位置 {x, y}
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function alignItemGroupToArtboardAnchor(items, artboardRect, anchorCode, offset, previewState) {
        var groupBounds = getUnionBounds(items);
        if (!groupBounds) return;
        var artboardAnchor = getAnchorPoint(artboardRect, anchorCode);
        var groupAnchor = getAnchorPoint(groupBounds, anchorCode);
        translateItems(
            items,
            artboardAnchor[0] + offset.x - groupAnchor[0],
            artboardAnchor[1] + offset.y - groupAnchor[1],
            previewState
        );
    }

    /**
     * アクティブアートボード上の選択を基準に、他のアートボードの選択を同じ相対位置へ整列する
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} itemsByArtboard - アートボードごとに振り分けたアイテム
     * @param {string} anchorCode - アンカーコード
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function alignUsingActiveArtboardReference(doc, itemsByArtboard, anchorCode, previewState) {
        var activeArtboardIndex = doc.artboards.getActiveArtboardIndex();
        if (activeArtboardIndex < 0) return;

        var referenceItems = itemsByArtboard[activeArtboardIndex];
        if (!referenceItems) return;

        var referenceBounds = getUnionBounds(referenceItems);
        if (!referenceBounds) return;

        /* 基準アートボードのアンカーから見た、基準オブジェクト群の相対位置
           Relative position of the reference group as seen from the artboard anchor */
        var referenceAnchor = getAnchorPoint(doc.artboards[activeArtboardIndex].artboardRect, anchorCode);
        var groupAnchor = getAnchorPoint(referenceBounds, anchorCode);
        var offset = { x: groupAnchor[0] - referenceAnchor[0], y: groupAnchor[1] - referenceAnchor[1] };

        for (var artboardIndex in itemsByArtboard) {
            if (!itemsByArtboard.hasOwnProperty(artboardIndex)) continue;
            var targetIndex = parseInt(artboardIndex, 10);
            if (targetIndex === activeArtboardIndex) continue;
            alignItemGroupToArtboardAnchor(
                itemsByArtboard[artboardIndex],
                doc.artboards[targetIndex].artboardRect,
                anchorCode,
                offset,
                previewState
            );
        }
    }

    /**
     * 設定に従って整列を実行する（プレビューと確定で共用）
     * @param {Document} doc - 対象ドキュメント
     * @param {Array} items - 整列するアイテムの配列
     * @param {Object} settings - 整列設定 {anchorCode, useActiveArtboardAsReference, marginX, marginY}
     * @param {Object} previewState - プレビュー状態。確定時は null
     * @returns {void}
     */
    function alignItems(doc, items, settings, previewState) {
        var itemsByArtboard = groupItemsByArtboard(doc, items);

        if (settings.useActiveArtboardAsReference) {
            alignUsingActiveArtboardReference(doc, itemsByArtboard, settings.anchorCode, previewState);
            return;
        }

        for (var artboardIndex in itemsByArtboard) {
            if (!itemsByArtboard.hasOwnProperty(artboardIndex)) continue;
            alignItemsToArtboardAnchor(
                itemsByArtboard[artboardIndex],
                doc.artboards[parseInt(artboardIndex, 10)].artboardRect,
                settings.anchorCode,
                settings.marginX,
                settings.marginY,
                previewState
            );
        }
    }

    // =========================================
    // 入力ユーティリティ / Input utilities
    // =========================================

    /**
     * 文字列を数値に変換する（不正値は既定値）
     * @param {string} inputText - 入力文字列
     * @param {number} defaultValue - 変換できないときに返す値
     * @returns {number} 変換した数値
     */
    function parseNumericInput(inputText, defaultValue) {
        var trimmedText = ("" + inputText).replace(/^\s+|\s+$/g, "");
        if (trimmedText === "") return defaultValue;
        var parsedValue = Number(trimmedText);
        return isNaN(parsedValue) ? defaultValue : parsedValue;
    }

    /**
     * 数値入力欄を上下矢印キーで増減できるようにする（Shift=10単位、Option=0.1単位）
     * @param {EditText} numberField - 対象の入力欄
     * @param {number} minimumValue - 下限値
     * @param {Function} onValueChanged - 値が変わったときに呼ぶ関数
     * @returns {void}
     */
    function enableArrowKeyStepping(numberField, minimumValue, onValueChanged) {
        numberField.addEventListener("keydown", function (event) {
            /* Up / Down 以外は素通り（数値入力を妨げない） / Pass through non-arrow keys */
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var currentValue = parseNumericInput(numberField.text, 0);
            var stepDirection = (event.keyName === "Up") ? 1 : -1;
            var keyboardState = ScriptUI.environment.keyboardState;

            if (keyboardState.shiftKey) {
                /* Shift：10の倍数にスナップ / Snap to multiples of 10 */
                currentValue = (stepDirection > 0)
                    ? Math.ceil((currentValue + 1) / 10) * 10
                    : Math.floor((currentValue - 1) / 10) * 10;
            } else if (keyboardState.altKey) {
                /* Option：0.1単位 / Increment by 0.1 */
                currentValue = Math.round((currentValue + stepDirection * 0.1) * 10) / 10;
            } else {
                /* 通常：1単位、整数丸め / Default: by 1, integer */
                currentValue = Math.round(currentValue + stepDirection);
            }

            numberField.text = (currentValue < minimumValue) ? minimumValue : currentValue;
            onValueChanged();
            event.preventDefault();
        });
    }

    // =========================================
    // 9軸ウィジェットの描画 / Anchor widget drawing
    // =========================================

    /* 外周の□どうしをつなぐケイ線の組み合わせ（中央は独立） / Rules joining the outer squares (center stands alone) */
    var ANCHOR_CONNECTIONS = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];

    /* UIの明暗に合わせて initAnchorColors() で設定 / Set from the light/dark UI in initAnchorColors() */
    var anchorLineColor = [0.6, 0.6, 0.6, 1];
    var anchorSelectedFillColor = [0.4, 0.4, 0.4, 1];
    var anchorDimmedLineColor = [0.8, 0.8, 0.8, 1];

    /**
     * IllustratorのUIが明るいテーマか判定する
     * @returns {boolean} 明るいテーマなら true
     */
    function isLightUI() {
        return app.preferences.getRealPreference("uiBrightness") > 0.5;
    }

    /**
     * 9軸ウィジェットの配色をUIの明暗に合わせて決める
     * @returns {void}
     */
    function initAnchorColors() {
        var lightUI = isLightUI();
        /* 選択セルの塗り：ライトは濃いグレー、ダークは明るいグレー / Selected-cell fill: dark gray in light UI, bright gray in dark UI */
        anchorLineColor = lightUI ? [0.6, 0.6, 0.6, 1] : [0.55, 0.55, 0.55, 1];
        anchorSelectedFillColor = lightUI ? [0.4, 0.4, 0.4, 1] : [0.8, 0.8, 0.8, 1];
        anchorDimmedLineColor = lightUI ? [0.8, 0.8, 0.8, 1] : [0.38, 0.38, 0.38, 1];
    }

    /**
     * 3×3のグリッド位置を 0〜2 に収める
     * @param {number} gridPosition - 計算した行または列
     * @returns {number} 0〜2 に丸めた値
     */
    function clampGridIndex(gridPosition) {
        if (gridPosition < 0) return 0;
        if (gridPosition > 2) return 2;
        return gridPosition;
    }

    /**
     * アンカーコードから ANCHOR_DEFINITIONS 上のインデックスを求める
     * @param {string} anchorCode - アンカーコード
     * @returns {number} インデックス。見つからない場合は 0
     */
    function findAnchorIndexByCode(anchorCode) {
        for (var i = 0; i < ANCHOR_DEFINITIONS.length; i++) {
            if (ANCHOR_DEFINITIONS[i].code === anchorCode) return i;
        }
        return 0;
    }

    /**
     * 9軸ウィジェットを再描画する
     * @param {Button} anchorWidget - 対象ウィジェット
     * @returns {void}
     */
    function redrawAnchorWidget(anchorWidget) {
        /* notify は環境により例外を投げ得るので保護 / notify can throw in some environments */
        try {
            anchorWidget.notify("onDraw");
        } catch (e) { }
    }

    /**
     * 9軸ウィジェットを描画する（外周の□をケイ線でつなぎ、選択セルを塗る）
     * @param {Button} anchorWidget - 対象ウィジェット
     * @returns {void}
     */
    function drawAnchorWidget(anchorWidget) {
        var graphics = anchorWidget.graphics;
        var widgetWidth = anchorWidget.size[0];
        var widgetHeight = anchorWidget.size[1];
        var lineColor = anchorWidget.enabled ? anchorLineColor : anchorDimmedLineColor;

        /* 背景はコントロールの地色で塗り、パネルと同色に見せる / Paint the control background so the widget blends into the panel */
        try {
            graphics.rectPath(0, 0, widgetWidth, widgetHeight);
            graphics.fillPath(graphics.backgroundColor);
        } catch (e) { }

        var cellStep = ANCHOR_CELL_SIZE + ANCHOR_CELL_GAP;
        var gridSize = ANCHOR_CELL_SIZE * 3 + ANCHOR_CELL_GAP * 2;
        var originX = Math.round((widgetWidth - gridSize) / 2);
        var originY = Math.round((widgetHeight - gridSize) / 2);

        /**
         * セルの左上座標を返す
         * @param {number} cellIndex - 0〜8のセル番号（行優先）
         * @returns {number[]} [X座標, Y座標]
         */
        function getCellOrigin(cellIndex) {
            return [
                originX + (cellIndex % 3) * cellStep,
                originY + Math.floor(cellIndex / 3) * cellStep
            ];
        }

        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1);
        for (var i = 0; i < ANCHOR_CONNECTIONS.length; i++) {
            var fromCell = getCellOrigin(ANCHOR_CONNECTIONS[i][0]);
            var toCell = getCellOrigin(ANCHOR_CONNECTIONS[i][1]);
            var isHorizontal = (ANCHOR_CONNECTIONS[i][1] - ANCHOR_CONNECTIONS[i][0] === 1);
            graphics.newPath();
            if (isHorizontal) {
                /* 横方向：右隣の□へ / Horizontal: to the square on the right */
                graphics.moveTo(fromCell[0] + ANCHOR_CELL_SIZE, fromCell[1] + ANCHOR_CELL_SIZE / 2);
                graphics.lineTo(toCell[0], toCell[1] + ANCHOR_CELL_SIZE / 2);
            } else {
                /* 縦方向：下の□へ / Vertical: to the square below */
                graphics.moveTo(fromCell[0] + ANCHOR_CELL_SIZE / 2, fromCell[1] + ANCHOR_CELL_SIZE);
                graphics.lineTo(toCell[0] + ANCHOR_CELL_SIZE / 2, toCell[1]);
            }
            graphics.strokePath(linePen);
        }

        for (var cellIndex = 0; cellIndex < ANCHOR_DEFINITIONS.length; cellIndex++) {
            var cellOrigin = getCellOrigin(cellIndex);
            drawAnchorCell(
                graphics,
                cellOrigin[0],
                cellOrigin[1],
                cellIndex === anchorWidget.selectedAnchorIndex,
                lineColor
            );
        }
    }

    /**
     * 9軸ウィジェットのセルを1つ描画する（選択時のみ塗りつぶす）
     * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
     * @param {number} cellX - セルの左端
     * @param {number} cellY - セルの上端
     * @param {boolean} isSelected - 選択中なら true
     * @param {number[]} lineColor - 枠線の色
     * @returns {void}
     */
    function drawAnchorCell(graphics, cellX, cellY, isSelected, lineColor) {
        /**
         * セルの四角形パスを作る
         * @returns {void}
         */
        function addCellPath() {
            graphics.newPath();
            graphics.moveTo(cellX, cellY);
            graphics.lineTo(cellX + ANCHOR_CELL_SIZE, cellY);
            graphics.lineTo(cellX + ANCHOR_CELL_SIZE, cellY + ANCHOR_CELL_SIZE);
            graphics.lineTo(cellX, cellY + ANCHOR_CELL_SIZE);
            graphics.closePath();
        }

        /* 選択中のみ塗る（枠は塗りの上に描く） / Fill only when selected; the border is drawn over the fill */
        if (isSelected) {
            addCellPath();
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, anchorSelectedFillColor));
        }
        addCellPath();
        graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, lineColor, 1));
    }

    // =========================================
    // ダイアログ UI / Dialog UI
    // =========================================

    /**
     * パネルの共通レイアウトを適用する
     * @param {Panel} panel - 対象パネル
     * @param {string} orientation - "column" または "row"
     * @returns {void}
     */
    function applyPanelLayout(panel, orientation) {
        panel.orientation = orientation;
        panel.margins = PANEL_MARGINS;
    }

    /**
     * 整列の基準パネル（すべてのアートボード／アクティブを基準）を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @param {Function} onSettingsChanged - 選択が変わったときに呼ぶ関数
     * @returns {Object} 基準の取得・設定用インターフェース
     */
    function buildAlignmentBasePanel(parentContainer, onSettingsChanged) {
        var basePanel = parentContainer.add("panel", undefined, getLabel("panel.alignmentBase"));
        applyPanelLayout(basePanel, "column");
        basePanel.alignChildren = ["left", "center"];

        var eachArtboardRadio = basePanel.add("radiobutton", undefined, getLabel("radio.eachArtboard"));
        var activeArtboardRadio = basePanel.add("radiobutton", undefined, getLabel("radio.activeArtboard"));
        eachArtboardRadio.helpTip = getLabel("tooltip.eachArtboard");
        activeArtboardRadio.helpTip = getLabel("tooltip.activeArtboard");

        /**
         * 整列の基準を切り替える
         * @param {boolean} useActiveArtboard - アクティブを基準にするなら true
         * @returns {void}
         */
        function selectAlignmentBase(useActiveArtboard) {
            eachArtboardRadio.value = !useActiveArtboard;
            activeArtboardRadio.value = useActiveArtboard;
        }

        selectAlignmentBase(DEFAULT_SETTINGS.useActiveArtboardAsReference);
        eachArtboardRadio.onClick = onSettingsChanged;
        activeArtboardRadio.onClick = onSettingsChanged;

        return {
            selectAlignmentBase: selectAlignmentBase,
            usesActiveArtboard: function () {
                return activeArtboardRadio.value === true;
            }
        };
    }

    /**
     * 整列先パネル（3×3の9軸ウィジェット）を構築する
     * @param {Group} parentContainer - 追加先のグループ
     * @param {Function} onSettingsChanged - 選択が変わったときに呼ぶ関数
     * @returns {Object} パネル参照とアンカー取得・ショートカット処理をまとめたオブジェクト
     */
    function buildAnchorPanel(parentContainer, onSettingsChanged) {
        var anchorPanel = parentContainer.add("panel", undefined, getLabel("panel.anchor"));
        applyPanelLayout(anchorPanel, "column");
        anchorPanel.margins = ANCHOR_PANEL_MARGINS;
        anchorPanel.alignChildren = ["center", "top"];

        var anchorWidget = anchorPanel.add("button", undefined, "");
        anchorWidget.preferredSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.minimumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.maximumSize = [ANCHOR_WIDGET_SIZE, ANCHOR_WIDGET_SIZE];
        anchorWidget.selectedAnchorIndex = findAnchorIndexByCode(DEFAULT_SETTINGS.anchorCode);
        anchorWidget.onDraw = function () {
            drawAnchorWidget(this);
        };

        /**
         * 指定したセルを選択し、ウィジェットとツールチップを更新する
         * @param {number} selectedIndex - ANCHOR_DEFINITIONS 上のインデックス
         * @returns {void}
         */
        function selectAnchorAt(selectedIndex) {
            var anchorDefinition = ANCHOR_DEFINITIONS[selectedIndex];
            anchorWidget.selectedAnchorIndex = selectedIndex;
            anchorWidget.helpTip = getLabel("tooltip.anchor") + "\n" +
                getLabel("anchor." + anchorDefinition.labelKey) + " (" + anchorDefinition.shortcutKey + ")";
            anchorPanel.helpTip = anchorWidget.helpTip;
            redrawAnchorWidget(anchorWidget);
        }

        /* クリックした3×3のセルを整列先にする（座標はウィジェット基準）
           Set the anchor from the clicked 3x3 cell (coordinates are widget-relative) */
        anchorWidget.addEventListener("mousedown", function (event) {
            var column = clampGridIndex(Math.floor(event.clientX / (anchorWidget.size[0] / 3)));
            var row = clampGridIndex(Math.floor(event.clientY / (anchorWidget.size[1] / 3)));
            selectAnchorAt(row * 3 + column);
            onSettingsChanged();
        });

        selectAnchorAt(anchorWidget.selectedAnchorIndex);

        return {
            isEnabled: function () {
                return anchorWidget.enabled === true;
            },
            setEnabled: function (isEnabled) {
                /* パネルとウィジェットの両方を切り替える（描画の淡色化はウィジェット側を見る）
                   Toggle both the panel and the widget; the dimmed drawing follows the widget state */
                anchorPanel.enabled = isEnabled;
                anchorWidget.enabled = isEnabled;
                redrawAnchorWidget(anchorWidget);
            },
            getAnchorCode: function () {
                return ANCHOR_DEFINITIONS[anchorWidget.selectedAnchorIndex].code;
            },
            selectByShortcutKey: function (pressedKey) {
                for (var i = 0; i < ANCHOR_DEFINITIONS.length; i++) {
                    if (ANCHOR_DEFINITIONS[i].shortcutKey !== pressedKey) continue;
                    selectAnchorAt(i);
                    return true;
                }
                return false;
            }
        };
    }

    /**
     * マージンパネル（左右・上下・連動）を構築する
     * @param {Group} parentContainer - 追加先のグループ
     * @param {string} unitLabel - 現在のルーラー単位ラベル
     * @param {Function} onSettingsChanged - 入力が変わったときに呼ぶ関数
     * @returns {Object} パネル参照とマージン取得・フォーカス制御をまとめたオブジェクト
     */
    function buildMarginPanel(parentContainer, unitLabel, onSettingsChanged) {
        var marginPanel = parentContainer.add("panel", undefined, getLabel("panel.margin") + " (" + unitLabel + ")");
        applyPanelLayout(marginPanel, "row");
        marginPanel.alignChildren = ["fill", "center"];
        marginPanel.helpTip = getLabel("tooltip.margin");

        var fieldColumn = marginPanel.add("group");
        fieldColumn.orientation = "column";
        fieldColumn.alignChildren = ["left", "center"];

        /**
         * ラベル付きのマージン入力欄を作る
         * @param {string} fieldLabelPath - ラベルの LABELS パス
         * @returns {EditText} 追加した入力欄
         */
        function addMarginField(fieldLabelPath) {
            var fieldGroup = fieldColumn.add("group");
            fieldGroup.orientation = "row";
            fieldGroup.alignChildren = ["left", "center"];
            var fieldLabel = fieldGroup.add("statictext", undefined, getLabel(fieldLabelPath) + getLocalizedColon());
            fieldLabel.helpTip = marginPanel.helpTip;
            var marginField = fieldGroup.add("edittext", undefined, DEFAULT_SETTINGS.marginText);
            marginField.characters = MARGIN_FIELD_CHARACTERS;
            marginField.helpTip = marginPanel.helpTip;
            return marginField;
        }

        var horizontalField = addMarginField("fieldLabel.marginHorizontal");
        var verticalField = addMarginField("fieldLabel.marginVertical");

        var linkColumn = marginPanel.add("group");
        linkColumn.orientation = "column";
        linkColumn.alignChildren = ["left", "center"];
        var linkCheckbox = linkColumn.add("checkbox", undefined, getLabel("checkbox.linkMargins"));
        linkCheckbox.value = DEFAULT_SETTINGS.linkMargins;
        linkCheckbox.helpTip = getLabel("tooltip.linkMargins");

        /**
         * 入力欄の値を読み取る（マージンは内側へのオフセットなので負値は0として扱う）
         * @param {EditText} marginField - 対象の入力欄
         * @returns {number} 0以上の入力値
         */
        function readMarginField(marginField) {
            var marginValue = parseNumericInput(marginField.text, 0);
            return (marginValue < MINIMUM_MARGIN_VALUE) ? MINIMUM_MARGIN_VALUE : marginValue;
        }

        /**
         * 入力確定時に負値を下限へ丸めて表示に反映する
         * @param {EditText} marginField - 対象の入力欄
         * @returns {void}
         */
        function normalizeMarginField(marginField) {
            var marginValue = readMarginField(marginField);
            if (marginValue !== parseNumericInput(marginField.text, 0)) marginField.text = marginValue;
        }

        /**
         * 連動ONのときは左右の値を上下へ反映し、上下の入力欄を無効化する
         * @returns {void}
         */
        function syncLinkedMargin() {
            verticalField.enabled = !linkCheckbox.value;
            if (linkCheckbox.value) verticalField.text = horizontalField.text;
        }

        /**
         * 左右の入力が変わったときの処理
         * @returns {void}
         */
        function handleHorizontalMarginChanged() {
            syncLinkedMargin();
            onSettingsChanged();
        }

        /**
         * 上下の入力が変わったときの処理（連動ONのときは左右側で処理済み）
         * @returns {void}
         */
        function handleVerticalMarginChanged() {
            if (!linkCheckbox.value) onSettingsChanged();
        }

        horizontalField.onChanging = handleHorizontalMarginChanged;
        horizontalField.onChange = function () {
            normalizeMarginField(horizontalField);
            handleHorizontalMarginChanged();
        };
        verticalField.onChanging = handleVerticalMarginChanged;
        verticalField.onChange = function () {
            normalizeMarginField(verticalField);
            handleVerticalMarginChanged();
        };
        linkCheckbox.onClick = handleHorizontalMarginChanged;
        enableArrowKeyStepping(horizontalField, MINIMUM_MARGIN_VALUE, handleHorizontalMarginChanged);
        enableArrowKeyStepping(verticalField, MINIMUM_MARGIN_VALUE, handleVerticalMarginChanged);
        syncLinkedMargin();

        return {
            panel: marginPanel,
            linkCheckbox: linkCheckbox,
            getMarginInPoints: function () {
                var horizontalValue = readMarginField(horizontalField);
                var verticalValue = linkCheckbox.value ? horizontalValue : readMarginField(verticalField);
                return {
                    x: convertToPoints(horizontalValue, unitLabel),
                    y: convertToPoints(verticalValue, unitLabel)
                };
            },
            focusHorizontalField: function () {
                horizontalField.active = true;
                horizontalField.selection = [0, horizontalField.text.length];
            },
            bindEnterKey: function (okButton) {
                var marginFields = [horizontalField, verticalField];
                for (var i = 0; i < marginFields.length; i++) {
                    marginFields[i].addEventListener("keydown", function (event) {
                        if (event.keyName !== "Enter" && event.keyName !== "Return") return;
                        okButton.notify();
                        event.preventDefault();
                    });
                }
            }
        };
    }

    /**
     * プレビュー境界のチェックボックス行を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @param {Function} onSettingsChanged - 切り替え時に呼ぶ関数
     * @returns {void}
     */
    function buildBoundsOptionRow(parentContainer, onSettingsChanged) {
        var optionGroup = parentContainer.add("group");
        optionGroup.orientation = "row";
        optionGroup.alignment = ["left", "top"];
        optionGroup.margins = [4, 4, 4, 4];

        var boundsCheckbox = optionGroup.add("checkbox", undefined, getLabel("checkbox.useVisibleBounds"));
        boundsCheckbox.value = useVisibleBounds;
        boundsCheckbox.helpTip = getLabel("tooltip.useVisibleBounds");
        boundsCheckbox.onClick = function () {
            useVisibleBounds = boundsCheckbox.value;
            onSettingsChanged();
        };
    }

    /**
     * OK／キャンセルのボタン行を構築する
     * @param {Window} parentContainer - 追加先のダイアログ
     * @returns {Button} OKボタン
     */
    function buildDialogButtonRow(parentContainer) {
        var buttonGroup = parentContainer.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["center", "center"];
        buttonGroup.alignment = ["center", "bottom"];

        buttonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        return buttonGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    }

    /**
     * ダイアログのショートカットキー（1/2＝整列の基準、q〜c＝整列先）を登録する
     * @param {Window} dialog - 対象ダイアログ
     * @param {Object} baseControls - 整列の基準パネルのインターフェース
     * @param {Object} anchorControls - 整列先パネルのインターフェース
     * @param {Function} onSettingsChanged - 選択が変わったときに呼ぶ関数
     * @returns {void}
     */
    function registerShortcutKeys(dialog, baseControls, anchorControls, onSettingsChanged) {
        dialog.addEventListener("keydown", function (event) {
            /* テキスト入力中はショートカットを無効化 / Skip shortcuts while typing in edittext */
            if (event.target && event.target.type === "edittext") return;
            if (!event.keyName) return;

            var pressedKey = ("" + event.keyName).toLowerCase();
            if (pressedKey === "1" || pressedKey === "2") {
                baseControls.selectAlignmentBase(pressedKey === "2");
                event.preventDefault();
                onSettingsChanged();
                return;
            }

            /* 整列先が無効のときは9軸のショートカットも無効 / Skip anchor shortcuts when the widget is disabled */
            if (!anchorControls.isEnabled()) return;
            if (!anchorControls.selectByShortcutKey(pressedKey)) return;
            event.preventDefault();
            onSettingsChanged();
        });
    }

    /**
     * ダイアログを構築・表示し、ライブプレビューで整列を反映する
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function showAlignmentDialog(doc) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];

        var unitLabel = getCurrentUnitLabel();
        var previewState = createPreviewState(doc.selection);
        initAnchorColors();

        var baseControls = buildAlignmentBasePanel(dialog, handleSettingsChanged);

        var contentGroup = dialog.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = ["fill", "top"];

        var anchorColumn = contentGroup.add("group");
        anchorColumn.orientation = "column";
        anchorColumn.alignChildren = ["fill", "top"];
        var anchorControls = buildAnchorPanel(anchorColumn, handleSettingsChanged);

        var marginColumn = contentGroup.add("group");
        marginColumn.orientation = "column";
        marginColumn.alignChildren = ["fill", "top"];
        marginColumn.alignment = ["right", "top"];
        var marginControls = buildMarginPanel(marginColumn, unitLabel, handleSettingsChanged);

        buildBoundsOptionRow(dialog, handleSettingsChanged);
        var okButton = buildDialogButtonRow(dialog);
        dialog.defaultElement = okButton;
        marginControls.bindEnterKey(okButton);
        registerShortcutKeys(dialog, baseControls, anchorControls, handleSettingsChanged);

        /**
         * ダイアログの入力から整列設定を読み取る
         * @returns {Object} 整列設定 {anchorCode, useActiveArtboardAsReference, marginX, marginY}
         */
        function readSettingsFromDialog() {
            var anchorCode = anchorControls.getAnchorCode();
            var useActiveArtboardAsReference = baseControls.usesActiveArtboard();
            /* 中央整列時とアクティブ基準時はマージンを使わない / Margins are unused for center and active-artboard modes */
            var marginIsAvailable = (anchorCode !== CENTER_ANCHOR_CODE) && !useActiveArtboardAsReference;
            var margin = marginIsAvailable ? marginControls.getMarginInPoints() : { x: 0, y: 0 };
            return {
                anchorCode: anchorCode,
                useActiveArtboardAsReference: useActiveArtboardAsReference,
                marginIsAvailable: marginIsAvailable,
                marginX: margin.x,
                marginY: margin.y
            };
        }

        /**
         * 入力が変わるたびにパネルの有効状態を同期し、プレビューを再適用する
         * @returns {void}
         */
        function handleSettingsChanged() {
            var settings = readSettingsFromDialog();
            anchorControls.setEnabled(!settings.useActiveArtboardAsReference);
            marginControls.panel.enabled = settings.marginIsAvailable;
            marginControls.linkCheckbox.enabled = settings.marginIsAvailable;

            revertPreview(previewState);
            alignItems(doc, previewState.items, settings, previewState);
            app.redraw();
        }

        dialog.onShow = function () {
            marginControls.focusHorizontalField();
        };

        /* 初期状態でプレビューを適用 / Apply the initial preview */
        handleSettingsChanged();

        /* OKのときはプレビューの位置をそのまま確定 / OK keeps the previewed positions as the result */
        if (dialog.show() === 1) return;

        revertPreview(previewState);
        app.redraw();
    }

    // =========================================
    // メイン / Main
    // =========================================

    if (app.documents.length === 0) return;

    var activeDocument = app.activeDocument;
    if (!activeDocument.selection || activeDocument.selection.length === 0) return;

    showAlignmentDialog(activeDocument);

})();
