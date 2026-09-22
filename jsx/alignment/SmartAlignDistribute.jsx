#target illustrator
#targetengine "SmartAlignAndTileEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを縦または横に並べ、指定した間隔で分布します。
方向は自動判定でき、揃え（左右／上下）、プレビュー境界、ランダム並べ替えにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignDistribute.md

### Overview

Lines the selected objects up vertically or horizontally and distributes them at the spacing you specify.
The direction can be detected automatically, and alignment, preview bounds and random reordering are all supported.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignDistribute.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartAlignDistribute";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartAlignDistribute.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartAlignDistribute.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プレビューを再描画する最小間隔（ミリ秒）/ Minimum interval between preview renders (ms) */
    var PREVIEW_MIN_INTERVAL_MS = 80;

    /* 操作からプレビューを実行するまでの待ち時間（ミリ秒）/ Delay before a requested preview runs (ms) */
    var PREVIEW_SCHEDULE_MS = 60;

    // =========================================
    // 内部キー / Internal keys
    // =========================================

    /* ダイアログ位置をセッション内に控える $.global のキー / $.global key for the session dialog position */
    var DIALOG_POSITION_KEY = "__SmartAlignDistribute_DialogPosition__";

    /* 計測用に一時的に作るグループの名前 / name of the throwaway measuring group */
    var TEMP_MEASURE_GROUP_NAME = "__SmartAlignDistribute_TempMeasure__";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OPACITY = 0.97;               /* ダイアログの不透明度 / dialog opacity */
    var PANEL_MARGINS = [15, 20, 15, 10];    /* パネルの余白 [左,上,右,下] / panel margins */
    var OPTIONS_MARGINS = [15, 5, 15, 5];    /* オプション欄の余白 / options group margins */
    var SPACING_FIELD_CHARS = 3;             /* 間隔の入力欄の幅（文字数）/ width of the spacing field */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "整列と分布", en: "Align & Distribute" }
        },
        panel: {
            direction: { ja: "方向", en: "Direction" },
            spacing: { ja: "間隔", en: "Spacing" },
            alignHorizontal: { ja: "揃え（左右）", en: "Align (H)" },
            alignVertical: { ja: "揃え（上下）", en: "Align (V)" }
        },
        radio: {
            directionAuto: { ja: "自動", en: "Auto" },
            directionVertical: { ja: "縦", en: "Vertical" },
            directionHorizontal: { ja: "横", en: "Horizontal" },
            alignNone: { ja: "なし", en: "None" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            alignTop: { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" }
        },
        checkbox: {
            usePreviewBounds: { ja: "プレビュー境界を使用", en: "Use preview bounds" },
            measureText: { ja: "テキストの高さを計測", en: "Measure text height" },
            random: { ja: "ランダム", en: "Random" }
        },
        tooltip: {
            directionAuto: {
                ja: "選択範囲が横長なら横並び、縦長なら縦並びとして扱います。",
                en: "Lays the objects out in a row when the selection is wider than tall, in a column otherwise."
            },
            directionVertical: { ja: "上から下へ縦に並べます。", en: "Stacks the objects from top to bottom." },
            directionHorizontal: { ja: "左から右へ横に並べます。", en: "Lays the objects out from left to right." },
            spacing: {
                ja: "オブジェクト間のすき間。マイナス値で重ねられます。",
                en: "Gap between objects. Negative values overlap them."
            },
            alignHorizontal: {
                ja: "縦に並べたときの左右の揃え方です。N／L／C／R キーでも切り替えられます。",
                en: "Horizontal alignment used when stacking vertically. The keys N / L / C / R switch it."
            },
            alignVertical: {
                ja: "横に並べたときの上下の揃え方です。N／T／M／B キーでも切り替えられます。",
                en: "Vertical alignment used when laying out horizontally. The keys N / T / M / B switch it."
            },
            usePreviewBounds: {
                ja: "線や効果を含む見た目の境界で整列します。オフはパスのみのジオメトリ境界。",
                en: "Align by visible bounds (incl. strokes/effects). Off uses geometric (path-only) bounds."
            },
            measureText: {
                ja: "縦並び時のみ、テキストを一度だけ複製→アウトライン化して境界を計測します（ダイアログ中だけキャッシュ）。",
                en: "Only in vertical layout, measures text by duplicating and outlining once (cached for this dialog only)."
            },
            random: {
                ja: "並び順をランダムに入れ替えます（左上の位置は維持）。",
                en: "Shuffle the stacking order at random (top-left position is kept)."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            errorPrefix: { ja: "エラーが発生しました: ", en: "An error has occurred: " }
        }
    };

    /**
     * ドット区切りのパスで LABELS から現在の言語の文字列を取り出す（{slash} は "/" に置き換える）
     * @param {string} labelPath - "panel.direction" のようなドット区切りのキー
     * @returns {string} 現在の言語の文字列（見つからなければ labelPath そのもの）
     */
    function getLabel(labelPath) {
        var labelNode = LABELS;
        var pathKeys = labelPath.split(".");
        for (var i = 0; i < pathKeys.length; i++) {
            if (labelNode == null) return labelPath;
            labelNode = labelNode[pathKeys[i]];
        }
        if (labelNode == null) return labelPath;
        var localizedText = (labelNode[uiLang] != null) ? labelNode[uiLang] : labelNode.en;
        if (localizedText == null) return labelPath;
        return String(localizedText).replace(/\{slash\}/g, "/");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * プレビューを作り直す（前回分を undo してから processFn を実行）
     * @param {{isUndo: boolean}} previewState - プレビューの状態（undo すべき変更があるか）
     * @param {function} processFn - プレビューとして実行する処理
     * @param {boolean} isEnabled - プレビューを表示するか（false なら前回分を戻すだけ）
     * @returns {void}
     */
    function runPreview(previewState, processFn, isEnabled) {
        /* app.undo() と DOM 操作は状況によって失敗する / undo and DOM edits may fail */
        try {
            if (isEnabled) {
                if (previewState.isUndo) app.undo();
                else previewState.isUndo = true;
                processFn();
                app.redraw();
            } else if (previewState.isUndo) {
                app.undo();
                app.redraw();
                previewState.isUndo = false;
            }
        } catch (err) { }
    }

    /**
     * プレビュー分を巻き戻す（確定処理の直前・キャンセル時）
     * @param {{isUndo: boolean}} previewState - プレビューの状態
     * @returns {void}
     */
    function undoPreview(previewState) {
        /* 取り消す履歴が無いと app.undo() が失敗することがある / undo may fail with no history */
        try {
            if (previewState.isUndo) app.undo();
        } catch (err) { }
        previewState.isUndo = false;
    }

    /**
     * セッション内に控えたダイアログ位置を返す
     * @returns {number[]|null} [x, y]（控えが無ければ null）
     */
    function loadDialogPosition() {
        var savedPosition = $.global[DIALOG_POSITION_KEY];
        return (savedPosition && savedPosition.length === 2) ? [savedPosition[0], savedPosition[1]] : null;
    }

    /**
     * ダイアログ位置をセッション内に控える
     * @param {number[]} dialogLocation - ダイアログの位置 [x, y]
     * @returns {void}
     */
    function saveDialogPosition(dialogLocation) {
        if (!dialogLocation || dialogLocation.length !== 2) return;
        $.global[DIALOG_POSITION_KEY] = [Math.round(dialogLocation[0]), Math.round(dialogLocation[1])];
    }

    // =========================================
    // 境界と並び順 / Bounds and ordering
    // =========================================

    /**
     * クリップグループのクリッピングパスを返す
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {PageItem|null} クリッピングパス（クリップグループでなければ null）
     */
    function getClippingPath(targetItem) {
        if (!targetItem || targetItem.typename !== "GroupItem" || !targetItem.clipped) return null;

        var childItems = targetItem.pageItems;
        for (var i = 0; i < childItems.length; i++) {
            var childItem = childItems[i];
            /* clipping を持たない種類のアイテムが混ざると例外になる / some item kinds do not expose clipping */
            try {
                if (childItem.clipping === true) return childItem;
            } catch (e) { }
            /* 複合パスはマスク本体に clipping が無く、内部パスに付く / Compound path: clipping flag sits on inner path, not the wrapper */
            if (childItem.typename === "CompoundPathItem" && childItem.pathItems && childItem.pathItems.length) {
                try {
                    if (childItem.pathItems[0].clipping === true) return childItem;
                } catch (e) { }
            }
        }
        return null;
    }

    /**
     * オブジェクトの境界を返す（クリップグループはクリッピングパスを測る）
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界（visibleBounds）を使うか
     * @returns {number[]} [左, 上, 右, 下]
     */
    function getItemBounds(targetItem, usePreviewBounds) {
        var clipPath = getClippingPath(targetItem);
        var measuredItem = clipPath ? clipPath : targetItem;
        return usePreviewBounds ? measuredItem.visibleBounds : measuredItem.geometricBounds;
    }

    /**
     * 選択全体の幅・高さを返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {{spanX: number, spanY: number}} 幅と高さ
     */
    function getSelectionSpan(targetItems, usePreviewBounds) {
        var minLeft = null, maxRight = null, maxTop = null, minBottom = null;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (!targetItem) continue;
            var bounds = getItemBounds(targetItem, usePreviewBounds);
            if (minLeft === null || bounds[0] < minLeft) minLeft = bounds[0];
            if (maxRight === null || bounds[2] > maxRight) maxRight = bounds[2];
            if (maxTop === null || bounds[1] > maxTop) maxTop = bounds[1];
            if (minBottom === null || bounds[3] < minBottom) minBottom = bounds[3];
        }
        if (minLeft === null) return { spanX: 0, spanY: 0 };
        return { spanX: (maxRight - minLeft), spanY: (maxTop - minBottom) };
    }

    /**
     * 選択の縦横比から並べる方向を判定する（geometricBounds で安定させる）
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {string} "horizontal" または "vertical"
     */
    function detectDirection(targetItems) {
        var selectionSpan = getSelectionSpan(targetItems, false);
        return (selectionSpan.spanX >= selectionSpan.spanY) ? "horizontal" : "vertical";
    }

    /**
     * 上→下（同じ高さなら左→右）に並べ替えた新しい配列を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortTopToBottom(targetItems) {
        var sortedItems = targetItems.slice();
        sortedItems.sort(function (itemA, itemB) {
            if (itemA.top !== itemB.top) return itemB.top - itemA.top;
            return itemA.left - itemB.left;
        });
        return sortedItems;
    }

    /**
     * 左→右（同じ位置なら上→下）に並べ替えた新しい配列を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortLeftToRight(targetItems) {
        var sortedItems = targetItems.slice();
        sortedItems.sort(function (itemA, itemB) {
            if (itemA.left !== itemB.left) return itemA.left - itemB.left;
            return itemB.top - itemA.top;
        });
        return sortedItems;
    }

    /**
     * 0〜count-1 のインデックスを Fisher-Yates でシャッフルした配列を返す
     * @param {number} count - 要素数
     * @returns {number[]} シャッフルしたインデックス
     */
    function makeShuffledIndices(count) {
        var indices = [];
        for (var i = 0; i < count; i++) indices.push(i);
        for (var j = indices.length - 1; j > 0; j--) {
            var k = Math.floor(Math.random() * (j + 1));
            var swappedIndex = indices[j]; indices[j] = indices[k]; indices[k] = swappedIndex;
        }
        return indices;
    }

    /**
     * オブジェクト（または子孫）がテキストを含むかを返す
     * @param {PageItem} targetItem - 対象のオブジェクト
     * @returns {boolean} テキストを含めば true
     */
    function containerHasText(targetItem) {
        if (!targetItem) return false;
        /* textFrames / pageItems を持たない種類がある / not every item kind has textFrames or pageItems */
        try {
            if (targetItem.typename === "TextFrame") return true;
            if (targetItem.textFrames && targetItem.textFrames.length > 0) return true;
            if (targetItem.pageItems && targetItem.pageItems.length) {
                for (var i = 0; i < targetItem.pageItems.length; i++) {
                    if (containerHasText(targetItem.pageItems[i])) return true;
                }
            }
        } catch (e) { }
        return false;
    }

    /**
     * いずれかのオブジェクトがテキストを含むかを返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @returns {boolean} テキストを含むものがあれば true
     */
    function anyItemHasText(targetItems) {
        for (var i = 0; i < targetItems.length; i++) {
            if (containerHasText(targetItems[i])) return true;
        }
        return false;
    }

    // =========================================
    // キー操作 / Keyboard
    // =========================================

    /**
     * 入力欄で ↑↓ キーによる値の増減を有効にする（Shift で 10 刻みにスナップ）
     * @param {EditText} editText - 対象の入力欄
     * @param {boolean} allowNegative - マイナス値を許すか
     * @param {function} [onUpdate] - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative, onUpdate) {
        editText.addEventListener("keydown", function (event) {
            if (editText.text.length === 0) return;
            var currentValue = Number(editText.text);
            if (isNaN(currentValue)) return;

            var keyboardState = ScriptUI.environment.keyboardState;
            if (event.keyName == "Up" || event.keyName == "Down") {
                var isUp = event.keyName == "Up";
                var delta = 1;
                if (keyboardState.shiftKey) {
                    /* 10の倍数にスナップ / Snap to multiples of 10 */
                    currentValue = Math.floor(currentValue / 10) * 10;
                    delta = 10;
                }
                currentValue += isUp ? delta : -delta;
                if (!allowNegative && currentValue < 0) currentValue = 0;

                event.preventDefault();
                editText.text = currentValue;
                if (typeof onUpdate === "function") onUpdate();
            }
        });
    }

    /* 揃えのショートカットキーと、対応するラジオのキー / Shortcut keys mapped to a radio in each row */
    var HORIZONTAL_ALIGN_RADIO_BY_KEY = { "N": "none", "L": "left", "C": "center", "R": "right" };
    var VERTICAL_ALIGN_RADIO_BY_KEY = { "N": "none", "T": "top", "M": "middle", "B": "bottom" };

    /**
     * N / L / C / R / T / M / B キーで揃えのラジオを切り替える
     * @param {object} keyTarget - キーを受けるダイアログまたはコントロール
     * @param {{horizontal: Object, vertical: Object, getDirection: function}} alignRadios - 揃えのラジオ一式と方向の取得関数
     * @param {function} [onUpdate] - 切り替えたあとに呼ぶ処理
     * @returns {void}
     */
    function addAlignKeyHandler(keyTarget, alignRadios, onUpdate) {
        keyTarget.addEventListener("keydown", function (event) {
            /* 横並びのときは上下の揃え、縦並びのときは左右の揃えを操作する
               A horizontal layout adjusts the vertical align, and vice versa */
            var isHorizontalLayout = (alignRadios.getDirection() === "horizontal");
            var radioByKey = isHorizontalLayout ? VERTICAL_ALIGN_RADIO_BY_KEY : HORIZONTAL_ALIGN_RADIO_BY_KEY;
            var radioSet = isHorizontalLayout ? alignRadios.vertical : alignRadios.horizontal;

            var radioKey = radioByKey[event.keyName];
            if (!radioKey) return;

            for (var key in radioSet) {
                radioSet[key].value = (key === radioKey);
            }
            event.preventDefault();
            if (onUpdate) onUpdate();
        });
    }

    // =========================================
    // アウトライン計測 / Outline measurement
    // =========================================

    /**
     * コンテナ内のすべてのテキストをアウトライン化する
     * @param {PageItem} containerItem - テキスト、またはテキストを含むグループ
     * @returns {void}
     */
    function outlineAllTextInContainer(containerItem) {
        if (!containerItem) return;
        /* createOutline() は空のテキストなどで失敗する / createOutline() can fail, e.g. on empty text */
        try {
            if (containerItem.typename === "TextFrame") {
                containerItem.createOutline();
                return;
            }
            if (containerItem.textFrames && containerItem.textFrames.length) {
                /* createOutline で要素が消えるため後ろから走査 / iterate backwards: createOutline removes the frame */
                for (var i = containerItem.textFrames.length - 1; i >= 0; i--) {
                    try { containerItem.textFrames[i].createOutline(); } catch (e1) { }
                }
            }
        } catch (e) { }
    }

    /**
     * 複製をアウトライン化して境界を測る（boundsCache に控えて再利用）
     * @param {PageItem} originalItem - 測るオブジェクト
     * @param {Array} boundsCache - [{item: PageItem, bounds: number[]}] の控え
     * @returns {number[]|null} [左, 上, 右, 下]（テキストを含まない・測れないときは null）
     */
    function measureOutlineBoundsOnce(originalItem, boundsCache) {
        if (!originalItem) return null;

        for (var i = 0; i < boundsCache.length; i++) {
            if (boundsCache[i].item === originalItem) return boundsCache[i].bounds;
        }
        if (!containerHasText(originalItem)) return null;

        var doc = app.activeDocument;
        var measureGroup = null;
        try {
            var measureLayer = null;
            try { measureLayer = originalItem.layer; } catch (eLayer) { }
            if (!measureLayer) measureLayer = doc.activeLayer;
            measureGroup = measureLayer.groupItems.add();
            measureGroup.name = TEMP_MEASURE_GROUP_NAME;

            var duplicatedItem = null;
            try {
                duplicatedItem = originalItem.duplicate(measureGroup, ElementPlacement.PLACEATEND);
            } catch (eDup) {
                duplicatedItem = originalItem.duplicate();
                duplicatedItem.move(measureGroup, ElementPlacement.PLACEATEND);
            }

            outlineAllTextInContainer(duplicatedItem);

            var measuredBounds = null;
            try { measuredBounds = measureGroup.visibleBounds; }
            catch (eVisible) { try { measuredBounds = measureGroup.geometricBounds; } catch (eGeometric) { } }
            if (!measuredBounds) return null;

            boundsCache.push({ item: originalItem, bounds: [measuredBounds[0], measuredBounds[1], measuredBounds[2], measuredBounds[3]] });
            return [measuredBounds[0], measuredBounds[1], measuredBounds[2], measuredBounds[3]];
        } catch (e) {
            return null;
        } finally {
            try { if (measureGroup) measureGroup.remove(); } catch (eRemove) { }
        }
    }

    // =========================================
    // レイアウト適用 / Layout application
    // =========================================

    /**
     * 並べる順（並べ替え、またはシャッフル）と、ランダム時の基準位置を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {string} direction - "horizontal" / "vertical"
     * @param {boolean} useRandom - ランダムに並べるか
     * @param {number[]|null} previousShuffleOrder - 前回のシャッフル順（件数が同じなら使い回す）
     * @returns {{items: PageItem[], baseLeft: number, baseTop: number, shuffleOrder: number[]}} 並べる順と基準位置
     */
    function buildArrangementOrder(targetItems, direction, useRandom, previousShuffleOrder) {
        var arrangement = { items: [], baseLeft: null, baseTop: null, shuffleOrder: previousShuffleOrder };

        if (useRandom) {
            if (!arrangement.shuffleOrder || arrangement.shuffleOrder.length !== targetItems.length) {
                arrangement.shuffleOrder = makeShuffledIndices(targetItems.length);
            }
            for (var i = 0; i < arrangement.shuffleOrder.length; i++) {
                arrangement.items.push(targetItems[arrangement.shuffleOrder[i]]);
            }
            /* ランダム時は元の左上を基準として保持 / Preserve top-left base for random */
            for (var k = 0; k < targetItems.length; k++) {
                var targetItem = targetItems[k];
                if (!targetItem) continue;
                if (arrangement.baseLeft === null || targetItem.left < arrangement.baseLeft) arrangement.baseLeft = targetItem.left;
                if (arrangement.baseTop === null || targetItem.top > arrangement.baseTop) arrangement.baseTop = targetItem.top;
            }
            if (arrangement.baseLeft === null && arrangement.items[0]) arrangement.baseLeft = arrangement.items[0].left;
            if (arrangement.baseTop === null && arrangement.items[0]) arrangement.baseTop = arrangement.items[0].top;
        } else {
            arrangement.shuffleOrder = null;
            arrangement.items = (direction === "horizontal") ? sortLeftToRight(targetItems) : sortTopToBottom(targetItems);
        }
        return arrangement;
    }

    /**
     * オブジェクトの幅または高さの最大値を返す
     * @param {PageItem[]} targetItems - 対象のオブジェクト
     * @param {function} getBounds - 境界 [左, 上, 右, 下] を返す関数
     * @param {boolean} measureWidth - true なら幅、false なら高さ
     * @returns {number} 最大値
     */
    function computeMaxSize(targetItems, getBounds, measureWidth) {
        var maxSize = 0;
        for (var i = 0; i < targetItems.length; i++) {
            var bounds = getBounds(targetItems[i]);
            var itemSize = measureWidth ? (bounds[2] - bounds[0]) : (bounds[1] - bounds[3]);
            if (itemSize > maxSize) maxSize = itemSize;
        }
        return maxSize;
    }

    /**
     * 左→右に横並びで配置し、上下揃えを適用する
     * @param {PageItem[]} targetItems - 並べる順のオブジェクト
     * @param {number} startX - 先頭の左端
     * @param {number} startY - 揃えの基準にする上端
     * @param {number} referenceHeight - 揃えの基準にする高さ
     * @param {number} spacingPt - 間隔（pt）
     * @param {string} vAlignMode - "none" / "top" / "middle" / "bottom"
     * @param {function} getBounds - 境界を返す関数
     * @returns {void}
     */
    function applyHorizontalLayout(targetItems, startX, startY, referenceHeight, spacingPt, vAlignMode, getBounds) {
        var currentX = startX;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (!targetItem) continue;
            var bounds = getBounds(targetItem);
            var itemWidth = bounds[2] - bounds[0];

            /* Xは左基準で配置 / Place by left edge on X */
            targetItem.left = targetItem.left + (currentX - bounds[0]);

            /* 上下揃え / Vertical alignment */
            if (vAlignMode !== "none") {
                var cellTop = startY;
                var cellBottom = startY - referenceHeight;
                var dy = 0;
                if (vAlignMode === "middle") dy = ((cellTop + cellBottom) / 2) - ((bounds[1] + bounds[3]) / 2);
                else if (vAlignMode === "bottom") dy = cellBottom - bounds[3];
                else dy = cellTop - bounds[1]; /* top */
                targetItem.top = targetItem.top + dy;
            }

            if (i < targetItems.length - 1) {
                currentX += itemWidth + spacingPt;
            }
        }
    }

    /**
     * 上→下に縦並びで配置し、左右揃えを適用する
     * @param {PageItem[]} targetItems - 並べる順のオブジェクト
     * @param {number} startX - 揃えの基準にする左端
     * @param {number} startY - 先頭の上端
     * @param {number} referenceWidth - 揃えの基準にする幅
     * @param {number} spacingPt - 間隔（pt）
     * @param {string} hAlignMode - "none" / "left" / "center" / "right"
     * @param {function} getBounds - 境界を返す関数
     * @returns {void}
     */
    function applyVerticalLayout(targetItems, startX, startY, referenceWidth, spacingPt, hAlignMode, getBounds) {
        var currentY = startY;
        for (var i = 0; i < targetItems.length; i++) {
            var targetItem = targetItems[i];
            if (!targetItem) continue;
            var bounds = getBounds(targetItem);
            var itemHeight = bounds[1] - bounds[3];

            /* 左右揃え / Horizontal alignment */
            if (hAlignMode !== "none") {
                var cellLeft = startX;
                var cellRight = cellLeft + referenceWidth;
                var dx = 0;
                if (hAlignMode === "center") dx = ((cellLeft + cellRight) / 2) - ((bounds[0] + bounds[2]) / 2);
                else if (hAlignMode === "right") dx = cellRight - bounds[2];
                else dx = cellLeft - bounds[0]; /* left */
                targetItem.left = targetItem.left + dx;
            }

            /* YはTop揃えで積む / Stack by top on Y */
            targetItem.top = targetItem.top + (currentY - bounds[1]);

            if (i < targetItems.length - 1) {
                currentY -= itemHeight + spacingPt;
            }
        }
    }

    /**
     * ランダム時に、先頭のオブジェクトが元の左上に来るよう全体をずらす
     * @param {PageItem[]} targetItems - 並べたオブジェクト
     * @param {number} baseLeft - 元の左端
     * @param {number} baseTop - 元の上端
     * @returns {void}
     */
    function applyRandomBaseOffset(targetItems, baseLeft, baseTop) {
        if (!targetItems.length) return;
        var dx = baseLeft - targetItems[0].left;
        var dy = baseTop - targetItems[0].top;
        for (var i = 0; i < targetItems.length; i++) {
            if (!targetItems[i]) continue;
            targetItems[i].left += dx;
            targetItems[i].top += dy;
        }
    }

    /**
     * 並べる順に従ってオブジェクトを配置する
     * @param {{items: PageItem[], baseLeft: number, baseTop: number}} arrangement - buildArrangementOrder() の結果
     * @param {{direction: string, spacingPt: number, useRandom: boolean, hAlignMode: string, vAlignMode: string}} layoutSettings - 並べ方
     * @param {function} getBounds - 境界 [左, 上, 右, 下] を返す関数
     * @returns {void}
     */
    function placeArrangement(arrangement, layoutSettings, getBounds) {
        var arrangedItems = arrangement.items;
        if (!arrangedItems.length) return;

        var startBounds = getBounds(arrangedItems[0]);
        var startX = startBounds[0];
        var startY = startBounds[1];

        if (layoutSettings.direction === "horizontal") {
            var referenceHeight = computeMaxSize(arrangedItems, getBounds, false);
            applyHorizontalLayout(arrangedItems, startX, startY, referenceHeight, layoutSettings.spacingPt, layoutSettings.vAlignMode, getBounds);
        } else {
            var referenceWidth = computeMaxSize(arrangedItems, getBounds, true);
            applyVerticalLayout(arrangedItems, startX, startY, referenceWidth, layoutSettings.spacingPt, layoutSettings.hAlignMode, getBounds);
        }

        if (layoutSettings.useRandom) {
            applyRandomBaseOffset(arrangedItems, arrangement.baseLeft, arrangement.baseTop);
        }
    }

    // =========================================
    // ダイアログ部品 / Dialog parts
    // =========================================

    /**
     * ［方向］パネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{auto: RadioButton, vertical: RadioButton, horizontal: RadioButton}} 方向のラジオ
     */
    function addDirectionPanel(parentDialog) {
        var directionPanel = parentDialog.add("panel", undefined, getLabel("panel.direction"));
        directionPanel.orientation = "row";
        directionPanel.alignChildren = ["center", "center"];
        directionPanel.margins = PANEL_MARGINS;
        var directionRadios = {
            auto: addTooltipRadio(directionPanel, getLabel("radio.directionAuto"), getLabel("tooltip.directionAuto")),
            vertical: addTooltipRadio(directionPanel, getLabel("radio.directionVertical"), getLabel("tooltip.directionVertical")),
            horizontal: addTooltipRadio(directionPanel, getLabel("radio.directionHorizontal"), getLabel("tooltip.directionHorizontal"))
        };
        directionRadios.vertical.value = true; /* デフォルトを「縦」に / Default to Vertical */
        return directionRadios;
    }

    /**
     * ツールチップ付きのラジオボタンを追加する
     * @param {object} parentGroup - 追加先のパネルまたはグループ
     * @param {string} radioText - ラジオの表示名
     * @param {string} tooltipText - ツールチップ
     * @returns {RadioButton} 追加したラジオ
     */
    function addTooltipRadio(parentGroup, radioText, tooltipText) {
        var radioButton = parentGroup.add("radiobutton", undefined, radioText);
        radioButton.helpTip = tooltipText;
        return radioButton;
    }

    /**
     * ［間隔］パネルを追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {EditText} 間隔の入力欄
     */
    function addSpacingPanel(parentDialog) {
        var spacingPanel = parentDialog.add("panel", undefined, getLabel("panel.spacing"));
        spacingPanel.orientation = "column";
        spacingPanel.alignChildren = ["center", "center"];
        spacingPanel.margins = PANEL_MARGINS;
        var spacingRowGroup = spacingPanel.add("group");
        spacingRowGroup.orientation = "row";
        spacingRowGroup.alignChildren = ["left", "center"];
        var spacingInput = spacingRowGroup.add("edittext", undefined, "0");
        spacingInput.characters = SPACING_FIELD_CHARS;
        spacingInput.helpTip = getLabel("tooltip.spacing");
        spacingRowGroup.add("statictext", undefined, getUnitInfo().label);
        return spacingInput;
    }

    /**
     * 揃えのラジオを1行ぶん追加する
     * @param {Panel} alignPanel - 追加先のパネル
     * @param {Array} radioDefs - [ラジオのキー, LABELS のパス] の配列
     * @param {string} defaultKey - 最初に選ぶラジオのキー
     * @param {string} tooltipText - 行のラジオすべてに付けるツールチップ
     * @returns {{rowGroup: Group, radios: Object}} 行のグループと、キー → ラジオ
     */
    function addAlignRadioRow(alignPanel, radioDefs, defaultKey, tooltipText) {
        var rowGroup = alignPanel.add("group");
        rowGroup.orientation = "row";
        rowGroup.alignChildren = ["left", "center"];
        var alignRadios = {};
        for (var i = 0; i < radioDefs.length; i++) {
            alignRadios[radioDefs[i][0]] = rowGroup.add("radiobutton", undefined, getLabel(radioDefs[i][1]));
        }
        alignRadios[defaultKey].value = true;
        for (var radioKey in alignRadios) {
            alignRadios[radioKey].helpTip = tooltipText;
        }
        return { rowGroup: rowGroup, radios: alignRadios };
    }

    /**
     * ［揃え］パネル（左右の揃えと上下の揃えの2行）を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{panel: Panel, horizontalRow: Group, verticalRow: Group, horizontal: Object, vertical: Object}} パネルと各行のラジオ
     */
    function addAlignPanel(parentDialog) {
        var alignPanel = parentDialog.add("panel", undefined, "");
        alignPanel.orientation = "column";
        alignPanel.alignChildren = ["left", "center"];
        alignPanel.margins = PANEL_MARGINS;

        /* 既定は「中央」/ Default to Center (Middle) */
        var horizontalRow = addAlignRadioRow(alignPanel, [
            ["none", "radio.alignNone"], ["left", "radio.alignLeft"], ["center", "radio.alignCenter"], ["right", "radio.alignRight"]
        ], "center", getLabel("tooltip.alignHorizontal"));
        var verticalRow = addAlignRadioRow(alignPanel, [
            ["none", "radio.alignNone"], ["top", "radio.alignTop"], ["middle", "radio.alignMiddle"], ["bottom", "radio.alignBottom"]
        ], "middle", getLabel("tooltip.alignVertical"));

        return {
            panel: alignPanel,
            horizontalRow: horizontalRow.rowGroup,
            verticalRow: verticalRow.rowGroup,
            horizontal: horizontalRow.radios,
            vertical: verticalRow.radios
        };
    }

    /**
     * オプションのチェックボックス（プレビュー境界・テキスト計測・ランダム）を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {{usePreviewBounds: Checkbox, measureText: Checkbox, random: Checkbox}} チェックボックス
     */
    function addOptionsGroup(parentDialog) {
        var optionsGroup = parentDialog.add("group");
        optionsGroup.orientation = "column";
        optionsGroup.alignChildren = ["left", "center"];
        optionsGroup.alignment = ["fill", "top"];
        optionsGroup.margins = OPTIONS_MARGINS;

        return {
            usePreviewBounds: addOptionCheckbox(optionsGroup, getLabel("checkbox.usePreviewBounds"), getLabel("tooltip.usePreviewBounds"), true),
            measureText: addOptionCheckbox(optionsGroup, getLabel("checkbox.measureText"), getLabel("tooltip.measureText"), false),
            random: addOptionCheckbox(optionsGroup, getLabel("checkbox.random"), getLabel("tooltip.random"), false)
        };
    }

    /**
     * ツールチップ付きのチェックボックスを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} checkboxText - 表示名
     * @param {string} tooltipText - ツールチップ
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentGroup, checkboxText, tooltipText, initialValue) {
        var optionCheckbox = parentGroup.add("checkbox", undefined, checkboxText);
        optionCheckbox.value = initialValue;
        optionCheckbox.helpTip = tooltipText;
        return optionCheckbox;
    }

    /**
     * ［キャンセル］［OK］のボタン行を追加する
     * @param {Window} parentDialog - 追加先のダイアログ
     * @returns {void}
     */
    function addButtonRow(parentDialog) {
        var btnRowGroup = parentDialog.add("group");
        btnRowGroup.alignment = "center";
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    }

    /**
     * ラジオの組から選ばれているもののキーを返す
     * @param {object} radioSet - キー → ラジオ
     * @param {string} fallbackKey - どれも選ばれていないときのキー
     * @returns {string} 選ばれているラジオのキー
     */
    function getCheckedKey(radioSet, fallbackKey) {
        for (var radioKey in radioSet) {
            if (radioSet[radioKey].value) return radioKey;
        }
        return fallbackKey;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを表示し、プレビューしながら整列・分布を実行する
     * @returns {boolean} OK で確定したら true、キャンセルなら false
     */
    function showArrangeDialog() {
        var alignDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        alignDialog.orientation = "column";
        alignDialog.alignChildren = "fill";
        alignDialog.opacity = DIALOG_OPACITY;

        var lastPosition = loadDialogPosition();
        if (lastPosition) alignDialog.location = lastPosition;

        /* キャンセル時に復元するため現在のプリファレンスを保存 / Preserve preference for cancel */
        var originalIncludeStrokeInBounds = app.preferences.getBooleanPreference("includeStrokeInBounds");

        /* 選択スナップショット / Snapshot selection */
        var originalSelection = app.activeDocument.selection.slice();
        var previewState = { isUndo: false };
        var detectedDirection = detectDirection(originalSelection);
        var selectionHasText = anyItemHasText(originalSelection);

        var directionRadios = addDirectionPanel(alignDialog);
        var spacingInput = addSpacingPanel(alignDialog);
        var alignControls = addAlignPanel(alignDialog);
        var optionCheckboxes = addOptionsGroup(alignDialog);
        addButtonRow(alignDialog);

        /* ---- 内部状態 / Internal state ---- */
        var randomOrderCache = null;
        var outlineMeasureCache = []; /* [{item:PageItem, bounds:[l,t,r,b]}] */
        var isPreviewUpdating = false;
        var lastPreviewTime = 0;
        var previewTaskId = 0;

        /**
         * 選択中のラジオから実際に並べる方向を返す（自動は判定結果）
         * @returns {string} "horizontal" または "vertical"
         */
        function getEffectiveDirection() {
            if (directionRadios.vertical.value) return "vertical";
            if (directionRadios.horizontal.value) return "horizontal";
            return detectedDirection;
        }

        /**
         * レイアウト用の境界を返す（縦並びでテキスト計測が ON ならアウトラインの実測値）
         * @param {PageItem} targetItem - 対象のオブジェクト
         * @returns {number[]} [左, 上, 右, 下]
         */
        function getBoundsForLayout(targetItem) {
            if (optionCheckboxes.measureText.value && getEffectiveDirection() !== "horizontal") {
                var outlinedBounds = measureOutlineBoundsOnce(targetItem, outlineMeasureCache);
                if (outlinedBounds) return outlinedBounds;
            }
            return getItemBounds(targetItem, optionCheckboxes.usePreviewBounds.value);
        }

        /**
         * 間隔の入力値を現在の単位から pt に換算する
         * @returns {number} 間隔（pt）
         */
        function getSpacingPt() {
            var spacingValue = parseFloat(spacingInput.text);
            if (isNaN(spacingValue)) spacingValue = 0;
            return spacingValue * getUnitInfo().pointsPerUnit;
        }

        /**
         * 現在の設定で選択を並べる
         * @returns {void}
         */
        function applyLayoutToSelection() {
            if (!originalSelection || originalSelection.length === 0) return;
            var direction = getEffectiveDirection();
            var spacingPt = getSpacingPt();
            var useRandom = optionCheckboxes.random.value;
            var arrangement = buildArrangementOrder(originalSelection, direction, useRandom, randomOrderCache);
            randomOrderCache = arrangement.shuffleOrder;
            placeArrangement(arrangement, {
                direction: direction,
                spacingPt: spacingPt,
                useRandom: useRandom,
                hAlignMode: getCheckedKey(alignControls.horizontal, "left"),
                vAlignMode: getCheckedKey(alignControls.vertical, "top")
            }, getBoundsForLayout);
        }

        /**
         * 方向に合わせて揃えパネルの見出しと有効状態、テキスト計測の有効状態を切り替える
         * @returns {void}
         */
        function syncAlignUI() {
            var isHorizontal = (getEffectiveDirection() === "horizontal");
            alignControls.panel.text = getLabel(isHorizontal ? "panel.alignVertical" : "panel.alignHorizontal");
            alignControls.horizontalRow.enabled = !isHorizontal;
            alignControls.verticalRow.enabled = isHorizontal;

            /* テキスト計測は縦並び時かつ選択にテキストがある場合のみ有効 / Enable text-measure only when vertical & selection has text */
            optionCheckboxes.measureText.enabled = !isHorizontal && selectionHasText;

            /* 表示前のレイアウトは失敗することがある / layout before show may fail */
            try { alignDialog.layout.layout(true); } catch (e) { }
        }

        /**
         * 連続呼び出しを間引きながらプレビューを作り直す
         * @returns {void}
         */
        function updatePreview() {
            if (isPreviewUpdating) return;
            var now = new Date().getTime();
            if (now - lastPreviewTime < PREVIEW_MIN_INTERVAL_MS) return;
            lastPreviewTime = now;

            isPreviewUpdating = true;
            try {
                try {
                    app.preferences.setBooleanPreference("includeStrokeInBounds", optionCheckboxes.usePreviewBounds.value);
                } catch (e) { }
                runPreview(previewState, applyLayoutToSelection, true);
            } finally {
                isPreviewUpdating = false;
            }
        }

        /**
         * プレビューを予約する（ScriptUI のイベント中に DOM を触ると落ちるため scheduleTask で遅らせる）
         * @returns {void}
         */
        function requestPreviewUpdate() {
            try { if (previewTaskId) app.cancelTask(previewTaskId); } catch (e) { }
            $.global.__SAT_updatePreview = updatePreview;
            try {
                previewTaskId = app.scheduleTask('$.global.__SAT_updatePreview && $.global.__SAT_updatePreview();', PREVIEW_SCHEDULE_MS, false);
            } catch (e) {
                updatePreview();
            }
        }

        /**
         * 方向を変えたら揃えパネルを同期してプレビューを更新する
         * @returns {void}
         */
        function onDirectionChanged() {
            syncAlignUI();
            requestPreviewUpdate();
        }

        /* ---- イベントバインド / Event bindings ---- */
        changeValueByArrowKey(spacingInput, true, requestPreviewUpdate);

        var alignRadios = {
            horizontal: alignControls.horizontal,
            vertical: alignControls.vertical,
            getDirection: getEffectiveDirection
        };
        addAlignKeyHandler(alignDialog, alignRadios, requestPreviewUpdate);
        addAlignKeyHandler(spacingInput, alignRadios, requestPreviewUpdate);

        var radioKey;
        for (radioKey in alignControls.horizontal) alignControls.horizontal[radioKey].onClick = requestPreviewUpdate;
        for (radioKey in alignControls.vertical) alignControls.vertical[radioKey].onClick = requestPreviewUpdate;

        directionRadios.auto.onClick = onDirectionChanged;
        directionRadios.vertical.onClick = onDirectionChanged;
        directionRadios.horizontal.onClick = onDirectionChanged;

        optionCheckboxes.usePreviewBounds.onClick = function () {
            /* 境界モードが変わるとアウトライン計測結果も再計算 / Outline cache invalid when bounds mode changes */
            outlineMeasureCache = [];
            requestPreviewUpdate();
        };
        optionCheckboxes.random.onClick = function () {
            randomOrderCache = null;
            requestPreviewUpdate();
        };
        optionCheckboxes.measureText.onClick = function () {
            outlineMeasureCache = [];
            requestPreviewUpdate();
        };

        /* 初期化 / Init */
        syncAlignUI();
        requestPreviewUpdate();
        spacingInput.active = true;

        var dialogResult = alignDialog.show();
        try { if (previewTaskId) app.cancelTask(previewTaskId); } catch (e) { }
        saveDialogPosition(alignDialog.location);

        if (dialogResult !== 1) {
            /* キャンセル: プレビューを undo してプリファレンスも戻す / Cancel: undo preview & restore preference */
            undoPreview(previewState);
            app.preferences.setBooleanPreference("includeStrokeInBounds", originalIncludeStrokeInBounds);
            app.redraw();
            return false;
        }

        /* 確定: プレビュー分を巻き戻して本実行（undo 履歴を 1 件にまとめる） / OK: undo preview, then commit as single history entry */
        undoPreview(previewState);
        try { app.preferences.setBooleanPreference("includeStrokeInBounds", optionCheckboxes.usePreviewBounds.value); } catch (e) { }
        applyLayoutToSelection();
        app.redraw();
        return true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択を確かめてダイアログを開く
     * @returns {void}
     */
    function main() {
        /* ドキュメントが無いと app.activeDocument が例外になる / activeDocument throws with no document */
        try {
            var selectedItems = app.activeDocument.selection;
            if (!selectedItems || selectedItems.length === 0) {
                alert(getLabel("alert.noSelection"));
                return;
            }
            showArrangeDialog();
        } catch (e) {
            alert(getLabel("alert.errorPrefix") + e.message);
        }
    }

    main();

})();
