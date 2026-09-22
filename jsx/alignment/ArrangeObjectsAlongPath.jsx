#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

複数のオブジェクトを、選択範囲内の1本のパスに沿って等間隔に自動配置します。
基準にするパスは「自動（面積最大）／最前面／最背面」から選べます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeObjectsAlongPath.md

### Overview

Distributes several objects evenly along a single path taken from the selection.
The reference path is chosen automatically by largest area, or as the frontmost or backmost object.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeObjectsAlongPath.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ArrangeObjectsAlongPath";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.5.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ArrangeObjectsAlongPath.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ArrangeObjectsAlongPath.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 配置後、さらに接線方向へ回転する（簡易）/ Also rotate along the tangent after placing (simple) */
    var ROTATE_ALONG_TANGENT = false;
    /* true: 始点〜終点を含めて等分 / false: 端を避ける（開いたパス）/ true: include both ends, false: keep off the ends (open paths) */
    var USE_ENDPOINTS = true;
    /* 曲線1区間あたりのサンプル数（増やすほど精度が上がり重くなる）/ Samples per Bezier segment (more is precise but slower) */
    var SAMPLES_PER_SEGMENT = 30;
    /* ランダム間隔の強さの初期値（ステップに対する比率、0.1〜1.0）/ Initial random spacing strength (ratio of the step, 0.1-1.0) */
    var DEFAULT_SPACING_JITTER_RATIO = 0.4;
    /* 複製数の範囲 / Duplicate count range */
    var DUPLICATE_COUNT_MIN = 2;
    var DUPLICATE_COUNT_MAX = 20;

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* プレビュー用の一時レイヤー名 / Name of the temporary preview layer */
    var PREVIEW_LAYER_NAME = "__PREVIEW_ArrangeAlongPath";

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ダイアログの位置と不透明度 / Dialog position and opacity */
    var DIALOG_OFFSET_X = 300;
    var DIALOG_OFFSET_Y = 0;
    var DIALOG_OPACITY = 0.98;

    var COLUMN_SPACING = 15;                        /* 左右カラムの間隔 / Gap between the columns */
    var OUTER_PANEL_MARGINS = [15, 20, 15, 15];     /* 「対象パス」パネルの余白 / Margins of the Target Path panel */
    var PANEL_MARGINS = [15, 20, 15, 10];           /* その他のパネルの余白 / Margins of the other panels */
    var SLIDER_WIDTH = 180;                         /* スライダーの幅 / Slider width */

    /**
     * 表示時にダイアログを指定量ずらす
     * @param {Window} targetDialog - 対象のダイアログ
     * @param {number} offsetX - 横方向のずらし量
     * @param {number} offsetY - 縦方向のずらし量
     * @returns {void}
     */
    function shiftDialogPosition(targetDialog, offsetX, offsetY) {
        targetDialog.onShow = function () {
            var currentX = targetDialog.location[0];
            var currentY = targetDialog.location[1];
            targetDialog.location = [currentX + offsetX, currentY + offsetY];
        };
    }

    /**
     * ダイアログの不透明度を設定する
     * @param {Window} targetDialog - 対象のダイアログ
     * @param {number} opacityValue - 不透明度（0〜1）
     * @returns {void}
     */
    function setDialogOpacity(targetDialog, opacityValue) {
        try {
            targetDialog.opacity = opacityValue;
        } catch (e) { /* 環境によっては opacity を持たない / opacity is not supported in some environments */ }
    }

    /**
     * 数値欄に ↑↓ キーでの増減を付ける（shift で 10 刻み、option で 0.1 刻み）
     * @param {EditText} editText - 対象の数値欄
     * @param {boolean} allowNegative - 負の値を許すか
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative) {
        editText.addEventListener("keydown", function (event) {
            if (!event || !event.keyName) return;
            if (event.keyName !== "Up" && event.keyName !== "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboardState = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName === "Up");

            if (keyboardState.shiftKey) {
                /* shift 押下時は 10 の倍数にスナップ / Snap to multiples of 10 with shift */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboardState.altKey) {
                /* option 押下時は 0.1 単位で増減 / Step by 0.1 with option */
                value += isUp ? 0.1 : -0.1;
            } else {
                value += isUp ? 1 : -1;
            }

            if (keyboardState.altKey) {
                /* 小数第1位までに丸め / Round to one decimal place */
                value = Math.round(value * 10) / 10;
            } else {
                /* 整数に丸め / Round to an integer */
                value = Math.round(value);
            }

            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;

            /* 既存の onChange を呼んでプレビューと UI をそろえる / Fire onChange so the preview and UI stay in sync */
            try {
                if (typeof editText.onChange === "function") editText.onChange();
            } catch (e) { /* ハンドラー内の DOM 操作が失敗してもキー操作は止めない / keep key handling alive if the handler's DOM work fails */ }
        });
    }

    /**
     * タイトル付きのパネルを追加する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} titlePath - パネルタイトルの LABELS パス
     * @param {string} orientation - "row" または "column"
     * @param {string|string[]} childAlignment - alignChildren に入れる値
     * @param {number[]} [panelMargins] - 余白（省略時は PANEL_MARGINS）
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentContainer, titlePath, orientation, childAlignment, panelMargins) {
        var optionPanel = parentContainer.add("panel", undefined, getLabel(titlePath));
        optionPanel.orientation = orientation;
        optionPanel.alignChildren = childAlignment;
        optionPanel.margins = panelMargins || PANEL_MARGINS;
        return optionPanel;
    }

    /**
     * 上揃えの縦カラムを追加する
     * @param {Group} parentGroup - 追加先
     * @returns {Group} 追加したカラム
     */
    function addColumn(parentGroup) {
        var columnGroup = parentGroup.add("group");
        columnGroup.orientation = "column";
        columnGroup.alignChildren = "fill";
        columnGroup.alignment = "top";
        return columnGroup;
    }

    /**
     * 余りの幅を吸う伸縮スペーサーを追加する
     * @param {Group} parentGroup - 追加先
     * @returns {void}
     */
    function addStretchSpacer(parentGroup) {
        var stretchSpacer = parentGroup.add("statictext", undefined, "");
        stretchSpacer.alignment = "fill";
        stretchSpacer.minimumSize.width = 10;
        stretchSpacer.maximumSize.width = 10000;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    var LABELS = {
        dialog: {
            title: { ja: "パスに沿って配置", en: "Arrange Objects Along Path" }
        },
        panel: {
            targetPath: { ja: "対象パス", en: "Target Path" },
            basePathRule: { ja: "基準", en: "Base Path" },
            basePathHandling: { ja: "パス処理", en: "Base Path Handling" },
            placeObjects: { ja: "配置するオブジェクト", en: "Objects to Arrange" },
            duplicate: { ja: "複製", en: "Duplicate" },
            rotation: { ja: "回転", en: "Rotation" },
            order: { ja: "順番", en: "Order" },
            spacing: { ja: "間隔", en: "Spacing" }
        },
        fieldLabel: {
            duplicateCount: { ja: "複製数", en: "Count" }
        },
        checkbox: {
            rotationFlip180: { ja: "反転", en: "Flip" },
            groupPlaced: { ja: "グループ化", en: "Group placed objects" },
            allRandom: { ja: "一括ランダム", en: "Random" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        radio: {
            autoLargest: { ja: "自動（面積最大）", en: "Auto (Largest)" },
            frontmost: { ja: "最前面", en: "Frontmost" },
            backmost: { ja: "最背面", en: "Backmost" },
            basePathModeNone: { ja: "何もしない", en: "Do nothing" },
            basePathModeHide: { ja: "「塗り／線」なし", en: "No fill / no stroke" },
            basePathModeDelete: { ja: "削除", en: "Delete" },
            rotationNone: { ja: "正立", en: "Upright" },
            rotationPerpendicular: { ja: "それぞれ垂直", en: "Perpendicular" },
            rotationPathPerpendicular: { ja: "パスに沿う（接線）", en: "Follow Path (Tangent)" },
            rotationRandom: { ja: "ランダム", en: "Random" },
            rotationAngle: { ja: "角度指定", en: "Angle" },
            orderCurrent: { ja: "正順", en: "Current" },
            orderReverse: { ja: "逆順", en: "Reverse" },
            orderRandom: { ja: "ランダム", en: "Random" },
            spacingEven: { ja: "均等（現状）", en: "Even (Current)" },
            spacingRandom: { ja: "ランダム", en: "Random" }
        },
        tooltip: {
            basePathRule: { ja: "選択の中から、どれを基準のパス（B）とみなすかを決めます。", en: "Which of the selected paths is treated as the base path (B)." },
            autoLargest: {
                ja: "面積が最も大きいパスを基準にします。開いたパスは外接矩形の面積で比べます。",
                en: "Uses the path with the largest area. Open paths are compared by their bounding box."
            },
            basePathHandling: { ja: "基準にしたパス（B）を、配置後にどう扱うかを決めます。", en: "What to do with the base path (B) once the objects are placed." },
            duplicateEnabled: {
                ja: "オンにすると、それぞれのオブジェクトを「複製数」の数まで増やしてから配置します。選択が2つのときは最初からオンです。",
                en: "When on, each object is multiplied up to the Count before arranging. Starts on when exactly two objects are selected."
            },
            duplicateCount: {
                ja: "選択したオブジェクトを複製して数を増やしてから配置します。1 なら複製しません。",
                en: "Duplicates the selection to this many copies before arranging. 1 means no duplication."
            },
            rotation: { ja: "配置したオブジェクトの向きの決め方です。", en: "How each placed object is rotated." },
            rotationNone: {
                ja: "オブジェクトの向きを変えません（［反転］がオンなら 180° 回転します）。",
                en: "Keeps each object's current rotation (turns it 180° when Flip is on)."
            },
            rotationPerpendicular: {
                ja: "基準パスの中心から放射状に、外向きに立つよう回転します。",
                en: "Rotates each object to stand outward, radiating from the center of the base path."
            },
            rotationPathPerpendicular: {
                ja: "パスの進む向き（接線）に合わせて回転します。",
                en: "Rotates each object to follow the direction of the path (its tangent)."
            },
            rotationRandom: {
                ja: "オブジェクトごとに −180°〜180° のランダムな角度だけ回転します。",
                en: "Rotates each object by a random angle between -180° and 180°."
            },
            rotationAngle: { ja: "「角度指定」を選んだときに適用する角度です。", en: "The angle applied when Angle is selected." },
            rotationFlip180: { ja: "回転に 180° を加えて、向きを反対にします。", en: "Adds 180° to the rotation to turn each object around." },
            order: { ja: "オブジェクトをパスに沿って並べる順序です。", en: "The order the objects are laid along the path." },
            orderCurrent: {
                ja: "重ね順で前面にあるものから、パスの始点側に並べます。",
                en: "Lays the objects from the start of the path in stacking order, frontmost first."
            },
            orderReverse: {
                ja: "重ね順で背面にあるものから、パスの始点側に並べます。",
                en: "Lays the objects from the start of the path in stacking order, backmost first."
            },
            spacing: { ja: "パス上に配置する間隔の決め方です。", en: "How the objects are spaced along the path." },
            spacingEven: { ja: "パスに沿って等間隔に配置します。", en: "Places the objects at equal intervals along the path." },
            spacingRandom: {
                ja: "等間隔の位置から、下のスライダーの強さでランダムにずらします。",
                en: "Shifts each position randomly away from even spacing, by the strength set with the slider below."
            },
            spacingJitter: {
                ja: "ランダム間隔のばらつきの強さです。右ほど均等な位置から大きくずれます。「ランダム」のときだけ使えます。",
                en: "How far the random spacing strays from even spacing; further right strays more. Available only with Random."
            },
            groupPlaced: { ja: "配置したオブジェクトを1つのグループにまとめます。", en: "Groups the placed objects into a single group." },
            allRandom: { ja: "順番・間隔・回転をまとめてランダムに設定します。", en: "Sets order, spacing and rotation all to random at once." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original state."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントがありません。", en: "No document is open." },
            needSelection: {
                ja: "A（複数オブジェクト）とB（基準パス）を選択してください。基準パス（B）は「自動（面積最大）/ 最前面 / 最背面」で指定できます。",
                en: "Select A (objects) and B (a path). Choose the base path by Auto (largest), Frontmost, or Backmost."
            },
            noBasePath: {
                ja: "選択範囲に基準となるパス（B）が見つかりません。基準パスにしたいパス（PathItem）を含めて選択してください。",
                en: "No base path (B) was found. Include a PathItem to be used as the base path."
            },
            noItems: { ja: "配置対象（A）が見つかりません。", en: "No placeable objects (A) were found." },
            pathTooShort: { ja: "パスが短すぎます。", en: "The path is too short." },
            pathAnalyzeFailed: { ja: "パスの解析に失敗しました。", en: "Failed to analyze the path." },
            pathLengthZero: { ja: "パス長が0です。", en: "The path length is zero." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "panel.order" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // 配列と乱数 / Arrays and random numbers
    // =========================================

    /**
     * 配列の浅いコピーを返す
     * @param {Array} sourceList - 元の配列
     * @returns {Array} コピー
     */
    function copyArray(sourceList) {
        var copiedList = [];
        for (var i = 0; i < sourceList.length; i++) copiedList.push(sourceList[i]);
        return copiedList;
    }

    /**
     * シード付きの乱数関数を作る（xorshift32）
     * @param {number} seed - シード
     * @returns {Function} 0 以上 1 未満の数を返す関数
     */
    function createSeededRandom(seed) {
        var state = (seed | 0);
        if (state === 0) state = 123456789;
        return function () {
            state ^= (state << 13);
            state ^= (state >>> 17);
            state ^= (state << 5);
            return (state >>> 0) / 4294967296;
        };
    }

    /**
     * 現在時刻から並べ替え用のシードを作る
     * @returns {number} シード
     */
    function createShuffleSeed() {
        return (new Date().getTime() & 0x7fffffff) >>> 0;
    }

    /**
     * 配列をその場で並べ替える（Fisher-Yates）
     * @param {Array} targetList - 並べ替える配列
     * @param {Function} [randomFn] - 乱数関数（省略時は Math.random）
     * @returns {void}
     */
    function shuffleInPlace(targetList, randomFn) {
        for (var i = targetList.length - 1; i > 0; i--) {
            var randomValue = randomFn ? randomFn() : Math.random();
            var j = Math.floor(randomValue * (i + 1));
            var swappedItem = targetList[i];
            targetList[i] = targetList[j];
            targetList[j] = swappedItem;
        }
    }

    /**
     * 複製数を範囲内の整数にそろえる
     * @param {number|string} value - 入力値
     * @returns {number} DUPLICATE_COUNT_MIN〜DUPLICATE_COUNT_MAX の整数
     */
    function clampDuplicateCount(value) {
        var count = Math.round(Number(value));
        if (isNaN(count)) count = DUPLICATE_COUNT_MIN;
        if (count < DUPLICATE_COUNT_MIN) count = DUPLICATE_COUNT_MIN;
        if (count > DUPLICATE_COUNT_MAX) count = DUPLICATE_COUNT_MAX;
        return count;
    }

    /**
     * ランダム間隔の強さを 0.1〜1.0 にそろえる
     * @param {number} value - スライダーの値
     * @returns {number} 強さ
     */
    function clampJitterRatio(value) {
        var ratio = Number(value);
        if (isNaN(ratio)) ratio = DEFAULT_SPACING_JITTER_RATIO;
        if (ratio < 0.1) ratio = 0.1;
        if (ratio > 1.0) ratio = 1.0;
        return ratio;
    }

    // =========================================
    // 選択と重ね順 / Selection and stacking order
    // =========================================

    /**
     * 2点以上あるパスか
     * @param {PageItem} candidate - 調べるオブジェクト
     * @returns {boolean} パスなら true
     */
    function isPathItem(candidate) {
        return candidate && candidate.typename === "PathItem" && candidate.pathPoints && candidate.pathPoints.length >= 2;
    }

    /**
     * パスに沿って配置できる種類のオブジェクトか
     * @param {PageItem} candidate - 調べるオブジェクト
     * @returns {boolean} 配置できるなら true
     */
    function isPlaceableItem(candidate) {
        if (!candidate) return false;
        var typeName = candidate.typename;
        return (
            typeName === "PathItem" ||
            typeName === "CompoundPathItem" ||
            typeName === "GroupItem" ||
            typeName === "TextFrame" ||
            typeName === "PlacedItem" ||
            typeName === "RasterItem" ||
            typeName === "SymbolItem" ||
            typeName === "MeshItem"
        );
    }

    /**
     * 選択から、基準パス以外の配置できるオブジェクトを集める
     * @param {PageItem[]} selectionItems - 選択
     * @param {PathItem} basePath - 基準パス
     * @returns {PageItem[]} 配置するオブジェクト
     */
    function collectPlaceableItems(selectionItems, basePath) {
        var placeableItems = [];
        for (var i = 0; i < selectionItems.length; i++) {
            if (selectionItems[i] === basePath) continue;
            if (isPlaceableItem(selectionItems[i])) placeableItems.push(selectionItems[i]);
        }
        return placeableItems;
    }

    /**
     * 重ね順の位置を返す
     * @param {PageItem} pageItem - 対象
     * @returns {number|null} zOrderPosition。取れなければ null
     */
    function getZOrderPosition(pageItem) {
        var zOrder = null;
        try { zOrder = pageItem.zOrderPosition; } catch (e) { /* 取れないオブジェクトがある / not available for some items */ }
        if (zOrder === null || zOrder === undefined) return null;
        return zOrder;
    }

    /**
     * 重ね順（最前面→最背面）に並べた配列を返す。重ね順が取れないものは後ろに元の順で並べる
     * @param {PageItem[]} sourceItems - 並べるオブジェクト
     * @returns {PageItem[]} 並べ替えた配列
     */
    function sortByStackingOrder(sourceItems) {
        var decoratedItems = [];
        for (var i = 0; i < sourceItems.length; i++) {
            var zOrder = getZOrderPosition(sourceItems[i]);
            decoratedItems.push({ item: sourceItems[i], zOrder: zOrder, hasZOrder: (zOrder !== null), index: i });
        }

        decoratedItems.sort(function (a, b) {
            if (a.hasZOrder && b.hasZOrder) {
                if (a.zOrder === b.zOrder) return a.index - b.index;
                return b.zOrder - a.zOrder; /* 大きいほど前面 / larger is frontmost */
            }
            if (a.hasZOrder) return -1;
            if (b.hasZOrder) return 1;
            return a.index - b.index;
        });

        var sortedItems = [];
        for (var j = 0; j < decoratedItems.length; j++) sortedItems.push(decoratedItems[j].item);
        return sortedItems;
    }

    /**
     * 控えておいた順に並べ直す。控えに無いものは後ろに付ける
     * @param {PageItem[]} baseItems - 並べ直すオブジェクト
     * @param {PageItem[]|null} storedOrder - 控えておいた順
     * @returns {PageItem[]|null} 並べ直した配列。1つも一致しなければ null（別の選択とみなす）
     */
    function reorderByStoredOrder(baseItems, storedOrder) {
        if (!storedOrder || !storedOrder.length) return null;

        var usedFlags = [];
        for (var i = 0; i < baseItems.length; i++) usedFlags[i] = false;

        var reorderedItems = [];
        var matchedCount = 0;

        for (var s = 0; s < storedOrder.length; s++) {
            for (var j = 0; j < baseItems.length; j++) {
                if (!usedFlags[j] && baseItems[j] === storedOrder[s]) {
                    usedFlags[j] = true;
                    reorderedItems.push(baseItems[j]);
                    matchedCount++;
                    break;
                }
            }
        }

        /* 残りを後ろに付ける / Append the leftovers */
        for (var k = 0; k < baseItems.length; k++) {
            if (!usedFlags[k]) reorderedItems.push(baseItems[k]);
        }

        if (matchedCount === 0) return null;
        return reorderedItems;
    }

    /**
     * 外接矩形の面積を返す（線幅を含む visibleBounds を優先）
     * @param {PageItem} pageItem - 対象
     * @returns {number} 面積
     */
    function getBoundsArea(pageItem) {
        var bounds = null;
        try { bounds = pageItem.visibleBounds; } catch (e) { bounds = null; }
        if (!bounds) {
            try { bounds = pageItem.geometricBounds; } catch (e) { bounds = null; }
        }
        if (!bounds || bounds.length < 4) return 0;
        var width = Math.abs(bounds[2] - bounds[0]);
        var height = Math.abs(bounds[1] - bounds[3]);
        return width * height;
    }

    /**
     * 面積が最も大きいパスを返す。開いたパスは外接矩形の面積で比べる
     * @param {PageItem[]} selectionItems - 選択
     * @returns {PathItem|null} パス。無ければ null
     */
    function getLargestPathItem(selectionItems) {
        var largestPath = null;
        var largestArea = null;

        for (var i = 0; i < selectionItems.length; i++) {
            var candidate = selectionItems[i];
            if (!isPathItem(candidate)) continue;

            var area = null;
            if (!candidate.closed) {
                area = getBoundsArea(candidate);
            } else {
                /* 閉じたパスは area（向きで負になる）を使う / Closed paths use area, which can be negative */
                try {
                    area = Math.abs(candidate.area);
                } catch (e) {
                    area = null;
                }
                if (area === null || area === undefined || isNaN(area) || area <= 0) {
                    area = getBoundsArea(candidate);
                }
            }

            if (largestArea === null || area > largestArea) {
                largestArea = area;
                largestPath = candidate;
            }
        }
        return largestPath;
    }

    /**
     * 重ね順で最前面または最背面のパスを返す
     * 重ね順が取れないときは、最前面なら選択の最後、最背面なら選択の最初のパスを返す
     * @param {PageItem[]} selectionItems - 選択
     * @param {boolean} wantFrontmost - 最前面なら true、最背面なら false
     * @returns {PathItem|null} パス。無ければ null
     */
    function findPathItemByStacking(selectionItems, wantFrontmost) {
        var foundPath = null;
        var foundZOrder = null;

        for (var i = 0; i < selectionItems.length; i++) {
            var candidate = selectionItems[i];
            if (!isPathItem(candidate)) continue;

            var zOrder = getZOrderPosition(candidate);
            if (zOrder === null) continue;

            if (foundZOrder === null || (wantFrontmost ? zOrder > foundZOrder : zOrder < foundZOrder)) {
                foundZOrder = zOrder;
                foundPath = candidate;
            }
        }
        if (foundPath) return foundPath;

        if (wantFrontmost) {
            for (var j = selectionItems.length - 1; j >= 0; j--) {
                if (isPathItem(selectionItems[j])) return selectionItems[j];
            }
        } else {
            for (var k = 0; k < selectionItems.length; k++) {
                if (isPathItem(selectionItems[k])) return selectionItems[k];
            }
        }
        return null;
    }

    // =========================================
    // パスの形状 / Path geometry
    // =========================================

    /**
     * 2点間の距離を返す
     * @param {{x: number, y: number}} pointA - 点A
     * @param {{x: number, y: number}} pointB - 点B
     * @returns {number} 距離
     */
    function getDistance(pointA, pointB) {
        var dx = pointB.x - pointA.x;
        var dy = pointB.y - pointA.y;
        return Math.sqrt(dx * dx + dy * dy);
    }

    /**
     * 3次ベジェ曲線上の点を返す
     * @param {{x: number, y: number}} startPoint - 始点
     * @param {{x: number, y: number}} control1 - 制御点1
     * @param {{x: number, y: number}} control2 - 制御点2
     * @param {{x: number, y: number}} endPoint - 終点
     * @param {number} t - 媒介変数（0〜1）
     * @returns {{x: number, y: number}} 曲線上の点
     */
    function cubicBezier(startPoint, control1, control2, endPoint, t) {
        var mt = 1 - t;
        var mt2 = mt * mt;
        var t2 = t * t;

        var x =
            startPoint.x * mt2 * mt +
            3 * control1.x * mt2 * t +
            3 * control2.x * mt * t2 +
            endPoint.x * t2 * t;
        var y =
            startPoint.y * mt2 * mt +
            3 * control1.y * mt2 * t +
            3 * control2.y * mt * t2 +
            endPoint.y * t2 * t;

        return { x: x, y: y };
    }

    /**
     * ベジェパスを折れ線に分割する
     * @param {PathItem} pathItem - パス
     * @param {number} samplesPerSegment - 1区間あたりのサンプル数
     * @returns {Array<{x: number, y: number}>} 折れ線の点
     */
    function buildPolylineFromBezierPath(pathItem, samplesPerSegment) {
        var pathPoints = pathItem.pathPoints;
        var isClosed = pathItem.closed;

        var polyline = [];
        var segmentCount = isClosed ? pathPoints.length : (pathPoints.length - 1);

        for (var i = 0; i < segmentCount; i++) {
            var startAnchor = pathPoints[i];
            var endAnchor = pathPoints[(i + 1) % pathPoints.length];

            var startPoint = { x: startAnchor.anchor[0], y: startAnchor.anchor[1] };
            var control1 = { x: startAnchor.rightDirection[0], y: startAnchor.rightDirection[1] };
            var control2 = { x: endAnchor.leftDirection[0], y: endAnchor.leftDirection[1] };
            var endPoint = { x: endAnchor.anchor[0], y: endAnchor.anchor[1] };

            for (var s = 0; s <= samplesPerSegment; s++) {
                /* 区間のつなぎ目は重複させない / Do not repeat the joint between segments */
                if (i > 0 && s === 0) continue;
                polyline.push(cubicBezier(startPoint, control1, control2, endPoint, s / samplesPerSegment));
            }
        }
        return polyline;
    }

    /**
     * 折れ線の各点までの累積長を返す
     * @param {Array<{x: number, y: number}>} polyline - 折れ線
     * @returns {number[]} 累積長（先頭は 0）
     */
    function getCumulativeLengths(polyline) {
        var cumulativeLengths = [0];
        for (var i = 1; i < polyline.length; i++) {
            cumulativeLengths[i] = cumulativeLengths[i - 1] + getDistance(polyline[i - 1], polyline[i]);
        }
        return cumulativeLengths;
    }

    /**
     * 始点からの距離にある折れ線上の点を返す
     * @param {Array<{x: number, y: number}>} polyline - 折れ線
     * @param {number[]} cumulativeLengths - 累積長
     * @param {number} distance - 始点からの距離
     * @returns {{x: number, y: number}} 点
     */
    function pointAtDistance(polyline, cumulativeLengths, distance) {
        if (distance <= 0) return { x: polyline[0].x, y: polyline[0].y };
        var totalLength = cumulativeLengths[cumulativeLengths.length - 1];
        if (distance >= totalLength) return { x: polyline[polyline.length - 1].x, y: polyline[polyline.length - 1].y };

        var index = 1;
        while (index < cumulativeLengths.length && cumulativeLengths[index] < distance) index++;

        var startLength = cumulativeLengths[index - 1];
        var endLength = cumulativeLengths[index];
        var ratio = (distance - startLength) / (endLength - startLength);

        var startPoint = polyline[index - 1];
        var endPoint = polyline[index];

        return {
            x: startPoint.x + (endPoint.x - startPoint.x) * ratio,
            y: startPoint.y + (endPoint.y - startPoint.y) * ratio
        };
    }

    /**
     * 始点からの距離にある接線の角度を返す（端でも安定させる）
     * @param {Array<{x: number, y: number}>} polyline - 折れ線
     * @param {number[]} cumulativeLengths - 累積長
     * @param {number} distance - 始点からの距離
     * @param {boolean} isClosed - 閉じたパスか
     * @returns {number} 角度（度）
     */
    function tangentDegAtDistance(polyline, cumulativeLengths, distance, isClosed) {
        var totalLength = cumulativeLengths[cumulativeLengths.length - 1];
        if (totalLength <= 0) return 0;

        if (distance < 0) distance = 0;
        if (distance > totalLength) distance = totalLength;

        /* cumulativeLengths[index] >= distance となる区間を探す（distance は全長以下なので index は末尾を超えない）
           Find the segment where the length reaches the distance; index never passes the last point */
        var index = 1;
        while (index < cumulativeLengths.length && cumulativeLengths[index] < distance) index++;

        var startPoint = polyline[index - 1];
        var endPoint = polyline[index];

        /* 長さ 0 の区間なら隣を使う（尖った点・重なった点で有効）/ Use a neighbor for a zero-length segment */
        if ((endPoint.x === startPoint.x) && (endPoint.y === startPoint.y)) {
            if (index + 1 < polyline.length) {
                endPoint = polyline[index + 1];
            } else if (isClosed && polyline.length > 2) {
                endPoint = polyline[1];
            }
        }

        return Math.atan2(endPoint.y - startPoint.y, endPoint.x - startPoint.x) * 180 / Math.PI;
    }

    /**
     * 均等配置の位置（始点からの距離）を返す
     * @param {number} totalLength - パスの長さ
     * @param {number} itemCount - 配置する数
     * @param {boolean} isClosed - 閉じたパスか
     * @param {boolean} useEndpoints - 開いたパスで始点・終点を含めるか
     * @returns {number[]} 距離の配列
     */
    function computeEvenDistances(totalLength, itemCount, isClosed, useEndpoints) {
        var distances = [];
        if (itemCount === 1) {
            distances.push(totalLength / 2);
        } else if (isClosed) {
            for (var i = 0; i < itemCount; i++) distances.push((totalLength * i) / itemCount);
        } else if (useEndpoints) {
            for (var j = 0; j < itemCount; j++) distances.push((totalLength * j) / (itemCount - 1));
        } else {
            for (var k = 0; k < itemCount; k++) distances.push(totalLength * (k + 0.5) / itemCount);
        }
        return distances;
    }

    /**
     * 均等配置の位置にばらつきを加えた位置を返す（始点から昇順）
     * @param {number} totalLength - パスの長さ
     * @param {number} itemCount - 配置する数
     * @param {boolean} isClosed - 閉じたパスか
     * @param {boolean} useEndpoints - 開いたパスで始点・終点を含めるか
     * @param {number} jitterRatio - ばらつきの強さ（間隔に対する比率）
     * @returns {number[]} 距離の配列
     */
    function computeJitteredDistances(totalLength, itemCount, isClosed, useEndpoints, jitterRatio) {
        /* 取りうる範囲 / Allowed range */
        var minDistance = 0;
        var maxDistance = totalLength;
        if (!isClosed && !useEndpoints) {
            var padding = totalLength * 0.02;
            if (padding > 0) {
                minDistance = padding;
                maxDistance = Math.max(padding, totalLength - padding);
            }
        }

        /* 均等配置から少しだけずらす / Start from even spacing and add a little jitter */
        var evenDistances = computeEvenDistances(totalLength, itemCount, isClosed, useEndpoints);
        var step = (itemCount <= 1) ? totalLength : (isClosed ? (totalLength / itemCount) : (useEndpoints ? (totalLength / (itemCount - 1)) : (totalLength / itemCount)));
        var jitter = step * jitterRatio;

        var distances = [];
        for (var i = 0; i < itemCount; i++) {
            var distance = evenDistances[i];
            /* 開いたパスで端を含めるときは両端を動かさない / Keep both ends fixed on an open path with endpoints */
            var isFixedEnd = !isClosed && useEndpoints && itemCount > 1 && (i === 0 || i === itemCount - 1);
            if (!isFixedEnd) distance += (Math.random() * 2 - 1) * jitter;

            if (distance < minDistance) distance = minDistance;
            if (distance > maxDistance) distance = maxDistance;
            distances.push(distance);
        }

        /* 間隔はばらつくが、並びはパスに沿って昇順 / Random gaps, but monotonic along the path */
        distances.sort(function (a, b) { return a - b; });
        return distances;
    }

    // =========================================
    // オブジェクトの操作 / Object operations
    // =========================================

    /**
     * 外接矩形の中心を返す
     * @param {PageItem} pageItem - 対象
     * @returns {{x: number, y: number}} 中心
     */
    function getItemCenter(pageItem) {
        var bounds = pageItem.geometricBounds; /* [left, top, right, bottom] */
        return { x: (bounds[0] + bounds[2]) / 2, y: (bounds[1] + bounds[3]) / 2 };
    }

    /**
     * 現在の回転角を返す
     * @param {PageItem} pageItem - 対象
     * @returns {number} 角度（度）。matrix を持たないオブジェクトは 0
     */
    function getItemRotationDeg(pageItem) {
        try {
            var itemMatrix = pageItem.matrix;
            return Math.atan2(itemMatrix.mValueB, itemMatrix.mValueA) * 180 / Math.PI;
        } catch (e) {
            /* matrix を持つのは配置画像・ラスター・テキストなどだけ / only some item types have a matrix */
            return 0;
        }
    }

    /**
     * 中心を基準に回転する（失敗しても続行）
     * @param {PageItem} pageItem - 対象
     * @param {number} angle - 回転角（度）
     * @returns {void}
     */
    function rotateItemSafely(pageItem, angle) {
        try {
            pageItem.rotate(angle, true, true, true, true, Transformation.CENTER);
        } catch (e) { }
    }

    /**
     * 指定の角度になるよう回転する
     * @param {PageItem} pageItem - 対象
     * @param {number} desiredDeg - 目標の角度（度）
     * @returns {void}
     */
    function rotateItemToDeg(pageItem, desiredDeg) {
        var delta = desiredDeg - getItemRotationDeg(pageItem);
        while (delta > 180) delta -= 360;
        while (delta < -180) delta += 360;
        rotateItemSafely(pageItem, delta);
    }

    /**
     * 塗りと線をなしにする
     * @param {PathItem} pathItem - 対象
     * @returns {void}
     */
    function clearFillAndStroke(pathItem) {
        try {
            pathItem.filled = false;
            pathItem.stroked = false;
        } catch (e) { }
    }

    /**
     * 回転の設定に従って、配置したオブジェクトを回転する
     * @param {PageItem} pageItem - 配置したオブジェクト
     * @param {{mode: string, angle: number, flip180: boolean}} rotationSettings - 回転の設定
     * @param {{x: number, y: number}} baseCenter - 基準パスの中心
     * @param {{polyline: Array, cumulativeLengths: number[], isClosed: boolean}} pathGeometry - パスの形状
     * @param {number} distance - 始点からの距離
     * @returns {void}
     */
    function rotatePlacedItem(pageItem, rotationSettings, baseCenter, pathGeometry, distance) {
        var flipAngle = rotationSettings.flip180 ? 180 : 0;
        var mode = rotationSettings.mode;

        if (mode === "angle") {
            rotateItemSafely(pageItem, (rotationSettings.angle || 0) + flipAngle);
        } else if (mode === "random") {
            rotateItemSafely(pageItem, (Math.random() * 360) - 180 + flipAngle);
        } else if (mode === "perp") {
            /* 基準パスの中心から外向きに立てる / Stand outward from the center of the base path */
            var itemCenter = getItemCenter(pageItem);
            var outwardDeg = Math.atan2(itemCenter.y - baseCenter.y, itemCenter.x - baseCenter.x) * 180 / Math.PI + 270;
            rotateItemToDeg(pageItem, outwardDeg + flipAngle);
        } else if (mode === "path_perp") {
            /* パスの接線方向に合わせる / Follow the path tangent */
            var tangentDeg = tangentDegAtDistance(pathGeometry.polyline, pathGeometry.cumulativeLengths, distance, pathGeometry.isClosed);
            rotateItemToDeg(pageItem, tangentDeg + flipAngle);
        } else if (flipAngle !== 0) {
            /* 「正立」でも反転は効かせる / Flip still applies in Upright mode */
            rotateItemSafely(pageItem, flipAngle);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを出し、OK なら選択したオブジェクトをパスに沿って配置する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            showDedupedAlert("alert.noDocument");
            return;
        }

        var doc = app.activeDocument;
        var initialSelection = doc.selection;

        /* 選択が0または1ならダイアログを出さずに終了 / Exit without the dialog when fewer than two objects are selected */
        if (!initialSelection || initialSelection.length < 2) {
            showDedupedAlert("alert.needSelection");
            return;
        }

        /* ランダム間隔の強さ（スライダーで変わる）/ Random spacing strength, changed by the slider */
        var spacingJitterRatio = DEFAULT_SPACING_JITTER_RATIO;

        /* プレビューの状態 / Preview state */
        var previewLayer = null;
        var previewSelectionSnapshot = null;   /* 元を隠して選択が外れても作り直せるように控える / kept for rebuilds while originals are hidden */
        var previewHiddenEntries = [];         /* { item, wasHidden } */

        /* ランダムの結果を控え、OK でプレビューと同じ結果にする / Keep random results so OK matches the latest preview */
        var lastRandomOrder = null;            /* PageItem[] */
        var lastRandomSpacing = null;          /* { totalLength: number, distances: number[] } */
        var lastShuffleSeed = null;            /* { seed: number, itemCount: number, duplicateCount: number } 複製込みのランダム順 */

        /* 警告の多重表示を防ぐ / Prevent stacked or repeated alerts */
        var alertLocked = false;
        var alertLastTimeByKey = {};
        var alertLastSignatureByKey = {};

        /* ダイアログのコントロール（buildDialog で作る）/ Dialog controls, created in buildDialog() */
        var rbAutoLargest, rbFrontmost, rbBackmost;
        var rbBasePathKeep, rbBasePathHide, rbBasePathDelete;
        var cbDuplicateEnabled, etDuplicateCount, sldDuplicateCount;
        var rbRotNone, rbRotPerp, rbRotPathPerp, rbRotRandom, rbRotAngle, etRotAngle, cbRotFlip180;
        var rbOrderCurrent, rbOrderReverse, rbOrderRandom;
        var rbSpacingEven, rbSpacingRandom, sldSpacingJitter;
        var cbGroupPlaced, cbAllRandom, cbPreview;

        var arrangeDialog = buildDialog();
        bindDialogEvents();
        updateDuplicateUI();
        updateRotationUI();
        updateSpacingUI();

        var dialogResult = arrangeDialog.show();

        /* 閉じたらプレビューを片付ける / Clean up the preview on close */
        clearPreview();

        if (dialogResult !== 1) return;
        applyArrangement();

        // -----------------------------------------
        // 警告 / Alerts
        // -----------------------------------------

        /**
         * 警告の重複判定に使う、選択と基準パスの規則の要約を返す
         * @returns {string} 要約
         */
        function getSelectionSignature() {
            try {
                var selectionItems = doc && doc.selection ? doc.selection : null;
                if (!selectionItems || selectionItems.length === 0) return "sel:0";
                var countsByType = {};
                for (var i = 0; i < selectionItems.length; i++) {
                    var typeName = (selectionItems[i] && selectionItems[i].typename) ? selectionItems[i].typename : "?";
                    countsByType[typeName] = (countsByType[typeName] || 0) + 1;
                }
                var signatureParts = ["sel:" + selectionItems.length];
                for (var typeKey in countsByType) {
                    if (countsByType.hasOwnProperty(typeKey)) signatureParts.push(typeKey + ":" + countsByType[typeKey]);
                }
                /* 基準パスの規則（自動／最前面／最背面）も含める / Include the base path rule */
                var ruleName = (rbAutoLargest && rbAutoLargest.value) ? "auto" : ((rbFrontmost && rbFrontmost.value) ? "front" : "back");
                signatureParts.push("rule:" + ruleName);
                return signatureParts.join("|");
            } catch (e) {
                return "sel:?";
            }
        }

        /**
         * 警告を出す。同じ選択で同じ警告は 15 秒間出し直さない
         * @param {string} labelPath - 警告文の LABELS パス
         * @returns {void}
         */
        function showDedupedAlert(labelPath) {
            if (alertLocked) return;

            var now = new Date().getTime();
            var lastTime = alertLastTimeByKey[labelPath] || 0;
            var signature = getSelectionSignature();
            var lastSignature = alertLastSignatureByKey[labelPath] || "";

            if (signature === lastSignature && (now - lastTime) < 15000) return;

            alertLastTimeByKey[labelPath] = now;
            alertLastSignatureByKey[labelPath] = signature;

            alertLocked = true;
            alert(getLabel(labelPath));
            alertLocked = false;
        }

        // -----------------------------------------
        // ダイアログ / Dialog
        // -----------------------------------------

        /**
         * ダイアログを組み立てる
         * @returns {Window} ダイアログ
         */
        function buildDialog() {
            var arrangeDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
            arrangeDialog.orientation = "column";
            arrangeDialog.alignChildren = "fill";

            setDialogOpacity(arrangeDialog, DIALOG_OPACITY);
            shiftDialogPosition(arrangeDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

            /* 2カラム / Two columns */
            var columnsGroup = arrangeDialog.add("group");
            columnsGroup.orientation = "row";
            columnsGroup.alignChildren = ["fill", "top"];
            columnsGroup.spacing = COLUMN_SPACING;

            var leftColumn = addColumn(columnsGroup);
            var rightColumn = addColumn(columnsGroup);

            addTargetPathPanel(leftColumn);

            var placeObjectsPanel = addOptionPanel(rightColumn, "panel.placeObjects", "column", "fill");
            addDuplicatePanel(placeObjectsPanel);
            addRotationPanel(placeObjectsPanel);
            addOrderPanel(placeObjectsPanel);
            addSpacingPanel(placeObjectsPanel);
            addGroupOptionsRow(placeObjectsPanel);

            addButtonRow(arrangeDialog);
            return arrangeDialog;
        }

        /**
         * 「対象パス」パネル（基準にするパスとその扱い）を作る
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function addTargetPathPanel(parentColumn) {
            var targetPathPanel = addOptionPanel(parentColumn, "panel.targetPath", "column", "left", OUTER_PANEL_MARGINS);

            /* 基準にするパス / Which path is the base */
            var basePathRulePanel = addOptionPanel(targetPathPanel, "panel.basePathRule", "column", "left");
            basePathRulePanel.helpTip = getLabel("tooltip.basePathRule");
            rbAutoLargest = basePathRulePanel.add("radiobutton", undefined, getLabel("radio.autoLargest"));
            rbAutoLargest.helpTip = getLabel("tooltip.autoLargest");
            rbFrontmost = basePathRulePanel.add("radiobutton", undefined, getLabel("radio.frontmost"));
            rbBackmost = basePathRulePanel.add("radiobutton", undefined, getLabel("radio.backmost"));
            rbAutoLargest.value = true;

            /* 基準パスの扱い / What to do with the base path */
            var basePathHandlingPanel = addOptionPanel(targetPathPanel, "panel.basePathHandling", "column", "left");
            basePathHandlingPanel.helpTip = getLabel("tooltip.basePathHandling");
            rbBasePathKeep = basePathHandlingPanel.add("radiobutton", undefined, getLabel("radio.basePathModeNone"));
            rbBasePathHide = basePathHandlingPanel.add("radiobutton", undefined, getLabel("radio.basePathModeHide"));
            rbBasePathDelete = basePathHandlingPanel.add("radiobutton", undefined, getLabel("radio.basePathModeDelete"));
            rbBasePathHide.value = true;
        }

        /**
         * 「複製」パネルを作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addDuplicatePanel(parentPanel) {
            var duplicatePanel = addOptionPanel(parentPanel, "panel.duplicate", "column", "fill");
            duplicatePanel.helpTip = getLabel("tooltip.duplicateCount");

            var duplicateCountRow = duplicatePanel.add("group");
            duplicateCountRow.orientation = "row";
            duplicateCountRow.alignChildren = ["left", "center"];
            duplicateCountRow.spacing = 10;

            cbDuplicateEnabled = duplicateCountRow.add("checkbox", undefined, "");
            /* 選択がちょうど2つのときだけ最初からオン / On by default only when exactly two objects are selected */
            cbDuplicateEnabled.value = (initialSelection.length === 2);
            cbDuplicateEnabled.preferredSize.width = 18;
            cbDuplicateEnabled.helpTip = getLabel("tooltip.duplicateEnabled");

            duplicateCountRow.add("statictext", undefined, getLabel("fieldLabel.duplicateCount"));
            etDuplicateCount = duplicateCountRow.add("edittext", undefined, String(DUPLICATE_COUNT_MIN));
            etDuplicateCount.helpTip = getLabel("tooltip.duplicateCount");
            etDuplicateCount.characters = 4;
            changeValueByArrowKey(etDuplicateCount, false);

            var duplicateSliderRow = duplicatePanel.add("group");
            duplicateSliderRow.orientation = "row";
            duplicateSliderRow.alignChildren = ["left", "center"];

            sldDuplicateCount = duplicateSliderRow.add("slider", undefined, DUPLICATE_COUNT_MIN, DUPLICATE_COUNT_MIN, DUPLICATE_COUNT_MAX);
            sldDuplicateCount.preferredSize.width = SLIDER_WIDTH;
            sldDuplicateCount.helpTip = getLabel("tooltip.duplicateCount");
        }

        /**
         * 「回転」パネルを作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addRotationPanel(parentPanel) {
            var rotationPanel = addOptionPanel(parentPanel, "panel.rotation", "column", "left");
            rotationPanel.helpTip = getLabel("tooltip.rotation");

            /* 回転のラジオは1つのグループにまとめる（排他は setRotationMode で管理）/ One group for the rotation radios; exclusivity is handled in setRotationMode() */
            var rotationRadioGroup = rotationPanel.add("group");
            rotationRadioGroup.orientation = "column";
            rotationRadioGroup.alignChildren = "left";

            rbRotNone = rotationRadioGroup.add("radiobutton", undefined, getLabel("radio.rotationNone"));
            rbRotPerp = rotationRadioGroup.add("radiobutton", undefined, getLabel("radio.rotationPerpendicular"));
            rbRotPerp.helpTip = getLabel("tooltip.rotationPerpendicular");
            rbRotPathPerp = rotationRadioGroup.add("radiobutton", undefined, getLabel("radio.rotationPathPerpendicular"));
            rbRotPathPerp.helpTip = getLabel("tooltip.rotationPathPerpendicular");
            rbRotRandom = rotationRadioGroup.add("radiobutton", undefined, getLabel("radio.rotationRandom"));
            rbRotRandom.helpTip = getLabel("tooltip.rotationRandom");
            rbRotNone.helpTip = getLabel("tooltip.rotationNone");
            rbRotNone.value = true;

            /* 角度指定（1行）/ Angle on one row */
            var rotationAngleRow = rotationRadioGroup.add("group");
            rotationAngleRow.orientation = "row";
            rotationAngleRow.alignChildren = ["left", "center"];
            rotationAngleRow.spacing = 6;

            rbRotAngle = rotationAngleRow.add("radiobutton", undefined, getLabel("radio.rotationAngle"));
            etRotAngle = rotationAngleRow.add("edittext", undefined, "0");
            etRotAngle.helpTip = getLabel("tooltip.rotationAngle");
            etRotAngle.characters = 3;
            changeValueByArrowKey(etRotAngle, true);
            rotationAngleRow.add("statictext", undefined, "°");

            cbRotFlip180 = rotationPanel.add("checkbox", undefined, getLabel("checkbox.rotationFlip180"));
            cbRotFlip180.helpTip = getLabel("tooltip.rotationFlip180");
            cbRotFlip180.value = false;
        }

        /**
         * 「順番」パネルを作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addOrderPanel(parentPanel) {
            var orderPanel = addOptionPanel(parentPanel, "panel.order", "row", ["left", "center"]);
            orderPanel.helpTip = getLabel("tooltip.order");
            orderPanel.spacing = 12;

            rbOrderCurrent = orderPanel.add("radiobutton", undefined, getLabel("radio.orderCurrent"));
            rbOrderCurrent.helpTip = getLabel("tooltip.orderCurrent");
            rbOrderReverse = orderPanel.add("radiobutton", undefined, getLabel("radio.orderReverse"));
            rbOrderReverse.helpTip = getLabel("tooltip.orderReverse");
            rbOrderRandom = orderPanel.add("radiobutton", undefined, getLabel("radio.orderRandom"));
            rbOrderCurrent.value = true;
        }

        /**
         * 「間隔」パネルを作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addSpacingPanel(parentPanel) {
            var spacingPanel = addOptionPanel(parentPanel, "panel.spacing", "column", "fill");
            spacingPanel.helpTip = getLabel("tooltip.spacing");

            var spacingRadioGroup = spacingPanel.add("group");
            spacingRadioGroup.orientation = "row";
            spacingRadioGroup.alignChildren = ["left", "center"];
            spacingRadioGroup.spacing = 12;

            rbSpacingEven = spacingRadioGroup.add("radiobutton", undefined, getLabel("radio.spacingEven"));
            rbSpacingRandom = spacingRadioGroup.add("radiobutton", undefined, getLabel("radio.spacingRandom"));
            rbSpacingEven.helpTip = getLabel("tooltip.spacingEven");
            rbSpacingRandom.helpTip = getLabel("tooltip.spacingRandom");
            rbSpacingEven.value = true;

            /* ばらつきの強さ（「ランダム」のときだけ有効）/ Jitter strength, enabled only with Random */
            var spacingSliderRow = spacingPanel.add("group");
            spacingSliderRow.orientation = "row";
            spacingSliderRow.alignChildren = ["left", "center"];

            sldSpacingJitter = spacingSliderRow.add("slider", undefined, spacingJitterRatio, 0.1, 1.0);
            sldSpacingJitter.preferredSize.width = SLIDER_WIDTH;
            sldSpacingJitter.helpTip = getLabel("tooltip.spacingJitter");
        }

        /**
         * 「グループ化」「一括ランダム」の行を中央寄せで作る
         * @param {Panel} parentPanel - 追加先
         * @returns {void}
         */
        function addGroupOptionsRow(parentPanel) {
            var groupOptionsRow = parentPanel.add("group");
            groupOptionsRow.orientation = "row";
            groupOptionsRow.alignment = "fill";
            groupOptionsRow.alignChildren = ["center", "center"];

            addStretchSpacer(groupOptionsRow);

            var groupOptionsInner = groupOptionsRow.add("group");
            groupOptionsInner.orientation = "row";
            groupOptionsInner.alignChildren = ["center", "center"];
            groupOptionsInner.spacing = 12;

            cbGroupPlaced = groupOptionsInner.add("checkbox", undefined, getLabel("checkbox.groupPlaced"));
            cbGroupPlaced.helpTip = getLabel("tooltip.groupPlaced");
            cbGroupPlaced.value = true;

            cbAllRandom = groupOptionsInner.add("checkbox", undefined, getLabel("checkbox.allRandom"));
            cbAllRandom.helpTip = getLabel("tooltip.allRandom");
            cbAllRandom.value = false;

            addStretchSpacer(groupOptionsRow);
        }

        /**
         * 下部のボタンエリア（左：プレビュー、右：キャンセル／OK）を作る
         * @param {Window} targetDialog - ダイアログ
         * @returns {void}
         */
        function addButtonRow(targetDialog) {
            var btnRowGroup = targetDialog.add("group");
            btnRowGroup.orientation = "row";
            btnRowGroup.alignment = "fill";
            btnRowGroup.alignChildren = ["left", "center"];

            var btnLeftGroup = btnRowGroup.add("group");
            btnLeftGroup.orientation = "row";
            btnLeftGroup.alignChildren = ["left", "center"];
            cbPreview = btnLeftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
            cbPreview.helpTip = getLabel("tooltip.preview");
            cbPreview.value = false;
            btnLeftGroup.margins = [0, 0, 0, 0];

            var spacer = btnRowGroup.add("group");
            spacer.alignment = ["fill", "fill"];
            spacer.minimumSize.width = 0;

            var btnRightGroup = btnRowGroup.add("group");
            btnRightGroup.orientation = "row";
            btnRightGroup.alignment = "right";
            btnRightGroup.alignChildren = ["right", "center"];

            var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

            /* キャンセルで必ず閉じる / Always close on Cancel */
            btnCancel.onClick = function () {
                targetDialog.close(0);
            };
        }

        /**
         * ダイアログのイベントを結び付ける
         * @returns {void}
         */
        function bindDialogEvents() {
            /* プレビューを作り直すだけのコントロール / Controls that only rebuild the preview */
            var previewTriggers = [
                rbAutoLargest, rbFrontmost, rbBackmost,
                rbBasePathKeep, rbBasePathHide, rbBasePathDelete,
                cbRotFlip180,
                rbOrderCurrent, rbOrderReverse, rbOrderRandom
            ];
            for (var i = 0; i < previewTriggers.length; i++) {
                previewTriggers[i].onClick = function () { rebuildPreviewIfNeeded(); };
            }

            cbDuplicateEnabled.onClick = function () {
                if (cbDuplicateEnabled.value) {
                    /* オンにしたら複製数を最小値に戻す / Reset the count when turned on */
                    etDuplicateCount.text = String(DUPLICATE_COUNT_MIN);
                    sldDuplicateCount.value = DUPLICATE_COUNT_MIN;
                }
                updateDuplicateUI();
                rebuildPreviewIfNeeded();
            };
            sldDuplicateCount.onChanging = function () {
                /* ドラッグ中は数字だけ更新 / Update only the text while dragging */
                etDuplicateCount.text = String(clampDuplicateCount(sldDuplicateCount.value));
            };
            sldDuplicateCount.onChange = function () { syncDuplicateCountFromText(); };
            etDuplicateCount.onChange = function () { syncDuplicateCountFromText(); };

            rbRotNone.onClick = function () { onRotationModeClicked("none"); };
            rbRotPerp.onClick = function () { onRotationModeClicked("perp"); };
            rbRotPathPerp.onClick = function () { onRotationModeClicked("path_perp"); };
            rbRotAngle.onClick = function () { onRotationModeClicked("angle"); };
            rbRotRandom.onClick = function () { onRotationModeClicked("random"); };
            etRotAngle.onChange = function () {
                /* 角度を編集したら「角度指定」にする / Editing the angle selects Angle mode */
                setRotationMode("angle");
                rebuildPreviewIfNeeded();
            };

            rbSpacingEven.onClick = function () {
                lastRandomSpacing = null;
                updateSpacingUI();
                rebuildPreviewIfNeeded();
            };
            rbSpacingRandom.onClick = function () {
                updateSpacingUI();
                rebuildPreviewIfNeeded();
            };
            sldSpacingJitter.onChanging = function () {
                /* ドラッグ中は値だけ更新し、プレビューは作り直さない / Update the value only; no rebuild while dragging */
                spacingJitterRatio = clampJitterRatio(sldSpacingJitter.value);
            };
            sldSpacingJitter.onChange = function () {
                spacingJitterRatio = clampJitterRatio(sldSpacingJitter.value);
                /* 新しい強さでランダム間隔を作り直す / Regenerate the random spacing with the new strength */
                lastRandomSpacing = null;
                rebuildPreviewIfNeeded();
            };

            cbAllRandom.onClick = function () { onAllRandomClicked(); };
            cbPreview.onClick = function () { onPreviewClicked(); };
        }

        // -----------------------------------------
        // ダイアログの状態 / Dialog state
        // -----------------------------------------

        /**
         * 複製の欄を、チェックボックスに合わせて有効／無効にする（値はそのまま）
         * @returns {void}
         */
        function updateDuplicateUI() {
            var isEnabled = !!cbDuplicateEnabled.value;
            etDuplicateCount.enabled = isEnabled;
            sldDuplicateCount.enabled = isEnabled;
        }

        /**
         * 複製数の欄の値を範囲内に直してスライダーにそろえ、プレビューを作り直す
         * @returns {void}
         */
        function syncDuplicateCountFromText() {
            var count = clampDuplicateCount(etDuplicateCount.text);
            etDuplicateCount.text = String(count);
            sldDuplicateCount.value = count;
            rebuildPreviewIfNeeded();
        }

        /**
         * 複製数を返す（複製しないときは 1）
         * @returns {number} 複製数
         */
        function getDuplicateCount() {
            if (!cbDuplicateEnabled.value) return 1;
            return clampDuplicateCount(etDuplicateCount.text);
        }

        /**
         * 回転モードのラジオを設定する（別グループのラジオも含めて排他にする）
         * @param {string} mode - "none" / "angle" / "perp" / "path_perp" / "random"
         * @returns {void}
         */
        function setRotationMode(mode) {
            rbRotNone.value = (mode === "none");
            rbRotAngle.value = (mode === "angle");
            rbRotPerp.value = (mode === "perp");
            rbRotPathPerp.value = (mode === "path_perp");
            rbRotRandom.value = (mode === "random");
            updateRotationUI();
        }

        /**
         * 角度の欄を「角度指定」のときだけ有効にする
         * @returns {void}
         */
        function updateRotationUI() {
            etRotAngle.enabled = !!rbRotAngle.value;
        }

        /**
         * ばらつきのスライダーを「ランダム」のときだけ有効にする
         * @returns {void}
         */
        function updateSpacingUI() {
            sldSpacingJitter.enabled = !!rbSpacingRandom.value;
        }

        /**
         * 回転の設定を読み取る
         * @returns {{mode: string, angle: number, flip180: boolean}} 回転の設定
         */
        function getRotationSettings() {
            var mode = "none";
            if (rbRotAngle.value) mode = "angle";
            else if (rbRotPerp.value) mode = "perp";
            else if (rbRotPathPerp.value) mode = "path_perp";
            else if (rbRotRandom.value) mode = "random";

            var angle = Number(etRotAngle.text);
            if (isNaN(angle)) angle = 0;

            return { mode: mode, angle: angle, flip180: !!cbRotFlip180.value };
        }

        /**
         * 並べる順番のモードを返す
         * @returns {string} "current" / "reverse" / "random"
         */
        function getOrderMode() {
            if (rbOrderReverse.value) return "reverse";
            if (rbOrderRandom.value) return "random";
            return "current";
        }

        /**
         * 間隔のモードを返す
         * @returns {string} "even" / "random"
         */
        function getSpacingMode() {
            return rbSpacingRandom.value ? "random" : "even";
        }

        /**
         * 回転のラジオをクリックしたときの処理
         * @param {string} mode - 回転モード
         * @returns {void}
         */
        function onRotationModeClicked(mode) {
            setRotationMode(mode);

            if (mode === "angle") {
                /* 角度指定にしたら 10° を入れ、すぐ ↑↓ で変えられるようにする / Default to 10° and focus the field so ↑↓ works right away */
                etRotAngle.text = "10";
                try {
                    etRotAngle.active = true;
                    etRotAngle.selection = [0, etRotAngle.text.length];
                } catch (e) { /* 環境によってフォーカス／選択範囲を設定できない / focus or selection may be unsupported */ }
            }

            rebuildPreviewIfNeeded();
        }

        /**
         * 「一括ランダム」：オンで順番・間隔・回転をランダムに、オフで既定に戻す
         * @returns {void}
         */
        function onAllRandomClicked() {
            var isAllRandom = !!cbAllRandom.value;

            setRotationMode(isAllRandom ? "random" : "none");
            if (isAllRandom) {
                rbOrderRandom.value = true;
                rbSpacingRandom.value = true;
            } else {
                rbOrderCurrent.value = true;
                rbSpacingEven.value = true;
            }
            updateSpacingUI();

            /* ランダムの結果を作り直させる / Force new random results */
            lastRandomOrder = null;
            lastRandomSpacing = null;
            if (isAllRandom) lastShuffleSeed = null;

            rebuildPreviewIfNeeded();
        }

        // -----------------------------------------
        // 基準パスと配置対象 / Base path and objects
        // -----------------------------------------

        /**
         * 選択されている規則で基準パス（B）を探す
         * @param {PageItem[]} selectionItems - 選択
         * @returns {PathItem|null} 基準パス
         */
        function findBasePath(selectionItems) {
            if (rbAutoLargest.value) return getLargestPathItem(selectionItems);
            return findPathItemByStacking(selectionItems, !!rbFrontmost.value);
        }

        /**
         * プレビューに使う選択を返す。元を隠して選択が外れたときは控えを使う
         * @returns {Array|null} 2つ以上の選択。足りなければ null
         */
        function getPreviewSourceSelection() {
            var sourceSelection = doc.selection;
            if ((!sourceSelection || sourceSelection.length < 2) && previewSelectionSnapshot && previewSelectionSnapshot.length >= 2) {
                sourceSelection = previewSelectionSnapshot;
            }
            if (!sourceSelection || sourceSelection.length < 2) return null;
            return sourceSelection;
        }

        /**
         * プレビューを作れない理由を返す
         * @returns {string} 警告文の LABELS パス。作れるなら空文字
         */
        function findPreviewProblem() {
            var sourceSelection = getPreviewSourceSelection();
            if (!sourceSelection) return "alert.needSelection";
            var basePath = findBasePath(sourceSelection);
            if (!basePath) return "alert.noBasePath";
            if (collectPlaceableItems(sourceSelection, basePath).length === 0) return "alert.noItems";
            return "";
        }

        /**
         * 並べる順番を適用した配列を返す。ランダムは OK のときにプレビューの順を使い回す
         * @param {PageItem[]} sourceItems - 並べるオブジェクト
         * @param {string} orderMode - "current" / "reverse" / "random"
         * @param {boolean} isPreview - プレビューか
         * @returns {PageItem[]} 並べた配列
         */
        function applyOrderToArray(sourceItems, orderMode, isPreview) {
            var orderedItems = sortByStackingOrder(sourceItems);

            if (orderMode !== "random") {
                lastRandomOrder = null;
                if (orderMode === "reverse") orderedItems.reverse();
                return orderedItems;
            }

            if (!isPreview) {
                var reusedItems = reorderByStoredOrder(orderedItems, lastRandomOrder);
                if (reusedItems) return reusedItems;
            }

            /* 新しく混ぜる（最新の順はプレビューが作る）/ New shuffle; the preview produces the latest order */
            var shuffledItems = copyArray(orderedItems);
            shuffleInPlace(shuffledItems);
            lastRandomOrder = copyArray(shuffledItems);
            return shuffledItems;
        }

        /**
         * 各オブジェクトを複製して、元を含めて duplicateCount 個ずつにする
         * @param {PageItem[]} sourceItems - 元のオブジェクト
         * @param {number} duplicateCount - 1つあたりの個数
         * @returns {PageItem[]} 元と複製を並べた配列
         */
        function expandWithDuplicates(sourceItems, duplicateCount) {
            var expandedItems = [];
            for (var i = 0; i < sourceItems.length; i++) {
                expandedItems.push(sourceItems[i]);
                for (var copyIndex = 1; copyIndex < duplicateCount; copyIndex++) {
                    var duplicatedItem = duplicateBeside(sourceItems[i]);
                    if (duplicatedItem) expandedItems.push(duplicatedItem);
                }
            }
            return expandedItems;
        }

        /**
         * オブジェクトを複製する。親の後ろに置けないときは現在のレイヤーの末尾に置く
         * @param {PageItem} sourceItem - 元のオブジェクト
         * @returns {PageItem|null} 複製。できなければ null
         */
        function duplicateBeside(sourceItem) {
            try {
                return sourceItem.duplicate(sourceItem.parent, ElementPlacement.PLACEAFTER);
            } catch (e) {
                try {
                    return sourceItem.duplicate(doc.activeLayer, ElementPlacement.PLACEATEND);
                } catch (err) {
                    return null;
                }
            }
        }

        // -----------------------------------------
        // 配置 / Arrange
        // -----------------------------------------

        /**
         * ランダム間隔の位置を返す。OK ではプレビューと同じ位置を使い回す
         * @param {number} totalLength - パスの長さ
         * @param {number} itemCount - 配置する数
         * @param {boolean} isClosed - 閉じたパスか
         * @param {boolean} isPreview - プレビューか
         * @returns {number[]} 距離の配列
         */
        function getRandomDistances(totalLength, itemCount, isClosed, isPreview) {
            if (!isPreview && lastRandomSpacing && lastRandomSpacing.distances && lastRandomSpacing.distances.length === itemCount &&
                Math.abs(Number(lastRandomSpacing.totalLength) - Number(totalLength)) < 0.01) {
                return copyArray(lastRandomSpacing.distances);
            }
            var distances = computeJitteredDistances(totalLength, itemCount, isClosed, USE_ENDPOINTS, clampJitterRatio(spacingJitterRatio));
            lastRandomSpacing = { totalLength: totalLength, distances: copyArray(distances) };
            return distances;
        }

        /**
         * オブジェクトをパスに沿って配置し、回転する
         * @param {PathItem} pathItem - 基準パス
         * @param {PageItem[]} targetItems - 配置するオブジェクト（この順に始点から並べる）
         * @param {{mode: string, angle: number, flip180: boolean}} rotationSettings - 回転の設定
         * @param {string} spacingMode - "even" / "random"
         * @param {boolean} isPreview - プレビューか
         * @returns {boolean} 配置できたら true
         */
        function arrangeAlongPath(pathItem, targetItems, rotationSettings, spacingMode, isPreview) {
            var pathPoints = pathItem.pathPoints;
            if (!pathPoints || pathPoints.length < 2) {
                showDedupedAlert("alert.pathTooShort");
                return false;
            }

            var polyline = buildPolylineFromBezierPath(pathItem, SAMPLES_PER_SEGMENT);
            if (polyline.length < 2) {
                showDedupedAlert("alert.pathAnalyzeFailed");
                return false;
            }

            var cumulativeLengths = getCumulativeLengths(polyline);
            var totalLength = cumulativeLengths[cumulativeLengths.length - 1];
            if (totalLength <= 0) {
                showDedupedAlert("alert.pathLengthZero");
                return false;
            }

            var itemCount = targetItems.length;
            var isClosedPath = !!pathItem.closed;
            var pathGeometry = { polyline: polyline, cumulativeLengths: cumulativeLengths, isClosed: isClosedPath };

            /* 「それぞれ垂直」の中心 / Center for the Perpendicular mode */
            var baseCenter = getItemCenter(pathItem);

            var distances;
            if (spacingMode === "random") {
                distances = getRandomDistances(totalLength, itemCount, isClosedPath, isPreview);
            } else {
                lastRandomSpacing = null;
                distances = computeEvenDistances(totalLength, itemCount, isClosedPath, USE_ENDPOINTS);
            }

            for (var j = 0; j < itemCount; j++) {
                var distance = distances[j];
                var position = pointAtDistance(polyline, cumulativeLengths, distance);

                var itemCenter = getItemCenter(targetItems[j]);
                var dx = position.x - itemCenter.x;
                var dy = position.y - itemCenter.y;
                targetItems[j].translate(dx, dy);

                rotatePlacedItem(targetItems[j], rotationSettings, baseCenter, pathGeometry, distance);

                if (ROTATE_ALONG_TANGENT) {
                    var aheadPosition = pointAtDistance(polyline, cumulativeLengths, Math.min(totalLength, distance + totalLength * 0.001));
                    var tangentAngle = Math.atan2(aheadPosition.y - position.y, aheadPosition.x - position.x) * 180 / Math.PI;
                    targetItems[j].rotate(tangentAngle, true, true, true, true, Transformation.CENTER);
                }
            }
            return true;
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------

        /**
         * プレビュー中に隠したオブジェクトの表示状態を戻す
         * @returns {void}
         */
        function restoreHiddenItems() {
            for (var i = 0; i < previewHiddenEntries.length; i++) {
                try {
                    previewHiddenEntries[i].item.hidden = previewHiddenEntries[i].wasHidden;
                } catch (e) { /* 既に無いオブジェクト / the item may be gone */ }
            }
            previewHiddenEntries = [];
        }

        /**
         * プレビュー中だけ元のオブジェクトを隠す（元の表示状態を控える）
         * @param {PageItem} pageItem - 隠すオブジェクト
         * @returns {void}
         */
        function hideForPreview(pageItem) {
            if (!pageItem) return;
            for (var i = 0; i < previewHiddenEntries.length; i++) {
                if (previewHiddenEntries[i].item === pageItem) return;
            }
            var wasHidden = false;
            try { wasHidden = !!pageItem.hidden; } catch (e) { wasHidden = false; }
            previewHiddenEntries.push({ item: pageItem, wasHidden: wasHidden });
            try { pageItem.hidden = true; } catch (e) { }
        }

        /**
         * プレビューを消し、隠したオブジェクトと選択を戻す
         * @returns {void}
         */
        function clearPreview() {
            if (previewLayer) {
                try {
                    previewLayer.locked = false;
                    previewLayer.remove();
                } catch (e) { /* 既に消えている / already gone */ }
            }
            previewLayer = null;
            restoreHiddenItems();

            /* 元を隠したときに選択が外れていたら戻す / Restore the selection if hiding the originals dropped it */
            try {
                var currentSelection = doc.selection;
                if ((!currentSelection || currentSelection.length === 0) && previewSelectionSnapshot && previewSelectionSnapshot.length) {
                    doc.selection = previewSelectionSnapshot;
                }
            } catch (e) { }

            app.redraw();
        }

        /**
         * 作りかけのプレビュー用レイヤーを捨てる
         * @returns {void}
         */
        function discardPreviewLayer() {
            try { previewLayer.remove(); } catch (e) { }
            previewLayer = null;
        }

        /**
         * 各オブジェクトを duplicateCount 個ずつ、グループの末尾に複製する
         * @param {PageItem[]} sourceItems - 元のオブジェクト
         * @param {GroupItem} targetGroup - 複製先のグループ
         * @param {number} duplicateCount - 1つあたりの複製数
         * @returns {PageItem[]} 複製
         */
        function duplicateIntoGroup(sourceItems, targetGroup, duplicateCount) {
            var duplicatedItems = [];
            for (var i = 0; i < sourceItems.length; i++) {
                for (var copyIndex = 0; copyIndex < duplicateCount; copyIndex++) {
                    try {
                        duplicatedItems.push(sourceItems[i].duplicate(targetGroup, ElementPlacement.PLACEATEND));
                    } catch (e) { /* 複製できないものは飛ばす / skip items that cannot be duplicated */ }
                }
            }
            return duplicatedItems;
        }

        /**
         * プレビューを作り直す（複製を最前面の一時レイヤーに並べ、元は隠す）
         * @returns {boolean} 作れたら true
         */
        function buildPreview() {
            clearPreview();

            var sourceSelection = getPreviewSourceSelection();
            if (!sourceSelection) return false;
            previewSelectionSnapshot = copyArray(sourceSelection);

            var basePath = findBasePath(sourceSelection);
            if (!basePath) return false;

            var sourceItems = collectPlaceableItems(sourceSelection, basePath);
            if (sourceItems.length === 0) return false;

            var orderMode = getOrderMode();
            var orderedSourceItems = applyOrderToArray(sourceItems, orderMode, true);
            var duplicateCount = getDuplicateCount();

            /* プレビュー用レイヤーを最前面に作る / Create the preview layer on top */
            try {
                previewLayer = doc.layers.add();
                previewLayer.name = PREVIEW_LAYER_NAME;
            } catch (e) {
                previewLayer = null;
                return false;
            }

            /* 形状の計算用に基準パスを複製 / Duplicate the base path for the geometry */
            var previewGroup = null;
            var previewPath = null;
            try {
                previewGroup = previewLayer.groupItems.add();
                previewGroup.name = "Preview_ArrangedAlongPath";
                previewPath = basePath.duplicate(previewGroup, ElementPlacement.PLACEATBEGINNING);
            } catch (e) {
                discardPreviewLayer();
                return false;
            }

            /* 「塗り／線」なしはプレビュー側だけに反映 / Apply "no fill / no stroke" to the preview copy only */
            if (rbBasePathHide.value) clearFillAndStroke(previewPath);

            var previewItems = duplicateIntoGroup(orderedSourceItems, previewGroup, duplicateCount);
            if (previewItems.length === 0) {
                clearPreview();
                return false;
            }

            /* ランダム順で複製ありなら、複製も含めてまとめて混ぜ、OK 用にシードを控える / Shuffle all copies together and keep the seed for OK */
            if (orderMode === "random" && duplicateCount > 1) {
                lastShuffleSeed = { seed: createShuffleSeed(), itemCount: previewItems.length, duplicateCount: duplicateCount };
                shuffleInPlace(previewItems, createSeededRandom(lastShuffleSeed.seed));
            } else {
                lastShuffleSeed = null;
            }

            if (!arrangeAlongPath(previewPath, previewItems, getRotationSettings(), getSpacingMode(), true)) {
                clearPreview();
                return false;
            }

            /* 「削除」なら配置後にプレビューのパスを消す / With Delete, remove the preview path after arranging */
            if (rbBasePathDelete.value) {
                try { previewPath.remove(); } catch (e) { }
            }

            /* プレビュー中は元のオブジェクトを隠す / Hide the originals while previewing */
            hideForPreview(basePath);
            for (var h = 0; h < sourceItems.length; h++) hideForPreview(sourceItems[h]);

            /* 選択を戻す（複製で変わることがある）/ Restore the selection, which duplication may change */
            try { doc.selection = previewSelectionSnapshot; } catch (e) { }
            app.redraw();

            return true;
        }

        /**
         * プレビューがオンなら作り直す。作れなければプレビューをオフにする
         * @returns {void}
         */
        function rebuildPreviewIfNeeded() {
            if (!cbPreview.value) return;
            var isBuilt = false;
            try {
                isBuilt = buildPreview();
            } catch (e) { /* 想定外の DOM エラーでもダイアログは閉じない / keep the dialog alive on unexpected DOM errors */ }
            if (!isBuilt) {
                cbPreview.value = false;
                clearPreview();
            }
        }

        /**
         * 「プレビュー」をクリックしたときの処理。作る前に選択を確かめ、警告が重ならないようにする
         * @returns {void}
         */
        function onPreviewClicked() {
            if (!cbPreview.value) {
                clearPreview();
                return;
            }

            var problemPath = findPreviewProblem();
            if (problemPath) {
                showDedupedAlert(problemPath);
                cbPreview.value = false;
                clearPreview();
                return;
            }

            rebuildPreviewIfNeeded();
        }

        // -----------------------------------------
        // 実行 / Apply
        // -----------------------------------------

        /**
         * OK で閉じたあと、選択したオブジェクトをパスに沿って配置する
         * @returns {void}
         */
        function applyArrangement() {
            /* OK の時点の選択を取り直す / Re-read the selection at OK time */
            var currentSelection = doc.selection;

            if (!currentSelection || currentSelection.length < 2) {
                /* プレビュー中に選択が外れていたら控えから戻す / Fall back to the snapshot if the preview dropped the selection */
                if (previewSelectionSnapshot && previewSelectionSnapshot.length >= 2) {
                    try { doc.selection = previewSelectionSnapshot; } catch (e) { }
                    try { currentSelection = doc.selection; } catch (e) { }
                }
            }

            if (!currentSelection || currentSelection.length < 2) {
                showDedupedAlert("alert.needSelection");
                return;
            }

            /* B = 選ばれた規則で決まる基準パス / B = base path by the chosen rule */
            var basePath = findBasePath(currentSelection);
            if (!basePath) {
                showDedupedAlert("alert.noBasePath");
                return;
            }

            if (rbBasePathHide.value) clearFillAndStroke(basePath);

            /* A = 基準パス（B）以外で配置できるもの / A = placeable items other than the base path (B) */
            var arrangeItems = collectPlaceableItems(currentSelection, basePath);
            if (arrangeItems.length === 0) {
                showDedupedAlert("alert.noItems");
                return;
            }

            var orderMode = getOrderMode();
            arrangeItems = applyOrderToArray(arrangeItems, orderMode, false);

            var duplicateCount = getDuplicateCount();
            if (duplicateCount > 1) {
                arrangeItems = expandWithDuplicates(arrangeItems, duplicateCount);

                /* ランダム順なら複製も含めて混ぜる。数が同じならプレビューのシードを使う / Shuffle all copies; reuse the preview seed when the counts match */
                if (orderMode === "random") {
                    var canReuseSeed = (lastShuffleSeed !== null && lastShuffleSeed.itemCount === arrangeItems.length && lastShuffleSeed.duplicateCount === duplicateCount);
                    var shuffleSeed = canReuseSeed ? lastShuffleSeed.seed : createShuffleSeed();
                    shuffleInPlace(arrangeItems, createSeededRandom(shuffleSeed));
                }
            }

            if (!arrangeAlongPath(basePath, arrangeItems, getRotationSettings(), getSpacingMode(), false)) return;

            /* 配置したオブジェクトをグループ化（基準パスは含めない）/ Group the placed objects, excluding the base path */
            if (cbGroupPlaced.value) {
                try {
                    var arrangedGroup = doc.groupItems.add();
                    arrangedGroup.name = "ArrangedAlongPath";
                    for (var i = 0; i < arrangeItems.length; i++) {
                        arrangeItems[i].move(arrangedGroup, ElementPlacement.PLACEATEND);
                    }
                } catch (e) { }
            }

            /* 基準パスを削除 / Remove the base path */
            if (rbBasePathDelete.value) {
                try { basePath.remove(); } catch (e) { }
            }
        }
    }

    main();

})();
