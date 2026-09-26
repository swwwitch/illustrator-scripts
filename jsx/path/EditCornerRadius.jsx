#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択・現在のアートボード・ドキュメント全体のいずれかにある角丸長方形の角の半径を、ダイアログで変更します。
ダイアログには現在の半径を計測して表示し、OK で角を指定の半径にそろえてパスを作り直します（半径 0 の角は、既定では角のまま残します）。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditCornerRadius.md

### 注意

水平・垂直に置かれた長方形（角丸を含む）だけが対象です。回転した長方形、長方形以外のパス、ロック・非表示のオブジェクトは変更しません。選択が対象のときは、グループの中身は直接選択してください。

### Overview

Edits the corner radius of rounded rectangles in the selection, on the current artboard, or in the entire document, using a dialog.
The dialog shows the measured radius, and OK rebuilds each path with the corners set to that radius (corners with zero radius stay square by default).

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/EditCornerRadius.md

### Notes

Only rectangles (rounded or not) aligned to the horizontal and vertical axes are changed. Rotated rectangles, other paths, and locked or hidden objects are left as they are. When the target is the selection, select the contents of groups directly.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "EditCornerRadius";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/EditCornerRadius.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/EditCornerRadius.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プレビューの初期状態 / Initial state of the preview checkbox */
    var PREVIEW_DEFAULT = true;

    /* ［0の半径は0のままに］の初期状態 / Initial state of the "keep zero radii" checkbox */
    var KEEP_ZERO_RADII_DEFAULT = true;

    /* ［「角を丸くする」効果を含む］の初期状態 / Initial state of the "include Round Corners effect" checkbox */
    var INCLUDE_EFFECT_DEFAULT = false;

    /* ［「効果」に変換］の初期状態 / Initial state of the "convert to effect" checkbox */
    var CONVERT_TO_EFFECT_DEFAULT = false;

    // =========================================
    // 計測設定 / Measurement settings
    // =========================================

    /* 長さ・座標を 0 とみなす許容値（pt）/ Tolerance treated as zero for lengths and coordinates (pt) */
    var GEOMETRY_TOLERANCE = 0.01;

    /* 直線とみなす接線の角度（ラジアン）/ Tangent angle treated as straight (radians) */
    var MIN_ARC_ANGLE = 0.001;

    /* 90° の円弧をベジェで近似するときのハンドル長の係数 / Handle length ratio for a 90-degree Bezier arc */
    var ARC_HANDLE_RATIO = 0.5522847498;

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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // レイアウト / Layout
    // =========================================
    var WINDOW_MARGINS        = 16;                /* ウィンドウ外周の余白 */
    var WINDOW_SPACING        = 12;                /* ウィンドウ内の要素間隔 */
    var PANEL_MARGINS         = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING         = 6;                 /* パネル内の要素間隔 */
    var FIELD_SPACING         = 6;                 /* 名前・入力欄・単位の間隔 */
    var FIELD_CHARACTERS      = 4;                 /* 半径の入力欄の文字数 */
    var BUTTON_ROW_TOP_MARGIN = 5;                 /* ボタンエリアの上余白 */

    /**
     * パネルの共通レイアウトを設定する
     * @param {Panel} targetPanel - 対象パネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = PANEL_SPACING;
    }

    /**
     * 横並びグループの共通レイアウトを設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = FIELD_SPACING;
    }

    /**
     * 入力欄に↑↓キーでの値増減を追加する（Shift で 10 単位、Option で 0.1 単位）
     * @param {EditText} inputField - 対象の入力欄
     * @param {function} [onValueChanged] - 値を更新したあとに呼ぶコールバック
     * @returns {void}
     */
    function changeValueByArrowKey(inputField, onValueChanged) {
        inputField.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var currentValue = Number(inputField.text);
            if (isNaN(currentValue)) return;

            var keyboardState = ScriptUI.environment.keyboardState;
            var stepDirection = (event.keyName == "Up") ? 1 : -1;
            if (keyboardState.shiftKey) {
                /* 10 の倍数にスナップ / Snap to multiples of 10 */
                currentValue = Math.floor(currentValue / 10) * 10 + stepDirection * 10;
            } else if (keyboardState.altKey) {
                /* 小数第 1 位に丸める / Round to one decimal place */
                currentValue = Math.round((currentValue + stepDirection * 0.1) * 10) / 10;
            } else {
                currentValue += stepDirection;
            }
            if (!keyboardState.altKey && currentValue < 0) currentValue = 0;

            event.preventDefault();
            inputField.text = currentValue;
            if (typeof onValueChanged === "function") onValueChanged();
        });
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var currentLanguage = ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        dialog: {
            title: { ja: "角丸の半径を変更", en: "Edit Corner Radius" }
        },
        panel: {
            targetScope: { ja: "対象", en: "Target" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            selection: { ja: "選択したオブジェクトのみ", en: "Selected objects only" },
            artboard: { ja: "現在のアートボードのみ", en: "Current artboard only" },
            document: { ja: "ドキュメント全体", en: "Entire document" }
        },
        fieldLabel: {
            radius: { ja: "半径", en: "Radius" }
        },
        checkbox: {
            keepZeroRadii: { ja: "0の半径は0のままに", en: "Keep zero radii at zero" },
            includeEffect: { ja: "「角を丸くする」効果を含む", en: "Include the Round Corners effect" },
            convertToEffect: { ja: "「角を丸くする」効果に変換", en: "Convert to Round Corners effect" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            artboard: {
                ja: "現在のアートボードに一部でも重なる長方形が対象です。グループの中も含みます（ロック・非表示は除く）",
                en: "Rectangles that overlap the current artboard, including those inside groups (locked or hidden ones are skipped)"
            },
            document: {
                ja: "ドキュメント内のすべての長方形が対象です。グループの中も含みます（ロック・非表示は除く）",
                en: "All rectangles in the document, including those inside groups (locked or hidden ones are skipped)"
            },
            radiusField: {
                ja: "短辺の半分を超える値は、短辺の半分に制限されます。↑↓で増減（Shift：10、Option：0.1）",
                en: "Values over half the shorter side are limited to half of it. Up/Down to change (Shift: 10, Option: 0.1)"
            },
            keepZeroRadii: {
                ja: "オンのときは角丸の無い角を角のまま残します",
                en: "When on, corners without rounding stay square"
            },
            includeEffect: {
                ja: "オンのときは「角を丸くする」効果で角丸になった長方形も対象にし、効果を付け直します。付け直すとほかの効果は外れます（塗り・線・不透明度は残ります）",
                en: "When on, rectangles rounded by the Round Corners effect are included and the effect is reapplied. Other effects are removed (fill, stroke and opacity are kept)"
            },
            convertToEffect: {
                ja: "オンのときは、4つの角がすべて角丸の長方形を角のない長方形に戻し、「角を丸くする」効果で角丸を付けます",
                en: "When on, rectangles with all four corners rounded are made square and rounded with the Round Corners effect"
            }
        },
        status: {
            skippedCount: {
                ja: "対象外のオブジェクト：{count} 個（水平・垂直の長方形のみ変更します）",
                en: "Skipped objects: {count} (only axis-aligned rectangles are changed)"
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * ドット区切りのキーからラベルを取得する
     * @param {string} key - "dialog.title" のようなキー
     * @returns {string} 現在の言語のラベル
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (labelNode == null) return key;
            labelNode = labelNode[keyParts[i]];
        }
        if (labelNode == null) return key;
        return labelNode[currentLanguage] || labelNode.en || key;
    }

    /**
     * コロン付きのラベルを返す（日本語は全角、英語は半角）
     * @param {string} key - ラベルのキー
     * @returns {string} コロン付きラベル
     */
    function labelText(key) {
        return getLabel(key) + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // 半径の計測 / Radius measurement
    // =========================================

    /**
     * 2 点間の距離を返す
     * @param {number[]} fromPoint - [x, y]
     * @param {number[]} toPoint - [x, y]
     * @returns {number} 距離
     */
    function getDistance(fromPoint, toPoint) {
        var dx = toPoint[0] - fromPoint[0];
        var dy = toPoint[1] - fromPoint[1];
        return Math.sqrt(dx * dx + dy * dy);
    }

    /**
     * 正規化した方向ベクトルを返す
     * @param {number[]} fromPoint - 始点 [x, y]
     * @param {number[]} toPoint - 終点 [x, y]
     * @returns {number[]|null} 単位ベクトル（長さが 0 なら null）
     */
    function getUnitVector(fromPoint, toPoint) {
        var vectorLength = getDistance(fromPoint, toPoint);
        if (vectorLength < GEOMETRY_TOLERANCE) return null;
        return [(toPoint[0] - fromPoint[0]) / vectorLength, (toPoint[1] - fromPoint[1]) / vectorLength];
    }

    /**
     * 曲線セグメントを円弧とみなして半径を求める
     * 両端の接線がなす角 θ と弦の長さ c から r = c / (2 sin(θ/2))
     * @param {PathPoint} startPoint - セグメントの始点
     * @param {PathPoint} endPoint - セグメントの終点
     * @returns {number|null} 半径（直線・算出できないときは null）
     */
    function getSegmentRadius(startPoint, endPoint) {
        var startAnchor = startPoint.anchor;
        var startHandle = startPoint.rightDirection;
        var endAnchor = endPoint.anchor;
        var endHandle = endPoint.leftDirection;
        if (getDistance(startAnchor, startHandle) < GEOMETRY_TOLERANCE &&
            getDistance(endAnchor, endHandle) < GEOMETRY_TOLERANCE) return null;

        /* ハンドルが片側だけのときは、もう一方のハンドルへ向かう線を接線に使う
           When only one handle exists, use the line toward the other handle as the tangent */
        var startTangent = getUnitVector(startAnchor, startHandle) || getUnitVector(startAnchor, endHandle);
        var endTangent = getUnitVector(endHandle, endAnchor) || getUnitVector(startHandle, endAnchor);
        if (!startTangent || !endTangent) return null;

        var dotProduct = startTangent[0] * endTangent[0] + startTangent[1] * endTangent[1];
        var arcAngle = Math.acos(Math.max(-1, Math.min(1, dotProduct)));
        if (arcAngle < MIN_ARC_ANGLE) return null;

        return getDistance(startAnchor, endAnchor) / (2 * Math.sin(arcAngle / 2));
    }

    /**
     * パスの 4 隅の角丸半径を求める
     * 弦の中点がバウンディングボックス中心のどちら側にあるかで隅を判定し、曲線の無い隅は 0 とする
     * @param {PathItem} pathItem - 対象パス
     * @returns {number[]} [左上, 右上, 右下, 左下] の半径（pt）
     */
    function getCornerRadii(pathItem) {
        var bounds = pathItem.geometricBounds; /* [左, 上, 右, 下] / [left, top, right, bottom] */
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        var cornerRadii = [0, 0, 0, 0];
        var pathPoints = pathItem.pathPoints;

        for (var i = 0; i < pathPoints.length; i++) {
            var startPoint = pathPoints[i];
            var endPoint = pathPoints[(i + 1) % pathPoints.length];
            var radius = getSegmentRadius(startPoint, endPoint);
            if (radius === null) continue;

            var isTop = (startPoint.anchor[1] + endPoint.anchor[1]) / 2 > centerY;
            var isLeft = (startPoint.anchor[0] + endPoint.anchor[0]) / 2 < centerX;
            var cornerIndex = isTop ? (isLeft ? 0 : 1) : (isLeft ? 3 : 2);
            cornerRadii[cornerIndex] = radius;
        }
        return cornerRadii;
    }

    /**
     * 角丸のある角の半径の平均を返す（角丸が 1 つも無ければ 0）
     * @param {Object[]} measurements - measureCorners() の戻り値の配列
     * @returns {number} 平均の半径（pt）
     */
    function getAverageRadius(measurements) {
        var radiusSum = 0;
        var roundedCount = 0;
        for (var i = 0; i < measurements.length; i++) {
            var cornerRadii = measurements[i].effectRadii || measurements[i].pathRadii;
            for (var j = 0; j < cornerRadii.length; j++) {
                if (cornerRadii[j] < GEOMETRY_TOLERANCE) continue;
                radiusSum += cornerRadii[j];
                roundedCount++;
            }
        }
        return (roundedCount > 0) ? radiusSum / roundedCount : 0;
    }

    /**
     * 角丸のある角（半径が 0 より大きい角）の数を返す
     * @param {number[]} cornerRadii - [左上, 右上, 右下, 左下] の半径（pt）
     * @returns {number} 角丸の数（0〜4）
     */
    function countRoundedCorners(cornerRadii) {
        var roundedCount = 0;
        for (var i = 0; i < cornerRadii.length; i++) {
            if (cornerRadii[i] >= GEOMETRY_TOLERANCE) roundedCount++;
        }
        return roundedCount;
    }

    /**
     * 範囲の短辺の半分（角丸の半径の上限）を返す
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @returns {number} 半径の上限（pt）
     */
    function getMaxRadius(bounds) {
        return Math.min(bounds[2] - bounds[0], bounds[1] - bounds[3]) / 2;
    }

    /**
     * ベクトルが水平または垂直か（長さ 0 も可）を返す
     * @param {number[]} fromPoint - 始点 [x, y]
     * @param {number[]} toPoint - 終点 [x, y]
     * @returns {boolean} 軸に沿っていれば true
     */
    function isAxisAligned(fromPoint, toPoint) {
        return Math.abs(toPoint[0] - fromPoint[0]) < GEOMETRY_TOLERANCE ||
            Math.abs(toPoint[1] - fromPoint[1]) < GEOMETRY_TOLERANCE;
    }

    /**
     * 水平・垂直に置かれた長方形（角丸を含む）のパスかを判定する
     * すべてのアンカーがバウンディングボックスの辺上にあり、ハンドルと直線セグメントが軸に沿っていること
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 対象なら true
     */
    function isAxisAlignedRectangle(pageItem) {
        if (pageItem.typename !== "PathItem" || !pageItem.closed) return false;
        var pathPoints = pageItem.pathPoints;
        if (pathPoints.length < 4 || pathPoints.length > 8) return false;

        var bounds = pageItem.geometricBounds;
        for (var i = 0; i < pathPoints.length; i++) {
            var pathPoint = pathPoints[i];
            var anchor = pathPoint.anchor;
            var isOnEdge = Math.abs(anchor[0] - bounds[0]) < GEOMETRY_TOLERANCE ||
                Math.abs(anchor[0] - bounds[2]) < GEOMETRY_TOLERANCE ||
                Math.abs(anchor[1] - bounds[1]) < GEOMETRY_TOLERANCE ||
                Math.abs(anchor[1] - bounds[3]) < GEOMETRY_TOLERANCE;
            if (!isOnEdge) return false;
            if (!isAxisAligned(anchor, pathPoint.leftDirection) || !isAxisAligned(anchor, pathPoint.rightDirection)) return false;

            /* 回転した長方形は直線セグメントが斜めになる / Rotated rectangles have diagonal straight segments */
            var nextPoint = pathPoints[(i + 1) % pathPoints.length];
            if (getSegmentRadius(pathPoint, nextPoint) === null && !isAxisAligned(anchor, nextPoint.anchor)) return false;
        }
        return true;
    }

    // =========================================
    // 対象の収集 / Target collection
    // =========================================

    /**
     * オブジェクトと親（グループ・レイヤー）がすべてロック解除・表示中かを返す
     * @param {PageItem} pageItem - 判定するオブジェクト
     * @returns {boolean} 編集できれば true
     */
    function isEditable(pageItem) {
        var ancestorItem = pageItem;
        while (ancestorItem && ancestorItem.typename !== "Document") {
            if (ancestorItem.typename === "Layer") {
                if (ancestorItem.locked || !ancestorItem.visible) return false;
            } else if (ancestorItem.locked || ancestorItem.hidden) {
                return false;
            }
            ancestorItem = ancestorItem.parent;
        }
        return true;
    }

    /**
     * 2 つの範囲が重なるかを返す
     * @param {number[]} itemBounds - [左, 上, 右, 下]
     * @param {number[]} areaBounds - [左, 上, 右, 下]
     * @returns {boolean} 重なっていれば true
     */
    function boundsOverlap(itemBounds, areaBounds) {
        return itemBounds[0] < areaBounds[2] && itemBounds[2] > areaBounds[0] &&
            itemBounds[1] > areaBounds[3] && itemBounds[3] < areaBounds[1];
    }

    /**
     * ドキュメント内の編集できる長方形を集める（グループの中も含む。複合パスの一部とガイドは除く）
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} [areaBounds] - 指定したときは、この範囲に重なるものだけ
     * @returns {PathItem[]} 対象パス
     */
    function collectDocumentRectangles(doc, areaBounds) {
        var targetPaths = [];
        var pathItems = doc.pathItems;
        for (var i = 0; i < pathItems.length; i++) {
            var pathItem = pathItems[i];
            if (pathItem.guides || pathItem.parent.typename === "CompoundPathItem") continue;
            if (!isAxisAlignedRectangle(pathItem) || !isEditable(pathItem)) continue;
            if (areaBounds && !boundsOverlap(pathItem.geometricBounds, areaBounds)) continue;
            targetPaths.push(pathItem);
        }
        return targetPaths;
    }

    // =========================================
    // パスの作り直し / Path rebuild
    // =========================================

    /**
     * アンカーの並びが反時計回り（y 上向き）かを返す
     * @param {PathItem} pathItem - 対象パス
     * @returns {boolean} 反時計回りなら true
     */
    function isCounterClockwise(pathItem) {
        var pathPoints = pathItem.pathPoints;
        var signedArea = 0;
        for (var i = 0; i < pathPoints.length; i++) {
            var currentAnchor = pathPoints[i].anchor;
            var nextAnchor = pathPoints[(i + 1) % pathPoints.length].anchor;
            signedArea += currentAnchor[0] * nextAnchor[1] - nextAnchor[0] * currentAnchor[1];
        }
        return signedArea > 0;
    }

    /**
     * 角丸長方形のアンカーとハンドルを時計回りで生成する
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @param {number[]} cornerRadii - [左上, 右上, 右下, 左下] の半径（pt）
     * @returns {Object[]} { anchor, leftDirection, rightDirection } の配列
     */
    function buildRoundedRectPoints(bounds, cornerRadii) {
        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        var maxRadius = getMaxRadius(bounds);
        /* 各隅の角と、そこへ入る向き・出る向き（時計回り）/ Corner, incoming and outgoing directions (clockwise) */
        var cornerSpecs = [
            { corner: [left, top], incoming: [0, 1], outgoing: [1, 0] },
            { corner: [right, top], incoming: [1, 0], outgoing: [0, -1] },
            { corner: [right, bottom], incoming: [0, -1], outgoing: [-1, 0] },
            { corner: [left, bottom], incoming: [-1, 0], outgoing: [0, 1] }
        ];
        var pointSpecs = [];

        for (var i = 0; i < cornerSpecs.length; i++) {
            var cornerPoint = cornerSpecs[i].corner;
            var incoming = cornerSpecs[i].incoming;
            var outgoing = cornerSpecs[i].outgoing;
            /* 短辺の半分を超えないようにする / Limit to half the shorter side */
            var radius = Math.min(Math.max(cornerRadii[i], 0), maxRadius);

            if (radius < GEOMETRY_TOLERANCE) {
                pointSpecs.push({ anchor: cornerPoint, leftDirection: cornerPoint, rightDirection: cornerPoint });
                continue;
            }
            var handleLength = radius * ARC_HANDLE_RATIO;
            var arcStart = [cornerPoint[0] - incoming[0] * radius, cornerPoint[1] - incoming[1] * radius];
            var arcEnd = [cornerPoint[0] + outgoing[0] * radius, cornerPoint[1] + outgoing[1] * radius];
            pointSpecs.push({
                anchor: arcStart,
                leftDirection: arcStart,
                rightDirection: [arcStart[0] + incoming[0] * handleLength, arcStart[1] + incoming[1] * handleLength]
            });
            pointSpecs.push({
                anchor: arcEnd,
                leftDirection: [arcEnd[0] - outgoing[0] * handleLength, arcEnd[1] - outgoing[1] * handleLength],
                rightDirection: arcEnd
            });
        }
        return pointSpecs;
    }

    /**
     * ポイントの並びを逆順にし、左右のハンドルを入れ替える
     * @param {Object[]} pointSpecs - buildRoundedRectPoints() の戻り値
     * @returns {Object[]} 逆順のポイント
     */
    function reversePointSpecs(pointSpecs) {
        var reversedSpecs = [];
        for (var i = pointSpecs.length - 1; i >= 0; i--) {
            reversedSpecs.push({
                anchor: pointSpecs[i].anchor,
                leftDirection: pointSpecs[i].rightDirection,
                rightDirection: pointSpecs[i].leftDirection
            });
        }
        return reversedSpecs;
    }

    /**
     * 同じ位置に重なった隣り合うアンカーを 1 つにまとめる（半径が短辺の半分に達したときに生じる）
     * まとめたアンカーは、前のポイントの左ハンドルと後ろのポイントの右ハンドルを持つ
     * @param {Object[]} pointSpecs - buildRoundedRectPoints() の戻り値
     * @returns {Object[]} 重なりを除いたポイント
     */
    function mergeCoincidentPoints(pointSpecs) {
        var mergedSpecs = [];
        for (var i = 0; i < pointSpecs.length; i++) {
            var previousSpec = mergedSpecs[mergedSpecs.length - 1];
            if (previousSpec && getDistance(previousSpec.anchor, pointSpecs[i].anchor) < GEOMETRY_TOLERANCE) {
                previousSpec.rightDirection = pointSpecs[i].rightDirection;
            } else {
                mergedSpecs.push(pointSpecs[i]);
            }
        }
        /* 末尾と先頭の重なり / Overlap between the last and first points */
        var lastSpec = mergedSpecs[mergedSpecs.length - 1];
        if (mergedSpecs.length > 1 && getDistance(lastSpec.anchor, mergedSpecs[0].anchor) < GEOMETRY_TOLERANCE) {
            mergedSpecs[0].leftDirection = lastSpec.leftDirection;
            mergedSpecs.pop();
        }
        return mergedSpecs;
    }

    /**
     * パスを指定の角丸半径の長方形に作り直す（元のアンカーの向きは保つ）
     * @param {PathItem} pathItem - 対象パス
     * @param {number[]} cornerRadii - [左上, 右上, 右下, 左下] の半径（pt）
     * @returns {void}
     */
    function rebuildPath(pathItem, cornerRadii) {
        var pointSpecs = mergeCoincidentPoints(buildRoundedRectPoints(pathItem.geometricBounds, cornerRadii));
        if (isCounterClockwise(pathItem)) pointSpecs = reversePointSpecs(pointSpecs);

        var anchors = [];
        for (var i = 0; i < pointSpecs.length; i++) anchors.push(pointSpecs[i].anchor);
        pathItem.setEntirePath(anchors);

        for (var j = 0; j < pointSpecs.length; j++) {
            var pathPoint = pathItem.pathPoints[j];
            pathPoint.leftDirection = pointSpecs[j].leftDirection;
            pathPoint.rightDirection = pointSpecs[j].rightDirection;
        }
    }

    /**
     * 角丸を指定の半径にする
     * - 効果で角丸になっている長方形：アピアランスを消去して効果を付け直す
     * - ［「効果」に変換］がオンで 4 つの角が角丸の長方形：角のない長方形に戻して効果を付ける
     * - それ以外：パスを作り直す（［0の半径は0のままに］なら角丸の無い角は残す）
     * @param {PathItem} pathItem - 対象パス
     * @param {Object} measurement - measureCorners() の戻り値
     * @param {number} cornerRadius - 半径（pt）
     * @param {Object} cornerOptions - { keepZeroRadii, convertToEffect }
     * @returns {PathItem} 処理後のパス（アピアランスの消去で作り直されたときは新しい参照）
     */
    function applyCornerRadius(pathItem, measurement, cornerRadius, cornerOptions) {
        if (measurement.effectRadii) {
            /* 消去でオブジェクトが作り直され、元の参照が無効になることがある
               Clearing may recreate the object and invalidate the old reference */
            var clearedPath = clearAppearance(pathItem);
            applyRoundCornersEffect(clearedPath, cornerRadius);
            return clearedPath;
        }
        /* ［0の半径は0のままに］がオンのときは変換しない（ダイアログでもディム）
           No conversion while "keep zero radii" is on (dimmed in the dialog) */
        if (cornerOptions.convertToEffect && !cornerOptions.keepZeroRadii &&
            countRoundedCorners(measurement.pathRadii) === measurement.pathRadii.length) {
            rebuildPath(pathItem, [0, 0, 0, 0]);
            applyRoundCornersEffect(pathItem, cornerRadius);
            return pathItem;
        }
        var cornerRadii = [];
        for (var i = 0; i < measurement.pathRadii.length; i++) {
            var isKept = cornerOptions.keepZeroRadii && measurement.pathRadii[i] < GEOMETRY_TOLERANCE;
            cornerRadii.push(isKept ? 0 : cornerRadius);
        }
        rebuildPath(pathItem, cornerRadii);
        return pathItem;
    }

    // =========================================
    // 「角を丸くする」効果 / Round Corners effect
    // =========================================

    /**
     * 「角を丸くする」効果の LiveEffect XML を返す
     * @param {number} radius - 半径（pt）
     * @returns {string} applyEffect() に渡す XML
     */
    function buildRoundCornersXml(radius) {
        return '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>';
    }

    /**
     * 「角を丸くする」効果を付ける（半径は短辺の半分まで、0 なら付けない）
     * @param {PathItem} pathItem - 対象パス
     * @param {number} cornerRadius - 半径（pt）
     * @returns {void}
     */
    function applyRoundCornersEffect(pathItem, cornerRadius) {
        if (cornerRadius < GEOMETRY_TOLERANCE) return;
        pathItem.applyEffect(buildRoundCornersXml(Math.min(cornerRadius, getMaxRadius(pathItem.geometricBounds))));
    }

    /**
     * カラー値を複製する
     * @param {Object} sourceColor - 複製元のカラー
     * @returns {Object|null} 複製したカラー
     */
    function cloneColorValue(sourceColor) {
        if (!sourceColor) return null;
        var clonedColor;
        switch (sourceColor.typename) {
            case "RGBColor":
                clonedColor = new RGBColor();
                clonedColor.red = sourceColor.red;
                clonedColor.green = sourceColor.green;
                clonedColor.blue = sourceColor.blue;
                return clonedColor;
            case "CMYKColor":
                clonedColor = new CMYKColor();
                clonedColor.cyan = sourceColor.cyan;
                clonedColor.magenta = sourceColor.magenta;
                clonedColor.yellow = sourceColor.yellow;
                clonedColor.black = sourceColor.black;
                return clonedColor;
            case "GrayColor":
                clonedColor = new GrayColor();
                clonedColor.gray = sourceColor.gray;
                return clonedColor;
            case "SpotColor":
                clonedColor = new SpotColor();
                clonedColor.spot = sourceColor.spot;
                clonedColor.tint = sourceColor.tint;
                return clonedColor;
            case "PatternColor":
                clonedColor = new PatternColor();
                clonedColor.pattern = sourceColor.pattern;
                return clonedColor;
            case "GradientColor":
                clonedColor = new GradientColor();
                clonedColor.gradient = sourceColor.gradient;
                clonedColor.angle = sourceColor.angle;
                clonedColor.length = sourceColor.length;
                clonedColor.matrix = sourceColor.matrix;
                clonedColor.origin = sourceColor.origin;
                clonedColor.hiliteAngle = sourceColor.hiliteAngle;
                clonedColor.hiliteLength = sourceColor.hiliteLength;
                return clonedColor;
            case "NoColor":
                return new NoColor();
            default:
                return sourceColor;
        }
    }

    /**
     * パスの基本の塗り・線・不透明度を控える
     * @param {PathItem} pathItem - 対象パス
     * @returns {Object} 控えたスタイル
     */
    function capturePathStyle(pathItem) {
        return {
            filled: pathItem.filled,
            fillColor: pathItem.filled ? cloneColorValue(pathItem.fillColor) : null,
            stroked: pathItem.stroked,
            strokeColor: pathItem.stroked ? cloneColorValue(pathItem.strokeColor) : null,
            strokeWidth: pathItem.strokeWidth,
            strokeDashes: pathItem.strokeDashes,
            strokeDashOffset: pathItem.strokeDashOffset,
            strokeCap: pathItem.strokeCap,
            strokeJoin: pathItem.strokeJoin,
            strokeMiterLimit: pathItem.strokeMiterLimit,
            opacity: pathItem.opacity,
            blendingMode: pathItem.blendingMode
        };
    }

    /**
     * 控えたスタイルをパスに戻す
     * @param {PathItem} pathItem - 対象パス
     * @param {Object} pathStyle - capturePathStyle() の戻り値
     * @returns {void}
     */
    function restorePathStyle(pathItem, pathStyle) {
        pathItem.opacity = pathStyle.opacity;
        pathItem.blendingMode = pathStyle.blendingMode;
        pathItem.filled = pathStyle.filled;
        if (pathStyle.filled && pathStyle.fillColor) pathItem.fillColor = cloneColorValue(pathStyle.fillColor);
        pathItem.stroked = pathStyle.stroked;
        if (pathStyle.stroked && pathStyle.strokeColor) {
            pathItem.strokeColor = cloneColorValue(pathStyle.strokeColor);
            pathItem.strokeWidth = pathStyle.strokeWidth;
            pathItem.strokeDashes = pathStyle.strokeDashes;
            pathItem.strokeDashOffset = pathStyle.strokeDashOffset;
            pathItem.strokeCap = pathStyle.strokeCap;
            pathItem.strokeJoin = pathStyle.strokeJoin;
            pathItem.strokeMiterLimit = pathStyle.strokeMiterLimit;
        }
    }

    /* ［アピアランスを消去］のダイナミックアクション（セット「Appearance」／アクション「clear」）
       Dynamic action for Clear Appearance (set "Appearance", action "clear") */
    var CLEAR_APPEARANCE_ACTION = [
        "/version 3",
        "/name [ 10",
        " 417070656172616e6365",
        "]",
        "/isOpen 1",
        "/actionCount 1",
        "/action-1 {",
        " /name [ 5",
        " 636c656172",
        " ]",
        " /keyIndex 0",
        " /colorIndex 0",
        " /isOpen 1",
        " /eventCount 1",
        " /event-1 {",
        " /useRulersIn1stQuadrant 0",
        " /internalName (ai_plugin_appearance)",
        " /localizedName [ 18",
        " e382a2e38394e382a2e383a9e383b3e382b9",
        " ]",
        " /isOpen 1",
        " /isOn 1",
        " /hasDialog 0",
        " /parameterCount 1",
        " /parameter-1 {",
        " /key 1835363957",
        " /showInPalette 4294967295",
        " /type (enumerated)",
        " /name [ 27",
        " e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb",
        " ]",
        " /value 6",
        " }",
        " }",
        "}"
    ].join("\n");

    /**
     * オブジェクトだけを選択する（メニューコマンドやアクションの対象にする）
     * @param {PageItem} pageItem - 対象オブジェクト（ロック・非表示でないこと）
     * @returns {void}
     */
    function selectOnly(pageItem) {
        app.activeDocument.selection = null;
        pageItem.selected = true;
    }

    /**
     * 選択の先頭を返す（何も選択されていなければ null）
     * @returns {PageItem|null} 選択の先頭
     */
    function getFirstSelectedItem() {
        var currentSelection = app.activeDocument.selection;
        return (currentSelection.length > 0) ? currentSelection[0] : null;
    }

    /**
     * ダイナミックアクションで［アピアランスを消去］を実行し、基本の塗り・線・不透明度を戻す
     * （効果を外すためのメニューコマンドは無いのでアクションで実行する）
     * @param {PathItem} pathItem - 対象パス
     * @returns {PathItem} 処理後のパス（作り直されたときは新しい参照）
     */
    function clearAppearance(pathItem) {

        var pathStyle = capturePathStyle(pathItem);
        var actionFile = new File(Folder.temp.fsName + "/EditCornerRadius_clear_" + new Date().getTime() + ".aia");
        actionFile.open("w");
        actionFile.write(CLEAR_APPEARANCE_ACTION);
        actionFile.close();
        app.loadAction(actionFile);
        /* 読み込んだ時点でパース済みなので、すぐ消す / Already parsed on load, so remove right away */
        actionFile.remove();

        var clearedPath = pathItem;
        /* doScript が失敗しても読み込んだアクションを残さない / Unload the action even if doScript fails */
        try {
            selectOnly(pathItem);
            app.doScript("clear", "Appearance", false);
            clearedPath = findFirstPathItem(getFirstSelectedItem()) || pathItem;
        } finally {
            app.unloadAction("Appearance", "");
        }
        restorePathStyle(clearedPath, pathStyle);
        return clearedPath;
    }

    /**
     * オブジェクトだけを選択してメニューコマンドを実行し、実行後の選択の先頭を返す
     * @param {PageItem} pageItem - 対象オブジェクト
     * @param {string} commandName - executeMenuCommand のコマンド名
     * @returns {PageItem|null} 実行後に選択されているオブジェクト
     */
    function runMenuCommand(pageItem, commandName) {
        selectOnly(pageItem);
        app.redraw();
        app.executeMenuCommand(commandName);
        return getFirstSelectedItem();
    }

    /**
     * グループの中から最初のパスを探す
     * @param {PageItem} pageItem - 探す対象
     * @returns {PathItem|null} 見つかったパス
     */
    function findFirstPathItem(pageItem) {
        if (!pageItem) return null;
        if (pageItem.typename === "PathItem") return pageItem;
        if (pageItem.typename !== "GroupItem") return null;
        for (var i = 0; i < pageItem.pageItems.length; i++) {
            var foundPath = findFirstPathItem(pageItem.pageItems[i]);
            if (foundPath) return foundPath;
        }
        return null;
    }

    /**
     * 効果を含めた 4 隅の角丸半径を求める（複製のアピアランスを分割して計測する）
     * @param {PathItem} pathItem - 対象パス（非表示でないこと）
     * @returns {number[]} [左上, 右上, 右下, 左下] の半径（pt）
     */
    function getEffectiveCornerRadii(pathItem) {
        var measureCopy = pathItem.duplicate();
        var expandedItem = runMenuCommand(measureCopy, "expandStyle");
        var measuredPath = findFirstPathItem(expandedItem);
        var cornerRadii = (measuredPath && isAxisAlignedRectangle(measuredPath))
            ? getCornerRadii(measuredPath) : getCornerRadii(pathItem);
        if (expandedItem) {
            expandedItem.remove();
        } else {
            measureCopy.remove();
        }
        return cornerRadii;
    }

    /**
     * パスの角丸を計測する
     * 効果を含めるときは、パスに角丸が無いものだけ複製を分割して効果の角丸を調べる
     * @param {PathItem} pathItem - 対象パス（非表示でないこと）
     * @param {boolean} includeEffect - true なら「角を丸くする」効果も調べる
     * @returns {{pathRadii: number[], effectRadii: number[]|null}} パスの半径と、効果による半径（無ければ null）
     */
    function measureCorners(pathItem, includeEffect) {
        var pathRadii = getCornerRadii(pathItem);
        var effectRadii = null;
        if (includeEffect && countRoundedCorners(pathRadii) === 0) {
            var expandedRadii = getEffectiveCornerRadii(pathItem);
            if (countRoundedCorners(expandedRadii) > 0) effectRadii = expandedRadii;
        }
        return { pathRadii: pathRadii, effectRadii: effectRadii };
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* プレビュー用の複製と、一時的に隠した元のパス / Preview copies and temporarily hidden originals */
    var previewCopies = [];
    var hiddenOriginals = [];

    /**
     * プレビューを消して、隠した元のパスを表示に戻す
     * @returns {void}
     */
    function clearPreview() {
        /* 複製の削除より先に元を表示に戻す（途中で止まっても元が隠れたままにならない）
           Unhide the originals first so they never stay hidden if a removal fails */
        for (var j = 0; j < hiddenOriginals.length; j++) hiddenOriginals[j].hidden = false;
        for (var i = 0; i < previewCopies.length; i++) previewCopies[i].remove();
        previewCopies = [];
        hiddenOriginals = [];
    }

    /**
     * 複製に半径を適用してプレビューを表示する（元のパスは一時的に隠す。前のプレビューは消してから呼ぶ）
     * @param {PathItem[]} targetPaths - 対象パス
     * @param {Object[]} measurements - パスごとの計測結果（targetPaths と同じ並び）
     * @param {number} cornerRadius - 半径（pt）
     * @param {Object} cornerOptions - { keepZeroRadii, convertToEffect }
     * @returns {void}
     */
    function showPreview(targetPaths, measurements, cornerRadius, cornerOptions) {
        for (var i = 0; i < targetPaths.length; i++) {
            /* duplicate() は hidden を引き継ぐので、隠す前に複製する / duplicate() inherits hidden, so copy first */
            var previewCopy = targetPaths[i].duplicate();
            targetPaths[i].hidden = true;
            hiddenOriginals.push(targetPaths[i]);
            previewCopies.push(applyCornerRadius(previewCopy, measurements[i], cornerRadius, cornerOptions));
        }
        app.redraw();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 入力欄の値を pt で読む（数値でない・負の値は 0）
     * @param {EditText} inputField - 半径の入力欄
     * @param {number} pointsPerUnit - 1 単位あたりの pt
     * @returns {number} 半径（pt）
     */
    function readRadiusField(inputField, pointsPerUnit) {
        var fieldValue = parseFloat(inputField.text);
        if (isNaN(fieldValue) || fieldValue < 0) return 0;
        return fieldValue * pointsPerUnit;
    }

    /**
     * pt の値を定規単位の表示用文字列にする（小数第 2 位まで）
     * @param {number} points - 値（pt）
     * @param {number} pointsPerUnit - 1 単位あたりの pt
     * @returns {string} 表示用の文字列
     */
    function formatRadius(points, pointsPerUnit) {
        return String(Math.round(points / pointsPerUnit * 100) / 100);
    }

    /**
     * 「半径：［入力欄］単位」の行を左右中央に追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} initialText - 入力欄の初期値
     * @param {string} unitLabel - 単位の表示
     * @returns {EditText} 半径の入力欄
     */
    function addRadiusRow(parentWindow, initialText, unitLabel) {
        var radiusRow = parentWindow.add("group");
        setupRow(radiusRow, "center");
        radiusRow.add("statictext", undefined, labelText("fieldLabel.radius"));
        var radiusField = radiusRow.add("edittext", undefined, initialText);
        radiusField.characters = FIELD_CHARACTERS;
        radiusField.helpTip = getLabel("tooltip.radiusField");
        radiusRow.add("statictext", undefined, unitLabel);
        return radiusField;
    }

    /**
     * ［対象］パネルとラジオボタンを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} initialScope - 最初に選ぶ対象のキー
     * @param {boolean} hasSelection - 選択に対象の長方形があるか（無ければ［選択したオブジェクトのみ］をディム）
     * @returns {Object} { selection, artboard, document } のラジオボタン
     */
    function addScopePanel(parentWindow, initialScope, hasSelection) {
        var scopePanel = parentWindow.add("panel", undefined, getLabel("panel.targetScope"));
        setupPanel(scopePanel);
        var scopeRadios = {
            selection: scopePanel.add("radiobutton", undefined, getLabel("radio.selection")),
            artboard: scopePanel.add("radiobutton", undefined, getLabel("radio.artboard")),
            document: scopePanel.add("radiobutton", undefined, getLabel("radio.document"))
        };
        scopeRadios.artboard.helpTip = getLabel("tooltip.artboard");
        scopeRadios.document.helpTip = getLabel("tooltip.document");
        scopeRadios.selection.enabled = hasSelection;
        scopeRadios[initialScope].value = true;
        return scopeRadios;
    }

    /**
     * ラベルと tooltip の付いたチェックボックスを追加する
     * @param {Object} parentContainer - 追加先（パネル・ダイアログ）
     * @param {string} labelKey - LABELS.checkbox のキー
     * @param {boolean} initialValue - 初期状態
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentContainer, labelKey, initialValue) {
        var optionCheckbox = parentContainer.add("checkbox", undefined, getLabel("checkbox." + labelKey));
        var tooltipText = getLabel("tooltip." + labelKey);
        if (tooltipText !== "tooltip." + labelKey) optionCheckbox.helpTip = tooltipText;
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * キャンセル・OK のボタン行を左右中央に追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @returns {Button} OK ボタン
     */
    function addButtonRow(parentWindow) {
        var btnRowGroup = parentWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["center", "bottom"];
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        return btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    }

    /**
     * 角丸の半径を入力するダイアログを表示する
     * @param {Object} scopeTargets - 対象ごとのパス { selection, artboard, document }（artboard・document は未収集なら null）
     * @param {function} collectScopeTargets - 対象のキーを受け取り、パスを集めて返す関数
     * @param {number} skippedCount - 選択のうち対象外のオブジェクト数
     * @returns {Object|null} OK なら { targetPaths, measurements, radius, cornerOptions }、キャンセルなら null
     */
    function showRadiusDialog(scopeTargets, collectScopeTargets, skippedCount) {
        var unitInfo = getUnitInfo();
        var hasSelection = scopeTargets.selection.length > 0;
        var currentScope = hasSelection ? "selection" : "artboard";
        var targetPaths = collectScopeTargets(currentScope);

        /* チェックボックスの状態は変数で持つ。表示前に代入した value は読み戻すと false になることがある
           Track checkbox states in variables: a value assigned before show() can read back as false */
        var cornerOptions = {
            keepZeroRadii: KEEP_ZERO_RADII_DEFAULT,
            includeEffect: INCLUDE_EFFECT_DEFAULT,
            convertToEffect: CONVERT_TO_EFFECT_DEFAULT
        };
        var isPreviewOn = PREVIEW_DEFAULT;

        /* 効果を含めた計測は重いので対象ごとに控える / Cache the effect-inclusive measurement per scope */
        var effectMeasurementsByScope = {};

        /**
         * 対象パスごとの計測結果を返す
         * プレビューの複製を消し、元のパスが表示に戻ってから呼ぶ
         * @returns {Object[]} measureCorners() の戻り値の配列
         */
        function getMeasurements() {
            var includeEffect = cornerOptions.includeEffect;
            if (includeEffect && effectMeasurementsByScope[currentScope]) return effectMeasurementsByScope[currentScope];
            var measurements = [];
            for (var i = 0; i < targetPaths.length; i++) measurements.push(measureCorners(targetPaths[i], includeEffect));
            if (includeEffect) effectMeasurementsByScope[currentScope] = measurements;
            return measurements;
        }

        var radiusDialog = new Window("dialog", getLabel("dialog.title"));
        radiusDialog.orientation = "column";
        radiusDialog.alignChildren = ["fill", "top"];
        radiusDialog.margins = WINDOW_MARGINS;
        radiusDialog.spacing = WINDOW_SPACING;

        /* 初期値は対象の角丸の平均 / Start from the average radius of the targets */
        var radiusField = addRadiusRow(radiusDialog,
            formatRadius(getAverageRadius(getMeasurements()), unitInfo.pointsPerUnit), unitInfo.label);
        var scopeRadios = addScopePanel(radiusDialog, currentScope, hasSelection);

        var optionsPanel = radiusDialog.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel);
        var keepZeroCheckbox = addOptionCheckbox(optionsPanel, "keepZeroRadii", KEEP_ZERO_RADII_DEFAULT);
        var includeEffectCheckbox = addOptionCheckbox(optionsPanel, "includeEffect", INCLUDE_EFFECT_DEFAULT);
        var convertToEffectCheckbox = addOptionCheckbox(optionsPanel, "convertToEffect", CONVERT_TO_EFFECT_DEFAULT);
        /* ［0の半径は0のままに］がオンのあいだは変換できない / Conversion is unavailable while keeping zero radii */
        convertToEffectCheckbox.enabled = !KEEP_ZERO_RADII_DEFAULT;

        /* 対象外の数は「選択したオブジェクトのみ」のときだけ表示 / Skipped count shows only for the selection scope */
        var skippedText = null;
        var skippedLabel = getLabel("status.skippedCount").replace("{count}", skippedCount);
        if (skippedCount > 0) skippedText = radiusDialog.add("statictext", undefined, skippedLabel);

        var previewCheckbox = addOptionCheckbox(radiusDialog, "preview", PREVIEW_DEFAULT);
        previewCheckbox.alignment = ["center", "top"];

        var btnOk = addButtonRow(radiusDialog);

        /**
         * プレビューの表示状態を入力に合わせる
         * @returns {void}
         */
        function refreshPreview() {
            /* 計測は元のパスが表示に戻ってから / Measure after the originals are visible again */
            clearPreview();
            if (isPreviewOn) {
                showPreview(targetPaths, getMeasurements(), readRadiusField(radiusField, unitInfo.pointsPerUnit),
                    cornerOptions);
            } else {
                app.redraw();
            }
        }

        /**
         * 対象を切り替えて、対象外の表示・OK の可否・プレビューを更新する
         * @param {string} scopeKey - "selection" / "artboard" / "document"
         * @returns {void}
         */
        function changeScope(scopeKey) {
            /* プレビューで隠した元のパスを取りこぼさないよう、集める前に消す
               Clear the preview first so hidden originals are not skipped while collecting */
            clearPreview();
            currentScope = scopeKey;
            targetPaths = collectScopeTargets(scopeKey);
            if (skippedText) skippedText.text = (scopeKey === "selection") ? skippedLabel : "";
            btnOk.enabled = targetPaths.length > 0;
            refreshPreview();
        }

        for (var scopeKey in scopeRadios) {
            scopeRadios[scopeKey].onClick = (function (clickedScope) {
                return function () { changeScope(clickedScope); };
            })(scopeKey);
        }
        radiusField.onChanging = refreshPreview;
        changeValueByArrowKey(radiusField, refreshPreview);
        keepZeroCheckbox.onClick = function () {
            cornerOptions.keepZeroRadii = keepZeroCheckbox.value;
            convertToEffectCheckbox.enabled = !cornerOptions.keepZeroRadii;
            refreshPreview();
        };
        includeEffectCheckbox.onClick = function () {
            cornerOptions.includeEffect = includeEffectCheckbox.value;
            refreshPreview();
        };
        convertToEffectCheckbox.onClick = function () {
            cornerOptions.convertToEffect = convertToEffectCheckbox.value;
            refreshPreview();
        };
        previewCheckbox.onClick = function () {
            isPreviewOn = previewCheckbox.value;
            refreshPreview();
        };

        radiusDialog.onShow = function () {
            radiusField.active = true;
            changeScope(currentScope);
        };

        var dialogResult = radiusDialog.show();
        clearPreview();
        if (dialogResult !== 1) return null;
        return {
            targetPaths: targetPaths,
            measurements: getMeasurements(),
            radius: readRadiusField(radiusField, unitInfo.pointsPerUnit),
            cornerOptions: cornerOptions
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログの入力値を対象のパスに適用する
     * @param {Object} dialogValues - showRadiusDialog() の戻り値
     * @param {PageItem[]} initialSelection - 実行前の選択
     * @returns {PageItem[]} 戻す選択（アピアランスの消去で作り直されたパスは新しい参照に置き換える）
     */
    function applyDialogValues(dialogValues, initialSelection) {
        var restoredSelection = initialSelection.slice(0);
        for (var i = 0; i < dialogValues.targetPaths.length; i++) {
            var targetPath = dialogValues.targetPaths[i];
            var resultPath = applyCornerRadius(targetPath, dialogValues.measurements[i],
                dialogValues.radius, dialogValues.cornerOptions);
            for (var j = 0; j < restoredSelection.length; j++) {
                if (restoredSelection[j] === targetPath) restoredSelection[j] = resultPath;
            }
        }
        return restoredSelection;
    }

    /**
     * 選択・アートボード・ドキュメントのいずれかにある長方形の角丸半径をダイアログで変更する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var initialSelection = doc.selection || [];

        var scopeTargets = { selection: [], artboard: null, document: null };
        for (var i = 0; i < initialSelection.length; i++) {
            if (isAxisAlignedRectangle(initialSelection[i])) scopeTargets.selection.push(initialSelection[i]);
        }

        /**
         * 対象のパスを返す（アートボード・ドキュメントは初回だけ集める）
         * @param {string} scopeKey - "selection" / "artboard" / "document"
         * @returns {PathItem[]} 対象パス
         */
        function collectScopeTargets(scopeKey) {
            if (scopeTargets[scopeKey] === null) {
                var artboardRect = null;
                if (scopeKey === "artboard") {
                    artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
                }
                scopeTargets[scopeKey] = collectDocumentRectangles(doc, artboardRect);
            }
            return scopeTargets[scopeKey];
        }

        var dialogValues = showRadiusDialog(scopeTargets, collectScopeTargets,
            initialSelection.length - scopeTargets.selection.length);
        /* 効果の計測・付け直しで変わる選択を元に戻す / Restore the selection changed by measuring and reapplying */
        var restoredSelection = dialogValues ? applyDialogValues(dialogValues, initialSelection) : initialSelection;
        doc.selection = (restoredSelection.length > 0) ? restoredSelection : null;
        app.redraw();
    }

    main();

})();
