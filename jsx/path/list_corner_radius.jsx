#target illustrator

/*

### 概要

選択した角丸長方形の半径を、左上・右上・右下・左下の 4 つの値としてダイアログで変更します。
ダイアログには現在の半径を計測して表示し、OK でパスを指定の半径で作り直します。

### 注意

水平・垂直に置かれた長方形（角丸を含む）だけが対象です。回転した長方形、グループ、長方形以外のパスは変更しません。

### Overview

Edits the corner radii of the selected rounded rectangles as four values: top left, top right, bottom right, and bottom left.
The dialog shows the measured radii, and OK rebuilds each path with the new radii.

### Notes

Only rectangles (rounded or not) aligned to the horizontal and vertical axes are changed. Rotated rectangles, groups, and other paths are left as they are.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "EditCornerRadius";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-26";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-26";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* プレビューの初期状態 / Initial state of the preview checkbox */
    var PREVIEW_DEFAULT = true;

    // =========================================
    // 計測設定 / Measurement settings
    // =========================================

    /* 長さ・座標を 0 とみなす許容値（pt）/ Tolerance treated as zero for lengths and coordinates (pt) */
    var GEOMETRY_TOLERANCE = 0.01;

    /* 直線とみなす接線の角度（ラジアン）/ Tangent angle treated as straight (radians) */
    var MIN_ARC_ANGLE = 0.001;

    /* 90° の円弧をベジェで近似するときのハンドル長の係数 / Handle length ratio for a 90-degree Bezier arc */
    var ARC_HANDLE_RATIO = 0.5522847498;

    /* 隅の順番（LABELS.fieldLabel のキー）/ Corner order (keys of LABELS.fieldLabel) */
    var CORNER_KEYS = ["topLeft", "topRight", "bottomRight", "bottomLeft"];

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
    var PANEL_SPACING         = 12;                /* パネル内の要素間隔 */
    var COLUMN_SPACING        = 20;                /* 左右の隅の間隔 */
    var FIELD_LABEL_WIDTH     = 40;                /* 隅の名前の幅 */
    var FIELD_CHARACTERS      = 6;                 /* 半径の入力欄の文字数 */
    var BUTTON_ROW_TOP_MARGIN = 5;                 /* ボタンエリアの上余白 */

    /**
     * パネルの共通レイアウトを設定する
     * @param {Panel} targetPanel - 対象パネル
     * @param {number} [spacing] - 要素間隔
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 横並びグループの共通レイアウトを設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え
     * @param {number} [spacing] - 要素間隔
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
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

            var keyboard = ScriptUI.environment.keyboardState;
            var stepDirection = (event.keyName == "Up") ? 1 : -1;
            if (keyboard.shiftKey) {
                /* 10 の倍数にスナップ / Snap to multiples of 10 */
                currentValue = Math.floor(currentValue / 10) * 10 + stepDirection * 10;
            } else if (keyboard.altKey) {
                /* 小数第 1 位に丸める / Round to one decimal place */
                currentValue = Math.round((currentValue + stepDirection * 0.1) * 10) / 10;
            } else {
                currentValue += stepDirection;
            }
            if (!keyboard.altKey && currentValue < 0) currentValue = 0;

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
            radius: { ja: "半径", en: "Radius" }
        },
        fieldLabel: {
            topLeft: { ja: "左上", en: "Top left" },
            topRight: { ja: "右上", en: "Top right" },
            bottomRight: { ja: "右下", en: "Bottom right" },
            bottomLeft: { ja: "左下", en: "Bottom left" }
        },
        checkbox: {
            linkCorners: { ja: "連動", en: "Link" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            radiusField: {
                ja: "短辺の半分を超える値は、短辺の半分に制限されます。↑↓で増減（Shift：10、Option：0.1）",
                en: "Values over half the shorter side are limited to half of it. Up/Down to change (Shift: 10, Option: 0.1)"
            },
            linkCorners: {
                ja: "オンのときは左上の値をほかの3つの角にも使います",
                en: "When on, the top-left value is used for the other three corners"
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
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: {
                ja: "オブジェクトが選択されていません。\n角丸を変更したい長方形を選択してください。",
                en: "Nothing is selected.\nSelect the rectangles whose corners you want to change."
            },
            noTarget: {
                ja: "変更できる長方形がありません。\n水平・垂直に置かれた長方形（パス）を選択してください。グループの中身は直接選択してください。",
                en: "There are no rectangles to change.\nSelect axis-aligned rectangle paths. Select the contents of groups directly."
            }
        }
    };

    /**
     * ドット区切りのキーからラベルを取得する
     * @param {string} key - "dialog.title" のようなキー
     * @returns {string} 現在の言語のラベル
     */
    function getLabel(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return key;
            node = node[parts[i]];
        }
        if (node == null) return key;
        return node[currentLanguage] || node.en || key;
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
    function distance(fromPoint, toPoint) {
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
    function unitVector(fromPoint, toPoint) {
        var length = distance(fromPoint, toPoint);
        if (length < GEOMETRY_TOLERANCE) return null;
        return [(toPoint[0] - fromPoint[0]) / length, (toPoint[1] - fromPoint[1]) / length];
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
        if (distance(startAnchor, startHandle) < GEOMETRY_TOLERANCE &&
            distance(endAnchor, endHandle) < GEOMETRY_TOLERANCE) return null;

        /* ハンドルが片側だけのときは、もう一方のハンドルへ向かう線を接線に使う
           When only one handle exists, use the line toward the other handle as the tangent */
        var startTangent = unitVector(startAnchor, startHandle) || unitVector(startAnchor, endHandle);
        var endTangent = unitVector(endHandle, endAnchor) || unitVector(startHandle, endAnchor);
        if (!startTangent || !endTangent) return null;

        var dotProduct = startTangent[0] * endTangent[0] + startTangent[1] * endTangent[1];
        var arcAngle = Math.acos(Math.max(-1, Math.min(1, dotProduct)));
        if (arcAngle < MIN_ARC_ANGLE) return null;

        return distance(startAnchor, endAnchor) / (2 * Math.sin(arcAngle / 2));
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
            /* CORNER_KEYS の並び（左上・右上・右下・左下）に合わせる / Match CORNER_KEYS order */
            var cornerIndex = isTop ? (isLeft ? 0 : 1) : (isLeft ? 3 : 2);
            cornerRadii[cornerIndex] = radius;
        }
        return cornerRadii;
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
            var point = pathPoints[i];
            var anchor = point.anchor;
            var isOnEdge = Math.abs(anchor[0] - bounds[0]) < GEOMETRY_TOLERANCE ||
                Math.abs(anchor[0] - bounds[2]) < GEOMETRY_TOLERANCE ||
                Math.abs(anchor[1] - bounds[1]) < GEOMETRY_TOLERANCE ||
                Math.abs(anchor[1] - bounds[3]) < GEOMETRY_TOLERANCE;
            if (!isOnEdge) return false;
            if (!isAxisAligned(anchor, point.leftDirection) || !isAxisAligned(anchor, point.rightDirection)) return false;

            /* 回転した長方形は直線セグメントが斜めになる / Rotated rectangles have diagonal straight segments */
            var nextPoint = pathPoints[(i + 1) % pathPoints.length];
            if (getSegmentRadius(point, nextPoint) === null && !isAxisAligned(anchor, nextPoint.anchor)) return false;
        }
        return true;
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
            var current = pathPoints[i].anchor;
            var next = pathPoints[(i + 1) % pathPoints.length].anchor;
            signedArea += current[0] * next[1] - next[0] * current[1];
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
        var maxRadius = Math.min(right - left, top - bottom) / 2;
        /* 各隅の角と、そこへ入る向き・出る向き（時計回り）/ Corner, incoming and outgoing directions (clockwise) */
        var corners = [
            { corner: [left, top], incoming: [0, 1], outgoing: [1, 0] },
            { corner: [right, top], incoming: [1, 0], outgoing: [0, -1] },
            { corner: [right, bottom], incoming: [0, -1], outgoing: [-1, 0] },
            { corner: [left, bottom], incoming: [-1, 0], outgoing: [0, 1] }
        ];
        var pointSpecs = [];

        for (var i = 0; i < corners.length; i++) {
            var cornerPoint = corners[i].corner;
            var incoming = corners[i].incoming;
            var outgoing = corners[i].outgoing;
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
            if (previousSpec && distance(previousSpec.anchor, pointSpecs[i].anchor) < GEOMETRY_TOLERANCE) {
                previousSpec.rightDirection = pointSpecs[i].rightDirection;
            } else {
                mergedSpecs.push(pointSpecs[i]);
            }
        }
        /* 末尾と先頭の重なり / Overlap between the last and first points */
        var lastSpec = mergedSpecs[mergedSpecs.length - 1];
        if (mergedSpecs.length > 1 && distance(lastSpec.anchor, mergedSpecs[0].anchor) < GEOMETRY_TOLERANCE) {
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
    function applyCornerRadii(pathItem, cornerRadii) {
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
        for (var i = 0; i < previewCopies.length; i++) previewCopies[i].remove();
        for (var j = 0; j < hiddenOriginals.length; j++) hiddenOriginals[j].hidden = false;
        previewCopies = [];
        hiddenOriginals = [];
    }

    /**
     * 複製に半径を適用してプレビューを表示する（元のパスは一時的に隠す）
     * @param {PathItem[]} targetPaths - 対象パス
     * @param {number[]} cornerRadii - [左上, 右上, 右下, 左下] の半径（pt）
     * @returns {void}
     */
    function showPreview(targetPaths, cornerRadii) {
        clearPreview();
        for (var i = 0; i < targetPaths.length; i++) {
            /* duplicate() は hidden を引き継ぐので、隠す前に複製する / duplicate() inherits hidden, so copy first */
            var previewCopy = targetPaths[i].duplicate();
            targetPaths[i].hidden = true;
            applyCornerRadii(previewCopy, cornerRadii);
            previewCopies.push(previewCopy);
            hiddenOriginals.push(targetPaths[i]);
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
     * 隅 1 つ分の「名前：［入力欄］単位」を追加する
     * @param {Group} parentRow - 追加先の行
     * @param {string} cornerKey - CORNER_KEYS のキー
     * @param {string} unitLabel - 単位の表示
     * @returns {EditText} 追加した入力欄
     */
    function addRadiusField(parentRow, cornerKey, unitLabel) {
        var cornerGroup = parentRow.add("group");
        setupRow(cornerGroup, "left", 6);
        var cornerLabel = cornerGroup.add("statictext", undefined, labelText("fieldLabel." + cornerKey));
        cornerLabel.preferredSize.width = FIELD_LABEL_WIDTH;
        cornerLabel.justify = "right";
        var radiusField = cornerGroup.add("edittext", undefined, "0");
        radiusField.characters = FIELD_CHARACTERS;
        radiusField.helpTip = getLabel("tooltip.radiusField");
        cornerGroup.add("statictext", undefined, unitLabel);
        return radiusField;
    }

    /**
     * 角丸の半径を入力するダイアログを表示する
     * @param {PathItem[]} targetPaths - 対象パス
     * @param {number} skippedCount - 対象外のオブジェクト数
     * @returns {number[]|null} OK なら [左上, 右上, 右下, 左下] の半径（pt）、キャンセルなら null
     */
    function showRadiusDialog(targetPaths, skippedCount) {
        var unitInfo = getUnitInfo();
        var initialRadii = getCornerRadii(targetPaths[0]);

        var dialog = new Window("dialog", getLabel("dialog.title"));
        dialog.orientation = "column";
        dialog.alignChildren = ["fill", "top"];
        dialog.margins = WINDOW_MARGINS;
        dialog.spacing = WINDOW_SPACING;

        var radiusPanel = dialog.add("panel", undefined, getLabel("panel.radius"));
        setupPanel(radiusPanel);

        /* 画面上の配置に合わせて 左上・右上 / 左下・右下 の 2 行に並べる
           Lay out in two rows that mirror the corners: top left, top right / bottom left, bottom right */
        var radiusFields = [];
        var topRow = radiusPanel.add("group");
        setupRow(topRow, "left", COLUMN_SPACING);
        var bottomRow = radiusPanel.add("group");
        setupRow(bottomRow, "left", COLUMN_SPACING);
        radiusFields[0] = addRadiusField(topRow, CORNER_KEYS[0], unitInfo.label);
        radiusFields[1] = addRadiusField(topRow, CORNER_KEYS[1], unitInfo.label);
        radiusFields[3] = addRadiusField(bottomRow, CORNER_KEYS[3], unitInfo.label);
        radiusFields[2] = addRadiusField(bottomRow, CORNER_KEYS[2], unitInfo.label);

        var linkCornersCheckbox = radiusPanel.add("checkbox", undefined, getLabel("checkbox.linkCorners"));
        linkCornersCheckbox.helpTip = getLabel("tooltip.linkCorners");

        var isUniform = true;
        for (var i = 0; i < CORNER_KEYS.length; i++) {
            radiusFields[i].text = formatRadius(initialRadii[i], unitInfo.pointsPerUnit);
            if (radiusFields[i].text !== radiusFields[0].text) isUniform = false;
        }
        linkCornersCheckbox.value = isUniform;
        updateLinkedFieldsEnabled();

        if (skippedCount > 0) {
            dialog.add("statictext", undefined,
                getLabel("status.skippedCount").replace("{count}", skippedCount));
        }

        var previewCheckbox = dialog.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.value = PREVIEW_DEFAULT;

        /* ボタンエリア / Button area */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 入力欄の値を pt の配列で返す
         * @returns {number[]} [左上, 右上, 右下, 左下] の半径（pt）
         */
        function readAllRadii() {
            var radii = [];
            for (var i = 0; i < radiusFields.length; i++) {
                radii.push(readRadiusField(radiusFields[i], unitInfo.pointsPerUnit));
            }
            return radii;
        }

        /**
         * プレビューの表示状態を入力に合わせる
         * @returns {void}
         */
        function refreshPreview() {
            if (previewCheckbox.value) {
                showPreview(targetPaths, readAllRadii());
            } else {
                clearPreview();
                app.redraw();
            }
        }

        /**
         * 連動がオンなら、左上の値をほかの欄へ写す
         * @returns {void}
         */
        function syncLinkedFields() {
            if (!linkCornersCheckbox.value) return;
            for (var i = 1; i < radiusFields.length; i++) radiusFields[i].text = radiusFields[0].text;
        }

        /**
         * 連動がオンのときは左上以外の隅（名前・入力欄・単位）をディムにする
         * @returns {void}
         */
        function updateLinkedFieldsEnabled() {
            for (var i = 1; i < radiusFields.length; i++) {
                radiusFields[i].parent.enabled = !linkCornersCheckbox.value;
            }
        }

        /**
         * 入力欄の変更時の処理を登録する
         * @param {EditText} radiusField - 半径の入力欄
         * @returns {void}
         */
        function attachFieldHandlers(radiusField) {
            var onFieldChanged = function () {
                syncLinkedFields();
                refreshPreview();
            };
            radiusField.onChanging = onFieldChanged;
            changeValueByArrowKey(radiusField, onFieldChanged);
        }

        for (var j = 0; j < radiusFields.length; j++) attachFieldHandlers(radiusFields[j]);

        linkCornersCheckbox.onClick = function () {
            updateLinkedFieldsEnabled();
            syncLinkedFields();
            refreshPreview();
        };
        previewCheckbox.onClick = refreshPreview;

        dialog.onShow = function () {
            radiusFields[0].active = true;
            refreshPreview();
        };

        var dialogResult = dialog.show();
        var enteredRadii = readAllRadii();
        clearPreview();
        return (dialogResult === 1) ? enteredRadii : null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択中の角丸長方形の半径をダイアログで変更する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var selection = app.activeDocument.selection;
        if (!selection || selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var targetPaths = [];
        for (var i = 0; i < selection.length; i++) {
            if (isAxisAlignedRectangle(selection[i])) targetPaths.push(selection[i]);
        }
        if (targetPaths.length === 0) {
            alert(getLabel("alert.noTarget"));
            return;
        }

        var enteredRadii = showRadiusDialog(targetPaths, selection.length - targetPaths.length);
        if (!enteredRadii) return;
        for (var j = 0; j < targetPaths.length; j++) applyCornerRadii(targetPaths[j], enteredRadii);
    }

    main();

})();
