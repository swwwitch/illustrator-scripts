#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトを「動かす対象」と「スナップ基準」に分類し、対象のバウンディングボックスを最寄りの基準線へ合わせます。
パスはアンカーポイントを直接変形するため、クリップグループ内の子パスにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FillSnapper.md

### Overview

Classifies the current selection into items to move and snap references, then snaps each item's bounding box to the nearest reference line.
Paths are transformed at the anchor level, so child paths inside clipping groups are handled too.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FillSnapper.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FillSnapper";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FillSnapper.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FillSnapper.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ダイアログの初期値（tolerance / maxDistance は pt、maxDistance 0 は距離制限なし）
       Initial dialog values (tolerance / maxDistance in pt; maxDistance 0 means no limit) */
    var DEFAULT_SNAP_OPTIONS = {
        filled: true,
        strokedOnly: true,
        blank: true,
        group: true,
        clipGroup: false,
        unrotate: true,
        tolerance: 0.5,
        maxDistance: 0,
        includeGuides: true,
        includeArtboard: true,
        preview: true
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 16;                /* ダイアログの余白 / Dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] / Panel margins [L,T,R,B] */
    var OPTION_LABEL_WIDTH = 120;           /* 数値欄の項目名の幅 / Width of the field labels */

    /**
     * パネルの共通レイアウトを設定する
     * @param {Panel} targetPanel - 設定するパネル
     * @param {number} [spacing] - 子要素の間隔
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            targetPanel.spacing = spacing;
        }
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

    /**
     * 入力欄に表示するため、数値を小数第3位で丸めて文字列にする
     * @param {number} value - 表示する数値
     * @returns {string} 整形した文字列
     */
    function formatUnitValue(value) {
        return String(Math.round(value * 1000) / 1000);
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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "塗りを線にスナップ", en: "Snap Fill to Lines" }
        },
        panel: {
            target: { ja: "動かす対象", en: "Items to move" },
            snapBasis: { ja: "スナップ基準", en: "Snap references" },
            option: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            filled: { ja: "塗りのあるクローズパス", en: "Filled closed paths" },
            group: { ja: "グループ内のパスも対象にする", en: "Include paths inside groups" },
            clipGroup: { ja: "クリップグループ内のパスも対象にする", en: "Include paths inside clipping groups" },
            strokedOnly: { ja: "線だけのパス", en: "Stroke-only paths" },
            blank: { ja: "塗り／線のないオープンパス", en: "Open paths without fill or stroke" },
            includeGuides: { ja: "ガイドライン", en: "Guide lines" },
            includeArtboard: { ja: "アートボードのエッジ", en: "Artboard edges" },
            unrotate: { ja: "回転補正", en: "Rotation correction" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        fieldLabel: {
            tolerance: { ja: "線判定の許容差", en: "Line detection tolerance" },
            maxDistance: { ja: "最大スナップ距離", en: "Max snap distance" }
        },
        tooltip: {
            filled: { ja: "塗りのあるオブジェクトを対象にします。", en: "Targets objects that have a fill." },
            group: { ja: "グループも対象にします。", en: "Targets groups as well." },
            clipGroup: { ja: "クリップグループも対象にします。", en: "Targets clipping groups as well." },
            strokedOnly: {
                ja: "線だけのオブジェクトを、吸着先の基準として使います。",
                en: "Uses stroke-only objects as the edges to snap to."
            },
            blank: {
                ja: "塗りも線もないオブジェクトも基準に含めます。",
                en: "Includes objects with neither fill nor stroke as edges."
            },
            includeGuides: { ja: "ガイドも吸着先の基準に含めます。", en: "Includes guides as edges to snap to." },
            includeArtboard: {
                ja: "アートボードの端も吸着先の基準に含めます。",
                en: "Includes the artboard edges as edges to snap to."
            },
            unrotate: {
                ja: "回転しているオブジェクトを、いったん角度0に戻してから合わせます。",
                en: "Straightens rotated objects before snapping them."
            },
            tolerance: {
                ja: "同じ位置とみなす許容差です。",
                en: "How far apart two edges can be and still count as aligned."
            },
            maxDistance: { ja: "この距離までの基準にだけ吸着します。", en: "Only snaps to edges within this distance." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original layout."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectTarget: { ja: "スナップ対象のオブジェクトを選択してください。", en: "Select a snap target object." },
            noTarget: { ja: "変形対象が見つかりません。", en: "No transformable target found." },
            noReferenceLines: {
                ja: "基準となる罫線や枠（塗りなしパス）が見つかりません。",
                en: "No reference lines or frames (unfilled paths) found."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) {
                return labelPath;
            }
        }
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // 共通ユーティリティ / Common utilities
    // =========================================

    /**
     * 配列の末尾に別の配列（またはコレクション）の要素をすべて追加する
     * @param {Array} targetList - 追加先の配列
     * @param {Array} sourceList - 追加する要素
     * @returns {void}
     */
    function appendAll(targetList, sourceList) {
        for (var i = 0; i < sourceList.length; i++) {
            targetList.push(sourceList[i]);
        }
    }

    /**
     * 選択からスナップの処理対象候補を再帰的に集める（グループ・複合パスは中身を展開）
     * @param {PageItem[]} pageItems - 走査するアイテム
     * @param {{includeNormalGroups: boolean, includeClipGroups: boolean}} groupOptions - グループを展開するかどうか
     * @returns {PageItem[]} 処理対象候補
     */
    function collectProcessableItems(pageItems, groupOptions) {
        var includeNormalGroups = (groupOptions.includeNormalGroups !== false);
        var includeClipGroups = (groupOptions.includeClipGroups === true);
        var processableItems = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename === "GroupItem") {
                if (pageItem.clipped === true) {
                    /* クリップグループ：マスクパスは基準線化しやすいため除外し、中身だけを再帰 / Clip group: skip the clipping mask path and recurse into contents */
                    if (!includeClipGroups) continue;
                    var clipGroupChildren = pageItem.pageItems;
                    for (var j = 0; j < clipGroupChildren.length; j++) {
                        var childItem = clipGroupChildren[j];
                        if (childItem.typename === "PathItem" && childItem.clipping === true) continue;
                        appendAll(processableItems, collectProcessableItems([childItem], groupOptions));
                    }
                } else if (includeNormalGroups) {
                    appendAll(processableItems, collectProcessableItems(pageItem.pageItems, groupOptions));
                }
            } else if (pageItem.typename === "CompoundPathItem") {
                appendAll(processableItems, collectProcessableItems(pageItem.pathItems, groupOptions));
            } else {
                processableItems.push(pageItem);
            }
        }
        return processableItems;
    }

    /**
     * アイテムが編集できるかどうかを、親グループとレイヤーのロック・表示までたどって判定する
     * @param {PageItem} pageItem - 判定するアイテム
     * @returns {boolean} 編集できるとき true
     */
    function isEditableItem(pageItem) {
        /* 親をたどる途中でプロパティを持たない型に当たることがある / some parents may lack these properties */
        try {
            if (!pageItem) return false;
            if (pageItem.locked === true || pageItem.hidden === true) return false;

            var parentItem = pageItem.parent;
            while (parentItem && parentItem.typename !== "Document") {
                if (parentItem.locked === true || parentItem.hidden === true) return false;
                parentItem = parentItem.parent;
            }

            if (pageItem.layer) {
                if (pageItem.layer.locked === true || pageItem.layer.visible === false) return false;
            }
        } catch (e) {
            return false;
        }
        return true;
    }

    /**
     * プレビューの復元用に、編集できるパスの元の形状を保存する
     * @param {PageItem[]} pageItems - 保存する候補
     * @returns {Object[]} パスとアンカーポイントの座標の組
     */
    function captureOriginalGeometry(pageItems) {
        var snapshotData = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename !== "PathItem") continue;
            if (!isEditableItem(pageItem)) continue;

            /* 読み取れないパスは保存せずに飛ばす / skip paths whose points cannot be read */
            var savedPoints = [];
            try {
                var pathPointList = pageItem.pathPoints;
                for (var j = 0; j < pathPointList.length; j++) {
                    var pathPoint = pathPointList[j];
                    savedPoints.push({
                        anchor: [pathPoint.anchor[0], pathPoint.anchor[1]],
                        leftDirection: [pathPoint.leftDirection[0], pathPoint.leftDirection[1]],
                        rightDirection: [pathPoint.rightDirection[0], pathPoint.rightDirection[1]]
                    });
                }
            } catch (e) {
                savedPoints = [];
            }
            if (savedPoints.length > 0) {
                snapshotData.push({ item: pageItem, points: savedPoints });
            }
        }
        return snapshotData;
    }

    /**
     * 保存した元の形状へ戻す
     * @param {Object[]} snapshotData - captureOriginalGeometry() の戻り値
     * @returns {void}
     */
    function restoreOriginalGeometry(snapshotData) {
        for (var i = 0; i < snapshotData.length; i++) {
            var snapshotEntry = snapshotData[i];
            if (!snapshotEntry || !snapshotEntry.item || !isEditableItem(snapshotEntry.item)) continue;

            /* 書き戻せないパスはそこで打ち切り、次のパスへ / stop at a path that cannot be written and move on */
            try {
                var pathPointList = snapshotEntry.item.pathPoints;
                var pointCount = Math.min(snapshotEntry.points.length, pathPointList.length);
                for (var j = 0; j < pointCount; j++) {
                    var savedPoint = snapshotEntry.points[j];
                    pathPointList[j].anchor = savedPoint.anchor;
                    pathPointList[j].leftDirection = savedPoint.leftDirection;
                    pathPointList[j].rightDirection = savedPoint.rightDirection;
                }
            } catch (e) {
                /* 続行 / continue with the next path */
            }
        }
    }

    /**
     * 候補値の中から最も近い値を返す（最大距離を超えるときは元の値のまま）
     * @param {number} targetValue - 基準の値
     * @param {number[]} candidates - 候補の値
     * @param {number} maxDistance - 最大距離。0 なら制限なし
     * @returns {number} 最も近い候補、または targetValue
     */
    function findClosestCoordinate(targetValue, candidates, maxDistance) {
        if (!candidates || candidates.length === 0) return targetValue;
        var closest = candidates[0];
        var minDiff = Math.abs(targetValue - closest);
        for (var i = 1; i < candidates.length; i++) {
            var diff = Math.abs(targetValue - candidates[i]);
            if (diff < minDiff) {
                minDiff = diff;
                closest = candidates[i];
            }
        }
        if (maxDistance && maxDistance > 0 && minDiff > maxDistance) return targetValue;
        return closest;
    }

    /**
     * パスのアンカーポイントとハンドルを、指定の範囲に収まるよう直接拡大・縮小して移動する
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} left - 範囲の左端
     * @param {number} top - 範囲の上端
     * @param {number} right - 範囲の右端
     * @param {number} bottom - 範囲の下端
     * @returns {boolean} 変形できたとき true
     */
    function fitPathItemToTargetBounds(pathItem, left, top, right, bottom) {
        var geometricBounds = pathItem.geometricBounds; /* [left, top, right, bottom] */
        var sourceLeft = geometricBounds[0];
        var sourceTop = geometricBounds[1];
        var sourceWidth = geometricBounds[2] - sourceLeft;
        var sourceHeight = sourceTop - geometricBounds[3];
        var targetWidthPt = right - left;
        var targetHeightPt = top - bottom;

        if (Math.abs(sourceWidth) <= 0.0001 || Math.abs(sourceHeight) <= 0.0001) return false;
        if (Math.abs(targetWidthPt) <= 0.0001 || Math.abs(targetHeightPt) <= 0.0001) return false;

        var scaleX = targetWidthPt / sourceWidth;
        var scaleY = targetHeightPt / sourceHeight;
        var pathPointList = pathItem.pathPoints;

        /**
         * 元の範囲の座標を、指定の範囲の座標に写す
         * @param {number[]} pointArray - [x, y]
         * @returns {number[]} 写した [x, y]
         */
        function mapPoint(pointArray) {
            return [
                left + (pointArray[0] - sourceLeft) * scaleX,
                top - (sourceTop - pointArray[1]) * scaleY
            ];
        }

        /* アンカーの書き込みに失敗したら失敗扱い / treat a failed anchor write as failure */
        try {
            for (var i = 0; i < pathPointList.length; i++) {
                var pathPoint = pathPointList[i];
                pathPoint.anchor = mapPoint(pathPoint.anchor);
                pathPoint.leftDirection = mapPoint(pathPoint.leftDirection);
                pathPoint.rightDirection = mapPoint(pathPoint.rightDirection);
            }
        } catch (e) {
            return false;
        }
        return true;
    }

    /**
     * 外接矩形から基準線を追加する（細い縦長は左端、細い横長は上端、それ以外は4辺）
     * @param {number[]} geometricBounds - [left, top, right, bottom]
     * @param {number} tolerance - 線とみなす幅（pt）
     * @param {{horizontalLines: number[], verticalLines: number[]}} referenceLines - 追加先
     * @returns {void}
     */
    function addReferenceEdges(geometricBounds, tolerance, referenceLines) {
        var boundsWidthPt = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var boundsHeightPt = Math.abs(geometricBounds[1] - geometricBounds[3]);
        if (boundsWidthPt < tolerance) {
            referenceLines.verticalLines.push(geometricBounds[0]);
        } else if (boundsHeightPt < tolerance) {
            referenceLines.horizontalLines.push(geometricBounds[1]);
        } else {
            /* 矩形などは4辺すべてを採用 / Rectangles: use all four sides */
            referenceLines.horizontalLines.push(geometricBounds[1]);
            referenceLines.horizontalLines.push(geometricBounds[3]);
            referenceLines.verticalLines.push(geometricBounds[0]);
            referenceLines.verticalLines.push(geometricBounds[2]);
        }
    }

    /**
     * アクティブなアートボードの4辺を基準線として返す
     * @param {Document} doc - 対象のドキュメント
     * @returns {{horizontalLines: number[], verticalLines: number[]}} 基準線
     */
    function gatherArtboardEdges(doc) {
        var artboardIndex;
        /* 取得できない環境では先頭のアートボード / fall back to the first artboard */
        try { artboardIndex = doc.artboards.getActiveArtboardIndex(); } catch (e) { artboardIndex = 0; }
        var artboardRect = doc.artboards[artboardIndex].artboardRect; /* [left, top, right, bottom] */
        return {
            horizontalLines: [artboardRect[1], artboardRect[3]],
            verticalLines: [artboardRect[0], artboardRect[2]]
        };
    }

    /**
     * ドキュメント内のガイドから水平線・垂直線を集める
     * @param {Document} doc - 対象のドキュメント
     * @param {number} tolerance - 線とみなす幅（pt）
     * @returns {{horizontalLines: number[], verticalLines: number[]}} 基準線
     */
    function gatherDocumentGuides(doc, tolerance) {
        var guideLines = { horizontalLines: [], verticalLines: [] };
        var documentPathItems = doc.pathItems;
        for (var i = 0; i < documentPathItems.length; i++) {
            var guidePathItem = documentPathItems[i];
            if (guidePathItem.guides !== true) continue;
            addReferenceEdges(guidePathItem.geometricBounds, tolerance, guideLines);
        }
        return guideLines;
    }

    // =========================================
    // 角度・回転補正 / Angle & Rotation correction
    // =========================================

    /**
     * パスの最初の辺の角度を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 角度（度）。点が2つ未満なら 0
     */
    function getEdgeAngleDeg(pathItem) {
        var pathPoints = pathItem.pathPoints;
        if (!pathPoints || pathPoints.length < 2) return 0;
        var firstAnchor = pathPoints[0].anchor;
        var secondAnchor = pathPoints[1].anchor;
        return Math.atan2(secondAnchor[1] - firstAnchor[1], secondAnchor[0] - firstAnchor[0]) * 180 / Math.PI;
    }

    /**
     * 最寄りの 90 度の倍数からのずれを返す
     * @param {number} angleDeg - 角度（度）
     * @returns {number} -45〜45 のずれ（度）
     */
    function getNearestRightAngleOffset(angleDeg) {
        var modulus = angleDeg % 90;
        if (modulus > 45) modulus -= 90;
        if (modulus < -45) modulus += 90;
        return modulus;
    }

    // =========================================
    // スナップ判定・実行 / Snap classification & execution
    // =========================================

    /**
     * 処理対象候補を、動かす対象と基準線に振り分ける
     * @param {PageItem[]} pageItems - 処理対象候補
     * @param {Object} snapOptions - ダイアログの設定（readSnapOptions() の戻り値）
     * @returns {{targets: PathItem[], horizontalLines: number[], verticalLines: number[]}} 振り分けの結果
     */
    function classifyTargetsAndReferences(pageItems, snapOptions) {
        var tolerance = snapOptions.tolerance;
        var classification = { targets: [], horizontalLines: [], verticalLines: [] };

        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            if (pageItem.typename !== "PathItem") continue;

            var geometricBounds = pageItem.geometricBounds; /* [left, top, right, bottom] */
            var itemWidthPt = Math.abs(geometricBounds[2] - geometricBounds[0]);
            var itemHeightPt = Math.abs(geometricBounds[1] - geometricBounds[3]);

            var isFilledClosedPath = pageItem.filled && pageItem.closed;
            var isStrokedOpenPath = !pageItem.filled && pageItem.stroked && !pageItem.closed;
            var isBlankOpenPath = !pageItem.filled && !pageItem.stroked && !pageItem.closed;

            /* 変形対象：塗りのあるクローズパスのみ / Targets: filled closed paths only */
            if (isFilledClosedPath) {
                if (!snapOptions.filled) continue;
                if (itemWidthPt > tolerance && itemHeightPt > tolerance) {
                    classification.targets.push(pageItem);
                    continue;
                }
                /* 極端に細い filled closed は基準として扱う / Extremely thin filled closed: treat as reference */
            }

            /* スナップ基準（補助）：トグル OFF 時はスキップ / Snap references (auxiliary): skip if toggle is OFF */
            if (isStrokedOpenPath && !snapOptions.strokedOnly) continue;
            if (isBlankOpenPath && !snapOptions.blank) continue;

            /* 残りはすべて基準線として追加 / Remaining items contribute as reference lines */
            addReferenceEdges(geometricBounds, tolerance, classification);
        }
        return classification;
    }

    /**
     * 1つの対象を、最寄りの基準線に合わせて変形する
     * @param {PathItem} pathItem - 動かす対象
     * @param {number[]} horizontalLines - 水平の基準線（Y 座標）
     * @param {number[]} verticalLines - 垂直の基準線（X 座標）
     * @param {Object} snapOptions - ダイアログの設定
     * @returns {boolean} 変形できたとき true
     */
    function applySnapToSingleTarget(pathItem, horizontalLines, verticalLines, snapOptions) {
        if (!isEditableItem(pathItem)) return false;
        if (snapOptions.unrotate) {
            var rotationOffsetDeg = getNearestRightAngleOffset(getEdgeAngleDeg(pathItem));
            if (Math.abs(rotationOffsetDeg) > 0.001) {
                /* 回転できないアイテムはスキップ / skip items that cannot be rotated */
                try {
                    pathItem.rotate(-rotationOffsetDeg, true, true, true, true, Transformation.CENTER);
                } catch (e) {
                    return false;
                }
            }
        }

        var maxDistance = snapOptions.maxDistance;
        var geometricBounds = pathItem.geometricBounds;
        var nearestTop = findClosestCoordinate(geometricBounds[1], horizontalLines, maxDistance);
        var nearestBottom = findClosestCoordinate(geometricBounds[3], horizontalLines, maxDistance);
        var nearestLeft = findClosestCoordinate(geometricBounds[0], verticalLines, maxDistance);
        var nearestRight = findClosestCoordinate(geometricBounds[2], verticalLines, maxDistance);

        var top = Math.max(nearestTop, nearestBottom);
        var bottom = Math.min(nearestTop, nearestBottom);
        var left = Math.min(nearestLeft, nearestRight);
        var right = Math.max(nearestLeft, nearestRight);

        if (Math.abs(right - left) <= 0 || Math.abs(top - bottom) <= 0) return false;

        return fitPathItemToTargetBounds(pathItem, left, top, right, bottom);
    }

    /**
     * すべての対象にスナップを適用する
     * @param {PageItem[]} pageItems - 処理対象候補
     * @param {Object} snapOptions - ダイアログの設定
     * @returns {{snapped: number, targets: number, horizontalLines: number, verticalLines: number}} 適用結果の件数
     */
    function applySnapToTargets(pageItems, snapOptions) {
        var classification = classifyTargetsAndReferences(pageItems, snapOptions);
        var doc = app.activeDocument;

        /* ガイドラインを基準線として合算 / Merge document guides as additional reference lines */
        if (snapOptions.includeGuides) {
            var guideLines = gatherDocumentGuides(doc, snapOptions.tolerance);
            appendAll(classification.horizontalLines, guideLines.horizontalLines);
            appendAll(classification.verticalLines, guideLines.verticalLines);
        }

        /* アートボードのエッジを基準線として合算 / Merge artboard edges as additional reference lines */
        if (snapOptions.includeArtboard) {
            var artboardEdges = gatherArtboardEdges(doc);
            appendAll(classification.horizontalLines, artboardEdges.horizontalLines);
            appendAll(classification.verticalLines, artboardEdges.verticalLines);
        }

        var snappedTargetCount = 0;
        for (var i = 0; i < classification.targets.length; i++) {
            if (applySnapToSingleTarget(classification.targets[i], classification.horizontalLines, classification.verticalLines, snapOptions)) {
                snappedTargetCount++;
            }
        }
        return {
            snapped: snappedTargetCount,
            targets: classification.targets.length,
            horizontalLines: classification.horizontalLines.length,
            verticalLines: classification.verticalLines.length
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * tooltip 付きのチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, initialValue) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * 「項目名＋数値欄＋単位」の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 項目名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {number} valuePt - 初期値（pt）
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 表示する単位
     * @returns {EditText} 数値欄
     */
    function addUnitField(parentPanel, labelPath, tooltipPath, valuePt, rulerUnit) {
        var fieldGroup = parentPanel.add("group");
        var fieldLabel = fieldGroup.add("statictext", undefined, labelText(labelPath));
        fieldLabel.preferredSize = [OPTION_LABEL_WIDTH, -1];
        var valueInput = fieldGroup.add("edittext", undefined, formatUnitValue(valuePt / rulerUnit.pointsPerUnit));
        valueInput.helpTip = getLabel(tooltipPath);
        valueInput.characters = 3;
        fieldGroup.add("statictext", undefined, rulerUnit.label);
        return valueInput;
    }

    /**
     * ダイアログを組み立てる
     * @param {Object} initialOptions - 初期値
     * @param {{label: string, pointsPerUnit: number}} rulerUnit - 数値欄の単位
     * @returns {Object} ダイアログ（snapDialog）と各コントロール
     */
    function buildSnapDialog(initialOptions, rulerUnit) {
        var snapDialog = new Window("dialog", getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        snapDialog.orientation = "column";
        snapDialog.alignChildren = "fill";
        snapDialog.margins = DIALOG_MARGINS;

        /* 「変形するもの」パネル：動かされる側 / Targets panel: items being moved */
        var targetPanel = snapDialog.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, 6);
        var filledCheckbox = addOptionCheckbox(targetPanel, "checkbox.filled", "tooltip.filled", initialOptions.filled);
        var groupCheckbox = addOptionCheckbox(targetPanel, "checkbox.group", "tooltip.group", initialOptions.group);
        var clipGroupCheckbox = addOptionCheckbox(targetPanel, "checkbox.clipGroup", "tooltip.clipGroup", initialOptions.clipGroup);

        /* 「スナップ基準」パネル：基準として扱うパス・追加基準 / Snap references panel: paths and extra references */
        var snapBasisPanel = snapDialog.add("panel", undefined, getLabel("panel.snapBasis"));
        setupPanel(snapBasisPanel, 6);
        var strokedOnlyCheckbox = addOptionCheckbox(snapBasisPanel, "checkbox.strokedOnly", "tooltip.strokedOnly", initialOptions.strokedOnly);
        var blankCheckbox = addOptionCheckbox(snapBasisPanel, "checkbox.blank", "tooltip.blank", initialOptions.blank);
        var includeGuidesCheckbox = addOptionCheckbox(snapBasisPanel, "checkbox.includeGuides", "tooltip.includeGuides", initialOptions.includeGuides);
        var includeArtboardCheckbox = addOptionCheckbox(snapBasisPanel, "checkbox.includeArtboard", "tooltip.includeArtboard", initialOptions.includeArtboard);

        /* 「オプション」パネル / Options panel */
        var optionPanel = snapDialog.add("panel", undefined, getLabel("panel.option"));
        setupPanel(optionPanel, 6);
        var unrotateCheckbox = addOptionCheckbox(optionPanel, "checkbox.unrotate", "tooltip.unrotate", initialOptions.unrotate);
        var toleranceInput = addUnitField(optionPanel, "fieldLabel.tolerance", "tooltip.tolerance", initialOptions.tolerance, rulerUnit);
        var maxDistanceInput = addUnitField(optionPanel, "fieldLabel.maxDistance", "tooltip.maxDistance", initialOptions.maxDistance, rulerUnit);

        /* 下段：左＝プレビュー、中央＝余白、右＝ボタン / Bottom row: left=preview, center=spacer, right=buttons */
        var btnRowGroup = snapDialog.add("group");
        btnRowGroup.alignment = "fill";
        btnRowGroup.alignChildren = ["fill", "center"];

        var previewGroup = btnRowGroup.add("group");
        previewGroup.alignment = ["left", "center"];
        var previewCheckbox = addOptionCheckbox(previewGroup, "checkbox.preview", "tooltip.preview", initialOptions.preview);

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignment = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            snapDialog: snapDialog,
            filledCheckbox: filledCheckbox,
            groupCheckbox: groupCheckbox,
            clipGroupCheckbox: clipGroupCheckbox,
            strokedOnlyCheckbox: strokedOnlyCheckbox,
            blankCheckbox: blankCheckbox,
            includeGuidesCheckbox: includeGuidesCheckbox,
            includeArtboardCheckbox: includeArtboardCheckbox,
            unrotateCheckbox: unrotateCheckbox,
            toleranceInput: toleranceInput,
            maxDistanceInput: maxDistanceInput,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * ダイアログの現在の入力値を読み取る（数値は pt に換算）
     * @param {Object} dialogControls - buildSnapDialog() の戻り値
     * @param {{pointsPerUnit: number}} rulerUnit - 数値欄の単位
     * @returns {Object} スナップの設定
     */
    function readSnapOptions(dialogControls, rulerUnit) {
        var parsedTolerance = parseFloat(dialogControls.toleranceInput.text) * rulerUnit.pointsPerUnit;
        if (isNaN(parsedTolerance) || parsedTolerance <= 0) parsedTolerance = 0.5;
        var parsedMaxDistance = parseFloat(dialogControls.maxDistanceInput.text) * rulerUnit.pointsPerUnit;
        if (isNaN(parsedMaxDistance) || parsedMaxDistance < 0) parsedMaxDistance = 0;
        return {
            filled: dialogControls.filledCheckbox.value,
            strokedOnly: dialogControls.strokedOnlyCheckbox.value,
            blank: dialogControls.blankCheckbox.value,
            group: dialogControls.groupCheckbox.value,
            clipGroup: dialogControls.clipGroupCheckbox.value,
            unrotate: dialogControls.unrotateCheckbox.value,
            tolerance: parsedTolerance,
            maxDistance: parsedMaxDistance,
            includeGuides: dialogControls.includeGuidesCheckbox.value,
            includeArtboard: dialogControls.includeArtboardCheckbox.value,
            preview: dialogControls.previewCheckbox.value
        };
    }

    /**
     * ダイアログを表示し、確定した設定を返す
     * @param {Object} initialOptions - 初期値
     * @param {Function} onPreview - 設定が変わるたびに呼ぶ関数（引数は設定）
     * @returns {Object|null} 確定した設定。キャンセルなら null
     */
    function showSnapDialog(initialOptions, onPreview) {
        var rulerUnit = getUnitInfo();
        var dialogControls = buildSnapDialog(initialOptions, rulerUnit);

        /**
         * 現在の設定でプレビューを呼ぶ
         * @returns {void}
         */
        function notifyPreview() {
            onPreview(readSnapOptions(dialogControls, rulerUnit));
        }

        dialogControls.filledCheckbox.onClick = notifyPreview;
        dialogControls.strokedOnlyCheckbox.onClick = notifyPreview;
        dialogControls.blankCheckbox.onClick = notifyPreview;
        dialogControls.groupCheckbox.onClick = notifyPreview;
        dialogControls.clipGroupCheckbox.onClick = notifyPreview;
        dialogControls.unrotateCheckbox.onClick = notifyPreview;
        dialogControls.includeGuidesCheckbox.onClick = notifyPreview;
        dialogControls.includeArtboardCheckbox.onClick = notifyPreview;
        dialogControls.previewCheckbox.onClick = notifyPreview;
        dialogControls.toleranceInput.onChange = notifyPreview;
        dialogControls.maxDistanceInput.onChange = notifyPreview;

        if (dialogControls.snapDialog.show() !== 1) return null;
        return readSnapOptions(dialogControls, rulerUnit);
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * プレビュー付きのダイアログで設定を決め、選択中の対象をスナップする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var selectedPageItems = app.activeDocument.selection;
        if (selectedPageItems.length < 1) {
            alert(getLabel("alert.selectTarget"));
            return;
        }

        /* スナップショットは（プレビュー切替で復元できるように）通常グループもクリップグループも展開して取得 */
        /* Snapshot covers everything — both normal and clip-group descendants — so preview toggles can reset cleanly */
        var allProcessableItemsForPreviewReset = collectProcessableItems(selectedPageItems, { includeNormalGroups: true, includeClipGroups: true });
        var originalGeometrySnapshot = captureOriginalGeometry(allProcessableItemsForPreviewReset);
        var shouldRestoreOriginalGeometry = true;

        /**
         * 現在の設定に応じて処理対象候補を集める
         * @param {Object} snapOptions - スナップの設定
         * @returns {PageItem[]} 処理対象候補
         */
        function collectItemsForOptions(snapOptions) {
            return collectProcessableItems(selectedPageItems, {
                includeNormalGroups: snapOptions.group === true,
                includeClipGroups: snapOptions.clipGroup === true
            });
        }

        /**
         * 元の形状に戻してから、プレビューが ON ならスナップを適用する
         * @param {Object} snapOptions - スナップの設定
         * @returns {void}
         */
        function applyPreview(snapOptions) {
            restoreOriginalGeometry(originalGeometrySnapshot);
            if (snapOptions.preview) {
                applySnapToTargets(collectItemsForOptions(snapOptions), snapOptions);
            }
            app.redraw();
        }

        /* 途中で抜けたとき（キャンセル・エラー）は元の形状に戻す / restore on cancel or error */
        try {
            /* 初回プレビュー / Initial preview */
            applyPreview(DEFAULT_SNAP_OPTIONS);

            var confirmedOptions = showSnapDialog(DEFAULT_SNAP_OPTIONS, applyPreview);
            if (!confirmedOptions) {
                /* キャンセル：状態を巻き戻し / Cancel: roll back to original state */
                return;
            }

            /* OK：スナップショットから本適用 / OK: re-apply cleanly from snapshot */
            restoreOriginalGeometry(originalGeometrySnapshot);
            var snapApplyResult = applySnapToTargets(collectItemsForOptions(confirmedOptions), confirmedOptions);

            if (snapApplyResult.targets === 0) {
                alert(getLabel("alert.noTarget"));
                return;
            }
            if (snapApplyResult.horizontalLines === 0 && snapApplyResult.verticalLines === 0) {
                alert(getLabel("alert.noReferenceLines"));
                return;
            }

            shouldRestoreOriginalGeometry = false;
        } finally {
            if (shouldRestoreOriginalGeometry) {
                restoreOriginalGeometry(originalGeometrySnapshot);
                app.redraw();
            }
        }
    }

    main();

})();
