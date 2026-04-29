#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要
選択中のオブジェクトを「動かす対象」と「スナップ基準」に分類し、対象の
バウンディングボックスを最寄りの基準線へ合わせるスクリプト。パスは
アンカーポイントを直接変形するため、クリップグループ内の子パスにも対応する。

「線判定の許容差」は、細長いパスを水平線／垂直線として扱うための判定幅。
「最大スナップ距離」は、対象の辺からどれだけ離れた基準線まで吸着するかの制限。
0の場合は距離制限なし。

Overview
Classifies the current selection into items to move and snap references, then
fits each target's bounding box to the nearest references. Path items are fitted
by directly transforming anchor points, so child paths inside clipping groups are supported.

"Line detection tolerance" controls how thin a path must be to count as a
horizontal or vertical reference line. "Max snap distance" limits how far a
target edge may move to the nearest reference; 0 means no distance limit.
*/

// =========================================
// バージョンとローカライズ / Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.0.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    ui: {
        dialogTitle: { ja: "塗りを線にスナップ", en: "Snap Fill to Lines" },
        panelTarget: { ja: "動かす対象", en: "Items to move" },
        cbFilled: { ja: "塗りのあるクローズパス", en: "Filled closed paths" },
        panelSnapBasis: { ja: "スナップ基準", en: "Snap references" },
        cbStrokedOnly: { ja: "線だけのパス", en: "Stroke-only paths" },
        cbBlank: { ja: "塗り／線のないオープンパス", en: "Open paths without fill or stroke" },
        cbGroup: { ja: "グループ内のパスも対象にする", en: "Include paths inside groups" },
        cbClipGroup: { ja: "クリップグループ内のパスも対象にする", en: "Include paths inside clipping groups" },
        panelOption: { ja: "オプション", en: "Options" },
        cbUnrotate: { ja: "回転補正", en: "Rotation correction" },
        tolerance: { ja: "線判定の許容差", en: "Line detection tolerance" },
        maxDistance: { ja: "最大スナップ距離", en: "Max snap distance" },
        cbIncludeGuides: { ja: "ガイドライン", en: "Guide lines" },
        cbIncludeArtboard: { ja: "アートボードのエッジ", en: "Artboard edges" },
        cbPreview: { ja: "プレビュー", en: "Preview" },
        btnOK: { ja: "OK", en: "OK" },
        btnCancel: { ja: "キャンセル", en: "Cancel" }
    },
    error: {
        errNoDoc: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        errSelect: { ja: "スナップ対象のオブジェクトを選択してください。", en: "Select a snap target object." }, errNoTarget: { ja: "変形対象が見つかりません。", en: "No transformable target found." },
        errNoLines: { ja: "基準となる罫線や枠（塗りなしパス）が見つかりません。", en: "No reference lines or frames (unfilled paths) found." }
    }
};

/* ラベル取得 / Get localized label */
function getLabel(key) {
    var categories = [LABELS.ui, LABELS.error];
    for (var i = 0; i < categories.length; i++) {
        var entry = categories[i][key];
        if (entry) return entry[lang] || entry.en || key;
    }
    return key;
}
var L = getLabel;

/* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
function labelText(key) {
    return getLabel(key) + (lang === 'ja' ? '：' : ':');
}


// =========================================
// 設定 / Configuration
// =========================================

var PANEL_MARGINS = [15, 20, 15, 10];
var OPTION_LABEL_WIDTH = 120;

var unitLabelMap = {
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

var pointsPerUnitMap = {
    0: 72, // in
    1: 72 / 25.4, // mm
    2: 1, // pt
    3: 12, // pica
    4: 72 / 2.54, // cm
    5: 2.834645669291339 / 4, // Q/H
    6: 1, // px: Illustrator generally treats 1 px as 1 pt
    7: 864, // ft/in
    8: 72 / 0.0254, // m
    9: 2592, // yd
    10: 864 // ft
};

/* 現在の定規単位コードを取得 / Get current ruler unit code */
function getCurrentRulerUnitCode() {
    try {
        return app.preferences.getIntegerPreference("rulerType");
    } catch (e) {
        return 2;
    }
}

/* 単位コードからラベルを取得 / Get ruler unit label from unit code */
function getUnitLabel(unitCode) {
    return unitLabelMap[unitCode] || "pt";
}

/* 単位コードから pt 換算係数を取得 / Get points-per-unit factor from unit code */
function getPointsPerUnit(unitCode) {
    return pointsPerUnitMap[unitCode] || 1;
}

/* pt を指定単位へ変換 / Convert points to the specified ruler unit */
function pointsToUnit(valuePt, unitCode) {
    return valuePt / getPointsPerUnit(unitCode);
}

/* 指定単位を pt へ変換 / Convert the specified ruler unit to points */
function unitToPoints(value, unitCode) {
    return value * getPointsPerUnit(unitCode);
}

/* 入力欄用に数値を整形 / Format numeric value for edit fields */
function formatUnitValue(value) {
    var rounded = Math.round(value * 1000) / 1000;
    return String(rounded);
}


// =========================================
// 共通ユーティリティ / Common utilities
// =========================================

/* パネルの共通設定 / Apply common panel layout */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = "left";
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    if (typeof spacing === "number") {
        panel.spacing = spacing;
    }
}

/* オプションに応じて処理対象候補を再帰収集 / Recursively collect processable item candidates by options */
function collectProcessableItems(pageItems, options) {
    if (!options) options = {};
    var includeNormalGroups = (options.includeNormalGroups !== false);
    var includeClipGroups = (options.includeClipGroups === true);
    var processableItems = [];
    for (var i = 0; i < pageItems.length; i++) {
        var pageItem = pageItems[i];
        if (pageItem.typename === "GroupItem") {
            if (pageItem.clipped === true) {
                /* クリップグループ：マスクパスは基準線化しやすいため除外し、中身だけを再帰 / Clip group: skip the clipping mask path and recurse into contents */
                if (!includeClipGroups) continue;
                var clippingGroupItems = pageItem.pageItems;
                for (var j = 0; j < clippingGroupItems.length; j++) {
                    var childItem = clippingGroupItems[j];
                    if (childItem.typename === "PathItem" && childItem.clipping === true) continue;
                    var nestedProcessableItems = collectProcessableItems([childItem], options);
                    for (var k = 0; k < nestedProcessableItems.length; k++) processableItems.push(nestedProcessableItems[k]);
                }
            } else {
                if (!includeNormalGroups) continue;
                var nestedProcessableItems = collectProcessableItems(pageItem.pageItems, options);
                for (var m = 0; m < nestedProcessableItems.length; m++) processableItems.push(nestedProcessableItems[m]);
            }
        } else if (pageItem.typename === "CompoundPathItem") {
            var compoundPathItems = collectProcessableItems(pageItem.pathItems, options);
            for (var n = 0; n < compoundPathItems.length; n++) processableItems.push(compoundPathItems[n]);
        } else {
            processableItems.push(pageItem);
        }
    }
    return processableItems;
}

/* アイテムが編集可能かを安全に判定 / Safely check whether an item is editable */
function isEditableItem(pageItem) {
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

/* プレビュー復元用に元のパス形状を保存 / Capture original path geometry for preview restore */
function captureOriginalGeometry(pageItems) {
    var snapshotData = [];
    for (var i = 0; i < pageItems.length; i++) {
        var pageItem = pageItems[i];
        if (pageItem.typename !== "PathItem") continue;
        if (!isEditableItem(pageItem)) continue;

        var pathPointList;
        try {
            pathPointList = pageItem.pathPoints;
        } catch (e) {
            continue;
        }

        var savedPoints = [];
        for (var j = 0; j < pathPointList.length; j++) {
            var pathPoint = pathPointList[j];
            try {
                savedPoints.push({
                    anchor: [pathPoint.anchor[0], pathPoint.anchor[1]],
                    leftDirection: [pathPoint.leftDirection[0], pathPoint.leftDirection[1]],
                    rightDirection: [pathPoint.rightDirection[0], pathPoint.rightDirection[1]]
                });
            } catch (e2) {
                savedPoints = [];
                break;
            }
        }
        if (savedPoints.length > 0) {
            snapshotData.push({ item: pageItem, points: savedPoints });
        }
    }
    return snapshotData;
}

/* 保存した元のパス形状へ戻す / Restore saved original path geometry */
function restoreOriginalGeometry(snapshotData) {
    for (var i = 0; i < snapshotData.length; i++) {
        var snapshotEntry = snapshotData[i];
        if (!snapshotEntry || !snapshotEntry.item || !isEditableItem(snapshotEntry.item)) continue;

        var pathPointList;
        try {
            pathPointList = snapshotEntry.item.pathPoints;
        } catch (e) {
            continue;
        }

        var limit = Math.min(snapshotEntry.points.length, pathPointList.length);
        for (var j = 0; j < limit; j++) {
            var savedPoint = snapshotEntry.points[j];
            try {
                pathPointList[j].anchor = savedPoint.anchor;
                pathPointList[j].leftDirection = savedPoint.leftDirection;
                pathPointList[j].rightDirection = savedPoint.rightDirection;
            } catch (e2) {
                break;
            }
        }
    }
}

/* 候補値の中から最も近い値を返す（最大距離付き）/ Return the closest value within max distance */
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

/* パスのアンカーポイントを指定範囲に直接フィット / Fit path points directly to target bounds */
function fitPathItemToTargetBounds(pathItem, left, top, right, bottom) {
    if (!pathItem || pathItem.typename !== "PathItem") return false;
    if (!isEditableItem(pathItem)) return false;

    var geometricBounds = pathItem.geometricBounds; /* [left, top, right, bottom] */
    var sourceLeft = geometricBounds[0];
    var sourceTop = geometricBounds[1];
    var sourceRight = geometricBounds[2];
    var sourceBottom = geometricBounds[3];
    var sourceWidth = sourceRight - sourceLeft;
    var sourceHeight = sourceTop - sourceBottom;
    var targetWidthPt = right - left;
    var targetHeightPt = top - bottom;

    if (Math.abs(sourceWidth) <= 0.0001 || Math.abs(sourceHeight) <= 0.0001) return false;
    if (Math.abs(targetWidthPt) <= 0.0001 || Math.abs(targetHeightPt) <= 0.0001) return false;

    var scaleX = targetWidthPt / sourceWidth;
    var scaleY = targetHeightPt / sourceHeight;
    var pathPointList = pathItem.pathPoints;

    function mapPoint(pointArray) {
        return [
            left + (pointArray[0] - sourceLeft) * scaleX,
            top - (sourceTop - pointArray[1]) * scaleY
        ];
    }

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

/* アクティブなアートボードの 4 辺を収集 / Gather edges of the active artboard */
function gatherArtboardEdges(doc) {
    if (!doc) return { horizontalLines: [], verticalLines: [] };
    var index;
    try { index = doc.artboards.getActiveArtboardIndex(); } catch (e) { index = 0; }
    var rect = doc.artboards[index].artboardRect; /* [left, top, right, bottom] */
    return {
        horizontalLines: [rect[1], rect[3]],
        verticalLines: [rect[0], rect[2]]
    };
}

/* ドキュメント内のガイドラインから水平線・垂直線を収集 / Gather guide lines from the document */
function gatherDocumentGuides(doc, tolerance) {
    if (tolerance == null) tolerance = 0.5;
    var horizontalLines = [];
    var verticalLines = [];
    var documentPathItems = doc.pathItems;
    for (var i = 0; i < documentPathItems.length; i++) {
        var guidePathItem = documentPathItems[i];
        if (guidePathItem.guides !== true) continue;
        var geometricBounds = guidePathItem.geometricBounds;
        var guideWidthPt = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var guideHeightPt = Math.abs(geometricBounds[1] - geometricBounds[3]);
        if (guideWidthPt < tolerance) {
            verticalLines.push(geometricBounds[0]);
        } else if (guideHeightPt < tolerance) {
            horizontalLines.push(geometricBounds[1]);
        } else {
            /* 矩形ガイドなどは4辺すべてを採用 / Rectangular guide: use all four sides */
            horizontalLines.push(geometricBounds[1]);
            horizontalLines.push(geometricBounds[3]);
            verticalLines.push(geometricBounds[0]);
            verticalLines.push(geometricBounds[2]);
        }
    }
    return { horizontalLines: horizontalLines, verticalLines: verticalLines };
}


// =========================================
// 角度・回転補正 / Angle & Rotation correction
// =========================================

/* 最初のエッジの角度（度）/ Angle of the first edge in degrees */
function getEdgeAngleDeg(item) {
    var pathPoints = item.pathPoints;
    if (!pathPoints || pathPoints.length < 2) return 0;
    var firstAnchor = pathPoints[0].anchor;
    var secondAnchor = pathPoints[1].anchor;
    return Math.atan2(secondAnchor[1] - firstAnchor[1], secondAnchor[0] - firstAnchor[0]) * 180 / Math.PI;
}

/* 最寄りの 90 度倍数からのずれ / Offset from nearest right angle */
function getNearestRightAngleOffset(angleDeg) {
    var modulus = angleDeg % 90;
    if (modulus > 45) modulus -= 90;
    if (modulus < -45) modulus += 90;
    return modulus;
}


// =========================================
// スナップ判定・実行 / Snap classification & execution
// =========================================

/* 処理対象候補をスナップ対象と基準線に分類 / Classify processable candidates into targets and reference lines */
function classifyTargetsAndReferences(pageItems, tolerance, options) {
    if (tolerance == null) tolerance = 0.5;
    if (!options) options = { filled: true, strokedOnly: true, blank: true };
    var classification = { targets: [], horizontalLines: [], verticalLines: [] };

    for (var i = 0; i < pageItems.length; i++) {
        var pageItem = pageItems[i];
        if (pageItem.typename === "TextFrame") continue;
        if (pageItem.typename !== "PathItem") continue;

        var geometricBounds = pageItem.geometricBounds; /* [left, top, right, bottom] */
        var itemWidthPt = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var itemHeightPt = Math.abs(geometricBounds[1] - geometricBounds[3]);

        var isFilledClosedPath = pageItem.filled && pageItem.closed;
        var isStrokedOpenPath = !pageItem.filled && pageItem.stroked && !pageItem.closed;
        var isBlankOpenPath = !pageItem.filled && !pageItem.stroked && !pageItem.closed;

        /* 変形対象：塗りのあるクローズパスのみ / Targets: filled closed paths only */
        if (isFilledClosedPath) {
            if (!options.filled) continue;
            if (itemWidthPt > tolerance && itemHeightPt > tolerance) {
                classification.targets.push(pageItem);
                continue;
            }
            /* 極端に細い filled closed は基準として扱う / Extremely thin filled closed: treat as reference */
        }

        /* スナップ基準（補助）：トグル OFF 時はスキップ / Snap references (auxiliary): skip if toggle is OFF */
        if (isStrokedOpenPath && !options.strokedOnly) continue;
        if (isBlankOpenPath && !options.blank) continue;

        /* 残りはすべて基準線として追加 / Remaining items contribute as reference lines */
        if (itemWidthPt < tolerance) {
            classification.verticalLines.push(geometricBounds[0]);
        } else if (itemHeightPt < tolerance) {
            classification.horizontalLines.push(geometricBounds[1]);
        } else {
            classification.horizontalLines.push(geometricBounds[1]);
            classification.horizontalLines.push(geometricBounds[3]);
            classification.verticalLines.push(geometricBounds[0]);
            classification.verticalLines.push(geometricBounds[2]);
        }
    }
    return classification;
}

/* 1つのスナップ対象を最寄りの基準線に合わせる / Apply snap to one target using nearest reference lines */
function applySnapToSingleTarget(pageItem, horizontalLines, verticalLines, options) {
    if (!isEditableItem(pageItem)) return false;
    if (options && options.unrotate) {
        var rotationOffsetDeg = getNearestRightAngleOffset(getEdgeAngleDeg(pageItem));
        if (Math.abs(rotationOffsetDeg) > 0.001) {
            try {
                pageItem.rotate(-rotationOffsetDeg, true, true, true, true, Transformation.CENTER);
            } catch (e) {
                return false;
            }
        }
    }

    var maxDistance = options ? options.maxDistance : 0;
    var geometricBounds = pageItem.geometricBounds;
    var nearestTop = findClosestCoordinate(geometricBounds[1], horizontalLines, maxDistance);
    var nearestBottom = findClosestCoordinate(geometricBounds[3], horizontalLines, maxDistance);
    var nearestLeft = findClosestCoordinate(geometricBounds[0], verticalLines, maxDistance);
    var nearestRight = findClosestCoordinate(geometricBounds[2], verticalLines, maxDistance);

    var top = Math.max(nearestTop, nearestBottom);
    var bottom = Math.min(nearestTop, nearestBottom);
    var left = Math.min(nearestLeft, nearestRight);
    var right = Math.max(nearestLeft, nearestRight);

    var targetWidthPt = Math.abs(right - left);
    var targetHeightPt = Math.abs(top - bottom);
    if (targetWidthPt <= 0 || targetHeightPt <= 0) return false;

    if (pageItem.typename !== "PathItem") return false;
    return fitPathItemToTargetBounds(pageItem, left, top, right, bottom);
}

/* すべてのスナップ対象へ適用 / Apply snap to all targets */
function applySnapToTargets(pageItems, tolerance, options) {
    var classification = classifyTargetsAndReferences(pageItems, tolerance, options);

    /* ガイドラインを基準線として合算 / Merge document guides as additional reference lines */
    if (options && options.includeGuides && app.documents.length > 0) {
        var guideLines = gatherDocumentGuides(app.activeDocument, tolerance);
        for (var g = 0; g < guideLines.horizontalLines.length; g++) classification.horizontalLines.push(guideLines.horizontalLines[g]);
        for (var v = 0; v < guideLines.verticalLines.length; v++) classification.verticalLines.push(guideLines.verticalLines[v]);
    }

    /* アートボードのエッジを基準線として合算 / Merge artboard edges as additional reference lines */
    if (options && options.includeArtboard && app.documents.length > 0) {
        var artboardEdges = gatherArtboardEdges(app.activeDocument);
        for (var ah = 0; ah < artboardEdges.horizontalLines.length; ah++) classification.horizontalLines.push(artboardEdges.horizontalLines[ah]);
        for (var av = 0; av < artboardEdges.verticalLines.length; av++) classification.verticalLines.push(artboardEdges.verticalLines[av]);
    }

    var snappedTargetCount = 0;
    for (var i = 0; i < classification.targets.length; i++) {
        if (applySnapToSingleTarget(classification.targets[i], classification.horizontalLines, classification.verticalLines, options)) {
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

function showSnapDialog(defaults, onPreview) {
    var dialog = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
    dialog.orientation = "column";
    dialog.alignChildren = "fill";
    dialog.margins = 16;
    var currentUnitCode = getCurrentRulerUnitCode();
    var currentUnitLabel = getUnitLabel(currentUnitCode);

    /* 「変形するもの」パネル：動かされる側 / Targets panel: items being moved */
    var targetPanel = dialog.add("panel", undefined, L('panelTarget'));
    setupPanel(targetPanel, 6);
    var checkboxFilled = targetPanel.add("checkbox", undefined, L('cbFilled'));
    checkboxFilled.value = defaults.filled;
    var checkboxGroup = targetPanel.add("checkbox", undefined, L('cbGroup'));
    checkboxGroup.value = defaults.group;
    var checkboxClipGroup = targetPanel.add("checkbox", undefined, L('cbClipGroup'));
    checkboxClipGroup.value = defaults.clipGroup;

    /* 「スナップ基準」パネル：基準として扱うパス・追加基準 / Snap references panel: paths and extra references */
    var basisPanel = dialog.add("panel", undefined, L('panelSnapBasis'));
    setupPanel(basisPanel, 6);
    var checkboxStrokedOnly = basisPanel.add("checkbox", undefined, L('cbStrokedOnly'));
    checkboxStrokedOnly.value = defaults.strokedOnly;
    var checkboxBlank = basisPanel.add("checkbox", undefined, L('cbBlank'));
    checkboxBlank.value = defaults.blank;
    var checkboxIncludeGuides = basisPanel.add("checkbox", undefined, L('cbIncludeGuides'));
    checkboxIncludeGuides.value = defaults.includeGuides;
    var checkboxIncludeArtboard = basisPanel.add("checkbox", undefined, L('cbIncludeArtboard'));
    checkboxIncludeArtboard.value = defaults.includeArtboard;

    /* 「オプション」パネル / Options panel */
    var optionPanel = dialog.add("panel", undefined, L('panelOption'));
    setupPanel(optionPanel, 6);
    var checkboxUnrotate = optionPanel.add("checkbox", undefined, L('cbUnrotate'));
    checkboxUnrotate.value = defaults.unrotate;

    var toleranceGroup = optionPanel.add("group");
    var toleranceLabel = toleranceGroup.add("statictext", undefined, labelText('tolerance'));
    toleranceLabel.preferredSize = [OPTION_LABEL_WIDTH, -1];
    var toleranceInput = toleranceGroup.add("edittext", undefined, formatUnitValue(pointsToUnit(defaults.tolerance, currentUnitCode)));
    toleranceInput.characters = 3;
    toleranceGroup.add("statictext", undefined, currentUnitLabel);

    var maxDistanceGroup = optionPanel.add("group");
    var maxDistanceLabel = maxDistanceGroup.add("statictext", undefined, labelText('maxDistance'));
    maxDistanceLabel.preferredSize = [OPTION_LABEL_WIDTH, -1];
    var maxDistanceInput = maxDistanceGroup.add("edittext", undefined, formatUnitValue(pointsToUnit(defaults.maxDistance, currentUnitCode))); maxDistanceInput.characters = 3;
    maxDistanceGroup.add("statictext", undefined, currentUnitLabel);

    /* 下段：左＝プレビュー、中央＝余白、右＝ボタン / Bottom row: left=preview, center=spacer, right=buttons */
    var bottomGroup = dialog.add("group");
    bottomGroup.alignment = "fill";
    bottomGroup.alignChildren = ["fill", "center"];

    var previewGroup = bottomGroup.add("group");
    previewGroup.alignment = ["left", "center"];
    var checkboxPreview = previewGroup.add("checkbox", undefined, L('cbPreview'));
    checkboxPreview.value = defaults.preview;

    var spacer = bottomGroup.add("group");
    spacer.alignment = ["fill", "fill"];

    var buttonGroup = bottomGroup.add("group");
    buttonGroup.alignment = ["right", "center"];
    buttonGroup.add("button", undefined, L('btnCancel'), { name: "cancel" });
    buttonGroup.add("button", undefined, L('btnOK'), { name: "ok" });

    /* 現在の入力値を取得 / Collect current input values */
    function collectOptions() {
        var parsedTolerance = unitToPoints(parseFloat(toleranceInput.text), currentUnitCode);
        if (isNaN(parsedTolerance) || parsedTolerance <= 0) parsedTolerance = 0.5;
        var parsedMaxDistance = unitToPoints(parseFloat(maxDistanceInput.text), currentUnitCode);
        if (isNaN(parsedMaxDistance) || parsedMaxDistance < 0) parsedMaxDistance = 0;
        return {
            filled: checkboxFilled.value,
            strokedOnly: checkboxStrokedOnly.value,
            blank: checkboxBlank.value,
            group: checkboxGroup.value,
            clipGroup: checkboxClipGroup.value,
            unrotate: checkboxUnrotate.value,
            tolerance: parsedTolerance,
            maxDistance: parsedMaxDistance,
            includeGuides: checkboxIncludeGuides.value,
            includeArtboard: checkboxIncludeArtboard.value,
            preview: checkboxPreview.value
        };
    }

    /* プレビューコールバックを発火 / Fire preview callback */
    function notifyPreview() {
        if (typeof onPreview === "function") {
            onPreview(collectOptions());
        }
    }

    checkboxFilled.onClick = notifyPreview;
    checkboxStrokedOnly.onClick = notifyPreview;
    checkboxBlank.onClick = notifyPreview;
    checkboxGroup.onClick = notifyPreview;
    checkboxClipGroup.onClick = notifyPreview;
    checkboxUnrotate.onClick = notifyPreview;
    checkboxIncludeGuides.onClick = notifyPreview;
    checkboxIncludeArtboard.onClick = notifyPreview;
    checkboxPreview.onClick = notifyPreview;
    toleranceInput.onChange = notifyPreview;
    maxDistanceInput.onChange = notifyPreview;

    if (dialog.show() !== 1) return null;
    return collectOptions();
}


// =========================================
// メイン / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(L('errNoDoc'));
        return;
    }
    var selectedPageItems = app.activeDocument.selection;
    if (selectedPageItems.length < 1) {
        alert(L('errSelect'));
        return;
    }

    /* スナップショットは（プレビュー切替で復元できるように）通常グループもクリップグループも展開して取得 */
    /* Snapshot covers everything — both normal and clip-group descendants — so preview toggles can reset cleanly */
    var allProcessableItemsForPreviewReset = collectProcessableItems(selectedPageItems, { includeNormalGroups: true, includeClipGroups: true });
    var originalGeometrySnapshot = captureOriginalGeometry(allProcessableItemsForPreviewReset);
    var shouldRestoreOriginalGeometry = true;

    try {
        /* 現在のオプションに応じて処理対象を取得 / Resolve items to process under current options */
        function collectItemsForCurrentOptions(currentOptions) {
            return collectProcessableItems(selectedPageItems, {
                includeNormalGroups: currentOptions.group === true,
                includeClipGroups: currentOptions.clipGroup === true
            });
        }

        /* プレビュー反映 / Apply preview */
        function applyPreview(currentOptions) {
            restoreOriginalGeometry(originalGeometrySnapshot);
            if (currentOptions.preview) {
                applySnapToTargets(collectItemsForCurrentOptions(currentOptions), currentOptions.tolerance, currentOptions);
            }
            app.redraw();
        }

        var defaultDialogOptions = {
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

        /* 初回プレビュー / Initial preview */
        applyPreview(defaultDialogOptions);

        var confirmedDialogOptions = showSnapDialog(defaultDialogOptions, applyPreview);
        if (!confirmedDialogOptions) {
            /* キャンセル：状態を巻き戻し / Cancel: roll back to original state */
            return;
        }

        /* OK：スナップショットから本適用 / OK: re-apply cleanly from snapshot */
        restoreOriginalGeometry(originalGeometrySnapshot);
        var snapApplyResult = applySnapToTargets(collectItemsForCurrentOptions(confirmedDialogOptions), confirmedDialogOptions.tolerance, confirmedDialogOptions);

        if (snapApplyResult.targets === 0) {
            alert(L('errNoTarget'));
            return;
        }
        if (snapApplyResult.horizontalLines === 0 && snapApplyResult.verticalLines === 0) {
            alert(L('errNoLines'));
            return;
        }

        shouldRestoreOriginalGeometry = false;
    } finally {
        if (shouldRestoreOriginalGeometry) {
            restoreOriginalGeometry(originalGeometrySnapshot);
            app.redraw();
        }
    }
})();
