#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
概要
選択した外枠の長方形と、その内側にある縦罫・横罫を、外枠の中で
均等に再配置します。横罫・縦罫はそれぞれ ON/OFF でき、プレビュー
切り替えにも対応します。

Overview
Evenly redistributes the vertical and horizontal rules inside the selected
outer rectangle. Horizontal and vertical rules can be toggled
independently, with a live preview switch.
*/


// =========================================
// バージョンとローカライズ / Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: {
        ja: "罫線を平均化",
        en: "Distribute Rules"
    },
    panelTarget: {
        ja: "対象",
        en: "Target"
    },
    averageHorizontal: {
        ja: "横罫を平均化",
        en: "Distribute horizontal rules"
    },
    averageVertical: {
        ja: "縦罫を平均化",
        en: "Distribute vertical rules"
    },
    matchRuleLengths: {
        ja: "長さを揃える",
        en: "Match rule lengths"
    },
    preview: {
        ja: "プレビュー",
        en: "Preview"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    errNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    errNoSelection: {
        ja: "外枠の長方形と罫線を選択してください。",
        en: "Please select the outer rectangle and the rules."
    },
    errNoPathInSelection: {
        ja: "選択内にパスがありません。",
        en: "No paths found in the selection."
    },
    errOuterRectNotFound: {
        ja: "外枠の長方形が見つかりませんでした。",
        en: "Outer rectangle not found."
    },
    errNoRulesFound: {
        ja: "縦罫または横罫が見つかりませんでした。",
        en: "No vertical or horizontal rules were found."
    }
};

/* ローカライズされたラベル取得 / Get a localized label */
function L(key) {
    var entry = LABELS[key];
    if (!entry) {
        return key;
    }
    return entry[lang] || entry.en;
}



// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(L("errNoDocument"));
        return;
    }

    var doc = app.activeDocument;

    if (doc.selection.length === 0) {
        alert(L("errNoSelection"));
        return;
    }

    /*
        true  : 縦罫の上下、横罫の左右を外枠に合わせる
                Snap rule endpoints to the outer rectangle.
        false : 罫線の長さは変えず、位置だけ平均化する
                Keep rule lengths; redistribute positions only.
    */
    var MATCH_LINES_TO_RECT = true;

    var TOLERANCE = 0.5;
    var PANEL_MARGINS = [15, 20, 15, 10];

    var pathItems = [];
    collectPathItems(doc.selection, pathItems);

    if (pathItems.length === 0) {
        alert(L("errNoPathInSelection"));
        return;
    }

    var outerRect = findOuterRectangle(pathItems);

    if (!outerRect) {
        alert(L("errOuterRectNotFound"));
        return;
    }

    var outerBounds = outerRect.geometricBounds;
    var rectLeft = outerBounds[0];
    var rectTop = outerBounds[1];
    var rectRight = outerBounds[2];
    var rectBottom = outerBounds[3];

    var rectWidth = rectRight - rectLeft;
    var rectHeight = rectTop - rectBottom;

    var verticalLines = [];
    var horizontalLines = [];

    for (var pathIndex = 0; pathIndex < pathItems.length; pathIndex++) {
        var pathItem = pathItems[pathIndex];

        if (pathItem === outerRect) {
            continue;
        }

        if (isVerticalLineInRect(pathItem, rectLeft, rectTop, rectRight, rectBottom, rectHeight, TOLERANCE)) {
            verticalLines.push(pathItem);
        } else if (isHorizontalLineInRect(pathItem, rectLeft, rectTop, rectRight, rectBottom, rectWidth, TOLERANCE)) {
            horizontalLines.push(pathItem);
        }
    }

    /* 縦罫は左から右へ、横罫は上から下へソート / Sort vertical L→R, horizontal T→B */
    verticalLines.sort(function (a, b) {
        return getCenterX(a) - getCenterX(b);
    });

    horizontalLines.sort(function (a, b) {
        return getCenterY(b) - getCenterY(a);
    });

    if (verticalLines.length === 0 && horizontalLines.length === 0) {
        alert(L("errNoRulesFound"));
        return;
    }

    var originalState = savePathState(pathItems);

    /* ダイアログを作成 / Create dialog */
    function createDialog(horizontalLineCount, verticalLineCount) {
        var dialog = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);

        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        var targetPanel = dialog.add("panel", undefined, L("panelTarget"));
        setupPanel(targetPanel, 6);

        var horizontalRulesCheckbox = targetPanel.add("checkbox", undefined, L("averageHorizontal"));
        horizontalRulesCheckbox.value = false;
        horizontalRulesCheckbox.enabled = horizontalLineCount > 0;

        var verticalRulesCheckbox = targetPanel.add("checkbox", undefined, L("averageVertical"));
        verticalRulesCheckbox.value = false;
        verticalRulesCheckbox.enabled = verticalLineCount > 0;

        var matchRuleLengthsCheckbox = targetPanel.add("checkbox", undefined, L("matchRuleLengths"));
        matchRuleLengthsCheckbox.value = false;
        matchRuleLengthsCheckbox.enabled = horizontalLineCount > 0 || verticalLineCount > 0;

        var previewGroup = dialog.add("group");
        previewGroup.orientation = "row";
        previewGroup.alignment = "center";

        var previewCheckbox = previewGroup.add("checkbox", undefined, L("preview"));
        previewCheckbox.value = true;

        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "center";

        var cancelButton = buttonGroup.add("button", undefined, L("cancel"));
        var okButton = buttonGroup.add("button", undefined, L("ok"));

        return {
            dialog: dialog,
            horizontalRulesCheckbox: horizontalRulesCheckbox,
            verticalRulesCheckbox: verticalRulesCheckbox,
            matchRuleLengthsCheckbox: matchRuleLengthsCheckbox,
            previewCheckbox: previewCheckbox,
            cancelButton: cancelButton,
            okButton: okButton
        };
    }

    var dialogUI = createDialog(horizontalLines.length, verticalLines.length);
    var dialog = dialogUI.dialog;
    var horizontalRulesCheckbox = dialogUI.horizontalRulesCheckbox;
    var verticalRulesCheckbox = dialogUI.verticalRulesCheckbox;
    var matchRuleLengthsCheckbox = dialogUI.matchRuleLengthsCheckbox;
    var previewCheckbox = dialogUI.previewCheckbox;
    var cancelButton = dialogUI.cancelButton;
    var okButton = dialogUI.okButton;

    horizontalRulesCheckbox.onClick = function () {
        handleRuleCheckboxClick(horizontalRulesCheckbox);
    };

    verticalRulesCheckbox.onClick = function () {
        handleRuleCheckboxClick(verticalRulesCheckbox);
    };

    previewCheckbox.onClick = function () {
        updatePreview();
    };

    okButton.onClick = function () {
        restorePathState(originalState);

        applyAverage(
            horizontalRulesCheckbox.value,
            verticalRulesCheckbox.value
        );

        app.redraw();
        dialog.close(1);
    };

    cancelButton.onClick = function () {
        restorePathState(originalState);
        app.redraw();
        dialog.close(0);
    };

    updatePreview();

    dialog.show();

    /* 罫線チェックボックスのクリック処理 / Handle rule checkbox clicks */
    function handleRuleCheckboxClick(clickedCheckbox) {
        if (ScriptUI.environment.keyboardState.altKey) {
            toggleBothRuleCheckboxes(clickedCheckbox.value);
        }

        updatePreview();
    }

    /* 横罫・縦罫をまとめてON/OFF / Toggle horizontal and vertical rule checkboxes together */
    function toggleBothRuleCheckboxes(checkboxValue) {
        if (horizontalRulesCheckbox.enabled) {
            horizontalRulesCheckbox.value = checkboxValue;
        }

        if (verticalRulesCheckbox.enabled) {
            verticalRulesCheckbox.value = checkboxValue;
        }
    }

    /* プレビュー更新 / Refresh preview */
    function updatePreview() {
        restorePathState(originalState);

        if (previewCheckbox.value) {
            applyAverage(
                horizontalRulesCheckbox.value,
                verticalRulesCheckbox.value
            );
        }

        app.redraw();
    }

    /* チェックボックスの状態に応じて平均化を実行 / Apply averaging based on checkbox states */
    function applyAverage(shouldDistributeHorizontalRules, shouldDistributeVerticalRules) {
        if (shouldDistributeVerticalRules) {
            redistributeVerticalRules();
        }

        if (shouldDistributeHorizontalRules) {
            redistributeHorizontalRules();
        }
    }

    function redistributeVerticalRules() {
        var lineCount = verticalLines.length;

        if (lineCount === 0) {
            return;
        }

        for (var lineIndex = 0; lineIndex < lineCount; lineIndex++) {
            var verticalRule = verticalLines[lineIndex];

            var targetX = rectLeft + rectWidth * (lineIndex + 1) / (lineCount + 1);

            if (MATCH_LINES_TO_RECT && verticalRule.pathPoints.length === 2) {
                setTwoPointVerticalLine(verticalRule, targetX, rectTop, rectBottom);
            } else {
                movePathToX(verticalRule, targetX);
            }
        }
    }

    function redistributeHorizontalRules() {
        var lineCount = horizontalLines.length;

        if (lineCount === 0) {
            return;
        }

        for (var lineIndex = 0; lineIndex < lineCount; lineIndex++) {
            var horizontalRule = horizontalLines[lineIndex];

            var targetY = rectTop - rectHeight * (lineIndex + 1) / (lineCount + 1);

            if (MATCH_LINES_TO_RECT && horizontalRule.pathPoints.length === 2) {
                setTwoPointHorizontalLine(horizontalRule, targetY, rectLeft, rectRight);
            } else {
                movePathToY(horizontalRule, targetY);
            }
        }
    }

    /* パネル共通設定 / Common panel setup */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            panel.spacing = spacing;
        }
    }

    /* 選択中のパスを再帰的に収集 / Recursively collect path items */
    function collectPathItems(sourceItems, collectedPathItems) {
        for (var i = 0; i < sourceItems.length; i++) {
            var pageItem = sourceItems[i];

            if (pageItem.locked || pageItem.hidden) {
                continue;
            }

            if (pageItem.typename === "PathItem") {
                collectedPathItems.push(pageItem);
            } else if (pageItem.typename === "GroupItem") {
                collectPathItems(pageItem.pageItems, collectedPathItems);
            } else if (pageItem.typename === "CompoundPathItem") {
                collectPathItems(pageItem.pathItems, collectedPathItems);
            }
        }
    }

    /* もっとも面積の大きい閉じたパスを外枠とみなす / Treat the largest closed path as the outer rectangle */
    function findOuterRectangle(pathItems) {
        var outerRectangleCandidate = null;
        var maxArea = 0;

        for (var i = 0; i < pathItems.length; i++) {
            var pathItem = pathItems[i];

            if (!pathItem.closed) {
                continue;
            }

            var geometricBounds = pathItem.geometricBounds;
            var pathWidth = geometricBounds[2] - geometricBounds[0];
            var pathHeight = geometricBounds[1] - geometricBounds[3];

            if (pathWidth <= 0 || pathHeight <= 0) {
                continue;
            }

            var pathArea = pathWidth * pathHeight;

            if (pathArea > maxArea) {
                maxArea = pathArea;
                outerRectangleCandidate = pathItem;
            }
        }

        return outerRectangleCandidate;
    }

    /* 縦罫判定（外枠内かつ縦長）/ Detect a vertical rule inside the outer rectangle */
    function isVerticalLineInRect(pathItem, left, top, right, bottom, outerHeight, tolerance) {
        if (pathItem.closed) {
            return false;
        }

        var geometricBounds = pathItem.geometricBounds;

        var pathWidth = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var pathHeight = Math.abs(geometricBounds[1] - geometricBounds[3]);

        if (pathHeight <= 0) {
            return false;
        }

        if (!(pathWidth <= tolerance || pathHeight > pathWidth * 5)) {
            return false;
        }

        var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
        var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;

        if (centerX <= left + tolerance || centerX >= right - tolerance) {
            return false;
        }

        if (centerY < bottom - tolerance || centerY > top + tolerance) {
            return false;
        }

        if (pathHeight < outerHeight * 0.2) {
            return false;
        }

        return true;
    }

    /* 横罫判定（外枠内かつ横長）/ Detect a horizontal rule inside the outer rectangle */
    function isHorizontalLineInRect(pathItem, left, top, right, bottom, outerWidth, tolerance) {
        if (pathItem.closed) {
            return false;
        }

        var geometricBounds = pathItem.geometricBounds;

        var pathWidth = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var pathHeight = Math.abs(geometricBounds[1] - geometricBounds[3]);

        if (pathWidth <= 0) {
            return false;
        }

        if (!(pathHeight <= tolerance || pathWidth > pathHeight * 5)) {
            return false;
        }

        var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
        var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;

        if (centerY <= bottom + tolerance || centerY >= top - tolerance) {
            return false;
        }

        if (centerX < left - tolerance || centerX > right + tolerance) {
            return false;
        }

        if (pathWidth < outerWidth * 0.2) {
            return false;
        }

        return true;
    }

    function getCenterX(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[0] + geometricBounds[2]) / 2;
    }

    function getCenterY(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[1] + geometricBounds[3]) / 2;
    }

    function movePathToX(pathItem, targetX) {
        var currentX = getCenterX(pathItem);
        var deltaX = targetX - currentX;

        for (var i = 0; i < pathItem.pathPoints.length; i++) {
            movePathPoint(pathItem.pathPoints[i], deltaX, 0);
        }
    }

    function movePathToY(pathItem, targetY) {
        var currentY = getCenterY(pathItem);
        var deltaY = targetY - currentY;

        for (var i = 0; i < pathItem.pathPoints.length; i++) {
            movePathPoint(pathItem.pathPoints[i], 0, deltaY);
        }
    }

    function setTwoPointVerticalLine(pathItem, x, top, bottom) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];

        var firstY = firstPoint.anchor[1];
        var secondY = secondPoint.anchor[1];

        var topPoint;
        var bottomPoint;

        if (firstY >= secondY) {
            topPoint = firstPoint;
            bottomPoint = secondPoint;
        } else {
            topPoint = secondPoint;
            bottomPoint = firstPoint;
        }

        movePointTo(topPoint, x, top);
        movePointTo(bottomPoint, x, bottom);
    }

    function setTwoPointHorizontalLine(pathItem, y, left, right) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];

        var firstX = firstPoint.anchor[0];
        var secondX = secondPoint.anchor[0];

        var leftPoint;
        var rightPoint;

        if (firstX <= secondX) {
            leftPoint = firstPoint;
            rightPoint = secondPoint;
        } else {
            leftPoint = secondPoint;
            rightPoint = firstPoint;
        }

        movePointTo(leftPoint, left, y);
        movePointTo(rightPoint, right, y);
    }

    function movePointTo(pathPoint, targetX, targetY) {
        var anchorX = pathPoint.anchor[0];
        var anchorY = pathPoint.anchor[1];

        var deltaX = targetX - anchorX;
        var deltaY = targetY - anchorY;

        movePathPoint(pathPoint, deltaX, deltaY);
    }

    function movePathPoint(pathPoint, deltaX, deltaY) {
        pathPoint.anchor = [
            pathPoint.anchor[0] + deltaX,
            pathPoint.anchor[1] + deltaY
        ];

        pathPoint.leftDirection = [
            pathPoint.leftDirection[0] + deltaX,
            pathPoint.leftDirection[1] + deltaY
        ];

        pathPoint.rightDirection = [
            pathPoint.rightDirection[0] + deltaX,
            pathPoint.rightDirection[1] + deltaY
        ];
    }

    /* 編集前のパス座標を保存 / Save original path coordinates */
    function savePathState(pathItems) {
        var pathStates = [];

        for (var pathIndex = 0; pathIndex < pathItems.length; pathIndex++) {
            var pathItem = pathItems[pathIndex];
            var pointStates = [];

            for (var pointIndex = 0; pointIndex < pathItem.pathPoints.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];

                pointStates.push({
                    anchor: [
                        pathPoint.anchor[0],
                        pathPoint.anchor[1]
                    ],
                    leftDirection: [
                        pathPoint.leftDirection[0],
                        pathPoint.leftDirection[1]
                    ],
                    rightDirection: [
                        pathPoint.rightDirection[0],
                        pathPoint.rightDirection[1]
                    ],
                    pointType: pathPoint.pointType
                });
            }

            pathStates.push({
                item: pathItem,
                points: pointStates
            });
        }

        return pathStates;
    }

    /* 保存しておいた座標へ戻す / Restore previously saved coordinates */
    function restorePathState(pathStates) {
        for (var pathIndex = 0; pathIndex < pathStates.length; pathIndex++) {
            var pathItem = pathStates[pathIndex].item;
            var pointStates = pathStates[pathIndex].points;

            for (var pointIndex = 0; pointIndex < pointStates.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];
                var pointState = pointStates[pointIndex];

                pathPoint.anchor = [
                    pointState.anchor[0],
                    pointState.anchor[1]
                ];

                pathPoint.leftDirection = [
                    pointState.leftDirection[0],
                    pointState.leftDirection[1]
                ];

                pathPoint.rightDirection = [
                    pointState.rightDirection[0],
                    pointState.rightDirection[1]
                ];

                pathPoint.pointType = pointState.pointType;
            }
        }
    }
})();
