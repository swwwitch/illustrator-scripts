#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した外枠の長方形と、その内側にある縦罫・横罫を、外枠の中で均等に再配置します。
横罫・縦罫はそれぞれON/OFFでき、プレビューの切り替えにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AverageRulesInRectangle.md

### Overview

Evenly redistributes the vertical and horizontal rules inside the selected outer rectangle.
Horizontal and vertical rules can be toggled independently, with a live preview switch.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AverageRulesInRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AverageRulesInRectangle";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AverageRulesInRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AverageRulesInRectangle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /*
        true  : 縦罫の上下、横罫の左右を外枠に合わせる
                Snap rule endpoints to the outer rectangle.
        false : 罫線の長さは変えず、位置だけ平均化する
                Keep rule lengths; redistribute positions only.
    */
    var MATCH_LINES_TO_RECT = true;

    /* 罫線の判定に使う許容値（pt） / Tolerance for detecting rules (pt) */
    var TOLERANCE = 0.5;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] / Panel margins [L,T,R,B] */

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
            title: { ja: "罫線を平均化", en: "Distribute Rules" }
        },
        panel: {
            target: { ja: "対象", en: "Target" }
        },
        checkbox: {
            averageHorizontal: { ja: "横罫を平均化", en: "Distribute horizontal rules" },
            averageVertical: { ja: "縦罫を平均化", en: "Distribute vertical rules" },
            matchRuleLengths: { ja: "長さを揃える", en: "Match rule lengths" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            averageHorizontal: {
                ja: "長方形の中にある横罫の間隔を均等にします。",
                en: "Evens out the spacing of the horizontal rules inside the rectangle."
            },
            averageVertical: {
                ja: "長方形の中にある縦罫の間隔を均等にします。",
                en: "Evens out the spacing of the vertical rules inside the rectangle."
            },
            matchRuleLengths: {
                ja: "罫の長さを長方形の辺にそろえます。",
                en: "Matches the length of the rules to the sides of the rectangle."
            },
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
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "外枠の長方形と罫線を選択してください。", en: "Please select the outer rectangle and the rules." },
            noPathInSelection: { ja: "選択内にパスがありません。", en: "No paths found in the selection." },
            outerRectNotFound: { ja: "外枠の長方形が見つかりませんでした。", en: "Outer rectangle not found." },
            noRulesFound: { ja: "縦罫または横罫が見つかりませんでした。", en: "No vertical or horizontal rules were found." }
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
            if (labelNode === undefined || labelNode === null) {
                return labelPath;
            }
        }
        return labelNode[uiLang] || labelNode.en;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から外枠と罫線を見つけ、ダイアログで平均化を確定する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;

        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var selectedPathItems = [];
        collectPathItems(doc.selection, selectedPathItems);

        if (selectedPathItems.length === 0) {
            alert(getLabel("alert.noPathInSelection"));
            return;
        }

        var outerRect = findOuterRectangle(selectedPathItems);

        if (!outerRect) {
            alert(getLabel("alert.outerRectNotFound"));
            return;
        }

        var rectBounds = getRectBounds(outerRect);
        var rules = classifyRules(selectedPathItems, outerRect, rectBounds);

        if (rules.vertical.length === 0 && rules.horizontal.length === 0) {
            alert(getLabel("alert.noRulesFound"));
            return;
        }

        var originalState = savePathState(selectedPathItems);

        var dialogControls = buildDialog(rules.horizontal.length, rules.vertical.length);
        var horizontalRulesCheckbox = dialogControls.horizontalRulesCheckbox;
        var verticalRulesCheckbox = dialogControls.verticalRulesCheckbox;
        var previewCheckbox = dialogControls.previewCheckbox;

        /**
         * 罫線チェックボックスのクリック処理（Option／Alt キーで横罫・縦罫をまとめて切り替える）
         * @param {Checkbox} clickedCheckbox - クリックされたチェックボックス
         * @returns {void}
         */
        function handleRuleCheckboxClick(clickedCheckbox) {
            if (ScriptUI.environment.keyboardState.altKey) {
                toggleBothRuleCheckboxes(clickedCheckbox.value);
            }
            updatePreview();
        }

        /**
         * 有効な横罫・縦罫のチェックボックスをまとめて同じ状態にする
         * @param {boolean} checkboxValue - 設定する値
         * @returns {void}
         */
        function toggleBothRuleCheckboxes(checkboxValue) {
            if (horizontalRulesCheckbox.enabled) {
                horizontalRulesCheckbox.value = checkboxValue;
            }
            if (verticalRulesCheckbox.enabled) {
                verticalRulesCheckbox.value = checkboxValue;
            }
        }

        /**
         * 元の座標に戻したうえで、チェックボックスの状態どおりに平均化する
         * @returns {void}
         */
        function applyCheckedRules() {
            restorePathState(originalState);
            if (verticalRulesCheckbox.value) {
                redistributeVerticalRules(rules.vertical, rectBounds);
            }
            if (horizontalRulesCheckbox.value) {
                redistributeHorizontalRules(rules.horizontal, rectBounds);
            }
        }

        /**
         * プレビューを更新する（プレビューが OFF のときは元の座標に戻すだけ）
         * @returns {void}
         */
        function updatePreview() {
            if (previewCheckbox.value) {
                applyCheckedRules();
            } else {
                restorePathState(originalState);
            }
            app.redraw();
        }

        horizontalRulesCheckbox.onClick = function () {
            handleRuleCheckboxClick(horizontalRulesCheckbox);
        };

        verticalRulesCheckbox.onClick = function () {
            handleRuleCheckboxClick(verticalRulesCheckbox);
        };

        previewCheckbox.onClick = function () {
            updatePreview();
        };

        dialogControls.btnOK.onClick = function () {
            applyCheckedRules();
            app.redraw();
            dialogControls.rulesDialog.close(1);
        };

        dialogControls.btnCancel.onClick = function () {
            restorePathState(originalState);
            app.redraw();
            dialogControls.rulesDialog.close(0);
        };

        updatePreview();

        dialogControls.rulesDialog.show();
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを組み立てる
     * @param {number} horizontalRuleCount - 見つかった横罫の数
     * @param {number} verticalRuleCount - 見つかった縦罫の数
     * @returns {{rulesDialog: Window, horizontalRulesCheckbox: Checkbox, verticalRulesCheckbox: Checkbox, matchRuleLengthsCheckbox: Checkbox, previewCheckbox: Checkbox, btnCancel: Button, btnOK: Button}} ダイアログとコントロール
     */
    function buildDialog(horizontalRuleCount, verticalRuleCount) {
        var rulesDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);

        rulesDialog.orientation = "column";
        rulesDialog.alignChildren = "fill";

        var targetPanel = rulesDialog.add("panel", undefined, getLabel("panel.target"));
        setupPanel(targetPanel, 6);

        var horizontalRulesCheckbox = addOptionCheckbox(targetPanel,
            "checkbox.averageHorizontal", "tooltip.averageHorizontal", horizontalRuleCount > 0);
        var verticalRulesCheckbox = addOptionCheckbox(targetPanel,
            "checkbox.averageVertical", "tooltip.averageVertical", verticalRuleCount > 0);
        var matchRuleLengthsCheckbox = addOptionCheckbox(targetPanel,
            "checkbox.matchRuleLengths", "tooltip.matchRuleLengths", horizontalRuleCount > 0 || verticalRuleCount > 0);

        var previewGroup = rulesDialog.add("group");
        previewGroup.orientation = "row";
        previewGroup.alignment = "center";

        var previewCheckbox = previewGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");
        previewCheckbox.value = true;

        var btnRowGroup = rulesDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "center";

        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"));
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"));

        return {
            rulesDialog: rulesDialog,
            horizontalRulesCheckbox: horizontalRulesCheckbox,
            verticalRulesCheckbox: verticalRulesCheckbox,
            matchRuleLengthsCheckbox: matchRuleLengthsCheckbox,
            previewCheckbox: previewCheckbox,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * OFF で始まるチェックボックスを tooltip 付きで追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - 表示名の LABELS パス
     * @param {string} tooltipPath - tooltip の LABELS パス
     * @param {boolean} isEnabled - 操作できるかどうか
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, isEnabled) {
        var optionCheckbox = parentPanel.add("checkbox", undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = false;
        optionCheckbox.enabled = isEnabled;
        return optionCheckbox;
    }

    // =========================================
    // 外枠と罫線の判定 / Detecting the frame and rules
    // =========================================

    /**
     * 選択中のパスを再帰的に収集する（ロック中・非表示のものは除く）
     * @param {PageItem[]} sourceItems - 走査するアイテム
     * @param {PathItem[]} collectedPathItems - 見つけたパスを追加する配列
     * @returns {void}
     */
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

    /**
     * もっとも面積の大きい閉じたパスを外枠とみなす
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @returns {PathItem|null} 外枠のパス。見つからなければ null
     */
    function findOuterRectangle(candidatePathItems) {
        var outerRectangleCandidate = null;
        var maxArea = 0;

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];

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

    /**
     * 外枠の座標と寸法をまとめて返す
     * @param {PathItem} outerRect - 外枠のパス
     * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}} 外枠の範囲
     */
    function getRectBounds(outerRect) {
        var outerBounds = outerRect.geometricBounds;
        return {
            left: outerBounds[0],
            top: outerBounds[1],
            right: outerBounds[2],
            bottom: outerBounds[3],
            width: outerBounds[2] - outerBounds[0],
            height: outerBounds[1] - outerBounds[3]
        };
    }

    /**
     * 外枠の中にある縦罫・横罫を振り分け、縦罫は左から右、横罫は上から下に並べる
     * @param {PathItem[]} candidatePathItems - 候補のパス
     * @param {PathItem} outerRect - 外枠のパス（対象から除く）
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @returns {{vertical: PathItem[], horizontal: PathItem[]}} 縦罫と横罫
     */
    function classifyRules(candidatePathItems, outerRect, rectBounds) {
        var verticalRules = [];
        var horizontalRules = [];

        for (var i = 0; i < candidatePathItems.length; i++) {
            var pathItem = candidatePathItems[i];

            if (pathItem === outerRect) {
                continue;
            }

            if (isRuleInRect(pathItem, rectBounds, true)) {
                verticalRules.push(pathItem);
            } else if (isRuleInRect(pathItem, rectBounds, false)) {
                horizontalRules.push(pathItem);
            }
        }

        /* 縦罫は左から右へ、横罫は上から下へソート / Sort vertical L→R, horizontal T→B */
        verticalRules.sort(function (firstRule, secondRule) {
            return getCenterX(firstRule) - getCenterX(secondRule);
        });

        horizontalRules.sort(function (firstRule, secondRule) {
            return getCenterY(secondRule) - getCenterY(firstRule);
        });

        return { vertical: verticalRules, horizontal: horizontalRules };
    }

    /**
     * 外枠の中にある縦罫（または横罫）かどうかを判定する
     * 縦罫は「長さ＝高さ、並ぶ方向＝X」、横罫は「長さ＝幅、並ぶ方向＝Y」として同じ条件で見る
     * @param {PathItem} pathItem - 判定するパス
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @param {boolean} isVertical - true なら縦罫、false なら横罫として判定する
     * @returns {boolean} 罫線とみなせるとき true
     */
    function isRuleInRect(pathItem, rectBounds, isVertical) {
        if (pathItem.closed) {
            return false;
        }

        var geometricBounds = pathItem.geometricBounds;
        var pathWidth = Math.abs(geometricBounds[2] - geometricBounds[0]);
        var pathHeight = Math.abs(geometricBounds[1] - geometricBounds[3]);
        var centerX = (geometricBounds[0] + geometricBounds[2]) / 2;
        var centerY = (geometricBounds[1] + geometricBounds[3]) / 2;

        if (isVertical) {
            return hasRuleShape(pathHeight, pathWidth, rectBounds.height) &&
                isInsideRange(centerX, rectBounds.left, rectBounds.right) &&
                isNearRange(centerY, rectBounds.bottom, rectBounds.top);
        }
        return hasRuleShape(pathWidth, pathHeight, rectBounds.width) &&
            isInsideRange(centerY, rectBounds.bottom, rectBounds.top) &&
            isNearRange(centerX, rectBounds.left, rectBounds.right);
    }

    /**
     * 細長い形で、外枠の辺の 20% 以上の長さがあるかどうか
     * @param {number} ruleLength - 罫線の方向の長さ
     * @param {number} ruleThickness - 罫線と直交する方向の幅
     * @param {number} outerLength - 同じ方向の外枠の長さ
     * @returns {boolean} 罫線の形とみなせるとき true
     */
    function hasRuleShape(ruleLength, ruleThickness, outerLength) {
        if (ruleLength <= 0) {
            return false;
        }
        if (!(ruleThickness <= TOLERANCE || ruleLength > ruleThickness * 5)) {
            return false;
        }
        return !(ruleLength < outerLength * 0.2);
    }

    /**
     * 値が範囲の内側に、許容値より離れて収まっているかどうか（両端ちょうどは外）
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 内側なら true
     */
    function isInsideRange(value, rangeMin, rangeMax) {
        return value > rangeMin + TOLERANCE && value < rangeMax - TOLERANCE;
    }

    /**
     * 値が範囲に、許容値の はみ出しまで含めて収まっているかどうか
     * @param {number} value - 調べる値
     * @param {number} rangeMin - 範囲の最小値
     * @param {number} rangeMax - 範囲の最大値
     * @returns {boolean} 収まっていれば true
     */
    function isNearRange(value, rangeMin, rangeMax) {
        return value >= rangeMin - TOLERANCE && value <= rangeMax + TOLERANCE;
    }

    /**
     * パスの外接矩形の中心 X を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の X 座標
     */
    function getCenterX(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[0] + geometricBounds[2]) / 2;
    }

    /**
     * パスの外接矩形の中心 Y を返す
     * @param {PathItem} pathItem - 対象のパス
     * @returns {number} 中心の Y 座標
     */
    function getCenterY(pathItem) {
        var geometricBounds = pathItem.geometricBounds;
        return (geometricBounds[1] + geometricBounds[3]) / 2;
    }

    // =========================================
    // 罫線の再配置 / Redistributing rules
    // =========================================

    /**
     * 縦罫を外枠の幅で等間隔に並べ直す
     * @param {PathItem[]} verticalRules - 左から右に並んだ縦罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @returns {void}
     */
    function redistributeVerticalRules(verticalRules, rectBounds) {
        var ruleCount = verticalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var verticalRule = verticalRules[i];
            var targetX = rectBounds.left + rectBounds.width * (i + 1) / (ruleCount + 1);

            if (MATCH_LINES_TO_RECT && verticalRule.pathPoints.length === 2) {
                setTwoPointVerticalLine(verticalRule, targetX, rectBounds.top, rectBounds.bottom);
            } else {
                translatePath(verticalRule, targetX - getCenterX(verticalRule), 0);
            }
        }
    }

    /**
     * 横罫を外枠の高さで等間隔に並べ直す
     * @param {PathItem[]} horizontalRules - 上から下に並んだ横罫
     * @param {Object} rectBounds - getRectBounds() の戻り値
     * @returns {void}
     */
    function redistributeHorizontalRules(horizontalRules, rectBounds) {
        var ruleCount = horizontalRules.length;

        for (var i = 0; i < ruleCount; i++) {
            var horizontalRule = horizontalRules[i];
            var targetY = rectBounds.top - rectBounds.height * (i + 1) / (ruleCount + 1);

            if (MATCH_LINES_TO_RECT && horizontalRule.pathPoints.length === 2) {
                setTwoPointHorizontalLine(horizontalRule, targetY, rectBounds.left, rectBounds.right);
            } else {
                translatePath(horizontalRule, 0, targetY - getCenterY(horizontalRule));
            }
        }
    }

    /**
     * パスのすべてのアンカーポイントとハンドルを平行移動する
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function translatePath(pathItem, deltaX, deltaY) {
        for (var i = 0; i < pathItem.pathPoints.length; i++) {
            movePathPoint(pathItem.pathPoints[i], deltaX, deltaY);
        }
    }

    /**
     * 2点の縦罫を、指定の X 位置で上端から下端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} x - 配置する X 座標
     * @param {number} top - 上端の Y 座標
     * @param {number} bottom - 下端の Y 座標
     * @returns {void}
     */
    function setTwoPointVerticalLine(pathItem, x, top, bottom) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnTop = firstPoint.anchor[1] >= secondPoint.anchor[1];

        movePointTo(isFirstOnTop ? firstPoint : secondPoint, x, top);
        movePointTo(isFirstOnTop ? secondPoint : firstPoint, x, bottom);
    }

    /**
     * 2点の横罫を、指定の Y 位置で左端から右端までの線にする
     * @param {PathItem} pathItem - 2点のパス
     * @param {number} y - 配置する Y 座標
     * @param {number} left - 左端の X 座標
     * @param {number} right - 右端の X 座標
     * @returns {void}
     */
    function setTwoPointHorizontalLine(pathItem, y, left, right) {
        var firstPoint = pathItem.pathPoints[0];
        var secondPoint = pathItem.pathPoints[1];
        var isFirstOnLeft = firstPoint.anchor[0] <= secondPoint.anchor[0];

        movePointTo(isFirstOnLeft ? firstPoint : secondPoint, left, y);
        movePointTo(isFirstOnLeft ? secondPoint : firstPoint, right, y);
    }

    /**
     * アンカーポイントを指定の座標へ移動する（ハンドルも同じ量だけ動かす）
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} targetX - 移動先の X 座標
     * @param {number} targetY - 移動先の Y 座標
     * @returns {void}
     */
    function movePointTo(pathPoint, targetX, targetY) {
        movePathPoint(pathPoint, targetX - pathPoint.anchor[0], targetY - pathPoint.anchor[1]);
    }

    /**
     * アンカーポイントと両ハンドルを同じ量だけ動かす
     * @param {PathPoint} pathPoint - 対象のアンカーポイント
     * @param {number} deltaX - X 方向の移動量
     * @param {number} deltaY - Y 方向の移動量
     * @returns {void}
     */
    function movePathPoint(pathPoint, deltaX, deltaY) {
        pathPoint.anchor = [pathPoint.anchor[0] + deltaX, pathPoint.anchor[1] + deltaY];
        pathPoint.leftDirection = [pathPoint.leftDirection[0] + deltaX, pathPoint.leftDirection[1] + deltaY];
        pathPoint.rightDirection = [pathPoint.rightDirection[0] + deltaX, pathPoint.rightDirection[1] + deltaY];
    }

    // =========================================
    // 座標の保存と復元 / Saving and restoring coordinates
    // =========================================

    /**
     * 編集前のパスの座標を保存する
     * @param {PathItem[]} targetPathItems - 保存するパス
     * @returns {Object[]} パスごとのアンカーポイントの状態
     */
    function savePathState(targetPathItems) {
        var pathStates = [];

        for (var pathIndex = 0; pathIndex < targetPathItems.length; pathIndex++) {
            var pathItem = targetPathItems[pathIndex];
            var pointStates = [];

            for (var pointIndex = 0; pointIndex < pathItem.pathPoints.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];

                pointStates.push({
                    anchor: [pathPoint.anchor[0], pathPoint.anchor[1]],
                    leftDirection: [pathPoint.leftDirection[0], pathPoint.leftDirection[1]],
                    rightDirection: [pathPoint.rightDirection[0], pathPoint.rightDirection[1]],
                    pointType: pathPoint.pointType
                });
            }

            pathStates.push({ pathItem: pathItem, pointStates: pointStates });
        }

        return pathStates;
    }

    /**
     * 保存しておいた座標へ戻す
     * @param {Object[]} pathStates - savePathState() の戻り値
     * @returns {void}
     */
    function restorePathState(pathStates) {
        for (var pathIndex = 0; pathIndex < pathStates.length; pathIndex++) {
            var pathItem = pathStates[pathIndex].pathItem;
            var pointStates = pathStates[pathIndex].pointStates;

            for (var pointIndex = 0; pointIndex < pointStates.length; pointIndex++) {
                var pathPoint = pathItem.pathPoints[pointIndex];
                var pointState = pointStates[pointIndex];

                pathPoint.anchor = [pointState.anchor[0], pointState.anchor[1]];
                pathPoint.leftDirection = [pointState.leftDirection[0], pointState.leftDirection[1]];
                pathPoint.rightDirection = [pointState.rightDirection[0], pointState.rightDirection[1]];
                pathPoint.pointType = pointState.pointType;
            }
        }
    }

    main();

})();
