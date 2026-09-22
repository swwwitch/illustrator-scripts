#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した外枠の長方形を基準に、内部の縦罫・横罫を等間隔に再配置します。
最大の長方形を外枠として判定し、縦罫は左右、横罫は上下方向に均等配置します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TableRuleAverager.md

### Overview

Evenly redistributes the internal vertical and horizontal rules, using the largest selected rectangle as the outer frame.
Vertical rules are spaced horizontally and horizontal rules vertically.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TableRuleAverager.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TableRuleAverager";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TableRuleAverager.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TableRuleAverager.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 罫線／長方形分類用のしきい値（pt） / Thresholds for rule/rect classification (pt) */
    var RULE_THICKNESS_THRESHOLD = 2;   /* これより細いものを罫線とみなす / Thinner than this counts as a rule */
    var RULE_LENGTH_THRESHOLD = 10;     /* 罫線とみなす最小の長さ / Minimum length of a rule */
    var RECT_SIZE_THRESHOLD = 5;        /* 外枠とみなす最小の幅・高さ / Minimum width and height of the frame */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 20;                /* ダイアログの余白 / Dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];   /* パネル余白 [左,上,右,下] / Panel margins [L,T,R,B] */

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
            title: { ja: "罫線の均等配置", en: "Even Rule Distribution" }
        },
        panel: {
            averaging: { ja: "均等配置の対象", en: "Equalize Targets" }
        },
        checkbox: {
            vertical: { ja: "縦罫", en: "Vertical" },
            horizontal: { ja: "横罫", en: "Horizontal" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            vertical: { ja: "縦罫の間隔を均等にします。", en: "Evens out the spacing of the vertical rules." },
            horizontal: { ja: "横罫の間隔を均等にします。", en: "Evens out the spacing of the horizontal rules." },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original layout."
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectRectAndRules: { ja: "長方形と罫線を選択してください。", en: "Please select a rectangle and rules." },
            noRect: { ja: "外枠となる長方形が見つかりませんでした。", en: "Could not find a bounding rectangle." }
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

    // =========================================
    // ダイアログUI / Dialog UI
    // =========================================

    /**
     * ダイアログを組み立て、プレビューとボタンの処理を結び付けて返す
     * @param {Object} boundingRect - 外枠の範囲（bounds / width / height）
     * @param {Object[]} verticalLines - 縦罫の記録（createRuleRecord() の戻り値）
     * @param {Object[]} horizontalLines - 横罫の記録（createRuleRecord() の戻り値）
     * @returns {Window} 組み立てたダイアログ
     */
    function buildDialog(boundingRect, verticalLines, horizontalLines) {

        var averagerDialog = new Window("dialog", getLabel("dialog.title") + ' ' + SCRIPT_VERSION);
        averagerDialog.orientation = "column";
        averagerDialog.alignChildren = ["left", "top"];
        averagerDialog.margins = DIALOG_MARGINS;

        var averagingPanel = averagerDialog.add("panel", undefined, getLabel("panel.averaging"));
        averagingPanel.orientation = "column";
        averagingPanel.alignChildren = ["left", "top"];
        averagingPanel.alignment = ["fill", "top"];
        averagingPanel.margins = PANEL_MARGINS;
        var verticalCheckbox = averagingPanel.add("checkbox", undefined, getLabel("checkbox.vertical"));
        verticalCheckbox.helpTip = getLabel("tooltip.vertical");
        var horizontalCheckbox = averagingPanel.add("checkbox", undefined, getLabel("checkbox.horizontal"));
        horizontalCheckbox.helpTip = getLabel("tooltip.horizontal");

        var previewGroup = averagerDialog.add("group");
        previewGroup.alignment = "center";
        var previewCheckbox = previewGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.helpTip = getLabel("tooltip.preview");

        /* 初期状態：見つかった罫線のチェックをON / Initial state: enable checkboxes for found rules */
        verticalCheckbox.value = (verticalLines.length > 0);
        horizontalCheckbox.value = (horizontalLines.length > 0);
        previewCheckbox.value = true;

        var btnRowGroup = averagerDialog.add("group");
        btnRowGroup.alignment = "center";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 元の位置に戻してから、チェックの状態に合わせて罫線を均等配置する
         * @returns {void}
         */
        function updatePreview() {
            resetRulePositions(verticalLines, horizontalLines);

            if (previewCheckbox.value) {
                if (verticalCheckbox.value) {
                    distributeVerticalRules(boundingRect, verticalLines);
                }
                if (horizontalCheckbox.value) {
                    distributeHorizontalRules(boundingRect, horizontalLines);
                }
            }

            app.redraw();
        }

        /**
         * チェックボックスのクリック処理を設定する。option＋クリックでもう一方を逆の状態にする
         * @param {Checkbox} clickedCheckbox - クリックされるチェックボックス
         * @param {Checkbox} otherCheckbox - 逆の状態にするチェックボックス
         * @returns {void}
         */
        function bindRuleCheckbox(clickedCheckbox, otherCheckbox) {
            clickedCheckbox.onClick = function () {
                if (ScriptUI.environment.keyboardState.altKey) {
                    otherCheckbox.value = !clickedCheckbox.value;
                }
                updatePreview();
            };
        }

        /* option+クリックで縦罫／横罫を互い違いに / Option+click toggles V and H to opposite states */
        bindRuleCheckbox(verticalCheckbox, horizontalCheckbox);
        bindRuleCheckbox(horizontalCheckbox, verticalCheckbox);
        previewCheckbox.onClick = updatePreview;

        btnCancel.onClick = function () {
            /* キャンセル時は元に戻して閉じる / Restore positions on cancel */
            resetRulePositions(verticalLines, horizontalLines);
            app.redraw();
            averagerDialog.close();
        };

        btnOK.onClick = function () {
            /* プレビューOFFでもOK時は最終設定を反映 / Apply final settings even if preview is off */
            previewCheckbox.value = true;
            updatePreview();
            averagerDialog.close();
        };

        /* 初期表示時のプレビュー反映 / Apply preview on initial display */
        updatePreview();

        return averagerDialog;
    }

    // =========================================
    // 罫線の配置 / Rule placement
    // =========================================

    /**
     * すべての罫線を元の位置・サイズに戻す
     * @param {Object[]} verticalLines - 縦罫の記録
     * @param {Object[]} horizontalLines - 横罫の記録
     * @returns {void}
     */
    function resetRulePositions(verticalLines, horizontalLines) {
        var i;
        for (i = 0; i < verticalLines.length; i++) {
            verticalLines[i].pathItem.height = verticalLines[i].originalHeight;
            verticalLines[i].pathItem.left = verticalLines[i].originalLeft;
            verticalLines[i].pathItem.top = verticalLines[i].originalTop;
        }
        for (i = 0; i < horizontalLines.length; i++) {
            horizontalLines[i].pathItem.width = horizontalLines[i].originalWidth;
            horizontalLines[i].pathItem.left = horizontalLines[i].originalLeft;
            horizontalLines[i].pathItem.top = horizontalLines[i].originalTop;
        }
    }

    /**
     * 縦罫を外枠の幅の n+1 分割位置に並べる
     * @param {Object} boundingRect - 外枠の範囲
     * @param {Object[]} verticalLines - 左から右に並んだ縦罫の記録
     * @returns {void}
     */
    function distributeVerticalRules(boundingRect, verticalLines) {
        var verticalSpacing = boundingRect.width / (verticalLines.length + 1);
        for (var i = 0; i < verticalLines.length; i++) {
            var targetX = boundingRect.bounds[0] + verticalSpacing * (i + 1);
            var horizontalDelta = targetX - verticalLines[i].originalX;
            verticalLines[i].pathItem.left = verticalLines[i].originalLeft + horizontalDelta;
        }
    }

    /**
     * 横罫を外枠の高さの n+1 分割位置に並べる
     * @param {Object} boundingRect - 外枠の範囲
     * @param {Object[]} horizontalLines - 上から下に並んだ横罫の記録
     * @returns {void}
     */
    function distributeHorizontalRules(boundingRect, horizontalLines) {
        var horizontalSpacing = boundingRect.height / (horizontalLines.length + 1);
        for (var i = 0; i < horizontalLines.length; i++) {
            var targetY = boundingRect.bounds[1] - horizontalSpacing * (i + 1);
            var verticalDelta = targetY - horizontalLines[i].originalY;
            horizontalLines[i].pathItem.top = horizontalLines[i].originalTop + verticalDelta;
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 罫線の元の位置・サイズを控えた記録を作る
     * @param {PathItem} pathItem - 罫線のパス
     * @param {number[]} itemBounds - pathItem の geometricBounds
     * @returns {Object} 罫線の記録
     */
    function createRuleRecord(pathItem, itemBounds) {
        return {
            pathItem: pathItem,
            originalLeft: pathItem.left,
            originalTop: pathItem.top,
            originalWidth: pathItem.width,
            originalHeight: pathItem.height,
            originalX: itemBounds[0],
            originalY: itemBounds[1]
        };
    }

    /**
     * 選択中のパスを縦罫・横罫・外枠（最大の長方形）に振り分ける
     * @param {PageItem[]} selectedItems - 選択中のアイテム
     * @returns {{boundingRect: Object, verticalLines: Object[], horizontalLines: Object[]}} 振り分けの結果（外枠が無ければ boundingRect は null）
     */
    function classifySelection(selectedItems) {
        var boundingRect = null;
        var verticalLines = [];
        var horizontalLines = [];
        var maxRectArea = 0;

        for (var i = 0; i < selectedItems.length; i++) {
            var pageItem = selectedItems[i];
            if (pageItem.typename !== "PathItem") {
                continue;
            }

            var itemBounds = pageItem.geometricBounds; // [left, top, right, bottom]
            var itemWidth = itemBounds[2] - itemBounds[0];
            var itemHeight = itemBounds[1] - itemBounds[3];

            if (itemWidth < RULE_THICKNESS_THRESHOLD && itemHeight > RULE_LENGTH_THRESHOLD) {
                verticalLines.push(createRuleRecord(pageItem, itemBounds));
            } else if (itemHeight < RULE_THICKNESS_THRESHOLD && itemWidth > RULE_LENGTH_THRESHOLD) {
                horizontalLines.push(createRuleRecord(pageItem, itemBounds));
            } else if (itemWidth > RECT_SIZE_THRESHOLD && itemHeight > RECT_SIZE_THRESHOLD) {
                var itemArea = itemWidth * itemHeight;
                if (itemArea > maxRectArea) {
                    maxRectArea = itemArea;
                    boundingRect = { bounds: itemBounds, width: itemWidth, height: itemHeight };
                }
            }
        }

        return { boundingRect: boundingRect, verticalLines: verticalLines, horizontalLines: horizontalLines };
    }

    /**
     * 選択を振り分けてダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var selectedItems = doc.selection;

        if (selectedItems.length < 2) {
            alert(getLabel("alert.selectRectAndRules"));
            return;
        }

        /* 選択オブジェクトを分類して初期位置を保存 / Classify selection and save initial positions */
        var classified = classifySelection(selectedItems);

        if (!classified.boundingRect) {
            alert(getLabel("alert.noRect"));
            return;
        }

        /* 縦罫はX座標で昇順、横罫はY座標で降順にソート / Sort V rules left-to-right, H rules top-to-bottom */
        classified.verticalLines.sort(function (firstRule, secondRule) {
            return firstRule.originalX - secondRule.originalX;
        });
        classified.horizontalLines.sort(function (firstRule, secondRule) {
            return secondRule.originalY - firstRule.originalY;
        });

        var averagerDialog = buildDialog(classified.boundingRect, classified.verticalLines, classified.horizontalLines);
        averagerDialog.show();
    }

    main();

})();
