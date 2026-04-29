#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要 / Overview

選択した外枠の長方形を基準に、内部の縦罫／横罫を等間隔に再配置します。
最大の長方形を外枠として判定し、縦罫は左右、横罫は上下方向に均等配置します。
プレビューは初期ON。縦罫／横罫の切替に対応し、option+クリックで一方ON・他方OFF。
キャンセルで元の位置とサイズに復元します。

Evenly redistributes internal vertical and horizontal rules based on the
largest selected rectangle as the outer frame. Vertical rules are spaced
horizontally and horizontal rules vertically. Preview is ON by default.
Option+click toggles one direction ON and the other OFF. Cancel restores
original positions and sizes.
*/

(function () {

    // =========================================
    // バージョンとローカライズ / Version and Localization
    // =========================================

    var SCRIPT_VERSION = "v1.0";

    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "罫線の均等配置",
            en: "Even Rule Distribution"
        },
        panelAveraging: {
            ja: "均等配置の対象",
            en: "Equalize Targets"
        },
        checkboxVertical: {
            ja: "縦罫",
            en: "Vertical"
        },
        checkboxHorizontal: {
            ja: "横罫",
            en: "Horizontal"
        },
        checkboxPreview: {
            ja: "プレビュー",
            en: "Preview"
        },
        buttonCancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        buttonOk: {
            ja: "OK",
            en: "OK"
        },
        errNoDoc: {
            ja: "ドキュメントを開いてください。",
            en: "Please open a document."
        },
        errSelect: {
            ja: "長方形と罫線を選択してください。",
            en: "Please select a rectangle and rules."
        },
        errNoRect: {
            ja: "外枠となる長方形が見つかりませんでした。",
            en: "Could not find a bounding rectangle."
        }
    };

    /* ラベル取得。未定義キーはキー名を返す / Get a localized label. Return the key when missing. */
    function getLabel(key) {
        if (!LABELS[key]) return key;
        return LABELS[key][currentLanguage] || LABELS[key].en || key;
    }


    // =========================================
    // ダイアログUI / Dialog UI
    // =========================================

    /* ダイアログを構築して返す / Build and return the dialog */
    function createDialog(boundingRect, verticalLines, horizontalLines) {

        var dialog = new Window("dialog", getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = ["left", "top"];
        dialog.margins = 20;

        var averagingPanel = dialog.add("panel", undefined, getLabel('panelAveraging'));
        averagingPanel.orientation = "column";
        averagingPanel.alignChildren = ["left", "top"];
        averagingPanel.alignment = ["fill", "top"];
        averagingPanel.margins = [15, 20, 15, 10];
        var verticalCheckbox = averagingPanel.add("checkbox", undefined, getLabel('checkboxVertical'));
        var horizontalCheckbox = averagingPanel.add("checkbox", undefined, getLabel('checkboxHorizontal'));

        var previewGroup = dialog.add("group");
        previewGroup.alignment = "center";
        var previewCheckbox = previewGroup.add("checkbox", undefined, getLabel('checkboxPreview'));

        /* 初期状態：見つかった罫線のチェックをON / Initial state: enable checkboxes for found rules */
        verticalCheckbox.value = (verticalLines.length > 0);
        horizontalCheckbox.value = (horizontalLines.length > 0);
        previewCheckbox.value = true;

        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "center";
        var cancelButton = buttonGroup.add("button", undefined, getLabel('buttonCancel'), { name: "cancel" });
        var okButton = buttonGroup.add("button", undefined, getLabel('buttonOk'), { name: "ok" });

        /* すべての罫線を元の位置・サイズに戻す / Reset all rules to original position and size */
        function resetPositions() {
            var i;
            for (i = 0; i < verticalLines.length; i++) {
                verticalLines[i].item.height = verticalLines[i].originalHeight;
                verticalLines[i].item.left = verticalLines[i].originalLeft;
                verticalLines[i].item.top = verticalLines[i].originalTop;
            }
            for (i = 0; i < horizontalLines.length; i++) {
                horizontalLines[i].item.width = horizontalLines[i].originalWidth;
                horizontalLines[i].item.left = horizontalLines[i].originalLeft;
                horizontalLines[i].item.top = horizontalLines[i].originalTop;
            }
        }

        /* プレビュー更新 / Update preview */
        function updatePreview() {
            resetPositions();

            if (previewCheckbox.value) {
                var i;
                /* 縦罫を均等配置（n+1分割）/ Distribute vertical rules (n+1 divisions) */
                if (verticalCheckbox.value && verticalLines.length > 0) {
                    var verticalSpacing = boundingRect.width / (verticalLines.length + 1);
                    for (i = 0; i < verticalLines.length; i++) {
                        var targetX = boundingRect.bounds[0] + verticalSpacing * (i + 1);
                        var horizontalDelta = targetX - verticalLines[i].originalX;
                        verticalLines[i].item.left = verticalLines[i].originalLeft + horizontalDelta;
                    }
                }

                /* 横罫を均等配置 / Distribute horizontal rules */
                if (horizontalCheckbox.value && horizontalLines.length > 0) {
                    var horizontalSpacing = boundingRect.height / (horizontalLines.length + 1);
                    for (i = 0; i < horizontalLines.length; i++) {
                        var targetY = boundingRect.bounds[1] - horizontalSpacing * (i + 1);
                        var verticalDelta = targetY - horizontalLines[i].originalY;
                        horizontalLines[i].item.top = horizontalLines[i].originalTop + verticalDelta;
                    }
                }
            }

            app.redraw();
        }

        /* option+クリックで縦罫／横罫を互い違いに / Option+click toggles V and H to opposite states */
        verticalCheckbox.onClick = function () {
            if (ScriptUI.environment.keyboardState.altKey) {
                horizontalCheckbox.value = !verticalCheckbox.value;
            }
            updatePreview();
        };
        horizontalCheckbox.onClick = function () {
            if (ScriptUI.environment.keyboardState.altKey) {
                verticalCheckbox.value = !horizontalCheckbox.value;
            }
            updatePreview();
        };
        previewCheckbox.onClick = updatePreview;

        cancelButton.onClick = function () {
            /* キャンセル時は元に戻して閉じる / Restore positions on cancel */
            resetPositions();
            app.redraw();
            dialog.close();
        };

        okButton.onClick = function () {
            /* プレビューOFFでもOK時は最終設定を反映 / Apply final settings even if preview is off */
            previewCheckbox.value = true;
            updatePreview();
            dialog.close();
        };

        /* 初期表示時のプレビュー反映 / Apply preview on initial display */
        updatePreview();

        return dialog;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('errNoDoc'));
            return;
        }

        var doc = app.activeDocument;
        var selection = doc.selection;

        if (selection.length < 2) {
            alert(getLabel('errSelect'));
            return;
        }

        var boundingRect = null;
        var verticalLines = [];
        var horizontalLines = [];
        var maxRectArea = 0;

        /* 罫線／長方形分類用のしきい値 / Thresholds for rule/rect classification */
        var RULE_THICKNESS_THRESHOLD = 2;
        var RULE_LENGTH_THRESHOLD = 10;
        var RECT_SIZE_THRESHOLD = 5;

        /* 選択オブジェクトを分類して初期位置を保存 / Classify selection and save initial positions */
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename === "PathItem") {
                var bounds = item.geometricBounds; // [left, top, right, bottom]
                var width = bounds[2] - bounds[0];
                var height = bounds[1] - bounds[3];

                if (width < RULE_THICKNESS_THRESHOLD && height > RULE_LENGTH_THRESHOLD) {
                    var lineRecord = {
                        item: item,
                        originalLeft: item.left,
                        originalTop: item.top,
                        originalWidth: item.width,
                        originalHeight: item.height,
                        originalX: bounds[0],
                        originalY: bounds[1]
                    };
                    verticalLines.push(lineRecord);
                } else if (height < RULE_THICKNESS_THRESHOLD && width > RULE_LENGTH_THRESHOLD) {
                    var lineRecord = {
                        item: item,
                        originalLeft: item.left,
                        originalTop: item.top,
                        originalWidth: item.width,
                        originalHeight: item.height,
                        originalX: bounds[0],
                        originalY: bounds[1]
                    };
                    horizontalLines.push(lineRecord);
                } else if (width > RECT_SIZE_THRESHOLD && height > RECT_SIZE_THRESHOLD) {
                    var area = width * height;
                    if (area > maxRectArea) {
                        maxRectArea = area;
                        boundingRect = { bounds: bounds, width: width, height: height };
                    }
                }
            }
        }

        if (!boundingRect) {
            alert(getLabel('errNoRect'));
            return;
        }

        /* 縦罫はX座標で昇順、横罫はY座標で降順にソート / Sort V rules left-to-right, H rules top-to-bottom */
        verticalLines.sort(function (a, b) { return a.originalX - b.originalX; });
        horizontalLines.sort(function (a, b) { return b.originalY - a.originalY; });

        var dialog = createDialog(boundingRect, verticalLines, horizontalLines);
        dialog.show();
    }

    main();

})();