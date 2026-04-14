#target illustrator

/*
### スクリプト名：

ReorderArtboardsByPosition.jsx

### 概要

- 見た目の配置を基準に、アートボードを左上（Y優先）で並べ替えます
- 許容差を調整しながら並びを確認し、必要に応じて再配置できます

### 主な機能

- 左上基準（Y優先）でのアートボード順並べ替え
- 許容差の自動計算とスライダーによる調整
- 行単位の並びをリストで確認（例：1 | 2 | 3）
- オプションで再配置の ON / OFF を切り替え
- 再配置する場合は、横列数・間隔（ドキュメント単位）を指定可能
- 間隔入力欄は現在の定規単位を表示
- 横列数・間隔は ↑↓キーで ±1、Shift + ↑↓キーで ±10 増減

### 処理の流れ

1. 許容差を調整して並びを確認
2. 必要に応じて［オプション］で再配置の有無、横列数、間隔を設定
3. ［OK］で左上基準の並べ替えを実行
4. ［アートボード］パネル順を更新し、再配置が ON の場合は GridByRow で再配置

### オリジナル、謝辞

- m1b 氏: https://community.adobe.com/t5/illustrator-discussions/randomly-order-artboards/m-p/12692397
- https://community.adobe.com/t5/illustrator-discussions/illustrator-script-to-renumber-reorder-the-artboards-with-there-position/m-p/12752568

### 更新履歴

- v1.0 (20231115) : 初期バージョン（Andrew_BJ による UI 改良と上限拡張）
- v1.2.0 (20260415) : 左上基準専用ツールとしてUIと構成を整理、再配置設定とプレビュー表示を調整

---

### Script Name:

ReorderArtboardsByPosition.jsx

### Overview

- Reorders artboards by visual layout using a Top Left (Y priority) order
- Adjust tolerance while previewing the order, and optionally rearrange afterward

### Main Features

- Reorder artboards using Top Left base order (Y priority)
- Auto-calculate tolerance and fine-tune it with a slider
- Preview row-based order in a list (e.g. 1 | 2 | 3)
- Toggle rearrangement on or off in Options
- When rearrangement is enabled, set column count and spacing (document units)
- Show the current ruler unit next to the spacing input
- Increase or decrease column count and spacing with Up/Down keys, or by 10 with Shift + Up/Down

### Process Flow

1. Adjust tolerance and review the order
2. In Options, set whether to rearrange, and configure column count and spacing if needed
3. Click OK to sort using Top Left order
4. Update Artboards panel order, and rearrange with GridByRow when enabled

### Original / Credit

- m1b: https://community.adobe.com/t5/illustrator-discussions/randomly-order-artboards/m-p/12692397
- https://community.adobe.com/t5/illustrator-discussions/illustrator-script-to-renumber-reorder-the-artboards-with-there-position/m-p/12752568

### Changelog

- v1.0 (20231115): Initial version (UI improvements and limit extension by Andrew_BJ)
- v1.2.0 (20260415): Refined the UI and structure as a dedicated Top Left reorder tool, and adjusted rearrange settings and preview display
*/

    (function () {

        var SCRIPT_VERSION = "v1.2.0";
        var COORDINATE_PRECISION_DIGITS = 3;

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

        app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);


        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var lang = getCurrentLang();

        function getCurrentUnitLabel() {
            var unitCode = app.preferences.getIntegerPreference("rulerType");
            return unitLabelMap[unitCode] || "pt";
        }
        var currentUnitLabel = getCurrentUnitLabel();

        /* ------------------------------- */
        /* 日英ラベル定義 / Japanese-English label definitions */
        /* ------------------------ */

        var LABELS = {
            dialogTitle: {
                ja: "アートボードの並べ替え " + SCRIPT_VERSION,
                en: "Reorder Artboards " + SCRIPT_VERSION
            },
            sort: { ja: "ソート", en: "Sort" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            tolerance: { ja: "許容差", en: "Tolerance" },
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noArtboards: { ja: "アートボードが存在しません。", en: "No artboards found." },
            preview: { ja: "並び替え", en: "Reorder" },
            execute: { ja: "OK", en: "OK" },
            close: { ja: "キャンセル", en: "Cancel" },
            rearrange: { ja: "オプション", en: "Options" },
            rearrangeEnable: { ja: "再配置", en: "Rearrange" },
            columns: { ja: "横列数：", en: "Columns:" },
            spacing: { ja: "間隔：", en: "Spacing:" },
            errorPrefix: { ja: "エラー: ", en: "Error: " }
        };


        function main() {
            if (app.documents.length === 0) {
                alert(LABELS.noDocument[lang]);
                return;
            }
            var doc = app.activeDocument;
            if (doc.artboards.length === 0) {
                alert(LABELS.noArtboards[lang]);
                return;
            }

            var context = buildContext(doc);
            var ui = buildDialogUI(context.defaultTolerance, context.sliderMax);

            bindEvents(doc, context, ui);

            ui.dialog.center();
            ui.dialog.show();
        }

        function buildContext(doc) {
            /* アートボード情報を配列に格納（プレビュー・自動計算用） */
            var artboardEntries = [];
            for (var i = 0; i < doc.artboards.length; i++) {
                artboardEntries.push({
                    name: doc.artboards[i].name,
                    artboardRect: doc.artboards[i].artboardRect
                });
            }

            /* スライダー最大値をアートボードの最大高さに設定 */
            var maxHeight = 0;
            for (var i = 0; i < artboardEntries.length; i++) {
                var r = artboardEntries[i].artboardRect;
                var h = Math.abs(r[1] - r[3]);
                if (h > maxHeight) maxHeight = h;
            }
            var sliderMax = Math.round(maxHeight);

            /* スライダー初期値を自動計算値×1.1に設定 */
            var autoTol = calculateAutoTolerance(artboardEntries);
            var defaultTolerance = Math.round(autoTol * 1.1);
            defaultTolerance = Math.min(defaultTolerance, sliderMax);

            return {
                artboardEntries: artboardEntries,
                sliderMax: sliderMax,
                defaultTolerance: defaultTolerance
            };
        }

        function bindEvents(doc, context, ui) {
            updateReorderPreview(context, ui, context.defaultTolerance);

            ui.toleranceSlider.onChanging = function () {
                var tol = Math.round(ui.toleranceSlider.value);
                updateReorderPreview(context, ui, tol);
            };

            ui.executeBtn.onClick = function () {
                executeReorder(doc, context, ui);
            };

            ui.closeBtn.onClick = function () {
                ui.dialog.close(-1);
            };
        }

        function executeReorder(doc, context, ui) {
            var tolerance = Math.round(ui.toleranceSlider.value) || 0;
            var topLeftSorter = function (_artboards, dp) {
                sortArtboardsTopLeftWithTolerance(_artboards, dp, tolerance);
            };
            rebuildArtboardsInSortedOrder(doc, topLeftSorter, COORDINATE_PRECISION_DIGITS);

            /* 再配置を実行 */
            if (ui.rearrangeEnableCheckbox.value) {
                var columns = parseInt(ui.columnsInput.text, 10) || 4;
                var spacing = parseInt(ui.spacingInput.text, 10) || 100;
                try {
                    doc.rearrangeArtboards(DocumentArtboardLayout.GridByRow, columns, spacing, true);
                } catch (e) {
                    alert(LABELS.errorPrefix[lang] + e.message);
                }
            }
            ui.dialog.close(1);
        }

        function updateReorderPreview(context, ui, tol) {
            var dp = Math.pow(10, COORDINATE_PRECISION_DIGITS);
            var sorted = [];
            for (var i = 0; i < context.artboardEntries.length; i++) {
                sorted.push({
                    name: context.artboardEntries[i].name,
                    artboardRect: context.artboardEntries[i].artboardRect
                });
            }
            sortArtboardsTopLeftWithTolerance(sorted, dp, tol);

            /* 行ごとにグルーピング */
            var rows = [];
            var currentRow = [sorted[0].name];
            for (var i = 1; i < sorted.length; i++) {
                var prevTop = Math.round(sorted[i - 1].artboardRect[1] * dp) / dp;
                var currTop = Math.round(sorted[i].artboardRect[1] * dp) / dp;
                if (Math.abs(prevTop - currTop) > tol) {
                    rows.push(currentRow);
                    currentRow = [];
                }
                currentRow.push(sorted[i].name);
            }
            rows.push(currentRow);

            ui.reorderList.removeAll();
            for (var i = 0; i < rows.length; i++) {
                ui.reorderList.add("item", rows[i].join(" | "));
            }
        }

        function buildDialogUI(defaultTolerance, sliderMax) {
            var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
            dialog.orientation = "column";
            dialog.alignChildren = ['fill', 'top'];
            dialog.spacing = 15;

            var mainColumn = dialog.add("group");
            mainColumn.orientation = "column";
            mainColumn.alignChildren = ['fill', 'top'];
            mainColumn.alignment = ['fill', 'top'];

            /* プレビューリスト */
            var reorderPanel = mainColumn.add("panel", undefined, LABELS.preview[lang]);
            reorderPanel.orientation = "column";
            reorderPanel.alignChildren = ['fill', 'top'];
            reorderPanel.margins = [15, 20, 15, 10];

            /* 許容差パネル（プレビュー上部） */
            var toleranceGroup = reorderPanel.add("group");
            toleranceGroup.orientation = "column";
            toleranceGroup.alignChildren = ['fill', 'top'];

            var tolRow = toleranceGroup.add("group");
            tolRow.orientation = "row";
            tolRow.alignChildren = ['left', 'center'];

            var toleranceSlider = tolRow.add("slider", undefined, defaultTolerance, 0, sliderMax);
            toleranceSlider.preferredSize = [250, 20];

            var reorderList = reorderPanel.add("listbox", undefined, [],
                { multiselect: false });
            reorderList.preferredSize.width = 160;
            reorderList.minimumSize.height = 120;
            reorderList.alignment = ['fill', 'fill'];

            /* 再配置パネルを mainColumn 内に追加 */
            var rearrangeGroup = mainColumn.add("panel", undefined, LABELS.rearrange[lang]);
            rearrangeGroup.orientation = "column";
            rearrangeGroup.alignChildren = ['fill', 'top'];
            rearrangeGroup.margins = [15, 20, 15, 10];

            var rearrangeEnableCheckbox = rearrangeGroup.add("checkbox", undefined, LABELS.rearrangeEnable[lang]);
            rearrangeEnableCheckbox.value = true;

            var rearrangeSettingsGroup = rearrangeGroup.add("group");
            rearrangeSettingsGroup.orientation = "column";
            rearrangeSettingsGroup.alignChildren = ['fill', 'top'];

            var colRow = rearrangeSettingsGroup.add("group");
            colRow.orientation = "row";
            colRow.alignChildren = ['left', 'center'];
            var colLabel = colRow.add("statictext", undefined, LABELS.columns[lang]);
            colLabel.preferredSize.width = 60;
            colLabel.justify = 'right';
            var columnsInput = colRow.add("edittext", undefined, "4");
            columnsInput.characters = 5;
            changeValueByArrowKey(columnsInput);

            var spcRow = rearrangeSettingsGroup.add("group");
            spcRow.orientation = "row";
            spcRow.alignChildren = ['left', 'center'];
            var spcLabel = spcRow.add("statictext", undefined, LABELS.spacing[lang]);
            spcLabel.preferredSize.width = 60;
            spcLabel.justify = 'right';
            var spacingInput = spcRow.add("edittext", undefined, "100");
            spacingInput.characters = 5;
            changeValueByArrowKey(spacingInput);
            var spacingUnit = spcRow.add("statictext", undefined, currentUnitLabel);

            function updateRearrangeSettingsEnabled() {
                rearrangeSettingsGroup.enabled = rearrangeEnableCheckbox.value;
            }
            rearrangeEnableCheckbox.onClick = updateRearrangeSettingsEnabled;
            updateRearrangeSettingsEnabled();

            var buttonGroup = dialog.add("group");
            buttonGroup.orientation = "row";
            buttonGroup.alignment = ['center', 'bottom'];

            var closeBtn = buttonGroup.add('button', undefined, LABELS.close[lang]);
            var executeBtn = buttonGroup.add('button', undefined, LABELS.execute[lang]);

            var btnWidth = Math.max(executeBtn.preferredSize.width, closeBtn.preferredSize.width);
            executeBtn.preferredSize.width = btnWidth;
            closeBtn.preferredSize.width = btnWidth;

            return {
                dialog: dialog,
                toleranceSlider: toleranceSlider,
                reorderList: reorderList,
                columnsInput: columnsInput,
                spacingInput: spacingInput,
                rearrangeEnableCheckbox: rearrangeEnableCheckbox,
                closeBtn: closeBtn,
                executeBtn: executeBtn
            };
        }

        function changeValueByArrowKey(editText) {
            editText.addEventListener("keydown", function (event) {
                var value = Number(editText.text);
                if (isNaN(value)) return;

                var keyboard = ScriptUI.environment.keyboardState;
                var delta = 1;

                if (keyboard.shiftKey) {
                    delta = 10;
                    // Shiftキー押下時は10の倍数にスナップ
                    if (event.keyName == "Up") {
                        value = Math.ceil((value + 1) / delta) * delta;
                        event.preventDefault();
                    } else if (event.keyName == "Down") {
                        value = Math.floor((value - 1) / delta) * delta;
                        if (value < 0) value = 0;
                        event.preventDefault();
                    }
                } else {
                    delta = 1;
                    if (event.keyName == "Up") {
                        value += delta;
                        event.preventDefault();
                    } else if (event.keyName == "Down") {
                        value -= delta;
                        if (value < 0) value = 0;
                        event.preventDefault();
                    }
                }

                value = Math.round(value);
                editText.text = value;
            });
        }

        /* アートボードを並べ替えて［アートボード］パネル順を再構築 / Rebuild Artboards panel order based on sorted result */
        function rebuildArtboardsInSortedOrder(doc, sorterFn, precisionDigits) {
            var decimalPlaces = Math.pow(10, precisionDigits || 3);

            /* アートボード情報を配列に格納 / Store artboard info in array */
            var _artboards = [];
            var originalArtboards = [];
            for (var i = 0; i < doc.artboards.length; i++) {
                var a = doc.artboards[i];
                _artboards.push({
                    name: a.name,
                    artboardRect: a.artboardRect,
                    srcIndex: i
                });
                originalArtboards.push({
                    name: a.name,
                    artboardRect: a.artboardRect,
                    rulerOrigin: a.rulerOrigin,
                    rulerPAR: a.rulerPAR,
                    showCenter: a.showCenter,
                    showCrossHairs: a.showCrossHairs,
                    showSafeAreas: a.showSafeAreas
                });
            }

            /* ソート実行 / Execute sort */
            sorterFn(_artboards, decimalPlaces);

            /* ソート結果順に末尾へ複製を追加 / Append duplicates in sorted order */
            for (var i = 0; i < _artboards.length; i++) {
                var s = _artboards[i];
                var src = originalArtboards[s.srcIndex];
                var b = doc.artboards.add(src.artboardRect);
                b.name = src.name;
                b.rulerOrigin = src.rulerOrigin;
                b.rulerPAR = src.rulerPAR;
                b.showCenter = src.showCenter;
                b.showCrossHairs = src.showCrossHairs;
                b.showSafeAreas = src.showSafeAreas;
            }

            /* 元のアートボードを後ろから削除 / Remove original artboards from back to front */
            for (var i = originalArtboards.length - 1; i >= 0; i--) {
                doc.artboards[i].remove();
            }
        }

        /* 許容差付き 左上基準ソート / With tolerance Top Left base */
        function sortArtboardsTopLeftWithTolerance(_artboards, decimalPlaces, tolerance) {
            sortByPositionWithTolerance(_artboards, decimalPlaces, 1, false, 0, true, tolerance);
        }

        /* 許容差を考慮した共通位置ソート関数 / Common position sort function with tolerance */
        function sortByPositionWithTolerance(_artboards, decimalPlaces, primaryIndex, primaryAsc, secondaryIndex, secondaryAsc, tolerance) {
            decimalPlaces = decimalPlaces || 1000;
            _artboards.sort(function (a, b) {
                var primaryA = Math.round(a.artboardRect[primaryIndex] * decimalPlaces) / decimalPlaces;
                var primaryB = Math.round(b.artboardRect[primaryIndex] * decimalPlaces) / decimalPlaces;
                var secondaryA = Math.round(a.artboardRect[secondaryIndex] * decimalPlaces) / decimalPlaces;
                var secondaryB = Math.round(b.artboardRect[secondaryIndex] * decimalPlaces) / decimalPlaces;

                /* 上辺の差が tolerance 以内なら、secondary で比較 / If top edge difference within tolerance, compare secondary */
                if (Math.abs(primaryA - primaryB) <= tolerance) {
                    return (secondaryAsc ? secondaryA - secondaryB : secondaryB - secondaryA);
                }
                return (primaryAsc ? primaryA - primaryB : primaryB - primaryA);
            });
        }

        /* 許容差自動計算関数 / Auto calculate tolerance function */
        function calculateAutoTolerance(_artboards) {
            var tops = [];
            for (var i = 0; i < _artboards.length; i++) {
                tops.push(_artboards[i].artboardRect[1]);
            }
            tops.sort(function (a, b) {
                return b - a;
            }); /* 上から下に並べる / Sort top to bottom */

            var diffs = [];
            for (var i = 1; i < tops.length; i++) {
                var diff = Math.abs(tops[i] - tops[i - 1]);
                if (diff > 0) {
                    diffs.push(diff);
                }
            }

            if (diffs.length === 0) {
                return 5; /* 差が無ければデフォルト / Default if no difference */
            }

            var minDiff = Math.min.apply(null, diffs);
            return minDiff + 2; /* 少しマージンを加える / Add some margin */
        }
        main();
    })();