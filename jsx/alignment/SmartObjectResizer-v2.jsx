#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartObjectResizer.jsx

### 概要

- 選択中のオブジェクトを指定した基準（最大、最小、指定サイズ、アートボード、裁ち落とし、面積）に基づいて柔軟にリサイズできるIllustrator用スクリプトです。
- 縦横比保持モードや片辺のみモード、整列オプション、リアルタイムプレビュー機能を備えています。

### 主な機能

- 縦横比保持と片辺のみの切り替え
- 各種基準（最大、最小、指定サイズ、アートボード、裁ち落とし、面積）でのスケーリング
- 整列オプション（左、中央、均等、0間隔、上、中央、横均等、横0間隔）
- リアルタイムプレビューとリセット機能
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログでリサイズモードと基準を選択
2. 選択モードに従ってスケーリング実行（必要に応じて一時グループ化）
3. 整列オプションや面積一致処理を適用
4. OKで確定、キャンセルやリセットで元に戻す

### 更新履歴

- v1.0.0 (20250601) : 初期バージョン
- v1.1.0 (20250601) : 「片辺のみ」＋「指定サイズ」での変形サポート
- v1.2.0 (20250601) : バグ修正、細部改善

---

### Script Name:

SmartObjectResizer.jsx

### Overview

- An Illustrator script that flexibly resizes selected objects based on criteria such as Max, Min, Fixed Size, Artboard, Bleed, or Area.
- Supports "Keep Aspect" and "One Side Only" modes, alignment options, and real-time preview.

### Main Features

- Toggle between "Keep Aspect" and "One Side Only"
- Scaling based on Max, Min, Fixed Size, Artboard, Bleed, or Area
- Alignment options (Left, Center, Evenly, Zero Gap, Top, Middle, Horizontal Evenly, Horizontal Zero Gap)
- Real-time preview and reset function
- Japanese and English UI support

### Process Flow

1. Select resize mode and base in the dialog
2. Execute scaling according to selected mode (temporarily grouping if necessary)
3. Apply alignment or area-matching options
4. Confirm with OK, or revert with Cancel or Reset

### Update History

- v1.0.0 (20250601): Initial version
- v1.1.0 (20250601): Supported "One Side Only" with "Fixed Size"
- v1.2.0 (20250601): Bug fixes and improvements
*/

(function() {
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var lang = getCurrentLang();

    var labels = {
        keepAspect: { ja: "縦横比保持", en: "Keep aspect" },
        oneSideOnly: { ja: "片辺のみ", en: "One side only" },
        max: { ja: "最大：", en: "Max:" },
        min: { ja: "最小：", en: "Min:" },
        fixed: { ja: "指定サイズ：", en: "Fixed Size:" },
        width: { ja: "幅", en: "Width" },
        height: { ja: "高さ", en: "Height" },
        base: { ja: "基準：", en: "Base:" },
        longSide: { ja: "長辺", en: "Long side" },
        shortSide: { ja: "短辺", en: "Short side" },
        area: { ja: "面積：", en: "Area:" },
        areaMax: { ja: "最大", en: "Max" },
        areaMin: { ja: "最小", en: "Min" },
        artboard: { ja: "アートボード：", en: "Artboard:" },
        bleed: { ja: "裁ち落とし：", en: "Bleed" },
        previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
        alignTitle: { ja: "整列", en: "Alignment" },
        alignVerticalLabel: { ja: "縦に並べる", en: "Vertical" },
        alignLeft: { ja: "左", en: "Left" },
        alignCenter: { ja: "中央", en: "Center" },
        alignEven: { ja: "均等", en: "Distribute evenly" },
        alignZero: { ja: "0", en: "0" },
        alignHorizontalLabel: { ja: "横に並べる", en: "Horizontal" },
        alignTop: { ja: "上", en: "Top" },
        alignMiddle: { ja: "中央", en: "Middle" },
        ok: { ja: "OK", en: "OK" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        selectObject: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
        reset: { ja: "リセット", en: "Reset" },
        dialogTitle: { ja: "オブジェクトのリサイズ", en: "SmartObjectResizer" }
    };

    var doc = app.activeDocument;
    var selectedItems = doc.selection;
    if (!selectedItems || selectedItems.length === 0) {
        alert(labels.selectObject[lang]);
        return;
    }

    var originalStates = [];
    for (var i = 0; i < selectedItems.length; i++) {
        var item = selectedItems[i];
        originalStates.push({
            item: item,
            width: item.width,
            height: item.height,
            left: item.left,
            top: item.top,
            // strokeWidth: item.strokeWidth
        });
    }

    var unitLabel;
    switch (doc.rulerUnits) {
        case RulerUnits.Millimeters:
            unitLabel = "mm";
            break;
        case RulerUnits.Centimeters:
            unitLabel = "cm";
            break;
        case RulerUnits.Inches:
            unitLabel = "inch";
            break;
        case RulerUnits.Pixels:
            unitLabel = "px";
            break;
        case RulerUnits.Picas:
            unitLabel = "pica";
            break;
        default:
            unitLabel = "pt";
            break;
    }

    var dialog = new Window("dialog", labels.dialogTitle ? labels.dialogTitle[lang] : "SmartObjectResizer");
    dialog.alignChildren = ["left", "top"];
    dialog.margins = [0, 0, 0, 10];

    // --- 縦横比と片辺オプションのグループ ---
    // 縦横比・片辺選択ラジオボタン（mainGroupより前に配置）
    var ratioGroup = dialog.add("group");
    ratioGroup.orientation = "row";
    ratioGroup.alignChildren = ["left", "center"];
    ratioGroup.margins = [20, 20, 0, 0];
    ratioGroup.alignment = ["center", "top"]; // 中央揃え

    // 「縦横比保持」「片辺のみ」ラジオボタン
    var keepRatioRadio = ratioGroup.add("radiobutton", undefined, labels.keepAspect[lang]);
    var oneSideOnlyRadio = ratioGroup.add("radiobutton", undefined, labels.oneSideOnly[lang]);
    // ラジオボタン間のスペースを明示的に設定
    keepRatioRadio.margins = [0, 0, 20, 0]; // 右側に20px相当の余白
    oneSideOnlyRadio.margins = [0, 0, 0, 0];

    keepRatioRadio.value = true; // デフォルトで「縦横比保持」を選択

    // 「片辺のみ」選択時に即時ディム表示
    oneSideOnlyRadio.onClick = function () {
        keepRatioRadio.value = false;
        this.value = true;

        var baseGroup = baseRadios[0].parent;
        baseGroup.enabled = false;
        var areaGroup = areaRadios[0].parent;
        areaGroup.enabled = false;
        // アートボードラジオグループもディム
        var artboardGroup = artboardRadios[0].parent;
        artboardGroup.enabled = false;
        // 裁ち落としラジオグループもディム
        var bleedGroup = bleedRadios[0].parent;
        bleedGroup.enabled = false;

        // 指定サイズセクション全体（ラジオ＋テキストフィールド）を有効化
        if (createRadioGroup.sizeInput && createRadioGroup.sizeInput.parent) {
            createRadioGroup.sizeInput.parent.enabled = true; // 入力行
        }
        if (createRadioGroup.widthRadio && createRadioGroup.widthRadio.parent) {
            createRadioGroup.widthRadio.parent.enabled = true; // ラジオボタン行
        }

        onAnyRadioClick.call(this);
    };

    // 「縦横比保持」選択時に即時有効化
    keepRatioRadio.onClick = function () {
        oneSideOnlyRadio.value = false;
        this.value = true;

        var baseGroup = baseRadios[0].parent;
        baseGroup.enabled = true;
        var areaGroup = areaRadios[0].parent;
        areaGroup.enabled = true;
        // アートボードラジオグループも有効化
        var artboardGroup = artboardRadios[0].parent;
        artboardGroup.enabled = true;
        // 裁ち落としラジオグループも有効化
        var bleedGroup = bleedRadios[0].parent;
        bleedGroup.enabled = true;

        // 指定サイズセクション全体（ラジオ＋テキストフィールド）を有効化
        if (createRadioGroup.sizeInput && createRadioGroup.sizeInput.parent) {
            createRadioGroup.sizeInput.parent.enabled = true;
        }
        if (createRadioGroup.widthRadio && createRadioGroup.widthRadio.parent) {
            createRadioGroup.widthRadio.parent.enabled = true;
        }

        onAnyRadioClick.call(this);
    };

    var mainGroup = dialog.add("group");
    mainGroup.orientation = "row";
    // 左右ペインを上揃えに
    mainGroup.alignChildren = ["top", "top"];

    // 左ペイン（各種設定）
    var leftPane = mainGroup.add("group");
    leftPane.orientation = "column";
    leftPane.alignChildren = ["left", "top"];
    leftPane.margins = [20, 20, 10, 0];
    leftPane.margins = [leftPane.margins[0], 0, leftPane.margins[2], leftPane.margins[3]];


    // containerPanel を左ペイン内にパネルとして追加
    var containerPanel = leftPane.add("panel", undefined, "基準");
    containerPanel.orientation = "column";
    containerPanel.alignChildren = "left";
    containerPanel.margins = [15, 20, 15, 15];
    containerPanel.alignment = ["fill", "top"];

    var allRadioButtons = [];
    var radioGroups = [];

    function createRadioGroup(label, optionLabels, parent, isFixedSizeGroup) {
        if (!isFixedSizeGroup) {
            var group = parent.add("group");
            group.orientation = "row";
            group.alignChildren = ["left", "center"];
            var labelText = group.add("statictext", undefined, label);
            // 「最小」のラベルに幅指定を追加
            if (label === labels.min[lang]) {
                labelText.preferredSize.width = 80;
            } else {
                labelText.preferredSize.width = 80;
            }
            var buttons = [];
            for (var i = 0; i < optionLabels.length; i++) {
                var btn = group.add("radiobutton", undefined, optionLabels[i]);
                buttons.push(btn);
                allRadioButtons.push(btn);
            }
            return buttons;
        } else {
            // --- 指定サイズ＋数値入力欄を2行構成に分割 ---
            // 1行目：ラベルとラジオボタン
            var sizeHeaderGroup = parent.add("group");
            sizeHeaderGroup.orientation = "row";
            sizeHeaderGroup.alignChildren = ["left", "center"];
            var sizeLabel = sizeHeaderGroup.add("statictext", undefined, label);
            sizeLabel.preferredSize.width = 80;
            var widthRadio = sizeHeaderGroup.add("radiobutton", undefined, optionLabels[0]);
            var heightRadio = sizeHeaderGroup.add("radiobutton", undefined, optionLabels[1]);
            allRadioButtons.push(widthRadio, heightRadio);
            // 2行目：空ラベルと入力欄
            var sizeInputGroup = parent.add("group");
            sizeInputGroup.orientation = "row";
            sizeInputGroup.alignChildren = ["left", "center"];
           var labelSpacer = sizeInputGroup.add("statictext", undefined, ""); // 空ラベル
           labelSpacer.preferredSize.width = 80;
            // 平均幅を初期値に
            var totalWidth = 0;
            for (var i = 0; i < selectedItems.length; i++) {
                var bounds = getReferenceBounds(selectedItems[i], true);
                totalWidth += bounds.width;
            }
            var avgWidth = selectedItems.length > 0 ? (totalWidth / selectedItems.length) : 100;
            var sizeInput = sizeInputGroup.add("edittext", undefined, avgWidth.toFixed(0));
            sizeInput.characters = 5;
            var sizeUnit = sizeInputGroup.add("statictext", undefined, unitLabel);
            // イベント連携
            sizeInput.onChange = function() {
                // 指定サイズグループはラジオボタン2つのどちらか選択時に有効
                if ((widthRadio.value || heightRadio.value) && !isNaN(parseFloat(sizeInput.text))) {
                    restoreOriginalSizes();
                    if (oneSideOnlyRadio.value) {
                        var parsed = parseFloat(sizeInput.text);
                        var factor = 1;
                        if (unitLabel === "mm") factor = 2.83464567;
                        else if (unitLabel === "cm") factor = 28.3464567;
                        else if (unitLabel === "inch") factor = 72;
                        else if (unitLabel === "pica") factor = 12;
                        var referenceValue = parsed * factor;

                        for (var i = 0; i < selectedItems.length; i++) {
                            var bounds = getReferenceBounds(selectedItems[i], true);
                            var currentSide = widthRadio.value ? bounds.width : bounds.height;
                            if (currentSide === 0) continue;
                            var scale = getScaleFactor(currentSide, referenceValue);
                            selectedItems[i].resize(
                                widthRadio.value ? scale : 100,
                                heightRadio.value ? scale : 100
                            );
                        }
                        app.redraw();
                    } else {
                        applyResizeBySelection();
                    }
                }
            };
            // 保存: sizeInput, widthRadio, heightRadio, sizeUnit
            createRadioGroup.sizeInput = sizeInput;
            createRadioGroup.widthRadio = widthRadio;
            createRadioGroup.heightRadio = heightRadio;
            createRadioGroup.sizeUnit = sizeUnit;
            // 返却: [widthRadio, heightRadio]（従来のラジオボタン配列として）
            return [widthRadio, heightRadio];
        }
    }

    // 新しい順序でラジオボタンとグループを作成
    // 1. 最大
    var maxRadios = createRadioGroup(labels.max[lang], [labels.width[lang], labels.height[lang]], containerPanel);
    // 2. 最小
    var minRadios = createRadioGroup(labels.min[lang], [labels.width[lang], labels.height[lang]], containerPanel);
    // 3. 指定サイズ（ラジオ＋数値欄一体）
    var fixedRadios = createRadioGroup(labels.fixed[lang], [labels.width[lang], labels.height[lang]], containerPanel, true);
    // 5. 基準
    var baseRadios = createRadioGroup(labels.base[lang], [labels.longSide[lang], labels.shortSide[lang]], containerPanel);
    // 6. 面積
    var areaRadios = createRadioGroup(labels.area[lang], [labels.areaMax[lang], labels.areaMin[lang]], containerPanel);
    // 7. --- ディバイダー ---
    var dividerLine = containerPanel.add("statictext", undefined, "  ───────────────  ");
    // 8. アートボード
    var artboardRadios = createRadioGroup(labels.artboard[lang], [labels.width[lang], labels.height[lang]], containerPanel);
    // 9. 裁ち落とし
    var bleedRadios = createRadioGroup(labels.bleed[lang], [labels.width[lang], labels.height[lang]], containerPanel);

    // radioGroupsの並び順も新順に
    // [最大, 最小, 指定サイズ, 基準, 面積, アートボード, 裁ち落とし]
    radioGroups.push(maxRadios, minRadios, fixedRadios, baseRadios, areaRadios, artboardRadios, bleedRadios);

    // --- 整列チェック群 ---
    // 整列チェックボックスをすべてOFFにする共通関数
    function resetAlignChecks() {
        alignLeftCheck.value = false;
        alignCenterCheck.value = false;
        alignTopCheck.value = false;
        alignMiddleCheck.value = false;
        alignEvenCheck.value = false;
        alignHorizontalEvenCheck.value = false;
    }

    function updateRadioGroupStates() {
        var baseGroup = baseRadios[0].parent;
        var areaGroup = areaRadios[0].parent;
        var artboardGroup = artboardRadios[0].parent;
        var bleedGroup = bleedRadios[0].parent;
        if (oneSideOnlyRadio.value) {
            baseGroup.enabled = false;
            areaGroup.enabled = false;
            artboardGroup.enabled = false;
            bleedGroup.enabled = false;
        } else {
            baseGroup.enabled = true;
            areaGroup.enabled = true;
            artboardGroup.enabled = true;
            bleedGroup.enabled = true;
        }
    }

    // ラジオボタン選択時の共通処理
    function onAnyRadioClick() {
        for (var i = 0; i < allRadioButtons.length; i++) {
            if (allRadioButtons[i] !== this) {
                allRadioButtons[i].value = false;
            }
        }
        resetAlignChecks();
        updateInputState();
        // 一時グループが存在すれば解除
        var tempGroup = applyResizeBySelection.tempGroup;
        if (tempGroup && tempGroup.pageItems.length > 0) {
            while (tempGroup.pageItems.length > 0) {
                tempGroup.pageItems[0].move(doc, ElementPlacement.PLACEATEND);
            }
            tempGroup.remove();
            selectedItems = doc.selection;
            applyResizeBySelection.tempGroup = null;
        }
        restoreOriginalSizes();
        applyResizeBySelection();
        updateRadioGroupStates();
        // 指定サイズ入力有効化制御
        var sizeInput = createRadioGroup.sizeInput;
        var widthRadio = createRadioGroup.widthRadio;
        var heightRadio = createRadioGroup.heightRadio;
        var sizeGroup = sizeInput && sizeInput.parent;
        var isFixed = false;
        if (widthRadio && widthRadio.value) isFixed = true;
        if (heightRadio && heightRadio.value) isFixed = true;
        var targetIsWidth = widthRadio && widthRadio.value;
        var targetIsHeight = heightRadio && heightRadio.value;
        if (sizeInput && sizeGroup) {
            if (isFixed && !oneSideOnlyRadio.value) {
                sizeInput.enabled = true;
            } else if (isFixed && oneSideOnlyRadio.value && (targetIsWidth || targetIsHeight)) {
                sizeInput.enabled = true;
            } else {
                sizeInput.enabled = false;
            }
        }
        // --- ディバイダーのディム制御 ---
        if (typeof dividerLine !== "undefined" && dividerLine) {
            if (oneSideOnlyRadio.value) {
                dividerLine.enabled = false;
            } else {
                dividerLine.enabled = true;
            }
        }
    }

    for (var i = 0; i < allRadioButtons.length; i++) {
        // keepRatioRadio, oneSideOnlyRadioには個別のonClickを設定済み
        if (allRadioButtons[i] !== keepRatioRadio && allRadioButtons[i] !== oneSideOnlyRadio) {
            allRadioButtons[i].onClick = onAnyRadioClick;
        }
    }

    // 初期値: 最大・幅
    radioGroups[0][0].value = true;

    // 指定サイズ入力欄の有効化制御
    function updateInputState() {
        var widthRadio = createRadioGroup.widthRadio;
        var heightRadio = createRadioGroup.heightRadio;
        var sizeInput = createRadioGroup.sizeInput;
        var sizeGroup = sizeInput && sizeInput.parent;
        var isFixed = false;
        if (widthRadio && widthRadio.value) isFixed = true;
        if (heightRadio && heightRadio.value) isFixed = true;
        var targetIsWidth = widthRadio && widthRadio.value;
        var targetIsHeight = heightRadio && heightRadio.value;
        if (sizeInput && sizeGroup) {
            if (isFixed && !oneSideOnlyRadio.value) {
                sizeInput.enabled = true;
            } else if (isFixed && oneSideOnlyRadio.value && (targetIsWidth || targetIsHeight)) {
                sizeInput.enabled = true;
            } else {
                sizeInput.enabled = false;
            }
        }
    }

    // Preview boundary checkbox in a group with top margin
    var previewGroup = leftPane.add("group");
    previewGroup.margins = [0, 10, 0, 0]; // top margin

    var previewCheck = previewGroup.add("checkbox", undefined, labels.previewBounds[lang]);
    previewCheck.value = true;
    previewCheck.onClick = function() {
        resetAlignChecks(); // ← 整列チェックボックスをすべてOFFに
        restoreOriginalSizes();
        applyResizeBySelection();
    };

    // 右ペイン（ボタン）
    var rightPane = mainGroup.add("group");
    rightPane.orientation = "column";
    rightPane.alignChildren = ["fill", "top"];
    rightPane.margins = [10, 20, 20, 20];
    // 上マージンを0に
    rightPane.margins = [rightPane.margins[0], 0, rightPane.margins[2], rightPane.margins[3]];

    // 整列パネル
    var alignPanel = rightPane.add("panel", undefined, labels.alignTitle[lang]);
    alignPanel.orientation = "column";
    alignPanel.alignChildren = ["fill", "top"];
    // より見た目が整うようマージンを調整
    alignPanel.margins = [10, 20, 10, 10];

    // 横並びラベル（実際は縦並びラベル）
    alignPanel.add("statictext", undefined, labels.alignVerticalLabel[lang]);

    // 「左」「中央」チェックボックスを横並びに配置するグループ
    var alignLeftGroup = alignPanel.add("group");
    alignLeftGroup.orientation = "row";
    alignLeftGroup.alignChildren = ["left", "center"];

    var alignLeftCheck = alignLeftGroup.add("checkbox", undefined, labels.alignLeft[lang]);
    var alignCenterCheck = alignLeftGroup.add("checkbox", undefined, labels.alignCenter[lang]);
    alignLeftCheck.value = false;
    alignCenterCheck.value = false;

    alignLeftCheck.onClick = function() {
        if (alignLeftCheck.value) {
            // 横方向の他チェックをOFF
            alignCenterCheck.value = false;
            alignHorizontalEvenCheck.value = false;
            alignHorizontalEvenZeroCheck.value = false;
        }
        if (alignLeftCheck.value) {
            var minLeft = null;
            for (var i = 0; i < selectedItems.length; i++) {
                var bounds = getReferenceBounds(selectedItems[i], true);
                if (minLeft === null || bounds.left < minLeft) {
                    minLeft = bounds.left;
                }
            }
            if (minLeft !== null) {
                for (var i = 0; i < selectedItems.length; i++) {
                    var item = selectedItems[i];
                    var bounds = getReferenceBounds(item, true);
                    var delta = minLeft - bounds.left;
                    item.left += delta;
                }
            }
            app.redraw();
        } else {
            restoreOriginalSizes();
        }
    };

    alignCenterCheck.onClick = function() {
        if (alignCenterCheck.value) {
            // 横方向の他チェックをOFF
            alignLeftCheck.value = false;
            alignHorizontalEvenCheck.value = false;
            alignHorizontalEvenZeroCheck.value = false;

            var centerSum = 0;
            for (var i = 0; i < selectedItems.length; i++) {
                var bounds = getReferenceBounds(selectedItems[i], true);
                centerSum += bounds.left + bounds.width / 2;
            }
            var averageCenter = centerSum / selectedItems.length;

            for (var j = 0; j < selectedItems.length; j++) {
                var item = selectedItems[j];
                var bounds = getReferenceBounds(item, true);
                var itemCenter = bounds.left + bounds.width / 2;
                var delta = averageCenter - itemCenter;
                item.left += delta;
            }
            app.redraw();
        } else {
            restoreOriginalSizes();
        }
    };

    // 「均等」チェックボックスを新たなグループで配置
    var alignEvenGroup = alignPanel.add("group");
    alignEvenGroup.orientation = "row";
    alignEvenGroup.alignChildren = ["left", "center"];
    var alignEvenCheck = alignEvenGroup.add("checkbox", undefined, labels.alignEven[lang]);
    alignEvenCheck.value = false;
    // 追加: 0間隔チェックボックス（ラベル多言語対応）
    var alignEvenZeroCheck = alignEvenGroup.add("checkbox", undefined, labels.alignZero[lang]);
    alignEvenZeroCheck.value = false;

    alignEvenCheck.onClick = function() {
        if (alignEvenCheck.value) {
            // 縦方向の他チェックをOFF
            alignTopCheck.value = false;
            alignMiddleCheck.value = false;
            alignEvenZeroCheck.value = false;
        }
        if (alignEvenCheck.value && selectedItems.length > 1) {
            // 0 チェックボックスと排他
            alignEvenZeroCheck.value = false;
            // visibleBounds.top を基準に上から順にソート
            selectedItems.sort(function(a, b) {
                var topA = getReferenceBounds(a, true).top;
                var topB = getReferenceBounds(b, true).top;
                return topB - topA;
            });

            var topMost = getReferenceBounds(selectedItems[0], true).top;
            var bottomMost = getReferenceBounds(selectedItems[selectedItems.length - 1], true).top -
                getReferenceBounds(selectedItems[selectedItems.length - 1], true).height;

            var totalHeight = 0;
            for (var i = 0; i < selectedItems.length; i++) {
                totalHeight += getReferenceBounds(selectedItems[i], true).height;
            }

            var gap = (topMost - bottomMost - totalHeight) / (selectedItems.length - 1);
            var currentY = topMost;

            for (var j = 0; j < selectedItems.length; j++) {
                var item = selectedItems[j];
                var bounds = getReferenceBounds(item, true);
                var height = bounds.height;
                item.top = currentY;
                currentY -= (height + gap);
            }

            app.redraw();
        } else if (!alignEvenCheck.value) {
            restoreOriginalSizes();
        }
    };

    // 追加: 0間隔チェックボックスの挙動
    alignEvenZeroCheck.onClick = function() {
        if (alignEvenZeroCheck.value) {
            // 縦方向の他チェックをOFF
            alignTopCheck.value = false;
            alignMiddleCheck.value = false;
            alignEvenCheck.value = false;
        }
        if (alignEvenZeroCheck.value && selectedItems.length > 1) {
            // restoreOriginalSizes(); // ← 削除: ON時はリサイズ後の状態を維持してゼロ間隔配置
            alignEvenCheck.value = false;

            selectedItems.sort(function(a, b) {
                var topA = getReferenceBounds(a, true).top;
                var topB = getReferenceBounds(b, true).top;
                return topB - topA;
            });

            var topMost = getReferenceBounds(selectedItems[0], true).top;

            var currentY = topMost;
            for (var j = 0; j < selectedItems.length; j++) {
                var item = selectedItems[j];
                var bounds = getReferenceBounds(item, true);
                var height = bounds.height;
                item.top = currentY;
                currentY -= height; // gap = 0
            }

            app.redraw();
        }

        // OFFになったら元に戻す
        if (!alignEvenZeroCheck.value) {
            restoreOriginalSizes();
        }
    };

    // ディバイダー（横並びの前、中央に見えるよう調整）
    alignPanel.add("statictext", undefined, "  ─────  ");
    // 縦並びラベル（実際は横並びラベル）
    alignPanel.add("statictext", undefined, labels.alignHorizontalLabel[lang]);

    // 「上」「中央」チェックボックスを横並びに配置するグループ
    var alignTopGroup = alignPanel.add("group");
    alignTopGroup.orientation = "row";
    alignTopGroup.alignChildren = ["left", "center"];

    var alignTopCheck = alignTopGroup.add("checkbox", undefined, labels.alignTop[lang]);
    var alignMiddleCheck = alignTopGroup.add("checkbox", undefined, labels.alignMiddle[lang]);
    alignTopCheck.value = false;
    alignMiddleCheck.value = false;

    alignTopCheck.onClick = function() {
        if (alignTopCheck.value) {
            // 縦方向の他チェックをOFF
            alignMiddleCheck.value = false;
            alignEvenCheck.value = false;
            alignEvenZeroCheck.value = false;
        }
        if (alignTopCheck.value) {
            var maxTop = null;
            for (var i = 0; i < selectedItems.length; i++) {
                var bounds = getReferenceBounds(selectedItems[i], true);
                if (maxTop === null || bounds.top > maxTop) {
                    maxTop = bounds.top;
                }
            }
            if (maxTop !== null) {
                for (var i = 0; i < selectedItems.length; i++) {
                    var item = selectedItems[i];
                    var bounds = getReferenceBounds(item, true);
                    var delta = maxTop - bounds.top;
                    item.top += delta;
                }
            }
            app.redraw();
        } else {
            restoreOriginalSizes();
        }
    };

    alignMiddleCheck.onClick = function() {
        if (alignMiddleCheck.value) {
            // 縦方向の他チェックをOFF
            alignTopCheck.value = false;
            alignEvenCheck.value = false;
            alignEvenZeroCheck.value = false;
        }
        if (alignMiddleCheck.value) {
            var centerSum = 0;
            for (var i = 0; i < selectedItems.length; i++) {
                var bounds = getReferenceBounds(selectedItems[i], true);
                centerSum += bounds.top - bounds.height / 2;
            }
            var averageCenter = centerSum / selectedItems.length;

            for (var j = 0; j < selectedItems.length; j++) {
                var item = selectedItems[j];
                var bounds = getReferenceBounds(item, true);
                var itemCenter = bounds.top - bounds.height / 2;
                var delta = averageCenter - itemCenter;
                item.top += delta;
            }
            app.redraw();
        } else {
            restoreOriginalSizes();
        }
    };

    // 「均等」チェックボックス（横方向）を新たなグループで配置
    var alignHorizontalEvenGroup = alignPanel.add("group");
    alignHorizontalEvenGroup.orientation = "row";
    alignHorizontalEvenGroup.alignChildren = ["left", "center"];
    var alignHorizontalEvenCheck = alignHorizontalEvenGroup.add("checkbox", undefined, labels.alignEven[lang]);
    alignHorizontalEvenCheck.value = false;
    // 「0」チェックボックスを横方向にも追加（ラベル多言語対応）
    var alignHorizontalEvenZeroCheck = alignHorizontalEvenGroup.add("checkbox", undefined, labels.alignZero[lang]);
    alignHorizontalEvenZeroCheck.value = false;

    alignHorizontalEvenCheck.onClick = function() {
        if (alignHorizontalEvenCheck.value) {
            // 横方向の他チェックをOFF
            alignLeftCheck.value = false;
            alignCenterCheck.value = false;
            alignHorizontalEvenZeroCheck.value = false;
        }
        if (alignHorizontalEvenCheck.value && selectedItems.length > 1) {
            alignHorizontalEvenZeroCheck.value = false;

            selectedItems.sort(function(a, b) {
                var leftA = getReferenceBounds(a, true).left;
                var leftB = getReferenceBounds(b, true).left;
                return leftA - leftB;
            });

            var leftMost = getReferenceBounds(selectedItems[0], true).left;
            var rightMost = getReferenceBounds(selectedItems[selectedItems.length - 1], true).left +
                getReferenceBounds(selectedItems[selectedItems.length - 1], true).width;

            var totalWidth = 0;
            for (var i = 0; i < selectedItems.length; i++) {
                totalWidth += getReferenceBounds(selectedItems[i], true).width;
            }

            var gap = (rightMost - leftMost - totalWidth) / (selectedItems.length - 1);
            var currentX = leftMost;

            for (var j = 0; j < selectedItems.length; j++) {
                var item = selectedItems[j];
                var bounds = getReferenceBounds(item, true);
                var width = bounds.width;
                item.left = currentX;
                currentX += (width + gap);
            }

            app.redraw();
        } else if (!alignHorizontalEvenCheck.value) {
            restoreOriginalSizes();
        }
    };

    // 追加: 0間隔チェックボックスの挙動（横方向）
    alignHorizontalEvenZeroCheck.onClick = function() {
        if (alignHorizontalEvenZeroCheck.value) {
            // 横方向の他チェックをOFF
            alignLeftCheck.value = false;
            alignCenterCheck.value = false;
            alignHorizontalEvenCheck.value = false;
        }
        if (alignHorizontalEvenZeroCheck.value && selectedItems.length > 1) {
            // restoreOriginalSizes(); // ← 削除: ON時はリサイズ後の状態を維持してゼロ間隔配置
            alignHorizontalEvenCheck.value = false;

            selectedItems.sort(function(a, b) {
                var leftA = getReferenceBounds(a, true).left;
                var leftB = getReferenceBounds(b, true).left;
                return leftA - leftB;
            });

            var leftMost = getReferenceBounds(selectedItems[0], true).left;

            var currentX = leftMost;
            for (var j = 0; j < selectedItems.length; j++) {
                var item = selectedItems[j];
                var bounds = getReferenceBounds(item, true);
                var width = bounds.width;
                item.left = currentX;
                currentX += width; // gap = 0
            }

            app.redraw();
        }

        // OFFになったら元に戻す
        if (!alignHorizontalEvenZeroCheck.value) {
            restoreOriginalSizes();
        }
    };


    // スペーサー
    var spacer = rightPane.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.height = 10;

    // --- リセットボタン追加 ---
    // リセットボタンはキャンセルボタンの直前に配置
    var resetButton = rightPane.add("button", undefined, labels.reset ? labels.reset[lang] : "Reset");
    resetButton.onClick = function() {
        restoreOriginalSizes();

        // --- 一時グループ解除（明示的なリセット時） ---
        var tempGroup = applyResizeBySelection.tempGroup;
        if (tempGroup && tempGroup.pageItems.length > 0) {
            while (tempGroup.pageItems.length > 0) {
                tempGroup.pageItems[0].move(doc, ElementPlacement.PLACEATEND);
            }
            tempGroup.remove();
            selectedItems = doc.selection;
            applyResizeBySelection.tempGroup = null;
        }
    };

    // ボタン類はスペーサーの後に配置
    var cancelButton = rightPane.add("button", undefined, labels.cancel[lang], {
        name: "cancel"
    });
    cancelButton.onClick = function() {
        restoreOriginalSizes();
        dialog.close();
    };

    var okButton = rightPane.add("button", undefined, labels.ok[lang], {
        name: "ok"
    });

    // --- 一時グループ解除: OKボタン押下時 ---
    okButton.onClick = function() {
        // 一時グループを解除（アートボード／裁ち落とし時）
        var tempGroup = applyResizeBySelection.tempGroup;
        if (tempGroup && tempGroup.pageItems.length > 0) {
            // 保持順序のため逆順でPLACEATBEGINNING
            var items = [];
            while (tempGroup.pageItems.length > 0) {
                items.push(tempGroup.pageItems[0]);
                tempGroup.pageItems[0].move(doc, ElementPlacement.PLACEATBEGINNING);
            }
            // 逆順に配置して元の重ね順を保持
            for (var i = items.length - 1; i >= 0; i--) {
                items[i].move(doc, ElementPlacement.PLACEATBEGINNING);
            }
            tempGroup.remove();
        }
        applyResizeBySelection.tempGroup = null;
        dialog.close();
    };

    dialog.show();


    function getScaleFactor(current, target) {
        return (target / current) * 100;
    }

    // --- リサイズ適用 ---
    function applyResizeBySelection() {
        var option = null;
        for (var g = 0; g < radioGroups.length; g++) {
            if (radioGroups[g][0].value) {
                option = [g, 0];
                break;
            }
            if (radioGroups[g][1].value) {
                option = [g, 1];
                break;
            }
        }
        if (!option) return;

        var isArtboard = option[0] === 5;
        var isFixed = option[0] === 2;
        var isBleed = option[0] === 6;

        if (typeof applyResizeBySelection.tempGroup === "undefined") {
            applyResizeBySelection.tempGroup = null;
        }
        var tempGroup = null;
        tempGroup = applyResizeBySelection.tempGroup;
        if (tempGroup && tempGroup.pageItems.length > 0) {
            while (tempGroup.pageItems.length > 0) {
                tempGroup.pageItems[0].move(doc, ElementPlacement.PLACEATEND);
            }
            tempGroup.remove();
            selectedItems = doc.selection;
        }
        applyResizeBySelection.tempGroup = null;
        if ((isArtboard || isBleed) && selectedItems.length > 1) {
            restoreOriginalSizes();
            tempGroup = doc.groupItems.add();
            // 重ね順を維持するため先に順番を記録
            var itemsInOrder = [];
            for (var i = 0; i < selectedItems.length; i++) {
                itemsInOrder.push(selectedItems[i]);
            }
            // 重ね順を保持するため、逆順で moveToBeginning を使う
            for (var j = itemsInOrder.length - 1; j >= 0; j--) {
                itemsInOrder[j].moveToBeginning(tempGroup);
            }
            selectedItems = [tempGroup];
            applyResizeBySelection.tempGroup = tempGroup;
        }

        var referenceValue = null;
        var isWidth = option[1] === 0;
        var targetIsWidth = isWidth;
        var isMax = option[0] === 0;
        var isMin = option[0] === 1;
        var isLong = option[0] === 3 && option[1] === 0;
        var isShort = option[0] === 3 && option[1] === 1;
        var isArea = option[0] === 4;
        var isAreaMax = isArea && option[1] === 0;
        var isAreaMin = isArea && option[1] === 1;

        if (isArea) {
            var areas = [];
            for (var i = 0; i < selectedItems.length; i++) {
                var b = getReferenceBounds(selectedItems[i]);
                areas.push(b.width * b.height);
            }
            var baseArea = isAreaMax ?
                Math.max.apply(null, areas) :
                Math.min.apply(null, areas);
            if (baseArea === null || baseArea <= 0) return;
            for (var i = 0; i < selectedItems.length; i++) {
                resizeToMatchArea(selectedItems[i], baseArea);
            }
            app.redraw();
            return;
        }

        if (isFixed) {
            var sizeInput = createRadioGroup.sizeInput;
            var parsed = parseFloat(sizeInput.text);
            if (isNaN(parsed) || parsed <= 0) return;
            var factor = 1;
            if (unitLabel === "mm") factor = 2.83464567;
            else if (unitLabel === "cm") factor = 28.3464567;
            else if (unitLabel === "inch") factor = 72;
            else if (unitLabel === "pica") factor = 12;
            referenceValue = parsed * factor;
        } else if (isArtboard) {
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            referenceValue = isWidth ? (ab[2] - ab[0]) : (ab[1] - ab[3]);
        } else if (isBleed) {
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            var abWidth = ab[2] - ab[0];
            var abHeight = ab[1] - ab[3];
            var bleedOffset = 6 * 2.83464567;
            if (option[1] === 0) {
                referenceValue = abWidth + bleedOffset;
            } else if (option[1] === 1) {
                referenceValue = abHeight + bleedOffset;
            }
        } else {
            for (var i = 0; i < selectedItems.length; i++) {
                var bounds = getReferenceBounds(selectedItems[i]);
                var value;
                if (isLong) {
                    value = Math.max(bounds.width, bounds.height);
                } else if (isShort) {
                    value = Math.min(bounds.width, bounds.height);
                } else {
                    value = isWidth ? bounds.width : bounds.height;
                }
                if (isMin && value === 0) continue;
                if (referenceValue === null) referenceValue = value;
                else if (isMax && value > referenceValue) referenceValue = value;
                else if (isMin && value < referenceValue) referenceValue = value;
            }
        }

        if (referenceValue === null || referenceValue <= 0) return;

        var keepOneSideOnly = oneSideOnlyRadio.value;
        for (var i = 0; i < selectedItems.length; i++) {
            var bounds = getReferenceBounds(selectedItems[i]);
            var current;
            if (isLong) {
                current = Math.max(bounds.width, bounds.height);
            } else if (isShort) {
                current = Math.min(bounds.width, bounds.height);
            } else {
                current = isWidth ? bounds.width : bounds.height;
            }
            if (current === 0) continue;
            var scale = getScaleFactor(current, referenceValue);

            if (keepOneSideOnly) {
                if (isFixed) {
                    // referenceValueは既にテキストフィールドから取得済み
                    var currentSide = targetIsWidth ? bounds.width : bounds.height;
                    if (currentSide === 0) continue;
                    var scaleFixed = getScaleFactor(currentSide, referenceValue);
                    selectedItems[i].resize(
                        targetIsWidth ? scaleFixed : 100,
                        targetIsHeight ? scaleFixed : 100
                    );
                } else {
                    if (targetIsWidth) {
                        selectedItems[i].resize(scale, 100);
                    } else {
                        selectedItems[i].resize(100, scale);
                    }
                }
            } else {
                selectedItems[i].resize(scale, scale, true, true, true, true, scale, Transformation.TOPLEFT);
            }
        }

        if (isArtboard || isBleed) {
            var ab = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            var centerX, centerY;
            if (isBleed) {
                var bleedOffset = 6 * 2.83464567;
                var bleedRect = [
                    ab[0] - bleedOffset / 2,
                    ab[1] + bleedOffset / 2,
                    ab[2] + bleedOffset / 2,
                    ab[3] - bleedOffset / 2
                ];
                centerX = (bleedRect[0] + bleedRect[2]) / 2;
                centerY = (bleedRect[1] + bleedRect[3]) / 2;
            } else {
                centerX = (ab[0] + ab[2]) / 2;
                centerY = (ab[1] + ab[3]) / 2;
            }
            var groupBounds = {
                left: null,
                top: null,
                right: null,
                bottom: null
            };
            for (var i = 0; i < selectedItems.length; i++) {
                var b = getReferenceBounds(selectedItems[i]);
                if (groupBounds.left === null || b.left < groupBounds.left) groupBounds.left = b.left;
                if (groupBounds.top === null || b.top > groupBounds.top) groupBounds.top = b.top;
                if (groupBounds.right === null || (b.left + b.width) > groupBounds.right) groupBounds.right = b.left + b.width;
                if (groupBounds.bottom === null || (b.top - b.height) < groupBounds.bottom) groupBounds.bottom = b.top - b.height;
            }
            var groupWidth = groupBounds.right - groupBounds.left;
            var groupHeight = groupBounds.top - groupBounds.bottom;
            var groupCenterX = groupBounds.left + groupWidth / 2;
            var groupCenterY = groupBounds.top - groupHeight / 2;
            var deltaX = centerX - groupCenterX;
            var deltaY = centerY - groupCenterY;
            for (var i = 0; i < selectedItems.length; i++) {
                selectedItems[i].left += deltaX;
                selectedItems[i].top += deltaY;
            }
            doc.selection = selectedItems;
        }
        app.redraw();
    }

    function resizeToMatchArea(item, targetArea) {
        var bounds = getReferenceBounds(item);
        var w = bounds.width;
        var h = bounds.height;
        var area = w * h;
        if (area === 0) return;
        var scale = Math.sqrt(targetArea / area) * 100;
        item.resize(scale, scale, true, true, true, true, scale, Transformation.TOPLEFT);
    }

    function restoreOriginalSizes() {
        for (var i = 0; i < originalStates.length; i++) {
            var item = originalStates[i].item;
            var state = originalStates[i];
            var scaleW = (state.width === 0) ? 100 : (state.width / item.width) * 100;
            var scaleH = (state.height === 0) ? 100 : (state.height / item.height) * 100;
            item.resize(scaleW, scaleH, true, true, true, true, scaleW, Transformation.TOPLEFT);
            item.left = state.left;
            item.top = state.top;
        }
        app.redraw();
    }

    function getReferenceBounds(item, forceVisible) {
        var boundsArray;
        if (forceVisible || (previewCheck && previewCheck.value)) {
            boundsArray = item.visibleBounds;
        } else {
            boundsArray = item.geometricBounds;
        }
        var width = boundsArray[2] - boundsArray[0];
        var height = boundsArray[1] - boundsArray[3];
        return {
            width: width,
            height: height,
            left: boundsArray[0],
            top: boundsArray[1]
        };
    }

    function drawReferenceBackgrounds(x, y, width, height) {
        var cellRect = doc.activeLayer.pathItems.rectangle(y, x, width, height);
    }
})();