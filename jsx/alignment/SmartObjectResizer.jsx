#targetengine "TMK_SOR_Engine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

SmartObjectResizer.jsx

### 概要

- 選択中のオブジェクトを指定した基準（最大、最小、指定サイズ、アートボード、裁ち落とし、面積）に基づいて柔軟にリサイズできるIllustrator用スクリプトです。
- 縦横比保持モードや片辺のみモード、整列オプション、リアルタイムプレビュー機能を備えています。整列処理は visibleBounds ではなく geometricBounds ベースで計算し、テキスト整列の安定性を高めています。

### 主な機能

- 縦横比保持と片辺のみの切り替え
- 各種基準（最大、最小、指定サイズ、アートボード、裁ち落とし、面積）でのスケーリング
- テキストをアウトライン化したサイズで計算（複製→分割→アウトライン→計測→即削除）
- 整列オプション（左、中央、均等、0間隔、上、中央、横均等、横0間隔）
- リアルタイムプレビューとリセット機能
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログでリサイズモードと基準を選択
2. 選択モードに従ってスケーリング実行（テキストは複製→分割→アウトライン化→計測し、元オブジェクトに反映）
3. 整列オプションや面積一致処理を適用
4. OKで確定、キャンセルやリセットで元に戻す

### 紹介記事（note）

https://note.com/dtp_tranist/n/n6f35bd4000ec

### 更新履歴

- v1.0.0 (20250601) : 初期バージョン
- v1.1.0 (20250601) : 「片辺のみ」＋「指定サイズ」での変形サポート
- v1.2.0 (20250601) : バグ修正、細部改善
- v1.2.1 (20260227) : ダイアログ初期表示時に、選択中モードを1回適用（初期状態でもリサイズが反映されるように）
- v1.2.2 (20260227) : 指定サイズの数値欄で、↑↓/Shift+↑↓/Option+↑↓ による増減を追加（±1 / ±10 / ±0.1）
- v1.2.3 (20260227) : ダイアログ位置をセッション内で記憶し、次回起動時に復元（Illustrator終了でリセット）
- v1.3.0 (20260322) : テキストのアウトライン基準計測を再設計（差分検出＋即時削除＋キャッシュ化で残骸問題を解消、幾何変化対応のキャッシュ最適化）
- v1.3.1 (20260322) : テキスト計測用の複製オブジェクトに、アウトライン化前の分割（expandStyle）を追加
- v1.3.2 (20260322) : 整列処理のみ visibleBounds 依存をやめ、geometricBounds ベースに分離してテキスト整列精度を改善

----

### Script Name:

SmartObjectResizer.jsx

### Overview

- An Illustrator script that flexibly resizes selected objects based on criteria such as Max, Min, Fixed Size, Artboard, Bleed, or Area.
- Supports "Keep Aspect" and "One Side Only" modes, alignment options, and real-time preview. Alignment calculations use geometricBounds instead of visibleBounds for more stable text alignment.

### Main Features

- Toggle between "Keep Aspect" and "One Side Only"
- Scaling based on Max, Min, Fixed Size, Artboard, Bleed, or Area
- Optional text measurement using outlined bounds (duplicate → expand appearance → outline → measure → immediate cleanup)
- Alignment options (Left, Center, Evenly, Zero Gap, Top, Middle, Horizontal Evenly, Horizontal Zero Gap)
- Real-time preview and reset function
- Japanese and English UI support

### Process Flow

1. Select resize mode and base in the dialog
2. Execute scaling according to selected mode (text objects are measured via duplicate → expand appearance → outline → bounds calculation and applied back to originals)
3. Apply alignment or area-matching options
4. Confirm with OK, or revert with Cancel or Reset

### Update History

- v1.0.0 (20250601): Initial version
- v1.1.0 (20250601): Supported "One Side Only" with "Fixed Size"
- v1.2.0 (20250601): Bug fixes and improvements
- v1.2.1 (20260227): Apply the selected mode once on dialog show (so the default selection is applied immediately)
- v1.2.2 (20260227): Added arrow-key value stepping for the Fixed Size input (±1 / ±10 with Shift / ±0.1 with Option)
- v1.2.3 (20260227): Remember dialog position within the session and restore next run (reset when Illustrator quits)
- v1.3.0 (20260322): Reworked text outline measurement using diff-based cleanup and caching (eliminates leftover outline artifacts and optimizes cache invalidation for geometry changes)
- v1.3.1 (20260322): Added expand appearance (expandStyle) before outlining duplicated text objects used for measurement
- v1.3.2 (20260322): Separated alignment calculations from visibleBounds and switched them to geometricBounds for more stable text alignment
*/

(function () {
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var lang = getCurrentLang();

    // スクリプトバージョン / Script version
    var SCRIPT_VERSION = "v1.3.2";

    // ダイアログ位置をセッション内で記憶（Illustrator終了でリセット）
    // Remember dialog position within this Illustrator session (resets when Illustrator quits)
    var __SOR_SESSION_KEY = "SmartObjectResizer_dialogPos";
    if (typeof $.global[__SOR_SESSION_KEY] === "undefined") {
        $.global[__SOR_SESSION_KEY] = null;
    }

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
        textOutlineBounds: { ja: "テキスト（アウトライン化して計算）", en: "Text (outline for bounds)" },
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
    var originalSelectedItems = doc.selection;
    var workingItems = originalSelectedItems;
    if (!originalSelectedItems || originalSelectedItems.length === 0) {
        alert(labels.selectObject[lang]);
        return;
    }

    var originalStates = [];
    for (var i = 0; i < originalSelectedItems.length; i++) {
        var item = originalSelectedItems[i];
        originalStates.push({
            item: item,
            width: item.width,
            height: item.height,
            left: item.left,
            top: item.top,
            // strokeWidth: item.strokeWidth
        });
    }
    var resizeBaseStates = [];

    var outlineBoundsCache = {};
    var outlineBoundsCacheSeq = 1;
    var outlineIdMap = [];

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

    var dialog = new Window("dialog", (labels.dialogTitle ? labels.dialogTitle[lang] : "SmartObjectResizer") + " " + SCRIPT_VERSION);
    dialog.alignChildren = ["left", "top"];
    dialog.margins = [0, 0, 0, 10];

    // 前回のダイアログ位置を復元（セッション内のみ）
    try {
        var _pos = $.global[__SOR_SESSION_KEY];
        if (_pos && _pos.length === 2 && !isNaN(_pos[0]) && !isNaN(_pos[1])) {
            dialog.location = _pos;
        }
    } catch (_) { }

    // --- 初期表示時に、選択中のモードを1回だけ適用 ---
    // ※ 初期状態で「最大：幅」などが選択されていても、クリックするまで反映されない問題の対策
    dialog.onShow = function () {
        try {
            // 整列は初期OFF、状態を元に戻してから適用
            resetAlignChecks();
            clearOutlineBoundsCache();
            restoreOriginalGeometry();
            restoreOriginalPosition();
            app.redraw();

            // UI状態に応じた有効/無効も更新
            updateInputState();
            updateRadioGroupStates();

            updateTextOutlineOptionState();

            // 現在選択されているラジオのモードを1回だけ適用
            applyResizeBySelection();
        } catch (e) {
            // 何もしない（ダイアログ表示は継続）
        }
    };

    // 閉じるときに位置を記憶（セッション内のみ）
    dialog.onClose = function () {
        try {
            $.global[__SOR_SESSION_KEY] = [dialog.location[0], dialog.location[1]];
        } catch (_) { }
        return true;
    };

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

    /**
     * ↑↓キーで数値を増減する / Change numeric value by arrow keys
     * - Up/Down: ±1
     * - Shift + Up/Down: ±10 (snap to tens)
     * - Option(Alt) + Up/Down: ±0.1
     * @param {EditText} editText
     * @param {Boolean} allowNegative
     */
    function changeValueByArrowKey(editText, allowNegative) {
        if (!editText) return;
        if (typeof allowNegative === "undefined") allowNegative = false;

        editText.addEventListener("keydown", function (event) {
            // ScriptUI keyName: "Up" / "Down"
            if (!(event && (event.keyName === "Up" || event.keyName === "Down"))) return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                // Shiftキー押下時は10の倍数にスナップ
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else {
                    value = Math.floor((value - 1) / delta) * delta;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            } else {
                delta = 1;
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            }

            // 丸め / Rounding
            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10; // 小数第1位まで
            } else {
                value = Math.round(value); // 整数
            }

            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;

            // onChange を明示的に呼ぶ（矢印キーでは onChange が発火しないことがある）
            try {
                if (typeof editText.onChange === "function") editText.onChange();
            } catch (_) { }
        });
    }

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
            for (var i = 0; i < originalSelectedItems.length; i++) {
                var bounds = getReferenceBounds(originalSelectedItems[i], true);
                totalWidth += bounds.width;
            }
            var avgWidth = originalSelectedItems.length > 0 ? (totalWidth / originalSelectedItems.length) : 100;
            var sizeInput = sizeInputGroup.add("edittext", undefined, avgWidth.toFixed(0));
            sizeInput.characters = 5;
            changeValueByArrowKey(sizeInput, false);
            var sizeUnit = sizeInputGroup.add("statictext", undefined, unitLabel);
            // イベント連携
            sizeInput.onChange = function () {
                if ((widthRadio.value || heightRadio.value) && !isNaN(parseFloat(sizeInput.text))) {
                    clearOutlineBoundsCache();
                    restoreOriginalGeometry();
                    restoreOriginalPosition();
                    app.redraw();

                    if (oneSideOnlyRadio.value) {
                        var parsed = parseFloat(sizeInput.text);
                        var factor = 1;
                        if (unitLabel === "mm") factor = 2.83464567;
                        else if (unitLabel === "cm") factor = 28.3464567;
                        else if (unitLabel === "inch") factor = 72;
                        else if (unitLabel === "pica") factor = 12;
                        var referenceValue = parsed * factor;

                        for (var i = 0; i < workingItems.length; i++) {
                            var bounds = getReferenceBounds(workingItems[i], true);
                            var currentSide = widthRadio.value ? bounds.width : bounds.height;
                            if (currentSide === 0) continue;
                            var scale = getScaleFactor(currentSide, referenceValue);
                            workingItems[i].resize(
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
        clearOutlineBoundsCache();
        updateInputState();
        // 一時グループが存在すれば解除
        releaseTempGroup(ElementPlacement.PLACEATEND);
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
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

    // Preview-related checkboxes
    var previewGroup = leftPane.add("group");
    previewGroup.orientation = "column";
    previewGroup.alignChildren = ["left", "top"];
    previewGroup.margins = [0, 10, 0, 0]; // top margin

    var textOutlineBoundsCheck = previewGroup.add("checkbox", undefined, labels.textOutlineBounds[lang]);
    textOutlineBoundsCheck.value = false;
    textOutlineBoundsCheck.onClick = function () {
        resetAlignChecks();
        clearOutlineBoundsCache();
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
        applyResizeBySelection();
    };

    function updateTextOutlineOptionState() {
        var hasText = false;

        for (var i = 0; i < originalSelectedItems.length; i++) {
            if (containsTextForOutlineBounds(originalSelectedItems[i])) {
                hasText = true;
                break;
            }
        }

        if (hasText) {
            textOutlineBoundsCheck.enabled = true;
            textOutlineBoundsCheck.value = true;
        } else {
            textOutlineBoundsCheck.value = false;
            textOutlineBoundsCheck.enabled = false;
        }
    }

    var previewCheck = previewGroup.add("checkbox", undefined, labels.previewBounds[lang]);
    previewCheck.value = true;
    previewCheck.onClick = function () {
        resetAlignChecks(); // ← 整列チェックボックスをすべてOFFに
        clearOutlineBoundsCache();
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
        applyResizeBySelection();
    };

    updateTextOutlineOptionState();

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

    alignLeftCheck.onClick = function () {
        if (alignLeftCheck.value) {
            // 横方向の他チェックをOFF
            alignCenterCheck.value = false;
            alignHorizontalEvenCheck.value = false;
            alignHorizontalEvenZeroCheck.value = false;
            restoreResizeBaseState();
        }
        if (alignLeftCheck.value) {
            var minLeft = null;
            for (var i = 0; i < workingItems.length; i++) {
                var bounds = getAlignmentBounds(workingItems[i]);
                if (minLeft === null || bounds.left < minLeft) {
                    minLeft = bounds.left;
                }
            }
            if (minLeft !== null) {
                for (var i = 0; i < workingItems.length; i++) {
                    var item = workingItems[i];
                    var bounds = getAlignmentBounds(item);
                    var delta = minLeft - bounds.left;
                    item.left += delta;
                }
            }
            app.redraw();
        } else {
            restoreResizeBaseState();
        }
    };

    alignCenterCheck.onClick = function () {
        if (alignCenterCheck.value) {
            // 横方向の他チェックをOFF
            alignLeftCheck.value = false;
            alignHorizontalEvenCheck.value = false;
            alignHorizontalEvenZeroCheck.value = false;
            restoreResizeBaseState();

            var centerSum = 0;
            for (var i = 0; i < workingItems.length; i++) {
                var bounds = getAlignmentBounds(workingItems[i]);
                centerSum += bounds.left + bounds.width / 2;
            }
            var averageCenter = centerSum / workingItems.length;

            for (var j = 0; j < workingItems.length; j++) {
                var item = workingItems[j];
                var bounds = getAlignmentBounds(item);
                var itemCenter = bounds.left + bounds.width / 2;
                var delta = averageCenter - itemCenter;
                item.left += delta;
            }
            app.redraw();
        } else {
            restoreResizeBaseState();
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

    alignEvenCheck.onClick = function () {
        if (alignEvenCheck.value) {
            // 縦方向の他チェックをOFF
            alignTopCheck.value = false;
            alignMiddleCheck.value = false;
            alignEvenZeroCheck.value = false;
            restoreResizeBaseState();
        }
        if (alignEvenCheck.value && workingItems.length > 1) {
            // 0 チェックボックスと排他
            alignEvenZeroCheck.value = false;
            // geometricBounds.top を基準に上から順にソート
            var sortedItems = workingItems.slice(0).sort(function (a, b) {
                var topA = getAlignmentBounds(a).top;
                var topB = getAlignmentBounds(b).top;
                return topB - topA;
            });

            var topMost = getAlignmentBounds(sortedItems[0]).top;
            var bottomMost = getAlignmentBounds(sortedItems[sortedItems.length - 1]).top -
                getAlignmentBounds(sortedItems[sortedItems.length - 1]).height;

            var totalHeight = 0;
            for (var i = 0; i < sortedItems.length; i++) {
                totalHeight += getAlignmentBounds(sortedItems[i]).height;
            }

            var gap = (topMost - bottomMost - totalHeight) / (sortedItems.length - 1);
            var currentY = topMost;

            for (var j = 0; j < sortedItems.length; j++) {
                var item = sortedItems[j];
                var bounds = getAlignmentBounds(item);
                var height = bounds.height;
                item.top = currentY;
                currentY -= (height + gap);
            }

            app.redraw();
        } else if (!alignEvenCheck.value) {
            restoreResizeBaseState();
        }
    };

    // 追加: 0間隔チェックボックスの挙動
    alignEvenZeroCheck.onClick = function () {
        if (alignEvenZeroCheck.value) {
            // 縦方向の他チェックをOFF
            alignTopCheck.value = false;
            alignMiddleCheck.value = false;
            alignEvenCheck.value = false;
            restoreResizeBaseState();
        }
        if (alignEvenZeroCheck.value && workingItems.length > 1) {
            alignEvenCheck.value = false;

            var sortedItems = workingItems.slice(0).sort(function (a, b) {
                var topA = getAlignmentBounds(a).top;
                var topB = getAlignmentBounds(b).top;
                return topB - topA;
            });

            var topMost = getAlignmentBounds(sortedItems[0]).top;

            var currentY = topMost;
            for (var j = 0; j < sortedItems.length; j++) {
                var item = sortedItems[j];
                var bounds = getAlignmentBounds(item);
                var height = bounds.height;
                item.top = currentY;
                currentY -= height; // gap = 0
            }

            app.redraw();
        }

        // OFFになったらリサイズ後の基準状態に戻す
        if (!alignEvenZeroCheck.value) {
            restoreResizeBaseState();
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

    alignTopCheck.onClick = function () {
        if (alignTopCheck.value) {
            // 縦方向の他チェックをOFF
            alignMiddleCheck.value = false;
            alignEvenCheck.value = false;
            alignEvenZeroCheck.value = false;
            restoreResizeBaseState();
        }
        if (alignTopCheck.value) {
            var maxTop = null;
            for (var i = 0; i < workingItems.length; i++) {
                var bounds = getAlignmentBounds(workingItems[i]);
                if (maxTop === null || bounds.top > maxTop) {
                    maxTop = bounds.top;
                }
            }
            if (maxTop !== null) {
                for (var i = 0; i < workingItems.length; i++) {
                    var item = workingItems[i];
                    var bounds = getAlignmentBounds(item);
                    var delta = maxTop - bounds.top;
                    item.top += delta;
                }
            }
            app.redraw();
        } else {
            restoreResizeBaseState();
        }
    };

    alignMiddleCheck.onClick = function () {
        if (alignMiddleCheck.value) {
            // 縦方向の他チェックをOFF
            alignTopCheck.value = false;
            alignEvenCheck.value = false;
            alignEvenZeroCheck.value = false;
            restoreResizeBaseState();
        }
        if (alignMiddleCheck.value) {
            var centerSum = 0;
            for (var i = 0; i < workingItems.length; i++) {
                var bounds = getAlignmentBounds(workingItems[i]);
                centerSum += bounds.top - bounds.height / 2;
            }
            var averageCenter = centerSum / workingItems.length;

            for (var j = 0; j < workingItems.length; j++) {
                var item = workingItems[j];
                var bounds = getAlignmentBounds(item);
                var itemCenter = bounds.top - bounds.height / 2;
                var delta = averageCenter - itemCenter;
                item.top += delta;
            }
            app.redraw();
        } else {
            restoreResizeBaseState();
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

    alignHorizontalEvenCheck.onClick = function () {
        if (alignHorizontalEvenCheck.value) {
            // 横方向の他チェックをOFF
            alignLeftCheck.value = false;
            alignCenterCheck.value = false;
            alignHorizontalEvenZeroCheck.value = false;
            restoreResizeBaseState();
        }
        if (alignHorizontalEvenCheck.value && workingItems.length > 1) {
            alignHorizontalEvenZeroCheck.value = false;

            var sortedItems = workingItems.slice(0).sort(function (a, b) {
                var leftA = getAlignmentBounds(a).left;
                var leftB = getAlignmentBounds(b).left;
                return leftA - leftB;
            });

            var leftMost = getAlignmentBounds(sortedItems[0]).left;
            var rightMost = getAlignmentBounds(sortedItems[sortedItems.length - 1]).left +
                getAlignmentBounds(sortedItems[sortedItems.length - 1]).width;

            var totalWidth = 0;
            for (var i = 0; i < sortedItems.length; i++) {
                totalWidth += getAlignmentBounds(sortedItems[i]).width;
            }

            var gap = (rightMost - leftMost - totalWidth) / (sortedItems.length - 1);
            var currentX = leftMost;

            for (var j = 0; j < sortedItems.length; j++) {
                var item = sortedItems[j];
                var bounds = getAlignmentBounds(item);
                var width = bounds.width;
                item.left = currentX;
                currentX += (width + gap);
            }

            app.redraw();
        } else if (!alignHorizontalEvenCheck.value) {
            restoreResizeBaseState();
        }
    };

    // 追加: 0間隔チェックボックスの挙動（横方向）
    alignHorizontalEvenZeroCheck.onClick = function () {
        if (alignHorizontalEvenZeroCheck.value) {
            // 横方向の他チェックをOFF
            alignLeftCheck.value = false;
            alignCenterCheck.value = false;
            alignHorizontalEvenCheck.value = false;
            restoreResizeBaseState();
        }
        if (alignHorizontalEvenZeroCheck.value && workingItems.length > 1) {
            alignHorizontalEvenCheck.value = false;

            var sortedItems = workingItems.slice(0).sort(function (a, b) {
                var leftA = getAlignmentBounds(a).left;
                var leftB = getAlignmentBounds(b).left;
                return leftA - leftB;
            });

            var leftMost = getAlignmentBounds(sortedItems[0]).left;

            var currentX = leftMost;
            for (var j = 0; j < sortedItems.length; j++) {
                var item = sortedItems[j];
                var bounds = getAlignmentBounds(item);
                var width = bounds.width;
                item.left = currentX;
                currentX += width; // gap = 0
            }

            app.redraw();
        }

        // OFFになったらリサイズ後の基準状態に戻す
        if (!alignHorizontalEvenZeroCheck.value) {
            restoreResizeBaseState();
        }
    };

    // スペーサー
    var spacer = rightPane.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.height = 10;

    // --- リセットボタン追加 ---
    // リセットボタンはキャンセルボタンの直前に配置
    var resetButton = rightPane.add("button", undefined, labels.reset ? labels.reset[lang] : "Reset");
    resetButton.onClick = function () {
        // --- 一時グループ解除（明示的なリセット時） ---
        releaseTempGroup(ElementPlacement.PLACEATEND);
        clearOutlineBoundsCache();
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
    };

    // ボタン類はスペーサーの後に配置
    var cancelButton = rightPane.add("button", undefined, labels.cancel[lang], {
        name: "cancel"
    });
    cancelButton.onClick = function () {
        releaseTempGroup(ElementPlacement.PLACEATEND);
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
        clearOutlineBoundsCache();
        dialog.close();
    };

    var okButton = rightPane.add("button", undefined, labels.ok[lang], {
        name: "ok"
    });

    // --- 一時グループ解除: OKボタン押下時 ---
    okButton.onClick = function () {
        // 一時グループを解除（アートボード／裁ち落とし時）
        var tempGroup = getTempGroup();
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
        setTempGroup(null);
        clearOutlineBoundsCache();
        dialog.close();
    };

    dialog.show();

    function getTempGroup() {
        return applyResizeBySelection.tempGroup || null;
    }

    function setTempGroup(group) {
        applyResizeBySelection.tempGroup = group || null;
        workingItems = group ? [group] : originalSelectedItems;
    }

    function releaseTempGroup(placeMode) {
        var tempGroup = getTempGroup();
        if (!tempGroup || tempGroup.pageItems.length === 0) {
            setTempGroup(null);
            return null;
        }

        while (tempGroup.pageItems.length > 0) {
            tempGroup.pageItems[0].move(doc, placeMode);
        }
        tempGroup.remove();
        setTempGroup(null);
        return null;
    }

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
        clearOutlineBoundsCache();

        var isArtboard = option[0] === 5;
        var isFixed = option[0] === 2;
        var isBleed = option[0] === 6;

        if (typeof applyResizeBySelection.tempGroup === "undefined") {
            applyResizeBySelection.tempGroup = null;
        }
        var tempGroup = getTempGroup();
        if (tempGroup && tempGroup.pageItems.length > 0) {
            releaseTempGroup(ElementPlacement.PLACEATEND);
        }
        if ((isArtboard || isBleed) && originalSelectedItems.length > 1) {
            restoreOriginalGeometry();
            restoreOriginalPosition();
            app.redraw();
            tempGroup = doc.groupItems.add();
            // 重ね順を維持するため先に順番を記録
            var itemsInOrder = [];
            for (var i = 0; i < originalSelectedItems.length; i++) {
                itemsInOrder.push(originalSelectedItems[i]);
            }
            // 重ね順を保持するため、逆順で moveToBeginning を使う
            for (var j = itemsInOrder.length - 1; j >= 0; j--) {
                itemsInOrder[j].moveToBeginning(tempGroup);
            }
            setTempGroup(tempGroup);
        } else {
            setTempGroup(null);
        }

        var referenceValue = null;
        var isWidth = option[1] === 0;
        var targetIsWidth = isWidth;
        var targetIsHeight = !isWidth;
        var isMax = option[0] === 0;
        var isMin = option[0] === 1;
        var isLong = option[0] === 3 && option[1] === 0;
        var isShort = option[0] === 3 && option[1] === 1;
        var isArea = option[0] === 4;
        var isAreaMax = isArea && option[1] === 0;
        var isAreaMin = isArea && option[1] === 1;

        if (isArea) {
            var areas = [];
            for (var i = 0; i < workingItems.length; i++) {
                var b = getReferenceBounds(workingItems[i]);
                areas.push(b.width * b.height);
            }
            var baseArea = isAreaMax ?
                Math.max.apply(null, areas) :
                Math.min.apply(null, areas);
            if (baseArea === null || baseArea <= 0) return;
            for (var i = 0; i < workingItems.length; i++) {
                resizeToMatchArea(workingItems[i], baseArea);
            }
            captureResizeBaseState();
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
            for (var i = 0; i < workingItems.length; i++) {
                var bounds = getReferenceBounds(workingItems[i]);
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
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getReferenceBounds(workingItems[i]);
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
                    workingItems[i].resize(
                        targetIsWidth ? scaleFixed : 100,
                        targetIsHeight ? scaleFixed : 100
                    );
                } else {
                    if (targetIsWidth) {
                        workingItems[i].resize(scale, 100);
                    } else {
                        workingItems[i].resize(100, scale);
                    }
                }
            } else {
                workingItems[i].resize(scale, scale, true, true, true, true, scale, Transformation.TOPLEFT);
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
            for (var i = 0; i < workingItems.length; i++) {
                var b = getReferenceBounds(workingItems[i]);
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
            for (var i = 0; i < workingItems.length; i++) {
                workingItems[i].left += deltaX;
                workingItems[i].top += deltaY;
            }
            doc.selection = workingItems;
        }
        captureResizeBaseState();
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

    function restoreOriginalGeometry() {
        for (var i = 0; i < originalStates.length; i++) {
            var item = originalStates[i].item;
            var state = originalStates[i];
            var scaleW = (state.width === 0) ? 100 : (state.width / item.width) * 100;
            var scaleH = (state.height === 0) ? 100 : (state.height / item.height) * 100;
            item.resize(scaleW, scaleH, true, true, true, true, scaleW, Transformation.TOPLEFT);
        }
    }

    function restoreOriginalPosition() {
        for (var i = 0; i < originalStates.length; i++) {
            var item = originalStates[i].item;
            var state = originalStates[i];
            item.left = state.left;
            item.top = state.top;
        }
    }

    function restoreOriginalSizes() {
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
    }

    function getReferenceBounds(item, forceVisible) {
        var useVisibleBounds = !!(forceVisible || (previewCheck && previewCheck.value));
        if (textOutlineBoundsCheck && textOutlineBoundsCheck.value && containsTextForOutlineBounds(item)) {
            return getOutlinedBoundsCached(item, useVisibleBounds);
        }

        var boundsArray;
        if (useVisibleBounds) {
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

    // 整列処理は visibleBounds ではなく geometricBounds ベースで計算する
    // Alignment calculations use geometricBounds instead of visibleBounds
    function getAlignmentBounds(item) {
        if (textOutlineBoundsCheck && textOutlineBoundsCheck.value && containsTextForOutlineBounds(item)) {
            return getOutlinedBoundsCached(item, false);
        }
        return getPageItemBoundsObject(item, false);
    }

    function containsTextForOutlineBounds(item) {
        if (!item) return false;
        if (item.typename === "TextFrame") return true;
        if (item.typename === "GroupItem") {
            return groupHasTextFrames(item);
        }
        return false;
    }

    function groupHasTextFrames(groupItem) {
        if (!groupItem || !groupItem.pageItems) return false;
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            var child = groupItem.pageItems[i];
            if (child.typename === "TextFrame") return true;
            if (child.typename === "GroupItem" && groupHasTextFrames(child)) return true;
        }
        return false;
    }

    function collectTextFramesInGroup(groupItem, result) {
        if (!groupItem || !groupItem.pageItems) return;
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            var child = groupItem.pageItems[i];
            if (child.typename === "TextFrame") {
                result.push(child);
            } else if (child.typename === "GroupItem") {
                collectTextFramesInGroup(child, result);
            }
        }
    }

    // 幾何変化はキャッシュキー側でも吸収するが、モード切替時は明示的に全消去する
    function clearOutlineBoundsCache() {
        outlineBoundsCache = {};
        outlineIdMap = [];
    }

    function getOutlineCacheKey(item, useVisibleBounds) {
        var id = getOutlineId(item);
        return id + "_" + getOutlineCacheGeometrySignature(item, useVisibleBounds);
    }

    function getOutlineId(item) {
        for (var i = 0; i < outlineIdMap.length; i++) {
            if (outlineIdMap[i].item === item) {
                return outlineIdMap[i].id;
            }
        }
        var newId = "sor_" + (outlineBoundsCacheSeq++);
        outlineIdMap.push({ item: item, id: newId });
        return newId;
    }

    function getOutlineCacheGeometrySignature(item, useVisibleBounds) {
        var boundsArray = useVisibleBounds ? item.visibleBounds : item.geometricBounds;
        return (useVisibleBounds ? "v_" : "g_") +
            roundCacheCoord(boundsArray[0]) + "_" +
            roundCacheCoord(boundsArray[1]) + "_" +
            roundCacheCoord(boundsArray[2]) + "_" +
            roundCacheCoord(boundsArray[3]);
    }

    function roundCacheCoord(value) {
        return Math.round(value * 1000) / 1000;
    }

    function getOutlinedBoundsCached(item, useVisibleBounds) {
        var key = getOutlineCacheKey(item, useVisibleBounds);
        if (outlineBoundsCache.hasOwnProperty(key)) {
            return outlineBoundsCache[key];
        }
        var measured = measureOutlinedBoundsByDuplicate(item, useVisibleBounds);
        outlineBoundsCache[key] = measured;
        return measured;
    }

    function measureOutlinedBoundsByDuplicate(item, useVisibleBounds) {
        var beforeItems = snapshotAllPageItems(doc);
        try {
            var duplicateItem = item.duplicate();

            // テキスト計測用の複製では、アウトライン化の前に分割を適用する
            // For duplicated text used for measurement, apply expand appearance before outlining
            if (containsTextForOutlineBounds(duplicateItem)) {
                var previousSelection = null;
                try {
                    previousSelection = doc.selection;
                } catch (_) { }
                try {
                    doc.selection = null;
                    duplicateItem.selected = true;
                    app.executeMenuCommand('expandStyle');
                    if (doc.selection && doc.selection.length > 0) {
                        duplicateItem = doc.selection[0];
                    }
                } catch (_) {
                } finally {
                    try {
                        doc.selection = previousSelection;
                    } catch (__restoreSelErr) { }
                }
            }

            if (duplicateItem.typename === "TextFrame") {
                duplicateItem = duplicateItem.createOutline();
            } else if (duplicateItem.typename === "GroupItem") {
                outlineTextFramesInGroupDuplicate(duplicateItem);
            }

            var newItems = collectNewPageItems(doc, beforeItems);
            if (newItems.length > 0) {
                return getBoundsFromItems(newItems, useVisibleBounds);
            }
            return getPageItemBoundsObject(duplicateItem, useVisibleBounds);
        } catch (_) {
            return getPageItemBoundsObject(item, useVisibleBounds);
        } finally {
            var createdItems = collectNewPageItems(doc, beforeItems);
            removeItemsSafe(createdItems);
        }
    }

    function outlineTextFramesInGroupDuplicate(groupItem) {
        var textFrames = [];
        collectTextFramesInGroup(groupItem, textFrames);
        for (var i = textFrames.length - 1; i >= 0; i--) {
            if (textFrames[i] && textFrames[i].isValid) {
                textFrames[i].createOutline();
            }
        }
    }

    function getPageItemBoundsObject(item, useVisibleBounds) {
        var boundsArray = useVisibleBounds ? item.visibleBounds : item.geometricBounds;
        return {
            width: boundsArray[2] - boundsArray[0],
            height: boundsArray[1] - boundsArray[3],
            left: boundsArray[0],
            top: boundsArray[1]
        };
    }



    function snapshotAllPageItems(docRef) {
        var items = [];
        for (var i = 0; i < docRef.pageItems.length; i++) {
            items.push(docRef.pageItems[i]);
        }
        return items;
    }

    function collectNewPageItems(docRef, beforeItems) {
        var result = [];
        for (var i = 0; i < docRef.pageItems.length; i++) {
            var item = docRef.pageItems[i];
            if (!containsPageItemRef(beforeItems, item)) {
                result.push(item);
            }
        }
        return result;
    }

    function containsPageItemRef(list, item) {
        for (var i = 0; i < list.length; i++) {
            if (list[i] === item) return true;
        }
        return false;
    }

    function getBoundsFromItems(items, useVisibleBounds) {
        if (!items || items.length === 0) {
            return { width: 0, height: 0, left: 0, top: 0 };
        }

        var left = null;
        var top = null;
        var right = null;
        var bottom = null;

        for (var i = 0; i < items.length; i++) {
            var boundsArray = useVisibleBounds ? items[i].visibleBounds : items[i].geometricBounds;
            if (left === null || boundsArray[0] < left) left = boundsArray[0];
            if (top === null || boundsArray[1] > top) top = boundsArray[1];
            if (right === null || boundsArray[2] > right) right = boundsArray[2];
            if (bottom === null || boundsArray[3] < bottom) bottom = boundsArray[3];
        }

        return {
            width: right - left,
            height: top - bottom,
            left: left,
            top: top
        };
    }

    function removeItemsSafe(items) {
        if (!items || items.length === 0) return;
        for (var i = items.length - 1; i >= 0; i--) {
            try {
                items[i].locked = false;
            } catch (_) { }
            try {
                items[i].hidden = false;
            } catch (_) { }
            try {
                items[i].remove();
            } catch (_) { }
        }
    }


    function drawReferenceBackgrounds(x, y, width, height) {
        var cellRect = doc.activeLayer.pathItems.rectangle(y, x, width, height);
    }

    function captureResizeBaseState() {
        resizeBaseStates = [];
        for (var i = 0; i < workingItems.length; i++) {
            var item = workingItems[i];
            resizeBaseStates.push({
                item: item,
                width: item.width,
                height: item.height,
                left: item.left,
                top: item.top
            });
        }
    }

    function restoreResizeBaseState() {
        if (!resizeBaseStates || resizeBaseStates.length === 0) return;
        for (var i = 0; i < resizeBaseStates.length; i++) {
            var item = resizeBaseStates[i].item;
            var state = resizeBaseStates[i];
            var scaleW = (state.width === 0) ? 100 : (state.width / item.width) * 100;
            var scaleH = (state.height === 0) ? 100 : (state.height / item.height) * 100;
            item.resize(scaleW, scaleH, true, true, true, true, scaleW, Transformation.TOPLEFT);
            item.left = state.left;
            item.top = state.top;
        }
        app.redraw();
    }
})();