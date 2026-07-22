#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択したオブジェクトの線幅と矢印（始点／終点の形状・倍率・先端位置）をまとめて設定します。
- 矢印は Illustrator の DOM から操作できないため、一時アクション（ai_plugin_setStroke）を生成して実行します。
- ダイアログでプレビューできます。線幅は DOM で即時反映、矢印を含む設定はアクション実行＋取り消しで反映します。

### 処理の流れ

1. ドキュメントと選択オブジェクトの有無を確認
2. ダイアログで線幅・矢印・先端位置を入力（プレビュー可）
3. 入力値から .aia（アクション）ソースを生成
4. 一時ファイルとして書き出し → 読み込み → 実行 → 破棄

### 注意

- 矢印名・先端位置名は Illustrator の UI 表示名と一致している必要があります（言語に依存）。
- 矢印の倍率キー（asc1 / asc2）は推定値です。

---

### Overview

- Sets stroke width and arrowheads (shape, scale and tip alignment for both ends) of the selection at once.
- Arrowheads are not exposed to the Illustrator DOM, so a temporary action (ai_plugin_setStroke) is generated and played.
- The dialog supports preview: stroke width updates through the DOM instantly, arrowhead settings are previewed by playing the action and undoing it.

### Flow

1. Verify that a document is open and objects are selected
2. Enter stroke width, arrowheads and tip alignment in the dialog (with preview)
3. Build an .aia (action) source from the entered values
4. Write it as a temporary file, load, play, then discard it

### Notes

- Arrowhead and tip alignment names must match the Illustrator UI labels (language dependent).
- The arrowhead scale keys (asc1 / asc2) are estimated values.

*/

(function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================

    var SCRIPT_NAME     = "SetStrokeAndArrowheads";       /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
    var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
    var SCRIPT_RELEASED = "2026-07-22";                   /* 最初のリリース日 / first release date */
    var SCRIPT_UPDATED  = "2026-07-22";                   /* 更新日 / last updated */

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 一時アクション / Temporary action */
    var ACTION_SET_NAME  = "SwwwitchTempStrokeSet";
    var ACTION_NAME      = "SwwwitchTempStroke";
    var ACTION_FILE_NAME = File(Folder.temp).fsName + "/swwwitch_temp_stroke.aia";

    /* 既定値 / Defaults */
    var DEFAULT_STROKE_WIDTH = 5;            /* 線幅の初期値 / initial stroke width */
    var DEFAULT_ARROW_SCALE  = 100;          /* 倍率の初期値 / initial arrowhead scale */
    var ARROW_COUNT          = 39;           /* 矢印の種類数 / number of arrowhead presets */

    /* パラメータキー / Parameter keys（記録した .aia から採取 / taken from a recorded .aia） */
    var KEY_STROKE_WIDTH  = 2003072104;      /* 線幅 / stroke width */
    var KEY_CAP           = 1667330094;      /* 線端 / cap */
    var KEY_JOIN          = 1785686382;      /* 角の形状 / join */
    var KEY_DASH_INT      = 1684825454;      /* 破線（整数）/ dash (integer) */
    var KEY_DASH_BOOL     = 1684104298;      /* 破線（真偽）/ dash (boolean) */
    var KEY_ARROW_HEAD_1  = 1634231345;      /* ahd1: 始点の形状 / start arrowhead */
    var KEY_ARROW_HEAD_2  = 1634231346;      /* ahd2: 終点の形状 / end arrowhead */
    var KEY_ARROW_SCALE_1 = 1634951985;      /* asc1: 始点の倍率（推定値）/ start scale (estimated) */
    var KEY_ARROW_SCALE_2 = 1634951986;      /* asc2: 終点の倍率 / end scale */
    var KEY_ARROW_ALIGN   = 1634230636;      /* ahal: 矢印の配置 / tip alignment */
    var KEY_ALIGN         = 1634494318;      /* algn: 線の位置 / stroke alignment */

    /* UIレイアウト：余白と間隔 / UI layout: margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /* UIレイアウト：コントロールの寸法 / UI layout: control metrics */
    var LABEL_WIDTH      = 40;               /* 行頭ラベルの幅 / row label width */
    var SHAPE_LIST_WIDTH = 90;               /* 形状プルダウンの幅 / shape dropdown width */
    var FIELD_CHARACTERS = 4;                /* 数値入力欄の文字数 / numeric field width */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 実行環境の言語を判定 / Detect the runtime language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "線と矢印を設定", en: "Set Stroke and Arrowheads" }
        },
        panel: {
            stroke: { ja: "線", en: "Stroke" },
            arrow: { ja: "矢印", en: "Arrowheads" },
            start: { ja: "始点", en: "Start" },
            end: { ja: "終点", en: "End" },
            tipAlign: { ja: "先端位置", en: "Tip Alignment" }
        },
        field: {
            strokeWidth: { ja: "線幅：", en: "Weight:" },
            shape: { ja: "形状：", en: "Shape:" },
            scale: { ja: "倍率：", en: "Scale:" },
            unitPt: { ja: "pt", en: "pt" },
            unitPercent: { ja: "%", en: "%" }
        },
        checkbox: {
            linkEnds: { ja: "始点と終点を連動", en: "Link start and end" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        arrow: {
            none: { ja: "[なし]", en: "[None]" },
            prefix: { ja: "矢印 ", en: "Arrow " }
        },
        /* 先端位置：label は UI 表示用、name はアクションに埋め込む Illustrator の表示名 */
        /* Tip alignment: label is for the UI, name is the Illustrator label embedded in the action */
        tipAlign: {
            atEndLabel: {
                ja: "矢印の先端をパスの終点に配置",
                en: "Place arrow tip at end of path"
            },
            atEndName: {
                ja: "パスの終点に配置",
                en: "Place Arrow Tip At End of Path"
            },
            beyondEndLabel: {
                ja: "矢印の先端をパスの終点から配置",
                en: "Extend arrow tip beyond end of path"
            },
            beyondEndName: {
                ja: "パスの終点から配置",
                en: "Extend Arrow Tip Beyond End of Path"
            }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントを開いてください。",
                en: "Please open a document."
            },
            noSelection: {
                ja: "オブジェクトを選択してください。",
                en: "Please select at least one object."
            },
            invalidWidth: {
                ja: "線幅には 0 以上の数値を入力してください。",
                en: "Enter a stroke weight of 0 or greater."
            },
            actionFailed: {
                ja: "一時アクションファイルを開けませんでした。",
                en: "Failed to open the temporary action file."
            }
        }
    };

    /* ドット区切りのキーからローカライズ文字列を取得 / Resolve a dotted key to a localized string */
    function L(keyPath) {
        var parts = String(keyPath).split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (!node) return keyPath;
            node = node[parts[i]];
        }
        if (!node) return keyPath;
        return (node[currentLanguage] !== undefined) ? node[currentLanguage] : node.en;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    var UNIT_POINT = 592476268;              /* ポイント / point（parameter /unit） */

    /* 先端位置の選択肢 / Tip alignment options（ahal の enumerated 値 / enumerated values） */
    var ARROW_ALIGN_OPTIONS = [
        { label: L("tipAlign.atEndLabel"),     name: L("tipAlign.atEndName"),     value: 0 },
        { label: L("tipAlign.beyondEndLabel"), name: L("tipAlign.beyondEndName"), value: 1 }
    ];

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* タブの共通設定 / Apply shared tab layout */
    function setupTab(tab, spacing) {
        tab.orientation = "column";
        tab.alignChildren = "fill";
        tab.margins = TAB_MARGINS;
        if (typeof spacing === "number") tab.spacing = spacing;
    }

    /* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
    function trimButtonHeight(button, px) {
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) {}
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* 線を持てるオブジェクトを再帰的に集める / Collect strokable items recursively */
    function collectStrokeTargets(items) {
        var targets = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            try {
                if (item.typename === "PathItem") {
                    targets.push(item);
                } else if (item.typename === "CompoundPathItem") {
                    targets = targets.concat(collectStrokeTargets(item.pathItems));
                } else if (item.typename === "GroupItem") {
                    targets = targets.concat(collectStrokeTargets(item.pageItems));
                }
            } catch (e) {}
        }
        return targets;
    }

    /*
     * プレビュー制御 / Preview controller
     * - 線幅は DOM で即時プレビュー（元の値を保存して書き戻すので undo に依存しない）
     * - 矢印は DOM に存在しないため、アクション実行 + app.undo() でしかプレビューできない
     * - Stroke width previews through the DOM; arrowheads require play + undo
     */
    function createPreviewController() {
        var targets = [];
        var originalStroked = [];
        var originalWidths = [];
        var isActionPreviewApplied = false;

        /* 現在の線設定を退避する / Capture the current stroke state */
        function capture() {
            targets = collectStrokeTargets(app.activeDocument.selection);
            originalStroked = [];
            originalWidths = [];
            for (var i = 0; i < targets.length; i++) {
                try {
                    originalStroked.push(targets[i].stroked);
                    originalWidths.push(targets[i].stroked ? targets[i].strokeWidth : 1);
                } catch (e) {
                    originalStroked.push(false);
                    originalWidths.push(1);
                }
            }
        }

        /* 退避した線設定を書き戻す / Restore the captured stroke state */
        function restoreStroke() {
            for (var i = 0; i < targets.length; i++) {
                try {
                    targets[i].stroked = originalStroked[i];
                    if (originalStroked[i]) targets[i].strokeWidth = originalWidths[i];
                } catch (e) {}
            }
        }

        /* アクションによるプレビューを undo で取り消す / Undo the action based preview */
        function clearActionPreview() {
            if (!isActionPreviewApplied) return;
            try { app.undo(); } catch (e) {}
            isActionPreviewApplied = false;
            /* undo でオブジェクト参照が無効になるため取り直す / References die on undo, re-capture */
            capture();
        }

        capture();

        return {
            /* 線幅のみ DOM で即時プレビュー / Preview stroke width only, via the DOM */
            previewStrokeWidth: function (strokeWidth) {
                clearActionPreview();
                if (isNaN(strokeWidth) || strokeWidth < 0) return;
                for (var i = 0; i < targets.length; i++) {
                    try {
                        targets[i].stroked = true;
                        targets[i].strokeWidth = strokeWidth;
                    } catch (e) {}
                }
                app.redraw();
            },

            /* 矢印を含む全設定をアクションで 1 回だけプレビュー / Preview all settings via the action */
            previewAction: function (settings) {
                clearActionPreview();
                restoreStroke();
                var source = buildStrokeActionSource(ACTION_SET_NAME, ACTION_NAME, settings);
                playTemporaryAction(source, ACTION_SET_NAME, ACTION_NAME, ACTION_FILE_NAME);
                isActionPreviewApplied = true;
                app.redraw();
            },

            /* アクションによるプレビューだけを取り消す / Clear only the action based preview */
            clearAction: function () {
                if (!isActionPreviewApplied) return;
                clearActionPreview();
                app.redraw();
            },

            /* すべてのプレビューを取り消して元の状態に戻す / Revert everything */
            reset: function () {
                clearActionPreview();
                restoreStroke();
                app.redraw();
            }
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /*
     * ↑↓キーで値を増減する / Step a value with the arrow keys
     * - ↑↓ で ±1 / plus or minus 1
     * - shift 併用で ±10（10 の倍数にスナップ）/ plus or minus 10, snapped to multiples of 10
     * - option 併用で ±0.1 / plus or minus 0.1
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 */
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                /* Optionキー押下時は0.1単位で増減 / Step by 0.1 */
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                /* 小数第1位までに丸め / Round to one decimal place */
                value = Math.round(value * 10) / 10;
            } else {
                /* 整数に丸め / Round to an integer */
                value = Math.round(value);
            }

            editText.text = value;

            /* プログラムからの書き換えではイベントが発火しないため明示的に呼ぶ */
            /* Programmatic changes do not fire the event, so call it explicitly */
            if (typeof editText.onChanging === "function") editText.onChanging();
        });
    }

    /* 矢印名の一覧を作る / Build the arrowhead name list */
    function buildArrowNameList() {
        var names = [L("arrow.none")];
        for (var i = 1; i <= ARROW_COUNT; i++) {
            names.push(L("arrow.prefix") + i);
        }
        return names;
    }

    /*
     * 設定用ダイアログを表示する / Show the settings dialog
     * @returns {object|null} 入力値。キャンセル時は null / entered values, or null when cancelled
     */
    function showSettingsDialog() {
        var arrowNames = buildArrowNameList();
        var preview = createPreviewController();

        var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(dialog);

        /* 線幅 / Stroke width */
        var strokePanel = dialog.add("panel", undefined, L("panel.stroke"));
        setupPanel(strokePanel);

        var strokeRow = strokePanel.add("group");
        setupRow(strokeRow);
        strokeRow.add("statictext", undefined, L("field.strokeWidth"));
        var strokeWidthInput = strokeRow.add("edittext", undefined, String(DEFAULT_STROKE_WIDTH));
        strokeWidthInput.characters = FIELD_CHARACTERS;
        strokeRow.add("statictext", undefined, L("field.unitPt"));

        /* 矢印 / Arrowheads */
        var arrowPanel = dialog.add("panel", undefined, L("panel.arrow"));
        setupPanel(arrowPanel);

        var linkCheckbox = arrowPanel.add("checkbox", undefined, L("checkbox.linkEnds"));
        linkCheckbox.value = false;
        linkCheckbox.alignment = "left";

        /* 始点／終点を 2 カラムで並べる / Lay out start and end in two columns */
        var arrowColumns = arrowPanel.add("group");
        arrowColumns.orientation = "row";
        arrowColumns.alignChildren = ["fill", "top"];
        arrowColumns.spacing = COLUMN_SPACING;

        /* 形状・倍率をひと組にしたカラムを作る / Build a shape + scale column */
        function addArrowColumn(parent, title) {
            var panel = parent.add("panel", undefined, title);
            setupPanel(panel);

            var shapeRow = panel.add("group");
            setupRow(shapeRow);
            var shapeLabel = shapeRow.add("statictext", undefined, L("field.shape"));
            shapeLabel.preferredSize.width = LABEL_WIDTH;
            var shapeList = shapeRow.add("dropdownlist", undefined, arrowNames);
            shapeList.selection = 0;
            shapeList.preferredSize.width = SHAPE_LIST_WIDTH;

            var scaleRow = panel.add("group");
            setupRow(scaleRow);
            var scaleLabel = scaleRow.add("statictext", undefined, L("field.scale"));
            scaleLabel.preferredSize.width = LABEL_WIDTH;
            var scaleInput = scaleRow.add("edittext", undefined, String(DEFAULT_ARROW_SCALE));
            scaleInput.characters = FIELD_CHARACTERS;
            scaleRow.add("statictext", undefined, L("field.unitPercent"));

            return { shapeList: shapeList, scaleInput: scaleInput };
        }

        var startColumn = addArrowColumn(arrowColumns, L("panel.start"));
        var endColumn   = addArrowColumn(arrowColumns, L("panel.end"));

        var startShapeList  = startColumn.shapeList;
        var startScaleInput = startColumn.scaleInput;
        var endShapeList    = endColumn.shapeList;
        var endScaleInput   = endColumn.scaleInput;

        /* 先端位置 / Tip alignment */
        var alignPanel = dialog.add("panel", undefined, L("panel.tipAlign"));
        setupPanel(alignPanel, 6);

        var alignRadios = [];
        for (var a = 0; a < ARROW_ALIGN_OPTIONS.length; a++) {
            var radio = alignPanel.add("radiobutton", undefined, ARROW_ALIGN_OPTIONS[a].label);
            /* ボタン類はパネル幅いっぱいに広げない / Keep buttons at their natural width */
            radio.alignment = "left";
            alignRadios.push(radio);
        }
        alignRadios[0].value = true;

        /* ボタン（左：プレビュー／中央：スペーサー／右：キャンセル・OK） */
        /* Buttons (left: preview, center: spacer, right: cancel and OK) */
        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignChildren = ["fill", "center"];

        var buttonLeft = buttonGroup.add("group");
        buttonLeft.alignment = ["left", "center"];
        var previewCheckbox = buttonLeft.add("checkbox", undefined, L("checkbox.preview"));
        previewCheckbox.value = false;

        var buttonCenter = buttonGroup.add("group");
        buttonCenter.alignment = ["fill", "center"];

        var buttonRight = buttonGroup.add("group");
        buttonRight.alignment = ["right", "center"];
        buttonRight.add("button", undefined, L("button.cancel"), { name: "cancel" });
        buttonRight.add("button", undefined, "OK", { name: "ok" });

        /* 連動時は始点の値を終点にコピー / Mirror start values onto end when linked */
        function applyLink() {
            if (!linkCheckbox.value) return;
            endShapeList.selection = startShapeList.selection.index;
            endScaleInput.text = startScaleInput.text;
        }

        /* 形状が [なし] のときは倍率を無効化。連動時は終点側を無効化 */
        /* Disable scale when the shape is [None]; disable the end column when linked */
        function syncEnabled() {
            var isLinked = linkCheckbox.value;
            var hasArrow = (startShapeList.selection.index !== 0) || (endShapeList.selection.index !== 0);

            startScaleInput.enabled = (startShapeList.selection.index !== 0);
            endShapeList.enabled    = !isLinked;
            endScaleInput.enabled   = !isLinked && (endShapeList.selection.index !== 0);

            /* 矢印がひとつも無ければ先端位置は無効 / No arrowheads means no tip alignment */
            alignPanel.enabled = hasArrow;
        }

        /* 選択中の先端位置を返す / Return the selected tip alignment */
        function getSelectedAlign() {
            for (var i = 0; i < alignRadios.length; i++) {
                if (alignRadios[i].value) return ARROW_ALIGN_OPTIONS[i];
            }
            return ARROW_ALIGN_OPTIONS[0];
        }

        /* 連動と有効／無効をまとめて更新 / Refresh linking and enabled states */
        function refresh() {
            applyLink();
            syncEnabled();
        }

        /* 入力値を読み取る。線幅が不正なら null / Read input; null when the width is invalid */
        function readSettings() {
            var strokeWidth = parseFloat(strokeWidthInput.text);
            var startScale  = parseFloat(startScaleInput.text);
            var endScale    = parseFloat(endScaleInput.text);

            if (isNaN(strokeWidth) || strokeWidth < 0) return null;
            if (isNaN(startScale) || startScale <= 0) startScale = DEFAULT_ARROW_SCALE;
            if (isNaN(endScale)   || endScale   <= 0) endScale   = DEFAULT_ARROW_SCALE;

            return {
                strokeWidth: strokeWidth,
                startArrow: startShapeList.selection.text,
                startScale: startScale,
                endArrow: endShapeList.selection.text,
                endScale: endScale,
                arrowAlign: getSelectedAlign()
            };
        }

        /* プレビューを現在の入力値で更新する（矢印を含むためアクションを実行） */
        /* Refresh the preview with the current values (plays the action for arrowheads) */
        function updatePreview() {
            if (!previewCheckbox.value) {
                preview.reset();
                return;
            }
            var settings = readSettings();
            if (!settings) return;
            preview.previewAction(settings);
        }

        /* 値を変更したらアクションのプレビューを取り消す（確定時に貼り直す） */
        /* Drop the action preview while editing; it is reapplied on commit */
        function invalidatePreview() {
            if (previewCheckbox.value) preview.clearAction();
        }

        previewCheckbox.onClick = updatePreview;

        linkCheckbox.onClick = function () {
            refresh();
            updatePreview();
        };
        startShapeList.onChange = function () {
            refresh();
            updatePreview();
        };
        endShapeList.onChange = function () {
            syncEnabled();
            updatePreview();
        };

        startScaleInput.onChanging = function () {
            if (linkCheckbox.value) endScaleInput.text = startScaleInput.text;
            invalidatePreview();
        };
        startScaleInput.onChange = updatePreview;
        endScaleInput.onChanging = invalidatePreview;
        endScaleInput.onChange = updatePreview;

        /* 線幅は入力中も DOM で即時プレビュー、確定時にアクションで貼り直す */
        /* Stroke width previews live through the DOM, then via the action on commit */
        strokeWidthInput.onChanging = function () {
            if (previewCheckbox.value) preview.previewStrokeWidth(parseFloat(strokeWidthInput.text));
        };
        strokeWidthInput.onChange = updatePreview;

        for (var r = 0; r < alignRadios.length; r++) {
            alignRadios[r].onClick = updatePreview;
        }

        /* ↑↓キーで値を増減 / Enable arrow key stepping */
        changeValueByArrowKey(strokeWidthInput);
        changeValueByArrowKey(startScaleInput);
        changeValueByArrowKey(endScaleInput);

        refresh();

        var isAccepted = (dialog.show() === 1);

        /* プレビューを必ず取り消してから本適用に進む / Always revert the preview before applying */
        preview.reset();

        if (!isAccepted) return null;

        var settings = readSettings();
        if (!settings) {
            alert(L("alert.invalidWidth"));
            return null;
        }
        return settings;
    }

    // =========================================
    // 一時アクション生成 / Temporary action generation
    // =========================================

    /*
     * 線幅・矢印を設定する一時アクションのソースを生成する
     * Build the source of a temporary action that sets stroke width and arrowheads
     * @param {string} setName    アクションセット名（ASCII 推奨）/ action set name
     * @param {string} actionName アクション名（ASCII 推奨）/ action name
     * @param {object} settings   showSettingsDialog() の戻り値 / result of showSettingsDialog()
     */
    function buildStrokeActionSource(setName, actionName, settings) {
        return ''
            + '/version 3\n'
            + buildNameLine(setName)
            + '/isOpen 1\n'
            + '/actionCount 1\n'
            + '/action-1 {\n'
            + '\t' + buildNameLine(actionName)
            + '\t/keyIndex 0\n'
            + '\t/colorIndex 0\n'
            + '\t/isOpen 1\n'
            + '\t/eventCount 1\n'
            + '\t/event-1 {\n'
            + '\t\t/useRulersIn1stQuadrant 0\n'
            + '\t\t/internalName (ai_plugin_setStroke)\n'
            + '\t\t/localizedName [ 0 \n\t\t]\n'
            + '\t\t/isOpen 1\n'
            + '\t\t/isOn 1\n'
            + '\t\t/hasDialog 0\n'
            + '\t\t/parameterCount 11\n'
            /* 可変 / variable */
            + buildUnitRealParam(1, KEY_STROKE_WIDTH, settings.strokeWidth, UNIT_POINT)
            + buildUStrParam(6, KEY_ARROW_HEAD_1, settings.startArrow)
            + buildUStrParam(7, KEY_ARROW_HEAD_2, settings.endArrow)
            + buildRealParam(8, KEY_ARROW_SCALE_1, settings.startScale)
            + buildRealParam(9, KEY_ARROW_SCALE_2, settings.endScale)
            + buildEnumParamByName(10, KEY_ARROW_ALIGN, settings.arrowAlign.name, settings.arrowAlign.value) /* 先端位置 / tip alignment */
            /* 以下は記録した .aia のまま / recorded as-is */
            + buildEnumParam(2, KEY_CAP,  'e4b8b8e59e8be7b79ae7abaf', 12, 1)             /* 線端: 丸型線端 / round cap */
            + buildEnumParam(3, KEY_JOIN, 'e383a9e382a6e383b3e38389e7b590e59088', 18, 1) /* 角の形状: ラウンド結合 / round join */
            + buildIntParam(4,  KEY_DASH_INT,  0)
            + buildBoolParam(5, KEY_DASH_BOOL, 0)
            + buildEnumParam(11, KEY_ALIGN, 'e4b8ade5a4ae', 6, 0)                        /* 線の位置: 中央 / center */
            + '\t}\n'
            + '}\n';
    }

    // --- パラメータ組み立てヘルパー / parameter builders ---

    /* パラメータブロックの外枠を作る / Wrap a parameter body in its block */
    function buildParamBlock(index, key, body) {
        return ''
            + '\t\t/parameter-' + index + ' {\n'
            + '\t\t\t/key ' + key + '\n'
            + '\t\t\t/showInPalette 4294967295\n'
            + body
            + '\t\t}\n';
    }

    /* 単位付き実数パラメータ / A unit real parameter */
    function buildUnitRealParam(index, key, value, unitCode) {
        return buildParamBlock(index, key, ''
            + '\t\t\t/type (unit real)\n'
            + '\t\t\t/value ' + toRealString(value) + '\n'
            + '\t\t\t/unit ' + unitCode + '\n');
    }

    /* 列挙パラメータ（hex 指定）/ An enumerated parameter from a hex name */
    function buildEnumParam(index, key, nameHex, byteLength, value) {
        return buildParamBlock(index, key, ''
            + '\t\t\t/type (enumerated)\n'
            + '\t\t\t/name [ ' + byteLength + ' \n\t\t\t\t' + nameHex + '\n\t\t\t]\n'
            + '\t\t\t/value ' + value + '\n');
    }

    /* enumerated パラメータを表示名から組み立てる（hex は自動生成）/ Enumerated parameter from a display name */
    function buildEnumParamByName(index, key, name, value) {
        var hex = stringToUtf8Hex(name);
        return buildEnumParam(index, key, hex, hex.length / 2, value);
    }

    /* 整数パラメータ / An integer parameter */
    function buildIntParam(index, key, value) {
        return buildParamBlock(index, key, '\t\t\t/type (integer)\n\t\t\t/value ' + value + '\n');
    }

    /* 真偽値パラメータ / A boolean parameter */
    function buildBoolParam(index, key, value) {
        return buildParamBlock(index, key, '\t\t\t/type (boolean)\n\t\t\t/value ' + (value ? 1 : 0) + '\n');
    }

    /* 実数パラメータ / A real parameter */
    function buildRealParam(index, key, value) {
        return buildParamBlock(index, key, '\t\t\t/type (real)\n\t\t\t/value ' + toRealString(value) + '\n');
    }

    /* Unicode 文字列パラメータ / A ustring parameter */
    function buildUStrParam(index, key, value) {
        var hex = stringToUtf8Hex(value);
        return buildParamBlock(index, key, ''
            + '\t\t\t/type (ustring)\n'
            + '\t\t\t/value [ ' + (hex.length / 2) + ' \n\t\t\t\t' + hex + '\n\t\t\t]\n');
    }

    // --- 文字列・数値ユーティリティ / utilities ---

    /* アクションセット名・アクション名の /name 行を作る / Build a /name line */
    function buildNameLine(name) {
        var hex = stringToUtf8Hex(name);
        return '/name [ ' + (hex.length / 2) + ' \n\t' + hex + '\n]\n';
    }

    /* UTF-8 バイト列の 16 進表現にする（長さもバイト数で数える）/ Encode a string as UTF-8 hex */
    function stringToUtf8Hex(sourceText) {
        var hexText = "";
        for (var i = 0; i < sourceText.length; i++) {
            var code = sourceText.charCodeAt(i);
            var bytes;
            if (code < 0x80) {
                bytes = [code];
            } else if (code < 0x800) {
                bytes = [0xC0 | (code >> 6), 0x80 | (code & 0x3F)];
            } else {
                bytes = [0xE0 | (code >> 12), 0x80 | ((code >> 6) & 0x3F), 0x80 | (code & 0x3F)];
            }
            for (var j = 0; j < bytes.length; j++) {
                var h = bytes[j].toString(16);
                if (h.length < 2) h = "0" + h;
                hexText += h;
            }
        }
        return hexText;
    }

    /* 5 → "5.0" のように必ず小数点を含む文字列にする / Force a decimal point */
    function toRealString(value) {
        var text = String(Number(value));
        if (text.indexOf('.') === -1 && text.indexOf('e') === -1) text += '.0';
        return text;
    }

    // =========================================
    // 一時アクション実行 / Temporary action playback
    // =========================================

    /* アクションを書き出して読み込み、実行後に破棄する / Write, load, play, then discard the action */
    function playTemporaryAction(actionSource, setName, actionName, actionFilePath) {
        var actionFile = new File(actionFilePath);
        var isActionLoaded = false;
        var isActionFileOpen = false;

        try { app.unloadAction(setName, ""); } catch (e) {}

        try {
            actionFile.encoding = "BINARY";
            if (!actionFile.open("w")) {
                throw new Error(L("alert.actionFailed"));
            }
            isActionFileOpen = true;

            actionFile.write(actionSource);
            actionFile.close();
            isActionFileOpen = false;

            app.loadAction(actionFile);
            isActionLoaded = true;

            app.doScript(actionName, setName, false);

        } finally {
            if (isActionFileOpen) {
                try { actionFile.close(); } catch (e) {}
            }
            if (actionFile.exists) {
                try { actionFile.remove(); } catch (e) {}
            }
            if (isActionLoaded) {
                try { app.unloadAction(setName, ""); } catch (e) {}
            }
        }
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    /* 前提条件を確認し、ダイアログの入力をアクションとして適用する */
    /* Verify preconditions, then apply the dialog input through the action */
    function main() {
        if (app.documents.length === 0) {
            alert(L("alert.noDocument"));
            return;
        }
        if (app.activeDocument.selection.length === 0) {
            alert(L("alert.noSelection"));
            return;
        }

        var settings = showSettingsDialog();
        if (!settings) return;

        var source = buildStrokeActionSource(ACTION_SET_NAME, ACTION_NAME, settings);
        playTemporaryAction(source, ACTION_SET_NAME, ACTION_NAME, ACTION_FILE_NAME);
    }

    main();

})();
