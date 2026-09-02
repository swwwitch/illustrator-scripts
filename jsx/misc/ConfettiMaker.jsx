#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクト（未選択のときはアートボード）の領域に、紙吹雪を散らします。
形状・生成数・分布・ランダム量をプレビューで確かめながら調整し、［OK］で Confetti レイヤーへ出力します。

詳細は README を参照してください。

### Overview

Scatters confetti across the selected object, or across the artboard when nothing is selected.
Shape, count, distribution and randomness are tuned against a live preview, and OK commits the result onto a Confetti layer.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ConfettiMaker";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.7.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConfettiMaker.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConfettiMaker.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n5a41fb524a5a"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 生成物のレイヤー名・グループ名 / Names of generated layers and groups */
    var PREVIEW_LAYER_NAME = "__ConfettiPreview__"; /* プレビュー専用レイヤー / preview-only layer */
    var OUTPUT_LAYER_NAME  = "Confetti";            /* 確定時の出力先レイヤー / output layer */
    var MASK_GROUP_NAME    = "__ConfettiMasked__";  /* マスク処理用グループ / group to be clipped */
    var OUTPUT_GROUP_NAME  = "__ConfettiGroup__";   /* マスクOFF時のまとめグループ / group used when masking is off */
    var SYMBOL_WRAP_NAME   = "__ConfettiSymbol__";  /* シンボルを包むグループ / wrapper group for a symbol */

    /* 初期値 / Initial values */
    var DEFAULT_BASE_SIZE_PT = 6;   /* 基準サイズ（pt）/ base size in points */
    var DEFAULT_COUNT        = 150; /* 生成数 / number of confetti */
    var DEFAULT_OPACITY_MIN  = 70;  /* 不透明度の下限（%）/ lower bound of random opacity */
    var DEFAULT_ROTATE_MAX   = 360; /* 最大回転角（度）/ maximum rotation in degrees */
    var DEFAULT_STRENGTH       = 2.0; /* 分布の強度 / distribution strength */

    /* 挙動の調整値 / Behavior tuning */
    var SOLO_RANDOM_SIZE      = 167; /* Option+クリックの単独選択時に設定する大きさランダム量 / size randomness applied on solo click */
    var MARGIN_STEP_ON_ENABLE = 30;  /* マージンをONにしたときに加算する量（pt）/ amount added when margin is switched on */
    var PREVIEW_DEBOUNCE_MS   = 140; /* プレビュー再描画の間引き時間（ms）/ debounce delay for preview redraw */
    var SKEW_MAX_DEG          = 45;  /* 歪みの上限（度）/ upper bound of skew */

    /* カラフルな配色 / Confetti colors */
    var CONFETTI_COLORS = [
        [255, 77, 77],   /* 赤 / red */
        [255, 153, 51],  /* オレンジ / orange */
        [255, 255, 51],  /* 黄色 / yellow */
        [102, 255, 102], /* 緑 / green */
        [51, 153, 255],  /* 青 / blue */
        [153, 102, 255], /* 紫 / purple */
        [255, 102, 178], /* ピンク / pink */
        [255, 204, 51]   /* 金色 / gold */
    ];

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS         = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING         = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS          = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING          = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING         = 12;               /* 2カラムの間隔 / gap between columns */
    var LABEL_WIDTH            = 80;               /* 行ラベルの共通幅 / shared width of row labels */
    var SLIDER_WIDTH           = 240;              /* スライダーの幅 / slider width */
    var NARROW_SLIDER_WIDTH    = 200;              /* 補助スライダーの幅 / width of a secondary slider */
    var SYMBOL_DROPDOWN_WIDTH  = 120;              /* シンボル選択の幅 / width of the symbol dropdown */
    var BUTTON_ROW_TOP_MARGIN  = 10;               /* ボタンエリアの上余白 / top margin of the button row */
    var DIALOG_OFFSET_X        = 300;              /* ダイアログの横方向オフセット / horizontal offset of the dialog */
    var DIALOG_OFFSET_Y        = 0;                /* ダイアログの縦方向オフセット / vertical offset of the dialog */
    var DIALOG_OPACITY         = 0.98;             /* ダイアログの不透明度 / opacity of the dialog */

    /**
     * ダイアログ全体の並びと余白を設定する
     * @param {Window} targetWindow - 対象ウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの並びと余白を設定する
     * @param {Panel} targetPanel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 横並びの行グループを生成する
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @returns {Group} 生成したグループ
     */
    function addRow(parentContainer, horizontalAlign) {
        var createdGroup = parentContainer.add("group");
        setupRow(createdGroup, horizontalAlign);
        return createdGroup;
    }

    /**
     * 共通幅で右揃えした行ラベルを追加する
     * @param {Group} parentGroup - 追加先グループ
     * @param {string} text - ラベル文字列
     * @returns {StaticText} 追加した statictext
     */
    function addRowLabel(parentGroup, text) {
        var rowLabel = parentGroup.add("statictext", undefined, text);
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * スライダーを追加する
     * @param {Group} parentGroup - 追加先グループ
     * @param {number} value - 初期値
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @param {number} [width] - 幅（省略時は SLIDER_WIDTH）
     * @returns {Slider} 追加したスライダー
     */
    function addSlider(parentGroup, value, minValue, maxValue, width) {
        var slider = parentGroup.add("slider", undefined, value, minValue, maxValue);
        slider.preferredSize.width = (typeof width === "number") ? width : SLIDER_WIDTH;
        return slider;
    }

    /**
     * チェックボックスのラベル幅を最長のものにそろえる
     * @param {Array<Checkbox>} checkboxList - 対象チェックボックス
     * @returns {void}
     */
    function unifyCheckboxLabelWidth(checkboxList) {
        var unifiedWidth = 0;
        for (var i = 0; i < checkboxList.length; i++) {
            var checkboxText = String(checkboxList[i].text || "");
            var textWidth;
            try {
                /* measureString が使える場合は実測 / Measure the string when the API is available */
                textWidth = checkboxList[i].graphics.measureString(checkboxText).width;
            } catch (_) {
                /* フォールバック: 文字数ベース / Fall back to a character-count estimate */
                textWidth = checkboxText.length * 7;
            }
            var candidateWidth = Math.ceil(textWidth + 18); /* チェック部分の余白を足す / Add room for the check box itself */
            if (candidateWidth > unifiedWidth) unifiedWidth = candidateWidth;
        }
        for (var j = 0; j < checkboxList.length; j++) {
            checkboxList[j].preferredSize.width = unifiedWidth;
        }
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の表示言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + "";
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "紙吹雪を生成", en: "Generate Confetti" }
        },
        panel: {
            basic:  { ja: "基本設定", en: "Basic" },
            shape:  { ja: "形状", en: "Shapes" },
            random: { ja: "ランダム", en: "Randomize" }
        },
        fieldLabel: {
            baseSize:     { ja: "基準サイズ", en: "Base Size" },
            count:        { ja: "生成数", en: "Count" },
            distribution: { ja: "分布", en: "Distribution" },
            strength:     { ja: "強度", en: "Strength" },
            zoom:         { ja: "画面ズーム", en: "Zoom" }
        },
        radio: {
            distEven:   { ja: "全体に均等", en: "Uniform" },
            distGrad:   { ja: "垂直方向", en: "Top to Bottom" },
            distHollow: { ja: "放射状", en: "Radial Outward" }
        },
        checkbox: {
            circle:     { ja: "円", en: "Circle" },
            rect:       { ja: "長方形", en: "Rectangle" },
            square:     { ja: "正方形", en: "Square" },
            triangle:   { ja: "三角形", en: "Triangle" },
            star:       { ja: "スター", en: "Star" },
            sparkleA:   { ja: "キラキラA", en: "Sparkle A" },
            sparkleB:   { ja: "キラキラB", en: "Sparkle B" },
            heart:      { ja: "ハート", en: "Heart" },
            ribbon:     { ja: "リボン", en: "Ribbon" },
            symbol:     { ja: "シンボル", en: "Symbol" },
            mask:       { ja: "マスク処理", en: "Mask" },
            margin:     { ja: "マージン", en: "Margin" },
            randomSize: { ja: "大きさ", en: "Size" },
            opacity:    { ja: "不透明度", en: "Opacity" },
            skew:       { ja: "歪み", en: "Skew" },
            rotate:     { ja: "回転", en: "Rotation" }
        },
        dropdown: {
            symbolNone: { ja: "（なし）", en: "(None)" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('panel','shape')）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} 該当するラベル（見つからない場合は空文字）
     */
    function getLabel() {
        var labelNode = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[arguments[i]];
        }
        return (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} コロンを付けたラベル
     */
    function labelText() {
        var label = getLabel.apply(null, arguments);
        return label + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * デバッグ用にエラー内容を ExtendScript コンソールへ出力する
     * @param {Error} err - 捕捉した例外
     * @param {string} context - 発生箇所を示す文字列
     * @returns {void}
     */
    function logError(err, context) {
        var message = "[" + SCRIPT_NAME + "] " + String(context || "Error");
        try {
            if (err) {
                message += " :: " + String(err.message || err);
                if (err.line) message += " (line: " + String(err.line) + ")";
            }
            $.writeln(message);
        } catch (_) { }
    }

    if (!app.documents.length) {
        alert(getLabel("alert", "noDocument"));
        return;
    }

    // =========================================
    // 対象の判定 / Target detection
    // =========================================

    var doc = app.activeDocument;

    /* 選択オブジェクト（なければアートボードを対象にする）/ Selected object, or the artboard when nothing is selected */
    var selectedItem = (doc.selection.length > 0) ? doc.selection[0] : null;
    var useArtboardBounds = !selectedItem;
    var isTextSelection = !!(selectedItem && selectedItem.typename === "TextFrame");

    // =========================================
    // ビュー・ズーム / View & zoom
    // =========================================

    /**
     * 現在のビュー状態（ズーム・中心）を控える
     * @param {Document} targetDoc - 対象ドキュメント
     * @returns {object} ビュー状態 { view, zoom, center }
     */
    function captureViewState(targetDoc) {
        var viewState = { view: null, zoom: null, center: null };
        try {
            viewState.view = targetDoc.activeView;
            viewState.zoom = viewState.view.zoom;
            viewState.center = viewState.view.centerPoint;
        } catch (eCaptureView) {
            logError(eCaptureView, "captureViewState");
        }
        return viewState;
    }

    /**
     * 控えておいたビュー状態を復元する
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {object} viewState - captureViewState() の戻り値
     * @returns {void}
     */
    function restoreViewState(targetDoc, viewState) {
        if (!viewState) return;
        try {
            var view = viewState.view || targetDoc.activeView;
            if (!view) return;
            if (viewState.zoom != null) view.zoom = viewState.zoom;
            if (viewState.center != null) view.centerPoint = viewState.center;
        } catch (eRestoreView) {
            logError(eRestoreView, "restoreViewState");
        }
    }

    /**
     * 画面ズーム用のスライダー行を追加する
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @param {Document} targetDoc - 対象ドキュメント
     * @param {string} zoomLabelText - 行ラベル
     * @param {object} initialViewState - 初期ビュー状態
     * @param {object} zoomOptions - { min, max, sliderWidth, redraw }
     * @returns {object} ズーム操作用のインターフェイス
     */
    function addZoomControls(parentContainer, targetDoc, zoomLabelText, initialViewState, zoomOptions) {
        var minZoom = zoomOptions.min;
        var maxZoom = zoomOptions.max;
        var doRedraw = (zoomOptions.redraw !== false);

        var zoomGroup = addRow(parentContainer, "center");
        var zoomLabel = addRowLabel(zoomGroup, String(zoomLabelText));

        var initialZoom = 1;
        try {
            initialZoom = Number((initialViewState && initialViewState.zoom != null) ? initialViewState.zoom : targetDoc.activeView.zoom);
        } catch (eReadZoom) {
            logError(eReadZoom, "addZoomControls.readZoom");
        }
        if (!initialZoom || isNaN(initialZoom)) initialZoom = 1;

        var zoomSlider = addSlider(zoomGroup, initialZoom, minZoom, maxZoom, zoomOptions.sliderWidth);

        /**
         * 指定倍率をビューへ適用する
         * @param {number} zoomValue - 倍率
         * @returns {void}
         */
        function applyZoom(zoomValue) {
            try {
                var view = (initialViewState && initialViewState.view) ? initialViewState.view : targetDoc.activeView;
                if (!view) return;
                view.zoom = zoomValue;
                if (doRedraw) app.redraw();
            } catch (eApplyZoom) {
                logError(eApplyZoom, "addZoomControls.applyZoom");
            }
        }

        /**
         * 現在のビュー倍率をスライダーへ反映する
         * @returns {void}
         */
        function syncFromView() {
            try {
                var view = (initialViewState && initialViewState.view) ? initialViewState.view : targetDoc.activeView;
                if (view) zoomSlider.value = view.zoom;
            } catch (eSyncZoom) {
                logError(eSyncZoom, "addZoomControls.syncFromView");
            }
        }

        zoomSlider.onChanging = function () {
            applyZoom(Number(zoomSlider.value));
        };

        return {
            group: zoomGroup,
            label: zoomLabel,
            slider: zoomSlider,
            applyZoom: applyZoom,
            syncFromView: syncFromView,
            restoreInitial: function () { restoreViewState(targetDoc, initialViewState); }
        };
    }

    var initialViewState = captureViewState(doc);

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログの表示位置をずらす
     * @param {Window} targetWindow - 対象ウィンドウ
     * @param {number} dx - 横方向の移動量
     * @param {number} dy - 縦方向の移動量
     * @returns {void}
     */
    function shiftDialogPosition(targetWindow, dx, dy) {
        targetWindow.location = [targetWindow.location[0] + dx, targetWindow.location[1] + dy];
    }

    var dialog = new Window("dialog", getLabel("dialog", "title") + " " + SCRIPT_VERSION);
    setupWindow(dialog);
    dialog.opacity = DIALOG_OPACITY;

    /* 基本設定 / Basic */
    var basicPanel = addPanel(dialog, getLabel("panel", "basic"));

    var baseSizeRow = addRow(basicPanel);
    addRowLabel(baseSizeRow, labelText("fieldLabel", "baseSize"));
    /* スライダーの 0 が DEFAULT_BASE_SIZE_PT に対応する相対指定 / Slider 0 maps to DEFAULT_BASE_SIZE_PT */
    var baseSizeSlider = addSlider(baseSizeRow, 0, -5, 45);

    var countRow = addRow(basicPanel);
    addRowLabel(countRow, labelText("fieldLabel", "count"));
    var countSlider = addSlider(countRow, DEFAULT_COUNT, 10, 500);
    var confettiCount = DEFAULT_COUNT;

    /* マスク処理・マージン / Mask & margin */
    var maskMarginGroup = basicPanel.add("group");
    maskMarginGroup.orientation = "column";
    maskMarginGroup.alignChildren = ["left", "top"];
    maskMarginGroup.margins = [0, 10, 0, 10];

    var maskRow = addRow(maskMarginGroup);
    var maskCheckbox = maskRow.add("checkbox", undefined, getLabel("checkbox", "mask"));
    maskCheckbox.value = true;
    if (isTextSelection || useArtboardBounds) {
        maskCheckbox.value = false;
        maskCheckbox.enabled = false;
    }

    var marginRow = addRow(maskMarginGroup);
    var marginCheckbox = marginRow.add("checkbox", undefined, getLabel("checkbox", "margin"));
    marginCheckbox.value = false;
    var marginSlider = addSlider(marginRow, 0, 0, 50);
    marginSlider.enabled = marginCheckbox.value;
    var generationMarginPt = 0;

    /* 配置分布 / Distribution */
    var distributionGroup = basicPanel.add("group");
    distributionGroup.orientation = "column";
    distributionGroup.alignChildren = ["left", "top"];

    var distributionRow = addRow(distributionGroup);
    addRowLabel(distributionRow, labelText("fieldLabel", "distribution"));
    var evenRadio = distributionRow.add("radiobutton", undefined, getLabel("radio", "distEven"));
    var verticalRadio = distributionRow.add("radiobutton", undefined, getLabel("radio", "distGrad"));
    var radialRadio = distributionRow.add("radiobutton", undefined, getLabel("radio", "distHollow"));
    evenRadio.value = true;

    var strengthRow = addRow(distributionGroup);
    addRowLabel(strengthRow, labelText("fieldLabel", "strength"));
    var strengthSlider = addSlider(strengthRow, DEFAULT_STRENGTH, 1.0, 6.0, NARROW_SLIDER_WIDTH);
    var distributionStrength = DEFAULT_STRENGTH;
    strengthRow.enabled = false;

    /* 形状 / Shapes */
    var shapePanel = addPanel(dialog, getLabel("panel", "shape"));

    var shapeColumnsGroup = shapePanel.add("group");
    shapeColumnsGroup.orientation = "row";
    shapeColumnsGroup.alignment = ["left", "top"];
    shapeColumnsGroup.alignChildren = ["left", "top"];
    shapeColumnsGroup.spacing = 0;

    var shapeColumns = [];
    for (var i = 0; i < 3; i++) {
        var shapeColumn = shapeColumnsGroup.add("group");
        shapeColumn.orientation = "column";
        shapeColumn.alignChildren = ["left", "top"];
        shapeColumns.push(shapeColumn);
    }

    /**
     * 形状の定義（表示位置・既定値・生成関数・単独選択時のプリセット）
     * @typedef {object} ShapeDef
     * @property {string} key - LABELS.checkbox のキー
     * @property {number} column - 表示するカラム番号（0..2）
     * @property {boolean} defaultOn - 初期状態でONにするか
     * @property {number} sizeScale - 基準サイズに掛ける倍率
     * @property {string} soloPreset - Option+クリック時のプリセット（"" / "noRotate" / "noRotateNoSkew"）
     * @property {function} create - 形状を生成する関数
     */
    var SHAPE_DEFS = [
        { key: "circle",   column: 0, defaultOn: true,  sizeScale: 1.0,        soloPreset: "noRotate",       create: createCircle },
        { key: "triangle", column: 0, defaultOn: true,  sizeScale: 1.0,        soloPreset: "",               create: createTriangle },
        { key: "heart",    column: 0, defaultOn: false, sizeScale: 1.3 * 0.9,  soloPreset: "noRotateNoSkew", create: createHeart },
        { key: "square",   column: 1, defaultOn: false, sizeScale: 1.0,        soloPreset: "",               create: createSquare },
        { key: "rect",     column: 1, defaultOn: true,  sizeScale: 1.0,        soloPreset: "",               create: createRectangle },
        { key: "ribbon",   column: 1, defaultOn: false, sizeScale: 1.0,        soloPreset: "",               create: createRibbon },
        { key: "star",     column: 2, defaultOn: false, sizeScale: 1.2,        soloPreset: "",               create: createStar },
        { key: "sparkleA", column: 2, defaultOn: false, sizeScale: 1.4,        soloPreset: "",               create: createSparkleA },
        { key: "sparkleB", column: 2, defaultOn: false, sizeScale: 1.6,        soloPreset: "noRotateNoSkew", create: createSparkleB }
    ];

    /* 形状チェックボックスの一覧（シンボルを含む）/ Shape toggles including the symbol entry */
    var shapeToggles = [];
    for (var i = 0; i < SHAPE_DEFS.length; i++) {
        var shapeItem = SHAPE_DEFS[i];
        shapeItem.checkbox = shapeColumns[shapeItem.column].add("checkbox", undefined, getLabel("checkbox", shapeItem.key));
        shapeItem.checkbox.value = shapeItem.defaultOn;
        shapeToggles.push(shapeItem);
    }

    /* シンボル行（チェック + ドロップダウン）/ Symbol row */
    var symbolRow = addRow(shapePanel);
    var symbolCheckbox = symbolRow.add("checkbox", undefined, getLabel("checkbox", "symbol"));
    symbolCheckbox.value = false;
    symbolCheckbox.enabled = false; /* シンボル未選択のうちはディム / Dimmed until a symbol is picked */
    var symbolDropdown = symbolRow.add("dropdownlist", undefined, [getLabel("dropdown", "symbolNone")]);
    symbolDropdown.selection = 0;
    symbolDropdown.preferredSize.width = SYMBOL_DROPDOWN_WIDTH;

    var symbolToggle = { key: "symbol", checkbox: symbolCheckbox, soloPreset: "" };
    shapeToggles.push(symbolToggle);

    var shapeCheckboxes = [];
    for (var i = 0; i < shapeToggles.length; i++) {
        shapeCheckboxes.push(shapeToggles[i].checkbox);
    }
    unifyCheckboxLabelWidth(shapeCheckboxes);

    /* ランダム / Randomize */
    var randomPanel = addPanel(dialog, getLabel("panel", "random"));

    var randomSizeRow = addRow(randomPanel);
    var randomSizeCheckbox = randomSizeRow.add("checkbox", undefined, getLabel("checkbox", "randomSize"));
    randomSizeCheckbox.value = true;
    var randomSizeSlider = addSlider(randomSizeRow, 100, 100, 300);
    randomSizeSlider.enabled = randomSizeCheckbox.value;
    var randomSizeStrength = 100;

    var opacityRow = addRow(randomPanel);
    var opacityCheckbox = opacityRow.add("checkbox", undefined, getLabel("checkbox", "opacity"));
    opacityCheckbox.value = true;
    /* スライダーは反転指定（値 = 100 − 不透明度の下限）/ Reversed slider: value = 100 − minimum opacity */
    var opacitySlider = addSlider(opacityRow, 100 - DEFAULT_OPACITY_MIN, 0, 100);
    opacitySlider.enabled = opacityCheckbox.value;
    var opacityMin = DEFAULT_OPACITY_MIN;

    var skewRow = addRow(randomPanel);
    var skewCheckbox = skewRow.add("checkbox", undefined, getLabel("checkbox", "skew"));
    skewCheckbox.value = false;
    var skewSlider = addSlider(skewRow, 0, 0, SKEW_MAX_DEG);
    skewSlider.enabled = skewCheckbox.value;
    var skewMaxDeg = 0;

    var rotateRow = addRow(randomPanel);
    var rotateCheckbox = rotateRow.add("checkbox", undefined, getLabel("checkbox", "rotate"));
    rotateCheckbox.value = true;
    var rotateSlider = addSlider(rotateRow, DEFAULT_ROTATE_MAX, 0, 360);
    rotateSlider.enabled = rotateCheckbox.value;
    var rotateMaxDeg = DEFAULT_ROTATE_MAX;
    var previousRotateMaxDeg = DEFAULT_ROTATE_MAX; /* 回転を一時OFFにしたときの復元値 / Value restored when rotation is turned back on */

    unifyCheckboxLabelWidth([randomSizeCheckbox, opacityCheckbox, skewCheckbox, rotateCheckbox]);

    /* ズーム / Zoom */
    var zoomControls = addZoomControls(dialog, doc, labelText("fieldLabel", "zoom"), initialViewState, {
        min: 0.1,
        max: 16,
        sliderWidth: SLIDER_WIDTH,
        redraw: true
    });

    /* ボタンエリア / Button row */
    var btnRowGroup = dialog.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
    btnRowGroup.alignment = ["fill", "bottom"];

    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];
    var btnCancel = btnRightGroup.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" });
    var btnOK = btnRightGroup.add("button", undefined, getLabel("button", "ok"), { name: "ok" });
    dialog.defaultElement = btnOK;

    // =========================================
    // 生成エリアの算出 / Generation area
    // =========================================

    /**
     * 基準となる領域（選択オブジェクトまたはアートボード）を取得する
     * @returns {Array<number>|null} [左, 上, 右, 下]（取得できない場合は null）
     */
    function getTargetBounds() {
        try {
            if (useArtboardBounds) {
                return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            }
            if (selectedItem) return selectedItem.geometricBounds;
        } catch (eTargetBounds) {
            logError(eTargetBounds, "getTargetBounds");
        }
        return null;
    }

    /**
     * マージンを反映した生成エリアを取得する
     * @returns {Array<number>|null} [左, 上, 右, 下]（取得できない場合は null）
     */
    function getEffectiveBounds() {
        var bounds = getTargetBounds();
        if (!bounds) return null;

        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        if (marginCheckbox.value && generationMarginPt !== 0) {
            left -= generationMarginPt;
            top += generationMarginPt;
            right += generationMarginPt;
            bottom -= generationMarginPt;
        }
        return [left, top, right, bottom];
    }

    /**
     * 対象サイズからマージンの上限値を求める（幅と高さの合計の 1/6）
     * @returns {number} マージンの最大値（pt）
     */
    function computeMaxMarginFromTarget() {
        var bounds = getTargetBounds();
        if (!bounds) return 50;

        var width = Math.abs(Number(bounds[2]) - Number(bounds[0]));
        var height = Math.abs(Number(bounds[1]) - Number(bounds[3]));
        var maxMargin = (width + height) / 6;

        if (!maxMargin || isNaN(maxMargin) || maxMargin < 0) return 50;
        return Math.min(maxMargin, 100000); /* 念のため上限 / Guard against absurd values */
    }

    /**
     * マージンスライダーの上限を対象サイズに合わせて更新する
     * @returns {void}
     */
    function applyMarginMaxToUI() {
        var maxMargin = computeMaxMarginFromTarget();
        marginSlider.maxvalue = maxMargin;
        if (Number(marginSlider.value) > maxMargin) marginSlider.value = maxMargin;
        if (generationMarginPt > maxMargin) generationMarginPt = maxMargin;
    }

    /**
     * 基準サイズ（pt）を取得する
     * @returns {number} 基準サイズ（0.1pt 刻み）
     */
    function getBaseSizePt() {
        var sizePt = DEFAULT_BASE_SIZE_PT + Number(baseSizeSlider.value);
        if (isNaN(sizePt)) sizePt = DEFAULT_BASE_SIZE_PT;
        if (sizePt < 0.1) sizePt = 0.1;
        return Math.round(sizePt * 10) / 10;
    }

    /**
     * 1個ぶんの紙吹雪サイズを求める（ランダム量を反映）
     * @returns {number} サイズ（pt）
     */
    function getConfettiSize() {
        var baseSize = getBaseSizePt();
        if (!randomSizeCheckbox.value) return baseSize;

        var strength = Math.min(Math.max(randomSizeStrength, 100), 300);
        /* 100 でほぼ固定、300 で最大 ±150% の揺れ / 100 keeps the size fixed, 300 allows ±150% */
        var maxRange = baseSize * 1.5 * ((strength - 100) / 200);
        var size = baseSize + randomBetween(-maxRange, maxRange);
        return (size < 0.1) ? 0.1 : size;
    }

    /**
     * min 以上 max 未満の乱数を返す
     * @param {number} min - 下限
     * @param {number} max - 上限
     * @returns {number} 乱数
     */
    function randomBetween(min, max) {
        return Math.random() * (max - min) + min;
    }

    /**
     * 分布設定に従って生成位置をサンプリングする
     * @param {Array<number>} bounds - [左, 上, 右, 下]
     * @returns {object} 生成位置 { x, y }
     */
    function pickConfettiPoint(bounds) {
        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        var centerX = (left + right) * 0.5;
        var centerY = (top + bottom) * 0.5;

        /* 全体に均等 / Uniform */
        if (evenRadio.value) {
            return { x: randomBetween(left, right), y: randomBetween(bottom, top) };
        }

        /* 垂直方向（上を濃く、下を薄く）/ Top-biased gradient */
        if (verticalRadio.value) {
            var height = top - bottom;
            var gradientPower = 1 + (distributionStrength - 1) * 0.6; /* 1..4 程度へマップ / Map roughly to 1..4 */
            var verticalRatio = 1 - Math.pow(1 - Math.random(), gradientPower);
            return { x: randomBetween(left, right), y: bottom + height * verticalRatio };
        }

        /* 放射状（中心ほど出にくい）/ Radial outward */
        var angle = randomBetween(0, Math.PI * 2);
        var cos = Math.cos(angle);
        var sin = Math.sin(angle);

        /* その角度で矩形内に収まる最大半径 / Largest radius that stays inside the rectangle */
        var maxRadiusX = (cos === 0) ? 1e12 : (cos > 0 ? (right - centerX) / cos : (left - centerX) / cos);
        var maxRadiusY = (sin === 0) ? 1e12 : (sin > 0 ? (top - centerY) / sin : (bottom - centerY) / sin);
        var maxRadius = Math.min(Math.abs(maxRadiusX), Math.abs(maxRadiusY));

        var strength = Math.min(Math.max(distributionStrength, 1), 6);
        var radius = maxRadius * (1 - Math.pow(1 - Math.random(), strength));

        return { x: centerX + cos * radius, y: centerY + sin * radius };
    }

    // =========================================
    // 形状の生成 / Shape creation
    // =========================================

    /**
     * 長方形を生成する（横長寄りの比率）
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 基準サイズ
     * @returns {PathItem} 生成したパス
     */
    function createRectangle(container, left, top, size) {
        return container.pathItems.rectangle(top, left, size * 1.4, size * 1.4 * randomBetween(0.3, 0.6));
    }

    /**
     * 正方形を生成する
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 一辺の長さ
     * @returns {PathItem} 生成したパス
     */
    function createSquare(container, left, top, size) {
        return container.pathItems.rectangle(top, left, size, size);
    }

    /**
     * 円を生成する
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 直径
     * @returns {PathItem} 生成したパス
     */
    function createCircle(container, left, top, size) {
        return container.pathItems.ellipse(top, left, size, size);
    }

    /**
     * 三角形を生成する
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 底辺の長さ
     * @returns {PathItem} 生成したパス
     */
    function createTriangle(container, left, top, size) {
        var triangle = container.pathItems.add();
        var height = size * 1.2;
        triangle.setEntirePath([
            [left, top],
            [left + size, top],
            [left + size * 0.5, top - height]
        ]);
        triangle.closed = true;
        return triangle;
    }

    /**
     * 星型のパスを生成する（外接正方形の左上基準）
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 外接正方形の一辺
     * @param {number} pointCount - 頂点の数
     * @param {number} innerRatio - 外側半径に対する内側半径の比
     * @param {number} startDeg - 最初の頂点の角度
     * @returns {PathItem} 生成したパス
     */
    function createStarPath(container, left, top, size, pointCount, innerRatio, startDeg) {
        var outerRadius = size * 0.5;
        var innerRadius = outerRadius * innerRatio;
        var centerX = left + outerRadius;
        var centerY = top - outerRadius;

        var vertexCount = pointCount * 2;
        var angleStep = 360 / vertexCount;
        var points = [];
        for (var i = 0; i < vertexCount; i++) {
            var radius = (i % 2 === 0) ? outerRadius : innerRadius;
            var angle = (startDeg + i * angleStep) * Math.PI / 180;
            points.push([centerX + Math.cos(angle) * radius, centerY + Math.sin(angle) * radius]);
        }

        var star = container.pathItems.add();
        star.setEntirePath(points);
        star.closed = true;
        return star;
    }

    /**
     * 五芒星を生成する
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 外接正方形の一辺
     * @returns {PathItem} 生成したパス
     */
    function createStar(container, left, top, size) {
        return createStarPath(container, left, top, size, 5, 0.5, 90);
    }

    /**
     * キラキラA（尖った四芒星）を生成する
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 外接正方形の一辺
     * @returns {PathItem} 生成したパス
     */
    function createSparkleA(container, left, top, size) {
        return createStarPath(container, left, top, size, 4, 0.25, -90);
    }

    /**
     * キラキラB（ベジエで辺をへこませた四芒星）を生成する
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 外接正方形の一辺
     * @returns {PathItem} 生成したパス
     */
    function createSparkleB(container, left, top, size) {
        var curveFactor = 0.8; /* 中心へハンドルを引き込む量 / How far handles are pulled toward the center */
        var centerX = left + size / 2;
        var centerY = top - size / 2;
        var halfSize = size / 2;
        var handleLength = halfSize * curveFactor;

        /* 上・右・下・左の4頂点と、そのハンドル位置 / Four vertices and their handle positions */
        var vertices = [
            { anchor: [centerX, centerY + halfSize], handle: [centerX, centerY + halfSize - handleLength] },
            { anchor: [centerX + halfSize, centerY], handle: [centerX + halfSize - handleLength, centerY] },
            { anchor: [centerX, centerY - halfSize], handle: [centerX, centerY - halfSize + handleLength] },
            { anchor: [centerX - halfSize, centerY], handle: [centerX - halfSize + handleLength, centerY] }
        ];

        var sparkle = container.pathItems.add();
        for (var i = 0; i < vertices.length; i++) {
            var pathPoint = sparkle.pathPoints.add();
            pathPoint.anchor = vertices[i].anchor;
            pathPoint.leftDirection = vertices[i].handle;
            pathPoint.rightDirection = vertices[i].handle;
            pathPoint.pointType = PointType.CORNER;
        }
        sparkle.closed = true;
        return sparkle;
    }

    /**
     * リボン（1/4円弧を2つつないだS字の帯）を生成する
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 基準サイズ
     * @returns {PathItem} 生成したパス
     */
    function createRibbon(container, left, top, size) {
        var centerX = left + size * 0.5;
        var centerY = top - size * 0.5;
        var radius = Math.max(0.8, size * 0.72);   /* 円弧の半径 / arc radius */
        var halfWidth = Math.max(0.25, size * 0.165); /* 帯の半幅 / half width of the band */
        var segmentCount = 10;                     /* 円弧の分割数 / arc subdivisions */

        /**
         * 円周上の点を返す
         * @param {number} circleCenterX - 円の中心X
         * @param {number} circleCenterY - 円の中心Y
         * @param {number} deg - 角度（度）
         * @returns {Array<number>} [x, y]
         */
        function pointOnCircle(circleCenterX, circleCenterY, deg) {
            var angle = deg * Math.PI / 180;
            return [circleCenterX + Math.cos(angle) * radius, circleCenterY + Math.sin(angle) * radius];
        }

        /* 中心線: 1つ目は (0,0) を中心に 180°→90°、2つ目は (0,2r) を中心に -90°→0° / Center line built from two quarter arcs */
        var centerPoints = [];
        for (var i = 0; i <= segmentCount; i++) {
            centerPoints.push(pointOnCircle(0, 0, 180 - 90 * (i / segmentCount)));
        }
        for (var j = 1; j <= segmentCount; j++) {
            centerPoints.push(pointOnCircle(0, 2 * radius, -90 + 90 * (j / segmentCount)));
        }

        /* 中心線のYレンジは 0..2r なので、その中点を (centerX, centerY) に合わせる / Center the 0..2r span on the target point */
        var offsetX = centerX;
        var offsetY = centerY - radius;

        var upperPoints = [];
        var lowerPoints = [];
        for (var k = 0; k < centerPoints.length; k++) {
            var previousPoint = centerPoints[(k === 0) ? 0 : (k - 1)];
            var currentPoint = centerPoints[k];
            var nextPoint = centerPoints[(k === centerPoints.length - 1) ? k : (k + 1)];

            /* 接線から左法線を作り、上下にオフセットして輪郭にする / Offset along the normal to build the outline */
            var tangentX = nextPoint[0] - previousPoint[0];
            var tangentY = nextPoint[1] - previousPoint[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            var normalX = -(tangentY / tangentLength);
            var normalY = tangentX / tangentLength;

            var x = currentPoint[0] + offsetX;
            var y = currentPoint[1] + offsetY;
            upperPoints.push([x + normalX * halfWidth, y + normalY * halfWidth]);
            lowerPoints.push([x - normalX * halfWidth, y - normalY * halfWidth]);
        }

        var points = [];
        for (var u = 0; u < upperPoints.length; u++) points.push(upperPoints[u]);
        for (var d = lowerPoints.length - 1; d >= 0; d--) points.push(lowerPoints[d]);

        var ribbon = container.pathItems.add();
        ribbon.closed = true;
        ribbon.setEntirePath(points);

        /* 角を立てず滑らかに見せる / Smooth every point so the band reads as a ribbon */
        try {
            for (var m = 0; m < ribbon.pathPoints.length; m++) {
                ribbon.pathPoints[m].pointType = PointType.SMOOTH;
            }
        } catch (eSmoothRibbon) {
            logError(eSmoothRibbon, "createRibbon.smooth");
        }
        ribbon.filled = true;
        ribbon.stroked = false;
        return ribbon;
    }

    /**
     * ハートを生成する（幅120・高さ80の基準座標を size にスケーリング）
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 基準サイズ
     * @returns {PathItem} 生成したパス
     */
    function createHeart(container, left, top, size) {
        var scale = size / 120;
        var centerX = left + size / 2;
        var centerY = top - size / 2;

        /**
         * 基準座標をスケーリングして実座標へ変換する
         * @param {number} x - 基準座標のX
         * @param {number} y - 基準座標のY
         * @returns {Array<number>} [x, y]
         */
        function scaled(x, y) {
            return [centerX + x * scale, centerY + y * scale];
        }

        /* 上のくぼみ → 右のふくらみ → 下の尖り → 左のふくらみ / Cleft, right lobe, tip, left lobe */
        var vertices = [
            { anchor: scaled(0, 30),    left: scaled(-20, 60),  right: scaled(20, 60),   type: PointType.CORNER },
            { anchor: scaled(60, 15),   left: scaled(60, 50),   right: scaled(60, -15),  type: PointType.SMOOTH },
            { anchor: scaled(0, -55),   left: scaled(20, -20),  right: scaled(-20, -20), type: PointType.CORNER },
            { anchor: scaled(-60, 15),  left: scaled(-60, -15), right: scaled(-60, 50),  type: PointType.SMOOTH }
        ];

        var heart = container.pathItems.add();
        for (var i = 0; i < vertices.length; i++) {
            var pathPoint = heart.pathPoints.add();
            pathPoint.anchor = vertices[i].anchor;
            pathPoint.leftDirection = vertices[i].left;
            pathPoint.rightDirection = vertices[i].right;
            pathPoint.pointType = vertices[i].type;
        }
        heart.closed = true;
        return heart;
    }

    // =========================================
    // シンボル / Symbols
    // =========================================

    var selectedSymbolName = "";
    var selectedSymbolRef = null;

    /**
     * ドキュメント内のシンボルでドロップダウンを組み直す
     * @returns {void}
     */
    function refreshSymbolDropdown() {
        symbolDropdown.removeAll();
        symbolDropdown.add("item", getLabel("dropdown", "symbolNone"));

        var hasSymbols = (doc.symbols.length > 0);
        for (var i = 0; i < doc.symbols.length; i++) {
            var symbolDef = doc.symbols[i];
            var symbolEntry = symbolDropdown.add("item", String(symbolDef.name || ("Symbol " + (i + 1))));
            symbolEntry._symbolDef = symbolDef;
        }
        symbolDropdown.enabled = hasSymbols;
        symbolDropdown.selection = 0;

        /* シンボル未選択なので形状パネルの「シンボル」はディム / Dim the symbol shape until one is picked */
        selectedSymbolName = "";
        selectedSymbolRef = null;
        symbolCheckbox.value = false;
        symbolCheckbox.enabled = false;
    }

    /**
     * シンボルが選択されているかを返す
     * @returns {boolean} 選択されていれば true
     */
    function hasSelectedSymbol() {
        return !!(selectedSymbolRef || selectedSymbolName);
    }

    /**
     * 選択中のシンボル定義を取得する（名前からの再探索を含む）
     * @returns {Symbol|null} シンボル定義（見つからない場合は null）
     */
    function resolveSelectedSymbol() {
        if (selectedSymbolRef) return selectedSymbolRef;
        if (!selectedSymbolName) return null;
        for (var i = 0; i < doc.symbols.length; i++) {
            if (String(doc.symbols[i].name) === String(selectedSymbolName)) return doc.symbols[i];
        }
        return null;
    }

    /**
     * シンボルインスタンスを1個生成し、指定サイズへ収める
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} size - 最大辺の長さ
     * @returns {PageItem|null} 生成したアイテム（失敗時は null）
     */
    function createSymbolConfetti(container, size) {
        var symbolDef = resolveSelectedSymbol();
        if (!symbolDef) return null;

        var wrapGroup = null;
        var symbolItem = null;
        try {
            wrapGroup = container.groupItems.add();
            wrapGroup.name = SYMBOL_WRAP_NAME;
            /* SymbolItem はレイヤー直下に作るのが確実 / Creating the instance on a layer is the reliable path */
            symbolItem = doc.activeLayer.symbolItems.add(symbolDef);
            symbolItem.move(wrapGroup, ElementPlacement.PLACEATEND);
        } catch (eCreateSymbol) {
            logError(eCreateSymbol, "createSymbolConfetti");
            if (symbolItem && !wrapGroup) {
                try { symbolItem.remove(); } catch (_) { }
            }
            if (wrapGroup) {
                try { wrapGroup.remove(); } catch (_) { }
            }
            return null;
        }

        /* 最大辺が size になるよう等比スケール / Scale proportionally so the longest side matches size */
        var longestSide = Math.max(Number(wrapGroup.width), Number(wrapGroup.height));
        if (longestSide > 0) {
            var scalePercent = Math.max(1, (size / longestSide) * 100);
            wrapGroup.resize(scalePercent, scalePercent);
        }
        return wrapGroup;
    }

    // =========================================
    // 変形・体裁 / Transform & appearance
    // =========================================

    /**
     * 中心を基準にX方向のシアー（歪み）を適用する
     * @param {PageItem} item - 対象アイテム
     * @param {number} deg - シアー角度（度）
     * @returns {void}
     */
    function applyShearToItem(item, deg) {
        if (!item || !deg) return;
        var shearMatrix = new Matrix();
        /* X方向シアー: [1 tan; 0 1] / Shear along X */
        shearMatrix.mValueA = 1;
        shearMatrix.mValueB = Math.tan(deg * Math.PI / 180);
        shearMatrix.mValueC = 0;
        shearMatrix.mValueD = 1;
        shearMatrix.mValueTX = 0;
        shearMatrix.mValueTY = 0;

        try {
            item.transform(shearMatrix, true, true, true, true, 1, Transformation.CENTER);
        } catch (_) {
            /* フォールバック: shear API（環境によってはこちらが効く）/ Fall back to the shear API */
            try { item.shear(deg); } catch (eShear) { logError(eShear, "applyShearToItem"); }
        }
    }

    /**
     * 設定に従ってランダムな回転を適用する
     * @param {PageItem} item - 対象アイテム
     * @returns {void}
     */
    function applyRandomRotate(item) {
        if (!rotateCheckbox.value) return;
        var maxDeg = Math.min(Math.max(rotateMaxDeg, 0), 360);
        if (maxDeg > 0) item.rotate(randomBetween(-maxDeg, maxDeg));
    }

    /**
     * 設定に従ってランダムな歪みを適用する
     * @param {PageItem} item - 対象アイテム
     * @returns {void}
     */
    function applyRandomSkew(item) {
        if (!skewCheckbox.value) return;
        var maxDeg = Math.min(Math.max(skewMaxDeg, 0), SKEW_MAX_DEG);
        if (maxDeg > 0) applyShearToItem(item, randomBetween(-maxDeg, maxDeg));
    }

    /**
     * 設定に従ってランダムな不透明度を適用する
     * @param {PageItem} item - 対象アイテム
     * @returns {void}
     */
    function applyRandomOpacity(item) {
        if (!opacityCheckbox.value) {
            item.opacity = 100;
            return;
        }
        item.opacity = randomBetween(Math.min(Math.max(opacityMin, 0), 100), 100);
    }

    /**
     * アイテムの左上を指定座標へ合わせる
     * @param {PageItem} item - 対象アイテム
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @returns {void}
     */
    function alignItemTopLeft(item, left, top) {
        var bounds = item.geometricBounds; /* [左, 上, 右, 下] / [L, T, R, B] */
        item.left = Number(item.left) + (Number(left) - Number(bounds[0]));
        item.top = Number(item.top) + (Number(top) - Number(bounds[1]));
    }

    /**
     * 紙吹雪に色を設定する（開いたパスは線、閉じたパスは塗り）
     * @param {PathItem} item - 対象パス
     * @param {Array<number>} color - [R, G, B]
     * @returns {void}
     */
    function applyConfettiColor(item, color) {
        var rgbColor = new RGBColor();
        rgbColor.red = color[0];
        rgbColor.green = color[1];
        rgbColor.blue = color[2];

        if (item.closed === false) {
            item.filled = false;
            item.stroked = true;
            item.strokeColor = rgbColor;
        } else {
            item.filled = true;
            item.stroked = false;
            item.fillColor = rgbColor;
        }
    }

    /**
     * ONになっている形状からランダムに1つ選ぶ
     * @returns {ShapeDef|null} 選ばれた形状（候補がない場合は null）
     */
    function pickEnabledShape() {
        var candidates = [];
        for (var i = 0; i < shapeToggles.length; i++) {
            var shapeToggle = shapeToggles[i];
            if (!shapeToggle.checkbox.value) continue;
            if (shapeToggle.key === "symbol" && !hasSelectedSymbol()) continue;
            candidates.push(shapeToggle);
        }
        if (candidates.length === 0) return null;
        return candidates[Math.floor(Math.random() * candidates.length)];
    }

    /**
     * 紙吹雪を1個生成する
     * @param {Layer|GroupItem} container - 追加先
     * @param {number} left - 左端
     * @param {number} top - 上端
     * @param {number} size - 基準サイズ
     * @param {Array<number>} color - [R, G, B]
     * @returns {PageItem|null} 生成したアイテム（生成できなかった場合は null）
     */
    function createConfetti(container, left, top, size, color) {
        var shapeToggle = pickEnabledShape();
        if (!shapeToggle) return null;

        if (shapeToggle.key === "symbol") {
            var symbolConfetti = createSymbolConfetti(container, size);
            if (!symbolConfetti) return null;
            applyRandomRotate(symbolConfetti);
            applyRandomSkew(symbolConfetti);
            alignItemTopLeft(symbolConfetti, left, top);
            applyRandomOpacity(symbolConfetti);
            return symbolConfetti;
        }

        var confetti = shapeToggle.create(container, left, top, size * shapeToggle.sizeScale);
        if (!confetti) return null;
        applyConfettiColor(confetti, color);
        applyRandomRotate(confetti);
        applyRandomSkew(confetti);
        applyRandomOpacity(confetti);
        return confetti;
    }

    // =========================================
    // マスク処理 / Masking
    // =========================================

    /**
     * アイテムにクリッピング指定を立てる
     * @param {PageItem} item - 対象アイテム
     * @returns {boolean} 指定できたら true
     */
    function setClippingFlag(item) {
        if (!item) return false;
        try {
            if (item.typename === "PathItem") {
                item.clipping = true;
                return true;
            }
            if (item.typename === "CompoundPathItem" && item.pathItems.length > 0) {
                item.pathItems[0].clipping = true;
                return true;
            }
            if (item.typename === "GroupItem") {
                /* Pathfinder の結果がグループになる場合があるため内側のパスを使う / Pathfinder can return a group */
                if (item.compoundPathItems.length > 0 && item.compoundPathItems[0].pathItems.length > 0) {
                    item.compoundPathItems[0].pathItems[0].clipping = true;
                    return true;
                }
                if (item.pathItems.length > 0) {
                    item.pathItems[0].clipping = true;
                    return true;
                }
            }
        } catch (eSetClipping) {
            logError(eSetClipping, "setClippingFlag");
        }
        return false;
    }

    /**
     * マスク処理の設定に応じて、紙吹雪を入れる器を用意する
     * @param {Layer} parentLayer - 追加先レイヤー
     * @returns {object} { container: 紙吹雪の追加先, maskGroup: マスク対象グループ（不要なら null） }
     */
    function createConfettiContainer(parentLayer) {
        if (!maskCheckbox.value) return { container: parentLayer, maskGroup: null };
        var maskGroup = parentLayer.groupItems.add();
        maskGroup.name = MASK_GROUP_NAME;
        return { container: maskGroup, maskGroup: maskGroup };
    }

    /**
     * グループを複製して合体し、単一のパスにする
     * @param {GroupItem} container - 複製先
     * @param {GroupItem} groupItem - 元グループ
     * @returns {PageItem|null} 合体結果（失敗時は null）
     */
    function uniteGroupToSinglePath(container, groupItem) {
        var previousSelection = doc.selection;
        var duplicated = null;
        try {
            duplicated = groupItem.duplicate(container, ElementPlacement.PLACEATEND);
        } catch (eDuplicateGroup) {
            logError(eDuplicateGroup, "uniteGroupToSinglePath.duplicate");
            return null;
        }

        try {
            /* 選択を複製グループに切り替えて合体→展開 / Select the copy, then unite and expand */
            doc.selection = null;
            duplicated.selected = true;
            app.executeMenuCommand('Live Pathfinder Add');
            app.executeMenuCommand('expandStyle');
            if (doc.selection && doc.selection.length > 0) return doc.selection[0];
        } catch (eUniteGroup) {
            logError(eUniteGroup, "uniteGroupToSinglePath.pathfinder");
            try { duplicated.remove(); } catch (_) { }
            return null;
        } finally {
            doc.selection = null;
            var itemsToReselect = (previousSelection && previousSelection.length) ? previousSelection : [selectedItem];
            for (var i = 0; i < itemsToReselect.length; i++) {
                try { itemsToReselect[i].selected = true; } catch (_) { }
            }
        }
        return null;
    }

    /**
     * クリップグループからクリッピングパスの候補を探す
     * @param {GroupItem} groupItem - 対象グループ
     * @returns {PageItem|null} 候補のパス（見つからない場合は null）
     */
    function findClippingCandidate(groupItem) {
        /* 穴あき形状を想定して CompoundPath を先に探す / Look at compound paths first */
        for (var i = 0; i < groupItem.compoundPathItems.length; i++) {
            var compoundPath = groupItem.compoundPathItems[i];
            if (compoundPath.pathItems.length > 0 && compoundPath.pathItems[0].clipping) return compoundPath;
        }
        for (var j = 0; j < groupItem.pathItems.length; j++) {
            if (groupItem.pathItems[j].clipping) return groupItem.pathItems[j];
        }
        /* クリッピング指定が見つからない場合は最初のパスを使う（最後の手段）/ Fall back to the first path */
        if (groupItem.compoundPathItems.length > 0) return groupItem.compoundPathItems[0];
        if (groupItem.pathItems.length > 0) return groupItem.pathItems[0];
        return null;
    }

    /**
     * 選択オブジェクトに合わせたマスク図形を作る
     * @param {GroupItem} container - 追加先
     * @param {PageItem} sourceItem - 元になる選択オブジェクト
     * @returns {PageItem|null} マスク図形（作れない場合は null）
     */
    function createMaskShapeFromSelection(container, sourceItem) {
        if (!container || !sourceItem) return null;

        try {
            /* Path / CompoundPath / TextFrame はそのまま複製 / Duplicate these as-is */
            var typeName = sourceItem.typename;
            if (typeName === "PathItem" || typeName === "CompoundPathItem" || typeName === "TextFrame") {
                return sourceItem.duplicate(container, ElementPlacement.PLACEATEND);
            }
            if (typeName === "GroupItem") {
                var clipCandidate = findClippingCandidate(sourceItem);
                if (clipCandidate) return clipCandidate.duplicate(container, ElementPlacement.PLACEATEND);
                /* 通常グループなど候補が取れない場合は合体して単一パス化 / Unite a plain group into one path */
                return uniteGroupToSinglePath(container, sourceItem);
            }
        } catch (eMaskShape) {
            logError(eMaskShape, "createMaskShapeFromSelection");
        }

        /* フォールバック: バウンディング矩形 / Fall back to the bounding rectangle */
        var bounds;
        try {
            bounds = sourceItem.geometricBounds; /* [左, 上, 右, 下] / [L, T, R, B] */
        } catch (eMaskBounds) {
            logError(eMaskBounds, "createMaskShapeFromSelection.bounds");
            return null;
        }
        var width = bounds[2] - bounds[0];
        var height = bounds[1] - bounds[3];
        if (!width || !height) return null;

        var maskRect = container.pathItems.rectangle(bounds[1], bounds[0], width, height);
        /* クリッピングパスは存在が必要なので、塗りをNoColorにして不可視にする / Keep it present but invisible */
        maskRect.stroked = false;
        maskRect.filled = true;
        maskRect.fillColor = new NoColor();
        return maskRect;
    }

    /**
     * グループへクリッピングマスクを適用する
     * @param {GroupItem} group - 対象グループ
     * @param {PageItem} sourceItem - マスクの元になる選択オブジェクト
     * @returns {void}
     */
    function applyMaskToGroup(group, sourceItem) {
        if (!group || !sourceItem) return;

        var maskItem = createMaskShapeFromSelection(group, sourceItem);
        if (!maskItem) return;

        /* クリッピングパスはグループの最前面に置く / A clipping path must sit at the front */
        try {
            maskItem.move(group, ElementPlacement.PLACEATEND);
            maskItem.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (eMoveMask) {
            logError(eMoveMask, "applyMaskToGroup.moveMask");
        }

        /* TextFrame は clipping を立てられないため makeMask を使う / Text frames need the makeMask command */
        if (maskItem.typename === "TextFrame") {
            try {
                doc.selection = null;
                group.selected = true;
                app.executeMenuCommand('makeMask');
            } catch (eMakeMask) {
                logError(eMakeMask, "applyMaskToGroup.makeMask");
            }
            return;
        }

        setClippingFlag(maskItem);
        try {
            group.clipped = true;
        } catch (eSetClipped) {
            logError(eSetClipped, "applyMaskToGroup.setGroupClipped");
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    var previewLayer = null;
    var debounceTaskId = 0;

    /**
     * プレビュー専用レイヤーを用意する
     * @returns {void}
     */
    function ensurePreviewLayer() {
        if (previewLayer) return;
        previewLayer = doc.layers.add();
        previewLayer.name = PREVIEW_LAYER_NAME;
    }

    /**
     * プレビューの中身を消す（グループ・マスクを含む）
     * @returns {void}
     */
    function clearPreview() {
        if (previewLayer) {
            while (previewLayer.pageItems.length > 0) {
                try {
                    previewLayer.pageItems[0].remove();
                } catch (eRemovePreviewItem) {
                    logError(eRemovePreviewItem, "clearPreview");
                    break;
                }
            }
        }
        app.redraw();
    }

    /**
     * 現在の設定でプレビューを描き直す
     * @returns {void}
     */
    function drawPreview() {
        ensurePreviewLayer();
        clearPreview();

        try {
            var confettiContainer = createConfettiContainer(previewLayer);
            var bounds = getEffectiveBounds();
            if (!bounds) return;

            for (var i = 0; i < confettiCount; i++) {
                var point = pickConfettiPoint(bounds);
                var color = CONFETTI_COLORS[Math.floor(Math.random() * CONFETTI_COLORS.length)];
                createConfetti(confettiContainer.container, point.x, point.y, getConfettiSize(), color);
            }
            if (confettiContainer.maskGroup && !useArtboardBounds) {
                applyMaskToGroup(confettiContainer.maskGroup, selectedItem);
            }
        } catch (eDrawPreview) {
            logError(eDrawPreview, "drawPreview");
        }
        app.redraw();
    }

    /* scheduleTask はグローバルスコープで動くため、呼び出し口を公開する / scheduleTask runs in the global scope */
    $.global.__ConfettiMaker_runDebouncedPreview = function () {
        try {
            drawPreview();
        } catch (eDebouncedPreview) {
            logError(eDebouncedPreview, "runDebouncedPreview");
        }
    };

    /**
     * 公開したグローバル関数を片付ける
     * @returns {void}
     */
    function cleanupScheduledGlobals() {
        try { delete $.global.__ConfettiMaker_runDebouncedPreview; } catch (_) { }
    }

    /**
     * 予約済みのプレビュー再描画を取り消す
     * @returns {void}
     */
    function cancelScheduledPreview() {
        if (!debounceTaskId) return;
        try { app.cancelTask(debounceTaskId); } catch (_) { }
        debounceTaskId = 0;
    }

    /**
     * プレビューの再描画を間引いて予約する（スライダードラッグ対策）
     * @returns {void}
     */
    function requestPreviewDebounced() {
        cancelScheduledPreview();
        try {
            debounceTaskId = app.scheduleTask("__ConfettiMaker_runDebouncedPreview()", PREVIEW_DEBOUNCE_MS, false);
        } catch (eScheduleTask) {
            logError(eScheduleTask, "requestPreviewDebounced.scheduleTask");
            drawPreview(); /* 予約できない環境では即時描画 / Draw immediately when scheduling is unavailable */
        }
    }

    // =========================================
    // イベント / Events
    // =========================================

    /**
     * 修飾キーの状態を読む（ScriptUI のイベントではなく keyboardState を使う）
     * @returns {object} { altKey, metaKey }
     */
    function getKeyboardState() {
        try {
            var keyboardState = ScriptUI.environment.keyboardState;
            return { altKey: !!keyboardState.altKey, metaKey: !!keyboardState.metaKey };
        } catch (_) {
            return { altKey: false, metaKey: false };
        }
    }

    /**
     * 指定した形状だけをONにする
     * @param {ShapeDef} activeShape - 単独で残す形状
     * @returns {void}
     */
    function setOnlyShape(activeShape) {
        for (var i = 0; i < shapeToggles.length; i++) {
            shapeToggles[i].checkbox.value = (shapeToggles[i] === activeShape);
        }
    }

    /**
     * 指定した形状だけをOFFにする（それ以外をON）
     * @param {ShapeDef} activeShape - OFFにする形状
     * @returns {void}
     */
    function setOnlyOtherShapes(activeShape) {
        for (var i = 0; i < shapeToggles.length; i++) {
            shapeToggles[i].checkbox.value = (shapeToggles[i] !== activeShape);
        }
    }

    /**
     * 回転をOFFにして0固定にする
     * @returns {void}
     */
    function disableRotate() {
        if (rotateMaxDeg > 0) previousRotateMaxDeg = rotateMaxDeg;
        rotateCheckbox.value = false;
        rotateMaxDeg = 0;
        rotateSlider.enabled = false;
        rotateSlider.value = 0;
    }

    /**
     * 歪みをOFFにして0固定にする
     * @returns {void}
     */
    function disableSkew() {
        skewCheckbox.value = false;
        skewMaxDeg = 0;
        skewSlider.enabled = false;
        skewSlider.value = 0;
    }

    /**
     * Option+クリックで単独選択したときのプリセットを適用する
     * @param {string} presetName - "noRotate" または "noRotateNoSkew"（それ以外は何もしない）
     * @returns {void}
     */
    function applySoloPreset(presetName) {
        if (presetName !== "noRotate" && presetName !== "noRotateNoSkew") return;
        disableRotate();
        if (presetName !== "noRotateNoSkew") return;

        disableSkew();
        randomSizeCheckbox.value = true;
        randomSizeSlider.enabled = true;
        randomSizeSlider.value = SOLO_RANDOM_SIZE;
        randomSizeStrength = SOLO_RANDOM_SIZE;
    }

    /* 形状チェックボックス（Option+クリックで単独選択、⌘+Option+クリックで反転）/ Shape toggles with solo and invert shortcuts */
    for (var i = 0; i < shapeToggles.length; i++) {
        (function (shapeToggle) {
            shapeToggle.checkbox.onClick = function () {
                var keyboardState = getKeyboardState();
                if (keyboardState.metaKey && keyboardState.altKey) {
                    setOnlyOtherShapes(shapeToggle);
                } else if (keyboardState.altKey) {
                    setOnlyShape(shapeToggle);
                    applySoloPreset(shapeToggle.soloPreset);
                }
                drawPreview();
            };
        })(shapeToggles[i]);
    }

    symbolDropdown.onChange = function () {
        if (!symbolDropdown.selection || symbolDropdown.selection.index === 0) {
            /* （なし）を選んだらシンボル形状をディム / Dim the symbol shape when "(None)" is picked */
            selectedSymbolName = "";
            selectedSymbolRef = null;
            symbolCheckbox.value = false;
            symbolCheckbox.enabled = false;
        } else {
            selectedSymbolName = String(symbolDropdown.selection.text);
            selectedSymbolRef = symbolDropdown.selection._symbolDef || null;
            symbolCheckbox.enabled = true;
            symbolCheckbox.value = true; /* 選んだ時点で自動でON / Turn it on as soon as a symbol is chosen */
        }
        drawPreview();
    };

    baseSizeSlider.onChanging = function () {
        requestPreviewDebounced();
    };

    countSlider.onChanging = function () {
        confettiCount = Math.round(countSlider.value);
        requestPreviewDebounced();
    };

    maskCheckbox.onClick = function () {
        drawPreview();
    };

    marginCheckbox.onClick = function () {
        marginSlider.enabled = marginCheckbox.value;
        if (marginCheckbox.value) {
            /* ONにしたときは少し余白が付いた状態から始める / Start with a visible margin */
            var newValue = Number(marginSlider.value) + MARGIN_STEP_ON_ENABLE;
            if (isNaN(newValue)) newValue = MARGIN_STEP_ON_ENABLE;
            marginSlider.value = Math.min(newValue, Number(marginSlider.maxvalue));
            generationMarginPt = Math.round(Number(marginSlider.value));
            /* マージンONとマスク処理は両立させない / Margin and masking are mutually exclusive */
            if (maskCheckbox.enabled) maskCheckbox.value = false;
        }
        requestPreviewDebounced();
    };

    marginSlider.onChanging = function () {
        generationMarginPt = Math.round(marginSlider.value);
        requestPreviewDebounced();
    };

    evenRadio.onClick = function () {
        strengthRow.enabled = false;
        drawPreview();
    };

    verticalRadio.onClick = function () {
        strengthRow.enabled = true;
        drawPreview();
    };

    radialRadio.onClick = function () {
        strengthRow.enabled = true;
        drawPreview();
    };

    strengthSlider.onChanging = function () {
        distributionStrength = Math.round(strengthSlider.value * 10) / 10; /* 0.1刻み / step of 0.1 */
        requestPreviewDebounced();
    };

    randomSizeCheckbox.onClick = function () {
        randomSizeSlider.enabled = randomSizeCheckbox.value;
        drawPreview();
    };

    randomSizeSlider.onChanging = function () {
        randomSizeStrength = Math.round(randomSizeSlider.value);
        if (randomSizeCheckbox.value) requestPreviewDebounced();
    };

    opacityCheckbox.onClick = function () {
        opacitySlider.enabled = opacityCheckbox.value;
        drawPreview();
    };

    opacitySlider.onChanging = function () {
        /* スライダーは反転指定 / The slider is reversed */
        opacityMin = Math.min(Math.max(100 - Math.round(opacitySlider.value), 0), 100);
        if (opacityCheckbox.value) requestPreviewDebounced();
    };

    skewCheckbox.onClick = function () {
        skewSlider.enabled = skewCheckbox.value;
        drawPreview();
    };

    skewSlider.onChanging = function () {
        skewMaxDeg = Math.min(Math.max(Math.round(skewSlider.value), 0), SKEW_MAX_DEG);
        if (skewCheckbox.value) requestPreviewDebounced();
    };

    rotateCheckbox.onClick = function () {
        rotateSlider.enabled = rotateCheckbox.value;
        if (!rotateCheckbox.value) {
            disableRotate();
        } else {
            /* ONに戻したときは前回値（なければ既定値）へ復元 / Restore the previous amount */
            if (rotateMaxDeg <= 0) rotateMaxDeg = (previousRotateMaxDeg > 0) ? previousRotateMaxDeg : DEFAULT_ROTATE_MAX;
            rotateSlider.enabled = true;
            rotateSlider.value = rotateMaxDeg;
        }
        drawPreview();
    };

    rotateSlider.onChanging = function () {
        rotateMaxDeg = Math.min(Math.max(Math.round(rotateSlider.value), 0), 360);
        if (rotateCheckbox.value) requestPreviewDebounced();
    };

    dialog.onShow = function () {
        shiftDialogPosition(dialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        refreshSymbolDropdown();
        applyMarginMaxToUI();
        zoomControls.syncFromView();
        drawPreview();
    };

    // =========================================
    // 実行 / Run
    // =========================================

    var dialogResult = dialog.show();
    cancelScheduledPreview();

    if (dialogResult !== 1) {
        clearPreview();
        try { if (previewLayer) previewLayer.remove(); } catch (eRemovePreviewLayer) { logError(eRemovePreviewLayer, "cancel.removePreviewLayer"); }
        zoomControls.restoreInitial();
        cleanupScheduledGlobals();
        return;
    }

    /* プレビューの見た目そのままで確定する（再生成しない）/ Keep exactly what the preview shows */
    if (previewLayer && previewLayer.pageItems.length === 0) drawPreview();

    /**
     * 出力先レイヤーを取得する（既存があれば再利用）
     * @returns {Layer} 出力先レイヤー
     */
    function getOutputLayer() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === OUTPUT_LAYER_NAME) return doc.layers[i];
        }
        var createdLayer = doc.layers.add();
        createdLayer.name = OUTPUT_LAYER_NAME;
        return createdLayer;
    }

    var confettiLayer = getOutputLayer();

    /* マスクOFF時は、この実行で生成したものだけをグループ化する / Group this run's output when masking is off */
    var confettiGroup = null;
    if (!maskCheckbox.value) {
        confettiGroup = confettiLayer.groupItems.add();
        confettiGroup.name = OUTPUT_GROUP_NAME;
    }

    /* プレビューの中身を丸ごと移動（マスクグループ含む）/ Move the whole preview, clipping group included */
    var movedItems = [];
    try {
        while (previewLayer && previewLayer.pageItems.length > 0) {
            var previewItem = previewLayer.pageItems[0];
            try {
                previewItem.move(confettiGroup || confettiLayer, ElementPlacement.PLACEATEND);
                movedItems.push(previewItem);
            } catch (eMoveItem) {
                logError(eMoveItem, "finalize.moveItem");
                try { previewItem.remove(); } catch (_) { break; }
            }
        }

        /* 既存アイテムより生成物が前面になるようにする / Bring the confetti in front of the existing artwork */
        if (confettiGroup) {
            confettiGroup.zOrder(ZOrderMethod.BRINGTOFRONT);
        } else {
            for (var zi = 0; zi < movedItems.length; zi++) {
                movedItems[zi].zOrder(ZOrderMethod.BRINGTOFRONT);
            }
        }
    } catch (eFinalizeMove) {
        logError(eFinalizeMove, "finalize.movePreviewItems");
    }

    try { if (previewLayer) previewLayer.remove(); } catch (eRemoveLayer) { logError(eRemoveLayer, "finalize.removePreviewLayer"); }

    /* 生成したコンフェティ全体を選択状態にする / Leave the generated confetti selected */
    try {
        doc.selection = null;
        if (confettiGroup) {
            confettiGroup.selected = true;
        } else {
            for (var si = 0; si < confettiLayer.pageItems.length; si++) {
                confettiLayer.pageItems[si].selected = true;
            }
        }
    } catch (eSelectResult) {
        logError(eSelectResult, "finalize.selectResult");
    }

    /* 画面ズームはプレビュー中の操作を尊重して復元しない / Keep whatever zoom the user landed on */
    cleanupScheduledGlobals();

})();
