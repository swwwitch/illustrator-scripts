#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ダイアログで方向・位置・単位・対象（カンバス／アートボード）を指定してガイドを作成します。

詳細は README を参照してください。

### Overview

Creates guides by specifying direction, position, unit, and target (canvas or artboard) in a dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "NewGuideMaker";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/NewGuideMaker.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/NewGuideMaker.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1085336d7265"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ガイドを作成するレイヤー名 / Layer that receives the guides */
    var GUIDE_LAYER_NAME = "_guide";

    /* カンバス端まで届く十分な長さ（Illustrator の最大カンバス 227inch 相当）/ Length long enough to span the canvas (227 inch ≈ Illustrator max canvas) */
    var CANVAS_SPAN_PT = 227 * 72;

    /* プレビュー線の色（CMYK ドキュメント用）/ Preview stroke color for CMYK documents */
    var PREVIEW_COLOR_CMYK = { cyan: 70, magenta: 50, yellow: 0, black: 0 };

    /* プレビュー線の色（CMYK 以外へのフォールバック）/ Preview stroke color for non-CMYK documents */
    var PREVIEW_COLOR_RGB = { red: 74, green: 132, blue: 255 };

    /* プレビュー線の太さ（pt）/ Stroke width of the preview paths (pt) */
    var PREVIEW_STROKE_WIDTH = 1.0;

    /* ガイド化後の線幅（pt）/ Stroke width once converted to a guide (pt) */
    var GUIDE_STROKE_WIDTH = 0.1;

    /* リピートで一度に作れるガイドの上限（桁の打ち間違いで固まるのを防ぐ）/ Cap on repeated guides, so a mistyped digit cannot freeze Illustrator */
    var MAX_REPEAT_COUNT = 1000;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;               /* 2カラムの間隔 / gap between columns */
    var FIELD_ROW_SPACING  = 6;                /* ラベル・入力欄・単位表記の間隔 / gap inside a labeled field row */
    var UNIT_TEXT_WIDTH    = 34;               /* 数値欄に添える単位表記の幅 / width of the unit label next to a field */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */
    var BUTTON_BAR_SPACING = 10;               /* ボタンバー内グループの要素間隔 / spacing inside the button bar groups */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の表示言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        /* "ja" で始まるロケール（ja, ja_JP など）は日本語扱い / Treat "ja*" locales as Japanese */
        if (localeText.indexOf("ja") === 0) {
            return "ja";
        }
        return "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ガイド作成", en: "Create Guide" }
        },
        target: {
            panelTitle: { ja: "対象", en: "Target" },
            canvas:     { ja: "カンバス", en: "Canvas" },
            artboard:   { ja: "アートボード", en: "Artboard" },
            extension:  { ja: "延長", en: "Extension" }
        },
        direction: {
            panelTitle: { ja: "方向", en: "Direction" },
            horizontal: { ja: "水平方向", en: "Horizontal" },
            vertical:   { ja: "垂直方向", en: "Vertical" },
            position:   { ja: "開始位置", en: "Start Position" }
        },
        layer: {
            panelTitle:  { ja: "作成レイヤー", en: "Target Layer" },
            guideLayer:  { ja: "_guideレイヤー", en: "_guide Layer" },
            activeLayer: { ja: "現在のレイヤー", en: "Current Layer" }
        },
        repeat: {
            panelTitle: { ja: "リピート", en: "Repeat" },
            count:      { ja: "ガイド数", en: "Guide Count" },
            distance:   { ja: "距離", en: "Distance" }
        },
        unit: {
            fieldLabel: { ja: "単位", en: "Unit" }
        },
        tooltip: {
            extension: { ja: "ガイドをアートボードの外側へ伸ばす量（アートボード対象時のみ）", en: "How far to extend guides beyond the artboard (artboard target only)" },
            position:  { ja: "ガイドの開始位置。↑↓で増減、Shift+↑↓で10単位スナップ", en: "Guide start position. Up/Down to step, Shift+Up/Down snaps to 10" },
            count:     { ja: "作成するガイドの本数", en: "Number of guides to create" },
            distance:  { ja: "リピート時のガイドの間隔", en: "Spacing between repeated guides" },
            direction: { ja: "H / V キーでも切り替えできます", en: "Toggle with the H / V keys too" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            lockedLayer: { ja: "アクティブレイヤーがロックされています。", en: "The active layer is locked." },
            noDocument:  { ja: "ドキュメントが開かれていません。", en: "No document is open." }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('target','canvas')）
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
     * ラベル末尾に付けるコロンを返す（日本語は全角、英語は半角）
     * @returns {string} コロン記号
     */
    function getUiColon() {
        return (uiLang === "ja") ? "：" : ":";
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; array index equals the rulerType code (0=in, 1=mm, 2=pt …) */
    var UNITS = [
        { label: "in",    factor: 72.0 },                /* 0 */
        { label: "mm",    factor: 72.0 / 25.4 },         /* 1 */
        { label: "pt",    factor: 1.0 },                 /* 2 */
        { label: "pica",  factor: 12.0 },                /* 3 */
        { label: "cm",    factor: 72.0 / 2.54 },         /* 4 */
        { label: "Q/H",   factor: 72.0 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    factor: 1.0 },                 /* 6 */
        { label: "ft/in", factor: 72.0 * 12.0 },         /* 7 */
        { label: "m",     factor: 72.0 / 25.4 * 1000.0 },/* 8 */
        { label: "yd",    factor: 72.0 * 36.0 },         /* 9 */
        { label: "ft",    factor: 72.0 * 12.0 }          /* 10 */
    ];

    /* pt の添字（単位が特定できないときのフォールバック）/ Index of pt, used as the fallback unit */
    var POINT_UNIT_INDEX = 2;

    /**
     * 単位ラベルの一覧を返す（ドロップダウン用）
     * @returns {string[]} 単位ラベルの配列
     */
    function getUnitLabels() {
        var labelList = [];
        for (var i = 0; i < UNITS.length; i++) {
            labelList.push(UNITS[i].label);
        }
        return labelList;
    }

    /**
     * ルーラー環境設定の単位インデックスを取得する（= rulerType コード）
     * @returns {number} UNITS の添字（範囲外なら POINT_UNIT_INDEX）
     */
    function getRulerUnitIndex() {
        var rulerTypeCode = app.preferences.getIntegerPreference("rulerType");
        return (rulerTypeCode >= 0 && rulerTypeCode < UNITS.length) ? rulerTypeCode : POINT_UNIT_INDEX;
    }

    /**
     * 値と単位ラベルから pt へ変換する
     * @param {string|number} inputValue - 変換する値（数値以外は0扱い）
     * @param {string} unitLabel - 単位ラベル（"mm" など）
     * @returns {number} pt に変換した値
     */
    function convertToPt(inputValue, unitLabel) {
        var numericValue = Number(inputValue);
        if (isNaN(numericValue)) {
            return 0;
        }
        for (var i = 0; i < UNITS.length; i++) {
            if (UNITS[i].label === unitLabel) {
                return numericValue * UNITS[i].factor;
            }
        }
        return numericValue; /* 見つからなければ pt 扱い / Fall back to pt */
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * パネルに共通レイアウトを適用する
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
     * 左寄せの縦並びグループを生成する（ラジオ列など）
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @returns {Group} 生成したグループ
     */
    function addLeftAlignedColumn(parentContainer) {
        var createdGroup = parentContainer.add("group");
        createdGroup.orientation = "column";
        createdGroup.alignChildren = ["left", "center"];
        return createdGroup;
    }

    /**
     * 2択のラジオボタン列を生成する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} firstLabel - 1つ目のラベル
     * @param {string} secondLabel - 2つ目のラベル
     * @param {number} selectedIndex - 初期選択（0=1つ目、1=2つ目）
     * @returns {RadioButton[]} [1つ目, 2つ目] のラジオボタン
     */
    function addRadioPair(parentContainer, firstLabel, secondLabel, selectedIndex) {
        var radioColumn = addLeftAlignedColumn(parentContainer);
        var firstRadio = radioColumn.add("radiobutton", undefined, firstLabel);
        var secondRadio = radioColumn.add("radiobutton", undefined, secondLabel);
        ((selectedIndex === 1) ? secondRadio : firstRadio).value = true;
        return [firstRadio, secondRadio];
    }

    /**
     * 幅いっぱいに広がる縦並びカラムを生成する（パネルを積む用）
     * @param {Group} parentRow - 追加先の行グループ
     * @returns {Group} 生成したカラム
     */
    function addSettingsColumn(parentRow) {
        var createdColumn = parentRow.add("group");
        createdColumn.orientation = "column";
        createdColumn.alignChildren = ["fill", "top"];
        createdColumn.spacing = WINDOW_SPACING;
        return createdColumn;
    }

    /**
     * ダイアログウィンドウを生成する（上段に内容、下段にボタンの縦構成）
     * @param {string} windowTitle - ウィンドウタイトル
     * @returns {Window} 生成したダイアログ
     */
    function createDialogWindow(windowTitle) {
        var dialogWindow = new Window("dialog", windowTitle);
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];
        dialogWindow.spacing = WINDOW_SPACING;
        dialogWindow.margins = WINDOW_MARGINS;
        return dialogWindow;
    }

    // =========================================
    // 入力欄の補助 / Input field helpers
    // =========================================

    /**
     * 入力欄に↑↓キーでの値増減を追加する（Shift併用で10単位スナップ）
     * @param {EditText} inputField - 対象の入力欄
     * @param {function} [onValueChanged] - 値を更新したあとに呼ぶコールバック
     * @returns {void}
     */
    function changeValueByArrowKey(inputField, onValueChanged) {
        inputField.addEventListener("keydown", function(event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;
            var currentValue = Number(inputField.text);
            if (isNaN(currentValue)) return;

            /* 修飾キーは event から読む（keyboardState は macOS で誤報あり）/ Read the modifier from event (keyboardState misreports on macOS) */
            var shiftPressed = event.shiftKey;
            if (shiftPressed === undefined) {
                shiftPressed = ScriptUI.environment.keyboardState.shiftKey;
            }

            var stepDirection = (event.keyName == "Up") ? 1 : -1;
            if (shiftPressed) {
                /* Shift押下時は「10の倍数」スナップ / Snap to multiples of 10 when Shift is pressed */
                currentValue = Math.round(currentValue / 10) * 10 + stepDirection * 10;
            } else {
                currentValue += stepDirection;
            }

            event.preventDefault();
            inputField.text = currentValue;
            if (typeof onValueChanged === "function") {
                onValueChanged(inputField.text);
            }
        });
    }

    // =========================================
    // ガイド座標と外観 / Guide geometry and appearance
    // =========================================

    /**
     * ガイド1本分の始点・終点座標を返す（カンバスはドキュメント原点基準＝既定のルーラー0点）
     * @param {boolean} isCanvasTarget - カンバス基準なら true、アートボード基準なら false
     * @param {boolean} isHorizontal - 水平ガイドなら true、垂直ガイドなら false
     * @param {number} positionPt - ガイド位置（pt）
     * @param {number} extensionPt - アートボード外への延長量（pt）
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]（カンバス基準では未使用）
     * @returns {number[][]} 始点・終点の座標配列
     */
    function getGuidePathPoints(isCanvasTarget, isHorizontal, positionPt, extensionPt, artboardRect) {
        if (isCanvasTarget) {
            /* Y は上方向が正なので、下向きの位置は減算 / Y is up, so a downward position subtracts */
            return isHorizontal
                ? [[-CANVAS_SPAN_PT, -positionPt], [CANVAS_SPAN_PT, -positionPt]]
                : [[positionPt, CANVAS_SPAN_PT], [positionPt, -CANVAS_SPAN_PT]];
        }
        var artboardLeft = artboardRect[0];
        var artboardTop = artboardRect[1];
        var artboardRight = artboardRect[2];
        var artboardBottom = artboardRect[3];
        return isHorizontal
            ? [[artboardLeft - extensionPt, artboardTop - positionPt], [artboardRight + extensionPt, artboardTop - positionPt]]
            : [[artboardLeft + positionPt, artboardTop + extensionPt], [artboardLeft + positionPt, artboardBottom - extensionPt]];
    }

    /**
     * プレビュー線に使う色を生成する
     * @param {DocumentColorSpace} docColorSpace - ドキュメントのカラースペース
     * @returns {CMYKColor|RGBColor} プレビュー線の色
     */
    function createPreviewColor(docColorSpace) {
        if (docColorSpace === DocumentColorSpace.CMYK) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = PREVIEW_COLOR_CMYK.cyan;
            cmykColor.magenta = PREVIEW_COLOR_CMYK.magenta;
            cmykColor.yellow = PREVIEW_COLOR_CMYK.yellow;
            cmykColor.black = PREVIEW_COLOR_CMYK.black;
            return cmykColor;
        }
        /* CMYK 以外は青の RGB にフォールバック / Fall back to a blue RGB for non-CMYK modes */
        var rgbColor = new RGBColor();
        rgbColor.red = PREVIEW_COLOR_RGB.red;
        rgbColor.green = PREVIEW_COLOR_RGB.green;
        rgbColor.blue = PREVIEW_COLOR_RGB.blue;
        return rgbColor;
    }

    /**
     * プレビュー線の見た目を設定する
     * @param {PathItem} previewPath - 対象パス
     * @param {CMYKColor|RGBColor} previewColor - 線の色
     * @returns {void}
     */
    function stylePreviewPath(previewPath, previewColor) {
        previewPath.stroked = true;
        previewPath.filled = false;
        previewPath.strokeWidth = PREVIEW_STROKE_WIDTH;
        previewPath.strokeColor = previewColor;
        previewPath.guides = false; /* プレビューはガイド化しない / Preview is not a guide */
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * @typedef {object} PreviewSettings
     * @property {boolean} isCanvasTarget - カンバス基準なら true
     * @property {boolean} isHorizontal - 水平ガイドなら true
     * @property {number} positionPt - 1本目のガイド位置（pt）
     * @property {number} extensionPt - アートボード外への延長量（pt）
     * @property {number[]} artboardRect - アートボードの矩形（カンバス基準では null）
     * @property {number} repeatCount - 作成するガイドの本数
     * @property {number} repeatDistancePt - リピート間隔（pt）
     */

    /**
     * ガイド作成ダイアログを構築する
     * @returns {Window} 構築済みのダイアログ
     */
    function createGuideDialog() {
        var doc = app.activeDocument;

        /* 単位オプションと現在単位のインデックス / Unit options and current unit index */
        var unitOptions = getUnitLabels();
        var currentUnitIndex = getRulerUnitIndex();
        var unitSuffixTexts = []; /* 各数値欄の単位表記（共有ドロップダウンに追従）/ Per-field unit labels (follow the shared dropdown) */

        /* ガイド用レイヤーは使用時のみ遅延生成 / The guide layer is created lazily, only when used */
        var guideLayer = null;
        var guideLayerWasCreated = false; /* このスクリプトで新規作成したか / Whether this run created the layer */
        var guideLayerWasLocked = false;  /* 元のロック状態 / The layer's lock state before this run */
        var guidesCommitted = false;      /* OKでガイド化を確定したか / Whether OK committed the guides */

        /* ロック警告は選択ごとに1回だけ出す / Warn about a locked layer only once per selection */
        var lockedLayerAlerted = false;

        /* 現在描画中のプレビュー線（常に配列で保持）/ Currently drawn preview paths (always an array) */
        var activePreviewPaths = null;

        /* ダイアログ内で共有する主要コントロール / Controls shared across the dialog's handlers */
        var canvasRadio, artboardRadio, extensionRow, extensionInput;
        var guideLayerRadio, activeLayerRadio;
        var horizontalRadio, verticalRadio, positionInput;
        var repeatCountInput, repeatDistanceInput;
        var unitDropdown;

        /**
         * ガイド用レイヤーを取得する（なければ作成）
         * @returns {Layer} ガイド用レイヤー
         */
        function ensureGuideLayer() {
            if (guideLayer) {
                return guideLayer;
            }
            var docLayers = doc.layers;
            for (var i = 0; i < docLayers.length; i++) {
                if (docLayers[i].name === GUIDE_LAYER_NAME) {
                    guideLayer = docLayers[i];
                    guideLayerWasLocked = guideLayer.locked;
                    return guideLayer;
                }
            }
            guideLayer = docLayers.add();
            guideLayer.name = GUIDE_LAYER_NAME;
            guideLayerWasCreated = true;
            return guideLayer;
        }

        /**
         * ガイド化せずに閉じたとき、ガイド用レイヤーを元の状態に戻す
         * （このスクリプトで作った空レイヤーは削除、既存レイヤーはロック状態を復元）
         * @returns {void}
         */
        function restoreGuideLayer() {
            if (!guideLayer) return;
            guideLayer.locked = false;
            if (guideLayerWasCreated) {
                /* 空のまま残さない（最後の1枚は削除できないので残す）/ Do not leave an empty layer behind (the last layer cannot be removed) */
                if (guideLayer.pageItems.length === 0 && doc.layers.length > 1) {
                    guideLayer.remove();
                }
            } else {
                guideLayer.locked = guideLayerWasLocked;
            }
            guideLayer = null;
        }

        /**
         * プレビュー線を削除する（モーダル中はドキュメント編集不可なのでパスは常に有効）
         * @returns {void}
         */
        function removePreviewPaths() {
            if (!activePreviewPaths) return;
            for (var i = 0; i < activePreviewPaths.length; i++) {
                var previewPath = activePreviewPaths[i];
                if (previewPath && !previewPath.locked && previewPath.layer && !previewPath.layer.locked) {
                    previewPath.remove();
                }
            }
            activePreviewPaths = null;
        }

        /**
         * プレビューの描画先レイヤーを決定する
         * @returns {Layer|null} 描画先レイヤー（アクティブレイヤーがロック中なら null）
         */
        function resolvePreviewLayer() {
            if (guideLayerRadio.value) {
                var targetGuideLayer = ensureGuideLayer();
                targetGuideLayer.locked = false;
                lockedLayerAlerted = false;
                return targetGuideLayer;
            }
            var activeLayer = doc.activeLayer;
            if (activeLayer.locked) {
                /* プレビューは打鍵ごとに走るので、警告は選択ごとに1回だけ / The preview runs on every keystroke, so warn only once per selection */
                if (!lockedLayerAlerted) {
                    lockedLayerAlerted = true;
                    alert(getLabel('alert', 'lockedLayer'));
                }
                return null;
            }
            lockedLayerAlerted = false;
            return activeLayer;
        }

        /**
         * 入力欄の内容をプレビュー用の設定値に読み取る
         * @returns {PreviewSettings|null} 設定値（数値として読めない欄があれば null）
         */
        function readPreviewSettings() {
            /* 単位は全フィールド共通 / A single shared unit for all fields */
            var unitLabel = unitDropdown.selection.text;

            var positionValue = parseFloat(positionInput.text);
            if (isNaN(positionValue)) {
                return null;
            }

            var repeatCount = parseInt(repeatCountInput.text, 10);
            if (isNaN(repeatCount) || repeatCount < 1) repeatCount = 1;
            /* 桁の打ち間違いで大量生成しないよう上限で丸める / Clamp so a mistyped digit cannot spawn a huge batch */
            if (repeatCount > MAX_REPEAT_COUNT) repeatCount = MAX_REPEAT_COUNT;
            var repeatDistancePt = convertToPt(repeatDistanceInput.text, unitLabel);
            /* 距離0以下なら重複を避けて1本に / Avoid overlapping guides when distance is 0 or less */
            if (repeatDistancePt <= 0) {
                repeatDistancePt = 0;
                repeatCount = 1;
            }

            /* アートボード対象時の延長量と矩形を一度だけ算出 / Compute extension amount and rect once when targeting the artboard */
            var isCanvasTarget = canvasRadio.value;
            var extensionPt = 0;
            var artboardRect = null;
            if (!isCanvasTarget) {
                var extensionValue = parseFloat(extensionInput.text);
                if (isNaN(extensionValue)) {
                    return null;
                }
                extensionPt = convertToPt(extensionValue, unitLabel);
                artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            }

            return {
                isCanvasTarget: isCanvasTarget,
                isHorizontal: horizontalRadio.value,
                positionPt: convertToPt(positionValue, unitLabel),
                extensionPt: extensionPt,
                artboardRect: artboardRect,
                repeatCount: repeatCount,
                repeatDistancePt: repeatDistancePt
            };
        }

        /**
         * 設定値どおりにプレビュー線を描画する
         * @param {PreviewSettings} settings - 入力欄から読み取った設定値
         * @param {Layer} targetLayer - 描画先レイヤー
         * @returns {PathItem[]} 描画したプレビュー線
         */
        function drawPreviewPaths(settings, targetLayer) {
            var previewColor = createPreviewColor(doc.documentColorSpace);
            var drawnPaths = [];
            for (var i = 0; i < settings.repeatCount; i++) {
                var currentPositionPt = settings.positionPt + i * settings.repeatDistancePt;
                var previewPath = targetLayer.pathItems.add();
                previewPath.setEntirePath(getGuidePathPoints(
                    settings.isCanvasTarget,
                    settings.isHorizontal,
                    currentPositionPt,
                    settings.extensionPt,
                    settings.artboardRect
                ));
                stylePreviewPath(previewPath, previewColor);
                drawnPaths.push(previewPath);
            }
            return drawnPaths;
        }

        /**
         * 現在の入力値でプレビュー線を描き直す
         * @returns {void}
         */
        function drawPreview() {
            removePreviewPaths();
            var settings = readPreviewSettings();
            if (!settings) {
                return;
            }
            var targetLayer = resolvePreviewLayer();
            if (!targetLayer) {
                return;
            }
            activePreviewPaths = drawPreviewPaths(settings, targetLayer);
            app.redraw();
        }

        /**
         * プレビュー線をガイドに変換する
         * @returns {void}
         */
        function convertPreviewToGuides() {
            if (!activePreviewPaths) return;
            /* ガイド化のため対象レイヤーのロックを一時解除 / Temporarily unlock so paths can be converted */
            if (guideLayer && guideLayer.locked) {
                guideLayer.locked = false;
            }
            for (var i = 0; i < activePreviewPaths.length; i++) {
                var previewPath = activePreviewPaths[i];
                if (previewPath && previewPath.layer && !previewPath.layer.locked) {
                    previewPath.guides = true;
                    previewPath.strokeWidth = GUIDE_STROKE_WIDTH;
                }
            }
        }

        /**
         * 数値入力欄を生成する（↑↓キーのステップと逐次プレビュー付き）
         * @param {Group} parentRow - 追加先の行グループ
         * @param {string} defaultValue - 初期値
         * @param {number} [widthInChars] - 入力欄の幅（文字数、省略時は2）
         * @returns {EditText} 生成した入力欄
         */
        function addNumberField(parentRow, defaultValue, widthInChars) {
            var numberField = parentRow.add('edittext {characters: ' + ((typeof widthInChars === "number") ? widthInChars : 2) + '}');
            numberField.text = defaultValue;
            changeValueByArrowKey(numberField, drawPreview);
            numberField.addEventListener("changing", drawPreview);
            return numberField;
        }

        /**
         * 数値欄の右に単位表記を追加する（共有ドロップダウンに追従）
         * @param {Group} parentRow - 追加先の行グループ
         * @returns {StaticText} 生成した単位表記
         */
        function addUnitSuffixText(parentRow) {
            var unitText = parentRow.add("statictext", undefined, unitOptions[currentUnitIndex]);
            unitText.preferredSize.width = UNIT_TEXT_WIDTH;
            unitSuffixTexts.push(unitText);
            return unitText;
        }

        /**
         * ラベル付きの数値欄を生成する
         * @param {Group|Panel} parentContainer - 追加先
         * @param {string} labelText - ラベル文字列（コロンは自動付与）
         * @param {string} defaultValue - 初期値
         * @param {boolean} showUnitSuffix - 単位表記を添えるかどうか
         * @param {string} [helpTipText] - ツールチップ
         * @param {number} [widthInChars] - 入力欄の幅（文字数）
         * @returns {{row: Group, input: EditText}} 行グループと入力欄
         */
        function addLabeledField(parentContainer, labelText, defaultValue, showUnitSuffix, helpTipText, widthInChars) {
            var fieldRow = parentContainer.add("group");
            setupRow(fieldRow, "left", FIELD_ROW_SPACING);
            var fieldLabel = fieldRow.add("statictext", undefined, labelText + getUiColon());
            var numberField = addNumberField(fieldRow, defaultValue, widthInChars);
            if (showUnitSuffix) {
                addUnitSuffixText(fieldRow);
            }
            if (helpTipText) {
                fieldLabel.helpTip = helpTipText;
                numberField.helpTip = helpTipText;
            }
            return { row: fieldRow, input: numberField };
        }

        /**
         * 対象（カンバス／アートボード）パネルを組み立てる
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function buildTargetPanel(parentColumn) {
            var targetPanel = addPanel(parentColumn, getLabel('target', 'panelTitle'));
            var targetRadios = addRadioPair(targetPanel, getLabel('target', 'canvas'), getLabel('target', 'artboard'), 1);
            canvasRadio = targetRadios[0];
            artboardRadio = targetRadios[1];

            /* 延長：ガイドをアートボード外へ伸ばす量 / Extension: how far to extend guides beyond the artboard */
            var extensionField = addLabeledField(targetPanel, getLabel('target', 'extension'), "0", true, getLabel('tooltip', 'extension'));
            extensionRow = extensionField.row;
            extensionInput = extensionField.input;
        }

        /**
         * 作成レイヤーパネルを組み立てる
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function buildLayerPanel(parentColumn) {
            var layerPanel = addPanel(parentColumn, getLabel('layer', 'panelTitle'));
            var layerRadios = addRadioPair(layerPanel, getLabel('layer', 'guideLayer'), getLabel('layer', 'activeLayer'), 0);
            guideLayerRadio = layerRadios[0];
            activeLayerRadio = layerRadios[1];
        }

        /**
         * 方向パネルを組み立てる
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function buildDirectionPanel(parentColumn) {
            var directionPanel = addPanel(parentColumn, getLabel('direction', 'panelTitle'));
            directionPanel.helpTip = getLabel('tooltip', 'direction');
            var directionRadios = addRadioPair(directionPanel, getLabel('direction', 'horizontal'), getLabel('direction', 'vertical'), 0);
            horizontalRadio = directionRadios[0];
            verticalRadio = directionRadios[1];
            positionInput = addLabeledField(directionPanel, getLabel('direction', 'position'), "0", true, getLabel('tooltip', 'position')).input;
        }

        /**
         * リピートパネルを組み立てる
         * @param {Group} parentColumn - 追加先のカラム
         * @returns {void}
         */
        function buildRepeatPanel(parentColumn) {
            var repeatPanel = addPanel(parentColumn, getLabel('repeat', 'panelTitle'));
            var repeatColumn = addLeftAlignedColumn(repeatPanel);
            repeatCountInput = addLabeledField(repeatColumn, getLabel('repeat', 'count'), "1", false, getLabel('tooltip', 'count')).input;
            repeatDistanceInput = addLabeledField(repeatColumn, getLabel('repeat', 'distance'), "0", true, getLabel('tooltip', 'distance'), 3).input;
        }

        /**
         * 下部のボタンバー（左＝単位、右＝キャンセル＋OK）を組み立てる
         * @param {Window} dialogWindow - 追加先のダイアログ
         * @returns {Button} OKボタン
         */
        function buildButtonBar(dialogWindow) {
            var buttonBarGroup = dialogWindow.add("group");
            setupRow(buttonBarGroup, "fill");
            buttonBarGroup.margins = BUTTON_BAR_MARGINS;

            /* 左側グループ：単位 / Left-side group: unit */
            var unitSelectGroup = buttonBarGroup.add("group");
            setupRow(unitSelectGroup, "left", BUTTON_BAR_SPACING);
            unitSelectGroup.add("statictext", undefined, getLabel('unit', 'fieldLabel') + getUiColon());
            unitDropdown = unitSelectGroup.add("dropdownlist", undefined, unitOptions);
            unitDropdown.selection = currentUnitIndex;
            unitDropdown.onChange = function() {
                /* 各数値欄の単位表記を更新 / Update the per-field unit labels */
                var selectedUnitLabel = unitDropdown.selection.text;
                for (var i = 0; i < unitSuffixTexts.length; i++) {
                    unitSuffixTexts[i].text = selectedUnitLabel;
                }
                drawPreview();
            };

            /* スペーサー（伸縮）/ Spacer (stretchable) */
            var buttonBarSpacer = buttonBarGroup.add("group");
            buttonBarSpacer.alignment = ["fill", "fill"];
            buttonBarSpacer.minimumSize.width = 0;

            /* 右側グループ：キャンセル＋OK（Mac 規約で Cancel → OK）/ Right-side group: Cancel + OK (Cancel → OK per macOS) */
            var dialogButtonGroup = buttonBarGroup.add("group");
            setupRow(dialogButtonGroup, "right", BUTTON_BAR_SPACING);
            /* キャンセルは既定動作で閉じ、後片付けは dialog.onClose が行う / Cancel closes by default; cleanup happens in dialog.onClose */
            dialogButtonGroup.add("button", undefined, getLabel('button', 'cancel'), { name: "cancel" });
            return dialogButtonGroup.add("button", undefined, getLabel('button', 'ok'), { name: "ok" });
        }

        var dialog = createDialogWindow(getLabel('dialog', 'title') + ' ' + SCRIPT_VERSION);

        /* 上段：2カラムを横並びに収める行 / Top area: a row holding the two columns */
        var columnsRow = dialog.add("group");
        columnsRow.orientation = "row";
        columnsRow.alignChildren = ["fill", "top"];
        columnsRow.spacing = COLUMN_SPACING;

        /* 左カラム：対象・作成レイヤー / Left column: target, layer */
        var leftColumn = addSettingsColumn(columnsRow);
        buildTargetPanel(leftColumn);
        buildLayerPanel(leftColumn);

        /* 右カラム：方向・リピート / Right column: direction, repeat */
        var rightColumn = addSettingsColumn(columnsRow);
        buildDirectionPanel(rightColumn);
        buildRepeatPanel(rightColumn);

        var okButton = buildButtonBar(dialog);

        /* OKボタン：プレビュー線をガイド化してレイヤーを再ロック / OK: convert the preview to guides, then re-lock the layer */
        okButton.onClick = function() {
            convertPreviewToGuides();
            if (guideLayer) {
                guideLayer.locked = true;
            }
            activePreviewPaths = null;
            guidesCommitted = true;
            dialog.close();
        };

        /**
         * 対象の選択状態をUIに反映してプレビューを更新する
         * @returns {void}
         */
        function syncTargetState() {
            /* カンバス対象では延長が効かないので行をディム / The extension has no effect on the canvas, so dim the row */
            extensionRow.enabled = artboardRadio.value;
            drawPreview();
        }
        canvasRadio.onClick = syncTargetState;
        artboardRadio.onClick = syncTargetState;

        /* イベント：その他はプレビュー再描画のみ / Events: others just redraw the preview */
        horizontalRadio.onClick = drawPreview;
        verticalRadio.onClick = drawPreview;
        guideLayerRadio.onClick = drawPreview;
        activeLayerRadio.onClick = drawPreview;

        /* H/Vキーで方向切り替え。数値欄に文字が入らないようキャプチャフェーズで受ける / Switch direction with H/V; capture phase keeps the letter out of the numeric fields */
        dialog.addEventListener("keydown", function(event) {
            var pressedKey = (event.keyName || "").toUpperCase();
            if (pressedKey !== "H" && pressedKey !== "V") return;
            horizontalRadio.value = (pressedKey === "H");
            verticalRadio.value = !horizontalRadio.value;
            drawPreview();
            event.preventDefault();
        }, true);

        /* 初期状態を反映し、そのまま初回プレビューを描画 / Apply the initial state and draw the first preview */
        syncTargetState();

        /* ダイアログ表示時に「開始位置」入力欄へフォーカス / Focus the position input on dialog show */
        positionInput.active = true;

        /* ダイアログを閉じたら後片付け（キャンセル・ESCも含む）/ Clean up on close (Cancel and ESC included) */
        dialog.onClose = function() {
            removePreviewPaths();
            if (!guidesCommitted) {
                /* ガイド化せずに閉じたので、レイヤーを元の状態へ / Closed without committing, so put the layer back */
                restoreGuideLayer();
            }
        };
        return dialog;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * ガイド作成ダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert', 'noDocument'));
            return;
        }
        createGuideDialog().show();
    }

    main();

})();
