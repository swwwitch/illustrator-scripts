#targetengine "MyScriptEngine"
#target illustrator
#include "../stroke-table/ColorPicker.jsx"

/*

### 概要

- 円・正多角形・スター・スーパー楕円・ルーロー（定幅図形）を1つのダイアログから作成します。
- 辺の数、幅、回転、塗りと線、角丸、アンカーポイントの各オプションをプレビューしながら設定します。
- 機能の詳細と使い方はREADME（readme-ja/SmartShapeMaker.md）を参照してください。

### 仕様・注意

- 数値はIllustratorの定規単位で入力します。線幅だけは線の単位に従います。
- ［幅］は図形の外接円の直径です。正方形だけは既定の回転角（45°）で1辺の長さと一致します。
- 図形はドキュメントウィンドウの中央に作成します。
- プレビューはUndo履歴を汚さず、確定した図形は1ステップで取り消せます。
  ただし［ライブシェイプ化］と［ラフ効果で追加］はメニューコマンドを使うため、それぞれ取り消しが1ステップ増えます。
- ダイアログの値はIllustratorの起動中のみ保持し、再起動でリセットされます。［アンカーポイントで分割］だけは毎回OFFで開きます。

### キーボードショートカット

- E：円（0）／A：回転／S：スター／P：五芒星／D：アンカーポイントで分割
- L：三角形（左）／R：三角形（右）／B：三角形（下）。いずれも辺の数を3にします。

*/

/*

### Overview

- Builds circles, regular polygons, stars, superellipses and Reuleaux (constant-width) shapes from one dialog.
- Sides, width, rotation, fill and stroke, corner smoothing and anchor options are all set with a live preview.
- See the README (readme-en/SmartShapeMaker.md) for the full feature list and usage.

### Notes

- Values are entered in Illustrator's ruler unit; the stroke width follows the stroke unit.
- "Width" is the diameter of the circumscribed circle. Only a square matches its edge length, at the default 45 degree rotation.
- Shapes are created at the center of the document window.
- The preview leaves the undo history clean and the confirmed shape is undone in a single step.
  "Live Shape" and "Add Anchors (Roughen)" run menu commands, so each of them adds one more undo step.
- Dialog values persist only while Illustrator is running. "Split at Anchor Points" always opens off.

### Keyboard shortcuts

- E: circle (0) / A: rotate / S: star / P: pentagram / D: split at anchor points
- L: triangle left / R: triangle right / B: triangle down. Each one also sets the side count to 3.

*/

(function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================
    var SCRIPT_NAME     = "SmartShapeMaker";              /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v2.2.0";                       /* バージョン / version */
    var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
    var SCRIPT_RELEASED = "2025-05-02";                   /* 最初のリリース日 / first release date */
    var SCRIPT_UPDATED  = "2026-07-31";                   /* 更新日 / last updated */

    // README (Japanese)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartShapeMaker.md
    // README (English)
    // https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartShapeMaker.md
    var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n005a7087f9c3"; /* 紹介記事 / article URL */

    // Released under the MIT license
    // http://opensource.org/licenses/mit-license.php

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 外部JSX読み込み時の警告を抑止 / Suppress the warning raised when an external JSX is loaded */
    app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

    /* ダイアログとUIの外観 / Dialog and UI appearance */
    var UI_CONFIG = {
        dialogOffsetX: 300,    /* ダイアログの表示位置オフセットX / dialog offset X */
        dialogOffsetY: 0,      /* ダイアログの表示位置オフセットY / dialog offset Y */
        dialogOpacity: 0.98,   /* ダイアログの不透明度 / dialog opacity */
        swatchSize: 16,        /* カラースウォッチの一辺（px） / color swatch size in px */
        sliderWidth: 200,      /* 標準スライダー幅 / default slider width */
        shortSliderWidth: 150, /* 短いスライダー幅 / short slider width */
        sidesSliderWidth: 100, /* ［辺の数］スライダー幅（入力欄と同じ行に収める） / slider width of the side count, kept on the field's row */
        zoomSliderWidth: 300   /* ［画面ズーム］スライダー幅 / slider width of the view zoom */
    };

    /* UIレイアウト：余白と間隔 / UI layout: margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /* 各パネル内の要素間隔（パネルが多い密なダイアログなので既定より詰める）
       Panel spacing used in this dialog; tighter than the default because many panels are stacked */
    var DIALOG_PANEL_SPACING = 6;

    /* 各入力の初期値 / Default values of the inputs */
    var SHAPE_DEFAULTS = {
        size: "100",             /* 幅 / width */
        rotation: "90",          /* 回転角（度） / rotation angle in degrees */
        customSides: 12,         /* ［それ以外］の辺の数 / custom side count */
        circleAnchors: 4,        /* 円のアンカーポイント数 / anchor count of a circle */
        innerRatio: 30,          /* スターの第2半径（%） / star inner radius in percent */
        superExponent: 2.5,      /* スーパー楕円の指数 / superellipse exponent */
        strokeWidth: "1",        /* 線幅 / stroke width */
        opacity: 100,            /* 不透明度（%） / opacity in percent */
        cornerRadius: "15",      /* 角丸の半径 / corner radius */
        cornerRadiusRatio: 0.15, /* 角丸半径の既定値＝幅×この比率 / corner radius default ratio */
        smoothing: 60,           /* スムージング（%） / smoothing in percent */
        reuleauxAmount: 100,     /* ルーローの度合い（%） / Reuleaux amount in percent */
        roughenDetail: "1",      /* ラフ効果の詳細 / roughen detail */
        segmentStrokeWidth: 0.3  /* 分割時に線がないときの既定線幅（pt）
                                    fallback stroke width of split segments, in pt */
    };

    /* 入力の下限と上限 [min, max] / Lower and upper bounds of the inputs */
    var SHAPE_RANGES = {
        customSides: [3, 36],     /* 辺の数 / side count */
        innerRatio: [0, 100],     /* 第2半径（%） / inner radius */
        superExponent: [1.5, 6],  /* スーパー楕円の指数 / superellipse exponent */
        smoothing: [0, 150],      /* スムージング（%） / smoothing */
        reuleauxAmount: [0, 200], /* ルーローの度合い（%） / Reuleaux amount */
        opacity: [0, 100],        /* 不透明度（%） / opacity */
        zoom: [0.1, 16]           /* 画面ズーム倍率 / view zoom factor */
    };

    /* 図形生成の内部パラメーター / Internal parameters of the shape generation */
    var SHAPE_CONFIG = {
        superEllipsePoints: 8,    /* スーパー楕円のサンプル点数 / sample point count of a superellipse */
        superEllipseHandle: 0.35, /* スーパー楕円のハンドル長比率 / handle length ratio of a superellipse */
        smoothingArmFactor: 0.8   /* 角丸のアーム長係数 / arm length factor of the corner smoothing */
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のロケールからUI言語を判定する。
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var lang = getCurrentLang();

    /* UI文言の定義 / UI string definitions */
    var LABELS = {
        dialog: {
            title: { ja: "基本図形の作成", en: "Shape Builder" },
            colorPicker: { ja: "カラーピッカー", en: "Color Picker" }
        },
        panel: {
            sides: { ja: "辺の数", en: "Sides" },
            rotation: { ja: "回転", en: "Rotate" },
            triangle: { ja: "三角形", en: "Triangle" },
            fillAndStroke: { ja: "塗りと線", en: "Fill & Stroke" },
            width: { ja: "幅", en: "Width" },
            star: { ja: "スター", en: "Star" },
            circle: { ja: "円", en: "Circle" },
            anchor: { ja: "アンカーポイント", en: "Anchor Points" },
            anchorOps: { ja: "アンカーポイントの操作", en: "Anchor Point Operations" },
            cornerSmoothing: { ja: "角丸", en: "Corner Smoothing" },
            option: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            superEllipse: { ja: "スーパー楕円", en: "Superellipse" },
            star: { ja: "スター", en: "Star" },
            pentagram: { ja: "五芒星", en: "Pentagram" },
            fill: { ja: "塗り", en: "Fill" },
            stroke: { ja: "線", en: "Stroke" },
            cornerRadius: { ja: "半径：", en: "Radius:" },
            liveShape: { ja: "ライブシェイプ化", en: "Live Shape" },
            reuleaux: { ja: "ルーロー（定幅図形）", en: "Reuleaux (Constant-Width)" },
            splitAtAnchors: { ja: "アンカーポイントで分割", en: "Split at Anchor Points" },
            roughenAnchors: { ja: "ラフ効果で追加：", en: "Add Anchors (Roughen):" }
        },
        radio: {
            circleWithZero: { ja: "0（円）", en: "0 (Circle)" },
            triangleRight: { ja: "右", en: "Right" },
            triangleLeft: { ja: "左", en: "Left" },
            triangleDown: { ja: "下", en: "Down" },
            capButt: { ja: "なし", en: "Butt" },
            capRound: { ja: "丸型", en: "Round" },
            capProjecting: { ja: "突出", en: "Projecting" }
        },
        label: {
            innerRadius: { ja: "第2半径：", en: "Inner Radius:" },
            opacity: { ja: "不透明度：", en: "Opacity:" },
            smoothing: { ja: "スムージング：", en: "Smoothing:" },
            strokeCap: { ja: "線端：", en: "Line Cap:" },
            viewZoom: { ja: "画面ズーム", en: "View Zoom" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        alert: {
            previewError: { ja: "プレビューエラー：", en: "Preview Error: " },
            finalError: { ja: "確定エラー：", en: "Final Error: " },
            noDocument: {
                ja: "ドキュメントを開いてから実行してください。",
                en: "Open a document before running this script."
            }
        }
    };

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * 定規単位のラベルとpt換算係数を取得する。
     * @returns {{label: string, factor: number}} 単位ラベルとpt換算係数
     */
    function getRulerUnitInfo() {
        var rulerType = app.preferences.getIntegerPreference("rulerType");
        var unitInfo = { label: "pt", factor: 1.0 };
        if (rulerType === 0) unitInfo = { label: "inch", factor: 72.0 };
        else if (rulerType === 1) unitInfo = { label: "mm", factor: 72.0 / 25.4 };
        else if (rulerType === 3) unitInfo = { label: "pica", factor: 12.0 };
        else if (rulerType === 4) unitInfo = { label: "cm", factor: 72.0 / 2.54 };
        else if (rulerType === 5) unitInfo = { label: "Q", factor: 72.0 / 25.4 * 0.25 };
        else if (rulerType === 6) unitInfo = { label: "px", factor: 1.0 };
        return unitInfo;
    }

    /**
     * 線の単位（strokeUnits）のラベルとpt換算係数を取得する。
     * @returns {{label: string, factorToPt: number}} 単位ラベルとpt換算係数
     */
    function getStrokeUnitInfo() {
        var strokeUnitCode;
        try {
            strokeUnitCode = app.preferences.getIntegerPreference("strokeUnits");
        } catch (e) {
            strokeUnitCode = 2; /* ptにフォールバック / fall back to pt */
        }
        var label, factor;
        switch (strokeUnitCode) {
            case 0: label = "inch"; factor = 72; break;
            case 1: label = "mm"; factor = 72 / 25.4; break;
            case 3: label = "pica"; factor = 12; break;
            case 4: label = "cm"; factor = 72 / 2.54; break;
            case 5: label = "Q"; factor = (72 / 25.4) * 0.25; break;
            case 6: label = "px"; factor = 1; break;
            default: label = "pt"; factor = 1; break; /* case 2 */
        }
        return { label: label, factorToPt: factor };
    }

    // =========================================
    // セッション状態 / Session state
    // =========================================

    /* Illustratorの起動中だけダイアログの値を保持する（#targetengineの常駐エンジンを利用）
       Dialog values are kept only while Illustrator is running, on the engine named by #targetengine */
    var SESSION_STATE_KEY = "__SmartShapeMaker_State__";
    if (!$.global[SESSION_STATE_KEY]) {
        $.global[SESSION_STATE_KEY] = {};
    }

    /**
     * セッション状態オブジェクトを取得する。
     * @returns {object} Illustratorの起動中だけ保持される状態オブジェクト
     */
    function getSessionState() {
        return $.global[SESSION_STATE_KEY];
    }

    // =========================================
    // 確定時に参照する状態 / State referenced on finalize
    // =========================================

    var previewShape = null;                   /* プレビュー中の図形 / the shape currently previewed */
    var applyLiveShape = true;                 /* ライブシェイプ化するか / whether to convert to a live shape */
    var roughenAnchorsDetail = 0;              /* ラフ効果の詳細（0で無効） / roughen detail, 0 disables it */
    var roughenAnchorsUseMenuFallback = false; /* メニューコマンドで代替するか / whether to fall back to the menu command */

    // =========================================
    // 数学ヘルパー / Math helpers
    // =========================================

    /**
     * 角度表示を整える（小数第3位まで、整数なら整数表記）。
     * @param {number} angle - 角度（度）
     * @returns {string} 表示用の角度文字列
     */
    function formatAngle(angle) {
        var rounded = Math.round(angle * 1000) / 1000;
        return (rounded % 1 === 0) ? String(Math.round(rounded)) : String(rounded);
    }

    /**
     * 数値の符号を返す（0と-0はそのまま返す）。
     * @param {number} value - 対象の数値
     * @returns {number} 1、-1、または0
     */
    function signOf(value) {
        return ((value > 0) - (value < 0)) || +value;
    }

    /**
     * 曲率半径を保つベジェハンドルの長さを求める。
     * @param {number} arm - コーナーから直線部までのアーム長
     * @param {number} radius - 角丸の半径
     * @returns {number} ハンドルの長さ
     */
    function computeCornerHandleLength(arm, radius) {
        if (arm <= 0 || radius <= 0) return 0;
        var kappa = 16 / (3 * Math.sqrt(2));
        var discriminant = radius * (8 * kappa * arm + kappa * kappa * radius);
        return ((4 * arm + kappa * radius) - Math.sqrt(discriminant)) / 2;
    }

    // =========================================
    // プレビュー管理 / Preview manager
    // =========================================

    /**
     * プレビューの適用とUndoによる巻き戻しを管理する。
     * rollback()でプレビューをすべて取り消し、confirm()で巻き戻したうえで確定処理を1回だけ実行する。
     * @constructor
     */
    function PreviewManager() {
        this.undoDepth = 0;

        /**
         * プレビュー処理を1ステップ実行し、Undo段数を記録する。
         * @param {function} stepAction - プレビューを描画する処理
         * @returns {void}
         */
        this.addStep = function (stepAction) {
            try {
                stepAction();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert(LABELS.alert.previewError[lang] + e);
            }
        };

        /**
         * 記録済みのプレビュー処理をすべて取り消す。
         * @returns {void}
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                try { app.undo(); } catch (e) { break; }
                this.undoDepth--;
            }
            try { app.redraw(); } catch (e) { }
        };

        /**
         * プレビューを巻き戻したうえで確定処理を1回だけ実行する。
         * @param {function} finalAction - 確定時に実行する処理
         * @returns {void}
         */
        this.confirm = function (finalAction) {
            if (finalAction) {
                this.rollback();
                try { finalAction(); } catch (e) { alert(LABELS.alert.finalError[lang] + e); }
                this.undoDepth = 0;
            } else {
                this.undoDepth = 0;
            }
        };
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /**
     * ウィンドウの共通レイアウトを適用する。
     * @param {Window} win - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時はWINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通レイアウトを適用する。
     * @param {Panel} panel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時はPANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通レイアウトを適用する。
     * @param {Group} tab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は既定のまま）
     * @returns {void}
     */
    function setupTab(tab, spacing) {
        tab.orientation = "column";
        tab.alignChildren = "fill";
        tab.margins = TAB_MARGINS;
        if (typeof spacing === "number") tab.spacing = spacing;
    }

    /**
     * 行グループの共通レイアウトを適用する（ボタン列など）。
     * @param {Group} group - 対象のグループ
     * @param {string|Array} [alignment] - 整列指定（省略時は"left"）
     * @param {number} [spacing] - 要素間隔（省略時はPANEL_SPACING）
     * @returns {void}
     */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定px詰める（レイアウト確定後に呼ぶ）。
     * @param {Button} button - 対象のボタン
     * @param {number} px - 詰める高さ（px）
     * @returns {void}
     */
    function trimButtonHeight(button, px) {
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) { }
    }

    // =========================================
    // UIヘルパー / UI helpers
    // =========================================

    /**
     * 入力欄の値を上下矢印キーで増減できるようにする（Shiftで±10）。
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;
            var keyboard = ScriptUI.environment.keyboardState;
            var delta = keyboard.shiftKey ? 10 : 1;
            if (event.keyName == "Up") {
                value += delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value -= delta;
                event.preventDefault();
            }
            editText.text = value;

            /* プレビュー用のonChangingがあれば呼び出して矢印キー編集も反映する
               Invoke onChanging, when present, so arrow-key edits update the preview */
            if (typeof editText.onChanging === "function") {
                try { editText.onChanging(); } catch (e) { }
            }
        });
    }

    /**
     * ラジオボタンまたは手動入力から辺の数を取得する。
     * @param {Array} sideRadios - 辺の数のラジオボタン
     * @param {EditText} customSidesInput - ［それ以外］の入力欄
     * @returns {number} 辺の数（0は円）
     */
    function getSelectedSideValue(sideRadios, customSidesInput) {
        for (var i = 0; i < sideRadios.length; i++) {
            if (sideRadios[i].value) {
                return (i === 6) ? parseInt(customSidesInput.text, 10) : [0, 3, 4, 5, 6, 8][i];
            }
        }
        return 4;
    }

    // =========================================
    // パス生成 / Path builders
    // =========================================

    /**
     * スーパー楕円のパスを作成する（サンプル点とスムーズハンドルで構成）。
     * @param {Document} doc - 対象ドキュメント
     * @param {number} sizePt - 幅（pt）
     * @param {number} exponent - スーパー楕円の指数
     * @param {number} [pointCount] - サンプル点数
     * @returns {PathItem} 作成したパス
     */
    function createSuperellipsePath(doc, sizePt, exponent, pointCount) {
        exponent = (typeof exponent === 'number' && exponent > 0) ? exponent : SHAPE_DEFAULTS.superExponent;
        pointCount = (typeof pointCount === 'number' && pointCount >= SHAPE_CONFIG.superEllipsePoints)
            ? Math.round(pointCount)
            : SHAPE_CONFIG.superEllipsePoints;

        var layer = doc.activeLayer;
        layer.locked = false;
        layer.visible = true;

        var viewCenter = doc.activeView.centerPoint;
        var centerX = viewCenter[0];
        var centerY = viewCenter[1];

        var shapeWidth = sizePt;
        var shapeHeight = sizePt;

        var TWO_PI = Math.PI * 2;
        var anchors = [];
        for (var i = 0; i < pointCount; i++) {
            var theta = (TWO_PI * i) / pointCount;
            var cosTheta = Math.cos(theta);
            var sinTheta = Math.sin(theta);
            var x = Math.pow(Math.abs(cosTheta), 2 / exponent) * (shapeWidth / 2) * signOf(cosTheta);
            var y = Math.pow(Math.abs(sinTheta), 2 / exponent) * (shapeHeight / 2) * signOf(sinTheta);
            anchors.push([x + centerX, y + centerY]);
        }

        var pathItem = layer.pathItems.add();
        pathItem.setEntirePath(anchors);
        pathItem.closed = true;

        /* ハンドルをスムーズにする / Smooth the handles */
        try {
            var pathPoints = pathItem.pathPoints;
            var anchorCount = pathPoints.length;
            if (anchorCount >= 4) {
                for (var k = 0; k < anchorCount; k++) {
                    var prevAnchor = anchors[(k - 1 + anchorCount) % anchorCount];
                    var currentAnchor = anchors[k];
                    var nextAnchor = anchors[(k + 1) % anchorCount];

                    /* 接線ベクトル（次点−前点） / Tangent vector (next - prev) */
                    var tangentX = nextAnchor[0] - prevAnchor[0];
                    var tangentY = nextAnchor[1] - prevAnchor[1];
                    var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY);
                    if (tangentLength === 0) continue;
                    tangentX /= tangentLength;
                    tangentY /= tangentLength;

                    /* 前後のセグメント長 / Lengths of the neighbouring segments */
                    var prevDeltaX = currentAnchor[0] - prevAnchor[0];
                    var prevDeltaY = currentAnchor[1] - prevAnchor[1];
                    var nextDeltaX = nextAnchor[0] - currentAnchor[0];
                    var nextDeltaY = nextAnchor[1] - currentAnchor[1];
                    var prevLength = Math.sqrt(prevDeltaX * prevDeltaX + prevDeltaY * prevDeltaY);
                    var nextLength = Math.sqrt(nextDeltaX * nextDeltaX + nextDeltaY * nextDeltaY);

                    var handleLength = Math.min(prevLength, nextLength) * SHAPE_CONFIG.superEllipseHandle;
                    var leftHandle = [currentAnchor[0] - tangentX * handleLength, currentAnchor[1] - tangentY * handleLength];
                    var rightHandle = [currentAnchor[0] + tangentX * handleLength, currentAnchor[1] + tangentY * handleLength];

                    pathPoints[k].anchor = currentAnchor;
                    pathPoints[k].leftDirection = leftHandle;
                    pathPoints[k].rightDirection = rightHandle;
                    pathPoints[k].pointType = PointType.SMOOTH;
                }
            }
        } catch (e) { }

        /* 見た目はcreateShapeの既定に合わせる / Match the defaults used by createShape */
        pathItem.filled = true;
        pathItem.fillColor = doc.defaultFillColor;
        pathItem.stroked = false;

        return pathItem;
    }

    /**
     * 指定したアンカー数（2以上）で円形の閉じたパスを作成する。
     * ハンドル長は k = 4/3 * tan(pi/(2N)) を用いる。
     * @param {Document} doc - 対象ドキュメント
     * @param {number} sizePt - 直径（pt）
     * @param {number} anchorCount - アンカーポイント数
     * @returns {PathItem} 作成したパス
     */
    function createCirclePathWithNAnchors(doc, sizePt, anchorCount) {
        var layer = doc.activeLayer;
        layer.locked = false;
        layer.visible = true;

        var viewCenter = doc.activeView.centerPoint;
        var centerX = viewCenter[0];
        var centerY = viewCenter[1];

        var radius = sizePt / 2;
        anchorCount = (typeof anchorCount === 'number') ? Math.round(anchorCount) : SHAPE_DEFAULTS.circleAnchors;
        if (anchorCount < 2) anchorCount = 2;

        var handleRatio = (4 / 3) * Math.tan(Math.PI / (2 * anchorCount));
        var handleLength = radius * handleRatio;

        /* 頂点が真上に来るよう-90°から並べる / Start at -90 degrees so one anchor sits on top */
        var anchors = [];
        for (var i = 0; i < anchorCount; i++) {
            var angle = (-Math.PI / 2) + (2 * Math.PI * i) / anchorCount;
            anchors.push([centerX + radius * Math.cos(angle), centerY + radius * Math.sin(angle)]);
        }

        var pathItem = layer.pathItems.add();
        pathItem.setEntirePath(anchors);
        pathItem.closed = true;

        /* 接線方向にスムーズハンドルを置く / Place smooth handles along the tangents */
        try {
            var pathPoints = pathItem.pathPoints;
            for (var j = 0; j < anchorCount; j++) {
                var anchorX = anchors[j][0];
                var anchorY = anchors[j][1];
                var anchorAngle = (-Math.PI / 2) + (2 * Math.PI * j) / anchorCount;

                /* 接線は半径に直交する [-sin, cos] / The tangent is perpendicular to the radius */
                var tangentX = -Math.sin(anchorAngle);
                var tangentY = Math.cos(anchorAngle);

                pathPoints[j].anchor = [anchorX, anchorY];
                pathPoints[j].leftDirection = [anchorX - tangentX * handleLength, anchorY - tangentY * handleLength];
                pathPoints[j].rightDirection = [anchorX + tangentX * handleLength, anchorY + tangentY * handleLength];
                pathPoints[j].pointType = PointType.SMOOTH;
            }
        } catch (e) { }

        pathItem.filled = true;
        pathItem.fillColor = doc.defaultFillColor;
        pathItem.stroked = false;

        return pathItem;
    }

    /**
     * 奇数辺の正多角形の各辺を円弧に置き換えてルーロー図形にする。
     * 参考スクリプト reuleaux_polygon.jsx のロジックを移植。
     * @param {PathItem} pathItem - 対象のパス（奇数個のアンカーを持つ多角形）
     * @param {number} amount - 度合い（1.0が標準、0.0〜2.0）
     * @returns {PathItem} 変換後のパス
     */
    function applyReuleauxToPolygon(pathItem, amount) {
        try {
            if (!pathItem || pathItem.typename !== "PathItem") return pathItem;
            if (!pathItem.pathPoints || pathItem.pathPoints.length < 3) return pathItem;
            var pathPoints = pathItem.pathPoints;
            var pointCount = pathPoints.length;
            if (pointCount % 2 === 0) return pathItem; /* 奇数辺のみ / odd side counts only */

            amount = (typeof amount === "number") ? amount : 1;
            if (isNaN(amount)) amount = 1;
            if (amount < 0) amount = 0;
            if (amount > 2) amount = 2;

            /* アンカー座標をキャッシュ / Cache the anchor coordinates */
            var anchorCoords = [];
            for (var i = 0; i < pointCount; i++) {
                anchorCoords.push([pathPoints[i].anchor[0], pathPoints[i].anchor[1]]);
            }

            for (i = 0; i < pointCount; i++) {
                var startIndex = i;
                var endIndex = (i + 1) % pointCount;
                var centerIndex = (i + Math.floor((pointCount + 1) / 2)) % pointCount;

                var arcStart = anchorCoords[startIndex];
                var arcEnd = anchorCoords[endIndex];
                var arcCenter = anchorCoords[centerIndex];

                var vectorToStart = [arcStart[0] - arcCenter[0], arcStart[1] - arcCenter[1]];
                var vectorToEnd = [arcEnd[0] - arcCenter[0], arcEnd[1] - arcCenter[1]];

                var crossZ = vectorToStart[0] * vectorToEnd[1] - vectorToStart[1] * vectorToEnd[0];

                var radiusToStart = Math.sqrt(vectorToStart[0] * vectorToStart[0] + vectorToStart[1] * vectorToStart[1]);
                var radiusToEnd = Math.sqrt(vectorToEnd[0] * vectorToEnd[0] + vectorToEnd[1] * vectorToEnd[1]);
                if (radiusToStart === 0 || radiusToEnd === 0) continue;
                var arcRadius = (radiusToStart + radiusToEnd) / 2;

                var dotProduct = vectorToStart[0] * vectorToEnd[0] + vectorToStart[1] * vectorToEnd[1];
                var cosTheta = dotProduct / (radiusToStart * radiusToEnd);
                if (cosTheta < -1) cosTheta = -1;
                if (cosTheta > 1) cosTheta = 1;
                var deltaTheta = Math.acos(cosTheta);

                var handleLength = arcRadius * (4 / 3) * Math.tan(deltaTheta / 4);
                handleLength *= amount; /* 度合いを反映 / apply the amount */

                var tangentStart, tangentEnd;
                if (crossZ > 0) {
                    tangentStart = [-vectorToStart[1], vectorToStart[0]];
                    tangentEnd = [vectorToEnd[1], -vectorToEnd[0]];
                } else {
                    tangentStart = [vectorToStart[1], -vectorToStart[0]];
                    tangentEnd = [-vectorToEnd[1], vectorToEnd[0]];
                }

                var tangentStartLength = Math.sqrt(tangentStart[0] * tangentStart[0] + tangentStart[1] * tangentStart[1]);
                var tangentEndLength = Math.sqrt(tangentEnd[0] * tangentEnd[0] + tangentEnd[1] * tangentEnd[1]);
                if (tangentStartLength === 0 || tangentEndLength === 0) continue;

                var rightHandle = [
                    arcStart[0] + handleLength * tangentStart[0] / tangentStartLength,
                    arcStart[1] + handleLength * tangentStart[1] / tangentStartLength
                ];
                var leftHandle = [
                    arcEnd[0] + handleLength * tangentEnd[0] / tangentEndLength,
                    arcEnd[1] + handleLength * tangentEnd[1] / tangentEndLength
                ];

                pathPoints[startIndex].rightDirection = rightHandle;
                pathPoints[endIndex].leftDirection = leftHandle;

                pathPoints[startIndex].pointType = PointType.CORNER;
                pathPoints[endIndex].pointType = PointType.CORNER;
            }

            pathItem.closed = true;
        } catch (e) { }
        return pathItem;
    }

    /**
     * スムージングを効かせた角丸長方形のパスを作成する。
     * 黒野真吾さんの corner_smoothing.jsx をもとにしている。
     * @param {Document} doc - 対象ドキュメント
     * @param {number} left - 左端のX座標
     * @param {number} top - 上端のY座標
     * @param {number} rectWidth - 幅
     * @param {number} rectHeight - 高さ
     * @param {number} radius - 角丸の半径
     * @param {number} smoothing - スムージング量（0以上、1.0＝100%）
     * @returns {PathItem} 作成したパス
     */
    function buildSmoothedRect(doc, left, top, rectWidth, rectHeight, radius, smoothing) {
        radius = Math.min(Math.abs(radius), rectWidth / 2, rectHeight / 2);
        smoothing = Math.max(0, smoothing);

        var armLength = radius * (1 + SHAPE_CONFIG.smoothingArmFactor * smoothing);
        var armX = Math.min(armLength, rectWidth / 2);
        var armY = Math.min(armLength, rectHeight / 2);

        var handleX = computeCornerHandleLength(armX, radius);
        var handleY = computeCornerHandleLength(armY, radius);

        var edgeLeft = left;
        var edgeTop = top;
        var edgeRight = left + rectWidth;
        var edgeBottom = top - rectHeight;
        var midX = (edgeLeft + edgeRight) / 2;
        var midY = (edgeTop + edgeBottom) / 2;

        var mergeTopBottom = (armLength >= rectWidth / 2 - 0.01);
        var mergeSides = (armLength >= rectHeight / 2 - 0.01);

        var pointSpecs = [];

        /**
         * アンカーと左右ハンドルの組を記録する。
         * @param {Array} anchor - アンカー座標
         * @param {Array} leftHandle - 左方向ハンドルの座標
         * @param {Array} rightHandle - 右方向ハンドルの座標
         * @returns {void}
         */
        function addPoint(anchor, leftHandle, rightHandle) {
            pointSpecs.push({ anchor: anchor, leftHandle: leftHandle, rightHandle: rightHandle });
        }

        /* 上辺 / Top edge */
        if (mergeTopBottom) {
            addPoint([midX, edgeTop], [midX - handleX, edgeTop], [midX + handleX, edgeTop]);
        } else {
            addPoint([edgeLeft + armX, edgeTop], [edgeLeft + armX - handleX, edgeTop], [edgeLeft + armX, edgeTop]);
            addPoint([edgeRight - armX, edgeTop], [edgeRight - armX, edgeTop], [edgeRight - armX + handleX, edgeTop]);
        }
        /* 右辺 / Right edge */
        if (mergeSides) {
            addPoint([edgeRight, midY], [edgeRight, midY + handleY], [edgeRight, midY - handleY]);
        } else {
            addPoint([edgeRight, edgeTop - armY], [edgeRight, edgeTop - armY + handleY], [edgeRight, edgeTop - armY]);
            addPoint([edgeRight, edgeBottom + armY], [edgeRight, edgeBottom + armY], [edgeRight, edgeBottom + armY - handleY]);
        }
        /* 下辺 / Bottom edge */
        if (mergeTopBottom) {
            addPoint([midX, edgeBottom], [midX + handleX, edgeBottom], [midX - handleX, edgeBottom]);
        } else {
            addPoint([edgeRight - armX, edgeBottom], [edgeRight - armX + handleX, edgeBottom], [edgeRight - armX, edgeBottom]);
            addPoint([edgeLeft + armX, edgeBottom], [edgeLeft + armX, edgeBottom], [edgeLeft + armX - handleX, edgeBottom]);
        }
        /* 左辺 / Left edge */
        if (mergeSides) {
            addPoint([edgeLeft, midY], [edgeLeft, midY - handleY], [edgeLeft, midY + handleY]);
        } else {
            addPoint([edgeLeft, edgeBottom + armY], [edgeLeft, edgeBottom + armY - handleY], [edgeLeft, edgeBottom + armY]);
            addPoint([edgeLeft, edgeTop - armY], [edgeLeft, edgeTop - armY], [edgeLeft, edgeTop - armY + handleY]);
        }

        var layer = doc.activeLayer;
        var pathItem = layer.pathItems.add();
        pathItem.closed = true;

        for (var i = 0; i < pointSpecs.length; i++) {
            var pathPoint = pathItem.pathPoints.add();
            pathPoint.anchor = pointSpecs[i].anchor;
            pathPoint.leftDirection = pointSpecs[i].leftHandle;
            pathPoint.rightDirection = pointSpecs[i].rightHandle;
            pathPoint.pointType = PointType.CORNER;
        }

        return pathItem;
    }

    /**
     * 閉じたパスをアンカーポイントごとに分割し、開いたパスのグループにする。
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} pathItem - 分割元のパス
     * @param {object} strokeOpts - 線の設定 {enabled, color, widthPt}
     * @param {StrokeCap} strokeCap - 線端の種類
     * @returns {GroupItem|PathItem} 分割後のグループ（分割できない場合は元のパス）
     */
    function splitPathAtAnchors(doc, pathItem, strokeOpts, strokeCap) {
        if (!pathItem || !pathItem.pathPoints || pathItem.pathPoints.length < 2) return pathItem;

        var layer = doc.activeLayer;
        var segmentGroup = layer.groupItems.add();

        var pathPoints = pathItem.pathPoints;
        var pointCount = pathPoints.length;
        var isClosed = pathItem.closed;

        for (var i = 0; i < pointCount; i++) {
            var j = i + 1;
            if (j >= pointCount) {
                if (!isClosed) break;
                j = 0;
            }

            var startPoint = pathPoints[i];
            var endPoint = pathPoints[j];

            /* このセグメント用に開いたパスを作る / Create an open path for this segment */
            var segmentPath = segmentGroup.pathItems.add();
            segmentPath.closed = false;

            /* 先にアンカーを設定 / Set the anchors first */
            segmentPath.setEntirePath([startPoint.anchor, endPoint.anchor]);

            /* ハンドルを引き継ぐ。2点の開いたパスでは始点はright、終点はleftを使う
               Copy the handles: the start point uses rightDirection, the end point uses leftDirection */
            segmentPath.pathPoints[0].leftDirection = startPoint.anchor;
            segmentPath.pathPoints[0].rightDirection = startPoint.rightDirection;
            segmentPath.pathPoints[0].pointType = startPoint.pointType;

            segmentPath.pathPoints[1].leftDirection = endPoint.leftDirection;
            segmentPath.pathPoints[1].rightDirection = endPoint.anchor;
            segmentPath.pathPoints[1].pointType = endPoint.pointType;

            /* 開いたパスに塗りは合わないので線だけにする / Open paths get a stroke, not a fill */
            segmentPath.filled = false;
            segmentPath.stroked = true;

            try {
                if (strokeOpts && strokeOpts.enabled) {
                    segmentPath.strokeColor = strokeOpts.color;
                } else {
                    var fallbackColor;
                    if (doc && doc.documentColorSpace === DocumentColorSpace.CMYK) {
                        fallbackColor = new CMYKColor();
                        fallbackColor.cyan = 0; fallbackColor.magenta = 0; fallbackColor.yellow = 0; fallbackColor.black = 100;
                    } else {
                        fallbackColor = new RGBColor();
                        fallbackColor.red = 0; fallbackColor.green = 0; fallbackColor.blue = 0;
                    }
                    segmentPath.strokeColor = fallbackColor;
                }
            } catch (e) { }
            try {
                segmentPath.strokeWidth = (strokeOpts && strokeOpts.enabled) ? strokeOpts.widthPt : SHAPE_DEFAULTS.segmentStrokeWidth;
            } catch (e) { }
            try {
                if (strokeCap) segmentPath.strokeCap = strokeCap;
            } catch (e) { }
        }

        /* 元のパスを削除 / Remove the original path */
        try { pathItem.remove(); } catch (e) { }

        return segmentGroup;
    }

    /**
     * 指定したパラメーターから図形を作成し、選択状態にする。
     * @param {Document} doc - 対象ドキュメント
     * @param {number} sizePt - 幅（pt）
     * @param {number} sides - 辺の数（0は円）
     * @param {boolean} isStar - スターにするか
     * @param {number} innerRatio - スターの第2半径（%）
     * @param {boolean} rotateEnabled - 回転を適用するか
     * @param {number} rotateAngle - 回転角（度）
     * @param {boolean} splitAtAnchors - アンカーポイントで分割するか
     * @param {boolean} useSuperEllipse - スーパー楕円にするか
     * @param {number} superExponent - スーパー楕円の指数
     * @param {number} circleAnchorCount - 円のアンカーポイント数
     * @param {boolean} useReuleaux - ルーロー図形にするか
     * @param {number} reuleauxAmount - ルーローの度合い（1.0が標準）
     * @param {object} fillOpts - 塗りの設定 {enabled, color}
     * @param {object} strokeOpts - 線の設定 {enabled, color, widthPt}
     * @param {object} cornerSmoothing - 角丸の設定 {radius, smoothing}（不要ならnull）
     * @param {StrokeCap} strokeCap - 線端の種類
     * @param {number} opacity - 不透明度（%）
     * @returns {PathItem|GroupItem} 作成した図形
     */
    function createShape(doc, sizePt, sides, isStar, innerRatio, rotateEnabled, rotateAngle, splitAtAnchors, useSuperEllipse, superExponent, circleAnchorCount, useReuleaux, reuleauxAmount, fillOpts, strokeOpts, cornerSmoothing, strokeCap, opacity) {
        var layer = doc.activeLayer;
        layer.locked = false;
        layer.visible = true;
        var viewCenter = doc.activeView.centerPoint;
        var radius = sizePt / 2;
        var innerRadius = radius * (innerRatio / 100);
        var shape;

        if (sides === 0) {
            if (useSuperEllipse) {
                shape = createSuperellipsePath(doc, sizePt, superExponent);
            } else {
                /* 既定の4アンカーはIllustratorの楕円、それ以外は独自のスムーズパス
                   Four anchors use Illustrator's ellipse; other counts build a custom smooth path */
                var anchorCount = (typeof circleAnchorCount === 'number') ? Math.round(circleAnchorCount) : SHAPE_DEFAULTS.circleAnchors;
                if (anchorCount < 2) anchorCount = 2;
                if (anchorCount === 4) {
                    shape = layer.pathItems.ellipse(viewCenter[1] + radius, viewCenter[0] - radius, sizePt, sizePt);
                } else {
                    shape = createCirclePathWithNAnchors(doc, sizePt, anchorCount);
                }
            }
        } else if (isStar) {
            shape = doc.pathItems.star(viewCenter[0], viewCenter[1], radius, innerRadius, sides);
        } else if (sides === 4 && cornerSmoothing && cornerSmoothing.radius > 0 && cornerSmoothing.smoothing > 0) {
            /* スムージングありの角丸は独自のベジェパス / Corner smoothing above zero builds a custom bezier path */
            var smoothRadius = cornerSmoothing.radius;
            var smoothAmount = cornerSmoothing.smoothing / 100;
            var smoothLeft = viewCenter[0] - sizePt / 2;
            var smoothTop = viewCenter[1] + sizePt / 2;
            shape = buildSmoothedRect(doc, smoothLeft, smoothTop, sizePt, sizePt, smoothRadius, smoothAmount);
        } else if (sides === 4 && cornerSmoothing && cornerSmoothing.radius > 0 && cornerSmoothing.smoothing === 0) {
            /* スムージング0の角丸は通常の正方形＋［角を丸くする］効果
               A zero smoothing value uses a plain square plus the Round Corners live effect */
            var roundedSquareRadius = sizePt / Math.sqrt(2);
            shape = doc.pathItems.polygon(viewCenter[0], viewCenter[1], roundedSquareRadius, sides);
            try {
                var roundCornersXml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + cornerSmoothing.radius + ' "/></LiveEffect>';
                shape.applyEffect(roundCornersXml);
            } catch (e) { }
        } else if (sides === 4) {
            /* 正方形は1辺の長さを幅として扱うので外接円の半径に換算する（既定の45°回転が前提）
               A square is sized by its edge, so convert to the circumscribed radius; assumes the default 45 degree rotation */
            var squareRadius = sizePt / Math.sqrt(2);
            shape = doc.pathItems.polygon(viewCenter[0], viewCenter[1], squareRadius, sides);
        } else {
            /* 正方形以外はsizePtを外接円の直径として扱う（バウンディングボックスの幅とは一致しない）
               Other polygons treat sizePt as the circumscribed diameter, which is not the bounding box width */
            shape = doc.pathItems.polygon(viewCenter[0], viewCenter[1], radius, sides);
        }

        /* 塗りと線を適用 / Apply the fill and stroke options */
        if (fillOpts && fillOpts.enabled) {
            shape.filled = true;
            shape.fillColor = fillOpts.color;
        } else {
            shape.filled = false;
        }
        if (strokeOpts && strokeOpts.enabled) {
            shape.stroked = true;
            shape.strokeColor = strokeOpts.color;
            shape.strokeWidth = strokeOpts.widthPt;
        } else {
            shape.stroked = false;
        }

        var bounds = shape.geometricBounds;
        var shapeCenterX = (bounds[0] + bounds[2]) / 2;
        var shapeCenterY = (bounds[1] + bounds[3]) / 2;
        shape.translate(viewCenter[0] - shapeCenterX, viewCenter[1] - shapeCenterY);

        /* 奇数辺の多角形をルーロー（定幅図形）に変換 / Convert odd-sided polygons into constant-width shapes */
        if (useReuleaux && !isStar && sides > 0 && (sides % 2 === 1)) {
            try {
                shape = applyReuleauxToPolygon(shape, reuleauxAmount);
            } catch (e) { }
        }

        if (rotateEnabled && !isNaN(rotateAngle)) {
            shape.rotate(rotateAngle, true, true, true, true, Transformation.CENTER);
        }
        if (splitAtAnchors) {
            shape = splitPathAtAnchors(doc, shape, strokeOpts, strokeCap);
        }
        if (typeof opacity === "number" && opacity < 100) {
            try { shape.opacity = opacity; } catch (e) { }
        }
        doc.selection = [shape];
        return shape;
    }

    /**
     * ラフ効果でアンカーポイントを追加する（変形量0なので位置は動かない）。
     * @param {PathItem|GroupItem} target - 対象のオブジェクト
     * @param {number} detail - ラフ効果の詳細（0以下なら何もしない）
     * @returns {void}
     */
    function applyRoughenEffect(target, detail) {
        if (!target || !(detail > 0)) return;
        try {
            var roughenXml = '<LiveEffect name="Adobe Roughen"><Dict data="R asiz 0 R size 0 R absoluteness 0 R dtal ' + detail + ' R roundness 0 "/></LiveEffect>';
            target.applyEffect(roughenXml);
        } catch (e) { }
    }

    /**
     * プレビュー中の図形を選択状態にし、プレビュー参照をクリアする。
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function finalizeShape(doc) {
        if (!previewShape) return;
        doc.selection = [previewShape];
        previewShape = null;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを表示し、OKが押されたら確定用の状態を整える。
     * @param {string} unitLabel - 定規単位のラベル
     * @param {number} unitFactor - 定規単位のpt換算係数
     * @param {object} strokeUnitInfo - 線の単位情報 {label, factorToPt}
     * @returns {boolean} OKで確定できたときtrue、キャンセルや失敗時はnull
     */
    function showInputDialog(unitLabel, unitFactor, strokeUnitInfo) {
        var dialog = new Window("dialog", LABELS.dialog.title[lang] + " " + SCRIPT_VERSION);
        var previewManager = new PreviewManager();
        var doc = app.activeDocument;

        /**
         * ダイアログの表示位置を相対的にずらす。
         * @param {Window} targetDialog - 対象のダイアログ
         * @param {number} dx - X方向の移動量
         * @param {number} dy - Y方向の移動量
         * @returns {void}
         */
        function shiftDialogPosition(targetDialog, dx, dy) {
            try {
                var currentX = targetDialog.location[0];
                var currentY = targetDialog.location[1];
                targetDialog.location = [currentX + dx, currentY + dy];
            } catch (e) { }
        }

        /**
         * セッション状態に保存されたダイアログ位置を取得する。
         * @returns {Array} [x, y] の配列、保存がなければnull
         */
        function getSavedDialogLocation() {
            try {
                var state = getSessionState();
                if (state && state.dialogLocation && state.dialogLocation.length === 2) {
                    var x = Number(state.dialogLocation[0]);
                    var y = Number(state.dialogLocation[1]);
                    if (!isNaN(x) && !isNaN(y)) return [x, y];
                }
            } catch (e) { }
            return null;
        }

        /**
         * ダイアログの表示位置をセッション状態に保存する。
         * @param {Window} targetDialog - 対象のダイアログ
         * @returns {void}
         */
        function saveDialogLocation(targetDialog) {
            try {
                if (!targetDialog || !targetDialog.location) return;
                var state = getSessionState();
                state.dialogLocation = [Number(targetDialog.location[0]), Number(targetDialog.location[1])];
            } catch (e) { }
        }

        /**
         * 円のアンカーポイント数をラジオボタンから取得する。
         * @returns {number} アンカーポイント数
         */
        function getCircleAnchorCountFromRadios() {
            try {
                return circleAnchorRadios.anchors2.value ? 2 :
                    (circleAnchorRadios.anchors3.value ? 3 :
                        (circleAnchorRadios.anchors5.value ? 5 :
                            (circleAnchorRadios.anchors6.value ? 6 : SHAPE_DEFAULTS.circleAnchors)));
            } catch (e) {
                return SHAPE_DEFAULTS.circleAnchors;
            }
        }

        /**
         * ダイアログの不透明度を設定する。
         * @param {Window} targetDialog - 対象のダイアログ
         * @param {number} opacityValue - 不透明度（0.0〜1.0）
         * @returns {void}
         */
        function setDialogOpacity(targetDialog, opacityValue) {
            try {
                targetDialog.opacity = opacityValue;
            } catch (e) { }
        }

        /**
         * 入力欄にフォーカスがあるかどうかを判定する。
         * @param {object} target - イベントの発生元
         * @returns {boolean} 入力欄ならtrue
         */
        function isTextInputTarget(target) {
            try {
                return !!(target && target.type === "edittext");
            } catch (e) {
                return false;
            }
        }

        /* キーボードショートカット / Keyboard shortcuts */
        dialog.addEventListener("keydown", function (event) {
            if (!event || !event.keyName) return;

            /* 入力欄の編集中はショートカットを発火させない
               Shortcuts must not fire while a text field is being edited */
            if (isTextInputTarget(event.target)) return;

            switch (event.keyName.toUpperCase()) {

                case "E":
                    /* 辺の数を0（円）にする / Set the side count to 0 (circle) */
                    for (var i = 0; i < sideRadios.length; i++) sideRadios[i].value = false;
                    sideRadios[0].value = true;
                    setCustomSidesEnabled(false);
                    applyAutoRotationForSides(0);
                    updatePreview();
                    event.preventDefault();
                    break;

                case "L":
                    /* 三角形（左） / Triangle pointing left */
                    for (var i = 0; i < sideRadios.length; i++) sideRadios[i].value = false;
                    sideRadios[1].value = true;
                    setCustomSidesEnabled(false);
                    applyAutoRotationForSides(3);
                    triangleLeftRadio.value = true;
                    onTriangleDirectionChange();
                    event.preventDefault();
                    break;

                case "R":
                    /* 三角形（右） / Triangle pointing right */
                    for (var i = 0; i < sideRadios.length; i++) sideRadios[i].value = false;
                    sideRadios[1].value = true;
                    setCustomSidesEnabled(false);
                    applyAutoRotationForSides(3);
                    triangleRightRadio.value = true;
                    onTriangleDirectionChange();
                    event.preventDefault();
                    break;

                case "B":
                    /* 三角形（下） / Triangle pointing down */
                    for (var i = 0; i < sideRadios.length; i++) sideRadios[i].value = false;
                    sideRadios[1].value = true;
                    setCustomSidesEnabled(false);
                    applyAutoRotationForSides(3);
                    triangleDownRadio.value = true;
                    onTriangleDirectionChange();
                    event.preventDefault();
                    break;

                case "D":
                    /* ［アンカーポイントで分割］の切り替え / Toggle "split at anchor points" */
                    splitAtAnchorsCheck.value = !splitAtAnchorsCheck.value;
                    if (typeof splitAtAnchorsCheck.onClick === "function") {
                        splitAtAnchorsCheck.onClick();
                    } else {
                        updatePreview();
                    }
                    event.preventDefault();
                    break;

                case "A":
                    /* ［回転］の切り替え / Toggle the rotation */
                    rotateCheck.value = !rotateCheck.value;
                    rotateInput.enabled = rotateCheck.value;
                    rotateUnitLabel.enabled = rotateCheck.value;
                    if (rotateCheck.value) {
                        applyDefaultRotationWhenEnablingRotate();
                    }
                    updatePreview();
                    event.preventDefault();
                    break;

                case "S":
                    /* ［スター］の切り替え / Toggle the star */
                    starCheck.value = !starCheck.value;
                    if (!starCheck.value) {
                        pentagramCheck.value = false;
                    }
                    updatePreview();
                    event.preventDefault();
                    break;

                case "P":
                    /* ［五芒星］の切り替え（スターがONのときのみ） / Toggle the pentagram, the star must be on */
                    if (!starCheck.value) {
                        starCheck.value = true;
                    }
                    pentagramCheck.value = !pentagramCheck.value;
                    updatePreview();
                    event.preventDefault();
                    break;
            }
        });
        setupWindow(dialog);

        var sideRadios = [], customSidesInput;
        var contentGroup = dialog.add("group");
        contentGroup.orientation = "row";
        contentGroup.alignChildren = ["fill", "top"];
        contentGroup.spacing = COLUMN_SPACING;

        /* 左カラム（辺の数・回転・塗りと線・幅） / Left column: sides, rotation, fill and stroke, width */
        var leftColumn = contentGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";
        leftColumn.alignment = "top";
        leftColumn.spacing = DIALOG_PANEL_SPACING;

        var sidesPanel = leftColumn.add("panel", undefined, LABELS.panel.sides[lang]);
        setupPanel(sidesPanel, DIALOG_PANEL_SPACING);

        sideRadios[0] = sidesPanel.add("radiobutton", undefined, LABELS.radio.circleWithZero[lang]);
        sideRadios[1] = sidesPanel.add("radiobutton", undefined, "3");
        sideRadios[2] = sidesPanel.add("radiobutton", undefined, "4");
        sideRadios[3] = sidesPanel.add("radiobutton", undefined, "5");
        sideRadios[4] = sidesPanel.add("radiobutton", undefined, "6");
        sideRadios[5] = sidesPanel.add("radiobutton", undefined, "8");

        /* ［それ以外］はラベルを持たず、ラジオ・入力欄・スライダーを1行に並べる
           The custom side count has no label; its radio, field and slider share one row */
        var customSidesRow = sidesPanel.add("group");
        customSidesRow.orientation = "row";
        customSidesRow.alignChildren = ["left", "center"];

        sideRadios[6] = customSidesRow.add("radiobutton", undefined, "");
        customSidesInput = customSidesRow.add("edittext", undefined, String(SHAPE_DEFAULTS.customSides));
        customSidesInput.characters = 3;
        customSidesInput.enabled = false;
        changeValueByArrowKey(customSidesInput);

        var customSidesSlider = customSidesRow.add("slider", undefined, SHAPE_DEFAULTS.customSides, SHAPE_RANGES.customSides[0], SHAPE_RANGES.customSides[1]);
        customSidesSlider.preferredSize.width = UI_CONFIG.sidesSliderWidth;
        customSidesSlider.enabled = false;
        sideRadios[2].value = true;

        /**
         * ［それ以外］の入力欄とスライダーの有効・無効をまとめて切り替える。
         * @param {boolean} isEnabled - 有効にするかどうか
         * @returns {void}
         */
        function setCustomSidesEnabled(isEnabled) {
            customSidesInput.enabled = isEnabled;
            customSidesSlider.enabled = isEnabled;
        }

        /* 回転パネル / Rotation panel */
        var rotatePanel = leftColumn.add("panel", undefined, LABELS.panel.rotation[lang]);
        setupPanel(rotatePanel, DIALOG_PANEL_SPACING);

        var rotateRow = rotatePanel.add("group");
        rotateRow.orientation = "row";
        rotateRow.alignChildren = ["left", "center"];
        rotateRow.spacing = 6;

        var rotateCheck = rotateRow.add("checkbox", undefined, "");

        var rotateInput = rotateRow.add("edittext", undefined, SHAPE_DEFAULTS.rotation);
        rotateInput.characters = 4;
        changeValueByArrowKey(rotateInput);

        var rotateUnitLabel = rotateRow.add("statictext", undefined, "°");
        /* 初期状態ではチェック時だけ手動入力を有効にする / Manual entry starts enabled only when checked */
        rotateInput.enabled = rotateCheck.value;
        rotateUnitLabel.enabled = rotateCheck.value;

        /* 塗りと線パネル / Fill and stroke panel */
        var fillStrokePanel = leftColumn.add("panel", undefined, LABELS.panel.fillAndStroke[lang]);
        setupPanel(fillStrokePanel, DIALOG_PANEL_SPACING);

        var fillStrokeLabelWidth = (lang === 'ja') ? 30 : 50;

        /**
         * Illustratorのカラーオブジェクトを ColorPicker 用の文字列に変換する。
         * @param {object} aiColor - Illustratorのカラーオブジェクト
         * @returns {string} ColorPickerが受け取る色文字列
         */
        function aiColorToPickerString(aiColor) {
            try {
                if (aiColor.typename === "RGBColor") {
                    return ColorPicker.rgbToHex(aiColor.red, aiColor.green, aiColor.blue);
                } else if (aiColor.typename === "CMYKColor") {
                    return "cmyk:" + Math.round(aiColor.cyan) + "," + Math.round(aiColor.magenta) + "," + Math.round(aiColor.yellow) + "," + Math.round(aiColor.black);
                } else if (aiColor.typename === "GrayColor") {
                    return "cmyk:0,0,0," + Math.round(aiColor.gray);
                }
            } catch (e) { }
            return "000000";
        }

        /**
         * ColorPickerの戻り値をIllustratorのカラーオブジェクトに変換する。
         * @param {string} pickerString - ColorPickerが返した色文字列
         * @returns {object} Illustratorのカラーオブジェクト
         */
        function pickerStringToAiColor(pickerString) {
            if (ColorPicker.isCmykString(pickerString)) {
                var cmykValues = ColorPicker.parseCmykString(pickerString);
                var cmykColor = new CMYKColor();
                cmykColor.cyan = cmykValues.c;
                cmykColor.magenta = cmykValues.m;
                cmykColor.yellow = cmykValues.y;
                cmykColor.black = cmykValues.k;
                return cmykColor;
            } else {
                var rgbValues = ColorPicker.hexToRGB(pickerString);
                var rgbColor = new RGBColor();
                rgbColor.red = rgbValues.r;
                rgbColor.green = rgbValues.g;
                rgbColor.blue = rgbValues.b;
                return rgbColor;
            }
        }

        /**
         * Illustratorのカラーから、スウォッチ描画用のブラシを作る。
         * @param {object} graphics - ScriptUIのgraphicsオブジェクト
         * @param {object} aiColor - Illustratorのカラーオブジェクト
         * @returns {object} ブラシ（NoColorのときnull）
         */
        function aiColorToScriptUIBrush(graphics, aiColor) {
            try {
                if (aiColor.typename === "RGBColor") {
                    return graphics.newBrush(graphics.BrushType.SOLID_COLOR, [aiColor.red / 255, aiColor.green / 255, aiColor.blue / 255, 1]);
                } else if (aiColor.typename === "CMYKColor") {
                    /* 表示用にCMYKをRGBへ近似 / Approximate CMYK as RGB for display */
                    var redValue = 1 - Math.min(1, aiColor.cyan / 100 + aiColor.black / 100);
                    var greenValue = 1 - Math.min(1, aiColor.magenta / 100 + aiColor.black / 100);
                    var blueValue = 1 - Math.min(1, aiColor.yellow / 100 + aiColor.black / 100);
                    return graphics.newBrush(graphics.BrushType.SOLID_COLOR, [redValue, greenValue, blueValue, 1]);
                } else if (aiColor.typename === "GrayColor") {
                    var grayValue = 1 - (aiColor.gray / 100);
                    return graphics.newBrush(graphics.BrushType.SOLID_COLOR, [grayValue, grayValue, grayValue, 1]);
                } else if (aiColor.typename === "NoColor") {
                    return null;
                }
            } catch (e) { }
            return graphics.newBrush(graphics.BrushType.SOLID_COLOR, [1, 1, 1, 1]);
        }

        /**
         * クリックでカラーピッカーを開けるカラースウォッチを作る。
         * @param {Group} parent - 追加先のコンテナ
         * @param {object} aiColor - 初期表示に使うIllustratorのカラー
         * @returns {Group} スウォッチのグループ
         */
        function createSwatchPanel(parent, aiColor) {
            var swatchGroup = parent.add("group");
            swatchGroup.preferredSize = [UI_CONFIG.swatchSize, UI_CONFIG.swatchSize];
            swatchGroup.minimumSize = [UI_CONFIG.swatchSize, UI_CONFIG.swatchSize];
            swatchGroup._aiColor = aiColor;
            swatchGroup.onDraw = function () {
                var graphics = this.graphics;
                var brush = aiColorToScriptUIBrush(graphics, this._aiColor);
                if (brush) {
                    graphics.rectPath(0, 0, this.size[0], this.size[1]);
                    graphics.fillPath(brush);
                }
                var borderPen = graphics.newPen(graphics.PenType.SOLID_COLOR, [0.5, 0.5, 0.5, 1], 1);
                graphics.rectPath(0, 0, this.size[0], this.size[1]);
                graphics.strokePath(borderPen);
            };
            return swatchGroup;
        }

        /* 塗りの行 / Fill row */
        var fillRow = fillStrokePanel.add("group");
        fillRow.orientation = "row";
        fillRow.alignChildren = ["left", "center"];
        fillRow.spacing = 6;

        var fillCheck = fillRow.add("checkbox", undefined, LABELS.checkbox.fill[lang]);
        fillCheck.preferredSize.width = fillStrokeLabelWidth + 16;
        fillCheck.value = true;
        var fillSwatch = createSwatchPanel(fillRow, doc.defaultFillColor);
        fillSwatch.addEventListener("click", function () {
            var pickedColor = ColorPicker.show({
                value: aiColorToPickerString(fillSwatch._aiColor),
                title: LABELS.dialog.colorPicker[lang],
                lang: lang
            });
            if (pickedColor !== null) {
                fillSwatch._aiColor = pickerStringToAiColor(pickedColor);
                try { fillSwatch.hide(); fillSwatch.show(); } catch (e) { }
                updatePreview();
            }
        });

        /* 線の行 / Stroke row */
        var strokeRow = fillStrokePanel.add("group");
        strokeRow.orientation = "row";
        strokeRow.alignChildren = ["left", "center"];
        strokeRow.spacing = 6;

        var strokeCheck = strokeRow.add("checkbox", undefined, LABELS.checkbox.stroke[lang]);
        strokeCheck.preferredSize.width = fillStrokeLabelWidth + 16;
        strokeCheck.value = false;
        var strokeSwatch = createSwatchPanel(strokeRow, doc.defaultStrokeColor);
        strokeSwatch.addEventListener("click", function () {
            var pickedColor = ColorPicker.show({
                value: aiColorToPickerString(strokeSwatch._aiColor),
                title: LABELS.dialog.colorPicker[lang],
                lang: lang
            });
            if (pickedColor !== null) {
                strokeSwatch._aiColor = pickerStringToAiColor(pickedColor);
                try { strokeSwatch.hide(); strokeSwatch.show(); } catch (e) { }
                updatePreview();
            }
        });

        /* 線幅は［線］と同じ行に置く（ラベルなし） / The stroke width sits on the stroke row, without a label */
        var strokeWidthInput = strokeRow.add("edittext", undefined, SHAPE_DEFAULTS.strokeWidth);
        strokeWidthInput.characters = 4;
        changeValueByArrowKey(strokeWidthInput);
        var strokeWidthUnitLabel = strokeRow.add("statictext", undefined, strokeUnitInfo.label);

        /**
         * 線幅の入力欄を［線］チェックの状態に合わせて有効・無効にする。
         * @returns {void}
         */
        function updateStrokeWidthEnabled() {
            var isEnabled = strokeCheck.value;
            strokeWidthInput.enabled = isEnabled;
            strokeWidthUnitLabel.enabled = isEnabled;
        }
        fillCheck.onClick = function () {
            updatePreview();
        };
        strokeCheck.onClick = function () {
            updateStrokeWidthEnabled();
            updatePreview();
        };
        strokeWidthInput.onChanging = function () {
            updatePreview();
        };
        updateStrokeWidthEnabled();

        /* 不透明度の行 / Opacity row */
        var opacityRow = fillStrokePanel.add("group");
        opacityRow.orientation = "row";
        opacityRow.alignChildren = ["left", "center"];
        opacityRow.spacing = 6;

        opacityRow.add("statictext", undefined, LABELS.label.opacity[lang]);
        var opacityInput = opacityRow.add("edittext", undefined, String(SHAPE_DEFAULTS.opacity));
        opacityInput.characters = 4;
        changeValueByArrowKey(opacityInput);
        opacityRow.add("statictext", undefined, "%");
        var opacitySlider = fillStrokePanel.add("slider", undefined, SHAPE_DEFAULTS.opacity, SHAPE_RANGES.opacity[0], SHAPE_RANGES.opacity[1]);
        opacitySlider.preferredSize.width = UI_CONFIG.sliderWidth;

        /**
         * 不透明度を有効範囲に収める（0は有効な値なので既定値に丸めない）。
         * @param {string|number} value - 入力値
         * @returns {number} 整数に丸めた不透明度（%）
         */
        function clampOpacity(value) {
            /* 入力途中の空欄はNumber()が0になってしまうので既定値に戻す
               An empty field would become 0 through Number(), so fall back to the default */
            if (typeof value === "string" && !/\S/.test(value)) return SHAPE_DEFAULTS.opacity;
            value = Math.round(Number(value));
            if (isNaN(value)) value = SHAPE_DEFAULTS.opacity;
            if (value < SHAPE_RANGES.opacity[0]) value = SHAPE_RANGES.opacity[0];
            if (value > SHAPE_RANGES.opacity[1]) value = SHAPE_RANGES.opacity[1];
            return value;
        }

        opacityInput.onChanging = function () {
            opacitySlider.value = clampOpacity(opacityInput.text);
            updatePreview();
        };

        opacitySlider.onChanging = function () {
            var value = Math.round(opacitySlider.value);
            var keyboard = ScriptUI.environment.keyboardState;
            if (keyboard.shiftKey) {
                value = Math.round(value / 10) * 10;
                opacitySlider.value = value;
            }
            opacityInput.text = String(value);
            updatePreview();
        };

        /* 幅パネル / Width panel */
        var widthPanel = leftColumn.add("panel", undefined, LABELS.panel.width[lang]);
        setupPanel(widthPanel, DIALOG_PANEL_SPACING);

        var widthRow = widthPanel.add("group");
        widthRow.orientation = "row";
        widthRow.alignChildren = ["left", "center"];

        var sizeInput = widthRow.add("edittext", undefined, SHAPE_DEFAULTS.size);
        sizeInput.characters = 5;
        changeValueByArrowKey(sizeInput);
        widthRow.add("statictext", undefined, unitLabel);

        /* 右カラム（スター・円・角丸・アンカー・オプション） / Right column */
        var rightColumn = contentGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";
        rightColumn.alignment = "top";
        rightColumn.spacing = DIALOG_PANEL_SPACING;

        /* スターパネル / Star panel */
        var starPanel = rightColumn.add("panel", undefined, LABELS.panel.star[lang]);
        setupPanel(starPanel, DIALOG_PANEL_SPACING);

        var starRow = starPanel.add("group");
        starRow.orientation = "row";
        starRow.alignChildren = ["left", "center"];
        starRow.spacing = 10;

        var starCheck = starRow.add("checkbox", undefined, LABELS.checkbox.star[lang]);
        var pentagramCheck = starRow.add("checkbox", undefined, LABELS.checkbox.pentagram[lang]);
        pentagramCheck.value = false;

        var innerRadiusRow = starPanel.add("group");
        var innerRadiusLabel = innerRadiusRow.add("statictext", undefined, LABELS.label.innerRadius[lang]);
        var innerRatioInput = innerRadiusRow.add("edittext", undefined, String(SHAPE_DEFAULTS.innerRatio));
        innerRatioInput.characters = 4;
        changeValueByArrowKey(innerRatioInput);
        var innerPercentLabel = innerRadiusRow.add("statictext", undefined, "%");

        var innerRatioSliderRow = starPanel.add("group");
        innerRatioSliderRow.orientation = "row";
        innerRatioSliderRow.alignChildren = ["left", "center"];

        var innerRatioSlider = innerRatioSliderRow.add("slider", undefined, SHAPE_DEFAULTS.innerRatio, SHAPE_RANGES.innerRatio[0], SHAPE_RANGES.innerRatio[1]);
        innerRatioSlider.preferredSize.width = UI_CONFIG.sliderWidth;

        /* 円パネル / Circle panel */
        var circlePanel = rightColumn.add("panel", undefined, LABELS.panel.circle[lang]);
        setupPanel(circlePanel, DIALOG_PANEL_SPACING);

        var superEllipseCheck = circlePanel.add("checkbox", undefined, LABELS.checkbox.superEllipse[lang]);
        superEllipseCheck.value = false;

        var superExponentRow = circlePanel.add("group");
        superExponentRow.orientation = "row";
        superExponentRow.alignChildren = ["left", "center"];
        superExponentRow.spacing = 8;

        var superExponentInput = superExponentRow.add("edittext", undefined, String(SHAPE_DEFAULTS.superExponent));
        superExponentInput.characters = 4;
        changeValueByArrowKey(superExponentInput);

        var superExponentSlider = superExponentRow.add("slider", undefined, SHAPE_DEFAULTS.superExponent, SHAPE_RANGES.superExponent[0], SHAPE_RANGES.superExponent[1]);
        superExponentSlider.preferredSize.width = UI_CONFIG.shortSliderWidth;

        /* 円のアンカーポイント数パネル / Anchor count panel of the circle */
        var circleAnchorPanel = circlePanel.add("panel", undefined, LABELS.panel.anchor[lang]);
        setupPanel(circleAnchorPanel, DIALOG_PANEL_SPACING);

        var circleAnchorColumn = circleAnchorPanel.add("group");
        circleAnchorColumn.orientation = "column";
        circleAnchorColumn.alignChildren = "left";
        circleAnchorColumn.spacing = 10;

        var circleAnchorRow = circleAnchorColumn.add("group");
        circleAnchorRow.orientation = "row";
        circleAnchorRow.alignChildren = ["left", "center"];
        circleAnchorRow.spacing = 10;

        var circleAnchorRadios = {};
        circleAnchorRadios.anchors2 = circleAnchorRow.add("radiobutton", undefined, "2");
        circleAnchorRadios.anchors3 = circleAnchorRow.add("radiobutton", undefined, "3");
        circleAnchorRadios.anchors4 = circleAnchorRow.add("radiobutton", undefined, "4");
        circleAnchorRadios.anchors5 = circleAnchorRow.add("radiobutton", undefined, "5");
        circleAnchorRadios.anchors6 = circleAnchorRow.add("radiobutton", undefined, "6");
        circleAnchorRadios.anchors4.value = true;

        /**
         * 円パネルの有効・無効を辺の数に応じて切り替える。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateCirclePanelEnabled(sidesValue) {
            var isEnabled = (sidesValue === 0);
            circlePanel.enabled = isEnabled;
            try { circleAnchorPanel.enabled = isEnabled; } catch (e) { }
            if (!isEnabled) {
                try { superEllipseCheck.value = false; } catch (e) { }
                try {
                    circleAnchorRadios.anchors2.value = false;
                    circleAnchorRadios.anchors3.value = false;
                    circleAnchorRadios.anchors4.value = true;
                    circleAnchorRadios.anchors5.value = false;
                    circleAnchorRadios.anchors6.value = false;
                    circleAnchorRadios.anchors2.enabled = true;
                    circleAnchorRadios.anchors3.enabled = true;
                    circleAnchorRadios.anchors4.enabled = true;
                    circleAnchorRadios.anchors5.enabled = true;
                    circleAnchorRadios.anchors6.enabled = true;
                } catch (e) { }
            }
        }

        /**
         * スターパネルの有効・無効を辺の数に応じて切り替える。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateStarPanelEnabled(sidesValue) {
            /* 円（0）にはスターの設定を適用しない / Star options do not apply to a circle */
            var isEnabled = (sidesValue !== 0);
            starPanel.enabled = isEnabled;
            if (!isEnabled) {
                starCheck.value = false;
                pentagramCheck.value = false;
                pentagramCheck.enabled = false;
            }
        }

        /**
         * 第2半径の入力群を［スター］チェックの状態に合わせて有効・無効にする。
         * @returns {void}
         */
        function updateInnerRadiusEnabled() {
            var isEnabled = !!starCheck.value;
            try {
                innerRadiusLabel.enabled = isEnabled;
                innerRatioInput.enabled = isEnabled;
                innerPercentLabel.enabled = isEnabled;
                innerRatioSlider.enabled = isEnabled;
            } catch (e) { }
        }

        /**
         * ルーローのチェックボックスを、辺の数とスターの状態に応じて有効・無効にする。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateReuleauxAvailability(sidesValue) {
            try {
                /* スターがONのときは奇数判定を行わない（スター側の制御を優先）
                   While the star is on, the odd-side rule is skipped and the star logic wins */
                if (starCheck && starCheck.value) {
                    try { updateReuleauxAmountEnabled(); } catch (e) { }
                    return;
                }

                /* ルーローは奇数辺（3、5、7…）のみ。円（0）と偶数辺は対象外
                   Reuleaux applies to odd side counts only, never to a circle or an even count */
                var isEnabled = (typeof sidesValue === "number") && (sidesValue > 0) && (sidesValue % 2 === 1);
                reuleauxCheck.enabled = isEnabled;
                if (!isEnabled) reuleauxCheck.value = false;
                updateReuleauxAmountEnabled();
            } catch (e) { }
        }

        /**
         * スーパー楕円の指数を有効範囲に収める。
         * @param {number} value - 入力値
         * @returns {number} 小数第1位に丸めた指数
         */
        function clampSuperExponent(value) {
            value = Number(value);
            if (isNaN(value)) value = SHAPE_DEFAULTS.superExponent;
            if (value < SHAPE_RANGES.superExponent[0]) value = SHAPE_RANGES.superExponent[0];
            if (value > SHAPE_RANGES.superExponent[1]) value = SHAPE_RANGES.superExponent[1];
            value = Math.round(value * 10) / 10;
            return value;
        }

        /**
         * スーパー楕円の指数を入力欄とスライダーの両方に反映する。
         * @param {number} value - 入力値
         * @returns {number} 反映した指数
         */
        function syncSuperExponentUI(value) {
            value = clampSuperExponent(value);
            try {
                superExponentInput.text = String(value);
                superExponentSlider.value = value;
            } catch (e) { }
            return value;
        }

        /**
         * ルーローの度合いを有効範囲に収める。
         * @param {number} value - 入力値
         * @returns {number} 整数に丸めた度合い（%）
         */
        function clampReuleauxAmount(value) {
            value = Math.round(Number(value));
            if (isNaN(value)) value = SHAPE_DEFAULTS.reuleauxAmount;
            if (value < SHAPE_RANGES.reuleauxAmount[0]) value = SHAPE_RANGES.reuleauxAmount[0];
            if (value > SHAPE_RANGES.reuleauxAmount[1]) value = SHAPE_RANGES.reuleauxAmount[1];
            return value;
        }

        /**
         * ルーローの度合いを入力欄とスライダーの両方に反映する。
         * @param {number} value - 入力値
         * @returns {number} 反映した度合い（%）
         */
        function syncReuleauxAmountUI(value) {
            value = clampReuleauxAmount(value);
            try {
                reuleauxAmountInput.text = String(value);
                reuleauxAmountSlider.value = value;
            } catch (e) { }
            return value;
        }

        /**
         * ルーローの度合いの入力群を有効・無効にする。
         * @returns {void}
         */
        function updateReuleauxAmountEnabled() {
            try {
                var isEnabled = (reuleauxCheck.enabled && reuleauxCheck.value);
                reuleauxAmountInput.enabled = isEnabled;
                reuleauxAmountSlider.enabled = isEnabled;
            } catch (e) { }
        }

        /**
         * スーパー楕円の指数入力と、円のアンカーポイントパネルの有効状態を更新する。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateSuperEllipseControlsEnabled(sidesValue) {
            var isSuperEllipseActive = (superEllipseCheck.value && sidesValue === 0);
            try {
                superExponentInput.enabled = isSuperEllipseActive;
                superExponentSlider.enabled = isSuperEllipseActive;
                /* スーパー楕円がONのときはアンカーポイント数を選べない
                   The anchor count cannot be chosen while the superellipse is on */
                circleAnchorPanel.enabled = !isSuperEllipseActive;
            } catch (e) { }
        }

        /* 三角形パネルは回転パネルの中に置く / The triangle panel lives inside the rotation panel */
        var trianglePanel = rotatePanel.add("panel", undefined, LABELS.panel.triangle[lang]);
        setupPanel(trianglePanel, DIALOG_PANEL_SPACING);

        /* 角丸パネル / Corner smoothing panel */
        var cornerSmoothingPanel = rightColumn.add("panel", undefined, LABELS.panel.cornerSmoothing[lang]);
        setupPanel(cornerSmoothingPanel, DIALOG_PANEL_SPACING);

        var cornerRadiusRow = cornerSmoothingPanel.add("group");
        cornerRadiusRow.orientation = "row";
        cornerRadiusRow.alignChildren = ["left", "center"];
        cornerRadiusRow.spacing = 6;

        var cornerRadiusCheck = cornerRadiusRow.add("checkbox", undefined, LABELS.checkbox.cornerRadius[lang]);
        cornerRadiusCheck.value = false;
        var cornerRadiusInput = cornerRadiusRow.add("edittext", undefined, SHAPE_DEFAULTS.cornerRadius);
        cornerRadiusInput.characters = 5;
        changeValueByArrowKey(cornerRadiusInput);
        cornerRadiusRow.add("statictext", undefined, unitLabel);

        /**
         * 角丸の入力群を［半径］チェックの状態に合わせて有効・無効にする。
         * @returns {void}
         */
        function updateCornerRadiusInputEnabled() {
            var isEnabled = cornerRadiusCheck.value;
            cornerRadiusInput.enabled = isEnabled;
            smoothingSlider.enabled = isEnabled;
            smoothingValueLabel.enabled = isEnabled;
        }

        var smoothingLabelRow = cornerSmoothingPanel.add("group");
        smoothingLabelRow.orientation = "row";
        smoothingLabelRow.alignChildren = ["left", "center"];
        smoothingLabelRow.spacing = 8;

        smoothingLabelRow.add("statictext", undefined, LABELS.label.smoothing[lang]);
        var smoothingValueLabel = smoothingLabelRow.add("statictext", undefined, String(SHAPE_DEFAULTS.smoothing));
        smoothingValueLabel.characters = 4;

        var smoothingSlider = cornerSmoothingPanel.add("slider", undefined, SHAPE_DEFAULTS.smoothing, SHAPE_RANGES.smoothing[0], SHAPE_RANGES.smoothing[1]);
        smoothingSlider.preferredSize.width = UI_CONFIG.sliderWidth;

        /**
         * 角丸の半径の既定値を、現在の幅に対する比率から求めて入力欄に入れる。
         * @returns {void}
         */
        function applyDefaultCornerRadius() {
            try {
                var currentWidth = parseFloat(sizeInput.text);
                if (isNaN(currentWidth) || currentWidth <= 0) return;
                var defaultRadius = Math.round(currentWidth * SHAPE_DEFAULTS.cornerRadiusRatio * 10) / 10;
                cornerRadiusInput.text = String(defaultRadius);
            } catch (e) { }
        }

        /**
         * 角丸パネルの有効・無効を辺の数に応じて切り替える。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function updateCornerSmoothingEnabled(sidesValue) {
            var isEnabled = (sidesValue === 4);
            cornerSmoothingPanel.enabled = isEnabled;

            /* 正方形以外ではパネルを無効にするだけで入力値は破棄しない。
               角丸は sides === 4 のときしか参照されないため、値を残しても影響はない
               Outside a square the panel is only disabled; the values are kept, and they are
               read only when sides === 4, so keeping them is harmless */
            if (isEnabled && !(parseFloat(cornerRadiusInput.text) > 0)) {
                /* 有効化したときは幅に対する比率で既定値を入れる / Restore the ratio-based default when enabled */
                applyDefaultCornerRadius();
            }
        }

        cornerRadiusCheck.onClick = function () {
            updateCornerRadiusInputEnabled();
            refreshLiveShapeAvailabilityFromUI();
            updatePreview();
        };
        updateCornerRadiusInputEnabled();

        cornerRadiusInput.onChanging = function () {
            refreshLiveShapeAvailabilityFromUI();
            updatePreview();
        };

        smoothingSlider.onChanging = function () {
            var value = Math.round(smoothingSlider.value);
            smoothingValueLabel.text = String(value);
            updatePreview();
        };

        /* アンカーポイント操作パネル / Anchor operations panel */
        var anchorOpsPanel = rightColumn.add("panel", undefined, LABELS.panel.anchorOps[lang]);
        setupPanel(anchorOpsPanel, DIALOG_PANEL_SPACING);

        /* オプションパネル / Options panel */
        var optionPanel = rightColumn.add("panel", undefined, LABELS.panel.option[lang]);
        setupPanel(optionPanel, DIALOG_PANEL_SPACING);

        var liveShapeCheck = optionPanel.add("checkbox", undefined, LABELS.checkbox.liveShape[lang]);
        liveShapeCheck.value = true;

        /* ラフ効果でアンカーを追加 / Add anchors with the Roughen effect */
        var roughenAnchorsRow = anchorOpsPanel.add("group");
        roughenAnchorsRow.orientation = "row";
        roughenAnchorsRow.alignChildren = ["left", "center"];
        roughenAnchorsRow.spacing = 4;
        var roughenAnchorsCheck = roughenAnchorsRow.add("checkbox", undefined, LABELS.checkbox.roughenAnchors[lang]);
        roughenAnchorsCheck.value = false;
        var roughenAnchorsInput = roughenAnchorsRow.add("edittext", undefined, SHAPE_DEFAULTS.roughenDetail);
        roughenAnchorsInput.characters = 3;
        roughenAnchorsInput.enabled = false;
        changeValueByArrowKey(roughenAnchorsInput);

        roughenAnchorsCheck.onClick = function () {
            var isRoughenOn = roughenAnchorsCheck.value;
            roughenAnchorsInput.enabled = isRoughenOn;

            if (isRoughenOn) {
                splitAtAnchorsCheck.value = false;
                splitAtAnchorsCheck.enabled = false;
                liveShapeCheck.value = false;
                liveShapeCheck.enabled = false;
            } else {
                splitAtAnchorsCheck.enabled = true;
            }

            refreshLiveShapeAvailabilityFromUI();
            if (isRoughenOn) {
                liveShapeCheck.value = false;
                liveShapeCheck.enabled = false;
            }
            updatePreview();
        };
        roughenAnchorsInput.onChanging = function () { updatePreview(); };

        /* アンカーポイントで分割 / Split at anchor points */
        var splitAtAnchorsCheck = anchorOpsPanel.add("checkbox", undefined, LABELS.checkbox.splitAtAnchors[lang]);
        splitAtAnchorsCheck.value = false;

        var strokeCapRow = anchorOpsPanel.add("group");
        strokeCapRow.orientation = "row";
        strokeCapRow.alignChildren = ["left", "center"];
        strokeCapRow.spacing = 4;

        var strokeCapLabel = strokeCapRow.add("statictext", undefined, LABELS.label.strokeCap[lang]);
        var capButtRadio = strokeCapRow.add("radiobutton", undefined, LABELS.radio.capButt[lang]);
        var capRoundRadio = strokeCapRow.add("radiobutton", undefined, LABELS.radio.capRound[lang]);
        var capProjectingRadio = strokeCapRow.add("radiobutton", undefined, LABELS.radio.capProjecting[lang]);
        capButtRadio.value = true;

        /**
         * 線端の選択肢を［アンカーポイントで分割］の状態に合わせて有効・無効にする。
         * @returns {void}
         */
        function updateStrokeCapEnabled() {
            var isEnabled = splitAtAnchorsCheck.value;
            strokeCapLabel.enabled = isEnabled;
            capButtRadio.enabled = isEnabled;
            capRoundRadio.enabled = isEnabled;
            capProjectingRadio.enabled = isEnabled;
        }
        updateStrokeCapEnabled();

        capButtRadio.onClick = function () { updatePreview(); };
        capRoundRadio.onClick = function () { updatePreview(); };
        capProjectingRadio.onClick = function () { updatePreview(); };

        /**
         * 選択されている線端の種類を取得する。
         * @returns {StrokeCap} 線端の種類
         */
        function getSelectedStrokeCap() {
            if (capRoundRadio.value) return StrokeCap.ROUNDENDCAP;
            if (capProjectingRadio.value) return StrokeCap.PROJECTINGENDCAP;
            return StrokeCap.BUTTENDCAP;
        }

        /* ルーロー（定幅図形） / Reuleaux (constant-width) */
        var reuleauxCheck = optionPanel.add("checkbox", undefined, LABELS.checkbox.reuleaux[lang]);
        reuleauxCheck.value = false;

        var reuleauxAmountRow = optionPanel.add("group");
        reuleauxAmountRow.orientation = "row";
        reuleauxAmountRow.alignChildren = ["left", "center"];
        reuleauxAmountRow.spacing = 8;

        var reuleauxAmountInput = reuleauxAmountRow.add("edittext", undefined, String(SHAPE_DEFAULTS.reuleauxAmount));
        reuleauxAmountInput.characters = 4;
        changeValueByArrowKey(reuleauxAmountInput);

        var reuleauxAmountSlider = reuleauxAmountRow.add("slider", undefined, SHAPE_DEFAULTS.reuleauxAmount, SHAPE_RANGES.reuleauxAmount[0], SHAPE_RANGES.reuleauxAmount[1]);
        reuleauxAmountSlider.preferredSize.width = UI_CONFIG.shortSliderWidth;

        /**
         * ライブシェイプ化の可否を、排他条件から決める。
         * @param {boolean} isSplit - アンカーポイントで分割するか
         * @param {boolean} isSuperEllipseActive - スーパー楕円が有効か
         * @param {boolean} isCustomCircleAnchors - 円のアンカー数が4以外か
         * @param {boolean} isReuleaux - ルーロー、またはラフ効果が有効か
         * @param {boolean} isCornerSmoothing - 角丸が有効か
         * @returns {void}
         */
        function updateLiveShapeAvailability(isSplit, isSuperEllipseActive, isCustomCircleAnchors, isReuleaux, isCornerSmoothing) {
            if (isSplit || isSuperEllipseActive || isCustomCircleAnchors || isReuleaux || isCornerSmoothing) {
                liveShapeCheck.value = false;
                liveShapeCheck.enabled = false;
            } else {
                liveShapeCheck.enabled = true;
            }
        }

        /**
         * 現在のUIの状態からライブシェイプ化の可否を再計算する。
         * @returns {void}
         */
        function refreshLiveShapeAvailabilityFromUI() {
            try {
                var currentSides = getSelectedSideValue(sideRadios, customSidesInput);
                var isSuperEllipseActive = (superEllipseCheck.value && currentSides === 0);
                var anchorCount = getCircleAnchorCountFromRadios();
                if (!anchorCount || anchorCount < 2) anchorCount = 2;
                /* ライブシェイプ化できるのは円のアンカーが4のときだけ（スーパー楕円を除く）
                   A live shape is possible only with four circle anchors and no superellipse */
                var isCustomCircleAnchors = (currentSides === 0 && !isSuperEllipseActive && anchorCount !== 4);
                var isCornerSmoothingActive = (currentSides === 4 && cornerRadiusCheck.value && parseFloat(cornerRadiusInput.text) > 0);
                var isRoughenActive = roughenAnchorsCheck.value;
                updateLiveShapeAvailability(splitAtAnchorsCheck.value, isSuperEllipseActive, isCustomCircleAnchors, reuleauxCheck.value || isRoughenActive, isCornerSmoothingActive);
            } catch (e) {
                /* 判定できないときは安全側に倒す / Fall back to the safe side when the state cannot be read */
                try {
                    liveShapeCheck.value = false;
                    liveShapeCheck.enabled = false;
                } catch (err) { }
            }
        }

        /**
         * 回転がOFFのときに使う自動角度を、回転の入力欄に反映する。
         * 辺の数が変わったときは回転がONでも呼び出す。
         * @param {number} sidesValue - 現在の辺の数
         * @returns {void}
         */
        function applyAutoRotationForSides(sidesValue) {
            var angle;
            if (sidesValue === 0) {
                angle = 45;
            } else if (sidesValue >= 3) {
                angle = 360 / (sidesValue * 2);
            } else {
                return;
            }
            rotateInput.text = formatAngle(angle);
        }

        /**
         * 回転をONにしたときの既定角度を決めて入力欄に反映する。
         * 円（辺の数0）はアンカー数ごとに 2→90、3→180、4→45、5→180、6→30 を使う。
         * @returns {void}
         */
        function applyDefaultRotationWhenEnablingRotate() {
            try {
                var currentSides = getSelectedSideValue(sideRadios, customSidesInput);
                var isSuperEllipseActive = (superEllipseCheck.value && currentSides === 0);
                var anchorCount = getCircleAnchorCountFromRadios();
                if (!anchorCount || anchorCount < 2) anchorCount = 2;
                if (currentSides === 0 && !isSuperEllipseActive) {
                    var angle;
                    if (anchorCount === 2) angle = 90;
                    else if (anchorCount === 3) angle = 180;
                    else if (anchorCount === 4) angle = 45;
                    else if (anchorCount === 5) angle = 180;
                    else if (anchorCount === 6) angle = 30;
                    else angle = 45;
                    rotateInput.text = formatAngle(angle);
                } else {
                    applyAutoRotationForSides(currentSides);
                    /* 三角形のときは［下］（60°）を既定にする / A triangle defaults to "down" (60 degrees) */
                    if (currentSides === 3) {
                        try {
                            triangleDownRadio.value = true;
                        } catch (e) { }
                    }
                }
            } catch (e) { }
        }

        /**
         * 回転を強制的にOFFにする。
         * @returns {void}
         */
        function forceRotateOff() {
            rotateCheck.value = false;
            rotateInput.enabled = false;
            rotateUnitLabel.enabled = false;
        }

        splitAtAnchorsCheck.onClick = function () {
            if (splitAtAnchorsCheck.value) {
                fillCheck.value = false;
                strokeCheck.value = true;
                updateStrokeWidthEnabled();
            }
            updateStrokeCapEnabled();
            refreshLiveShapeAvailabilityFromUI();
            updateSuperEllipseControlsEnabled(getSelectedSideValue(sideRadios, customSidesInput));
            updatePreview();
        };

        var triangleRow = trianglePanel.add("group");
        triangleRow.orientation = "row";
        triangleRow.alignChildren = ["left", "center"];
        triangleRow.spacing = 10;

        var triangleRightRadio = triangleRow.add("radiobutton", undefined, LABELS.radio.triangleRight[lang]);
        var triangleLeftRadio = triangleRow.add("radiobutton", undefined, LABELS.radio.triangleLeft[lang]);
        var triangleDownRadio = triangleRow.add("radiobutton", undefined, LABELS.radio.triangleDown[lang]);
        triangleRightRadio.value = true;

        /**
         * 三角形の向きを変えたときの処理。回転を必ずONにしてプレビューを更新する。
         * @returns {void}
         */
        function onTriangleDirectionChange() {
            rotateCheck.value = true;
            rotateInput.enabled = true;
            rotateUnitLabel.enabled = true;
            updatePreview();
        }

        triangleRightRadio.onClick = onTriangleDirectionChange;
        triangleLeftRadio.onClick = onTriangleDirectionChange;
        triangleDownRadio.onClick = onTriangleDirectionChange;

        /**
         * スターと五芒星、およびルーローとの排他関係を整える。
         * @returns {void}
         */
        function validateStarAndPentagram() {
            /* スターパネルが無効（円）ならスター関連を強制的にOFF / Force the star options off while the panel is disabled */
            if (!starPanel.enabled) {
                starCheck.enabled = false;
                starCheck.value = false;
                pentagramCheck.value = false;
                pentagramCheck.enabled = false;
                return;
            }

            starCheck.enabled = true;
            if (!starCheck.value) pentagramCheck.value = false;
            pentagramCheck.enabled = starCheck.value;

            /* ルーローはスターと併用できない / Reuleaux is not compatible with a star */
            if (starCheck.value) {
                reuleauxCheck.value = false;
                reuleauxCheck.enabled = false;
            }
            try { updateReuleauxAmountEnabled(); } catch (e) { }

            if (pentagramCheck.value) {
                /* ［それ以外］は別グループなのでラジオの排他が効かない。全件を明示的に解除する
                   The custom radio lives in another group, so clear every radio explicitly */
                for (var i = 0; i < sideRadios.length; i++) sideRadios[i].value = false;
                sideRadios[3].value = true;
                setCustomSidesEnabled(false);
                applyAutoRotationForSides(5);
                forceRotateOff();
            }

            /* スターがOFFに戻ったら奇数辺の条件でルーローを復帰させる
               Once the star is off again, restore Reuleaux under the odd-side rule */
            if (!starCheck.value) {
                try {
                    updateReuleauxAvailability(getSelectedSideValue(sideRadios, customSidesInput));
                } catch (e) { }
            }
            updateInnerRadiusEnabled();
        }

        /**
         * 現在のUIから図形生成のパラメーターを組み立てる（プレビューと確定の両方で使う）。
         * @returns {object} createShapeに渡すパラメーター一式
         */
        function getCurrentParams() {
            validateStarAndPentagram();

            var sides = getSelectedSideValue(sideRadios, customSidesInput);
            trianglePanel.enabled = (sides === 3);
            updateCirclePanelEnabled(sides);
            updateCornerSmoothingEnabled(sides);
            updateStarPanelEnabled(sides);
            updateReuleauxAvailability(sides);

            var size = parseFloat(sizeInput.text) * unitFactor;
            var innerRatio = parseFloat(innerRatioInput.text);
            var isStar = starCheck.value;
            var isPentagram = pentagramCheck.value;
            var rotateEnabled = rotateCheck.value;
            var angle = parseFloat(rotateInput.text);
            var splitAtAnchors = splitAtAnchorsCheck.value;
            var useSuperEllipse = superEllipseCheck.value && (sides === 0);
            roughenAnchorsUseMenuFallback = (sides !== 0 && Math.round(Number(roughenAnchorsInput.text)) === 1);

            /* ラフ効果の詳細。メニューコマンド経由の経路はプレビューできないので0にする
               Roughen detail; the menu-command path cannot be previewed, so it is reported as 0 */
            var roughenDetail = 0;
            if (roughenAnchorsCheck.value && !roughenAnchorsUseMenuFallback) {
                roughenDetail = parseFloat(roughenAnchorsInput.text);
                if (isNaN(roughenDetail) || roughenDetail < 0) roughenDetail = 0;
            }

            var useReuleaux = reuleauxCheck.value;
            var reuleauxAmount = clampReuleauxAmount(reuleauxAmountInput.text) / 100;
            var superExponent = clampSuperExponent(superExponentInput.text);

            /* 円のアンカーポイント数は、円かつスーパー楕円OFFのときだけ有効
               The circle anchor count only matters for a circle without the superellipse */
            var anchorCount = getCircleAnchorCountFromRadios();
            if (!anchorCount || anchorCount < 2) anchorCount = 2;
            var circleAnchorCount = (sides === 0 && !useSuperEllipse) ? anchorCount : SHAPE_DEFAULTS.circleAnchors;

            /* refreshLiveShapeAvailabilityFromUI()が同じ判定を行うので、そちらに任せる
               refreshLiveShapeAvailabilityFromUI() recomputes the same condition, so let it decide */
            refreshLiveShapeAvailabilityFromUI();

            /* スーパー楕円は回転を強制的にOFFにする / The superellipse forces the rotation off */
            if (useSuperEllipse) {
                forceRotateOff();
            }

            /* 回転がOFFのときだけ自動角度を使う（ONなら入力値を尊重）
               The auto angle is used only while the rotation is off; manual mode keeps the entered angle */
            if (!rotateEnabled) {
                if (sides === 0) {
                    angle = 45;
                    rotateInput.text = "45";
                } else if (sides >= 3) {
                    angle = 360 / (sides * 2);
                    rotateInput.text = formatAngle(angle);
                }
            }

            /* 三角形の向きごとの回転角。右＝-90°、左＝90°、下＝60°
               Rotation per triangle direction: right -90, left 90, down 60 */
            if (sides === 3 && trianglePanel.enabled) {
                if (triangleRightRadio.value) {
                    angle = -90;
                } else if (triangleLeftRadio.value) {
                    angle = 90;
                } else if (triangleDownRadio.value) {
                    angle = 60;
                }
                rotateInput.text = formatAngle(angle);
            }

            if (isStar && isPentagram && sides === 5) {
                innerRatio = (3 - Math.sqrt(5)) / 2 * 100;
                innerRatioInput.text = innerRatio.toFixed(2);
            }

            return {
                size: size,
                sides: sides,
                isStar: isStar,
                innerRatio: innerRatio,
                rotateEnabled: rotateEnabled,
                angle: angle,
                splitAtAnchors: splitAtAnchors,
                strokeCap: splitAtAnchors ? getSelectedStrokeCap() : null,
                useSuperEllipse: useSuperEllipse,
                superExponent: superExponent,
                circleAnchorCount: circleAnchorCount,
                useReuleaux: useReuleaux,
                reuleauxAmount: reuleauxAmount,
                fillOpts: {
                    enabled: fillCheck.value,
                    color: fillSwatch._aiColor
                },
                strokeOpts: {
                    enabled: strokeCheck.value,
                    color: strokeSwatch._aiColor,
                    widthPt: parseFloat(strokeWidthInput.text) * strokeUnitInfo.factorToPt
                },
                cornerSmoothing: (sides === 4 && cornerRadiusCheck.value) ? {
                    radius: parseFloat(cornerRadiusInput.text) * unitFactor,
                    smoothing: Math.round(smoothingSlider.value)
                } : null,
                opacity: clampOpacity(opacityInput.text),
                roughenDetail: roughenDetail
            };
        }

        /**
         * 入力内容からプレビューを描き直す（Undo履歴を汚さない）。
         * @returns {void}
         */
        function updatePreview() {
            /* 新しいプレビューの前に必ず前回分を巻き戻す / Always roll back the previous preview first */
            previewManager.rollback();
            previewShape = null;

            var params = getCurrentParams();
            if (!isNaN(params.size) && !isNaN(params.innerRatio)) {
                previewManager.addStep(function () {
                    previewShape = createShape(app.activeDocument, params.size, params.sides, params.isStar, params.innerRatio, params.rotateEnabled, params.angle, params.splitAtAnchors, params.useSuperEllipse, params.superExponent, params.circleAnchorCount, params.useReuleaux, params.reuleauxAmount, params.fillOpts, params.strokeOpts, params.cornerSmoothing, params.strokeCap, params.opacity);
                    /* ラフ効果も同じステップ内で適用してプレビューに反映する
                       The roughen effect runs inside the same step so the preview shows it */
                    applyRoughenEffect(previewShape, params.roughenDetail);
                });
            }
        }

        superExponentInput.onChanging = function () {
            syncSuperExponentUI(superExponentInput.text);
            updatePreview();
        };

        superExponentSlider.onChanging = function () {
            syncSuperExponentUI(superExponentSlider.value);
            updatePreview();
        };

        /* イベントの割り当て / Event bindings */
        starCheck.onClick = function () {
            validateStarAndPentagram();
            updateInnerRadiusEnabled();
            updatePreview();
        };
        pentagramCheck.onClick = function () {
            if (pentagramCheck.value) {
                forceRotateOff();
            }
            updatePreview();
        };
        superEllipseCheck.onClick = function () {
            /* スーパー楕円は円（辺の数0）でだけ効く / The superellipse only applies to a circle */
            var currentSides = getSelectedSideValue(sideRadios, customSidesInput);
            if (superEllipseCheck.value && currentSides === 0) {
                forceRotateOff();
            }
            try {
                var isSuperEllipseActive = (superEllipseCheck.value && currentSides === 0);
                circleAnchorRadios.anchors2.enabled = !isSuperEllipseActive;
                circleAnchorRadios.anchors3.enabled = !isSuperEllipseActive;
                circleAnchorRadios.anchors4.enabled = !isSuperEllipseActive;
                circleAnchorRadios.anchors5.enabled = !isSuperEllipseActive;
                circleAnchorRadios.anchors6.enabled = !isSuperEllipseActive;
            } catch (e) { }
            refreshLiveShapeAvailabilityFromUI();
            updateSuperEllipseControlsEnabled(currentSides);
            updatePreview();
        };
        innerRatioInput.onChanging = function () {
            if (pentagramCheck.value) pentagramCheck.value = false;
            var value = Math.round(Number(innerRatioInput.text));
            if (isNaN(value)) value = SHAPE_DEFAULTS.innerRatio;
            if (value < SHAPE_RANGES.innerRatio[0]) value = SHAPE_RANGES.innerRatio[0];
            if (value > SHAPE_RANGES.innerRatio[1]) value = SHAPE_RANGES.innerRatio[1];
            innerRatioInput.text = String(value);
            innerRatioSlider.value = value;
            updatePreview();
        };

        innerRatioSlider.onChanging = function () {
            var value = Math.round(innerRatioSlider.value);
            innerRatioInput.text = String(value);
            updatePreview();
        };
        sizeInput.onChanging = updatePreview;

        reuleauxCheck.onClick = function () {
            if (reuleauxCheck.value) {
                /* 有効にするたび既定値（100%）に戻す / Reset to the default amount whenever it is enabled */
                syncReuleauxAmountUI(SHAPE_DEFAULTS.reuleauxAmount);
            }
            updateReuleauxAmountEnabled();
            updatePreview();
        };

        reuleauxAmountInput.onChanging = function () {
            syncReuleauxAmountUI(reuleauxAmountInput.text);
            updatePreview();
        };

        reuleauxAmountSlider.onChanging = function () {
            syncReuleauxAmountUI(reuleauxAmountSlider.value);
            updatePreview();
        };

        rotateInput.onChanging = updatePreview;
        rotateCheck.onClick = function () {
            rotateInput.enabled = rotateCheck.value;
            rotateUnitLabel.enabled = rotateCheck.value;
            if (rotateCheck.value) {
                applyDefaultRotationWhenEnablingRotate();
            }
            updatePreview();
        };
        customSidesInput.onChanging = function () {
            var value = Math.round(Number(customSidesInput.text));
            if (isNaN(value)) value = SHAPE_DEFAULTS.customSides;
            if (value < SHAPE_RANGES.customSides[0]) value = SHAPE_RANGES.customSides[0];
            if (value > SHAPE_RANGES.customSides[1]) value = SHAPE_RANGES.customSides[1];
            customSidesInput.text = String(value);
            customSidesSlider.value = value;
            try { updateReuleauxAvailability(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            updatePreview();
        };

        customSidesSlider.onChanging = function () {
            var value = Math.round(customSidesSlider.value);
            customSidesInput.text = String(value);
            try { updateReuleauxAvailability(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            updatePreview();
        };

        /**
         * 円のアンカーポイント数を切り替えたときの処理。
         * @returns {void}
         */
        function onCircleAnchorRadioClick() {
            refreshLiveShapeAvailabilityFromUI();
            updatePreview();
        }
        circleAnchorRadios.anchors2.onClick = onCircleAnchorRadioClick;
        circleAnchorRadios.anchors3.onClick = onCircleAnchorRadioClick;
        circleAnchorRadios.anchors4.onClick = onCircleAnchorRadioClick;
        circleAnchorRadios.anchors5.onClick = onCircleAnchorRadioClick;
        circleAnchorRadios.anchors6.onClick = onCircleAnchorRadioClick;

        for (var i = 0; i <= 6; i++) {
            (function (radioIndex) {
                sideRadios[radioIndex].onClick = function () {
                    if (pentagramCheck.value) pentagramCheck.value = false;
                    if (radioIndex === 6) {
                        for (var j = 0; j < 6; j++) sideRadios[j].value = false;
                        setCustomSidesEnabled(true);
                    } else {
                        sideRadios[6].value = false;
                        setCustomSidesEnabled(false);
                    }
                    /* 辺の数が変わったら回転がONでも角度を更新する / Update the angle on a side change, even while rotation is on */
                    try { applyAutoRotationForSides(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
                    try { updateReuleauxAvailability(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
                    try { updateCornerSmoothingEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
                    refreshLiveShapeAvailabilityFromUI();
                    updateSuperEllipseControlsEnabled(getSelectedSideValue(sideRadios, customSidesInput));
                    updatePreview();
                };
            })(i);
        }

        /**
         * セッション状態をUIに復元する。
         * @param {object} state - セッション状態
         * @returns {void}
         */
        function applyStateToUI(state) {
            if (!state) return;

            /* 辺の数の選択 / Side count selection */
            if (typeof state.selectedSideIndex === "number" && state.selectedSideIndex >= 0 && state.selectedSideIndex < sideRadios.length) {
                for (var i = 0; i < sideRadios.length; i++) sideRadios[i].value = false;
                sideRadios[state.selectedSideIndex].value = true;
                if (state.selectedSideIndex === 6) {
                    setCustomSidesEnabled(true);
                    if (typeof state.customSidesText === "string") customSidesInput.text = state.customSidesText;
                    /* スライダーのつまみも復元した辺の数に合わせる / Move the slider thumb to the restored side count */
                    var restoredCustomSides = Math.round(Number(customSidesInput.text));
                    if (!isNaN(restoredCustomSides)) {
                        if (restoredCustomSides < SHAPE_RANGES.customSides[0]) restoredCustomSides = SHAPE_RANGES.customSides[0];
                        if (restoredCustomSides > SHAPE_RANGES.customSides[1]) restoredCustomSides = SHAPE_RANGES.customSides[1];
                        customSidesSlider.value = restoredCustomSides;
                    }
                } else {
                    setCustomSidesEnabled(false);
                }
            }

            if (typeof state.sizeText === "string") sizeInput.text = state.sizeText;

            /* 塗りと線 / Fill and stroke */
            if (typeof state.fillCheck === "boolean") fillCheck.value = state.fillCheck;
            if (typeof state.strokeCheck === "boolean") strokeCheck.value = state.strokeCheck;
            if (typeof state.strokeWidthText === "string") strokeWidthInput.text = state.strokeWidthText;
            if (state.fillColorString) {
                try { fillSwatch._aiColor = pickerStringToAiColor(state.fillColorString); } catch (e) { }
            }
            if (state.strokeColorString) {
                try { strokeSwatch._aiColor = pickerStringToAiColor(state.strokeColorString); } catch (e) { }
            }
            try { updateStrokeWidthEnabled(); } catch (e) { }

            /* 不透明度 / Opacity */
            if (typeof state.opacityText === "string") {
                opacityInput.text = state.opacityText;
                try { opacitySlider.value = clampOpacity(state.opacityText); } catch (e) { }
            }

            /* 角丸 / Corner smoothing */
            if (typeof state.cornerRadiusCheck === "boolean") cornerRadiusCheck.value = state.cornerRadiusCheck;
            if (typeof state.cornerRadiusText === "string") cornerRadiusInput.text = state.cornerRadiusText;
            if (typeof state.smoothingValue === "number") {
                var restoredSmoothing = Math.round(state.smoothingValue);
                if (restoredSmoothing < SHAPE_RANGES.smoothing[0]) restoredSmoothing = SHAPE_RANGES.smoothing[0];
                if (restoredSmoothing > SHAPE_RANGES.smoothing[1]) restoredSmoothing = SHAPE_RANGES.smoothing[1];
                smoothingSlider.value = restoredSmoothing;
                smoothingValueLabel.text = String(restoredSmoothing);
            }
            try { updateCornerRadiusInputEnabled(); } catch (e) { }

            /* 回転 / Rotation */
            if (typeof state.rotateCheck === "boolean") rotateCheck.value = state.rotateCheck;
            if (typeof state.rotateText === "string") rotateInput.text = state.rotateText;
            rotateInput.enabled = rotateCheck.value;
            rotateUnitLabel.enabled = rotateCheck.value;

            /* スターと五芒星 / Star and pentagram */
            if (typeof state.starCheck === "boolean") starCheck.value = state.starCheck;
            if (typeof state.pentagramCheck === "boolean") pentagramCheck.value = state.pentagramCheck;
            if (typeof state.innerRatioText === "string") innerRatioInput.text = state.innerRatioText;
            try {
                var restoredInnerRatio = Math.round(Number(innerRatioInput.text));
                if (!isNaN(restoredInnerRatio)) innerRatioSlider.value = restoredInnerRatio;
            } catch (e) { }
            try { updateInnerRadiusEnabled(); } catch (e) { }
            if (typeof state.superEllipseCheck === "boolean") superEllipseCheck.value = state.superEllipseCheck;
            if (typeof state.superExponentText === "string") {
                superExponentInput.text = state.superExponentText;
                superExponentSlider.value = clampSuperExponent(state.superExponentText);
            }
            if (typeof state.circleAnchorsValue === "number") {
                circleAnchorRadios.anchors2.value = (state.circleAnchorsValue === 2);
                circleAnchorRadios.anchors3.value = (state.circleAnchorsValue === 3);
                circleAnchorRadios.anchors4.value = (state.circleAnchorsValue === 4);
                circleAnchorRadios.anchors5.value = (state.circleAnchorsValue === 5);
                circleAnchorRadios.anchors6.value = (state.circleAnchorsValue === 6);
                if (!circleAnchorRadios.anchors2.value && !circleAnchorRadios.anchors3.value && !circleAnchorRadios.anchors4.value && !circleAnchorRadios.anchors5.value && !circleAnchorRadios.anchors6.value) {
                    circleAnchorRadios.anchors4.value = true;
                }
            }

            /* 三角形の向き / Triangle direction */
            if (state.triangleDir === "right") {
                triangleRightRadio.value = true;
            } else if (state.triangleDir === "left") {
                triangleLeftRadio.value = true;
            } else if (state.triangleDir === "down") {
                triangleDownRadio.value = true;
            }

            /* ［アンカーポイントで分割］は毎回OFFで開く（復元しない）
               "Split at anchor points" always opens off and is never restored */
            splitAtAnchorsCheck.value = false;
            if (state.strokeCapValue === "round") {
                capRoundRadio.value = true;
            } else if (state.strokeCapValue === "projecting") {
                capProjectingRadio.value = true;
            } else {
                capButtRadio.value = true;
            }
            updateStrokeCapEnabled();
            if (typeof state.reuleauxCheck === "boolean") reuleauxCheck.value = state.reuleauxCheck;
            if (typeof state.liveShapeCheck === "boolean") liveShapeCheck.value = state.liveShapeCheck;
            if (typeof state.roughenAnchorsCheck === "boolean") roughenAnchorsCheck.value = state.roughenAnchorsCheck;
            if (typeof state.roughenAnchorsText === "string") roughenAnchorsInput.text = state.roughenAnchorsText;
            roughenAnchorsInput.enabled = roughenAnchorsCheck.value;
            if (roughenAnchorsCheck.value) {
                splitAtAnchorsCheck.value = false;
                splitAtAnchorsCheck.enabled = false;
                liveShapeCheck.value = false;
                liveShapeCheck.enabled = false;
            } else {
                splitAtAnchorsCheck.enabled = true;
            }

            /* 分割・スーパー楕円・アンカー数とライブシェイプ化の依存関係を反映
               Apply the dependency between split, superellipse, anchor count and the live shape */
            try {
                refreshLiveShapeAvailabilityFromUI();
            } catch (e) {
                if (splitAtAnchorsCheck.value) {
                    liveShapeCheck.value = false;
                    liveShapeCheck.enabled = false;
                } else {
                    liveShapeCheck.enabled = true;
                }
            }

            if (!starCheck.value) pentagramCheck.value = false;

            /* 五芒星やスーパー楕円が有効なら回転を強制的にOFF / Force the rotation off for a pentagram or a superellipse */
            try {
                var restoredSides = getSelectedSideValue(sideRadios, customSidesInput);
                if (pentagramCheck.value || (superEllipseCheck.value && restoredSides === 0)) {
                    forceRotateOff();
                }
                var isSuperEllipseActive = (superEllipseCheck.value && restoredSides === 0);
                circleAnchorRadios.anchors2.enabled = !isSuperEllipseActive;
                circleAnchorRadios.anchors3.enabled = !isSuperEllipseActive;
                circleAnchorRadios.anchors4.enabled = !isSuperEllipseActive;
                circleAnchorRadios.anchors5.enabled = !isSuperEllipseActive;
                circleAnchorRadios.anchors6.enabled = !isSuperEllipseActive;
            } catch (e) { }

            /* 選択内容に合わせて各パネルの有効状態を更新 / Refresh every panel against the current selection */
            try { updateCirclePanelEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            try { updateCornerSmoothingEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            try { updateStarPanelEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            try { updateReuleauxAvailability(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
        }

        /**
         * 現在のUIの状態をセッション状態に保存する。
         * @param {object} state - セッション状態
         * @returns {void}
         */
        function saveStateFromUI(state) {
            if (!state) return;

            var selectedIndex = 0;
            for (var i = 0; i < sideRadios.length; i++) {
                if (sideRadios[i].value) { selectedIndex = i; break; }
            }
            state.selectedSideIndex = selectedIndex;
            state.customSidesText = customSidesInput.text;

            state.sizeText = sizeInput.text;

            /* 塗りと線・不透明度・角丸 / Fill and stroke, opacity, corner smoothing */
            state.fillCheck = fillCheck.value;
            state.strokeCheck = strokeCheck.value;
            state.strokeWidthText = strokeWidthInput.text;
            /* カラーは常駐エンジンに残るオブジェクト参照ではなく文字列で保存する
               Colors are stored as strings, not as object references kept alive by the engine */
            try { state.fillColorString = aiColorToPickerString(fillSwatch._aiColor); } catch (e) { }
            try { state.strokeColorString = aiColorToPickerString(strokeSwatch._aiColor); } catch (e) { }
            state.opacityText = opacityInput.text;
            state.cornerRadiusCheck = cornerRadiusCheck.value;
            state.cornerRadiusText = cornerRadiusInput.text;
            state.smoothingValue = Math.round(smoothingSlider.value);

            state.rotateCheck = rotateCheck.value;
            state.rotateText = rotateInput.text;
            state.starCheck = starCheck.value;
            state.pentagramCheck = pentagramCheck.value;
            state.innerRatioText = innerRatioInput.text;
            state.superEllipseCheck = superEllipseCheck.value;
            state.superExponentText = superExponentInput.text;
            state.circleAnchorsValue = getCircleAnchorCountFromRadios();

            state.triangleDir = triangleRightRadio.value ? "right" : (triangleLeftRadio.value ? "left" : "down");

            /* 次回はOFFで開くが、今回の値としては保存しておく
               The option opens off next time, but the current value is still recorded */
            state.splitAtAnchorsCheck = splitAtAnchorsCheck.value;
            state.strokeCapValue = capRoundRadio.value ? "round" : (capProjectingRadio.value ? "projecting" : "butt");
            state.liveShapeCheck = liveShapeCheck.value;
            state.roughenAnchorsCheck = roughenAnchorsCheck.value;
            state.roughenAnchorsText = roughenAnchorsInput.text;
            state.reuleauxCheck = reuleauxCheck.value;
        }

        /* 最初のプレビューの前に復元した状態を反映 / Apply the restored state before the first preview */
        try { applyStateToUI(getSessionState()); } catch (e) { }
        try { updateReuleauxAmountEnabled(); } catch (e) { }
        try { updateInnerRadiusEnabled(); } catch (e) { }
        try { updateSuperEllipseControlsEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
        try { updateCornerSmoothingEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }

        dialog.onClose = function () {
            try { saveDialogLocation(dialog); } catch (err) { }
            try { saveStateFromUI(getSessionState()); } catch (e) { }

            /* キャンセル時はプレビューを巻き戻してUndo履歴を汚さない / Roll the preview back on cancel */
            if (!isConfirmed) {
                try { previewManager.rollback(); } catch (e) { }
                previewShape = null;
                try { if (initialView && initialZoom != null) initialView.zoom = initialZoom; } catch (err) { }
            }
        };

        dialog.onShow = function () {
            /* 環境によっては無視されるため不透明度を先に適用 / Apply the opacity first, some hosts ignore it */
            setDialogOpacity(dialog, UI_CONFIG.dialogOpacity);
            /* ［アンカーポイントで分割］は毎回OFFで開く / The split option always opens off */
            try { splitAtAnchorsCheck.value = false; } catch (err) { }
            try { refreshLiveShapeAvailabilityFromUI(); } catch (err) { }

            try { updateCirclePanelEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            try { updateCornerSmoothingEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            try { updateSuperEllipseControlsEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            try { updateStarPanelEnabled(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            try { updateReuleauxAvailability(getSelectedSideValue(sideRadios, customSidesInput)); } catch (e) { }
            try { updateReuleauxAmountEnabled(); } catch (e) { }
            try { updateInnerRadiusEnabled(); } catch (e) { }
            updatePreview();

            /* Illustratorの起動中は前回のダイアログ位置を再利用 / Reuse the last dialog position while Illustrator is running */
            var savedLocation = null;
            try { savedLocation = getSavedDialogLocation(); } catch (err) { savedLocation = null; }
            if (savedLocation) {
                try { dialog.location = savedLocation; } catch (err) { }
            } else {
                dialog.center();
                shiftDialogPosition(dialog, UI_CONFIG.dialogOffsetX, UI_CONFIG.dialogOffsetY);
            }
        };

        /* 画面ズーム / View zoom */
        var initialView = null;
        var initialZoom = null;

        var zoomGroup = dialog.add("group");
        zoomGroup.orientation = "row";
        zoomGroup.alignChildren = ["center", "center"];
        zoomGroup.alignment = "center";
        try { zoomGroup.margins = [0, 7, 0, 5]; } catch (err) { }

        zoomGroup.add("statictext", undefined, LABELS.label.viewZoom[lang]);

        try { initialZoom = (doc && doc.activeView) ? Number(doc.activeView.zoom) : 1; } catch (err) { initialZoom = 1; }
        var zoomSlider = zoomGroup.add("slider", undefined, (initialZoom != null ? initialZoom : 1), SHAPE_RANGES.zoom[0], SHAPE_RANGES.zoom[1]);
        zoomSlider.preferredSize.width = UI_CONFIG.zoomSliderWidth;

        /**
         * ドキュメントウィンドウのズーム倍率を変更する。
         * @param {number} zoomValue - ズーム倍率
         * @returns {void}
         */
        function applyZoom(zoomValue) {
            try {
                if (!initialView) initialView = doc.activeView;
                if (!initialView) return;
                initialView.zoom = zoomValue;
            } catch (err) { }
        }

        zoomSlider.onChanging = function () {
            applyZoom(Number(zoomSlider.value));
            try { app.redraw(); } catch (err) { }
        };

        var buttonArea = dialog.add("group");
        buttonArea.orientation = "row";
        buttonArea.alignChildren = ["fill", "center"];
        buttonArea.alignment = "fill";
        buttonArea.margins = [0, 10, 0, 0];

        var buttonLeftGroup = buttonArea.add("group");
        setupRow(buttonLeftGroup, ["left", "center"]);

        /* ドキュメントは通常プレビュー表示で開くので、ONから始めて表示と実態を合わせる
           A document normally opens in preview mode, so start on to keep the label truthful */
        var isViewPreviewOn = true;
        var previewButton = buttonLeftGroup.add("button", undefined, LABELS.button.preview[lang]);
        /* ボタンはコンテナいっぱいに広げない / Buttons must not stretch to the container width */
        previewButton.alignment = "left";

        var buttonRightGroup = buttonArea.add("group");
        setupRow(buttonRightGroup, ["right", "center"], 10);

        var cancelButton = buttonRightGroup.add("button", undefined, LABELS.button.cancel[lang], { name: "cancel" });
        var okButton = buttonRightGroup.add("button", undefined, LABELS.button.ok[lang], { name: "ok" });
        cancelButton.alignment = "right";
        okButton.alignment = "right";

        /**
         * ［プレビュー］ボタンの表示を現在の状態に合わせて更新する。
         * @returns {void}
         */
        function updatePreviewButtonText() {
            previewButton.text = isViewPreviewOn
                ? "● " + LABELS.button.preview[lang]
                : LABELS.button.preview[lang];
        }
        updatePreviewButtonText();

        previewButton.onClick = function () {
            isViewPreviewOn = !isViewPreviewOn;
            updatePreviewButtonText();
            try { app.executeMenuCommand('preview'); } catch (e) { }
        };

        var isConfirmed = false;
        var finalParams = null;
        cancelButton.onClick = function () {
            try { previewManager.rollback(); } catch (e) { }
            previewShape = null;
            try { if (initialView && initialZoom != null) initialView.zoom = initialZoom; } catch (err) { }
            dialog.close();
        };
        okButton.onClick = function () {
            /* ウィジェットが生きているうちにパラメーターを確定させる
               Capture the parameters while the dialog widgets are still alive */
            try { finalParams = getCurrentParams(); } catch (e) { finalParams = null; }
            applyLiveShape = liveShapeCheck.value;
            roughenAnchorsDetail = roughenAnchorsCheck.value ? parseFloat(roughenAnchorsInput.text) : 0;
            isConfirmed = true;
            dialog.close();
        };

        dialog.show();

        /* OKなら1回のUndoで取り消せる形で確定する / Finalize as a single undoable action when OK was pressed */
        if (isConfirmed) {
            previewManager.confirm(function () {
                if (!finalParams) return;
                if (isNaN(finalParams.size) || isNaN(finalParams.innerRatio)) return;
                previewShape = createShape(app.activeDocument, finalParams.size, finalParams.sides, finalParams.isStar, finalParams.innerRatio, finalParams.rotateEnabled, finalParams.angle, finalParams.splitAtAnchors, finalParams.useSuperEllipse, finalParams.superExponent, finalParams.circleAnchorCount, finalParams.useReuleaux, finalParams.reuleauxAmount, finalParams.fillOpts, finalParams.strokeOpts, finalParams.cornerSmoothing, finalParams.strokeCap, finalParams.opacity);
            });
        }

        if (!isConfirmed || !previewShape) {
            /* キャンセル時にプレビューが残らないようにする / Make sure no preview survives a cancel */
            try { previewManager.rollback(); } catch (e) { }
            previewShape = null;
            return null;
        }

        return true;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを開き、確定した図形を作成して後処理を行う。
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(LABELS.alert.noDocument[lang]);
            return;
        }
        var doc = app.activeDocument;
        var rulerUnit = getRulerUnitInfo();
        var strokeUnitInfo = getStrokeUnitInfo();
        var dialogResult = showInputDialog(rulerUnit.label, rulerUnit.factor, strokeUnitInfo);
        if (!dialogResult) return;

        /* ダイアログ表示中にドキュメントが切り替わった場合に備えて取得し直す
           Re-acquire the active document in case it changed while the dialog was open */
        doc = app.activeDocument;

        finalizeShape(doc);
        if (roughenAnchorsDetail > 0) {
            try {
                if (roughenAnchorsUseMenuFallback && Math.round(roughenAnchorsDetail) === 1) {
                    /* この経路だけはメニューコマンドなのでプレビューには出ない
                       Only this path uses a menu command, so it cannot appear in the preview */
                    app.executeMenuCommand('Add Anchor Points2');
                } else {
                    for (var i = 0; i < doc.selection.length; i++) {
                        applyRoughenEffect(doc.selection[i], roughenAnchorsDetail);
                    }
                }
            } catch (e) { }
        }
        if (applyLiveShape) app.executeMenuCommand('Convert to Shape');
    }

    main();

})();
