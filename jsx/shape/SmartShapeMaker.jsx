#targetengine "MyScriptEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

1つのダイアログから、正円・多角形・星形・スーパー楕円・ルーローの三角形などのカスタム形状を作成します。
リアルタイムプレビューで、辺の数・幅・回転・詳細オプションを調整できます。

詳細は README を参照してください。

### Overview

Creates custom shapes — circle, polygon, star, superellipse, Reuleaux-style — from a single dialog.
A real-time preview lets you adjust the number of sides, the width, the rotation and the advanced options.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartShapeMaker";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-15";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartShapeMaker.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartShapeMaker.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n005a7087f9c3"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 外部JSX読み込み時の警告を抑止 / Suppress the warning raised when an external JSX is loaded */
    app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

    #include "../stroke-table/ColorPicker.jsx"

    (function () {

        // =========================================
        // ユーザー設定 / User settings
        // =========================================

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

        /* ［画面にフィット］したときの余裕（1.0で図形が画面いっぱい） / Breathing room when fitting the view, 1.0 fills the window */
        var FIT_VIEW_MARGIN = 1.05;

        /* 図形生成の内部パラメーター / Internal parameters of the shape generation */
        var SHAPE_GEOMETRY = {
            superEllipsePoints: 8,    /* スーパー楕円のサンプル点数 / sample point count of a superellipse */
            superEllipseHandle: 0.35, /* スーパー楕円のハンドル長比率 / handle length ratio of a superellipse */
            smoothingArmFactor: 0.8   /* 角丸のアーム長係数 / arm length factor of the corner smoothing */
        };

        /* ［辺の数］ラジオの選択肢（0は円）。最後に［それ以外］の手動入力が続く
           Choices of the side-count radios, 0 being a circle; the custom field follows the last one */
        var SIDE_CHOICES = [0, 3, 4, 5, 6, 8];
        var CUSTOM_SIDES_INDEX = SIDE_CHOICES.length;

        /* 円のアンカーポイント数の選択肢 / Choices of the circle anchor count */
        var CIRCLE_ANCHOR_CHOICES = [2, 3, 4, 5, 6];

        /* 回転をONにしたときの円の既定角度（アンカー数ごと） / Default angle of a circle per anchor count, applied when the rotation is turned on */
        var CIRCLE_ROTATION_DEFAULTS = { 2: 90, 3: 180, 4: 45, 5: 180, 6: 30 };

        /* 三角形の向きごとの回転角 / Rotation angle per triangle direction */
        var TRIANGLE_ANGLES = { right: -90, left: 90, down: 60 };

        // =========================================
        // レイアウト / Layout
        // =========================================

        /* ウィンドウ・パネルの余白と間隔 / Window and panel margins and spacing */
        var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
        var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
        var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
        var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
        var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

        /* パネルが多い密なダイアログなので、パネル内と行内は既定より詰める
           This dialog stacks many panels, so panels and rows are tighter than the defaults */
        var DENSE_PANEL_SPACING = 6;   /* パネル内の要素間隔 / spacing inside a panel */
        var ROW_SPACING         = 6;   /* 行内の標準間隔 / default spacing inside a row */
        var TIGHT_ROW_SPACING   = 4;   /* 要素の多い行の間隔 / spacing of a crowded row */
        var WIDE_ROW_SPACING    = 10;  /* ラジオを横に並べる行の間隔 / spacing of a row of radios */

        /* コントロールの寸法 / Control sizes */
        var SWATCH_SIZE        = 16;   /* カラースウォッチの一辺（px） / color swatch size in px */
        var SLIDER_WIDTH       = 200;  /* 標準スライダー幅 / default slider width */
        var SHORT_SLIDER_WIDTH = 150;  /* 短いスライダー幅 / short slider width */
        var SIDES_SLIDER_WIDTH = 100;  /* ［辺の数］スライダー幅（入力欄と同じ行に収める） / slider width of the side count, kept on the field's row */
        var ZOOM_SLIDER_WIDTH  = 300;  /* ［画面ズーム］スライダー幅 / slider width of the view zoom */

        /* ダイアログの表示位置と不透明度 / Dialog position and opacity */
        var DIALOG_OFFSET_X = 300;    /* 初回表示位置のオフセットX / dialog offset X */
        var DIALOG_OFFSET_Y = 0;      /* 初回表示位置のオフセットY / dialog offset Y */
        var DIALOG_OPACITY  = 0.98;   /* ダイアログの不透明度 / dialog opacity */

        // =========================================
        // ローカライズ / Localization
        // =========================================

        /**
         * 実行環境のロケールからUI言語を判定する。
         * @returns {string} "ja" または "en"
         */
        function getUiLang() {
            return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
        }
        var uiLang = getUiLang();

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
                roughenAnchors: { ja: "ラフ効果で追加：", en: "Add Anchors (Roughen):" },
                fitView: { ja: "画面にフィット", en: "Fit View" }
            },
            tooltip: {
                fitView: {
                    ja: "［幅］を変えたときに、作成する図形が収まるよう表示倍率を合わせます。",
                    en: "Refits the view to the shape being created whenever the width changes."
                }
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
         * Illustratorの単位コードから、単位ラベルとpt換算係数を取得する。
         * @param {number} unitCode - 単位コード（rulerType / strokeUnits 共通）
         * @returns {{label: string, factorToPt: number}} 単位ラベルとpt換算係数
         */
        function getUnitInfo(unitCode) {
            switch (unitCode) {
                case 0: return { label: "inch", factorToPt: 72 };
                case 1: return { label: "mm", factorToPt: 72 / 25.4 };
                case 3: return { label: "pica", factorToPt: 12 };
                case 4: return { label: "cm", factorToPt: 72 / 2.54 };
                case 5: return { label: "Q", factorToPt: (72 / 25.4) * 0.25 };
                case 6: return { label: "px", factorToPt: 1 };
                default: return { label: "pt", factorToPt: 1 }; /* case 2 */
            }
        }

        /**
         * 環境設定から単位情報を取得する（取得できないときはptとして扱う）。
         * @param {string} preferenceKey - 単位の環境設定キー（"rulerType" または "strokeUnits"）
         * @returns {{label: string, factorToPt: number}} 単位ラベルとpt換算係数
         */
        function getUnitInfoFromPreference(preferenceKey) {
            var unitCode = 2; /* ptにフォールバック / fall back to pt */
            try {
                unitCode = app.preferences.getIntegerPreference(preferenceKey);
            } catch (e) { }
            return getUnitInfo(unitCode);
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
         * 数値を有効範囲に収める。空欄や数値でない入力は既定値に戻す。
         * @param {string|number} value - 入力値
         * @param {Array<number>} range - [下限, 上限]
         * @param {number} fallbackValue - 数値として読めないときの既定値
         * @param {boolean} [asInteger] - 整数に丸めるかどうか
         * @returns {number} 範囲内に収めた数値
         */
        function clampNumber(value, range, fallbackValue, asInteger) {
            /* 入力途中の空欄はNumber()が0になってしまうので既定値に戻す
               An empty field would become 0 through Number(), so fall back to the default */
            if (typeof value === "string" && !/\S/.test(value)) value = fallbackValue;
            value = Number(value);
            if (isNaN(value)) value = fallbackValue;
            if (asInteger) value = Math.round(value);
            if (value < range[0]) value = range[0];
            if (value > range[1]) value = range[1];
            return value;
        }

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
                    alert(LABELS.alert.previewError[uiLang] + e);
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
                    try { finalAction(); } catch (e) { alert(LABELS.alert.finalError[uiLang] + e); }
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
         * 見出し付きパネルを追加し、共通レイアウトを適用する。
         * @param {object} parent - 追加先のコンテナ
         * @param {string} labelText - パネルの見出し
         * @param {number} [spacing] - 要素間隔（省略時はDENSE_PANEL_SPACING）
         * @returns {Panel} 追加したパネル
         */
        function addPanel(parent, labelText, spacing) {
            var panel = parent.add("panel", undefined, labelText);
            setupPanel(panel, (typeof spacing === "number") ? spacing : DENSE_PANEL_SPACING);
            return panel;
        }

        /**
         * 左揃え・上下中央の行グループを追加する。
         * @param {object} parent - 追加先のコンテナ
         * @param {number} [spacing] - 要素間隔（省略時はROW_SPACING）
         * @returns {Group} 追加したグループ
         */
        function addControlRow(parent, spacing) {
            var row = parent.add("group");
            row.orientation = "row";
            row.alignChildren = ["left", "center"];
            row.spacing = (typeof spacing === "number") ? spacing : ROW_SPACING;
            return row;
        }

        /**
         * 上下矢印キーで増減できる数値入力欄を追加する。
         * @param {object} parent - 追加先のコンテナ
         * @param {string|number} value - 初期値
         * @param {number} charCount - 入力欄の文字幅
         * @returns {EditText} 追加した入力欄
         */
        function addNumberField(parent, value, charCount) {
            var editText = parent.add("edittext", undefined, String(value));
            editText.characters = charCount;
            changeValueByArrowKey(editText);
            return editText;
        }

        /**
         * 幅を指定したスライダーを追加する。
         * @param {object} parent - 追加先のコンテナ
         * @param {number} value - 初期値
         * @param {Array<number>} range - [下限, 上限]
         * @param {number} width - スライダー幅（px）
         * @returns {Slider} 追加したスライダー
         */
        function addSlider(parent, value, range, width) {
            var slider = parent.add("slider", undefined, value, range[0], range[1]);
            slider.preferredSize.width = width;
            return slider;
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

                /* 矢印キーは確定した編集として扱う。onChangeがあれば丸めまで、なければプレビューだけ反映する
                   An arrow key is a committed edit: onChange also rounds the value, onChanging only redraws */
                var editHandler = (typeof editText.onChange === "function") ? editText.onChange : editText.onChanging;
                if (typeof editHandler === "function") {
                    try { editHandler(); } catch (e) { }
                }
            });
        }

        /**
         * 入力途中で書き戻すと編集できなくなる文字列かどうかを判定する。
         * 空欄・符号だけ・小数点で終わる値は、まだ確定していないものとして扱う。
         * @param {string} text - 入力欄の文字列
         * @returns {boolean} 入力途中ならtrue
         */
        function isPartialNumberInput(text) {
            return /^\s*[+-]?\s*$/.test(text) || /\.\s*$/.test(text);
        }

        /**
         * 選択中のラジオボタンのインデックスを返す。
         * @param {Array} radios - ラジオボタンの配列
         * @returns {number} 選択中のインデックス（未選択なら-1）
         */
        function getSelectedRadioIndex(radios) {
            for (var i = 0; i < radios.length; i++) {
                if (radios[i].value) return i;
            }
            return -1;
        }

        /**
         * 指定したインデックスのラジオボタンだけをONにする。
         * ［それ以外］は別グループにあるため排他が効かず、全件を明示的に解除する必要がある。
         * @param {Array} radios - ラジオボタンの配列
         * @param {number} selectedIndex - ONにするインデックス
         * @returns {void}
         */
        function selectRadio(radios, selectedIndex) {
            for (var i = 0; i < radios.length; i++) {
                radios[i].value = (i === selectedIndex);
            }
        }

        // =========================================
        // パス生成 / Path builders
        // =========================================

        /**
         * 図形を作る前にアクティブレイヤーの編集を許可し、ドキュメントウィンドウの中心を得る。
         * @param {Document} doc - 対象ドキュメント
         * @returns {{layer: Layer, centerX: number, centerY: number}} レイヤーと中心座標
         */
        function prepareActiveLayer(doc) {
            var layer = doc.activeLayer;
            layer.locked = false;
            layer.visible = true;
            var viewCenter = doc.activeView.centerPoint;
            return { layer: layer, centerX: viewCenter[0], centerY: viewCenter[1] };
        }

        /**
         * 自前で組んだパスの見た目をcreateShapeの既定に合わせる。
         * @param {Document} doc - 対象ドキュメント
         * @param {PathItem} pathItem - 対象のパス
         * @returns {PathItem} 見た目を適用したパス
         */
        function applyDefaultAppearance(doc, pathItem) {
            pathItem.filled = true;
            pathItem.fillColor = doc.defaultFillColor;
            pathItem.stroked = false;
            return pathItem;
        }

        /**
         * 座標の配列から閉じたパスを作る。
         * @param {Layer} layer - 追加先のレイヤー
         * @param {Array} anchors - アンカー座標の配列
         * @returns {PathItem} 作成したパス
         */
        function createClosedPath(layer, anchors) {
            var pathItem = layer.pathItems.add();
            pathItem.setEntirePath(anchors);
            pathItem.closed = true;
            return pathItem;
        }

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
            pointCount = (typeof pointCount === 'number' && pointCount >= SHAPE_GEOMETRY.superEllipsePoints)
                ? Math.round(pointCount)
                : SHAPE_GEOMETRY.superEllipsePoints;

            var placement = prepareActiveLayer(doc);
            var radius = sizePt / 2;

            var anchors = [];
            for (var i = 0; i < pointCount; i++) {
                var theta = (Math.PI * 2 * i) / pointCount;
                var cosTheta = Math.cos(theta);
                var sinTheta = Math.sin(theta);
                var x = Math.pow(Math.abs(cosTheta), 2 / exponent) * radius * signOf(cosTheta);
                var y = Math.pow(Math.abs(sinTheta), 2 / exponent) * radius * signOf(sinTheta);
                anchors.push([x + placement.centerX, y + placement.centerY]);
            }

            var pathItem = createClosedPath(placement.layer, anchors);

            /* 接線方向（次点−前点）にスムーズハンドルを置く / Place smooth handles along the tangents (next - prev) */
            var pathPoints = pathItem.pathPoints;
            var anchorCount = pathPoints.length;
            for (var k = 0; anchorCount >= 4 && k < anchorCount; k++) {
                var prevAnchor = anchors[(k - 1 + anchorCount) % anchorCount];
                var currentAnchor = anchors[k];
                var nextAnchor = anchors[(k + 1) % anchorCount];

                var tangentX = nextAnchor[0] - prevAnchor[0];
                var tangentY = nextAnchor[1] - prevAnchor[1];
                var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY);
                if (tangentLength === 0) continue;
                tangentX /= tangentLength;
                tangentY /= tangentLength;

                /* 前後のセグメント長のうち短いほうに合わせる / Follow the shorter of the neighbouring segments */
                var prevDeltaX = currentAnchor[0] - prevAnchor[0];
                var prevDeltaY = currentAnchor[1] - prevAnchor[1];
                var nextDeltaX = nextAnchor[0] - currentAnchor[0];
                var nextDeltaY = nextAnchor[1] - currentAnchor[1];
                var prevLength = Math.sqrt(prevDeltaX * prevDeltaX + prevDeltaY * prevDeltaY);
                var nextLength = Math.sqrt(nextDeltaX * nextDeltaX + nextDeltaY * nextDeltaY);
                var handleLength = Math.min(prevLength, nextLength) * SHAPE_GEOMETRY.superEllipseHandle;

                pathPoints[k].anchor = currentAnchor;
                pathPoints[k].leftDirection = [currentAnchor[0] - tangentX * handleLength, currentAnchor[1] - tangentY * handleLength];
                pathPoints[k].rightDirection = [currentAnchor[0] + tangentX * handleLength, currentAnchor[1] + tangentY * handleLength];
                pathPoints[k].pointType = PointType.SMOOTH;
            }

            return applyDefaultAppearance(doc, pathItem);
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
            var placement = prepareActiveLayer(doc);
            var radius = sizePt / 2;
            anchorCount = (typeof anchorCount === 'number') ? Math.round(anchorCount) : SHAPE_DEFAULTS.circleAnchors;
            if (anchorCount < 2) anchorCount = 2;

            var handleLength = radius * (4 / 3) * Math.tan(Math.PI / (2 * anchorCount));

            /* 頂点が真上に来るよう-90°から並べる / Start at -90 degrees so one anchor sits on top */
            var angles = [];
            var anchors = [];
            for (var i = 0; i < anchorCount; i++) {
                var angle = (-Math.PI / 2) + (2 * Math.PI * i) / anchorCount;
                angles.push(angle);
                anchors.push([placement.centerX + radius * Math.cos(angle), placement.centerY + radius * Math.sin(angle)]);
            }

            var pathItem = createClosedPath(placement.layer, anchors);

            /* 接線は半径に直交する [-sin, cos] / The tangent is perpendicular to the radius */
            var pathPoints = pathItem.pathPoints;
            for (var j = 0; j < anchorCount; j++) {
                var tangentX = -Math.sin(angles[j]);
                var tangentY = Math.cos(angles[j]);

                pathPoints[j].anchor = anchors[j];
                pathPoints[j].leftDirection = [anchors[j][0] - tangentX * handleLength, anchors[j][1] - tangentY * handleLength];
                pathPoints[j].rightDirection = [anchors[j][0] + tangentX * handleLength, anchors[j][1] + tangentY * handleLength];
                pathPoints[j].pointType = PointType.SMOOTH;
            }

            return applyDefaultAppearance(doc, pathItem);
        }

        /**
         * 奇数辺の正多角形の各辺を円弧に置き換えてルーロー図形にする。
         * 参考スクリプト reuleaux_polygon.jsx のロジックを移植。
         * @param {PathItem} pathItem - 対象のパス（奇数個のアンカーを持つ多角形）
         * @param {number} amount - 度合い（1.0が標準、0.0〜2.0）
         * @returns {PathItem} 変換後のパス
         */
        function applyReuleauxToPolygon(pathItem, amount) {
            if (!pathItem || pathItem.typename !== "PathItem") return pathItem;
            if (!pathItem.pathPoints || pathItem.pathPoints.length < 3) return pathItem;
            var pathPoints = pathItem.pathPoints;
            var pointCount = pathPoints.length;
            if (pointCount % 2 === 0) return pathItem; /* 奇数辺のみ / odd side counts only */

            amount = clampNumber(amount, [0, 2], 1);

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

                var radiusToStart = Math.sqrt(vectorToStart[0] * vectorToStart[0] + vectorToStart[1] * vectorToStart[1]);
                var radiusToEnd = Math.sqrt(vectorToEnd[0] * vectorToEnd[0] + vectorToEnd[1] * vectorToEnd[1]);
                if (radiusToStart === 0 || radiusToEnd === 0) continue;
                var arcRadius = (radiusToStart + radiusToEnd) / 2;

                var dotProduct = vectorToStart[0] * vectorToEnd[0] + vectorToStart[1] * vectorToEnd[1];
                var cosTheta = clampNumber(dotProduct / (radiusToStart * radiusToEnd), [-1, 1], 0);
                var deltaTheta = Math.acos(cosTheta);

                /* 円弧をベジェで近似したハンドル長に度合いを掛ける / Bezier approximation of the arc, scaled by the amount */
                var handleLength = arcRadius * (4 / 3) * Math.tan(deltaTheta / 4) * amount;

                /* 中心から見た回り方向に合わせて接線の向きを決める / Pick the tangent side from the winding around the arc center */
                var crossZ = vectorToStart[0] * vectorToEnd[1] - vectorToStart[1] * vectorToEnd[0];
                var tangentSign = (crossZ > 0) ? 1 : -1;
                var tangentStart = [-vectorToStart[1] * tangentSign, vectorToStart[0] * tangentSign];
                var tangentEnd = [vectorToEnd[1] * tangentSign, -vectorToEnd[0] * tangentSign];

                pathPoints[startIndex].rightDirection = [
                    arcStart[0] + handleLength * tangentStart[0] / radiusToStart,
                    arcStart[1] + handleLength * tangentStart[1] / radiusToStart
                ];
                pathPoints[endIndex].leftDirection = [
                    arcEnd[0] + handleLength * tangentEnd[0] / radiusToEnd,
                    arcEnd[1] + handleLength * tangentEnd[1] / radiusToEnd
                ];

                pathPoints[startIndex].pointType = PointType.CORNER;
                pathPoints[endIndex].pointType = PointType.CORNER;
            }

            pathItem.closed = true;
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

            var armLength = radius * (1 + SHAPE_GEOMETRY.smoothingArmFactor * smoothing);
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
         * ドキュメントのカラーモードに合わせた黒を作る。
         * @param {Document} doc - 対象ドキュメント
         * @returns {CMYKColor|RGBColor} 黒のカラーオブジェクト
         */
        function createBlackColor(doc) {
            if (doc && doc.documentColorSpace === DocumentColorSpace.CMYK) {
                var cmykColor = new CMYKColor();
                cmykColor.cyan = 0;
                cmykColor.magenta = 0;
                cmykColor.yellow = 0;
                cmykColor.black = 100;
                return cmykColor;
            }
            var rgbColor = new RGBColor();
            rgbColor.red = 0;
            rgbColor.green = 0;
            rgbColor.blue = 0;
            return rgbColor;
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

                /* 開いたパスに塗りは合わないので線だけにする。線がOFFなら黒の細線で見えるようにする
                   Open paths get a stroke, not a fill; without a stroke option they fall back to a thin black line */
                segmentPath.filled = false;
                segmentPath.stroked = true;
                var hasStroke = !!(strokeOpts && strokeOpts.enabled);
                try {
                    segmentPath.strokeColor = hasStroke ? strokeOpts.color : createBlackColor(doc);
                    segmentPath.strokeWidth = hasStroke ? strokeOpts.widthPt : SHAPE_DEFAULTS.segmentStrokeWidth;
                    if (strokeCap) segmentPath.strokeCap = strokeCap;
                } catch (e) { }
            }

            /* 元のパスを削除 / Remove the original path */
            pathItem.remove();

            return segmentGroup;
        }

        /**
         * 図形生成のパラメーター一式。ダイアログのgetCurrentParams()が組み立てる。
         * @typedef {object} ShapeParams
         * @property {number} size - 幅（pt）
         * @property {number} sides - 辺の数（0は円）
         * @property {boolean} isStar - スターにするか
         * @property {number} innerRatio - スターの第2半径（%）
         * @property {boolean} rotateEnabled - 回転を適用するか
         * @property {number} angle - 回転角（度）
         * @property {boolean} splitAtAnchors - アンカーポイントで分割するか
         * @property {StrokeCap} strokeCap - 分割時の線端の種類
         * @property {boolean} useSuperEllipse - スーパー楕円にするか
         * @property {number} superExponent - スーパー楕円の指数
         * @property {number} circleAnchorCount - 円のアンカーポイント数
         * @property {boolean} useReuleaux - ルーロー図形にするか
         * @property {number} reuleauxAmount - ルーローの度合い（1.0が標準）
         * @property {object} fillOpts - 塗りの設定 {enabled, color}
         * @property {object} strokeOpts - 線の設定 {enabled, color, widthPt}
         * @property {object} cornerSmoothing - 角丸の設定 {radius, smoothing}（不要ならnull）
         * @property {number} opacity - 不透明度（%）
         * @property {number} roughenDetail - ラフ効果の詳細（0で無効）
         */

        /**
         * 円のもとになるパスを作成する。
         * @param {Document} doc - 対象ドキュメント
         * @param {ShapeParams} params - 図形生成のパラメーター
         * @param {Array} viewCenter - ドキュメントウィンドウの中心座標
         * @returns {PathItem} 作成したパス
         */
        function createCircleBasePath(doc, params, viewCenter) {
            if (params.useSuperEllipse) {
                return createSuperellipsePath(doc, params.size, params.superExponent);
            }
            /* 既定の4アンカーはIllustratorの楕円、それ以外は独自のスムーズパス
               Four anchors use Illustrator's ellipse; other counts build a custom smooth path */
            var anchorCount = (typeof params.circleAnchorCount === 'number') ? Math.round(params.circleAnchorCount) : SHAPE_DEFAULTS.circleAnchors;
            if (anchorCount < 2) anchorCount = 2;
            if (anchorCount !== 4) return createCirclePathWithNAnchors(doc, params.size, anchorCount);

            var radius = params.size / 2;
            return doc.activeLayer.pathItems.ellipse(viewCenter[1] + radius, viewCenter[0] - radius, params.size, params.size);
        }

        /**
         * 正方形のもとになるパスを作成する（角丸の指定に応じて作り方を変える）。
         * @param {Document} doc - 対象ドキュメント
         * @param {ShapeParams} params - 図形生成のパラメーター
         * @param {Array} viewCenter - ドキュメントウィンドウの中心座標
         * @returns {PathItem} 作成したパス
         */
        function createSquareBasePath(doc, params, viewCenter) {
            var cornerSmoothing = params.cornerSmoothing;
            var hasCornerRadius = !!(cornerSmoothing && cornerSmoothing.radius > 0);

            if (hasCornerRadius && cornerSmoothing.smoothing > 0) {
                /* スムージングありの角丸は独自のベジェパス / Corner smoothing above zero builds a custom bezier path */
                return buildSmoothedRect(doc, viewCenter[0] - params.size / 2, viewCenter[1] + params.size / 2,
                    params.size, params.size, cornerSmoothing.radius, cornerSmoothing.smoothing / 100);
            }

            /* 正方形は1辺の長さを幅として扱うので外接円の半径に換算する（既定の45°回転が前提）
               A square is sized by its edge, so convert to the circumscribed radius; assumes the default 45 degree rotation */
            var squarePath = doc.pathItems.polygon(viewCenter[0], viewCenter[1], params.size / Math.sqrt(2), 4);

            if (hasCornerRadius) {
                /* スムージング0の角丸は通常の正方形＋［角を丸くする］効果
                   A zero smoothing value uses a plain square plus the Round Corners live effect */
                try {
                    squarePath.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + cornerSmoothing.radius + ' "/></LiveEffect>');
                } catch (e) { }
            }
            return squarePath;
        }

        /**
         * 辺の数と各オプションから、変形前のもとになるパスを作成する。
         * @param {Document} doc - 対象ドキュメント
         * @param {ShapeParams} params - 図形生成のパラメーター
         * @param {Array} viewCenter - ドキュメントウィンドウの中心座標
         * @returns {PathItem} 作成したパス
         */
        function createBasePath(doc, params, viewCenter) {
            var radius = params.size / 2;
            if (params.sides === 0) return createCircleBasePath(doc, params, viewCenter);
            if (params.isStar) {
                return doc.pathItems.star(viewCenter[0], viewCenter[1], radius, radius * (params.innerRatio / 100), params.sides);
            }
            if (params.sides === 4) return createSquareBasePath(doc, params, viewCenter);

            /* 正方形以外はsizeを外接円の直径として扱う（バウンディングボックスの幅とは一致しない）
               Other polygons treat the size as the circumscribed diameter, which is not the bounding box width */
            return doc.pathItems.polygon(viewCenter[0], viewCenter[1], radius, params.sides);
        }

        /**
         * 塗りと線の設定をオブジェクトに適用する。
         * @param {PathItem} shape - 対象の図形
         * @param {object} fillOpts - 塗りの設定 {enabled, color}
         * @param {object} strokeOpts - 線の設定 {enabled, color, widthPt}
         * @returns {void}
         */
        function applyFillAndStroke(shape, fillOpts, strokeOpts) {
            shape.filled = !!(fillOpts && fillOpts.enabled);
            if (shape.filled) shape.fillColor = fillOpts.color;

            shape.stroked = !!(strokeOpts && strokeOpts.enabled);
            if (shape.stroked) {
                shape.strokeColor = strokeOpts.color;
                shape.strokeWidth = strokeOpts.widthPt;
            }
        }

        /**
         * 指定したパラメーターから図形を作成し、選択状態にする。
         * @param {Document} doc - 対象ドキュメント
         * @param {ShapeParams} params - 図形生成のパラメーター
         * @returns {PathItem|GroupItem} 作成した図形
         */
        function createShape(doc, params) {
            var placement = prepareActiveLayer(doc);
            var viewCenter = [placement.centerX, placement.centerY];

            var shape = createBasePath(doc, params, viewCenter);
            applyFillAndStroke(shape, params.fillOpts, params.strokeOpts);

            /* ドキュメントウィンドウの中央にそろえる / Center the shape in the document window */
            var bounds = shape.geometricBounds;
            shape.translate(viewCenter[0] - (bounds[0] + bounds[2]) / 2, viewCenter[1] - (bounds[1] + bounds[3]) / 2);

            /* 奇数辺の多角形をルーロー（定幅図形）に変換 / Convert odd-sided polygons into constant-width shapes */
            if (params.useReuleaux && !params.isStar && params.sides > 0 && (params.sides % 2 === 1)) {
                shape = applyReuleauxToPolygon(shape, params.reuleauxAmount);
            }

            if (params.rotateEnabled && !isNaN(params.angle)) {
                shape.rotate(params.angle, true, true, true, true, Transformation.CENTER);
            }
            if (params.splitAtAnchors) {
                shape = splitPathAtAnchors(doc, shape, params.strokeOpts, params.strokeCap);
            }
            if (typeof params.opacity === "number" && params.opacity < 100) {
                shape.opacity = params.opacity;
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

        /**
         * 指定したオブジェクトが収まるようにドキュメントウィンドウの表示倍率を合わせる。
         * ColorPaletteFromImage.jsx の fitViewToItems() をもとにしている。
         * @param {Document} doc - 対象ドキュメント
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function fitViewToItem(doc, item) {
            if (!item) return;

            var bounds = item.geometricBounds;
            var itemWidth = bounds[2] - bounds[0];
            var itemHeight = bounds[1] - bounds[3];

            var activeView = doc.activeView;
            activeView.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
            if (itemWidth <= 0 || itemHeight <= 0) return;

            /* 中心をそろえたあとの表示範囲を基準に倍率を求める / Scale from the view bounds after the center has moved */
            var viewBounds = activeView.bounds;
            var scale = Math.min(
                (viewBounds[2] - viewBounds[0]) / itemWidth,
                (viewBounds[1] - viewBounds[3]) / itemHeight
            ) / FIT_VIEW_MARGIN;
            activeView.zoom = activeView.zoom * scale;
        }

        // =========================================
        // カラーの変換 / Color conversion
        // =========================================

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
         * @returns {CMYKColor|RGBColor} Illustratorのカラーオブジェクト
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
            }
            var rgbValues = ColorPicker.hexToRGB(pickerString);
            var rgbColor = new RGBColor();
            rgbColor.red = rgbValues.r;
            rgbColor.green = rgbValues.g;
            rgbColor.blue = rgbValues.b;
            return rgbColor;
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
        function addColorSwatch(parent, aiColor) {
            var swatchGroup = parent.add("group");
            swatchGroup.preferredSize = [SWATCH_SIZE, SWATCH_SIZE];
            swatchGroup.minimumSize = [SWATCH_SIZE, SWATCH_SIZE];
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

        /**
         * スウォッチを描き直す（色を変えたあとに呼ぶ）。
         * @param {Group} swatchGroup - 対象のスウォッチ
         * @returns {void}
         */
        function redrawSwatch(swatchGroup) {
            try {
                swatchGroup.hide();
                swatchGroup.show();
            } catch (e) { }
        }

        // =========================================
        // ダイアログ位置の記憶 / Stored dialog position
        // =========================================

        /**
         * セッション状態に保存されたダイアログ位置を取得する。
         * @returns {Array<number>} [x, y] の配列、保存がなければnull
         */
        function getSavedDialogLocation() {
            var savedLocation = getSessionState().dialogLocation;
            if (!savedLocation || savedLocation.length !== 2) return null;
            var x = Number(savedLocation[0]);
            var y = Number(savedLocation[1]);
            return (isNaN(x) || isNaN(y)) ? null : [x, y];
        }

        /**
         * ダイアログの表示位置をセッション状態に保存する。
         * @param {Window} targetDialog - 対象のダイアログ
         * @returns {void}
         */
        function saveDialogLocation(targetDialog) {
            if (!targetDialog || !targetDialog.location) return;
            getSessionState().dialogLocation = [Number(targetDialog.location[0]), Number(targetDialog.location[1])];
        }

        /**
         * 入力欄にフォーカスがあるかどうかを判定する。
         * @param {object} target - イベントの発生元
         * @returns {boolean} 入力欄ならtrue
         */
        function isTextInputTarget(target) {
            return !!(target && target.type === "edittext");
        }

        // =========================================
        // ダイアログ / Dialog
        // =========================================

        /**
         * 設定ダイアログを表示し、OKが押されたら確定用の状態を整える。
         * @param {object} rulerUnitInfo - 定規単位の情報 {label, factorToPt}
         * @param {object} strokeUnitInfo - 線の単位情報 {label, factorToPt}
         * @returns {boolean} OKで確定できたときtrue、キャンセルや失敗時はnull
         */
        function showInputDialog(rulerUnitInfo, strokeUnitInfo) {
            var dialog = new Window("dialog", LABELS.dialog.title[uiLang] + " " + SCRIPT_VERSION);
            var previewManager = new PreviewManager();
            var doc = app.activeDocument;

            /* 確定に関わるダイアログの状態 / Dialog state referenced when finalizing */
            var isConfirmed = false;
            var finalParams = null;
            var documentView = null;
            var initialZoom = null;

            /* 辺の数 / Side count */
            var sidesPanel, sideRadios = [], customSidesInput, customSidesSlider;
            /* 回転と三角形 / Rotation and triangle */
            var rotatePanel, rotateCheck, rotateInput, rotateUnitLabel;
            var trianglePanel, triangleRightRadio, triangleLeftRadio, triangleDownRadio;
            /* 塗りと線 / Fill and stroke */
            var fillCheck, fillSwatch, strokeCheck, strokeSwatch, strokeWidthInput, strokeWidthUnitLabel;
            var opacityInput, opacitySlider;
            /* 幅と表示 / Width and view */
            var sizeInput, fitViewCheck;
            /* スター / Star */
            var starPanel, starCheck, pentagramCheck;
            var innerRadiusLabel, innerRatioInput, innerPercentLabel, innerRatioSlider;
            /* 円 / Circle */
            var circlePanel, superEllipseCheck, superExponentInput, superExponentSlider;
            var circleAnchorPanel, circleAnchorRadios = [];
            /* 角丸 / Corner smoothing */
            var cornerSmoothingPanel, cornerRadiusCheck, cornerRadiusInput, smoothingValueLabel, smoothingSlider;
            /* アンカーポイントの操作 / Anchor point operations */
            var roughenAnchorsCheck, roughenAnchorsInput, splitAtAnchorsCheck;
            var strokeCapLabel, capButtRadio, capRoundRadio, capProjectingRadio;
            /* オプション / Options */
            var liveShapeCheck, reuleauxCheck, reuleauxAmountInput, reuleauxAmountSlider;

            // -----------------------------------------
            // 現在の入力の読み取り / Reading the current input
            // -----------------------------------------

            /**
             * ラジオボタンまたは手動入力から辺の数を取得する。
             * @returns {number} 辺の数（0は円）
             */
            function getCurrentSides() {
                var selectedIndex = getSelectedRadioIndex(sideRadios);
                if (selectedIndex < 0) return 4;
                return (selectedIndex === CUSTOM_SIDES_INDEX) ? parseInt(customSidesInput.text, 10) : SIDE_CHOICES[selectedIndex];
            }

            /**
             * 円のアンカーポイント数をラジオボタンから取得する。
             * @returns {number} アンカーポイント数
             */
            function getCircleAnchorCount() {
                var selectedIndex = getSelectedRadioIndex(circleAnchorRadios);
                return (selectedIndex < 0) ? SHAPE_DEFAULTS.circleAnchors : CIRCLE_ANCHOR_CHOICES[selectedIndex];
            }

            /**
             * スーパー楕円が実際に効く状態かどうかを判定する。
             * @param {number} sidesValue - 現在の辺の数
             * @returns {boolean} 円かつスーパー楕円がONならtrue
             */
            function isSuperEllipseActive(sidesValue) {
                return !!(superEllipseCheck.value && sidesValue === 0);
            }

            /**
             * 選択されている線端の種類を取得する。
             * @returns {StrokeCap} 線端の種類
             */
            function getSelectedStrokeCap() {
                if (capRoundRadio.value) return StrokeCap.ROUNDENDCAP;
                if (capProjectingRadio.value) return StrokeCap.PROJECTINGENDCAP;
                return StrokeCap.BUTTENDCAP;
            }

            /**
             * スーパー楕円の指数を有効範囲に収める（小数第1位まで）。
             * @param {string|number} value - 入力値
             * @returns {number} 丸めた指数
             */
            function clampSuperExponent(value) {
                return Math.round(clampNumber(value, SHAPE_RANGES.superExponent, SHAPE_DEFAULTS.superExponent) * 10) / 10;
            }

            /**
             * ルーローの度合いを有効範囲に収める。
             * @param {string|number} value - 入力値
             * @returns {number} 整数に丸めた度合い（%）
             */
            function clampReuleauxAmount(value) {
                return clampNumber(value, SHAPE_RANGES.reuleauxAmount, SHAPE_DEFAULTS.reuleauxAmount, true);
            }

            /**
             * 不透明度を有効範囲に収める（0は有効な値なので既定値に丸めない）。
             * @param {string|number} value - 入力値
             * @returns {number} 整数に丸めた不透明度（%）
             */
            function clampOpacity(value) {
                return clampNumber(value, SHAPE_RANGES.opacity, SHAPE_DEFAULTS.opacity, true);
            }

            // -----------------------------------------
            // パネルの組み立て / Panel construction
            // -----------------------------------------

            /**
             * ［辺の数］パネルを組み立てる。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildSidesPanel(parent) {
                sidesPanel = addPanel(parent, LABELS.panel.sides[uiLang]);

                sideRadios[0] = sidesPanel.add("radiobutton", undefined, LABELS.radio.circleWithZero[uiLang]);
                for (var i = 1; i < SIDE_CHOICES.length; i++) {
                    sideRadios[i] = sidesPanel.add("radiobutton", undefined, String(SIDE_CHOICES[i]));
                }

                /* ［それ以外］はラベルを持たず、ラジオ・入力欄・スライダーを1行に並べる
                   The custom side count has no label; its radio, field and slider share one row */
                var customSidesRow = addControlRow(sidesPanel);
                sideRadios[CUSTOM_SIDES_INDEX] = customSidesRow.add("radiobutton", undefined, "");
                customSidesInput = addNumberField(customSidesRow, SHAPE_DEFAULTS.customSides, 3);
                customSidesSlider = addSlider(customSidesRow, SHAPE_DEFAULTS.customSides, SHAPE_RANGES.customSides, SIDES_SLIDER_WIDTH);

                selectRadio(sideRadios, 2); /* 既定は4辺 / Four sides by default */
                setCustomSidesEnabled(false);
            }

            /**
             * ［回転］パネルと、その中の［三角形］パネルを組み立てる。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildRotatePanel(parent) {
                rotatePanel = addPanel(parent, LABELS.panel.rotation[uiLang]);

                var rotateRow = addControlRow(rotatePanel);
                rotateCheck = rotateRow.add("checkbox", undefined, "");
                rotateInput = addNumberField(rotateRow, SHAPE_DEFAULTS.rotation, 4);
                rotateUnitLabel = rotateRow.add("statictext", undefined, "°");
                /* 手動入力はチェック時だけ有効 / Manual entry is enabled only while checked */
                rotateInput.enabled = rotateCheck.value;
                rotateUnitLabel.enabled = rotateCheck.value;

                /* 三角形パネルは回転パネルの中に置く / The triangle panel lives inside the rotation panel */
                trianglePanel = addPanel(rotatePanel, LABELS.panel.triangle[uiLang]);
                var triangleRow = addControlRow(trianglePanel, WIDE_ROW_SPACING);
                triangleRightRadio = triangleRow.add("radiobutton", undefined, LABELS.radio.triangleRight[uiLang]);
                triangleLeftRadio = triangleRow.add("radiobutton", undefined, LABELS.radio.triangleLeft[uiLang]);
                triangleDownRadio = triangleRow.add("radiobutton", undefined, LABELS.radio.triangleDown[uiLang]);
                triangleRightRadio.value = true;
            }

            /**
             * ［塗りと線］パネルを組み立てる。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildFillStrokePanel(parent) {
                var fillStrokePanel = addPanel(parent, LABELS.panel.fillAndStroke[uiLang]);
                var checkboxWidth = (uiLang === 'ja') ? 46 : 66;

                var fillRow = addControlRow(fillStrokePanel);
                fillCheck = fillRow.add("checkbox", undefined, LABELS.checkbox.fill[uiLang]);
                fillCheck.preferredSize.width = checkboxWidth;
                fillCheck.value = true;
                fillSwatch = addColorSwatch(fillRow, doc.defaultFillColor);

                var strokeRow = addControlRow(fillStrokePanel);
                strokeCheck = strokeRow.add("checkbox", undefined, LABELS.checkbox.stroke[uiLang]);
                strokeCheck.preferredSize.width = checkboxWidth;
                strokeCheck.value = false;
                strokeSwatch = addColorSwatch(strokeRow, doc.defaultStrokeColor);

                /* 線幅は［線］と同じ行に置く（ラベルなし） / The stroke width sits on the stroke row, without a label */
                strokeWidthInput = addNumberField(strokeRow, SHAPE_DEFAULTS.strokeWidth, 4);
                strokeWidthUnitLabel = strokeRow.add("statictext", undefined, strokeUnitInfo.label);

                var opacityRow = addControlRow(fillStrokePanel);
                opacityRow.add("statictext", undefined, LABELS.label.opacity[uiLang]);
                opacityInput = addNumberField(opacityRow, SHAPE_DEFAULTS.opacity, 4);
                opacityRow.add("statictext", undefined, "%");
                opacitySlider = addSlider(fillStrokePanel, SHAPE_DEFAULTS.opacity, SHAPE_RANGES.opacity, SHORT_SLIDER_WIDTH);
            }

            /**
             * ［幅］パネルを組み立てる。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildWidthPanel(parent) {
                var widthPanel = addPanel(parent, LABELS.panel.width[uiLang]);
                var widthRow = addControlRow(widthPanel);
                sizeInput = addNumberField(widthRow, SHAPE_DEFAULTS.size, 5);
                widthRow.add("statictext", undefined, rulerUnitInfo.label);
            }

            /**
             * ［画面にフィット］の行を組み立てる（幅パネルの下に置く）。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildFitViewRow(parent) {
                var fitViewRow = addControlRow(parent);
                fitViewRow.margins = [20, 0, 0, 0];
                fitViewCheck = fitViewRow.add("checkbox", undefined, LABELS.checkbox.fitView[uiLang]);
                fitViewCheck.value = false;
                fitViewCheck.helpTip = LABELS.tooltip.fitView[uiLang];
            }

            /**
             * ［スター］パネルを組み立てる。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildStarPanel(parent) {
                starPanel = addPanel(parent, LABELS.panel.star[uiLang]);

                var starRow = addControlRow(starPanel, WIDE_ROW_SPACING);
                starCheck = starRow.add("checkbox", undefined, LABELS.checkbox.star[uiLang]);
                pentagramCheck = starRow.add("checkbox", undefined, LABELS.checkbox.pentagram[uiLang]);
                pentagramCheck.value = false;

                var innerRadiusRow = addControlRow(starPanel);
                innerRadiusLabel = innerRadiusRow.add("statictext", undefined, LABELS.label.innerRadius[uiLang]);
                innerRatioInput = addNumberField(innerRadiusRow, SHAPE_DEFAULTS.innerRatio, 4);
                innerPercentLabel = innerRadiusRow.add("statictext", undefined, "%");

                innerRatioSlider = addSlider(starPanel, SHAPE_DEFAULTS.innerRatio, SHAPE_RANGES.innerRatio, SLIDER_WIDTH);
            }

            /**
             * ［円］パネルと、その中の［アンカーポイント］パネルを組み立てる。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildCirclePanel(parent) {
                circlePanel = addPanel(parent, LABELS.panel.circle[uiLang]);

                superEllipseCheck = circlePanel.add("checkbox", undefined, LABELS.checkbox.superEllipse[uiLang]);
                superEllipseCheck.value = false;

                var superExponentRow = addControlRow(circlePanel);
                superExponentInput = addNumberField(superExponentRow, SHAPE_DEFAULTS.superExponent, 4);
                superExponentSlider = addSlider(superExponentRow, SHAPE_DEFAULTS.superExponent, SHAPE_RANGES.superExponent, SHORT_SLIDER_WIDTH);

                circleAnchorPanel = addPanel(circlePanel, LABELS.panel.anchor[uiLang]);
                var circleAnchorRow = addControlRow(circleAnchorPanel, WIDE_ROW_SPACING);
                for (var i = 0; i < CIRCLE_ANCHOR_CHOICES.length; i++) {
                    circleAnchorRadios[i] = circleAnchorRow.add("radiobutton", undefined, String(CIRCLE_ANCHOR_CHOICES[i]));
                }
                selectRadio(circleAnchorRadios, 2); /* 既定は4アンカー / Four anchors by default */
            }

            /**
             * ［角丸］パネルを組み立てる。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildCornerSmoothingPanel(parent) {
                cornerSmoothingPanel = addPanel(parent, LABELS.panel.cornerSmoothing[uiLang]);

                var cornerRadiusRow = addControlRow(cornerSmoothingPanel);
                cornerRadiusCheck = cornerRadiusRow.add("checkbox", undefined, LABELS.checkbox.cornerRadius[uiLang]);
                cornerRadiusCheck.value = false;
                cornerRadiusInput = addNumberField(cornerRadiusRow, SHAPE_DEFAULTS.cornerRadius, 5);
                cornerRadiusRow.add("statictext", undefined, rulerUnitInfo.label);

                var smoothingLabelRow = addControlRow(cornerSmoothingPanel);
                smoothingLabelRow.add("statictext", undefined, LABELS.label.smoothing[uiLang]);
                smoothingValueLabel = smoothingLabelRow.add("statictext", undefined, String(SHAPE_DEFAULTS.smoothing));
                smoothingValueLabel.characters = 4;

                smoothingSlider = addSlider(cornerSmoothingPanel, SHAPE_DEFAULTS.smoothing, SHAPE_RANGES.smoothing, SLIDER_WIDTH);
            }

            /**
             * ［アンカーポイントの操作］パネルを組み立てる。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildAnchorOpsPanel(parent) {
                var anchorOpsPanel = addPanel(parent, LABELS.panel.anchorOps[uiLang]);

                var roughenAnchorsRow = addControlRow(anchorOpsPanel, TIGHT_ROW_SPACING);
                roughenAnchorsCheck = roughenAnchorsRow.add("checkbox", undefined, LABELS.checkbox.roughenAnchors[uiLang]);
                roughenAnchorsCheck.value = false;
                roughenAnchorsInput = addNumberField(roughenAnchorsRow, SHAPE_DEFAULTS.roughenDetail, 3);
                roughenAnchorsInput.enabled = false;

                splitAtAnchorsCheck = anchorOpsPanel.add("checkbox", undefined, LABELS.checkbox.splitAtAnchors[uiLang]);
                splitAtAnchorsCheck.value = false;

                var strokeCapRow = addControlRow(anchorOpsPanel, TIGHT_ROW_SPACING);
                strokeCapLabel = strokeCapRow.add("statictext", undefined, LABELS.label.strokeCap[uiLang]);
                capButtRadio = strokeCapRow.add("radiobutton", undefined, LABELS.radio.capButt[uiLang]);
                capRoundRadio = strokeCapRow.add("radiobutton", undefined, LABELS.radio.capRound[uiLang]);
                capProjectingRadio = strokeCapRow.add("radiobutton", undefined, LABELS.radio.capProjecting[uiLang]);
                capButtRadio.value = true;
            }

            /**
             * ［オプション］パネルを組み立てる。
             * @param {Group} parent - 追加先のコンテナ
             * @returns {void}
             */
            function buildOptionsPanel(parent) {
                var optionPanel = addPanel(parent, LABELS.panel.option[uiLang]);

                liveShapeCheck = optionPanel.add("checkbox", undefined, LABELS.checkbox.liveShape[uiLang]);
                liveShapeCheck.value = true;

                reuleauxCheck = optionPanel.add("checkbox", undefined, LABELS.checkbox.reuleaux[uiLang]);
                reuleauxCheck.value = false;

                var reuleauxAmountRow = addControlRow(optionPanel);
                reuleauxAmountInput = addNumberField(reuleauxAmountRow, SHAPE_DEFAULTS.reuleauxAmount, 4);
                reuleauxAmountSlider = addSlider(reuleauxAmountRow, SHAPE_DEFAULTS.reuleauxAmount, SHAPE_RANGES.reuleauxAmount, SHORT_SLIDER_WIDTH);
            }

            /**
             * ［画面ズーム］の行を組み立てる。
             * @param {Window} parent - 追加先のコンテナ
             * @returns {Slider} ズームスライダー
             */
            function buildViewZoomRow(parent) {
                var viewZoomRow = parent.add("group");
                viewZoomRow.orientation = "row";
                viewZoomRow.alignChildren = ["center", "center"];
                viewZoomRow.alignment = "center";
                viewZoomRow.margins = [0, 7, 0, 5];
                viewZoomRow.add("statictext", undefined, LABELS.label.viewZoom[uiLang]);

                try {
                    initialZoom = (doc && doc.activeView) ? Number(doc.activeView.zoom) : 1;
                } catch (e) {
                    initialZoom = 1;
                }
                return addSlider(viewZoomRow, initialZoom, SHAPE_RANGES.zoom, ZOOM_SLIDER_WIDTH);
            }

            /**
             * ダイアログ下部のボタン列を組み立てる。
             * @param {Window} parent - 追加先のコンテナ
             * @returns {{previewButton: Button, cancelButton: Button, okButton: Button}} 各ボタン
             */
            function buildButtonRow(parent) {
                var buttonRow = parent.add("group");
                buttonRow.orientation = "row";
                buttonRow.alignChildren = ["fill", "center"];
                buttonRow.alignment = "fill";
                buttonRow.margins = [0, 10, 0, 0];

                var previewButtonGroup = buttonRow.add("group");
                setupRow(previewButtonGroup, ["left", "center"]);
                var previewButton = previewButtonGroup.add("button", undefined, LABELS.button.preview[uiLang]);

                var confirmButtonGroup = buttonRow.add("group");
                setupRow(confirmButtonGroup, ["right", "center"], WIDE_ROW_SPACING);
                var cancelButton = confirmButtonGroup.add("button", undefined, LABELS.button.cancel[uiLang], { name: "cancel" });
                var okButton = confirmButtonGroup.add("button", undefined, LABELS.button.ok[uiLang], { name: "ok" });

                /* ボタンはコンテナいっぱいに広げない / Buttons must not stretch to the container width */
                previewButton.alignment = "left";
                cancelButton.alignment = "right";
                okButton.alignment = "right";

                return { previewButton: previewButton, cancelButton: cancelButton, okButton: okButton };
            }

            // -----------------------------------------
            // 有効・無効の制御 / Enabled state
            // -----------------------------------------

            /**
             * ［それ以外］の入力欄とスライダーの有効・無効をまとめて切り替える。
             * @param {boolean} isEnabled - 有効にするかどうか
             * @returns {void}
             */
            function setCustomSidesEnabled(isEnabled) {
                customSidesInput.enabled = isEnabled;
                customSidesSlider.enabled = isEnabled;
            }

            /**
             * 円のアンカーポイント数のラジオをまとめて有効・無効にする。
             * @param {boolean} isEnabled - 有効にするかどうか
             * @returns {void}
             */
            function setCircleAnchorRadiosEnabled(isEnabled) {
                for (var i = 0; i < circleAnchorRadios.length; i++) {
                    circleAnchorRadios[i].enabled = isEnabled;
                }
            }

            /**
             * 線幅の入力欄を［線］チェックの状態に合わせて有効・無効にする。
             * @returns {void}
             */
            function updateStrokeWidthEnabled() {
                strokeWidthInput.enabled = strokeCheck.value;
                strokeWidthUnitLabel.enabled = strokeCheck.value;
            }

            /**
             * 円パネルの有効・無効を辺の数に応じて切り替える。
             * @param {number} sidesValue - 現在の辺の数
             * @returns {void}
             */
            function updateCirclePanelEnabled(sidesValue) {
                var isEnabled = (sidesValue === 0);
                circlePanel.enabled = isEnabled;
                circleAnchorPanel.enabled = isEnabled;
                if (isEnabled) return;

                /* 円以外に切り替えたら円用の設定を既定へ戻す / Reset the circle options when leaving the circle */
                superEllipseCheck.value = false;
                selectRadio(circleAnchorRadios, 2);
                setCircleAnchorRadiosEnabled(true);
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
                innerRadiusLabel.enabled = isEnabled;
                innerRatioInput.enabled = isEnabled;
                innerPercentLabel.enabled = isEnabled;
                innerRatioSlider.enabled = isEnabled;
            }

            /**
             * スーパー楕円の指数入力と、円のアンカーポイントパネルの有効状態を更新する。
             * @param {number} sidesValue - 現在の辺の数
             * @returns {void}
             */
            function updateSuperEllipseControlsEnabled(sidesValue) {
                var isActive = isSuperEllipseActive(sidesValue);
                superExponentInput.enabled = isActive;
                superExponentSlider.enabled = isActive;
                /* スーパー楕円がONのときはアンカーポイント数を選べない
                   The anchor count cannot be chosen while the superellipse is on */
                circleAnchorPanel.enabled = !isActive;
                setCircleAnchorRadiosEnabled(!isActive);
            }

            /**
             * ルーローの度合いの入力群を有効・無効にする。
             * @returns {void}
             */
            function updateReuleauxAmountEnabled() {
                var isEnabled = (reuleauxCheck.enabled && reuleauxCheck.value);
                reuleauxAmountInput.enabled = isEnabled;
                reuleauxAmountSlider.enabled = isEnabled;
            }

            /**
             * ルーローのチェックボックスを、辺の数とスターの状態に応じて有効・無効にする。
             * @param {number} sidesValue - 現在の辺の数
             * @returns {void}
             */
            function updateReuleauxAvailability(sidesValue) {
                /* スターがONのときは奇数判定を行わない（スター側の制御を優先）
                   While the star is on, the odd-side rule is skipped and the star logic wins */
                if (!starCheck.value) {
                    /* ルーローは奇数辺（3、5、7…）のみ。円（0）と偶数辺は対象外
                       Reuleaux applies to odd side counts only, never to a circle or an even count */
                    var isEnabled = (typeof sidesValue === "number") && (sidesValue > 0) && (sidesValue % 2 === 1);
                    reuleauxCheck.enabled = isEnabled;
                    if (!isEnabled) reuleauxCheck.value = false;
                }
                updateReuleauxAmountEnabled();
            }

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
                    var currentWidth = parseFloat(sizeInput.text);
                    if (!isNaN(currentWidth) && currentWidth > 0) {
                        cornerRadiusInput.text = String(Math.round(currentWidth * SHAPE_DEFAULTS.cornerRadiusRatio * 10) / 10);
                    }
                }
            }

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

            /**
             * 現在のUIの状態からライブシェイプ化の可否を再計算する。
             * 分割・スーパー楕円・4以外のアンカー数・ルーロー・ラフ効果・角丸のいずれかが有効なら使えない。
             * @returns {void}
             */
            function refreshLiveShapeAvailability() {
                var sidesValue = getCurrentSides();
                var isSuperEllipse = isSuperEllipseActive(sidesValue);
                /* ライブシェイプ化できるのは円のアンカーが4のときだけ（スーパー楕円を除く）
                   A live shape is possible only with four circle anchors and no superellipse */
                var isCustomCircleAnchors = (sidesValue === 0 && !isSuperEllipse && getCircleAnchorCount() !== 4);
                var isCornerSmoothing = (sidesValue === 4 && cornerRadiusCheck.value && parseFloat(cornerRadiusInput.text) > 0);

                var isBlocked = splitAtAnchorsCheck.value || isSuperEllipse || isCustomCircleAnchors ||
                    reuleauxCheck.value || roughenAnchorsCheck.value || isCornerSmoothing;
                if (isBlocked) liveShapeCheck.value = false;
                liveShapeCheck.enabled = !isBlocked;
            }

            /**
             * 辺の数に依存する各パネルの有効状態をまとめて更新する。
             * @returns {number} 現在の辺の数
             */
            function refreshPanelStates() {
                var sidesValue = getCurrentSides();
                trianglePanel.enabled = (sidesValue === 3);
                updateCirclePanelEnabled(sidesValue);
                updateCornerSmoothingEnabled(sidesValue);
                updateSuperEllipseControlsEnabled(sidesValue);
                updateStarPanelEnabled(sidesValue);
                updateReuleauxAvailability(sidesValue);
                updateInnerRadiusEnabled();
                refreshLiveShapeAvailability();
                return sidesValue;
            }

            // -----------------------------------------
            // 入力どうしの整合 / Keeping the inputs consistent
            // -----------------------------------------

            /**
             * 回転を強制的にOFFにする。
             * @returns {void}
             */
            function forceRotateOff() {
                rotateCheck.value = false;
                rotateInput.enabled = false;
                rotateUnitLabel.enabled = false;
            }

            /**
             * 回転がOFFのときに使う自動角度を、回転の入力欄に反映する。
             * 辺の数が変わったときは回転がONでも呼び出す。
             * @param {number} sidesValue - 現在の辺の数
             * @returns {number} 反映した角度（対象外の辺の数ならNaN）
             */
            function applyAutoRotationForSides(sidesValue) {
                var angle = NaN;
                if (sidesValue === 0) angle = 45;
                else if (sidesValue >= 3) angle = 360 / (sidesValue * 2);
                if (!isNaN(angle)) rotateInput.text = formatAngle(angle);
                return angle;
            }

            /**
             * 回転をONにしたときの既定角度を決めて入力欄に反映する。
             * 円（辺の数0）はアンカー数ごとに 2→90、3→180、4→45、5→180、6→30 を使う。
             * @returns {void}
             */
            function applyDefaultRotationWhenEnablingRotate() {
                var sidesValue = getCurrentSides();
                if (sidesValue === 0 && !isSuperEllipseActive(sidesValue)) {
                    var angle = CIRCLE_ROTATION_DEFAULTS[getCircleAnchorCount()];
                    rotateInput.text = formatAngle((typeof angle === "number") ? angle : 45);
                    return;
                }
                applyAutoRotationForSides(sidesValue);
                /* 三角形のときは［下］（60°）を既定にする / A triangle defaults to "down" (60 degrees) */
                if (sidesValue === 3) triangleDownRadio.value = true;
            }

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
                    updateReuleauxAmountEnabled();
                }

                if (pentagramCheck.value) {
                    selectRadio(sideRadios, 3); /* 五芒星は5辺 / A pentagram has five sides */
                    setCustomSidesEnabled(false);
                    applyAutoRotationForSides(5);
                    forceRotateOff();
                }

                /* スターがOFFに戻ったら奇数辺の条件でルーローを復帰させる
                   Once the star is off again, restore Reuleaux under the odd-side rule */
                if (!starCheck.value) updateReuleauxAvailability(getCurrentSides());
                updateInnerRadiusEnabled();
            }

            /**
             * スーパー楕円の指数を入力欄とスライダーの両方に反映する。
             * @param {string|number} value - 入力値
             * @returns {number} 反映した指数
             */
            function syncSuperExponentUI(value) {
                value = clampSuperExponent(value);
                superExponentInput.text = String(value);
                superExponentSlider.value = value;
                return value;
            }

            /**
             * ルーローの度合いを入力欄とスライダーの両方に反映する。
             * @param {string|number} value - 入力値
             * @returns {number} 反映した度合い（%）
             */
            function syncReuleauxAmountUI(value) {
                value = clampReuleauxAmount(value);
                reuleauxAmountInput.text = String(value);
                reuleauxAmountSlider.value = value;
                return value;
            }

            // -----------------------------------------
            // プレビュー / Preview
            // -----------------------------------------

            /**
             * 現在のUIから図形生成のパラメーターを組み立てる（プレビューと確定の両方で使う）。
             * @returns {ShapeParams} createShapeに渡すパラメーター一式
             */
            function getCurrentParams() {
                validateStarAndPentagram();
                var sides = refreshPanelStates();

                var useSuperEllipse = isSuperEllipseActive(sides);
                /* スーパー楕円は回転を強制的にOFFにする / The superellipse forces the rotation off */
                if (useSuperEllipse) forceRotateOff();

                /* 回転がOFFのときだけ自動角度を使う（ONなら入力値を尊重）
                   The auto angle is used only while the rotation is off; manual mode keeps the entered angle */
                var angle = parseFloat(rotateInput.text);
                if (!rotateCheck.value) {
                    var autoAngle = applyAutoRotationForSides(sides);
                    if (!isNaN(autoAngle)) angle = autoAngle;
                }

                /* 三角形は向きごとの回転角で上書きする / A triangle overrides the angle with its direction */
                if (sides === 3 && trianglePanel.enabled) {
                    if (triangleRightRadio.value) angle = TRIANGLE_ANGLES.right;
                    else if (triangleLeftRadio.value) angle = TRIANGLE_ANGLES.left;
                    else if (triangleDownRadio.value) angle = TRIANGLE_ANGLES.down;
                    rotateInput.text = formatAngle(angle);
                }

                /* 五芒星の第2半径は黄金比から決まる / A pentagram derives its inner radius from the golden ratio */
                var innerRatio = parseFloat(innerRatioInput.text);
                if (starCheck.value && pentagramCheck.value && sides === 5) {
                    innerRatio = (3 - Math.sqrt(5)) / 2 * 100;
                    innerRatioInput.text = innerRatio.toFixed(2);
                }

                /* 詳細が1の多角形はメニューコマンドでアンカーを追加する。この経路はプレビューできない
                   A detail of 1 on a polygon adds anchors through a menu command, which cannot be previewed */
                roughenAnchorsUseMenuFallback = (sides !== 0 && Math.round(Number(roughenAnchorsInput.text)) === 1);
                var roughenDetail = 0;
                if (roughenAnchorsCheck.value && !roughenAnchorsUseMenuFallback) {
                    roughenDetail = parseFloat(roughenAnchorsInput.text);
                    if (isNaN(roughenDetail) || roughenDetail < 0) roughenDetail = 0;
                }

                /* 円のアンカーポイント数は、円かつスーパー楕円OFFのときだけ有効
                   The circle anchor count only matters for a circle without the superellipse */
                var circleAnchorCount = (sides === 0 && !useSuperEllipse) ? getCircleAnchorCount() : SHAPE_DEFAULTS.circleAnchors;

                return {
                    size: parseFloat(sizeInput.text) * rulerUnitInfo.factorToPt,
                    sides: sides,
                    isStar: starCheck.value,
                    innerRatio: innerRatio,
                    rotateEnabled: rotateCheck.value,
                    angle: angle,
                    splitAtAnchors: splitAtAnchorsCheck.value,
                    strokeCap: splitAtAnchorsCheck.value ? getSelectedStrokeCap() : null,
                    useSuperEllipse: useSuperEllipse,
                    superExponent: clampSuperExponent(superExponentInput.text),
                    circleAnchorCount: circleAnchorCount,
                    useReuleaux: reuleauxCheck.value,
                    reuleauxAmount: clampReuleauxAmount(reuleauxAmountInput.text) / 100,
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
                        radius: parseFloat(cornerRadiusInput.text) * rulerUnitInfo.factorToPt,
                        smoothing: Math.round(smoothingSlider.value)
                    } : null,
                    opacity: clampOpacity(opacityInput.text),
                    roughenDetail: roughenDetail
                };
            }

            /**
             * ［画面にフィット］がONのとき、プレビューが収まるよう表示倍率を合わせる。
             * 呼ぶのは［幅］を変えたときとダイアログを開いたときだけ。ほかの設定でも呼ぶと
             * 1キーごとに倍率が動いて落ち着かないため。
             * 倍率の変更はUndo履歴に残らないので、キャンセル時はdiscardPreview()で戻す。
             * @returns {void}
             */
            function applyFitView() {
                if (!fitViewCheck.value || !previewShape) return;
                try {
                    var activeDoc = app.activeDocument;
                    documentView = activeDoc.activeView;
                    fitViewToItem(activeDoc, previewShape);
                    /* ズームスライダーの表示も合わせる / Keep the zoom slider in step */
                    zoomSlider.value = clampNumber(documentView.zoom, SHAPE_RANGES.zoom, 1);
                    app.redraw();
                } catch (e) { }
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
                if (isNaN(params.size) || isNaN(params.innerRatio)) return;

                previewManager.addStep(function () {
                    previewShape = createShape(app.activeDocument, params);
                    /* ラフ効果も同じステップ内で適用してプレビューに反映する
                       The roughen effect runs inside the same step so the preview shows it */
                    applyRoughenEffect(previewShape, params.roughenDetail);
                });
            }

            /**
             * ［幅］を変えたときの処理。プレビューを更新してから画面にフィットさせる。
             * @returns {void}
             */
            function onSizeChange() {
                updatePreview();
                applyFitView();
            }

            // -----------------------------------------
            // セッション状態 / Session state
            // -----------------------------------------

            /**
             * セッション状態をUIに復元する。
             * @param {object} sessionState - セッション状態
             * @returns {void}
             */
            function applyStateToUI(sessionState) {
                if (!sessionState) return;

                /* 辺の数 / Side count */
                if (typeof sessionState.selectedSideIndex === "number" && sessionState.selectedSideIndex >= 0 && sessionState.selectedSideIndex < sideRadios.length) {
                    selectRadio(sideRadios, sessionState.selectedSideIndex);
                    var isCustomSides = (sessionState.selectedSideIndex === CUSTOM_SIDES_INDEX);
                    if (isCustomSides && typeof sessionState.customSidesText === "string") {
                        customSidesInput.text = sessionState.customSidesText;
                        /* スライダーのつまみも復元した辺の数に合わせる / Move the slider thumb to the restored side count */
                        customSidesSlider.value = clampNumber(customSidesInput.text, SHAPE_RANGES.customSides, SHAPE_DEFAULTS.customSides, true);
                    }
                    setCustomSidesEnabled(isCustomSides);
                }
                if (typeof sessionState.sizeText === "string") sizeInput.text = sessionState.sizeText;
                if (typeof sessionState.fitViewCheck === "boolean") fitViewCheck.value = sessionState.fitViewCheck;

                /* 塗りと線 / Fill and stroke */
                if (typeof sessionState.fillCheck === "boolean") fillCheck.value = sessionState.fillCheck;
                if (typeof sessionState.strokeCheck === "boolean") strokeCheck.value = sessionState.strokeCheck;
                if (typeof sessionState.strokeWidthText === "string") strokeWidthInput.text = sessionState.strokeWidthText;
                if (sessionState.fillColorString) fillSwatch._aiColor = pickerStringToAiColor(sessionState.fillColorString);
                if (sessionState.strokeColorString) strokeSwatch._aiColor = pickerStringToAiColor(sessionState.strokeColorString);
                updateStrokeWidthEnabled();

                /* 不透明度 / Opacity */
                if (typeof sessionState.opacityText === "string") {
                    opacityInput.text = sessionState.opacityText;
                    opacitySlider.value = clampOpacity(sessionState.opacityText);
                }

                /* 角丸 / Corner smoothing */
                if (typeof sessionState.cornerRadiusCheck === "boolean") cornerRadiusCheck.value = sessionState.cornerRadiusCheck;
                if (typeof sessionState.cornerRadiusText === "string") cornerRadiusInput.text = sessionState.cornerRadiusText;
                if (typeof sessionState.smoothingValue === "number") {
                    var restoredSmoothing = clampNumber(sessionState.smoothingValue, SHAPE_RANGES.smoothing, SHAPE_DEFAULTS.smoothing, true);
                    smoothingSlider.value = restoredSmoothing;
                    smoothingValueLabel.text = String(restoredSmoothing);
                }
                updateCornerRadiusInputEnabled();

                /* 回転と三角形 / Rotation and triangle */
                if (typeof sessionState.rotateCheck === "boolean") rotateCheck.value = sessionState.rotateCheck;
                if (typeof sessionState.rotateText === "string") rotateInput.text = sessionState.rotateText;
                rotateInput.enabled = rotateCheck.value;
                rotateUnitLabel.enabled = rotateCheck.value;
                if (sessionState.triangleDir === "left") triangleLeftRadio.value = true;
                else if (sessionState.triangleDir === "down") triangleDownRadio.value = true;
                else if (sessionState.triangleDir === "right") triangleRightRadio.value = true;

                /* スターと五芒星 / Star and pentagram */
                if (typeof sessionState.starCheck === "boolean") starCheck.value = sessionState.starCheck;
                if (typeof sessionState.pentagramCheck === "boolean") pentagramCheck.value = sessionState.pentagramCheck;
                if (!starCheck.value) pentagramCheck.value = false;
                if (typeof sessionState.innerRatioText === "string") innerRatioInput.text = sessionState.innerRatioText;
                innerRatioSlider.value = clampNumber(innerRatioInput.text, SHAPE_RANGES.innerRatio, SHAPE_DEFAULTS.innerRatio, true);

                /* 円 / Circle */
                if (typeof sessionState.superEllipseCheck === "boolean") superEllipseCheck.value = sessionState.superEllipseCheck;
                if (typeof sessionState.superExponentText === "string") syncSuperExponentUI(sessionState.superExponentText);
                if (typeof sessionState.circleAnchorsValue === "number") {
                    var anchorIndex = 2; /* 見つからないときは4アンカー / Fall back to four anchors */
                    for (var i = 0; i < CIRCLE_ANCHOR_CHOICES.length; i++) {
                        if (CIRCLE_ANCHOR_CHOICES[i] === sessionState.circleAnchorsValue) anchorIndex = i;
                    }
                    selectRadio(circleAnchorRadios, anchorIndex);
                }

                /* ［アンカーポイントで分割］は毎回OFFで開く（復元しない）
                   "Split at anchor points" always opens off and is never restored */
                splitAtAnchorsCheck.value = false;
                if (sessionState.strokeCapValue === "round") capRoundRadio.value = true;
                else if (sessionState.strokeCapValue === "projecting") capProjectingRadio.value = true;
                else capButtRadio.value = true;
                updateStrokeCapEnabled();

                /* オプション / Options */
                if (typeof sessionState.reuleauxCheck === "boolean") reuleauxCheck.value = sessionState.reuleauxCheck;
                if (typeof sessionState.liveShapeCheck === "boolean") liveShapeCheck.value = sessionState.liveShapeCheck;
                if (typeof sessionState.roughenAnchorsCheck === "boolean") roughenAnchorsCheck.value = sessionState.roughenAnchorsCheck;
                if (typeof sessionState.roughenAnchorsText === "string") roughenAnchorsInput.text = sessionState.roughenAnchorsText;
                applyRoughenExclusions();

                /* 五芒星やスーパー楕円が有効なら回転を強制的にOFF / Force the rotation off for a pentagram or a superellipse */
                if (pentagramCheck.value || isSuperEllipseActive(getCurrentSides())) forceRotateOff();
                refreshPanelStates();
            }

            /**
             * 現在のUIの状態をセッション状態に保存する。
             * @param {object} sessionState - セッション状態
             * @returns {void}
             */
            function saveStateFromUI(sessionState) {
                if (!sessionState) return;

                var selectedSideIndex = getSelectedRadioIndex(sideRadios);
                sessionState.selectedSideIndex = (selectedSideIndex < 0) ? 0 : selectedSideIndex;
                sessionState.customSidesText = customSidesInput.text;
                sessionState.sizeText = sizeInput.text;
                sessionState.fitViewCheck = fitViewCheck.value;

                /* 塗りと線・不透明度 / Fill and stroke, opacity */
                sessionState.fillCheck = fillCheck.value;
                sessionState.strokeCheck = strokeCheck.value;
                sessionState.strokeWidthText = strokeWidthInput.text;
                /* カラーは常駐エンジンに残るオブジェクト参照ではなく文字列で保存する
                   Colors are stored as strings, not as object references kept alive by the engine */
                sessionState.fillColorString = aiColorToPickerString(fillSwatch._aiColor);
                sessionState.strokeColorString = aiColorToPickerString(strokeSwatch._aiColor);
                sessionState.opacityText = opacityInput.text;

                /* 角丸 / Corner smoothing */
                sessionState.cornerRadiusCheck = cornerRadiusCheck.value;
                sessionState.cornerRadiusText = cornerRadiusInput.text;
                sessionState.smoothingValue = Math.round(smoothingSlider.value);

                /* 回転と三角形 / Rotation and triangle */
                sessionState.rotateCheck = rotateCheck.value;
                sessionState.rotateText = rotateInput.text;
                sessionState.triangleDir = triangleRightRadio.value ? "right" : (triangleLeftRadio.value ? "left" : "down");

                /* スターと円 / Star and circle */
                sessionState.starCheck = starCheck.value;
                sessionState.pentagramCheck = pentagramCheck.value;
                sessionState.innerRatioText = innerRatioInput.text;
                sessionState.superEllipseCheck = superEllipseCheck.value;
                sessionState.superExponentText = superExponentInput.text;
                sessionState.circleAnchorsValue = getCircleAnchorCount();

                /* アンカーポイントの操作とオプション / Anchor point operations and options */
                sessionState.strokeCapValue = capRoundRadio.value ? "round" : (capProjectingRadio.value ? "projecting" : "butt");
                sessionState.liveShapeCheck = liveShapeCheck.value;
                sessionState.roughenAnchorsCheck = roughenAnchorsCheck.value;
                sessionState.roughenAnchorsText = roughenAnchorsInput.text;
                sessionState.reuleauxCheck = reuleauxCheck.value;
            }

            // -----------------------------------------
            // イベントの割り当て / Event bindings
            // -----------------------------------------

            /**
             * ラフ効果と併用できない項目を、ラフ効果の状態に合わせて整える。
             * @returns {void}
             */
            function applyRoughenExclusions() {
                var isRoughenOn = roughenAnchorsCheck.value;
                roughenAnchorsInput.enabled = isRoughenOn;
                if (isRoughenOn) splitAtAnchorsCheck.value = false;
                splitAtAnchorsCheck.enabled = !isRoughenOn;
                updateStrokeCapEnabled();
                refreshLiveShapeAvailability();
            }

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

            /**
             * 辺の数のラジオを選び直したときの処理。
             * @param {number} selectedIndex - 選択したインデックス
             * @returns {void}
             */
            function selectSides(selectedIndex) {
                if (pentagramCheck.value) pentagramCheck.value = false;
                selectRadio(sideRadios, selectedIndex);
                setCustomSidesEnabled(selectedIndex === CUSTOM_SIDES_INDEX);
                /* 辺の数が変わったら回転がONでも角度を更新する / Update the angle on a side change, even while rotation is on */
                applyAutoRotationForSides(getCurrentSides());
                refreshPanelStates();
            }

            /**
             * カラースウォッチをクリックしたときにカラーピッカーを開く。
             * @param {Group} swatchGroup - 対象のスウォッチ
             * @returns {void}
             */
            function pickSwatchColor(swatchGroup) {
                var pickedColor = ColorPicker.show({
                    value: aiColorToPickerString(swatchGroup._aiColor),
                    title: LABELS.dialog.colorPicker[uiLang],
                    lang: uiLang
                });
                if (pickedColor === null) return;
                swatchGroup._aiColor = pickerStringToAiColor(pickedColor);
                redrawSwatch(swatchGroup);
                updatePreview();
            }

            /**
             * 入力欄とスライダーを連動させ、変更のたびにプレビューを更新する。
             * 入力中は入力欄へ書き戻さず（小数点の途中でカーソルが飛ぶため）、
             * 確定（Enter・フォーカス移動）とスライダー操作のときだけ丸めた値を書き戻す。
             * @param {EditText} inputField - 入力欄
             * @param {Slider} slider - スライダー
             * @param {function} clampValue - 値を有効範囲に収める関数
             * @param {function} [afterChange] - 値を反映したあとに呼ぶ処理
             * @returns {void}
             */
            function bindValueAndSlider(inputField, slider, clampValue, afterChange) {
                /**
                 * 値を反映してプレビューを更新する。
                 * @param {string|number} value - 反映する値
                 * @param {boolean} writeBackText - 丸めた値を入力欄へ書き戻すかどうか
                 * @returns {void}
                 */
                function applyValue(value, writeBackText) {
                    value = clampValue(value);
                    if (writeBackText) inputField.text = String(value);
                    slider.value = value;
                    if (afterChange) afterChange();
                    updatePreview();
                }
                inputField.onChanging = function () {
                    /* 入力途中は確定を待つ / Wait for the rest of a value that is still being typed */
                    if (isPartialNumberInput(inputField.text)) return;
                    applyValue(inputField.text, false);
                };
                inputField.onChange = function () { applyValue(inputField.text, true); };
                slider.onChanging = function () { applyValue(slider.value, true); };
            }

            /**
             * 辺の数・回転・三角形のハンドラーを割り当てる。
             * @returns {void}
             */
            function bindShapeHandlers() {
                for (var i = 0; i < sideRadios.length; i++) {
                    (function (radioIndex) {
                        sideRadios[radioIndex].onClick = function () {
                            selectSides(radioIndex);
                            updatePreview();
                        };
                    })(i);
                }
                bindValueAndSlider(customSidesInput, customSidesSlider, function (value) {
                    return clampNumber(value, SHAPE_RANGES.customSides, SHAPE_DEFAULTS.customSides, true);
                }, function () {
                    updateReuleauxAvailability(getCurrentSides());
                });

                sizeInput.onChanging = onSizeChange;
                fitViewCheck.onClick = applyFitView;
                rotateInput.onChanging = updatePreview;
                rotateCheck.onClick = function () {
                    rotateInput.enabled = rotateCheck.value;
                    rotateUnitLabel.enabled = rotateCheck.value;
                    if (rotateCheck.value) applyDefaultRotationWhenEnablingRotate();
                    updatePreview();
                };
                triangleRightRadio.onClick = onTriangleDirectionChange;
                triangleLeftRadio.onClick = onTriangleDirectionChange;
                triangleDownRadio.onClick = onTriangleDirectionChange;
            }

            /**
             * 塗りと線、不透明度のハンドラーを割り当てる。
             * @returns {void}
             */
            function bindAppearanceHandlers() {
                fillCheck.onClick = updatePreview;
                fillSwatch.addEventListener("click", function () { pickSwatchColor(fillSwatch); });
                strokeCheck.onClick = function () {
                    updateStrokeWidthEnabled();
                    updatePreview();
                };
                strokeSwatch.addEventListener("click", function () { pickSwatchColor(strokeSwatch); });
                strokeWidthInput.onChanging = updatePreview;

                /* Shiftドラッグは10%刻み / Shift-dragging snaps to steps of ten percent */
                bindValueAndSlider(opacityInput, opacitySlider, function (value) {
                    if (typeof value === "number" && ScriptUI.environment.keyboardState.shiftKey) {
                        value = Math.round(value / 10) * 10;
                    }
                    return clampOpacity(value);
                });
            }

            /**
             * スター・円・角丸・アンカーポイントの操作・オプションのハンドラーを割り当てる。
             * @returns {void}
             */
            function bindShapeOptionHandlers() {
                starCheck.onClick = function () {
                    validateStarAndPentagram();
                    updatePreview();
                };
                pentagramCheck.onClick = function () {
                    if (pentagramCheck.value) forceRotateOff();
                    updatePreview();
                };
                bindValueAndSlider(innerRatioInput, innerRatioSlider, function (value) {
                    return clampNumber(value, SHAPE_RANGES.innerRatio, SHAPE_DEFAULTS.innerRatio, true);
                }, function () {
                    /* 第2半径を手で決めたら五芒星の固定値から外れる / A hand-picked inner radius leaves the pentagram preset */
                    if (pentagramCheck.value) pentagramCheck.value = false;
                });

                superEllipseCheck.onClick = function () {
                    /* スーパー楕円は円（辺の数0）でだけ効く / The superellipse only applies to a circle */
                    if (isSuperEllipseActive(getCurrentSides())) forceRotateOff();
                    refreshPanelStates();
                    updatePreview();
                };
                bindValueAndSlider(superExponentInput, superExponentSlider, clampSuperExponent);
                for (var i = 0; i < circleAnchorRadios.length; i++) {
                    circleAnchorRadios[i].onClick = function () {
                        refreshLiveShapeAvailability();
                        updatePreview();
                    };
                }

                cornerRadiusCheck.onClick = function () {
                    updateCornerRadiusInputEnabled();
                    refreshLiveShapeAvailability();
                    updatePreview();
                };
                cornerRadiusInput.onChanging = function () {
                    refreshLiveShapeAvailability();
                    updatePreview();
                };
                smoothingSlider.onChanging = function () {
                    smoothingValueLabel.text = String(Math.round(smoothingSlider.value));
                    updatePreview();
                };

                roughenAnchorsCheck.onClick = function () {
                    applyRoughenExclusions();
                    updatePreview();
                };
                roughenAnchorsInput.onChanging = updatePreview;
                splitAtAnchorsCheck.onClick = function () {
                    /* 分割後は開いたパスになるので、塗りではなく線で見せる / Split segments are open paths, so show them with a stroke */
                    if (splitAtAnchorsCheck.value) {
                        fillCheck.value = false;
                        strokeCheck.value = true;
                        updateStrokeWidthEnabled();
                    }
                    updateStrokeCapEnabled();
                    refreshPanelStates();
                    updatePreview();
                };
                capButtRadio.onClick = updatePreview;
                capRoundRadio.onClick = updatePreview;
                capProjectingRadio.onClick = updatePreview;

                reuleauxCheck.onClick = function () {
                    /* 有効にするたび既定値（100%）に戻す / Reset to the default amount whenever it is enabled */
                    if (reuleauxCheck.value) syncReuleauxAmountUI(SHAPE_DEFAULTS.reuleauxAmount);
                    updateReuleauxAmountEnabled();
                    refreshLiveShapeAvailability();
                    updatePreview();
                };
                bindValueAndSlider(reuleauxAmountInput, reuleauxAmountSlider, clampReuleauxAmount);
            }

            /**
             * キーボードショートカットを割り当てる。
             * E：円（0）／A：回転／S：スター／P：五芒星／D：アンカーポイントで分割／L・R・B：三角形の向き。
             * @returns {void}
             */
            function bindKeyboardShortcuts() {
                /**
                 * 三角形のショートカット。辺の数を3にして向きを決める。
                 * @param {RadioButton} directionRadio - 向きのラジオボタン
                 * @returns {void}
                 */
                function applyTriangleShortcut(directionRadio) {
                    selectSides(1);
                    directionRadio.value = true;
                    onTriangleDirectionChange();
                }

                dialog.addEventListener("keydown", function (event) {
                    if (!event || !event.keyName) return;
                    /* 入力欄の編集中はショートカットを発火させない
                       Shortcuts must not fire while a text field is being edited */
                    if (isTextInputTarget(event.target)) return;

                    switch (event.keyName.toUpperCase()) {
                        case "E":
                            selectSides(0);
                            updatePreview();
                            break;
                        case "L":
                            applyTriangleShortcut(triangleLeftRadio);
                            break;
                        case "R":
                            applyTriangleShortcut(triangleRightRadio);
                            break;
                        case "B":
                            applyTriangleShortcut(triangleDownRadio);
                            break;
                        case "D":
                            splitAtAnchorsCheck.value = !splitAtAnchorsCheck.value;
                            splitAtAnchorsCheck.onClick();
                            break;
                        case "A":
                            rotateCheck.value = !rotateCheck.value;
                            rotateCheck.onClick();
                            break;
                        case "S":
                            starCheck.value = !starCheck.value;
                            starCheck.onClick();
                            break;
                        case "P":
                            /* 五芒星はスターがONのときだけ / The pentagram needs the star to be on */
                            starCheck.value = true;
                            pentagramCheck.value = !pentagramCheck.value;
                            pentagramCheck.onClick();
                            break;
                        default:
                            return;
                    }
                    event.preventDefault();
                });
            }

            // -----------------------------------------
            // ダイアログの組み立てと表示 / Building and showing the dialog
            // -----------------------------------------

            setupWindow(dialog);

            var columnsGroup = dialog.add("group");
            columnsGroup.orientation = "row";
            columnsGroup.alignChildren = ["fill", "top"];
            columnsGroup.spacing = COLUMN_SPACING;

            /* 左カラム（辺の数・回転・塗りと線・幅） / Left column: sides, rotation, fill and stroke, width */
            var leftColumn = columnsGroup.add("group");
            leftColumn.orientation = "column";
            leftColumn.alignChildren = "fill";
            leftColumn.alignment = "top";
            leftColumn.spacing = DENSE_PANEL_SPACING;

            buildSidesPanel(leftColumn);
            buildRotatePanel(leftColumn);
            buildFillStrokePanel(leftColumn);
            buildWidthPanel(leftColumn);
            buildFitViewRow(leftColumn);

            /* 右カラム（スター・円・角丸・アンカー・オプション） / Right column: star, circle, corners, anchors, options */
            var rightColumn = columnsGroup.add("group");
            rightColumn.orientation = "column";
            rightColumn.alignChildren = "fill";
            rightColumn.alignment = "top";
            rightColumn.spacing = DENSE_PANEL_SPACING;

            buildStarPanel(rightColumn);
            buildCirclePanel(rightColumn);
            buildCornerSmoothingPanel(rightColumn);
            buildAnchorOpsPanel(rightColumn);
            buildOptionsPanel(rightColumn);

            var zoomSlider = buildViewZoomRow(dialog);
            var dialogButtons = buildButtonRow(dialog);

            bindShapeHandlers();
            bindAppearanceHandlers();
            bindShapeOptionHandlers();
            bindKeyboardShortcuts();

            /* レイアウトが決まる前に復元しておく。壊れた保存値でも既定値で開けるようにする
               Restore before the layout is measured; a broken stored value must still let the dialog open */
            try { applyStateToUI(getSessionState()); } catch (e) { }
            updateStrokeWidthEnabled();
            updateCornerRadiusInputEnabled();
            updateStrokeCapEnabled();

            zoomSlider.onChanging = function () {
                try {
                    if (!documentView) documentView = doc.activeView;
                    documentView.zoom = Number(zoomSlider.value);
                    app.redraw();
                } catch (e) { }
            };

            /* ドキュメントは通常プレビュー表示で開くので、ONから始めて表示と実態を合わせる
               A document normally opens in preview mode, so start on to keep the label truthful */
            var isViewPreviewOn = true;
            dialogButtons.previewButton.text = "● " + LABELS.button.preview[uiLang];
            dialogButtons.previewButton.onClick = function () {
                isViewPreviewOn = !isViewPreviewOn;
                dialogButtons.previewButton.text = (isViewPreviewOn ? "● " : "") + LABELS.button.preview[uiLang];
                try { app.executeMenuCommand('preview'); } catch (e) { }
            };

            /**
             * プレビューを巻き戻し、画面ズームを開いたときの倍率へ戻す。
             * @returns {void}
             */
            function discardPreview() {
                previewManager.rollback();
                previewShape = null;
                try {
                    if (documentView && initialZoom != null) documentView.zoom = initialZoom;
                } catch (e) { }
            }

            dialogButtons.cancelButton.onClick = function () {
                discardPreview();
                dialog.close();
            };
            dialogButtons.okButton.onClick = function () {
                /* ウィジェットが生きているうちにパラメーターを確定させる
                   Capture the parameters while the dialog widgets are still alive */
                try { finalParams = getCurrentParams(); } catch (e) { finalParams = null; }
                applyLiveShape = liveShapeCheck.value;
                roughenAnchorsDetail = roughenAnchorsCheck.value ? parseFloat(roughenAnchorsInput.text) : 0;
                isConfirmed = true;
                dialog.close();
            };

            dialog.onShow = function () {
                /* 環境によっては無視されるため不透明度を先に適用 / Apply the opacity first, some hosts ignore it */
                try { dialog.opacity = DIALOG_OPACITY; } catch (e) { }

                /* ［アンカーポイントで分割］は毎回OFFで開く / The split option always opens off */
                splitAtAnchorsCheck.value = false;
                updateStrokeCapEnabled();
                refreshPanelStates();
                /* 開いた時点でも一度フィットさせる / Fit once when the dialog opens, too */
                onSizeChange();

                /* Illustratorの起動中は前回のダイアログ位置を再利用 / Reuse the last dialog position while Illustrator is running */
                var savedLocation = getSavedDialogLocation();
                if (savedLocation) {
                    dialog.location = savedLocation;
                } else {
                    dialog.center();
                    dialog.location = [dialog.location[0] + DIALOG_OFFSET_X, dialog.location[1] + DIALOG_OFFSET_Y];
                }
            };

            dialog.onClose = function () {
                /* 保存に失敗しても閉じる処理は続ける（OK時の確定を巻き添えにしない）
                   A failed save must not abort the close, which would also abort the confirmed shape */
                try {
                    saveDialogLocation(dialog);
                    saveStateFromUI(getSessionState());
                } catch (e) { }

                /* キャンセル時はプレビューを巻き戻してUndo履歴を汚さない / Roll the preview back on cancel */
                if (!isConfirmed) discardPreview();
            };

            dialog.show();

            /* OKなら1回のUndoで取り消せる形で確定する / Finalize as a single undoable action when OK was pressed */
            if (isConfirmed) {
                previewManager.confirm(function () {
                    if (!finalParams) return;
                    if (isNaN(finalParams.size) || isNaN(finalParams.innerRatio)) return;
                    previewShape = createShape(app.activeDocument, finalParams);
                });
            }

            if (!isConfirmed || !previewShape) {
                /* キャンセル時にプレビューが残らないようにする / Make sure no preview survives a cancel */
                previewManager.rollback();
                previewShape = null;
                return null;
            }
            return true;
        }

        // =========================================
        // メイン処理 / Main
        // =========================================

        /**
         * 確定後にラフ効果でアンカーポイントを追加する。
         * @param {Document} doc - 対象ドキュメント
         * @returns {void}
         */
        function addAnchorsAfterConfirm(doc) {
            if (!(roughenAnchorsDetail > 0)) return;
            try {
                if (roughenAnchorsUseMenuFallback && Math.round(roughenAnchorsDetail) === 1) {
                    /* この経路だけはメニューコマンドなのでプレビューには出ない
                       Only this path uses a menu command, so it cannot appear in the preview */
                    app.executeMenuCommand('Add Anchor Points2');
                    return;
                }
                for (var i = 0; i < doc.selection.length; i++) {
                    applyRoughenEffect(doc.selection[i], roughenAnchorsDetail);
                }
            } catch (e) { }
        }

        /**
         * ダイアログを開き、確定した図形を作成して後処理を行う。
         * @returns {void}
         */
        function main() {
            if (app.documents.length === 0) {
                alert(LABELS.alert.noDocument[uiLang]);
                return;
            }
            var rulerUnitInfo = getUnitInfoFromPreference("rulerType");
            var strokeUnitInfo = getUnitInfoFromPreference("strokeUnits");
            if (!showInputDialog(rulerUnitInfo, strokeUnitInfo)) return;

            /* ダイアログ表示中にドキュメントが切り替わった場合に備えて取得し直す
               Re-acquire the active document in case it changed while the dialog was open */
            var doc = app.activeDocument;
            finalizeShape(doc);
            addAnchorsAfterConfirm(doc);
            if (applyLiveShape) app.executeMenuCommand('Convert to Shape');
        }

        main();

    })();

})();
