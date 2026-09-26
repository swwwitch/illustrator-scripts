#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブまたはすべてのアートボードと同じサイズの長方形を、オフセットを考慮して描画します。
カラー・配置位置・対象をライブプレビューで確かめながら指定でき、描画後に「ガイドに変換」「ライブシェイプ化」を適用できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDrawArtboardRectangle.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n1ba88513a9c8

### Overview

Draws a rectangle the size of the active artboard, or of every artboard, taking an offset into account.
Color, placement and target scope are set with a live preview, and the result can be converted to guides or to a live shape.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDrawArtboardRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartDrawArtboardRectangle";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartDrawArtboardRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartDrawArtboardRectangle.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n1ba88513a9c8"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 入力中のプレビュー遅延（タイプしやすさ優先）/ Preview delay while typing */
    var PREVIEW_DELAY_TYPING_MS = 110; /* 推奨 100–120ms / recommend 100–120ms */

    /* K100モードの不透明度（%）/ Opacity (%) used by the K100 color mode */
    var K100_OPACITY = 15;

    /* 「bgレイヤー」配置で使うレイヤー名 / Layer name used by the "bg layer" placement */
    var BG_LAYER_NAME = 'bg';

    /* 描画先が見つからないときに作るレイヤー名 / Layer created when no writable layer is found */
    var FALLBACK_LAYER_NAME = '_auto_draw';

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ダイアログの初期位置・不透明度 / Dialog position & opacity */
    var DIALOG_OFFSET_X = 300;  /* 右(+)／左(-) / shift right (+) / left (-) */
    var DIALOG_OFFSET_Y = 0;    /* 下(+)／上(-) / shift down (+) / up (-) */
    var DIALOG_OPACITY = 0.98;  /* 0.0 - 1.0 */

    /* 余白と間隔 / Margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] */
    var PANEL_SPACING = 8;                /* パネル内の要素間隔 */
    var COLUMN_SPACING = 12;              /* 2カラムの間隔 */
    var STACK_SPACING = 10;               /* カラム内のパネル間隔・広めの行間 */
    var TIGHT_SPACING = 6;                /* 詰めた行間 */

    /* CMYK入力欄の固定幅（ラベルと桁を揃える）/ Fixed width that aligns CMYK labels and fields */
    var CMYK_FIELD_WIDTH = 40;

    /**
     * パネルの共通設定
     * @param {Panel} panel - 対象パネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
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
     * グループの共通設定（row/column で整列を切り替え）
     * @param {Group} group - 対象グループ
     * @param {string} [orientation] - "row" または "column"（省略時は "column"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupGroup(group, orientation, spacing) {
        var groupOrientation = orientation || "column";
        group.orientation = groupOrientation;
        /* row は横並びなので縦中央、column は縦並びなので左揃え / row: vertically centered, column: left-aligned */
        group.alignChildren = (groupOrientation === "row") ? ["left", "center"] : ["left", "top"];
        group.alignment = "fill";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // カラーモード / Color modes
    // =========================================

    /* 塗りの決め方を表す定数 / How the fill color is decided */
    var ColorMode = {
        NONE: 'none',
        K100: 'k100',
        HEX: 'hex',
        CMYK: 'cmyk'
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* ラベル定義（カテゴリ別）/ Label definitions (by category) */
    var LABELS = {
        dialog: {
            title: { ja: "アートボードサイズの長方形を描画", en: "Draw Artboard Size Rectangle" }
        },
        panel: {
            offset: { ja: "オフセット", en: "Offset" },
            color: { ja: "カラー", en: "Color" },
            placement: { ja: "配置位置", en: "Placement" },
            target: { ja: "対象", en: "Target" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            colorNone: { ja: "なし", en: "None" },
            colorK100: { ja: "K100、不透明度15%", en: "K100, Opacity 15%" },
            colorHex: { ja: "HEX", en: "HEX" },
            colorCmyk: { ja: "CMYK", en: "CMYK" },
            placeFront: { ja: "最前面", en: "Front" },
            placeBack: { ja: "最背面", en: "Back" },
            placeBgLayer: { ja: "bgレイヤー", en: "bg Layer" },
            currentArtboard: { ja: "現在のアートボード", en: "Current Artboard" },
            allArtboards: { ja: "すべてのアートボード", en: "All Artboards" }
        },
        checkbox: {
            bleed: { ja: "裁ち落とし", en: "Bleed" },
            makeGuide: { ja: "ガイドに変換", en: "Convert to Guides" },
            convertToLiveShape: { ja: "ライブシェイプ化", en: "Convert to Live Shape" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            previewOutline: { ja: "アウトライン表示", en: "Outline" },
            previewPreview: { ja: "プレビュー表示", en: "Preview" }
        },
        tooltip: {
            offsetInput: {
                ja: "アートボード境界から外側へ広げる量を指定します。負の値で内側へ縮めます。",
                en: "Set how far the bounds expand outward from the artboard. Use a negative value to shrink inward."
            },
            bleed: {
                ja: "現在の単位に応じて、裁ち落とし相当の値を自動入力します。",
                en: "Automatically fills a bleed-equivalent offset based on the current unit."
            },
            colorNone: { ja: "塗りも線もない長方形を描画します。", en: "Draws the rectangle with no fill and no stroke." },
            hexInput: {
                ja: "#RRGGBB（#RGB 短縮・red などの色名・gray50 も可）で塗りカラーを指定します。",
                en: "Enter a fill color: #RRGGBB (also #RGB shorthand, color names like red, or gray50)."
            },
            colorCmyk: {
                ja: "CMYK値で塗りを指定します。RGBドキュメントではRGBに換算して塗ります。",
                en: "Sets the fill from CMYK values. In an RGB document the values are converted to RGB."
            },
            cmykInput: {
                ja: "0〜100の範囲でCMYK値を指定します。未入力は0として扱います。",
                en: "Enter CMYK values from 0 to 100. Empty fields are treated as 0."
            },
            placeFront: { ja: "現在のレイヤー内で最前面に配置します。", en: "Places the rectangle at the front of the current layer." },
            placeBack: { ja: "現在のレイヤー内で最背面に配置します。", en: "Places the rectangle at the back of the current layer." },
            bgLayer: {
                ja: "bgレイヤーを作成または使用し、レイヤーの最背面へ配置します。",
                en: "Creates or uses the bg layer and places it at the back of the layer stack."
            },
            previewToggle: {
                ja: "Illustratorのアウトライン表示／プレビュー表示を切り替えます。",
                en: "Toggles Illustrator's Outline and Preview display modes."
            },
            convertToLiveShape: {
                ja: "Illustratorのメニューコマンドで長方形をライブシェイプ化します。中心点も表示されます。",
                en: "Uses Illustrator's menu command to convert rectangles to Live Shapes. The center point is also shown."
            },
            makeGuide: { ja: "描画した長方形をガイドに変換します。", en: "Converts the drawn rectangles to guides." },
            stepUp: {
                ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepDown: {
                ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
            },
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        warning: {
            hexInvalid: { ja: "正しい #RRGGBB を入力してください", en: "Enter a valid #RRGGBB value" },
            hexEmpty: { ja: "HEX未入力（# のみ）", en: "HEX not entered (# only)" },
            cmykRange: {
                ja: "0–100 の範囲にしてください（未入力は 0 として扱います）",
                en: "Enter a value from 0 to 100 (empty fields are treated as 0)"
            },
            singleArtboard: { ja: "アートボードが1つのため選択できません", en: "Disabled: only one artboard exists" }
        },
        objectName: {
            previewLayer: { ja: "_preview", en: "_preview" },
            rect: { ja: "<長方形>", en: "<Rectangle>" },
            guide: { ja: "<ガイド>", en: "<Guide>" },
            previewRect: { ja: "__プレビュー_アートボードサイズの長方形", en: "__Preview_ArtboardSizeRectangle" }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー、{slash}→「/」に展開）
     * @param {string} key - "panel.offset" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(key) {
        var labelNode = LABELS;
        var keyParts = String(key).split('.');
        for (var i = 0; i < keyParts.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[keyParts[i]];
        }
        var labelValue = (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : key;
        return String(labelValue).replace(/\{slash\}/g, '/');
    }

    // =========================================
    // ダイアログ共通ユーティリティ / Dialog utilities
    // =========================================

    /* =========================================
     * DialogPersist util (extractable)
     * ダイアログの不透明度・初期位置を共通化するユーティリティ。
     * 使い方:
     *   DialogPersist.setOpacity(dialog, 0.95);
     *   DialogPersist.applyInitialOffset(dialog, offsetX, offsetY); // onShow などで
     * ========================================= */
    (function (globalObject) {
        if (!globalObject.DialogPersist) {
            globalObject.DialogPersist = {
                setOpacity: function (dialog, opacity) {
                    try { dialog.opacity = opacity; } catch (e) { }
                },
                applyInitialOffset: function (dialog, offsetX, offsetY) {
                    try {
                        var currentLocation = dialog.location;
                        dialog.location = [currentLocation[0] + (offsetX | 0), currentLocation[1] + (offsetY | 0)];
                    } catch (e) { }
                }
            };
        }
    })($.global);

    /* 入力中のホットキー抑止用に、フォーカス中のコントロールを保持 / Control that currently owns focus */
    var focusedField = null;

    /**
     * 入力欄のフォーカスを追跡し、入力中はダイアログのホットキーを無効にする
     * 単一 boolean だと「新フィールドの focus → 旧フィールドの blur」の順で false に落ちるため、
     * コントロール自体を保持して自分の blur のときだけクリアする（順序非依存）。
     * @param {EditText} fieldControl - 追跡対象の入力欄
     * @returns {void}
     */
    function trackFocusForHotkeys(fieldControl) {
        fieldControl.addEventListener('focus', function () {
            focusedField = fieldControl;
        });
        fieldControl.addEventListener('blur', function () {
            if (focusedField === fieldControl) focusedField = null;
        });
    }

    /**
     * 入力欄を淡黄色でハイライト表示する（選択中のカラーモードを示す）
     * @param {EditText} fieldControl - 対象の入力欄
     * @param {boolean} highlighted - true でハイライト、false で通常表示
     * @returns {void}
     */
    function setFieldHighlight(fieldControl, highlighted) {
        try {
            var graphics = fieldControl.graphics;
            var backgroundRgb = highlighted ? [1, 1, 0.85] : [1, 1, 1];
            var foregroundRgb = highlighted ? [0.2, 0.2, 0] : [0, 0, 0];
            graphics.backgroundColor = graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundRgb);
            graphics.foregroundColor = graphics.newPen(graphics.PenType.SOLID_COLOR, foregroundRgb, 1);
            fieldControl.notify('onDraw');
        } catch (e) { }
    }

    /**
     * 入力欄の文字色を警告(赤)／通常(黒)に切り替える
     * @param {EditText} fieldControl - 対象の入力欄
     * @param {boolean} isWarning - true で赤、false で黒
     * @returns {void}
     */
    function setFieldWarnColor(fieldControl, isWarning) {
        try {
            var graphics = fieldControl.graphics;
            graphics.foregroundColor = graphics.newPen(graphics.PenType.SOLID_COLOR, isWarning ? [1, 0, 0] : [0, 0, 0], 1);
        } catch (e) { }
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に↑↓キーの増減処理を別に付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    // -----------------------------------------
    // ステップボタンの寸法・増減量 / Stepper metrics and steps
    // -----------------------------------------
    var STEPPER_BUTTON_WIDTH   = 20;  /* ∧∨ボタンの幅 / button width */
    var STEPPER_BUTTON_HEIGHT  = 11;  /* ∧∨ボタン1つの高さ（2つ重ねた全体の高さは22） / button height (22 for the pair) */
    var STEPPER_CORNER_RADIUS  = 2;   /* 枠の角丸の半径（ScriptUIは円弧を描けないため短い線分で近似） / corner radius, approximated with segments */
    var STEPPER_FIELD_SPACING  = 3;   /* 項目名と∧∨の間隔 / spacing between the label and the stepper */
    var STEPPER_SIDE_MARGIN    = 3;   /* ∧∨の左に足す余白（右は入力欄に突き合わせる） / extra space left of the stepper */
    var STEPPER_SHIFT_MULTIPLE = 10;  /* shift＋クリックでそろえる倍数 / Shift-click snaps to multiples of this */
    var STEPPER_OPTION_STEP    = 0.1; /* option＋クリックの増減量 / Option-click step */

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    /**
     * UIがダークテーマかどうかを判定する（Illustrator・InDesign の両方に対応）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkStepperUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    var STEPPER_UI_DARK           = isDarkStepperUI();
    /* UIの明るさは4段階あり、段階ごとに背景色が違う。どの段階でも背景に対する差で見せるよう、黒・白の半透明を重ねる。
       ダーク側は Illustrator 標準のスピナー（［グリッドに分割］）で実測、明るい側は最も明るい段階（背景 約0.94）から逆算
       UI brightness has four levels with different backgrounds, so colors are translucent overlays that follow the
       dialog background. Dark values are measured from Illustrator's own spinner; light values derived for the lightest level */
    var STEPPER_FILL_COLOR        = STEPPER_UI_DARK ? [0, 0, 0, 0.10]  : [1, 1, 1, 0.50];  /* 地 / background */
    var STEPPER_FRAME_COLOR       = STEPPER_UI_DARK ? [1, 1, 1, 0.07]  : [0, 0, 0, 0.10];  /* 枠線 / frame */
    var STEPPER_PRESSED_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 0.12]  : [0, 0, 0, 0.13];  /* 押下中 / pressed */
    var STEPPER_CHEVRON_COLOR     = STEPPER_UI_DARK ? [1, 1, 1, 1]     : [0, 0, 0, 0.70];  /* 山形の線 / chevron */
    var STEPPER_DIM_FILL_COLOR    = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [1, 1, 1, 0.30];  /* 無効時の地 / background when disabled */
    var STEPPER_DIM_FRAME_COLOR   = STEPPER_UI_DARK ? [1, 1, 1, 0.035] : [0, 0, 0, 0.05];  /* 無効時の枠線（ダークは地と同じで見せない） / frame when disabled */
    var STEPPER_DIM_CHEVRON_COLOR = STEPPER_UI_DARK ? [1, 1, 1, 0.20]  : [0, 0, 0, 0.25];  /* 無効時の山形 / chevron when disabled */

    // -----------------------------------------
    // 数値欄を作る（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 「項目名・∧∨・入力欄」をひと組にした数値欄を追加する。
     * ↑↓キーでも∧∨と同じように増減する。直接入力した値も、フォーカスが外れたときに
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
     * @param {Group|Panel} parent - 追加先
     * @param {Object} fieldOptions - label（コロン込みの項目名）/ labelWidth / text / characters /
     *     step / min / max / integer（true で整数のみ）/ unit / onStep
     * @returns {EditText} 入力欄（項目名は .fieldLabel、∧∨は .stepperGroup で参照できる）
     */
    function addSteppedField(parent, fieldOptions) {
        var fieldRowGroup = parent.add("group");
        fieldRowGroup.orientation = "row";
        fieldRowGroup.alignChildren = ["left", "center"];
        fieldRowGroup.spacing = STEPPER_FIELD_SPACING;

        var fieldLabel = fieldRowGroup.add("statictext", undefined, fieldOptions.label || "");
        if (fieldOptions.labelWidth) {
            fieldLabel.preferredSize.width = fieldOptions.labelWidth;
            fieldLabel.justify = "right";
        }

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperInputGroup = fieldRowGroup.add("group");
        stepperInputGroup.orientation = "row";
        stepperInputGroup.alignChildren = ["left", "center"];
        stepperInputGroup.spacing = 0;
        stepperInputGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, fieldOptions);
        numberInput = stepperInputGroup.add("edittext", undefined, fieldOptions.text || "");
        numberInput.characters = fieldOptions.characters || 6;
        numberInput.fieldLabel = fieldLabel;
        numberInput.stepperGroup = stepperGroup;

        /* ↑↓キーも∧∨と同じ処理で増減する（増減量・下限・上限・単位・修飾キーをそろえる） / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(numberInput, stepperGroup);

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) {
                numberInput.text = numberInput.lastValidText;
                return;
            }
            writeSteppedValue(numberInput, value, fieldOptions);
        };
        return numberInput;
    }

    /**
     * 数値欄の有効／無効を、項目名・∧∨ごとまとめて切り替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setSteppedFieldEnabled(numberInput, isEnabled) {
        numberInput.enabled = isEnabled;
        numberInput.fieldLabel.enabled = isEnabled;
        numberInput.stepperGroup.enabled = isEnabled;
        /* ∧∨は自作描画なので、描き直してディム表示を切り替える / redraw the custom-drawn buttons to update the dimming */
        for (var i = 0; i < numberInput.stepperGroup.children.length; i++) {
            redrawStepperGroup(numberInput.stepperGroup.children[i]);
        }
    }

    /**
     * 入力欄の値を増減する∧∨ボタンを、隙間なく縦に積んで追加する
     * @param {Group|Panel} parent - 追加先
     * @param {Function} getNumberInput - 対象の入力欄を返す関数（入力欄を∧∨より後に作れるよう、クリック時に引く）
     * @param {Object} stepOptions - step（増減量）/ min / max / integer / unit（例 " mm"）/ onStep(numberInput)
     * @returns {Group} ∧∨をまとめた group（.stepBy(direction) で同じ増減を呼べる）
     */
    function addStepper(parent, getNumberInput, stepOptions) {
        var stepperGroup = parent.add("group");
        stepperGroup.orientation = "column";
        stepperGroup.spacing = 0; /* 2つのボタンをつなげて1つの枠に見せる / join the buttons into one frame */
        stepperGroup.margins = [STEPPER_SIDE_MARGIN, 0, 0, 0]; /* 右は入力欄に突き合わせる / butt against the field on the right */
        stepperGroup.alignment = ["left", "center"];

        /**
         * 入力欄の値を増減する（shift を押しながらなら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ。下限・上限で止める）
         * @param {number} direction - 増やすなら 1、減らすなら -1
         * @returns {void}
         */
        function stepBy(direction) {
            var numberInput = getNumberInput();
            if (!isStepperEnabledInTree(numberInput)) return; /* 入力欄か親が無効の間は動かさない */
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? "tooltip.stepUpInteger" : "tooltip.stepUp";
        var downTooltip = stepOptions.integer ? "tooltip.stepDownInteger" : "tooltip.stepDown";
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
     * @param {EditText} numberInput - 対象の入力欄
     * @param {Group} stepperGroup - addStepper() で作った∧∨
     * @returns {void}
     */
    function bindSteppedArrowKeys(numberInput, stepperGroup) {
        numberInput.addEventListener("keydown", function (event) {
            if (event.keyName !== "Up" && event.keyName !== "Down") return;
            stepperGroup.stepBy(event.keyName === "Up" ? 1 : -1);
            event.preventDefault(); /* カーソル移動を止める / keep the caret from moving */
        });
    }

    // -----------------------------------------
    // 値の計算 / Value helpers
    // -----------------------------------------
    /**
     * 押された修飾キーに応じて、1回分増減した値を返す
     * （shift なら STEPPER_SHIFT_MULTIPLE の倍数へ、option なら STEPPER_OPTION_STEP ずつ、それ以外は step の倍数へ（1.5→2、1.5→1）。
     * 整数の欄では option を無視して step の倍数へ）
     * @param {number} value - 元の値
     * @param {number} direction - 増やすなら 1、減らすなら -1
     * @param {Object} stepOptions - step（通常の増減量。省略時は 1）/ integer
     * @returns {number} 増減した値（下限・上限は未適用）
     */
    function computeSteppedValue(value, direction, stepOptions) {
        var keyState = ScriptUI.environment.keyboardState;
        if (keyState.shiftKey) return snapStepperToNextMultiple(value, STEPPER_SHIFT_MULTIPLE, direction);
        if (keyState.altKey && !stepOptions.integer) return value + direction * STEPPER_OPTION_STEP;
        return snapStepperToNextMultiple(value, stepOptions.step || 1, direction);
    }

    /**
     * 値を、指定した方向にある次の倍数へ移す（230→240、232→240、下げるときは 232→230、230→220）
     * @param {number} value - 元の値
     * @param {number} multiple - 倍数の単位（例 10）
     * @param {number} direction - 上げるなら 1、下げるなら -1
     * @returns {number} 移した値
     */
    function snapStepperToNextMultiple(value, multiple, direction) {
        if (direction > 0) return Math.floor(value / multiple) * multiple + multiple;
        return Math.ceil(value / multiple) * multiple - multiple;
    }

    /**
     * 値を下限・上限の範囲に収める
     * @param {number} value - 数値
     * @param {Object} rangeOptions - min / max（どちらも省略可）
     * @returns {number} 範囲に収めた値
     */
    function clampSteppedValue(value, rangeOptions) {
        if (rangeOptions.min !== undefined && value < rangeOptions.min) return rangeOptions.min;
        if (rangeOptions.max !== undefined && value > rangeOptions.max) return rangeOptions.max;
        return value;
    }

    /**
     * 値を整数化・下限・上限でそろえ、単位を付けて入力欄に書き込む（直前の正しい値としても控える）
     * @param {EditText} numberInput - 書き込む入力欄
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {void}
     */
    function writeSteppedValue(numberInput, value, valueOptions) {
        numberInput.text = formatSteppedValue(value, valueOptions);
        numberInput.lastValidText = numberInput.text;
    }

    /**
     * 値を整数化・下限・上限でそろえ、丸めて単位を付けた表示用の文字列にする。
     * 整数化してから下限で止めるので、「整数・下限1」の欄に 0.4 が入っても 1 になる
     * @param {number} value - 数値
     * @param {Object} valueOptions - integer / min / max / unit（どれも省略可）
     * @returns {string} 入力欄に入れる文字列（例 "20 mm"）
     */
    function formatSteppedValue(value, valueOptions) {
        if (valueOptions.integer) value = Math.round(value);
        return formatStepperNumber(clampSteppedValue(value, valueOptions)) + (valueOptions.unit || "");
    }

    /**
     * 小数第2位で丸めた数値を文字列で返す
     * @param {number} value - 数値
     * @returns {string} 表示用の数値文字列
     */
    function formatStepperNumber(value) {
        return String(Math.round(value * 100) / 100);
    }

    // -----------------------------------------
    // ∧∨ボタンの描画 / Drawing
    // -----------------------------------------
    /**
     * 山形（∧／∨）の極小ボタンを作成する。
     * 上下2つを隙間なく積んで1つの枠に見えるよう、枠線は外側の辺だけ描き（上ボタンは上側、下ボタンは下側）、
     * 継ぎ目に線は引かない
     * @param {Group|Panel} parent - 追加先
     * @param {string} direction - "up" または "down"
     * @param {Function} onClickFn - クリック時の処理
     * @returns {Group} ボタンとして使う group
     */
    function makeStepperChevronButton(parent, direction, onClickFn) {
        var buttonWidth = STEPPER_BUTTON_WIDTH;
        var buttonHeight = STEPPER_BUTTON_HEIGHT;
        var isUp = (direction === "up");
        var chevronBox = parent.add("group");
        chevronBox.margins = 0;
        chevronBox.spacing = 0;
        chevronBox.preferredSize = [buttonWidth, buttonHeight];
        chevronBox.minimumSize = [buttonWidth, buttonHeight];
        chevronBox.maximumSize = [buttonWidth, buttonHeight];
        chevronBox.isPressed = false;
        chevronBox.isStepperButton = true; /* redrawSteppersIn() の目印 / marker for redrawSteppersIn() */

        chevronBox.onDraw = function () {
            var boxGraphics = chevronBox.graphics;
            /* 自作描画は自動でディムにならないため、無効なら薄い色で描く。親の無効化は子の enabled に出ないので親も見る
               Custom drawing is not dimmed automatically; the parent's state does not reach the child's enabled */
            var isDimmed = !isStepperEnabledInTree(chevronBox);

            /* 枠線の内側の地（押下中は押下色） / background inside the frame, pressed color while pressed */
            var fillColor = isDimmed ? STEPPER_DIM_FILL_COLOR : (chevronBox.isPressed ? STEPPER_PRESSED_COLOR : STEPPER_FILL_COLOR);
            boxGraphics.newPath();
            boxGraphics.rectPath(1, isUp ? 1 : 0, buttonWidth - 2, buttonHeight - 1);
            boxGraphics.fillPath(boxGraphics.newBrush(boxGraphics.BrushType.SOLID_COLOR, fillColor));

            drawStepperFrame(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_FRAME_COLOR : STEPPER_FRAME_COLOR);
            drawStepperChevron(boxGraphics, buttonWidth, buttonHeight, isUp, isDimmed ? STEPPER_DIM_CHEVRON_COLOR : STEPPER_CHEVRON_COLOR);
        };

        /**
         * 押下状態を変えて描き直す
         * @param {boolean} isPressed - 押下中なら true
         * @returns {void}
         */
        function repaint(isPressed) {
            if (chevronBox.isPressed === isPressed) return;
            chevronBox.isPressed = isPressed;
            redrawStepperGroup(chevronBox);
        }
        chevronBox.addEventListener("mousedown", function () {
            if (!isStepperEnabledInTree(chevronBox)) return;
            repaint(true);
            if (onClickFn) onClickFn();
        });
        chevronBox.addEventListener("mouseup", function () { repaint(false); });
        /* 押したまま外へ出たときも押下色を残さない / reset when the pointer leaves while pressed */
        chevronBox.addEventListener("mouseout", function () { repaint(false); });
        return chevronBox;
    }

    /**
     * 外側の辺だけの枠を描く（角は丸める）。継ぎ目側は開けておき、上下2つで1つの枠に見せる。
     * ScriptUI は円弧を描けないため、角丸は短い線分で近似する
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - 上のボタンなら true（上側に枠を描く）
     * @param {number[]} frameColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperFrame(boxGraphics, boxWidth, boxHeight, isUp, frameColor) {
        var frameLeft = 0.5;
        var frameRight = boxWidth - 0.5;
        var outerY = isUp ? 0.5 : boxHeight - 0.5;
        var seamY = isUp ? boxHeight : 0;
        var towardSeam = isUp ? 1 : -1; /* 外側の辺から継ぎ目へ向かう向き / direction from the outer edge to the seam */
        var radius = STEPPER_CORNER_RADIUS;
        var arcSteps = 4; /* 角丸1つを何本の線分で近似するか / segments per corner */
        var angle, k;

        boxGraphics.newPath();
        boxGraphics.moveTo(frameLeft, seamY);
        /* 左の角丸 / left corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameLeft + radius - radius * Math.cos(angle), outerY + towardSeam * (radius - radius * Math.sin(angle)));
        }
        /* 右の角丸 / right corner */
        for (k = 0; k <= arcSteps; k++) {
            angle = (Math.PI / 2) * k / arcSteps;
            boxGraphics.lineTo(frameRight - radius + radius * Math.sin(angle), outerY + towardSeam * (radius - radius * Math.cos(angle)));
        }
        boxGraphics.lineTo(frameRight, seamY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, frameColor, 1));
    }

    /**
     * 山形（∧／∨）を描く。文字グリフの▲▼は上下で大きさやベースラインが揃わないため、線で描く
     * @param {ScriptUIGraphics} boxGraphics - 描画先
     * @param {number} boxWidth - ボタンの幅
     * @param {number} boxHeight - ボタンの高さ
     * @param {boolean} isUp - ∧なら true、∨なら false
     * @param {number[]} chevronColor - [r, g, b, a]
     * @returns {void}
     */
    function drawStepperChevron(boxGraphics, boxWidth, boxHeight, isUp, chevronColor) {
        var centerX = boxWidth / 2;
        var centerY = isUp ? boxHeight / 2 + 0.5 : boxHeight / 2 - 0.5; /* 継ぎ目から少し離す / nudged away from the seam */
        var halfWidth = 3.6; /* 山形の半幅（高さ1.8に対して開き約127°） / half width of the chevron */
        var tipOffsetY = isUp ? -1.8 : 1.8; /* 頂点の中心からのずれ（上向きは上、下向きは下） */
        boxGraphics.newPath();
        boxGraphics.moveTo(centerX - halfWidth, centerY - tipOffsetY);
        boxGraphics.lineTo(centerX, centerY + tipOffsetY);
        boxGraphics.lineTo(centerX + halfWidth, centerY - tipOffsetY);
        boxGraphics.strokePath(boxGraphics.newPen(boxGraphics.PenType.SOLID_COLOR, chevronColor, 1.2));
    }

    /**
     * コントロールと、その親をたどってすべて有効かを返す（親の無効化は子の enabled に出ない）
     * @param {Object} control - 対象のコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isStepperEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * コンテナ以下にある∧∨ボタンをすべて描き直す。行やパネルの enabled を切り替えたあとに呼ぶ
     * @param {Object} container - 行・グループ・パネルなど
     * @returns {void}
     */
    function redrawSteppersIn(container) {
        if (!container.children) return;
        for (var i = 0; i < container.children.length; i++) {
            var child = container.children[i];
            if (child.isStepperButton) redrawStepperGroup(child);
            else redrawSteppersIn(child);
        }
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} targetGroup - 描き直す group
     * @returns {void}
     */
    function redrawStepperGroup(targetGroup) {
        targetGroup.hide();
        targetGroup.show();
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 入力欄の表示値と内部pt値を一元的に解決する（裁ち落としプリセットを含む）
     * @param {string} offsetText - 入力欄の現在のテキスト
     * @param {number} unitCode - rulerType の単位コード
     * @param {boolean} bleedEnabled - 裁ち落としがONかどうか
     * @returns {object} { pt: number, displayText: string, disabled: boolean }
     */
    function resolveOffsetToPt(offsetText, unitCode, bleedEnabled) {
        var displayText = String(offsetText == null ? '' : offsetText);

        if (bleedEnabled) {
            /* 単位ごとの裁ち落とし相当値。表示値と pt 値を必ず同じ量にする
               Bleed preset per unit; the shown value and the pt value always describe the same amount */
            var bleedAmount = 3;      /* 既定は 3mm 相当 / defaults to 3mm */
            var bleedUnitCode = 1;
            if (unitCode === 5) {          /* Q/H */
                bleedAmount = 12;
                bleedUnitCode = 5;
            } else if (unitCode === 2) {   /* pt（0.125in = 9pt）*/
                bleedAmount = 9;
                bleedUnitCode = 2;
            } else if (unitCode === 1) {   /* mm */
                bleedAmount = 3;
                bleedUnitCode = 1;
            } else {
                /* mm・Q/H・pt 以外は 3mm 相当を現在の単位へ換算して表示
                   For other units, convert the 3mm equivalent into the current unit */
                bleedAmount = 3 * UNITS[1].pointsPerUnit / (UNITS[unitCode] ? UNITS[unitCode].pointsPerUnit : 1);
                bleedAmount = Math.round(bleedAmount * 1000) / 1000;
                bleedUnitCode = unitCode;
            }
            return {
                pt: bleedAmount * (UNITS[bleedUnitCode] ? UNITS[bleedUnitCode].pointsPerUnit : 1),
                displayText: String(bleedAmount),
                disabled: true
            };
        }

        /* 通常時は現在の単位の係数を掛ける / Normal case: multiply by the current unit factor */
        var offsetValue = parseFloat(displayText);
        if (isNaN(offsetValue)) offsetValue = 0;
        return {
            pt: offsetValue * (UNITS[unitCode] ? UNITS[unitCode].pointsPerUnit : 1),
            displayText: displayText,
            disabled: false
        };
    }

    // =========================================
    // カラー / Color
    // =========================================

    /* 色名テーブル（RGB/CMYK 両方を持つ）/ Named colors, with both an RGB and a CMYK value */
    var NAMED_COLOR_TABLE = {
        black: { rgb: [0, 0, 0], cmyk: [0, 0, 0, 100] },
        white: { rgb: [255, 255, 255], cmyk: [0, 0, 0, 0] },
        red: { rgb: [255, 0, 0], cmyk: [0, 100, 100, 0] },
        green: { rgb: [0, 128, 0], cmyk: [100, 0, 100, 50] },
        blue: { rgb: [0, 0, 255], cmyk: [100, 100, 0, 0] },
        cyan: { rgb: [0, 255, 255], cmyk: [100, 0, 0, 0] },
        magenta: { rgb: [255, 0, 255], cmyk: [0, 100, 0, 0] },
        yellow: { rgb: [255, 255, 0], cmyk: [0, 0, 100, 0] },
        orange: { rgb: [255, 165, 0], cmyk: [0, 35, 100, 0] }
    };

    /**
     * 数値を指定範囲に収める
     * @param {number} value - 対象の値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 範囲内に収めた値
     */
    function clampValue(value, minValue, maxValue) {
        return value < minValue ? minValue : (value > maxValue ? maxValue : value);
    }

    /**
     * RGBColor を生成する（0–255にクランプ）
     * @param {number} red - 赤（0–255）
     * @param {number} green - 緑（0–255）
     * @param {number} blue - 青（0–255）
     * @returns {RGBColor} 生成した色
     */
    function makeRgbColor(red, green, blue) {
        var rgbColor = new RGBColor();
        rgbColor.red = clampValue(Math.round(red), 0, 255);
        rgbColor.green = clampValue(Math.round(green), 0, 255);
        rgbColor.blue = clampValue(Math.round(blue), 0, 255);
        return rgbColor;
    }

    /**
     * CMYKColor を生成する（0–100にクランプ）
     * @param {number} cyan - シアン（0–100）
     * @param {number} magenta - マゼンタ（0–100）
     * @param {number} yellow - イエロー（0–100）
     * @param {number} black - ブラック（0–100）
     * @returns {CMYKColor} 生成した色
     */
    function makeCmykColor(cyan, magenta, yellow, black) {
        var cmykColor = new CMYKColor();
        cmykColor.cyan = clampValue(cyan, 0, 100);
        cmykColor.magenta = clampValue(magenta, 0, 100);
        cmykColor.yellow = clampValue(yellow, 0, 100);
        cmykColor.black = clampValue(black, 0, 100);
        return cmykColor;
    }

    /**
     * CMYK値をRGB値へ変換する
     * @param {number} cyan - シアン（0–100）
     * @param {number} magenta - マゼンタ（0–100）
     * @param {number} yellow - イエロー（0–100）
     * @param {number} black - ブラック（0–100）
     * @returns {number[]} [R, G, B]（0–255）
     */
    function cmykToRgb(cyan, magenta, yellow, black) {
        var cyanRatio = clampValue(cyan, 0, 100) / 100;
        var magentaRatio = clampValue(magenta, 0, 100) / 100;
        var yellowRatio = clampValue(yellow, 0, 100) / 100;
        var blackRatio = clampValue(black, 0, 100) / 100;
        return [
            Math.round(255 * (1 - cyanRatio) * (1 - blackRatio)),
            Math.round(255 * (1 - magentaRatio) * (1 - blackRatio)),
            Math.round(255 * (1 - yellowRatio) * (1 - blackRatio))
        ];
    }

    /**
     * ドキュメントのカラースペースに合わせた黒を生成する
     * @param {Document} doc - 対象ドキュメント
     * @returns {RGBColor|CMYKColor} 黒
     */
    function createBlackColor(doc) {
        if (doc.documentColorSpace == DocumentColorSpace.RGB) return makeRgbColor(0, 0, 0);
        return makeCmykColor(0, 0, 0, 100);
    }

    /**
     * カラー入力欄の文字列を色として解釈する
     * #RRGGBB／#RGB・#RG・#R の短縮形／色名（red など）／grayNN（0–100）を受け付ける
     * @param {Document} doc - 対象ドキュメント（カラースペース判定に使う）
     * @param {string} colorText - 入力文字列
     * @returns {RGBColor|CMYKColor|null} 解釈できた色。できなければ null
     */
    function parseColorText(doc, colorText) {
        if (!colorText) return null;
        var normalizedText = String(colorText).replace(/^\s+|\s+$/g, '').toLowerCase();
        if (!normalizedText) return null;

        /* 全角の空白・読点・数字・記号をASCIIへ正規化 / Normalize full-width characters to ASCII */
        normalizedText = normalizedText.replace(/　/g, ' ').replace(/[，、]/g, ',');
        normalizedText = normalizedText.replace(/[０-９]/g, function (fullWidthDigit) {
            return String.fromCharCode(fullWidthDigit.charCodeAt(0) - 0xFF10 + 0x30);
        });
        normalizedText = normalizedText.replace(/．/g, '.').replace(/／/g, '/');

        /* 短縮HEXを #RRGGBB へ展開 / Expand shorthand hex notations */
        if (normalizedText.charAt(0) === '#') {
            var digits = normalizedText.substr(1);
            if (digits.length === 1) {          /* #R → #RRRRRR */
                normalizedText = '#' + digits + digits + digits + digits + digits + digits;
            } else if (digits.length === 2) {   /* #RG → #RGRGRG */
                normalizedText = '#' + digits + digits + digits;
            } else if (digits.length === 3) {   /* #RGB → #RRGGBB */
                normalizedText = '#' + digits.charAt(0) + digits.charAt(0) +
                    digits.charAt(1) + digits.charAt(1) +
                    digits.charAt(2) + digits.charAt(2);
            }
        }

        /* #RRGGBB */
        if (/^#[0-9a-f]{6}$/.test(normalizedText)) {
            return makeRgbColor(parseInt(normalizedText.substr(1, 2), 16), parseInt(normalizedText.substr(3, 2), 16), parseInt(normalizedText.substr(5, 2), 16));
        }

        /* 色名 / Named colors — ドキュメントのカラースペースを優先 */
        var namedColor = NAMED_COLOR_TABLE[normalizedText];
        if (namedColor) {
            if (doc && doc.documentColorSpace == DocumentColorSpace.CMYK) {
                return makeCmykColor(namedColor.cmyk[0], namedColor.cmyk[1], namedColor.cmyk[2], namedColor.cmyk[3]);
            }
            return makeRgbColor(namedColor.rgb[0], namedColor.rgb[1], namedColor.rgb[2]);
        }

        /* grayNN（0–100）/ grayNN (0-100) */
        var grayMatch = normalizedText.match(/^gray\s*(\d{1,3})$/);
        if (grayMatch) {
            var grayLevel = clampValue(parseInt(grayMatch[1], 10), 0, 100);
            if (doc && doc.documentColorSpace == DocumentColorSpace.CMYK) return makeCmykColor(0, 0, 0, grayLevel);
            var grayByte = Math.round(255 * (100 - grayLevel) / 100);
            return makeRgbColor(grayByte, grayByte, grayByte);
        }

        return null;
    }

    /**
     * CMYK入力値からドキュメントのカラースペースに合う色を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {object} cmykValues - { c, m, y, k }（0–100）
     * @returns {RGBColor|CMYKColor|null} 4値が揃っていなければ null
     */
    function buildCmykFillColor(doc, cmykValues) {
        if (!cmykValues) return null;
        var channels = [cmykValues.c, cmykValues.m, cmykValues.y, cmykValues.k];
        for (var i = 0; i < channels.length; i++) {
            if (typeof channels[i] !== 'number' || isNaN(channels[i])) return null;
        }
        if (doc && doc.documentColorSpace == DocumentColorSpace.RGB) {
            var rgbValues = cmykToRgb(channels[0], channels[1], channels[2], channels[3]);
            return makeRgbColor(rgbValues[0], rgbValues[1], rgbValues[2]);
        }
        return makeCmykColor(channels[0], channels[1], channels[2], channels[3]);
    }

    /**
     * カラーモードに応じた塗りを適用する（プレビューと本描画で共通）
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} targetRectangle - 塗りを適用する長方形
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function applyFillByMode(doc, targetRectangle, drawSettings) {
        var fillColor = null;
        var fillOpacity = 100;

        if (drawSettings.colorMode === ColorMode.K100) {
            fillColor = createBlackColor(doc);
            fillOpacity = K100_OPACITY;
        } else if (drawSettings.colorMode === ColorMode.HEX) {
            fillColor = parseColorText(doc, drawSettings.customValue);
        } else if (drawSettings.colorMode === ColorMode.CMYK) {
            fillColor = buildCmykFillColor(doc, drawSettings.customCMYK);
        }

        /* 解釈できない値・「なし」は塗りなし。線は呼び出し側（プレビュー）で付け直す
           Unparsable values and "None" mean no fill; the caller re-applies any stroke */
        targetRectangle.stroked = false;
        targetRectangle.filled = !!fillColor;
        if (fillColor) targetRectangle.fillColor = fillColor;
        targetRectangle.opacity = fillOpacity;
    }

    // =========================================
    // 長方形とレイヤーの共通処理 / Rectangle and layer helpers
    // =========================================

    /**
     * 座標系をドキュメント座標へ切り替える
     * @returns {CoordinateSystem|null} 切り替え前の座標系（取得できなければ null）
     */
    function switchToDocumentCoordinates() {
        var previousCoordinateSystem = null;
        try {
            previousCoordinateSystem = app.coordinateSystem;
            app.coordinateSystem = CoordinateSystem.DOCUMENTCOORDINATESYSTEM;
        } catch (e) { }
        return previousCoordinateSystem;
    }

    /**
     * switchToDocumentCoordinates() で控えた座標系へ戻す
     * @param {CoordinateSystem|null} previousCoordinateSystem - 切り替え前の座標系
     * @returns {void}
     */
    function restoreCoordinateSystem(previousCoordinateSystem) {
        try {
            if (previousCoordinateSystem !== null) app.coordinateSystem = previousCoordinateSystem;
        } catch (e) { }
    }

    /**
     * アートボードをオフセットぶん広げた長方形の位置と寸法を求める
     * @param {number[]} artboardRect - アートボードの [left, top, right, bottom]
     * @param {number} offsetPt - 外側へ広げる量（pt、負の値で内側）
     * @returns {{top: number, left: number, width: number, height: number}} 長方形の位置と寸法
     */
    function getOffsetRectangleBounds(artboardRect, offsetPt) {
        return {
            top: artboardRect[1] + offsetPt,
            left: artboardRect[0] - offsetPt,
            width: (artboardRect[2] - artboardRect[0]) + offsetPt * 2,
            height: (artboardRect[1] - artboardRect[3]) + offsetPt * 2
        };
    }

    /**
     * 配置位置の設定どおりにレイヤー内の重ね順を変える（bgレイヤーは作成順のまま）
     * @param {PathItem} targetRectangle - 対象の長方形
     * @param {string} zOrder - "front" / "back" / "bg"
     * @returns {void}
     */
    function applyZOrder(targetRectangle, zOrder) {
        if (zOrder === 'front') targetRectangle.zOrder(ZOrderMethod.BRINGTOFRONT);
        else if (zOrder === 'back') targetRectangle.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 名前が一致する最上位レイヤーを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー。無ければ null
     */
    function findLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) return doc.layers[i];
        }
        return null;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* =========================================
     * PreviewHistory util (extractable)
     * ヒストリーを残さないプレビューのための小さなユーティリティ。
     * 使い方:
     *   PreviewHistory.start();      // ダイアログ表示時などにカウンタ初期化
     *   PreviewHistory.bump();       // プレビュー描画ごとにカウント(+1)
     *   PreviewHistory.undo();       // 閉じる/キャンセル時に一括Undo
     *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
     * ========================================= */
    (function (globalObject) {
        if (!globalObject.PreviewHistory) {
            globalObject.PreviewHistory = {
                start: function () {
                    globalObject.__previewUndoCount = 0;
                },
                bump: function () {
                    globalObject.__previewUndoCount = (globalObject.__previewUndoCount | 0) + 1;
                },
                undo: function () {
                    var undoCount = globalObject.__previewUndoCount | 0;
                    try {
                        for (var i = 0; i < undoCount; i++) app.executeMenuCommand('undo');
                    } catch (e) { }
                    globalObject.__previewUndoCount = 0;
                },
                cancelTask: function (taskId) {
                    try { if (taskId) app.cancelTask(taskId); } catch (e) { }
                }
            };
        }
    })($.global);

    /* デバウンス中のプレビュータスクID / Task id of the pending debounced preview */
    var previewDebounceTaskId = null;

    /**
     * プレビュー描画を遅延スケジュールする（デバウンス）
     * @param {object} drawSettings - ダイアログの設定値
     * @param {number} delayMs - 遅延ミリ秒
     * @returns {void}
     */
    function schedulePreview(drawSettings, delayMs) {
        PreviewHistory.cancelTask(previewDebounceTaskId);
        /* scheduleTask の文字列はグローバルスコープで評価されるため、IIFE内の関数を $.global 経由で渡す
           scheduleTask runs its string in global scope, so expose the renderer via $.global */
        $.global.__previewSettings = drawSettings;
        $.global.__previewRenderer = renderPreview;
        var scheduledCode = 'try{$.global.__previewRenderer(app.activeDocument, $.global.__previewSettings);}catch(e){}';
        try {
            previewDebounceTaskId = app.scheduleTask(scheduledCode, Math.max(0, delayMs | 0), false);
        } catch (e) {
            try { renderPreview(app.activeDocument, drawSettings); } catch (err) { }
        }
    }

    /**
     * プレビュー専用レイヤーの名前かどうかを判定する
     * @param {string} layerName - レイヤー名
     * @returns {boolean} プレビュー専用レイヤーなら true
     */
    function isPreviewLayerName(layerName) {
        return layerName === getLabel('objectName.previewLayer') || layerName === '_preview';
    }

    /**
     * プレビューを片付ける
     * @param {boolean} removeLayer - true でレイヤーごと削除、false ではプレビュー長方形を隠すだけ
     * @returns {void}
     */
    function clearPreview(removeLayer) {
        try {
            var doc = app.activeDocument;
            var previewItemPrefix = getLabel('objectName.previewRect') + "#";
            for (var i = doc.layers.length - 1; i >= 0; i--) {
                var layer = doc.layers[i];
                if (!isPreviewLayerName(layer.name)) continue;
                if (removeLayer) {
                    layer.remove();
                    continue;
                }
                /* 入力中は削除せず、このスクリプトが作った長方形だけ隠す
                   While typing, hide only the items this script created instead of deleting them */
                for (var k = layer.pathItems.length - 1; k >= 0; k--) {
                    if (String(layer.pathItems[k].name || "").indexOf(previewItemPrefix) === 0) {
                        layer.pathItems[k].hidden = true;
                    }
                }
            }
        } catch (e) { }
    }

    /**
     * プレビュー専用レイヤーを取得する（なければ作成し最前面へ）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} プレビュー専用レイヤー
     */
    function getOrCreatePreviewLayer(doc) {
        var previewLayerName = getLabel('objectName.previewLayer');
        var previewLayer = findLayerByName(doc, previewLayerName);
        if (!previewLayer) {
            previewLayer = doc.layers.add();
            previewLayer.name = previewLayerName;
        }
        previewLayer.visible = true;
        previewLayer.locked = false;
        try {
            previewLayer.move(doc, ElementPlacement.PLACEATBEGINNING);
        } catch (e) { }
        return previewLayer;
    }

    /**
     * アートボード番号に対応するプレビュー長方形を取得する（なければ作成、あれば再利用）
     * @param {Layer} previewLayer - プレビュー専用レイヤー
     * @param {number} artboardIndex - アートボード番号
     * @param {{top: number, left: number, width: number, height: number}} rectangleBounds - 長方形の位置と寸法
     * @returns {PathItem} プレビュー長方形
     */
    function getOrCreatePreviewRectangle(previewLayer, artboardIndex, rectangleBounds) {
        var previewItemName = getLabel('objectName.previewRect') + "#" + artboardIndex;
        for (var i = 0; i < previewLayer.pathItems.length; i++) {
            var existingRectangle = previewLayer.pathItems[i];
            if (existingRectangle.name !== previewItemName) continue;
            /* 既存を使い回して再作成のヒストリーを増やさない / Reuse in place to avoid extra history entries */
            existingRectangle.top = rectangleBounds.top;
            existingRectangle.left = rectangleBounds.left;
            existingRectangle.width = rectangleBounds.width;
            existingRectangle.height = rectangleBounds.height;
            existingRectangle.hidden = false;
            return existingRectangle;
        }
        var previewRectangle = previewLayer.pathItems.rectangle(
            rectangleBounds.top, rectangleBounds.left, rectangleBounds.width, rectangleBounds.height);
        previewRectangle.name = previewItemName;
        return previewRectangle;
    }

    /**
     * プレビュー用の50%グレー（ドキュメントのカラースペースに合わせる）
     * @param {Document} doc - 対象ドキュメント
     * @returns {RGBColor|CMYKColor} 線色
     */
    function getPreviewStrokeColor(doc) {
        if (doc.documentColorSpace == DocumentColorSpace.RGB) return makeRgbColor(128, 128, 128);
        return makeCmykColor(0, 0, 0, 50);
    }

    /**
     * 1つのアートボードぶんのプレビュー長方形を描く
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} previewLayer - プレビュー専用レイヤー
     * @param {number} artboardIndex - アートボード番号
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function drawPreviewRectangle(doc, previewLayer, artboardIndex, drawSettings) {
        var rectangleBounds = getOffsetRectangleBounds(doc.artboards[artboardIndex].artboardRect, drawSettings.offset || 0);
        var previewRectangle = getOrCreatePreviewRectangle(previewLayer, artboardIndex, rectangleBounds);

        applyFillByMode(doc, previewRectangle, drawSettings);

        /* 塗りなしでも位置が分かるよう、プレビューは常に破線で縁取る
           Always outline the preview so it stays visible even with no fill */
        previewRectangle.stroked = true;
        previewRectangle.strokeWidth = 1;
        previewRectangle.strokeDashes = [6, 4];
        previewRectangle.strokeColor = getPreviewStrokeColor(doc);
        previewRectangle.selected = false;

        applyZOrder(previewRectangle, drawSettings.zOrder);
    }

    /**
     * 対象外のアートボードに対応するプレビュー長方形を隠す
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} previewLayer - プレビュー専用レイヤー
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function hidePreviewItemsOutOfScope(doc, previewLayer, drawSettings) {
        /* すべてのアートボードが対象のときは -1（範囲外の番号だけ隠す）/ -1 means "every artboard is in scope" */
        var visibleArtboardIndex = (drawSettings.target === 'all') ? -1 : doc.artboards.getActiveArtboardIndex();
        for (var i = 0; i < previewLayer.pathItems.length; i++) {
            var previewItem = previewLayer.pathItems[i];
            /* アイテム名は "__Preview_...#<アートボード番号>" 形式 */
            var indexMatch = /#(\d+)$/.exec(previewItem.name || "");
            if (!indexMatch) continue;
            var itemArtboardIndex = parseInt(indexMatch[1], 10);
            var keepVisible = (visibleArtboardIndex < 0) ?
                (itemArtboardIndex < doc.artboards.length) :
                (itemArtboardIndex === visibleArtboardIndex);
            if (!keepVisible) previewItem.hidden = true;
        }
    }

    /**
     * プレビューを描画する（専用レイヤーへ一時オブジェクトを生成）
     * @param {Document} doc - 対象ドキュメント
     * @param {object} drawSettings - ダイアログの設定値
     * @returns {void}
     */
    function renderPreview(doc, drawSettings) {
        /* レイヤーごと消さずに既存プレビューを隠す / Hide existing preview items instead of deleting the layer */
        clearPreview(false);
        if (!doc || !drawSettings) return;

        var previousCoordinateSystem = switchToDocumentCoordinates();
        var previewLayer = getOrCreatePreviewLayer(doc);
        if (drawSettings.target === 'all') {
            for (var i = 0; i < doc.artboards.length; i++) drawPreviewRectangle(doc, previewLayer, i, drawSettings);
        } else {
            drawPreviewRectangle(doc, previewLayer, doc.artboards.getActiveArtboardIndex(), drawSettings);
        }
        hidePreviewItemsOutOfScope(doc, previewLayer, drawSettings);
        restoreCoordinateSystem(previousCoordinateSystem);

        PreviewHistory.bump();
        app.redraw();
    }

    // =========================================
    // 入力欄の検証 / Field validation
    // =========================================

    /**
     * HEX入力欄の警告表示を切り替える
     * @param {EditText} hexField - HEX入力欄
     * @param {boolean} isWarning - true で警告表示
     * @param {string} [warningKey] - 警告時のヘルプチップのラベルキー
     * @returns {void}
     */
    function setHexWarning(hexField, isWarning, warningKey) {
        setFieldWarnColor(hexField, isWarning);
        hexField.helpTip = isWarning ? getLabel(warningKey || 'warning.hexInvalid') : getLabel('tooltip.hexInput');
        try { hexField.notify('onDraw'); } catch (e) { }
    }

    /**
     * CMYK入力欄の警告表示を切り替える
     * @param {EditText} channelInput - CMYK各チャンネルの入力欄
     * @param {boolean} isWarning - true で警告表示
     * @returns {void}
     */
    function setCmykWarning(channelInput, isWarning) {
        setFieldWarnColor(channelInput, isWarning);
        channelInput.helpTip = isWarning ? getLabel('warning.cmykRange') : getLabel('tooltip.cmykInput');
    }

    /**
     * CMYK入力欄の値が0–100の範囲かを判定して警告表示に反映する（入力途中は警告しない）
     * @param {EditText} channelInput - CMYK各チャンネルの入力欄
     * @returns {void}
     */
    function validateCmykField(channelInput) {
        var fieldText = String(channelInput.text || '');
        if (fieldText === '') {
            setCmykWarning(channelInput, false);
            return;
        }
        var channelValue = parseFloat(fieldText);
        setCmykWarning(channelInput, isNaN(channelValue) || channelValue < 0 || channelValue > 100);
    }

    /**
     * CMYK入力欄の値を0–100へ丸める（未入力は空のまま。計算時に0として扱う）
     * @param {EditText} channelInput - CMYK各チャンネルの入力欄
     * @returns {void}
     */
    function clampCmykField(channelInput) {
        var fieldText = String(channelInput.text || '').replace(/^\s+|\s+$/g, '');
        if (fieldText !== '') {
            var channelValue = parseFloat(fieldText);
            if (isNaN(channelValue)) channelValue = 0;
            channelInput.text = String(clampValue(channelValue, 0, 100));
        }
        setCmykWarning(channelInput, false);
    }

    /**
     * CMYK入力欄へ共通のハンドラをまとめて登録する
     * @param {EditText} channelInput - CMYK各チャンネルの入力欄（∧∨付き）
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred }
     * @returns {void}
     */
    function bindCmykField(channelInput, previewHooks) {
        channelInput.addEventListener('focus', function () {
            /* ちょうど "0" のときは入力しやすいようクリア / Clear a lone "0" so typing replaces it */
            if (String(channelInput.text) === '0') channelInput.text = '';
        });

        channelInput.addEventListener('keydown', function (event) {
            /* 先頭ゼロ（"03"）を作らせない。小数 "0.5" は触らない
               Prevent leading-zero integers; leave decimals like "0.5" alone */
            var typedKey = String(event.keyName || '');
            if (!/^[0-9]$/.test(typedKey)) return;
            var fieldText = String(channelInput.text || '');
            if (/\./.test(fieldText)) return;
            if (/^0+$/.test(fieldText)) channelInput.text = '';
            else if (/^0\d+$/.test(fieldText)) channelInput.text = fieldText.replace(/^0+/, '');
        });

        channelInput.onChanging = function () {
            var fieldText = String(channelInput.text || '');
            if (/^0\d+$/.test(fieldText)) channelInput.text = fieldText.replace(/^0+/, '');
            validateCmykField(channelInput);
            previewHooks.deferred();
        };

        channelInput.onChange = function () {
            clampCmykField(channelInput);
            previewHooks.immediate();
        };

        /* ∧∨と↑↓キーは同じ処理で増減する（0〜100に収める） / steppers and arrow keys share one path (0-100) */
        channelInput.stepperOptions.onStep = function () {
            clampCmykField(channelInput);
            previewHooks.deferred();
        };
        bindSteppedArrowKeys(channelInput, channelInput.stepperGroup);

        trackFocusForHotkeys(channelInput);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * オフセットパネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred }
     * @returns {object} { offsetInput, bleedCheckbox, initFieldState }
     */
    function buildOffsetPanel(parentGroup, previewHooks) {
        var offsetPanel = parentGroup.add('panel', undefined, getLabel('panel.offset'));
        setupPanel(offsetPanel);

        var offsetRow = offsetPanel.add('group');
        setupGroup(offsetRow, 'row');
        offsetRow.alignChildren = 'center';
        offsetRow.alignment = 'center';

        /* ∧∨と入力欄は隙間0で突き合わせる。負の値（内側へ縮小）も可 / butt the stepper against the field; negatives allowed */
        var offsetFieldGroup = offsetRow.add('group');
        offsetFieldGroup.orientation = 'row';
        offsetFieldGroup.alignChildren = ['left', 'center'];
        offsetFieldGroup.spacing = 0;
        offsetFieldGroup.margins = 0;
        var offsetInput;
        var offsetStepper = addStepper(offsetFieldGroup, function () { return offsetInput; }, {
            onStep: function () { previewHooks.deferred(); }
        });
        offsetInput = offsetFieldGroup.add('edittext', undefined, '0');
        offsetInput.characters = 4;
        offsetInput.helpTip = getLabel('tooltip.offsetInput');
        offsetRow.add('statictext', undefined, getUnitInfo().label);

        var bleedRow = offsetPanel.add('group');
        setupGroup(bleedRow, 'row');
        bleedRow.alignChildren = 'center';
        bleedRow.alignment = 'center';

        var bleedCheckbox = bleedRow.add('checkbox', undefined, getLabel('checkbox.bleed'));
        bleedCheckbox.alignment = 'center';
        bleedCheckbox.value = false; /* デフォルトOFF / default OFF */
        bleedCheckbox.helpTip = getLabel('tooltip.bleed');

        /* 裁ち落としON/OFFの往復で手入力値を失わないよう控えておく
           Remember the manual offset so toggling Bleed does not lose it */
        var manualOffsetText = '0';

        /* 裁ち落としの状態を入力欄へ反映 / Reflect the current Bleed state in the field */
        function applyBleedState(refreshPreview) {
            if (bleedCheckbox.value) {
                var resolvedOffset = resolveOffsetToPt(offsetInput.text, getUnitInfo().code, true);
                offsetInput.text = resolvedOffset.displayText;
                offsetInput.enabled = !resolvedOffset.disabled;
            } else {
                offsetInput.text = manualOffsetText;
                offsetInput.enabled = true;
            }
            /* ∧∨も入力欄にそろえて切り替え、描き直す / sync and redraw the stepper */
            offsetStepper.enabled = offsetInput.enabled;
            redrawSteppersIn(offsetStepper);
            if (refreshPreview) previewHooks.deferred();
        }

        bleedCheckbox.onClick = function () {
            if (bleedCheckbox.value) manualOffsetText = String(offsetInput.text);
            applyBleedState(true);
        };

        offsetInput.onChanging = previewHooks.deferred;
        offsetInput.onChange = previewHooks.immediate;
        offsetInput.addEventListener('keydown', function (event) {
            /* Enterでも即座に反映 / Enter refreshes the preview immediately */
            if (event.keyName == 'Enter') previewHooks.immediate();
        });
        bindSteppedArrowKeys(offsetInput, offsetStepper);
        trackFocusForHotkeys(offsetInput);

        return {
            offsetInput: offsetInput,
            bleedCheckbox: bleedCheckbox,
            initFieldState: function () {
                manualOffsetText = String(offsetInput.text);
                applyBleedState(false);
            }
        };
    }

    /**
     * 配列の指定番目の入力欄を返す関数を作る（ループ内で∧∨に渡すため、添字を閉じ込める）
     * @param {EditText[]} inputList - 入力欄の配列（あとから push されてもよい）
     * @param {number} inputIndex - 添字
     * @returns {Function} 入力欄を返す関数
     */
    function makeInputGetter(inputList, inputIndex) {
        return function () { return inputList[inputIndex]; };
    }

    /**
     * カラーパネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred }
     * @returns {object} 各ラジオ・入力欄をまとめたオブジェクト
     */
    function buildColorPanel(parentGroup, previewHooks) {
        var colorPanel = parentGroup.add('panel', undefined, getLabel('panel.color'));
        setupPanel(colorPanel, STACK_SPACING); /* やや広めの行間 / a bit more vertical gap */

        var noneRadio = colorPanel.add('radiobutton', undefined, getLabel('radio.colorNone'));
        noneRadio.helpTip = getLabel('tooltip.colorNone');
        var k100Radio = colorPanel.add('radiobutton', undefined, getLabel('radio.colorK100'));

        /* HEXはラジオと入力欄を同じ行に / HEX radio and its field share one row */
        var hexRow = colorPanel.add('group');
        setupGroup(hexRow, 'row', TIGHT_SPACING);
        var hexRadio = hexRow.add('radiobutton', undefined, getLabel('radio.colorHex'));
        var hexInput = hexRow.add('edittext', undefined, '#');
        hexInput.characters = 14; /* カラム幅が伸びないよう控えめに / narrow enough to keep the column width */
        hexInput.helpTip = getLabel('tooltip.hexInput');

        var cmykRadio = colorPanel.add('radiobutton', undefined, getLabel('radio.colorCmyk'));
        cmykRadio.helpTip = getLabel('tooltip.colorCmyk');

        /* ラベル行とフィールド行の2段グリッド / Two-row grid: labels on top, fields below */
        var cmykGrid = colorPanel.add('group');
        setupGroup(cmykGrid, 'column', 4);
        var cmykLabelRow = cmykGrid.add('group');
        setupGroup(cmykLabelRow, 'row', STACK_SPACING);
        var cmykFieldRow = cmykGrid.add('group');
        setupGroup(cmykFieldRow, 'row', STACK_SPACING);

        var cmykChannelTexts = ['  C', '  M', '  Y', '  K'];
        var cmykLabels = [];
        var cmykInputs = [];
        for (var i = 0; i < cmykChannelTexts.length; i++) {
            /* 項目名は入力欄の真上に来るよう、∧∨の幅だけ左をあける / indent the label by the stepper width to sit above the field */
            var channelLabelCell = cmykLabelRow.add('group');
            channelLabelCell.margins = [STEPPER_SIDE_MARGIN + STEPPER_BUTTON_WIDTH, 0, 0, 0];
            channelLabelCell.spacing = 0;
            var channelLabel = channelLabelCell.add('statictext', undefined, cmykChannelTexts[i]);
            channelLabel.preferredSize.width = CMYK_FIELD_WIDTH;

            /* ∧∨と入力欄は隙間0で突き合わせる。onStep は bindCmykField() で入れる / onStep is set in bindCmykField() */
            var channelFieldGroup = cmykFieldRow.add('group');
            channelFieldGroup.orientation = 'row';
            channelFieldGroup.alignChildren = ['left', 'center'];
            channelFieldGroup.spacing = 0;
            channelFieldGroup.margins = 0;
            var channelStepperOptions = { min: 0, max: 100 };
            var channelStepper = addStepper(channelFieldGroup, makeInputGetter(cmykInputs, i), channelStepperOptions);
            var channelInput = channelFieldGroup.add('edittext', undefined, '');
            channelInput.characters = 3;
            channelInput.preferredSize.width = CMYK_FIELD_WIDTH;
            channelInput.helpTip = getLabel('tooltip.cmykInput');
            channelInput.stepperGroup = channelStepper;
            channelInput.stepperOptions = channelStepperOptions;
            cmykLabels.push(channelLabel);
            cmykInputs.push(channelInput);
            bindCmykField(channelInput, previewHooks);
        }

        hexInput.onChanging = function () {
            var hexText = String(hexInput.text || '').replace(/\s+/g, '');
            if (hexText === '') setHexWarning(hexInput, false);
            else if (hexText === '#') setHexWarning(hexInput, true, 'warning.hexEmpty');
            /* parseColorText が解釈できる入力（#RRGGBB／短縮HEX／色名／grayNN）はすべて有効
               Anything parseColorText can resolve is valid */
            else setHexWarning(hexInput, !parseColorText(app.activeDocument, hexText));
            previewHooks.deferred();
        };

        hexInput.onChange = function () {
            var hexText = String(hexInput.text || '').replace(/\s+/g, '');
            if (/^#?[0-9a-fA-F]{6}$/.test(hexText)) {
                /* 6桁HEXは # 付き・大文字へ正規化 / Normalize 6-digit hex to "#" + uppercase */
                hexInput.text = '#' + hexText.replace(/^#/, '').toUpperCase();
                setHexWarning(hexInput, false);
            } else if (hexText === '#') {
                setHexWarning(hexInput, true, 'warning.hexEmpty');
            } else {
                /* 色名・短縮HEX・grayNN は整形せずそのまま受理 / Names, shorthand hex and grayNN pass through */
                setHexWarning(hexInput, !parseColorText(app.activeDocument, hexText));
            }
            previewHooks.immediate();
        };

        trackFocusForHotkeys(hexInput);

        /* ラジオ選択に応じて入力欄の有効・無効を反映 / Sync field enable state with the radios */
        function updateColorFieldStates() {
            hexInput.enabled = !!hexRadio.value;
            var cmykEnabled = !!cmykRadio.value;
            for (var i = 0; i < cmykInputs.length; i++) {
                cmykInputs[i].enabled = cmykEnabled;
                cmykInputs[i].stepperGroup.enabled = cmykEnabled;
                redrawSteppersIn(cmykInputs[i].stepperGroup);
                cmykLabels[i].enabled = cmykEnabled;
                if (!cmykEnabled) setCmykWarning(cmykInputs[i], false);
            }
        }

        /* カラーモードを排他選択し、ハイライト・フォーカス・プレビューを更新
           Select a color mode exclusively, then sync highlight, focus and preview */
        function selectColorMode(colorMode) {
            noneRadio.value = (colorMode === ColorMode.NONE);
            k100Radio.value = (colorMode === ColorMode.K100);
            hexRadio.value = (colorMode === ColorMode.HEX);
            cmykRadio.value = (colorMode === ColorMode.CMYK);
            updateColorFieldStates();
            setFieldHighlight(hexInput, colorMode === ColorMode.HEX);
            setFieldHighlight(cmykInputs[0], colorMode === ColorMode.CMYK);
            try {
                if (colorMode === ColorMode.HEX) hexInput.active = true;
                else if (colorMode === ColorMode.CMYK) cmykInputs[0].active = true;
            } catch (e) { }
            previewHooks.immediate();
        }

        var colorRadioModes = [
            [noneRadio, ColorMode.NONE],
            [k100Radio, ColorMode.K100],
            [hexRadio, ColorMode.HEX],
            [cmykRadio, ColorMode.CMYK]
        ];
        for (var j = 0; j < colorRadioModes.length; j++) {
            (function (radio, colorMode) {
                radio.onClick = radio.onChanging = function () { selectColorMode(colorMode); };
            })(colorRadioModes[j][0], colorRadioModes[j][1]);
        }

        k100Radio.value = true; /* デフォルトはK100 / default to K100 */
        updateColorFieldStates();

        return {
            noneRadio: noneRadio,
            k100Radio: k100Radio,
            hexRadio: hexRadio,
            cmykRadio: cmykRadio,
            hexInput: hexInput,
            cmykInputs: cmykInputs
        };
    }

    /**
     * 配置位置（重ね順）パネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred }
     * @returns {object} { frontRadio, backRadio, bgLayerRadio }
     */
    function buildPlacementPanel(parentGroup, previewHooks) {
        var placementPanel = parentGroup.add('panel', undefined, getLabel('panel.placement'));
        setupPanel(placementPanel, TIGHT_SPACING);

        var frontRadio = placementPanel.add('radiobutton', undefined, getLabel('radio.placeFront'));
        var backRadio = placementPanel.add('radiobutton', undefined, getLabel('radio.placeBack'));
        var bgLayerRadio = placementPanel.add('radiobutton', undefined, getLabel('radio.placeBgLayer'));
        frontRadio.helpTip = getLabel('tooltip.placeFront');
        backRadio.helpTip = getLabel('tooltip.placeBack');
        bgLayerRadio.helpTip = getLabel('tooltip.bgLayer');

        frontRadio.value = true; /* デフォルトは最前面 / default to Bring to Front */

        var placementRadios = [frontRadio, backRadio, bgLayerRadio];
        for (var i = 0; i < placementRadios.length; i++) {
            placementRadios[i].onClick = previewHooks.immediate;
        }

        return { frontRadio: frontRadio, backRadio: backRadio, bgLayerRadio: bgLayerRadio };
    }

    /**
     * 対象アートボードのパネルを構築する
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @param {object} previewHooks - プレビュー更新コールバック { immediate, deferred }
     * @returns {object} { currentArtboardRadio, allArtboardsRadio }
     */
    function buildTargetPanel(parentGroup, previewHooks) {
        var targetPanel = parentGroup.add('panel', undefined, getLabel('panel.target'));
        setupPanel(targetPanel);

        var currentArtboardRadio = targetPanel.add('radiobutton', undefined, getLabel('radio.currentArtboard'));
        var allArtboardsRadio = targetPanel.add('radiobutton', undefined, getLabel('radio.allArtboards'));

        /* 常に「現在のアートボード」をデフォルト選択 / Always default to the current artboard */
        currentArtboardRadio.value = true;
        allArtboardsRadio.value = false;

        /* 1枚しかない場合は「すべてのアートボード」をディム / Dim "All Artboards" when there is only one */
        var artboardCount = app.documents.length ? app.activeDocument.artboards.length : 0;
        if (artboardCount <= 1) {
            allArtboardsRadio.enabled = false;
            allArtboardsRadio.helpTip = getLabel('warning.singleArtboard');
        }

        /* クリックだけでなくキーボード操作（onChanging）でも更新 / Refresh on click and on keyboard change */
        currentArtboardRadio.onClick = currentArtboardRadio.onChanging = previewHooks.immediate;
        allArtboardsRadio.onClick = allArtboardsRadio.onChanging = previewHooks.immediate;

        return { currentArtboardRadio: currentArtboardRadio, allArtboardsRadio: allArtboardsRadio };
    }

    /**
     * オプションパネル（ガイド化／ライブシェイプ変換）を構築する
     * 各オプションは独立。必要なものだけ描画後に適用（applyDrawOptions）
     * @param {Group} parentGroup - 追加先のカラムグループ
     * @returns {object} { makeGuideCheckbox, convertToLiveShapeCheckbox }
     */
    function buildOptionsPanel(parentGroup) {
        var optionsPanel = parentGroup.add('panel', undefined, getLabel('panel.options'));
        setupPanel(optionsPanel);

        var makeGuideCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.makeGuide'));
        makeGuideCheckbox.value = false; /* デフォルトOFF / default OFF */
        makeGuideCheckbox.helpTip = getLabel('tooltip.makeGuide');

        var convertToLiveShapeCheckbox = optionsPanel.add('checkbox', undefined, getLabel('checkbox.convertToLiveShape'));
        convertToLiveShapeCheckbox.value = true; /* デフォルトON / default ON */
        convertToLiveShapeCheckbox.helpTip = getLabel('tooltip.convertToLiveShape');

        return { makeGuideCheckbox: makeGuideCheckbox, convertToLiveShapeCheckbox: convertToLiveShapeCheckbox };
    }

    /**
     * ダイアログのホットキーを登録する（F/B/L=重ね順、C/A=対象、G=ガイド化）
     * @param {Window} settingsDialog - 対象ダイアログ
     * @param {object} dialogControls - 各パネルのコントロール
     * @param {function} refreshPreview - プレビューを即時更新するコールバック
     * @returns {void}
     */
    function addDialogHotkeys(settingsDialog, dialogControls, refreshPreview) {
        settingsDialog.addEventListener('keydown', function (event) {
            if (focusedField) return; /* 入力中は無効 / ignore while typing in a field */
            var pressedKey = (event && event.keyName) ? String(event.keyName).toUpperCase() : '';

            if (pressedKey === 'G') {
                /* ガイド化は描画後の処理なのでプレビューには反映しない
                   Make-guides is a post-draw option and is not previewed */
                var makeGuideCheckbox = dialogControls.options.makeGuideCheckbox;
                makeGuideCheckbox.value = !makeGuideCheckbox.value;
                event.preventDefault();
                return;
            }

            var selectedRadio = null;
            if (pressedKey === 'F') selectedRadio = dialogControls.placement.frontRadio;
            else if (pressedKey === 'B') selectedRadio = dialogControls.placement.backRadio;
            else if (pressedKey === 'L') selectedRadio = dialogControls.placement.bgLayerRadio;
            else if (pressedKey === 'C') selectedRadio = dialogControls.target.currentArtboardRadio;
            else if (pressedKey === 'A') selectedRadio = dialogControls.target.allArtboardsRadio;
            else return;

            if (selectedRadio.enabled) {
                selectedRadio.value = true;
                refreshPreview();
            }
            event.preventDefault();
        });
    }

    /**
     * ダイアログの入力内容から描画設定を組み立てる（プレビューと確定値で共通）
     * @param {object} dialogControls - 各パネルのコントロール
     * @returns {object} 描画設定
     */
    function collectDrawSettings(dialogControls) {
        var colorControls = dialogControls.color;
        var placementControls = dialogControls.placement;
        var offsetControls = dialogControls.offset;

        var colorMode = ColorMode.NONE;
        if (colorControls.k100Radio.value) colorMode = ColorMode.K100;
        else if (colorControls.hexRadio.value) colorMode = ColorMode.HEX;
        else if (colorControls.cmykRadio.value) colorMode = ColorMode.CMYK;

        var zOrder = placementControls.frontRadio.value ? 'front' :
            (placementControls.bgLayerRadio.value ? 'bg' : 'back');

        /* オフセット計算は resolveOffsetToPt に一元化 / All offset math lives in resolveOffsetToPt */
        var resolvedOffset = resolveOffsetToPt(offsetControls.offsetInput.text, getUnitInfo().code, !!offsetControls.bleedCheckbox.value);

        /* 各欄を0–100にクランプ（空欄・不正は0）/ Clamp each field to 0-100 (empty or invalid becomes 0) */
        var cmykChannelKeys = ['c', 'm', 'y', 'k'];
        var cmykValues = { c: 0, m: 0, y: 0, k: 0 };
        for (var i = 0; i < colorControls.cmykInputs.length; i++) {
            var channelValue = parseFloat(colorControls.cmykInputs[i].text);
            if (isNaN(channelValue)) channelValue = 0;
            cmykValues[cmykChannelKeys[i]] = clampValue(channelValue, 0, 100);
        }

        return {
            colorMode: colorMode,
            customValue: String(colorControls.hexInput.text || '').replace(/^\s+|\s+$/g, ''), /* HEX文字列 */
            customCMYK: cmykValues,
            offset: resolvedOffset.pt,
            zOrder: zOrder,
            target: dialogControls.target.allArtboardsRadio.value ? 'all' : 'current',
            makeGuide: !!dialogControls.options.makeGuideCheckbox.value,
            convertToLiveShape: !!dialogControls.options.convertToLiveShapeCheckbox.value
        };
    }

    /**
     * ボタン行（左：表示モード切り替え、右：キャンセル／OK）を構築する
     * @param {Window} settingsDialog - 追加先のダイアログ
     * @returns {{btnOK: Button, btnCancel: Button}} OK・キャンセルボタン
     */
    function buildButtonRow(settingsDialog) {
        var btnRowGroup = settingsDialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignChildren = ['fill', 'center'];
        btnRowGroup.alignment = 'fill';

        var btnLeftGroup = btnRowGroup.add('group');
        setupGroup(btnLeftGroup, 'row');

        var isPreviewDisplayMode = true;
        var btnDisplayToggle = btnLeftGroup.add('button', undefined, getLabel('button.previewOutline'));
        btnDisplayToggle.helpTip = getLabel('tooltip.previewToggle');

        var spacer = btnRowGroup.add('group');
        spacer.alignment = ['fill', 'fill'];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add('group');
        btnRightGroup.orientation = 'row';
        btnRightGroup.alignment = ['right', 'center']; /* 右カラムは右揃え / right-align the right column */

        var btnCancel = btnRightGroup.add('button', undefined, getLabel('button.cancel'));
        var btnOK = btnRightGroup.add('button', undefined, getLabel('button.ok'));

        btnDisplayToggle.onClick = function () {
            try {
                app.executeMenuCommand('preview');
                isPreviewDisplayMode = !isPreviewDisplayMode;
                btnDisplayToggle.text = isPreviewDisplayMode ? getLabel('button.previewOutline') : getLabel('button.previewPreview');
            } catch (e) { }
        };

        return { btnOK: btnOK, btnCancel: btnCancel };
    }

    /**
     * 設定ダイアログを構築して結果を返す
     * @returns {object|null} 描画設定。キャンセル時は null
     */
    function showDialog() {
        var settingsDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        DialogPersist.setOpacity(settingsDialog, DIALOG_OPACITY);
        settingsDialog.alignChildren = 'left';

        /* 各パネルより先に定義してコールバックとして配る（実行はパネル構築後）
           Declared before the panels so they can be handed out as callbacks */
        var dialogControls = null;

        function updatePreviewImmediately() {
            PreviewHistory.cancelTask(previewDebounceTaskId);
            try {
                renderPreview(app.activeDocument, collectDrawSettings(dialogControls));
            } catch (e) { }
        }

        function updatePreviewDeferred() {
            try {
                schedulePreview(collectDrawSettings(dialogControls), PREVIEW_DELAY_TYPING_MS);
            } catch (e) { }
        }

        var previewHooks = {
            immediate: updatePreviewImmediately,
            deferred: updatePreviewDeferred
        };

        /* 2カラム構成 / Two-column layout */
        var mainColumnsGroup = settingsDialog.add('group');
        setupGroup(mainColumnsGroup, 'row', COLUMN_SPACING);
        mainColumnsGroup.alignChildren = ['fill', 'top']; /* 2カラムを上揃え・横いっぱいに */

        var leftColumnGroup = mainColumnsGroup.add('group');
        setupGroup(leftColumnGroup, 'column', STACK_SPACING);
        leftColumnGroup.alignChildren = 'fill'; /* パネルを列幅いっぱいに / panels fill the column */

        var rightColumnGroup = mainColumnsGroup.add('group');
        setupGroup(rightColumnGroup, 'column', STACK_SPACING);
        rightColumnGroup.alignChildren = 'fill';

        dialogControls = {
            offset: buildOffsetPanel(leftColumnGroup, previewHooks),
            placement: buildPlacementPanel(leftColumnGroup, previewHooks),
            options: buildOptionsPanel(leftColumnGroup),
            color: buildColorPanel(rightColumnGroup, previewHooks),
            target: buildTargetPanel(rightColumnGroup, previewHooks)
        };

        addDialogHotkeys(settingsDialog, dialogControls, updatePreviewImmediately);

        var dialogButtons = buildButtonRow(settingsDialog);

        /* プレビューを片付けてから閉じる / Clean the preview up, then close */
        function closeWithCleanup(resultCode) {
            PreviewHistory.cancelTask(previewDebounceTaskId);
            PreviewHistory.undo();
            clearPreview(true); /* undo回数に依存せず _preview レイヤーを確実に削除 */
            settingsDialog.close(resultCode);
        }

        dialogButtons.btnOK.onClick = function () { closeWithCleanup(1); };
        dialogButtons.btnCancel.onClick = function () { closeWithCleanup(0); };

        settingsDialog.onShow = function () {
            DialogPersist.applyInitialOffset(settingsDialog, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
            dialogControls.offset.initFieldState();
            try { dialogControls.offset.offsetInput.active = true; } catch (e) { }
            PreviewHistory.start(); /* プレビューのUndoカウンタを初期化 */
            updatePreviewImmediately();
        };

        if (settingsDialog.show() != 1) return null;

        /* 確定値もプレビューと同じ計算経路から取る / Final values come from the same computation as the preview */
        return collectDrawSettings(dialogControls);
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * 編集可能なレイヤーを取得する（なければ作成）
     * テンプレートレイヤーは locked が false でも編集できない（Error 8705）ので除外し、
     * プレビュー用レイヤーは本番の描画先にしない（残存時の誤描画防止）。
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 描画先レイヤー
     */
    function getWritableLayer(doc) {
        function isWritableLayer(layer) {
            return !!layer && !layer.locked && layer.visible && !layer.template && !isPreviewLayerName(layer.name);
        }
        try {
            if (isWritableLayer(doc.activeLayer)) return doc.activeLayer;
            for (var i = 0; i < doc.layers.length; i++) {
                if (isWritableLayer(doc.layers[i])) return doc.layers[i];
            }
            var newLayer = doc.layers.add();
            newLayer.name = FALLBACK_LAYER_NAME;
            return newLayer;
        } catch (e) { }
        return doc.activeLayer;
    }

    /**
     * 「bg」レイヤーを取得する（なければ作成し、最背面へ移動）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} bgレイヤー
     */
    function getOrCreateBgLayer(doc) {
        var bgLayer = findLayerByName(doc, BG_LAYER_NAME);
        if (!bgLayer) {
            bgLayer = doc.layers.add();
            bgLayer.name = BG_LAYER_NAME;
        }
        /* 見える＆編集可能にしてから最背面へ / Make it visible and editable, then send it to the back */
        bgLayer.visible = true;
        bgLayer.locked = false;
        try {
            bgLayer.printable = true;
            bgLayer.move(doc, ElementPlacement.PLACEATEND);
        } catch (e) { }
        return bgLayer;
    }

    /**
     * アートボードと同サイズ（オフセット込み）の長方形を1枚描画する
     * @param {Document} doc - 対象ドキュメント
     * @param {Artboard} artboard - 対象アートボード
     * @param {object} drawSettings - 描画設定
     * @returns {PathItem} 描画した長方形
     */
    function drawRectangleForArtboard(doc, artboard, drawSettings) {
        var rectangleBounds = getOffsetRectangleBounds(artboard.artboardRect, drawSettings.offset);
        var targetLayer = (drawSettings.zOrder === 'bg') ? getOrCreateBgLayer(doc) : getWritableLayer(doc);
        /* アクティブレイヤーがロックされたままだと、別の編集可能レイヤーへ作成しても
           Illustrator が Error 8705（対象レイヤーは編集できません）を投げる。
           Make the target layer editable AND active before creating, or a locked active layer
           triggers Error 8705 "Target layer cannot be modified". */
        try {
            targetLayer.locked = false;
            targetLayer.visible = true;
            doc.activeLayer = targetLayer;
        } catch (e) { }

        var artboardRectangle = targetLayer.pathItems.rectangle(
            rectangleBounds.top, rectangleBounds.left, rectangleBounds.width, rectangleBounds.height);

        applyFillByMode(doc, artboardRectangle, drawSettings);
        artboardRectangle.name = getLabel('objectName.rect');
        artboardRectangle.selected = true;
        applyZOrder(artboardRectangle, drawSettings.zOrder);

        return artboardRectangle;
    }

    /**
     * 中心の○（属性パネル「中心点を表示」）を選択オブジェクトへ適用する
     * API・メニューコマンドからは設定できないため、記録済みアクション(.aia)を一時ファイルへ書き出して
     * loadAction→doScript で再生する。呼び出し側で「対象だけを選択した状態」にしてから実行すること。
     * @returns {void}
     */
    function showShapeCenterWidget() {
        var ACTION_SET_NAME = 'SmartDrawArtboardRectangle';
        var ACTION_NAME = 'CenterPoint';
        var ACTION_BODY = [
            '/version 3',
            '/name [ 26',
            '\t536d61727444726177417274626f61726452656374616e676c65',
            ']',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            '\t/name [ 11',
            '\t\t43656e746572506f696e74',
            '\t]',
            '\t/keyIndex 0',
            '\t/colorIndex 0',
            '\t/isOpen 1',
            '\t/eventCount 1',
            '\t/event-1 {',
            '\t\t/useRulersIn1stQuadrant 0',
            '\t\t/internalName (adobe_attributePalette)',
            '\t\t/localizedName [ 12',
            '\t\t\te5b19ee680a7e8a8ade5ae9a',
            '\t\t]',
            '\t\t/isOpen 1',
            '\t\t/isOn 1',
            '\t\t/hasDialog 0',
            '\t\t/parameterCount 1',
            '\t\t/parameter-1 {',
            '\t\t\t/key 1668183154',
            '\t\t\t/showInPalette 4294967295',
            '\t\t\t/type (boolean)',
            '\t\t\t/value 1',
            '\t\t}',
            '\t}',
            '}',
            ''
        ].join('\n');

        var actionFile = null;
        try {
            /* 一時ファイルへ書き出し / Write the recorded action to a temp file */
            actionFile = new File(Folder.temp + '/SmartDrawArtboardRectangle_center.aia');
            actionFile.encoding = 'UTF-8';
            actionFile.open('w');
            actionFile.write(ACTION_BODY);
            actionFile.close();

            /* 同名セットを解放してからロード→再生 / Unload any same-named set, then load and play */
            try { app.unloadAction(ACTION_SET_NAME, ''); } catch (e) { }
            app.loadAction(actionFile);
            app.doScript(ACTION_NAME, ACTION_SET_NAME, false);
        } catch (e) {
        } finally {
            try { app.unloadAction(ACTION_SET_NAME, ''); } catch (e) { }
            try { if (actionFile && actionFile.exists) actionFile.remove(); } catch (e) { }
        }
    }

    /**
     * 指定アイテムだけを選択状態にする
     * @param {PathItem[]} items - 選択したいアイテム
     * @returns {void}
     */
    function selectOnly(items) {
        try { app.executeMenuCommand('deselectall'); } catch (e) { }
        try {
            for (var i = 0; i < items.length; i++) items[i].selected = true;
        } catch (e) { }
    }

    /**
     * 描画後のオプション（ライブシェイプ化／中心の○表示／ガイド化）を適用する
     * @param {PathItem[]} createdRectangles - 描画した長方形
     * @param {object} drawSettings - 描画設定
     * @returns {void}
     */
    function applyDrawOptions(createdRectangles, drawSettings) {
        if (!createdRectangles || !createdRectangles.length) return;

        if (drawSettings.convertToLiveShape) {
            /* 選択ベースのメニューコマンドなので、対象だけを選択してから実行 */
            selectOnly(createdRectangles);
            try { app.executeMenuCommand('Convert to Shape'); } catch (e) { }
            /* 中心の○はライブシェイプにのみ表示される。変換後に選択し直してから適用
               The center widget only renders on live shapes, so re-select after converting */
            selectOnly(createdRectangles);
            showShapeCenterWidget();
        }

        if (drawSettings.makeGuide) {
            /* PathItem.guides を直接立てる（選択・メニュー状態に依存せず確実）
               Set PathItem.guides directly - robust, independent of selection and menu state */
            for (var i = 0; i < createdRectangles.length; i++) {
                createdRectangles[i].name = getLabel('objectName.guide');
                createdRectangles[i].guides = true;
            }
            try { app.executeMenuCommand('deselectall'); } catch (e) { }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリポイント：ダイアログ→描画→オプション適用
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;

        var drawSettings = showDialog();
        if (drawSettings === null) return;

        var doc = app.activeDocument;
        var previousCoordinateSystem = switchToDocumentCoordinates();

        app.executeMenuCommand('deselectall'); /* 既存選択を解除 / clear any existing selection */

        var createdRectangles = [];
        if (drawSettings.target === 'all') {
            for (var i = 0; i < doc.artboards.length; i++) {
                createdRectangles.push(drawRectangleForArtboard(doc, doc.artboards[i], drawSettings));
            }
        } else {
            var currentArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
            createdRectangles.push(drawRectangleForArtboard(doc, currentArtboard, drawSettings));
        }

        applyDrawOptions(createdRectangles, drawSettings);
        restoreCoordinateSystem(previousCoordinateSystem);
    }

    main();

})();
