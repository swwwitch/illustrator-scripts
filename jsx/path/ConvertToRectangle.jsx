#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの境界に合わせて長方形を作成します。
作成単位、マージン、角丸、塗り・線プリセット、元オブジェクトの扱いをダイアログで指定でき、プレビューにも対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToRectangle.md

### Overview

Creates rectangles that match the bounds of the selected objects.
The unit of creation, margin, corner radius, fill and stroke presets, and what happens to the originals are set in a dialog, with a preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ConvertToRectangle";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertToRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertToRectangle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    // ==============================
    // ラベル定義 / Label definitions
    // ==============================
    var LABELS = {
        /* Dialog / ダイアログ */
        dialogTitle: { ja: "長方形を作成", en: "Create Rectangle" },
        cancel: { ja: "キャンセル", en: "Cancel" },
        preview: { ja: "プレビュー", en: "Preview" },
        previewTip: { ja: "オンで作成される長方形をその場で表示します（マスク／削除はOK後に反映）", en: "Shows the resulting rectangle while you adjust (mask/delete applied on OK)" },

        /* Panels / パネル */
        panelTarget: { ja: "作成単位", en: "Creation unit" },
        panelOptions: { ja: "オプション", en: "Options" },
        panelAppearance: { ja: "塗りと線", en: "Fill and stroke" },
        panelCorner: { ja: "角丸", en: "Rounded corners" },
        panelOrder: { ja: "長方形の重ね順", en: "Rectangle order" },
        panelOriginal: { ja: "元オブジェクトの扱い", en: "Original handling" },

        /* Target radios / 作成単位 */
        targetEach: { ja: "オブジェクトごと", en: "Per object" },
        targetEachTip: { ja: "選択中の各オブジェクトに対して、それぞれ長方形を作成します", en: "Creates a separate rectangle for each selected object" },
        targetGroup: { ja: "選択範囲全体", en: "Whole selection" },
        targetGroupTip: { ja: "選択中のオブジェクト全体を1つの境界として、長方形を1つ作成します", en: "Creates one rectangle using the bounds of the entire selection" },

        /* Options / オプション */
        previewBounds: { ja: "プレビュー境界で計測", en: "Use preview bounds" },
        boundsTip: { ja: "オンで線幅・効果を含む見た目のサイズを基準にします。オフではパス本体の境界を使います", en: "When on, measures the visible size including stroke and effects. When off, uses the path bounds" },
        outlineText: { ja: "テキストはアウトライン化して計測", en: "Measure text as outlines" },
        outlineTextTip: { ja: "元のテキストは変更せず、複製を一時的にアウトライン化して境界を計測します", en: "Measures a temporary outlined copy without modifying the original text" },
        outlineTextDisabledTip: { ja: "選択にテキストが含まれていないため無効です", en: "Disabled because the selection does not include text" },
        useMargin: { ja: "マージンを追加", en: "Add margin" },
        marginTip: { ja: "オンにすると、計測した境界から指定値だけ外側へ広げます。単位は Illustrator の定規単位に連動します（負値で内側）", en: "When on, expands the measured bounds outward by this amount. The unit follows Illustrator's ruler unit (negative shrinks)" },
        dimSelection: { ja: "実行中は不透明度を下げる", en: "Dim selection while running" },
        dimSelectionTip: { ja: "ダイアログを開いている間、選択オブジェクトの不透明度を一時的に 50% にして、閉じたときに元に戻します", en: "Temporarily sets the selected objects to 50% opacity while the dialog is open, and restores it when the dialog closes" },

        /* Appearance / 塗りと線 */
        currentTip: { ja: "ドキュメントの現在の既定値（直前に使った塗り・線）をそのまま使います", en: "Uses the document's current default fill and stroke" },
        imageNotice: { ja: "画像では直前の塗り・線を取得できないため選択できません", en: "Current fill/stroke is not available for image items" },

        /* Corner / 角丸 */
        cornerRadius: { ja: "半径", en: "Radius" },
        cornerRadiusTip: { ja: "生成する長方形に角丸ライブエフェクトを適用します。0 で角丸なし（単位は定規単位）。マスクのクリップ形状はパス本体に従うため、丸めるには「アピアランスを分割」で実体化が必要", en: "Applies a Round Corners live effect to the new rectangle. 0 means no rounding (unit follows the ruler). Clipping uses the geometric path, so use Expand Appearance to bake the corners when masking" },

        /* Order radios / 重ね順 */
        orderFront: { ja: "前面に作成", en: "Create in front" },
        orderFrontTip: { ja: "生成する長方形を元オブジェクトの前面に配置します。「元オブジェクトの扱い」が「残す」のときだけ有効です", en: "Places the new rectangle in front of the original. Available only when the original is kept" },
        orderBack: { ja: "背面に作成", en: "Create behind" },
        orderBackTip: { ja: "生成する長方形を元オブジェクトの背面に配置します。「元オブジェクトの扱い」が「残す」のときだけ有効です", en: "Places the new rectangle behind the original. Available only when the original is kept" },

        /* Original handling / 元オブジェクト */
        originalKeep: { ja: "残す", en: "Keep" },
        originalKeepTip: { ja: "元オブジェクトを残し、長方形だけを追加します", en: "Keeps the original object and adds the rectangle" },
        originalMask: { ja: "クリッピングマスクにする", en: "Make clipping mask" },
        maskTip: { ja: "生成した長方形をクリップパスとして、元オブジェクトをクリッピングマスク化します", en: "Uses the new rectangle as the clipping path for the original object" },
        maskImageOnly: { ja: "選択がリンク画像／埋め込み画像のときだけ選べます", en: "Available only when the selection is linked or embedded images" },
        originalDelete: { ja: "削除", en: "Delete" },
        originalDeleteTip: { ja: "長方形の作成後、元オブジェクトを削除します", en: "Deletes the original object after creating the rectangle" },

        /* Alerts / 警告 */
        noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
        noSelection: { ja: "オブジェクトを選択してください。", en: "Please select one or more objects." },
        noResult: { ja: "長方形を作成できるオブジェクトがありませんでした。", en: "No rectangles could be created." },

        /* Stepper / ステップボタン */
        stepUpTip: {
            ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepDownTip: {
            ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
            en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
        },
        stepUpIntegerTip: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
        stepDownIntegerTip: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
    };

    // ラベル取得 / Get label
    function getLabel(labelKey) {
        var labelEntry = LABELS[labelKey];
        if (!labelEntry) return labelKey;
        return labelEntry[currentLanguage] || labelEntry.en || labelKey;
    }

    // ==============================
    // 塗り・線プリセット / Fill & stroke presets
    // ==============================
    // applyAppearance: true … 生成した長方形にこのプリセットの塗り・線を適用する
    // applyAppearance: false（直前の塗り・線）… 適用せず、作成時の既定値のままにする
    // strokeWidth: Illustrator の「線幅」環境設定単位で指定し、適用時に pt へ変換する
    // opacity (省略時 100): プリセットで不透明度を変える場合に指定
    var FILL_STROKE_PRESETS = [
        { ja: "塗り：なし、線：なし", en: "Fill: none, Stroke: none", applyAppearance: true, filled: false, stroked: false, strokeWidth: 0 },
        { ja: "線：黒、{width}{unit}", en: "Stroke: black, {width}{unit}", applyAppearance: true, filled: false, stroked: true, strokeWidth: 1 },
        { ja: "線：黒、{width}{unit}", en: "Stroke: black, {width}{unit}", applyAppearance: true, filled: false, stroked: true, strokeWidth: 0.25 },
        { ja: "塗り：黒、不透明度：40%", en: "Fill: black 40%, Stroke: none", applyAppearance: true, filled: true, stroked: false, strokeWidth: 0, opacity: 40 },
        { ja: "直前の塗り・線の情報", en: "Current fill / stroke", applyAppearance: false }
    ];
    var DEFAULT_PRESET_INDEX = 1; // 既定：線：黒、1（線幅設定単位）

    // 「直前の塗り・線の情報」プリセットの位置（applyAppearance: false）
    var CURRENT_PRESET_INDEX = (function () {
        for (var i = 0; i < FILL_STROKE_PRESETS.length; i++) {
            if (!FILL_STROKE_PRESETS[i].applyAppearance) return i;
        }
        return -1;
    })();

    // ==============================
    // 単位変換 / Unit conversion
    // ==============================

    // Illustrator の strokeUnits 設定値を取得する / Get Illustrator strokeUnits preference
    function getStrokeUnitsPreference() {
        var strokeUnits = 2; // 既定：pt / Default: points
        try {
            strokeUnits = app.preferences.getIntegerPreference("strokeUnits");
        } catch (ignoreStrokeUnits) { }
        return strokeUnits;
    }

    // Illustrator の単位設定値ごとの換算情報 / Conversion info for Illustrator unit preferences
    var UNIT_INFO_BY_TYPE = {
        0: { factor: 72, label: "in" },        // inch
        1: { factor: 72 / 25.4, label: "mm" }, // mm
        2: { factor: 1, label: "pt" },         // pt
        3: { factor: 12, label: "pc" },        // pica
        4: { factor: 72 / 2.54, label: "cm" }, // cm
        5: { factor: 0.25, label: "Q" },       // Q/H
        6: { factor: 1, label: "px" }          // px
    };

    // Illustrator の単位設定値に対応する換算情報を返す / Get conversion info for an Illustrator unit preference value
    function getUnitInfo(unitType) {
        return UNIT_INFO_BY_TYPE[unitType] || UNIT_INFO_BY_TYPE[2];
    }

    // Illustrator の単位設定値を pt 変換係数に変換する / Convert Illustrator unit preference value to a points factor
    function getUnitToPointFactor(unitType) {
        return getUnitInfo(unitType).factor;
    }

    // Illustrator の単位設定値を表示用単位ラベルに変換する / Convert Illustrator unit preference value to a display label
    function getUnitLabel(unitType) {
        return getUnitInfo(unitType).label;
    }

    // Illustrator の strokeUnits 設定値を pt 変換係数に変換する / Convert Illustrator strokeUnits to a points factor
    function getStrokeUnitToPointFactor() {
        return getUnitToPointFactor(getStrokeUnitsPreference());
    }

    // Illustrator の strokeUnits 設定値を表示用単位ラベルに変換する / Convert Illustrator strokeUnits to a display label
    function getStrokeUnitLabel() {
        return getUnitLabel(getStrokeUnitsPreference());
    }

    // Illustrator の rulerType 設定値を取得する / Get Illustrator rulerType preference
    function getRulerTypePreference() {
        var rulerType = 2; // 既定：pt / Default: points
        try {
            rulerType = app.preferences.getIntegerPreference("rulerType");
        } catch (ignoreRulerType) { }
        return rulerType;
    }

    // Illustrator の rulerType 設定値を pt 変換係数に変換する / Convert Illustrator rulerType to a points factor
    function getRulerUnitToPointFactor() {
        return getUnitToPointFactor(getRulerTypePreference());
    }

    // Illustrator の rulerType 設定値を表示用単位ラベルに変換する / Convert Illustrator rulerType to a display label
    function getRulerUnitLabel() {
        return getUnitLabel(getRulerTypePreference());
    }

    // 定規単位の値を pt に変換する / Convert ruler-unit value to points
    function rulerUnitValueToPoints(value) {
        return value * getRulerUnitToPointFactor();
    }

    // 線幅設定単位の値を pt に変換する / Convert stroke-unit value to points
    function strokeUnitValueToPoints(value) {
        return value * getStrokeUnitToPointFactor();
    }

    // ドキュメントのカラースペースに合わせた黒を返す / Returns black for the document color space
    function makeBlack(isRgbDocument) {
        if (isRgbDocument) {
            var rgb = new RGBColor();
            rgb.red = rgb.green = rgb.blue = 0;
            return rgb;
        }
        var cmyk = new CMYKColor();
        cmyk.cyan = cmyk.magenta = cmyk.yellow = 0;
        cmyk.black = 100;
        return cmyk;
    }

    // 線幅プリセット名に現在の線幅単位を反映する / Apply the current stroke unit to preset labels
    function formatFillStrokePresetLabel(preset) {
        var presetLabel = preset[currentLanguage] || preset.en;
        if (presetLabel.indexOf("{width}") < 0) return presetLabel;

        return presetLabel
            .replace("{width}", preset.strokeWidth)
            .replace("{unit}", getStrokeUnitLabel());
    }

    // ==============================
    // 選択オブジェクトの判定 / Selection inspection
    // ==============================

    // リンク画像（PlacedItem）または配置画像（RasterItem）かどうか
    function isImageItem(item) {
        return item.typename === "PlacedItem" || item.typename === "RasterItem";
    }

    // 選択オブジェクトがすべて画像かどうか / True when every selected item is an image
    function isSelectionAllImages(pageItems) {
        if (!pageItems || pageItems.length === 0) return false;
        for (var i = 0; i < pageItems.length; i++) {
            if (!isImageItem(pageItems[i])) return false;
        }
        return true;
    }

    // 選択オブジェクトに TextFrame が含まれるかどうか / True when any selected item is a TextFrame
    function selectionHasText(pageItems) {
        if (!pageItems || pageItems.length === 0) return false;
        for (var i = 0; i < pageItems.length; i++) {
            if (pageItems[i].typename === "TextFrame") return true;
        }
        return false;
    }

    // ==============================
    // パネル構築ヘルパー / Panel building helpers
    // ==============================
    var PANEL_MARGINS = [15, 20, 15, 10];
    var PANEL_SPACING = 8;

    // パネル共通設定。spacing 省略時は PANEL_SPACING
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ['fill', 'top'];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // 縦並びパネルを追加（チェックボックス・ラジオを縦に並べる用）
    function addColumnPanel(parent, labelKey) {
        var panel = parent.add("panel", undefined, getLabel(labelKey));
        setupPanel(panel);
        return panel;
    }

    // ==============================
    // ダイアログ用ユーティリティ / Dialog utilities
    // ==============================

    // value が true のラジオの index を返す。見つからなければ fallback
    function findSelectedRadioIndex(radios, fallback) {
        for (var i = 0; i < radios.length; i++) {
            if (radios[i].value) return i;
        }
        return fallback;
    }

    // 元オブジェクトの扱いラジオから "keep" | "mask" | "delete" を返す
    function getOriginalMode(maskRadio, deleteRadio) {
        if (maskRadio.value) return "mask";
        if (deleteRadio.value) return "delete";
        return "keep";
    }

    // ダイアログ全体にキーボードショートカットを付与する
    // handlers: { "S": function(){...}, "G": function(){...}, ... }
    // テキスト入力にフォーカスがあるとき、および修飾キー併用時は発火しない
    function addKeyboardShortcuts(dialog, handlers) {
        dialog.addEventListener("keydown", function (event) {
            // テキスト編集中はショートカットを横取りしない
            if (event.target && event.target.type === "edittext") return;
            // 修飾キー併用時はシステムショートカットの可能性があるため無視
            var keyboard = ScriptUI.environment.keyboardState;
            if (keyboard.shiftKey || keyboard.ctrlKey || keyboard.altKey || keyboard.metaKey) return;
            var handler = handlers[event.keyName];
            if (handler) {
                handler();
                event.preventDefault();
            }
        });
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    // 2. コピー先の LABELS に stepUpTip / stepDownTip / stepUpIntegerTip / stepDownIntegerTip を足す（このファイルの LABELS から写す）。
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
        var upTooltip = stepOptions.integer ? "stepUpIntegerTip" : "stepUpTip";
        var downTooltip = stepOptions.integer ? "stepDownIntegerTip" : "stepDownTip";
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

    /**
     * ∧∨と数値入力欄を隙間0で突き合わせて追加する（↑↓キーと onStep は showSettingsDialog でつなぐ）
     * @param {Group} parentRow - 追加先の行
     * @param {string} initialText - 入力欄の初期値
     * @param {number} characters - 入力欄の文字数
     * @param {Object} stepOptions - ∧∨の設定（min など）
     * @returns {EditText} 追加した入力欄（∧∨は .stepperGroup、設定は .stepOptions で参照できる）
     */
    function addStepperInput(parentRow, initialText, characters, stepOptions) {
        var stepperFieldGroup = parentRow.add("group");
        stepperFieldGroup.orientation = "row";
        stepperFieldGroup.alignChildren = ["left", "center"];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var numberInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return numberInput; }, stepOptions);
        numberInput = stepperFieldGroup.add("edittext", undefined, initialText);
        numberInput.characters = characters;
        numberInput.stepperGroup = stepperGroup;
        numberInput.stepOptions = stepOptions;
        return numberInput;
    }

    // ==============================
    // 設定ダイアログ / Settings dialog
    // ==============================
    function createSettingsDialogWindow() {
        var dialog = new Window("dialog", getLabel("dialogTitle") + "  " + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.margins = 16;
        dialog.spacing = 12;
        return dialog;
    }

    function createDialogColumns(dialog) {
        var columnsGroup = dialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = "top";
        columnsGroup.spacing = 12;

        var leftColumn = columnsGroup.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "fill";
        leftColumn.spacing = 12;

        var rightColumn = columnsGroup.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "fill";
        rightColumn.spacing = 12;

        return {
            leftColumn: leftColumn,
            rightColumn: rightColumn
        };
    }

    function buildTargetPanel(parent) {
        var targetPanel = addColumnPanel(parent, "panelTarget");
        var eachRadio = targetPanel.add("radiobutton", undefined, getLabel("targetEach"));
        eachRadio.helpTip = getLabel("targetEachTip");
        var groupRadio = targetPanel.add("radiobutton", undefined, getLabel("targetGroup"));
        groupRadio.helpTip = getLabel("targetGroupTip");
        eachRadio.value = true;
        return {
            eachRadio: eachRadio,
            groupRadio: groupRadio
        };
    }

    function buildOptionsPanel(parent, selectionHasTextItem) {
        var optionsPanel = addColumnPanel(parent, "panelOptions");
        var previewBoundsCheck = optionsPanel.add("checkbox", undefined, getLabel("previewBounds"));
        previewBoundsCheck.helpTip = getLabel("boundsTip");
        var outlineTextCheck = optionsPanel.add("checkbox", undefined, getLabel("outlineText"));
        outlineTextCheck.helpTip = selectionHasTextItem ? getLabel("outlineTextTip") : getLabel("outlineTextDisabledTip");
        outlineTextCheck.enabled = selectionHasTextItem;

        var marginRow = optionsPanel.add("group");
        marginRow.orientation = "row";
        marginRow.alignChildren = ["left", "center"];
        var useMarginCheck = marginRow.add("checkbox", undefined, getLabel("useMargin"));
        useMarginCheck.helpTip = getLabel("marginTip");
        /* マージンは負の値（内側）も可 / negative margins shrink inward */
        var marginInput = addStepperInput(marginRow, "0", 5, {});
        marginInput.enabled = false;
        marginInput.stepperGroup.enabled = false;
        marginInput.helpTip = getLabel("marginTip");
        var marginUnitLabel = marginRow.add("statictext", undefined, getRulerUnitLabel());
        marginUnitLabel.enabled = false;

        var dimCheck = optionsPanel.add("checkbox", undefined, getLabel("dimSelection"));
        dimCheck.helpTip = getLabel("dimSelectionTip");
        dimCheck.value = true;

        return {
            previewBoundsCheck: previewBoundsCheck,
            outlineTextCheck: outlineTextCheck,
            useMarginCheck: useMarginCheck,
            marginInput: marginInput,
            marginUnitLabel: marginUnitLabel,
            dimCheck: dimCheck
        };
    }

    function buildOriginalPanel(parent, selectionAllImages) {
        var originalPanel = addColumnPanel(parent, "panelOriginal");
        var keepRadio = originalPanel.add("radiobutton", undefined, getLabel("originalKeep"));
        keepRadio.helpTip = getLabel("originalKeepTip");
        var maskRadio = originalPanel.add("radiobutton", undefined, getLabel("originalMask"));
        maskRadio.helpTip = selectionAllImages ? getLabel("maskTip") : getLabel("maskImageOnly");
        maskRadio.enabled = selectionAllImages;
        var deleteRadio = originalPanel.add("radiobutton", undefined, getLabel("originalDelete"));
        deleteRadio.helpTip = getLabel("originalDeleteTip");
        keepRadio.value = true;
        return {
            keepRadio: keepRadio,
            maskRadio: maskRadio,
            deleteRadio: deleteRadio
        };
    }

    function buildAppearancePanel(parent, selectionAllImages) {
        var appearancePanel = addColumnPanel(parent, "panelAppearance");
        var radios = [];
        for (var i = 0; i < FILL_STROKE_PRESETS.length; i++) {
            var presetLabel = formatFillStrokePresetLabel(FILL_STROKE_PRESETS[i]);
            radios.push(appearancePanel.add("radiobutton", undefined, presetLabel));
        }
        radios[DEFAULT_PRESET_INDEX].value = true;

        if (CURRENT_PRESET_INDEX >= 0) {
            radios[CURRENT_PRESET_INDEX].helpTip = getLabel("currentTip");
        }
        if (selectionAllImages && CURRENT_PRESET_INDEX >= 0) {
            radios[CURRENT_PRESET_INDEX].enabled = false;
            radios[CURRENT_PRESET_INDEX].helpTip = getLabel("imageNotice");
        }
        return radios;
    }

    function buildCornerPanel(parent) {
        var cornerPanel = addColumnPanel(parent, "panelCorner");
        var cornerRow = cornerPanel.add("group");
        cornerRow.orientation = "row";
        cornerRow.alignChildren = ["left", "center"];
        var cornerRadiusLabel = cornerRow.add("statictext", undefined, getLabel("cornerRadius"));
        cornerRadiusLabel.helpTip = getLabel("cornerRadiusTip");
        var cornerRadiusInput = addStepperInput(cornerRow, "0", 5, { min: 0 });
        cornerRadiusInput.helpTip = getLabel("cornerRadiusTip");
        var cornerUnitLabel = cornerRow.add("statictext", undefined, getRulerUnitLabel());
        return {
            cornerRadiusInput: cornerRadiusInput,
            cornerUnitLabel: cornerUnitLabel
        };
    }

    function buildOrderPanel(parent) {
        var orderPanel = addColumnPanel(parent, "panelOrder");
        var frontRadio = orderPanel.add("radiobutton", undefined, getLabel("orderFront"));
        frontRadio.helpTip = getLabel("orderFrontTip");
        var backRadio = orderPanel.add("radiobutton", undefined, getLabel("orderBack"));
        backRadio.helpTip = getLabel("orderBackTip");
        frontRadio.value = true;
        return {
            frontRadio: frontRadio,
            backRadio: backRadio
        };
    }

    function buildButtonRow(dialog) {
        var buttonRow = dialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = "fill";
        buttonRow.alignChildren = ["fill", "center"];

        var buttonLeft = buttonRow.add("group");
        buttonLeft.alignment = ["left", "center"];
        var previewCheck = buttonLeft.add("checkbox", undefined, getLabel("preview"));
        previewCheck.helpTip = getLabel("previewTip");

        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];

        var buttonRight = buttonRow.add("group");
        buttonRight.alignment = ["right", "center"];
        var cancelButton = buttonRight.add("button", undefined, getLabel("cancel"), { name: "cancel" });
        var okButton = buttonRight.add("button", undefined, "OK", { name: "ok" });

        return {
            previewCheck: previewCheck,
            cancelButton: cancelButton,
            okButton: okButton
        };
    }

    function getMarginValue(ui) {
        if (!ui.options.useMarginCheck.value) return 0;
        return rulerUnitValueToPoints(parseFloat(ui.options.marginInput.text) || 0);
    }

    function getCornerRadiusValue(ui) {
        var raw = parseFloat(ui.corner.cornerRadiusInput.text) || 0;
        if (raw <= 0) return 0;
        return rulerUnitValueToPoints(raw);
    }

    function buildDialogSettings(ui, forPreview) {
        var appearanceIndex = findSelectedRadioIndex(ui.appearanceRadios, DEFAULT_PRESET_INDEX);
        return {
            groupAsOne: ui.target.groupRadio.value,
            useVisibleBounds: ui.options.previewBoundsCheck.value,
            outlineText: ui.options.outlineTextCheck.value,
            margin: getMarginValue(ui),
            appearancePreset: FILL_STROKE_PRESETS[appearanceIndex],
            cornerRadius: getCornerRadiusValue(ui),
            placeInFront: ui.order.frontRadio.value,
            originalMode: forPreview ? "keep" : getOriginalMode(ui.original.maskRadio, ui.original.deleteRadio)
        };
    }
    function syncMarginEnabled(ui) {
        ui.options.marginInput.enabled = ui.options.useMarginCheck.value;
        ui.options.marginInput.stepperGroup.enabled = ui.options.useMarginCheck.value;
        redrawSteppersIn(ui.options.marginInput.stepperGroup);
        ui.options.marginUnitLabel.enabled = ui.options.useMarginCheck.value;
    }

    // ==============================
    // プレビュー undo ヘルパー / Preview undo helpers
    // ==============================
    // 設定変更のたびに「仮配置 → undo → 再生成」するパターン
    // state: { isUndo: boolean, items: [...] } を呼び出し側で保持する
    // process() の戻り値（追加された PageItem 配列）が state.items に格納される

    // 残ったプレビューアイテムを直接削除する（app.undo() が複数ステップを巻き戻さない場合の保険）
    function removePreviewRemnants(state) {
        if (!state.items) {
            state.items = [];
            return;
        }
        for (var i = 0; i < state.items.length; i++) {
            try { state.items[i].remove(); } catch (ignoreRemoveRemnant) { }
        }
        state.items = [];
    }

    // プレビューを再生成する / Re-render preview
    // processFn は追加された PageItem の配列を返すこと
    function runPreview(state, processFn, isEnabled) {
        try {
            if (isEnabled) {
                if (state.isUndo) {
                    app.undo();
                    removePreviewRemnants(state);
                } else {
                    state.isUndo = true;
                }
                state.items = processFn() || [];
                app.redraw();
            } else if (state.isUndo) {
                app.undo();
                removePreviewRemnants(state);
                app.redraw();
                state.isUndo = false;
            }
        } catch (ignorePreviewRun) { }
    }

    // 確定処理の直前にプレビュー分を巻き戻す / Undo preview before commit
    function undoPreview(state) {
        try {
            if (state.isUndo) {
                app.undo();
                removePreviewRemnants(state);
            }
        } catch (ignorePreviewUndo) { }
        state.isUndo = false;
    }

    // ダイアログクローズ時のクリーンアップ / Cleanup on dialog close
    function cleanupPreview(state) {
        try {
            if (state.isUndo) {
                app.undo();
                removePreviewRemnants(state);
            }
            state.isUndo = false;
        } catch (ignorePreviewCleanup) { }
    }

    // 元の不透明度を退避（後で確実に復元するため）
    function captureItemOpacities(items) {
        var snapshots = [];
        for (var i = 0; i < items.length; i++) {
            var current = 100;
            try { current = items[i].opacity; } catch (ignoreOpacityRead) { }
            snapshots.push({ item: items[i], opacity: current });
        }
        return snapshots;
    }

    function applyDimmedOpacity(items) {
        for (var i = 0; i < items.length; i++) {
            try { items[i].opacity = 50; } catch (ignoreDim) { }
        }
    }

    function restoreItemOpacities(snapshots) {
        for (var i = 0; i < snapshots.length; i++) {
            try { snapshots[i].item.opacity = snapshots[i].opacity; } catch (ignoreRestore) { }
        }
    }

    function renderDialogPreview(doc, eligibleItems, ui, previewState, outlinedBoundsCache) {
        runPreview(previewState, function () {
            var previewSettings = buildDialogSettings(ui, true);
            return produceRectangles(doc, eligibleItems, previewSettings, outlinedBoundsCache);
        }, ui.buttons.previewCheck.value);
    }

    function createDialogState(doc, eligibleItems, ui) {
        var previewState = { isUndo: false, items: [] };
        var outlinedBoundsCache = [];
        var opacitySnapshots = captureItemOpacities(eligibleItems);
        var dimmed = false;
        /* 不透明度付きプリセットで強制 OFF にする前の値を保存する（戻すときに使う）/
           Saves the dim value before being forced off by an opacity preset */
        var savedDimValue = null;

        function syncDimAvailability() {
            var appearanceIndex = findSelectedRadioIndex(ui.appearanceRadios, DEFAULT_PRESET_INDEX);
            var preset = FILL_STROKE_PRESETS[appearanceIndex];
            var presetHasOpacity = preset && typeof preset.opacity === "number";
            if (presetHasOpacity) {
                if (savedDimValue === null) {
                    savedDimValue = ui.options.dimCheck.value;
                }
                ui.options.dimCheck.value = false;
                ui.options.dimCheck.enabled = false;
            } else {
                if (savedDimValue !== null) {
                    ui.options.dimCheck.value = savedDimValue;
                    savedDimValue = null;
                }
                ui.options.dimCheck.enabled = true;
            }
        }

        function syncDimming() {
            var shouldDim = ui.options.dimCheck.value;
            if (shouldDim && !dimmed) {
                applyDimmedOpacity(eligibleItems);
                dimmed = true;
            } else if (!shouldDim && dimmed) {
                restoreItemOpacities(opacitySnapshots);
                dimmed = false;
            }
            app.redraw();
        }

        function restoreOpacity() {
            if (dimmed) {
                restoreItemOpacities(opacitySnapshots);
                dimmed = false;
            }
        }

        return {
            buildSettings: function (forPreview) {
                return buildDialogSettings(ui, forPreview);
            },
            undoPreview: function () {
                undoPreview(previewState);
            },
            cleanupPreview: function () {
                cleanupPreview(previewState);
            },
            renderPreview: function () {
                renderDialogPreview(doc, eligibleItems, ui, previewState, outlinedBoundsCache);
            },
            syncDimAvailability: syncDimAvailability,
            syncDimming: syncDimming,
            restoreOpacity: restoreOpacity
        };
    }

    function syncOrderEnabled(ui) {
        ui.order.frontRadio.enabled = ui.original.keepRadio.value;
        ui.order.backRadio.enabled = ui.original.keepRadio.value;
    }

    function bindPreviewEvents(ui, state) {
        ui.target.eachRadio.onClick = state.renderPreview;
        ui.target.groupRadio.onClick = state.renderPreview;
        ui.options.previewBoundsCheck.onClick = state.renderPreview;
        ui.options.outlineTextCheck.onClick = state.renderPreview;
        ui.options.useMarginCheck.onClick = function () {
            syncMarginEnabled(ui);
            state.renderPreview();
        };
        ui.options.marginInput.onChange = state.renderPreview;
        /* ディム変更は opacity 書き換えで undo エントリを作るので、プレビュー巻き戻しを先に行う */
        ui.options.dimCheck.onClick = function () {
            state.undoPreview();
            state.syncDimming();
            state.renderPreview();
        };
        for (var appearanceRadioIndex = 0; appearanceRadioIndex < ui.appearanceRadios.length; appearanceRadioIndex++) {
            ui.appearanceRadios[appearanceRadioIndex].onClick = function () {
                state.undoPreview();
                state.syncDimAvailability();
                state.syncDimming();
                state.renderPreview();
            };
        }
        ui.corner.cornerRadiusInput.onChange = state.renderPreview;
        ui.order.frontRadio.onClick = state.renderPreview;
        ui.order.backRadio.onClick = state.renderPreview;
        ui.buttons.previewCheck.onClick = state.renderPreview;
    }

    function bindOriginalHandlingEvents(ui) {
        ui.original.keepRadio.onClick = function () { syncOrderEnabled(ui); };
        ui.original.maskRadio.onClick = function () { syncOrderEnabled(ui); };
        ui.original.deleteRadio.onClick = function () { syncOrderEnabled(ui); };
    }

    function bindDialogKeyboardShortcuts(dialog, ui, state) {
        addKeyboardShortcuts(dialog, {
            "S": function () { ui.target.eachRadio.value = true; state.renderPreview(); },
            "G": function () { ui.target.groupRadio.value = true; state.renderPreview(); },
            "P": function () { ui.options.previewBoundsCheck.value = !ui.options.previewBoundsCheck.value; state.renderPreview(); },
            "O": function () {
                if (!ui.options.outlineTextCheck.enabled) return; // テキスト非選択時はスキップ
                ui.options.outlineTextCheck.value = !ui.options.outlineTextCheck.value;
                state.renderPreview();
            },
            "A": function () { ui.options.useMarginCheck.value = !ui.options.useMarginCheck.value; syncMarginEnabled(ui); state.renderPreview(); },
            "T": function () {
                if (!ui.options.dimCheck.enabled) return; // 不透明度プリセット選択中は無効化されているのでスキップ
                state.undoPreview();
                ui.options.dimCheck.value = !ui.options.dimCheck.value;
                state.syncDimming();
                state.renderPreview();
            },
            "F": function () { ui.order.frontRadio.value = true; state.renderPreview(); },
            "B": function () { ui.order.backRadio.value = true; state.renderPreview(); },
            "N": function () { ui.original.keepRadio.value = true; syncOrderEnabled(ui); },
            "M": function () {
                if (!ui.original.maskRadio.enabled) return; // 画像以外ではマスク選択不可
                ui.original.maskRadio.value = true;
                syncOrderEnabled(ui);
            },
            "D": function () { ui.original.deleteRadio.value = true; syncOrderEnabled(ui); }
        });
    }

    function bindDialogEvents(dialog, ui, state) {
        bindPreviewEvents(ui, state);
        bindOriginalHandlingEvents(ui);
        bindDialogKeyboardShortcuts(dialog, ui, state);
    }

    function buildDialogUI(dialog, selectionAllImages, selectionHasTextItem) {
        var columns = createDialogColumns(dialog);
        return {
            target: buildTargetPanel(columns.leftColumn),
            options: buildOptionsPanel(columns.leftColumn, selectionHasTextItem),
            original: buildOriginalPanel(columns.leftColumn, selectionAllImages),
            appearanceRadios: buildAppearancePanel(columns.rightColumn, selectionAllImages),
            corner: buildCornerPanel(columns.rightColumn),
            order: buildOrderPanel(columns.rightColumn),
            buttons: buildButtonRow(dialog)
        };
    }

    function showSettingsDialog(doc, eligibleItems, selectionAllImages, selectionHasTextItem) {
        var dialog = createSettingsDialogWindow();
        var ui = buildDialogUI(dialog, selectionAllImages, selectionHasTextItem);
        var state = createDialogState(doc, eligibleItems, ui);
        /* ∧∨と↑↓キーは同じ処理で増減し、増減後にプレビューを更新する / steppers and arrow keys share one path */
        var steppedInputs = [ui.options.marginInput, ui.corner.cornerRadiusInput];
        for (var i = 0; i < steppedInputs.length; i++) {
            steppedInputs[i].stepOptions.onStep = state.renderPreview;
            bindSteppedArrowKeys(steppedInputs[i], steppedInputs[i].stepperGroup);
        }
        syncMarginEnabled(ui);
        syncOrderEnabled(ui);
        bindDialogEvents(dialog, ui, state);
        state.syncDimAvailability();
        state.syncDimming();

        dialog.onClose = function () {
            state.cleanupPreview();
            state.restoreOpacity();
        };

        var dialogResult = null;
        ui.buttons.okButton.onClick = function () {
            dialogResult = state.buildSettings(false);
            state.undoPreview();
            dialog.close();
        };
        ui.buttons.cancelButton.onClick = function () {
            dialog.close();
        };

        dialog.show();
        return dialogResult;
    }

    // ==============================
    // 角丸エフェクト / Round Corners effect
    // ==============================
    // 生成した長方形に「角を丸くする」ライブエフェクトを適用する（半径は pt）
    function applyCornerRadius(rect, radiusPt) {
        if (!radiusPt || radiusPt <= 0) return;
        try {
            var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>';
            rect.applyEffect(xml);
        } catch (ignoreCornerApply) { }
    }

    // ==============================
    // 塗り・線の適用 / Fill & stroke
    // ==============================
    // 生成した長方形そのものにプリセットの塗り・線・不透明度を適用する（元の選択オブジェクトには影響しない）
    function applyRectAppearance(rect, preset, isRgbDocument) {
        if (!preset.applyAppearance) {
            return; // 直前の塗り・線：作成時の既定値のままにする
        }
        rect.filled = preset.filled;
        rect.stroked = preset.stroked;
        if (preset.filled) {
            rect.fillColor = makeBlack(isRgbDocument);
        }
        if (preset.stroked) {
            rect.strokeColor = makeBlack(isRgbDocument);
            rect.strokeWidth = strokeUnitValueToPoints(preset.strokeWidth);
        }
        if (typeof preset.opacity === "number") {
            rect.opacity = preset.opacity;
        }
    }

    // ==============================
    // 長方形の作成 / Rectangle creation
    // ==============================

    // オブジェクトの外接矩形を取得する
    // geometricBounds: パス本体の境界 / visibleBounds: 線幅・効果を含むプレビュー境界
    function getItemBounds(item, useVisibleBounds) {
        return useVisibleBounds ? item.visibleBounds : item.geometricBounds;
    }

    // テキストをアウトライン化して計測した境界をキャッシュする
    // 各エントリ: { item: TextFrame参照, geometric: [...], visible: [...] }
    function findOutlinedBoundsCache(outlinedBoundsCache, item) {
        for (var i = 0; i < outlinedBoundsCache.length; i++) {
            if (outlinedBoundsCache[i].item === item) return outlinedBoundsCache[i];
        }
        return null;
    }

    // 計測用の外接矩形を取得する
    // テキスト＋「テキストをアウトライン化」設定時は、複製をアウトライン化して計測し複製は破棄する
    // 一度測定した結果は渡された outlinedBoundsCache に保存し、以降は再アウトライン化せず使い回す
    function measureItemBounds(item, settings, outlinedBoundsCache) {
        if (settings.outlineText && item.typename === "TextFrame") {
            var cached = findOutlinedBoundsCache(outlinedBoundsCache, item);
            if (cached) {
                return settings.useVisibleBounds ? cached.visible : cached.geometric;
            }
            var duplicatedText = null;
            var outlinedGroup = null;
            try {
                duplicatedText = item.duplicate();
                outlinedGroup = duplicatedText.createOutline(); // 成功すると duplicatedText は消費され GroupItem になる
                // useVisibleBounds が後で切り替わっても再計測しなくて済むよう両方を保存する
                var geometricBounds = getItemBounds(outlinedGroup, false);
                var visibleBounds = getItemBounds(outlinedGroup, true);
                outlinedGroup.remove();
                outlinedBoundsCache.push({ item: item, geometric: geometricBounds, visible: visibleBounds });
                return settings.useVisibleBounds ? visibleBounds : geometricBounds;
            } catch (e) {
                // 計測用の複製・アウトラインがドキュメントに残らないよう後始末する
                if (outlinedGroup) {
                    try { outlinedGroup.remove(); } catch (ignoreOutlined) { }
                } else if (duplicatedText) {
                    try { duplicatedText.remove(); } catch (ignoreDuplicate) { }
                }
            }
        }
        return getItemBounds(item, settings.useVisibleBounds);
    }

    // 複数オブジェクトをまとめた外接矩形を返す / Combined bounding box of multiple objects
    function getCombinedBounds(pageItems, settings, outlinedBoundsCache) {
        if (pageItems.length === 0) return null;
        var first = measureItemBounds(pageItems[0], settings, outlinedBoundsCache);
        var combined = [first[0], first[1], first[2], first[3]];
        for (var i = 1; i < pageItems.length; i++) {
            var b = measureItemBounds(pageItems[i], settings, outlinedBoundsCache);
            if (b[0] < combined[0]) combined[0] = b[0];
            if (b[1] > combined[1]) combined[1] = b[1];
            if (b[2] > combined[2]) combined[2] = b[2];
            if (b[3] < combined[3]) combined[3] = b[3];
        }
        return combined;
    }

    // 指定した外接矩形から長方形を作成し、プリセットの塗り・線を適用する
    function createRectFromBounds(doc, bounds, settings, referenceItem) {
        // マージン分だけ境界を外側へ拡張（負値なら内側に縮む）
        var margin = (typeof settings.margin === "number") ? settings.margin : 0;
        var rectLeft = bounds[0] - margin;
        var rectTop = bounds[1] + margin;
        var rectRight = bounds[2] + margin;
        var rectBottom = bounds[3] - margin;

        var rectWidth = rectRight - rectLeft;
        var rectHeight = rectTop - rectBottom;
        if (rectWidth <= 0 || rectHeight <= 0) {
            return null;
        }

        var rect = doc.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
        // keep モードのみ「前面/背面」を反映。mask は後でグループ内に再配置、delete は元の位置に置く
        var placement;
        if (settings.originalMode === "keep") {
            placement = settings.placeInFront ? ElementPlacement.PLACEBEFORE : ElementPlacement.PLACEAFTER;
        } else {
            placement = ElementPlacement.PLACEBEFORE;
        }
        // move() の戻り値は環境により undefined になるため、元の参照をそのまま使う
        rect.move(referenceItem, placement);

        // 塗り・線は生成した長方形に直接適用する（選択中の元オブジェクトには触れない）
        var isRgbDocument = (doc.documentColorSpace === DocumentColorSpace.RGB);
        applyRectAppearance(rect, settings.appearancePreset, isRgbDocument);
        applyCornerRadius(rect, settings.cornerRadius);
        return rect;
    }

    // ==============================
    // 元のオブジェクトの処理 / Handling the original object
    // ==============================

    // 元オブジェクト群のうち最前面（z 順で最も上）の項目を返す
    function findTopMostItem(items) {
        var top = items[0];
        for (var i = 1; i < items.length; i++) {
            if (items[i].absoluteZOrderPosition > top.absoluteZOrderPosition) {
                top = items[i];
            }
        }
        return top;
    }

    // 長方形をクリッピングマスクとして適用し、マスクグループを返す
    function applyClippingMask(doc, clipRect, contentItems) {
        // 元の最前面項目の位置にグループを差し込み、元の z 順序を維持する
        var topMost = findTopMostItem(contentItems);
        var clipGroup = doc.groupItems.add();
        clipGroup.move(topMost, ElementPlacement.PLACEBEFORE);
        // 元オブジェクトをグループへ移動
        for (var i = 0; i < contentItems.length; i++) {
            contentItems[i].move(clipGroup, ElementPlacement.PLACEATEND);
        }
        // クリップパス（長方形）は最前面に置く
        clipRect.move(clipGroup, ElementPlacement.PLACEATBEGINNING);
        clipRect.clipping = true;
        clipGroup.clipped = true;
        return clipGroup;
    }

    // 元のオブジェクトの扱い（そのまま／マスク／削除）を適用し、選択対象の項目を返す
    function applyOriginalMode(doc, rect, originalItems, originalMode) {
        if (originalMode === "mask") {
            return applyClippingMask(doc, rect, originalItems);
        }
        if (originalMode === "delete") {
            for (var i = 0; i < originalItems.length; i++) {
                originalItems[i].remove();
            }
            return rect;
        }
        return rect; // そのまま / keep
    }

    // ==============================
    // ロック・非表示判定 / Locked & hidden checks
    // ==============================

    // 親レイヤー・親グループを含めてロック状態かどうかを返す
    function isItemLocked(item) {
        var current = item;
        while (current && current.typename !== "Document") {
            if (current.locked) return true;
            current = current.parent;
        }
        return false;
    }

    // 親レイヤー・親グループを含めて非表示状態かどうかを返す
    function isItemHidden(item) {
        var current = item;
        while (current && current.typename !== "Document") {
            if (current.hidden) return true;
            current = current.parent;
        }
        return false;
    }

    // ==============================
    // メイン処理のステップ / Main steps
    // ==============================

    // doc.selection を配列にコピー（後段で selection が書き換わっても参照を保持するため）
    function collectSelectedItems(doc) {
        var items = [];
        for (var i = 0; i < doc.selection.length; i++) {
            items.push(doc.selection[i]);
        }
        return items;
    }

    // ロック・非表示オブジェクトを除いた処理対象を返す
    function filterEligibleItems(items) {
        var eligible = [];
        for (var i = 0; i < items.length; i++) {
            if (!isItemLocked(items[i]) && !isItemHidden(items[i])) {
                eligible.push(items[i]);
            }
        }
        return eligible;
    }

    // 元オブジェクト群のうち最背面（z 順で最も下）の項目を返す
    function findBottomMostItem(items) {
        var bottom = items[0];
        for (var i = 1; i < items.length; i++) {
            if (items[i].absoluteZOrderPosition < bottom.absoluteZOrderPosition) {
                bottom = items[i];
            }
        }
        return bottom;
    }

    // 選択範囲全体をまとめて1つの長方形（またはマスクグループ）にする
    function produceGroupedRectangle(doc, items, settings, outlinedBoundsCache) {
        var combinedBounds = getCombinedBounds(items, settings, outlinedBoundsCache);
        if (!combinedBounds) return [];

        // 「前面に作成」は選択中の最前面項目、「背面に作成」は最背面項目を基準にする
        var referenceItem = settings.placeInFront ? findTopMostItem(items) : findBottomMostItem(items);
        var groupRect = createRectFromBounds(doc, combinedBounds, settings, referenceItem);
        if (!groupRect) return [];

        return [applyOriginalMode(doc, groupRect, items, settings.originalMode)];
    }

    // オブジェクトごとに長方形（またはマスクグループ）を作成する
    function produceIndividualRectangles(doc, items, settings, outlinedBoundsCache) {
        var created = [];
        for (var i = 0; i < items.length; i++) {
            var itemBounds = measureItemBounds(items[i], settings, outlinedBoundsCache);
            var rect = createRectFromBounds(doc, itemBounds, settings, items[i]);
            if (rect) {
                created.push(applyOriginalMode(doc, rect, [items[i]], settings.originalMode));
            }
        }
        return created;
    }

    // 設定に従って長方形（またはマスクグループ）を生成し、結果配列を返す
    function produceRectangles(doc, items, settings, outlinedBoundsCache) {
        if (!outlinedBoundsCache) outlinedBoundsCache = [];

        if (settings.groupAsOne) {
            return produceGroupedRectangle(doc, items, settings, outlinedBoundsCache);
        }
        return produceIndividualRectangles(doc, items, settings, outlinedBoundsCache);
    }

    // 表示切替コマンドを実行し、成功したかどうかを返す / Toggle a display command and return whether it succeeded
    function toggleDisplayCommand(commandName) {
        try {
            app.executeMenuCommand(commandName);
            return true;
        } catch (ignoreDisplayToggle) {
            return false;
        }
    }

    // ==============================
    // メイン処理 / Main
    // ==============================
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("noDocument"));
            return;
        }
        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel("noSelection"));
            return;
        }

        var selectedItems = collectSelectedItems(doc);
        var eligibleItems = filterEligibleItems(selectedItems);
        if (eligibleItems.length === 0) {
            alert(getLabel("noResult"));
            return;
        }

        var didToggleBoundingBox = toggleDisplayCommand('AI Bounding Box Toggle');
        var didToggleEdges = toggleDisplayCommand('edge');
        try {
            var settings = showSettingsDialog(doc, eligibleItems, isSelectionAllImages(eligibleItems), selectionHasText(eligibleItems));
            if (!settings) {
                return; // キャンセル / Cancelled
            }

            var createdItems = produceRectangles(doc, eligibleItems, settings);

            if (createdItems.length === 0) {
                alert(getLabel("noResult"));
                return;
            }

            // 作成した長方形（マスク時はマスクグループ）のみを選択
            doc.selection = createdItems;
        } finally {
            if (didToggleEdges) toggleDisplayCommand('edge');
            if (didToggleBoundingBox) toggleDisplayCommand('AI Bounding Box Toggle');
        }
    }

    main();

})();
