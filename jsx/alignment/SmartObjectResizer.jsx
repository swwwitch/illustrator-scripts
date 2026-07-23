#targetengine "SOR_Engine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

- 選択中のオブジェクトを、指定した基準（最大／最小／指定サイズ／基準辺／面積／アートボード／裁ち落とし）に基づいて柔軟にリサイズできるIllustrator用スクリプトです。
- 縦横比保持・片辺のみモード、横位置／縦位置の整列、リアルタイムプレビューを備えています。ダイアログを開いた直後はどの基準も選択されておらず変形は起こらず、基準ラジオをクリックした時点でリサイズが始まります。整列はリサイズと同じ境界基準で計算します（「プレビュー境界で計測」ON＝visibleBounds／OFF＝geometricBounds）。基準を切り替えても、選択中の基準と整列は保持したまま再計算します。

### 主な機能

- 縦横比保持と片辺のみの切り替え
- 各種基準（最大／最小／指定サイズ／基準辺／面積／アートボード／裁ち落とし）でのスケーリング
- テキストをアウトライン境界で計測（複製→分割→アウトライン→計測→即削除）
- 整列（右カラムの「整列（横）」「整列（縦）」パネル）
  - 整列（横）パネル：左／中央／右（X座標）＋ 均等／0間隔（Y方向の分配）
  - 整列（縦）パネル：上／中央／下（Y座標）＋ 均等／0間隔（X方向の分配）
  - 横位置と縦位置は同時指定可（横=left・縦=top で独立）。基準を変えても維持される
- 主要オプションにツールチップ（helpTip）を表示
- リアルタイムプレビュー、リセット、フッターの リセット／キャンセル／OK
- 日本語／英語インターフェース対応

### 処理の流れ

1. ダイアログでリサイズ基準（と縦横比／片辺）を選択（開いた直後は未選択＝変形なし。基準を選ぶとリサイズ開始）
2. 選択基準でスケーリング（テキストは複製→分割→アウトライン化→計測し、元オブジェクトに反映）
3. 整列（横位置／縦位置）や面積一致を適用。整列チェックを切り替えると、基準状態に戻してから両軸の整列を再適用
4. OKで確定、キャンセル／リセットで元に戻す

----

### Script Name:

SmartObjectResizer.jsx

### Overview

- An Illustrator script that flexibly resizes selected objects based on a chosen criterion (Max / Min / Fixed Size / Ref. side / Area / Artboard / Bleed).
- Supports Keep Aspect / One Side Only modes, horizontal- and vertical-position alignment, and real-time preview. No base is selected when the dialog opens, so nothing is transformed until you click a base radio. Alignment uses the same bounds basis as resizing ("Measure by preview bounds" ON = visibleBounds / OFF = geometricBounds). Switching the base keeps the current criterion and alignment and recomputes them.

### Main Features

- Toggle between Keep Aspect and One Side Only
- Scaling based on Max / Min / Fixed Size / Ref. side / Area / Artboard / Bleed
- Text measured by outline bounds (duplicate → expand → outline → measure → immediate cleanup)
- Alignment (the "Align (H)" / "Align (V)" panels in the right column)
  - Align (H) panel: Left / Center / Right (X) + Distribute evenly / Zero gap (along Y)
  - Align (V) panel: Top / Middle / Bottom (Y) + Distribute evenly / Zero gap (along X)
  - Horizontal and vertical can be combined (independent: H = left, V = top); preserved across base changes
- Tooltips (helpTips) on key options
- Real-time preview, Reset, and a footer with Reset / Cancel / OK
- Japanese and English UI support

### Process Flow

1. Choose the resize base (and Keep Aspect / One Side Only) in the dialog (nothing is selected on open = no transform; picking a base starts resizing)
2. Scale by the selected base (text is duplicated → expanded → outlined → measured, then applied back to originals)
3. Apply alignment (horizontal / vertical) or area matching; toggling an alignment restores the base state and re-applies both axes
4. Confirm with OK, or revert with Cancel / Reset

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartObjectResizer";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-22";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectResizer.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectResizer.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6f35bd4000ec"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* UIレイアウトの余白・間隔（必要に応じて調整） / UI layout margins & spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 15;                 /* 2カラムの間隔 / gap between columns */
    var LABEL_WIDTH    = 90;                 /* 行ラベルの共通幅 / shared row-label width */
    var BLEED_OFFSET_PT = 6 * 2.83464567;    /* 裁ち落とし: 片側3mm＝幅/高さそれぞれ+6mm / bleed 3mm per side */

    /* ダイアログ位置をセッション内で記憶（Illustrator終了でリセット） */
    /* Remember dialog position within this Illustrator session (resets on quit) */
    var __SOR_SESSION_KEY = "SmartObjectResizer_dialogPos";
    if (typeof $.global[__SOR_SESSION_KEY] === "undefined") {
        $.global[__SOR_SESSION_KEY] = null;
    }

    /* 共通レイアウト適用ヘルパー / Shared layout helpers */
    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 縦方向のスペーサー（区切り用の空グループ） / Vertical spacer group */
    function addSpacer(parent, height) {
        var spacer = parent.add("group");
        spacer.minimumSize.height = height;
        spacer.maximumSize.height = height;
        return spacer;
    }

    /* helpTip（ツールチップ）を1コントロールまたはコントロール配列へ設定 */
    /* Set helpTip on a single control or an array of controls */
    function setHelpTip(target, tip) {
        if (!tip || !target) return;
        if (typeof target !== "string" && typeof target.length === "number") {
            for (var i = 0; i < target.length; i++) {
                if (target[i]) target[i].helpTip = tip;
            }
        } else {
            target.helpTip = tip;
        }
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var currentLanguage = getCurrentLang();

    /*
    LABELS のカテゴリ規約 / Category rules
      dialog   : ダイアログタイトル / dialog title
      panel    : パネル見出し / panel headers
      field    : 行ラベル（コロンは含めず、描画時に fieldLabel() が付与） / row field labels (colon added by fieldLabel())
      radio    : ラジオボタン / radio buttons
      checkbox : チェックボックス / checkboxes
      button   : ボタン（Cancel / Reset。OK は非ローカライズの "OK" 直書き）
      alert    : 警告メッセージ / alerts
    記号は {colon} / {slash} / {comma} / {openParen} / {closeParen} で記述し、
    applyUISymbols() が言語に応じた全角/半角へ展開する。
    */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトのリサイズ", en: "SmartObjectResizer" }
        },
        panel: {
            base:  { ja: "リサイズ基準", en: "Resize base" },
            // 整列（横）: 左/中央/右 ＋ 縦方向の分配 / horizontal alignment
            hAlign: { ja: "整列{openParen}横{closeParen}", en: "Align {openParen}H{closeParen}" },
            // 整列（縦）: 上/中央/下 ＋ 横方向の分配 / vertical alignment
            vAlign: { ja: "整列{openParen}縦{closeParen}", en: "Align {openParen}V{closeParen}" }
        },
        field: {
            max:      { ja: "最大",         en: "Max" },
            min:      { ja: "最小",         en: "Min" },
            fixed:    { ja: "指定サイズ",   en: "Fixed Size" },
            base:     { ja: "基準辺",       en: "Ref. side" },
            area:     { ja: "面積",         en: "Area" },
            artboard: { ja: "アートボード", en: "Artboard" },
            bleed:    { ja: "裁ち落とし",   en: "Bleed" }
        },
        radio: {
            keepAspect:  { ja: "縦横比保持", en: "Keep aspect" },
            oneSideOnly: { ja: "片辺のみ",   en: "One side only" },
            width:       { ja: "幅",         en: "Width" },
            height:      { ja: "高さ",       en: "Height" },
            longSide:    { ja: "長辺",       en: "Long side" },
            shortSide:   { ja: "短辺",       en: "Short side" },
            areaMax:     { ja: "最大",       en: "Max" },
            areaMin:     { ja: "最小",       en: "Min" }
        },
        checkbox: {
            textOutlineBounds: { ja: "テキストをアウトライン境界で計測", en: "Measure text by outline bounds" },
            previewBounds:     { ja: "プレビュー境界で計測", en: "Measure by preview bounds" },
            alignLeft:   { ja: "左",     en: "Left" },
            alignCenter: { ja: "中央",   en: "Center" },
            alignRight:  { ja: "右",     en: "Right" },
            alignEven:   { ja: "均等",   en: "Distribute evenly" },
            alignZero:   { ja: "0間隔",  en: "Zero gap" },
            alignTop:    { ja: "上",   en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下",   en: "Bottom" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset:  { ja: "リセット",   en: "Reset" }
        },
        // ツールチップ（helpTip）: 意味が自明でないコントロールにのみ設定 / Tooltips for non-obvious controls
        tooltip: {
            keepAspect:  { ja: "縦横比を保ったまま拡大／縮小します。", en: "Scale while keeping the aspect ratio." },
            oneSideOnly: { ja: "幅または高さの片辺だけを変更します（縦横比は保持しません）。", en: "Change only one side (width or height); aspect ratio is not preserved." },
            base:        { ja: "基準となる辺（長辺／短辺）の長さに合わせて、各オブジェクトをリサイズします。", en: "Resize each object to match the chosen reference side (long / short)." },
            area:        { ja: "選択オブジェクトの面積を、最大／最小のものにそろえます。", en: "Match object areas to the largest / smallest in the selection." },
            artboard:    { ja: "選択全体をアートボードの幅／高さに合わせ、中央に配置します。", en: "Fit the whole selection to the artboard width / height and center it." },
            bleed:       { ja: "アートボード＋裁ち落とし（片側3mm）の幅／高さに合わせ、中央に配置します。", en: "Fit to the artboard plus bleed (3mm per side) and center it." },
            alignEven:   { ja: "オブジェクトの間隔が均等になるように分配します（3つ以上で有効）。", en: "Distribute objects with equal gaps (needs 3+ objects)." },
            alignZero:   { ja: "オブジェクトを間隔0で隙間なく並べます（2つ以上で有効）。", en: "Place objects with zero gap, no spacing (needs 2+ objects)." },
            textOutline: { ja: "テキストをアウトライン化した実際の字形の境界で計測します。", en: "Measure text by the actual outlined glyph bounds." },
            preview:     { ja: "線幅や効果を含むプレビュー境界で計測します（オフは幾何境界）。", en: "Measure by preview bounds incl. strokes / effects (off = geometric bounds)." },
            reset:       { ja: "サイズ・位置・整列をすべて元の状態に戻します。", en: "Revert size, position, and alignment to the original state." }
        },
        alert: {
            noDocument:   { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectObject: { ja: "オブジェクトを選択してください。", en: "Please select an object." }
        }
    };

    // "panel.base" のようなドット区切りパスで LABELS から文字列を取得
    // Resolve a dotted path like "panel.base" from LABELS
    function L(path) {
        var parts = path.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            node = node[parts[i]];
            if (!node) return path;
        }
        var text = node[currentLanguage] || node.en;
        if (!text) return path;
        return applyUISymbols(text);
    }

    // 行ラベルに末尾コロンを付与（コロンのロケール差は uiSymbol("colon") に一元化）
    // Field row label with a trailing locale-aware colon (full-width JA ： / half-width EN :)
    function fieldLabel(path) {
        return L(path) + uiSymbol("colon");
    }

    // LABELS 内の {colon} などのプレースホルダを言語別記号へ展開
    // Expand {colon}-style placeholders into locale-specific symbols
    function applyUISymbols(text) {
        return text
            .replace(/\{colon\}/g,      uiSymbol("colon"))
            .replace(/\{slash\}/g,      uiSymbol("slash"))
            .replace(/\{comma\}/g,      uiSymbol("comma"))
            .replace(/\{openParen\}/g,  uiSymbol("openParen"))
            .replace(/\{closeParen\}/g, uiSymbol("closeParen"));
    }

    // 言語別の記号セット（日本語は全角、英語は半角）
    // Locale symbol set (full-width for JA, half-width for EN)
    function uiSymbol(name) {
        if (currentLanguage === "ja") {
            switch (name) {
                case "slash":      return "／";
                case "colon":      return "：";
                case "comma":      return "、";
                case "openParen":  return "（";
                case "closeParen": return "）";
            }
        }
        switch (name) {
            case "slash":      return "/";
            case "colon":      return ":";
            case "comma":      return ", ";
            case "openParen":  return "(";
            case "closeParen": return ")";
        }
        return "";
    }

    // =========================================
    // 単位 / Units
    // =========================================

    if (app.documents.length === 0) {
        alert(L("alert.noDocument"));
        return;
    }
    var doc = app.activeDocument;

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

    // =========================================
    // 選択オブジェクトと初期状態 / Selection & initial state
    // =========================================

    if (!doc.selection || doc.selection.length === 0) {
        alert(L("alert.selectObject"));
        return;
    }
    // doc.selection はライブ参照になりうるため、配列にコピーして固定する
    var originalSelectedItems = [];
    for (var i = 0; i < doc.selection.length; i++) {
        originalSelectedItems.push(doc.selection[i]);
    }
    var workingItems = originalSelectedItems;

    var originalStates = [];
    for (var i = 0; i < originalSelectedItems.length; i++) {
        var item = originalSelectedItems[i];
        originalStates.push({
            item: item,
            width: item.width,
            height: item.height,
            left: item.left,
            top: item.top
        });
    }
    var resizeBaseStates = [];

    var outlineBoundsCache = {};
    var outlineBoundsCacheSeq = 1;
    var outlineIdMap = [];

    var dialog = new Window("dialog", L("dialog.title") + " " + SCRIPT_VERSION);
    dialog.alignChildren = ["left", "top"];
    // 左右インセットと下余白はここで一括管理（各ペイン／フッターの左右マージンは 0）
    dialog.margins = [20, 0, 20, 20];

    // 前回のダイアログ位置を復元（セッション内のみ）
    var savedDialogPos = $.global[__SOR_SESSION_KEY];
    if (savedDialogPos && savedDialogPos.length === 2 && !isNaN(savedDialogPos[0]) && !isNaN(savedDialogPos[1])) {
        dialog.location = savedDialogPos;
    }

    // --- 初期表示時の状態リセットとプレビュー ---
    // 基準は初期未選択なので applyResizeBySelection() は何もしない（＝開いた直後は変形しない）。
    // 復元とUI状態の初期化のみ行う。基準が選択済みの場合のみ、そのモードを1回だけ反映する。
    dialog.onShow = function () {
        // UI 状態の初期化（DOM に触れないので保護不要。失敗したら顕在化させる）
        resetAlignChecks();
        updateInputState();
        updateRadioGroupStates();
        updateTextOutlineOptionState();

        // 初期プレビュー（DOM 操作）のみ保護。失敗しても表示は継続するが、黙殺せずログに残す
        try {
            clearOutlineBoundsCache();
            restoreOriginalGeometry();
            restoreOriginalPosition();
            app.redraw();
            applyResizeBySelection(); // 基準が選択済みの場合のみ、そのモードを1回だけ適用（未選択なら何もしない）
        } catch (e) {
            $.writeln("SmartObjectResizer: 初期プレビューに失敗 / initial preview failed — " + e);
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
    // 縦横比・片辺選択ラジオボタン（columnsGroupより前に配置）
    var ratioGroup = dialog.add("group");
    ratioGroup.orientation = "row";
    ratioGroup.alignChildren = ["left", "center"];
    ratioGroup.margins = [20, 20, 0, 0];
    ratioGroup.alignment = ["center", "top"]; // 中央揃え

    // 「縦横比保持」「片辺のみ」ラジオボタン
    var keepRatioRadio = ratioGroup.add("radiobutton", undefined, L("radio.keepAspect"));
    var oneSideOnlyRadio = ratioGroup.add("radiobutton", undefined, L("radio.oneSideOnly"));
    // ラジオボタン間のスペースを明示的に設定
    keepRatioRadio.margins = [0, 0, 20, 0]; // 右側に20px相当の余白
    oneSideOnlyRadio.margins = [0, 0, 0, 0];

    keepRatioRadio.value = true; // デフォルトで「縦横比保持」を選択
    setHelpTip(keepRatioRadio, L("tooltip.keepAspect"));
    setHelpTip(oneSideOnlyRadio, L("tooltip.oneSideOnly"));

    // 「片辺のみ」でディムされる基準（基準辺／面積／アートボード／裁ち落とし）の選択を解除する。
    // ディムするだけだと value が残り、getSelectedResizeMode() が enabled を見ないため
    // 「片辺のみのはずが縦横比保持でリサイズされる」状態になる。
    // Clear bases that "one side only" dims — leaving them checked would keep them active.
    function clearDimmedBaseSelections() {
        var dimmedGroups = [baseRadios, areaRadios, artboardRadios, bleedRadios];
        for (var g = 0; g < dimmedGroups.length; g++) {
            for (var r = 0; r < dimmedGroups[g].length; r++) {
                dimmedGroups[g][r].value = false;
            }
        }
    }

    // 縦横比保持／片辺のみトグルの共通処理（有効・無効の切替は updateRadioGroupStates() が担当）
    // Shared handler for the keep-aspect / one-side-only toggle.
    function onRatioModeChanged(useOneSideOnly) {
        keepRatioRadio.value = !useOneSideOnly;
        oneSideOnlyRadio.value = useOneSideOnly;

        if (useOneSideOnly) clearDimmedBaseSelections();

        // 指定サイズセクション全体（ラジオ＋テキストフィールド）はどちらのモードでも有効
        if (createRadioGroup.sizeInput && createRadioGroup.sizeInput.parent) {
            createRadioGroup.sizeInput.parent.enabled = true;   // 入力行
        }
        if (createRadioGroup.widthRadio && createRadioGroup.widthRadio.parent) {
            createRadioGroup.widthRadio.parent.enabled = true;  // ラジオボタン行
        }

        reapplyCurrentSelection(); // 内部で updateRadioGroupStates() を呼ぶ
    }

    oneSideOnlyRadio.onClick = function () { onRatioModeChanged(true); };
    keepRatioRadio.onClick   = function () { onRatioModeChanged(false); };

    var columnsGroup = dialog.add("group");
    columnsGroup.orientation = "row";
    // 左右ペインを上揃えに
    columnsGroup.alignChildren = ["top", "top"];
    // 2カラム（左ペイン／右ペイン）の間隔
    columnsGroup.spacing = COLUMN_SPACING;

    // 左ペイン（各種設定）
    var leftPane = columnsGroup.add("group");
    leftPane.orientation = "column";
    leftPane.alignChildren = ["left", "top"];
    // 余白は dialog.margins（左右）と columnsGroup.spacing（カラム間）で管理


    // resizeBasePanel を左ペイン内にパネルとして追加
    var resizeBasePanel = leftPane.add("panel", undefined, L("panel.base"));
    setupPanel(resizeBasePanel);

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

    // 単純な「ラベル＋ラジオ複数」の1行を作る
    function createRadioGroup(label, optionLabels, parent) {
        var group = parent.add("group");
        group.orientation = "row";
        group.alignChildren = ["left", "center"];
        var labelText = group.add("statictext", undefined, label);
        labelText.preferredSize.width = LABEL_WIDTH;
        var buttons = [];
        for (var i = 0; i < optionLabels.length; i++) {
            var radioButton = group.add("radiobutton", undefined, optionLabels[i]);
            buttons.push(radioButton);
            allRadioButtons.push(radioButton);
        }
        return buttons;
    }

    // 指定サイズ用: 「ラベル＋幅/高さラジオ」行と「数値入力＋単位」行の2段構成
    // sizeInput 等は従来どおり createRadioGroup 上に保持して外部から参照する
    function createFixedSizeGroup(label, optionLabels, parent) {
        // 1行目：ラベルとラジオボタン
        var sizeHeaderGroup = parent.add("group");
        sizeHeaderGroup.orientation = "row";
        sizeHeaderGroup.alignChildren = ["left", "center"];
        var sizeLabel = sizeHeaderGroup.add("statictext", undefined, label);
        sizeLabel.preferredSize.width = LABEL_WIDTH;
        var widthRadio = sizeHeaderGroup.add("radiobutton", undefined, optionLabels[0]);
        var heightRadio = sizeHeaderGroup.add("radiobutton", undefined, optionLabels[1]);
        allRadioButtons.push(widthRadio, heightRadio);

        // 2行目：空ラベルと入力欄
        var sizeInputGroup = parent.add("group");
        sizeInputGroup.orientation = "row";
        sizeInputGroup.alignChildren = ["left", "center"];
        var labelSpacer = sizeInputGroup.add("statictext", undefined, ""); // 空ラベル
        labelSpacer.preferredSize.width = LABEL_WIDTH;

        // 選択オブジェクトの平均幅を初期値に
        var totalWidth = 0;
        for (var i = 0; i < originalSelectedItems.length; i++) {
            totalWidth += getReferenceBounds(originalSelectedItems[i], true).width;
        }
        var avgWidth = originalSelectedItems.length > 0 ? (totalWidth / originalSelectedItems.length) : 100;
        var sizeInput = sizeInputGroup.add("edittext", undefined, avgWidth.toFixed(0));
        sizeInput.characters = 5;
        changeValueByArrowKey(sizeInput, false);
        var sizeUnit = sizeInputGroup.add("statictext", undefined, unitLabel);

        sizeInput.onChange = function () {
            if ((widthRadio.value || heightRadio.value) && !isNaN(parseFloat(sizeInput.text))) {
                clearOutlineBoundsCache();
                restoreOriginalGeometry();
                restoreOriginalPosition();
                app.redraw();
                // 片辺のみ／縦横比保持のどちらも applyResizeBySelection() に集約
                // （指定サイズの単位換算・片辺スケールは同関数側で処理）
                applyResizeBySelection();
            }
        };

        // 外部参照用に保持（updateInputState / computeReferenceValue などが参照）
        createRadioGroup.sizeInput = sizeInput;
        createRadioGroup.widthRadio = widthRadio;
        createRadioGroup.heightRadio = heightRadio;
        createRadioGroup.sizeUnit = sizeUnit;
        return [widthRadio, heightRadio];
    }

    // 新しい順序でラジオボタンとグループを作成
    // 1. 最大
    var maxRadios = createRadioGroup(fieldLabel("field.max"), [L("radio.width"), L("radio.height")], resizeBasePanel);
    // 2. 最小
    var minRadios = createRadioGroup(fieldLabel("field.min"), [L("radio.width"), L("radio.height")], resizeBasePanel);
    // 3. 指定サイズ（ラジオ＋数値欄一体）
    var fixedRadios = createFixedSizeGroup(fieldLabel("field.fixed"), [L("radio.width"), L("radio.height")], resizeBasePanel);
    // 5. 基準辺
    var baseRadios = createRadioGroup(fieldLabel("field.base"), [L("radio.longSide"), L("radio.shortSide")], resizeBasePanel);
    // 6. 面積
    var areaRadios = createRadioGroup(fieldLabel("field.area"), [L("radio.areaMax"), L("radio.areaMin")], resizeBasePanel);
    // 7. --- ディバイダー ---
    var dividerLine = resizeBasePanel.add("statictext", undefined, "  ───────────────  ");
    // 8. アートボード
    var artboardRadios = createRadioGroup(fieldLabel("field.artboard"), [L("radio.width"), L("radio.height")], resizeBasePanel);
    // 9. 裁ち落とし
    var bleedRadios = createRadioGroup(fieldLabel("field.bleed"), [L("radio.width"), L("radio.height")], resizeBasePanel);

    // radioGroupsの並び順も新順に
    // [最大, 最小, 指定サイズ, 基準辺, 面積, アートボード, 裁ち落とし]
    radioGroups.push(maxRadios, minRadios, fixedRadios, baseRadios, areaRadios, artboardRadios, bleedRadios);

    // 意味が自明でない基準にツールチップを設定（最大／最小／指定サイズは自明なため付けない）
    setHelpTip(baseRadios, L("tooltip.base"));
    setHelpTip(areaRadios, L("tooltip.area"));
    setHelpTip(artboardRadios, L("tooltip.artboard"));
    setHelpTip(bleedRadios, L("tooltip.bleed"));

    // --- 整列チェック群 ---
    // 整列チェックボックスをすべてOFFにする共通関数
    function resetAlignChecks() {
        alignLeftCheck.value = false;
        alignCenterCheck.value = false;
        alignRightCheck.value = false;
        alignTopCheck.value = false;
        alignMiddleCheck.value = false;
        alignBottomCheck.value = false;
        alignEvenCheck.value = false;
        alignEvenZeroCheck.value = false;
        alignHorizontalEvenCheck.value = false;
        alignHorizontalEvenZeroCheck.value = false;
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

    // 現在の選択モードを再適用（ラジオのクリアはしない）
    // Re-apply the current selection WITHOUT clearing any radio.
    // 縦横比／片辺トグルから使う（選択中の基準を保持したまま再計算する）
    function reapplyCurrentSelection() {
        clearOutlineBoundsCache();
        updateInputState();
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
        applyResizeBySelection();
        // チェック中の整列は維持し、新しいサイズに対して再適用する
        reapplyActiveAlignments();
        updateRadioGroupStates();
        // 指定サイズ入力の有効化は updateInputState() 済み。ここではディバイダーのみ制御
        if (typeof dividerLine !== "undefined" && dividerLine) {
            dividerLine.enabled = !oneSideOnlyRadio.value;
        }
    }

    // ラジオボタン選択時の共通処理: 他の基準ラジオをOFFにしてから再適用
    function onAnyRadioClick() {
        for (var i = 0; i < allRadioButtons.length; i++) {
            if (allRadioButtons[i] !== this) {
                allRadioButtons[i].value = false;
            }
        }
        reapplyCurrentSelection();
    }

    // allRadioButtons は createRadioGroup / createFixedSizeGroup が作った基準ラジオのみ。
    // keepRatioRadio / oneSideOnlyRadio は含まれず、個別に onClick を設定済み。
    for (var i = 0; i < allRadioButtons.length; i++) {
        allRadioButtons[i].onClick = onAnyRadioClick;
    }

    // 初期状態ではどの基準も選択しない。
    // → ダイアログを開いた直後は getSelectedResizeMode() が null を返し、変形は一切起こらない。
    //    ユーザーが基準ラジオをクリックした時点で初めてリサイズが実行される。
    // No base is selected initially, so opening the dialog performs no transform;
    // resizing starts only when the user clicks a base radio.

    // 指定サイズ入力欄の有効化制御
    function updateInputState() {
        var widthRadio = createRadioGroup.widthRadio;
        var heightRadio = createRadioGroup.heightRadio;
        var sizeInput = createRadioGroup.sizeInput;
        if (!sizeInput || !sizeInput.parent) return;
        // 指定サイズ（幅/高さ）が選択されているときだけ入力欄を有効化
        sizeInput.enabled = !!((widthRadio && widthRadio.value) || (heightRadio && heightRadio.value));
    }

    // Preview-related checkboxes
    var previewGroup = leftPane.add("group");
    previewGroup.orientation = "column";
    previewGroup.alignChildren = ["left", "top"];
    previewGroup.margins = [0, 5, 0, 0]; // top margin

    // プレビュー系チェック（アウトライン計測／プレビュー境界）切替時の共通処理
    function onPreviewOptionChanged() {
        resetAlignChecks();
        clearOutlineBoundsCache();
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
        applyResizeBySelection();
    }

    var textOutlineBoundsCheck = previewGroup.add("checkbox", undefined, L("checkbox.textOutlineBounds"));
    textOutlineBoundsCheck.value = false;
    textOutlineBoundsCheck.onClick = onPreviewOptionChanged;
    setHelpTip(textOutlineBoundsCheck, L("tooltip.textOutline"));

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

    var previewCheck = previewGroup.add("checkbox", undefined, L("checkbox.previewBounds"));
    previewCheck.value = true;
    previewCheck.onClick = onPreviewOptionChanged;
    setHelpTip(previewCheck, L("tooltip.preview"));

    updateTextOutlineOptionState();

    // 右ペイン（整列）
    var rightPane = columnsGroup.add("group");
    rightPane.orientation = "column";
    rightPane.alignChildren = ["fill", "top"];

    // 横位置パネル（左/中央/右 ＋ 縦方向の分配）: チェックは縦並び
    var hAlignPanel = rightPane.add("panel", undefined, L("panel.hAlign"));
    setupPanel(hAlignPanel, 5);
    var alignLeftCheck = hAlignPanel.add("checkbox", undefined, L("checkbox.alignLeft"));
    var alignCenterCheck = hAlignPanel.add("checkbox", undefined, L("checkbox.alignCenter"));
    var alignRightCheck = hAlignPanel.add("checkbox", undefined, L("checkbox.alignRight"));
    addSpacer(hAlignPanel, 5); // 「均等」の上の余白（整列⇔分配の区切り）
    var alignEvenCheck = hAlignPanel.add("checkbox", undefined, L("checkbox.alignEven"));
    var alignEvenZeroCheck = hAlignPanel.add("checkbox", undefined, L("checkbox.alignZero"));
    setHelpTip(alignEvenCheck, L("tooltip.alignEven"));
    setHelpTip(alignEvenZeroCheck, L("tooltip.alignZero"));
    alignLeftCheck.value = false;
    alignCenterCheck.value = false;
    alignRightCheck.value = false;
    alignEvenCheck.value = false;
    alignEvenZeroCheck.value = false;

    // 縦位置パネル（上/中央/下 ＋ 横方向の分配）: チェックは縦並び
    var vAlignPanel = rightPane.add("panel", undefined, L("panel.vAlign"));
    setupPanel(vAlignPanel, 5);
    var alignTopCheck = vAlignPanel.add("checkbox", undefined, L("checkbox.alignTop"));
    var alignMiddleCheck = vAlignPanel.add("checkbox", undefined, L("checkbox.alignMiddle"));
    var alignBottomCheck = vAlignPanel.add("checkbox", undefined, L("checkbox.alignBottom"));
    addSpacer(vAlignPanel, 5); // 「均等」の上の余白（整列⇔分配の区切り）
    var alignHorizontalEvenCheck = vAlignPanel.add("checkbox", undefined, L("checkbox.alignEven"));
    var alignHorizontalEvenZeroCheck = vAlignPanel.add("checkbox", undefined, L("checkbox.alignZero"));
    setHelpTip(alignHorizontalEvenCheck, L("tooltip.alignEven"));
    setHelpTip(alignHorizontalEvenZeroCheck, L("tooltip.alignZero"));
    alignTopCheck.value = false;
    alignMiddleCheck.value = false;
    alignBottomCheck.value = false;
    alignHorizontalEvenCheck.value = false;
    alignHorizontalEvenZeroCheck.value = false;

    // =========================================
    // 整列ハンドラ（同一軸内で排他） / Alignment handlers (exclusive per axis)
    // =========================================

    // クリック時: ON なら同一軸の兄弟を OFF にし、base状態へ戻してから、
    // 現在チェック中の整列（横軸・縦軸それぞれ最大1つ）を reapplyActiveAlignments() で再適用する。
    // ※横軸=left / 縦軸=top は独立だが、restoreResizeBaseState() は left/top を両方戻すため、
    //   自分の軸だけ再適用するともう一方の軸の整列が消える。常に両軸まとめて適用してこれを防ぐ。
    //   OFF 時も同様に、残った軸の整列を維持するため両軸を再適用する。
    function makeAlignHandler(check, siblings) {
        return function () {
            if (check.value) {
                for (var i = 0; i < siblings.length; i++) siblings[i].value = false;
            }
            restoreResizeBaseState();
            reapplyActiveAlignments();
        };
    }

    // 端揃え・中央揃え（getAlignmentBounds 基準、base状態から実行）
    function alignToMinLeft() {
        var minLeft = null;
        for (var i = 0; i < workingItems.length; i++) {
            var left = getAlignmentBounds(workingItems[i]).left;
            if (minLeft === null || left < minLeft) minLeft = left;
        }
        if (minLeft === null) return;
        for (var i = 0; i < workingItems.length; i++) {
            workingItems[i].left += minLeft - getAlignmentBounds(workingItems[i]).left;
        }
    }
    function alignToCenterX() {
        var sum = 0;
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getAlignmentBounds(workingItems[i]);
            sum += bounds.left + bounds.width / 2;
        }
        var average = sum / workingItems.length;
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getAlignmentBounds(workingItems[i]);
            workingItems[i].left += average - (bounds.left + bounds.width / 2);
        }
    }
    function alignToMaxRight() {
        var maxRight = null;
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getAlignmentBounds(workingItems[i]);
            var right = bounds.left + bounds.width;
            if (maxRight === null || right > maxRight) maxRight = right;
        }
        if (maxRight === null) return;
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getAlignmentBounds(workingItems[i]);
            workingItems[i].left += maxRight - (bounds.left + bounds.width);
        }
    }
    function alignToMaxTop() {
        var maxTop = null;
        for (var i = 0; i < workingItems.length; i++) {
            var top = getAlignmentBounds(workingItems[i]).top;
            if (maxTop === null || top > maxTop) maxTop = top;
        }
        if (maxTop === null) return;
        for (var i = 0; i < workingItems.length; i++) {
            workingItems[i].top += maxTop - getAlignmentBounds(workingItems[i]).top;
        }
    }
    function alignToCenterY() {
        var sum = 0;
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getAlignmentBounds(workingItems[i]);
            sum += bounds.top - bounds.height / 2;
        }
        var average = sum / workingItems.length;
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getAlignmentBounds(workingItems[i]);
            workingItems[i].top += average - (bounds.top - bounds.height / 2);
        }
    }
    function alignToMinBottom() {
        var minBottom = null;
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getAlignmentBounds(workingItems[i]);
            var bottom = bounds.top - bounds.height;
            if (minBottom === null || bottom < minBottom) minBottom = bottom;
        }
        if (minBottom === null) return;
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getAlignmentBounds(workingItems[i]);
            workingItems[i].top += minBottom - (bounds.top - bounds.height);
        }
    }

    // 分配（縦=top降順 / 横=left昇順にソートして順に配置。useGap=false で 0 間隔）
    // ※ item.top / item.left は visibleBounds 基準なので、getAlignmentBounds() の値を
    //   直接代入すると「プレビュー境界で計測」OFF（geometricBounds）のとき線幅の半分ずれる。
    //   端揃え系と同様、必ず差分（delta）で動かすこと。
    //   item.top/left are visible-bounds based; always move by delta so the
    //   distribution follows the same bounds basis as the alignment functions.
    function distributeVertical(useGap) {
        var sortedItems = workingItems.slice(0).sort(function (a, b) {
            return getAlignmentBounds(b).top - getAlignmentBounds(a).top;
        });
        var topMost = getAlignmentBounds(sortedItems[0]).top;
        var gap = 0;
        if (useGap) {
            var lastBounds = getAlignmentBounds(sortedItems[sortedItems.length - 1]);
            var bottomMost = lastBounds.top - lastBounds.height;
            var totalHeight = 0;
            for (var i = 0; i < sortedItems.length; i++) totalHeight += getAlignmentBounds(sortedItems[i]).height;
            gap = (topMost - bottomMost - totalHeight) / (sortedItems.length - 1);
        }
        var currentY = topMost;
        for (var j = 0; j < sortedItems.length; j++) {
            var bounds = getAlignmentBounds(sortedItems[j]);
            sortedItems[j].top += currentY - bounds.top;
            currentY -= (bounds.height + gap);
        }
    }
    function distributeHorizontal(useGap) {
        var sortedItems = workingItems.slice(0).sort(function (a, b) {
            return getAlignmentBounds(a).left - getAlignmentBounds(b).left;
        });
        var leftMost = getAlignmentBounds(sortedItems[0]).left;
        var gap = 0;
        if (useGap) {
            var lastBounds = getAlignmentBounds(sortedItems[sortedItems.length - 1]);
            var rightMost = lastBounds.left + lastBounds.width;
            var totalWidth = 0;
            for (var i = 0; i < sortedItems.length; i++) totalWidth += getAlignmentBounds(sortedItems[i]).width;
            gap = (rightMost - leftMost - totalWidth) / (sortedItems.length - 1);
        }
        var currentX = leftMost;
        for (var j = 0; j < sortedItems.length; j++) {
            var bounds = getAlignmentBounds(sortedItems[j]);
            sortedItems[j].left += currentX - bounds.left;
            currentX += (bounds.width + gap);
        }
    }

    // 整列の定義テーブル（軸ごとに排他）: onClick 割り当てと再適用の唯一のソース
    // Alignment table (mutually exclusive within each axis) — drives both the
    // onClick handlers and reapplyActiveAlignments (single source of truth).
    // minItems はその整列が成立する最小オブジェクト数。
    // 「均等」は両端を固定して間を分けるので3個以上、「0間隔」は2個以上で意味を持つ。
    // minItems is the smallest object count for which the entry is meaningful:
    // 3+ for "distribute evenly" (both ends are pinned), 2+ for "zero gap".
    var alignAxes = [
        // 横位置（X座標を変更）: 左 / 中央 / 右 / 横均等 / 横0
        [
            { check: alignLeftCheck,               minItems: 1, apply: alignToMinLeft },
            { check: alignCenterCheck,             minItems: 1, apply: alignToCenterX },
            { check: alignRightCheck,              minItems: 1, apply: alignToMaxRight },
            { check: alignHorizontalEvenCheck,     minItems: 3, apply: function () { distributeHorizontal(true); } },
            { check: alignHorizontalEvenZeroCheck, minItems: 2, apply: function () { distributeHorizontal(false); } }
        ],
        // 縦位置（Y座標を変更）: 上 / 中央 / 下 / 均等 / 0
        [
            { check: alignTopCheck,      minItems: 1, apply: alignToMaxTop },
            { check: alignMiddleCheck,   minItems: 1, apply: alignToCenterY },
            { check: alignBottomCheck,   minItems: 1, apply: alignToMinBottom },
            { check: alignEvenCheck,     minItems: 3, apply: function () { distributeVertical(true); } },
            { check: alignEvenZeroCheck, minItems: 2, apply: function () { distributeVertical(false); } }
        ]
    ];

    // onClick を割り当て（同一軸の他チェックを siblings として排他）
    for (var i = 0; i < alignAxes.length; i++) {
        var axis = alignAxes[i];
        for (var j = 0; j < axis.length; j++) {
            var siblings = [];
            for (var k = 0; k < axis.length; k++) {
                if (k !== j) siblings.push(axis[k].check);
            }
            axis[j].check.onClick = makeAlignHandler(axis[j].check, siblings);
        }
    }

    // 基準変更などで再リサイズした後、チェック中の整列を新サイズに対して再適用する
    // 横軸・縦軸は独立（left/top を別々に動かす）ため両方同時に再適用（各軸1つだけ有効）
    function reapplyActiveAlignments() {
        var changed = false;
        for (var i = 0; i < alignAxes.length; i++) {
            var axis = alignAxes[i];
            for (var j = 0; j < axis.length; j++) {
                var entry = axis[j];
                if (entry.check.value && workingItems.length >= entry.minItems) {
                    entry.apply();
                    changed = true;
                    break; // 同一軸は1つだけ
                }
            }
        }
        if (changed) app.redraw();
    }

    // =========================================
    // フッター（左右分割: 左=リセット / スペーサー / 右=キャンセル・OK）
    // Footer (split): Reset (left) / spacer / Cancel・OK (right)
    // =========================================

    // メイングループ（横並び） / Main group (horizontal layout)
    var btnRowGroup = dialog.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.alignment = ["fill", "bottom"];   // 左右下の余白は dialog.margins が担当
    btnRowGroup.margins = [0, 5, 0, 0];           // 整列パネルとの間隔（上のみ）

    // 左側グループ / Left-side button group
    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.alignChildren = ["left", "center"];
    var btnReset = btnLeftGroup.add("button", undefined, L("button.reset"));
    setHelpTip(btnReset, L("tooltip.reset"));
    btnReset.onClick = function () {
        // 基準ラジオ・整列チェック・基準状態をすべて解除して、UI と実状態を初期状態にそろえる。
        // resizeBaseStates を消さないと、この後に整列をクリックしたとき
        // restoreResizeBaseState() がリセット前のリサイズ結果を復元してしまう。
        // Clearing resizeBaseStates is required: otherwise a later alignment click would
        // restore the pre-reset (resized) geometry via restoreResizeBaseState().
        for (var i = 0; i < allRadioButtons.length; i++) {
            allRadioButtons[i].value = false;
        }
        resizeBaseStates = [];
        resetAlignChecks();
        updateInputState();
        clearOutlineBoundsCache();
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
    };

    // スペーサー（伸縮）/ Spacer (stretchable)
    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    // 右側グループ / Right-side button group（Mac 規約で Cancel → OK の順）
    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.alignChildren = ["right", "center"];

    // 確定／破棄の判定は dialog.show() の戻り値に一本化する。
    // ボタンの onClick だけで復元すると、ESC キーやウィンドウの閉じるボタンでは
    // onClick が発火しないため、プレビュー中の変形がそのまま確定してしまう。
    // Deciding commit vs. discard from the return value of show() (not from the button
    // handlers) is what makes ESC / the window close box behave as a real cancel.
    var DIALOG_RESULT_OK = 1;
    var DIALOG_RESULT_CANCEL = 2;

    var btnCancel = btnRightGroup.add("button", undefined, L("button.cancel"), { name: "cancel" });
    btnCancel.onClick = function () {
        dialog.close(DIALOG_RESULT_CANCEL);
    };

    var btnOK = btnRightGroup.add("button", undefined, "OK", { name: "ok" });
    btnOK.onClick = function () {
        // 確定（一時グループを使わないので親階層の復元処理は不要）
        dialog.close(DIALOG_RESULT_OK);
    };

    var dialogResult = dialog.show();
    if (dialogResult !== DIALOG_RESULT_OK) {
        // キャンセル／ESC／ウィンドウを閉じる: プレビュー中の変形をすべて破棄
        restoreOriginalGeometry();
        restoreOriginalPosition();
        app.redraw();
    }
    clearOutlineBoundsCache();

    function getScaleFactor(current, target) {
        return (target / current) * 100;
    }

    // 現在の定規単位1つ分をポイントに換算 / Points per current ruler unit
    function pointsPerUnit() {
        switch (unitLabel) {
            case "mm":   return 2.83464567;
            case "cm":   return 28.3464567;
            case "inch": return 72;
            case "pica": return 12;
            default:     return 1; // pt / px
        }
    }

    // アクティブアートボードの矩形 [left, top, right, bottom]
    function getActiveArtboardRect() {
        return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
    }

    // モードに応じて基準とする1辺の長さを返す（長辺／短辺／幅／高さ）
    function measureSide(bounds, mode) {
        if (mode.isLong) return Math.max(bounds.width, bounds.height);
        if (mode.isShort) return Math.min(bounds.width, bounds.height);
        return mode.isWidth ? bounds.width : bounds.height;
    }

    // 選択中のラジオからリサイズモードをフラグ付きで取得
    // 基準グループの並び: 0=最大 1=最小 2=指定サイズ 3=基準(長/短) 4=面積 5=アートボード 6=裁ち落とし
    function getSelectedResizeMode() {
        for (var groupIndex = 0; groupIndex < radioGroups.length; groupIndex++) {
            var sideIndex = -1;
            if (radioGroups[groupIndex][0].value) sideIndex = 0;
            else if (radioGroups[groupIndex][1].value) sideIndex = 1;
            if (sideIndex < 0) continue;
            return {
                group: groupIndex,
                side: sideIndex,
                isWidth: sideIndex === 0,
                isMax: groupIndex === 0,
                isMin: groupIndex === 1,
                isFixed: groupIndex === 2,
                isLong: groupIndex === 3 && sideIndex === 0,
                isShort: groupIndex === 3 && sideIndex === 1,
                isArea: groupIndex === 4,
                isAreaMax: groupIndex === 4 && sideIndex === 0,
                isAreaMin: groupIndex === 4 && sideIndex === 1,
                isArtboard: groupIndex === 5,
                isBleed: groupIndex === 6
            };
        }
        return null;
    }

    // 選択全体（クラスタ）の合成バウンズ（getReferenceBounds ベース）
    function getCombinedReferenceBounds(items) {
        var left = null, top = null, right = null, bottom = null;
        for (var i = 0; i < items.length; i++) {
            var itemBounds = getReferenceBounds(items[i]);
            if (left === null || itemBounds.left < left) left = itemBounds.left;
            if (top === null || itemBounds.top > top) top = itemBounds.top;
            if (right === null || (itemBounds.left + itemBounds.width) > right) right = itemBounds.left + itemBounds.width;
            if (bottom === null || (itemBounds.top - itemBounds.height) < bottom) bottom = itemBounds.top - itemBounds.height;
        }
        return { left: left, top: top, right: right, bottom: bottom, width: right - left, height: top - bottom };
    }

    // アートボード／裁ち落とし基準: 選択全体をひとまとまりとして等倍スケール（グループ化しない）。
    // クラスタ左上を原点に、各アイテムのサイズと相対位置を同じ倍率でスケールする。
    // → 親階層・重ね順を一切変更しないので、旧・一時グループ方式の復元リスクを構造的に回避する。
    function applyClusterResize(mode, referenceValue) {
        var clusterBounds = getCombinedReferenceBounds(workingItems);
        var current = mode.isWidth ? clusterBounds.width : clusterBounds.height;
        if (!current) return;
        var scaleFactor = referenceValue / current;

        // リサイズで位置がずれる前に、各アイテムの左上と原点（クラスタ左上）を記録
        var originLeft = null, originTop = null;
        var originalPositions = [];
        for (var i = 0; i < workingItems.length; i++) {
            var itemLeft = workingItems[i].left, itemTop = workingItems[i].top;
            originalPositions.push({ left: itemLeft, top: itemTop });
            if (originLeft === null || itemLeft < originLeft) originLeft = itemLeft;
            if (originTop === null || itemTop > originTop) originTop = itemTop;
        }

        var scalePct = scaleFactor * 100;
        for (var i = 0; i < workingItems.length; i++) {
            var item = workingItems[i];
            item.resize(scalePct, scalePct, true, true, true, true, scalePct, Transformation.TOPLEFT);
            // 原点からの相対位置も同倍率でスケール（クラスタとして拡大縮小）
            item.left = originLeft + (originalPositions[i].left - originLeft) * scaleFactor;
            item.top = originTop - (originTop - originalPositions[i].top) * scaleFactor;
        }
    }

    // 目標寸法（ポイント）を算出。面積モードは applyAreaResize で別処理
    function computeReferenceValue(mode) {
        if (mode.isFixed) {
            var parsed = parseFloat(createRadioGroup.sizeInput.text);
            if (isNaN(parsed) || parsed <= 0) return null;
            return parsed * pointsPerUnit();
        }
        if (mode.isArtboard) {
            var rect = getActiveArtboardRect();
            return mode.isWidth ? (rect[2] - rect[0]) : (rect[1] - rect[3]);
        }
        if (mode.isBleed) {
            var bleedBase = getActiveArtboardRect();
            var sideLength = mode.isWidth ? (bleedBase[2] - bleedBase[0]) : (bleedBase[1] - bleedBase[3]);
            return sideLength + BLEED_OFFSET_PT;
        }
        // 最大／最小／長辺／短辺: 全アイテムから基準値を集計
        var referenceValue = null;
        for (var i = 0; i < workingItems.length; i++) {
            var value = measureSide(getReferenceBounds(workingItems[i]), mode);
            if (mode.isMin && value === 0) continue;
            if (referenceValue === null) referenceValue = value;
            else if (mode.isMax && value > referenceValue) referenceValue = value;
            else if (mode.isMin && value < referenceValue) referenceValue = value;
        }
        return referenceValue;
    }

    // 片辺のみのスケール。引数を省略した 2 引数版は基準点が中心になり、他モード（TOPLEFT）と
    // 挙動が食い違うため、常に明示指定する。線幅は非等倍では変えない（100%）。
    // Always spell out the anchor: the 2-argument form scales about the center,
    // which would be inconsistent with every other mode. Line widths stay at 100%.
    function resizeOneSide(item, isWidth, scale) {
        var scaleX = isWidth ? scale : 100;
        var scaleY = isWidth ? 100 : scale;
        item.resize(scaleX, scaleY, true, true, true, true, 100, Transformation.TOPLEFT);
    }

    // 各アイテムを基準値に合わせてスケール（縦横比保持／片辺のみ）
    function resizeItemsToReference(mode, referenceValue) {
        var keepOneSideOnly = oneSideOnlyRadio.value;
        for (var i = 0; i < workingItems.length; i++) {
            var bounds = getReferenceBounds(workingItems[i]);
            var current = measureSide(bounds, mode);
            if (current === 0) continue;
            var scale = getScaleFactor(current, referenceValue);

            if (!keepOneSideOnly) {
                workingItems[i].resize(scale, scale, true, true, true, true, scale, Transformation.TOPLEFT);
            } else if (mode.isFixed) {
                var currentSide = mode.isWidth ? bounds.width : bounds.height;
                if (currentSide === 0) continue;
                var scaleFixed = getScaleFactor(currentSide, referenceValue);
                resizeOneSide(workingItems[i], mode.isWidth, scaleFixed);
            } else {
                resizeOneSide(workingItems[i], mode.isWidth, scale);
            }
        }
    }

    // アートボード中心へ選択全体を移動
    // （裁ち落としのオフセットは上下左右対称なので、中心はアートボードと一致する）
    function centerItemsOnArtboard() {
        var rect = getActiveArtboardRect();
        var centerX = (rect[0] + rect[2]) / 2;
        var centerY = (rect[1] + rect[3]) / 2;
        var clusterBounds = getCombinedReferenceBounds(workingItems);
        var dx = centerX - (clusterBounds.left + clusterBounds.width / 2);
        var dy = centerY - (clusterBounds.top - clusterBounds.height / 2);
        for (var i = 0; i < workingItems.length; i++) {
            workingItems[i].left += dx;
            workingItems[i].top += dy;
        }
    }

    // 面積基準（最大／最小の面積に全アイテムを合わせる）
    function applyAreaResize(mode) {
        var areas = [];
        for (var i = 0; i < workingItems.length; i++) {
            var itemBounds = getReferenceBounds(workingItems[i]);
            areas.push(itemBounds.width * itemBounds.height);
        }
        var baseArea = mode.isAreaMax ? Math.max.apply(null, areas) : Math.min.apply(null, areas);
        if (!baseArea || baseArea <= 0) return;
        for (var i = 0; i < workingItems.length; i++) {
            resizeToMatchArea(workingItems[i], baseArea);
        }
        captureResizeBaseState();
        app.redraw();
    }

    // --- リサイズ適用（オーケストレーション） / Apply resize by current selection ---
    function applyResizeBySelection() {
        var mode = getSelectedResizeMode();
        if (!mode) return;
        clearOutlineBoundsCache();

        if (mode.isArea) {
            applyAreaResize(mode);
            return;
        }

        var referenceValue = computeReferenceValue(mode);
        if (referenceValue === null || referenceValue <= 0) return;

        if (mode.isArtboard || mode.isBleed) {
            // 選択全体をアートボード（＋裁ち落とし）に合わせて等倍スケールし、中心へ配置
            applyClusterResize(mode, referenceValue);
            centerItemsOnArtboard();
        } else {
            // 最大／最小／指定サイズ／長辺／短辺: 各アイテムを個別にスケール
            resizeItemsToReference(mode, referenceValue);
        }
        captureResizeBaseState();
        app.redraw();
    }

    function resizeToMatchArea(item, targetArea) {
        var bounds = getReferenceBounds(item);
        var area = bounds.width * bounds.height;
        if (area === 0) return;
        var scale = Math.sqrt(targetArea / area) * 100;
        item.resize(scale, scale, true, true, true, true, scale, Transformation.TOPLEFT);
    }

    // item を目標サイズへリサイズ。除数 0（線のみ・空テキスト等）は Infinity 回避のため 100%＝現状維持とする。
    // ※ 幅/高さが 0 の要素は resize では拡大できないため「厳密復元」ではなく現状スキップに近い（実用上は許容）
    function resizeItemToSize(item, targetWidth, targetHeight) {
        var scaleW = (targetWidth === 0 || item.width === 0) ? 100 : (targetWidth / item.width) * 100;
        var scaleH = (targetHeight === 0 || item.height === 0) ? 100 : (targetHeight / item.height) * 100;
        item.resize(scaleW, scaleH, true, true, true, true, scaleW, Transformation.TOPLEFT);
    }

    function restoreOriginalGeometry() {
        for (var i = 0; i < originalStates.length; i++) {
            resizeItemToSize(originalStates[i].item, originalStates[i].width, originalStates[i].height);
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

    function getReferenceBounds(item, forceVisible) {
        var useVisibleBounds = !!(forceVisible || (previewCheck && previewCheck.value));
        if (textOutlineBoundsCheck && textOutlineBoundsCheck.value && containsTextForOutlineBounds(item)) {
            return getOutlinedBoundsCached(item, useVisibleBounds);
        }
        return getPageItemBoundsObject(item, useVisibleBounds);
    }

    // 整列も getReferenceBounds と同じ基準で計算する。
    // 「プレビュー境界で計測」ON のときは visibleBounds（線幅・効果込み＝見た目の端）、
    // OFF のときは geometricBounds（パスの端）。リサイズと基準を揃えることで、
    // 線幅のあるオブジェクトでも「見た目の端」で整列がそろう。
    // Alignment uses the same basis as getReferenceBounds: visibleBounds when
    // "measure by preview bounds" is on (visual edges incl. strokes/effects),
    // geometricBounds otherwise — so resize and alignment stay consistent.
    function getAlignmentBounds(item) {
        var useVisibleBounds = !!(previewCheck && previewCheck.value);
        if (textOutlineBoundsCheck && textOutlineBoundsCheck.value && containsTextForOutlineBounds(item)) {
            return getOutlinedBoundsCached(item, useVisibleBounds);
        }
        return getPageItemBoundsObject(item, useVisibleBounds);
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
                var previousSelection = doc.selection;
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
            var state = resizeBaseStates[i];
            resizeItemToSize(state.item, state.width, state.height);
            state.item.left = state.left;
            state.item.top = state.top;
        }
        app.redraw();
    }
})();