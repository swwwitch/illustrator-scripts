#target illustrator
#targetengine "StrokeGraphEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.StrokeGraph
Meta-Default-language=ja
Name=線を2色の破線グラフにする
Name[en]=Turn strokes into two-color dashed graphs
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.1
Release-Date=2026-11-20
Tags=color,カラー
X-com.swwwitch-Script-Name=StrokeGraph
X-com.swwwitch-Updated=2026-11-22
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/StrokeGraph.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/StrokeGraph.md
X-com.swwwitch-Article-URL=https://note.com/dtp_tranist/n/ne78abaa5610e
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/StrokeGraph.jsx
Description-BEGIN
[ja]
選択したパスの線を、ベースとアクセントの2色で塗り分けた丸型の破線にして、円グラフや進捗バーのような線のグラフを作ります。
パスの色をアクセントにし、グラデーションの無い線にも、埋め込んだグラフィックスタイルで［パスに沿ってグラデーションを適用］の線を付けます。
[en]
Turns the strokes of the selected paths into round dashes split into a base color and an accent color, making stroke graphs such as ring charts and progress bars.
The path's own color becomes the accent, and strokes without a gradient get Apply Gradient Along Stroke from an embedded graphic style.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 楕円のライブシェイプに入れる角度（度。変形パネルの［楕円形のプロパティ］の角度） / Angle set on live ellipses (degrees) */
    var ELLIPSE_ANGLE = 90;

    /* ［キャンセル］で入れ直す楕円の角度（元の角度は読めないため、作成時の 0° に戻す） / Angle restored on Cancel (the original cannot be read) */
    var ELLIPSE_ANGLE_ON_CANCEL = 0;

    /* 楕円の角度がこの差（度）以内なら、すでに ELLIPSE_ANGLE になっているとみなす / Angle difference treated as already at ELLIPSE_ANGLE */
    var ELLIPSE_ANGLE_TOLERANCE = 0.01;

    /* アクセントの割合の初期値（線がまだ2色に分かれていないとき） / Default accent ratio when the stroke is not split yet */
    var DEFAULT_ACCENT_PERCENT = 20;

    /* 線幅の初期値（pt。線がまだグラデーションでないとき） / Default weight in pt when the stroke has no gradient yet */
    var DEFAULT_STROKE_WIDTH = 10;

    /* アクセントの割合の欄の増減量・範囲・単位 / Step, range and unit of the accent ratio field */
    var ACCENT_FIELD_OPTIONS = { step: 1, min: 0, max: 100, unit: "%" };

    /* 線幅の欄の増減量・範囲（線の単位） / Step and range of the weight field (stroke units) */
    var WEIGHT_FIELD_OPTIONS = { step: 1, min: 0, max: 1000 };

    /* 入れる破線（pt）。[線分] だけなら間隔は空白（線分と同じ間隔） / Dash pattern (pt); a single value leaves the gap blank */
    var STROKE_DASH_PATTERN = [1];

    /* 色の切り替わりの幅がこれ以内（%）なら2色に分かれているとみなし、今の境界からアクセントの割合を出す / Max transition width (%) treated as already split */
    var SEPARATE_EDGE_MAX_GAP = 1;

    /* 線にグラデーションが無いパスで使うベース・アクセントの初期値（ドキュメントのカラーモードで使い分け）
       Default base and accent colors for paths without a gradient stroke, per document color mode */
    var DEFAULT_BASE_COLOR   = { cmyk: [0, 0, 0, 10], rgb: [230, 230, 230] };
    var DEFAULT_ACCENT_COLOR = { cmyk: [0, 90, 85, 0], rgb: [230, 62, 50] };

    /* パスごとに作るグラデーションの名前 / Name of the gradient made for each path */
    var GRAPH_GRADIENT_BASE_NAME = "Stroke Graph";

    /* 線にグラデーションが無いパスに先に適用するグラフィックスタイル（［パスに沿ってグラデーションを適用］の線）。
       スクリプトの末尾に埋め込んだ AI ファイルから取り込む。埋め込むファイルは helpers/embed-file-base64.py で差し替える
       Graphic style applied first to paths without a gradient stroke (a gradient applied along the stroke).
       It is imported from the AI file embedded at the end of this script; replace that file with helpers/embed-file-base64.py */
    var STROKE_GRADIENT_STYLE_NAME = "stroke-gradient";

    /* グラフィックスタイルを取り込むときに一時的に貼り付けるレイヤー / Temporary layer the style is pasted onto while importing */
    var IMPORT_LAYER_NAME = "// _imported";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var FIELD_LABEL_WIDTH = 115;     /* 項目名の幅（「アクセントの割合 :」が収まる幅） / label width */
    var FIELD_CHARACTERS  = 6;       /* 入力欄の文字数 / characters of the fields */
    var COLOR_CHIP_SIZE   = 16;      /* カラーの見本の大きさ / color chip size */
    var COLOR_PANEL_SPACING = 10;    /* ［カラー］パネルの行の間隔 / row spacing of the Color panel */
    var COLOR_CHIP_TEXT_GAP = 8;     /* カラーの見本と値の間隔 / gap between the color chip and its value */
    var REVERSE_ROW_TOP_GAP = 4;     /* 線幅と［方向反転］の間に足すアキ / extra space between the weight and Reverse Direction */
    var COLOR_LABEL_CHIP_GAP = 8;    /* カラーの項目名（「ベース :」など）と見本の間隔 / gap between the color label and the chip */
    var COLOR_TEXT_WIDTH  = 150;     /* カラーの値の欄の幅（「C100 M100 Y100 K100」が収まる幅） / width of the color value text */

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
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
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 240; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // UI の明暗（再利用パーツ） / UI theme (reusable)

    /**
     * UI がダークテーマかどうかを判定する（Illustrator は uiBrightness、InDesign は uiBrightnessPreference）
     * @returns {boolean} ダークなら true。取得できない環境では false（明るいUI扱い）
     */
    function isDarkUI() {
        try {
            if (app.preferences && app.preferences.getRealPreference) {
                return app.preferences.getRealPreference("uiBrightness") <= 0.5; /* Illustrator */
            }
            return app.generalPreferences.uiBrightnessPreference <= 0.5; /* InDesign */
        } catch (e) {
            return false;
        }
    }

    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme

    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)

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

    /* 単位の換算は UnitValue に任せる（in / ft / yd / mm / cm / m / pt / pc / px ほか、単数形・複数形も可）。
       UnitValue に無い単位だけ、ここで UnitValue の単位に読み替える（値は「1単位＝何 unit か」）。
       「p」は「1p6」（1パイカ6ポイント）の形にも使う
       Units UnitValue lacks, mapped onto UnitValue units (how many of `unit` make one) */
    var STEPPER_UNIT_ALIASES = {
        "q": { unit: "mm", amount: 0.25 },    /* 級 / Q */
        "h": { unit: "mm", amount: 0.25 },    /* 歯 / H */
        "p": { unit: "pc", amount: 1 },       /* パイカ / pica */
        "ft/in": { unit: "ft", amount: 1 },   /* Illustrator の単位コード7の表示 / Illustrator unit code 7 */
        "c": { unit: "ci", amount: 1 },       /* シセロ（InDesign の表示） / ciceros as InDesign shows them */
        "ag": { unit: "in", amount: 1 / 14 }, /* アゲート / agates */
        "ap": { unit: "tpt", amount: 1 }      /* アメリカンポイント / American points */
    };

    // -----------------------------------------
    // ステップボタンの配色 / Stepper colors
    // -----------------------------------------
    var STEPPER_UI_DARK           = isDarkUI();
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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す。
     * 四則演算（+ - * / と括弧）を入れると、確定時に計算した値にする。欄と違う単位で入れた値は欄の単位へ換算する（mm の欄に「1 in」→「25.4 mm」）
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

        /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
        fieldLabel.addEventListener("click", function () { focusNumberInput(numberInput); });

        /* 直接入力をそろえる。計算式は計算し、数値でなければ直前の値に戻す / normalize typed values; evaluate arithmetic, revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = evaluateArithmetic(numberInput.text, fieldOptions.unit);
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
     * 数値欄の単位を差し替える（単位の設定やドロップダウンを切り替えたとき用）。
     * shouldConvert が true なら値を新しい単位へ換算し（10 mm → 28.35 pt）、false なら数値はそのままで単位だけ付け替える
     * @param {EditText} numberInput - addSteppedField() で作った入力欄、または bindSteppedArrowKeys() を呼んだ入力欄
     * @param {string} unit - 新しい単位（例 " pt"。単位なしは ""）
     * @param {boolean} [shouldConvert] - 値も換算するなら true
     * @returns {void}
     */
    function setSteppedFieldUnit(numberInput, unit, shouldConvert) {
        var stepOptions = numberInput.stepperGroup.stepOptions;
        var oldUnit = stepOptions.unit || "";
        var value = parseFloat(numberInput.text);
        stepOptions.unit = unit;
        if (isNaN(value)) return;
        if (shouldConvert) {
            var converted = evaluateArithmetic(String(value) + oldUnit, unit);
            if (!isNaN(converted)) value = converted;
        }
        numberInput.text = formatStepperNumber(value) + unit;
        numberInput.lastValidText = numberInput.text;
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
            var value = evaluateArithmetic(numberInput.text, stepOptions.unit); /* 確定前の計算式も計算してから増減 / evaluate an uncommitted expression first */
            if (isNaN(value)) value = parseFloat(numberInput.text); /* 計算できなければ従来どおり先頭の数値 / fall back to the leading number */
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /**
         * ∧∨を離したときに入力欄へフォーカスを移す（mousedown で移しても、離したときに外れる）
         * @param {Group} chevronButton - makeStepperChevronButton() で作ったボタン
         * @returns {Group} 渡したボタン
         */
        function focusInputOnRelease(chevronButton) {
            chevronButton.addEventListener("mouseup", function () {
                var numberInput = getNumberInput();
                if (isStepperEnabledInTree(numberInput)) focusNumberInput(numberInput);
            });
            return chevronButton;
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); })).helpTip = getLabel(upTooltip);
        focusInputOnRelease(makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); })).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        stepperGroup.stepOptions = stepOptions; /* 確定時の計算で欄の単位を引けるよう公開 / lets the commit-time evaluation find the unit */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し。
     * あわせて、確定時に計算式・単位付きの値を計算して書き戻す（各スクリプトの onChange より先に呼ばれるので、onChange は計算後の値を読む）
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
        numberInput.addEventListener("change", function () {
            var fieldUnit = stepperGroup.stepOptions ? stepperGroup.stepOptions.unit : undefined;
            var value = evaluateArithmetic(numberInput.text, fieldUnit);
            if (isNaN(value)) return; /* 計算できなければ各スクリプトの処理に任せる / leave it to the script's own handler */
            /* 式か、換算で値が変わったときだけ書き戻す（ただの数値は書式を崩さない） / rewrite only expressions and converted values */
            var hasOperator = /[*\/()\u00D7\u00F7\uFF0A\uFF0F\uFF08\uFF09]|[\d.\uFF10-\uFF19][^\d.\uFF10-\uFF19]*[+\-\u2212\uFF0B\uFF0D]/.test(numberInput.text);
            if (!hasOperator && value === parseFloat(numberInput.text)) {
                /* 単位を省いて入れた数値には、欄の単位だけ付け足す（桁は丸めない） / append the field unit to a bare number */
                var trimmedText = numberInput.text.replace(/^\s+|\s+$/g, "");
                if (fieldUnit && /[\d.]$/.test(trimmedText)) numberInput.text = trimmedText + fieldUnit;
                return;
            }
            numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
        });
        numberInput.stepperGroup = stepperGroup; /* setSteppedFieldUnit() から∧∨の設定を引けるようにする / lets setSteppedFieldUnit() find the options */
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
        /* 0.29 / 0.01 = 28.999… のような浮動小数の誤差で同じ値に戻らないよう、商を丸めてから切り捨て・切り上げる
           round the quotient first so float error (0.29 / 0.01 = 28.999…) does not step back to the same value */
        var quotient = Math.round(value / multiple * 1e6) / 1e6;
        if (direction > 0) return Math.round((Math.floor(quotient) + 1) * multiple * 1e6) / 1e6;
        return Math.round((Math.ceil(quotient) - 1) * multiple * 1e6) / 1e6;
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
     * 入力欄の文字列を四則演算（+ - * / と括弧）として計算する。eval は使わない。
     * 数値の後ろの単位は UnitValue で欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
     * 全角の数字・記号と × ÷ は半角に直す
     * @param {string} text - 入力欄の文字列
     * @param {string} [fieldUnit] - 欄の単位（例 " mm"。前後の空白は無視）
     * @returns {number} 欄の単位での計算結果（式として読めない・換算できない単位・0で割ったときは NaN）
     */
    function evaluateArithmetic(text, fieldUnit) {
        var source = String(text)
            .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
            .replace(/×/g, "*")
            .replace(/÷/g, "/")
            .replace(/[−–—]/g, "-")
            .replace(/\s/g, "");
        if (source === "") return NaN;
        var fieldUnitKey = String(fieldUnit || "").replace(/^\s+|\s+$/g, "").toLowerCase();
        var fieldUnitValue = createStepperUnitValue(1, fieldUnitKey); /* 欄の単位の1単位（換算できない欄は null） / one field unit */
        var position = 0;

        /**
         * 加減算の並び（項 ± 項 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readSum() {
            var total = readProduct();
            while (position < source.length && (source.charAt(position) === "+" || source.charAt(position) === "-")) {
                var operator = source.charAt(position++);
                var operand = readProduct();
                total = (operator === "+") ? total + operand : total - operand;
            }
            return total;
        }

        /**
         * 乗除算の並び（因子 × 因子 …）を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readProduct() {
            var total = readFactor();
            while (position < source.length && (source.charAt(position) === "*" || source.charAt(position) === "/")) {
                var operator = source.charAt(position++);
                var operand = readFactor();
                if (operator === "/" && operand === 0) return NaN;
                total = (operator === "*") ? total * operand : total / operand;
            }
            return total;
        }

        /**
         * 符号付きの数値（単位付きなら欄の単位へ換算）か、括弧で囲んだ式を読む
         * @returns {number} 値（読めなければ NaN）
         */
        function readFactor() {
            var ch = source.charAt(position);
            if (ch === "+" || ch === "-") {
                position++;
                var signedValue = readFactor();
                return (ch === "-") ? -signedValue : signedValue;
            }
            if (ch === "(") {
                position++;
                var innerValue = readSum();
                if (source.charAt(position) !== ")") return NaN;
                position++;
                return innerValue;
            }
            var numberMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
            if (!numberMatch) return NaN;
            position += numberMatch[0].length;
            return readUnitSuffix(parseFloat(numberMatch[0]));
        }

        /**
         * 数値の直後の単位を読み、欄の単位へ換算する
         * @param {number} value - 単位の前の数値
         * @returns {number} 欄の単位での値（換算できない単位なら NaN）
         */
        function readUnitSuffix(value) {
            var unitMatch = /^(ft\/in|[A-Za-z]+|%|°)/i.exec(source.substring(position));
            if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
            position += unitMatch[0].length;
            var unitKey = unitMatch[0].toLowerCase();
            if (unitKey === fieldUnitKey) return value;
            var typedValue = createStepperUnitValue(value, unitKey);
            if (!typedValue || !fieldUnitValue) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
            var points = typedValue.as("pt");
            /* 「1p6」＝1パイカ6ポイント / pica-point notation */
            if (unitKey === "p") {
                var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                if (pointMatch) {
                    position += pointMatch[0].length;
                    points += parseFloat(pointMatch[0]);
                }
            }
            return points / fieldUnitValue.as("pt");
        }

        var result = readSum();
        if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
        return result;
    }

    /**
     * 数値と単位から UnitValue を作る。Q・H・p は STEPPER_UNIT_ALIASES で UnitValue の単位に読み替える。
     * %（percent）は基準の長さが無いと換算できないので扱わない
     * @param {number} value - 数値
     * @param {string} unitKey - 単位（小文字。例 "mm"、"inches"、"q"）
     * @returns {UnitValue|null} UnitValue（UnitValue が知らない単位・空・% なら null）
     */
    function createStepperUnitValue(value, unitKey) {
        if (unitKey === "" || unitKey === "%") return null;
        var alias = STEPPER_UNIT_ALIASES[unitKey];
        var unitValue = alias ? new UnitValue(value * alias.amount, alias.unit) : new UnitValue(value, unitKey);
        if (unitValue.type === "?" || unitValue.type === "%") return null; /* 知らない単位は例外にならず "?" になる。"percent" も除く / unknown units become "?" */
        return unitValue;
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
     * 入力欄にフォーカスを移す
     * @param {EditText} numberInput - 対象の入力欄
     * @returns {void}
     */
    function focusNumberInput(numberInput) {
        numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
        numberInput.active = true;
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

    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper

    // バージョン表示（再利用パーツ） / Version footer (reusable)

    var VERSION_FOOTER_TOP_MARGIN = 4;    /* バージョンの行の上に足す余白（px。ウィンドウの間隔に足す）/ extra space above the version row, added to the window spacing */
    var VERSION_FOOTER_BOTTOM_MARGIN = 8; /* バージョンの行の下の余白（px。ウィンドウの下余白をこの値にする）/ space below the version row (replaces the window's bottom margin) */
    var VERSION_FOOTER_PIXEL_SIZE = 1;    /* ドット1つの大きさ（px）/ size of one dot (px) */
    var VERSION_FOOTER_LETTER_GAP = 1;    /* 字の間（ドット）/ gap between glyphs (dots) */
    var VERSION_FOOTER_BRAND = "DTP Transit"; /* 右下に出す名前（字形にある文字だけ）/ name shown at the bottom right (glyph characters only) */
    var VERSION_FOOTER_COLOR = isDarkUI() ? [0.45, 0.45, 0.45, 1] : [0.62, 0.62, 0.62, 1]; /* 文字の色（ディム表示に近い灰色）/ text color, a gray close to the dimmed look */

    /* 字形：Sigi 7px Regular（Rasmus Andersson、CC0。https://github.com/rsms/sigi-pixel-font）の大文字・数字・「. - _ +」・空白を
       9 行のドット（# が点。上7行が大文字の高さ、下2行は小文字の下に出る部分）に写したもの。Sigi に小文字が無いので、a〜z は同じ大きさで描き足した。I は l と見分けるよう上下に横棒を足した
       / Glyphs: uppercase, digits, ". - _ +" and space from Sigi 7px Regular by Rasmus Andersson (CC0, https://github.com/rsms/sigi-pixel-font),
         nine rows each ("#" is a dot; seven rows of cap height plus two for descenders). Sigi has no lowercase, so a-z were drawn to match. I has top and bottom bars to tell it from l */
    var VERSION_FOOTER_GLYPH_HEIGHT = 9;  /* 字の高さ（ドット。下2行は g・j・p・q・y の下に出る部分）/ glyph height (dots), including two descender rows */
    var VERSION_FOOTER_GLYPHS = {
        "A": [".###.", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "B": ["####.", "#...#", "#...#", "####.", "#...#", "#...#", "####.", ".....", "....."],
        "C": [".###.", "#...#", "#....", "#....", "#....", "#...#", ".###.", ".....", "....."],
        "D": ["####.", "#...#", "#...#", "#...#", "#...#", "#...#", "####.", ".....", "....."],
        "E": ["#####", "#....", "#....", "####.", "#....", "#....", "#####", ".....", "....."],
        "F": ["#####", "#....", "#....", "####.", "#....", "#....", "#....", ".....", "....."],
        "G": [".###.", "#...#", "#....", "#..##", "#...#", "#...#", ".###.", ".....", "....."],
        "H": ["#...#", "#...#", "#...#", "#####", "#...#", "#...#", "#...#", ".....", "....."],
        "I": ["###", ".#.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "J": ["....#", "....#", "....#", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "K": ["#...#", "#..#.", "#.#..", "##...", "#.#..", "#..#.", "#...#", ".....", "....."],
        "L": ["#...", "#...", "#...", "#...", "#...", "#...", "####", "....", "...."],
        "M": ["#...#", "##.##", "#.#.#", "#...#", "#...#", "#...#", "#...#", ".....", "....."],
        "N": ["#...#", "#...#", "##..#", "#.#.#", "#..##", "#...#", "#...#", ".....", "....."],
        "O": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "P": ["####.", "#...#", "#...#", "####.", "#....", "#....", "#....", ".....", "....."],
        "Q": [".###.", "#...#", "#...#", "#...#", "#...#", "#..##", ".####", ".....", "....."],
        "R": ["####.", "#...#", "#...#", "####.", "#.#..", "#..#.", "#...#", ".....", "....."],
        "S": [".###.", "#...#", "#....", ".###.", "....#", "#...#", ".###.", ".....", "....."],
        "T": ["#####", "..#..", "..#..", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "U": ["#...#", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "V": ["#...#", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", "..#..", ".....", "....."],
        "W": ["#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", "#..#..#", ".##.##.", ".......", "......."],
        "X": ["#...#", "#...#", ".#.#.", "..#..", ".#.#.", "#...#", "#...#", ".....", "....."],
        "Y": ["#...#", "#...#", ".#.#.", "..#..", "..#..", "..#..", "..#..", ".....", "....."],
        "Z": ["#####", "....#", "...#.", "..#..", ".#...", "#....", "#####", ".....", "....."],
        "a": ["....", "....", ".##.", "...#", ".###", "#..#", ".###", "....", "...."],
        "b": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "###.", "....", "...."],
        "c": ["....", "....", ".###", "#...", "#...", "#...", ".###", "....", "...."],
        "d": ["...#", "...#", ".###", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "e": ["....", "....", ".##.", "#..#", "####", "#...", ".###", "....", "...."],
        "f": ["..#", ".#.", "###", ".#.", ".#.", ".#.", ".#.", "...", "..."],
        "g": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "h": ["#...", "#...", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "i": ["#", ".", "#", "#", "#", "#", "#", ".", "."],
        "j": [".#", "..", ".#", ".#", ".#", ".#", ".#", ".#", "#."],
        "k": ["#...", "#...", "#..#", "#.#.", "##..", "#.#.", "#..#", "....", "...."],
        "l": ["#.", "#.", "#.", "#.", "#.", "#.", ".#", "..", ".."],
        "m": [".....", ".....", "####.", "#.#.#", "#.#.#", "#.#.#", "#.#.#", ".....", "....."],
        "n": ["....", "....", "###.", "#..#", "#..#", "#..#", "#..#", "....", "...."],
        "o": ["....", "....", ".##.", "#..#", "#..#", "#..#", ".##.", "....", "...."],
        "p": ["....", "....", "###.", "#..#", "#..#", "#..#", "###.", "#...", "#..."],
        "q": ["....", "....", ".###", "#..#", "#..#", "#..#", ".###", "...#", "...#"],
        "r": ["...", "...", "#.#", "##.", "#..", "#..", "#..", "...", "..."],
        "s": ["....", "....", ".###", "#...", ".##.", "...#", "###.", "....", "...."],
        "t": ["...", ".#.", "###", ".#.", ".#.", ".#.", ".##", "...", "..."],
        "u": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "....", "...."],
        "v": [".....", ".....", "#...#", "#...#", ".#.#.", ".#.#.", "..#..", ".....", "....."],
        "w": [".....", ".....", "#...#", "#...#", "#.#.#", "#.#.#", ".#.#.", ".....", "....."],
        "x": ["....", "....", "#..#", "#..#", ".##.", "#..#", "#..#", "....", "...."],
        "y": ["....", "....", "#..#", "#..#", "#..#", "#..#", ".###", "...#", "###."],
        "z": ["....", "....", "####", "...#", ".##.", "#...", "####", "....", "...."],
        "0": [".###.", "#...#", "#...#", "#...#", "#...#", "#...#", ".###.", ".....", "....."],
        "1": [".#.", "##.", ".#.", ".#.", ".#.", ".#.", "###", "...", "..."],
        "2": [".###.", "#...#", "....#", "...#.", "..#..", ".#...", "#####", ".....", "....."],
        "3": [".###.", "#...#", "....#", "..##.", "....#", "#...#", ".###.", ".....", "....."],
        "4": ["...#.", "..##.", ".#.#.", "#..#.", "#####", "...#.", "...#.", ".....", "....."],
        "5": ["#####", "#....", "####.", "....#", "....#", "#...#", ".###.", ".....", "....."],
        "6": [".###.", "#...#", "#....", "####.", "#...#", "#...#", ".###.", ".....", "....."],
        "7": ["#####", "....#", "....#", "...#.", "...#.", "..#..", "..#..", ".....", "....."],
        "8": [".###.", "#...#", "#...#", ".###.", "#...#", "#...#", ".###.", ".....", "....."],
        "9": [".###.", "#...#", "#...#", ".####", "....#", "#...#", ".###.", ".....", "....."],
        ".": [".", ".", ".", ".", ".", ".", "#", ".", "."],
        "-": ["...", "...", "...", "###", "...", "...", "...", "...", "..."],
        "_": ["....", "....", "....", "....", "....", "....", "####", "....", "...."],
        "+": [".....", "..#..", "..#..", "#####", "..#..", "..#..", ".....", ".....", "....."],
        " ": ["...", "...", "...", "...", "...", "...", "...", "...", "..."]
    };

    /**
     * ウィンドウの下端の左にスクリプト名とバージョン（SCRIPT_NAME・SCRIPTMETA の Version）、右に VERSION_FOOTER_BRAND を小さく薄く入れる。部品をすべて足したあと、show() の直前に呼ぶ。
     * 直前の行（ボタン行など）の下余白を 0 にし、ウィンドウの下余白を VERSION_FOOTER_BOTTOM_MARGIN にする。
     * InDesign の ScriptUI は文字サイズの指定を無視するので、onDraw でドットを並べて描く（UITheme の isDarkUI() を使う）
     * @param {Window} targetWindow - ダイアログまたはパレット（縦並び）
     * @returns {Group|null} バージョンを描くグループ。すでに足してあれば null
     */
    function addVersionFooter(targetWindow) {
        if (targetWindow.versionFooter) return null;

        /* ボタン行の下余白（BUTTON_ROW_BOTTOM_MARGIN）は、この行が引き受ける / this row takes over the button row's bottom margin */
        var lastChild = targetWindow.children[targetWindow.children.length - 1];
        if (lastChild && lastChild.type === "group") {
            var rowMargins = lastChild.margins;
            lastChild.margins = [rowMargins[0], rowMargins[1], rowMargins[2], 0];
        }
        var windowMargins = targetWindow.margins;
        targetWindow.margins = [windowMargins[0], windowMargins[1], windowMargins[2], VERSION_FOOTER_BOTTOM_MARGIN];

        var footerGroup = targetWindow.add("group");
        footerGroup.orientation = "row";
        footerGroup.alignment = ["fill", "bottom"];
        footerGroup.alignChildren = ["left", "center"];
        footerGroup.margins = [0, VERSION_FOOTER_TOP_MARGIN, 0, 0];

        /* 左は「スクリプト名 バージョン」。字形に無い文字（日本語など）を含む名前や、幅が足りないときはバージョンだけ。
           左の欄は余りの幅まで伸ばし、ダイアログの幅はバージョンの分しか求めない（名前でダイアログを広げない）
           / Left: "script name version", or just the version when the name has characters without glyphs (e.g. Japanese)
             or does not fit. The left area takes the extra width and only asks for the version's width, so the name never widens the dialog */
        var versionText = getVersionInfoValue("version");
        var leftCandidates = [versionText];
        if (typeof SCRIPT_NAME === "string" && hasVersionFooterGlyphs(SCRIPT_NAME)) leftCandidates.unshift(SCRIPT_NAME + " " + versionText);
        var versionCanvas = addVersionFooterCanvas(footerGroup, leftCandidates);
        versionCanvas.alignment = ["fill", "center"];
        bindVersionInfoClick(versionCanvas); /* クリックで「このスクリプトについて」を開く / click opens About This Script */
        var brandCanvas = addVersionFooterCanvas(footerGroup, [VERSION_FOOTER_BRAND]);
        brandCanvas.alignment = ["right", "center"];
        brandCanvas.maximumSize = brandCanvas.preferredSize;
        targetWindow.versionFooter = footerGroup;
        return versionCanvas;
    }

    /**
     * 文字列のすべての文字に字形があるかを返す
     * @param {string} dotText - 文字列
     * @returns {boolean} すべてあれば true
     */
    function hasVersionFooterGlyphs(dotText) {
        for (var i = 0; i < dotText.length; i++) {
            if (!VERSION_FOOTER_GLYPHS.hasOwnProperty(dotText.charAt(i))) return false;
        }
        return dotText.length > 0;
    }

    /**
     * ドットで文字列を描くグループを足す。候補のうち、欄の幅に収まる最初のものを描く
     * @param {Group} parentGroup - 足す先
     * @param {string[]} dotTexts - 描く文字列の候補（長い順。いちばん短いものの幅を欄の幅として求める）
     * @returns {Group} 描くグループ
     */
    function addVersionFooterCanvas(parentGroup, dotTexts) {
        var dotCanvas = parentGroup.add("group");
        var shortestText = dotTexts[dotTexts.length - 1];
        dotCanvas.preferredSize = [measureVersionFooterDots(shortestText) * VERSION_FOOTER_PIXEL_SIZE, VERSION_FOOTER_GLYPH_HEIGHT * VERSION_FOOTER_PIXEL_SIZE];
        dotCanvas.helpTip = dotTexts[0];
        dotCanvas.onDraw = function () {
            var availableWidth = this.size ? this.size.width : 0;
            var drawText = shortestText;
            for (var i = 0; i < dotTexts.length; i++) {
                if (measureVersionFooterDots(dotTexts[i]) * VERSION_FOOTER_PIXEL_SIZE <= availableWidth) {
                    drawText = dotTexts[i];
                    break;
                }
            }
            drawVersionFooterDots(this.graphics, drawText);
        };
        return dotCanvas;
    }

    /**
     * 文字列の幅をドットの数で返す（字形に無い文字は数えない）
     * @param {string} versionText - 文字列
     * @returns {number} 幅（ドット）
     */
    function measureVersionFooterDots(versionText) {
        var dotCount = 0;
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (dotCount > 0) dotCount += VERSION_FOOTER_LETTER_GAP;
            dotCount += glyphRows[0].length;
        }
        return dotCount;
    }

    /**
     * 文字列をドットで描く。横に続くドットは1つの四角にまとめて塗る
     * @param {ScriptUIGraphics} canvasGraphics - 描く先の graphics
     * @param {string} versionText - 文字列
     * @returns {void}
     */
    function drawVersionFooterDots(canvasGraphics, versionText) {
        var footerBrush = canvasGraphics.newBrush(canvasGraphics.BrushType.SOLID_COLOR, VERSION_FOOTER_COLOR);
        var dotSize = VERSION_FOOTER_PIXEL_SIZE;
        var cursorDots = 0;
        canvasGraphics.newPath();
        for (var i = 0; i < versionText.length; i++) {
            var glyphRows = VERSION_FOOTER_GLYPHS[versionText.charAt(i)];
            if (!glyphRows) continue;
            if (cursorDots > 0) cursorDots += VERSION_FOOTER_LETTER_GAP;
            for (var row = 0; row < glyphRows.length; row++) {
                var rowText = glyphRows[row];
                var runStart = -1;
                for (var col = 0; col <= rowText.length; col++) {
                    var isDot = col < rowText.length && rowText.charAt(col) === "#";
                    if (isDot && runStart < 0) runStart = col;
                    if (!isDot && runStart >= 0) {
                        canvasGraphics.rectPath((cursorDots + runStart) * dotSize, row * dotSize, (col - runStart) * dotSize, dotSize);
                        runStart = -1;
                    }
                }
            }
            cursorDots += glyphRows[0].length;
        }
        canvasGraphics.fillPath(footerBrush);
    }

    // -----------------------------------------
    // このスクリプトについて / About this script
    // -----------------------------------------
    /* 左下のスクリプト名・バージョンをクリックすると開く。SCRIPTMETA の項目を getScriptMetaValue() で直接読み（スクリプト情報の部品）、
       この部品だけで動くよう、文言（VERSION_INFO_LABELS）とレイアウトの値も部品の中に持つ
       Opened by clicking the name and version at the bottom left. Reads the SCRIPTMETA fields directly with getScriptMetaValue(),
       and keeps its own labels and layout values so the part works on its own */

    var VERSION_INFO_LANG = (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en"; /* 文言の言語 / label language */
    var VERSION_INFO_IS_INDESIGN = (String(app.name).indexOf("InDesign") >= 0);         /* InDesign で動いているか / running in InDesign */
    /* 実行中のスクリプトのファイル。ボタンを押すころには $.fileName が別のものを指すことがあるので、読み込んだ時点で控える
       The running script file, captured at load time; $.fileName may point elsewhere by the time the button is clicked */
    var VERSION_INFO_SCRIPT_FILE = new File($.fileName);
    var VERSION_INFO_REVEAL_APP_PATH = "/Applications/RevealInFinder.app";   /* Finder で選択表示する補助アプリ / helper app that reveals a file in the Finder */
    var VERSION_INFO_REVEAL_PATH_FILE = "/tmp/illustrator_reveal_path.txt";  /* 補助アプリが読むパス（アプリの旧名のまま。変えると読めない）/ path the helper reads (keeps the app's old name) */
    var VERSION_INFO_FILE_VIEWER_APP_PATH = "/Applications/OpenInFileViewer.app";   /* Path Finder が起動中なら Path Finder、そうでなければ Finder で選択表示する補助アプリ / helper app that reveals a file in Path Finder when it is running, otherwise in the Finder */
    var VERSION_INFO_FILE_VIEWER_PATH_FILE = "/tmp/open_in_file_viewer_path.txt"; /* OpenInFileViewer.app が読むパス / path the helper reads */

    var VERSION_INFO_WINDOW_MARGINS = 16;               /* ウィンドウの余白 / window margin */
    var VERSION_INFO_WINDOW_SPACING = 12;               /* ウィンドウ内の間隔 / window spacing */
    var VERSION_INFO_PANEL_MARGINS = [16, 20, 16, 12];  /* パネルの余白 [左,上,右,下] / panel margins */
    var VERSION_INFO_ROW_SPACING = 6;                   /* 項目名と値の間隔 / gap between a label and its value */
    var VERSION_INFO_BUTTON_SPACING = 10;               /* ボタンどうしの間隔 / gap between buttons */
    var VERSION_INFO_BUTTON_ROW_MARGINS = [0, 5, 0, 14]; /* ボタン行の余白（Illustrator 標準のダイアログに合わせる）/ button row margins */
    var VERSION_INFO_COPY_BUTTON_SIZE = 22;             /* コピーボタンの一辺 / copy button size */
    /* コピーボタンの配色。∧∨（ステップボタン）と同じく、背景に重ねる半透明 / Copy button colors, translucent overlays like the stepper */
    var VERSION_INFO_COPY_FILL_COLOR    = isDarkUI() ? [0, 0, 0, 0.10] : [1, 1, 1, 0.50]; /* 地 / background */
    var VERSION_INFO_COPY_FRAME_COLOR   = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.15]; /* 枠線 / frame */
    var VERSION_INFO_COPY_PRESSED_COLOR = isDarkUI() ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 押下中 / pressed */
    var VERSION_INFO_COPY_ICON_COLOR    = isDarkUI() ? [1, 1, 1, 0.85] : [0, 0, 0, 0.70]; /* アイコン / icon */

    var VERSION_INFO_LABELS = {
        dialog: {
            title: { ja: "このスクリプトについて", en: "About This Script" }
        },
        panel: {
            generalInfo: { ja: "基本情報", en: "General" },
            links: { ja: "リンク", en: "Links" }
        },
        fieldLabel: {
            scriptName: { ja: "スクリプト名", en: "Script Name" },
            scriptFeature: { ja: "機能", en: "Feature" },
            scriptVersion: { ja: "バージョン", en: "Version" }
        },
        button: {
            openReadme: { ja: "README", en: "README" },
            openSourceCode: { ja: "ソースコード", en: "Source Code" },
            openArticle: { ja: "記事（note）", en: "Article (note)" },
            revealInFinder: { ja: "Finder に表示", en: "Show in Finder" },
            revealInExplorer: { ja: "エクスプローラーに表示", en: "Show in Explorer" },
            close: { ja: "閉じる", en: "Close" }
        },
        tooltip: {
            versionInfo: { ja: "クリックで、このスクリプトについての情報を表示", en: "Click for info about this script" },
            copyScriptName: { ja: "スクリプト名をコピー", en: "Copy Script Name" }
        },
        alert: {
            copyFailed: { ja: "スクリプト名をコピーできませんでした。", en: "Could not copy the script name." }
        }
    };

    /**
     * VERSION_INFO_LABELS の文言を今の言語で返す
     * @param {Object} labelEntry - { ja, en }
     * @returns {string} 文言
     */
    function getVersionInfoLabel(labelEntry) {
        return labelEntry[VERSION_INFO_LANG] || labelEntry.en;
    }

    /**
     * 「このスクリプトについて」とフッターに出す値を返す。SCRIPTMETA の項目を getScriptMetaValue() で直接読み、名前は SCRIPT_NAME。
     * getScriptMetaValue() が無いスクリプトは従来の SCRIPT_* を typeof で読む（未定義でも例外にならない）
     * @param {string} variableKey - "name" / "version" / "feature" / "readme" / "source" / "article"
     * @returns {string} 値（バージョンは「v1.2.0」の形）。無ければ空文字
     */
    function getVersionInfoValue(variableKey) {
        var infoValue = "";
        if (variableKey === "name") {
            infoValue = (typeof SCRIPT_NAME !== "undefined") ? SCRIPT_NAME : "";
        } else if (typeof getScriptMetaValue === "function") {
            var isJapanese = (VERSION_INFO_LANG === "ja");
            if (variableKey === "version") infoValue = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : "";
            else if (variableKey === "feature") infoValue = getScriptMetaValue("Name");
            else if (variableKey === "readme") infoValue = getScriptMetaValue(isJapanese ? "X-com.swwwitch-README-JA" : "X-com.swwwitch-README-EN");
            else if (variableKey === "source") infoValue = getScriptMetaValue("X-com.swwwitch-Source-URL");
            else if (variableKey === "article") infoValue = getScriptMetaValue("X-com.swwwitch-Article-URL");
        } else {
            if (variableKey === "version") infoValue = (typeof SCRIPT_VERSION !== "undefined") ? SCRIPT_VERSION : "";
            else if (variableKey === "feature") infoValue = (typeof SCRIPT_DESCRIPTION_JA !== "undefined") ? SCRIPT_DESCRIPTION_JA : "";
            else if (variableKey === "readme") {
                if (VERSION_INFO_LANG === "ja") infoValue = (typeof SCRIPT_README_JA !== "undefined") ? SCRIPT_README_JA : "";
                else infoValue = (typeof SCRIPT_README_EN !== "undefined") ? SCRIPT_README_EN : "";
            }
            else if (variableKey === "source") infoValue = (typeof SCRIPT_SOURCE_URL !== "undefined") ? SCRIPT_SOURCE_URL : "";
            else if (variableKey === "article") infoValue = (typeof SCRIPT_ARTICLE_URL !== "undefined") ? SCRIPT_ARTICLE_URL : "";
        }
        return (infoValue === undefined || infoValue === null) ? "" : String(infoValue);
    }

    /**
     * 左下のスクリプト名・バージョンをクリックしたら、「このスクリプトについて」のダイアログを開くようにする
     * @param {Group} versionCanvas - addVersionFooterCanvas() で作った左の欄
     * @returns {void}
     */
    function bindVersionInfoClick(versionCanvas) {
        versionCanvas.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.versionInfo);
        /* mousedown で開くと、離したときのイベントが閉じたあとの欄に届くので mouseup で開く
           Open on mouseup so the release does not land on the footer after the dialog closes */
        versionCanvas.addEventListener("mouseup", function () {
            showVersionInfoDialog(versionCanvas.window.type === "palette");
        });
    }

    /**
     * 「このスクリプトについて」のダイアログを開く。スクリプト名（コピーボタン付き）・機能・バージョンの［基本情報］パネル、
     * README・ソースコード・記事を開く［リンク］パネル、［Finder に表示］［閉じる］を並べる
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true（Illustrator のコピーを BridgeTalk 経由にする）
     * @returns {void}
     */
    function showVersionInfoDialog(isFromPalette) {
        var infoDialog = new Window("dialog", getVersionInfoLabel(VERSION_INFO_LABELS.dialog.title));
        infoDialog.orientation = "column";
        infoDialog.alignChildren = "fill";
        infoDialog.margins = VERSION_INFO_WINDOW_MARGINS;
        infoDialog.spacing = VERSION_INFO_WINDOW_SPACING;

        var infoPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.generalInfo);
        /* コピーボタンはスクリプト名だけ / Only the script name gets a copy button */
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptName, getVersionInfoValue("name"), isFromPalette);
        /* 機能は README の見出しで日本語だけなので、日本語の画面だけに出す / The feature is the Japanese README title, shown only in the Japanese UI */
        if (VERSION_INFO_LANG === "ja" && getVersionInfoValue("feature") !== "") {
            addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptFeature, getVersionInfoValue("feature"), null);
        }
        addVersionInfoRow(infoPanel, VERSION_INFO_LABELS.fieldLabel.scriptVersion, getVersionInfoValue("version"), null);
        alignVersionInfoLabels(infoPanel);

        /* URL が無いボタンはディム表示にして、並びはスクリプトごとに変えない / Dim buttons without a URL; keep the same buttons everywhere */
        var linksPanel = addVersionInfoPanel(infoDialog, VERSION_INFO_LABELS.panel.links);
        var linkButtonRow = linksPanel.add("group");
        linkButtonRow.orientation = "row";
        linkButtonRow.alignment = ["left", "center"];
        linkButtonRow.alignChildren = ["left", "center"];
        linkButtonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openReadme, getVersionInfoValue("readme"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openSourceCode, getVersionInfoValue("source"));
        addVersionInfoLinkButton(infoDialog, linkButtonRow, VERSION_INFO_LABELS.button.openArticle, getVersionInfoValue("article"));

        /* ボタン行：左に［Finder に表示］、右に［閉じる］ / Button row: Show in Finder at the left, Close at the right */
        var buttonRow = infoDialog.add("group");
        buttonRow.orientation = "row";
        buttonRow.alignment = ["fill", "bottom"];
        buttonRow.alignChildren = ["left", "center"];
        buttonRow.margins = VERSION_INFO_BUTTON_ROW_MARGINS;
        buttonRow.spacing = VERSION_INFO_BUTTON_SPACING;
        var revealLabel = (File.fs === "Windows") ? VERSION_INFO_LABELS.button.revealInExplorer : VERSION_INFO_LABELS.button.revealInFinder;
        var revealButton = buttonRow.add("button", undefined, getVersionInfoLabel(revealLabel));
        if (VERSION_INFO_SCRIPT_FILE.exists) {
            revealButton.helpTip = VERSION_INFO_SCRIPT_FILE.fsName;
            revealButton.onClick = function () {
                revealVersionInfoScriptFile(VERSION_INFO_SCRIPT_FILE);
                infoDialog.close();
            };
        } else {
            revealButton.enabled = false;
        }
        /* 余りの幅を吸って、［閉じる］を右端に寄せる / Absorbs the extra width so Close sits at the right edge */
        var buttonSpacer = buttonRow.add("group");
        buttonSpacer.alignment = ["fill", "fill"];
        buttonSpacer.minimumSize.width = 0;
        var closeButton = buttonRow.add("button", undefined, getVersionInfoLabel(VERSION_INFO_LABELS.button.close), { name: "ok" });
        closeButton.alignment = ["right", "center"];
        /* return でも esc でも閉じる / Close with both Return and Esc */
        infoDialog.defaultElement = closeButton;
        infoDialog.cancelElement = closeButton;
        infoDialog.onShow = function () {
            pinVersionInfoCopyButton(infoPanel);
        };
        infoDialog.show();
    }

    /**
     * 「このスクリプトについて」のダイアログにパネルを足す
     * @param {Window} infoDialog - 足す先のダイアログ
     * @param {Object} titleEntry - パネル名の文言 { ja, en }
     * @returns {Panel} パネル
     */
    function addVersionInfoPanel(infoDialog, titleEntry) {
        var infoPanel = infoDialog.add("panel", undefined, getVersionInfoLabel(titleEntry));
        infoPanel.orientation = "column";
        infoPanel.alignChildren = ["fill", "top"];
        infoPanel.alignment = "fill";
        infoPanel.margins = VERSION_INFO_PANEL_MARGINS;
        infoPanel.spacing = VERSION_INFO_WINDOW_SPACING;
        return infoPanel;
    }

    /**
     * ［基本情報］パネルに「項目名 : 値」の1行を足す（値は見せるだけなので statictext）
     * @param {Panel} infoPanel - 足す先のパネル
     * @param {Object} labelEntry - 項目名の文言 { ja, en }
     * @param {string} valueText - 表示する値
     * @param {boolean|null} copyFromPalette - 値の右にコピーボタンを置くなら true / false（常駐パレットから開いたか）、置かないなら null
     * @returns {StaticText} 値の statictext
     */
    function addVersionInfoRow(infoPanel, labelEntry, valueText, copyFromPalette) {
        var infoRow = infoPanel.add("group");
        infoRow.orientation = "row";
        infoRow.alignment = ["fill", "center"];
        infoRow.alignChildren = ["left", "center"];
        infoRow.spacing = VERSION_INFO_ROW_SPACING;
        /* 項目名のコロンは日本語が「 :」、英語が「:」 / Colon: " :" in Japanese, ":" in English */
        var rowLabel = infoRow.add("statictext", undefined, getVersionInfoLabel(labelEntry) + (VERSION_INFO_LANG === "ja" ? " :" : ":"));
        rowLabel.justify = "right";
        var valueLabel = infoRow.add("statictext", undefined, valueText);
        if (copyFromPalette === null) return valueLabel;
        /* statictext の fill では伸びないことがあるので、伸びるスペーサーでコピーボタンをパネルの右端に寄せる
           A stretching spacer (statictext fill is unreliable) pushes the copy button to the panel's right edge */
        var copySpacer = infoRow.add("group");
        copySpacer.alignment = ["fill", "fill"];
        copySpacer.minimumSize.width = 0;
        infoRow.copyButton = addVersionInfoCopyButton(infoRow, valueLabel, copyFromPalette);
        return valueLabel;
    }

    /**
     * コピーボタンを［基本情報］パネルの右端に置き直す（表示時、レイアウトが決まったあとに呼ぶ）。
     * 行の fill・スペーサーだけでは、項目の値が短い行がパネルの幅まで伸びないことがあるので、行の幅とボタンの位置を直接入れる
     * @param {Panel} infoPanel - ［基本情報］パネル
     * @returns {void}
     */
    function pinVersionInfoCopyButton(infoPanel) {
        var innerWidth = infoPanel.size.width - VERSION_INFO_PANEL_MARGINS[0] - VERSION_INFO_PANEL_MARGINS[2];
        for (var i = 0; i < infoPanel.children.length; i++) {
            var infoRow = infoPanel.children[i];
            if (!infoRow.copyButton || !infoRow.size || infoRow.size.width >= innerWidth) continue;
            infoRow.size = [innerWidth, infoRow.size.height];
            infoRow.copyButton.location = [innerWidth - infoRow.copyButton.size.width, infoRow.copyButton.location[1]];
        }
    }

    /**
     * ［基本情報］パネルの項目名を、いちばん長いものの幅にそろえる
     * @param {Panel} infoPanel - 対象のパネル
     * @returns {void}
     */
    function alignVersionInfoLabels(infoPanel) {
        var rowLabels = [];
        var labelWidth = 0;
        for (var i = 0; i < infoPanel.children.length; i++) {
            var rowLabel = infoPanel.children[i].children[0]; /* 行の先頭は項目名 / each row starts with its label */
            rowLabels.push(rowLabel);
            labelWidth = Math.max(labelWidth, rowLabel.graphics.measureString(rowLabel.text)[0]);
        }
        for (var j = 0; j < rowLabels.length; j++) {
            rowLabels[j].preferredSize.width = Math.ceil(labelWidth) + 2;
        }
    }

    /**
     * URL を開くボタンを足す。URL が無ければディム表示にする。開いたらダイアログは閉じる
     * @param {Window} infoDialog - ボタンを置くダイアログ
     * @param {Group} buttonGroup - 足す先のグループ
     * @param {Object} labelEntry - ボタンの文言 { ja, en }
     * @param {string} targetUrl - 開く URL（空ならディム表示）
     * @returns {void}
     */
    function addVersionInfoLinkButton(infoDialog, buttonGroup, labelEntry, targetUrl) {
        var linkButton = buttonGroup.add("button", undefined, getVersionInfoLabel(labelEntry));
        if (!targetUrl) {
            linkButton.enabled = false;
            return;
        }
        linkButton.helpTip = targetUrl;
        linkButton.onClick = function () {
            openVersionInfoUrl(targetUrl);
            infoDialog.close();
        };
    }

    /**
     * 値をクリップボードへコピーするアイコンのボタンを足す。コピーしたらアイコンをチェックマークに変え、ポインタが外れたら戻す
     * @param {Group} infoRow - 足す先の行
     * @param {StaticText} valueLabel - コピーする値の statictext
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {Group} ボタンとして使う group
     */
    function addVersionInfoCopyButton(infoRow, valueLabel, isFromPalette) {
        var buttonSize = VERSION_INFO_COPY_BUTTON_SIZE;
        var copyButton = infoRow.add("group");
        copyButton.preferredSize = [buttonSize, buttonSize];
        copyButton.minimumSize = [buttonSize, buttonSize];
        copyButton.maximumSize = [buttonSize, buttonSize];
        copyButton.helpTip = getVersionInfoLabel(VERSION_INFO_LABELS.tooltip.copyScriptName);
        copyButton.isPressed = false;
        copyButton.isCopied = false;

        copyButton.onDraw = function () {
            var buttonGraphics = copyButton.graphics;
            var fillColor = copyButton.isPressed ? VERSION_INFO_COPY_PRESSED_COLOR : VERSION_INFO_COPY_FILL_COLOR;
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0, 0, buttonSize, buttonSize);
            buttonGraphics.fillPath(buttonGraphics.newBrush(buttonGraphics.BrushType.SOLID_COLOR, fillColor));
            buttonGraphics.newPath();
            buttonGraphics.rectPath(0.5, 0.5, buttonSize - 1, buttonSize - 1);
            buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_FRAME_COLOR, 1));
            if (copyButton.isCopied) drawVersionInfoCheckIcon(buttonGraphics, buttonSize);
            else drawVersionInfoCopyIcon(buttonGraphics, buttonSize);
        };

        /**
         * 状態を変えて描き直す（group には notify() が無いので隠して再表示する）
         * @param {boolean} isPressed - 押下中なら true
         * @param {boolean} isCopied - コピー済みなら true
         * @returns {void}
         */
        function repaint(isPressed, isCopied) {
            if (copyButton.isPressed === isPressed && copyButton.isCopied === isCopied) return;
            copyButton.isPressed = isPressed;
            copyButton.isCopied = isCopied;
            copyButton.hide();
            copyButton.show();
        }
        copyButton.addEventListener("mousedown", function () { repaint(true, false); });
        copyButton.addEventListener("mouseup", function () {
            if (!copyButton.isPressed) return;
            var hasCopied = copyVersionInfoText(valueLabel.text, isFromPalette);
            repaint(false, hasCopied);
            if (!hasCopied) alert(getVersionInfoLabel(VERSION_INFO_LABELS.alert.copyFailed));
        });
        /* 押したまま外へ出たとき・コピーのあと離れたときは元の見た目に戻す / Reset when the pointer leaves */
        copyButton.addEventListener("mouseout", function () { repaint(false, false); });
        return copyButton;
    }

    /**
     * コピーのアイコン（後ろの四角に前の四角を重ねた形）を描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCopyIcon(buttonGraphics, buttonSize) {
        var iconPen = buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1);
        var sheetSize = 9;
        var offset = 3;
        var frontLeft = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var frontTop = Math.round((buttonSize - sheetSize - offset) / 2) + offset + 0.5;
        var backLeft = frontLeft - offset;
        var backTop = frontTop - offset;
        /* 後ろの四角は、前の四角に隠れる部分を描かない（上と左の辺だけ）/ Draw only the visible top and left of the back sheet */
        buttonGraphics.newPath();
        buttonGraphics.moveTo(frontLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop + sheetSize);
        buttonGraphics.lineTo(backLeft, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, backTop);
        buttonGraphics.lineTo(backLeft + sheetSize, frontTop);
        buttonGraphics.strokePath(iconPen);
        buttonGraphics.newPath();
        buttonGraphics.rectPath(frontLeft, frontTop, sheetSize, sheetSize);
        buttonGraphics.strokePath(iconPen);
    }

    /**
     * コピーしたあとのチェックマークを描く
     * @param {ScriptUIGraphics} buttonGraphics - 描画先
     * @param {number} buttonSize - ボタンの一辺
     * @returns {void}
     */
    function drawVersionInfoCheckIcon(buttonGraphics, buttonSize) {
        var center = buttonSize / 2;
        buttonGraphics.newPath();
        buttonGraphics.moveTo(center - 4.5, center);
        buttonGraphics.lineTo(center - 1.5, center + 3);
        buttonGraphics.lineTo(center + 4.5, center - 3.5);
        buttonGraphics.strokePath(buttonGraphics.newPen(buttonGraphics.PenType.SOLID_COLOR, VERSION_INFO_COPY_ICON_COLOR, 1.5));
    }

    /**
     * 文字列をクリップボードにコピーする。
     * InDesign は AppleScript（Windows は VBScript）の doScript で書き込む。
     * Illustrator は文字列を直接送る API が無いので、一時レイヤーのテキストフレームを選んでコピーする（プレーンテキストも入る）。
     * Illustrator の常駐パレットからは DOM に触れないので、同じ処理を BridgeTalk でメインエンジンに渡す
     * @param {string} copyText - コピーする文字列
     * @param {boolean} isFromPalette - 常駐パレットから開いたなら true
     * @returns {boolean} コピーの操作まで進めたら true
     */
    function copyVersionInfoText(copyText, isFromPalette) {
        try {
            if (VERSION_INFO_IS_INDESIGN) {
                if (File.fs === "Windows") {
                    app.doScript('CreateObject("htmlfile").parentWindow.clipboardData.setData "text", "' + copyText.replace(/"/g, '""') + '"', ScriptLanguage.VISUAL_BASIC);
                } else {
                    app.doScript('set the clipboard to "' + copyText.replace(/\\/g, "\\\\").replace(/"/g, '\\"') + '"', ScriptLanguage.APPLESCRIPT_LANGUAGE);
                }
                return true;
            }
            var copyScript = buildVersionInfoCopyScript(copyText);
            if (!isFromPalette) return eval(copyScript) === true;
            var copyMessage = new BridgeTalk();
            copyMessage.target = BridgeTalk.appSpecifier; /* 複数バージョンを起動していても、今のバージョンに送る / the running version, even with several installed */
            copyMessage.body = copyScript;
            copyMessage.send();
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * Illustrator でコピーする処理を文字列で組み立てる（BridgeTalk でも送れるよう、関数の toString() は使わない）。
     * 一時レイヤーにテキストフレームを作り、フレームを選んでメニューコマンドでコピーして、レイヤーごと消す。
     * 文字として選択（textRange.select()）してコピーすると「オブジェクトをコピーできません」の警告が出るので使わない。
     * 選択と作業レイヤーは元に戻す。ドキュメントが無ければ一時ドキュメントを作って閉じる
     * @param {string} copyText - コピーする文字列
     * @returns {string} 実行すると true / false を返すスクリプト
     */
    function buildVersionInfoCopyScript(copyText) {
        return [
            "(function (copyText) {",
            "    var usingTempDoc = (app.documents.length === 0);",
            "    var targetDoc = usingTempDoc ? app.documents.add() : app.activeDocument;",
            "    var savedSelection = (!usingTempDoc && targetDoc.selection instanceof Array) ? targetDoc.selection : [];",
            "    var savedActiveLayer = usingTempDoc ? null : targetDoc.activeLayer;",
            "    var tempLayer = null;",
            "    var hasCopied = false;",
            "    try {",
            "        tempLayer = targetDoc.layers.add();",
            "        var tempFrame = tempLayer.textFrames.add();",
            "        tempFrame.contents = copyText;",
            "        app.redraw();",
            "        app.executeMenuCommand('deselectall');",
            "        tempFrame.selected = true;",
            "        app.redraw();",
            "        app.executeMenuCommand('copy');",
            "        app.redraw();",
            "        hasCopied = true;",
            "    } catch (e) { }",
            "    if (tempLayer) { try { tempLayer.remove(); } catch (e) { } }",
            "    if (usingTempDoc) { targetDoc.close(SaveOptions.DONOTSAVECHANGES); return hasCopied; }",
            "    try { targetDoc.activeLayer = savedActiveLayer; } catch (e) { }",
            "    targetDoc.selection = null;",
            "    for (var i = 0; i < savedSelection.length; i++) { try { savedSelection[i].selected = true; } catch (e) { } }",
            "    app.redraw();",
            "    return hasCopied;",
            "})(" + quoteVersionInfoString(copyText) + ");"
        ].join(" "); /* 行コメントを含まないので1行につなぐ（BridgeTalk の転送で改行が化けない）/ joined on one line; there are no line comments */
    }

    /**
     * 文字列を、スクリプトに埋め込める二重引用符のリテラルにする
     * @param {string} sourceText - 文字列
     * @returns {string} "…" の形のリテラル
     */
    function quoteVersionInfoString(sourceText) {
        return '"' + String(sourceText).replace(/\\/g, "\\\\").replace(/"/g, '\\"').replace(/\r/g, "\\r").replace(/\n/g, "\\n") + '"';
    }

    /**
     * ファイルを選択した状態で表示する。OpenInFileViewer.app があれば Path Finder が起動中なら Path Finder、そうでなければ Finder で表示し、
     * 無ければ RevealInFinder.app で Finder に表示する。どちらも無いときは入っているフォルダーを開く
     * @param {File} scriptFile - 表示するファイル
     * @returns {void}
     */
    function revealVersionInfoScriptFile(scriptFile) {
        if (launchVersionInfoHelperApp(VERSION_INFO_FILE_VIEWER_APP_PATH, VERSION_INFO_FILE_VIEWER_PATH_FILE, scriptFile.fsName)) return;
        if (launchVersionInfoHelperApp(VERSION_INFO_REVEAL_APP_PATH, VERSION_INFO_REVEAL_PATH_FILE, scriptFile.fsName)) return;
        /* アプリが無い・起動できないときは、何も起きないままにせずフォルダーを開く / Otherwise open the enclosing folder */
        scriptFile.parent.execute();
    }

    /**
     * 補助アプリが読む一時ファイルにパスを書き出してから、アプリを起動する
     * @param {string} appPath - 補助アプリ（.app）のパス
     * @param {string} pathFilePath - アプリが読む一時ファイルのパス
     * @param {string} targetPath - 表示するファイルの絶対パス
     * @returns {boolean} 起動できたら true。Mac 以外・アプリが無い・起動できないときは false
     */
    function launchVersionInfoHelperApp(appPath, pathFilePath, targetPath) {
        /* .app は実体がフォルダーなので File.exists が false になる環境がある。Folder でも確かめる / An .app is a folder, so check with Folder too */
        if (File.fs !== "Macintosh" || !(new Folder(appPath).exists || new File(appPath).exists)) return false;
        return writeVersionInfoRevealPath(pathFilePath, targetPath) && new File(appPath).execute();
    }

    /**
     * 補助アプリに渡すパスを一時ファイルへ書き出す（UTF-8、改行なし）
     * @param {string} pathFilePath - 一時ファイルのパス
     * @param {string} targetPath - 絶対パス
     * @returns {boolean} 書き出せたら true
     */
    function writeVersionInfoRevealPath(pathFilePath, targetPath) {
        var pathFile = new File(pathFilePath);
        pathFile.encoding = "UTF-8";
        pathFile.lineFeed = "Unix";
        if (!pathFile.open("w")) return false;
        var hasWritten = pathFile.write(targetPath);
        pathFile.close();
        return hasWritten;
    }

    /**
     * URL を既定のブラウザーで開く（一時的なショートカットファイルを作って開く）
     * @param {string} targetUrl - 開く URL
     * @returns {void}
     */
    function openVersionInfoUrl(targetUrl) {
        var isMac = (File.fs === "Macintosh");
        var shortcutFile = new File(Folder.temp.fsName + "/VersionInfoLink" + (isMac ? ".webloc" : ".url"));
        shortcutFile.encoding = "UTF-8";
        shortcutFile.open("w");
        if (isMac) {
            var escapedUrl = targetUrl.replace(/&/g, "&amp;").replace(/</g, "&lt;");
            shortcutFile.write('<?xml version="1.0" encoding="UTF-8"?>\n'
                + '<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">\n'
                + '<plist version="1.0"><dict><key>URL</key><string>' + escapedUrl + '</string></dict></plist>\n');
        } else {
            shortcutFile.write("[InternetShortcut]\r\nURL=" + targetUrl + "\r\n");
        }
        shortcutFile.close();
        shortcutFile.execute();
    }

    // バージョン表示（再利用パーツ）ここまで / End of the reusable version footer

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "線のグラフ", en: "Stroke Graph" }
        },
        panel: {
            color: { ja: "カラー", en: "Color" },
            stroke: { ja: "線", en: "Stroke" }
        },
        fieldLabel: {
            baseColor: { ja: "ベース", en: "Base" },
            accentColor: { ja: "アクセント", en: "Accent" },
            accentPercent: { ja: "アクセントの割合", en: "Accent Ratio" },
            strokeWidth: { ja: "線幅", en: "Weight" }
        },
        colorText: {
            none: { ja: "なし", en: "None" },
            unknown: { ja: "（不明）", en: "(Unknown)" }
        },
        checkbox: {
            reverse: { ja: "方向反転", en: "Reverse Direction" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            reverse: {
                ja: "パスの方向を反転して、ベースとアクセントの位置を入れ替えます",
                en: "Reverses the path direction, swapping where the base and accent go"
            },
            colorChip: { ja: "クリックしてカラーを変更", en: "Click to change the color" },
            accentPercent: {
                ja: "線の終点側からアクセントにする長さ（20% なら境界は 80%）",
                en: "Length colored with the accent, from the end of the stroke (20% puts the boundary at 80%)"
            },
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
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            noPath: { ja: "パスを選択してください。", en: "Please select a path." },
            invalidValues: {
                ja: "アクセントの割合は 0% より大きく 100% より小さい値を、線幅は 0 より大きい値を入力してください。",
                en: "Enter an accent ratio greater than 0% and less than 100%, and a weight greater than 0."
            },
            dashAlignFailed: {
                ja: "［コーナーやパス先端に破線の先端を整列］を設定できませんでした。破線と丸型線端だけ設定しました。",
                en: "Could not set Align Dashes to Corners and Path Ends. Only the dashes and round caps were set."
            },
            restoreFailed: {
                ja: "埋め込んだファイルを書き出せませんでした。\n{fileName}",
                en: "Could not write out the embedded file.\n{fileName}"
            },
            openFailed: {
                ja: "埋め込んだファイルを開けませんでした。\n{fileName}",
                en: "Could not open the embedded file.\n{fileName}"
            },
            styleNotFound: {
                ja: "グラフィックスタイル「{styleName}」が見つかりません。",
                en: "The graphic style \"{styleName}\" was not found."
            }
        }
    };

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * 単位の値を pt に換算する
     * @param {number} value - 単位の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} pt の値
     */
    function unitToPt(value, unit) {
        return value * unit.pointsPerUnit;
    }

    /**
     * pt の値を単位の値に換算する
     * @param {number} ptValue - pt の値
     * @param {{pointsPerUnit: number}} unit - getUnitInfo() の戻り値か UNITS[i]
     * @returns {number} 単位の値
     */
    function ptToUnit(ptValue, unit) {
        return ptValue / unit.pointsPerUnit;
    }

    // =========================================
    // 破線の一時アクション / Dash temporary action
    // =========================================

    var DASH_ACTION_SET_NAME = "SwwwitchTempDashSet";
    var DASH_ACTION_NAME     = "SwwwitchTempDash";
    var ELLIPSE_ACTION_SET_NAME = "SwwwitchTempEllipseSet";
    var ELLIPSE_ACTION_NAME     = "SwwwitchTempEllipseAngle";
    var ACTION_UNIT_POINT    = 592476268;  /* ポイント / point（parameter /unit） */

    /* 線の設定（ai_plugin_setStroke）のパラメーターキー。記録した .aia から採取 / Parameter keys of ai_plugin_setStroke, taken from a recorded .aia */
    var STROKE_KEY_WIDTH      = 2003072104; /* wdth: 線幅 / stroke width */
    var STROKE_KEY_CAP        = 1667330094; /* cap : 線端 / cap */
    var STROKE_KEY_MITER      = 1836344690; /* mitr: 比率 / miter limit */
    var STROKE_KEY_JOIN       = 1785686382; /* join: 角の形状 / join */
    var STROKE_KEY_DASH_COUNT = 1684825454; /* dlen: 破線の値の数 / number of dash values */
    var STROKE_KEY_DASH_1     = 1685284913; /* dsh1: 線分 / first dash */
    var STROKE_KEY_DASH_ADJUST = 1684104298; /* dadj: コーナーやパス先端に破線の先端を整列 / align dashes to corners and path ends */
    var STROKE_KEY_ARROW_1    = 1634231345; /* ahd1: 始点の矢印 / start arrowhead */
    var STROKE_KEY_ARROW_2    = 1634231346; /* ahd2: 終点の矢印 / end arrowhead */
    var STROKE_KEY_TIP_ALIGN  = 1634230636; /* ahal: 矢印の配置 / arrow tip alignment */
    var STROKE_KEY_ALIGN      = 1634494318; /* algn: 線の位置 / stroke alignment */

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel("alert.noDocument"));
        return;
    }
    var activeDoc = app.activeDocument;
    if (!activeDoc.selection || activeDoc.selection.length === 0) {
        alert(getLabel("alert.noSelection"));
        return;
    }
    var graphPaths = collectGraphPaths(activeDoc.selection);
    if (graphPaths.length === 0) {
        alert(getLabel("alert.noPath"));
        return;
    }
    runStrokeGraph(activeDoc, graphPaths);

    // =========================================
    // 線のグラフ / Stroke graph
    // =========================================

    /**
     * 選択からパスを集める（グループ・複合パスの中も。ガイドとクリッピングパスは外す）
     * @param {*} selectionSource - doc.selection
     * @returns {PathItem[]} パス
     */
    function collectGraphPaths(selectionSource) {
        return collectSelectionItems(selectionSource, {
            accept: function (item) { return item.typename === "PathItem"; },
            enterCompoundPaths: true,
            skipGuides: true,
            skipClipMasks: true
        });
    }

    /**
     * パスの線がグラデーションかを返す
     * @param {PathItem} pathItem - パス
     * @returns {boolean} 線がグラデーションなら true
     */
    function hasGradientStroke(pathItem) {
        try {
            return pathItem.stroked && pathItem.strokeColor.typename === "GradientColor";
        } catch (e) {
            return false;
        }
    }

    /**
     * ベース・アクセント・線幅の初期値を返す。線がグラデーションなら始点側の色をベース、終点側の色をアクセント、今の線幅にし、
     * そうでなければ DEFAULT_BASE_COLOR をベース、パスの色（線、無ければ塗りの単色。どちらも無ければ DEFAULT_ACCENT_COLOR）をアクセント、
     * DEFAULT_STROKE_WIDTH を線幅にする
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} pathItem - 最初のパス
     * @returns {{baseColor: Color, accentColor: Color, accentPercent: number, strokeWidth: number}} 初期値（線幅は pt）
     */
    function getInitialGraphValues(doc, pathItem) {
        if (hasGradientStroke(pathItem)) {
            var bands = extractColorBands(readGradientStops(pathItem.strokeColor.gradient));
            return {
                baseColor: bands[0].color,
                accentColor: bands[bands.length - 1].color,
                accentPercent: getInitialAccent(bands),
                strokeWidth: pathItem.strokeWidth
            };
        }
        return {
            baseColor: createDocumentColor(doc, DEFAULT_BASE_COLOR),
            accentColor: readSolidPathColor(pathItem) || createDocumentColor(doc, DEFAULT_ACCENT_COLOR),
            accentPercent: DEFAULT_ACCENT_PERCENT,
            strokeWidth: DEFAULT_STROKE_WIDTH
        };
    }

    /**
     * パスの単色を返す。線の色を優先し、線が無ければ塗りの色（パターン・グラデーション・なしは使わない）
     * @param {PathItem} pathItem - パス
     * @returns {Color|null} 色。使える色が無ければ null
     */
    function readSolidPathColor(pathItem) {
        if (pathItem.stroked && isSolidColor(pathItem.strokeColor)) return pathItem.strokeColor;
        if (pathItem.filled && isSolidColor(pathItem.fillColor)) return pathItem.fillColor;
        return null;
    }

    /**
     * ベース・アクセントに使える単色かを返す
     * @param {Color} color - 色
     * @returns {boolean} 単色なら true
     */
    function isSolidColor(color) {
        var colorType = color.typename;
        return colorType !== "NoColor" && colorType !== "PatternColor" && colorType !== "GradientColor";
    }

    /**
     * ドキュメントのカラーモードに合わせた色を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {{cmyk: number[], rgb: number[]}} colorValues - CMYK と RGB の値
     * @returns {Color} CMYKColor か RGBColor
     */
    function createDocumentColor(doc, colorValues) {
        if (doc.documentColorSpace === DocumentColorSpace.RGB) {
            var rgbColor = new RGBColor();
            rgbColor.red = colorValues.rgb[0];
            rgbColor.green = colorValues.rgb[1];
            rgbColor.blue = colorValues.rgb[2];
            return rgbColor;
        }
        var cmykColor = new CMYKColor();
        cmykColor.cyan = colorValues.cmyk[0];
        cmykColor.magenta = colorValues.cmyk[1];
        cmykColor.yellow = colorValues.cmyk[2];
        cmykColor.black = colorValues.cmyk[3];
        return cmykColor;
    }

    /**
     * 線のグラフにする。先にパスを線のグラフの形にそろえ（prepareGraphPaths）、
     * ダイアログボックスの入力に合わせてストップと線幅を書き込み、［キャンセル］ではすべて元に戻す（revertGraphPaths）
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem[]} graphPaths - collectGraphPaths() の戻り値
     * @returns {void}
     */
    function runStrokeGraph(doc, graphPaths) {
        /* 初期値は、そろえる前の最初のパスから / defaults from the first path, before it is changed */
        var initialValues = getInitialGraphValues(doc, graphPaths[0]);
        var baseChoice = { label: labelText("fieldLabel.baseColor"), color: initialValues.baseColor };
        var accentChoice = { label: labelText("fieldLabel.accentColor"), color: initialValues.accentColor };
        var graphState = prepareGraphPaths(doc, graphPaths);

        /**
         * ベース・アクセントの2色のストップと線幅を書き込む
         * @param {{accentPercent: number, strokeWidth: number}} graphValues - アクセントの割合と線幅（pt）
         * @returns {void}
         */
        function applyGraphValues(graphValues) {
            writeGraphValues(graphState.gradientSwaps, [baseChoice.color, accentChoice.color], graphValues);
        }

        var graphValues = showStrokeGraphDialog({
            colorChoices: [baseChoice, accentChoice],
            initialAccent: initialValues.accentPercent,
            initialStrokeWidth: initialValues.strokeWidth,
            initialReverse: graphState.isReversed,
            onPreview: applyGraphValues,
            onReverseChange: function (shouldReverse) { setPathsReversed(graphState, shouldReverse); }
        });
        if (graphValues) applyGraphValues(graphValues);
        else revertGraphPaths(graphState);
    }

    /**
     * パスを線のグラフの形にそろえる：塗りをなしにし、パスごとの専用のグラデーションへ付け替え、破線・丸型線端を入れ、
     * 方向を反転して楕円の角度を入れる。［キャンセル］で戻すための控えを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem[]} graphPaths - collectGraphPaths() の戻り値
     * @returns {{graphPaths: PathItem[], strokeSnapshots: Object[], anglePaths: PathItem[], fillSnapshots: Object[], gradientSwaps: Object[], isReversed: boolean}} 控え
     */
    function prepareGraphPaths(doc, graphPaths) {
        var graphState = { graphPaths: graphPaths, isReversed: false };
        graphState.strokeSnapshots = saveStrokeSnapshots(graphPaths);
        /* 楕円の角度がすでに ELLIPSE_ANGLE のパスは角度を変えない（［キャンセル］でも戻さない） / leave paths already at ELLIPSE_ANGLE alone */
        graphState.anglePaths = collectPathsNotAtAngle(graphPaths, ELLIPSE_ANGLE);
        /* 塗りはなしにする（［キャンセル］で戻す） / remove fills; Cancel brings them back */
        graphState.fillSnapshots = removeFills(graphPaths);
        /* 線にグラデーションが無いパスがあるときだけ、［パスに沿って…］のスタイルを用意する / prepare the along-stroke style only when needed */
        var strokeGradientStyle = hasPathWithoutGradientStroke(graphPaths) ? getStrokeGradientStyle(doc) : null;
        graphState.gradientSwaps = attachOwnGradients(doc, graphPaths, strokeGradientStyle);
        if (!applyDashSettings(graphPaths)) alert(getLabel("alert.dashAlignFailed"));
        setPathsReversed(graphState, true);
        setEllipseAngle(graphState.anglePaths, ELLIPSE_ANGLE);
        return graphState;
    }

    /**
     * prepareGraphPaths() でそろえたパスを元に戻す（掛けた順の逆）
     * @param {Object} graphState - prepareGraphPaths() の戻り値
     * @returns {void}
     */
    function revertGraphPaths(graphState) {
        setEllipseAngle(graphState.anglePaths, ELLIPSE_ANGLE_ON_CANCEL);
        setPathsReversed(graphState, false);
        restoreStrokeSnapshots(graphState.strokeSnapshots, true);
        detachOwnGradients(graphState.gradientSwaps);
        restoreFills(graphState.fillSnapshots);
    }

    /**
     * ［方向反転］に合わせてパスの方向を反転する（今の状態と同じなら何もしない）
     * @param {Object} graphState - prepareGraphPaths() の控え（isReversed を書き換える）
     * @param {boolean} shouldReverse - 反転するなら true
     * @returns {void}
     */
    function setPathsReversed(graphState, shouldReverse) {
        if (shouldReverse === graphState.isReversed) return;
        reversePathDirections(graphState.graphPaths);
        graphState.isReversed = shouldReverse;
    }

    /**
     * パスごとの専用のグラデーションに2色のストップを、パスに線幅を書き込む
     * @param {Object[]} gradientSwaps - attachOwnGradients() の戻り値
     * @param {Color[]} bandColors - [ベース, アクセント]
     * @param {{accentPercent: number, strokeWidth: number}} graphValues - アクセントの割合と線幅（pt）
     * @returns {void}
     */
    function writeGraphValues(gradientSwaps, bandColors, graphValues) {
        var bands = [{ color: bandColors[0], opacity: 100 }, { color: bandColors[1], opacity: 100 }];
        var graphStops = buildSeparateStops(bands, [100 - graphValues.accentPercent]);
        for (var i = 0; i < gradientSwaps.length; i++) {
            var gradientSwap = gradientSwaps[i];
            /* ロック中などで書けなくても、残りは続ける / keep going when one path refuses */
            try {
                writeGradientStops(gradientSwap.ownGradient.gradientStops, graphStops);
                /* 入れ直して、パスに反映させる / reassign so the path picks up the change */
                var strokeGradient = gradientSwap.item.strokeColor;
                gradientSwap.item.strokeColor = strokeGradient;
                gradientSwap.item.strokeWidth = graphValues.strokeWidth;
            } catch (e) { $.writeln("writeGraphValues: " + e); }
        }
        app.redraw();
    }

    /**
     * パスの塗りをなしにする。塗りのあったパスだけ、元の塗りを控えて返す
     * @param {PathItem[]} graphPaths - パス
     * @returns {{item: PathItem, fillColor: Color}[]} 塗りをなしにしたパスと元の塗り
     */
    function removeFills(graphPaths) {
        var fillSnapshots = [];
        for (var i = 0; i < graphPaths.length; i++) {
            var graphPath = graphPaths[i];
            /* ロック中などで変えられなくても、残りは続ける / keep going when one path refuses */
            try {
                if (!graphPath.filled) continue;
                fillSnapshots.push({ item: graphPath, fillColor: graphPath.fillColor });
                graphPath.filled = false;
            } catch (e) { $.writeln("removeFills: " + e); }
        }
        return fillSnapshots;
    }

    /**
     * removeFills() で消した塗りを戻す
     * @param {{item: PathItem, fillColor: Color}[]} fillSnapshots - removeFills() の戻り値
     * @returns {void}
     */
    function restoreFills(fillSnapshots) {
        for (var i = 0; i < fillSnapshots.length; i++) {
            try {
                fillSnapshots[i].item.fillColor = fillSnapshots[i].fillColor;
                fillSnapshots[i].item.filled = true;
            } catch (e) { $.writeln("restoreFills: " + e); }
        }
        app.redraw();
    }

    /**
     * アクセントの割合の初期値を返す。すでに2色に分かれていれば、今の境界から出す
     * @param {Object[]} bands - extractColorBands() の戻り値
     * @returns {number} アクセントの割合
     */
    function getInitialAccent(bands) {
        if (bands.length !== 2 || bands[1].start - bands[0].end > SEPARATE_EDGE_MAX_GAP) return DEFAULT_ACCENT_PERCENT;
        var boundary = (bands[0].end + bands[1].start) / 2;
        return Math.round((100 - boundary) * 100) / 100;
    }

    /**
     * グラデーションのストップを位置の順に読む
     * @param {Gradient} gradient - グラデーション
     * @returns {Object[]} {color, rampPoint, opacity, midPoint} の配列（位置の昇順）
     */
    function readGradientStops(gradient) {
        var orderedStops = getStopsByPosition(gradient.gradientStops);
        var stops = [];
        for (var i = 0; i < orderedStops.length; i++) {
            stops.push({
                color: orderedStops[i].color,
                rampPoint: orderedStops[i].rampPoint,
                opacity: orderedStops[i].opacity,
                midPoint: orderedStops[i].midPoint
            });
        }
        return stops;
    }

    /**
     * パスごとに専用のグラデーションを作って線に付ける。
     * 線がグラデーションなら今のストップを写して差し替える（複製したオブジェクトどうしは同じグラデーションを共有しており、
     * そのままストップを書き換えると選択していない方も変わるため）。
     * グラデーションでなければ、［パスに沿ってグラデーションを適用］の線のグラフィックスタイルを先に適用し、その線を同じように差し替える
     * （線のグラデーションの種類は DOM で設定できないため）。スタイルが無いか適用できなければ、線をオンにして新しく塗る
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem[]} graphPaths - パス
     * @param {GraphicStyle|null} strokeGradientStyle - getStrokeGradientStyle() の戻り値（用意できなければ null）
     * @returns {Object[]} {item, ownGradient, originalGradient, originalStroked, originalStrokeColor} の配列
     *     （originalGradient はもとからグラデーションだった線だけ。付けられなかったパスは入らない）
     */
    function attachOwnGradients(doc, graphPaths, strokeGradientStyle) {
        var gradientSwaps = [];
        for (var i = 0; i < graphPaths.length; i++) {
            var graphPath = graphPaths[i];
            try {
                var gradientSwap = { item: graphPath, ownGradient: null, originalGradient: null, originalStroked: graphPath.stroked, originalStrokeColor: graphPath.strokeColor };
                var hadGradientStroke = hasGradientStroke(graphPath);
                if (!hadGradientStroke && strokeGradientStyle) applyStyleKeepingOpacity(strokeGradientStyle, graphPath);
                var ownGradient = doc.gradients.add();
                ownGradient.type = GradientType.LINEAR;
                if (hasGradientStroke(graphPath)) {
                    var strokeGradient = graphPath.strokeColor;
                    /* スタイルで付けたグラデーションは、［キャンセル］で元の単色・線なしに戻す / a gradient from the style goes back to the original stroke on Cancel */
                    if (hadGradientStroke) gradientSwap.originalGradient = strokeGradient.gradient;
                    ownGradient.type = strokeGradient.gradient.type;
                    writeGradientStops(ownGradient.gradientStops, readGradientStops(strokeGradient.gradient));
                    /* 角度などの設定は GradientColor に残るので、グラデーションだけ差し替える / swap only the gradient; angle and the rest stay on the GradientColor */
                    strokeGradient.gradient = ownGradient;
                    graphPath.strokeColor = strokeGradient;
                } else {
                    var newGradientColor = new GradientColor();
                    newGradientColor.gradient = ownGradient;
                    graphPath.stroked = true;
                    graphPath.strokeColor = newGradientColor;
                }
                ownGradient.name = uniqueName(GRAPH_GRADIENT_BASE_NAME, doc.gradients);
                gradientSwap.ownGradient = ownGradient;
                gradientSwaps.push(gradientSwap);
            } catch (e) {
                $.writeln("attachOwnGradients: " + e);
            }
        }
        return gradientSwaps;
    }

    /**
     * attachOwnGradients() で付けたパスを元の線（グラデーションか単色・線なし）に戻し、専用のグラデーションを削除する
     * @param {Object[]} gradientSwaps - attachOwnGradients() の戻り値
     * @returns {void}
     */
    function detachOwnGradients(gradientSwaps) {
        for (var i = 0; i < gradientSwaps.length; i++) {
            var gradientSwap = gradientSwaps[i];
            try {
                if (gradientSwap.originalGradient) {
                    var strokeGradient = gradientSwap.item.strokeColor;
                    strokeGradient.gradient = gradientSwap.originalGradient;
                    gradientSwap.item.strokeColor = strokeGradient;
                } else {
                    gradientSwap.item.strokeColor = gradientSwap.originalStrokeColor;
                    gradientSwap.item.stroked = gradientSwap.originalStroked;
                }
                gradientSwap.ownGradient.remove();
            } catch (e) {
                $.writeln("detachOwnGradients: " + e);
            }
        }
        app.redraw();
    }

    // =========================================
    // セパレートのストップ / Segmented stops
    // =========================================

    /**
     * 境界の位置が使えるか調べる（0〜100% の内側で、昇順に並び、同じ位置が無い）
     * @param {number[]} boundaries - 境界の位置（%）
     * @returns {boolean} 使えるなら true
     */
    function areBoundariesValid(boundaries) {
        var previousEdge = 0;
        for (var i = 0; i < boundaries.length; i++) {
            var boundary = boundaries[i];
            if (typeof boundary !== "number" || isNaN(boundary)) return false;
            if (boundary <= previousEdge) return false;
            previousEdge = boundary;
        }
        return previousEdge < 100;
    }

    /**
     * 色の帯の並びと境界の位置から、セパレートグラデーションのストップ列を作る。
     * 両端に1つずつ、境界ごとに前の色と次の色を同じ位置に置く（間に中間色が出ないように）
     * @param {Object[]} bands - {color, opacity} の配列（始点側から順）
     * @param {number[]} boundaries - 境界の位置（%、bands.length - 1 個、昇順）
     * @returns {Object[]} {color, rampPoint, opacity, midPoint} の配列（位置の昇順）
     */
    function buildSeparateStops(bands, boundaries) {
        var stops = [{ color: bands[0].color, rampPoint: 0, opacity: bands[0].opacity, midPoint: 50 }];
        for (var i = 0; i < boundaries.length; i++) {
            stops.push({ color: bands[i].color, rampPoint: boundaries[i], opacity: bands[i].opacity, midPoint: 50 });
            stops.push({ color: bands[i + 1].color, rampPoint: boundaries[i], opacity: bands[i + 1].opacity, midPoint: 50 });
        }
        var lastBand = bands[bands.length - 1];
        stops.push({ color: lastBand.color, rampPoint: 100, opacity: lastBand.opacity, midPoint: 50 });
        return stops;
    }

    /**
     * ストップ列を、同じ色・不透明度が続く「色の帯」にまとめる。
     * なめらかなグラデーションはストップ1つが1帯、セパレートは同じ色の2つのストップが1帯になる
     * @param {Object[]} stops - readGradientStops() か buildSeparateStops() の戻り値（位置の昇順）
     * @returns {Object[]} {color, opacity, start, end} の配列（始点側から順）
     */
    function extractColorBands(stops) {
        var bands = [];
        var lastKey = null;
        for (var i = 0; i < stops.length; i++) {
            var bandKey = describeColor(stops[i].color) + "@" + stops[i].opacity;
            if (bandKey === lastKey) {
                bands[bands.length - 1].end = stops[i].rampPoint;
                continue;
            }
            lastKey = bandKey;
            bands.push({ color: stops[i].color, opacity: stops[i].opacity, start: stops[i].rampPoint, end: stops[i].rampPoint });
        }
        return bands;
    }

    /**
     * ストップに変更後の値を書き込む（数をそろえ、位置を隣を追い越さない順で動かしてから、位置の順に色・不透明度・中間点を書く）
     * @param {GradientStops} gradientStops - 書き込む先のストップ
     * @param {Object[]} stopsToWrite - readGradientStops() か buildSeparateStops() の戻り値（位置の昇順）
     * @returns {void}
     */
    function writeGradientStops(gradientStops, stopsToWrite) {
        /* 複製・セパレートで増えた分を足し、元に戻すときは減らす / add stops for duplicates and segments, remove them when rolling back */
        while (gradientStops.length < stopsToWrite.length) gradientStops.add();
        while (gradientStops.length > stopsToWrite.length) gradientStops[gradientStops.length - 1].remove();
        var orderedStops = getStopsByPosition(gradientStops);
        moveStopsWithoutOverlap(orderedStops, stopsToWrite);
        var stopCount = orderedStops.length;
        for (var i = 0; i < stopCount; i++) {
            orderedStops[i].color = stopsToWrite[i].color;
            orderedStops[i].opacity = stopsToWrite[i].opacity;
            /* 最後のストップの中間点は使われないので、書けなくても続ける / the last stop's midpoint is unused */
            try { orderedStops[i].midPoint = stopsToWrite[i].midPoint; } catch (e) { }
        }
    }

    /**
     * グラデーションのストップを今の位置の昇順に並べた配列を返す（同じ位置は元の順のまま）
     * @param {GradientStops} gradientStops - グラデーションのストップ
     * @returns {GradientStop[]} 位置の昇順のストップ
     */
    function getStopsByPosition(gradientStops) {
        var orderedStops = [];
        for (var i = 0; i < gradientStops.length; i++) {
            var gradientStop = gradientStops[i];
            var insertAt = orderedStops.length;
            while (insertAt > 0 && orderedStops[insertAt - 1].rampPoint > gradientStop.rampPoint) insertAt--;
            orderedStops.splice(insertAt, 0, gradientStop);
        }
        return orderedStops;
    }

    /**
     * ストップを、隣のストップを追い越さない順で目標の位置へ動かす。
     * Illustrator は前後のストップを越える位置を入れると例外になるので、
     * 始点側へ動くものは始点側から、終点側へ動くものは終点側から順に動かす
     * @param {GradientStop[]} orderedStops - 位置の昇順のストップ
     * @param {Object[]} targetStops - 目標の {rampPoint} の配列（位置の昇順、orderedStops と同じ数）
     * @returns {void}
     */
    function moveStopsWithoutOverlap(orderedStops, targetStops) {
        var i;
        for (i = 0; i < orderedStops.length; i++) {
            if (targetStops[i].rampPoint < orderedStops[i].rampPoint) orderedStops[i].rampPoint = targetStops[i].rampPoint;
        }
        for (i = orderedStops.length - 1; i >= 0; i--) {
            if (targetStops[i].rampPoint > orderedStops[i].rampPoint) orderedStops[i].rampPoint = targetStops[i].rampPoint;
        }
    }

    // =========================================
    // 破線と線の向き / Dashes and path direction
    // =========================================

    /**
     * 線の設定を控える（アクションで変わる線幅・角の形状・比率と、［キャンセル］で戻す破線・線端）
     * @param {PathItem[]} strokeItems - collectGraphPaths() の戻り値
     * @returns {Object[]} {item, strokeWidth, strokeJoin, strokeMiterLimit, strokeDashes, strokeCap} の配列
     */
    function saveStrokeSnapshots(strokeItems) {
        var strokeSnapshots = [];
        for (var i = 0; i < strokeItems.length; i++) {
            var strokeItem = strokeItems[i];
            strokeSnapshots.push({
                item: strokeItem,
                strokeWidth: strokeItem.strokeWidth,
                strokeJoin: strokeItem.strokeJoin,
                strokeMiterLimit: strokeItem.strokeMiterLimit,
                strokeDashes: strokeItem.strokeDashes,
                strokeCap: strokeItem.strokeCap
            });
        }
        return strokeSnapshots;
    }

    /**
     * 控えた線の設定に戻す
     * @param {Object[]} strokeSnapshots - saveStrokeSnapshots() の戻り値
     * @param {boolean} includeDashes - 破線と線端も戻すなら true（線幅・角の形状・比率だけなら false）
     * @returns {void}
     */
    function restoreStrokeSnapshots(strokeSnapshots, includeDashes) {
        for (var i = 0; i < strokeSnapshots.length; i++) {
            var snapshot = strokeSnapshots[i];
            /* ロック中などで戻せなくても、残りは続ける / keep going when one item refuses */
            try {
                snapshot.item.strokeWidth = snapshot.strokeWidth;
                snapshot.item.strokeJoin = snapshot.strokeJoin;
                snapshot.item.strokeMiterLimit = snapshot.strokeMiterLimit;
                if (includeDashes) {
                    snapshot.item.strokeDashes = snapshot.strokeDashes;
                    snapshot.item.strokeCap = snapshot.strokeCap;
                }
            } catch (e) { $.writeln("restoreStrokeSnapshots: " + e); }
        }
        app.redraw();
    }

    /**
     * パスに、破線（STROKE_DASH_PATTERN）・丸型線端・［コーナーやパス先端に破線の先端を整列］を設定する。
     * 整列は DOM に無いので線の設定のアクションで入れ、アクションが既定値に戻す線幅・角の形状・比率は控えた値に戻す。
     * 破線と線端は DOM でも確かめて入れ直す
     * @param {PathItem[]} strokeItems - collectGraphPaths() の戻り値
     * @returns {boolean} 整列まで設定できたら true
     */
    function applyDashSettings(strokeItems) {
        var strokeSnapshots = saveStrokeSnapshots(strokeItems); /* アクションの前の線幅・角の形状・比率 / weight, join and miter before the action */
        var isAligned = runOnItems(strokeItems, function () {
            return runTemporaryAction(buildDashActionSource(strokeSnapshots[0].strokeWidth), DASH_ACTION_SET_NAME, DASH_ACTION_NAME);
        }) === true;
        restoreStrokeSnapshots(strokeSnapshots, false);
        for (var i = 0; i < strokeItems.length; i++) {
            try {
                if (String(strokeItems[i].strokeDashes) !== String(STROKE_DASH_PATTERN)) strokeItems[i].strokeDashes = STROKE_DASH_PATTERN;
                if (strokeItems[i].strokeCap !== StrokeCap.ROUNDENDCAP) strokeItems[i].strokeCap = StrokeCap.ROUNDENDCAP;
            } catch (e) { $.writeln("applyDashSettings: " + e); }
        }
        app.redraw();
        return isAligned;
    }

    /**
     * 破線・丸型線端・［コーナーやパス先端に破線の先端を整列］を設定する、線の設定（ai_plugin_setStroke）の一時アクションを作る。
     * 省いたパラメーターは既定値に戻るので、すべて書く（線幅・角の形状・比率はあとで DOM で戻す。矢印はなし、線の位置は中央になる）
     * @param {number} strokeWidth - 線幅（pt）
     * @returns {string} アクション定義のテキスト
     */
    function buildDashActionSource(strokeWidth) {
        var actionLines = ["/version 3"]
            .concat(buildActionNameLines("", DASH_ACTION_SET_NAME))
            .concat(["/isOpen 1", "/actionCount 1", "/action-1 {"])
            .concat(buildActionNameLines("\t", DASH_ACTION_NAME))
            .concat(["\t/keyIndex 0", "\t/colorIndex 0", "\t/isOpen 1", "\t/eventCount 1", "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0", "\t\t/internalName (ai_plugin_setStroke)"])
            .concat(buildActionNameLines("\t\t", "線を設定", "localizedName"))
            .concat(["\t\t/isOpen 0", "\t\t/isOn 1", "\t\t/hasDialog 0", "\t\t/parameterCount " + (10 + STROKE_DASH_PATTERN.length)]);
        var paramIndex = 0;

        /**
         * パラメーターを1つ足す
         * @param {number} paramKey - パラメーターキー
         * @param {string[]} bodyLines - /type から後の行（字下げなし）
         * @returns {void}
         */
        function addParam(paramKey, bodyLines) {
            paramIndex++;
            actionLines.push("\t\t/parameter-" + paramIndex + " {", "\t\t\t/key " + paramKey, "\t\t\t/showInPalette 4294967295");
            for (var i = 0; i < bodyLines.length; i++) actionLines.push("\t\t\t" + bodyLines[i]);
            actionLines.push("\t\t}");
        }

        /**
         * 列挙パラメーターの本文を作る
         * @param {string} displayName - 表示名
         * @param {number} enumValue - 値
         * @returns {string[]} 本文の行
         */
        function enumBody(displayName, enumValue) {
            return ["/type (enumerated)"].concat(buildActionNameLines("", displayName), ["/value " + enumValue]);
        }

        addParam(STROKE_KEY_WIDTH, ["/type (unit real)", "/value " + formatActionReal(strokeWidth), "/unit " + ACTION_UNIT_POINT]);
        addParam(STROKE_KEY_CAP, enumBody("丸型線端", 1));
        addParam(STROKE_KEY_MITER, ["/type (real)", "/value 10.0"]);
        addParam(STROKE_KEY_JOIN, enumBody("マイター結合", 0));
        addParam(STROKE_KEY_DASH_COUNT, ["/type (integer)", "/value " + STROKE_DASH_PATTERN.length]);
        for (var d = 0; d < STROKE_DASH_PATTERN.length; d++) {
            addParam(STROKE_KEY_DASH_1 + d, ["/type (unit real)", "/value " + formatActionReal(STROKE_DASH_PATTERN[d]), "/unit " + ACTION_UNIT_POINT]);
        }
        addParam(STROKE_KEY_DASH_ADJUST, ["/type (boolean)", "/value 1"]);
        /* 矢印は「[なし]」 / no arrowheads */
        addParam(STROKE_KEY_ARROW_1, ["/type (ustring)"].concat(buildActionNameLines("", "[なし]", "value")));
        addParam(STROKE_KEY_ARROW_2, ["/type (ustring)"].concat(buildActionNameLines("", "[なし]", "value")));
        addParam(STROKE_KEY_TIP_ALIGN, enumBody("パスの終点に配置", 0));
        addParam(STROKE_KEY_ALIGN, enumBody("中央", 0));
        actionLines.push("\t}", "}", "");
        return actionLines.join(String.fromCharCode(13));
    }

    /**
     * アクション定義の実数の書き方にする（整数でも「.0」を付ける）
     * @param {number} value - 数値
     * @returns {string} 「1.0」「0.5」の形の文字列
     */
    function formatActionReal(value) {
        var realText = String(Math.round(value * 10000) / 10000);
        return (realText.indexOf(".") < 0) ? realText + ".0" : realText;
    }

    /**
     * パスを選択して、楕円のライブシェイプの角度を入れる一時アクション（楕円形のシェイププロパティ）を実行し、選択を戻す。
     * 楕円でないパスには効かない
     * @param {PathItem[]} ellipsePaths - パス
     * @param {number} angle - 角度（度）
     * @returns {void}
     */
    function setEllipseAngle(ellipsePaths, angle) {
        if (ellipsePaths.length === 0) return; /* 選択が空のままアクションを流さない / do not run the action on an empty selection */
        runOnItems(ellipsePaths, function () {
            runTemporaryAction(buildEllipseAngleActionSource(angle), ELLIPSE_ACTION_SET_NAME, ELLIPSE_ACTION_NAME);
        });
    }

    /**
     * 楕円のライブシェイプの角度を BBAccumRotation タグから読む。
     * 角度は DOM に無いが、パネルの角度がこのタグ（ラジアン）に入る。0° の楕円にはタグが付かない
     * @param {PathItem} pathItem - パス
     * @returns {number} 角度（度）。タグが無ければ 0
     */
    function readEllipseAngle(pathItem) {
        var rotationTag = findItemByName(pathItem.tags, "BBAccumRotation");
        return rotationTag ? Number(rotationTag.value) * 180 / Math.PI : 0;
    }

    /**
     * 楕円の角度がまだ指定の角度になっていないパスを集める
     * @param {PathItem[]} graphPaths - パス
     * @param {number} angle - 角度（度）
     * @returns {PathItem[]} 角度を入れるパス
     */
    function collectPathsNotAtAngle(graphPaths, angle) {
        var anglePaths = [];
        for (var i = 0; i < graphPaths.length; i++) {
            if (Math.abs(readEllipseAngle(graphPaths[i]) - angle) > ELLIPSE_ANGLE_TOLERANCE) anglePaths.push(graphPaths[i]);
        }
        return anglePaths;
    }

    /**
     * 楕円のライブシェイプの角度を入れる一時アクション（ai_liveshape_ellipse）を作る
     * @param {number} angle - 角度（度）
     * @returns {string} アクション定義のテキスト
     */
    function buildEllipseAngleActionSource(angle) {
        return ["/version 3"]
            .concat(buildActionNameLines("", ELLIPSE_ACTION_SET_NAME))
            .concat(["/isOpen 1", "/actionCount 1", "/action-1 {"])
            .concat(buildActionNameLines("\t", ELLIPSE_ACTION_NAME))
            .concat(["\t/keyIndex 0", "\t/colorIndex 0", "\t/isOpen 1", "\t/eventCount 1", "\t/event-1 {",
                "\t\t/useRulersIn1stQuadrant 0", "\t\t/internalName (ai_liveshape_ellipse)"])
            .concat(buildActionNameLines("\t\t", "楕円形のシェイププロパティ", "localizedName"))
            .concat(["\t\t/isOpen 1", "\t\t/isOn 1", "\t\t/hasDialog 0", "\t\t/parameterCount 3",
                /* shap / type：角度の項目 / which ellipse property to set (the angle) */
                "\t\t/parameter-1 {", "\t\t\t/key 1936220528", "\t\t\t/showInPalette 4294967295", "\t\t\t/type (enumerated)",
                "\t\t\t/name [ 0", "", "\t\t\t]", "\t\t\t/value 9", "\t\t}",
                "\t\t/parameter-2 {", "\t\t\t/key 1954115685", "\t\t\t/showInPalette 4294967295", "\t\t\t/type (enumerated)",
                "\t\t\t/name [ 0", "", "\t\t\t]", "\t\t\t/value 6", "\t\t}",
                /* angl：角度 / angle */
                "\t\t/parameter-3 {", "\t\t\t/key 1634625388", "\t\t\t/showInPalette 4294967295", "\t\t\t/type (unit real)",
                "\t\t\t/value " + formatActionReal(angle), "\t\t\t/unit 591490663", "\t\t}",
                "\t}", "}", ""])
            .join(String.fromCharCode(13));
    }

    /**
     * パスを選択して［パスの方向反転］を実行し、選択を戻す
     * @param {PathItem[]} targetPaths - パス
     * @returns {void}
     */
    function reversePathDirections(targetPaths) {
        runOnItems(targetPaths, function () {
            app.executeMenuCommand("Reverse Path Direction");
        });
    }

    /**
     * 項目を選択してから処理を実行し、選択を戻して再描画する（アクション・メニューコマンドは選択に掛かるため）
     * @param {PageItem[]} targetItems - 選択する項目
     * @param {Function} selectionOperation - 実行する処理
     * @returns {*} 処理の戻り値（失敗したら undefined）
     */
    function runOnItems(targetItems, selectionOperation) {
        var doc = app.activeDocument;
        var savedSelection = doc.selection;
        var operationResult;
        /* ロック中の項目は選択に代入すると例外になる / assigning locked items to the selection throws */
        try {
            doc.selection = targetItems;
            operationResult = selectionOperation();
        } catch (e) {
            $.writeln("runOnItems: " + e);
        }
        restoreSelection(doc, savedSelection);
        app.redraw();
        return operationResult;
    }

    /**
     * 控えた選択に戻す（戻せなくても続ける）
     * @param {Document} doc - ドキュメント
     * @param {*} savedSelection - 控えた doc.selection
     * @returns {void}
     */
    function restoreSelection(doc, savedSelection) {
        try {
            if (savedSelection instanceof Array) doc.selection = savedSelection;
            else if (savedSelection && savedSelection.select) savedSelection.select();
        } catch (e) { /* 選択を戻せなくても処理は済んでいる / the work is already done */ }
    }

    /**
     * 名前の衝突を避けた一意な名前を作る（"名前", "名前 1", "名前 2" …）
     * @param {string} baseName - 元の名前
     * @param {Object} namedCollection - 名前の重複を調べるコレクション（getByName を持つ）
     * @returns {string} 使われていない名前
     */
    function uniqueName(baseName, namedCollection) {
        var candidateName = baseName;
        for (var suffix = 1; findItemByName(namedCollection, candidateName); suffix++) {
            candidateName = baseName + " " + suffix;
        }
        return candidateName;
    }

    /**
     * コレクションから名前で項目を探す（グラデーション・グラフィックスタイル・レイヤー・タグ）
     * @param {Object} namedCollection - getByName を持つコレクション
     * @param {string} itemName - 探す名前
     * @returns {Object|null} 見つかった項目。無ければ null
     */
    function findItemByName(namedCollection, itemName) {
        /* getByName は見つからないと例外 / getByName throws when the name is missing */
        try {
            return namedCollection.getByName(itemName);
        } catch (e) {
            return null;
        }
    }

    // =========================================
    // ［パスに沿って…］の線のスタイル / Along-stroke gradient style
    // =========================================

    /**
     * 線にグラデーションが無いパスがあるかを返す
     * @param {PathItem[]} graphPaths - パス
     * @returns {boolean} 1つでもあれば true
     */
    function hasPathWithoutGradientStroke(graphPaths) {
        for (var i = 0; i < graphPaths.length; i++) {
            if (!hasGradientStroke(graphPaths[i])) return true;
        }
        return false;
    }

    /**
     * ［パスに沿ってグラデーションを適用］の線のグラフィックスタイルを返す。
     * ドキュメントにあればそれを使い（埋め込んだファイルは開かない）、無ければ埋め込んだ AI ファイルから取り込む
     * @param {Document} doc - 対象ドキュメント
     * @returns {GraphicStyle|null} グラフィックスタイル。用意できなければ警告して null
     */
    function getStrokeGradientStyle(doc) {
        var strokeGradientStyle = findItemByName(doc.graphicStyles, STROKE_GRADIENT_STYLE_NAME);
        if (strokeGradientStyle) return strokeGradientStyle;
        if (!importEmbeddedStyles(doc)) return null; /* importEmbeddedStyles が警告済み / already warned */
        strokeGradientStyle = findItemByName(doc.graphicStyles, STROKE_GRADIENT_STYLE_NAME);
        if (!strokeGradientStyle) alert(getLabel("alert.styleNotFound", { styleName: STROKE_GRADIENT_STYLE_NAME }));
        return strokeGradientStyle;
    }

    /**
     * グラフィックスタイルを適用し、スタイルで変わる不透明度・描画モードを元に戻す（塗りはスタイルどおりなし）。
     * 適用できなくても例外にしない（呼び出し元は線がグラデーションになったかで判断する）
     * @param {GraphicStyle} graphicStyle - 適用するグラフィックスタイル
     * @param {PathItem} pathItem - パス
     * @returns {void}
     */
    function applyStyleKeepingOpacity(graphicStyle, pathItem) {
        var originalOpacity = pathItem.opacity;
        var originalBlendingMode = pathItem.blendingMode;
        try {
            graphicStyle.applyTo(pathItem);
            pathItem.opacity = originalOpacity;
            pathItem.blendingMode = originalBlendingMode;
        } catch (e) {
            $.writeln("applyStyleKeepingOpacity: " + e);
        }
    }

    // =========================================
    // 埋め込みファイルの復元 / Restore the embedded file
    // =========================================

    /**
     * Base64 の行の配列をバイト列の文字列（1文字が1バイト）に戻す。
     * 1文字ずつ += でつなぐと遅いので、行ごとに文字コードを配列に溜めて fromCharCode.apply でまとめる
     * @param {string[]} base64Lines - Base64 の行（どの行も4文字単位で、行ごとにデコードできる）
     * @returns {string} バイト列の文字列
     */
    function decodeBase64Lines(base64Lines) {
        var base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        /* 文字コード → 6ビットの値。indexOf を毎回呼ぶより速い / char code -> 6-bit value, faster than indexOf */
        var codeTable = [];
        for (var i = 0; i < base64Chars.length; i++) codeTable[base64Chars.charCodeAt(i)] = i;
        var paddingCode = "=".charCodeAt(0);

        var byteParts = [];
        for (var lineIndex = 0; lineIndex < base64Lines.length; lineIndex++) {
            var base64Line = base64Lines[lineIndex];
            var byteCodes = [];
            for (var j = 0; j + 3 < base64Line.length; j += 4) {
                var thirdCode = base64Line.charCodeAt(j + 2);
                var fourthCode = base64Line.charCodeAt(j + 3);
                var groupBits = (codeTable[base64Line.charCodeAt(j)] << 18) |
                    (codeTable[base64Line.charCodeAt(j + 1)] << 12) |
                    ((thirdCode === paddingCode ? 0 : codeTable[thirdCode]) << 6) |
                    (fourthCode === paddingCode ? 0 : codeTable[fourthCode]);
                byteCodes.push((groupBits >> 16) & 255);
                if (thirdCode !== paddingCode) byteCodes.push((groupBits >> 8) & 255);
                if (fourthCode !== paddingCode) byteCodes.push(groupBits & 255);
            }
            byteParts.push(String.fromCharCode.apply(null, byteCodes));
        }
        return byteParts.join("");
    }

    /**
     * 埋め込んだ AI ファイルを一時フォルダーに書き出す。同じ中身を書き出し済みならデコードせずにそれを使う
     * @returns {File|null} 書き出したファイル。書き出せなければ null
     */
    function restoreEmbeddedFile() {
        /* 中身が変わると ID も変わるので、ID ごとのフォルダーにして古いファイルと取り違えない
           The ID changes with the content, so a folder per ID never picks up an older file */
        var restoreFolder = new Folder(Folder.temp.fsName + "/" + SCRIPT_NAME + "/" + getEmbeddedFileId());
        var restoredFile = new File(restoreFolder.fsName + "/" + getEmbeddedFileName());
        if (restoredFile.exists && restoredFile.length === getEmbeddedFileSize()) return restoredFile;

        if (!restoreFolder.parent.exists) restoreFolder.parent.create();
        if (!restoreFolder.exists && !restoreFolder.create()) return null;
        restoredFile.encoding = "BINARY";
        if (!restoredFile.open("w")) return null;
        restoredFile.write(decodeBase64Lines(getEmbeddedFileBase64Lines()));
        restoredFile.close();
        /* 書き出しが途中で失敗していたら次回また書き直せるよう、壊れたファイルは残さない
           Remove a short file so the next run writes it again */
        if (restoredFile.length !== getEmbeddedFileSize()) {
            restoredFile.remove();
            return null;
        }
        return restoredFile;
    }

    // =========================================
    // スタイルの取り込み / Import styles
    // =========================================

    /**
     * 埋め込んだ AI ファイルを復元してグラフィックスタイルを取り込む。
     * ファイルを開いてアートボード内をコピーし、取り込み先に貼り付けてすぐ消す（グラフィックスタイルだけが残る）
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @returns {boolean} 取り込めたら true。復元・オープンに失敗したら警告して false
     */
    function importEmbeddedStyles(destinationDoc) {
        var styleFile = restoreEmbeddedFile();
        if (!styleFile) {
            alert(getLabel("alert.restoreFailed", { fileName: getEmbeddedFileName() }));
            return false;
        }
        /* 復元したファイルが開けないときは Illustrator の例外になる / app.open throws when the file cannot be opened */
        try {
            copyArtboardFromStyleFile(styleFile);
        } catch (e) {
            app.activeDocument = destinationDoc;
            alert(getLabel("alert.openFailed", { fileName: getEmbeddedFileName() }) + "\n" + e);
            return false;
        }
        app.activeDocument = destinationDoc;
        pasteAndDiscard(destinationDoc);
        return true;
    }

    /**
     * スタイルの AI ファイルを開き、作業アートボード内をコピーする。
     * スクリプトが開いたときだけ保存せずに閉じる（すでに開いていたものは閉じない）
     * @param {File} styleFile - スタイルの AI ファイル
     * @returns {void}
     */
    function copyArtboardFromStyleFile(styleFile) {
        /* 開いているファイルを app.open しても手前に出るだけで、ドキュメントの数は変わらない
           Opening a file that is already open only brings it to the front, so the document count stays the same */
        var documentCountBefore = app.documents.length;
        var styleSourceDoc = app.open(styleFile);
        var wasAlreadyOpen = (app.documents.length === documentCountBefore);

        var selectionBeforeCopy = styleSourceDoc.selection;
        app.executeMenuCommand("selectallinartboard");
        app.executeMenuCommand("copy");
        if (wasAlreadyOpen) restoreSelection(styleSourceDoc, selectionBeforeCopy);
        else styleSourceDoc.close(SaveOptions.DONOTSAVECHANGES);
    }

    /**
     * 取り込み用の一時レイヤーを返す（無ければ作る）。ロックと非表示は外す
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @returns {Layer} 一時レイヤー
     */
    function getOrCreateImportLayer(destinationDoc) {
        var importLayer = findItemByName(destinationDoc.layers, IMPORT_LAYER_NAME);
        if (!importLayer) {
            importLayer = destinationDoc.layers.add();
            importLayer.name = IMPORT_LAYER_NAME;
        }
        importLayer.locked = false;
        importLayer.visible = true;
        return importLayer;
    }

    /**
     * 配列に同じ DOM の参照が入っているか
     * @param {Array} referenceList - 調べる配列
     * @param {Object} targetReference - 探す参照
     * @returns {boolean} 入っていれば true
     */
    function containsReference(referenceList, targetReference) {
        for (var i = 0; i < referenceList.length; i++) {
            if (referenceList[i] === targetReference) return true;
        }
        return false;
    }

    /**
     * レイヤーとサブレイヤーをすべて集める（親が先、子が後）
     * @param {Document|Layer} layerContainer - 集めるレイヤーの親
     * @param {Layer[]} layerList - 追加先の配列
     * @returns {Layer[]} layerList
     */
    function collectLayersDeep(layerContainer, layerList) {
        for (var i = 0; i < layerContainer.layers.length; i++) {
            layerList.push(layerContainer.layers[i]);
            collectLayersDeep(layerContainer.layers[i], layerList);
        }
        return layerList;
    }

    /**
     * クリップボードの中身を取り込み先に貼り付け、すぐに消す（グラフィックスタイルだけがドキュメントに残る）。
     * ［コピー元のレイヤーにペースト］がオンだと一時レイヤー以外にも入るので、一時レイヤーの外に貼り付いたものと、
     * 貼り付けで増えたレイヤー・サブレイヤーも消す。最後に貼り付ける前の選択に戻す
     * @param {Document} destinationDoc - 取り込み先のドキュメント
     * @returns {void}
     */
    function pasteAndDiscard(destinationDoc) {
        var selectionBeforePaste = destinationDoc.selection;
        var activeLayerBeforePaste = destinationDoc.activeLayer;
        var layersBeforePaste = collectLayersDeep(destinationDoc, []);

        var importLayer = getOrCreateImportLayer(destinationDoc);
        destinationDoc.activeLayer = importLayer;
        /* 選択を残したまま貼り付けると、選択中のものや文字が置き換わることがある / Pasting over a selection can replace the selected items or text */
        destinationDoc.selection = null;
        app.executeMenuCommand("paste");

        /* 一時レイヤーのものはレイヤーごと消えるので、外に貼り付いたものだけ消す / Items on the temporary layer go with it; remove only the ones pasted elsewhere */
        var pastedItems = destinationDoc.selection;
        if (pastedItems && pastedItems.typename !== "TextRange") {
            for (var i = pastedItems.length - 1; i >= 0; i--) {
                if (pastedItems[i].layer !== importLayer) pastedItems[i].remove();
            }
        }

        /* 子から先に消す。最後の1枚のレイヤーは消せない / Remove children first; the last top-level layer cannot be removed */
        var layersAfterPaste = collectLayersDeep(destinationDoc, []);
        for (var j = layersAfterPaste.length - 1; j >= 0; j--) {
            var candidateLayer = layersAfterPaste[j];
            if (candidateLayer !== importLayer && containsReference(layersBeforePaste, candidateLayer)) continue;
            if (candidateLayer.parent === destinationDoc && destinationDoc.layers.length <= 1) continue;
            candidateLayer.locked = false;
            candidateLayer.remove();
        }

        /* 作業中のレイヤーと選択を戻す（このあとの破線・方向反転は選択に掛ける）
           Restore the active layer and the selection; the dash and reverse steps work on the selection */
        try { destinationDoc.activeLayer = activeLayerBeforePaste; } catch (e) { /* 消えたレイヤーなら戻さない / the layer is gone */ }
        destinationDoc.selection = null;
        restoreSelection(destinationDoc, selectionBeforePaste);
    }

    // =========================================
    // カラーの表示 / Color display
    // =========================================

    /**
     * 数値を小数第1位で丸めた文字列にする（.0 は付けない）
     * @param {number} value - 数値
     * @returns {string} 表示用の文字列
     */
    function formatNumber(value) {
        return String(Math.round(value * 10) / 10);
    }

    /**
     * カラーの値を文字列にする（例「C0 M50 Y100 K0」「R255 G0 B0」「K50」、スウォッチは名前と濃度）
     * @param {Color} color - カラー
     * @returns {string} 表示用の文字列
     */
    function describeColor(color) {
        try {
            switch (color.typename) {
                case "CMYKColor":
                    return "C" + formatNumber(color.cyan) + " M" + formatNumber(color.magenta) +
                        " Y" + formatNumber(color.yellow) + " K" + formatNumber(color.black);
                case "RGBColor":
                    return "R" + formatNumber(color.red) + " G" + formatNumber(color.green) + " B" + formatNumber(color.blue);
                case "GrayColor":
                    return "K" + formatNumber(color.gray);
                case "LabColor":
                    return "L" + formatNumber(color.l) + " a" + formatNumber(color.a) + " b" + formatNumber(color.b);
                case "SpotColor":
                    /* グローバルプロセスカラーも SpotColor で返る / global process colors also come back as SpotColor */
                    return color.spot.name + (color.tint < 100 ? " " + formatNumber(color.tint) + "%" : "");
                case "NoColor":
                    return getLabel("colorText.none");
            }
        } catch (e) { /* 値を読めない / values are unreadable */ }
        return getLabel("colorText.unknown");
    }

    /**
     * カラーを画面表示用の RGB（0〜255）にする
     * @param {Color} color - カラー
     * @returns {number[]|null} [R, G, B]（変換できなければ null）
     */
    function toScreenRgb(color) {
        try {
            switch (color.typename) {
                case "RGBColor":
                    return [color.red, color.green, color.blue];
                case "CMYKColor":
                    return app.convertSampleColor(ImageColorSpace.CMYK, [color.cyan, color.magenta, color.yellow, color.black],
                        ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
                case "GrayColor":
                    return app.convertSampleColor(ImageColorSpace.GrayScale, [color.gray],
                        ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
                case "LabColor":
                    return app.convertSampleColor(ImageColorSpace.LAB, [color.l, color.a, color.b],
                        ImageColorSpace.RGB, ColorConvertPurpose.defaultpurpose);
                case "SpotColor":
                    var baseRgb = toScreenRgb(color.spot.color);
                    if (!baseRgb) return null;
                    /* 濃度の分だけ白に寄せる / lighten toward white by the tint */
                    var tintRatio = color.tint / 100;
                    return [
                        255 - (255 - baseRgb[0]) * tintRatio,
                        255 - (255 - baseRgb[1]) * tintRatio,
                        255 - (255 - baseRgb[2]) * tintRatio
                    ];
            }
        } catch (e) { /* 変換できない / cannot convert */ }
        return null;
    }

    /**
     * 標準のカラーピッカーでカラーを選び直す
     * @param {Color} currentColor - 今のカラー
     * @returns {Color|null} 選んだカラー（キャンセル・同じ色・ピッカーが受け付けない色なら null）
     */
    function pickColor(currentColor) {
        var pickedColor;
        /* スウォッチのカラーなど、ピッカーが受け付けない色では例外になることがある / the picker may refuse some colors, such as swatch colors */
        try {
            pickedColor = app.showColorPicker(currentColor);
        } catch (e) {
            return null;
        }
        /* キャンセルでは渡した色がそのまま返る / Cancel returns the color that was passed in */
        if (!pickedColor || describeColor(pickedColor) === describeColor(currentColor)) return null;
        return pickedColor;
    }

    /**
     * 見本とカラーの文字を、カラーに合わせて描き直す
     * @param {Group} colorChip - addColorChip() の見本
     * @param {StaticText} colorText - カラーの値の文字
     * @param {Color} color - カラー
     * @returns {void}
     */
    function showChipColor(colorChip, colorText, color) {
        colorText.text = describeColor(color);
        colorChip.rgbValues = toScreenRgb(color);
        /* group には notify() が無いので、隠して再表示して描き直す / groups have no notify(), so hide and show to redraw */
        colorChip.hide();
        colorChip.show();
    }

    /**
     * カラーの見本を足す（変換できない色は斜線）。色は .rgbValues を書き換えて描き直す
     * @param {Group} chipRow - 追加先の行
     * @param {number[]|null} rgbValues - [R, G, B]（0〜255）
     * @returns {Group} 見本
     */
    function addColorChip(chipRow, rgbValues) {
        var colorChip = chipRow.add("group");
        colorChip.preferredSize = [COLOR_CHIP_SIZE, COLOR_CHIP_SIZE];
        colorChip.minimumSize = [COLOR_CHIP_SIZE, COLOR_CHIP_SIZE];
        colorChip.maximumSize = [COLOR_CHIP_SIZE, COLOR_CHIP_SIZE];
        colorChip.rgbValues = rgbValues;
        var frameColor = isDarkUI() ? [1, 1, 1, 0.35] : [0, 0, 0, 0.35];
        colorChip.onDraw = function () {
            var chipGraphics = colorChip.graphics;
            var chipWidth = colorChip.size.width;
            var chipHeight = colorChip.size.height;
            var chipRgb = colorChip.rgbValues;
            var framePen = chipGraphics.newPen(chipGraphics.PenType.SOLID_COLOR, frameColor, 1);
            if (chipRgb) {
                chipGraphics.newPath();
                chipGraphics.rectPath(0, 0, chipWidth, chipHeight);
                chipGraphics.fillPath(chipGraphics.newBrush(chipGraphics.BrushType.SOLID_COLOR,
                    [chipRgb[0] / 255, chipRgb[1] / 255, chipRgb[2] / 255, 1]));
            } else {
                chipGraphics.newPath();
                chipGraphics.moveTo(chipWidth, 0);
                chipGraphics.lineTo(0, chipHeight);
                chipGraphics.strokePath(framePen);
            }
            chipGraphics.newPath();
            chipGraphics.rectPath(0.5, 0.5, chipWidth - 1, chipHeight - 1);
            chipGraphics.strokePath(framePen);
        };
        return colorChip;
    }

    // =========================================
    // ダイアログUI / Dialog UI
    // =========================================

    /**
     * ［線のグラフ］のダイアログを表示する（ベース・アクセントのカラー、アクセントの割合、線幅）。入力に合わせて onPreview で反映する
     * @param {Object} dialogSetup - 表示の設定
     * @param {{label: string, color: Color}[]} dialogSetup.colorChoices - カラーの行（見本のクリックで選び直すと .color を書き換える）
     * @param {number} dialogSetup.initialAccent - アクセントの割合の初期値
     * @param {number} dialogSetup.initialStrokeWidth - 線幅の初期値（pt）
     * @param {boolean} dialogSetup.initialReverse - ［方向反転］の初期値
     * @param {function(Object): void} dialogSetup.onPreview - 使える値が入力されるたびに {accentPercent, strokeWidth} で呼ぶ
     * @param {function(boolean): void} dialogSetup.onReverseChange - ［方向反転］を切り替えるたびにチェックの状態で呼ぶ
     * @returns {{accentPercent: number, strokeWidth: number}|null} アクセントの割合と線幅（pt）。キャンセルなら null
     */
    function showStrokeGraphDialog(dialogSetup) {
        var graphDialog = new Window("dialog", getLabel("dialog.title"));
        setupWindow(graphDialog);

        var accentInput = addColorPanel(graphDialog, dialogSetup, refreshPreview);
        /* 線幅は環境設定の線の単位で表示する / show the weight in the stroke unit preference */
        var strokeUnit = getUnitInfo("strokeUnits");
        var strokeWidthInput = addStrokePanel(graphDialog, dialogSetup, strokeUnit, refreshPreview);

        var buttonRow = addButtonRow(graphDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var okButton = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * 入力欄の値を読む（使えない値なら null）
         * @returns {{accentPercent: number, strokeWidth: number}|null} アクセントの割合と線幅（pt）
         */
        function readGraphValues() {
            var accentPercent = parseFloat(accentInput.text);
            var strokeWidth = unitToPt(parseFloat(strokeWidthInput.text), strokeUnit);
            if (!areBoundariesValid([100 - accentPercent])) return null;
            if (isNaN(strokeWidth) || strokeWidth <= 0) return null;
            return { accentPercent: accentPercent, strokeWidth: strokeWidth };
        }

        /**
         * 入力した値をプレビューに反映する（使えない値のあいだは前の表示のまま）
         * @returns {void}
         */
        function refreshPreview() {
            var graphValues = readGraphValues();
            if (graphValues) dialogSetup.onPreview(graphValues);
        }

        okButton.onClick = function () {
            if (!readGraphValues()) {
                alert(getLabel("alert.invalidValues"));
                return;
            }
            graphDialog.close(1);
        };

        refreshPreview();
        alignRightOnlyButtonRow(buttonRow);
        addVersionFooter(graphDialog);
        prepareDialogWindow(graphDialog, SCRIPT_NAME);
        if (graphDialog.show() !== 1) return null;
        return readGraphValues();
    }

    /**
     * ［カラー］パネル（ベース・アクセントの行とアクセントの割合）を足す
     * @param {Window} graphDialog - ダイアログボックス
     * @param {Object} dialogSetup - showStrokeGraphDialog() の表示の設定
     * @param {Function} onInput - 値やカラーが変わったときに呼ぶ関数
     * @returns {EditText} アクセントの割合の入力欄
     */
    function addColorPanel(graphDialog, dialogSetup, onInput) {
        var colorPanel = graphDialog.add("panel", undefined, getLabel("panel.color"));
        setupPanel(colorPanel, COLOR_PANEL_SPACING);
        for (var i = 0; i < dialogSetup.colorChoices.length; i++) {
            addColorChoiceRow(colorPanel, dialogSetup.colorChoices[i], FIELD_LABEL_WIDTH, onInput);
        }
        var accentInput = addSteppedField(colorPanel, {
            label: labelText("fieldLabel.accentPercent"),
            labelWidth: FIELD_LABEL_WIDTH,
            text: formatStepperNumber(dialogSetup.initialAccent) + ACCENT_FIELD_OPTIONS.unit,
            characters: FIELD_CHARACTERS,
            step: ACCENT_FIELD_OPTIONS.step,
            min: ACCENT_FIELD_OPTIONS.min,
            max: ACCENT_FIELD_OPTIONS.max,
            unit: ACCENT_FIELD_OPTIONS.unit,
            onStep: onInput
        });
        accentInput.helpTip = getLabel("tooltip.accentPercent");
        appendOnChange(accentInput, onInput);
        return accentInput;
    }

    /**
     * ［線］パネル（線幅と［方向反転］）を足す
     * @param {Window} graphDialog - ダイアログボックス
     * @param {Object} dialogSetup - showStrokeGraphDialog() の表示の設定
     * @param {{label: string, pointsPerUnit: number}} strokeUnit - 線の単位（getUnitInfo("strokeUnits")）
     * @param {Function} onInput - 線幅が変わったときに呼ぶ関数
     * @returns {EditText} 線幅の入力欄
     */
    function addStrokePanel(graphDialog, dialogSetup, strokeUnit, onInput) {
        var strokePanel = graphDialog.add("panel", undefined, getLabel("panel.stroke"));
        setupPanel(strokePanel, 6);
        var strokeWidthInput = addSteppedField(strokePanel, {
            label: labelText("fieldLabel.strokeWidth"),
            labelWidth: FIELD_LABEL_WIDTH,
            text: formatStepperNumber(ptToUnit(dialogSetup.initialStrokeWidth, strokeUnit)) + " " + strokeUnit.label,
            characters: FIELD_CHARACTERS,
            step: WEIGHT_FIELD_OPTIONS.step,
            min: WEIGHT_FIELD_OPTIONS.min,
            max: WEIGHT_FIELD_OPTIONS.max,
            unit: " " + strokeUnit.label,
            onStep: onInput
        });
        appendOnChange(strokeWidthInput, onInput);

        /* 方向反転（切り替えるとその場で反転する）。左端を線幅の入力欄にそろえる（項目名・間隔・∧∨の幅だけ空ける）
           Reverse direction, applied right away; its left edge lines up with the weight field */
        var reverseRow = strokePanel.add("group");
        setupRow(reverseRow, "left", 0);
        reverseRow.margins = [FIELD_LABEL_WIDTH + STEPPER_FIELD_SPACING + STEPPER_SIDE_MARGIN + STEPPER_BUTTON_WIDTH, REVERSE_ROW_TOP_GAP, 0, 0];
        var reverseCheckbox = reverseRow.add("checkbox", undefined, getLabel("checkbox.reverse"));
        reverseCheckbox.value = dialogSetup.initialReverse;
        reverseCheckbox.helpTip = getLabel("tooltip.reverse");
        reverseCheckbox.onClick = function () {
            dialogSetup.onReverseChange(reverseCheckbox.value);
        };
        return strokeWidthInput;
    }

    /**
     * カラーの行（項目名・見本・カラーの値）を足す。見本のクリックでカラーを選び直す
     * @param {Panel} colorPanel - 追加先のパネル
     * @param {{label: string, color: Color}} colorChoice - 項目名（コロン込み）とカラー（選び直すと .color を書き換える）
     * @param {number} labelWidth - 項目名の幅
     * @param {Function} onColorChanged - カラーを選び直したときに呼ぶ関数
     * @returns {void}
     */
    function addColorChoiceRow(colorPanel, colorChoice, labelWidth, onColorChanged) {
        var choiceRow = colorPanel.add("group");
        setupRow(choiceRow, "left", COLOR_LABEL_CHIP_GAP); /* コロンのあとに半角ほど空ける / about a half-width space after the colon */
        var choiceLabel = choiceRow.add("statictext", undefined, colorChoice.label);
        choiceLabel.preferredSize.width = labelWidth;
        choiceLabel.justify = "right";
        /* 見本と値も少し離す / keep a gap between the chip and its value */
        var chipGroup = choiceRow.add("group");
        setupRow(chipGroup, "left", COLOR_CHIP_TEXT_GAP);
        var colorChip = addColorChip(chipGroup, toScreenRgb(colorChoice.color));
        colorChip.helpTip = getLabel("tooltip.colorChip");
        var colorText = chipGroup.add("statictext", undefined, describeColor(colorChoice.color));
        colorText.preferredSize.width = COLOR_TEXT_WIDTH;
        colorChip.addEventListener("click", function () {
            var pickedColor = pickColor(colorChoice.color);
            if (!pickedColor) return;
            colorChoice.color = pickedColor;
            showChipColor(colorChip, colorText, pickedColor);
            onColorChanged();
        });
    }

    /**
     * 入力欄の確定時の処理に、あとから処理を足す（StepperButtons の onChange で値をそろえたあとに呼ぶ）
     * @param {EditText} numberInput - addSteppedField() で作った入力欄
     * @param {Function} afterChange - 足す処理
     * @returns {void}
     */
    function appendOnChange(numberInput, afterChange) {
        var baseOnChange = numberInput.onChange;
        numberInput.onChange = function () {
            if (baseOnChange) baseOnChange();
            afterChange();
        };
    }

    // =========================================
    // 埋め込みファイル / Embedded file
    // =========================================
    /* 中身が大きいので最後に置く（関数宣言なので、上のメイン処理からも呼べる）
       Placed last because it is large; function declarations are hoisted, so the main code above can call them */

    /* EMBEDDED-FILE-BEGIN（helpers/embed-file-base64.py が書き直す。手で編集しない / rewritten by helpers/embed-file-base64.py; do not edit） */

    /**
     * 埋め込んだファイルの元のファイル名を返す
     * @returns {string} ファイル名
     */
    function getEmbeddedFileName() {
        return "strokegradient.ai";
    }

    /**
     * 埋め込んだファイルのバイト数を返す
     * @returns {number} バイト数
     */
    function getEmbeddedFileSize() {
        return 113006;
    }

    /**
     * 埋め込んだファイルの中身の ID（SHA-1 の先頭8桁）を返す。中身が変わると変わる
     * @returns {string} ID
     */
    function getEmbeddedFileId() {
        return "0a62489b";
    }

    /**
     * 埋め込んだファイルの Base64 を行の配列で返す（1行 3072 バイト分。行ごとにデコードできる）
     * @returns {string[]} Base64 の行
     */
    function getEmbeddedFileBase64Lines() {
        return [
            "JVBERi0xLjYNJeLjz9MNCjEgMCBvYmoNPDwvTWV0YWRhdGEgMiAwIFIvUGFnZXMgMyAwIFIvVHlwZS9DYXRhbG9nPj4NZW5kb2JqDTIgMCBvYmoNPDwvTGVuZ3RoIDIyMzc0L1N1YnR5cGUvWE1ML1R5cGUvTWV0YWRhdGE+PnN0cmVhbQ0KPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgMTAuMC1jMDAwIDEuMDAwMDAwLCAwMDAwLzAwLzAwLTAwOjAwOjAwICAgICAgICAiPgogICA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPgogICAgICA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIgogICAgICAgICAgICB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iCiAgICAgICAgICAgIHhtbG5zOnhtcD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wLyIKICAgICAgICAgICAgeG1sbnM6eG1wR0ltZz0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL2cvaW1nLyIKICAgICAgICAgICAgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iCiAgICAgICAgICAgIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIgogICAgICAgICAgICB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIgogICAgICAgICAgICB4bWxuczppbGx1c3RyYXRvcj0iaHR0cDovL25zLmFkb2JlLmNvbS9pbGx1c3RyYXRvci8xLjAvIgogICAgICAgICAgICB4bWxuczp4bXBUUGc9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC90L3BnLyIKICAgICAgICAgICAgeG1sbnM6c3REaW09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9zVHlwZS9EaW1lbnNpb25zIyIKICAgICAgICAgICAgeG1sbnM6eG1wRz0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL2cvIgogICAgICAgICAgICB4bWxuczpwZGY9Imh0dHA6Ly9ucy5hZG9iZS5jb20vcGRmLzEuMy8iPgogICAgICAgICA8ZGM6Zm9ybWF0PmFwcGxpY2F0aW9uL3ZuZC5hZG9iZS5pbGx1c3RyYXRvcjwvZGM6Zm9ybWF0PgogICAgICAgICA8ZGM6dGl0bGU+CiAgICAgICAgICAgIDxyZGY6QWx0PgogICAgICAgICAgICAgICA8cmRmOmxpIHhtbDpsYW5nPSJ4LWRlZmF1bHQiPnN0cm9rZWdyYWRpZW50PC9yZGY6bGk+CiAgICAgICAgICAgIDwvcmRmOkFsdD4KICAgICAgICAgPC9kYzp0aXRsZT4KICAgICAgICAgPHhtcDpNZXRhZGF0YURhdGU+MjAyNi0xMS0yMVQxNDo1MDoxNiswOTowMDwveG1wOk1ldGFkYXRhRGF0ZT4KICAgICAgICAgPHhtcDpNb2RpZnlEYXRlPjIwMjYtMTEtMjFUMTQ6NTA6MTYrMDk6MDA8L3htcDpNb2RpZnlEYXRlPgogICAgICAgICA8eG1wOkNyZWF0ZURhdGU+MjAyNi0xMS0yMVQxNDo1MDoxNiswOTowMDwveG1wOkNyZWF0ZURhdGU+CiAgICAgICAgIDx4bXA6Q3JlYXRvclRvb2w+QWRvYmUgSWxsdXN0cmF0b3IgMzEuMCAoTWFjaW50b3NoKTwveG1wOkNyZWF0b3JUb29sPgogICAgICAgICA8eG1wOlRodW1ibmFpbHM+CiAgICAgICAgICAgIDxyZGY6QWx0PgogICAgICAgICAgICAgICA8cmRmOmxpIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgICAgICAgPHhtcEdJbWc6d2lkdGg+MjU2PC94bXBHSW1nOndpZHRoPgogICAgICAgICAgICAgICAgICA8eG1wR0ltZzpoZWlnaHQ+MjU2PC94bXBHSW1nOmhlaWdodD4KICAgICAgICAgICAgICAgICAgPHhtcEdJbWc6Zm9ybWF0PkpQRUc8L3htcEdJbWc6Zm9ybWF0PgogICAgICAgICAgICAgICAgICA8eG1wR0ltZzppbWFnZT4vOWovNEFBUVNrWkpSZ0FCQWdFQUFBQUFBQUQvN1FBc1VHaHZkRzl6YUc5d0lETXVNQUE0UWtsTkErMEFBQUFBQUJBQUFBQUFBQUVBJiN4QTtBUUFBQUFBQUFRQUIvKzRBRGtGa2IySmxBR1RBQUFBQUFmL2JBSVFBQmdRRUJBVUVCZ1VGQmdrR0JRWUpDd2dHQmdnTERBb0tDd29LJiN4QTtEQkFNREF3TURBd1FEQTRQRUE4T0RCTVRGQlFURXh3Ykd4c2NIeDhmSHg4Zkh4OGZId0VIQndjTkRBMFlFQkFZR2hVUkZSb2ZIeDhmJiN4QTtIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGZIeDhmSHg4Zkh4OGYvOEFBRVFnQkFBRUFBd0VSJiN4QTtBQUlSQVFNUkFmL0VBYUlBQUFBSEFRRUJBUUVBQUFBQUFBQUFBQVFGQXdJR0FRQUhDQWtLQ3dFQUFnSURBUUVCQVFFQUFBQUFBQUFBJiN4QTtBUUFDQXdRRkJnY0lDUW9MRUFBQ0FRTURBZ1FDQmdjREJBSUdBbk1CQWdNUkJBQUZJUkl4UVZFR0UyRWljWUVVTXBHaEJ4V3hRaVBCJiN4QTtVdEhoTXhaaThDUnlndkVsUXpSVGtxS3lZM1BDTlVRbms2T3pOaGRVWkhURDB1SUlKb01KQ2hnWmhKUkZScVMwVnROVktCcnk0L1BFJiN4QTsxT1QwWlhXRmxhVzF4ZFhsOVdaMmhwYW10c2JXNXZZM1IxZG5kNGVYcDdmSDErZjNPRWhZYUhpSW1LaTR5TmpvK0NrNVNWbHBlWW1aJiN4QTtxYm5KMmVuNUtqcEtXbXA2aXBxcXVzcmE2dm9SQUFJQ0FRSURCUVVFQlFZRUNBTURiUUVBQWhFREJDRVNNVUVGVVJOaElnWnhnWkV5JiN4QTtvYkh3Rk1IUjRTTkNGVkppY3ZFekpEUkRnaGFTVXlXaVk3TENCM1BTTmVKRWd4ZFVrd2dKQ2hnWkpqWkZHaWRrZEZVMzhxT3p3eWdwJiN4QTswK1B6aEpTa3RNVFU1UFJsZFlXVnBiWEYxZVgxUmxabWRvYVdwcmJHMXViMlIxZG5kNGVYcDdmSDErZjNPRWhZYUhpSW1LaTR5TmpvJiN4QTsrRGxKV1dsNWlabXB1Y25aNmZrcU9rcGFhbnFLbXFxNnl0cnErdi9hQUF3REFRQUNFUU1SQUQ4QTlVNHE3RlhZcTdGWFlxN0ZYWXE3JiN4QTtGWFlxN0ZVUGVhallXVWZxWGR4SEF0Q1FaR0MxcDRBN242TXJ5Wm9RRnlJRE9HS1V6VVJiR3RSL01yUmJma3RwSEplT09qQWVuR2ZwJiN4QTtiNHYrRnpXWnUyY1VmcEJsOWcvSHdkaGk3THlTK3Fvc2R2UHpOMXlXb3Q0WWJkZXhvWGNmU3g0LzhMbXZ5ZHM1VHlBRG13N0t4am1TJiN4QTtVbXVmTi9tYTROWk5SbVgvQUl4a1JmOEFKc0xtSlBYNXBjNUg3dnVjdUdp",
            "d3g1UkgzL2Vsc3VvWDhwL2UzTXNuK3M3TitzNVFja2p6JiN4QTtKYnhqaU9RQ0ZPUWJGcHdwVm85UXY0VFdLNWxqUGlyc3Y2ams0NUpEa1N3T09KNWdJKzI4NSthYlkxajFLWnYrTXBFdi9Kd1BtVERYJiN4QTtabzhwSDcvdmFaYUhETG5FZmQ5eWMyWDVyYTlDUUxxR0c2VHVhR04vdlU4ZitGekx4OXI1UnpBTGk1T3lNWjVFaGtlbmZtcm9OeFJiJiN4QTsyS1d5YzlXSTlXTWZTdnhmOExtZmk3V3h5K29HUDIvajVPQmw3SXlENlNKZlorUG15dXcxWFRkUWk5V3h1WTdoTy9wc0NSOHgxSDA1JiN4QTtzY2VhRXhjVGJyc21HY0RVZ1Fpc3NhM1lxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcWxtc2VaTkkwbFQ5JiN4QTthbXJOMUZ2SFJwRC9BTEh0OU5NeE5UcmNlSDZqdjNkWEl3YVdlVDZSdDNzRTFqOHhOV3VpMGRpQlpRZEF3K0tVai9XT3crZ2ZUbWgxJiN4QTtIYkdTZTBQU1B0ZHhnN01oSGVYcVAyTVhrZTV1cCtUczg4OGhwVWt1N0g2YWs1cXlaU08rNWRnQUlqdUNkYWY1RTh4M29ER0FXMFovJiN4QTtidUR3L3dDRkZYL0RNL0QyWm1uMDRmZitMY1hKMmhpaDF2M01pc3Z5dHRRQWIyOWR6M1NGUWcvNEp1ZGZ1elk0K3hJL3hTK1RnNU8xJiN4QTtqL0RINXB6YmVRdks4QUZiUXpNUDJwWGR2d0JDL2htYkRzdkJIcGZ2Y1NYYU9ZOWFUR0x5OW9NUXBIcDFzUGYwa0orOGl1WkVkSmlIJiN4QTtLTWZrMG5VNUR6a2ZtaWtzYkpEVkxlTlQwcUVVZnd5MFlvamtBMW5KSTlTNTdLeWZkN2VOaU9sVVUvckdKeFJQTUJSa2tPcFFzM2w3JiN4QTtRWmgrOTA2MmIzOUpLL2ZTdVZ5MHVJODR4K1RaSFU1QnlrZm1sbDErWHZsUzRCcFptRmorMUU3citCSlg4TXg1OW1ZWmRLYjRkcFpvJiN4QTs5YlNHL3dEeWp0V3ExaGZ2R2V5VHFIQi8yUzhLZmRtSms3R0g4TXZtNW1QdGsveFIrVEZ0VS9MenpSWWdzTFlYY1kvYnRqelAvQUVCJiN4QTsvd0RoYzErWHM3TkRwZnVkamk3U3d6NjE3Mk9wTGQyZHh5amVTMnVZajlwUzBicVI4cUVaaUFtSjIyTG5FUmtOOXd5N1EvelQxcXlLJiN4QTt4YWdvdjdjYkZqOEVvSHN3MlAwajZjMldEdFRKSGFYcUgydXMxSFpPT1c4ZlNmc2VqYUY1dDBMVzFIMUs0SHIwcTF0SjhFby8yUGY1JiN4QTtyVVp1c0dyeDVmcE8vZDFkSHFOSGt4ZlVOdS9vbkdaTGl1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFZTdXJ1MnRJR251JiN4QTtaRmloUVZaMk5Ca01tU01CeFNOQmxDQmthQXNzQTh3Zm1EY3o4cmZTZ1lJZWh1R0g3eHY5VWZzajhmbG5PYXp0aVV2VGoySGYxL1k3JiN4QTtyVGRtZ2J6M1BkMFlsRGIzbDljK25DajNGeEthMEZXWWs5U2Y2NXFJeGxrbFFzeUxzcFNqQVdkZ3pIUnZ5Mmxla3VyUytrdlg2dkVRJiN4QTtXLzJUN2dmUlg1NXU5TjJLVHZrTmVRZFhuN1VBMmdQaVdaNmRvdWw2YW5HeXRraDdGd0t1Zm01cXgrL04zaDAyUEY5SXAxV1hQUEo5JiN4QTtSdEc1ZTFPeFYyS3V4VjJLdXhWMkt1eFYyS3V4Vkw5VjBEUjlXajRhaGFSem1sRmtJbzQrVGlqRDc4cHk2ZUdUNmhiZGgxR1RHZlNhJiN4QTtZQnIvQU9VczhZYWJSWi9XSFg2ck9RSC9BTmkreW42YWZQTlJuN0pJM2diOGk3clQ5c0E3WkJYbUdBWE50ZjZkZG1LZU9TMXVvaldqJiN4QTtBbzZrZENQNEVacXBSbEEwZGk3cUVvempZM0JadjVXL05TN3RTbHJyZ056YjlGdTFINzFmOWNmdGo4Zm5tMDAzYWNvN1QzSGYxZFJxJiN4QTsreUl5M3g3SHU2UFVMSytzNzYyUzZzNWxudDVCVkpFTlI4dlkrSXpkd25HUXVKc1BQNU1jb0dwQ2lyNU5nN0ZYWXE3RlhZcTdGWFlxJiN4QTs3RlhZcTdGWFlxbHV1YS9ZYVBiR1c0YmxLd1Bvd0w5cHovQWVKekUxZXNoZ2pjdWZRZDdrYWZUeXltaHk3M2wrdDY5cUdyM0hxWEwwJiN4QTtqVS91b0Yrd2c5aDQrK2NqcXRaUE5LNWN1NTZIQnA0NGhRVER5OTVLdjlVNHp6MXRySTdoeVBqY2Y1QVA2eitPWmVpN01ubTNQcGg5JiN4QTsvdWFOVHJvNDlodko2TnBlajZkcGtIbzJjSWpCKzAvVjJQaXpIYzUwK0RUUXhDb0NuUlpjODhodVJSbVh0VHNWZGlyc1ZkaXJzVmRpJiN4QTtyc1ZkaXJzVmRpcnNWZGlyc1ZTL1dkQTBuV2JiNnZxRUN5cVBzUDBkRDRxdzNHVTVzRU1ncVFiOEdvbmlOeE5QSlBOdjVkNm5vdk82JiN4QTt0cTNtbWpjeXFQM2tZLzRzVWR2OG9iZkxOQnF0QkxGdU40dlNhUHRLR1hZK21YMys1Si9Mbm1uVmRBdXZXczVLeE1SNjlzMVRISVBjJiN4QTtkajRFWlJwOVRQRWJpNVdwMGtNMGFsejczdGZsbnpWcGZtQ3o5YTBiak1sUFh0bUk1b2Y0cjROblI2YlZSeWl4ejdubE5WcEo0WlZMJiN4QTtsMEtjWmt1SzdGWFlxN0ZYWXE3RlhZcTdGWFlxbEhtUHpIYTZOYThtcEpkU0Q5eEI0LzVUZUNqTUhYYTZPQ1BmSThnNVdsMHB5bnllJiN4QTtXM3Q3ZWFoZHRQY09aWjVEVHgrU3FQRHdHY2RseXp5eTRwR3lYb29ZNHdqUTJETlBLL2taSXdsN3F5Y3Blc1ZvZWkrNytKOXZ2em9PJiN4QTt6K3lRUFhsRy9kK3QxT3I3UXYwdytiTmMzN3FYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWEVBaWgzQjZqJiN4QTtGWG5IbmY4QUxKSmhKcU9oUmhKaFZwckZkbGZ1VEY0SC9KNzl2ZlQ2enM2L1ZqK1g2bmU2RHRVajBaT1hmK3Q1dHB1cGFocE44bDFaJiN4QTt5TkJjeEduaDBPNnNPNDhRYzArUEpLRXJHeGQ5bHhSeVI0WkN3WHVIay96aFplWXJMa3RJcjZJRDZ6YlY2ZjVTK0tuOE02WFNhc1pZJiN4QTsvd0JMcUhrZGJvcFlKZDhUeUxJTXkzQ2RpcnNWZGlyc1ZkaXJzVlM3WHRidHRJc1d1SmFOSWZoZ2k3dS85QjNPWW1zMWNjRU9JOCtnJiN4QTs3Mi9UNERsbFFlVTMxNWRhaGVQY1RzWko1VC90S284QjJ6aXN1V1dXWmxMY2w2VEhDTUkwT1FaOTVSOG9ycDZyZlh5aHI1aFdPTTdpJiN4QTtJSC9qYk9uN003TThMMXordjd2MnVsMXV0NC9USDZmdlpUbTVkYzdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYJiN4QTtZcTdGWFlxN0ZXRGZtQjVBVFZrZlU5TVFMcWFpc3NRMkU0SC9BQnY0SHZtczEyaDQvVkg2dnZkdjJkMmo0Zm9uOUgzZnNlVDZicVY5JiN4QTtwT29SM2xxNWl1WUc3MUhUcXJEd1BRak5IanlTaEt4ekQwbVhGSEpFeGx1Qzk0OHFlWjdQekRwaTNjRkVuU2kzVnZYZU42ZjhSUDdKJiN4QTt6cHROcVJsalk1OVhqdFpwSllKOEo1ZENuT1pMaXV4VjJLdXhW",
            "MktxVjFjd1d0dkpjVHNFaGlVczdId0dReVpJd2laUzVCbENCa2FIJiN4QTtNdko5ZTFtNDFhL2U1a3FzWStHQ0xzcWRoOC9IT0cxbXJsbm1aSGwwSGNIcGRQZ0dLTkJsWGtqeXdJMVRWYnhQM2piMnNaL1pIODU5JiN4QTt6MnplZGtkbjBQRm56NmZyL1U2N1g2dS9SSDRzenpvSFV1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWJiN4QTsyS3V4VjJLdXhWNXYrWjNrZFpvNU5lMDZQOThueFg4S2o3U2ovZG9IaVAydkhyODlQMmpvNy9lUitQNjNmZGxhK2o0YytYVDlUQXZLJiN4QTszbVM2OHY2dEhldzFhSS9CY3dWMmtqSjNIekhVSHh6V2FiT2NVK0lPNTFlbGptaHdubjBlL1dGOWEzOW5EZVdyaVMzblFQRzQ4RDQrJiN4QTs0N2pPcGhNU0FJNUY0dkpqTUpHSjVoWHlUQjJLdXhWMktzQjgrYTZaNS8wWEEzN21FMXVDUDJwT3kvSmYxL0xPVzdhMXZGTHdvOGh6JiN4QTs5LzdQdmQxMmRwNkhHZVo1SUh5ZjVmOEEwbmUrdk90YksySUxnOUhmcUUvaWN4dXl0RDQwK0tYMFIrM3lidGJxZkRqUStvdlRBQUJRJiN4QTtkTTdKNTkyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWeEFJSUlxRDFHS3ZFUHpIJiN4QTs4by9vUFUvclZxbE5OdkNXaUE2UnlkV2orWGRmYjVaemV1MHZoeXNmU1hyZXpOWjRzS1AxeCszelRUOHAvTlp0cnM2RmR2OEE2UGNrJiN4QTt0Wmsvc3lucW55ZnQ3L1BMK3pkVHduZ1BJOHZlNC9iR2o0bytKSG1PZnUvWTliemV2Tk94VjJLcGQ1ZzFWZEwwdVc1cis5UHdRRHhrJiN4QTticDkzWDZNdzlmcXZBeEdYWGtQZTM2YkQ0a3dPblY1WEREY1hsMmtTVmtubmVncjFMTWVwT2NOQ0Vza3dCdktSZWtsSVFqWjVCNjFwJiN4QTtPbXc2YllSV2tYUkI4YmZ6TWZ0TWZtYzczUzZjWWNZZ09qek9iS2Nraklvdk1ocWRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzJiN4QTtWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcVhlWWRFdHRiMGk0MDZmWVNyV09UcVVrRzZNUGtmd3luUGhHU0JpVy9UWnppbUpEJiN4QTtvK2VicTN2Tk0xQ1MzbEJodTdTUXFhYkZYUTlRZnhCemxwUk1aVWVZZTNoS09TTmplSkQzM3lmNWdUWGRBdHI2bzllbnBYU2p0TW16JiN4QTtmOEY5b2V4enA5TG04U0FsMTZ2R2EzVGVEbE1lblQzSjFtUTRqc1ZlZCtlOVVOenFZczBQN20wRkQ3eU51MzNDZ3prTzI5Vng1ZUFjJiN4QTtvZmU3M3M3RHd3NHVza1orWCtrODVaZFRrRzBkWW9LL3pFZkUzMEEwekk3QjB0azVUMDJINldydExOUUVCOFdjWjA3cDNZcTdGWFlxJiN4QTs3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxOG8vT1B5OEk1NE5kZ1g0WnFRWGRQJiN4QTs1d1AzYkg1cU9QMEROSjJwZ29pWTY4M3BPeE5UWU9NOU54K2xMdnlrMTgyT3ZOcGtyVXQ5UldpZ25ZVElDVlAreUZWKzdLK3pjM0RQJiN4QTtoUEtYM3QzYk9uNDhmR09jZnVlMFp2M2xWRyt1MHRMT2E2azNXRkdjaXRLMEZhYitQVEtzK1VZNEdaL2hGczhjREtRaU9yeUYybXViJiN4QTtobWFyelRPU2ZGbWMvd0FTYzg4SmxPVjg1RS9lOVFBSWp5RDFuU2JCTERUb0xSZjkxSUF4OFdPN0g2U2M5QTB1QVlzY1lEb1A3WG1jJiN4QTsyVGptWmQ2THpJYW5ZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcWx2JiN4QTttUFI0OVkwUzgwNTZWbmpJalk5RmtIeEkzME1CbE9mRjRrREZ2MDJZNHNnbjNQbktLVzZzYjFKRnJGZFdzZ1phN0ZaSTJxUHBCR2N3JiN4QTtDWW0rb2U2SUU0MXpCRDZVMGpVb2RUMHUxMUNIYU82aldRTFd2RXNOMVA4QXFuYk9weFpCT0lrT3J3ZWJFY2N6RTlDa25uNjg5TFNZJiN4QTs3WlQ4VnpJS2ovSVQ0ai93M0hOUDI5bTRjSWovQURqOWcvQWN2czdIYzc3bU0rVHJFWFd1d2xoVkxjR1p2OWo5bi9oaU0wZlkrSHhOJiN4QTtRTzZPL3dDUGk3SFhaT0hHZlBaNlpuY1BQT3hWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1JiN4QTt4VjJLdXhWMkt1eFYyS3V4VjRIK2FHbERUZk45MFZYakZlZ1hVZnprcUgvNGRXem5OZGk0Y3A4OTNzK3ljM2lZQjN4Mi9Id1o3K1RXJiN4QTtyZld2TGsxaXpWa3NKaUZIaEhOOGEvOEFEODgyWFpzN2dZOXhkTjIzaDRjb2wvT0gyajhCM242NjlUVm9vQWZoZ2lGUi9sT1NUK0ZNJiN4QTswSGIrWGl6Q1A4MGZmK0FuczJGUUo3eW1INWVXbElMdTdJKzJ5eEtmOVVjai93QVNHWm5zOWk5TXArZE5IYWM5eEZsK2RJNnQyS3V4JiN4QTtWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt2TC96eDA2dGxwdXBxJiN4QTtONHBIdDVENGlSZWEvZHdiNzgxWGFjTmhMNFBROWdaZlZLSGxhUi9rcnFaZzh5M0ZrekVSM2x1YUwyTWtSRExYNUlYeWpzNmRaSzd3JiN4QTs1ZmJ1SzhRbC9OUDMvZ0oxNWxuOWZYcjEvQ1FwL3dBaXdFLzQxem0rMDhuRnFKbnpyNWJPTnBJMWlqN21jZVQ0UFI4djIyMUdrNVNOJiN4QTs5TEduNEFaMW5ZMlBoMDBmT3o5cnA5ZEs4cFRyTm80anNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpJiN4QTtyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVllsK2F0bUxueVBma0NyMjVqbVQyNHlBTi93aE9ZZXVqZUkrVHMreUo4T29qNTJQc2VOZmw5JiN4QTtlbTE4NmFSS0RUbGNMRC95T0JpLzQzelVhV1ZaSSs5Nm50S0hGcDVqeXY1YnMwdlpETGVUeW5jeVNPeEo5Mkp6bU04K0tjajNrdXJ4JiN4QTt4cUlIazlTMGFNUjZSWko0UVIxK2ZFRTU2Qm9vOE9HQS9vajdubk01dWN2ZVVabVUxT3hWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4JiN4QTtWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VktmTjBBdVBLdXJ3bjlxem40LzZ3allqOGNwMUF2SEwzT1RvJiN4QTs1Y09hQi9wRDczemJvTjhMTFhkT3ZDdk1XMTFETVVCcHk5T1JXcFgzcG5QNDVjTWdlNHZlYWpIeFk1Ujc0a2ZZOUZwbktPbGV2V3FsJiN4QTtiV0ZUMVZGSDNBWjZiaUZRQThnOHRNN2xWeXhpN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYJiN4QTtZcTdGWFlxN0ZYWXE3RlhZcTdGVU5xY1prMDI3akJv",
            "WGhrVUUrNkVaRElMaWZjMllqVXdmTjhtOHZmT2JwOUpwNjFUT1VlY2V1V3JGJiN4QTtyV0ZqMUtLVDlJR2VuWVRjQWZJUEx6SHFLcmxqRjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLJiN4QTt1eFYyS3V4VjJLdXhWMkt1eFYyS3JaYWVrOVJVY1RVSG9kc3J6R29FK1JUSG13ai9BQXo1Wi82dEZsLzBqeGY4MDU1UitieS96NWZNJiN4QTt1OS9NNWY1MHZtVWcxQ0VRWDl6Q0R5RVVyb0c4ZUxFWm5hcUhCbGxIdWtSOXJrWXBYRUh2RDB6UjVCSnBObTQ3d1IxK2ZFQTU2SG9aJiN4QTtjV0NCL29qN25uczRxY3ZlVVptVTFPeFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLJiN4QTt1eFYyS3V4VjJLdXhWU3VtNDJzcmVDTityTVB0Q2ZEcDhoL29TKzVuakZ5REhvQnpualRweVlDdnpPZVg2UEh4NW9SL25TaVBtWFp6JiN4QTtOQWxqSG1PRDBkY3ZFOFpDL3dEd2Z4LzhiWnQrMVljR3BtUE8vbnU1bWxsZU9QdVpyNVNuOVhRYmZlcGo1SWZvWTAvQWpPdjdFeWNXJiN4QTtsajVXUHRkUnJZMWxLY1p0bkVkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyJiN4QTtzVmRpcnNWZGlxRTFhVGhZeWVMVVVmU2MwM3RCbDROSlB6b2ZiK3B2MDR1WVNiVGw1M3NJL3dBcXYvQTcvd0FNNGJzYkh4NnZHUDZWJiN4QTsvTGY5RG5aalVDbFhubTBNZXBSWEFId1R4MHI0c2hvZndLNXZmYVBEdzVoUHBJZmFQMlUyZG5UdUJIY1VkNUN1YXdYVnFUdWpDUlIvJiN4QTtyRGlmK0lqTTMyYXpYR2NPNDM4LzdHbnRLRzRreXZPb2RZN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGJiN4QTtYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGVXAxK2FrY1VYaVN4K2pZZnJ6a1BhelBVSVkrODM4dHYwdVpvNDdrcUdoUjhycHBLYkl2JiN4QTtYM1A5bGMxL3NyZzRzOHA5SXgrMC9zdHQxY3FqU241MnMvVzBwYmdLQzl2SUNXN2hIK0UvOE54em9QYUxCeFlCUHJBL1lkdnZwajJmJiN4QTtPcDEzc2Q4cDNndGRaaUJORW5CaGIvWmZaLzRZRE9kN0QxSGg2a1h5bDZmbnkrMXo5Ymo0c1o4dDNvZWVndWhkaXJzVmRpcnNWZGlyJiN4QTtzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyR3RZbjlXK2NBL0RIOEEranIrT2VaJiN4QTsrME9wOFhWU3JsRDAvTG45dHUwMDBLaDcwejBHSGhhTkllc3JiZkpkdjExenB2WmJUY0duTXovR2ZzRzMzMjR1cmxjcTdrWmUycVhkJiN4QTtwTmJQOW1WQ3RmQWtiSDZNNkRVWVJseHlnZjRoVFJqbnd5QjdubDdKTGJ6c3Bxa3NMRUh4REtmNjU1aElTeHpybEtKKzBQUmdpUThpJiN4QTs5TjB5OVc5c0libGY5MktDd0haaHN3Ky9QUzlGcVJteFJtT28rM3E4N214OEVqRkU1bE5ic1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpJiN4QTtyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVlVydWNXOXRKTWYyUnNQZm9QeHpFMTJxR0RETElmNFI5dlQ3V2VPJiN4QTtIRklCaVlEeXlBRDRuYzArWkp6eWVJbGxuWE9VajlwZHhzQXk2Q0pZWVVpWG9nQXoxelRZQmh4eGdPVVJUcHB5NGlTdnk5aXdqem5wJiN4QTtvZ3ZrdTBGSTdrZkdCMmRhQTl1NHArT2NQN1I2VHc4b3lEbFA3eCt2OWJ1ZEJsdVBDZWlJOGs2bHhrazArUTdQKzhoci9NQjhRKzdmJiN4QTtNajJhMXRFNFQxM0g2V3Z0REZZRXd5L093ZFU3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZJiN4QTtxN0ZYWXE3RlhZcTdGVWo4d1hkV1cyVS9aK0tUNTlobkVlMVd2c2pCSHB2TDlBL1Q4blAwbVA4QWlVdEJ0VExjbVpoOEVYVC9BRmowJiN4QTt6RTltTkY0bWJ4VDlPUDd6K1B1WjZ2SlVhNzJRNTZFNjEyS29IV3RPWFVOT2x0LzkyZmFpSjdPdlQ3K21ZSGFXakdvd21IWG1QZjhBJiN4QTtqWnUwK1hnbUM4N2hrbnRibFpFcWswTFZGZXpLZWh6emVHU2VLWWtOcFJQM08va0JJVjBMMG5UYitLK3M0N21QbzQrSmY1V0hVWjZiJiN4QTtvdFhIVVloa2oxK3c5UTg5bHhtRWlDaWN5bXQyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4JiN4QTtWMkt1eFYyS3FOM2NwYlc3elAwVWJEeFBZWmlhN1Z4MCtLV1NYVDdUMERQSEF5bFFZazhrazB4ZHZpa2tOZm1UbmsrWEpQTmtNanZPJiN4QTtSKzkzSUFpSzZNcjArMUZyYXBGKzExYy81UjY1NmwyWG9ocHNFWWZ4Y3o3eitLZFJteWNjclJHYkZxZGlyc1ZZZjV2MGIwcGYwaENwJiN4QTs5T1EwdUFCc3JkbS8yWDYvbm5GZTBmWnZETHg0RFkvVjVIditQMys5MjJoejJPQS9CQ2VXZFkrb1hYcFN0UzFuSUQxNkszWnY2NWc5JiN4QTtoOXAvbDhuREkvdTVmWWUvOWY3RzNXWU9PTmo2Z3p2UFEzU094VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1JiN4QTt4VjJLdXhWMkt1eFYyS3V4VmkydGFsOWFuOU9NL3VJalJmOEFLUGMvMHp6WDJnN1YvTTVlQ0IvZHcrMDkvd0NyOXJ0dE5oNEJaNWxFJiN4QTsrWDdBeVNmV3BCOENiUmp4Yngrak03Mlk3TTQ1K1BNZW1QMCsvdjhBaDkvdWF0WGxvY0laQm5lT3VkaXJzVmRpcXllQ0tlRjRabER4JiN4QTtTRGk2bnVEbGVYRkhKRXhrTGlXVVpHSnNjM25lcmFYTHAxNDBEL0VoK0tLVCtaZjYrT2VZOXBhQ1dseW1CNWREM2o4YzNmNE13eVJ0JiN4QTtrUGxiWFE2cnA5eTN4cnRidWU0L2xQdU8yZEw3UDlyOFFHRElmVi9DZS95L1U0R3QwMWV1UHhaTm5XdXRkaXJzVmRpcnNWZGlyc1ZkJiN4QTtpcnNWZGlyc1ZkaXJzVmRpcnNWZGlyc1ZkaXJzVmRpcnNWZGlxUjY5cXdRTlp3TjhaMm1ZZGgvTC9YT1A5bysyZUVIQmpQcVAxSHU4JiN4QTt2MXVkcGNGK28vQktOT3NwTHk0RVM3S041SDhGemxlek96cGFyS0lEbDFQY1B4eWN6TmxFSTJ6R0tLT0tOWTR4eFJCUlJucXVIREhGJiN4QTtBUWlLakYwMHBFbXl1eTFEc1ZkaXJzVmRpcUMxYlM0ZFJ0VEMvd0FMamVLVHVyZjA4YzEvYVhaOE5WajREejZIdVA0NXQyRE1jY3JEJiN4QTtBTG0ydUxPNWFHVUZKWXovQUxSQnp6VFVZTW1ESVlTMmxIOFdIZXdtSml4eVpkNWU4d2k4VVd0MFFMb0Q0WDZDUUQvamJPMzdFN2JHJiN4QTtjREhrUDd6cC9TL2E2clZhWGc5VWZwKzVQ",
            "czZSd1hZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZJiN4QTtxbEd0YTBMWUczdHpXNFAybTdKL2JuTWR1OXVqVGc0c1IvZTlUL04vYTVtbTAzRnVlVEhMZUNhNW5XS01GcEhQKzJUbkNhZlQ1TlJrJiN4QTtFSTd5bCtMTHNweUVSWjVNeDArd2lzcmNScHV4M2tmdVRucVhablowTkppNEk4K3A3eitPVHBzMlV6TmxFNXNHcDJLdXhWMkt1eFYyJiN4QTtLdXhWTDlaMGFEVW9LR2lUb1AzVXZoN0gyelY5cWRsdzFjS08weHlQNDZPUnA5UWNaOG1DM05yYzJsd1labE1jcUgvYUlPZWJhalQ1JiN4QTtNR1RobU9HUS9GaDNjSmlZc2NtUzZINW5EaGJhL2FqOUVuUFEremUvdm5YZGorMFFsV1BPZCtrdS93Qi82M1c2blIxNm9mSmt1ZGU2JiN4QTs1Mkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLcERxL21GVURRV2JWZm84dzZEL1Y5L2ZPUDdaOW94RzhlJiN4QTtBM0xyTHU5M241dWZwOUplOHZra052QlBjemlPTUY1SFArWkp6amRQcDhtb3ljTUJ4U2wrTExuemtJaXp5WmZwZWx4V01WQjhVemYzJiN4QTtrbjhCN1o2WjJUMlREU1FvYnpQTS9vSGs2alBuTXo1STNOczBPeFYyS3V4VjJLdXhWMkt1eFYyS29QVTlLdGRSaDRUQ2pqKzdsSDJsJiN4QTtQOVBiTmYyaDJiaTFVT0dZMzZIcVB4M04ySFBMR2JEQ2RUMG03MDZiaE1LbzMySlYreTM5dnRublBhUFptWFN5cWZMb2VoL2I1TzV3JiN4QTs1NDVCWVJla2VZN3F4cEZMV2EyL2xKK0pSL2tuK0daM1pmYitUVFZHWHJ4L2FQZCtyN21yUHBJejNHeFpmWmFoYVhzWHFXOGdjZnRMJiN4QTswWWZNWjN1ajEyTFVSNHNadjd4N3c2bkppbEEwUWlNeTJ0Mkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMktxRjVmV3RwSHpuJiN4QTtjTC9LdlZqOGhtSnJOZGkwMGVMSkt2dlB1RFpqeFNtYUFZdnFldjNGNVdPS3NWdi9BQ2o3VEQvS1A4TTREdFQyZ3k2bTRROUdQN1Q3JiN4QTsvd0JYM3UwdzZVUTNPNVF1bjZkYzMwbkdKZmhIMjVEOWtacit6K3pNdXJsVUJ0MVBRZmp1Yk11YU1CdXkvVDlOdDdHTGhFS3VmdHlIJiN4QTtxeHowbnM3c3pGcEljTU9aNW5xZngzT295NWpNN29yTmkxT3hWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3JKb1lwNG1pbVFQRzRveW5vJiN4QTtjcnk0WTVJbU14Y1QwWlJrWW14ell0cXZsU1dJdExZMWtqNm1FL2JIeThmMTV4SGFmc3pPRnp3ZXFQOEFONi9EdisvM3V6d2E0SGFXJiN4QTt5UlJTM0Z0TnpqWm9wVVBVVkJIdG5MNDh1VERPNGt4bVBnNTBvaVEzM0RJTlA4M3V0RXZrNWovZnFVQitsZW4zWjFlZzlxeVBUbkYvJiN4QTswaCtrZnErVGdaZEFPY1dRMm1vV1YydGJlWlpQRlJzdythbmZPdDB1dnc1eGVPUWw5L3k1dXZuaWxEbUVSbVcxdXhWMkt1eFYyS3V4JiN4QTtWMkt1eFYyS3V4VlF1cit6dEZyY1NxbmdwM1kvSURmTVRWYS9EcHhlU1FqOS93QXViWkRGS1hJSkJmZWEzYXFXU2NCL3YxOXo5QzlQJiN4QTt2emtkZjdXRStuQUsvcEg5QS9YOG5QeGFFYzVKSEpKUGNTOG5acFpYTkttcEo5czVQSmx5WnAzSW1jejhTNXdBaU50Z25lbWVXWlpLJiN4QTtTM2xZMDZpSWZhUHo4TTZqc3oyWG5PcDUvVEgrYjErUGQ5L3VjTE5yUU5vN3NraGhpaGpFY1NCRVhvb3p1TU9HR0tJakFDTVE2MlVqJiN4QTtJMlYrV3NYWXE3RlhZcTdGWFlxN0ZYWXE3RlhZcTdGWFlxN0ZYWXFnNy9TTEMrWDkvSCs4N1NyczQrbnY5T2E3WGRsWU5TUFhIMWZ6JiN4QTtoc2Z4Nzdic1dlVU9SWTNmZVU3NkVscllpNGo2MCt5NCtnN0g2TTQzVyt5K2ZIdmkvZVIrUi9IdStUc2NXdWlmcTJTWmxtZ2tvd2FLJiN4QTtWVDBOVllIOWVjNUtNOFV0d1l5SHdMbUFpUTd3bU5yNWsxVzNvREw2eWo5bVVjdngyYjhjMitsOW90WGkyNHVNZjB0L3Q1L2E0ODlIJiN4QTtqbDByM0pyYitjWVR0Y1c3TC9sUmtOK0JwbSt3ZTE4RC9lUUkvcW0vdnB4Wjlubm9Vd2g4eWFQSi91N2dmQjFZZmpTbWJiRjdSYU9mJiN4QTs4ZGU4SCt4eDVhUElPaUtUVTlPY2ZEZFJIMjVyWDllWjBPMHROTGxraC9wZzFIRE1kQ3FpNHR5YUNWQ2ZBTU15Qm54bmxJZk5qd0h1JiN4QTtjYmkzQm9aVUI4Q3d4T2ZHT2NoODE0RDNLVDZucHlmYXVvaDdjMXIrdk1lZmFXbWp6eVEvMHdaRERNOUNoWmZNZWtSLzd1NW53UldQJiN4QTs0MHBtRGw5b3RIRCtPL2NEL1kyeDBlUTlFRFA1d2hGUmIyN01lek9RdjRDdWFuUDdYd0g5M0FuM212dXR5STlubnFVcXVmTWVxWEZSJiN4QTs2Z2hVOW94eC9IYy9qbWgxWHRGcTh1M0Z3RCtqdDl2UDdYSmhwTWNlbG9CVm1ta29vYVNSdXdxekU1cDR4bmxsc0RLUitKY2drUkhjJiN4QTtFNHNmSzk3TlJyZ2lDUHdPN242TzJkSG9mWmZQazN5L3U0L00vajhVNGVYV3hITGRrVmpwVmxaRDl6SDhmZVJ0MlAwLzB6czlEMlZnJiN4QTswdzlFZlYzbmMvajNPdnk1NVQ1bEY1c1dsMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLdXhWMkt1eFYyS3V4VjJLcVZ4YTIxd3ZHJiN4QTtlSlpCL2xBR255OE1vejZiSGxGVGlKRHpES001UjVHa3B1ZktXblNWTUxQQXg2QUhrbytnNy9qbWcxUHNycHAyWUdVRDh4OXUvd0JyJiN4QTtsdzEweHozU3U0OG82Z2hKaGtTVmUzVlcrNDdmam1peit5ZW9qOUVveUh5UDZ2dGNxT3ZnZVlJUU11aDZ0RjlxMWMvNm81LzhScm1xJiN4QTt5OWlhdUhQSEw0Yi9BSFczeDFPTTlVTTF0Y3A5dUoxK2FrWmhUMHVXUE9NaDhDMkNjVDFVOG9aT3hWVVMxdVgreEU3ZkpTY3ZocGNzJiN4QTt1VVpINEZpWnhIVkZSYUpxMHYyYlZ4L3JEaC94S21adUxzVFZ6NVk1ZkhiNzZhNWFuR09xT2c4cGFnKzhycEVQbVdQNGJmam0xd2V5JiN4QTtlb2w5WmpIN1QrUGkwUzE4QnlzcG5iZVU3Q09obWQ1MjdqN0svY04veHpkNmIyVTA4TjVtVXo4aDltLzJ1TFBYVFBMWk5yZTF0cmRlJiN4QTtNRVN4anZ4RksvUE9od2FYRmhGWTRpSThuRW5PVXVaVmN2WXV4VjJLdXhWMkt1eFYyS3V4VjJLdXhWLy8yUT09PC94bXBHSW1nOmltYWdlPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgPC9yZGY6QWx0PgogICAgICAgICA8L3htcDpUaHVtYm5haWxzPgogICAgICAgICA8eG1wTU06SW5zdGFuY2VJRD51dWlkOjBmZGM3NjU3LTAwNjgtYWE0Yi04MGQzLWFiNmFlYzJhOTlmYzwveG1wTU06SW5zdGFuY2VJRD4KICAgICAgICAgPHhtcE1NOkRv",
            "Y3VtZW50SUQ+eG1wLmRpZDpiOGY0MGFjNy03NDI0LTQ0ZGQtYTY3NS0zM2RiZDU2NjI1Y2U8L3htcE1NOkRvY3VtZW50SUQ+CiAgICAgICAgIDx4bXBNTTpPcmlnaW5hbERvY3VtZW50SUQ+dXVpZDo1RDIwODkyNDkzQkZEQjExOTE0QTg1OTBEMzE1MDhDODwveG1wTU06T3JpZ2luYWxEb2N1bWVudElEPgogICAgICAgICA8eG1wTU06UmVuZGl0aW9uQ2xhc3M+ZGVmYXVsdDwveG1wTU06UmVuZGl0aW9uQ2xhc3M+CiAgICAgICAgIDx4bXBNTTpEZXJpdmVkRnJvbSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgIDxzdFJlZjppbnN0YW5jZUlEPnV1aWQ6ZTY4OGEwYTEtYWE3Ny1mNzQzLThjOTUtYmM2OGYyMTc0OWY2PC9zdFJlZjppbnN0YW5jZUlEPgogICAgICAgICAgICA8c3RSZWY6ZG9jdW1lbnRJRD54bXAuZGlkOjg4ZDFlODM1LWQ2OWEtNGM2NS1iNGY2LWQ1YmU2NDVkMWI1Yjwvc3RSZWY6ZG9jdW1lbnRJRD4KICAgICAgICAgICAgPHN0UmVmOm9yaWdpbmFsRG9jdW1lbnRJRD51dWlkOjVEMjA4OTI0OTNCRkRCMTE5MTRBODU5MEQzMTUwOEM4PC9zdFJlZjpvcmlnaW5hbERvY3VtZW50SUQ+CiAgICAgICAgICAgIDxzdFJlZjpyZW5kaXRpb25DbGFzcz5wcm9vZjpwZGY8L3N0UmVmOnJlbmRpdGlvbkNsYXNzPgogICAgICAgICA8L3htcE1NOkRlcml2ZWRGcm9tPgogICAgICAgICA8eG1wTU06SGlzdG9yeT4KICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6YWN0aW9uPnNhdmVkPC9zdEV2dDphY3Rpb24+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDppbnN0YW5jZUlEPnhtcC5paWQ6ODhkMWU4MzUtZDY5YS00YzY1LWI0ZjYtZDViZTY0NWQxYjViPC9zdEV2dDppbnN0YW5jZUlEPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6d2hlbj4yMDI2LTExLTIxVDE0OjM0OjM0KzA5OjAwPC9zdEV2dDp3aGVuPgogICAgICAgICAgICAgICAgICA8c3RFdnQ6c29mdHdhcmVBZ2VudD5BZG9iZSBJbGx1c3RyYXRvciAzMS4wIChNYWNpbnRvc2gpPC9zdEV2dDpzb2Z0d2FyZUFnZW50PgogICAgICAgICAgICAgICAgICA8c3RFdnQ6Y2hhbmdlZD4vPC9zdEV2dDpjaGFuZ2VkPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDphY3Rpb24+c2F2ZWQ8L3N0RXZ0OmFjdGlvbj4KICAgICAgICAgICAgICAgICAgPHN0RXZ0Omluc3RhbmNlSUQ+eG1wLmlpZDpiOGY0MGFjNy03NDI0LTQ0ZGQtYTY3NS0zM2RiZDU2NjI1Y2U8L3N0RXZ0Omluc3RhbmNlSUQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDp3aGVuPjIwMjYtMTEtMjFUMTQ6NTA6MTYrMDk6MDA8L3N0RXZ0OndoZW4+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpzb2Z0d2FyZUFnZW50PkFkb2JlIElsbHVzdHJhdG9yIDMxLjAgKE1hY2ludG9zaCk8L3N0RXZ0OnNvZnR3YXJlQWdlbnQ+CiAgICAgICAgICAgICAgICAgIDxzdEV2dDpjaGFuZ2VkPi88L3N0RXZ0OmNoYW5nZWQ+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICA8L3JkZjpTZXE+CiAgICAgICAgIDwveG1wTU06SGlzdG9yeT4KICAgICAgICAgPGlsbHVzdHJhdG9yOlR5cGU+RG9jdW1lbnQ8L2lsbHVzdHJhdG9yOlR5cGU+CiAgICAgICAgIDxpbGx1c3RyYXRvcjpTdGFydHVwUHJvZmlsZT5QcmludDwvaWxsdXN0cmF0b3I6U3RhcnR1cFByb2ZpbGU+CiAgICAgICAgIDxpbGx1c3RyYXRvcjpDcmVhdG9yU3ViVG9vbD5BZG9iZSBJbGx1c3RyYXRvcjwvaWxsdXN0cmF0b3I6Q3JlYXRvclN1YlRvb2w+CiAgICAgICAgIDxpbGx1c3RyYXRvcjpJc0ZpbGVTYXZlZFZpYUluc3RhbnRTYXZlPkZhbHNlPC9pbGx1c3RyYXRvcjpJc0ZpbGVTYXZlZFZpYUluc3RhbnRTYXZlPgogICAgICAgICA8eG1wVFBnOkhhc1Zpc2libGVPdmVycHJpbnQ+RmFsc2U8L3htcFRQZzpIYXNWaXNpYmxlT3ZlcnByaW50PgogICAgICAgICA8eG1wVFBnOkhhc1Zpc2libGVUcmFuc3BhcmVuY3k+RmFsc2U8L3htcFRQZzpIYXNWaXNpYmxlVHJhbnNwYXJlbmN5PgogICAgICAgICA8eG1wVFBnOk5QYWdlcz4xPC94bXBUUGc6TlBhZ2VzPgogICAgICAgICA8eG1wVFBnOk1heFBhZ2VTaXplIHJkZjpwYXJzZVR5cGU9IlJlc291cmNlIj4KICAgICAgICAgICAgPHN0RGltOnc+NDYuMjEzODg5PC9zdERpbTp3PgogICAgICAgICAgICA8c3REaW06aD40Ni4yMTM4ODk8L3N0RGltOmg+CiAgICAgICAgICAgIDxzdERpbTp1bml0Pk1pbGxpbWV0ZXJzPC9zdERpbTp1bml0PgogICAgICAgICA8L3htcFRQZzpNYXhQYWdlU2l6ZT4KICAgICAgICAgPHhtcFRQZzpQbGF0ZU5hbWVzPgogICAgICAgICAgICA8cmRmOlNlcT4KICAgICAgICAgICAgICAgPHJkZjpsaT5DeWFuPC9yZGY6bGk+CiAgICAgICAgICAgICAgIDxyZGY6bGk+TWFnZW50YTwvcmRmOmxpPgogICAgICAgICAgICAgICA8cmRmOmxpPlllbGxvdzwvcmRmOmxpPgogICAgICAgICAgICA8L3JkZjpTZXE+CiAgICAgICAgIDwveG1wVFBnOlBsYXRlTmFtZXM+CiAgICAgICAgIDx4bXBUUGc6U3dhdGNoR3JvdXBzPgogICAgICAgICAgICA8cmRmOlNlcT4KICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgIDx4bXBHOmdyb3VwTmFtZT7liJ3mnJ/oqK3lrprjga7jgrnjgqbjgqnjg4Pjg4HjgrDjg6vjg7zjg5c8L3htcEc6Z3JvdXBOYW1lPgogICAgICAgICAgICAgICAgICA8eG1wRzpncm91cFR5cGU+MDwveG1wRzpncm91cFR5cGU+CiAgICAgICAgICAgICAgICAgIDx4bXBHOkNvbG9yYW50cz4KICAgICAgICAgICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpzd2F0Y2hOYW1lPuODm+ODr+OCpOODiDwveG1wRzpzd2F0Y2hOYW1lPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpt",
            "b2RlPkNNWUs8L3htcEc6bW9kZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6dHlwZT5QUk9DRVNTPC94bXBHOnR5cGU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOmN5YW4+MC4wMDAwMDA8L3htcEc6Y3lhbj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6bWFnZW50YT4wLjAwMDAwMDwveG1wRzptYWdlbnRhPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzp5ZWxsb3c+MC4wMDAwMDA8L3htcEc6eWVsbG93PgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpibGFjaz4wLjAwMDAwMDwveG1wRzpibGFjaz4KICAgICAgICAgICAgICAgICAgICAgICAgPC9yZGY6bGk+CiAgICAgICAgICAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpzd2F0Y2hOYW1lPuODluODqeODg+OCrzwveG1wRzpzd2F0Y2hOYW1lPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzptb2RlPkNNWUs8L3htcEc6bW9kZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6dHlwZT5QUk9DRVNTPC94bXBHOnR5cGU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOmN5YW4+MC4wMDAwMDA8L3htcEc6Y3lhbj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6bWFnZW50YT4wLjAwMDAwMDwveG1wRzptYWdlbnRhPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzp5ZWxsb3c+MC4wMDAwMDA8L3htcEc6eWVsbG93PgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpibGFjaz4xMDAuMDAwMDAwPC94bXBHOmJsYWNrPgogICAgICAgICAgICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgICAgICAgICAgPC9yZGY6U2VxPgogICAgICAgICAgICAgICAgICA8L3htcEc6Q29sb3JhbnRzPgogICAgICAgICAgICAgICA8L3JkZjpsaT4KICAgICAgICAgICAgICAgPHJkZjpsaSByZGY6cGFyc2VUeXBlPSJSZXNvdXJjZSI+CiAgICAgICAgICAgICAgICAgIDx4bXBHOmdyb3VwTmFtZT7jgrDjg6zjg7w8L3htcEc6Z3JvdXBOYW1lPgogICAgICAgICAgICAgICAgICA8eG1wRzpncm91cFR5cGU+MDwveG1wRzpncm91cFR5cGU+CiAgICAgICAgICAgICAgICAgIDx4bXBHOkNvbG9yYW50cz4KICAgICAgICAgICAgICAgICAgICAgPHJkZjpTZXE+CiAgICAgICAgICAgICAgICAgICAgICAgIDxyZGY6bGkgcmRmOnBhcnNlVHlwZT0iUmVzb3VyY2UiPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzpzd2F0Y2hOYW1lPkM9MCBNPTAgWT0wIEs9MTAwPC94bXBHOnN3YXRjaE5hbWU+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOm1vZGU+Q01ZSzwveG1wRzptb2RlPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzp0eXBlPlBST0NFU1M8L3htcEc6dHlwZT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgPHhtcEc6Y3lhbj4wLjAwMDAwMDwveG1wRzpjeWFuPgogICAgICAgICAgICAgICAgICAgICAgICAgICA8eG1wRzptYWdlbnRhPjAuMDAwMDAwPC94bXBHOm1hZ2VudGE+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOnllbGxvdz4wLjAwMDAwMDwveG1wRzp5ZWxsb3c+CiAgICAgICAgICAgICAgICAgICAgICAgICAgIDx4bXBHOmJsYWNrPjEwMC4wMDAwMDA8L3htcEc6YmxhY2s+CiAgICAgICAgICAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICAgICAgICAgICA8L3JkZjpTZXE+CiAgICAgICAgICAgICAgICAgIDwveG1wRzpDb2xvcmFudHM+CiAgICAgICAgICAgICAgIDwvcmRmOmxpPgogICAgICAgICAgICA8L3JkZjpTZXE+CiAgICAgICAgIDwveG1wVFBnOlN3YXRjaEdyb3Vwcz4KICAgICAgICAgPHBkZjpQcm9kdWNlcj5BZG9iZSBQREYgbGlicmFyeSAxOC4wMDwvcGRmOlByb2R1Y2VyPgogICAgICA8L3JkZjpEZXNjcmlwdGlvbj4KICAgPC9yZGY6UkRGPgo8L3g6eG1wbWV0YT4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAg",
            "ICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgIAo8P3hwYWNrZXQgZW5kPSJ3Ij8+DWVuZHN0cmVhbQ1lbmRvYmoNMyAwIG9iag08PC9Db3VudCAxL0tpZHNbOCAwIFJdL1R5cGUvUGFnZXM+Pg1lbmRvYmoNOCAwIG9iag08PC9BcnRCb3hbMTEuNTA0MiAxMS41MDQyIDEzNi41MDQgMTM2LjUwNF0vQmxlZWRCb3hbMC4wIDAuMCAxNDguMDA4IDE0OC4wMDhdL0NvbnRlbnRzIDkgMCBSL0Nyb3BCb3hbMC4wIDAuMCAxNDguMDA4IDE0OC4wMDhdL0xhc3RNb2RpZmllZChEOjIwMjYxMTIxMTQ1MDE2KzA5JzAwJykvTWVkaWFCb3hbMC4wIDAuMCAxNDguMDA4IDE0OC4wMDhdL1BhcmVudCAzIDAgUi9SZXNvdXJjZXM8PC9YT2JqZWN0PDwvRm0wIDcgMCBSPj4+Pi9UaHVtYiAxMCAwIFIvVHJpbUJveFs4LjUwMzg5IDguNTAzODkgMTM5LjUwNCAxMzkuNTA0XS9UeXBlL1BhZ2UvUGllY2VJbmZvPDwvSWxsdXN0cmF0b3IgMTEgMCBSPj4+Pg1lbmRvYmoNOSAwIG9iag08PC9GaWx0ZXIvRmxhdGVEZWNvZGUvTGVuZ3RoIDE1NT4+c3RyZWFtDQpIiZzQMQ7CMAxA0d2n8AXq2omTODtiYqASEjOqOiBRoBHcny6gLAyNPNjTk/wXYBQ1Yrbf7r5HmeCMd1hAkNfpBMUceU2sIaP4SEmNNWYcZ+iPUxmn5+t9uWG5ruzp0O9nxt0Dhn+Ey0xRhZ2kVkJVKXifY5RWImQjdTErb3pEkyO24DhIa4uaaGxRE40tamJLiwE+AgwAhdKHYw1lbmRzdHJlYW0NZW5kb2JqDTEwIDAgb2JqDTw8L0JpdHNQZXJDb21wb25lbnQgOC9Db2xvclNwYWNlIDEyIDAgUi9GaWx0ZXJbL0FTQ0lJODVEZWNvZGUvRmxhdGVEZWNvZGVdL0hlaWdodCAxOC9MZW5ndGggNTkvV2lkdGggMTg+PnN0cmVhbQ0KODtYcCwqPE5VVVsyQykoLldDcUUpWzlecDpfaFJvITJnX1hgaFM4Kl0uaF9DPzJ0T0UhLTlAMkY4fj4NZW5kc3RyZWFtDWVuZG9iag0xMSAwIG9iag08PC9MYXN0TW9kaWZpZWQoRDoyMDI2MTEyMTE0NTAxNiswOScwMCcpL1ByaXZhdGUgMTMgMCBSPj4NZW5kb2JqDTEzIDAgb2JqDTw8L0FJTWV0YURhdGEgMTQgMCBSL0FJUHJpdmF0ZURhdGExIDE1IDAgUi9BSVByaXZhdGVEYXRhMiAxNiAwIFIvQ29udGFpbmVyVmVyc2lvbiAxMi9DcmVhdG9yVmVyc2lvbiAzMS9OdW1CbG9jayAyL1JvdW5kdHJpcFN0cmVhbVR5cGUgMi9Sb3VuZHRyaXBWZXJzaW9uIDI0Pj4NZW5kb2JqDTE0IDAgb2JqDTw8L0xlbmd0aCAxNDI1Pj5zdHJlYW0NCiUhUFMtQWRvYmUtMy4wIA0lJUNyZWF0b3I6IEFkb2JlIElsbHVzdHJhdG9yKFIpIDI0LjANJSVBSThfQ3JlYXRvclZlcnNpb246IDMxLjAuMA0lJUZvcjogKG1hc2FoaXJvIHRha2FubykgKCkNJSVUaXRsZTogKHN0cm9rZS1ncmFkaWVudC5haSkNJSVDcmVhdGlvbkRhdGU6IDIwMjYvMTEvMjEgMTQ6NTANJSVDYW52YXNzaXplOiAxNjM4Mw0lJUJvdW5kaW5nQm94OiAyMTIgLTQ3MiAzMzcgLTM0Nw0lJUhpUmVzQm91bmRpbmdCb3g6IDIxMi4wMDAyNTI4NTcxMzkgLTQ3MS45OTk3NDcxNDI4NjEgMzM3IC0zNDcuMDAwMjUyODI2NTkxDSUlRG9jdW1lbnRQcm9jZXNzQ29sb3JzOiBDeWFuIE1hZ2VudGEgWWVsbG93DSVBSTVfRmlsZUZvcm1hdCAxNC4wDSVBSTEyX0J1aWxkTnVtYmVyOiA4NA0lQUkzX0NvbG9yVXNhZ2U6IENvbG9yDSVBSTdfSW1hZ2VTZXR0aW5nczogMA0lJUNNWUtQcm9jZXNzQ29sb3I6IDEgMSAxIDEgKFvjg6zjgrjjgrnjg4jjg6zjg7zjgrfjg6fjg7NdKQ0lQUkzX0Nyb3BtYXJrczogMjA5IC00NzUgMzQwIC0zNDQNJUFJM19UZW1wbGF0ZUJveDogMjk4LjUgLTQyMS41IDI5OC41IC00MjEuNQ0lQUkzX1RpbGVCb3g6IC01IC03ODkuNSA1NTQgLTYuNQ0lQUkzX0RvY3VtZW50UHJldmlldzogTm9uZQ0lQUk1X0FydFNpemU6IDE0NDAwIDE0NDAwDSVBSTVfUnVsZXJVbml0czogMQ0lQUkyNF9MYXJnZUNhbnZhc1NjYWxlOiAxDSVBSTlfQ29sb3JNb2RlbDogMg0lQUk1X0FydEZsYWdzOiAwIDAgMCAxIDAgMCAxIDAgMA0lQUk1X1RhcmdldFJlc29sdXRpb246IDgwMA0lQUk1X051bUxheWVyczogMQ0lQUkxN19CZWdpbl9Db250ZW50X2lmX3ZlcnNpb25fZ3Q6MjQgNA0lQUkxMF9PcGVuVG9WaWU6IDI5IC0zMjQgNS4zODEyMTU0Njk2MTMyNiAwIDgxNTAuMjAzMjg1NDIwOTUgODE3Ny4xNDg4NzA2MzY1NSAyMTg2IDk5NSAxOCAwIDAg",
            "NzggMTEzIDAgMCAwIDEgMSAwIDEgMSAwIDANJUFJMTdfQWx0ZXJuYXRlX0NvbnRlbnQNJUFJOV9PcGVuVG9WaWV3OiAyOSAtMzI0IDUuMzgxMjE1NDY5NjEzMjYgMjE4NiA5OTUgMTggMCAwIDc4IDExMyAwIDAgMCAxIDEgMCAxIDEgMCAwDSVBSTE3X0VuZF9WZXJzaW9uZWRfQ29udGVudA0lQUk1X09wZW5WaWV3TGF5ZXJzOiA3DSVBSTE3X0JlZ2luX0NvbnRlbnRfaWZfdmVyc2lvbl9ndDoyNCA0DSVBSTE3X0FsdGVybmF0ZV9Db250ZW50DSVBSTE3X0VuZF9WZXJzaW9uZWRfQ29udGVudA0lJVBhZ2VPcmlnaW46MCAwDSVBSTdfR3JpZFNldHRpbmdzOiA3MiA4IDcyIDggMCAwIDAuODAwMDAwMDExOTIwOTI5IDAuODAwMDAwMDExOTIwOTI5IDAuODAwMDAwMDExOTIwOTI5IDAuODk5OTk5OTc2MTU4MTQyIDAuODk5OTk5OTc2MTU4MTQyIDAuODk5OTk5OTc2MTU4MTQyDSVBSTlfRmxhdHRlbjogMQ0lQUkxMl9DTVNldHRpbmdzOiAwMC5NUw0lJUVuZENvbW1lbnRzDQ1lbmRzdHJlYW0NZW5kb2JqDTE1IDAgb2JqDTw8L0xlbmd0aCA2NTUzNj4+c3RyZWFtDQolQUkyNF9aU3RhbmRhcmRfRGF0YSi1L/0AWFT4AGr54C9I4FIbN6o0MMwwwwyVqpKYmZnstv+fgNFPiH+BUvNj2nXrdx+9xccaLX3e06gYpfN1LFFQslVQJEBrEqGV1bfqPI0Xw9g8McIL3wLUAv8CH/fg3hwfHeIkKcaEpRQjgprg3hzfCUWePce9Xw/3EHTBoJc7vxbpJZMDHZ4wg1vjAV2kl2iTpBiNSdIPOykGZJIUM1GLSVKMh+8N3qSXPOYX7CSffCnSS0ZMkmJITEkxdVLMJCmGM0mKqRPFTEoxVJ0wc1KM9JLuYQY0Gv9hE5nv7uEOKdobffERN7fQmSOClowVIuLo3HXmFr1DZD5ixxfOBje3RCJR2DX2eLb2aS4SiS4ZFCkztAzVuPcLdrROy06qui56SXrNKS2K9bT6MrfIVBWdOZlrYM61RKoSOHfYD2c+zZW4LrBIhPvVHHcQy1TzAotE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N3JcSS5IdSZlMGQmRTwuJylmRVJyL2wrKlcsKXFeRCphaTU8dXVMWC43Zy8+JFhLcmNZcDBuK1hsX25VKk8oCmxbJDZObitaX05xMF1zN2hzXWBYWDFuWjgmOTRhXH4+DWVuZHN0cmVhbQ1lbmRvYmoNNyAwIG9iag08PC9CQm94Wy0xMTUuODk2IDY0LjEyMjEgMTE1Ljg5NSAtNjQuMTIyMV0vTGVuZ3RoIDE0MDMvTWF0cml4WzEuMCAwLjAgMC4wIDEuMCAwLjAgMC4wXS9SZXNvdXJjZXM8PC9FeHRHU3RhdGU8PC9HUzAgMTggMCBSPj4vRm9udDw8L0MyXzAgNSAwIFIvVFQwIDYgMCBSPj4vUHJvY1NldFsvUERGL1RleHRdPj4vU3VidHlwZS9Gb3JtPj5zdHJlYW0NCkJUCjAgMCAwIDEgawovUGVyY2VwdHVhbCByaQovR1MwIGdzCi9DMl8wIDEgVGYKMCBUYyAwIFR3IDAgVHMgMTAwIFR6IDAgVHIgMTIgMCAwIC0xMiAtMTE1Ljg5NTUgLTU0LjA3ODEgVG0KPDAwMDAwMDAwMDAwMD5UagovVFQwIDEgVGYKKCBQREYgKVRqCi9DMl8wIDEgVGYKPDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA+VGoKL1RUMCAxIFRmCiggKVRqCjAgLTEuMiBUZApbKEEpMTIgKGRvYmVcMjU2IEkpLTEwLjEgKGxsdXN0cik1IChhKTQgKHQpNiAob3JcMjU2ICldVEoKL0MyXzAgMSBUZgo8MDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMD5UagovVFQwIDEgVGYKKCApVGoKL0MyXzAgMSBUZgowIC0xLjIgVGQKPDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwPlRqCi9UVDAgMSBUZgooIClUagovQzJfMCAxIFRmCjAgLTEuMiBUZAo8MDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA+VGoKL1RUMCAxIFRmClsoQSkxMiAoZG9iZSBJKS0xMCAobGx1c3RyKTUgKGEpMy45ICh0KTYuMSAob3IgKV1USgovQzJfMCAxIFRmCjwwMDAwMDAwMD5UagovVFQwIDEgVGYKKFBERiApVGoKL0MyXzAgMSBUZgo8MDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMD5UagovVFQwIDEgVGYKKCApVGoKL0MyXzAgMSBUZgowIC0xLjIgVGQKPDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMD5UagovVFQwIDEgVGYKKCApVGoKL0MyXzAgMSBUZgowIC0xLjIgVGQKPDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMD5UagovVFQwIDEgVGYKWyhBKTEyIChkb2JlIEkpLTEwIChsbHVzdHIpNSAoYSkzLjkgKHQpNi4xIChvciApXVRKCi9DMl8wIDEgVGYKPDAwMDAwMDAwMDAwMDAwMDAwMDAwPlRqCi9UVDAgMSBUZgooIClUagovQzJfMCAxIFRmCjAgLTEuMiBUZAo8MDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMD5UagovVFQwIDEgVGYKKCApVGoKL0MyXzAgMSBUZgowIC0xLjIgVGQKPDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA+VGoKL1RUMCAxIFRmClsoSSktMTAgKGxsdXN0cik0LjkgKGEpNC4xICh0KTYgKG9yICldVEoKL0MyXzAgMSBUZgo8MDAwMDAwMDAwMDAwMDAwMDAwMDA+VGoKL1RUMCAxIFRmCiggKVRqCi9DMl8wIDEgVGYKMCAtMS4yIFRkCjwwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDA+VGoKRVQKDWVuZHN0cmVhbQ1lbmRvYmoNNSAwIG9iag08PC9CYXNlRm9udC9KWlhOREErTXlyaWFkUHJvLVJlZ3VsYXIvRGVzY2VuZGFudEZvbnRzIDE5IDAgUi9FbmNvZGluZy9JZGVudGl0eS1IL1N1YnR5cGUvVHlwZTAvVG9Vbmljb2RlIDIwIDAgUi9UeXBlL0ZvbnQ+Pg1lbmRvYmoNNiAwIG9iag08PC9CYXNlRm9udC9KWlhOREErTXlyaWFkUHJvLVJlZ3VsYXIvRW5jb2RpbmcvV2luQW5zaUVuY29kaW5nL0ZpcnN0Q2hhciAzMi9Gb250RGVzY3JpcHRvciAyMSAwIFIvTGFzdENoYXIgMTc0L1N1YnR5cGUvVHJ1ZVR5cGUvVHlwZS9Gb250L1dpZHRoc1syMTIgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDYxMiAwIDAgNjY2IDAgNDg3IDAgMCAyMzkgMCAwIDAgMCAwIDAgNTMyIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAg",
            "NDgyIDU2OSAwIDU2NCA1MDEgMCAwIDAgMCAwIDAgMjM2IDAgMCA1NDkgMCAwIDMyNyAzOTYgMzMxIDU1MSAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgMCAwIDAgNDE5XT4+DWVuZG9iag0yMSAwIG9iag08PC9Bc2NlbnQgOTUyL0NhcEhlaWdodCA2NzQvRGVzY2VudCAtMjUwL0ZsYWdzIDMyL0ZvbnRCQm94Wy0xNTcgLTI1MCAxMTI2IDk1Ml0vRm9udEZhbWlseShNeXJpYWQgUHJvKS9Gb250RmlsZTIgMjIgMCBSL0ZvbnROYW1lL0paWE5EQStNeXJpYWRQcm8tUmVndWxhci9Gb250U3RyZXRjaC9Ob3JtYWwvRm9udFdlaWdodCA0MDAvSXRhbGljQW5nbGUgMC9TdGVtViA4OC9UeXBlL0ZvbnREZXNjcmlwdG9yL1hIZWlnaHQgNDg0Pj4NZW5kb2JqDTIyIDAgb2JqDTw8L0ZpbHRlci9GbGF0ZURlY29kZS9MZW5ndGggNjI2Ny9MZW5ndGgxIDEyMjM1Pj5zdHJlYW0NCkiJfJYHdBbVFoX3PvcnQEIJIVQhTH4CiDwFKaH5fFIfT3ov0kIKARIIEJrSpSV0LAgiVUGKEVS6CEiv0jvkT6GXUCdAIHlnArh4a7HezDozc++5s2bPnbu/MyCAvBgNgxbNW1es3KXhjPnac0Gje2h0SMzg6FErAH4A5PKEDo61sD6XB/AuoXl3REzPaK+O86pr+0MdM7Bn1LCI6wmL2wM+azRKRIaHhBVJDR0KFD6k44MjtSPnDq/aQBGj7aDI6NihDbuNWKbtciqiQ1S/0JDYSv2DNDUK8EqIDhkak6Of1AUqvKvjrb4h0eGnkht30nZTwFUipt/A2KwNqAxU7OXkYwaEx9i1d+7Sdpy268LFZM5ADrgkXjpoT+MXZ3ZFZboB8clhjMuIuCzIRM07c5G91alftwE+AjJccifLCzDpFAtc4OSYKoudp8F5B3l5j7+29ErywIuPtZ0b6dl5Z3t1/v+bgUuVeiEncundveGDPPpd8iE/fFEAfiiozyiEwiiC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qQrKD6lKIbl/+kq1nb0bfnefs3t9Xzj5HjxD2YlZ8uQd5CBLBE/4XFMMEYRajvoVYAC4LW/iDWVuZHN0cmVhbQ1lbmRvYmoNMjMgMCBvYmoNPDwvQmFzZUZvbnQvSlpYTkRBK015cmlhZFByby1SZWd1bGFyL0NJRFN5c3RlbUluZm8gMjQgMCBSL0NJRFRvR0lETWFwL0lkZW50aXR5L0RXIDEwMDAvRm9udERlc2NyaXB0b3IgMjUgMCBSL1N1YnR5cGUvQ0lERm9udFR5cGUyL1R5cGUvRm9udC9XWzBbNTAwXV0+Pg1lbmRvYmoNMjQgMCBvYmoNPDwvT3JkZXJpbmcoSWRlbnRpdHkpL1JlZ2lzdHJ5KEFkb2JlKS9TdXBwbGVtZW50IDA+Pg1lbmRvYmoNMjUgMCBvYmoNPDwvQXNjZW50IDk1Mi9DSURTZXQgMjYgMCBSL0NhcEhlaWdodCA2NzQvRGVzY2VudCAtMjUwL0ZsYWdzIDQvRm9udEJCb3hbLTE1NyAtMjUwIDExMjYgOTUyXS9Gb250RmFtaWx5KE15cmlhZCBQcm8pL0ZvbnRGaWxlMiAyNyAwIFIvRm9udE5hbWUvSlpYTkRBK015cmlhZFByby1SZWd1bGFyL0ZvbnRTdHJldGNoL05vcm1hbC9Gb250V2VpZ2h0IDQwMC9JdGFsaWNBbmdsZSAwL1N0ZW1WIDg4L1R5cGUvRm9udERlc2NyaXB0b3IvWEhlaWdodCA0ODQ+Pg1lbmRvYmoNMjYgMCBvYmoNPDwvRmlsdGVyL0ZsYXRlRGVjb2RlL0xlbmd0aCAxMT4+c3RyZWFtDQpIiWoACDAAAIEAgQ1lbmRzdHJlYW0NZW5kb2JqDTI3IDAgb2JqDTw8L0ZpbHRlci9GbGF0ZURlY29kZS9MZW5ndGggMjgwMy9MZW5ndGgxIDcxNzE+PnN0cmVhbQ0KSInsV3twVFcZ/865m11CAuQBJGTzuJtnk92QJUl5haZJ2CSliRAIxV06FJZkSUIDiZuIoFMBpdbZQcWOgwOt7bSVcYxUl5kOBmVKaNXRmY6ofVBndEoRKROptDVtbCm5/r5zb9J1G5zxP8fJ3vnd7zuv75z7Pc+SIKIkOkAaede1V1RuaTryBHrOAts69gzq9NNZl4hENtr5O/q7dtkDjy9DuxYY6Ordt6Ps0f4s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MDk0NCAwMDAwMCBuDQowMDAwMTAxMDgyIDAwMDAwIG4NCjAwMDAwOTkzMTggMDAwMDAgbg0KMDAwMDAyMjU3OCAwMDAwMCBuDQowMDAwMDIyOTM4IDAwMDAwIG4NCjAwMDAwMjMxNjEgMDAwMDAgbg0KMDAwMDAyMzM2MCAwMDAwMCBuDQowMDAwMDk4NzU3IDAwMDAwIG4NCjAwMDAwMjM0MzQgMDAwMDAgbg0KMDAwMDAyMzYwOCAwMDAwMCBuDQowMDAwMDI1MDg0IDAwMDAwIG4NCjAwMDAwOTA2NzIgMDAwMDAgbg0KMDAwMDA5ODgwNSAwMDAwMCBuDQowMDAwMTExOTY3IDAwMDAwIG4NCjAwMDAxMDgxNzUgMDAwMDAgbg0KMDAwMDEwODIwMCAwMDAwMCBuDQowMDAwMTAxNTY0IDAwMDAwIG4NCjAwMDAxMDE4MjQgMDAwMDAgbg0KMDAwMDEwODQ5MiAwMDAwMCBuDQowMDAwMTA4NjYwIDAwMDAwIG4NCjAwMDAxMDg3MjkgMDAwMDAgbg0KMDAwMDEwOTAwMiAwMDAwMCBuDQowMDAwMTA5MDgxIDAwMDAwIG4NCjAwMDAxMTIwODAgMDAwMDAgbg0KdHJhaWxlcg08PC9TaXplIDI5L1Jvb3QgMSAwIFIvSW5mbyAyOCAwIFIvSURbPDIzQkE1QzZCRDU1QTQ0OTc4RjVDNUI0MkQ0Q0JFMDcwPjxGRUIyNzNGRjI3Qzc0MzAxQUVCODUyRjZEMjBFNkUwNj5dPj4Nc3RhcnR4cmVmDTExMjI3NA0lJUVPRg0="
        ];
    }

    /* EMBEDDED-FILE-END */

})();
