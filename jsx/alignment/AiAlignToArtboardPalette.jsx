#target illustrator
#targetengine "AiAlignToArtboard"
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択したオブジェクトを、アートボードを対象に整列する常駐パレットです。
3×3のボタンで8方向へ寄せ、押すたびにガイド・アートボードのエッジ・裁ち落としへと寄せ先が進みます。中央揃えはマージンの内側の中央へ寄せ、マージンや分割のガイドも引けます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAlignToArtboardPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n42952a7adcb6

### Overview

A persistent palette that aligns the selected objects to the artboard.
A 3x3 grid of buttons moves the selection in eight directions, stepping the destination outwards on each
press: the guide, the artboard edge, then the bleed. The centred buttons align to the centre inside the
margin, and it also draws margin and division guides.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAlignToArtboardPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiAlignToArtboardPalette";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-08-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-03";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiAlignToArtboardPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiAlignToArtboardPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n42952a7adcb6"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    /* 常駐エンジンに残すパレット参照（GC回避と多重起動防止を兼ねる）
       var の初期化は再実行のたびに走るため、既存の参照を消さないよう $.global から引き継ぐ
       The palette reference lives in the persistent engine; it is carried over from $.global so a
       re-run does not wipe it before closeExistingPalette() can close the old window */
    var paletteWindow = $.global.__aiAlignToArtboardWindow || null;

    (function() {

        // =========================================
        // ユーザー設定 / User settings
        // =========================================
        /* チェックボックスの初期状態。［プレビュー境界］［字形の境界に整列］はパレットを開くときに
           環境設定の現在値で上書きするので、この値は読めなかったときの控えになる
           （整列のあいだだけワーカーが環境設定へ書き込み、終わったら元の設定に戻す）
           Initial checkbox states. Preview Bounds and Align to Glyph Bounds are overwritten with the current
           preferences when the palette opens, so these are only the fallback; the worker applies them for the
           duration of an align and restores the previous preferences afterwards */
        var DEFAULT_PREVIEW_BOUNDS       = false; /* プレビュー境界 / preview bounds */
        var DEFAULT_GLYPH_BOUNDS         = true;  /* 字形の境界に整列 / align to glyph bounds */
        var DEFAULT_CHANGE_JUSTIFICATION = true;  /* 行揃えを変更 / change justification */
        var DEFAULT_OPTICAL_ADJUST       = true;  /* 見た目の調整 / optical adjustment */
        var DEFAULT_LINK_MARGINS         = true;  /* マージンの4値を連動させる / keep the four margins in sync */
        var DEFAULT_ALIGN_TO_BLEED       = true;  /* 裁ち落としに整列 / align to the bleed */
        var DEFAULT_ALIGN_PER_ARTBOARD   = false; /* アートボードごとに整列 / align per artboard */
        /* 裁ち落としの量（mm）。パレットには数値欄を置かないので、変えたいときはここを書き換える
           裁ち落としは印刷の値なので、定規の単位に追従させず mm で扱う
           The bleed in millimetres; the palette has no field for it, so change it here.
           Being a print value it stays in millimetres instead of following the ruler */
        var BLEED_MM          = 3;
        /* ［見た目の調整］の強さ（%）。100で字面の中心をぴったり揃える（約物の空きまで埋めると寄せすぎに見える）
           Strength of Optical Adjustment (%); 100 centers the glyphs exactly, which looks overdone with punctuation */
        var OPTICAL_KERNING_STRENGTH = 35;
        var BLEED_UNIT_POINTS = 72.0 / 25.4;
        var DEFAULT_SHOW_GUIDE           = false; /* ガイドを追加 / add the margin guide */
        /* 分割ガイドの初期状態。行・列はマージンの内側をいくつに分けるかで、1 のときはガイドを引かない
           行間・列間・伸張は定規の単位で扱う（伸張はアートボードの外へ伸ばす距離）
           Division guides: rows and columns split the area inside the margin, so 1 draws no guide;
           the gutters and the extension are in ruler units, the extension reaching outside the artboard */
        var DEFAULT_DIVIDE_MODE          = "none";
        var DEFAULT_ARTBOARD_EDGE        = false; /* アートボードのエッジにガイドを引く / draw guides on the artboard edges */
        /* 伸張の初期値は単位ごとに決まるので、ここには持たない（UNIT_DEFAULTS の defaultExtension を使う）
           The extension's default comes from the unit, so it is not listed here */
        var DEFAULT_DIVIDE_VALUES        = { rows: 1, columns: 2, rowGutter: 0, columnGutter: 0 };
        /* ［十字］の行数・列数（縦横とも2等分＝中央に十字のガイドが1本ずつ）
           "Cross" halves both directions, leaving one guide on each axis through the centre */
        var CROSS_DIVIDE_COUNTS          = { rows: 2, columns: 2 };
        /* 行数・列数の上限（大きな値でガイドを何千本も作らないための歯止め）/ Cap on the counts, so a large number cannot flood the document */
        var DIVIDE_COUNT_MAX             = 100;

        /* マージンのガイドを作るレイヤー名（他のガイド系スクリプトと共通）
           Layer that receives the margin guide, shared with the other guide scripts */
        var GUIDE_LAYER_NAME = "_guide";
        /* このスクリプトが作るガイドの名前。張り替えるときの目印にする
           The name given to the guide, used to find and replace it */
        var GUIDE_NAME = "AiAlignToArtboard-margin";
        /* 分割ガイドの名前。マージンのガイドと分けて、片方だけ張り替えられるようにする
           The name given to the division guides, kept apart from the margin guide */
        var DIVIDE_GUIDE_NAME = "AiAlignToArtboard-divide";

        /* 常駐エンジン（$.global）に控える値のキー
           ガイドの設定はパレットを開き直しても引き継ぐ（ドキュメントに残したガイドと入力欄の値が食い違わないように。保存は settingsStore）
           Keys kept on $.global: the guide settings survive a close-and-reopen, so the guides left in the
           document keep matching the fields */
        var LEGACY_SETTINGS_KEY = "__aiAlignToArtboardSettings"; /* 旧版の置き場所（読み継ぎだけに使う）/ old location, read only for migration */
        /* メインエンジンへ送り込んだワーカー定義の刻印を控えるキー / Key holding the stamp of the loaded worker source */
        var WORKER_STAMP_KEY = "__aiAlignToArtboardWorkerStamp";
        /* 閉じたときのパレットの位置を控えるキー / Key holding the palette's location when it was closed */
        var WINDOW_LOCATION_KEY = "__aiAlignToArtboardWindowLocation";
        /* 控えた位置を使う条件。左上がこのぶん画面の内側にあること（掴めない位置に出さないため）
           The stored location is used only when its top-left sits at least this far inside a screen */
        var WINDOW_ONSCREEN_MARGIN = 60;

        /* メインエンジンからの応答を待つ秒数 / seconds to wait for the main engine */
        var WORKER_TIMEOUT = 10;
        /* 整列先が［アートボード］かを判定するための仮移動量（pt）
           整列してもオブジェクトが動かなかったときだけ、このぶん内側へずらして整列し直し、戻ってくるかを見る
           Probe distance (pt): used only when an align moved nothing, to tell "already aligned" from a wrong target */
        var ALIGN_PROBE_PT = 4;
        /* ガイドが水平・垂直かを判定する許容値（pt）。これを超える幅・高さがあれば長方形とみなす
           Tolerance (pt) for calling a guide horizontal or vertical; anything thicker counts as a rectangle */
        var GUIDE_ORIENTATION_TOLERANCE = 0.01;
        /* 移動先がここまで近ければ「すでにその位置にいる」とみなす許容値（pt）
           手で吸着させたオブジェクトの辺とガイドは 1e-12 ほどずれることがあり、そのままでは行き先に選ばれて動かなくなる
           Tolerance (pt) for "already there"; a hand-snapped edge and its guide can differ by ~1e-12,
           which would otherwise be picked as the destination and move nothing */
        var MOVE_MIN_DELTA_PT = 0.001;
        /* 選択を取り直す最短間隔（mouseover は何度も発生するため間引く）/ Throttle for the mouseover refresh */
        var SELECTION_POLL_INTERVAL_MS = 400;

        // =========================================
        // レイアウト / Layout
        // =========================================

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

        var ICON_SIZE       = 30;   /* 整列アイコン1個の大きさ（px）/ size of each align icon (px) */
        var RADIO_GAP       = 14;   /* ラジオボタンどうしの間隔 / gap between the radio buttons */
        var CROSS_GAP       = 2;    /* 移動ボタン（十字）どうしの間隔 / gap between the move buttons */
        var CENTER_ROW_TOP  = 8;    /* 十字と中央揃えボタンの間隔 / gap between the cross and the centre-align buttons */
        var EDGE_ROW_TOP    = 5;    /* 分割の欄と［アートボードのエッジ］の間隔 / gap above the artboard-edge checkbox */
        var FIELD_CHARS     = 3;    /* マージン入力欄の文字数 / width of the margin field */
        var LABEL_FIELD_SPACING = 4; /* 入力欄と単位ラベルの間隔（既定は広すぎる）/ gap between the field and its unit label */
        /* マージン欄を3×3に並べるときの1セルの幅（日英で文字数が違うので分ける）
           Width of one cell in the 3x3 margin grid; the labels differ in length by language */
        var MARGIN_CELL_WIDTH = { ja: 70, en: 84 };
        /* 分割ガイドの項目名の幅。右そろえにして数値欄の頭をそろえる（日英で文字数が違うので分ける）
           Width of the division labels; right-aligned so the fields line up, and the labels differ by language */
        var DIVIDE_LABEL_WIDTH        = { ja: 48, en: 72 };
        var DIVIDE_GUTTER_LABEL_WIDTH = { ja: 48, en: 52 };
        var STATUS_WIDTH    = 260;  /* 状況表示の幅（中身でパレット幅が変わらないよう固定）/ fixed width of the status line */
        var OPTION_SPACING  = 4;    /* オプションのチェックボックスどうしの間隔 / gap between the option checkboxes */

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

        /* 入力された単位を欄の単位へ換算するための、1単位あたりのポイント数（値は UNITS 表と同じ。キーは小文字）。
           「p」は「1p6」（1パイカ6ポイント）の形にも使う
           Points per unit for converting typed units into the field's unit (same values as the UNITS table; lowercase keys) */
        var STEPPER_POINTS_PER_UNIT = {
            "in": 72, "inch": 72, "mm": 72 / 25.4, "cm": 72 / 2.54, "m": 72 / 25.4 * 1000,
            "pt": 1, "px": 1, "p": 12, "pc": 12, "pica": 12,
            "q": 72 / 25.4 * 0.25, "h": 72 / 25.4 * 0.25, "ft": 72 * 12, "yd": 72 * 36
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
                if (!hasOperator && value === parseFloat(numberInput.text)) return;
                numberInput.text = formatStepperNumber(value) + (fieldUnit || "");
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
         * 数値の後ろの単位は欄の単位へ換算する（mm の欄に「1in」→ 25.4、「1p6」は1パイカ6ポイント）。単位のない数値は欄の単位とみなす。
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
            var fieldPointsPerUnit = STEPPER_POINTS_PER_UNIT[fieldUnitKey];
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
                var unitMatch = /^([A-Za-z]+|%|°)/.exec(source.substring(position));
                if (!unitMatch) return value; /* 単位なしは欄の単位 / no unit means the field's unit */
                position += unitMatch[0].length;
                var unitKey = unitMatch[0].toLowerCase();
                if (unitKey === fieldUnitKey) return value;
                var pointsPerUnit = STEPPER_POINTS_PER_UNIT[unitKey];
                if (pointsPerUnit === undefined || fieldPointsPerUnit === undefined) return NaN; /* 知らない単位・単位のない欄 / unknown unit or unitless field */
                var points = value * pointsPerUnit;
                /* 「1p6」＝1パイカ6ポイント / pica-point notation */
                if (unitKey === "p") {
                    var pointMatch = /^(\d+\.?\d*|\.\d+)/.exec(source.substring(position));
                    if (pointMatch) {
                        position += pointMatch[0].length;
                        points += parseFloat(pointMatch[0]);
                    }
                }
                return points / fieldPointsPerUnit;
            }

            var result = readSum();
            if (position !== source.length || !isFinite(result)) return NaN; /* 読み残しがあれば式として不正 / leftovers mean a malformed expression */
            return result;
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

        // =========================================
        // アイコンの寸法 / Icon metrics
        // =========================================
        /* すべて ICON_SIZE に対する比率で持ち、アイコンサイズを変えても形が崩れないようにする
           All ratios of ICON_SIZE so the glyphs keep their shape when the icon size changes */
        var ICON_BAR_THICKNESS  = 0.25;  /* オブジェクトを表すバーの太さ / thickness of the bars standing for objects */
        var ICON_BAR_GAP        = 0.07;  /* バー2本の間隔 / gap between the two bars */
        var ICON_BAR_LONG       = 0.55;  /* 長いほうのバーの長さ / length of the longer bar */
        var ICON_BAR_SHORT      = 0.35;  /* 短いほうのバーの長さ / length of the shorter bar */
        var ICON_RULE_INSET     = 0.12;  /* 基準線の端の余白 / inset at both ends of the reference rule */
        var ICON_RULE_OFFSET    = 0.17;  /* 端に置く基準線の位置 / position of the rule when it sits at an edge */
        var ICON_RULE_CLEARANCE = 0.05;  /* 端の基準線とバーのすき間（中央の基準線はバーの下を通す）/ gap between an edge rule and the bars (a center rule runs behind them) */
        /* アイコンに置くオブジェクトの大きさ。どのアイコンも同じ横長の長方形にする
           The object block: every icon uses the same landscape rectangle */
        var ICON_BLOCK_WIDTH    = 0.38;  /* オブジェクトの幅 / width of the object block */
        var ICON_BLOCK_HEIGHT   = 0.30;  /* オブジェクトの高さ / height of the object block */

        // =========================================
        // 整列コマンド / Align commands
        // =========================================
        /* 「オブジェクト > 整列」のメニューコマンド名と、マージンぶん内側へ動かす向き
           整列コマンドはアートボードの辺にぴったり寄せるので、そこから offset 方向へマージンぶん動かす
           Y は上が正のため、上揃えは -1（下へ）・下揃えは +1（上へ）になる
           axis は整列する軸で、整列先の判定に使う仮移動の向きを決める
           mode はその軸のどこに寄せるか（start＝左・上／center＝中央／end＝右・下）で、
           字形の境界での補正（btApplyGlyphCorrection）が目標位置を計算するのに使う
           justification は水平方向の整列に合わせる行揃え（垂直方向は行揃えを変えないので null）
           Menu command names under Object > Align, with the direction to move by the margin;
           Y grows upward, so top align moves -1 (down) and bottom align +1 (up).
           axis is the axis being aligned, which sets the direction of the probe used to check the align target.
           justification is the paragraph justification to match; vertical aligns leave it alone (null) */
        var ALIGN_COMMANDS = {
            horizontalLeft:   { command: "Horizontal Align Left",   axis: "x", mode: "start",  offsetX:  1, offsetY:  0, justification: "LEFT" },
            horizontalCenter: { command: "Horizontal Align Center",  axis: "x", mode: "center", offsetX:  0, offsetY:  0, justification: "CENTER" },
            horizontalRight:  { command: "Horizontal Align Right",   axis: "x", mode: "end",    offsetX: -1, offsetY:  0, justification: "RIGHT" },
            verticalTop:      { command: "Vertical Align Top",       axis: "y", mode: "start",  offsetX:  0, offsetY: -1, justification: null },
            verticalCenter:   { command: "Vertical Align Center",    axis: "y", mode: "center", offsetX:  0, offsetY:  0, justification: null },
            verticalBottom:   { command: "Vertical Align Bottom",    axis: "y", mode: "end",    offsetX:  0, offsetY:  1, justification: null }
        };

        // =========================================
        // 定規の単位 / Ruler units
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

        /* 単位ごとのマージン欄・伸張欄の初期値。換算値ではなく、その単位で扱いやすい丸めた数
           Initial margin and extension per unit: round numbers that read well in that unit */
        var UNIT_DEFAULTS = [
            { defaultMargin: 0.25,  defaultExtension: 0.5 },   /* 0 in */
            { defaultMargin: 5,     defaultExtension: 10 },    /* 1 mm */
            { defaultMargin: 20,    defaultExtension: 20 },    /* 2 pt */
            { defaultMargin: 1.5,   defaultExtension: 2 },     /* 3 pica */
            { defaultMargin: 0.5,   defaultExtension: 1 },     /* 4 cm */
            { defaultMargin: 20,    defaultExtension: 40 },    /* 5 Q/H */
            { defaultMargin: 20,    defaultExtension: 20 },    /* 6 px */
            { defaultMargin: 0.02,  defaultExtension: 0.04 },  /* 7 ft/in */
            { defaultMargin: 0.005, defaultExtension: 0.01 },  /* 8 m */
            { defaultMargin: 0.006, defaultExtension: 0.012 }, /* 9 yd */
            { defaultMargin: 0.02,  defaultExtension: 0.04 }   /* 10 ft */
        ];

        /**
         * 表示用の単位情報に、その単位の初期値を足して返す
         * @returns {{label: string, points: number, defaultMargin: number, defaultExtension: number}} 単位の情報
         */
        function getUnitInfoWithDefaults() {
            var unit = getUnitInfo();
            var defaults = UNIT_DEFAULTS[unit.code] || UNIT_DEFAULTS[2];
            return {
                label: unit.label,
                points: unit.pointsPerUnit,
                defaultMargin: defaults.defaultMargin,
                defaultExtension: defaults.defaultExtension
            };
        }
        /* 単位が取れないときの既定 / Fallback when the ruler unit cannot be read */
        var FALLBACK_UNIT_INFO = { label: "pt", points: 1.0, defaultMargin: 20, defaultExtension: 20 };

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
         * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
         * @param {string|Object} labelRef - getLabel と同じ
         * @param {Object|Array} [placeholderValues] - getLabel と同じ
         * @returns {string} コロン付きの文言
         */
        function labelText(labelRef, placeholderValues) {
            return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
        }

        /**
         * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
         * @param {string|Object} labelRef - getLabel と同じ
         * @param {string|number} value - コロンのあとに続ける値
         * @returns {string} 項目名と値をつないだ文字列
         */
        function labelValueText(labelRef, value) {
            return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
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

        // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)

        /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
        var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

        /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
        var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

        /* 修飾キーの別名 / Aliases accepted for the modifiers */
        var KEY_SHORTCUT_MODIFIER_ALIASES = {
            SHIFT: "SHIFT",
            ALT: "ALT", OPTION: "ALT", OPT: "ALT",
            CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
        };

        /**
         * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
         * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
         * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
         */
        function normalizeKeyShortcutSpec(keySpec) {
            var specParts = String(keySpec).split("+");
            var baseKey = specParts.pop().toUpperCase();
            var modifierFlags = {};
            for (var i = 0; i < specParts.length; i++) {
                var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
                if (modifierName) modifierFlags[modifierName] = true;
            }
            return buildKeyShortcutSpec(modifierFlags, baseKey);
        }

        /**
         * 修飾キーの状態とキー名から照合用の表記を組み立てる
         * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
         * @param {string} baseKey - 大文字のキー名
         * @returns {string} 照合用の表記
         */
        function buildKeyShortcutSpec(modifierFlags, baseKey) {
            var specText = "";
            for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
                if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
            }
            return specText + baseKey;
        }

        /**
         * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
         * @param {Object} keyEvent - keydown イベント
         * @returns {string} 照合用の表記。キー名が無いときは空文字
         */
        function readKeyShortcutSpec(keyEvent) {
            if (!keyEvent || !keyEvent.keyName) return "";
            var keyboardState = {};
            try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
            var modifierFlags = {
                SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
                ALT: !!(keyEvent.altKey || keyboardState.altKey),
                CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
            };
            return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
        }

        /**
         * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
         * @param {Object} control - コントロール
         * @returns {boolean} 押せるなら true
         */
        function isKeyShortcutControlUsable(control) {
            for (var node = control; node; node = node.parent) {
                if (node.enabled === false || node.visible === false) return false;
            }
            return true;
        }

        /**
         * キーを受けたコントロールが、文字を入力する欄か
         * @param {Object} focusedControl - イベントの発生元
         * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
         * @returns {boolean} 入力中としてショートカットを止めるなら true
         */
        function isKeyShortcutTypingTarget(focusedControl, numericFields) {
            if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
            for (var i = 0; i < numericFields.length; i++) {
                if (numericFields[i] === focusedControl) return false;
            }
            return true;
        }

        /**
         * コントロールをクリックしたときと同じ動作をする
         * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
         * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
         * @returns {void}
         */
        function pressKeyShortcutControl(control) {
            if (control.type === "radiobutton") {
                /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
                var siblings = control.parent ? control.parent.children : [];
                for (var i = 0; i < siblings.length; i++) {
                    if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
                }
                control.value = true;
            } else if (control.type === "checkbox") {
                control.value = !control.value;
            }
            if (typeof control.onClick === "function") {
                control.onClick.call(control);
            } else if (control.type === "button" && typeof control.notify === "function") {
                /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
                control.notify("onClick");
            }
        }

        /**
         * 1つのショートカットを実行する
         * @param {Object|Function} shortcutTarget - コントロール、または関数
         * @param {Object} keyEvent - keydown イベント
         * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
         */
        function runKeyShortcutTarget(shortcutTarget, keyEvent) {
            var targetControl = shortcutTarget;
            if (typeof shortcutTarget === "function") {
                var runResult = shortcutTarget(keyEvent);
                if (runResult === false || runResult === null) return false;
                if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
                targetControl = runResult;
            }
            /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
            if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
            return true;
        }

        /**
         * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
         * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
         * @returns {string} 表示用の表記（"Shift+R" など）
         */
        function formatKeyShortcutLabel(normalizedSpec) {
            var isMac = ($.os.indexOf("Mac") === 0);
            var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
            var specParts = normalizedSpec.split("+");
            var baseKey = specParts.pop();
            var labelText = "";
            for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
            if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
            return labelText + baseKey;
        }

        /**
         * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
         * @param {Object} control - コントロール
         * @param {string} normalizedSpec - 照合用の表記
         * @returns {void}
         */
        function appendKeyShortcutToTip(control, normalizedSpec) {
            var keyLabel = formatKeyShortcutLabel(normalizedSpec);
            var currentTip = control.helpTip ? String(control.helpTip) : "";
            if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
            var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
            control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
        }

        /**
         * ダイアログ・パレットに文字キーのショートカットを付ける
         * @param {Window} targetWindow - キーを受けるダイアログ・パレット
         * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
         * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
         * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
         */
        function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
            var shortcutSettings = shortcutOptions || {};
            var numericFields = shortcutSettings.numericFields || [];
            var bindingTable = {};

            for (var keySpec in shortcutMap) {
                if (!shortcutMap.hasOwnProperty(keySpec)) continue;
                var mapEntry = shortcutMap[keySpec];
                if (!mapEntry) continue;
                var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
                var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
                bindingTable[normalizedSpec] = {
                    target: isWrapped ? mapEntry.target : mapEntry,
                    inFields: !!(isWrapped && mapEntry.inFields)
                };
                var tipControl = bindingTable[normalizedSpec].target;
                if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                    appendKeyShortcutToTip(tipControl, normalizedSpec);
                }
            }

            /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
            targetWindow.addEventListener("keydown", function (keyEvent) {
                var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
                if (!binding) return;
                if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
                if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
                if (keyEvent.preventDefault) keyEvent.preventDefault();
                if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
            }, true);

            return bindingTable;
        }

        // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts

        /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
        var LABELS = {
            dialog: {
                title: { ja: "ガイドやアートボードのエッジに整列", en: "Align to Guides or Artboard Edges" }
            },
            panel: {
                guide:       { ja: "マージン", en: "Margin" },
                divide:      { ja: "分割とエッジのガイド", en: "Division & Edge Guides" },
                options:     { ja: "オプション", en: "Options" }
            },
            fieldLabel: {
                top:    { ja: "上", en: "Top" },
                bottom: { ja: "下", en: "Bottom" },
                left:   { ja: "左", en: "Left" },
                right:  { ja: "右", en: "Right" },
                rows:         { ja: "行数", en: "Rows" },
                columns:      { ja: "列数", en: "Columns" },
                rowGutter:    { ja: "行間", en: "Gutter" },
                columnGutter: { ja: "列間", en: "Gutter" },
                extension:    { ja: "伸張", en: "Extension" }
            },
            direction: {
                up:        { ja: "上", en: "Up" },
                left:      { ja: "左", en: "Left" },
                right:     { ja: "右", en: "Right" },
                down:      { ja: "下", en: "Down" },
                upLeft:    { ja: "左上", en: "Top left" },
                upRight:   { ja: "右上", en: "Top right" },
                downLeft:  { ja: "左下", en: "Bottom left" },
                downRight: { ja: "右下", en: "Bottom right" }
            },
            tooltip: {
                stepUp: {
                    ja: "値を増やす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                    en: "Increase (Shift-click to snap to 10s, Option-click by 0.1)"
                },
                stepDown: {
                    ja: "値を減らす（shift＋クリックで10の倍数へ、option＋クリックで0.1ずつ）",
                    en: "Decrease (Shift-click to snap to 10s, Option-click by 0.1)"
                },
                stepUpInteger:   { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
                stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
                alignCenterH:   { ja: "水平方向中央に整列", en: "Horizontal Align Center" },
                alignCenterV:   { ja: "垂直方向中央に整列", en: "Vertical Align Center" },
                alignCenterAll: { ja: "水平・垂直方向中央に整列", en: "Align Center on Both Axes" },
                panelGuide: {
                    ja: "整列の寄せ先になる余白。ガイドとしても引ける",
                    en: "The inset that alignment steps to; it can be drawn as a guide too"
                },
                panelDivide: {
                    ja: "マージンの内側を等分する位置と、アートボードのエッジにガイドを引く",
                    en: "Draw guides that split the area inside the margin, and guides on the artboard edges"
                },
                margin: { ja: "アートボードのエッジから空ける距離", en: "Distance to keep from the artboard edge" },
                linkMargins: {
                    ja: "上下左右のマージンを同じ値にする（どれかを変えると残りもそろえる）",
                    en: "Keep the four margins equal; changing one updates the rest"
                },
                alignToBleed: {
                    ja: "アートボードのエッジの次の寄せ先として、その外側の裁ち落とし（{0}mm）の位置を使う",
                    en: "Add the bleed ({0} mm) outside the artboard as the stop after the artboard edge"
                },
                perArtboard: {
                    ja: "選択したオブジェクトを、それぞれが乗っているアートボードに整列（OFFのときは1つのアートボードにまとめて整列）。ガイドもすべてのアートボードに描く",
                    en: "Align each object to the artboard it sits on (off: everything goes to a single artboard). Guides are drawn on every artboard too"
                },
                moveToEdge: {
                    ja: "その方向のガイド → アートボードのエッジ → 裁ち落とし の順に寄せる（↑↓←→キーでも実行）",
                    en: "Step to the guide in that direction, then the artboard edge, then the bleed (the arrow keys do the same)"
                },
                showGuide: {
                    ja: "マージンの位置に長方形のガイドを作る（「_guide」レイヤー、アクティブなアートボードに1つ）。パレットを閉じても残る。OFFのあいだは上下左右の欄をディムにする",
                    en: "Draw a rectangle guide at the margin (on the \"_guide\" layer, one on the active artboard). It stays when the palette closes, and the four fields are dimmed while this is off"
                },
                divideNone: {
                    ja: "分割のガイドを引かない",
                    en: "Draw no division guides"
                },
                divideCross: {
                    ja: "マージンの内側の中央に、縦横1本ずつの十字のガイドを引く",
                    en: "Draw a cross through the centre of the area inside the margin"
                },
                divideCustom: {
                    ja: "行数・列数のぶんだけ、マージンの内側を等分するガイドを引く",
                    en: "Split the area inside the margin into the number of rows and columns below"
                },
                divideRows: {
                    ja: "マージンの内側（マージンが0ならアートボード全体）を上下に何段に分けるか（1のときはガイドを引かない）",
                    en: "How many rows to split the area inside the margin into, or the whole artboard at margin 0 (1 draws no guide)"
                },
                divideColumns: {
                    ja: "マージンの内側（マージンが0ならアートボード全体）を左右に何列に分けるか（1のときはガイドを引かない）",
                    en: "How many columns to split the area inside the margin into, or the whole artboard at margin 0 (1 draws no guide)"
                },
                divideRowGutter: {
                    ja: "行と行のあいだに空ける距離（0より大きいと、段の上下2本のガイドを引く）",
                    en: "Space between two rows; above zero, each split gets a guide on both sides"
                },
                divideColumnGutter: {
                    ja: "列と列のあいだに空ける距離（0より大きいと、列の左右2本のガイドを引く）",
                    en: "Space between two columns; above zero, each split gets a guide on both sides"
                },
                divideExtension: {
                    ja: "分割のガイドをアートボードの外へ伸ばす距離",
                    en: "Distance the division guides reach outside the artboard"
                },
                artboardEdge: {
                    ja: "アートボードの上下左右4辺にガイドを引く（伸張のぶんだけ外へ伸ばす）",
                    en: "Draw guides on the four edges of the artboard, reaching outside by the extension"
                },
                optionGlyphBounds: { ja: "Option＋クリックで字形の境界に整列", en: "Option-click to align to glyph bounds" },
                optionArtboardCenter: {
                    ja: "Option＋クリックでアートボードの中央に整列（通常はマージンの内側の中央）",
                    en: "Option-click to centre on the artboard instead of inside the margin"
                },
                optionNoMargin: {
                    ja: "Option＋クリックでマージンなし・字形の境界に整列",
                    en: "Option-click: ignore the margin, align to glyph bounds"
                },
                previewBounds: {
                    ja: "整列でプレビュー境界（線幅・効果を含む）を使用",
                    en: "Use preview bounds (incl. stroke & effects) when aligning"
                },
                glyphBounds: {
                    ja: "ポイント文字・エリア内文字を字形の境界で整列",
                    en: "Align point & area type to glyph bounds"
                },
                changeJustification: {
                    ja: "水平方向の整列に合わせて、1行だけのテキスト1つの行揃えも変える",
                    en: "Match the justification of a lone single-line text object to the horizontal alignment"
                },
                opticalAdjust: {
                    ja: "中央揃えのポイント文字で、行末の「、」などの空きで寄って見える行を、行頭のカーニングで戻す（OFFなら行頭のカーニングを0に）",
                    en: "Kern the start of each line of centered point type so lines that look off because of punctuation such as 、 are pulled back (off resets the line-start kerning to 0)"
                }
            },
            radio: {
                divideNone:   { ja: "なし", en: "None" },
                divideCross:  { ja: "十字", en: "Cross" },
                divideCustom: { ja: "カスタム", en: "Custom" }
            },
            checkbox: {
                showGuide:     { ja: "ガイドを追加", en: "Add Guides" },
                artboardEdge:  { ja: "アートボードのエッジ", en: "Artboard Edges" },
                previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
                glyphBounds:   { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
                alignToBleed:  { ja: "裁ち落としに整列", en: "Align to Bleed" },
                perArtboard:   { ja: "アートボードごと", en: "Per Artboard" },
                changeJustification: { ja: "行揃えを変更", en: "Change Justification" },
                opticalAdjust:       { ja: "見た目の調整", en: "Optical Adjustment" }
            },
            status: {
                done:           { ja: "整列しました。", en: "Aligned." },
                moved:          { ja: "寄せました。", en: "Moved." },
                noBounds:       { ja: "境界を取得できません。", en: "Could not measure the selection." },
                doneJustified:  { ja: "整列し、行揃えを{0}に変更しました。", en: "Aligned; justification set to {0}." },
                movedJustified: { ja: "寄せて、行揃えを{0}に変更しました。", en: "Moved; justification set to {0}." },
                noDocument:     { ja: "ドキュメントが開かれていません。", en: "No document is open." },
                noSelection:    { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
                multipleLayers: { ja: "レイヤーをまたぐ選択は整列できません。", en: "Cannot align a selection spanning layers." },
                /* 状況表示は STATUS_WIDTH で切り詰められるため、全角20字ほどに収める
                   （切れても helpTip で全文を読める）
                   Keep it within the fixed status width; the full text is still available as a helpTip */
                alignTarget: {
                    ja: "整列先を［アートボード］にしてください。",
                    en: "Set Align To: Artboard."
                },
                noResponse:     { ja: "Illustrator から応答がありません。", en: "No response from Illustrator." },
                genericError:   { ja: "エラー：", en: "Error: " },
                /* doneJustified の {0} に入れる行揃えの名前 / Names substituted into doneJustified */
                justification: {
                    LEFT:   { ja: "左揃え",   en: "Left" },
                    CENTER: { ja: "中央揃え", en: "Center" },
                    RIGHT:  { ja: "右揃え",   en: "Right" }
                }
            }
        };

        // =========================================
        // 配色 / Colors
        // =========================================
        /* アイコンの配色（initIconColors() で UI 明暗から設定）/ Icon colors (set from the light/dark UI in initIconColors()) */
        var iconColor, iconBaseBg, iconHoverBg, iconBorderColor;

        /**
         * UI 明度（0..1）を取得する
         * @returns {number} 0〜1 にクランプした明度（取得失敗時は 0＝暗い側）
         */
        function getUIBrightness() {
            try {
                var brightness = app.preferences.getRealPreference("uiBrightness");
                if (brightness < 0) { brightness = 0; }
                if (brightness > 1) { brightness = 1; }
                return brightness;
            } catch (e) {
                return 0;
            }
        }

        /**
         * グレーの RGBA を作る
         * @param {number} brightness - 明度（0..1 にクランプ）
         * @returns {number[]} [r, g, b, a] の配列
         */
        function grayColor(brightness) {
            if (brightness < 0) { brightness = 0; }
            if (brightness > 1) { brightness = 1; }
            return [brightness, brightness, brightness, 1];
        }

        /**
         * UI の明暗に合わせてアイコン色とマウスオーバー時の背景色を決める
         * @returns {void}
         */
        function initIconColors() {
            var uiBrightness = getUIBrightness();
            var lightUI = !isDarkUI();
            iconColor = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1];
            /* 通常時の背景はパレットの地色に近いグレー。graphics.backgroundColor は iconbutton などで取得できず
               fillPath() が例外を投げ、再描画のたびにボタンが消えるため、必ず明示色で塗る
               Always paint an explicit gray; graphics.backgroundColor is unavailable on some controls and makes
               fillPath() throw, which blanks the button on every redraw */
            iconBaseBg  = lightUI ? grayColor(uiBrightness)        : [0.28, 0.28, 0.28, 1];
            /* マウスオーバー時の背景（ライトは少し暗く、ダークは少し明るく）/ Hover background (slightly darker in light, lighter in dark) */
            iconHoverBg = lightUI ? grayColor(uiBrightness - 0.10) : [0.38, 0.38, 0.38, 1];
            /* マウスオーバー時の枠線（ライトは薄いグレー、ダークは背景より明るいグレー）
               Hover border: light gray in light UI, gray brighter than the background in dark UI */
            iconBorderColor = lightUI ? [0.65, 0.65, 0.65, 1] : [0.45, 0.45, 0.45, 1];
        }

        // =========================================
        // 描画ヘルパー / Drawing helpers
        // =========================================

        /**
         * 塗りつぶした矩形を描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} x - 左端
         * @param {number} y - 上端
         * @param {number} width - 幅
         * @param {number} height - 高さ
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function fillRect(graphics, x, y, width, height, color) {
            graphics.newPath();
            graphics.moveTo(x, y);
            graphics.lineTo(x + width, y);
            graphics.lineTo(x + width, y + height);
            graphics.lineTo(x, y + height);
            graphics.closePath();
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
        }

        /**
         * 整列の基準となる位置（左端・中央・右端）を求める
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} alignMode - "start" / "center" / "end"
         * @returns {number} 基準線の座標
         */
        function getRulePosition(size, alignMode) {
            if (alignMode === "start") { return Math.round(size * ICON_RULE_OFFSET) + 0.5; }
            if (alignMode === "end") { return Math.round(size * (1 - ICON_RULE_OFFSET)) - 0.5; }
            return Math.round(size / 2) + 0.5;
        }

        /**
         * バー2本の並び方向の開始座標を求める（バー・間隔・バーの合計を中央に置く）
         * @param {number} size - アイコンの一辺の長さ
         * @returns {number} 1本目のバーの開始座標
         */
        function getBarStackOrigin(size) {
            var stackLength = size * (ICON_BAR_THICKNESS * 2 + ICON_BAR_GAP);
            return Math.round((size - stackLength) / 2);
        }

        /**
         * 基準線に対するバーの開始座標を求める（端の基準線からはすき間を空け、中央の基準線はバーの下を通す）
         * @param {number} rulePosition - 基準線の座標
         * @param {number} barLength - バーの長さ
         * @param {string} alignMode - "start" / "center" / "end"
         * @param {number} clearance - 端の基準線とバーのすき間
         * @returns {number} バーの開始座標
         */
        function getBarOrigin(rulePosition, barLength, alignMode, clearance) {
            /* 基準線は rulePosition を中心とした太さ1なので、その両端から数えて左右・上下を対称にする
               The rule is 1 unit thick around rulePosition, so measure from its edges to keep both ends symmetrical */
            if (alignMode === "start") { return rulePosition + 0.5 + clearance; }
            if (alignMode === "end") { return rulePosition - 0.5 - clearance - barLength; }
            return Math.round(rulePosition - barLength / 2);
        }

        /**
         * バー2本の長さを並び順に返す（水平は上が短く、90°回した垂直は左が長い）
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} iconType - "horizontal" または "vertical"
         * @returns {number[]} 並び順のバーの長さ
         */
        function getBarLengths(size, iconType) {
            var longBar = Math.round(size * ICON_BAR_LONG);
            var shortBar = Math.round(size * ICON_BAR_SHORT);
            return (iconType === "vertical") ? [longBar, shortBar] : [shortBar, longBar];
        }

        /**
         * 向きに合わせて矩形を描く（垂直方向のアイコンは水平方向の座標を縦横入れ替えて描く）
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {boolean} isVertical - 垂直方向のアイコンなら true
         * @param {number} alignPos - 整列する軸の座標（水平は x、垂直は y）
         * @param {number} stackPos - バーが並ぶ軸の座標（水平は y、垂直は x）
         * @param {number} alignLen - 整列する軸方向の長さ
         * @param {number} stackLen - バーが並ぶ軸方向の長さ
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function fillOrientedRect(graphics, isVertical, alignPos, stackPos, alignLen, stackLen, color) {
            if (isVertical) {
                fillRect(graphics, stackPos, alignPos, stackLen, alignLen, color);
            } else {
                fillRect(graphics, alignPos, stackPos, alignLen, stackLen, color);
            }
        }

        /**
         * 整列アイコン（バー2本＋基準線）を描く
         * 水平方向と垂直方向は縦横が入れ替わるだけなので、同じ手順で描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} alignMode - "start"＝左・上 / "center"＝中央 / "end"＝右・下
         * @param {number[]} color - RGBA の配列
         * @param {string} iconType - "horizontal" または "vertical"
         * @returns {void}
         */
        function drawAlignIcon(graphics, size, alignMode, color, iconType) {
            var isVertical = (iconType === "vertical");
            var rulePosition = getRulePosition(size, alignMode);
            var barThickness = Math.round(size * ICON_BAR_THICKNESS);
            var barGap = Math.round(size * ICON_BAR_GAP);
            var barStackOrigin = getBarStackOrigin(size);
            var barLengths = getBarLengths(size, iconType);
            var clearance = Math.round(size * ICON_RULE_CLEARANCE);
            var ruleInset = Math.round(size * ICON_RULE_INSET);

            /* 基準線を先に描き、バーを上に重ねる（中央の基準線がバーの下を通って見える）
               Draw the rule first and the bars on top, so a center rule runs behind them */
            fillOrientedRect(graphics, isVertical, rulePosition - 0.5, ruleInset, 1, size - ruleInset * 2, color);

            for (var i = 0; i < barLengths.length; i++) {
                fillOrientedRect(graphics, isVertical,
                    getBarOrigin(rulePosition, barLengths[i], alignMode, clearance),
                    barStackOrigin + i * (barThickness + barGap),
                    barLengths[i], barThickness, color);
            }
        }

        /**
         * アイコンの中心線の座標を返す（ケイ線もオブジェクトもここを中心に置く）
         * 太さ1のケイ線がピクセル境界に乗るよう、中心は .5 の位置に取る
         * @param {number} size - アイコンの一辺の長さ
         * @returns {number} 中心の座標
         */
        function getIconCenter(size) {
            return Math.round(size / 2) + 0.5;
        }

        /**
         * 辺の長さを奇数に丸める
         * 中心が .5 の位置にあるので、奇数にすると両側が同じ幅で割り振られ、中心がぴったり合う
         * @param {number} rawLength - 丸める前の長さ
         * @returns {number} 奇数の長さ
         */
        function roundToOddLength(rawLength) {
            var rounded = Math.round(rawLength);
            if (rounded % 2 !== 0) { return rounded; }
            return (rounded > 1) ? rounded - 1 : 1;
        }

        /**
         * アイコンに置くオブジェクトの大きさを求める（どのアイコンも同じ長方形）
         * @param {number} size - アイコンの一辺の長さ
         * @returns {number[]} [幅, 高さ]
         */
        function getBlockSize(size) {
            return [roundToOddLength(size * ICON_BLOCK_WIDTH), roundToOddLength(size * ICON_BLOCK_HEIGHT)];
        }

        /**
         * 端のケイ線（寄せ先を表す線）を1本描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} side - "top" / "bottom" / "left" / "right"
         * @param {number[]} color - RGBA の配列
         * @returns {number} 引いたケイ線の座標
         */
        function drawEdgeRule(graphics, size, side, color) {
            var ruleInset = Math.round(size * ICON_RULE_INSET);
            var ruleLength = size - ruleInset * 2;
            var isTopOrLeft = (side === "top" || side === "left");
            var rulePosition = getRulePosition(size, isTopOrLeft ? "start" : "end");
            if (side === "top" || side === "bottom") {
                fillRect(graphics, ruleInset, rulePosition - 0.5, ruleLength, 1, color);
            } else {
                fillRect(graphics, rulePosition - 0.5, ruleInset, 1, ruleLength, color);
            }
            return rulePosition;
        }

        /**
         * アイコンの中央を貫くケイ線を1本描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} ruleDirection - "vertical"＝縦のケイ線 / "horizontal"＝横のケイ線
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function drawCenterRule(graphics, size, ruleDirection, color) {
            var ruleInset = Math.round(size * ICON_RULE_INSET);
            var ruleLength = size - ruleInset * 2;
            var iconCenter = getIconCenter(size);
            if (ruleDirection === "vertical") {
                fillRect(graphics, iconCenter - 0.5, ruleInset, 1, ruleLength, color);
            } else {
                fillRect(graphics, ruleInset, iconCenter - 0.5, ruleLength, 1, color);
            }
        }

        /**
         * オブジェクトをアイコンの中央に置く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {number[]} blockSize - [幅, 高さ]
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function fillCenteredBlock(graphics, size, blockSize, color) {
            var iconCenter = getIconCenter(size);
            fillRect(graphics,
                iconCenter - blockSize[0] / 2, iconCenter - blockSize[1] / 2,
                blockSize[0], blockSize[1], color);
        }

        /**
         * 水平・垂直の中央に整列するアイコン（十字のケイ線＋中央に置いたオブジェクト）を描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function drawCenterBothIcon(graphics, size, color) {
            /* ケイ線を先に描き、オブジェクトを上に重ねる（ケイ線がオブジェクトの下を通って見える）
               Draw the rules first and the block on top, so they run behind it */
            drawCenterRule(graphics, size, "vertical", color);
            drawCenterRule(graphics, size, "horizontal", color);
            fillCenteredBlock(graphics, size, getBlockSize(size), color);
        }

        /**
         * 片方の軸だけ中央に整列するアイコン（中央を貫くケイ線1本＋オブジェクト）を描く
         * 水平方向中央は縦のケイ線、垂直方向中央は横のケイ線で見分ける
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {number} size - アイコンの一辺の長さ
         * @param {string} iconType - "horizontal" または "vertical"
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function drawCenterOneIcon(graphics, size, iconType, color) {
            drawCenterRule(graphics, size, (iconType === "horizontal") ? "vertical" : "horizontal", color);
            fillCenteredBlock(graphics, size, getBlockSize(size), color);
        }

        /**
         * 方向ボタンのアイコン（寄せ先のケイ線と、そこへ寄せたオブジェクト）を描く
         * @param {ScriptUIGraphics} graphics - 描画対象のグラフィックス
         * @param {string} directionKey - MOVE_ICON_RULES のキー
         * @param {number} size - ボタンの一辺
         * @param {number[]} color - RGBA の配列
         * @returns {void}
         */
        function drawMoveIcon(graphics, directionKey, size, color) {
            var ruleSides, blockSize, clearance, iconCenter, blockX, blockY, side, rulePosition, i;
            ruleSides = MOVE_ICON_RULES[directionKey];
            if (!ruleSides) { return; }
            blockSize = getBlockSize(size);
            clearance = Math.round(size * ICON_RULE_CLEARANCE);
            iconCenter = getIconCenter(size);
            /* ケイ線のない軸では中央に置く / Centred on any axis without a rule */
            blockX = iconCenter - blockSize[0] / 2;
            blockY = iconCenter - blockSize[1] / 2;

            for (i = 0; i < ruleSides.length; i++) {
                side = ruleSides[i];
                rulePosition = drawEdgeRule(graphics, size, side, color);
                /* ケイ線は太さ1なので、その端から数えてすき間を取る / The rule is 1 unit thick, so measure from its edge */
                if (side === "top")    { blockY = rulePosition + 0.5 + clearance; }
                if (side === "bottom") { blockY = rulePosition - 0.5 - clearance - blockSize[1]; }
                if (side === "left")   { blockX = rulePosition + 0.5 + clearance; }
                if (side === "right")  { blockX = rulePosition - 0.5 - clearance - blockSize[0]; }
            }

            fillRect(graphics, blockX, blockY, blockSize[0], blockSize[1], color);
        }

        /**
         * ホバー状態に応じた背景色を返す
         * @param {Button} iconButton - 対象のボタン
         * @returns {number[]} 背景色の RGBA
         */
        function hoverBackground(iconButton) {
            return (iconButton.isHover === true) ? iconHoverBg : iconBaseBg;
        }

        /**
         * ボタンの下地（背景と、マウスオーバー中だけの枠線）を描く
         * @param {Button} iconButton - 対象のボタン
         * @returns {void}
         */
        function drawButtonBase(iconButton) {
            var graphics = iconButton.graphics;
            var width = iconButton.size[0];
            var height = iconButton.size[1];

            try {
                graphics.rectPath(0, 0, width, height);
                graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, hoverBackground(iconButton)));
            } catch (e) {
                try { graphics.drawOSControl(); } catch (osControlError) {}
            }

            /* マウスオーバー中だけ、ボタン領域（正方形）のエッジにグレーの枠を描く
               0.5 ずらすと1pxの線がピクセル境界に乗ってくっきり出る
               A gray border on the square button's edge while hovered; the 0.5 offset keeps the 1px line crisp */
            if (iconButton.isHover === true) {
                try {
                    graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                    graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, iconBorderColor, 1));
                } catch (borderError) {}
            }
        }

        /**
         * 整列アイコンボタンを描画する
         * @param {Button} alignButton - 対象のボタン（iconType と alignMode を持つ）
         * @returns {void}
         */
        function drawAlignButton(alignButton) {
            drawButtonBase(alignButton);
            if (alignButton.iconType === "center") {
                drawCenterBothIcon(alignButton.graphics, alignButton.size[0], iconColor);
            } else if (alignButton.alignMode === "center") {
                drawCenterOneIcon(alignButton.graphics, alignButton.size[0], alignButton.iconType, iconColor);
            } else {
                drawAlignIcon(alignButton.graphics, alignButton.size[0], alignButton.alignMode, iconColor, alignButton.iconType);
            }
        }

        /**
         * 移動ボタン（寄せ先のケイ線とオブジェクト）を描画する
         * @param {Button} moveButton - 対象のボタン（directionKey を持つ）
         * @returns {void}
         */
        function drawMoveButton(moveButton) {
            drawButtonBase(moveButton);
            drawMoveIcon(moveButton.graphics, moveButton.directionKey, moveButton.size[0], iconColor);
        }

        // =========================================
        // UI 部品 / UI helpers
        // =========================================

        /**
         * コントロールのサイズを固定する（最小・推奨・最大を同じ値でそろえる）
         * @param {object} targetControl - 対象のコントロール
         * @param {number} width - 幅
         * @param {number} height - 高さ
         * @returns {void}
         */
        function fixControlSize(targetControl, width, height) {
            targetControl.minimumSize = [width, height];
            targetControl.preferredSize = [width, height];
            targetControl.maximumSize = [width, height];
        }

        /**
         * コントロールを再描画する（notify は環境により例外を投げ得るので保護）
         * @param {object} targetControl - 対象のコントロール
         * @returns {void}
         */
        function redrawControl(targetControl) {
            try { targetControl.notify("onDraw"); } catch (e) {}
        }

        /**
         * 行に∧∨と入力欄を隙間0で突き合わせて追加する。↑↓キーも∧∨と同じ処理で増減し、
         * 増減のたびに入力欄の onChange（確定時と同じ処理）を呼ぶ
         * @param {Group} parentRow - 追加先の行グループ
         * @param {string} initialText - 入力欄の初期値
         * @param {object} stepOptions - addStepper() に渡す増減設定（min / max / integer）
         * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
         */
        function addSteppedInput(parentRow, initialText, stepOptions) {
            var stepperInputGroup = parentRow.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;

            var numberInput;
            /* ∧∨での増減は確定値なので、確定時と同じ処理を走らせる（プログラムからの変更では onChange が発火しない）
               A step is a committed value; programmatic changes do not fire onChange, so call it */
            stepOptions.onStep = function (steppedInput) {
                if (typeof steppedInput.onChange === "function") { steppedInput.onChange(); }
            };
            var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
            numberInput = stepperInputGroup.add("edittext", undefined, initialText);
            numberInput.characters = FIELD_CHARS;
            numberInput.stepperGroup = stepperGroup;
            bindSteppedArrowKeys(numberInput, stepperGroup);
            return numberInput;
        }

        /**
         * addSteppedInput() で作った入力欄と∧∨の有効／無効をまとめて切り替え、∧∨を描き直す
         * @param {EditText} numberInput - 対象の入力欄
         * @param {boolean} isEnabled - 有効にするなら true
         * @returns {void}
         */
        function setSteppedInputEnabled(numberInput, isEnabled) {
            numberInput.enabled = isEnabled;
            numberInput.stepperGroup.enabled = isEnabled;
            redrawSteppersIn(numberInput.stepperGroup);
        }

        /**
         * Option（Alt）キーが押されているか判定する
         * @returns {boolean} 押されていれば true（取得できない環境では false）
         */
        function isAltPressed() {
            try {
                return ScriptUI.environment.keyboardState.altKey === true;
            } catch (e) {
                return false;
            }
        }

        /**
         * マウスオーバーの状態を iconButton.isHover に反映して再描画する
         * @param {Button} iconButton - 対象のボタン
         * @returns {void}
         */
        function attachHover(iconButton) {
            iconButton.addEventListener("mouseover", function() { iconButton.isHover = true; redrawControl(iconButton); });
            iconButton.addEventListener("mouseout", function() { iconButton.isHover = false; redrawControl(iconButton); });
        }

        // リンクアイコン（再利用パーツ） / Link toggle (reusable)

        // -----------------------------------------
        // リンクアイコンの寸法 / Link toggle metrics
        // -----------------------------------------
        var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
        var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
        var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
        var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
        var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
        var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

        // -----------------------------------------
        // リンクアイコンの配色 / Link toggle colors
        // -----------------------------------------
        var LINK_UI_DARK = isDarkUI();
        /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
           Translucent overlays that follow the dialog background; same values as the stepper buttons */
        var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
        var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
        var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
        var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

        // -----------------------------------------
        // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
        // -----------------------------------------
        /**
         * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
         * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
         * @param {Group} parent - 追加先
         * @param {boolean} initialValue - 連動の初期値
         * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
         * @returns {Group} アイコン（.value で連動中かを読む）
         */
        function addLinkToggle(parent, initialValue, onToggle) {
            var linkToggle = parent.add("group");
            linkToggle.preferredSize = LINK_ICON_SIZE;
            linkToggle.minimumSize = LINK_ICON_SIZE;
            linkToggle.maximumSize = LINK_ICON_SIZE;
            linkToggle.value = initialValue;

            linkToggle.onDraw = function () {
                var iconGraphics = linkToggle.graphics;
                var iconWidth = LINK_ICON_SIZE[0];
                var iconHeight = LINK_ICON_SIZE[1];
                /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
                var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
                /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
                if (linkToggle.value && !isDimmed) {
                    iconGraphics.newPath();
                    iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                    iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                    iconGraphics.newPath();
                    iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                    iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
                }
                drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR);
            };

            linkToggle.addEventListener("mousedown", function () {
                if (!isLinkToggleEnabledInTree(linkToggle)) return;
                linkToggle.value = !linkToggle.value;
                redrawLinkToggle(linkToggle);
                if (onToggle) onToggle();
            });
            return linkToggle;
        }

        /**
         * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
         * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
         * @param {boolean} isLinked - 連動にするなら true
         * @returns {void}
         */
        function setLinkToggleValue(linkToggle, isLinked) {
            if (linkToggle.value === isLinked) return;
            linkToggle.value = isLinked;
            redrawLinkToggle(linkToggle);
        }

        /**
         * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
         * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
         * @param {boolean} isEnabled - 有効にするなら true
         * @returns {void}
         */
        function setLinkToggleEnabled(linkToggle, isEnabled) {
            if (linkToggle.enabled === isEnabled) return;
            linkToggle.enabled = isEnabled;
            redrawLinkToggle(linkToggle);
        }

        /**
         * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
         * @param {Object} control - 判定するコントロール
         * @returns {boolean} すべて有効なら true
         */
        function isLinkToggleEnabledInTree(control) {
            for (var node = control; node; node = node.parent) {
                if (!node.enabled) return false;
            }
            return true;
        }

        /**
         * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
         * @param {Group} linkToggle - 描き直すアイコン
         * @returns {void}
         */
        function redrawLinkToggle(linkToggle) {
            linkToggle.hide();
            linkToggle.show();
        }

        // -----------------------------------------
        // アイコンの形 / Icon geometry
        // -----------------------------------------
        /**
         * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
         * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
         * @param {ScriptUIGraphics} iconGraphics - 描画先
         * @param {number} iconWidth - 描画範囲の幅
         * @param {number} iconHeight - 描画範囲の高さ
         * @param {boolean} isLinked - 連動中なら true
         * @param {number[]} iconColor - [r, g, b, a]
         * @returns {void}
         */
        function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor) {
            var iconScale = Math.min(iconWidth, iconHeight) / 22;
            var offsetX = (iconWidth - 22 * iconScale) / 2;
            var offsetY = (iconHeight - 22 * iconScale) / 2;
            var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
            for (var i = 0; i < strokes.length; i++) {
                var strokePoints = strokes[i].points;
                /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
                iconGraphics.newPath();
                for (var j = 0; j < strokePoints.length; j++) {
                    var pointX = offsetX + strokePoints[j][0] * iconScale;
                    var pointY = offsetY + strokePoints[j][1] * iconScale;
                    if (j === 0) iconGraphics.moveTo(pointX, pointY);
                    else iconGraphics.lineTo(pointX, pointY);
                }
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
            }
        }

        /**
         * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
         * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
         * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
         * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
         */
        function buildLinkedChainStrokes() {
            /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
               Keep only a stub on the left so it stays clear of the lower ring's hook */
            var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
                .concat([[14.5, 11.2]])
                .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
            var ringStart = upperRing[0];
            var ringEnd = upperRing[upperRing.length - 1];
            var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
            /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
            var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
            var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

            var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
                var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
                trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
                return [trimmed];
            });
            var strokes = [];
            for (var i = 0; i < upperStrands.length; i++) {
                strokes.push(upperStrands[i]);
                strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
            }
            return strokes;
        }

        /**
         * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
         * @param {Array<number[]>} centerline - 中心線の点列
         * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
         * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
         */
        function buildStrandStrokes(centerline, clipStrand) {
            var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
            var strokes = [];
            for (var k = 0; k < LINK_STRAND_COUNT; k++) {
                /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
                var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
                var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
                for (var j = 0; j < strandPieces.length; j++) {
                    /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                    if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
                }
            }
            return strokes;
        }

        /**
         * 点列の両端を、端の向きのまま length だけ延ばす。
         * @param {Array<number[]>} points - 点列
         * @param {number} length - 延ばす長さ
         * @returns {Array<number[]>} 延ばした点列
         */
        function extendPolylineEnds(points, length) {
            /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
            function extendBeyond(from, to) {
                var dx = to[0] - from[0];
                var dy = to[1] - from[1];
                var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
                return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
            }
            var lastIndex = points.length - 1;
            return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
        }

        /**
         * 点が直線のどちら側にあるかを符号で返す。
         * @param {number[]} point - 点
         * @param {number[]} linePoint - 直線上の1点
         * @param {number[]} direction - 直線の向き
         * @returns {number} 正・負で側を表す値
         */
        function sideOfLine(point, linePoint, direction) {
            return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
        }

        /**
         * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
         * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
         * @param {Array<number[]>} points - 点列
         * @param {number[]} cutPoint - 切る直線上の1点
         * @param {number[]} direction - 切る直線の向き
         * @param {number} outsideSign - 切り捨てる側の符号
         * @returns {Array<number[]>} 切った点列
         */
        function trimPolylineTail(points, cutPoint, direction, outsideSign) {
            var lastIndex = points.length - 1;
            var searchLimit = Math.max(0, lastIndex - 12);
            var index = lastIndex;
            while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
            if (index === lastIndex) return points.slice(0);
            var inside = points[index];
            var outside = points[index + 1];
            var insideSide = sideOfLine(inside, cutPoint, direction);
            var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
            return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
        }

        /**
         * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
         * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
         * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
         * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
         */
        function buildUnlinkedChainStrokes() {
            var slashStart = [3.5, 3.5];
            var slashEnd = [18.5, 18.5];
            var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
            var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

            /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
            function clipAroundSlash(strandPoints) {
                return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
            }
            var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
            strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
            return strokes;
        }

        /**
         * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
         * @param {Array<number[]>} points - 点列
         * @returns {Array<number[]>} 細かくした点列
         */
        function densifyPoints(points) {
            var densePoints = [points[0]];
            for (var i = 1; i < points.length; i++) {
                var from = points[i - 1];
                var to = points[i];
                var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
                for (var j = 1; j <= steps; j++) {
                    densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
                }
            }
            return densePoints;
        }

        /**
         * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
         * @param {Array<number[]>} points - 点列
         * @param {number} offset - ずらす距離
         * @returns {Array<number[]>} ずらした点列
         */
        function offsetPolyline(points, offset) {
            var shifted = [];
            for (var i = 0; i < points.length; i++) {
                var before = points[Math.max(0, i - 1)];
                var after = points[Math.min(points.length - 1, i + 1)];
                var tangentX = after[0] - before[0];
                var tangentY = after[1] - before[1];
                var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
                shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
            }
            return shifted;
        }

        /**
         * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
         * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
         * @param {Array<number[]>} points - 点列
         * @param {number[]} lineStart - 直線上の1点
         * @param {number[]} lineEnd - 直線上のもう1点
         * @param {number} clearance - 空ける距離
         * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
         */
        function clipOutsideBand(points, lineStart, lineEnd, clearance) {
            var directionX = lineEnd[0] - lineStart[0];
            var directionY = lineEnd[1] - lineStart[1];
            var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

            /* 直線からの符号付き距離 / signed distance from the line */
            function signedDistance(point) {
                return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
            }
            /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
            function interpolateAt(from, to, fromDistance, toDistance, boundary) {
                var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
                return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
            }

            var pieces = [];
            var currentPiece = [];
            for (var i = 0; i < points.length; i++) {
                var distance = signedDistance(points[i]);
                var isOutside = Math.abs(distance) >= clearance;
                if (i > 0) {
                    var previousDistance = signedDistance(points[i - 1]);
                    var wasOutside = Math.abs(previousDistance) >= clearance;
                    if (wasOutside && !isOutside) {
                        /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                        currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                        if (currentPiece.length > 1) pieces.push(currentPiece);
                        currentPiece = [];
                    } else if (!wasOutside && isOutside) {
                        /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                        currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                    }
                }
                if (isOutside) currentPiece.push(points[i]);
            }
            if (currentPiece.length > 1) pieces.push(currentPiece);
            return pieces;
        }

        /**
         * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
         * @param {number} centerX - 中心X
         * @param {number} centerY - 中心Y
         * @param {number} radiusX - 横の半径
         * @param {number} radiusY - 縦の半径
         * @param {number} startDegrees - 開始角度
         * @param {number} endDegrees - 終了角度
         * @returns {Array<number[]>} 点列
         */
        function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
            var arcSteps = 12;
            var arcPoints = [];
            for (var i = 0; i <= arcSteps; i++) {
                var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
                arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
            }
            return arcPoints;
        }

        /**
         * 点列を 22px 四方の中心で180度回す。
         * @param {Array<number[]>} points - 点列
         * @returns {Array<number[]>} 回した点列
         */
        function rotatePointsHalfTurn(points) {
            var rotated = [];
            for (var i = 0; i < points.length; i++) {
                rotated.push([22 - points[i][0], 22 - points[i][1]]);
            }
            return rotated;
        }

        // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle

        // 設定の保存（再利用パーツ） / Settings store (reusable)

        var SETTINGS_STORE_FOLDER_NAME = "illustrator-scripts"; /* Folder.userData の下に作るフォルダー / folder created under Folder.userData */
        var SETTINGS_STORE_MAX_DEPTH = 32;                                /* 入れ子の上限（循環参照よけ）/ nesting limit (guards against cycles) */

        /**
         * 設定の保存先を作る。寿命は "session"（Illustrator の終了まで）か "persistent"（ファイルに保存）
         * @param {string} storeName - 保存名（ふつうは SCRIPT_NAME）。ファイル名と $.global のキーに使う
         * @param {string} lifetime - "session" または "persistent"
         * @param {Object} [storeOptions] - { legacy: function () → 旧形式の保存値のオブジェクト|null }
         * @returns {{load: Function, save: Function, clear: Function}} 読み込み・保存・消去の関数
         */
        function createSettingsStore(storeName, lifetime, storeOptions) {
            var isPersistent = (lifetime === "persistent");
            var legacyReader = (storeOptions && typeof storeOptions.legacy === "function") ? storeOptions.legacy : null;
            var safeStoreName = String(storeName).replace(/[\\\/:*?"<>|]/g, "_");
            var sessionKey = "__" + safeStoreName + "_Settings";
            var settingsFile = isPersistent
                ? new File(Folder.userData + "/" + SETTINGS_STORE_FOLDER_NAME + "/" + safeStoreName + ".json")
                : null;

            /**
             * 保存してある文字列を返す
             * @returns {string|null} 保存文字列。1度も保存していなければ null
             */
            function readStoredText() {
                if (!isPersistent) {
                    return (typeof $.global[sessionKey] === "string") ? $.global[sessionKey] : null;
                }
                return settingsStoreReadTextFile(settingsFile);
            }

            /**
             * 文字列を保存する
             * @param {string} storedText - 保存する文字列
             * @returns {boolean} 保存できたら true
             */
            function writeStoredText(storedText) {
                if (!isPersistent) {
                    $.global[sessionKey] = storedText;
                    return true;
                }
                return settingsStoreWriteTextFile(settingsFile, storedText);
            }

            /**
             * 保存値を読み込み、既定値と突き合わせて返す（型の合わない値・知らない項目は捨てる）
             * @param {Object} defaultSettings - 既定値
             * @returns {Object} 設定（毎回新しいオブジェクト）
             */
            function load(defaultSettings) {
                var savedSettings = null;
                try {
                    var storedText = readStoredText();
                    if (storedText !== null) {
                        savedSettings = settingsStoreParse(storedText);
                    } else if (legacyReader) {
                        savedSettings = legacyReader();
                    }
                } catch (e) {
                    $.writeln("SettingsStore.load(" + storeName + "): " + e);
                    savedSettings = null;
                }
                return settingsStoreMerge(defaultSettings, savedSettings);
            }

            /**
             * 設定を保存する
             * @param {Object} settingValues - 保存する値
             * @returns {boolean} 保存できたら true
             */
            function save(settingValues) {
                try {
                    return writeStoredText(settingsStoreSerialize(settingValues, "", 0));
                } catch (e) {
                    $.writeln("SettingsStore.save(" + storeName + "): " + e);
                    return false;
                }
            }

            /**
             * 保存を消す。旧形式を読み継ぐストアでは空の保存を書き、旧設定が戻らないようにする
             * @returns {boolean} 消せたら true
             */
            function clear() {
                if (legacyReader) return writeStoredText("{}");
                if (!isPersistent) {
                    try { delete $.global[sessionKey]; } catch (e) { $.global[sessionKey] = undefined; }
                    return true;
                }
                try {
                    return settingsFile.exists ? settingsFile.remove() : true;
                } catch (e) {
                    $.writeln("SettingsStore.clear(" + storeName + "): " + e);
                    return false;
                }
            }

            return { load: load, save: save, clear: clear };
        }

        /**
         * 旧形式の設定ファイルを読む（key=value の行 / toSource / JSON を自動判別。eval は使わない）
         * @param {File|string} legacyFileOrPath - 旧ファイルかそのパス
         * @returns {Object|null} 読み込んだ値（key=value は値がすべて文字列）。無い・読めないときは null
         */
        function readSettingsLegacyFile(legacyFileOrPath) {
            try {
                var legacyFile = (legacyFileOrPath instanceof File) ? legacyFileOrPath : new File(legacyFileOrPath);
                var legacyText = settingsStoreReadTextFile(legacyFile);
                return (legacyText === null) ? null : settingsStoreParseLegacyText(legacyText);
            } catch (e) {
                $.writeln("readSettingsLegacyFile: " + e);
                return null;
            }
        }

        /**
         * app.preferences に文字列で保存していた旧設定を読む（形式は readSettingsLegacyFile と同じく自動判別）
         * @param {string} preferenceKey - 環境設定のキー
         * @returns {Object|null} 読み込んだ値。無い・読めないときは null
         */
        function readSettingsLegacyPreference(preferenceKey) {
            try {
                var legacyText = app.preferences.getStringPreference(preferenceKey);
                if (!legacyText) return null;
                return settingsStoreParseLegacyText(String(legacyText));
            } catch (e) {
                $.writeln("readSettingsLegacyPreference: " + e);
                return null;
            }
        }

        /**
         * テキストファイルを UTF-8 で読む
         * @param {File} textFile - 読むファイル
         * @returns {string|null} 中身。ファイルが無ければ null
         */
        function settingsStoreReadTextFile(textFile) {
            if (!textFile.exists) return null;
            textFile.encoding = "UTF-8";
            if (!textFile.open("r")) throw new Error("cannot open " + textFile.fsName);
            try {
                return textFile.read().replace(/^\uFEFF/, "");
            } finally {
                textFile.close();
            }
        }

        /**
         * テキストファイルを UTF-8 で書く（フォルダーが無ければ作る）
         * @param {File} textFile - 書くファイル
         * @param {string} fileText - 中身
         * @returns {boolean} 書けたら true
         */
        function settingsStoreWriteTextFile(textFile, fileText) {
            try {
                var parentFolder = textFile.parent;
                if (!parentFolder.exists && !parentFolder.create()) throw new Error("cannot create " + parentFolder.fsName);
                textFile.encoding = "UTF-8";
                textFile.lineFeed = "Unix";
                if (!textFile.open("w")) throw new Error("cannot open " + textFile.fsName);
                try {
                    textFile.write(fileText);
                } finally {
                    textFile.close();
                }
                return true;
            } catch (e) {
                $.writeln("SettingsStore write: " + e);
                return false;
            }
        }

        /**
         * 値が配列か
         * @param {*} checkedValue - 調べる値
         * @returns {boolean} 配列なら true
         */
        function settingsStoreIsArray(checkedValue) {
            return Object.prototype.toString.call(checkedValue) === "[object Array]";
        }

        /**
         * 値が素のオブジェクト（{ } で作ったもの）か
         * @param {*} checkedValue - 調べる値
         * @returns {boolean} 素のオブジェクトなら true
         */
        function settingsStoreIsPlainObject(checkedValue) {
            return checkedValue !== null && typeof checkedValue === "object"
                && Object.prototype.toString.call(checkedValue) === "[object Object]"
                && checkedValue.constructor === Object;
        }

        /**
         * 文字列を JSON の文字列リテラルにする（ASCII 以外は \uXXXX にして、文字コードの取り違えに強くする）
         * @param {string} sourceText - 文字列
         * @returns {string} 引用符つきの文字列
         */
        function settingsStoreQuote(sourceText) {
            var quotedText = "\"";
            for (var i = 0; i < sourceText.length; i++) {
                var charCode = sourceText.charCodeAt(i);
                var oneChar = sourceText.charAt(i);
                if (oneChar === "\"" || oneChar === "\\") quotedText += "\\" + oneChar;
                else if (oneChar === "\n") quotedText += "\\n";
                else if (oneChar === "\r") quotedText += "\\r";
                else if (oneChar === "\t") quotedText += "\\t";
                else if (charCode < 0x20 || charCode > 0x7E) quotedText += "\\u" + ("0000" + charCode.toString(16)).slice(-4);
                else quotedText += oneChar;
            }
            return quotedText + "\"";
        }

        /**
         * 値を JSON の文字列にする（オブジェクトは1項目1行、中身が値だけの配列は1行）。
         * undefined・関数・DOM オブジェクトは項目ごと省き、配列の中では null にする。有限でない数値は null
         * @param {*} sourceValue - 値
         * @param {string} indentText - 今の字下げ
         * @param {number} depth - 入れ子の深さ
         * @returns {string|undefined} JSON の文字列。書けない値は undefined
         */
        function settingsStoreSerialize(sourceValue, indentText, depth) {
            if (depth > SETTINGS_STORE_MAX_DEPTH) throw new Error("settings are nested too deeply");
            if (sourceValue === null) return "null";
            var valueType = typeof sourceValue;
            if (valueType === "boolean") return sourceValue ? "true" : "false";
            if (valueType === "number") return isFinite(sourceValue) ? String(sourceValue) : "null";
            if (valueType === "string") return settingsStoreQuote(sourceValue);
            var innerIndent = indentText + "  ";
            var itemTexts = [];
            var i;
            if (settingsStoreIsArray(sourceValue)) {
                var hasNested = false;
                for (i = 0; i < sourceValue.length; i++) {
                    var itemText = settingsStoreSerialize(sourceValue[i], innerIndent, depth + 1);
                    itemTexts.push(itemText === undefined ? "null" : itemText);
                    if (sourceValue[i] !== null && typeof sourceValue[i] === "object") hasNested = true;
                }
                if (!itemTexts.length) return "[]";
                if (!hasNested) return "[" + itemTexts.join(", ") + "]";
                return "[\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "]";
            }
            if (settingsStoreIsPlainObject(sourceValue)) {
                for (var key in sourceValue) {
                    if (!sourceValue.hasOwnProperty(key)) continue;
                    var memberText = settingsStoreSerialize(sourceValue[key], innerIndent, depth + 1);
                    if (memberText !== undefined) itemTexts.push(settingsStoreQuote(key) + ": " + memberText);
                }
                if (!itemTexts.length) return "{}";
                return "{\n" + innerIndent + itemTexts.join(",\n" + innerIndent) + "\n" + indentText + "}";
            }
            return undefined; /* 関数・DOM オブジェクトなど / functions, DOM objects, etc. */
        }

        /**
         * JSON（と toSource の出力）を読む。eval は使わない。
         * キーの引用符なし・'…' の文字列・全体の ( ) ・末尾のカンマ・(void 0) も受け付ける
         * @param {string} sourceText - 読む文字列
         * @returns {*} 読み込んだ値
         */
        function settingsStoreParse(sourceText) {
            var readPos = 0;
            var textLength = sourceText.length;

            /**
             * 読み取り位置で失敗を知らせる
             * @param {string} reasonText - 理由
             * @returns {void}
             */
            function fail(reasonText) {
                throw new Error("settings parse error at " + readPos + ": " + reasonText);
            }

            /**
             * 空白を読み飛ばす
             * @returns {void}
             */
            function skipSpaces() {
                while (readPos < textLength && /\s/.test(sourceText.charAt(readPos))) readPos++;
            }

            /**
             * 識別子（英数字・_・$）を読む
             * @returns {string} 識別子。無ければ空文字
             */
            function readWord() {
                var startPos = readPos;
                while (readPos < textLength && /[\w$]/.test(sourceText.charAt(readPos))) readPos++;
                return sourceText.substring(startPos, readPos);
            }

            /**
             * 引用符で囲んだ文字列を読む（" と ' のどちらでも）
             * @returns {string} 文字列
             */
            function readString() {
                var quoteChar = sourceText.charAt(readPos++);
                var resultText = "";
                while (readPos < textLength) {
                    var oneChar = sourceText.charAt(readPos++);
                    if (oneChar === quoteChar) return resultText;
                    if (oneChar !== "\\") { resultText += oneChar; continue; }
                    var escapeChar = sourceText.charAt(readPos++);
                    if (escapeChar === "n") resultText += "\n";
                    else if (escapeChar === "r") resultText += "\r";
                    else if (escapeChar === "t") resultText += "\t";
                    else if (escapeChar === "b") resultText += "\b";
                    else if (escapeChar === "f") resultText += "\f";
                    else if (escapeChar === "v") resultText += "\v";
                    else if (escapeChar === "0") resultText += "\0";
                    else if (escapeChar === "u" || escapeChar === "x") {
                        var hexLength = (escapeChar === "u") ? 4 : 2;
                        var hexText = sourceText.substr(readPos, hexLength);
                        if (!new RegExp("^[0-9A-Fa-f]{" + hexLength + "}$").test(hexText)) fail("bad escape");
                        resultText += String.fromCharCode(parseInt(hexText, 16));
                        readPos += hexLength;
                    } else resultText += escapeChar;
                }
                fail("unterminated string");
            }

            /**
             * 値を1つ読む
             * @param {number} depth - 入れ子の深さ
             * @returns {*} 値
             */
            function readValue(depth) {
                if (depth > SETTINGS_STORE_MAX_DEPTH) fail("nested too deeply");
                skipSpaces();
                var oneChar = sourceText.charAt(readPos);
                if (oneChar === "{") return readObject(depth);
                if (oneChar === "[") return readArray(depth);
                if (oneChar === "\"" || oneChar === "'") return readString();
                if (oneChar === "(") {
                    readPos++;
                    var innerValue = readValue(depth + 1);
                    skipSpaces();
                    if (sourceText.charAt(readPos) !== ")") fail("expected )");
                    readPos++;
                    return innerValue;
                }
                var numberMatch = /^-?(\d+\.?\d*|\.\d+)([eE][+\-]?\d+)?/.exec(sourceText.substring(readPos, readPos + 64));
                if (numberMatch) {
                    readPos += numberMatch[0].length;
                    return Number(numberMatch[0]);
                }
                var wordText = readWord();
                if (wordText === "true") return true;
                if (wordText === "false") return false;
                if (wordText === "null") return null;
                if (wordText === "NaN") return NaN;
                if (wordText === "Infinity") return Infinity;
                if (wordText === "void") { readValue(depth + 1); return undefined; } /* toSource の (void 0) */
                fail("unexpected " + (wordText || oneChar || "end of text"));
            }

            /**
             * 配列を読む
             * @param {number} depth - 入れ子の深さ
             * @returns {Array} 配列
             */
            function readArray(depth) {
                var resultArray = [];
                readPos++;
                skipSpaces();
                while (sourceText.charAt(readPos) !== "]") {
                    resultArray.push(readValue(depth + 1));
                    skipSpaces();
                    if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                    if (sourceText.charAt(readPos) !== "]") fail("expected , or ]");
                }
                readPos++;
                return resultArray;
            }

            /**
             * オブジェクトを読む（__proto__ のキーは捨てる）
             * @param {number} depth - 入れ子の深さ
             * @returns {Object} オブジェクト
             */
            function readObject(depth) {
                var resultObject = {};
                readPos++;
                skipSpaces();
                while (sourceText.charAt(readPos) !== "}") {
                    var keyChar = sourceText.charAt(readPos);
                    var memberKey = (keyChar === "\"" || keyChar === "'") ? readString() : readWord();
                    if (memberKey === "") fail("expected a key");
                    skipSpaces();
                    if (sourceText.charAt(readPos) !== ":") fail("expected :");
                    readPos++;
                    var memberValue = readValue(depth + 1);
                    if (memberKey !== "__proto__") resultObject[memberKey] = memberValue;
                    skipSpaces();
                    if (sourceText.charAt(readPos) === ",") { readPos++; skipSpaces(); continue; }
                    if (sourceText.charAt(readPos) !== "}") fail("expected , or }");
                }
                readPos++;
                return resultObject;
            }

            var parsedValue = readValue(0);
            skipSpaces();
            if (readPos < textLength) fail("unexpected text after the value");
            return parsedValue;
        }

        /**
         * 旧形式の文字列を読む。{ [ ( で始まれば JSON / toSource、それ以外は key=value の行とみなす
         * @param {string} legacyText - 旧形式の文字列
         * @returns {Object|null} 読み込んだ値
         */
        function settingsStoreParseLegacyText(legacyText) {
            var trimmedText = legacyText.replace(/^\uFEFF/, "").replace(/^\s+|\s+$/g, "");
            if (trimmedText === "") return null;
            if (/^[\{\[\(]/.test(trimmedText)) return settingsStoreParse(trimmedText);
            var keyValues = {};
            var textLines = trimmedText.split(/\r\n|\r|\n/);
            for (var i = 0; i < textLines.length; i++) {
                var separatorIndex = textLines[i].indexOf("=");
                if (separatorIndex < 1) continue;
                var lineKey = textLines[i].substring(0, separatorIndex).replace(/^\s+|\s+$/g, "");
                if (lineKey !== "" && lineKey !== "__proto__") keyValues[lineKey] = textLines[i].substring(separatorIndex + 1);
            }
            return keyValues;
        }

        /**
         * 値を深くコピーする（素のデータだけ。関数・DOM オブジェクトは null）
         * @param {*} sourceValue - コピー元
         * @returns {*} コピー
         */
        function settingsStoreClone(sourceValue) {
            if (sourceValue === null || typeof sourceValue !== "object") {
                return (typeof sourceValue === "function" || sourceValue === undefined) ? null : sourceValue;
            }
            var i;
            if (settingsStoreIsArray(sourceValue)) {
                var arrayCopy = [];
                for (i = 0; i < sourceValue.length; i++) arrayCopy.push(settingsStoreClone(sourceValue[i]));
                return arrayCopy;
            }
            if (!settingsStoreIsPlainObject(sourceValue)) return null;
            var objectCopy = {};
            for (var key in sourceValue) {
                if (sourceValue.hasOwnProperty(key)) objectCopy[key] = settingsStoreClone(sourceValue[key]);
            }
            return objectCopy;
        }

        /**
         * 保存値を既定値と突き合わせる。型は既定値に合わせ、合わなければ既定値を使う。
         * 既定値が {} か null なら中身を問わず受け取り、配列は配列なら受け取る。既定値に無い項目は捨てる
         * @param {*} defaultValue - 既定値
         * @param {*} savedValue - 保存値
         * @returns {*} 突き合わせた値（新しいオブジェクト）
         */
        function settingsStoreMerge(defaultValue, savedValue) {
            if (defaultValue === null || defaultValue === undefined) {
                return (savedValue === undefined) ? null : settingsStoreClone(savedValue);
            }
            var defaultType = typeof defaultValue;
            var savedType = typeof savedValue;
            if (defaultType === "boolean") {
                if (savedType === "boolean") return savedValue;
                if (savedValue === 1 || savedValue === "1" || savedValue === "true") return true;
                if (savedValue === 0 || savedValue === "0" || savedValue === "false") return false;
                return defaultValue;
            }
            if (defaultType === "number") {
                if (savedType === "number" && isFinite(savedValue)) return savedValue;
                if (savedType === "string" && /\S/.test(savedValue)) {
                    var parsedNumber = Number(savedValue);
                    if (isFinite(parsedNumber)) return parsedNumber;
                }
                return defaultValue;
            }
            if (defaultType === "string") {
                if (savedType === "string") return savedValue;
                if (savedType === "number" && isFinite(savedValue)) return String(savedValue);
                if (savedType === "boolean") return String(savedValue);
                return defaultValue;
            }
            if (settingsStoreIsArray(defaultValue)) {
                return settingsStoreClone(settingsStoreIsArray(savedValue) ? savedValue : defaultValue);
            }
            if (defaultType === "object") {
                var savedIsObject = settingsStoreIsPlainObject(savedValue);
                var hasDefaultKeys = false;
                var mergedObject = {};
                for (var key in defaultValue) {
                    if (!defaultValue.hasOwnProperty(key)) continue;
                    hasDefaultKeys = true;
                    mergedObject[key] = settingsStoreMerge(defaultValue[key], savedIsObject ? savedValue[key] : undefined);
                }
                /* 既定値が {} なら自由な入れ物として中身ごと受け取る / an empty default {} is a free-form map */
                if (!hasDefaultKeys && savedIsObject) return settingsStoreClone(savedValue);
                return mergedObject;
            }
            return defaultValue;
        }

        // 設定の保存（再利用パーツ）ここまで / End of the reusable settings store

        // =========================================
        // 設定の記憶 / Settings memory
        // =========================================
        /* パレットの設定は Illustrator を終了するまで残す（パレット側のエンジンに置く。ワーカーには入れない）。
           旧版の $.global のキーを1度だけ読み継ぐ
           Palette settings last until Illustrator quits, kept in the palette's engine (never in the workers);
           the old $.global key is read once */
        var settingsStore = createSettingsStore(SCRIPT_NAME, "session", {
            legacy: function () { return $.global[LEGACY_SETTINGS_KEY] || null; }
        });

        // =========================================
        // パレット構築 / Palette builder
        // =========================================

        /* チェックボックスの参照（アイコンのクリック時に読む）/ The checkboxes, read when an icon is clicked */
        var previewBoundsCheckbox = null;
        var glyphBoundsCheckbox = null;
        var changeJustificationCheckbox = null;
        var opticalAdjustCheckbox = null;
        /* マージン欄・単位ラベル・ガイド表示チェックボックスの参照
           The margin field, its unit label, and the guide checkbox */
        /* 辺の名前をキーにしたマージンの入力欄 / Margin fields, keyed by side */
        var marginFields = {};
        var marginPanel = null;
        var linkMarginsToggle = null;
        /* マージン欄がまだ既定値のままか。定規の単位が分かった時点で、その単位の既定値に入れ直す
           Whether the margin fields still hold defaults; they are refilled once the ruler unit is known */
        var marginsAreDefault = false;
        /* 伸張欄がまだ既定値のままか。マージン欄と同じく、定規の単位が分かった時点でその単位の既定値に入れ直す
           Whether the extension field still holds a default; like the margins, it is refilled once the ruler unit is known */
        var extensionIsDefault = false;
        var alignToBleedCheckbox = null;
        var alignPerArtboardCheckbox = null;
        var showGuideCheckbox = null;
        /* 分割ガイドの切り替えラジオと分割数の入力欄（どちらも軸・モードの名前をキーにする）
           The division mode radios and count fields, keyed by mode and axis */
        var divideRadios = {};
        var divideFields = {};
        var dividePanel = null;
        var artboardEdgeCheckbox = null;
        /* パレットを組み立て終えたか。組み立てのあいだは控えを書かない
           （まだ作っていないコントロールを既定値として読み、引き継いだ設定を消してしまうため）
           Whether the palette is fully built; settings are not stored until it is, or controls that do not
           exist yet would be read as defaults and wipe the carried-over settings */
        var isPaletteReady = false;
        /* 現在の定規単位（パレットへフォーカスが来るたびに取り直す）/ The current ruler unit, re-read on focus */
        var currentUnitInfo = FALLBACK_UNIT_INFO;
        /* 状況表示の参照 / The status line */
        var statusText = null;

        /* 水平・垂直方向中央に整列するボタンの定義。移動ボタンの十字の中央に置く
           The centre-on-both-axes button; it sits in the middle of the move-button cross */
        var CENTER_ALIGN_BUTTON = { iconType: "center", alignMode: "center", tooltip: "alignCenterAll", alignKeys: ["horizontalCenter", "verticalCenter"] };

        /* 片方の軸だけ中央に整列するボタン2つ。十字の下に横に並べる
           The two single-axis centre buttons, in a row below the cross */
        var CENTER_ALIGN_BUTTONS = [
            { iconType: "horizontal", alignMode: "center", tooltip: "alignCenterH", alignKeys: ["horizontalCenter"] },
            { iconType: "vertical",   alignMode: "center", tooltip: "alignCenterV", alignKeys: ["verticalCenter"] }
        ];

        /* ショートカットから引くためのボタン定義の一覧
           端に寄せる整列は十字の外周ボタンが受け持つので、整列ボタンは中央揃えの3つだけになる
           Every align button definition in one list; the cross handles the edge alignments, so only the
           three centred buttons remain */
        var ALL_ALIGN_BUTTONS = CENTER_ALIGN_BUTTONS.concat([CENTER_ALIGN_BUTTON]);

        /* 移動ボタンの並び（中央は移動ではなく、水平・垂直方向中央に整列するボタン）
           Layout of the move buttons; the middle one aligns on both axes instead of moving */
        var MOVE_BUTTON_ROWS = [
            ["upLeft",   "up",     "upRight"],
            ["left",     "center", "right"],
            ["downLeft", "down",   "downRight"]
        ];

        /* 値を持つ辺（設定の読み書きはこの並びで回す）/ The sides that hold a value */
        var MARGIN_SIDES = ["top", "bottom", "left", "right"];

        /* 分割ガイドの入力欄（設定の読み書きはこの並びで回す。ラベルのキーも兼ねる）
           The division fields; the settings are read and written in this order, and the keys double as label keys */
        var DIVIDE_VALUE_KEYS = ["rows", "columns", "rowGutter", "columnGutter", "extension"];

        /* 分割数と、その間隔を横に並べる2行（伸張はこの下に単独で置く）
           The two rows pairing a count with its gutter; the extension sits below them on its own */
        var DIVIDE_FIELD_ROWS = [
            { count: "rows",    gutter: "rowGutter" },
            { count: "columns", gutter: "columnGutter" }
        ];

        /* 分割ガイドのモード（ラジオボタンの並び）/ The division modes, in the order the radio buttons appear */
        var DIVIDE_MODES = [
            { key: "none",   labelKey: "divideNone" },
            { key: "cross",  labelKey: "divideCross" },
            { key: "custom", labelKey: "divideCustom" }
        ];

        /* 数値欄のツールチップのキー（項目名のキーから引く）/ Tooltip keys, looked up by the field key */
        var DIVIDE_TOOLTIP_KEYS = {
            rows:         "divideRows",
            columns:      "divideColumns",
            rowGutter:    "divideRowGutter",
            columnGutter: "divideColumnGutter",
            extension:    "divideExtension"
        };

        /* マージン欄の並び（3×3、真ん中は連動のリンクアイコン、"" は位置合わせの空セル）
           The 3x3 margin grid: the link icon sits in the middle, "" is an empty spacer cell */
        var MARGIN_FIELD_ROWS = [
            ["",     "top",    ""],
            ["left", "link",   "right"],
            ["",     "bottom", ""]
        ];

        /* 方向キーと、寄せる辺の対応（斜めは2辺へ同時に寄せる）
           The edges each direction moves to; a diagonal moves to both of them at once */
        var MOVE_SIDES_BY_DIRECTION = {
            up:        ["top"],
            left:      ["left"],
            right:     ["right"],
            down:      ["bottom"],
            upLeft:    ["left", "top"],
            upRight:   ["right", "top"],
            downLeft:  ["left", "bottom"],
            downRight: ["right", "bottom"]
        };

        /* 文字キーで実行する整列（値は整列ボタンの定義の tooltip キー）
           C は左右中央（Center）、M は上下中央（Middle）、X は上下左右中央
           The letter keys that run an align: C centres horizontally, M vertically, X on both axes */
        var ALIGN_KEY_SHORTCUTS = { C: "alignCenterH", M: "alignCenterV", X: "alignCenterAll" };

        /* ［裁ち落としに整列］を切り替えるキー / The key that toggles Align to Bleed */
        var BLEED_TOGGLE_KEY = "B";

        /* 矢印キーで実行する方向ボタン（斜めはキーが無いので4方向だけ）
           The arrow keys that fire a move button; the diagonals have no key of their own */
        var MOVE_DIRECTION_BY_KEY_NAME = { Up: "up", Left: "left", Right: "right", Down: "down" };

        /* 方向ボタンのアイコンで、寄せ先を表すケイ線を引く辺（斜めは2辺）
           The edges that get a rule in each move-button icon; a diagonal gets two */
        var MOVE_ICON_RULES = {
            up:        ["top"],
            down:      ["bottom"],
            left:      ["left"],
            right:     ["right"],
            upLeft:    ["top", "left"],
            upRight:   ["top", "right"],
            downLeft:  ["bottom", "left"],
            downRight: ["bottom", "right"]
        };

        /**
         * 整列アイコンボタンを1つ生成する
         * @param {Group} parentRow - 追加先の行グループ
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {void}
         */
        function addAlignButton(parentRow, buttonDef) {
            /* iconbutton ではなく button を使う（画像なしの iconbutton はクリックが届かないことがある）
               Use button, not iconbutton: an image-less iconbutton does not always receive clicks */
            var alignButton = parentRow.add("button", undefined, "");
            fixControlSize(alignButton, ICON_SIZE, ICON_SIZE);
            /* キー操作はラベルに出さず helpTip に書く / Shortcuts belong in the tooltip, not the label */
            alignButton.helpTip = getLabel("tooltip." + buttonDef.tooltip) + shortcutSuffix(buttonDef) + "  —  " +
                getLabel("tooltip." + optionTooltipKey(buttonDef));
            alignButton.iconType = buttonDef.iconType;
            alignButton.alignMode = buttonDef.alignMode;
            alignButton.isHover = false;
            alignButton.onDraw = function() { drawAlignButton(this); };
            alignButton.onClick = function() { runAlignButton(buttonDef); };
            attachHover(alignButton);
        }

        /**
         * そのボタンに割り当てたショートカットを「（C）」の形で返す（無ければ空文字）
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {string} helpTip に添える文字列
         */
        function shortcutSuffix(buttonDef) {
            for (var keyName in ALIGN_KEY_SHORTCUTS) {
                if (!ALIGN_KEY_SHORTCUTS.hasOwnProperty(keyName)) { continue; }
                if (ALIGN_KEY_SHORTCUTS[keyName] === buttonDef.tooltip) { return keyHint(keyName); }
            }
            return "";
        }

        /**
         * ショートカットのキーを「（B）」の形にする（日本語は全角括弧、英語は半角）
         * @param {string} keyName - キー名
         * @returns {string} helpTip に添える文字列
         */
        function keyHint(keyName) {
            return (uiLang === "ja") ? "（" + keyName + "）" : " (" + keyName + ")";
        }

        /**
         * パレットのイベント（選択の取り直し・キー操作・閉じるときの後始末）を結線する
         * @param {Window} alignPalette - 対象のパレット
         * @returns {void}
         */
        function attachPaletteEvents(alignPalette) {
            /* 選択の変化はパレットへ操作しに来た瞬間に拾う（Illustrator にタイマーAPIが無いため）
               Illustrator has no timer API, so the selection is re-read when the user comes to the palette */
            alignPalette.onActivate = function() { onPaletteFocus(true); };
            alignPalette.addEventListener("mouseover", function() { onPaletteFocus(false); });

            bindPaletteShortcuts(alignPalette);

            /* 閉じるとき：ガイドはドキュメントに残したまま、参照だけ解放する
               （消したいときは［ガイドを追加］を外すか、分割を［なし］にする）
               On close: the guides stay in the document and only the reference is released */
            alignPalette.onClose = function() {
                saveWindowLocation(alignPalette);
                paletteWindow = null;
                $.global.__aiAlignToArtboardWindow = null;
                return true;
            };
        }

        /**
         * パレット上のキー操作を登録する
         * Esc で閉じ、↑↓←→ で方向ボタン、C/M/X で中央揃え、B で裁ち落としの切り替えを行う
         * @param {Window} alignPalette - 対象のパレット
         * @returns {void}
         */
        function bindPaletteShortcuts(alignPalette) {
            var shortcutMap = {
                "Escape": { target: function() { alignPalette.close(); }, inFields: true }
            };
            shortcutMap[BLEED_TOGGLE_KEY] = function() {
                /* 無いときは文字をそのまま通す / Pass the key through when there is no checkbox */
                return (alignToBleedCheckbox !== null) ? alignToBleedCheckbox : false;
            };
            var keyName;
            for (keyName in MOVE_DIRECTION_BY_KEY_NAME) {
                if (!MOVE_DIRECTION_BY_KEY_NAME.hasOwnProperty(keyName)) { continue; }
                shortcutMap[keyName] = makeMoveShortcut(MOVE_DIRECTION_BY_KEY_NAME[keyName]);
            }
            for (keyName in ALIGN_KEY_SHORTCUTS) {
                if (!ALIGN_KEY_SHORTCUTS.hasOwnProperty(keyName)) { continue; }
                shortcutMap[keyName] = makeAlignShortcut(ALIGN_KEY_SHORTCUTS[keyName]);
            }

            /* 数値欄にフォーカスがあっても文字キーは効かせる（↑↓←→は欄の増減・カーソル移動に残す）
               Letter keys also work in the number fields; the arrows stay with the fields */
            var numericFields = [];
            for (keyName in marginFields) {
                if (marginFields.hasOwnProperty(keyName)) { numericFields.push(marginFields[keyName]); }
            }
            for (keyName in divideFields) {
                if (divideFields.hasOwnProperty(keyName)) { numericFields.push(divideFields[keyName]); }
            }
            addKeyShortcuts(alignPalette, shortcutMap, { numericFields: numericFields });
        }

        /**
         * 矢印キーで方向ボタンを実行するショートカットを作る（入力欄では何もしない）
         * @param {string} directionKey - MOVE_SIDES_BY_DIRECTION のキー
         * @returns {Function} addKeyShortcuts に渡す関数
         */
        function makeMoveShortcut(directionKey) {
            return function(keyEvent) {
                /* 入力欄の↑↓は値の増減、←→はカーソル移動 / In a field the arrows step the value or move the caret */
                if (keyEvent && keyEvent.target && keyEvent.target.type === "edittext") { return false; }
                runMoveButton(directionKey);
                return true;
            };
        }

        /**
         * 文字キーで中央揃えの整列ボタンを実行するショートカットを作る
         * @param {string} tooltipKey - 整列ボタンの定義の tooltip
         * @returns {Function} addKeyShortcuts に渡す関数
         */
        function makeAlignShortcut(tooltipKey) {
            return function() {
                var buttonDef = findAlignButtonDef(tooltipKey);
                if (buttonDef === null) { return false; }
                runAlignButton(buttonDef);
                return true;
            };
        }

        /**
         * 整列ボタンを1つ実行する（クリックとショートカットで共通）
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {void}
         */
        function runAlignButton(buttonDef) {
            var workerResult = null;
            var didRun = runExclusive(function() { workerResult = runAlign(buttonDef); });
            /* 実行後の選択に合わせてディムと単位を更新してから、今回の結果を表示する
               （先に表示すると、この更新で選択が変わったと見なされて消えてしまう）
               Refresh first, then show this run's result; showing it first would be wiped by the refresh */
            onPaletteFocus(true);
            if (didRun) { showWorkerResult(workerResult); }
        }

        /**
         * tooltip キーから整列ボタンの定義を探す
         * @param {string} tooltipKey - 整列ボタンの定義の tooltip
         * @returns {object} 見つかったボタン定義。無ければ null
         */
        function findAlignButtonDef(tooltipKey) {
            for (var i = 0; i < ALL_ALIGN_BUTTONS.length; i++) {
                if (ALL_ALIGN_BUTTONS[i].tooltip === tooltipKey) { return ALL_ALIGN_BUTTONS[i]; }
            }
            return null;
        }

        /**
         * 移動ボタンを1つ生成する（directionKey が "center" なら中央揃えの整列ボタン）
         * @param {Group} parentRow - 追加先の行グループ
         * @param {string} directionKey - MOVE_SIDES_BY_DIRECTION のキー（"center" は十字の中央）
         * @returns {void}
         */
        function addMoveButton(parentRow, directionKey) {
            /* 十字の中央は、整列アイコンの列と同じ「水平・垂直方向中央に整列」にする
               The middle of the cross is the same centre-on-both-axes align button as in the icon row */
            if (directionKey === "center") {
                addAlignButton(parentRow, CENTER_ALIGN_BUTTON);
                return;
            }
            /* 整列アイコンと同じ理由で iconbutton ではなく button を使う / Same reason as the align icons: button, not iconbutton */
            var moveButton = parentRow.add("button", undefined, "");
            fixControlSize(moveButton, ICON_SIZE, ICON_SIZE);
            moveButton.helpTip = getLabel("direction." + directionKey) + "  —  " + getLabel("tooltip.moveToEdge");
            moveButton.directionKey = directionKey;
            moveButton.isHover = false;
            moveButton.onDraw = function() { drawMoveButton(this); };
            moveButton.onClick = function() { runMoveButton(directionKey); };
            attachHover(moveButton);
        }

        /**
         * 方向ボタンを1つ実行する（クリックと矢印キーで共通）
         * @param {string} directionKey - MOVE_SIDES_BY_DIRECTION のキー
         * @returns {void}
         */
        function runMoveButton(directionKey) {
            var workerResult = null;
            var didRun = runExclusive(function() { workerResult = runMove(directionKey); });
            /* 整列ボタンと同じ順序：先に選択を取り直してから今回の結果を出す
               Same order as the align buttons: refresh first, then show this run's result */
            onPaletteFocus(true);
            if (didRun) { showWorkerResult(workerResult); }
        }

        /**
         * 移動ボタンを十字に並べる
         * @param {Group} parentColumn - 追加先の行グループ（2カラムの左側）
         * @returns {void}
         */
        function addMoveButtonCross(parentColumn) {
            var crossGroup = parentColumn.add("group");
            crossGroup.orientation = "column";
            /* 右カラムの高さに引き伸ばさず、真ん中に置く / Centered in the column instead of stretched to its height */
            crossGroup.alignment = ["center", "center"];
            crossGroup.alignChildren = ["center", "top"];
            crossGroup.spacing = CROSS_GAP;

            for (var i = 0; i < MOVE_BUTTON_ROWS.length; i++) {
                var crossRow = crossGroup.add("group");
                setupRow(crossRow, "center", CROSS_GAP);
                for (var j = 0; j < MOVE_BUTTON_ROWS[i].length; j++) {
                    addMoveButton(crossRow, MOVE_BUTTON_ROWS[i][j]);
                }
            }

            addCenterAlignRow(crossGroup);
        }

        /**
         * 十字の下に、片方の軸だけ中央に整列するボタン2つを並べる
         * @param {Group} crossGroup - 追加先の十字のグループ
         * @returns {void}
         */
        function addCenterAlignRow(crossGroup) {
            var centerRow = crossGroup.add("group");
            setupRow(centerRow, "center", CROSS_GAP);
            /* 十字と続けて見えないよう、上に少し余白を取る / A little room so it does not read as a fourth row of the cross */
            centerRow.margins = [0, CENTER_ROW_TOP, 0, 0];
            for (var i = 0; i < CENTER_ALIGN_BUTTONS.length; i++) {
                addAlignButton(centerRow, CENTER_ALIGN_BUTTONS[i]);
            }
        }

        /**
         * パレットの上段を2カラムに分け、左にボタンの十字・右にオプションのパネルを置く
         * マージンガイドと分割ガイドはここには入れず、その下に幅いっぱいで並べる
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addColumnsRow(targetWindow) {
            var columnsRow = targetWindow.add("group");
            columnsRow.orientation = "row";
            columnsRow.alignment = ["fill", "top"];
            columnsRow.alignChildren = ["fill", "fill"];
            columnsRow.spacing = COLUMN_SPACING;

            addMoveButtonCross(columnsRow);
            addOptionsPanel(columnsRow);
        }

        /**
         * マージンガイドのパネル（上下左右の数値欄＋定規の単位ラベル）を組み立てる
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addMarginPanel(targetWindow) {
            marginPanel = targetWindow.add("panel", undefined, panelTitleWithUnit("guide"));
            marginPanel.helpTip = getLabel("tooltip.panelGuide");
            setupPanel(marginPanel, OPTION_SPACING);

            /* 前回の設定を引き継ぐ（残したガイドと表示が食い違わないように）/ Carry over the previous settings */
            var paletteSettings = loadPaletteSettings();
            marginsAreDefault = paletteSettings.usesDefaultMargins;

            showGuideCheckbox = marginPanel.add("checkbox", undefined, getLabel("checkbox.showGuide"));
            showGuideCheckbox.helpTip = getLabel("tooltip.showGuide");
            showGuideCheckbox.value = paletteSettings.showGuide;
            showGuideCheckbox.onClick = function() {
                /* ガイドを出すときは四辺をそろえたいことが多いので、連動も入れて4つの値をそろえる
                   Turning the guide on usually means an even inset, so the link goes on with it */
                if (showGuideCheckbox.value === true && linkMarginsToggle !== null && linkMarginsToggle.value !== true) {
                    setLinkToggleValue(linkMarginsToggle, true);
                    copyMarginToLinkedFields(marginFields.top);
                }
                syncMarginFields();
                savePaletteSettings();
                runExclusive(refreshMarginGuide);
            };

            addMarginFieldRows(marginPanel, paletteSettings);
        }

        /**
         * パネル名に、いまの定規の単位を添える
         * 数値欄が多いので、単位は欄ごとではなくパネル名にまとめて出す
         * @param {string} panelKey - LABELS.panel のキー
         * @returns {string} パネル名
         */
        function panelTitleWithUnit(panelKey) {
            var unitLabel = currentUnitInfo.label;
            return getLabel("panel." + panelKey) + (uiLang === "ja" ? "（" + unitLabel + "）" : " (" + unitLabel + ")");
        }

        /**
         * 上下左右のマージン欄を3×3で組み立てる
         * @param {Panel} marginPanel - 追加先のパネル
         * @param {object} paletteSettings - 引き継いだ設定
         * @returns {void}
         */
        function addMarginFieldRows(marginPanel, paletteSettings) {
            for (var i = 0; i < MARGIN_FIELD_ROWS.length; i++) {
                var fieldRow = marginPanel.add("group");
                fieldRow.orientation = "row";
                /* パネルの幅いっぱいに広げず、真ん中に置く / Centred in the panel instead of stretched */
                fieldRow.alignment = ["center", "top"];
                fieldRow.alignChildren = ["center", "center"];
                fieldRow.spacing = LABEL_FIELD_SPACING;
                for (var j = 0; j < MARGIN_FIELD_ROWS[i].length; j++) {
                    addMarginCell(fieldRow, MARGIN_FIELD_ROWS[i][j], paletteSettings);
                }
            }
            syncMarginFields();
        }

        /**
         * ［ガイドを追加］と［連動］の状態に合わせてマージン欄の使える・使えないを切り替える
         * ガイドを出していないあいだは4辺まとめてディムにし、連動しているあいだは［上］の1つで4辺が決まるので、ほかをディムにする
         * @returns {void}
         */
        function syncMarginFields() {
            var addsGuide, isLinked, side, i;
            if (linkMarginsToggle === null) { return; }
            addsGuide = showGuideCheckbox !== null && showGuideCheckbox.value === true;
            isLinked = linkMarginsToggle.value === true;
            setLinkToggleEnabled(linkMarginsToggle, addsGuide);
            for (i = 0; i < MARGIN_SIDES.length; i++) {
                side = MARGIN_SIDES[i];
                if (!marginFields[side]) { continue; }
                setSteppedInputEnabled(marginFields[side], addsGuide && (isLinked ? (side === "top") : true));
            }
        }

        /**
         * 3×3の1セルを生成する（辺なら「ラベル＋入力欄」、"link" なら連動のリンクアイコン、"" なら空セル）
         * どのセルも同じ幅にして、上下左右が十字に並ぶようにする
         * @param {Group} fieldRow - 追加先の行グループ
         * @param {string} cellKey - "top" / "bottom" / "left" / "right" / "link" / ""
         * @param {object} paletteSettings - 引き継いだ設定
         * @returns {void}
         */
        function addMarginCell(fieldRow, cellKey, paletteSettings) {
            var cellGroup = fieldRow.add("group");
            cellGroup.orientation = "row";
            cellGroup.alignment = ["center", "center"];
            cellGroup.alignChildren = ["center", "center"];
            cellGroup.spacing = LABEL_FIELD_SPACING;
            /* 入力欄の左に∧∨が入るぶん、どのセルも同じだけ広げて十字をそろえる
               Widen every cell by the stepper so the cross stays aligned */
            cellGroup.minimumSize.width = MARGIN_CELL_WIDTH[uiLang] + STEPPER_BUTTON_WIDTH + STEPPER_SIDE_MARGIN;
            if (cellKey === "") { return; }

            if (cellKey === "link") {
                /* 連動に切り替えた時点で、4つを上のマージンにそろえる
                   Turning the link back on levels the four values off against the top margin */
                linkMarginsToggle = addLinkToggle(cellGroup, paletteSettings.linkMargins, function() {
                    copyMarginToLinkedFields(marginFields.top);
                    syncMarginFields();
                    runExclusive(refreshMarginGuide);
                    savePaletteSettings();
                });
                linkMarginsToggle.helpTip = getLabel("tooltip.linkMargins");
                return;
            }

            cellGroup.add("statictext", undefined, labelText("fieldLabel." + cellKey));
            var marginField = addSteppedInput(cellGroup, paletteSettings.margins[cellKey], { min: 0 });
            marginField.helpTip = getLabel("tooltip.margin");
            /* 確定（Enter・フォーカス移動）でガイドを描き直す
               入力途中で毎回描き直すとそのつど委譲が走るため、描き直しは onChange だけにする
               Only redraw the guide when the field commits, not on every keystroke */
            marginField.onChange = function() {
                /* 一度でも触られたら、単位が変わっても既定値では上書きしない */
                marginsAreDefault = false;
                copyMarginToLinkedFields(marginField);
                runExclusive(refreshMarginGuide);
                savePaletteSettings();
            };
            marginFields[cellKey] = marginField;
        }

        /**
         * まだ触られていないマージン欄を、いまの定規の単位の既定値で埋め直す
         * パネルを組み立てる時点では定規の単位が分からないため、分かった時点で入れ直す
         * @returns {void}
         */
        function fillDefaultMargins() {
            var defaultText, side, i;
            if (!marginsAreDefault) { return; }
            defaultText = String(currentUnitInfo.defaultMargin);
            for (i = 0; i < MARGIN_SIDES.length; i++) {
                side = MARGIN_SIDES[i];
                if (!marginFields[side]) { continue; }
                marginFields[side].text = defaultText;
            }
            savePaletteSettings();
        }

        /**
         * 伸張欄がまだ既定値のままなら、いまの定規の単位の既定値に入れ直す
         * @returns {void}
         */
        function fillDefaultExtension() {
            if (!extensionIsDefault || !divideFields.extension) { return; }
            divideFields.extension.text = String(currentUnitInfo.defaultExtension);
            savePaletteSettings();
        }

        /**
         * ［連動］がONのとき、編集した欄の値を残りのマージン欄へ写す
         * @param {EditText} sourceField - 値の元にする入力欄
         * @returns {void}
         */
        function copyMarginToLinkedFields(sourceField) {
            if (linkMarginsToggle === null || linkMarginsToggle.value !== true) { return; }
            if (!sourceField) { return; }
            for (var side in marginFields) {
                if (!marginFields.hasOwnProperty(side)) { continue; }
                if (marginFields[side] === sourceField) { continue; }
                marginFields[side].text = sourceField.text;
            }
        }

        /**
         * 分割ガイドのパネル（なし・十字・カスタムの切り替えと、行・列・行間・列間・伸張）を組み立てる
         * マージンの内側を等分する位置にガイドを引き、伸張のぶんだけアートボードの外へ伸ばす
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addDividePanel(targetWindow) {
            var paletteSettings, edgeRow, extensionRow, i;

            /* 前回の設定を引き継ぐ（残したガイドと表示が食い違わないように）/ Carry over the previous settings */
            paletteSettings = loadPaletteSettings();
            extensionIsDefault = paletteSettings.usesDefaultExtension;

            dividePanel = targetWindow.add("panel", undefined, panelTitleWithUnit("divide"));
            dividePanel.helpTip = getLabel("tooltip.panelDivide");
            setupPanel(dividePanel, OPTION_SPACING);

            addDivideModeRow(dividePanel, paletteSettings);
            for (i = 0; i < DIVIDE_FIELD_ROWS.length; i++) {
                addDivideFieldRow(dividePanel, DIVIDE_FIELD_ROWS[i], paletteSettings);
            }

            /* 分割の指定とは別の設定なので、上に少し余白を取って切り分ける
               It is separate from the division settings, so a little room above sets it apart */
            edgeRow = dividePanel.add("group");
            setupRow(edgeRow, "left", LABEL_FIELD_SPACING);
            edgeRow.margins = [0, EDGE_ROW_TOP, 0, 0];

            artboardEdgeCheckbox = edgeRow.add("checkbox", undefined, getLabel("checkbox.artboardEdge"));
            artboardEdgeCheckbox.helpTip = getLabel("tooltip.artboardEdge");
            artboardEdgeCheckbox.value = paletteSettings.artboardEdge;
            artboardEdgeCheckbox.onClick = function() {
                syncDivideFields();
                runExclusive(refreshMarginGuide);
            };

            /* 伸張は行・列とアートボードのエッジのどれにも掛かるので、対にせず、すべての下に単独で置く
               The extension applies to the rows, the columns and the artboard edges alike, so it sits on
               its own row below all of them */
            extensionRow = dividePanel.add("group");
            setupRow(extensionRow, "left", LABEL_FIELD_SPACING);
            addDivideField(extensionRow, "extension", paletteSettings, DIVIDE_LABEL_WIDTH[uiLang]);

            syncDivideFields();
        }

        /**
         * 分割の仕方を選ぶラジオボタンの行（なし・十字・カスタム）を組み立てる
         * @param {Panel} dividePanel - 追加先のパネル
         * @param {object} paletteSettings - 引き継いだ設定
         * @returns {void}
         */
        function addDivideModeRow(dividePanel, paletteSettings) {
            var modeRow, modeDef, modeRadio, i;
            modeRow = dividePanel.add("group");
            setupRow(modeRow, "left", RADIO_GAP);
            for (i = 0; i < DIVIDE_MODES.length; i++) {
                modeDef = DIVIDE_MODES[i];
                modeRadio = modeRow.add("radiobutton", undefined, getLabel("radio." + modeDef.labelKey));
                modeRadio.helpTip = getLabel("tooltip." + modeDef.labelKey);
                modeRadio.value = (modeDef.key === paletteSettings.divideMode);
                modeRadio.onClick = function() {
                    syncDivideFields();
                    runExclusive(refreshMarginGuide);
                };
                divideRadios[modeDef.key] = modeRadio;
            }
        }

        /**
         * 分割数と、その間隔を横に並べた1行を組み立てる
         * @param {Panel} dividePanel - 追加先のパネル
         * @param {object} rowDef - DIVIDE_FIELD_ROWS の1行
         * @param {object} paletteSettings - 引き継いだ設定
         * @returns {void}
         */
        function addDivideFieldRow(dividePanel, rowDef, paletteSettings) {
            var fieldRow = dividePanel.add("group");
            setupRow(fieldRow, "left", LABEL_FIELD_SPACING);
            addDivideField(fieldRow, rowDef.count, paletteSettings, DIVIDE_LABEL_WIDTH[uiLang]);
            addDivideField(fieldRow, rowDef.gutter, paletteSettings, DIVIDE_GUTTER_LABEL_WIDTH[uiLang]);
        }

        /**
         * 分割ガイドの「項目名＋数値欄」を1組ぶん作る
         * 項目名は右そろえにして、日英で長さが違っても数値欄の頭がそろうようにする
         * @param {Group} fieldRow - 追加先の行グループ
         * @param {string} valueKey - DIVIDE_VALUE_KEYS のキー
         * @param {object} paletteSettings - 引き継いだ設定
         * @param {number} labelWidth - 項目名の幅
         * @returns {void}
         */
        function addDivideField(fieldRow, valueKey, paletteSettings, labelWidth) {
            var divideLabel, divideField;

            divideLabel = fieldRow.add("statictext", undefined, labelText("fieldLabel." + valueKey));
            divideLabel.preferredSize.width = labelWidth;
            divideLabel.justify = "right";

            /* 分割数は1〜DIVIDE_COUNT_MAXの整数、間隔・伸張は0以上 / counts are whole numbers from 1; lengths are non-negative */
            var isCountField = (valueKey === "rows" || valueKey === "columns");
            divideField = addSteppedInput(fieldRow, paletteSettings.divideValues[valueKey],
                isCountField ? { integer: true, min: 1, max: DIVIDE_COUNT_MAX } : { min: 0 });
            divideField.helpTip = getLabel("tooltip." + DIVIDE_TOOLTIP_KEYS[valueKey]);
            /* 確定（Enter・フォーカス移動）でガイドを描き直す / Only redraw the guide when the field commits */
            divideField.onChange = function() {
                /* 一度でも触られたら、単位が変わっても既定値では上書きしない */
                if (valueKey === "extension") { extensionIsDefault = false; }
                runExclusive(refreshMarginGuide);
                savePaletteSettings();
            };
            divideFields[valueKey] = divideField;
        }

        /**
         * 分割の仕方に合わせて数値欄のディムを切り替える
         * ［なし］はすべて、［十字］は行・列と間隔をディムにする（伸張は十字にも効く）
         * @returns {void}
         */
        function syncDivideFields() {
            var divideMode, isCustom, usesExtension, valueKey, i;
            divideMode = readDivideMode();
            isCustom = (divideMode === "custom");
            /* 伸張は分割のガイドとアートボードのエッジの両方に掛かるので、どちらかがあれば使える
               The extension applies to the division guides and to the artboard edges alike */
            usesExtension = (divideMode !== "none") || readArtboardEdge();
            for (i = 0; i < DIVIDE_VALUE_KEYS.length; i++) {
                valueKey = DIVIDE_VALUE_KEYS[i];
                if (!divideFields[valueKey]) { continue; }
                setSteppedInputEnabled(divideFields[valueKey], (valueKey === "extension") ? usesExtension : isCustom);
            }
            savePaletteSettings();
        }

        /**
         * ［アートボードのエッジ］が入っているかを返す
         * @returns {boolean} 入っていれば true
         */
        function readArtboardEdge() {
            return artboardEdgeCheckbox !== null && artboardEdgeCheckbox.value === true;
        }

        /**
         * いま選ばれている分割の仕方を返す
         * @returns {string} "none" / "cross" / "custom"
         */
        function readDivideMode() {
            var divideMode, i;
            for (i = 0; i < DIVIDE_MODES.length; i++) {
                divideMode = DIVIDE_MODES[i].key;
                if (divideRadios[divideMode] && divideRadios[divideMode].value === true) { return divideMode; }
            }
            return DEFAULT_DIVIDE_MODE;
        }

        /**
         * ガイドを描くものがあるか（マージンのガイドか分割ガイドのどちらかが有効なら true）
         * @returns {boolean} 描くものがあれば true
         */
        function hasGuideToDraw() {
            if (showGuideCheckbox !== null && showGuideCheckbox.value === true) { return true; }
            if (readArtboardEdge()) { return true; }
            return readDivideMode() !== "none";
        }

        /**
         * 前回のパレット設定を常駐エンジンから読み出す（控えが無ければ既定値）
         * @returns {object} { showGuide, usesDefaultMargins, margins, divideMode, divideValues, usesDefaultExtension, artboardEdge, linkMargins, alignToBleed, perArtboard }
         */
        function loadPaletteSettings() {
            /* 項目ごとの確認は下で行うため、既定値は {}（中身を問わず受け取る）/ Each field is checked below, so the default is {} (accepts any content) */
            var storedSettings = settingsStore.load({});
            var hasStoredSettings = false;
            for (var storedKey in storedSettings) {
                if (storedSettings.hasOwnProperty(storedKey)) { hasStoredSettings = true; break; }
            }
            if (!hasStoredSettings) {
                return {
                    showGuide:    DEFAULT_SHOW_GUIDE,
                    usesDefaultMargins: true,
                    margins:      marginStrings(null, String(currentUnitInfo.defaultMargin)),
                    divideMode:   DEFAULT_DIVIDE_MODE,
                    divideValues: divideValueStrings(null),
                    usesDefaultExtension: true,
                    artboardEdge: DEFAULT_ARTBOARD_EDGE,
                    linkMargins:  DEFAULT_LINK_MARGINS,
                    alignToBleed: DEFAULT_ALIGN_TO_BLEED,
                    perArtboard:  DEFAULT_ALIGN_PER_ARTBOARD
                };
            }
            return {
                showGuide:    storedSettings.showGuide === true,
                usesDefaultMargins: storedSettings.margins == null,
                margins:      marginStrings(storedSettings.margins, String(currentUnitInfo.defaultMargin)),
                divideMode:   knownDivideMode(storedSettings.divideMode),
                divideValues: divideValueStrings(storedSettings.divideValues),
                usesDefaultExtension: !(storedSettings.divideValues && storedSettings.divideValues.extension != null),
                artboardEdge: storedSettings.artboardEdge === true,
                linkMargins:  (storedSettings.linkMargins != null) ? (storedSettings.linkMargins === true) : DEFAULT_LINK_MARGINS,
                alignToBleed: (storedSettings.alignToBleed != null) ? (storedSettings.alignToBleed === true) : DEFAULT_ALIGN_TO_BLEED,
                perArtboard:  storedSettings.perArtboard === true
            };
        }

        /**
         * 控えのマージン4値を、欠けている辺を既定値で埋めた文字列の組にそろえる
         * @param {object} storedMargins - 控えの値（無ければ null）
         * @param {string} fallback - 欠けている辺に使う値
         * @returns {object} { top: string, bottom: string, left: string, right: string }
         */
        function marginStrings(storedMargins, fallback) {
            var marginTexts, side, i;
            marginTexts = {};
            for (i = 0; i < MARGIN_SIDES.length; i++) {
                side = MARGIN_SIDES[i];
                marginTexts[side] = (storedMargins && storedMargins[side] != null) ? String(storedMargins[side]) : fallback;
            }
            return marginTexts;
        }

        /**
         * 控えの分割の仕方が並びにあるものか確かめる（無ければ既定値）
         * @param {string} storedMode - 控えの値
         * @returns {string} "none" / "cross" / "custom"
         */
        function knownDivideMode(storedMode) {
            for (var i = 0; i < DIVIDE_MODES.length; i++) {
                if (DIVIDE_MODES[i].key === storedMode) { return storedMode; }
            }
            return DEFAULT_DIVIDE_MODE;
        }

        /**
         * 控えの分割ガイドの値を、欠けている項目を既定値で埋めた文字列の組にそろえる
         * @param {object} storedValues - 控えの値（無ければ null）
         * @returns {object} DIVIDE_VALUE_KEYS をキーにした文字列の組
         */
        function divideValueStrings(storedValues) {
            var divideTexts, valueKey, i;
            divideTexts = {};
            for (i = 0; i < DIVIDE_VALUE_KEYS.length; i++) {
                valueKey = DIVIDE_VALUE_KEYS[i];
                divideTexts[valueKey] = (storedValues && storedValues[valueKey] != null) ?
                    String(storedValues[valueKey]) : defaultDivideValue(valueKey);
            }
            return divideTexts;
        }

        /**
         * 分割ガイドの欄の初期値を返す（伸張だけは定規の単位ごとに決まる）
         * @param {string} valueKey - DIVIDE_VALUE_KEYS のキー
         * @returns {string} 初期値
         */
        function defaultDivideValue(valueKey) {
            if (valueKey === "extension") { return String(currentUnitInfo.defaultExtension); }
            return String(DEFAULT_DIVIDE_VALUES[valueKey]);
        }

        /**
         * 現在のパレット設定を常駐エンジンに控える
         * @returns {void}
         */
        function savePaletteSettings() {
            if (!isPaletteReady) { return; }
            settingsStore.save({
                showGuide:    showGuideCheckbox !== null && showGuideCheckbox.value === true,
                margins:      readMarginTexts(),
                divideMode:   readDivideMode(),
                divideValues: readDivideValueTexts(),
                artboardEdge: readArtboardEdge(),
                linkMargins:  linkMarginsToggle !== null && linkMarginsToggle.value === true,
                alignToBleed: alignToBleedCheckbox !== null && alignToBleedCheckbox.value === true,
                perArtboard:  alignPerArtboardCheckbox !== null && alignPerArtboardCheckbox.value === true
            });
        }

        /**
         * オプションパネル（境界の測り方・行揃えの変更・寄せ先の指定）を組み立てる
         * @param {Group} parentRow - 追加先の2カラムの行グループ
         * @returns {void}
         */
        function addOptionsPanel(parentRow) {
            var paletteSettings = loadPaletteSettings();

            var optionsPanel = parentRow.add("panel", undefined, getLabel("panel.options"));
            setupPanel(optionsPanel, OPTION_SPACING);

            /* ここでは控えの値を置くだけで、実際の初期値は loadBoundsPreferences() が入れ直す
               （常駐パレットの app は当てにならないため、環境設定はメインエンジンへ問い合わせる）
               Only the fallbacks are set here; loadBoundsPreferences() replaces them with the real preferences */
            previewBoundsCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
            previewBoundsCheckbox.helpTip = getLabel("tooltip.previewBounds");
            previewBoundsCheckbox.value = DEFAULT_PREVIEW_BOUNDS;

            /* 字形の境界はポイント文字・エリア内文字をまとめてON/OFFする / Glyph bounds toggles point & area type together */
            glyphBoundsCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.glyphBounds"));
            glyphBoundsCheckbox.helpTip = getLabel("tooltip.glyphBounds");
            glyphBoundsCheckbox.value = DEFAULT_GLYPH_BOUNDS;

            /* 行揃えはこのスクリプト内だけの設定 / Justification is script-local */
            changeJustificationCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.changeJustification"));
            changeJustificationCheckbox.helpTip = getLabel("tooltip.changeJustification");
            changeJustificationCheckbox.value = DEFAULT_CHANGE_JUSTIFICATION;

            opticalAdjustCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.opticalAdjust"));
            opticalAdjustCheckbox.helpTip = getLabel("tooltip.opticalAdjust");
            opticalAdjustCheckbox.value = DEFAULT_OPTICAL_ADJUST;

            /* 裁ち落としの量は BLEED_MM 固定なので、チェックボックスだけを置く
               The bleed amount is fixed at BLEED_MM, so only the checkbox is needed */
            alignToBleedCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.alignToBleed"));
            alignToBleedCheckbox.helpTip = getLabel("tooltip.alignToBleed").replace("{0}", BLEED_MM) + keyHint(BLEED_TOGGLE_KEY);
            alignToBleedCheckbox.value = paletteSettings.alignToBleed;
            alignToBleedCheckbox.onClick = function() { savePaletteSettings(); };

            alignPerArtboardCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.perArtboard"));
            alignPerArtboardCheckbox.helpTip = getLabel("tooltip.perArtboard");
            alignPerArtboardCheckbox.value = paletteSettings.perArtboard;
            alignPerArtboardCheckbox.onClick = function() {
                savePaletteSettings();
                /* ガイドを描く範囲（作業中のアートボードだけ／すべて）が変わるので描き直す */
                if (hasGuideToDraw()) { runExclusive(refreshMarginGuide); }
            };
        }

        /**
         * ［プレビュー境界］［字形の境界に整列］の現在の環境設定をチェックボックスへ反映する
         * 常駐パレットの app は当てにならないためメインエンジンへ問い合わせ、取れなければ DEFAULT_* のままにする
         * @returns {void}
         */
        function loadBoundsPreferences() {
            if (previewBoundsCheckbox === null || glyphBoundsCheckbox === null) { return; }
            var workerResult = runWorker("btGetPreferenceFlags();");
            if (workerResult === null || workerResult.indexOf("ERR:") === 0) { return; }
            var preferenceFlags = workerResult.split("|");
            if (preferenceFlags.length < 2) { return; }
            previewBoundsCheckbox.value = (preferenceFlags[0] === "1");
            glyphBoundsCheckbox.value = (preferenceFlags[1] === "1");
        }

        /**
         * 状況表示を書き換える（固定幅で切り詰められるため、全文は helpTip に入れる）
         * @param {string} message - 表示する文言
         * @returns {void}
         */
        function setStatus(message) {
            if (statusText === null) { return; }
            statusText.text = message;
            statusText.helpTip = message;
        }

        /**
         * 状況表示の行を組み立てる（中身で幅が変わらないよう固定幅にする）
         * @param {Window} targetWindow - 追加先のパレット
         * @returns {void}
         */
        function addStatusLine(targetWindow) {
            statusText = targetWindow.add("statictext", undefined, "", { truncate: "end" });
            statusText.alignment = ["fill", "center"];
            statusText.preferredSize.width = STATUS_WIDTH;
            statusText.maximumSize.width = STATUS_WIDTH;
        }

        /**
         * パレットを閉じた位置を控える（次に開いたときに同じ場所へ出すため）
         * @param {Window} alignPalette - 対象のパレット
         * @returns {void}
         */
        function saveWindowLocation(alignPalette) {
            $.global[WINDOW_LOCATION_KEY] = [alignPalette.location[0], alignPalette.location[1]];
        }

        /**
         * 控えておいた位置にパレットを移す（画面の外に出る位置なら使わない）
         * @param {Window} alignPalette - 対象のパレット
         * @returns {void}
         */
        function restoreWindowLocation(alignPalette) {
            var storedLocation = $.global[WINDOW_LOCATION_KEY];
            if (!storedLocation || storedLocation.length !== 2) { return; }
            if (!isLocationOnScreen(storedLocation[0], storedLocation[1])) { return; }
            try { alignPalette.location = [storedLocation[0], storedLocation[1]]; } catch (locationError) {}
        }

        /**
         * その座標がいずれかのディスプレイの内側かを判定する
         * ディスプレイ構成が変わったときに、掴めない位置へパレットを出さないために見る
         * @param {number} x - 左端
         * @param {number} y - 上端
         * @returns {boolean} 内側なら true（判定できない環境では true）
         */
        function isLocationOnScreen(x, y) {
            var screens, screenBounds, i;
            try {
                screens = $.screens;
            } catch (screenError) {
                return true;
            }
            if (!screens || screens.length === 0) { return true; }
            for (i = 0; i < screens.length; i++) {
                screenBounds = screens[i];
                if (x >= screenBounds.left && x <= screenBounds.right - WINDOW_ONSCREEN_MARGIN &&
                    y >= screenBounds.top && y <= screenBounds.bottom - WINDOW_ONSCREEN_MARGIN) { return true; }
            }
            return false;
        }

        /**
         * すでに開いているパレットがあれば閉じる（多重起動防止と、修正後のコードで開き直すため）
         * @returns {void}
         */
        function closeExistingPalette() {
            try {
                if (paletteWindow) { paletteWindow.close(); }
            } catch (staleReferenceError) {} /* 参照が無効なら閉じる必要もない / A stale reference needs no closing */
            paletteWindow = null;
            $.global.__aiAlignToArtboardWindow = null;
        }

        /**
         * 整列パレットを組み立てて表示する
         * @returns {void}
         */
        function showPalette() {
            /* 多重起動防止：開いているパレットは必ず閉じてから作り直す / Close any open palette first */
            closeExistingPalette();
            isPaletteReady = false;
            initIconColors();

            var alignPalette = new Window("palette", getLabel("dialog.title") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
            setupWindow(alignPalette);

            addColumnsRow(alignPalette);
            addMarginPanel(alignPalette);
            addDividePanel(alignPalette);
            addStatusLine(alignPalette);
            /* ここまでで全コントロールがそろうので、以降は操作のたびに控える
               Every control now exists, so from here on each change is stored */
            isPaletteReady = true;

            attachPaletteEvents(alignPalette);

            /* 常駐参照：GC 回避と多重起動の検出を兼ねる / Persistent reference: avoids GC and detects a second launch */
            paletteWindow = alignPalette;
            $.global.__aiAlignToArtboardWindow = alignPalette;
            /* 環境設定の現在値をチェックボックスに入れてから表示する（表示後だと目の前で切り替わる）
               Load the preferences before showing, so the checkboxes do not flip in front of the user */
            loadBoundsPreferences();
            alignPalette.layout.layout(true);
            restoreWindowLocation(alignPalette);
            alignPalette.show();
            refreshPaletteState();
            /* 引き継いだ設定でガイドを描き直す（前回残したガイドと入力欄の値をそろえる）
               Redraw the guide from the carried-over settings so it matches the field */
            if (hasGuideToDraw()) { runExclusive(refreshMarginGuide); }
        }

        // =========================================
        // BridgeTalk ワーカー / BridgeTalk workers
        // =========================================
        /* 以下の bt* 関数は toString() で連結してメインエンジンへ送るため、次の制約がある
           これらにJSDocを付けないのも同じ理由（出力が壊れて構文エラーになる）
           These bt* functions are stringified with toString() and shipped to the main engine, so:
           - 行コメント（//）は使わず、ブロックコメント（/* *\/）だけにする
           - 文は必ずセミコロンで終える（toString で改行が失われても壊れないように）
           - パレット側の変数は参照しない。必要な値は workerOptions で受け取る */

        function btAlignSelection(workerOptions) {
            var doc, selectedItems, alignResult;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            selectedItems = btResolveSelection(doc);
            if (selectedItems === null) { return "NOSEL"; }
            alignResult = btRunPerArtboard(doc, selectedItems, workerOptions, btAlignItems, "OK");
            if (alignResult.indexOf("OK") !== 0) { return alignResult; }
            /* ガイドは最後に1回だけ描く（アートボードごとに回すと、そのつど張り替えることになる）*/
            try {
                btDrawMarginGuide(doc, workerOptions);
            } catch (guideError) {
                return "ERR:" + guideError;
            }
            return alignResult;
        }

        function btResolveSelection(doc) {
            var selectedItems;
            selectedItems = doc.selection;
            /* 文字を選択しているときは、そのストーリーのテキストオブジェクトに置き換える */
            if (selectedItems && !(selectedItems instanceof Array)) {
                selectedItems = btPromoteTextRange(doc, selectedItems);
            }
            if (!(selectedItems instanceof Array) || selectedItems.length === 0) { return null; }
            return selectedItems;
        }

        function btSelectItems(doc, targetItems) {
            var i;
            doc.selection = null;
            for (i = 0; i < targetItems.length; i++) {
                try { targetItems[i].selected = true; } catch (selectError) {}
            }
        }

        function btRunPerArtboard(doc, selectedItems, workerOptions, runBucket, okMarker) {
            var buckets, bucketResult, lastResult, i;
            /* ［アートボードごとに整列］がOFF、またはアートボードが1つなら、まとめて1回で済ませる */
            if (workerOptions.perArtboard !== true || doc.artboards.length < 2) {
                return runBucket(doc, selectedItems, workerOptions);
            }
            buckets = btGroupItemsByArtboard(doc, selectedItems);
            lastResult = okMarker;
            for (i = 0; i < buckets.length; i++) {
                /* 整列も移動も選択を見て動くので、そのアートボードのぶんだけ選び直す */
                btSelectItems(doc, buckets[i]);
                bucketResult = runBucket(doc, buckets[i], workerOptions);
                /* 1つでも失敗したら、そこで止めて理由を返す */
                if (bucketResult.indexOf(okMarker) !== 0) { return bucketResult; }
                if (bucketResult !== okMarker) { lastResult = bucketResult; }
            }
            /* アートボードごとに選び直したので、最後に元の選択へ戻す */
            btSelectItems(doc, selectedItems);
            return lastResult;
        }

        function btGroupItemsByArtboard(doc, targetItems) {
            var buckets, artboardIndexes, itemBounds, artboardIndex, bucketPosition, i, j;
            buckets = [];
            artboardIndexes = [];
            for (i = 0; i < targetItems.length; i++) {
                itemBounds = btUnionBounds([targetItems[i]], true, false);
                artboardIndex = btFindOverlappingArtboardIndex(doc, itemBounds, doc.artboards.getActiveArtboardIndex());
                if (artboardIndex < 0) { artboardIndex = btFindNearestArtboardIndex(doc, itemBounds); }
                bucketPosition = -1;
                for (j = 0; j < artboardIndexes.length; j++) {
                    if (artboardIndexes[j] === artboardIndex) { bucketPosition = j; break; }
                }
                if (bucketPosition < 0) {
                    artboardIndexes.push(artboardIndex);
                    buckets.push([]);
                    bucketPosition = buckets.length - 1;
                }
                buckets[bucketPosition].push(targetItems[i]);
            }
            return buckets;
        }

        function btCopyOptions(workerOptions) {
            var optionsCopy, optionKey;
            optionsCopy = {};
            for (optionKey in workerOptions) {
                if (workerOptions.hasOwnProperty(optionKey)) { optionsCopy[optionKey] = workerOptions[optionKey]; }
            }
            return optionsCopy;
        }

        function btAlignItems(doc, selectedItems, workerOptions) {
            var needsGroup, didGroup, previousPreferences, justification, opticalKernings, failure;
            /* 寄せ先はこの組だけの判定で決まるので、呼び出し元の workerOptions は書き換えない */
            workerOptions = btCopyOptions(workerOptions);
            needsGroup = selectedItems.length > 1;
            if (needsGroup && btSpansMultipleLayers(selectedItems)) { return "MULTILAYER"; }
            btActivateArtboardForSelection(doc, selectedItems);
            /* すでにマージンの位置に寄っているなら、次はその外のアートボードのエッジへ寄せる（方向ボタンと同じ二段階）*/
            workerOptions.marginPt = btResolveStepMargin(doc, selectedItems, workerOptions);
            previousPreferences = btReadPreferences();
            justification = { previous: null, changed: null };
            opticalKernings = [];
            /* 実際にグループ化できたかを控える。needsGroup で解除すると、グループ化の前で例外が出たときに
               選択していた既存のグループを解除してしまう
               Track whether the group actually happened; keying the ungroup off needsGroup would
               dissolve the user's own groups when something throws before the group runs */
            didGroup = false;
            /* 整列できなかった理由。finally で後始末をしてから返す / Why the align failed; returned after the teardown */
            failure = null;
            try {
                btWritePreferences(workerOptions.previewBounds === true, workerOptions.glyphBounds === true, workerOptions.glyphBounds === true);
                /* 書いた環境設定を整列コマンドに拾わせるため、いったん反映させる
                   ここを飛ばすと「字形の境界に整列」が効かないまま整列されることがある */
                app.redraw();
                justification = btApplyJustification(selectedItems, workerOptions);
                /* 字面が変わるので、整列で測る前に入れる */
                if (workerOptions.opticalAdjust === true) {
                    opticalKernings = btApplyOpticalKerning(selectedItems, workerOptions.opticalStrength);
                } else {
                    /* OFF のときは、前に入れた調整を消す（各行の1文字目のカーニングを0に）*/
                    opticalKernings = btResetLineStartKerning(selectedItems);
                }
                if (needsGroup) {
                    app.executeMenuCommand("group");
                    didGroup = true;
                }
                if (!btRunAlignCommands(doc, workerOptions)) {
                    failure = "NOTARGET";
                } else {
                    btNudgeSelection(doc,
                        workerOptions.offsetX * workerOptions.marginPt + workerOptions.centerOffsetX,
                        workerOptions.offsetY * workerOptions.marginPt + workerOptions.centerOffsetY);
                }
            } catch (alignError) {
                failure = "ERR:" + alignError;
            } finally {
                btFinishAlign(selectedItems, didGroup, previousPreferences, justification.previous, failure);
                /* 整列できなかったときは、入れたカーニングも元に戻す */
                if (failure !== null) { btRestoreOpticalKerning(opticalKernings); }
            }
            if (failure !== null) { return failure; }
            try {
                /* 整列コマンドが「字形の境界に整列」を拾えていないことがあるため、
                   字形を実測して目標位置との差を打ち消す（拾えていれば差は0で何も動かない）*/
                btApplyGlyphCorrection(doc, workerOptions);
            } catch (correctionError) {
                return "ERR:" + correctionError;
            }
            if (justification.changed !== null) { return "OK:" + justification.changed; }
            return "OK";
        }

        function btApplyJustification(selectedItems, workerOptions) {
            var previousJustification;
            if (!workerOptions.changeJustification || !workerOptions.justification) { return { previous: null, changed: null }; }
            if (!btIsSingleLineTextFrame(selectedItems)) { return { previous: null, changed: null }; }
            previousJustification = btSetJustification(selectedItems[0], workerOptions.justification);
            if (previousJustification === null) { return { previous: null, changed: null }; }
            /* もとから同じ行揃えだったときは「変更した」と言わない */
            if (previousJustification === btJustificationByName(workerOptions.justification)) {
                return { previous: previousJustification, changed: null };
            }
            return { previous: previousJustification, changed: workerOptions.justification };
        }

        function btFinishAlign(selectedItems, didGroup, previousPreferences, previousJustification, failure) {
            if (didGroup) { app.executeMenuCommand("ungroup"); }
            /* 整列が環境設定を使い終えてから戻す（先に戻すと反映前の値で整列されることがある）*/
            app.redraw();
            btWritePreferences(previousPreferences.previewBounds, previousPreferences.pointText, previousPreferences.areaText);
            /* 整列できなかったときは、先に変えた行揃えも元に戻す（例外で抜けた場合も同じ）
               Roll the justification back whenever the align did not go through, exceptions included */
            if (failure === null) { return; }
            btRestoreJustification(selectedItems, previousJustification);
        }

        function btRestoreJustification(selectedItems, previousJustification) {
            if (previousJustification === null) { return; }
            try {
                selectedItems[0].textRange.paragraphAttributes.justification = previousJustification;
            } catch (restoreError) {}
        }

        function btMoveToEdgeOrGuide(workerOptions) {
            var doc, selectedItems;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            selectedItems = btResolveSelection(doc);
            if (selectedItems === null) { return "NOSEL"; }
            if (doc.artboards.length === 0) { return "NODOC"; }
            return btRunPerArtboard(doc, selectedItems, workerOptions, btMoveItems, "MOVED");
        }

        function btMoveItems(doc, selectedItems, workerOptions) {
            var needsGroup, didGroup, justification, failure, bounds, artboardRect, deltaX, deltaY, side, delta, i;
            needsGroup = selectedItems.length > 1;
            if (needsGroup && btSpansMultipleLayers(selectedItems)) { return "MULTILAYER"; }
            btActivateArtboardForSelection(doc, selectedItems);
            justification = { previous: null, changed: null };
            /* 実際にグループ化できたかを控える（例外で抜けたときに、選択していた既存のグループを解除しないため）*/
            didGroup = false;
            failure = null;
            try {
                /* 行揃えを先に変える。ポイント文字は行揃えを変えると文字が動くので、
                   変えたあとの境界で移動量を出さないと、そのぶんズレる */
                justification = btApplyJustification(selectedItems, workerOptions);
                bounds = btUnionBounds(selectedItems, workerOptions.previewBounds === true, workerOptions.glyphBounds === true);
                if (bounds === null) {
                    failure = "NOBOUNDS";
                } else {
                    artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
                    deltaX = 0;
                    deltaY = 0;
                    /* 斜めは2辺ぶんの移動量を、どちらも動かす前の境界から出して1回で動かす
                       先に片方だけ動かすと、もう一方のガイド探しが動いたあとの位置で走ってしまう */
                    for (i = 0; i < workerOptions.sides.length; i++) {
                        side = workerOptions.sides[i];
                        delta = btEdgeDelta(doc, artboardRect, bounds, side, workerOptions);
                        if (side === "left" || side === "right") { deltaX = delta; } else { deltaY = delta; }
                    }
                    if (needsGroup) {
                        app.executeMenuCommand("group");
                        didGroup = true;
                    }
                    btNudgeSelection(doc, deltaX, deltaY);
                }
            } catch (moveError) {
                failure = "ERR:" + moveError;
            } finally {
                if (didGroup) { app.executeMenuCommand("ungroup"); }
                /* 動かせなかったときは、先に変えた行揃えも元に戻す */
                if (failure !== null) { btRestoreJustification(selectedItems, justification.previous); }
            }
            if (failure !== null) { return failure; }
            if (justification.changed !== null) { return "MOVED:" + justification.changed; }
            return "MOVED";
        }

        function btEdgeDelta(doc, artboardRect, bounds, side, workerOptions) {
            var selectionEdge, artboardEdge, bleedEdge, targetEdge, delta;
            selectionEdge = btEdgeValue(bounds, side);
            artboardEdge = btEdgeValue(artboardRect, side);
            /* 進む向きにガイドがあればそこで止め、無ければアートボードのエッジまで動かす */
            targetEdge = btFindGuideSnapValue(doc, artboardRect, bounds, side, workerOptions.guideTolerance, workerOptions.minDeltaPt);
            if (targetEdge === null) {
                targetEdge = artboardEdge;
                if (workerOptions.bleedPt > 0) {
                    bleedEdge = btOutsetValue(artboardEdge, side, workerOptions.bleedPt);
                    /* 裁ち落としまで出ていればそこに留め、アートボードのエッジに乗っていれば次は裁ち落としへ */
                    if (Math.abs(bleedEdge - selectionEdge) < workerOptions.minDeltaPt) { return 0; }
                    if (Math.abs(artboardEdge - selectionEdge) < workerOptions.minDeltaPt) { targetEdge = bleedEdge; }
                }
            }
            delta = targetEdge - selectionEdge;
            if (Math.abs(delta) < workerOptions.minDeltaPt) { delta = 0; }
            return delta;
        }

        function btOutsetValue(edgeValue, side, bleedPt) {
            if (side === "left") { return edgeValue - bleedPt; }
            if (side === "right") { return edgeValue + bleedPt; }
            if (side === "top") { return edgeValue + bleedPt; }
            return edgeValue - bleedPt;
        }

        function btEdgeValue(bounds, side) {
            if (side === "left") { return bounds[0]; }
            if (side === "top") { return bounds[1]; }
            if (side === "right") { return bounds[2]; }
            return bounds[3];
        }

        function btFindGuideSnapValue(doc, artboardRect, bounds, side, tolerance, epsilon) {
            var selectionEdge, oppositeEdge, ahead, aheadDistance, behindNear, behindNearDistance, behindFar, behindFarDistance,
                guideItems, i, guideValues, j, guideValue, distance;
            selectionEdge = btEdgeValue(bounds, side);
            oppositeEdge = btEdgeValue(bounds, btOppositeSide(side));
            ahead = null;
            aheadDistance = null;
            behindNear = null;
            behindNearDistance = null;
            behindFar = null;
            behindFarDistance = null;
            guideItems = doc.pathItems;
            for (i = 0; i < guideItems.length; i++) {
                if (guideItems[i].guides !== true) { continue; }
                guideValues = btGuideEdgeValues(guideItems[i].geometricBounds, bounds, side, tolerance, epsilon);
                for (j = 0; j < guideValues.length; j++) {
                    guideValue = guideValues[j];
                    if (!btIsInsideArtboard(guideValue, artboardRect, side)) { continue; }
                    distance = Math.abs(guideValue - selectionEdge);
                    /* すでにその辺が乗っているガイドは行き先にしない（手で吸着させたときの微小なずれを吸収する） */
                    if (distance <= epsilon) { continue; }
                    if (btIsAhead(guideValue, selectionEdge, side, epsilon)) {
                        if (ahead === null || distance < aheadDistance) { ahead = guideValue; aheadDistance = distance; }
                    } else {
                        if (behindNear === null || distance < behindNearDistance) { behindNear = guideValue; behindNearDistance = distance; }
                        if (behindFar === null || distance > behindFarDistance) { behindFar = guideValue; behindFarDistance = distance; }
                    }
                }
            }
            if (ahead !== null) { return ahead; }
            /* 進む向きにガイドが無く、なおかつオブジェクトがガイド全体をまたいでいる（ガイドの間隔より大きい）ときだけ、
               戻る向きの最も近いガイドに合わせる。またいでいなければ null を返し、呼び出し側がアートボードのエッジを使う */
            if (behindNear !== null && btIsAhead(behindFar, oppositeEdge, side, epsilon)) { return behindNear; }
            return null;
        }

        function btOppositeSide(side) {
            if (side === "left") { return "right"; }
            if (side === "right") { return "left"; }
            if (side === "top") { return "bottom"; }
            return "top";
        }

        function btGuideEdgeValues(guideBounds, bounds, side, tolerance, epsilon) {
            var isVertical, isHorizontal;
            isVertical = Math.abs(guideBounds[2] - guideBounds[0]) <= tolerance;
            isHorizontal = Math.abs(guideBounds[1] - guideBounds[3]) <= tolerance;
            /* 線のガイドはその1本、長方形のガイド（マージンガイドなど）は向かい合う2辺を候補にする
               動かす向きと直交する方向で選択範囲と重ならないガイド＝そのまま動かしてもぶつからないガイドは対象にしない
               （ルーラーのガイドはアートボードの外まで伸びているので、いつでも重なる） */
            if (side === "left" || side === "right") {
                if (isHorizontal) { return []; }
                if (guideBounds[1] - bounds[3] <= epsilon || bounds[1] - guideBounds[3] <= epsilon) { return []; }
                return isVertical ? [guideBounds[0]] : [guideBounds[0], guideBounds[2]];
            }
            if (isVertical) { return []; }
            if (guideBounds[2] - bounds[0] <= epsilon || bounds[2] - guideBounds[0] <= epsilon) { return []; }
            return isHorizontal ? [guideBounds[1]] : [guideBounds[1], guideBounds[3]];
        }

        function btIsInsideArtboard(edgeValue, artboardRect, side) {
            if (side === "left" || side === "right") { return edgeValue >= artboardRect[0] && edgeValue <= artboardRect[2]; }
            return edgeValue <= artboardRect[1] && edgeValue >= artboardRect[3];
        }

        function btIsAhead(edgeValue, referenceEdge, side, epsilon) {
            if (side === "left") { return edgeValue < referenceEdge - epsilon; }
            if (side === "right") { return edgeValue > referenceEdge + epsilon; }
            if (side === "top") { return edgeValue > referenceEdge + epsilon; }
            return edgeValue < referenceEdge - epsilon;
        }

        function btUpdateMarginGuide(workerOptions) {
            var doc;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            btDrawMarginGuide(doc, workerOptions);
            return "OK";
        }

        function btDrawMarginGuide(doc, workerOptions) {
            var guideLayer, previousLayerState, artboardIndexes, artboardRect, marginArea, frameAreas, divideLines, i;
            btRemoveGuidesByName(doc, workerOptions.guideName, workerOptions.guideLayerName);
            btRemoveGuidesByName(doc, workerOptions.divideName, workerOptions.guideLayerName);
            if (doc.artboards.length === 0) { return; }
            /* ［アートボードごと］がONなら、すべてのアートボードに描く。OFFなら作業中のアートボードだけ */
            artboardIndexes = [];
            if (workerOptions.perArtboard === true) {
                for (i = 0; i < doc.artboards.length; i++) { artboardIndexes.push(i); }
            } else {
                artboardIndexes.push(doc.artboards.getActiveArtboardIndex());
            }
            frameAreas = [];
            divideLines = [];
            for (i = 0; i < artboardIndexes.length; i++) {
                artboardRect = doc.artboards[artboardIndexes[i]].artboardRect;
                /* マージンが大きすぎて内側が残らないときは null。マージンに依らないエッジのガイドは引ける */
                marginArea = btMarginArea(artboardRect, workerOptions.guideMargins);
                if (marginArea !== null && workerOptions.showGuide === true && btHasMargin(workerOptions.guideMargins)) {
                    frameAreas.push(marginArea);
                }
                divideLines = divideLines.concat(btDivideLines(marginArea, artboardRect, workerOptions.divisions));
            }
            /* 描くものが無いときは、ガイド用レイヤーを作らずに戻る */
            if (frameAreas.length === 0 && divideLines.length === 0) { return; }
            guideLayer = btGetGuideLayer(doc, workerOptions.guideLayerName);
            previousLayerState = btUnlockLayer(guideLayer);
            try {
                for (i = 0; i < frameAreas.length; i++) { btAddGuideRectangle(guideLayer, frameAreas[i], workerOptions.guideName); }
                btAddGuideLines(guideLayer, divideLines, workerOptions.divideName);
            } finally {
                btRestoreLayer(guideLayer, previousLayerState);
            }
        }

        function btMarginArea(artboardRect, margins) {
            var marginArea = {
                left:   artboardRect[0] + margins.left,
                top:    artboardRect[1] - margins.top,
                right:  artboardRect[2] - margins.right,
                bottom: artboardRect[3] + margins.bottom
            };
            if (marginArea.right <= marginArea.left || marginArea.top <= marginArea.bottom) { return null; }
            return marginArea;
        }

        function btAddGuideRectangle(guideLayer, marginArea, guideName) {
            btMakeGuide(guideLayer.pathItems.rectangle(marginArea.top, marginArea.left, marginArea.right - marginArea.left, marginArea.top - marginArea.bottom), guideName);
        }

        function btAddGuideLines(guideLayer, guideLines, guideName) {
            var guideLine, i;
            for (i = 0; i < guideLines.length; i++) {
                guideLine = guideLayer.pathItems.add();
                guideLine.setEntirePath(guideLines[i]);
                btMakeGuide(guideLine, guideName);
            }
        }

        function btMakeGuide(pathItem, guideName) {
            var isGuide = false;
            try {
                /* 名前を先に付ける。ガイドにできずに消し損ねても、次の描き直しで名前を頼りに消せる */
                pathItem.name = guideName;
                pathItem.stroked = false;
                pathItem.filled = false;
                pathItem.guides = true;
                isGuide = (pathItem.guides === true);
            } catch (guideError) {}
            if (isGuide) { return; }
            /* ガイドにできなかったときは、描画時の塗り・線を持ったオブジェクトとして残さずに消す */
            try { pathItem.remove(); } catch (removeError) {}
        }

        function btDivideLines(marginArea, artboardRect, divisions) {
            var guideLines, extension, span;
            guideLines = [];
            if (!divisions) { return guideLines; }
            /* 分割の線はアートボードの幅・高さいっぱいに引き、伸張のぶんだけその外へ伸ばす */
            extension = divisions.extension > 0 ? divisions.extension : 0;
            span = {
                left:   artboardRect[0] - extension,
                top:    artboardRect[1] + extension,
                right:  artboardRect[2] + extension,
                bottom: artboardRect[3] - extension
            };
            btPushRowLines(guideLines, marginArea, divisions, span);
            btPushColumnLines(guideLines, marginArea, divisions, span);
            btPushArtboardEdgeLines(guideLines, artboardRect, divisions, span);
            return guideLines;
        }

        function btPushArtboardEdgeLines(guideLines, artboardRect, divisions, span) {
            if (divisions.artboardEdge !== true) { return; }
            guideLines.push([[span.left, artboardRect[1]], [span.right, artboardRect[1]]]);
            guideLines.push([[span.left, artboardRect[3]], [span.right, artboardRect[3]]]);
            guideLines.push([[artboardRect[0], span.top], [artboardRect[0], span.bottom]]);
            guideLines.push([[artboardRect[2], span.top], [artboardRect[2], span.bottom]]);
        }

        function btPushRowLines(guideLines, marginArea, divisions, span) {
            var usableHeight, cellHeight, lineY, i;
            if (marginArea === null || !(divisions.rows > 1)) { return; }
            usableHeight = (marginArea.top - marginArea.bottom) - (divisions.rows - 1) * divisions.rowGutter;
            /* 行間が大きすぎて段の高さが残らないときは引かない */
            if (usableHeight <= 0) { return; }
            cellHeight = usableHeight / divisions.rows;
            lineY = marginArea.top;
            /* 最後の段は area.bottom にちょうど着地するので、内側の区切りだけを引く */
            for (i = 0; i < divisions.rows - 1; i++) {
                lineY -= cellHeight;
                guideLines.push([[span.left, lineY], [span.right, lineY]]);
                /* 行間が0のときは同じ位置に重なるので引かない */
                if (divisions.rowGutter > 0) {
                    lineY -= divisions.rowGutter;
                    guideLines.push([[span.left, lineY], [span.right, lineY]]);
                }
            }
        }

        function btPushColumnLines(guideLines, marginArea, divisions, span) {
            var usableWidth, cellWidth, lineX, i;
            if (marginArea === null || !(divisions.columns > 1)) { return; }
            usableWidth = (marginArea.right - marginArea.left) - (divisions.columns - 1) * divisions.columnGutter;
            if (usableWidth <= 0) { return; }
            cellWidth = usableWidth / divisions.columns;
            lineX = marginArea.left;
            for (i = 0; i < divisions.columns - 1; i++) {
                lineX += cellWidth;
                guideLines.push([[lineX, span.top], [lineX, span.bottom]]);
                if (divisions.columnGutter > 0) {
                    lineX += divisions.columnGutter;
                    guideLines.push([[lineX, span.top], [lineX, span.bottom]]);
                }
            }
        }

        function btHasMargin(margins) {
            return margins.top > 0 || margins.bottom > 0 || margins.left > 0 || margins.right > 0;
        }

        function btGetGuideLayer(doc, layerName) {
            var guideLayer;
            try {
                guideLayer = doc.layers.getByName(layerName);
            } catch (missingLayerError) {
                guideLayer = doc.layers.add();
                guideLayer.name = layerName;
            }
            return guideLayer;
        }

        function btUnlockLayer(targetLayer) {
            var previousLayerState;
            /* ロック・非表示のままでは書き換えられないので一時的に外し、戻せるよう控える */
            previousLayerState = { locked: targetLayer.locked, visible: targetLayer.visible };
            targetLayer.locked = false;
            targetLayer.visible = true;
            return previousLayerState;
        }

        function btRestoreLayer(targetLayer, previousLayerState) {
            try {
                targetLayer.visible = previousLayerState.visible;
                targetLayer.locked = previousLayerState.locked;
            } catch (restoreLayerError) {}
        }

        function btRemoveGuidesByName(doc, guideName, layerName) {
            var guideLayer, previousLayerState, layerPaths, i, pathItem;
            try {
                guideLayer = doc.layers.getByName(layerName);
            } catch (missingLayerError) {
                return;
            }
            previousLayerState = btUnlockLayer(guideLayer);
            try {
                layerPaths = guideLayer.pathItems;
                for (i = layerPaths.length - 1; i >= 0; i--) {
                    pathItem = layerPaths[i];
                    if (pathItem.name !== guideName) { continue; }
                    try {
                        pathItem.locked = false;
                        pathItem.hidden = false;
                        pathItem.remove();
                    } catch (removeError) {}
                }
            } finally {
                btRestoreLayer(guideLayer, previousLayerState);
            }
        }

        function btAlignMovedSelection(doc, workerOptions) {
            var boundsBefore, boundsAfter, i;
            boundsBefore = btUnionBounds(doc.selection, true, false);
            for (i = 0; i < workerOptions.alignCommands.length; i++) {
                app.executeMenuCommand(workerOptions.alignCommands[i]);
            }
            boundsAfter = btUnionBounds(doc.selection, true, false);
            return !btSameBounds(boundsBefore, boundsAfter);
        }

        function btRunAlignCommands(doc, workerOptions) {
            if (btAlignMovedSelection(doc, workerOptions)) { return true; }
            return btProbeAlignTarget(doc, workerOptions);
        }

        function btProbeAlignTarget(doc, workerOptions) {
            var moved;
            /* 整列で動かなかったときだけ、いったんずらして整列し直し、
               戻ってくるかどうかで整列先がアートボードかを見る */
            btNudgeSelection(doc, workerOptions.probeX, workerOptions.probeY);
            try {
                moved = btAlignMovedSelection(doc, workerOptions);
            } catch (probeError) {
                btNudgeSelection(doc, -workerOptions.probeX, -workerOptions.probeY);
                throw probeError;
            }
            if (!moved) { btNudgeSelection(doc, -workerOptions.probeX, -workerOptions.probeY); }
            return moved;
        }

        function btSameBounds(boundsA, boundsB) {
            var i;
            for (i = 0; i < 4; i++) {
                if (Math.abs(boundsA[i] - boundsB[i]) > 0.0001) { return false; }
            }
            return true;
        }

        function btApplyGlyphCorrection(doc, workerOptions) {
            var selectedItems, bounds, artboardRect, deltaX, deltaY;
            if (workerOptions.glyphBounds !== true) { return; }
            selectedItems = doc.selection;
            if (!(selectedItems instanceof Array) || selectedItems.length === 0) { return; }
            if (!btHasGlyphBoundsTarget(selectedItems)) { return; }
            if (doc.artboards.length === 0) { return; }
            bounds = btUnionBounds(selectedItems, workerOptions.previewBounds === true, true);
            if (bounds === null) { return; }
            artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            /* artboardRect は [左, 上, 右, 下] で、X は右へ・Y は下へ向かって内側になるため符号が逆になる */
            deltaX = btAlignDelta(bounds, artboardRect, workerOptions.modeX, workerOptions.marginPt, workerOptions.centerOffsetX, 0, 2, 1);
            deltaY = btAlignDelta(bounds, artboardRect, workerOptions.modeY, workerOptions.marginPt, workerOptions.centerOffsetY, 1, 3, -1);
            if (Math.abs(deltaX) < 0.001) { deltaX = 0; }
            if (Math.abs(deltaY) < 0.001) { deltaY = 0; }
            btNudgeSelection(doc, deltaX, deltaY);
        }

        function btAlignDelta(bounds, artboardRect, alignMode, marginPt, centerOffset, startIndex, endIndex, inwardSign) {
            if (alignMode === "start") { return (artboardRect[startIndex] + inwardSign * marginPt) - bounds[startIndex]; }
            if (alignMode === "end") { return (artboardRect[endIndex] - inwardSign * marginPt) - bounds[endIndex]; }
            /* 中央はアートボードの中心。Option＋クリックのときは centerOffset のぶんマージンの内側の中心へずらす */
            if (alignMode === "center") {
                return (artboardRect[startIndex] + artboardRect[endIndex]) / 2 + centerOffset - (bounds[startIndex] + bounds[endIndex]) / 2;
            }
            return 0;
        }

        function btAlignStops(workerOptions) {
            var alignStops;
            /* 内側から順に、マージン → アートボードのエッジ → 裁ち落とし。
               値はアートボードの辺からの内向きの量なので、裁ち落としは負になる */
            alignStops = [];
            if (workerOptions.marginPt > 0) { alignStops.push(workerOptions.marginPt); }
            alignStops.push(0);
            if (workerOptions.bleedPt > 0) { alignStops.push(-workerOptions.bleedPt); }
            return alignStops;
        }

        function btResolveStepMargin(doc, selectedItems, workerOptions) {
            var alignStops, bounds, artboardRect, i;
            alignStops = btAlignStops(workerOptions);
            if (alignStops.length < 2) { return alignStops[0]; }
            if (doc.artboards.length === 0) { return alignStops[0]; }
            bounds = btUnionBounds(selectedItems, workerOptions.previewBounds === true, workerOptions.glyphBounds === true);
            if (bounds === null) { return alignStops[0]; }
            artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
            /* いま乗っている停止位置の次へ進める。いちばん外まで来ていればそこに留める */
            for (i = 0; i < alignStops.length; i++) {
                if (btIsAtAlignTarget(bounds, artboardRect, workerOptions, alignStops[i])) {
                    return (i + 1 < alignStops.length) ? alignStops[i + 1] : alignStops[i];
                }
            }
            return alignStops[0];
        }

        function btIsAtAlignTarget(bounds, artboardRect, workerOptions, marginPt) {
            if (Math.abs(btAlignDelta(bounds, artboardRect, workerOptions.modeX, marginPt, workerOptions.centerOffsetX, 0, 2, 1)) > workerOptions.minDeltaPt) { return false; }
            if (Math.abs(btAlignDelta(bounds, artboardRect, workerOptions.modeY, marginPt, workerOptions.centerOffsetY, 1, 3, -1)) > workerOptions.minDeltaPt) { return false; }
            return true;
        }

        function btHasGlyphBoundsTarget(targetItems) {
            var i, typeName;
            for (i = 0; i < targetItems.length; i++) {
                typeName = targetItems[i].typename;
                if (typeName === "TextFrame" || typeName === "SymbolItem") { return true; }
            }
            return false;
        }

        function btUnionBounds(targetItems, usePreviewBounds, useGlyphBounds) {
            var bounds, itemBounds, i;
            bounds = null;
            for (i = 0; i < targetItems.length; i++) {
                /* 削除済みの参照が混ざることがあるので、測れないものは飛ばす */
                try {
                    itemBounds = btGetMeasureBounds(targetItems[i], usePreviewBounds, useGlyphBounds);
                } catch (boundsError) {
                    continue;
                }
                if (!itemBounds) { continue; }
                if (bounds === null) {
                    bounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
                } else {
                    if (itemBounds[0] < bounds[0]) { bounds[0] = itemBounds[0]; }
                    if (itemBounds[1] > bounds[1]) { bounds[1] = itemBounds[1]; }
                    if (itemBounds[2] > bounds[2]) { bounds[2] = itemBounds[2]; }
                    if (itemBounds[3] < bounds[3]) { bounds[3] = itemBounds[3]; }
                }
            }
            return bounds;
        }

        function btGetMeasureBounds(targetItem, usePreviewBounds, useGlyphBounds) {
            var bounds;
            bounds = null;
            if (useGlyphBounds === true) {
                if (targetItem.typename === "TextFrame") { bounds = btGetOutlineBounds(targetItem, usePreviewBounds); }
                else if (targetItem.typename === "SymbolItem") { bounds = btGetSymbolGlyphBounds(targetItem, usePreviewBounds); }
                if (bounds !== null) { return bounds; }
            }
            return btGetPlainBounds(targetItem, usePreviewBounds);
        }

        function btGetPlainBounds(targetItem, usePreviewBounds) {
            var bounds;
            /* クリップグループの境界はマスクを無視した中身全体を返すため、マスクのパスで測り直す */
            if (targetItem.typename === "GroupItem" && targetItem.clipped === true) {
                bounds = btGetClipPathBounds(targetItem, usePreviewBounds);
                if (bounds !== null) { return bounds; }
            }
            return usePreviewBounds ? targetItem.visibleBounds : targetItem.geometricBounds;
        }

        function btGetClipPathBounds(groupItem, usePreviewBounds) {
            var clipPath;
            clipPath = btFindClipPath(groupItem);
            if (clipPath === null) { return null; }
            return usePreviewBounds ? clipPath.visibleBounds : clipPath.geometricBounds;
        }

        function btFindClipPath(groupItem) {
            var childItem, typeName, i;
            for (i = 0; i < groupItem.pageItems.length; i++) {
                childItem = groupItem.pageItems[i];
                typeName = childItem.typename;
                if (typeName === "PathItem" && childItem.clipping === true) { return childItem; }
                /* 複合パスのマスクは、束ねているパスの clipping にだけ立つ */
                if (typeName === "CompoundPathItem" && childItem.pathItems.length > 0 && childItem.pathItems[0].clipping === true) { return childItem; }
            }
            return null;
        }

        function btGetOutlineBounds(textFrame, usePreviewBounds) {
            var duplicated, outlined, bounds;
            duplicated = null;
            outlined = null;
            bounds = null;
            try {
                duplicated = textFrame.duplicate();
                outlined = duplicated.createOutline();
                bounds = usePreviewBounds ? outlined.visibleBounds : outlined.geometricBounds;
            } catch (outlineError) {
                bounds = null;
            } finally {
                btSafeRemove(outlined);
                btSafeRemove(duplicated);
            }
            return bounds;
        }

        function btGetSymbolGlyphBounds(symbolItem, usePreviewBounds) {
            var doc, symbolLayer, selectionBefore, itemsBefore, layersBefore, brokenItems, hasText, bounds, i;
            symbolLayer = symbolItem.layer;
            if (!symbolLayer) { return null; }
            doc = app.activeDocument;
            selectionBefore = doc.selection;
            itemsBefore = btCollectionToArray(symbolLayer.pageItems);
            layersBefore = btCollectionToArray(symbolLayer.layers);
            bounds = null;
            try {
                /* 複製をレイヤー直下に置いてからリンクを解除する
                   グループの中で解除すると生成物の行き先が読めないため、いったん外へ出す */
                symbolItem.duplicate(symbolLayer, ElementPlacement.PLACEATBEGINNING).breakLink();
                /* breakLink は複製そのものを別のオブジェクトに置き換えるので、参照ではなく
                   レイヤー直下の差分で生成物を拾う（静的シンボルはサブレイヤーを作る）*/
                brokenItems = btCollectBrokenItems(symbolLayer, itemsBefore, layersBefore);
                hasText = false;
                for (i = 0; i < brokenItems.length; i++) {
                    if (btOutlineTextFramesIn(brokenItems[i])) { hasText = true; }
                }
                /* 文字が無ければシンボルの通常の境界と変わらないので、呼び出し元の既定に任せる */
                if (hasText) {
                    /* アウトライン化でテキストは別のオブジェクトに置き換わるため、拾い直してから測る */
                    bounds = btUnionBounds(btCollectBrokenItems(symbolLayer, itemsBefore, layersBefore), usePreviewBounds, false);
                }
            } catch (symbolError) {
                bounds = null;
            } finally {
                btRemoveBrokenItems(symbolLayer, itemsBefore, layersBefore);
                /* breakLink は生成物を選択状態にするため、元の選択に戻す */
                if (selectionBefore instanceof Array) {
                    try { doc.selection = selectionBefore; } catch (selectionError) {}
                }
            }
            return bounds;
        }

        function btCollectionToArray(collection) {
            var collectionItems, i;
            collectionItems = [];
            for (i = 0; i < collection.length; i++) { collectionItems.push(collection[i]); }
            return collectionItems;
        }

        function btCollectNewEntries(currentEntries, existingEntries) {
            var newEntries, isExisting, i, j;
            newEntries = [];
            for (i = 0; i < currentEntries.length; i++) {
                isExisting = false;
                for (j = 0; j < existingEntries.length; j++) {
                    if (existingEntries[j] === currentEntries[i]) { isExisting = true; break; }
                }
                if (!isExisting) { newEntries.push(currentEntries[i]); }
            }
            return newEntries;
        }

        function btCollectBrokenItems(symbolLayer, itemsBefore, layersBefore) {
            var brokenItems, newLayers, i, j;
            brokenItems = btCollectNewEntries(symbolLayer.pageItems, itemsBefore);
            newLayers = btCollectNewEntries(symbolLayer.layers, layersBefore);
            for (i = 0; i < newLayers.length; i++) {
                for (j = 0; j < newLayers[i].pageItems.length; j++) { brokenItems.push(newLayers[i].pageItems[j]); }
            }
            return brokenItems;
        }

        function btRemoveBrokenItems(symbolLayer, itemsBefore, layersBefore) {
            var newItems, newLayers, i;
            try {
                newItems = btCollectNewEntries(symbolLayer.pageItems, itemsBefore);
                newLayers = btCollectNewEntries(symbolLayer.layers, layersBefore);
            } catch (collectError) {
                return;
            }
            for (i = 0; i < newItems.length; i++) { btSafeRemove(newItems[i]); }
            for (i = 0; i < newLayers.length; i++) { btSafeRemove(newLayers[i]); }
        }

        function btOutlineTextFramesIn(targetItem) {
            var textFrames, outlined, i;
            textFrames = [];
            btCollectTextFrames(targetItem, textFrames);
            /* 先に集めてからアウトライン化する。createOutline は元のテキストを置き換えるので、
               pageItems をたどりながらだと取りこぼす */
            outlined = false;
            for (i = 0; i < textFrames.length; i++) {
                try {
                    textFrames[i].createOutline();
                    outlined = true;
                } catch (outlineError) {}
            }
            return outlined;
        }

        function btCollectTextFrames(targetItem, textFrames) {
            var i;
            if (!targetItem) { return; }
            if (targetItem.typename === "TextFrame") { textFrames.push(targetItem); return; }
            if (targetItem.typename !== "GroupItem") { return; }
            for (i = 0; i < targetItem.pageItems.length; i++) { btCollectTextFrames(targetItem.pageItems[i], textFrames); }
        }

        function btSafeRemove(targetItem) {
            try {
                if (targetItem) { targetItem.remove(); }
            } catch (removeError) {}
        }

        function btGetSelectionKind() {
            var doc, selectedItems, i;
            if (app.documents.length === 0) { return "NODOC"; }
            doc = app.activeDocument;
            selectedItems = doc.selection;
            if (selectedItems && !(selectedItems instanceof Array)) { return "TEXT"; }
            if (!(selectedItems instanceof Array) || selectedItems.length === 0) { return "NONE"; }
            for (i = 0; i < selectedItems.length; i++) {
                if (selectedItems[i].typename !== "TextFrame") { return "OTHER"; }
            }
            return "TEXT";
        }

        function btGetPaletteState() {
            var doc, selectionCount, artboardCount;
            selectionCount = 0;
            artboardCount = 0;
            if (app.documents.length > 0) {
                doc = app.activeDocument;
                if (doc.selection instanceof Array) { selectionCount = doc.selection.length; }
                else if (doc.selection) { selectionCount = 1; }
                artboardCount = doc.artboards.length;
            }
            return btGetSelectionKind() + "|" + app.preferences.getIntegerPreference("rulerType") +
                "|" + selectionCount + "|" + artboardCount;
        }

        function btNudgeSelection(doc, deltaX, deltaY) {
            var selectedItems, i;
            if (deltaX === 0 && deltaY === 0) { return; }
            selectedItems = doc.selection;
            if (!(selectedItems instanceof Array)) { return; }
            for (i = 0; i < selectedItems.length; i++) {
                selectedItems[i].translate(deltaX, deltaY);
            }
        }

        function btReadPreferences() {
            return {
                previewBounds: app.preferences.getBooleanPreference("includeStrokeInBounds") === true,
                pointText: app.preferences.getBooleanPreference("EnableActualPointTextSpaceAlign") === true,
                areaText: app.preferences.getBooleanPreference("EnableActualAreaTextSpaceAlign") === true
            };
        }

        function btGetPreferenceFlags() {
            var preferences;
            preferences = btReadPreferences();
            /* 字形の境界はポイント文字・エリア内文字のどちらかがONならON扱いにする */
            return (preferences.previewBounds ? "1" : "0") + "|" + ((preferences.pointText || preferences.areaText) ? "1" : "0");
        }

        function btWritePreferences(previewBounds, pointText, areaText) {
            app.preferences.setBooleanPreference("includeStrokeInBounds", previewBounds === true);
            app.preferences.setBooleanPreference("EnableActualPointTextSpaceAlign", pointText === true);
            app.preferences.setBooleanPreference("EnableActualAreaTextSpaceAlign", areaText === true);
        }

        function btPromoteTextRange(doc, textRange) {
            var storyFrames, targetFrames, i;
            storyFrames = textRange.story.textFrames;
            targetFrames = [];
            for (i = 0; i < storyFrames.length; i++) { targetFrames.push(storyFrames[i]); }
            app.executeMenuCommand("deselectall");
            for (i = 0; i < targetFrames.length; i++) { targetFrames[i].selected = true; }
            return doc.selection;
        }

        function btGetLayerKey(targetLayer) {
            var keyParts, layerNode;
            keyParts = [];
            layerNode = targetLayer;
            while (layerNode && layerNode.typename === "Layer") {
                keyParts.push(layerNode.zOrderPosition + ":" + layerNode.name);
                layerNode = layerNode.parent;
            }
            return keyParts.join("/");
        }

        function btSpansMultipleLayers(selectedItems) {
            var firstKey, i;
            firstKey = btGetLayerKey(selectedItems[0].layer);
            for (i = 1; i < selectedItems.length; i++) {
                if (btGetLayerKey(selectedItems[i].layer) !== firstKey) { return true; }
            }
            return false;
        }

        function btGetOverlapArea(boundsA, boundsB) {
            var overlapWidth, overlapHeight;
            overlapWidth = Math.min(boundsA[2], boundsB[2]) - Math.max(boundsA[0], boundsB[0]);
            overlapHeight = Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]);
            if (overlapWidth <= 0 || overlapHeight <= 0) { return 0; }
            return overlapWidth * overlapHeight;
        }

        function btFindOverlappingArtboardIndex(doc, selectionBounds, currentIndex) {
            var searchOrder, i, j, bestIndex, bestArea, overlapArea;
            searchOrder = [currentIndex];
            for (i = 0; i < doc.artboards.length; i++) {
                if (i !== currentIndex) { searchOrder.push(i); }
            }
            bestIndex = -1;
            bestArea = 0;
            for (j = 0; j < searchOrder.length; j++) {
                overlapArea = btGetOverlapArea(selectionBounds, doc.artboards[searchOrder[j]].artboardRect);
                if (overlapArea > bestArea) { bestArea = overlapArea; bestIndex = searchOrder[j]; }
            }
            return bestIndex;
        }

        function btFindNearestArtboardIndex(doc, selectionBounds) {
            var centerX, centerY, nearestIndex, nearestDistance, i, artboardRect, offsetX, offsetY, distance;
            centerX = (selectionBounds[0] + selectionBounds[2]) / 2;
            centerY = (selectionBounds[1] + selectionBounds[3]) / 2;
            nearestIndex = 0;
            nearestDistance = null;
            for (i = 0; i < doc.artboards.length; i++) {
                artboardRect = doc.artboards[i].artboardRect;
                offsetX = centerX - (artboardRect[0] + artboardRect[2]) / 2;
                offsetY = centerY - (artboardRect[1] + artboardRect[3]) / 2;
                distance = offsetX * offsetX + offsetY * offsetY;
                if (nearestDistance === null || distance < nearestDistance) {
                    nearestDistance = distance;
                    nearestIndex = i;
                }
            }
            return nearestIndex;
        }

        function btActivateArtboardForSelection(doc, selectedItems) {
            var selectionBounds, currentIndex, targetIndex;
            if (doc.artboards.length < 2) { return; }
            selectionBounds = btUnionBounds(selectedItems, true, false);
            currentIndex = doc.artboards.getActiveArtboardIndex();
            targetIndex = btFindOverlappingArtboardIndex(doc, selectionBounds, currentIndex);
            if (targetIndex < 0) { targetIndex = btFindNearestArtboardIndex(doc, selectionBounds); }
            if (targetIndex !== currentIndex) { doc.artboards.setActiveArtboardIndex(targetIndex); }
        }

        function btIsSingleLineTextFrame(selectedItems) {
            if (selectedItems.length !== 1 || selectedItems[0].typename !== "TextFrame") { return false; }
            return selectedItems[0].lines.length === 1;
        }

        function btJustificationByName(justificationName) {
            if (justificationName === "LEFT") { return Justification.LEFT; }
            if (justificationName === "CENTER") { return Justification.CENTER; }
            if (justificationName === "RIGHT") { return Justification.RIGHT; }
            return null;
        }

        function btSetJustification(textFrame, justificationName) {
            var justification, previousJustification;
            justification = btJustificationByName(justificationName);
            if (justification === null) { return null; }
            previousJustification = textFrame.textRange.paragraphAttributes.justification;
            textFrame.textRange.paragraphAttributes.justification = justification;
            return previousJustification;
        }

        function btApplyOpticalKerning(selectedItems, strength) {
            var textFrames, applied, i;
            /* 見た目の調整：中央揃えのポイント文字の行ごとに字面の中心を測り、行頭のカーニングで寄せる。
               行頭に k（1/1000 em）を入れると行幅が k 伸び、字面は k/2 右へ動くので k = 2d / サイズ x 1000。
               _templates/OpticalCenterKerning.jsx のワーカー版。改行は fromCharCode で書く（body の逆スラッシュは二重になる）*/
            textFrames = [];
            for (i = 0; i < selectedItems.length; i++) { btCollectTextFrames(selectedItems[i], textFrames); }
            applied = [];
            for (i = 0; i < textFrames.length; i++) {
                try {
                    btApplyOpticalKerningToFrame(textFrames[i], strength, applied);
                } catch (opticalError) {}
            }
            return applied;
        }

        function btApplyOpticalKerningToFrame(textFrame, strength, applied) {
            var characters, centerX, lineTexts, kernings, lineStart, lineEnd, ink, fontSize, kerning, previousKerning, i;
            if (textFrame.kind !== TextType.POINTTEXT || textFrame.locked || textFrame.hidden) { return; }
            characters = textFrame.characters;
            centerX = textFrame.anchor[0];
            lineTexts = textFrame.contents.split(String.fromCharCode(13));
            /* 先にすべての行を測ってから書き込む */
            kernings = [];
            lineStart = 0;
            for (i = 0; i < lineTexts.length; i++) {
                lineEnd = lineStart + lineTexts[i].length;
                if (lineEnd > lineStart && characters[lineStart].paragraphAttributes.justification === Justification.CENTER) {
                    ink = btMeasureLineInk(textFrame, lineStart, lineEnd);
                    if (ink !== null) {
                        fontSize = characters[lineStart].characterAttributes.size;
                        kerning = Math.round(2 * (centerX - (ink[0] + ink[1]) / 2) / fontSize * 1000 * strength / 100);
                        kernings.push({ index: lineStart, kerning: kerning });
                    }
                }
                lineStart = lineEnd + 1;
            }
            for (i = 0; i < kernings.length; i++) {
                /* 手動の値が無い位置で読むと例外になる */
                previousKerning = 0;
                try { previousKerning = characters[kernings[i].index].kerning; } catch (readError) {}
                characters[kernings[i].index].kerning = kernings[i].kerning;
                applied.push({ textFrame: textFrame, index: kernings[i].index, previous: previousKerning });
            }
        }

        function btMeasureLineInk(textFrame, lineStart, lineEnd) {
            var dup, outline, bounds, hasGlyphs, i;
            dup = textFrame.duplicate();
            /* 対象行以外の文字を右から消し、補正前の状態で測る（createOutline は複製を消費する）*/
            for (i = dup.characters.length - 1; i >= lineEnd; i--) { dup.characters[i].remove(); }
            for (i = lineStart - 1; i >= 0; i--) { dup.characters[i].remove(); }
            try { dup.characters[0].kerning = 0; } catch (kerningError) {}
            outline = dup.createOutline();
            bounds = outline.geometricBounds;
            hasGlyphs = outline.pageItems.length > 0;
            outline.remove();
            return hasGlyphs ? [bounds[0], bounds[2]] : null;
        }

        function btResetLineStartKerning(selectedItems) {
            var textFrames, applied, characters, lineTexts, lineStart, previousKerning, i, j;
            textFrames = [];
            for (i = 0; i < selectedItems.length; i++) { btCollectTextFrames(selectedItems[i], textFrames); }
            applied = [];
            for (i = 0; i < textFrames.length; i++) {
                if (textFrames[i].kind !== TextType.POINTTEXT || textFrames[i].locked || textFrames[i].hidden) { continue; }
                characters = textFrames[i].characters;
                lineTexts = textFrames[i].contents.split(String.fromCharCode(13));
                lineStart = 0;
                for (j = 0; j < lineTexts.length; j++) {
                    if (lineTexts[j].length > 0) {
                        /* 手動の値が無い位置は読むと例外になる。もともと0なので触らない */
                        previousKerning = null;
                        try { previousKerning = characters[lineStart].kerning; } catch (readError) {}
                        if (previousKerning !== null && previousKerning !== 0) {
                            try {
                                characters[lineStart].kerning = 0;
                                applied.push({ textFrame: textFrames[i], index: lineStart, previous: previousKerning });
                            } catch (writeError) {}
                        }
                    }
                    lineStart += lineTexts[j].length + 1;
                }
            }
            return applied;
        }

        function btRestoreOpticalKerning(applied) {
            var i;
            for (i = applied.length - 1; i >= 0; i--) {
                try { applied[i].textFrame.characters[applied[i].index].kerning = applied[i].previous; } catch (restoreError) {}
            }
        }

        /* 送信するワーカー関数の一覧（追加したらここにも必ず登録する）/ Every worker function shipped to the main engine */
        var WORKER_FUNCS = [
            btAlignSelection, btResolveSelection, btSelectItems, btRunPerArtboard, btAlignItems,
            btGroupItemsByArtboard, btCopyOptions, btApplyJustification, btFinishAlign,
            btMoveItems, btAlignMovedSelection, btRunAlignCommands, btProbeAlignTarget, btSameBounds,
            btMoveToEdgeOrGuide, btRestoreJustification, btEdgeDelta, btOutsetValue, btEdgeValue, btOppositeSide, btFindGuideSnapValue, btGuideEdgeValues, btIsInsideArtboard, btIsAhead,
            btUpdateMarginGuide, btDrawMarginGuide, btMarginArea, btAddGuideRectangle, btAddGuideLines,
            btDivideLines, btPushRowLines, btPushColumnLines, btPushArtboardEdgeLines, btMakeGuide,
            btHasMargin, btGetGuideLayer, btUnlockLayer, btRestoreLayer, btRemoveGuidesByName,
            btApplyGlyphCorrection, btAlignDelta, btAlignStops, btResolveStepMargin, btIsAtAlignTarget, btHasGlyphBoundsTarget,
            btUnionBounds, btGetMeasureBounds, btGetPlainBounds, btGetClipPathBounds, btFindClipPath,
            btGetOutlineBounds, btSafeRemove,
            btGetSymbolGlyphBounds, btCollectionToArray, btCollectNewEntries, btCollectBrokenItems,
            btRemoveBrokenItems, btOutlineTextFramesIn, btCollectTextFrames,
            btGetPaletteState, btGetSelectionKind, btJustificationByName,
            btNudgeSelection, btReadPreferences, btGetPreferenceFlags, btWritePreferences, btPromoteTextRange,
            btGetLayerKey, btSpansMultipleLayers,
            btGetOverlapArea,
            btFindOverlappingArtboardIndex, btFindNearestArtboardIndex, btActivateArtboardForSelection,
            btIsSingleLineTextFrame, btSetJustification,
            btApplyOpticalKerning, btApplyOpticalKerningToFrame, btMeasureLineInk, btResetLineStartKerning, btRestoreOpticalKerning
        ];

        // =========================================
        // メインエンジンへの委譲 / Delegating to the main engine
        // =========================================
        /* 常駐パレットの app は表示中に DOM 接続を失うため、DOM を触る処理は毎回メインエンジンへ送る
           A persistent palette loses its DOM connection while shown, so every DOM touch is delegated */

        /* 実行中フラグ（連打による多重実行を防ぐ）/ Guard against double execution from rapid clicks */
        var isBusy = false;

        /* ワーカーの戻り値マーカーと status ラベルの対応（「OK:CENTER」は showWorkerResult で組み立てる）
           Worker markers mapped to status labels; "OK:CENTER" is composed in showWorkerResult */
        var STATUS_BY_MARKER = {
            OK:         "done",
            MOVED:      "moved",
            NOBOUNDS:   "noBounds",
            NODOC:      "noDocument",
            NOSEL:      "noSelection",
            MULTILAYER: "multipleLayers",
            NOTARGET:   "alignTarget"
        };

        /**
         * 再入防止つきで処理を実行する（連打による多重実行を防ぐ）
         * @param {function} exclusiveTask - 実行する処理
         * @returns {boolean} 実行したら true（実行中で見送ったときは false）
         */
        function runExclusive(exclusiveTask) {
            if (isBusy) { return false; }
            isBusy = true;
            try {
                exclusiveTask();
            } finally {
                isBusy = false;
            }
            return true;
        }

        /**
         * 関数のソースから宣言行〜閉じ括弧行だけを切り出す
         * ExtendScript の toString() は改行を CR で返し、前後のコメント断片を閉じ「*」「/」を落として
         * 巻き込むことがあるため、行区切りを LF に正規化したうえで関数本体だけを取り出す
         * @param {function} targetFunction - 文字列化する関数
         * @returns {string} 関数宣言だけのソース文字列
         */
        function sliceFunctionSource(targetFunction) {
            var sourceLines = String(targetFunction).replace(/\r\n?/g, "\n").split("\n");
            var firstIndex = -1;
            var lastIndex = -1;
            for (var i = 0; i < sourceLines.length; i++) {
                if (firstIndex < 0 && /^\s*function\s/.test(sourceLines[i])) { firstIndex = i; }
                if (firstIndex >= 0 && /^\s*\}[;\s]*$/.test(sourceLines[i])) { lastIndex = i; }
            }
            if (firstIndex < 0) { return String(targetFunction); }
            if (lastIndex < firstIndex) {
                /* 1行で書かれた関数は、その行だけを取り出す / A function written on one line: keep just that line */
                return /\}[;\s]*$/.test(sourceLines[firstIndex]) ? sourceLines[firstIndex] : sourceLines.slice(firstIndex).join("\n");
            }
            return sourceLines.slice(firstIndex, lastIndex + 1).join("\n");
        }

        /* 連結済みのワーカーソースと、その刻印（1回だけ組み立てて使い回す）
           The assembled worker source and its stamp, built once and reused */
        var workerSourceCache = null;
        var workerStampCache = null;

        /**
         * ソースの取り違えを防ぐ刻印を作る（内容が1文字でも変われば別の値になる）
         * @param {string} workerSource - 対象のソース
         * @returns {string} 刻印
         */
        function buildWorkerStamp(workerSource) {
            var checksum = 0;
            for (var i = 0; i < workerSource.length; i++) {
                checksum = (checksum * 31 + workerSource.charCodeAt(i)) % 2147483647;
            }
            return SCRIPT_VERSION + "-" + workerSource.length + "-" + checksum;
        }

        /**
         * ワーカー関数の定義をひとつの文字列にまとめる（2回目以降はキャッシュを返す）
         * @returns {string} 連結したワーカー関数のソース
         */
        function buildWorkerSource() {
            if (workerSourceCache !== null) { return workerSourceCache; }
            var functionSources = [];
            for (var i = 0; i < WORKER_FUNCS.length; i++) {
                functionSources.push(sliceFunctionSource(WORKER_FUNCS[i]));
            }
            workerSourceCache = functionSources.join("\n");
            workerStampCache = buildWorkerStamp(workerSourceCache);
            return workerSourceCache;
        }

        /* メインエンジンに送り込んだはずのワーカー定義の刻印（送り直しの要否を判断する）
           The stamp of the worker source believed to be loaded in the main engine */
        var loadedWorkerStamp = $.global[WORKER_STAMP_KEY] || null;

        /**
         * 本文をメインエンジンで同期実行し、結果を受け取る
         * @param {string} messageBody - メインエンジンで評価する本文
         * @returns {string} 評価結果（応答がなければ null）
         */
        function sendToMainEngine(messageBody) {
            /* 同期送信の結果は resultHolder 経由で受け取る / The synchronous send hands its result back through resultHolder */
            var resultHolder = { result: null };
            var bridgeTalk = new BridgeTalk();
            bridgeTalk.target = "illustrator";
            bridgeTalk.body = messageBody;
            bridgeTalk.onResult = function(response) { resultHolder.result = String(response.body); };
            bridgeTalk.onError = function(response) { resultHolder.result = "ERR:" + String(response.body); };
            bridgeTalk.send(WORKER_TIMEOUT);
            return resultHolder.result;
        }

        /**
         * ワーカー定義ごと送る本文を組み立てる
         * 定義をメインエンジンのグローバルに残したいので、eval はトップレベルで実行する
         * （関数の中で eval するとその関数のローカルになり、次の呼び出しから見えない）
         * @param {string} workerSource - 連結したワーカー関数のソース
         * @param {string} stamp - そのソースの刻印
         * @param {string} functionCall - 評価する呼び出し式
         * @returns {string} 送信する本文
         */
        function buildFullBody(workerSource, stamp, functionCall) {
            /* バックスラッシュ・多バイト文字・改行が途中で壊れないよう、ソースはURIエンコードして送る
               URI-encode the source so backslashes, multi-byte characters and newlines survive the trip */
            /* 全体をひとつの式にして、最後に評価される呼び出しの値がそのまま結果として返るようにする
               One comma expression, so the value of the trailing call is what comes back */
            return "eval(decodeURIComponent(\"" + encodeURIComponent(workerSource) + "\")), " +
                "$.global." + WORKER_STAMP_KEY + " = \"" + stamp + "\", " +
                functionCall;
        }

        /**
         * 呼び出し式だけを送る本文を組み立てる（定義が残っていなければ "RELOAD" を返させる）
         * @param {string} stamp - 送り込んだはずのソースの刻印
         * @param {string} functionCall - 評価する呼び出し式
         * @returns {string} 送信する本文
         */
        function buildCallOnlyBody(stamp, functionCall) {
            return "(function() {" +
                "if ($.global." + WORKER_STAMP_KEY + " !== \"" + stamp + "\") { return \"RELOAD\"; }" +
                "if (typeof btAlignSelection !== \"function\") { return \"RELOAD\"; }" +
                "return " + functionCall +
                "})();";
        }

        /**
         * ワーカー関数の呼び出しをメインエンジンで同期実行し、戻り値のマーカーを受け取る
         * 定義はメインエンジンのグローバルに残るので、2回目以降は呼び出し式だけを送る
         * （毎回ソースを丸ごと送ると、選択を問い合わせるたびに10KB近いやり取りになる）
         * @param {string} functionCall - メインエンジンで評価する呼び出し式
         * @returns {string} ワーカーが返したマーカー（応答がなければ null）
         */
        function runWorker(functionCall) {
            var workerSource = buildWorkerSource();
            var stamp = workerStampCache;
            try {
                var workerResult = null;
                var needsFullBody = (loadedWorkerStamp !== stamp);
                if (!needsFullBody) {
                    workerResult = sendToMainEngine(buildCallOnlyBody(stamp, functionCall));
                    /* 定義が消えていた（別のエンジンが再起動した）ときは送り直す
                       Ship the source again when the definitions are gone */
                    needsFullBody = (workerResult === "RELOAD");
                }
                if (needsFullBody) {
                    workerResult = sendToMainEngine(buildFullBody(workerSource, stamp, functionCall));
                    if (workerResult !== null) {
                        loadedWorkerStamp = stamp;
                        $.global[WORKER_STAMP_KEY] = stamp;
                    }
                }
                return workerResult;
            } catch (bridgeError) {
                /* BridgeTalk が使えない環境では、このエンジンで直接実行する
                   Fallback: run in this engine when BridgeTalk is unavailable */
                try {
                    return String(eval(workerSource + "\n" + functionCall));
                } catch (evalError) {
                    return "ERR:" + evalError;
                }
            }
        }

        /**
         * 数値の入力欄を読んで pt に換算する（数値以外と負数は0に丸め、欄の表示もそろえる）
         * @param {EditText} valueField - 読み取る入力欄
         * @param {number} pointsPerUnit - 1単位あたりの pt
         * @returns {number} 入力値（pt）
         */
        function readFieldPt(valueField, pointsPerUnit) {
            if (valueField === null) { return 0; }
            var fieldValue = Number(valueField.text);
            if (isNaN(fieldValue) || fieldValue < 0) { fieldValue = 0; }
            /* 手入力が丸められたときは欄の表示も実際に使う値にそろえる / Show the value actually used */
            if (String(fieldValue) !== valueField.text) { valueField.text = fieldValue; }
            return fieldValue * pointsPerUnit;
        }

        /**
         * マージン4欄の表示中の文字列を辺ごとに読む（控え用）
         * @returns {object} { top: string, bottom: string, left: string, right: string }
         */
        function readMarginTexts() {
            var marginTexts, side;
            marginTexts = {};
            for (side in marginFields) {
                if (!marginFields.hasOwnProperty(side)) { continue; }
                marginTexts[side] = String(marginFields[side].text);
            }
            return marginTexts;
        }

        /**
         * 分割ガイドの数値欄を、入力されたままの文字列で読む（控え用）
         * @returns {object} DIVIDE_VALUE_KEYS をキーにした文字列の組
         */
        function readDivideValueTexts() {
            var divideTexts, valueKey, i;
            divideTexts = {};
            for (i = 0; i < DIVIDE_VALUE_KEYS.length; i++) {
                valueKey = DIVIDE_VALUE_KEYS[i];
                divideTexts[valueKey] = (divideFields[valueKey]) ?
                    String(divideFields[valueKey].text) : defaultDivideValue(valueKey);
            }
            return divideTexts;
        }

        /**
         * 分割数の欄を読む（読めない値と1未満は1＝分割なし、上限は DIVIDE_COUNT_MAX）
         * @param {string} valueKey - "rows" または "columns"
         * @returns {number} 分割数
         */
        function readDivideCount(valueKey) {
            var count = Math.floor(Number(divideFields[valueKey] ? divideFields[valueKey].text : ""));
            if (isNaN(count) || count < 1) { return 1; }
            return (count > DIVIDE_COUNT_MAX) ? DIVIDE_COUNT_MAX : count;
        }

        /**
         * 分割ガイドの設定を読んでワーカーに渡す形にする
         * ［なし］は分割なし、［十字］は縦横2等分（伸張はどちらにも効く）
         * @returns {object} { rows: number, columns: number, rowGutter: number, columnGutter: number, extension: number }（長さは pt）
         */
        function readDivisions() {
            var divideMode, divisions;
            divideMode = readDivideMode();
            divisions = {
                rows:         1,
                columns:      1,
                rowGutter:    0,
                columnGutter: 0,
                extension:    readFieldPt(divideFields.extension || null, currentUnitInfo.points),
                artboardEdge: readArtboardEdge()
            };
            if (divideMode === "cross") {
                divisions.rows    = CROSS_DIVIDE_COUNTS.rows;
                divisions.columns = CROSS_DIVIDE_COUNTS.columns;
                return divisions;
            }
            if (divideMode === "custom") {
                divisions.rows         = readDivideCount("rows");
                divisions.columns      = readDivideCount("columns");
                divisions.rowGutter    = readFieldPt(divideFields.rowGutter || null, currentUnitInfo.points);
                divisions.columnGutter = readFieldPt(divideFields.columnGutter || null, currentUnitInfo.points);
            }
            return divisions;
        }

        /**
         * マージン4欄を読んで pt に換算する
         * @returns {object} { top: number, bottom: number, left: number, right: number }（pt）
         */
        function readMarginsPt() {
            var marginsPt, side, i;
            marginsPt = {};
            for (i = 0; i < MARGIN_SIDES.length; i++) {
                side = MARGIN_SIDES[i];
                marginsPt[side] = readFieldPt(marginFields[side] || null, currentUnitInfo.points);
            }
            return marginsPt;
        }

        /**
         * その整列が使うマージンを、寄せる辺から1つ選ぶ（中央揃えは使わないので0）
         * @param {object} alignSpec - readAlignSpec() の戻り値
         * @param {object} marginsPt - readMarginsPt() の戻り値
         * @returns {number} マージン（pt）
         */
        function marginForAlign(alignSpec, marginsPt) {
            if (alignSpec.modeX === "start") { return marginsPt.left; }
            if (alignSpec.modeX === "end") { return marginsPt.right; }
            if (alignSpec.modeY === "start") { return marginsPt.top; }
            if (alignSpec.modeY === "end") { return marginsPt.bottom; }
            return 0;
        }

        /**
         * 裁ち落としの距離を pt で読む（［裁ち落としに整列］がOFFなら0）
         * @returns {number} 裁ち落とし（pt）
         */
        function readBleedPt() {
            if (alignToBleedCheckbox === null || alignToBleedCheckbox.value !== true) { return 0; }
            return BLEED_MM * BLEED_UNIT_POINTS;
        }

        /**
         * ボタン定義から、そのクリックに必要な値をまとめて求める
         * 実行するメニューコマンド、整列後にマージンぶん動かす向き（中央揃えは0）、整列先の判定に使う仮移動量、
         * 軸ごとの寄せ先（字形の境界での補正が目標位置の計算に使う）、合わせる行揃えを、ALIGN_COMMANDS の1周で得る
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {object} { commands: string[], offsetX: number, offsetY: number, probeX: number, probeY: number, modeX: string, modeY: string, justification: string }
         */
        function readAlignSpec(buttonDef) {
            var alignSpec = {
                commands: [],
                offsetX: 0,
                offsetY: 0,
                probeX: 0,
                probeY: 0,
                modeX: null,
                modeY: null,
                justification: null
            };
            for (var i = 0; i < buttonDef.alignKeys.length; i++) {
                var alignCommand = ALIGN_COMMANDS[buttonDef.alignKeys[i]];
                alignSpec.commands.push(alignCommand.command);
                alignSpec.offsetX += alignCommand.offsetX;
                alignSpec.offsetY += alignCommand.offsetY;
                /* 仮移動は端揃えなら内側へ、中央揃えは向きがないので＋方向へ（整列が空振りしたときだけ使う）
                   The probe moves inward for an edge align; a center align has no direction, so it uses the plus side */
                if (alignCommand.axis === "x") {
                    alignSpec.probeX = (alignCommand.offsetX !== 0 ? alignCommand.offsetX : 1) * ALIGN_PROBE_PT;
                    alignSpec.modeX = alignCommand.mode;
                } else {
                    alignSpec.probeY = (alignCommand.offsetY !== 0 ? alignCommand.offsetY : 1) * ALIGN_PROBE_PT;
                    alignSpec.modeY = alignCommand.mode;
                }
                /* 行揃えは水平方向の整列にだけ付いているので、最初に見つかったものを使う
                   Only horizontal aligns carry a justification, so the first one found wins */
                if (alignSpec.justification === null && alignCommand.justification) {
                    alignSpec.justification = alignCommand.justification;
                }
            }
            return alignSpec;
        }

        /**
         * 中央に寄せる軸を持つ整列か判定する
         * @param {object} alignSpec - readAlignSpec() の戻り値
         * @returns {boolean} 片方でも中央に寄せるなら true
         */
        function isCenterAlign(alignSpec) {
            return alignSpec.modeX === "center" || alignSpec.modeY === "center";
        }

        /**
         * そのボタンの Option＋クリックの説明を選ぶ
         * 中央揃えはアートボードの中央へ、端に寄せる整列はマージンを無視して字形の境界に寄せる
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {string} tooltip のキー
         */
        function optionTooltipKey(buttonDef) {
            var alignSpec = readAlignSpec(buttonDef);
            if (isCenterAlign(alignSpec)) { return "optionArtboardCenter"; }
            /* マージンの影響を受けるのは上下左右に寄せるボタンだけ / Only the edge alignments use a margin */
            return (alignSpec.offsetX !== 0 || alignSpec.offsetY !== 0) ? "optionNoMargin" : "optionGlyphBounds";
        }

        /**
         * マージンの内側の中央へ寄せるための、アートボードの中央からのずれを求める
         * 中央に寄せる軸だけ、向かい合うマージンの差の半分だけ内側へずらす
         * @param {object} alignSpec - readAlignSpec() の戻り値
         * @param {object} marginsPt - readMarginsPt() の戻り値
         * @returns {object} { x: number, y: number }（pt）
         */
        function centerOffsetInMargin(alignSpec, marginsPt) {
            return {
                x: (alignSpec.modeX === "center") ? (marginsPt.left - marginsPt.right) / 2 : 0,
                y: (alignSpec.modeY === "center") ? (marginsPt.bottom - marginsPt.top) / 2 : 0
            };
        }

        /**
         * ガイドの作成に必要な値を組み立てる
         * マージンは Option＋クリックの影響を受けない（ガイドは入力欄の値をそのまま表す）
         * @returns {object} ワーカーへ渡すガイドのオプション
         */
        function buildGuideOptions() {
            return {
                showGuide:      showGuideCheckbox !== null && showGuideCheckbox.value === true,
                guideMargins:   readMarginsPt(),
                divisions:      readDivisions(),
                perArtboard:    alignPerArtboardCheckbox !== null && alignPerArtboardCheckbox.value === true,
                guideName:      GUIDE_NAME,
                divideName:     DIVIDE_GUIDE_NAME,
                guideLayerName: GUIDE_LAYER_NAME
            };
        }

        /**
         * マージンのガイドを作り直す（チェックがOFFなら消すだけ）
         * 結果は状況表示に出さず、エラーのときだけ知らせる
         * @returns {void}
         */
        function refreshMarginGuide() {
            if (showGuideCheckbox === null) { return; }
            var workerResult = runWorker("btUpdateMarginGuide(" + buildGuideOptions().toSource() + ");");
            if (workerResult !== null && workerResult.indexOf("ERR:") === 0) { showWorkerResult(workerResult); }
        }

        /**
         * ワーカーに渡すオプションを組み立てる（パレット側の状態はすべてここで値にする）
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {object} ワーカーへ渡すオプション
         */
        function buildAlignOptions(buttonDef) {
            var alignSpec = readAlignSpec(buttonDef);
            /* Option＋クリックの意味はボタンによって変わる
               中央揃え：マージンの内側ではなくアートボードの中央へ寄せる
               上下左右：字形の境界をONにしたうえでマージンを無視し、アートボードの辺にぴったり寄せる
               Option-click means different things per button: centre on the artboard instead of inside
               the margin for the centred alignments, and flush against the artboard edge with glyph
               bounds for the edge ones */
            var altPressed = isAltPressed();
            var centersInMargin = isCenterAlign(alignSpec) && !altPressed;
            var ignoresMargin = altPressed && !isCenterAlign(alignSpec);
            var alignOptions = buildGuideOptions();
            var centerOffset = centersInMargin ?
                centerOffsetInMargin(alignSpec, alignOptions.guideMargins) : { x: 0, y: 0 };
            alignOptions.alignCommands       = alignSpec.commands;
            alignOptions.offsetX             = alignSpec.offsetX;
            alignOptions.offsetY             = alignSpec.offsetY;
            alignOptions.probeX              = alignSpec.probeX;
            alignOptions.probeY              = alignSpec.probeY;
            alignOptions.modeX               = alignSpec.modeX;
            alignOptions.modeY               = alignSpec.modeY;
            /* 中央からのずれ。Option を押していない中央揃えだけ 0 以外になる
               The shift from the artboard centre; only a centre alignment without Option sets it */
            alignOptions.centerOffsetX       = centerOffset.x;
            alignOptions.centerOffsetY       = centerOffset.y;
            /* Option＋クリックは段階を踏まず、アートボードのエッジにぴったり寄せる
               Option-click skips the steps and sits flush against the artboard edge */
            alignOptions.marginPt            = ignoresMargin ? 0 : marginForAlign(alignSpec, alignOptions.guideMargins);
            alignOptions.bleedPt             = ignoresMargin ? 0 : readBleedPt();
            alignOptions.perArtboard         = alignPerArtboardCheckbox !== null && alignPerArtboardCheckbox.value === true;
            alignOptions.minDeltaPt          = MOVE_MIN_DELTA_PT;
            alignOptions.previewBounds       = previewBoundsCheckbox !== null && previewBoundsCheckbox.value === true;
            alignOptions.glyphBounds         = ignoresMargin || (glyphBoundsCheckbox !== null && glyphBoundsCheckbox.value === true);
            alignOptions.changeJustification = changeJustificationCheckbox !== null && changeJustificationCheckbox.value === true;
            alignOptions.justification       = alignSpec.justification;
            alignOptions.opticalAdjust       = opticalAdjustCheckbox !== null && opticalAdjustCheckbox.value === true;
            alignOptions.opticalStrength     = OPTICAL_KERNING_STRENGTH;
            return alignOptions;
        }

        /**
         * 移動ボタンのオプションを組み立てる
         * マージンは使わず（ガイドかアートボードのエッジにぴったり寄せる）、境界の測り方だけオプションに従う
         * @param {string} directionKey - MOVE_SIDES_BY_DIRECTION のキー
         * @returns {object} ワーカーへ渡すオプション
         */
        function buildMoveOptions(directionKey) {
            var sides = MOVE_SIDES_BY_DIRECTION[directionKey];
            return {
                sides:          sides,
                bleedPt:        readBleedPt(),
                perArtboard:    alignPerArtboardCheckbox !== null && alignPerArtboardCheckbox.value === true,
                previewBounds:  previewBoundsCheckbox !== null && previewBoundsCheckbox.value === true,
                glyphBounds:    glyphBoundsCheckbox !== null && glyphBoundsCheckbox.value === true,
                changeJustification: changeJustificationCheckbox !== null && changeJustificationCheckbox.value === true,
                justification:  justificationForSides(sides),
                guideTolerance: GUIDE_ORIENTATION_TOLERANCE,
                minDeltaPt:     MOVE_MIN_DELTA_PT
            };
        }

        /**
         * その移動で合わせる行揃えを求める（左右へ寄せるときだけ。上下では変えない）
         * @param {string[]} sides - 寄せる辺
         * @returns {string} 行揃えの名前。合わせないときは null
         */
        function justificationForSides(sides) {
            for (var i = 0; i < sides.length; i++) {
                if (sides[i] === "left") { return ALIGN_COMMANDS.horizontalLeft.justification; }
                if (sides[i] === "right") { return ALIGN_COMMANDS.horizontalRight.justification; }
            }
            return null;
        }

        /**
         * ワーカーの戻り値を状況表示に反映する
         * @param {string} workerResult - ワーカーが返したマーカー
         * @returns {void}
         */
        function showWorkerResult(workerResult) {
            if (workerResult === null) {
                setStatus(getLabel("status.noResponse"));
                return;
            }
            if (workerResult.indexOf("ERR:") === 0) {
                setStatus(getLabel("status.genericError") + workerResult.substring(4));
                return;
            }
            /* 「OK:CENTER」のように行揃えを変えたときは、変更後の行揃えも知らせる
               An "OK:CENTER" marker also reports the justification that was applied */
            if (workerResult.indexOf("OK:") === 0) {
                showJustifiedResult(workerResult.substring(3), "doneJustified", "done");
                return;
            }
            if (workerResult.indexOf("MOVED:") === 0) {
                showJustifiedResult(workerResult.substring(6), "movedJustified", "moved");
                return;
            }
            var statusKey = STATUS_BY_MARKER[workerResult];
            setStatus(statusKey ? getLabel("status." + statusKey) : workerResult);
        }

        /**
         * 行揃えを変えたときの状況表示を出す（名前が引けなければ、変えなかったときと同じ文言にする）
         * @param {string} justificationName - ワーカーが返した行揃えの名前
         * @param {string} justifiedKey - 行揃えを添える文言のキー
         * @param {string} plainKey - 添えないときの文言のキー
         * @returns {void}
         */
        function showJustifiedResult(justificationName, justifiedKey, plainKey) {
            var justificationLabel = getLabel("status.justification." + justificationName);
            setStatus(justificationLabel ?
                getLabel("status." + justifiedKey).replace("{0}", justificationLabel) :
                getLabel("status." + plainKey));
        }

        /* 取り直しの状態 / State of the refresh */
        var isRefreshingSelection = false;
        var lastSelectionRefreshTime = 0;
        /* 直前の選択（種類＋個数）。変わったら前回の実行結果の表示を消す
           The previous selection (kind and count); a change clears the last result from the status line */
        var lastSelectionSignature = null;

        /**
         * 選択の種類・定規の単位・選択数・アートボード数をメインエンジンに問い合わせ、ディムと単位ラベルを更新する
         * 1往復でまとめて受け取り（"TEXT|2|3|1" の形）、選択が変わっていれば状況表示も消す
         * @returns {void}
         */
        function refreshPaletteState() {
            /* 整列の実行中や取り直しの最中に割り込ませない（同期送信の待ち時間にイベントが入り得るため）
               Never nest inside a running align or another refresh; events can fire while the send waits */
            if (isBusy || isRefreshingSelection) { return; }
            isRefreshingSelection = true;
            try {
                var workerResult = runWorker("btGetPaletteState();");
                /* 応答なし・エラーのときは、当てにならない値で表示を書き換えない
                   Leave the palette as it is when there is no usable answer */
                if (workerResult === null || workerResult.indexOf("ERR:") === 0) { return; }
                var stateParts = workerResult.split("|");
                if (changeJustificationCheckbox !== null) {
                    changeJustificationCheckbox.enabled = (stateParts[0] === "TEXT");
                }
                if (opticalAdjustCheckbox !== null) {
                    opticalAdjustCheckbox.enabled = (stateParts[0] === "TEXT");
                }
                /* アートボードが1つしかないなら束ね分ける先が無いので、［アートボードごとに整列］はディムにする */
                if (alignPerArtboardCheckbox !== null) {
                    alignPerArtboardCheckbox.enabled = (Number(stateParts[3]) > 1);
                }
                currentUnitInfo = getUnitInfoWithDefaults();
                if (marginPanel !== null) { marginPanel.text = panelTitleWithUnit("guide"); }
                if (dividePanel !== null) { dividePanel.text = panelTitleWithUnit("divide"); }
                fillDefaultExtension();
                fillDefaultMargins();
                var selectionSignature = stateParts[0] + "|" + stateParts[2];
                if (lastSelectionSignature !== null && selectionSignature !== lastSelectionSignature) {
                    setStatus("");
                }
                lastSelectionSignature = selectionSignature;
            } finally {
                isRefreshingSelection = false;
            }
        }

        /**
         * パレットへフォーカスが来たときに選択と定規の単位を取り直す
         * Illustrator にタイマーAPIが無いため、変化はこの瞬間に拾う
         * @param {boolean} force - true なら間引きを無視して必ず取り直す
         * @returns {void}
         */
        function onPaletteFocus(force) {
            var now = (new Date()).getTime();
            if (!force && (now - lastSelectionRefreshTime) < SELECTION_POLL_INTERVAL_MS) { return; }
            lastSelectionRefreshTime = now;
            refreshPaletteState();
        }

        /**
         * 整列をメインエンジンへ委譲する
         * @param {object} buttonDef - 整列ボタンの定義
         * @returns {string} ワーカーが返したマーカー（応答がなければ null）
         */
        function runAlign(buttonDef) {
            var alignOptions = buildAlignOptions(buttonDef);
            return runWorker("btAlignSelection(" + alignOptions.toSource() + ");");
        }

        /**
         * 端またはガイドへの移動をメインエンジンへ委譲する
         * @param {string} directionKey - "up" / "left" / "right" / "down"
         * @returns {string} ワーカーが返したマーカー（応答がなければ null）
         */
        function runMove(directionKey) {
            var moveOptions = buildMoveOptions(directionKey);
            return runWorker("btMoveToEdgeOrGuide(" + moveOptions.toSource() + ");");
        }

        showPalette();

    })();

})();
