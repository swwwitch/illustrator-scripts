#target illustrator
#targetengine "VTRainMakerEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

雨のような罫線を生成します。
雨の種類（雨粒・春雨・五月雨・鉄砲雨）や粒の形状、線幅・角度・密度・本数を指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/VTRainMaker.md

### Overview

Generates rain-like rules.
The rain type (raindrop, spring rain, early-summer rain, driving rain), the drop shape, and the stroke weight, angle, density and count are all configurable.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/VTRainMaker.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "VTRainMaker";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.6";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-01";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/VTRainMaker.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/VTRainMaker.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "雨の罫線",
            en: "Rain Lines"
        },
        noDocument: {
            ja: "ドキュメントを開いてから実行してください。",
            en: "Open a document before running this script."
        },
        panelRainType: {
            ja: "雨の種類",
            en: "Rain Type"
        },
        panelOptions: {
            ja: "オプション",
            en: "Options"
        },
        panelRaindrop: {
            ja: "雨粒",
            en: "Raindrop"
        },
        rainHarusame: {
            ja: "春雨",
            en: "Harusame"
        },
        rainSamidare: {
            ja: "五月雨",
            en: "Samidare"
        },
        rainTeppouame: {
            ja: "鉄砲雨",
            en: "Driving Rain"
        },
        rainAmatsubu: {
            ja: "雨粒",
            en: "Raindrops"
        },
        rainMizutama: {
            ja: "水玉",
            en: "Polka Dots"
        },
        labelStrokeWidth: {
            ja: "線幅",
            en: "Stroke Width"
        },
        labelAngle: {
            ja: "角度",
            en: "Angle"
        },
        labelDensity: {
            ja: "密度",
            en: "Density"
        },
        unitLines: {
            ja: "本数",
            en: "Lines"
        },
        unitDegree: {
            ja: "°",
            en: "°"
        },
        labelMargin: {
            ja: "外側拡張",
            en: "Margin Expansion"
        },
        labelSpacing: {
            ja: "子罫オフセット",
            en: "Child Line Offset"
        },
        labelScale: {
            ja: "全体スケール",
            en: "Overall Scale"
        },
        unitPercent: {
            ja: "%",
            en: "%"
        },
        chkRaindropFill: {
            ja: "塗り",
            en: "Fill"
        },
        chkRaindropStroke: {
            ja: "線",
            en: "Stroke"
        },
        raindropShapeA: {
            ja: "A",
            en: "A"
        },
        raindropShapeB: {
            ja: "B",
            en: "B"
        },
        raindropShapeC: {
            ja: "C",
            en: "C"
        },
        invalidScale: {
            ja: "全体スケールには25〜300の数値を入力してください。",
            en: "Enter a value between 25 and 300 for the overall scale."
        },
        invalidAngle: {
            ja: "角度には数値を入力してください。",
            en: "Enter a numeric value for the angle."
        },
        preview: {
            ja: "プレビュー",
            en: "Preview"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        invalidStrokeWidth: {
            ja: "線幅には正の数値を入力してください。",
            en: "Enter a positive number for the stroke width."
        },
        invalidDensity: {
            ja: "密度には正の整数を入力してください。",
            en: "Enter a positive integer for the density."
        },
        invalidSpacing: {
            ja: "子罫オフセットには0以上の数値を入力してください。",
            en: "Enter a number greater than or equal to 0 for the child offset."
        },
        invalidMargin: {
            ja: "外側拡張には0以上の数値を入力してください。",
            en: "Enter a number greater than or equal to 0 for the margin expansion."
        },
        tipHarusame: {
            ja: "細く短い線を降らせます。",
            en: "Draws fine, short lines."
        },
        tipSamidare: {
            ja: "中くらいの長さの線を降らせます。",
            en: "Draws medium-length lines."
        },
        tipTeppouame: {
            ja: "太く長い線を降らせます。",
            en: "Draws thick, long lines."
        },
        tipAmatsubu: {
            ja: "しずくの形を散らします。",
            en: "Scatters droplet shapes."
        },
        tipMizutama: {
            ja: "丸い粒を散らします。",
            en: "Scatters round drops."
        },
        tipScale: {
            ja: "線幅・長さ・間隔をまとめて拡大・縮小します（25〜300%）。",
            en: "Scales stroke width, length, and spacing together (25-300%)."
        },
        tipStrokeWidth: {
            ja: "1本あたりの線の太さです。",
            en: "Thickness of each line."
        },
        tipAngle: {
            ja: "雨が降る向きです。0度で垂直になります。",
            en: "Direction of the rain. 0 degrees is vertical."
        },
        tipDensity: {
            ja: "描く本数です。多いほど雨が密になります。",
            en: "Number of lines to draw. More lines make denser rain."
        },
        tipSpacing: {
            ja: "1本の親罫に寄り添う子罫をずらす距離です。0でずらしません。",
            en: "Distance the companion line is offset from its parent line. 0 draws no offset."
        },
        tipMargin: {
            ja: "対象範囲の外へ描き足す幅です。端で雨が途切れるのを防ぎます。",
            en: "How far to draw beyond the target area, so the rain is not cut off at the edges."
        },
        tipRaindropFill: {
            ja: "しずく・水玉を塗りで描きます。",
            en: "Draws droplets and drops with a fill."
        },
        tipRaindropStroke: {
            ja: "しずく・水玉を線で描きます。",
            en: "Draws droplets and drops with a stroke."
        },
        tipRaindropShape: {
            ja: "しずく・水玉の形のバリエーションです。",
            en: "Shape variation for the droplets and drops."
        },
        tipPreview: {
            ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
            en: "Shows the result on the canvas. Cancel restores the original state."
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
            stepUpInteger: { ja: "値を増やす（shift＋クリックで10の倍数へ）", en: "Increase (Shift-click to snap to 10s)" },
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        }
    };

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
    // 単位ユーティリティ / Unit utilities
    // =========================================
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

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する。
     * @param {string} prefKey - 環境設定キー
     * @returns {object} code / label / pointsPerUnit を持つオブジェクト
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /* pt 値を現在の単位値へ変換 / Convert a pt value to the current display unit */
    function ptToUnitValue(ptValue, prefKey, decimals) {
        var info = getUnitInfo(prefKey);
        var unitValue = ptValue / info.pointsPerUnit;
        if (typeof decimals === "number") {
            return parseFloat(unitValue.toFixed(decimals));
        }
        return unitValue;
    }

    /* 現在の単位値を pt 値へ変換 / Convert a current-unit value to pt */

    function unitValueToPt(unitValue, prefKey) {
        var info = getUnitInfo(prefKey);
        return unitValue * info.pointsPerUnit;
    }

    function parseMarginPtFromText(text) {
        return unitValueToPt(parseFloat(text), "rulerType");
    }

    function parseScaleValueFromText(text) {
        return parseFloat(text);
    }

    function isValidMarginPt(marginPt) {
        return !isNaN(marginPt) && marginPt >= 0;
    }

    function clampScaleValue(scaleValue) {
        if (scaleValue < 25) return 25;
        if (scaleValue > 300) return 300;
        return scaleValue;
    }

    function isValidScaleValue(scaleValue) {
        return !isNaN(scaleValue) && scaleValue >= 25 && scaleValue <= 300;
    }

    function isRectanglePathItem(item) {
        if (!item || item.typename !== "PathItem") return false;
        if (!item.closed || item.guides || item.clipping) return false;
        if (item.pathPoints.length !== 4) return false;

        var gb = item.geometricBounds;
        var left = gb[0];
        var top = gb[1];
        var right = gb[2];
        var bottom = gb[3];
        var eps = 0.01;
        var hasLeft = false;
        var hasRight = false;
        var hasTop = false;
        var hasBottom = false;

        for (var i = 0; i < item.pathPoints.length; i++) {
            var anchor = item.pathPoints[i].anchor;
            var x = anchor[0];
            var y = anchor[1];

            if (Math.abs(x - left) <= eps) hasLeft = true;
            if (Math.abs(x - right) <= eps) hasRight = true;
            if (Math.abs(y - top) <= eps) hasTop = true;
            if (Math.abs(y - bottom) <= eps) hasBottom = true;
        }

        return hasLeft && hasRight && hasTop && hasBottom;
    }

    function getTargetBounds(doc) {
        var target = null;
        if (doc.selection && doc.selection.length === 1 && isRectanglePathItem(doc.selection[0])) {
            target = doc.selection[0].geometricBounds;
        } else {
            target = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        }

        return {
            left: target[0],
            top: target[1],
            right: target[2],
            bottom: target[3]
        };
    }

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

    function drawRandomParentChildLines() {
        if (app.documents.length === 0) {
            alert(getLabel("noDocument"));
            return;
        }

        var doc = app.activeDocument;

        // 対象範囲を取得（選択長方形があればその範囲、なければ現在のアートボード）
        var targetBounds = getTargetBounds(doc);
        var left = targetBounds.left;
        var top = targetBounds.top;
        var right = targetBounds.right;
        var bottom = targetBounds.bottom;

        var width = right - left;
        var height = top - bottom;

        var layer = doc.activeLayer;

        var rulerUnitInfo = getUnitInfo("rulerType");
        var strokeUnitInfo = getUnitInfo("strokeUnits");

        function parseMarginPt() {
            return parseMarginPtFromText(inputMargin.text);
        }

        function parseScaleValue() {
            return parseScaleValueFromText(inputScale.text);
        }

        var strokeColor = new CMYKColor();
        strokeColor.cyan = 0;
        strokeColor.magenta = 0;
        strokeColor.yellow = 0;
        strokeColor.black = 100;

        /* プリセット定義 / Preset definitions */
        // { lineLength, angleDeg, defaultStrokeWidth, defaultAngle, defaultDensity, defaultSpacing, defaultScale, roundCap, noChild, colors, penetrate, raindrop, polkadot }
        var presets = {
            harusame: { lineLength: 60, angleDeg: 45, defaultStrokeWidth: "2.0", defaultAngle: "45", defaultDensity: "50", defaultSpacing: "3", defaultScale: "100", avoidOverlap: true, placementGap: 0.95 },
            samidare: { lineLength: 80, angleDeg: 45, defaultStrokeWidth: "10.0", defaultAngle: "45", defaultDensity: "35", defaultSpacing: "0", defaultScale: "100", roundCap: true, noChild: true, colors: [[66, 122, 190], [207, 167, 204]], avoidOverlap: true, placementGap: 1.05 },
            teppouame: { lineLength: 0, angleDeg: 0, defaultStrokeWidth: "100", defaultAngle: "45", defaultDensity: "15", defaultSpacing: "0", defaultScale: "100", noChild: true, penetrate: true },
            amatsubu: { defaultStrokeWidth: "8.0", defaultAngle: "0", defaultDensity: "40", defaultSpacing: "0", defaultScale: "100", noChild: true, raindrop: true },
            mizutama: { defaultStrokeWidth: "8.0", defaultAngle: "0", defaultDensity: "35", defaultSpacing: "0", defaultScale: "100", noChild: true, polkadot: true, colors: [[177, 208, 236], [102, 197, 236]] }
        };

        /* プレビュー用グループの参照 / References to preview groups */
        var previewGroups = [];

        /* 現在のプリセットを取得 / Get the current preset */
        function getSelectedPreset() {
            if (rbSamidare.value) return presets.samidare;
            if (rbTeppouame.value) return presets.teppouame;
            if (rbAmatsubu.value) return presets.amatsubu;
            if (rbMizutama.value) return presets.mizutama;
            return presets.harusame;
        }

        /* プリセットのcolorsからランダムに1色選択 / Pick one random color from preset colors */
        function pickColor(colors) {
            var rgb = colors[Math.floor(Math.random() * colors.length)];
            var c = new RGBColor();
            c.red = rgb[0];
            c.green = rgb[1];
            c.blue = rgb[2];
            return c;
        }

        /* 雨粒用の塗り色を作成 / Create a fill color for raindrops */
        function createRaindropColor() {
            var c = new RGBColor();
            c.red = 0;
            c.green = 150;
            c.blue = 255;
            return c;
        }

        /* 雨粒用の線色を作成 / Create a stroke color for raindrops */
        function createRaindropStrokeColor() {
            var c = new RGBColor();
            c.red = 0;
            c.green = 100;
            c.blue = 200;
            return c;
        }

        /* ティアドロップ形状を作成 / Create a teardrop shape */
        function createRaindropShape(parentGroup, centerX, centerY, widthPt, heightPt, fillColor, angleDeg, styleOptions) {
            function rotatePoint(px, py) {
                var rad = (angleDeg || 0) * Math.PI / 180;
                var dx = px - centerX;
                var dy = py - centerY;
                var cosA = Math.cos(rad);
                var sinA = Math.sin(rad);
                return [
                    centerX + dx * cosA - dy * sinA,
                    centerY + dx * sinA + dy * cosA
                ];
            }

            var opt = {
                heightRatio: 1.8,
                midRatio: 0.55,
                bottomRatio: 0.80,
                downHandleScale: 0.35,
                topHandleLength: 0.40,
                topHandleInset: 0.15,
                strokeWidth: 2,
                filled: true,
                stroked: false,
                strokeColor: createRaindropStrokeColor()
            };

            if (styleOptions) {
                for (var key in styleOptions) {
                    if (styleOptions.hasOwnProperty(key)) {
                        opt[key] = styleOptions[key];
                    }
                }
            }

            var width = widthPt;
            var height = (typeof heightPt === "number" && heightPt > 0) ? heightPt : width * opt.heightRatio;
            var hw = width / 2;
            var hh = height / 2;

            var cx = centerX;
            var cy = centerY + hh * 0.80;

            var top = cy;
            var bottom = cy - height * opt.bottomRatio;
            var mid = cy - height * opt.midRatio;

            var k = 0.5523;
            var rx = hw;
            var ry = hh * 0.85;

            var pathPoints = [
                {
                    anchor: [cx, top],
                    leftDir: [cx - hw * opt.topHandleInset, top - hh * opt.topHandleLength],
                    rightDir: [cx + hw * opt.topHandleInset, top - hh * opt.topHandleLength],
                    pointType: PointType.CORNER
                },
                {
                    anchor: [cx + rx, mid],
                    leftDir: [cx + rx, mid + ry * k],
                    rightDir: [cx + rx, mid - ry * opt.downHandleScale],
                    pointType: PointType.SMOOTH
                },
                {
                    anchor: [cx, bottom],
                    leftDir: [cx + rx * k, bottom],
                    rightDir: [cx - rx * k, bottom],
                    pointType: PointType.SMOOTH
                },
                {
                    anchor: [cx - rx, mid],
                    leftDir: [cx - rx, mid - ry * opt.downHandleScale],
                    rightDir: [cx - rx, mid + ry * k],
                    pointType: PointType.SMOOTH
                }
            ];

            var item = parentGroup.pathItems.add();
            item.stroked = !!opt.stroked;
            if (item.stroked) {
                item.strokeColor = opt.strokeColor;
                item.strokeWidth = opt.strokeWidth;
            }
            item.filled = !!opt.filled;
            if (item.filled) {
                item.fillColor = fillColor;
            }
            item.closed = true;
            item.setEntirePath([
                rotatePoint(pathPoints[0].anchor[0], pathPoints[0].anchor[1]),
                rotatePoint(pathPoints[1].anchor[0], pathPoints[1].anchor[1]),
                rotatePoint(pathPoints[2].anchor[0], pathPoints[2].anchor[1]),
                rotatePoint(pathPoints[3].anchor[0], pathPoints[3].anchor[1])
            ]);

            var pts = item.pathPoints;
            for (var i = 0; i < pts.length; i++) {
                pts[i].leftDirection = rotatePoint(pathPoints[i].leftDir[0], pathPoints[i].leftDir[1]);
                pts[i].rightDirection = rotatePoint(pathPoints[i].rightDir[0], pathPoints[i].rightDir[1]);
                pts[i].pointType = pathPoints[i].pointType;
            }

            return item;
        }

        /* 水玉を作成 / Create a polka dot */
        function createPolkaDot(parentGroup, centerX, centerY, diameterPt, fillColor) {
            var ellipse = parentGroup.pathItems.ellipse(
                centerY + diameterPt / 2,
                centerX - diameterPt / 2,
                diameterPt,
                diameterPt
            );
            ellipse.stroked = false;
            ellipse.filled = true;
            ellipse.fillColor = fillColor;
            return ellipse;
        }

        /* 描画関数 / Draw lines */
        function drawLines(mainStrokeWidth, numLines) {

            var scaleValue = parseScaleValue();
            if (!isValidScaleValue(scaleValue)) return;
            var scaleFactor = scaleValue / 100;
            var preset = getSelectedPreset();
            var margin = parseMarginPt();
            if (!isValidMarginPt(margin)) return;

            if (preset.raindrop) {
                for (var r = 0; r < numLines; r++) {
                    var dropX = (left - margin) + Math.random() * (width + margin * 2);
                    var dropY = (bottom - margin) + Math.random() * (height + margin * 2);
                    var baseDropWidth = Math.max((mainStrokeWidth * scaleFactor) * 1.1, 2);
                    var dropWidth = Math.max((baseDropWidth * 0.6) + Math.random() * (baseDropWidth * 0.8), 2);
                    var dropHeightRatio = 1.5;
                    var dropOpacity = 35 + Math.random() * 65;
                    var dropGroup = layer.groupItems.add();
                    var raindropStyle = getRaindropAppearanceOptions();

                    if (rbRaindropB.value) {
                        dropHeightRatio = 1.8;
                    } else if (rbRaindropC.value) {
                        dropHeightRatio = 1.3;
                    }

                    var dropHeight = Math.max(dropWidth * dropHeightRatio, 3);

                    createRaindropShape(
                        dropGroup,
                        dropX,
                        dropY,
                        dropWidth,
                        dropHeight,
                        createRaindropColor(),
                        parseFloat(inputAngle.text) || 0,
                        raindropStyle
                    );
                    dropGroup.opacity = dropOpacity;
                    previewGroups.push(dropGroup);

                }
                app.redraw();
                return;
            }

            if (preset.polkadot) {
                for (var p = 0; p < numLines; p++) {
                    var dotX = (left - margin) + Math.random() * (width + margin * 2);
                    var dotY = (bottom - margin) + Math.random() * (height + margin * 2);
                    var dotBase = mainStrokeWidth * scaleFactor;
                    var dotDiameter = Math.max((dotBase * 0.5) + Math.random() * (dotBase * 1.5), 2);
                    var dotGroup = layer.groupItems.add();
                    createPolkaDot(dotGroup, dotX, dotY, dotDiameter, pickColor(preset.colors));
                    previewGroups.push(dotGroup);
                }
                app.redraw();
                return;
            }

            var angleDeg = parseFloat(inputAngle.text);
            if (isNaN(angleDeg)) angleDeg = parseFloat(preset.defaultAngle || preset.angleDeg || 45);
            var angleRad = angleDeg * Math.PI / 180;
            var dx = preset.lineLength * scaleFactor * Math.cos(angleRad);
            var dy = preset.lineLength * scaleFactor * Math.sin(angleRad);
            var childOffset = unitValueToPt(parseFloat(inputSpacing.text) || 0, "strokeUnits") * scaleFactor;
            var childOffsetX = 0;
            var childOffsetY = 0;
            if (childOffset !== 0) {
                var lineLength = Math.sqrt(dx * dx + dy * dy);
                if (lineLength > 0) {
                    childOffsetX = -dy / lineLength * childOffset;
                    childOffsetY = dx / lineLength * childOffset;
                }
            }
            var placedCenters = [];
            var placementRadius = Math.max(preset.lineLength * scaleFactor * (preset.placementGap || 1), mainStrokeWidth * scaleFactor * 4);
            var maxPlacementTries = 200;

            for (var i = 0; i < numLines; i++) {
                var startX, startY, endX, endY;
                var currentStrokeWidth = mainStrokeWidth;
                var scaledStrokeWidth = mainStrokeWidth * scaleFactor;

                if (preset.penetrate) {
                    var randAngle = (75 + Math.random() * 30) * Math.PI / 180;
                    var penetrateLen = (height + margin * 2) / Math.sin(randAngle);
                    startX = (left - margin) + Math.random() * (width + margin * 2);
                    startY = top + margin;
                    endX = startX + penetrateLen * Math.cos(randAngle);
                    endY = startY - penetrateLen * Math.sin(randAngle);
                    currentStrokeWidth = 0.5 + Math.random() * 4.5;
                } else {
                    var placed = false;
                    var tryCount;
                    for (tryCount = 0; tryCount < maxPlacementTries; tryCount++) {
                        startX = (left - margin) + Math.random() * (width + margin * 2);
                        startY = (bottom - margin) + Math.random() * (height + margin * 2);
                        endX = startX + dx;
                        endY = startY + dy;

                        if (!preset.avoidOverlap) {
                            placed = true;
                            break;
                        }

                        var centerX = (startX + endX) / 2;
                        var centerY = (startY + endY) / 2;
                        var overlaps = false;
                        for (var j = 0; j < placedCenters.length; j++) {
                            var pc = placedCenters[j];
                            var ddx = centerX - pc.x;
                            var ddy = centerY - pc.y;
                            if (Math.sqrt(ddx * ddx + ddy * ddy) < placementRadius) {
                                overlaps = true;
                                break;
                            }
                        }
                        if (!overlaps) {
                            placed = true;
                            placedCenters.push({ x: centerX, y: centerY });
                            break;
                        }
                    }

                    if (!placed) {
                        startX = (left - margin) + Math.random() * (width + margin * 2);
                        startY = (bottom - margin) + Math.random() * (height + margin * 2);
                        endX = startX + dx;
                        endY = startY + dy;
                        if (preset.avoidOverlap) {
                            placedCenters.push({ x: (startX + endX) / 2, y: (startY + endY) / 2 });
                        }
                    }
                }

                var lineColor = preset.colors ? pickColor(preset.colors) : strokeColor;

                var group = layer.groupItems.add();

                /* 子罫線（点線）— noChild の場合はスキップ / Child line (dashed) — skip when noChild is enabled */
                if (!preset.noChild) {
                    var childLine = group.pathItems.add();
                    /* 親線に対する法線方向へ平行移動 / Offset parallel to the parent line along its normal */
                    childLine.setEntirePath([[startX + childOffsetX, startY + childOffsetY], [endX + childOffsetX, endY + childOffsetY]]);
                    childLine.stroked = true;
                    childLine.strokeWidth = scaledStrokeWidth;
                    childLine.filled = false;
                    childLine.strokeColor = lineColor;
                    childLine.strokeCap = StrokeCap.ROUNDENDCAP;
                    childLine.strokeDashes = [0, childOffset > 0 ? childOffset : 3];
                }

                /* 親罫線（実線） / Parent line (solid) */
                var mainLine = group.pathItems.add();
                mainLine.setEntirePath([[startX, startY], [endX, endY]]);
                mainLine.stroked = true;
                mainLine.strokeWidth = preset.penetrate ? currentStrokeWidth : scaledStrokeWidth;
                mainLine.filled = false;
                mainLine.strokeColor = lineColor;
                mainLine.strokeDashes = [];
                mainLine.strokeCap = preset.roundCap ? StrokeCap.ROUNDENDCAP : StrokeCap.BUTTENDCAP;

                previewGroups.push(group);
            }
            app.redraw();
        }

        /* プレビュー削除関数 / Remove preview */
        function removePreview() {
            for (var i = previewGroups.length - 1; i >= 0; i--) {
                try {
                    previewGroups[i].remove();
                } catch (e) {
                    $.writeln("[VTRainMaker] Failed to remove preview group: " + e);
                }
            }
            previewGroups = [];
            app.redraw();
        }

        /* プレビュー更新 / Update preview */
        function updatePreview() {
            removePreview();
            if (!chkPreview.value) return;

            var preset = getSelectedPreset();
            var sw = unitValueToPt(parseFloat(inputStrokeWidth.text), "strokeUnits");
            var angle = parseFloat(inputAngle.text);
            var nl = parseInt(inputDensity.text, 10);
            var spacing = unitValueToPt(parseFloat(inputSpacing.text), "strokeUnits");
            var margin = parseMarginPt();
            var scale = parseScaleValue();
            if ((!preset.penetrate && (isNaN(sw) || sw <= 0)) || (!preset.penetrate && !preset.polkadot && isNaN(angle)) || isNaN(nl) || nl <= 0 || !isValidScaleValue(scale) || !isValidMarginPt(margin) || (!preset.noChild && !preset.raindrop && !preset.polkadot && (isNaN(spacing) || spacing < 0))) return;
            drawLines(sw, nl);
        }

        /* ラジオボタン切り替え時：デフォルト値を反映 / Apply preset defaults when the radio button changes */
        function onPresetChange() {
            var preset = getSelectedPreset();
            inputStrokeWidth.text = ptToUnitValue(parseFloat(preset.defaultStrokeWidth), "strokeUnits", 2);
            inputAngle.text = preset.defaultAngle || String(preset.angleDeg || 45);
            inputDensity.text = preset.defaultDensity;
            inputSpacing.text = ptToUnitValue(parseFloat(preset.defaultSpacing || "0"), "strokeUnits", 2);
            inputScale.text = preset.defaultScale || "100";
            syncScaleFromInput();
            normalizeRaindropAppearance();
            reflectEnabledUI();
            updatePreview();
        }

        /* UIの有効状態を反映 / Reflect enabled UI state */
        function reflectEnabledUI() {
            var preset = getSelectedPreset();
            var strokeEnabled = !preset.penetrate;
            var angleEnabled = !preset.penetrate && !preset.polkadot;
            var spacingEnabled = !preset.noChild && !preset.raindrop && !preset.polkadot;
            /* 入力欄・項目名・∧∨をまとめて切り替える / toggles the field, its label and the stepper together */
            setSteppedFieldEnabled(inputStrokeWidth, strokeEnabled);
            txtStrokeUnit.enabled = strokeEnabled;
            setSteppedFieldEnabled(inputAngle, angleEnabled);
            txtAngleUnit.enabled = angleEnabled;
            setSteppedFieldEnabled(inputSpacing, spacingEnabled);
            txtSpacingUnit.enabled = spacingEnabled;
            var raindropEnabled = !!preset.raindrop;
            raindropPanel.enabled = raindropEnabled;
            chkRaindropFill.enabled = raindropEnabled;
            chkRaindropStroke.enabled = raindropEnabled;
            rbRaindropA.enabled = raindropEnabled;
            rbRaindropB.enabled = raindropEnabled;
            rbRaindropC.enabled = raindropEnabled;
        }

        function normalizeRaindropAppearance(changedControl) {
            if (chkRaindropFill.value || chkRaindropStroke.value) return;

            if (changedControl === chkRaindropFill) {
                chkRaindropStroke.value = true;
            } else if (changedControl === chkRaindropStroke) {
                chkRaindropFill.value = true;
            } else {
                chkRaindropFill.value = true;
            }
        }

        function getRaindropAppearanceOptions() {
            var fillOn = chkRaindropFill.value;
            var strokeOn = chkRaindropStroke.value;

            if (fillOn && strokeOn) {
                if (Math.random() < 0.5) {
                    fillOn = true;
                    strokeOn = false;
                } else {
                    fillOn = false;
                    strokeOn = true;
                }
            }

            return {
                filled: fillOn,
                stroked: strokeOn,
                strokeColor: createRaindropStrokeColor(),
                strokeWidth: 2
            };
        }

        function syncScaleFromInput() {
            var v = parseScaleValue();
            if (isNaN(v)) return;
            v = clampScaleValue(v);
            if (Math.round(v) !== v) {
                inputScale.text = String(Math.round(v));
                v = Math.round(v);
            } else {
                inputScale.text = String(v);
            }
            sldScale.value = v;
        }

        function syncScaleFromSlider() {
            inputScale.text = String(Math.round(sldScale.value));
        }

        /* ダイアログボックス / Dialog box */
        var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
        setupWindow(dialog);

        var mainColumns = dialog.add("group");
        mainColumns.orientation = "row";
        mainColumns.alignChildren = ["fill", "top"];
        mainColumns.alignment = ["fill", "top"];
        mainColumns.spacing = COLUMN_SPACING;

        var leftCol = mainColumns.add("group");
        leftCol.orientation = "column";
        leftCol.alignChildren = ["fill", "top"];
        leftCol.alignment = ["left", "fill"];

        var rightCol = mainColumns.add("group");
        rightCol.orientation = "column";
        rightCol.alignChildren = ["left", "top"];
        rightCol.alignment = ["fill", "fill"];

        /* 雨の種類ラジオボタン / Rain type radio buttons */
        var radioPanel = leftCol.add("panel", undefined, getLabel("panelRainType"));
        setupPanel(radioPanel, 6);
        var rbHarusame = radioPanel.add("radiobutton", undefined, getLabel("rainHarusame"));
        rbHarusame.helpTip = getLabel("tipHarusame");
        var rbSamidare = radioPanel.add("radiobutton", undefined, getLabel("rainSamidare"));
        rbSamidare.helpTip = getLabel("tipSamidare");
        var rbTeppouame = radioPanel.add("radiobutton", undefined, getLabel("rainTeppouame"));
        rbTeppouame.helpTip = getLabel("tipTeppouame");
        var rbAmatsubu = radioPanel.add("radiobutton", undefined, getLabel("rainAmatsubu"));
        rbAmatsubu.helpTip = getLabel("tipAmatsubu");
        var rbMizutama = radioPanel.add("radiobutton", undefined, getLabel("rainMizutama"));
        rbMizutama.helpTip = getLabel("tipMizutama");
        rbHarusame.value = true;

        var optionPanel = rightCol.add("panel", undefined, getLabel("panelOptions"));
        setupPanel(optionPanel);
        optionPanel.alignChildren = ["left", "top"];
        optionPanel.alignment = ["fill", "top"];

        var labelWidth = 96;

        /**
         * 行に「∧∨＋入力欄」を追加する。∧∨・↑↓キーで増減したら onChanging を呼んでプレビューなどを更新する
         * @param {Group} rowGroup - 追加先の行
         * @param {StaticText} fieldLabel - 項目名（有効／無効をまとめて切り替えるため）
         * @param {string} initialText - 初期値
         * @param {Object} stepOptions - min / max / integer（addStepper() に渡す）
         * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
         */
        function addSteppedInput(rowGroup, fieldLabel, initialText, stepOptions) {
            /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
            var stepperInputGroup = rowGroup.add("group");
            stepperInputGroup.orientation = "row";
            stepperInputGroup.alignChildren = ["left", "center"];
            stepperInputGroup.spacing = 0;
            stepperInputGroup.margins = 0;

            var numberInput;
            stepOptions.onStep = function (steppedInput) {
                if (typeof steppedInput.onChanging === "function") steppedInput.onChanging();
            };
            var stepperGroup = addStepper(stepperInputGroup, function () { return numberInput; }, stepOptions);
            numberInput = stepperInputGroup.add("edittext", undefined, initialText);
            numberInput.fieldLabel = fieldLabel;
            numberInput.stepperGroup = stepperGroup;
            bindSteppedArrowKeys(numberInput, stepperGroup);

            /* 項目名のクリックで入力欄にフォーカスを移す / clicking the label focuses the field */
            fieldLabel.addEventListener("click", function () {
                numberInput.active = false; /* 一度外さないとフォーカスが移らないことがある / reset first or focus may not move */
                numberInput.active = true;
            });
            return numberInput;
        }

        var grpScale = optionPanel.add("group");
        grpScale.alignment = "left";
        var lblScale = grpScale.add("statictext", undefined, labelText("labelScale"), { justify: "right" });
        lblScale.preferredSize = [labelWidth, -1];
        lblScale.justify = "right";
        var inputScale = addSteppedInput(grpScale, lblScale, presets.harusame.defaultScale || "100", { min: 25, max: 300, integer: true });
        inputScale.helpTip = getLabel("tipScale");
        inputScale.characters = 4;
        var txtScaleUnit = grpScale.add("statictext", undefined, getLabel("unitPercent"));

        var grpScaleSlider = optionPanel.add("group");
        grpScaleSlider.alignment = ["fill", "top"];
        var sldScale = grpScaleSlider.add("slider", undefined, 100, 25, 300);
        sldScale.helpTip = getLabel("tipScale");
        sldScale.preferredSize.width = 180;

        var grpStroke = optionPanel.add("group");
        grpStroke.alignment = "left";
        var lblStroke = grpStroke.add("statictext", undefined, labelText("labelStrokeWidth"), { justify: "right" });
        lblStroke.preferredSize = [labelWidth, -1];
        lblStroke.justify = "right";
        var inputStrokeWidth = addSteppedInput(grpStroke, lblStroke, ptToUnitValue(parseFloat(presets.harusame.defaultStrokeWidth), "strokeUnits", 2), { min: 0 });
        inputStrokeWidth.helpTip = getLabel("tipStrokeWidth");
        inputStrokeWidth.characters = 6;
        var txtStrokeUnit = grpStroke.add("statictext", undefined, strokeUnitInfo.label);
        var grpAngle = optionPanel.add("group");
        grpAngle.alignment = "left";
        var lblAngle = grpAngle.add("statictext", undefined, labelText("labelAngle"), { justify: "right" });
        lblAngle.preferredSize = [labelWidth, -1];
        lblAngle.justify = "right";
        var inputAngle = addSteppedInput(grpAngle, lblAngle, presets.harusame.defaultAngle || String(presets.harusame.angleDeg || 45), { min: 0 });
        inputAngle.helpTip = getLabel("tipAngle");
        inputAngle.characters = 6;
        var txtAngleUnit = grpAngle.add("statictext", undefined, getLabel("unitDegree"));

        var grpDensity = optionPanel.add("group");
        grpDensity.alignment = "left";
        var lblDensity = grpDensity.add("statictext", undefined, labelText("labelDensity"), { justify: "right" });
        lblDensity.preferredSize = [labelWidth, -1];
        lblDensity.justify = "right";
        var inputDensity = addSteppedInput(grpDensity, lblDensity, presets.harusame.defaultDensity, { min: 1, integer: true });
        inputDensity.helpTip = getLabel("tipDensity");
        inputDensity.characters = 6;
        grpDensity.add("statictext", undefined, getLabel("unitLines"));

        var grpSpacing = optionPanel.add("group");
        grpSpacing.alignment = "left";
        var lblSpacing = grpSpacing.add("statictext", undefined, labelText("labelSpacing"), { justify: "right" });
        lblSpacing.preferredSize = [labelWidth, -1];
        lblSpacing.justify = "right";
        var inputSpacing = addSteppedInput(grpSpacing, lblSpacing, ptToUnitValue(parseFloat(presets.harusame.defaultSpacing || "0"), "strokeUnits", 2), { min: 0 });
        inputSpacing.helpTip = getLabel("tipSpacing");
        inputSpacing.characters = 6;
        var txtSpacingUnit = grpSpacing.add("statictext", undefined, strokeUnitInfo.label);

        var grpMargin = optionPanel.add("group");
        grpMargin.alignment = "left";
        var lblMargin = grpMargin.add("statictext", undefined, labelText("labelMargin"), { justify: "right" });
        lblMargin.preferredSize = [labelWidth, -1];
        lblMargin.justify = "right";
        var inputMargin = addSteppedInput(grpMargin, lblMargin, ptToUnitValue(20, "rulerType", 2), { min: 0 });
        inputMargin.helpTip = getLabel("tipMargin");
        inputMargin.characters = 6;
        grpMargin.add("statictext", undefined, rulerUnitInfo.label);

        var raindropPanel = optionPanel.add("panel", undefined, getLabel("panelRaindrop"));
        setupPanel(raindropPanel);
        raindropPanel.alignChildren = ["left", "top"];
        raindropPanel.alignment = ["fill", "top"];

        var grpRaindropAppearance = raindropPanel.add("group");
        grpRaindropAppearance.orientation = "row";
        grpRaindropAppearance.alignChildren = ["left", "center"];

        var chkRaindropFill = grpRaindropAppearance.add("checkbox", undefined, getLabel("chkRaindropFill"));
        chkRaindropFill.helpTip = getLabel("tipRaindropFill");
        chkRaindropFill.value = true;

        var chkRaindropStroke = grpRaindropAppearance.add("checkbox", undefined, getLabel("chkRaindropStroke"));
        chkRaindropStroke.helpTip = getLabel("tipRaindropStroke");
        chkRaindropStroke.value = false;

        var grpRaindropShapes = raindropPanel.add("group");
        grpRaindropShapes.orientation = "row";
        grpRaindropShapes.alignChildren = ["left", "center"];

        var rbRaindropA = grpRaindropShapes.add("radiobutton", undefined, getLabel("raindropShapeA"));
        rbRaindropA.helpTip = getLabel("tipRaindropShape");
        var rbRaindropB = grpRaindropShapes.add("radiobutton", undefined, getLabel("raindropShapeB"));
        rbRaindropB.helpTip = getLabel("tipRaindropShape");
        var rbRaindropC = grpRaindropShapes.add("radiobutton", undefined, getLabel("raindropShapeC"));
        rbRaindropC.helpTip = getLabel("tipRaindropShape");
        rbRaindropB.value = true;

        /* 2カラムにしないUI / UI outside the two-column layout */
        var buttonRow = addButtonRow(dialog);

        var chkPreview = buttonRow.leftGroup.add("checkbox", undefined, getLabel("preview"));
        chkPreview.helpTip = getLabel("tipPreview");
        chkPreview.value = false;

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("ok"), { name: "ok" });

        /* イベントハンドラ / Event handlers */
        rbHarusame.onClick = onPresetChange;
        rbSamidare.onClick = onPresetChange;
        rbTeppouame.onClick = onPresetChange;
        rbAmatsubu.onClick = onPresetChange;
        rbMizutama.onClick = onPresetChange;
        chkPreview.onClick = updatePreview;
        inputStrokeWidth.onChanging = updatePreview;
        inputAngle.onChanging = updatePreview;
        inputDensity.onChanging = updatePreview;
        inputSpacing.onChanging = updatePreview;
        inputMargin.onChanging = updatePreview;
        chkRaindropFill.onClick = function () {
            normalizeRaindropAppearance(chkRaindropFill);
            updatePreview();
        };
        chkRaindropStroke.onClick = function () {
            normalizeRaindropAppearance(chkRaindropStroke);
            updatePreview();
        };
        rbRaindropA.onClick = updatePreview;
        rbRaindropB.onClick = updatePreview;
        rbRaindropC.onClick = updatePreview;
        reflectEnabledUI();

        inputScale.onChanging = function () {
            syncScaleFromInput();
            updatePreview();
        };

        sldScale.onChanging = function () {
            syncScaleFromSlider();
            updatePreview();
        };

        alignRightOnlyButtonRow(buttonRow);
        prepareDialogWindow(dialog, SCRIPT_NAME);
        var result = dialog.show();

        if (result !== 1) {
            removePreview();
            return;
        }

        var preset = getSelectedPreset();
        var mainStrokeWidth = preset.penetrate ? 1 : unitValueToPt(parseFloat(inputStrokeWidth.text), "strokeUnits");
        var angle = parseFloat(inputAngle.text);
        var numLines = parseInt(inputDensity.text, 10);
        var spacing = unitValueToPt(parseFloat(inputSpacing.text), "strokeUnits");
        var margin = parseMarginPt();
        var scale = parseScaleValue();
        normalizeRaindropAppearance();

        if (!preset.penetrate && (isNaN(mainStrokeWidth) || mainStrokeWidth <= 0)) {
            removePreview();
            alert(getLabel("invalidStrokeWidth"));
            return;
        }
        if (!preset.penetrate && !preset.polkadot && isNaN(angle)) {
            removePreview();
            alert(getLabel("invalidAngle"));
            return;
        }
        if (isNaN(numLines) || numLines <= 0) {
            removePreview();
            alert(getLabel("invalidDensity"));
            return;
        }
        if (!isValidScaleValue(scale)) {
            removePreview();
            alert(getLabel("invalidScale"));
            return;
        }
        if (!isValidMarginPt(margin)) {
            removePreview();
            alert(getLabel("invalidMargin"));
            return;
        }
        if (!preset.noChild && !preset.raindrop && !preset.polkadot && (isNaN(spacing) || spacing < 0)) {
            removePreview();
            alert(getLabel("invalidSpacing"));
            return;
        }

        /* OK → プレビューがあればそのまま確定、なければ描画 / On OK, keep the preview as final if it exists; otherwise draw */
        if (previewGroups.length === 0) {
            drawLines(mainStrokeWidth, numLines);
        }
    }

    drawRandomParentChildLines();

})();
