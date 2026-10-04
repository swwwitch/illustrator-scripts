#target illustrator
#targetengine "SwwwitchPalettes"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

キー入力・整列・変形・ペースト・描画まわりの、よく使う環境設定を常駐パレットで切り替えます。
操作したその場で反映されるため、環境設定ダイアログを開く手間がありません。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiQuickPrefsPalette.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n41d8dc1961be

### Overview

A persistent palette for the preferences you change most often: keyboard increment, align, transform, paste and drawing.
Every change applies the moment you make it, so there is no need to open the Preferences dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiQuickPrefsPalette.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiQuickPrefsPalette";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.4.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiQuickPrefsPalette.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiQuickPrefsPalette.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n41d8dc1961be"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var PALETTE_OPACITY = 0.98;   /* パレットの不透明度 / Palette opacity */

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

    /* 日英ラベル定義（カテゴリ分け）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "クイック環境設定", en: "Quick Preferences" }
        },
        panel: {
            keyInput: { ja: "キー入力", en: "Keyboard Increment" },
            align: { ja: "整列オプション", en: "Align Options" },
            transform: { ja: "変形オプション", en: "Transform Options" },
            copyPaste: { ja: "コピー/ペースト", en: "Copy / Paste" },
            drawing: { ja: "描画", en: "Drawing" }
        },
        checkbox: {
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            previewBounds: { ja: "プレビュー境界を使用", en: "Use Preview Bounds" },
            transformPattern: { ja: "パターン", en: "Pattern Tiles" },
            scaleCorners: { ja: "角", en: "Corners" },
            scaleStroke: { ja: "線幅と効果", en: "Strokes & Effects" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集", en: "Real-time Drawing & Editing" },
            pastePlain: { ja: "書式なしペースト", en: "Paste without Formatting" },
            pastePreserve: { ja: "コピー元のレイヤーにペースト", en: "Paste Remembers Layers" }
        },
        tooltip: {
            cursorStep: {
                ja: "矢印キーでの移動量（環境設定 > 一般 > キー入力）。∧∨・↑↓で次の整数へ、shift＋で10の倍数へ、option＋で0.1ずつ",
                en: "Keyboard increment (Preferences > General). Steppers and Up/Down go to the next whole number; Shift to the next multiple of 10; Option by 0.1"
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" },
            unit: { ja: "定規の単位を切り替え", en: "Switch the ruler unit" },
            previewBounds: { ja: "整列・分布でプレビュー境界（線幅・効果を含む）を使用", en: "Use preview bounds (incl. stroke & effects) for align/distribute" },
            glyphBounds: { ja: "ポイント文字・エリア内文字を字形の境界で整列", en: "Align point & area type to glyph bounds" },
            transformPattern: { ja: "変形時にパターンも変形する", en: "Transform pattern tiles when transforming" },
            scaleCorners: { ja: "拡大・縮小時に角（ライブコーナー）の半径も拡大・縮小", en: "Scale corner radius when scaling" },
            scaleStroke: { ja: "拡大・縮小時に線幅と効果も拡大・縮小", en: "Scale strokes & effects when scaling" },
            pastePlain: { ja: "書式を保持せずにペースト", en: "Paste without keeping formatting" },
            pastePreserve: { ja: "コピー元と同じレイヤーにペースト", en: "Paste into the original (source) layer" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集を切り替え", en: "Toggle real-time drawing & editing" },
            refreshGpuPreview: { ja: "GPUプレビューを更新（再描画）", en: "Refresh the GPU preview (redraw)" }
        },
        button: {
            refreshGpuPreview: { ja: "プレビュー更新", en: "Refresh Preview" }
        }
    };

    // =========================================
    // 単位 / Unit
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       popup：単位ポップアップに並べるかどうか（このスクリプト独自の項目）
       Unit code -> display label and points per unit.
       popup: whether the unit appears in the unit popup (specific to this script). */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72,               popup: true },   /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4,        popup: true },   /* 1 */
        { label: "pt",    pointsPerUnit: 1,                popup: true },   /* 2 */
        { label: "pica",  pointsPerUnit: 12,               popup: true },   /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54,        popup: true },   /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25, popup: true },   /* 5 */
        { label: "px",    pointsPerUnit: 1,                popup: true },   /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12,          popup: false },  /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000, popup: false },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36,          popup: false },  /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12,          popup: false }   /* 10 */
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

    /* ポップアップに表示する単位コード（表示順、UNITS から派生）/ Unit codes shown in the popup (in order, derived from UNITS) */
    var UNIT_POPUP_CODES = (function () {
        var codes = [];
        for (var i = 0; i < UNITS.length; i++) {
            if (UNITS[i].popup) codes.push(i);
        }
        return codes;
    })();

    /**
     * 単位コードを単位ポップアップの位置に変換する
     * @param {number} unitCode - 単位コード
     * @returns {number} ポップアップの位置（並べていない単位は -1）
     */
    function unitCodeToPopupIndex(unitCode) {
        for (var i = 0; i < UNIT_POPUP_CODES.length; i++) {
            if (UNIT_POPUP_CODES[i] === unitCode) return i;
        }
        return -1;
    }

    // =========================================
    // 状態（キャッシュ）/ State (cache)
    // =========================================

    /* 常駐エンジンでは app.preferences への都度アクセスを避け、読み出した値をここへ保持 */
    /* In a persistent engine we avoid per-event app.preferences access; fetched values are cached here */
    var PREF_STATE = {
        rulerType: 2,
        cursorKeyLengthPt: 1.0
    };

    /**
     * 控えてある定規単位の1単位あたりのポイント数を返す（未知のコードは pt）
     * @returns {number} 1単位あたりのポイント数
     */
    function getCurrentPtPerUnit() {
        return (UNITS[PREF_STATE.rulerType] || UNITS[2]).pointsPerUnit;
    }

    // =========================================
    // 環境設定の読み出し（同期・直接）/ Reading preferences (direct & synchronous)
    // 読み出しはエンジンを跨いでも安全なため、パレットエンジンで直接取得する
    // Reads are safe across engines, so fetch them directly in the palette engine
    // =========================================

    /**
     * パレットに表示する環境設定をまとめて読み出す
     * @returns {Object} 環境設定キー（一部は短い名前）と値の組
     */
    function readAllPreferences() {
        var appPreferences = app.preferences;
        return {
            EnableActualPointTextSpaceAlign: appPreferences.getBooleanPreference("EnableActualPointTextSpaceAlign"),
            EnableActualAreaTextSpaceAlign: appPreferences.getBooleanPreference("EnableActualAreaTextSpaceAlign"),
            includeStrokeInBounds: appPreferences.getBooleanPreference("includeStrokeInBounds"),
            transformPatterns: appPreferences.getBooleanPreference("transformPatterns"),
            scaleLineWeight: appPreferences.getBooleanPreference("scaleLineWeight"),
            LiveEdit_State_Machine: appPreferences.getBooleanPreference("LiveEdit_State_Machine"),
            pastePlain: appPreferences.getBooleanPreference("plugin/FileClipboard/pasteWithoutFormatting"),
            pastePreserve: appPreferences.getBooleanPreference("layers/pastePreserve"),
            policyForPreservingCorners: appPreferences.getIntegerPreference("policyForPreservingCorners"),
            rulerType: getUnitInfo().code,
            cursorKeyLength: appPreferences.getRealPreference("cursorKeyLength")
        };
    }

    // =========================================
    // BridgeTalk 委譲（書き込み）/ BridgeTalk delegation (writes)
    // =========================================

    /**
     * メインエンジン（target="illustrator"）へコードを送って実行する
     * @param {string} bodyCode - 実行するコード
     * @returns {void}
     */
    function runInMainEngine(bodyCode) {
        try {
            var bridge = new BridgeTalk();
            bridge.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
            bridge.body = bodyCode;
            bridge.onError = function (message) {
                /* エラーは意図的に握りつぶす（常駐パレットなので alert は出さない）。
                   完全に無音だと失敗に気づけないため、デバッグ用に $.writeln にだけ残す。
                   Intentionally swallowed (no alert in a persistent palette);
                   logged to $.writeln only so failures stay noticeable while debugging. */
                $.writeln("AiQuickPrefsPalette BridgeTalk error: " + message.body);
            };
            bridge.send();
        } catch (e) {
            /* BridgeTalk 不可時は同一エンジンで直接実行（失敗は無視）/ Fallback: run directly in this engine (failures ignored) */
            try {
                eval(bodyCode);
            } catch (e2) {}
        }
    }

    /**
     * 環境設定をメインエンジンで書き込む（値の書式は型別の関数がそろえる）
     * @param {string} setterName - "setBooleanPreference" などのメソッド名
     * @param {string} prefKey - 環境設定キー
     * @param {string|number} valueLiteral - コードに埋め込む値
     * @returns {void}
     */
    function btSetPreference(setterName, prefKey, valueLiteral) {
        runInMainEngine('app.preferences.' + setterName + '("' + prefKey + '", ' + valueLiteral + ');');
    }

    /**
     * 真偽値の環境設定を書き込む
     * @param {string} prefKey - 環境設定キー
     * @param {boolean} value - 値
     * @returns {void}
     */
    function btSetBooleanPreference(prefKey, value) { btSetPreference('setBooleanPreference', prefKey, value ? 'true' : 'false'); }

    /**
     * 整数の環境設定を書き込む
     * @param {string} prefKey - 環境設定キー
     * @param {number} value - 値
     * @returns {void}
     */
    function btSetIntegerPreference(prefKey, value) { btSetPreference('setIntegerPreference', prefKey, parseInt(value, 10)); }

    /**
     * 実数の環境設定を書き込む
     * @param {string} prefKey - 環境設定キー
     * @param {number} value - 値
     * @returns {void}
     */
    function btSetRealPreference(prefKey, value) { btSetPreference('setRealPreference', prefKey, Number(value)); }

    // =========================================
    // カーソル移動量 / Cursor step
    // =========================================

    /**
     * 今の定規単位の値を cursorKeyLength（pt）として保存する
     * @param {number} unitValue - 今の定規単位での値
     * @returns {boolean} 保存したら true（数値でない・負の値なら false）
     */
    function saveCursorKeyLengthInCurrentUnit(unitValue) {
        if (isNaN(unitValue) || unitValue < 0) return false;
        PREF_STATE.cursorKeyLengthPt = unitValue * getCurrentPtPerUnit();
        btSetRealPreference("cursorKeyLength", PREF_STATE.cursorKeyLengthPt);
        return true;
    }

    /**
     * 控えてある cursorKeyLength（pt）を今の定規単位の文字列（小数第1位まで）で返す
     * @returns {string} 表示用の値
     */
    function readCursorKeyLengthInCurrentUnit() {
        return (PREF_STATE.cursorKeyLengthPt / getCurrentPtPerUnit()).toFixed(1);
    }

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
    // UI ヘルパー / UI helpers
    // =========================================

    /**
     * ツールチップ付きのチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.checkbox のキー
     * @param {string} tooltipKey - LABELS.tooltip のキー
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addCheckbox(parentPanel, labelKey, tooltipKey) {
        var newCheckbox = parentPanel.add('checkbox', undefined, getLabel('checkbox.' + labelKey));
        newCheckbox.helpTip = getLabel('tooltip.' + tooltipKey);
        return newCheckbox;
    }

    /**
     * 真偽値の環境設定とつないだチェックボックスを追加する（クリックでその場で書き込む）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelKey - LABELS.checkbox と LABELS.tooltip のキー
     * @param {string} prefKey - 環境設定キー
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addBooleanPrefCheckbox(parentPanel, labelKey, prefKey) {
        var prefCheckbox = addCheckbox(parentPanel, labelKey, labelKey);
        prefCheckbox.onClick = function () {
            btSetBooleanPreference(prefKey, prefCheckbox.value === true);
        };
        return prefCheckbox;
    }

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

    // =========================================
    // パネル構築 / Panel builders
    // =========================================

    /**
     * キー入力パネル（移動量＋単位ポップアップ）を作る。編集中フラグと単位の同期はこの中に閉じ込める
     * @param {Group} parent - 追加先
     * @returns {{field: EditText, isEditing: Function, syncDisplay: Function}} 入力欄・編集中かどうか・表示の更新
     */
    function buildKeyInputPanel(parent) {
        var keyInputPanel = parent.add('panel', undefined, getLabel('panel.keyInput'));
        setupPanel(keyInputPanel, 6);
        keyInputPanel.orientation = "row"; /* 入力欄と単位を横に並べる / field and unit side by side */
        keyInputPanel.alignChildren = ["left", "center"];

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var cursorStepGroup = keyInputPanel.add('group');
        cursorStepGroup.orientation = 'row';
        cursorStepGroup.alignChildren = ['left', 'center'];
        cursorStepGroup.spacing = 0;
        cursorStepGroup.margins = 0;
        var cursorStepStepper = addStepper(cursorStepGroup, function () { return cursorStepField; }, {
            step: 1, min: 0,
            /* 常に小数第1位で表示し、その場で保存 / Always show 1 decimal and save right away */
            onStep: function (numberInput) {
                numberInput.text = parseFloat(numberInput.text).toFixed(1);
                saveCursorKeyLengthInCurrentUnit(parseFloat(numberInput.text));
            }
        });
        var cursorStepField = cursorStepGroup.add('edittext', undefined, "1.0");
        cursorStepField.characters = 4;
        cursorStepField.helpTip = getLabel('tooltip.cursorStep');
        bindSteppedArrowKeys(cursorStepField, cursorStepStepper);

        /* 編集中フラグ：入力欄にフォーカスがある間は外部同期で値を上書きしない / Editing flag: don't let external sync overwrite while the field has focus */
        var isEditingCursorStep = false;
        cursorStepField.onActivate = function () { isEditingCursorStep = true; };
        cursorStepField.onDeactivate = function () { isEditingCursorStep = false; };

        /* 確定時に保存（不正値は元へ戻す）/ Save on commit (restore on invalid input) */
        cursorStepField.onChange = function () {
            if (!saveCursorKeyLengthInCurrentUnit(parseFloat(cursorStepField.text))) {
                cursorStepField.text = readCursorKeyLengthInCurrentUnit();
            }
        };

        var unitDropdown = keyInputPanel.add('dropdownlist', undefined, []);
        for (var i = 0; i < UNIT_POPUP_CODES.length; i++) {
            /* 定規単位なので単位コード5は「H」/ ruler unit, so unit code 5 shows as H */
            var popupUnitCode = UNIT_POPUP_CODES[i];
            unitDropdown.add('item', (popupUnitCode === 5) ? "H" : UNITS[popupUnitCode].label);
        }
        unitDropdown.preferredSize.width = 55;
        unitDropdown.helpTip = getLabel('tooltip.unit');

        /* 単位ポップアップ：選んだ単位を定規単位(rulerType)へ反映し、表示を再計算 / Unit popup: apply the chosen unit to rulerType and recompute the display */
        var isSyncingUnit = false;
        unitDropdown.onChange = function () {
            if (isSyncingUnit || !unitDropdown.selection) return;
            var rulerUnitCode = UNIT_POPUP_CODES[unitDropdown.selection.index];
            PREF_STATE.rulerType = rulerUnitCode;
            btSetIntegerPreference("rulerType", rulerUnitCode);
            /* 保存済み pt 値は不変。新しい単位で再表示 / Stored pt value is unchanged; redisplay in the new unit */
            cursorStepField.text = readCursorKeyLengthInCurrentUnit();
        };

        /**
         * PREF_STATE から表示（単位ポップアップ＋数値）を更新する。onChange は発火させない
         * @returns {void}
         */
        function syncDisplay() {
            var unitPopupIndex = unitCodeToPopupIndex(PREF_STATE.rulerType);
            if (unitPopupIndex >= 0) {
                isSyncingUnit = true;
                unitDropdown.selection = unitPopupIndex;
                isSyncingUnit = false;
            }
            cursorStepField.text = readCursorKeyLengthInCurrentUnit();
        }

        return {
            field: cursorStepField,
            isEditing: function () { return isEditingCursorStep; },
            syncDisplay: syncDisplay
        };
    }

    /**
     * 整列オプションパネル（プレビュー境界を使用／字形の境界に整列）を作る
     * @param {Group} parent - 追加先
     * @returns {{previewBounds: Checkbox, glyphBounds: Checkbox}} チェックボックス
     */
    function buildAlignPanel(parent) {
        var alignPanel = parent.add('panel', undefined, getLabel('panel.align'));
        setupPanel(alignPanel, 6);

        var checkboxPreviewBounds = addBooleanPrefCheckbox(alignPanel, 'previewBounds', 'includeStrokeInBounds');

        /* 字形の境界に整列：ポイント文字・エリア内文字の両方をまとめてON/OFF / Align to glyph bounds: toggle both point & area type together */
        var checkboxGlyphBounds = addCheckbox(alignPanel, 'glyphBounds', 'glyphBounds');
        checkboxGlyphBounds.onClick = function () {
            var isOn = checkboxGlyphBounds.value === true;
            btSetBooleanPreference("EnableActualPointTextSpaceAlign", isOn);
            btSetBooleanPreference("EnableActualAreaTextSpaceAlign", isOn);
        };

        return { previewBounds: checkboxPreviewBounds, glyphBounds: checkboxGlyphBounds };
    }

    /**
     * 変形オプションパネル（パターン／角／線幅と効果）を作る。option＋クリックで3つをそろえて切り替える
     * @param {Group} parent - 追加先
     * @returns {{transformPattern: Checkbox, scaleCorners: Checkbox, scaleStroke: Checkbox}} チェックボックス
     */
    function buildTransformPanel(parent) {
        var transformPanel = parent.add('panel', undefined, getLabel('panel.transform'));
        setupPanel(transformPanel, 6);

        var checkboxTransformPattern = addCheckbox(transformPanel, 'transformPattern', 'transformPattern');
        /* 角は真偽値でなく整数（1=ON, 2=OFF）/ Corners is an integer pref (1=ON, 2=OFF) */
        var checkboxScaleCorners = addCheckbox(transformPanel, 'scaleCorners', 'scaleCorners');
        var checkboxScaleStroke = addCheckbox(transformPanel, 'scaleStroke', 'scaleStroke');

        /* チェックボックスと、その値を環境設定へ書き込む処理の組 / Each checkbox paired with the write for its preference */
        var transformOptions = [
            { checkbox: checkboxTransformPattern, apply: function () { btSetBooleanPreference("transformPatterns", checkboxTransformPattern.value === true); } },
            { checkbox: checkboxScaleCorners, apply: function () { btSetIntegerPreference("policyForPreservingCorners", checkboxScaleCorners.value ? 1 : 2); } },
            { checkbox: checkboxScaleStroke, apply: function () { btSetBooleanPreference("scaleLineWeight", checkboxScaleStroke.value === true); } }
        ];

        /**
         * クリックしたチェックボックスを書き込む。option を押していれば3つを同じ値にそろえて書き込む
         * @param {Object} clickedOption - transformOptions の要素
         * @returns {void}
         */
        function onTransformOptionClick(clickedOption) {
            if (!ScriptUI.environment.keyboardState.altKey) {
                clickedOption.apply();
                return;
            }
            var newValue = clickedOption.checkbox.value === true;
            for (var i = 0; i < transformOptions.length; i++) {
                transformOptions[i].checkbox.value = newValue;
                transformOptions[i].apply();
            }
        }

        for (var i = 0; i < transformOptions.length; i++) {
            (function (transformOption) {
                transformOption.checkbox.onClick = function () { onTransformOptionClick(transformOption); };
            })(transformOptions[i]);
        }

        return { transformPattern: checkboxTransformPattern, scaleCorners: checkboxScaleCorners, scaleStroke: checkboxScaleStroke };
    }

    /**
     * コピー/ペーストパネル（書式なしペースト／コピー元のレイヤーにペースト）を作る
     * @param {Group} parent - 追加先
     * @returns {{pastePlain: Checkbox, pastePreserve: Checkbox}} チェックボックス
     */
    function buildCopyPastePanel(parent) {
        var copyPastePanel = parent.add('panel', undefined, getLabel('panel.copyPaste'));
        setupPanel(copyPastePanel, 6);
        return {
            pastePlain: addBooleanPrefCheckbox(copyPastePanel, 'pastePlain', 'plugin/FileClipboard/pasteWithoutFormatting'),
            pastePreserve: addBooleanPrefCheckbox(copyPastePanel, 'pastePreserve', 'layers/pastePreserve')
        };
    }

    /**
     * 描画パネル（リアルタイムの描画と編集＋プレビュー更新ボタン）を作る
     * @param {Group} parent - 追加先
     * @returns {{realtimeDrawing: Checkbox, refreshButton: Button}} チェックボックスと、レイアウト確定後に高さを詰めるボタン
     */
    function buildDrawingPanel(parent) {
        var drawingPanel = parent.add('panel', undefined, getLabel('panel.drawing'));
        setupPanel(drawingPanel, 6);

        var checkboxRealtimeDrawing = addBooleanPrefCheckbox(drawingPanel, 'realtimeDrawing', 'LiveEdit_State_Machine');

        /* GPU プレビューを更新（View using GPU を2回トグルして再描画）/ Refresh GPU preview (toggle View using GPU twice to redraw) */
        var btnRefreshGpuPreview = drawingPanel.add('button', undefined, getLabel('button.refreshGpuPreview'));
        btnRefreshGpuPreview.alignment = ['left', 'top']; /* 幅いっぱいにしない（ラベル幅）/ Do not fill width (size to label) */
        btnRefreshGpuPreview.helpTip = getLabel('tooltip.refreshGpuPreview');
        btnRefreshGpuPreview.onClick = function () {
            runInMainEngine('try{app.executeMenuCommand("View using GPU");app.executeMenuCommand("View using GPU");}catch(e){}');
        };

        return { realtimeDrawing: checkboxRealtimeDrawing, refreshButton: btnRefreshGpuPreview };
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /**
     * パレットとパネルを組み立てる
     * @returns {Object} パレットと各パネルのコントロール
     */
    function buildPalette() {
        var palette = new Window('palette', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        setupWindow(palette);
        palette.opacity = PALETTE_OPACITY;

        var panelColumn = palette.add('group');
        panelColumn.orientation = 'column';
        panelColumn.alignChildren = ['fill', 'top'];

        return {
            palette: palette,
            keyInput: buildKeyInputPanel(panelColumn),
            align: buildAlignPanel(panelColumn),
            transform: buildTransformPanel(panelColumn),
            copyPaste: buildCopyPastePanel(panelColumn),
            drawing: buildDrawingPanel(panelColumn)
        };
    }

    /**
     * 読み出した環境設定をパレットの表示と PREF_STATE へ反映する
     * @param {Object} paletteControls - buildPalette() の戻り値
     * @param {Object} prefValues - readAllPreferences() の戻り値
     * @returns {void}
     */
    function applyPreferencesToControls(paletteControls, prefValues) {
        paletteControls.align.previewBounds.value = prefValues.includeStrokeInBounds === true;
        /* ポイント文字・エリア内文字が両方ONのときだけチェック / Checked only when both point & area type are on */
        paletteControls.align.glyphBounds.value = prefValues.EnableActualPointTextSpaceAlign === true && prefValues.EnableActualAreaTextSpaceAlign === true;
        paletteControls.transform.transformPattern.value = prefValues.transformPatterns === true;
        paletteControls.transform.scaleCorners.value = parseInt(prefValues.policyForPreservingCorners, 10) === 1;
        paletteControls.transform.scaleStroke.value = prefValues.scaleLineWeight === true;
        paletteControls.copyPaste.pastePlain.value = prefValues.pastePlain === true;
        paletteControls.copyPaste.pastePreserve.value = prefValues.pastePreserve === true;
        paletteControls.drawing.realtimeDrawing.value = prefValues.LiveEdit_State_Machine === true;

        var rulerUnitCode = parseInt(prefValues.rulerType, 10);
        if (!isNaN(rulerUnitCode)) PREF_STATE.rulerType = rulerUnitCode;
        var cursorKeyLengthPt = parseFloat(prefValues.cursorKeyLength);
        if (!isNaN(cursorKeyLengthPt)) PREF_STATE.cursorKeyLengthPt = cursorKeyLengthPt;

        /* 単位ポップアップと数値表示を PREF_STATE に同期 / Sync the unit popup and value display to PREF_STATE */
        paletteControls.keyInput.syncDisplay();
    }

    /**
     * パレットを表示する。すでに開いていれば前面に出す
     * @returns {void}
     */
    function main() {
        if ($.global.__aiQuickPrefsPalette) {
            /* 閉じたあとの参照が残っていると show() で例外になる / show() throws on a stale reference left after closing */
            try {
                $.global.__aiQuickPrefsPalette.show();
                return;
            } catch (e) {
                $.global.__aiQuickPrefsPalette = null;
            }
        }

        var paletteControls = buildPalette();
        var palette = paletteControls.palette;
        $.global.__aiQuickPrefsPalette = palette;

        /* 閉じたら参照をクリア（次回は再構築）/ Clear the reference on close (rebuild next time) */
        palette.onClose = function () {
            $.global.__aiQuickPrefsPalette = null;
        };

        /* パレットがアクティブなとき esc キーで閉じる / Close the palette with Esc while it is active */
        palette.addEventListener('keydown', function (event) {
            if (event.keyName === 'Escape') {
                palette.close();
            }
        });

        /* 構築直後に現在値を反映 / Populate from current values right after building */
        applyPreferencesToControls(paletteControls, readAllPreferences());

        /* 表示時：レイアウト再計算（描画欠け防止）＋フォーカス＋最新値の再読込 / On show: recalc layout (avoid partial rendering), focus, and reload current values */
        palette.onShow = function () {
            palette.layout.layout(true);
            palette.layout.resize();
            paletteControls.keyInput.field.active = true;
            applyPreferencesToControls(paletteControls, readAllPreferences());
        };

        /* 再アクティブ時：外部変更（環境設定ダイアログ等）へクリックで追従。ただし編集中は同期しない（入力値の上書き防止）/ On re-activate: follow external changes via click, but skip while editing (avoid clobbering input) */
        palette.onActivate = function () {
            if (paletteControls.keyInput.isEditing()) return;
            applyPreferencesToControls(paletteControls, readAllPreferences());
        };

        palette.show();

        /* ボタンの高さを4px詰める（レイアウト確定後に1回）/ Trim the button height by 4px (once, after layout) */
        trimButtonHeight(paletteControls.drawing.refreshButton, 4);
    }

    main();

}());
