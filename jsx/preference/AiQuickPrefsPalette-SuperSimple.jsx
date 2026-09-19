#target illustrator
#targetengine "AiQuickPrefsPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustratorの使用頻度の高い環境設定を、常駐パレットでまとめて切り替えます。
チェックや入力を操作したその場で反映されるため、環境設定ダイアログを開き直す手間がありません。

詳細は README を参照してください。

### Overview

A persistent palette for toggling the Illustrator preferences you use most often.
Every checkbox and field applies as you touch it, so there is no reopening of the Preferences dialog.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiQuickPrefsPalette-SuperSimple"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiQuickPrefsPalette-SuperSimple.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiQuickPrefsPalette-SuperSimple.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n41d8dc1961be"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DIALOG_OPACITY = 0.98;   /* パレットの不透明度 / Palette opacity */
    var SAVE_DEBOUNCE_MS = 40;   /* 保存デバウンス(ms) / Save debounce (ms) */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 現在のUI言語を取得 / Get the current UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義（カテゴリ分け）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "クイック環境設定", en: "Quick Preferences" }
        },
        panel: {
            keyInput: { ja: "キー増加", en: "Key Input" },
            align: { ja: "整列オプション", en: "Align Options" },
            transform: { ja: "変形オプション", en: "Transform Options" },
            copyPaste: { ja: "コピー/ペースト", en: "Copy / Paste" },
            drawing: { ja: "描画", en: "Drawing" },
            view: { ja: "表示", en: "View" }
        },
        checkbox: {
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
            transformPattern: { ja: "パターン", en: "Pattern Tiles" },
            scaleCorners: { ja: "角", en: "Corners" },
            scaleStroke: { ja: "線幅と効果", en: "Strokes & Effects" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集", en: "Real-time Drawing & Editing" },
            pastePlain: { ja: "書式なしペースト", en: "Paste without Formatting" },
            pastePreserve: { ja: "コピー元のレイヤーに", en: "Paste Remembers Layers" }
        },
        tooltip: {
            cursorStep: { ja: "矢印キーでの移動量（環境設定 > 一般 > キー増加）。↑↓ / Shift=±10 / Option=±0.1", en: "Keyboard increment (Preferences > General). Up/Down / Shift=±10 / Option=±0.1" },
            unit: { ja: "定規の単位を切り替え", en: "Switch the ruler unit" },
            previewBounds: { ja: "整列・分布でプレビュー境界（線幅・効果を含む）を使用", en: "Use preview bounds (incl. stroke & effects) for align/distribute" },
            glyphBounds: { ja: "ポイント文字・エリア内文字を字形の境界で整列", en: "Align point & area type to glyph bounds" },
            transformPattern: { ja: "変形時にパターンも変形する", en: "Transform pattern tiles when transforming" },
            scaleCorners: { ja: "拡大・縮小時に角（ライブコーナー）の半径も拡大・縮小", en: "Scale corner radius when scaling" },
            scaleStroke: { ja: "拡大・縮小時に線幅と効果も拡大・縮小", en: "Scale strokes & effects when scaling" },
            pastePlain: { ja: "書式を保持せずにペースト", en: "Paste without keeping formatting" },
            pastePreserve: { ja: "コピー元と同じレイヤーにペースト", en: "Paste into the original (source) layer" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集を切り替え", en: "Toggle real-time drawing & editing" },
            refreshGpuPreview: { ja: "GPUプレビューを更新（再描画）", en: "Refresh the GPU preview (redraw)" },
            edges: { ja: "エッジとバウンディングボックスの表示をまとめて切り替え", en: "Toggle edges and the bounding box together" },
            artboard: { ja: "アートボードの表示を切り替え", en: "Toggle artboard visibility" },
            videoRuler: { ja: "ビデオ定規の表示を切り替え", en: "Toggle video ruler visibility" },
            canvasColor: { ja: "カンバスカラーを「UIに合わせる」と「ホワイト」で切り替え", en: "Switch the canvas color between Match Brightness and White" }
        },
        button: {
            refreshGpuPreview: { ja: "プレビュー更新", en: "Refresh Preview" },
            edges: { ja: "境界線", en: "Edges" },
            artboard: { ja: "アートボード", en: "Artboards" },
            videoRuler: { ja: "ビデオ定規", en: "Video Ruler" },
            canvasColor: { ja: "カンバスカラー", en: "Canvas Color" }
        }
    };

    /* ドット区切りキーで LABELS を辿り、現在言語の文言を返す / Resolve a dot-path key in LABELS to the current-language text */
    function getLabel(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return key;
            node = node[parts[i]];
        }
        if (node == null) return key;
        return node[currentLanguage] || node.en || "";
    }

    // =========================================
    // 単位 / Unit
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）
       decimals：1pt 未満に潰れないように、大きい単位ほど桁数を増やす（in で 1mm ≒ 0.039）
       popup：単位ポップアップに並べるかどうか
       Unit table; the array index equals the rulerType code.
       decimals: larger units need more digits so small values do not collapse to 0 (1mm is 0.039in).
       popup: whether the unit appears in the unit popup. */
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
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    /* コードから単位定義を取得（未対応コードは pt 相当）/ Find a unit definition by code (unsupported codes fall back to pt) */
    function getUnitByCode(unitCode) {
        return UNITS[unitCode] || UNITS[2];
    }

    /* 単位ラベルを取得（定規単位なので歯は H 表示）/ Get the unit label (ruler unit, so unit code 5 shows as H) */
    function getUnitLabel(unitCode) {
        var unit = getUnitByCode(unitCode);
        return (unitCode === 5) ? "H" : unit.label;
    }

    /* 単位コードから pt への換算係数を取得 / Get the pt conversion factor from a unit code */
    function getPtFactorFromUnitCode(unitCode) {
        return getUnitByCode(unitCode).pointsPerUnit;
    }

    /* ポップアップに表示する単位コード（表示順、UNITS から派生）/ Unit codes shown in the popup (in order, derived from UNITS) */
    var UNIT_POPUP_CODES = (function () {
        var codes = [];
        for (var i = 0; i < UNITS.length; i++) {
            if (UNITS[i].popup) codes.push(i);
        }
        return codes;
    })();

    /* 単位コード → ポップアップのインデックス（未対応コードは -1）/ Unit code -> popup index (-1 if unsupported) */
    function unitCodeToPopupIndex(code) {
        for (var i = 0; i < UNIT_POPUP_CODES.length; i++) {
            if (UNIT_POPUP_CODES[i] === code) return i;
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

    /* 現在の定規単位の pt 換算係数を取得 / Get the pt factor for the current ruler unit */
    function getCurrentPtPerUnit() {
        return getPtFactorFromUnitCode(PREF_STATE.rulerType);
    }

    // =========================================
    // 環境設定の読み出し（同期・直接）/ Reading preferences (direct & synchronous)
    // 読み出しはエンジンを跨いでも安全なため、パレットエンジンで直接取得する
    // Reads are safe across engines, so fetch them directly in the palette engine
    // =========================================

    function readAllPreferences() {
        var prefs = app.preferences;
        function getBool(key) { try { return prefs.getBooleanPreference(key); } catch (e) { return false; } }
        function getInt(key) { try { return prefs.getIntegerPreference(key); } catch (e) { return 0; } }
        function getReal(key) { try { return prefs.getRealPreference(key); } catch (e) { return 0; } }
        return {
            EnableActualPointTextSpaceAlign: getBool("EnableActualPointTextSpaceAlign"),
            EnableActualAreaTextSpaceAlign: getBool("EnableActualAreaTextSpaceAlign"),
            includeStrokeInBounds: getBool("includeStrokeInBounds"),
            transformPatterns: getBool("transformPatterns"),
            scaleLineWeight: getBool("scaleLineWeight"),
            LiveEdit_State_Machine: getBool("LiveEdit_State_Machine"),
            pastePlain: getBool("plugin/FileClipboard/pasteWithoutFormatting"),
            pastePreserve: getBool("layers/pastePreserve"),
            policyForPreservingCorners: getInt("policyForPreservingCorners"),
            rulerType: getInt("rulerType"),
            cursorKeyLength: getReal("cursorKeyLength")
        };
    }

    // =========================================
    // BridgeTalk 委譲（書き込み）/ BridgeTalk delegation (writes)
    // =========================================

    /* メインエンジン（target="illustrator"）へ環境設定コードを送って実行 / Send preference code to the main engine and run it */
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
            /* BridgeTalk 不可時は同一エンジンで直接実行 / Fallback: run directly in this engine */
            try {
                eval(bodyCode);
            } catch (e2) {
                // no-op
            }
        }
    }

    /* 環境設定をメインエンジンで設定（値リテラルは型別ラッパーが整形）/ Set a preference on the main engine (per-type wrappers format the value literal) */
    function btSetPreference(method, prefKey, valueLiteral) {
        runInMainEngine('app.preferences.' + method + '("' + prefKey + '", ' + valueLiteral + ');');
    }

    /* 型別の薄いラッパー / Thin per-type wrappers */
    function btSetBooleanPreference(prefKey, value) { btSetPreference('setBooleanPreference', prefKey, value ? 'true' : 'false'); }
    function btSetIntegerPreference(prefKey, value) { btSetPreference('setIntegerPreference', prefKey, parseInt(value, 10)); }
    function btSetRealPreference(prefKey, value)    { btSetPreference('setRealPreference', prefKey, Number(value)); }

    // =========================================
    // カーソル移動量 / Cursor step
    // =========================================

    /* 現在単位の値を cursorKeyLength(pt) として保存 / Save value (in current unit) to cursorKeyLength as pt */
    function saveCursorKeyLengthInCurrentUnit(unitValue) {
        if (isNaN(unitValue) || unitValue < 0) return false;
        PREF_STATE.cursorKeyLengthPt = unitValue * getCurrentPtPerUnit();
        btSetRealPreference("cursorKeyLength", PREF_STATE.cursorKeyLengthPt);
        return true;
    }

    /* キャッシュ済み cursorKeyLength(pt) を現在単位の文字列(小数1桁)で取得 / Read cached cursorKeyLength as a current-unit string */
    function readCursorKeyLengthInCurrentUnit() {
        return (PREF_STATE.cursorKeyLengthPt / getCurrentPtPerUnit()).toFixed(1);
    }

    /* ===== デバウンス保存 / Debounced saving ===== */
    var __cursorKeyDebounceTaskId = null;
    var __cursorKeyPendingText = null;

    /* 保留中のテキストを実際に保存（scheduleTask から呼ばれる）。内部は例外を投げないため try 不要 / Save the pending text (called from scheduleTask); no try needed since the body cannot throw */
    function __runSaveCursorKeyLength() {
        if (__cursorKeyPendingText !== null) {
            saveCursorKeyLengthInCurrentUnit(parseFloat(__cursorKeyPendingText));
        }
        __cursorKeyDebounceTaskId = null;
    }
    /* scheduleTask の文字列はグローバルスコープで評価されるため $.global 経由で公開 / scheduleTask strings run in global scope, so expose via $.global */
    $.global.__aiQuickPrefsRunSave = __runSaveCursorKeyLength;

    /* 一定遅延後に保存をスケジュール（不可時は即時保存）/ Schedule a save after a short delay (immediate if unavailable) */
    function scheduleSaveCursorKeyLengthDebounced(editText, delayMs) {
        try {
            __cursorKeyPendingText = String(editText.text);
            if (__cursorKeyDebounceTaskId) {
                app.cancelTask(__cursorKeyDebounceTaskId);
                __cursorKeyDebounceTaskId = null;
            }
            __cursorKeyDebounceTaskId = app.scheduleTask("$.global.__aiQuickPrefsRunSave()", delayMs, false);
        } catch (e) {
            /* scheduleTask 不可時は即時保存 / Save immediately if scheduleTask is unavailable */
            saveCursorKeyLengthInCurrentUnit(parseFloat(editText.text));
        }
    }

    /* ↑↓キーで値を増減（Shift=±10で10の倍数にスナップ / Option=±0.1 / 通常=±1）。負値は0でクランプ */
    /* Arrow keys adjust the value (Shift = ±10 snapped to multiples of 10, Option = ±0.1, otherwise ±1); clamps at 0 */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var keyName = event.keyName;
            if (keyName !== "Up" && keyName !== "Down") {
                return; /* 非矢印キーでは処理しない（Tabで0に丸められるのを防止）/ Ignore non-arrow keys (prevents Tab rounding to 0) */
            }
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var direction = (keyName === "Up") ? 1 : -1; /* ↑=+ / ↓=- */

            if (keyboard.shiftKey) {
                /* Shift=±10：次の10の倍数へスナップ / Shift: snap to the next multiple of 10 */
                value = (direction > 0)
                    ? Math.ceil((value + 1) / 10) * 10
                    : Math.floor((value - 1) / 10) * 10;
            } else {
                /* Option=±0.1 / 通常=±1 / Option = ±0.1, otherwise ±1 */
                value += direction * (keyboard.altKey ? 0.1 : 1);
            }
            if (value < 0) value = 0;

            /* Option のみ小数第1位、その他は整数に丸め / Round to 1 decimal with Option, otherwise to integer */
            value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);

            event.preventDefault();
            /* 常に小数第1位で表示し、デバウンス保存 / Always show 1 decimal and save (debounced) */
            editText.text = (Math.round(value * 10) / 10).toFixed(1);
            scheduleSaveCursorKeyLengthDebounced(editText, SAVE_DEBOUNCE_MS);
        });
    }

    // =========================================
    // UI ヘルパー / UI helpers
    // =========================================

    /* Boolean 環境設定にバインドしたチェックボックスを生成して返す（tooltipKey 指定時は helpTip も設定）/ Create a checkbox bound to a boolean preference (sets helpTip when tooltipKey is given) */
    function addBooleanCheckbox(parent, labelKey, prefKey, tooltipKey) {
        var checkbox = parent.add('checkbox', undefined, getLabel(labelKey));
        if (tooltipKey) checkbox.helpTip = getLabel(tooltipKey);
        checkbox.onClick = function () {
            btSetBooleanPreference(prefKey, checkbox.value === true);
        };
        return checkbox;
    }

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
    function trimButtonHeight(button, px) {
        button.size = [button.size.width, button.size.height - px];
    }

    // =========================================
    // パネル構築 / Panel builders
    // =========================================

    /* キー増加パネルを構築（カーソル移動量＋単位ポップアップ）。編集中フラグと単位同期はビルダー内に閉じ込め、{field, isEditing(), syncDisplay()} を返す */
    /* Build the Key Input panel (cursor step + unit popup); the editing flag and unit-sync are kept inside, returns {field, isEditing(), syncDisplay()} */
    function buildKeyInputPanel(parent) {
        var panel = parent.add('panel', undefined, getLabel('panel.keyInput'));
        panel.orientation = 'row';
        panel.alignChildren = ['left', 'center'];
        panel.margins = PANEL_MARGINS;

        var cursorStepField = panel.add('edittext', undefined, "1.0");
        cursorStepField.characters = 4;
        cursorStepField.helpTip = getLabel('tooltip.cursorStep');

        /* 編集中フラグ：入力欄にフォーカスがある間は外部同期で値を上書きしない / Editing flag: don't let external sync overwrite while the field has focus */
        var isEditingCursorStep = false;
        cursorStepField.onActivate = function () { isEditingCursorStep = true; };
        cursorStepField.onDeactivate = function () { isEditingCursorStep = false; };

        var suppressUnitChange = false;
        var unitDropdown = panel.add('dropdownlist', undefined, []);
        for (var i = 0; i < UNIT_POPUP_CODES.length; i++) {
            unitDropdown.add('item', getUnitLabel(UNIT_POPUP_CODES[i]));
        }
        unitDropdown.preferredSize.width = 55;
        unitDropdown.helpTip = getLabel('tooltip.unit');

        /* 単位ポップアップ：選んだ単位を定規単位(rulerType)へ反映し、表示を再計算 / Unit popup: apply the chosen unit to rulerType and recompute the display */
        unitDropdown.onChange = function () {
            if (suppressUnitChange || !unitDropdown.selection) return;
            var code = UNIT_POPUP_CODES[unitDropdown.selection.index];
            PREF_STATE.rulerType = code;
            btSetIntegerPreference("rulerType", code);
            /* 保存済み pt 値は不変。新しい単位で再表示 / Stored pt value is unchanged; redisplay in the new unit */
            cursorStepField.text = readCursorKeyLengthInCurrentUnit();
        };

        changeValueByArrowKey(cursorStepField);

        /* キー増加の確定時に保存（不正値は元へ戻す）/ Save on commit (restore on invalid input) */
        cursorStepField.onChange = function () {
            if (!saveCursorKeyLengthInCurrentUnit(parseFloat(cursorStepField.text))) {
                cursorStepField.text = readCursorKeyLengthInCurrentUnit();
            }
        };

        /* PREF_STATE から表示（単位ポップアップ＋数値）を更新。onChange は発火させない / Refresh the display (unit popup + value) from PREF_STATE without firing onChange */
        function syncDisplay() {
            var unitPopupIndex = unitCodeToPopupIndex(PREF_STATE.rulerType);
            if (unitPopupIndex >= 0) {
                suppressUnitChange = true;
                unitDropdown.selection = unitPopupIndex;
                suppressUnitChange = false;
            }
            cursorStepField.text = readCursorKeyLengthInCurrentUnit();
        }

        return {
            field: cursorStepField,
            isEditing: function () { return isEditingCursorStep; },
            syncDisplay: syncDisplay
        };
    }

    /* 整列オプションパネルを構築（プレビュー境界／字形の境界に整列）。2チェックを {preview, glyphBounds} で返す。字形の境界はポイント文字・エリア内文字を両方まとめてトグル */
    /* Build the Align Options panel (Preview Bounds / Align to Glyph Bounds); returns {preview, glyphBounds}. Glyph bounds toggles both point & area type together */
    function buildAlignPanel(parent) {
        var panel = parent.add('panel', undefined, getLabel('panel.align'));
        setupPanel(panel);

        /* プレビュー境界 / Preview bounds */
        var checkboxPreview = addBooleanCheckbox(panel, 'checkbox.previewBounds', 'includeStrokeInBounds', 'tooltip.previewBounds');

        /* 字形の境界に整列：ポイント文字・エリア内文字の両方をまとめてON/OFF / Align to glyph bounds: toggle both point & area type together */
        var checkboxGlyphBounds = panel.add('checkbox', undefined, getLabel('checkbox.glyphBounds'));
        checkboxGlyphBounds.helpTip = getLabel('tooltip.glyphBounds');
        checkboxGlyphBounds.onClick = function () {
            var on = checkboxGlyphBounds.value === true;
            btSetBooleanPreference("EnableActualPointTextSpaceAlign", on);
            btSetBooleanPreference("EnableActualAreaTextSpaceAlign", on);
        };

        return { preview: checkboxPreview, glyphBounds: checkboxGlyphBounds };
    }

    /* 変形オプションパネルを構築（パターン／角／線幅と効果）。3チェックを {pattern, corner, stroke} で返す。Option+クリックで3つまとめてトグル */
    /* Build the Transform Options panel (Pattern / Corners / Strokes & Effects); returns {pattern, corner, stroke}. Option+click toggles all three */
    function buildTransformPanel(parent) {
        var panel = parent.add('panel', undefined, getLabel('panel.transform'));
        setupPanel(panel);

        /* パターンを変形 / Transform patterns */
        var checkboxPattern = panel.add('checkbox', undefined, getLabel('checkbox.transformPattern'));
        checkboxPattern.helpTip = getLabel('tooltip.transformPattern');

        /* 角を拡大・縮小（1=ON, 2=OFF。Boolean でなく整数）/ Scale corners (1=ON, 2=OFF; integer pref) */
        var checkboxCorner = panel.add('checkbox', undefined, getLabel('checkbox.scaleCorners'));
        checkboxCorner.helpTip = getLabel('tooltip.scaleCorners');

        /* 線幅と効果も拡大・縮小 / Scale strokes and effects */
        var checkboxStroke = panel.add('checkbox', undefined, getLabel('checkbox.scaleStroke'));
        checkboxStroke.helpTip = getLabel('tooltip.scaleStroke');

        /* 各チェックの現在値を環境設定へ反映（角は 1/2 の整数）/ Apply each checkbox value to its preference (corners is the integer 1/2) */
        function applyPattern() { btSetBooleanPreference("transformPatterns", checkboxPattern.value === true); }
        function applyCorner() { btSetIntegerPreference("policyForPreservingCorners", checkboxCorner.value ? 1 : 2); }
        function applyStroke() { btSetBooleanPreference("scaleLineWeight", checkboxStroke.value === true); }

        /* クリック時：Option 併用なら3つを同じ値に揃えてまとめて適用、通常は単独適用 / On click: with Option, set all three to the same value and apply together; otherwise apply just this one */
        function onTransformOptionClick(clicked) {
            if (ScriptUI.environment.keyboardState.altKey) {
                var newValue = clicked.value === true;
                checkboxPattern.value = newValue;
                checkboxCorner.value = newValue;
                checkboxStroke.value = newValue;
                applyPattern();
                applyCorner();
                applyStroke();
                return;
            }
            if (clicked === checkboxPattern) applyPattern();
            else if (clicked === checkboxCorner) applyCorner();
            else applyStroke();
        }

        checkboxPattern.onClick = function () { onTransformOptionClick(checkboxPattern); };
        checkboxCorner.onClick = function () { onTransformOptionClick(checkboxCorner); };
        checkboxStroke.onClick = function () { onTransformOptionClick(checkboxStroke); };

        return { pattern: checkboxPattern, corner: checkboxCorner, stroke: checkboxStroke };
    }

    /* コピー/ペーストパネルを構築（書式なしペースト／コピー元のレイヤーにペースト）。2チェックを {pastePlain, pastePreserve} で返す */
    /* Build the Copy / Paste panel (Paste without Formatting / Paste Remembers Layers); returns {pastePlain, pastePreserve} */
    function buildCopyPastePanel(parent) {
        var panel = parent.add('panel', undefined, getLabel('panel.copyPaste'));
        setupPanel(panel);

        /* 書式なしペースト / Paste without formatting */
        var checkboxPastePlain = addBooleanCheckbox(panel, 'checkbox.pastePlain', 'plugin/FileClipboard/pasteWithoutFormatting', 'tooltip.pastePlain');

        /* コピー元のレイヤーにペースト / Paste remembers layers */
        var checkboxPastePreserve = addBooleanCheckbox(panel, 'checkbox.pastePreserve', 'layers/pastePreserve', 'tooltip.pastePreserve');

        return { pastePlain: checkboxPastePlain, pastePreserve: checkboxPastePreserve };
    }

    /* 描画パネルを構築（リアルタイムの描画と編集＋プレビュー更新ボタン）。{realtime, refreshButton} を返す。refreshButton はレイアウト確定後の trimButtonHeight 用 */
    /* Build the Drawing panel (Real-time Drawing & Editing + Refresh Preview button); returns {realtime, refreshButton}. refreshButton is for trimButtonHeight after layout */
    function buildDrawingPanel(parent) {
        var panel = parent.add('panel', undefined, getLabel('panel.drawing'));
        setupPanel(panel);

        /* リアルタイムの描画と編集（上段）＋ 更新ボタン（次の行）を縦並び / Real-time editing checkbox (top) + Refresh button (next line), stacked */

        /* リアルタイムの描画と編集 / Real-time drawing & editing */
        var checkboxRealtime = addBooleanCheckbox(panel, 'checkbox.realtimeDrawing', 'LiveEdit_State_Machine', 'tooltip.realtimeDrawing');

        /* GPU プレビューを更新（View using GPU を2回トグルして再描画）/ Refresh GPU preview (toggle View using GPU twice to redraw) */
        var btnRefreshGpuPreview = panel.add('button', undefined, getLabel('button.refreshGpuPreview'));
        btnRefreshGpuPreview.alignment = ['left', 'top']; /* 幅いっぱいにしない（ラベル幅）/ Do not fill width (size to label) */
        btnRefreshGpuPreview.helpTip = getLabel('tooltip.refreshGpuPreview');
        btnRefreshGpuPreview.onClick = function () {
            runInMainEngine('try{app.executeMenuCommand("View using GPU");app.executeMenuCommand("View using GPU");}catch(e){}');
        };

        return { realtime: checkboxRealtime, refreshButton: btnRefreshGpuPreview };
    }

    /* 表示パネルのボタン行を追加（左右中央・2列）/ Add a button row to the View panel (centered, two columns) */
    function addViewButtonRow(panel) {
        var row = panel.add('group');
        row.orientation = 'row';
        row.alignment = ['fill', 'top'];
        row.alignChildren = ['center', 'top']; /* 左右中央 / Center horizontally */
        row.spacing = 8;
        return row;
    }

    /* 表示パネルを構築：境界線／アートボード／ビデオ定規／カンバスカラーのトグルボタン。各ボタンはレイアウト確定後の trimButtonHeight 用に返す */
    /* Build the View panel: Edges / Artboards / Video Ruler / Canvas Color toggle buttons; each button is returned for trimButtonHeight after layout */
    function buildViewPanel(parent) {
        var panel = parent.add('panel', undefined, getLabel('panel.view'));
        setupPanel(panel);

        /* 幅を抑えるため2列×2行に配置 / Two columns by two rows, to keep the palette narrow */
        var firstRow = addViewButtonRow(panel);
        var secondRow = addViewButtonRow(panel);

        var btnEdges = firstRow.add('button', undefined, getLabel('button.edges'));
        btnEdges.helpTip = getLabel('tooltip.edges');
        btnEdges.onClick = function () {
            /* edge＝エッジの表示/隠す、AI Bounding Box Toggle＝バウンディングボックスの表示/隠す。2つをまとめてトグル */
            /* edge = Show/Hide Edges, AI Bounding Box Toggle = Show/Hide Bounding Box; toggle both at once */
            runInMainEngine("try{app.executeMenuCommand('edge');app.executeMenuCommand('AI Bounding Box Toggle');}catch(e){}");
        };

        /* アートボードの表示/隠すをトグル / Toggle Show/Hide Artboards */
        var btnArtboard = firstRow.add('button', undefined, getLabel('button.artboard'));
        btnArtboard.helpTip = getLabel('tooltip.artboard');
        btnArtboard.onClick = function () {
            runInMainEngine("try{app.executeMenuCommand('artboard');}catch(e){}");
        };

        /* ビデオ定規の表示/隠すをトグル / Toggle Show/Hide Video Rulers */
        var btnVideoRuler = secondRow.add('button', undefined, getLabel('button.videoRuler'));
        btnVideoRuler.helpTip = getLabel('tooltip.videoRuler');
        btnVideoRuler.onClick = function () {
            runInMainEngine("try{app.executeMenuCommand('videoruler');}catch(e){}");
        };

        /* カンバスカラー：uiCanvasIsWhite を 0（UIに合わせる）と 1（ホワイト）で交互に切替 */
        /* Canvas Color: flip uiCanvasIsWhite between 0 (Match Brightness) and 1 (White) */
        var btnCanvasColor = secondRow.add('button', undefined, getLabel('button.canvasColor'));
        btnCanvasColor.helpTip = getLabel('tooltip.canvasColor');
        btnCanvasColor.onClick = function () {
            /* 現在値の読み出しから反転・保存・再描画までをメインエンジン側で完結（エンジン間で値がずれないように） */
            /* Read, flip, save, and repaint all in the main engine so the two engines can't disagree on the value */
            /* 反映には画面の強制再描画が必要。redraw() だけでは足りないため、ズーム操作で描き直す（PresetManager と同じ手当て） */
            /* Applying it needs a forced repaint; redraw() alone is not enough, so toggle the zoom (same workaround as PresetManager) */
            var body =
                'var isWhite = 0; ' +
                'try { isWhite = app.preferences.getIntegerPreference("uiCanvasIsWhite"); } catch (e) {} ' +
                'app.preferences.setIntegerPreference("uiCanvasIsWhite", (isWhite === 1) ? 0 : 1); ' +
                'try { app.redraw(); app.executeMenuCommand("zoomout"); app.executeMenuCommand("zoomin"); } catch (e) {}';
            runInMainEngine(body);
        };

        return {
            edgesButton: btnEdges,
            artboardButton: btnArtboard,
            videoRulerButton: btnVideoRuler,
            canvasColorButton: btnCanvasColor
        };
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /* パレットを構築して表示 / Build and show the palette */
    function main() {

        /* すでにパレットが開いていれば前面に出して終了 / If a palette already exists, bring it forward and return */
        try {
            if ($.global.__aiQuickPrefsPalette) {
                $.global.__aiQuickPrefsPalette.show();
                return;
            }
        } catch (e) {
            $.global.__aiQuickPrefsPalette = null;
        }

        /* 初期表示用に現在値を読み込む（PREF_STATE への反映は後段の applyPreferencesToUI(initialPrefs) が担う）/ Load current values (applyPreferencesToUI(initialPrefs) below seeds PREF_STATE with the same fallbacks) */
        var initialPrefs = readAllPreferences();

        var dialog = new Window('palette', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];
        dialog.opacity = DIALOG_OPACITY;
        $.global.__aiQuickPrefsPalette = dialog;

        /* 閉じたら参照をクリア（次回は再構築）/ Clear the reference on close (rebuild next time) */
        dialog.onClose = function () {
            $.global.__aiQuickPrefsPalette = null;
        };

        /* パレットがアクティブなとき esc キーで閉じる / Close the palette with Esc while it is active */
        dialog.addEventListener('keydown', function (event) {
            if (event.keyName === 'Escape') {
                dialog.close();
            }
        });

        var mainGroup = dialog.add('group');
        mainGroup.orientation = 'column';
        mainGroup.alignChildren = ['fill', 'top'];

        /* ----- キー増加 / 整列オプション / 変形オプション（1カラムで縦積み）/ Key input / Align Options / Transform Options (single column, stacked) ----- */

        /* キー増加パネル（カーソル移動量＋単位ポップアップ）/ Key input panel (cursor step + unit popup) */
        var keyInput = buildKeyInputPanel(mainGroup);
        var cursorStepField = keyInput.field;

        /* 整列オプションパネル（キー増加の下）。2チェックの参照を受け取る / Align Options panel (below Key input); receive the 2 checkbox refs */
        var alignControls = buildAlignPanel(mainGroup);
        var checkboxPreview = alignControls.preview;
        var checkboxGlyphBounds = alignControls.glyphBounds;

        /* 変形オプションパネル（整列オプションの下）。3チェックの参照を受け取る / Transform Options panel (below Align Options); receive the 3 checkbox refs */
        var transformControls = buildTransformPanel(mainGroup);
        var checkboxPattern = transformControls.pattern;
        var checkboxCorner = transformControls.corner;
        var checkboxStroke = transformControls.stroke;

        /* ----- 全幅：コピー/ペースト / 描画 / Full width: Copy / Paste / Drawing ----- */

        /* コピー/ペーストパネル。2チェックの参照を受け取る / Copy / Paste panel; receive the 2 checkbox refs */
        var copyPasteControls = buildCopyPastePanel(mainGroup);
        var checkboxPastePlain = copyPasteControls.pastePlain;
        var checkboxPastePreserve = copyPasteControls.pastePreserve;

        /* 描画パネル（コピー/ペーストの下）。realtime チェックと更新ボタンを受け取る / Drawing panel (below Copy / Paste); receive the realtime checkbox and refresh button */
        var drawingControls = buildDrawingPanel(mainGroup);
        var checkboxRealtime = drawingControls.realtime;
        var btnRefreshGpuPreview = drawingControls.refreshButton;

        /* 表示パネル（最下部）。境界線／アートボード／ビデオ定規／カンバスカラーの切替ボタン / View panel (bottom); Edges / Artboards / Video Ruler / Canvas Color toggles */
        var viewControls = buildViewPanel(mainGroup);

        /* 読み出した環境設定を UI へ反映 / Apply fetched preferences to the UI */
        function applyPreferencesToUI(prefValues) {
            function asBool(prefKey) {
                return prefValues[prefKey] === true;
            }
            checkboxPreview.value = asBool('includeStrokeInBounds');
            /* ポイント文字・エリア内文字が両方ONのときだけチェック / Checked only when both point & area type are on */
            checkboxGlyphBounds.value = asBool('EnableActualPointTextSpaceAlign') && asBool('EnableActualAreaTextSpaceAlign');
            checkboxPattern.value = asBool('transformPatterns');
            checkboxStroke.value = asBool('scaleLineWeight');
            checkboxRealtime.value = asBool('LiveEdit_State_Machine');
            checkboxPastePlain.value = asBool('pastePlain');
            checkboxPastePreserve.value = asBool('pastePreserve');
            checkboxCorner.value = (parseInt(prefValues['policyForPreservingCorners'], 10) === 1);

            var rulerTypeCode = parseInt(prefValues['rulerType'], 10);
            PREF_STATE.rulerType = isNaN(rulerTypeCode) ? PREF_STATE.rulerType : rulerTypeCode;
            var cursorKeyLengthPt = parseFloat(prefValues['cursorKeyLength']);
            PREF_STATE.cursorKeyLengthPt = isNaN(cursorKeyLengthPt) ? PREF_STATE.cursorKeyLengthPt : cursorKeyLengthPt;

            /* 単位ポップアップと数値表示を PREF_STATE に同期 / Sync the unit popup and value display to PREF_STATE */
            keyInput.syncDisplay();
        }

        /* 構築直後に現在値を反映（常に最新の環境設定を表示）/ Populate from current values right after building */
        applyPreferencesToUI(initialPrefs);

        /* 表示時：レイアウト再計算（描画欠け防止）＋フォーカス＋最新値の再読込 / On show: recalc layout (avoid partial rendering), focus, and reload current values */
        dialog.onShow = function () {
            dialog.layout.layout(true);
            dialog.layout.resize();
            cursorStepField.active = true;
            applyPreferencesToUI(readAllPreferences());
        };

        /* 再アクティブ時：外部変更（環境設定ダイアログ等）へクリックで追従。ただし編集中は同期しない（入力値の上書き防止）/ On re-activate: follow external changes via click, but skip while editing (avoid clobbering input) */
        dialog.onActivate = function () {
            if (keyInput.isEditing()) return;
            applyPreferencesToUI(readAllPreferences());
        };

        dialog.show();

        /* ボタンの高さを4px詰める（レイアウト確定後に1回）/ Trim button heights by 4px (once, after layout) */
        trimButtonHeight(btnRefreshGpuPreview, 4);
        trimButtonHeight(viewControls.edgesButton, 4);
        trimButtonHeight(viewControls.artboardButton, 4);
        trimButtonHeight(viewControls.videoRulerButton, 4);
        trimButtonHeight(viewControls.canvasColorButton, 4);
    }

    main();

}());
