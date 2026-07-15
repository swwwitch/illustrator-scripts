#target illustrator
#targetengine "AiQuickPrefsPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustrator の各種環境設定の切り替えを、常駐パレットでまとめて操作するユーティリティです。
操作した時点で即時反映されます。

- キー増加：カーソル移動量（cursorKeyLength）。単位ポップアップで定規単位を切替、↑↓ / Shift / Option で増減
- 整列オプション：プレビュー境界／字形の境界に整列（ポイント文字・エリア内文字を両方まとめてON/OFF）
- 変形オプション：パターン／角／線幅と効果
- コピー/ペースト：書式なしペースト／コピー元のレイヤーにペースト
- 描画：リアルタイムの描画と編集／プレビュー更新（GPU プレビューを更新）
- 明るさ：インターフェイスカラー（UIの明るさ）を4段階で切り替え。インターフェイスカラーは直接反映できないため、スウォッチをクリックすると環境設定（ユーザーインターフェイス）が開きます。矢印キー＋Return で確定してください


### 紹介記事（note）

https://note.com/dtp_tranist/n/n41d8dc1961be

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v2.0.3";

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
            brightness: { ja: "明るさ", en: "Brightness" }
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
            brightness: { ja: "インターフェイスカラーは直接反映できないため、クリックで環境設定が開きます。矢印キー＋Return で確定してください", en: "Interface color can't be applied directly, so clicking opens Preferences. Confirm with the arrow keys + Return." }
        },
        button: {
            refreshGpuPreview: { ja: "プレビュー更新", en: "Refresh Preview" }
        },
        brightness: {
            dark: { ja: "暗", en: "Dark" },
            mediumDark: { ja: "やや暗", en: "Medium Dark" },
            mediumLight: { ja: "やや明", en: "Medium Light" },
            light: { ja: "明", en: "Light" }
        }
    };

    /* ドット区切りキーで LABELS を辿り、現在言語の文言を返す / Resolve a dot-path key in LABELS to the current-language text */
    function L(key) {
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

    /* 単位の定義を1か所に集約（コード／ラベル／pt換算係数／ポップアップ表示）*/
    /* Single source of unit definitions (code, label, pt factor, popup visibility) */
    var UNITS = [
        { code: 0,  label: "in",    factor: 72.0,                 popup: true },
        { code: 1,  label: "mm",    factor: 72.0 / 25.4,          popup: true },
        { code: 2,  label: "pt",    factor: 1.0,                  popup: true },
        { code: 3,  label: "pica",  factor: 12.0,                 popup: true },
        { code: 4,  label: "cm",    factor: 72.0 / 2.54,          popup: true },
        { code: 5,  label: "Q/H",   factor: 72.0 / 25.4 * 0.25,   popup: true },
        { code: 6,  label: "px",    factor: 1.0,                  popup: true },
        { code: 7,  label: "ft/in", factor: 72.0 * 12.0,          popup: false },
        { code: 8,  label: "m",     factor: 72.0 / 25.4 * 1000.0, popup: false },
        { code: 9,  label: "yd",    factor: 72.0 * 36.0,          popup: false },
        { code: 10, label: "ft",    factor: 72.0 * 12.0,          popup: false }
    ];

    /* コードから単位定義を取得 / Find a unit definition by code */
    function getUnitByCode(code) {
        for (var i = 0; i < UNITS.length; i++) {
            if (UNITS[i].code === code) return UNITS[i];
        }
        return null;
    }

    /* 単位ラベルを取得 / Get the unit label */
    function getUnitLabel(code) {
        var unit = getUnitByCode(code);
        return unit ? unit.label : "pt";
    }

    /* 単位コードから pt への換算係数を取得 / Get the pt conversion factor from a unit code */
    function getPtFactorFromUnitCode(code) {
        var unit = getUnitByCode(code);
        return unit ? unit.factor : 1.0;
    }

    /* ポップアップに表示する単位コード（表示順、UNITS から派生）/ Unit codes shown in the popup (in order, derived from UNITS) */
    var UNIT_POPUP_CODES = (function () {
        var codes = [];
        for (var i = 0; i < UNITS.length; i++) {
            if (UNITS[i].popup) codes.push(UNITS[i].code);
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
            cursorKeyLength: getReal("cursorKeyLength"),
            uiBrightness: getReal("uiBrightness")
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
        var checkbox = parent.add('checkbox', undefined, L(labelKey));
        if (tooltipKey) checkbox.helpTip = L(tooltipKey);
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
        var panel = parent.add('panel', undefined, L('panel.keyInput'));
        panel.orientation = 'row';
        panel.alignChildren = ['left', 'center'];
        panel.margins = PANEL_MARGINS;

        var cursorStepField = panel.add('edittext', undefined, "1.0");
        cursorStepField.characters = 4;
        cursorStepField.helpTip = L('tooltip.cursorStep');

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
        unitDropdown.helpTip = L('tooltip.unit');

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
        var panel = parent.add('panel', undefined, L('panel.align'));
        setupPanel(panel);

        /* プレビュー境界 / Preview bounds */
        var checkboxPreview = addBooleanCheckbox(panel, 'checkbox.previewBounds', 'includeStrokeInBounds', 'tooltip.previewBounds');

        /* 字形の境界に整列：ポイント文字・エリア内文字の両方をまとめてON/OFF / Align to glyph bounds: toggle both point & area type together */
        var checkboxGlyphBounds = panel.add('checkbox', undefined, L('checkbox.glyphBounds'));
        checkboxGlyphBounds.helpTip = L('tooltip.glyphBounds');
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
        var panel = parent.add('panel', undefined, L('panel.transform'));
        setupPanel(panel);

        /* パターンを変形 / Transform patterns */
        var checkboxPattern = panel.add('checkbox', undefined, L('checkbox.transformPattern'));
        checkboxPattern.helpTip = L('tooltip.transformPattern');

        /* 角を拡大・縮小（1=ON, 2=OFF。Boolean でなく整数）/ Scale corners (1=ON, 2=OFF; integer pref) */
        var checkboxCorner = panel.add('checkbox', undefined, L('checkbox.scaleCorners'));
        checkboxCorner.helpTip = L('tooltip.scaleCorners');

        /* 線幅と効果も拡大・縮小 / Scale strokes and effects */
        var checkboxStroke = panel.add('checkbox', undefined, L('checkbox.scaleStroke'));
        checkboxStroke.helpTip = L('tooltip.scaleStroke');

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
        var panel = parent.add('panel', undefined, L('panel.copyPaste'));
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
        var panel = parent.add('panel', undefined, L('panel.drawing'));
        setupPanel(panel);

        /* リアルタイムの描画と編集（上段）＋ 更新ボタン（次の行）を縦並び / Real-time editing checkbox (top) + Refresh button (next line), stacked */

        /* リアルタイムの描画と編集 / Real-time drawing & editing */
        var checkboxRealtime = addBooleanCheckbox(panel, 'checkbox.realtimeDrawing', 'LiveEdit_State_Machine', 'tooltip.realtimeDrawing');

        /* GPU プレビューを更新（View using GPU を2回トグルして再描画）/ Refresh GPU preview (toggle View using GPU twice to redraw) */
        var btnRefreshGpuPreview = panel.add('button', undefined, L('button.refreshGpuPreview'));
        btnRefreshGpuPreview.alignment = ['left', 'top']; /* 幅いっぱいにしない（ラベル幅）/ Do not fill width (size to label) */
        btnRefreshGpuPreview.helpTip = L('tooltip.refreshGpuPreview');
        btnRefreshGpuPreview.onClick = function () {
            runInMainEngine('try{app.executeMenuCommand("View using GPU");app.executeMenuCommand("View using GPU");}catch(e){}');
        };

        return { realtime: checkboxRealtime, refreshButton: btnRefreshGpuPreview };
    }

    /* UI明るさの4プリセット：uiBrightness 値・表示シェード（RGB 0..1）・ラベル / Four UI-brightness presets: uiBrightness value, swatch shade (RGB 0..1), label */
    /* 値は連続値ではなく離散プリセット（0.5 と 0.50999999 が別段階）/ Values are discrete presets, not a continuous scale (0.5 vs 0.50999999 are distinct steps) */
    var BRIGHTNESS_LEVELS = [
        { value: 0.0,               shade: [0.22, 0.22, 0.22], labelKey: 'brightness.dark' },
        { value: 0.5,               shade: [0.33, 0.33, 0.33], labelKey: 'brightness.mediumDark' },
        { value: 0.50999999046326,  shade: [0.70, 0.70, 0.70], labelKey: 'brightness.mediumLight' },
        { value: 1.0,               shade: [0.94, 0.94, 0.94], labelKey: 'brightness.light' }
    ];
    var BRIGHTNESS_SWATCH_SIZE = 23;                    /* スウォッチの一辺(px) / Swatch side length (px) */
    var BRIGHTNESS_SELECTED_BORDER = [0.15, 0.5, 0.92]; /* 選択枠の青 / Blue selection border */
    var BRIGHTNESS_SWATCH_OUTLINE = [0.5, 0.5, 0.5];    /* 通常時の細枠 / Thin outline when not selected */

    /* uiBrightness 値に最も近いプリセットの index を返す / Index of the preset closest to a uiBrightness value */
    function nearestBrightnessIndex(value) {
        var best = 0, bestDiff = 1e9;
        for (var i = 0; i < BRIGHTNESS_LEVELS.length; i++) {
            var diff = Math.abs(BRIGHTNESS_LEVELS[i].value - value);
            if (diff < bestDiff) { bestDiff = diff; best = i; }
        }
        return best;
    }

    /* 明るさパネルを構築：4段階のシェードを onDraw で自前描画するボタン。{setByValue} を返す */
    /* Build the Brightness panel: four shade swatches drawn by onDraw. Returns {setByValue} */
    function buildBrightnessPanel(parent) {
        var panel = parent.add('panel', undefined, L('panel.brightness'));
        setupPanel(panel);

        var row = panel.add('group');
        row.orientation = 'row';
        row.alignment = ['center', 'top'];
        row.spacing = 8;

        var buttons = [];
        var selectedIndex = -1;

        /* onDraw ハンドラ：塗り＋（選択時のみ）青い枠を描く / onDraw handler: fill + (only when selected) a blue border */
        function makeDraw(index) {
            return function () {
                var g = this.graphics;
                var w = this.size[0], h = this.size[1];
                var level = BRIGHTNESS_LEVELS[index];

                /* シェードで塗りつぶし / Fill with the shade */
                var brush = g.newBrush(g.BrushType.SOLID_COLOR, level.shade.concat(1));
                g.newPath();
                g.rectPath(0, 0, w, h);
                g.fillPath(brush);

                if (index === selectedIndex) {
                    /* 選択：太めの青枠 / Selected: thicker blue border */
                    var penSel = g.newPen(g.PenType.SOLID_COLOR, BRIGHTNESS_SELECTED_BORDER.concat(1), 2);
                    g.newPath();
                    g.rectPath(1, 1, w - 2, h - 2);
                    g.strokePath(penSel);
                } else {
                    /* 非選択：細いグレー枠（明シェードを明背景でも視認）/ Not selected: thin gray outline (keep light swatches visible) */
                    var penOutline = g.newPen(g.PenType.SOLID_COLOR, BRIGHTNESS_SWATCH_OUTLINE.concat(1), 1);
                    g.newPath();
                    g.rectPath(0.5, 0.5, w - 1, h - 1);
                    g.strokePath(penOutline);
                }
            };
        }

        /* 全ボタンを再描画（hide/show で onDraw を確実に再実行し、旧選択枠を残さない＝排他表示）/ Force all buttons to repaint (hide/show reliably re-runs onDraw so the old selection border never lingers = exclusive) */
        function refresh() {
            for (var i = 0; i < buttons.length; i++) {
                buttons[i].hide();
                buttons[i].show();
            }
        }

        /* uiBrightness 値から選択状態だけ更新（適用はしない）。値が変わらなければ再描画しない */
        /* Update selection only, from a uiBrightness value (no apply). Skip repaint when unchanged */
        /* onActivate は毎クリック発火するため、無駄な hide/show を避けてウィンドウ z-order の乱れを防ぐ */
        /* onActivate fires on every click, so avoid needless hide/show that disturbs window z-order */
        function setByValue(value) {
            var nextIndex = nearestBrightnessIndex(value);
            if (nextIndex === selectedIndex) return;
            selectedIndex = nextIndex;
            refresh();
        }

        for (var i = 0; i < BRIGHTNESS_LEVELS.length; i++) {
            var button = row.add('iconbutton', undefined, undefined, { style: 'toolbutton' });
            button.preferredSize = [BRIGHTNESS_SWATCH_SIZE, BRIGHTNESS_SWATCH_SIZE];
            button.helpTip = L(BRIGHTNESS_LEVELS[i].labelKey) + " / " + L('tooltip.brightness');
            button.onDraw = makeDraw(i);
            (function (index, level) {
                button.onClick = function () {
                    selectedIndex = index;
                    /* 明るさを適用 → redraw → 環境設定(ユーザーインターフェイス)を開く。あとは矢印＋Return を手動で確定して反映 */
                    /* Apply brightness → redraw → open Preferences (User Interface). Confirm with arrows + Return manually to apply */
                    /* redraw はドキュメント未オープンだと例外になるため try で保護（以降の UIPref を止めない）/ Guard redraw (it throws with no open document) so it never blocks the following UIPref */
                    var body =
                        'app.preferences.setRealPreference("uiBrightness", ' + Number(level.value) + '); ' +
                        'try { app.redraw(); } catch (e) {} ' +
                        'app.executeMenuCommand("UIPref");';
                    runInMainEngine(body);
                    refresh();
                };
            })(i, BRIGHTNESS_LEVELS[i]);
            buttons.push(button);
        }

        return { setByValue: setByValue };
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

        var dialog = new Window('palette', L('dialog.title') + ' ' + SCRIPT_VERSION);
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

        /* 明るさパネル（最下部）。4段階のUI明るさスウォッチ / Brightness panel (bottom); four UI-brightness swatches */
        var brightnessControls = buildBrightnessPanel(mainGroup);


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

            /* UI明るさの選択スウォッチを現在値に同期 / Sync the selected brightness swatch to the current value */
            var uiBrightnessValue = parseFloat(prefValues['uiBrightness']);
            if (!isNaN(uiBrightnessValue)) brightnessControls.setByValue(uiBrightnessValue);

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
    }

    main();

}());
