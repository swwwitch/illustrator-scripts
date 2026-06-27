#target illustrator
#targetengine "AiQuickPrefsPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustrator の各種環境設定を、常駐パレットでまとめて切り替えるユーティリティです。設定は操作した時点で即時反映されます。

- パレット（常駐エンジン）で表示し、書き込みは BridgeTalk でメインエンジンへ委譲（読み出しは同期で直接取得）
- 2カラム構成：左＝キー入力／変形と整列、右＝字形の境界に整列／ガイドと定規、最下部（全幅）＝アートボード名と枠線／その他
- 環境設定ダイアログ等の外部変更は、パレットをクリック（再アクティブ）で同期

#### パネルと項目

- キー入力：カーソル移動量（cursorKeyLength）。単位ポップアップで定規単位を切替、↑↓ / Shift / Option で増減
- 変形と整列：プレビュー境界／パターンを変形／角を拡大・縮小／線幅と効果も拡大・縮小
- 字形の境界に整列：ポイント文字／エリア内文字
- ガイドと定規：ガイドを表示／ガイドをロック／ビデオ定規（トグル）
- アートボード名と枠線：アートボード名を表示／カンバスカラーをホワイトに／枠線のカラー・幅
- その他：リアルタイムの描画と編集／書式なしペースト


### 紹介記事（note）

https://note.com/dtp_tranist/n/n41d8dc1961be

### Overview

A persistent-palette utility for batch-toggling various Illustrator preferences. Every setting applies immediately when changed.

- Runs in a persistent-engine palette; writes are delegated to the main engine via BridgeTalk (reads are fetched directly/synchronously)
- Two-column layout: left = Key input / Transform & Align, right = Glyph bounds / Guides & Rulers, bottom (full width) = Artboard / Other
- External changes (e.g. the Preferences dialog) sync when you click (re-activate) the palette

#### Panels & options

- Key input: cursor step (cursorKeyLength); switch ruler unit via popup, adjust with Up/Down / Shift / Option
- Transform & Align: Preview Bounds / Transform Patterns / Scale Corners / Scale Strokes & Effects
- Align to Glyph Bounds: Point Type / Area Type
- Guides & Rulers: Show Guides / Lock Guides / Video Ruler (toggle)
- Artboard Name & Border: Show Artboard Name / Canvas color to white / border color & width
- Other: Real-time Drawing & Editing / Paste without Formatting

### 更新履歴 / Change Log

- v1.6.0 (20260627): 標準フォーマットへ整理（IIFE 化、ローカライズ構造、ブロックコメント）。パレット化＋BridgeTalk 委譲、キー入力の単位ポップアップ、クリック同期、ガイド／アートボードパネルを追加、2カラムレイアウト。
- v1.0 (20250804): 初期バージョン。

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v1.6.0";

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
    var lang = getCurrentLang();

    /* 日英ラベル定義（カテゴリ分け）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        dialog: {
            title: { ja: "環境設定：変形と整列", en: "Preferences: Transform & Align" }
        },
        panel: {
            keyInput: { ja: "キー入力", en: "Key Input" },
            transform: { ja: "変形と整列", en: "Transform & Align" },
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            guide: { ja: "ガイドと定規", en: "Guides & Rulers" },
            artboard: { ja: "アートボード名と枠線", en: "Artboard Name & Border" },
            artboardBorder: { ja: "アートボードの枠線", en: "Artboard Border" },
            etc: { ja: "その他", en: "Other" }
        },
        checkbox: {
            pointText: { ja: "ポイント文字", en: "Point Type" },
            areaText: { ja: "エリア内文字", en: "Area Type" },
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
            transformPattern: { ja: "パターンを変形", en: "Transform Pattern Tiles" },
            scaleCorners: { ja: "角を拡大・縮小", en: "Scale Corners" },
            scaleStroke: { ja: "線幅と効果も拡大・縮小", en: "Scale Strokes & Effects" },
            guideShow: { ja: "ガイドを表示", en: "Show Guides" },
            guideLock: { ja: "ガイドをロック", en: "Lock Guides" },
            showArtboardName: { ja: "アートボード名を表示", en: "Show Artboard Name" },
            canvasWhite: { ja: "カンバスカラーをホワイトに", en: "Set Canvas Color to White" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集", en: "Real-time Drawing & Editing" },
            pastePlain: { ja: "書式なしペースト", en: "Paste without Formatting" }
        },
        label: {
            strokeColor: { ja: "ハイライトのカラー", en: "Highlight Color" },
            strokeWidth: { ja: "ストロークの幅", en: "Stroke Width" }
        },
        button: {
            videoRuler: { ja: "ビデオ定規", en: "Video Ruler" }
        },
        color: {
            lightBlue: { ja: "ライトブルー", en: "Light Blue" },
            red: { ja: "サーモンピンク", en: "Light Red" },
            green: { ja: "グリーン", en: "Green" },
            blue: { ja: "ミディアムブルー", en: "Medium Blue" },
            magenta: { ja: "マゼンタ", en: "Magenta" },
            cyan: { ja: "シアン", en: "Cyan" },
            grey: { ja: "ライトグレー", en: "Light Gray" },
            black: { ja: "ブラック", en: "Black" },
            yellow: { ja: "イエロー", en: "Yellow" }
        }
    };

    /* ドット区切りキーで LABELS を辿り現在言語の文言を返す（{slash}→/）/ Resolve a dot-path key in LABELS to the current-language text ({slash}→/) */
    function getLabel(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node == null) return key;
            node = node[parts[i]];
        }
        if (node == null) return key;
        var text = node[lang] || node.en || "";
        return text.replace(/\{slash\}/g, "/");
    }

    /* 現在言語のラベル文字列を返す / Return the current-language label string */
    function L(key) {
        return getLabel(key);
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return getLabel(key) + (lang === "ja" ? "：" : ":");
    }

    /* 件数付きラベル（日本語は全角括弧、英語は半角括弧）/ Label with count (full-width JA parentheses, half-width EN parentheses) */
    function labelWithCount(key, count) {
        if (lang === "ja") {
            return L(key) + "（" + count + "）";
        }
        return L(key) + " (" + count + ")";
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
    // アートボード枠線 / Artboard border
    // =========================================

    /* 枠線カラーのプリセット（ドロップダウンの並び順）/ Border color presets (dropdown order) */
    var STROKE_COLOR_PRESETS = [
        { labelKey: "color.lightBlue", r: 0.29, g: 0.52, b: 1.0 },
        { labelKey: "color.red",       r: 1.0,  g: 0.29, b: 0.29 },
        { labelKey: "color.green",     r: 0.0,  g: 0.65, b: 0.31 },
        { labelKey: "color.blue",      r: 0.0,  g: 0.45, b: 0.78 },
        { labelKey: "color.magenta",   r: 1.0,  g: 0.0,  b: 1.0 },
        { labelKey: "color.cyan",      r: 0.0,  g: 1.0,  b: 1.0 },
        { labelKey: "color.grey",      r: 0.65, g: 0.65, b: 0.65 },
        { labelKey: "color.black",     r: 0.0,  g: 0.0,  b: 0.0 },
        { labelKey: "color.yellow",    r: 1.0,  g: 1.0,  b: 0.0 }
    ];
    var STROKE_COLOR_BLACK_INDEX = 7;

    /* ドロップダウン用のカラー名配列を生成 / Build the color name list for the dropdown */
    function buildStrokeColorNames() {
        var names = [];
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            names.push(L(STROKE_COLOR_PRESETS[i].labelKey));
        }
        return names;
    }

    /* RGB に最も近いプリセットの index を返す / Return the index of the preset closest to the given RGB */
    function findClosestStrokeColor(r, g, b) {
        var bestIdx = 0;
        var bestDist = Infinity;
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            var p = STROKE_COLOR_PRESETS[i];
            var dist = Math.abs(p.r - r) + Math.abs(p.g - g) + Math.abs(p.b - b);
            if (dist < bestDist) {
                bestDist = dist;
                bestIdx = i;
            }
        }
        return bestIdx;
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
        var p = app.preferences;
        function gb(key) { try { return p.getBooleanPreference(key); } catch (e) { return false; } }
        function gi(key) { try { return p.getIntegerPreference(key); } catch (e) { return 0; } }
        function gr(key) { try { return p.getRealPreference(key); } catch (e) { return 0; } }
        return {
            EnableActualPointTextSpaceAlign: gb("EnableActualPointTextSpaceAlign"),
            EnableActualAreaTextSpaceAlign: gb("EnableActualAreaTextSpaceAlign"),
            includeStrokeInBounds: gb("includeStrokeInBounds"),
            transformPatterns: gb("transformPatterns"),
            scaleLineWeight: gb("scaleLineWeight"),
            showGuides: gb("showGuides"),
            lockGuides: gb("lockGuides"),
            showArtboardLabelOnCanvas: gb("showArtboardLabelOnCanvas"),
            ArtboardBBColorRed: gr("ArtboardBBColorRed"),
            ArtboardBBColorGreen: gr("ArtboardBBColorGreen"),
            ArtboardBBColorBlue: gr("ArtboardBBColorBlue"),
            ArtboardBBWidth: gr("ArtboardBBWidth"),
            LiveEdit_State_Machine: gb("LiveEdit_State_Machine"),
            pastePlain: gb("plugin/FileClipboard/pasteWithoutFormatting"),
            policyForPreservingCorners: gi("policyForPreservingCorners"),
            uiCanvasIsWhite: gi("uiCanvasIsWhite"),
            rulerType: gi("rulerType"),
            cursorKeyLength: gr("cursorKeyLength")
        };
    }

    // =========================================
    // BridgeTalk 委譲（書き込み）/ BridgeTalk delegation (writes)
    // =========================================

    /* メインエンジン（target="illustrator"）へ環境設定コードを送って実行 / Send preference code to the main engine and run it */
    function runInMainEngine(bodyCode) {
        try {
            var bt = new BridgeTalk();
            bt.target = "illustrator"; /* #targetengine 指定なし＝メインエンジン / no engine = main engine */
            bt.body = bodyCode;
            bt.onError = function (message) {
                /* no-op: 失敗時は既存の値を保持 / keep existing values on failure */
            };
            bt.send();
        } catch (e) {
            /* BridgeTalk 不可時は同一エンジンで直接実行 / Fallback: run directly in this engine */
            try {
                eval(bodyCode);
            } catch (e2) {
                // no-op
            }
        }
    }

    /* Boolean 環境設定をメインエンジンで設定 / Set a boolean preference on the main engine */
    function btSetBooleanPreference(prefKey, value) {
        runInMainEngine('app.preferences.setBooleanPreference("' + prefKey + '", ' + (value ? 'true' : 'false') + ');');
    }

    /* Integer 環境設定をメインエンジンで設定 / Set an integer preference on the main engine */
    function btSetIntegerPreference(prefKey, value) {
        runInMainEngine('app.preferences.setIntegerPreference("' + prefKey + '", ' + parseInt(value, 10) + ');');
    }

    /* Real 環境設定をメインエンジンで設定 / Set a real preference on the main engine */
    function btSetRealPreference(prefKey, value) {
        runInMainEngine('app.preferences.setRealPreference("' + prefKey + '", ' + Number(value) + ');');
    }

    /* アートボード名・枠線カラー・幅をまとめて設定し、キャンバスを再描画 / Apply artboard name/border color/width together, then refresh the canvas */
    function btApplyArtboardBorder(showName, r, g, b, width) {
        var body = '' +
            'var p=app.preferences;' +
            'p.setBooleanPreference("showArtboardLabelOnCanvas",' + (showName ? 'true' : 'false') + ');' +
            'p.setRealPreference("ArtboardBBColorRed",' + Number(r) + ');' +
            'p.setRealPreference("ArtboardBBColorGreen",' + Number(g) + ');' +
            'p.setRealPreference("ArtboardBBColorBlue",' + Number(b) + ');' +
            'p.setRealPreference("ArtboardBBWidth",' + Number(width) + ');' +
            'try{app.executeMenuCommand("zoomout");app.executeMenuCommand("zoomin");}catch(e){}';
        runInMainEngine(body);
    }

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

    /* キー入力フィールドの値を保存 / Save the value from the key field */
    function saveCursorKeyLengthFromField(editText) {
        saveCursorKeyLengthInCurrentUnit(parseFloat(editText.text));
    }

    /* ===== デバウンス保存 / Debounced saving ===== */
    var __cursorKeyDebounceTaskId = null;
    var __cursorKeyPendingText = null;

    /* 保留中のテキストを実際に保存（scheduleTask から呼ばれる）/ Save the pending text (called from scheduleTask) */
    function __runSaveCursorKeyLength() {
        try {
            if (__cursorKeyPendingText !== null) {
                saveCursorKeyLengthInCurrentUnit(parseFloat(__cursorKeyPendingText));
            }
        } catch (e) {
            // no-op on failure
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
            saveCursorKeyLengthFromField(editText);
        }
    }

    /* ↑↓キーで値を増減（Shift=±10で10の倍数にスナップ / Option=±0.1 / 通常=±1）。負値は0でクランプ */
    /* Arrow keys adjust the value (Shift = ±10 snapped to multiples of 10, Option = ±0.1, otherwise ±1); clamps at 0 */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var k = event.keyName;
            if (k !== "Up" && k !== "Down") {
                return; /* 非矢印キーでは処理しない（Tabで0に丸められるのを防止）/ Ignore non-arrow keys (prevents Tab rounding to 0) */
            }
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 with Shift */
                if (k == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (k == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                /* Optionキー押下時は0.1単位で増減 / Step by 0.1 with Option */
                if (k == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (k == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (k == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (k == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10; /* 小数第1位までに丸め / Round to 1 decimal */
            } else {
                value = Math.round(value); /* 整数に丸め / Round to integer */
            }

            /* 常に小数第1位で表示し、デバウンス保存 / Always show 1 decimal and save (debounced) */
            editText.text = (Math.round(value * 10) / 10).toFixed(1);
            scheduleSaveCursorKeyLengthDebounced(editText, SAVE_DEBOUNCE_MS);
        });
    }

    // =========================================
    // UI ヘルパー / UI helpers
    // =========================================

    /* 複数チェックボックスを Boolean 環境設定キーにバインド / Bind multiple checkboxes to boolean preference keys */
    function bindCheckboxes(pairs) {
        for (var i = 0; i < pairs.length; i++) {
            (function (pair) {
                pair.checkbox.onClick = function () {
                    btSetBooleanPreference(pair.prefKey, pair.checkbox.value === true);
                };
            })(pairs[i]);
        }
    }

    /* 標準パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel) {
        panel.orientation = 'column';
        panel.alignChildren = ['left', 'top'];
        panel.margins = [8, 20, 8, 15];
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

        /* 初期表示用に現在値を読み込む / Load current values for the initial display */
        var initialPrefs = readAllPreferences();
        PREF_STATE.rulerType = parseInt(initialPrefs.rulerType, 10);
        if (isNaN(PREF_STATE.rulerType)) PREF_STATE.rulerType = 2;
        PREF_STATE.cursorKeyLengthPt = parseFloat(initialPrefs.cursorKeyLength);
        if (isNaN(PREF_STATE.cursorKeyLengthPt)) PREF_STATE.cursorKeyLengthPt = 1.0;

        var dialog = new Window('palette', L('dialog.title') + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];
        dialog.opacity = DIALOG_OPACITY;
        $.global.__aiQuickPrefsPalette = dialog;

        /* 閉じたら参照をクリア（次回は再構築）/ Clear the reference on close (rebuild next time) */
        dialog.onClose = function () {
            $.global.__aiQuickPrefsPalette = null;
        };

        var mainGroup = dialog.add('group');
        mainGroup.orientation = 'column';
        mainGroup.alignChildren = ['fill', 'top'];

        /* ===== 2カラム行（左右の列）/ Two-column row ===== */
        var columnsRow = mainGroup.add('group');
        columnsRow.orientation = 'row';
        columnsRow.alignChildren = ['fill', 'top'];

        var leftColumn = columnsRow.add('group');
        leftColumn.orientation = 'column';
        leftColumn.alignChildren = ['fill', 'top'];

        var rightColumn = columnsRow.add('group');
        rightColumn.orientation = 'column';
        rightColumn.alignChildren = ['fill', 'top'];

        /* ----- 左列：キー入力 / 変形と整列 / Left column: Key input / Transform & Align ----- */

        /* キー入力パネル（カーソル移動量）と単位ポップアップ / Key input panel (cursor step) with the unit popup */
        var keyInputPanel = leftColumn.add('panel', undefined, L('panel.keyInput'));
        keyInputPanel.orientation = 'row';
        keyInputPanel.alignChildren = ['left', 'center'];
        keyInputPanel.margins = [8, 20, 8, 15];

        var keyField = keyInputPanel.add('edittext', undefined, "1.0");
        keyField.characters = 4;

        var suppressUnitChange = false;
        var unitDropdown = keyInputPanel.add('dropdownlist', undefined, []);
        for (var u = 0; u < UNIT_POPUP_CODES.length; u++) {
            unitDropdown.add('item', getUnitLabel(UNIT_POPUP_CODES[u]));
        }
        unitDropdown.preferredSize.width = 70;

        /* 単位ポップアップ：選んだ単位を定規単位(rulerType)へ反映し、表示を再計算 / Unit popup: apply the chosen unit to rulerType and recompute the display */
        unitDropdown.onChange = function () {
            if (suppressUnitChange || !unitDropdown.selection) return;
            var code = UNIT_POPUP_CODES[unitDropdown.selection.index];
            PREF_STATE.rulerType = code;
            btSetIntegerPreference("rulerType", code);
            /* 保存済み pt 値は不変。新しい単位で再表示 / Stored pt value is unchanged; redisplay in the new unit */
            keyField.text = readCursorKeyLengthInCurrentUnit();
        };

        changeValueByArrowKey(keyField);

        /* 変形と整列パネル / Transform & Align panel */
        var transformPanel = leftColumn.add('panel', undefined, L('panel.transform'));
        setupPanel(transformPanel);

        /* プレビュー境界 / Preview bounds */
        var checkboxPreview = transformPanel.add('checkbox', undefined, L('checkbox.previewBounds'));
        checkboxPreview.onClick = function () {
            btSetBooleanPreference("includeStrokeInBounds", checkboxPreview.value === true);
        };

        /* パターンを変形 / Transform patterns */
        var checkboxPattern = transformPanel.add('checkbox', undefined, L('checkbox.transformPattern'));
        checkboxPattern.onClick = function () {
            btSetBooleanPreference("transformPatterns", checkboxPattern.value === true);
        };

        /* 角を拡大・縮小（1=ON, 2=OFF）/ Scale corners (1=ON, 2=OFF) */
        var checkboxCorner = transformPanel.add('checkbox', undefined, L('checkbox.scaleCorners'));
        checkboxCorner.onClick = function () {
            btSetIntegerPreference("policyForPreservingCorners", checkboxCorner.value ? 1 : 2);
        };

        /* 線幅と効果も拡大・縮小 / Scale strokes and effects */
        var checkboxStroke = transformPanel.add('checkbox', undefined, L('checkbox.scaleStroke'));
        checkboxStroke.onClick = function () {
            btSetBooleanPreference("scaleLineWeight", checkboxStroke.value === true);
        };

        /* ----- 右列：字形の境界に整列 / ガイドと定規 / Right column: Glyph bounds / Guides & Rulers ----- */

        /* 字形の境界に整列パネル / Align to glyph bounds panel */
        var glyphPanel = rightColumn.add('panel', undefined, L('panel.glyphBounds'));
        setupPanel(glyphPanel);

        var checkboxPoint = glyphPanel.add('checkbox', undefined, L('checkbox.pointText'));
        var checkboxArea = glyphPanel.add('checkbox', undefined, L('checkbox.areaText'));

        bindCheckboxes([
            { checkbox: checkboxPoint, prefKey: 'EnableActualPointTextSpaceAlign' },
            { checkbox: checkboxArea, prefKey: 'EnableActualAreaTextSpaceAlign' }
        ]);

        /* ガイドと定規パネル / Guides & Rulers panel */
        var guidePanel = rightColumn.add('panel', undefined, L('panel.guide'));
        setupPanel(guidePanel);

        /* ガイドを表示 / Show guides */
        var checkboxGuideShow = guidePanel.add('checkbox', undefined, L('checkbox.guideShow'));
        checkboxGuideShow.onClick = function () {
            btSetBooleanPreference("showGuides", checkboxGuideShow.value === true);
        };

        /* ガイドをロック / Lock guides */
        var checkboxGuideLock = guidePanel.add('checkbox', undefined, L('checkbox.guideLock'));
        checkboxGuideLock.onClick = function () {
            btSetBooleanPreference("lockGuides", checkboxGuideLock.value === true);
        };

        /* ビデオ定規（メニューコマンドのトグル）/ Video ruler (menu-command toggle) */
        var btnVideoRuler = guidePanel.add('button', undefined, L('button.videoRuler'));
        btnVideoRuler.alignment = ['left', 'top']; /* 幅いっぱいにしない（ラベル幅）/ Do not fill width (size to label) */
        btnVideoRuler.onClick = function () {
            runInMainEngine('try{app.executeMenuCommand("videoruler");}catch(e){}');
        };

        /* ----- 全幅：アートボード名と枠線 / その他（一番下）/ Full width: Artboard / Other (bottom) ----- */

        /* アートボード名と枠線パネル / Artboard name & border panel */
        var artboardPanel = mainGroup.add('panel', undefined, L('panel.artboard'));
        artboardPanel.orientation = 'column';
        artboardPanel.alignChildren = ['fill', 'top'];
        artboardPanel.margins = [8, 20, 8, 15];

        var suppressArtboardChange = false;

        /* アートボード名を表示 / Show artboard name */
        var cbShowArtboardName = artboardPanel.add('checkbox', undefined, L('checkbox.showArtboardName'));
        cbShowArtboardName.onClick = function () {
            applyArtboard();
        };

        /* カンバスカラーをホワイトに（ON=1, OFF=0）/ Canvas color white (ON=1, OFF=0) */
        var checkboxCanvasWhite = artboardPanel.add('checkbox', undefined, L('checkbox.canvasWhite'));
        checkboxCanvasWhite.onClick = function () {
            btSetIntegerPreference("uiCanvasIsWhite", checkboxCanvasWhite.value ? 1 : 0);
        };

        /* アートボードの枠線サブパネル / Artboard border sub-panel */
        var artboardBorderPanel = artboardPanel.add('panel', undefined, L('panel.artboardBorder'));
        artboardBorderPanel.orientation = 'column';
        artboardBorderPanel.alignChildren = ['left', 'top'];
        artboardBorderPanel.margins = [8, 20, 8, 15];

        /* ハイライトのカラー / Highlight color */
        var strokeColorRow = artboardBorderPanel.add('group');
        strokeColorRow.orientation = 'row';
        strokeColorRow.alignChildren = ['left', 'center'];
        strokeColorRow.add('statictext', undefined, labelText('label.strokeColor'));
        var ddStrokeColor = strokeColorRow.add('dropdownlist', undefined, buildStrokeColorNames());
        ddStrokeColor.onChange = function () {
            if (suppressArtboardChange || !ddStrokeColor.selection) return;
            applyArtboard();
        };

        /* ストロークの幅（1〜4）/ Stroke width (1-4) */
        var strokeWidthRow = artboardBorderPanel.add('group');
        strokeWidthRow.orientation = 'row';
        strokeWidthRow.alignChildren = ['left', 'center'];
        strokeWidthRow.add('statictext', undefined, labelText('label.strokeWidth'));
        var rbStrokeWidth1 = strokeWidthRow.add('radiobutton', undefined, '1');
        var rbStrokeWidth2 = strokeWidthRow.add('radiobutton', undefined, '2');
        var rbStrokeWidth3 = strokeWidthRow.add('radiobutton', undefined, '3');
        var rbStrokeWidth4 = strokeWidthRow.add('radiobutton', undefined, '4');
        var rbStrokeWidths = [rbStrokeWidth1, rbStrokeWidth2, rbStrokeWidth3, rbStrokeWidth4];
        for (var sw = 0; sw < rbStrokeWidths.length; sw++) {
            rbStrokeWidths[sw].onClick = function () {
                applyArtboard();
            };
        }

        /* 選択中のストローク幅（1〜4）を取得 / Get the selected stroke width (1-4) */
        function getSelectedStrokeWidth() {
            for (var i = 0; i < rbStrokeWidths.length; i++) {
                if (rbStrokeWidths[i].value) return i + 1;
            }
            return 1;
        }

        /* アートボード名・枠線の現在 UI 値をまとめて保存 / Save the current artboard name/border UI state */
        function applyArtboard() {
            var idx = ddStrokeColor.selection ? ddStrokeColor.selection.index : STROKE_COLOR_BLACK_INDEX;
            var c = STROKE_COLOR_PRESETS[idx];
            btApplyArtboardBorder(cbShowArtboardName.value === true, c.r, c.g, c.b, getSelectedStrokeWidth());
        }

        /* その他パネル（一番下）/ Other panel (bottom) */
        var etcPanel = mainGroup.add('panel', undefined, L('panel.etc'));
        setupPanel(etcPanel);

        /* リアルタイムの描画と編集 / Real-time drawing & editing */
        var checkboxRealtime = etcPanel.add('checkbox', undefined, L('checkbox.realtimeDrawing'));
        checkboxRealtime.onClick = function () {
            btSetBooleanPreference("LiveEdit_State_Machine", checkboxRealtime.value === true);
        };

        /* 書式なしペースト / Paste without formatting */
        var checkboxPastePlain = etcPanel.add('checkbox', undefined, L('checkbox.pastePlain'));
        checkboxPastePlain.onClick = function () {
            btSetBooleanPreference("plugin/FileClipboard/pasteWithoutFormatting", checkboxPastePlain.value === true);
        };

        /* 読み出した環境設定を UI へ反映 / Apply fetched preferences to the UI */
        function applyPreferencesToUI(map) {
            function asBool(prefKey) {
                return String(map[prefKey]) === "true";
            }
            checkboxPoint.value = asBool('EnableActualPointTextSpaceAlign');
            checkboxArea.value = asBool('EnableActualAreaTextSpaceAlign');
            checkboxPreview.value = asBool('includeStrokeInBounds');
            checkboxPattern.value = asBool('transformPatterns');
            checkboxStroke.value = asBool('scaleLineWeight');
            checkboxGuideShow.value = asBool('showGuides');
            checkboxGuideLock.value = asBool('lockGuides');
            checkboxRealtime.value = asBool('LiveEdit_State_Machine');
            checkboxPastePlain.value = asBool('pastePlain');
            checkboxCorner.value = (parseInt(map['policyForPreservingCorners'], 10) === 1);

            /* アートボード名と枠線 / Artboard name & border */
            cbShowArtboardName.value = asBool('showArtboardLabelOnCanvas');
            checkboxCanvasWhite.value = (parseInt(map['uiCanvasIsWhite'], 10) === 1);
            var cr = parseFloat(map['ArtboardBBColorRed']);
            var cg = parseFloat(map['ArtboardBBColorGreen']);
            var cb2 = parseFloat(map['ArtboardBBColorBlue']);
            if (isNaN(cr)) cr = 0;
            if (isNaN(cg)) cg = 0;
            if (isNaN(cb2)) cb2 = 0;
            suppressArtboardChange = true;
            ddStrokeColor.selection = findClosestStrokeColor(cr, cg, cb2);
            suppressArtboardChange = false;
            var w = Math.round(parseFloat(map['ArtboardBBWidth']));
            if (isNaN(w)) w = 1;
            if (w < 1) w = 1;
            if (w > 4) w = 4;
            for (var wi = 0; wi < rbStrokeWidths.length; wi++) {
                rbStrokeWidths[wi].value = (wi === w - 1);
            }

            var ruler = parseInt(map['rulerType'], 10);
            PREF_STATE.rulerType = isNaN(ruler) ? PREF_STATE.rulerType : ruler;
            var ckl = parseFloat(map['cursorKeyLength']);
            PREF_STATE.cursorKeyLengthPt = isNaN(ckl) ? PREF_STATE.cursorKeyLengthPt : ckl;

            /* 単位ポップアップを現在の rulerType に同期（onChange を発火させない）/ Sync the unit popup to the current rulerType (without firing onChange) */
            var unitIdx = unitCodeToPopupIndex(PREF_STATE.rulerType);
            if (unitIdx >= 0) {
                suppressUnitChange = true;
                unitDropdown.selection = unitIdx;
                suppressUnitChange = false;
            }
            keyField.text = readCursorKeyLengthInCurrentUnit();
        }

        /* キー入力の確定時に保存（不正値は元へ戻す）/ Save on commit (restore on invalid input) */
        keyField.onChange = function () {
            if (!saveCursorKeyLengthInCurrentUnit(parseFloat(keyField.text))) {
                keyField.text = readCursorKeyLengthInCurrentUnit();
            }
        };

        /* 構築直後に現在値を反映（常に最新の環境設定を表示）/ Populate from current values right after building */
        applyPreferencesToUI(initialPrefs);

        /* 表示時：フォーカスと最新値の再読込 / On show: focus and reload current values */
        dialog.onShow = function () {
            keyField.active = true;
            applyPreferencesToUI(readAllPreferences());
        };

        /* 再アクティブ時：外部変更（環境設定ダイアログ等）へクリックで追従 / On re-activate: follow external changes via click */
        dialog.onActivate = function () {
            applyPreferencesToUI(readAllPreferences());
        };

        dialog.show();
    }

    main();

}());
