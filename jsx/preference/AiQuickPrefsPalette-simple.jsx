#target illustrator
#targetengine "AiQuickPrefsPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustrator の各種環境設定の切り替えと、選択オブジェクトの反転・回転を、常駐パレットでまとめて操作するユーティリティです。操作した時点で即時反映されます。

- パレット（常駐エンジン）で表示し、書き込み・DOM 操作は BridgeTalk でメインエンジンへ委譲（読み出しは同期で直接取得）
- 上段は2カラム（左:キー入力／整列オプション／変形オプション・右:反転と回転／字形の境界に整列）、その下に全幅でコピー/ペースト・描画
- 反転・回転はアイコンボタンで実行し、9軸（3×3）ウィジェットで基点を指定。アイコンはライト／ダーク UI に合わせて配色を自動切り替え
- 環境設定ダイアログ等の外部変更は、パレットをクリック（再アクティブ）で同期
- パレットがアクティブなとき esc キーで閉じる

#### パネルと項目

- キー入力：カーソル移動量（cursorKeyLength）。単位ポップアップで定規単位を切替、↑↓ / Shift / Option で増減
- 整列オプション：プレビュー境界
- 字形の境界に整列：ポイント文字／エリア内文字
- 変形オプション：パターン／角／線幅と効果
- 変形：左右反転／上下反転／回転（反時計回り・時計回り）をアイコンボタンで実行。9軸（3×3）の基準点ウィジェットで反転・回転の基点を指定（既定は中央）。基点は選択全体の可視バウンディングを基準に算出。アイコンはライト／ダーク UI に合わせて配色を自動切り替え
- コピー/ペースト：書式なしペースト／コピー元のレイヤーにペースト
- 描画：リアルタイムの描画と編集／プレビュー更新（GPU プレビューを更新）


### 紹介記事（note）

https://note.com/dtp_tranist/n/n41d8dc1961be

### Overview

A persistent-palette utility for batch-toggling various Illustrator preferences and flipping/rotating the selection. Every action applies immediately when triggered.

- Runs in a persistent-engine palette; writes and DOM operations are delegated to the main engine via BridgeTalk (reads are fetched directly/synchronously)
- Top is a two-column row (left = Key input / Align Options / Transform Options, right = Flip & Rotate / Align to Glyph Bounds); below it, full width = Copy / Paste and Drawing
- Flip and rotate run from icon buttons, with a 9-axis (3x3) widget to set the pivot; icon colors adapt to the light / dark UI
- External changes (e.g. the Preferences dialog) sync when you click (re-activate) the palette
- Press Esc to close while the palette is active

#### Panels & options

- Key input: cursor step (cursorKeyLength); switch ruler unit via popup, adjust with Up/Down / Shift / Option
- Align Options: Preview Bounds
- Align to Glyph Bounds: Point Type / Area Type
- Transform Options: Pattern Tiles / Corners / Strokes & Effects
- Transform: Flip horizontal / vertical and rotate (counterclockwise / clockwise) from icon buttons. A 9-axis (3x3) anchor widget sets the pivot (center by default), computed from the selection's overall visible bounds. Icon colors adapt to the light / dark UI
- Copy / Paste: Paste without Formatting / Paste Remembers Layers
- Drawing: Real-time Drawing & Editing / Refresh Preview (GPU preview)

### 更新履歴 / Change Log

- v2.0.0 (20260630): 「反転と回転」パネルを FlipRotatePalette のアイコン UI へ差し替え（左右反転／上下反転／回転 CCW・CW のアイコンボタン＋9軸の基準点ウィジェット、ライト／ダーク対応）。反転・回転の基点を選択中心固定から 9 軸の任意基準点に変更（btTransformSelection＋getAnchorExpressions に共通化、適用後に app.redraw()）。方向ラジオと水平／垂直／45°回転のテキストボタンを廃止。
- v1.0 (20250804): 初期バージョン。

*/

// =========================================
// バージョン / Version
// =========================================

var SCRIPT_VERSION = "v2.0.0";

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DIALOG_OPACITY = 0.98;   /* パレットの不透明度 / Palette opacity */
    var SAVE_DEBOUNCE_MS = 40;   /* 保存デバウンス(ms) / Save debounce (ms) */
    var ROTATE_ANGLE = 45;       /* 回転アイコンの回転角（度）。正＝反時計回り / Rotate icon angle (deg); positive = CCW */
    var COLUMNS_PER_ROW = 2;     /* アイコンを何個ごとに改行するか / Icons per row before wrapping */

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
            keyInput: { ja: "キー入力", en: "Key Input" },
            align: { ja: "整列オプション", en: "Align Options" },
            transform: { ja: "変形オプション", en: "Transform Options" },
            flip: { ja: "反転と回転", en: "Flip & Rotate" },
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            copyPaste: { ja: "コピー/ペースト", en: "Copy / Paste" },
            drawing: { ja: "描画", en: "Drawing" }
        },
        checkbox: {
            pointText: { ja: "ポイント文字", en: "Point Type" },
            areaText: { ja: "エリア内文字", en: "Area Type" },
            previewBounds: { ja: "プレビュー境界", en: "Preview Bounds" },
            transformPattern: { ja: "パターン", en: "Pattern Tiles" },
            scaleCorners: { ja: "角", en: "Corners" },
            scaleStroke: { ja: "線幅と効果", en: "Strokes & Effects" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集", en: "Real-time Drawing & Editing" },
            pastePlain: { ja: "書式なしペースト", en: "Paste without Formatting" },
            pastePreserve: { ja: "コピー元のレイヤーにペースト", en: "Paste Remembers Layers" }
        },
        tooltip: {
            cursorStep: { ja: "矢印キーでの移動量（環境設定 > 一般 > キー入力）。↑↓ / Shift=±10 / Option=±0.1", en: "Keyboard increment (Preferences > General). Up/Down / Shift=±10 / Option=±0.1" },
            unit: { ja: "定規の単位を切り替え", en: "Switch the ruler unit" },
            previewBounds: { ja: "整列・分布でプレビュー境界（線幅・効果を含む）を使用", en: "Use preview bounds (incl. stroke & effects) for align/distribute" },
            pointText: { ja: "ポイント文字を字形の境界で整列", en: "Align point type to glyph bounds" },
            areaText: { ja: "エリア内文字を字形の境界で整列", en: "Align area type to glyph bounds" },
            transformPattern: { ja: "変形時にパターンも変形する", en: "Transform pattern tiles when transforming" },
            scaleCorners: { ja: "拡大・縮小時に角（ライブコーナー）の半径も拡大・縮小", en: "Scale corner radius when scaling" },
            scaleStroke: { ja: "拡大・縮小時に線幅と効果も拡大・縮小", en: "Scale strokes & effects when scaling" },
            flipHorizontal: { ja: "選択を水平方向に反転（基準点が基点）", en: "Flip the selection horizontally (about the anchor point)" },
            flipVertical: { ja: "選択を垂直方向に反転（基準点が基点）", en: "Flip the selection vertically (about the anchor point)" },
            rotateCW: { ja: "選択を時計回りに45°回転（基準点が基点）", en: "Rotate the selection 45° clockwise (about the anchor point)" },
            rotateCCW: { ja: "選択を反時計回りに45°回転（基準点が基点）", en: "Rotate the selection 45° counterclockwise (about the anchor point)" },
            anchor: { ja: "基準点（反転・回転の基点）", en: "Anchor point (pivot for flip / rotate)" },
            pastePlain: { ja: "書式を保持せずにペースト", en: "Paste without keeping formatting" },
            pastePreserve: { ja: "コピー元と同じレイヤーにペースト", en: "Paste into the original (source) layer" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集を切り替え", en: "Toggle real-time drawing & editing" },
            refreshGpuPreview: { ja: "GPUプレビューを更新（再描画）", en: "Refresh the GPU preview (redraw)" }
        },
        button: {
            refreshGpuPreview: { ja: "プレビュー更新", en: "Refresh Preview" }
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
        var text = node[currentLanguage] || node.en || "";
        return text.replace(/\{slash\}/g, "/");
    }

    /* 現在言語のラベル文字列を返す / Return the current-language label string */
    function L(key) {
        return getLabel(key);
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
                try { $.writeln("AiQuickPrefsPalette BridgeTalk error: " + message.body); } catch (e) {}
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

    /* 選択した 9 軸の基準点に対応する anchorX/anchorY の式（body 内の left/right/top/bottom を参照）を返す */
    /* Return the anchorX/anchorY expressions for the selected 9-axis point (referencing left/right/top/bottom in the body) */
    function getAnchorExpressions() {
        var col = selectedAnchorIndex % 3;
        var row = Math.floor(selectedAnchorIndex / 3);
        var xExpr = (col === 0) ? "left" : (col === 1) ? "((left+right)/2)" : "right";
        var yExpr = (row === 0) ? "top" : (row === 1) ? "((top+bottom)/2)" : "bottom";
        return "var anchorX=" + xExpr + ",anchorY=" + yExpr + ";";
    }

    /* 共通の委譲本体：可視バウンディング測定→基準点→合成行列を1回適用→再描画。matrixCode が matrix を組み立てる */
    /* Shared delegation body: measure visible bounds -> anchor -> apply one matrix -> redraw; matrixCode builds the matrix */
    /* visibleBounds が取れない/変形できない項目（ロック・非表示・ガイド等）は try/catch でスキップ。測定できた項目が無ければ中断 */
    /* Skip items whose bounds/transform fail (locked, hidden, guides, ...) via try/catch; abort if nothing could be measured */
    function btTransformSelection(matrixCode) {
        var body = '' +
            'if(app.documents.length>0){' +
            'var doc=app.activeDocument,selection=doc.selection;' +
            'if(selection&&selection.length>0){' +
            'var left=Infinity,top=-Infinity,right=-Infinity,bottom=Infinity,measured=false;' +
            'for(var i=0;i<selection.length;i++){try{var b=selection[i].visibleBounds;if(b[0]<left)left=b[0];if(b[1]>top)top=b[1];if(b[2]>right)right=b[2];if(b[3]<bottom)bottom=b[3];measured=true;}catch(e){}}' +
            'if(measured){' +
            getAnchorExpressions() +
            'var matrix=app.getIdentityMatrix();' + matrixCode +
            'for(var i=0;i<selection.length;i++){try{selection[i].transform(matrix,true,true,true,true,1,Transformation.DOCUMENTORIGIN);}catch(e){}}' +
            'app.redraw();' +
            '}}}';
        runInMainEngine(body);
    }

    /* 選択を、選択した 9 軸基準点を基点に反転（水平=-100,100／垂直=100,-100）。DOM 操作なのでメインエンジンへ委譲 */
    /* Flip the selection about the selected 9-axis anchor (H=-100,100 / V=100,-100); delegated to the main engine since it touches the DOM */
    /* 係数は数値化して埋め込む（'1-' + (-1) だと "1--1" になりデクリメント解釈で構文エラーになるため）*/
    /* Coefficients are precomputed as numbers ('1-' + (-1) would yield "1--1", a decrement → syntax error) */
    function btFlipSelection(scaleX, scaleY) {
        var scaleFractionX = Number(scaleX) / 100; /* -1 or 1 */
        var scaleFractionY = Number(scaleY) / 100;
        btTransformSelection(
            'matrix.mValueA=' + scaleFractionX + ';matrix.mValueD=' + scaleFractionY + ';' +
            'matrix.mValueTX=anchorX*' + (1 - scaleFractionX) + ';matrix.mValueTY=anchorY*' + (1 - scaleFractionY) + ';'
        );
    }

    /* 選択を、選択した 9 軸基準点を基点に回転（angleDegrees：正＝反時計回り／負＝時計回り）。DOM 操作なのでメインエンジンへ委譲 */
    /* Rotate the selection about the selected 9-axis anchor (angleDegrees: positive = CCW / negative = CW); delegated to the main engine since it touches the DOM */
    function btRotateSelection(angleDegrees) {
        var radians = Number(angleDegrees) * Math.PI / 180;
        var cosAngle = Math.cos(radians);
        var sinAngle = Math.sin(radians);
        var oneMinusCos = 1 - cosAngle;   /* 係数は数値化（"1--0.7" のような構文エラーを避ける）/ Precompute to avoid "1--0.7"-style syntax errors */
        btTransformSelection(
            'matrix.mValueA=' + cosAngle + ';matrix.mValueB=' + sinAngle + ';matrix.mValueC=' + (-sinAngle) + ';matrix.mValueD=' + cosAngle + ';' +
            'matrix.mValueTX=anchorX*' + oneMinusCos + '+anchorY*' + sinAngle + ';' +
            'matrix.mValueTY=anchorY*' + oneMinusCos + '-anchorX*' + sinAngle + ';'
        );
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
    var COLUMN_SPACING = 12; /* 2カラムの間隔 / Gap between the two columns */

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
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) {}
    }

    // =========================================
    // 反転・回転（アイコン UI）/ Flip & Rotate (icon UI)
    // =========================================

    /* 基準点（9軸）: 0..8 を行優先（0=左上, 4=中央, 8=右下）。反転・回転の基点に使う / 9-axis anchor 0..8 row-major (0=top-left, 4=center, 8=bottom-right); used as the flip/rotate pivot */
    var selectedAnchorIndex = 4;

    /* 4つのアイコンボタン定義（左右反転／上下反転／回転CCW／回転CW）/ The four icon buttons (flip H / flip V / rotate CCW / rotate CW) */
    var iconButtonDefs = [
        { name: "FLIP_HORIZONTAL", icon: "flipHorizontal", tooltip: "tooltip.flipHorizontal" },
        { name: "FLIP_VERTICAL",   icon: "flipVertical",   tooltip: "tooltip.flipVertical" },
        { name: "ROTATE",          icon: "rotate",         tooltip: "tooltip.rotateCCW" },
        { name: "ROTATE_FLIP",     icon: "rotateFlip",     tooltip: "tooltip.rotateCW" }
    ];

    /* アイコンの配色（main() で UI 明暗から設定）/ Icon colors (set in main() from the light/dark UI) */
    var iconColor, iconBorderColor, iconBaseBg, iconHoverBg;

    /* UI 明度(0..1)を取得、取得失敗時は 1（明るい）を返す / Get UI brightness (0..1); 1 on failure */
    function getUIBrightness() {
        try {
            var brightness = app.preferences.getRealPreference("uiBrightness");
            if (brightness < 0) brightness = 0;
            if (brightness > 1) brightness = 1;
            return brightness;
        } catch (e) {
            return 1;
        }
    }

    /* UI 明度 > 0.5 ならライト、取得失敗時はダーク扱い / Light if uiBrightness > 0.5; dark on failure */
    function isLightUI() {
        try {
            return app.preferences.getRealPreference("uiBrightness") > 0.5;
        } catch (e) {
            return false;
        }
    }

    /* グレーの RGBA を返す（0..1 にクランプ）/ Build a clamped gray RGBA */
    function grayColor(value) {
        if (value < 0) value = 0;
        if (value > 1) value = 1;
        return [value, value, value, 1];
    }

    /* UI の明暗に合わせてアイコン色・背景色・枠線を決める（main() から呼ぶ）/ Decide icon, background and border colors from the light/dark UI (called from main()) */
    function initIconColors() {
        var lightUI = isLightUI();
        var uiBrightness = getUIBrightness();
        iconColor       = lightUI ? [0.25, 0.25, 0.25, 1] : [0.85, 0.85, 0.85, 1];
        /* ライトモードは背景を塗らず（ウィンドウ色）薄いグレーの枠を付ける。ダークは枠なし / Light: no fill (window color) + light gray border; dark: no border */
        iconBorderColor = lightUI ? [0.65, 0.65, 0.65, 1] : null;
        iconBaseBg      = lightUI ? grayColor(uiBrightness)        : [0.28, 0.28, 0.28, 1];
        /* マウスオーバー時の背景（ライトは少し暗く、ダークは少し明るく）/ Hover background (slightly darker in light, lighter in dark) */
        iconHoverBg     = lightUI ? grayColor(uiBrightness - 0.10) : [0.38, 0.38, 0.38, 1];
    }

    /* コントロールを再描画（notify は環境により例外を投げ得るので保護）/ Redraw a control (notify can throw in some environments) */
    function redrawControl(control) {
        try { control.notify("onDraw"); } catch (e) {}
    }

    /* アイコンに対応する変形を実行する / Run the transform for an icon */
    function handleIconAction(name) {
        if (name === "FLIP_HORIZONTAL") {
            btFlipSelection(-100, 100);
        } else if (name === "FLIP_VERTICAL") {
            btFlipSelection(100, -100);
        } else if (name === "ROTATE") {
            btRotateSelection(ROTATE_ANGLE);     /* 反時計回り / counterclockwise */
        } else if (name === "ROTATE_FLIP") {
            btRotateSelection(-ROTATE_ANGLE);    /* 時計回り / clockwise */
        }
    }

    /* アイコンボタンを1つ生成して登録 / Create and register one icon button */
    function addIconButton(parentGroup, buttonDef) {
        var button = parentGroup.add("button", undefined, "");
        button.helpTip = L(buttonDef.tooltip);
        button.preferredSize = [26, 26];
        button.minimumSize = [26, 26];
        button.maximumSize = [26, 26];
        button.iconType = buttonDef.icon;
        button.isHover = false;
        button.onDraw = function () {
            drawIcon(this);
        };
        button.onClick = function () {
            handleIconAction(buttonDef.name);
        };
        /* マウスオーバーで背景を付けるためホバー状態を切り替える / Toggle hover to show a background on mouseover */
        try {
            button.addEventListener("mouseover", function () { button.isHover = true; redrawControl(button); });
            button.addEventListener("mouseout", function () { button.isHover = false; redrawControl(button); });
        } catch (e) {}
    }

    /* 9軸（3×3）の基準点ウィジェットを生成 / Create the 9-axis (3x3) anchor widget */
    function addAnchorWidget(parentGroup) {
        var widget = parentGroup.add("button", undefined, "");
        widget.helpTip = L('tooltip.anchor');
        widget.preferredSize = [44, 44];
        widget.minimumSize = [44, 44];
        widget.maximumSize = [44, 44];
        widget.onDraw = function () {
            drawAnchorWidget(this);
        };
        widget.onClick = function () {
            /* クリック位置の判定は mousedown 側で行う（onClick は何もしない）/ Cell hit-testing happens in mousedown; onClick is a no-op */
        };
        /* クリックした 3×3 のセルを基準点に設定（クリック座標はコントロール基準）/ Set the anchor from the clicked 3x3 cell (coords are control-relative) */
        try {
            widget.addEventListener("mousedown", function (event) {
                var cellWidth = widget.size[0] / 3;
                var cellHeight = widget.size[1] / 3;
                var col = Math.floor(event.clientX / cellWidth);
                var row = Math.floor(event.clientY / cellHeight);
                if (col < 0) col = 0;
                if (col > 2) col = 2;
                if (row < 0) row = 0;
                if (row > 2) row = 2;
                selectedAnchorIndex = row * 3 + col;
                redrawControl(widget);
            });
        } catch (e) {}
        return widget;
    }

    /* 9軸ウィジェットを描画（外周の□をケイ線でつなぐ・中央は独立）/ Draw the 9-axis widget (outer squares joined by rules; center stands alone) */
    function drawAnchorWidget(widget) {
        var graphics = widget.graphics;
        var width = widget.size[0];
        var height = widget.size[1];

        try {
            /* 9軸ウィジェットは枠線なし（背景のみ、ウィンドウ色に馴染ませる）/ No border (background only, blending into the window color) */
            graphics.rectPath(0, 0, width, height);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, iconBaseBg));
        } catch (e0) {}

        var cellSize = 6;   /* 四角のサイズ / Square size */
        var cellGap = 5;    /* 四角どうしの間隔 / Gap between squares */
        var cellStep = cellSize + cellGap;
        var gridSize = cellSize * 3 + cellGap * 2;
        var originX = Math.round((width - gridSize) / 2);
        var originY = Math.round((height - gridSize) / 2);

        function anchorCellX(index) { return originX + (index % 3) * cellStep; }
        function anchorCellY(index) { return originY + Math.floor(index / 3) * cellStep; }

        /* 中央(4)を除く外周の□どうしをケイ線（隙間部分）でつなぐ / Join the outer squares (except center 4) with rules across the gaps */
        var connections = [[0, 1], [1, 2], [6, 7], [7, 8], [0, 3], [3, 6], [2, 5], [5, 8]];
        var linePen = graphics.newPen(graphics.PenType.SOLID_COLOR, iconColor, 1);
        for (var i = 0; i < connections.length; i++) {
            var cellA = connections[i][0];
            var cellB = connections[i][1];
            var cellAX = anchorCellX(cellA);
            var cellAY = anchorCellY(cellA);
            var cellBX = anchorCellX(cellB);
            var cellBY = anchorCellY(cellB);
            graphics.newPath();
            if (cellB - cellA === 1) {
                /* 横方向：右隣の□へ / Horizontal: to the square on the right */
                graphics.moveTo(cellAX + cellSize, cellAY + cellSize / 2);
                graphics.lineTo(cellBX, cellBY + cellSize / 2);
            } else {
                /* 縦方向：下の□へ / Vertical: to the square below */
                graphics.moveTo(cellAX + cellSize / 2, cellAY + cellSize);
                graphics.lineTo(cellBX + cellSize / 2, cellBY);
            }
            graphics.strokePath(linePen);
        }

        for (var index = 0; index < 9; index++) {
            drawAnchorCell(graphics, anchorCellX(index), anchorCellY(index), cellSize, index === selectedAnchorIndex);
        }
    }

    /* 基準点セルの□を1つ描画（選択中は塗り、非選択は枠線）/ Draw one anchor-cell square (filled if selected, outlined otherwise) */
    function drawAnchorCell(graphics, x, y, size, selected) {
        graphics.newPath();
        graphics.moveTo(x, y);
        graphics.lineTo(x + size, y);
        graphics.lineTo(x + size, y + size);
        graphics.lineTo(x, y + size);
        graphics.closePath();
        if (selected) {
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, iconColor));
        } else {
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, iconColor, 1));
        }
    }

    /* アイコンボタンの背景・枠線を描き、種類に応じた図柄を描画 / Draw the icon button background/border and dispatch to the right glyph */
    function drawIcon(button) {
        var graphics = button.graphics;
        var width = button.size[0];
        var height = button.size[1];
        var backgroundColor = (button.isHover === true) ? iconHoverBg : iconBaseBg;

        try {
            graphics.rectPath(0, 0, width, height);
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, backgroundColor));

            /* ライトモードは薄いグレーの枠線を描く / Light mode draws a light gray border */
            if (iconBorderColor) {
                graphics.rectPath(0.5, 0.5, width - 1, height - 1);
                graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, iconBorderColor, 1));
            }
        } catch (e1) {
            try {
                graphics.drawOSControl();
            } catch (e2) {}
        }

        if (button.iconType === "rotate") {
            drawRotateIcon(graphics, width, height, false);
        } else if (button.iconType === "rotateFlip") {
            drawRotateIcon(graphics, width, height, true);
        } else {
            drawFlipIcon(graphics, button.iconType, width, height);
        }
    }

    /* 反転アイコン（軸の点線＋向かい合う三角形）を描画 / Draw a flip icon (dotted axis + opposing triangles) */
    function drawFlipIcon(graphics, iconType, width, height) {
        var color = iconColor;
        var centerX = width / 2;
        var centerY = height / 2;

        if (iconType === "flipVertical") {
            /* 横の点線を軸に、上向き／下向きの三角形を配置する / Up/down triangles about a horizontal dotted axis */
            drawDottedLine(graphics, 5, centerY, width - 5, centerY, color);
            drawTriangle(graphics, [[centerX - 5, 4], [centerX + 5, 4], [centerX, centerY - 2]], color, true);
            drawTriangle(graphics, [[centerX - 5, height - 4], [centerX + 5, height - 4], [centerX, centerY + 2]], color, false);
        } else {
            /* 縦の点線を軸に、左向き／右向きの三角形を配置する / Left/right triangles about a vertical dotted axis */
            drawDottedLine(graphics, centerX, 5, centerX, height - 5, color);
            drawTriangle(graphics, [[4, centerY - 5], [4, centerY + 5], [centerX - 2, centerY]], color, true);
            drawTriangle(graphics, [[width - 4, centerY - 5], [width - 4, centerY + 5], [centerX + 2, centerY]], color, false);
        }
    }

    /* 回転アイコン（一定幅の実線弧＋点線弧＋矢じり）を描画。mirror で左右反転 / Draw a rotate icon (uniform-width solid arc + dotted arc + arrowhead); mirror flips it horizontally */
    function drawRotateIcon(graphics, width, height, mirror) {
        var color = iconColor;
        var centerX = width / 2;
        var centerY = height / 2 + 1;
        var radius = 7.5;
        var mirrorSign = mirror ? -1 : 1;   /* 左右反転のときは x をミラーする / Mirror x when flipped */
        var headDeg = 232;                  /* 矢じりの位置（左上）/ Arrowhead position (top-left) */

        /* 実線の弧は一定幅・単一パスでなめらかに描く / Solid arc: uniform width, single path for smoothness */
        strokeArc(graphics, color, centerX, centerY, radius, headDeg, 410, mirrorSign, 1.8, 1.8);
        /* 下側は四角い点線 / Square-dotted arc on the lower side */
        drawDottedArc(graphics, color, centerX, centerY, radius, 50, 150, mirrorSign, 1.8);

        /* 左上（反転時は右上）に大きめの矢じりを付ける / Add a larger arrowhead top-left (top-right when mirrored) */
        var headRad = headDeg * Math.PI / 180;
        var headX = centerX + radius * Math.cos(headRad);
        var headY = centerY + radius * Math.sin(headRad);
        var tangentX = Math.sin(headRad);   /* 反時計回り（角度が減る向き）の接線 / Tangent for the CCW (decreasing angle) direction */
        var tangentY = -Math.cos(headRad);
        var perpX = -tangentY;
        var perpY = tangentX;
        var tipForward = 4;       /* 矢じり先端の前方への張り出し / Arrowhead tip extent (forward) */
        var tipBack = 2;          /* 矢じり後方への張り出し / Arrowhead extent (backward) */
        var tipHalfWidth = 4.5;   /* 矢じりの片側の幅 / Arrowhead half width */

        var arrowPoints = [
            [headX + tangentX * tipForward, headY + tangentY * tipForward],
            [headX - tangentX * tipBack + perpX * tipHalfWidth, headY - tangentY * tipBack + perpY * tipHalfWidth],
            [headX - tangentX * tipBack - perpX * tipHalfWidth, headY - tangentY * tipBack - perpY * tipHalfWidth]
        ];

        if (mirror) {
            for (var i = 0; i < arrowPoints.length; i++) {
                arrowPoints[i][0] = 2 * centerX - arrowPoints[i][0];
            }
        }

        drawTriangle(graphics, arrowPoints, color, true);
    }

    /* 円弧を線分で近似して描く。線幅一定なら継ぎ目の出ない単一パス、変化させる場合のみセグメントごとに描く（先細り）/
       Stroke an arc as line segments: a single seamless path when the width is constant, per-segment only when it tapers */
    function strokeArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign, startWidth, endWidth) {
        var segments = Math.max(8, Math.round(Math.abs(endDeg - startDeg) / 5));

        function arcX(ratio) { return centerX + mirrorSign * radius * Math.cos((startDeg + (endDeg - startDeg) * ratio) * Math.PI / 180); }
        function arcY(ratio) { return centerY + radius * Math.sin((startDeg + (endDeg - startDeg) * ratio) * Math.PI / 180); }

        /* 幅一定：1本の連続パスで描くのでガタつかない / Constant width: one continuous path, so no jaggedness */
        if (startWidth === endWidth) {
            graphics.newPath();
            graphics.moveTo(arcX(0), arcY(0));
            for (var i = 1; i <= segments; i++) {
                graphics.lineTo(arcX(i / segments), arcY(i / segments));
            }
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, startWidth));
            return;
        }

        /* 幅可変：セグメントごとに線幅を補間 / Variable width: interpolate per segment */
        for (var j = 1; j <= segments; j++) {
            var ratio = j / segments;
            var lineWidth = startWidth + (endWidth - startWidth) * ratio;
            graphics.newPath();
            graphics.moveTo(arcX((j - 1) / segments), arcY((j - 1) / segments));
            graphics.lineTo(arcX(ratio), arcY(ratio));
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, lineWidth));
        }
    }

    /* 円弧に沿って四角い点線を描く / Draw a square-dotted line along an arc */
    function drawDottedArc(graphics, color, centerX, centerY, radius, startDeg, endDeg, mirrorSign, dotWidth) {
        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, dotWidth);
        var stepDeg = 13;
        var dashHalf = 0.9;

        for (var deg = startDeg; deg <= endDeg; deg += stepDeg) {
            var rad = deg * Math.PI / 180;
            var x = centerX + mirrorSign * radius * Math.cos(rad);
            var y = centerY + radius * Math.sin(rad);
            var tangentX = mirrorSign * Math.sin(rad);
            var tangentY = -Math.cos(rad);

            graphics.newPath();
            graphics.moveTo(x - tangentX * dashHalf, y - tangentY * dashHalf);
            graphics.lineTo(x + tangentX * dashHalf, y + tangentY * dashHalf);
            graphics.strokePath(pen);
        }
    }

    /* 3点の三角形を塗り or 線で描く / Draw a triangle from 3 points, filled or stroked */
    function drawTriangle(graphics, points, color, fill) {
        graphics.newPath();
        graphics.moveTo(points[0][0], points[0][1]);
        graphics.lineTo(points[1][0], points[1][1]);
        graphics.lineTo(points[2][0], points[2][1]);
        graphics.closePath();
        if (fill) {
            graphics.fillPath(graphics.newBrush(graphics.BrushType.SOLID_COLOR, color));
        } else {
            graphics.strokePath(graphics.newPen(graphics.PenType.SOLID_COLOR, color, 1.2));
        }
    }

    /* 水平または垂直の点線を描く / Draw a horizontal or vertical dotted line */
    function drawDottedLine(graphics, x1, y1, x2, y2, color) {
        var pen = graphics.newPen(graphics.PenType.SOLID_COLOR, color, 1);
        var isHorizontal = (y1 === y2);
        var totalLength = isHorizontal ? (x2 - x1) : (y2 - y1);
        var dashStep = 3;

        for (var pos = 0; pos < totalLength; pos += dashStep) {
            graphics.newPath();
            if (isHorizontal) {
                graphics.moveTo(x1 + pos, y1);
                graphics.lineTo(Math.min(x1 + pos + 1.5, x2), y1);
            } else {
                graphics.moveTo(x1, y1 + pos);
                graphics.lineTo(x1, Math.min(y1 + pos + 1.5, y2));
            }
            graphics.strokePath(pen);
        }
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

        /* UI の明暗からアイコンの配色を決定 / Decide icon colors from the light/dark UI */
        initIconColors();

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

        /* ===== 2カラム行（左右の列）/ Two-column row ===== */
        var columnsRow = mainGroup.add('group');
        columnsRow.orientation = 'row';
        columnsRow.alignChildren = ['fill', 'top'];
        columnsRow.spacing = COLUMN_SPACING;

        var leftColumn = columnsRow.add('group');
        leftColumn.orientation = 'column';
        leftColumn.alignChildren = ['fill', 'top'];

        var rightColumn = columnsRow.add('group');
        rightColumn.orientation = 'column';
        rightColumn.alignChildren = ['fill', 'top'];

        /* ----- 左列：キー入力 / 整列オプション / 変形オプション / Left column: Key input / Align Options / Transform Options ----- */

        /* キー入力パネル（カーソル移動量）と単位ポップアップ / Key input panel (cursor step) with the unit popup */
        var keyInputPanel = leftColumn.add('panel', undefined, L('panel.keyInput'));
        keyInputPanel.orientation = 'row';
        keyInputPanel.alignChildren = ['left', 'center'];
        keyInputPanel.margins = PANEL_MARGINS;

        var cursorStepField = keyInputPanel.add('edittext', undefined, "1.0");
        cursorStepField.characters = 4;
        cursorStepField.helpTip = L('tooltip.cursorStep');

        /* 編集中フラグ：キー入力欄にフォーカスがある間は外部同期で値を上書きしない / Editing flag: don't let external sync overwrite while the field has focus */
        var isEditingCursorStep = false;
        cursorStepField.onActivate = function () { isEditingCursorStep = true; };
        cursorStepField.onDeactivate = function () { isEditingCursorStep = false; };

        var suppressUnitChange = false;
        var unitDropdown = keyInputPanel.add('dropdownlist', undefined, []);
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

        /* 整列パネル（左列・キー入力の下）/ Align panel (left column, below Key input) */
        var alignPanel = leftColumn.add('panel', undefined, L('panel.align'));
        setupPanel(alignPanel);

        /* プレビュー境界 / Preview bounds */
        var checkboxPreview = addBooleanCheckbox(alignPanel, 'checkbox.previewBounds', 'includeStrokeInBounds', 'tooltip.previewBounds');

        /* ----- 右列：反転と回転 / 字形の境界に整列 / Right column: Flip & Rotate / Align to Glyph Bounds ----- */

        /* 反転と回転パネル（右列・字形の境界に整列の上）。アイコンボタン（2×2）＋9軸の基準点ウィジェット / Flip & Rotate panel (right column, above Align to Glyph Bounds); icon buttons (2x2) + 9-axis anchor widget */
        var flipPanel = rightColumn.add('panel', undefined, L('panel.flip'));
        setupPanel(flipPanel);
        flipPanel.alignChildren = ['center', 'top'];

        /* アイコンボタンを COLUMNS_PER_ROW 個ごとに改行して並べる / Lay out icon buttons, wrapping every COLUMNS_PER_ROW */
        var iconGrid = flipPanel.add('group');
        iconGrid.orientation = 'column';
        iconGrid.alignChildren = ['center', 'center'];
        iconGrid.spacing = 7;

        var iconRow = null;
        for (var iconIndex = 0; iconIndex < iconButtonDefs.length; iconIndex++) {
            if ((iconIndex % COLUMNS_PER_ROW) === 0) {
                iconRow = iconGrid.add('group');
                iconRow.orientation = 'row';
                iconRow.alignChildren = ['center', 'center'];
                iconRow.spacing = 7;
            }
            addIconButton(iconRow, iconButtonDefs[iconIndex]);
        }

        /* 9軸（3×3）の基準点ウィジェット（反転・回転の基点を指定）/ 9-axis (3x3) anchor widget (sets the flip/rotate pivot) */
        var anchorRow = flipPanel.add('group');
        anchorRow.orientation = 'row';
        anchorRow.alignChildren = ['center', 'center'];
        anchorRow.margins = [0, 6, 0, 0];
        addAnchorWidget(anchorRow);

        /* 字形の境界に整列パネル（右列・反転と回転の下）/ Align to glyph bounds panel (right column, below Flip & Rotate) */
        var glyphBoundsPanel = rightColumn.add('panel', undefined, L('panel.glyphBounds'));
        setupPanel(glyphBoundsPanel);

        /* 字形の境界に整列の2チェック（ポイント文字／エリア内文字）。Option+クリックで2つをまとめてON/OFF */
        /* The two Align-to-Glyph-Bounds checks (Point Type / Area Type); Option+click toggles both at once */
        var checkboxPoint = glyphBoundsPanel.add('checkbox', undefined, L('checkbox.pointText'));
        checkboxPoint.helpTip = L('tooltip.pointText');
        var checkboxArea = glyphBoundsPanel.add('checkbox', undefined, L('checkbox.areaText'));
        checkboxArea.helpTip = L('tooltip.areaText');

        /* 各チェックの現在値を環境設定へ反映 / Apply each checkbox value to its preference */
        function applyGlyphPoint() { btSetBooleanPreference("EnableActualPointTextSpaceAlign", checkboxPoint.value === true); }
        function applyGlyphArea() { btSetBooleanPreference("EnableActualAreaTextSpaceAlign", checkboxArea.value === true); }

        /* クリック時：Option 併用なら2つを同じ値に揃えてまとめて適用、通常は単独適用 / On click: with Option, set both to the same value and apply together; otherwise apply just this one */
        function onGlyphBoundsClick(clicked) {
            if (ScriptUI.environment.keyboardState.altKey) {
                var newValue = clicked.value === true;
                checkboxPoint.value = newValue;
                checkboxArea.value = newValue;
                applyGlyphPoint();
                applyGlyphArea();
                return;
            }
            if (clicked === checkboxPoint) applyGlyphPoint();
            else applyGlyphArea();
        }

        checkboxPoint.onClick = function () { onGlyphBoundsClick(checkboxPoint); };
        checkboxArea.onClick = function () { onGlyphBoundsClick(checkboxArea); };

        /* 変形オプションパネル（左列・整列オプションの下）/ Transform Options panel (left column, below Align Options) */
        var transformPanel = leftColumn.add('panel', undefined, L('panel.transform'));
        setupPanel(transformPanel);

        /* 変形オプションの3チェック（パターン／角／線幅と効果）。Option+クリックで3つをまとめてON/OFF */
        /* The three Transform Options (Pattern / Corners / Strokes & Effects); Option+click toggles all three at once */

        /* パターンを変形 / Transform patterns */
        var checkboxPattern = transformPanel.add('checkbox', undefined, L('checkbox.transformPattern'));
        checkboxPattern.helpTip = L('tooltip.transformPattern');

        /* 角を拡大・縮小（1=ON, 2=OFF。Boolean でなく整数）/ Scale corners (1=ON, 2=OFF; integer pref) */
        var checkboxCorner = transformPanel.add('checkbox', undefined, L('checkbox.scaleCorners'));
        checkboxCorner.helpTip = L('tooltip.scaleCorners');

        /* 線幅と効果も拡大・縮小 / Scale strokes and effects */
        var checkboxStroke = transformPanel.add('checkbox', undefined, L('checkbox.scaleStroke'));
        checkboxStroke.helpTip = L('tooltip.scaleStroke');

        /* 各チェックの現在値を環境設定へ反映（角は 1/2 の整数）/ Apply each checkbox value to its preference (corners is the integer 1/2) */
        function applyTransformPattern() { btSetBooleanPreference("transformPatterns", checkboxPattern.value === true); }
        function applyTransformCorner() { btSetIntegerPreference("policyForPreservingCorners", checkboxCorner.value ? 1 : 2); }
        function applyTransformStroke() { btSetBooleanPreference("scaleLineWeight", checkboxStroke.value === true); }

        /* クリック時：Option 併用なら3つを同じ値に揃えてまとめて適用、通常は単独適用 / On click: with Option, set all three to the same value and apply together; otherwise apply just this one */
        function onTransformOptionClick(clicked) {
            if (ScriptUI.environment.keyboardState.altKey) {
                var newValue = clicked.value === true;
                checkboxPattern.value = newValue;
                checkboxCorner.value = newValue;
                checkboxStroke.value = newValue;
                applyTransformPattern();
                applyTransformCorner();
                applyTransformStroke();
                return;
            }
            if (clicked === checkboxPattern) applyTransformPattern();
            else if (clicked === checkboxCorner) applyTransformCorner();
            else applyTransformStroke();
        }

        checkboxPattern.onClick = function () { onTransformOptionClick(checkboxPattern); };
        checkboxCorner.onClick = function () { onTransformOptionClick(checkboxCorner); };
        checkboxStroke.onClick = function () { onTransformOptionClick(checkboxStroke); };

        /* ----- 全幅：コピー/ペースト / 描画 / Full width: Copy / Paste / Drawing ----- */

        /* コピー/ペーストパネル / Copy / Paste panel */
        var copyPastePanel = mainGroup.add('panel', undefined, L('panel.copyPaste'));
        setupPanel(copyPastePanel);

        /* 書式なしペースト / Paste without formatting */
        var checkboxPastePlain = addBooleanCheckbox(copyPastePanel, 'checkbox.pastePlain', 'plugin/FileClipboard/pasteWithoutFormatting', 'tooltip.pastePlain');

        /* コピー元のレイヤーにペースト / Paste remembers layers */
        var checkboxPastePreserve = addBooleanCheckbox(copyPastePanel, 'checkbox.pastePreserve', 'layers/pastePreserve', 'tooltip.pastePreserve');

        /* 描画パネル（コピー/ペーストの下）/ Drawing panel (below Copy / Paste) */
        var drawingPanel = mainGroup.add('panel', undefined, L('panel.drawing'));
        setupPanel(drawingPanel);

        /* リアルタイムの描画と編集（上段）＋ 更新ボタン（次の行）を縦並び / Real-time editing checkbox (top) + Refresh button (next line), stacked */

        /* リアルタイムの描画と編集 / Real-time drawing & editing */
        var checkboxRealtime = addBooleanCheckbox(drawingPanel, 'checkbox.realtimeDrawing', 'LiveEdit_State_Machine', 'tooltip.realtimeDrawing');

        /* GPU プレビューを更新（View using GPU を2回トグルして再描画）/ Refresh GPU preview (toggle View using GPU twice to redraw) */
        var btnRefreshGpuPreview = drawingPanel.add('button', undefined, L('button.refreshGpuPreview'));
        btnRefreshGpuPreview.alignment = ['left', 'top']; /* 幅いっぱいにしない（ラベル幅）/ Do not fill width (size to label) */
        btnRefreshGpuPreview.helpTip = L('tooltip.refreshGpuPreview');
        btnRefreshGpuPreview.onClick = function () {
            runInMainEngine('try{app.executeMenuCommand("View using GPU");app.executeMenuCommand("View using GPU");}catch(e){}');
        };


        /* 読み出した環境設定を UI へ反映 / Apply fetched preferences to the UI */
        function applyPreferencesToUI(prefValues) {
            function asBool(prefKey) {
                return prefValues[prefKey] === true;
            }
            checkboxPoint.value = asBool('EnableActualPointTextSpaceAlign');
            checkboxArea.value = asBool('EnableActualAreaTextSpaceAlign');
            checkboxPreview.value = asBool('includeStrokeInBounds');
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

            /* 単位ポップアップを現在の rulerType に同期（onChange を発火させない）/ Sync the unit popup to the current rulerType (without firing onChange) */
            var unitPopupIndex = unitCodeToPopupIndex(PREF_STATE.rulerType);
            if (unitPopupIndex >= 0) {
                suppressUnitChange = true;
                unitDropdown.selection = unitPopupIndex;
                suppressUnitChange = false;
            }
            cursorStepField.text = readCursorKeyLengthInCurrentUnit();
        }

        /* キー入力の確定時に保存（不正値は元へ戻す）/ Save on commit (restore on invalid input) */
        cursorStepField.onChange = function () {
            if (!saveCursorKeyLengthInCurrentUnit(parseFloat(cursorStepField.text))) {
                cursorStepField.text = readCursorKeyLengthInCurrentUnit();
            }
        };

        /* 構築直後に現在値を反映（常に最新の環境設定を表示）/ Populate from current values right after building */
        applyPreferencesToUI(initialPrefs);

        /* 表示時：フォーカスと最新値の再読込 / On show: focus and reload current values */
        dialog.onShow = function () {
            cursorStepField.active = true;
            applyPreferencesToUI(readAllPreferences());
        };

        /* 再アクティブ時：外部変更（環境設定ダイアログ等）へクリックで追従。ただし編集中は同期しない（入力値の上書き防止）/ On re-activate: follow external changes via click, but skip while editing (avoid clobbering input) */
        dialog.onActivate = function () {
            if (isEditingCursorStep) return;
            applyPreferencesToUI(readAllPreferences());
        };

        dialog.show();

        /* ボタンの高さを2px詰める（レイアウト確定後に1回）/ Trim button heights by 2px (once, after layout) */
        trimButtonHeight(btnRefreshGpuPreview, 2);
    }

    main();

}());
