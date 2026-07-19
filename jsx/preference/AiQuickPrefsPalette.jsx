#target illustrator
#targetengine "AiQuickPrefsPalette"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

Illustrator の各種環境設定の切り替えと、選択オブジェクトの反転・回転を、常駐パレットでまとめて操作するユーティリティです。操作した時点で即時反映されます。

- パレット（常駐エンジン）で表示し、書き込み・DOM 操作は BridgeTalk でメインエンジンへ委譲（読み出しは同期で直接取得）
- 2カラム構成：左＝キー増加／整列オプション／字形の境界に整列、右＝変形オプション／変形、最下部（全幅）＝アートボード／コピー/ペースト／描画
- 環境設定ダイアログ等の外部変更は、パレットをクリック（再アクティブ）で同期
- 「字形の境界に整列」「変形オプション」のチェックボックスは Option+クリックでパネル内の全項目を一括切替（通常クリックは個別）
- パレットがアクティブなとき esc キーで閉じる

#### パネルと項目

- キー増加：カーソル移動量（cursorKeyLength）。単位ポップアップで定規単位を切替、↑↓ / Shift / Option で増減
- 整列オプション：プレビュー境界
- 字形の境界に整列：ポイント文字／エリア内文字（Option+クリックで2つを一括切替）
- 変形オプション：パターン／角／線幅と効果（Option+クリックで3つを一括切替）
- 変形：水平反転／垂直反転（横並び）、45°回転（方向はラジオで指定・デフォルト反時計回り）。いずれも選択全体の可視バウンディング中心を基点
- アートボード：アートボード名を表示／ビデオ定規の表示／枠線のカラー・幅
- コピー/ペースト：書式なしペースト／コピー元のレイヤーにペースト
- 描画：リアルタイムの描画と編集／プレビュー更新（GPU プレビューを更新）


### 紹介記事（note）

https://note.com/dtp_tranist/n/n41d8dc1961be

### Overview

A persistent-palette utility for batch-toggling various Illustrator preferences and flipping/rotating the selection. Every action applies immediately when triggered.

- Runs in a persistent-engine palette; writes and DOM operations are delegated to the main engine via BridgeTalk (reads are fetched directly/synchronously)
- Two-column layout: left = Key input / Align Options / Align to Glyph Bounds, right = Transform Options / Transform, bottom (full width) = Artboard / Copy / Paste / Drawing
- External changes (e.g. the Preferences dialog) sync when you click (re-activate) the palette
- In "Align to Glyph Bounds" and "Transform Options", Option-click a checkbox to toggle every item in that panel together (a normal click toggles one)
- Press Esc to close while the palette is active

#### Panels & options

- Key input: cursor step (cursorKeyLength); switch ruler unit via popup, adjust with Up/Down / Shift / Option
- Align Options: Preview Bounds
- Align to Glyph Bounds: Point Type / Area Type (Option-click toggles both together)
- Transform Options: Pattern Tiles / Corners / Strokes & Effects (Option-click toggles all three together)
- Transform: Flip Horizontal / Vertical (side by side), Rotate 45° (direction via radios, default counterclockwise). All pivot about the center of the selection's visible bounds
- Artboard: Show Artboard Name / Show Video Ruler / border color & width
- Copy / Paste: Paste without Formatting / Paste Remembers Layers
- Drawing: Real-time Drawing & Editing / Refresh Preview (GPU preview)

### 更新履歴 / Change Log

- v1.8.1 (20260629): 「字形の境界に整列」（2つ）と「変形オプション」（3つ）の各チェックボックスを Option+クリックでグループ一括切替に対応。連動処理を linkCheckboxGroup に共通化（角は整数プリファレンスを apply 関数で吸収）。
- v1.8.0 (20260628): 「45°回転」ボタンを追加（方向はラジオで指定・デフォルト反時計回り・選択中心基準）。反転ボタンを水平／垂直で横並び化し「水平反転／垂直反転」へ改称、パネル名を「変形」に短縮。整列パネルに「整列オプション」タイトルを付与、アートボードパネル名を「アートボード」に短縮。2カラム間隔を COLUMN_SPACING で調整。
- v1.7.1 (20260627): 「その他」を「コピー/ペースト」へ改称し、「描画」パネルを新設（リアルタイムの描画と編集を移動）。描画パネルにプレビュー更新ボタンを追加。
- v1.7.0 (20260627): 変形（反転）パネルを追加（選択中心基準・ロック／ガイド等はスキップ・合成行列で高速化）。レイアウト刷新（左＝キー入力／整列／字形の境界、右＝変形オプション／変形（反転））。テキスト／単位パネルとカンバスカラーを削除、「コピー元のレイヤーにペースト」を追加。esc キーで閉じる、編集中は同期を回避。命名整理、チェックボックス生成・ボタン高さ調整の共通化。
- v1.6.0 (20260627): 標準フォーマットへ整理（IIFE 化、ローカライズ構造、ブロックコメント）。パレット化＋BridgeTalk 委譲、キー入力の単位ポップアップ、クリック同期、ガイド／アートボードパネルを追加、2カラムレイアウト。
- v1.0 (20250804): 初期バージョン。

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiQuickPrefsPalette";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.8.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

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
            flip: { ja: "変形", en: "Transform (Flip)" },
            glyphBounds: { ja: "字形の境界に整列", en: "Align to Glyph Bounds" },
            artboard: { ja: "アートボード", en: "Artboard" },
            artboardBorder: { ja: "アートボードの枠線", en: "Artboard Border" },
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
            showArtboardName: { ja: "アートボード名を表示", en: "Show Artboard Name" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集", en: "Real-time Drawing & Editing" },
            pastePlain: { ja: "書式なしペースト", en: "Paste without Formatting" },
            pastePreserve: { ja: "コピー元のレイヤーにペースト", en: "Paste Remembers Layers" }
        },
        label: {
            strokeColor: { ja: "ハイライトのカラー", en: "Highlight Color" },
            strokeWidth: { ja: "ストロークの幅", en: "Stroke Width" }
        },
        tooltip: {
            cursorStep: { ja: "矢印キーでの移動量（環境設定 > 一般 > キー増加）。↑↓ / Shift=±10 / Option=±0.1", en: "Keyboard increment (Preferences > General). Up/Down / Shift=±10 / Option=±0.1" },
            unit: { ja: "定規の単位を切り替え", en: "Switch the ruler unit" },
            previewBounds: { ja: "整列・分布でプレビュー境界（線幅・効果を含む）を使用", en: "Use preview bounds (incl. stroke & effects) for align/distribute" },
            pointText: { ja: "ポイント文字を字形の境界で整列（Option+クリックで2つを一括切替）", en: "Align point type to glyph bounds (Option-click toggles both)" },
            areaText: { ja: "エリア内文字を字形の境界で整列（Option+クリックで2つを一括切替）", en: "Align area type to glyph bounds (Option-click toggles both)" },
            transformPattern: { ja: "変形時にパターンも変形する（Option+クリックで3つを一括切替）", en: "Transform pattern tiles when transforming (Option-click toggles all three)" },
            scaleCorners: { ja: "拡大・縮小時に角（ライブコーナー）の半径も拡大・縮小（Option+クリックで3つを一括切替）", en: "Scale corner radius when scaling (Option-click toggles all three)" },
            scaleStroke: { ja: "拡大・縮小時に線幅と効果も拡大・縮小（Option+クリックで3つを一括切替）", en: "Scale strokes & effects when scaling (Option-click toggles all three)" },
            flipHorizontal: { ja: "選択を水平方向に反転（選択全体の中心が基点）", en: "Flip the selection horizontally (about the selection center)" },
            flipVertical: { ja: "選択を垂直方向に反転（選択全体の中心が基点）", en: "Flip the selection vertically (about the selection center)" },
            rotate: { ja: "選択を45°回転（選択全体の中心が基点。方向は右のラジオで指定）", en: "Rotate the selection 45° (about the selection center; direction set by the radios)" },
            rotateCW: { ja: "時計回りに回転", en: "Rotate clockwise" },
            rotateCCW: { ja: "反時計回りに回転", en: "Rotate counterclockwise" },
            showArtboardName: { ja: "アートボード名をカンバスに表示", en: "Show artboard names on the canvas" },
            videoRuler: { ja: "ビデオ定規の表示を切り替え", en: "Toggle the video ruler" },
            strokeColor: { ja: "アートボードの枠線（ハイライト）のカラー", en: "Artboard border (highlight) color" },
            strokeWidth: { ja: "アートボードの枠線の幅（1〜4）", en: "Artboard border width (1-4)" },
            pastePlain: { ja: "書式を保持せずにペースト", en: "Paste without keeping formatting" },
            pastePreserve: { ja: "コピー元と同じレイヤーにペースト", en: "Paste into the original (source) layer" },
            realtimeDrawing: { ja: "リアルタイムの描画と編集を切り替え", en: "Toggle real-time drawing & editing" },
            refreshGpuPreview: { ja: "GPUプレビューを更新（再描画）", en: "Refresh the GPU preview (redraw)" }
        },
        button: {
            videoRuler: { ja: "ビデオ定規の表示", en: "Show Video Ruler" },
            flipHorizontal: { ja: "水平反転", en: "Horizontal" },
            flipVertical: { ja: "垂直反転", en: "Vertical" },
            rotate: { ja: "45°回転", en: "Rotate 45°" },
            rotateCW: { ja: "時計回り", en: "Clockwise" },
            rotateCCW: { ja: "反時計回り", en: "Counterclockwise" },
            refreshGpuPreview: { ja: "プレビュー更新", en: "Refresh Preview" }
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
        var text = node[currentLanguage] || node.en || "";
        return text.replace(/\{slash\}/g, "/");
    }

    /* 現在言語のラベル文字列を返す / Return the current-language label string */
    function L(key) {
        return getLabel(key);
    }

    /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
    function labelText(key) {
        return getLabel(key) + (currentLanguage === "ja" ? "：" : ":");
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
        var closestIndex = 0;
        var closestDistance = Infinity;
        for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
            var preset = STROKE_COLOR_PRESETS[i];
            var distance = Math.abs(preset.r - r) + Math.abs(preset.g - g) + Math.abs(preset.b - b);
            if (distance < closestDistance) {
                closestDistance = distance;
                closestIndex = i;
            }
        }
        return closestIndex;
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
            showArtboardLabelOnCanvas: getBool("showArtboardLabelOnCanvas"),
            ArtboardBBColorRed: getReal("ArtboardBBColorRed"),
            ArtboardBBColorGreen: getReal("ArtboardBBColorGreen"),
            ArtboardBBColorBlue: getReal("ArtboardBBColorBlue"),
            ArtboardBBWidth: getReal("ArtboardBBWidth"),
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

    /* 選択オブジェクトを、選択全体の可視バウンディング中心を基点に反転（水平=-100,100／垂直=100,-100）。DOM 操作なのでメインエンジンへ委譲 */
    /* Flip the selection around the center of its overall visible bounds (H=-100,100 / V=100,-100); delegated to the main engine since it touches the DOM */
    /* 平行移動→反転→平行移動の3変形を1つの合成行列にまとめ、item.transform をオブジェクトあたり1回に削減（高速化）*/
    /* Collapse the 3-step translate/scale/translate into a single matrix so each item is transformed only once (faster) */
    /* visibleBounds が取れない/変形できない項目（ロック・非表示・ガイド等）は try/catch でスキップ。測定できた項目が無ければ中断 */
    /* Skip items whose bounds/transform fail (locked, hidden, guides, ...) via try/catch; abort if nothing could be measured */
    function btFlipSelection(scaleX, scaleY) {
        var scaleFractionX = Number(scaleX) / 100; /* -1 or 1 */
        var scaleFractionY = Number(scaleY) / 100;
        var body = '' +
            'if(app.documents.length>0){' +
            'var doc=app.activeDocument;' +
            'var selection=doc.selection;' +
            'if(selection&&selection.length>0){' +
            'var sx=' + scaleFractionX + ',sy=' + scaleFractionY + ';' +
            'var left=Infinity,top=-Infinity,right=-Infinity,bottom=Infinity;' +
            'var measured=false;' +
            'for(var i=0;i<selection.length;i++){try{var bounds=selection[i].visibleBounds;if(bounds[0]<left)left=bounds[0];if(bounds[1]>top)top=bounds[1];if(bounds[2]>right)right=bounds[2];if(bounds[3]<bottom)bottom=bounds[3];measured=true;}catch(e){}}' +
            'if(measured){' +
            'var anchorX=(left+right)/2,anchorY=(top+bottom)/2;' +
            /* 中心 (anchorX,anchorY) を基点に倍率 (sx,sy) で反転する原点基準行列 / Origin-based matrix reflecting about (anchorX,anchorY) by (sx,sy) */
            'var matrix=app.getIdentityMatrix();matrix.mValueA=sx;matrix.mValueD=sy;matrix.mValueTX=anchorX*(1-sx);matrix.mValueTY=anchorY*(1-sy);' +
            'for(var i=0;i<selection.length;i++){try{selection[i].transform(matrix,true,true,true,true,1,Transformation.DOCUMENTORIGIN);}catch(e){}}' +
            '}}}';
        runInMainEngine(body);
    }

    /* 選択オブジェクトを、選択全体の可視バウンディング中心を基点に回転（angleDegrees：正＝反時計回り／負＝時計回り）。DOM 操作なのでメインエンジンへ委譲 */
    /* Rotate the selection around the center of its overall visible bounds (angleDegrees: positive = CCW / negative = CW); delegated to the main engine since it touches the DOM */
    /* 平行移動→回転→平行移動を1つの合成行列にまとめ、cos/sin はコントローラ側で算出して数値を埋め込む（btFlipSelection と同様）*/
    /* Collapse translate/rotate/translate into one matrix; cos/sin are computed controller-side and embedded as numbers (same approach as btFlipSelection) */
    function btRotateSelection(angleDegrees) {
        var radians = Number(angleDegrees) * Math.PI / 180;
        var cos = Math.cos(radians);
        var sin = Math.sin(radians);
        var body = '' +
            'if(app.documents.length>0){' +
            'var doc=app.activeDocument;' +
            'var selection=doc.selection;' +
            'if(selection&&selection.length>0){' +
            'var ca=' + cos + ',sa=' + sin + ';' +
            'var left=Infinity,top=-Infinity,right=-Infinity,bottom=Infinity;' +
            'var measured=false;' +
            'for(var i=0;i<selection.length;i++){try{var bounds=selection[i].visibleBounds;if(bounds[0]<left)left=bounds[0];if(bounds[1]>top)top=bounds[1];if(bounds[2]>right)right=bounds[2];if(bounds[3]<bottom)bottom=bounds[3];measured=true;}catch(e){}}' +
            'if(measured){' +
            'var anchorX=(left+right)/2,anchorY=(top+bottom)/2;' +
            /* 中心 (anchorX,anchorY) を基点に角度で回転する原点基準行列（x'=A*x+C*y+TX, y'=B*x+D*y+TY）/ Origin-based matrix rotating about (anchorX,anchorY) */
            'var matrix=app.getIdentityMatrix();matrix.mValueA=ca;matrix.mValueB=sa;matrix.mValueC=-sa;matrix.mValueD=ca;matrix.mValueTX=anchorX*(1-ca)+anchorY*sa;matrix.mValueTY=anchorY*(1-ca)-anchorX*sa;' +
            'for(var i=0;i<selection.length;i++){try{selection[i].transform(matrix,true,true,true,true,1,Transformation.DOCUMENTORIGIN);}catch(e){}}' +
            '}}}';
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

    /* チェックボックス群を Option+クリックで連動させる（通常クリックは個別）。各項目は { checkbox, apply } で、apply(value) が値の書き込みを担う */
    /* Link a group of checkboxes so Option-click toggles them together (a normal click toggles one). Each item is { checkbox, apply }, where apply(value) writes the preference */
    function linkCheckboxGroup(items) {
        for (var i = 0; i < items.length; i++) {
            (function (self) {
                self.checkbox.onClick = function () {
                    self.apply(self.checkbox.value === true);
                    if (ScriptUI.environment.keyboardState.altKey) {
                        for (var j = 0; j < items.length; j++) {
                            if (items[j] === self) continue;
                            items[j].checkbox.value = self.checkbox.value;
                            items[j].apply(items[j].checkbox.value === true);
                        }
                    }
                };
            })(items[i]);
        }
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

        /* ----- 左列：キー増加 / 整列 / Left column: Key input / Align ----- */

        /* キー増加パネル（カーソル移動量）と単位ポップアップ / Key input panel (cursor step) with the unit popup */
        var keyInputPanel = leftColumn.add('panel', undefined, L('panel.keyInput'));
        keyInputPanel.orientation = 'row';
        keyInputPanel.alignChildren = ['left', 'center'];
        keyInputPanel.margins = PANEL_MARGINS;

        var cursorStepField = keyInputPanel.add('edittext', undefined, "1.0");
        cursorStepField.characters = 4;
        cursorStepField.helpTip = L('tooltip.cursorStep');

        /* 編集中フラグ：キー増加欄にフォーカスがある間は外部同期で値を上書きしない / Editing flag: don't let external sync overwrite while the field has focus */
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

        /* 整列パネル / Align panel */
        var alignPanel = leftColumn.add('panel', undefined, L('panel.align'));
        setupPanel(alignPanel);

        /* プレビュー境界 / Preview bounds */
        var checkboxPreview = addBooleanCheckbox(alignPanel, 'checkbox.previewBounds', 'includeStrokeInBounds', 'tooltip.previewBounds');

        /* 字形の境界に整列パネル（整列パネルとは独立、左列に並べる）/ Align to glyph bounds panel (independent of Align, stacked in the left column) */
        var glyphBoundsPanel = leftColumn.add('panel', undefined, L('panel.glyphBounds'));
        setupPanel(glyphBoundsPanel);

        var checkboxPoint = addBooleanCheckbox(glyphBoundsPanel, 'checkbox.pointText', 'EnableActualPointTextSpaceAlign', 'tooltip.pointText');
        var checkboxArea = addBooleanCheckbox(glyphBoundsPanel, 'checkbox.areaText', 'EnableActualAreaTextSpaceAlign', 'tooltip.areaText');

        /* Option+クリックでポイント文字／エリア内文字の2つを同じ状態に連動 / Option-click links Point/Area type to the same state */
        linkCheckboxGroup([
            { checkbox: checkboxPoint, apply: function (value) { btSetBooleanPreference('EnableActualPointTextSpaceAlign', value); } },
            { checkbox: checkboxArea,  apply: function (value) { btSetBooleanPreference('EnableActualAreaTextSpaceAlign', value); } }
        ]);

        /* ----- 右列：変形 / Right column: Transform ----- */

        /* 変形パネル / Transform panel */
        var transformPanel = rightColumn.add('panel', undefined, L('panel.transform'));
        setupPanel(transformPanel);

        /* パターンを変形 / Transform patterns */
        var checkboxPattern = addBooleanCheckbox(transformPanel, 'checkbox.transformPattern', 'transformPatterns', 'tooltip.transformPattern');

        /* 角を拡大・縮小（1=ON, 2=OFF。Boolean でなく整数なので個別ハンドラ）/ Scale corners (1=ON, 2=OFF; integer, so a dedicated handler) */
        var checkboxCorner = transformPanel.add('checkbox', undefined, L('checkbox.scaleCorners'));
        checkboxCorner.helpTip = L('tooltip.scaleCorners');
        checkboxCorner.onClick = function () {
            btSetIntegerPreference("policyForPreservingCorners", checkboxCorner.value ? 1 : 2);
        };

        /* 線幅と効果も拡大・縮小 / Scale strokes and effects */
        var checkboxStroke = addBooleanCheckbox(transformPanel, 'checkbox.scaleStroke', 'scaleLineWeight', 'tooltip.scaleStroke');

        /* Option+クリックでパターン／角／線幅と効果の3つを同じ状態に連動（角は整数プリファレンス 1=ON,2=OFF）/ Option-click links Pattern/Corners/Strokes to the same state (Corners uses an integer pref: 1=ON, 2=OFF) */
        linkCheckboxGroup([
            { checkbox: checkboxPattern, apply: function (value) { btSetBooleanPreference('transformPatterns', value); } },
            { checkbox: checkboxCorner,  apply: function (value) { btSetIntegerPreference('policyForPreservingCorners', value ? 1 : 2); } },
            { checkbox: checkboxStroke,  apply: function (value) { btSetBooleanPreference('scaleLineWeight', value); } }
        ]);

        /* 変形（反転）パネル（変形オプションの下・横並び・ラベル幅・左右中央）/ Transform (Flip) panel (below Transform Options, side by side, sized to label, centered) */
        var flipPanel = rightColumn.add('panel', undefined, L('panel.flip'));
        setupPanel(flipPanel);
        flipPanel.alignChildren = ['center', 'top'];

        /* 2ボタンを横並びに / Lay out the two buttons in a row */
        var flipRow = flipPanel.add('group');
        flipRow.orientation = 'row';
        flipRow.alignChildren = ['center', 'center'];

        /* 水平方向に反転（選択全体の中心を基点に左右反転）/ Flip horizontal (around selection center) */
        var btnFlipHorizontal = flipRow.add('button', undefined, L('button.flipHorizontal'));
        btnFlipHorizontal.helpTip = L('tooltip.flipHorizontal');
        btnFlipHorizontal.onClick = function () {
            btFlipSelection(-100, 100);
        };

        /* 垂直方向に反転（選択全体の中心を基点に上下反転）/ Flip vertical (around selection center) */
        var btnFlipVertical = flipRow.add('button', undefined, L('button.flipVertical'));
        btnFlipVertical.helpTip = L('tooltip.flipVertical');
        btnFlipVertical.onClick = function () {
            btFlipSelection(100, -100);
        };

        /* 45°回転：ボタン1つ＋方向ラジオ（時計回り／反時計回り、デフォルトは反時計回り）/ 45° rotate: a single button plus direction radios (CW / CCW, default CCW) */
        var rotateGroup = flipPanel.add('group');
        rotateGroup.orientation = 'row';
        rotateGroup.alignChildren = ['left', 'center'];

        var btnRotate = rotateGroup.add('button', undefined, L('button.rotate'));
        btnRotate.helpTip = L('tooltip.rotate');

        var rotateDirGroup = rotateGroup.add('group');
        rotateDirGroup.orientation = 'column';
        rotateDirGroup.alignChildren = ['left', 'center'];

        var radioRotateCCW = rotateDirGroup.add('radiobutton', undefined, L('button.rotateCCW'));
        radioRotateCCW.helpTip = L('tooltip.rotateCCW');
        var radioRotateCW = rotateDirGroup.add('radiobutton', undefined, L('button.rotateCW'));
        radioRotateCW.helpTip = L('tooltip.rotateCW');
        radioRotateCCW.value = true; /* デフォルトは反時計回り / Default: counterclockwise */

        btnRotate.onClick = function () {
            /* 正＝反時計回り／負＝時計回り / positive = CCW, negative = CW */
            btRotateSelection(radioRotateCW.value ? -45 : 45);
        };

        /* ----- 全幅：アートボード名と枠線 / その他（一番下）/ Full width: Artboard / Other (bottom) ----- */

        /* アートボード名と枠線パネル / Artboard name & border panel */
        var artboardPanel = mainGroup.add('panel', undefined, L('panel.artboard'));
        setupPanel(artboardPanel);

        var suppressArtboardChange = false;

        /* 上部を2カラムに：左＝アートボード名チェック、右＝ビデオ定規ボタン（互いに天地中央）/ Top row in two columns: left = artboard-name checkbox, right = video ruler button (vertically centered) */
        var artboardTopRow = artboardPanel.add('group');
        artboardTopRow.orientation = 'row';
        artboardTopRow.alignChildren = ['fill', 'center'];

        var artboardLeftCol = artboardTopRow.add('group');
        artboardLeftCol.orientation = 'column';
        artboardLeftCol.alignChildren = ['left', 'top'];

        var artboardRightCol = artboardTopRow.add('group');
        artboardRightCol.orientation = 'column';
        artboardRightCol.alignChildren = ['left', 'top'];

        /* アートボード名を表示 / Show artboard name */
        var checkboxShowArtboardName = artboardLeftCol.add('checkbox', undefined, L('checkbox.showArtboardName'));
        checkboxShowArtboardName.helpTip = L('tooltip.showArtboardName');
        checkboxShowArtboardName.onClick = function () {
            applyArtboard();
        };

        /* ビデオ定規（メニューコマンドのトグル）/ Video ruler (menu-command toggle) */
        var btnVideoRuler = artboardRightCol.add('button', undefined, L('button.videoRuler'));
        btnVideoRuler.alignment = ['left', 'top']; /* 幅いっぱいにしない（ラベル幅）/ Do not fill width (size to label) */
        btnVideoRuler.helpTip = L('tooltip.videoRuler');
        btnVideoRuler.onClick = function () {
            runInMainEngine('try{app.executeMenuCommand("videoruler");}catch(e){}');
        };

        /* アートボードの枠線サブパネル / Artboard border sub-panel */
        var artboardBorderPanel = artboardPanel.add('panel', undefined, L('panel.artboardBorder'));
        setupPanel(artboardBorderPanel);

        /* ハイライトのカラー / Highlight color */
        var strokeColorRow = artboardBorderPanel.add('group');
        strokeColorRow.orientation = 'row';
        strokeColorRow.alignChildren = ['left', 'center'];
        strokeColorRow.add('statictext', undefined, labelText('label.strokeColor'));
        var strokeColorDropdown = strokeColorRow.add('dropdownlist', undefined, buildStrokeColorNames());
        strokeColorDropdown.helpTip = L('tooltip.strokeColor');
        strokeColorDropdown.onChange = function () {
            if (suppressArtboardChange || !strokeColorDropdown.selection) return;
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
        for (var i = 0; i < rbStrokeWidths.length; i++) {
            rbStrokeWidths[i].helpTip = L('tooltip.strokeWidth');
            rbStrokeWidths[i].onClick = function () {
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
            var colorIndex = strokeColorDropdown.selection ? strokeColorDropdown.selection.index : STROKE_COLOR_BLACK_INDEX;
            var colorPreset = STROKE_COLOR_PRESETS[colorIndex];
            btApplyArtboardBorder(checkboxShowArtboardName.value === true, colorPreset.r, colorPreset.g, colorPreset.b, getSelectedStrokeWidth());
        }

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

        /* リアルタイムの描画と編集（左）＋ 更新ボタン（右端）を横並び。チェックボックスを伸縮させて余白を吸収し、ボタンを右へ寄せる */
        /* Real-time editing checkbox (left) + Refresh button (right edge); the checkbox absorbs the slack to push the button right */
        var drawingRow = drawingPanel.add('group');
        drawingRow.orientation = 'row';
        drawingRow.alignChildren = ['left', 'center'];
        drawingRow.alignment = ['fill', 'top'];

        /* リアルタイムの描画と編集 / Real-time drawing & editing */
        var checkboxRealtime = addBooleanCheckbox(drawingRow, 'checkbox.realtimeDrawing', 'LiveEdit_State_Machine', 'tooltip.realtimeDrawing');
        checkboxRealtime.alignment = ['fill', 'center']; /* 余白を吸収してボタンを右端へ / Absorb slack to push the button to the right */

        /* GPU プレビューを更新（View using GPU を2回トグルして再描画）/ Refresh GPU preview (toggle View using GPU twice to redraw) */
        var btnRefreshGpuPreview = drawingRow.add('button', undefined, L('button.refreshGpuPreview'));
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

            /* アートボード名と枠線 / Artboard name & border */
            checkboxShowArtboardName.value = asBool('showArtboardLabelOnCanvas');
            var borderColorRed = parseFloat(prefValues['ArtboardBBColorRed']);
            var borderColorGreen = parseFloat(prefValues['ArtboardBBColorGreen']);
            var borderColorBlue = parseFloat(prefValues['ArtboardBBColorBlue']);
            if (isNaN(borderColorRed)) borderColorRed = 0;
            if (isNaN(borderColorGreen)) borderColorGreen = 0;
            if (isNaN(borderColorBlue)) borderColorBlue = 0;
            suppressArtboardChange = true;
            strokeColorDropdown.selection = findClosestStrokeColor(borderColorRed, borderColorGreen, borderColorBlue);
            suppressArtboardChange = false;
            var borderWidth = Math.round(parseFloat(prefValues['ArtboardBBWidth']));
            if (isNaN(borderWidth)) borderWidth = 1;
            if (borderWidth < 1) borderWidth = 1;
            if (borderWidth > 4) borderWidth = 4;
            for (var i = 0; i < rbStrokeWidths.length; i++) {
                rbStrokeWidths[i].value = (i === borderWidth - 1);
            }

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

        /* キー増加の確定時に保存（不正値は元へ戻す）/ Save on commit (restore on invalid input) */
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
        trimButtonHeight(btnVideoRuler, 2);
        trimButtonHeight(btnFlipHorizontal, 2);
        trimButtonHeight(btnFlipVertical, 2);
        trimButtonHeight(btnRefreshGpuPreview, 2);
    }

    main();

}());
