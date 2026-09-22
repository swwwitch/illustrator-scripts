#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "DialogEngine"

/*

### 概要

アートボードのサイズを「操作」×「対象」×「サイズ」の組み合わせで自動調整します。
選択オブジェクトに合わせるほか、各アートボード内のオブジェクトに合わせて全アートボードを個別に調整できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitArtboardWithMargin.md

note記事も参照してください。
https://note.com/dtp_transit/n/n15d3c6c5a1e5

### Overview

Adjusts artboard size by operation, target and size (width & height).
Fits to the selection, or fits every artboard individually to the objects it contains.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitArtboardWithMargin.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitArtboardWithMargin";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.9.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-20";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitArtboardWithMargin.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitArtboardWithMargin.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_transit/n/n15d3c6c5a1e5"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値（保存済みの設定があればそちらが優先） / Dialog defaults (stored settings win) */
    var DIALOG_DEFAULTS = {
        marginByUnit: {
            mm: '5',
            px: '20',
            pt: '10',
            _fallback: '0'
        },
        previewBounds: true,        // プレビュー境界(visibleBounds)を既定に / use visibleBounds by default
        roundMode: 'pixelGrid',     // 既定の丸めモード（pixelGrid / currentUnit / none）/ default rounding mode
        link: true                  // 上下左右の連動を既定ON / link margins by default
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var BUTTON_ROW_TOP_MARGIN = 5;           /* ボタンエリアの上余白 / top margin of the button row */

    /* ダイアログの不透明度と、初回表示時の画面中央からの横オフセット / Dialog opacity and first-run offset from screen center */
    var DIALOG_OPACITY = 0.98;
    var DIALOG_FIRST_RUN_OFFSET_X = 300;

    /* 行ラベルの幅（言語別） / Row label widths per language */
    var BASIS_LABEL_WIDTHS = { ja: 40, en: 76 };    /* 調整基準パネル / adjustment basis panel */
    var MARGIN_LABEL_WIDTHS = { ja: 32, en: 62 };   /* マージンパネル / margin panel */

    /* マージン入力欄の桁数 / Margin field width in characters */
    var MARGIN_FIELD_CHARACTERS = 4;

    /**
     * ウィンドウの共通設定を適用する
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
     * パネルの共通設定を適用する
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
     * 横並びの行グループ（ボタン列など）の共通設定を適用する
     * @param {Group} rowGroup - 対象のグループ
     * @param {string} [alignment] - グループ自身の配置（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, alignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = alignment || "left";
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

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
     * 現在の定規単位を UnitValue に渡せる文字列で返す（UnitValue が扱えない単位は pt に寄せる）
     * @returns {string} 単位の文字列
     */
    function getRulerUnitString() {
        var unit = getUnitInfo();
        return (unit.code >= 0 && unit.code <= 6) ? unit.label : 'pt';
    }

    /**
     * 単位ごとの初期マージン値を返す
     * @param {string} unit - 単位の文字列
     * @returns {string} 初期マージン値
     */
    function getDefaultMargin(unit) {
        return DIALOG_DEFAULTS.marginByUnit.hasOwnProperty(unit) ?
            DIALOG_DEFAULTS.marginByUnit[unit] :
            DIALOG_DEFAULTS.marginByUnit._fallback;
    }

    /**
     * 数値＋単位を pt に変換する
     * @param {number|string} value - 値
     * @param {string} unit - 単位の文字列
     * @returns {number} pt。変換できなければ NaN
     */
    function toPt(value, unit) {
        var numericValue = Number(value);
        if (isNaN(numericValue)) return NaN;
        // 歯/Q は UnitValue が非対応のため手計算（1H = 1Q = 0.25mm、1mm = 72/25.4pt） / H and Q are unsupported by UnitValue
        if (unit === "H" || unit === "Q") {
            return numericValue * 0.25 * 72 / 25.4;
        }
        /* UnitValue が単位を受け付けないときに備える / guard against units UnitValue rejects */
        try {
            return new UnitValue(numericValue, unit).as('pt');
        } catch (e) {
            return NaN;
        }
    }

    // =========================================
    // ローカライズ / Localize
    // =========================================

    /**
     * 実行環境の言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    var LABELS = {
        dialog: {
            title: { ja: "アートボードサイズを調整", en: "Adjust Artboard Size" }
        },
        panel: {
            adjustmentBasis: { ja: "調整基準", en: "Adjustment basis" },
            margin: { ja: "マージン", en: "Margin" },
            fineTuning: { ja: "アートボードサイズの微調整", en: "Artboard size fine-tuning" }
        },
        fieldLabel: {
            operation: { ja: "操作", en: "Operation" },
            scope: { ja: "対象", en: "Target" },
            size: { ja: "サイズ", en: "Size" },
            vertical: { ja: "上下", en: "Vertical" },
            horizontal: { ja: "左右", en: "Horizontal" }
        },
        // 操作（fit/expand）と対象（current/all）の2軸、丸めモード / operation (fit/expand), scope (current/all), rounding mode
        radio: {
            fit: { ja: "オブジェクトに合わせる", en: "Fit to objects" },
            expand: { ja: "アートボードを拡張", en: "Expand artboard" },
            currentArtboard: { ja: "現在のアートボード", en: "Current artboard" },
            allArtboards: { ja: "すべてのアートボード", en: "All artboards" },
            roundPixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize to pixel grid" },
            roundCurrentUnit: { ja: "現在の単位で値を整数値に", en: "Round values in current unit" },
            roundNone: { ja: "何もしない", en: "Do nothing" }
        },
        checkbox: {
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            linked: { ja: "連動", en: "Linked" },
            previewBounds: { ja: "プレビュー境界", en: "Preview bounds" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてから実行してください。", en: "Please open a document first." },
            enterNumber: { ja: "数値を入力してください。", en: "Please enter a number." },
            errorOccurred: { ja: "エラーが発生しました：", en: "An error occurred: " },
            marginTooLarge: {
                ja: "マージンが大きすぎて有効なサイズにできないため、適用をスキップしました。",
                en: "The margin is too large to produce a valid size; skipped."
            }
        },
        tooltip: {
            fit: {
                ja: "オブジェクトの外接＋マージンのサイズにアートボードを合わせます（選択が無いときは各アートボード内のオブジェクトが対象）",
                en: "Resize artboards to the objects' bounds plus margins (with no selection, each artboard uses the objects it contains)"
            },
            expand: {
                ja: "アートボード自身のサイズにマージンを加減します（マイナス値で縮小）",
                en: "Grow/shrink the artboards themselves by the margins (negative shrinks)"
            },
            currentArtboard: {
                ja: "現在のアートボードのみを対象にします",
                en: "Apply to the current artboard only"
            },
            allArtboards: {
                ja: "すべてのアートボードを対象にします（「合わせる」では各アートボード内のオブジェクトに合わせます）",
                en: "Apply to all artboards (Fit uses the objects each artboard contains)"
            },
            marginInput: {
                ja: "↑↓で±1、Shift+↑↓で10の倍数にスナップ、Option+↑↓で±0.1",
                en: "Arrow: ±1, Shift: snap to 10, Option: ±0.1"
            },
            axisEnable: {
                ja: "OFFにすると実行時のサイズのまま固定します（連動は自動でOFF）。Option+クリックでこちらだけON",
                en: "Off keeps this dimension at its original size (auto-unlinks). Option-click to solo it"
            },
            linked: {
                ja: "上下の値を左右にも自動で適用します",
                en: "Apply the vertical value to horizontal as well"
            },
            previewBounds: {
                ja: "ON：線・効果を含む見た目の境界（プレビュー境界）で計測／OFF：パスの幾何境界で計測",
                en: "On: measure with preview (visible) bounds incl. strokes/effects; Off: geometric path bounds"
            },
            roundPixelGrid: {
                ja: "座標とサイズを整数ピクセルに丸めます",
                en: "Round position and size to integer pixels"
            },
            roundCurrentUnit: {
                ja: "現在の定規単位で座標とサイズを整数に丸めます",
                en: "Round position and size to integers in the current ruler unit"
            },
            roundNone: {
                ja: "丸めずに計測値のまま設定します",
                en: "Apply the measured values without rounding"
            }
        }
    };

    /**
     * ローカライズ文字列を取得する（キー漏れ時は英語へフォールバック）
     * @param {Object} labelSet - { ja, en } のラベル
     * @returns {string} 表示言語の文字列
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        if (labelSet[uiLang] != null) return labelSet[uiLang];
        return (labelSet.en != null) ? labelSet.en : "";
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {Object} labelSet - ラベル
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelSet) {
        return getLabel(labelSet) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // エラー処理 / Error handling
    // =========================================

    /**
     * Error を行番号・ファイル名付きで読みやすく整形する
     * @param {Error} error - 例外
     * @returns {string} 整形した文字列
     */
    function formatError(error) {
        var messageText = (error && error.message) ? String(error.message) : String(error);
        var lineText = (error && error.line) ? (" line " + error.line) : "";
        var fileText = (error && error.fileName) ? (" (" + error.fileName + ")") : "";
        return messageText + lineText + fileText;
    }

    // =========================================
    // プレビュー管理 / Preview manager
    // =========================================

    /**
     * プレビューの適用と復元を制御するクラス / Preview apply/restore manager
     *
     * - updatePreview() のたびに rollback() で開いた時点の状態へ戻してから addStep() で最新状態を適用
     * - OK/Cancel 時に rollback() してプレビューを開いた時点へ戻す
     *
     * app.undo() の回数に依存すると、複数アートボード書き換え時に undo 粒度とズレて
     * 戻しすぎ/戻し不足が起きる。そのため巻き戻しは restoreFn（スナップショット復元）で行う。
     * @param {function} restoreFn - プレビュー前の状態へ戻す関数
     */
    function PreviewManager(restoreFn) {
        this.restoreFn = restoreFn;
        this.dirty = false; // プレビューによる変更が未復元か / preview changes pending restore

        /**
         * 変更操作を実行し、未復元フラグを立てる
         * @param {function} func - 変更操作
         * @returns {void}
         */
        this.addStep = function (func) {
            try {
                func();
                this.dirty = true;
                app.redraw();
            } catch (e) {
                $.writeln("[PreviewManager] addStep error: " + e);
            }
        };

        /**
         * プレビューを開いた時点の状態へ戻す
         * @returns {void}
         */
        this.rollback = function () {
            if (this.dirty && typeof this.restoreFn === "function") {
                try {
                    this.restoreFn();
                } catch (e) {
                    $.writeln("[PreviewManager] rollback error: " + e);
                }
            }
            this.dirty = false;
            app.redraw();
        };

        /**
         * 現在の状態を確定する（OK時）
         * OK時は一度 rollback() で元に戻してから main() 側で本処理を1回だけ実行するため、ここでは rollback のみ。
         * @returns {void}
         */
        this.confirm = function () {
            this.rollback();
        };
    }

    // =========================================
    // ダイアログ位置の記憶 / Dialog position persistence
    // =========================================
    // 共通エンジン名でセッションをまたいで位置を記憶し、key で保存先を分離する。
    // Share session state across scripts; separate each dialog by key.

    var DIALOG_POSITION_KEY = "__FitArtboardWithMargin_Dialog";

    /**
     * 保存済みのダイアログ位置を取得する
     * @param {string} storageKey - $.global のキー
     * @returns {number[]|null} [x, y]。無ければ null
     */
    function getStoredLocation(storageKey) {
        return $.global[storageKey] && $.global[storageKey].length === 2 ? $.global[storageKey] : null;
    }

    /**
     * ダイアログ位置をセッションに保存する
     * @param {string} storageKey - $.global のキー
     * @param {number[]} location - [x, y]
     * @returns {void}
     */
    function storeLocation(storageKey, location) {
        $.global[storageKey] = [location[0], location[1]];
    }

    /**
     * 位置を画面内に収める
     * @param {number[]} location - [x, y]
     * @returns {number[]} 画面内に収めた [x, y]
     */
    function clampLocationToScreen(location) {
        /* 画面情報が取れない環境では元の位置のまま / keep the location when screen info is unavailable */
        try {
            var visibleBounds = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
            var clampedX = Math.max(visibleBounds[0] + 10, Math.min(location[0], visibleBounds[2] - 10));
            var clampedY = Math.max(visibleBounds[1] + 10, Math.min(location[1], visibleBounds[3] - 10));
            return [clampedX, clampedY];
        } catch (e) {
            return location;
        }
    }

    /**
     * ダイアログ位置の記憶を設定し、保存関数を返す
     * 保存位置があれば表示時に復元、無ければ初回はセンターからのオフセットで表示する。
     * @param {Window} dialogWindow - 対象のダイアログ
     * @param {string} positionKey - $.global のキー
     * @param {number} firstRunOffsetX - 初回表示時の中央からの横オフセット
     * @returns {function} 現在位置を保存する関数
     */
    function attachPositionPersistence(dialogWindow, positionKey, firstRunOffsetX) {
        var savedLocation = getStoredLocation(positionKey);

        var persist = function () {
            storeLocation(positionKey, [dialogWindow.location[0], dialogWindow.location[1]]);
        };

        if (savedLocation) {
            dialogWindow.onShow = function () {
                dialogWindow.location = clampLocationToScreen(savedLocation);
            };
        } else {
            dialogWindow.onShow = function () {
                dialogWindow.layout.layout(true);
                var screenWidth = $.screens[0].right - $.screens[0].left;
                var screenHeight = $.screens[0].bottom - $.screens[0].top;
                var centerX = screenWidth / 2 - dialogWindow.bounds.width / 2;
                var centerY = screenHeight / 2 - dialogWindow.bounds.height / 2;
                dialogWindow.location = [centerX + firstRunOffsetX, centerY];
            };
        }

        dialogWindow.onMove = persist;
        return persist;
    }

    // =========================================
    // 設定の記憶（セッション内） / Settings persistence (session only)
    // =========================================
    // $.global に設定を保持。#targetengine のためセッション中は保持されるが、再起動でリセット。
    // Kept in $.global; persists during the session but resets when Illustrator restarts.

    var SETTINGS_KEY = "__FitArtboardWithMargin_Settings";

    /**
     * 保存済みの設定を取得する
     * @returns {Object|null} 設定。無ければ null
     */
    function getStoredSettings() {
        var stored = $.global[SETTINGS_KEY];
        return (stored && typeof stored === "object") ? stored : null;
    }

    /**
     * 設定をセッションに保存する
     * @param {Object} settings - 設定
     * @returns {void}
     */
    function storeSettings(settings) {
        $.global[SETTINGS_KEY] = settings;
    }

    /**
     * 保存済み設定と文脈から、ダイアログの初期値を解決する
     * operation: "fit"（オブジェクトに合わせる、要選択）/ "expand"（アートボードを拡張）
     * scope: "current"（現在のアートボード）/ "all"（すべてのアートボード）
     * @param {string} defaultMargin - 保存が無いときのマージン
     * @param {number} artboardCount - アートボードの数
     * @param {boolean} hasSelection - 計測できる選択があるか
     * @returns {Object} ダイアログの初期値
     */
    function resolveInitialSettings(defaultMargin, artboardCount, hasSelection) {
        var saved = getStoredSettings();

        // 操作：選択があれば fit、無ければ expand を既定に / operation default
        var operation = hasSelection ? "fit" : "expand";
        if (saved && (saved.operation === "fit" || saved.operation === "expand")) {
            operation = saved.operation;
        }

        // 対象：選択なし・複数アートボードなら all、それ以外は current を既定に / scope default
        var scope = (!hasSelection && artboardCount > 1) ? "all" : "current";
        if (saved && (saved.scope === "current" || saved.scope === "all")) {
            scope = saved.scope;
        }
        // 「合わせる」で選択が無いときは、各アートボード内のオブジェクトが対象になるため「すべて」固定
        // Fit without a selection works per artboard, so the scope is locked to all.
        if (operation === "fit" && !hasSelection) scope = "all";

        var link = (saved && typeof saved.link === "boolean") ? saved.link : DIALOG_DEFAULTS.link;
        // 連動ONのときは上下・左右とも有効に揃える / when linked, both axes are enabled
        var verticalEnabled = link ? true : ((saved && typeof saved.verticalEnabled === "boolean") ? saved.verticalEnabled : true);
        var horizontalEnabled = link ? true : ((saved && typeof saved.horizontalEnabled === "boolean") ? saved.horizontalEnabled : true);

        return {
            marginV: (saved && saved.marginV != null) ? saved.marginV : defaultMargin,
            marginH: (saved && saved.marginH != null) ? saved.marginH : defaultMargin,
            link: link,
            verticalEnabled: verticalEnabled,
            horizontalEnabled: horizontalEnabled,
            previewBounds: (saved && typeof saved.previewBounds === "boolean") ? saved.previewBounds : DIALOG_DEFAULTS.previewBounds,
            roundMode: (saved && saved.roundMode) ? saved.roundMode : DIALOG_DEFAULTS.roundMode,
            operation: operation,
            scope: scope
        };
    }

    // =========================================
    // 矩形・境界のユーティリティ / Rect & bounds utilities
    // =========================================

    /**
     * 矩形にマージンを加えた新しい矩形を返す
     * Illustrator の artboardRect は [left, top, right, bottom]（上が大・下が小）。
     * @param {number[]} rect - 元の矩形
     * @param {number} verticalMarginPt - 上下のマージン（pt）
     * @param {number} horizontalMarginPt - 左右のマージン（pt）
     * @returns {number[]} マージンを加えた矩形
     */
    function expandRectByMargin(rect, verticalMarginPt, horizontalMarginPt) {
        return [
            rect[0] - horizontalMarginPt,
            rect[1] + verticalMarginPt,
            rect[2] + horizontalMarginPt,
            rect[3] - verticalMarginPt
        ];
    }

    /**
     * アートボード矩形をピクセルグリッドに最適化する（X/Y/W/H を各1回だけ整数化）
     * X(左)・Y(上)・幅・高さの4値をそれぞれ整数に丸め、右下は X+幅 / Y−高さ で再構成する（二重丸めしない）。
     * @param {number[]} rect - 元の矩形
     * @returns {number[]} 丸めた矩形
     */
    function snapRectToPixelGrid(rect) {
        var x = Math.round(rect[0]);
        var y = Math.round(rect[1]);
        var width = Math.round(rect[2] - rect[0]);
        var height = Math.round(rect[1] - rect[3]);
        return [x, y, x + width, y - height];
    }

    /**
     * 矩形の X/Y/W/H を指定単位で各1回だけ整数化する
     * 単位変換に失敗した場合はピクセルグリッドにフォールバック。
     * @param {number[]} rect - 元の矩形
     * @param {string} unit - 単位の文字列
     * @returns {number[]} 丸めた矩形
     */
    function snapRectToUnitGrid(rect, unit) {
        var ptPerUnit = toPt(1, unit);
        if (isNaN(ptPerUnit) || ptPerUnit === 0) return snapRectToPixelGrid(rect);
        var x = Math.round(rect[0] / ptPerUnit) * ptPerUnit;
        var y = Math.round(rect[1] / ptPerUnit) * ptPerUnit;
        var width = Math.round((rect[2] - rect[0]) / ptPerUnit) * ptPerUnit;
        var height = Math.round((rect[1] - rect[3]) / ptPerUnit) * ptPerUnit;
        return [x, y, x + width, y - height];
    }

    /**
     * 矩形の幅・高さが正か（left<right かつ bottom<top）
     * @param {number[]} rect - 矩形
     * @returns {boolean} 有効なら true
     */
    function isValidRect(rect) {
        return rect[2] > rect[0] && rect[1] > rect[3];
    }

    /**
     * 無効化した軸を元アートボードの座標に固定する（＝その軸は動かさない）
     * 横(左右)OFFなら left/right を、縦(上下)OFFなら top/bottom を元アートボード値に戻す。
     * @param {number[]} rect - 計算した矩形
     * @param {number[]} artboardRect - 元のアートボード矩形
     * @param {boolean} verticalEnabled - 上下（高さ）を調整するか
     * @param {boolean} horizontalEnabled - 左右（幅）を調整するか
     * @returns {number[]} 無効軸を固定した矩形
     */
    function lockDisabledAxes(rect, artboardRect, verticalEnabled, horizontalEnabled) {
        var lockedRect = rect.slice();
        if (!horizontalEnabled) { lockedRect[0] = artboardRect[0]; lockedRect[2] = artboardRect[2]; }
        if (!verticalEnabled) { lockedRect[1] = artboardRect[1]; lockedRect[3] = artboardRect[3]; }
        return lockedRect;
    }

    /**
     * 丸めモードに従って矩形を整数化する
     * "pixelGrid" = ピクセル整数、"currentUnit" = 現在の単位で整数、"none" = 丸めなし。
     * @param {number[]} rect - 元の矩形
     * @param {string} roundMode - 丸めモード
     * @param {string} unit - 単位の文字列
     * @returns {number[]} 丸めた矩形
     */
    function applyRounding(rect, roundMode, unit) {
        if (roundMode === "pixelGrid") return snapRectToPixelGrid(rect);
        if (roundMode === "currentUnit") return snapRectToUnitGrid(rect, unit);
        return rect; // "none"
    }

    /**
     * 2つの矩形が実質的に同一か判定する
     * マージン0などで値が変わらない場合は書き換えを省き、取り消し履歴を増やさないために使う。
     * @param {number[]} rectA - 矩形A
     * @param {number[]} rectB - 矩形B
     * @returns {boolean} 同一なら true
     */
    function rectsEqual(rectA, rectB) {
        if (!rectA || !rectB) return false;
        for (var i = 0; i < 4; i++) {
            if (Math.abs(rectA[i] - rectB[i]) > 0.0001) return false;
        }
        return true;
    }

    /**
     * オブジェクトのバウンディングボックスを取得する
     * usePreviewBounds=true なら visibleBounds（塗り/線を含む）、false なら geometricBounds（パス外形のみ）。
     * @param {PageItem} item - 対象
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]} [left, top, right, bottom]
     */
    function getItemBounds(item, usePreviewBounds) {
        return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
    }

    /**
     * 複数アイテムの外接バウンディングボックスを取得する
     * @param {PageItem[]} items - 対象
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]|null} 外接矩形。空なら null
     */
    function getUnionBounds(items, usePreviewBounds) {
        if (!items || items.length === 0) return null;
        var unionBounds = getItemBounds(items[0], usePreviewBounds);
        for (var i = 1; i < items.length; i++) {
            var itemBounds = getItemBounds(items[i], usePreviewBounds);
            unionBounds[0] = Math.min(unionBounds[0], itemBounds[0]);
            unionBounds[1] = Math.max(unionBounds[1], itemBounds[1]);
            unionBounds[2] = Math.max(unionBounds[2], itemBounds[2]);
            unionBounds[3] = Math.min(unionBounds[3], itemBounds[3]);
        }
        return unionBounds;
    }

    /**
     * 計測可能なページアイテムか（geometricBounds を持つか）
     * TextRange 等の非ページアイテムを除外し、選択の型不整合を防ぐ。
     * @param {Object} item - 選択の要素
     * @returns {boolean} 計測できれば true
     */
    function isMeasurableItem(item) {
        if (!item) return false;
        /* TextRange などは geometricBounds で例外になる / non-page items throw on geometricBounds */
        try {
            var bounds = item.geometricBounds;
            return (bounds && bounds.length === 4);
        } catch (e) {
            return false;
        }
    }

    /**
     * クリップグループのクリッピングパスを返す
     * @param {GroupItem} groupItem - クリップグループ
     * @returns {PageItem|null} クリッピングパス。無ければ null
     */
    function getClippingPath(groupItem) {
        try {
            for (var j = 0; j < groupItem.pageItems.length; j++) {
                if (groupItem.pageItems[j].clipping) return groupItem.pageItems[j];
            }
        } catch (e) { /* ignore */ }
        return null;
    }

    /**
     * 選択アイテムを正規化する
     * ・計測できない要素（TextRange 等）は除外
     * ・クリップグループはクリッピングパスのみを採用、それ以外はそのまま
     * @param {Object[]} items - 選択の要素
     * @returns {PageItem[]} 計測に使うアイテム
     */
    function collectEffectiveItems(items) {
        var effectiveItems = [];
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!isMeasurableItem(item)) continue; // 非ページアイテムをスキップ / skip non-page items
            if (item.typename === "GroupItem" && item.clipped) {
                var clippingPath = getClippingPath(item);
                if (clippingPath) effectiveItems.push(clippingPath);
            } else {
                effectiveItems.push(item);
            }
        }
        return effectiveItems;
    }

    /**
     * 計測対象を再帰的に収集する
     * ・TextFrame は複製をアウトライン化（元は不変）。グループ内テキストも対象。
     * ・クリップグループはクリッピングパスのみ。通常グループは中身へ再帰。
     * 一時複製・アウトラインは tempObjects に登録し、呼び出し側で必ず削除する。
     * @param {PageItem[]} items - 対象
     * @param {PageItem[]} tempObjects - 一時オブジェクトの登録先
     * @param {PageItem[]} measureItems - 計測対象の格納先
     * @returns {void}
     */
    function collectMeasureTargets(items, tempObjects, measureItems) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!isMeasurableItem(item)) continue;
            if (item.typename === "TextFrame") {
                // 複製を先に登録 → アウトライン化（失敗時も複製を削除できる） / track clone before outlining
                var clone = item.duplicate();
                tempObjects.push(clone);
                var outlined = clone.createOutline(); // GroupItem を返す / returns a GroupItem
                tempObjects.push(outlined);
                measureItems.push(outlined);
            } else if (item.typename === "GroupItem" && item.clipped) {
                var clippingPath = getClippingPath(item);
                if (clippingPath) measureItems.push(clippingPath);
            } else if (item.typename === "GroupItem") {
                // 通常グループは中身を再帰（ネストされたテキストも非破壊計測） / recurse into groups
                collectMeasureTargets(item.pageItems, tempObjects, measureItems);
            } else {
                measureItems.push(item);
            }
        }
    }

    /**
     * 選択オブジェクトの外接境界を取得する
     * テキストは複製をアウトライン化して計測し、計測後に一時オブジェクトを必ず削除する。
     * 元の TextFrame には一切触れないため、ID・重なり順・名前・タグ・ノート・Variable 等が保持される。
     * @param {PageItem[]} items - 対象
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]|null} 外接矩形。空なら null
     */
    function measureSelectionBounds(items, usePreviewBounds) {
        var tempObjects = []; // 計測用に作った一時複製・アウトライン（必ず削除） / temp objects to remove
        try {
            var measureItems = [];
            collectMeasureTargets(items, tempObjects, measureItems);
            return getUnionBounds(measureItems, usePreviewBounds);
        } finally {
            // 途中で例外が起きても一時オブジェクトは必ず削除 / always remove temp objects, even on error
            for (var k = 0; k < tempObjects.length; k++) {
                /* アウトライン化で消費された複製の remove() は例外になる / a clone consumed by createOutline() throws on remove() */
                try { tempObjects[k].remove(); } catch (e) { }
            }
        }
    }

    /**
     * 計測に使えるアイテムか（ロック・非表示・ガイド、およびそのレイヤーを除外）
     * @param {PageItem} item - 対象
     * @returns {boolean} 使えるなら true
     */
    function isUsableItem(item) {
        try {
            if (!item || item.locked || item.hidden || item.guides) return false;
            var parentLayer = item.parent;
            while (parentLayer && parentLayer.typename === "Layer") {
                if (!parentLayer.visible || parentLayer.locked) return false;
                parentLayer = parentLayer.parent;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * アートボードに重なるページアイテムを収集する
     * レイヤー直下のアイテムだけを見る（グループの中身はグループごと1件として扱う）。
     * @param {number[]} artboardRect - アートボード矩形
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {PageItem[]} 重なるアイテム
     */
    function getItemsInArtboard(artboardRect, usePreviewBounds) {
        var doc = app.activeDocument;
        var overlappingItems = [];
        for (var i = 0; i < doc.pageItems.length; i++) {
            var item = doc.pageItems[i];
            try {
                if (item.parent.typename !== "Layer") continue; // 入れ子はグループと一緒に扱う / nested items travel with their group
                if (!isUsableItem(item)) continue;
                var bounds = getItemBounds(item, usePreviewBounds);
                // 一辺でも外れていれば非交差 / no overlap when any edge clears the artboard
                if (bounds[2] <= artboardRect[0] || bounds[0] >= artboardRect[2] || bounds[3] >= artboardRect[1] || bounds[1] <= artboardRect[3]) continue;
                overlappingItems.push(item);
            } catch (e) { /* ignore */ }
        }
        return overlappingItems;
    }

    /**
     * アートボード内オブジェクトの外接境界を取得する
     * @param {number[]} artboardRect - アートボード矩形
     * @param {boolean} usePreviewBounds - プレビュー境界を使うか
     * @returns {number[]|null} 外接矩形。対象が無ければ null
     */
    function measureArtboardContentBounds(artboardRect, usePreviewBounds) {
        var items = getItemsInArtboard(artboardRect, usePreviewBounds);
        if (items.length === 0) return null;
        return measureSelectionBounds(items, usePreviewBounds);
    }

    // =========================================
    // アートボード矩形の計算 / Artboard rect planning
    // =========================================
    // プレビューと確定で同じ計算（丸め・無効軸固定を含む）を使い、プレビューと結果を一致させる。
    // marginSettings: { verticalPt, horizontalPt, roundMode, unit, verticalEnabled, horizontalEnabled }

    /**
     * マージン適用の共通パイプライン：拡張 → 丸め → 無効軸を元座標に固定
     * @param {number[]} baseRect - 基準の矩形（アートボード自身、またはオブジェクトの外接）
     * @param {number[]} artboardOriginalRect - 実行時のアートボード矩形（無効軸の固定用）
     * @param {Object} marginSettings - マージン適用の設定
     * @returns {number[]} 新しいアートボード矩形
     */
    function computeMarginRect(baseRect, artboardOriginalRect, marginSettings) {
        var rect = expandRectByMargin(baseRect, marginSettings.verticalPt, marginSettings.horizontalPt);
        rect = applyRounding(rect, marginSettings.roundMode, marginSettings.unit);
        return lockDisabledAxes(rect, artboardOriginalRect, marginSettings.verticalEnabled, marginSettings.horizontalEnabled);
    }

    /**
     * 全アートボードの矩形を控える
     * @param {Artboards} artboards - アートボードのコレクション
     * @returns {number[][]} 矩形の配列
     */
    function snapshotArtboardRects(artboards) {
        var rects = [];
        for (var i = 0; i < artboards.length; i++) {
            rects.push(artboards[i].artboardRect.slice());
        }
        return rects;
    }

    /**
     * 操作と対象から、書き換えるアートボードの番号と新しい矩形を求める
     * ・拡張：アートボード自身の矩形にマージンを加減
     * ・合わせる×すべて：各アートボードを、その内側のオブジェクトに合わせる（対象が無いアートボードは据え置き）
     * ・合わせる×現在：現在のアートボードを選択の外接に合わせる
     * @param {string} operation - "fit" / "expand"
     * @param {string} scope - "current" / "all"
     * @param {number} activeIndex - 現在のアートボードの番号
     * @param {number[][]} originalRects - 実行時のアートボード矩形
     * @param {Object} boundsProvider - getContentBounds(index) と getSelectionBounds() を持つ計測元
     * @param {Object} marginSettings - マージン適用の設定
     * @returns {Object[]} { index, rect } の配列
     */
    function planArtboardRects(operation, scope, activeIndex, originalRects, boundsProvider, marginSettings) {
        var targetIndexes = [];
        if (scope === "all") {
            for (var i = 0; i < originalRects.length; i++) targetIndexes.push(i);
        } else {
            targetIndexes.push(activeIndex);
        }

        var rectPlans = [];
        for (var j = 0; j < targetIndexes.length; j++) {
            var artboardIndex = targetIndexes[j];
            var baseRect;
            if (operation === "expand") {
                baseRect = originalRects[artboardIndex];
            } else if (scope === "all") {
                baseRect = boundsProvider.getContentBounds(artboardIndex);
            } else {
                baseRect = boundsProvider.getSelectionBounds();
            }
            if (!baseRect) continue; // 計測対象が無い / nothing to measure
            rectPlans.push({ index: artboardIndex, rect: computeMarginRect(baseRect, originalRects[artboardIndex], marginSettings) });
        }
        return rectPlans;
    }

    /**
     * プレビューとしてアートボード矩形を書き換える
     * 無効な矩形と変化の無い矩形は書き換えない（取り消し履歴のノイズを減らす）。
     * @param {Object[]} rectPlans - planArtboardRects() の結果
     * @returns {void}
     */
    function previewArtboardRects(rectPlans) {
        var artboards = app.activeDocument.artboards;
        for (var i = 0; i < rectPlans.length; i++) {
            var rectPlan = rectPlans[i];
            if (isValidRect(rectPlan.rect) && !rectsEqual(artboards[rectPlan.index].artboardRect, rectPlan.rect)) {
                artboards[rectPlan.index].artboardRect = rectPlan.rect;
            }
        }
    }

    // =========================================
    // 入力ユーティリティ / Input utilities
    // =========================================

    /**
     * edittext に矢印キーでの増減を付与する
     * ↑↓で±1、Shift+↑↓で10の倍数にスナップ（例 36→40 / 36→30）、Option(Alt)+↑↓で±0.1。
     * @param {EditText} editText - 対象の入力欄
     * @param {function} onUpdate - 値を変えたあとに呼ぶ関数（新しい文字列を受け取る）
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onUpdate) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            // 修飾キーは event から読む（keyboardState は macOS で誤報あり）。取得不可時のみフォールバック。
            var keyboardState = ScriptUI.environment.keyboardState;
            var isShiftPressed = (event.shiftKey !== undefined) ? event.shiftKey : keyboardState.shiftKey;
            var isOptionPressed = (event.altKey !== undefined) ? event.altKey : keyboardState.altKey;

            if (event.keyName == "Up" || event.keyName == "Down") {
                var isUp = (event.keyName == "Up");
                if (isShiftPressed) {
                    // 10の倍数にスナップ（倍数上ならさらに±10） / snap to the next multiple of 10
                    value = isUp ? (Math.floor(value / 10) * 10 + 10) : (Math.ceil(value / 10) * 10 - 10);
                } else {
                    var step = isOptionPressed ? 0.1 : 1;
                    value = value + (isUp ? step : -step);
                }
                // 浮動小数の誤差を丸める（0.1刻み対応） / trim float error for 0.1 steps
                value = Math.round(value * 10000) / 10000;

                event.preventDefault();
                editText.text = value;
                if (typeof onUpdate === "function") onUpdate(editText.text);
            }
        });
    }

    // =========================================
    // ダイアログの構築 / Dialog construction
    // =========================================

    /**
     * 調整基準パネルに「項目名＋コントロール群」の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelSet - 項目名のラベル
     * @param {string} contentOrientation - コントロール群の並び（"column" = 縦 / "row" = 横）
     * @returns {Group} コントロールを入れるグループ
     */
    function addBasisRow(parentPanel, labelSet, contentOrientation) {
        var isColumn = (contentOrientation === "column");
        var basisRow = parentPanel.add("group");
        basisRow.orientation = "row";
        basisRow.alignChildren = ["left", isColumn ? "top" : "center"];
        var rowLabel = basisRow.add("statictext", undefined, labelText(labelSet));
        rowLabel.preferredSize.width = BASIS_LABEL_WIDTHS[uiLang];
        var contentGroup = basisRow.add("group");
        contentGroup.orientation = contentOrientation;
        contentGroup.alignChildren = isColumn ? "left" : ["left", "center"];
        return contentGroup;
    }

    /**
     * 調整基準パネル（操作＋対象＋サイズ）を作る
     * @param {Window} marginDialog - ダイアログ
     * @param {Object} initialSettings - ダイアログの初期値
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildBasisPanel(marginDialog, initialSettings, dialogControls) {
        var basisPanel = marginDialog.add("panel", undefined, getLabel(LABELS.panel.adjustmentBasis));
        setupPanel(basisPanel);

        /* 操作：オブジェクトに合わせる / アートボードを拡張（ラジオは縦並び） / Operation group (vertical radios) */
        var operationGroup = addBasisRow(basisPanel, LABELS.fieldLabel.operation, "column");
        dialogControls.fitRadio = operationGroup.add("radiobutton", undefined, getLabel(LABELS.radio.fit));
        dialogControls.fitRadio.helpTip = getLabel(LABELS.tooltip.fit);
        dialogControls.expandRadio = operationGroup.add("radiobutton", undefined, getLabel(LABELS.radio.expand));
        dialogControls.expandRadio.helpTip = getLabel(LABELS.tooltip.expand);

        /* 対象：現在のアートボード / すべてのアートボード（ラジオは縦並び） / Scope group (vertical radios) */
        var scopeGroup = addBasisRow(basisPanel, LABELS.fieldLabel.scope, "column");
        dialogControls.currentRadio = scopeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.currentArtboard));
        dialogControls.currentRadio.helpTip = getLabel(LABELS.tooltip.currentArtboard);
        dialogControls.allRadio = scopeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.allArtboards));
        dialogControls.allRadio.helpTip = getLabel(LABELS.tooltip.allArtboards);

        /* サイズ：幅／高さ（OFFにした方は実行時のサイズのまま固定） / Axis targets: width & height (off keeps the original size) */
        var axisGroup = addBasisRow(basisPanel, LABELS.fieldLabel.size, "row");
        axisGroup.spacing = COLUMN_SPACING;
        dialogControls.widthCheckbox = axisGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.width));
        dialogControls.widthCheckbox.value = initialSettings.horizontalEnabled;
        dialogControls.widthCheckbox.helpTip = getLabel(LABELS.tooltip.axisEnable);
        dialogControls.heightCheckbox = axisGroup.add("checkbox", undefined, getLabel(LABELS.checkbox.height));
        dialogControls.heightCheckbox.value = initialSettings.verticalEnabled;
        dialogControls.heightCheckbox.helpTip = getLabel(LABELS.tooltip.axisEnable);

        /* 初期値を適用 / apply initial selection */
        dialogControls.fitRadio.value = (initialSettings.operation === "fit");
        dialogControls.expandRadio.value = (initialSettings.operation !== "fit");
        dialogControls.currentRadio.value = (initialSettings.scope === "current");
        dialogControls.allRadio.value = (initialSettings.scope === "all");
    }

    /**
     * マージンの入力行（項目名＋入力欄）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {Object} labelSet - 項目名のラベル
     * @param {string} initialText - 入力欄の初期値
     * @returns {{label: StaticText, input: EditText}} 項目名と入力欄
     */
    function addMarginField(parentGroup, labelSet, initialText) {
        var fieldRow = parentGroup.add("group");
        fieldRow.orientation = "row";
        fieldRow.alignChildren = ["left", "center"];
        var fieldLabel = fieldRow.add("statictext", undefined, getLabel(labelSet));
        fieldLabel.justify = "right";
        fieldLabel.preferredSize.width = MARGIN_LABEL_WIDTHS[uiLang]; /* 入力欄の位置を揃える / line up the inputs */
        var fieldInput = fieldRow.add("edittext", undefined, initialText);
        fieldInput.characters = MARGIN_FIELD_CHARACTERS;
        fieldInput.helpTip = getLabel(LABELS.tooltip.marginInput);
        return { label: fieldLabel, input: fieldInput };
    }

    /**
     * マージン入力パネル（入力欄＋連動・プレビュー境界の2カラム）を作る
     * @param {Window} marginDialog - ダイアログ
     * @param {string} rulerUnit - 定規単位
     * @param {Object} initialSettings - ダイアログの初期値
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildMarginPanel(marginDialog, rulerUnit, initialSettings, dialogControls) {
        var marginPanel = marginDialog.add("panel", undefined, getLabel(LABELS.panel.margin) + " (" + rulerUnit + ")");
        setupPanel(marginPanel);
        marginPanel.orientation = "row";
        marginPanel.alignChildren = ["left", "top"];
        marginPanel.spacing = COLUMN_SPACING;

        var marginFieldsColumn = marginPanel.add("group");
        marginFieldsColumn.orientation = "column";
        marginFieldsColumn.alignChildren = ["left", "center"];

        var linkColumn = marginPanel.add("group");
        linkColumn.orientation = "column";
        linkColumn.alignChildren = ["left", "center"];
        linkColumn.alignment = ["left", "center"];

        /* 上下マージン（「高さ」OFFで無効）・左右マージン（「幅」OFFで無効） / Vertical & horizontal margin inputs */
        var verticalField = addMarginField(marginFieldsColumn, LABELS.fieldLabel.vertical, initialSettings.marginV);
        dialogControls.verticalLabel = verticalField.label;
        dialogControls.verticalInput = verticalField.input;
        var horizontalField = addMarginField(marginFieldsColumn, LABELS.fieldLabel.horizontal,
            initialSettings.link ? initialSettings.marginV : initialSettings.marginH);
        dialogControls.horizontalLabel = horizontalField.label;
        dialogControls.horizontalInput = horizontalField.input;

        /* 連動チェックボックス / Linked checkbox */
        dialogControls.linkCheckbox = linkColumn.add("checkbox", undefined, getLabel(LABELS.checkbox.linked));
        dialogControls.linkCheckbox.value = initialSettings.link;
        dialogControls.linkCheckbox.helpTip = getLabel(LABELS.tooltip.linked);

        /* プレビュー境界（visibleBounds を採用するか。「合わせる」のときのみ有効） / use visibleBounds; only for fit */
        dialogControls.previewBoundsCheckbox = linkColumn.add("checkbox", undefined, getLabel(LABELS.checkbox.previewBounds));
        dialogControls.previewBoundsCheckbox.value = initialSettings.previewBounds;
        dialogControls.previewBoundsCheckbox.helpTip = getLabel(LABELS.tooltip.previewBounds);
        dialogControls.previewBoundsCheckbox.enabled = (initialSettings.operation === "fit");
    }

    /**
     * アートボードサイズの微調整パネル（丸めモード）を作る
     * プレビューにも確定と同じ丸めを反映する。
     * @param {Window} marginDialog - ダイアログ
     * @param {Object} initialSettings - ダイアログの初期値
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildFineTuningPanel(marginDialog, initialSettings, dialogControls) {
        var fineTuningPanel = marginDialog.add("panel", undefined, getLabel(LABELS.panel.fineTuning));
        setupPanel(fineTuningPanel);

        var roundModeGroup = fineTuningPanel.add("group");
        roundModeGroup.orientation = "column";
        roundModeGroup.alignChildren = "left";
        dialogControls.roundPixelRadio = roundModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.roundPixelGrid));
        dialogControls.roundUnitRadio = roundModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.roundCurrentUnit));
        dialogControls.roundNoneRadio = roundModeGroup.add("radiobutton", undefined, getLabel(LABELS.radio.roundNone));
        dialogControls.roundPixelRadio.helpTip = getLabel(LABELS.tooltip.roundPixelGrid);
        dialogControls.roundUnitRadio.helpTip = getLabel(LABELS.tooltip.roundCurrentUnit);
        dialogControls.roundNoneRadio.helpTip = getLabel(LABELS.tooltip.roundNone);
        dialogControls.roundUnitRadio.value = (initialSettings.roundMode === "currentUnit");
        dialogControls.roundNoneRadio.value = (initialSettings.roundMode === "none");
        dialogControls.roundPixelRadio.value = !dialogControls.roundUnitRadio.value && !dialogControls.roundNoneRadio.value; // 既定 / default
    }

    /**
     * ボタン行（左右中央：Cancel → OK の順）を作る
     * @param {Window} marginDialog - ダイアログ
     * @param {Object} dialogControls - 作ったコントロールの格納先
     * @returns {void}
     */
    function buildButtonRow(marginDialog, dialogControls) {
        var btnRowGroup = marginDialog.add("group");
        setupRow(btnRowGroup, "center");
        btnRowGroup.alignChildren = ["center", "center"];
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
        dialogControls.btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        dialogControls.btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
    }

    // =========================================
    // ダイアログの状態 / Dialog state
    // =========================================

    /**
     * 選んでいる操作を返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "fit" / "expand"
     */
    function getOperation(dialogControls) {
        return dialogControls.fitRadio.value ? "fit" : "expand";
    }

    /**
     * 選んでいる対象を返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "current" / "all"
     */
    function getScope(dialogControls) {
        return dialogControls.allRadio.value ? "all" : "current";
    }

    /**
     * 選んでいる丸めモードを返す
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {string} "pixelGrid" / "currentUnit" / "none"
     */
    function getRoundMode(dialogControls) {
        if (dialogControls.roundPixelRadio.value) return "pixelGrid";
        return dialogControls.roundUnitRadio.value ? "currentUnit" : "none";
    }

    /**
     * 「合わせる」で選択が無いときは対象を「すべて」に固定してディムする
     * @param {Object} dialogControls - ダイアログのコントロール
     * @param {boolean} hasSelection - 計測できる選択があるか
     * @returns {void}
     */
    function refreshScopeState(dialogControls, hasSelection) {
        var forceAll = (getOperation(dialogControls) === "fit" && !hasSelection);
        if (forceAll) {
            dialogControls.allRadio.value = true;
            dialogControls.currentRadio.value = false;
        }
        dialogControls.currentRadio.enabled = !forceAll;
    }

    /**
     * 入力欄と連動チェックの有効/無効を現在の状態から更新する
     * 上下=「高さ」ON、左右=「幅」ON かつ 非連動、連動=幅・高さが両方ONのときだけ有効。
     * @param {Object} dialogControls - ダイアログのコントロール
     * @returns {void}
     */
    function refreshMarginInputStates(dialogControls) {
        var heightOn = dialogControls.heightCheckbox.value;
        var widthOn = dialogControls.widthCheckbox.value;
        dialogControls.verticalLabel.enabled = heightOn;
        dialogControls.verticalInput.enabled = heightOn;
        dialogControls.horizontalLabel.enabled = widthOn;
        dialogControls.horizontalInput.enabled = widthOn && !dialogControls.linkCheckbox.value;
        // 幅・高さのどちらかがOFFなら連動は使えない（自動OFFのうえディム） / disable link when either axis is off
        dialogControls.linkCheckbox.enabled = heightOn && widthOn;
    }

    /**
     * マージン欄の文字列を pt に変換する。無効な軸は 0 とみなす
     * @param {boolean} axisEnabled - その軸を調整するか
     * @param {string} marginText - 入力欄の文字列
     * @param {string} unit - 単位の文字列
     * @returns {number} pt。数値でなければ NaN
     */
    function marginTextToPoints(axisEnabled, marginText, unit) {
        return axisEnabled ? toPt(parseFloat(marginText), unit) : 0;
    }

    /**
     * ダイアログの入力から、マージン適用の設定を作る
     * @param {Object} dialogControls - ダイアログのコントロール
     * @param {string} rulerUnit - 定規単位
     * @returns {Object} マージン適用の設定。valid は有効な軸の値がすべて数値のとき true
     */
    function readMarginSettings(dialogControls, rulerUnit) {
        var verticalEnabled = dialogControls.heightCheckbox.value;
        var horizontalEnabled = dialogControls.widthCheckbox.value;
        var verticalPt = marginTextToPoints(verticalEnabled, dialogControls.verticalInput.text, rulerUnit);
        var horizontalPt = marginTextToPoints(horizontalEnabled, dialogControls.horizontalInput.text, rulerUnit);
        return {
            valid: !isNaN(verticalPt) && !isNaN(horizontalPt),
            verticalPt: verticalPt,
            horizontalPt: horizontalPt,
            roundMode: getRoundMode(dialogControls),
            unit: rulerUnit,
            verticalEnabled: verticalEnabled,
            horizontalEnabled: horizontalEnabled
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * マージン入力ダイアログを表示し設定を返す（ライブプレビュー付き）
     * @param {string} defaultMargin - 保存が無いときのマージン
     * @param {string} rulerUnit - 定規単位
     * @param {number} artboardCount - アートボードの数
     * @param {boolean} hasSelection - 計測できる選択があるか
     * @param {PageItem[]} selectionItems - ダイアログ表示時に固定した選択アイテム（「合わせる」で使用）
     * @returns {Object|null} { operation, scope, previewBounds, marginSettings }。キャンセル時は null
     */
    function showMarginDialog(defaultMargin, rulerUnit, artboardCount, hasSelection, selectionItems) {
        var marginDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);

        // ダイアログ位置の記憶（保存関数を受け取る） / wire position persistence, get the saver
        var persistDialogLocation = attachPositionPersistence(marginDialog, DIALOG_POSITION_KEY, DIALOG_FIRST_RUN_OFFSET_X);
        marginDialog.opacity = DIALOG_OPACITY;
        setupWindow(marginDialog);

        // 保存済み設定（セッション内）から初期値を解決 / resolve initial values from stored settings
        var initialSettings = resolveInitialSettings(defaultMargin, artboardCount, hasSelection);

        var dialogControls = {};
        buildBasisPanel(marginDialog, initialSettings, dialogControls);
        buildMarginPanel(marginDialog, rulerUnit, initialSettings, dialogControls);
        buildFineTuningPanel(marginDialog, initialSettings, dialogControls);
        buildButtonRow(marginDialog, dialogControls);
        refreshScopeState(dialogControls, hasSelection);
        refreshMarginInputStates(dialogControls);

        var verticalInput = dialogControls.verticalInput;
        var horizontalInput = dialogControls.horizontalInput;
        var widthCheckbox = dialogControls.widthCheckbox;
        var heightCheckbox = dialogControls.heightCheckbox;
        var linkCheckbox = dialogControls.linkCheckbox;
        var previewBoundsCheckbox = dialogControls.previewBoundsCheckbox;

        /* 現在／全アートボードの rect を保存（プレビュー復元用） / Save artboard rects for preview restore */
        var artboards = app.activeDocument.artboards;
        var activeArtboardIndex = artboards.getActiveArtboardIndex();
        var originalArtboardRects = snapshotArtboardRects(artboards);

        /* プレビュー復元：開いた時点の全 artboardRect へ書き戻す（undo回数に依存しない） / Restore all artboard rects to the opening snapshot */
        var previewManager = new PreviewManager(function () {
            var currentArtboards = app.activeDocument.artboards;
            for (var i = 0; i < currentArtboards.length && i < originalArtboardRects.length; i++) {
                if (!rectsEqual(currentArtboards[i].artboardRect, originalArtboardRects[i])) {
                    currentArtboards[i].artboardRect = originalArtboardRects[i];
                }
            }
        });

        /* 計測結果のキャッシュ（元の矩形と固定した選択で計測するのでダイアログ表示中は不変）
           Measured bounds cache; stable while the dialog is open */
        var boundsCache = {};
        var previewBoundsProvider = {
            getContentBounds: function (index) {
                var cacheKey = index + (previewBoundsCheckbox.value ? ":v" : ":g");
                if (!boundsCache.hasOwnProperty(cacheKey)) {
                    boundsCache[cacheKey] = measureArtboardContentBounds(originalArtboardRects[index], previewBoundsCheckbox.value);
                }
                return boundsCache[cacheKey];
            },
            getSelectionBounds: function () {
                var cacheKey = previewBoundsCheckbox.value ? "v" : "g";
                if (!boundsCache.hasOwnProperty(cacheKey)) {
                    boundsCache[cacheKey] = (selectionItems && selectionItems.length > 0) ?
                        measureSelectionBounds(selectionItems, previewBoundsCheckbox.value) : null;
                }
                return boundsCache[cacheKey];
            }
        };

        /* プレビュー更新：直前分を rollback してから最新状態を1回だけ適用 / Refresh preview via PreviewManager */
        function updatePreview() {
            previewManager.rollback();

            var marginSettings = readMarginSettings(dialogControls, rulerUnit);
            if (!marginSettings.valid) return;
            var operation = getOperation(dialogControls);
            var scope = getScope(dialogControls);

            previewManager.addStep(function () {
                previewArtboardRects(planArtboardRects(operation, scope, activeArtboardIndex, originalArtboardRects, previewBoundsProvider, marginSettings));
            });
        }

        /* 操作/対象の変更時：対象の固定・プレビュー境界の有効/無効を切り替えてプレビュー更新 / On change: refresh scope & previewBounds */
        function onBasisChange() {
            refreshScopeState(dialogControls, hasSelection); // 選択なしの「合わせる」は対象を「すべて」に固定 / fit without selection locks scope to all
            previewBoundsCheckbox.enabled = (getOperation(dialogControls) === "fit"); // 合わせる時のみ有効 / only for fit
            updatePreview();
        }

        /* 矢印キー・入力・ラジオ・連動のハンドラ登録 / Wire up input handlers */
        changeValueByArrowKey(verticalInput, function (newText) {
            if (linkCheckbox.value) horizontalInput.text = newText;
            updatePreview();
        });
        changeValueByArrowKey(horizontalInput, function () {
            if (linkCheckbox.value) return;
            updatePreview();
        });
        verticalInput.onChanging = function () {
            if (linkCheckbox.value) horizontalInput.text = verticalInput.text;
            updatePreview();
        };
        horizontalInput.onChanging = function () {
            if (linkCheckbox.value) return; // 連動中は水平の直接編集は無効
            updatePreview();
        };

        /* Option(Alt)クリック検出：mousedown で event.altKey を捕捉（onClick では修飾キーを取得できない） / capture Alt on mousedown */
        var axisSoloRequested = { width: false, height: false };
        widthCheckbox.addEventListener("mousedown", function (event) { axisSoloRequested.width = (event.altKey === true); });
        heightCheckbox.addEventListener("mousedown", function (event) { axisSoloRequested.height = (event.altKey === true); });

        /**
         * 幅・高さチェックの変更処理
         * Option+クリック＝ソロ（クリックした方のみON、もう片方OFF）。どちらかOFFなら連動を自動OFF。
         * @param {boolean} solo - Option+クリックか
         * @param {Checkbox} clickedCheckbox - クリックしたチェックボックス
         * @param {Checkbox} otherCheckbox - もう片方のチェックボックス
         * @returns {void}
         */
        function handleAxisToggle(solo, clickedCheckbox, otherCheckbox) {
            if (solo) {
                clickedCheckbox.value = true;   // クリックした軸をON / keep the clicked axis on
                otherCheckbox.value = false;    // もう片方をOFF / turn the other off
            }
            if (!heightCheckbox.value || !widthCheckbox.value) linkCheckbox.value = false;
            refreshMarginInputStates(dialogControls);
            updatePreview();
        }
        widthCheckbox.onClick = function () {
            var solo = axisSoloRequested.width; axisSoloRequested.width = false;
            handleAxisToggle(solo, widthCheckbox, heightCheckbox);
        };
        heightCheckbox.onClick = function () {
            var solo = axisSoloRequested.height; axisSoloRequested.height = false;
            handleAxisToggle(solo, heightCheckbox, widthCheckbox);
        };

        linkCheckbox.onClick = function () {
            if (linkCheckbox.value) {
                // 連動ON：幅・高さを有効に揃え、値を上下に統一 / linking re-enables both axes and mirrors V→H
                heightCheckbox.value = true;
                widthCheckbox.value = true;
                horizontalInput.text = verticalInput.text;
            }
            refreshMarginInputStates(dialogControls);
            updatePreview();
        };
        previewBoundsCheckbox.onClick = updatePreview;
        dialogControls.fitRadio.onClick = onBasisChange;
        dialogControls.expandRadio.onClick = onBasisChange;
        dialogControls.currentRadio.onClick = onBasisChange;
        dialogControls.allRadio.onClick = onBasisChange;
        dialogControls.roundPixelRadio.onClick = updatePreview;
        dialogControls.roundUnitRadio.onClick = updatePreview;
        dialogControls.roundNoneRadio.onClick = updatePreview;

        var dialogResult = null;
        dialogControls.btnOK.onClick = function () {
            var marginSettings = readMarginSettings(dialogControls, rulerUnit);
            if (!marginSettings.valid) {
                alert(getLabel(LABELS.alert.enterNumber));
                return;
            }
            dialogResult = {
                operation: getOperation(dialogControls),
                scope: getScope(dialogControls),
                previewBounds: previewBoundsCheckbox.value,
                marginSettings: marginSettings
            };
            // 設定をセッションに保存（次回の初期値に） / store settings for next run
            storeSettings({
                marginV: verticalInput.text,
                marginH: horizontalInput.text,
                link: linkCheckbox.value,
                verticalEnabled: marginSettings.verticalEnabled,
                horizontalEnabled: marginSettings.horizontalEnabled,
                previewBounds: dialogResult.previewBounds,
                roundMode: marginSettings.roundMode,
                operation: dialogResult.operation,
                scope: dialogResult.scope
            });
            persistDialogLocation();
            // プレビューを開いた時点へ復元してから閉じる / restore to the opening snapshot, then close
            previewManager.confirm();
            marginDialog.close(1);
        };
        dialogControls.btnCancel.onClick = function () {
            persistDialogLocation();
            // キャンセル時は必ずロールバックして閉じる / rollback preview and close
            previewManager.rollback();
            marginDialog.close(0);
        };

        updatePreview();

        // 開いたら有効な方のマージン入力欄にフォーカス（上下→無ければ左右） / focus a margin field on open
        if (verticalInput.enabled) {
            verticalInput.active = true;
        } else if (horizontalInput.enabled) {
            horizontalInput.active = true;
        }

        marginDialog.show();
        return dialogResult;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 対象を決定し、ダイアログの設定に従ってアートボードを調整する
     * @returns {void}
     */
    function main() {
        try {
            // ドキュメント未オープンなら分かりやすく案内して終了 / friendly guard when no document is open
            if (app.documents.length === 0) {
                alert(getLabel(LABELS.alert.noDocument));
                return;
            }
            var doc = app.activeDocument;

            // ダイアログ表示時点の選択を固定（「合わせる」で共通利用） / freeze selection at dialog time
            var selectionItems = collectEffectiveItems(doc.selection);
            var hasSelection = selectionItems.length > 0; // 計測可能な選択があるか / measurable selection?

            var artboards = doc.artboards;
            var rulerUnit = getRulerUnitString();

            /* 単位ごとの初期マージン値。選択なし・複数アートボード時は0（保存済み設定があればそちらが優先される）
               Default margin for the unit; 0 for the expand-all default (stored settings win) */
            var defaultMarginValue = (!hasSelection && artboards.length > 1) ? '0' : getDefaultMargin(rulerUnit);

            var dialogResult = showMarginDialog(defaultMarginValue, rulerUnit, artboards.length, hasSelection, selectionItems);
            if (!dialogResult) return; // キャンセル時は選択ツールにも切り替えない / cancel: leave tool unchanged

            var originalRects = snapshotArtboardRects(artboards);
            var previewBounds = dialogResult.previewBounds;
            var confirmBoundsProvider = {
                getContentBounds: function (index) { return measureArtboardContentBounds(originalRects[index], previewBounds); },
                getSelectionBounds: function () { return measureSelectionBounds(selectionItems, previewBounds); }
            };
            var rectPlans = planArtboardRects(dialogResult.operation, dialogResult.scope, artboards.getActiveArtboardIndex(),
                originalRects, confirmBoundsProvider, dialogResult.marginSettings);

            // 負マージン等で無効サイズになるものは適用しない / skip rects that end up invalid
            var skippedCount = 0;
            for (var i = 0; i < rectPlans.length; i++) {
                if (isValidRect(rectPlans[i].rect)) artboards[rectPlans[i].index].artboardRect = rectPlans[i].rect;
                else skippedCount++;
            }
            if (skippedCount > 0) alert(getLabel(LABELS.alert.marginTooLarge));

            // 適用が完了したときのみ選択ツールへ切り替え（キャンセル・エラー時は切り替えない） / switch tool only after a successful run
            app.selectTool("Adobe Select Tool");

        } catch (e) {
            $.writeln("[FitArtboardWithMargin] ERROR: " + formatError(e));
            alert(getLabel(LABELS.alert.errorOccurred) + formatError(e));
        }
    }

    main();

})();
