#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "DialogEngine"

/*

### スクリプト名：

FitArtboardWithMargin.jsx

### 概要

- 更新日：20260716
- アートボードのサイズを「操作」×「対象」の2軸で選んで自動調整します。
  - 操作：オブジェクトに合わせる（選択の外接＋マージン）／ アートボードを拡張（自身のサイズにマージン加減）
  - 対象：現在のアートボード ／ すべてのアートボード
  - ※「合わせる」は現在のアートボード専用（選択が必要。対象は「現在」に固定されます）
- 即時プレビュー付きダイアログで、確定結果と同じ計測・丸めをリアルタイムに確認できます。
- 定規単位（mm・px・pt・歯/Q など）に応じた初期マージン値を用意します。

### 主な機能

- 調整基準を「操作（合わせる／拡張）×対象（現在／すべて）」の2グループで指定
- 上下・左右マージンの個別入力・連動、各軸の個別ON/OFF（OFFの軸は実行時サイズを維持）
- サイズ丸めモードを3択（ピクセルグリッド／現在の単位で整数／なし）※X/Y/W/H を各1回だけ丸め
- テキストは複製をアウトライン化して非破壊で正確に計測（グループ内テキストも対象・元オブジェクトは不変）
- 確定と一致する即時プレビュー、キャンセル時はスナップショット復元で完全に元へ戻す
- ↑↓＝±1／Shift＝10の倍数にスナップ／Option＝±0.1 のキー入力、設定はセッション内で記憶

### 処理の流れ

1. 操作（合わせる／拡張）と対象（現在／すべてのアートボード）を選ぶ
   ※「合わせる」を選ぶと対象は「現在」に固定
2. 上下・左右マージンと丸めモードを設定（即時プレビュー対応）
3. OK で確定し、設定に基づきアートボードを自動調整

### オリジナル、謝辞

Gorolib Design
https://gorolib.blog.jp/archives/71820861.html

### オリジナルからの変更点

- ダイアログを閉じずにプレビュー更新（確定と同じ計測・丸めを反映）
- 単位系（mm、px など）によってデフォルト値を切り替え、歯/Q にも対応
- アートボードのX/Y/W/Hを整数化（ピクセル／現在単位／なしの3択）
- 「操作（合わせる／拡張）×対象（現在／すべて）」の2軸で調整基準を指定
- 空ドキュメントでもアートボード拡張が可能
- 上下・左右の個別ON/OFF（Option+クリックでソロ）、各軸の連動
- ↑↓／Shift＋↑↓／Option＋↑↓によるキー入力、ダイアログ位置と設定の記憶

### note

https://note.com/dtp_transit/n/n15d3c6c5a1e5

### 更新履歴

- v1.0 (20250420) : 初期バージョン
- v1.9.0 (20260715) : 丸めモード3択化・微調整パネル・UI共通化・プレビュー巻き戻し修正／テキスト計測を非破壊化（元テキスト不変）・プレビューと確定の計測ロジック統一・rulerType未対応値はptフォールバック・対象アイテム固定・Shift=±10/Option=±0.1・プレビュー境界は選択時のみ有効・丸めをX/Y/W/Hに統一・無効サイズ検証・文言改善・ラベル幅統一・ローカライズ共通化・ボタン右寄せ・上下/左右のマージンを個別ON/OFF（OFFで連動自動解除・実行時サイズを保持・Option+クリックでソロ）・空ドキュメントでもダイアログ表示・一時アウトラインをtry/finallyで確実削除・選択の型安全化（非ページアイテム除外）・選択ツール切替は成功時のみ・丸めをプレビューにも反映・ドキュメント未オープンガード・調整基準を「操作（合わせる/拡張）×対象（現在/すべて）」の2グループ化
- v1.9.1 (20260716) : 合わせるは現在のアートボード固定＋対象ディム・Shiftを10の倍数スナップに・開いたらマージン欄にフォーカス・グループ内テキストも非破壊計測（再帰）・一時複製の確実削除（createOutline失敗時も）・getMaxBounds空配列ガード・UIコロンを言語別に・英語コメント/概要を実装へ同期

---

### Script Name:

FitArtboardWithMargin.jsx

### Overview

- Last Updated: 20260716
- Automatically resizes artboards by choosing an operation and a scope:
  - Operation: Fit to objects (the selected objects' bounds plus margins) / Expand artboard (grow/shrink the artboard itself by the margins)
  - Scope: Current artboard / All artboards
  - Note: "Fit to objects" applies to the current artboard only (scope is locked to Current).
- The live-preview dialog reflects the exact same measurement and rounding as the final result.
- Provides unit-based default margins (mm, px, pt, Ha/Q, etc.).

### Main Features

- Basis split into two groups: operation (fit / expand) x scope (current / all)
- Vertical/horizontal margins with linking and per-axis on/off (a disabled axis keeps its original size)
- Three size-rounding modes (pixel grid / integers in current unit / none); X/Y/W/H rounded once each
- Non-destructive text measurement by outlining duplicates, including text nested in groups (originals stay untouched)
- Live preview that matches the final result; cancel fully restores via snapshot
- Key input: Arrow=±1 / Shift=snap to 10 / Option=±0.1; settings remembered within the session

### Workflow

1. Choose the operation (fit / expand) and scope (current / all artboards)
2. Set vertical/horizontal margins and rounding mode (with live preview)
3. Click OK to apply the adjustment

### Changelog

- v1.0 (20250420): Initial version
- v1.9.1 (20260716): Fit is current-artboard only (scope locked & dimmed), Shift snaps to multiples of 10, focus a margin field on open, non-destructive measurement of text nested in groups (recursive), reliable temp-duplicate cleanup (even if createOutline fails), getMaxBounds empty guard, language-aware UI colon, English comments/overview synced to the implementation
- v1.9.0 (20260715): Three-way rounding mode, fine-tuning panel, shared UI, preview rollback fix / non-destructive text measurement (originals untouched), unified preview & confirm measurement, pt fallback for unknown rulerType, frozen target items, Shift=±10/Option=±0.1, preview-bounds enabled only for selection, X/Y/W/H rounding, invalid-size guard, wording, aligned labels, shared localization, right-aligned buttons, per-axis on/off for vertical/horizontal margins (auto-unlinks, keeps original size, Option-click to solo), dialog shows on empty docs, temp outlines removed via try/finally, selection type-safety (non-page items excluded), tool switch only on success, rounding reflected in preview, no-document guard, basis split into two groups (operation: fit/expand × scope: current/all)

*/

// =========================================
// バージョン / Version
// =========================================
var SCRIPT_VERSION = "v1.9.1";

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================
    var CONFIG = {
        // rulerType → 単位。未対応の rulerType は pt にフォールバック / rulerType to unit; unknown falls back to pt
        rulerTypeToUnit: { 0: 'inch', 1: 'mm', 2: 'pt', 3: 'pica', 4: 'cm', 5: 'H', 6: 'px' },
        defaultMarginByUnit: {
            mm: '5',
            px: '20',
            pt: '10',
            _fallback: '0'
        },
        previewBoundsDefault: true,        // プレビュー境界(visibleBounds)を既定に / use visibleBounds by default
        roundModeDefault: 'pixelGrid',     // 既定の丸めモード（pixelGrid / currentUnit / none）/ default rounding mode
        linkDefault: true,                 // 上下左右の連動を既定ON / link margins by default
        dialogOpacity: 0.98,
        offsetX: 300
    };

    // =========================================
    // ローカライズ / Localize
    // =========================================

    /* 実行環境の言語を判定（ja / en） / Detect UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();
    var UI_COLON = (currentLanguage === "ja") ? "：" : ": "; // UIラベル用コロン（言語別）/ label colon per language

    /* ローカライズ文字列を取得（キー漏れ時は英語へフォールバック） / Get localized text, fallback to English */
    function getLocalizedText(entry) {
        if (!entry) return "";
        if (entry[currentLanguage] != null) return entry[currentLanguage];
        return (entry.en != null) ? entry.en : "";
    }

    var LABELS = {
        dialog: {
            title: { ja: "アートボードサイズを調整", en: "Adjust Artboard Size" }
        },
        panel: {
            target: { ja: "調整基準", en: "Adjustment basis" },
            margin: { ja: "マージン", en: "Margin" },
            options: { ja: "アートボードサイズの微調整", en: "Artboard size fine-tuning" }
        },
        // 操作（fit/expand）と対象（current/all）の2軸 / operation (fit/expand) & scope (current/all)
        radio: {
            fit: { ja: "オブジェクトに合わせる", en: "Fit to objects" },
            expand: { ja: "アートボードを拡張", en: "Expand artboard" },
            current: { ja: "現在のアートボード", en: "Current artboard" },
            all: { ja: "すべてのアートボード", en: "All artboards" }
        },
        groupLabel: {
            operation: { ja: "操作", en: "Operation" },
            scope: { ja: "対象", en: "Target" }
        },
        field: {
            vertical: { ja: "上下", en: "Vertical" },
            horizontal: { ja: "左右", en: "Horizontal" }
        },
        checkbox: {
            linked: { ja: "連動", en: "Linked" },
            previewBounds: { ja: "プレビュー境界", en: "Preview bounds" }
        },
        roundMode: {
            pixelGrid: { ja: "ピクセルグリッドに最適化", en: "Optimize to pixel grid" },
            currentUnit: { ja: "現在の単位で値を整数値に", en: "Round values in current unit" },
            none: { ja: "何もしない", en: "Do nothing" }
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
            opFit: {
                ja: "選択オブジェクトの外接＋マージンのサイズにアートボードを合わせます（選択が必要）",
                en: "Resize artboards to the selected objects' bounds plus margins (requires a selection)"
            },
            opExpand: {
                ja: "アートボード自身のサイズにマージンを加減します（マイナス値で縮小）",
                en: "Grow/shrink the artboards themselves by the margins (negative shrinks)"
            },
            scopeCurrent: {
                ja: "現在のアートボードのみを対象にします",
                en: "Apply to the current artboard only"
            },
            scopeAll: {
                ja: "すべてのアートボードを対象にします",
                en: "Apply to all artboards"
            },
            marginInput: {
                ja: "↑↓で±1、Shift+↑↓で10の倍数にスナップ、Option+↑↓で±0.1",
                en: "Arrow: ±1, Shift: snap to 10, Option: ±0.1"
            },
            axisEnable: {
                ja: "OFFにするとその方向は実行時のサイズのまま（連動は自動でOFF）。Option+クリックでこの軸だけON",
                en: "Off keeps this axis at its original size (auto-unlinks). Option-click to solo this axis"
            },
            link: {
                ja: "上下の値を左右にも自動で適用します",
                en: "Apply the vertical value to horizontal as well"
            },
            previewBounds: {
                ja: "ON：線・効果を含む見た目の境界（プレビュー境界）で計測／OFF：パスの幾何境界で計測",
                en: "On: measure with preview (visible) bounds incl. strokes/effects; Off: geometric path bounds"
            },
            roundPixel: {
                ja: "座標とサイズを整数ピクセルに丸めます",
                en: "Round position and size to integer pixels"
            },
            roundUnit: {
                ja: "現在の定規単位で座標とサイズを整数に丸めます",
                en: "Round position and size to integers in the current ruler unit"
            },
            roundNone: {
                ja: "丸めずに計測値のまま設定します",
                en: "Apply the measured values without rounding"
            }
        }
    };

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位ごとの初期マージン値を返す / Return default margin string for a unit */
    function getDefaultMargin(unit) {
        return CONFIG.defaultMarginByUnit.hasOwnProperty(unit) ?
            CONFIG.defaultMarginByUnit[unit] :
            CONFIG.defaultMarginByUnit._fallback;
    }

    /* 数値＋単位を pt に変換（失敗時は NaN） / Convert value+unit to points */
    function toPt(val, unit) {
        var n = Number(val);
        if (isNaN(n)) return NaN;
        // 歯/Q は UnitValue が非対応のため手計算（1H = 1Q = 0.25mm、1mm = 72/25.4pt） / H and Q are unsupported by UnitValue
        if (unit === "H" || unit === "Q") {
            return n * 0.25 * 72 / 25.4;
        }
        try {
            return new UnitValue(n, unit).as('pt');
        } catch (e) {
            return NaN;
        }
    }

    /* rulerType から単位文字列を得る（未対応は pt） / Map rulerType to a unit string, fallback to pt */
    function rulerUnitFromType(rulerType) {
        return CONFIG.rulerTypeToUnit.hasOwnProperty(rulerType) ? CONFIG.rulerTypeToUnit[rulerType] : 'pt';
    }

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */

    /* ウィンドウの共通設定 / Apply shared window layout */
    function setupWindow(win, spacing) {
        win.orientation = "column";
        win.alignChildren = "fill";
        win.margins = WINDOW_MARGINS;
        win.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 行グループの共通設定（ボタン列など） / Apply a horizontal row group */
    function setupRow(group, alignment, spacing) {
        group.orientation = "row";
        group.alignment = alignment || "left";
        group.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    // =========================================
    // エラー処理 / Error handling
    // =========================================

    /* Error を行番号・ファイル名付きで読みやすく整形 / Format an Error with line number */
    function formatError(e) {
        try {
            var msg = (e && e.message) ? String(e.message) : String(e);
            var ln = (e && e.line) ? (" line " + e.line) : "";
            var fn = (e && e.fileName) ? (" (" + e.fileName + ")") : "";
            return msg + ln + fn;
        } catch (_) {
            return String(e);
        }
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

        /* 変更操作を実行し、未復元フラグを立てる / Run a change and mark dirty */
        this.addStep = function (func) {
            try {
                func();
                this.dirty = true;
                app.redraw();
            } catch (e) {
                $.writeln("[PreviewManager] addStep error: " + e);
            }
        };

        /* プレビューを開いた時点の状態へ戻す / Restore to the pre-preview snapshot */
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
         * 現在の状態を確定する（OK時） / Confirm current state (on OK)
         * OK時は一度 rollback() で元に戻してから main() 側で本処理を1回だけ実行するため、ここでは rollback のみ。
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

    /* 保存済みのダイアログ位置を取得 / Get stored dialog location */
    function getStoredLocation(storageKey) {
        return $.global[storageKey] && $.global[storageKey].length === 2 ? $.global[storageKey] : null;
    }

    /* ダイアログ位置をセッションに保存 / Store dialog location */
    function storeLocation(storageKey, location) {
        $.global[storageKey] = [location[0], location[1]];
    }

    /* 位置を画面内に収める / Clamp a location within the screen */
    function clampLocationToScreen(location) {
        try {
            var visibleBounds = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0, 0, 1920, 1080];
            var clampedX = Math.max(visibleBounds[0] + 10, Math.min(location[0], visibleBounds[2] - 10));
            var clampedY = Math.max(visibleBounds[1] + 10, Math.min(location[1], visibleBounds[3] - 10));
            return [clampedX, clampedY];
        } catch (e) {
            return location;
        }
    }

    /* ダイアログ位置の記憶を設定し、保存関数を返す / Wire up dialog position persistence; return the saver
       保存位置があれば表示時に復元、無ければ初回はセンターからのオフセットで表示する。 */
    function attachPositionPersistence(dialog, positionKey, firstRunOffsetX) {
        if ($.global[positionKey] === undefined) $.global[positionKey] = null;
        var savedLocation = getStoredLocation(positionKey);

        var persist = function () {
            storeLocation(positionKey, [dialog.location[0], dialog.location[1]]);
        };

        if (savedLocation) {
            dialog.onShow = function () {
                dialog.location = clampLocationToScreen(savedLocation);
            };
        } else {
            dialog.onShow = function () {
                dialog.layout.layout(true);
                var screenWidth = $.screens[0].right - $.screens[0].left;
                var screenHeight = $.screens[0].bottom - $.screens[0].top;
                var centerX = screenWidth / 2 - dialog.bounds.width / 2;
                var centerY = screenHeight / 2 - dialog.bounds.height / 2;
                dialog.location = [centerX + firstRunOffsetX, centerY];
            };
        }

        dialog.onMove = persist;
        return persist;
    }

    // =========================================
    // 設定の記憶（セッション内） / Settings persistence (session only)
    // =========================================
    // $.global に設定を保持。#targetengine のためセッション中は保持されるが、再起動でリセット。
    // Kept in $.global; persists during the session but resets when Illustrator restarts.

    var SETTINGS_KEY = "__FitArtboardWithMargin_Settings";

    /* 保存済みの設定を取得（無ければ null） / Get stored settings */
    function getStoredSettings() {
        var stored = $.global[SETTINGS_KEY];
        return (stored && typeof stored === "object") ? stored : null;
    }

    /* 設定をセッションに保存 / Store settings */
    function storeSettings(settings) {
        $.global[SETTINGS_KEY] = settings;
    }

    /* 保存済み設定と文脈から、ダイアログの初期値を解決する / Resolve dialog initial values from stored settings + context
       operation: "fit"（オブジェクトに合わせる、要選択）/ "expand"（アートボードを拡張）
       scope: "current"（現在のアートボード）/ "all"（すべてのアートボード） */
    function resolveInitialSettings(defaultMargin, artboardCount, hasSelection) {
        var saved = getStoredSettings();

        // 操作：選択があれば fit、無ければ expand を既定に / operation default
        var operation = hasSelection ? "fit" : "expand";
        if (saved && (saved.operation === "fit" || saved.operation === "expand")) {
            operation = saved.operation;
        }
        if (operation === "fit" && !hasSelection) operation = "expand"; // fit は選択必須 / fit needs a selection

        // 対象：選択なし・複数アートボードなら all、それ以外は current を既定に / scope default
        var scope = (!hasSelection && artboardCount > 1) ? "all" : "current";
        if (saved && (saved.scope === "current" || saved.scope === "all")) {
            scope = saved.scope;
        }
        if (operation === "fit") scope = "current"; // 合わせるは現在のアートボード専用 / fit is current-only

        var link = (saved && typeof saved.link === "boolean") ? saved.link : CONFIG.linkDefault;
        // 連動ONのときは上下・左右とも有効に揃える / when linked, both axes are enabled
        var verticalEnabled = link ? true : ((saved && typeof saved.verticalEnabled === "boolean") ? saved.verticalEnabled : true);
        var horizontalEnabled = link ? true : ((saved && typeof saved.horizontalEnabled === "boolean") ? saved.horizontalEnabled : true);

        return {
            marginV: (saved && saved.marginV != null) ? saved.marginV : defaultMargin,
            marginH: (saved && saved.marginH != null) ? saved.marginH : defaultMargin,
            link: link,
            verticalEnabled: verticalEnabled,
            horizontalEnabled: horizontalEnabled,
            previewBounds: (saved && typeof saved.previewBounds === "boolean") ? saved.previewBounds : CONFIG.previewBoundsDefault,
            roundMode: (saved && saved.roundMode) ? saved.roundMode : CONFIG.roundModeDefault,
            operation: operation,
            scope: scope
        };
    }

    // =========================================
    // 矩形・境界のユーティリティ / Rect & bounds utilities
    // =========================================

    /* 矩形にマージンを加えた新しい矩形を返す / Return a new rect expanded by margin
       Illustrator の artboardRect は [left, top, right, bottom]（上が大・下が小）。 */
    function expandRectByMargin(rect, verticalMarginPt, horizontalMarginPt) {
        return [
            rect[0] - horizontalMarginPt,
            rect[1] + verticalMarginPt,
            rect[2] + horizontalMarginPt,
            rect[3] - verticalMarginPt
        ];
    }

    /* アートボード矩形をピクセルグリッドに最適化（X/Y/W/H を各1回だけ整数化） / Snap rect X/Y/W/H to pixel integers
       仕様：X(左)・Y(上)・幅・高さの4値をそれぞれ整数に丸め、右下は X+幅 / Y−高さ で再構成する（二重丸めしない）。 */
    function snapRectToPixelGrid(rect) {
        var x = Math.round(rect[0]);
        var y = Math.round(rect[1]);
        var width = Math.round(rect[2] - rect[0]);
        var height = Math.round(rect[1] - rect[3]);
        return [x, y, x + width, y - height];
    }

    /* 矩形の X/Y/W/H を指定単位で各1回だけ整数化 / Snap rect X/Y/W/H to integers in the given unit
       X・Y・幅・高さを単位グリッドに丸める。単位変換に失敗した場合はピクセルグリッドにフォールバック。 */
    function snapRectToUnitGrid(rect, unit) {
        var ptPerUnit = toPt(1, unit);
        if (isNaN(ptPerUnit) || ptPerUnit === 0) return snapRectToPixelGrid(rect);
        var x = Math.round(rect[0] / ptPerUnit) * ptPerUnit;
        var y = Math.round(rect[1] / ptPerUnit) * ptPerUnit;
        var width = Math.round((rect[2] - rect[0]) / ptPerUnit) * ptPerUnit;
        var height = Math.round((rect[1] - rect[3]) / ptPerUnit) * ptPerUnit;
        return [x, y, x + width, y - height];
    }

    /* 矩形の幅・高さが正か（left<right かつ bottom<top） / Valid rect: positive width & height */
    function isValidRect(rect) {
        return rect[2] > rect[0] && rect[1] > rect[3];
    }

    /* 無効化した軸を元アートボードの座標に固定（＝その軸は動かさない） / Lock disabled axes to the original artboard rect
       横(左右)OFFなら left/right を、縦(上下)OFFなら top/bottom を元アートボード値に戻す。 */
    function lockDisabledAxes(rect, artboardRect, verticalEnabled, horizontalEnabled) {
        var out = rect.slice();
        if (!horizontalEnabled) { out[0] = artboardRect[0]; out[2] = artboardRect[2]; }
        if (!verticalEnabled) { out[1] = artboardRect[1]; out[3] = artboardRect[3]; }
        return out;
    }

    /* 丸めモードに従って矩形を整数化 / Round rect according to the selected mode
       "pixelGrid" = ピクセル整数、"currentUnit" = 現在の単位で整数、"none" = 丸めなし。 */
    function applyRounding(rect, roundMode, unit) {
        if (roundMode === "pixelGrid") return snapRectToPixelGrid(rect);
        if (roundMode === "currentUnit") return snapRectToUnitGrid(rect, unit);
        return rect; // "none"
    }

    /* 2つの矩形が実質的に同一か判定 / Check if two rects are effectively equal
       マージン0などで値が変わらない場合はUndoが生成されないため、プレビューのUndoカウント除外に使う。 */
    function rectsEqual(rectA, rectB) {
        if (!rectA || !rectB) return false;
        for (var i = 0; i < 4; i++) {
            if (Math.abs(rectA[i] - rectB[i]) > 0.0001) return false;
        }
        return true;
    }

    /* オブジェクトのバウンディングボックスを取得 / Get bounding box of a single object
       usePreviewBounds=true なら visibleBounds（塗り/線を含む）、false なら geometricBounds（パス外形のみ）。 */
    function getBounds(item, usePreviewBounds) {
        return usePreviewBounds ? item.visibleBounds : item.geometricBounds;
    }

    /* 複数アイテムから最大の外接バウンディングボックスを取得（空なら null） / Union bounding box, null if empty */
    function getMaxBounds(items, usePreviewBounds) {
        if (!items || items.length === 0) return null;
        var bounds = getBounds(items[0], usePreviewBounds);
        for (var i = 1; i < items.length; i++) {
            var itemBounds = getBounds(items[i], usePreviewBounds);
            bounds[0] = Math.min(bounds[0], itemBounds[0]);
            bounds[1] = Math.max(bounds[1], itemBounds[1]);
            bounds[2] = Math.max(bounds[2], itemBounds[2]);
            bounds[3] = Math.min(bounds[3], itemBounds[3]);
        }
        return bounds;
    }

    /* 計測可能なページアイテムか（geometricBounds を持つか） / Is a measurable page item?
       TextRange 等の非ページアイテムを除外し、選択の型不整合を防ぐ。 */
    function isMeasurableItem(it) {
        if (!it) return false;
        try {
            var b = it.geometricBounds;
            return (b && b.length === 4);
        } catch (e) {
            return false;
        }
    }

    /* クリップグループのクリッピングパスを返す（無ければ null） / Clipping path of a clip group, or null */
    function getClippingPath(groupItem) {
        try {
            for (var j = 0; j < groupItem.pageItems.length; j++) {
                if (groupItem.pageItems[j].clipping) return groupItem.pageItems[j];
            }
        } catch (e) { /* ignore */ }
        return null;
    }

    /* 選択アイテムの正規化 / Normalize selected items
       ・計測できない要素（TextRange 等）は除外
       ・クリップグループはクリッピングパスのみを採用、それ以外はそのまま
       プレビューと本処理でアイテム収集を共通化する。 */
    function collectEffectiveItems(items) {
        var out = [];
        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            if (!isMeasurableItem(it)) continue; // 非ページアイテムをスキップ / skip non-page items
            if (it.typename === "GroupItem" && it.clipped) {
                var clip = getClippingPath(it);
                if (clip) out.push(clip);
            } else {
                out.push(it);
            }
        }
        return out;
    }

    /* 計測対象を再帰的に収集 / Recursively collect measure targets
       ・TextFrame は複製をアウトライン化（元は不変）。グループ内テキストも対象。
       ・クリップグループはクリッピングパスのみ。通常グループは中身へ再帰。
       一時複製・アウトラインは tempObjects に登録し、呼び出し側で必ず削除する。 */
    function collectMeasureTargets(items, tempObjects, out) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (!isMeasurableItem(item)) continue;
            if (item.typename === "TextFrame") {
                // 複製を先に登録 → アウトライン化（失敗時も複製を削除できる） / track clone before outlining
                var clone = item.duplicate();
                tempObjects.push(clone);
                var outlined = clone.createOutline(); // GroupItem を返す / returns a GroupItem
                tempObjects.push(outlined);
                out.push(outlined);
            } else if (item.typename === "GroupItem" && item.clipped) {
                var clip = getClippingPath(item);
                if (clip) out.push(clip);
            } else if (item.typename === "GroupItem") {
                // 通常グループは中身を再帰（ネストされたテキストも非破壊計測） / recurse into groups
                collectMeasureTargets(item.pageItems, tempObjects, out);
            } else {
                out.push(item);
            }
        }
    }

    /* 選択オブジェクトの外接境界を取得（空なら null） / Measure selection bounds (null if empty)
       テキストは複製をアウトライン化して計測し、計測後に一時オブジェクトを必ず削除する。
       元の TextFrame には一切触れないため、ID・重なり順・名前・タグ・ノート・Variable 等が保持される。 */
    function measureSelectionBounds(items, usePreviewBounds) {
        var tempObjects = []; // 計測用に作った一時複製・アウトライン（必ず削除） / temp objects to remove
        try {
            var measureItems = [];
            collectMeasureTargets(items, tempObjects, measureItems);
            return getMaxBounds(measureItems, usePreviewBounds);
        } finally {
            // 途中で例外が起きても一時オブジェクトは必ず削除 / always remove temp objects, even on error
            for (var k = 0; k < tempObjects.length; k++) {
                try { tempObjects[k].remove(); } catch (e) { }
            }
        }
    }

    /* マージン適用の共通パイプライン：拡張 → 丸め → 無効軸を元座標に固定 / Shared pipeline: expand → round → lock
       プレビューと確定で同一の矩形を得るために両者から使う。 */
    function computeMarginRect(baseRect, verticalMarginPt, horizontalMarginPt, roundMode, unit, artboardOriginalRect, verticalEnabled, horizontalEnabled) {
        var rect = expandRectByMargin(baseRect, verticalMarginPt, horizontalMarginPt);
        rect = applyRounding(rect, roundMode, unit);
        rect = lockDisabledAxes(rect, artboardOriginalRect, verticalEnabled, horizontalEnabled);
        return rect;
    }

    /* 1枚のアートボードにマージンを適用（拡張→丸め→無効軸固定）。
       負マージン等で無効サイズになる場合は適用せず false を返す / Apply margin; skip & return false if invalid */
    function applyMarginToArtboard(artboard, verticalMarginPt, horizontalMarginPt, roundMode, unit, verticalEnabled, horizontalEnabled) {
        var baseRect = artboard.artboardRect; // スクリプト実行時の矩形 / rect at execution time
        var rect = computeMarginRect(baseRect, verticalMarginPt, horizontalMarginPt, roundMode, unit, baseRect, verticalEnabled, horizontalEnabled);
        if (!isValidRect(rect)) return false;
        artboard.artboardRect = rect;
        return true;
    }

    // =========================================
    // プレビュー適用 / Preview application
    // =========================================
    // 巻き戻しはスナップショット復元（PreviewManager.restoreFn）で行う。
    // 確定時と同じ computeMarginRect（丸め・無効軸固定を含む）を使い、プレビューと結果を一致させる。
    // 矩形が変わらない/無効な場合は書き換えをスキップし、undo履歴のノイズを減らす。

    // --- 拡張（アートボード自身のサイズにマージンを加減） / Expand: grow the artboard by margins ---

    /* すべてのアートボードを拡張プレビュー / Preview: expand all artboards */
    function previewExpandAll(originalRects, verticalMarginPt, horizontalMarginPt, roundMode, unit, verticalEnabled, horizontalEnabled) {
        var artboards = app.activeDocument.artboards;
        for (var i = 0; i < artboards.length; i++) {
            var rect = computeMarginRect(originalRects[i], verticalMarginPt, horizontalMarginPt, roundMode, unit, originalRects[i], verticalEnabled, horizontalEnabled);
            if (!isValidRect(rect)) continue;
            if (rectsEqual(artboards[i].artboardRect, rect)) continue;
            artboards[i].artboardRect = rect;
        }
        app.redraw();
    }

    /* 現在のアートボードを拡張プレビュー / Preview: expand the active artboard */
    function previewExpandCurrent(originalRects, index, verticalMarginPt, horizontalMarginPt, roundMode, unit, verticalEnabled, horizontalEnabled) {
        var artboards = app.activeDocument.artboards;
        var rect = computeMarginRect(originalRects[index], verticalMarginPt, horizontalMarginPt, roundMode, unit, originalRects[index], verticalEnabled, horizontalEnabled);
        if (isValidRect(rect) && !rectsEqual(artboards[index].artboardRect, rect)) {
            artboards[index].artboardRect = rect;
        }
        app.redraw();
    }

    // --- 合わせる（選択の外接＋マージンにアートボードを合わせる） / Fit: resize artboard(s) to selection bounds ---
    // boundsRect は確定時と同一ロジック（measureSelectionBounds）で得た計測済み矩形。
    // 無効軸は各アートボードの実行時座標（originalRect）に固定する。

    /* 現在のアートボードを選択に合わせるプレビュー / Preview: fit the active artboard to selection */
    function previewFitCurrent(index, boundsRect, verticalMarginPt, horizontalMarginPt, roundMode, unit, verticalEnabled, horizontalEnabled, artboardOriginalRect) {
        if (!boundsRect) return;
        var artboards = app.activeDocument.artboards;
        var rect = computeMarginRect(boundsRect, verticalMarginPt, horizontalMarginPt, roundMode, unit, artboardOriginalRect, verticalEnabled, horizontalEnabled);
        if (isValidRect(rect) && !rectsEqual(artboards[index].artboardRect, rect)) {
            artboards[index].artboardRect = rect;
        }
        app.redraw();
    }

    // =========================================
    // 入力ユーティリティ / Input utilities
    // =========================================

    /* edittext に矢印キーでの増減を付与 / Add arrow-key increment/decrement to an edittext
       ↑↓で±1、Shift+↑↓で10の倍数にスナップ（例 36→40 / 36→30）、Option(Alt)+↑↓で±0.1。 */
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
    // ダイアログ / Dialog
    // =========================================

    /* マージン入力ダイアログを表示し設定を返す（ライブプレビュー付き） / Show margin dialog with live preview
       selectionItems はダイアログ表示時に固定した選択アイテム配列（「合わせる」で使用）。 */
    function showMarginDialog(defaultMargin, rulerUnit, artboardCount, hasSelection, selectionItems) {
        var dialog = new Window("dialog", getLocalizedText(LABELS.dialog.title) + " " + SCRIPT_VERSION);

        // ダイアログ位置の記憶（保存関数を受け取る） / wire position persistence, get the saver
        var persistDialogLocation = attachPositionPersistence(dialog, "__FitArtboardWithMargin_Dialog", CONFIG.offsetX);
        dialog.opacity = CONFIG.dialogOpacity;
        setupWindow(dialog);

        // 保存済み設定（セッション内）から初期値を解決 / resolve initial values from stored settings
        var initial = resolveInitialSettings(defaultMargin, artboardCount, hasSelection);

        // 選択境界のキャッシュ（プレビューと確定で同一の計測結果を使う） / cache measured selection bounds per previewBounds flag
        var selectionBoundsCache = {};
        function getSelectionBounds(usePreviewBounds) {
            var key = usePreviewBounds ? "v" : "g";
            if (!selectionBoundsCache.hasOwnProperty(key)) {
                selectionBoundsCache[key] = (selectionItems && selectionItems.length > 0) ?
                    measureSelectionBounds(selectionItems, usePreviewBounds) : null;
            }
            return selectionBoundsCache[key];
        }

        /* 調整基準パネル（操作＋対象の2グループ） / Basis panel: operation + scope */
        var targetPanel = dialog.add("panel", undefined, getLocalizedText(LABELS.panel.target));
        setupPanel(targetPanel);

        var basisLabelWidth = (currentLanguage === "ja") ? 40 : 76;

        /* 操作：オブジェクトに合わせる / アートボードを拡張（ラジオは縦並び） / Operation group (vertical radios) */
        var operationRow = targetPanel.add("group");
        operationRow.orientation = "row";
        operationRow.alignChildren = ["left", "top"];
        var operationLabel = operationRow.add("statictext", undefined, getLocalizedText(LABELS.groupLabel.operation) + UI_COLON);
        operationLabel.preferredSize.width = basisLabelWidth;
        var operationGroup = operationRow.add("group");
        operationGroup.orientation = "column";
        operationGroup.alignChildren = "left";
        var fitRadio = operationGroup.add("radiobutton", undefined, getLocalizedText(LABELS.radio.fit));
        fitRadio.enabled = hasSelection; // 合わせるは選択必須 / fit needs a selection
        fitRadio.helpTip = getLocalizedText(LABELS.tooltip.opFit);
        var expandRadio = operationGroup.add("radiobutton", undefined, getLocalizedText(LABELS.radio.expand));
        expandRadio.helpTip = getLocalizedText(LABELS.tooltip.opExpand);

        /* 対象：現在のアートボード / すべてのアートボード（ラジオは縦並び） / Scope group (vertical radios) */
        var scopeRow = targetPanel.add("group");
        scopeRow.orientation = "row";
        scopeRow.alignChildren = ["left", "top"];
        var scopeLabel = scopeRow.add("statictext", undefined, getLocalizedText(LABELS.groupLabel.scope) + UI_COLON);
        scopeLabel.preferredSize.width = basisLabelWidth;
        var scopeGroup = scopeRow.add("group");
        scopeGroup.orientation = "column";
        scopeGroup.alignChildren = "left";
        var currentRadio = scopeGroup.add("radiobutton", undefined, getLocalizedText(LABELS.radio.current));
        currentRadio.helpTip = getLocalizedText(LABELS.tooltip.scopeCurrent);
        var allRadio = scopeGroup.add("radiobutton", undefined, getLocalizedText(LABELS.radio.all));
        allRadio.helpTip = getLocalizedText(LABELS.tooltip.scopeAll);

        /* 現在の操作・対象を取得 / current operation & scope */
        function getOperation() { return fitRadio.value ? "fit" : "expand"; }
        function getScope() { return allRadio.value ? "all" : "current"; }

        /* 「合わせる」は現在のアートボード専用。fit のときは対象を現在に固定してディム。 / fit is current-only: lock & dim scope */
        function refreshScopeState() {
            var isFit = (getOperation() === "fit");
            if (isFit) {
                currentRadio.value = true;
                allRadio.value = false;
            }
            scopeLabel.enabled = !isFit;
            currentRadio.enabled = !isFit;
            allRadio.enabled = !isFit;
        }

        /* 初期値を適用 / apply initial selection */
        fitRadio.value = (initial.operation === "fit");
        expandRadio.value = (initial.operation !== "fit");
        currentRadio.value = (initial.scope === "current");
        allRadio.value = (initial.scope === "all");
        refreshScopeState();

        /* マージン入力パネル（2カラム） / Margin input panel (two columns) */
        var marginPanel = dialog.add("panel", undefined, getLocalizedText(LABELS.panel.margin) + " (" + rulerUnit + ")");
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

        /* 上下マージン入力欄（チェックボックスで有効/無効） / Vertical margin (checkbox toggles the axis) */
        var verticalMarginRow = marginFieldsColumn.add("group");
        verticalMarginRow.orientation = "row";
        var verticalEnabledCheckbox = verticalMarginRow.add("checkbox", undefined, getLocalizedText(LABELS.field.vertical));
        verticalEnabledCheckbox.value = initial.verticalEnabled;
        verticalEnabledCheckbox.helpTip = getLocalizedText(LABELS.tooltip.axisEnable);
        var verticalMarginInput = verticalMarginRow.add("edittext", undefined, initial.marginV);
        verticalMarginInput.characters = 4;
        verticalMarginInput.helpTip = getLocalizedText(LABELS.tooltip.marginInput);

        /* 左右マージン入力欄（チェックボックスで有効/無効） / Horizontal margin (checkbox toggles the axis) */
        var horizontalMarginRow = marginFieldsColumn.add("group");
        horizontalMarginRow.orientation = "row";
        var horizontalEnabledCheckbox = horizontalMarginRow.add("checkbox", undefined, getLocalizedText(LABELS.field.horizontal));
        horizontalEnabledCheckbox.value = initial.horizontalEnabled;
        horizontalEnabledCheckbox.helpTip = getLocalizedText(LABELS.tooltip.axisEnable);
        var horizontalMarginInput = horizontalMarginRow.add("edittext", undefined, initial.link ? initial.marginV : initial.marginH);
        horizontalMarginInput.characters = 4;
        horizontalMarginInput.helpTip = getLocalizedText(LABELS.tooltip.marginInput);

        /* 上下・左右チェックボックスの幅を揃えて入力欄位置を統一 / Align checkbox widths so inputs line up */
        var fieldLabelWidth = (currentLanguage === "ja") ? 62 : 96;
        verticalEnabledCheckbox.preferredSize.width = fieldLabelWidth;
        horizontalEnabledCheckbox.preferredSize.width = fieldLabelWidth;

        /* 連動チェックボックス / Linked checkbox */
        var linkCheckbox = linkColumn.add("checkbox", undefined, getLocalizedText(LABELS.checkbox.linked));
        linkCheckbox.value = initial.link;
        linkCheckbox.helpTip = getLocalizedText(LABELS.tooltip.link);

        /* 入力欄と連動チェックの有効/無効を現在の状態から更新 / Refresh input & link enabled states
           上下=自軸チェック、左右=自軸チェック かつ 非連動、連動=上下・左右が両方ONのときだけ有効。 */
        function refreshMarginInputStates() {
            verticalMarginInput.enabled = verticalEnabledCheckbox.value;
            horizontalMarginInput.enabled = horizontalEnabledCheckbox.value && !linkCheckbox.value;
            // どちらかの軸がOFFなら連動は使えない（自動OFFのうえディム） / disable link when either axis is off
            linkCheckbox.enabled = verticalEnabledCheckbox.value && horizontalEnabledCheckbox.value;
        }
        refreshMarginInputStates();

        /* プレビュー境界（visibleBounds を採用するか。「合わせる」系のときのみ有効） / use visibleBounds; only for fit modes */
        var previewBoundsCheckbox = linkColumn.add("checkbox", undefined, getLocalizedText(LABELS.checkbox.previewBounds));
        previewBoundsCheckbox.value = initial.previewBounds;
        previewBoundsCheckbox.helpTip = getLocalizedText(LABELS.tooltip.previewBounds);
        previewBoundsCheckbox.enabled = (initial.operation === "fit"); // 合わせる時のみ有効 / only for fit
        previewBoundsCheckbox.onClick = updatePreview;

        /* 現在／全アートボードの rect を保存（プレビュー復元用） / Save artboard rects for preview restore */
        var activeArtboardIndex = app.activeDocument.artboards.getActiveArtboardIndex();
        var originalArtboardRects = [];
        for (var i = 0; i < app.activeDocument.artboards.length; i++) {
            originalArtboardRects.push(app.activeDocument.artboards[i].artboardRect.slice());
        }

        /* プレビュー復元：開いた時点の全 artboardRect へ書き戻す（undo回数に依存しない） / Restore all artboard rects to the opening snapshot */
        var previewManager = new PreviewManager(function () {
            var artboards = app.activeDocument.artboards;
            for (var i = 0; i < artboards.length && i < originalArtboardRects.length; i++) {
                if (!rectsEqual(artboards[i].artboardRect, originalArtboardRects[i])) {
                    artboards[i].artboardRect = originalArtboardRects[i];
                }
            }
        });

        /* 有効チェックを反映した実効マージン（pt）を取得。無効な軸は 0 とみなす / Effective margins in pt; disabled axis = 0 */
        function getEffectiveMargins() {
            var vEnabled = verticalEnabledCheckbox.value;
            var hEnabled = horizontalEnabledCheckbox.value;
            var vPt = vEnabled ? toPt(parseFloat(verticalMarginInput.text), rulerUnit) : 0;
            var hPt = hEnabled ? toPt(parseFloat(horizontalMarginInput.text), rulerUnit) : 0;
            // 有効な軸だけ数値妥当性を要求 / only enabled axes must be valid numbers
            var valid = (!vEnabled || !isNaN(vPt)) && (!hEnabled || !isNaN(hPt));
            return { valid: valid, vPt: vPt, hPt: hPt };
        }

        /* プレビュー更新：直前分を rollback してから最新状態を1回だけ適用 / Refresh preview via PreviewManager
           対象モードに応じてトップレベルのプレビュー関数へ委譲する。 */
        function updatePreview() {
            previewManager.rollback();

            var margins = getEffectiveMargins();
            if (!margins.valid) return;
            var verticalMarginPt = margins.vPt;
            var horizontalMarginPt = margins.hPt;

            var operation = getOperation();
            var scope = getScope();

            var verticalEnabled = verticalEnabledCheckbox.value;
            var horizontalEnabled = horizontalEnabledCheckbox.value;
            // 確定時と同じ丸めモードをプレビューにも反映 / apply the same rounding as confirm
            var roundMode = roundPixelRadio.value ? "pixelGrid" : (roundUnitRadio.value ? "currentUnit" : "none");

            previewManager.addStep(function () {
                if (operation === "expand") {
                    if (scope === "all") { previewExpandAll(originalArtboardRects, verticalMarginPt, horizontalMarginPt, roundMode, rulerUnit, verticalEnabled, horizontalEnabled); }
                    else { previewExpandCurrent(originalArtboardRects, activeArtboardIndex, verticalMarginPt, horizontalMarginPt, roundMode, rulerUnit, verticalEnabled, horizontalEnabled); }
                    return;
                }
                // 合わせる：常に現在のアートボードを選択の外接に合わせる（無効軸は実行時座標に固定） / fit is current-only
                var boundsRect = getSelectionBounds(previewBoundsCheckbox.value);
                previewFitCurrent(activeArtboardIndex, boundsRect, verticalMarginPt, horizontalMarginPt, roundMode, rulerUnit, verticalEnabled, horizontalEnabled, originalArtboardRects[activeArtboardIndex]);
            });
        }

        /* 操作/対象の変更時：対象の固定・プレビュー境界の有効/無効を切り替えてプレビュー更新 / On change: refresh scope & previewBounds */
        function onBasisChange() {
            refreshScopeState(); // fit のときは対象を現在に固定＋ディム / fit locks scope to current
            previewBoundsCheckbox.enabled = (getOperation() === "fit"); // 合わせる時のみ有効 / only for fit
            updatePreview();
        }

        /* 矢印キー・入力・ラジオ・連動のハンドラ登録 / Wire up input handlers */
        changeValueByArrowKey(verticalMarginInput, function (val) {
            if (linkCheckbox.value) horizontalMarginInput.text = val;
            updatePreview();
        });
        changeValueByArrowKey(horizontalMarginInput, function () {
            if (linkCheckbox.value) return;
            updatePreview();
        });
        verticalMarginInput.onChanging = function () {
            if (linkCheckbox.value) horizontalMarginInput.text = verticalMarginInput.text;
            updatePreview();
        };
        horizontalMarginInput.onChanging = function () {
            if (linkCheckbox.value) return; // 連動中は水平の直接編集は無効
            updatePreview();
        };

        /* Option(Alt)クリック検出：mousedown で event.altKey を捕捉（onClick では修飾キーを取得できない） / capture Alt on mousedown */
        var axisSoloRequested = { vertical: false, horizontal: false };
        verticalEnabledCheckbox.addEventListener("mousedown", function (event) { axisSoloRequested.vertical = (event.altKey === true); });
        horizontalEnabledCheckbox.addEventListener("mousedown", function (event) { axisSoloRequested.horizontal = (event.altKey === true); });

        /* 軸チェックの変更処理 / axis toggle handler
           Option+クリック＝ソロ（クリックした軸のみON、もう片方OFF）。どちらかOFFなら連動を自動OFF。 */
        function handleAxisToggle(solo, clickedCheckbox, otherCheckbox) {
            if (solo) {
                clickedCheckbox.value = true;   // クリックした軸をON / keep the clicked axis on
                otherCheckbox.value = false;    // もう片方をOFF / turn the other off
            }
            if (!verticalEnabledCheckbox.value || !horizontalEnabledCheckbox.value) linkCheckbox.value = false;
            refreshMarginInputStates();
            updatePreview();
        }
        verticalEnabledCheckbox.onClick = function () {
            var solo = axisSoloRequested.vertical; axisSoloRequested.vertical = false;
            handleAxisToggle(solo, verticalEnabledCheckbox, horizontalEnabledCheckbox);
        };
        horizontalEnabledCheckbox.onClick = function () {
            var solo = axisSoloRequested.horizontal; axisSoloRequested.horizontal = false;
            handleAxisToggle(solo, horizontalEnabledCheckbox, verticalEnabledCheckbox);
        };

        linkCheckbox.onClick = function () {
            if (linkCheckbox.value) {
                // 連動ON：上下・左右を有効に揃え、値を上下に統一 / linking re-enables both axes and mirrors V→H
                verticalEnabledCheckbox.value = true;
                horizontalEnabledCheckbox.value = true;
                horizontalMarginInput.text = verticalMarginInput.text;
            }
            refreshMarginInputStates();
            updatePreview();
        };
        fitRadio.onClick = onBasisChange;
        expandRadio.onClick = onBasisChange;
        currentRadio.onClick = onBasisChange;
        allRadio.onClick = onBasisChange;

        /* アートボードの微調整パネル / Artboard fine-tuning panel */
        var optionsPanel = dialog.add("panel", undefined, getLocalizedText(LABELS.panel.options));
        setupPanel(optionsPanel);

        // 丸めモード（XYWHの整数化方法） / rounding mode for artboard X/Y/W/H
        // プレビューにも確定と同じ丸めを反映する。 / preview reflects the same rounding as confirm
        var roundModeGroup = optionsPanel.add("group");
        roundModeGroup.orientation = "column";
        roundModeGroup.alignChildren = "left";
        var roundPixelRadio = roundModeGroup.add("radiobutton", undefined, getLocalizedText(LABELS.roundMode.pixelGrid));
        var roundUnitRadio = roundModeGroup.add("radiobutton", undefined, getLocalizedText(LABELS.roundMode.currentUnit));
        var roundNoneRadio = roundModeGroup.add("radiobutton", undefined, getLocalizedText(LABELS.roundMode.none));
        roundPixelRadio.helpTip = getLocalizedText(LABELS.tooltip.roundPixel);
        roundUnitRadio.helpTip = getLocalizedText(LABELS.tooltip.roundUnit);
        roundNoneRadio.helpTip = getLocalizedText(LABELS.tooltip.roundNone);
        roundUnitRadio.value = (initial.roundMode === "currentUnit");
        roundNoneRadio.value = (initial.roundMode === "none");
        roundPixelRadio.value = !roundUnitRadio.value && !roundNoneRadio.value; // 既定 / default
        roundPixelRadio.onClick = updatePreview;
        roundUnitRadio.onClick = updatePreview;
        roundNoneRadio.onClick = updatePreview;

        /* ボタングループ（右寄せ：左は空き、Cancel → OK の順で OK を一番右に） / Button group: right-aligned, OK rightmost */
        var buttonGroup = dialog.add("group");
        setupRow(buttonGroup, "right");
        var cancelButton = buttonGroup.add("button", undefined, getLocalizedText(LABELS.button.cancel), { name: "cancel" });
        var okButton = buttonGroup.add("button", undefined, getLocalizedText(LABELS.button.ok), { name: "ok" });

        var dialogResult = null;
        okButton.onClick = function () {
            if (getEffectiveMargins().valid) {
                dialogResult = {
                    marginV: verticalMarginInput.text,
                    marginH: horizontalMarginInput.text,
                    verticalEnabled: verticalEnabledCheckbox.value,
                    horizontalEnabled: horizontalEnabledCheckbox.value,
                    operation: getOperation(),
                    scope: getScope(),
                    previewBounds: previewBoundsCheckbox.value,
                    roundMode: roundPixelRadio.value ? "pixelGrid" : (roundUnitRadio.value ? "currentUnit" : "none")
                };
                // 設定をセッションに保存（次回の初期値に） / store settings for next run
                storeSettings({
                    marginV: dialogResult.marginV,
                    marginH: dialogResult.marginH,
                    link: linkCheckbox.value,
                    verticalEnabled: dialogResult.verticalEnabled,
                    horizontalEnabled: dialogResult.horizontalEnabled,
                    previewBounds: dialogResult.previewBounds,
                    roundMode: dialogResult.roundMode,
                    operation: dialogResult.operation,
                    scope: dialogResult.scope
                });
                persistDialogLocation();
                // プレビューを開いた時点へ復元してから閉じる / restore to the opening snapshot, then close
                previewManager.confirm();
                dialog.close(1);
            } else {
                alert(getLocalizedText(LABELS.alert.enterNumber));
            }
        };
        cancelButton.onClick = function () {
            persistDialogLocation();
            // キャンセル時は必ずロールバックして閉じる / rollback preview and close
            previewManager.rollback();
            dialog.close(0);
        };

        updatePreview();

        // 開いたら有効な方のマージン入力欄にフォーカス（上下→無ければ左右） / focus a margin field on open
        if (verticalMarginInput.enabled) {
            verticalMarginInput.active = true;
        } else if (horizontalMarginInput.enabled) {
            horizontalMarginInput.active = true;
        }

        dialog.show();
        return dialogResult;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /* 対象を決定し、ダイアログの設定に従ってアートボードを調整する / Resolve target and apply margins */
    function main() {
        try {
            // ドキュメント未オープンなら分かりやすく案内して終了 / friendly guard when no document is open
            if (app.documents.length === 0) {
                alert(getLocalizedText(LABELS.alert.noDocument));
                return;
            }
            var doc = app.activeDocument;

            var hasSelection = doc.selection.length > 0;
            // ダイアログ表示時点の選択を固定（「合わせる」で共通利用） / freeze selection at dialog time
            var selectionItems = collectEffectiveItems(doc.selection);
            hasSelection = selectionItems.length > 0; // 計測可能な選択があるか / measurable selection?

            var artboards = doc.artboards;
            var rulerType = app.preferences.getIntegerPreference("rulerType");
            var rulerUnit = rulerUnitFromType(rulerType);

            /* 単位ごとの初期マージン値（保存済み設定があればそちらが優先される） / default margin for the unit (stored settings win) */
            var defaultMarginValue = getDefaultMargin(rulerUnit);

            /* 選択なし・複数アートボード時は初期マージンを0に（保存が無い初回のみ有効） / start at 0 for expand-all default (first run only) */
            if (!hasSelection && artboards.length > 1) {
                defaultMarginValue = '0';
            }

            var userInput = showMarginDialog(defaultMarginValue, rulerUnit, artboards.length, hasSelection, selectionItems);
            if (!userInput) return; // キャンセル時は選択ツールにも切り替えない / cancel: leave tool unchanged

            // 無効化した軸のマージンは 0 として扱う / disabled axes contribute zero margin
            var verticalMarginPt = userInput.verticalEnabled ? toPt(parseFloat(userInput.marginV), rulerUnit) : 0;
            var horizontalMarginPt = userInput.horizontalEnabled ? toPt(parseFloat(userInput.marginH), rulerUnit) : 0;
            var roundMode = userInput.roundMode;
            var verticalEnabled = userInput.verticalEnabled;
            var horizontalEnabled = userInput.horizontalEnabled;
            var skippedCount = 0;

            if (userInput.operation === "expand") {
                // 拡張：各アートボード自身のサイズにマージンを加減 / expand each artboard by margins
                if (userInput.scope === "all") {
                    for (var i = 0; i < artboards.length; i++) {
                        if (!applyMarginToArtboard(artboards[i], verticalMarginPt, horizontalMarginPt, roundMode, rulerUnit, verticalEnabled, horizontalEnabled)) skippedCount++;
                    }
                } else {
                    if (!applyMarginToArtboard(artboards[artboards.getActiveArtboardIndex()], verticalMarginPt, horizontalMarginPt, roundMode, rulerUnit, verticalEnabled, horizontalEnabled)) skippedCount++;
                }
            } else if (userInput.operation === "fit" && selectionItems.length > 0) {
                // 合わせる：現在のアートボードを選択の外接＋マージンに合わせる（プレビューと同一パイプライン） / fit current artboard to selection
                var activeArtboard = artboards[artboards.getActiveArtboardIndex()];
                var artboardOriginalRect = activeArtboard.artboardRect.slice(); // 実行時座標（無効軸の固定用） / for locking disabled axes
                var fitBounds = measureSelectionBounds(selectionItems, userInput.previewBounds);
                if (fitBounds) { // 計測可能なジオメトリがある場合のみ / only when measurable geometry exists
                    var rect = computeMarginRect(fitBounds, verticalMarginPt, horizontalMarginPt, roundMode, rulerUnit, artboardOriginalRect, verticalEnabled, horizontalEnabled);
                    if (!isValidRect(rect)) skippedCount++;
                    else activeArtboard.artboardRect = rect;
                }
            }

            if (skippedCount > 0) alert(getLocalizedText(LABELS.alert.marginTooLarge));

            // 適用が完了したときのみ選択ツールへ切り替え（キャンセル・エラー時は切り替えない） / switch tool only after a successful run
            app.selectTool("Adobe Select Tool");

        } catch (e) {
            $.writeln("[FitArtboardWithMargin] ERROR: " + formatError(e));
            alert(getLocalizedText(LABELS.alert.errorOccurred) + formatError(e));
        }
    }

    main();

})();
