#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "ExtendLinesEngine"

/*

### 概要

選択オブジェクト内のパス（グループ／複合パスを含む）から隣接するアンカーポイントのペアを取り、補助線として「直線を描画範囲いっぱいに延長した線」を描画します。
［円弧から円］でBezier曲線セグメントから円を推定することもできます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExtendLines.md

### Overview

Takes pairs of adjacent anchor points from the paths in the selection — groups and compound paths included — and draws each as a construction line extended across the drawing area.
The "arc to circle" option can also estimate a circle from a Bézier segment.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExtendLines.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExtendLines";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-27";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExtendLines.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExtendLines.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

var SCRIPT_MARKER = "__ExtendLines__";

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 線幅の既定値（0.1mm を pt で）/ Default stroke width (0.1 mm in points) */
    var DEFAULT_STROKE_WIDTH_PT = 0.1 * 72.0 / 25.4;

    /* ［別レイヤーに］の描画先レイヤー名 / Layer used by "Separate layer" */
    var CONSTRUCTION_LAYER_NAME = "_construction";

    /* プレビュー用レイヤー名のもと（重複時は連番を付ける）/ Base name of the preview layer (numbered when taken) */
    var PREVIEW_LAYER_BASE_NAME = "__ExtendLines_Preview";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 15;                 /* ダイアログの余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];    /* パネルの余白 [左,上,右,下] / panel margins */
    var COLUMN_SPACING = 10;                 /* 2カラムの間隔 / gap between the two columns */
    var STROKE_ROW_SPACING = 6;              /* 線幅の行の間隔 / spacing in the stroke-width row */
    var STROKE_INPUT_CHARS = 6;              /* 線幅の入力欄の幅（文字数）/ width of the stroke-width field */

    /**
     * パネルを縦並び・左揃えにして共通の余白を付ける
     * @param {Panel} targetPanel - 対象のパネル
     * @returns {void}
     */
    function applyPanelLayout(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "補助線の描画", en: "Extend Lines" }
        },
        panel: {
            auxLines: { ja: "補助線を描画", en: "Draw guide lines" },
            arcOptions: { ja: "円弧オプション", en: "Arc options" },
            options: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            straightOnly: { ja: "直線", en: "Straight" },
            arcToCircle: { ja: "円弧から円", en: "Create circle from arc" },
            group: { ja: "グループ化", en: "Group" },
            separateLayer: { ja: "別レイヤーに", en: "Separate layer" },
            guide: { ja: "ガイド化", en: "Convert to guides" },
            dedup: { ja: "線のダブりを削除", en: "Remove duplicates" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        radio: {
            arcIgnore: { ja: "無視", en: "Ignore" },
            arcStraight: { ja: "直線", en: "Straight" },
            arcExtend: { ja: "直線（延長）", en: "Straight (extend)" }
        },
        fieldLabel: {
            strokeWidth: { ja: "線幅", en: "Stroke width" }
        },
        tooltip: {
            arcOptions: { ja: "完全な円弧以外の場合", en: "If not a perfect circular arc" },
            straightOnly: {
                ja: "オンのときは直線のセグメントだけ、オフのときは曲線のセグメントだけを延長します。",
                en: "On: extends straight segments only. Off: extends curved segments only."
            },
            arcToCircle: {
                ja: "円弧になっている曲線のセグメントから、その円を描きます。",
                en: "Draws the full circle of each curved segment that is a circular arc."
            },
            arcIgnore: { ja: "円弧でない曲線には何も描きません。", en: "Draws nothing for curves that are not circular arcs." },
            arcStraight: {
                ja: "円弧でない曲線には、両端のアンカーポイントを結ぶ線分を描きます。",
                en: "For curves that are not circular arcs, draws the segment between their end anchor points."
            },
            arcExtend: {
                ja: "円弧でない曲線には、両端のアンカーポイントを結ぶ直線を描画範囲いっぱいに延長して描きます。",
                en: "For curves that are not circular arcs, draws the line through their end anchor points across the drawing area."
            },
            separateLayer: {
                ja: "「_construction」レイヤーに描きます。再実行時は、このスクリプトで描いた線を消してから描き直します。",
                en: "Draws on the \"_construction\" layer. On a rerun, lines this script drew earlier are cleared first."
            },
            dedup: { ja: "重なる延長線を1本にまとめます。", en: "Draws only one line where extended lines overlap." },
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
        },
        itemName: {
            lineGroup: { ja: "補助線", en: "Aux Lines" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            noValidPath: { ja: "有効なパスが見つかりません。", en: "No valid paths were found." },
            error: { ja: "エラー: ", en: "Error: " }
        }
    };

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
            if (!selectedItems || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // UI の明暗（再利用パーツ）ここまで / End of the reusable UI theme
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ステップボタン（再利用パーツ） / Stepper buttons (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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
     * 整数化・下限・上限・単位（「20 mm」の形）へそろえ、数値でなければ直前の値に戻す
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

        /* 直接入力をそろえる。数値でなければ直前の値に戻す / normalize typed values; revert non-numbers */
        numberInput.lastValidText = numberInput.text;
        numberInput.onChange = function () {
            var value = parseFloat(numberInput.text);
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
            var value = parseFloat(numberInput.text);
            if (isNaN(value)) value = 0;
            writeSteppedValue(numberInput, computeSteppedValue(value, direction, stepOptions), stepOptions);
            if (stepOptions.onStep) stepOptions.onStep(numberInput);
        }

        /* 整数の欄では option＋クリックの0.1刻みが効かないので、説明から外す / integer fields have no 0.1 step */
        var upTooltip = stepOptions.integer ? LABELS.tooltip.stepUpInteger : LABELS.tooltip.stepUp;
        var downTooltip = stepOptions.integer ? LABELS.tooltip.stepDownInteger : LABELS.tooltip.stepDown;
        makeStepperChevronButton(stepperGroup, "up", function () { stepBy(1); }).helpTip = getLabel(upTooltip);
        makeStepperChevronButton(stepperGroup, "down", function () { stepBy(-1); }).helpTip = getLabel(downTooltip);
        stepperGroup.stepBy = stepBy; /* ↑↓キーからも同じ処理で増減できるよう公開 / shared with the arrow keys */
        return stepperGroup;
    }

    /**
     * 入力欄の↑↓キーを、∧∨と同じ処理で増減させる。ほかのキーは素通し
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ステップボタン（再利用パーツ）ここまで / End of the reusable stepper
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 入力の補助 / Input helpers
    // =========================================

    /**
     * 数値の入力文字列を読む。空白を除き、全角のカンマ・小数点を半角にし、
     * 「1,5」のような小数カンマは小数点に、「1,000」のような桁区切りは取り除く
     * @param {string} inputText - 入力された文字列
     * @returns {number} 読み取った数値（読めなければ NaN）
     */
    function parseLocaleNumber(inputText) {
        var normalizedText = String(inputText);
        normalizedText = normalizedText.replace(/\s+/g, "").replace(/，/g, ",").replace(/．/g, ".");
        if (normalizedText.indexOf(",") >= 0 && normalizedText.indexOf(".") < 0) {
            normalizedText = normalizedText.replace(/,/g, ".");
        } else {
            normalizedText = normalizedText.replace(/,/g, "");
        }
        return Number(normalizedText);
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントを確認して補助線の描画を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        try {
            runExtendLines(app.activeDocument);
        } catch (e) {
            alert(getLabel("alert.error") + e);
        }
    }

    /**
     * 選択からパスを集め、ダイアログの設定で補助線を描く
     * @param {Document} doc - 対象のドキュメント
     * @returns {void}
     */
    function runExtendLines(doc) {
        var currentSelection = doc.selection;

        if (!currentSelection || currentSelection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 選択からパスを抽出し、テキストは一時的に複製・アウトライン化してから追加
           / Collect paths from the selection; text is duplicated and outlined temporarily, then added */
        var targetPaths = collectSelectionPathItems(currentSelection, { unique: false });
        var tempOutlineRoots = outlineTextFromSelection(currentSelection);
        if (tempOutlineRoots.length > 0) {
            targetPaths = targetPaths.concat(collectSelectionPathItems(tempOutlineRoots, { unique: false }));
        }

        if (targetPaths.length === 0) {
            alert(getLabel("alert.noValidPath"));
            cleanupTempOutlines(tempOutlineRoots);
            return;
        }

        var drawSettings = showDialog(doc, currentSelection, targetPaths);
        if (drawSettings === null) {
            cleanupTempOutlines(tempOutlineRoots);
            return; /* キャンセル / cancelled */
        }

        var drawBounds = getDrawBounds(doc, currentSelection);
        var targetLayer = prepareTargetLayer(doc, currentSelection, drawSettings.separateLayer);
        var lineContainer = createLineContainer(targetLayer, drawSettings.group);

        try {
            drawConstructionLines(lineContainer, targetPaths, drawSettings, drawBounds);
        } finally {
            /* 一時アウトラインの後始末（失敗時も確実に削除）/ Always remove the temporary outlines */
            cleanupTempOutlines(tempOutlineRoots);
        }
    }

    /**
     * 描画範囲を求める。通常はアクティブなアートボードで、選択がアートボードと
     * まったく交差しないときだけ、選択の中心を中心に幅・高さ4倍の矩形を使う
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 選択オブジェクト
     * @returns {{left: number, top: number, right: number, bottom: number}} 描画範囲
     */
    function getDrawBounds(doc, currentSelection) {
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var drawBounds = { left: artboardRect[0], top: artboardRect[1], right: artboardRect[2], bottom: artboardRect[3] };

        var selectionBounds = getClipAwareUnionBounds(currentSelection, false);
        if (!selectionBounds) return drawBounds;

        var selLeft = selectionBounds[0], selTop = selectionBounds[1], selRight = selectionBounds[2], selBottom = selectionBounds[3];

        /* Illustrator の座標は top > bottom（Yが上に行くほど増える）/ Illustrator's Y axis points up (top > bottom) */
        var intersects = !(selRight < drawBounds.left || selLeft > drawBounds.right || selTop < drawBounds.bottom || selBottom > drawBounds.top);
        if (intersects) return drawBounds;

        var selWidth = selRight - selLeft;
        var selHeight = selTop - selBottom;
        /* 幅・高さが極端に小さい場合の安全策 / Guard against a degenerate selection */
        if (selWidth < 1) selWidth = 1;
        if (selHeight < 1) selHeight = 1;

        var centerX = (selLeft + selRight) / 2;
        var centerY = (selTop + selBottom) / 2;
        var drawWidth = selWidth * 4;
        var drawHeight = selHeight * 4;

        return {
            left: centerX - drawWidth / 2,
            top: centerY + drawHeight / 2,
            right: centerX + drawWidth / 2,
            bottom: centerY - drawHeight / 2
        };
    }

    /**
     * 描画先のレイヤーを用意する。［別レイヤーに］がオフならアクティブレイヤー。
     * オンなら _construction レイヤーを再利用し（このスクリプトの生成物は消す）、
     * 選択が _construction 上にあるときは既存を _backup に改名して新しく作る
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 選択オブジェクト
     * @param {boolean} shouldSeparateLayer - ［別レイヤーに］の値
     * @returns {Layer} 描画先のレイヤー
     */
    function prepareTargetLayer(doc, currentSelection, shouldSeparateLayer) {
        if (!shouldSeparateLayer) return doc.activeLayer;

        var targetLayer;
        if (isSelectionOnLayer(currentSelection, CONSTRUCTION_LAYER_NAME)) {
            /* 元のオブジェクトを消さないため、既存をバックアップ名へ改名して新しく作る / Keep the originals: rename the existing layer and start a new one */
            var existingLayer = findLayerByName(doc, CONSTRUCTION_LAYER_NAME);
            if (existingLayer) {
                var backupName = createUniqueLayerName(doc, CONSTRUCTION_LAYER_NAME + "_backup");
                try { existingLayer.name = backupName; } catch (e) { }
            }
            targetLayer = doc.layers.add();
            targetLayer.name = CONSTRUCTION_LAYER_NAME;
        } else {
            targetLayer = findLayerByName(doc, CONSTRUCTION_LAYER_NAME);
            if (!targetLayer) {
                targetLayer = doc.layers.add();
                targetLayer.name = CONSTRUCTION_LAYER_NAME;
            }
            /* 再実行時は、このスクリプトが生成したものだけクリア / On a rerun, clear only what this script generated */
            clearGeneratedItemsInLayer(targetLayer);
        }

        /* 最前面へ移動 / Bring to front */
        try {
            targetLayer.zOrder(ZOrderMethod.BRINGTOFRONT);
        } catch (e) { }
        return targetLayer;
    }

    /**
     * 補助線の格納先を用意する（［グループ化］がオンなら目印付きのグループ、オフならレイヤーそのもの）
     * @param {Layer} targetLayer - 描画先のレイヤー
     * @param {boolean} shouldGroup - ［グループ化］の値
     * @returns {Layer|GroupItem} 補助線の格納先
     */
    function createLineContainer(targetLayer, shouldGroup) {
        if (!shouldGroup) return targetLayer;
        var lineGroup = targetLayer.groupItems.add();
        lineGroup.name = SCRIPT_MARKER + "_" + getLabel("itemName.lineGroup");
        try { lineGroup.note = SCRIPT_MARKER; } catch (e) { }
        return lineGroup;
    }

    /**
     * 対象パスから補助線（と［円弧から円］の円）を描く。プレビューと確定の両方で使う
     * @param {Layer|GroupItem} targetContainer - 描画先
     * @param {PathItem[]} targetPaths - 対象のパス
     * @param {Object} drawSettings - readDialogSettings() の戻り値と同じ形の設定
     * @param {{left: number, top: number, right: number, bottom: number}} drawBounds - 描画範囲
     * @returns {void}
     */
    function drawConstructionLines(targetContainer, targetPaths, drawSettings, drawBounds) {
        var dedupMap = {}; /* 線のダブり検出用 / for detecting duplicate lines */

        /* 円弧から円（先に）/ Circles from arcs first */
        if (drawSettings.arcToCircle) {
            for (var arcPathIndex = 0; arcPathIndex < targetPaths.length; arcPathIndex++) {
                createCirclesFromArcPath(targetPaths[arcPathIndex], targetContainer, drawSettings, drawBounds, dedupMap);
            }
        }

        for (var pathIndex = 0; pathIndex < targetPaths.length; pathIndex++) {
            var pathItem = targetPaths[pathIndex];
            var pathPoints = pathItem.pathPoints;
            if (!pathPoints || pathPoints.length < 2) continue;

            /* 隣接するアンカーポイントのペア（閉じたパスは最後と最初もつなぐ）/ Adjacent anchor pairs (closed paths also join the last and first) */
            var pointPairs = [];
            for (var i = 0; i < pathPoints.length - 1; i++) {
                pointPairs.push([i, i + 1]);
            }
            if (pathItem.closed && pathPoints.length >= 3) {
                pointPairs.push([pathPoints.length - 1, 0]);
            }

            for (var j = 0; j < pointPairs.length; j++) {
                var startPoint = pathPoints[pointPairs[j][0]];
                var endPoint = pathPoints[pointPairs[j][1]];
                var isStraight = isStraightSegment(startPoint, endPoint);

                /* STRAIGHT は直線セグメントだけ、CURVE は曲線セグメントだけ / STRAIGHT keeps straight segments, CURVE keeps curved ones */
                if (drawSettings.mode === "STRAIGHT" && !isStraight) continue;
                if (drawSettings.mode === "CURVE" && isStraight) continue;

                /* 「円弧から円」ON のときは曲線を円の処理に任せる（ここで結ぶと円弧が直線になる）
                   / With "arc to circle" on, curves are left to the circle routine (joining them here would flatten the arc) */
                if (drawSettings.arcToCircle && !isStraight) continue;

                var startAnchor = startPoint.anchor;
                var endAnchor = endPoint.anchor;

                /* 2点が同じ座標（ゴミパスなど）はスキップ / Skip coincident points (stray paths and the like) */
                if (Math.abs(startAnchor[0] - endAnchor[0]) < 0.001 && Math.abs(startAnchor[1] - endAnchor[1]) < 0.001) continue;

                drawLineAcrossArtboard(targetContainer, startAnchor, endAnchor, drawBounds, drawSettings, dedupMap);
            }
        }
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */

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
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];
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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ボタン行（再利用パーツ）ここまで / End of the reusable button row
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログと各コントロールを作る
     * @returns {Object} ダイアログ（dialog）と各コントロールの参照
     */
    function buildDialog() {
        var extendDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        extendDialog.orientation = "column";
        extendDialog.alignChildren = ["left", "top"];
        extendDialog.margins = DIALOG_MARGINS;

        /* 2カラムレイアウト / Two-column layout */
        var columnsGroup = extendDialog.add("group");
        columnsGroup.orientation = "row";
        columnsGroup.alignChildren = ["fill", "top"];
        columnsGroup.alignment = ["fill", "top"];
        columnsGroup.spacing = COLUMN_SPACING;

        /* 左カラム：補助線を描画 / Left column: draw construction lines */
        var auxLinesPanel = columnsGroup.add("panel", undefined, getLabel("panel.auxLines"));
        applyPanelLayout(auxLinesPanel);
        auxLinesPanel.alignment = ["fill", "top"];

        var auxLinesGroup = auxLinesPanel.add("group");
        auxLinesGroup.orientation = "column";
        auxLinesGroup.alignChildren = ["left", "top"];
        auxLinesGroup.margins = [0, 0, 0, 0];

        var straightOnlyRow = auxLinesGroup.add("group");
        straightOnlyRow.orientation = "row";
        straightOnlyRow.alignChildren = ["left", "center"];
        straightOnlyRow.spacing = 0;

        var straightOnlyCheckbox = straightOnlyRow.add("checkbox", undefined, getLabel("checkbox.straightOnly"));
        straightOnlyCheckbox.helpTip = getLabel("tooltip.straightOnly");
        straightOnlyCheckbox.value = true; /* デフォルトON / on by default */

        var arcToCircleCheckbox = auxLinesGroup.add("checkbox", undefined, getLabel("checkbox.arcToCircle"));
        arcToCircleCheckbox.helpTip = getLabel("tooltip.arcToCircle");
        arcToCircleCheckbox.value = true; /* デフォルトON / on by default */

        /* 円弧オプション（［円弧から円］がオフのときはディム）/ Arc options (dimmed while "arc to circle" is off) */
        var arcOptionsPanel = auxLinesGroup.add("panel", undefined, getLabel("panel.arcOptions"));
        applyPanelLayout(arcOptionsPanel);
        arcOptionsPanel.helpTip = getLabel("tooltip.arcOptions");

        var arcFallbackGroup = arcOptionsPanel.add("group");
        arcFallbackGroup.orientation = "column";
        arcFallbackGroup.alignChildren = ["left", "top"];

        var arcIgnoreRadio = arcFallbackGroup.add("radiobutton", undefined, getLabel("radio.arcIgnore"));
        arcIgnoreRadio.helpTip = getLabel("tooltip.arcIgnore");
        var arcStraightRadio = arcFallbackGroup.add("radiobutton", undefined, getLabel("radio.arcStraight"));
        arcStraightRadio.helpTip = getLabel("tooltip.arcStraight");
        var arcExtendRadio = arcFallbackGroup.add("radiobutton", undefined, getLabel("radio.arcExtend"));
        arcExtendRadio.helpTip = getLabel("tooltip.arcExtend");
        arcIgnoreRadio.value = true; /* デフォルト：無視 / default: ignore */

        arcOptionsPanel.enabled = arcToCircleCheckbox.value;

        /* 線幅（線の単位で入力。既定値は 0.1mm を換算）/ Stroke width (entered in the stroke unit; defaults to 0.1 mm converted) */
        var strokeRow = auxLinesGroup.add("group");
        strokeRow.orientation = "row";
        strokeRow.alignChildren = ["left", "center"];
        strokeRow.spacing = STROKE_ROW_SPACING;

        strokeRow.add("statictext", undefined, labelText("fieldLabel.strokeWidth"));

        var strokeUnit = getUnitInfo("strokeUnits");
        var defaultUnitValue = DEFAULT_STROKE_WIDTH_PT / strokeUnit.pointsPerUnit;
        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var strokeStepperGroup = strokeRow.add("group");
        strokeStepperGroup.orientation = "row";
        strokeStepperGroup.alignChildren = ["left", "center"];
        strokeStepperGroup.spacing = 0;
        strokeStepperGroup.margins = 0;
        var strokeWidthInput;
        var strokeStepper = addStepper(strokeStepperGroup, function () { return strokeWidthInput; }, {
            min: 0,
            /* 線幅の変更をプレビューへ反映する（onChanging と同じ処理）。描き直しで DOM が失敗しても操作は続ける
               / Same as onChanging; keep stepping even if the preview redraw fails */
            onStep: function (numberInput) {
                if (typeof numberInput.onChanging === "function") {
                    try { numberInput.onChanging(); } catch (e) { }
                }
            }
        });
        strokeWidthInput = strokeStepperGroup.add("edittext", undefined, defaultUnitValue.toFixed(3));
        strokeWidthInput.characters = STROKE_INPUT_CHARS;
        bindSteppedArrowKeys(strokeWidthInput, strokeStepper);

        strokeRow.add("statictext", undefined, strokeUnit.label);

        /* 右カラム：オプション / Right column: options */
        var optionsPanel = columnsGroup.add("panel", undefined, getLabel("panel.options"));
        applyPanelLayout(optionsPanel);
        optionsPanel.alignment = ["fill", "top"];

        var groupCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.group"));
        groupCheckbox.value = true; /* デフォルトON / on by default */

        var separateLayerCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.separateLayer"));
        separateLayerCheckbox.helpTip = getLabel("tooltip.separateLayer");
        separateLayerCheckbox.value = true; /* デフォルトON / on by default */

        var guideCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.guide"));
        guideCheckbox.value = false; /* デフォルトOFF / off by default */

        var dedupCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.dedup"));
        dedupCheckbox.helpTip = getLabel("tooltip.dedup");
        dedupCheckbox.value = true; /* デフォルトON / on by default */

        /* 2カラムの下に余白 / Spacer below the two columns */
        extendDialog.add("panel", undefined, undefined);

        /* ボタンエリア（左：プレビュー／右：キャンセル・OK）/ Button row: preview on the left, Cancel/OK on the right */
        var buttonRow = addButtonRow(extendDialog);

        var previewCheckbox = buttonRow.leftGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.value = false;

        var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            dialog: extendDialog,
            straightOnlyCheckbox: straightOnlyCheckbox,
            arcToCircleCheckbox: arcToCircleCheckbox,
            arcOptionsPanel: arcOptionsPanel,
            arcIgnoreRadio: arcIgnoreRadio,
            arcStraightRadio: arcStraightRadio,
            arcExtendRadio: arcExtendRadio,
            strokeWidthInput: strokeWidthInput,
            groupCheckbox: groupCheckbox,
            separateLayerCheckbox: separateLayerCheckbox,
            guideCheckbox: guideCheckbox,
            dedupCheckbox: dedupCheckbox,
            previewCheckbox: previewCheckbox
        };
    }

    /**
     * ダイアログを表示して設定を取得する（プレビューはダイアログ中だけ描き、閉じたら消す）
     * @param {Document} doc - 対象のドキュメント
     * @param {PageItem[]} currentSelection - 選択オブジェクト
     * @param {PathItem[]} targetPaths - 対象のパス
     * @returns {Object|null} 設定（キャンセル時は null）
     */
    function showDialog(doc, currentSelection, targetPaths) {
        var dialogUI = buildDialog();

        /* 線幅の最後の有効値（入力途中で NaN や 0 になったときに使う）/ Last valid stroke width, used while the field is mid-edit */
        var lastValidStrokeWidthPt = DEFAULT_STROKE_WIDTH_PT;

        /* プレビュー用レイヤー名（ダイアログごとに一意）/ Preview layer name (unique per dialog) */
        var previewLayerName = createUniqueLayerName(doc, PREVIEW_LAYER_BASE_NAME);

        /**
         * 線幅の入力を pt に換算する（読めないときは最後の有効値）
         * @returns {number} 線幅（pt）
         */
        function readStrokeWidthPt() {
            var pointsPerUnit = getUnitInfo("strokeUnits").pointsPerUnit;
            var enteredValue = parseLocaleNumber(dialogUI.strokeWidthInput.text);
            if (!(enteredValue > 0)) return lastValidStrokeWidthPt;

            var strokeWidthPt = enteredValue * pointsPerUnit;
            if (!(strokeWidthPt > 0)) return lastValidStrokeWidthPt;

            lastValidStrokeWidthPt = strokeWidthPt;
            return strokeWidthPt;
        }

        /**
         * ダイアログの状態を設定オブジェクトにまとめる
         * @returns {Object} mode / group / separateLayer / guide / dedup / arcToCircle / strokeWidthPt / arcFallback
         */
        function readDialogSettings() {
            return {
                mode: dialogUI.straightOnlyCheckbox.value ? "STRAIGHT" : "CURVE",
                group: dialogUI.groupCheckbox.value,
                separateLayer: dialogUI.separateLayerCheckbox.value,
                guide: dialogUI.guideCheckbox.value,
                dedup: dialogUI.dedupCheckbox.value,
                arcToCircle: dialogUI.arcToCircleCheckbox.value,
                strokeWidthPt: readStrokeWidthPt(),
                arcFallback: dialogUI.arcIgnoreRadio.value ? "IGNORE" : (dialogUI.arcExtendRadio.value ? "EXTEND" : "STRAIGHT")
            };
        }

        /**
         * プレビューを消す
         * @returns {void}
         */
        function clearPreview() {
            removeLayerIfExists(doc, previewLayerName);
            app.redraw();
        }

        /**
         * プレビューを描き直す（既存のプレビューは消して作り直す。ユーザーの元オブジェクトは触らない）
         * @returns {void}
         */
        function updatePreviewFromUI() {
            if (!dialogUI.previewCheckbox.value) return;

            removeLayerIfExists(doc, previewLayerName);

            var previewLayer = doc.layers.add();
            previewLayer.name = previewLayerName;
            try { previewLayer.zOrder(ZOrderMethod.BRINGTOFRONT); } catch (e) { }

            var previewGroup = previewLayer.groupItems.add();
            previewGroup.name = SCRIPT_MARKER + "_PREVIEW";
            try { previewGroup.note = SCRIPT_MARKER; } catch (e) { }

            drawConstructionLines(previewGroup, targetPaths, readDialogSettings(), getDrawBounds(doc, currentSelection));
            app.redraw();
        }

        /**
         * プレビューがオンなら描き直す
         * @returns {void}
         */
        function refreshPreviewIfOn() {
            if (dialogUI.previewCheckbox.value) updatePreviewFromUI();
        }

        /**
         * 線幅が変わったとき、最後の有効値を更新してプレビューへ反映する
         * @returns {void}
         */
        function onStrokeWidthChanged() {
            readStrokeWidthPt();
            refreshPreviewIfOn();
        }

        /**
         * 円弧オプションのラジオを排他で選ぶ（環境差で同時にオンになる事故を防ぐ）
         * @param {string} fallbackMode - "IGNORE" / "STRAIGHT" / "EXTEND"
         * @returns {void}
         */
        function setArcFallback(fallbackMode) {
            dialogUI.arcIgnoreRadio.value = (fallbackMode === "IGNORE");
            dialogUI.arcStraightRadio.value = (fallbackMode === "STRAIGHT");
            dialogUI.arcExtendRadio.value = (fallbackMode === "EXTEND");
        }

        /**
         * 円弧オプションのラジオにクリック処理を付ける
         * @param {RadioButton} fallbackRadio - 対象のラジオ
         * @param {string} fallbackMode - そのラジオが表すモード
         * @returns {void}
         */
        function bindArcFallbackRadio(fallbackRadio, fallbackMode) {
            fallbackRadio.onClick = function () {
                setArcFallback(fallbackMode);
                refreshPreviewIfOn();
            };
        }

        dialogUI.previewCheckbox.onClick = function () {
            if (dialogUI.previewCheckbox.value) updatePreviewFromUI();
            else clearPreview();
        };

        dialogUI.arcToCircleCheckbox.onClick = function () {
            dialogUI.arcOptionsPanel.enabled = dialogUI.arcToCircleCheckbox.value;
            refreshPreviewIfOn();
        };

        var previewTriggerCheckboxes = [dialogUI.straightOnlyCheckbox, dialogUI.groupCheckbox, dialogUI.separateLayerCheckbox, dialogUI.guideCheckbox, dialogUI.dedupCheckbox];
        for (var i = 0; i < previewTriggerCheckboxes.length; i++) {
            previewTriggerCheckboxes[i].onClick = refreshPreviewIfOn;
        }

        bindArcFallbackRadio(dialogUI.arcIgnoreRadio, "IGNORE");
        bindArcFallbackRadio(dialogUI.arcStraightRadio, "STRAIGHT");
        bindArcFallbackRadio(dialogUI.arcExtendRadio, "EXTEND");

        /* 線幅の変更をプレビューに反映し、常に最後の有効値も更新 / Reflect stroke-width edits in the preview and keep the last valid value */
        dialogUI.strokeWidthInput.onChanging = onStrokeWidthChanged;

        prepareDialogWindow(dialogUI.dialog, SCRIPT_NAME);
        var dialogResult = dialogUI.dialog.show();

        /* ダイアログ終了時は必ずプレビューを消す / Always remove the preview when the dialog closes */
        clearPreview();

        return (dialogResult === 1) ? readDialogSettings() : null;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    // =========================================
    // 対象の収集 / Target collection
    // =========================================

    /**
     * 選択内のテキストを一時的に複製してアウトライン化し、アウトライン（グループ等）を返す。
     * 元のテキストは変更しない
     * @param {PageItem[]} items - 選択オブジェクト
     * @returns {PageItem[]} 一時アウトライン（テキストが無ければ空配列）
     */
    function outlineTextFromSelection(items) {
        var outlineRoots = [];

        /* グループの中のテキストも拾う / Text inside groups is included */
        var textFrames = collectSelectionTextFrames(items, { unique: false });
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            try {
                /* 同一レイヤー末尾に複製し、複製だけアウトライン化 / Duplicate to the end of the same layer and outline only the copy */
                var duplicatedText = textFrame.duplicate(textFrame.layer, ElementPlacement.PLACEATEND);
                var outlined = duplicatedText.createOutline();
                /* createOutline は複製を消費するので remove は例外になりうる / createOutline consumes the copy, so remove() may throw */
                try { duplicatedText.remove(); } catch (e) { }

                if (outlined) {
                    outlineRoots.push(outlined);
                }
            } catch (e) {
                /* 失敗しても全体は止めない / Keep going on failure */
            }
        }

        return outlineRoots;
    }

    /**
     * 一時アウトライン（グループ等）を削除する
     * @param {PageItem[]} outlineRoots - outlineTextFromSelection() の戻り値
     * @returns {void}
     */
    function cleanupTempOutlines(outlineRoots) {
        if (!outlineRoots || outlineRoots.length === 0) return;
        for (var i = outlineRoots.length - 1; i >= 0; i--) {
            try {
                if (outlineRoots[i] && outlineRoots[i].remove) outlineRoots[i].remove();
            } catch (e) { }
        }
    }

    // =========================================
    // レイヤー / Layers
    // =========================================

    /**
     * 選択がすべて指定レイヤー上にあるかを返す
     * @param {PageItem[]} items - 選択オブジェクト
     * @param {string} layerName - レイヤー名
     * @returns {boolean} すべてそのレイヤーに属していれば true
     */
    function isSelectionOnLayer(items, layerName) {
        try {
            if (!items || items.length === 0) return false;
            for (var i = 0; i < items.length; i++) {
                if (!items[i]) return false;
                var itemLayer = items[i].layer;
                if (!itemLayer || itemLayer.name !== layerName) return false;
            }
            return true;
        } catch (e) { }
        return false;
    }

    /**
     * 名前でレイヤーを探す
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {Layer|null} 見つかったレイヤー（無ければ null）
     */
    function findLayerByName(doc, layerName) {
        /* getByName は見つからないと例外 / getByName throws when missing */
        try { return doc.layers.getByName(layerName); } catch (e) { }
        return null;
    }

    /**
     * 既存と重ならないレイヤー名を作る（重なれば _2, _3 … を付ける）
     * @param {Document} doc - 対象のドキュメント
     * @param {string} baseName - もとの名前
     * @returns {string} 使われていないレイヤー名
     */
    function createUniqueLayerName(doc, baseName) {
        var layerName = baseName;
        var suffixNumber = 2;
        while (findLayerByName(doc, layerName)) {
            layerName = baseName + "_" + suffixNumber;
            suffixNumber++;
        }
        return layerName;
    }

    /**
     * 指定した名前のレイヤーがあれば削除する
     * @param {Document} doc - 対象のドキュメント
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function removeLayerIfExists(doc, layerName) {
        try {
            var targetLayer = findLayerByName(doc, layerName);
            if (targetLayer) targetLayer.remove();
        } catch (e) { }
    }

    /**
     * このスクリプトが生成したオブジェクトだけを削除する（ユーザーの既存オブジェクトは残す）
     * @param {Layer} targetLayer - 対象のレイヤー
     * @returns {void}
     */
    function clearGeneratedItemsInLayer(targetLayer) {
        if (!targetLayer) return;
        try {
            /* groupItems を後ろから走査 / Walk the groups from the back */
            for (var i = targetLayer.groupItems.length - 1; i >= 0; i--) {
                var groupItem = targetLayer.groupItems[i];
                if (!groupItem) continue;
                var groupNote = "";
                try { groupNote = groupItem.note || ""; } catch (e) { }
                if (groupNote === SCRIPT_MARKER) {
                    try { groupItem.remove(); } catch (e) { }
                    continue;
                }
                var groupName = "";
                try { groupName = groupItem.name || ""; } catch (e) { }
                if (groupName.indexOf(SCRIPT_MARKER) === 0) {
                    try { groupItem.remove(); } catch (e) { }
                }
            }

            /* レイヤー直下に描いた線もあるので pathItems も後ろから走査 / Lines may sit directly on the layer, so walk the paths too */
            for (var j = targetLayer.pathItems.length - 1; j >= 0; j--) {
                var pathItem = targetLayer.pathItems[j];
                if (!pathItem) continue;
                var pathNote = "";
                try { pathNote = pathItem.note || ""; } catch (e) { }
                if (pathNote === SCRIPT_MARKER) {
                    try { pathItem.remove(); } catch (e) { }
                }
            }
        } catch (e) { }
    }

    // =========================================
    // セグメントの判定 / Segment helpers
    // =========================================

    /**
     * 2つのアンカーポイント間が直線（ハンドルが出ていない）かどうかを判定する
     * @param {PathPoint} startPoint - 始点
     * @param {PathPoint} endPoint - 終点
     * @returns {boolean} 直線なら true
     */
    function isStraightSegment(startPoint, endPoint) {
        var epsilon = 0.001;

        /* 始点の出力ハンドル・終点の入力ハンドルがアンカーと同じ位置か / Is each handle on its anchor? */
        var rightDx = Math.abs(startPoint.rightDirection[0] - startPoint.anchor[0]);
        var rightDy = Math.abs(startPoint.rightDirection[1] - startPoint.anchor[1]);
        var leftDx = Math.abs(endPoint.leftDirection[0] - endPoint.anchor[0]);
        var leftDy = Math.abs(endPoint.leftDirection[1] - endPoint.anchor[1]);

        return (rightDx < epsilon && rightDy < epsilon && leftDx < epsilon && leftDy < epsilon);
    }

    /**
     * 2点から線の重複判定キーを作る（向きは無視）
     * @param {number[]} pointA - 端点
     * @param {number[]} pointB - もう一方の端点
     * @returns {string} 重複判定キー
     */
    function makeLineKey(pointA, pointB) {
        var keyA = pointKey(pointA);
        var keyB = pointKey(pointB);
        return (keyA < keyB) ? (keyA + "|" + keyB) : (keyB + "|" + keyA);
    }

    /**
     * 点を 0.001pt 単位に丸めたキーにする（浮動小数の揺れを抑える）
     * @param {number[]} point - [x, y]
     * @returns {string} 点のキー
     */
    function pointKey(point) {
        return (Number(point[0]).toFixed(3) + "," + Number(point[1]).toFixed(3));
    }

    /**
     * 曲線になっている隣接セグメント（どちらかのハンドルが出ているもの）を集める
     * @param {PathItem} pathItem - 対象のパス
     * @returns {Object[]} セグメントの配列（startPoint / endPoint）
     */
    function getCurvedAdjacentSegments(pathItem) {
        var pathPoints = pathItem.pathPoints;
        var pointCount = pathPoints.length;
        var tolerance = 1e-6;
        var isClosed = pathItem.closed;
        var curvedSegments = [];

        for (var i = 0; i < pointCount; i++) {
            if (!isClosed && i === pointCount - 1) break;

            var startPoint = pathPoints[i];
            var endPoint = pathPoints[(i + 1) % pointCount];

            var startRight = startPoint.rightDirection;
            var startAnchor = startPoint.anchor;
            var endLeft = endPoint.leftDirection;
            var endAnchor = endPoint.anchor;

            var startHandleOut = Math.abs(startRight[0] - startAnchor[0]) > tolerance || Math.abs(startRight[1] - startAnchor[1]) > tolerance;
            var endHandleOut = Math.abs(endLeft[0] - endAnchor[0]) > tolerance || Math.abs(endLeft[1] - endAnchor[1]) > tolerance;

            if (startHandleOut || endHandleOut) {
                curvedSegments.push({ startPoint: startPoint, endPoint: endPoint });
            }
        }

        return curvedSegments;
    }

    // =========================================
    // 円弧・円の計算 / Arc and circle math
    // =========================================

    /**
     * 2アンカーの円弧から円の中心を求める（両端の接線に垂直な直線の交点）
     * @param {number[]} startAnchor - 始点のアンカー
     * @param {number[]} startHandle - 始点の出力ハンドル
     * @param {number[]} endAnchor - 終点のアンカー
     * @param {number[]} endHandle - 終点の入力ハンドル
     * @returns {number[]|null} 中心 [x, y]（平行で求まらなければ null）
     */
    function getCircleCenterFrom2AnchorArc(startAnchor, startHandle, endAnchor, endHandle) {
        var startTangent = [startHandle[0] - startAnchor[0], startHandle[1] - startAnchor[1]];
        var endTangent = [endAnchor[0] - endHandle[0], endAnchor[1] - endHandle[1]];

        /* 法線（半径方向）/ Normals (radius directions) */
        return intersectLines(startAnchor, rotate90(startTangent), endAnchor, rotate90(endTangent));
    }

    /**
     * ベクトルを90°回す
     * @param {number[]} vector - [x, y]
     * @returns {number[]} 回したベクトル
     */
    function rotate90(vector) {
        return [-vector[1], vector[0]];
    }

    /**
     * 点と方向で表した2直線の交点を求める
     * @param {number[]} pointA - 直線Aの通る点
     * @param {number[]} directionA - 直線Aの方向
     * @param {number[]} pointB - 直線Bの通る点
     * @param {number[]} directionB - 直線Bの方向
     * @returns {number[]|null} 交点（平行なら null）
     */
    function intersectLines(pointA, directionA, pointB, directionB) {
        var denominator = directionA[0] * directionB[1] - directionA[1] * directionB[0];
        if (Math.abs(denominator) < 1e-9) return null;

        var dx = pointB[0] - pointA[0];
        var dy = pointB[1] - pointA[1];

        var distanceA = (dx * directionB[1] - dy * directionB[0]) / denominator;
        return [pointA[0] + directionA[0] * distanceA, pointA[1] + directionA[1] * distanceA];
    }

    /**
     * 2次元ベクトルの内積
     * @param {number[]} vectorA - [x, y]
     * @param {number[]} vectorB - [x, y]
     * @returns {number} 内積
     */
    function dotProduct(vectorA, vectorB) {
        return vectorA[0] * vectorB[0] + vectorA[1] * vectorB[1];
    }

    /**
     * 2次元ベクトルの長さ
     * @param {number[]} vector - [x, y]
     * @returns {number} 長さ
     */
    function vectorLength(vector) {
        return Math.sqrt(vector[0] * vector[0] + vector[1] * vector[1]);
    }

    /**
     * 2次元ベクトルの差（a − b）
     * @param {number[]} vectorA - [x, y]
     * @param {number[]} vectorB - [x, y]
     * @returns {number[]} 差のベクトル
     */
    function subtractVectors(vectorA, vectorB) {
        return [vectorA[0] - vectorB[0], vectorA[1] - vectorB[1]];
    }

    /**
     * 3次ベジェ曲線上の点を求める
     * @param {number[]} p0 - 始点
     * @param {number[]} p1 - 始点側の制御点
     * @param {number[]} p2 - 終点側の制御点
     * @param {number[]} p3 - 終点
     * @param {number} t - パラメーター（0〜1）
     * @returns {number[]} 曲線上の点
     */
    function cubicBezierPoint(p0, p1, p2, p3, t) {
        var u = 1 - t;
        var uu = u * u;
        var uuu = uu * u;
        var tt = t * t;
        var ttt = tt * t;

        return [
            uuu * p0[0] + 3 * uu * t * p1[0] + 3 * u * tt * p2[0] + ttt * p3[0],
            uuu * p0[1] + 3 * uu * t * p1[1] + 3 * u * tt * p2[1] + ttt * p3[1]
        ];
    }

    /**
     * 円弧とみなせるかを判定する（両端と中点が同じ円上にあり、両端の接線が半径に垂直）
     * @param {number[]} startAnchor - 始点のアンカー
     * @param {number[]} startHandle - 始点の出力ハンドル
     * @param {number[]} endHandle - 終点の入力ハンドル
     * @param {number[]} endAnchor - 終点のアンカー
     * @param {number[]} center - 推定した中心
     * @param {number} radius - 推定した半径
     * @returns {boolean} 円弧とみなせれば true
     */
    function isApproxCircularArc(startAnchor, startHandle, endHandle, endAnchor, center, radius) {
        if (!center || !(radius > 0)) return false;

        /* 半径の1%（最低 0.2pt）を許容 / Tolerance: 1% of the radius, at least 0.2 pt */
        var tolerance = Math.max(0.2, radius * 0.01);

        var startDistance = vectorLength(subtractVectors(startAnchor, center));
        var endDistance = vectorLength(subtractVectors(endAnchor, center));
        if (Math.abs(startDistance - radius) > tolerance) return false;
        if (Math.abs(endDistance - radius) > tolerance) return false;

        var midPoint = cubicBezierPoint(startAnchor, startHandle, endHandle, endAnchor, 0.5);
        var midDistance = vectorLength(subtractVectors(midPoint, center));
        if (Math.abs(midDistance - radius) > tolerance) return false;

        /* 両端で「半径・接線 ≈ 0」/ radius · tangent ≈ 0 at both ends */
        var startTangent = subtractVectors(startHandle, startAnchor);
        var endTangent = subtractVectors(endAnchor, endHandle);
        var startRadius = subtractVectors(startAnchor, center);
        var endRadius = subtractVectors(endAnchor, center);

        /* 接線が短すぎるものは円弧ではないとみなす / Treat vanishing tangents as not an arc */
        if (vectorLength(startTangent) < 1e-6 || vectorLength(endTangent) < 1e-6) return false;

        if (Math.abs(dotProduct(startRadius, startTangent)) > tolerance * vectorLength(startTangent)) return false;
        if (Math.abs(dotProduct(endRadius, endTangent)) > tolerance * vectorLength(endTangent)) return false;

        return true;
    }

    // =========================================
    // 描画 / Drawing
    // =========================================

    /**
     * 補助線（延長線）と同じスタイルを適用する
     * @param {PathItem} pathItem - 対象のパス
     * @param {boolean} shouldGuide - ガイドにするか
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {void}
     */
    function applyAuxStyle(pathItem, shouldGuide, strokeWidthPt) {
        try { pathItem.note = SCRIPT_MARKER; } catch (e) { }
        pathItem.filled = false;
        try { pathItem.fillColor = new NoColor(); } catch (e) { }
        if (shouldGuide) {
            pathItem.stroked = false;
            try { pathItem.guides = true; } catch (e) { }
        } else {
            pathItem.stroked = true;

            var blackColor = new CMYKColor();
            blackColor.cyan = 0;
            blackColor.magenta = 0;
            blackColor.yellow = 0;
            blackColor.black = 100;
            pathItem.strokeColor = blackColor;

            pathItem.strokeWidth = (strokeWidthPt && strokeWidthPt > 0) ? strokeWidthPt : DEFAULT_STROKE_WIDTH_PT;
        }
    }

    /**
     * アンカー同士を結ぶ線分（弦）を描く
     * @param {Layer|GroupItem} targetContainer - 描画先
     * @param {number[]} startAnchor - 始点
     * @param {number[]} endAnchor - 終点
     * @param {boolean} shouldGuide - ガイドにするか
     * @param {number} strokeWidthPt - 線幅（pt）
     * @returns {PathItem|null} 描いた線（失敗時は null）
     */
    function createChordLine(targetContainer, startAnchor, endAnchor, shouldGuide, strokeWidthPt) {
        try {
            var chordLine = targetContainer.pathItems.add();
            chordLine.setEntirePath([startAnchor, endAnchor]);
            chordLine.closed = false;
            applyAuxStyle(chordLine, shouldGuide, strokeWidthPt);
            return chordLine;
        } catch (e) { }
        return null;
    }

    /**
     * 曲線セグメントごとに円弧から円を推定して描く。円弧とみなせないものは円弧オプションに従う
     * @param {PathItem} arcPath - 対象のパス
     * @param {Layer|GroupItem} targetContainer - 描画先
     * @param {Object} drawSettings - 設定（guide / arcFallback / dedup / strokeWidthPt）
     * @param {{left: number, top: number, right: number, bottom: number}} drawBounds - 描画範囲
     * @param {Object} dedupMap - 線のダブり検出用の表
     * @returns {PathItem[]} 描いた円
     */
    function createCirclesFromArcPath(arcPath, targetContainer, drawSettings, drawBounds, dedupMap) {
        var circles = [];
        try {
            var curvedSegments = getCurvedAdjacentSegments(arcPath);
            if (!curvedSegments || curvedSegments.length === 0) return circles;

            for (var i = 0; i < curvedSegments.length; i++) {
                var startAnchor = curvedSegments[i].startPoint.anchor;
                var endAnchor = curvedSegments[i].endPoint.anchor;
                var startHandle = curvedSegments[i].startPoint.rightDirection;
                var endHandle = curvedSegments[i].endPoint.leftDirection;

                var center = getCircleCenterFrom2AnchorArc(startAnchor, startHandle, endAnchor, endHandle);
                if (!center) continue;

                var dx = startAnchor[0] - center[0];
                var dy = startAnchor[1] - center[1];
                var radius = Math.sqrt(dx * dx + dy * dy);
                if (!(radius > 0)) continue;

                /* 正確な円弧の一部でない場合の扱い / Segments that are not true arcs */
                if (!isApproxCircularArc(startAnchor, startHandle, endHandle, endAnchor, center, radius)) {
                    if (drawSettings.arcFallback === "STRAIGHT") {
                        /* 直線：アンカー同士を結ぶ弦 / Straight: the chord between the anchors */
                        createChordLine(targetContainer, startAnchor, endAnchor, drawSettings.guide, drawSettings.strokeWidthPt);
                    } else if (drawSettings.arcFallback === "EXTEND") {
                        /* 直線（延長）：弦を描画範囲いっぱいに延長 / Straight (extend): the chord extended across the drawing area */
                        drawLineAcrossArtboard(targetContainer, startAnchor, endAnchor, drawBounds, drawSettings, dedupMap);
                    }
                    /* IGNORE は何もしない / IGNORE draws nothing */
                    continue;
                }

                var circle = targetContainer.pathItems.ellipse(
                    center[1] + radius,
                    center[0] - radius,
                    radius * 2,
                    radius * 2
                );
                /* 塗りなしは applyAuxStyle が設定する / applyAuxStyle clears the fill */
                applyAuxStyle(circle, drawSettings.guide, drawSettings.strokeWidthPt);
                circles.push(circle);
            }
        } catch (e) { }

        return circles;
    }

    /**
     * 2点を通る直線を描画範囲の端まで延長して描く（［線のダブりを削除］がオンなら同じ線は1本だけ）
     * @param {Layer|GroupItem} targetContainer - 描画先
     * @param {number[]} startAnchor - 通る点
     * @param {number[]} endAnchor - もう一つの通る点
     * @param {{left: number, top: number, right: number, bottom: number}} drawBounds - 描画範囲
     * @param {Object} drawSettings - 設定（guide / dedup / strokeWidthPt）
     * @param {Object} dedupMap - 線のダブり検出用の表
     * @returns {void}
     */
    function drawLineAcrossArtboard(targetContainer, startAnchor, endAnchor, drawBounds, drawSettings, dedupMap) {
        var left = drawBounds.left, top = drawBounds.top, right = drawBounds.right, bottom = drawBounds.bottom;
        var x1 = startAnchor[0], y1 = startAnchor[1];
        var x2 = endAnchor[0], y2 = endAnchor[1];

        var intersections = [];
        var epsilon = 0.001;

        if (Math.abs(x1 - x2) < epsilon) {
            /* 垂直線 / vertical */
            intersections.push([x1, top]);
            intersections.push([x1, bottom]);
        } else if (Math.abs(y1 - y2) < epsilon) {
            /* 水平線 / horizontal */
            intersections.push([left, y1]);
            intersections.push([right, y1]);
        } else {
            /* 傾きがある場合は4辺との交点を調べる / Sloped: test all four edges */
            var slope = (y2 - y1) / (x2 - x1);
            var intercept = y1 - slope * x1;

            var yAtLeft = slope * left + intercept;
            if (yAtLeft <= top + epsilon && yAtLeft >= bottom - epsilon) intersections.push([left, yAtLeft]);

            var yAtRight = slope * right + intercept;
            if (yAtRight <= top + epsilon && yAtRight >= bottom - epsilon) intersections.push([right, yAtRight]);

            var xAtTop = (top - intercept) / slope;
            if (xAtTop >= left - epsilon && xAtTop <= right + epsilon) intersections.push([xAtTop, top]);

            var xAtBottom = (bottom - intercept) / slope;
            if (xAtBottom >= left - epsilon && xAtBottom <= right + epsilon) intersections.push([xAtBottom, bottom]);
        }

        if (intersections.length < 2) return;

        /* 最初の交点と重ならない交点を1つ選ぶ / Pick the first intersection that differs from the first one */
        var uniquePoints = [intersections[0]];
        for (var j = 1; j < intersections.length; j++) {
            var dx = intersections[j][0] - uniquePoints[0][0];
            var dy = intersections[j][1] - uniquePoints[0][1];
            if (Math.sqrt(dx * dx + dy * dy) > epsilon) {
                uniquePoints.push(intersections[j]);
                break;
            }
        }
        if (uniquePoints.length !== 2) return;

        /* 線のダブりを削除（同じ2点で構成される延長線は1本だけ描く）/ Draw each extended line only once */
        if (drawSettings.dedup && dedupMap) {
            var lineKey = makeLineKey(uniquePoints[0], uniquePoints[1]);
            if (dedupMap[lineKey]) return;
            dedupMap[lineKey] = true;
        }
        var newLine = targetContainer.pathItems.add();
        newLine.setEntirePath(uniquePoints);
        newLine.closed = false;
        applyAuxStyle(newLine, drawSettings.guide, drawSettings.strokeWidthPt);
    }

    main();

})();
