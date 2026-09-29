#target illustrator
#targetengine "PresetManagerArtboardEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード名表示とアートボード枠線の表示設定を、ダイアログでまとめて切り替えます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerArtboard.md

### Overview

Switches the artboard-name and artboard-border display preferences from a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerArtboard.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PresetManagerArtboard";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-29";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerArtboard.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerArtboard.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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
        dialogTitle: {
            ja: "アートボード名と枠線の設定",
            en: "Artboard Name & Border Settings"
        },
        OK: {
            ja: "閉じる",
            en: "Close"
        },
        Cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        VideoRuler: {
            ja: "ビデオ定規",
            en: "Video Ruler"
        },

        // プリセット / Preset
        presetDefault: {
            ja: "デフォルト",
            en: "Default"
        },
        presetEmphasis: {
            ja: "強調",
            en: "Emphasis"
        },
        presetLight: {
            ja: "ライト",
            en: "Light"
        },

        // アートボード / Artboard
        panelArtboardTitle: {
            ja: "アートボード",
            en: "Artboard"
        },
        cbShowArtboardName: {
            ja: "アートボード名を表示",
            en: "Show Artboard Name"
        },
        panelArtboardBorderTitle: {
            ja: "アートボードの枠線",
            en: "Artboard Border"
        },
        artboardStrokeColor: {
            ja: "ハイライトのカラー",
            en: "Highlight Color"
        },
        artboardStrokeWidth: {
            ja: "ストロークの幅",
            en: "Stroke Width"
        },
        artboardColorBlack: {
            ja: "ブラック",
            en: "Black"
        },
        artboardColorLightBlue: {
            ja: "ライトブルー",
            en: "Light Blue"
        },
        artboardColorRed: {
            ja: "サーモンピンク",
            en: "Light Red"
        },
        artboardColorGreen: {
            ja: "グリーン",
            en: "Green"
        },
        artboardColorBlue: {
            ja: "ミディアムブルー",
            en: "Medium Blue"
        },
        artboardColorCyan: {
            ja: "シアン",
            en: "Cyan"
        },
        artboardColorMagenta: {
            ja: "マゼンタ",
            en: "Magenta"
        },
        artboardColorYellow: {
            ja: "イエロー",
            en: "Yellow"
        },
        artboardColorGrey: {
            ja: "ライトグレー",
            en: "Light Gray"
        },
    };

    // =========================================
    // UIレイアウトの共通設定 / Shared UI layout
    // =========================================

    var PANEL_MARGINS = [15, 20, 15, 10]; /* パネル余白 [左,上,右,下] / panel margins */
    var PRESET_ROW_BOTTOM_MARGIN = 5;     /* プリセット行の下余白 / preset row bottom margin */

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

    function main() {

        var prefs = app.preferences;

        // =========================================
        // Constants / 定数定義
        // =========================================
        var STROKE_COLOR_PRESETS = [
            { key: "LIGHT_BLUE", label: getLabel("artboardColorLightBlue"), r: 0.29, g: 0.52, b: 1.0 },
            { key: "RED", label: getLabel("artboardColorRed"), r: 1.0, g: 0.29, b: 0.29 },
            { key: "GREEN", label: getLabel("artboardColorGreen"), r: 0.0, g: 0.65, b: 0.31 },
            { key: "BLUE", label: getLabel("artboardColorBlue"), r: 0.0, g: 0.45, b: 0.78 },
            { key: "MAGENTA", label: getLabel("artboardColorMagenta"), r: 1.0, g: 0.0, b: 1.0 },
            { key: "CYAN", label: getLabel("artboardColorCyan"), r: 0.0, g: 1.0, b: 1.0 },
            { key: "GREY", label: getLabel("artboardColorGrey"), r: 0.65, g: 0.65, b: 0.65 },
            { key: "BLACK", label: getLabel("artboardColorBlack"), r: 0.0, g: 0.0, b: 0.0 },
            { key: "YELLOW", label: getLabel("artboardColorYellow"), r: 1.0, g: 1.0, b: 0.0 }
        ];
        var STROKE_COLOR_INDEX = {
            LIGHT_BLUE: 0,
            RED: 1,
            GREEN: 2,
            BLUE: 3,
            MAGENTA: 4,
            CYAN: 5,
            GREY: 6,
            BLACK: 7,
            YELLOW: 8
        };

        // =========================================
        // Utility functions / ユーティリティ関数
        // =========================================

        function getReal(key, fb) {
            try { return prefs.getRealPreference(key); } catch (e) { return fb; }
        }

        function getBool(key, fb) {
            try { return prefs.getBooleanPreference(key); } catch (e) { return fb; }
        }

        function clamp(n, min, max) {
            return Math.max(min, Math.min(max, n));
        }

        function buildStrokeColorNames() {
            var names = [];
            for (var i = 0; i < STROKE_COLOR_PRESETS.length; i++) {
                names.push(STROKE_COLOR_PRESETS[i].label);
            }
            return names;
        }

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

        function getSelectedStrokeWidth() {
            for (var i = 0; i < rbStrokeWidths.length; i++) {
                if (rbStrokeWidths[i].value) return i + 1;
            }
            return 1;
        }

        function applyCurrentSettings() {
            prefs.setBooleanPreference("showArtboardLabelOnCanvas", cbShowArtboardName.value);
            var scIdx = ddStrokeColor.selection ? ddStrokeColor.selection.index : STROKE_COLOR_INDEX.BLACK;
            var scPreset = STROKE_COLOR_PRESETS[scIdx];
            prefs.setRealPreference("ArtboardBBColorRed", scPreset.r);
            prefs.setRealPreference("ArtboardBBColorGreen", scPreset.g);
            prefs.setRealPreference("ArtboardBBColorBlue", scPreset.b);
            prefs.setRealPreference("ArtboardBBWidth", getSelectedStrokeWidth());
            forceScreenRefresh();
        }

        /**
         * 環境設定の変更後に画面を強制再描画（redraw だけではハイライトが更新されないため）
         * Force a redraw after preference changes (redraw alone leaves the highlight stale)
         * ドキュメントが開いていないときは何もしない / Does nothing when no document is open
         * @returns {void}
         */
        function forceScreenRefresh() {
            if (app.documents.length === 0) return;
            app.executeMenuCommand('zoomout');
            app.executeMenuCommand('zoomin');
        }

        function applyPreset(preset) {
            var i;
            if (preset === "default") {
                cbShowArtboardName.value = true;
                ddStrokeColor.selection = STROKE_COLOR_INDEX.BLACK;
                for (i = 0; i < rbStrokeWidths.length; i++) rbStrokeWidths[i].value = (i === 0);
            } else if (preset === "emphasis") {
                cbShowArtboardName.value = false;
                ddStrokeColor.selection = STROKE_COLOR_INDEX.RED;
                for (i = 0; i < rbStrokeWidths.length; i++) rbStrokeWidths[i].value = (i === 2);
            } else if (preset === "light") {
                cbShowArtboardName.value = false;
                ddStrokeColor.selection = STROKE_COLOR_INDEX.GREY;
                for (i = 0; i < rbStrokeWidths.length; i++) rbStrokeWidths[i].value = (i === 0);
            }
        }

        /*
      Build main dialog / ダイアログ生成
    */
        var dlg = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        var mainGroup = dlg.add("group");
        mainGroup.orientation = "column";
        mainGroup.alignChildren = "left";

        /* Preset radio buttons / プリセットラジオボタン */
        var presetRow = mainGroup.add("group");
        presetRow.orientation = "row";
        presetRow.alignment = "center";
        presetRow.margins = [0, 0, 0, PRESET_ROW_BOTTOM_MARGIN];
        var rbPresetDefault = presetRow.add("radiobutton", undefined, getLabel("presetDefault"));
        var rbPresetEmphasis = presetRow.add("radiobutton", undefined, getLabel("presetEmphasis"));
        var rbPresetLight = presetRow.add("radiobutton", undefined, getLabel("presetLight"));

        /*
      Artboard panel / ［アートボード］
    */
        var panelArtboard = mainGroup.add("panel", undefined, getLabel("panelArtboardTitle"));
        panelArtboard.orientation = "column";
        panelArtboard.alignChildren = ["fill", "top"];
        panelArtboard.alignment = ["fill", "top"];
        panelArtboard.margins = PANEL_MARGINS;

        var cbShowArtboardName = panelArtboard.add("checkbox", undefined, getLabel("cbShowArtboardName"));
        cbShowArtboardName.helpTip = LABELS.cbShowArtboardName.ja + " / " + LABELS.cbShowArtboardName.en;

        // Artboard border panel / アートボードの枠線パネル
        var panelArtboardBorder = panelArtboard.add("panel", undefined, getLabel("panelArtboardBorderTitle"));
        panelArtboardBorder.orientation = "column";
        panelArtboardBorder.alignChildren = ["fill", "top"];
        panelArtboardBorder.alignment = ["fill", "top"];
        panelArtboardBorder.margins = PANEL_MARGINS;

        // Stroke color (dropdown) / ストロークのカラー（ドロップダウン）
        var strokeColorRow = panelArtboardBorder.add("group");
        strokeColorRow.orientation = "row";
        strokeColorRow.alignChildren = ["left", "center"];
        strokeColorRow.add("statictext", undefined, labelText("artboardStrokeColor"));

        var ddStrokeColor = strokeColorRow.add("dropdownlist", undefined, buildStrokeColorNames());

        // Stroke width (1-4, radio buttons) / ストロークの幅（1〜4、ラジオボタン）
        var strokeWidthRow = panelArtboardBorder.add("group");
        strokeWidthRow.orientation = "row";
        strokeWidthRow.alignChildren = ["left", "center"];
        strokeWidthRow.add("statictext", undefined, labelText("artboardStrokeWidth"));
        var rbStrokeWidth1 = strokeWidthRow.add("radiobutton", undefined, "1");
        var rbStrokeWidth2 = strokeWidthRow.add("radiobutton", undefined, "2");
        var rbStrokeWidth3 = strokeWidthRow.add("radiobutton", undefined, "3");
        var rbStrokeWidth4 = strokeWidthRow.add("radiobutton", undefined, "4");
        var rbStrokeWidths = [rbStrokeWidth1, rbStrokeWidth2, rbStrokeWidth3, rbStrokeWidth4];

        /* 下部ボタン行（左：ビデオ定規／右：閉じる）/ Bottom button row (left: Video Ruler, right: Close) */
        var buttonRow = addButtonRow(mainGroup);
        var btnVideoRuler = buttonRow.leftGroup.add("button", undefined, getLabel("VideoRuler"));
        var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("OK"), { name: "ok" });

        // =========================================
        // Reflect current values / 値反映
        // =========================================
        cbShowArtboardName.value = !!getBool("showArtboardLabelOnCanvas", false);
        var curSCR = getReal("ArtboardBBColorRed", 0.0);
        var curSCG = getReal("ArtboardBBColorGreen", 0.0);
        var curSCB = getReal("ArtboardBBColorBlue", 0.0);
        var closestIdx = findClosestStrokeColor(curSCR, curSCG, curSCB);
        ddStrokeColor.selection = closestIdx;
        if (ddStrokeColor.selection === null || ddStrokeColor.selection < 0) {
            ddStrokeColor.selection = STROKE_COLOR_INDEX.BLACK;
        }

        var curStrokeWidth = Math.round(getReal("ArtboardBBWidth", 1.0));
        var swIdx = clamp(curStrokeWidth, 1, 4) - 1;
        rbStrokeWidths[swIdx].value = true;

        /* 現在値がいずれかのプリセットと一致していればそのラジオを選ぶ / Select the preset radio that matches the current values */
        if (cbShowArtboardName.value === true && closestIdx === STROKE_COLOR_INDEX.BLACK && curStrokeWidth === 1) {
            rbPresetDefault.value = true;
        } else if (cbShowArtboardName.value === false && closestIdx === STROKE_COLOR_INDEX.RED && curStrokeWidth === 3) {
            rbPresetEmphasis.value = true;
        } else if (cbShowArtboardName.value === false && closestIdx === STROKE_COLOR_INDEX.GREY && curStrokeWidth === 1) {
            rbPresetLight.value = true;
        }

        // =========================================
        // Event wiring / イベント設定
        // =========================================

        rbPresetDefault.onClick = function () {
            applyPreset("default");
            applyCurrentSettings();
        };
        rbPresetEmphasis.onClick = function () {
            applyPreset("emphasis");
            applyCurrentSettings();
        };
        rbPresetLight.onClick = function () {
            applyPreset("light");
            applyCurrentSettings();
        };

        cbShowArtboardName.onClick = function () {
            applyCurrentSettings();
        };

        ddStrokeColor.onChange = function () {
            applyCurrentSettings();
        };

        rbStrokeWidth1.onClick = function () {
            applyCurrentSettings();
        };
        rbStrokeWidth2.onClick = function () {
            applyCurrentSettings();
        };
        rbStrokeWidth3.onClick = function () {
            applyCurrentSettings();
        };
        rbStrokeWidth4.onClick = function () {
            applyCurrentSettings();
        };

        btnVideoRuler.onClick = function () {
            app.executeMenuCommand('videoruler');
        };

        btnOK.onClick = function () {
            applyCurrentSettings();
            dlg.close();
        };

        dlg.center();
        prepareDialogWindow(dlg, SCRIPT_NAME);
        dlg.show();
    }

    main();

})();
