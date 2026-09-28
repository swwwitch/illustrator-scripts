#target illustrator
#targetengine "GridArrangeArtboardsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アートボード名を解析し、行列グリッドとして再配置します。
名前全体が「行番号＋区切り文字（-, _, x）＋列番号」に一致するアートボードを、行列指定として扱います。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GridArrangeArtboards.md

### Overview

Parses the artboard names and re-lays the artboards out as a row-column grid.
An artboard counts as positioned when its whole name matches "row number + separator (-, _, x) + column number".

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GridArrangeArtboards.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "GridArrangeArtboards";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/GridArrangeArtboards.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/GridArrangeArtboards.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* ダイアログの初期値（間隔はアクティブなアートボード幅の1/8） / Dialog defaults (the gap is 1/8 of the active artboard width) */
    var DIALOG_DEFAULTS = {
        exceptionMode: 'rowEnd',            // 'rowEnd'（直前の行の末尾）| 'lastRow'（最後の行にまとめる）
        linkSpacing: true,
        excludeLockedHiddenLayers: false,
        excludeLockedHiddenItems: false,
        changeArtboardOrder: false
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_MARGINS = 16;                    /* ダイアログ外周の余白 / dialog margins */
    var PANEL_MARGINS = [15, 20, 15, 10];       /* パネルの余白 / panel margins */
    var PANEL_SPACING = 6;                      /* パネル内の間隔 / spacing inside panels */
    var SPACING_FIELD_CHARACTERS = 5;           /* 間隔欄の桁数 / width of the gap fields */

    /**
     * パネルを共通スタイルで初期化する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は変更しない）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = "left";
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        if (typeof spacing === "number") {
            targetPanel.spacing = spacing;
        }
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
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
            title: { ja: "アートボードを再配置（行列）", en: "Rearrange Artboards (Grid)" }
        },
        panel: {
            spacing: { ja: "間隔", en: "Spacing" },
            exclusion: { ja: "ロック／非表示を除外", en: "Exclude Locked / Hidden" },
            exception: { ja: "未指定／重複", en: "Unmatched / Duplicate" },
            options: { ja: "オプション", en: "Options" }
        },
        fieldLabel: {
            spacingX: { ja: "列間", en: "Column gap" },
            spacingY: { ja: "行間", en: "Row gap" }
        },
        checkbox: {
            excludeLockedHiddenLayers: { ja: "レイヤー", en: "Layers" },
            excludeLockedHiddenItems: { ja: "オブジェクト", en: "Objects" },
            changeArtboardOrder: { ja: "パネル上の並び順を変更", en: "Reorder in Artboards panel" }
        },
        radio: {
            exceptionRowEnd: { ja: "直前の行の末尾に配置", en: "Append to current row" },
            exceptionLastRow: { ja: "最後の行にまとめる", en: "Collect in final row" }
        },
        tooltip: {
            spacingX: { ja: "左右に隣り合うアートボードのあいだにあける間隔です。", en: "Gap left between horizontally adjacent artboards." },
            spacingY: { ja: "上下に隣り合うアートボードのあいだにあける間隔です。", en: "Gap left between vertically adjacent artboards." },
            linkSpacing: { ja: "列間と同じ値を行間にも使います。", en: "Uses the column gap for the row gap too." },
            excludeLayers: {
                ja: "ロックまたは非表示のレイヤーに載っているオブジェクトは動かしません。",
                en: "Leaves objects on locked or hidden layers where they are."
            },
            excludeItems: {
                ja: "ロックまたは非表示のオブジェクトは動かしません。",
                en: "Leaves locked or hidden objects where they are."
            },
            exceptionRowEnd: {
                ja: "アートボード名から行列を読み取れなかったものを、それぞれの行の末尾に置きます。",
                en: "Puts artboards with no readable row-column at the end of each row."
            },
            exceptionLastRow: {
                ja: "アートボード名から行列を読み取れなかったものを、最終行の次の行にまとめます。",
                en: "Collects artboards with no readable row-column into a row after the last one."
            },
            changeArtboardOrder: {
                ja: "アートボードパネルの並び順も、配置後の順序に合わせて入れ替えます。",
                en: "Also reorders the Artboards panel to match the new layout."
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
            stepDownInteger: { ja: "値を減らす（shift＋クリックで10の倍数へ）", en: "Decrease (Shift-click to snap to 10s)" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noMatch: {
                ja: "「行-列」（例: 1-2）または「接頭辞-番号」（例: banner-1）の形式のアートボード名が見つかりませんでした。",
                en: "No artboard names in '<row><sep><column>' format (e.g., 1-2) or '<prefix><sep><number>' format (e.g., banner-1) were found."
            }
        }
    };

    // =========================================
    // 単位と数値入力 / Units and numeric input
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

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    /**
     * pt値を表示単位に変換する
     * @param {number} points - pt値
     * @param {Object} unitInfo - getUnitInfo() の結果
     * @returns {number} 表示単位の値
     */
    function pointsToDisplayUnit(points, unitInfo) {
        return points / unitInfo.pointsPerUnit;
    }

    /**
     * 表示単位をpt値に変換する
     * @param {number} value - 表示単位の値
     * @param {Object} unitInfo - getUnitInfo() の結果
     * @returns {number} pt値
     */
    function displayUnitToPoints(value, unitInfo) {
        return value * unitInfo.pointsPerUnit;
    }

    /**
     * 表示用の数値を整える（小数点以下3桁まで）
     * @param {number} value - 値
     * @returns {string} 表示用の文字列
     */
    function formatDisplayNumber(value) {
        var rounded = Math.round(value * 1000) / 1000;
        return String(rounded);
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は DIALOG_* / prepareDialogWindow / *DialogLeft* / getSelectionViewSpan の名前
    // 2. スクリプトの先頭（#target の次の行）に #targetengine "<SCRIPT_NAME>Engine" を置く。
    //    #targetengine が無いと $.global が実行ごとに消え、位置を覚えられない。すでにあればそのまま使う
    // 3. ダイアログの show() の直前で prepareDialogWindow(dialog, SCRIPT_NAME) を呼ぶ。
    //    それまでに入れた onShow / onMove / onClose はそのまま生かし、あとに位置の復元・記録をつなぐ
    //      prepareDialogWindow(mainDialog, SCRIPT_NAME);
    //      var dialogResult = mainDialog.show();
    //    同じスクリプトで複数のダイアログを開くときは、2つ目以降のキーを変える（SCRIPT_NAME + "_colorPicker" など）
    //    同じダイアログを何度も開くときも、毎回 show() の直前で呼んでよい（2回目からは選択範囲を測り直すだけ）
    // 4. 初めて開くとき（記録が無いとき）は、スクリプト側の配置（中央・オフセットなど）がそのまま効く
    // 5. 開く位置が選択中のオブジェクトに重なりそうなら左右の反対側へずらす（Illustrator のみ）。
    //    ずらした位置は記録せず、ユーザーが動かしたときだけ記録する
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    var DIALOG_OPACITY = 0.97;       /* ダイアログの不透明度 / dialog opacity */
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

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ボタン行（再利用パーツ） / Button row (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ダイアログを作る関数より前）に貼る。
    //    識別子は BUTTON_ROW_* / addButtonRow
    // 2. ダイアログの最後で行を作り、ボタンは btn 接頭辞の変数で左右のグループに足す（キャンセル → OK の順）
    //      var buttonRow = addButtonRow(dialog);
    //      var btnPreferences = buttonRow.leftGroup.add("button", undefined, getLabel("button.preferences"));
    //      var btnCancel = buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    //      var btnOK = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
    //    左右中央に並べるときは addButtonRow(dialog, { centered: true }) にして、buttonRow.rowGroup に直接足す
    // 3. 行の上の余白は BUTTON_ROW_TOP_MARGIN で決める。左右の余白はダイアログの margins に任せる
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
    // UI の明暗（再利用パーツ） / UI theme (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る（StepperButtons・LinkToggle の部品より前）。識別子は isDarkUI
    // 2. 配色を明暗で切り替えるときは isDarkUI() を1回だけ呼んで定数に控える
    //      var MY_UI_DARK = isDarkUI();
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
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内（ローカライズより前）に貼る。
    //    識別子はすべて STEPPER_* / *Stepper* / *Stepped* の名前なので、既存の名前とはぶつからない
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. コピー先の LABELS.tooltip に stepUp / stepDown / stepUpInteger / stepDownInteger を足す（このファイルの LABELS から写す）。
    //    getLabel() と uiLang はコピー先のものをそのまま使う
    // 3. 数値欄を addSteppedField() で作る。項目名・∧∨・入力欄がひと組で入り、↑↓キーも∧∨と同じ処理で増減する
    //      var widthInput = addSteppedField(parentPanel, {
    //          label: labelText(LABELS.fieldLabel.width), labelWidth: 60,
    //          text: "210 mm", characters: 8, step: 1, min: 1, unit: " mm",
    //          onStep: function (numberInput) { updatePreview(); }
    //      });
    //    値の種類は options で切り分ける:
    //      小数あり（幅・位置など）   … 指定なし（option＋クリックで0.1ずつ）
    //      整数・1以上（段数・個数など）… integer: true, min: 1（0・小数・負数は受け付けず、option＋クリックも1ずつ）
    //      整数・0以上（間隔の数など）  … integer: true, min: 0
    //      範囲つき（％など）           … min: 0, max: 100, unit: "%"
    // 4. 有効／無効は setSteppedFieldEnabled(widthInput, isEnabled)（∧∨のディム表示も切り替わる）。
    //    行・パネルなど親の enabled を切り替えたときは、そのあとで redrawSteppersIn(親) を呼んで∧∨を描き直す
    //    （∧∨は親をたどって無効を判定し、無効の間はクリックも↑↓キーも効かない）
    // 5. 値は parseFloat(widthInput.text) で読む（unit 付きの欄は「210 mm」の形で入っている）
    // 6. この欄に別の↑↓キー処理を付けない（↑↓キーが二重に効く）
    // 既存の edittext をそのまま使うときは、同じ行の group（spacing 0）に addStepper() → edittext の順で置き、
    // bindSteppedArrowKeys(edittext, stepperGroup) を呼ぶ
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
    // ダイアログ / Dialog
    // =========================================

    /**
     * 間隔の入力行（項目名＋入力欄）を追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} labelPath - 項目名のラベルのパス
     * @param {string} tooltipPath - ツールチップのラベルのパス
     * @param {number} initialPoints - 初期値（pt）
     * @param {Object} unitInfo - 定規単位の情報
     * @param {Function} [onStep] - ∧∨・↑↓キーで増減したあとに呼ぶ処理
     * @returns {EditText} 入力欄（∧∨は .stepperGroup で参照できる）
     */
    function addSpacingField(parentGroup, labelPath, tooltipPath, initialPoints, unitInfo, onStep) {
        var fieldGroup = parentGroup.add('group');
        fieldGroup.add('statictext', undefined, labelText(labelPath));

        /* ∧∨と入力欄は隙間0で突き合わせる / butt the stepper against the field */
        var stepperFieldGroup = fieldGroup.add('group');
        stepperFieldGroup.orientation = 'row';
        stepperFieldGroup.alignChildren = ['left', 'center'];
        stepperFieldGroup.spacing = 0;
        stepperFieldGroup.margins = 0;

        var fieldInput;
        var stepperGroup = addStepper(stepperFieldGroup, function () { return fieldInput; }, { min: 0, onStep: onStep });
        fieldInput = stepperFieldGroup.add('edittext', undefined, formatDisplayNumber(pointsToDisplayUnit(initialPoints, unitInfo)));
        fieldInput.helpTip = getLabel(tooltipPath);
        fieldInput.characters = SPACING_FIELD_CHARACTERS;
        fieldInput.stepperGroup = stepperGroup;
        /* ↑↓キーも∧∨と同じ処理で増減する / arrow keys share the stepper's logic */
        bindSteppedArrowKeys(fieldInput, stepperGroup);
        return fieldInput;
    }

    /**
     * 間隔欄の値を pt で読み取る。数値でないか負の値なら初期値を使う
     * @param {EditText} spacingInput - 間隔欄
     * @param {number} fallbackPoints - 読み取れないときの値（pt）
     * @param {Object} unitInfo - 定規単位の情報
     * @returns {number} 間隔（pt）
     */
    function readSpacingPoints(spacingInput, fallbackPoints, unitInfo) {
        var parsedValue = parseFloat(spacingInput.text);
        if (isNaN(parsedValue) || parsedValue < 0) return fallbackPoints;
        return displayUnitToPoints(parsedValue, unitInfo);
    }

    /**
     * チェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelPath - ラベルのパス
     * @param {string} tooltipPath - ツールチップのラベルのパス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} チェックボックス
     */
    function addOptionCheckbox(parentPanel, labelPath, tooltipPath, initialValue) {
        var optionCheckbox = parentPanel.add('checkbox', undefined, getLabel(labelPath));
        optionCheckbox.helpTip = getLabel(tooltipPath);
        optionCheckbox.value = !!initialValue;
        return optionCheckbox;
    }

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // リンクアイコン（再利用パーツ） / Link toggle (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子はすべて LINK_* / *LinkToggle* / *Link* の名前か、描画の下請け関数（buildArcPoints など）
    //    UI の明暗は UITheme 部品の isDarkUI() を使う（先に UITheme の ▼〜▲ も貼っておく）
    // 2. アイコンを addLinkToggle(親, 初期値, 切り替え後の関数) で作る。helpTip はコピー先で付ける
    //      var linkToggle = addLinkToggle(fieldsRowGroup, true, function () { syncFields(); });
    //      linkToggle.helpTip = getLabel(LABELS.tooltip.linkToggle);
    // 3. 連動中かは linkToggle.value で読む。コードから変えるときは setLinkToggleValue(linkToggle, true)
    // 4. 有効／無効は setLinkToggleEnabled(linkToggle, isEnabled)（無効の間はクリックが効かず、薄く描く）
    // 5. 2つの入力欄の右に置くときは、行 group の中に「入力欄を縦に積んだ group」とアイコンを並べ、
    //    行 group の alignChildren を ["left", "center"] にすると上下の中央に来る
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

    // -----------------------------------------
    // リンクアイコンの寸法 / Link toggle metrics
    // -----------------------------------------
    var LINK_ICON_SIZE          = [22, 22]; /* アイコンの大きさ / icon size */
    var LINK_ICON_STROKE        = 1.5;      /* 線幅 / stroke width */
    var LINK_CUT_DIRECTION      = [1, 0];   /* 連動中の左辺の切れ目の向き（水平）/ direction of the left-leg cut when linked (horizontal) */
    var LINK_HOOK_CUT_DIRECTION = [0, 1];   /* 連動中の巻き込みの切れ目の向き（垂直）/ direction of the hook cut when linked (vertical) */
    var LINK_STRAND_COUNT       = 4;        /* 切れ目の向きをそろえるための細い線の本数 / strands used to shape the cuts */
    var LINK_SLASH_CLEARANCE    = 2.2;      /* 連動OFFの斜線とフックの間（22px 基準）/ gap between the slash and the hooks when unlinked */

    // -----------------------------------------
    // リンクアイコンの配色 / Link toggle colors
    // -----------------------------------------
    var LINK_UI_DARK = isDarkUI();
    /* ダイアログの地に重ねる半透明の黒・白（UIの明るさの段階に追従する）。値はステップボタンの配色と同じ
       Translucent overlays that follow the dialog background; same values as the stepper buttons */
    var LINK_PRESSED_COLOR  = LINK_UI_DARK ? [1, 1, 1, 0.12] : [0, 0, 0, 0.13]; /* 連動中の地 / background while linked */
    var LINK_FRAME_COLOR    = LINK_UI_DARK ? [1, 1, 1, 0.07] : [0, 0, 0, 0.10]; /* 連動中の枠 / frame while linked */
    var LINK_ICON_COLOR     = LINK_UI_DARK ? [1, 1, 1, 1]    : [0, 0, 0, 0.70]; /* アイコンの線 / icon strokes */
    var LINK_DIM_ICON_COLOR = LINK_UI_DARK ? [1, 1, 1, 0.20] : [0, 0, 0, 0.25]; /* 無効時の線 / strokes when disabled */

    // -----------------------------------------
    // アイコンを作る・切り替える（外から呼ぶ関数） / Public API
    // -----------------------------------------
    /**
     * 連動の ON／OFF を切り替えるリンクアイコンを追加する（onDraw で自作描画）。
     * クリックで切り替わる。連動中は押し込んだボタンのように地と枠を描く。
     * @param {Group} parent - 追加先
     * @param {boolean} initialValue - 連動の初期値
     * @param {Function} onToggle - 切り替えたあとに呼ぶ関数
     * @returns {Group} アイコン（.value で連動中かを読む）
     */
    function addLinkToggle(parent, initialValue, onToggle) {
        var linkToggle = parent.add("group");
        linkToggle.preferredSize = LINK_ICON_SIZE;
        linkToggle.minimumSize = LINK_ICON_SIZE;
        linkToggle.maximumSize = LINK_ICON_SIZE;
        linkToggle.value = initialValue;

        linkToggle.onDraw = function () {
            var iconGraphics = linkToggle.graphics;
            var iconWidth = LINK_ICON_SIZE[0];
            var iconHeight = LINK_ICON_SIZE[1];
            /* 自作描画は自動でディムにならないため、親もたどって判定する / Custom drawing is not dimmed automatically */
            var isDimmed = !isLinkToggleEnabledInTree(linkToggle);
            /* 連動中は押し込んだボタンのように地と枠を描く / While linked, draw it like a pressed button */
            if (linkToggle.value && !isDimmed) {
                iconGraphics.newPath();
                iconGraphics.rectPath(0, 0, iconWidth, iconHeight);
                iconGraphics.fillPath(iconGraphics.newBrush(iconGraphics.BrushType.SOLID_COLOR, LINK_PRESSED_COLOR));
                iconGraphics.newPath();
                iconGraphics.rectPath(0.5, 0.5, iconWidth - 1, iconHeight - 1);
                iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, LINK_FRAME_COLOR, 1));
            }
            drawLinkIcon(iconGraphics, iconWidth, iconHeight, linkToggle.value, isDimmed ? LINK_DIM_ICON_COLOR : LINK_ICON_COLOR);
        };

        linkToggle.addEventListener("mousedown", function () {
            if (!isLinkToggleEnabledInTree(linkToggle)) return;
            linkToggle.value = !linkToggle.value;
            redrawLinkToggle(linkToggle);
            if (onToggle) onToggle();
        });
        return linkToggle;
    }

    /**
     * 連動の状態をコードから変えて描き直す（onToggle は呼ばない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isLinked - 連動にするなら true
     * @returns {void}
     */
    function setLinkToggleValue(linkToggle, isLinked) {
        if (linkToggle.value === isLinked) return;
        linkToggle.value = isLinked;
        redrawLinkToggle(linkToggle);
    }

    /**
     * アイコンの有効／無効を切り替えて描き直す（変わらないときは描き直さない）
     * @param {Group} linkToggle - addLinkToggle() で作ったアイコン
     * @param {boolean} isEnabled - 有効にするなら true
     * @returns {void}
     */
    function setLinkToggleEnabled(linkToggle, isEnabled) {
        if (linkToggle.enabled === isEnabled) return;
        linkToggle.enabled = isEnabled;
        redrawLinkToggle(linkToggle);
    }

    /**
     * コントロールと親がすべて有効かを判定する（親の無効化は子の enabled に出ないため、親もたどる）
     * @param {Object} control - 判定するコントロール
     * @returns {boolean} すべて有効なら true
     */
    function isLinkToggleEnabledInTree(control) {
        for (var node = control; node; node = node.parent) {
            if (!node.enabled) return false;
        }
        return true;
    }

    /**
     * group の onDraw を呼び直す。group には notify() が無いため、隠して再表示して描き直させる
     * @param {Group} linkToggle - 描き直すアイコン
     * @returns {void}
     */
    function redrawLinkToggle(linkToggle) {
        linkToggle.hide();
        linkToggle.show();
    }

    // -----------------------------------------
    // アイコンの形 / Icon geometry
    // -----------------------------------------
    /**
     * 連動アイコンを描く。Illustrator の［縦横比を固定］に合わせ、連動中は縦につながったチェーン、
     * 連動していないときは上下に分かれたチェーンに斜線を重ねる。座標は 22px 四方を基準に拡大縮小する。
     * @param {ScriptUIGraphics} iconGraphics - 描画先
     * @param {number} iconWidth - 描画範囲の幅
     * @param {number} iconHeight - 描画範囲の高さ
     * @param {boolean} isLinked - 連動中なら true
     * @param {number[]} iconColor - [r, g, b, a]
     * @returns {void}
     */
    function drawLinkIcon(iconGraphics, iconWidth, iconHeight, isLinked, iconColor) {
        var iconScale = Math.min(iconWidth, iconHeight) / 22;
        var offsetX = (iconWidth - 22 * iconScale) / 2;
        var offsetY = (iconHeight - 22 * iconScale) / 2;
        var strokes = isLinked ? buildLinkedChainStrokes() : buildUnlinkedChainStrokes();
        for (var i = 0; i < strokes.length; i++) {
            var strokePoints = strokes[i].points;
            /* newPath() を呼ばないとパスが前の描画に積み重なる / Without newPath() the paths accumulate */
            iconGraphics.newPath();
            for (var j = 0; j < strokePoints.length; j++) {
                var pointX = offsetX + strokePoints[j][0] * iconScale;
                var pointY = offsetY + strokePoints[j][1] * iconScale;
                if (j === 0) iconGraphics.moveTo(pointX, pointY);
                else iconGraphics.lineTo(pointX, pointY);
            }
            iconGraphics.strokePath(iconGraphics.newPen(iconGraphics.PenType.SOLID_COLOR, iconColor, strokes[i].width * iconScale));
        }
    }

    /**
     * 連動中のチェーン（縦に組み合った2つの輪）の線を返す。
     * 上の輪は左辺の途中から上端を回って右辺を下り、下端で内側へ巻き込む。下の輪はそれを180度回したもの。
     * 切れ目の向きをそろえるため、輪を細い線の束にし、両端を延ばしてから直線で切る（左辺は水平、巻き込みは垂直）
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildLinkedChainStrokes() {
        /* 左辺は上端の丸みだけ残して短く切り、下の輪の巻き込みとの間を空ける
           Keep only a stub on the left so it stays clear of the lower ring's hook */
        var upperRing = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360)
            .concat([[14.5, 11.2]])
            .concat(buildArcPoints(11, 11.2, 3.5, 2.3, 0, 115)));
        var ringStart = upperRing[0];
        var ringEnd = upperRing[upperRing.length - 1];
        var extendedRing = extendPolylineEnds(upperRing, LINK_ICON_STROKE);
        /* 延ばした先がどちら側かで、切り捨てる側を決める / The extended tips tell which side to cut away */
        var startOutsideSign = sideOfLine(extendedRing[0], ringStart, LINK_CUT_DIRECTION);
        var endOutsideSign = sideOfLine(extendedRing[extendedRing.length - 1], ringEnd, LINK_HOOK_CUT_DIRECTION);

        var upperStrands = buildStrandStrokes(extendedRing, function (strandPoints) {
            var trimmed = trimPolylineTail(strandPoints, ringEnd, LINK_HOOK_CUT_DIRECTION, endOutsideSign);
            trimmed = trimPolylineTail(trimmed.reverse(), ringStart, LINK_CUT_DIRECTION, startOutsideSign).reverse();
            return [trimmed];
        });
        var strokes = [];
        for (var i = 0; i < upperStrands.length; i++) {
            strokes.push(upperStrands[i]);
            strokes.push({ points: rotatePointsHalfTurn(upperStrands[i].points), width: upperStrands[i].width });
        }
        return strokes;
    }

    /**
     * 中心線を線幅の中で等分した細い線に分け、clipStrand で切った結果を線として返す。
     * @param {Array<number[]>} centerline - 中心線の点列
     * @param {Function} clipStrand - 細い線の点列を受け取り、残す点列の配列を返す関数
     * @returns {Array<{points: Array<number[]>, width: number}>} 細い線ごとの点列と線幅
     */
    function buildStrandStrokes(centerline, clipStrand) {
        var strandWidth = LINK_ICON_STROKE / LINK_STRAND_COUNT;
        var strokes = [];
        for (var k = 0; k < LINK_STRAND_COUNT; k++) {
            /* 線幅の中を等分した位置に細い線を並べる / Lay the strands evenly across the stroke width */
            var strandOffset = -LINK_ICON_STROKE / 2 + strandWidth * (k + 0.5);
            var strandPieces = clipStrand(offsetPolyline(centerline, strandOffset));
            for (var j = 0; j < strandPieces.length; j++) {
                /* 隣の線と少し重ねて隙間を埋める / Overlap neighbours slightly so no seams show */
                if (strandPieces[j].length > 1) strokes.push({ points: strandPieces[j], width: strandWidth * 1.4 });
            }
        }
        return strokes;
    }

    /**
     * 点列の両端を、端の向きのまま length だけ延ばす。
     * @param {Array<number[]>} points - 点列
     * @param {number} length - 延ばす長さ
     * @returns {Array<number[]>} 延ばした点列
     */
    function extendPolylineEnds(points, length) {
        /* from から to の向きへ、to から length 先の点 / point length beyond to, heading from from to to */
        function extendBeyond(from, to) {
            var dx = to[0] - from[0];
            var dy = to[1] - from[1];
            var segmentLength = Math.sqrt(dx * dx + dy * dy) || 1;
            return [to[0] + dx / segmentLength * length, to[1] + dy / segmentLength * length];
        }
        var lastIndex = points.length - 1;
        return [extendBeyond(points[1], points[0])].concat(points, [extendBeyond(points[lastIndex - 1], points[lastIndex])]);
    }

    /**
     * 点が直線のどちら側にあるかを符号で返す。
     * @param {number[]} point - 点
     * @param {number[]} linePoint - 直線上の1点
     * @param {number[]} direction - 直線の向き
     * @returns {number} 正・負で側を表す値
     */
    function sideOfLine(point, linePoint, direction) {
        return direction[0] * (point[1] - linePoint[1]) - direction[1] * (point[0] - linePoint[0]);
    }

    /**
     * 点列の終わり側で、直線より outsideSign の側にはみ出した部分を切り、直線との交点で止める。
     * 輪の別の場所が同じ直線をまたいでも切らないよう、終わりから数点の範囲だけを見る。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} cutPoint - 切る直線上の1点
     * @param {number[]} direction - 切る直線の向き
     * @param {number} outsideSign - 切り捨てる側の符号
     * @returns {Array<number[]>} 切った点列
     */
    function trimPolylineTail(points, cutPoint, direction, outsideSign) {
        var lastIndex = points.length - 1;
        var searchLimit = Math.max(0, lastIndex - 12);
        var index = lastIndex;
        while (index > searchLimit && sideOfLine(points[index], cutPoint, direction) * outsideSign > 0) index--;
        if (index === lastIndex) return points.slice(0);
        var inside = points[index];
        var outside = points[index + 1];
        var insideSide = sideOfLine(inside, cutPoint, direction);
        var ratio = insideSide / (insideSide - sideOfLine(outside, cutPoint, direction));
        return points.slice(0, index + 1).concat([[inside[0] + (outside[0] - inside[0]) * ratio, inside[1] + (outside[1] - inside[1]) * ratio]]);
    }

    /**
     * 連動していないときのチェーン（上下に分かれた輪と斜線）の線を返す。
     * フックは斜線の近くで切る。線の端は進む向きに直角にしか切れないため、フックを細い線の束にして
     * 1本ずつ斜線と平行な境界で切り、切り口が斜線に沿って見えるようにする。
     * @returns {Array<{points: Array<number[]>, width: number}>} 線ごとの点列と線幅（22px 四方の座標）
     */
    function buildUnlinkedChainStrokes() {
        var slashStart = [3.5, 3.5];
        var slashEnd = [18.5, 18.5];
        var upperHook = densifyPoints(buildArcPoints(11, 7, 3.5, 3.5, 180, 360).concat([[14.5, 11.5]]));
        var hooks = [upperHook, rotatePointsHalfTurn(upperHook)];

        /* 斜線の近くの帯を切り取る / Cut away the band around the slash */
        function clipAroundSlash(strandPoints) {
            return clipOutsideBand(strandPoints, slashStart, slashEnd, LINK_SLASH_CLEARANCE);
        }
        var strokes = buildStrandStrokes(hooks[0], clipAroundSlash).concat(buildStrandStrokes(hooks[1], clipAroundSlash));
        strokes.push({ points: [slashStart, slashEnd], width: LINK_ICON_STROKE });
        return strokes;
    }

    /**
     * 点の間隔が 0.5 以下になるよう、線分の間に点を足す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 細かくした点列
     */
    function densifyPoints(points) {
        var densePoints = [points[0]];
        for (var i = 1; i < points.length; i++) {
            var from = points[i - 1];
            var to = points[i];
            var steps = Math.max(1, Math.ceil(Math.sqrt(Math.pow(to[0] - from[0], 2) + Math.pow(to[1] - from[1], 2)) / 0.5));
            for (var j = 1; j <= steps; j++) {
                densePoints.push([from[0] + (to[0] - from[0]) * j / steps, from[1] + (to[1] - from[1]) * j / steps]);
            }
        }
        return densePoints;
    }

    /**
     * 点列を、進む向きの左側へ offset だけずらした点列を返す（負の値なら右側）。
     * @param {Array<number[]>} points - 点列
     * @param {number} offset - ずらす距離
     * @returns {Array<number[]>} ずらした点列
     */
    function offsetPolyline(points, offset) {
        var shifted = [];
        for (var i = 0; i < points.length; i++) {
            var before = points[Math.max(0, i - 1)];
            var after = points[Math.min(points.length - 1, i + 1)];
            var tangentX = after[0] - before[0];
            var tangentY = after[1] - before[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY) || 1;
            shifted.push([points[i][0] - tangentY / tangentLength * offset, points[i][1] + tangentX / tangentLength * offset]);
        }
        return shifted;
    }

    /**
     * 直線（線分を延長したもの）から clearance 未満の帯に入る部分を切り取り、残りを点列に分けて返す。
     * 帯の境界で線分を補間して切るので、切り口は直線と平行にそろう。
     * @param {Array<number[]>} points - 点列
     * @param {number[]} lineStart - 直線上の1点
     * @param {number[]} lineEnd - 直線上のもう1点
     * @param {number} clearance - 空ける距離
     * @returns {Array<Array<number[]>>} 帯の外側に残った点列（2点未満のものは除く）
     */
    function clipOutsideBand(points, lineStart, lineEnd, clearance) {
        var directionX = lineEnd[0] - lineStart[0];
        var directionY = lineEnd[1] - lineStart[1];
        var directionLength = Math.sqrt(directionX * directionX + directionY * directionY);

        /* 直線からの符号付き距離 / signed distance from the line */
        function signedDistance(point) {
            return (directionX * (point[1] - lineStart[1]) - directionY * (point[0] - lineStart[0])) / directionLength;
        }
        /* 2点の間で、距離が boundary になる点 / point between two points where the distance equals boundary */
        function interpolateAt(from, to, fromDistance, toDistance, boundary) {
            var ratio = (boundary - fromDistance) / (toDistance - fromDistance);
            return [from[0] + (to[0] - from[0]) * ratio, from[1] + (to[1] - from[1]) * ratio];
        }

        var pieces = [];
        var currentPiece = [];
        for (var i = 0; i < points.length; i++) {
            var distance = signedDistance(points[i]);
            var isOutside = Math.abs(distance) >= clearance;
            if (i > 0) {
                var previousDistance = signedDistance(points[i - 1]);
                var wasOutside = Math.abs(previousDistance) >= clearance;
                if (wasOutside && !isOutside) {
                    /* 帯に入る: 境界で止める / entering the band: stop at the boundary */
                    currentPiece.push(interpolateAt(points[i - 1], points[i], previousDistance, distance, previousDistance > 0 ? clearance : -clearance));
                    if (currentPiece.length > 1) pieces.push(currentPiece);
                    currentPiece = [];
                } else if (!wasOutside && isOutside) {
                    /* 帯から出る: 境界から始める / leaving the band: start at the boundary */
                    currentPiece = [interpolateAt(points[i - 1], points[i], previousDistance, distance, distance > 0 ? clearance : -clearance)];
                }
            }
            if (isOutside) currentPiece.push(points[i]);
        }
        if (currentPiece.length > 1) pieces.push(currentPiece);
        return pieces;
    }

    /**
     * 楕円弧の点列を返す（角度は右が0度、下が90度の画面座標）。
     * @param {number} centerX - 中心X
     * @param {number} centerY - 中心Y
     * @param {number} radiusX - 横の半径
     * @param {number} radiusY - 縦の半径
     * @param {number} startDegrees - 開始角度
     * @param {number} endDegrees - 終了角度
     * @returns {Array<number[]>} 点列
     */
    function buildArcPoints(centerX, centerY, radiusX, radiusY, startDegrees, endDegrees) {
        var arcSteps = 12;
        var arcPoints = [];
        for (var i = 0; i <= arcSteps; i++) {
            var angle = (startDegrees + (endDegrees - startDegrees) * i / arcSteps) * Math.PI / 180;
            arcPoints.push([centerX + radiusX * Math.cos(angle), centerY + radiusY * Math.sin(angle)]);
        }
        return arcPoints;
    }

    /**
     * 点列を 22px 四方の中心で180度回す。
     * @param {Array<number[]>} points - 点列
     * @returns {Array<number[]>} 回した点列
     */
    function rotatePointsHalfTurn(points) {
        var rotated = [];
        for (var i = 0; i < points.length; i++) {
            rotated.push([22 - points[i][0], 22 - points[i][1]]);
        }
        return rotated;
    }

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // リンクアイコン（再利用パーツ）ここまで / End of the reusable link toggle
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * 設定ダイアログを表示し、ユーザーが入力した値を返す
     * @param {Object} defaultSettings - 初期値（間隔は pt）
     * @returns {Object|null} 設定。キャンセル時は null
     */
    function showSettingsDialog(defaultSettings) {
        var settingsDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        settingsDialog.orientation = 'column';
        settingsDialog.alignChildren = 'fill';
        settingsDialog.margins = DIALOG_MARGINS;

        var rulerUnitInfo = getUnitInfo();

        /* 間隔パネル（列間・行間と連動アイコン）/ Spacing panel (column/row gaps with the link icon) */
        var spacingPanel = settingsDialog.add('panel', undefined, getLabel('panel.spacing') + '（' + rulerUnitInfo.label + '）');
        setupPanel(spacingPanel, PANEL_SPACING);

        /* 列間・行間の2行の右に連動アイコンを置く / Put the link icon to the right of the two gap rows */
        var spacingContentGroup = spacingPanel.add('group');
        spacingContentGroup.orientation = 'row';
        spacingContentGroup.alignChildren = ['left', 'center'];

        var spacingInputColumn = spacingContentGroup.add('group');
        spacingInputColumn.orientation = 'column';
        spacingInputColumn.alignChildren = 'left';
        spacingInputColumn.spacing = PANEL_SPACING;

        /* 列間を増減したら、連動時は行間にもミラーする / mirror the column gap to the row gap when linked */
        var spacingXInput = addSpacingField(spacingInputColumn, 'fieldLabel.spacingX', 'tooltip.spacingX', defaultSettings.spacingX, rulerUnitInfo,
            function () { syncLinkedSpacingValue(); });
        spacingXInput.active = true;
        var spacingYInput = addSpacingField(spacingInputColumn, 'fieldLabel.spacingY', 'tooltip.spacingY', defaultSettings.spacingY, rulerUnitInfo);

        var linkToggle = addLinkToggle(spacingContentGroup, !!defaultSettings.linkSpacing, function () {
            syncLinkedSpacingValue();
            syncLinkedSpacingInputState();
        });
        linkToggle.helpTip = getLabel('tooltip.linkSpacing');

        /* 連動時は列間を行間にミラーし、行間欄をディムする / When linked, mirror the column gap to the row gap and dim the row field */
        function syncLinkedSpacingValue() {
            if (linkToggle.value) spacingYInput.text = spacingXInput.text;
        }
        function syncLinkedSpacingInputState() {
            spacingYInput.enabled = !linkToggle.value;
            spacingYInput.stepperGroup.enabled = spacingYInput.enabled;
            redrawSteppersIn(spacingYInput.stepperGroup);
        }
        syncLinkedSpacingInputState();

        spacingXInput.onChanging = syncLinkedSpacingValue;

        /* ロック／非表示の除外パネル / Locked / hidden exclusion panel */
        var exclusionPanel = settingsDialog.add('panel', undefined, getLabel('panel.exclusion'));
        setupPanel(exclusionPanel, PANEL_SPACING);
        var excludeLayersCheckbox = addOptionCheckbox(exclusionPanel, 'checkbox.excludeLockedHiddenLayers', 'tooltip.excludeLayers', defaultSettings.excludeLockedHiddenLayers);
        var excludeItemsCheckbox = addOptionCheckbox(exclusionPanel, 'checkbox.excludeLockedHiddenItems', 'tooltip.excludeItems', defaultSettings.excludeLockedHiddenItems);

        /* 未指定／重複の扱いパネル / Unmatched / duplicate handling panel */
        var exceptionPanel = settingsDialog.add('panel', undefined, getLabel('panel.exception'));
        setupPanel(exceptionPanel, PANEL_SPACING);
        var rowEndRadio = exceptionPanel.add('radiobutton', undefined, getLabel('radio.exceptionRowEnd'));
        rowEndRadio.helpTip = getLabel('tooltip.exceptionRowEnd');
        var lastRowRadio = exceptionPanel.add('radiobutton', undefined, getLabel('radio.exceptionLastRow'));
        lastRowRadio.helpTip = getLabel('tooltip.exceptionLastRow');
        if (defaultSettings.exceptionMode === 'lastRow') {
            lastRowRadio.value = true;
        } else {
            rowEndRadio.value = true;
        }

        /* オプションパネル / Options panel */
        var optionsPanel = settingsDialog.add('panel', undefined, getLabel('panel.options'));
        setupPanel(optionsPanel, PANEL_SPACING);
        var changeArtboardOrderCheckbox = addOptionCheckbox(optionsPanel, 'checkbox.changeArtboardOrder', 'tooltip.changeArtboardOrder', defaultSettings.changeArtboardOrder);

        /* OK / キャンセルボタン / OK and Cancel buttons */
        var buttonRow = addButtonRow(settingsDialog);
        var btnCancel = buttonRow.rightGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        var btnOK = buttonRow.rightGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        prepareDialogWindow(settingsDialog, SCRIPT_NAME);
        if (settingsDialog.show() !== 1) return null;

        return {
            spacingX: readSpacingPoints(spacingXInput, defaultSettings.spacingX, rulerUnitInfo),
            spacingY: readSpacingPoints(spacingYInput, defaultSettings.spacingY, rulerUnitInfo),
            exceptionMode: lastRowRadio.value ? 'lastRow' : 'rowEnd',
            linkSpacing: linkToggle.value,
            excludeLockedHiddenLayers: excludeLayersCheckbox.value,
            excludeLockedHiddenItems: excludeItemsCheckbox.value,
            changeArtboardOrder: changeArtboardOrderCheckbox.value
        };
    }

    // =========================================
    // 配置処理 / Placement
    // =========================================

    /**
     * アクティブアートボード幅の1/8を初期間隔として取得する
     * @param {Document} documentRef - 対象のドキュメント
     * @returns {number} 間隔（pt）
     */
    function getDefaultSpacingFromActiveArtboard(documentRef) {
        var activeArtboardIndex = documentRef.artboards.getActiveArtboardIndex();
        var activeArtboardRect = documentRef.artboards[activeArtboardIndex].artboardRect;
        var activeArtboardWidth = activeArtboardRect[2] - activeArtboardRect[0];
        return Math.round(activeArtboardWidth / 8);
    }

    /**
     * アートボード名を解析する
     * 「数字 + 区切り文字 + 数字」（区切り文字 -, _, x。大文字小文字無視）を行列指定とし、
     * 行列指定の形でない場合のみ「接頭辞 + 区切り文字 + 数字」を接頭辞指定として判定する。
     * 0行・0列・0番、数字だけの接頭辞（行列指定と紛らわしい）は無効
     * @param {string} artboardName - アートボード名
     * @returns {Object|null} { matchType: 'rowColumn' | 'prefixNumber', rowNumber, columnNumber, prefixName }。どちらでもなければ null
     */
    function parseArtboardName(artboardName) {
        var rowColumnMatch = artboardName.match(/^(\d+)[-_x](\d+)$/i);
        if (rowColumnMatch) {
            var rowNumber = parseInt(rowColumnMatch[1], 10);
            var columnNumber = parseInt(rowColumnMatch[2], 10);
            if (rowNumber < 1 || columnNumber < 1) return null;
            return { matchType: 'rowColumn', rowNumber: rowNumber, columnNumber: columnNumber, prefixName: '' };
        }

        var prefixNumberMatch = artboardName.match(/^(.+?)[-_x](\d+)$/i);
        if (!prefixNumberMatch || /^\d+$/.test(prefixNumberMatch[1])) return null;
        var prefixColumnNumber = parseInt(prefixNumberMatch[2], 10);
        if (prefixColumnNumber < 1) return null;
        return { matchType: 'prefixNumber', rowNumber: 0, columnNumber: prefixColumnNumber, prefixName: prefixNumberMatch[1] };
    }

    /**
     * 第1パス：アートボードの情報を収集し、行列指定と接頭辞指定を判定する
     * @param {Artboards} artboards - アートボードのコレクション
     * @returns {Object} placementItems と、最大列番号・最大行番号・接頭辞の出現順などをまとめたグリッド情報
     */
    function collectArtboardPlacements(artboards) {
        var gridInfo = {
            placementItems: [],
            maxColumnNumber: 0,
            maxRowNumber: 0,
            prefixOrder: [],
            prefixRowOffsetByName: {},
            hasMatchedArtboard: false
        };

        for (var artboardIndex = 0; artboardIndex < artboards.length; artboardIndex++) {
            var artboard = artboards[artboardIndex];
            var artboardRect = artboard.artboardRect; // [left, top, right, bottom]
            var placementItem = {
                artboard: artboard,
                width: artboardRect[2] - artboardRect[0],
                height: artboardRect[1] - artboardRect[3],
                matched: false,
                matchType: '',
                prefixName: '',
                rowNumber: 0,
                columnNumber: 0,
                assignedRow: 0,
                assignedColumn: 0
            };

            var nameMatch = parseArtboardName(artboard.name);
            if (nameMatch) {
                placementItem.matched = true;
                placementItem.matchType = nameMatch.matchType;
                placementItem.prefixName = nameMatch.prefixName;
                placementItem.rowNumber = nameMatch.rowNumber;
                placementItem.columnNumber = nameMatch.columnNumber;

                if (nameMatch.matchType === 'rowColumn') {
                    if (nameMatch.rowNumber > gridInfo.maxRowNumber) gridInfo.maxRowNumber = nameMatch.rowNumber;
                } else if (gridInfo.prefixRowOffsetByName[nameMatch.prefixName] === undefined) {
                    /* 接頭辞は出現順に1行ずつ割り当てる / one row per prefix, in order of appearance */
                    gridInfo.prefixRowOffsetByName[nameMatch.prefixName] = gridInfo.prefixOrder.length + 1;
                    gridInfo.prefixOrder.push(nameMatch.prefixName);
                }
                if (nameMatch.columnNumber > gridInfo.maxColumnNumber) gridInfo.maxColumnNumber = nameMatch.columnNumber;
                gridInfo.hasMatchedArtboard = true;
            }
            gridInfo.placementItems.push(placementItem);
        }

        /* 接頭辞の行は通常の行列指定の下に続く / prefix rows follow the numeric rows */
        gridInfo.maxMatchedRowNumber = gridInfo.maxRowNumber + gridInfo.prefixOrder.length;
        return gridInfo;
    }

    /**
     * 第2パス：順番に走査して行・列番号を割り当てる
     * 例外処理モード / Exception handling modes:
     *   'rowEnd'  : 非一致名は直前にマッチした行に追従し、列は (maxColumnNumber + 1) から右へ並べる
     *               先頭の非一致は行1に配置 / Leading unmatched goes to row 1
     *   'lastRow' : 非一致名は (maxMatchedRowNumber + 1) 行目に、出現順に列 1, 2, 3, ... で並べる
     * 接頭辞指定は、通常の行列指定の下に接頭辞ごとの行として配置する。
     * 重複名は元の (row, col) と衝突した時点で、その行の末尾へ押し出す。
     * @param {Object} gridInfo - collectArtboardPlacements() の結果（assignedRow / assignedColumn を書き込む）
     * @param {string} exceptionMode - 'rowEnd' | 'lastRow'
     * @returns {void}
     */
    function assignGridSlots(gridInfo, exceptionMode) {
        var occupiedSlots = {};                                     // "row,col" → true
        var nextFreeColumnByRow = {};                               // row → next column to try
        var rowEndFirstColumn = gridInfo.maxColumnNumber + 1;       // 行末に押し出すときの開始列 / first column at a row end
        var exceptionRowNumber = gridInfo.maxMatchedRowNumber + 1;  // 'lastRow' 用の行番号 / row for 'lastRow'
        var currentRowNumber = 1;

        for (var assignIndex = 0; assignIndex < gridInfo.placementItems.length; assignIndex++) {
            var assignTarget = gridInfo.placementItems[assignIndex];

            if (assignTarget.matched) {
                var targetRowNumber = assignTarget.rowNumber;
                if (assignTarget.matchType === 'prefixNumber') {
                    targetRowNumber = gridInfo.maxRowNumber + gridInfo.prefixRowOffsetByName[assignTarget.prefixName];
                }

                currentRowNumber = targetRowNumber;
                assignTarget.assignedRow = targetRowNumber;
                var requestedKey = targetRowNumber + ',' + assignTarget.columnNumber;
                if (!occupiedSlots[requestedKey]) {
                    /* 通常配置 / Normal placement */
                    occupiedSlots[requestedKey] = true;
                    assignTarget.assignedColumn = assignTarget.columnNumber;
                } else {
                    /* 重複 → その行の末尾へ / Duplicate → push to row end */
                    assignTarget.assignedColumn = reserveNextFreeColumn(targetRowNumber, rowEndFirstColumn, occupiedSlots, nextFreeColumnByRow);
                }
            } else if (exceptionMode === 'lastRow') {
                assignTarget.assignedRow = exceptionRowNumber;
                assignTarget.assignedColumn = reserveNextFreeColumn(exceptionRowNumber, 1, occupiedSlots, nextFreeColumnByRow);
            } else {
                /* 'rowEnd' */
                assignTarget.assignedRow = currentRowNumber;
                assignTarget.assignedColumn = reserveNextFreeColumn(currentRowNumber, rowEndFirstColumn, occupiedSlots, nextFreeColumnByRow);
            }
        }
    }

    /**
     * 第3〜5パス：列ごとの最大幅・行ごとの最大高さから、各アートボードの新しい位置と移動量を求める
     * 例外領域に入る境界（未指定・重複の列、未指定の行）では間隔を3倍にする。
     * IllustratorのY座標は上がプラス、下がマイナス。
     * @param {Object[]} placementItems - 行列を割り当てたアートボード情報（oldRect / newLeft / newTop / deltaX / deltaY を書き込む）
     * @param {number} spacingX - 列間（pt）
     * @param {number} spacingY - 行間（pt）
     * @param {number} columnBoundary - 間隔を3倍にする列番号
     * @param {number} rowBoundary - 間隔を3倍にする行番号
     * @returns {void}
     */
    function computeArtboardDestinations(placementItems, spacingX, spacingY, columnBoundary, rowBoundary) {
        var originX = 0;    // 配置の開始X座標 / X origin
        var originY = 0;    // 配置の開始Y座標 / Y origin

        /* 列ごとの最大幅・行ごとの最大高さを集計する / Per-column max width and per-row max height */
        var columnWidthByNumber = {};
        var rowHeightByNumber = {};
        for (var aggregateIndex = 0; aggregateIndex < placementItems.length; aggregateIndex++) {
            var aggregateItem = placementItems[aggregateIndex];
            var columnKey = aggregateItem.assignedColumn;
            var rowKey = aggregateItem.assignedRow;
            if (columnWidthByNumber[columnKey] === undefined || aggregateItem.width > columnWidthByNumber[columnKey]) {
                columnWidthByNumber[columnKey] = aggregateItem.width;
            }
            if (rowHeightByNumber[rowKey] === undefined || aggregateItem.height > rowHeightByNumber[rowKey]) {
                rowHeightByNumber[rowKey] = aggregateItem.height;
            }
        }

        /* 使われている列番号・行番号の昇順で累積オフセットを計算する / Cumulative offsets over the used numbers */
        var columnOffsetByNumber = computeCumulativeOffsets(
            collectSortedNumericKeys(columnWidthByNumber), columnWidthByNumber, spacingX, columnBoundary
        );
        var rowOffsetByNumber = computeCumulativeOffsets(
            collectSortedNumericKeys(rowHeightByNumber), rowHeightByNumber, spacingY, rowBoundary
        );

        for (var placeIndex = 0; placeIndex < placementItems.length; placeIndex++) {
            var placeItem = placementItems[placeIndex];
            var oldRect = placeItem.artboard.artboardRect;
            placeItem.oldRect = oldRect;
            placeItem.newLeft = originX + columnOffsetByNumber[placeItem.assignedColumn];
            placeItem.newTop = originY - rowOffsetByNumber[placeItem.assignedRow];
            placeItem.deltaX = placeItem.newLeft - oldRect[0];
            placeItem.deltaY = placeItem.newTop - oldRect[1];
        }
    }

    /**
     * 第6パス：アートボード上のオブジェクトを、アートボードと同じだけ移動する
     * オブジェクトの中心が元のアートボード矩形内にあるものを対象にする。
     * @param {Layers} layerCollection - ドキュメントのレイヤー
     * @param {Object[]} placementItems - 移動量を求めたアートボード情報
     * @param {boolean} excludeLockedHiddenLayers - ロック／非表示レイヤーのオブジェクトを動かさないか
     * @param {boolean} excludeLockedHiddenItems - ロック／非表示のオブジェクトを動かさないか
     * @returns {void}
     */
    function translateItemsWithArtboards(layerCollection, placementItems, excludeLockedHiddenLayers, excludeLockedHiddenItems) {
        var itemTranslations = [];
        appendItemTranslations(layerCollection, placementItems, itemTranslations, excludeLockedHiddenLayers, excludeLockedHiddenItems);
        for (var translateIndex = 0; translateIndex < itemTranslations.length; translateIndex++) {
            var pendingMove = itemTranslations[translateIndex];
            if (pendingMove.deltaX === 0 && pendingMove.deltaY === 0) continue;
            try {
                pendingMove.item.translate(pendingMove.deltaX, pendingMove.deltaY);
            } catch (translateError) {
                /* ロック・非表示などで失敗したオブジェクトはスキップ
                   Skip items that can't be translated (locked, hidden layer, etc.) */
            }
        }
    }

    /**
     * 第7パス：アートボードの矩形を更新する
     * @param {Object[]} placementItems - 新しい位置を求めたアートボード情報
     * @returns {void}
     */
    function applyArtboardRects(placementItems) {
        for (var applyIndex = 0; applyIndex < placementItems.length; applyIndex++) {
            var applyItem = placementItems[applyIndex];
            var newRight = applyItem.newLeft + applyItem.width;
            var newBottom = applyItem.newTop - applyItem.height;
            applyItem.artboard.artboardRect = [applyItem.newLeft, applyItem.newTop, newRight, newBottom];
        }
    }

    /**
     * アートボード名を解析してグリッド状に再配置する
     * 「行-列」は行列指定、「接頭辞-番号」は接頭辞ごとの行として扱い、
     * 列幅・行高は「列ごと・行ごとの最大サイズ」で計算する。
     * @returns {void}
     */
    function arrangeArtboardsByRowColumnName() {
        if (app.documents.length === 0) {
            return;
        }

        var documentRef = app.activeDocument;
        var defaultSpacing = getDefaultSpacingFromActiveArtboard(documentRef);

        var arrangeSettings = showSettingsDialog({
            spacingX: defaultSpacing,
            spacingY: defaultSpacing,
            exceptionMode: DIALOG_DEFAULTS.exceptionMode,
            linkSpacing: DIALOG_DEFAULTS.linkSpacing,
            excludeLockedHiddenLayers: DIALOG_DEFAULTS.excludeLockedHiddenLayers,
            excludeLockedHiddenItems: DIALOG_DEFAULTS.excludeLockedHiddenItems,
            changeArtboardOrder: DIALOG_DEFAULTS.changeArtboardOrder
        });
        if (!arrangeSettings) return;

        var gridInfo = collectArtboardPlacements(documentRef.artboards);
        if (!gridInfo.hasMatchedArtboard) {
            alert(getLabel('alert.noMatch'));
            return;
        }

        assignGridSlots(gridInfo, arrangeSettings.exceptionMode);
        computeArtboardDestinations(gridInfo.placementItems, arrangeSettings.spacingX, arrangeSettings.spacingY,
            gridInfo.maxColumnNumber + 1, gridInfo.maxMatchedRowNumber + 1);

        /* オブジェクトを先に動かしてからアートボードを動かす / Move the objects first, then the artboards */
        translateItemsWithArtboards(documentRef.layers, gridInfo.placementItems,
            arrangeSettings.excludeLockedHiddenLayers, arrangeSettings.excludeLockedHiddenItems);
        applyArtboardRects(gridInfo.placementItems);

        /* 第8パス: アートボードパネル上の順番を、行→列の順に並べ替える / Reorder the Artboards panel row-then-column */
        if (arrangeSettings.changeArtboardOrder) {
            reorderArtboardsByGridOrder(documentRef, gridInfo.placementItems);
        }

        /* 画面を全体表示に更新 / Fit all in view */
        app.executeMenuCommand('fitall');
    }

    // =========================================
    // 配置ヘルパー / Placement helpers
    // =========================================

    /**
     * 指定行の firstColumn 以降で空きスロットを予約する
     * @param {number} rowNumber - 行番号
     * @param {number} firstColumn - その行で最初に試す列番号
     * @param {Object} occupiedSlots - "row,col" → true の表
     * @param {Object} nextFreeColumnByRow - 行番号 → 次に試す列番号の表
     * @returns {number} 予約した列番号
     */
    function reserveNextFreeColumn(rowNumber, firstColumn, occupiedSlots, nextFreeColumnByRow) {
        if (nextFreeColumnByRow[rowNumber] === undefined) {
            nextFreeColumnByRow[rowNumber] = firstColumn;
        }
        while (true) {
            var candidateColumn = nextFreeColumnByRow[rowNumber]++;
            var candidateSlotKey = rowNumber + ',' + candidateColumn;
            if (!occupiedSlots[candidateSlotKey]) {
                occupiedSlots[candidateSlotKey] = true;
                return candidateColumn;
            }
        }
    }

    /**
     * オブジェクトの数値キーを昇順で取り出す
     * @param {Object} numericKeyedObject - 数値キーの表
     * @returns {number[]} 昇順のキー
     */
    function collectSortedNumericKeys(numericKeyedObject) {
        var sortedKeys = [];
        for (var numericKey in numericKeyedObject) {
            if (numericKeyedObject.hasOwnProperty(numericKey)) {
                sortedKeys.push(parseInt(numericKey, 10));
            }
        }
        sortedKeys.sort(function (a, b) { return a - b; });
        return sortedKeys;
    }

    /**
     * レイヤーを再帰的に走査し、直下の page item の移動量を集める
     * 必要に応じてロック／非表示レイヤー・オブジェクトを除外する。
     * @param {Layers} layerCollection - 走査するレイヤー
     * @param {Object[]} placementItems - 移動量を求めたアートボード情報
     * @param {Object[]} itemTranslations - { item, deltaX, deltaY } の格納先
     * @param {boolean} excludeLockedHiddenLayers - ロック／非表示レイヤーを除外するか
     * @param {boolean} excludeLockedHiddenItems - ロック／非表示オブジェクトを除外するか
     * @returns {void}
     */
    function appendItemTranslations(layerCollection, placementItems, itemTranslations, excludeLockedHiddenLayers, excludeLockedHiddenItems) {
        for (var layerIndex = 0; layerIndex < layerCollection.length; layerIndex++) {
            var layer = layerCollection[layerIndex];
            if (excludeLockedHiddenLayers && shouldSkipLayerForItemMove(layer)) continue;

            for (var itemIndex = 0; itemIndex < layer.pageItems.length; itemIndex++) {
                var pageItem = layer.pageItems[itemIndex];
                if (pageItem.parent !== layer) continue;
                if (excludeLockedHiddenItems && shouldSkipPageItemForMove(pageItem)) continue;
                appendTranslationForItem(pageItem, placementItems, itemTranslations);
            }

            if (layer.layers && layer.layers.length > 0) {
                appendItemTranslations(layer.layers, placementItems, itemTranslations, excludeLockedHiddenLayers, excludeLockedHiddenItems);
            }
        }
    }

    /**
     * 移動対象から除外するレイヤーか判定する
     * @param {Layer} layer - 対象のレイヤー
     * @returns {boolean} ロックまたは非表示なら true
     */
    function shouldSkipLayerForItemMove(layer) {
        try {
            return layer.locked || !layer.visible;
        } catch (layerStateError) {
            return false;
        }
    }

    /**
     * 移動対象から除外するオブジェクトか判定する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {boolean} ロックまたは非表示なら true
     */
    function shouldSkipPageItemForMove(pageItem) {
        try {
            return pageItem.locked || pageItem.hidden;
        } catch (pageItemStateError) {
            return false;
        }
    }

    /**
     * 単一オブジェクトの移動量を追加する（中心を含む最初の元アートボードの移動量）
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @param {Object[]} placementItems - 移動量を求めたアートボード情報
     * @param {Object[]} itemTranslations - { item, deltaX, deltaY } の格納先
     * @returns {void}
     */
    function appendTranslationForItem(pageItem, placementItems, itemTranslations) {
        var center = getItemGeometricCenter(pageItem);
        if (!center) return;

        for (var placementIndex = 0; placementIndex < placementItems.length; placementIndex++) {
            var placement = placementItems[placementIndex];
            if (isCenterInsideRect(center, placement.oldRect)) {
                itemTranslations.push({ item: pageItem, deltaX: placement.deltaX, deltaY: placement.deltaY });
                break;
            }
        }
    }

    /**
     * オブジェクトの幾何中心を取得する
     * @param {PageItem} pageItem - 対象のオブジェクト
     * @returns {number[]|null} [x, y]。取得できない場合は null
     */
    function getItemGeometricCenter(pageItem) {
        try {
            var bounds = pageItem.geometricBounds; // [left, top, right, bottom]
            return [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (geometricBoundsError) {
            return null;
        }
    }

    /**
     * 中心座標が矩形内にあるか（artboardRect の Y は上がプラス、下がマイナス）
     * @param {number[]} center - [x, y]
     * @param {number[]} rect - [left, top, right, bottom]
     * @returns {boolean} 矩形内なら true
     */
    function isCenterInsideRect(center, rect) {
        return center[0] >= rect[0] && center[0] <= rect[2] &&
            center[1] <= rect[1] && center[1] >= rect[3];
    }

    /**
     * アートボードパネル上の順序を、割り当てた (assignedRow, assignedColumn) の昇順に再構築する
     * 一時名にリネームしてから新規作成→旧アートボードを削除することで名称重複を回避する。
     * @param {Document} documentRef - 対象のドキュメント
     * @param {Object[]} placementItems - 行列を割り当てたアートボード情報
     * @returns {void}
     */
    function reorderArtboardsByGridOrder(documentRef, placementItems) {
        var sortedPlacements = placementItems.slice();
        sortedPlacements.sort(function (placementA, placementB) {
            if (placementA.assignedRow !== placementB.assignedRow) return placementA.assignedRow - placementB.assignedRow;
            return placementA.assignedColumn - placementB.assignedColumn;
        });

        var artboards = documentRef.artboards;
        var originalCount = artboards.length;

        var orderedArtboardSpecs = [];
        for (var collectIndex = 0; collectIndex < sortedPlacements.length; collectIndex++) {
            var sourceArtboard = sortedPlacements[collectIndex].artboard;
            var sourceRect = sourceArtboard.artboardRect;
            orderedArtboardSpecs.push({
                name: sourceArtboard.name,
                rect: [sourceRect[0], sourceRect[1], sourceRect[2], sourceRect[3]]
            });
        }

        for (var renameIndex = 0; renameIndex < originalCount; renameIndex++) {
            artboards[renameIndex].name = '__reorder_tmp__' + renameIndex;
        }

        for (var addIndex = 0; addIndex < orderedArtboardSpecs.length; addIndex++) {
            var newArtboard = artboards.add(orderedArtboardSpecs[addIndex].rect);
            newArtboard.name = orderedArtboardSpecs[addIndex].name;
        }

        for (var removeIndex = originalCount - 1; removeIndex >= 0; removeIndex--) {
            artboards[removeIndex].remove();
        }

        documentRef.artboards.setActiveArtboardIndex(0);
    }

    /**
     * 各番号（列または行）の累積オフセットを計算する
     * boundaryNumber と一致するキーの直前のギャップは3倍に広げる。
     * @param {number[]} sortedNumbers - 昇順の番号
     * @param {Object} sizeByNumber - 番号 → 幅または高さ
     * @param {number} gap - 間隔（pt）
     * @param {number} boundaryNumber - 間隔を3倍にする番号
     * @returns {Object} 番号 → オフセット
     */
    function computeCumulativeOffsets(sortedNumbers, sizeByNumber, gap, boundaryNumber) {
        var offsetByNumber = {};
        var cumulative = 0;
        for (var orderIndex = 0; orderIndex < sortedNumbers.length; orderIndex++) {
            var currentNumber = sortedNumbers[orderIndex];
            if (orderIndex > 0) {
                var currentGap = gap;
                if (currentNumber === boundaryNumber) currentGap *= 3;
                cumulative += currentGap;
            }
            offsetByNumber[currentNumber] = cumulative;
            cumulative += sizeByNumber[currentNumber];
        }
        return offsetByNumber;
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================

    arrangeArtboardsByRowColumnName();

})();
