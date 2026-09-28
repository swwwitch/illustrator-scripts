#target illustrator
#targetengine "ShuffleObjectColorsEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクト（パス、グループ、複合シェイプ）の塗り・線カラーを再配色します。
RGB／CMYK／グレースケール／特色／グラデーション／パターンに対応します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShuffleObjectColors.md

### Overview

Reshuffles the fill and stroke colors of the selected paths, groups and compound shapes.
RGB, CMYK, grayscale, spot colors, gradients and patterns are all supported.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShuffleObjectColors.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ShuffleObjectColors";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-06-24";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ShuffleObjectColors.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ShuffleObjectColors.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // バージョンとローカライズ
    // =========================================

    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }

    var uiLang = getCurrentLang();

    /* ラベル定義 / Labels */
    var LABELS = {
        /* --- ダイアログ / Dialog --- */
        dialogTitle: { ja: "選択オブジェクトのカラーをシャッフル", en: "Shuffle Selected Object Colors" },

        /* --- パネル / Panels --- */
        panelTarget: { ja: "対象", en: "Target" },
        panelExclude: { ja: "除外", en: "Exclude" },

        /* --- 対象オプション / Target options --- */
        fill: { ja: "塗り", en: "Fill" },
        stroke: { ja: "線", en: "Stroke" },

        /* --- 除外オプション / Exclude options --- */
        black: { ja: "黒", en: "Black" },
        white: { ja: "白", en: "White" },

        /* --- 適用オプション / Apply options --- */
        random: { ja: "ランダム順", en: "Random order" },
        balance: { ja: "バランスを保持", en: "Preserve balance" },

        /* --- ボタン / Buttons --- */
        apply: { ja: "適用", en: "Apply" },
        cancel: { ja: "キャンセル", en: "Cancel" },

        /* --- ツールチップ / Tooltips --- */
        tipFill: { ja: "塗りカラーをシャッフルの対象にする", en: "Include fill colors in the shuffle" },
        tipStroke: { ja: "線カラーをシャッフルの対象にする", en: "Include stroke colors in the shuffle" },
        tipBlack: { ja: "黒（および近いカラー）は変更しない", en: "Leave black (and near-black) colors untouched" },
        tipWhite: { ja: "白は変更しない", en: "Leave white untouched" },
        tipRandom: { ja: "ON：ランダム順／OFF：元の順序で適用", en: "ON: random order / OFF: keep original order" },
        tipBalance: { ja: "同じカラーの出現回数を保持して配色", en: "Keep the original frequency of each color" },
        tipApply: { ja: "ダイアログを閉じずにプレビュー", en: "Preview without closing the dialog" },

        /* --- アラート / Alerts --- */
        alertSelect: { ja: "オブジェクトを選択してください。", en: "Please select some objects." },
        alertEmpty: { ja: "使用可能なカラーが見つかりません（白と黒以外）。", en: "No usable colors found (excluding white and black)." },
        alertChoice: { ja: "塗りまたは線のいずれかを選択してください。", en: "Please select either fill or stroke." }
    };

    function getLabel(key) {
        return LABELS[key][uiLang];
    }

    // =========================================
    // UI 状態 / UI state
    // =========================================

    var fillCheckbox, strokeCheckbox;
    var blackCheckbox, whiteCheckbox;
    var randomCheckbox, balanceCheckbox;
    var previewApplied = false;

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

    main();

    // =========================================
    // メイン処理 / Main
    // =========================================

    function main() {
        if (app.documents.length === 0 || !app.activeDocument.selection || app.activeDocument.selection.length === 0) {
            alert(getLabel("alertSelect"));
            return;
        }
        showDialog();
    }

    function showDialog() {
        var dialog = new Window("dialog", getLabel("dialogTitle") + " " + SCRIPT_VERSION);
        dialog.orientation = "row";
        dialog.alignChildren = "top";

        var leftGroup = dialog.add("group");
        leftGroup.orientation = "column";
        leftGroup.alignChildren = "left";

        buildTargetPanel(leftGroup);
        buildExcludePanel(leftGroup);

        randomCheckbox = leftGroup.add("checkbox", undefined, getLabel("random"));
        randomCheckbox.value = true;
        randomCheckbox.helpTip = getLabel("tipRandom");

        balanceCheckbox = leftGroup.add("checkbox", undefined, getLabel("balance"));
        balanceCheckbox.value = true;
        balanceCheckbox.helpTip = getLabel("tipBalance");

        buildButtons(dialog);

        prepareDialogWindow(dialog, SCRIPT_NAME);
        dialog.show();
    }

    function buildTargetPanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panelTarget"));
        panel.preferredSize.width = 120;
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = [10, 25, 10, 10];

        var row = panel.add("group");
        row.orientation = "row";
        row.alignChildren = "left";

        fillCheckbox = row.add("checkbox", undefined, getLabel("fill"));
        fillCheckbox.value = true;
        fillCheckbox.helpTip = getLabel("tipFill");

        strokeCheckbox = row.add("checkbox", undefined, getLabel("stroke"));
        strokeCheckbox.value = true;
        strokeCheckbox.helpTip = getLabel("tipStroke");
    }

    function buildExcludePanel(parent) {
        var panel = parent.add("panel", undefined, getLabel("panelExclude"));
        panel.preferredSize.width = 120;
        panel.orientation = "column";
        panel.alignChildren = "left";
        panel.margins = [10, 25, 10, 10];

        var row = panel.add("group");
        row.orientation = "row";
        row.alignChildren = "left";

        blackCheckbox = row.add("checkbox", undefined, getLabel("black"));
        blackCheckbox.value = true;
        blackCheckbox.helpTip = getLabel("tipBlack");

        whiteCheckbox = row.add("checkbox", undefined, getLabel("white"));
        whiteCheckbox.value = true;
        whiteCheckbox.helpTip = getLabel("tipWhite");
    }

    function buildButtons(dialog) {
        var rightGroup = dialog.add("group");
        rightGroup.orientation = "column";
        rightGroup.alignChildren = "right";

        /* Mac 規約: 縦並びは OK が上 / Vertical: OK on top */
        var okBtn = rightGroup.add("button", undefined, "OK");
        okBtn.preferredSize.width = 80;

        var cancelBtn = rightGroup.add("button", undefined, getLabel("cancel"));
        cancelBtn.preferredSize.width = 80;

        // スペーサー（縦に伸びる）
        var verticalSpacer = rightGroup.add("statictext", undefined, "");
        verticalSpacer.alignment = ["fill", "fill"];

        // 適用ボタン用グループ（右下）
        var applyGroup = rightGroup.add("group");
        applyGroup.orientation = "column";
        applyGroup.alignChildren = ["fill", "bottom"];

        var applyBtn = applyGroup.add("button", undefined, getLabel("apply"));
        applyBtn.preferredSize.width = 80;
        applyBtn.helpTip = getLabel("tipApply");

        applyBtn.onClick = function () {
            applyBtn.enabled = false;
            applyWithPreview();
            applyBtn.enabled = true;
        };

        okBtn.onClick = function () {
            applyWithPreview();
            dialog.close(1);
        };

        cancelBtn.onClick = function () {
            if (previewApplied) {
                app.undo();
                previewApplied = false;
            }
            dialog.close(0);
        };
    }

    function applyWithPreview() {
        if (previewApplied) {
            app.undo();
            previewApplied = false;
        }
        if (applyColors()) {
            previewApplied = true;
            app.redraw();
        }
    }

    // =========================================
    // カラー収集 / Color collection
    // =========================================

    function applyColors() {
        var changeFill = fillCheckbox.value;
        var changeStroke = strokeCheckbox.value;
        var excludeBlack = blackCheckbox.value;
        var excludeWhite = whiteCheckbox.value;
        var balancePreserve = balanceCheckbox.value;
        var randomize = randomCheckbox.value;

        if (!changeFill && !changeStroke) {
            alert(getLabel("alertChoice"));
            return false;
        }

        var pathItems = [];
        collectPathItems(app.activeDocument.selection, pathItems);

        var colorPool = buildColorPool(pathItems, changeFill, changeStroke, excludeBlack, excludeWhite, balancePreserve);
        if (colorPool.length === 0) {
            alert(getLabel("alertEmpty"));
            return false;
        }

        var shuffled = randomize ? shuffleArray(colorPool) : colorPool;
        reapplyColors(pathItems, shuffled, changeFill, changeStroke, excludeBlack, excludeWhite);
        return true;
    }

    function buildColorPool(pathItems, changeFill, changeStroke, excludeBlack, excludeWhite, balancePreserve) {
        var colorCountMap = {};
        var uniqueColors = {};

        for (var i = 0; i < pathItems.length; i++) {
            var pathItem = pathItems[i];
            if (changeFill && pathItem.filled) {
                addColorToPool(pathItem.fillColor, colorCountMap, uniqueColors, excludeBlack, excludeWhite, balancePreserve);
            }
            if (changeStroke && pathItem.stroked && pathItem.strokeWidth > 0) {
                addColorToPool(pathItem.strokeColor, colorCountMap, uniqueColors, excludeBlack, excludeWhite, balancePreserve);
            }
        }

        var pool = [];
        if (balancePreserve) {
            for (var countKey in colorCountMap) {
                var entry = colorCountMap[countKey];
                for (var r = 0; r < entry.count; r++) {
                    pool.push(cloneColor(entry.color));
                }
            }
        } else {
            for (var uniqueKey in uniqueColors) {
                pool.push(cloneColor(uniqueColors[uniqueKey]));
            }
        }
        return pool;
    }

    function addColorToPool(rawColor, colorCountMap, uniqueColors, excludeBlack, excludeWhite, balancePreserve) {
        var color = cloneColor(rawColor);
        if (!color) return;
        if (excludeBlack && isNearBlack(color)) return;
        if (excludeWhite && isPureWhite(color)) return;

        var colorKey = getColorKey(color);
        if (!colorKey) return;

        if (!uniqueColors[colorKey]) {
            uniqueColors[colorKey] = cloneColor(color);
        }
        if (balancePreserve) {
            if (!colorCountMap[colorKey]) {
                colorCountMap[colorKey] = { color: cloneColor(color), count: 0 };
            }
            colorCountMap[colorKey].count++;
        }
    }

    // 対象パス収集（再帰） / Collect path items recursively
    function collectPathItems(items, results) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === "GroupItem") {
                collectPathItems(item.pageItems, results);
            } else if (item.typename === "CompoundPathItem") {
                collectPathItems(item.pathItems, results);
            } else if (item.typename === "PathItem") {
                results.push(item);
            }
        }
    }

    // =========================================
    // カラー適用 / Color application
    // =========================================

    function reapplyColors(pathItems, colorPool, changeFill, changeStroke, excludeBlack, excludeWhite) {
        var colorIndex = 0;
        for (var i = 0; i < pathItems.length; i++) {
            var pathItem = pathItems[i];
            colorIndex = applyToSlot(pathItem, true, changeFill, colorPool, colorIndex, excludeBlack, excludeWhite);
            colorIndex = applyToSlot(pathItem, false, changeStroke, colorPool, colorIndex, excludeBlack, excludeWhite);
        }
    }

    function applyToSlot(pathItem, isFill, changeFlag, colorPool, colorIndex, excludeBlack, excludeWhite) {
        if (!changeFlag) return colorIndex;

        var currentColor = cloneColor(isFill ? pathItem.fillColor : pathItem.strokeColor);
        if (!currentColor) return colorIndex;
        if (excludeBlack && isNearBlack(currentColor)) return colorIndex;
        if (excludeWhite && isPureWhite(currentColor)) return colorIndex;

        var nextColor = cloneColor(colorPool[colorIndex % colorPool.length]);
        if (isFill) {
            pathItem.filled = true;
            pathItem.fillColor = nextColor;
        } else {
            pathItem.stroked = true;
            pathItem.strokeColor = nextColor;
        }
        return colorIndex + 1;
    }

    // =========================================
    // ユーティリティ / Utilities
    // =========================================

    // Fisher–Yates シャッフル / Fisher–Yates shuffle
    function shuffleArray(arr) {
        var shuffled = arr.slice();
        for (var i = shuffled.length - 1; i > 0; i--) {
            var j = Math.floor(Math.random() * (i + 1));
            var temp = shuffled[i];
            shuffled[i] = shuffled[j];
            shuffled[j] = temp;
        }
        return shuffled;
    }

    /**
     * 色オブジェクトのプロパティを写す
     * 種類によっては持っていないプロパティがあり、代入で例外になるため1つずつ受け流す。
     * @param {object} targetColor - 写し先の色
     * @param {object} sourceColor - 写し元の色
     * @param {string[]} propertyNames - 写すプロパティ名
     * @returns {void}
     */
    function copyColorProperties(targetColor, sourceColor, propertyNames) {
        for (var i = 0; i < propertyNames.length; i++) {
            try {
                targetColor[propertyNames[i]] = sourceColor[propertyNames[i]];
            } catch (e) {}
        }
    }

    function getColorKey(color) {
        if (color.typename === "RGBColor") {
            return "rgb:" + color.red + "," + color.green + "," + color.blue;
        }
        if (color.typename === "CMYKColor") {
            return "cmyk:" + color.cyan + "," + color.magenta + "," + color.yellow + "," + color.black;
        }
        if (color.typename === "GrayColor") {
            return "gray:" + color.gray;
        }
        if (color.typename === "SpotColor") {
            return "spot:" + color.spot.name + "@" + color.tint;
        }
        if (color.typename === "GradientColor") {
            return "grad:" + color.gradient.name;
        }
        if (color.typename === "PatternColor") {
            return "pat:" + color.pattern.name;
        }
        return null;
    }

    // カラー複製（RGB/CMYK/Gray/Spot/Gradient/Pattern 対応） / Clone color (RGB/CMYK/Gray/Spot/Gradient/Pattern)
    function cloneColor(color) {
        if (!color || !color.typename) return null;
        if (color.typename === "RGBColor") {
            var rgb = new RGBColor();
            rgb.red = color.red;
            rgb.green = color.green;
            rgb.blue = color.blue;
            return rgb;
        }
        if (color.typename === "CMYKColor") {
            var cmyk = new CMYKColor();
            cmyk.cyan = color.cyan;
            cmyk.magenta = color.magenta;
            cmyk.yellow = color.yellow;
            cmyk.black = color.black;
            return cmyk;
        }
        if (color.typename === "GrayColor") {
            var gr = new GrayColor();
            gr.gray = color.gray;
            return gr;
        }
        if (color.typename === "SpotColor") {
            /* 見当合わせ色（レジストレーション）はシャッフル対象外 / Skip registration */
            if (color.spot.colorType === ColorModel.REGISTRATION) return null;
            var sc = new SpotColor();
            sc.spot = color.spot;
            sc.tint = color.tint;
            return sc;
        }
        if (color.typename === "GradientColor") {
            var gc = new GradientColor();
            gc.gradient = color.gradient;
            copyColorProperties(gc, color, ["angle", "length", "origin", "hiliteAngle", "hiliteLength", "matrix"]);
            return gc;
        }
        if (color.typename === "PatternColor") {
            var pc = new PatternColor();
            pc.pattern = color.pattern;
            copyColorProperties(pc, color, ["matrix", "shiftAngle", "shiftDistance", "reflect",
                "reflectAngle", "rotation", "scaleFactor", "shearAngle", "shearAxis"]);
            return pc;
        }
        return null;
    }

    // 黒近似判定 / Near-black detection
    function isNearBlack(color) {
        if (color.typename === "RGBColor") {
            return color.red <= 51 && color.green <= 51 && color.blue <= 51;
        }
        if (color.typename === "CMYKColor") {
            return color.black >= 0.8 || (color.cyan <= 0.2 && color.magenta <= 0.2 && color.yellow <= 0.2 && color.black >= 0.7);
        }
        if (color.typename === "GrayColor") {
            return color.gray >= 80;
        }
        if (color.typename === "SpotColor") {
            /* tint 100% のときだけ基底色で判定 / Only when tint is 100% */
            return color.tint >= 99.999 && isNearBlack(color.spot.color);
        }
        return false;
    }

    // 純白判定 / Pure-white detection
    function isPureWhite(color) {
        if (color.typename === "RGBColor") {
            return color.red === 255 && color.green === 255 && color.blue === 255;
        }
        if (color.typename === "CMYKColor") {
            return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
        }
        if (color.typename === "GrayColor") {
            return color.gray === 0;
        }
        if (color.typename === "SpotColor") {
            /* tint 0% は紙色（白） / Tint 0% = paper (white) */
            if (color.tint <= 0.001) return true;
            return color.tint >= 99.999 && isPureWhite(color.spot.color);
        }
        return false;
    }

})();
