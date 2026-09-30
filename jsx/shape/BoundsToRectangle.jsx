#target illustrator
#targetengine "session"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した複数オブジェクト全体の外接矩形をもとに、ひとつの長方形へ統合します。
キーオブジェクトがあれば、その塗りと線を引き継ぎます。
塗りと線の引き継ぎ元やプレビュー境界／オブジェクト境界の切り替え、元の図形を残すオプションを指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BoundsToRectangle.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nd4afdd8315f0

### Overview

Merges the selected objects into a single rectangle based on their overall bounding box.
A key object, if set, supplies the fill and stroke.
You can choose which object's fill and stroke to inherit, switch between preview and geometric bounds, and keep the originals.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/BoundsToRectangle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "BoundsToRectangle";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/BoundsToRectangle.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/BoundsToRectangle.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd4afdd8315f0"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* ［プレビュー境界を使用］の初期値 / Initial state of "Use Preview Bounds" */
    var DEFAULT_USE_VISIBLE_BOUNDS = true;
    /* ［元の図形を残す］の初期値 / Initial state of "Keep Original Objects" */
    var DEFAULT_KEEP_ORIGINALS = false;
    /* ［属性パネルで中心を表示］の初期値 / Initial state of "Show Center" */
    var DEFAULT_SHOW_CENTER = true;

    /* 引き継ぎ元を選ぶキー（1番目から順に） / Keys that pick the source object, in order */
    var SOURCE_SHORTCUT_KEYS = ["J", "K", "L", "Semicolon", "A", "S", "D", "F", "G"];

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [15, 20, 15, 10];

    /**
     * パネルの揃えと余白を設定する
     * @param {Panel} targetPanel - 対象のパネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.alignment = ["fill", "fill"];
        targetPanel.alignChildren = ["left", "top"];
        targetPanel.margins = PANEL_MARGINS;
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /* 引き継ぎ元を示す枠の線幅 / Stroke width of the source highlight frame */
    var HIGHLIGHT_STROKE_WIDTH = 2;

    // =========================================
    // キーオブジェクトの検出 / Key object detection
    // =========================================

    /* 整列後に「動いていない」とみなす差（pt） / Tolerance for "did not move" after aligning */
    var KEY_DETECT_TOLERANCE_PT = 0.001;

    /**
     * スマートガイドの表示を切り替える（検出の前後で同じ状態に戻す）
     * @returns {void}
     */
    function toggleSmartGuides() {
        /* メニューコマンドが使えない状態がある / The menu command may be unavailable */
        try {
            app.executeMenuCommand("edge");
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] toggleSmartGuides error: " + e);
        }
    }

    /**
     * 控えておいた位置へ戻す（1つ失敗しても残りは戻す）
     * @param {PageItem[]} items - 対象のオブジェクト
     * @param {number[][]} positions - [[left, top], ...]
     * @returns {void}
     */
    function restorePositions(items, positions) {
        for (var i = 0; i < items.length; i++) {
            /* ロック中などで動かせないことがある / Some items may refuse to move (e.g. locked) */
            try {
                items[i].left = positions[i][0];
                items[i].top = positions[i][1];
            } catch (e) {
                $.writeln("[" + SCRIPT_NAME + "] restorePositions error: " + e);
            }
        }
    }

    /**
     * 選択オブジェクトからキーオブジェクトを検出する
     * DOM にキーオブジェクトを示すプロパティは無いため、整列コマンドを実行して
     * 「どの向きに整列しても動かないもの」を実測で特定する。
     * 判定中は app.redraw() を呼ばない（描画すると整列がそのつど取り消し履歴に積まれる）
     * @param {PageItem[]} items - 選択中のオブジェクト
     * @returns {number} キーオブジェクトの番号。判定できないときは -1
     */
    function detectKeyObjectIndex(items) {
        var alignCommands = ["Horizontal Align Left", "Horizontal Align Right", "Vertical Align Top", "Vertical Align Bottom"];
        var stayedPut = [];
        var originPositions = [];
        var i;
        for (i = 0; i < items.length; i++) {
            stayedPut.push(true);
            originPositions.push([items[i].left, items[i].top]);
        }

        try {
            for (var commandIndex = 0; commandIndex < alignCommands.length; commandIndex++) {
                /* 2回目以降だけ元の位置へ戻す / Restore only from the second pass on */
                if (commandIndex > 0) restorePositions(items, originPositions);
                app.executeMenuCommand(alignCommands[commandIndex]);
                for (i = 0; i < items.length; i++) {
                    if (!stayedPut[i]) continue;
                    if (Math.abs(items[i].left - originPositions[i][0]) > KEY_DETECT_TOLERANCE_PT ||
                        Math.abs(items[i].top - originPositions[i][1]) > KEY_DETECT_TOLERANCE_PT) {
                        stayedPut[i] = false;
                    }
                }
            }
        } finally {
            /* 例外で抜けるときも整列結果を残さない / Never leave the aligned positions behind */
            restorePositions(items, originPositions);
        }

        var foundIndex = -1;
        for (i = 0; i < items.length; i++) {
            if (!stayedPut[i]) continue;
            if (foundIndex !== -1) return -1; /* 複数残った＝判定不能 / More than one stayed: undecidable */
            foundIndex = i;
        }
        return foundIndex;
    }

    /**
     * キーオブジェクトの番号を返す（2つ以上選択しているときだけ判定する）
     * @param {PageItem[]} items - 選択中のオブジェクト
     * @returns {number} キーオブジェクトの番号。無いときは -1
     */
    function findKeyObjectIndex(items) {
        if (items.length < 2) return -1;
        toggleSmartGuides();
        /* 検出できなくても処理は続ける / Carry on even if detection fails */
        try {
            return detectKeyObjectIndex(items);
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] detectKeyObjectIndex error: " + e);
            return -1;
        } finally {
            toggleSmartGuides();
        }
    }

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

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
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
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

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

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

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

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

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    // キーボードショートカット（再利用パーツ） / Keyboard shortcuts (reusable)

    /* 入力中はショートカットを止めるコントロールの種類 / Control types that swallow keys while focused */
    var KEY_SHORTCUT_TYPING_TYPES = { edittext: true, dropdownlist: true, listbox: true };

    /* 修飾キーの並び順（キーの表記をそろえる）/ Canonical order of modifiers in a key spec */
    var KEY_SHORTCUT_MODIFIERS = ["SHIFT", "ALT", "CMD"];

    /* 修飾キーの別名 / Aliases accepted for the modifiers */
    var KEY_SHORTCUT_MODIFIER_ALIASES = {
        SHIFT: "SHIFT",
        ALT: "ALT", OPTION: "ALT", OPT: "ALT",
        CMD: "CMD", COMMAND: "CMD", META: "CMD", CTRL: "CMD", CONTROL: "CMD"
    };

    /**
     * キーの指定（"Shift+R" など）を、照合用の表記（"SHIFT+R"）にそろえる
     * @param {string} keySpec - キーの指定。修飾キーは "Shift+" / "Alt+" / "Cmd+" を前に付ける
     * @returns {string} 照合用の表記（大文字、修飾キーは SHIFT → ALT → CMD の順）
     */
    function normalizeKeyShortcutSpec(keySpec) {
        var specParts = String(keySpec).split("+");
        var baseKey = specParts.pop().toUpperCase();
        var modifierFlags = {};
        for (var i = 0; i < specParts.length; i++) {
            var modifierName = KEY_SHORTCUT_MODIFIER_ALIASES[specParts[i].toUpperCase()];
            if (modifierName) modifierFlags[modifierName] = true;
        }
        return buildKeyShortcutSpec(modifierFlags, baseKey);
    }

    /**
     * 修飾キーの状態とキー名から照合用の表記を組み立てる
     * @param {Object} modifierFlags - { SHIFT: true, ALT: true, CMD: true } のうち押されているもの
     * @param {string} baseKey - 大文字のキー名
     * @returns {string} 照合用の表記
     */
    function buildKeyShortcutSpec(modifierFlags, baseKey) {
        var specText = "";
        for (var i = 0; i < KEY_SHORTCUT_MODIFIERS.length; i++) {
            if (modifierFlags[KEY_SHORTCUT_MODIFIERS[i]]) specText += KEY_SHORTCUT_MODIFIERS[i] + "+";
        }
        return specText + baseKey;
    }

    /**
     * keydown イベントから照合用の表記を作る。修飾キーはイベントと keyboardState の両方を見る
     * @param {Object} keyEvent - keydown イベント
     * @returns {string} 照合用の表記。キー名が無いときは空文字
     */
    function readKeyShortcutSpec(keyEvent) {
        if (!keyEvent || !keyEvent.keyName) return "";
        var keyboardState = {};
        try { keyboardState = ScriptUI.environment.keyboardState; } catch (e) { }
        var modifierFlags = {
            SHIFT: !!(keyEvent.shiftKey || keyboardState.shiftKey),
            ALT: !!(keyEvent.altKey || keyboardState.altKey),
            CMD: !!(keyEvent.metaKey || keyEvent.ctrlKey || keyboardState.metaKey || keyboardState.ctrlKey)
        };
        return buildKeyShortcutSpec(modifierFlags, String(keyEvent.keyName).toUpperCase());
    }

    /**
     * コントロールが押せる状態か（自分と親がすべて有効で表示中か）を返す
     * @param {Object} control - コントロール
     * @returns {boolean} 押せるなら true
     */
    function isKeyShortcutControlUsable(control) {
        for (var node = control; node; node = node.parent) {
            if (node.enabled === false || node.visible === false) return false;
        }
        return true;
    }

    /**
     * キーを受けたコントロールが、文字を入力する欄か
     * @param {Object} focusedControl - イベントの発生元
     * @param {Object[]} numericFields - 数値だけの欄（ショートカットを効かせる）
     * @returns {boolean} 入力中としてショートカットを止めるなら true
     */
    function isKeyShortcutTypingTarget(focusedControl, numericFields) {
        if (!focusedControl || !KEY_SHORTCUT_TYPING_TYPES[focusedControl.type]) return false;
        for (var i = 0; i < numericFields.length; i++) {
            if (numericFields[i] === focusedControl) return false;
        }
        return true;
    }

    /**
     * コントロールをクリックしたときと同じ動作をする
     * ラジオは同じ親のラジオを外して選び、チェックボックスは反転してから onClick を呼ぶ
     * @param {Object} control - ラジオボタン・チェックボックス・ボタンなど
     * @returns {void}
     */
    function pressKeyShortcutControl(control) {
        if (control.type === "radiobutton") {
            /* 同じ親の直下だけが排他になるので、クリックと同じく兄弟を外す / Clear siblings like a click would */
            var siblings = control.parent ? control.parent.children : [];
            for (var i = 0; i < siblings.length; i++) {
                if (siblings[i] !== control && siblings[i].type === "radiobutton") siblings[i].value = false;
            }
            control.value = true;
        } else if (control.type === "checkbox") {
            control.value = !control.value;
        }
        if (typeof control.onClick === "function") {
            control.onClick.call(control);
        } else if (control.type === "button" && typeof control.notify === "function") {
            /* onClick の無い OK・キャンセルは notify で既定の動作（閉じる）を起こす / Let default buttons close the dialog */
            control.notify("onClick");
        }
    }

    /**
     * 1つのショートカットを実行する
     * @param {Object|Function} shortcutTarget - コントロール、または関数
     * @param {Object} keyEvent - keydown イベント
     * @returns {boolean} キーを使ったなら true（false なら文字をそのまま通す）
     */
    function runKeyShortcutTarget(shortcutTarget, keyEvent) {
        var targetControl = shortcutTarget;
        if (typeof shortcutTarget === "function") {
            var runResult = shortcutTarget(keyEvent);
            if (runResult === false || runResult === null) return false;
            if (!runResult || typeof runResult !== "object" || !runResult.type) return true;
            targetControl = runResult;
        }
        /* 無効なコントロールのキーも使ったことにして、数値欄へ文字を入れない / Consume the key even when disabled */
        if (isKeyShortcutControlUsable(targetControl)) pressKeyShortcutControl(targetControl);
        return true;
    }

    /**
     * キーの指定に修飾キーの表示名を当てて、ツールチップ用の表記にする
     * @param {string} normalizedSpec - 照合用の表記（"SHIFT+R" など）
     * @returns {string} 表示用の表記（"Shift+R" など）
     */
    function formatKeyShortcutLabel(normalizedSpec) {
        var isMac = ($.os.indexOf("Mac") === 0);
        var displayNames = { SHIFT: "Shift", ALT: isMac ? "Option" : "Alt", CMD: isMac ? "Cmd" : "Ctrl" };
        var specParts = normalizedSpec.split("+");
        var baseKey = specParts.pop();
        var labelText = "";
        for (var i = 0; i < specParts.length; i++) labelText += displayNames[specParts[i]] + "+";
        if (baseKey.length > 1) baseKey = baseKey.charAt(0) + baseKey.substring(1).toLowerCase();
        return labelText + baseKey;
    }

    /**
     * コントロールのツールチップの末尾にキーを足す（すでに書いてあれば足さない）
     * @param {Object} control - コントロール
     * @param {string} normalizedSpec - 照合用の表記
     * @returns {void}
     */
    function appendKeyShortcutToTip(control, normalizedSpec) {
        var keyLabel = formatKeyShortcutLabel(normalizedSpec);
        var currentTip = control.helpTip ? String(control.helpTip) : "";
        if (currentTip.indexOf("（" + keyLabel) >= 0 || currentTip.indexOf("(" + keyLabel) >= 0) return;
        var keySuffix = (uiLang === "ja") ? "（" + keyLabel + "）" : " (" + keyLabel + ")";
        control.helpTip = currentTip ? currentTip + keySuffix : keyLabel;
    }

    /**
     * ダイアログ・パレットに文字キーのショートカットを付ける
     * @param {Window} targetWindow - キーを受けるダイアログ・パレット
     * @param {Object} shortcutMap - { "L": ラジオ, "Shift+R": ボタン, "G": 関数, "Escape": { target: 関数, inFields: true } }
     * @param {Object} [shortcutOptions] - numericFields（数値だけの欄の配列）/ afterKey（キーを使ったあとに呼ぶ関数）/ showInTip（ツールチップにキーを足す）
     * @returns {Object} 照合用の表記 → { target, inFields } の表（テスト・デバッグ用）
     */
    function addKeyShortcuts(targetWindow, shortcutMap, shortcutOptions) {
        var shortcutSettings = shortcutOptions || {};
        var numericFields = shortcutSettings.numericFields || [];
        var bindingTable = {};

        for (var keySpec in shortcutMap) {
            if (!shortcutMap.hasOwnProperty(keySpec)) continue;
            var mapEntry = shortcutMap[keySpec];
            if (!mapEntry) continue;
            var isWrapped = (typeof mapEntry === "object" && !mapEntry.type && mapEntry.target);
            var normalizedSpec = normalizeKeyShortcutSpec(keySpec);
            bindingTable[normalizedSpec] = {
                target: isWrapped ? mapEntry.target : mapEntry,
                inFields: !!(isWrapped && mapEntry.inFields)
            };
            var tipControl = bindingTable[normalizedSpec].target;
            if (shortcutSettings.showInTip && typeof tipControl === "object" && tipControl.type) {
                appendKeyShortcutToTip(tipControl, normalizedSpec);
            }
        }

        /* キャプチャで受けて、数値欄に文字が入る前に止める / Capture phase keeps the letter out of numeric fields */
        targetWindow.addEventListener("keydown", function (keyEvent) {
            var binding = bindingTable[readKeyShortcutSpec(keyEvent)];
            if (!binding) return;
            if (!binding.inFields && isKeyShortcutTypingTarget(keyEvent.target, numericFields)) return;
            if (!runKeyShortcutTarget(binding.target, keyEvent)) return;
            if (keyEvent.preventDefault) keyEvent.preventDefault();
            if (typeof shortcutSettings.afterKey === "function") shortcutSettings.afterKey(keyEvent);
        }, true);

        return bindingTable;
    }

    // キーボードショートカット（再利用パーツ）ここまで / End of the reusable keyboard shortcuts

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "外接矩形を長方形に", en: "Merge Bounds into One Rectangle" }
        },
        panel: {
            inheritSource: { ja: "塗りと線の引き継ぎ元", en: "Take Fill and Stroke From" },
            options: { ja: "オプション", en: "Options" }
        },
        radio: {
            keyObjectSuffix: { ja: "（キーオブジェクト）", en: " (Key Object)" }
        },
        checkbox: {
            useVisibleBounds: { ja: "プレビュー境界を使用", en: "Use Preview Bounds" },
            keepOriginals: { ja: "元の図形を残す", en: "Keep Original Objects" },
            showCenter: { ja: "属性パネルで中心を表示", en: "Show Center in Attributes Panel" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        tooltip: {
            shortcutKey: { ja: "ショートカット", en: "Shortcut" },
            useVisibleBounds: {
                ja: "線幅や効果を含む見た目の範囲で計算。OFFのときはパスの形状で計算",
                en: "Measure including stroke width and effects. When off, use the path geometry"
            },
            keepOriginals: { ja: "元のオブジェクトを残して新しい長方形を作成", en: "Keep the originals and create a new rectangle" },
            showCenter: {
                ja: "OK後、作成した長方形に属性パネルの［中心を表示］を適用",
                en: "After OK, turn on Show Center in the Attributes panel for the new rectangle"
            },
            preview: {
                ja: "作成される長方形を表示し、元のオブジェクトを一時的に隠す",
                en: "Show the resulting rectangle and temporarily hide the originals"
            }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection: { ja: "オブジェクトを選択してから実行してください。", en: "Please select objects before running this script." }
        },
        /* 名前のないオブジェクトの表示名 / Display names for unnamed objects */
        fallbackName: {
            PathItem: { ja: "パス", en: "Path" },
            CompoundPathItem: { ja: "複合パス", en: "Compound Path" },
            GroupItem: { ja: "グループ", en: "Group" },
            TextFrame: { ja: "テキスト", en: "Text" },
            PlacedItem: { ja: "配置画像", en: "Placed Image" },
            RasterItem: { ja: "ラスター画像", en: "Raster Image" },
            SymbolItem: { ja: "シンボル", en: "Symbol" },
            MeshItem: { ja: "メッシュ", en: "Mesh" },
            PluginItem: { ja: "プラグインアイテム", en: "Plugin Item" }
        },
        layerName: {
            previewTemp: { ja: "_選択プレビュー（一時）", en: "_Selection Preview (Temp)" }
        }
    };

    // =========================================
    // 属性パネルの［中心を表示］アクション / "Show Center" action
    // =========================================

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    /**
     * 選択中のオブジェクトに属性パネルの［中心を表示］を適用する
     * @returns {void}
     */
    function runShowCenterAction() {
        var actionSetName = "Attribute";
        var actionName = "ShowCenter";
        var actionSource = [
            "/version 3",
            "/name [ 9 417474726962757465 ]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            " /name [ 10 53686f7743656e746572 ]",
            " /keyIndex 0",
            " /colorIndex 0",
            " /isOpen 1",
            " /eventCount 1",
            " /event-1 {",
            "  /useRulersIn1stQuadrant 0",
            "  /internalName (adobe_attributePalette)",
            "  /localizedName [ 12 e5b19ee680a7e8a8ade5ae9a ]",
            "  /isOpen 1",
            "  /isOn 1",
            "  /hasDialog 0",
            "  /parameterCount 1",
            "  /parameter-1 {",
            "   /key 1668183154",
            "   /showInPalette 4294967295",
            "   /type (boolean)",
            "   /value 1",
            "  }",
            " }",
            "}"
        ].join("\n");

        /* 失敗は従来どおり例外で伝える / Report a failure as an exception, as before */
        if (!runTemporaryAction(actionSource, actionSetName, actionName)) {
            throw new Error("Could not run the action: " + actionSetName + " / " + actionName);
        }
    }

    // =========================================
    // 外接矩形の計算 / Bounds calculation
    // =========================================

    /**
     * [左, 上, 右, 下] の配列を名前付きの外接矩形にする
     * @param {number[]} boundsArray - visibleBounds / geometricBounds の値
     * @returns {{left: number, top: number, right: number, bottom: number}} 外接矩形
     */
    function toBoundsObject(boundsArray) {
        return { left: boundsArray[0], top: boundsArray[1], right: boundsArray[2], bottom: boundsArray[3] };
    }

    /**
     * 選択の状態を控える（ダイアログ中に選択やレイヤーが変わっても使えるように）
     * @param {Array} selectedItems - doc.selection
     * @param {number} keyObjectIndex - キーオブジェクトの番号（無いときは -1）
     * @returns {{items: PageItem[], visibleBounds: Object[], geometricBounds: Object[], hiddenStates: boolean[], keyObjectIndex: number}} 控えた状態
     */
    function takeSelectionSnapshot(selectedItems, keyObjectIndex) {
        var snapshot = { items: [], visibleBounds: [], geometricBounds: [], hiddenStates: [], keyObjectIndex: keyObjectIndex };
        for (var i = 0; i < selectedItems.length; i++) {
            var item = selectedItems[i];
            snapshot.items.push(item);
            snapshot.hiddenStates.push(item.hidden);
            snapshot.visibleBounds.push(toBoundsObject(item.visibleBounds));
            snapshot.geometricBounds.push(toBoundsObject(item.geometricBounds));
        }
        return snapshot;
    }

    /**
     * 使う境界の一覧を返す
     * @param {Object} snapshot - takeSelectionSnapshot() の戻り値
     * @param {boolean} useVisibleBounds - プレビュー境界を使うか
     * @returns {Object[]} 各オブジェクトの外接矩形
     */
    function getBoundsList(snapshot, useVisibleBounds) {
        return useVisibleBounds ? snapshot.visibleBounds : snapshot.geometricBounds;
    }

    /**
     * 外接矩形の一覧を包む外接矩形を返す
     * @param {Object[]} boundsList - 外接矩形の一覧
     * @returns {{left: number, top: number, right: number, bottom: number}} 全体の外接矩形
     */
    function getMergedBounds(boundsList) {
        var merged = { left: Infinity, top: -Infinity, right: -Infinity, bottom: Infinity };
        for (var i = 0; i < boundsList.length; i++) {
            var bounds = boundsList[i];
            if (bounds.left < merged.left) merged.left = bounds.left;
            if (bounds.top > merged.top) merged.top = bounds.top;
            if (bounds.right > merged.right) merged.right = bounds.right;
            if (bounds.bottom < merged.bottom) merged.bottom = bounds.bottom;
        }
        return merged;
    }

    /**
     * 線の位置（中央・内側・外側）を返す
     * @param {PageItem} sourceItem - 引き継ぎ元
     * @returns {string} "center" / "inside" / "outside"
     */
    function getStrokeAlignmentName(sourceItem) {
        /* strokeAlignment を持たない種類・バージョンがある / Not every item type or version has strokeAlignment */
        try {
            var alignmentText = String(sourceItem.strokeAlignment).toLowerCase();
            if (alignmentText.indexOf("inside") >= 0) return "inside";
            if (alignmentText.indexOf("outside") >= 0) return "outside";
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] getStrokeAlignmentName error: " + e);
        }
        return "center";
    }

    /**
     * 作る長方形のパスの外接矩形を返す。プレビュー境界のときは、見た目が合うよう線幅ぶん内側に寄せる
     * @param {Object} mergedBounds - 全体の外接矩形
     * @param {PageItem} sourceItem - 引き継ぎ元
     * @param {boolean} useVisibleBounds - プレビュー境界を使うか
     * @returns {{left: number, top: number, right: number, bottom: number}} パスの外接矩形
     */
    function getRectanglePathBounds(mergedBounds, sourceItem, useVisibleBounds) {
        var pathBounds = {
            left: mergedBounds.left,
            top: mergedBounds.top,
            right: mergedBounds.right,
            bottom: mergedBounds.bottom
        };
        if (!useVisibleBounds || !sourceItem.stroked) return pathBounds;

        var strokeWidth = Number(sourceItem.strokeWidth);
        if (!(strokeWidth > 0)) return pathBounds;

        var alignmentName = getStrokeAlignmentName(sourceItem);
        var inset = 0;
        if (alignmentName === "center") inset = strokeWidth / 2;
        else if (alignmentName === "outside") inset = strokeWidth;

        pathBounds.left += inset;
        pathBounds.top -= inset;
        pathBounds.right -= inset;
        pathBounds.bottom += inset;

        /* 線幅より小さいときは中央に潰す / Collapse to the center when smaller than the stroke */
        if (pathBounds.right < pathBounds.left) {
            pathBounds.left = pathBounds.right = (pathBounds.left + pathBounds.right) / 2;
        }
        if (pathBounds.top < pathBounds.bottom) {
            pathBounds.top = pathBounds.bottom = (pathBounds.top + pathBounds.bottom) / 2;
        }
        return pathBounds;
    }

    // =========================================
    // 長方形の作成 / Rectangle creation
    // =========================================

    /**
     * 外接矩形どおりの長方形を追加する
     * @param {Layer} targetLayer - 追加先のレイヤー
     * @param {Object} bounds - 外接矩形
     * @returns {PathItem} 追加した長方形
     */
    function addRectangleFromBounds(targetLayer, bounds) {
        return targetLayer.pathItems.rectangle(bounds.top, bounds.left, bounds.right - bounds.left, bounds.top - bounds.bottom);
    }

    /**
     * 水平・垂直の辺でできた4点の長方形パスか
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} 長方形なら true
     */
    function isAxisAlignedRectanglePath(item) {
        if (item.typename !== "PathItem" || !item.closed || item.pathPoints.length !== 4) return false;

        var xValues = {};
        var yValues = {};
        var xCount = 0;
        var yCount = 0;
        for (var i = 0; i < 4; i++) {
            var anchor = item.pathPoints[i].anchor;
            if (!xValues[anchor[0]]) { xValues[anchor[0]] = true; xCount++; }
            if (!yValues[anchor[1]]) { yValues[anchor[1]] = true; yCount++; }
        }
        return (xCount === 2 && yCount === 2);
    }

    /**
     * 長方形パスを外接矩形どおりに描き直す（ハンドルは消す）
     * @param {PathItem} rectanglePath - 描き直す長方形
     * @param {Object} bounds - 外接矩形
     * @returns {void}
     */
    function reshapeRectanglePath(rectanglePath, bounds) {
        rectanglePath.setEntirePath([
            [bounds.left, bounds.top],
            [bounds.right, bounds.top],
            [bounds.right, bounds.bottom],
            [bounds.left, bounds.bottom]
        ]);
        rectanglePath.closed = true;

        for (var i = 0; i < rectanglePath.pathPoints.length; i++) {
            var point = rectanglePath.pathPoints[i];
            point.leftDirection = point.anchor;
            point.rightDirection = point.anchor;
            point.pointType = PointType.CORNER;
        }
    }

    /* 線ありのときに引き継ぐ線の属性 / Stroke properties copied when stroked */
    var STROKE_PROPERTY_NAMES = ["strokeDashes", "strokeDashOffset", "strokeCap", "strokeJoin", "strokeMiterLimit", "strokeOverprint", "strokeAlignment"];
    /* 常に引き継ぐ属性 / Properties always copied */
    var COMMON_PROPERTY_NAMES = ["fillOverprint", "opacity", "blendingMode"];

    /**
     * 属性を1つずつ写す
     * @param {PageItem} sourceItem - 写し元
     * @param {PageItem} targetItem - 写し先
     * @param {string[]} propertyNames - 属性名の一覧
     * @returns {void}
     */
    function copyProperties(sourceItem, targetItem, propertyNames) {
        for (var i = 0; i < propertyNames.length; i++) {
            /* 引き継ぎ元の種類によっては持たない属性がある / Some source types lack these properties */
            try {
                targetItem[propertyNames[i]] = sourceItem[propertyNames[i]];
            } catch (e) {
                $.writeln("[" + SCRIPT_NAME + "] copy " + propertyNames[i] + " error: " + e);
            }
        }
    }

    /**
     * 塗り・線・不透明度などを写す
     * @param {PageItem} sourceItem - 写し元
     * @param {PathItem} targetPath - 写し先
     * @returns {void}
     */
    function copyBasicAppearance(sourceItem, targetPath) {
        targetPath.filled = !!sourceItem.filled;
        if (targetPath.filled) targetPath.fillColor = sourceItem.fillColor;

        targetPath.stroked = !!sourceItem.stroked;
        if (targetPath.stroked) {
            targetPath.strokeColor = sourceItem.strokeColor;
            targetPath.strokeWidth = sourceItem.strokeWidth;
            copyProperties(sourceItem, targetPath, STROKE_PROPERTY_NAMES);
        }
        copyProperties(sourceItem, targetPath, COMMON_PROPERTY_NAMES);
    }

    /**
     * 引き継ぎ元の見た目で長方形を作る
     * @param {Layer} targetLayer - 追加先のレイヤー
     * @param {Object} mergedBounds - 全体の外接矩形
     * @param {PageItem} sourceItem - 引き継ぎ元
     * @param {boolean} useVisibleBounds - プレビュー境界を使うか
     * @returns {PathItem} 作った長方形
     */
    function createMergedRectangle(targetLayer, mergedBounds, sourceItem, useVisibleBounds) {
        var rectanglePath = addRectangleFromBounds(targetLayer, getRectanglePathBounds(mergedBounds, sourceItem, useVisibleBounds));
        copyBasicAppearance(sourceItem, rectanglePath);
        return rectanglePath;
    }

    /**
     * 選択オブジェクトをひとつの長方形にまとめる
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} snapshot - takeSelectionSnapshot() の戻り値
     * @param {{sourceIndex: number, useVisibleBounds: boolean, keepOriginals: boolean}} mergeSettings - ダイアログの設定
     * @returns {void}
     */
    function applyMergeResult(doc, snapshot, mergeSettings) {
        var items = snapshot.items;
        var sourceItem = items[mergeSettings.sourceIndex];
        var mergedBounds = getMergedBounds(getBoundsList(snapshot, mergeSettings.useVisibleBounds));
        var resultPath;

        if (!mergeSettings.keepOriginals && isAxisAlignedRectanglePath(sourceItem)) {
            /* 引き継ぎ元が長方形なら、それ自体を変形して使う（見た目の写しは不要） / Reuse the source rectangle itself */
            reshapeRectanglePath(sourceItem, getRectanglePathBounds(mergedBounds, sourceItem, mergeSettings.useVisibleBounds));
            resultPath = sourceItem;
        } else {
            resultPath = createMergedRectangle(doc.activeLayer, mergedBounds, sourceItem, mergeSettings.useVisibleBounds);
        }

        if (!mergeSettings.keepOriginals) {
            for (var i = items.length - 1; i >= 0; i--) {
                if (items[i] !== resultPath) items[i].remove();
            }
        }
        resultPath.selected = true;
    }

    // =========================================
    // プレビュー表示 / Preview display
    // =========================================

    /**
     * ダイアログ中のプレビューを管理する
     * @param {Document} doc - 対象のドキュメント
     * @param {Object} snapshot - takeSelectionSnapshot() の戻り値
     * @returns {{update: Function, dispose: Function}} 更新と後片付け
     */
    function createPreviewController(doc, snapshot) {
        var originalActiveLayer = doc.activeLayer;
        var previewLayer = null;

        var highlightColor = new RGBColor();
        highlightColor.red = 255;
        highlightColor.green = 0;
        highlightColor.blue = 0;

        /**
         * 元のオブジェクトの表示を戻す
         * @returns {void}
         */
        function restoreHiddenStates() {
            for (var i = 0; i < snapshot.items.length; i++) {
                snapshot.items[i].hidden = snapshot.hiddenStates[i];
            }
        }

        /**
         * プレビュー用レイヤーを空にして返す（無ければ作る）
         * @returns {Layer} プレビュー用レイヤー
         */
        function getEmptyPreviewLayer() {
            if (!previewLayer) {
                previewLayer = doc.layers.add();
                previewLayer.name = getLabel("layerName.previewTemp");
            }
            while (previewLayer.pageItems.length > 0) {
                previewLayer.pageItems[0].remove();
            }
            return previewLayer;
        }

        /**
         * プレビューを描き直す
         * @param {{sourceIndex: number, useVisibleBounds: boolean, showMergedRectangle: boolean}} previewSettings - ダイアログの設定
         * @returns {void}
         */
        function update(previewSettings) {
            var targetLayer = getEmptyPreviewLayer();
            var boundsList = getBoundsList(snapshot, previewSettings.useVisibleBounds);

            if (previewSettings.showMergedRectangle) {
                /* 仕上がりの長方形を描き、元のオブジェクトは一時的に隠す / Draw the result and hide the originals */
                createMergedRectangle(targetLayer, getMergedBounds(boundsList), snapshot.items[previewSettings.sourceIndex], previewSettings.useVisibleBounds);
                for (var i = 0; i < snapshot.items.length; i++) {
                    snapshot.items[i].hidden = true;
                }
            } else {
                /* 引き継ぎ元を赤枠で示す / Outline the source object in red */
                var highlightFrame = addRectangleFromBounds(targetLayer, boundsList[previewSettings.sourceIndex]);
                highlightFrame.filled = false;
                highlightFrame.stroked = true;
                highlightFrame.strokeColor = highlightColor;
                highlightFrame.strokeWidth = HIGHLIGHT_STROKE_WIDTH;
                restoreHiddenStates();
            }
            app.redraw();
        }

        /**
         * プレビュー用レイヤーを消し、表示とアクティブレイヤーを戻す
         * @returns {void}
         */
        function dispose() {
            restoreHiddenStates();
            if (previewLayer) {
                previewLayer.remove();
                previewLayer = null;
            }
            /* ロック中などで戻せないことがある / The original layer may refuse activation (e.g. locked) */
            try {
                doc.activeLayer = originalActiveLayer;
            } catch (e) {
                $.writeln("[" + SCRIPT_NAME + "] restore activeLayer error: " + e);
            }
            app.redraw();
        }

        return { update: update, dispose: dispose };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ラジオボタンに表示するオブジェクト名を返す
     * @param {PageItem} item - 対象のオブジェクト
     * @param {number} index - 選択内の番号（0始まり）
     * @returns {string} "1: パス" のような表示名
     */
    function getItemDisplayName(item, index) {
        var itemName = item.name;
        if (!itemName) {
            var fallbackKey = "fallbackName." + item.typename;
            itemName = getLabel(fallbackKey);
            if (itemName === fallbackKey) itemName = item.typename;
        }
        return (index + 1) + ": " + itemName;
    }

    /**
     * オンになっているラジオボタンの番号を返す
     * @param {RadioButton[]} radioButtons - ラジオボタンの一覧
     * @returns {number} 番号（どれもオフなら 0）
     */
    function getCheckedIndex(radioButtons) {
        for (var i = 0; i < radioButtons.length; i++) {
            if (radioButtons[i].value) return i;
        }
        return 0;
    }

    /**
     * キー名を tooltip 用の表記にする
     * @param {string} keyName - ScriptUI のキー名
     * @returns {string} 表示用の文字
     */
    function getKeyDisplayText(keyName) {
        return (keyName === "Semicolon") ? ";" : keyName;
    }

    /**
     * 塗りと線の引き継ぎ元パネルを作る
     * @param {Window} dialog - 親のダイアログ
     * @param {PageItem[]} items - 選択オブジェクト
     * @param {number} keyObjectIndex - キーオブジェクトの番号（無いときは -1）
     * @returns {RadioButton[]} 作ったラジオボタン
     */
    function buildSourcePanel(dialog, items, keyObjectIndex) {
        var sourcePanel = dialog.add("panel", undefined, getLabel("panel.inheritSource"));
        setupPanel(sourcePanel);

        var sourceRadios = [];
        for (var i = 0; i < items.length; i++) {
            var radioText = getItemDisplayName(items[i], i);
            if (i === keyObjectIndex) radioText += getLabel("radio.keyObjectSuffix");
            var sourceRadio = sourcePanel.add("radiobutton", undefined, radioText);
            if (i < SOURCE_SHORTCUT_KEYS.length) {
                sourceRadio.helpTip = labelText("tooltip.shortcutKey") + getKeyDisplayText(SOURCE_SHORTCUT_KEYS[i]);
            }
            sourceRadios.push(sourceRadio);
        }
        /* キーオブジェクトがあれば、その属性を優先する / Prefer the key object's attributes */
        sourceRadios[(keyObjectIndex >= 0) ? keyObjectIndex : 0].value = true;
        return sourceRadios;
    }

    /**
     * オプションパネルを作る
     * @param {Window} dialog - 親のダイアログ
     * @returns {{useVisibleBounds: Checkbox, keepOriginals: Checkbox, showCenter: Checkbox}} 作ったチェックボックス
     */
    function buildOptionsPanel(dialog) {
        var optionsPanel = dialog.add("panel", undefined, getLabel("panel.options"));
        setupPanel(optionsPanel);

        var useVisibleBoundsCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.useVisibleBounds"));
        useVisibleBoundsCheckbox.value = DEFAULT_USE_VISIBLE_BOUNDS;
        useVisibleBoundsCheckbox.helpTip = getLabel("tooltip.useVisibleBounds");

        var keepOriginalsCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.keepOriginals"));
        keepOriginalsCheckbox.value = DEFAULT_KEEP_ORIGINALS;
        keepOriginalsCheckbox.helpTip = getLabel("tooltip.keepOriginals");

        var showCenterCheckbox = optionsPanel.add("checkbox", undefined, getLabel("checkbox.showCenter"));
        showCenterCheckbox.value = DEFAULT_SHOW_CENTER;
        showCenterCheckbox.helpTip = getLabel("tooltip.showCenter");

        return {
            useVisibleBounds: useVisibleBoundsCheckbox,
            keepOriginals: keepOriginalsCheckbox,
            showCenter: showCenterCheckbox
        };
    }

    /**
     * ダイアログを表示して設定を返す
     * @param {Object} snapshot - takeSelectionSnapshot() の戻り値
     * @param {{update: Function}} previewController - プレビューの管理
     * @returns {Object|null} 設定。キャンセルなら null
     */
    function showMergeDialog(snapshot, previewController) {
        var dialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        dialog.alignChildren = ["left", "top"];

        var sourceRadios = buildSourcePanel(dialog, snapshot.items, snapshot.keyObjectIndex);
        var optionCheckboxes = buildOptionsPanel(dialog);

        var previewGroup = dialog.add("group");
        previewGroup.alignment = ["fill", "top"];
        previewGroup.alignChildren = ["center", "center"];
        var previewCheckbox = previewGroup.add("checkbox", undefined, getLabel("checkbox.preview"));
        previewCheckbox.value = false;
        previewCheckbox.helpTip = getLabel("tooltip.preview");

        /* ボタン行（左右中央） / Button row (centered) */
        var buttonRow = addButtonRow(dialog, { centered: true });
        var btnCancel = buttonRow.rowGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = buttonRow.rowGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        /**
         * プレビューをいまの設定で描き直す
         * @returns {void}
         */
        function refreshPreview() {
            previewController.update({
                sourceIndex: getCheckedIndex(sourceRadios),
                useVisibleBounds: optionCheckboxes.useVisibleBounds.value,
                showMergedRectangle: previewCheckbox.value
            });
        }

        for (var i = 0; i < sourceRadios.length; i++) {
            sourceRadios[i].onClick = refreshPreview;
        }
        optionCheckboxes.useVisibleBounds.onClick = refreshPreview;
        previewCheckbox.onClick = refreshPreview;

        /* キーで引き継ぎ元を選ぶ / Pick the source object by key */
        var sourceShortcutMap = {};
        for (var keyIndex = 0; keyIndex < SOURCE_SHORTCUT_KEYS.length && keyIndex < sourceRadios.length; keyIndex++) {
            sourceShortcutMap[SOURCE_SHORTCUT_KEYS[keyIndex]] = sourceRadios[keyIndex];
        }
        addKeyShortcuts(dialog, sourceShortcutMap);

        dialog.layout.layout(true);
        refreshPreview();

        prepareDialogWindow(dialog, SCRIPT_NAME);
        var dialogResult = dialog.show();
        if (dialogResult !== 1) return null;

        return {
            sourceIndex: getCheckedIndex(sourceRadios),
            useVisibleBounds: optionCheckboxes.useVisibleBounds.value,
            keepOriginals: optionCheckboxes.keepOriginals.value,
            showCenter: optionCheckboxes.showCenter.value
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * エントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* 選択を解除する前にキーオブジェクトを調べる / Detect the key object while the selection is intact */
        var selectedItems = doc.selection;
        var snapshot = takeSelectionSnapshot(selectedItems, findKeyObjectIndex(selectedItems));

        /* 赤枠が見やすいように選択を解除 / Clear the selection so the red frame stands out */
        doc.selection = null;

        var previewController = createPreviewController(doc, snapshot);
        var mergeSettings = showMergeDialog(snapshot, previewController);
        previewController.dispose();

        if (!mergeSettings) {
            /* キャンセル: 選択を戻す / Cancelled: restore the selection */
            for (var i = 0; i < snapshot.items.length; i++) {
                snapshot.items[i].selected = true;
            }
            return;
        }

        applyMergeResult(doc, snapshot, mergeSettings);

        if (mergeSettings.showCenter) {
            runShowCenterAction();
        }
    }

    main();

})();
