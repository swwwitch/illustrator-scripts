#target illustrator
#targetengine "AutoKerningEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの自動カーニング方式（和文等幅／0／メトリクス／オプティカル）をまとめて設定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerning.md

note記事も参照してください。
https://note.com/dtp_tranist/n/ne7a198a4f527

### Overview

Sets the auto-kerning method — Metrics (Roman Only), 0, Metrics or Optical — on the selected text.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerning.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoKerning";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerning.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerning.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne7a198a4f527"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PANEL_MARGINS = [16, 20, 16, 12];  /* パネル余白 [左,上,右,下] / panel margins */
    var BUTTON_WIDTH  = 90;                /* OK／キャンセルの幅 / OK and Cancel width */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILanguage();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "自動カーニング", en: "Auto Kerning" }
        },
        panel: {
            autoKern: { ja: "自動カーニング", en: "Auto Kerning" }
        },
        radio: {
            mono: { ja: "和文等幅", en: "Metrics - Roman Only" },
            zero: { ja: "0", en: "0" },
            metrics: { ja: "メトリクス", en: "Metrics" },
            optical: { ja: "オプティカル", en: "Optical" }
        },
        checkbox: {
            propMetrics: { ja: "プロポーショナルメトリクス", en: "Proportional Metrics" }
        },
        tooltip: {
            propMetrics: {
                ja: "カーニング方式に連動し、メトリクスのときだけオンになります。単独で切り替えると、カーニング方式は変えずにプロポーショナルメトリクスだけを設定します。",
                en: "Follows the kerning method and turns on only for Metrics. Toggling it on its own sets proportional metrics without changing the kerning method."
            }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            selectText: { ja: "テキストを選択してください", en: "Please select text" }
        }
    };

    /**
     * 表示言語のラベル文字列を返す
     * @param {Object} labelSet - { ja, en } のラベル
     * @returns {string} 表示言語の文字列（無ければ ja → en → 空文字の順に代替）
     */
    function getLabel(labelSet) {
        if (!labelSet) return "";
        return labelSet[uiLang] || labelSet.ja || labelSet.en || "";
    }

    // =========================================
    // 選択取得 / Selection
    // =========================================

    /**
     * 選択中のテキスト範囲を集める
     * テキスト編集モードでは selection が配列でなく TextRange になる
     * @returns {TextRange[]} 選択中のテキスト範囲（テキストフレームは全文の範囲）
     */
    function getSelectedTextRanges() {
        var currentSelection = app.activeDocument.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モード / Text-edit mode: the selection itself is a TextRange */
        if (currentSelection.typename === "TextRange") {
            selectedRanges.push(currentSelection);
            return selectedRanges;
        }
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            if (selectedItem.typename === "TextFrame") {
                selectedRanges.push(selectedItem.textRange);
            } else if (selectedItem.typename === "TextRange") {
                selectedRanges.push(selectedItem);
            }
        }
        return selectedRanges;
    }

    // =========================================
    // カーニング処理 / Kerning
    // =========================================

    /**
     * 自動カーニングの選択肢を作る
     * 和文等幅は欧文のみメトリクス＝和文は等幅（METRICSROMANONLY）
     * @returns {Object[]} { label: ラベル, value: AutoKernType } の配列
     */
    function createAutoKernOptions() {
        return [
            { label: LABELS.radio.mono, value: AutoKernType.METRICSROMANONLY },
            { label: LABELS.radio.zero, value: AutoKernType.NOAUTOKERN },
            { label: LABELS.radio.metrics, value: AutoKernType.AUTO },
            { label: LABELS.radio.optical, value: AutoKernType.OPTICAL }
        ];
    }

    /**
     * テキスト範囲にカーニング方式とプロポーショナルメトリクスを設定する
     * undefined の値は設定しない。設定できない範囲は飛ばす
     * @param {TextRange} textRange - 対象のテキスト範囲
     * @param {AutoKernType|undefined} kerningMethod - カーニング方式（undefined なら変えない）
     * @param {boolean|undefined} useProportionalMetrics - プロポーショナルメトリクス（undefined なら変えない）
     * @returns {void}
     */
    function setKerningAttributes(textRange, kerningMethod, useProportionalMetrics) {
        try {
            var charAttrs = textRange.characterAttributes;
            if (kerningMethod !== undefined) charAttrs.kerningMethod = kerningMethod;
            if (useProportionalMetrics !== undefined) charAttrs.proportionalMetrics = useProportionalMetrics;
        } catch (e) {
            /* 適用できない範囲はスキップ / Skip ranges that can't take these attributes */
        }
    }

    /**
     * 複数のテキスト範囲にカーニング方式とプロポーショナルメトリクスを設定する
     * @param {TextRange[]} textRanges - 対象のテキスト範囲
     * @param {AutoKernType|undefined} kerningMethod - カーニング方式（undefined なら変えない）
     * @param {boolean|undefined} useProportionalMetrics - プロポーショナルメトリクス（undefined なら変えない）
     * @returns {void}
     */
    function applyKerningAttributes(textRanges, kerningMethod, useProportionalMetrics) {
        for (var i = 0; i < textRanges.length; i++) {
            setKerningAttributes(textRanges[i], kerningMethod, useProportionalMetrics);
        }
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


    // =========================================
    // UI構築 / Build UI
    // =========================================

    /**
     * ダイアログを組み立てて参照を返す（イベントは未接続）
     * @param {Object[]} autoKernOptions - createAutoKernOptions() の選択肢
     * @returns {Object} kerningDialog / kernRadios / propMetricsCheckbox / btnOK / btnCancel
     */
    function buildDialog(autoKernOptions) {
        var kerningDialog = new Window("dialog", getLabel(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        kerningDialog.alignChildren = "fill";

        var autoKernPanel = kerningDialog.add("panel", undefined, getLabel(LABELS.panel.autoKern));
        autoKernPanel.orientation = "column";
        autoKernPanel.alignChildren = ["left", "top"];
        autoKernPanel.margins = PANEL_MARGINS;

        var kernRadios = [];
        for (var i = 0; i < autoKernOptions.length; i++) {
            var kernRadio = autoKernPanel.add("radiobutton", undefined, getLabel(autoKernOptions[i].label));
            kernRadio.value = (i === 0);
            kernRadio.index = i;
            kernRadios.push(kernRadio);
        }

        /* パネルの外に置く（方式に連動しつつ単独でも操作できる）
           Sits outside the panel: follows the method, but can also be toggled on its own */
        var propMetricsCheckbox = kerningDialog.add("checkbox", undefined, getLabel(LABELS.checkbox.propMetrics));
        propMetricsCheckbox.value = false;
        propMetricsCheckbox.alignment = "left";
        propMetricsCheckbox.helpTip = getLabel(LABELS.tooltip.propMetrics);

        var btnRowGroup = kerningDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";
        var btnCancel = btnRowGroup.add("button", undefined, getLabel(LABELS.button.cancel), { name: "cancel" });
        var btnOK = btnRowGroup.add("button", undefined, getLabel(LABELS.button.ok), { name: "ok" });
        btnOK.preferredSize.width = BUTTON_WIDTH;
        btnCancel.preferredSize.width = BUTTON_WIDTH;

        return {
            kerningDialog: kerningDialog,
            kernRadios: kernRadios,
            propMetricsCheckbox: propMetricsCheckbox,
            btnOK: btnOK,
            btnCancel: btnCancel
        };
    }

    /**
     * ダイアログにライブプレビューのイベントを接続する
     * プレビュー状態（確定フラグ・選択中インデックス）はこの関数内に保持
     * @param {Object} dialogControls - buildDialog() の戻り値
     * @param {TextRange[]} targetRanges - 対象のテキスト範囲
     * @param {Object[]} autoKernOptions - createAutoKernOptions() の選択肢
     * @returns {{finalizePreview: Function}} show() の後に呼ぶ後始末
     */
    function bindDialogEvents(dialogControls, targetRanges, autoKernOptions) {
        var selectedKernIndex = 0;
        /* OK で閉じたときだけ確定。Esc・ウィンドウを閉じる等は未確定のまま復元する
           Commit only when closed via OK; Esc / window-close etc. stay uncommitted and get reverted */
        var isCommitted = false;

        /* プレビュー前の元値を保持。app.undo() の段数に頼らず、元値の再代入で戻す
           Snapshot the original attributes; restore by reassigning them instead of counting app.undo() steps */
        var originalAttributes = [];
        for (var i = 0; i < targetRanges.length; i++) {
            var sourceAttrs = targetRanges[i].characterAttributes;
            originalAttributes.push({
                kerningMethod: sourceAttrs.kerningMethod,
                proportionalMetrics: sourceAttrs.proportionalMetrics
            });
        }

        /**
         * 選択中の方式を適用し、チェックボックスを連動させる
         * 毎回両方のプロパティを上書きするので、前回のプレビューを戻す必要はない
         * @returns {void}
         */
        function applyPreview() {
            var kerningMethod = autoKernOptions[selectedKernIndex].value;
            var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
            /* メトリクスのときだけプロポーショナルメトリクスをオン / Proportional metrics on only for Metrics */
            applyKerningAttributes(targetRanges, kerningMethod, useProportionalMetrics);
            dialogControls.propMetricsCheckbox.value = useProportionalMetrics;
            app.redraw();
        }

        /**
         * 元値を再代入してダイアログを開く前に戻す
         * 混在などで取得できなかった値（undefined）は戻さない
         * @returns {void}
         */
        function restoreOriginal() {
            for (var i = 0; i < targetRanges.length; i++) {
                setKerningAttributes(targetRanges[i], originalAttributes[i].kerningMethod, originalAttributes[i].proportionalMetrics);
            }
            app.redraw();
        }

        for (var r = 0; r < dialogControls.kernRadios.length; r++) {
            dialogControls.kernRadios[r].onClick = function () { selectedKernIndex = this.index; applyPreview(); };
        }

        /* カーニング方式は変えずに単独で適用 / Applied on its own, leaving the kerning method alone */
        dialogControls.propMetricsCheckbox.onClick = function () {
            applyKerningAttributes(targetRanges, undefined, this.value ? true : false);
            app.redraw();
        };

        /* OK のときだけ確定フラグを立てて閉じる / Set the commit flag and close only on OK */
        dialogControls.btnOK.onClick = function () {
            isCommitted = true;
            dialogControls.kerningDialog.close(1);
        };

        /* キャンセルは確定せず閉じる（復元は show() 後の finalizePreview に任せる）
           Cancel closes without committing (revert is handled by finalizePreview after show()) */
        dialogControls.btnCancel.onClick = function () {
            dialogControls.kerningDialog.close(2);
        };

        /* 初期プレビュー / Initial preview */
        applyPreview();

        return {
            /* show() 後に呼び出し、未確定なら元値を再代入して開く前へ戻す
               Call after show(): if not committed, restore the original values */
            finalizePreview: function () {
                if (!isCommitted) {
                    restoreOriginal();
                }
            }
        };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択テキストを確かめ、ダイアログを開いてプレビューを確定または破棄する
     * @returns {void}
     */
    function main() {
        if (app.documents.length <= 0) {
            return;
        }

        var targetRanges = getSelectedTextRanges();
        if (targetRanges.length === 0) {
            alert(getLabel(LABELS.alert.selectText));
            return;
        }

        var autoKernOptions = createAutoKernOptions();
        var dialogControls = buildDialog(autoKernOptions);
        var previewSession = bindDialogEvents(dialogControls, targetRanges, autoKernOptions);
        prepareDialogWindow(dialogControls.kerningDialog, SCRIPT_NAME);
        dialogControls.kerningDialog.show();
        /* OK 以外で閉じた場合はプレビューを破棄 / Discard preview unless closed via OK */
        previewSession.finalizePreview();
    }

    main();

})();
