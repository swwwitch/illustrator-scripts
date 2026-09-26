#target illustrator
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
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

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
        dialogControls.kerningDialog.show();
        /* OK 以外で閉じた場合はプレビューを破棄 / Discard preview unless closed via OK */
        previewSession.finalizePreview();
    }

    main();

})();
