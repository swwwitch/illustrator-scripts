#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの自動カーニング方式（和文等幅／0／メトリクス／オプティカル）を
まとめて設定するスクリプトです。

- ラジオボタンで方式を選ぶとライブプレビューで確認できる
- テキストオブジェクト全体（選択範囲全体）に適用する
- 「メトリクス」を選んだときだけプロポーショナルメトリクスをON、それ以外はOFF
- キャンセルで開く前の状態に戻す

### Overview

Sets the auto-kerning method (Metrics - Roman Only / 0 / Metrics / Optical)
for the selected text.

- Pick a method with the radio buttons; live preview shows the result
- Applies to the whole text object (entire selection)
- Proportional metrics is turned ON only for "Metrics", OFF otherwise
- Cancel restores the state before opening

*/

(function () {

    // =========================================
    // バージョン / Version
    // =========================================
    var SCRIPT_VERSION = "v1.0";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 言語判定 / Detect UI language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* ラベル定義 / Label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "自動カーニング", en: "Auto Kerning" }
        },
        field: {
            autoKern: { ja: "自動カーニング", en: "Auto Kerning" }
        },
        autoKern: {
            mono: { ja: "和文等幅", en: "Metrics - Roman Only" },
            zero: { ja: "0", en: "0" },
            metrics: { ja: "メトリクス", en: "Metrics" },
            optical: { ja: "オプティカル", en: "Optical" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            selectText: { ja: "テキストを選択してください", en: "Please select text" }
        }
    };

    /* 言語に応じたラベル文字列を取得 / Resolve a label string for the current language */
    function getLocalizedText(entry) {
        if (!entry) return "";
        return entry[currentLanguage] || entry.ja || entry.en || "";
    }

    // =========================================
    // 選択取得 / Selection
    // =========================================

    /* 型名を安全に取得（host オブジェクトは typename を優先、JS オブジェクトは constructor.name）
       Safely resolve a type name: typename for host objects, constructor.name for JS objects */
    function getTypeName(obj) {
        if (obj === null || obj === undefined) return "";
        if (obj.typename) return obj.typename;
        try {
            return obj.constructor ? obj.constructor.name : "";
        } catch (e) {
            return "";
        }
    }

    /* 選択中のテキスト範囲を取得 / Get selected text ranges from current document */
    function getSelectedTextRanges() {
        var activeDoc = app.activeDocument;
        var currentSelection = activeDoc.selection;
        var selectedRanges = [];
        if (!currentSelection) return selectedRanges;
        /* テキスト編集モードでは selection が配列でなく TextRange になる / In text-edit mode the selection is a TextRange, not an array */
        if (getTypeName(currentSelection) === "TextRange") {
            selectedRanges.push(currentSelection);
            return selectedRanges;
        }
        if (currentSelection.length === 0) return selectedRanges;
        for (var i = 0; i < currentSelection.length; i++) {
            var selectedItem = currentSelection[i];
            var itemType = getTypeName(selectedItem);
            if (itemType === "TextFrame") {
                selectedRanges.push(selectedItem.textRange);
            } else if (itemType === "TextRange") {
                selectedRanges.push(selectedItem);
            }
        }
        return selectedRanges;
    }

    // =========================================
    // カーニング処理 / Kerning
    // =========================================

    /* 自動カーニングの選択肢を生成 / Build the auto-kerning option list
       和文等幅は欧文のみメトリクス＝和文は等幅（METRICSROMANONLY）
       Japanese equal width = metrics for Roman only (METRICSROMANONLY) */
    function createAutoKernOptions() {
        return [
            { label: LABELS.autoKern.mono, value: AutoKernType.METRICSROMANONLY },
            { label: LABELS.autoKern.zero, value: AutoKernType.NOAUTOKERN },
            { label: LABELS.autoKern.metrics, value: AutoKernType.AUTO },
            { label: LABELS.autoKern.optical, value: AutoKernType.OPTICAL }
        ];
    }

    /* 選択範囲にカーニング方式を適用 / Apply a kerning method to the given ranges
       メトリクスのときのみプロポーショナルメトリクスをON、それ以外はOFF
       Proportional metrics ON only for Metrics, OFF otherwise */
    function applyKerningToRanges(ranges, kerningMethod) {
        var useProportionalMetrics = (kerningMethod === AutoKernType.AUTO);
        for (var i = 0; i < ranges.length; i++) {
            try {
                ranges[i].characterAttributes.kerningMethod = kerningMethod;
                ranges[i].characterAttributes.proportionalMetrics = useProportionalMetrics;
            } catch (e) {
                // 適用できない範囲はスキップ / Skip ranges that can't take these attributes
            }
        }
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================

    /* ダイアログを組み立てて参照を返す（イベント未接続）/ Build the dialog and return references (events not wired yet) */
    function createDialogUI(autoKernOptions) {
        var dialog = new Window("dialog", getLocalizedText(LABELS.dialog.title) + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";

        var autoKernPanel = dialog.add("panel", undefined, getLocalizedText(LABELS.field.autoKern));
        autoKernPanel.orientation = "column";
        autoKernPanel.alignChildren = ["left", "top"];
        autoKernPanel.margins = [16, 20, 16, 12];

        var kernRadios = [];
        for (var i = 0; i < autoKernOptions.length; i++) {
            var kernRadio = autoKernPanel.add("radiobutton", undefined, getLocalizedText(autoKernOptions[i].label));
            kernRadio.value = (i === 0);
            kernRadio.index = i;
            kernRadios.push(kernRadio);
        }

        var buttonGroup = dialog.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "right";
        var cancelButton = buttonGroup.add("button", undefined, getLocalizedText(LABELS.button.cancel), { name: "cancel" });
        var okButton = buttonGroup.add("button", undefined, getLocalizedText(LABELS.button.ok), { name: "ok" });
        okButton.preferredSize.width = 90;
        cancelButton.preferredSize.width = 90;

        return {
            dialog: dialog,
            kernRadios: kernRadios,
            okButton: okButton,
            cancelButton: cancelButton
        };
    }

    /* ダイアログにライブプレビューのイベントを接続 / Wire live-preview events to the dialog
       プレビュー状態（適用済みフラグ・選択中インデックス）はこの関数内に保持
       Preview state (applied flag, selected index) is kept inside this function */
    function bindDialogEvents(ui, targetRanges, autoKernOptions) {
        var selectedKernIndex = 0;
        // OK で閉じたときだけ確定。Esc・ウィンドウ閉じる等は未確定のまま復元する
        // Commit only when closed via OK; Esc / window-close etc. stay uncommitted and get reverted
        var isCommitted = false;

        // ---- プレビュー前の元値を保持 / Snapshot the original attributes before previewing ----
        // app.undo() の段数前提に依存せず、復元は元値の再代入で行う
        // Avoid relying on app.undo() step counting; restore by reassigning the captured values
        var originalAttributes = [];
        for (var s = 0; s < targetRanges.length; s++) {
            var sourceAttrs = targetRanges[s].characterAttributes;
            originalAttributes.push({
                kerningMethod: sourceAttrs.kerningMethod,
                proportionalMetrics: sourceAttrs.proportionalMetrics
            });
        }

        function applyPreview() {
            // 各適用で両プロパティを一律に上書きするため、前回プレビューの取り消しは不要
            // Each apply overwrites both properties uniformly, so no prior revert is needed
            applyKerningToRanges(targetRanges, autoKernOptions[selectedKernIndex].value);
            app.redraw();
        }

        function restoreOriginal() {
            for (var r = 0; r < targetRanges.length; r++) {
                try {
                    var targetAttrs = targetRanges[r].characterAttributes;
                    var original = originalAttributes[r];
                    // 混在等で取得できなかった値（undefined）は復元しない / Skip values that couldn't be captured (undefined)
                    if (original.kerningMethod !== undefined) {
                        targetAttrs.kerningMethod = original.kerningMethod;
                    }
                    if (original.proportionalMetrics !== undefined) {
                        targetAttrs.proportionalMetrics = original.proportionalMetrics;
                    }
                } catch (e) {
                    // 復元できない範囲はスキップ / Skip ranges that can't be restored
                }
            }
            app.redraw();
        }

        for (var i = 0; i < ui.kernRadios.length; i++) {
            ui.kernRadios[i].onClick = function () { selectedKernIndex = this.index; applyPreview(); };
        }

        // OK のときだけ確定フラグを立てて閉じる / Set the commit flag and close only on OK
        ui.okButton.onClick = function () {
            isCommitted = true;
            ui.dialog.close(1);
        };

        // キャンセルは確定せず閉じる（復元は show() 後の finalizePreview に委譲）
        // Cancel closes without committing (revert is handled by finalizePreview after show())
        ui.cancelButton.onClick = function () {
            ui.dialog.close(2);
        };

        // 初期プレビュー / Initial preview
        applyPreview();

        /* show() 後に呼び出し、未確定なら元値を再代入して開く前へ戻す
           Call after show(): if not committed, restore the original values */
        return {
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
    function main() {
        if (app.documents.length <= 0) {
            return;
        }

        var targetRanges = getSelectedTextRanges();
        if (targetRanges.length === 0) {
            alert(getLocalizedText(LABELS.alert.selectText));
            return;
        }

        var autoKernOptions = createAutoKernOptions();
        var ui = createDialogUI(autoKernOptions);
        var preview = bindDialogEvents(ui, targetRanges, autoKernOptions);
        ui.dialog.show();
        // OK 以外で閉じた場合はプレビューを破棄 / Discard preview unless closed via OK
        preview.finalizePreview();
    }

    main();

})();
