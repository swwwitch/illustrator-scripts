#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトを高解像度でラスタライズし、PNG相当のビットマップとしてクリップボードへコピーします。
解像度・背景色・アンチエイリアス・余白は、実行時のダイアログで指定できます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyAsPngLikeFigmaWithDialog.md

### Overview

Rasterizes the selection at high resolution and copies it to the clipboard as a PNG-equivalent bitmap.
Resolution, background color, anti-aliasing and margin are all set in a dialog when the script runs.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyAsPngLikeFigmaWithDialog.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CopyAsPngLikeFigmaWithDialog"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyAsPngLikeFigmaWithDialog.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyAsPngLikeFigmaWithDialog.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DPI_CHOICES       = ["72", "150", "300", "600", "1200"]; /* 解像度の選択肢（dpi）/ Resolution choices (dpi) */
    var DEFAULT_DPI_INDEX = 3;                                   /* 初期選択（600dpi）/ Initial choice (600 dpi) */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var MARGIN_INPUT_CHARS  = 4;                  /* 余白欄の文字数 / Margin field width in characters */
    var PROGRESS_BAR_BOUNDS = [20, 20, 300, 10];  /* プログレスバーの bounds / Progress bar bounds */

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var uiLang = ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';

    var LABELS = {
        dialog: {
            title: { ja: "ビットマップとしてコピー", en: "Copy as PNG" }
        },
        fieldLabel: {
            dpi: { ja: "解像度", en: "Resolution" },
            background: { ja: "背景", en: "Background" },
            margin: { ja: "余白", en: "Margin" }
        },
        unit: {
            dpi: { ja: "（dpi）", en: "(dpi)" }
        },
        radio: {
            transparent: { ja: "透明", en: "Transparent" },
            white: { ja: "白", en: "White" },
            black: { ja: "黒", en: "Black" }
        },
        checkbox: {
            antialias: { ja: "アンチエイリアスを有効にする", en: "Enable Anti-Aliasing" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        status: {
            start: { ja: "処理を開始しています...", en: "Starting process..." },
            createLayer: { ja: "一時レイヤー作成中...", en: "Creating temporary layer..." },
            duplicate: { ja: "複製中...", en: "Duplicating..." },
            createArea: { ja: "ラスタライズ範囲作成中...", en: "Creating rasterize bounds..." },
            rasterize: { ja: "ラスタライズ中...", en: "Rasterizing..." },
            resize: { ja: "拡大中...", en: "Resizing..." },
            copy: { ja: "コピー中...", en: "Copying..." },
            cleanUp: { ja: "一時オブジェクト削除中...", en: "Deleting temporary objects..." },
            done: { ja: "完了しました。", en: "Done." }
        },
        alert: {
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            copiedAfterDpi: { ja: "dpi（", en: "dpi (" },
            copiedAfterRatio: { ja: "）でビットマップコピーしました。", en: ") bitmap copied." }
        },
        tooltip: {
            dpi: {
                ja: "コピーするビットマップの解像度です。数値が大きいほど精細になり、処理時間も伸びます。",
                en: "Resolution of the copied bitmap. Higher values are sharper but take longer."
            },
            background: {
                ja: "ビットマップの背景です。「透明」はアルファ付きでコピーします。",
                en: "Background of the bitmap. Transparent copies it with an alpha channel."
            },
            margin: { ja: "選択範囲の外側に足す余白です。", en: "Extra space added around the selection." },
            antialias: {
                ja: "輪郭をなめらかにします。オフにするとドットのままコピーします。",
                en: "Smooths the edges. Off copies the pixels as they are."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで現在の言語のラベルを取得する（例: getLabel("status.copy")）
     * @param {string} labelPath - LABELS を辿るパス
     * @returns {string} 該当するラベル
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
        }
        return labelNode[uiLang];
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 設定ダイアログを組み立てる
     * @returns {{dialog: Window, controls: Object}} ダイアログと設定を読むコントロール
     */
    function buildDialog() {
        var rasterizeDialog = new Window("dialog", getLabel("dialog.title"));
        rasterizeDialog.orientation = "column";
        rasterizeDialog.alignChildren = "left";

        /* 解像度 / Resolution */
        var dpiGroup = rasterizeDialog.add("group");
        dpiGroup.orientation = "row";
        dpiGroup.add("statictext", undefined, labelText("fieldLabel.dpi"));
        var dpiDropdown = dpiGroup.add("dropdownlist", undefined, DPI_CHOICES);
        dpiDropdown.helpTip = getLabel("tooltip.dpi");
        dpiGroup.add("statictext", undefined, getLabel("unit.dpi"));
        dpiDropdown.selection = DEFAULT_DPI_INDEX;

        /* 背景 / Background */
        var backgroundGroup = rasterizeDialog.add("group");
        backgroundGroup.orientation = "row";
        backgroundGroup.alignChildren = "left";
        backgroundGroup.add("statictext", undefined, labelText("fieldLabel.background"));
        var transparentRadio = backgroundGroup.add("radiobutton", undefined, getLabel("radio.transparent"));
        transparentRadio.helpTip = getLabel("tooltip.background");
        var whiteRadio = backgroundGroup.add("radiobutton", undefined, getLabel("radio.white"));
        whiteRadio.helpTip = getLabel("tooltip.background");
        var blackRadio = backgroundGroup.add("radiobutton", undefined, getLabel("radio.black"));
        blackRadio.helpTip = getLabel("tooltip.background");
        whiteRadio.value = true;

        /* 余白 / Margin */
        var marginGroup = rasterizeDialog.add("group");
        marginGroup.orientation = "row";
        marginGroup.add("statictext", undefined, labelText("fieldLabel.margin"));
        var marginInput = marginGroup.add("edittext", undefined, "0");
        marginInput.helpTip = getLabel("tooltip.margin");
        marginInput.characters = MARGIN_INPUT_CHARS;

        /* アンチエイリアス / Anti-aliasing */
        var antialiasGroup = rasterizeDialog.add("group");
        var antialiasCheckbox = antialiasGroup.add("checkbox", undefined, getLabel("checkbox.antialias"));
        antialiasCheckbox.helpTip = getLabel("tooltip.antialias");
        antialiasCheckbox.value = true;

        /* ボタン / Buttons */
        var btnRowGroup = rasterizeDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = "right";
        btnRowGroup.add("button", undefined, getLabel("button.cancel"));
        btnRowGroup.add("button", undefined, getLabel("button.ok"));

        return {
            dialog: rasterizeDialog,
            controls: {
                dpiDropdown: dpiDropdown,
                transparentRadio: transparentRadio,
                blackRadio: blackRadio,
                antialiasCheckbox: antialiasCheckbox,
                marginInput: marginInput
            }
        };
    }

    /**
     * ダイアログのコントロールから設定を読み取る
     * @param {Object} dialogControls - buildDialog() が返したコントロール
     * @returns {{dpi: number, background: string, antiAliasing: boolean, margin: number}} 設定
     */
    function readDialogSettings(dialogControls) {
        return {
            dpi: parseInt(dialogControls.dpiDropdown.selection.text, 10),
            background: dialogControls.transparentRadio.value ? "transparent" : (dialogControls.blackRadio.value ? "black" : "white"),
            antiAliasing: dialogControls.antialiasCheckbox.value,
            margin: parseInt(dialogControls.marginInput.text, 10)
        };
    }

    /**
     * 設定ダイアログを表示し、OK なら設定を返す
     * @returns {Object|null} 設定（キャンセル時は null）
     */
    function showRasterizeDialog() {
        var dialogParts = buildDialog();
        if (dialogParts.dialog.show() != 1) return null;
        return readDialogSettings(dialogParts.controls);
    }

    // =========================================
    // 進行状況 / Progress
    // =========================================

    /**
     * 進行状況を表示するパレットを開き、更新・終了用の関数を返す
     * @param {string} title - パレットのタイトル
     * @param {string} initialText - 最初に表示する文言
     * @returns {{update: Function, close: Function}} update(値, 文言) と close()
     */
    function openProgressWindow(title, initialText) {
        var progressWin = new Window("palette", title);
        var progressBar = progressWin.add("progressbar", PROGRESS_BAR_BOUNDS, 0, 100);
        var progressText = progressWin.add("statictext", undefined, initialText);
        progressWin.show();
        return {
            update: function (value, text) {
                progressBar.value = value;
                progressText.text = text;
                progressWin.update();
            },
            close: function () {
                progressWin.close();
            }
        };
    }

    // =========================================
    // ラスタライズ / Rasterize
    // =========================================

    /**
     * 作業用の一時レイヤーを作成する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 一時レイヤー
     */
    function createTempLayer(doc) {
        var tempLayer = doc.layers.add();
        tempLayer.name = "__TEMP_LAYER__";
        tempLayer.locked = false;
        tempLayer.visible = true;
        return tempLayer;
    }

    /**
     * オブジェクトを指定グループの末尾へ複製する
     * @param {PageItem[]} sourceItems - 複製元のオブジェクト
     * @param {GroupItem} targetGroup - 複製先のグループ
     * @returns {PageItem[]} 複製したオブジェクト
     */
    function duplicateIntoGroup(sourceItems, targetGroup) {
        var duplicatedItems = [];
        for (var i = 0; i < sourceItems.length; i++) {
            duplicatedItems.push(sourceItems[i].duplicate(targetGroup, ElementPlacement.PLACEATEND));
        }
        return duplicatedItems;
    }

    /**
     * 余白を足したラスタライズ範囲の長方形（塗り・線なし）を一時レイヤーの最前面に作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} tempLayer - 一時レイヤー
     * @param {number[]} areaBounds - 範囲 [left, top, right, bottom]
     * @param {number} margin - 外側に足す余白（pt）
     * @returns {PathItem} 範囲の長方形
     */
    function createRasterAreaRect(doc, tempLayer, areaBounds, margin) {
        var areaRect = doc.pathItems.rectangle(
            areaBounds[1] + margin,
            areaBounds[0] - margin,
            (areaBounds[2] - areaBounds[0]) + margin * 2,
            (areaBounds[1] - areaBounds[3]) + margin * 2
        );
        areaRect.stroked = false;
        areaRect.filled = false;
        areaRect.move(tempLayer, ElementPlacement.PLACEATBEGINNING);
        return areaRect;
    }

    /**
     * 設定からラスタライズオプションを作る
     * @param {Object} settings - ダイアログの設定
     * @returns {RasterizeOptions} ラスタライズオプション
     */
    function buildRasterizeOptions(settings) {
        var rasterizeOptions = new RasterizeOptions();
        rasterizeOptions.resolution = settings.dpi;
        rasterizeOptions.transparency = (settings.background === "transparent");
        rasterizeOptions.antiAliasing = settings.antiAliasing;
        if (!rasterizeOptions.transparency) {
            rasterizeOptions.backgroundBlack = (settings.background === "black");
        }
        return rasterizeOptions;
    }

    /**
     * オブジェクトを削除する。削除済み・削除できないものは飛ばす
     * @param {Object} target - 削除するオブジェクトまたはレイヤー
     * @returns {void}
     */
    function removeQuietly(target) {
        /* ラスタライズで元のグループは消えるので、参照が無効なことがある / Rasterizing disposes the source group, so references may be stale */
        try {
            target.remove();
        } catch (e) {}
    }

    /**
     * 選択オブジェクトを一時レイヤーで複製・ラスタライズ・拡大してクリップボードへコピーし、一時オブジェクトを消す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} sourceItems - 選択オブジェクト
     * @param {Object} settings - ダイアログの設定
     * @param {{update: Function, close: Function}} progress - 進行状況のパレット
     * @returns {number} 拡大率（%）
     */
    function copySelectionAsBitmap(doc, sourceItems, settings, progress) {
        progress.update(10, getLabel("status.createLayer"));
        var tempLayer = createTempLayer(doc);

        progress.update(20, getLabel("status.duplicate"));
        var tempGroup = tempLayer.groupItems.add();
        var duplicatedItems = duplicateIntoGroup(sourceItems, tempGroup);

        progress.update(35, getLabel("status.createArea"));
        var areaRect = createRasterAreaRect(doc, tempLayer, tempGroup.visibleBounds, settings.margin || 0);

        progress.update(50, getLabel("status.rasterize"));
        var rasterItem = doc.rasterize(tempGroup, areaRect.geometricBounds, buildRasterizeOptions(settings));

        progress.update(70, getLabel("status.resize"));
        var resizeRatio = Math.round((settings.dpi / 72) * 100);
        rasterItem.resize(resizeRatio, resizeRatio, true, true, true, true, resizeRatio, Transformation.CENTER);

        progress.update(85, getLabel("status.copy"));
        doc.selection = [rasterItem];
        app.executeMenuCommand("copy");

        progress.update(95, getLabel("status.cleanUp"));
        var tempObjects = [rasterItem, areaRect].concat(duplicatedItems, [tempGroup, tempLayer]);
        for (var i = 0; i < tempObjects.length; i++) {
            removeQuietly(tempObjects[i]);
        }

        doc.selection = sourceItems;

        progress.update(100, getLabel("status.done"));
        return resizeRatio;
    }

    /**
     * 進行状況を表示しながらビットマップコピーを実行し、結果を知らせる
     * @param {Object} settings - ダイアログの設定
     * @returns {void}
     */
    function runRasterizeWithSettings(settings) {
        var doc = app.activeDocument;
        var originalSelection = doc.selection;

        var progress = openProgressWindow("処理中...", getLabel("status.start"));

        /* 途中で失敗したらパレットを閉じてエラーを表示 / Close the palette and report the error on failure */
        try {
            var resizeRatio = copySelectionAsBitmap(doc, originalSelection, settings, progress);
            progress.close();
            alert(settings.dpi + getLabel("alert.copiedAfterDpi") + resizeRatio + "%" + getLabel("alert.copiedAfterRatio"));
        } catch (err) {
            progress.close();
            alert("エラーが発生しました: " + err.message);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 設定ダイアログを出し、選択オブジェクトをビットマップとしてクリップボードへコピーする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var settings = showRasterizeDialog();
        if (settings) {
            runRasterizeWithSettings(settings);
        }
    }

    main();

})();
