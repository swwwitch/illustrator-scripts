#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択中のオブジェクトを高解像度でラスタライズし、Figmaの「Copy as PNG」のようにビットマップとしてクリップボードへコピーします。
600dpiでラスタライズしたあと、72ppi相当の偶数整数倍率に調整します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyAsPngLikeFigma.md

note記事も参照してください。
https://note.com/dtp_tranist/n/nf5f269788086

### Overview

Rasterizes the selection at high resolution and copies it to the clipboard as a bitmap, the way Figma's "Copy as PNG" does.
It rasterizes at 600 dpi, then scales the result to an even integer multiple of 72 ppi.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyAsPngLikeFigma.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CopyAsPngLikeFigma";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-05-02";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CopyAsPngLikeFigma.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CopyAsPngLikeFigma.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nf5f269788086"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var RASTER_RESOLUTION = 600; /* ラスタライズの解像度（dpi）/ Rasterize resolution (dpi) */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var PROGRESS_BAR_BOUNDS = [20, 20, 300, 10]; /* プログレスバーの bounds / Progress bar bounds */

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
     * 項目名の文言の末尾にコロンを付ける（日本語は半角スペース＋半角コロン「 :」、英語は「:」。Illustrator の線パネルなどの項目名に合わせる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? " :" : ":");
    }

    /**
     * 「項目名 : 値」の1行を返す（日本語は「件数 : 5」、英語は「Count: 5」。どちらもコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + " " + value;
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

    var LABELS = {
        dialog: {
            progressTitle: { ja: "処理中...", en: "Processing..." }
        },
        status: {
            start: { ja: "処理を開始しています...", en: "Starting process..." },
            createLayer: { ja: "一時レイヤー作成中...", en: "Creating temp layer..." },
            duplicate: { ja: "複製中...", en: "Duplicating..." },
            createArea: { ja: "ラスタライズ範囲作成中...", en: "Creating raster area..." },
            rasterize: { ja: "ラスタライズ中...", en: "Rasterizing..." },
            resize: { ja: "拡大中...", en: "Resizing..." },
            copy: { ja: "コピー中...", en: "Copying..." },
            cleanUp: { ja: "一時オブジェクト削除中...", en: "Cleaning up..." },
            done: { ja: "完了しました。", en: "Completed." }
        },
        alert: {
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select objects." },
            error: { ja: "エラーが発生しました: ", en: "An error occurred: " }
        }
    };

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
     * ラスタライズ範囲になる塗り・線なしの長方形を一時レイヤーの最前面に作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Layer} tempLayer - 一時レイヤー
     * @param {number[]} areaBounds - 範囲 [left, top, right, bottom]
     * @returns {PathItem} 範囲の長方形
     */
    function createRasterAreaRect(doc, tempLayer, areaBounds) {
        var areaRect = doc.pathItems.rectangle(areaBounds[1], areaBounds[0], areaBounds[2] - areaBounds[0], areaBounds[1] - areaBounds[3]);
        areaRect.stroked = false;
        areaRect.filled = false;
        areaRect.move(tempLayer, ElementPlacement.PLACEATBEGINNING);
        return areaRect;
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
     * @param {{update: Function, close: Function}} progress - 進行状況のパレット
     * @returns {void}
     */
    function copySelectionAsBitmap(doc, sourceItems, progress) {
        progress.update(10, getLabel("status.createLayer"));
        var tempLayer = createTempLayer(doc);

        progress.update(20, getLabel("status.duplicate"));
        var tempGroup = tempLayer.groupItems.add();
        var duplicatedItems = duplicateIntoGroup(sourceItems, tempGroup);

        progress.update(35, getLabel("status.createArea"));
        var areaRect = createRasterAreaRect(doc, tempLayer, tempGroup.visibleBounds);

        progress.update(50, getLabel("status.rasterize"));
        var rasterizeOptions = new RasterizeOptions();
        rasterizeOptions.resolution = RASTER_RESOLUTION;
        rasterizeOptions.transparency = false;
        rasterizeOptions.backgroundBlack = false;
        rasterizeOptions.antiAliasing = true;

        var rasterItem = doc.rasterize(tempGroup, areaRect.geometricBounds, rasterizeOptions);

        progress.update(70, getLabel("status.resize"));
        /* 拡大倍率を「72ppi相当」を基準に偶数の整数へ切り上げ / Round the scale up to an even integer based on 72 ppi */
        var baseRatio = (RASTER_RESOLUTION / 72) * 100;
        var resizeRatio = Math.ceil(baseRatio / 2) * 2;
        rasterItem.resize(resizeRatio, resizeRatio);

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
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択オブジェクトをビットマップとしてクリップボードへコピーする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0 || app.activeDocument.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var doc = app.activeDocument;
        var originalSelection = doc.selection;

        var progress = openProgressWindow(getLabel("dialog.progressTitle"), getLabel("status.start"));

        /* 途中で失敗したらパレットを閉じてエラーを表示 / Close the palette and report the error on failure */
        try {
            copySelectionAsBitmap(doc, originalSelection, progress);
        } catch (err) {
            progress.close();
            alert(getLabel("alert.error") + err.message);
            return;
        }

        progress.close();
    }

    main();

})();
