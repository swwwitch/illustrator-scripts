#targetengine "AutoKerningPanelEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

自動カーニング（和文等幅／0／メトリクス／オプティカル）とプロポーショナルメトリクスだけを
1枚の常駐パレットにまとめ、ラジオやチェックボックスを操作した時点で選択中のテキストへ即時適用します。
統合文字組みパネル（UnifiedTypePanel）からカーニングまわりだけを抜き出したものです。

### 注意

- 適用範囲は選択が触れた段落全体です。テキスト編集モードで数文字だけ選んでも、その段落全体に当たります。
- 「メトリクス」を選んだときだけプロポーショナルメトリクスが ON になります。チェックボックスは単独でも
  操作でき、その場合はカーニング方式を変えずに切り替えます。
- UI に表示する値は、選択のうち最初に読めた段落の先頭文字です。設定が混在している選択では実態とズレます。

### Overview

A persistent palette limited to auto kerning (Japanese equal width / 0 / Metrics / Optical) and
proportional metrics. Operating a radio or the checkbox applies the change to the selected text
immediately. It is the kerning part of the Unified Type Panel, pulled out on its own.

### Notes

- Settings apply to every paragraph the selection touches, even when only a few characters are selected.
- Choosing Metrics turns proportional metrics on; any other choice turns it off. The checkbox can also be
  toggled on its own, which leaves the kerning method alone.
- The values shown come from the first readable paragraph's first character, so a mixed selection may not
  be reflected accurately.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoKerningPanel";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-22";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-07-29";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoKerningPanel.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoKerningPanel.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne7a198a4f527"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var currentLanguage = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        title: { ja: "カーニング設定", en: "Kerning Settings" },
        autoKern: { ja: "自動カーニング", en: "Auto Kerning" },
        proportionalMetrics: { ja: "プロポーショナルメトリクス", en: "Proportional Metrics" },
        tipAutoKern: { ja: "自動カーニング方式（和文等幅／0／メトリクス／オプティカル）。選択が触れた段落全体に適用します。", en: "Auto-kerning method (Japanese equal width / 0 / Metrics / Optical), applied to every paragraph the selection touches." },
        tipProportional: { ja: "プロポーショナルメトリクスの ON／OFF。メトリクスを選ぶと自動で ON になります。", en: "Turns proportional metrics on or off. Choosing Metrics turns it on automatically." },
        applyError: { ja: "適用に失敗しました", en: "Apply failed" }
    };

    /* 並び順がラジオの並び順 / Array order = radio order */
    var AUTO_KERN_OPTIONS = [
        { id: "mono", ja: "和文等幅", en: "Metrics - Roman Only" },
        { id: "zero", ja: "0", en: "0" },
        { id: "metrics", ja: "メトリクス", en: "Metrics" },
        { id: "optical", ja: "オプティカル", en: "Optical" }
    ];

    function getLocalizedText(entry) {
        return entry[currentLanguage] || entry.ja || entry.en || "";
    }

    // =========================================
    // メインエンジンで実行する DOM 処理 / DOM helpers run on the main engine
    //
    // toString() で連結して BridgeTalk 本文に同梱し、メインエンジンで eval される。
    // 注意: ここの関数には JSDoc を付けない（toString() の結果が壊れる）。
    // Shipped via toString(); never attach JSDoc — it corrupts the output.
    // =========================================

    /* グループは中を再帰 / Descend into groups */
    function collectTextRangesFromItem(item, selectedRanges) {
        if (!item) return;
        if (item.typename === "TextFrame") {
            selectedRanges.push(item.textRange);
        } else if (item.typename === "TextRange") {
            selectedRanges.push(item);
        } else if (item.typename === "GroupItem" && item.pageItems) {
            for (var i = 0; i < item.pageItems.length; i++) {
                collectTextRangesFromItem(item.pageItems[i], selectedRanges);
            }
        }
    }

    /* 選択が触れた段落を段落全体の範囲で取得 / Get the full paragraphs the selection touches */
    function getSelectedParagraphRanges() {
        var currentSelection = app.activeDocument.selection;
        var ranges = [];
        if (!currentSelection) return ranges;
        /* テキスト編集モードでは selection が TextRange / A text-edit selection is a TextRange */
        if (currentSelection.typename === "TextRange") {
            ranges.push(currentSelection);
        } else {
            for (var i = 0; i < currentSelection.length; i++) {
                collectTextRangesFromItem(currentSelection[i], ranges);
            }
        }
        var paragraphRanges = [];
        for (var j = 0; j < ranges.length; j++) {
            try {
                var paragraphs = ranges[j].paragraphs;
                for (var k = 0; k < paragraphs.length; k++) paragraphRanges.push(paragraphs[k]);
            } catch (e) { }
        }
        return paragraphRanges;
    }

    /* 和文等幅＝欧文のみメトリクス / Japanese equal width = metrics for Roman only */
    function resolveAutoKernType(methodId) {
        if (methodId === "mono") return AutoKernType.METRICSROMANONLY;
        if (methodId === "metrics") return AutoKernType.AUTO;
        if (methodId === "optical") return AutoKernType.OPTICAL;
        return AutoKernType.NOAUTOKERN;
    }

    function kernMethodToId(value) {
        var valueText = String(value);
        if (valueText === String(AutoKernType.AUTO)) return "metrics";
        if (valueText === String(AutoKernType.OPTICAL)) return "optical";
        if (valueText === String(AutoKernType.METRICSROMANONLY)) return "mono";
        if (valueText === String(AutoKernType.NOAUTOKERN)) return "zero";
        return "";
    }

    /* 戻り値は "OK:<payload>" / "ERR:<msg>"。getState の payload は kernId|proportionalMetrics
       Returns "OK:<payload>" / "ERR:<msg>"; getState's payload is kernId|proportionalMetrics */
    function dispatch(action, params) {
        if (app.documents.length === 0) return "ERR:nodoc";
        try { app.activeDocument; } catch (e) { return "ERR:nodoc"; }

        /* いずれも段落単位で適用 / Everything applies per paragraph */
        var ranges = getSelectedParagraphRanges();

        if (action === "getState") {
            /* 最初に読めた段落の先頭文字を代表値にする / The first readable paragraph's first character represents the selection */
            for (var i = 0; i < ranges.length; i++) {
                try {
                    var readAttributes = ranges[i].characters[0].characterAttributes;
                    return "OK:" + kernMethodToId(readAttributes.kerningMethod) + "|" + (readAttributes.proportionalMetrics ? 1 : 0);
                } catch (eRead) { }
            }
            return "OK:";
        }

        if (ranges.length === 0) return "OK:0";
        /* メトリクスのときだけプロポーショナルを ON / Proportional ON only for Metrics */
        var kerningMethod = (action === "apply") ? resolveAutoKernType(params.method) : null;
        var useProportional = (action === "apply") ? (kerningMethod === AutoKernType.AUTO) : params.proportionalMetrics;
        for (var j = 0; j < ranges.length; j++) {
            try {
                var charAttributes = ranges[j].characterAttributes;
                /* 方式を先に確定させてから上書き / Set the method first, then override */
                if (kerningMethod !== null) charAttributes.kerningMethod = kerningMethod;
                charAttributes.proportionalMetrics = useProportional ? true : false;
            } catch (eApply) { }
        }
        app.redraw();
        return "OK:" + ranges.length;
    }

    // =========================================
    // メインエンジン委譲 / Main-engine delegation (BridgeTalk)
    // =========================================

    var WORKER_FUNCS = [collectTextRangesFromItem, getSelectedParagraphRanges, resolveAutoKernType, kernMethodToId, dispatch];
    var WORKER_LIB_SRC = "";
    for (var workerFuncIndex = 0; workerFuncIndex < WORKER_FUNCS.length; workerFuncIndex++) {
        WORKER_LIB_SRC += WORKER_FUNCS[workerFuncIndex].toString() + "\n";
    }

    /* 非同期なので連打による委譲の重複を抑止 / Guard against overlapping async delegations */
    var workerBusy = false;

    /* BridgeTalk は本文のバックスラッシュをエスケープするため、コード全体を
       encodeURIComponent で包んで送り、ターゲットで復元する
       BridgeTalk escapes backslashes, so the code is sent percent-encoded and restored there */
    function runWorker(actionId, params, onDone) {
        if (workerBusy) return;
        workerBusy = true;

        /* method は自前の ID なのでエスケープ不要 / method is our own ASCII id */
        var parts = [];
        if (params && params.method !== undefined) parts.push('method:"' + params.method + '"');
        if (params && params.proportionalMetrics !== undefined) parts.push("proportionalMetrics:" + (params.proportionalMetrics ? "true" : "false"));
        var code = WORKER_LIB_SRC + '\nvar __r=dispatch("' + actionId + '",{' + parts.join(",") + "});__r;";

        var bridge = new BridgeTalk();
        bridge.target = "illustrator";
        bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(code) + "\"));";
        bridge.onResult = function (response) {
            workerBusy = false;
            var payload = response.body || "";
            var colonIndex = payload.indexOf(":");
            /* 失敗は握りつぶさず可視化 / Surface failures instead of swallowing them */
            if (payload.substring(0, colonIndex) === "OK") onDone(payload.substring(colonIndex + 1));
            else alert("⚠ " + getLocalizedText(LABELS.applyError) + " [" + actionId + "]: " + payload);
        };
        bridge.onError = function (response) {
            workerBusy = false;
            alert("⚠ " + getLocalizedText(LABELS.applyError) + " [" + actionId + "]: " + (response && response.body ? response.body : "BridgeTalk error"));
        };
        bridge.send();
    }

    // =========================================
    // パレット位置の保存・復元 / Palette position persistence
    // 移動のたびの位置を userData に "x,y" で保存し、次回起動時に同じ位置で開く。
    // =========================================

    function getPaletteLocationFile() {
        return File(Folder.userData.fsName + "/AutoKerningPanel_location.txt");
    }

    /* モニタ構成が変わっても画面外に復元しないための検証 / Never restore off-screen after a layout change */
    function isLocationOnScreen(x, y) {
        var screens;
        try { screens = $.screens; } catch (e) { return true; }
        if (!screens || !screens.length) return true;
        var margin = 40; /* タイトルバーを掴める余白 / Keep the title bar reachable */
        for (var i = 0; i < screens.length; i++) {
            var screen = screens[i];
            if (x >= screen.left - margin && x <= screen.right - margin &&
                y >= screen.top && y <= screen.bottom - margin) return true;
        }
        return false;
    }

    function savePaletteLocation(win) {
        var file = getPaletteLocationFile();
        file.encoding = "UTF-8";
        if (!file.open("w")) return;
        file.write(Math.round(win.location[0]) + "," + Math.round(win.location[1]));
        file.close();
    }

    function applySavedPaletteLocation(win) {
        var file = getPaletteLocationFile();
        if (!file.exists) return;
        file.encoding = "UTF-8";
        if (!file.open("r")) return;
        var fields = file.read().split(",");
        file.close();
        var x = parseInt(fields[0], 10), y = parseInt(fields[1], 10);
        if (isNaN(x) || isNaN(y) || !isLocationOnScreen(x, y)) return;
        win.location = [x, y];
    }

    // =========================================
    // UI構築 / Build UI
    // =========================================

    function createPalette() {
        var palette = new Window("palette", getLocalizedText(LABELS.title) + " " + SCRIPT_VERSION);
        palette.alignChildren = "fill";
        palette.margins = 15;
        palette.spacing = 10;

        var autoKernPanel = palette.add("panel", undefined, getLocalizedText(LABELS.autoKern));
        autoKernPanel.orientation = "column";
        autoKernPanel.alignChildren = ["left", "top"];
        autoKernPanel.margins = [10, 15, 10, 10];
        autoKernPanel.spacing = 6;
        autoKernPanel.helpTip = getLocalizedText(LABELS.tipAutoKern);

        var proportionalCheck = palette.add("checkbox", undefined, getLocalizedText(LABELS.proportionalMetrics));
        proportionalCheck.value = false;
        proportionalCheck.alignment = "left";
        proportionalCheck.helpTip = getLocalizedText(LABELS.tipProportional);

        var kernRadios = [];
        for (var i = 0; i < AUTO_KERN_OPTIONS.length; i++) {
            var kernRadio = autoKernPanel.add("radiobutton", undefined, getLocalizedText(AUTO_KERN_OPTIONS[i]));
            kernRadio.value = false;
            kernRadio.index = i;
            /* 表示も適用と同じ連動にする / Keep the checkbox in step with what is applied */
            kernRadio.onClick = function () {
                var methodId = AUTO_KERN_OPTIONS[this.index].id;
                proportionalCheck.value = (methodId === "metrics");
                runWorker("apply", { method: methodId }, function () { });
            };
            kernRadios.push(kernRadio);
        }

        /* カーニング方式は変えずに単独で適用 / Applied on its own, leaving the kerning method alone */
        proportionalCheck.onClick = function () {
            runWorker("applyProportional", { proportionalMetrics: this.value }, function () { });
        };

        function refreshState() {
            runWorker("getState", null, function (payload) {
                var fields = payload.split("|");
                if (!fields[0]) return; /* 読み取れなければ現在の表示を維持 / Keep the display when nothing could be read */
                for (var i = 0; i < kernRadios.length; i++) kernRadios[i].value = (AUTO_KERN_OPTIONS[i].id === fields[0]);
                proportionalCheck.value = (fields[1] === "1");
            });
        }

        /* 表示・フォーカス復帰のたびに選択状況を反映 / Reflect the selection on show and on regaining focus */
        palette.onShow = refreshState;
        palette.onActivate = refreshState;

        palette.addEventListener("keydown", function (event) {
            if (event.keyName === "Escape") palette.close();
        });

        return palette;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    function main() {
        /* 二重起動ガード：既存パレットがあれば前面化 / Single-instance guard */
        if ($.global.__AutoKerningPanel && $.global.__AutoKerningPanel.window) {
            try {
                var existingPalette = $.global.__AutoKerningPanel.window;
                existingPalette.show();
                existingPalette.active = true;
                return;
            } catch (existingError) {
                /* 参照が死んでいたら作り直す / Stale reference, rebuild */
                $.global.__AutoKerningPanel = null;
            }
        }

        var palette = createPalette();
        $.global.__AutoKerningPanel = { window: palette };

        /* 閉じた位置がそのまま次回の初期位置になる / Wherever it was closed becomes next launch's spot */
        palette.onMove = function () { savePaletteLocation(palette); };
        palette.addEventListener("close", function () {
            savePaletteLocation(palette);
            $.global.__AutoKerningPanel = null;
        });

        applySavedPaletteLocation(palette);
        palette.show();
        /* show 後にレイアウトが確定するため当て直す / Re-apply once the layout is final */
        applySavedPaletteLocation(palette);
    }

    main();

})();
