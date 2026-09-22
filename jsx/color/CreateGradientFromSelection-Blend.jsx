#targetengine "CreateGradientFromSelectionEngine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの塗り／線カラーを配置順（左→右、上→下）で抽出し、スウォッチグループに登録してグラデーションを自動生成します。ブレンド版です。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGradientFromSelection-Blend.md

### Overview

Extracts the fill and stroke colors of the selection in layout order, registers them as a swatch group, and builds a gradient from them. This is the blend variant.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGradientFromSelection-Blend.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CreateGradientFromSelection-Blend"; /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGradientFromSelection-Blend.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGradientFromSelection-Blend.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 長方形サイズの既定値（pt）/ Default rectangle size in points */
    var DEFAULT_RECT_SIZE = 100;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* パネルの余白 [左, 上, 右, 下] / Panel margins [left, top, right, bottom] */
    var PANEL_MARGINS = [15, 20, 15, 10];

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "グラデーション作成", en: "Create Gradient" }
        },
        panel: {
            color: { ja: "カラー", en: "Colors" },
            rect: { ja: "長方形（適用）", en: "Rectangle" }
        },
        checkbox: {
            globalColor: { ja: "グローバルカラー", en: "Global colors" },
            createGradient: { ja: "グラデーションを作成", en: "Create gradient" },
            createRect: { ja: "長方形を作成し、グラデーションを適用", en: "Create rectangle and apply gradient" },
            useSelectionSize: { ja: "選択オブジェクトに合わせてサイズ指定", en: "Match rectangle size to selection" },
            registerGraphicStyle: { ja: "グラフィックスタイルに登録", en: "Save as Graphic Style" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            globalColor: {
                ja: "選択オブジェクトの塗りをグローバルカラーとしてスウォッチに登録します。",
                en: "Registers the fills of the selection as global color swatches."
            },
            createGradient: {
                ja: "選択オブジェクトの塗りを順に並べたグラデーションを作ります。",
                en: "Builds a gradient from the fills of the selection, in order."
            },
            createRect: {
                ja: "作ったグラデーションを適用した長方形を描きます。",
                en: "Draws a rectangle filled with the new gradient."
            },
            useSelectionSize: {
                ja: "長方形のサイズを選択範囲に合わせます。オフのときは既定のサイズで描きます。",
                en: "Matches the rectangle to the size of the selection. Off uses the default size."
            },
            registerGraphicStyle: {
                ja: "作ったグラデーションをグラフィックスタイルとして登録します。",
                en: "Saves the new gradient as a graphic style."
            }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "checkbox.globalColor" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelPath;
    }

    // =========================================
    // セッション設定 / Session Settings
    // =========================================

    /**
     * ダイアログの値を保持するオブジェクトを返す（targetengine 内だけで保持し、Illustrator の再起動で消える）
     * @returns {Object} 設定の保存先
     */
    function getSessionSettings() {
        if (!$.global.__CGFS_SETTINGS) $.global.__CGFS_SETTINGS = {};
        return $.global.__CGFS_SETTINGS;
    }

    /**
     * 保持している真偽値を読み出す
     * @param {string} settingKey - 設定のキー
     * @param {boolean} defaultValue - 保持していないときの値
     * @returns {boolean} 保持している値、または defaultValue
     */
    function loadBool(settingKey, defaultValue) {
        var sessionSettings = getSessionSettings();
        if (typeof sessionSettings[settingKey] === "boolean") return sessionSettings[settingKey];
        return defaultValue;
    }

    /**
     * 真偽値を保持する
     * @param {string} settingKey - 設定のキー
     * @param {boolean} value - 保持する値
     * @returns {void}
     */
    function saveBool(settingKey, value) {
        getSessionSettings()[settingKey] = !!value;
    }

    // =========================================
    // 選択範囲の解析 / Selection Analysis
    // =========================================

    /**
     * 現在の選択を配列に写し取る
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択中のオブジェクト（選択が無ければ空配列）
     */
    function snapshotSelection(doc) {
        var selectedItems = [];
        var currentSelection = doc.selection;
        if (currentSelection && currentSelection.length) {
            for (var i = 0; i < currentSelection.length; i++) selectedItems.push(currentSelection[i]);
        }
        return selectedItems;
    }

    /**
     * 選択を指定のオブジェクトに戻す（できる範囲で）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択に戻すオブジェクト（null や空なら選択解除だけ）
     * @returns {void}
     */
    function restoreSelection(doc, selectedItems) {
        try {
            doc.selection = null;
            if (selectedItems && selectedItems.length) doc.selection = selectedItems;
        } catch (e) { /* 削除済み・ロック中のオブジェクトは選択できない / Removed or locked items cannot be selected */ }
    }

    /**
     * アイテムの左上座標を取得する
     * @param {PageItem} item - 対象のオブジェクト
     * @returns {{left: number, top: number}} 左上座標（取得できないときは 0, 0）
     */
    function getItemTopLeft(item) {
        try {
            var bounds = item.geometricBounds; /* [left, top, right, bottom] */
            return { left: bounds[0], top: bounds[1] };
        } catch (e) {
            return { left: 0, top: 0 };
        }
    }

    /**
     * 選択全体の外接バウンディングを取得する
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {{left: number, top: number, right: number, bottom: number}|null} 外接矩形（求められないときは null）
     */
    function getSelectionBounds(selectedItems) {
        if (!selectedItems || selectedItems.length === 0) return null;

        var left = 1e12, top = -1e12, right = -1e12, bottom = 1e12;
        var hasBounds = false;

        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (!selectedItem) continue;
            try {
                var bounds = selectedItem.geometricBounds; /* [left, top, right, bottom] */
                if (bounds[0] < left) left = bounds[0];
                if (bounds[1] > top) top = bounds[1];
                if (bounds[2] > right) right = bounds[2];
                if (bounds[3] < bottom) bottom = bounds[3];
                hasBounds = true;
            } catch (e) { /* 空白だけのテキストなどは境界が取れないので無視 / skip items without bounds */ }
        }

        if (!hasBounds || left > right || bottom > top) return null;
        return { left: left, top: top, right: right, bottom: bottom };
    }

    /**
     * 選択オブジェクトが横並びか縦並びかを、各オブジェクトの中心の散らばりから推定する
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {string} "horizontal" / "vertical" / "mixed" / "unknown"
     */
    function detectSelectionOrientation(selectedItems) {
        if (!selectedItems || selectedItems.length < 2) return "unknown";

        var minX = 1e12, maxX = -1e12;
        var minY = 1e12, maxY = -1e12;

        for (var i = 0; i < selectedItems.length; i++) {
            var selectedItem = selectedItems[i];
            if (!selectedItem) continue;
            try {
                var bounds = selectedItem.geometricBounds; /* [left, top, right, bottom] */
                var centerX = (bounds[0] + bounds[2]) / 2;
                var centerY = (bounds[1] + bounds[3]) / 2;
                if (centerX < minX) minX = centerX;
                if (centerX > maxX) maxX = centerX;
                if (centerY < minY) minY = centerY;
                if (centerY > maxY) maxY = centerY;
            } catch (e) { /* bounds を取得できないものは無視 / skip items without bounds */ }
        }

        if (minX > maxX || minY > maxY) return "unknown";

        var dx = Math.abs(maxX - minX);
        var dy = Math.abs(maxY - minY);
        if (dx > dy) return "horizontal";
        if (dy > dx) return "vertical";
        return "mixed";
    }

    // =========================================
    // ブレンド / Blend
    // =========================================

    /**
     * 選択オブジェクトを複製し、複製側でブレンドを作る（元の選択はそのまま残す）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function duplicateSelectionAndBlend(doc) {
        try {
            if (!doc.selection || doc.selection.length < 2) return;

            var originalItems = snapshotSelection(doc);

            /* 複製する / Duplicate the items */
            var duplicatedItems = [];
            for (var i = 0; i < originalItems.length; i++) {
                try { duplicatedItems.push(originalItems[i].duplicate()); } catch (eD) { }
            }
            if (duplicatedItems.length < 2) {
                restoreSelection(doc, originalItems);
                return;
            }

            /* 複製だけを選択する / Select only the duplicates */
            doc.selection = null;
            for (var j = 0; j < duplicatedItems.length; j++) {
                /* 非表示・ロックを引き継いだ複製は選択できない / Duplicates that inherit hidden or locked cannot be selected */
                try { duplicatedItems[j].selected = true; } catch (eS) { }
            }

            /* ブレンドのメニューコマンド名は環境で異なるため順に試す / Try known Blend menu commands (varies by locale/version) */
            try {
                app.executeMenuCommand('Make Blend');
            } catch (e1) {
                try {
                    app.executeMenuCommand('Blend Make');
                } catch (e2) {
                    try { app.executeMenuCommand('blend'); } catch (e3) { }
                }
            }

            restoreSelection(doc, originalItems);
        } catch (e) {
            /* ブレンドを作れなくても無言で続行 / Continue silently when the blend fails */
        }
    }

    // =========================================
    // カラーユーティリティ / Color Utilities
    // =========================================

    /**
     * 「なし」の色かどうかを判定する
     * @param {Color} color - 判定する色
     * @returns {boolean} null または NoColor なら true
     */
    function isNoColor(color) {
        try {
            return (color == null) || (color.typename === "NoColor");
        } catch (e) {
            return true;
        }
    }

    /**
     * 重複除去用のカラーキーを作る
     * @param {Color} color - 対象の色
     * @returns {string} 同じ色なら同じになるキー
     */
    function colorKey(color) {
        if (!color) return "null";
        var typeName = color.typename;
        try {
            if (typeName === "RGBColor") return "RGB:" + [color.red, color.green, color.blue].join(",");
            if (typeName === "CMYKColor") return "CMYK:" + [color.cyan, color.magenta, color.yellow, color.black].join(",");
            if (typeName === "GrayColor") return "Gray:" + color.gray;
            if (typeName === "SpotColor") {
                /* スポットはスポット名＋濃度 / Spot name plus tint */
                var spotName = (color.spot && color.spot.name) ? color.spot.name : "(spot)";
                return "Spot:" + spotName + ":" + color.tint;
            }
            if (typeName === "PatternColor") {
                var patternName = (color.pattern && color.pattern.name) ? color.pattern.name : "(pattern)";
                return "Pattern:" + patternName;
            }
            if (typeName === "GradientColor") {
                var gradientName = (color.gradient && color.gradient.name) ? color.gradient.name : "(gradient)";
                return "Gradient:" + gradientName;
            }
        } catch (e) { /* 名前を読めない色は種類だけのキーにする / fall back to the type name */ }
        return "Other:" + typeName;
    }

    /**
     * オブジェクトから塗り・線の色を位置付きで集める（グループ・複合パスは再帰）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {Object[]} outEntries - {left, top, color} を追加する配列
     * @returns {void}
     */
    function collectColorEntries(item, outEntries) {
        if (!item) return;

        try {
            if (item.typename === "GroupItem") {
                for (var i = 0; i < item.pageItems.length; i++) {
                    collectColorEntries(item.pageItems[i], outEntries);
                }
                return;
            }

            if (item.typename === "CompoundPathItem") {
                for (var j = 0; j < item.pathItems.length; j++) {
                    collectColorEntries(item.pathItems[j], outEntries);
                }
                return;
            }

            if (item.typename === "TextFrame") {
                var textPos = getItemTopLeft(item);
                var textFill = item.textRange.characterAttributes.fillColor;
                if (!isNoColor(textFill)) outEntries.push({ left: textPos.left, top: textPos.top, color: textFill });
                try {
                    var textStroke = item.textRange.characterAttributes.strokeColor;
                    if (!isNoColor(textStroke)) outEntries.push({ left: textPos.left, top: textPos.top, color: textStroke });
                } catch (eTS) { /* 線の色を読めないテキストは塗りだけ / fill only when the stroke is unreadable */ }
                return;
            }

            if (typeof item.filled !== "undefined" || typeof item.stroked !== "undefined") {
                var pathPos = getItemTopLeft(item);
                if (item.filled) {
                    var pathFill = item.fillColor;
                    if (!isNoColor(pathFill)) outEntries.push({ left: pathPos.left, top: pathPos.top, color: pathFill });
                }
                if (item.stroked) {
                    var pathStroke = item.strokeColor;
                    if (!isNoColor(pathStroke)) outEntries.push({ left: pathPos.left, top: pathPos.top, color: pathStroke });
                }
            }
        } catch (e) { /* 取得できないアイテムは無視 / skip unreadable items */ }
    }

    /**
     * 選択から色を「左→右、上→下」の順で集める（同じ色は最初の1つだけ）
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {Color[]} 重複を除いた色
     */
    function collectColorsFromSelection(selectedItems) {
        var entries = [];
        for (var i = 0; i < selectedItems.length; i++) {
            collectColorEntries(selectedItems[i], entries);
        }

        /* 左→右（left 昇順）、上→下（top は上ほど大きいので降順） / left ascending, then top descending */
        entries.sort(function (a, b) {
            if (a.left < b.left) return -1;
            if (a.left > b.left) return 1;
            if (a.top > b.top) return -1;
            if (a.top < b.top) return 1;
            return 0;
        });

        var colors = [];
        var seenKeys = {};
        for (var k = 0; k < entries.length; k++) {
            var entryColor = entries[k].color;
            if (isNoColor(entryColor)) continue;
            var entryKey = colorKey(entryColor);
            if (seenKeys[entryKey]) continue;
            seenKeys[entryKey] = true;
            colors.push(entryColor);
        }
        return colors;
    }

    // =========================================
    // アクション定義 / Action Definitions
    // =========================================

    /**
     * アクション定義を一時ファイルに書き出して読み込み、実行してから取り除く（失敗しても無言）
     * @param {string[]} actionLines - アクション定義（.aia）の各行
     * @param {string} actionSetName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {string} tempFileName - 一時ファイル名
     * @returns {void}
     */
    function runTempAction(actionLines, actionSetName, actionName, tempFileName) {
        /* 改行は CR / Lines are joined with CR */
        var actionCode = actionLines.join(String.fromCharCode(13));
        try {
            var tempFile = new File(Folder.temp + "/" + tempFileName);
            tempFile.open("w");
            tempFile.write(actionCode);
            tempFile.close();

            app.loadAction(tempFile);
            app.doScript(actionName, actionSetName);
            app.unloadAction(actionSetName, "");

            try { tempFile.remove(); } catch (eDel) { }
        } catch (e) {
            /* 読み込み済みなら取り除く / Unload the set if it was loaded */
            try { app.unloadAction(actionSetName, ""); } catch (e2) { }
        }
    }

    /**
     * 選択中のオブジェクトのグラデーション角度を 90° にするアクションを実行する
     * @returns {void}
     */
    function runGradientAngle90Action() {
        runTempAction([
            " /version 3",
            "/name [ 8",
            "\t6772616469656e74",
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "\t/name [ 8",
            "\t\t3930646567726565",
            "\t]",
            "\t/keyIndex 0",
            "\t/colorIndex 0",
            "\t/isOpen 1",
            "\t/eventCount 1",
            "\t/event-1 {",
            "\t\t/useRulersIn1stQuadrant 0",
            "\t\t/internalName (ai_plugin_setGradient)",
            "\t\t/localizedName [ 30",
            "\t\t\te382b0e383a9e38387e383bce382b7e383a7e383b3e38292e8a8ade5ae9a",
            "\t\t]",
            "\t\t/isOpen 1",
            "\t\t/isOn 1",
            "\t\t/hasDialog 0",
            "\t\t/parameterCount 1",
            "\t\t/parameter-1 {",
            "\t\t\t/key 1634625388",
            "\t\t\t/showInPalette 4294967295",
            "\t\t\t/type (unit real)",
            "\t\t\t/value -90.0",
            "\t\t\t/unit 591490663",
            "\t\t}",
            "\t}",
            "}",
            ""
        ], "gradient", "90degree", "temp_action.aia");
    }

    /**
     * 「新規グラフィックスタイル」をアクションで実行する
     * @returns {void}
     */
    function runGraphicStyleAction() {
        runTempAction([
            " /version 3",
            "/name [ 12",
            "\t477261706869635374796c65",
            "]",
            "/isOpen 1",
            "/actionCount 1",
            "/action-1 {",
            "\t/name [ 3",
            "\t\t6e6577",
            "\t]",
            "\t/keyIndex 0",
            "\t/colorIndex 0",
            "\t/isOpen 1",
            "\t/eventCount 1",
            "\t/event-1 {",
            "\t\t/useRulersIn1stQuadrant 0",
            "\t\t/internalName (ai_plugin_styles)",
            "\t\t/localizedName [ 30",
            "\t\t\te382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab",
            "\t\t]",
            "\t\t/isOpen 1",
            "\t\t/isOn 1",
            "\t\t/hasDialog 1",
            "\t\t/showDialog 0",
            "\t\t/parameterCount 1",
            "\t\t/parameter-1 {",
            "\t\t\t/key 1835363957",
            "\t\t\t/showInPalette 4294967295",
            "\t\t\t/type (enumerated)",
            "\t\t\t/name [ 36",
            "\t\t\t\te696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382",
            "\t\t\t\ta4e383ab",
            "\t\t\t]",
            "\t\t\t/value 1",
            "\t\t}",
            "\t}",
            "}",
            ""
        ], "GraphicStyle", "new", "temp_graphicstyle.aia");
    }

    // =========================================
    // スウォッチ操作 / Swatch Operations
    // =========================================

    /**
     * コレクションに同じ名前の項目があるかを調べる
     * @param {Object} namedCollection - doc.swatches / doc.swatchGroups / doc.gradients など getByName を持つコレクション
     * @param {string} itemName - 調べる名前
     * @returns {boolean} あれば true
     */
    function hasItemNamed(namedCollection, itemName) {
        try {
            namedCollection.getByName(itemName); /* 見つからないと例外 / throws when not found */
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 名前の衝突を避けた一意な名前を作る（"名前", "名前 1", "名前 2" …）
     * @param {string} baseName - 元の名前
     * @param {Object} namedCollection - 名前の重複を調べるコレクション
     * @returns {string} 使われていない名前
     */
    function uniqueName(baseName, namedCollection) {
        var candidateName = baseName;
        var suffix = 1;
        while (hasItemNamed(namedCollection, candidateName)) {
            candidateName = baseName + " " + suffix;
            suffix++;
        }
        return candidateName;
    }

    /**
     * 色をグローバルカラー（プロセス）に変換する
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} baseColor - 元の色
     * @param {string} spotName - 作るスポットの名前
     * @returns {Color} グローバルカラー（失敗したときは元の色）
     */
    function toGlobalProcessColor(doc, baseColor, spotName) {
        try {
            var globalSpot = doc.spots.add();
            globalSpot.name = spotName;
            globalSpot.colorType = ColorModel.PROCESS; /* グローバル（プロセス） / Global process */
            globalSpot.color = baseColor;

            var spotColor = new SpotColor();
            spotColor.spot = globalSpot;
            spotColor.tint = 100;
            return spotColor;
        } catch (e) {
            return baseColor;
        }
    }

    /**
     * 1色をスウォッチに登録する（必要ならグローバルカラーにする）
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} sourceColor - 登録する色
     * @param {string} baseName - スウォッチ名の元
     * @param {boolean} makeGlobal - グローバルカラーにするか
     * @returns {Swatch} 作ったスウォッチ
     */
    function addSwatchForColor(doc, sourceColor, baseName, makeGlobal) {
        var swatch = doc.swatches.add();
        var swatchName = uniqueName(baseName, doc.swatches);
        swatch.name = swatchName;
        swatch.color = makeGlobal ? toGlobalProcessColor(doc, sourceColor, swatchName) : sourceColor;
        try { swatch.selected = false; } catch (e) { }
        return swatch;
    }

    /**
     * 新しいスウォッチグループを作り、色をスウォッチとして登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 登録する色
     * @param {boolean} makeGlobal - グローバルカラーにするか
     * @returns {Swatch[]} 作ったスウォッチ（colors と同じ順）
     */
    function registerColorSwatches(doc, colors, makeGlobal) {
        var groupName = uniqueName("AutoGradient", doc.swatchGroups);
        var swatchGroup = doc.swatchGroups.add();
        swatchGroup.name = groupName;

        var createdSwatches = [];
        for (var i = 0; i < colors.length; i++) {
            var createdSwatch = addSwatchForColor(doc, colors[i], "AutoColor", makeGlobal);
            createdSwatches.push(createdSwatch);
            try { swatchGroup.addSwatch(createdSwatch); } catch (eAdd) { }
        }
        return createdSwatches;
    }

    // =========================================
    // レイヤー操作 / Layer Helpers
    // =========================================

    /**
     * 描画できるレイヤー（ロックも非表示もされていないもの）を返す。アクティブレイヤーを優先する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer|null} 描画できるレイヤー（無ければ null）
     */
    function getUnlockedVisibleLayer(doc) {
        try {
            if (doc.activeLayer && !doc.activeLayer.locked && doc.activeLayer.visible) return doc.activeLayer;
        } catch (e) { }
        for (var i = 0; i < doc.layers.length; i++) {
            try {
                if (!doc.layers[i].locked && doc.layers[i].visible) return doc.layers[i];
            } catch (e2) { }
        }
        return null;
    }

    // =========================================
    // グラフィックスタイル登録 / Graphic Style Registration
    // =========================================

    /**
     * 選択中のオブジェクトの見た目を、1つずつ新規グラフィックスタイルとして登録する（名前は既定のまま）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function registerGraphicStyleFromSelected(doc) {
        try {
            if (!doc.graphicStyles) return;

            var selectedItems = snapshotSelection(doc);
            for (var i = 0; i < selectedItems.length; i++) {
                /* 1つだけ選択してアクションで登録 / Select just this item and register it via the action */
                doc.selection = null;
                try {
                    doc.selection = [selectedItems[i]];
                } catch (eSetSel) {
                    try { selectedItems[i].selected = true; } catch (eSel2) { }
                }
                runGraphicStyleAction();
            }
        } catch (e) { }
    }

    // =========================================
    // ダイアログ / Options Dialog
    // =========================================

    /**
     * 前回の値（無ければ既定値）からオプションを作る
     * @returns {Object} makeGlobal / makeGradient / makeRect / useSelectionSize / registerGraphicStyle
     */
    function loadGradientOptions() {
        return {
            makeGlobal: loadBool('makeGlobal', true),
            makeGradient: loadBool('makeGradient', true),
            makeRect: loadBool('makeRect', true),
            useSelectionSize: loadBool('useSelectionSize', true),
            registerGraphicStyle: loadBool('registerGraphicStyle', true)
        };
    }

    /**
     * チェックボックスを並べるパネルを追加する
     * @param {Window} parentWindow - 追加先のダイアログ
     * @param {string} titlePath - パネル名のラベルのパス
     * @returns {Panel} 追加したパネル
     */
    function addOptionPanel(parentWindow, titlePath) {
        var optionPanel = parentWindow.add('panel', undefined, getLabel(titlePath));
        optionPanel.orientation = 'column';
        optionPanel.alignChildren = ['fill', 'top'];
        optionPanel.margins = PANEL_MARGINS;
        return optionPanel;
    }

    /**
     * tooltip 付きのチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} textPath - 項目名のラベルのパス
     * @param {string} tipPath - tooltip のラベルのパス
     * @param {boolean} initialValue - 初期値
     * @returns {Checkbox} 追加したチェックボックス
     */
    function addOptionCheckbox(parentPanel, textPath, tipPath, initialValue) {
        var optionCheckbox = parentPanel.add('checkbox', undefined, getLabel(textPath));
        optionCheckbox.helpTip = getLabel(tipPath);
        optionCheckbox.value = initialValue;
        return optionCheckbox;
    }

    /**
     * オプションダイアログを表示する。OK のときは選んだ値を gradientOptions に書き込む
     * @param {Object} gradientOptions - 初期値（loadGradientOptions() の戻り値）
     * @returns {boolean} OK なら true、キャンセルなら false
     */
    function showOptionsDialog(gradientOptions) {
        var optionsDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        optionsDialog.orientation = 'column';
        optionsDialog.alignChildren = ['fill', 'top'];

        var colorPanel = addOptionPanel(optionsDialog, 'panel.color');
        var globalColorCheckbox = addOptionCheckbox(colorPanel, 'checkbox.globalColor', 'tooltip.globalColor', gradientOptions.makeGlobal);
        var gradientCheckbox = addOptionCheckbox(colorPanel, 'checkbox.createGradient', 'tooltip.createGradient', gradientOptions.makeGradient);

        var rectPanel = addOptionPanel(optionsDialog, 'panel.rect');
        var rectCheckbox = addOptionCheckbox(rectPanel, 'checkbox.createRect', 'tooltip.createRect', gradientOptions.makeRect);
        var selectionSizeCheckbox = addOptionCheckbox(rectPanel, 'checkbox.useSelectionSize', 'tooltip.useSelectionSize', gradientOptions.useSelectionSize);
        var graphicStyleCheckbox = addOptionCheckbox(rectPanel, 'checkbox.registerGraphicStyle', 'tooltip.registerGraphicStyle', gradientOptions.registerGraphicStyle);

        /**
         * チェック状態の連動（グラデーションを作らないときは長方形・スタイルを OFF にしてディム）
         * @returns {void}
         */
        function syncEnable() {
            var gradientOn = gradientCheckbox.value;
            if (!gradientOn) {
                rectCheckbox.value = false;
                selectionSizeCheckbox.value = false;
                graphicStyleCheckbox.value = false;
            }
            rectCheckbox.enabled = gradientOn;
            selectionSizeCheckbox.enabled = gradientOn && rectCheckbox.value;
            /* 長方形を出力しなくても一時長方形でスタイルは作れる / Graphic style can be created from a temporary rectangle */
            graphicStyleCheckbox.enabled = gradientOn;
        }
        gradientCheckbox.onClick = syncEnable;
        rectCheckbox.onClick = syncEnable;
        syncEnable();

        var btnRowGroup = optionsDialog.add('group');
        btnRowGroup.alignment = 'right';
        btnRowGroup.add('button', undefined, getLabel('button.cancel'), { name: 'cancel' });
        btnRowGroup.add('button', undefined, getLabel('button.ok'), { name: 'ok' });

        /**
         * チェック状態をセッション設定に保存する
         * @returns {void}
         */
        function persistFromUI() {
            saveBool('makeGlobal', globalColorCheckbox.value);
            saveBool('makeGradient', gradientCheckbox.value);
            saveBool('makeRect', rectCheckbox.value);
            saveBool('useSelectionSize', selectionSizeCheckbox.value);
            saveBool('registerGraphicStyle', graphicStyleCheckbox.value);
        }
        optionsDialog.onClose = persistFromUI;

        if (optionsDialog.show() !== 1) return false;

        gradientOptions.makeGlobal = !!globalColorCheckbox.value;
        gradientOptions.makeGradient = !!gradientCheckbox.value;
        gradientOptions.makeRect = !!rectCheckbox.value;
        gradientOptions.useSelectionSize = !!selectionSizeCheckbox.value;
        gradientOptions.registerGraphicStyle = !!graphicStyleCheckbox.value;
        persistFromUI();
        return true;
    }

    // =========================================
    // グラデーション生成 / Gradient Construction
    // =========================================

    /**
     * 色を等間隔に並べた線形グラデーションを作る（ストップには登録したスウォッチの色を優先して使う）
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 並べる色
     * @param {Swatch[]} createdSwatches - colors と同じ順に登録したスウォッチ
     * @returns {Gradient} 作ったグラデーション
     */
    function buildNormalGradient(doc, colors, createdSwatches) {
        var newGradient = doc.gradients.add();
        newGradient.type = GradientType.LINEAR;

        /* ストップ数を色数に合わせる / Match the stop count to the color count */
        while (newGradient.gradientStops.length < colors.length) {
            newGradient.gradientStops.add();
        }
        while (newGradient.gradientStops.length > colors.length) {
            newGradient.gradientStops[newGradient.gradientStops.length - 1].remove();
        }

        for (var i = 0; i < colors.length; i++) {
            var gradientStop = newGradient.gradientStops[i];
            gradientStop.rampPoint = (i / (colors.length - 1)) * 100;

            /* スウォッチ登録時に作ったグローバルカラーを優先 / Prefer the global color made for the swatch */
            try {
                if (createdSwatches[i] && createdSwatches[i].color) {
                    gradientStop.color = createdSwatches[i].color;
                } else {
                    gradientStop.color = colors[i];
                }
            } catch (eStopColor) {
                try { gradientStop.color = colors[i]; } catch (e2) { }
            }

            gradientStop.midPoint = 50;
            gradientStop.opacity = 100;
        }

        newGradient.name = uniqueName("New Gradient", doc.gradients);
        return newGradient;
    }

    // =========================================
    // 長方形配置 / Rectangle Placement
    // =========================================

    /**
     * グラデーションを適用する長方形の位置とサイズを決める
     * 横並びなら選択の左端にそろえて長方形1個分下に、縦並びなら上端にそろえて長方形1個分右に置く。それ以外はビューの中央
     * @param {number} viewCenterX - ビュー中心の X
     * @param {number} viewCenterY - ビュー中心の Y
     * @param {boolean} useSelectionSize - 選択の外接サイズを使うか
     * @param {Object|null} selectionBounds - 選択の外接矩形 {left, top, right, bottom}
     * @param {string} selectionOrientation - detectSelectionOrientation() の結果
     * @returns {{top: number, left: number, width: number, height: number}} 長方形の位置とサイズ
     */
    function computeRectFrame(viewCenterX, viewCenterY, useSelectionSize, selectionBounds, selectionOrientation) {
        var rectWidth = DEFAULT_RECT_SIZE;
        var rectHeight = DEFAULT_RECT_SIZE;
        if (useSelectionSize && selectionBounds) {
            rectWidth = Math.abs(selectionBounds.right - selectionBounds.left);
            rectHeight = Math.abs(selectionBounds.top - selectionBounds.bottom);
        }
        if (!rectWidth || rectWidth <= 0) rectWidth = DEFAULT_RECT_SIZE;
        if (!rectHeight || rectHeight <= 0) rectHeight = DEFAULT_RECT_SIZE;

        var rectTop = viewCenterY + rectHeight / 2;
        var rectLeft = viewCenterX - rectWidth / 2;
        if (selectionBounds) {
            if (selectionOrientation === "horizontal") {
                rectLeft = selectionBounds.left;
                rectTop = selectionBounds.bottom - rectHeight;
            } else if (selectionOrientation === "vertical") {
                rectLeft = selectionBounds.right + rectWidth;
                rectTop = selectionBounds.top;
            }
        }
        return { top: rectTop, left: rectLeft, width: rectWidth, height: rectHeight };
    }

    /**
     * 一時レイヤーを作ってアクティブにする（長方形を出力せずにスタイルだけ登録するとき用）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer|null} 作ったレイヤー（失敗したときは null）
     */
    function createTempLayer(doc) {
        try {
            var tempLayer = doc.layers.add();
            tempLayer.name = "__TempGraphicStyle";
            tempLayer.visible = true;
            tempLayer.locked = false;
            try { doc.activeLayer = tempLayer; } catch (eSetAL) { }
            return tempLayer;
        } catch (e) {
            return null;
        }
    }

    /**
     * 長方形を作ってグラデーションを適用し、必要ならグラフィックスタイルに登録する
     * 長方形を出力しない設定では、一時レイヤーの一時長方形でスタイルを登録してから片付ける
     * @param {Document} doc - 対象ドキュメント
     * @param {Gradient} newGradient - 適用するグラデーション
     * @param {Object} gradientOptions - ダイアログのオプション
     * @param {Object|null} selectionBounds - 元の選択の外接矩形
     * @param {string} selectionOrientation - 元の選択の並び
     * @returns {void}
     */
    function createGradientRect(doc, newGradient, gradientOptions, selectionBounds, selectionOrientation) {
        var targetLayer = getUnlockedVisibleLayer(doc);
        if (!targetLayer) return;

        var tempRectForStyle = (!gradientOptions.makeRect && gradientOptions.registerGraphicStyle);
        var previousActiveLayer = null;
        try { previousActiveLayer = doc.activeLayer; } catch (eAL) { }
        var previousSelection = snapshotSelection(doc);

        var tempLayer = tempRectForStyle ? createTempLayer(doc) : null;
        var drawLayer = tempLayer || targetLayer;

        var viewCenter = doc.activeView.centerPoint;
        var rectFrame = computeRectFrame(viewCenter[0], viewCenter[1], gradientOptions.useSelectionSize, selectionBounds, selectionOrientation);
        var gradientRect = drawLayer.pathItems.rectangle(rectFrame.top, rectFrame.left, rectFrame.width, rectFrame.height);

        /* 作った長方形を選択し、元の選択が縦並びならグラデーション角度を 90° に / Select it; rotate the gradient when vertical */
        try {
            doc.selection = null;
            gradientRect.selected = true;
            if (selectionOrientation === "vertical") runGradientAngle90Action();
        } catch (eSelRect) { }

        gradientRect.stroked = false;
        gradientRect.filled = true;
        var gradientFill = new GradientColor();
        gradientFill.gradient = newGradient;
        gradientRect.fillColor = gradientFill;

        if (gradientOptions.registerGraphicStyle) {
            try {
                doc.selection = null;
                gradientRect.selected = true;
                registerGraphicStyleFromSelected(doc);
            } catch (eGS) { }
        }

        /* 一時長方形だった場合の後始末 / Clean up the temporary rectangle */
        if (tempRectForStyle) {
            try { gradientRect.remove(); } catch (eRm) { }
            if (tempLayer) {
                try { tempLayer.remove(); } catch (eLayRm) { }
            }
            try {
                if (previousActiveLayer) doc.activeLayer = previousActiveLayer;
            } catch (eRestoreAL) { }
            restoreSelection(doc, previousSelection);
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から色を集め、ブレンド・スウォッチ・グラデーション・長方形を作る
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) return;

        /* 元の選択の並びと外接矩形（長方形の配置とサイズに使う） / Orientation and bounds of the original selection */
        var selectionOrientation = detectSelectionOrientation(doc.selection);
        var selectionBounds = getSelectionBounds(doc.selection);

        /* 複製したオブジェクトでブレンドを作る（元の選択は維持） / Blend the duplicates, keeping the original selection */
        duplicateSelectionAndBlend(doc);

        var gradientOptions = loadGradientOptions();
        try {
            if (!showOptionsDialog(gradientOptions)) return; /* キャンセル時は無言で終了 / silent on cancel */
        } catch (eDlg) {
            /* ダイアログを作れなくても既定値のまま無言で続行 / Continue with the defaults if the dialog fails */
        }

        var colors = collectColorsFromSelection(doc.selection);
        if (colors.length < 2) return;

        try {
            var createdSwatches = registerColorSwatches(doc, colors, gradientOptions.makeGlobal);

            /* 以降は選択に依存しない / Nothing below depends on the selection */
            doc.selection = null;

            var newGradient = gradientOptions.makeGradient ? buildNormalGradient(doc, colors, createdSwatches) : null;

            /* 長方形・グラフィックスタイル / Rectangle and Graphic Style */
            if ((gradientOptions.makeRect || gradientOptions.registerGraphicStyle) && newGradient) {
                try { createGradientRect(doc, newGradient, gradientOptions, selectionBounds, selectionOrientation); } catch (eRect) { }
            }

            /* 最後に追加されたスウォッチ（= 作ったグラデーション）を選択 / Select the last swatch (the new gradient) */
            if (newGradient) {
                try {
                    var lastIndex = doc.swatches.length - 1;
                    if (lastIndex >= 0) doc.swatches[lastIndex].selected = true;
                } catch (eSw) { }
            }
        } catch (e) {
            /* 無言（エラー通知しない） / silent */
        }
    }

    main();

})();
