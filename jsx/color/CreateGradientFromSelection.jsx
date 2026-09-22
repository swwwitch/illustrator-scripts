#target illustrator
#targetengine "CreateGradientFromSelectionEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの塗り／線カラーを配置順（左→右、上→下）で抽出し、スウォッチグループに登録してグラデーションを自動生成します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGradientFromSelection.md

### Overview

Extracts the fill and stroke colors of the selection in layout order (left to right, top to bottom), registers them as a swatch group, and builds a gradient from them.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGradientFromSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CreateGradientFromSelection";  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.9.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-05-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CreateGradientFromSelection.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CreateGradientFromSelection.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* セパレートグラデーションで許可する最大色数 / Max colors allowed for Separate gradients */
    var SEPARATE_MAX_COLORS = 6;

    /* 長方形サイズの既定値（pt）/ Default rectangle size in points */
    var DEFAULT_RECT_SIZE = 100;

    /* スウォッチ由来時の固定サイズ（pt）/ Fixed size when invoked from swatches */
    var SWATCH_RECT_WIDTH = 200;
    var SWATCH_RECT_HEIGHT = 100;

    /* 新規スウォッチグループの既定名 / Default name for the new swatch group */
    var SWATCH_GROUP_BASE_NAME = "AutoGradient";
    var SWATCH_BASE_NAME = "AutoColor";
    var GRADIENT_BASE_NAME = "New Gradient";

    /* ダイアログのデフォルトオプション / Default dialog values */
    var DEFAULT_OPTIONS = {
        makeGlobal: true,
        makeGradient: true,
        makeRect: true,
        useSelectionSize: true,
        registerGraphicStyle: true,
        separateGradient: false
    };

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

    var LABELS = {
        dialog: {
            title: { ja: "グラデーション作成", en: "Create Gradient" }
        },
        panel: {
            color: { ja: "カラー", en: "Colors" },
            rect: { ja: "長方形", en: "Rectangle" }
        },
        checkbox: {
            globalColor: { ja: "グローバルカラー化", en: "Make Global colors" },
            createGradient: { ja: "グラデーションを作成", en: "Create gradient" },
            createRect: { ja: "長方形を作成してグラデーションを適用", en: "Create rectangle and apply gradient" },
            useSelectionSize: { ja: "選択オブジェクトのサイズに合わせる", en: "Match selection size" },
            registerGraphicStyle: { ja: "グラフィックスタイルとして登録", en: "Save as Graphic Style" }
        },
        radio: {
            normal: { ja: "通常", en: "Smooth" },
            separate: { ja: "セパレート", en: "Segmented" }
        },
        button: {
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        tooltip: {
            globalColor: {
                ja: "スウォッチをグローバルカラー（プロセス）として登録し、後から一括で色を変更可能にします。",
                en: "Register swatches as Global Process colors so they can be edited together later."
            },
            createGradient: {
                ja: "抽出した色を並べた線形グラデーションを作ります。OFF のときはスウォッチの登録だけを行います。",
                en: "Builds a linear gradient from the extracted colors. When off, only the swatches are registered."
            },
            normal: {
                ja: "色を等間隔に置き、なめらかにつなぐグラデーションにします。",
                en: "Places the colors evenly and blends them smoothly."
            },
            separate: {
                ja: "色の境界をくっきり分割するグラデーション（最大 6 色）。選択が 7 つ以上のときは使えません。",
                en: "Hard-edged segmented gradient (up to 6 colors). Disabled when 7 or more items are selected."
            },
            createRect: {
                ja: "作ったグラデーションで塗った長方形を描きます。横並びの選択なら下、縦並びなら右に、長方形1個分の間隔をあけて置きます。",
                en: "Draws a rectangle filled with the new gradient, leaving a one-rectangle gap below a horizontal selection or to the right of a vertical one."
            },
            useSelectionSize: {
                ja: "選択オブジェクトの外接サイズに合わせて長方形を作成します。",
                en: "Use the bounding size of the selection for the rectangle."
            },
            registerGraphicStyle: {
                ja: "作成した長方形の見た目をグラフィックスタイルに登録します。長方形作成 OFF のときは一時長方形で登録します。",
                en: "Register the rectangle's appearance as a Graphic Style. When rectangle output is off, a temporary rectangle is used."
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
            labelNode = labelNode && labelNode[labelPathKeys[i]];
        }
        if (labelNode && labelNode[uiLang]) return labelNode[uiLang];
        if (labelNode && labelNode.en) return labelNode.en;
        return labelPath;
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
        if (typeof sessionSettings[settingKey] === 'boolean') return sessionSettings[settingKey];
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
            try {
                var bounds = selectedItems[i].geometricBounds;
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
     * 選択オブジェクトが横並びか縦並びかを、各オブジェクトの中心の散らばりから判定する
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {string} "horizontal" / "vertical" / "mixed" / "unknown"
     */
    function detectSelectionOrientation(selectedItems) {
        if (!selectedItems || selectedItems.length < 2) return "unknown";

        var minX = 1e12, maxX = -1e12;
        var minY = 1e12, maxY = -1e12;

        for (var i = 0; i < selectedItems.length; i++) {
            try {
                var bounds = selectedItems[i].geometricBounds;
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
                } catch (e) { /* 線の色を読めないテキストは塗りだけ / fill only when the stroke is unreadable */ }
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
     * 選択から重複を除いた色を、長辺方向（横なら左→右／縦なら上→下）の順で返す
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @param {string} selectionOrientation - detectSelectionOrientation() の結果
     * @returns {Color[]} 重複を除いた色
     */
    function collectColorsFromSelection(selectedItems, selectionOrientation) {
        var entries = [];
        for (var i = 0; i < selectedItems.length; i++) {
            collectColorEntries(selectedItems[i], entries);
        }

        /* 縦並び（dy > dx）なら上→下を優先キーに / Use top→bottom as primary key when vertical */
        var verticalPrimary = (selectionOrientation === "vertical");

        entries.sort(function (a, b) {
            if (verticalPrimary) {
                if (a.top > b.top) return -1;
                if (a.top < b.top) return 1;
                if (a.left < b.left) return -1;
                if (a.left > b.left) return 1;
                return 0;
            }
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
     * 一時アクションを読み込んで実行し、後始末をする（失敗しても無言）
     * @param {string} actionCode - アクション定義（.aia）のテキスト
     * @param {string} actionSetName - アクションセット名
     * @param {string} actionName - アクション名
     * @returns {void}
     */
    function runTempAction(actionCode, actionSetName, actionName) {
        var tempFile = new File(Folder.temp + "/temp_action_" + actionSetName + ".aia");
        try {
            tempFile.open("w");
            tempFile.write(actionCode);
            tempFile.close();

            app.loadAction(tempFile);
            app.doScript(actionName, actionSetName);
            app.unloadAction(actionSetName, "");

            try { tempFile.remove(); } catch (e) { /* 無視 / ignore */ }
        } catch (e) {
            /* 読み込み済みなら取り除く / Unload the set if it was loaded */
            try { app.unloadAction(actionSetName, ""); } catch (e2) { /* 無視 / ignore */ }
        }
    }

    /**
     * 選択中のオブジェクトのグラデーション角度を 90° にするアクションを実行する
     * @returns {void}
     */
    function runGradientAngle90Action() {
        var CR = String.fromCharCode(13);
        var actionCode = [
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
        ].join(CR);
        runTempAction(actionCode, "gradient", "90degree");
    }

    /**
     * 「新規グラフィックスタイル」をアクションで実行する
     * @returns {void}
     */
    function runGraphicStyleAction() {
        var CR = String.fromCharCode(13);
        var actionCode = [
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
        ].join(CR);
        runTempAction(actionCode, "GraphicStyle", "new");
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
            globalSpot.colorType = ColorModel.PROCESS;
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
        try { swatch.selected = false; } catch (e) { /* 無視 / ignore */ }
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
        var groupName = uniqueName(SWATCH_GROUP_BASE_NAME, doc.swatchGroups);
        var swatchGroup = doc.swatchGroups.add();
        swatchGroup.name = groupName;

        var createdSwatches = [];
        for (var i = 0; i < colors.length; i++) {
            var createdSwatch = addSwatchForColor(doc, colors[i], SWATCH_BASE_NAME, makeGlobal);
            createdSwatches.push(createdSwatch);
            try { swatchGroup.addSwatch(createdSwatch); } catch (e) { /* 無視 / ignore */ }
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
        var activeLayer = doc.activeLayer;
        if (activeLayer && !activeLayer.locked && activeLayer.visible) return activeLayer;
        for (var i = 0; i < doc.layers.length; i++) {
            var layer = doc.layers[i];
            if (!layer.locked && layer.visible) return layer;
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
        if (!doc.graphicStyles) return;

        var selectedItems = snapshotSelection(doc);
        for (var i = 0; i < selectedItems.length; i++) {
            /* 1つだけ選択してアクションで登録 / Select just this item and register it via the action */
            doc.selection = null;
            try { doc.selection = [selectedItems[i]]; }
            catch (e) { try { selectedItems[i].selected = true; } catch (e2) { /* 無視 / ignore */ } }

            runGraphicStyleAction();
        }
    }

    // =========================================
    // 入力収集 / Input Collection
    // =========================================

    /**
     * 選択オブジェクト、または（選択が無ければ）選択中のスウォッチから、色と関連情報を集める
     * @param {Document} doc - 対象ドキュメント
     * @returns {{colors: Color[], fromSwatches: boolean, selectionBounds: Object, selectionOrientation: string, itemCount: number}} 集めた色と関連情報
     */
    function collectInputColors(doc) {
        var colorInput = {
            colors: [],
            fromSwatches: false,
            selectionBounds: null,
            selectionOrientation: "unknown",
            itemCount: 0
        };

        if (doc.selection && doc.selection.length > 0) {
            colorInput.selectionOrientation = detectSelectionOrientation(doc.selection);
            colorInput.colors = collectColorsFromSelection(doc.selection, colorInput.selectionOrientation);
            colorInput.selectionBounds = getSelectionBounds(doc.selection);
            colorInput.itemCount = doc.selection.length;
            return colorInput;
        }

        var selectedSwatches = null;
        try { selectedSwatches = doc.swatches.getSelected(); } catch (e) { selectedSwatches = null; }
        if (!selectedSwatches || selectedSwatches.length < 2) return colorInput;

        colorInput.fromSwatches = true;
        for (var i = 0; i < selectedSwatches.length; i++) {
            try {
                var swatchColor = selectedSwatches[i].color;
                if (swatchColor && swatchColor.typename !== "NoColor") colorInput.colors.push(swatchColor);
            } catch (e) { /* 無視 / ignore */ }
        }
        colorInput.itemCount = colorInput.colors.length;
        return colorInput;
    }

    // =========================================
    // ダイアログ / Options Dialog
    // =========================================

    /**
     * 前回の値（無ければ既定値）からダイアログの初期値を作る
     * @param {boolean} disallowSeparate - セパレートを選べないとき true（常に OFF にする）
     * @returns {Object} makeGlobal / makeGradient / makeRect / useSelectionSize / registerGraphicStyle / separateGradient
     */
    function loadDialogOptions(disallowSeparate) {
        return {
            makeGlobal: loadBool('makeGlobal', DEFAULT_OPTIONS.makeGlobal),
            makeGradient: loadBool('makeGradient', DEFAULT_OPTIONS.makeGradient),
            makeRect: loadBool('makeRect', DEFAULT_OPTIONS.makeRect),
            useSelectionSize: loadBool('useSelectionSize', DEFAULT_OPTIONS.useSelectionSize),
            registerGraphicStyle: loadBool('registerGraphicStyle', DEFAULT_OPTIONS.registerGraphicStyle),
            separateGradient: (disallowSeparate ? false : loadBool('separateGradient', DEFAULT_OPTIONS.separateGradient))
        };
    }

    /**
     * 縦並びのパネルを追加する
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
     * オプションダイアログを表示し、確定値を返す
     * @param {boolean} disallowSeparate - セパレートを選べないとき true
     * @param {boolean} fromSwatches - 色をスウォッチから集めたとき true（［選択オブジェクトのサイズに合わせる］を使えない）
     * @returns {Object|null} 確定したオプション（キャンセル時は null）
     */
    function showOptionsDialog(disallowSeparate, fromSwatches) {
        var initialOptions = loadDialogOptions(disallowSeparate);

        var optionsDialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        optionsDialog.orientation = 'column';
        optionsDialog.alignChildren = ['fill', 'top'];

        /* カラー関連パネル / Color-related panel */
        var colorPanel = addOptionPanel(optionsDialog, 'panel.color');

        var globalColorCheckbox = colorPanel.add('checkbox', undefined, getLabel('checkbox.globalColor'));
        globalColorCheckbox.value = initialOptions.makeGlobal;
        globalColorCheckbox.helpTip = getLabel('tooltip.globalColor');

        var gradientCheckbox = colorPanel.add('checkbox', undefined, getLabel('checkbox.createGradient'));
        gradientCheckbox.value = initialOptions.makeGradient;
        gradientCheckbox.helpTip = getLabel('tooltip.createGradient');

        var gradientTypeGroup = colorPanel.add('group');
        gradientTypeGroup.orientation = 'row';
        gradientTypeGroup.alignChildren = ['left', 'center'];

        var normalRadio = gradientTypeGroup.add('radiobutton', undefined, getLabel('radio.normal'));
        normalRadio.helpTip = getLabel('tooltip.normal');
        var separateRadio = gradientTypeGroup.add('radiobutton', undefined, getLabel('radio.separate'));
        separateRadio.helpTip = getLabel('tooltip.separate');
        separateRadio.value = !!initialOptions.separateGradient;
        normalRadio.value = !separateRadio.value;

        /* 長方形パネル / Rectangle panel */
        var rectPanel = addOptionPanel(optionsDialog, 'panel.rect');

        var rectCheckbox = rectPanel.add('checkbox', undefined, getLabel('checkbox.createRect'));
        rectCheckbox.value = initialOptions.makeRect;
        rectCheckbox.helpTip = getLabel('tooltip.createRect');

        var selectionSizeCheckbox = rectPanel.add('checkbox', undefined, getLabel('checkbox.useSelectionSize'));
        selectionSizeCheckbox.value = initialOptions.useSelectionSize;
        selectionSizeCheckbox.helpTip = getLabel('tooltip.useSelectionSize');

        var graphicStyleCheckbox = rectPanel.add('checkbox', undefined, getLabel('checkbox.registerGraphicStyle'));
        graphicStyleCheckbox.value = initialOptions.registerGraphicStyle;
        graphicStyleCheckbox.helpTip = getLabel('tooltip.registerGraphicStyle');

        /**
         * チェック状態の連動（グラデーションを作らないときは長方形・スタイル・種類を OFF にしてディム）
         * @returns {void}
         */
        function syncEnable() {
            rectCheckbox.enabled = gradientCheckbox.value;
            selectionSizeCheckbox.enabled = gradientCheckbox.value && rectCheckbox.value && !fromSwatches;
            graphicStyleCheckbox.enabled = gradientCheckbox.value;

            normalRadio.enabled = gradientCheckbox.value;
            separateRadio.enabled = gradientCheckbox.value && !disallowSeparate;
            gradientTypeGroup.enabled = gradientCheckbox.value;

            if (!gradientCheckbox.value) {
                rectCheckbox.value = false;
                selectionSizeCheckbox.value = false;
                graphicStyleCheckbox.value = false;
            }
            if (fromSwatches) selectionSizeCheckbox.value = false;
            if (!gradientCheckbox.value || disallowSeparate) {
                separateRadio.value = false;
                normalRadio.value = true;
            }
        }
        gradientCheckbox.onClick = syncEnable;
        rectCheckbox.onClick = syncEnable;
        syncEnable();

        /* OK／キャンセル / OK and Cancel */
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
            saveBool('separateGradient', (disallowSeparate ? false : separateRadio.value));
        }
        optionsDialog.onClose = persistFromUI;

        if (optionsDialog.show() !== 1) return null;

        return {
            makeGlobal: !!globalColorCheckbox.value,
            makeGradient: !!gradientCheckbox.value,
            makeRect: !!rectCheckbox.value,
            useSelectionSize: !!selectionSizeCheckbox.value,
            registerGraphicStyle: !!graphicStyleCheckbox.value,
            separateGradient: (disallowSeparate ? false : !!separateRadio.value)
        };
    }

    // =========================================
    // グラデーション生成 / Gradient Construction
    // =========================================

    /**
     * ストップに使う色を返す（作成済みスウォッチの色を優先し、無ければ元の色）
     * @param {Swatch[]} createdSwatches - 登録したスウォッチ
     * @param {Color[]} colors - 元の色
     * @param {number} index - 色の番号
     * @returns {Color} ストップに使う色
     */
    function pickStopColor(createdSwatches, colors, index) {
        try {
            if (createdSwatches && createdSwatches[index] && createdSwatches[index].color) {
                return createdSwatches[index].color;
            }
        } catch (e) { /* 無視 / ignore */ }
        return colors[index];
    }

    /**
     * グラデーションのストップ数を targetCount に合わせる
     * @param {Gradient} gradient - 対象のグラデーション
     * @param {number} targetCount - ストップ数
     * @returns {void}
     */
    function resizeGradientStops(gradient, targetCount) {
        while (gradient.gradientStops.length < targetCount) gradient.gradientStops.add();
        while (gradient.gradientStops.length > targetCount) {
            gradient.gradientStops[gradient.gradientStops.length - 1].remove();
        }
    }

    /**
     * 色を等間隔に並べた通常（スムーズ）グラデーションを作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 並べる色
     * @param {Swatch[]} createdSwatches - colors と同じ順に登録したスウォッチ
     * @returns {Gradient} 作ったグラデーション
     */
    function buildNormalGradient(doc, colors, createdSwatches) {
        var gradient = doc.gradients.add();
        gradient.type = GradientType.LINEAR;
        resizeGradientStops(gradient, colors.length);

        for (var i = 0; i < colors.length; i++) {
            var stop = gradient.gradientStops[i];
            stop.rampPoint = (i / (colors.length - 1)) * 100;
            try { stop.color = pickStopColor(createdSwatches, colors, i); }
            catch (e) { try { stop.color = colors[i]; } catch (e2) { /* 無視 / ignore */ } }
            stop.midPoint = 50;
            stop.opacity = 100;
        }

        gradient.name = uniqueName(GRADIENT_BASE_NAME, doc.gradients);
        return gradient;
    }

    /**
     * セパレート（境界がくっきり）グラデーションを作る（2〜SEPARATE_MAX_COLORS 色）
     * 境界ごとに左右 0.01% の位置へ同じ色の組を置き、色の帯を作る
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 並べる色
     * @param {Swatch[]} createdSwatches - colors と同じ順に登録したスウォッチ
     * @returns {Gradient} 作ったグラデーション
     */
    function buildSeparateGradient(doc, colors, createdSwatches) {
        var colorCount = colors.length;
        var epsilon = 0.01;
        var bandWidth = 100 / colorCount;
        if (colorCount === 3 || colorCount === 6) bandWidth = Math.round(bandWidth * 10) / 10;

        var stopPoints = [0];
        var stopColors = [pickStopColor(createdSwatches, colors, 0)];

        for (var k = 1; k <= colorCount - 1; k++) {
            var boundary = bandWidth * k;
            if (colorCount === 3 || colorCount === 6) boundary = Math.round(boundary * 10) / 10;
            var leftPoint = Math.max(0, boundary - epsilon);
            var rightPoint = Math.min(100, boundary + epsilon);

            stopPoints.push(leftPoint);
            stopColors.push(pickStopColor(createdSwatches, colors, k - 1));
            stopPoints.push(rightPoint);
            stopColors.push(pickStopColor(createdSwatches, colors, k));
        }

        stopPoints.push(100);
        stopColors.push(pickStopColor(createdSwatches, colors, colorCount - 1));

        var gradient = doc.gradients.add();
        gradient.type = GradientType.LINEAR;
        resizeGradientStops(gradient, stopPoints.length);

        for (var i = 0; i < stopPoints.length; i++) {
            var stop = gradient.gradientStops[i];
            stop.rampPoint = stopPoints[i];
            try { stop.color = stopColors[i]; }
            catch (e) {
                try { stop.color = colors[Math.min(colors.length - 1, Math.max(0, Math.floor(i / 2)))]; } catch (e2) { /* 無視 / ignore */ }
            }
            stop.midPoint = 50;
            stop.opacity = 100;
        }
        return gradient;
    }

    // =========================================
    // 長方形配置 / Rectangle Placement
    // =========================================

    /**
     * 長方形のサイズを決める
     * @param {Object} gradientOptions - ダイアログのオプション
     * @param {Object} colorInput - collectInputColors() の結果
     * @returns {{width: number, height: number}} 長方形のサイズ
     */
    function computeRectSize(gradientOptions, colorInput) {
        if (colorInput.fromSwatches) return { width: SWATCH_RECT_WIDTH, height: SWATCH_RECT_HEIGHT };
        if (gradientOptions.useSelectionSize && colorInput.selectionBounds) {
            var boundsWidth = Math.abs(colorInput.selectionBounds.right - colorInput.selectionBounds.left);
            var boundsHeight = Math.abs(colorInput.selectionBounds.top - colorInput.selectionBounds.bottom);
            if (boundsWidth > 0 && boundsHeight > 0) return { width: boundsWidth, height: boundsHeight };
        }
        return { width: DEFAULT_RECT_SIZE, height: DEFAULT_RECT_SIZE };
    }

    /**
     * 長方形の配置（左上座標）を決める
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} colorInput - collectInputColors() の結果
     * @param {{width: number, height: number}} rectSize - 長方形のサイズ
     * @returns {{left: number, top: number}} 長方形の左上
     */
    function computeRectPosition(doc, colorInput, rectSize) {
        var viewCenter = doc.activeView.centerPoint;
        var left = viewCenter[0] - rectSize.width / 2;
        var top = viewCenter[1] + rectSize.height / 2;

        if (!colorInput.fromSwatches && colorInput.selectionBounds) {
            if (colorInput.selectionOrientation === "horizontal") {
                /* 横並び: 選択の左端揃え／真下に 1 個分離す / Horizontal: align to left edge, offset below */
                left = colorInput.selectionBounds.left;
                top = colorInput.selectionBounds.bottom - rectSize.height;
            } else if (colorInput.selectionOrientation === "vertical") {
                /* 縦並び: 選択の上端揃え／右に 1 個分離す / Vertical: align to top edge, offset to right */
                left = colorInput.selectionBounds.right + rectSize.width;
                top = colorInput.selectionBounds.top;
            }
        }
        return { left: left, top: top };
    }

    /**
     * 長方形を作ってグラデーションを適用し、必要ならグラフィックスタイルに登録する
     * 長方形を出力しない設定では、一時レイヤーの一時長方形でスタイルを登録してから片付ける
     * @param {Document} doc - 対象ドキュメント
     * @param {Gradient} gradient - 適用するグラデーション
     * @param {Object} gradientOptions - ダイアログのオプション
     * @param {Object} colorInput - collectInputColors() の結果
     * @returns {void}
     */
    function createGradientRect(doc, gradient, gradientOptions, colorInput) {
        var targetLayer = getUnlockedVisibleLayer(doc);
        if (!targetLayer) return;

        var tempRectForStyle = (!gradientOptions.makeRect && gradientOptions.registerGraphicStyle);
        var previousActiveLayer = null;
        var tempLayer = null;

        try { previousActiveLayer = doc.activeLayer; } catch (e) { /* 無視 / ignore */ }
        var previousSelection = snapshotSelection(doc);

        if (tempRectForStyle) {
            try {
                tempLayer = doc.layers.add();
                tempLayer.name = "__TempGraphicStyle";
                doc.activeLayer = tempLayer;
            } catch (e) { tempLayer = null; }
        }

        var drawLayer = (tempRectForStyle && tempLayer) ? tempLayer : targetLayer;
        var rectSize = computeRectSize(gradientOptions, colorInput);
        var rectPosition = computeRectPosition(doc, colorInput, rectSize);

        var gradientRect = drawLayer.pathItems.rectangle(rectPosition.top, rectPosition.left, rectSize.width, rectSize.height);
        doc.selection = null;
        gradientRect.selected = true;

        /* 縦並びならグラデーション角度を 90° に / If vertical, rotate gradient by action */
        if (colorInput.selectionOrientation === "vertical") runGradientAngle90Action();

        gradientRect.stroked = false;
        gradientRect.filled = true;
        var gradientFill = new GradientColor();
        gradientFill.gradient = gradient;
        gradientRect.fillColor = gradientFill;

        if (gradientOptions.registerGraphicStyle) {
            doc.selection = null;
            gradientRect.selected = true;
            try { registerGraphicStyleFromSelected(doc); } catch (e) { /* 無視 / ignore */ }
        }

        /* 一時長方形だった場合の後始末 / Clean up the temporary rectangle */
        if (tempRectForStyle) {
            try { gradientRect.remove(); } catch (e) { /* 無視 / ignore */ }
            if (tempLayer) { try { tempLayer.remove(); } catch (e) { /* 無視 / ignore */ } }
            try { if (previousActiveLayer) doc.activeLayer = previousActiveLayer; } catch (e) { /* 無視 / ignore */ }
            try {
                doc.selection = null;
                if (previousSelection.length) doc.selection = previousSelection;
            } catch (e) { /* 削除済みのオブジェクトは選択できない / removed items cannot be selected */ }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 全体フロー: 入力 → ダイアログ → スウォッチ登録 → グラデーション → 長方形・スタイル
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) return;
        var doc = app.activeDocument;

        var colorInput = collectInputColors(doc);
        if (colorInput.colors.length < 2) return;

        var disallowSeparate = (colorInput.itemCount >= 7);
        var gradientOptions = showOptionsDialog(disallowSeparate, colorInput.fromSwatches);
        if (!gradientOptions) return;

        try {
            /* 新規スウォッチグループに抽出色を登録 / Register extracted colors in a new swatch group */
            var createdSwatches = registerColorSwatches(doc, colorInput.colors, gradientOptions.makeGlobal);

            doc.selection = null;

            /* グラデーション作成 / Build the gradient */
            var gradient = null;
            if (gradientOptions.makeGradient) {
                var canSeparate = gradientOptions.separateGradient
                    && colorInput.colors.length >= 2
                    && colorInput.colors.length <= SEPARATE_MAX_COLORS;
                gradient = canSeparate
                    ? buildSeparateGradient(doc, colorInput.colors, createdSwatches)
                    : buildNormalGradient(doc, colorInput.colors, createdSwatches);
            }

            /* 長方形・グラフィックスタイル / Rectangle and Graphic Style */
            if ((gradientOptions.makeRect || gradientOptions.registerGraphicStyle) && gradient) {
                try { createGradientRect(doc, gradient, gradientOptions, colorInput); } catch (e) { /* 無視 / ignore */ }
            }

            /* 作成した最後のスウォッチ（= グラデーション）を選択 / Select the last created swatch */
            if (gradient) {
                var lastIndex = doc.swatches.length - 1;
                if (lastIndex >= 0) {
                    try { doc.swatches[lastIndex].selected = true; } catch (e) { /* 無視 / ignore */ }
                }
            }
        } catch (e) {
            /* 無言で終了 / silent */
        }
    }

    main();

})();
