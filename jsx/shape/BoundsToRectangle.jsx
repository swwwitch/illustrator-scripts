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
var SCRIPT_VERSION  = "v1.4.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

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

    // =========================================
    // ダイアログ位置の記憶 / Dialog position memory
    // =========================================
    var dialogState = $.global.__BoundsToRectangleDialogState__ || ($.global.__BoundsToRectangleDialogState__ = {});

    /**
     * ダイアログの位置を控える
     * @param {Window} dialog - 対象のダイアログ
     * @returns {void}
     */
    function saveDialogLocation(dialog) {
        var x = dialog.location[0];
        var y = dialog.location[1];
        if (isNaN(x) || isNaN(y)) return;
        dialogState.location = [x, y];
    }

    /**
     * 控えた位置にダイアログを戻す
     * @param {Window} dialog - 対象のダイアログ
     * @returns {void}
     */
    function restoreDialogLocation(dialog) {
        var savedLocation = dialogState.location;
        if (!savedLocation || savedLocation.length !== 2 || isNaN(savedLocation[0]) || isNaN(savedLocation[1])) return;
        dialog.location = [savedLocation[0], savedLocation[1]];
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

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

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} key - "checkbox.preview" のようなキー
     * @returns {string} 現在の言語の文言。見つからなければキーそのもの
     */
    function getLabel(key) {
        var keyParts = key.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (labelNode && labelNode[keyParts[i]] !== undefined) {
                labelNode = labelNode[keyParts[i]];
            } else {
                return key;
            }
        }
        return (labelNode[uiLang] !== undefined) ? labelNode[uiLang] : key;
    }

    /**
     * コロン付きラベル（日本語は全角、英語は半角）
     * @param {string} key - ラベルのキー
     * @returns {string} コロンを付けた文言
     */
    function labelText(key) {
        return getLabel(key) + (uiLang === "ja" ? "：" : ": ");
    }

    // =========================================
    // 属性パネルの［中心を表示］アクション / "Show Center" action
    // =========================================

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

        var actionFile = new File(Folder.temp + "/" + SCRIPT_NAME + "_ShowCenter.aia");
        actionFile.open("w");
        actionFile.write(actionSource);
        actionFile.close();

        app.loadAction(actionFile);
        actionFile.remove();
        app.doScript(actionName, actionSetName, false);
        app.unloadAction(actionSetName, "");
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

        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = ["center", "top"];
        buttonGroup.alignChildren = ["center", "center"];
        buttonGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        buttonGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

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
        dialog.addEventListener("keydown", function (e) {
            for (var keyIndex = 0; keyIndex < SOURCE_SHORTCUT_KEYS.length && keyIndex < sourceRadios.length; keyIndex++) {
                if (e.keyName !== SOURCE_SHORTCUT_KEYS[keyIndex]) continue;
                for (var j = 0; j < sourceRadios.length; j++) {
                    sourceRadios[j].value = (j === keyIndex);
                }
                refreshPreview();
                e.preventDefault();
                break;
            }
        });

        dialog.layout.layout(true);
        restoreDialogLocation(dialog);
        refreshPreview();

        var dialogResult = dialog.show();
        saveDialogLocation(dialog);
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
