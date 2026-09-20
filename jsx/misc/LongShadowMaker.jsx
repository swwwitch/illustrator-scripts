#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

距離・角度・スケールを指定して、選択したオブジェクトからロングシャドウを生成します。
プレビューを見ながらプリセットやオフセットで調整でき、生成後に「パスの単純化」を実行できます。

詳細は README を参照してください。

### Overview

Generates a long shadow from the selected object using a distance, an angle and a scale.
Presets and an offset are adjusted with a live preview, and a Simplify Path pass can be run afterwards.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "LongShadowMaker";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-20";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/LongShadowMaker.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/LongShadowMaker.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n0be484dab7fc"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @author こじらせたクマー（オリジナルアイデア） / original idea
 * @discussion https://note.com/nice_lotus120/n/nf406fb3ae2b4
 */

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 影の面を作るときに1セグメントを何点でサンプリングするか / Samples taken per curve segment */
    var CURVE_SAMPLE_COUNT = 8;

    /* 影の塗りの彩度（1=元のまま、0=無彩色）/ Saturation of the shadow fill (1 = as-is, 0 = gray) */
    var SHADOW_SATURATION_FACTOR = 0.7;

    /* オフセット初期値を「幅と高さの平均」の何分の1にするか / Divisor for the initial offset amount */
    var OFFSET_SIZE_DIVISOR = 20;

    /* プレビューの中間コピー（比率と不透明度）/ Intermediate preview copies (ratio and opacity) */
    var PREVIEW_STEPS = [
        { ratio: 1, opacity: 20 },
        { ratio: 0.75, opacity: 40 },
        { ratio: 0.5, opacity: 60 },
        { ratio: 0.25, opacity: 80 }
    ];

    /* プリセット（スケール／角度）。距離は変更しない / Presets (scale / angle); distance is left untouched */
    var PRESET_ITEMS = [
        "100% /  45°",
        "100% /  30°",
        "100% /  60°",
        " 50% /  90°",
        "  1% /  90°",
        "100% / 135°",
        "100% / 120°",
        "100% / 150°"
    ];

    /* 各スライダーの範囲 / Range of each slider */
    var DISTANCE_SLIDER_MIN_RANGE = 500;
    var ANGLE_MIN = -180;
    var ANGLE_MAX = 180;
    var SCALE_MIN = 1;
    var SCALE_MAX = 300;

    // =========================================
    // レイアウト / Layout
    // =========================================
    var ROW_LABEL_WIDTH = 60;     /* 行ラベルの幅 / Width of a row label */
    var UNIT_LABEL_WIDTH = 24;    /* 単位ラベルの幅 / Width of a unit label */
    var NUMBER_FIELD_CHARS = 4;   /* 数値欄の文字数 / Character width of a number field */
    var SLIDER_WIDTH = 170;       /* スライダーの幅 / Width of a slider */
    var PANEL_MARGINS = [15, 20, 15, 10];    /* パネルの余白 / Panel margins */
    var SIMPLIFY_ROW_MARGINS = [0, 10, 0, 0]; /* 単純化行の余白 / Margins of the simplify row */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 実行環境のUI言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "ロングシャドウメーカー", en: "Long Shadow Maker" }
        },
        panel: {
            settings: { ja: "設定", en: "Settings" },
            offset: { ja: "オフセット", en: "Offset" }
        },
        fieldLabel: {
            preset: { ja: "プリセット", en: "Preset" },
            join: { ja: "形状", en: "Join" },
            distance: { ja: "距離", en: "Distance" },
            angle: { ja: "角度", en: "Angle" },
            scale: { ja: "スケール", en: "Scale" }
        },
        radio: {
            joinMiter: { ja: "マイター", en: "Miter" },
            joinRound: { ja: "ラウンド", en: "Round" },
            joinBevel: { ja: "ベベル", en: "Bevel" }
        },
        checkbox: {
            simplify: { ja: "パスの単純化", en: "Simplify" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        tooltip: {
            preset: {
                ja: "スケールと角度の組み合わせをまとめて設定します。",
                en: "Sets the scale and angle together."
            },
            offsetEnabled: {
                ja: "影を作る前に、元の形を太らせます。",
                en: "Grows the original shape before the shadow is built."
            },
            offsetValue: { ja: "太らせる量です。", en: "How much to grow the shape." },
            join: {
                ja: "太らせたときの角の処理です。",
                en: "How corners are treated when the shape is grown."
            },
            distance: { ja: "影を伸ばす長さです。", en: "Length of the shadow." },
            angle: { ja: "影が伸びる向きです。", en: "Direction the shadow extends." },
            scale: {
                ja: "影の先端の大きさです。100%で元の形と同じ大きさになります。",
                en: "Size of the far end of the shadow. 100% matches the original shape."
            },
            simplify: {
                ja: "影のアンカーポイントを減らして、軽いパスにします。",
                en: "Reduces the number of anchor points in the shadow."
            },
            preview: {
                ja: "結果を画面で確認します。キャンセルすると元に戻ります。",
                en: "Shows the result on the canvas. Cancel restores the original state."
            }
        },
        alert: {
            noDocument: {
                ja: "ドキュメントを開いてください。",
                en: "Please open a document."
            },
            selectSingleShape: {
                ja: "閉パス（単一パス／複合パス／グループ）を1つだけ選択してください。",
                en: "Please select exactly one closed shape (path / compound path / group / text)."
            },
            selectClosedPath: {
                ja: "閉パスを選択してください。",
                en: "Please select a closed path."
            },
            selectClosedGroup: {
                ja: "閉パスのグループを選択してください。",
                en: "Please select a group that consists of closed paths."
            },
            groupBuildFailed: {
                ja: "グループから形状を作成できませんでした。閉パスのグループを選択してください。",
                en: "Could not build a shape from the group. Please select a group of closed paths."
            },
            notGroupItem: { ja: "GroupItem ではありません。", en: "This is not a GroupItem." },
            mergeResultMissing: {
                ja: "グループの合体結果を取得できませんでした。",
                en: "Could not retrieve the merged result from the group."
            },
            groupMergeError: {
                ja: "グループの合体中にエラーが発生しました: ",
                en: "An error occurred while merging the group: "
            },
            notTextFrame: { ja: "TextFrame ではありません。", en: "This is not a TextFrame." },
            outlineFailed: { ja: "アウトライン化に失敗しました。", en: "Failed to create outlines." },
            textMergeError: {
                ja: "テキストの合体中にエラーが発生しました: ",
                en: "An error occurred while processing the text: "
            }
        }
    };

    /**
     * ドット区切りのキーからUI言語のラベルを取得する
     * @param {string} labelPath - "dialog.title" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            if (!labelNode) return labelPath;
            labelNode = labelNode[pathKeys[i]];
        }
        if (!labelNode) return labelPath;
        return labelNode[uiLang] || labelNode.en || labelPath;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    (function () {
        if (app.documents.length === 0) {
            alert(getLabel('alert.noDocument'));
            return;
        }

        var doc = app.activeDocument;
        var currentSelection = doc.selection;

        /* 一時ベースの残骸を回収するためのタグ名 / Name tag used to sweep temporary base items */
        var TEMP_BASE_NAME = "__LongShadowTempBase__";

        // -----------------------------------------
        // 選択オブジェクトの検証 / Selection checks
        // -----------------------------------------

        /**
         * ロングシャドウの元にできる種類かどうかを判定する
         * @param {PageItem} pageItem - 判定対象
         * @returns {boolean} 対応している種類なら true
         */
        function isSupportedSourceItem(pageItem) {
            return !!pageItem && (
                pageItem.typename === "PathItem" ||
                pageItem.typename === "CompoundPathItem" ||
                pageItem.typename === "GroupItem" ||
                pageItem.typename === "TextFrame"
            );
        }

        /**
         * PathItem / CompoundPathItem / GroupItem から PathItem を再帰的に集める
         * @param {PageItem} pageItem - 走査対象
         * @returns {PathItem[]} 見つかった PathItem の配列
         */
        function collectSubPaths(pageItem) {
            var subPaths = [];
            if (!pageItem) return subPaths;

            if (pageItem.typename === "PathItem") {
                subPaths.push(pageItem);
                return subPaths;
            }

            if (pageItem.typename === "CompoundPathItem") {
                for (var i = 0; i < pageItem.pathItems.length; i++) subPaths.push(pageItem.pathItems[i]);
                return subPaths;
            }

            if (pageItem.typename === "GroupItem") {
                collectSubPathsFromContainer(pageItem, subPaths);
                return subPaths;
            }

            return subPaths;
        }

        /**
         * グループの中身をたどって PathItem を集める
         * @param {GroupItem} container - 走査対象のグループ
         * @param {PathItem[]} collected - 集めた PathItem を追加する配列
         * @returns {void}
         */
        function collectSubPathsFromContainer(container, collected) {
            if (!container) return;

            for (var i = 0; i < container.pageItems.length; i++) {
                var childItem = container.pageItems[i];
                if (!childItem) continue;

                if (childItem.typename === "PathItem") {
                    collected.push(childItem);
                } else if (childItem.typename === "CompoundPathItem") {
                    for (var j = 0; j < childItem.pathItems.length; j++) collected.push(childItem.pathItems[j]);
                } else if (childItem.typename === "GroupItem") {
                    collectSubPathsFromContainer(childItem, collected);
                }
            }
        }

        /**
         * 渡されたパスがすべて閉じているかを判定する
         * @param {PathItem[]} subPaths - 判定対象のパス
         * @returns {boolean} 1つ以上あり、すべて閉じていれば true
         */
        function areAllPathsClosed(subPaths) {
            if (!subPaths || subPaths.length === 0) return false;
            for (var i = 0; i < subPaths.length; i++) {
                /* 生成直後のパスは closed を読めないことがある / closed may not be readable yet */
                try {
                    if (!subPaths[i].closed) return false;
                } catch (e) {
                    return false;
                }
            }
            return true;
        }

        if (currentSelection.length !== 1 || !isSupportedSourceItem(currentSelection[0])) {
            alert(getLabel('alert.selectSingleShape'));
            return;
        }

        var sourceItem = currentSelection[0];
        var sourceSubPaths = collectSubPaths(sourceItem);

        /* Path/Compound はここで閉パス検証。Group/Text は実行時に一時パスへ変換して検証する
           Paths and compound paths are checked here; groups and text are checked after conversion */
        if (sourceItem.typename !== "GroupItem" && sourceItem.typename !== "TextFrame") {
            if (!areAllPathsClosed(sourceSubPaths)) {
                alert(getLabel('alert.selectClosedPath'));
                return;
            }
        }

        // -----------------------------------------
        // 一時ベースの後始末 / Temporary base cleanup
        // -----------------------------------------

        /**
         * 一時ベースとその子要素に名前タグを付け外しする
         * @param {PageItem} pageItem - 対象アイテム
         * @param {string} tagName - 付ける名前（空文字でタグを外す）
         * @returns {void}
         */
        function setTempBaseTag(pageItem, tagName) {
            if (!pageItem) return;
            /* 名前を持たない種類のアイテムがある / some item types reject a name */
            try { pageItem.name = tagName; } catch (e) { }

            if (pageItem.typename === 'GroupItem') {
                for (var i = 0; i < pageItem.pageItems.length; i++) setTempBaseTag(pageItem.pageItems[i], tagName);
            } else if (pageItem.typename === 'CompoundPathItem') {
                for (var j = 0; j < pageItem.pathItems.length; j++) {
                    try { pageItem.pathItems[j].name = tagName; } catch (e) { }
                }
            }
        }

        /**
         * アイテム自身と親（グループ／レイヤー）のロックと非表示を解除する
         * @param {PageItem} pageItem - 対象アイテム
         * @returns {void}
         */
        function unlockItemAndAncestors(pageItem) {
            try { pageItem.locked = false; } catch (e) { }
            try { pageItem.hidden = false; } catch (e) { }

            var parent = null;
            try { parent = pageItem.parent; } catch (e) { parent = null; }

            var guard = 0;
            while (parent && guard++ < 50) {
                if (parent.typename === 'Layer') {
                    try { parent.locked = false; } catch (e) { }
                    try { parent.visible = true; } catch (e) { }
                    break;
                }
                if (parent.typename === 'GroupItem') {
                    try { parent.locked = false; } catch (e) { }
                    try { parent.hidden = false; } catch (e) { }
                }
                try { parent = parent.parent; } catch (e) { parent = null; }
            }
        }

        /**
         * ロックや非表示を解除したうえでアイテムを確実に削除する
         * @param {PageItem} pageItem - 削除するアイテム
         * @returns {void}
         */
        function forceRemoveItem(pageItem) {
            try { if (!pageItem || !pageItem.isValid) return; } catch (e) { return; }

            unlockItemAndAncestors(pageItem);

            /* まず通常の削除 / first try a plain remove */
            try { pageItem.remove(); return; } catch (e) { }

            /* 削除できないときは選択してカット / fall back to selecting and clearing */
            try {
                doc.selection = null;
                pageItem.selected = true;
                app.executeMenuCommand('clear');
            } catch (e) { }
            doc.selection = null;

            try { if (pageItem && pageItem.isValid) pageItem.remove(); } catch (e) { }
        }

        /**
         * ドキュメント内に残った一時ベース（名前タグ付き）をすべて削除する
         * @returns {void}
         */
        function removeTempBaseItemsByName() {
            /* 一時ベースはグループであることが多いので先に走査 / temp bases are usually groups */
            var groups = doc.groupItems;
            for (var i = groups.length - 1; i >= 0; i--) {
                if (readItemName(groups[i]) === TEMP_BASE_NAME) forceRemoveItem(groups[i]);
            }

            /* 取りこぼし（グループ以外）を掃除 / sweep the non-group leftovers */
            var items = doc.pageItems;
            for (var j = items.length - 1; j >= 0; j--) {
                if (readItemName(items[j]) === TEMP_BASE_NAME) forceRemoveItem(items[j]);
            }
        }

        /**
         * アイテム名を安全に読み取る
         * @param {PageItem} pageItem - 対象アイテム
         * @returns {string} 読み取れない場合は空文字
         */
        function readItemName(pageItem) {
            try {
                if (!pageItem || !pageItem.isValid) return '';
                return pageItem.name;
            } catch (e) {
                return '';
            }
        }

        // -----------------------------------------
        // 一時ベースの生成 / Building the temporary base
        // -----------------------------------------

        /**
         * 一時生成物を控えて、あとでまとめて削除するための入れ物を作る
         * @returns {{track: function, disposeAll: function}} 追跡用と削除用の関数
         */
        function createTempItemTracker() {
            var temporaryItems = [];

            return {
                track: function (pageItem) {
                    if (pageItem) temporaryItems.push(pageItem);
                    return pageItem;
                },
                disposeAll: function () {
                    doc.selection = null;
                    for (var i = temporaryItems.length - 1; i >= 0; i--) {
                        try {
                            if (temporaryItems[i] && temporaryItems[i].isValid) temporaryItems[i].remove();
                        } catch (e) { }
                    }
                }
            };
        }

        /* 複製を合体して単一のパスにする / Merge the duplicate down to a single path */
        function executePathfinderAddAndExpand() {
            app.executeMenuCommand('Live Pathfinder Add');
            app.executeMenuCommand('expandStyle');
        }

        /**
         * GroupItem を一時的に単一パス／複合パスへ合体して返す（元グループは残す）
         * @param {GroupItem} groupItem - 合体するグループ
         * @returns {{item: PageItem, cleanup: function, ok: boolean, message: string}} 合体結果
         */
        function buildMergedItemFromGroup(groupItem) {
            var result = { item: null, cleanup: function () { }, ok: false, message: "" };

            if (!groupItem || groupItem.typename !== "GroupItem") {
                result.message = getLabel('alert.notGroupItem');
                return result;
            }

            var tracker = createTempItemTracker();
            result.cleanup = tracker.disposeAll;

            try {
                var duplicatedGroup = tracker.track(groupItem.duplicate());

                doc.selection = null;
                duplicatedGroup.selected = true;
                executePathfinderAddAndExpand();

                var expandedItems = doc.selection;
                doc.selection = null;
                trackAll(tracker, expandedItems);

                if (!expandedItems || expandedItems.length === 0) {
                    result.message = getLabel('alert.mergeResultMissing');
                    return result;
                }

                var mergedItem = expandedItems[0];

                /* 複数残った場合は一度グループ化して再合体 / regroup and merge again when several remain */
                if (expandedItems.length > 1) {
                    var regrouped = tracker.track(doc.groupItems.add());
                    for (var i = 0; i < expandedItems.length; i++) {
                        try { expandedItems[i].move(regrouped, ElementPlacement.PLACEATEND); } catch (e) { }
                    }
                    doc.selection = null;
                    regrouped.selected = true;
                    executePathfinderAddAndExpand();

                    var remergedItems = doc.selection;
                    doc.selection = null;
                    trackAll(tracker, remergedItems);

                    if (remergedItems && remergedItems.length > 0) mergedItem = remergedItems[0];
                }

                tracker.track(mergedItem);
                setTempBaseTag(mergedItem, TEMP_BASE_NAME);
                result.item = mergedItem;
                result.ok = true;
                return result;

            } catch (e) {
                result.message = getLabel('alert.groupMergeError') + e;
                return result;
            }
        }

        /**
         * 選択結果の配列をまとめて一時生成物として控える
         * @param {{track: function}} tracker - 一時生成物の入れ物
         * @param {PageItem[]} items - 控えるアイテム（null 可）
         * @returns {void}
         */
        function trackAll(tracker, items) {
            if (!items || !items.length) return;
            for (var i = 0; i < items.length; i++) tracker.track(items[i]);
        }

        /**
         * TextFrame を一時的にアウトライン化して返す（元テキストは残す）
         * @param {TextFrame} textFrame - アウトライン化するテキスト
         * @returns {{item: PageItem, cleanup: function, ok: boolean, message: string}} アウトライン化結果
         */
        function buildMergedItemFromText(textFrame) {
            var result = { item: null, cleanup: function () { }, ok: false, message: "" };

            if (!textFrame || textFrame.typename !== "TextFrame") {
                result.message = getLabel('alert.notTextFrame');
                return result;
            }

            var tracker = createTempItemTracker();
            result.cleanup = tracker.disposeAll;

            try {
                /* 複製に対してアウトライン化するので、オリジナルは残る
                   createOutline() consumes the duplicate, so the original survives */
                var duplicatedText = tracker.track(textFrame.duplicate());

                var outlinedItem = null;
                try { outlinedItem = duplicatedText.createOutline(); } catch (e) { outlinedItem = null; }

                if (!outlinedItem) {
                    result.message = getLabel('alert.outlineFailed');
                    return result;
                }

                /* 控えるのはアウトライン化で生成されたルートだけにする。子要素まで控えると、
                   後段で移動・合体された要素を巻き込んで生成済みの影が消えることがある
                   Track only the outlined root; tracking its children would delete the finished shadow */
                tracker.track(outlinedItem);

                /* 選択状態が残ると後段の処理に巻き込まれるので解除 / clear the selection before moving on */
                doc.selection = null;
                try { outlinedItem.selected = false; } catch (e) { }

                setTempBaseTag(outlinedItem, TEMP_BASE_NAME);
                result.item = outlinedItem;
                result.ok = true;
                return result;

            } catch (e) {
                result.message = getLabel('alert.textMergeError') + e;
                return result;
            }
        }

        // -----------------------------------------
        // 色処理 / Color utilities
        // -----------------------------------------

        /**
         * 0〜1の範囲に丸める
         * @param {number} value - 丸める値
         * @returns {number} 0〜1に収めた値
         */
        function clamp01(value) {
            return Math.max(0, Math.min(1, value));
        }

        /**
         * 塗り色をRGBの成分に変換する（GrayColor と未対応の色は null）
         * @param {Color} fillColor - 元の塗り色
         * @returns {{red: number, green: number, blue: number}|null} 0〜255のRGB成分
         */
        function toRgbComponents(fillColor) {
            if (fillColor.typename === "RGBColor") {
                return { red: fillColor.red, green: fillColor.green, blue: fillColor.blue };
            }

            if (fillColor.typename === "CMYKColor") {
                /* 簡易 CMYK -> RGB（0-100 を 0-1 に換算）/ rough CMYK to RGB conversion */
                var cyan = fillColor.cyan / 100.0;
                var magenta = fillColor.magenta / 100.0;
                var yellow = fillColor.yellow / 100.0;
                var black = fillColor.black / 100.0;
                return {
                    red: 255 * (1 - cyan) * (1 - black),
                    green: 255 * (1 - magenta) * (1 - black),
                    blue: 255 * (1 - yellow) * (1 - black)
                };
            }

            return null;
        }

        /**
         * HSLの中間値から1チャンネル分の値を求める
         * @param {number} lowerBound - 下側の値
         * @param {number} upperBound - 上側の値
         * @param {number} hueFraction - 0〜1に正規化した色相
         * @returns {number} 0〜1のチャンネル値
         */
        function hueToChannel(lowerBound, upperBound, hueFraction) {
            var t = hueFraction;
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1 / 6) return lowerBound + (upperBound - lowerBound) * 6 * t;
            if (t < 1 / 2) return upperBound;
            if (t < 2 / 3) return lowerBound + (upperBound - lowerBound) * (2 / 3 - t) * 6;
            return lowerBound;
        }

        /**
         * 元の塗り色をもとに、彩度を下げた影用の色を作る
         * @param {Color} fillColor - 元の塗り色
         * @param {number} factor - 彩度の倍率（1=元のまま、0=無彩色）
         * @returns {Color|null} 影用の色。色が無い場合は null
         */
        function desaturateColorFromFill(fillColor, factor) {
            if (!fillColor) return null;
            if (factor === undefined || factor === null) factor = SHADOW_SATURATION_FACTOR;

            /* グレーはそのまま複製する / gray is copied as-is */
            if (fillColor.typename === "GrayColor") {
                var grayCopy = new GrayColor();
                grayCopy.gray = fillColor.gray;
                return grayCopy;
            }

            var rgb = toRgbComponents(fillColor);
            /* 未対応（NoColor / グラデーション / パターン）はそのまま返す / unsupported fills pass through */
            if (!rgb) return fillColor;

            var red = clamp01(rgb.red / 255.0);
            var green = clamp01(rgb.green / 255.0);
            var blue = clamp01(rgb.blue / 255.0);

            var maxChannel = Math.max(red, green, blue);
            var minChannel = Math.min(red, green, blue);
            var chroma = maxChannel - minChannel;
            var lightness = (maxChannel + minChannel) / 2;
            var hue = 0;
            var saturation = 0;

            if (chroma !== 0) {
                saturation = chroma / (1 - Math.abs(2 * lightness - 1));
                switch (maxChannel) {
                    case red: hue = ((green - blue) / chroma) % 6; break;
                    case green: hue = ((blue - red) / chroma) + 2; break;
                    case blue: hue = ((red - green) / chroma) + 4; break;
                }
                hue = hue * 60;
                if (hue < 0) hue += 360;
            }

            saturation = clamp01(saturation * factor);

            var resultRed, resultGreen, resultBlue;
            if (saturation === 0) {
                resultRed = resultGreen = resultBlue = lightness;
            } else {
                var upperBound = lightness < 0.5 ? lightness * (1 + saturation) : (lightness + saturation - lightness * saturation);
                var lowerBound = 2 * lightness - upperBound;
                var hueFraction = hue / 360;
                resultRed = hueToChannel(lowerBound, upperBound, hueFraction + 1 / 3);
                resultGreen = hueToChannel(lowerBound, upperBound, hueFraction);
                resultBlue = hueToChannel(lowerBound, upperBound, hueFraction - 1 / 3);
            }

            var desaturatedColor = new RGBColor();
            desaturatedColor.red = Math.round(resultRed * 255);
            desaturatedColor.green = Math.round(resultGreen * 255);
            desaturatedColor.blue = Math.round(resultBlue * 255);
            return desaturatedColor;
        }

        // -----------------------------------------
        // 影の面を作る / Building the shadow faces
        // -----------------------------------------

        /**
         * 3次ベジェ曲線上の点を求める
         * @param {number[]} p0 - 始点
         * @param {number[]} p1 - 始点側の方向点
         * @param {number[]} p2 - 終点側の方向点
         * @param {number[]} p3 - 終点
         * @param {number} t - 0〜1の位置
         * @returns {number[]} [x, y] 座標
         */
        function bezierPoint(p0, p1, p2, p3, t) {
            var u = 1 - t;
            var tt = t * t;
            var uu = u * u;
            var uuu = uu * u;
            var ttt = tt * t;
            return [
                uuu * p0[0] + 3 * uu * t * p1[0] + 3 * u * tt * p2[0] + ttt * p3[0],
                uuu * p0[1] + 3 * uu * t * p1[1] + 3 * u * tt * p2[1] + ttt * p3[1]
            ];
        }

        /**
         * アイテムを複製し、中心を基準に拡大縮小してから移動する
         * @param {PageItem} pageItem - 複製するアイテム
         * @param {number} dx - 水平方向の移動量（pt）
         * @param {number} dy - 垂直方向の移動量（pt）
         * @param {number} scalePercent - 拡大率（100=等倍）
         * @returns {PageItem|null} 複製したアイテム
         */
        function duplicateWithOffsetAndScale(pageItem, dx, dy, scalePercent) {
            if (!pageItem) return null;

            var duplicatedItem = null;
            try { duplicatedItem = pageItem.duplicate(); } catch (e) { duplicatedItem = null; }
            if (!duplicatedItem) return null;

            resizeFromCenter(duplicatedItem, scalePercent);
            try { duplicatedItem.translate(dx, dy); } catch (e) { }
            return duplicatedItem;
        }

        /**
         * 中心を基準にアイテムを拡大縮小する
         * @param {PageItem} pageItem - 対象アイテム
         * @param {number} scalePercent - 拡大率（100=等倍なら何もしない）
         * @returns {void}
         */
        function resizeFromCenter(pageItem, scalePercent) {
            var scaleValue = isFinite(scalePercent) ? Number(scalePercent) : 100;
            if (scaleValue === 100) return;
            /* テキストや効果付きのアイテムで resize が失敗することがある / resize can fail on some items */
            try {
                pageItem.resize(scaleValue, scaleValue, true, true, true, true, true, Transformation.CENTER);
            } catch (e) { }
        }

        /**
         * 閉パスを曲線も含めて点列にする
         * @param {PathItem} pathItem - 点列にするパス
         * @param {number} curveSamplesPerSegment - 1セグメントあたりのサンプル数
         * @returns {number[][]} [x, y] の配列
         */
        function samplePathToPoints(pathItem, curveSamplesPerSegment) {
            if (!curveSamplesPerSegment) curveSamplesPerSegment = CURVE_SAMPLE_COUNT;
            var points = [];
            if (!pathItem) return points;

            /* 一時アイテムは走査中に無効化されることがある / temporary items can go invalid mid-scan */
            try {
                var pathPoints = pathItem.pathPoints;
                var pointCount = pathPoints.length;
                if (pointCount < 2) return points;

                for (var i = 0; i < pointCount; i++) {
                    var currentPoint = pathPoints[i];
                    var nextPoint = pathPoints[(i + 1) % pointCount];

                    for (var j = 0; j < curveSamplesPerSegment; j++) {
                        points.push(bezierPoint(
                            currentPoint.anchor,
                            currentPoint.rightDirection,
                            nextPoint.leftDirection,
                            nextPoint.anchor,
                            j / curveSamplesPerSegment
                        ));
                    }
                }
            } catch (e) { }

            return points;
        }

        /**
         * 重なり順を保ったまま閉じた PathItem だけを集める
         * @param {PageItem} pageItem - 走査対象
         * @returns {PathItem[]} 閉じた PathItem の配列
         */
        function collectClosedPaths(pageItem) {
            var closedPaths = [];
            if (!pageItem) return closedPaths;

            if (pageItem.typename === 'PathItem') {
                pushIfClosed(closedPaths, pageItem);
                return closedPaths;
            }

            if (pageItem.typename === 'CompoundPathItem') {
                for (var i = 0; i < pageItem.pathItems.length; i++) {
                    pushIfClosed(closedPaths, pageItem.pathItems[i]);
                }
                return closedPaths;
            }

            if (pageItem.typename === 'GroupItem') {
                /* pageItems 順にたどる（見た目の重なり順に近い）/ follow pageItems order */
                for (var j = 0; j < pageItem.pageItems.length; j++) {
                    var nestedPaths = collectClosedPaths(pageItem.pageItems[j]);
                    for (var k = 0; k < nestedPaths.length; k++) closedPaths.push(nestedPaths[k]);
                }
            }

            return closedPaths;
        }

        /**
         * 閉じた PathItem のときだけ配列に加える
         * @param {PathItem[]} closedPaths - 追加先の配列
         * @param {PageItem} pathCandidate - 判定するアイテム
         * @returns {void}
         */
        function pushIfClosed(closedPaths, pathCandidate) {
            /* closed を読めない種類のアイテムが混ざる / closed is not readable on every item type */
            try {
                if (pathCandidate && pathCandidate.typename === 'PathItem' && pathCandidate.closed) {
                    closedPaths.push(pathCandidate);
                }
            } catch (e) { }
        }

        /**
         * 2つの点列の間を四角形の面で埋める
         * @param {GroupItem} parentGroup - 面を追加するグループ
         * @param {number[][]} nearPoints - 元の形の点列
         * @param {number[][]} farPoints - 影の先端側の点列
         * @param {Color} faceFill - 面の塗り色
         * @returns {PathItem[]} 生成した面
         */
        function buildSideFaces(parentGroup, nearPoints, farPoints, faceFill) {
            var createdFaces = [];
            if (!parentGroup || !nearPoints || !farPoints) return createdFaces;

            var pointCount = Math.min(nearPoints.length, farPoints.length);
            if (pointCount < 2) return createdFaces;

            for (var i = 0; i < pointCount; i++) {
                var nextIndex = (i + 1) % pointCount;
                var facePath = null;
                /* 極端に小さい面はパス生成が失敗することがある / very small faces can fail to build */
                try {
                    facePath = parentGroup.pathItems.add();
                    facePath.setEntirePath([nearPoints[i], farPoints[i], farPoints[nextIndex], nearPoints[nextIndex]]);
                    facePath.closed = true;
                    facePath.stroked = false;
                    facePath.filled = true;
                    if (faceFill) facePath.fillColor = faceFill;
                    createdFaces.push(facePath);
                } catch (e) {
                    try { if (facePath && facePath.isValid) facePath.remove(); } catch (err) { }
                }
            }
            return createdFaces;
        }

        /**
         * 元の形と複製した形の間を面でつないで影のグループを作る
         * @param {PageItem} baseItem - 影の元になる形
         * @param {number} dx - 水平方向の移動量（pt）
         * @param {number} dy - 垂直方向の移動量（pt）
         * @param {number} scalePercent - 先端の拡大率（100=等倍）
         * @returns {GroupItem|null} 生成した影のグループ
         */
        function buildShadowFaces(baseItem, dx, dy, scalePercent) {
            if (!baseItem) return null;

            var nearPaths = collectClosedPaths(baseItem);
            if (nearPaths.length === 0) return null;

            var farItem = duplicateWithOffsetAndScale(baseItem, dx, dy, scalePercent);
            if (!farItem) return null;

            var farPaths = collectClosedPaths(farItem);
            /* 数が違う場合は少ない方に合わせる（落とさない優先）/ pair up to the smaller count */
            var pairCount = Math.min(nearPaths.length, farPaths.length);
            if (pairCount === 0) {
                try { farItem.remove(); } catch (e) { }
                return null;
            }

            var shadowGroup = doc.groupItems.add();

            var shadowFill = null;
            /* 塗りを持たないパスがある / some paths carry no fill */
            try {
                shadowFill = desaturateColorFromFill(nearPaths[0].fillColor, SHADOW_SATURATION_FACTOR);
            } catch (e) {
                shadowFill = null;
            }

            for (var i = 0; i < pairCount; i++) {
                var nearPoints = samplePathToPoints(nearPaths[i], CURVE_SAMPLE_COUNT);
                var farPoints = samplePathToPoints(farPaths[i], CURVE_SAMPLE_COUNT);
                if (nearPoints.length < 2 || farPoints.length < 2) continue;
                buildSideFaces(shadowGroup, nearPoints, farPoints, shadowFill);
            }

            try { farItem.remove(); } catch (e) { }
            return shadowGroup;
        }

        // -----------------------------------------
        // オフセット効果 / Offset Path live effect
        // -----------------------------------------

        /**
         * 選択中の角の処理に対応する Offset Path のコードを返す
         * @returns {number} 0=ラウンド、1=ベベル、2=マイター
         */
        function getJoinCode() {
            if (joinRoundRadio.value) return 0;
            if (joinBevelRadio.value) return 1;
            return 2;
        }

        /**
         * Offset Path のライブ効果XMLを組み立てる
         * @param {number} offsetPt - オフセット量（pt）
         * @param {number} joinCode - 角の処理（0=ラウンド、1=ベベル、2=マイター）
         * @returns {string} applyEffect() に渡すXML
         */
        function buildOffsetEffectXML(offsetPt, joinCode) {
            /* mlim はマイター制限（既定4）、ofst は pt / mlim is the miter limit, ofst is in points */
            return '<LiveEffect name="Adobe Offset Path"><Dict data="R mlim 4 R ofst ' + offsetPt + ' I jntp ' + joinCode + ' "/></LiveEffect>';
        }

        /**
         * 生成した影にオフセット（ライブ効果）を適用する
         * @param {GroupItem} shadowGroup - 適用先のグループ
         * @param {boolean} isFinalRun - 本実行なら true（プレビューでは適用しない）
         * @returns {void}
         */
        function applyOffsetEffect(shadowGroup, isFinalRun) {
            if (!shadowGroup || !isFinalRun) return;
            if (!offsetCheckbox.value) return;

            var offsetPt = Number(offsetInput.text);
            if (isNaN(offsetPt)) offsetPt = 0;

            /* 0でもONなら適用する（結果が変わらないだけ）/ apply even at 0; it simply changes nothing */
            var joinCode = getJoinCode();
            try { shadowGroup.applyEffect(buildOffsetEffectXML(offsetPt, joinCode)); } catch (e) { }

            /* ラウンドのときは後処理として Pathfinder Merge を実行 / round joins need a merge pass */
            if (joinCode === 0) {
                try {
                    doc.selection = null;
                    shadowGroup.selected = true;
                    app.executeMenuCommand('Live Pathfinder Merge');
                } catch (e) { }
                doc.selection = null;
            }
        }

        // -----------------------------------------
        // パスの単純化 / Simplify paths
        // -----------------------------------------

        /**
         * 単純化の対象になるパスを集める
         * @param {PageItem} container - 走査対象
         * @param {PageItem[]} collected - 集めたアイテムを追加する配列
         * @returns {void}
         */
        function collectSimplifyTargets(container, collected) {
            if (!container) return;

            if (container.typename === 'PathItem' || container.typename === 'CompoundPathItem') {
                collected.push(container);
                return;
            }
            if (container.typename === 'GroupItem') {
                for (var i = 0; i < container.pageItems.length; i++) {
                    collectSimplifyTargets(container.pageItems[i], collected);
                }
            }
        }

        /**
         * 「パスの単純化」を実行する（Illustratorの仕様でダイアログが開く）
         * @param {PageItem} pageItem - 単純化するアイテム
         * @returns {void}
         */
        function simplifyPathsInItem(pageItem) {
            if (!pageItem || !simplifyCheckbox.value) return;

            var simplifyTargets = [];
            collectSimplifyTargets(pageItem, simplifyTargets);
            if (!simplifyTargets.length) return;

            doc.selection = null;
            for (var i = 0; i < simplifyTargets.length; i++) {
                try { simplifyTargets[i].selected = true; } catch (e) { }
            }

            /* このメニューコマンドは単純化ダイアログを開く（Illustratorの制限）
               This menu command opens the Simplify dialog (Illustrator limitation) */
            try { app.executeMenuCommand("simplify menu item"); } catch (e) { }

            doc.selection = null;
        }

        // -----------------------------------------
        // ダイアログ / Dialog
        // -----------------------------------------

        var isPreviewing = false;
        var hasFinished = false;

        /* 数値欄とスライダーの同期関数。ダイアログ構築より前に用意する
           Sync handlers per number field; must exist before the dialog is built */
        var fieldSyncHandlers = [];

        var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        dialog.orientation = "column";
        dialog.alignChildren = "fill";

        /* プリセット行（左右中央に配置）/ Preset row, centered in the dialog */
        var presetRowGroup = dialog.add("group");
        presetRowGroup.orientation = "row";
        presetRowGroup.alignChildren = ["center", "center"];
        presetRowGroup.alignment = "center";

        var presetGroup = presetRowGroup.add("group");
        presetGroup.orientation = "row";
        presetGroup.alignChildren = ["left", "center"];
        presetGroup.add("statictext", undefined, getLabel('fieldLabel.preset'));

        var presetDropdown = presetGroup.add("dropdownlist", undefined, PRESET_ITEMS);
        presetDropdown.selection = 0;
        presetDropdown.helpTip = getLabel('tooltip.preset');

        var panelColumnGroup = dialog.add("group");
        panelColumnGroup.orientation = "column";
        panelColumnGroup.alignChildren = ["fill", "top"];
        panelColumnGroup.alignment = "fill";

        var settingsPanel = panelColumnGroup.add("panel", undefined, getLabel('panel.settings'));
        settingsPanel.orientation = "column";
        settingsPanel.alignChildren = "left";
        settingsPanel.margins = PANEL_MARGINS;

        var offsetPanel = panelColumnGroup.add("panel", undefined, getLabel('panel.offset'));
        offsetPanel.orientation = "column";
        offsetPanel.alignChildren = "left";
        offsetPanel.margins = PANEL_MARGINS;

        /* オフセットパネル内は2カラム（左=量／右=角の処理）/ Offset panel holds amount and join columns */
        var offsetRowGroup = offsetPanel.add("group");
        offsetRowGroup.orientation = "row";
        offsetRowGroup.alignChildren = ["fill", "top"];

        var offsetValueGroup = offsetRowGroup.add("group");
        offsetValueGroup.orientation = "column";
        offsetValueGroup.alignChildren = ["left", "top"];

        var offsetInputGroup = offsetValueGroup.add("group");
        offsetInputGroup.orientation = "row";
        offsetInputGroup.alignChildren = ["left", "center"];

        var offsetCheckbox = offsetInputGroup.add("checkbox", undefined, "");
        offsetCheckbox.helpTip = getLabel('tooltip.offsetEnabled');
        offsetCheckbox.value = false;

        var offsetInput = offsetInputGroup.add("edittext", undefined, String(getInitialOffsetPt()));
        offsetInput.characters = NUMBER_FIELD_CHARS;
        offsetInput.helpTip = getLabel('tooltip.offsetValue');

        var offsetUnitLabel = offsetInputGroup.add("statictext", undefined, "pt");

        var joinColumnGroup = offsetRowGroup.add("group");
        joinColumnGroup.orientation = "column";
        joinColumnGroup.alignChildren = ["fill", "top"];

        var joinRowGroup = joinColumnGroup.add("group");
        joinRowGroup.orientation = "row";
        joinRowGroup.alignChildren = ["left", "top"];
        joinRowGroup.margins = [0, 0, 0, 0];

        var joinLabelGroup = joinRowGroup.add("group");
        joinLabelGroup.orientation = "column";
        joinLabelGroup.alignChildren = ["left", "top"];
        var joinLabel = joinLabelGroup.add("statictext", undefined, getLabel('fieldLabel.join'));
        joinLabel.preferredSize.width = ROW_LABEL_WIDTH;
        joinLabel.justify = "right";

        var joinRadioGroup = joinRowGroup.add("group");
        joinRadioGroup.orientation = "column";
        joinRadioGroup.alignChildren = ["left", "center"];

        var joinMiterRadio = joinRadioGroup.add("radiobutton", undefined, getLabel('radio.joinMiter'));
        var joinRoundRadio = joinRadioGroup.add("radiobutton", undefined, getLabel('radio.joinRound'));
        var joinBevelRadio = joinRadioGroup.add("radiobutton", undefined, getLabel('radio.joinBevel'));
        joinMiterRadio.helpTip = getLabel('tooltip.join');
        joinRoundRadio.helpTip = getLabel('tooltip.join');
        joinBevelRadio.helpTip = getLabel('tooltip.join');
        joinRoundRadio.value = true;

        offsetCheckbox.onClick = updateOffsetControlsEnabled;
        updateOffsetControlsEnabled();

        /* 距離の初期値は元のオブジェクトの「幅＋高さ」/ Default distance is width plus height */
        var sourceBoundsPt = sourceItem.geometricBounds; /* [left, top, right, bottom] */
        var defaultDistancePt = Math.round((sourceBoundsPt[2] - sourceBoundsPt[0]) + (sourceBoundsPt[1] - sourceBoundsPt[3]));
        var maxDistancePt = Math.max(DISTANCE_SLIDER_MIN_RANGE, defaultDistancePt * 3);

        var distanceInput = addSliderRow(settingsPanel, 'fieldLabel.distance', 'tooltip.distance',
            String(defaultDistancePt), "pt", 0, maxDistancePt, defaultDistancePt, false);
        var angleInput = addSliderRow(settingsPanel, 'fieldLabel.angle', 'tooltip.angle',
            "45", "°", ANGLE_MIN, ANGLE_MAX, 45, true);
        var scaleInput = addSliderRow(settingsPanel, 'fieldLabel.scale', 'tooltip.scale',
            "100", "%", SCALE_MIN, SCALE_MAX, 100, false);

        var simplifyGroup = settingsPanel.add("group");
        simplifyGroup.orientation = "row";
        simplifyGroup.alignChildren = ["center", "center"];
        simplifyGroup.alignment = "center";
        simplifyGroup.margins = SIMPLIFY_ROW_MARGINS;

        var simplifyCheckbox = simplifyGroup.add("checkbox", undefined, getLabel('checkbox.simplify'));
        simplifyCheckbox.helpTip = getLabel('tooltip.simplify');
        simplifyCheckbox.value = true;
        simplifyCheckbox.alignment = "center";

        /* ボタン行（左=プレビュー／中央=スペーサー／右=キャンセル・OK）/ Button row */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "top"];
        btnRowGroup.alignChildren = ["left", "center"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.orientation = "row";
        btnLeftGroup.alignChildren = ["left", "center"];

        var previewCheckbox = btnLeftGroup.add("checkbox", undefined, getLabel('checkbox.preview'));
        previewCheckbox.helpTip = getLabel('tooltip.preview');
        previewCheckbox.value = true;

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.alignment = ["right", "center"];

        var btnCancel = btnRightGroup.add("button", undefined, getLabel('button.cancel'), { name: "cancel" });
        var btnOk = btnRightGroup.add("button", undefined, getLabel('button.ok'), { name: "ok" });
        btnOk.active = true;

        changeValueByArrowKey(offsetInput, false, false);

        /**
         * 元のオブジェクトのサイズからオフセットの初期値を求める
         * @returns {number} オフセットの初期値（pt）
         */
        function getInitialOffsetPt() {
            var bounds = null;
            /* 種類によって geometricBounds を読めないことがある / geometricBounds is not always readable */
            try { bounds = sourceItem.geometricBounds; } catch (e) { bounds = null; }
            if (!bounds) {
                try { bounds = sourceItem.visibleBounds; } catch (e) { bounds = null; }
            }
            if (!bounds || bounds.length !== 4) return 0;

            var widthPt = Math.abs(bounds[2] - bounds[0]);
            var heightPt = Math.abs(bounds[1] - bounds[3]);
            var averageSizePt = (widthPt + heightPt) / 2;
            var offsetBasePt = averageSizePt / OFFSET_SIZE_DIVISOR;

            return isNaN(offsetBasePt) ? 0 : Math.round(offsetBasePt);
        }

        /**
         * 「ラベル＋数値欄＋単位＋スライダー」の1行を作る
         * @param {Panel} parentPanel - 追加先のパネル
         * @param {string} labelPath - 行ラベルのラベルキー
         * @param {string} tooltipPath - ツールチップのラベルキー
         * @param {string} initialText - 数値欄の初期値
         * @param {string} unitText - 単位の表示
         * @param {number} minValue - スライダーの最小値
         * @param {number} maxValue - スライダーの最大値
         * @param {number} initialValue - スライダーの初期値
         * @param {boolean} allowNegative - 負の値を許可するか
         * @returns {EditText} 作成した数値欄
         */
        function addSliderRow(parentPanel, labelPath, tooltipPath, initialText, unitText, minValue, maxValue, initialValue, allowNegative) {
            var rowGroup = parentPanel.add("group");

            var rowLabel = rowGroup.add("statictext", undefined, getLabel(labelPath));
            rowLabel.preferredSize.width = ROW_LABEL_WIDTH;
            rowLabel.justify = "right";

            var numberInput = rowGroup.add("edittext", undefined, initialText);
            numberInput.characters = NUMBER_FIELD_CHARS;
            numberInput.helpTip = getLabel(tooltipPath);

            var unitLabel = rowGroup.add("statictext", undefined, unitText);
            unitLabel.preferredSize.width = UNIT_LABEL_WIDTH;

            var rowSlider = rowGroup.add("slider", undefined, initialValue, minValue, maxValue);
            rowSlider.preferredSize.width = SLIDER_WIDTH;
            rowSlider.helpTip = getLabel(tooltipPath);

            changeValueByArrowKey(numberInput, allowNegative, true);
            bindSliderToInput(numberInput, rowSlider, minValue, maxValue);

            return numberInput;
        }

        /**
         * オフセットのコントロールをチェックボックスに合わせて有効／無効にする
         * @returns {void}
         */
        function updateOffsetControlsEnabled() {
            var isEnabled = !!offsetCheckbox.value;
            offsetInput.enabled = isEnabled;
            offsetUnitLabel.enabled = isEnabled;
            joinRowGroup.enabled = isEnabled;
            joinMiterRadio.enabled = isEnabled;
            joinRoundRadio.enabled = isEnabled;
            joinBevelRadio.enabled = isEnabled;
        }

        // -----------------------------------------
        // 数値欄の操作 / Number field behavior
        // -----------------------------------------

        /**
         * 数値欄の内容をスライダーへ反映する（登録済みの欄のみ）
         * @param {EditText} editText - 対象の数値欄
         * @returns {void}
         */
        function syncFieldToSlider(editText) {
            for (var i = 0; i < fieldSyncHandlers.length; i++) {
                if (fieldSyncHandlers[i].input === editText) {
                    fieldSyncHandlers[i].sync();
                    return;
                }
            }
        }

        /**
         * 値を最小値と最大値の間に丸める
         * @param {number} value - 丸める値
         * @param {number} minValue - 最小値
         * @param {number} maxValue - 最大値
         * @returns {number} 範囲内に収めた値
         */
        function clamp(value, minValue, maxValue) {
            return Math.max(minValue, Math.min(maxValue, value));
        }

        /**
         * 数値欄とスライダーを双方向に同期させる
         * @param {EditText} editText - 数値欄
         * @param {Slider} slider - スライダー
         * @param {number} minValue - 最小値
         * @param {number} maxValue - 最大値
         * @returns {void}
         */
        function bindSliderToInput(editText, slider, minValue, maxValue) {
            var isSyncing = false;

            function setInputFromSlider() {
                if (isSyncing) return;
                isSyncing = true;
                editText.text = String(Math.round(slider.value));
                isSyncing = false;
                refreshPreviewIfEnabled();
            }

            function setSliderFromInput() {
                if (isSyncing) return;
                isSyncing = true;
                var value = Number(editText.text);
                if (isNaN(value)) value = 0;
                slider.value = Math.round(clamp(value, minValue, maxValue));
                isSyncing = false;
            }

            setSliderFromInput();
            fieldSyncHandlers.push({ input: editText, sync: setSliderFromInput });

            slider.onChanging = setInputFromSlider;
            slider.onChange = setInputFromSlider;

            editText.onChanging = function () {
                setSliderFromInput();
                refreshPreviewIfEnabled();
            };
        }

        /**
         * ↑↓キーで数値を増減する（Shift=±10、Option=±0.1）
         * @param {EditText} editText - 対象の数値欄
         * @param {boolean} allowNegative - 負の値を許可するか
         * @param {boolean} enablePreviewUpdate - 変更時にプレビューを更新するか
         * @returns {void}
         */
        function changeValueByArrowKey(editText, allowNegative, enablePreviewUpdate) {
            editText.addEventListener("keydown", function (event) {
                if (event.keyName !== "Up" && event.keyName !== "Down") return;

                var value = Number(editText.text);
                if (isNaN(value)) return;

                var keyboard = ScriptUI.environment.keyboardState;
                var isUp = (event.keyName === "Up");

                if (keyboard.shiftKey) {
                    /* Shift押下時は10の倍数にスナップ / snap to multiples of 10 */
                    value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
                } else if (keyboard.altKey) {
                    /* Option押下時は0.1単位で増減 / step by 0.1 */
                    value = Math.round((value + (isUp ? 0.1 : -0.1)) * 10) / 10;
                } else {
                    value = Math.round(value + (isUp ? 1 : -1));
                }
                event.preventDefault();

                if (!allowNegative && value < 0) value = 0;
                editText.text = value;
                syncFieldToSlider(editText);

                if (enablePreviewUpdate !== false) refreshPreviewIfEnabled();
            });
        }

        // -----------------------------------------
        // プレビュー / Preview
        // -----------------------------------------

        /**
         * プレビューがONのときだけ再描画する
         * @returns {void}
         */
        function refreshPreviewIfEnabled() {
            if (previewCheckbox && previewCheckbox.value) updatePreview();
        }

        /**
         * 直前のプレビューを取り消して描き直す
         * @returns {void}
         */
        function updatePreview() {
            undoPreview();
            if (previewCheckbox.value) {
                buildLongShadow(false);
                isPreviewing = true;
            }
            app.redraw();
        }

        /**
         * 表示中のプレビューを取り消す
         * @returns {void}
         */
        function undoPreview() {
            if (!isPreviewing) return;
            app.undo();
            isPreviewing = false;
        }

        /**
         * プレビュー用の半透明コピーを1つ作る
         * @param {PageItem} previewSource - 複製元のアイテム
         * @param {number} ratio - 影の先端までの比率（0〜1）
         * @param {number} opacityPercent - 不透明度（%）
         * @param {number} dx - 影の先端までの水平移動量（pt）
         * @param {number} dy - 影の先端までの垂直移動量（pt）
         * @param {number} scalePercent - 影の先端の拡大率
         * @returns {PageItem|null} 作成したコピー
         */
        function addPreviewCopy(previewSource, ratio, opacityPercent, dx, dy, scalePercent) {
            var previewCopy = null;
            try { previewCopy = previewSource.duplicate(); } catch (e) { previewCopy = null; }
            if (!previewCopy) return null;

            /* 一時ベースの複製はタグを外さないと後始末で消えてしまう
               Clear the temp tag, otherwise the cleanup sweep deletes the preview */
            setTempBaseTag(previewCopy, "");

            resizeFromCenter(previewCopy, 100 + (scalePercent - 100) * ratio);
            try { previewCopy.translate(dx * ratio, dy * ratio); } catch (e) { }
            try { previewCopy.move(sourceItem, ElementPlacement.PLACEBEFORE); } catch (e) { }
            try { previewCopy.opacity = opacityPercent; } catch (e) { }
            return previewCopy;
        }

        // -----------------------------------------
        // 実行 / Execution
        // -----------------------------------------

        /**
         * スケールを実際に使う倍率へ変換する
         * @param {number} rawScalePercent - 入力されたスケール（%）
         * @returns {number} 実際に使う倍率（%）
         */
        function normalizeScalePercent(rawScalePercent) {
            var value = Number(rawScalePercent);
            if (isNaN(value)) return 100;
            /* プリセット「1%」は極小の影を作るため0.01%として扱う / the 1% preset runs as 0.01% */
            if (value === 1) return 0.01;
            return value;
        }

        /**
         * 入力欄から距離・角度・スケールを読み取る
         * @returns {{dx: number, dy: number, scalePercent: number}} 影の先端までの移動量と拡大率
         */
        function readShadowParameters() {
            var distancePt = parseFloat(distanceInput.text) || 0;
            var scalePercent = normalizeScalePercent(parseFloat(scaleInput.text));
            if (isNaN(scalePercent) || scalePercent <= 0) scalePercent = 100;

            /* 入力角度を反転して画面座標に合わせる / flip the angle to match screen coordinates */
            var angleRadians = -(parseFloat(angleInput.text) || 0) * Math.PI / 180;

            return {
                dx: distancePt * Math.cos(angleRadians),
                dy: distancePt * Math.sin(angleRadians),
                scalePercent: scalePercent
            };
        }

        /**
         * 影の元になる形を用意する（グループとテキストは一時パスへ変換する）
         * @param {boolean} isFinalRun - 本実行なら true
         * @returns {{item: PageItem, cleanup: function, ok: boolean, message: string}} 影の元になる形
         */
        function prepareShadowBaseItem(isFinalRun) {
            var result = { item: sourceItem, cleanup: function () { }, ok: true, message: "" };

            var merged = null;
            if (sourceItem.typename === "GroupItem") {
                merged = buildMergedItemFromGroup(sourceItem);
                if (!merged.ok || !merged.item) {
                    merged.message = merged.message || getLabel('alert.groupBuildFailed');
                    return merged;
                }
            } else if (isFinalRun && sourceItem.typename === "TextFrame") {
                merged = buildMergedItemFromText(sourceItem);
                if (!merged.ok || !merged.item) {
                    merged.message = merged.message || getLabel('alert.selectClosedPath');
                    return merged;
                }
            }

            if (!merged) return result;

            if (!areAllPathsClosed(collectSubPaths(merged.item))) {
                merged.ok = false;
                merged.message = (sourceItem.typename === "GroupItem")
                    ? getLabel('alert.selectClosedGroup')
                    : getLabel('alert.selectClosedPath');
            }
            return merged;
        }

        /**
         * 生成した影を元のオブジェクトの背面へ置く
         * @param {PageItem} shadowItem - 生成した影
         * @param {PageItem} originalItem - 元のオブジェクト
         * @returns {void}
         */
        function placeShadowBehindOriginal(shadowItem, originalItem) {
            if (!shadowItem || !originalItem) return;

            /* まずは元オブジェクトの直後（背面側）へ / first try placing it right behind the original */
            try {
                shadowItem.move(originalItem, ElementPlacement.PLACEAFTER);
                return;
            } catch (e) { }

            /* TextFrame などで move が失敗する場合は同一レイヤーの最背面へ
               When move fails, fall back to the back of the same layer */
            try {
                var ownerLayer = originalItem.layer;
                if (ownerLayer) {
                    shadowItem.move(ownerLayer, ElementPlacement.PLACEATBEGINNING);
                    return;
                }
            } catch (e) { }

            /* 最後の手段。これで見えなくなることもあるので最後に試す / last resort */
            try { shadowItem.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
        }

        /**
         * 影の面を合体させ、オフセットと単純化を適用したグループを返す
         * @param {GroupItem} shadowGroup - 面を集めたグループ
         * @param {boolean} isFinalRun - 本実行なら true
         * @returns {GroupItem} 仕上げた影のグループ
         */
        function finishShadowGroup(shadowGroup, isFinalRun) {
            /* パスの穴を潰すため、Merge → Add → Expand の順で実行 / merge, add, then expand */
            try {
                doc.selection = null;
                shadowGroup.selected = true;
                app.executeMenuCommand('Live Pathfinder Merge');
                app.executeMenuCommand('Live Pathfinder Add');
                app.executeMenuCommand('expandStyle');
            } catch (e) { }

            var mergedShadowItems = doc.selection;
            doc.selection = null;

            var shadowResultGroup = shadowGroup;
            if (mergedShadowItems && mergedShadowItems.length) {
                /* アクティブレイヤーがロックされているとグループを追加できない
                   groupItems.add() fails when the active layer is locked */
                try { shadowResultGroup = createGroupFrom(mergedShadowItems); } catch (e) { }
            }

            simplifyPathsInItem(shadowResultGroup);
            applyOffsetEffect(shadowResultGroup, isFinalRun);
            return shadowResultGroup;
        }

        /**
         * 渡された順序のままアイテムを新しいグループにまとめる
         * @param {PageItem[]} items - まとめるアイテム
         * @returns {GroupItem} 作成したグループ
         */
        function createGroupFrom(items) {
            var newGroup = doc.groupItems.add();
            for (var i = 0; i < items.length; i++) {
                try { items[i].move(newGroup, ElementPlacement.PLACEATEND); } catch (e) { }
            }
            return newGroup;
        }

        /**
         * ロングシャドウを生成する
         * @param {boolean} isFinalRun - 本実行なら true、プレビューなら false
         * @returns {void}
         */
        function buildLongShadow(isFinalRun) {
            var params = readShadowParameters();
            var base = prepareShadowBaseItem(isFinalRun);

            function cleanupTempBase() {
                base.cleanup();
                removeTempBaseItemsByName();
            }

            if (!base.ok) {
                cleanupTempBase();
                alert(base.message);
                return;
            }

            /* プレビューでは影を合体させず、半透明のコピーを並べるだけにする
               The preview stacks translucent copies instead of building the merged shadow */
            if (!isFinalRun) {
                for (var i = 0; i < PREVIEW_STEPS.length; i++) {
                    addPreviewCopy(base.item, PREVIEW_STEPS[i].ratio, PREVIEW_STEPS[i].opacity,
                        params.dx, params.dy, params.scalePercent);
                }
                doc.selection = null;
                cleanupTempBase();
                app.redraw();
                return;
            }

            var shadowGroup = buildShadowFaces(base.item, params.dx, params.dy, params.scalePercent);
            if (!shadowGroup) {
                cleanupTempBase();
                alert(getLabel('alert.selectClosedPath'));
                return;
            }

            var shadowResultGroup = finishShadowGroup(shadowGroup, isFinalRun);
            placeShadowBehindOriginal(shadowResultGroup, sourceItem);
            cleanupTempBase();

            /* 単一パスから作ったときは、生成した影を選択状態にする
               Select the generated shadow when the source was a single path */
            doc.selection = null;
            try {
                if (sourceItem.typename === "PathItem" && shadowResultGroup.isValid) {
                    doc.selection = [shadowResultGroup];
                }
            } catch (e) { }

            app.redraw();
        }

        // -----------------------------------------
        // イベントリスナー / Event listeners
        // -----------------------------------------

        /* プリセット選択時：スケールと角度に反映（距離は変更しない）
           Presets set the scale and the angle; the distance is left as it is */
        presetDropdown.onChange = function () {
            if (!presetDropdown.selection) return;

            var matched = String(presetDropdown.selection.text)
                .match(/^\s*(-?\d+(?:\.\d+)?)\s*%\s*\/\s*(-?\d+(?:\.\d+)?)\s*°\s*$/);
            if (!matched) return;

            scaleInput.text = matched[1];
            angleInput.text = matched[2];
            syncFieldToSlider(scaleInput);
            syncFieldToSlider(angleInput);

            refreshPreviewIfEnabled();
        };

        simplifyCheckbox.onClick = refreshPreviewIfEnabled;
        previewCheckbox.onClick = updatePreview;

        btnOk.onClick = function () {
            /* プレビューを消してから確定実行 / drop the preview before the real run */
            undoPreview();
            hasFinished = true;
            buildLongShadow(true);
            dialog.close();
        };

        btnCancel.onClick = function () {
            undoPreview();
            hasFinished = true;
            removeTempBaseItemsByName();
            dialog.close();
        };

        /* 閉じるボタンで閉じられたときもプレビューを後片付けする
           Clean the preview up when the dialog is dismissed by its close box */
        dialog.onClose = function () {
            if (hasFinished) return;
            undoPreview();
            removeTempBaseItemsByName();
            app.redraw();
        };

        /* ダイアログを開いた時点でプレビューを表示 / show the preview as the dialog opens */
        refreshPreviewIfEnabled();
        dialog.show();

    })();

})();
