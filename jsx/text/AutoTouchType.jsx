#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの各文字に対して、ベースライン・比率・回転・カーニング・トラッキングをランダムに付与します。
seed付きの乱数でプレビューの見た目を安定させ、文字回転を適用したときはトラッキングを自動補正します。

詳細は README を参照してください。

### Overview

Randomizes the baseline shift, scale, rotation, kerning and tracking of each character in the selected text.
A seeded RNG keeps the preview stable, and the tracking is corrected automatically when character rotation is applied.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AutoTouchType";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.9";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoTouchType.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AutoTouchType.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/ne6545c4717af"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 「犯行声明文」風の背景を置くレイヤー名とグループ名 / Layer and group that hold the ransom-note backgrounds */
    var RANSOM_BG_LAYER_NAME = "__AutoTouchType_BG__";
    var RANSOM_BG_GROUP_NAME = "__BGRects__";

    /* 背景長方形の余白・ゆがみ量（pt）とグレー濃度（%）/ Padding, jitter (pt) and gray range (%) of the background rectangles */
    var RANSOM_RECT_PADDING_PT = 1;
    var RANSOM_RECT_JITTER_PT = 1;
    var RANSOM_GRAY_MIN = 10;
    var RANSOM_GRAY_MAX = 50;

    /* 「犯行声明文」風のトラッキング初期値（1/1000em）/ Default ransom-note tracking (1/1000 em) */
    var RANSOM_TRACKING_DEFAULT = 200;

    /* 文字タッチの初期値 / Default touch amounts */
    var DEFAULT_SCALE_PERCENT = 10;
    var DEFAULT_KERNING_EM = 50;
    var DEFAULT_ROTATION_DEG = 5;

    /* スライダーの範囲 / Slider ranges */
    var SCALE_SLIDER_MAX = 200;
    var KERNING_SLIDER_MIN = -200;
    var KERNING_SLIDER_MAX = 200;
    var ROTATION_SLIDER_MAX = 30;
    var BASELINE_SLIDER_MAX_FALLBACK = 50;
    var ZOOM_MIN_PERCENT = 10;
    var ZOOM_MAX_PERCENT = 1600;

    /* 画面更新の間引き（ミリ秒）/ Redraw throttling in milliseconds */
    var PREVIEW_DELAY_MS = 120;
    var ZOOM_THROTTLE_MS = 150;

    /* 文字回転に対するトラッキング補正の係数 / Factors of the rotation tracking compensation */
    var ROTATION_GAP_FACTOR_AFTER_POSITIVE = 680;
    var ROTATION_GAP_FACTOR_AFTER_NEGATIVE = 1050;
    var ROTATION_GAP_FACTOR_BEFORE = 180;
    var ROTATION_GAP_UPPERCASE_BEFORE = 0.6;
    var ROTATION_GAP_UPPERCASE_AFTER = 1.05;

    /* ダイアログ位置を覚える環境設定キー / Preference keys that remember the dialog position */
    var PREF_KEY_DIALOG_X = "AutoTouchType/dialogX";
    var PREF_KEY_DIALOG_Y = "AutoTouchType/dialogY";

    // =========================================
    // レイアウト / Layout
    // =========================================

    var DIALOG_OPACITY = 0.98;
    var PANEL_MARGINS = [15, 20, 15, 10];
    var SLIDER_WIDTH = 180;
    var ZOOM_SLIDER_WIDTH = 270;
    var TOGGLE_WIDTH = 15;
    var VALUE_FIELD_CHARS = 4;
    var RANSOM_FIELD_CHARS = 5;
    var SMALL_BUTTON_SIZE = [74, 22];
    var BUTTON_ROW_TOP_MARGIN = 15;
    var TOUCH_BUTTON_ROW_TOP_MARGIN = 10;

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UIの表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "オート文字タッチツール", en: "Auto Touch Type Tool" }
        },
        panel: {
            touch: {
                ja: "位置・スケール・回転などの「ゆらぎ」",
                en: "Position, Scale and Rotation Variation"
            },
            font: { ja: "フォント", en: "Font" },
            ransom: { ja: "オプション", en: "Options" }
        },
        checkbox: {
            randomFont: { ja: "1文字ごとに変更", en: "Change per character" },
            japaneseOnly: { ja: "和文フォントに限定", en: "Japanese only" },
            ransomEnabled: { ja: "「犯行声明文」風", en: "Ransom-note style" },
            ransomTracking: { ja: "トラッキング調整", en: "Adjust tracking" },
            rotationTracking: {
                ja: "文字回転によるトラッキング補正",
                en: "Tracking compensation for rotation"
            },
            lightMode: { ja: "軽量モード", en: "Light mode" }
        },
        fieldLabel: {
            baseline: { ja: "ベースライン", en: "Baseline" },
            scale: { ja: "水平／垂直比率", en: "Scale" },
            rotation: { ja: "文字回転", en: "Rotation" },
            kerning: { ja: "カーニング", en: "Kerning" },
            zoom: { ja: "ズーム", en: "Zoom" }
        },
        button: {
            randomize: { ja: "ランダム", en: "Random" },
            reset: { ja: "リセット", en: "Reset" },
            allOn: { ja: "すべてON", en: "All ON" },
            allOff: { ja: "すべてOFF", en: "All OFF" },
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok: { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectText: { ja: "テキストを選択してください。", en: "Please select text." },
            selectTextRange: {
                ja: "テキスト、または文字を選択してください。",
                en: "Please select a text object or characters."
            },
            enterNumber: { ja: "数値を入力してください。", en: "Please enter a number." },
            noJapaneseFonts: {
                ja: "対象の和文フォント（Pr6 / Pr6N）が見つかりません。",
                en: "No target JP fonts (Pr6 / Pr6N) were found."
            }
        },
        tooltip: {
            randomFont: { ja: "1文字ずつフォントを入れ替えます。", en: "Swaps the font of each character." },
            japaneseOnly: {
                ja: "入れ替え先を和文フォント（Pr6／Pr6N）だけに絞ります。",
                en: "Limits the replacement fonts to Japanese fonts (Pr6 / Pr6N)."
            },
            ransomEnabled: {
                ja: "1文字ずつ大きさと書体をばらつかせ、切り貼りしたような見た目にします。",
                en: "Varies the size and typeface of each character, like letters cut from a magazine."
            },
            ransomTracking: {
                ja: "ばらついた文字幅に合わせて字間を詰めます。",
                en: "Tightens the spacing to match the varied character widths."
            },
            ransomTrackingValue: { ja: "詰める量です。単位は1/1000em。", en: "How much to tighten, in 1/1000 em." },
            baselineEnabled: { ja: "ベースラインシフトをかけるかどうかです。", en: "Whether to apply a baseline shift." },
            baseline: { ja: "1文字ずつ上下にずらす最大量です。", en: "Maximum amount each character is shifted up or down." },
            scaleEnabled: { ja: "文字を長体・平体にするかどうかです。", en: "Whether to condense or extend the characters." },
            scale: {
                ja: "1文字ずつ変える水平／垂直比率の最大量です。",
                en: "Maximum change applied to each character's horizontal and vertical scale."
            },
            kerningEnabled: { ja: "字間をばらつかせるかどうかです。", en: "Whether to vary the spacing between characters." },
            kerning: { ja: "1文字ずつ変える字間の最大量です。単位は1/1000em。", en: "Maximum spacing change per character, in 1/1000 em." },
            rotationEnabled: { ja: "文字を回転させるかどうかです。", en: "Whether to rotate the characters." },
            rotation: { ja: "1文字ずつ回転させる最大角度です。", en: "Maximum rotation applied to each character." },
            rotationTracking: {
                ja: "回転で広がった見た目の幅を、字間で打ち消します。",
                en: "Offsets the apparent width added by the rotation with the letter spacing."
            },
            allOn: { ja: "文字タッチの4項目をすべてオンにします。", en: "Turns on all four touch settings." },
            allOff: { ja: "文字タッチの4項目をすべてオフにします。", en: "Turns off all four touch settings." },
            zoom: {
                ja: "作業中の画面表示倍率を変えます。結果には影響しません。",
                en: "Changes the view zoom while you work. It does not affect the result."
            },
            lightMode: {
                ja: "ドラッグ中は画面表示倍率を変えず、離したときにまとめて反映します。",
                en: "Applies the zoom only when the drag ends, instead of while dragging."
            },
            randomize: { ja: "同じ設定のまま、乱数だけ振り直します。", en: "Re-rolls the randomness while keeping the same settings." },
            reset: {
                ja: "選択したテキストの文字属性を初期状態に戻します。",
                en: "Restores the character attributes of the selected text."
            }
        }
    };

    /**
     * ドット区切りキーで文言を引く
     * @param {string} labelKey - "panel.touch" のようなキー
     * @returns {string} 表示言語の文言（見つからないときはキーそのもの）
     */
    function getLabel(labelKey) {
        var keyParts = labelKey.split(".");
        var node = LABELS;
        for (var i = 0; i < keyParts.length; i++) {
            if (node == null) return labelKey;
            node = node[keyParts[i]];
        }
        if (node == null) return labelKey;
        return node[uiLang] || node.ja || node.en || labelKey;
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelKey - ドット区切りキー
     * @returns {string} コロンを付けた文言
     */
    function labelText(labelKey) {
        return getLabel(labelKey) + (uiLang === "ja" ? "：" : ": ");
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /* Q ではなく H と表示する設定キー / Preference keys that display H instead of Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 設定キーごとの単位情報を取得する
     * @param {string} prefKey - 環境設定キー（省略時は "rulerType"）
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        var unit = UNITS[unitCode] || UNITS[2];
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) { alert(getLabel("alert.noDocument")); return; }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) { alert(getLabel("alert.selectText")); return; }

    // -----------------------------------------
    // 画面表示倍率 / View zoom
    // -----------------------------------------

    /**
     * 複数オブジェクトを囲む可視バウンディングボックスを返す
     * @param {Object[]} items - ページアイテムの配列
     * @returns {number[]} [left, top, right, bottom]
     */
    function getUnionVisibleBounds(items) {
        var bounds = items[0].visibleBounds;
        var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
        for (var i = 1; i < items.length; i++) {
            var itemBounds = items[i].visibleBounds;
            if (itemBounds[0] < left) left = itemBounds[0];
            if (itemBounds[1] > top) top = itemBounds[1];
            if (itemBounds[2] > right) right = itemBounds[2];
            if (itemBounds[3] < bottom) bottom = itemBounds[3];
        }
        return [left, top, right, bottom];
    }

    /**
     * バウンディングボックスの中心点を返す
     * @param {number[]} bounds - [left, top, right, bottom]
     * @returns {number[]} [x, y]
     */
    function getBoundsCenter(bounds) {
        return [bounds[0] + (bounds[2] - bounds[0]) / 2, bounds[1] + (bounds[3] - bounds[1]) / 2];
    }

    /**
     * 表示倍率をスライダーの範囲に収める
     * @param {number} zoomPercent - 表示倍率（%）
     * @returns {number} 範囲内に丸めた表示倍率（%）
     */
    function clampZoomPercent(zoomPercent) {
        var percent = Math.round(zoomPercent);
        if (percent < ZOOM_MIN_PERCENT) percent = ZOOM_MIN_PERCENT;
        if (percent > ZOOM_MAX_PERCENT) percent = ZOOM_MAX_PERCENT;
        return percent;
    }

    var docView = doc.views[0];
    var originalZoomFactor = docView.zoom;
    var originalCenterPoint = docView.centerPoint;
    var zoomCenterPoint = originalCenterPoint;
    /* 選択に TextRange が混じると visibleBounds を持たない / A TextRange in the selection has no visibleBounds */
    try {
        zoomCenterPoint = getBoundsCenter(getUnionVisibleBounds(doc.selection));
    } catch (e) { }

    /**
     * 指定した表示倍率でドキュメントを表示し直す
     * @param {number} zoomPercent - 表示倍率（%）
     * @returns {void}
     */
    function applyZoomPercent(zoomPercent) {
        docView.zoom = clampZoomPercent(zoomPercent) / 100.0;
        if (zoomCenterPoint) docView.centerPoint = zoomCenterPoint;
        app.redraw();
    }

    /**
     * 開始時の表示倍率と表示位置に戻す
     * @returns {void}
     */
    function restoreOriginalView() {
        docView.zoom = originalZoomFactor;
        docView.centerPoint = originalCenterPoint;
        app.redraw();
    }

    // -----------------------------------------
    // 「犯行声明文」風の背景 / Ransom-note backgrounds
    // -----------------------------------------

    /**
     * 選択範囲に含まれるテキストフレームを重複なく集める
     * @param {Object[]} selection - 選択オブジェクトの配列
     * @returns {TextFrame[]} テキストフレームの配列
     */
    function getSelectionTextFrames(selection) {
        var textFrames = [];
        if (!selection || selection.length === 0) return textFrames;

        /* 同じフレームを二重に入れない / Keep the list unique */
        function pushUnique(textFrame) {
            if (!textFrame || textFrame.typename !== "TextFrame") return;
            for (var k = 0; k < textFrames.length; k++) {
                if (textFrames[k] === textFrame) return;
            }
            textFrames.push(textFrame);
        }

        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (!item) continue;
            if (item.typename === "TextFrame") {
                pushUnique(item);
            } else if (item.typename === "TextRange") {
                /* 文字編集中の選択は親をたどってフレームにする / A text-editing selection resolves to its frame */
                var owner = item.parent;
                if (owner && owner.typename === "TextFrame") pushUnique(owner);
                else if (owner && owner.parent && owner.parent.typename === "TextFrame") pushUnique(owner.parent);
            }
        }
        return textFrames;
    }

    /**
     * 背景用レイヤーを探し、無ければ作る
     * @returns {Layer} 背景用レイヤー
     */
    function ensureRansomBgLayer() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === RANSOM_BG_LAYER_NAME) return doc.layers[i];
        }
        var bgLayer = doc.layers.add();
        bgLayer.name = RANSOM_BG_LAYER_NAME;
        return bgLayer;
    }

    /**
     * 背景用グループを削除する
     * @param {Layer} bgLayer - 背景用レイヤー
     * @returns {void}
     */
    function removeRansomBgGroup(bgLayer) {
        if (!bgLayer) return;
        for (var i = bgLayer.groupItems.length - 1; i >= 0; i--) {
            if (bgLayer.groupItems[i].name === RANSOM_BG_GROUP_NAME) bgLayer.groupItems[i].remove();
        }
    }

    /**
     * min以上max以下の整数を返す
     * @param {number} minValue - 最小値
     * @param {number} maxValue - 最大値
     * @returns {number} 乱数
     */
    function randomIntBetween(minValue, maxValue) {
        return Math.floor(Math.random() * (maxValue - minValue + 1)) + minValue;
    }

    /**
     * 濃度をばらつかせたグレーを作る
     * @returns {GrayColor} 背景用のグレー
     */
    function createRandomGrayFill() {
        var grayColor = new GrayColor();
        grayColor.gray = randomIntBetween(RANSOM_GRAY_MIN, RANSOM_GRAY_MAX);
        return grayColor;
    }

    /**
     * 長方形の4隅を外側へランダムにずらす
     * @param {PathItem} rect - 対象の長方形
     * @param {number} maxOffsetPt - ずらす最大量（pt）
     * @returns {void}
     */
    function expandRectCornersRandomly(rect, maxOffsetPt) {
        if (!rect.pathPoints || rect.pathPoints.length < 4) return;
        var bounds = rect.geometricBounds;
        var centerX = (bounds[0] + bounds[2]) / 2;
        var centerY = (bounds[1] + bounds[3]) / 2;
        for (var i = 0; i < rect.pathPoints.length; i++) {
            var pathPoint = rect.pathPoints[i];
            var anchorX = pathPoint.anchor[0];
            var anchorY = pathPoint.anchor[1];
            var dx = anchorX - centerX;
            var dy = anchorY - centerY;
            var distance = Math.sqrt(dx * dx + dy * dy);
            if (!distance) continue;
            var offset = Math.random() * maxOffsetPt;
            var offsetX = (dx / distance) * offset;
            var offsetY = (dy / distance) * offset;
            pathPoint.anchor = [anchorX + offsetX, anchorY + offsetY];
            pathPoint.leftDirection = [pathPoint.leftDirection[0] + offsetX, pathPoint.leftDirection[1] + offsetY];
            pathPoint.rightDirection = [pathPoint.rightDirection[0] + offsetX, pathPoint.rightDirection[1] + offsetY];
        }
    }

    /**
     * 1文字分のバウンディングボックスから背景の長方形を作る
     * @param {GroupItem} bgGroup - 背景用グループ
     * @param {number[]} charBounds - [left, top, right, bottom]
     * @returns {PathItem|null} 作成した長方形（作れないときは null）
     */
    function addRansomRect(bgGroup, charBounds) {
        var rectLeft = charBounds[0] - RANSOM_RECT_PADDING_PT;
        var rectTop = charBounds[1] + RANSOM_RECT_PADDING_PT;
        var rectWidth = (charBounds[2] + RANSOM_RECT_PADDING_PT) - rectLeft;
        var rectHeight = rectTop - (charBounds[3] - RANSOM_RECT_PADDING_PT);
        if (rectWidth <= 0 || rectHeight <= 0) return null;
        var rect = bgGroup.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
        rect.stroked = false;
        rect.filled = true;
        rect.fillColor = createRandomGrayFill();
        expandRectCornersRandomly(rect, RANSOM_RECT_JITTER_PT);
        return rect;
    }

    /**
     * アウトライン化した結果から1文字ぶんのまとまりを集める
     * @param {Object} outlinedItem - createOutline() の戻り値
     * @param {Object[]} collected - 集めた結果を入れる配列
     * @returns {void}
     */
    function collectOutlinedCharItems(outlinedItem, collected) {
        if (!outlinedItem) return;

        /* 末端のパスまで降りて拾う / Walk down to the leaf paths */
        function pushLeafItems(container) {
            if (!container || !container.pageItems) return;
            for (var k = 0; k < container.pageItems.length; k++) {
                var item = container.pageItems[k];
                if (!item) continue;
                if (item.typename === "GroupItem") pushLeafItems(item);
                else if (item.typename === "PathItem" || item.typename === "CompoundPathItem") collected.push(item);
            }
        }

        if (outlinedItem.typename !== "GroupItem") {
            collected.push(outlinedItem);
            return;
        }

        /* 行グループ → 文字グループの順に入れ子になっている / Line groups hold the character groups */
        var pushedCharGroup = false;
        for (var i = 0; i < outlinedItem.groupItems.length; i++) {
            var lineGroup = outlinedItem.groupItems[i];
            if (!lineGroup || lineGroup.typename !== "GroupItem") continue;
            if (lineGroup.groupItems.length > 0) {
                for (var j = 0; j < lineGroup.groupItems.length; j++) {
                    collected.push(lineGroup.groupItems[j]);
                    pushedCharGroup = true;
                }
            } else {
                collected.push(lineGroup);
                pushedCharGroup = true;
            }
        }
        if (pushedCharGroup) return;

        pushLeafItems(outlinedItem);
        if (collected.length === 0) collected.push(outlinedItem);
    }

    /**
     * テキストを複製してアウトライン化し、1文字ごとの背景長方形を作る
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {void}
     */
    function createRansomBgRects(textFrames) {
        if (!textFrames || textFrames.length === 0) return;
        var bgLayer = ensureRansomBgLayer();
        removeRansomBgGroup(bgLayer);
        var bgGroup = bgLayer.groupItems.add();
        bgGroup.name = RANSOM_BG_GROUP_NAME;

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!textFrame || textFrame.typename !== "TextFrame") continue;

            /* ロックや非表示のレイヤー上では複製・アウトライン化に失敗することがある
               Duplicating or outlining can fail on locked or hidden layers */
            var duplicatedFrame = null;
            var outlinedGroup = null;
            try {
                duplicatedFrame = textFrame.duplicate();
                duplicatedFrame.move(textFrame, ElementPlacement.PLACEAFTER);
            } catch (e) { }
            if (!duplicatedFrame) continue;
            try {
                outlinedGroup = duplicatedFrame.createOutline();
            } catch (e) { }
            /* createOutline() は複製自体をアウトライン化して消費するため、残っているときだけ片付ける
               createOutline() consumes the duplicate, so remove it only when it is still there */
            try { duplicatedFrame.remove(); } catch (e) { }
            if (!outlinedGroup) continue;

            var charItems = [];
            collectOutlinedCharItems(outlinedGroup, charItems);
            for (var j = 0; j < charItems.length; j++) {
                /* スペースなど中身の無い文字グループは geometricBounds を取れない
                   An empty character group, such as a space, has no geometricBounds */
                try {
                    addRansomRect(bgGroup, charItems[j].geometricBounds);
                } catch (e) { }
            }
            outlinedGroup.remove();
        }

        bgGroup.zOrder(ZOrderMethod.SENDTOBACK);
        bgLayer.zOrder(ZOrderMethod.SENDTOBACK);
    }

    /**
     * 背景の長方形を消す（無ければ何もしない）
     * @returns {void}
     */
    function clearRansomBgRects() {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === RANSOM_BG_LAYER_NAME) {
                removeRansomBgGroup(doc.layers[i]);
                return;
            }
        }
    }

    // -----------------------------------------
    // フォント / Fonts
    // -----------------------------------------

    var allFonts = app.textFonts;

    /**
     * ひらがな・カタカナ・漢字を含むか判定する
     * @param {string} text - 判定する文字列
     * @returns {boolean} 含むとき true
     */
    function hasJapaneseCharacters(text) {
        if (!text) return false;
        return /[぀-ゟ゠-ヿ一-鿿]/.test(String(text));
    }

    /* 名前に含まれていたら和文として扱わない語 / Markers that rule a font out */
    var NON_JAPANESE_FONT_MARKERS = [
        "Apple LiGothic",
        "RyoGothicStd",
        "-KO", "-KL", "LogoArl",
        "Kana"
    ];

    /* 和文フォントとみなす語（和文ファミリー名とベンダー名）/ Keywords that mark a font as Japanese */
    var JAPANESE_FONT_KEYWORDS = [
        "ゴシック", "明朝", "丸ゴ", "教科書", "楷書",
        "Mincho", "Maru",
        "Hiragino", "ヒラギノ",
        "Yu Gothic", "Yu Mincho", "游ゴシック", "游明朝",
        "Meiryo", "メイリオ",
        "MS Gothic", "MS Mincho", "MS ゴシック", "MS 明朝",
        "Kozuka", "小塚",
        "Morisawa", "モリサワ",
        "Ryumin", "Shin Go", "新ゴ",
        "Heisei", "平成",
        "Klee", "クレー",
        "Tsukushi", "筑紫",
        "A-OTF", "AP-OTF ", "-OTF",
        "FOT", "Pr6N", "Pr6",
        "Noto Sans JP", "Noto Serif JP",
        "Source Han", "源ノ角", "源ノ明",
        "Min2"
    ];

    /**
     * 和文フォントかどうかを名前から判定する
     * @param {TextFont} font - 判定するフォント
     * @returns {boolean} 和文とみなせるとき true
     */
    function isJapaneseFont(font) {
        if (!font) return false;
        var names = [];
        /* 環境にないフォントは名前を取れないことがある / A font missing from the system may not expose its names */
        try {
            names = [String(font.name || ""), String(font.family || ""), String(font.fullName || ""), String(font.postScriptName || "")];
        } catch (e) {
            return false;
        }

        var i, j;
        for (i = 0; i < NON_JAPANESE_FONT_MARKERS.length; i++) {
            for (j = 0; j < names.length; j++) {
                if (names[j].indexOf(NON_JAPANESE_FONT_MARKERS[i]) !== -1) return false;
            }
        }

        /* 名前そのものが和文表記ならそれで判定できる / A Japanese name settles it */
        for (j = 0; j < 3; j++) {
            if (hasJapaneseCharacters(names[j])) return true;
        }

        for (i = 0; i < JAPANESE_FONT_KEYWORDS.length; i++) {
            for (j = 0; j < names.length; j++) {
                if (names[j].indexOf(JAPANESE_FONT_KEYWORDS[i]) !== -1) return true;
            }
        }
        return false;
    }

    /* 和文フォントの一覧は作るのが重いので使い回す / Building the list is heavy, so cache it */
    var japaneseFontsCache = null;

    /**
     * 環境にある和文フォントの一覧を返す
     * @returns {TextFont[]} 和文フォントの配列
     */
    function getJapaneseFonts() {
        if (japaneseFontsCache) return japaneseFontsCache;
        japaneseFontsCache = [];
        for (var i = 0; i < allFonts.length; i++) {
            if (isJapaneseFont(allFonts[i])) japaneseFontsCache.push(allFonts[i]);
        }
        return japaneseFontsCache;
    }

    // -----------------------------------------
    // 選択テキスト / Selected text
    // -----------------------------------------

    /**
     * 選択オブジェクトから TextRange を集める
     * @param {Object[]} selection - 選択オブジェクトの配列
     * @returns {TextRange[]} TextRange の配列
     */
    function collectTextRanges(selection) {
        var collected = [];
        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (!item) continue;
            if (item.typename === "TextRange") collected.push(item);
            else if (item.typename === "TextFrame") collected.push(item.textRange);
        }
        return collected;
    }

    /**
     * 英数字以外（改行を除く）を含むか判定する
     * @param {TextRange[]} targetRanges - 判定する TextRange
     * @returns {boolean} 含むとき true
     */
    function containsNonAlphanumeric(targetRanges) {
        for (var i = 0; i < targetRanges.length; i++) {
            var text = targetRanges[i].contents;
            for (var j = 0; j < text.length; j++) {
                var oneChar = text.charAt(j);
                if (oneChar === "\r" || oneChar === "\n") continue;
                if (!/[A-Za-z0-9]/.test(oneChar)) return true;
            }
        }
        return false;
    }

    /**
     * 手動カーニング値を読む（読めない環境では0）
     * @param {Object} character - 対象の文字
     * @returns {number} カーニング値（1/1000em）
     */
    function getCharacterKerning(character) {
        /* 自動カーニング中の文字では Error 9551 になる / Reading kerning throws 9551 while auto-kerning is on */
        try {
            var kerning = character.kerning;
            return (typeof kerning === "number") ? kerning : 0;
        } catch (e) {
            return 0;
        }
    }

    var textRanges = collectTextRanges(doc.selection);
    var selectedTextFrames = getSelectionTextFrames(doc.selection);
    if (textRanges.length === 0) { alert(getLabel("alert.selectTextRange")); return; }

    // -----------------------------------------
    // 文字属性のスナップショット / Character snapshots
    // -----------------------------------------

    /* 元の文字属性の控え / The attributes each character started with */
    var charSnapshots = [];

    /**
     * 選択中の全文字の属性を控える
     * @returns {void}
     */
    function takeCharSnapshots() {
        charSnapshots = [];
        for (var r = 0; r < textRanges.length; r++) {
            var textRange = textRanges[r];
            for (var c = 0; c < textRange.length; c++) {
                var character = textRange.characters[c];
                var attributes = character.characterAttributes;
                charSnapshots.push({
                    character: character,
                    baselineShift: attributes.baselineShift,
                    horizontalScale: attributes.horizontalScale,
                    verticalScale: attributes.verticalScale,
                    rotation: attributes.rotation,
                    kerning: getCharacterKerning(character),
                    tracking: (typeof attributes.tracking === "number") ? attributes.tracking : 0,
                    textFont: attributes.textFont
                });
            }
        }
    }

    /**
     * 控えた文字属性に戻す
     * @returns {void}
     */
    function restoreCharSnapshots() {
        for (var i = 0; i < charSnapshots.length; i++) {
            var snapshot = charSnapshots[i];
            var attributes = snapshot.character.characterAttributes;
            attributes.baselineShift = snapshot.baselineShift;
            attributes.horizontalScale = snapshot.horizontalScale;
            attributes.verticalScale = snapshot.verticalScale;
            attributes.rotation = snapshot.rotation;
            /* カーニング・トラッキング・フォントは書き込めない環境がある / These three are not writable everywhere */
            try {
                attributes.kerningMethod = AutoKernType.NOAUTOKERN;
                snapshot.character.kerning = snapshot.kerning;
            } catch (e) { }
            try {
                attributes.tracking = snapshot.tracking;
            } catch (e) { }
            try {
                if (snapshot.textFont) attributes.textFont = snapshot.textFont;
            } catch (e) { }
        }
    }

    // -----------------------------------------
    // 文字回転によるトラッキング補正 / Tracking compensation for rotation
    // -----------------------------------------

    /**
     * 1文字の回転量から、前後に足すトラッキング量を求める
     * @param {number} rotationDeg - 回転角（度）
     * @param {string} charContent - その文字の内容
     * @returns {{previousGap: number, currentGap: number}} 前の文字と自身に足す量（1/1000em）
     */
    function calcRotationTrackingPair(rotationDeg, charContent) {
        var sine = Math.sin(rotationDeg * (Math.PI / 180));
        var factorAfter = (rotationDeg > 0) ? ROTATION_GAP_FACTOR_AFTER_POSITIVE : ROTATION_GAP_FACTOR_AFTER_NEGATIVE;
        var currentGap = Math.abs(sine) * factorAfter * -1;
        var previousGap = sine * ROTATION_GAP_FACTOR_BEFORE * -1;

        /* 大文字は字形が大きいぶん前後の効き方が変わる / Uppercase letters need a different balance */
        if (charContent && charContent === charContent.toUpperCase() && charContent !== charContent.toLowerCase()) {
            previousGap = previousGap * ROTATION_GAP_UPPERCASE_BEFORE;
            currentGap = currentGap * ROTATION_GAP_UPPERCASE_AFTER;
        }
        return { previousGap: previousGap, currentGap: currentGap };
    }

    /**
     * 文字の回転角を読む
     * @param {Object} character - 対象の文字
     * @returns {number} 回転角（度）
     */
    function getCharacterRotation(character) {
        /* バージョンによって rotation を持たないことがある / Some builds do not expose rotation */
        try {
            if (typeof character.rotation === "number") return character.rotation;
            if (character.characterAttributes && typeof character.characterAttributes.rotation === "number") {
                return character.characterAttributes.rotation;
            }
        } catch (e) { }
        return 0;
    }

    /**
     * 1つの TextRange にトラッキング補正をかける
     * @param {TextRange} textRange - 対象の TextRange
     * @returns {void}
     */
    function applyRotationTrackingToRange(textRange) {
        var characters = textRange.characters;
        var charCount = characters.length;
        if (charCount <= 1) return;

        var trackingDeltas = [];
        var i;
        for (i = 0; i < charCount; i++) trackingDeltas[i] = 0;

        /* 1周目：回転している文字が前後に必要とする量を足し合わせる / Pass 1: accumulate the required gaps */
        for (i = 0; i < charCount; i++) {
            var rotationDeg = getCharacterRotation(characters[i]);
            if (Math.abs(rotationDeg) <= 1.0) continue;
            var gapPair = calcRotationTrackingPair(rotationDeg, characters[i].contents);
            trackingDeltas[i] += gapPair.currentGap;
            if (i > 0) trackingDeltas[i - 1] += gapPair.previousGap;
        }

        /* 2周目：まとめて書き込む / Pass 2: write the values */
        for (i = 0; i < charCount; i++) {
            if (Math.abs(trackingDeltas[i]) <= 0.5) continue;
            var content = characters[i].contents;
            if (content === "\r" || content === "\n") continue;
            /* トラッキングを書き込めない環境がある / Tracking is not writable everywhere */
            try {
                characters[i].characterAttributes.tracking = trackingDeltas[i];
            } catch (e) { }
        }
    }

    /**
     * 選択中のすべての TextRange にトラッキング補正をかける
     * @param {TextRange[]} targetRanges - 対象の TextRange
     * @returns {void}
     */
    function applyRotationTrackingToRanges(targetRanges) {
        for (var r = 0; r < targetRanges.length; r++) {
            applyRotationTrackingToRange(targetRanges[r]);
        }
    }

    // -----------------------------------------
    // ランダム適用 / Randomization
    // -----------------------------------------

    /**
     * seedから同じ並びを再現できる乱数生成器を作る
     * @param {number} seed - 乱数の種
     * @returns {function} 0以上1未満の乱数を返す関数
     */
    function createSeededRandom(seed) {
        var state = seed >>> 0;
        return function () {
            state = (1664525 * state + 1013904223) >>> 0;
            return state / 4294967296;
        };
    }

    /**
     * -1〜1の乱数を返す
     * @param {function} random - 乱数生成器
     * @returns {number} -1以上1未満の値
     */
    function randomSigned(random) {
        return random() * 2.0 - 1.0;
    }

    /**
     * 入れ替え先のフォント一覧を決める
     * @param {Object} options - 適用オプション
     * @returns {TextFont[]|null} フォントの配列（入れ替えないときは null）
     */
    function resolveFontPool(options) {
        if (!options.randomFont) return null;
        var fontPool = options.japaneseOnly ? options.japaneseFonts : options.allFonts;
        return (fontPool && fontPool.length > 0) ? fontPool : null;
    }

    /**
     * 選択中の各文字にランダムな文字タッチを適用する
     * @param {{baselinePt: number, scalePercent: number, rotationDeg: number, kerningEm: number}} amounts - 各項目の最大量
     * @param {number} seed - 乱数の種
     * @param {Object} options - フォント入れ替えや「犯行声明文」風などのオプション
     * @returns {void}
     */
    function applyRandomTouch(amounts, seed, options) {
        var random = createSeededRandom(seed);
        var fontPool = resolveFontPool(options);

        /* 「犯行声明文」風のトラッキングは固定値で上書きする / Ransom-note tracking overrides the random spacing */
        var ransomTrackEnabled = (options.ransomTrack !== false);
        var useFixedTracking = (options.ransom && ransomTrackEnabled) || options.previewRansomTracking;
        var fixedTracking = (typeof options.ransomTrackValue === "number") ? options.ransomTrackValue : RANSOM_TRACKING_DEFAULT;
        var addFixedToExisting = (options.rotationTracking !== false);

        for (var i = 0; i < charSnapshots.length; i++) {
            var snapshot = charSnapshots[i];
            var attributes = snapshot.character.characterAttributes;
            var content = snapshot.character.contents;

            if (fontPool && !(content === "\r" || content === "\n" || content === " ")) {
                /* 環境にないフォントは適用できない / A font missing from the system cannot be applied */
                try {
                    attributes.textFont = fontPool[Math.floor(random() * fontPool.length)];
                } catch (e) { }
            }

            attributes.baselineShift = randomSigned(random) * amounts.baselinePt;
            attributes.rotation = snapshot.rotation + (randomSigned(random) * amounts.rotationDeg);

            var scaleFactor = 1.0 + (randomSigned(random) * (amounts.scalePercent / 100.0));
            attributes.horizontalScale = snapshot.horizontalScale * scaleFactor;
            attributes.verticalScale = snapshot.verticalScale * scaleFactor;

            /* カーニングとトラッキングは書き込めない環境がある / Kerning and tracking are not writable everywhere */
            try {
                attributes.kerningMethod = AutoKernType.NOAUTOKERN;
                snapshot.character.kerning = snapshot.kerning + (randomSigned(random) * amounts.kerningEm);
            } catch (e) { }
            try {
                if (!useFixedTracking) attributes.tracking = snapshot.tracking;
                else if (addFixedToExisting) attributes.tracking = snapshot.tracking + fixedTracking;
                else attributes.tracking = fixedTracking;
            } catch (e) { }
        }

        /* 回転で広がった見た目の幅を字間で打ち消す / Offset the width the rotation added */
        if (Math.abs(amounts.rotationDeg) > 0.0001 && options.rotationTracking !== false && !useFixedTracking) {
            applyRotationTrackingToRanges(textRanges);
        }

        if (!options.skipRedraw) app.redraw();
    }

    /**
     * 文字列を数値に変換する
     * @param {string} text - 入力文字列
     * @returns {number|null} 数値（数値にならないときは null）
     */
    function parseNumber(text) {
        var value = parseFloat(text);
        return isNaN(value) ? null : value;
    }

    takeCharSnapshots();
    var seed = (new Date()).getTime() & 0xffffffff;

    /* 前回の実行が残した背景を消してから始める / Clear any background left by a previous run */
    clearRansomBgRects();

    // -----------------------------------------
    // プレビュー管理 / Preview state
    // -----------------------------------------

    /* プレビューは取り消し（アンドゥ）ではなく上書きで反映する。キャンセル時だけ控えに戻す
       The preview overwrites the attributes instead of using undo; Cancel restores the snapshots */
    var previewApplied = false;
    /* OKで閉じたときは onClose の後始末をしない / A close via OK must not roll the result back */
    var closedByOK = false;
    /* リセット直後にOKされたら、もう一度ランダム化しない / OK right after Reset must not re-randomize */
    var didReset = false;

    /**
     * プレビューを適用し、キャンセル時に戻せるよう記録する
     * @param {function} previewFn - 文字属性を書き換える処理
     * @returns {void}
     */
    function runPreview(previewFn) {
        /* 入力のたびに失敗をアラートしない / Do not alert on every keystroke */
        try { previewFn(); } catch (e) { }
        previewApplied = true;
    }

    /**
     * プレビューを取り消して元の文字属性に戻す
     * @returns {void}
     */
    function cancelPreview() {
        if (!previewApplied) return;
        try { restoreCharSnapshots(); } catch (e) { }
        previewApplied = false;
    }

    // -----------------------------------------
    // ダイアログ / Dialog
    // -----------------------------------------

    var dlg = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = ["fill", "top"];

    /* 前回閉じた位置を復元する（未保存のときは 0 が返る）/ Restore the last position; an unset key returns 0 */
    var savedDialogX = app.preferences.getIntegerPreference(PREF_KEY_DIALOG_X);
    var savedDialogY = app.preferences.getIntegerPreference(PREF_KEY_DIALOG_Y);
    if (savedDialogX !== 0 || savedDialogY !== 0) dlg.location = [savedDialogX, savedDialogY];

    /* 文字タッチ / Touch panel */
    var touchPanel = dlg.add("panel", undefined, getLabel("panel.touch"));
    touchPanel.orientation = "column";
    touchPanel.alignChildren = ["fill", "top"];
    touchPanel.margins = PANEL_MARGINS;

    /* フォントと「犯行声明文」風を横に並べる / Font and ransom-note panels sit side by side */
    var fontRowGroup = dlg.add("group");
    fontRowGroup.orientation = "row";
    fontRowGroup.alignChildren = ["fill", "top"];

    var fontPanel = fontRowGroup.add("panel", undefined, getLabel("panel.font"));
    fontPanel.orientation = "column";
    fontPanel.alignChildren = ["left", "top"];
    fontPanel.margins = PANEL_MARGINS;
    fontPanel.alignment = ["fill", "top"];

    var chkRandomFont = fontPanel.add("checkbox", undefined, getLabel("checkbox.randomFont"));
    chkRandomFont.helpTip = getLabel("tooltip.randomFont");
    var chkJapaneseOnly = fontPanel.add("checkbox", undefined, getLabel("checkbox.japaneseOnly"));
    chkJapaneseOnly.helpTip = getLabel("tooltip.japaneseOnly");

    var ransomPanel = fontRowGroup.add("panel", undefined, getLabel("panel.ransom"));
    ransomPanel.orientation = "column";
    ransomPanel.alignChildren = ["left", "top"];
    ransomPanel.margins = PANEL_MARGINS;
    ransomPanel.alignment = ["fill", "top"];

    var chkRansomEnabled = ransomPanel.add("checkbox", undefined, getLabel("checkbox.ransomEnabled"));
    chkRansomEnabled.helpTip = getLabel("tooltip.ransomEnabled");

    var ransomTrackingGroup = ransomPanel.add("group");
    ransomTrackingGroup.orientation = "row";
    ransomTrackingGroup.alignChildren = ["left", "center"];

    var chkRansomTracking = ransomTrackingGroup.add("checkbox", undefined, getLabel("checkbox.ransomTracking"));
    chkRansomTracking.helpTip = getLabel("tooltip.ransomTracking");
    var edtRansomTracking = ransomTrackingGroup.add("edittext", undefined, String(RANSOM_TRACKING_DEFAULT));
    edtRansomTracking.helpTip = getLabel("tooltip.ransomTrackingValue");
    edtRansomTracking.characters = RANSOM_FIELD_CHARS;

    chkRandomFont.value = false;
    chkJapaneseOnly.value = false;
    chkRansomEnabled.value = false;
    /* トラッキング調整は既定でON。ただし「有効」がOFFの間は操作できない
       Tracking is on by default but stays dimmed until the style is enabled */
    chkRansomTracking.value = true;
    chkRansomTracking.enabled = false;
    edtRansomTracking.enabled = false;

    /**
     * 「ランダム」がOFFのときは「和文フォントに限定」を無効にする
     * @returns {void}
     */
    function updateFontOptionState() {
        chkJapaneseOnly.enabled = chkRandomFont.value;
        if (!chkRandomFont.value) chkJapaneseOnly.value = false;
    }
    updateFontOptionState();

    /**
     * 英数字以外を含む選択では「和文フォントに限定」を自動でONにする
     * @returns {void}
     */
    function autoEnableJapaneseOnly() {
        if (chkJapaneseOnly.enabled && containsNonAlphanumeric(textRanges)) chkJapaneseOnly.value = true;
    }

    /* ベースラインの単位は環境設定の「文字」の単位に従う / The baseline unit follows the "text/asianunits" preference */
    var baselineUnit = getUnitInfo("text/asianunits");
    var baselineUnitLabel = baselineUnit.label;
    var baselinePointsPerUnit = baselineUnit.pointsPerUnit;

    /* 初期値は文字サイズの1/12、スライダーの上限は文字サイズ / The default is 1/12 of the type size; the slider tops out at the type size */
    var firstCharSizePt = (charSnapshots.length > 0) ? charSnapshots[0].character.characterAttributes.size : 0;
    var defaultBaselineValue = firstCharSizePt ? Math.round(Math.round(firstCharSizePt / 12) / baselinePointsPerUnit) : 0;
    if (isNaN(defaultBaselineValue)) defaultBaselineValue = 0;
    var baselineSliderMax = firstCharSizePt
        ? Math.max(1, Math.round(Math.max(1, Math.round(firstCharSizePt)) / baselinePointsPerUnit))
        : BASELINE_SLIDER_MAX_FALLBACK;

    /**
     * 「チェックボックス＋項目名＋数値欄＋単位＋スライダー」の行を作る
     * @param {Panel} parentPanel - 行を置くパネル
     * @param {string} labelKey - 項目名のキー（tooltipのキーにも使う）
     * @param {string} unitLabel - 単位の表示文字列
     * @param {number} defaultValue - 数値欄の初期値
     * @param {number} minValue - スライダーの下限
     * @param {number} maxValue - スライダーの上限
     * @returns {{toggle: Checkbox, label: StaticText, field: EditText, unit: StaticText, slider: Slider}} 作成したコントロール
     */
    function addTouchRow(parentPanel, labelKey, unitLabel, defaultValue, minValue, maxValue) {
        var row = parentPanel.add("group");

        var toggle = row.add("checkbox", undefined, "");
        toggle.helpTip = getLabel("tooltip." + labelKey + "Enabled");
        toggle.value = true;
        toggle.preferredSize.width = TOGGLE_WIDTH;

        var label = row.add("statictext", undefined, labelText("fieldLabel." + labelKey));

        var field = row.add("edittext", undefined, String(defaultValue));
        field.characters = VALUE_FIELD_CHARS;
        field.helpTip = getLabel("tooltip." + labelKey);

        var unit = row.add("statictext", undefined, unitLabel);

        var slider = row.add("slider", undefined, defaultValue, minValue, maxValue);
        slider.preferredSize.width = SLIDER_WIDTH;
        slider.helpTip = getLabel("tooltip." + labelKey);

        return { toggle: toggle, label: label, field: field, unit: unit, slider: slider };
    }

    var baselineRow = addTouchRow(touchPanel, "baseline", baselineUnitLabel, defaultBaselineValue, 0, baselineSliderMax);
    var scaleRow = addTouchRow(touchPanel, "scale", "%", DEFAULT_SCALE_PERCENT, 0, SCALE_SLIDER_MAX);
    var kerningRow = addTouchRow(touchPanel, "kerning", "em", DEFAULT_KERNING_EM, KERNING_SLIDER_MIN, KERNING_SLIDER_MAX);
    var rotationRow = addTouchRow(touchPanel, "rotation", "°", DEFAULT_ROTATION_DEG, 0, ROTATION_SLIDER_MAX);

    /**
     * 複数のラベルの幅を、いちばん広いものに揃える
     * @param {StaticText[]} controls - 幅を揃えるラベル
     * @returns {void}
     */
    function alignLabelWidths(controls) {
        var widest = 0;
        var i;
        for (i = 0; i < controls.length; i++) {
            if (controls[i].preferredSize.width > widest) widest = controls[i].preferredSize.width;
        }
        for (i = 0; i < controls.length; i++) {
            controls[i].preferredSize.width = widest;
            controls[i].justify = "left";
        }
    }
    alignLabelWidths([baselineRow.label, scaleRow.label, kerningRow.label, rotationRow.label]);
    alignLabelWidths([baselineRow.unit, scaleRow.unit, kerningRow.unit, rotationRow.unit]);

    /* 文字タッチパネル下部：左に一括ON／OFF、右にトラッキング補正
       Bottom of the touch panel: bulk toggles on the left, tracking compensation on the right */
    var touchButtonRow = touchPanel.add("group");
    touchButtonRow.orientation = "row";
    touchButtonRow.alignChildren = ["fill", "center"];
    touchButtonRow.alignment = "fill";
    touchButtonRow.margins = [0, TOUCH_BUTTON_ROW_TOP_MARGIN, 0, 0];

    var touchButtonLeftGroup = touchButtonRow.add("group");
    touchButtonLeftGroup.orientation = "row";
    touchButtonLeftGroup.alignChildren = ["left", "center"];

    var btnAllOn = touchButtonLeftGroup.add("button", undefined, getLabel("button.allOn"));
    btnAllOn.helpTip = getLabel("tooltip.allOn");
    btnAllOn.preferredSize = SMALL_BUTTON_SIZE;
    var btnAllOff = touchButtonLeftGroup.add("button", undefined, getLabel("button.allOff"));
    btnAllOff.helpTip = getLabel("tooltip.allOff");
    btnAllOff.preferredSize = SMALL_BUTTON_SIZE;

    var touchButtonSpacer = touchButtonRow.add("group");
    touchButtonSpacer.alignment = ["fill", "fill"];
    touchButtonSpacer.minimumSize.width = 0;

    var touchButtonRightGroup = touchButtonRow.add("group");
    touchButtonRightGroup.orientation = "row";
    touchButtonRightGroup.alignChildren = ["right", "center"];

    var chkRotationTracking = touchButtonRightGroup.add("checkbox", undefined, getLabel("checkbox.rotationTracking"));
    chkRotationTracking.helpTip = getLabel("tooltip.rotationTracking");
    chkRotationTracking.value = true;

    /* ズーム / Zoom row */
    var zoomGroup = dlg.add("group");
    zoomGroup.orientation = "row";
    zoomGroup.alignChildren = ["center", "center"];
    zoomGroup.margins = [0, 0, 0, 0];

    zoomGroup.add("statictext", undefined, labelText("fieldLabel.zoom"));
    var sldZoom = zoomGroup.add("slider", undefined, clampZoomPercent(originalZoomFactor * 100), ZOOM_MIN_PERCENT, ZOOM_MAX_PERCENT);
    sldZoom.helpTip = getLabel("tooltip.zoom");
    sldZoom.preferredSize.width = ZOOM_SLIDER_WIDTH;

    var chkLightMode = zoomGroup.add("checkbox", undefined, getLabel("checkbox.lightMode"));
    chkLightMode.helpTip = getLabel("tooltip.lightMode");
    chkLightMode.value = false;

    /* ボタンエリア / Button row */
    var btnRowGroup = dlg.add("group");
    btnRowGroup.orientation = "row";
    btnRowGroup.alignChildren = ["fill", "center"];
    btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, 0];

    var btnLeftGroup = btnRowGroup.add("group");
    btnLeftGroup.orientation = "row";
    btnLeftGroup.alignChildren = ["left", "center"];

    var btnRandomize = btnLeftGroup.add("button", undefined, getLabel("button.randomize"));
    btnRandomize.helpTip = getLabel("tooltip.randomize");
    var btnReset = btnLeftGroup.add("button", undefined, getLabel("button.reset"));
    btnReset.helpTip = getLabel("tooltip.reset");

    var spacer = btnRowGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 0;

    var btnRightGroup = btnRowGroup.add("group");
    btnRightGroup.orientation = "row";
    btnRightGroup.alignChildren = ["right", "center"];
    var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
    var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

    // -----------------------------------------
    // 入力値の同期 / Value syncing
    // -----------------------------------------

    /**
     * 値を範囲内に収める
     * @param {number} value - 対象の値
     * @param {number} minValue - 下限
     * @param {number} maxValue - 上限
     * @returns {number} 範囲内に収めた値
     */
    function clamp(value, minValue, maxValue) {
        if (value < minValue) return minValue;
        if (value > maxValue) return maxValue;
        return value;
    }

    /**
     * 数値欄の値をスライダーに反映する
     * @param {EditText} field - 数値欄
     * @param {Slider} slider - スライダー
     * @returns {void}
     */
    function syncSliderFromField(field, slider) {
        var value = parseNumber(field.text);
        if (value === null) return;
        slider.value = clamp(value, slider.minvalue, slider.maxvalue);
    }

    /**
     * スライダーの値を数値欄に反映する
     * @param {Slider} slider - スライダー
     * @param {EditText} field - 数値欄
     * @returns {void}
     */
    function syncFieldFromSlider(slider, field) {
        field.text = String(Math.round(slider.value));
    }

    /**
     * 数値欄を↑↓キーで増減できるようにする
     * @param {EditText} field - 対象の数値欄
     * @param {Slider} slider - 連動するスライダー（無いときは null）
     * @param {boolean} allowNegative - マイナス値を許すかどうか
     * @returns {void}
     */
    function changeValueByArrowKey(field, slider, allowNegative) {
        field.addEventListener("keydown", function (event) {
            if (!(event && (event.keyName === "Up" || event.keyName === "Down"))) return;

            var value = Number(field.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;
            if (keyboard.shiftKey) {
                /* Shiftキー押下時は10の倍数にスナップ / Shift snaps to multiples of ten */
                delta = 10;
                value = (event.keyName === "Up")
                    ? Math.ceil((value + 1) / delta) * delta
                    : Math.floor((value - 1) / delta) * delta;
            } else {
                /* Optionキー押下時は0.1単位 / Option steps by 0.1 */
                delta = keyboard.altKey ? 0.1 : 1;
                value += (event.keyName === "Up") ? delta : -delta;
            }
            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            field.text = String(value);
            if (slider) syncSliderFromField(field, slider);
            requestPreview();
        });
    }

    // -----------------------------------------
    // プレビューの実行 / Running the preview
    // -----------------------------------------

    /* 入力のたびに再描画すると重いので、少し待ってからまとめて反映する
       Redrawing on every keystroke is heavy, so the preview is debounced */
    var previewTaskId = null;

    /* scheduleTask はグローバルに置いた関数名でしか呼べない / scheduleTask can only call a global by name */
    $.global.__AutoTouchTypePreview = function () {
        previewTaskId = null;
        updatePreview();
    };

    /**
     * プレビューの更新を予約する
     * @returns {void}
     */
    function requestPreview() {
        /* scheduleTask が使えない環境ではその場で更新する / Fall back to an immediate update */
        try {
            if (previewTaskId != null) app.cancelTask(previewTaskId);
            previewTaskId = app.scheduleTask("$.global.__AutoTouchTypePreview()", PREVIEW_DELAY_MS, false);
        } catch (e) {
            updatePreview();
        }
    }

    /**
     * 文字タッチ4項目の入力値を読む
     * @returns {{baselinePt: number|null, scalePercent: number|null, rotationDeg: number|null, kerningEm: number|null}} 各項目の最大量（空欄は null）
     */
    function readTouchAmounts() {
        var baselineValue = baselineRow.toggle.value ? parseNumber(baselineRow.field.text) : 0;
        return {
            baselinePt: (baselineValue === null) ? null : baselineValue * baselinePointsPerUnit,
            scalePercent: scaleRow.toggle.value ? parseNumber(scaleRow.field.text) : 0,
            rotationDeg: rotationRow.toggle.value ? parseNumber(rotationRow.field.text) : 0,
            kerningEm: kerningRow.toggle.value ? parseNumber(kerningRow.field.text) : 0
        };
    }

    /**
     * 空欄の項目があるか判定する
     * @param {Object} amounts - readTouchAmounts() の戻り値
     * @returns {boolean} 空欄があるとき true
     */
    function hasEmptyAmount(amounts) {
        return amounts.baselinePt === null || amounts.scalePercent === null ||
            amounts.rotationDeg === null || amounts.kerningEm === null;
    }

    /**
     * 適用オプションを組み立てる
     * @param {boolean} forPreview - プレビュー用かどうか
     * @returns {Object} applyRandomTouch() に渡すオプション
     */
    function buildTouchOptions(forPreview) {
        var ransomOn = chkRansomEnabled.value && chkRansomTracking.value;
        return {
            randomFont: chkRandomFont.value,
            japaneseOnly: chkJapaneseOnly.value,
            /* 背景の生成はプレビューしない（OK時のみ）/ The backgrounds are created on OK only */
            ransom: forPreview ? false : chkRansomEnabled.value,
            ransomTrack: chkRansomTracking.value,
            /* プレビューではトラッキングの固定値だけ反映する / The preview only reflects the fixed tracking */
            previewRansomTracking: forPreview && ransomOn,
            ransomTrackValue: null,
            rotationTracking: chkRotationTracking.value,
            allFonts: allFonts,
            japaneseFonts: null,
            skipRedraw: forPreview
        };
    }

    /**
     * 和文フォントに限定するとき、一覧をオプションに入れる
     * @param {Object} options - 適用オプション
     * @returns {boolean} 続行できるとき true（和文フォントが無いときは false）
     */
    function fillJapaneseFonts(options) {
        if (!(options.randomFont && options.japaneseOnly)) return true;
        options.japaneseFonts = getJapaneseFonts();
        if (options.japaneseFonts.length > 0) return true;
        alert(getLabel("alert.noJapaneseFonts"));
        return false;
    }

    /**
     * 現在の設定でプレビューを更新する
     * @returns {void}
     */
    function updatePreview() {
        didReset = false;

        var amounts = readTouchAmounts();
        if (hasEmptyAmount(amounts)) return;

        var options = buildTouchOptions(true);
        if (!fillJapaneseFonts(options)) return;
        if (options.previewRansomTracking) {
            var trackingValue = parseNumber(edtRansomTracking.text);
            options.ransomTrackValue = (trackingValue === null) ? RANSOM_TRACKING_DEFAULT : trackingValue;
        }

        runPreview(function () {
            applyRandomTouch(amounts, seed, options);
        });
        app.redraw();
    }

    // -----------------------------------------
    // イベント / Events
    // -----------------------------------------

    var touchRows = [baselineRow, scaleRow, kerningRow, rotationRow];
    var allowNegativeRows = { kerning: true };

    /**
     * 文字タッチ4項目がすべてOFFか判定する
     * @returns {boolean} すべてOFFのとき true
     */
    function isTouchAllOff() {
        for (var i = 0; i < touchRows.length; i++) {
            if (touchRows[i].toggle.value) return false;
        }
        return true;
    }

    /**
     * 変化を生まない設定のときは「ランダム」ボタンを無効にする
     * @returns {void}
     */
    function updateRandomizeEnabled() {
        btnRandomize.enabled = !(isTouchAllOff() && !chkRandomFont.value);
    }

    /**
     * 文字タッチ4項目をまとめて切り替える
     * @param {boolean} enabled - ONにするとき true
     * @returns {void}
     */
    function setAllTouchToggles(enabled) {
        for (var i = 0; i < touchRows.length; i++) {
            touchRows[i].toggle.value = enabled;
        }
        updatePreview();
        updateRandomizeEnabled();
    }

    /* 数値欄・スライダー・チェックボックスの操作をプレビューにつなぐ
       Wire the fields, sliders and toggles to the preview */
    var rowKeys = ["baseline", "scale", "kerning", "rotation"];
    for (var rowIndex = 0; rowIndex < touchRows.length; rowIndex++) {
        (function (row, allowNegative) {
            row.field.onChanging = function () {
                syncSliderFromField(row.field, row.slider);
                requestPreview();
            };
            row.slider.onChanging = function () {
                syncFieldFromSlider(row.slider, row.field);
                requestPreview();
            };
            row.toggle.onClick = function () {
                requestPreview();
                updateRandomizeEnabled();
            };
            changeValueByArrowKey(row.field, row.slider, allowNegative);
            syncSliderFromField(row.field, row.slider);
        })(touchRows[rowIndex], !!allowNegativeRows[rowKeys[rowIndex]]);
    }

    edtRansomTracking.onChanging = function () { requestPreview(); };
    changeValueByArrowKey(edtRansomTracking, null, true);

    chkRotationTracking.onClick = function () { requestPreview(); };

    chkRandomFont.onClick = function () {
        updateFontOptionState();
        autoEnableJapaneseOnly();
        requestPreview();
        updateRandomizeEnabled();
    };

    chkJapaneseOnly.onClick = function () {
        requestPreview();
        updateRandomizeEnabled();
    };

    chkRansomEnabled.onClick = function () {
        chkRansomTracking.enabled = chkRansomEnabled.value;
        edtRansomTracking.enabled = chkRansomEnabled.value && chkRansomTracking.value;
        updateRandomizeEnabled();
        requestPreview();
    };

    chkRansomTracking.onClick = function () {
        edtRansomTracking.enabled = chkRansomEnabled.value && chkRansomTracking.value;
        requestPreview();
    };

    btnAllOn.onClick = function () { setAllTouchToggles(true); };
    btnAllOff.onClick = function () { setAllTouchToggles(false); };

    /* ドラッグ中の表示倍率の反映は重いので間引く / Applying the zoom while dragging is heavy, so throttle it */
    var lastZoomAppliedMs = 0;

    sldZoom.onChanging = function () {
        /* 軽量モードではドラッグ中は何もしない / Light mode skips the update while dragging */
        if (chkLightMode.value) return;
        var now = (new Date()).getTime();
        if (now - lastZoomAppliedMs < ZOOM_THROTTLE_MS) return;
        lastZoomAppliedMs = now;
        applyZoomPercent(this.value);
    };

    sldZoom.onChange = function () { applyZoomPercent(this.value); };
    chkLightMode.onClick = function () { applyZoomPercent(sldZoom.value); };

    /**
     * 文字属性をリセットし、控えも初期値にそろえる
     * @returns {void}
     */
    function resetCharAttributes() {
        for (var i = 0; i < charSnapshots.length; i++) {
            var snapshot = charSnapshots[i];
            var attributes = snapshot.character.characterAttributes;

            attributes.baselineShift = 0;
            attributes.horizontalScale = 100;
            attributes.verticalScale = 100;
            attributes.rotation = 0;
            snapshot.baselineShift = 0;
            snapshot.horizontalScale = 100;
            snapshot.verticalScale = 100;
            snapshot.rotation = 0;

            /* カーニングとトラッキングは書き込めない環境がある / Kerning and tracking are not writable everywhere */
            try {
                attributes.kerningMethod = AutoKernType.NOAUTOKERN;
                snapshot.character.kerning = 0;
                snapshot.kerning = 0;
            } catch (e) { }
            try {
                attributes.tracking = 0;
                snapshot.tracking = 0;
            } catch (e) { }
        }
    }

    /**
     * すべての文字のフォントを先頭文字のフォントにそろえる
     * @returns {void}
     */
    function unifyFontToFirstChar() {
        if (charSnapshots.length === 0) return;
        var firstFont = charSnapshots[0].character.characterAttributes.textFont;
        if (!firstFont) return;
        for (var i = 0; i < charSnapshots.length; i++) {
            var snapshot = charSnapshots[i];
            var content = snapshot.character.contents;
            if (content === "\r" || content === "\n") continue;
            /* 環境にないフォントは適用できない / A font missing from the system cannot be applied */
            try {
                snapshot.character.characterAttributes.textFont = firstFont;
                snapshot.textFont = firstFont;
            } catch (e) { }
        }
    }

    btnRandomize.onClick = function () {
        didReset = false;
        seed = (new Date()).getTime() & 0xffffffff;
        requestPreview();
        autoEnableJapaneseOnly();
        updateRandomizeEnabled();
    };

    btnReset.onClick = function () {
        /* チェックボックスの状態は変えず、文字属性だけ初期化する / Only the text attributes are reset */
        cancelPreview();
        clearRansomBgRects();
        resetCharAttributes();
        unifyFontToFirstChar();
        app.redraw();

        updateFontOptionState();
        autoEnableJapaneseOnly();
        updateRandomizeEnabled();
        didReset = true;
    };

    btnOK.onClick = function () {
        var amounts = readTouchAmounts();
        /* ベースラインの空欄は0として扱う / An empty baseline field counts as zero */
        if (amounts.baselinePt === null) amounts.baselinePt = 0;
        if (hasEmptyAmount(amounts)) {
            alert(getLabel("alert.enterNumber"));
            return;
        }

        var options = buildTouchOptions(false);
        if (options.ransom && options.ransomTrack) {
            options.ransomTrackValue = parseNumber(edtRansomTracking.text);
            if (options.ransomTrackValue === null) {
                alert(getLabel("alert.enterNumber"));
                return;
            }
        }
        if (!fillJapaneseFonts(options)) return;

        /* リセット直後でプレビューが無ければ、その結果をそのまま確定する
           Right after Reset with no preview, keep the reset result as is */
        if (!(didReset && !previewApplied)) {
            try {
                applyRandomTouch(amounts, seed, options);
            } catch (e) {
                /* 選択が変わって控えが無効になったときは取り直してやり直す
                   Re-derive the selection when the snapshots went stale */
                textRanges = collectTextRanges(doc.selection);
                selectedTextFrames = getSelectionTextFrames(doc.selection);
                takeCharSnapshots();
                applyRandomTouch(amounts, seed, options);
            }
        }

        /* 背景の長方形はOK時だけ作る / The background rectangles are created on OK only */
        if (options.ransom) createRansomBgRects(selectedTextFrames);
        else clearRansomBgRects();

        closedByOK = true;
        didReset = false;
        dlg.close(1);
    };

    btnCancel.onClick = function () {
        cancelPreview();
        clearRansomBgRects();
        restoreOriginalView();
        closedByOK = false;
        dlg.close(0);
    };

    dlg.onClose = function () {
        var location = dlg.location;
        if (location && location.length === 2) {
            app.preferences.setIntegerPreference(PREF_KEY_DIALOG_X, Math.round(location[0]));
            app.preferences.setIntegerPreference(PREF_KEY_DIALOG_Y, Math.round(location[1]));
        }
        /* OKで閉じたときは確定済みなので後始末しない / A close via OK keeps the committed result */
        if (closedByOK) return true;

        cancelPreview();
        clearRansomBgRects();
        restoreOriginalView();
        return true;
    };

    updatePreview();
    updateRandomizeEnabled();
    dlg.opacity = DIALOG_OPACITY;
    dlg.show();
})();
