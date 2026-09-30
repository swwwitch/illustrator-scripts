#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択しているオブジェクトの塗り色を、配置順（左→右、上→下）で抽出してスウォッチグループに登録します。
グループ・複合パス・テキストは再帰的に処理し、線色も対象にします。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwatchGroupFromSelection.md

### Overview

Extracts the fill colors of the selected objects in layout order (left to right, top to bottom) and registers them as a swatch group.
Groups, compound paths and text are processed recursively, and stroke colors are collected as well.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwatchGroupFromSelection.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SwatchGroupFromSelection";     /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-01-28";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-30";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SwatchGroupFromSelection.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SwatchGroupFromSelection.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var SWATCH_GROUP_BASE_NAME = "AutoGradient"; /* スウォッチグループ名（重複時は連番） / Swatch group name; numbered when taken */
    var SWATCH_BASE_NAME       = "AutoColor";    /* スウォッチ名（重複時は連番） / Swatch name; numbered when taken */

    // =========================================
    // カラー判定 / Color keys
    // =========================================

    /**
     * カラーが「なし」かを判定する
     * @param {Color} color - 判定するカラー
     * @returns {boolean} null または NoColor なら true
     */
    function isNoColor(color) {
        return (color == null) || (color.typename === "NoColor");
    }

    /**
     * 重複除去と対応付けに使う、カラーの簡易キーを作る
     * @param {Color} color - 対象のカラー
     * @returns {string} "RGB:255,0,0" のようなキー
     */
    function colorKey(color) {
        if (!color) return "null";
        var colorType = color.typename;
        /* スポット・パターン・グラデーションの参照先が読めないことがある / Referenced spot/pattern/gradient may be unreadable */
        try {
            if (colorType === "RGBColor") {
                return "RGB:" + [color.red, color.green, color.blue].join(",");
            }
            if (colorType === "CMYKColor") {
                return "CMYK:" + [color.cyan, color.magenta, color.yellow, color.black].join(",");
            }
            if (colorType === "GrayColor") {
                return "Gray:" + color.gray;
            }
            if (colorType === "SpotColor") {
                /* スポットはスポット名＋濃度 / Spot name plus tint */
                var spotName = (color.spot && color.spot.name) ? color.spot.name : "(spot)";
                return "Spot:" + spotName + ":" + color.tint;
            }
            if (colorType === "PatternColor") {
                var patternName = (color.pattern && color.pattern.name) ? color.pattern.name : "(pattern)";
                return "Pattern:" + patternName;
            }
            if (colorType === "GradientColor") {
                var gradientName = (color.gradient && color.gradient.name) ? color.gradient.name : "(gradient)";
                return "Gradient:" + gradientName;
            }
        } catch (e) { }
        return "Other:" + colorType;
    }

    // =========================================
    // 塗り・線の走査 / Paint traversal
    // =========================================

    /**
     * 塗りまたは線を1つ読み、「なし」でなければ visitor に渡す
     * @param {Function} visitor - function(item, owner, propName, color)
     * @param {PageItem} item - 対象のオブジェクト
     * @param {Object} owner - カラーを持つオブジェクト（パス本体、またはテキストの characterAttributes）
     * @param {string} propName - "fillColor" または "strokeColor"
     * @returns {void}
     */
    function visitPaint(visitor, item, owner, propName) {
        /* 読み書きできない塗り・線は無視 / Ignore paints that cannot be read or written */
        try {
            var color = owner[propName];
            if (isNoColor(color)) return;
            visitor(item, owner, propName, color);
        } catch (e) { }
    }

    /**
     * オブジェクトの塗りと線をたどる（グループ・複合パスは再帰、テキストは全体の文字属性）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {Function} visitor - function(item, owner, propName, color)
     * @returns {void}
     */
    function visitPaintTargets(item, visitor) {
        if (!item) return;
        var i;
        if (item.typename === "GroupItem") {
            for (i = 0; i < item.pageItems.length; i++) {
                visitPaintTargets(item.pageItems[i], visitor);
            }
            return;
        }
        if (item.typename === "CompoundPathItem") {
            for (i = 0; i < item.pathItems.length; i++) {
                visitPaintTargets(item.pathItems[i], visitor);
            }
            return;
        }
        if (item.typename === "TextFrame") {
            var textAttributes = item.textRange.characterAttributes;
            visitPaint(visitor, item, textAttributes, "fillColor");
            visitPaint(visitor, item, textAttributes, "strokeColor");
            return;
        }
        /* パスなど（filled / stroked を持つもの） / Paths and other painted items */
        if (item.filled) visitPaint(visitor, item, item, "fillColor");
        if (item.stroked) visitPaint(visitor, item, item, "strokeColor");
    }

    /**
     * 選択から色を「左→右、上→下」の順で集める（同じ色は最初の1つだけ）
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {Color[]} 重複を除いたカラー
     */
    function collectColorsFromSelection(selectedItems) {
        var entries = [];
        /* 位置はオブジェクトの左上（geometricBounds: [left, top, right, bottom]） / Position is the item's top-left */
        var collectEntry = function (item, owner, propName, color) {
            var bounds = item.geometricBounds;
            entries.push({ left: bounds[0], top: bounds[1], color: color });
        };
        for (var i = 0; i < selectedItems.length; i++) {
            visitPaintTargets(selectedItems[i], collectEntry);
        }

        /* 左→右（left 昇順）、上→下（top は上ほど大きいので降順） / Left to right, then top to bottom */
        entries.sort(function (a, b) {
            if (a.left < b.left) return -1;
            if (a.left > b.left) return 1;
            if (a.top > b.top) return -1;
            if (a.top < b.top) return 1;
            return 0;
        });

        var colors = [];
        var seenKeys = {};
        for (var j = 0; j < entries.length; j++) {
            var key = colorKey(entries[j].color);
            if (seenKeys[key]) continue;
            seenKeys[key] = true;
            colors.push(entries[j].color);
        }
        return colors;
    }

    /**
     * 選択中のオブジェクトの塗りと線を、対応するグローバルカラーに置き換える
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @param {Object} colorMap - カラーのキー → 置き換え先のカラー
     * @returns {void}
     */
    function applyGlobalColorsToSelection(selectedItems, colorMap) {
        var applyMappedColor = function (item, owner, propName, color) {
            var mappedColor = colorMap[colorKey(color)];
            if (mappedColor) owner[propName] = mappedColor;
        };
        for (var i = 0; i < selectedItems.length; i++) {
            visitPaintTargets(selectedItems[i], applyMappedColor);
        }
    }

    // =========================================
    // スウォッチ登録 / Swatch registration
    // =========================================

    /**
     * コレクションに同名の項目があるかを調べる
     * @param {Object} collection - doc.swatches や doc.swatchGroups
     * @param {string} name - 名前
     * @returns {boolean} あれば true
     */
    function hasNamedItem(collection, name) {
        /* getByName は見つからないと例外 / getByName throws when missing */
        try {
            collection.getByName(name);
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 重複しない名前を作る（"名前", "名前 1", "名前 2", …）
     * @param {string} baseName - 基本の名前
     * @param {Object} collection - 重複を調べるコレクション
     * @returns {string} 使われていない名前
     */
    function uniqueName(baseName, collection) {
        var name = baseName;
        for (var n = 1; hasNamedItem(collection, name); n++) {
            name = baseName + " " + n;
        }
        return name;
    }

    /**
     * カラーをグローバルカラー（プロセス）のスポットにして返す
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} baseColor - 元のカラー
     * @param {string} spotName - スポット名
     * @returns {Color} グローバルカラーの SpotColor（作れなければ元のカラー）
     */
    function toGlobalProcessColor(doc, baseColor, spotName) {
        /* 名前の衝突などで作れないときは元のカラーを使う / Fall back to the original color when the spot cannot be made */
        try {
            var spot = doc.spots.add();
            spot.name = spotName;
            spot.colorType = ColorModel.PROCESS;
            spot.color = baseColor;

            var spotColor = new SpotColor();
            spotColor.spot = spot;
            spotColor.tint = 100;
            return spotColor;
        } catch (e) {
            return baseColor;
        }
    }

    /**
     * カラーをグローバルカラーのスウォッチとして追加する
     * @param {Document} doc - 対象ドキュメント
     * @param {Color} color - 登録するカラー
     * @param {string} baseName - スウォッチ名の基本
     * @returns {Swatch} 追加したスウォッチ
     */
    function addSwatchForColor(doc, color, baseName) {
        var swatch = doc.swatches.add();
        var swatchName = uniqueName(baseName, doc.swatches);
        swatch.name = swatchName;
        swatch.color = toGlobalProcessColor(doc, color, swatchName);
        /* Swatch に selected が無い環境がある / Some versions have no Swatch.selected */
        try { swatch.selected = false; } catch (e) { }
        return swatch;
    }

    /**
     * スウォッチグループを作り、カラーを順に登録する
     * @param {Document} doc - 対象ドキュメント
     * @param {Color[]} colors - 登録するカラー
     * @returns {Object} 元のカラーのキー → 登録したグローバルカラー
     */
    function registerSwatchGroup(doc, colors) {
        var swatchGroup = doc.swatchGroups.add();
        swatchGroup.name = uniqueName(SWATCH_GROUP_BASE_NAME, doc.swatchGroups);

        var colorMap = {};
        for (var i = 0; i < colors.length; i++) {
            var swatch = addSwatchForColor(doc, colors[i], SWATCH_BASE_NAME);
            /* グループに入れられない種類は単独のまま / Leave swatches that cannot join the group */
            try { swatchGroup.addSwatch(swatch); } catch (e) { }
            if (swatch.color) {
                colorMap[colorKey(colors[i])] = swatch.color;
            }
        }
        return colorMap;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択の色をスウォッチグループに登録し、選択中のオブジェクトへグローバルカラーとして適用し直す
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            return;
        }
        var doc = app.activeDocument;
        var originalSelection = doc.selection;
        /* 文字ツールで文字を選択しているときは TextRange が返り、length は文字数になる / With characters selected by the Type tool, selection is a TextRange whose length is the character count */
        if (!originalSelection || originalSelection.typename === "TextRange" || originalSelection.length === 0) {
            return;
        }

        var colors = collectColorsFromSelection(originalSelection);
        if (colors.length < 2) {
            return;
        }

        /* 失敗しても通知せずに終える（元の仕様） / Fail silently, as before */
        try {
            var colorMap = registerSwatchGroup(doc, colors);
            applyGlobalColorsToSelection(originalSelection, colorMap);
        } catch (e) { }
    }

    main();

})();
