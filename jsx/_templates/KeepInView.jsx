#target illustrator

/*

### スクリプト名：

KeepInView.jsx（テンプレート）

### 概要：

- 生成・変更したオブジェクトが画面から外れたときだけ、見える位置へ表示を移す再利用テンプレート
- すでに見えているときは動かさないので、操作のたびに画面が揺れない
- 収まらないときはズームアウトのみ行い、拡大はしない（倍率が勝手に上がらない）

### 使い方：

1. 下の `var KeepInView = (function () { ... })();` のブロックまるごとを、対象スクリプトのIIFE内へコピーする
   （`#include` は使わない。各スクリプトは1ファイルで完結させる方針）
2. チェックボックスを置く。ラベルとツールチップは日英を内蔵しているので、文言を用意する必要はない

        var cbKeepInView = KeepInView.addCheckbox(grpViewOptions, { value: true });

3. 結果を作り終えたところで呼ぶ

        if (cbKeepInView.value) KeepInView.ensureVisible(createdItems, { doc: doc, fitRatio: 0.9 });

チェックボックスを自前で作る場合は `addCheckbox` を使わず、`ensureVisible` だけを呼ぶ。
`doc` を省略すると最前面のドキュメント、`fitRatio` を省略すると `DEFAULT_FIT_RATIO` が使われる。

### 使用例：

- `jsx/text/DynamicTextGenerator.jsx`（「結果を画面内に表示」）

### 更新履歴：

- v1.0.0 : 初版（テンプレート）

*/

// =========================================
// 再利用パーツ / Reusable part
//   ↓↓↓ ここから下のブロックを対象スクリプトへコピーして使う ↓↓↓
// =========================================
var KeepInView = (function () {

    // =========================================
    // 基本情報 / Basic info
    // =========================================
    var SCRIPT_NAME     = "KeepInView";                   /* スクリプト名 / script name */
    var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
    var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
    var SCRIPT_RELEASED = "2026-08-14";                   /* 最初のリリース日 / first release date */
    var SCRIPT_UPDATED  = "2026-08-14";                   /* 更新日 / last updated */

    // Released under the MIT license
    // http://opensource.org/licenses/mit-license.php

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    /* 結果が可視領域からはみ出したときに合わせる倍率（1で余白なし。小さいほど余白が増える）
       呼び出し側から options.fitRatio で上書きできる */
    var DEFAULT_FIT_RATIO = 0.9;

    // =========================================
    // ローカライズ / Localization
    // =========================================
    var LABELS = {
        checkbox: { ja: "結果を画面内に表示", en: "Keep the result in view" },
        tooltip: {
            ja: "変換した結果が画面から外れたときに、見える位置へ表示を移します。すでに見えているときは動かしません。",
            en: "Moves the view so the converted result stays visible. The view is left alone when it is already in sight."
        }
    };

    /**
     * UI言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    /**
     * 内蔵ラベルを取り出す
     * @param {string} key - "checkbox" または "tooltip"
     * @param {string} lang - "ja" または "en"（省略時はUI言語）
     * @returns {string} ラベル文字列（見つからない場合は空文字）
     */
    function getLabel(key, lang) {
        var entry = LABELS[key];
        if (!entry) return "";
        if (!lang) lang = getCurrentLang();
        return (entry[lang] != null) ? entry[lang] : entry.en;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 「結果を画面内に表示」チェックボックスを作る（ラベルとツールチップは内蔵）
     * @param {Group|Panel|Window} parent - 追加先のコンテナ
     * @param {object} options - value: 初期値（既定 true）／lang: 表示言語／text: ラベルの差し替え
     * @returns {Checkbox} 作成したチェックボックス
     */
    function addCheckbox(parent, options) {
        if (!options) options = {};

        var lang = options.lang || getCurrentLang();
        var checkbox = parent.add('checkbox', undefined, options.text || getLabel('checkbox', lang));
        checkbox.helpTip = getLabel('tooltip', lang);
        /* 明示的に false を渡したときだけOFFで始める / only an explicit false starts it unchecked */
        checkbox.value = (options.value !== false);
        return checkbox;
    }

    /**
     * 複数アイテムを囲む外接範囲を求める
     * @param {PageItem[]} items - 対象アイテム
     * @returns {{left: number, top: number, right: number, bottom: number}|null} 外接範囲（求められない場合は null）
     */
    function getItemsBounds(items) {
        var bounds = null;

        for (var i = 0; i < items.length; i++) {
            var itemBounds;
            try {
                itemBounds = items[i].visibleBounds; // [left, top, right, bottom]
            } catch (e) {
                continue;
            }
            if (bounds === null) {
                bounds = { left: itemBounds[0], top: itemBounds[1], right: itemBounds[2], bottom: itemBounds[3] };
                continue;
            }
            if (itemBounds[0] < bounds.left) bounds.left = itemBounds[0];
            if (itemBounds[1] > bounds.top) bounds.top = itemBounds[1];
            if (itemBounds[2] > bounds.right) bounds.right = itemBounds[2];
            if (itemBounds[3] < bounds.bottom) bounds.bottom = itemBounds[3];
        }
        return bounds;
    }

    /**
     * 結果が可視領域に収まっていなければ、見えるように表示位置とズームを合わせる
     * すでに見えているときは何もしないので、操作のたびに画面が動くことはない
     * @param {PageItem[]} items - 見えるようにしたいアイテム
     * @param {object} options - doc: 対象ドキュメント（省略時は最前面）／fitRatio: 収めるときの倍率
     * @returns {boolean} 表示を動かしたら true
     */
    function ensureVisible(items, options) {
        if (!items || items.length === 0) return false;
        if (!options) options = {};

        var bounds = getItemsBounds(items);
        if (bounds === null) return false;

        var activeView;
        try {
            var targetDoc = options.doc || app.activeDocument;
            activeView = targetDoc.views[0];
        } catch (e) {
            return false;
        }
        if (!activeView) return false;

        var viewBounds = activeView.bounds; // [left, top, right, bottom]
        /* すでに全体が見えているなら動かさない / leave the view alone when everything is already visible */
        if (bounds.left >= viewBounds[0] && bounds.right <= viewBounds[2] &&
            bounds.top <= viewBounds[1] && bounds.bottom >= viewBounds[3]) return false;

        var fitRatio = (options.fitRatio > 0) ? options.fitRatio : DEFAULT_FIT_RATIO;

        var targetWidth = bounds.right - bounds.left;
        var targetHeight = bounds.top - bounds.bottom;
        var visibleWidth = viewBounds[2] - viewBounds[0];
        var visibleHeight = viewBounds[1] - viewBounds[3];

        /* 収まらないときだけズームアウトする。拡大はしない（操作のたびに倍率が変わると落ち着かない）
           Only zoom out when it does not fit; never zoom in, so the magnification stays predictable */
        var targetZoom = activeView.zoom;
        if (targetWidth > visibleWidth || targetHeight > visibleHeight) {
            var zoomByWidth = (targetWidth > 0) ? targetZoom * visibleWidth / targetWidth : targetZoom;
            var zoomByHeight = (targetHeight > 0) ? targetZoom * visibleHeight / targetHeight : targetZoom;
            targetZoom = Math.min(zoomByWidth, zoomByHeight) * fitRatio;
        }

        /* 中心を合わせてからズームする（ズームは中心を保つ）/ center first, then zoom about that center */
        activeView.centerPoint = [bounds.left + targetWidth / 2, bounds.top - targetHeight / 2];
        activeView.zoom = targetZoom;
        return true;
    }

    return {
        addCheckbox: addCheckbox,
        ensureVisible: ensureVisible,
        getItemsBounds: getItemsBounds,
        getLabel: getLabel
    };

})();
