#targetengine "SOR_Engine"
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したオブジェクトを、最大／最小／キーオブジェクト／指定サイズ／基準辺／面積／アートボード／裁ち落としのいずれかの基準でリサイズし、あわせて横位置・縦位置の整列も行えます。
縦横比保持と片辺のみを切り替えでき、操作はリアルタイムにプレビューされます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectResizer.md

note記事も参照してください。
https://note.com/dtp_tranist/n/n6f35bd4000ec

### Overview

Resizes the selected objects to the largest, the smallest, the key object, a given size, a reference edge, an area, the artboard, or the bleed, and can align them horizontally and vertically at the same time.
You can switch between keeping the aspect ratio and constraining a single edge, with a real-time preview.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectResizer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SmartObjectResizer";           /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.5";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-04-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SmartObjectResizer.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartObjectResizer.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6f35bd4000ec"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/**
 * @discussion 参考 / Reference
 * キーオブジェクトの取得方法
 * 自分用メモ (@mute_racoon3631)
 * https://note.com/mute_racoon3631/n/n5dfae854988a
 *
 * ロゴなどの大きさ調整（面積を使うアイデア）
 * Gorolib Design
 * https://gorolib.blog.jp/archives/75031515.html
 */

(function () {
    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    var BLEED_OFFSET_PT = 6 * 2.83464567;    /* 裁ち落とし: 片側3mm＝幅/高さそれぞれ+6mm / bleed 3mm per side */

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* UIレイアウトの余白・間隔 / UI layout margins & spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 8;                  /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 15;                 /* 2カラムの間隔 / gap between columns */
    var LABEL_WIDTH    = 122;                /* 行ラベルの共通幅（右揃え）/ shared row-label width (right-aligned) */

    /**
     * パネルに共通のレイアウトを適用する
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * 縦方向のスペーサー（区切り用の空グループ）を追加する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {number} height - 高さ（px）
     * @returns {Group} 追加したスペーサー
     */
    function addSpacer(parentContainer, height) {
        var spacerGroup = parentContainer.add("group");
        spacerGroup.minimumSize.height = height;
        spacerGroup.maximumSize.height = height;
        return spacerGroup;
    }

    /**
     * helpTip（ツールチップ）を1コントロールまたはコントロール配列へ設定する
     * @param {Object|Object[]} targetControls - コントロール、またはその配列
     * @param {string} tipText - ツールチップの文言
     * @returns {void}
     */
    function setHelpTip(targetControls, tipText) {
        if (!tipText || !targetControls) return;
        if (typeof targetControls !== "string" && typeof targetControls.length === "number") {
            for (var i = 0; i < targetControls.length; i++) {
                if (targetControls[i]) targetControls[i].helpTip = tipText;
            }
        } else {
            targetControls.helpTip = tipText;
        }
    }

    /**
     * 共通幅で右揃えの行ラベルを追加する
     * @param {Group} parentGroup - 追加先のグループ
     * @param {string} rowLabelText - ラベル文字列（空文字なら幅だけのスペーサーになる）
     * @returns {StaticText} 追加したラベル
     */
    function addRowLabel(parentGroup, rowLabelText) {
        var rowLabel = parentGroup.add("statictext", undefined, rowLabelText);
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * ↑↓キーで数値を増減する（shift で 10 刻み、option で 0.1 刻み）
     * @param {EditText} editText - 対象の数値欄
     * @param {boolean} [allowNegative] - 負の値を許すか（省略時は false）
     * @returns {void}
     */
    function changeValueByArrowKey(editText, allowNegative) {
        if (!editText) return;
        if (typeof allowNegative === "undefined") allowNegative = false;

        editText.addEventListener("keydown", function (event) {
            /* ScriptUI の keyName は "Up" / "Down" / ScriptUI key names */
            if (!(event && (event.keyName === "Up" || event.keyName === "Down"))) return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                /* shift 押下時は 10 の倍数にスナップ / Snap to multiples of 10 with shift */
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                } else {
                    value = Math.floor((value - 1) / delta) * delta;
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            } else {
                delta = 1;
                if (event.keyName === "Up") value += delta;
                else value -= delta;
            }

            /* 丸め（option は小数第1位、それ以外は整数）/ Rounding */
            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            if (!allowNegative && value < 0) value = 0;

            event.preventDefault();
            editText.text = value;

            /* onChange を明示的に呼ぶ（矢印キーでは発火しないことがある）。中の DOM 操作が失敗してもキー操作は止めない
               Fire onChange explicitly; keep key handling alive even if its DOM work fails */
            try {
                if (typeof editText.onChange === "function") editText.onChange();
            } catch (e) { }
        });
    }

    // =========================================
    // セッション / Session
    // =========================================

    /* ダイアログ位置をセッション内で記憶（Illustrator 終了でリセット）
       Remember the dialog position within this Illustrator session (resets on quit) */
    var SESSION_POSITION_KEY = "SmartObjectResizer_dialogPos";
    if (typeof $.global[SESSION_POSITION_KEY] === "undefined") {
        $.global[SESSION_POSITION_KEY] = null;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator のロケールから表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var uiLang = detectUILanguage();

    /*
    LABELS のカテゴリ規約 / Category rules
      dialog     : ダイアログタイトル / dialog title
      panel      : パネル見出し / panel headers
      fieldLabel : 行ラベル（コロンは含めず、描画時に labelText() が付与） / row labels (colon added by labelText())
      radio      : ラジオボタン / radio buttons
      checkbox   : チェックボックス / checkboxes
      button     : ボタン（Cancel / Reset。OK は非ローカライズの "OK" 直書き）
      tooltip    : 意味が自明でないコントロールのツールチップ / tooltips for non-obvious controls
      alert      : 警告メッセージ / alerts
    記号は {colon} / {slash} / {comma} / {openParen} / {closeParen} で記述し、
    applyUISymbols() が言語に応じた全角/半角へ展開する。
    */
    var LABELS = {
        dialog: {
            title: { ja: "オブジェクトのリサイズ", en: "SmartObjectResizer" }
        },
        panel: {
            base: { ja: "リサイズ基準", en: "Resize base" },
            /* 整列（横）: 左/中央/右 ＋ 縦方向の分配 / horizontal alignment */
            hAlign: { ja: "整列{openParen}横{closeParen}", en: "Align {openParen}H{closeParen}" },
            /* 整列（縦）: 上/中央/下 ＋ 横方向の分配 / vertical alignment */
            vAlign: { ja: "整列{openParen}縦{closeParen}", en: "Align {openParen}V{closeParen}" }
        },
        fieldLabel: {
            max: { ja: "最大", en: "Max" },
            min: { ja: "最小", en: "Min" },
            key: { ja: "キーオブジェクト", en: "Key object" },
            fixed: { ja: "指定サイズ", en: "Fixed Size" },
            base: { ja: "基準辺", en: "Ref. side" },
            area: { ja: "面積", en: "Area" },
            artboard: { ja: "アートボード", en: "Artboard" },
            bleed: { ja: "裁ち落とし", en: "Bleed" }
        },
        radio: {
            keepAspect: { ja: "縦横比保持", en: "Keep aspect" },
            oneSideOnly: { ja: "片辺のみ", en: "One side only" },
            width: { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            longSide: { ja: "長辺", en: "Long side" },
            shortSide: { ja: "短辺", en: "Short side" },
            areaMax: { ja: "最大", en: "Max" },
            areaMin: { ja: "最小", en: "Min" }
        },
        checkbox: {
            textOutlineBounds: { ja: "テキストをアウトライン境界で計測", en: "Measure text by outline bounds" },
            previewBounds: { ja: "プレビュー境界で計測", en: "Measure by preview bounds" },
            alignLeft: { ja: "左", en: "Left" },
            alignCenter: { ja: "中央", en: "Center" },
            alignRight: { ja: "右", en: "Right" },
            alignEven: { ja: "均等", en: "Distribute evenly" },
            alignZero: { ja: "0間隔", en: "Zero gap" },
            alignTop: { ja: "上", en: "Top" },
            alignMiddle: { ja: "中央", en: "Middle" },
            alignBottom: { ja: "下", en: "Bottom" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            reset: { ja: "リセット", en: "Reset" }
        },
        tooltip: {
            keepAspect: { ja: "縦横比を保ったまま拡大／縮小します。", en: "Scale while keeping the aspect ratio." },
            oneSideOnly: {
                ja: "幅または高さの片辺だけを変更します（縦横比は保持しません）。",
                en: "Change only one side (width or height); aspect ratio is not preserved."
            },
            key: {
                ja: "キーオブジェクト（整列の基準に指定したオブジェクト）の幅／高さに、各オブジェクトをそろえます。",
                en: "Match every object to the width / height of the key object (the one set as the align target)."
            },
            keyNone: {
                ja: "キーオブジェクトが設定されていません。選択したうえで、基準にしたいオブジェクトをもう一度クリックしてください。",
                en: "No key object is set. With the objects selected, click the one you want as the key again."
            },
            base: {
                ja: "基準となる辺（長辺／短辺）の長さに合わせて、各オブジェクトをリサイズします。",
                en: "Resize each object to match the chosen reference side (long / short)."
            },
            area: {
                ja: "選択オブジェクトの面積を、最大／最小のものにそろえます。",
                en: "Match object areas to the largest / smallest in the selection."
            },
            artboard: {
                ja: "選択全体をアートボードの幅／高さに合わせ、中央に配置します。",
                en: "Fit the whole selection to the artboard width / height and center it."
            },
            bleed: {
                ja: "アートボード＋裁ち落とし（片側3mm）の幅／高さに合わせ、中央に配置します。",
                en: "Fit to the artboard plus bleed (3mm per side) and center it."
            },
            alignEven: {
                ja: "オブジェクトの間隔が均等になるように分配します（3つ以上で有効）。",
                en: "Distribute objects with equal gaps (needs 3+ objects)."
            },
            alignZero: {
                ja: "オブジェクトを間隔0で隙間なく並べます（2つ以上で有効）。",
                en: "Place objects with zero gap, no spacing (needs 2+ objects)."
            },
            textOutline: {
                ja: "テキストをアウトライン化した実際の字形の境界で計測します。",
                en: "Measure text by the actual outlined glyph bounds."
            },
            preview: {
                ja: "線幅や効果を含むプレビュー境界で計測します（オフは幾何境界）。",
                en: "Measure by preview bounds incl. strokes / effects (off = geometric bounds)."
            },
            reset: {
                ja: "サイズ・位置・整列をすべて元の状態に戻します。",
                en: "Revert size, position, and alignment to the original state."
            }
        },
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            selectObject: { ja: "オブジェクトを選択してください。", en: "Please select an object." }
        }
    };

    /**
     * "panel.base" のようなドット区切りのパスで LABELS から文字列を取得する
     * @param {string} labelPath - ドット区切りのキー
     * @returns {string} 表示言語の文字列（記号を展開済み。見つからなければ labelPath）
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        var localizedText = labelNode[uiLang] || labelNode.en;
        if (!localizedText) return labelPath;
        return applyUISymbols(localizedText);
    }

    /**
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {string} labelPath - ラベルのパス
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelPath) {
        return getLabel(labelPath) + uiSymbol("colon");
    }

    /**
     * {colon} などのプレースホルダを言語別の記号へ展開する
     * @param {string} sourceText - 展開前の文字列
     * @returns {string} 展開後の文字列
     */
    function applyUISymbols(sourceText) {
        return sourceText
            .replace(/\{colon\}/g,      uiSymbol("colon"))
            .replace(/\{slash\}/g,      uiSymbol("slash"))
            .replace(/\{comma\}/g,      uiSymbol("comma"))
            .replace(/\{openParen\}/g,  uiSymbol("openParen"))
            .replace(/\{closeParen\}/g, uiSymbol("closeParen"));
    }

    /**
     * 言語別の記号を返す（日本語は全角、英語は半角）
     * @param {string} symbolName - "slash" / "colon" / "comma" / "openParen" / "closeParen"
     * @returns {string} 記号。未知の名前なら空文字
     */
    function uiSymbol(symbolName) {
        if (uiLang === "ja") {
            switch (symbolName) {
                case "slash":      return "／";
                case "colon":      return "：";
                case "comma":      return "、";
                case "openParen":  return "（";
                case "closeParen": return "）";
            }
        }
        switch (symbolName) {
            case "slash":      return "/";
            case "colon":      return ":";
            case "comma":      return ", ";
            case "openParen":  return "(";
            case "closeParen": return ")";
        }
        return "";
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /**
     * ドキュメントの定規単位の表示ラベルを返す（Q／H などは pt として扱う）
     * @param {Document} doc - 対象ドキュメント
     * @returns {string} "mm" / "cm" / "inch" / "px" / "pica" / "pt"
     */
    function getRulerUnitLabel(doc) {
        switch (doc.rulerUnits) {
            case RulerUnits.Millimeters: return "mm";
            case RulerUnits.Centimeters: return "cm";
            case RulerUnits.Inches:      return "inch";
            case RulerUnits.Pixels:      return "px";
            case RulerUnits.Picas:       return "pica";
            default:                     return "pt";
        }
    }

    /**
     * 定規単位1つ分をポイントに換算する
     * @param {string} unitLabel - getRulerUnitLabel() の戻り値
     * @returns {number} 1単位あたりの pt
     */
    function getPointsPerUnit(unitLabel) {
        switch (unitLabel) {
            case "mm":   return 2.83464567;
            case "cm":   return 28.3464567;
            case "inch": return 72;
            case "pica": return 12;
            default:     return 1; /* pt / px */
        }
    }

    // =========================================
    // 境界の計測 / Bounds measuring
    // =========================================

    /**
     * オブジェクトの境界を { width, height, left, top } で返す
     * @param {PageItem} pageItem - 対象
     * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
     * @returns {{width: number, height: number, left: number, top: number}} 境界
     */
    function getPageItemBoundsObject(pageItem, useVisibleBounds) {
        var boundsArray = useVisibleBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
        return {
            width: boundsArray[2] - boundsArray[0],
            height: boundsArray[1] - boundsArray[3],
            left: boundsArray[0],
            top: boundsArray[1]
        };
    }

    /**
     * 複数のオブジェクトを包む境界を { width, height, left, top } で返す
     * @param {PageItem[]} pageItems - 対象
     * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
     * @returns {{width: number, height: number, left: number, top: number}} 境界（空なら 0）
     */
    function getBoundsFromItems(pageItems, useVisibleBounds) {
        if (!pageItems || pageItems.length === 0) {
            return { width: 0, height: 0, left: 0, top: 0 };
        }

        var left = null;
        var top = null;
        var right = null;
        var bottom = null;

        for (var i = 0; i < pageItems.length; i++) {
            var boundsArray = useVisibleBounds ? pageItems[i].visibleBounds : pageItems[i].geometricBounds;
            if (left === null || boundsArray[0] < left) left = boundsArray[0];
            if (top === null || boundsArray[1] > top) top = boundsArray[1];
            if (right === null || boundsArray[2] > right) right = boundsArray[2];
            if (bottom === null || boundsArray[3] < bottom) bottom = boundsArray[3];
        }

        return {
            width: right - left,
            height: top - bottom,
            left: left,
            top: top
        };
    }

    /**
     * アウトライン境界で計測すべきテキストを含むか（テキスト、またはテキストを含むグループ）
     * @param {PageItem} pageItem - 対象
     * @returns {boolean} 含むなら true
     */
    function containsTextForOutlineBounds(pageItem) {
        if (!pageItem) return false;
        if (pageItem.typename === "TextFrame") return true;
        if (pageItem.typename === "GroupItem") {
            return groupHasTextFrames(pageItem);
        }
        return false;
    }

    /**
     * グループ（入れ子を含む）にテキストフレームがあるか
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function groupHasTextFrames(groupItem) {
        if (!groupItem || !groupItem.pageItems) return false;
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            var childItem = groupItem.pageItems[i];
            if (childItem.typename === "TextFrame") return true;
            if (childItem.typename === "GroupItem" && groupHasTextFrames(childItem)) return true;
        }
        return false;
    }

    /**
     * グループ（入れ子を含む）のテキストフレームを集める
     * @param {GroupItem} groupItem - 対象のグループ
     * @param {TextFrame[]} collectedFrames - 集めた結果を追加する配列
     * @returns {void}
     */
    function collectTextFramesInGroup(groupItem, collectedFrames) {
        if (!groupItem || !groupItem.pageItems) return;
        for (var i = 0; i < groupItem.pageItems.length; i++) {
            var childItem = groupItem.pageItems[i];
            if (childItem.typename === "TextFrame") {
                collectedFrames.push(childItem);
            } else if (childItem.typename === "GroupItem") {
                collectTextFramesInGroup(childItem, collectedFrames);
            }
        }
    }

    /**
     * 計測用に複製したグループの中のテキストをアウトライン化する（グループ自体は残る）
     * @param {GroupItem} groupItem - 複製したグループ
     * @returns {void}
     */
    function outlineTextFramesInGroupDuplicate(groupItem) {
        var textFrames = [];
        collectTextFramesInGroup(groupItem, textFrames);
        for (var i = textFrames.length - 1; i >= 0; i--) {
            if (textFrames[i] && textFrames[i].isValid) {
                textFrames[i].createOutline();
            }
        }
    }

    /**
     * 計測用の一時オブジェクトを削除する（ロック・非表示を解いてから）
     * @param {PageItem[]} pageItems - 削除するオブジェクト
     * @returns {void}
     */
    function removeItemsSafe(pageItems) {
        if (!pageItems || pageItems.length === 0) return;
        for (var i = pageItems.length - 1; i >= 0; i--) {
            /* 無効になった参照もあるので1操作ずつ守る / Guard each step; some references may be stale */
            try {
                pageItems[i].locked = false;
            } catch (e) { }
            try {
                pageItems[i].hidden = false;
            } catch (e) { }
            try {
                pageItems[i].remove();
            } catch (e) { }
        }
    }

    /**
     * キャッシュキー用に座標を丸める
     * @param {number} coordinate - 座標
     * @returns {number} 小数第3位までに丸めた値
     */
    function roundCacheCoord(coordinate) {
        return Math.round(coordinate * 1000) / 1000;
    }

    /**
     * ドキュメントの選択を配列にコピーする（doc.selection はライブ参照になりうるため）
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択のコピー（選択なしなら空配列）
     */
    function copySelectionItems(doc) {
        var copiedItems = [];
        var currentSelection = doc.selection;
        if (!currentSelection) return copiedItems;
        for (var i = 0; i < currentSelection.length; i++) copiedItems.push(currentSelection[i]);
        return copiedItems;
    }

    /**
     * 各オブジェクトのサイズと位置を控える
     * @param {PageItem[]} pageItems - 対象
     * @returns {Array<{item: PageItem, width: number, height: number, left: number, top: number}>} 控え
     */
    function captureItemStates(pageItems) {
        var itemStates = [];
        for (var i = 0; i < pageItems.length; i++) {
            var pageItem = pageItems[i];
            itemStates.push({
                item: pageItem,
                width: pageItem.width,
                height: pageItem.height,
                left: pageItem.left,
                top: pageItem.top
            });
        }
        return itemStates;
    }

    // =========================================
    // リサイズの基本操作 / Resize primitives
    // =========================================

    /**
     * 現在の長さを目標の長さにする倍率（%）を返す
     * @param {number} currentLength - 現在の長さ
     * @param {number} targetLength - 目標の長さ
     * @returns {number} 倍率（%）
     */
    function getScalePercent(currentLength, targetLength) {
        return (targetLength / currentLength) * 100;
    }

    /**
     * モードに応じて基準とする1辺の長さを返す（長辺／短辺／幅／高さ）
     * @param {{width: number, height: number}} bounds - 境界
     * @param {Object} mode - getSelectedResizeMode() の戻り値
     * @returns {number} 辺の長さ
     */
    function measureSide(bounds, mode) {
        if (mode.isLong) return Math.max(bounds.width, bounds.height);
        if (mode.isShort) return Math.min(bounds.width, bounds.height);
        return mode.isWidth ? bounds.width : bounds.height;
    }

    /**
     * 片辺のみをスケールする（基準点は左上、線幅は変えない）
     * 引数を省略した 2 引数版は基準点が中心になり、他モード（TOPLEFT）と挙動が食い違うため、常に明示指定する。
     * @param {PageItem} pageItem - 対象
     * @param {boolean} isWidth - 幅なら true、高さなら false
     * @param {number} scalePct - 倍率（%）
     * @returns {void}
     */
    function resizeOneSide(pageItem, isWidth, scalePct) {
        var scaleX = isWidth ? scalePct : 100;
        var scaleY = isWidth ? 100 : scalePct;
        pageItem.resize(scaleX, scaleY, true, true, true, true, 100, Transformation.TOPLEFT);
    }

    /**
     * 目標サイズへリサイズする。除数 0（線のみ・空テキスト等）は Infinity 回避のため 100%＝現状維持とする
     * 線幅倍率は呼び出し側が明示する（省略時は 100%＝据え置き）。scaleW を流用すると、
     * 線幅を変えていない片辺のみのリサイズを戻すときに線幅だけが縮んでいく。
     * ※ 幅/高さが 0 の要素は resize では拡大できないため「厳密復元」ではなく現状スキップに近い（実用上は許容）
     * @param {PageItem} pageItem - 対象
     * @param {number} targetWidth - 目標の幅
     * @param {number} targetHeight - 目標の高さ
     * @param {number} [lineScalePct] - 線幅の倍率（%）
     * @returns {void}
     */
    function resizeItemToSize(pageItem, targetWidth, targetHeight, lineScalePct) {
        var scaleW = (targetWidth === 0 || pageItem.width === 0) ? 100 : (targetWidth / pageItem.width) * 100;
        var scaleH = (targetHeight === 0 || pageItem.height === 0) ? 100 : (targetHeight / pageItem.height) * 100;
        var lineScale = (typeof lineScalePct === "number") ? lineScalePct : 100;
        pageItem.resize(scaleW, scaleH, true, true, true, true, lineScale, Transformation.TOPLEFT);
    }

    // =========================================
    // キーオブジェクトの検出 / Key object detection
    // =========================================
    /* Illustrator の DOM にキーオブジェクトを示すプロパティは無いため、整列コマンドで実測して特定する。
       キーオブジェクトが設定されていると、どの向きに整列してもそのオブジェクトだけは動かない。
       左右上下の4方向すべてで動かなかったものだけを採用する。1方向だけだと「たまたま端にいた
       オブジェクト」を拾ってしまうが、4方向すべての端を兼ねることは（同一バウンズでない限り）無い。
       候補が0個または2個以上のときは判定不能として null を返し、UI 側でこの基準をディムする。
       Illustrator exposes no key-object property, so probe it: run the align commands and see
       which item stays put in all four directions. Ambiguous results yield null. */

    /* 整列後の位置差をどこまで「動いていない」とみなすか（pt） / Move tolerance in points */
    var KEY_DETECT_TOLERANCE_PT = 0.001;

    /**
     * 選択オブジェクトからキーオブジェクトを検出する
     * @param {PageItem[]} pageItems - 判定対象のオブジェクト
     * @returns {PageItem|null} キーオブジェクト。判定できないときは null
     */
    function detectKeyObject(pageItems) {
        if (!pageItems || pageItems.length < 2) return null;
        var alignCommands = ["Horizontal Align Left", "Horizontal Align Right", "Vertical Align Top", "Vertical Align Bottom"];
        var stayedPut = [];
        for (var i = 0; i < pageItems.length; i++) stayedPut.push(true);

        for (var commandIndex = 0; commandIndex < alignCommands.length; commandIndex++) {
            var savedPositions = [];
            for (var j = 0; j < pageItems.length; j++) savedPositions.push([pageItems[j].left, pageItems[j].top]);
            app.redraw(); /* executeMenuCommand は直前の DOM 変更が反映されていないと空振りする / the command misfires without a redraw */
            app.executeMenuCommand(alignCommands[commandIndex]);
            for (var k = 0; k < pageItems.length; k++) {
                if (Math.abs(pageItems[k].left - savedPositions[k][0]) > KEY_DETECT_TOLERANCE_PT ||
                    Math.abs(pageItems[k].top - savedPositions[k][1]) > KEY_DETECT_TOLERANCE_PT) {
                    stayedPut[k] = false;
                }
                /* 整列は検出のための試行なので、その場で元の位置へ戻す / Undo the probe move right away */
                pageItems[k].left = savedPositions[k][0];
                pageItems[k].top = savedPositions[k][1];
            }
        }
        app.redraw();

        var keyCandidate = null;
        for (var m = 0; m < pageItems.length; m++) {
            if (!stayedPut[m]) continue;
            if (keyCandidate !== null) return null; /* 複数残った＝判定不能 / more than one left: ambiguous */
            keyCandidate = pageItems[m];
        }
        return keyCandidate;
    }

    // =========================================
    // 整列の座標 / Alignment edges
    // =========================================

    /* 境界から各辺・中心の座標を取り出す関数 / Accessors for each edge and center of a bounds object */
    var EDGE_OF = {
        left: function (bounds) { return bounds.left; },
        right: function (bounds) { return bounds.left + bounds.width; },
        centerX: function (bounds) { return bounds.left + bounds.width / 2; },
        top: function (bounds) { return bounds.top; },
        bottom: function (bounds) { return bounds.top - bounds.height; },
        centerY: function (bounds) { return bounds.top - bounds.height / 2; }
    };

    /* X 座標を動かす辺 / Edges that move along X */
    var HORIZONTAL_EDGES = { left: true, right: true, centerX: true };

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ダイアログを表示し、選択オブジェクトのリサイズと整列を行う
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var unitLabel = getRulerUnitLabel(doc);

        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel("alert.selectObject"));
            return;
        }
        /* doc.selection はライブ参照になりうるため、配列にコピーして固定する / Freeze the selection into an array */
        var targetItems = copySelectionItems(doc);
        var originalStates = captureItemStates(targetItems);
        var resizeBaseStates = [];

        /* 直近の変形で線幅に掛けた倍率（%）。並びは originalStates と同じ。
           片辺のみは線幅を変えない（100%）ので、復元時に一律 scaleW を掛けると線幅だけがずれていく。
           Line-width percentage applied by the last transform, per item (same order as originalStates). */
        var appliedLineScales = [];
        resetAppliedLineScales();

        var keyObject = null;
        try {
            keyObject = detectKeyObject(targetItems);
        } catch (e) {
            /* 検出に失敗してもスクリプト自体は続行する（この基準がディムされるだけ）/ Keep going; the key row just dims */
            $.writeln("SmartObjectResizer: キーオブジェクトの検出に失敗 / key object detection failed — " + e);
            keyObject = null;
        }

        /* アウトライン計測のキャッシュ / Outline measurement cache */
        var outlineBoundsCache = {};
        var outlineBoundsCacheSeq = 1;
        var outlineIdMap = [];

        /* 確定／破棄の判定は dialog.show() の戻り値に一本化する。
           ボタンの onClick だけで復元すると、ESC キーやウィンドウの閉じるボタンでは
           onClick が発火しないため、プレビュー中の変形がそのまま確定してしまう。
           Deciding commit vs. discard from the return value of show() (not from the button
           handlers) is what makes ESC / the window close box behave as a real cancel. */
        var DIALOG_RESULT_OK = 1;
        var DIALOG_RESULT_CANCEL = 2;

        /* ダイアログのコントロール（buildDialog で作る）/ Dialog controls, created in buildDialog() */
        var keepRatioRadio, oneSideOnlyRadio;
        var allRadioButtons = [];       /* 基準ラジオのみ（縦横比保持／片辺のみは含まない）/ base radios only */
        var resizeBaseRadioGroups = []; /* [最大, 最小, キーオブジェクト, 指定サイズ, 基準辺, 面積, アートボード, 裁ち落とし] */
        var baseRadios, areaRadios;
        var fixedWidthRadio, fixedHeightRadio, fixedSizeInput;
        var textOutlineBoundsCheck, previewBoundsCheck;
        var alignLeftCheck, alignCenterCheck, alignRightCheck, verticalEvenCheck, verticalZeroGapCheck;
        var alignTopCheck, alignMiddleCheck, alignBottomCheck, horizontalEvenCheck, horizontalZeroGapCheck;
        var alignAxes = [];

        var resizeDialog = buildDialog();
        var dialogResult = resizeDialog.show();
        if (dialogResult !== DIALOG_RESULT_OK) {
            /* キャンセル／ESC／ウィンドウを閉じる: プレビュー中の変形をすべて破棄 / Discard every preview transform */
            restoreOriginalState();
        }

        // -----------------------------------------
        // ダイアログ / Dialog
        // -----------------------------------------

        /**
         * ダイアログを組み立てる
         * @returns {Window} ダイアログ
         */
        function buildDialog() {
            var dialogWindow = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
            dialogWindow.alignChildren = ["left", "top"];
            /* 左右インセットと下余白はここで一括管理（各ペイン／フッターの左右マージンは 0）/ Insets live on the dialog */
            dialogWindow.margins = [20, 0, 20, 20];

            /* 前回のダイアログ位置を復元（セッション内のみ）/ Restore the last position within the session */
            var savedDialogPos = $.global[SESSION_POSITION_KEY];
            if (savedDialogPos && savedDialogPos.length === 2 && !isNaN(savedDialogPos[0]) && !isNaN(savedDialogPos[1])) {
                dialogWindow.location = savedDialogPos;
            }

            dialogWindow.onShow = onDialogShow;

            /* 閉じるときに位置を記憶（セッション内のみ）/ Remember the position on close */
            dialogWindow.onClose = function () {
                $.global[SESSION_POSITION_KEY] = [dialogWindow.location[0], dialogWindow.location[1]];
                return true;
            };

            addAspectModeGroup(dialogWindow);

            var columnsGroup = dialogWindow.add("group");
            columnsGroup.orientation = "row";
            columnsGroup.alignChildren = ["top", "top"]; /* 左右ペインを上揃えに / top-align both panes */
            columnsGroup.spacing = COLUMN_SPACING;

            /* 左ペイン（リサイズ基準と計測オプション）。余白は dialogWindow.margins と columnsGroup.spacing で管理 */
            var leftPane = columnsGroup.add("group");
            leftPane.orientation = "column";
            leftPane.alignChildren = ["left", "top"];
            addResizeBasePanel(leftPane);
            addMeasureOptions(leftPane);

            /* 右ペイン（整列）/ Right pane: alignment */
            var rightPane = columnsGroup.add("group");
            rightPane.orientation = "column";
            rightPane.alignChildren = ["fill", "top"];
            addAlignPanels(rightPane);
            bindAlignChecks();

            addButtonRow(dialogWindow);
            return dialogWindow;
        }

        /**
         * 表示時：UI 状態を初期化し、元の状態へ戻す（基準が選択済みならそのモードを1回だけ反映）
         * 基準は初期未選択なので applyResizeBySelection() は何もしない（＝開いた直後は変形しない）。
         * @returns {void}
         */
        function onDialogShow() {
            /* UI 状態の初期化（DOM に触れないので保護不要。失敗したら顕在化させる）/ UI only; failures should surface */
            resetAlignChecks();
            updateInputState();
            updateRadioGroupStates();
            updateTextOutlineOptionState();

            /* 初期プレビュー（DOM 操作）のみ保護。失敗しても表示は継続するが、黙殺せずログに残す / Guard only the DOM work and log failures */
            try {
                resizeFromOriginal();
            } catch (e) {
                $.writeln("SmartObjectResizer: 初期プレビューに失敗 / initial preview failed — " + e);
            }
        }

        /**
         * 「縦横比保持」「片辺のみ」の行を作る
         * @param {Window} dialogWindow - ダイアログ
         * @returns {void}
         */
        function addAspectModeGroup(dialogWindow) {
            var aspectModeGroup = dialogWindow.add("group");
            aspectModeGroup.orientation = "row";
            aspectModeGroup.alignChildren = ["left", "center"];
            aspectModeGroup.margins = [20, 20, 0, 0];
            /* コントロール単体の margins は ScriptUI では無視されるため、間隔は親の spacing で取る
               Per-control margins are ignored in ScriptUI; use the parent group's spacing instead. */
            aspectModeGroup.spacing = 20;
            aspectModeGroup.alignment = ["center", "top"];

            keepRatioRadio = aspectModeGroup.add("radiobutton", undefined, getLabel("radio.keepAspect"));
            oneSideOnlyRadio = aspectModeGroup.add("radiobutton", undefined, getLabel("radio.oneSideOnly"));

            keepRatioRadio.value = true;
            setHelpTip(keepRatioRadio, getLabel("tooltip.keepAspect"));
            setHelpTip(oneSideOnlyRadio, getLabel("tooltip.oneSideOnly"));

            oneSideOnlyRadio.onClick = function () { onRatioModeChanged(true); };
            keepRatioRadio.onClick = function () { onRatioModeChanged(false); };
        }

        /**
         * 「リサイズ基準」パネルを作る
         * @param {Group} leftPane - 追加先の左ペイン
         * @returns {void}
         */
        function addResizeBasePanel(leftPane) {
            var resizeBasePanel = leftPane.add("panel", undefined, getLabel("panel.base"));
            setupPanel(resizeBasePanel);

            var widthHeightLabels = [getLabel("radio.width"), getLabel("radio.height")];
            var maxRadios = addRadioRow(labelText("fieldLabel.max"), widthHeightLabels, resizeBasePanel);
            var minRadios = addRadioRow(labelText("fieldLabel.min"), widthHeightLabels, resizeBasePanel);
            var keyRadios = addRadioRow(labelText("fieldLabel.key"), widthHeightLabels, resizeBasePanel);
            var fixedRadios = addFixedSizeRows(labelText("fieldLabel.fixed"), widthHeightLabels, resizeBasePanel);
            baseRadios = addRadioRow(labelText("fieldLabel.base"), [getLabel("radio.longSide"), getLabel("radio.shortSide")], resizeBasePanel);
            areaRadios = addRadioRow(labelText("fieldLabel.area"), [getLabel("radio.areaMax"), getLabel("radio.areaMin")], resizeBasePanel);
            resizeBasePanel.add("statictext", undefined, "  ───────────────  ");
            var artboardRadios = addRadioRow(labelText("fieldLabel.artboard"), widthHeightLabels, resizeBasePanel);
            var bleedRadios = addRadioRow(labelText("fieldLabel.bleed"), widthHeightLabels, resizeBasePanel);

            /* キーオブジェクトが特定できなかったときは、この行だけディムする / Dim the key row when no key object was found */
            keyRadios[0].parent.enabled = !!keyObject;

            resizeBaseRadioGroups.push(maxRadios, minRadios, keyRadios, fixedRadios, baseRadios, areaRadios, artboardRadios, bleedRadios);

            /* 意味が自明でない基準にツールチップを設定（最大／最小／指定サイズは自明なため付けない）/ Tooltips for non-obvious bases */
            setHelpTip(keyRadios, keyObject ? getLabel("tooltip.key") : getLabel("tooltip.keyNone"));
            setHelpTip(baseRadios, getLabel("tooltip.base"));
            setHelpTip(areaRadios, getLabel("tooltip.area"));
            setHelpTip(artboardRadios, getLabel("tooltip.artboard"));
            setHelpTip(bleedRadios, getLabel("tooltip.bleed"));

            /* 初期状態ではどの基準も選択しない（開いた直後は変形しない）/ No base is selected initially, so opening performs no transform */
            for (var i = 0; i < allRadioButtons.length; i++) {
                allRadioButtons[i].onClick = onResizeBaseRadioClick;
            }
        }

        /**
         * 「ラベル＋ラジオ複数」の1行を作る
         * @param {string} rowLabelText - 行ラベル
         * @param {string[]} optionLabels - ラジオの文言
         * @param {Panel} parentPanel - 追加先
         * @returns {RadioButton[]} 作ったラジオ
         */
        function addRadioRow(rowLabelText, optionLabels, parentPanel) {
            var rowGroup = parentPanel.add("group");
            rowGroup.orientation = "row";
            rowGroup.alignChildren = ["left", "center"];
            addRowLabel(rowGroup, rowLabelText);
            var rowRadios = [];
            for (var i = 0; i < optionLabels.length; i++) {
                var radioButton = rowGroup.add("radiobutton", undefined, optionLabels[i]);
                rowRadios.push(radioButton);
                allRadioButtons.push(radioButton);
            }
            return rowRadios;
        }

        /**
         * 指定サイズの2行（「ラベル＋幅/高さラジオ」と「数値入力＋単位」）を作る
         * @param {string} rowLabelText - 行ラベル
         * @param {string[]} optionLabels - 幅・高さの文言
         * @param {Panel} parentPanel - 追加先
         * @returns {RadioButton[]} 幅・高さのラジオ
         */
        function addFixedSizeRows(rowLabelText, optionLabels, parentPanel) {
            /* 1行目：ラベルとラジオボタン / Row 1: label and radios */
            var sizeHeaderGroup = parentPanel.add("group");
            sizeHeaderGroup.orientation = "row";
            sizeHeaderGroup.alignChildren = ["left", "center"];
            addRowLabel(sizeHeaderGroup, rowLabelText);
            fixedWidthRadio = sizeHeaderGroup.add("radiobutton", undefined, optionLabels[0]);
            fixedHeightRadio = sizeHeaderGroup.add("radiobutton", undefined, optionLabels[1]);
            allRadioButtons.push(fixedWidthRadio, fixedHeightRadio);

            /* 2行目：空ラベル（入力欄の左端をそろえる）と入力欄 / Row 2: blank label and the field */
            var sizeInputGroup = parentPanel.add("group");
            sizeInputGroup.orientation = "row";
            sizeInputGroup.alignChildren = ["left", "center"];
            addRowLabel(sizeInputGroup, "");

            /* 選択オブジェクトの平均幅を初期値に。getReferenceBounds() は pt を返すが、入力欄は定規単位で扱うため必ず換算する。
               換算を忘れると mm 定規で「283」と表示され、そのまま 283mm にリサイズされる。
               Bounds are in points but this field is in ruler units — always convert. */
            var totalWidth = 0;
            for (var i = 0; i < targetItems.length; i++) {
                totalWidth += getReferenceBounds(targetItems[i], true).width;
            }
            var avgWidthPt = targetItems.length > 0 ? (totalWidth / targetItems.length) : 100;
            var avgWidth = avgWidthPt / getPointsPerUnit(unitLabel);
            fixedSizeInput = sizeInputGroup.add("edittext", undefined, avgWidth.toFixed(0));
            fixedSizeInput.characters = 5;
            changeValueByArrowKey(fixedSizeInput, false);
            sizeInputGroup.add("statictext", undefined, unitLabel);

            fixedSizeInput.onChange = function () {
                if ((fixedWidthRadio.value || fixedHeightRadio.value) && !isNaN(parseFloat(fixedSizeInput.text))) {
                    /* 片辺のみ／縦横比保持のどちらも applyResizeBySelection() に集約 / Both modes go through applyResizeBySelection() */
                    resizeFromOriginal();
                }
            };
            return [fixedWidthRadio, fixedHeightRadio];
        }

        /**
         * 計測オプション（アウトライン境界／プレビュー境界）のチェックボックスを作る
         * @param {Group} leftPane - 追加先の左ペイン
         * @returns {void}
         */
        function addMeasureOptions(leftPane) {
            var measureOptionsGroup = leftPane.add("group");
            measureOptionsGroup.orientation = "column";
            measureOptionsGroup.alignChildren = ["left", "top"];
            measureOptionsGroup.margins = [0, 5, 0, 0];

            textOutlineBoundsCheck = measureOptionsGroup.add("checkbox", undefined, getLabel("checkbox.textOutlineBounds"));
            textOutlineBoundsCheck.value = false;
            textOutlineBoundsCheck.onClick = onMeasureOptionChanged;
            setHelpTip(textOutlineBoundsCheck, getLabel("tooltip.textOutline"));

            previewBoundsCheck = measureOptionsGroup.add("checkbox", undefined, getLabel("checkbox.previewBounds"));
            previewBoundsCheck.value = true;
            previewBoundsCheck.onClick = onMeasureOptionChanged;
            setHelpTip(previewBoundsCheck, getLabel("tooltip.preview"));

            updateTextOutlineOptionState();
        }

        /**
         * 整列（横）・整列（縦）のパネルを作る
         * 横位置パネルは左/中央/右＋縦方向の分配、縦位置パネルは上/中央/下＋横方向の分配
         * @param {Group} rightPane - 追加先の右ペイン
         * @returns {void}
         */
        function addAlignPanels(rightPane) {
            var hAlignPanel = rightPane.add("panel", undefined, getLabel("panel.hAlign"));
            setupPanel(hAlignPanel, 5);
            alignLeftCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignLeft"));
            alignCenterCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignCenter"));
            alignRightCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignRight"));
            addSpacer(hAlignPanel, 5); /* 「均等」の上の余白（整列⇔分配の区切り）/ gap between align and distribute */
            verticalEvenCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignEven"));
            verticalZeroGapCheck = hAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignZero"));
            setHelpTip(verticalEvenCheck, getLabel("tooltip.alignEven"));
            setHelpTip(verticalZeroGapCheck, getLabel("tooltip.alignZero"));

            var vAlignPanel = rightPane.add("panel", undefined, getLabel("panel.vAlign"));
            setupPanel(vAlignPanel, 5);
            alignTopCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignTop"));
            alignMiddleCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignMiddle"));
            alignBottomCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignBottom"));
            addSpacer(vAlignPanel, 5); /* 「均等」の上の余白（整列⇔分配の区切り）/ gap between align and distribute */
            horizontalEvenCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignEven"));
            horizontalZeroGapCheck = vAlignPanel.add("checkbox", undefined, getLabel("checkbox.alignZero"));
            setHelpTip(horizontalEvenCheck, getLabel("tooltip.alignEven"));
            setHelpTip(horizontalZeroGapCheck, getLabel("tooltip.alignZero"));
        }

        /**
         * 整列の定義テーブルを作り、クリック時の処理と有効・無効を設定する
         * テーブルは onClick の割り当てと再適用の唯一のソース。minItems はその整列が成立する最小オブジェクト数
         * （「均等」は両端を固定して間を分けるので3個以上、「0間隔」は2個以上）。
         * minItems を有効・無効にも反映しないと、2個選択で「均等」をチェックできるのに何も起きない。
         * @returns {void}
         */
        function bindAlignChecks() {
            alignAxes = [
                /* 横位置（X座標を変更）: 左 / 中央 / 右 / 横均等 / 横0 / Horizontal (moves X) */
                [
                    { check: alignLeftCheck,         minItems: 1, apply: function () { alignToExtremeEdge("left", false); } },
                    { check: alignCenterCheck,       minItems: 1, apply: function () { alignToSelectionCenter("centerX"); } },
                    { check: alignRightCheck,        minItems: 1, apply: function () { alignToExtremeEdge("right", true); } },
                    { check: horizontalEvenCheck,    minItems: 3, apply: function () { distributeHorizontal(true); } },
                    { check: horizontalZeroGapCheck, minItems: 2, apply: function () { distributeHorizontal(false); } }
                ],
                /* 縦位置（Y座標を変更）: 上 / 中央 / 下 / 均等 / 0 / Vertical (moves Y) */
                [
                    { check: alignTopCheck,        minItems: 1, apply: function () { alignToExtremeEdge("top", true); } },
                    { check: alignMiddleCheck,     minItems: 1, apply: function () { alignToSelectionCenter("centerY"); } },
                    { check: alignBottomCheck,     minItems: 1, apply: function () { alignToExtremeEdge("bottom", false); } },
                    { check: verticalEvenCheck,    minItems: 3, apply: function () { distributeVertical(true); } },
                    { check: verticalZeroGapCheck, minItems: 2, apply: function () { distributeVertical(false); } }
                ]
            ];

            for (var i = 0; i < alignAxes.length; i++) {
                var axisEntries = alignAxes[i];
                for (var j = 0; j < axisEntries.length; j++) {
                    var siblingChecks = [];
                    for (var k = 0; k < axisEntries.length; k++) {
                        if (k !== j) siblingChecks.push(axisEntries[k].check);
                    }
                    axisEntries[j].check.onClick = makeAlignHandler(axisEntries[j].check, siblingChecks);
                    axisEntries[j].check.enabled = targetItems.length >= axisEntries[j].minItems;
                }
            }
        }

        /**
         * フッター（左=リセット / スペーサー / 右=キャンセル・OK）を作る
         * @param {Window} dialogWindow - ダイアログ
         * @returns {void}
         */
        function addButtonRow(dialogWindow) {
            var btnRowGroup = dialogWindow.add("group");
            btnRowGroup.orientation = "row";
            btnRowGroup.alignment = ["fill", "bottom"];   /* 左右下の余白は dialogWindow.margins が担当 / insets come from dialogWindow.margins */
            btnRowGroup.margins = [0, 5, 0, 0];           /* 整列パネルとの間隔（上のみ）/ gap above */

            var btnLeftGroup = btnRowGroup.add("group");
            btnLeftGroup.alignChildren = ["left", "center"];
            var btnReset = btnLeftGroup.add("button", undefined, getLabel("button.reset"));
            setHelpTip(btnReset, getLabel("tooltip.reset"));
            btnReset.onClick = resetToOriginal;

            var spacer = btnRowGroup.add("group");
            spacer.alignment = ["fill", "fill"];
            spacer.minimumSize.width = 0;

            /* Mac 規約で Cancel → OK の順 / Cancel then OK, per macOS convention */
            var btnRightGroup = btnRowGroup.add("group");
            btnRightGroup.alignChildren = ["right", "center"];

            var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
            btnCancel.onClick = function () {
                dialogWindow.close(DIALOG_RESULT_CANCEL);
            };

            var btnOK = btnRightGroup.add("button", undefined, "OK", { name: "ok" });
            btnOK.onClick = function () {
                /* 確定（一時グループを使わないので親階層の復元処理は不要）/ Commit; no temporary groups to unwind */
                dialogWindow.close(DIALOG_RESULT_OK);
            };
        }

        // -----------------------------------------
        // UI の状態 / UI state
        // -----------------------------------------

        /**
         * 整列チェックボックスをすべてオフにする
         * @returns {void}
         */
        function resetAlignChecks() {
            for (var i = 0; i < alignAxes.length; i++) {
                for (var j = 0; j < alignAxes[i].length; j++) alignAxes[i][j].check.value = false;
            }
        }

        /**
         * 「片辺のみ」のとき基準辺（長辺／短辺）と面積をディムする
         * アートボード／裁ち落としは片辺スケールにも対応するため常に有効のままにする。
         * @returns {void}
         */
        function updateRadioGroupStates() {
            var keepAspectOnly = !oneSideOnlyRadio.value;
            baseRadios[0].parent.enabled = keepAspectOnly;
            areaRadios[0].parent.enabled = keepAspectOnly;
        }

        /**
         * 指定サイズ（幅/高さ）が選択されているときだけ入力欄を有効にする
         * @returns {void}
         */
        function updateInputState() {
            fixedSizeInput.enabled = !!(fixedWidthRadio.value || fixedHeightRadio.value);
        }

        /**
         * 選択にテキストがあるときだけ「テキストをアウトライン境界で計測」を有効にしてオンにする
         * @returns {void}
         */
        function updateTextOutlineOptionState() {
            var hasText = false;
            for (var i = 0; i < targetItems.length; i++) {
                if (containsTextForOutlineBounds(targetItems[i])) {
                    hasText = true;
                    break;
                }
            }

            if (hasText) {
                textOutlineBoundsCheck.enabled = true;
                textOutlineBoundsCheck.value = true;
            } else {
                textOutlineBoundsCheck.value = false;
                textOutlineBoundsCheck.enabled = false;
            }
        }

        /**
         * 「片辺のみ」でディムされる基準（基準辺／面積）の選択を解除する
         * ディムするだけだと value が残り、getSelectedResizeMode() が enabled を見ないため
         * 「片辺のみのはずが縦横比保持でリサイズされる」状態になる。
         * @returns {void}
         */
        function clearDimmedBaseSelections() {
            var dimmedGroups = [baseRadios, areaRadios];
            for (var groupIndex = 0; groupIndex < dimmedGroups.length; groupIndex++) {
                for (var radioIndex = 0; radioIndex < dimmedGroups[groupIndex].length; radioIndex++) {
                    dimmedGroups[groupIndex][radioIndex].value = false;
                }
            }
        }

        /**
         * 選択中のラジオからリサイズモードをフラグ付きで取得する
         * @returns {Object|null} モード。基準が未選択なら null
         */
        function getSelectedResizeMode() {
            for (var groupIndex = 0; groupIndex < resizeBaseRadioGroups.length; groupIndex++) {
                var sideIndex = -1;
                if (resizeBaseRadioGroups[groupIndex][0].value) sideIndex = 0;
                else if (resizeBaseRadioGroups[groupIndex][1].value) sideIndex = 1;
                if (sideIndex < 0) continue;
                return {
                    isWidth: sideIndex === 0,
                    isMax: groupIndex === 0,
                    isMin: groupIndex === 1,
                    isKey: groupIndex === 2,
                    isFixed: groupIndex === 3,
                    isLong: groupIndex === 4 && sideIndex === 0,
                    isShort: groupIndex === 4 && sideIndex === 1,
                    isArea: groupIndex === 5,
                    isAreaMax: groupIndex === 5 && sideIndex === 0,
                    isArtboard: groupIndex === 6,
                    isBleed: groupIndex === 7
                };
            }
            return null;
        }

        // -----------------------------------------
        // イベント / Events
        // -----------------------------------------

        /**
         * 縦横比保持／片辺のみを切り替える（有効・無効の切り替えは updateRadioGroupStates() が担当）
         * @param {boolean} useOneSideOnly - 片辺のみなら true
         * @returns {void}
         */
        function onRatioModeChanged(useOneSideOnly) {
            keepRatioRadio.value = !useOneSideOnly;
            oneSideOnlyRadio.value = useOneSideOnly;

            if (useOneSideOnly) clearDimmedBaseSelections();
            reapplyCurrentSelection();
        }

        /**
         * 基準ラジオのクリック：他の基準ラジオをオフにしてから再適用する（this はクリックしたラジオ）
         * @returns {void}
         */
        function onResizeBaseRadioClick() {
            for (var i = 0; i < allRadioButtons.length; i++) {
                if (allRadioButtons[i] !== this) {
                    allRadioButtons[i].value = false;
                }
            }
            reapplyCurrentSelection();
        }

        /**
         * 計測オプション（アウトライン境界／プレビュー境界）の切り替え
         * @returns {void}
         */
        function onMeasureOptionChanged() {
            resetAlignChecks();
            resizeFromOriginal();
        }

        /**
         * 整列チェックのクリック処理を作る
         * ON なら同一軸の兄弟を OFF にし、基準状態へ戻してから、チェック中の整列を両軸まとめて再適用する。
         * restoreResizeBaseState() は left/top を両方戻すため、自分の軸だけ再適用するともう一方の軸の整列が消える。
         * @param {Checkbox} alignCheck - クリックされたチェックボックス
         * @param {Checkbox[]} siblingChecks - 同じ軸の他のチェックボックス
         * @returns {Function} onClick に設定する関数
         */
        function makeAlignHandler(alignCheck, siblingChecks) {
            return function () {
                if (alignCheck.value) {
                    for (var i = 0; i < siblingChecks.length; i++) siblingChecks[i].value = false;
                }
                restoreResizeBaseState();
                reapplyActiveAlignments();
            };
        }

        /**
         * リセット：基準ラジオ・整列チェック・基準状態をすべて解除し、UI と実状態を初期状態にそろえる
         * resizeBaseStates を消さないと、この後に整列をクリックしたとき
         * restoreResizeBaseState() がリセット前のリサイズ結果を復元してしまう。
         * @returns {void}
         */
        function resetToOriginal() {
            for (var i = 0; i < allRadioButtons.length; i++) {
                allRadioButtons[i].value = false;
            }
            resizeBaseStates = [];
            resetAlignChecks();
            updateInputState();
            restoreOriginalState();
        }

        // -----------------------------------------
        // 復元と再適用 / Restore and reapply
        // -----------------------------------------

        /**
         * 線幅倍率の記録を 100% に戻す
         * @returns {void}
         */
        function resetAppliedLineScales() {
            appliedLineScales = [];
            for (var i = 0; i < targetItems.length; i++) appliedLineScales.push(100);
        }

        /**
         * 元のサイズへ戻す（掛けた線幅倍率の逆数で線幅も戻す）
         * @returns {void}
         */
        function restoreOriginalGeometry() {
            for (var i = 0; i < originalStates.length; i++) {
                var appliedLineScale = appliedLineScales[i] || 100;
                resizeItemToSize(originalStates[i].item, originalStates[i].width, originalStates[i].height, 10000 / appliedLineScale);
            }
            resetAppliedLineScales();
        }

        /**
         * 元の位置へ戻す
         * @returns {void}
         */
        function restoreOriginalPosition() {
            for (var i = 0; i < originalStates.length; i++) {
                originalStates[i].item.left = originalStates[i].left;
                originalStates[i].item.top = originalStates[i].top;
            }
        }

        /**
         * キャッシュを捨て、サイズと位置を元に戻して再描画する
         * @returns {void}
         */
        function restoreOriginalState() {
            clearOutlineBoundsCache();
            restoreOriginalGeometry();
            restoreOriginalPosition();
            app.redraw();
        }

        /**
         * 元の状態に戻してから、選択中の基準でリサイズし直す
         * @returns {void}
         */
        function resizeFromOriginal() {
            restoreOriginalState();
            applyResizeBySelection();
        }

        /**
         * 現在の選択モードを再適用する（ラジオはクリアしない）。チェック中の整列も新しいサイズに対して再適用する
         * @returns {void}
         */
        function reapplyCurrentSelection() {
            updateInputState();
            resizeFromOriginal();
            reapplyActiveAlignments();
            updateRadioGroupStates();
        }

        /**
         * リサイズ直後の状態を整列の基準として控える
         * @returns {void}
         */
        function captureResizeBaseState() {
            resizeBaseStates = captureItemStates(targetItems);
        }

        /**
         * 控えたリサイズ直後の状態へ戻す
         * @returns {void}
         */
        function restoreResizeBaseState() {
            if (!resizeBaseStates || resizeBaseStates.length === 0) return;
            for (var i = 0; i < resizeBaseStates.length; i++) {
                var itemState = resizeBaseStates[i];
                resizeItemToSize(itemState.item, itemState.width, itemState.height);
                itemState.item.left = itemState.left;
                itemState.item.top = itemState.top;
            }
            app.redraw();
        }

        /**
         * チェック中の整列を、両軸それぞれ最大1つずつ再適用する
         * @returns {void}
         */
        function reapplyActiveAlignments() {
            var changed = false;
            for (var i = 0; i < alignAxes.length; i++) {
                var axisEntries = alignAxes[i];
                for (var j = 0; j < axisEntries.length; j++) {
                    var alignEntry = axisEntries[j];
                    if (alignEntry.check.value && targetItems.length >= alignEntry.minItems) {
                        alignEntry.apply();
                        changed = true;
                        break; /* 同一軸は1つだけ / one per axis */
                    }
                }
            }
            if (changed) app.redraw();
        }

        // -----------------------------------------
        // 境界 / Bounds
        // -----------------------------------------

        /**
         * リサイズと整列に使う境界を返す
         * 「プレビュー境界で計測」ON のときは visibleBounds（線幅・効果込み＝見た目の端）、OFF のときは geometricBounds。
         * テキストは「アウトライン境界で計測」ON なら、アウトライン化した複製で計測する。
         * @param {PageItem} pageItem - 対象
         * @param {boolean} [forceVisible] - 設定に関係なくプレビュー境界を使うなら true
         * @returns {{width: number, height: number, left: number, top: number}} 境界
         */
        function getReferenceBounds(pageItem, forceVisible) {
            var useVisibleBounds = !!(forceVisible || (previewBoundsCheck && previewBoundsCheck.value));
            if (textOutlineBoundsCheck && textOutlineBoundsCheck.value && containsTextForOutlineBounds(pageItem)) {
                return getOutlinedBoundsCached(pageItem, useVisibleBounds);
            }
            return getPageItemBoundsObject(pageItem, useVisibleBounds);
        }

        /**
         * 選択全体（クラスタ）の合成境界を返す
         * @param {PageItem[]} pageItems - 対象
         * @returns {{left: number, top: number, right: number, bottom: number, width: number, height: number}} 境界
         */
        function getCombinedReferenceBounds(pageItems) {
            var left = null, top = null, right = null, bottom = null;
            for (var i = 0; i < pageItems.length; i++) {
                var itemBounds = getReferenceBounds(pageItems[i]);
                if (left === null || itemBounds.left < left) left = itemBounds.left;
                if (top === null || itemBounds.top > top) top = itemBounds.top;
                if (right === null || (itemBounds.left + itemBounds.width) > right) right = itemBounds.left + itemBounds.width;
                if (bottom === null || (itemBounds.top - itemBounds.height) < bottom) bottom = itemBounds.top - itemBounds.height;
            }
            return { left: left, top: top, right: right, bottom: bottom, width: right - left, height: top - bottom };
        }

        /**
         * アウトライン計測のキャッシュを全消去する（幾何変化はキーでも吸収するが、モード切り替え時は明示的に消す）
         * @returns {void}
         */
        function clearOutlineBoundsCache() {
            outlineBoundsCache = {};
            outlineIdMap = [];
        }

        /**
         * アウトライン計測のキャッシュキーを返す（オブジェクトの ID ＋境界）
         * @param {PageItem} pageItem - 対象
         * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
         * @returns {string} キー
         */
        function getOutlineCacheKey(pageItem, useVisibleBounds) {
            var boundsArray = useVisibleBounds ? pageItem.visibleBounds : pageItem.geometricBounds;
            return getOutlineId(pageItem) + "_" + (useVisibleBounds ? "v_" : "g_") +
                roundCacheCoord(boundsArray[0]) + "_" +
                roundCacheCoord(boundsArray[1]) + "_" +
                roundCacheCoord(boundsArray[2]) + "_" +
                roundCacheCoord(boundsArray[3]);
        }

        /**
         * キャッシュ用にオブジェクトへ振った ID を返す（無ければ振る）
         * @param {PageItem} pageItem - 対象
         * @returns {string} ID
         */
        function getOutlineId(pageItem) {
            for (var i = 0; i < outlineIdMap.length; i++) {
                if (outlineIdMap[i].item === pageItem) {
                    return outlineIdMap[i].id;
                }
            }
            var newId = "sor_" + (outlineBoundsCacheSeq++);
            outlineIdMap.push({ item: pageItem, id: newId });
            return newId;
        }

        /**
         * アウトライン境界をキャッシュ経由で返す
         * @param {PageItem} pageItem - 対象
         * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
         * @returns {{width: number, height: number, left: number, top: number}} 境界
         */
        function getOutlinedBoundsCached(pageItem, useVisibleBounds) {
            var cacheKey = getOutlineCacheKey(pageItem, useVisibleBounds);
            if (outlineBoundsCache.hasOwnProperty(cacheKey)) {
                return outlineBoundsCache[cacheKey];
            }
            var measuredBounds = measureOutlinedBoundsByDuplicate(pageItem, useVisibleBounds);
            outlineBoundsCache[cacheKey] = measuredBounds;
            return measuredBounds;
        }

        /**
         * 複製をアウトライン化して境界を計測する（一時オブジェクトは必ず削除）
         * 計測用の一時オブジェクトは、ドキュメント全体の差分ではなく生成物を直接ためて追跡する
         * （全件差分は参照の線形探索と組み合わさって O(N^2) になる）。
         * @param {PageItem} pageItem - 対象
         * @param {boolean} useVisibleBounds - プレビュー境界を使うなら true
         * @returns {{width: number, height: number, left: number, top: number}} 境界（失敗時は元の境界）
         */
        function measureOutlinedBoundsByDuplicate(pageItem, useVisibleBounds) {
            var createdItems = [];
            try {
                var duplicateItem = pageItem.duplicate();
                createdItems.push(duplicateItem);

                /* テキスト計測用の複製では、アウトライン化の前に分割を適用する / Expand appearance before outlining */
                if (containsTextForOutlineBounds(duplicateItem)) {
                    var expandedItems = expandDuplicateAppearance(duplicateItem);
                    if (expandedItems.length > 0) {
                        /* expandStyle は元の複製を作り直すので、追跡対象ごと差し替える / expandStyle rebuilds the duplicate */
                        createdItems = expandedItems;
                        duplicateItem = expandedItems[0];
                    }
                }

                if (duplicateItem.typename === "TextFrame") {
                    /* createOutline() は元のテキストフレームを置き換える / createOutline() replaces the frame */
                    duplicateItem = duplicateItem.createOutline();
                    createdItems = [duplicateItem];
                } else if (duplicateItem.typename === "GroupItem") {
                    /* グループ自体は残り、中のテキストフレームだけが置き換わる / The group stays; only its text is replaced */
                    outlineTextFramesInGroupDuplicate(duplicateItem);
                }

                if (createdItems.length === 0) return getPageItemBoundsObject(pageItem, useVisibleBounds);
                return getBoundsFromItems(createdItems, useVisibleBounds);
            } catch (e) {
                return getPageItemBoundsObject(pageItem, useVisibleBounds);
            } finally {
                removeItemsSafe(createdItems);
            }
        }

        /**
         * 複製に「アピアランスを分割」を適用し、生成された新しいアイテムの配列を返す
         * @param {PageItem} duplicateItem - 計測用の複製
         * @returns {PageItem[]} 生成されたアイテム（失敗時は空配列。呼び出し側は元の複製をそのまま使う）
         */
        function expandDuplicateAppearance(duplicateItem) {
            var expandedItems = [];
            var previousSelection = copySelectionItems(doc);
            try {
                doc.selection = null;
                duplicateItem.selected = true;
                app.executeMenuCommand('expandStyle');
                expandedItems = copySelectionItems(doc);
            } catch (e) {
            } finally {
                restoreSelectionItems(previousSelection);
            }
            return expandedItems;
        }

        /**
         * 選択を復元する（配列を doc.selection に直接代入せず、1件ずつ selected を立てる）
         * @param {PageItem[]} pageItems - 選択し直すオブジェクト
         * @returns {void}
         */
        function restoreSelectionItems(pageItems) {
            try {
                doc.selection = null;
                for (var i = 0; i < pageItems.length; i++) {
                    if (pageItems[i] && pageItems[i].isValid) pageItems[i].selected = true;
                }
            } catch (e) { }
        }

        // -----------------------------------------
        // リサイズ / Resize
        // -----------------------------------------

        /**
         * アクティブアートボードの矩形を返す
         * @returns {number[]} [左, 上, 右, 下]
         */
        function getActiveArtboardRect() {
            return doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        }

        /**
         * 目標寸法（pt）を求める（面積モードは applyAreaResize で別処理）
         * @param {Object} mode - getSelectedResizeMode() の戻り値
         * @returns {number|null} 目標寸法。求められなければ null
         */
        function computeReferenceValue(mode) {
            if (mode.isKey) {
                if (!keyObject) return null;
                var keyBounds = getReferenceBounds(keyObject);
                return mode.isWidth ? keyBounds.width : keyBounds.height;
            }
            if (mode.isFixed) {
                var parsedSize = parseFloat(fixedSizeInput.text);
                if (isNaN(parsedSize) || parsedSize <= 0) return null;
                return parsedSize * getPointsPerUnit(unitLabel);
            }
            if (mode.isArtboard) {
                var artboardRect = getActiveArtboardRect();
                return mode.isWidth ? (artboardRect[2] - artboardRect[0]) : (artboardRect[1] - artboardRect[3]);
            }
            if (mode.isBleed) {
                var bleedBase = getActiveArtboardRect();
                var sideLength = mode.isWidth ? (bleedBase[2] - bleedBase[0]) : (bleedBase[1] - bleedBase[3]);
                return sideLength + BLEED_OFFSET_PT;
            }
            /* 最大／最小／長辺／短辺: 全アイテムから基準値を集計 / Max, min, long, short: aggregate over all items */
            var referenceValue = null;
            for (var i = 0; i < targetItems.length; i++) {
                var measuredLength = measureSide(getReferenceBounds(targetItems[i]), mode);
                if (mode.isMin && measuredLength === 0) continue;
                if (referenceValue === null) referenceValue = measuredLength;
                else if (mode.isMax && measuredLength > referenceValue) referenceValue = measuredLength;
                else if (mode.isMin && measuredLength < referenceValue) referenceValue = measuredLength;
            }
            return referenceValue;
        }

        /**
         * 選択全体をひとまとまりとしてスケールする（アートボード／裁ち落とし基準。グループ化しない）
         * クラスタ左上を原点に、各アイテムのサイズと相対位置を同じ倍率でスケールする。
         * 「片辺のみ」では基準辺の軸だけ倍率を掛け、もう一方の軸は等倍のまま残す。
         * 親階層・重ね順を一切変更しないので、一時グループ方式の復元リスクを構造的に回避する。
         * @param {Object} mode - getSelectedResizeMode() の戻り値
         * @param {number} referenceValue - 目標寸法（pt）
         * @returns {void}
         */
        function applyClusterResize(mode, referenceValue) {
            var clusterBounds = getCombinedReferenceBounds(targetItems);
            var currentLength = mode.isWidth ? clusterBounds.width : clusterBounds.height;
            if (!currentLength) return;
            var scaleFactor = referenceValue / currentLength;
            var isOneSideOnly = oneSideOnlyRadio.value;
            var factorX = (!isOneSideOnly || mode.isWidth) ? scaleFactor : 1;
            var factorY = (!isOneSideOnly || !mode.isWidth) ? scaleFactor : 1;

            /* リサイズで位置がずれる前に、各アイテムの左上と原点（クラスタ左上）を記録 / Record positions before resizing */
            var originLeft = null, originTop = null;
            var originalPositions = [];
            for (var i = 0; i < targetItems.length; i++) {
                var itemLeft = targetItems[i].left, itemTop = targetItems[i].top;
                originalPositions.push({ left: itemLeft, top: itemTop });
                if (originLeft === null || itemLeft < originLeft) originLeft = itemLeft;
                if (originTop === null || itemTop > originTop) originTop = itemTop;
            }

            /* 線幅は等倍のときだけ追従させる（非等倍では resizeOneSide と同じく 100%）/ Line widths follow uniform scaling only */
            var lineScalePct = isOneSideOnly ? 100 : scaleFactor * 100;
            for (var j = 0; j < targetItems.length; j++) {
                var pageItem = targetItems[j];
                pageItem.resize(factorX * 100, factorY * 100, true, true, true, true, lineScalePct, Transformation.TOPLEFT);
                appliedLineScales[j] = lineScalePct;
                /* 原点からの相対位置も同倍率でスケール（クラスタとして拡大縮小）/ Scale the offsets from the origin too */
                pageItem.left = originLeft + (originalPositions[j].left - originLeft) * factorX;
                pageItem.top = originTop - (originTop - originalPositions[j].top) * factorY;
            }
        }

        /**
         * 各アイテムを基準値に合わせてスケールする（縦横比保持／片辺のみ）
         * @param {Object} mode - getSelectedResizeMode() の戻り値
         * @param {number} referenceValue - 目標寸法（pt）
         * @returns {void}
         */
        function resizeItemsToReference(mode, referenceValue) {
            var keepOneSideOnly = oneSideOnlyRadio.value;
            for (var i = 0; i < targetItems.length; i++) {
                var currentLength = measureSide(getReferenceBounds(targetItems[i]), mode);
                if (currentLength === 0) continue;
                var scalePct = getScalePercent(currentLength, referenceValue);

                if (!keepOneSideOnly) {
                    targetItems[i].resize(scalePct, scalePct, true, true, true, true, scalePct, Transformation.TOPLEFT);
                    appliedLineScales[i] = scalePct;
                } else {
                    /* 指定サイズも measureSide() が幅／高さを返すので、ここは共通で扱える / Fixed size shares this path */
                    resizeOneSide(targetItems[i], mode.isWidth, scalePct);
                    appliedLineScales[i] = 100;
                }
            }
        }

        /**
         * 選択全体をアートボードの中心へ移動する（裁ち落としのオフセットは上下左右対称なので、中心はアートボードと一致する）
         * @returns {void}
         */
        function centerItemsOnArtboard() {
            var artboardRect = getActiveArtboardRect();
            var centerX = (artboardRect[0] + artboardRect[2]) / 2;
            var centerY = (artboardRect[1] + artboardRect[3]) / 2;
            var clusterBounds = getCombinedReferenceBounds(targetItems);
            var dx = centerX - (clusterBounds.left + clusterBounds.width / 2);
            var dy = centerY - (clusterBounds.top - clusterBounds.height / 2);
            for (var i = 0; i < targetItems.length; i++) {
                targetItems[i].left += dx;
                targetItems[i].top += dy;
            }
        }

        /**
         * 面積を目標の面積に合わせる
         * @param {PageItem} pageItem - 対象
         * @param {number} targetArea - 目標の面積
         * @returns {number} 適用した線幅倍率（%）。復元時の逆算に使う
         */
        function resizeToMatchArea(pageItem, targetArea) {
            var bounds = getReferenceBounds(pageItem);
            var area = bounds.width * bounds.height;
            if (area === 0) return 100;
            var scalePct = Math.sqrt(targetArea / area) * 100;
            pageItem.resize(scalePct, scalePct, true, true, true, true, scalePct, Transformation.TOPLEFT);
            return scalePct;
        }

        /**
         * 面積基準：全アイテムの面積を最大／最小の面積に合わせる
         * @param {Object} mode - getSelectedResizeMode() の戻り値
         * @returns {void}
         */
        function applyAreaResize(mode) {
            var areas = [];
            for (var i = 0; i < targetItems.length; i++) {
                var itemBounds = getReferenceBounds(targetItems[i]);
                areas.push(itemBounds.width * itemBounds.height);
            }
            var baseArea = mode.isAreaMax ? Math.max.apply(null, areas) : Math.min.apply(null, areas);
            if (!baseArea || baseArea <= 0) return;
            for (var j = 0; j < targetItems.length; j++) {
                appliedLineScales[j] = resizeToMatchArea(targetItems[j], baseArea);
            }
            captureResizeBaseState();
            app.redraw();
        }

        /**
         * 選択中の基準でリサイズする（呼び出し側は元のジオメトリへ戻してから呼ぶ）
         * @returns {void}
         */
        function applyResizeBySelection() {
            /* ここで基準状態も捨てておく。捨てないと、以降の早期 return（基準未選択・指定サイズが 0 など）で古いリサイズ結果が
               resizeBaseStates に残り、整列クリック時の restoreResizeBaseState() がそれを復元してしまう。
               Drop the stale base state first, or an early return would leave a previous resize to be restored. */
            resizeBaseStates = [];

            var mode = getSelectedResizeMode();
            if (!mode) return;
            clearOutlineBoundsCache();

            if (mode.isArea) {
                applyAreaResize(mode);
                return;
            }

            var referenceValue = computeReferenceValue(mode);
            if (referenceValue === null || referenceValue <= 0) return;

            if (mode.isArtboard || mode.isBleed) {
                /* 選択全体をアートボード（＋裁ち落とし）に合わせてスケールし、中心へ配置 / Fit the cluster and center it */
                applyClusterResize(mode, referenceValue);
                centerItemsOnArtboard();
            } else {
                /* 最大／最小／指定サイズ／長辺／短辺: 各アイテムを個別にスケール / Scale each item on its own */
                resizeItemsToReference(mode, referenceValue);
            }
            captureResizeBaseState();
            app.redraw();
        }

        // -----------------------------------------
        // 整列と分配 / Align and distribute
        // -----------------------------------------

        /**
         * 各アイテムの指定の辺（または中心）を targetValue へ動かす（位置は差分で動かす）
         * item.top / item.left は visibleBounds 基準なので、境界の値を直接代入すると
         * 「プレビュー境界で計測」OFF（geometricBounds）のとき線幅の半分ずれる。
         * @param {number} targetValue - 揃える座標
         * @param {string} edgeName - EDGE_OF のキー
         * @returns {void}
         */
        function shiftItemsToEdge(targetValue, edgeName) {
            var isHorizontal = HORIZONTAL_EDGES[edgeName] === true;
            for (var i = 0; i < targetItems.length; i++) {
                var delta = targetValue - EDGE_OF[edgeName](getReferenceBounds(targetItems[i]));
                if (isHorizontal) {
                    targetItems[i].left += delta;
                } else {
                    targetItems[i].top += delta;
                }
            }
        }

        /**
         * いちばん外側の辺に揃える（左・下は最小、右・上は最大）
         * @param {string} edgeName - "left" / "right" / "top" / "bottom"
         * @param {boolean} useMaximum - 最大値に揃えるなら true
         * @returns {void}
         */
        function alignToExtremeEdge(edgeName, useMaximum) {
            var extremeValue = null;
            for (var i = 0; i < targetItems.length; i++) {
                var edgeValue = EDGE_OF[edgeName](getReferenceBounds(targetItems[i]));
                if (extremeValue === null || (useMaximum ? edgeValue > extremeValue : edgeValue < extremeValue)) extremeValue = edgeValue;
            }
            if (extremeValue === null) return;
            shiftItemsToEdge(extremeValue, edgeName);
        }

        /**
         * 選択全体の境界の中心に揃える（各中心の平均ではなく、Illustrator 標準の整列と同じ基準）
         * @param {string} centerName - "centerX" / "centerY"
         * @returns {void}
         */
        function alignToSelectionCenter(centerName) {
            shiftItemsToEdge(EDGE_OF[centerName](getCombinedReferenceBounds(targetItems)), centerName);
        }

        /**
         * 縦方向に分配する（上から順に並べる。useGap=false で 0 間隔）
         * @param {boolean} useGap - 間隔を均等にするなら true、0 間隔なら false
         * @returns {void}
         */
        function distributeVertical(useGap) {
            var sortedItems = targetItems.slice(0).sort(function (itemA, itemB) {
                return getReferenceBounds(itemB).top - getReferenceBounds(itemA).top;
            });
            var topMost = getReferenceBounds(sortedItems[0]).top;
            var gap = 0;
            if (useGap) {
                var lastBounds = getReferenceBounds(sortedItems[sortedItems.length - 1]);
                var bottomMost = lastBounds.top - lastBounds.height;
                var totalHeight = 0;
                for (var i = 0; i < sortedItems.length; i++) totalHeight += getReferenceBounds(sortedItems[i]).height;
                gap = (topMost - bottomMost - totalHeight) / (sortedItems.length - 1);
            }
            var currentY = topMost;
            for (var j = 0; j < sortedItems.length; j++) {
                var bounds = getReferenceBounds(sortedItems[j]);
                sortedItems[j].top += currentY - bounds.top;
                currentY -= (bounds.height + gap);
            }
        }

        /**
         * 横方向に分配する（左から順に並べる。useGap=false で 0 間隔）
         * @param {boolean} useGap - 間隔を均等にするなら true、0 間隔なら false
         * @returns {void}
         */
        function distributeHorizontal(useGap) {
            var sortedItems = targetItems.slice(0).sort(function (itemA, itemB) {
                return getReferenceBounds(itemA).left - getReferenceBounds(itemB).left;
            });
            var leftMost = getReferenceBounds(sortedItems[0]).left;
            var gap = 0;
            if (useGap) {
                var lastBounds = getReferenceBounds(sortedItems[sortedItems.length - 1]);
                var rightMost = lastBounds.left + lastBounds.width;
                var totalWidth = 0;
                for (var i = 0; i < sortedItems.length; i++) totalWidth += getReferenceBounds(sortedItems[i]).width;
                gap = (rightMost - leftMost - totalWidth) / (sortedItems.length - 1);
            }
            var currentX = leftMost;
            for (var j = 0; j < sortedItems.length; j++) {
                var bounds = getReferenceBounds(sortedItems[j]);
                sortedItems[j].left += currentX - bounds.left;
                currentX += (bounds.width + gap);
            }
        }
    }

    main();
})();
