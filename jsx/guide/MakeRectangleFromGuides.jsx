#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ガイドの交点で区切られた区画に、塗りつぶした長方形を一括生成します。対象にするガイドと作成先はダイアログで指定できます。
詳細はREADMEを参照してください。

### Overview

Fills every area bounded by guide intersections with a generated rectangle. A dialog selects which guides to use and where the rectangles go.
See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MakeRectangleFromGuides";      /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-13";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-17";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeRectangleFromGuides.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MakeRectangleFromGuides.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n4907511336ad"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 作成先に専用レイヤーを選んだときのレイヤー名 / Layer name used when the dedicated destination is selected */
    var OUTPUT_LAYER_NAME = "Generated Rectangles";

    /* 生成した長方形の不透明度（%）/ Opacity of the generated rectangles (%) */
    var RECT_OPACITY = 50;

    /* 2つの座標を同一とみなす許容誤差（pt）/ Tolerance for treating two coordinates as identical (pt) */
    var COORD_TOLERANCE = 0.01;

    /* ルーラーガイド判定：アートボードをこの長さ以上はみ出すガイドを全面ガイドとみなす（pt）/ A guide overhanging the artboard by at least this much counts as a ruler guide (pt) */
    var RULER_GUIDE_MARGIN_PT = 100;

    /* プレビューで描く長方形の上限（多すぎるとダイアログ操作が重くなる）/ Cap on previewed rectangles, beyond which the dialog would crawl */
    var PREVIEW_MAX_RECTS = 500;

    // =========================================
    // レイアウト / Layout
    // =========================================

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;               /* 2カラムの間隔 / gap between the two columns */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */
    var BUTTON_BAR_SPACING = 10;               /* ボタンバー内の要素間隔 / spacing inside the button bar */
    var FIELD_ROW_SPACING  = 6;                /* ラジオと入力欄の間隔 / gap between a radio and its input */
    var LAYER_NAME_CHARS   = 15;               /* レイヤー名入力欄の最小幅（文字数）/ minimum width of the layer name field (characters) */
    var INDENT_MARGINS     = [20, 0, 0, 0];    /* ラジオの下に続く行の字下げ / indent for a row that follows its radio */
    var MESSAGE_MARGINS    = [0, 0, 0, 10];    /* パネル最上部のメッセージの下余白 / space under the message at the top of a panel */
    var OFFSET_CHARS       = 5;                /* オフセット入力欄の幅（文字数）/ width of the offset field (characters) */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の表示言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + ""; /* 文字列化して扱う / Ensure a string */
        /* "ja" で始まるロケール（ja, ja_JP など）は日本語扱い / Treat "ja*" locales as Japanese */
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title:   { ja: "ガイドから長方形を作成", en: "Create Rectangles from Guides" },
            preview: { ja: "プレビュー", en: "Preview" }
        },
        guideSource: {
            panelTitle: { ja: "対象となるガイド", en: "Source Guides" }
        },
        guideType: {
            panelTitle: { ja: "ガイドの種類", en: "Guide Type" },
            all:        { ja: "すべてのガイド", en: "All Guides" },
            rulerOnly:  { ja: "ルーラーガイドのみ", en: "Ruler Guides Only" }
        },
        targetLayer: {
            panelTitle: { ja: "レイヤー", en: "Layers" },
            allLayers:  { ja: "すべてのレイヤー", en: "All Layers" },
            activeOnly: { ja: "現在のレイヤーのみ", en: "Current Layer Only" }
        },
        guideOption: {
            panelTitle:         { ja: "絞り込み", en: "Filters" },
            activeArtboardOnly: { ja: "現在のアートボードのみ", en: "Current Artboard Only" },
            includeLocked:      { ja: "ロックされたレイヤーを含む", en: "Include Locked Layers" }
        },
        rectangle: {
            panelTitle: { ja: "作成する長方形", en: "Rectangles to Create" }
        },
        destination: {
            panelTitle:  { ja: "作成先", en: "Destination" },
            activeLayer: { ja: "現在のレイヤー", en: "Current Layer" },
            outputLayer: { ja: "指定レイヤー", en: "Specific Layer" }
        },
        rectangleOption: {
            panelTitle:     { ja: "後処理", en: "After Creation" },
            offset:         { ja: "オフセット：", en: "Offset:" },
            mergeAdjacent:  { ja: "長方形を1つに結合", en: "Merge Into a Single Path" },
            convertToShape: { ja: "シェイプに変換", en: "Convert to Shape" }
        },
        summary: {
            guides: {
                ja: "縦#vertical#本・横#horizontal#本",
                en: "#vertical# vertical / #horizontal# horizontal"
            },
            count: {
                ja: "#count#個",
                en: "#count#"
            },
            merged: {
                ja: "#count#個 → 結合して1つ",
                en: "#count# → merged into 1"
            }
        },
        tooltip: {
            guideType: {
                ja: "「ルーラーガイドのみ」は、アートボードをまたぐ長さのガイドだけを拾います",
                en: "\"Ruler Guides Only\" keeps just the guides that run past the artboard edges"
            },
            targetLayer: {
                ja: "ガイドを探す範囲です。長方形の作成先とは別の設定です",
                en: "Where to look for guides. This is separate from where the rectangles are created"
            },
            activeArtboardOnly: {
                ja: "現在のアートボードの範囲内にあるガイドだけを使います",
                en: "Uses only the guides that fall inside the current artboard"
            },
            includeLocked: {
                ja: "ロックされたレイヤーを一時的に解除してガイドを集め、集め終わったらロックを戻します",
                en: "Temporarily unlocks locked layers to collect their guides, then restores the lock"
            },
            destination: {
                ja: "長方形を作成するレイヤーです",
                en: "Layer that receives the rectangles"
            },
            outputLayerName: {
                ja: "作成先のレイヤー名。同名のレイヤーがすでにあれば、そのレイヤーを再利用します",
                en: "Name of the destination layer. An existing layer with the same name is reused"
            },
            offset: {
                ja: "各長方形を四辺とも広げます（パスのオフセットと同じで、マイナス値なら縮みます）。単位はルーラーに従います",
                en: "Grows every rectangle on all four sides, like Offset Path; a negative value shrinks it. The unit follows the ruler"
            },
            mergeAdjacent: {
                ja: "作成した長方形をすべて合体して1つのパスにします",
                en: "Unites every generated rectangle into a single path"
            },
            convertToShape: {
                ja: "作成した長方形をライブシェイプに変換します。結合する場合は結合後に変換します",
                en: "Converts the result into live shapes. When merging, the conversion runs after the merge"
            },
            guideCount: {
                ja: "下の条件で見つかったガイドの本数です。同じ位置に重なったガイドは1本として数えます",
                en: "Guides found under the settings below. Guides at the same position count as one"
            },
            preview: {
                ja: "作成される長方形をカンバス上に仮表示します。結合とシェイプ変換はOKを押したときに適用されます",
                en: "Draws the rectangles on the canvas. Merging and shape conversion are applied when you press OK"
            },
            summaryCount: {
                ja: "OKを押したときに作成される長方形の数です",
                en: "How many rectangles pressing OK will create"
            }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:        { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            notEnoughGuides:   { ja: "条件に合うガイドが足りません。縦・横それぞれ2本以上必要です。", en: "Not enough matching guides. At least two vertical and two horizontal guides are required." },
            lockedActiveLayer: { ja: "現在のレイヤーがロックまたは非表示のため作成できません。ロックを解除するか、作成先を「指定レイヤー」にしてください。", en: "Cannot draw because the current layer is locked or hidden. Unlock it, or set the destination to the specific layer." }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('button','ok')）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} 該当するラベル（見つからない場合は空文字）
     */
    function getLabel() {
        var labelNode = LABELS;
        for (var i = 0; i < arguments.length; i++) {
            if (labelNode == null) break;
            labelNode = labelNode[arguments[i]];
        }
        return (labelNode && labelNode[uiLang] != null) ? labelNode[uiLang] : "";
    }

    /**
     * ラベル内の #キー# を値に差し替える
     * @param {string} template - 差し替え前の文字列
     * @param {object} values - キーと値の対応
     * @returns {string} 差し替え後の文字列
     */
    function formatMessage(template, values) {
        var formatted = template;
        for (var key in values) {
            formatted = formatted.replace(new RegExp("#" + key + "#", "g"), values[key]);
        }
        return formatted;
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; array index equals the rulerType code (0=in, 1=mm, 2=pt …) */
    var UNITS = [
        { label: "in",    factor: 72.0 },                /* 0 */
        { label: "mm",    factor: 72.0 / 25.4 },         /* 1 */
        { label: "pt",    factor: 1.0 },                 /* 2 */
        { label: "pica",  factor: 12.0 },                /* 3 */
        { label: "cm",    factor: 72.0 / 2.54 },         /* 4 */
        { label: "Q/H",   factor: 72.0 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    factor: 1.0 },                 /* 6 */
        { label: "ft/in", factor: 72.0 * 12.0 },         /* 7 */
        { label: "m",     factor: 72.0 / 25.4 * 1000.0 },/* 8 */
        { label: "yd",    factor: 72.0 * 36.0 },         /* 9 */
        { label: "ft",    factor: 72.0 * 12.0 }          /* 10 */
    ];

    /* pt の添字（単位が特定できないときのフォールバック）/ Index of pt, used as the fallback unit */
    var POINT_UNIT_INDEX = 2;

    /**
     * ルーラー環境設定の単位を取得する
     * @returns {{label: string, factor: number}} 単位ラベルと pt 換算係数
     */
    function getRulerUnit() {
        var rulerTypeCode = app.preferences.getIntegerPreference("rulerType");
        return (rulerTypeCode >= 0 && rulerTypeCode < UNITS.length) ? UNITS[rulerTypeCode] : UNITS[POINT_UNIT_INDEX];
    }

    /**
     * 入力値を pt に変換する
     * @param {string} inputText - 入力欄の値
     * @param {number} factor - 単位の pt 換算係数
     * @returns {number} pt に変換した値（数値として読めなければ0）
     */
    function convertToPt(inputText, factor) {
        var numericValue = parseFloat(inputText);
        return isNaN(numericValue) ? 0 : numericValue * factor;
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * パネルに共通レイアウトを適用する
     * @param {Panel} targetPanel - 対象パネル
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
     * グループを横並びの行として設定する
     * @param {Group} targetGroup - 対象グループ
     * @param {string} [horizontalAlign] - 横方向の揃え（省略時は "left"）
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(targetGroup, horizontalAlign, spacing) {
        targetGroup.orientation = "row";
        /* 揃えは横と天地を対で指定し、親の fill 継承を打ち消す / Pair both axes to cancel the parent's fill */
        targetGroup.alignment = [horizontalAlign || "left", "center"];
        targetGroup.alignChildren = ["left", "center"];
        targetGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ラベル付きパネルを生成する（共通レイアウト適用）
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @returns {Panel} 生成したパネル
     */
    function addPanel(parentContainer, panelTitle) {
        var createdPanel = parentContainer.add("panel");
        createdPanel.text = panelTitle;
        setupPanel(createdPanel);
        return createdPanel;
    }

    /**
     * 左寄せの縦並びグループを生成する（ラジオ列・チェックボックス列など）
     * @param {Window|Group|Panel} parentContainer - 追加先
     * @returns {Group} 生成したグループ
     */
    function addLeftAlignedColumn(parentContainer) {
        var createdGroup = parentContainer.add("group");
        createdGroup.orientation = "column";
        createdGroup.alignment = ["left", "top"];
        createdGroup.alignChildren = ["left", "center"];
        return createdGroup;
    }

    /**
     * 2択のラジオボタン列を生成する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} firstLabel - 1つ目のラベル
     * @param {string} secondLabel - 2つ目のラベル
     * @param {number} selectedIndex - 初期選択（0=1つ目、1=2つ目）
     * @param {string} [helpTipText] - ツールチップ（親コンテナと両ボタンに付与）
     * @returns {RadioButton[]} [1つ目, 2つ目] のラジオボタン
     */
    function addRadioPair(parentContainer, firstLabel, secondLabel, selectedIndex, helpTipText) {
        var radioColumn = addLeftAlignedColumn(parentContainer);
        var firstRadio = radioColumn.add("radiobutton", undefined, firstLabel);
        var secondRadio = radioColumn.add("radiobutton", undefined, secondLabel);
        ((selectedIndex === 1) ? secondRadio : firstRadio).value = true;
        if (helpTipText) {
            /* パネルの余白でもボタン上でも出るように両方へ付ける / Attach to both so the tip shows over the panel and the buttons */
            parentContainer.helpTip = helpTipText;
            firstRadio.helpTip = helpTipText;
            secondRadio.helpTip = helpTipText;
        }
        return [firstRadio, secondRadio];
    }

    /**
     * パネル内にチェックボックスを追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} labelText - ラベル
     * @param {boolean} defaultValue - 初期状態
     * @param {string} [helpTipText] - ツールチップ
     * @returns {Checkbox} 生成したチェックボックス
     */
    function addPanelCheckbox(parentPanel, labelText, defaultValue, helpTipText) {
        var checkbox = parentPanel.add("checkbox", undefined, labelText);
        /* パネルの fill を打ち消して幅いっぱいに広げない / Cancel the panel's fill so the control keeps its natural width */
        checkbox.alignment = ["left", "center"];
        checkbox.value = defaultValue;
        if (helpTipText) {
            checkbox.helpTip = helpTipText;
        }
        return checkbox;
    }

    /**
     * パネル最上部に置く1行メッセージを生成する（左右中央、下に余白）
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {string} [helpTipText] - ツールチップ
     * @returns {StaticText} 生成したメッセージ欄
     */
    function addPanelMessage(parentPanel, helpTipText) {
        var messageRow = parentPanel.add("group");
        setupRow(messageRow, "fill", 0);
        messageRow.margins = MESSAGE_MARGINS;
        /* 幅いっぱいの箱にして中央揃えで描く。中身の幅で中央寄せすると、文字数が増えたとき収まらない
           Give it the full width and center the text inside; sizing to the content would clip a longer message */
        var messageText = messageRow.add('statictext {justify: "center"}');
        messageText.alignment = ["fill", "center"];
        if (helpTipText) {
            messageText.helpTip = helpTipText;
        }
        return messageText;
    }

    /**
     * ダイアログウィンドウを生成する（縦積みの構成）
     * @param {string} windowTitle - ウィンドウタイトル
     * @returns {Window} 生成したダイアログ
     */
    function createDialogWindow(windowTitle) {
        var dialogWindow = new Window("dialog", windowTitle);
        dialogWindow.orientation = "column";
        dialogWindow.alignChildren = ["fill", "top"];
        dialogWindow.spacing = WINDOW_SPACING;
        dialogWindow.margins = WINDOW_MARGINS;
        return dialogWindow;
    }

    // =========================================
    // 入力欄の補助 / Input field helpers
    // =========================================

    /**
     * 入力欄に↑↓キーでの値増減を追加する（Shiftで±10・10の倍数スナップ、Optionで±0.1）
     * @param {EditText} editText - 対象の入力欄
     * @returns {void}
     */
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 when Shift is held */
                if (event.keyName === "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                /* Optionキー押下時は0.1単位で増減 / Step by 0.1 when Option is held */
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName === "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName === "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                /* 小数第1位までに丸め / Round to one decimal place */
                value = Math.round(value * 10) / 10;
            } else {
                /* 整数に丸め / Round to an integer */
                value = Math.round(value);
            }

            editText.text = value;
        });
    }

    // =========================================
    // レイヤー操作 / Layer handling
    // =========================================

    /**
     * 名前でレイヤーを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {string} layerName - 探すレイヤー名
     * @returns {Layer|null} 見つかったレイヤー（なければ null）
     */
    function findLayerByName(doc, layerName) {
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === layerName) {
                return doc.layers[i];
            }
        }
        return null;
    }

    /**
     * レイヤー名の入力値を整える（前後の空白を落とし、空なら既定名に戻す）
     * @param {string} inputText - 入力欄の値
     * @returns {string} 使用するレイヤー名
     */
    function normalizeLayerName(inputText) {
        var trimmedName = (inputText || "").replace(/^\s+|\s+$/g, "");
        return trimmedName || OUTPUT_LAYER_NAME;
    }

    /**
     * 長方形の作成先レイヤーを用意してアクティブにする（メッセージは出さない）
     * @param {Document} doc - 対象ドキュメント
     * @param {GuideRectOptions} options - ダイアログの設定値
     * @returns {{layer: Layer, created: boolean}|null} 用意できた作成先（用意できなければ null）
     */
    function activateOutputLayer(doc, options) {
        if (!options.useOutputLayer) {
            /* 現在のレイヤーに描くので、ロック・非表示なら描けない / Drawing into the current layer, so a locked or hidden one is unusable */
            var activeLayer = doc.activeLayer;
            if (activeLayer.locked || !activeLayer.visible) {
                return null;
            }
            return { layer: activeLayer, created: false };
        }
        /* 同名レイヤーがあれば再利用し、なければ作る / Reuse a layer with the same name, or create one */
        var outputLayer = findLayerByName(doc, options.outputLayerName);
        var wasCreated = false;
        if (!outputLayer) {
            outputLayer = doc.layers.add();
            outputLayer.name = options.outputLayerName;
            wasCreated = true;
        }
        /* 使い回すレイヤーは描ける状態に戻す / Make a reused layer drawable again */
        outputLayer.locked = false;
        outputLayer.visible = true;
        doc.activeLayer = outputLayer;
        return { layer: outputLayer, created: wasCreated };
    }

    /**
     * 作成先レイヤーを用意し、用意できなければ理由を知らせる
     * @param {Document} doc - 対象ドキュメント
     * @param {GuideRectOptions} options - ダイアログの設定値
     * @returns {boolean} 作成先を確保できたら true
     */
    function prepareOutputLayer(doc, options) {
        if (activateOutputLayer(doc, options)) {
            return true;
        }
        alert(getLabel('alert', 'lockedActiveLayer'));
        return false;
    }

    /**
     * ロックされているレイヤーを一時的に解除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer[]} 解除したレイヤー（復元用）
     */
    function unlockLockedLayers(doc) {
        var unlockedLayers = [];
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].locked) {
                doc.layers[i].locked = false;
                unlockedLayers.push(doc.layers[i]);
            }
        }
        return unlockedLayers;
    }

    /**
     * 一時解除したレイヤーのロックを元に戻す
     * @param {Layer[]} layers - 復元するレイヤー
     * @returns {void}
     */
    function relockLayers(layers) {
        for (var i = 0; i < layers.length; i++) {
            layers[i].locked = true;
        }
    }

    // =========================================
    // ガイドの収集と分類 / Collecting and classifying guides
    // =========================================

    /**
     * ガイドパスの向き・位置・長さを求める（縦横どちらでもないものは対象外）
     * @param {PathItem} guidePath - 対象のガイドパス
     * @returns {{isVertical: boolean, position: number, length: number}|null} 幾何情報（対象外なら null）
     */
    function getGuideGeometry(guidePath) {
        if (guidePath.pathPoints.length < 2) {
            return null;
        }
        var startAnchor = guidePath.pathPoints[0].anchor;
        var endAnchor = guidePath.pathPoints[1].anchor;
        var deltaX = endAnchor[0] - startAnchor[0];
        var deltaY = endAnchor[1] - startAnchor[1];
        var isVertical = Math.abs(deltaX) < COORD_TOLERANCE;
        if (!isVertical && Math.abs(deltaY) >= COORD_TOLERANCE) {
            return null;
        }
        return {
            isVertical: isVertical,
            position: isVertical ? startAnchor[0] : startAnchor[1],
            length: Math.sqrt(deltaX * deltaX + deltaY * deltaY)
        };
    }

    /**
     * アートボードをまたぐ長さのガイドか（＝ルーラーガイドとみなせるか）を判定する
     * @param {PathItem} guidePath - 判定するガイドパス
     * @param {number} artboardWidth - アートボードの幅（pt）
     * @param {number} artboardHeight - アートボードの高さ（pt）
     * @returns {boolean} ルーラーガイドとみなせるなら true
     */
    function isRulerGuide(guidePath, artboardWidth, artboardHeight) {
        /* 面積を持たない2点の直線だけが対象 / Only a two-point line with no area qualifies */
        if (guidePath.pathPoints.length !== 2 || Math.abs(guidePath.area) >= COORD_TOLERANCE) {
            return false;
        }
        var geometry = getGuideGeometry(guidePath);
        if (!geometry) {
            return false;
        }
        var artboardSpan = geometry.isVertical ? artboardHeight : artboardWidth;
        return geometry.length > artboardSpan + RULER_GUIDE_MARGIN_PT;
    }

    /**
     * ガイドが現在のアートボードの範囲内にあるかを判定する
     * @param {PathItem} guidePath - 判定するガイドパス
     * @param {number[]} artboardRect - アートボードの矩形 [左, 上, 右, 下]
     * @returns {boolean} 範囲内なら true
     */
    function isWithinArtboard(guidePath, artboardRect) {
        var geometry = getGuideGeometry(guidePath);
        if (!geometry) {
            return false;
        }
        /* 縦ガイドは左右、横ガイドは上下の範囲で判定する（Yは上が正）/ Verticals are bounded left-right, horizontals top-bottom (Y grows upward) */
        return geometry.isVertical
            ? (geometry.position >= artboardRect[0] && geometry.position <= artboardRect[2])
            : (geometry.position <= artboardRect[1] && geometry.position >= artboardRect[3]);
    }

    /**
     * 条件に合うガイドパスを集める
     * @param {Document} doc - 対象ドキュメント
     * @param {GuideRectOptions} options - ダイアログの設定値
     * @returns {PathItem[]} 集めたガイドパス
     */
    function collectGuidePaths(doc, options) {
        var artboardRect = doc.artboards[doc.artboards.getActiveArtboardIndex()].artboardRect;
        var artboardWidth = artboardRect[2] - artboardRect[0];
        var artboardHeight = artboardRect[1] - artboardRect[3];

        /* 「現在のレイヤーのみ」はレイヤー配下だけを列挙する（レイヤー同士の比較を避けられる）/ Enumerate just the active layer, which avoids comparing layer objects */
        /* pathItems は都度の DOM アクセスが重いので参照と件数を控える / Cache the collection and its length; each DOM access is costly */
        var pathItems = options.activeLayerOnly ? doc.activeLayer.pathItems : doc.pathItems;
        var guidePaths = [];
        for (var i = 0, itemCount = pathItems.length; i < itemCount; i++) {
            var pathItem = pathItems[i];
            if (!pathItem.guides || pathItem.locked || pathItem.hidden) continue;
            if (pathItem.layer.locked && !options.includeLocked) continue;
            if (options.useRulerOnly && !isRulerGuide(pathItem, artboardWidth, artboardHeight)) continue;
            if (options.activeArtboardOnly && !isWithinArtboard(pathItem, artboardRect)) continue;
            guidePaths.push(pathItem);
        }
        return guidePaths;
    }

    /**
     * ソート済みの座標配列から、重複とみなせる隣接値を取り除く
     * @param {number[]} sortedPositions - ソート済みの座標配列
     * @returns {number[]} 重複を除いた配列
     */
    function dedupeSortedPositions(sortedPositions) {
        var uniquePositions = [];
        for (var i = 0; i < sortedPositions.length; i++) {
            var previous = uniquePositions[uniquePositions.length - 1];
            if (i === 0 || Math.abs(sortedPositions[i] - previous) >= COORD_TOLERANCE) {
                uniquePositions.push(sortedPositions[i]);
            }
        }
        return uniquePositions;
    }

    /**
     * ガイドを縦・横に分類して座標を返す
     * @param {PathItem[]} guidePaths - 分類するガイドパス
     * @returns {{verticals: number[], horizontals: number[]}} 縦ガイドのX座標（昇順）と横ガイドのY座標（降順）
     */
    function classifyGuidePositions(guidePaths) {
        var verticalXs = [];
        var horizontalYs = [];
        for (var i = 0; i < guidePaths.length; i++) {
            var geometry = getGuideGeometry(guidePaths[i]);
            if (!geometry) continue;
            (geometry.isVertical ? verticalXs : horizontalYs).push(geometry.position);
        }
        /* 縦は左→右、横は上→下の順に並べる / Order verticals left to right and horizontals top to bottom */
        verticalXs.sort(function(a, b) {
            return a - b;
        });
        horizontalYs.sort(function(a, b) {
            return b - a;
        });
        /* 同座標に重なったガイドは1本として扱う（つぶれた長方形を作らない）/ Collapse overlapping guides so no zero-size rectangle is produced */
        return {
            verticals: dedupeSortedPositions(verticalXs),
            horizontals: dedupeSortedPositions(horizontalYs)
        };
    }

    /**
     * オプションに従ってガイドを集め、縦横の座標に分類する
     * @param {Document} doc - 対象ドキュメント
     * @param {GuideRectOptions} options - ダイアログの設定値
     * @returns {{verticals: number[], horizontals: number[]}} 縦ガイドのX座標（昇順）と横ガイドのY座標（降順）
     */
    function getGuidePositions(doc, options) {
        /* ロックされたレイヤー上のガイドも拾えるよう、収集の間だけロックを外す / Unlock during collection so guides on locked layers are visible to the loop */
        var relockTargets = options.includeLocked ? unlockLockedLayers(doc) : [];
        var guidePaths = collectGuidePaths(doc, options);
        relockLayers(relockTargets);
        return classifyGuidePositions(guidePaths);
    }

    // =========================================
    // 長方形の生成 / Creating the rectangles
    // =========================================

    /**
     * 長方形の塗り色を生成する（RGBは黒、それ以外はK100）
     * @param {DocumentColorSpace} docColorSpace - ドキュメントのカラースペース
     * @returns {RGBColor|CMYKColor} 塗り色
     */
    function createFillColor(docColorSpace) {
        if (docColorSpace === DocumentColorSpace.RGB) {
            var rgbColor = new RGBColor();
            rgbColor.red = 0;
            rgbColor.green = 0;
            rgbColor.blue = 0;
            return rgbColor;
        }
        var cmykColor = new CMYKColor();
        cmykColor.cyan = 0;
        cmykColor.magenta = 0;
        cmykColor.yellow = 0;
        cmykColor.black = 100;
        return cmykColor;
    }

    /**
     * ガイドの座標から、作成される長方形の数を求める
     * @param {{verticals: number[], horizontals: number[]}} guidePositions - 縦横ガイドの座標
     * @returns {number} 作成される長方形の数（足りなければ0）
     */
    function countRectangles(guidePositions) {
        var verticalCount = guidePositions.verticals.length;
        var horizontalCount = guidePositions.horizontals.length;
        if (verticalCount < 2 || horizontalCount < 2) {
            return 0;
        }
        return (verticalCount - 1) * (horizontalCount - 1);
    }

    /**
     * 縦横ガイドで区切られた区画それぞれに長方形を作成する
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} verticalXs - 縦ガイドのX座標（昇順）
     * @param {number[]} horizontalYs - 横ガイドのY座標（降順）
     * @param {number} offsetPt - 四辺を広げる量（pt、マイナスで縮む）
     * @returns {PathItem[]} 作成した長方形
     */
    function createRectangles(doc, verticalXs, horizontalYs, offsetPt) {
        /* 色は使い回す（アイテムごとに新規生成する必要はない）/ Reuse one color object for every rectangle */
        var fillColor = createFillColor(doc.documentColorSpace);
        var createdRects = [];
        for (var xi = 0; xi < verticalXs.length - 1; xi++) {
            for (var yi = 0; yi < horizontalYs.length - 1; yi++) {
                var rectLeft = verticalXs[xi] - offsetPt;
                var rectTop = horizontalYs[yi] + offsetPt;
                var rectWidth = verticalXs[xi + 1] - verticalXs[xi] + offsetPt * 2;
                var rectHeight = horizontalYs[yi] - horizontalYs[yi + 1] + offsetPt * 2;
                /* マイナスのオフセットで潰れる区画は作らない / Skip a cell that a negative offset collapses */
                if (rectWidth <= 0 || rectHeight <= 0) continue;
                var rectangle = doc.pathItems.rectangle(rectTop, rectLeft, rectWidth, rectHeight);
                rectangle.stroked = false;
                rectangle.filled = true;
                rectangle.fillColor = fillColor;
                rectangle.opacity = RECT_OPACITY;
                createdRects.push(rectangle);
            }
        }
        return createdRects;
    }

    /**
     * 指定したアイテムだけを選択状態にする
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} items - 選択するアイテム
     * @returns {void}
     */
    function selectOnly(doc, items) {
        doc.selection = null;
        for (var i = 0; i < items.length; i++) {
            items[i].selected = true;
        }
    }

    /**
     * 選択中の長方形を1つのパスに結合する
     * @returns {void}
     */
    function mergeSelectedRectangles() {
        app.executeMenuCommand('group');
        app.executeMenuCommand('Live Pathfinder Add');
        app.executeMenuCommand('expandStyle');
        app.executeMenuCommand('ungroup');
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * @typedef {object} GuideRectOptions
     * @property {boolean} useRulerOnly - ルーラーガイドのみを対象にするなら true
     * @property {boolean} activeLayerOnly - 現在のレイヤーのガイドだけを対象にするなら true
     * @property {boolean} activeArtboardOnly - 現在のアートボード内のガイドだけを対象にするなら true
     * @property {boolean} includeLocked - ロックされたレイヤー上のガイドも含めるなら true
     * @property {boolean} useOutputLayer - 指定レイヤーに長方形を作成するなら true（false なら現在のレイヤー）
     * @property {string} outputLayerName - 指定レイヤーの名前（同名があれば再利用）
     * @property {number} offsetPt - 各長方形の四辺を広げる量（pt、マイナスで縮む。オフセット未使用なら0）
     * @property {boolean} mergeAdjacent - 隣り合う長方形を結合するなら true
     * @property {boolean} convertToShape - 作成した長方形をライブシェイプに変換するなら true
     */

    /**
     * オプションダイアログを表示する
     * @param {Document} doc - 作成予定数の集計に使うドキュメント
     * @returns {GuideRectOptions|null} 設定値（キャンセル時は null）
     */
    function showOptionsDialog(doc) {
        var dialog = createDialogWindow(getLabel('dialog', 'title') + ' ' + SCRIPT_VERSION);

        /* 上段：左＝対象となるガイド、右＝作成する長方形 / Top area: target guides on the left, rectangles on the right */
        var columnsRow = dialog.add("group");
        columnsRow.orientation = "row";
        columnsRow.alignChildren = ["fill", "top"];
        columnsRow.spacing = COLUMN_SPACING;

        var guideSourcePanel = addPanel(columnsRow, getLabel('guideSource', 'panelTitle'));
        /* 各カラムは中身の高さのまま、上端をそろえて並べる / Each column keeps its own height and lines up at the top */
        guideSourcePanel.alignment = ["fill", "top"];

        /* 見つかったガイドの本数（絞り込みの結果がすぐ分かるよう最上部に置く）/ How many guides matched, kept at the top so the filter result is visible at once */
        var guideCountText = addPanelMessage(guideSourcePanel, getLabel('tooltip', 'guideCount'));

        var guideTypePanel = addPanel(guideSourcePanel, getLabel('guideType', 'panelTitle'));
        var guideTypeRadios = addRadioPair(guideTypePanel,
            getLabel('guideType', 'all'), getLabel('guideType', 'rulerOnly'), 0,
            getLabel('tooltip', 'guideType'));
        var rulerGuidesRadio = guideTypeRadios[1];

        var targetLayerPanel = addPanel(guideSourcePanel, getLabel('targetLayer', 'panelTitle'));
        var targetLayerRadios = addRadioPair(targetLayerPanel,
            getLabel('targetLayer', 'allLayers'), getLabel('targetLayer', 'activeOnly'), 0,
            getLabel('tooltip', 'targetLayer'));
        var activeLayerRadio = targetLayerRadios[1];

        var guideOptionPanel = addPanel(guideSourcePanel, getLabel('guideOption', 'panelTitle'));
        var activeArtboardCheckbox = addPanelCheckbox(guideOptionPanel,
            getLabel('guideOption', 'activeArtboardOnly'), false, getLabel('tooltip', 'activeArtboardOnly'));
        var includeLockedCheckbox = addPanelCheckbox(guideOptionPanel,
            getLabel('guideOption', 'includeLocked'), true, getLabel('tooltip', 'includeLocked'));

        var rectanglePanel = addPanel(columnsRow, getLabel('rectangle', 'panelTitle'));
        rectanglePanel.alignment = ["fill", "top"];

        /* 作成予定数（左カラムのガイド本数と同じく、パネルの最上部に置く）/ Expected count, kept at the top like the guide count in the left column */
        var summaryText = addPanelMessage(rectanglePanel, getLabel('tooltip', 'summaryCount'));

        var destinationPanel = addPanel(rectanglePanel, getLabel('destination', 'panelTitle'));
        var destinationRadios = addRadioPair(destinationPanel,
            getLabel('destination', 'activeLayer'), getLabel('destination', 'outputLayer'), 0,
            getLabel('tooltip', 'destination'));
        var outputLayerRadio = destinationRadios[1];

        /* 「指定レイヤー」の名前はラジオの次の行に字下げして置く / The layer name sits on the line below its radio, indented */
        var outputLayerNameRow = destinationPanel.add("group");
        setupRow(outputLayerNameRow, "fill", FIELD_ROW_SPACING);
        outputLayerNameRow.margins = INDENT_MARGINS;
        var outputLayerNameInput = outputLayerNameRow.add('edittext {characters: ' + LAYER_NAME_CHARS + '}');
        outputLayerNameInput.text = OUTPUT_LAYER_NAME;
        /* 余った幅で伸ばす（characters を増やすとダイアログごと広がる）/ Stretch into the leftover width; raising characters would widen the dialog */
        outputLayerNameInput.alignment = ["fill", "center"];
        outputLayerNameInput.helpTip = getLabel('tooltip', 'outputLayerName');

        var rectangleOptionPanel = addPanel(rectanglePanel, getLabel('rectangleOption', 'panelTitle'));

        /* オフセット：チェックボックス＋数値欄＋ルーラー単位 / Offset: checkbox, value field, and the ruler unit */
        var rulerUnit = getRulerUnit();
        var offsetRow = rectangleOptionPanel.add("group");
        setupRow(offsetRow, "left", FIELD_ROW_SPACING);
        var offsetCheckbox = offsetRow.add("checkbox", undefined, getLabel('rectangleOption', 'offset'));
        offsetCheckbox.helpTip = getLabel('tooltip', 'offset');
        var offsetInput = offsetRow.add('edittext {characters: ' + OFFSET_CHARS + '}');
        offsetInput.text = "0";
        offsetInput.helpTip = getLabel('tooltip', 'offset');
        changeValueByArrowKey(offsetInput);
        offsetRow.add("statictext", undefined, rulerUnit.label);

        var mergeAdjacentCheckbox = addPanelCheckbox(rectangleOptionPanel,
            getLabel('rectangleOption', 'mergeAdjacent'), false, getLabel('tooltip', 'mergeAdjacent'));
        var convertToShapeCheckbox = addPanelCheckbox(rectangleOptionPanel,
            getLabel('rectangleOption', 'convertToShape'), true, getLabel('tooltip', 'convertToShape'));

        /**
         * 指定レイヤーを選んでいるときだけ名前欄を使えるようにする
         * @returns {void}
         */
        function syncOutputLayerField() {
            outputLayerNameInput.enabled = outputLayerRadio.value;
        }

        /**
         * オフセットにチェックが入っているときだけ数値欄を使えるようにする
         * @returns {void}
         */
        function syncOffsetField() {
            offsetInput.enabled = offsetCheckbox.value;
        }

        /**
         * 現在の入力内容を設定値として読み取る
         * @returns {GuideRectOptions} 設定値
         */
        function readOptions() {
            return {
                useRulerOnly: rulerGuidesRadio.value,
                activeLayerOnly: activeLayerRadio.value,
                activeArtboardOnly: activeArtboardCheckbox.value,
                includeLocked: includeLockedCheckbox.value,
                useOutputLayer: outputLayerRadio.value,
                outputLayerName: normalizeLayerName(outputLayerNameInput.text),
                offsetPt: offsetCheckbox.value ? convertToPt(offsetInput.text, rulerUnit.factor) : 0,
                mergeAdjacent: mergeAdjacentCheckbox.value,
                convertToShape: convertToShapeCheckbox.value
            };
        }

        /**
         * ガイドの本数と作成予定数を各パネルの見出しに反映する
         * @param {GuideRectOptions} options - ダイアログの設定値
         * @param {{verticals: number[], horizontals: number[]}} guidePositions - 縦横ガイドの座標
         * @param {number} rectangleCount - 作成される長方形の数
         * @returns {void}
         */
        function updateSummary(options, guidePositions, rectangleCount) {
            guideCountText.text = formatMessage(getLabel('summary', 'guides'), {
                vertical: guidePositions.verticals.length,
                horizontal: guidePositions.horizontals.length
            });
            /* 判定は main() の結合条件と同じにそろえる / Mirror the merge condition used in main() */
            var willMerge = options.mergeAdjacent && rectangleCount > 1;
            summaryText.text = formatMessage(getLabel('summary', willMerge ? 'merged' : 'count'), {
                count: rectangleCount
            });
        }

        var buttonBarGroup = dialog.add("group");
        setupRow(buttonBarGroup, "fill", BUTTON_BAR_SPACING);
        buttonBarGroup.margins = BUTTON_BAR_MARGINS;
        var previewCheckbox = buttonBarGroup.add("checkbox", undefined, getLabel('dialog', 'preview'));
        previewCheckbox.value = true;
        previewCheckbox.helpTip = getLabel('tooltip', 'preview');

        /* スペーサー：ボタンを右端へ押し出す / Spacer that pushes the buttons to the right edge */
        var buttonBarSpacer = buttonBarGroup.add("group");
        buttonBarSpacer.alignment = ["fill", "fill"];
        buttonBarSpacer.minimumSize.width = 0;

        var dialogButtonGroup = buttonBarGroup.add("group");
        setupRow(dialogButtonGroup, "right", BUTTON_BAR_SPACING);
        dialogButtonGroup.add("button", undefined, getLabel('button', 'cancel'), { name: "cancel" });
        dialogButtonGroup.add("button", undefined, getLabel('button', 'ok'), { name: "ok" });

        /* プレビューで描いた長方形と、そのために新規作成したレイヤー / Previewed rectangles and any layer created just for them */
        var previewRects = null;
        var previewLayer = null;

        /**
         * プレビューの長方形を取り除く
         * @returns {void}
         */
        function clearPreview() {
            if (!previewRects) return;
            for (var i = 0; i < previewRects.length; i++) {
                var previewRect = previewRects[i];
                if (previewRect && !previewRect.locked && previewRect.layer && !previewRect.layer.locked) {
                    previewRect.remove();
                }
            }
            previewRects = null;
        }

        /**
         * プレビューのために作ったレイヤーを、空のままなら片付ける
         * @returns {void}
         */
        function removePreviewLayer() {
            if (!previewLayer) return;
            /* 最後の1枚は削除できないので残す / The last layer cannot be removed, so leave it */
            if (previewLayer.pageItems.length === 0 && doc.layers.length > 1) {
                previewLayer.remove();
            }
            previewLayer = null;
        }

        /**
         * 現在の設定でプレビューを描き直す
         * @param {GuideRectOptions} options - ダイアログの設定値
         * @param {{verticals: number[], horizontals: number[]}} guidePositions - 縦横ガイドの座標
         * @param {number} rectangleCount - 作成される長方形の数
         * @returns {void}
         */
        function updatePreview(options, guidePositions, rectangleCount) {
            clearPreview();
            /* 結合・シェイプ変換はメニューコマンドで重いため、プレビューでは適用しない
               Merging and shape conversion run as menu commands, so the preview leaves them out */
            if (previewCheckbox.value && rectangleCount > 0 && rectangleCount <= PREVIEW_MAX_RECTS) {
                var outputTarget = activateOutputLayer(doc, options);
                if (outputTarget) {
                    if (outputTarget.created) {
                        previewLayer = outputTarget.layer;
                    }
                    previewRects = createRectangles(doc, guidePositions.verticals, guidePositions.horizontals, options.offsetPt);
                }
            }
            app.redraw();
        }

        /**
         * 設定を読み直し、件数表示とプレビューをまとめて更新する
         * @returns {void}
         */
        function refresh() {
            /* 入力欄の有効・無効もここでまとめて反映する（ハンドラーを1本にして取りこぼしを防ぐ）
               Enable/disable the fields here too, so a single handler covers everything */
            syncOutputLayerField();
            syncOffsetField();
            var options = readOptions();
            var guidePositions = getGuidePositions(doc, options);
            var rectangleCount = countRectangles(guidePositions);
            updateSummary(options, guidePositions, rectangleCount);
            updatePreview(options, guidePositions, rectangleCount);
        }

        /* 結果の見込みが変わる操作はすべて反映し直す / Refresh after anything that changes the expected result */
        var refreshControls = guideTypeRadios.concat(targetLayerRadios, destinationRadios);
        refreshControls.push(activeArtboardCheckbox, includeLockedCheckbox, mergeAdjacentCheckbox,
            offsetCheckbox, previewCheckbox);
        for (var i = 0; i < refreshControls.length; i++) {
            /* onClick と onChange の両方に付ける（チェックボックスはどちらで届くかが環境で異なる）
               Bind both events; which one a checkbox delivers varies by platform */
            refreshControls[i].onClick = refresh;
            refreshControls[i].onChange = refresh;
        }
        offsetInput.addEventListener("changing", refresh);
        outputLayerNameInput.addEventListener("changing", refresh);
        refresh();

        var dialogResult = dialog.show();
        /* 本番の長方形は main() が作り直すので、プレビューは必ず片付ける / main() draws the real rectangles, so the preview always goes away */
        clearPreview();
        if (dialogResult !== 1) {
            removePreviewLayer();
            app.redraw();
            return null;
        }
        return readOptions();
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * ガイドの交点で区切られた区画に長方形を作成する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert', 'noDocument'));
            return;
        }

        var doc = app.activeDocument;
        var options = showOptionsDialog(doc);
        if (!options) return;

        var guidePositions = getGuidePositions(doc, options);
        if (countRectangles(guidePositions) === 0) {
            alert(getLabel('alert', 'notEnoughGuides'));
            return;
        }

        if (!prepareOutputLayer(doc, options)) return;

        /* selectOnly が元の選択を解除するので、後続のメニューコマンドは生成分にだけ効く / selectOnly clears the previous selection, so the menu commands below act only on the new rectangles */
        var rectangles = createRectangles(doc, guidePositions.verticals, guidePositions.horizontals, options.offsetPt);
        if (rectangles.length === 0) return;
        selectOnly(doc, rectangles);

        if (options.mergeAdjacent && rectangles.length > 1) {
            mergeSelectedRectangles();
        }
        if (options.convertToShape) {
            app.executeMenuCommand('Convert to Shape');
        }
    }

    main();

})();
