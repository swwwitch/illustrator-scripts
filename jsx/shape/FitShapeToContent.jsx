#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
#targetengine "FitShapeToContentSession"

/*

### 概要

テキストやグループに合わせて、背面の座布団形状を手早く作成・調整します。
図形も一緒に選ぶ（またはテキストと図形のグループを選ぶ）と、その図形を座布団として使います。

詳細は README を参照してください。

### Overview

Quickly creates and adjusts a backing shape that fits a text frame or a group.
Add a shape to the selection, or select a text+shape group, to reuse that shape instead of creating a rectangle.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FitShapeToContent";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v2.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-25";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-31";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FitShapeToContent.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FitShapeToContent.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n6e4a6a2b175f"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 自動作成した座布団の初期不透明度（%）/ Initial opacity of the auto-created backing shape (%) */
    var AUTO_CREATED_SHAPE_OPACITY = 20;

    /* パディングの初期値（pt）。定規の単位に換算して入力欄に入れる / Default padding (pt), shown in the ruler unit */
    var DEFAULT_PADDING_PT = 20;

    /* プレビュー図形の最小サイズ（pt）。0以下は Illustrator が受け付けない / Minimum preview size (pt) */
    var MIN_PREVIEW_SIZE = 0.1;

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 6;                /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING     = 12;               /* 2カラムの間隔 / gap between columns */
    var FIELD_LABEL_WIDTH  = 50;               /* 項目名の幅 / width of a field label */
    var FIELD_CHARS        = 5;                /* 数値入力欄の文字数 / character width of a numeric field */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */

    // =========================================
    // セッション記憶 / Session state
    // =========================================

    /* #targetengine 内に残す設定の保存キー / Storage key kept alive by #targetengine */
    var SESSION_KEY = "__FitShapeToContentSession__";

    /**
     * 前回のダイアログ設定を取得する（初回は既定値）
     * @param {number} unitFactor - 表示単位1つあたりの pt 数
     * @returns {object} 保存済みの設定
     */
    function getSessionState(unitFactor) {
        if (!$.global[SESSION_KEY]) {
            var defaultPadding = formatUnitValue(DEFAULT_PADDING_PT, unitFactor);
            $.global[SESSION_KEY] = {
                addW: defaultPadding,
                addH: defaultPadding,
                radius: "0",
                radiusEnabled: true,
                link: true,
                pill: false
            };
        }
        return $.global[SESSION_KEY];
    }

    /**
     * ダイアログ設定を保存する
     * @param {object} state - 保存する設定
     * @returns {void}
     */
    function saveSessionState(state) {
        $.global[SESSION_KEY] = state;
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在の表示言語を取得する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        var localeText = ($.locale || "") + "";
        return (localeText.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "座布団メーカー", en: "Fit Shape to Content" }
        },
        panel: {
            padding: { ja: "パディング", en: "Padding" },
            corner:  { ja: "角丸", en: "Rounded Corners" }
        },
        fieldLabel: {
            width:  { ja: "幅", en: "Width" },
            height: { ja: "高さ", en: "Height" },
            radius: { ja: "半径", en: "Radius" }
        },
        checkbox: {
            adjustEnabled: { ja: "座布団の調整", en: "Adjust Shape" },
            link:          { ja: "連動", en: "Link" },
            pill:          { ja: "ピル形状", en: "Pill Shape" }
        },
        button: {
            ok:     { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            selectError:   { ja: "選択エラー", en: "Selection Error" },
            invalidNumber: { ja: "数値を入力してください。", en: "Enter a numeric value." },
            invalidRadius: { ja: "角丸の半径は0以上の数値を入力してください。", en: "Enter a radius value of 0 or greater." },
            selectOne: {
                ja: "テキストまたはグループを1つ、もしくはテキスト/グループと図形を計2つ選択して実行してください。",
                en: "Select one text/group item, or one text/group and one shape (2 items total)."
            },
            clippingGroup: {
                ja: "クリッピンググループの計測には未対応です。クリッピングを解除するか、計測対象を単純なグループにしてください。",
                en: "Clipping groups are not supported for measurement. Release the clipping mask or use a simple group as the content item."
            },
            measureFailed: {
                ja: "コンテンツの計測に失敗しました。選択内容を確認してください。",
                en: "Failed to measure the content item. Check the selected objects and try again."
            }
        }
    };

    /**
     * LABELS から現在の言語のラベルを取得する
     * @param {...string} keyPath - LABELS を辿るキー（例: "alert", "selectOne"）
     * @returns {string} ラベル文字列（未定義なら空文字）
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
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {...string} keyPath - LABELS を辿るキー
     * @returns {string} コロンを付けたラベル
     */
    function labelText() {
        return getLabel.apply(null, arguments) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 選択エラーの警告を表示する
     * @param {string} messageKey - LABELS.alert 配下のメッセージキー
     * @returns {void}
     */
    function alertSelectionError(messageKey) {
        alert(getLabel("alert", messageKey), getLabel("alert", "selectError"));
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位テーブル（配列の添字が rulerType コードと一致：0=in, 1=mm, 2=pt …）/ Unit table; the array index equals the rulerType code */
    var UNITS = [
        { label: "in",    factor: 72.0 },
        { label: "mm",    factor: 72.0 / 25.4 },
        { label: "pt",    factor: 1.0 },
        { label: "pica",  factor: 12.0 },
        { label: "cm",    factor: 72.0 / 2.54 },
        { label: "H",     factor: 72.0 / 25.4 * 0.25 },
        { label: "px",    factor: 1.0 },
        { label: "ft/in", factor: 72.0 * 12.0 },
        { label: "m",     factor: 72.0 / 25.4 * 1000.0 },
        { label: "yd",    factor: 72.0 * 36.0 },
        { label: "ft",    factor: 72.0 * 12.0 }
    ];

    /**
     * 定規の単位（表示ラベルと pt 換算係数）を取得する
     * @returns {{label: string, factor: number}} 単位情報（不明なら pt）
     */
    function getRulerUnit() {
        return UNITS[app.preferences.getIntegerPreference("rulerType")] || UNITS[2];
    }

    /**
     * pt 値を表示単位の文字列に変換する（小数第2位まで）
     * @param {number} valueInPt - pt 単位の値
     * @param {number} unitFactor - 表示単位1つあたりの pt 数
     * @returns {string} 表示用の文字列
     */
    function formatUnitValue(valueInPt, unitFactor) {
        return String(Math.round((valueInPt / unitFactor) * 100) / 100);
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ウィンドウに共通レイアウトを適用する
     * @param {Window} targetWindow - 対象ウィンドウ
     * @returns {void}
     */
    function setupWindow(targetWindow) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = ["fill", "top"];
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = WINDOW_SPACING;
    }

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
     * @param {Window|Group} parentContainer - 追加先
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
     * 右寄せのボタンバーを生成する
     * @param {Window} targetWindow - 追加先のウィンドウ
     * @returns {{btnCancel: Button, btnOK: Button}} 生成したボタン
     */
    function addButtonBar(targetWindow) {
        var btnRowGroup = targetWindow.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.margins = BUTTON_BAR_MARGINS;

        /* 伸縮スペーサーで右側グループを右端へ押し出す / A stretchable spacer pushes the buttons right */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        return {
            btnCancel: btnRightGroup.add("button", undefined, getLabel("button", "cancel"), { name: "cancel" }),
            btnOK: btnRightGroup.add("button", undefined, getLabel("button", "ok"), { name: "ok" })
        };
    }

    /**
     * 「項目名＋数値入力欄＋単位」の行を生成する
     * @param {Group|Panel} parentContainer - 追加先
     * @param {string} fieldLabelText - 右揃えで表示する項目名
     * @param {string} initialText - 入力欄の初期値
     * @param {string} unitLabel - 入力欄の右に添える単位
     * @returns {EditText} 生成した入力欄
     */
    function addNumericFieldRow(parentContainer, fieldLabelText, initialText, unitLabel) {
        var fieldRow = parentContainer.add("group");
        setupRow(fieldRow);

        var fieldLabel = fieldRow.add("statictext", undefined, fieldLabelText);
        fieldLabel.preferredSize = [FIELD_LABEL_WIDTH, -1];
        fieldLabel.justify = "right";

        var inputField = fieldRow.add("edittext", undefined, initialText);
        inputField.characters = FIELD_CHARS;

        fieldRow.add("statictext", undefined, unitLabel);
        return inputField;
    }

    // =========================================
    // 汎用ヘルパー / Generic helpers
    // =========================================

    /**
     * 例外を握り潰して処理を実行する
     * @param {function} fn - 実行する処理
     * @returns {void}
     */
    function safeDo(fn) {
        try { fn(); } catch (e) { }
    }

    /**
     * 例外を握り潰してアイテムを削除する
     * @param {PageItem} item - 削除するアイテム
     * @returns {void}
     */
    function safeRemove(item) {
        safeDo(function () { if (item) item.remove(); });
    }

    /**
     * 文字列を数値として解析し、無効なら既定値を返す
     * @param {string} value - 解析する文字列
     * @param {number} defaultValue - 解析できないときに返す値
     * @returns {number} 解析結果
     */
    function parseNumberOrDefault(value, defaultValue) {
        var num = parseFloat(value);
        return isNaN(num) ? defaultValue : num;
    }

    /**
     * 数値入力欄を検証する（負値は不可）
     * @param {EditText} editText - 対象の入力欄
     * @param {string} messageKey - LABELS.alert 配下のメッセージキー
     * @returns {number|null} 妥当な数値、不正なら null（警告表示済み）
     */
    function validateNumericField(editText, messageKey) {
        var value = parseFloat(editText.text);
        if (isNaN(value) || value < 0) {
            alert(getLabel("alert", messageKey), getLabel("dialog", "title"));
            editText.active = true;
            editText.selection = [0, editText.text.length];
            return null;
        }
        return value;
    }

    /**
     * ↑↓キーで数値を増減できるようにする
     * @param {EditText} editText - 対象の入力欄
     * @param {function} [onChanged] - 値が変わったときに呼ぶ処理
     * @param {number} [minValue] - 下限値
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onChanged, minValue) {
        editText.addEventListener("keydown", function (event) {
            if (event.keyName != "Up" && event.keyName != "Down") return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName == "Up");
            event.preventDefault();

            if (keyboard.shiftKey) {
                /* Shift：10 単位にスナップ / Shift snaps to multiples of 10 */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                /* Option：0.1 刻み / Option steps by 0.1 */
                value = Math.round((value + (isUp ? 0.1 : -0.1)) * 10) / 10;
            } else {
                value = Math.round(value + (isUp ? 1 : -1));
            }

            if (typeof minValue === "number" && value < minValue) value = minValue;

            editText.text = value;
            if (typeof onChanged === "function") onChanged();
        });
    }

    // =========================================
    // オブジェクト判定・計測 / Item checks and measurement
    // =========================================

    /**
     * @typedef {object} ContentBounds
     * @property {number} left - 左端
     * @property {number} top - 上端
     * @property {number} width - 幅
     * @property {number} height - 高さ
     * @property {number} centerX - 中心X
     * @property {number} centerY - 中心Y
     */

    /**
     * アイテムの境界情報を取得する
     * @param {PageItem} item - 対象アイテム
     * @returns {ContentBounds} 境界情報
     */
    function getBoundsFromItem(item) {
        var vb = item.visibleBounds;
        return {
            left: vb[0],
            top: vb[1],
            width: vb[2] - vb[0],
            height: vb[1] - vb[3],
            centerX: vb[0] + ((vb[2] - vb[0]) / 2),
            centerY: vb[1] - ((vb[1] - vb[3]) / 2)
        };
    }

    /**
     * 図形の幾何学的中心（線幅・効果を除いたパスの中心）を指定座標に合わせる
     * @param {PageItem} item - 対象アイテム
     * @param {number} centerX - 合わせ先の中心X
     * @param {number} centerY - 合わせ先の中心Y
     * @returns {void}
     */
    function centerByGeometry(item, centerX, centerY) {
        var gb = item.geometricBounds;
        var geoCenterX = gb[0] + ((gb[2] - gb[0]) / 2);
        var geoCenterY = gb[1] - ((gb[1] - gb[3]) / 2);
        item.translate(centerX - geoCenterX, centerY - geoCenterY);
    }

    /**
     * パス自体（線幅を含まない）が指定サイズになるよう拡大縮小し、中心を合わせる。
     * width / height は線幅込みの実寸なので、比率を掛けてから設定する。
     * @param {PageItem} item - 対象アイテム
     * @param {number} targetWidth - パスの目標幅
     * @param {number} targetHeight - パスの目標高さ
     * @param {number} centerX - 合わせ先の中心X
     * @param {number} centerY - 合わせ先の中心Y
     * @returns {void}
     */
    function resizeAndCenterByGeometry(item, targetWidth, targetHeight, centerX, centerY) {
        var gb = item.geometricBounds;
        var geoWidth = gb[2] - gb[0];
        var geoHeight = gb[1] - gb[3];
        if (geoWidth > 0) item.width = item.width * (targetWidth / geoWidth);
        if (geoHeight > 0) item.height = item.height * (targetHeight / geoHeight);
        centerByGeometry(item, centerX, centerY);
    }

    /**
     * 座布団を敷く対象（テキスト／グループ）かどうかを判定する
     * @param {PageItem} item - 対象アイテム
     * @returns {boolean} コンテンツ対象なら true
     */
    function isContentItem(item) {
        return !!(item && (item.typename === "TextFrame" || item.typename === "GroupItem"));
    }

    /**
     * 座布団として使える図形かどうかを判定する
     * @param {PageItem} item - 対象アイテム
     * @returns {boolean} 図形なら true
     */
    function isShapeItem(item) {
        return !!(item && (item.typename === "PathItem" || item.typename === "CompoundPathItem"));
    }

    /**
     * クリッピンググループかどうかを判定する
     * @param {PageItem} item - 対象アイテム
     * @returns {boolean} クリッピンググループなら true
     */
    function isClippingGroupItem(item) {
        return !!(item && item.typename === "GroupItem" && item.clipped);
    }

    // =========================================
    // アピアランス退避・復元 / Appearance capture and restore
    // =========================================

    /**
     * カラー値を複製する
     * @param {object} color - 複製元のカラー
     * @returns {object|null} 複製したカラー
     */
    function cloneColorValue(color) {
        if (!color) return null;

        var cloned;
        switch (color.typename) {
            case "RGBColor":
                cloned = new RGBColor();
                cloned.red = color.red;
                cloned.green = color.green;
                cloned.blue = color.blue;
                return cloned;
            case "CMYKColor":
                cloned = new CMYKColor();
                cloned.cyan = color.cyan;
                cloned.magenta = color.magenta;
                cloned.yellow = color.yellow;
                cloned.black = color.black;
                return cloned;
            case "GrayColor":
                cloned = new GrayColor();
                cloned.gray = color.gray;
                return cloned;
            case "SpotColor":
                cloned = new SpotColor();
                cloned.spot = color.spot;
                cloned.tint = color.tint;
                return cloned;
            case "PatternColor":
                cloned = new PatternColor();
                cloned.pattern = color.pattern;
                return cloned;
            case "GradientColor":
                cloned = new GradientColor();
                cloned.gradient = color.gradient;
                cloned.angle = color.angle;
                cloned.length = color.length;
                cloned.matrix = color.matrix;
                cloned.origin = color.origin;
                cloned.hiliteAngle = color.hiliteAngle;
                cloned.hiliteLength = color.hiliteLength;
                return cloned;
            case "NoColor":
                return new NoColor();
            default:
                return color;
        }
    }

    /**
     * 図形の基本スタイルを退避する
     * @param {PageItem} shapeItem - 対象図形
     * @returns {object} 退避したスタイル情報
     */
    function captureShapeStyle(shapeItem) {
        return {
            filled: !!shapeItem.filled,
            fillColor: shapeItem.filled ? cloneColorValue(shapeItem.fillColor) : null,
            stroked: !!shapeItem.stroked,
            strokeColor: shapeItem.stroked ? cloneColorValue(shapeItem.strokeColor) : null,
            strokeWidth: shapeItem.strokeWidth,
            opacity: shapeItem.opacity
        };
    }

    /**
     * 退避した基本スタイルを図形に復元する
     * @param {PageItem} shapeItem - 対象図形
     * @param {object} styleInfo - captureShapeStyle() の戻り値
     * @returns {void}
     */
    function restoreShapeStyle(shapeItem, styleInfo) {
        if (!shapeItem || !styleInfo) return;

        shapeItem.opacity = styleInfo.opacity;

        shapeItem.filled = styleInfo.filled;
        if (styleInfo.filled && styleInfo.fillColor) {
            shapeItem.fillColor = cloneColorValue(styleInfo.fillColor);
        }

        shapeItem.stroked = styleInfo.stroked;
        if (styleInfo.stroked && styleInfo.strokeColor) {
            shapeItem.strokeColor = cloneColorValue(styleInfo.strokeColor);
            shapeItem.strokeWidth = styleInfo.strokeWidth;
        }
    }

    /**
     * ダイナミックアクションで「アピアランスを消去」を実行する（塗り・線・不透明度は復元）
     * @param {PageItem} targetItem - 対象アイテム
     * @returns {void}
     */
    function clearAppearanceByAction(targetItem) {
        if (!targetItem) return;

        var actionDefinition = [
            '/version 3',
            '/name [ 10',
            ' 417070656172616e6365',
            ']',
            '/isOpen 1',
            '/actionCount 1',
            '/action-1 {',
            ' /name [ 5',
            ' 636c656172',
            ' ]',
            ' /keyIndex 0',
            ' /colorIndex 0',
            ' /isOpen 1',
            ' /eventCount 1',
            ' /event-1 {',
            ' /useRulersIn1stQuadrant 0',
            ' /internalName (ai_plugin_appearance)',
            ' /localizedName [ 18',
            ' e382a2e38394e382a2e383a9e383b3e382b9',
            ' ]',
            ' /isOpen 1',
            ' /isOn 1',
            ' /hasDialog 0',
            ' /parameterCount 1',
            ' /parameter-1 {',
            ' /key 1835363957',
            ' /showInPalette 4294967295',
            ' /type (enumerated)',
            ' /name [ 27',
            ' e382a2e38394e382a2e383a9e383b3e382b9e38292e6b688e58ebb',
            ' ]',
            ' /value 6',
            ' }',
            ' }',
            '}'
        ].join('');

        var doc = app.activeDocument;
        var tempActionPath = Folder.temp.fsName + '/Appearance_clear_' + (new Date().getTime()) + '_' + Math.floor(Math.random() * 100000) + '.aia';
        var actionFile = new File(tempActionPath);
        var originalStyle = captureShapeStyle(targetItem);

        try {
            doc.selection = null;
            targetItem.selected = true;

            actionFile.open('w');
            actionFile.write(actionDefinition);
            actionFile.close();

            app.loadAction(actionFile);
            app.doScript('clear', 'Appearance', false);
            restoreShapeStyle(targetItem, originalStyle);
        } finally {
            safeDo(function () { app.unloadAction('Appearance', ''); });
            if (actionFile.exists) actionFile.remove();
            doc.selection = null;
        }
    }

    // =========================================
    // プレビュー undo ヘルパー / Preview undo helpers
    //
    // 「設定変更のたびに 前回 undo → 再生成」を回す共通パターン。
    // process() は ScriptUI 値から毎回再構築する純粋関数として書く。
    // - runPreview: UI 変更ごとに呼ぶ
    // - undoPreview: OK 直前・キャンセル・ダイアログクローズ時に前回プレビューを巻き戻す
    // =========================================

    /**
     * 直前のプレビューを巻き戻してから、プレビューを作り直す
     * @param {{isUndo: boolean}} previewState - プレビューの undo 状態
     * @param {function} processFn - プレビューを組み立てる処理
     * @returns {void}
     */
    function runPreview(previewState, processFn) {
        try {
            if (previewState.isUndo) app.undo();
            else previewState.isUndo = true;
            processFn();
            app.redraw();
        } catch (err) { }
    }

    /**
     * 残っているプレビューを巻き戻す
     * @param {{isUndo: boolean}} previewState - プレビューの undo 状態
     * @returns {void}
     */
    function undoPreview(previewState) {
        try {
            if (previewState.isUndo) app.undo();
        } catch (err) { }
        previewState.isUndo = false;
    }

    // =========================================
    // プレビュー値の計算 / Preview value math
    // =========================================

    /**
     * パディングを加えた座布団に収まる角丸半径の上限を求める（短辺の半分）
     * @param {number} addW - 幅のパディング（pt）
     * @param {number} addH - 高さのパディング（pt）
     * @param {ContentBounds} bounds - コンテンツの境界情報
     * @returns {number} 半径の上限（pt）
     */
    function getMaxRadius(addW, addH, bounds) {
        var shapeWidth = Math.max(MIN_PREVIEW_SIZE, bounds.width + addW);
        var shapeHeight = Math.max(MIN_PREVIEW_SIZE, bounds.height + addH);
        return Math.min(shapeWidth, shapeHeight) / 2;
    }

    /**
     * UI の入力値から、プレビューに使う実値（pt）と表示用テキストを求める
     * @param {{adjustEnabled:boolean, radiusEnabled:boolean, pill:boolean, addW:number, addH:number, radius:number}} uiValues - 入力欄の状態（数値は pt）
     * @param {ContentBounds} bounds - コンテンツの境界情報
     * @param {number} unitFactor - 表示単位1つあたりの pt 数
     * @returns {{addW:number, addH:number, radius:number, widthText:string, heightText:string, radiusText:string}} 計算結果
     */
    function computePreviewValues(uiValues, bounds, unitFactor) {
        var addW = uiValues.addW;
        var addH = uiValues.addH;
        var radius = (uiValues.radius < 0) ? 0 : uiValues.radius;

        if (!uiValues.adjustEnabled) {
            addW = 0;
            addH = 0;
            radius = 0;
        } else if (!uiValues.radiusEnabled) {
            radius = 0;
        } else if (uiValues.pill) {
            /* 半径＝高さの半分。幅も両端の半円分だけ広げる / Radius is half the height; widen by both caps */
            radius = Math.max(MIN_PREVIEW_SIZE, bounds.height + addH) / 2;
            addW = Math.max(MIN_PREVIEW_SIZE, radius * 2);
        } else {
            radius = Math.min(radius, getMaxRadius(addW, addH, bounds));
        }

        return {
            addW: addW,
            addH: addH,
            radius: radius,
            widthText: formatUnitValue(addW, unitFactor),
            heightText: formatUnitValue(addH, unitFactor),
            radiusText: formatUnitValue(radius, unitFactor)
        };
    }

    // =========================================
    // プレビューコントローラ / Preview controller
    // =========================================

    /**
     * @typedef {object} PreviewWidgets
     * @property {Checkbox} chkAdjustEnabled - 座布団の調整
     * @property {Checkbox} chkRadiusEnabled - 角丸の有効化
     * @property {Checkbox} chkPill - ピル形状
     * @property {Checkbox} chkLink - 幅と高さの連動
     * @property {EditText} inputW - 幅のパディング
     * @property {EditText} inputH - 高さのパディング
     * @property {EditText} inputR - 角丸の半径
     */

    /**
     * @typedef {object} PreviewTemplates
     * @property {PageItem|null} previewSourceShapeItem - 整列専用の複製（元のアピアランスを保持）
     * @property {PageItem} previewBaseShapeItem - 調整用の複製（アピアランス消去済み）
     */

    /**
     * プレビューの計算・描画・確定を束ねたコントローラを生成する
     * @param {PreviewWidgets} widgets - ダイアログの各ウィジェット
     * @param {ContentBounds} bounds - コンテンツの境界情報
     * @param {PreviewTemplates} templates - プレビュー元の複製図形
     * @param {{isUndo: boolean}} previewState - プレビューの undo 状態
     * @param {number} unitFactor - 入力欄の単位1つあたりの pt 数
     * @returns {{reflectEnabled: function, refresh: function, getFinalValues: function, commitFinal: function}} コントローラ
     */
    function createPreviewController(widgets, bounds, templates, previewState, unitFactor) {

        /**
         * チェック状態から各ウィジェットの有効／無効を反映する
         * @returns {void}
         */
        function reflectEnabled() {
            var isAdjustEnabled = widgets.chkAdjustEnabled.value;
            var isRadiusEnabled = widgets.chkRadiusEnabled.value;
            var isPill = widgets.chkPill.value;
            var isLinked = widgets.chkLink.value;

            /* ピル形状は幅を自動計算するので連動とは併用しない / Pill drives the width itself, so unlink */
            if (isPill && isLinked) {
                widgets.chkLink.value = false;
                isLinked = false;
            }

            widgets.inputW.enabled = isAdjustEnabled && !isPill;
            widgets.inputH.enabled = isAdjustEnabled && (isPill || !isLinked);
            widgets.inputR.enabled = isAdjustEnabled && isRadiusEnabled && !isPill;
            widgets.chkPill.enabled = isAdjustEnabled && isRadiusEnabled;
            widgets.chkLink.enabled = isAdjustEnabled && !isPill;
            widgets.chkAdjustEnabled.enabled = true;
            widgets.chkRadiusEnabled.enabled = isAdjustEnabled;
        }

        /**
         * ウィジェットの現在値を読み取る（数値は表示単位から pt に換算）
         * @returns {object} UI の入力値
         */
        function readUIValues() {
            return {
                adjustEnabled: widgets.chkAdjustEnabled.value,
                radiusEnabled: widgets.chkRadiusEnabled.value,
                pill: widgets.chkPill.value,
                addW: parseNumberOrDefault(widgets.inputW.text, 0) * unitFactor,
                addH: parseNumberOrDefault(widgets.inputH.text, 0) * unitFactor,
                radius: parseNumberOrDefault(widgets.inputR.text, 0) * unitFactor
            };
        }

        /**
         * 自動計算した値を入力欄に書き戻す
         * @param {object} previewValues - computePreviewValues() の戻り値
         * @returns {void}
         */
        function applyDerivedUIValues(previewValues) {
            if (!widgets.chkAdjustEnabled.value) return;
            if (!widgets.chkRadiusEnabled.value) {
                widgets.inputR.text = "0";
                return;
            }
            widgets.inputR.text = previewValues.radiusText;
            if (widgets.chkPill.value) {
                widgets.inputW.text = previewValues.widthText;
                widgets.inputH.text = previewValues.heightText;
            }
        }

        /**
         * 現在の UI 値からプレビュー図形をゼロから作り直す。
         * undo のタイミングは runPreview / undoPreview が管理するので、
         * ユーザーの操作履歴には最後の1回分だけが残る。
         * 確定後に特定できるよう、末尾でプレビュー図形を選択状態にする。
         * @returns {void}
         */
        function process() {
            var previewValues = computePreviewValues(readUIValues(), bounds, unitFactor);
            applyDerivedUIValues(previewValues);

            var doc = app.activeDocument;
            var useAdjustedPreview = widgets.chkAdjustEnabled.value || !templates.previewSourceShapeItem;
            var previewTemplate = useAdjustedPreview ? templates.previewBaseShapeItem : templates.previewSourceShapeItem;

            var previewItem = previewTemplate.duplicate();
            previewItem.hidden = false;

            if (useAdjustedPreview) {
                /* パディングは線の外側ではなくパス基準。線幅を変えても余白が変わらない /
                   Padding is measured from the path, not the stroke, so stroke weight does not shift it */
                var newWidth = Math.max(MIN_PREVIEW_SIZE, bounds.width + previewValues.addW);
                var newHeight = Math.max(MIN_PREVIEW_SIZE, bounds.height + previewValues.addH);
                resizeAndCenterByGeometry(
                    previewItem, newWidth, newHeight, bounds.centerX, bounds.centerY
                );

                if (previewValues.radius > 0) {
                    var roundCornersXml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + previewValues.radius + ' "/></LiveEffect>';
                    previewItem.applyEffect(roundCornersXml);
                }
            } else {
                /* 調整なしのときは元図形の寸法のまま、コンテンツの中心に合わせるだけ / Align only */
                centerByGeometry(previewItem, bounds.centerX, bounds.centerY);
            }

            /* 確定後に特定するため選択しておく / Select so the caller can identify it after OK */
            doc.selection = null;
            previewItem.selected = true;
        }

        /**
         * プレビューを更新する
         * @returns {void}
         */
        function refresh() {
            runPreview(previewState, process);
        }

        /**
         * 入力値を検証し、連動・ピル形状に合わせて入力欄を整える
         * @returns {boolean|null} 妥当なら true、不正なら null（警告表示済み）
         */
        function validateInputs() {
            if (!widgets.chkAdjustEnabled.value) return true;

            var addW = validateNumericField(widgets.inputW, "invalidNumber");
            if (addW === null) return null;

            var addH;
            if (widgets.chkLink.value) {
                addH = addW;
            } else {
                addH = validateNumericField(widgets.inputH, "invalidNumber");
                if (addH === null) return null;
            }
            widgets.inputH.text = String(addH);

            if (!widgets.chkRadiusEnabled.value) {
                widgets.inputR.text = "0";
            } else if (!widgets.chkPill.value) {
                var radius = validateNumericField(widgets.inputR, "invalidRadius");
                if (radius === null) return null;
                widgets.inputR.text = String(radius);
            }
            return true;
        }

        /**
         * 入力を検証し、確定処理に必要な値を返す
         * @returns {{shouldRunPathfinder: boolean}|null} 確定値、不正なら null（警告表示済み）
         */
        function getFinalValues() {
            if (validateInputs() === null) return null;
            return {
                /* ピル形状は角丸効果を実体化する必要がある / Pill shapes must flatten the round-corner effect */
                shouldRunPathfinder: widgets.chkAdjustEnabled.value &&
                    widgets.chkRadiusEnabled.value && widgets.chkPill.value
            };
        }

        /**
         * 前回プレビューを巻き戻し、本番として1回だけ生成する
         * @returns {void}
         */
        function commitFinal() {
            undoPreview(previewState);
            process();
        }

        return {
            reflectEnabled: reflectEnabled,
            refresh: refresh,
            getFinalValues: getFinalValues,
            commitFinal: commitFinal
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * @typedef {object} DialogContext
     * @property {ContentBounds} bounds - コンテンツの境界情報
     * @property {PageItem|null} previewSourceShapeItem - 整列専用の複製
     * @property {PageItem} previewBaseShapeItem - 調整用の複製
     * @property {boolean} shapeIsAutoCreated - 長方形を自動作成したかどうか
     */

    /**
     * ダイアログのウィジェットを組み立てる
     * @param {Window} targetWindow - 追加先のウィンドウ
     * @param {object} sessionState - 前回の設定
     * @param {{label: string, factor: number}} rulerUnit - 定規の単位
     * @param {boolean} shapeIsAutoCreated - 長方形を自動作成したかどうか
     * @returns {PreviewWidgets} 生成したウィジェット一式
     */
    function buildDialogWidgets(targetWindow, sessionState, rulerUnit, shapeIsAutoCreated) {
        var adjustRow = targetWindow.add("group");
        setupRow(adjustRow, "center");
        var chkAdjustEnabled = adjustRow.add("checkbox", undefined, getLabel("checkbox", "adjustEnabled"));
        chkAdjustEnabled.value = !!shapeIsAutoCreated;

        /* パディング / Padding */
        var paddingPanel = addPanel(targetWindow, getLabel("panel", "padding"));
        var paddingRow = paddingPanel.add("group");
        setupRow(paddingRow, "left", COLUMN_SPACING);

        var paddingFields = paddingRow.add("group");
        paddingFields.orientation = "column";
        paddingFields.alignChildren = ["left", "center"];
        var inputW = addNumericFieldRow(paddingFields, labelText("fieldLabel", "width"), sessionState.addW, rulerUnit.label);
        var inputH = addNumericFieldRow(paddingFields, labelText("fieldLabel", "height"), sessionState.addH, rulerUnit.label);

        var linkColumn = paddingRow.add("group");
        linkColumn.orientation = "column";
        linkColumn.alignChildren = ["left", "center"];
        var chkLink = linkColumn.add("checkbox", undefined, getLabel("checkbox", "link"));
        chkLink.value = sessionState.link;

        /* 角丸 / Rounded corners */
        var cornerPanel = addPanel(targetWindow, getLabel("panel", "corner"));
        var radiusRow = cornerPanel.add("group");
        setupRow(radiusRow);
        var chkRadiusEnabled = radiusRow.add("checkbox", undefined, labelText("fieldLabel", "radius"));
        chkRadiusEnabled.value = (sessionState.radiusEnabled !== false);
        var inputR = radiusRow.add("edittext", undefined, sessionState.radius);
        inputR.characters = FIELD_CHARS;
        radiusRow.add("statictext", undefined, rulerUnit.label);

        var pillRow = cornerPanel.add("group");
        setupRow(pillRow);
        var chkPill = pillRow.add("checkbox", undefined, getLabel("checkbox", "pill"));
        chkPill.value = !!sessionState.pill;

        return {
            chkAdjustEnabled: chkAdjustEnabled,
            chkRadiusEnabled: chkRadiusEnabled,
            chkPill: chkPill,
            chkLink: chkLink,
            inputW: inputW,
            inputH: inputH,
            inputR: inputR
        };
    }

    /**
     * 入力欄・チェックボックスにイベントを結び付ける
     * @param {PreviewWidgets} widgets - ダイアログのウィジェット
     * @param {{reflectEnabled: function, refresh: function}} ctrl - プレビューコントローラ
     * @returns {void}
     */
    function bindDialogEvents(widgets, ctrl) {

        /**
         * 連動がONなら高さを幅にそろえる
         * @returns {void}
         */
        function syncHeightToWidth() {
            if (widgets.chkLink.value) widgets.inputH.text = widgets.inputW.text;
        }

        /**
         * 連動がONなら幅を高さにそろえる
         * @returns {void}
         */
        function syncWidthToHeight() {
            if (widgets.chkLink.value) widgets.inputW.text = widgets.inputH.text;
        }

        /**
         * 入力欄の負値を 0 に丸める
         * @param {EditText} editText - 対象の入力欄
         * @returns {void}
         */
        function clampNonNegativeText(editText) {
            var value = parseFloat(editText.text);
            if (!isNaN(value) && value < 0) editText.text = "0";
        }

        /**
         * 幅を変えたあと、連動を反映してプレビューを更新する
         * @returns {void}
         */
        function syncAndPreviewW() {
            syncHeightToWidth();
            ctrl.refresh();
        }

        /**
         * 高さを変えたあと、連動を反映してプレビューを更新する
         * @returns {void}
         */
        function syncAndPreviewH() {
            syncWidthToHeight();
            ctrl.refresh();
        }

        /**
         * チェックボックス用のハンドラを作る（変更を適用してから有効状態とプレビューを更新）
         * @param {function|null} mutateFn - 先に適用する処理（不要なら null）
         * @returns {function} onClick に割り当てるハンドラ
         */
        function checkboxHandler(mutateFn) {
            return function () {
                if (mutateFn) mutateFn();
                ctrl.reflectEnabled();
                ctrl.refresh();
            };
        }

        widgets.inputW.onChanging = function () {
            clampNonNegativeText(widgets.inputW);
            syncAndPreviewW();
        };
        widgets.inputH.onChanging = function () {
            clampNonNegativeText(widgets.inputH);
            syncAndPreviewH();
        };
        widgets.inputR.onChanging = function () {
            clampNonNegativeText(widgets.inputR);
            ctrl.refresh();
        };

        changeValueByArrowKey(widgets.inputW, syncAndPreviewW, 0);
        changeValueByArrowKey(widgets.inputH, syncAndPreviewH, 0);
        changeValueByArrowKey(widgets.inputR, ctrl.refresh, 0);

        widgets.chkAdjustEnabled.onClick = checkboxHandler(null);
        widgets.chkLink.onClick = checkboxHandler(syncHeightToWidth);
        widgets.chkPill.onClick = checkboxHandler(function () {
            /* ピル解除で幅が手入力に戻るので、連動していれば高さを合わせ直す /
               Leaving pill mode hands the width back to the user; re-mirror it when linked */
            if (!widgets.chkPill.value) syncHeightToWidth();
        });
        widgets.chkRadiusEnabled.onClick = checkboxHandler(function () {
            if (!widgets.chkRadiusEnabled.value) {
                widgets.inputR.text = "0";
                widgets.chkPill.value = false;
            }
            syncHeightToWidth();
        });

        /* 初期状態：連動ONなら高さは幅に追従 / Initial state: height follows width when linked */
        syncHeightToWidth();
    }

    /**
     * 現在の UI 状態からセッション保存用の設定を作る
     * @param {PreviewWidgets} widgets - ダイアログのウィジェット
     * @returns {object} 保存する設定
     */
    function buildSessionPayload(widgets) {
        return {
            addW: widgets.inputW.text,
            addH: widgets.inputH.text,
            radius: widgets.inputR.text,
            radiusEnabled: widgets.chkRadiusEnabled.value,
            link: widgets.chkLink.value,
            pill: widgets.chkPill.value
        };
    }

    /**
     * 設定ダイアログを表示し、確定した内容を返す
     * @param {DialogContext} dialogContext - ダイアログに渡す情報
     * @returns {{shouldRunPathfinder: boolean, previewItem: PageItem}|null} 確定結果、キャンセル時は null
     */
    function showDialog(dialogContext) {
        var rulerUnit = getRulerUnit();
        var sessionState = getSessionState(rulerUnit.factor);
        var previewState = { isUndo: false };
        var confirmedValues = null;
        var finalPreviewItem = null;

        var win = new Window('dialog', getLabel('dialog', 'title') + ' ' + SCRIPT_VERSION);
        setupWindow(win);

        var widgets = buildDialogWidgets(win, sessionState, rulerUnit, dialogContext.shapeIsAutoCreated);
        var ctrl = createPreviewController(
            widgets,
            dialogContext.bounds,
            {
                previewSourceShapeItem: dialogContext.previewSourceShapeItem,
                previewBaseShapeItem: dialogContext.previewBaseShapeItem
            },
            previewState,
            rulerUnit.factor
        );
        bindDialogEvents(widgets, ctrl);

        var buttons = addButtonBar(win);

        buttons.btnOK.onClick = function () {
            confirmedValues = ctrl.getFinalValues();
            if (!confirmedValues) return;

            saveSessionState(buildSessionPayload(widgets));

            /* 前回プレビューを巻き戻して、本番として1回だけ確定 / Undo the preview, then commit once */
            ctrl.commitFinal();

            /* commitFinal はプレビュー図形を選択状態で残す / commitFinal leaves the item selected */
            var sel = app.activeDocument.selection;
            if (sel && sel.length > 0) finalPreviewItem = sel[0];

            win.close(1);
        };

        buttons.btnCancel.onClick = function () {
            saveSessionState(buildSessionPayload(widgets));
            /* 残ったプレビューは win.onClose で巻き戻す / onClose reverts the leftover preview */
            win.close(0);
        };

        /* OK / キャンセル / Esc / 閉じるボタン、どの経路でも残プレビューを片付ける /
           Catch-all: revert any leftover preview however the dialog closes */
        win.onClose = function () {
            undoPreview(previewState);
        };

        /* 初回プレビューは win.onShow（コールバック）から起動する。同期側で呼ぶと、
           後続コールバックの app.undo() が main のセットアップ（テンプレート作成等）まで巻き戻すおそれがある。
           Initial preview is fired from win.onShow — calling it synchronously risks a later
           app.undo() rolling back main()'s template setup. */
        win.onShow = function () {
            widgets.inputW.active = true;
            widgets.inputW.selection = [0, widgets.inputW.text.length];
            if (widgets.chkPill.value) widgets.chkLink.value = false;
            ctrl.reflectEnabled();
            ctrl.refresh();
        };

        if (win.show() !== 1 || !confirmedValues) {
            /* キャンセル時は onClose の undo で未適用に戻っている / onClose already reverted everything */
            safeRemove(finalPreviewItem);
            return null;
        }

        return {
            shouldRunPathfinder: confirmedValues.shouldRunPathfinder,
            previewItem: finalPreviewItem
        };
    }

    // =========================================
    // 選択の解釈 / Selection parsing
    // =========================================

    /**
     * 「コンテンツ1つ＋図形1つ」だけで構成されたグループから、その2つを取り出す。
     * グループは解除も削除もせず、条件に合わなければ null を返す。
     * @param {GroupItem} groupItem - 対象グループ
     * @returns {{contentItem: PageItem, shapeItem: PageItem}|null} 取り出した組、対象外なら null
     */
    function findContentAndShapeInGroup(groupItem) {
        if (!groupItem || groupItem.typename !== "GroupItem") return null;
        if (groupItem.clipped) return null;
        if (groupItem.pageItems.length !== 2) return null;

        var contentChild = null;
        var shapeChild = null;

        for (var i = 0; i < 2; i++) {
            var child = groupItem.pageItems[i];
            if (isContentItem(child) && !contentChild) {
                contentChild = child;
            } else if (isShapeItem(child) && !shapeChild) {
                shapeChild = child;
            } else {
                return null;
            }
        }

        if (!contentChild || !shapeChild) return null;

        return { contentItem: contentChild, shapeItem: shapeChild };
    }

    /**
     * 選択内容を検証してコンテンツと図形に振り分ける
     * @param {Array<PageItem>} sel - ドキュメントの選択
     * @returns {{contentItem: PageItem, shapeItem: PageItem|null, shapeIsAutoCreated: boolean}|null} 振り分け結果、不正なら null
     */
    function parseSelection(sel) {
        if (!sel || sel.length < 1 || sel.length > 2) {
            alertSelectionError("selectOne");
            return null;
        }

        var contentItem = null;
        var shapeItem = null;
        var shapeIsAutoCreated = false;

        if (sel.length === 2) {
            /* 2つ選択：テキスト/グループ＋図形 / Two items: text or group, plus a shape */
            for (var i = 0; i < sel.length; i++) {
                var item = sel[i];
                if (isContentItem(item) && !contentItem) {
                    contentItem = item;
                } else if (isShapeItem(item) && !shapeItem) {
                    shapeItem = item;
                } else {
                    alertSelectionError("selectOne");
                    return null;
                }
            }
        } else {
            /* 1つ選択：テキスト＋図形のグループなら、グループを保ったまま中身を使い分ける。
               それ以外はコンテンツとみなして長方形を自動作成する /
               One item: reuse the members of a text+shape group in place; otherwise auto-create a rectangle */
            var selectedItem = sel[0];
            var groupMembers = null;
            if (selectedItem && selectedItem.typename === "GroupItem" && !isClippingGroupItem(selectedItem)) {
                groupMembers = findContentAndShapeInGroup(selectedItem);
            }
            if (groupMembers) {
                contentItem = groupMembers.contentItem;
                shapeItem = groupMembers.shapeItem;
            } else {
                if (!isContentItem(selectedItem)) {
                    alertSelectionError("selectOne");
                    return null;
                }
                contentItem = selectedItem;
                shapeIsAutoCreated = true;
            }
        }

        if (isClippingGroupItem(contentItem)) {
            alertSelectionError("clippingGroup");
            return null;
        }

        return {
            contentItem: contentItem,
            shapeItem: shapeItem,
            shapeIsAutoCreated: shapeIsAutoCreated
        };
    }

    /**
     * コンテンツを複製して境界を計測する（テキストはアウトライン化して字面を測る）
     * @param {PageItem} contentItem - 計測対象
     * @returns {ContentBounds|null} 境界情報、失敗時は null（警告表示済み）
     */
    function measureContent(contentItem) {
        var measureItem = null;
        var dupText = null;
        var bounds = null;

        try {
            if (contentItem.typename === "TextFrame") {
                dupText = contentItem.duplicate();
                measureItem = dupText.createOutline();
            } else {
                measureItem = contentItem.duplicate();
            }
            bounds = getBoundsFromItem(measureItem);
        } catch (err) {
            bounds = null;
        } finally {
            safeRemove(measureItem);
            safeRemove(dupText);
        }

        if (!bounds) {
            alertSelectionError("measureFailed");
            return null;
        }
        return bounds;
    }

    // =========================================
    // 図形の準備と確定 / Shape preparation and commit
    // =========================================

    /**
     * 必要なら長方形を自動作成し、プレビュー用の複製を用意して元図形を隠す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} contentItem - コンテンツ
     * @param {PageItem|null} shapeItem - 選択された図形（自動作成時は null）
     * @param {boolean} shapeIsAutoCreated - 長方形を自動作成するかどうか
     * @param {ContentBounds} bounds - コンテンツの境界情報
     * @returns {{shapeItem: PageItem, previewSourceShapeItem: PageItem|null, previewBaseShapeItem: PageItem}} 準備結果
     */
    function prepareShapeAndPreviews(doc, contentItem, shapeItem, shapeIsAutoCreated, bounds) {
        var previewSourceShapeItem = null;
        var previewBaseShapeItem;

        if (shapeIsAutoCreated) {
            /* コンテンツと同じ大きさの長方形を作り、それをプレビューの元にする /
               Create a rectangle the size of the content and use it as the preview template */
            shapeItem = doc.pathItems.rectangle(bounds.top, bounds.left, bounds.width, bounds.height);
            shapeItem.opacity = AUTO_CREATED_SHAPE_OPACITY;
            shapeItem.move(contentItem, ElementPlacement.PLACEAFTER);
            previewBaseShapeItem = shapeItem.duplicate();
        } else {
            /* 既存図形：整列専用（アピアランスそのまま）と調整専用（アピアランス消去）の2本立て。
               アピアランスが載ったままだと拡大縮小で見た目が崩れる /
               Existing shape: one duplicate as-is for align-only, one cleared so resizing cannot distort it */
            previewSourceShapeItem = shapeItem.duplicate();
            previewSourceShapeItem.hidden = true;

            previewBaseShapeItem = shapeItem.duplicate();
            clearAppearanceByAction(previewBaseShapeItem);
        }

        previewBaseShapeItem.hidden = true;
        shapeItem.hidden = true;
        app.redraw();

        return {
            shapeItem: shapeItem,
            previewSourceShapeItem: previewSourceShapeItem,
            previewBaseShapeItem: previewBaseShapeItem
        };
    }

    /**
     * キャンセル時の後片付け（テンプレート削除、自動作成した長方形の削除、元図形の再表示）
     * @param {object} prepared - prepareShapeAndPreviews() の戻り値
     * @param {boolean} shapeIsAutoCreated - 長方形を自動作成したかどうか
     * @returns {void}
     */
    function cancelDialogResult(prepared, shapeIsAutoCreated) {
        safeRemove(prepared.previewSourceShapeItem);
        safeRemove(prepared.previewBaseShapeItem);
        if (shapeIsAutoCreated) {
            safeRemove(prepared.shapeItem);
        } else {
            safeDo(function () { prepared.shapeItem.hidden = false; });
        }
        app.redraw();
    }

    /**
     * OK 確定時の後処理（テンプレート削除、必要ならライブパスファインダー、元図形削除、最終選択）
     * @param {Document} doc - 対象ドキュメント
     * @param {object} result - showDialog() の戻り値
     * @param {object} prepared - prepareShapeAndPreviews() の戻り値
     * @param {PageItem} contentItem - コンテンツ
     * @param {boolean} shapeIsAutoCreated - 長方形を自動作成したかどうか
     * @returns {void}
     */
    function commitDialogResult(doc, result, prepared, contentItem, shapeIsAutoCreated) {
        /* プレビュー用テンプレートは複製元なので確定図形とは常に別物 / Templates are never the committed item */
        safeRemove(prepared.previewSourceShapeItem);
        safeRemove(prepared.previewBaseShapeItem);

        if (!result.previewItem) {
            /* 確定図形を特定できなかった：元の状態に戻す / Could not identify the item; roll back */
            cancelDialogResult(prepared, shapeIsAutoCreated);
            return;
        }

        var finalPreviewItem = result.previewItem;
        finalPreviewItem.hidden = false;

        if (result.shouldRunPathfinder) {
            /* ピル形状は角丸効果を実体化してから合成する / Flatten the round-corner effect for pill shapes */
            doc.selection = null;
            finalPreviewItem.selected = true;
            app.executeMenuCommand('Live Pathfinder Add');
            if (!doc.selection || doc.selection.length !== 1) {
                throw new Error('Live Pathfinder Add did not return exactly one selected item.');
            }
            finalPreviewItem = doc.selection[0];
            finalPreviewItem.hidden = false;
        }

        try {
            prepared.shapeItem.remove();
        } catch (removeError) {
            if (!shapeIsAutoCreated) prepared.shapeItem.hidden = false;
            throw removeError;
        }

        doc.selection = null;
        contentItem.selected = true;
        finalPreviewItem.selected = true;
        app.redraw();
    }

    // =========================================
    // メイン処理 / Main process
    // =========================================

    /**
     * スクリプトのエントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert", "noDocument"));
            return;
        }

        var doc = app.activeDocument;
        var parsed = parseSelection(doc.selection);
        if (!parsed) return;

        var bounds = measureContent(parsed.contentItem);
        if (!bounds) return;

        var prepared = prepareShapeAndPreviews(
            doc, parsed.contentItem, parsed.shapeItem, parsed.shapeIsAutoCreated, bounds
        );

        var result = showDialog({
            bounds: bounds,
            previewSourceShapeItem: prepared.previewSourceShapeItem,
            previewBaseShapeItem: prepared.previewBaseShapeItem,
            shapeIsAutoCreated: parsed.shapeIsAutoCreated
        });

        if (result) {
            commitDialogResult(doc, result, prepared, parsed.contentItem, parsed.shapeIsAutoCreated);
        } else {
            cancelDialogResult(prepared, parsed.shapeIsAutoCreated);
        }
    }

    main();

})();
