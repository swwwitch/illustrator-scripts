#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

#targetengine "MyScriptEngine"

/*

### 概要

選択したテキスト（またはオブジェクト）の背面に、見た目寸法に基づく図形を生成して配置します。
既存の背面図形があれば検出して置き換え、プレビューは［OK］時に1ステップで確定します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBackdrop.md

note記事も参照してください。
https://note.com/dtp_tranist/n/na8af4a7016ad

### Overview

Generates a shape sized to the visual bounds of the selected text (or objects) and places it behind them.
An existing backdrop is detected and replaced, and the undo-based preview commits in one step on OK.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddBackdrop.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AddBackdrop";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.6.4";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-12-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-23";                   /* 更新日 / last updated */

var SCRIPT_README_JA   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AddBackdrop.md"; /* README（日本語） */
var SCRIPT_README_EN   = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AddBackdrop.md"; /* README (English) */
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/na8af4a7016ad"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var SUPERELLIPSE_EXPONENT     = 2.5;   /* スーパー楕円の指数（2で円、大きいほど四角に近づく） / Superellipse exponent */
    var SUPERELLIPSE_POINT_COUNT  = 8;     /* スーパー楕円のアンカー数（既存図形の判定にも使う） / Anchor count, also used for detection */
    var SUPERELLIPSE_HANDLE_RATIO = 0.35;  /* ハンドル長の係数（大きいほど丸い） / Handle length ratio */
    var ONE_CHAR_DIAMETER_RATIO   = 1.5;   /* 1文字モードの直径（文字サイズに対する倍率） / Diameter per font size in single-character mode */
    var DEFAULT_MARGIN_DIVISOR    = 4;     /* マージンの初期値＝短辺÷この値 / Default margin = short side / this */
    var DEFAULT_ROUND_DIVISOR     = 5;     /* 角丸の初期値＝短辺÷この値 / Default corner radius = short side / this */

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OFFSET_X         = 300;               /* 表示時の横方向のずらし量 / Horizontal shift on show */
    var DIALOG_OFFSET_Y         = 0;                 /* 表示時の縦方向のずらし量 / Vertical shift on show */
    var DIALOG_OPACITY          = 0.98;              /* ダイアログの不透明度 / Dialog opacity */
    var PANEL_MARGINS           = [15, 20, 15, 10];  /* パネル余白 [左,上,右,下] / Panel margins [left, top, right, bottom] */
    var FIELD_CHARS             = 4;                 /* 数値欄の文字数 / Characters of numeric fields */
    var SHAPE_ROW_BOTTOM_MARGIN = 5;                 /* 形状ラジオの下の余白 / Space below the shape radios */
    var SHORT_FIELD_CHARS       = 3;                 /* 短い数値欄（倍率・マージン・CMYK）の文字数 / Characters of short fields */

    /**
     * 縦並びのパネルを追加する
     * @param {Group} parentContainer - 追加先
     * @param {string} panelTitle - パネルの見出し
     * @param {string[]} [childAlignment] - alignChildren（省略時は ["fill", "top"]）
     * @returns {Panel} 追加したパネル
     */
    function addPanel(parentContainer, panelTitle, childAlignment) {
        var createdPanel = parentContainer.add('panel', undefined, panelTitle);
        createdPanel.orientation = 'column';
        createdPanel.alignChildren = childAlignment || ['fill', 'top'];
        createdPanel.margins = PANEL_MARGINS;
        return createdPanel;
    }

    /**
     * 縦並びのグループを追加する
     * @param {Group} parentContainer - 追加先
     * @param {string[]} childAlignment - alignChildren
     * @returns {Group} 追加したグループ
     */
    function addColumnGroup(parentContainer, childAlignment) {
        var createdGroup = parentContainer.add('group');
        createdGroup.orientation = 'column';
        createdGroup.alignChildren = childAlignment;
        return createdGroup;
    }

    /**
     * 数値欄を追加する
     * @param {Group} parentContainer - 追加先
     * @param {string} initialText - 初期値
     * @param {string} tooltipText - ツールチップ
     * @returns {EditText} 追加した数値欄
     */
    function addNumberField(parentContainer, initialText, tooltipText) {
        var field = parentContainer.add('edittext', undefined, initialText);
        field.helpTip = tooltipText;
        field.characters = FIELD_CHARS;
        return field;
    }

    // =========================================
    // セッション記憶 / Session memory
    // =========================================
    /* #targetengine 下の $.global は Illustrator の起動中は残るので、前回の UI 値をここに控える
       $.global survives while Illustrator runs under #targetengine; the last UI values live here */
    if (!$.global.__AddBackdrop_State) {
        $.global.__AddBackdrop_State = {
            _hasSaved: false,
            shape: 'circle',          /* 'circle' | 'superellipse' | 'rect' */
            scale: '90',
            oneChar: false,
            marginV: '0',
            marginH: '0',
            marginLink: true,
            marginSquare: false,
            roundEnable: false,
            roundValue: '2',
            pill: false,
            groupWithText: true,
            exclude: false,
            offsetX: '0',
            offsetY: '0',
            kind: 'fill',             /* 'fill' | 'stroke' */
            strokeWidth: '1',
            colorMode: 'black',       /* 'text' | 'black' | 'white' | 'cmyk' */
            cmykC: '0',
            cmykM: '0',
            cmykY: '0',
            cmykK: '0',
            opacityApply: false,
            opacity: '60'
        };
    }
    var dialogState = $.global.__AddBackdrop_State;

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
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

    /* 単位コード5を「歯（H）」と表示する環境設定キー。文字サイズ（text/units）だけ「級（Q）」
       Preference keys that show unit code 5 as H; only the type size (text/units) shows Q */
    var HA_UNIT_PREF_KEYS = { "rulerType": true, "strokeUnits": true, "text/asianunits": true };

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitKey = prefKey || "rulerType";
        var unitCode = app.preferences.getIntegerPreference(unitKey);
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        /* 級（Q）と歯（H）は同じ長さだが、文字サイズは「Q」、距離は「H」と呼び分ける */
        var label = (unitCode === 5 && HA_UNIT_PREF_KEYS[unitKey]) ? "H" : unit.label;
        return { code: unitCode, label: label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * UI の言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "背面に図形を敷く", en: "Add Backdrop" }
        },
        panel: {
            scale:    { ja: "スケール", en: "Scale" },
            margin:   { ja: "マージン", en: "Margin" },
            round:    { ja: "角丸", en: "Round" },
            grouping: { ja: "グループ", en: "Group" },
            offset:   { ja: "位置の調整", en: "Offset" },
            kind:     { ja: "塗りと線", en: "Fill & Stroke" },
            color:    { ja: "カラー", en: "Color" },
            opacity:  { ja: "不透明度", en: "Opacity" }
        },
        fieldLabel: {
            scale:   { ja: "倍率", en: "Size" },
            marginV: { ja: "上下", en: "Vertical" },
            marginH: { ja: "左右", en: "Horizontal" },
            offsetX: { ja: "X", en: "X" },
            offsetY: { ja: "Y", en: "Y" }
        },
        radio: {
            circle:       { ja: "正円", en: "Circle" },
            superellipse: { ja: "スーパー楕円", en: "Superellipse" },
            rectangle:    { ja: "長方形", en: "Rectangle" },
            fill:         { ja: "塗り", en: "Fill" },
            stroke:       { ja: "線", en: "Stroke" },
            textColor:    { ja: "文字の色", en: "Text Color" },
            black:        { ja: "ブラック", en: "Black" },
            white:        { ja: "ホワイト", en: "White" },
            cmyk:         { ja: "CMYK", en: "CMYK" }
        },
        checkbox: {
            oneChar:       { ja: "1文字", en: "Single Character" },
            marginLink:    { ja: "連動", en: "Link" },
            square:        { ja: "正方形", en: "Square" },
            pill:          { ja: "ピル形状", en: "Pill shape" },
            groupWithText: { ja: "テキストとグループ化", en: "Group with Text" },
            exclude:       { ja: "中マド処理", en: "Exclude" }
        },
        tooltip: {
            circle:        { ja: "背景を正円にします。（E キー）", en: "Makes the backdrop a circle. (E key)" },
            superellipse: {
                ja: "背景をスーパー楕円（角の丸い四角に近い形）にします。（S キー）",
                en: "Makes the backdrop a superellipse, between a circle and a rounded square. (S key)"
            },
            rectangle:     { ja: "背景を長方形にします。（R キー）", en: "Makes the backdrop a rectangle. (R key)" },
            scale:         { ja: "文字に対する背景の大きさ（％）です。", en: "Size of the backdrop relative to the text, in percent." },
            oneChar: {
                ja: "文字サイズを基準に、1文字分の円を作ります。",
                en: "Sizes the circle from the font size to fit a single character."
            },
            marginV:       { ja: "文字の上下に足す余白です。", en: "Space added above and below the text." },
            marginH:       { ja: "文字の左右に足す余白です。", en: "Space added left and right of the text." },
            marginLink: {
                ja: "上下と左右の余白を同じ値にそろえます。",
                en: "Uses the same value for the vertical and horizontal margins."
            },
            square:        { ja: "背景を正方形にします。", en: "Makes the backdrop a square." },
            round: {
                ja: "背景の角を丸めます。半径は右の欄で指定します。",
                en: "Rounds the corners of the backdrop. The field on the right sets the radius."
            },
            pill:          { ja: "左右の端を半円にして、丸いピル型にします。", en: "Rounds both ends into a pill shape." },
            groupWithText: { ja: "背景と文字を1つのグループにまとめます。", en: "Groups the backdrop with the text." },
            exclude:       { ja: "背景と文字を「中マド」にして、文字部分を抜きます。", en: "Knocks the text out of the backdrop using Exclude." },
            offset:        { ja: "背景の位置を、この値だけずらします。", en: "Nudges the backdrop by this amount." },
            fill:          { ja: "背景に塗りを付けます。", en: "Fills the backdrop." },
            stroke: {
                ja: "背景に線を付けます。太さは右の欄で指定します。",
                en: "Strokes the backdrop. The field on the right sets the weight."
            },
            opacity:       { ja: "背景の不透明度（％）です。", en: "Opacity of the backdrop, in percent." },
            textColor:     { ja: "文字の色をそのまま背景の色に使います。", en: "Uses the text color for the backdrop." },
            black:         { ja: "背景を黒にします。", en: "Makes the backdrop black." },
            white:         { ja: "背景を白にします。", en: "Makes the backdrop white." },
            cmyk:          { ja: "背景の色をCMYKで指定します。", en: "Sets the backdrop color in CMYK." },
            cmykValue:     { ja: "この版の濃度（％）です。", en: "Ink percentage for this plate." },
            transparencyGrid: {
                ja: "透明グリッドの表示／非表示を切り替えます。白い背景の見え方を確かめるときに使います。",
                en: "Toggles the transparency grid, which helps check a white backdrop."
            }
        },
        button: {
            transparencyGrid: { ja: "透明グリッドを表示", en: "Show Transparency Grid" },
            ok:               { ja: "OK", en: "OK" },
            cancel:           { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument:    { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:   { ja: "オブジェクトを選択してください。", en: "Please select an object." },
            genericError:  { ja: "エラーが発生しました", en: "An error occurred" },
            previewError:  { ja: "プレビュー中にエラーが発生しました", en: "An error occurred during preview" },
            groupFailed:   { ja: "グループ化に失敗しました", en: "Failed to group objects" },
            excludeFailed: { ja: "中マド処理（Exclude）の適用に失敗しました", en: "Failed to apply Pathfinder Exclude" }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('radio', 'circle')）
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
     * 項目名にコロンを付ける（日本語は全角、英語は半角）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} コロン付きのラベル
     */
    function labelText() {
        return getLabel.apply(null, arguments) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 見出しに単位を添える（日本語は全角かっこ、英語は半角かっこ）
     * @param {string} titleText - 見出し
     * @param {string} unitLabel - 単位
     * @returns {string} 単位付きの見出し
     */
    function withUnit(titleText, unitLabel) {
        return (uiLang === "ja") ? titleText + "（" + unitLabel + "）" : titleText + " (" + unitLabel + ")";
    }

    // =========================================
    // 共通処理 / Helpers
    // =========================================

    /**
     * デバッグ用のログを ExtendScript Toolkit のコンソールに出す
     * @param {string} message - ログの内容
     * @returns {void}
     */
    function logDebug(message) {
        $.writeln('[' + SCRIPT_NAME + '] ' + message);
    }

    /**
     * オブジェクトを削除する（失敗してもログを出して続行）
     * @param {PageItem} item - 削除するオブジェクト
     * @param {string} itemLabel - ログに出す名前
     * @returns {void}
     */
    function safeRemove(item, itemLabel) {
        if (!item) return;
        try {
            item.remove();
        } catch (e) {
            logDebug('remove failed (' + itemLabel + '): ' + e);
        }
    }

    /**
     * オブジェクトだけを選択してメニューコマンドを実行し、結果のオブジェクトを返す
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} item - 対象オブジェクト
     * @param {string} menuCommand - executeMenuCommand に渡すコマンド名
     * @returns {PageItem} 実行後に選択されているオブジェクト（取れなければ元のオブジェクト）
     */
    function runMenuCommandOnItem(doc, item, menuCommand) {
        doc.selection = null;
        item.selected = true;
        app.executeMenuCommand(menuCommand);
        var resultItem = (doc.selection && doc.selection.length) ? doc.selection[0] : item;
        doc.selection = null;
        return resultItem;
    }

    /**
     * グループ内のオブジェクトをグループの直前（前面側）へ出す
     * @param {GroupItem} group - 解除するグループ
     * @param {PageItem} keepItem - グループに残すオブジェクト
     * @returns {void}
     */
    function releaseGroupChildren(group, keepItem) {
        for (var i = group.pageItems.length - 1; i >= 0; i--) {
            var child = group.pageItems[i];
            if (child === keepItem) continue;
            child.move(group, ElementPlacement.PLACEBEFORE);
        }
    }

    /**
     * 小数第1位に丸める
     * @param {number} value - 元の値
     * @returns {number} 丸めた値
     */
    function roundToTenth(value) {
        return Math.round(value * 10) / 10;
    }

    /**
     * グレーのカラーを作る
     * @param {number} grayValue - 濃度（0＝白、100＝黒）
     * @returns {GrayColor} カラー
     */
    function makeGrayColor(grayValue) {
        var color = new GrayColor();
        color.gray = grayValue;
        return color;
    }

    /**
     * 境界 [左, 上, 右, 下] を寸法付きの矩形情報にする
     * @param {number[]} bounds - [左, 上, 右, 下]
     * @returns {{left: number, top: number, right: number, bottom: number, w: number, h: number}} 矩形情報
     */
    function boundsToRect(bounds) {
        return {
            left: bounds[0],
            top: bounds[1],
            right: bounds[2],
            bottom: bounds[3],
            w: bounds[2] - bounds[0],
            h: bounds[1] - bounds[3]
        };
    }

    /**
     * オブジェクトの境界を取得する（visibleBounds を優先し、取れなければ geometricBounds）
     * @param {PageItem} item - 対象オブジェクト
     * @returns {number[]|null} [左, 上, 右, 下]
     */
    function getItemBounds(item) {
        var bounds = null;
        /* 種類によっては visibleBounds を持たないことがある / Some items may lack visibleBounds */
        try {
            bounds = item.visibleBounds;
        } catch (eVisible) {
            logDebug('visibleBounds read failed: ' + eVisible);
        }
        if (!bounds) {
            try {
                bounds = item.geometricBounds;
            } catch (eGeometric) {
                logDebug('geometricBounds read failed: ' + eGeometric);
            }
        }
        return bounds;
    }

    // =========================================
    // 数値欄 / Numeric fields
    // =========================================

    /**
     * 数値欄に許容範囲を設定する（入力時と↑↓キーの両方で使う）
     * @param {EditText} editText - 対象の数値欄
     * @param {number|null} minValue - 最小値（制限しないときは null）
     * @param {number|null} maxValue - 最大値（制限しないときは null）
     * @returns {void}
     */
    function setFieldRange(editText, minValue, maxValue) {
        editText.rangeMin = minValue;
        editText.rangeMax = maxValue;
    }

    /**
     * 数値を数値欄の許容範囲に収める
     * @param {EditText} editText - 対象の数値欄
     * @param {number} value - 元の値
     * @returns {number} 範囲内に収めた値
     */
    function clampToFieldRange(editText, value) {
        if (typeof editText.rangeMin === 'number' && value < editText.rangeMin) return editText.rangeMin;
        if (typeof editText.rangeMax === 'number' && value > editText.rangeMax) return editText.rangeMax;
        return value;
    }

    /**
     * ↑↓キーで数値欄の値を増減する（shift で10刻み、option で0.1刻み）
     * @param {EditText} editText - 対象の数値欄
     * @param {Function} onUpdate - 値を変えたあとに呼ぶ処理
     * @returns {void}
     */
    function changeValueByArrowKey(editText, onUpdate) {
        editText.addEventListener('keydown', function (event) {
            if (event.keyName != 'Up' && event.keyName != 'Down') return;

            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var isUp = (event.keyName === 'Up');
            if (keyboard.shiftKey) {
                /* 10刻み（10の倍数にスナップ）/ Step by 10, snapping to multiples of 10 */
                value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
            } else if (keyboard.altKey) {
                value = roundToTenth(value + (isUp ? 0.1 : -0.1));
            } else {
                value = Math.round(value + (isUp ? 1 : -1));
            }

            editText.text = clampToFieldRange(editText, value);
            onUpdate();
            if (event.preventDefault) event.preventDefault();
        });
    }

    /**
     * 数値欄の入力と↑↓キーに更新処理をつなぐ（入力中の値も許容範囲に収める）
     * @param {EditText} editText - 対象の数値欄
     * @param {Function} onUpdate - 値が変わったときに呼ぶ処理
     * @returns {void}
     */
    function bindNumberField(editText, onUpdate) {
        editText.onChanging = function () {
            var value = parseFloat(editText.text);
            if (!isNaN(value) && clampToFieldRange(editText, value) !== value) {
                editText.text = String(clampToFieldRange(editText, value));
            }
            onUpdate();
        };
        changeValueByArrowKey(editText, onUpdate);
    }

    // =========================================
    // ラジオボタン / Radio buttons
    // =========================================

    /**
     * ラジオボタンの対応表から、選択中のキーを返す
     * @param {Object} radioMap - キー → RadioButton
     * @param {string} fallbackKey - どれも選ばれていないときのキー
     * @returns {string} 選択中のキー
     */
    function getSelectedRadioKey(radioMap, fallbackKey) {
        for (var key in radioMap) {
            if (radioMap[key].value) return key;
        }
        return fallbackKey;
    }

    /**
     * ラジオボタンの対応表で、指定したキーのボタンだけを選ぶ
     * @param {Object} radioMap - キー → RadioButton
     * @param {string} selectedKey - 選ぶキー
     * @param {string} fallbackKey - selectedKey が対応表に無いときに選ぶキー
     * @returns {void}
     */
    function selectRadioKey(radioMap, selectedKey, fallbackKey) {
        if (!radioMap[selectedKey]) selectedKey = fallbackKey;
        for (var key in radioMap) {
            radioMap[key].value = (key === selectedKey);
        }
    }

    // =========================================
    // 形状 / Shape geometry
    // =========================================

    /**
     * 符号を返す
     * @param {number} value - 対象の値
     * @returns {number} 1 / -1 / 0
     */
    function signOf(value) {
        return ((value > 0) - (value < 0)) || +value;
    }

    /**
     * スーパー楕円のアンカーポイントを、角度で等分して求める
     * @param {number} cx - 中心X
     * @param {number} cy - 中心Y
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @param {number} exponent - 指数
     * @param {number} pointCount - アンカー数
     * @returns {number[][]} アンカー座標の配列
     */
    function buildSuperellipseAnchorPoints(cx, cy, width, height, exponent, pointCount) {
        var anchors = [];
        for (var i = 0; i < pointCount; i++) {
            var theta = (Math.PI * 2 * i) / pointCount;
            var cosTheta = Math.cos(theta);
            var sinTheta = Math.sin(theta);
            var x = Math.pow(Math.abs(cosTheta), 2 / exponent) * (width / 2) * signOf(cosTheta);
            var y = Math.pow(Math.abs(sinTheta), 2 / exponent) * (height / 2) * signOf(sinTheta);
            anchors.push([cx + x, cy + y]);
        }
        return anchors;
    }

    /**
     * パスをスーパー楕円に置き換え、各アンカーにスムーズなハンドルを付ける
     * @param {PathItem} pathItem - 対象のパス
     * @param {number} cx - 中心X
     * @param {number} cy - 中心Y
     * @param {number} width - 幅
     * @param {number} height - 高さ
     * @returns {void}
     */
    function morphPathToSuperellipse(pathItem, cx, cy, width, height) {
        var anchors = buildSuperellipseAnchorPoints(cx, cy, width, height, SUPERELLIPSE_EXPONENT, SUPERELLIPSE_POINT_COUNT);
        pathItem.setEntirePath(anchors);
        pathItem.closed = true;

        var pathPoints = pathItem.pathPoints;
        var pointCount = anchors.length;
        for (var i = 0; i < pointCount; i++) {
            var prevAnchor = anchors[(i - 1 + pointCount) % pointCount];
            var anchor = anchors[i];
            var nextAnchor = anchors[(i + 1) % pointCount];

            /* 前後のアンカーを結ぶ向きを接線にする / Tangent from the previous to the next anchor */
            var tangentX = nextAnchor[0] - prevAnchor[0];
            var tangentY = nextAnchor[1] - prevAnchor[1];
            var tangentLength = Math.sqrt(tangentX * tangentX + tangentY * tangentY);
            if (tangentLength === 0) continue;
            tangentX /= tangentLength;
            tangentY /= tangentLength;

            /* 短いほうの辺に合わせてハンドル長を決める / Handle length from the shorter neighboring segment */
            var prevLength = Math.sqrt(Math.pow(anchor[0] - prevAnchor[0], 2) + Math.pow(anchor[1] - prevAnchor[1], 2));
            var nextLength = Math.sqrt(Math.pow(nextAnchor[0] - anchor[0], 2) + Math.pow(nextAnchor[1] - anchor[1], 2));
            var handleLength = Math.min(prevLength, nextLength) * SUPERELLIPSE_HANDLE_RATIO;

            pathPoints[i].anchor = anchor;
            pathPoints[i].leftDirection = [anchor[0] - tangentX * handleLength, anchor[1] - tangentY * handleLength];
            pathPoints[i].rightDirection = [anchor[0] + tangentX * handleLength, anchor[1] + tangentY * handleLength];
            pathPoints[i].pointType = PointType.SMOOTH;
        }
    }

    /**
     * ライブ効果「角を丸くする」を適用する
     * @param {PageItem} item - 対象オブジェクト
     * @param {number} radius - 半径（pt）
     * @returns {void}
     */
    function applyRoundCornersLive(item, radius) {
        if (!radius || radius <= 0) return;
        try {
            item.applyEffect('<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radius + ' "/></LiveEffect>');
        } catch (e) {
            logDebug('applyRoundCornersLive failed: ' + e);
        }
    }

    // =========================================
    // 対象の特定 / Target resolution
    // =========================================

    /**
     * オブジェクトの並びから最初のテキストを探す（グループの中も辿る）
     * @param {PageItem[]|PageItems} items - 探す対象
     * @returns {TextFrame|null} 見つかったテキスト
     */
    function findFirstTextFrame(items) {
        for (var i = 0; i < items.length; i++) {
            var item = items[i];
            if (item.typename === 'TextFrame') return item;
            if (item.typename === 'GroupItem') {
                var foundText = findFirstTextFrame(item.pageItems);
                if (foundText) return foundText;
            }
        }
        return null;
    }

    /**
     * 選択から、置き換え対象の既存の背面図形を探す
     * - 単一グループ（テキスト＋図形）を選択しているとき：グループ内のテキスト以外の最初のオブジェクト
     * - 複数選択（テキストまたはグループ＋図形）のとき：最初のパス／複合パス
     * @param {PageItem[]} selectionItems - 選択オブジェクト
     * @param {TextFrame|null} textItem - 対象テキスト
     * @returns {{backdrop: PageItem, enclosingGroup: GroupItem, target: PageItem}} 見つからなければ各値は null
     */
    function findExistingBackdrop(selectionItems, textItem) {
        var found = { backdrop: null, enclosingGroup: null, target: null };
        var isTextMode = !!textItem;

        if (selectionItems.length === 1 && selectionItems[0].typename === 'GroupItem' && isTextMode) {
            var outerGroup = selectionItems[0];
            for (var i = 0; i < outerGroup.pageItems.length; i++) {
                var groupChild = outerGroup.pageItems[i];
                if (groupChild === textItem || groupChild.typename === 'TextFrame') continue;
                found.backdrop = groupChild;
                found.enclosingGroup = outerGroup;
                return found;
            }
        }

        if (selectionItems.length >= 2) {
            var otherTargets = [];
            var shapeCandidates = [];
            for (var j = 0; j < selectionItems.length; j++) {
                var selectedItem = selectionItems[j];
                if (selectedItem === textItem) continue;
                if (selectedItem.typename === 'TextFrame' || selectedItem.typename === 'GroupItem') {
                    otherTargets.push(selectedItem);
                } else if (selectedItem.typename === 'PathItem' || selectedItem.typename === 'CompoundPathItem') {
                    shapeCandidates.push(selectedItem);
                }
            }
            if (shapeCandidates.length > 0 && (isTextMode || otherTargets.length > 0)) {
                found.backdrop = shapeCandidates[0];
                if (!isTextMode) found.target = otherTargets[0];
            }
        }
        return found;
    }

    /**
     * グループ内から形状判定に使う最初のパスを探す
     * @param {GroupItem} group - 対象グループ
     * @returns {PathItem|null} 見つかったパス
     */
    function findFirstPathInGroup(group) {
        for (var i = 0; i < group.pageItems.length; i++) {
            var child = group.pageItems[i];
            if (child.typename === 'PathItem') return child;
            if (child.typename === 'CompoundPathItem' && child.pathItems.length > 0) return child.pathItems[0];
        }
        return null;
    }

    /**
     * 既存の図形から形状の種類を判定する
     * @param {PageItem} item - 対象オブジェクト
     * @returns {string} 'circle' | 'superellipse' | 'rect'
     */
    function detectShapeType(item) {
        try {
            var path = item;
            if (item.typename === 'CompoundPathItem') path = item.pathItems[0];
            if (item.typename === 'GroupItem') {
                path = findFirstPathInGroup(item);
                if (!path) {
                    /* パスが無ければ縦横比で判定 / Without a path, judge by the aspect ratio */
                    var groupRect = boundsToRect(item.geometricBounds);
                    return (Math.abs(groupRect.w - groupRect.h) < Math.max(groupRect.w, groupRect.h) * 0.05) ? 'circle' : 'rect';
                }
            }
            if (!path || path.typename !== 'PathItem') return 'rect';

            var pathPoints = path.pathPoints;
            if (pathPoints.length === SUPERELLIPSE_POINT_COUNT) return 'superellipse';
            if (pathPoints.length !== 4) return 'rect';

            /* 4点でハンドルがあれば楕円系（正円扱い）、無ければ長方形 / 4 points with handles = ellipse, else rectangle */
            for (var i = 0; i < 4; i++) {
                var anchor = pathPoints[i].anchor;
                var leftHandle = pathPoints[i].leftDirection;
                var rightHandle = pathPoints[i].rightDirection;
                var leftDistance = Math.abs(anchor[0] - leftHandle[0]) + Math.abs(anchor[1] - leftHandle[1]);
                var rightDistance = Math.abs(anchor[0] - rightHandle[0]) + Math.abs(anchor[1] - rightHandle[1]);
                if (leftDistance > 0.5 || rightDistance > 0.5) return 'circle';
            }
            return 'rect';
        } catch (e) {
            return 'rect';
        }
    }

    // =========================================
    // プレビュー / Preview
    // =========================================

    /**
     * Undo でプレビューを巻き戻す管理役
     * - プレビューの更新ごとに Undo 段数を数え、巻き戻しで消す
     * - OK 時は巻き戻してから確定処理を1回だけ実行する（Undo 1段）
     * @constructor
     */
    function PreviewManager() {
        this.undoDepth = 0;

        /**
         * プレビューを1段分実行する
         * @param {Function} previewAction - プレビューの処理
         * @returns {void}
         */
        this.addStep = function (previewAction) {
            try {
                previewAction();
                this.undoDepth++;
                app.redraw();
            } catch (e) {
                alert(getLabel('alert', 'previewError') + ": " + e);
            }
        };

        /**
         * これまでのプレビューをすべて巻き戻す
         * @returns {void}
         */
        this.rollback = function () {
            while (this.undoDepth > 0) {
                try {
                    app.undo();
                } catch (e) {
                    logDebug('rollback undo failed: ' + e);
                    break;
                }
                this.undoDepth--;
            }
            app.redraw();
        };

        /**
         * プレビューを巻き戻してから確定処理を実行する
         * @param {Function} finalAction - 確定処理
         * @returns {void}
         */
        this.confirm = function (finalAction) {
            this.rollback();
            finalAction();
            this.undoDepth = 0;
        };
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 形状のラジオボタンを追加する（ダイアログ上部）
     * @param {Window} dialog - ダイアログ
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addShapeRadios(dialog, ui) {
        var shapeRadioGroup = dialog.add('group');
        shapeRadioGroup.orientation = 'row';
        shapeRadioGroup.alignChildren = ['center', 'center'];
        shapeRadioGroup.alignment = 'center';
        shapeRadioGroup.margins = [0, 0, 0, SHAPE_ROW_BOTTOM_MARGIN];

        ui.rbCircle = shapeRadioGroup.add('radiobutton', undefined, getLabel('radio', 'circle'));
        ui.rbCircle.helpTip = getLabel('tooltip', 'circle');
        ui.rbSuperellipse = shapeRadioGroup.add('radiobutton', undefined, getLabel('radio', 'superellipse'));
        ui.rbSuperellipse.helpTip = getLabel('tooltip', 'superellipse');
        ui.rbRectangle = shapeRadioGroup.add('radiobutton', undefined, getLabel('radio', 'rectangle'));
        ui.rbRectangle.helpTip = getLabel('tooltip', 'rectangle');
        ui.rbCircle.value = true;

        ui.shapeRadios = { circle: ui.rbCircle, superellipse: ui.rbSuperellipse, rect: ui.rbRectangle };
    }

    /**
     * スケールパネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addScalePanel(column, ui) {
        ui.scalePanel = addPanel(column, getLabel('panel', 'scale'));

        var scaleRow = ui.scalePanel.add('group');
        scaleRow.add('statictext', undefined, labelText('fieldLabel', 'scale'));
        ui.scaleInput = addNumberField(scaleRow, '90', getLabel('tooltip', 'scale'));
        ui.scaleInput.characters = SHORT_FIELD_CHARS;
        scaleRow.add('statictext', undefined, '%');

        var oneCharRow = ui.scalePanel.add('group');
        oneCharRow.orientation = 'row';
        oneCharRow.alignChildren = ['left', 'center'];
        ui.cbOneChar = oneCharRow.add('checkbox', undefined, getLabel('checkbox', 'oneChar'));
        ui.cbOneChar.helpTip = getLabel('tooltip', 'oneChar');
    }

    /**
     * マージンの1行（ラベル・数値欄・単位）を追加する
     * @param {Group} parentContainer - 追加先
     * @param {string} labelKey - fieldLabel / tooltip のキー
     * @param {string} unitLabel - 単位
     * @returns {EditText} 追加した数値欄
     */
    function addMarginRow(parentContainer, labelKey, unitLabel) {
        var marginRow = parentContainer.add('group');
        marginRow.orientation = 'row';
        marginRow.alignChildren = ['left', 'center'];
        marginRow.spacing = 10;
        marginRow.margins = 0;
        marginRow.add('statictext', undefined, labelText('fieldLabel', labelKey));
        var marginInput = addNumberField(marginRow, '0', getLabel('tooltip', labelKey));
        marginInput.characters = SHORT_FIELD_CHARS;
        marginRow.add('statictext', undefined, unitLabel);
        return marginInput;
    }

    /**
     * マージンパネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @param {string} rulerUnitLabel - 定規の単位
     * @returns {void}
     */
    function addMarginPanel(column, ui, rulerUnitLabel) {
        ui.marginPanel = addPanel(column, getLabel('panel', 'margin'));

        var marginBodyRow = ui.marginPanel.add('group');
        marginBodyRow.orientation = 'row';
        marginBodyRow.alignChildren = ['left', 'top'];
        marginBodyRow.spacing = 10;
        marginBodyRow.margins = 0;

        var marginFieldsColumn = addColumnGroup(marginBodyRow, ['left', 'center']);
        marginFieldsColumn.spacing = 10;
        marginFieldsColumn.margins = 0;
        ui.marginVInput = addMarginRow(marginFieldsColumn, 'marginV', rulerUnitLabel);
        ui.marginHInput = addMarginRow(marginFieldsColumn, 'marginH', rulerUnitLabel);

        var marginLinkColumn = addColumnGroup(marginBodyRow, ['left', 'center']);
        marginLinkColumn.spacing = 0;
        marginLinkColumn.margins = 0;
        /* 2行の間にチェックボックスを置くためのスペーサー / Spacer to sit the checkbox between the two rows */
        marginLinkColumn.add('statictext', undefined, '');
        ui.cbMarginLink = marginLinkColumn.add('checkbox', undefined, getLabel('checkbox', 'marginLink'));
        ui.cbMarginLink.helpTip = getLabel('tooltip', 'marginLink');
        ui.cbMarginLink.value = true;

        ui.cbSquare = ui.marginPanel.add('checkbox', undefined, getLabel('checkbox', 'square'));
        ui.cbSquare.helpTip = getLabel('tooltip', 'square');
    }

    /**
     * 角丸パネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @param {string} rulerUnitLabel - 定規の単位
     * @returns {void}
     */
    function addRoundPanel(column, ui, rulerUnitLabel) {
        ui.roundPanel = addPanel(column, getLabel('panel', 'round'), ['left', 'top']);

        var roundRow = ui.roundPanel.add('group');
        roundRow.orientation = 'row';
        roundRow.alignChildren = ['left', 'center'];
        ui.cbRoundEnable = roundRow.add('checkbox', undefined, '');
        ui.cbRoundEnable.helpTip = getLabel('tooltip', 'round');
        ui.roundInput = addNumberField(roundRow, '2', getLabel('tooltip', 'round'));
        roundRow.add('statictext', undefined, rulerUnitLabel);
        /* OFF→ON で戻す半径 / Radius restored when rounding is turned back on */
        ui.lastRoundRadius = '2';

        var pillRow = ui.roundPanel.add('group');
        pillRow.orientation = 'row';
        pillRow.alignChildren = ['left', 'center'];
        ui.cbPill = pillRow.add('checkbox', undefined, getLabel('checkbox', 'pill'));
        ui.cbPill.helpTip = getLabel('tooltip', 'pill');
    }

    /**
     * グループパネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addGroupingPanel(column, ui) {
        var groupingPanel = addPanel(column, getLabel('panel', 'grouping'));
        ui.cbGroupWithText = groupingPanel.add('checkbox', undefined, getLabel('checkbox', 'groupWithText'));
        ui.cbGroupWithText.helpTip = getLabel('tooltip', 'groupWithText');
        ui.cbGroupWithText.value = true;

        ui.cbExclude = groupingPanel.add('checkbox', undefined, getLabel('checkbox', 'exclude'));
        ui.cbExclude.helpTip = getLabel('tooltip', 'exclude');
    }

    /**
     * 座標パネル（X / Y のずらし量）を追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @param {string} rulerUnitLabel - 定規の単位
     * @returns {void}
     */
    function addOffsetPanel(column, ui, rulerUnitLabel) {
        var offsetPanel = addPanel(column, withUnit(getLabel('panel', 'offset'), rulerUnitLabel));

        var offsetRow = offsetPanel.add('group');
        offsetRow.orientation = 'row';
        offsetRow.alignChildren = ['left', 'center'];

        offsetRow.add('statictext', undefined, labelText('fieldLabel', 'offsetX'));
        ui.offsetXInput = addNumberField(offsetRow, '0', getLabel('tooltip', 'offset'));
        offsetRow.add('statictext', undefined, labelText('fieldLabel', 'offsetY'));
        ui.offsetYInput = addNumberField(offsetRow, '0', getLabel('tooltip', 'offset'));
    }

    /**
     * 種別パネル（塗り / 線 / 線幅）を追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @param {string} strokeUnitLabel - 線幅の単位
     * @returns {void}
     */
    function addKindPanel(column, ui, strokeUnitLabel) {
        var kindPanel = addPanel(column, getLabel('panel', 'kind'));

        var kindRow = kindPanel.add('group');
        kindRow.orientation = 'row';
        kindRow.alignChildren = ['left', 'center'];
        kindRow.spacing = 10;

        ui.rbFill = kindRow.add('radiobutton', undefined, getLabel('radio', 'fill'));
        ui.rbFill.helpTip = getLabel('tooltip', 'fill');
        ui.rbStroke = kindRow.add('radiobutton', undefined, getLabel('radio', 'stroke'));
        ui.rbStroke.helpTip = getLabel('tooltip', 'stroke');
        ui.rbFill.value = true;
        ui.kindRadios = { fill: ui.rbFill, stroke: ui.rbStroke };

        ui.strokeWidthInput = addNumberField(kindRow, '1', getLabel('tooltip', 'stroke'));
        setFieldRange(ui.strokeWidthInput, 0, null);
        kindRow.add('statictext', undefined, strokeUnitLabel);
    }

    /**
     * CMYK の1版分（ラベルの下に数値欄）を追加する
     * @param {Group} parentContainer - 追加先
     * @param {string} plateLabel - 版の名前（C / M / Y / K）
     * @returns {EditText} 追加した数値欄
     */
    function addCmykPlateField(parentContainer, plateLabel) {
        var plateColumn = addColumnGroup(parentContainer, ['fill', 'top']);
        var plateText = plateColumn.add('statictext', undefined, plateLabel);
        plateText.justify = 'center';
        var plateInput = addNumberField(plateColumn, '0', getLabel('tooltip', 'cmykValue'));
        plateInput.characters = SHORT_FIELD_CHARS;
        setFieldRange(plateInput, 0, 100);
        return plateInput;
    }

    /**
     * カラーパネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addColorPanel(column, ui) {
        var colorPanel = addPanel(column, getLabel('panel', 'color'));

        var colorModeColumn = addColumnGroup(colorPanel, ['left', 'top']);
        ui.rbTextColor = colorModeColumn.add('radiobutton', undefined, getLabel('radio', 'textColor'));
        ui.rbTextColor.helpTip = getLabel('tooltip', 'textColor');
        ui.rbBlack = colorModeColumn.add('radiobutton', undefined, getLabel('radio', 'black'));
        ui.rbBlack.helpTip = getLabel('tooltip', 'black');
        ui.rbWhite = colorModeColumn.add('radiobutton', undefined, getLabel('radio', 'white'));
        ui.rbWhite.helpTip = getLabel('tooltip', 'white');
        ui.rbCmyk = colorModeColumn.add('radiobutton', undefined, getLabel('radio', 'cmyk'));
        ui.rbCmyk.helpTip = getLabel('tooltip', 'cmyk');
        ui.rbBlack.value = true;
        ui.colorRadios = { text: ui.rbTextColor, black: ui.rbBlack, white: ui.rbWhite, cmyk: ui.rbCmyk };

        ui.cmykGroup = colorPanel.add('group');
        ui.cmykGroup.orientation = 'row';
        ui.cmykGroup.alignChildren = ['center', 'top'];
        /* キーはセッション記憶の項目名と同じ / Keys match the session-memory fields */
        ui.cmykInputs = {
            cmykC: addCmykPlateField(ui.cmykGroup, 'C'),
            cmykM: addCmykPlateField(ui.cmykGroup, 'M'),
            cmykY: addCmykPlateField(ui.cmykGroup, 'Y'),
            cmykK: addCmykPlateField(ui.cmykGroup, 'K')
        };
    }

    /**
     * 不透明度パネルを追加する
     * @param {Group} column - 追加先の列
     * @param {Object} ui - コントロールの格納先
     * @returns {void}
     */
    function addOpacityPanel(column, ui) {
        var opacityPanel = addPanel(column, getLabel('panel', 'opacity'));

        var opacityRow = opacityPanel.add('group');
        opacityRow.orientation = 'row';
        ui.cbOpacityApply = opacityRow.add('checkbox', undefined, '');
        ui.cbOpacityApply.helpTip = getLabel('tooltip', 'opacity');
        ui.opacityInput = opacityRow.add('edittext', undefined, '60');
        ui.opacityInput.helpTip = getLabel('tooltip', 'opacity');
        ui.opacityInput.characters = 3;
        opacityRow.add('statictext', undefined, '%');
    }

    /**
     * ダイアログとコントロールを組み立てる（イベントはつながない）
     * @returns {Object} ダイアログ（dialog）と各コントロール
     */
    function buildDialog() {
        var ui = {};
        var rulerUnitLabel = getUnitInfo('rulerType').label;
        var strokeUnitLabel = getUnitInfo('strokeUnits').label;

        var dialog = new Window('dialog', getLabel('dialog', 'title') + ' ' + SCRIPT_VERSION);
        dialog.opacity = DIALOG_OPACITY;
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];
        ui.dialog = dialog;

        addShapeRadios(dialog, ui);

        var columnsRow = dialog.add('group');
        columnsRow.orientation = 'row';
        columnsRow.alignChildren = ['fill', 'top'];
        var leftColumn = addColumnGroup(columnsRow, ['fill', 'top']);
        var rightColumn = addColumnGroup(columnsRow, ['fill', 'top']);

        addScalePanel(leftColumn, ui);
        addMarginPanel(leftColumn, ui, rulerUnitLabel);
        addRoundPanel(leftColumn, ui, rulerUnitLabel);
        addGroupingPanel(leftColumn, ui);

        addOffsetPanel(rightColumn, ui, rulerUnitLabel);
        addKindPanel(rightColumn, ui, strokeUnitLabel);
        addColorPanel(rightColumn, ui);
        addOpacityPanel(rightColumn, ui);

        /* ボタンエリア（左・スペーサー・右） / Button row (left, spacer, right) */
        var btnRowGroup = dialog.add('group');
        btnRowGroup.orientation = 'row';
        btnRowGroup.alignment = ['fill', 'bottom'];

        var btnLeftGroup = btnRowGroup.add('group');
        btnLeftGroup.alignChildren = ['left', 'center'];
        ui.btnTransparencyGrid = btnLeftGroup.add('button', undefined, getLabel('button', 'transparencyGrid'));
        ui.btnTransparencyGrid.helpTip = getLabel('tooltip', 'transparencyGrid');

        var spacer = btnRowGroup.add('group');
        spacer.alignment = ['fill', 'fill'];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add('group');
        btnRightGroup.alignChildren = ['right', 'center'];
        ui.btnCancel = btnRightGroup.add('button', undefined, getLabel('button', 'cancel'), { name: 'cancel' });
        ui.btnOK = btnRightGroup.add('button', undefined, getLabel('button', 'ok'), { name: 'ok' });

        return ui;
    }

    // =========================================
    // ダイアログの連動 / Dialog state sync
    // =========================================

    /**
     * スケールパネルの有効／無効を切り替える（長方形は100%固定。ただし正方形ONのときは有効）
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncScalePanel(ui) {
        if (ui.rbRectangle.value && !ui.cbSquare.value) {
            ui.scaleInput.text = '100';
            ui.cbOneChar.value = false;
            ui.scalePanel.enabled = false;
        } else {
            ui.scalePanel.enabled = true;
        }
    }

    /**
     * マージンの「連動」を反映する（ONなら左右を上下にそろえて無効化）
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncMarginLink(ui) {
        if (ui.cbMarginLink.value) ui.marginHInput.text = String(ui.marginVInput.text);
        ui.marginHInput.enabled = !ui.cbMarginLink.value;
    }

    /**
     * マージンパネルは長方形のときだけ有効にする
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncMarginPanel(ui) {
        ui.marginPanel.enabled = ui.rbRectangle.value;
    }

    /**
     * 角丸の数値欄の有効／無効を切り替える（ピル形状は半径が自動なので無効）
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncRoundInput(ui) {
        if (ui.cbPill.value) {
            ui.roundInput.enabled = false;
            return;
        }
        if (ui.cbRoundEnable.value) {
            /* OFF→ON で控えた値を戻す / Restore the saved value when turned on */
            if (!ui.roundInput.enabled) ui.roundInput.text = String(ui.lastRoundRadius || ui.roundInput.text || '0');
            ui.roundInput.enabled = true;
        } else {
            /* ON→OFF で値を控えて無効化 / Save the value and disable when turned off */
            ui.lastRoundRadius = String(ui.roundInput.text);
            ui.roundInput.enabled = false;
        }
    }

    /**
     * 角丸パネルは長方形のときだけ有効にする
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncRoundPanel(ui) {
        ui.roundPanel.enabled = ui.rbRectangle.value;
        syncRoundInput(ui);
    }

    /**
     * 種別に合わせて線幅と中マド処理を切り替える（線のときは中マド処理を使えない）
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncKind(ui) {
        var strokeOn = ui.rbStroke.value;
        ui.strokeWidthInput.enabled = strokeOn;
        if (strokeOn) ui.cbExclude.value = false;
        ui.cbExclude.enabled = !strokeOn;
    }

    /**
     * CMYK 欄は CMYK を選んだときだけ有効にする
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncColor(ui) {
        ui.cmykGroup.enabled = ui.rbCmyk.value;
    }

    /**
     * 不透明度の数値欄はチェックONのときだけ有効にする
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncOpacity(ui) {
        ui.opacityInput.enabled = ui.cbOpacityApply.value;
    }

    /**
     * 形状に連動するパネルをまとめて切り替える
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncShapeDependentPanels(ui) {
        syncRoundPanel(ui);
        syncMarginPanel(ui);
        syncScalePanel(ui);
    }

    /**
     * すべての連動をまとめて反映する
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function syncAllPanels(ui) {
        syncColor(ui);
        syncKind(ui);
        syncMarginLink(ui);
        syncOpacity(ui);
        syncShapeDependentPanels(ui);
    }

    // =========================================
    // ダイアログの記憶 / Dialog memory
    // =========================================

    /**
     * 数値欄とセッション記憶の項目の対応表を返す
     * @param {Object} ui - コントロール
     * @returns {Object} 項目名 → EditText
     */
    function getTextFieldMap(ui) {
        var fieldMap = {
            scale: ui.scaleInput,
            marginV: ui.marginVInput,
            marginH: ui.marginHInput,
            roundValue: ui.roundInput,
            offsetX: ui.offsetXInput,
            offsetY: ui.offsetYInput,
            strokeWidth: ui.strokeWidthInput,
            opacity: ui.opacityInput
        };
        for (var key in ui.cmykInputs) fieldMap[key] = ui.cmykInputs[key];
        return fieldMap;
    }

    /**
     * チェックボックスとセッション記憶の項目の対応表を返す
     * @param {Object} ui - コントロール
     * @returns {Object} 項目名 → Checkbox
     */
    function getCheckboxMap(ui) {
        return {
            oneChar: ui.cbOneChar,
            marginLink: ui.cbMarginLink,
            marginSquare: ui.cbSquare,
            roundEnable: ui.cbRoundEnable,
            pill: ui.cbPill,
            groupWithText: ui.cbGroupWithText,
            exclude: ui.cbExclude,
            opacityApply: ui.cbOpacityApply
        };
    }

    /**
     * ダイアログの値をセッション記憶に控える
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function saveDialogState(ui) {
        var fieldMap = getTextFieldMap(ui);
        var checkboxMap = getCheckboxMap(ui);
        var key;
        for (key in fieldMap) dialogState[key] = String(fieldMap[key].text);
        for (key in checkboxMap) dialogState[key] = !!checkboxMap[key].value;
        dialogState.shape = getSelectedRadioKey(ui.shapeRadios, 'circle');
        dialogState.kind = getSelectedRadioKey(ui.kindRadios, 'fill');
        dialogState.colorMode = getSelectedRadioKey(ui.colorRadios, 'black');
        dialogState._hasSaved = true;
    }

    /**
     * セッション記憶の値をダイアログに戻す
     * @param {Object} ui - コントロール
     * @returns {void}
     */
    function restoreDialogState(ui) {
        if (!dialogState._hasSaved) return;
        var fieldMap = getTextFieldMap(ui);
        var checkboxMap = getCheckboxMap(ui);
        var key;
        for (key in fieldMap) fieldMap[key].text = String(dialogState[key]);
        for (key in checkboxMap) checkboxMap[key].value = !!dialogState[key];
        selectRadioKey(ui.shapeRadios, dialogState.shape, 'circle');
        selectRadioKey(ui.kindRadios, dialogState.kind, 'fill');
        selectRadioKey(ui.colorRadios, dialogState.colorMode, 'black');
        ui.lastRoundRadius = String(dialogState.roundValue);
    }

    // =========================================
    // 計測 / Measurement
    // =========================================

    /**
     * テキストの見た目の境界を測る（複製をアウトライン化して測り、結果は控えて使い回す）
     * @param {Object} context - 対象の情報
     * @returns {{left: number, top: number, right: number, bottom: number, w: number, h: number}} 矩形情報
     */
    function measureTextVisualBounds(context) {
        if (context.textVisualRect) return context.textVisualRect;
        var textDuplicate = null;
        var outlinedGroup = null;
        try {
            textDuplicate = context.textItem.duplicate();
            outlinedGroup = textDuplicate.createOutline();
            context.textVisualRect = boundsToRect(outlinedGroup.geometricBounds);
        } catch (e) {
            /* アウトライン化できないときはフレームの境界 / Fall back to the frame bounds */
            context.textVisualRect = boundsToRect(context.textItem.geometricBounds);
        } finally {
            safeRemove(outlinedGroup, 'outlinedGroup');
            safeRemove(textDuplicate, 'textDuplicate');
        }
        return context.textVisualRect;
    }

    /**
     * 選択オブジェクト全体の境界を測る（置き換える既存の背面図形は除く）
     * @param {Object} context - 対象の情報
     * @returns {{left: number, top: number, right: number, bottom: number, w: number, h: number}} 矩形情報
     */
    function measureSelectionRect(context) {
        var unionBounds = null;
        for (var i = 0; i < context.selectionItems.length; i++) {
            var item = context.selectionItems[i];
            if (!item || item === context.existingBackdrop) continue;
            var itemBounds = getItemBounds(item);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
            } else {
                unionBounds[0] = Math.min(unionBounds[0], itemBounds[0]);
                unionBounds[1] = Math.max(unionBounds[1], itemBounds[1]);
                unionBounds[2] = Math.max(unionBounds[2], itemBounds[2]);
                unionBounds[3] = Math.min(unionBounds[3], itemBounds[3]);
            }
        }
        return boundsToRect(unionBounds || getItemBounds(context.targetItem));
    }

    /**
     * 背面図形の基準にする矩形を返す
     * - テキスト：見た目の境界（1文字モードではフレームの境界）
     * - テキスト以外：選択全体の境界
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {{left: number, top: number, right: number, bottom: number, w: number, h: number}} 矩形情報
     */
    function measureBaseRect(ui, context) {
        if (!context.isTextMode) return measureSelectionRect(context);
        if (ui.cbOneChar.value) return boundsToRect(context.textItem.geometricBounds);
        return measureTextVisualBounds(context);
    }

    /**
     * マージンとスケールの初期値を、対象の短辺から決める（前回の値が無いときだけ呼ぶ）
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function applyDefaultsFromTarget(ui, context) {
        var baseRect = context.isTextMode ? measureTextVisualBounds(context) : measureSelectionRect(context);
        var shortSide = Math.min(Math.abs(baseRect.w), Math.abs(baseRect.h));
        if (!isFinite(shortSide) || shortSide <= 0) return;

        var defaultMargin = String(roundToTenth(shortSide / DEFAULT_MARGIN_DIVISOR));
        ui.marginVInput.text = defaultMargin;
        ui.marginHInput.text = defaultMargin;
        syncMarginLink(ui);

        ui.roundInput.text = String(roundToTenth(shortSide / DEFAULT_ROUND_DIVISOR));
        ui.lastRoundRadius = ui.roundInput.text;
    }

    /**
     * 数値欄の値を数値で読む
     * @param {EditText} editText - 数値欄
     * @param {number} fallbackValue - 数値にならないときの値
     * @returns {number} 読み取った値
     */
    function readNumber(editText, fallbackValue) {
        var value = parseFloat(editText.text);
        return isNaN(value) ? fallbackValue : value;
    }

    /**
     * ダイアログの値から、背面図形の中心と寸法を求める
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {{cx: number, cy: number, diameter: number, rectW: number, rectH: number, baseW: number, scaleRatio: number}} 図形の寸法
     */
    function computeBackdropGeometry(ui, context) {
        var baseRect = measureBaseRect(ui, context);

        var marginY = readNumber(ui.marginVInput, 0);
        var marginX = ui.cbMarginLink.value ? marginY : readNumber(ui.marginHInput, 0);
        var paddedW = baseRect.w + marginX * 2;
        var paddedH = baseRect.h + marginY * 2;

        var cx = baseRect.left + baseRect.w / 2;
        var cy = baseRect.top - baseRect.h / 2;

        var scalePercent = readNumber(ui.scaleInput, 100);
        if (scalePercent <= 0) scalePercent = 100;
        var scaleRatio = scalePercent / 100;

        var diameter;
        if (context.isTextMode && ui.cbOneChar.value) {
            /* 1文字モード：文字サイズを基準に、上端から文字サイズの半分を中心にする / Size from the font size */
            var fontSize;
            try {
                fontSize = context.textItem.textRange.characterAttributes.size;
            } catch (e) {
                fontSize = Math.max(baseRect.w, baseRect.h);
            }
            diameter = fontSize * ONE_CHAR_DIAMETER_RATIO * scaleRatio;
            cy = baseRect.top - fontSize / 2;
        } else {
            /* 正方形がすっぽり入る円の直径 / Diameter of the circle that encloses the square */
            diameter = Math.max(paddedW, paddedH) * Math.SQRT2 * scaleRatio;
        }

        cx += readNumber(ui.offsetXInput, 0);
        cy += readNumber(ui.offsetYInput, 0);

        /* 正方形ONのときは正円と同じ直径を一辺にする / A square uses the circle's diameter as its side */
        var squareOn = ui.cbSquare.value;
        return {
            cx: cx,
            cy: cy,
            diameter: diameter,
            rectW: squareOn ? diameter : paddedW * scaleRatio,
            rectH: squareOn ? diameter : paddedH * scaleRatio,
            baseW: baseRect.w,
            scaleRatio: scaleRatio
        };
    }

    // =========================================
    // 背面図形の作成 / Backdrop creation
    // =========================================

    /**
     * 背面図形のカラーを決める
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {Color} カラー
     */
    function resolveBackdropColor(ui, context) {
        if (ui.rbTextColor.value) {
            /* 混在した書式などで読めないことがある。そのときはブラック / Falls back to black when unreadable */
            try {
                var textColor = context.textItem.textRange.characterAttributes.fillColor;
                if (textColor) return textColor;
            } catch (e) { }
        }
        if (ui.rbWhite.value) return makeGrayColor(0);
        if (ui.rbCmyk.value) {
            var cmykColor = new CMYKColor();
            cmykColor.cyan = Math.min(100, Math.max(0, parseInt(ui.cmykInputs.cmykC.text, 10) || 0));
            cmykColor.magenta = Math.min(100, Math.max(0, parseInt(ui.cmykInputs.cmykM.text, 10) || 0));
            cmykColor.yellow = Math.min(100, Math.max(0, parseInt(ui.cmykInputs.cmykY.text, 10) || 0));
            cmykColor.black = Math.min(100, Math.max(0, parseInt(ui.cmykInputs.cmykK.text, 10) || 0));
            return cmykColor;
        }
        return makeGrayColor(100);
    }

    /**
     * 背面図形に塗り／線と不透明度を設定する
     * @param {PageItem} item - 背面図形
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function applyBackdropStyle(item, ui, context) {
        /* 図形の変換後などに色を受け付けないことがある。そのときは薄いグレーの塗り / Falls back to a light gray fill */
        try {
            var backdropColor = resolveBackdropColor(ui, context);
            if (ui.rbFill.value) {
                item.filled = true;
                item.stroked = false;
                item.fillColor = backdropColor;
            } else {
                item.filled = false;
                item.stroked = true;
                var strokeWidth = readNumber(ui.strokeWidthInput, 1);
                item.strokeWidth = (strokeWidth < 0) ? 1 : strokeWidth;
                item.strokeColor = backdropColor;
            }
        } catch (e) {
            try {
                item.fillColor = makeGrayColor(20);
                item.filled = true;
                item.stroked = false;
            } catch (eFallback) { }
        }

        var opacityPercent = Math.max(0, Math.min(100, readNumber(ui.opacityInput, 100)));
        item.opacity = ui.cbOpacityApply.value ? opacityPercent : 100;
    }

    /**
     * 長方形の背面図形を作る（角丸・ピル形状はライブ効果）
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} ui - コントロール
     * @param {Object} geometry - 図形の寸法
     * @returns {PathItem} 作った図形
     */
    function createRectangleBackdrop(doc, ui, geometry) {
        var rectTop = geometry.cy + geometry.rectH / 2;

        if (ui.cbPill.value) {
            /* ピル形状：高さの半分を半径にし、幅は少なくとも「元の幅＋高さ」 / Pill: radius is half the height */
            var minPillWidth = Math.abs(geometry.baseW) * geometry.scaleRatio + geometry.rectH;
            var pillWidth = Math.max(geometry.rectW, minPillWidth);
            var pillShape = doc.pathItems.rectangle(rectTop, geometry.cx - pillWidth / 2, pillWidth, geometry.rectH);
            applyRoundCornersLive(pillShape, geometry.rectH / 2);
            return pillShape;
        }

        var rectShape = doc.pathItems.rectangle(rectTop, geometry.cx - geometry.rectW / 2, geometry.rectW, geometry.rectH);
        if (ui.cbRoundEnable.value) applyRoundCornersLive(rectShape, readNumber(ui.roundInput, 0));
        return rectShape;
    }

    /**
     * 背面図形を作り、対象の背面に置いて色を付ける
     * @param {Document} doc - 対象ドキュメント
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {PathItem} 作った図形
     */
    function createBackdrop(doc, ui, context) {
        var geometry = computeBackdropGeometry(ui, context);
        var backdropShape;

        if (ui.rbRectangle.value) {
            backdropShape = createRectangleBackdrop(doc, ui, geometry);
        } else {
            var diameter = geometry.diameter;
            backdropShape = doc.pathItems.ellipse(geometry.cy + diameter / 2, geometry.cx - diameter / 2, diameter, diameter);
            if (ui.rbSuperellipse.value) morphPathToSuperellipse(backdropShape, geometry.cx, geometry.cy, diameter, diameter);
        }

        /* 対象の直後（背面）へ。移動できない位置なら最背面へ / Place behind the target, or send to back */
        try {
            backdropShape.move(context.targetItem, ElementPlacement.PLACEAFTER);
        } catch (e) {
            try {
                backdropShape.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (eZOrder) {
                logDebug('zOrder fallback failed: ' + eZOrder);
            }
        }

        applyBackdropStyle(backdropShape, ui, context);
        return backdropShape;
    }

    /**
     * 確定時の仕上げ：正円は「シェイプに変換」、長方形は「ライブパスファインダー（合体）」
     * @param {Document} doc - 対象ドキュメント
     * @param {PathItem} backdropShape - 背面図形
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {PageItem} 仕上げたオブジェクト
     */
    function finalizeBackdropShape(doc, backdropShape, ui, context) {
        if (ui.rbSuperellipse.value) return backdropShape;
        var menuCommand = ui.rbCircle.value ? 'Convert to Shape' : 'Live Pathfinder Add';
        try {
            var finalizedShape = runMenuCommandOnItem(doc, backdropShape, menuCommand);
            applyBackdropStyle(finalizedShape, ui, context);
            return finalizedShape;
        } catch (e) {
            logDebug(menuCommand + ' failed: ' + e);
            return backdropShape;
        }
    }

    /**
     * 背面図形を対象の背面に置く（グループ化がONなら対象と1つのグループにする）
     * @param {PageItem} backdropShape - 背面図形
     * @param {PageItem} targetItem - 対象オブジェクト
     * @param {boolean} groupWithTarget - 対象とグループ化するか
     * @returns {PageItem} 配置後のオブジェクト（グループ化したときはグループ）
     */
    function placeBackdrop(backdropShape, targetItem, groupWithTarget) {
        if (!groupWithTarget) {
            try {
                backdropShape.move(targetItem, ElementPlacement.PLACEAFTER);
            } catch (e) {
                logDebug('ungrouped backdrop move failed: ' + e);
            }
            return backdropShape;
        }

        try {
            var resultGroup = targetItem.parent.groupItems.add();
            try {
                resultGroup.move(targetItem, ElementPlacement.PLACEAFTER);
            } catch (eMoveGroup) {
                logDebug('group move failed: ' + eMoveGroup);
            }
            backdropShape.move(resultGroup, ElementPlacement.PLACEATEND);
            targetItem.move(resultGroup, ElementPlacement.PLACEATEND);
            try {
                backdropShape.zOrder(ZOrderMethod.SENDTOBACK);
            } catch (eSendToBack) {
                logDebug('grouped backdrop zOrder failed: ' + eSendToBack);
            }
            return resultGroup;
        } catch (eGroup) {
            alert(getLabel('alert', 'groupFailed') + ": " + eGroup);
            return backdropShape;
        }
    }

    /**
     * 「ライブパスファインダー（中マド）」で文字部分を抜く
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} resultItem - 配置後のオブジェクト
     * @returns {PageItem} 処理後のオブジェクト
     */
    function applyExclude(doc, resultItem) {
        try {
            return runMenuCommandOnItem(doc, resultItem, 'Live Pathfinder Exclude');
        } catch (e) {
            alert(getLabel('alert', 'excludeFailed') + ": " + e);
            doc.selection = null;
            return resultItem;
        }
    }

    // =========================================
    // 既存の背面図形の置き換え / Replacing an existing backdrop
    // =========================================

    /**
     * プレビュー用：既存の背面図形を隠す（グループはいったん解除して隠す。Undo で戻る）
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function hideReplacedBackdrop(context) {
        if (context.enclosingGroup) {
            try {
                releaseGroupChildren(context.enclosingGroup, context.existingBackdrop);
                context.enclosingGroup.hidden = true;
            } catch (e) {
                logDebug('hideReplacedBackdrop ungroup failed: ' + e);
            }
        }
        if (context.existingBackdrop) {
            try {
                context.existingBackdrop.hidden = true;
            } catch (e) {
                logDebug('hideReplacedBackdrop hide failed: ' + e);
            }
        }
    }

    /**
     * 確定用：既存の背面図形を削除する（グループは中身を出してから削除）
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function removeReplacedBackdrop(context) {
        if (context.enclosingGroup) {
            try {
                releaseGroupChildren(context.enclosingGroup, context.existingBackdrop);
            } catch (e) {
                logDebug('removeReplacedBackdrop ungroup failed: ' + e);
            }
        }
        safeRemove(context.existingBackdrop, 'existingBackdrop');
        safeRemove(context.enclosingGroup, 'enclosingGroup');
        context.existingBackdrop = null;
        context.enclosingGroup = null;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択から処理対象を決める
     * @param {PageItem[]} selectionItems - 選択オブジェクト
     * @returns {Object} 対象の情報（テキスト、配置の基準、既存の背面図形など）
     */
    function resolveContext(selectionItems) {
        var textItem = findFirstTextFrame(selectionItems);
        var found = findExistingBackdrop(selectionItems, textItem);
        return {
            selectionItems: selectionItems,
            textItem: textItem,
            isTextMode: !!textItem,
            /* テキストが無いときは、グループなど選択の先頭を配置の基準にする / Without text, the first selected item is the anchor */
            targetItem: textItem || found.target || selectionItems[0],
            existingBackdrop: found.backdrop,
            enclosingGroup: found.enclosingGroup,
            textVisualRect: null
        };
    }

    /**
     * 対象に合わせてダイアログの初期状態を調整する
     * @param {Object} ui - コントロール
     * @param {Object} context - 対象の情報
     * @returns {void}
     */
    function adjustDialogForTarget(ui, context) {
        /* 既存の背面図形があれば形状を合わせる / Match the shape of an existing backdrop */
        if (context.existingBackdrop) {
            selectRadioKey(ui.shapeRadios, detectShapeType(context.existingBackdrop), 'rect');
            syncShapeDependentPanels(ui);
        }

        if (!context.isTextMode) {
            /* テキストが無いときは「1文字」と「テキストカラー」を使わない / No single-character or text color without text */
            ui.cbOneChar.value = false;
            ui.cbOneChar.enabled = false;
            if (ui.rbTextColor.value) selectRadioKey(ui.colorRadios, 'black', 'black');
            ui.rbTextColor.enabled = false;
            return;
        }

        /* 空白を除いて1文字だけなら「1文字」をON / Turn on single-character mode for a single visible character */
        if (String(context.textItem.contents || '').replace(/\s+/g, '').length === 1) {
            ui.cbOneChar.value = true;
        }
    }

    /**
     * キー入力で形状を切り替える（E: 正円 / S: スーパー楕円 / R: 長方形）
     * @param {Object} ui - コントロール
     * @param {Function} onShapeChanged - 切り替えたあとに呼ぶ処理
     * @returns {void}
     */
    function addShapeKeyHandler(ui, onShapeChanged) {
        var shapeKeyMap = { E: 'circle', S: 'superellipse', R: 'rect' };
        ui.dialog.addEventListener('keydown', function (event) {
            var shapeKey = shapeKeyMap[event.keyName];
            if (!shapeKey) return;
            ui.shapeRadios[shapeKey].value = true;
            onShapeChanged();
            if (event.preventDefault) event.preventDefault();
        });
    }

    /**
     * ダイアログのイベントをつなぐ
     * @param {Object} ui - コントロール
     * @param {Function} updatePreview - プレビューを更新する処理
     * @returns {void}
     */
    function bindDialogEvents(ui, updatePreview) {
        /* 形状 / Shape */
        var onShapeChanged = function () {
            syncShapeDependentPanels(ui);
            updatePreview();
        };
        for (var shapeKey in ui.shapeRadios) ui.shapeRadios[shapeKey].onClick = onShapeChanged;
        addShapeKeyHandler(ui, onShapeChanged);

        /* スケール / Scale */
        bindNumberField(ui.scaleInput, updatePreview);
        ui.cbOneChar.onClick = updatePreview;

        /* マージン / Margin */
        bindNumberField(ui.marginVInput, function () {
            syncMarginLink(ui);
            updatePreview();
        });
        bindNumberField(ui.marginHInput, updatePreview);
        ui.cbMarginLink.onClick = function () {
            syncMarginLink(ui);
            updatePreview();
        };
        ui.cbSquare.onClick = function () {
            /* 正方形ONでピル形状をOFFにし、スケールを使えるようにする / Square turns pill off and enables scale */
            if (ui.cbSquare.value) {
                ui.cbPill.value = false;
                syncRoundInput(ui);
            }
            syncScalePanel(ui);
            if (ui.cbSquare.value) ui.scaleInput.active = true;
            updatePreview();
        };

        /* 角丸 / Round */
        bindNumberField(ui.roundInput, function () {
            ui.lastRoundRadius = String(ui.roundInput.text);
            updatePreview();
        });
        ui.cbRoundEnable.onClick = function () {
            syncRoundInput(ui);
            if (ui.cbRoundEnable.value && !ui.cbPill.value) ui.roundInput.active = true;
            updatePreview();
        };
        ui.cbPill.onClick = function () {
            syncRoundInput(ui);
            updatePreview();
        };

        /* グループ / Grouping */
        ui.cbExclude.onClick = function () {
            /* 中マド処理にはグループ化とテキストカラーが要る / Exclude needs grouping and the text color */
            if (ui.cbExclude.value) {
                ui.cbGroupWithText.value = true;
                selectRadioKey(ui.colorRadios, 'text', 'text');
                syncColor(ui);
            }
            updatePreview();
        };
        ui.cbGroupWithText.onClick = function () {
            if (!ui.cbGroupWithText.value) ui.cbExclude.value = false;
            updatePreview();
        };

        /* 座標 / Offset */
        bindNumberField(ui.offsetXInput, updatePreview);
        bindNumberField(ui.offsetYInput, updatePreview);

        /* 種別 / Kind */
        ui.rbFill.onClick = ui.rbStroke.onClick = function () {
            syncKind(ui);
            updatePreview();
        };
        bindNumberField(ui.strokeWidthInput, updatePreview);

        /* カラー / Color */
        var onColorChanged = function () {
            syncColor(ui);
            updatePreview();
        };
        for (var colorKey in ui.colorRadios) ui.colorRadios[colorKey].onClick = onColorChanged;
        for (var plateKey in ui.cmykInputs) bindNumberField(ui.cmykInputs[plateKey], updatePreview);

        /* 不透明度 / Opacity */
        ui.cbOpacityApply.onClick = function () {
            syncOpacity(ui);
            updatePreview();
        };
        bindNumberField(ui.opacityInput, updatePreview);
    }

    /**
     * エントリーポイント
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel('alert', 'noDocument'));
            return;
        }
        var doc = app.activeDocument;
        var selectionItems = doc.selection;
        if (!selectionItems || selectionItems.length === 0) {
            alert(getLabel('alert', 'noSelection'));
            return;
        }

        var context = resolveContext(selectionItems);
        var hadSavedState = dialogState._hasSaved;
        var previewManager = new PreviewManager();

        var ui = buildDialog();
        restoreDialogState(ui);
        syncAllPanels(ui);
        adjustDialogForTarget(ui, context);

        /**
         * プレビューを作り直し、ダイアログの値を控える
         * @returns {void}
         */
        function updatePreview() {
            previewManager.rollback();
            previewManager.addStep(function () {
                hideReplacedBackdrop(context);
                createBackdrop(doc, ui, context);
            });
            saveDialogState(ui);
        }

        bindDialogEvents(ui, updatePreview);

        ui.btnTransparencyGrid.onClick = function () {
            /* 表示の切り替えなので Undo 履歴には残らない / A view toggle, so it adds no undo step */
            app.executeMenuCommand('TransparencyGrid Menu Item');
            app.redraw();
        };

        ui.btnCancel.onClick = function () {
            previewManager.rollback();
            ui.dialog.close(0);
        };

        ui.dialog.onShow = function () {
            var location = ui.dialog.location;
            ui.dialog.location = [location[0] + DIALOG_OFFSET_X, location[1] + DIALOG_OFFSET_Y];
            ui.scaleInput.active = true;
            if (!hadSavedState) {
                /* 計測に失敗しても初期値のまま開く / Keep the stock values if measuring fails */
                try {
                    applyDefaultsFromTarget(ui, context);
                } catch (e) {
                    logDebug('applyDefaultsFromTarget failed: ' + e);
                }
            }
            updatePreview();
        };

        var dialogResult = ui.dialog.show();
        saveDialogState(ui);
        if (dialogResult !== 1) {
            previewManager.rollback();
            return;
        }

        /* プレビューを巻き戻し、確定処理を1回だけ実行（Undo 1段）/ Roll back, then run the final action once */
        previewManager.confirm(function () {
            removeReplacedBackdrop(context);
            var backdropShape = createBackdrop(doc, ui, context);
            backdropShape = finalizeBackdropShape(doc, backdropShape, ui, context);
            var resultItem = placeBackdrop(backdropShape, context.targetItem, ui.cbGroupWithText.value);
            if (ui.cbExclude.value) resultItem = applyExclude(doc, resultItem);
            doc.selection = null;
            try {
                resultItem.selected = true;
            } catch (e) {
                logDebug('result selection failed: ' + e);
            }
        });
    }

    try {
        main();
    } catch (err) {
        alert(getLabel('alert', 'genericError') + ": " + err);
    }

})();
