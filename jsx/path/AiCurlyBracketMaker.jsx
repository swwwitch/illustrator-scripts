#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

カーリーブラケット（波括弧）のパスを、半径・直線の長さ・線の太さ・角の形状・向き（上下左右）を指定して作成します。
オブジェクトを選択して実行するとその大きさと向きを初期値にして辺に沿わせ、値を変更するたびにプレビューが更新されます。

詳細は README を参照してください。

### Overview

Creates a curly bracket path from a radius, overall length, stroke width, corner shape and direction (up, down, left or right), centered on the artboard.
Running it with a selection seeds the size and direction from that selection and hugs its edge, and the preview updates as the values change.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "AiCurlyBracketMaker";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-05";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-05";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiCurlyBracketMaker.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/AiCurlyBracketMaker.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nd6b3e36ff79d"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    var DEFAULT_TOTAL_LENGTH_MM  = 100;         /* 全体の長さの初期値（mm）/ initial overall length (mm) */
    var DEFAULT_RADIUS_RATIO     = 1 / 15;      /* 半径の初期値＝全体の長さ×この比率 / initial radius = overall length x this ratio */
    var DEFAULT_LINK_RADIUS      = true;        /* 2つの半径を連動させて始めるか / whether the two radii start linked */
    var DEFAULT_CHAMFER          = false;       /* 面取りを初期状態でONにするか / whether the chamfer starts on */

    var DEFAULT_CENTER_OFFSET_MM = 0;           /* 中央の位置の初期値（mm、0で中央）/ initial center offset (mm, 0 = centered) */
    var DEFAULT_EXTENSION_PT     = 0;           /* 両端の延長の初期値（pt）/ initial extension at both ends (pt) */
    var DEFAULT_MARGIN_MM        = 2;           /* 選択オブジェクトとの余白の初期値（mm）/ initial gap from the selection (mm) */
    var DEFAULT_STROKE_WIDTH_PT  = 2;           /* 線の太さの初期値（pt）/ initial stroke width (pt) */
    var DEFAULT_DIRECTION        = "right";     /* 初期の向き（DIRECTION_KEYS のいずれか）/ initial direction (one of DIRECTION_KEYS) */
    var DEFAULT_STROKE_CAP       = "buttCap";   /* 初期の線端 / initial stroke cap */
    var DEFAULT_CORNER_JOIN      = "miterJoin"; /* 初期の角の形状 / initial corner shape */

    /* 生成したブラケットに付ける目印。値には向きを入れ、選び直したときの復元に使う / Marker added to the bracket we create; its value holds the direction, used when it is selected again */
    var BRACKET_TAG_NAME = "AiCurlyBracketMaker";

    /* 面取りに使うジグザグ効果（大きさ0・折り返し0・roundness 0＝直線的に）/ Zig Zag effect for the chamfer: amount 0, ridges 0, roundness 0 (corner points) */
    var CHAMFER_EFFECT_XML = '<LiveEffect name="Adobe Zigzag"><Dict data="R amount 0 R relAmount 0 R absoluteness 1 R ridges 0 R roundness 0 "/></LiveEffect>';

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 8;                /* パネル内の要素間隔 / panel spacing */
    var FIELD_ROW_SPACING  = 6;                /* ラベル・入力欄・単位表記の間隔 / gap inside a labeled row */
    var LABEL_WIDTH        = 98;               /* 行ラベルの共通幅 / shared width of row labels */
    var FIELD_CHARACTERS   = 3;                /* 数値欄の文字数（＝最小幅）/ characters of a numeric field */
    var BUTTON_BAR_MARGINS = [0, 10, 0, 0];    /* ボタンバーの余白 / margins of the bottom button bar */
    var BUTTON_BAR_SPACING = 10;               /* ボタンバー内の要素間隔 / spacing inside the button bar */

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
        if (localeText.indexOf("ja") === 0) {
            return "ja";
        }
        return "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "カーリーブラケットの作成", en: "Create Curly Bracket" }
        },
        panel: {
            shapeAndSize: { ja: "形状と大きさ", en: "Shape & Size" },
            stroke:       { ja: "線", en: "Stroke" }
        },
        fieldLabel: {
            centerRadius: { ja: "半径（中央）", en: "Center Radius" },
            endRadius:    { ja: "半径（両端）", en: "End Radius" },
            centerOffset: { ja: "中央の位置", en: "Center Offset" },
            totalLength:  { ja: "直線の長さ", en: "Length" },
            extension:    { ja: "延長", en: "Extension" },
            margin:       { ja: "余白", en: "Margin" },
            strokeWidth:  { ja: "線の太さ", en: "Stroke Width" },
            strokeCap:    { ja: "線端", en: "Cap" },
            cornerJoin:   { ja: "角の形状", en: "Corner Shape" },
            direction:    { ja: "向き", en: "Direction" }
        },
        checkbox: {
            linkRadius: { ja: "連動", en: "Link" },
            chamfer:    { ja: "面取り", en: "Chamfer" }
        },
        radio: {
            buttCap:   { ja: "なし", en: "None" },
            roundCap:  { ja: "丸型", en: "Round" },
            miterJoin: { ja: "マイター結合", en: "Miter Join" },
            roundJoin: { ja: "ラウンド結合", en: "Round Join" },
            up:        { ja: "上", en: "Up" },
            down:      { ja: "下", en: "Down" },
            left:      { ja: "左", en: "Left" },
            right:     { ja: "右", en: "Right" }
        },
        tooltip: {
            centerRadius: { ja: "中央の突起を作る円弧の半径。↑↓で増減、Shift+↑↓で10単位スナップ", en: "Radius of the arcs that form the point in the middle. Up/Down to step, Shift+Up/Down snaps to 10" },
            endRadius:    { ja: "両端で外へ折れ返る円弧の半径。0にすると円弧なしの直角になります", en: "Radius of the arcs that curl outward at both ends. 0 leaves a right angle with no arc" },
            linkRadius:   { ja: "両端の半径を中央に合わせる（ONのあいだ両端は編集できません）", en: "Keep the end radius equal to the center radius (the end field is disabled while on)" },
            chamfer:      { ja: "ジグザグ効果（大きさ0・折り返し0）を適用し、円弧を直線でつないだ面取りにします", en: "Applies a Zig Zag effect (size 0, ridges 0) so the arcs become straight chamfers" },
            centerOffset: { ja: "中央の突起を長さ方向にずらす量。0で中央、左右の向きでは上へ、上下の向きでは右へ動きます（マイナスで逆）", en: "Moves the middle point along the length. 0 keeps it centered; up for a left/right bracket, right for an up/down one (negative reverses)" },
            totalLength:  { ja: "ブラケット全体の長さ。半径を変えても総長は変わらず、直線部分が自動で伸縮します", en: "Overall length of the bracket. Changing the radius keeps this length and resizes the straight sections instead" },
            extension:    { ja: "両端から外側へ伸ばす直線の長さ（0で延長なし）", en: "Straight run added at both ends, away from the middle (0 adds none)" },
            margin:       { ja: "選択オブジェクトとブラケットのあいだの間隔（選択して実行したときのみ有効）", en: "Gap between the selection and the bracket (only when run with a selection)" },
            strokeWidth:  { ja: "ブラケットの線幅（pt）", en: "Stroke width of the bracket (pt)" },
            strokeCap:    { ja: "両端の線の先を丸めるかどうか", en: "Whether both ends of the stroke are rounded" },
            cornerJoin:   { ja: "中央の角を尖らせるか丸めるか", en: "Whether the corner in the middle is pointed or rounded" },
            direction:    { ja: "中央の突起を向ける方向。選択オブジェクトがあれば、その辺に沿って配置されます", en: "The direction the middle point faces. With a selection, the bracket hugs the matching edge" },
            create:       { ja: "プレビューの状態で確定します（大きさの参照にしたパスは削除されます）", en: "Commits exactly what the preview shows (a path used as the size reference is removed)" }
        },
        button: {
            create: { ja: "作成", en: "Create" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            lockedLayer:  { ja: "アクティブレイヤーがロックまたは非表示です。", en: "The active layer is locked or hidden." },
            invalidValue: { ja: "数値が正しくありません。半径（中央）と線の太さは0より大きい値、ほかの項目は0以上を入力してください。", en: "Some values are not valid. The center radius and the stroke width must be greater than 0, and the other fields 0 or more." }
        }
    };

    /**
     * LABELS からカテゴリを辿って現在の言語のラベルを取得する（例: getLabel('radio','right')）
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
     * コロン付きの項目名を返す（日本語は全角、英語は半角）
     * @param {...string} keys - LABELS を辿るキー列
     * @returns {string} コロンを付けたラベル
     */
    function labelText() {
        return getLabel.apply(null, arguments) + (uiLang === "ja" ? "：" : ":");
    }

    // =========================================
    // UIレイアウト補助 / UI layout helpers
    // =========================================

    /**
     * ダイアログ全体の並びと余白を設定する
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
     * パネルの並びと余白を設定する
     * @param {Panel} targetPanel - 対象パネル
     * @returns {void}
     */
    function setupPanel(targetPanel) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = PANEL_SPACING;
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
     * 共通幅で右揃えの行ラベルを追加する
     * @param {Group} parentRow - 追加先の行グループ
     * @param {string} rowLabelText - 表示するラベル
     * @returns {StaticText} 生成したラベル
     */
    function addRowLabel(parentRow, rowLabelText) {
        var rowLabel = parentRow.add("statictext", undefined, rowLabelText);
        rowLabel.preferredSize.width = LABEL_WIDTH;
        rowLabel.justify = "right";
        return rowLabel;
    }

    /**
     * 「ラベル＋数値欄＋単位」の1行を生成する
     * @param {Panel|Group} parentContainer - 追加先
     * @param {string} rowLabelText - 行ラベル（コロン付き）
     * @param {string} initialValue - 数値欄の初期値
     * @param {string} unitSuffix - 数値欄に添える単位表記
     * @param {string} helpTipText - 数値欄のツールチップ
     * @returns {EditText} 生成した数値欄
     */
    function addNumberFieldRow(parentContainer, rowLabelText, initialValue, unitSuffix, helpTipText) {
        var numberFieldRow = parentContainer.add("group");
        setupRow(numberFieldRow, "left", FIELD_ROW_SPACING);
        addRowLabel(numberFieldRow, rowLabelText);

        var numberField = numberFieldRow.add("edittext", undefined, initialValue);
        numberField.characters = FIELD_CHARACTERS;
        numberField.helpTip = helpTipText;

        numberFieldRow.add("statictext", undefined, unitSuffix);
        return numberField;
    }

    // =========================================
    // 入力欄の補助 / Input field helpers
    // =========================================

    /**
     * 入力欄に入れる数値を小数2桁までに丸めて文字列にする
     * @param {number} value - 表示したい値
     * @returns {string} 入力欄に入れる文字列
     */
    function formatFieldValue(value) {
        return String(Math.round(value * 100) / 100);
    }

    /**
     * 入力欄に↑↓キーでの値増減を追加する（Shiftで10単位スナップ、Optionで0.1刻み）
     * @param {EditText} editText - 対象の入力欄
     * @param {function} [onChanged] - 値を更新したあとに呼ぶコールバック
     * @param {number} [minValue] - 下限値（省略時は制限なし）
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
    // ブラケットの形状 / Bracket geometry
    // =========================================

    /* mm を pt に変換する係数 / Factor converting mm to pt */
    var MM_TO_PT = 72 / 25.4;

    /* ベジェ曲線で90度円弧を描くための制御点ハンドル係数 / Handle factor for a 90-degree Bezier arc */
    var KAPPA = 0.55228474983;

    /* 線端ごとの端点の形状 / Stroke cap per end shape */
    var STROKE_CAPS = {
        buttCap:  StrokeCap.BUTTENDCAP,
        roundCap: StrokeCap.ROUNDENDCAP
    };

    /* 角の形状ごとの線の結合 / Stroke join per corner shape */
    var CORNER_JOINS = {
        miterJoin: StrokeJoin.MITERENDJOIN,
        roundJoin: StrokeJoin.ROUNDENDJOIN
    };

    /* 選択できる向き（ラジオの並び順）/ Selectable directions, in the order the radios appear */
    var DIRECTION_KEYS = ["up", "down", "left", "right"];

    /* 右向きを基準にした向きごとの回転量（cos/sin）/ Rotation per direction, based on the right-facing bracket */
    var DIRECTION_ROTATIONS = {
        right: { cos:  1, sin:  0 },
        up:    { cos:  0, sin:  1 },
        left:  { cos: -1, sin:  0 },
        down:  { cos:  0, sin: -1 }
    };

    /**
     * 向きに対応する回転量を返す
     * @param {string} direction - 向きのキー（DIRECTION_KEYS のいずれか）
     * @returns {{cos: number, sin: number}} 回転量（未知のキーは既定の向き）
     */
    function getDirectionRotation(direction) {
        return DIRECTION_ROTATIONS[direction] || DIRECTION_ROTATIONS[DEFAULT_DIRECTION];
    }

    /**
     * 座標を中心まわりに回転する
     * @param {number[]} basePoint - 回転前の座標 [x, y]
     * @param {{cos: number, sin: number}} rotation - 回転量（cos/sin）
     * @param {number} centerX - 回転中心のX座標（pt）
     * @param {number} centerY - 回転中心のY座標（pt）
     * @returns {number[]} 回転後の座標 [x, y]
     */
    function rotateAroundCenter(basePoint, rotation, centerX, centerY) {
        var dx = basePoint[0] - centerX;
        var dy = basePoint[1] - centerY;
        return [
            centerX + dx * rotation.cos - dy * rotation.sin,
            centerY + dx * rotation.sin + dy * rotation.cos
        ];
    }

    /**
     * ハンドルを持たない（前後と直線でつながる）パスポイントを返す
     * @param {number} anchorX - アンカーのX座標（pt）
     * @param {number} anchorY - アンカーのY座標（pt）
     * @returns {{anchor: number[], left: number[], right: number[]}} パスポイント
     */
    function createStraightPoint(anchorX, anchorY) {
        return {
            anchor: [anchorX, anchorY],
            left:   [anchorX, anchorY],
            right:  [anchorX, anchorY]
        };
    }

    /**
     * @typedef {object} BracketSettings
     * @property {number} centerRadiusPt - 中央の突起を作る円弧の半径（pt）
     * @property {number} endRadiusPt - 両端で外へ折れ返る円弧の半径（pt）
     * @property {number} totalLengthPt - 全体の長さ（pt）
     * @property {number} centerOffsetPt - 中央の突起を長さ方向にずらす量（pt、0で中央）
     * @property {number} extensionPt - 両端から外側へ伸ばす長さ（pt）
     * @property {number} marginPt - 選択オブジェクトとの余白（pt）
     * @property {number} strokeWidthPt - 線の太さ（pt）
     * @property {string} strokeCap - 線端（STROKE_CAPS のキー）
     * @property {string} cornerJoin - 角の形状（CORNER_JOINS のキー）
     * @property {string} direction - 向き（DIRECTION_KEYS のいずれか）
     * @property {boolean} isChamfer - 面取り（ジグザグ効果）を適用するか
     */

    /**
     * 中央をずらす向きを返す（左右の向きでは上が＋、上下の向きでは右が＋）
     * @param {string} direction - 向きのキー（DIRECTION_KEYS のいずれか）
     * @returns {number} 右向き基準の座標系での符号（+1 または -1）
     */
    function getCenterOffsetSign(direction) {
        return (direction === "left" || direction === "up") ? -1 : 1;
    }

    /**
     * 座標を水平の中心線で反転する
     * @param {number[]} basePoint - 反転前の座標 [x, y]
     * @param {number} centerY - 中心線のY座標（pt）
     * @returns {number[]} 反転後の座標 [x, y]
     */
    function mirrorAcrossCenterY(basePoint, centerY) {
        return [basePoint[0], 2 * centerY - basePoint[1]];
    }

    /**
     * 上半分のパスポイントを反転して下半分のパスポイントにする
     * 進行方向が逆になるので、左右のハンドルも入れ替える
     * @param {{anchor: number[], left: number[], right: number[]}} halfPoint - 上半分のパスポイント
     * @param {number} centerY - 中心線のY座標（pt）
     * @returns {{anchor: number[], left: number[], right: number[]}} 下半分のパスポイント
     */
    function mirrorPointAcrossCenterY(halfPoint, centerY) {
        return {
            anchor: mirrorAcrossCenterY(halfPoint.anchor, centerY),
            left:   mirrorAcrossCenterY(halfPoint.right, centerY),
            right:  mirrorAcrossCenterY(halfPoint.left, centerY)
        };
    }

    /**
     * 同じ位置に来た2つのパスポイントを1つにまとめる（入りのハンドルは前から、出のハンドルは後ろから）
     * @param {{anchor: number[], left: number[], right: number[]}} firstPoint - 前のパスポイント
     * @param {{anchor: number[], left: number[], right: number[]}} secondPoint - 後ろのパスポイント
     * @returns {{anchor: number[], left: number[], right: number[]}} まとめたパスポイント
     */
    function mergePathPoints(firstPoint, secondPoint) {
        return {
            anchor: secondPoint.anchor,
            left:   firstPoint.left,
            right:  secondPoint.right
        };
    }

    /**
     * 片側のパスポイントを、腕の先から中央の突起の手前まで順に返す（右向き基準）
     * 下半分はこの結果を中央で反転して使う
     * @param {number} endRadius - 両端で外へ折れ返る円弧の半径（pt）
     * @param {number} centerRadius - 中央の突起を作る円弧の半径（pt）
     * @param {number} straightLength - この側の直線部分の長さ（pt）
     * @param {number} extension - 腕の先に足す延長（pt）
     * @param {number} centerX - 中心軸のX座標（pt）
     * @param {number} centerY - 中央の突起のY座標（pt）
     * @returns {Array<{anchor: number[], left: number[], right: number[]}>} 片側のパスポイント
     */
    function buildHalfPoints(endRadius, centerRadius, straightLength, extension, centerX, centerY) {
        var endHandle = endRadius * KAPPA;
        var centerHandle = centerRadius * KAPPA;
        var straightTopY = centerY + centerRadius + straightLength; /* 直線の上端 / Top of the straight section */
        var straightBottomY = centerY + centerRadius;               /* 直線の下端 / Bottom of the straight section */
        var armEndY = straightTopY + endRadius;                  /* 腕の先のY座標 / Y of the arm end */
        /* 腕の先（両端の円弧の外端）/ Arm end (outer end of the end arc) */
        var armEndPoint = {
            anchor: [centerX - endRadius, armEndY],
            left:   [centerX - endRadius, armEndY],
            right:  [centerX - (endRadius - endHandle), armEndY]
        };
        /* 直線の開始点（両端の円弧の内端）/ Start of the straight section */
        var straightTopPoint = {
            anchor: [centerX, straightTopY],
            left:   [centerX, straightTopY + endHandle],
            right:  [centerX, straightTopY]
        };
        /* 直線の終了点（中央の円弧の外端）/ End of the straight section */
        var straightBottomPoint = {
            anchor: [centerX, straightBottomY],
            left:   [centerX, straightBottomY],
            right:  [centerX, straightBottomY - centerHandle]
        };

        var halfPoints = [];

        /* 延長線（中央とは反対の外側へ伸ばす）/ Extension, running outward, away from the middle */
        if (extension > 0) {
            halfPoints.push(createStraightPoint(centerX - endRadius - extension, armEndY));
        }

        /* 半径0のときは腕の先と直線の開始点が重なり、円弧のない直角になる / A zero radius merges the arm end into the straight start, leaving a right angle */
        var armPoint = (endRadius > 0) ? armEndPoint : mergePathPoints(armEndPoint, straightTopPoint);
        halfPoints.push(armPoint);

        if (straightLength > 0) {
            if (endRadius > 0) halfPoints.push(straightTopPoint);
            halfPoints.push(straightBottomPoint);
        } else if (endRadius > 0) {
            /* 直線がないときは変曲点ひとつにまとめる / Without a straight section the two points merge into an inflection point */
            halfPoints.push(mergePathPoints(straightTopPoint, straightBottomPoint));
        } else {
            /* 直線も円弧もないときは、腕の先まで含めてひとつにまとめる / With neither, the arm end merges in as well */
            halfPoints[halfPoints.length - 1] = mergePathPoints(armPoint, straightBottomPoint);
        }
        return halfPoints;
    }

    /**
     * ブラケットのアンカーポイントと方向線を計算する
     * 上半分だけを組み立て、中央の突起をはさんで鏡像を並べ、最後に向きへ回転する
     * @param {BracketSettings} bracketSettings - ブラケットの設定値
     * @param {number} centerX - 配置位置の中心X（pt）
     * @param {number} centerY - 配置位置の中心Y（pt）
     * @returns {Array<{anchor: number[], left: number[], right: number[]}>} 向きを反映したパスポイント
     */
    function buildBracketPoints(bracketSettings, centerX, centerY) {
        var endRadius = bracketSettings.endRadiusPt;
        var centerRadius = bracketSettings.centerRadiusPt;
        var centerHandle = centerRadius * KAPPA;
        var extension = bracketSettings.extensionPt;

        /* 全体の長さから円弧ぶん（片側 両端＋中央）を差し引いた直線部分 / The straight run left after the arcs (end + center per side) */
        var straightLength = Math.max(0, bracketSettings.totalLengthPt / 2 - endRadius - centerRadius);

        /* 中央のずらし量。全長を保つため、直線が尽きるところで止める / Offset of the middle point, capped where the straight run runs out so the overall length holds */
        var centerOffset = getCenterOffsetSign(bracketSettings.direction) * bracketSettings.centerOffsetPt;
        centerOffset = Math.max(-straightLength, Math.min(straightLength, centerOffset));

        /* ずらしたぶん、上下で直線の長さが変わる / The offset makes the two halves differ */
        var beakY = centerY + centerOffset;
        var upperPoints = buildHalfPoints(endRadius, centerRadius, straightLength - centerOffset, extension, centerX, beakY);
        var lowerPoints = buildHalfPoints(endRadius, centerRadius, straightLength + centerOffset, extension, centerX, beakY);


        var bracketPoints = upperPoints.slice();
        var i;

        /* 中央の突起は上下の折り返し点なので反転しない / The middle point is where the halves meet, so it is not mirrored */
        bracketPoints.push({
            anchor: [centerX + centerRadius, beakY],
            left:   [centerX + (centerRadius - centerHandle), beakY],
            right:  [centerX + (centerRadius - centerHandle), beakY]
        });

        /* 下半分は中央で折り返した鏡像を逆順に並べる / The lower half is mirrored across the middle point, in reverse order */
        for (i = lowerPoints.length - 1; i >= 0; i--) {
            bracketPoints.push(mirrorPointAcrossCenterY(lowerPoints[i], beakY));
        }

        /* 右向きで組み立てた座標を、選んだ向きへまとめて回転 / Rotate the right-facing points into the chosen direction */
        var rotation = getDirectionRotation(bracketSettings.direction);
        for (i = 0; i < bracketPoints.length; i++) {
            bracketPoints[i].anchor = rotateAroundCenter(bracketPoints[i].anchor, rotation, centerX, centerY);
            bracketPoints[i].left = rotateAroundCenter(bracketPoints[i].left, rotation, centerX, centerY);
            bracketPoints[i].right = rotateAroundCenter(bracketPoints[i].right, rotation, centerX, centerY);
        }
        return bracketPoints;
    }

    /**
     * @typedef {object} SelectionReference
     * @property {number[]} bounds - 選択全体の外接矩形 [左, 上, 右, 下]
     * @property {PathItem|null} sizeReferenceItem - 大きさの参照にした開いたパス（囲む対象のときは null）
     * @property {string|null} bracketDirection - このスクリプトで作ったブラケットなら、その向き
     */

    /**
     * このスクリプトで作ったブラケットかどうかを目印のタグで判定する
     * @param {PathItem|null} pathItem - 調べるパス
     * @returns {string|null} 記録されていた向き（ブラケットでなければ null）
     */
    function readBracketDirection(pathItem) {
        if (!pathItem) return null;
        for (var i = 0; i < pathItem.tags.length; i++) {
            if (pathItem.tags[i].name === BRACKET_TAG_NAME) {
                return DIRECTION_ROTATIONS[pathItem.tags[i].value] ? pathItem.tags[i].value : null;
            }
        }
        return null;
    }

    /**
     * 選択オブジェクトを基準情報として読み取る
     * 開いたパス1本だけなら「大きさの参照」、それ以外は「囲む対象」として扱う
     * @param {Document} doc - 対象ドキュメント
     * @returns {SelectionReference|null} 基準の情報（境界を持つ選択がなければ null）
     */
    function readSelectionReference(doc) {
        var selectedItems = doc.selection;
        if (!selectedItems || selectedItems.length === 0) return null;

        var selectionBounds = null;
        var boundedItemCount = 0;
        var firstBoundedItem = null;

        for (var i = 0; i < selectedItems.length; i++) {
            var itemBounds = null;
            /* 文字選択（TextRange）には geometricBounds がないので、取れないものは飛ばす / A text range has no geometricBounds, so skip what cannot be read */
            try {
                itemBounds = selectedItems[i].geometricBounds;
            } catch (eBounds) {
                itemBounds = null;
            }
            if (!itemBounds) continue;

            boundedItemCount++;
            if (!firstBoundedItem) firstBoundedItem = selectedItems[i];

            if (!selectionBounds) {
                selectionBounds = [itemBounds[0], itemBounds[1], itemBounds[2], itemBounds[3]];
            } else {
                selectionBounds[0] = Math.min(selectionBounds[0], itemBounds[0]);
                selectionBounds[1] = Math.max(selectionBounds[1], itemBounds[1]);
                selectionBounds[2] = Math.max(selectionBounds[2], itemBounds[2]);
                selectionBounds[3] = Math.min(selectionBounds[3], itemBounds[3]);
            }
        }
        if (!selectionBounds) return null;

        /* 閉じたパスや図形・テキストは囲む対象なので参照にしない / Closed paths, shapes and text are things to bracket, not references */
        var isSizeReference = (boundedItemCount === 1 && firstBoundedItem.typename === "PathItem" && !firstBoundedItem.closed);
        var sizeReferenceItem = isSizeReference ? firstBoundedItem : null;
        return {
            bounds: selectionBounds,
            sizeReferenceItem: sizeReferenceItem,
            bracketDirection: readBracketDirection(sizeReferenceItem)
        };
    }

    /**
     * 大きさの参照にしたパスの表示・非表示を切り替える（囲む対象のときは何もしない）
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @param {boolean} isHidden - 隠すなら true、戻すなら false
     * @returns {void}
     */
    function setSizeReferenceHidden(selectionReference, isHidden) {
        if (!selectionReference || !selectionReference.sizeReferenceItem) return;
        selectionReference.sizeReferenceItem.hidden = isHidden;
    }

    /**
     * 大きさの参照にしたパスを削除する（囲む対象のときは何もしない）
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @returns {void}
     */
    function removeSizeReference(selectionReference) {
        if (!selectionReference || !selectionReference.sizeReferenceItem) return;
        selectionReference.sizeReferenceItem.remove();
    }

    /**
     * 向きに合わせて、基準の外接矩形から長さを取る
     * 左右の向きなら高さ、上下の向きなら幅
     * @param {SelectionReference} selectionReference - 基準の情報
     * @param {string} direction - 向きのキー
     * @returns {number} 全体の長さ（mm）
     */
    function getReferenceLengthMm(selectionReference, direction) {
        var referenceBounds = selectionReference.bounds;
        var isVertical = (direction === "left" || direction === "right");
        var referenceLengthPt = isVertical
            ? (referenceBounds[1] - referenceBounds[3])
            : (referenceBounds[2] - referenceBounds[0]);
        return referenceLengthPt / MM_TO_PT;
    }

    /**
     * ダイアログの初期値（長さ・半径・向き）を決める
     * 選択があれば選択から決め直し、選択がなければ前回ダイアログを閉じたときの値を復元する
     * 参照のパスは左（縦長）・上（横長）に、囲む対象は右（縦長）・下（横長）に置く
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @param {object} savedSettings - 同一セッション内に記憶していた値
     * @returns {{totalLengthMm: number, centerRadiusMm: number, endRadiusMm: number, direction: string}} 初期値
     */
    function getInitialValues(selectionReference, savedSettings) {
        if (!selectionReference) {
            /* 選択がないときは前回の値をそのまま復元する / With no selection, restore what was there last time */
            var savedLengthMm = readSavedNumber(savedSettings.totalLength, DEFAULT_TOTAL_LENGTH_MM);
            var savedRadiusMm = savedLengthMm * DEFAULT_RADIUS_RATIO;
            return {
                totalLengthMm: savedLengthMm,
                centerRadiusMm: readSavedNumber(savedSettings.centerRadius, savedRadiusMm),
                endRadiusMm: readSavedNumber(savedSettings.endRadius, savedRadiusMm),
                direction: DIRECTION_ROTATIONS[savedSettings.direction] ? savedSettings.direction : DEFAULT_DIRECTION
            };
        }

        var referenceBounds = selectionReference.bounds;
        var referenceWidth = referenceBounds[2] - referenceBounds[0];
        var referenceHeight = referenceBounds[1] - referenceBounds[3];
        var isSizeReference = (selectionReference.sizeReferenceItem !== null);
        var direction;

        if (selectionReference.bracketDirection) {
            /* このスクリプトで作ったブラケットは、その向きのまま描き直す / A bracket we made keeps the direction it was drawn with */
            direction = selectionReference.bracketDirection;
        } else if (referenceHeight >= referenceWidth) {
            direction = isSizeReference ? "left" : "right";
        } else {
            direction = isSizeReference ? "up" : "down";
        }

        var totalLengthMm = getReferenceLengthMm(selectionReference, direction);
        /* 半径（中央・両端とも）は全体の長さに追従させる（以後は個別に変更できる）/ Both radii follow the overall length, and can be changed on their own afterwards */
        var radiusMm = totalLengthMm * DEFAULT_RADIUS_RATIO;

        return {
            totalLengthMm: totalLengthMm,
            centerRadiusMm: radiusMm,
            endRadiusMm: radiusMm,
            direction: direction
        };
    }

    /**
     * アクティブアートボードの中心座標を返す
     * @param {Document} doc - 対象ドキュメント
     * @returns {{x: number, y: number}} 中心座標（pt）
     */
    function getActiveArtboardCenter(doc) {
        var activeArtboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var artboardRect = activeArtboard.artboardRect; /* [左, 上, 右, 下] / [left, top, right, bottom] */
        return {
            x: (artboardRect[0] + artboardRect[2]) / 2,
            y: (artboardRect[1] + artboardRect[3]) / 2
        };
    }

    /**
     * ブラケットを置く中心座標を返す
     * このスクリプトで作ったブラケットを選んでいれば、その中央に重ねる
     * それ以外の選択では、腕の先が外接矩形の辺に触れる位置に置く
     * @param {Document} doc - 対象ドキュメント
     * @param {BracketSettings} bracketSettings - ブラケットの設定値
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @returns {{x: number, y: number}} 中心座標（pt）
     */
    function getPlacementCenter(doc, bracketSettings, selectionReference) {
        if (!selectionReference) return getActiveArtboardCenter(doc);

        var referenceBounds = selectionReference.bounds;
        var rotation = getDirectionRotation(bracketSettings.direction);
        var referenceCenterX = (referenceBounds[0] + referenceBounds[2]) / 2;
        var referenceCenterY = (referenceBounds[1] + referenceBounds[3]) / 2;

        if (selectionReference.bracketDirection) {
            /* 突起と腕で伸び方が違うぶんを戻し、見た目の中央を選択に重ねる / Undo the asymmetry between the point and the arms so the visual center lands on the selection */
            var shapeOffset = (bracketSettings.centerRadiusPt - bracketSettings.endRadiusPt - bracketSettings.extensionPt) / 2;
            return {
                x: referenceCenterX - rotation.cos * shapeOffset,
                y: referenceCenterY - rotation.sin * shapeOffset
            };
        }

        /* 中央の突起の向きへ、外接矩形の半分＋腕の長さ＋余白ぶん寄せると腕の先が辺に触れる / Shifting along the facing direction by half the bounds plus the arm and the gap puts the arm ends on the edge */
        var halfWidth = (referenceBounds[2] - referenceBounds[0]) / 2;
        var halfHeight = (referenceBounds[1] - referenceBounds[3]) / 2;
        var halfSize = Math.abs(rotation.cos) * halfWidth + Math.abs(rotation.sin) * halfHeight;
        var placementOffset = halfSize + bracketSettings.endRadiusPt + bracketSettings.extensionPt + bracketSettings.marginPt;

        return {
            x: referenceCenterX + rotation.cos * placementOffset,
            y: referenceCenterY + rotation.sin * placementOffset
        };
    }

    /**
     * ブラケットのパスを生成して配置する
     * @param {Document} doc - 対象ドキュメント
     * @param {BracketSettings} bracketSettings - ブラケットの設定値
     * @param {SelectionReference|null} selectionReference - 基準の情報（null ならアートボード中央）
     * @returns {PathItem} 生成したパス
     */
    function createBracketPath(doc, bracketSettings, selectionReference) {
        var placementCenter = getPlacementCenter(doc, bracketSettings, selectionReference);
        var bracketPoints = buildBracketPoints(bracketSettings, placementCenter.x, placementCenter.y);

        var bracketPath = doc.pathItems.add();
        bracketPath.closed = false;
        bracketPath.filled = false;
        bracketPath.stroked = true;
        bracketPath.strokeWidth = bracketSettings.strokeWidthPt;
        bracketPath.strokeCap = STROKE_CAPS[bracketSettings.strokeCap] || STROKE_CAPS[DEFAULT_STROKE_CAP];
        bracketPath.strokeJoin = CORNER_JOINS[bracketSettings.cornerJoin] || CORNER_JOINS[DEFAULT_CORNER_JOIN];

        var bracketStrokeColor = new GrayColor();
        bracketStrokeColor.gray = 100;
        bracketPath.strokeColor = bracketStrokeColor;

        for (var i = 0; i < bracketPoints.length; i++) {
            var pathPoint = bracketPath.pathPoints.add();
            pathPoint.anchor = bracketPoints[i].anchor;
            pathPoint.leftDirection = bracketPoints[i].left;
            pathPoint.rightDirection = bracketPoints[i].right;
            pathPoint.pointType = PointType.CORNER;
        }

        /* 面取りは円弧を直線でつなぐジグザグ効果で表現する。形が決まってから適用する / The chamfer is a Zig Zag effect that straightens the arcs; apply it once the shape exists */
        if (bracketSettings.isChamfer) {
            bracketPath.applyEffect(CHAMFER_EFFECT_XML);
        }

        /* 次に選び直したとき同じ位置・向きで描き直せるよう目印を残す / Leave a marker so a later run can redraw it in place */
        var bracketTag = bracketPath.tags.add();
        bracketTag.name = BRACKET_TAG_NAME;
        bracketTag.value = bracketSettings.direction;

        return bracketPath;
    }

    // =========================================
    // セッション記憶 / Session memory
    // =========================================

    /* ダイアログの状態を覚えておく $.global 上のキー。Illustratorを終了すると消える / Key on $.global; it is gone once Illustrator quits */
    var SESSION_SETTINGS_KEY = "aiCurlyBracketMakerSettings";

    /**
     * 前回ダイアログを閉じたときの状態を読み出す
     * @returns {object} 記憶していた値（無ければ空オブジェクト）
     */
    function loadSessionSettings() {
        var savedSettings = $.global[SESSION_SETTINGS_KEY];
        if (!savedSettings) return {};

        /* 常駐オブジェクトを直接書き換えないよう写しを返す / Return a copy so the stored object is not mutated */
        var settings = {};
        for (var key in savedSettings) {
            if (savedSettings.hasOwnProperty(key)) settings[key] = savedSettings[key];
        }
        return settings;
    }

    /**
     * ダイアログの状態を同一セッション内だけ覚える
     * @param {object} settings - 記憶する値
     * @returns {void}
     */
    function saveSessionSettings(settings) {
        $.global[SESSION_SETTINGS_KEY] = settings;
    }

    /**
     * 記憶していた値を数値として読む
     * @param {*} savedValue - 記憶していた値
     * @param {number} fallbackValue - 読めないときに使う値
     * @returns {number} 数値
     */
    function readSavedNumber(savedValue, fallbackValue) {
        var value = Number(savedValue);
        return (savedValue === undefined || savedValue === "" || isNaN(value)) ? fallbackValue : value;
    }

    /**
     * 記憶していた値を入力欄の文字列として読む
     * @param {*} savedValue - 記憶していた値
     * @param {string} fallbackText - 読めないときに使う文字列
     * @returns {string} 入力欄に入れる文字列
     */
    function readSavedText(savedValue, fallbackText) {
        return (savedValue === undefined || savedValue === "") ? fallbackText : String(savedValue);
    }

    /**
     * 記憶していた値をチェックボックスの状態として読む
     * @param {*} savedValue - 記憶していた値
     * @param {boolean} fallbackFlag - 読めないときに使う状態
     * @returns {boolean} チェック状態
     */
    function readSavedFlag(savedValue, fallbackFlag) {
        return (savedValue === undefined) ? fallbackFlag : (savedValue === true);
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ブラケット作成ダイアログを構築する
     * @param {Document} doc - 対象ドキュメント
     * @returns {Window} 構築済みのダイアログ
     */
    function createBracketDialog(doc) {
        /* 現在描画中のプレビューパス / The preview path currently on the artboard */
        var previewPath = null;
        /* ［作成］で確定したか（確定後はプレビューを消さない）/ Whether Create committed the path */
        var bracketCommitted = false;

        /* 選択があれば、その大きさと向きを初期値にして辺に沿わせる / A selection seeds the size and direction, and the bracket hugs its edge */
        var selectionReference = readSelectionReference(doc);
        /* 同一セッション内に覚えていた前回の状態 / What the previous run left behind, within this session */
        var savedSettings = loadSessionSettings();
        var initialValues = getInitialValues(selectionReference, savedSettings);
        /* 参照にしたパスは開いた時点で隠す（確定で削除、キャンセルで元に戻す）/ Hide the reference path as the dialog opens: removed on Create, restored on Cancel */
        setSizeReferenceHidden(selectionReference, true);

        var bracketDialog = new Window("dialog", getLabel('dialog', 'title') + " " + SCRIPT_VERSION);
        setupWindow(bracketDialog);

        var shapePanel = addPanel(bracketDialog, getLabel('panel', 'shapeAndSize'));

        /* 半径は2行を1つの列にまとめ、連動をその右に天地中央で添える / Both radius rows share a column, with the link centered beside them */
        var radiusRow = shapePanel.add("group");
        setupRow(radiusRow, "left", FIELD_ROW_SPACING);

        var radiusColumn = radiusRow.add("group");
        radiusColumn.orientation = "column";
        radiusColumn.alignChildren = ["left", "center"];
        radiusColumn.spacing = PANEL_SPACING;

        var centerRadiusInput = addNumberFieldRow(
            radiusColumn, labelText('fieldLabel', 'centerRadius'),
            formatFieldValue(initialValues.centerRadiusMm), "mm", getLabel('tooltip', 'centerRadius')
        );
        var endRadiusInput = addNumberFieldRow(
            radiusColumn, labelText('fieldLabel', 'endRadius'),
            formatFieldValue(initialValues.endRadiusMm), "mm", getLabel('tooltip', 'endRadius')
        );

        var linkRadiusCheckbox = radiusRow.add("checkbox", undefined, getLabel('checkbox', 'linkRadius'));
        linkRadiusCheckbox.alignment = ["left", "center"];
        linkRadiusCheckbox.value = readSavedFlag(savedSettings.linkRadius, DEFAULT_LINK_RADIUS);
        linkRadiusCheckbox.helpTip = getLabel('tooltip', 'linkRadius');

        /* 面取り：ラベルは空のまま、字下げして数値欄の列にそろえる / Chamfer: an empty label keeps it aligned with the fields */
        var chamferRow = shapePanel.add("group");
        setupRow(chamferRow, "left", FIELD_ROW_SPACING);
        addRowLabel(chamferRow, "");
        var chamferCheckbox = chamferRow.add("checkbox", undefined, getLabel('checkbox', 'chamfer'));
        chamferCheckbox.value = readSavedFlag(savedSettings.chamfer, DEFAULT_CHAMFER);
        chamferCheckbox.helpTip = getLabel('tooltip', 'chamfer');

        /* 向き：ラベル幅を数値欄の行にそろえ、上下・左右の2行に分けて並べる / Direction: share the label width, split into up-down and left-right rows */
        var directionRow = shapePanel.add("group");
        setupRow(directionRow, "left", FIELD_ROW_SPACING);
        /* ラベルを1行目のラジオに合わせる / Align the label with the first row of radios */
        directionRow.alignChildren = ["left", "top"];
        addRowLabel(directionRow, labelText('fieldLabel', 'direction'));

        var directionColumn = directionRow.add("group");
        directionColumn.orientation = "column";
        directionColumn.alignChildren = ["left", "center"];
        directionColumn.spacing = FIELD_ROW_SPACING;

        /* キー順（上・下／左・右）に2つずつ並べる / Two per row, in key order */
        var directionRadios = {};
        for (var i = 0; i < DIRECTION_KEYS.length; i += 2) {
            var directionGridRow = directionColumn.add("group");
            setupRow(directionGridRow, "left", FIELD_ROW_SPACING);
            for (var j = i; j < i + 2 && j < DIRECTION_KEYS.length; j++) {
                var directionKey = DIRECTION_KEYS[j];
                var directionRadio = directionGridRow.add("radiobutton", undefined, getLabel('radio', directionKey));
                directionRadio.helpTip = getLabel('tooltip', 'direction');
                directionRadios[directionKey] = directionRadio;
            }
        }

        var centerOffsetInput = addNumberFieldRow(
            shapePanel, labelText('fieldLabel', 'centerOffset'),
            readSavedText(savedSettings.centerOffset, String(DEFAULT_CENTER_OFFSET_MM)), "mm", getLabel('tooltip', 'centerOffset')
        );

        var totalLengthInput = addNumberFieldRow(
            shapePanel, labelText('fieldLabel', 'totalLength'),
            formatFieldValue(initialValues.totalLengthMm), "mm", getLabel('tooltip', 'totalLength')
        );
        var extensionInput = addNumberFieldRow(
            shapePanel, labelText('fieldLabel', 'extension'),
            readSavedText(savedSettings.extension, String(DEFAULT_EXTENSION_PT)), "pt", getLabel('tooltip', 'extension')
        );
        var marginInput = addNumberFieldRow(
            shapePanel, labelText('fieldLabel', 'margin'),
            readSavedText(savedSettings.margin, String(DEFAULT_MARGIN_MM)), "mm", getLabel('tooltip', 'margin')
        );
        /* 余白が効くのは「囲む対象」があるときだけなので、それ以外は行ごとディム表示 / The gap only applies to something being bracketed, so dim the row otherwise */
        marginInput.parent.enabled = (selectionReference !== null && !selectionReference.bracketDirection);

        var strokePanel = addPanel(bracketDialog, getLabel('panel', 'stroke'));

        var strokeWidthInput = addNumberFieldRow(
            strokePanel, labelText('fieldLabel', 'strokeWidth'),
            readSavedText(savedSettings.strokeWidth, String(DEFAULT_STROKE_WIDTH_PT)), "pt", getLabel('tooltip', 'strokeWidth')
        );

        /* 線端：同じグループ内なのでScriptUIが排他にしてくれる / Stroke cap: one group, so ScriptUI keeps the radios exclusive */
        var strokeCapRow = strokePanel.add("group");
        setupRow(strokeCapRow, "left", FIELD_ROW_SPACING);
        addRowLabel(strokeCapRow, labelText('fieldLabel', 'strokeCap'));
        var buttCapRadio = strokeCapRow.add("radiobutton", undefined, getLabel('radio', 'buttCap'));
        var roundCapRadio = strokeCapRow.add("radiobutton", undefined, getLabel('radio', 'roundCap'));
        var initialStrokeCap = STROKE_CAPS[savedSettings.strokeCap] ? savedSettings.strokeCap : DEFAULT_STROKE_CAP;
        buttCapRadio.value = (initialStrokeCap === "buttCap");
        roundCapRadio.value = !buttCapRadio.value;
        buttCapRadio.helpTip = getLabel('tooltip', 'strokeCap');
        roundCapRadio.helpTip = getLabel('tooltip', 'strokeCap');

        /* 角の形状：ラジオは縦並び。同じグループ内なのでScriptUIが排他にしてくれる / Corner shape: radios stacked in one group, so ScriptUI keeps them exclusive */
        var cornerJoinRow = strokePanel.add("group");
        setupRow(cornerJoinRow, "left", FIELD_ROW_SPACING);
        /* ラベルを1行目のラジオに合わせる（天地中央だと2行ぶんの中央に落ちる）/ Align the label with the first radio instead of centering it over both rows */
        cornerJoinRow.alignChildren = ["left", "top"];
        addRowLabel(cornerJoinRow, labelText('fieldLabel', 'cornerJoin'));

        var cornerJoinColumn = cornerJoinRow.add("group");
        cornerJoinColumn.orientation = "column";
        cornerJoinColumn.alignChildren = ["left", "center"];
        cornerJoinColumn.spacing = FIELD_ROW_SPACING;
        var miterJoinRadio = cornerJoinColumn.add("radiobutton", undefined, getLabel('radio', 'miterJoin'));
        var roundJoinRadio = cornerJoinColumn.add("radiobutton", undefined, getLabel('radio', 'roundJoin'));
        var initialCornerJoin = CORNER_JOINS[savedSettings.cornerJoin] ? savedSettings.cornerJoin : DEFAULT_CORNER_JOIN;
        miterJoinRadio.value = (initialCornerJoin === "miterJoin");
        roundJoinRadio.value = !miterJoinRadio.value;
        miterJoinRadio.helpTip = getLabel('tooltip', 'cornerJoin');
        roundJoinRadio.helpTip = getLabel('tooltip', 'cornerJoin');

        /**
         * 選択中の向きを返す
         * @returns {string} DIRECTION_KEYS のいずれか（未選択なら DEFAULT_DIRECTION）
         */
        function readSelectedDirection() {
            for (var i = 0; i < DIRECTION_KEYS.length; i++) {
                if (directionRadios[DIRECTION_KEYS[i]].value) return DIRECTION_KEYS[i];
            }
            return DEFAULT_DIRECTION;
        }

        /**
         * 入力欄からブラケットの設定値を読み取る
         * @returns {BracketSettings|null} 設定値（数値として読めない欄があれば null）
         */
        function readBracketSettings() {
            var endRadiusMm = Number(endRadiusInput.text);
            var centerRadiusMm = Number(centerRadiusInput.text);
            var totalLengthMm = Number(totalLengthInput.text);
            var extensionPt = Number(extensionInput.text);
            var centerOffsetMm = Number(centerOffsetInput.text);
            var marginMm = Number(marginInput.text);
            var strokeWidthPt = Number(strokeWidthInput.text);

            if (isNaN(endRadiusMm) || endRadiusMm < 0) return null;
            if (isNaN(centerRadiusMm) || centerRadiusMm <= 0) return null;
            if (isNaN(totalLengthMm) || totalLengthMm < 0) return null;
            if (isNaN(extensionPt) || extensionPt < 0) return null;
            if (isNaN(centerOffsetMm)) return null;
            if (isNaN(marginMm) || marginMm < 0) return null;
            if (isNaN(strokeWidthPt) || strokeWidthPt <= 0) return null;

            return {
                endRadiusPt: endRadiusMm * MM_TO_PT,
                centerRadiusPt: centerRadiusMm * MM_TO_PT,
                totalLengthPt: totalLengthMm * MM_TO_PT,
                centerOffsetPt: centerOffsetMm * MM_TO_PT,
                extensionPt: extensionPt,
                marginPt: marginMm * MM_TO_PT,
                strokeWidthPt: strokeWidthPt,
                strokeCap: roundCapRadio.value ? "roundCap" : "buttCap",
                cornerJoin: roundJoinRadio.value ? "roundJoin" : "miterJoin",
                direction: readSelectedDirection(),
                isChamfer: chamferCheckbox.value
            };
        }

        /**
         * プレビューのパスを削除する
         * @returns {void}
         */
        function removePreview() {
            if (!previewPath) return;
            previewPath.remove();
            previewPath = null;
        }

        /**
         * 入力値でプレビューを描き直す
         * @returns {void}
         */
        function drawPreview() {
            var bracketSettings = readBracketSettings();
            /* 入力途中で読めないときは直前のプレビューを残す / Keep the last preview while the input is unreadable */
            if (!bracketSettings) return;

            removePreview();
            previewPath = createBracketPath(doc, bracketSettings, selectionReference);
            app.redraw();
        }

        /**
         * 連動がONのとき、両端の半径を中央に合わせる
         * @returns {void}
         */
        function syncLinkedRadius() {
            if (!linkRadiusCheckbox.value) return;
            if (endRadiusInput.text === centerRadiusInput.text) return;
            endRadiusInput.text = centerRadiusInput.text;
        }

        /**
         * 連動の状態を両端の行に反映する（連動は行の外にあるので操作できるまま残る）
         * @returns {void}
         */
        function updateEndRadiusEnabled() {
            endRadiusInput.parent.enabled = !linkRadiusCheckbox.value;
        }

        /**
         * 数値欄にプレビュー更新（キー入力・↑↓キー）を割り当てる
         * @param {EditText} numberField - 対象の入力欄
         * @param {function} onChanged - 値が変わったときの処理
         * @param {number} [minValue] - ↑↓キーでの下限（省略時は制限なし）
         * @returns {void}
         */
        function bindPreviewField(numberField, onChanged, minValue) {
            numberField.addEventListener("changing", onChanged);
            changeValueByArrowKey(numberField, onChanged, minValue);
        }

        /* 中央は連動の書き写しをはさむ。ほかの数値欄はそのままプレビュー更新 / The center radius syncs first; the other fields just refresh the preview */
        bindPreviewField(centerRadiusInput, function () {
            syncLinkedRadius();
            drawPreview();
        }, 0);
        bindPreviewField(endRadiusInput, drawPreview, 0);

        /* 中央の位置だけはマイナスを許す / Only the center offset may go negative */
        bindPreviewField(centerOffsetInput, drawPreview);

        var previewNumberFields = [totalLengthInput, extensionInput, marginInput, strokeWidthInput];
        for (var i = 0; i < previewNumberFields.length; i++) {
            bindPreviewField(previewNumberFields[i], drawPreview, 0);
        }

        /* 連動をONにした時点で、両端を中央の値にそろえてディム表示にする / Turning the link on snaps the end radius to the center radius and dims it */
        linkRadiusCheckbox.onClick = function () {
            updateEndRadiusEnabled();
            syncLinkedRadius();
            drawPreview();
        };
        /**
         * 線端に角の形状をそろえるクリック処理を作る（丸型→ラウンド結合、なし→マイター結合）
         * @param {boolean} isRoundCap - 丸型を選んだときの処理なら true
         * @returns {function} クリックハンドラ
         */
        function createStrokeCapClickHandler(isRoundCap) {
            return function () {
                roundJoinRadio.value = isRoundCap;
                miterJoinRadio.value = !isRoundCap;
                drawPreview();
            };
        }
        buttCapRadio.onClick = createStrokeCapClickHandler(false);
        roundCapRadio.onClick = createStrokeCapClickHandler(true);
        chamferCheckbox.onClick = drawPreview;
        miterJoinRadio.onClick = drawPreview;
        roundJoinRadio.onClick = drawPreview;
        /**
         * 指定した向きだけを選択状態にする（行をまたぐラジオはScriptUIでは排他にならない）
         * @param {string} selectedKey - 選択する向きのキー
         * @returns {void}
         */
        function selectDirection(selectedKey) {
            for (var i = 0; i < DIRECTION_KEYS.length; i++) {
                directionRadios[DIRECTION_KEYS[i]].value = (DIRECTION_KEYS[i] === selectedKey);
            }
        }

        /**
         * 向きラジオのクリック処理を作る（ループ変数を閉じ込める）
         * @param {string} directionKey - 対象の向きのキー
         * @returns {function} クリックハンドラ
         */
        function createDirectionClickHandler(directionKey) {
            return function () {
                selectDirection(directionKey);
                /* 選択があるときは、新しい向きに合わせて長さを取り直す / With a selection, re-read the length across the new direction */
                if (selectionReference) {
                    totalLengthInput.text = formatFieldValue(getReferenceLengthMm(selectionReference, directionKey));
                }
                drawPreview();
            };
        }
        for (var i = 0; i < DIRECTION_KEYS.length; i++) {
            directionRadios[DIRECTION_KEYS[i]].onClick = createDirectionClickHandler(DIRECTION_KEYS[i]);
        }

        /* ボタンエリア：左側に置くものがないので行ごと右寄せ / Button row: nothing sits on the left, so the row itself is right aligned */
        var btnRowGroup = bracketDialog.add("group");
        setupRow(btnRowGroup, "right", BUTTON_BAR_SPACING);
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        /* キャンセルは既定動作で閉じ、後片付けは bracketDialog.onClose が行う / Cancel closes by default; cleanup happens in bracketDialog.onClose */
        btnRowGroup.add("button", undefined, getLabel('button', 'cancel'), { name: "cancel" });
        var btnCreate = btnRowGroup.add("button", undefined, getLabel('button', 'create'), { name: "ok" });
        btnCreate.helpTip = getLabel('tooltip', 'create');

        /* ［作成］：プレビューをそのまま成果物として残す / Create: keep the preview as the result */
        btnCreate.onClick = function () {
            var bracketSettings = readBracketSettings();
            if (!bracketSettings) {
                alert(getLabel('alert', 'invalidValue'));
                return;
            }
            /* 入力途中で止まったプレビューを最新の値にそろえる / Bring the preview up to date with the current values */
            removePreview();
            previewPath = createBracketPath(doc, bracketSettings, selectionReference);
            /* 隠しておいた参照のパスは役目を終えたので削除 / The hidden reference path has done its job, so remove it */
            removeSizeReference(selectionReference);

            /* 実行後は作成したブラケットだけを選択状態にする / Leave only the new bracket selected */
            doc.selection = null;
            previewPath.selected = true;

            bracketCommitted = true;
            previewPath = null;
            bracketDialog.close();
        };

        /* ダイアログを閉じたら後片付け（キャンセル・ESCも含む）/ Clean up on close (Cancel and ESC included) */
        bracketDialog.onClose = function () {
            if (!bracketCommitted) {
                removePreview();
                setSizeReferenceHidden(selectionReference, false);
                app.redraw();
            }

            /* 閉じたときの状態を、同一セッション内の次回実行のために覚えておく / Remember the closing state for the next run in this session */
            saveSessionSettings({
                centerRadius: centerRadiusInput.text,
                endRadius: endRadiusInput.text,
                linkRadius: linkRadiusCheckbox.value,
                chamfer: chamferCheckbox.value,
                direction: readSelectedDirection(),
                centerOffset: centerOffsetInput.text,
                totalLength: totalLengthInput.text,
                extension: extensionInput.text,
                margin: marginInput.text,
                strokeWidth: strokeWidthInput.text,
                strokeCap: roundCapRadio.value ? "roundCap" : "buttCap",
                cornerJoin: roundJoinRadio.value ? "roundJoin" : "miterJoin"
            });
        };

        selectDirection(initialValues.direction);
        updateEndRadiusEnabled();
        centerRadiusInput.active = true;
        drawPreview(); /* 初回プレビュー / First preview */
        return bracketDialog;
    }

    // =========================================
    // メイン / Main
    // =========================================

    /**
     * ブラケット作成ダイアログを表示する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            app.documents.add();
        }
        var doc = app.activeDocument;

        /* プレビューを描けないレイヤーでは開始しない / Do not start when the preview cannot be drawn */
        var activeLayer = doc.activeLayer;
        if (activeLayer.locked || !activeLayer.visible) {
            alert(getLabel('alert', 'lockedLayer'));
            return;
        }

        createBracketDialog(doc).show();
    }

    main();

})();
