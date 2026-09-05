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

    var DEFAULT_TOTAL_LENGTH_MM = 100;         /* 全体の長さの初期値（mm）/ initial overall length (mm) */
    var DEFAULT_RADIUS_RATIO    = 1 / 15;      /* 半径の初期値＝全体の長さ×この比率 / initial radius = overall length x this ratio */
    var DEFAULT_EXTENSION_PT    = 0;           /* 両端の延長の初期値（pt）/ initial extension at both ends (pt) */
    var DEFAULT_MARGIN_MM       = 0;           /* 選択オブジェクトとの余白の初期値（mm）/ initial gap from the selection (mm) */
    var DEFAULT_STROKE_WIDTH_PT = 2;           /* 線の太さの初期値（pt）/ initial stroke width (pt) */
    var DEFAULT_DIRECTION       = "right";     /* 初期の向き（DIRECTION_KEYS のいずれか）/ initial direction (one of DIRECTION_KEYS) */
    var DEFAULT_STROKE_CAP      = "buttCap";   /* 初期の線端 / initial stroke cap */
    var DEFAULT_CORNER_JOIN     = "miterJoin"; /* 初期の角の形状 / initial corner shape */

    // =========================================
    // レイアウト / Layout
    // =========================================

    var WINDOW_MARGINS     = 16;               /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING     = 12;               /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS      = [16, 20, 16, 12]; /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING      = 8;                /* パネル内の要素間隔 / panel spacing */
    var FIELD_ROW_SPACING  = 6;                /* ラベル・入力欄・単位表記の間隔 / gap inside a labeled row */
    var LABEL_WIDTH        = 100;              /* 行ラベルの共通幅 / shared width of row labels */
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
            parameters: { ja: "パラメータ設定", en: "Parameters" },
            stroke:     { ja: "線", en: "Stroke" }
        },
        fieldLabel: {
            radius:      { ja: "半径", en: "Radius" },
            totalLength: { ja: "直線の長さ", en: "Length" },
            extension:   { ja: "延長", en: "Extension" },
            margin:      { ja: "余白", en: "Margin" },
            strokeWidth: { ja: "線の太さ", en: "Stroke Width" },
            strokeCap:   { ja: "線端", en: "Cap" },
            cornerJoin:  { ja: "角の形状", en: "Corner Shape" },
            direction:   { ja: "向き", en: "Direction" }
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
            radius:      { ja: "先端と両端を作る1/4円の半径。↑↓で増減、Shift+↑↓で10単位スナップ", en: "Radius of the quarter circles at the tip and the ends. Up/Down to step, Shift+Up/Down snaps to 10" },
            totalLength: { ja: "ブラケット全体の長さ。半径を変えても総長は変わらず、直線部分が自動で伸縮します", en: "Overall length of the bracket. Changing the radius keeps this length and resizes the straight sections instead" },
            extension:   { ja: "両端から先端の反対側へ伸ばす直線の長さ（0で延長なし）", en: "Straight run added at both ends, away from the tip (0 adds none)" },
            margin:      { ja: "選択オブジェクトとブラケットのあいだの間隔（選択して実行したときのみ有効）", en: "Gap between the selection and the bracket (only when run with a selection)" },
            strokeWidth: { ja: "ブラケットの線幅（pt）", en: "Stroke width of the bracket (pt)" },
            strokeCap:   { ja: "両端の線の先を丸めるかどうか", en: "Whether both ends of the stroke are rounded" },
            cornerJoin:  { ja: "先端の角を尖らせるか丸めるか", en: "Whether the tip corner is pointed or rounded" },
            direction:   { ja: "先端を向ける方向", en: "The direction the tip points to" }
        },
        button: {
            create: { ja: "作成", en: "Create" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            lockedLayer:  { ja: "アクティブレイヤーがロックまたは非表示です。", en: "The active layer is locked or hidden." },
            invalidValue: { ja: "半径・長さ・線の太さには数値を入力してください（半径と線の太さは0より大きい値）。", en: "Enter numbers for the radius, length and stroke width (radius and stroke width must be greater than 0)." }
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
     * @property {number} radiusPt - 1/4円の半径（pt）
     * @property {number} totalLengthPt - 全体の長さ（pt）
     * @property {number} extensionPt - 両端から先端の反対側へ伸ばす長さ（pt）
     * @property {number} marginPt - 選択オブジェクトとの余白（pt）
     * @property {number} strokeWidthPt - 線の太さ（pt）
     * @property {string} strokeCap - 線端（STROKE_CAPS のキー）
     * @property {string} cornerJoin - 角の形状（CORNER_JOINS のキー）
     * @property {string} direction - 向き（DIRECTION_KEYS のいずれか）
     */

    /**
     * ブラケットのアンカーポイントと方向線を計算する（右向きで組み立ててから向きに応じて回転）
     * @param {BracketSettings} bracketSettings - ブラケットの設定値
     * @param {number} centerX - 配置位置の中心X（pt）
     * @param {number} centerY - 配置位置の中心Y（pt）
     * @returns {Array<{anchor: number[], left: number[], right: number[]}>} 向きを反映したパスポイント
     */
    function buildBracketPoints(bracketSettings, centerX, centerY) {
        var radius = bracketSettings.radiusPt;
        var extension = bracketSettings.extensionPt;
        /* 全体の長さから円弧ぶん（片側 2×半径）を差し引いた直線部分 / The straight run left after the arcs (2 x radius per side) */
        var straightLength = Math.max(0, bracketSettings.totalLengthPt / 2 - 2 * radius);
        var handleLength = radius * KAPPA;
        var rotation = DIRECTION_ROTATIONS[bracketSettings.direction] || DIRECTION_ROTATIONS[DEFAULT_DIRECTION];
        var bracketPoints = [];

        /* 0. 上の延長線（先端の反対側へ伸ばす）/ Upper extension, running away from the tip */
        if (extension > 0) {
            bracketPoints.push(createStraightPoint(centerX - radius - extension, centerY + 2 * radius + straightLength));
        }

        /* 1. 上端点（左上円の上端）/ Top end (top of the upper-left arc) */
        bracketPoints.push({
            anchor: [centerX - radius, centerY + 2 * radius + straightLength],
            left:   [centerX - radius, centerY + 2 * radius + straightLength],
            right:  [centerX - (radius - handleLength), centerY + 2 * radius + straightLength]
        });

        if (straightLength > 0) {
            /* 2. 左上円の右端（直線の上開始点）/ Start of the upper straight section */
            bracketPoints.push({
                anchor: [centerX, centerY + radius + straightLength],
                left:   [centerX, centerY + radius + straightLength + handleLength],
                right:  [centerX, centerY + radius + straightLength]
            });

            /* 3. 中央上円の左端（直線の終了点）/ End of the upper straight section */
            bracketPoints.push({
                anchor: [centerX, centerY + radius],
                left:   [centerX, centerY + radius],
                right:  [centerX, centerY + radius - handleLength]
            });
        } else {
            /* 直線がない場合（変曲点）/ Inflection point when there is no straight section */
            bracketPoints.push({
                anchor: [centerX, centerY + radius],
                left:   [centerX, centerY + radius + handleLength],
                right:  [centerX, centerY + radius - handleLength]
            });
        }

        /* 4. 先端（中央の鋭角部分）/ Tip (the pointed center) */
        bracketPoints.push({
            anchor: [centerX + radius, centerY],
            left:   [centerX + (radius - handleLength), centerY],
            right:  [centerX + (radius - handleLength), centerY]
        });

        if (straightLength > 0) {
            /* 5. 中央下円の左端（直線の下開始点）/ Start of the lower straight section */
            bracketPoints.push({
                anchor: [centerX, centerY - radius],
                left:   [centerX, centerY - radius + handleLength],
                right:  [centerX, centerY - radius]
            });

            /* 6. 左下円の右端（直線の終了点）/ End of the lower straight section */
            bracketPoints.push({
                anchor: [centerX, centerY - radius - straightLength],
                left:   [centerX, centerY - radius - straightLength],
                right:  [centerX, centerY - radius - straightLength - handleLength]
            });
        } else {
            /* 直線がない場合（変曲点）/ Inflection point when there is no straight section */
            bracketPoints.push({
                anchor: [centerX, centerY - radius],
                left:   [centerX, centerY - radius + handleLength],
                right:  [centerX, centerY - radius - handleLength]
            });
        }

        /* 7. 下端点（左下円の下端）/ Bottom end (bottom of the lower-left arc) */
        bracketPoints.push({
            anchor: [centerX - radius, centerY - (2 * radius + straightLength)],
            left:   [centerX - (radius - handleLength), centerY - (2 * radius + straightLength)],
            right:  [centerX - radius, centerY - (2 * radius + straightLength)]
        });

        /* 8. 下の延長線（先端の反対側へ伸ばす）/ Lower extension, running away from the tip */
        if (extension > 0) {
            bracketPoints.push(createStraightPoint(centerX - radius - extension, centerY - (2 * radius + straightLength)));
        }

        /* 右向きで組み立てた座標を、選んだ向きへまとめて回転 / Rotate the right-facing points into the chosen direction */
        for (var i = 0; i < bracketPoints.length; i++) {
            bracketPoints[i].anchor = rotateAroundCenter(bracketPoints[i].anchor, rotation, centerX, centerY);
            bracketPoints[i].left = rotateAroundCenter(bracketPoints[i].left, rotation, centerX, centerY);
            bracketPoints[i].right = rotateAroundCenter(bracketPoints[i].right, rotation, centerX, centerY);
        }
        return bracketPoints;
    }

    /**
     * @typedef {object} SelectionReference
     * @property {number[]} bounds - 選択全体の外接矩形 [左, 上, 右, 下]
     * @property {PageItem[]} items - 基準にしたオブジェクト
     * @property {boolean} isSizeReference - 大きさの参照（開いたパス1本）なら true。確定時に削除する
     */

    /**
     * 選択オブジェクトを基準情報として読み取る
     * @param {Document} doc - 対象ドキュメント
     * @returns {SelectionReference|null} 基準の情報（境界を持つ選択がなければ null）
     */
    function readSelectionReference(doc) {
        var selectedItems = doc.selection;
        if (!selectedItems || selectedItems.length === 0) return null;

        var selectionBounds = null;
        var referenceItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            var itemBounds = null;
            try {
                itemBounds = selectedItems[i].geometricBounds;
            } catch (eBounds) {
                itemBounds = null; /* 文字選択など境界を持たないものは飛ばす / Skip selections without bounds, such as a text range */
            }
            if (!itemBounds) continue;

            referenceItems.push(selectedItems[i]);
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
        return {
            bounds: selectionBounds,
            items: referenceItems,
            isSizeReference: isSizeReference(referenceItems)
        };
    }

    /**
     * 大きさの参照として扱う選択かどうかを返す（開いたパス1本だけなら参照）
     * @param {PageItem[]} referenceItems - 対象のオブジェクト
     * @returns {boolean} 参照として扱うなら true
     */
    function isSizeReference(referenceItems) {
        if (referenceItems.length !== 1) return false;
        var referenceItem = referenceItems[0];
        /* 閉じたパスや図形・テキストは「囲む対象」として残す / Closed paths, shapes and text stay, as things to be bracketed */
        return (referenceItem.typename === "PathItem" && !referenceItem.closed);
    }

    /**
     * 大きさの参照にしたパスの表示・非表示を切り替える（囲む対象のときは何もしない）
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @param {boolean} isHidden - 隠すなら true、戻すなら false
     * @returns {void}
     */
    function setReferenceItemsHidden(selectionReference, isHidden) {
        if (!selectionReference || !selectionReference.isSizeReference) return;
        for (var i = 0; i < selectionReference.items.length; i++) {
            try {
                selectionReference.items[i].hidden = isHidden;
            } catch (eHideReference) {
                /* ロックなどで切り替えられないものはそのまま / Leave alone what cannot be toggled, such as a locked item */
            }
        }
    }

    /**
     * 大きさの参照にしたパスを削除する（囲む対象のときは何もしない）
     * @param {SelectionReference|null} selectionReference - 基準の情報
     * @returns {void}
     */
    function removeReferenceItems(selectionReference) {
        if (!selectionReference || !selectionReference.isSizeReference) return;
        for (var i = 0; i < selectionReference.items.length; i++) {
            try {
                selectionReference.items[i].remove();
            } catch (eRemoveReference) {
                /* ロックなどで消せないものは残す / Leave behind what cannot be removed, such as a locked item */
            }
        }
    }

    /**
     * 基準矩形から初期値を求める
     * 参照のパスは左（縦長）・上（横長）に、囲む対象は右（縦長）・下（横長）に置く
     * @param {number[]|null} referenceBounds - 基準の外接矩形 [左, 上, 右, 下]
     * @param {boolean} isSizeReferenceSelection - 大きさの参照なら true
     * @returns {{totalLengthMm: number, radiusMm: number, direction: string}} 全体の長さ・半径（mm）と向き
     */
    function getInitialValues(referenceBounds, isSizeReferenceSelection) {
        var totalLengthMm = DEFAULT_TOTAL_LENGTH_MM;
        var direction = DEFAULT_DIRECTION;

        if (referenceBounds) {
            var referenceWidth = referenceBounds[2] - referenceBounds[0];
            var referenceHeight = referenceBounds[1] - referenceBounds[3];
            var isVertical = (referenceHeight >= referenceWidth);
            totalLengthMm = (isVertical ? referenceHeight : referenceWidth) / MM_TO_PT;
            if (isVertical) {
                direction = isSizeReferenceSelection ? "left" : "right";
            } else {
                direction = isSizeReferenceSelection ? "up" : "down";
            }
        }

        return {
            totalLengthMm: totalLengthMm,
            /* 半径は全体の長さに追従させる（以後は個別に変更できる）/ The radius follows the overall length, and can be changed on its own afterwards */
            radiusMm: totalLengthMm * DEFAULT_RADIUS_RATIO,
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
     * ブラケットを置く中心座標を返す（基準矩形があれば、腕の先がその辺に触れる位置）
     * @param {Document} doc - 対象ドキュメント
     * @param {BracketSettings} bracketSettings - ブラケットの設定値
     * @param {number[]|null} referenceBounds - 基準の外接矩形 [左, 上, 右, 下]
     * @returns {{x: number, y: number}} 中心座標（pt）
     */
    function getPlacementCenter(doc, bracketSettings, referenceBounds) {
        if (!referenceBounds) return getActiveArtboardCenter(doc);

        /* 中心から腕の先までの距離（先端の反対側）に余白を足す / Distance from the center to the arm ends, plus the gap */
        var armOffset = bracketSettings.radiusPt + bracketSettings.extensionPt + bracketSettings.marginPt;
        var referenceCenterX = (referenceBounds[0] + referenceBounds[2]) / 2;
        var referenceCenterY = (referenceBounds[1] + referenceBounds[3]) / 2;

        if (bracketSettings.direction === "left") {
            return { x: referenceBounds[0] - armOffset, y: referenceCenterY };
        }
        if (bracketSettings.direction === "right") {
            return { x: referenceBounds[2] + armOffset, y: referenceCenterY };
        }
        if (bracketSettings.direction === "up") {
            return { x: referenceCenterX, y: referenceBounds[1] + armOffset };
        }
        return { x: referenceCenterX, y: referenceBounds[3] - armOffset };
    }

    /**
     * ブラケットのパスを生成して配置する
     * @param {Document} doc - 対象ドキュメント
     * @param {BracketSettings} bracketSettings - ブラケットの設定値
     * @param {number[]|null} referenceBounds - 基準の外接矩形（null ならアートボード中央）
     * @returns {PathItem} 生成したパス
     */
    function createBracketPath(doc, bracketSettings, referenceBounds) {
        var placementCenter = getPlacementCenter(doc, bracketSettings, referenceBounds);
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
        return bracketPath;
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
        var referenceBounds = selectionReference ? selectionReference.bounds : null;
        var initialValues = getInitialValues(referenceBounds, selectionReference ? selectionReference.isSizeReference : false);
        /* 参照にしたパスは開いた時点で隠す（確定で削除、キャンセルで元に戻す）/ Hide the reference path as the dialog opens: removed on Create, restored on Cancel */
        setReferenceItemsHidden(selectionReference, true);

        var bracketDialog = new Window("dialog", getLabel('dialog', 'title') + " " + SCRIPT_VERSION);
        setupWindow(bracketDialog);

        var parameterPanel = addPanel(bracketDialog, getLabel('panel', 'parameters'));

        var radiusInput = addNumberFieldRow(
            parameterPanel, labelText('fieldLabel', 'radius'),
            formatFieldValue(initialValues.radiusMm), "mm", getLabel('tooltip', 'radius')
        );

        /* 向き：ラベル幅を数値欄の行にそろえ、キー順（上・下・左・右）に1行で並べる / Direction: share the label width, one row in key order */
        var directionRow = parameterPanel.add("group");
        setupRow(directionRow, "left", FIELD_ROW_SPACING);
        addRowLabel(directionRow, labelText('fieldLabel', 'direction'));

        var directionRadios = {};
        for (var i = 0; i < DIRECTION_KEYS.length; i++) {
            var directionKey = DIRECTION_KEYS[i];
            var directionRadio = directionRow.add("radiobutton", undefined, getLabel('radio', directionKey));
            directionRadio.helpTip = getLabel('tooltip', 'direction');
            directionRadio.value = (directionKey === initialValues.direction);
            directionRadios[directionKey] = directionRadio;
        }

        var totalLengthInput = addNumberFieldRow(
            parameterPanel, labelText('fieldLabel', 'totalLength'),
            formatFieldValue(initialValues.totalLengthMm), "mm", getLabel('tooltip', 'totalLength')
        );
        var extensionInput = addNumberFieldRow(
            parameterPanel, labelText('fieldLabel', 'extension'),
            String(DEFAULT_EXTENSION_PT), "pt", getLabel('tooltip', 'extension')
        );
        var marginInput = addNumberFieldRow(
            parameterPanel, labelText('fieldLabel', 'margin'),
            String(DEFAULT_MARGIN_MM), "mm", getLabel('tooltip', 'margin')
        );
        /* 対象オブジェクトがないときは余白が効かないので行ごとディム表示 / The gap does nothing without a selection, so dim the whole row */
        marginInput.parent.enabled = (referenceBounds !== null);

        var strokePanel = addPanel(bracketDialog, getLabel('panel', 'stroke'));

        var strokeWidthInput = addNumberFieldRow(
            strokePanel, labelText('fieldLabel', 'strokeWidth'),
            String(DEFAULT_STROKE_WIDTH_PT), "pt", getLabel('tooltip', 'strokeWidth')
        );

        /* 線端：同じグループ内なのでScriptUIが排他にしてくれる / Stroke cap: one group, so ScriptUI keeps the radios exclusive */
        var strokeCapRow = strokePanel.add("group");
        setupRow(strokeCapRow, "left", FIELD_ROW_SPACING);
        addRowLabel(strokeCapRow, labelText('fieldLabel', 'strokeCap'));
        var buttCapRadio = strokeCapRow.add("radiobutton", undefined, getLabel('radio', 'buttCap'));
        var roundCapRadio = strokeCapRow.add("radiobutton", undefined, getLabel('radio', 'roundCap'));
        buttCapRadio.value = (DEFAULT_STROKE_CAP === "buttCap");
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
        miterJoinRadio.value = (DEFAULT_CORNER_JOIN === "miterJoin");
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
            var radiusMm = Number(radiusInput.text);
            var totalLengthMm = Number(totalLengthInput.text);
            var extensionPt = Number(extensionInput.text);
            var marginMm = Number(marginInput.text);
            var strokeWidthPt = Number(strokeWidthInput.text);

            if (isNaN(radiusMm) || radiusMm <= 0) return null;
            if (isNaN(totalLengthMm) || totalLengthMm < 0) return null;
            if (isNaN(extensionPt) || extensionPt < 0) return null;
            if (isNaN(marginMm) || marginMm < 0) return null;
            if (isNaN(strokeWidthPt) || strokeWidthPt <= 0) return null;

            return {
                radiusPt: radiusMm * MM_TO_PT,
                totalLengthPt: totalLengthMm * MM_TO_PT,
                extensionPt: extensionPt,
                marginPt: marginMm * MM_TO_PT,
                strokeWidthPt: strokeWidthPt,
                strokeCap: roundCapRadio.value ? "roundCap" : "buttCap",
                cornerJoin: roundJoinRadio.value ? "roundJoin" : "miterJoin",
                direction: readSelectedDirection()
            };
        }

        /**
         * プレビューのパスを削除する
         * @returns {void}
         */
        function removePreview() {
            if (!previewPath) return;
            try {
                previewPath.remove();
            } catch (eRemovePreview) {
                /* すでに失われている場合は何もしない / Nothing to do when it is already gone */
            }
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
            previewPath = createBracketPath(doc, bracketSettings, referenceBounds);
            app.redraw();
        }

        /* 数値欄はキー入力とキー増減の両方でプレビューを更新 / Numeric fields refresh the preview on typing and on arrow keys */
        var previewNumberFields = [radiusInput, totalLengthInput, extensionInput, marginInput, strokeWidthInput];
        for (var i = 0; i < previewNumberFields.length; i++) {
            previewNumberFields[i].addEventListener("changing", drawPreview);
            changeValueByArrowKey(previewNumberFields[i], drawPreview, 0);
        }
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
        miterJoinRadio.onClick = drawPreview;
        roundJoinRadio.onClick = drawPreview;
        for (var i = 0; i < DIRECTION_KEYS.length; i++) {
            directionRadios[DIRECTION_KEYS[i]].onClick = drawPreview;
        }

        /* ボタンエリア（右寄せ）/ Button row (right aligned) */
        var btnRowGroup = bracketDialog.add("group");
        setupRow(btnRowGroup, "fill", BUTTON_BAR_SPACING);
        btnRowGroup.margins = BUTTON_BAR_MARGINS;

        /* スペーサー（伸縮）/ Spacer (stretchable) */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        setupRow(btnRightGroup, "right", BUTTON_BAR_SPACING);
        /* キャンセルは既定動作で閉じ、後片付けは bracketDialog.onClose が行う / Cancel closes by default; cleanup happens in bracketDialog.onClose */
        btnRightGroup.add("button", undefined, getLabel('button', 'cancel'), { name: "cancel" });
        var btnCreate = btnRightGroup.add("button", undefined, getLabel('button', 'create'), { name: "ok" });

        /* ［作成］：プレビューをそのまま成果物として残す / Create: keep the preview as the result */
        btnCreate.onClick = function () {
            var bracketSettings = readBracketSettings();
            if (!bracketSettings) {
                alert(getLabel('alert', 'invalidValue'));
                return;
            }
            /* 入力途中で止まったプレビューを最新の値にそろえる / Bring the preview up to date with the current values */
            removePreview();
            previewPath = createBracketPath(doc, bracketSettings, referenceBounds);
            /* 隠しておいた参照のパスは役目を終えたので削除 / The hidden reference path has done its job, so remove it */
            removeReferenceItems(selectionReference);
            previewPath.selected = true;

            bracketCommitted = true;
            previewPath = null;
            bracketDialog.close();
        };

        /* ダイアログを閉じたら後片付け（キャンセル・ESCも含む）/ Clean up on close (Cancel and ESC included) */
        bracketDialog.onClose = function () {
            if (!bracketCommitted) {
                removePreview();
                setReferenceItemsHidden(selectionReference, false);
                app.redraw();
            }
        };

        radiusInput.active = true;
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
