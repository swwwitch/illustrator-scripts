#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択した長方形パスから、ミシン目・ギザギザ・コーナー処理・スリット／ホールを組み合わせたチケット風の形状を生成します。
専用のプレビューレイヤーで結果を確認しながら設定でき、［OK］したときだけ元のオブジェクトへ適用します。

詳細は README を参照してください。

### Overview

Turns a selected rectangle into a ticket-like shape combining perforations, zigzag edges, corner treatments and slits or holes.
The result is set up on a dedicated preview layer and applied to the original only when you confirm with OK.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CouponTicketMaker";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.4.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-03-08";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-08-13";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/CouponTicketMaker.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/CouponTicketMaker.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n2e949946228a"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* 塗りも線も持たないオブジェクトに使う代替色の濃度（%）/ Tint used when an object has neither fill nor stroke */
var FALLBACK_GRAY_TINT = 60;

/* 逆角丸で四隅だけを残すための破線間隔（pt）/ Dash gap that leaves only the corners for inverse rounding */
var INVERSE_CORNER_DASH_GAP = 1000;

/* 長方形判定に使う座標の許容誤差（pt）/ Coordinate tolerance for the rectangle check */
var RECTANGLE_TOLERANCE = 0.01;

// =========================================
// レイアウト / Layout
// =========================================

var PANEL_MARGINS        = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] */
var PANEL_SPACING        = 6;                  /* パネル内の要素間隔 */
var PRESET_ROW_MARGINS   = [20, 0, 20, 0];     /* プリセット行の余白 */
var PRESET_ROW_SPACING   = 12;                 /* プリセット行の要素間隔 */
var SAVE_BUTTON_SIZE     = [70, 24];           /* ［保存］ボタンのサイズ */
var OFFSET_SLIDER_WIDTH  = 200;                /* 分割位置スライダーの幅 */
var NUMBER_FIELD_CHARS   = 3;                  /* 数値入力欄の標準幅（文字数） */
var OFFSET_FIELD_CHARS   = 5;                  /* 分割位置入力欄の幅（文字数） */
var INSET_FIELD_CHARS    = 4;                  /* 負値を入れる入力欄の幅（文字数） */
var ZIGZAG_LABEL_WIDTH   = { ja: 60, en: 55 }; /* ギザギザパネルのラベル幅 */

/**
 * パネルへ共通のレイアウトを適用する
 * @param {Panel} panel - 対象パネル
 * @param {Array<string>} alignChildren - 子要素の整列指定（省略時は ["fill", "top"]）
 * @returns {void}
 */
function setupPanel(panel, alignChildren) {
    panel.orientation = 'column';
    panel.margins = PANEL_MARGINS;
    panel.spacing = PANEL_SPACING;
    panel.alignChildren = alignChildren || ['fill', 'top'];
}

/**
 * ラベル付きパネルを生成し、共通レイアウトを適用する
 * @param {Window|Panel|Group} parent - 追加先のコンテナ
 * @param {string} titleText - パネルのタイトル
 * @param {Array<string>} alignChildren - 子要素の整列指定（省略時は ["fill", "top"]）
 * @returns {Panel} 生成したパネル
 */
function addPanel(parent, titleText, alignChildren) {
    var panel = parent.add('panel', undefined, titleText);
    setupPanel(panel, alignChildren);
    return panel;
}

/**
 * 横並びのグループを生成する
 * @param {Window|Panel|Group} parent - 追加先のコンテナ
 * @returns {Group} 生成したグループ
 */
function addRow(parent) {
    var row = parent.add('group');
    row.orientation = 'row';
    row.alignChildren = ['left', 'center'];
    return row;
}

/**
 * 縦積みのグループを生成する
 * @param {Window|Panel|Group} parent - 追加先のコンテナ
 * @param {Array<string>} alignChildren - 子要素の整列指定（省略時は ["fill", "top"]）
 * @returns {Group} 生成したグループ
 */
function addColumn(parent, alignChildren) {
    var column = parent.add('group');
    column.orientation = 'column';
    column.alignChildren = alignChildren || ['fill', 'top'];
    return column;
}

/**
 * 「ラベル＋数値入力欄＋単位」の1行を生成する
 * @param {Panel|Group} parent - 追加先のコンテナ
 * @param {Object} options - 生成オプション
 * @param {string} options.label - ラベル文言
 * @param {string} options.value - 入力欄の初期値
 * @param {string} options.unit - 単位表記（不要なら省略）
 * @param {number} options.labelWidth - ラベル幅（省略時は指定しない）
 * @param {number} options.characters - 入力欄の幅（文字数、省略時は NUMBER_FIELD_CHARS）
 * @param {boolean} options.allowNegative - 負値を許可するか
 * @param {boolean} options.negativeOnly - 0以下だけを許可するか
 * @param {function} options.onChange - 値が変わったときに呼ぶ関数
 * @returns {Object} { row: Group, input: EditText }
 */
function addNumberField(parent, options) {
    var row = addRow(parent);
    var label = row.add('statictext', undefined, options.label);
    if (options.labelWidth) label.preferredSize.width = options.labelWidth;

    var input = row.add('edittext', undefined, options.value);
    input.characters = options.characters || NUMBER_FIELD_CHARS;
    bindArrowKeyStep(input, options);

    if (options.unit) row.add('statictext', undefined, options.unit);

    return { row: row, input: input };
}

/**
 * ↑↓キーで数値を増減できるようにする（Shiftで10単位、Optionで0.1単位）
 * @param {EditText} input - 対象の入力欄
 * @param {Object} options - 挙動オプション（allowNegative / negativeOnly / onChange）
 * @returns {void}
 */
function bindArrowKeyStep(input, options) {
    input.addEventListener('keydown', function (event) {
        var value = Number(input.text);
        if (isNaN(value)) return;
        if (event.keyName != 'Up' && event.keyName != 'Down') return;

        var keyboard = ScriptUI.environment.keyboardState;
        var isUp = (event.keyName == 'Up');
        event.preventDefault();

        if (keyboard.shiftKey) {
            /* 10単位でスナップ / Snap to 10 */
            value = isUp ? Math.ceil((value + 1) / 10) * 10 : Math.floor((value - 1) / 10) * 10;
        } else if (keyboard.altKey) {
            /* 0.1単位で増減 / Step by 0.1 */
            value = Math.round((value + (isUp ? 0.1 : -0.1)) * 10) / 10;
        } else {
            value = Math.round(isUp ? value + 1 : value - 1);
        }

        if (!options.allowNegative && value < 0) value = 0;
        if (options.negativeOnly && value > 0) value = 0;

        input.text = value;
        if (options.onChange) options.onChange();
    });
}

// =========================================
// ローカライズ / Localization
// =========================================

/**
 * 現在のロケールから表示言語を判定する
 * @returns {string} "ja" または "en"
 */
function getCurrentLang() {
    return (($.locale || '') + '').indexOf('ja') === 0 ? 'ja' : 'en';
}
var uiLang = getCurrentLang();

/* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
var LABELS = {
    dialog: {
        title: { ja: "チケットメーカー", en: "Ticket Maker" }
    },
    panel: {
        sides:           { ja: "左右", en: "L/R" },
        sidePerforation: { ja: "ミシン目", en: "Perforation" },
        zigzag:          { ja: "ギザギザ", en: "Zigzag" },
        centerSplit:     { ja: "左右分割", en: "Center Divider" },
        divider:         { ja: "分割線", en: "Divider Line" },
        edge:            { ja: "エッジ", en: "Edge" },
        corner:          { ja: "コーナー", en: "Corner" },
        hole:            { ja: "スリット／ホール", en: "Slit / Hole" }
    },
    radio: {
        none:      { ja: "なし", en: "None" },
        dot:       { ja: "ドット", en: "Dot" },
        dash:      { ja: "破線", en: "Dash" },
        circle:    { ja: "円", en: "Circle" },
        triangle:  { ja: "三角", en: "Triangle" },
        round:     { ja: "角丸", en: "Rounded" },
        inverse:   { ja: "逆角丸", en: "Inverse Round" },
        chamfer:   { ja: "面取り", en: "Chamfer" },
        leftRight: { ja: "左右", en: "L/R" },
        topBottom: { ja: "上下", en: "T/B" }
    },
    checkbox: {
        enable:           { ja: "有効", en: "Enable" },
        linkToCenter:     { ja: "分割線に連動", en: "Link to Split Line" },
        left:             { ja: "左", en: "Left" },
        right:            { ja: "右", en: "Right" },
        doubleRound:      { ja: "ダブル角丸", en: "Double Rounded" },
        edgesOnly:        { ja: "エッジのみ", en: "Edges Only" },
        preview:          { ja: "プレビュー", en: "Preview" },
        expandAppearance: { ja: "アピアランスを分割", en: "Expand Appearance" }
    },
    fieldLabel: {
        lineWidth:    { ja: "線幅:", en: "Weight:" },
        gap:          { ja: "間隔:", en: "Gap:" },
        inset:        { ja: "長さ:", en: "Inset Length:" },
        size:         { ja: "サイズ:", en: "Size:" },
        zigzagSize:   { ja: "大きさ:", en: "Size:" },
        zigzagRepeat: { ja: "繰り返し:", en: "Repeat:" }
    },
    button: {
        save:       { ja: "保存", en: "Save" },
        cancel:     { ja: "キャンセル", en: "Cancel" },
        ok:         { ja: "OK", en: "OK" },
        outlineOn:  { ja: "アウトライン表示", en: "Outline View" },
        outlineOff: { ja: "プレビュー表示", en: "Preview View" }
    },
    alert: {
        openDocument:      { ja: "ドキュメントを開いてください。", en: "Please open a document." },
        selectRectangle:   { ja: "長方形を選択してください。", en: "Please select a rectangle." },
        rectangleOnly:     { ja: "長方形のパスを1つだけ選択してください。", en: "Please select exactly one rectangular path." },
        presetName:        { ja: "プリセット名を入力してください。", en: "Enter a preset name." },
        singleOnly: {
            ja: "複数選択時は実行できません。オブジェクトを1つだけ選択してください。",
            en: "This script cannot run with multiple selections. Please select only one object."
        },
        groupNotAllowed: {
            ja: "グループを選択しているときは実行できません。単体のオブジェクトを選択してください。",
            en: "This script cannot run when a group is selected. Please select a single object."
        },
        enterValidNumbers: {
            ja: "数値欄に正しい数値を入力してください。",
            en: "Please enter valid numeric values in the numeric fields."
        },
        presetSaved: {
            ja: "プリセットを保存しました。\n\nコード組み込み用:\n",
            en: "Preset saved.\n\nCode snippet for embedding:\n"
        }
    },
    layerName: {
        preview: { ja: "プレビュー", en: "Preview" }
    },
    fallbackName: {
        customPreset: { ja: "カスタム", en: "Custom" }
    }
};

/**
 * ラベルノードから現在の言語の文言を取り出す
 * @param {Object} labelNode - LABELS 内の { ja, en } ノード
 * @returns {string} 表示する文言
 */
function getLabel(labelNode) {
    if (!labelNode) return '';
    return labelNode[uiLang] || labelNode.ja || labelNode.en || '';
}

// =========================================
// 単位ユーティリティ / Unit utilities
// =========================================

/* 環境設定の単位コードと表記の対応 / Unit code to label */
var UNIT_LABELS = {
    0: "in",
    1: "mm",
    2: "pt",
    3: "pica",
    4: "cm",
    6: "px",
    7: "ft/in",
    8: "m",
    9: "yd",
    10: "ft"
};

/* コード5（歯／級）でH表記になる環境設定キー / Preference keys that use "H" for unit code 5 */
var HA_UNIT_PREF_KEYS = {
    "text/asianunits": true,
    "rulerType": true,
    "strokeUnits": true
};

/**
 * 単位コードから表記文字列を返す
 * @param {number} unitCode - 環境設定の単位コード
 * @param {string} prefKey - 参照した環境設定キー
 * @returns {string} 単位表記
 */
function getUnitLabel(unitCode, prefKey) {
    if (unitCode === 5) return HA_UNIT_PREF_KEYS[prefKey] ? 'H' : 'Q';
    return UNIT_LABELS[unitCode] || 'pt';
}

/**
 * 単位コードからポイント換算の係数を返す
 * @param {number} unitCode - 環境設定の単位コード
 * @returns {number} 1単位あたりのポイント数
 */
function getPtFactorFromUnitCode(unitCode) {
    switch (unitCode) {
        case 0: return 72.0;
        case 1: return 72.0 / 25.4;
        case 2: return 1.0;
        case 3: return 12.0;
        case 4: return 72.0 / 2.54;
        case 5: return 72.0 / 25.4 * 0.25;
        case 6: return 1.0;
        case 7: return 72.0 * 12.0;
        case 8: return 72.0 / 25.4 * 1000.0;
        case 9: return 72.0 * 36.0;
        case 10: return 72.0 * 12.0;
        default: return 1.0;
    }
}

/**
 * 指定単位の値をポイントへ変換する
 * @param {number} value - 変換前の値
 * @param {number} factor - 1単位あたりのポイント数
 * @returns {number} ポイント値
 */
function toPt(value, factor) {
    return value * factor;
}

/**
 * ポイント値を指定単位へ変換する
 * @param {number} value - ポイント値
 * @param {number} factor - 1単位あたりのポイント数
 * @returns {number} 変換後の値
 */
function fromPt(value, factor) {
    return value / factor;
}

var rulerUnitCode = app.preferences.getIntegerPreference('rulerType');
var rulerUnitLabel = getUnitLabel(rulerUnitCode, 'rulerType');
var rulerPtFactor = getPtFactorFromUnitCode(rulerUnitCode);

/* 分割線の線幅だけは線の単位に従う / Stroke units apply to the divider line weight only */
var strokeUnitCode = app.preferences.getIntegerPreference('strokeUnits');
var strokeUnitLabel = getUnitLabel(strokeUnitCode, 'strokeUnits');
var strokePtFactor = getPtFactorFromUnitCode(strokeUnitCode);

// =========================================
// 一時アクション / Temporary action
// =========================================

/* 破線を「線の位置：中央」で描くための一時アクション定義 / Action data that draws dashes with centered alignment */
var STROKE_DOT_ACTION = '/version 3 /name [ 9 5374726f6b65446f74 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 3 646f74 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_setStroke) /localizedName [ 12 e7b79ae38292e8a8ade5ae9a ] /isOpen 0 /isOn 1 /hasDialog 0 /parameterCount 12 /parameter-1 { /key 2003072104 /showInPalette 4294967295 /type (unit real) /value 4.0 /unit 592476268 } /parameter-2 { /key 1667330094 /showInPalette 4294967295 /type (enumerated) /name [ 12 e4b8b8e59e8be7b79ae7abaf ] /value 1 } /parameter-3 { /key 1836344690 /showInPalette 4294967295 /type (real) /value 10.0 } /parameter-4 { /key 1785686382 /showInPalette 4294967295 /type (enumerated) /name [ 18 e3839ee382a4e382bfe383bce7b590e59088 ] /value 0 } /parameter-5 { /key 1684825454 /showInPalette 4294967295 /type (integer) /value 2 } /parameter-6 { /key 1685284913 /showInPalette 4294967295 /type (unit real) /value 0.0 /unit 592476268 } /parameter-7 { /key 1685284914 /showInPalette 4294967295 /type (unit real) /value 6.0 /unit 592476268 } /parameter-8 { /key 1684104298 /showInPalette 4294967295 /type (boolean) /value 1 } /parameter-9 { /key 1634231345 /showInPalette 4294967295 /type (ustring) /value [ 8 5be381aae381975d ] } /parameter-10 { /key 1634231346 /showInPalette 4294967295 /type (ustring) /value [ 8 5be381aae381975d ] } /parameter-11 { /key 1634230636 /showInPalette 4294967295 /type (enumerated) /name [ 24 e38391e382b9e381aee7b582e782b9e381abe9858de7bdae ] /value 0 } /parameter-12 { /key 1634494318 /showInPalette 4294967295 /type (enumerated) /name [ 6 e4b8ade5a4ae ] /value 0 } } }';

var isStrokeDotActionLoaded = false;

/**
 * StrokeDot アクションを一時ファイル経由で読み込む（読み込み済みなら何もしない）
 * @returns {void}
 */
function loadStrokeDotAction() {
    if (isStrokeDotActionLoaded) return;

    var actionFile = new File('~/StrokeDot.aia');
    actionFile.open('w');
    actionFile.write(STROKE_DOT_ACTION);
    actionFile.close();
    app.loadAction(actionFile);
    actionFile.remove();

    isStrokeDotActionLoaded = true;
}

/**
 * StrokeDot アクションを破棄する
 * @returns {void}
 */
function unloadStrokeDotAction() {
    if (!isStrokeDotActionLoaded) return;
    try {
        app.unloadAction('StrokeDot', '');
    } catch (e) {
        /* すでに破棄済みでも続行 / Ignore when it is already gone */
    }
    isStrokeDotActionLoaded = false;
}

/**
 * 対象パスへ StrokeDot アクションを適用する（選択状態は元に戻す）
 * @param {Document} doc - 対象ドキュメント
 * @param {PathItem} item - 対象パス
 * @returns {void}
 */
function applyStrokeDotAction(doc, item) {
    loadStrokeDotAction();
    var previousSelection = saveSelection(doc);
    try {
        selectItems(doc, [item]);
        app.doScript('dot', 'StrokeDot', false);
    } finally {
        restoreSelection(doc, previousSelection);
    }
}

// =========================================
// 選択・オブジェクト操作 / Selection and object helpers
// =========================================

/**
 * 現在の選択を安全に取得する
 * @param {Document} doc - 対象ドキュメント
 * @returns {Array} 選択オブジェクトの配列（取得できないときは空配列）
 */
function getSafeSelection(doc) {
    try {
        return doc.selection || [];
    } catch (e) {
        return [];
    }
}

/**
 * 現在の選択を配列として控える
 * @param {Document} doc - 対象ドキュメント
 * @returns {Array} 選択オブジェクトの配列
 */
function saveSelection(doc) {
    var items = [];
    var selection = getSafeSelection(doc);
    for (var i = 0; i < selection.length; i++) items.push(selection[i]);
    return items;
}

/**
 * 控えておいた選択を復元する（削除済みのオブジェクトは読み飛ばす）
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - 復元するオブジェクトの配列
 * @returns {void}
 */
function restoreSelection(doc, items) {
    doc.selection = null;
    if (!items) return;
    for (var i = 0; i < items.length; i++) {
        try {
            items[i].selected = true;
        } catch (e) {
            /* 生成前に消えたオブジェクトは無視 / Skip items that no longer exist */
        }
    }
}

/**
 * 指定オブジェクトだけを選択する
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - 選択するオブジェクトの配列
 * @returns {void}
 */
function selectItems(doc, items) {
    doc.selection = null;
    for (var i = 0; i < items.length; i++) {
        items[i].selected = true;
    }
}

/**
 * 選択が1つだけのときにそのオブジェクトを返す
 * @param {Document} doc - 対象ドキュメント
 * @returns {Object} 先頭の選択オブジェクト（なければ null）
 */
function getSingleSelection(doc) {
    var selection = getSafeSelection(doc);
    return selection.length > 0 ? selection[0] : null;
}

/**
 * 線をアウトライン化する
 * @param {Document} doc - 対象ドキュメント
 * @param {PathItem} item - 対象パス
 * @returns {Object} アウトライン化後のオブジェクト
 */
function outlineStrokeItem(doc, item) {
    var previousSelection = saveSelection(doc);
    try {
        selectItems(doc, [item]);
        app.executeMenuCommand('Live Outline Stroke');
        var outlined = getSingleSelection(doc);
        if (!outlined) throw new Error('Live Outline Stroke failed.');
        return outlined;
    } finally {
        restoreSelection(doc, previousSelection);
    }
}

/**
 * 複数のオブジェクトをグループ化する
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - グループ化するオブジェクトの配列
 * @returns {Object} グループ（1つだけならそのオブジェクト、空なら null）
 */
function groupItems(doc, items) {
    if (!items || items.length === 0) return null;
    if (items.length === 1) return items[0];

    var previousSelection = saveSelection(doc);
    try {
        selectItems(doc, items);
        app.executeMenuCommand('group');
        var grouped = getSingleSelection(doc);
        if (!grouped) throw new Error('Group command failed.');
        return grouped;
    } finally {
        restoreSelection(doc, previousSelection);
    }
}

/**
 * ドキュメントのカラースペースに応じたグレーを生成する
 * @param {Document} doc - 対象ドキュメント
 * @returns {CMYKColor|GrayColor} 生成した色
 */
function makeFallbackGray(doc) {
    if (doc.documentColorSpace == DocumentColorSpace.CMYK) {
        var cmyk = new CMYKColor();
        cmyk.cyan = 0;
        cmyk.magenta = 0;
        cmyk.yellow = 0;
        cmyk.black = FALLBACK_GRAY_TINT;
        return cmyk;
    }
    var gray = new GrayColor();
    gray.gray = FALLBACK_GRAY_TINT;
    return gray;
}

/**
 * パスファインダーで抜けるように、対象を「塗りのみ」の状態へそろえる
 * @param {Array} items - 対象オブジェクトの配列
 * @param {Document} doc - 対象ドキュメント
 * @returns {void}
 */
function normalizeInitialAppearance(items, doc) {
    for (var i = 0; i < items.length; i++) {
        var item = items[i];

        if (item.filled) {
            item.stroked = false;
            continue;
        }

        if (item.stroked) {
            try {
                item.fillColor = item.strokeColor;
            } catch (e) {
                item.fillColor = makeFallbackGray(doc);
            }
        } else {
            item.fillColor = makeFallbackGray(doc);
        }
        item.filled = true;
        item.stroked = false;
    }
}

/**
 * 長方形パスかどうかを判定する
 * @param {Object} item - 判定するオブジェクト
 * @returns {boolean} 4隅がバウンディングボックスに一致する長方形なら true
 */
function isRectanglePath(item) {
    if (!item || item.typename !== 'PathItem') return false;
    if (item.guides || item.clipping) return false;
    if (item.pathPoints.length !== 4) return false;

    var geometricBounds = item.geometricBounds;
    var left = geometricBounds[0];
    var top = geometricBounds[1];
    var right = geometricBounds[2];
    var bottom = geometricBounds[3];

    var hasLeftTop = false, hasRightTop = false, hasRightBottom = false, hasLeftBottom = false;

    for (var i = 0; i < 4; i++) {
        var anchor = item.pathPoints[i].anchor;
        var isLeft = Math.abs(anchor[0] - left) <= RECTANGLE_TOLERANCE;
        var isRight = Math.abs(anchor[0] - right) <= RECTANGLE_TOLERANCE;
        var isTop = Math.abs(anchor[1] - top) <= RECTANGLE_TOLERANCE;
        var isBottom = Math.abs(anchor[1] - bottom) <= RECTANGLE_TOLERANCE;

        if (isLeft && isTop) hasLeftTop = true;
        else if (isRight && isTop) hasRightTop = true;
        else if (isRight && isBottom) hasRightBottom = true;
        else if (isLeft && isBottom) hasLeftBottom = true;
        else return false;
    }

    return hasLeftTop && hasRightTop && hasRightBottom && hasLeftBottom;
}

// =========================================
// 形状の生成 / Shape builders
// =========================================

/**
 * 角丸のライブエフェクトを適用する
 * @param {Array} items - 対象オブジェクトの配列
 * @param {number} radiusPt - 角丸の半径（pt）
 * @returns {void}
 */
function applyRoundCornersEffect(items, radiusPt) {
    if (!(radiusPt > 0)) return;
    var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + radiusPt + ' "/></LiveEffect>';
    for (var i = 0; i < items.length; i++) {
        items[i].applyEffect(xml);
    }
}

/**
 * 塗りだけを持つ円を追加する
 * @param {Document} doc - 対象ドキュメント
 * @param {number} centerX - 中心のX座標（pt）
 * @param {number} centerY - 中心のY座標（pt）
 * @param {number} sizePt - 直径（pt）
 * @returns {PathItem} 生成した円
 */
function addFilledCircle(doc, centerX, centerY, sizePt) {
    var circle = doc.pathItems.ellipse(centerY + sizePt / 2, centerX - sizePt / 2, sizePt, sizePt);
    circle.filled = true;
    circle.stroked = false;
    return circle;
}

/**
 * 塗りだけを持つ正方形を45°回転させて追加する（三角スリット・ギザギザ用）
 * @param {Document} doc - 対象ドキュメント
 * @param {number} centerX - 中心のX座標（pt）
 * @param {number} centerY - 中心のY座標（pt）
 * @param {number} sizePt - 一辺の長さ（pt）
 * @returns {PathItem} 生成したひし形
 */
function addFilledDiamond(doc, centerX, centerY, sizePt) {
    var diamond = doc.pathItems.rectangle(centerY + sizePt / 2, centerX - sizePt / 2, sizePt, sizePt);
    diamond.filled = true;
    diamond.stroked = false;
    diamond.rotate(45);
    return diamond;
}

/**
 * 線だけを持つ直線を追加する
 * @param {Document} doc - 対象ドキュメント
 * @param {Array<Array<number>>} points - 始点と終点の座標
 * @returns {PathItem} 生成した直線
 */
function addStrokedLine(doc, points) {
    var line = doc.pathItems.add();
    line.setEntirePath(points);
    line.filled = false;
    line.stroked = true;
    return line;
}

/**
 * 破線の長さが間隔にできるだけ近づくよう、破線パターンを求める
 * @param {number} lineLength - 破線を敷く全体長（pt）
 * @param {number} gapPt - 間隔（pt）
 * @returns {Array<number>} [破線長, 間隔] の配列（求まらなければ null）
 */
function computeDashPattern(lineLength, gapPt) {
    if (!(gapPt > 0) || !(lineLength > 0)) return null;

    var targetGapCount = (lineLength - gapPt) / (gapPt * 2);
    var candidates = [
        Math.floor(targetGapCount),
        Math.ceil(targetGapCount),
        Math.floor(targetGapCount) - 1,
        Math.ceil(targetGapCount) + 1,
        1
    ];
    var bestDashLength = 0;
    var bestDifference = Number.MAX_VALUE;

    for (var i = 0; i < candidates.length; i++) {
        var gapCount = Math.max(1, candidates[i]);
        var dashLength = (lineLength - gapCount * gapPt) / (gapCount + 1);
        if (dashLength <= 0) continue;

        var difference = Math.abs(dashLength - gapPt);
        if (difference < bestDifference) {
            bestDifference = difference;
            bestDashLength = dashLength;
        }
    }

    return bestDashLength > 0 ? [bestDashLength, gapPt] : null;
}

/**
 * 中央の分割線（ミシン目／破線）を追加する
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - パスファインダー対象を追加する配列
 * @param {Object} bounds - 長方形の位置とサイズ
 * @param {Object} settings - UIから読み取った設定
 * @param {Object} geometry - ポイント換算済みの寸法
 * @returns {void}
 */
function addCenterPerforation(doc, items, bounds, settings, geometry) {
    if (!settings.center.enabled) return;
    if (settings.edge.shapeOnly && settings.edge.mode !== 'none') return;

    var insetPt = geometry.centerInsetPt;
    var line = addStrokedLine(doc, [
        [bounds.centerLineX, bounds.top - insetPt],
        [bounds.centerLineX, bounds.top - bounds.height + insetPt]
    ]);

    var useDotStyle = (settings.center.mode !== 'dash');
    if (useDotStyle || insetPt === 0) {
        applyStrokeDotAction(doc, line);
    }

    line.strokeWidth = geometry.centerWidthPt;
    if (useDotStyle) {
        line.strokeDashes = [0, geometry.centerGapPt];
    } else {
        line.strokeCap = StrokeCap.BUTTENDCAP;
        var dashPattern = (insetPt !== 0)
            ? computeDashPattern(Math.max(0, bounds.height - insetPt * 2), geometry.centerGapPt)
            : null;
        line.strokeDashes = dashPattern || [geometry.centerGapPt, geometry.centerGapPt];
    }

    items.push(outlineStrokeItem(doc, line));
}

/**
 * 左右のミシン目を追加する
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - パスファインダー対象を追加する配列
 * @param {Object} bounds - 長方形の位置とサイズ
 * @param {Object} settings - UIから読み取った設定
 * @param {Object} geometry - ポイント換算済みの寸法
 * @returns {void}
 */
function addSidePerforation(doc, items, bounds, settings, geometry) {
    if (!settings.lr.enabled) return;

    var insetPt = geometry.sideInsetPt;
    var positions = [bounds.left, bounds.left + bounds.width];

    for (var i = 0; i < positions.length; i++) {
        var line = addStrokedLine(doc, [
            [positions[i], bounds.top - insetPt],
            [positions[i], bounds.top - bounds.height + insetPt]
        ]);
        applyStrokeDotAction(doc, line);
        line.strokeWidth = geometry.sideWidthPt;
        line.strokeDashes = [0, geometry.sideGapPt];
        items.push(outlineStrokeItem(doc, line));
    }
}

/**
 * 辺のギザギザ（45°回転した正方形の連続）を追加する
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - パスファインダー対象を追加する配列
 * @param {Object} bounds - 長方形の位置とサイズ
 * @param {Object} settings - UIから読み取った設定
 * @param {Object} geometry - ポイント換算済みの寸法
 * @returns {void}
 */
function addZigzag(doc, items, bounds, settings, geometry) {
    if (settings.zigzag.mode === 'none' || !(geometry.zigzagSizePt > 0)) return;

    /* 指定サイズは対角線の長さなので、正方形の辺に変換する / The size is the diagonal, convert it to a side */
    var side = geometry.zigzagSizePt / Math.sqrt(2);
    var step = geometry.zigzagSizePt + geometry.zigzagGapPt;
    var total = step * geometry.zigzagRepeat - geometry.zigzagGapPt;
    var zigzagItems = [];
    var lineIndex, repeatIndex;

    if (settings.zigzag.mode === 'lr') {
        /* 左右：辺の垂直中央を基準に配置 / Left and right: centered vertically on each side */
        var startY = (bounds.top - bounds.height / 2) + total / 2 - geometry.zigzagSizePt / 2;
        var xPositions = [bounds.left, bounds.left + bounds.width];
        for (lineIndex = 0; lineIndex < xPositions.length; lineIndex++) {
            for (repeatIndex = 0; repeatIndex < geometry.zigzagRepeat; repeatIndex++) {
                zigzagItems.push(addFilledDiamond(doc, xPositions[lineIndex], startY - step * repeatIndex, side));
            }
        }
    } else {
        /* 上下：辺の水平中央を基準に配置 / Top and bottom: centered horizontally on each side */
        var startX = (bounds.left + bounds.width / 2) - total / 2 + geometry.zigzagSizePt / 2;
        var yPositions = [bounds.top, bounds.top - bounds.height];
        for (lineIndex = 0; lineIndex < yPositions.length; lineIndex++) {
            for (repeatIndex = 0; repeatIndex < geometry.zigzagRepeat; repeatIndex++) {
                zigzagItems.push(addFilledDiamond(doc, startX + step * repeatIndex, yPositions[lineIndex], side));
            }
        }
    }

    if (zigzagItems.length > 0) items.push(groupItems(doc, zigzagItems));
}

/**
 * 分割線の両端に置くエッジ形状を追加する
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - パスファインダー対象を追加する配列
 * @param {Object} bounds - 長方形の位置とサイズ
 * @param {Object} settings - UIから読み取った設定
 * @param {Object} geometry - ポイント換算済みの寸法
 * @returns {void}
 */
function addCenterEdges(doc, items, bounds, settings, geometry) {
    if (!settings.center.enabled) return;
    if (settings.edge.mode !== 'circle' && settings.edge.mode !== 'triangle') return;
    if (!(geometry.edgeSizePt > 0)) return;

    var yPositions = [bounds.top, bounds.top - bounds.height];
    for (var i = 0; i < yPositions.length; i++) {
        items.push(settings.edge.mode === 'circle'
            ? addFilledCircle(doc, bounds.centerLineX, yPositions[i], geometry.edgeSizePt)
            : addFilledDiamond(doc, bounds.centerLineX, yPositions[i], geometry.edgeSizePt));
    }
}

/**
 * 左右のスリット／ホールを追加する
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - パスファインダー対象を追加する配列
 * @param {Object} bounds - 長方形の位置とサイズ
 * @param {Object} settings - UIから読み取った設定
 * @param {Object} geometry - ポイント換算済みの寸法
 * @returns {void}
 */
function addHoles(doc, items, bounds, settings, geometry) {
    if (settings.hole.mode === 'none' || !(geometry.holeSizePt > 0)) return;

    var centerY = bounds.top - bounds.height / 2;
    var xPositions = [];
    if (settings.hole.left) xPositions.push(bounds.left);
    if (settings.hole.right) xPositions.push(bounds.left + bounds.width);

    var holeItems = [];
    for (var i = 0; i < xPositions.length; i++) {
        holeItems.push(settings.hole.mode === 'circle'
            ? addFilledCircle(doc, xPositions[i], centerY, geometry.holeSizePt)
            : addFilledDiamond(doc, xPositions[i], centerY, geometry.holeSizePt));
    }

    if (holeItems.length > 0) items.push(groupItems(doc, holeItems));
}

/**
 * 四隅の逆角丸を追加する（上下辺に太い破線を敷いて角だけを削る）
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - パスファインダー対象を追加する配列
 * @param {Object} bounds - 長方形の位置とサイズ
 * @param {Object} settings - UIから読み取った設定
 * @param {Object} geometry - ポイント換算済みの寸法
 * @returns {void}
 */
function addInverseCorners(doc, items, bounds, settings, geometry) {
    if (settings.corner.mode !== 'inverse' || !(geometry.cornerSizePt > 0)) return;

    var right = bounds.left + bounds.width;
    var yPositions = [bounds.top, bounds.top - bounds.height];

    for (var i = 0; i < yPositions.length; i++) {
        var line = addStrokedLine(doc, [[bounds.left, yPositions[i]], [right, yPositions[i]]]);
        applyStrokeDotAction(doc, line);
        line.strokeWidth = geometry.cornerSizePt;
        line.strokeDashes = [0, INVERSE_CORNER_DASH_GAP];
        items.push(outlineStrokeItem(doc, line));
    }
}

/**
 * 四隅の面取りを追加する
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - パスファインダー対象を追加する配列
 * @param {Object} bounds - 長方形の位置とサイズ
 * @param {Object} settings - UIから読み取った設定
 * @param {Object} geometry - ポイント換算済みの寸法
 * @returns {void}
 */
function addChamferCorners(doc, items, bounds, settings, geometry) {
    if (settings.corner.mode !== 'chamfer' || !(geometry.cornerSizePt > 0)) return;

    var right = bounds.left + bounds.width;
    var bottom = bounds.top - bounds.height;
    var corners = [
        [bounds.left, bounds.top],
        [right, bounds.top],
        [bounds.left, bottom],
        [right, bottom]
    ];

    for (var i = 0; i < corners.length; i++) {
        items.push(addFilledDiamond(doc, corners[i][0], corners[i][1], geometry.cornerSizePt));
    }
}

/**
 * ダブル角丸を適用する（分割線で2分割し、それぞれに角丸をかけて合体する）
 * @param {Document} doc - 対象ドキュメント
 * @param {PathItem} rect - 元の長方形
 * @param {Array} items - パスファインダー対象の配列（先頭を差し替える）
 * @param {Object} bounds - 長方形の位置とサイズ
 * @param {Object} settings - UIから読み取った設定
 * @param {Object} geometry - ポイント換算済みの寸法
 * @returns {void}
 */
function applyDoubleRoundEdge(doc, rect, items, bounds, settings, geometry) {
    if (!(geometry.edgeSizePt > 0)) return;

    var leftRect = doc.pathItems.rectangle(bounds.top, bounds.left, bounds.centerLineX - bounds.left, bounds.height);
    var rightRect = doc.pathItems.rectangle(bounds.top, bounds.centerLineX, bounds.left + bounds.width - bounds.centerLineX, bounds.height);
    var halves = [leftRect, rightRect];

    for (var i = 0; i < halves.length; i++) {
        halves[i].filled = true;
        halves[i].stroked = false;
    }
    try {
        leftRect.fillColor = rect.fillColor;
        rightRect.fillColor = rect.fillColor;
    } catch (e) {
        /* 元の塗りを引き継げないときは既定色のまま / Keep the default fill when it cannot be copied */
    }

    /* エッジサイズの角丸に、コーナーパネルの角丸を重ねる / Stack the corner radius on top of the edge radius */
    applyRoundCornersEffect(halves, geometry.edgeSizePt);
    if (settings.corner.mode === 'round') {
        applyRoundCornersEffect(halves, geometry.cornerSizePt);
    }

    rect.remove();
    selectItems(doc, [groupItems(doc, halves)]);
    app.executeMenuCommand('Live Pathfinder Add');
    items[0] = getSingleSelection(doc);
}

/**
 * 生成した形状をまとめて前面オブジェクトで型抜きする
 * @param {Document} doc - 対象ドキュメント
 * @param {Array} items - 型抜き対象の配列（先頭がベースの長方形）
 * @param {Array} results - 結果を追加する配列
 * @returns {void}
 */
function finalizeSubtract(doc, items, results) {
    selectItems(doc, [groupItems(doc, items)]);
    app.executeMenuCommand('Live Pathfinder Subtract');
    var subtracted = getSingleSelection(doc);
    if (subtracted) results.push(subtracted);
}

// =========================================
// プリセット / Presets
// =========================================

var PRESETS = [
    {
        name:    { ja: "0: クリア", en: "0: Clear" },
        lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
        center:  { enabled: false, offset: 30, mode: "dot", width: 3, gap: 6, length: -10 },
        zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
        corner:  { mode: "none", size: 5 },
        hole:    { mode: "none", left: true, right: true, size: 10 },
        edge:    { mode: "wround", size: 5, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    },
    {
        name:    { ja: "1: W角丸＋分割線", en: "1: Double Rounded + Divider" },
        lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
        center:  { enabled: true, offset: 30, mode: "dot", width: 3, gap: 3, length: -10 },
        zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
        corner:  { mode: "none", size: 5 },
        hole:    { mode: "none", left: true, right: true, size: 10 },
        edge:    { mode: "wround", size: 5, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    },
    {
        name:    { ja: "2: 左右に2ホール", en: "2: Two Side Holes" },
        lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
        center:  { enabled: false, offset: 30, mode: "dot", width: 3, gap: 6, length: -6 },
        zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
        corner:  { mode: "none", size: 5 },
        hole:    { mode: "circle", left: true, right: true, size: 15 },
        edge:    { mode: "none", size: 5, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    },
    {
        name:    { ja: "3: ミシン目", en: "3: Perforation" },
        lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
        center:  { enabled: true, offset: 35, mode: "dot", width: 3, gap: 6, length: 0 },
        zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
        corner:  { mode: "round", size: 6 },
        hole:    { mode: "none", left: true, right: true, size: 10 },
        edge:    { mode: "none", size: 5, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    },
    {
        name:    { ja: "4: 四隅に逆角丸", en: "4: Inverse Round Corners" },
        lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
        center:  { enabled: false, offset: 35, mode: "dot", width: 3, gap: 6, length: 0 },
        zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
        corner:  { mode: "inverse", size: 25 },
        hole:    { mode: "none", left: true, right: true, size: 10 },
        edge:    { mode: "none", size: 5, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    },
    {
        name:    { ja: "5: 左右に三角スリット", en: "5: Triangular Side Slits" },
        lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
        center:  { enabled: false, offset: 35, mode: "dot", width: 3, gap: 6, length: 0 },
        zigzag:  { mode: "none", size: 10, gap: 0, repeat: 3 },
        corner:  { mode: "round", size: 3 },
        hole:    { mode: "triangle", left: true, right: true, size: 15 },
        edge:    { mode: "none", size: 5, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    },
    {
        name:    { ja: "6: ミシン目＋ギザギザ上下", en: "6: Perforation + T/B Zigzag" },
        lr:      { enabled: true, linkCenter: false, width: 4, gap: 7, length: 0 },
        center:  { enabled: false, offset: 35, mode: "dot", width: 3, gap: 6, length: 0 },
        zigzag:  { mode: "tb", size: 10, gap: 0, repeat: 7 },
        corner:  { mode: "none", size: 3 },
        hole:    { mode: "none", left: true, right: true, size: 15 },
        edge:    { mode: "none", size: 5, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    },
    {
        name:    { ja: "7: 面取り＋分割線（破線）", en: "7: Chamfer + Dashed Divider" },
        lr:      { enabled: false, linkCenter: false, width: 4, gap: 7, length: 0 },
        center:  { enabled: true, offset: 35, mode: "dash", width: 1, gap: 3, length: 0 },
        zigzag:  { mode: "none", size: 10, gap: 0, repeat: 7 },
        corner:  { mode: "chamfer", size: 13 },
        hole:    { mode: "none", left: true, right: true, size: 15 },
        edge:    { mode: "none", size: 5, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    },
    {
        name:    { ja: "8: 分割線＋エッジ（円）＋ギザギザ", en: "8: Divider + Circle Edges + Zigzag" },
        lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
        center:  { enabled: true, offset: 30, mode: "dot", width: 3, gap: 6, length: -13 },
        zigzag:  { mode: "lr", size: 7, gap: 0, repeat: 7 },
        corner:  { mode: "none", size: 5 },
        hole:    { mode: "none", left: true, right: true, size: 10 },
        edge:    { mode: "circle", size: 10, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    },
    {
        name:    { ja: "9: 分割線（破線）＋三角エッジ＋左ホール", en: "9: Dashed Divider + Triangle Edges + Left Hole" },
        lr:      { enabled: false, linkCenter: true, width: 3, gap: 6, length: 0 },
        center:  { enabled: true, offset: 40, mode: "dash", width: 1, gap: 3, length: -7.2 },
        zigzag:  { mode: "none", size: 7, gap: 0, repeat: 7 },
        corner:  { mode: "round", size: 3 },
        hole:    { mode: "circle", left: true, right: false, size: 20 },
        edge:    { mode: "triangle", size: 6, shapeOnly: false },
        preview: { enabled: true, expandAppearance: false }
    }
];

/**
 * プリセットや設定オブジェクトを再帰的に複製する
 * @param {*} value - 複製する値
 * @returns {*} 複製した値
 */
function cloneSimpleValue(value) {
    if (value === null || typeof value !== 'object') return value;
    if (value instanceof Array) {
        var cloned = [];
        for (var i = 0; i < value.length; i++) cloned.push(cloneSimpleValue(value[i]));
        return cloned;
    }
    var out = {};
    for (var key in value) {
        if (value.hasOwnProperty(key)) out[key] = cloneSimpleValue(value[key]);
    }
    return out;
}

/**
 * 有限な数値かどうかを判定する
 * @param {*} value - 判定する値
 * @returns {boolean} 有限な数値なら true
 */
function isFiniteNumber(value) {
    return typeof value === 'number' && !isNaN(value) && isFinite(value);
}

/**
 * 値をJavaScriptのソース表記へ変換する
 * @param {*} value - 変換する値
 * @param {number} indentLevel - 字下げの深さ
 * @returns {string} ソース表記
 */
function toCodeValue(value, indentLevel) {
    var indent = new Array(indentLevel + 1).join('    ');
    var nextIndent = new Array(indentLevel + 2).join('    ');
    var parts = [];
    var i, key;

    if (typeof value === 'string') {
        return '"' + value.replace(/\\/g, '\\\\').replace(/"/g, '\\"').replace(/\r/g, '\\r').replace(/\n/g, '\\n') + '"';
    }
    if (typeof value === 'number') {
        return isFiniteNumber(value) ? String(value) : '0';
    }
    if (typeof value === 'boolean') {
        return value ? 'true' : 'false';
    }
    if (value === null) {
        return 'null';
    }
    if (value instanceof Array) {
        for (i = 0; i < value.length; i++) parts.push(toCodeValue(value[i], indentLevel + 1));
        return '[' + parts.join(', ') + ']';
    }
    if (typeof value === 'object') {
        for (key in value) {
            if (value.hasOwnProperty(key)) {
                parts.push(nextIndent + key + ': ' + toCodeValue(value[key], indentLevel + 1));
            }
        }
        return '{\n' + parts.join(',\n') + '\n' + indent + '}';
    }
    return 'null';
}

/**
 * プリセットを PRESETS へ貼り付けられるコードに変換する
 * @param {Object} preset - 対象プリセット
 * @returns {string} 組み込み用のコード
 */
function presetToCode(preset) {
    return ',\n' + toCodeValue(preset, 0);
}

/**
 * プリセットの表示名を取り出す
 * @param {Object} preset - 対象プリセット
 * @returns {string} 表示名
 */
function getPresetDisplayName(preset) {
    if (!preset) return '';
    if (typeof preset.name === 'string') return preset.name;
    if (preset.name && typeof preset.name === 'object') return getLabel(preset.name);
    return '';
}

/**
 * ドロップダウンの選択項目の文言を取り出す
 * @param {DropDownList} dropdown - 対象ドロップダウン
 * @returns {string} 選択中の文言（未選択なら空文字）
 */
function getDropdownSelectionText(dropdown) {
    return (dropdown && dropdown.selection) ? dropdown.selection.text : '';
}

/**
 * プリセット一覧をドロップダウンへ流し込む
 * @param {DropDownList} dropdown - 対象ドロップダウン
 * @param {Array<Object>} presets - プリセットの配列
 * @returns {void}
 */
function setDropdownItemsFromPresets(dropdown, presets) {
    dropdown.removeAll();
    for (var i = 0; i < presets.length; i++) {
        dropdown.add('item', getPresetDisplayName(presets[i]));
    }
    if (dropdown.items.length > 0) dropdown.selection = 0;
}

// =========================================
// 設定の解決 / Settings resolution
// =========================================

/**
 * 設定をポイント単位の寸法へ変換する
 * @param {Object} settings - UIから読み取った設定
 * @returns {Object} ポイント換算済みの寸法
 */
function toGeometry(settings) {
    /* ドットのミシン目に連動しているときは、左右も分割線と同じ値・同じ単位系を使う / Linked dot perforation shares the divider values and units */
    var linked = settings.lr.enabled && settings.lr.linkCenter && settings.center.mode === 'dot';
    var side = linked ? settings.center : settings.lr;

    return {
        linked: linked,
        sideWidthPt: toPt(side.width, linked ? strokePtFactor : rulerPtFactor),
        sideGapPt: toPt(side.gap, rulerPtFactor),
        sideInsetPt: toPt(Math.abs(Math.min(0, side.length)), rulerPtFactor),

        centerWidthPt: toPt(settings.center.width, strokePtFactor),
        centerGapPt: toPt(settings.center.gap, rulerPtFactor),
        centerInsetPt: toPt(Math.abs(Math.min(0, settings.center.length)), rulerPtFactor),
        offsetPt: toPt(settings.center.offset, rulerPtFactor),

        zigzagSizePt: toPt(settings.zigzag.size, rulerPtFactor),
        zigzagGapPt: toPt(settings.zigzag.gap || 0, rulerPtFactor),
        zigzagRepeat: Math.max(1, Math.round(settings.zigzag.repeat || 1)),

        cornerSizePt: toPt(settings.corner.size, rulerPtFactor),
        holeSizePt: toPt(settings.hole.size, rulerPtFactor),
        edgeSizePt: toPt(settings.edge.size, rulerPtFactor)
    };
}

/**
 * 寸法に数値以外が混じっていないかを確認する
 * @param {Object} geometry - ポイント換算済みの寸法
 * @returns {boolean} すべて数値なら true
 */
function isValidGeometry(geometry) {
    var keys = [
        'sideWidthPt', 'sideGapPt', 'sideInsetPt',
        'centerWidthPt', 'centerGapPt', 'centerInsetPt', 'offsetPt',
        'zigzagSizePt', 'zigzagGapPt', 'zigzagRepeat',
        'cornerSizePt', 'holeSizePt', 'edgeSizePt'
    ];
    for (var i = 0; i < keys.length; i++) {
        if (isNaN(geometry[keys[i]])) return false;
    }
    return true;
}

// =========================================
// 起動時チェック / Startup validation
// =========================================

/**
 * 実行できる選択状態かどうかを確認する
 * @returns {Object} { doc: Document, selectedItems: Array }（実行できないときは null）
 */
function validateStartupState() {
    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.openDocument));
        return null;
    }

    var doc = app.activeDocument;
    var selectedItems = getSafeSelection(doc);

    if (selectedItems.length === 0) {
        alert(getLabel(LABELS.alert.selectRectangle));
        return null;
    }
    if (selectedItems.length > 1) {
        alert(getLabel(LABELS.alert.singleOnly));
        return null;
    }
    if (selectedItems[0].typename === 'GroupItem') {
        alert(getLabel(LABELS.alert.groupNotAllowed));
        return null;
    }
    if (!isRectanglePath(selectedItems[0])) {
        alert(getLabel(LABELS.alert.rectangleOnly));
        return null;
    }

    return { doc: doc, selectedItems: selectedItems };
}

// =========================================
// メイン処理 / Main
// =========================================

/**
 * スクリプトのエントリーポイント
 * @returns {void}
 */
function main() {
    var startup = validateStartupState();
    if (!startup) return;

    var doc = startup.doc;
    var selectedItems = startup.selectedItems;

    app.executeMenuCommand('edge');
    app.executeMenuCommand('AI Bounding Box Toggle');

    try {
        var originalLayer = selectedItems[0].layer;

        /* プレビュー専用レイヤーは必要になるまで作らない / The preview layer is created lazily */
        var previewLayer = null;

        /**
         * レイヤーがまだ生きているかを判定する
         * @param {Layer} layer - 対象レイヤー
         * @returns {boolean} 参照できれば true
         */
        function hasUsableLayer(layer) {
            if (!layer) return false;
            try {
                return !!layer.name;
            } catch (e) {
                return false;
            }
        }

        /**
         * プレビュー専用レイヤーを用意する
         * @returns {Layer} プレビューレイヤー
         */
        function ensurePreviewLayer() {
            if (hasUsableLayer(previewLayer)) return previewLayer;
            previewLayer = doc.layers.add();
            previewLayer.name = getLabel(LABELS.layerName.preview);
            return previewLayer;
        }

        /**
         * プレビューレイヤーの中身を空にする
         * @returns {void}
         */
        function clearPreviewLayerItems() {
            if (!hasUsableLayer(previewLayer)) return;
            try {
                while (previewLayer.pageItems.length > 0) {
                    previewLayer.pageItems[0].remove();
                }
            } catch (e) {
                /* 途中で参照が切れても続行 / Continue even if the reference breaks */
            }
        }

        /**
         * プレビューを消して元オブジェクトを表示に戻す
         * @param {boolean} skipRedraw - 再描画を省略するか
         * @returns {void}
         */
        function removePreview(skipRedraw) {
            clearPreviewLayerItems();
            for (var i = 0; i < selectedItems.length; i++) {
                selectedItems[i].hidden = false;
            }
            if (!skipRedraw) app.redraw();
        }

        /**
         * プレビューレイヤーそのものを削除する
         * @returns {void}
         */
        function removePreviewLayer() {
            if (!hasUsableLayer(previewLayer)) return;
            clearPreviewLayerItems();
            try {
                previewLayer.remove();
            } catch (e) {
                /* すでに削除済みでも続行 / Continue even if it is already removed */
            }
            previewLayer = null;
        }

        /**
         * UIの状態をプリセットと同じ構造の設定として読み取る
         * @param {string} presetName - 設定に付ける名前（省略時は選択中のプリセット名）
         * @returns {Object} 設定オブジェクト
         */
        function readSettingsFromUI(presetName) {
            return {
                name: presetName || getDropdownSelectionText(ddPreset) || getLabel(LABELS.fallbackName.customPreset),
                lr: {
                    enabled: chkSidePerforation.value,
                    linkCenter: chkLinkToCenter.value,
                    width: parseFloat(txtSideWidth.text),
                    gap: parseFloat(txtSideGap.text),
                    length: parseFloat(txtSideInset.text)
                },
                center: {
                    enabled: chkCenterSplit.value,
                    offset: parseFloat(txtCenterOffset.text),
                    mode: rdDividerDash.value ? 'dash' : 'dot',
                    width: parseFloat(txtDividerWidth.text),
                    gap: parseFloat(txtDividerGap.text),
                    length: parseFloat(txtDividerInset.text)
                },
                zigzag: {
                    mode: rdZigzagLeftRight.value ? 'lr' : (rdZigzagTopBottom.value ? 'tb' : 'none'),
                    size: parseFloat(txtZigzagSize.text),
                    gap: parseFloat(txtZigzagGap.text),
                    repeat: parseFloat(txtZigzagRepeat.text)
                },
                corner: {
                    mode: rdCornerRound.value ? 'round' : (rdCornerInverse.value ? 'inverse' : (rdCornerChamfer.value ? 'chamfer' : 'none')),
                    size: parseFloat(txtCornerSize.text)
                },
                hole: {
                    mode: rdHoleCircle.value ? 'circle' : (rdHoleTriangle.value ? 'triangle' : 'none'),
                    left: chkHoleLeft.value,
                    right: chkHoleRight.value,
                    size: parseFloat(txtHoleSize.text)
                },
                edge: {
                    mode: chkEdgeDoubleRound.value ? 'wround' : (rdEdgeCircle.value ? 'circle' : (rdEdgeTriangle.value ? 'triangle' : 'none')),
                    size: parseFloat(txtEdgeSize.text),
                    shapeOnly: chkEdgeOnly.value
                },
                preview: {
                    enabled: chkPreview.value,
                    expandAppearance: chkExpandAppearance.value
                }
            };
        }

        /**
         * 設定をUIへ反映する
         * @param {Object} settings - 反映する設定
         * @returns {void}
         */
        function applySettingsToUI(settings) {
            if (!settings) return;

            chkSidePerforation.value = !!settings.lr.enabled;
            chkLinkToCenter.value = !!settings.lr.linkCenter;
            txtSideWidth.text = String(settings.lr.width);
            txtSideGap.text = String(settings.lr.gap);
            txtSideInset.text = String(settings.lr.length);

            chkCenterSplit.value = !!settings.center.enabled;
            txtCenterOffset.text = String(settings.center.offset);
            rdDividerDot.value = settings.center.mode !== 'dash';
            rdDividerDash.value = settings.center.mode === 'dash';
            txtDividerWidth.text = String(settings.center.width);
            txtDividerGap.text = String(settings.center.gap);
            txtDividerInset.text = String(settings.center.length);

            rdZigzagNone.value = settings.zigzag.mode === 'none';
            rdZigzagLeftRight.value = settings.zigzag.mode === 'lr';
            rdZigzagTopBottom.value = settings.zigzag.mode === 'tb';
            txtZigzagSize.text = String(settings.zigzag.size);
            txtZigzagGap.text = String(settings.zigzag.gap);
            txtZigzagRepeat.text = String(settings.zigzag.repeat);

            rdCornerNone.value = settings.corner.mode === 'none';
            rdCornerRound.value = settings.corner.mode === 'round';
            rdCornerInverse.value = settings.corner.mode === 'inverse';
            rdCornerChamfer.value = settings.corner.mode === 'chamfer';
            txtCornerSize.text = String(settings.corner.size);

            rdHoleNone.value = settings.hole.mode === 'none';
            rdHoleCircle.value = settings.hole.mode === 'circle';
            rdHoleTriangle.value = settings.hole.mode === 'triangle';
            chkHoleLeft.value = !!settings.hole.left;
            chkHoleRight.value = !!settings.hole.right;
            txtHoleSize.text = String(settings.hole.size);

            chkEdgeDoubleRound.value = settings.edge.mode === 'wround';
            /* ダブル角丸はエッジ形状と排他なので、ラジオは「なし」に寄せる / Double rounding excludes the edge shapes */
            rdEdgeNone.value = settings.edge.mode === 'none' || settings.edge.mode === 'wround';
            rdEdgeCircle.value = settings.edge.mode === 'circle';
            rdEdgeTriangle.value = settings.edge.mode === 'triangle';
            txtEdgeSize.text = String(settings.edge.size);
            chkEdgeOnly.value = !!settings.edge.shapeOnly;

            chkPreview.value = !!settings.preview.enabled;
            chkExpandAppearance.value = !!settings.preview.expandAppearance;

            syncOffsetSliderFromInput();

            updateHoleState();
            updateCornerState();
            updateCenterState();
            updateZigzagState();
            updateSideState();
            updateEdgeState();
        }

        /**
         * 現在の設定でチケット形状を生成する
         * @param {boolean} isPreview - true ならプレビューレイヤーへ複製して適用、false なら元オブジェクトへ直接適用
         * @returns {void}
         */
        function applyEffect(isPreview) {
            var settings = readSettingsFromUI();
            var geometry = toGeometry(settings);
            /* 入力途中で数値になっていないときは描かない（確定時は呼び出し側でアラート）/ Skip while a field is still being typed */
            if (!isValidGeometry(geometry)) return;

            var targets = [];
            var results = [];
            var i;

            for (i = 0; i < selectedItems.length; i++) {
                if (isPreview) {
                    var targetLayer = ensurePreviewLayer();
                    doc.activeLayer = targetLayer;
                    var duplicated = selectedItems[i].duplicate(targetLayer, ElementPlacement.PLACEATEND);
                    normalizeInitialAppearance([duplicated], doc);
                    selectedItems[i].hidden = true;
                    targets.push(duplicated);
                } else {
                    normalizeInitialAppearance([selectedItems[i]], doc);
                    targets.push(selectedItems[i]);
                }
            }

            /* ダブル角丸のときは分割後の半片に角丸をかけるので、ここでは適用しない / Double rounding applies the radius to each half instead */
            var useDoubleRound = settings.center.enabled && settings.edge.mode === 'wround';
            if (!useDoubleRound && settings.corner.mode === 'round') {
                applyRoundCornersEffect(targets, geometry.cornerSizePt);
            }

            for (i = 0; i < targets.length; i++) {
                var rect = targets[i];
                var bounds = {
                    left: rect.left,
                    top: rect.top,
                    width: rect.width,
                    height: rect.height,
                    centerLineX: rect.left + rect.width / 2 + geometry.offsetPt
                };
                var items = [rect];

                if (useDoubleRound) {
                    applyDoubleRoundEdge(doc, rect, items, bounds, settings, geometry);
                }
                addCenterPerforation(doc, items, bounds, settings, geometry);
                addSidePerforation(doc, items, bounds, settings, geometry);
                addZigzag(doc, items, bounds, settings, geometry);
                addCenterEdges(doc, items, bounds, settings, geometry);
                addHoles(doc, items, bounds, settings, geometry);
                addInverseCorners(doc, items, bounds, settings, geometry);
                addChamferCorners(doc, items, bounds, settings, geometry);

                finalizeSubtract(doc, items, results);
            }

            if (!isPreview) {
                selectItems(doc, results.length > 0 ? results : targets);
            }
            app.redraw();
        }

        /**
         * プレビューを描き直す
         * @returns {void}
         */
        function updatePreview() {
            updateZigzagState();
            removePreview(chkPreview.value);
            if (!chkPreview.value) return;
            applyEffect(true);
        }

        // -----------------------------------------
        // ダイアログ / Dialog
        // -----------------------------------------

        var dialog = new Window('dialog', getLabel(LABELS.dialog.title) + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];

        /* プリセット行 / Preset row */
        var presetRow = addRow(dialog);
        presetRow.alignChildren = ['fill', 'center'];
        presetRow.alignment = ['fill', 'top'];
        presetRow.margins = PRESET_ROW_MARGINS;
        presetRow.spacing = PRESET_ROW_SPACING;

        var ddPreset = presetRow.add('dropdownlist', undefined, []);
        ddPreset.alignment = ['fill', 'center'];

        var btnSavePreset = presetRow.add('button', undefined, getLabel(LABELS.button.save));
        btnSavePreset.preferredSize.width = SAVE_BUTTON_SIZE[0];
        btnSavePreset.preferredSize.height = SAVE_BUTTON_SIZE[1];
        btnSavePreset.alignment = ['right', 'center'];

        /* 1行目：コーナー／ギザギザ ＋ 左右 / Row 1: corner & zigzag, then the sides */
        var topRow = addRow(dialog);
        topRow.alignChildren = ['fill', 'top'];

        var leftColumn = addColumn(topRow);

        var panelCorner = addPanel(leftColumn, getLabel(LABELS.panel.corner), ['left', 'top']);
        var rdCornerNone = panelCorner.add('radiobutton', undefined, getLabel(LABELS.radio.none));
        var rdCornerRound = panelCorner.add('radiobutton', undefined, getLabel(LABELS.radio.round));
        var rdCornerInverse = panelCorner.add('radiobutton', undefined, getLabel(LABELS.radio.inverse));
        var rdCornerChamfer = panelCorner.add('radiobutton', undefined, getLabel(LABELS.radio.chamfer));
        rdCornerNone.value = true;

        var cornerSizeField = addNumberField(panelCorner, {
            label: getLabel(LABELS.fieldLabel.size),
            value: '5',
            unit: rulerUnitLabel,
            onChange: function () { updatePreview(); }
        });
        var txtCornerSize = cornerSizeField.input;

        var panelZigzag = addPanel(leftColumn, getLabel(LABELS.panel.zigzag));

        var zigzagModeRow = addRow(panelZigzag);
        var rdZigzagNone = zigzagModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.none));
        var rdZigzagLeftRight = zigzagModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.leftRight));
        var rdZigzagTopBottom = zigzagModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.topBottom));
        rdZigzagNone.value = true;

        var zigzagLabelWidth = ZIGZAG_LABEL_WIDTH[uiLang];

        var zigzagSizeField = addNumberField(panelZigzag, {
            label: getLabel(LABELS.fieldLabel.zigzagSize),
            value: '10',
            unit: rulerUnitLabel,
            labelWidth: zigzagLabelWidth,
            onChange: function () { updatePreview(); }
        });
        var txtZigzagSize = zigzagSizeField.input;

        var zigzagRepeatField = addNumberField(panelZigzag, {
            label: getLabel(LABELS.fieldLabel.zigzagRepeat),
            value: '3',
            labelWidth: zigzagLabelWidth,
            onChange: function () { updateZigzagState(); updatePreview(); }
        });
        var txtZigzagRepeat = zigzagRepeatField.input;

        var zigzagGapField = addNumberField(panelZigzag, {
            label: getLabel(LABELS.fieldLabel.gap),
            value: '0',
            unit: rulerUnitLabel,
            labelWidth: zigzagLabelWidth,
            allowNegative: true,
            onChange: function () { updatePreview(); }
        });
        var txtZigzagGap = zigzagGapField.input;

        /* 左右（ミシン目＋スリット／ホール）/ Sides: perforation and slit / hole */
        var panelSides = addPanel(topRow, getLabel(LABELS.panel.sides));

        var panelSidePerforation = addPanel(panelSides, getLabel(LABELS.panel.sidePerforation));

        var sideEnableRow = addRow(panelSidePerforation);
        var chkSidePerforation = sideEnableRow.add('checkbox', undefined, getLabel(LABELS.checkbox.enable));
        var chkLinkToCenter = sideEnableRow.add('checkbox', undefined, getLabel(LABELS.checkbox.linkToCenter));
        chkLinkToCenter.value = true;

        var sideWidthField = addNumberField(panelSidePerforation, {
            label: getLabel(LABELS.fieldLabel.lineWidth),
            value: '3',
            unit: rulerUnitLabel,
            onChange: function () { updatePreview(); }
        });
        var txtSideWidth = sideWidthField.input;

        var sideGapField = addNumberField(panelSidePerforation, {
            label: getLabel(LABELS.fieldLabel.gap),
            value: '6',
            unit: rulerUnitLabel,
            onChange: function () { updatePreview(); }
        });
        var txtSideGap = sideGapField.input;

        var sideInsetField = addNumberField(panelSidePerforation, {
            label: getLabel(LABELS.fieldLabel.inset),
            value: '0',
            unit: rulerUnitLabel,
            characters: INSET_FIELD_CHARS,
            allowNegative: true,
            negativeOnly: true,
            onChange: function () { updatePreview(); }
        });
        var txtSideInset = sideInsetField.input;

        var panelHole = addPanel(panelSides, getLabel(LABELS.panel.hole), ['left', 'top']);

        var holeModeRow = addRow(panelHole);
        var rdHoleNone = holeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.none));
        var rdHoleCircle = holeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.circle));
        var rdHoleTriangle = holeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.triangle));
        rdHoleNone.value = true;

        var holeSideRow = addRow(panelHole);
        var chkHoleLeft = holeSideRow.add('checkbox', undefined, getLabel(LABELS.checkbox.left));
        var chkHoleRight = holeSideRow.add('checkbox', undefined, getLabel(LABELS.checkbox.right));
        chkHoleLeft.value = true;
        chkHoleRight.value = true;

        var holeSizeField = addNumberField(panelHole, {
            label: getLabel(LABELS.fieldLabel.size),
            value: '10',
            unit: rulerUnitLabel,
            onChange: function () { updatePreview(); }
        });
        var txtHoleSize = holeSizeField.input;

        /* 左右分割（分割線＋エッジ）/ Center split: divider line and edges */
        var panelCenter = addPanel(dialog, getLabel(LABELS.panel.centerSplit));

        var centerTopRow = addRow(panelCenter);
        centerTopRow.alignChildren = ['center', 'center'];

        var chkCenterSplit = centerTopRow.add('checkbox', undefined, getLabel(LABELS.checkbox.enable));
        chkCenterSplit.value = true;
        var txtCenterOffset = centerTopRow.add('edittext', undefined, '0');
        txtCenterOffset.characters = OFFSET_FIELD_CHARS;
        var lblCenterOffsetUnit = centerTopRow.add('statictext', undefined, rulerUnitLabel);

        /* 複数選択時は最小幅の半分を上限にして、すべての対象で安全な範囲に制限 / Clamp to the narrowest item so every target stays valid */
        var minHalfWidthPt = null;
        for (var itemIndex = 0; itemIndex < selectedItems.length; itemIndex++) {
            var halfWidthPt = Math.abs(selectedItems[itemIndex].width) / 2;
            if (minHalfWidthPt === null || halfWidthPt < minHalfWidthPt) minHalfWidthPt = halfWidthPt;
        }
        var maxOffset = Math.round(fromPt(minHalfWidthPt || 0, rulerPtFactor) * 10) / 10;

        var sliderCenterOffset = centerTopRow.add('slider', undefined, 0, -maxOffset, maxOffset);
        sliderCenterOffset.preferredSize.width = OFFSET_SLIDER_WIDTH;

        var centerBodyRow = addRow(panelCenter);
        centerBodyRow.alignChildren = ['fill', 'top'];

        var panelDivider = addPanel(centerBodyRow, getLabel(LABELS.panel.divider));

        var dividerModeRow = addRow(panelDivider);
        var rdDividerDot = dividerModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.dot));
        var rdDividerDash = dividerModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.dash));
        rdDividerDot.value = true;

        var dividerWidthField = addNumberField(panelDivider, {
            label: getLabel(LABELS.fieldLabel.lineWidth),
            value: '3',
            unit: strokeUnitLabel,
            onChange: function () { syncSideFromCenter(); updatePreview(); }
        });
        var txtDividerWidth = dividerWidthField.input;

        var dividerGapField = addNumberField(panelDivider, {
            label: getLabel(LABELS.fieldLabel.gap),
            value: '6',
            unit: rulerUnitLabel,
            onChange: function () { syncSideFromCenter(); updatePreview(); }
        });
        var txtDividerGap = dividerGapField.input;

        var dividerInsetField = addNumberField(panelDivider, {
            label: getLabel(LABELS.fieldLabel.inset),
            value: '0',
            unit: rulerUnitLabel,
            characters: INSET_FIELD_CHARS,
            allowNegative: true,
            negativeOnly: true,
            onChange: function () { syncSideFromCenter(); updatePreview(); }
        });
        var txtDividerInset = dividerInsetField.input;

        var panelEdge = addPanel(centerBodyRow, getLabel(LABELS.panel.edge), ['left', 'top']);

        var edgeModeRow = addRow(panelEdge);
        var rdEdgeNone = edgeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.none));
        var rdEdgeCircle = edgeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.circle));
        var rdEdgeTriangle = edgeModeRow.add('radiobutton', undefined, getLabel(LABELS.radio.triangle));
        rdEdgeNone.value = true;

        var chkEdgeDoubleRound = panelEdge.add('checkbox', undefined, getLabel(LABELS.checkbox.doubleRound));

        var edgeSizeField = addNumberField(panelEdge, {
            label: getLabel(LABELS.fieldLabel.size),
            value: '10',
            unit: rulerUnitLabel,
            onChange: function () { syncDividerInsetFromEdgeSize(); updatePreview(); }
        });
        var txtEdgeSize = edgeSizeField.input;

        var chkEdgeOnly = panelEdge.add('checkbox', undefined, getLabel(LABELS.checkbox.edgesOnly));

        /* プレビュー行 / Preview row */
        var previewRow = addRow(dialog);
        previewRow.alignment = ['center', 'top'];

        var chkPreview = previewRow.add('checkbox', undefined, getLabel(LABELS.checkbox.preview));
        chkPreview.value = true;

        var chkExpandAppearance = previewRow.add('checkbox', undefined, getLabel(LABELS.checkbox.expandAppearance));
        chkExpandAppearance.value = false;

        /* ボタン行 / Button row */
        var buttonRow = addRow(dialog);
        buttonRow.alignChildren = ['fill', 'center'];

        var buttonLeft = addRow(buttonRow);
        var isOutlineMode = false;
        var btnOutlineToggle = buttonLeft.add('button', undefined, getLabel(LABELS.button.outlineOn));

        var buttonRight = addRow(buttonRow);
        buttonRight.alignment = ['right', 'center'];
        buttonRight.alignChildren = ['right', 'center'];
        buttonRight.add('button', undefined, getLabel(LABELS.button.cancel), { name: 'cancel' });
        buttonRight.add('button', undefined, getLabel(LABELS.button.ok), { name: 'ok' });

        // -----------------------------------------
        // ディム制御と連動 / Enable state and syncing
        // -----------------------------------------

        /**
         * 入力欄の値からスライダー位置を合わせる
         * @returns {void}
         */
        function syncOffsetSliderFromInput() {
            var offsetValue = parseFloat(txtCenterOffset.text);
            if (isNaN(offsetValue)) return;
            sliderCenterOffset.value = Math.max(-maxOffset, Math.min(maxOffset, offsetValue));
        }

        /**
         * 連動が有効なとき、左右のミシン目に分割線の値をコピーする
         * @returns {void}
         */
        function syncSideFromCenter() {
            if (!chkLinkToCenter.value || !rdDividerDot.value) return;
            txtSideWidth.text = txtDividerWidth.text;
            txtSideGap.text = txtDividerGap.text;
            txtSideInset.text = txtDividerInset.text;
        }

        /**
         * エッジのサイズに合わせて分割線の長さ（食い込み量）を決める
         * @returns {void}
         */
        function syncDividerInsetFromEdgeSize() {
            if (!chkCenterSplit.value) return;
            if (rdEdgeNone.value && !chkEdgeDoubleRound.value) return;

            var edgeSizeValue = parseFloat(txtEdgeSize.text);
            if (isNaN(edgeSizeValue)) return;

            var ratio = 1.2;
            if (rdEdgeCircle.value) ratio = 1.0;
            else if (chkEdgeDoubleRound.value) ratio = 2.0;

            txtDividerInset.text = String(-(Math.round(edgeSizeValue * ratio * 1000) / 1000));
        }

        /**
         * 左右のミシン目まわりのディムを更新する
         * @returns {void}
         */
        function updateSideState() {
            var isOn = chkSidePerforation.value;
            var isLinked = isOn && chkLinkToCenter.value && rdDividerDot.value;
            chkLinkToCenter.enabled = isOn;
            sideWidthField.row.enabled = isOn && !isLinked;
            sideGapField.row.enabled = isOn && !isLinked;
            sideInsetField.row.enabled = isOn && !isLinked;
            if (isLinked) syncSideFromCenter();
        }

        /**
         * ギザギザまわりのディムを更新する
         * @returns {void}
         */
        function updateZigzagState() {
            var isOn = !rdZigzagNone.value;
            zigzagSizeField.row.enabled = isOn;
            zigzagRepeatField.row.enabled = isOn;
            zigzagGapField.row.enabled = isOn && Math.round(parseFloat(txtZigzagRepeat.text) || 0) > 1;
        }

        /**
         * スリット／ホールまわりのディムを更新する
         * @returns {void}
         */
        function updateHoleState() {
            var isOn = !rdHoleNone.value;
            holeSideRow.enabled = isOn;
            holeSizeField.row.enabled = isOn;
        }

        /**
         * コーナーまわりのディムを更新する
         * @returns {void}
         */
        function updateCornerState() {
            var isDoubleRound = chkEdgeDoubleRound.value;
            panelCorner.enabled = !isDoubleRound;
            cornerSizeField.row.enabled = !isDoubleRound && !rdCornerNone.value;
        }

        /**
         * 左右分割まわりのディムを更新する
         * @returns {void}
         */
        function updateCenterState() {
            var isOn = chkCenterSplit.value;
            txtCenterOffset.enabled = isOn;
            lblCenterOffsetUnit.enabled = isOn;
            sliderCenterOffset.enabled = isOn;
            centerBodyRow.enabled = isOn;
        }

        /**
         * エッジまわりのディムを更新する
         * @returns {void}
         */
        function updateEdgeState() {
            var isDoubleRound = chkCenterSplit.value && chkEdgeDoubleRound.value;
            var isOn = (chkCenterSplit.value && !rdEdgeNone.value) || isDoubleRound;
            edgeModeRow.enabled = !isDoubleRound;
            edgeSizeField.row.enabled = isOn;
            chkEdgeOnly.enabled = isOn;
            panelDivider.enabled = chkCenterSplit.value && !(isOn && chkEdgeOnly.value);
            if (isOn) syncDividerInsetFromEdgeSize();
        }

        // -----------------------------------------
        // イベント / Events
        // -----------------------------------------

        ddPreset.onChange = function () {
            var index = ddPreset.selection ? ddPreset.selection.index : -1;
            if (index < 0 || index >= PRESETS.length) return;
            applySettingsToUI(cloneSimpleValue(PRESETS[index]));
            updatePreview();
        };

        btnSavePreset.onClick = function () {
            var baseName = getDropdownSelectionText(ddPreset) || getLabel(LABELS.fallbackName.customPreset);
            var newName = prompt(getLabel(LABELS.alert.presetName), baseName);
            if (newName === null) return;
            newName = String(newName).replace(/^\s+|\s+$/g, '');
            if (!newName) return;

            var preset = readSettingsFromUI(newName);
            if (!isValidGeometry(toGeometry(preset))) {
                alert(getLabel(LABELS.alert.enterValidNumbers));
                return;
            }

            PRESETS.push(cloneSimpleValue(preset));
            setDropdownItemsFromPresets(ddPreset, PRESETS);
            ddPreset.selection = ddPreset.items.length - 1;

            alert(getLabel(LABELS.alert.presetSaved) + presetToCode(preset));
        };

        chkSidePerforation.onClick = function () {
            /* 左右のミシン目とギザギザ（左右）は同じ辺を使うので併用しない / Side perforation and L/R zigzag share the same edges */
            if (chkSidePerforation.value && rdZigzagLeftRight.value) {
                rdZigzagNone.value = true;
                rdZigzagLeftRight.value = false;
                updateZigzagState();
            }
            updateSideState();
            updatePreview();
        };
        chkLinkToCenter.onClick = function () { updateSideState(); updatePreview(); };

        rdZigzagNone.onClick = function () { updateZigzagState(); updatePreview(); };
        rdZigzagLeftRight.onClick = function () { updateZigzagState(); updatePreview(); };
        rdZigzagTopBottom.onClick = function () { updateZigzagState(); updatePreview(); };

        rdHoleNone.onClick = function () { updateHoleState(); updatePreview(); };
        rdHoleCircle.onClick = function () { updateHoleState(); updatePreview(); };
        rdHoleTriangle.onClick = function () { updateHoleState(); updatePreview(); };
        chkHoleLeft.onClick = function () { updatePreview(); };
        chkHoleRight.onClick = function () { updatePreview(); };

        rdCornerNone.onClick = function () { updateCornerState(); updatePreview(); };
        rdCornerRound.onClick = function () { updateCornerState(); updatePreview(); };
        rdCornerInverse.onClick = function () { updateCornerState(); updatePreview(); };
        rdCornerChamfer.onClick = function () { updateCornerState(); updatePreview(); };

        chkCenterSplit.onClick = function () { updateCenterState(); updatePreview(); };
        rdDividerDot.onClick = function () { updateSideState(); updatePreview(); };
        rdDividerDash.onClick = function () { updateSideState(); updatePreview(); };

        rdEdgeNone.onClick = function () { updateEdgeState(); updatePreview(); };
        rdEdgeCircle.onClick = function () { updateEdgeState(); updatePreview(); };
        rdEdgeTriangle.onClick = function () { updateEdgeState(); updatePreview(); };
        chkEdgeOnly.onClick = function () { updateEdgeState(); updatePreview(); };

        chkEdgeDoubleRound.onClick = function () {
            if (chkEdgeDoubleRound.value) {
                /* ダブル角丸はコーナー処理・エッジ形状と併用しない / Double rounding replaces the corner and edge shapes */
                rdCornerNone.value = true;
                rdCornerRound.value = false;
                rdCornerInverse.value = false;
                rdCornerChamfer.value = false;
                rdEdgeNone.value = true;
                rdEdgeCircle.value = false;
                rdEdgeTriangle.value = false;
                txtEdgeSize.text = '5';
            } else if (rdCornerNone.value) {
                /* OFFに戻したときは、強制的に「なし」にしていたコーナーを角丸へ戻す / Restore the corner mode forced to "none" */
                rdCornerNone.value = false;
                rdCornerRound.value = true;
            }
            updateCornerState();
            syncDividerInsetFromEdgeSize();
            updateEdgeState();
            updatePreview();
        };

        chkPreview.onClick = function () { updatePreview(); };

        txtCenterOffset.onChanging = function () {
            syncOffsetSliderFromInput();
            updatePreview();
        };
        sliderCenterOffset.onChanging = function () {
            txtCenterOffset.text = String(Math.round(sliderCenterOffset.value * 10) / 10);
            updatePreview();
        };
        bindArrowKeyStep(txtCenterOffset, {
            allowNegative: true,
            onChange: function () { syncOffsetSliderFromInput(); updatePreview(); }
        });

        txtSideWidth.onChanging = updatePreview;
        txtSideGap.onChanging = updatePreview;
        txtSideInset.onChanging = function () { clampToNegative(txtSideInset); updatePreview(); };
        txtZigzagSize.onChanging = updatePreview;
        txtZigzagGap.onChanging = updatePreview;
        txtZigzagRepeat.onChanging = function () { updateZigzagState(); updatePreview(); };
        txtCornerSize.onChanging = updatePreview;
        txtHoleSize.onChanging = updatePreview;
        txtDividerWidth.onChanging = function () { syncSideFromCenter(); updatePreview(); };
        txtDividerGap.onChanging = function () { syncSideFromCenter(); updatePreview(); };
        txtDividerInset.onChanging = function () {
            clampToNegative(txtDividerInset);
            syncSideFromCenter();
            updatePreview();
        };
        txtEdgeSize.onChanging = function () { syncDividerInsetFromEdgeSize(); updatePreview(); };

        btnOutlineToggle.onClick = function () {
            try {
                app.executeMenuCommand('preview');
                isOutlineMode = !isOutlineMode;
                btnOutlineToggle.text = isOutlineMode ? getLabel(LABELS.button.outlineOff) : getLabel(LABELS.button.outlineOn);
            } catch (e) {
                /* 表示モードを切り替えられなくても続行 / Continue even if the view cannot be toggled */
            }
        };

        /**
         * 入力欄の値が正なら0に丸める（食い込み量は0以下のみ）
         * @param {EditText} input - 対象の入力欄
         * @returns {void}
         */
        function clampToNegative(input) {
            var value = parseFloat(input.text);
            if (!isNaN(value) && value > 0) input.text = '0';
        }

        // -----------------------------------------
        // 表示と確定 / Show and apply
        // -----------------------------------------

        setDropdownItemsFromPresets(ddPreset, PRESETS);
        if (PRESETS.length > 0) {
            applySettingsToUI(cloneSimpleValue(PRESETS[0]));
        }
        syncDividerInsetFromEdgeSize();
        updatePreview();

        var isConfirmed = (dialog.show() === 1);
        removePreview();
        removePreviewLayer();

        if (!isConfirmed) return;

        if (!isValidGeometry(toGeometry(readSettingsFromUI()))) {
            alert(getLabel(LABELS.alert.enterValidNumbers));
            return;
        }

        doc.activeLayer = originalLayer;
        applyEffect(false);
        if (chkExpandAppearance.value) {
            app.executeMenuCommand('expandStyle');
        }
    } finally {
        unloadStrokeDotAction();
        app.executeMenuCommand('edge');
        app.executeMenuCommand('AI Bounding Box Toggle');
    }
}

main();
