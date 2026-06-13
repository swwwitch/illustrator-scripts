#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

スクリプト名：

SwapStyle.jsx — 2つのオブジェクト間でスタイル／文字列を交換

概要：

* 選択した2つのオブジェクトの間で、見た目または文字列を交換
* ダイアログで「スタイル交換」「文字列交換」を切替
* スタイル交換は次の3系統を組み合わせ可能
  - グラフィックスタイル交換（現在のアピアランス全体）
  - 基本的な塗りや線（塗り／線のカラー／線幅）
  - 文字書式（フォントとスタイル／フォントサイズ）
* 基本的な塗りや線のいずれかが ON のときは、グラフィックスタイルパネルをディム表示し、グラフィックスタイル交換は実行しない（一時アクションの読み込みもスキップ）
* 文字書式・基本的な塗りや線は、誤適用を避けるため初期 OFF
* TextFrame 同士の場合、基本的な塗りや線は `textRange.characterAttributes` 経由で交換（線幅は `strokeWeight`）
* グラフィックスタイル交換は、現在のアピアランスを一時スタイル（_swapStyleA〜_swapStyleZ）として登録し、相手側へクロス適用
* 登録した一時スタイルは、必要に応じて終了時に削除（適用済みのアピアランスは各オブジェクトに残る）

主な機能：

* 交換方式（スタイル／文字列）の切替
* テキスト同士の場合のみ文字列交換を許可
* グラフィックスタイル交換の ON/OFF と一時スタイルの削除可否
* 基本的な塗りや線（塗り／線のカラー／線幅）の個別 ON/OFF
* 文字書式（フォントとスタイル／フォントサイズ）の個別 ON/OFF
* 基本のいずれかが ON のときは、グラフィックスタイル交換を自動的に OFF として扱う
* 交換用の一時スタイル名として _swapStyleA / _swapStyleB から順に未使用ペアを使用
* 「AddNewWithoutName」アクションを一時ファイル経由でロードし、現在のアピアランスを無名スタイルとして登録

処理の流れ：

1. 選択された2つのオブジェクトを保持し、ダイアログで交換方式と対象項目を選択
2. 文字列交換のときは、テキストフレーム同士で contents を入れ替え
3. スタイル交換では、グラフィックスタイル交換 → 基本的な塗りや線 → 文字書式 の順に、ON の項目だけクロス入れ替え
4. グラフィックスタイル交換では、現在のアピアランスを _swapStyleX / _swapStyleY として一時登録し、_swapStyleX を相手側、_swapStyleY をこちら側にクロス適用
5. 基本的な塗りや線は、TextFrame 同士なら characterAttributes 経由、それ以外は PageItem の fillColor / strokeColor / strokeWidth で交換
6. 設定に応じて、登録した一時スタイルを削除し、選択を元の2つに戻す

note：

* 文字列交換と文字書式交換は、どちらもテキストオブジェクト同士の場合にのみ有効
* 基本的な塗りや線の交換は、TextFrame 同士なら characterAttributes 経由（線幅は strokeWeight）、それ以外は PageItem レベル
* 基本的な塗りや線のいずれかが ON のときは、グラフィックスタイル交換は実行されない（パネルもディム表示）
* スタイル名の空き（_swapStyleA〜_swapStyleZ の連続ペア）がない場合は中断
* 一時スタイルの登録に失敗した場合は、登録済みの一時スタイルを掃除してから中断
* Illustrator の仕様上、最後の1件など完全削除できないケースがあります

更新履歴：

* v1.0.0 (2026-05-21) : 初期バージョン
* v1.1.0 (2026-05-23) : ダイアログにプレビューを追加

⸻

Script Name:

SwapStyle.jsx — Swap Style / Content between Two Objects

Overview:

* Swap appearance or text content between two selected objects
* Toggle between “Style Swap” and “Content Swap” in the dialog
* Style Swap can combine up to three subsystems:
  - Graphic-style swap (full current appearance)
  - Basic fill & stroke (fill / stroke color / stroke width)
  - Text formatting (font & style / font size)
* When any Basic Fill & Stroke checkbox is ON, the Graphic Style panel is dimmed and graphic-style swap is skipped (the temporary action is not loaded either)
* Text formatting and Basic Fill & Stroke options are OFF by default to avoid unintended changes
* For two TextFrames, Basic Fill & Stroke is swapped via `textRange.characterAttributes` (stroke width uses `strokeWeight`)
* Graphic-style swap registers each object’s current appearance as a temporary style (_swapStyleA–_swapStyleZ), then applies them crosswise
* Registered temporary styles can be removed on exit; the applied appearance remains on each object

Key Features:

* Mode toggle (Style / Content)
* Content swap is enabled only when both selected objects are text frames
* Graphic-style swap ON/OFF and optional deletion of registered temporary styles
* Individual Basic Fill & Stroke toggles for Fill / Stroke Color / Stroke Width
* Individual text-format toggles for Font & Style and Font Size
* When any Basic Fill & Stroke checkbox is ON, graphic-style swap is automatically treated as OFF
* Uses the next available temporary pair, starting with _swapStyleA / _swapStyleB
* Loads the “AddNewWithoutName” action via a temporary file to register the current appearance as an unnamed style

Process Flow:

1. Keep the two selected objects and choose mode / target items in the dialog
2. In Content Swap mode, swap contents between two text frames
3. In Style Swap mode, cross-swap only the enabled items in this order: graphic style → basic fill & stroke → text formatting
4. Graphic-style swap registers current appearances as temporary _swapStyleX / _swapStyleY, then applies _swapStyleX to the other object and _swapStyleY back
5. Basic Fill & Stroke uses characterAttributes for two TextFrames, otherwise PageItem-level fillColor / strokeColor / strokeWidth
6. Remove registered temporary styles according to the setting, then restore selection to the original two objects

note:

* Content swap and text-format swaps are available only between text objects
* Basic Fill & Stroke uses characterAttributes (strokeWeight) for two TextFrames, otherwise PageItem-level properties
* When any Basic Fill & Stroke checkbox is ON, graphic-style swap is not executed (the panel is also dimmed)
* Aborts if no free style-name pair (_swapStyleA–_swapStyleZ) is available
* If temporary style registration fails, already registered temporary styles are cleaned up before aborting
* Due to Illustrator specs, some entries (e.g. the last one) may not be fully removable

Changelog:

* v1.0.0 (2026-05-21): Initial release
* v1.1.0 (2026-05-23): Add preview to the dialog

*/

// =========================================
// バージョンとローカライズ / Version & Localization
// =========================================

var SCRIPT_VERSION = "v1.1.0";

/* 現在の言語を返す / Return the current language */
function getCurrentLanguage() {
    return ($.locale && $.locale.indexOf("ja") === 0) ? "ja" : "en";
}

var currentLanguage = getCurrentLanguage();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    // UI labels
    dialogTitle: {
        ja: "スタイル・文字列の交換",
        en: "Swap Style / Content"
    },
    modePanel: {
        ja: "交換方式",
        en: "Swap Mode"
    },
    modeStyle: {
        ja: "スタイル",
        en: "Style"
    },
    modeContent: {
        ja: "文字列",
        en: "Content"
    },
    modeCoordinate: {
        ja: "座標",
        en: "Position"
    },
    formatPanel: {
        ja: "文字書式",
        en: "Text Formatting"
    },
    fontAndStyle: {
        ja: "フォントとスタイル",
        en: "Font & Style"
    },
    fontSize: {
        ja: "フォントサイズ",
        en: "Font Size"
    },
    basicFillStrokePanel: {
        ja: "基本的な塗りや線",
        en: "Basic Fill & Stroke"
    },
    swapFill: {
        ja: "塗り",
        en: "Fill"
    },
    swapStrokeColor: {
        ja: "線のカラー",
        en: "Stroke Color"
    },
    swapStrokeWidth: {
        ja: "線幅",
        en: "Stroke Width"
    },
    graphicStylePanel: {
        ja: "グラフィックスタイル",
        en: "Graphic Style"
    },
    coordinatePanel: {
        ja: "座標",
        en: "Position"
    },
    tipAnchor: {
        ja: "選択した基準点で2つのオブジェクトの座標を交換します。",
        en: "Swap positions of the two objects by the selected reference point."
    },
    axisLabel: {
        ja: "交換する軸",
        en: "Axis"
    },
    axisBoth: {
        ja: "両方",
        en: "Both"
    },
    axisX: {
        ja: "横（X）",
        en: "X"
    },
    axisY: {
        ja: "縦（Y）",
        en: "Y"
    },
    swapZOrder: {
        ja: "重ね順",
        en: "Stacking order"
    },
    swapGraphicStyles: {
        ja: "交換する",
        en: "Swap"
    },
    deleteStyles: {
        ja: "一時スタイルを削除",
        en: "Delete Temporary Styles"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    preview: {
        ja: "プレビュー",
        en: "Preview"
    },

    // Tooltips
    tipContentSwap: {
        ja: "文字列交換はテキストオブジェクト同士でのみ実行できます。",
        en: "Content swap is available only between text objects."
    },
    tipDeleteStyles: {
        ja: "削除しても適用済みのアピアランスは各オブジェクトに残ります。",
        en: "Deleting styles does not remove the applied appearance from objects."
    },
    tipGraphicStyle: {
        ja: "現在のアピアランスを一時スタイルとして登録し、相互に適用します。",
        en: "Registers current appearances as temporary styles and applies them crosswise."
    },

    // Alerts
    alertNoDocument: {
        ja: "ドキュメントが開かれていません。",
        en: "No document is open."
    },
    alertSelectTwo: {
        ja: "オブジェクトをちょうど2つ選択してください。",
        en: "Please select exactly two objects."
    },
    alertContentTextOnly: {
        ja: "文字列交換はテキストオブジェクト同士でのみ実行できます。",
        en: "Content swap requires both objects to be text frames."
    },
    alertNoFreePair: {
        ja: "使用できるスタイル名（_swapStyleA〜_swapStyleZ）の空きがありません。",
        en: "No free style name pair available (_swapStyleA–_swapStyleZ)."
    },
    alertSwapFailed: {
        ja: "スタイルの登録に失敗しました。処理を中止します。",
        en: "Graphic style registration failed. Aborting."
    }
};

/* ラベルを取得（未定義時は英語またはキー名にフォールバック） / Get a localized label with fallback */
function getLabel(key) {
    if (LABELS[key] && LABELS[key][currentLanguage]) return LABELS[key][currentLanguage];
    if (LABELS[key] && LABELS[key].en) return LABELS[key].en;
    return key;
}

// =========================================
// ダイアログ / Dialog
// =========================================

var PANEL_MARGINS = [15, 20, 15, 10];
var PANEL_SPACING = 8;

/* パネルの共通設定 / Common panel setup */
function setupPanel(panel, spacing) {
    panel.orientation = "column";
    panel.alignChildren = ['fill', 'top'];
    panel.alignment = "fill";
    panel.margins = PANEL_MARGINS;
    panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
}

/* 交換方式パネルを作成 / Build the swap mode panel */
function buildModePanel(dialog, bothAreText) {
    var modePanel = dialog.add('panel', undefined, getLabel('modePanel'));
    setupPanel(modePanel);
    modePanel.orientation = "row";
    modePanel.alignChildren = ['center', 'center'];

    var contentSwapRadio = modePanel.add('radiobutton', undefined, getLabel('modeContent'));
    var styleSwapRadio = modePanel.add('radiobutton', undefined, getLabel('modeStyle'));
    var coordinateSwapRadio = modePanel.add('radiobutton', undefined, getLabel('modeCoordinate'));
    contentSwapRadio.enabled = bothAreText;
    contentSwapRadio.helpTip = getLabel('tipContentSwap');
    styleSwapRadio.value = true;

    return {
        panel: modePanel,
        styleSwapRadio: styleSwapRadio,
        contentSwapRadio: contentSwapRadio,
        coordinateSwapRadio: coordinateSwapRadio
    };
}

/* 文字書式パネルを作成 / Build the text formatting panel */
function buildFormatPanel(dialog) {
    var formatPanel = dialog.add('panel', undefined, getLabel('formatPanel'));
    setupPanel(formatPanel);

    var fontAndStyleCheckbox = formatPanel.add('checkbox', undefined, getLabel('fontAndStyle'));
    var fontSizeCheckbox = formatPanel.add('checkbox', undefined, getLabel('fontSize'));
    fontAndStyleCheckbox.value = false;
    fontSizeCheckbox.value = false;

    return {
        panel: formatPanel,
        fontAndStyleCheckbox: fontAndStyleCheckbox,
        fontSizeCheckbox: fontSizeCheckbox
    };
}

/* 基本的な塗りや線パネルを作成 / Build the basic fill & stroke panel */
function buildBasicFillStrokePanel(dialog) {
    var basicFillStrokePanel = dialog.add('panel', undefined, getLabel('basicFillStrokePanel'));
    setupPanel(basicFillStrokePanel);

    var fillCheckbox = basicFillStrokePanel.add('checkbox', undefined, getLabel('swapFill'));
    var strokeColorCheckbox = basicFillStrokePanel.add('checkbox', undefined, getLabel('swapStrokeColor'));
    var strokeWidthCheckbox = basicFillStrokePanel.add('checkbox', undefined, getLabel('swapStrokeWidth'));
    fillCheckbox.value = false;
    strokeColorCheckbox.value = false;
    strokeWidthCheckbox.value = false;

    return {
        panel: basicFillStrokePanel,
        fillCheckbox: fillCheckbox,
        strokeColorCheckbox: strokeColorCheckbox,
        strokeWidthCheckbox: strokeWidthCheckbox
    };
}

/* グラフィックスタイルパネルを作成 / Build the graphic style panel */
function buildGraphicStylePanel(dialog) {
    var graphicStylePanel = dialog.add('panel', undefined, getLabel('graphicStylePanel'));
    setupPanel(graphicStylePanel);
    graphicStylePanel.helpTip = getLabel('tipGraphicStyle');

    var swapGraphicStylesCheckbox = graphicStylePanel.add('checkbox', undefined, getLabel('swapGraphicStyles'));
    var deleteStylesCheckbox = graphicStylePanel.add('checkbox', undefined, getLabel('deleteStyles'));
    swapGraphicStylesCheckbox.value = true;
    deleteStylesCheckbox.value = true;
    deleteStylesCheckbox.helpTip = getLabel('tipDeleteStyles');

    return {
        panel: graphicStylePanel,
        swapGraphicStylesCheckbox: swapGraphicStylesCheckbox,
        deleteStylesCheckbox: deleteStylesCheckbox
    };
}

/* 座標パネルを作成（3×3 の基準点ラジオ） / Build the position panel (3x3 reference-point radios) */
function buildCoordinatePanel(parent) {
    var coordinatePanel = parent.add('panel', undefined, getLabel('coordinatePanel'));
    setupPanel(coordinatePanel);
    coordinatePanel.alignChildren = ['center', 'top'];
    coordinatePanel.helpTip = getLabel('tipAnchor');

    var grid = coordinatePanel.add('group');
    grid.orientation = 'column';
    grid.spacing = 4;
    grid.alignChildren = ['center', 'center'];

    // 行優先（row-major）の 9 個。index 0=左上 … 4=中央 … 8=右下
    var anchorRadios = [];
    for (var row = 0; row < 3; row++) {
        var rowGroup = grid.add('group');
        rowGroup.orientation = 'row';
        rowGroup.spacing = 4;
        for (var col = 0; col < 3; col++) {
            anchorRadios.push(rowGroup.add('radiobutton', undefined, ''));
        }
    }
    anchorRadios[4].value = true; // 中央を初期値

    // 軸ロック（両方／横X／縦Y）。同一グループ内なので自動排他
    var axisGroup = coordinatePanel.add('group');
    axisGroup.orientation = 'column';
    axisGroup.alignChildren = ['left', 'center'];
    axisGroup.spacing = 6;
    axisGroup.add('statictext', undefined, getLabel('axisLabel'));
    var axisBothRadio = axisGroup.add('radiobutton', undefined, getLabel('axisBoth'));
    var axisXRadio = axisGroup.add('radiobutton', undefined, getLabel('axisX'));
    var axisYRadio = axisGroup.add('radiobutton', undefined, getLabel('axisY'));
    axisBothRadio.value = true;

    // 重なり順（前後関係）も交換
    var zOrderCheckbox = coordinatePanel.add('checkbox', undefined, getLabel('swapZOrder'));
    zOrderCheckbox.value = false;

    return {
        panel: coordinatePanel,
        anchorRadios: anchorRadios,
        axisBothRadio: axisBothRadio,
        axisXRadio: axisXRadio,
        axisYRadio: axisYRadio,
        zOrderCheckbox: zOrderCheckbox
    };
}

/* 選択中の基準点インデックス（0〜8、未選択時は中央=4） / Selected reference-point index */
function getSelectedAnchorIndex(anchorRadios) {
    for (var i = 0; i < anchorRadios.length; i++) {
        if (anchorRadios[i].value) return i;
    }
    return 4;
}

/* フッター行を作成（左：プレビュー／中央：スペーサー／右：ボタン） / Build the footer row (left: preview / center: spacer / right: buttons) */
function buildFooter(dialog) {
    var footer = dialog.add('group');
    footer.orientation = 'row';
    footer.alignment = 'fill';
    footer.alignChildren = ['fill', 'center'];

    // 左：プレビュー
    var footerLeft = footer.add('group');
    footerLeft.alignment = ['left', 'center'];
    var previewCheckbox = footerLeft.add('checkbox', undefined, getLabel('preview'));
    previewCheckbox.value = false;

    // 中央：スペーサー（残り幅を吸収して左右を引き離す）
    var footerCenter = footer.add('group');
    footerCenter.alignment = ['fill', 'center'];

    // 右：キャンセル／OK
    var footerRight = footer.add('group');
    footerRight.alignment = ['right', 'center'];
    var cancelButton = footerRight.add('button', undefined, getLabel('cancel'), { name: 'cancel' });
    var okButton = footerRight.add('button', undefined, 'OK', { name: 'ok' });

    return {
        group: footer,
        previewCheckbox: previewCheckbox,
        cancelButton: cancelButton,
        okButton: okButton
    };
}

/* ダイアログの有効／無効状態を更新 / Update dialog enabled states */
function updateDialogEnabled(ui, bothAreText) {
    var isStyleSwap = ui.mode.styleSwapRadio.value;
    var isCoordinateSwap = ui.mode.coordinateSwapRadio.value;
    var anyBasicOn = ui.basicFillStroke.fillCheckbox.value
        || ui.basicFillStroke.strokeColorCheckbox.value
        || ui.basicFillStroke.strokeWidthCheckbox.value;
    ui.format.panel.enabled = isStyleSwap && bothAreText;
    ui.basicFillStroke.panel.enabled = isStyleSwap;
    ui.graphicStyle.panel.enabled = isStyleSwap && !anyBasicOn;
    // 座標パネルは「座標」選択時のみ有効（それ以外はディム）
    ui.coordinate.panel.enabled = isCoordinateSwap;
    if (isStyleSwap) {
        ui.graphicStyle.deleteStylesCheckbox.enabled = ui.graphicStyle.swapGraphicStylesCheckbox.value;
    }
}

/* ダイアログの選択内容を読む / Read dialog choices */
function readDialogOptions(ui) {
    var anyBasicOn = ui.basicFillStroke.fillCheckbox.value
        || ui.basicFillStroke.strokeColorCheckbox.value
        || ui.basicFillStroke.strokeWidthCheckbox.value;
    return {
        mode: ui.mode.styleSwapRadio.value ? 'style'
            : (ui.mode.coordinateSwapRadio.value ? 'coordinate' : 'content'),
        swapGraphicStyles: ui.graphicStyle.swapGraphicStylesCheckbox.value && !anyBasicOn,
        deleteStyles: ui.graphicStyle.deleteStylesCheckbox.value,
        includeFontAndStyle: ui.format.fontAndStyleCheckbox.value,
        includeFontSize: ui.format.fontSizeCheckbox.value,
        includeFill: ui.basicFillStroke.fillCheckbox.value,
        includeStrokeColor: ui.basicFillStroke.strokeColorCheckbox.value,
        includeStrokeWidth: ui.basicFillStroke.strokeWidthCheckbox.value,
        anchor: getSelectedAnchorIndex(ui.coordinate.anchorRadios),
        axis: ui.coordinate.axisXRadio.value ? 'x'
            : (ui.coordinate.axisYRadio.value ? 'y' : 'both'),
        swapZOrder: ui.coordinate.zOrderCheckbox.value
    };
}

/* オプションダイアログを表示し、プレビューと確定を駆動 / Show the options dialog and drive preview / commit */
function showOptionsDialog(bothAreText, performSwapFn) {
    var dialog = new Window('dialog', getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
    dialog.orientation = 'column';
    dialog.alignChildren = 'fill';

    var modeUi = buildModePanel(dialog, bothAreText);

    // 本体を2カラムに：左＝スタイル系、右＝座標 / Two-column body: left = style panels, right = position
    var columns = dialog.add('group');
    columns.orientation = 'row';
    columns.alignment = 'fill';
    columns.alignChildren = ['fill', 'top'];

    var leftColumn = columns.add('group');
    leftColumn.orientation = 'column';
    leftColumn.alignChildren = ['fill', 'top'];

    var rightColumn = columns.add('group');
    rightColumn.orientation = 'column';
    rightColumn.alignChildren = ['fill', 'top'];

    var ui = {
        mode: modeUi,
        format: buildFormatPanel(leftColumn),
        basicFillStroke: buildBasicFillStrokePanel(leftColumn),
        graphicStyle: buildGraphicStylePanel(leftColumn),
        coordinate: buildCoordinatePanel(rightColumn)
    };
    var footer = buildFooter(dialog);
    var previewUi = { previewCheckbox: footer.previewCheckbox };
    var buttons = footer;

    var previewState = { isUndo: false };

    /* UI 変更時：有効状態を更新し、プレビューを再描画 / On UI change: refresh enabled state and re-render preview */
    var refresh = function () {
        updateDialogEnabled(ui, bothAreText);
        runPreview(previewState, function () {
            performSwapFn(readDialogOptions(ui), true);
        }, previewUi.previewCheckbox.value);
    };

    ui.mode.styleSwapRadio.onClick = refresh;
    ui.mode.contentSwapRadio.onClick = refresh;
    ui.mode.coordinateSwapRadio.onClick = refresh;
    ui.graphicStyle.swapGraphicStylesCheckbox.onClick = refresh;
    ui.graphicStyle.deleteStylesCheckbox.onClick = refresh;
    ui.basicFillStroke.fillCheckbox.onClick = refresh;
    ui.basicFillStroke.strokeColorCheckbox.onClick = refresh;
    ui.basicFillStroke.strokeWidthCheckbox.onClick = refresh;
    ui.format.fontAndStyleCheckbox.onClick = refresh;
    ui.format.fontSizeCheckbox.onClick = refresh;
    // 3×3 ラジオは行ごとにグループが分かれ自動排他が効かないため、手動で排他制御
    var anchorRadios = ui.coordinate.anchorRadios;
    for (var ai = 0; ai < anchorRadios.length; ai++) {
        (function (index) {
            anchorRadios[index].onClick = function () {
                for (var k = 0; k < anchorRadios.length; k++) {
                    anchorRadios[k].value = (k === index);
                }
                refresh();
            };
        })(ai);
    }
    ui.coordinate.axisBothRadio.onClick = refresh;
    ui.coordinate.axisXRadio.onClick = refresh;
    ui.coordinate.axisYRadio.onClick = refresh;
    ui.coordinate.zOrderCheckbox.onClick = refresh;
    previewUi.previewCheckbox.onClick = refresh;

    updateDialogEnabled(ui, bothAreText);

    /* OK：プレビュー分を巻き戻してから本番として再実行 / OK: undo preview then commit cleanly */
    buttons.okButton.onClick = function () {
        undoPreview(previewState);
        performSwapFn(readDialogOptions(ui), false);
        dialog.close(1);
    };

    /* キャンセル含むクローズ時：残ったプレビューを巻き戻す / On any close (incl. cancel): undo leftover preview */
    dialog.onClose = function () {
        cleanupPreview(previewState, app.activeDocument);
    };

    return dialog.show() === 1;
}

// =========================================
// グラフィックスタイル関連 / Graphic Style Helpers
// =========================================

/* 指定名のグラフィックスタイルが存在するか / Check if a graphic style with the given name exists */
function hasGraphicStyle(graphicStyles, styleName) {
    try {
        graphicStyles.getByName(styleName);
        return true;
    } catch (e) {
        return false;
    }
}

/* 未使用の _swapStyleX/_swapStyleY ペアを返す（A/B → C/D → … → Y/Z） / Return the next unused _swapStyleX/_swapStyleY pair */
function findNextStyleNamePair(graphicStyles) {
    var suffixLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    for (var i = 0; i < suffixLetters.length - 1; i += 2) {
        var firstStyleName = '_swapStyle' + suffixLetters.charAt(i);
        var secondStyleName = '_swapStyle' + suffixLetters.charAt(i + 1);
        if (!hasGraphicStyle(graphicStyles, firstStyleName) && !hasGraphicStyle(graphicStyles, secondStyleName)) {
            return [firstStyleName, secondStyleName];
        }
    }
    return null;
}

// 既存のユーザーアクションセットと衝突しないようユニークな名前を使う
var TEMP_ACTION_SET_NAME = '__SwapStyleGS';
var TEMP_ACTION_NAME = 'AddNewWithoutName';

/* アクションを一時ファイルに書き出してロードする / Write & load the force-new-style action */
function loadForceNewGraphicStyleAction() {
    // hex の `[ 13 5f5f537761705374796c654753 ]` はセット名 "__SwapStyleGS" (UTF-8) を表す
    var actionScriptSource = '/version 3 /name [ 13 5f5f537761705374796c654753 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /value 1 } } }';

    var actionFile = new File(Folder.temp.fsName + '/__tmp_swap_styles.aia');
    actionFile.open('w');
    actionFile.write(actionScriptSource);
    actionFile.close();
    app.loadAction(actionFile);
    actionFile.remove();
}

/* ロード済みアクションを実行（選択オブジェクトをスタイル登録） / Run the loaded action on current selection */
function runForceNewGraphicStyleAction() {
    app.doScript(TEMP_ACTION_NAME, TEMP_ACTION_SET_NAME, false);
}

/* アクションセットをアンロード / Unload the action set */
function unloadForceNewGraphicStyleAction() {
    try { app.unloadAction(TEMP_ACTION_SET_NAME, ''); } catch (e) { }
}

// =========================================
// 座標の交換 / Position Swap
// =========================================

/* geometricBounds と基準点インデックスから基準点座標を返す / Reference point from geometricBounds + anchor index */
function anchorPoint(bounds, anchorIndex) {
    // geometricBounds = [left, top, right, bottom]、index は行優先（0=左上 … 4=中央 … 8=右下）
    var col = anchorIndex % 3;
    var row = Math.floor(anchorIndex / 3);
    var x = (col === 0) ? bounds[0] : (col === 2 ? bounds[2] : (bounds[0] + bounds[2]) / 2);
    var y = (row === 0) ? bounds[1] : (row === 2 ? bounds[3] : (bounds[1] + bounds[3]) / 2);
    return [x, y];
}

/* 2つのオブジェクトの位置を入れ替える（基準点・軸ロックを反映） / Swap positions by reference point with axis lock */
function swapPositions(itemA, itemB, options) {
    var anchorIndex = (typeof options.anchor === 'number') ? options.anchor : 4;
    var axis = options.axis || 'both';
    var moveX = (axis === 'both' || axis === 'x');
    var moveY = (axis === 'both' || axis === 'y');
    var boundsA = itemA.geometricBounds;
    var boundsB = itemB.geometricBounds;
    var pointA = anchorPoint(boundsA, anchorIndex);
    var pointB = anchorPoint(boundsB, anchorIndex);
    // translate(deltaX, deltaY)：deltaY は上方向が正なので geometricBounds と整合
    itemA.translate(moveX ? (pointB[0] - pointA[0]) : 0, moveY ? (pointB[1] - pointA[1]) : 0);
    itemB.translate(moveX ? (pointA[0] - pointB[0]) : 0, moveY ? (pointA[1] - pointB[1]) : 0);
}

/* 2つのオブジェクトの重なり順を入れ替える / Swap the stacking order of two objects */
function swapZOrder(itemA, itemB) {
    try {
        // 各オブジェクトの直前にマーカーを置いて元位置を記録し、相手側のマーカー位置へ移動
        var markerA = itemA.parent.pathItems.add();
        markerA.move(itemA, ElementPlacement.PLACEBEFORE);
        var markerB = itemB.parent.pathItems.add();
        markerB.move(itemB, ElementPlacement.PLACEBEFORE);
        itemA.move(markerB, ElementPlacement.PLACEBEFORE);
        itemB.move(markerA, ElementPlacement.PLACEBEFORE);
        markerA.remove();
        markerB.remove();
    } catch (e) { }
}

// =========================================
// 基本属性の交換 / Basic Fill & Stroke Swap
// =========================================

/* 基本的な塗りや線をクロス入れ替え（TextFrame は characterAttributes 経由） / Cross-swap basic fill & stroke */
function swapBasicFillStroke(itemA, itemB, options) {
    var bothAreText = itemA.typename === 'TextFrame' && itemB.typename === 'TextFrame';
    if (bothAreText) {
        swapTextCharacterFillStroke(itemA, itemB, options);
    } else {
        swapPageItemFillStroke(itemA, itemB, options);
    }
}

/* PageItem 系（PathItem など）の塗り・線をクロス入れ替え / Cross-swap fill & stroke at PageItem level */
function swapPageItemFillStroke(itemA, itemB, options) {
    if (options.includeFill) {
        try {
            var fillColorA = itemA.fillColor;
            var fillColorB = itemB.fillColor;
            var filledA = itemA.filled;
            var filledB = itemB.filled;
            itemA.fillColor = fillColorB;
            itemB.fillColor = fillColorA;
            itemA.filled = filledB;
            itemB.filled = filledA;
        } catch (e) { }
    }
    if (options.includeStrokeColor) {
        try {
            var strokeColorA = itemA.strokeColor;
            var strokeColorB = itemB.strokeColor;
            var strokedA = itemA.stroked;
            var strokedB = itemB.stroked;
            itemA.strokeColor = strokeColorB;
            itemB.strokeColor = strokeColorA;
            itemA.stroked = strokedB;
            itemB.stroked = strokedA;
        } catch (e) { }
    }
    if (options.includeStrokeWidth) {
        try {
            var strokeWidthA = itemA.strokeWidth;
            var strokeWidthB = itemB.strokeWidth;
            itemA.strokeWidth = strokeWidthB;
            itemB.strokeWidth = strokeWidthA;
        } catch (e) { }
    }
}

/* TextFrame の characterAttributes 経由で塗り・線をクロス入れ替え / Cross-swap text fill & stroke via characterAttributes */
function swapTextCharacterFillStroke(textFrameA, textFrameB, options) {
    var charAttrsA = textFrameA.textRange.characterAttributes;
    var charAttrsB = textFrameB.textRange.characterAttributes;
    if (options.includeFill) {
        try {
            var fillColorA = charAttrsA.fillColor;
            var fillColorB = charAttrsB.fillColor;
            charAttrsA.fillColor = fillColorB;
            charAttrsB.fillColor = fillColorA;
        } catch (e) { }
    }
    if (options.includeStrokeColor) {
        try {
            var strokeColorA = charAttrsA.strokeColor;
            var strokeColorB = charAttrsB.strokeColor;
            charAttrsA.strokeColor = strokeColorB;
            charAttrsB.strokeColor = strokeColorA;
        } catch (e) { }
    }
    if (options.includeStrokeWidth) {
        try {
            // CharacterAttributes は strokeWidth ではなく strokeWeight
            var strokeWeightA = charAttrsA.strokeWeight;
            var strokeWeightB = charAttrsB.strokeWeight;
            charAttrsA.strokeWeight = strokeWeightB;
            charAttrsB.strokeWeight = strokeWeightA;
        } catch (e) { }
    }
}

// =========================================
// プレビュー undo ヘルパー / Preview Undo Helpers
// =========================================
/*
 * 設定変更のたびに「仮配置 → undo → 再生成」する典型パターンを
 * 共通化したユーティリティ。
 * Reusable preview/undo pattern for Illustrator dialog scripts.
 *
 * 使い方の概略 / Usage:
 *   var previewState = { isUndo: false };
 *   // UI 変更ごとに: runPreview(previewState, process, isPreview.value);
 *   // OK 時に:       undoPreview(previewState); process(); win.close();
 *   // onClose 時:    cleanupPreview(previewState, app.activeDocument);
 *
 * 注意 / Notes:
 *  - process() 内で graphicStyles.add() / layers.add() など
 *    app.undo() で戻らない副作用を起こす場合、cleanupPreview だけでは
 *    残骸が残る可能性がある（呼び出し側で手動 .remove() を検討）。
 *  - process() は ScriptUI の選択値を参照する純粋な再描画関数として
 *    書くと、内部状態を保持しない実装にしやすい。
 */

/* プレビューを再描画する / Re-render preview */
function runPreview(state, processFn, isEnabled) {
    try {
        if (isEnabled) {
            if (state.isUndo) app.undo();
            else state.isUndo = true;
            processFn();
            app.redraw();
        } else if (state.isUndo) {
            app.undo();
            app.redraw();
            state.isUndo = false;
        }
    } catch (err) { }
}

/* 確定処理の直前にプレビュー分を巻き戻す / Undo preview before the final commit */
function undoPreview(state) {
    try {
        if (state.isUndo) app.undo();
    } catch (err) { }
    state.isUndo = false;
}

/* ダイアログクローズ時のクリーンアップ / Cleanup on dialog close */
function cleanupPreview(state, doc, tempLayerName) {
    try {
        if (state.isUndo) app.undo();
        state.isUndo = false;
    } catch (err) { }
    if (tempLayerName) {
        try {
            var tmpLay = doc.layers.getByName(tempLayerName);
            tmpLay.remove();
        } catch (err) { }
    }
}

// =========================================
// 交換処理本体 / Swap Operation
// =========================================

/* スタイル／文字列の交換を実行（プレビュー・本番共通） / Run the swap (used for both preview and commit) */
function performSwap(activeDoc, graphicStyles, targetItems, options, isPreview) {
    // 文字列交換モード：テキストオブジェクト同士で contents を入れ替えるだけ
    if (options.mode === 'content') {
        if (targetItems[0].typename !== 'TextFrame' || targetItems[1].typename !== 'TextFrame') {
            if (!isPreview) alert(getLabel('alertContentTextOnly'));
            return;
        }
        var contentA = targetItems[0].contents;
        targetItems[0].contents = targetItems[1].contents;
        targetItems[1].contents = contentA;
        activeDoc.selection = targetItems;
        return;
    }

    // 座標交換モード：2つのオブジェクトの位置（左上基準）を入れ替える
    if (options.mode === 'coordinate') {
        swapPositions(targetItems[0], targetItems[1], options);
        if (options.swapZOrder) {
            swapZOrder(targetItems[0], targetItems[1]);
        }
        activeDoc.selection = targetItems;
        return;
    }

    if (options.swapGraphicStyles) {
        var tempStyleNamePair = findNextStyleNamePair(graphicStyles);
        if (!tempStyleNamePair) {
            if (!isPreview) alert(getLabel('alertNoFreePair'));
            return;
        }

        // 途中失敗時の掃除のため、実際に登録できた名前だけを別配列で追跡
        var registeredStyleNames = [];
        var actionLoaded = false;
        try {
            loadForceNewGraphicStyleAction();
            actionLoaded = true;
            for (var i = 0; i < targetItems.length; i++) {
                activeDoc.selection = null;
                targetItems[i].selected = true;
                var styleCountBefore = graphicStyles.length;
                runForceNewGraphicStyleAction();
                // doScript が無音で失敗した場合、最後尾を取ると既存スタイルをリネームしてしまう
                if (graphicStyles.length <= styleCountBefore) {
                    throw new Error('registration failed');
                }
                graphicStyles[graphicStyles.length - 1].name = tempStyleNamePair[i];
                registeredStyleNames.push(tempStyleNamePair[i]);
            }

            // A に styleB、B に styleA を適用（クロス）
            graphicStyles.getByName(tempStyleNamePair[1]).applyTo(targetItems[0]);
            graphicStyles.getByName(tempStyleNamePair[0]).applyTo(targetItems[1]);

            // ON のとき登録したスタイルを削除（アピアランス自体は各オブジェクトに残る）
            if (options.deleteStyles) {
                for (var j = 0; j < registeredStyleNames.length; j++) {
                    try { graphicStyles.getByName(registeredStyleNames[j]).remove(); }
                    catch (e) { }
                }
                registeredStyleNames = [];
            }
        } catch (mainErr) {
            // 失敗時は登録済みの一時スタイルを掃除してから抜ける
            for (var k = 0; k < registeredStyleNames.length; k++) {
                try { graphicStyles.getByName(registeredStyleNames[k]).remove(); }
                catch (e) { }
            }
            if (actionLoaded) unloadForceNewGraphicStyleAction();
            if (!isPreview) alert(getLabel('alertSwapFailed'));
            return;
        }
        if (actionLoaded) unloadForceNewGraphicStyleAction();
    }

    // 基本的な塗りや線のクロス入れ替え
    if (options.includeFill || options.includeStrokeColor || options.includeStrokeWidth) {
        swapBasicFillStroke(targetItems[0], targetItems[1], options);
    }

    // 書式情報のクロス入れ替え（両方ともテキストオブジェクトのとき）
    if ((options.includeFontAndStyle || options.includeFontSize)
        && targetItems[0].typename === 'TextFrame'
        && targetItems[1].typename === 'TextFrame') {
        var charAttrsA = targetItems[0].textRange.characterAttributes;
        var charAttrsB = targetItems[1].textRange.characterAttributes;
        var fontA = charAttrsA.textFont;
        var fontB = charAttrsB.textFont;
        var sizeA = charAttrsA.size;
        var sizeB = charAttrsB.size;

        if (options.includeFontAndStyle) {
            charAttrsA.textFont = fontB;
            charAttrsB.textFont = fontA;
        }
        if (options.includeFontSize) {
            charAttrsA.size = sizeB;
            charAttrsB.size = sizeA;
        }
    }

    // 選択を元の2つに戻す
    activeDoc.selection = targetItems;
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        alert(getLabel('alertNoDocument'));
        return;
    }
    var activeDoc = app.activeDocument;
    var graphicStyles = activeDoc.graphicStyles;

    var selectedItems = activeDoc.selection;
    if (selectedItems.length !== 2) {
        alert(getLabel('alertSelectTwo'));
        return;
    }

    // selection は後続処理で変更するため、対象2件を配列として保持 / Keep the two targets as an array because selection changes later
    var targetItems = [selectedItems[0], selectedItems[1]];

    var bothAreText = (targetItems[0].typename === 'TextFrame'
        && targetItems[1].typename === 'TextFrame');

    var performSwapFn = function (options, isPreview) {
        performSwap(activeDoc, graphicStyles, targetItems, options, isPreview);
    };

    showOptionsDialog(bothAreText, performSwapFn);
})();