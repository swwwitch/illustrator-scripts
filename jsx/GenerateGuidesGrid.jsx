<<<<<<< HEAD
// #target illustrator
=======
#target illustrator
>>>>>>> a8cc9f637bc495f8fb94ef76fcaa29a31445aa76

/*
 * スクリプトの概要 / Script Overview：
 * Illustratorのアートボードを、指定した行数・列数に分割してガイドを自動生成します。
 * 必要に応じて、分割されたセルごとに長方形を描画することもできます。
 * ダイアログは日本語／英語に自動対応し、各種オプションをリアルタイムでプレビュー可能です。
 *
 * 主な機能 / Main Features：
 * - 行数・列数・ガター（行間／列間）・上下左右マージン・ガイドの伸張距離を自由に設定可能
 * - セルごとの長方形描画（K100またはR0G0B0、不透明度15%）
 * - ガイドを描画するか選択可能（デフォルトON）
 * - 裁ち落としガイドを描画可能（デフォルトOFF・3mm）
 * - セル長方形のプレビューにも対応（本番作成前に確認）
 * - 「すべてのアートボードに適用」オプションあり
 * - 「すべて同じにする」チェックON時、マージン個別入力欄をディム表示
 * - 現在の設定をプリセットとしてコード書き出し（ファイル名＝プリセット名、拡張子自動付与）
 * - プリセットに「ガイドを引く」「裁ち落としのガイド」も含められる
 * - 日本語／英語環境自動切り替え対応（$.localeによる判定）
 *
 * その他仕様 / Other Specifications：
 * - ガイドと長方形は、それぞれ「grid_guides」「cell-rectangle」レイヤーに整理して作成
 * - プリセット書き出し時、拡張子がない場合は.txtを自動付与
 * - カラーモード（CMYK/RGB）に応じて長方形の塗り色を自動調整
 *
 * オリジナルアイデア：スガサワ君β
 * https://note.com/sgswkn/n/nee8c3ec1a14c
 *
 * Created: 2025-04-24
 * Updated: 2025-04-27（ガイドを引く・裁ち落としガイド・プリセット対応版）
 */

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてください。\nPlease open a document.");
        return;
    }

    // 言語判定
    function getCurrentLang() {
        return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
    }
    var lang = getCurrentLang();

    // プリセット定義（drawGuides, drawBleedGuide追加）
    var presets = [
        { label: (lang === 'ja') ? "十字" : "Cross", x: 2, y: 2, ext: 0, top: 0, bottom: 0, left: 0, right: 0, rowGutter: 0, colGutter: 0, drawCells: false, drawGuides: true, drawBleedGuide: false },
        { label: (lang === 'ja') ? "シングル" : "Single", x: 1, y: 1, ext: 50, top: 100, bottom: 100, left: 100, right: 100, rowGutter: 0, colGutter: 0, drawCells: true, drawGuides: true, drawBleedGuide: false },
        { label: (lang === 'ja') ? "2行×2列" : "2 Rows × 2 Columns", x: 2, y: 2, ext: 20, top: 0, bottom: 0, left: 0, right: 0, rowGutter: 50, colGutter: 50, drawCells: true, drawGuides: true, drawBleedGuide: false },
        { label: (lang === 'ja') ? "1行×3列" : "1 Row × 3 Columns", x: 3, y: 1, ext: 0, top: 30, bottom: 30, left: 30, right: 30, rowGutter: 0, colGutter: 30, drawCells: true, drawGuides: true, drawBleedGuide: false },
        { label: (lang === 'ja') ? "4行×4列" : "4 Rows × 4 Columns", x: 4, y: 4, ext: 0, top: 0, bottom: 0, left: 0, right: 0, rowGutter: 20, colGutter: 20, drawCells: true, drawGuides: true, drawBleedGuide: false },
        { label: (lang === 'ja') ? "2行×3列" : "3 Rows × 3 Columns", x: 3, y: 2, ext: 0, top: 100, bottom: 100, left: 100, right: 100, rowGutter: 20, colGutter: 20, drawCells: true, drawGuides: true, drawBleedGuide: false },
        { label: (lang === 'ja') ? "3行×3列" : "3 Rows × 3 Columns", x: 3, y: 3, ext: 0, top: 0, bottom: 0, left: 200, right: 0, rowGutter: 0, colGutter: 0, drawCells: true, drawGuides: true, drawBleedGuide: false },
        { label: (lang === 'ja') ? "sp" : "sp", x: 1, y: 1, ext: 0, top: 220, bottom: 220, left: 0, right: 0, rowGutter: 0, colGutter: 0, drawCells: true, drawGuides: true, drawBleedGuide: false },
        { label: (lang === 'ja') ? "長方形のみ" : "just rectangle", x: 1, y: 1, ext: 10, top: 0, bottom: 0, left: 0, right: 0, rowGutter: 0, colGutter: 0, drawCells: true, drawGuides: false, drawBleedGuide: false }
    ];
    // UIラベル定義
    var dialogTitle = (lang === 'ja') ? '段組設定Pro' : 'Split into Grid Pro';
    var presetLabel = (lang === 'ja') ? 'プリセット：' : 'Preset:';
    var rowTitle = (lang === 'ja') ? '行（─ 横線）' : 'Rows (─ Horizontal)';
    var columnTitle = (lang === 'ja') ? '列（│ 縦線）' : 'Columns (│ Vertical)';
    var rowsLabel = (lang === 'ja') ? '行数：' : 'Rows:';
    var rowGutterLabel = (lang === 'ja') ? '行間' : 'Row Gutter';
    var columnsLabel = (lang === 'ja') ? '列数：' : 'Columns:';
    var colGutterLabel = (lang === 'ja') ? '列間' : 'Column Gutter';
    var marginTitle = (lang === 'ja') ? 'マージン設定' : 'Margin Settings';
    var topLabel = (lang === 'ja') ? '上：' : 'Top:';
    var leftLabel = (lang === 'ja') ? '左：' : 'Left:';
    var bottomLabel = (lang === 'ja') ? '下：' : 'Bottom:';
    var rightLabel = (lang === 'ja') ? '右：' : 'Right:';
    var commonMarginLabel = (lang === 'ja') ? 'すべて同じ値にする' : 'Same Value';
    var guideExtensionLabel = (lang === 'ja') ? 'ガイドの伸張：' : 'Guide Extension:';
    var bleedGuideLabel = (lang === 'ja') ? '裁ち落としのガイド：' : 'Bleed Guide:';
    var allBoardsLabel = (lang === 'ja') ? 'すべてのアートボードに適用' : 'Apply to All Artboards';
    var cellRectLabel = (lang === 'ja') ? 'セルを長方形化' : 'Create Cell Rectangles';
    var drawGuidesLabel = (lang === 'ja') ? 'ガイドを引く' : 'Draw Guides';
    var clearGuidesLabel = (lang === 'ja') ? 'grid_guidesレイヤーをクリア' : 'Clear grid_guides Layer';
    var cancelLabel = (lang === 'ja') ? 'キャンセル' : 'Cancel';
    var applyLabel = (lang === 'ja') ? '適用' : 'Apply';
    var okLabel = (lang === 'ja') ? 'OK' : 'OK';
    var exportPresetLabel = (lang === 'ja') ? 'プリセット書き出し' : 'Export Preset';

    var doc = app.activeDocument;
    var rulerUnit = app.preferences.getIntegerPreference("rulerType");
    var unitLabel = "pt";
    var unitFactor = 1.0;

    // 単位設定
    if (rulerUnit === 0) { unitLabel = "inch"; unitFactor = 72.0; }
    else if (rulerUnit === 1) { unitLabel = "mm"; unitFactor = 72.0 / 25.4; }
    else if (rulerUnit === 2) { unitLabel = "pt"; unitFactor = 1.0; }
    else if (rulerUnit === 3) { unitLabel = "pica"; unitFactor = 12.0; }
    else if (rulerUnit === 4) { unitLabel = "cm"; unitFactor = 72.0 / 2.54; }
    else if (rulerUnit === 5) { unitLabel = "Q"; unitFactor = 72.0 / 25.4 * 0.25; }
    else if (rulerUnit === 6) { unitLabel = "px"; unitFactor = 1.0; }
    // ダイアログ作成
    var dlg = new Window("dialog", dialogTitle);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    // grid_guidesレイヤークリアチェックボックス
    var clearGuidesCheckbox = dlg.add("checkbox", undefined, clearGuidesLabel);
    clearGuidesCheckbox.value = true;

    // プリセット選択＋書き出しボタングループ
    var presetGroup = dlg.add("group");
    presetGroup.orientation = "row";
    presetGroup.alignChildren = "center";
    presetGroup.margins = [0, 5, 0, 10];

    presetGroup.add("statictext", undefined, presetLabel);
    var presetDropdown = presetGroup.add("dropdownlist", undefined, []);
    presetDropdown.selection = 0;
    var btnExportPreset = presetGroup.add("button", undefined, exportPresetLabel);


btnExportPreset.onClick = function () {
    var saveFile = File.saveDialog((lang === 'ja') ? "プリセットを書き出す場所と名前を指定してください" : "Choose where to save the preset", "*.txt");
    if (!saveFile) {
        return;
    }

    // 拡張子がない場合は.txtをつける
    if (saveFile.name.indexOf(".") === -1) {
        saveFile = new File(saveFile.fsName + ".txt");
    }

    // ★ファイル名から.txtを正しく除去！
    var fileName = saveFile.name.replace(/\\.txt$/i, "");

    var currentPreset = {
        x: parseInt(inputXText.text, 10),
        y: parseInt(inputYText.text, 10),
        ext: parseFloat(inputExt.text),
        top: parseFloat(inputTop.text),
        bottom: parseFloat(inputBottom.text),
        left: parseFloat(inputLeft.text),
        right: parseFloat(inputRight.text),
        rowGutter: parseFloat(inputRowGutter.text),
        colGutter: parseFloat(inputColGutter.text),
        drawCells: cellRectCheckbox.value,
        drawGuides: drawGuidesCheckbox.value,
        drawBleedGuide: bleedGuideCheckbox.value
    };

    var presetString = '{ label: (lang === \'ja\') ? "' + fileName + '" : "' + fileName + '", ' +
        'x: ' + currentPreset.x + ', ' +
        'y: ' + currentPreset.y + ', ' +
        'ext: ' + currentPreset.ext + ', ' +
        'top: ' + currentPreset.top + ', ' +
        'bottom: ' + currentPreset.bottom + ', ' +
        'left: ' + currentPreset.left + ', ' +
        'right: ' + currentPreset.right + ', ' +
        'rowGutter: ' + currentPreset.rowGutter + ', ' +
        'colGutter: ' + currentPreset.colGutter + ', ' +
        'drawCells: ' + currentPreset.drawCells + ', ' +
        'drawGuides: ' + currentPreset.drawGuides + ', ' +
        'drawBleedGuide: ' + currentPreset.drawBleedGuide +
        ' }';

    if (saveFile.open("w")) {
        saveFile.write(presetString);
        saveFile.close();
        alert((lang === 'ja') ? "プリセットを書き出しました！" : "Preset exported!");
    } else {
        alert((lang === 'ja') ? "ファイルを書き込めませんでした。" : "Failed to write the file.");
    }
};

    // グリッド設定グループ
    var gridGroup = dlg.add("group");
    gridGroup.orientation = "row";
    gridGroup.alignChildren = "top";
    gridGroup.spacing = 20;

    // 行設定パネル
    var rowBlock = gridGroup.add("panel", undefined, rowTitle);
    rowBlock.orientation = "column";
    rowBlock.alignChildren = "left";
    rowBlock.margins = [15, 20, 15, 15];

    var inputY = rowBlock.add("group");
    inputY.add("statictext", undefined, rowsLabel);
    var inputYText = inputY.add("edittext", undefined, "2");
    inputYText.characters = 3;

    var rowGutterGroup = rowBlock.add("group");
    rowGutterGroup.add("statictext", undefined, rowGutterLabel + "：");
    var inputRowGutter = rowGutterGroup.add("edittext", undefined, "0");
    inputRowGutter.characters = 4;
    rowGutterGroup.add("statictext", undefined, unitLabel);

    // 列設定パネル
    var colBlock = gridGroup.add("panel", undefined, columnTitle);
    colBlock.orientation = "column";
    colBlock.alignChildren = "left";
    colBlock.margins = [15, 20, 15, 15];

    var inputX = colBlock.add("group");
    inputX.add("statictext", undefined, columnsLabel);
    var inputXText = inputX.add("edittext", undefined, "2");
    inputXText.characters = 3;

    var colGutterGroup = colBlock.add("group");
    colGutterGroup.add("statictext", undefined, colGutterLabel + "：");
    var inputColGutter = colGutterGroup.add("edittext", undefined, "0");
    inputColGutter.characters = 4;
    colGutterGroup.add("statictext", undefined, unitLabel);

// マージン全体パネル
var marginPanel = dlg.add("panel", undefined, marginTitle + " (" + unitLabel + ")");
marginPanel.orientation = "column"; // ★ここがcolumnに（横並び→縦並びに変更）
marginPanel.alignChildren = "left";
marginPanel.margins = [10, 15, 10, 15];

// --- 上段グループ（左／上下／右） ---
var upperGroup = marginPanel.add("group");
upperGroup.orientation = "row"; // 横並び
upperGroup.alignChildren = "top";

// --- 左だけグループ ---
var leftGroup = upperGroup.add("group"); // 枠なしグループ
leftGroup.orientation = "row";
leftGroup.alignChildren = "center";
leftGroup.margins = [0, 12, 0, 10]; 

leftGroup.add("statictext", undefined, leftLabel);
var inputLeft = leftGroup.add("edittext", undefined, "0");
inputLeft.characters = 5;

// --- 上下グループ ---
var topBottomGroup = upperGroup.add("group"); // 枠なしグループ
topBottomGroup.orientation = "column"; // 縦に並べる
topBottomGroup.alignChildren = "left";
topBottomGroup.margins = [5, 0, 5, 0]; // 右にスペース

var topGroup = topBottomGroup.add("group");
topGroup.orientation = "row";
topGroup.add("statictext", undefined, topLabel);
var inputTop = topGroup.add("edittext", undefined, "0");
inputTop.characters = 5;

var bottomGroup = topBottomGroup.add("group");
bottomGroup.orientation = "row";
bottomGroup.add("statictext", undefined, bottomLabel);
var inputBottom = bottomGroup.add("edittext", undefined, "0");
inputBottom.characters = 5;

// --- 右だけグループ ---
var rightGroup = upperGroup.add("group"); // 枠なしグループ
rightGroup.orientation = "row";
rightGroup.alignChildren = "center";
rightGroup.margins = [0, 12, 0, 10]; 

rightGroup.add("statictext", undefined, rightLabel);
var inputRight = rightGroup.add("edittext", undefined, "0");
inputRight.characters = 5;

// --- 共通マージングループ（下段に追加） ---
var commonGroup = marginPanel.add("group");
commonGroup.orientation = "row";
commonGroup.alignChildren = "center";
commonGroup.margins = [0, 10, 0, 0];

var commonMarginCheckbox = commonGroup.add("checkbox", undefined, commonMarginLabel);
var commonMarginInput = commonGroup.add("edittext", undefined, "0");
commonMarginInput.characters = 5;

    // オプション設定グループ（ガイドの伸張・裁ち落としガイド・セル長方形化・ガイドを引く）
    var optGroup = dlg.add("group");
    optGroup.orientation = "column";
    optGroup.alignChildren = "left";

    // ガイドの伸張設定
    var extGroup = optGroup.add("group");
    extGroup.margins = [0, 0, 0, 10];
    extGroup.add("statictext", undefined, guideExtensionLabel);
    var inputExt = extGroup.add("edittext", undefined, "10");
    inputExt.characters = 5;
    extGroup.add("statictext", undefined, unitLabel);

    // 裁ち落としガイド設定
    var bleedGroup = optGroup.add("group");
    bleedGroup.margins = [0, 0, 0, 10];
    var bleedGuideCheckbox = bleedGroup.add("checkbox", undefined, bleedGuideLabel);
    bleedGuideCheckbox.value = false; // デフォルトOFF
    var inputBleed = bleedGroup.add("edittext", undefined, "3"); // デフォルト3mm
    inputBleed.characters = 4;
    bleedGroup.add("statictext", undefined, "(mm)");

    var allBoardsCheckbox = optGroup.add("checkbox", undefined, allBoardsLabel);

    // セル長方形化・ガイドを引くを横並び
    var cellGuideGroup = optGroup.add("group");
    cellGuideGroup.orientation = "row";
    cellGuideGroup.alignChildren = "left";

    var cellRectCheckbox = cellGuideGroup.add("checkbox", undefined, cellRectLabel);
    var drawGuidesCheckbox = cellGuideGroup.add("checkbox", undefined, drawGuidesLabel);
    drawGuidesCheckbox.value = true;

    // === ボタンエリア（レイアウト変更版）===
    var outerGroup = dlg.add("group");
    outerGroup.orientation = "row";
    outerGroup.alignChildren = ["fill", "center"];
    outerGroup.margins = [0, 10, 0, 0];
    outerGroup.spacing = 0;

    // 左グループ（キャンセルボタン）
    var leftGroup = outerGroup.add("group");
    leftGroup.orientation = "row";
    leftGroup.alignChildren = "left";
    // leftGroup.spacing = 10;
    var btnCancel = leftGroup.add("button", undefined, cancelLabel, { name: "cancel" });

    // スペーサー（横に伸びる空白）
    var spacer = outerGroup.add("group");
    spacer.alignment = ["fill", "fill"];
    spacer.minimumSize.width = 70;
    spacer.maximumSize.height = 0;

    // 右グループ（適用・OKボタン）
    var rightGroup = outerGroup.add("group");
    rightGroup.orientation = "row";
    rightGroup.alignChildren = "right";
    rightGroup.spacing = 10;
    var btnApply = rightGroup.add("button", undefined, applyLabel);
    var btnOK = rightGroup.add("button", undefined, okLabel, { name: "ok" });

    // プリセットをドロップダウンに追加
    for (var i = 0; i < presets.length; i++) {
        presetDropdown.add("item", presets[i].label);
    }
    presetDropdown.selection = 0;

    // プリセット選択時、入力値を反映
    presetDropdown.onChange = function () {
        var p = presets[presetDropdown.selection.index];
        inputXText.text = p.x;
        inputYText.text = p.y;
        inputExt.text = p.ext;
        inputTop.text = p.top;
        inputBottom.text = p.bottom;
        inputLeft.text = p.left;
        inputRight.text = p.right;
        inputRowGutter.text = (p.rowGutter !== undefined) ? p.rowGutter : "0";
        inputColGutter.text = (p.colGutter !== undefined) ? p.colGutter : "0";
        cellRectCheckbox.value = (typeof p.drawCells !== "undefined") ? p.drawCells : false;
        drawGuidesCheckbox.value = (typeof p.drawGuides !== "undefined") ? p.drawGuides : true;
        bleedGuideCheckbox.value = (typeof p.drawBleedGuide !== "undefined") ? p.drawBleedGuide : false;
        updateGutterEnable();
        syncCommonMargin();
        drawGuides(true);
    };
    // 「すべて同じにする」同期処理
    function syncCommonMargin() {
        if (commonMarginCheckbox.value) {
            var val = commonMarginInput.text;
            inputTop.text = val;
            inputBottom.text = val;
            inputLeft.text = val;
            inputRight.text = val;
            inputTop.enabled = false;
            inputBottom.enabled = false;
            inputLeft.enabled = false;
            inputRight.enabled = false;
        } else {
            inputTop.enabled = true;
            inputBottom.enabled = true;
            inputLeft.enabled = true;
            inputRight.enabled = true;
        }
        drawGuides(true);
    }
    commonMarginInput.onChanging = syncCommonMargin;
    commonMarginCheckbox.onClick = syncCommonMargin;

    // ガター有効無効切り替え
    function updateGutterEnable() {
        var xVal = parseInt(inputXText.text, 10);
        var yVal = parseInt(inputYText.text, 10);
        inputRowGutter.enabled = (yVal > 1);
        inputColGutter.enabled = (xVal > 1);
    }
    inputXText.onChanging = inputYText.onChanging = function () {
        updateGutterEnable();
    };

    // 「ガイドを引く」オプション切り替え時、伸張・裁ち落としをディム制御
    drawGuidesCheckbox.onClick = function () {
        var enable = drawGuidesCheckbox.value;
        inputExt.enabled = enable;
        bleedGuideCheckbox.enabled = enable;
        inputBleed.enabled = enable;
        drawGuides(true);
    };

    // 適用ボタン押下時
    btnApply.onClick = function () {
        updateGutterEnable();
        if (clearGuidesCheckbox.value) {
            clearGuidesLayer();
        }
        drawGuides(true);
    };

    // OKボタン押下時
    btnOK.onClick = function () {
        updateGutterEnable();
        if (clearGuidesCheckbox.value) {
            clearGuidesLayer();
        }
        dlg.close(1);
    };
    // ガイド＆セル長方形＆裁ち落としガイドを描画
    function drawGuides(isPreview) {
        if (isPreview) {
            removePreviewGuides();
        }

        var xDiv = parseInt(inputXText.text, 10);
        var yDiv = parseInt(inputYText.text, 10);
        var ext = parseFloat(inputExt.text) * unitFactor;
        var top = parseFloat(inputTop.text) * unitFactor;
        var bottom = parseFloat(inputBottom.text) * unitFactor;
        var left = parseFloat(inputLeft.text) * unitFactor;
        var right = parseFloat(inputRight.text) * unitFactor;
        var rowGutter = parseFloat(inputRowGutter.text) * unitFactor;
        var colGutter = parseFloat(inputColGutter.text) * unitFactor;
        var bleed = parseFloat(inputBleed.text) * (72.0 / 25.4); // mm → pt換算
        var allBoards = allBoardsCheckbox.value;
        var drawCells = cellRectCheckbox.value;
        var drawGuidesNow = drawGuidesCheckbox.value;
        var drawBleedGuide = bleedGuideCheckbox.value;

        if (isNaN(xDiv) || xDiv <= 0 || isNaN(yDiv) || yDiv <= 0) return;

        var gridLayerName = isPreview ? "_Preview_Guides" : "grid_guides";
        var gridLayer;
        try {
            gridLayer = doc.layers.getByName(gridLayerName);
        } catch (e) {
            gridLayer = doc.layers.add();
            gridLayer.name = gridLayerName;
        }
        gridLayer.locked = false;

        var cellLayer = gridLayer;
        if (!isPreview && drawCells) {
            try {
                cellLayer = doc.layers.getByName("cell-rectangle");
            } catch (e) {
                cellLayer = doc.layers.add();
                cellLayer.name = "cell-rectangle";
            }
            cellLayer.locked = false;
        }
        for (var b = 0; b < doc.artboards.length; b++) {
            if (!allBoards && b !== doc.artboards.getActiveArtboardIndex()) continue;

            var ab = doc.artboards[b];
            var rect = ab.artboardRect;
            var abLeft = rect[0], abTop = rect[1], abRight = rect[2], abBottom = rect[3];
            var baseLeft = abLeft + left;
            var baseRight = abRight - right;
            var baseTop = abTop - top;
            var baseBottom = abBottom + bottom;

            var usableWidth = baseRight - baseLeft;
            var usableHeight = baseTop - baseBottom;
            var totalColGutter = (xDiv - 1) * colGutter;
            var totalRowGutter = (yDiv - 1) * rowGutter;
            var cellWidth = (usableWidth - totalColGutter) / xDiv;
            var cellHeight = (usableHeight - totalRowGutter) / yDiv;

            var guideLeft = abLeft - ext;
            var guideRight = abRight + ext;
            var guideTop = abTop + ext;
            var guideBottom = abBottom - ext;

            if (drawGuidesNow) {
                // 通常ガイド描画（行・列）
                var y = baseTop;
                var line = doc.pathItems.add();
                line.setEntirePath([[guideLeft, y], [guideRight, y]]);
                line.stroked = false;
                line.filled = false;
                line.guides = true;
                line.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                for (var j = 0; j < yDiv; j++) {
                    y -= cellHeight;
                    var line1 = doc.pathItems.add();
                    line1.setEntirePath([[guideLeft, y], [guideRight, y]]);
                    line1.stroked = false;
                    line1.filled = false;
                    line1.guides = true;
                    line1.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                    if (j < yDiv - 1) {
                        y -= rowGutter;
                        var line2 = doc.pathItems.add();
                        line2.setEntirePath([[guideLeft, y], [guideRight, y]]);
                        line2.stroked = false;
                        line2.filled = false;
                        line2.guides = true;
                        line2.move(gridLayer, ElementPlacement.PLACEATBEGINNING);
                    }
                }

                y = baseBottom;
                var lineBottom = doc.pathItems.add();
                lineBottom.setEntirePath([[guideLeft, y], [guideRight, y]]);
                lineBottom.stroked = false;
                lineBottom.filled = false;
                lineBottom.guides = true;
                lineBottom.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                var x = baseLeft;
                var vline = doc.pathItems.add();
                vline.setEntirePath([[x, guideTop], [x, guideBottom]]);
                vline.stroked = false;
                vline.filled = false;
                vline.guides = true;
                vline.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                for (var i = 0; i < xDiv; i++) {
                    x += cellWidth;
                    var vline1 = doc.pathItems.add();
                    vline1.setEntirePath([[x, guideTop], [x, guideBottom]]);
                    vline1.stroked = false;
                    vline1.filled = false;
                    vline1.guides = true;
                    vline1.move(gridLayer, ElementPlacement.PLACEATBEGINNING);

                    if (i < xDiv - 1) {
                        x += colGutter;
                        var vline2 = doc.pathItems.add();
                        vline2.setEntirePath([[x, guideTop], [x, guideBottom]]);
                        vline2.stroked = false;
                        vline2.filled = false;
                        vline2.guides = true;
                        vline2.move(gridLayer, ElementPlacement.PLACEATBEGINNING);
                    }
                }

                x = baseRight;
                var vlineRight = doc.pathItems.add();
                vlineRight.setEntirePath([[x, guideTop], [x, guideBottom]]);
                vlineRight.stroked = false;
                vlineRight.filled = false;
                vlineRight.guides = true;
                vlineRight.move(gridLayer, ElementPlacement.PLACEATBEGINNING);
            }

            if (drawGuidesNow && drawBleedGuide) {
                // ★裁ち落としガイド
                var bleedRect = doc.pathItems.rectangle(
                    abTop + bleed,
                    abLeft - bleed,
                    (abRight - abLeft) + bleed * 2,
                    (abTop - abBottom) + bleed * 2
                );
                bleedRect.stroked = false;
                bleedRect.filled = false;
                bleedRect.guides = true;
                bleedRect.move(gridLayer, ElementPlacement.PLACEATBEGINNING);
            }

            if (drawCells && cellLayer) {
                var startX = baseLeft;
                var startY = baseTop;
                for (var row = 0; row < yDiv; row++) {
                    var cellY = startY - (cellHeight + rowGutter) * row;
                    for (var col = 0; col < xDiv; col++) {
                        var cellX = startX + (cellWidth + colGutter) * col;
                        var rect = cellLayer.pathItems.rectangle(cellY, cellX, cellWidth, cellHeight);
                        rect.stroked = false;
                        rect.filled = true;
                        rect.fillColor = createBlackColor();
                        rect.opacity = 15;
                    }
                }
            }
        }

        if (!isPreview) {
            gridLayer.locked = true;
        }

        if (isPreview) {
            app.redraw();
        }
    }
    // 黒色作成（CMYK／RGB対応）
    function createBlackColor() {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 0;
            cmyk.yellow = 0;
            cmyk.black = 100;
            return cmyk;
        } else {
            var rgb = new RGBColor();
            rgb.red = 0;
            rgb.green = 0;
            rgb.blue = 0;
            return rgb;
        }
    }

    // プレビューガイド削除
    function removePreviewGuides() {
        try {
            var previewLayer = doc.layers.getByName("_Preview_Guides");
            if (previewLayer) previewLayer.remove();
        } catch (e) {}
    }

    // grid_guidesレイヤーのガイドだけ削除
    function clearGuidesLayer() {
        var guidesLayer = null;
        for (var i = 0; i < doc.layers.length; i++) {
            if (doc.layers[i].name === "grid_guides") {
                guidesLayer = doc.layers[i];
                break;
            }
        }
        if (guidesLayer) {
            if (guidesLayer.locked) {
                guidesLayer.locked = false;
            }
            for (var j = guidesLayer.pageItems.length - 1; j >= 0; j--) {
                var item = guidesLayer.pageItems[j];
                if (item.guides) {
                    item.remove();
                }
            }
        }
    }

    // ダイアログ初期プレビュー＆終了時処理
    updateGutterEnable();
    syncCommonMargin();
    drawGuides(true);

    if (dlg.show() === 1) {
        removePreviewGuides();
        if (clearGuidesCheckbox.value) {
            clearGuidesLayer();
        }
        drawGuides(false);
    } else {
        removePreviewGuides();
    }

})();
