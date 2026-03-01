
#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/**
 * クリップグループの形状変更 / Change clipping group shape
 * 更新日 / Updated: 2026-02-01
 *
 * 概要 / Overview:
 * - 選択した画像（配置/埋め込み）または既存のクリップグループを対象に、
 *   正方形 / 正円 / 六角形（A/B）のクリップ形状へ置き換えます。
 * - 「ケイ線を追加」ONで、生成後にケイ線を追加（メニューコマンド実行）します。
 * - 「角丸」ONで、生成したクリップグループに角丸 LiveEffect を適用します（※正円では無効）。
 * - 「複数オブジェクト > 大きさを揃える」ONで、選択内の最大/最小（面積）に合わせて
 *   すべての対象を同じ幅・高さにリサイズしてから処理します。
 *
 * 使い方 / How to use:
 * 1) 対象を選択して実行
 * 2) 形状とアピアランスを指定して OK
 *
 * 注意 / Notes:
 * - 「大きさを揃える」は非等比（幅・高さを別々に）で揃えるため、画像が歪む場合があります。
 */

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";

/* 形状モード / Shape mode */
var __shapeMode = 'square'; // 'square' | 'circle' | 'hexA' | 'hexB'
/* アピアランス設定 / Appearance settings */
var __appearanceMode = 'strokeOnly'; // 'strokeOnly' | 'clipGroup'
var __roundCornersEnabled = false;
var __roundCornerRadius = 0;
var __sameSizeEnabled = false; // 「大きさを揃える」
var __sameSizeMode = 'max'; // 'max' | 'min'  （最大/最小）


/*
正六角形パスを作成 / Create a regular hexagon path
- centerX/centerY を中心に、半径 r の外接円上に6点を配置
- rotationDeg で回転（0=フラットトップ寄り、30=ポイントトップ寄り）
*/
function createHexagonPath(targetLayer, centerX, centerY, r, rotationDeg) {
    var pts = [];
    var base = rotationDeg * Math.PI / 180;
    for (var i = 0; i < 6; i++) {
        var a = base + (Math.PI / 3) * i; // 60deg step
        var x = centerX + r * Math.cos(a);
        var y = centerY + r * Math.sin(a);
        pts.push([x, y]);
    }
    var p = targetLayer.pathItems.add();
    p.setEntirePath(pts);
    p.closed = true;
    p.stroked = false;
    p.filled = false;
    return p;
}

/* 角丸 LiveEffect / Round corners LiveEffect */
function createRoundCornersEffectXML(radius) {
    var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
    return xml.replace('#value#', radius);
}

function applyRoundCornersLiveEffect(targetItem, radius) {
    if (!targetItem) return false;
    var r = Number(radius);
    if (isNaN(r) || r < 0) return false;

    var xml = createRoundCornersEffectXML(r);

    try {
        targetItem.applyEffect(xml);
        return true;
    } catch (e) {
        return false;
    }
}

/*
作業用レイヤーを取得/作成 / Get or create a reusable work layer
- 名称 / Name: _clip_work
- 既存がロック/テンプレでもこのレイヤーは常に編集可能に設定
*/
function getOrCreateWorkLayer(doc) {
    var name = "_clip_work";
    var lyr = null;
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === name) {
            lyr = doc.layers[i];
            break;
        }
    }
    if (!lyr) {
        lyr = doc.layers.add();
        lyr.name = name;
    }
    // ensure editable
    lyr.locked = false;
    lyr.visible = true;
    lyr.isTemplate = false;
    return lyr;
}

// bounds utilities / バウンディング取得ユーティリティ
function __getItemBounds(item) {
    try {
        if (item && item.visibleBounds) return item.visibleBounds;   // [L, T, R, B]
    } catch (e) { }
    try {
        if (item && item.geometricBounds) return item.geometricBounds; // [L, T, R, B]
    } catch (e2) { }
    return null;
}

function __getBoundsSize(b) {
    if (!b || b.length !== 4) return null;
    var w = b[2] - b[0];
    var h = b[1] - b[3];
    return { w: Math.abs(w), h: Math.abs(h) };
}

// 選択内で最大/最小のオブジェクトに合わせてリサイズ（幅・高さを一致させる）
function __resizeSelectionToTarget(sel, mode) {
    if (!sel || sel.length < 2) return;
    var useMin = (mode === 'min');

    // 最大/最小（面積）を探す
    var targetW = 0, targetH = 0;
    var bestArea = useMin ? Infinity : -1;

    for (var i = 0; i < sel.length; i++) {
        var it = sel[i];
        // このスクリプトが扱うタイプ中心でOK（その他は無視）
        if (!(it.typename === 'PlacedItem' || it.typename === 'RasterItem' || it.typename === 'GroupItem')) continue;

        var b = __getItemBounds(it);
        var sz = __getBoundsSize(b);
        if (!sz) continue;

        var area = sz.w * sz.h;
        if (useMin) {
            if (area > 0 && area < bestArea) {
                bestArea = area;
                targetW = sz.w;
                targetH = sz.h;
            }
        } else {
            if (area > bestArea) {
                bestArea = area;
                targetW = sz.w;
                targetH = sz.h;
            }
        }
    }

    if ((useMin && !isFinite(bestArea)) || (!useMin && bestArea <= 0) || targetW <= 0 || targetH <= 0) return;

    // すべてをターゲットサイズへ（非等比も許容して一致させる）
    for (var j = 0; j < sel.length; j++) {
        var it2 = sel[j];
        if (!(it2.typename === 'PlacedItem' || it2.typename === 'RasterItem' || it2.typename === 'GroupItem')) continue;

        var b2 = __getItemBounds(it2);
        var sz2 = __getBoundsSize(b2);
        if (!sz2 || sz2.w === 0 || sz2.h === 0) continue;

        // すでに同サイズならスキップ
        if (Math.abs(sz2.w - targetW) < 0.001 && Math.abs(sz2.h - targetH) < 0.001) continue;

        var sx = (targetW / sz2.w) * 100;
        var sy = (targetH / sz2.h) * 100;

        try {
            it2.resize(sx, sy, true, true, true, true, true, Transformation.CENTER);
        } catch (e3) { }
    }
}

/*
画像に最小正方形を追加し、クリッピンググループを作成 / Build a clipping group with the minimal square
- 入力 / Input: doc (Document), image (PlacedItem|RasterItem), groups (Array)
- 動作 / Behavior: visibleBounds から正方形を作成→画像と同グループに配置→グループをクリップ化
*/
function processImage(doc, image) {
    var bounds = image.visibleBounds;
    var width = bounds[2] - bounds[0];
    var height = bounds[1] - bounds[3];

    // 基準サイズ / Base sizes
    var sideLength = Math.min(width, height); // 内接正方形の一辺 / inscribed square side
    var squareSize = sideLength; // 内接正方形の一辺 / inscribed square side
    var centerX = bounds[0] + width / 2;
    var centerY = bounds[1] - height / 2;

    // 現在のレイヤーが編集可能か確認 / Check if current layer is editable
    var parentLayer = image.layer;
    var isLockedOrTemplate = parentLayer.locked || parentLayer.isTemplate;

    // 編集可能なレイヤーを使用 / Choose a writable layer
    var targetLayer = isLockedOrTemplate ? getOrCreateWorkLayer(doc) : parentLayer;

    // 形状作成 / Create shape
    var shapePath = null;

    if (__shapeMode === 'circle') {
        // 正円（円） / Circle (perfect circle)
        shapePath = targetLayer.pathItems.ellipse(centerY + sideLength / 2, centerX - sideLength / 2, sideLength, sideLength);
        shapePath.stroked = false;
        shapePath.filled = false;
    } else if (__shapeMode === 'hexA') {
        // 六角形A / Hexagon A (rotation 0deg)
        shapePath = createHexagonPath(targetLayer, centerX, centerY, sideLength / 2, 0);
    } else if (__shapeMode === 'hexB') {
        // 六角形B / Hexagon B (rotation 30deg)
        shapePath = createHexagonPath(targetLayer, centerX, centerY, sideLength / 2, 30);
    } else {
        // 正方形 / Square
        shapePath = targetLayer.pathItems.rectangle(centerY + squareSize / 2, centerX - squareSize / 2, squareSize, squareSize);
        shapePath.stroked = false;
        shapePath.filled = false;
    }

    var group = targetLayer.groupItems.add();
    image.moveToBeginning(group);
    shapePath.moveToBeginning(group);
    group.clipped = true;
    group.selected = true; // 生成直後に即選択 / Select immediately after creation
    return group;
}

function main() {
    // 安全ガード / Safety guard
    if (!app.documents.length || !app.selection.length) {
        return;
    }
    var doc = app.activeDocument;
    var selectedItems = app.selection;
    // 「大きさを揃える」ONなら、選択内で最大/最小のオブジェクトにサイズを合わせる
    if (__sameSizeEnabled) {
        __resizeSelectionToTarget(selectedItems, __sameSizeMode);
    }
    doc.selection = null; // 先に選択をクリア / Clear selection first
    var createdGroups = [];

    for (var i = 0; i < selectedItems.length; i++) {
        var item = selectedItems[i];

        // クリッピングマスクの処理 / Handle clipping mask groups
        if (item.typename === 'GroupItem' && item.clipped) {
            item.clipped = false;

            var itemsToProcess = [];
            for (var j = item.pageItems.length - 1; j >= 0; j--) {
                var pageItem = item.pageItems[j];
                if (pageItem.typename === 'PlacedItem' || pageItem.typename === 'RasterItem') {
                    itemsToProcess.push(pageItem);
                } else {
                    pageItem.remove();
                }
            }

            // 解除後のグループ内にある画像を処理 / Process images extracted from the released group
            for (var k = 0; k < itemsToProcess.length; k++) {
                var imageItem = itemsToProcess[k];
                var g1 = processImage(doc, imageItem);
                if (g1) createdGroups.push(g1);
            }
        }
        // 単体の配置/埋め込み画像を処理 / Handle standalone placed/embedded images
        else if (item.typename === 'PlacedItem' || item.typename === 'RasterItem') {
            var g2 = processImage(doc, item);
            if (g2) createdGroups.push(g2);
        }
    }

    // クリップグループ後の処理 / Post process for clipping groups
    if (createdGroups.length > 0) {
        var prevSel = createdGroups.slice(0);

        // 角丸（チェック時のみ）
        if (__roundCornersEnabled) {
            for (var r = 0; r < createdGroups.length; r++) {
                applyRoundCornersLiveEffect(createdGroups[r], __roundCornerRadius);
            }
        }

        if (__appearanceMode === 'clipGroup') {
            // 「ケイ線を追加」OFF の場合は、クリップグループ作成のみ（ケイ線は付けない）
        } else {
            // 「ケイ線を追加」ON：ケイ線を追加 + Exclude
            for (var s = 0; s < createdGroups.length; s++) {
                try {
                    doc.selection = [createdGroups[s]];
                    app.executeMenuCommand('Adobe New Stroke Shortcut');
                    app.executeMenuCommand('Live Pathfinder Exclude');
                } catch (e) { }
            }
        }

        doc.selection = prevSel;
    }
}

/* ダイアログボックス / Dialog */
function showDialog() {
    var dlg = new Window('dialog', 'クリップグループの形状変更 ' + SCRIPT_VERSION);
    dlg.alignChildren = 'fill';

    // 2カラム / Two columns
    var cols = dlg.add('group');
    cols.orientation = 'row';
    cols.alignChildren = ['fill', 'top'];
    cols.spacing = 10;

    // 左：形状 / Left: Shape
    var shapePanel = cols.add('panel', undefined, '形状');
    shapePanel.orientation = 'column';
    shapePanel.alignChildren = 'left';
    shapePanel.margins = [15, 20, 15, 10];

    var rbNoChange = shapePanel.add('radiobutton', undefined, '変更なし');
    rbNoChange.value = true;

    var rbSquare = shapePanel.add('radiobutton', undefined, '正方形');
    var rbCircle = shapePanel.add('radiobutton', undefined, '正円');
    var rbHexA = shapePanel.add('radiobutton', undefined, '六角形A');
    var rbHexB = shapePanel.add('radiobutton', undefined, '六角形B');

    // 右カラム / Right column
    var rightCol = cols.add('group');
    rightCol.orientation = 'column';
    rightCol.alignChildren = ['fill', 'top'];
    rightCol.spacing = 10;

    // 右：アピアランス / Right: Appearance
    var appearancePanel = rightCol.add('panel', undefined, 'アピアランス');
    appearancePanel.orientation = 'column';
    appearancePanel.alignChildren = ['left', 'top'];
    appearancePanel.margins = [15, 20, 15, 10];

    // オプションパネル / Options panel（アピアランスの外）
    var optionPanel = rightCol.add('panel', undefined, '複数オブジェクト');
    optionPanel.orientation = 'column';
    optionPanel.alignChildren = ['left', 'top'];
    optionPanel.margins = [15, 20, 15, 10];

    // 選択が1つ以下なら「複数オブジェクト」パネルはディム表示
    var selCount = 0;
    try {
        if (app.documents.length && app.activeDocument.selection) {
            selCount = app.activeDocument.selection.length;
        }
    } catch (e) { selCount = 0; }
    optionPanel.enabled = (selCount > 1);

    var cbSameSize = optionPanel.add('checkbox', undefined, '大きさを揃える');
    cbSameSize.value = false;

    var sameSizeModeGroup = optionPanel.add('group');
    sameSizeModeGroup.orientation = 'column';
    sameSizeModeGroup.alignChildren = 'left';
    sameSizeModeGroup.margins = [20, 0, 0, 0];

    var rbSameSizeMax = sameSizeModeGroup.add('radiobutton', undefined, '最大');
    var rbSameSizeMin = sameSizeModeGroup.add('radiobutton', undefined, '最小');
    rbSameSizeMax.value = true;

    function updateSameSizeModeUI() {
        sameSizeModeGroup.enabled = (cbSameSize.value === true);
    }
    cbSameSize.onClick = updateSameSizeModeUI;
    updateSameSizeModeUI();

    var cbAddStrokeOnly = appearancePanel.add('checkbox', undefined, 'ケイ線を追加');

    // 選択オブジェクト全体の外接矩形から角丸のデフォルト値を算出（PlacedImageStroke.jsx）
    function calcDefaultRoundRadiusFromSelection(sel) {
        try {
            if (!sel || sel.length === 0) return 10;

            var top = -Infinity, left = Infinity, bottom = Infinity, right = -Infinity;
            for (var i = 0, n = sel.length; i < n; i++) {
                var b = sel[i].geometricBounds; // [top, left, bottom, right]
                if (!b || b.length !== 4) continue;
                if (b[0] > top) top = b[0];
                if (b[1] < left) left = b[1];
                if (b[2] < bottom) bottom = b[2];
                if (b[3] > right) right = b[3];
            }

            if (!isFinite(top) || !isFinite(left) || !isFinite(bottom) || !isFinite(right)) return 10;

            var height = Math.abs(top - bottom);
            var width = Math.abs(right - left);
            var A = height + width;
            var B = A / 2;
            var r = Math.max(0, B / 20);
            return r;
        } catch (e) {
            return 10;
        }
    }

    var defaultRoundRadius = 10;
    try {
        if (app.documents.length) {
            defaultRoundRadius = calcDefaultRoundRadiusFromSelection(app.activeDocument.selection);
        }
    } catch (e) { }
    defaultRoundRadius = Math.round(defaultRoundRadius);

    // 角丸（クリップグループ用）
    var roundRow = appearancePanel.add('group');
    roundRow.orientation = 'row';
    roundRow.alignChildren = ['left', 'center'];
    roundRow.margins = [0, 0, 0, 0];

    var cbRoundCorners = roundRow.add('checkbox', undefined, '角丸');
    cbRoundCorners.value = false;

    var editRoundRadius = roundRow.add('edittext', undefined, String(defaultRoundRadius));
    editRoundRadius.characters = 6;
    var roundUnit = roundRow.add('statictext', undefined, 'pt');

    // default
    cbAddStrokeOnly.value = false;


    function updateRoundRadiusUI() {
        // 現在の形状選択をUIから判定（「変更なし」の場合は現在値を参照）
        var effectiveShapeMode = __shapeMode;
        if (rbCircle.value) effectiveShapeMode = 'circle';
        else if (rbHexA.value) effectiveShapeMode = 'hexA';
        else if (rbHexB.value) effectiveShapeMode = 'hexB';
        else if (rbSquare.value) effectiveShapeMode = 'square';
        // rbNoChange の場合は __shapeMode のまま

        // 正円の場合は角丸は無効（ディム表示）にする
        if (effectiveShapeMode === 'circle') {
            cbRoundCorners.enabled = false;
            editRoundRadius.enabled = false;
            roundUnit.enabled = false;
            return;
        }

        // 角丸チェックは操作可能（ケイ線ON/OFFではディムにしない）
        cbRoundCorners.enabled = true;

        // 角丸チェックに連動して数値だけ enable/disable
        var radiusEnabled = (cbRoundCorners.value === true);
        editRoundRadius.enabled = radiusEnabled;
        roundUnit.enabled = radiusEnabled;
    }

    // UI変更時にUI状態を更新
    cbAddStrokeOnly.onClick = updateRoundRadiusUI;
    rbNoChange.onClick = updateRoundRadiusUI;
    rbSquare.onClick = updateRoundRadiusUI;
    rbCircle.onClick = updateRoundRadiusUI;
    rbHexA.onClick = updateRoundRadiusUI;
    rbHexB.onClick = updateRoundRadiusUI;
    cbRoundCorners.onClick = updateRoundRadiusUI;

    updateRoundRadiusUI();

    editRoundRadius.onChange = function () {
        var v = Number(editRoundRadius.text);
        if (isNaN(v)) return;
        editRoundRadius.text = String(v);
    };

    // ↑↓キーでの値変更（Shift: ±10 / Option: ±0.1）
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                if (event.keyName == "Up") {
                    value = Math.ceil((value + 1) / delta) * delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value = Math.floor((value - 1) / delta) * delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            } else if (keyboard.altKey) {
                delta = 0.1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    event.preventDefault();
                }
            } else {
                delta = 1;
                if (event.keyName == "Up") {
                    value += delta;
                    event.preventDefault();
                } else if (event.keyName == "Down") {
                    value -= delta;
                    if (value < 0) value = 0;
                    event.preventDefault();
                }
            }

            if (keyboard.altKey) {
                value = Math.round(value * 10) / 10;
            } else {
                value = Math.round(value);
            }

            editText.text = value;
        });
    }
    changeValueByArrowKey(editRoundRadius);

    // ボタン / Buttons
    var btnGroup = dlg.add('group');
    btnGroup.alignment = 'right';
    var btnCancel = btnGroup.add('button', undefined, 'キャンセル', { name: 'cancel' });
    var btnOk = btnGroup.add('button', undefined, 'OK', { name: 'ok' });

    btnOk.onClick = function () {
        if (rbNoChange.value) {
            // keep current __shapeMode
        } else if (rbCircle.value) {
            __shapeMode = 'circle';
        } else if (rbHexA.value) {
            __shapeMode = 'hexA';
        } else if (rbHexB.value) {
            __shapeMode = 'hexB';
        } else {
            __shapeMode = 'square';
        }

        // アピアランス設定の確定 / Confirm appearance settings
        __appearanceMode = cbAddStrokeOnly.value ? 'strokeOnly' : 'clipGroup';
        __sameSizeEnabled = (cbSameSize.value === true);
        __sameSizeMode = (rbSameSizeMin.value === true) ? 'min' : 'max';

        // 正円は角丸無効 + ケイ線ON時は角丸適用しない
        __roundCornersEnabled = (__shapeMode !== 'circle') && (cbRoundCorners.value === true);
        __roundCornerRadius = __roundCornersEnabled ? Number(editRoundRadius.text) : 0;
        if (__roundCornersEnabled && (isNaN(__roundCornerRadius) || __roundCornerRadius < 0)) {
            __roundCornerRadius = 0;
        }

        dlg.close(1);
    };

    btnCancel.onClick = function () {
        dlg.close(0);
    };

    return dlg.show() === 1;
}

if (showDialog()) {
    main();
}