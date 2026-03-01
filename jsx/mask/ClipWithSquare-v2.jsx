#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";

/* 形状モード / Shape mode */
var __shapeMode = 'square'; // 'square' | 'circle' | 'hexA' | 'hexB'
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
    // 正円は従来通り「内接（min）」基準に戻す / Circle: revert to inscribed (min) behavior
    var sideLength = Math.min(width, height); // 内接正方形の一辺 / inscribed square side
    // 正方形も従来通り「内接（min）」基準に戻す / Square: revert to inscribed (min) behavior
    var squareSize = sideLength; // 内接正方形の一辺 / inscribed square side
    var centerX = bounds[0] + width / 2;
    var centerY = bounds[1] - height / 2;

    // 現在のレイヤーが編集可能か確認 / Check if current layer is editable
    var parentLayer = image.layer;
    var isLockedOrTemplate = parentLayer.locked || parentLayer.isTemplate;

    // 編集可能なレイヤーを使用 / Choose a writable layer
    var targetLayer = isLockedOrTemplate ? getOrCreateWorkLayer(doc) : parentLayer;

    // 四角形とグループをレイヤーに追加 / Create square and group
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
}

function main() {
    // 安全ガード / Safety guard
    if (!app.documents.length || !app.selection.length) {
        return;
    }
    var doc = app.activeDocument;
    var selectedItems = app.selection;
    doc.selection = null; // 先に選択をクリア / Clear selection first

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
                processImage(doc, imageItem);
            }
        }
        // 単体の配置/埋め込み画像を処理 / Handle standalone placed/embedded images
        else if (item.typename === 'PlacedItem' || item.typename === 'RasterItem') {
            processImage(doc, item);
        }
    }
}

/* ダイアログボックス / Dialog */
function showDialog() {
    var dlg = new Window('dialog', 'ClipWithSquare ' + SCRIPT_VERSION);
    dlg.alignChildren = 'fill';

    // 形状パネル / Shape panel
    var shapePanel = dlg.add('panel', undefined, '形状');
    shapePanel.orientation = 'column';
    shapePanel.alignChildren = 'left';
    shapePanel.margins = [15, 20, 15, 10];

    var rbSquare = shapePanel.add('radiobutton', undefined, '正方形');
    rbSquare.value = true; // 現状は正方形のみ
    var rbCircle = shapePanel.add('radiobutton', undefined, '正円');
    var rbHexA = shapePanel.add('radiobutton', undefined, '六角形A');
    var rbHexB = shapePanel.add('radiobutton', undefined, '六角形B');

    // ボタン / Buttons
    var btnGroup = dlg.add('group');
    btnGroup.alignment = 'right';
    var btnCancel = btnGroup.add('button', undefined, 'キャンセル', { name: 'cancel' });
    var btnOk = btnGroup.add('button', undefined, 'OK', { name: 'ok' });

    btnOk.onClick = function () {
        if (rbCircle.value) {
            __shapeMode = 'circle';
        } else if (rbHexA.value) {
            __shapeMode = 'hexA';
        } else if (rbHexB.value) {
            __shapeMode = 'hexB';
        } else {
            __shapeMode = 'square';
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