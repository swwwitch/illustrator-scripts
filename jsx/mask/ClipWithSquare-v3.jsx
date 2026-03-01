#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";

/* 形状モード / Shape mode */
var __shapeMode = 'square'; // 'square' | 'circle' | 'hexA' | 'hexB'
/* アピアランス設定 / Appearance settings */
var __appearanceMode = 'strokeOnly'; // 'strokeOnly' | 'clipGroup'
var __roundCornersEnabled = false;
var __roundCornerRadius = 0;

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
    return group;
}

function main() {
    // 安全ガード / Safety guard
    if (!app.documents.length || !app.selection.length) {
        return;
    }
    var doc = app.activeDocument;
    var selectedItems = app.selection;
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

        if (__appearanceMode === 'clipGroup') {
            // 角丸（チェック時のみ）
            if (__roundCornersEnabled) {
                for (var r = 0; r < createdGroups.length; r++) {
                    applyRoundCornersLiveEffect(createdGroups[r], __roundCornerRadius);
                }
            }

            // クリップグループ向け（ケイ線追加 + Exclude）
            for (var m = 0; m < createdGroups.length; m++) {
                try {
                    doc.selection = [createdGroups[m]];
                    app.executeMenuCommand('Adobe New Stroke Shortcut');
                    app.executeMenuCommand('Live Pathfinder Exclude');
                } catch (e) { }
            }
        } else {
            // 「ケイ線のみを追加」：マスク（クリップパス）変更後に適用
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
    var dlg = new Window('dialog', 'ClipWithSquare ' + SCRIPT_VERSION);
    dlg.alignChildren = 'fill';

    // 形状パネル / Shape panel
    var shapePanel = dlg.add('panel', undefined, '形状');
    shapePanel.orientation = 'column';
    shapePanel.alignChildren = 'left';
    shapePanel.margins = [15, 20, 15, 10];

    var rbNoChange = shapePanel.add('radiobutton', undefined, '変更なし');
    rbNoChange.value = true;

    var rbSquare = shapePanel.add('radiobutton', undefined, '正方形');
    var rbCircle = shapePanel.add('radiobutton', undefined, '正円');
    var rbHexA = shapePanel.add('radiobutton', undefined, '六角形A');
    var rbHexB = shapePanel.add('radiobutton', undefined, '六角形B');

    // アピアランスパネル / Appearance panel
    var appearancePanel = dlg.add('panel', undefined, 'アピアランス');
    appearancePanel.orientation = 'column';
    appearancePanel.alignChildren = ['left', 'top'];
    appearancePanel.margins = [15, 20, 15, 10];

    var cbAddStrokeOnly = appearancePanel.add('checkbox', undefined, 'ケイ線のみを追加');
    // var rbClipGroupThen = appearancePanel.add('radiobutton', undefined, 'クリップグループ');

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
            var width  = Math.abs(right - left);
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
    defaultRoundRadius = Math.round(defaultRoundRadius * 100) / 100;

    // 角丸（クリップグループ用）
    var roundRow = appearancePanel.add('group');
    roundRow.orientation = 'row';
    roundRow.alignChildren = ['left', 'center'];
    // roundRow.margins = [20, 0, 0, 0]; // インデントして「クリップグループ」の下に見せる
    roundRow.margins = [0, 0, 0, 0];

    var cbRoundCorners = roundRow.add('checkbox', undefined, '角丸');
    cbRoundCorners.value = true;

    var editRoundRadius = roundRow.add('edittext', undefined, String(defaultRoundRadius));
    editRoundRadius.characters = 6;
    var roundUnit = roundRow.add('statictext', undefined, 'pt');

    // 初期状態（ラジオに連動して有効/無効）
    cbRoundCorners.enabled = false;
    editRoundRadius.enabled = false;
    roundUnit.enabled = false;

    // default
    cbAddStrokeOnly.value = false;
    // rbClipGroupThen.value = true;

    function updateRoundRadiusUI() {
        // var clipEnabled = rbClipGroupThen.value === true;
        var clipEnabled = (cbAddStrokeOnly.value !== true); // クリップグループ相当（ケイ線のみOFF）
        cbRoundCorners.enabled = clipEnabled;

        var radiusEnabled = clipEnabled && cbRoundCorners.value === true;
        editRoundRadius.enabled = radiusEnabled;
        roundUnit.enabled = radiusEnabled;
    }

    cbAddStrokeOnly.onClick = updateRoundRadiusUI;
    // rbClipGroupThen.onClick = updateRoundRadiusUI;
    cbRoundCorners.onClick = updateRoundRadiusUI;
    updateRoundRadiusUI();

    editRoundRadius.onChange = function () {
        var v = Number(editRoundRadius.text);
        if (isNaN(v)) return;
        editRoundRadius.text = String(v);
    };

    // ↑↓キーでの値変更（Shift: ±10 / Option: ±0.1）
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function(event) {
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
        // __roundCornersEnabled = (rbClipGroupThen.value && cbRoundCorners.value);
        __roundCornersEnabled = (!cbAddStrokeOnly.value && cbRoundCorners.value);
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