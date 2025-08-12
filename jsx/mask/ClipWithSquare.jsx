#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

ClipWithSquare.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/mask/blob/main/ClipWithSquare.jsx

### 概要：

- 選択された画像（配置画像/埋め込み画像）やクリッピングマスクグループ内の画像に対して、中心を基準とした最小の正方形パスを生成
- 生成した正方形で新しいクリッピンググループを作成

### 主な機能：

- 配置画像、埋め込み画像、およびそれらを含むクリッピングマスクの処理
- ロックレイヤーやテンプレートレイヤー上の画像も対象（作業用レイヤーを作成して処理）

### 処理の流れ：

1) 選択オブジェクトを走査
2) クリッピングマスクなら解除して画像のみ抽出
3) visibleBounds から最小正方形を作成
4) 正方形と画像を同一グループに入れてクリッピング化
5) 作成されたグループを選択状態に

### 更新履歴：

- v1.0 (20250403) : 初期バージョン
- v1.1 (20250813) : クリッピング解除→再構築の安定化、コメント整理
*/

/*
### Script Name:

ClipWithSquare.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Generates the smallest possible square path centered on selected placed/embedded images or images within clipping mask groups
- Creates a new clipping group using the generated square

### Main Features:

- Handles placed images, embedded images, and clipping masks containing them
- Includes images on locked/template layers by creating a work layer for processing

### Process Flow:

1) Iterate through selected objects
2) If a clipping mask, release it and extract only the images
3) Create the smallest square from visibleBounds
4) Place the square and image in the same group and clip
5) Select the created group

### Changelog:

- v1.0 (20250403) : Initial version
- v1.1 (20250813) : Stabilized release→rebuild of clipping, comment cleanup
*/

// スクリプトバージョン
var SCRIPT_VERSION = "v1.1";
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
    var sideLength = Math.min(width, height);
    var centerX = bounds[0] + width / 2;
    var centerY = bounds[1] - height / 2;

    // 現在のレイヤーが編集可能か確認 / Check if current layer is editable
    var parentLayer = image.layer;
    var isLockedOrTemplate = parentLayer.locked || parentLayer.isTemplate;

    // 編集可能なレイヤーを使用 / Choose a writable layer
    var targetLayer = isLockedOrTemplate ? getOrCreateWorkLayer(doc) : parentLayer;

    // 四角形とグループをレイヤーに追加 / Create square and group
    var square = targetLayer.pathItems.rectangle(centerY + sideLength / 2, centerX - sideLength / 2, sideLength, sideLength);
    var group = targetLayer.groupItems.add();

    image.moveToBeginning(group);
    square.moveToBeginning(group);
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

main();