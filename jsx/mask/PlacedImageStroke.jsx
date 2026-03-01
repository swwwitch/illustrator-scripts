#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

function main() {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    // --- UI ---
    var dlg = new Window('dialog', '配置画像へのケイ線追加');
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];
    dlg.margins = 15;

    var appearancePanel = dlg.add('panel', undefined, 'アピアランス');
    appearancePanel.orientation = 'column';
    appearancePanel.alignChildren = ['left', 'top'];
    appearancePanel.margins = [15, 20, 15, 10];

    var rbAddStrokeOnly = appearancePanel.add('radiobutton', undefined, 'ケイ線のみを追加');
    var rbClipGroupThen = appearancePanel.add('radiobutton', undefined, 'クリップグループ');

    // 選択オブジェクト全体の外接矩形から角丸のデフォルト値を算出
    // A = 高さ + 幅
    // B = A / 2
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

    var defaultRoundRadius = calcDefaultRoundRadiusFromSelection(doc.selection);
    defaultRoundRadius = Math.round(defaultRoundRadius * 100) / 100;

    // 角丸（クリップグループ用）
    var roundRow = appearancePanel.add('group');
    roundRow.orientation = 'row';
    roundRow.alignChildren = ['left', 'center'];
    roundRow.margins = [20, 0, 0, 0]; // インデントして「クリップグループ」の下に見せる

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
    rbAddStrokeOnly.value = true;

    function updateRoundRadiusUI() {
        var clipEnabled = rbClipGroupThen.value === true;
        cbRoundCorners.enabled = clipEnabled;

        var radiusEnabled = clipEnabled && cbRoundCorners.value === true;
        editRoundRadius.enabled = radiusEnabled;
        roundUnit.enabled = radiusEnabled;
    }

    rbAddStrokeOnly.onClick = updateRoundRadiusUI;
    rbClipGroupThen.onClick = updateRoundRadiusUI;
    cbRoundCorners.onClick = updateRoundRadiusUI;
    updateRoundRadiusUI();

    editRoundRadius.onChange = function () {
        var v = Number(editRoundRadius.text);
        if (isNaN(v)) return; // 無理に直さない
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
                // Shiftキー押下時は10の倍数にスナップ
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
                // Optionキー押下時は0.1単位で増減
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
                // 小数第1位までに丸め
                value = Math.round(value * 10) / 10;
            } else {
                // 整数に丸め
                value = Math.round(value);
            }

            editText.text = value;
        });
    }

    changeValueByArrowKey(editRoundRadius);

    var btns = dlg.add('group');
    btns.orientation = 'row';
    btns.alignChildren = ['center', 'center'];
    btns.alignment = 'center';


    var cancelBtn = btns.add('button', undefined, 'キャンセル', { name: 'cancel' }); var okBtn = btns.add('button', undefined, 'OK', { name: 'ok' });

    if (dlg.show() !== 1) {
        return;
    }

    if (!doc.selection || doc.selection.length === 0) {
        alert("オブジェクトを選択してください。");
        return;
    }

    // --- Clipping helpers (ported from MakeClippingMask.jsx) ---
    function isImageItem(item) {
        return item && (item.typename === 'PlacedItem' || item.typename === 'RasterItem');
    }

    // 最前面のパス（zOrderPosition が最大）を取得
    function getFrontmostPath(pathArray) {
        var topPath = null;
        var highestZ = -1;

        for (var i = 0; i < pathArray.length; i++) {
            var item = pathArray[i];
            if (item && item.typename === 'PathItem' && item.zOrderPosition > highestZ) {
                highestZ = item.zOrderPosition;
                topPath = item;
            }
        }

        return topPath;
    }

    // 画像1つ → 画像外接の矩形でクリップ
    function createClippingMaskGroup(imageItem) {
        var targetLayer = imageItem.layer;
        var wasLocked = targetLayer.locked;
        var wasVisible = targetLayer.visible;
        var wasTemplate = targetLayer.isTemplate;

        if (wasLocked) targetLayer.locked = false;
        if (!wasVisible) targetLayer.visible = true;
        if (wasTemplate) targetLayer.isTemplate = false;

        var rect = targetLayer.pathItems.rectangle(
            imageItem.top,
            imageItem.left,
            imageItem.width,
            imageItem.height
        );
        rect.stroked = false;
        rect.filled = false;

        var groupItem = targetLayer.groupItems.add();
        imageItem.moveToBeginning(groupItem);
        rect.moveToBeginning(groupItem);
        groupItem.clipped = true;

        if (wasLocked) targetLayer.locked = true;
        if (!wasVisible) targetLayer.visible = false;
        if (wasTemplate) targetLayer.isTemplate = true;

        return groupItem;
    }

    // 画像1つ + パス1つ → パスでクリップ
    function createMaskWithPath(imageItem, pathItem) {
        var targetLayer = imageItem.layer;
        if (pathItem.layer != targetLayer) {
            pathItem.move(targetLayer, ElementPlacement.PLACEATBEGINNING);
        }

        var groupItem = targetLayer.groupItems.add();
        imageItem.moveToBeginning(groupItem);
        pathItem.moveToBeginning(groupItem);
        groupItem.clipped = true;

        return groupItem;
    }

    // 現在の選択から「クリップグループ」を作成し、選択を更新
    // 戻り値: 作成した groupItem（複数作成時は先頭を返す） / 作成できない場合は null
    function makeClippingFromSelection(doc) {
        var sel = doc.selection;
        if (!sel || sel.length === 0) return null;

        // すでにクリップグループが1つ選択されているなら、そのまま
        if (sel.length === 1 && sel[0].typename === 'GroupItem' && sel[0].clipped) {
            return sel[0];
        }

        var images = [];
        var paths = [];
        var i;

        for (i = 0; i < sel.length; i++) {
            if (!sel[i]) continue;
            if (isImageItem(sel[i])) {
                images.push(sel[i]);
            } else if (sel[i].typename === 'PathItem') {
                paths.push(sel[i]);
            }
        }

        // 画像1 + パス1（選択がちょうど2つ）
        if (sel.length === 2 && images.length === 1 && paths.length === 1) {
            var g1 = createMaskWithPath(images[0], paths[0]);
            doc.selection = [g1];
            return g1;
        }

        // 画像1のみ
        if (sel.length === 1 && images.length === 1) {
            var g2 = createClippingMaskGroup(images[0]);
            doc.selection = [g2];
            return g2;
        }

        // パスが含まれている場合は、最前面パスをマスクとして全体をクリップ
        if (paths.length > 0) {
            var maskPath = getFrontmostPath(paths);
            if (!maskPath) return null;

            var targetLayer = maskPath.layer;
            var wasLocked = targetLayer.locked;
            var wasVisible = targetLayer.visible;
            var wasTemplate = targetLayer.isTemplate;

            if (wasLocked) targetLayer.locked = false;
            if (!wasVisible) targetLayer.visible = true;
            if (wasTemplate) targetLayer.isTemplate = false;

            var groupItem = targetLayer.groupItems.add();

            // マスク以外を先に移動（順序を保ちやすいよう末尾から）
            for (i = sel.length - 1; i >= 0; i--) {
                if (!sel[i]) continue;
                if (sel[i] === maskPath) continue;
                sel[i].moveToBeginning(groupItem);
            }
            maskPath.moveToBeginning(groupItem);
            groupItem.clipped = true;

            if (wasLocked) targetLayer.locked = true;
            if (!wasVisible) targetLayer.visible = false;
            if (wasTemplate) targetLayer.isTemplate = true;

            doc.selection = [groupItem];
            return groupItem;
        }

        // 複数画像のみ → それぞれ外接矩形で個別にクリップ
        if (images.length > 0) {
            var groups = [];
            for (i = 0; i < images.length; i++) {
                try {
                    groups.push(createClippingMaskGroup(images[i]));
                } catch (e) { }
            }
            if (groups.length > 0) {
                doc.selection = groups;
                return groups[0];
            }
        }

        return null;
    }

    // クリップグループ内の「クリッピングパス」を取得（見つからなければ null）
    function getClippingPathFromGroup(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== 'GroupItem') return null;

            // groupItem.pathItems はグループ内のパスを横断して取れることが多い
            // その中で clipping=true のものを探す
            var paths = groupItem.pathItems;
            if (paths && paths.length) {
                for (var i = 0; i < paths.length; i++) {
                    if (paths[i] && paths[i].typename === 'PathItem' && paths[i].clipping) {
                        return paths[i];
                    }
                }
            }
        } catch (e) { }
        return null;
    }

    // 角丸 LiveEffect XML を生成（シンプル版）
    function createRoundCornersEffectXML(radius) {
        var xml = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius #value# "/></LiveEffect>';
        return xml.replace('#value#', radius);
    }

    // 角丸（LiveEffect: Adobe Round Corners）を適用
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

    try {
        if (rbClipGroupThen.value) {
            var result = makeClippingFromSelection(doc);
            if (!result) {
                alert('クリップグループ化できる選択ではありません。\n（例：画像1つ、または 画像1つ+パス1つ、または 複数オブジェクト+パス）');
                return;
            }
        }

        if (rbAddStrokeOnly.value) {
            // --- Existing behavior: add stroke then outline ---
            app.executeMenuCommand('Adobe New Stroke Shortcut');
            app.executeMenuCommand('Live Outline Object');
        } else {
            // クリップグループ後の処理：角丸（LiveEffect）を適用（チェックOFFならスルー）
            var doRound = (cbRoundCorners.value === true);
            var ROUND_RADIUS = null;

            if (doRound) {
                ROUND_RADIUS = Number(editRoundRadius.text);
                if (isNaN(ROUND_RADIUS) || ROUND_RADIUS < 0) {
                    alert('角丸の値が不正です。0以上の数値を入力してください。');
                    return;
                }
            }

            // 選択がクリップグループ（単体 or 複数）になっている想定
            var targets = doc.selection;
            if (targets && targets.length) {
                for (var i = 0; i < targets.length; i++) {
                    var g = targets[i];
                    if (!g || g.typename !== 'GroupItem' || !g.clipped) continue;

                    if (doRound) {
                        applyRoundCornersLiveEffect(g, ROUND_RADIUS);
                    }

                    // メニューコマンドは「単体選択」で実行（複数選択だと失敗することがある）
                    doc.selection = [g];
                    app.executeMenuCommand('Adobe New Stroke Shortcut');
                    app.executeMenuCommand('Live Pathfinder Exclude');
                }

                // 選択を元に戻す
                doc.selection = targets;
            }
        }

    } catch (e) {
        alert("コマンドの実行中にエラーが発生しました。\n" + e);
    }
}

main();