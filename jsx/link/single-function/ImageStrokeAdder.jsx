#target illustrator
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択した配置画像に、アピアランスとしてケイ線を追加します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImageStrokeAdder.md

### Overview

Adds a stroke to the selected placed images as an appearance.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImageStrokeAdder.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImageStrokeAdder";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                             /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ImageStrokeAdder.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ImageStrokeAdder.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * 現在のUI言語を判定する
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "配置画像へのケイ線追加", en: "Add Stroke to Placed Images" }
        },
        panel: {
            appearance: { ja: "アピアランス", en: "Appearance" }
        },
        radio: {
            addStrokeOnly: { ja: "ケイ線のみを追加", en: "Add stroke only" },
            clipGroupThen: { ja: "クリップグループ", en: "Clipping group" }
        },
        checkbox: {
            roundCorners: { ja: "角丸", en: "Rounded corners" }
        },
        tooltip: {
            addStrokeOnly: {
                ja: "配置画像にそのまま線を追加します。画像の形は変わりません。",
                en: "Adds a stroke straight to the placed image, leaving its shape alone."
            },
            clipGroupThen: {
                ja: "配置画像を長方形でクリップしてから、その長方形に線を追加します。角丸にできます。",
                en: "Clips the placed image with a rectangle and strokes that rectangle, so the corners can be rounded."
            },
            roundCorners: {
                ja: "クリップした長方形の角を丸めます。半径は右の欄で指定します。",
                en: "Rounds the corners of the clipping rectangle. The field on the right sets the radius."
            },
            roundRadius: { ja: "角丸の半径（pt）です。", en: "Radius of the rounded corners, in points." }
        }
    };

    /**
     * ラベルを取得する（ドット区切りキー）
     * @param {string} labelPath - "panel.appearance" のようなドット区切りキー
     * @returns {string} 現在のUI言語のラベル（見つからなければキーそのもの）
     */
    function getLabel(labelPath) {
        var pathKeys = String(labelPath).split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return (labelNode[uiLang] != null) ? labelNode[uiLang] : labelPath;
    }

    function main() {
        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。");
            return;
        }

        var doc = app.activeDocument;

        // --- UI ---
        var dialog = new Window('dialog', getLabel('dialog.title') + ' ' + SCRIPT_VERSION);
        dialog.orientation = 'column';
        dialog.alignChildren = ['fill', 'top'];
        dialog.margins = 15;

        var appearancePanel = dialog.add('panel', undefined, getLabel('panel.appearance'));
        appearancePanel.orientation = 'column';
        appearancePanel.alignChildren = ['left', 'top'];
        appearancePanel.margins = [15, 20, 15, 10];

        var rbAddStrokeOnly = appearancePanel.add('radiobutton', undefined, getLabel('radio.addStrokeOnly'));
        rbAddStrokeOnly.helpTip = getLabel('tooltip.addStrokeOnly');
        var rbClipGroupThen = appearancePanel.add('radiobutton', undefined, getLabel('radio.clipGroupThen'));
        rbClipGroupThen.helpTip = getLabel('tooltip.clipGroupThen');

        // 選択オブジェクト全体の外接矩形から角丸のデフォルト値を算出
        // A = 高さ + 幅
        // B = A / 2
        function calcDefaultRoundRadiusFromSelection(currentSelection) {
            try {
                if (!currentSelection || currentSelection.length === 0) return 10;

                var top = -Infinity, left = Infinity, bottom = Infinity, right = -Infinity;
                for (var i = 0, n = currentSelection.length; i < n; i++) {
                    var b = currentSelection[i].geometricBounds; // [top, left, bottom, right]
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

        var cbRoundCorners = roundRow.add('checkbox', undefined, getLabel('checkbox.roundCorners'));
        cbRoundCorners.helpTip = getLabel('tooltip.roundCorners');
        cbRoundCorners.value = true;

        var editRoundRadius = roundRow.add('edittext', undefined, String(defaultRoundRadius));
        editRoundRadius.helpTip = getLabel('tooltip.roundRadius');
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

        var btns = dialog.add('group');
        btns.orientation = 'row';
        btns.alignChildren = ['center', 'center'];
        btns.alignment = 'center';

        var cancelBtn = btns.add('button', undefined, 'キャンセル', { name: 'cancel' }); var okBtn = btns.add('button', undefined, 'OK', { name: 'ok' });

        if (dialog.show() !== 1) {
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
            var currentSelection = doc.selection;
            if (!currentSelection || currentSelection.length === 0) return null;

            // すでにクリップグループが1つ選択されているなら、そのまま
            if (currentSelection.length === 1 && currentSelection[0].typename === 'GroupItem' && currentSelection[0].clipped) {
                return currentSelection[0];
            }

            var images = [];
            var paths = [];
            var i;

            for (i = 0; i < currentSelection.length; i++) {
                if (!currentSelection[i]) continue;
                if (isImageItem(currentSelection[i])) {
                    images.push(currentSelection[i]);
                } else if (currentSelection[i].typename === 'PathItem') {
                    paths.push(currentSelection[i]);
                }
            }

            // 画像1 + パス1（選択がちょうど2つ）
            if (currentSelection.length === 2 && images.length === 1 && paths.length === 1) {
                var g1 = createMaskWithPath(images[0], paths[0]);
                doc.selection = [g1];
                return g1;
            }

            // 画像1のみ
            if (currentSelection.length === 1 && images.length === 1) {
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
                for (i = currentSelection.length - 1; i >= 0; i--) {
                    if (!currentSelection[i]) continue;
                    if (currentSelection[i] === maskPath) continue;
                    currentSelection[i].moveToBeginning(groupItem);
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

})();
