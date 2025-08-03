#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

RandomizeObjects.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/RandomizeObjects.jsx

### 概要：

- 選択したオブジェクトをランダムに移動・変形・回転・不透明度を変更するスクリプト
- UIから各種パラメータを指定し、即時プレビューで結果を確認可能

### 主な機能：

- 移動距離を横・縦方向に個別設定、または連動
- 中央揃えオプションでオブジェクトを一箇所に集約
- 拡大縮小（変形）、回転、不透明度のランダム化
- プレビュー反映、キャンセルで元の状態に完全リセット

### 処理の流れ：

1. ドキュメントと選択状態を確認
2. ダイアログを表示（移動距離・変形率・回転・不透明度を入力）
3. 入力値をもとにランダム変形や移動を即時プレビュー
4. OKで確定、キャンセルでダイアログ起動前にリセット

### 更新履歴：

- v1.0 (20250802): 初期バージョン
- v1.1 (20250803): 長方形を水平にする機能追加、UIの改善
- v1.2 (20250803): 強制的に中央に集める機能を追加
- v1.3 (20250804): ランダムボタンを追加

---

### Script Name:

RandomizeObjects.jsx

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/RandomizeObjects.jsx

### Overview:

- Randomly move, scale, rotate, and change opacity of selected objects
- Specify parameters via UI with immediate preview

### Main Features:

- Set horizontal/vertical move distances individually or linked
- Gather all objects to center option
- Random scaling, rotation, and opacity
- Live preview, cancel restores original state completely

### Process Flow:

1. Check document and selection
2. Show dialog (distance, scale, rotate, opacity)
3. Preview random transform based on inputs
4. OK confirms, Cancel restores state before dialog

### Update History:

- v1.0 (20250802): Initial version
- v1.1 (20250803): Added horizontal rectangle alignment, improved UI
- v1.2 (20250803): Added forced gather to center option
- v1.3 (20250804): Added random button for quick setup

*/

var SCRIPT_VERSION = "v1.3";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    dialogTitle: {
        ja: "ランダム化 " + SCRIPT_VERSION,
        en: "Randomize " + SCRIPT_VERSION
    },
    distance: {
        ja: "移動距離",
        en: "Distance"
    },
    horizontal: {
        ja: "横:",
        en: "Horizontal:"
    },
    vertical: {
        ja: "縦:",
        en: "Vertical:"
    },
    link: {
        ja: "連動",
        en: "Link"
    },
    center: {
        ja: "中央に集める",
        en: "Gather to Center"
    },
    avoidOverlap: {
        ja: "重なりを避ける",
        en: "Avoid Overlap"
    },
    scale: {
        ja: "スケール",
        en: "Scale"
    },
    rotate: {
        ja: "回転",
        en: "Rotate"
    },
    opacity: {
        ja: "不透明度",
        en: "Opacity"
    },
    ok: {
        ja: "OK",
        en: "OK"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },
    reset: {
        ja: "リセット",
        en: "Reset"
    },
    random: {
        ja: "ランダム",
        en: "Random"
    }
};

/* 入力値を更新し、リンクされた入力とプレビューを処理 / Update input, linked input, and preview */
function updateLinkedInputAndPreview(input, linkInput, previewFunc) {
    if (linkInput && linkInput.enabled) {
        linkInput.text = input.text;
    }
    if (previewFunc) previewFunc();
}
/* ダイアログ設定を共通化 / Configure dialog position and opacity */
function configureDialog(dlg, options) {
    if (options.opacity !== undefined) {
        dlg.opacity = options.opacity;
    }
}

function makeRectangleEdgeHorizontal(item) {
    if (item.typename === "PathItem" && item.closed && item.pathPoints.length === 4) {
        var p0 = item.pathPoints[0].anchor;
        var p1 = item.pathPoints[1].anchor;

        var dx = p1[0] - p0[0];
        var dy = p1[1] - p0[1];
        var angleRad = Math.atan2(dy, dx);
        var angleDeg = angleRad * 180 / Math.PI;

        item.rotate(-angleDeg);
    }
}


/* 矢印キーによる新しい値を返す補助関数 / Helper to calculate new value */
function getNewValueByKey(event, value, keyboard) {
    var delta = 1;
    if (keyboard.shiftKey) {
        delta = 10;
        if (event.keyName == "Up") {
            return Math.ceil((value + 1) / delta) * delta;
        } else if (event.keyName == "Down") {
            return Math.floor((value - 1) / delta) * delta;
        }
    } else if (keyboard.altKey) {
        delta = 0.1;
        if (event.keyName == "Up") {
            return value + delta;
        } else if (event.keyName == "Down") {
            return Math.max(0, value - delta);
        }
    } else {
        delta = 1;
        if (event.keyName == "Up") {
            return value + delta;
        } else if (event.keyName == "Down") {
            return Math.max(0, value - delta);
        }
    }
    return value;
}

/* 矢印キーで数値を増減 / Change numeric value with arrow keys */
function changeValueByArrowKey(editText, previewFunc) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var newValue = getNewValueByKey(event, value, keyboard);

        if (newValue !== value) {
            event.preventDefault();
            if (keyboard.altKey) {
                newValue = Math.round(newValue * 10) / 10;
            } else {
                newValue = Math.round(newValue);
            }
            editText.text = newValue;
            if (previewFunc) {
                updateLinkedInputAndPreview(editText, null, previewFunc);
            }
        }
    });
}

var savedBounds = $.global.__RandomizeObjectsBounds || null;

function storeWinBounds(dlg) {
    $.global.__RandomizeObjectsBounds = [
        dlg.bounds.left, dlg.bounds.top,
        dlg.bounds.right, dlg.bounds.bottom
    ];
}
var winBounds = savedBounds instanceof Array ? savedBounds : undefined;

/* メイン処理 / Main process */
function main() {
    try {

        if (app.documents.length === 0) {
            alert("ドキュメントが開かれていません。 / No document is open.");
            return;
        }

        var doc = app.activeDocument;
        if (doc.selection.length === 0) {
            alert("オブジェクトが選択されていません。 / No object is selected.");
            return;
        }

        var dialog = new Window("dialog", LABELS.dialogTitle[lang], winBounds);
        // ダイアログを開いたら Live Corner Annotator を実行
        try {
            app.executeMenuCommand('Live Corner Annotator');
        } catch (e) {}
        dialog.orientation = "row";
        dialog.alignChildren = "top";
        dialog.spacing = 20;

        // 左カラム
        var leftPanel = dialog.add("group");
        leftPanel.orientation = "column";
        leftPanel.alignChildren = "left";
        leftPanel.alignment = ["fill", "fill"];

        // 右カラム
        var rightPanel = dialog.add("group");
        rightPanel.orientation = "column";
        rightPanel.alignChildren = "right";
        rightPanel.alignment = ["right", "fill"];

        configureDialog(dialog, {
            opacity: 0.97
        });

        /* UI生成部分を以下に維持 */

        /* 移動距離パネル / Distance panel */
        var unitLabel = "pt";
        var panelDistance = leftPanel.add("panel", undefined, LABELS.distance[lang] + " (" + unitLabel + ")");
        panelDistance.orientation = "column";
        panelDistance.alignChildren = ["center", "center"];
        panelDistance.alignment = ["fill", "center"];
        panelDistance.margins = [15, 20, 15, 10];


        var groupContainer = panelDistance.add("group");
        groupContainer.orientation = "row";
        groupContainer.alignChildren = "top";

        // 左カラム
        var leftColumn = groupContainer.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "left";

        var groupX = leftColumn.add("group");
        groupX.orientation = "row";
        groupX.alignChildren = "center";
        var chkDistanceX = groupX.add("checkbox", undefined, LABELS.horizontal[lang]);
        chkDistanceX.value = true;
        var inputDistanceX = groupX.add("edittext", undefined, "0");
        inputDistanceX.characters = 4;
        inputDistanceX.enabled = chkDistanceX.value;

        var groupY = leftColumn.add("group");
        groupY.orientation = "row";
        groupY.alignChildren = "center";
        var chkDistanceY = groupY.add("checkbox", undefined, LABELS.vertical[lang]);
        chkDistanceY.value = true;
        var inputDistanceY = groupY.add("edittext", undefined, "0");
        inputDistanceY.characters = 4;
        inputDistanceY.enabled = chkDistanceY.value;

        // 右カラム
        var rightColumn = groupContainer.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "center";
        rightColumn.alignment = ["fill", "center"];

        var chkLinkX = rightColumn.add("checkbox", undefined, LABELS.link[lang]);
        chkLinkX.value = true;
        // Disable vertical checkbox and input so they're dimmed when dialog opens
        chkDistanceY.value = false;
        chkDistanceY.enabled = false;
        inputDistanceY.enabled = false;
        inputDistanceY.text = inputDistanceX.text;

        var groupCenter = panelDistance.add("group");
        groupCenter.orientation = "row";
        groupCenter.alignChildren = ["left", "center"];
        groupCenter.alignment = ["left", "center"];
        groupCenter.margins = [0, 10, 0, 0];
        var chkCenter = groupCenter.add("checkbox", undefined, LABELS.center[lang]);
        chkCenter.value = false;
        chkCenter.onClick = function() {
            if (chkCenter.value) {
                chkAvoidOverlap.value = false;
                chkAvoidOverlap.enabled = false;
            } else {
                chkAvoidOverlap.enabled = true;
            }
        };
        // [強制]ボタン追加
        var btnForceCenter = groupCenter.add("button", undefined, "強制");
        btnForceCenter.preferredSize = [40, 26];

        // 「重なりを避ける」チェックボックスを追加（常時表示）
        var groupAvoidOverlap = panelDistance.add("group");
        groupAvoidOverlap.orientation = "row";
        groupAvoidOverlap.alignChildren = ["left", "center"];
        groupAvoidOverlap.alignment = ["left", "center"];
        groupAvoidOverlap.margins = [0, 5, 0, 0];
        var chkAvoidOverlap = groupAvoidOverlap.add("checkbox", undefined, LABELS.avoidOverlap[lang]);
        chkAvoidOverlap.value = true; // デフォルトON

        btnForceCenter.onClick = function() {
            try {
                chkCenter.value = false;

                // 全オブジェクトのバウンディングボックスを取得
                var totalLeft = Infinity,
                    totalTop = -Infinity;
                var totalRight = -Infinity,
                    totalBottom = Infinity;
                for (var i = 0; i < originalStates.length; i++) {
                    var bounds = originalStates[i].item.visibleBounds;
                    if (bounds[0] < totalLeft) totalLeft = bounds[0];
                    if (bounds[1] > totalTop) totalTop = bounds[1];
                    if (bounds[2] > totalRight) totalRight = bounds[2];
                    if (bounds[3] < totalBottom) totalBottom = bounds[3];
                }

                var centerX = (totalLeft + totalRight) / 2;
                var centerY = (totalTop + totalBottom) / 2;

                // 各オブジェクトを中央に配置
                for (var i = 0; i < originalStates.length; i++) {
                    var st = originalStates[i];
                    var itemBounds = st.item.visibleBounds;
                    var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
                    var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;

                    var shiftX = centerX - itemCenterX;
                    var shiftY = centerY - itemCenterY;
                    st.item.position = [st.item.position[0] + shiftX, st.item.position[1] + shiftY];
                }

                // 中央に集めた位置を新たな基準として originalStates を更新
                for (var j = 0; j < originalStates.length; j++) {
                    var st = originalStates[j];
                    st.position = [st.item.position[0], st.item.position[1]];
                    var boundsAfter = st.item.visibleBounds;
                    st.width = boundsAfter[2] - boundsAfter[0];
                    st.height = boundsAfter[1] - boundsAfter[3];
                }

                // 横・縦の入力値を0にリセット
                inputDistanceX.text = "0";
                inputDistanceY.text = "0";

                app.redraw();
            } catch (e) {
                alert("強制処理中にエラーが発生しました: " + e.message);
            }
        };




        // チェックボックスと入力欄を紐付け、プレビューも反映する関数 / Link checkbox & input with preview
        function bindInput(chk, input, previewFunc, linkInput) {
            chk.onClick = function() {
                input.enabled = chk.value;
                if (previewFunc) previewFunc();
            };
            input.onChanging = function() {
                updateLinkedInputAndPreview(input, linkInput, previewFunc);
            };
            changeValueByArrowKey(input, function() {
                updateLinkedInputAndPreview(input, linkInput, previewFunc);
            });
        }

        // 「連動」チェックボックス（横）のON/OFFで縦の入力欄をディム、値をコピー
        chkLinkX.onClick = function() {
            if (chkLinkX.value) {
                chkDistanceY.value = false;
                chkDistanceY.enabled = false;
                inputDistanceY.enabled = false;
                inputDistanceY.text = inputDistanceX.text;
            } else {
                chkDistanceY.value = true;
                chkDistanceY.enabled = true;
                inputDistanceY.enabled = true;
            }
        };

        bindInput(chkDistanceX, inputDistanceX, previewDistance, inputDistanceY);
        bindInput(chkDistanceY, inputDistanceY, previewDistance);

        /* 共通ラベル幅 / Common label width */
        var LABEL_WIDTH = 72;

        // チェックボックス＋入力欄の生成を簡素化する関数 / Simplify creation of checkbox + input
        function createCheckInput(parent, label, width, unit, defaultVal, isChecked) {
            var group = parent.add("group");
            group.orientation = "row";
            group.alignChildren = "center";

            var chk = group.add("checkbox", undefined, label);
            chk.value = isChecked;
            chk.preferredSize.width = width;

            group.add("statictext", undefined, "");

            var input = group.add("edittext", undefined, defaultVal);
            input.characters = 4;
            input.enabled = isChecked;

            if (unit) group.add("statictext", undefined, unit);

            chk.onClick = function() {
                input.enabled = chk.value;
            };

            return {
                chk: chk,
                input: input
            };
        }

        /* 変形・回転・不透明度をまとめたグループ / Group: Scale, Rotate, Opacity */
        // var groupTransform = dialog.add("group");
        var groupTransform = leftPanel.add("panel", undefined, undefined);

        groupTransform.orientation = "column";
        groupTransform.alignChildren = "left";
        groupTransform.margins = [15, 15, 15, 10];




        /* 変形 / Scale */
        var scaleUI = createCheckInput(groupTransform, LABELS.scale[lang], LABEL_WIDTH, "%", "0", true);
        var chkScale = scaleUI.chk;
        var inputScale = scaleUI.input;

        /* 回転 / Rotate */
        var rotateUI = createCheckInput(groupTransform, LABELS.rotate[lang], LABEL_WIDTH, "°", "0", true);
        var chkRotate = rotateUI.chk;
        var inputRotate = rotateUI.input;

        /* 不透明度 / Opacity */
        var opacityUI = createCheckInput(groupTransform, LABELS.opacity[lang], LABEL_WIDTH, "%", "0", true);
        var chkOpacity = opacityUI.chk;
        var inputOpacity = opacityUI.input;

        /* 選択オブジェクトの初期状態を記録（ダイアログ開始時点） */
        var initialStates = [];
        for (var i = 0; i < doc.selection.length; i++) {
            var item = doc.selection[i];
            var bounds = item.visibleBounds;
            initialStates.push({
                item: item,
                position: [item.position[0], item.position[1]],
                width: bounds[2] - bounds[0],
                height: bounds[1] - bounds[3],
                matrix: item.matrix,
                opacity: (item.opacity !== undefined) ? item.opacity : 100
            });
        }
        // originalStatesはプレビュー等で編集可。initialStatesはダイアログ開始時点を保持
        var originalStates = [];
        for (var i = 0; i < initialStates.length; i++) {
            // ディープコピー（item参照は同じだが値は複製）
            var st = initialStates[i];
            originalStates.push({
                item: st.item,
                position: [st.position[0], st.position[1]],
                width: st.width,
                height: st.height,
                matrix: st.matrix,
                opacity: st.opacity
            });
        }

        /* ポジションをリセットする共通関数 / Common function to reset positions */
        function resetPositions(states) {
            for (var i = 0; i < states.length; i++) {
                var st = states[i];
                st.item.position = [st.position[0], st.position[1]];
            }
        }

        /* 移動距離のプレビュー処理（リセット→再適用） / Preview distance (reset → reapply) */
        function previewDistance() {
            if (!chkDistanceX.value && !chkDistanceY.value && !chkLinkX.value && !chkCenter.value) return;
            var valX = chkDistanceX.value ? parseFloat(inputDistanceX.text) : 0;
            var valY;
            if (chkLinkX.value) {
                valY = parseFloat(inputDistanceX.text);
            } else {
                valY = chkDistanceY.value ? parseFloat(inputDistanceY.text) : 0;
            }
            if ((chkDistanceX.value && isNaN(valX)) || (chkDistanceY.value && isNaN(valY))) return;
            try {
                var totalLeft = Infinity,
                    totalTop = -Infinity;
                var totalRight = -Infinity,
                    totalBottom = Infinity;
                if (chkCenter.value) {
                    for (var i = 0; i < originalStates.length; i++) {
                        var bounds = originalStates[i].item.visibleBounds;
                        if (bounds[0] < totalLeft) totalLeft = bounds[0];
                        if (bounds[1] > totalTop) totalTop = bounds[1];
                        if (bounds[2] > totalRight) totalRight = bounds[2];
                        if (bounds[3] < totalBottom) totalBottom = bounds[3];
                    }
                }

                var centerX = 0,
                    centerY = 0;
                if (chkCenter.value) {
                    centerX = (totalLeft + totalRight) / 2;
                    centerY = (totalTop + totalBottom) / 2;
                }

                // なるべく重ならないように配置
                var placedItems = []; // 配置済みアイテムのバウンディングボックス

                for (var i = 0; i < originalStates.length; i++) {
                    var st = originalStates[i];
                    // Reset position
                    st.item.position = [st.position[0], st.position[1]];
                    var newX = st.position[0];
                    var newY = st.position[1];

                    if (chkCenter.value) {
                        var itemBounds = st.item.visibleBounds;
                        var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
                        var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;

                        if (itemCenterX > centerX) {
                            var randX = -(valX * 0.8 + Math.random() * (valX * 0.2));
                            newX = st.position[0] + randX;
                        } else if (itemCenterX < centerX) {
                            var randX = (valX * 0.8 + Math.random() * (valX * 0.2));
                            newX = st.position[0] + randX;
                        }

                        if (itemCenterY > centerY) {
                            var randY = -(valY * 0.8 + Math.random() * (valY * 0.2));
                            newY = st.position[1] + randY;
                        } else if (itemCenterY < centerY) {
                            var randY = (valY * 0.8 + Math.random() * (valY * 0.2));
                            newY = st.position[1] + randY;
                        }
                    } else {
                        if (chkDistanceX.value) {
                            var randX = (Math.random() * 2 - 1) * valX;
                            newX = st.position[0] + randX;
                        }
                        if (chkLinkX.value || chkDistanceY.value) {
                            var randY = (Math.random() * 2 - 1) * valY;
                            newY = st.position[1] + randY;
                        }
                    }

                    var attempt = 0;
                    var placed = false;
                    if (chkAvoidOverlap.value) {
                        while (attempt < 20) {
                            st.item.position = [newX, newY];
                            var vb = st.item.visibleBounds;
                            var overlap = false;

                            for (var j = 0; j < placedItems.length; j++) {
                                var vb2 = placedItems[j];
                                if (!(vb[2] < vb2[0] || vb[0] > vb2[2] || vb[1] < vb2[3] || vb[3] > vb2[1])) {
                                    overlap = true;
                                    break;
                                }
                            }

                            if (!overlap) {
                                placedItems.push(vb);
                                placed = true;
                                break;
                            } else {
                                newX = st.position[0] + (Math.random() * 2 - 1) * valX;
                                newY = st.position[1] + (Math.random() * 2 - 1) * valY;
                                attempt++;
                            }
                        }
                    }

                    if (!chkAvoidOverlap.value || !placed) {
                        st.item.position = [newX, newY]; // 通常配置または妥協配置
                        placedItems.push(st.item.visibleBounds);
                    }
                }
                app.redraw();
            } catch (e) {}
        }

        /* 変形率のプレビュー処理（インクリメンタル適用） / Preview scale (incremental) */
        function previewScale() {
            if (!chkScale.value) return;
            var val = parseFloat(inputScale.text);
            if (isNaN(val)) return;
            try {
                for (var i = 0; i < originalStates.length; i++) {
                    var st = originalStates[i];
                    // 現在の状態に加えてランダム拡大縮小
                    var factor = 1 + (Math.random() * 2 - 1) * (val / 100);
                    var randScale = 100 * factor;
                    if (randScale < 1) randScale = 1;
                    st.item.resize(randScale, randScale);
                }
                app.redraw();
            } catch (e) {}
        }

        function previewOpacity() {
            if (!chkOpacity.value) return;
            var val = parseFloat(inputOpacity.text);
            if (isNaN(val)) return;
            if (val < 0) val = 0;
            if (val > 100) val = 100;
            try {
                for (var i = 0; i < originalStates.length; i++) {
                    var st = originalStates[i];
                    if (st.item.opacity !== undefined) {
                        // 現在の不透明度に加えてランダム変化
                        var baseOpacity = st.item.opacity;
                        var minOpacity = baseOpacity - val;
                        var maxOpacity = baseOpacity + val;
                        if (minOpacity < 0) minOpacity = 0;
                        if (maxOpacity > 100) maxOpacity = 100;
                        var newOpacity = minOpacity + Math.random() * (maxOpacity - minOpacity);
                        st.item.opacity = newOpacity;
                    }
                }
                app.redraw();
            } catch (e) {}
        }

        // 回転リセットフラグ / Flag to ensure rotation reset only once
        var rotationResetDone = false;
        /* 回転プレビュー処理 / Preview rotation */
        function previewRotate() {
            if (!chkRotate.value) return;
            var val = parseFloat(inputRotate.text);
            if (isNaN(val)) return;
            try {
                for (var i = 0; i < originalStates.length; i++) {
                    var st = originalStates[i];
                    // Reset rotation only once when dialog opens
                    if (!rotationResetDone && st.matrix) {
                        st.item.matrix = st.matrix;
                    }
                    // -val ～ +val の範囲でランダムに回転
                    var randRotate = (Math.random() * 2 - 1) * val;
                    st.item.rotate(randRotate);
                }
                rotationResetDone = true;
                app.redraw();
            } catch (e) {}
        }

        /* プレビュー関数をまとめる配列 / Array of preview functions */
        var previewFuncs = [{
                chk: chkDistanceX,
                func: previewDistance
            },
            {
                chk: chkScale,
                func: previewScale
            },
            {
                chk: chkOpacity,
                func: previewOpacity
            },
            {
                chk: chkRotate,
                func: previewRotate
            }
        ];

        // 一括プレビュー実行関数 / Run all previews
        function runAllPreviews() {
            for (var i = 0; i < previewFuncs.length; i++) {
                if (previewFuncs[i].chk && previewFuncs[i].chk.value && previewFuncs[i].func) {
                    previewFuncs[i].func();
                }
            }
        }

        bindInput(chkScale, inputScale, runAllPreviews);
        bindInput(chkRotate, inputRotate, runAllPreviews);
        bindInput(chkOpacity, inputOpacity, runAllPreviews);

        /* ボタン群 / Buttons */
        var btnGroup = rightPanel.add("group");
        btnGroup.alignment = "top";
        btnGroup.orientation = "column";

        // OK button
        var okBtn = btnGroup.add("button", undefined, LABELS.ok[lang], {
            name: "ok"
        });

        // Cancel button
        var cancelBtn = btnGroup.add("button", undefined, LABELS.cancel[lang], {
            name: "cancel"
        });

        // Spacer with larger height
        var spacer = btnGroup.add("statictext", undefined, "");
        spacer.preferredSize.height = 100; // increase height as needed

        // Reset button
        var resetBtn = btnGroup.add("button", undefined, LABELS.reset[lang]);

        // Random button
        var randomBtn = btnGroup.add("button", undefined, LABELS.random[lang]);



        randomBtn.onClick = function() {
            try {
                // 全オブジェクトのバウンディングボックスから幅を計算
                var totalLeft = Infinity,
                    totalTop = -Infinity;
                var totalRight = -Infinity,
                    totalBottom = Infinity;
                for (var i = 0; i < originalStates.length; i++) {
                    var bounds = originalStates[i].item.visibleBounds;
                    if (bounds[0] < totalLeft) totalLeft = bounds[0];
                    if (bounds[1] > totalTop) totalTop = bounds[1];
                    if (bounds[2] > totalRight) totalRight = bounds[2];
                    if (bounds[3] < totalBottom) totalBottom = bounds[3];
                }
                var totalWidth = totalRight - totalLeft;
                var halfWidth = Math.round(totalWidth / 2);

                // 各フィールドにランダム推奨値をセット
                inputDistanceX.text = halfWidth.toString();
                inputDistanceY.text = halfWidth.toString();
                inputScale.text = "20";
                inputRotate.text = "10";
                inputOpacity.text = "10";

                chkDistanceX.value = true;
                chkDistanceY.value = true;
                chkLinkX.value = true; // ランダム時は連動をON
                chkScale.value = true;
                chkRotate.value = true;
                chkOpacity.value = true;
                chkCenter.value = false; // ランダム時は中央揃えをOFF

                // 連動ON時は縦入力をディムして値を同期
                chkDistanceY.value = false;
                chkDistanceY.enabled = false;
                inputDistanceY.enabled = false;
                inputDistanceY.text = inputDistanceX.text;

                // プレビューを実行
                runAllPreviews();
            } catch (e) {
                alert("ランダム処理中にエラーが発生しました: " + e.message);
            }
        };

        // 共通のボタン幅を設定 / Set common button width
        var BUTTON_WIDTH = 100;
        okBtn.preferredSize.width = BUTTON_WIDTH;
        cancelBtn.preferredSize.width = BUTTON_WIDTH;
        resetBtn.preferredSize.width = BUTTON_WIDTH;
        randomBtn.preferredSize.width = BUTTON_WIDTH;

        // Reset action: restore state from before dialog and set rotation to 0°
        resetBtn.onClick = function() {
            try {
                // 元の状態に復元
                restoreOriginalStates(originalStates);

                // 長方形を水平に
                for (var i = 0; i < originalStates.length; i++) {
                    try {
                        makeRectangleEdgeHorizontal(originalStates[i].item);
                    } catch (e2) {}
                }

                // 入力フィールドを0に戻す / Reset all input fields to 0
                inputDistanceX.text = "0";
                inputDistanceY.text = "0";
                inputScale.text = "0";
                inputRotate.text = "0";
                inputOpacity.text = "0";

                chkDistanceX.value = true;
                chkDistanceY.value = true;
                chkLinkX.value = true;
                chkCenter.value = false;
                chkScale.value = true;
                chkRotate.value = true;
                chkOpacity.value = true;

                app.redraw();
            } catch (e) {
                alert("リセット処理中にエラーが発生しました: " + e.message);
            }
        };

        /* 元の状態に戻す関数 / Restore original states */
        function restoreOriginalStates(states) {
            resetPositions(states);
            for (var i = 0; i < states.length; i++) {
                var st = states[i];
                if (st.matrix) {
                    st.item.matrix = st.matrix;
                }
                if (st.opacity !== undefined) {
                    st.item.opacity = st.opacity;
                }
            }
            app.redraw();
        }

        // Store bounds on move
        dialog.onMove = function() {
            storeWinBounds(dialog);
        };

        cancelBtn.onClick = function() {
            try {
                restoreOriginalStates(initialStates);
            } catch (e) {}
            try {
                app.executeMenuCommand('Live Corner Annotator');
            } catch (e2) {}
            storeWinBounds(dialog);
            dialog.close();
        };

        okBtn.onClick = function() {
            // プレビューで反映済みなので一括実行で最終反映
            // runAllPreviews();
            try {
                app.executeMenuCommand('Live Corner Annotator');
            } catch (e2) {}
            storeWinBounds(dialog);
            dialog.close();
        };

        dialog.show();

    } catch (e) {
        alert("エラーが発生しました：" + e.message + " / An error occurred: " + e.message);
    }
}

main();