#target illustrator
#targetengine "MyScriptEngine"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

RandomizeObjects.jsx

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/RandomizeObjects.jsx

### 更新日：

- 2026-02-27（完全シャッフル：ランダム色生成／カラーは［ランダム］で実行）

### 概要：

- 選択したオブジェクトをランダムに移動・変形・回転・不透明度を変更するスクリプト
- UIから各種パラメータを指定し、即時プレビューで結果を確認可能
- カラーの通常シャッフル／完全シャッフルに対応
- [リセット]でダイアログ起動前の状態に復元

### 主な機能：

- 移動距離を横・縦方向に個別設定、または連動
- 中央揃えオプションでオブジェクトを一箇所に集約
- 拡大縮小（変形）、回転、不透明度のランダム化
- プレビュー反映、キャンセルで元の状態に完全リセット
- UI状態の同期安定性向上（チェックON/OFFと入力欄enabledの整合性を統一）
- プレビューを毎回ベース状態から再計算（積み増し挙動を解消）

### 処理の流れ：

1. ドキュメントと選択状態を確認
2. ダイアログを表示（移動距離・変形率・回転・不透明度を入力）
3. 入力値をもとにランダム変形や移動を即時プレビュー
4. OKで確定、キャンセルでダイアログ起動前にリセット
*/

var SCRIPT_VERSION = "v2.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */
var LABELS = {
    dialogTitle: { ja: "ランダム化 " + SCRIPT_VERSION, en: "Randomize " + SCRIPT_VERSION },
    distance: { ja: "移動距離", en: "Distance" },
    horizontal: { ja: "横:", en: "Horizontal:" },
    vertical: { ja: "縦:", en: "Vertical:" },
    link: { ja: "連動", en: "Link" },
    gatherCenter: { ja: "中央に集める", en: "Gather to Center" },
    avoidOverlap: { ja: "重なりを避ける", en: "Avoid Overlap" },
    panelScaleTitle: { ja: "スケール（%）", en: "Scale (%)" },
    scale: { ja: "幅・高さ", en: "Width & Height" },
    scaleX: { ja: "幅", en: "Width" },
    scaleY: { ja: "高さ", en: "Height" },
    rotate: { ja: "回転", en: "Rotate" },
    opacity: { ja: "不透明度", en: "Opacity" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    reset: { ja: "リセット", en: "Reset" },
    random: { ja: "ランダム", en: "Random" },
    shuffle: { ja: "シャッフル", en: "Shuffle" },
    force: { ja: "強制", en: "Force" }
    , shuffleMode: { ja: "シャッフル方式", en: "Shuffle Mode" }
    , fullShuffle: { ja: "完全シャッフル", en: "Full Shuffle" }
    , apply: { ja: "実行", en: "Apply" }
    , none: { ja: "なし", en: "None" }
};

/* 入力値を更新し、リンクされた入力とプレビューを処理 / Update input, linked input, and preview */
function updateLinkedInputAndPreview(input, linkInput, previewFunc) {
    if (linkInput && linkInput.enabled) linkInput.text = input.text;
    if (previewFunc) previewFunc();
}

/* ダイアログ設定を共通化 / Configure dialog position and opacity */
function configureDialog(dlg, options) {
    if (options.opacity !== undefined) dlg.opacity = options.opacity;
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
        if (event.keyName == "Up") return Math.ceil((value + 1) / delta) * delta;
        if (event.keyName == "Down") return Math.floor((value - 1) / delta) * delta;
    } else if (keyboard.altKey) {
        delta = 0.1;
        if (event.keyName == "Up") return value + delta;
        if (event.keyName == "Down") return Math.max(0, value - delta);
    } else {
        if (event.keyName == "Up") return value + 1;
        if (event.keyName == "Down") return Math.max(0, value - 1);
    }
    return value;
}

/* 矢印キーで数値を増減 / Change numeric value with arrow keys */
function changeValueByArrowKey(editText, previewFunc) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var newValue = getNewValueByKey(event, value, keyboard);

        if (newValue !== value) {
            event.preventDefault();
            if (keyboard.altKey) newValue = Math.round(newValue * 10) / 10;
            else newValue = Math.round(newValue);
            editText.text = newValue;
            if (previewFunc) updateLinkedInputAndPreview(editText, null, previewFunc);
        }
    });
}

// ダイアログ位置の記憶（Illustrator終了で忘れる）/ Remember dialog location (cleared when Illustrator quits)
var savedLoc = $.global.__RandomizeObjectsDialogLoc || null;

function storeWinLoc(dlg) {
    try {
        // location is [left, top]
        $.global.__RandomizeObjectsDialogLoc = [dlg.location[0], dlg.location[1]];
    } catch (e) { }
}

function restoreWinLoc(dlg) {
    try {
        if (savedLoc && savedLoc instanceof Array && savedLoc.length === 2) {
            dlg.location = [savedLoc[0], savedLoc[1]];
        }
    } catch (e) { }
}

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

        var dialog = new Window("dialog", LABELS.dialogTitle[lang]);

        try { app.executeMenuCommand('Live Corner Annotator'); } catch (e) { }

        dialog.orientation = "column";
        dialog.alignChildren = "fill";
        dialog.spacing = 20;
        dialog.margins = [15, 15, 20, 15];

        // Content row (main left/right panels)
        var contentRow = dialog.add("group");
        contentRow.orientation = "row";
        contentRow.alignChildren = "top";
        contentRow.alignment = ["fill", "fill"];
        contentRow.spacing = 20;

        var leftPanel = contentRow.add("group");
        leftPanel.orientation = "column";
        leftPanel.alignChildren = "left";
        leftPanel.alignment = ["fill", "top"];

        var rightPanel = contentRow.add("group");
        rightPanel.orientation = "column";
        rightPanel.alignChildren = "left";
        rightPanel.alignment = ["right", "top"];

        configureDialog(dialog, { opacity: 0.97 });
        restoreWinLoc(dialog);

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

        var leftColumn = groupContainer.add("group");
        leftColumn.orientation = "column";
        leftColumn.alignChildren = "left";

        var groupX = leftColumn.add("group");
        groupX.orientation = "row";
        groupX.alignChildren = "center";
        var chkDistanceX = groupX.add("checkbox", undefined, LABELS.horizontal[lang]);
        chkDistanceX.value = false;
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

        var rightColumn = groupContainer.add("group");
        rightColumn.orientation = "column";
        rightColumn.alignChildren = "center";
        rightColumn.alignment = ["fill", "center"];

        var chkLinkX = rightColumn.add("checkbox", undefined, LABELS.link[lang]);
        chkLinkX.value = true;

        chkDistanceY.value = false;
        chkDistanceY.enabled = false;
        inputDistanceY.enabled = false;
        inputDistanceY.text = inputDistanceX.text;

        var groupCenter = panelDistance.add("group");
        groupCenter.orientation = "row";
        groupCenter.alignChildren = ["center", "center"];
        groupCenter.alignment = ["fill", "center"];
        groupCenter.margins = [0, 10, 0, 0];

        var chkCenter = { value: false }; // dummy
        var btnForceCenter = groupCenter.add("button", undefined, LABELS.gatherCenter[lang]);
        btnForceCenter.alignment = ["center", "center"];
        btnForceCenter.preferredSize = [120, 24];

        var chkAvoidOverlap = { value: false, enabled: true }; // dummy

        var groupAvoidOverlap = panelDistance.add("group");
        groupAvoidOverlap.orientation = "row";
        groupAvoidOverlap.alignChildren = ["center", "center"];
        groupAvoidOverlap.alignment = ["fill", "center"];
        groupAvoidOverlap.margins = [0, 2, 0, 0];

        var btnForceAvoid = groupAvoidOverlap.add("button", undefined, LABELS.avoidOverlap[lang]);
        btnForceAvoid.alignment = ["center", "center"];
        btnForceAvoid.preferredSize = [120, 24];

        // カラーパネル / Color panel
        var panelColor = leftPanel.add("panel", undefined, "カラー");
        panelColor.orientation = "column";
        panelColor.alignChildren = ["fill", "top"]; // カラムいっぱいに広げる
        panelColor.alignment = ["fill", "top"];      // 左カラム幅に追従
        panelColor.margins = [15, 20, 15, 10];

        // シャッフル方式 / Shuffle mode
        var modeGroup = panelColor.add("group");
        modeGroup.orientation = "column";
        modeGroup.alignChildren = ["fill", "center"]; // ラジオも横幅いっぱい
        modeGroup.alignment = ["fill", "top"];

        var rbNoneShuffle = modeGroup.add("radiobutton", undefined, LABELS.none[lang]);
        var rbShuffle = modeGroup.add("radiobutton", undefined, LABELS.shuffle[lang]);
        var rbFullShuffle = modeGroup.add("radiobutton", undefined, LABELS.fullShuffle[lang]);
        rbNoneShuffle.value = true; // default
        rbNoneShuffle.alignment = ["fill", "center"];
        rbShuffle.alignment = ["fill", "center"];
        rbFullShuffle.alignment = ["fill", "center"];

        // var btnApplyColorShuffle = panelColor.add("button", undefined, LABELS.apply[lang]);
        // オブジェクトの塗り色をシャッフル / Shuffle fill colors of selected objects
        function cloneColor(c) {
            if (!c) return null;
            try {
                var t = c.typename;
                if (t === "RGBColor") {
                    var nc = new RGBColor();
                    nc.red = c.red; nc.green = c.green; nc.blue = c.blue;
                    return nc;
                }
                if (t === "CMYKColor") {
                    var nc2 = new CMYKColor();
                    nc2.cyan = c.cyan; nc2.magenta = c.magenta; nc2.yellow = c.yellow; nc2.black = c.black;
                    return nc2;
                }
                if (t === "GrayColor") {
                    var nc3 = new GrayColor();
                    nc3.gray = c.gray;
                    return nc3;
                }
                // SpotColor / GradientColor / PatternColor etc.
                // These are typically safe to reuse as references.
                return c;
            } catch (e) {
                return c;
            }
        }

        function shuffleArray(array) {
            var arr = array.slice();
            for (var i = arr.length - 1; i > 0; i--) {
                var j = Math.floor(Math.random() * (i + 1));
                var temp = arr[i];
                arr[i] = arr[j];
                arr[j] = temp;
            }
            return arr;
        }

        // 塗り色を持つ対象を返す（PathItem / CompoundPathItem / TextFrame / GroupItem対応）
        function getFillTarget(item) {
            if (!item) return null;
            try {
                var t = item.typename;
                if (t === "PathItem") return item;
                if (t === "CompoundPathItem") {
                    if (item.pathItems && item.pathItems.length > 0) return item.pathItems[0];
                    return null;
                }
                if (t === "TextFrame") {
                    return item; // textRange.characterAttributes.fillColor
                }
                if (t === "GroupItem") {
                    if (!item.pageItems) return null;
                    for (var i = 0; i < item.pageItems.length; i++) {
                        var target = getFillTarget(item.pageItems[i]);
                        if (target) return target;
                    }
                }
                return null;
            } catch (e) {
                return null;
            }
        }

        function getFillColorFromTarget(target) {
            if (!target) return null;
            try {
                if (target.typename === "TextFrame") {
                    return target.textRange.characterAttributes.fillColor;
                }
                if (target.filled) return target.fillColor;
                return null;
            } catch (e) {
                return null;
            }
        }

        function setFillColorToTarget(target, color) {
            if (!target || !color) return false;
            try {
                if (target.typename === "TextFrame") {
                    target.textRange.characterAttributes.fillColor = color;
                    return true;
                }
                if (target.filled) {
                    target.fillColor = color;
                    return true;
                }
                return false;
            } catch (e) {
                return false;
            }
        }

        function shuffleSelectionFillColors() {
            var doc = app.activeDocument;
            var selection = doc.selection;

            if (!selection || selection.length < 2) {
                alert("2つ以上のオブジェクトを選択してください");
                return;
            }

            // Collect fill colors (null for no-fill targets)
            var colors = [];
            var targets = [];
            for (var i = 0; i < selection.length; i++) {
                var target = getFillTarget(selection[i]);
                targets.push(target);
                var c = getFillColorFromTarget(target);
                colors.push(c ? cloneColor(c) : null);
            }

            colors = shuffleArray(colors);

            // Apply shuffled colors
            for (var k = 0; k < targets.length; k++) {
                if (colors[k] != null) {
                    setFillColorToTarget(targets[k], colors[k]);
                }
            }

            app.redraw();
        }

        // 完全シャッフル：CMYK/RGB のランダム色を生成して適用 / Full shuffle: generate random CMYK/RGB colors
        function fullShuffleSelectionFillColors() {
            var doc = app.activeDocument;
            var selection = doc.selection;

            if (!selection || selection.length < 1) {
                alert("オブジェクトを選択してください");
                return;
            }

            function randomInt(min, max) {
                return Math.floor(Math.random() * (max - min + 1)) + min;
            }

            function createRandomDocColor() {
                try {
                    if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
                        var cmyk = new CMYKColor();
                        cmyk.cyan = randomInt(0, 100);
                        cmyk.magenta = randomInt(0, 100);
                        cmyk.yellow = randomInt(0, 100);
                        cmyk.black = randomInt(0, 30);
                        return cmyk;
                    }
                } catch (e) { }

                // default: RGB
                var rgb = new RGBColor();
                rgb.red = randomInt(0, 255);
                rgb.green = randomInt(0, 255);
                rgb.blue = randomInt(0, 255);
                return rgb;
            }

            var appliedCount = 0;

            for (var i = 0; i < selection.length; i++) {
                var target = getFillTarget(selection[i]);
                if (!target) continue;

                var col = createRandomDocColor();

                // Force fill ON for PathItem if needed
                try {
                    if (target.typename === "PathItem") {
                        if (!target.filled) target.filled = true;
                    }
                } catch (e2) { }

                if (setFillColorToTarget(target, col)) {
                    appliedCount++;
                }
            }

            if (appliedCount === 0) {
                alert("塗りカラーを適用できる対象が見つかりませんでした");
                return;
            }

            app.redraw();
        }


        // カラーシャッフルをラジオ選択に応じて実行 / Apply color shuffle based on radio selection
        function applyColorShuffleFromUI() {
            try {
                if (rbNoneShuffle && rbNoneShuffle.value) return;
                if (rbFullShuffle && rbFullShuffle.value) {
                    fullShuffleSelectionFillColors();
                } else {
                    shuffleSelectionFillColors();
                }
            } catch (e) {
                alert("カラーシャッフル中にエラーが発生しました: " + e.message);
            }
        }

        // NOTE: originalStates is defined later but referenced here; it will exist when the button is clicked.

        btnForceAvoid.onClick = function () {
            try {
                var padding = 5;
                var placedItems = [];

                var baseX = parseFloat(inputDistanceX.text || "50");
                var baseY = parseFloat(inputDistanceY.text || "50");
                if (isNaN(baseX) || baseX <= 0) baseX = 50;
                if (isNaN(baseY) || baseY <= 0) baseY = 50;

                var success = false;
                var scaleFactor = 1;

                while (!success && scaleFactor <= 20) {
                    placedItems = [];
                    success = true;

                    for (var i = 0; i < originalStates.length; i++) {
                        var st = originalStates[i];
                        var attempts = 0;
                        var placed = false;

                        while (attempts < 300) {
                            var randX = (Math.random() * 2 - 1) * baseX * scaleFactor;
                            var randY = (Math.random() * 2 - 1) * baseY * scaleFactor;
                            var newX = st.position[0] + randX;
                            var newY = st.position[1] + randY;

                            st.item.position = [newX, newY];
                            var vb = st.item.visibleBounds;
                            var overlap = false;

                            for (var j = 0; j < placedItems.length; j++) {
                                var vb2 = placedItems[j];
                                if (!(vb[2] + padding < vb2[0] || vb[0] - padding > vb2[2] || vb[1] + padding < vb2[3] || vb[3] - padding > vb2[1])) {
                                    overlap = true;
                                    break;
                                }
                            }

                            if (!overlap) {
                                placedItems.push(vb);
                                placed = true;
                                break;
                            }
                            attempts++;
                        }

                        if (!placed) {
                            success = false;
                            break;
                        }
                    }

                    if (!success) {
                        scaleFactor += 1;
                        inputDistanceX.text = String(Math.round(baseX * scaleFactor));
                        inputDistanceY.text = String(Math.round(baseY * scaleFactor));
                    }
                }

                if (!success) alert("十分な距離を確保できず、完全に非重複で配置できませんでした。");
                else app.redraw();
            } catch (e) {
                alert("強制処理中にエラーが発生しました: " + e.message);
            }
        };

        btnForceCenter.onClick = function () {
            try {
                chkCenter.value = false;

                var totalLeft = Infinity, totalTop = -Infinity;
                var totalRight = -Infinity, totalBottom = Infinity;

                for (var i = 0; i < originalStates.length; i++) {
                    var bounds = originalStates[i].item.visibleBounds;
                    if (bounds[0] < totalLeft) totalLeft = bounds[0];
                    if (bounds[1] > totalTop) totalTop = bounds[1];
                    if (bounds[2] > totalRight) totalRight = bounds[2];
                    if (bounds[3] < totalBottom) totalBottom = bounds[3];
                }

                var centerX = (totalLeft + totalRight) / 2;
                var centerY = (totalTop + totalBottom) / 2;

                for (var i2 = 0; i2 < originalStates.length; i2++) {
                    var st2 = originalStates[i2];
                    var itemBounds = st2.item.visibleBounds;
                    var itemCenterX = (itemBounds[0] + itemBounds[2]) / 2;
                    var itemCenterY = (itemBounds[1] + itemBounds[3]) / 2;

                    var shiftX = centerX - itemCenterX;
                    var shiftY = centerY - itemCenterY;
                    st2.item.position = [st2.item.position[0] + shiftX, st2.item.position[1] + shiftY];
                }

                for (var j2 = 0; j2 < originalStates.length; j2++) {
                    var st3 = originalStates[j2];
                    st3.position = [st3.item.position[0], st3.item.position[1]];
                    var boundsAfter = st3.item.visibleBounds;
                    st3.width = boundsAfter[2] - boundsAfter[0];
                    st3.height = boundsAfter[1] - boundsAfter[3];
                }

                inputDistanceX.text = "0";
                inputDistanceY.text = "0";

                app.redraw();
            } catch (e) {
                alert("強制処理中にエラーが発生しました: " + e.message);
            }
        };

        function bindInput(chk, input, previewFunc, linkInput, opts) {
            chk.onClick = function () {
                input.enabled = chk.value;

                // opts: { set20OnEnable: boolean, skipPreviewOnCheck: boolean }
                if (opts && opts.set20OnEnable && chk.value) {
                    input.text = "20";

                    // scale link: when enabling Width while link is ON, keep Height synced
                    if (typeof chkLinkScale !== "undefined" && chkLinkScale && chk === chkScaleX && chkLinkScale.value) {
                        inputScaleY.text = input.text;
                    }

                    // distance link: when enabling Horizontal while Link is ON, keep Vertical synced
                    if (chk === chkDistanceX && chkLinkX && chkLinkX.value && typeof inputDistanceY !== "undefined" && inputDistanceY) {
                        inputDistanceY.text = input.text;
                    }
                }

                if (!(opts && opts.skipPreviewOnCheck)) {
                    if (previewFunc) previewFunc();
                }
            };
            input.onChanging = function () {
                if (chkLinkX.value && linkInput) updateLinkedInputAndPreview(input, linkInput, previewFunc);
                else if (previewFunc) previewFunc();
            };
            changeValueByArrowKey(input, function () {
                if (chkLinkX.value && linkInput) updateLinkedInputAndPreview(input, linkInput, previewFunc);
                else if (previewFunc) previewFunc();
            });
        }

        chkLinkX.onClick = function () {
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

        /* ラベル幅 / Label widths */
        var LABEL_WIDTH_SCALE = 60;   // 幅/高さ用
        var LABEL_WIDTH_OPT = 80;     // 回転/不透明度用

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

            chk.onClick = function () { input.enabled = chk.value; };

            return { chk: chk, input: input };
        }

        // スケールパネル（右カラム） / Scale panel (right column)
        var panelScale = rightPanel.add("panel", undefined, LABELS.panelScaleTitle[lang]);
        panelScale.orientation = "column";
        panelScale.alignChildren = ["fill", "top"];
        panelScale.margins = [15, 20, 15, 10];

        // Distanceと同じ2カラム構成（左: 幅/高さ、右: 連動）
        var scaleContainer = panelScale.add("group");
        scaleContainer.orientation = "row";
        scaleContainer.alignChildren = "top";

        // 左カラム（幅/高さ）
        var scaleLeft = scaleContainer.add("group");
        scaleLeft.orientation = "column";
        scaleLeft.alignChildren = "left";

        var scaleRowX = scaleLeft.add("group");
        scaleRowX.orientation = "row";
        scaleRowX.alignChildren = "center";
        var chkScaleX = scaleRowX.add("checkbox", undefined, LABELS.scaleX[lang]);
        chkScaleX.value = false;
        chkScaleX.preferredSize.width = LABEL_WIDTH_SCALE;
        var inputScaleX = scaleRowX.add("edittext", undefined, "0");
        inputScaleX.characters = 4;
        inputScaleX.enabled = chkScaleX.value;

        var scaleRowY = scaleLeft.add("group");
        scaleRowY.orientation = "row";
        scaleRowY.alignChildren = "center";
        var chkScaleY = scaleRowY.add("checkbox", undefined, LABELS.scaleY[lang]);
        chkScaleY.value = false;
        chkScaleY.preferredSize.width = LABEL_WIDTH_SCALE;
        var inputScaleY = scaleRowY.add("edittext", undefined, "0");
        inputScaleY.characters = 4;
        inputScaleY.enabled = chkScaleY.value;

        // 右カラム（連動）
        var scaleRight = scaleContainer.add("group");
        scaleRight.orientation = "column";
        scaleRight.alignChildren = "center";
        scaleRight.alignment = ["fill", "center"];

        var chkLinkScale = scaleRight.add("checkbox", undefined, LABELS.link[lang]);
        chkLinkScale.value = true;

        // 起動時は連動ON: 高さはディム＆同期
        chkScaleY.value = false;
        chkScaleY.enabled = false;
        inputScaleY.enabled = false;
        inputScaleY.text = inputScaleX.text;

        // 回転・不透明度パネル（右カラム） / Rotate & Opacity panel (right column)
        var groupTransform = rightPanel.add("panel", undefined, "オプション");
        groupTransform.orientation = "column";
        groupTransform.alignChildren = "left";
        groupTransform.margins = [15, 20, 15, 10];

        var rotateUI = createCheckInput(groupTransform, LABELS.rotate[lang], LABEL_WIDTH_OPT, "°", "0", false);
        var chkRotate = rotateUI.chk;
        var inputRotate = rotateUI.input;

        // 回転スライダー（0〜180）
        var rotateSlider = groupTransform.add("slider", undefined, 0, 0, 180);
        rotateSlider.alignment = ["fill", "center"];

        // 入力→スライダー同期
        inputRotate.onChanging = function () {
            var v = parseFloat(inputRotate.text);
            if (!isNaN(v)) {
                if (v < 0) v = 0;
                if (v > 180) v = 180;
                rotateSlider.value = v;
            }
        };

        // スライダー→入力同期
        rotateSlider.onChanging = function () {
            // スライダー操作で回転を自動ON（チェックのonClickは呼ばない＝値上書きを防ぐ）
            chkRotate.value = true;
            inputRotate.enabled = true;

            inputRotate.text = Math.round(rotateSlider.value).toString();
            runAllPreviews();
        };

        var opacityUI = createCheckInput(groupTransform, LABELS.opacity[lang], LABEL_WIDTH_OPT, "%", "0", false);
        var chkOpacity = opacityUI.chk;
        var inputOpacity = opacityUI.input;

        // 不透明度スライダー（0〜100）
        var _syncingOpacitySlider = false;
        var opacitySlider = groupTransform.add("slider", undefined, 0, 0, 100);
        opacitySlider.alignment = ["fill", "center"];
        opacitySlider.enabled = true; // 常に有効

        // スライダー→入力（※edittextのonChangingは自動発火しないため明示的にrunAllPreviews）
        opacitySlider.onChanging = function () {
            if (_syncingOpacitySlider) return;

            // スライダー操作で不透明度を自動ON（チェックのonClickは呼ばない＝値上書きを防ぐ）
            chkOpacity.value = true;
            inputOpacity.enabled = true;

            _syncingOpacitySlider = true;
            inputOpacity.text = Math.round(opacitySlider.value).toString();
            _syncingOpacitySlider = false;

            runAllPreviews();
        };

        // チェックON時に既定値を自動入力 / Auto set default value when enabled
        function setDefault20OnEnable(chk, input) {
            if (!chk || !input) return;
            var prev = chk.onClick;
            chk.onClick = function () {
                if (typeof prev === "function") prev();
                input.enabled = chk.value;
                if (chk.value) {
                    input.text = "20";
                }
                // ON/OFFのみ行い、ここではランダム実行しない
            };
        }


        // スケール連動のUI反映 / Apply scale link UI state
        function applyScaleLinkUI() {
            if (!chkLinkScale) return;

            if (chkLinkScale.value) {
                chkScaleY.value = false;
                chkScaleY.enabled = false;
                inputScaleY.enabled = false;
                inputScaleY.text = inputScaleX.text;
            } else {
                chkScaleY.enabled = true;
                inputScaleY.enabled = chkScaleY.value;
            }
        }
        chkLinkScale.onClick = function () {
            applyScaleLinkUI();
            runAllPreviews();
        };
        applyScaleLinkUI(); // initial

        // UIの状態同期 / Sync UI state
        function syncUIState(options) {
            options = options || {};

            if (options.hasOwnProperty("chkDistanceX")) chkDistanceX.value = !!options.chkDistanceX;
            inputDistanceX.enabled = chkDistanceX.value;

            if (options.hasOwnProperty("chkLinkX")) chkLinkX.value = !!options.chkLinkX;

            if (chkLinkX.value) {
                chkDistanceY.value = false;
                chkDistanceY.enabled = false;
                inputDistanceY.enabled = false;
                inputDistanceY.text = inputDistanceX.text;
            } else {
                if (options.hasOwnProperty("chkDistanceY")) chkDistanceY.value = !!options.chkDistanceY;
                chkDistanceY.enabled = true;
                inputDistanceY.enabled = chkDistanceY.value;
            }

            if (options.hasOwnProperty("chkScaleX")) chkScaleX.value = !!options.chkScaleX;
            inputScaleX.enabled = chkScaleX.value;

            if (options.hasOwnProperty("chkLinkScale")) chkLinkScale.value = !!options.chkLinkScale;

            if (options.hasOwnProperty("chkScaleY")) chkScaleY.value = !!options.chkScaleY;
            inputScaleY.enabled = chkScaleY.value;

            if (options.hasOwnProperty("chkRotate")) chkRotate.value = !!options.chkRotate;
            inputRotate.enabled = chkRotate.value;

            if (options.hasOwnProperty("chkOpacity")) chkOpacity.value = !!options.chkOpacity;
            inputOpacity.enabled = chkOpacity.value;

            // Apply scale link dimming/sync
            if (typeof applyScaleLinkUI === "function") applyScaleLinkUI();
        }

        /* 選択オブジェクトの初期状態を記録（ダイアログ開始時点） */
        var initialStates = [];
        for (var s = 0; s < doc.selection.length; s++) {
            var it = doc.selection[s];
            var b = it.visibleBounds;
            initialStates.push({
                item: it,
                position: [it.position[0], it.position[1]],
                width: b[2] - b[0],
                height: b[1] - b[3],
                matrix: it.matrix,
                opacity: (it.opacity !== undefined) ? it.opacity : 100
            });
        }

        var originalStates = [];
        for (var t = 0; t < initialStates.length; t++) {
            var st0 = initialStates[t];
            originalStates.push({
                item: st0.item,
                position: [st0.position[0], st0.position[1]],
                width: st0.width,
                height: st0.height,
                matrix: st0.matrix,
                opacity: st0.opacity
            });
        }

        function resetPositions(states) {
            for (var r = 0; r < states.length; r++) {
                var stp = states[r];
                stp.item.position = [stp.position[0], stp.position[1]];
            }
        }

        function previewDistance() {
            if (!chkDistanceX.value && !chkDistanceY.value) return;

            var valX = chkDistanceX.value ? parseFloat(inputDistanceX.text) : 0;
            var valY = 0;

            if (chkLinkX.value && chkDistanceX.value) valY = parseFloat(inputDistanceX.text);
            else valY = chkDistanceY.value ? parseFloat(inputDistanceY.text) : 0;

            if ((chkDistanceX.value && isNaN(valX)) || (chkDistanceY.value && isNaN(valY))) return;

            try {
                for (var i3 = 0; i3 < originalStates.length; i3++) {
                    var st = originalStates[i3];
                    var newX = st.position[0];
                    var newY = st.position[1];

                    if (chkDistanceX.value) newX = st.position[0] + (Math.random() * 2 - 1) * valX;
                    if (chkLinkX.value && chkDistanceX.value) newY = st.position[1] + (Math.random() * 2 - 1) * valY;
                    else if (chkDistanceY.value) newY = st.position[1] + (Math.random() * 2 - 1) * valY;

                    st.item.position = [newX, newY];
                }
            } catch (e) { }
        }

        function previewScale() {
            // Link ON かつ 幅ON のときは高さも幅に連動して適用（高さチェックがOFF/ディムでも適用）
            var applyWidth = !!chkScaleX.value;
            var applyHeight = (!!chkScaleY.value) || (chkLinkScale && chkLinkScale.value && applyWidth);
            if (!applyWidth && !applyHeight) return;

            var valX2 = applyWidth ? parseFloat(inputScaleX.text) : 0;
            var valY2 = applyHeight ? parseFloat((chkLinkScale && chkLinkScale.value && applyWidth) ? inputScaleX.text : inputScaleY.text) : 0;

            if (applyWidth && isNaN(valX2)) return;
            if (applyHeight && isNaN(valY2)) return;

            try {
                for (var i4 = 0; i4 < originalStates.length; i4++) {
                    var st4 = originalStates[i4];

                    if (applyWidth) {
                        var fX = 1 + (Math.random() * 2 - 1) * (valX2 / 100);
                        var sX = Math.max(1, 100 * fX);
                        st4.item.resize(sX, 100);
                    }
                    if (applyHeight) {
                        var fY = 1 + (Math.random() * 2 - 1) * (valY2 / 100);
                        var sY = Math.max(1, 100 * fY);
                        st4.item.resize(100, sY);
                    }
                }
            } catch (e) { }
        }

        function previewRotate() {
            if (!chkRotate.value) return;
            var valR = parseFloat(inputRotate.text);
            if (isNaN(valR)) return;

            try {
                for (var i5 = 0; i5 < originalStates.length; i5++) {
                    var st5 = originalStates[i5];
                    st5.item.rotate((Math.random() * 2 - 1) * valR);
                }
            } catch (e) { }
        }

        function previewOpacity() {
            if (!chkOpacity.value) return;

            var valO = parseFloat(inputOpacity.text);
            if (isNaN(valO)) return;
            if (valO < 0) valO = 0;
            if (valO > 100) valO = 100;

            try {
                for (var i6 = 0; i6 < originalStates.length; i6++) {
                    var st6 = originalStates[i6];
                    if (st6.item.opacity !== undefined) {
                        var base = (st6.opacity !== undefined) ? st6.opacity : st6.item.opacity;
                        var minO = Math.max(0, base - valO);
                        var maxO = Math.min(100, base + valO);
                        st6.item.opacity = minO + Math.random() * (maxO - minO);
                    }
                }
            } catch (e) { }
        }

        function restoreOriginalStates(states, doRedraw) {
            if (doRedraw === undefined) doRedraw = true;
            resetPositions(states);
            for (var i7 = 0; i7 < states.length; i7++) {
                var st7 = states[i7];
                if (st7.matrix) st7.item.matrix = st7.matrix;
                if (st7.opacity !== undefined) st7.item.opacity = st7.opacity;
            }
            if (doRedraw) app.redraw();
        }

        function runAllPreviews() {
            restoreOriginalStates(originalStates, false);

            if (chkDistanceX.value || chkDistanceY.value) previewDistance();
            if (chkScaleX.value || chkScaleY.value) previewScale(); // ←追加
            if (chkRotate.value) previewRotate();
            if (chkOpacity.value) previewOpacity();

            app.redraw();
        }

        // Distance: ON/OFFのみ（プレビューは実行しない）
        bindInput(chkDistanceX, inputDistanceX, runAllPreviews, inputDistanceY, { set20OnEnable: true, skipPreviewOnCheck: true });
        bindInput(chkDistanceY, inputDistanceY, runAllPreviews, null, { skipPreviewOnCheck: true });

        // other (Width/Height/Rotate/Opacity: set 20 on enable, do not preview on ON/OFF)
        bindInput(chkScaleX, inputScaleX, runAllPreviews, null, { set20OnEnable: true, skipPreviewOnCheck: true });
        bindInput(chkScaleY, inputScaleY, runAllPreviews, null, { set20OnEnable: true, skipPreviewOnCheck: true });
        bindInput(chkRotate, inputRotate, runAllPreviews, null, { set20OnEnable: true, skipPreviewOnCheck: true });
        bindInput(chkOpacity, inputOpacity, runAllPreviews, null, { set20OnEnable: true, skipPreviewOnCheck: true });

        // 不透明度：入力→スライダー同期（bindInput が onChanging を上書きするため後付けでラップ）
        var _prevOpacityOnChanging = inputOpacity.onChanging;
        inputOpacity.onChanging = function () {
            if (!_syncingOpacitySlider) {
                var v = parseFloat(inputOpacity.text);
                if (!isNaN(v)) {
                    if (v < 0) v = 0;
                    if (v > 100) v = 100;
                    _syncingOpacitySlider = true;
                    opacitySlider.value = v;
                    _syncingOpacitySlider = false;
                }
            }
            if (typeof _prevOpacityOnChanging === "function") {
                _prevOpacityOnChanging();
            }
        };

        // 初期状態（スライダーは常時有効）
        opacitySlider.value = parseFloat(inputOpacity.text) || 0;

        // 幅のみの入力が変わったら、連動ON時は高さにも反映（bindInput が onChanging を上書きするため後付けでラップ）
        var _prevScaleXOnChanging2 = inputScaleX.onChanging;
        inputScaleX.onChanging = function () {
            if (chkLinkScale && chkLinkScale.value) {
                inputScaleY.text = inputScaleX.text;
            }
            if (typeof _prevScaleXOnChanging2 === "function") _prevScaleXOnChanging2();
        };

        /* Buttons (bottom row: 3 columns) */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.alignment = ["fill", "bottom"];
        btnRowGroup.margins = [0, 10, 0, 0];

        // Left column: Random / Reset
        var btnLeft = btnRowGroup.add("group");
        btnLeft.orientation = "row";
        btnLeft.alignChildren = ["left", "center"];
        btnLeft.alignment = ["left", "center"];

        var randomBtn = btnLeft.add("button", undefined, LABELS.random[lang]);
        var resetBtn = btnLeft.add("button", undefined, LABELS.reset[lang]);

        // Center column: spacer
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        // Right column: Cancel / OK
        var btnRight = btnRowGroup.add("group");
        btnRight.orientation = "row";
        btnRight.alignChildren = ["right", "center"];
        btnRight.alignment = ["right", "center"];

        var cancelBtn = btnRight.add("button", undefined, LABELS.cancel[lang], { name: "cancel" });
        var okBtn = btnRight.add("button", undefined, LABELS.ok[lang], { name: "ok" });


        // Move/copy handlers from original buttons to new ones
        randomBtn.onClick = function () {
            try {
                syncUIState({
                    chkDistanceX: chkDistanceX.value,
                    chkLinkX: chkLinkX.value,
                    chkDistanceY: chkDistanceY.value,
                    chkScaleX: chkScaleX.value,
                    chkScaleY: chkScaleY.value,
                    chkLinkScale: chkLinkScale.value,
                    chkRotate: chkRotate.value,
                    chkOpacity: chkOpacity.value
                });
                // カラー（なし/シャッフル/完全シャッフル）を反映
                applyColorShuffleFromUI();
                runAllPreviews();
            } catch (e) {
                alert("ランダム処理中にエラーが発生しました: " + e.message);
            }
        };

        resetBtn.onClick = function () {
            try {
                restoreOriginalStates(initialStates);

                for (var i8 = 0; i8 < initialStates.length; i8++) {
                    try { makeRectangleEdgeHorizontal(initialStates[i8].item); } catch (e2) { }
                }

                inputDistanceX.text = "0";
                inputDistanceY.text = "0";
                inputScaleX.text = "0";
                inputScaleY.text = "0";
                inputRotate.text = "0";
                inputOpacity.text = "0";

                syncUIState({
                    chkDistanceX: true,
                    chkLinkX: true,
                    chkLinkScale: true,
                    chkScaleX: false,
                    chkScaleY: false,
                    chkRotate: false,
                    chkOpacity: false
                });

                app.redraw();
            } catch (e) {
                alert("リセット処理中にエラーが発生しました: " + e.message);
            }
        };

        dialog.onMove = function () { storeWinLoc(dialog); };

        // Closeボタン（×）などで閉じた場合も位置を記録 / Save bounds on close as well
        dialog.onClose = function () {
            try { storeWinLoc(dialog); } catch (e) { }
            return true;
        };

        cancelBtn.onClick = function () {
            try { restoreOriginalStates(initialStates); } catch (e) { }
            try { app.executeMenuCommand('Live Corner Annotator'); } catch (e2) { }
            storeWinLoc(dialog);
            dialog.close();
        };

        okBtn.onClick = function () {
            try { app.executeMenuCommand('Live Corner Annotator'); } catch (e2) { }
            storeWinLoc(dialog);
            dialog.close();
        };

        dialog.show();

    } catch (e) {
        alert("エラーが発生しました：" + e.message + " / An error occurred: " + e.message);
    }
}

main();