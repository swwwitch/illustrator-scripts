#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したリンク画像を、指定した列数・行数のグリッドに分割します。
分割後の各画像はクリッピングマスクで切り出され、元の配置を保ったまま並びます。

詳細は README を参照してください。

### Overview

Splits the selected linked image into a grid of the given number of columns and rows.
Each piece is cut out with a clipping mask and keeps the original placement.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "SplitLinkedImage";             /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-09-19";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-19";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialogTitle: {
            ja: "リンク画像を分割",
            en: "Split Linked Image"
        },
        countLR: {
            ja: "左右",
            en: "Columns"
        },
        countTB: {
            ja: "上下",
            en: "Rows"
        },
        overlap: {
            ja: "オーバーラップ",
            en: "Overlap"
        },
        groupCheck: {
            ja: "グループ化",
            en: "Create Group"
        },
        tipColumns: { ja: "横に何枚へ分割するかです。", en: "How many pieces to cut the image into across." },
        tipRows: { ja: "縦に何枚へ分割するかです。", en: "How many pieces to cut the image into down." },
        tipOverlap: { ja: "隣り合う分割片を重ねる幅です。継ぎ目を目立たせたくないときに使います。", en: "How far neighbouring pieces overlap. Use it to hide the seams." },
        tipGroupCheck: { ja: "分割してできた画像を1つのグループにまとめます。", en: "Groups the resulting pieces into a single group." },
        tipRuleCheck: { ja: "分割した境目にケイ線を引きます。", en: "Draws a rule along each cut." },
        tipRoundCheck: { ja: "分割片の角を丸めます。半径は右の欄で指定します。", en: "Rounds the corners of each piece. The field on the right sets the radius." },
        tipRoundRadius: { ja: "角丸の半径です。", en: "Radius of the rounded corners." },
        optionsPanel: {
            ja: "オプション",
            en: "Options"
        },
        splitPanel: {
            ja: "分割数",
            en: "Split Count"
        },
        ruleCheck: {
            ja: "ケイ",
            en: "Add Stroke"
        },
        roundCheck: {
            ja: "角丸",
            en: "Apply Round Corners"
        },
        ok: {
            ja: "OK",
            en: "OK"
        },
        cancel: {
            ja: "キャンセル",
            en: "Cancel"
        },
        errNoDoc: {
            ja: "ドキュメントが開かれていません。",
            en: "No document is open."
        },
        errNoSelection: {
            ja: "リンク画像を1つ以上選択してください。",
            en: "Please select one or more linked images."
        },
        errLR: {
            ja: "左右の分割数は1以上の整数を入力してください。",
            en: "Columns must be an integer of 1 or more."
        },
        errTB: {
            ja: "上下の分割数は1以上の整数を入力してください。",
            en: "Rows must be an integer of 1 or more."
        },
        errBoth: {
            ja: "左右または上下のいずれかを2以上にしてください。",
            en: "Either columns or rows must be 2 or more."
        },
        errOverlap: {
            ja: "オーバーラップは0以上の数値を入力してください。",
            en: "Overlap must be a number of 0 or more."
        },
        processed: {
            ja: "分割したリンク画像",
            en: "Processed"
        },
        skipped: {
            ja: "スキップ",
            en: "Skipped"
        },
        unit: {
            ja: "個",
            en: " item(s)"
        }
    };

    function getLabel(key) {
        return LABELS[key][uiLang];
    }

    function labelText(key) {
        return getLabel(key) + (uiLang === 'ja' ? '：' : ':');
    }

    // =========================================
    // 単位 / Units
    // =========================================

    /* 単位コードに対応する表示ラベルと、1単位あたりのポイント数
       Unit code -> display label and points per unit */
    var UNITS = [
        { label: "in",    pointsPerUnit: 72 },                /* 0 */
        { label: "mm",    pointsPerUnit: 72 / 25.4 },         /* 1 */
        { label: "pt",    pointsPerUnit: 1 },                 /* 2 */
        { label: "pica",  pointsPerUnit: 12 },                /* 3 */
        { label: "cm",    pointsPerUnit: 72 / 2.54 },         /* 4 */
        { label: "Q",     pointsPerUnit: 72 / 25.4 * 0.25 },  /* 5 */
        { label: "px",    pointsPerUnit: 1 },                 /* 6 */
        { label: "ft/in", pointsPerUnit: 72 * 12 },           /* 7 */
        { label: "m",     pointsPerUnit: 72 / 25.4 * 1000 },  /* 8 */
        { label: "yd",    pointsPerUnit: 72 * 36 },           /* 9 */
        { label: "ft",    pointsPerUnit: 72 * 12 }            /* 10 */
    ];

    /**
     * 環境設定キーの単位を返す
     * @param {string} [prefKey] - "rulerType"（既定）/ "strokeUnits" / "text/units" / "text/asianunits"
     * @returns {{code: number, label: string, pointsPerUnit: number}} 単位の情報
     */
    function getUnitInfo(prefKey) {
        var unitCode = app.preferences.getIntegerPreference(prefKey || "rulerType");
        /* 未知のコードは pt に寄せる / unknown codes fall back to points */
        var unit = UNITS[unitCode] || UNITS[2];
        return { code: unitCode, label: unit.label, pointsPerUnit: unit.pointsPerUnit };
    }

    // =========================================
    // 矢印キーによる値変更 / Arrow Key Value Change
    // =========================================
    function changeValueByArrowKey(editText) {
        editText.addEventListener("keydown", function (event) {
            var value = Number(editText.text);
            if (isNaN(value)) return;

            var keyboard = ScriptUI.environment.keyboardState;
            var delta = 1;

            if (keyboard.shiftKey) {
                delta = 10;
                /* Shiftキー押下時は10の倍数にスナップ / Snap to multiples of 10 with Shift */
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
                /* Optionキー押下時は0.1単位で増減 / Increment by 0.1 with Option */
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
                /* 小数第1位までに丸め / Round to 1 decimal */
                value = Math.round(value * 10) / 10;
            } else {
                /* 整数に丸め / Round to integer */
                value = Math.round(value);
            }

            editText.text = value;
        });
    }

    // =========================================
    // 事前チェック / Pre-check
    // =========================================
    if (app.documents.length === 0) {
        alert(getLabel('errNoDoc'));
        return;
    }

    var doc = app.activeDocument;

    if (doc.selection.length === 0) {
        alert(getLabel('errNoSelection'));
        return;
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================
    var dlg = new Window("dialog", getLabel('dialogTitle') + ' ' + SCRIPT_VERSION);
    dlg.orientation = "column";
    dlg.alignChildren = "left";

    var splitPanel = dlg.add("panel", undefined, getLabel('splitPanel'));
    splitPanel.orientation = "column";
    splitPanel.alignChildren = "left";
    splitPanel.margins = [15, 20, 15, 10];

    var countLRGroup = splitPanel.add("group");
    countLRGroup.orientation = "row";
    countLRGroup.add("statictext", undefined, labelText('countLR'));
    var columnsInput = countLRGroup.add("edittext", undefined, "2");
    columnsInput.helpTip = getLabel('tipColumns');
    columnsInput.characters = 4;
    changeValueByArrowKey(columnsInput);

    var countTBGroup = splitPanel.add("group");
    countTBGroup.orientation = "row";
    countTBGroup.add("statictext", undefined, labelText('countTB'));
    var rowsInput = countTBGroup.add("edittext", undefined, "1");
    rowsInput.helpTip = getLabel('tipRows');
    rowsInput.characters = 4;
    changeValueByArrowKey(rowsInput);

    var optionsPanel = dlg.add("panel", undefined, getLabel('optionsPanel'));
    optionsPanel.orientation = "column";
    optionsPanel.alignChildren = "left";
    optionsPanel.margins = [15, 20, 15, 10];

    var overlapGroup = optionsPanel.add("group");
    overlapGroup.orientation = "row";
    overlapGroup.add("statictext", undefined, labelText('overlap'));
    var overlapInput = overlapGroup.add("edittext", undefined, "0");
    overlapInput.helpTip = getLabel('tipOverlap');
    overlapInput.characters = 4;
    changeValueByArrowKey(overlapInput);
    overlapGroup.add("statictext", undefined, getUnitInfo().label);

    var groupCheck = optionsPanel.add("checkbox", undefined, getLabel('groupCheck'));
    groupCheck.helpTip = getLabel('tipGroupCheck');
    groupCheck.value = false;

    var ruleCheck = optionsPanel.add("checkbox", undefined, getLabel('ruleCheck'));
    ruleCheck.helpTip = getLabel('tipRuleCheck');
    ruleCheck.value = false;

    var roundGroup = optionsPanel.add("group");
    roundGroup.orientation = "row";
    var roundCheck = roundGroup.add("checkbox", undefined, getLabel('roundCheck'));
    roundCheck.helpTip = getLabel('tipRoundCheck');
    roundCheck.value = false;
    var roundRadiusInput = roundGroup.add("edittext", undefined, "3");
    roundRadiusInput.helpTip = getLabel('tipRoundRadius');
    roundRadiusInput.characters = 5;
    changeValueByArrowKey(roundRadiusInput);
    roundGroup.add("statictext", undefined, getUnitInfo().label);

    var btnGroup = dlg.add("group");
    btnGroup.alignment = "center";
    btnGroup.add("button", undefined, getLabel('cancel'), { name: "cancel" });
    btnGroup.add("button", undefined, getLabel('ok'), { name: "ok" });

    if (dlg.show() != 1) {
        return;
    }

    var columnCount = parseInt(columnsInput.text, 10);
    var rowCount = parseInt(rowsInput.text, 10);
    /* オーバーラップは rulerType 入力、内部では pt 換算 / Overlap input in ruler unit, used in pt */
    var overlapInputValue = parseFloat(overlapInput.text);
    var overlapInPoints = (isNaN(overlapInputValue) ? NaN : overlapInputValue * getUnitInfo().pointsPerUnit);
    var shouldCreateGroup = groupCheck.value;
    var shouldAddStroke = ruleCheck.value;
    var shouldApplyRoundCorners = roundCheck.value;
    var roundRadiusInputValue = parseFloat(roundRadiusInput.text);
    /* rulerType → pt 変換 / Convert ruler unit to pt */
    var roundRadiusInPoints = (isNaN(roundRadiusInputValue) ? 0 : roundRadiusInputValue) * getUnitInfo().pointsPerUnit;

    if (isNaN(columnCount) || columnCount < 1) {
        alert(getLabel('errLR'));
        return;
    }

    if (isNaN(rowCount) || rowCount < 1) {
        alert(getLabel('errTB'));
        return;
    }

    if (columnCount < 2 && rowCount < 2) {
        alert(getLabel('errBoth'));
        return;
    }

    if (isNaN(overlapInPoints) || overlapInPoints < 0) {
        alert(getLabel('errOverlap'));
        return;
    }

    // =========================================
    // メイン処理 / Main Process
    // =========================================
    /* 選択項目を退避 / Save selection */
    var selectedItems = [];
    for (var selectedIndex = 0; selectedIndex < doc.selection.length; selectedIndex++) {
        selectedItems.push(doc.selection[selectedIndex]);
    }

    var processedCount = 0;
    var skippedCount = 0;
    /* 生成アイテムを保持 / Keep created items for reselection */
    var createdItems = [];

    for (var selectedIndex = 0; selectedIndex < selectedItems.length; selectedIndex++) {
        var selectedItem = selectedItems[selectedIndex];

        if (selectedItem.typename !== "PlacedItem") {
            skippedCount++;
            continue;
        }

        try {
            var left = selectedItem.left;
            var top = selectedItem.top;
            var width = selectedItem.width;
            var height = selectedItem.height;

            var partWidth = width / columnCount;
            var partHeight = height / rowCount;

            /* グループ化ON時は親グループを作成 / Create parent group when Group is ON */
            var parentGroup = shouldCreateGroup ? doc.groupItems.add() : null;

            for (var columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                for (var rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    var clippingGroup = parentGroup
                        ? parentGroup.groupItems.add()
                        : doc.groupItems.add();

                    /* 左右方向の矩形計算 / Horizontal rectangle calculation */
                    var rectLeft = left + (partWidth * columnIndex) - overlapInPoints / 2;
                    var rectWidth = partWidth + overlapInPoints;

                    if (rectLeft < left) {
                        rectWidth -= (left - rectLeft);
                        rectLeft = left;
                    }
                    if (rectLeft + rectWidth > left + width) {
                        rectWidth = (left + width) - rectLeft;
                    }

                    /* 上下方向の矩形計算 / Vertical rectangle calculation */
                    var rectTop = top - (partHeight * rowIndex) + overlapInPoints / 2;
                    var rectHeight = partHeight + overlapInPoints;

                    if (rectTop > top) {
                        rectHeight -= (rectTop - top);
                        rectTop = top;
                    }
                    if (rectTop - rectHeight < top - height) {
                        rectHeight = rectTop - (top - height);
                    }

                    var clippingRect = clippingGroup.pathItems.rectangle(
                        rectTop,
                        rectLeft,
                        rectWidth,
                        rectHeight
                    );
                    clippingRect.stroked = false;
                    clippingRect.filled = false;

                    selectedItem.duplicate(clippingGroup, ElementPlacement.PLACEATEND);

                    clippingRect.zOrder(ZOrderMethod.BRINGTOFRONT);
                    clippingGroup.clipped = true;

                    /* ケイ設定の適用 / Apply rule (stroke + pathfinder) */
                    if (shouldAddStroke) {
                        doc.selection = null;
                        clippingGroup.selected = true;
                        app.executeMenuCommand('Adobe New Stroke Shortcut');
                        app.executeMenuCommand('Live Pathfinder Add');
                    }

                    /* 角丸ライブエフェクトの適用 / Apply round corners live effect */
                    if (shouldApplyRoundCorners && roundRadiusInPoints > 0) {
                        var roundXML = '<LiveEffect name="Adobe Round Corners"><Dict data="R radius ' + roundRadiusInPoints + ' "/></LiveEffect>';
                        clippingGroup.applyEffect(roundXML);
                    }

                    if (!parentGroup) {
                        createdItems.push(clippingGroup);
                    }
                }
            }

            if (parentGroup) {
                createdItems.push(parentGroup);
            }

            selectedItem.remove();
            processedCount++;

        } catch (e) {
            skippedCount++;
        }
    }

    /* 実行後に生成アイテムを選択し直し / Reselect created items after execution */
    doc.selection = null;
    for (var createdIndex = 0; createdIndex < createdItems.length; createdIndex++) {
        createdItems[createdIndex].selected = true;
    }

    /* スキップがある場合のみ通知 / Notify only when some items were skipped */
    if (skippedCount > 0) {
        if (uiLang === 'ja') {
            alert(
                labelText('processed') + processedCount + getLabel('unit') + "\n" +
                labelText('skipped') + skippedCount + getLabel('unit')
            );
        } else {
            alert(
                getLabel('processed') + ': ' + processedCount + getLabel('unit') + "\n" +
                getLabel('skipped') + ': ' + skippedCount + getLabel('unit')
            );
        }
    }
})();