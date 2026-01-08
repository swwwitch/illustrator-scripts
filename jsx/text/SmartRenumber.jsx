#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
 * 連番振り直し / Smart Renumber
 * 
 * [概要 / Summary]
 * 更新日: 2026-01-09
 * 選択した数値テキストを、指定した並び順でソートして連番を振り直します。
 * 接頭辞・接尾辞の追加、ゼロ埋め、逆順などのオプションが利用可能です。
 * 
 * プレビューは app.undo() を使って「巻き戻し→再適用」するため、操作中のUndo履歴を汚しません。
 * OK確定時は一度すべて巻き戻してから本番を1回だけ適用し、Undo 1回で取り消せます。
 * 
 * ダイアログは見やすさのため、表示位置を右へオフセットし、透明度をわずかに下げています。
 * 
 * Sorts selected numeric text frames by the chosen order and renumbers them.
 * Supports prefix/suffix, zero padding, and reverse order.
 * 
 * Preview rolls back via app.undo() (rollback → reapply), so it doesn't pollute the Undo history.
 * On OK, rolls back and applies once so the whole operation is undoable in a single step.
 * 
 * The dialog is shifted to the right and slightly transparent for better visibility.
 */

var SCRIPT_VERSION = "v2.0";

// ダイアログ表示位置・透明度
var DIALOG_OFFSET_X = 300;
var DIALOG_OFFSET_Y = 0;
var DIALOG_OPACITY = 0.98;

function shiftDialogPosition(dlg, offsetX, offsetY) {
    var prevOnShow = dlg.onShow;
    dlg.onShow = function () {
        try { if (prevOnShow) prevOnShow(); } catch (e) {}
        try {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        } catch (e2) {}
    };
}

function setDialogOpacity(dlg, opacityValue) {
    try { dlg.opacity = opacityValue; } catch (e) {}
}

/**
 * 現在の言語を取得
 */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/**
 * 言語ラベル定義
 */
var LABELS = {
    dialogTitle: { ja: "連番振り直し", en: "Smart Renumber" },
    startNum: { ja: "開始番号:", en: "Start Number:" },
    sortOrder: { ja: "並び順", en: "Sort Order" },
    textAdd: { ja: "テキスト追加", en: "Text Addition" },
    prefix: { ja: "接頭辞", en: "Prefix" },
    suffix: { ja: "接尾辞", en: "Suffix" },
    currentVal: { ja: "現在の数値順", en: "Current Value Order" },
    vertical: { ja: "垂直方向（上から下）", en: "Vertical (Top to Bottom)" },
    horizontal: { ja: "水平方向（左から右）", en: "Horizontal (Left to Right)" },
    zPattern: { ja: "Z方向（左→右、上→下）", en: "Z-Pattern (Left-to-Right, Row-major)" },
    nPattern: { ja: "N方向（上→下、左→右）", en: "N-Pattern (Top-to-Bottom, Column-major)" },
    stackTop: { ja: "重ね順（前面から）", en: "Layer Order (Front to Back)" },
    reverse: { ja: "逆順", en: "Reverse" },
    zeroPad: { ja: "ゼロ埋め", en: "Zero Padding" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    errNoDoc: { ja: "ドキュメントが開かれていません。", en: "No document open." },
    errNoSelection: { ja: "テキストオブジェクトを選択してください。", en: "Please select text objects." },
    errNoNumeric: { ja: "数字が入力されたテキストオブジェクトが見つかりませんでした。", en: "No numeric text objects found." }
};

function L(key) { return LABELS[key][lang] || LABELS[key]["en"]; }

main();

function main() {
    if (app.documents.length === 0) { alert(L('errNoDoc')); return; }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) { alert(L('errNoSelection')); return; }

    var textObjects = [];

    /* オブジェクト情報の取得 */
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "TextFrame") {
            var content = sel[i].contents;
            if (!isNaN(parseFloat(content)) && isFinite(content)) {
                textObjects.push({
                    obj: sel[i],
                    value: parseFloat(content),
                    x: sel[i].left,
                    y: sel[i].top,
                    stackOrder: i,
                    original: content
                });
            }
        }
    }

    if (textObjects.length === 0) { alert(L('errNoNumeric')); return; }

    // プレビュー時のUndo管理
    var previewMgr = new PreviewManager();

    /* ダイアログボックスの作成 */
    var win = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
    win.orientation = "row";
    win.alignChildren = ["left", "top"];
    win.spacing = 20;
    win.margins = 20;

    // ダイアログの透明度と位置を調整
    setDialogOpacity(win, DIALOG_OPACITY);
    shiftDialogPosition(win, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

    /* 左カラム */
    var leftCol = win.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = ["left", "top"];
    leftCol.spacing = 15;

    // 開始番号
    var groupStart = leftCol.add("group");
    groupStart.add("statictext", undefined, L('startNum'));
    var inputNumber = groupStart.add("edittext", undefined, "");
    var initialMin = textObjects.slice().sort(function (a, b) { return a.value - b.value; })[0].value;
    inputNumber.text = initialMin;
    inputNumber.characters = 8;
    inputNumber.active = true;

    // 並び順パネル
    var sortPanel = leftCol.add("panel", undefined, L('sortOrder'));
    sortPanel.orientation = "column";
    sortPanel.alignChildren = ["left", "top"];
    sortPanel.margins = [15, 20, 15, 10];
    sortPanel.spacing = 4;

    var rbIgnore = sortPanel.add("radiobutton", undefined, L('currentVal'));
    var rbVertical = sortPanel.add("radiobutton", undefined, L('vertical'));
    var rbHorizontal = sortPanel.add("radiobutton", undefined, L('horizontal'));
    var rbZ = sortPanel.add("radiobutton", undefined, L('zPattern'));
    var rbN = sortPanel.add("radiobutton", undefined, L('nPattern'));
    var rbStackTop = sortPanel.add("radiobutton", undefined, L('stackTop'));
    rbIgnore.value = true;

    // オプション（横並び）
    var optionsGroup = leftCol.add("group");
    optionsGroup.orientation = "row";
    optionsGroup.spacing = 20;
    var chkReverse = optionsGroup.add("checkbox", undefined, L('reverse'));
    var chkZeroPad = optionsGroup.add("checkbox", undefined, L('zeroPad'));


    // テキスト追加パネル
    var textAddPanel = leftCol.add("panel", undefined, L('textAdd'));
    textAddPanel.orientation = "column";
    textAddPanel.alignChildren = ["right", "top"];
    textAddPanel.margins = [15, 20, 15, 15];
    textAddPanel.spacing = 8;

    var groupPrefix = textAddPanel.add("group");
    groupPrefix.add("statictext", undefined, L('prefix'));
    var inputPrefix = groupPrefix.add("edittext", undefined, "");
    inputPrefix.characters = 12;

    var groupSuffix = textAddPanel.add("group");
    groupSuffix.add("statictext", undefined, L('suffix'));
    var inputSuffix = groupSuffix.add("edittext", undefined, "");
    inputSuffix.characters = 12;


    /* 右カラム */
    var rightCol = win.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];
    rightCol.preferredSize.width = 100;
    rightCol.spacing = 10;

    var btnOK = rightCol.add("button", undefined, L('ok'), { name: "ok" });
    var btnCancel = rightCol.add("button", undefined, L('cancel'), { name: "cancel" });

    btnCancel.onClick = function () {
        previewMgr.rollback();
        win.close(0);
    };

    btnOK.onClick = function () {
        // 一度プレビューを全てUndoしてから、最後に1回だけ本番適用（Undo 1回で戻せる）
        previewMgr.confirm(function () {
            var startNum = parseFloat(inputNumber.text);
            if (isNaN(startNum)) return;

            var prefix = inputPrefix.text;
            var suffix = inputSuffix.text;

            var sortedList = textObjects.slice();
            var threshold = 10;

            sortedList.sort(function (a, b) {
                if (rbIgnore.value) return a.value - b.value;
                if (rbVertical.value) return b.y - a.y;
                if (rbHorizontal.value) return a.x - b.x;
                if (rbStackTop.value) return a.stackOrder - b.stackOrder;

                if (rbZ.value) {
                    if (Math.abs(a.y - b.y) > threshold) return b.y - a.y;
                    return a.x - b.x;
                }
                if (rbN.value) {
                    if (Math.abs(a.x - b.x) > threshold) return a.x - b.x;
                    return b.y - a.y;
                }
                return 0;
            });

            if (chkReverse.value) sortedList.reverse();

            var maxNum = startNum + sortedList.length - 1;
            var maxDigits = Math.floor(Math.abs(maxNum)).toString().length;

            for (var j = 0; j < sortedList.length; j++) {
                var newNum = startNum + j;
                var numStr = chkZeroPad.value ? zeroPad(newNum, maxDigits) : newNum.toString();
                sortedList[j].obj.contents = prefix + numStr + suffix;
            }
            app.redraw();
        });
        win.close(1);
    };

    // ダイアログを×で閉じた場合もロールバック
    win.onClose = function () {
        previewMgr.rollback();
        return true;
    };

    /* プレビュー更新（Undo履歴を汚さない） */
    var updatePreview = function () {
        // 直前のプレビューを必ず巻き戻す
        previewMgr.rollback();

        previewMgr.addStep(function () {
            var startNum = parseFloat(inputNumber.text);
            if (isNaN(startNum)) return;

            var prefix = inputPrefix.text;
            var suffix = inputSuffix.text;

            var sortedList = textObjects.slice();
            var threshold = 10;

            sortedList.sort(function (a, b) {
                if (rbIgnore.value) return a.value - b.value;
                if (rbVertical.value) return b.y - a.y;
                if (rbHorizontal.value) return a.x - b.x;
                if (rbStackTop.value) return a.stackOrder - b.stackOrder;

                if (rbZ.value) {
                    if (Math.abs(a.y - b.y) > threshold) return b.y - a.y;
                    return a.x - b.x;
                }
                if (rbN.value) {
                    if (Math.abs(a.x - b.x) > threshold) return a.x - b.x;
                    return b.y - a.y;
                }
                return 0;
            });

            if (chkReverse.value) sortedList.reverse();

            var maxNum = startNum + sortedList.length - 1;
            var maxDigits = Math.floor(Math.abs(maxNum)).toString().length;

            for (var j = 0; j < sortedList.length; j++) {
                var newNum = startNum + j;
                var numStr = chkZeroPad.value ? zeroPad(newNum, maxDigits) : newNum.toString();
                sortedList[j].obj.contents = prefix + numStr + suffix;
            }
        });
    };

    /* イベントリスナー */
    changeValueByArrowKey(inputNumber, updatePreview);
    inputNumber.onChanging = updatePreview;
    inputPrefix.onChanging = updatePreview;
    inputSuffix.onChanging = updatePreview;
    rbIgnore.onClick = rbVertical.onClick = rbHorizontal.onClick =
        rbZ.onClick = rbN.onClick = rbStackTop.onClick =
        chkReverse.onClick = chkZeroPad.onClick = updatePreview;

    updatePreview();

    win.show();
}

/**
 * プレビュー時の履歴管理と一括Undoを制御するクラス
 * - updatePreview() のたびに rollback() → addStep() で「履歴を汚さないプレビュー」を実現
 * - OK確定時は confirm(finalAction) で「1回のUndo」で戻せるようにする
 */
function PreviewManager() {
    this.undoDepth = 0;

    this.addStep = function (func) {
        try {
            func();
            this.undoDepth++;
            app.redraw();
        } catch (e) {
            alert("Preview Error: " + e);
        }
    };

    this.rollback = function () {
        while (this.undoDepth > 0) {
            try {
                app.undo();
            } catch (e) {
                // 何らかの理由でUndoできない場合は中断
                break;
            }
            this.undoDepth--;
        }
        app.redraw();
    };

    this.confirm = function (finalAction) {
        if (finalAction) {
            this.rollback();
            finalAction();
        } else {
            this.undoDepth = 0;
        }
    };
}

/**
 * 矢印キーでの数値増減
 */
function changeValueByArrowKey(editText, callback) {
    editText.addEventListener("keydown", function (event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;
        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;
        if (keyboard.shiftKey) {
            delta = 10;
            if (event.keyName == "Up") value = Math.ceil((value + 0.1) / delta) * delta;
            else if (event.keyName == "Down") value = Math.floor((value - 0.1) / delta) * delta;
            event.preventDefault();
        } else if (keyboard.altKey) {
            delta = 0.1;
            if (event.keyName == "Up") value += delta;
            else if (event.keyName == "Down") value -= delta;
            event.preventDefault();
        } else {
            if (event.keyName == "Up") value += 1;
            else if (event.keyName == "Down") value -= 1;
            if (event.keyName == "Up" || event.keyName == "Down") event.preventDefault();
        }
        value = keyboard.altKey ? Math.round(value * 10) / 10 : Math.round(value);
        editText.text = value;
        if (callback) callback();
    });
}

/**
 * ゼロ埋め処理
 */
function zeroPad(num, digits) {
    var isNegative = num < 0;
    var absNum = Math.abs(num);
    var str = absNum.toString();
    while (str.length < digits) { str = "0" + str; }
    return (isNegative ? "-" : "") + str;
}
