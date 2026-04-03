#target illustrator
/*
 * 連番振り直し / Smart Renumber
 * 
 * [概要 / Summary]
 * 選択した数値テキストを「現在の値」「垂直座標」「水平座標」のいずれかでソートし、連番を振り直します。
 * Sorts selected numeric text frames by "Current Value", "Vertical", or "Horizontal" position and renumbers them.
 * 
 * [ショートカットキー / Shortcut Keys]
 * - D: 現在の数値順に設定 / Set to Current Value Order
 * - V: 垂直方向（上から下）に設定 / Set to Vertical Order
 * - H: 水平方向（左から右）に設定 / Set to Horizontal Order
 * - Z: ゼロ埋めの切り替え / Toggle Zero Padding
 * - Up/Down Arrows: 開始番号の増減 (±1) / Adjust Start Number (±1)
 * - Shift + Up/Down: 開始番号の増減 (±10) / Adjust Start Number (±10)
 * - Alt + Up/Down: 開始番号の増減 (±0.1) / Adjust Start Number (±0.1)
 */

app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

var SCRIPT_VERSION = "v1.1";

/**
 * 現在の言語を取得 / Get current language
 * @returns {string} "ja" or "en"
 */
function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/**
 * 言語ラベル定義 / Language label definitions
 */
var LABELS = {
    dialogTitle: {
        ja: "連番振り直し",
        en: "Smart Renumber"
    },
    startNum: {
        ja: "開始番号:",
        en: "Start Number:"
    },
    sortOrder: {
        ja: "並び順",
        en: "Sort Order"
    },
    currentVal: {
        ja: "現在の数値順",
        en: "Current Value Order"
    },
    vertical: {
        ja: "垂直方向（上から下）",
        en: "Vertical (Top to Bottom)"
    },
    horizontal: {
        ja: "水平方向（左から右）",
        en: "Horizontal (Left to Right)"
    },
    zeroPad: {
        ja: "ゼロ埋め",
        en: "Zero Padding"
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
        en: "No document open."
    },
    errNoSelection: {
        ja: "テキストオブジェクトを選択してください。",
        en: "Please select text objects."
    },
    errNoNumeric: {
        ja: "数字が入力されたテキストオブジェクトが見つかりませんでした。",
        en: "No numeric text objects found."
    }
};

/**
 * ローカライズ文字列の取得 / Get localized string
 * @param {string} key 
 * @returns {string}
 */
function L(key) {
    return LABELS[key][lang] || LABELS[key]["en"];
}

/* メイン処理 / Main process */
main();

function main() {
    if (app.documents.length === 0) {
        alert(L('errNoDoc'));
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert(L('errNoSelection'));
        return;
    }

    var textObjects = [];

    /* オブジェクト情報の取得 / Collect object information */
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "TextFrame") {
            var content = sel[i].contents;
            if (!isNaN(parseFloat(content)) && isFinite(content)) {
                textObjects.push({
                    obj: sel[i],
                    value: parseFloat(content),
                    x: sel[i].left,
                    y: sel[i].top,
                    original: content
                });
            }
        }
    }

    if (textObjects.length === 0) {
        alert(L('errNoNumeric'));
        return;
    }

    /* ダイアログボックスの作成 / Create dialog box */
    var win = new Window("dialog", L('dialogTitle') + ' ' + SCRIPT_VERSION);
    win.orientation = "row";
    win.alignChildren = ["left", "top"];
    win.spacing = 20;
    win.margins = 20;

    /* 左カラム：設定 / Left column: Settings */
    var leftCol = win.add("group");
    leftCol.orientation = "column";
    leftCol.alignChildren = ["left", "top"];
    leftCol.spacing = 15;

    /* 開始番号入力 / Start number input */
    var group1 = leftCol.add("group");
    group1.alignChildren = ["left", "center"];
    group1.add("statictext", undefined, L('startNum'));
    var inputNumber = group1.add("edittext", undefined, "");

    // 初期値を最小値に設定 / Set initial value to minimum found
    var initialMin = textObjects.slice().sort(function (a, b) { return a.value - b.value; })[0].value;
    inputNumber.text = initialMin;
    inputNumber.characters = 8;
    inputNumber.active = true;

    /* 並び順パネル / Sort order panel */
    var sortPanel = leftCol.add("panel", undefined, L('sortOrder'));
    sortPanel.orientation = "column";
    sortPanel.alignChildren = ["left", "top"];
    sortPanel.margins = [15, 20, 15, 10];
    sortPanel.spacing = 8;

    var rbIgnore = sortPanel.add("radiobutton", undefined, L('currentVal'));
    var rbVertical = sortPanel.add("radiobutton", undefined, L('vertical'));
    var rbHorizontal = sortPanel.add("radiobutton", undefined, L('horizontal'));
    rbIgnore.value = true;

    /* オプション / Options */
    var chkZeroPad = leftCol.add("checkbox", undefined, L('zeroPad'));

    /* 右カラム：ボタン / Right column: Buttons */
    var rightCol = win.add("group");
    rightCol.orientation = "column";
    rightCol.alignChildren = ["fill", "top"];
    rightCol.preferredSize.width = 100;
    rightCol.spacing = 10;

    var btnOK = rightCol.add("button", undefined, L('ok'), { name: "ok" });
    var btnCancel = rightCol.add("button", undefined, L('cancel'), { name: "cancel" });

    /* プレビュー更新関数 / Preview update function */
    var updatePreview = function () {
        var startNum = parseFloat(inputNumber.text);
        if (isNaN(startNum)) return;

        var sortedList = textObjects.slice();

        if (rbIgnore.value) {
            sortedList.sort(function (a, b) { return a.value - b.value; });
        } else if (rbVertical.value) {
            sortedList.sort(function (a, b) { return b.y - a.y; });
        } else if (rbHorizontal.value) {
            sortedList.sort(function (a, b) { return a.x - b.x; });
        }

        var maxNum = startNum + sortedList.length - 1;
        var maxDigits = Math.floor(Math.abs(maxNum)).toString().length;

        for (var j = 0; j < sortedList.length; j++) {
            var newNum = startNum + j;
            var newStr = chkZeroPad.value ? zeroPad(newNum, maxDigits) : newNum.toString();
            sortedList[j].obj.contents = newStr;
        }
        app.redraw();
    };

    /* キー入力ハンドラ / Keyboard input handler */
    win.addEventListener("keydown", function (event) {
        var key = event.keyName;
        var handled = false;

        switch (key) {
            case "D":
                rbIgnore.value = true;
                handled = true;
                break;
            case "V":
                rbVertical.value = true;
                handled = true;
                break;
            case "H":
                rbHorizontal.value = true;
                handled = true;
                break;
            case "Z":
                chkZeroPad.value = !chkZeroPad.value;
                handled = true;
                break;
        }

        if (handled) {
            updatePreview();
            event.preventDefault();
        }
    });

    /* イベントリスナー / Event listeners */
    changeValueByArrowKey(inputNumber, updatePreview);
    inputNumber.onChanging = updatePreview;
    rbIgnore.onClick = updatePreview;
    rbVertical.onClick = updatePreview;
    rbHorizontal.onClick = updatePreview;
    chkZeroPad.onClick = updatePreview;

    updatePreview();

    if (win.show() === 1) {
        /* 確定 / Confirm */
    } else {
        /* キャンセル：復元 / Cancel: Restore original */
        for (var k = 0; k < textObjects.length; k++) {
            textObjects[k].obj.contents = textObjects[k].original;
        }
        app.redraw();
    }
}

/**
 * 矢印キーでの数値増減 / Incremental value change with arrow keys
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
 * ゼロ埋め処理 / Zero padding process
 */
function zeroPad(num, digits) {
    var isNegative = num < 0;
    var absNum = Math.abs(num);
    var str = absNum.toString();
    while (str.length < digits) { str = "0" + str; }
    return (isNegative ? "-" : "") + str;
}
