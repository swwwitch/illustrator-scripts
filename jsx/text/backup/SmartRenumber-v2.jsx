#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

main();

function main() {
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;
    var sel = doc.selection;

    if (sel.length === 0) {
        alert("テキストオブジェクトを選択してください。");
        return;
    }

    var textObjects = [];

    // 選択オブジェクトから対象を抽出
    for (var i = 0; i < sel.length; i++) {
        if (sel[i].typename === "TextFrame") {
            var content = sel[i].contents;
            if (!isNaN(parseFloat(content)) && isFinite(content)) {
                textObjects.push({
                    obj: sel[i],
                    value: parseFloat(content),
                    original: content // キャンセル復元用に元の値を保存
                });
            }
        }
    }

    if (textObjects.length === 0) {
        alert("数字が入力されたテキストオブジェクトが見つかりませんでした。");
        return;
    }

    // 数値の小さい順にソート
    textObjects.sort(function (a, b) {
        return a.value - b.value;
    });

    // --- ダイアログボックスの作成 ---
    var win = new Window("dialog", "ナンバリング設定 (プレビュー付)");
    win.orientation = "column";
    win.alignChildren = ["fill", "top"];
    win.spacing = 15;
    win.margins = 20;

    // 入力グループ
    var group1 = win.add("group");
    group1.alignChildren = ["left", "center"];
    group1.add("statictext", undefined, "開始番号:");
    
    var defaultStartValue = textObjects[0].value;
    var inputNumber = group1.add("edittext", undefined, defaultStartValue);
    inputNumber.characters = 10;
    inputNumber.active = true;

    // オプショングループ
    var group2 = win.add("group");
    group2.alignChildren = ["left", "center"];
    var chkZeroPad = group2.add("checkbox", undefined, "ゼロ埋め (最大桁に合わせる)");
    chkZeroPad.value = false; 

    // ボタン（キャンセルを左、OKを右）
    var btnGroup = win.add("group");
    btnGroup.alignment = "right";
    var btnCancel = btnGroup.add("button", undefined, "キャンセル", {name: "cancel"});
    var btnOK = btnGroup.add("button", undefined, "OK", {name: "ok"});

    // --- プレビュー更新関数 ---
    var updatePreview = function() {
        var startNum = parseFloat(inputNumber.text);
        if (isNaN(startNum)) return;

        // ゼロ埋め用の最大桁数計算
        var maxNum = startNum + textObjects.length - 1;
        var maxDigits = Math.floor(Math.abs(maxNum)).toString().length;

        for (var j = 0; j < textObjects.length; j++) {
            var newNum = startNum + j;
            var newStr;

            if (chkZeroPad.value) {
                newStr = zeroPad(newNum, maxDigits);
            } else {
                newStr = newNum.toString();
            }
            textObjects[j].obj.contents = newStr;
        }
        app.redraw(); // 画面を更新
    };

    // --- イベントリスナー ---
    
    // 矢印キー操作の適用
    changeValueByArrowKey(inputNumber, updatePreview);

    // 直接入力時の更新
    inputNumber.onChanging = function() {
        updatePreview();
    };

    // チェックボックス切り替え時の更新
    chkZeroPad.onClick = function() {
        updatePreview();
    };

    // 初回プレビュー実行
    updatePreview();

    // ダイアログ表示
    if (win.show() === 1) {
        // OK時はそのまま確定
    } else {
        // キャンセル時は元の値に戻す
        for (var k = 0; k < textObjects.length; k++) {
            textObjects[k].obj.contents = textObjects[k].original;
        }
        app.redraw();
    }
}

/**
 * 矢印キーでの数値増減の実装
 */
function changeValueByArrowKey(editText, callback) {
    editText.addEventListener("keydown", function(event) {
        var value = Number(editText.text);
        if (isNaN(value)) return;

        var keyboard = ScriptUI.environment.keyboardState;
        var delta = 1;

        if (keyboard.shiftKey) {
            delta = 10;
            if (event.keyName == "Up") {
                value = Math.ceil((value + 0.1) / delta) * delta;
                event.preventDefault();
            } else if (event.keyName == "Down") {
                value = Math.floor((value - 0.1) / delta) * delta;
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
                event.preventDefault();
            }
        }

        // 丸め処理
        if (keyboard.altKey) {
            value = Math.round(value * 10) / 10;
        } else {
            value = Math.round(value);
        }

        editText.text = value;
        
        // プレビュー更新
        if (callback) callback();
    });
}

/**
 * ゼロ埋め用ヘルパー関数
 */
function zeroPad(num, digits) {
    var isNegative = num < 0;
    var absNum = Math.abs(num);
    var str = absNum.toString();
    
    while (str.length < digits) {
        str = "0" + str;
    }
    return (isNegative ? "-" : "") + str;
}
