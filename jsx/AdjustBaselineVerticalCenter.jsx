#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
スクリプト名：AdjustBaselineVerticalCenter.jsx

概要:
選択したテキストフレーム内の指定した文字（1文字以上）を、基準文字に合わせてベースライン（垂直位置）を調整します。
Adjusts the baseline of one or more specified characters in one or multiple text frames to align vertically with a reference character.

処理の流れ:
1. ダイアログで対象文字と基準文字を指定（対象文字は自動入力、複数ある場合は最頻出記号。手動上書きも可）
2. 複製とアウトライン化で中心Y座標を比較
3. 差分をすべての対象文字に適用

対象:
- テキストフレーム（複数選択可、一括適用対応）

対象外:
- アウトライン済み、非テキストオブジェクト

オリジナルアイデア:
Egor Chistyakov https://x.com/tchegr

オリジナルからの変更点:
- 対象文字は自動入力（複数ある場合には最頻出記号を選択）
- 手動での上書き入力も可能
- 複数のテキストオブジェクトに対しても一括適用可能
- 対象文字に複数文字を同時指定できるよう対応

更新履歴:
- v1.0.0(2025-07-04): 初版リリース
- v1.0.6(2025-07-05): 複数の対象文字を指定し、一括で調整可能に対応
*/

/*
UIラベルを定義。
表示順に並べ替え、使用されていないラベルは削除。
*/
var LABELS = {
    dialogTitle: { ja: "ベースライン調整", en: "Adjust Baseline" },
    infoTextMsg: { ja: "対象文字を縦方向に揃えます。", en: "This will align selected symbol vertically." },
    targetCharLabel: { ja: "対象文字:", en: "Target Character:" },
    baseCharLabel: { ja: "基準文字:", en: "Reference Character:" },
    okBtnLabel: { ja: "調整", en: "Adjust" },
    cancelBtnLabel: { ja: "キャンセル", en: "Cancel" },
    selectFrameMsg: { ja: "テキストフレームを選択してください。", en: "Select one or more text frames." },
    docOpenMsg: { ja: "ドキュメントが開かれていません。", en: "No document open." },
    invalidCharMsg: { ja: "対象文字は1文字以上、基準文字は1文字を入力してください。", en: "Enter at least one target character and exactly one reference character." },
    notFoundMsg: { ja: "対象文字が含まれていません。", en: "Target character not found." },
    errorMsg: { ja: "エラー: ", en: "Error: " }
};

function getLang() {
    /*
    ロケールから言語を判定。日本語なら'ja'、それ以外は'en'。
    */
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}

function getCenterY(item) {
    /*
    アイテムのジオメトリック境界から中心のY座標を計算して返す。
    */
    var bounds = item.geometricBounds;
    return bounds[1] - (bounds[1] - bounds[3]) / 2;
}

function createOutlineAndGetCenterY(textFrame, character) {
    /*
    指定文字を含むテキストフレームを複製し、アウトライン化後に中心Y座標を取得する。
    */
    var tempFrame = textFrame.duplicate();
    tempFrame.contents = character;
    tempFrame.filled = false;
    tempFrame.stroked = false;
    var outline = tempFrame.createOutline();
    var centerY = getCenterY(outline);
    outline.remove();
    return centerY;
}

function getSymbolFrequency(sel) {
    /*
    選択テキスト内の記号・非英数字の出現頻度を集計して返す。
    */
    var charCount = {};
    for (var i = 0; i < sel.length; i++) {
        var item = sel[i];
        if (item.typename == "TextFrame") {
            var selText = item.contents;
            for (var j = 0; j < selText.length; j++) {
                var c = selText.charAt(j);
                if (!c.match(/^[A-Za-z0-9\s]$/)) {
                    charCount[c] = (charCount[c] || 0) + 1;
                }
            }
        }
    }
    return charCount;
}

function showDialog() {
    /*
    ユーザーに対象文字と基準文字を入力させるダイアログを表示し、入力結果を返す。
    */
    var lang = getLang();
    var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
    dialog.orientation = "column";
    dialog.alignChildren = "left";

    var infoText = dialog.add("statictext", undefined, LABELS.infoTextMsg[lang]);
    infoText.alignment = "left";

    /* デフォルト対象文字（複数選択でも最頻出記号を抽出） */
    var defaultTarget = "";
    var sel = app.activeDocument.selection;
    if (sel && sel.length > 0) {
        var charCount = getSymbolFrequency(sel);
        var maxCount = 0;
        for (var key in charCount) {
            if (charCount[key] > maxCount) {
                maxCount = charCount[key];
                defaultTarget = key;
            }
        }
    }

    var inputGroup = dialog.add("group");
    inputGroup.orientation = "column";
    inputGroup.alignChildren = "left";
    inputGroup.margins = [15, 5, 15, 5];

    var targetGroup = inputGroup.add("group");
    targetGroup.add("statictext", undefined, LABELS.targetCharLabel[lang]);
    var targetInput = targetGroup.add("edittext", undefined, defaultTarget);
    targetInput.characters = 5;
    targetInput.active = true;

    var refGroup = inputGroup.add("group");
    refGroup.add("statictext", undefined, LABELS.baseCharLabel[lang]);
    var refInput = refGroup.add("edittext", undefined, "0");
    refInput.characters = 5;

    var buttonGroup = dialog.add("group");
    buttonGroup.alignment = "center";
    var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancelBtnLabel[lang]);
    var okBtn = buttonGroup.add("button", undefined, LABELS.okBtnLabel[lang], { name: "ok" });

    okBtn.onClick = function () {
        if (targetInput.text.length == 0) {
            alert(LABELS.invalidCharMsg[lang]);
            return;
        }
        if (refInput.text.length != 1) {
            alert(LABELS.invalidCharMsg[lang]);
            return;
        }
        dialog.close(1);
    };
    cancelBtn.onClick = function () { dialog.close(0); };

    if (dialog.show() == 1) {
        return { target: targetInput.text, reference: refInput.text };
    }
    return null;
}

function main() {
    /*
    メイン処理。ドキュメントと選択状態をチェックし、ダイアログを表示して調整を実行。
    */
    try {
        if (app.documents.length == 0) {
            alert(LABELS.docOpenMsg[getLang()]);
            return;
        }

        var selection = app.activeDocument.selection;
        if (!selection || selection.length == 0) {
            alert(LABELS.selectFrameMsg[getLang()]);
            return;
        }

        var input = showDialog();
        if (!input) return;

        for (var i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (item.typename == "TextFrame") {
                var contents = item.contents;
                for (var c = 0; c < input.target.length; c++) {
                    var targetChar = input.target.charAt(c);
                    if (contents.indexOf(targetChar) == -1) {
                        continue;
                    }

                    var refCenterY = createOutlineAndGetCenterY(item, input.reference);
                    var targetCenterY = createOutlineAndGetCenterY(item, targetChar);
                    var yOffset = targetCenterY - refCenterY;

                    var chars = item.textRange.characters;
                    for (var j = 0; j < chars.length; j++) {
                        var ch = chars[j].contents;
                        if (ch && ch === targetChar) {
                            chars[j].characterAttributes.baselineShift = -yOffset;
                        }
                    }
                }
            }
        }

    } catch (e) {
        alert(LABELS.errorMsg[getLang()] + e);
    }
}

main();