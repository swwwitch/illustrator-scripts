#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
■ スクリプト名：ApplyShearAppearance.jsx

■ 概要：
選択オブジェクトにアピアランスとしてシアー（傾斜）変形を適用します。
角度（-44°〜44°）と方向（水平方向または垂直方向）を指定でき、リアルタイムプレビューが可能です。

■ 処理の流れ：
1. オブジェクトを選択
2. ダイアログで角度と方向を指定
3. プレビューで即反映（Undo可能）
4. OKで確定、キャンセルで元に戻す

■ 対象：
- PathItem や GroupItem など applyEffect() が使えるアートワーク
- テキストや画像など applyEffect を持たないものは対象外

■ クレジット：
Originally created by kawamoto_α（あるふぁ（仮））さん
アピアランスでシアー.jsx
https://sysys.blog.shinobi.jp/Entry/53/

■ 変更点
・プレビュー機能を追加
・選択しているオブジェクトから水平／垂直を判断
・en/ja ラベル対応
・ダイアログボックスを開いたら、値入力フィールドをアクティブに
・値変更の+ボタン、ｰボタンを追加

■ 更新履歴：
- v1.0.0（2025-06-09）初版
- v1.1.0（2025-07-05）+ / - ボタン追加、説明コメント整理
*/

// -------------------------------
// 日英ラベル定義
// -------------------------------
function getCurrentLang() {
    return ($.locale && $.locale.indexOf('ja') === 0) ? 'ja' : 'en';
}
var lang = getCurrentLang();
var LABELS = {
    dialogTitle:       { ja: "アピアランスでシアー", en: "Shear with Appearance" },
    angleLabel:        { ja: "角度（-44°～44°）:", en: "Angle (-44° to 44°):" },
    horizontal:        { ja: "水平", en: "Horizontal" },
    vertical:          { ja: "垂直", en: "Vertical" },
    preview:           { ja: "プレビュー", en: "Preview" },
    cancel:            { ja: "キャンセル", en: "Cancel" },
    ok:                { ja: "OK", en: "OK" },
    errorInvalidAngle: { ja: "角度は -44°〜44° の間で指定してください。", en: "Please enter an angle between -44° and 44°." },
    errorOccurred:     { ja: "エラーが発生しました：\n", en: "An error occurred:\n" }
};

main();

function main() {
    try {
        var selectedItems = app.selection;
        if (!selectedItems.length || !selectedItems[0].hasOwnProperty('applyEffect')) return;

        var previewEnabled = false;
        var lastPreviewParams = null;

        // 選択オブジェクトのサイズ取得
        var bounds = selectedItems[0].geometricBounds; // [y1, x1, y2, x2]
        var width = Math.abs(bounds[2] - bounds[0]);
        var height = Math.abs(bounds[3] - bounds[1]);

        // ダイアログ作成
        var dlg = new Window("dialog", LABELS.dialogTitle[lang]);
        dlg.orientation = "column";
        dlg.alignChildren = "left";

        // 角度入力グループ
        var angleGroup = dlg.add("group");
        angleGroup.orientation = "row";
        var angleLabel = angleGroup.add("statictext", undefined, LABELS.angleLabel[lang]);
        var angleInput = angleGroup.add("edittext", undefined, "0");
        angleInput.characters = 3;

        // + / - ボタンを縦に配置するグループを angleInput の右隣に追加
        var buttonColumn = angleGroup.add("group");
        buttonColumn.orientation = "column";
        var plusBtn = buttonColumn.add("button", [0, 0, 23, 20], "+");
        var minusBtn = buttonColumn.add("button", [0, 0, 23, 20], "-");
        buttonColumn.spacing = 2; // ボタン間のスペースを設定

        // ＋ボタン押下時の処理
        plusBtn.onClick = function () {
            var angle = parseFloat(angleInput.text);
            if (angle !== angle) angle = 0;
            angle += 1;
            if (angle > 44) angle = 44;
            angleInput.text = angle.toString();
            applyPreviewIfNeeded();
        };

        // −ボタン押下時の処理
        minusBtn.onClick = function () {
            var angle = parseFloat(angleInput.text);
            if (angle !== angle) angle = 0;
            angle -= 1;
            if (angle < -44) angle = -44;
            angleInput.text = angle.toString();
            applyPreviewIfNeeded();
        };

        // 方向選択グループ
        var directionGroup = dlg.add("group");
        directionGroup.orientation = "row";
        directionGroup.alignment = "center"; // 中央揃えに変更
        var horizontalRadio = directionGroup.add("radiobutton", undefined, LABELS.horizontal[lang]);
        var verticalRadio = directionGroup.add("radiobutton", undefined, LABELS.vertical[lang]);
        // サイズに基づき方向を自動設定
        if (width > height) {
            horizontalRadio.value = true;
            verticalRadio.value = false;
        } else {
            horizontalRadio.value = false;
            verticalRadio.value = true;
        }

        // ボタングループ
        var buttonGroup = dlg.add("group");
        buttonGroup.orientation = "row";
        buttonGroup.alignment = "right";
        var cancelBtn = buttonGroup.add("button", undefined, LABELS.cancel[lang]);
        var okBtn = buttonGroup.add("button", undefined, LABELS.ok[lang], { name: "ok" });

        // プレビュー適用処理
        function applyPreviewIfNeeded() {
            var angle = parseFloat(angleInput.text);
            var horiz = horizontalRadio.value;
            if (angle <= -45 || angle >= 45 || angle !== angle) return;

            var newParams = angle + "_" + (horiz ? "H" : "V");
            if (newParams === lastPreviewParams) return;

            undoPreview(); // 前のプレビューをUndo

            applyShearEffect(selectedItems, angle, horiz);
            app.redraw();

            lastPreviewParams = newParams;
            previewEnabled = true;
        }

        function undoPreview() {
            if (previewEnabled) {
                app.undo();
                previewEnabled = false;
            }
        }

        // イベント監視
        angleInput.onChanging = function () {
            applyPreviewIfNeeded();
        };
        angleInput.onChange = function () {
            applyPreviewIfNeeded();
        };
        horizontalRadio.onClick = verticalRadio.onClick = function () {
            applyPreviewIfNeeded();
        };

        // OKボタン処理
        okBtn.onClick = function () {
            var angle = parseFloat(angleInput.text);
            var horiz = horizontalRadio.value;
            if (angle <= -45 || angle >= 45 || angle !== angle) {
                alert(LABELS.errorInvalidAngle[lang]);
                return;
            }
            if (previewEnabled) {
                lastPreviewParams = null; // 適用済みをそのまま使用
            } else {
                applyShearEffect(selectedItems, angle, horiz);
            }
            dlg.close();
        };

        // キャンセルボタン処理
        cancelBtn.onClick = function () {
            undoPreview();
            dlg.close();
        };

        angleInput.active = true;
        dlg.show();

    } catch (e) {
        alert(LABELS.errorOccurred[lang] + e);
    }
}

// シアー変形適用関数
function applyShearEffect(targets, angleDeg, isHorizontal) {
    var angleRad = toRadians(angleDeg);
    var n = Math.tan((2 * Math.asin(Math.tan(-angleRad)) + Math.PI) / 4);
    var angleA = Math.atan(n);
    var wx = n / Math.SQRT2 / Math.sin(angleA);
    var hx = Math.cos(2 * angleA - Math.PI / 2) * wx;
    var hx2 = 1 / n;

    var transform1 = isHorizontal ?
        createTransformXML(wx, hx, toDegrees(angleA)) :
        createTransformXML(hx, wx, toDegrees(angleA));
    var transform2 = isHorizontal ?
        createTransformXML(1, hx2, -45) :
        createTransformXML(hx2, 1, -45);

    for (var i = 0; i < targets.length; i++) {
        targets[i].applyEffect(transform1);
        targets[i].applyEffect(transform2);
    }
}

// 角度変換
function toRadians(deg) {
    return deg * Math.PI / 180;
}
function toDegrees(rad) {
    return rad * 180 / Math.PI;
}

// Adobe Transform用XML生成
function createTransformXML(scaleH, scaleV, rotateDeg) {
    var xml = '<LiveEffect name="Adobe Transform"><Dict data="';
    xml += 'B transformPatterns 0 ';
    xml += 'B transformObjects 1 ';
    xml += 'B scaleLines 0 ';
    xml += 'B randomize 0 ';
    xml += 'B reflectX 0 ';
    xml += 'B reflectY 0 ';
    xml += 'I numCopies 0 ';
    xml += 'I pinPoint 4 ';
    xml += 'R moveH_Pts 0 ';
    xml += 'R moveV_Pts 0 ';
    xml += 'R scaleH_Percent ' + (scaleH * 100) + ' ';
    xml += 'R scaleH_Factor ' + scaleH + ' ';
    xml += 'R scaleV_Percent ' + (scaleV * 100) + ' ';
    xml += 'R scaleV_Factor ' + scaleV + ' ';
    xml += 'R rotate_Degrees ' + rotateDeg + ' ';
    xml += 'R rotate_Radians ' + toRadians(rotateDeg);
    xml += '"/></LiveEffect>';
    return xml;
}