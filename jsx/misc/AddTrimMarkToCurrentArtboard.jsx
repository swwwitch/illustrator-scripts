#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);
var SCRIPT_VERSION = "v1.1";

/*
### スクリプト名：

AddTrimMark.jsx

Version: v1.1

### 概要

- 常に現在のアートボードに対して、日本式トンボを作成するIllustrator用スクリプトです。
- 専用の「トンボ」レイヤーを最初に取得または作成し、その上にアートボード矩形を1つだけ作成します。
- その矩形を元にトリムマークを作成し、同じオブジェクトをそのままガイド化します。
- 「トンボ」レイヤーがロックされている場合はいったん解除して処理し、終了時に元のロック状態へ戻します。

### 主な機能

- 常に現在のアートボードを対象にトリムマークを作成
- 「トンボ」レイヤーを自動取得／未存在時は新規作成
- アートボード矩形1つのみを使用（複製なし）
- 同一オブジェクトでトンボ生成とガイド化を完結
- 「トンボ」レイヤーの元のロック状態を維持

### 処理の流れ

1. 「トンボ」レイヤーを取得し、なければ新規作成
2. 環境設定で日本式トンボをONに設定
3. 「トンボ」レイヤー上にアートボード矩形を1つ作成
4. その矩形を元にトリムマーク作成メニューを実行
5. 同じ矩形オブジェクトをガイド化
6. finally で選択解除を行い、「トンボ」レイヤーのロック状態を元に戻す

### 更新履歴

- v1.0 (20250205) : 初期バージョン
- v1.1 (20260401) : 「トンボ」レイヤーがロックされている場合はいったん解除して処理し、終了時に元のロック状態へ戻すように変更
*/

function main() {
    var doc = app.activeDocument;
    var targetObj = null;
    var trimLayer = null;

    /* 「トンボ」レイヤーを取得（なければ作成） / Get "Trim" layer (create if not exists) */
    for (var i = 0; i < doc.layers.length; i++) {
        if (doc.layers[i].name === "トンボ") {
            trimLayer = doc.layers[i];
            break;
        }
    }
    if (!trimLayer) {
        trimLayer = doc.layers.add();
        trimLayer.name = "トンボ";
    }
    var wasLocked = trimLayer.locked;
    if (wasLocked) {
        trimLayer.locked = false;
    }

    try {
        /* 日本式トンボをONに設定 / Enable Japanese-style trim marks */
        app.preferences.setBooleanPreference('cropMarkStyle', 1);

        /* 常にアクティブなアートボードを基に処理 / Always use the active artboard */
        var artboard = doc.artboards[doc.artboards.getActiveArtboardIndex()];
        var rect = artboard.artboardRect; // [左, 上, 右, 下]

        /* 「トンボ」レイヤー上にアートボード矩形を作成 / Create artboard rectangle on "Trim" layer */
        targetObj = trimLayer.pathItems.rectangle(rect[1], rect[0], rect[2] - rect[0], rect[1] - rect[3]);
        targetObj.filled = false;
        targetObj.stroked = false;

        /* 作成オブジェクトを選択状態にする / Select the created object */
        doc.selection = [targetObj];

        /* トリムマークを作成 / Create trim marks */
        app.executeMenuCommand('TrimMark v25');

        /* 作成オブジェクトをガイド化する / Convert the created object to a guide */
        targetObj.guides = true;
    } finally {
        doc.selection = null;

        if (trimLayer) {
            trimLayer.locked = wasLocked ? true : false;
        }
    }
}

main();