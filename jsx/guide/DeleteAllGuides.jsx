#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

$.localize = true;

/*
### スクリプト名：

DeleteAllGuides

### 概要

- ドキュメント内のすべてのガイドを削除します。
- ロックされたレイヤーも一時的にアンロックしてガイドを削除します。

### Main Features

- Remove all guides from the document
- Temporarily unlock locked layers to remove guides

### 処理の流れ

1. ドキュメントを確認（開いていない場合は終了）
2. すべてのレイヤーのロック状態を保存し、アンロック
3. ガイドロックを解除
4. すべてのガイドを削除
5. レイヤーのロック状態を元に戻す

### オリジナル、謝辞

特になし

### 更新履歴

- v1.0 (20250711) : 初期バージョン

### Script Name:

DeleteAllGuides

### Overview

- Removes all guides from the document.
- Temporarily unlocks locked layers to remove guides.

### Main Features

- Remove all guides from the document
- Temporarily unlock locked layers to remove guides

### Process Flow

1. Check if a document is open (exit if none)
2. Save lock states of all layers and unlock
3. Unlock document guides
4. Remove all guides
5. Restore original lock states

### Original / Credits

None

### Update History

- v1.0 (20250711) : Initial version
*/

// スクリプトバージョン / Script Version
var SCRIPT_VERSION = "v1.0";

function main() {
    if (app.documents.length === 0) {
        alert(LABELS.dialogTitle.ja + " / " + LABELS.dialogTitle.en); // ドキュメントが開かれていません / No document open
        return;
    }

    var doc = app.activeDocument;

    // すべてのレイヤーのロック状態を保存 / Save lock states of all layers
    var layers = doc.layers;
    var lockStates = [];
    for (var i = 0; i < layers.length; i++) {
        lockStates[i] = layers[i].locked;
        if (layers[i].locked) {
            layers[i].locked = false; // 一時的にアンロック / Temporarily unlock
        }
    }

    // ガイドロックを解除 / Unlock guides
    doc.guidesLocked = false;

    // すべてのガイドを削除 / Remove all guides
    var paths = doc.pathItems;
    for (var i = paths.length - 1; i >= 0; i--) {
        if (paths[i].guides) {
            try {
                paths[i].remove();
            } catch (e) {
                // 削除できない場合は無視 / Ignore if cannot remove
            }
        }
    }

    // ロック状態を元に戻す / Restore original lock states
    for (var j = 0; j < layers.length; j++) {
        layers[j].locked = lockStates[j];
    }
}

main();