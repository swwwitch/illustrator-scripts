#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ドキュメント内のすべてのガイドを削除します。
ロックされたレイヤーも一時的にロックを解除して対象にし、処理後に元のロック状態へ戻します。

詳細は README を参照してください。

### Overview

Deletes every guide in the document.
Locked layers are unlocked temporarily so their guides are included, then their lock state is restored.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "DeleteAllGuides";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-07-11";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2025-07-11";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DeleteAllGuides.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/DeleteAllGuides.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

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
