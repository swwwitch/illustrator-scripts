#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ConvertDocumentToGrayscale
Meta-Default-language=ja
Name=ドキュメント全体をグレースケールに変換してスウォッチを削除
Name[en]=Convert the whole document to grayscale and delete swatches
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.0.0
Release-Date=2026-11-21
Tags=color,カラー
X-com.swwwitch-Script-Name=ConvertDocumentToGrayscale
X-com.swwwitch-Updated=2026-11-21
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertDocumentToGrayscale.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertDocumentToGrayscale.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ConvertDocumentToGrayscale.jsx
Description-BEGIN
[ja]
ドキュメント内のすべてのオブジェクトを［編集］→［カラーを編集］→［グレースケールに変換］でモノクロにし、スウォッチを削除します。
[en]
Turns every object in the document monochrome with Edit > Edit Colors > Convert to Grayscale, then deletes the swatches.
Description-END
SCRIPTMETA-END
*/

/*
### 注意
- ロック・非表示のレイヤーとオブジェクトも一時的に解除して変換し、終わったら元に戻します。
- リンク画像、シンボルの定義、パターン・ブラシの中身、グラフィックスタイルは変換されません。
- ［なし］［レジストレーション］など削除できないスウォッチは残ります。

### Notes
- Locked or hidden layers and objects are unlocked and shown while converting, then restored.
- Linked images, symbol definitions, the contents of patterns and brushes, and graphic styles are not converted.
- Swatches that cannot be deleted, such as [None] and [Registration], remain.
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読んで SCRIPT_* に入れる（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; it is read into SCRIPT_* at run time */

/**
 * このスクリプトの先頭にある SCRIPTMETA の項目（Description の手前まで）を読む
 * @returns {Object} 項目名 → 値（同じ項目は最初の値）。ファイルが読めなければ空のオブジェクト
 */
function readScriptMeta() {
    var metaFields = {};
    var scriptFile = new File($.fileName);
    scriptFile.encoding = "UTF-8";
    if (!scriptFile.exists || !scriptFile.open("r")) return metaFields;
    /* 境界の文字列は分けて書く（コードの中に境界があると SCRIPTMETA を読むアプリが誤読するため）
       Marker strings are split so apps that scan for SCRIPTMETA do not find them in the code */
    var blockBegin = "SCRIPTMETA" + "-BEGIN";
    var blockEnd = "SCRIPTMETA" + "-END";
    var descriptionBegin = "Description" + "-BEGIN";
    var isInBlock = false;
    for (var lineCount = 0; lineCount < 200 && !scriptFile.eof; lineCount++) {
        var metaLine = scriptFile.readln().replace(/\r$/, "");
        if (metaLine === blockBegin) {
            isInBlock = true;
        } else if (isInBlock) {
            if (metaLine === blockEnd || metaLine === descriptionBegin) break;
            var separatorIndex = metaLine.indexOf("=");
            var fieldKey = metaLine.substring(0, separatorIndex);
            if (separatorIndex > 0 && !metaFields.hasOwnProperty(fieldKey)) metaFields[fieldKey] = metaLine.substring(separatorIndex + 1);
        }
    }
    scriptFile.close();
    return metaFields;
}

var SCRIPT_META = readScriptMeta();

/**
 * SCRIPTMETA の項目の値を返す
 * @param {string} fieldKey - 項目名（"Version"、"X-com.swwwitch-Updated" など）
 * @returns {string} 値。項目が無ければ空文字
 */
function getScriptMetaValue(fieldKey) {
    return SCRIPT_META.hasOwnProperty(fieldKey) ? SCRIPT_META[fieldKey] : "";
}

/* SCRIPTMETA が読めないとき（ファイルを経由しない実行）は、スクリプト名だけファイル名から補う / fall back to the file name */
var SCRIPT_NAME           = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名 / script name */
var SCRIPT_VERSION        = getScriptMetaValue("Version") ? "v" + getScriptMetaValue("Version") : ""; /* バージョン / version */
var SCRIPT_AUTHOR         = getScriptMetaValue("Author");                       /* 作者 / author */
var SCRIPT_RELEASED       = getScriptMetaValue("Release-Date");                 /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED        = getScriptMetaValue("X-com.swwwitch-Updated");       /* 更新日 / last updated */
var SCRIPT_DESCRIPTION_JA = getScriptMetaValue("Name");                         /* README の見出し / README title (Japanese) */
var SCRIPT_README_JA      = getScriptMetaValue("X-com.swwwitch-README-JA");     /* README（日本語） */
var SCRIPT_README_EN      = getScriptMetaValue("X-com.swwwitch-README-EN");     /* README (English) */
var SCRIPT_ARTICLE_URL    = getScriptMetaValue("X-com.swwwitch-Article-URL");   /* 紹介記事 / article URL */
var SCRIPT_SOURCE_URL     = getScriptMetaValue("X-com.swwwitch-Source-URL");    /* 最新版のソース / latest source */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ロック・表示の一時解除 / Temporary unlock and show
    // =========================================

    /**
     * レイヤーとサブレイヤーのロック・非表示を解除し、元の状態を控える
     * @param {Layers} layers - 対象のレイヤー
     * @param {Object[]} records - 控え（{target, key, value} を追加する）
     * @returns {void}
     */
    function unlockAndShowLayers(layers, records) {
        for (var i = 0; i < layers.length; i++) {
            var layer = layers[i];
            if (!layer.visible) {
                layer.visible = true;
                records.push({ target: layer, key: "visible", value: false });
            }
            if (layer.locked) {
                layer.locked = false;
                records.push({ target: layer, key: "locked", value: true });
            }
            unlockAndShowLayers(layer.layers, records);
        }
    }

    /**
     * ドキュメント内のオブジェクト（グループの中も含む）のロック・非表示を解除し、元の状態を控える
     * @param {Document} doc - 対象ドキュメント
     * @param {Object[]} records - 控え（{target, key, value} を追加する）
     * @returns {void}
     */
    function unlockAndShowItems(doc, records) {
        var pageItems = doc.pageItems;
        for (var i = 0; i < pageItems.length; i++) {
            var item = pageItems[i];
            try {
                if (item.hidden) {
                    item.hidden = false;
                    records.push({ target: item, key: "hidden", value: true });
                }
                if (item.locked) {
                    item.locked = false;
                    records.push({ target: item, key: "locked", value: true });
                }
            } catch (e) {
                /* 変更できないオブジェクトはそのまま / Leave items that cannot be changed */
            }
        }
    }

    /**
     * 控えた状態を逆順に戻す（オブジェクトを先に、レイヤーを後に戻すため、控えた順の逆にする）
     * @param {Object[]} records - unlockAndShowLayers / unlockAndShowItems の控え
     * @returns {void}
     */
    function restoreStates(records) {
        for (var i = records.length - 1; i >= 0; i--) {
            try {
                records[i].target[records[i].key] = records[i].value;
            } catch (e) {
                /* 戻せないものは飛ばす / Skip what cannot be restored */
            }
        }
    }

    // =========================================
    // スウォッチ / Swatches
    // =========================================

    /**
     * スウォッチをすべて削除する（［なし］［レジストレーション］など削除できないものは残る）
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function deleteAllSwatches(doc) {
        for (var i = doc.swatches.length - 1; i >= 0; i--) {
            try {
                doc.swatches[i].remove();
            } catch (e) {
                /* 削除できないスウォッチ / Not removable */
            }
        }
    }

    /**
     * 中身が空になったスウォッチグループを削除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function removeEmptySwatchGroups(doc) {
        for (var i = doc.swatchGroups.length - 1; i >= 0; i--) {
            try {
                if (doc.swatchGroups[i].getAllSwatches().length === 0) {
                    doc.swatchGroups[i].remove();
                }
            } catch (e) {
                /* 既定のグループなど削除できないものは残す / Leave groups that cannot be removed */
            }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメント全体をグレースケールに変換し、スウォッチを削除する
     * @returns {void}
     */
    function main() {
        if (!app.documents.length) return;
        var doc = app.activeDocument;

        /* 文字の編集中だと［すべてを選択］が文字の範囲になるので、選択ツールに持ち替えて選択を解除
           While editing text, Select All selects characters; switch to the Selection tool and deselect */
        try {
            app.selectTool("Adobe Select Tool");
        } catch (e) {
            /* 持ち替えられなくても続ける / Continue even if the tool cannot be switched */
        }
        doc.selection = null;

        /* レイヤーを先に解除しないと、中のオブジェクトを変更できない
           Layers come first; items inside a locked or hidden layer cannot be changed */
        var stateRecords = [];
        unlockAndShowLayers(doc.layers, stateRecords);
        unlockAndShowItems(doc, stateRecords);

        try {
            app.executeMenuCommand("selectall");
            if (doc.selection.length) {
                /* ［編集］→［カラーを編集］→［グレースケールに変換］ / Edit > Edit Colors > Convert to Grayscale */
                app.executeMenuCommand("Colors7");
            }
            doc.selection = null;
        } finally {
            restoreStates(stateRecords);
        }

        /* 変換後はスウォッチを参照するオブジェクトが無いので、まとめて削除
           After the conversion no object refers to a swatch, so delete them all */
        deleteAllSwatches(doc);
        removeEmptySwatchGroups(doc);
    }

    main();

})();
