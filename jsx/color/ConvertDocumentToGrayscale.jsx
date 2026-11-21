#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.ConvertDocumentToGrayscale
Meta-Default-language=ja
Name=ドキュメント全体を CMYK のモノクロにする（テキストは K100）
Name[en]=Turn the whole document into CMYK monochrome (text in K100)
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.1.0
Release-Date=2026-11-21
Tags=color,カラー
X-com.swwwitch-Script-Name=ConvertDocumentToGrayscale
X-com.swwwitch-Updated=2026-11-21
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ConvertDocumentToGrayscale.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ConvertDocumentToGrayscale.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/color/ConvertDocumentToGrayscale.jsx
Description-BEGIN
[ja]
ドキュメントを CMYK にしてから、すべてのオブジェクトとシンボルの中身をグレースケールに変換し、テキストを K100 にします。最後にスウォッチを削除します。
[en]
Switches the document to CMYK, converts every object and the contents of its symbols to grayscale, and sets text to K100. The swatches are deleted at the end.
Description-END
SCRIPTMETA-END
*/

/*
### 注意
- ［カラーを編集］→［CMYK に変換］を実行してから、RGB のドキュメントを CMYK カラーモードに変えます。
- テキストの塗りと線の単色は K100 にします。白（グレー 0%）は白のまま残します（KEEP_WHITE_TEXT）。グラデーション・パターンの文字はグレースケールのままです。
- ロック・非表示のレイヤーとオブジェクトも一時的に解除して処理し、終わったら元に戻します。
- シンボルは中身をグレースケールにした新しいシンボルを作り、インスタンスを差し替えて元のシンボルと入れ替えます（入れ子のシンボルも変換）。基準点、ダイナミックシンボルの種類、9スライスなどのシンボルオプションは引き継がれません。
- リンク画像、パターン・ブラシの中身、グラフィックスタイルは変換されません。
- ［なし］［レジストレーション］など削除できないスウォッチは残ります。

### Notes
- Runs Edit Colors > Convert to CMYK, then switches an RGB document to CMYK color mode.
- Solid fill and stroke colors of text become K100. White (0% gray) stays white (KEEP_WHITE_TEXT). Text with a gradient or pattern stays grayscale.
- Locked or hidden layers and objects are unlocked and shown while processing, then restored.
- Each symbol is replaced by a new symbol whose contents are converted to grayscale, and its instances are swapped to it (nested symbols too). Symbol options such as the registration point, the dynamic type and 9-slice scaling are not carried over.
- Linked images, the contents of patterns and brushes, and graphic styles are not converted.
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
    // ユーザー設定 / User Settings
    // =========================================

    /* 白い文字（グレー0%）は K100 にせず白のまま残す / Keep white text (0% gray) white instead of K100 */
    var KEEP_WHITE_TEXT = true;

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
    // シンボル / Symbols
    // =========================================

    /**
     * レイヤーとサブレイヤーの中身を、前面から背面の順に1つのグループへ移す
     * @param {Layer} layer - 移す元のレイヤー
     * @param {GroupItem} artGroup - 移す先のグループ
     * @returns {void}
     */
    function moveLayerContentsIntoGroup(layer, artGroup) {
        /* 移すと集合が縮むので、先に配列へ控える / Snapshot first; the collection shrinks while moving */
        var items = [];
        for (var i = 0; i < layer.pageItems.length; i++) {
            if (layer.pageItems[i] !== artGroup) items.push(layer.pageItems[i]);
        }
        for (var j = 0; j < items.length; j++) {
            items[j].move(artGroup, ElementPlacement.PLACEATEND);
        }
        /* 静的シンボルはサブレイヤーとして展開される / Static symbols expand into sublayers */
        for (var k = 0; k < layer.layers.length; k++) {
            moveLayerContentsIntoGroup(layer.layers[k], artGroup);
        }
    }

    /**
     * シンボルの定義を一時レイヤーに展開し、1つのグループにまとめる
     * @param {Document} doc - 対象ドキュメント
     * @param {Symbol} symbol - 展開するシンボル
     * @returns {{layer: Layer, art: GroupItem}} 一時レイヤーと展開したアートワーク
     */
    function expandSymbolArt(doc, symbol) {
        /* 展開ごとに空のレイヤーを使い、増えたものをそのまま生成物とみなす
           Use an empty layer per expansion so everything in it is the result */
        var workLayer = doc.layers.add();
        var instance = workLayer.symbolItems.add(symbol);
        /* breakLink() 後は instance の参照が無効になる / The instance reference is invalid after breakLink() */
        instance.breakLink();
        var artGroup = workLayer.groupItems.add();
        moveLayerContentsIntoGroup(workLayer, artGroup);
        return { layer: workLayer, art: artGroup };
    }

    /**
     * グループの中から指定した種類のオブジェクトを入れ子までたどって集める（group.textFrames などは直下しか返さない）
     * @param {GroupItem} container - 探すグループ
     * @param {string} itemTypename - 集める種類（"SymbolItem"、"TextFrame"）
     * @param {PageItem[]} found - 見つけたオブジェクト（追加する）
     * @returns {PageItem[]} found と同じ配列
     */
    function collectNestedItems(container, itemTypename, found) {
        for (var i = 0; i < container.pageItems.length; i++) {
            var item = container.pageItems[i];
            if (item.typename === itemTypename) {
                found.push(item);
            } else if (item.typename === "GroupItem") {
                collectNestedItems(item, itemTypename, found);
            }
        }
        return found;
    }

    /**
     * シンボルの中身をグレースケールにした新しいシンボルを作る。入れ子のシンボルは先に変換して差し替える
     * @param {Document} doc - 対象ドキュメント
     * @param {string} symbolName - 元のシンボル名
     * @param {Object} converted - 元の名前 → 新しいシンボル（変換済みの控え。失敗は null）
     * @param {string[]} completedNames - 変換し終えた順の元の名前（追加する）
     * @returns {Symbol|null} 新しいシンボル。変換できなければ null
     */
    function convertSymbolToGrayscale(doc, symbolName, converted, completedNames) {
        if (converted.hasOwnProperty(symbolName)) return converted[symbolName];
        /* 入れ子の循環に備えて先に印を付ける / Mark first in case of a cycle */
        converted[symbolName] = null;

        var expanded = null;
        try {
            expanded = expandSymbolArt(doc, doc.symbols.getByName(symbolName));

            /* 入れ子のインスタンスを変換済みのシンボルに差し替える / Point nested instances at converted symbols */
            var nestedItems = collectNestedItems(expanded.art, "SymbolItem", []);
            for (var i = 0; i < nestedItems.length; i++) {
                var nestedSymbol = convertSymbolToGrayscale(doc, nestedItems[i].symbol.name, converted, completedNames);
                if (nestedSymbol) nestedItems[i].symbol = nestedSymbol;
            }

            doc.selection = null;
            expanded.art.selected = true;
            /* ［編集］→［カラーを編集］→［グレースケールに変換］ / Edit > Edit Colors > Convert to Grayscale */
            app.executeMenuCommand("Colors7");
            doc.selection = null;
            applyK100ToTextFrames(collectNestedItems(expanded.art, "TextFrame", []), createK100Color());

            converted[symbolName] = doc.symbols.add(expanded.art);
            completedNames.push(symbolName);
        } catch (e) {
            /* 展開できないシンボルは元のまま / Leave symbols that cannot be expanded */
        } finally {
            if (expanded) {
                try {
                    expanded.layer.remove();
                } catch (removeError) {
                    /* 消せなくても続ける / Continue even if it cannot be removed */
                }
            }
        }
        return converted[symbolName];
    }

    /**
     * すべてのシンボルの中身をグレースケールにし、インスタンスを差し替えて元のシンボルと入れ替える
     * @param {Document} doc - 対象ドキュメント（ロック・非表示は解除済み）
     * @returns {void}
     */
    function convertAllSymbolsToGrayscale(doc) {
        if (!doc.symbols.length) return;

        var originalLayer = doc.activeLayer;
        var symbolNames = [];
        for (var i = 0; i < doc.symbols.length; i++) {
            symbolNames.push(doc.symbols[i].name);
        }

        /* 差し替える前に、元のシンボルごとのインスタンスを控える / Collect instances per original symbol before swapping */
        var instancesByName = {};
        for (var j = 0; j < doc.symbolItems.length; j++) {
            try {
                var instance = doc.symbolItems[j];
                var instanceSymbolName = instance.symbol.name;
                if (!instancesByName.hasOwnProperty(instanceSymbolName)) instancesByName[instanceSymbolName] = [];
                instancesByName[instanceSymbolName].push(instance);
            } catch (e) {
                /* 読めないインスタンスは飛ばす / Skip instances that cannot be read */
            }
        }

        var converted = {};
        var completedNames = [];
        for (var k = 0; k < symbolNames.length; k++) {
            convertSymbolToGrayscale(doc, symbolNames[k], converted, completedNames);
        }

        /* 外側のシンボルから先に消す（入れ子の元のシンボルは外側の定義が消えるまで使用中）
           Remove outer symbols first; a nested original stays in use until the outer definition is gone */
        for (var n = completedNames.length - 1; n >= 0; n--) {
            var symbolName = completedNames[n];
            var newSymbol = converted[symbolName];
            var instances = instancesByName[symbolName] || [];
            for (var m = 0; m < instances.length; m++) {
                try {
                    instances[m].symbol = newSymbol;
                } catch (e) {
                    /* 差し替えられないインスタンスは元のまま / Leave instances that cannot be swapped */
                }
            }
            try {
                doc.symbols.getByName(symbolName).remove();
                newSymbol.name = symbolName;
            } catch (e) {
                /* 元のシンボルが消せないときは、新しいシンボルを別名で残す / Keep the new symbol under another name */
                try {
                    newSymbol.name = symbolName + " (Grayscale)";
                } catch (renameError) {
                    /* 自動の名前のまま / Keep the automatic name */
                }
            }
        }

        try {
            doc.activeLayer = originalLayer;
        } catch (e) {
            /* 戻せなくても続ける / Continue even if it cannot be restored */
        }
    }

    // =========================================
    // テキストの K100 / K100 text
    // =========================================

    /**
     * K100（CMYK の黒）のカラーを作る
     * @returns {CMYKColor} C0 M0 Y0 K100
     */
    function createK100Color() {
        var blackColor = new CMYKColor();
        blackColor.black = 100;
        return blackColor;
    }

    /**
     * 白かどうかを返す（グレースケール変換後は GrayColor の 0% が白）
     * @param {Color} color - 調べるカラー
     * @returns {boolean} 白なら true
     */
    function isWhiteColor(color) {
        switch (color.typename) {
            case "GrayColor":
                return color.gray === 0;
            case "CMYKColor":
                return color.cyan === 0 && color.magenta === 0 && color.yellow === 0 && color.black === 0;
            case "RGBColor":
                return color.red === 255 && color.green === 255 && color.blue === 255;
            default:
                return false;
        }
    }

    /**
     * K100 に置き換えるカラーかどうかを返す（単色だけ。なし・グラデーション・パターンは対象外）
     * @param {Color} color - 文字の塗りか線のカラー
     * @returns {boolean} 置き換えるなら true
     */
    function shouldBecomeK100(color) {
        var solidTypes = { "GrayColor": true, "CMYKColor": true, "RGBColor": true, "SpotColor": true, "LabColor": true };
        if (!color || !solidTypes[color.typename]) return false;
        return !(KEEP_WHITE_TEXT && isWhiteColor(color));
    }

    /**
     * 対象の文字が続く範囲ごとに、文字の塗りか線へカラーを入れる
     * @param {Characters} characters - テキストフレームの文字
     * @param {boolean[]} targets - 文字ごとに置き換えるかどうか
     * @param {string} colorKey - "fillColor" か "strokeColor"
     * @param {Color} color - 入れるカラー
     * @returns {void}
     */
    function applyColorToCharacterRuns(characters, targets, colorKey, color) {
        var i = 0;
        while (i < targets.length) {
            if (!targets[i]) {
                i++;
                continue;
            }
            var runStart = i;
            while (i < targets.length && targets[i]) i++;
            /* characters[n] の length を伸ばすと範囲になる / Extending the length of characters[n] makes a range */
            var characterRun = characters[runStart];
            characterRun.length = i - runStart;
            characterRun.characterAttributes[colorKey] = color;
        }
    }

    /**
     * テキストの塗りと線の単色を K100 にする（先に全文字を読んでから書く）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {CMYKColor} blackColor - K100 のカラー
     * @returns {void}
     */
    function applyK100ToTextFrames(textFrames, blackColor) {
        for (var i = 0; i < textFrames.length; i++) {
            try {
                var characters = textFrames[i].characters;
                var fillTargets = [];
                var strokeTargets = [];
                for (var j = 0; j < characters.length; j++) {
                    var attributes = characters[j].characterAttributes;
                    fillTargets.push(shouldBecomeK100(attributes.fillColor));
                    strokeTargets.push(shouldBecomeK100(attributes.strokeColor));
                }
                applyColorToCharacterRuns(characters, fillTargets, "fillColor", blackColor);
                applyColorToCharacterRuns(characters, strokeTargets, "strokeColor", blackColor);
            } catch (e) {
                /* 変更できないテキストは飛ばす / Skip text that cannot be changed */
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
     * ドキュメント内のすべてのオブジェクトを選択してメニューコマンドを実行し、選択を解除する
     * @param {Document} doc - 対象ドキュメント（ロック・非表示は解除済み）
     * @param {string} menuCommandId - メニューコマンドの ID（"Colors8" など）
     * @returns {void}
     */
    function runOnAllArtwork(doc, menuCommandId) {
        app.executeMenuCommand("selectall");
        if (doc.selection.length) {
            app.executeMenuCommand(menuCommandId);
        }
        doc.selection = null;
    }

    /**
     * ドキュメントを CMYK にして全体をグレースケールに変換し、テキストを K100 にしてスウォッチを削除する
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
            /* ［編集］→［カラーを編集］→［CMYK に変換］でオブジェクトの色を CMYK にしてから、ドキュメントを CMYK に。
               先に CMYK にしておくと、K100 が RGB に置き換わらない
               Convert the objects with Edit > Edit Colors > Convert to CMYK, then switch the document to CMYK
               so K100 is not converted to RGB */
            runOnAllArtwork(doc, "Colors8");
            if (doc.documentColorSpace !== DocumentColorSpace.CMYK) {
                /* ［ファイル］→［ドキュメントのカラーモード］→［CMYK カラー］ / File > Document Color Mode > CMYK Color */
                app.executeMenuCommand("doc-color-cmyk");
            }

            /* シンボルの中身は［すべてを選択］では届かないので先に変換（一時レイヤーは処理後に消える）
               Select All does not reach symbol contents; convert them first (the work layers are removed) */
            convertAllSymbolsToGrayscale(doc);

            /* ［編集］→［カラーを編集］→［グレースケールに変換］ / Edit > Edit Colors > Convert to Grayscale */
            runOnAllArtwork(doc, "Colors7");

            /* テキストはグレーではなく K100 に / Text becomes K100 instead of gray */
            applyK100ToTextFrames(doc.textFrames, createK100Color());
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
