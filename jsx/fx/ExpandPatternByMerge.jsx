#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトの塗りを［分割・拡張］で展開し、グループ化して［合流］の効果を掛けてからアピアランスを分割します。3D 効果を使わずに、パターンなどの塗りを通常のパスにまとめます。

### Overview

Expands the fills of the selected objects with Object > Expand, groups them, applies the Pathfinder Merge effect, then expands the appearance. Turns pattern fills and the like into regular paths without a 3D effect.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExpandPatternByMerge";         /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 選択にグラデーションが混ざっていたときの分割数（［分割・拡張］の既定値） / Gradient steps if the selection includes gradients (Expand default) */
    var EXPAND_GRADIENT_STEPS = 255;

    // =========================================
    // 一時アクション設定 / Temporary action settings
    // =========================================

    /* 一時アクションのセット名（衝突回避のためユニーク名） / Temporary action set name (unique to avoid collisions) */
    var ACTION_SET_NAME = "ExpandPatternByMerge_tmp";

    /* 一時アクションのアクション名 / Temporary action name */
    var ACTION_NAME = "Expand-fill";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ローカライズ（再利用パーツ） / Localization (reusable)

    /**
     * UI の言語を返す（"ja" で始まるロケールは日本語、それ以外は英語）
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return (String($.locale || "").indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = getCurrentLang();

    /**
     * LABELS から今の UI 言語の文言を取り出す。
     * @param {string|Object} labelRef - "dialog.title" のようなパス、または { ja, en }
     * @param {Object|Array} [placeholderValues] - { name: 値 } なら {name} を、[値, …] なら %1, %2 … を差し込む
     * @returns {string} 文言。パスが見つからなければパスの文字列、{ ja, en } が無ければ空文字
     */
    function getLabel(labelRef, placeholderValues) {
        var labelEntry = labelRef;
        if (typeof labelRef === "string") {
            var labelPathKeys = labelRef.split(".");
            labelEntry = LABELS;
            for (var i = 0; i < labelPathKeys.length && labelEntry != null; i++) {
                labelEntry = labelEntry[labelPathKeys[i]];
            }
        }
        var labelString;
        if (typeof labelEntry === "string") labelString = labelEntry;
        else if (labelEntry != null && labelEntry[uiLang] != null) labelString = labelEntry[uiLang];
        else if (labelEntry != null && labelEntry.en != null) labelString = labelEntry.en;
        else return (typeof labelRef === "string") ? labelRef : "";
        return fillLabelPlaceholders(String(labelString), placeholderValues);
    }

    /**
     * 項目名の文言の末尾にコロンを付ける（日本語は全角「：」、英語は半角「:」）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {Object|Array} [placeholderValues] - getLabel と同じ
     * @returns {string} コロン付きの文言
     */
    function labelText(labelRef, placeholderValues) {
        return getLabel(labelRef, placeholderValues) + (uiLang === "ja" ? "：" : ":");
    }

    /**
     * 「項目名：値」の1行を返す（日本語は「件数：5」、英語は「Count: 5」とコロンのあとに空白を入れる）
     * @param {string|Object} labelRef - getLabel と同じ
     * @param {string|number} value - コロンのあとに続ける値
     * @returns {string} 項目名と値をつないだ文字列
     */
    function labelValueText(labelRef, value) {
        return labelText(labelRef) + (uiLang === "ja" ? "" : " ") + value;
    }

    /**
     * 文言の {name} や %1 に値を差し込む
     * @param {string} labelString - 文言
     * @param {Object|Array} [placeholderValues] - { name: 値 } または [値, …]
     * @returns {string} 差し込んだ文言
     */
    function fillLabelPlaceholders(labelString, placeholderValues) {
        if (placeholderValues == null) return labelString;
        if (placeholderValues instanceof Array) {
            /* 大きい番号から置き換え、%1 が %10 の一部を置き換えないようにする / Replace from the highest index so %1 does not eat into %10 */
            for (var i = placeholderValues.length; i >= 1; i--) {
                labelString = labelString.split("%" + i).join(String(placeholderValues[i - 1]));
            }
            return labelString;
        }
        for (var placeholderKey in placeholderValues) {
            if (!placeholderValues.hasOwnProperty(placeholderKey)) continue;
            labelString = labelString.split("{" + placeholderKey + "}").join(String(placeholderValues[placeholderKey]));
        }
        return labelString;
    }

    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "展開するオブジェクトを選択してください。", en: "Please select the objects to expand." },
            expandFailed: { ja: "［分割・拡張］を実行できませんでした。", en: "Could not run Object > Expand." }
        },
        fallbackName: {
            expandGroup: { ja: "パターンの分割", en: "Expanded Pattern" }
        }
    };

    // =========================================
    // 効果の XML / Effect XML
    // =========================================

    /**
     * パスファインダー［合流］効果の LiveEffect XML を返す。
     * メニューコマンドや isPre="1" では「内容」の上に入るため、isPre を付けずに「内容」の下へ積む。
     * パラメーターは mark1bean/live-effect-functions-for-illustrator の LE_PathFinder の既定値に準拠。
     * @returns {string} applyEffect() に渡すXML
     */
    function buildMergeEffectXML() {
        var mergeParams = [
            "I Command 8",              /* 8 = 合流 / Merge */
            "B ConvertCustom 1",
            "R Precision 10",           /* 精度（マイクロメートル） / precision (micrometers) */
            "B RemovePoints 1"          /* 余分なポイントを削除 / remove redundant points */
        ].join(" ");
        return '<LiveEffect name="Adobe Pathfinder"><Dict data="' + mergeParams + ' "/></LiveEffect>';
    }

    /**
     * 分割・拡張（ai_plugin_expand）で塗りだけを展開するアクションソースを組み立てる。
     * パラメーターは ExpandGradient.jsx と同じキー（objt / fill / strk / smth / step）。
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @returns {string} .aia のソース
     */
    function buildExpandFillActionSource(setName, actionName) {
        var parameters = [
            buildParameterLine(1, 1868720756, "boolean", 0, ""),   /* objt: オブジェクト / Object */
            buildParameterLine(2, 1718185068, "boolean", 1, ""),   /* fill: 塗り / Fill */
            buildParameterLine(3, 1937011307, "boolean", 0, ""),   /* strk: 線 / Stroke */
            buildParameterLine(4, 1936553064, "boolean", 0, ""),   /* smth: グラデーションメッシュ / Gradient Mesh */
            buildParameterLine(5, 1937007984, "integer", EXPAND_GRADIENT_STEPS, "")  /* step: グラデーションの分割数 / gradient steps */
        ];
        /* 表示名「分割・拡張」 / Localized name */
        return buildActionSource(setName, actionName, "ai_plugin_expand", "e58886e589b2e383bbe68ba1e5bcb5", parameters);
    }

    /**
     * 1イベントだけのアクションセットを組み立てる。
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {string} internalName - プラグインの内部名
     * @param {string} localizedNameHex - パネル表示名のUTF-8 16進表現
     * @param {string[]} parameters - パラメーター行の配列
     * @returns {string} .aia のソース
     */
    function buildActionSource(setName, actionName, internalName, localizedNameHex, parameters) {
        return ''
            + '/version 3'
            + buildActionNameLine(setName)
            + '/isOpen 1'
            + '/actionCount 1'
            + '/action-1 {'
            + ' ' + buildActionNameLine(actionName)
            + ' /keyIndex 0'
            + ' /colorIndex 0'
            + ' /isOpen 1'
            + ' /eventCount 1'
            + ' /event-1 {'
            + ' /useRulersIn1stQuadrant 0'
            + ' /internalName (' + internalName + ')'
            + ' /localizedName ' + buildHexTextBlock(localizedNameHex)
            + ' /isOpen 1'
            + ' /isOn 1'
            + ' /hasDialog 1'
            + ' /showDialog 0'
            + ' /parameterCount ' + parameters.length
            + parameters.join('')
            + ' }'
            + '}';
    }

    /**
     * アクションのパラメーター1行を組み立てる。
     * @param {number} index - パラメーター番号（1始まり）
     * @param {number} key - パラメーターキー
     * @param {string} type - 値の型（boolean / integer / enumerated）
     * @param {number} value - 値
     * @param {string} localizedNameHex - 表示名のUTF-8 16進表現（不要なら空文字）
     * @returns {string} パラメーター行
     */
    function buildParameterLine(index, key, type, value, localizedNameHex) {
        return ' /parameter-' + index
            + ' { /key ' + key
            + ' /showInPalette 4294967295'
            + ' /type (' + type + ')'
            + (localizedNameHex ? ' /name ' + buildHexTextBlock(localizedNameHex) : '')
            + ' /value ' + value + ' }';
    }

    /**
     * アクション名の行を組み立てる。
     * @param {string} actionName - アクション名（ASCII）
     * @returns {string} /name の行
     */
    function buildActionNameLine(actionName) {
        return '/name ' + buildHexTextBlock(toActionHex(actionName)) + '\n';
    }

    /**
     * .aia の文字列表記 [ バイト数 16進 ] を組み立てる。
     * @param {string} hexText - 16進表現
     * @returns {string} [ バイト数 16進 ] の形式
     */
    function buildHexTextBlock(hexText) {
        return '[ ' + (hexText.length / 2) + ' ' + hexText + ' ]';
    }

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

    /**
     * 文字列を UTF-8 のバイト列の16進にする（アクション定義の /name・/localizedName 用）
     * @param {string} sourceText - 変換する文字列
     * @returns {string} 16進の文字列（2文字で1バイト）
     */
    function toActionHex(sourceText) {
        var utf8Text = unescape(encodeURIComponent(String(sourceText)));
        var hexText = "";
        for (var i = 0; i < utf8Text.length; i++) {
            var hexByte = utf8Text.charCodeAt(i).toString(16);
            hexText += (hexByte.length < 2 ? "0" : "") + hexByte;
        }
        return hexText;
    }

    /**
     * アクション定義の「/name [ バイト数 16進 ]」の3行を返す
     * @param {string} indent - 行頭の字下げ（"\t" など）
     * @param {string} nameText - 名前
     * @param {string} [fieldName] - 項目名（既定は "name"。"localizedName" など）
     * @returns {string[]} 3行ぶんの配列
     */
    function buildActionNameLines(indent, nameText, fieldName) {
        var nameHex = toActionHex(nameText);
        return [
            indent + "/" + (fieldName || "name") + " [ " + (nameHex.length / 2),
            indent + "\t" + nameHex,
            indent + "]"
        ];
    }

    /**
     * アクション定義を一時ファイルに書き出してセットを読み込む。読み込んだら一時ファイルは消す
     * （読み込んだ時点で解釈済みなので、以降の失敗でファイルが残らない）
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @returns {boolean} 読み込めたら true
     */
    function loadTemporaryActionSet(actionSource, setName) {
        var actionFile = new File(Folder.temp + "/" + setName + "_" + new Date().getTime() + ".aia");
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) throw new Error("cannot open " + actionFile.fsName);
            actionFile.write(actionSource);
            actionFile.close();
            /* 前回の失敗で同じ名前のセットが残っていれば外す / Remove a same-name set left by an earlier failure */
            unloadTemporaryActionSet(setName);
            app.loadAction(actionFile);
            return true;
        } catch (e) {
            $.writeln("loadTemporaryActionSet: " + e);
            return false;
        } finally {
            try { actionFile.close(); } catch (closeError) { /* 閉じ済み / already closed */ }
            try { actionFile.remove(); } catch (removeError) { /* 消せなくても続ける / keep going */ }
        }
    }

    /**
     * 一時アクションのセットを解除する（読み込まれていなくてもエラーにしない）
     * @param {string} setName - アクションセット名
     * @returns {void}
     */
    function unloadTemporaryActionSet(setName) {
        try {
            app.unloadAction(setName, "");
        } catch (e) {
            /* 読み込まれていない / not loaded */
        }
    }

    /**
     * アクション定義を読み込んで1回実行し、解除する。途中で失敗しても解除は必ず試みる
     * @param {string} actionSource - アクション定義のテキスト
     * @param {string} setName - アクションセット名
     * @param {string} actionName - 実行するアクション名
     * @returns {boolean} 実行できたら true
     */
    function runTemporaryAction(actionSource, setName, actionName) {
        if (!loadTemporaryActionSet(actionSource, setName)) return false;
        try {
            app.doScript(actionName, setName);
            return true;
        } catch (e) {
            $.writeln("runTemporaryAction: " + e);
            return false;
        } finally {
            unloadTemporaryActionSet(setName);
        }
    }

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択の塗りを分割・拡張し、グループ化して［合流］を掛け、アピアランスを分割する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var targetDoc = app.activeDocument;
        if (targetDoc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        /* ［オブジェクト］＞［分割・拡張］（塗りのみ） / Object > Expand (fill only) */
        if (!runTemporaryAction(buildExpandFillActionSource(ACTION_SET_NAME, ACTION_NAME), ACTION_SET_NAME, ACTION_NAME)) {
            alert(getLabel("alert.expandFailed"));
            return;
        }

        app.executeMenuCommand("group"); /* 展開した結果をグループ化 / Group the expanded result */
        var expandGroup = targetDoc.selection[0];
        expandGroup.name = getLabel("fallbackName.expandGroup");
        expandGroup.applyEffect(buildMergeEffectXML());

        /* 効果の適用を確定させてから分割する（redraw しないと分割が効かない）
           Commit the effect before expanding; without redraw the expand has no effect */
        targetDoc.selection = [expandGroup];
        app.redraw();
        app.executeMenuCommand("expandStyle"); /* アピアランスを分割 / Expand Appearance */
    }

    main();

})();
