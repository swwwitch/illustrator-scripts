#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

アクティブレイヤーを「テンプレート」属性（ロック・印刷不可・画像を薄く表示）にします。ダイアログを出さずに即実行します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeTemplateLayer.md

### Overview

Turns the active layer into a template layer — locked, non-printing and dimmed images. It runs immediately, without a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MakeTemplateLayer.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MakeTemplateLayer";            /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.2";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-28";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MakeTemplateLayer.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MakeTemplateLayer.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 一時アクション設定 / Temporary action settings
    // =========================================
    var ACTION_SET_NAME = "DynamicActionMakeTemplate";
    var ACTION_NAME = "TemplateON";

    /* レイヤーオプションのプリセット（テンプレート ON 固定）/ Layer-option preset (template ON) */
    var TEMPLATE_ON_OPTIONS = {
        template: true,   /* tmpl テンプレート */
        show: true,       /* show 表示 */
        lock: true,       /* lock ロック */
        preview: true,    /* prvw プレビュー */
        print: false,     /* prnt プリント */
        dim: true,        /* dim. 画像を薄く表示 */
        dimPercent: 50    /* 薄く表示の％（dim が true のときだけ書き出す） */
    };

    // =========================================
    // ローカライズ / Localization
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // ローカライズ（再利用パーツ） / Localization (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内のローカライズ節（LABELS の直前）に貼る。
    //    uiLang を使うコード（StepperButtons・LinkToggle の部品など）より前に置く
    // 2. 識別子は uiLang / getCurrentLang / getLabel / labelText / labelValueText / fillLabelPlaceholders。
    //    同じ役割の既存の関数・変数（getCurrentLanguage、currentLanguage、formatLabel など）は消して、これに寄せる
    // 3. 呼び出しはどちらの形でもよい（混ぜてもよい）
    //      getLabel("dialog.title")        … パス
    //      getLabel(LABELS.dialog.title)   … { ja, en } を直接
    //      getLabel("alert.count", { count: 3 })  … "{count} 個" の {count} を差し込む
    //      getLabel("alert.range", [1, 10])       … "%1〜%2" の %1・%2 を差し込む
    //      labelText("fieldLabel.width")   … 末尾にコロン（日本語は全角「：」、英語は半角「:」）
    //      labelValueText("message.count", 5) … 「件数：5」／「Count: 5」（値が続く1行。英語はコロンのあとに空白）
    // 4. 見つからないパスはパスの文字列をそのまま返す（表示で気づけるように）。{ ja, en } が無いときは空文字
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // ローカライズ（再利用パーツ）ここまで / End of the reusable localization
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            layerLocked: { ja: "アクティブレイヤーがロックされているため、実行できません。", en: "The active layer is locked." },
            layerHidden: { ja: "アクティブレイヤーが非表示のため、実行できません。", en: "The active layer is hidden." },
            actionFailed: { ja: "テンプレート属性の適用に失敗しました。", en: "Failed to apply the template attribute." }
        }
    };

    // =========================================
    // 一時アクション生成 / Temporary action generation
    // =========================================

    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼
    // 一時アクション（再利用パーツ） / Temporary action (reusable)
    //
    // 【移植手順 / How to port】
    // 1. ▼〜▲ をまるごと、コピー先の IIFE 内に貼る。
    //    識別子は runTemporaryAction / loadTemporaryActionSet / unloadTemporaryActionSet / toActionHex / buildActionNameLines
    // 2. アクション定義は配列＋join("\n") で組み立てる（''' は ES3 の構文エラー）。
    //    セット名・アクション名は英数字にする。/name [ n 16進 ] は buildActionNameLines で作るとバイト数がずれない
    //      var actionSource = [
    //          "/version 3"
    //      ].concat(buildActionNameLines("", "MySet"), [
    //          "/isOpen 1", "/actionCount 1", "/action-1 {"
    //      ], buildActionNameLines("\t", "myAction"), [ … ]).join("\n");
    // 3. 1回だけ実行するとき:
    //      if (!runTemporaryAction(actionSource, "MySet", "myAction")) alert(getLabel("alert.actionFailed"));
    //    何度も実行するとき（オブジェクトごとなど）は、読み込み・解除を1回ずつにする:
    //      if (!loadTemporaryActionSet(actionSource, "MySet")) { alert(…); return; }
    //      try { for (…) app.doScript("myAction", "MySet"); } finally { unloadTemporaryActionSet("MySet"); }
    // 4. 失敗は例外にせず false で返す（$.writeln に理由を出す）。警告を出すかはコピー先で決める
    // ▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼▼

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

    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲
    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action
    // ▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲▲

    /**
     * 記録済みの .aia から採取した internalName / key を使い、オプションからアクション定義を組み立てる
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {string} layerName - 対象レイヤー名（リネームを防ぐため注入する）
     * @param {Object} layerOptions - レイヤーオプションのプリセット
     * @returns {string} .aia の内容
     */
    function buildActionSource(setName, actionName, layerName, layerOptions) {
        var parameterLines = buildLayerParameterLines(layerName, layerOptions);

        var parameterBlock = '';
        for (var i = 0; i < parameterLines.length; i++) {
            parameterBlock += ' /parameter-' + (i + 1) + ' { ' + parameterLines[i] + ' }\n';
        }

        return ''
            + '/version 3\n'
            + buildActionNameLine(setName)
            + '/isOpen 1\n'
            + '/actionCount 1\n'
            + '/action-1 {\n'
            + ' ' + buildActionNameLine(actionName)
            + ' /keyIndex 0\n'
            + ' /colorIndex 0\n'
            + ' /isOpen 1\n'
            + ' /eventCount 1\n'
            + ' /event-1 {\n'
            + ' /useRulersIn1stQuadrant 0\n'
            + ' /internalName (ai_plugin_Layer)\n'
            + ' /localizedName [ 9 e8a1a8e7a4ba203a20 ]\n'
            + ' /isOpen 1\n'
            + ' /isOn 1\n'
            + ' /hasDialog 1\n'
            + ' /showDialog 0\n'
            + ' /parameterCount ' + parameterLines.length + '\n'
            + parameterBlock
            + ' }\n'
            + '}\n';
    }

    /**
     * parameter-* を1行ずつ生成する。key は記録済み .aia の FourCC（tmpl/show/lock/prvw/prnt/dim.）
     * dim が true のときだけ末尾に「薄く表示の％」行を追加する
     * @param {string} layerName - 対象レイヤー名
     * @param {Object} layerOptions - レイヤーオプションのプリセット
     * @returns {string[]} パラメーター行
     */
    function buildLayerParameterLines(layerName, layerOptions) {
        var lines = [];
        lines.push('/key 1836411236 /showInPalette 4294967295 /type (integer) /value 4');                                   /* カラー（ラベル色） */
        lines.push('/key 1851878757 /showInPalette 4294967295 /type (ustring) /value [ 36 e383ace382a4e383a4e383bce38391e3838de383abe382aae38397e382b7e383a7e383b3 ]'); /* 名前ラベル */
        lines.push('/key 1953068140 /showInPalette 4294967295 /type (ustring) /value ' + buildUstringValue(layerName));     /* レイヤー名（動的注入） */
        lines.push('/key 1953329260 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(layerOptions.template));   /* tmpl テンプレート */
        lines.push('/key 1936224119 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(layerOptions.show));       /* show 表示 */
        lines.push('/key 1819239275 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(layerOptions.lock));       /* lock ロック */
        lines.push('/key 1886549623 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(layerOptions.preview));    /* prvw プレビュー */
        lines.push('/key 1886547572 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(layerOptions.print));      /* prnt プリント */
        lines.push('/key 1684630830 /showInPalette 4294967295 /type (boolean) /value ' + boolBit(layerOptions.dim));        /* dim. 画像を薄く表示 */
        if (layerOptions.dim) {
            lines.push('/key 1885564532 /showInPalette 4294967295 /type (unit real) /value ' + layerOptions.dimPercent.toFixed(1) + ' /unit 592474723'); /* 薄く表示の％ */
        }
        return lines;
    }

    /**
     * 真偽値を .aia の boolean 値（1 / 0）にする
     * @param {boolean} flag - 真偽値
     * @returns {number} 1 または 0
     */
    function boolBit(flag) {
        return flag ? 1 : 0;
    }

    /**
     * アクションセット名・アクション名の /name 行を作る（名前は英数字）
     * @param {string} actionName - 名前
     * @returns {string} /name 行
     */
    function buildActionNameLine(actionName) {
        var nameHex = toActionHex(actionName);
        return '/name [ ' + (nameHex.length / 2) + ' ' + nameHex + ' ]\n';
    }

    /**
     * ustring 値（[ バイト数 UTF-8のhex ]）を生成する。マルチバイトの名前にも対応
     * @param {string} sourceText - 元の文字列
     * @returns {string} ustring 値
     */
    function buildUstringValue(sourceText) {
        var valueHex = toActionHex(sourceText);
        return '[ ' + (valueHex.length / 2) + ' ' + valueHex + ' ]';
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * アクティブレイヤーをテンプレートレイヤーにする
     * @returns {void}
     */
    function makeTemplateLayer() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }

        var activeLayer = app.activeDocument.activeLayer;

        /* ロック・非表示のレイヤーには適用しない / Skip locked or hidden layers */
        if (activeLayer.locked) {
            alert(getLabel("alert.layerLocked"));
            return;
        }
        if (!activeLayer.visible) {
            alert(getLabel("alert.layerHidden"));
            return;
        }

        /* 実行前にアクティブレイヤー名を取得して注入する（リネーム防止）/ Inject the active layer name so the action does not rename it */
        var actionSource = buildActionSource(ACTION_SET_NAME, ACTION_NAME, activeLayer.name, TEMPLATE_ON_OPTIONS);
        if (!runTemporaryAction(actionSource, ACTION_SET_NAME, ACTION_NAME)) {
            alert(getLabel("alert.actionFailed"));
        }
    }

    makeTemplateLayer();

})();
