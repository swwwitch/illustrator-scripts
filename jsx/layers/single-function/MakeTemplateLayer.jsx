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
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

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
    var ACTION_FILE_NAME = "~/MakeTemplateLayerAction.aia";

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

    /**
     * UI言語を返す
     * @returns {string} "ja" または "en"
     */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = getCurrentLang();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            layerLocked: { ja: "アクティブレイヤーがロックされているため、実行できません。", en: "The active layer is locked." },
            layerHidden: { ja: "アクティブレイヤーが非表示のため、実行できません。", en: "The active layer is hidden." },
            actionFailed: { ja: "テンプレート属性の適用に失敗しました。", en: "Failed to apply the template attribute." },
            fileOpenFailed: { ja: "一時アクションファイルを作成できませんでした。", en: "Failed to open the temporary action file." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語のテキストを取り出す
     * @param {string} labelPath - "alert.noDocument" のようなパス
     * @returns {string} 表示言語のテキスト
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        return LABELS[labelPathKeys[0]][labelPathKeys[1]][uiLang];
    }

    // =========================================
    // 一時アクション生成 / Temporary action generation
    // =========================================

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
     * バイト列（1文字1バイトの文字列）を16進数の文字列にする
     * @param {string} byteString - 各文字が 0〜255 の文字列
     * @returns {string} 16進数の文字列
     */
    function bytesToHex(byteString) {
        var hexText = "";
        for (var i = 0; i < byteString.length; i++) {
            var hexValue = byteString.charCodeAt(i).toString(16);
            if (hexValue.length < 2) hexValue = "0" + hexValue;
            hexText += hexValue;
        }
        return hexText;
    }

    /**
     * アクションセット名・アクション名の /name 行を作る（名前は英数字）
     * @param {string} actionName - 名前
     * @returns {string} /name 行
     */
    function buildActionNameLine(actionName) {
        return '/name [ ' + actionName.length + ' ' + bytesToHex(actionName) + ' ]\n';
    }

    /**
     * ustring 値（[ バイト数 UTF-8のhex ]）を生成する。マルチバイトの名前にも対応
     * @param {string} sourceText - 元の文字列
     * @returns {string} ustring 値
     */
    function buildUstringValue(sourceText) {
        var byteString = unescape(encodeURIComponent(sourceText));
        return '[ ' + byteString.length + ' ' + bytesToHex(byteString) + ' ]';
    }

    // =========================================
    // 一時アクション実行 / Temporary action playback
    // =========================================

    /**
     * アクション定義を一時ファイルに書き出して読み込み、実行後にアンロードする
     * @param {string} actionSource - .aia の内容
     * @param {string} setName - アクションセット名
     * @param {string} actionName - アクション名
     * @param {string} actionFilePath - 一時ファイルのパス
     * @returns {void}
     */
    function playTemporaryAction(actionSource, setName, actionName, actionFilePath) {
        /* 前回の残りを外す（読み込まれていなければ例外）/ Unload a leftover set; throws when none is loaded */
        try { app.unloadAction(setName, ""); } catch (e) { }

        var actionFile = new File(actionFilePath);
        if (!actionFile.open("w")) {
            alert(getLabel("alert.actionFailed") + "\n\n" + getLabel("alert.fileOpenFailed"));
            return;
        }
        actionFile.write(actionSource);
        actionFile.close();

        var isActionLoaded = false;
        try {
            app.loadAction(actionFile);
            isActionLoaded = true;
            app.doScript(actionName, setName, false);
        } catch (e) {
            alert(getLabel("alert.actionFailed") + "\n\n" + e);
        } finally {
            /* 読み込み時点でパース済みなので一時ファイルは消してよい / The action is parsed on load, so the temp file can go */
            actionFile.remove();
            if (isActionLoaded) app.unloadAction(setName, "");
        }
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
        playTemporaryAction(actionSource, ACTION_SET_NAME, ACTION_NAME, ACTION_FILE_NAME);
    }

    makeTemplateLayer();

})();
