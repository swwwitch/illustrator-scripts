#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトをグループ化し、線を塗りに変換してから Pathfinder の「合流」をライブエフェクトとして適用し、アピアランスを分割します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MergeExpand.md

### Overview

Groups the selection, converts strokes to fills, applies Pathfinder Merge as a live effect, and then expands the appearance.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MergeExpand.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MergeExpand";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.1";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-27";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MergeExpand.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MergeExpand.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================

    /* 処理の最後にグループ解除するか / Whether to ungroup at the end */
    var UNGROUP_AT_END = false;

    // =========================================
    // ライブエフェクト / Live effect
    // =========================================

    /* Pathfinder Merge ライブエフェクトの XML（command 8 = Merge 固定）
       Live effect XML for Pathfinder Merge (command 8, all other params at defaults) */
    var PATHFINDER_MERGE_XML = '<LiveEffect name="Adobe Pathfinder" isPre="1">'
        + '<Dict data="I Command 8 B ConvertCustom 1 B ExtractUnpainted 1 R Mix 0.5 R Precision 10 B RemovePoints 1 R TrapAspect 1 B TrapConvertCustom 1 R TrapMaxTint 1 B TrapReverse 0 R TrapThickness 0.25 R TrapTint 0.4 R TrapTintTolerance 0.05">'
        + '<Entry name="DisplayString" value="Merge" valueType="S"/>'
        + '</Dict></LiveEffect>';

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var uiLang = detectUILang();

    var LABELS = {
        alert: {
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noSelection: { ja: "オブジェクトを選択してください。", en: "Please select one or more objects." }
        }
    };

    /**
     * LABELS からドット区切りのパスで表示言語の文字列を引く
     * @param {string} labelPath - "alert.noSelection" のようなドット区切りのキー
     * @returns {string} 表示言語のテキスト（見つからない場合は labelPath をそのまま返す）
     */
    function getLabel(labelPath) {
        var labelPathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < labelPathKeys.length; i++) {
            labelNode = labelNode[labelPathKeys[i]];
            if (!labelNode) return labelPath;
        }
        return labelNode[uiLang] || labelNode["en"] || labelPath;
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択をグループ化し、線を塗りに変換してからライブエフェクトを適用し、アピアランスを分割する
     * @param {Document} doc - 対象ドキュメント
     * @param {string} liveEffectXml - 適用するライブエフェクトの XML
     * @param {boolean} ungroupAtEnd - 最後にグループを解除するか
     * @returns {void}
     */
    function applyLiveEffectAndExpand(doc, liveEffectXml, ungroupAtEnd) {
        app.executeMenuCommand('group');
        /* 線を塗りに変換 / Convert strokes to fills */
        app.executeMenuCommand('OffsetPath v22');
        var mergedGroup = doc.selection[0];
        mergedGroup.applyEffect(liveEffectXml);
        app.redraw();
        doc.selection = null;
        mergedGroup.selected = true;
        app.executeMenuCommand('expandStyle');
        if (ungroupAtEnd) {
            app.executeMenuCommand('ungroup');
        }
    }

    /**
     * ドキュメントと選択を確認して処理を実行する
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        if (!doc.selection || doc.selection.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        applyLiveEffectAndExpand(doc, PATHFINDER_MERGE_XML, UNGROUP_AT_END);
    }

    main();

})();
