#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトをグループ化し、線を塗りに変換してから Pathfinder の「合流」をライブエフェクトとして適用し、アピアランスを分割します。

詳細は README を参照してください。

### Overview

Groups the selection, converts strokes to fills, applies Pathfinder Merge as a live effect, and then expands the appearance.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "MergeExpand";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/MergeExpand.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/MergeExpand.md

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

/* 処理の最後にグループ解除するか / Whether to ungroup at the end */
var UNGROUP_AT_END = false;

/* Pathfinder Merge ライブエフェクトの XML（command 8 = Merge 固定）
   Live effect XML for Pathfinder Merge (command 8, all other params at defaults) */
var PATHFINDER_MERGE_XML = '<LiveEffect name="Adobe Pathfinder" isPre="1">'
    + '<Dict data="I Command 8 B ConvertCustom 1 B ExtractUnpainted 1 R Mix 0.5 R Precision 10 B RemovePoints 1 R TrapAspect 1 B TrapConvertCustom 1 R TrapMaxTint 1 B TrapReverse 0 R TrapThickness 0.25 R TrapTint 0.4 R TrapTintTolerance 0.05">'
    + '<Entry name="DisplayString" value="Merge" valueType="S"/>'
    + '</Dict></LiveEffect>';

/* ====== ライブエフェクト適用＋アピアランス分割 / Apply live effect and expand ======
   選択をグループ化し、線を塗りに変換してから指定 XML のライブエフェクトを適用、アピアランスを分割、必要に応じてグループを解除
   Group selection, convert strokes to fills, apply the given live effect XML, expand appearance, then optionally ungroup */
function applyLiveEffectAndExpand(doc, liveEffectXml, ungroupAtEnd) {
    app.executeMenuCommand('group');
    /* 線を塗りに変換 / Convert strokes to fills */
    app.executeMenuCommand('OffsetPath v22');
    var group = doc.selection[0];
    group.applyEffect(liveEffectXml);
    app.redraw();
    doc.selection = null;
    group.selected = true;
    app.executeMenuCommand('expandStyle');
    if (ungroupAtEnd) {
        app.executeMenuCommand('ungroup');
    }
}

/* ====== メイン処理 / Main ====== */
function main() {
    if (app.documents.length === 0) {
        alert('ドキュメントを開いてください。\nPlease open a document.');
        return;
    }
    var doc = app.activeDocument;
    if (!doc.selection || doc.selection.length === 0) {
        alert('オブジェクトを選択してください。\nPlease select one or more objects.');
        return;
    }

    applyLiveEffectAndExpand(doc, PATHFINDER_MERGE_XML, UNGROUP_AT_END);
}

main();
