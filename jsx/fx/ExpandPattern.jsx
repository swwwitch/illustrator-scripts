#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択オブジェクトをグループ化し、3D 回転（クラシック・前面 0/0/0）と［合流］の効果を掛けてからアピアランスを分割し、パターンなどの塗りを通常のパスに展開します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandPattern.md

### Overview

Groups the selected objects, applies 3D Rotate (Classic, Front 0/0/0) and Pathfinder Merge effects, then expands the appearance to turn pattern fills into regular paths.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandPattern.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ExpandPattern";                /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-04";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-04";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ExpandPattern.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ExpandPattern.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

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
            noSelection: { ja: "展開するオブジェクトを選択してください。", en: "Please select the objects to expand." }
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
     * 3D効果（クラシック）回転を「前面」（回転 X/Y/Z すべて 0°）で適用する LiveEffect XML を返す。
     * パラメーターは mark1bean/live-effect-functions-for-illustrator の LE_3DEffect の既定値に準拠。
     * @returns {string} applyEffect() に渡すXML
     */
    function build3DFrontEffectXML() {
        /* 回転 0/0/0 なので回転行列は単位行列 / Rotation 0/0/0, so the matrix is the identity */
        var rotationMatrix = [
            "R mat_00 1", "R mat_01 0", "R mat_02 0", "R mat_03 0",
            "R mat_10 0", "R mat_11 1", "R mat_12 0", "R mat_13 0",
            "R mat_20 0", "R mat_21 0", "R mat_22 1", "R mat_23 0",
            "R mat_30 0", "R mat_31 0", "R mat_32 0", "R mat_33 1"
        ].join(" ");
        var effectParams = [
            rotationMatrix,
            "I effectStyle 2",          /* 2 = 回転 / Rotate */
            "R rotX 0", "R rotY 0", "R rotZ 0",
            "R cameraPerspective 0",
            "R extrudeDepth 100",
            "R surfaceAmbient 70",
            "I shadeMode 2",
            "I surfaceStyle 1",         /* 1 = 陰影なし / No shading */
            "R surfaceMatte 40",
            "R surfaceGloss 10",
            "R blendSteps 25",
            "B preserveSpots 0",
            "B extrudeCap 1",
            "R revolveAngle 360",
            "R revolveOffset 0",
            "B revolveCap 1",
            "I revolveAxisMode 0",
            "R bevelHeight 4",
            "B bevelExtentIn 1",
            "B shadeMaps 0",
            "B showHiddenSurfaces 0",
            "B invisibleGeo 0",
            "I numArtMaps 0",
            "I 3Dversion 2",
            "B paramsDictionaryInitialized 1",
            "I numLights 1"
        ].join(" ");
        var lightParams = "R lightIntensity 1 R lightDirX -0.99 R lightDirY 0 R lightDirZ -1 R lightPosX 0 R lightPosY 0 R lightPosZ -1";
        return [
            '<LiveEffect name="Adobe 3D Effect"><Dict data="', effectParams, ' ">',
            '<Entry name="shadeColor" valueType="F"><Fill color="5 0 0 0"/></Entry>',
            '<Entry name="light0" valueType="D"><Dict data="', lightParams, ' " /></Entry>',
            '</Dict></LiveEffect>'
        ].join("");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 選択をグループ化し、3D 回転と［合流］を掛けてアピアランスを分割する
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

        app.executeMenuCommand("group"); /* 選択オブジェクトをグループ化 / Group the selection */
        var expandGroup = targetDoc.selection[0];
        expandGroup.name = getLabel("fallbackName.expandGroup");

        /* applyEffect() は後から掛けた効果ほどパネルの下に並ぶので、上に置く 3D を先に掛ける
           Effects applied later sit lower in the Appearance panel, so apply 3D first to keep it on top */
        expandGroup.applyEffect(build3DFrontEffectXML());
        expandGroup.applyEffect(buildMergeEffectXML());

        /* 効果の適用を確定させてから分割する（redraw しないと分割が効かない）
           Commit the effects before expanding; without redraw the expand has no effect */
        targetDoc.selection = [expandGroup];
        app.redraw();
        app.executeMenuCommand("expandStyle"); /* アピアランスを分割 / Expand Appearance */
    }

    main();

})();
