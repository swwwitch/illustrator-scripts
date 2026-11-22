#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
SCRIPTMETA-BEGIN
Script-ID=com.swwwitch.illustrator.PresetManagerNoDialog
Meta-Default-language=ja
Name=決めておいた環境設定一式をダイアログなしで適用
Name[en]=Apply a fixed set of preferences without a dialog
Author=Masahiro Takano (@swwwitch)
Target-App=Illustrator
Version=1.2.0
Release-Date=2025-08-18
Tags=preferences,環境設定
X-com.swwwitch-Script-Name=PresetManagerNoDialog
X-com.swwwitch-Updated=2026-10-09
X-com.swwwitch-README-JA=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerNoDialog.md
X-com.swwwitch-README-EN=https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerNoDialog.md
X-com.swwwitch-Source-URL=https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/preference/PresetManagerNoDialog.jsx
Description-BEGIN
[ja]
PresetManager の［プリセット1］と同じ環境設定一式を、ダイアログを表示せずにまとめて適用します。アートボードの枠線は ArtboardDisplayPresetManagerPalette の［ライト］と同じです。
[en]
Applies the same preferences as \[Preset 1\] in PresetManager at once, without showing a dialog. The artboard border matches \[Light\] in ArtboardDisplayPresetManagerPalette.
Description-END
SCRIPTMETA-END
*/

// =========================================
// 基本情報 / Basic info
// =========================================
// スクリプト情報（再利用パーツ） / Script info (reusable)
/* 基本情報は先頭の SCRIPTMETA が元。実行時に読み、getScriptMetaValue("Version") のように直接読む（値は SCRIPTMETA を書き換え、helpers/sync-scriptmeta.py で同期）
   The basic info lives in the SCRIPTMETA block at the top; read its fields directly with getScriptMetaValue() */

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
var SCRIPT_NAME = getScriptMetaValue("X-com.swwwitch-Script-Name") || decodeURI(new File($.fileName).name).replace(/\.[^.]+$/, ""); /* スクリプト名（設定の保存名） / script name, the key for saved settings */
// スクリプト情報（再利用パーツ）ここまで / End of the reusable script info

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 書き込む環境設定。type: "bool"=真偽値 / "int"=整数 / "real"=実数 */
    /* Preferences to write; type: "bool" / "int" / "real" */
    var PREFERENCES = [
        /* 一般 / General */
        { key: "showRichToolTips", type: "bool", value: false },                   /* 詳細なツールヒントを表示 / Show Rich Tool Tips */
        { key: "Hello/ShowHomeScreenWS", type: "bool", value: false },             /* 「ホーム画面」を表示 / Show the Home Screen */
        { key: "Hello/NewDoc", type: "bool", value: true },                        /* 以前の「新規ドキュメント」インターフェイス / Legacy File > New */
        { key: "enablePrintBleedWidget", type: "bool", value: false },             /* 「裁ち落としを印刷」生成AIボタン / Print Bleed AI buttons */
        /* 選択範囲・アンカー表示 / Selection & Anchor Display */
        { key: "zoomToSelection", type: "bool", value: false },                    /* 選択範囲へズーム / Zoom to Selection */
        { key: "showLockIcon", type: "bool", value: false },                       /* カンバス上でロック解除 / Unlock on Canvas */
        { key: "anchorSizePref", type: "int", value: 7 },                          /* アンカーポイントのサイズ 5/7/9/11 / Anchor Point Size */
        { key: "hitShapeOnPreview", type: "int", value: 1 },                       /* オブジェクトの選択範囲をパスに制限：OFF（0がON）/ Object Selection by Path Only: off (0 = on) */
        { key: "hitTypeShapeOnPreview", type: "int", value: 1 },                   /* テキストオブジェクトの選択範囲をパスに制限：OFF（0がON）/ Type Object Selection by Path Only: off (0 = on) */
        /* アートボード / Artboard */
        { key: "moveLockedAndHiddenArt", type: "bool", value: true },              /* ロックまたは非表示オブジェクトを一緒に移動 / Move Locked and Hidden Artwork */
        { key: "showArtboardLabelOnCanvas", type: "bool", value: false },          /* アートボード名を表示 / Show Artboard Name */
        { key: "ArtboardBBColorRed", type: "real", value: 0.65 },                  /* ハイライトのカラー：ライトグレー / Highlight Color: Light Gray */
        { key: "ArtboardBBColorGreen", type: "real", value: 0.65 },
        { key: "ArtboardBBColorBlue", type: "real", value: 0.65 },
        { key: "ArtboardBBWidth", type: "real", value: 1 },                        /* ストロークの幅 1〜4 / Stroke Width */
        /* テキスト / Text */
        { key: "text/autoSizing", type: "bool", value: true },                     /* 新規エリア内文字の自動サイズ調整 / Auto Size New Area Type */
        { key: "text/recentFontMenu/showNEntries", type: "int", value: 15 },       /* 最近使用したフォントの表示数 1〜30（0で非表示）/ Number of Recent Fonts (0 hides the list) */
        { key: "text/doFontLocking", type: "bool", value: false },                 /* 見つからない字形の保護 / Enable Missing Glyph Protection */
        { key: "text/enableAlternateGlyph", type: "bool", value: false },          /* 選択された文字の異体字を表示 / Show Character Alternates */
        /* ユーザーインターフェイス / User Interface */
        { key: "uiCanvasIsWhite", type: "int", value: 1 },                         /* カンバスカラー：ホワイト / Canvas Color: White */
        /* ガイド / Guides */
        { key: "Guide/Color/red", type: "real", value: 0.29 },                     /* ガイドのカラー：ライトブルー / Guide Color: Light Blue */
        { key: "Guide/Color/green", type: "real", value: 0.52 },
        { key: "Guide/Color/blue", type: "real", value: 1.0 },
        { key: "Guide/Style", type: "int", value: 0 },                             /* ガイドのスタイル 0=ライン・1=ポイント / Guide Style 0=lines, 1=dots */
        /* スマートガイド / Smart Guides */
        { key: "smartGuides/showObjectHighlighting", type: "int", value: 0 },      /* オブジェクトのハイライト表示 / Object Highlighting */
        /* パフォーマンス / Performance */
        { key: "Performance/AnimZoom", type: "bool", value: false },               /* アニメーションズーム / Animated Zoom */
        { key: "maximumUndoDepth", type: "int", value: 50 },                       /* ヒストリー数 / History States */
        { key: "LiveEdit_State_Machine", type: "bool", value: false },             /* リアルタイムの描画と編集 / Real-Time Drawing and Editing */
        /* ファイル管理 / File Management */
        { key: "useSysDefEdit", type: "bool", value: true },                       /* 「オリジナルの編集」にシステムデフォルトを使用 / Use System Defaults */
        { key: "AutoActivateMissingFont", type: "bool", value: true },             /* Adobe Fonts を自動アクティベート / Auto-activate Adobe Fonts */
        { key: "AdobeSaveAsCloudDocumentPreference", type: "bool", value: false }, /* ファイルの保存先：コンピューター / Save Location: computer */
        { key: "plugin/FileClipboard/linkoptions", type: "int", value: 0 },        /* リンクを更新 0=自動・1=手動・2=確認 / Update Links 0=auto, 1=manual, 2=ask */
        /* クリップボードの処理 / Clipboard Handling */
        { key: "plugin/FileClipboard/copySVGCode", type: "bool", value: true }     /* SVGコードを含める / Include SVG Code */
    ];

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * 環境設定を1件書き込む。存在しないキーなどで失敗しても処理は止めない
     * Write one preference; a failure (e.g. a key missing in this version) is logged and skipped
     * @param {{key: string, type: string, value: *}} preference - PREFERENCES の1行
     * @returns {void}
     */
    function writePreference(preference) {
        try {
            if (preference.type === "bool") {
                app.preferences.setBooleanPreference(preference.key, preference.value);
            } else if (preference.type === "int") {
                app.preferences.setIntegerPreference(preference.key, preference.value);
            } else {
                app.preferences.setRealPreference(preference.key, preference.value);
            }
        } catch (e) {
            $.writeln("[" + SCRIPT_NAME + "] " + preference.key + " = " + preference.value + " / " + e);
        }
    }

    /**
     * 環境設定をまとめて書き込み、画面を更新する / Write every preference and refresh the screen
     * 反映には強制再描画が要り、redraw だけでは足りないためズームで描き直す
     * A forced repaint is needed and redraw alone is unreliable, so toggle the zoom
     * @returns {void}
     */
    function main() {
        for (var i = 0; i < PREFERENCES.length; i++) {
            writePreference(PREFERENCES[i]);
        }
        if (app.documents.length > 0) {
            app.executeMenuCommand("zoomout");
            app.executeMenuCommand("zoomin");
        }
    }

    main();
})();
