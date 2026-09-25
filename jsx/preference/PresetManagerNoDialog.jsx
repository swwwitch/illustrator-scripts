#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

PresetManager の［プリセット1］と同じ環境設定一式を、ダイアログを表示せずにまとめて適用します。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerNoDialog.md

### Overview

Applies the same preferences as [Preset 1] in PresetManager at once, without showing a dialog.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerNoDialog.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "PresetManagerNoDialog";        /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.1.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2025-08-18";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-25";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/PresetManagerNoDialog.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/PresetManagerNoDialog.md"; /* README (English) */

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
        { key: "ArtboardBBColorRed", type: "real", value: 0.0 },                   /* ハイライトのカラー：ブラック / Highlight Color: Black */
        { key: "ArtboardBBColorGreen", type: "real", value: 0.0 },
        { key: "ArtboardBBColorBlue", type: "real", value: 0.0 },
        { key: "ArtboardBBWidth", type: "real", value: 2 },                        /* ストロークの幅 1〜4 / Stroke Width */
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
