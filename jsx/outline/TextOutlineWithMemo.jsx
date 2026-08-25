#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*

### 概要

選択したテキストの文字・段落属性をオブジェクトのメモ（note）へ保存してから、アウトライン化します。

詳細は README を参照してください。

### Overview

Saves the character and paragraph attributes of the selected text into the object's note, then outlines it.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "TextOutlineWithMemo";          /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2";                         /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2024-07-23";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-01-11";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/TextOutlineWithMemo.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/TextOutlineWithMemo.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/n3e0f241508db"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    var LOCALE = (app.locale && app.locale.indexOf('ja') === 0) ? 'ja' : 'en';

    var I18N = {
        ja: {
            ERR_NO_DOC: 'ドキュメントが開かれていません。',
            ERR_NO_SELECTION: 'テキストオブジェクトが選択されていません。'
        },
        en: {
            ERR_NO_DOC: 'No document is open.',
            ERR_NO_SELECTION: 'No text objects are selected.'
        }
    };

    function _(key) {
        return (I18N[LOCALE] && I18N[LOCALE][key]) ? I18N[LOCALE][key] : key;
    }

    function main() {
        // ドキュメントが開かれていることを確認
        if (app.documents.length > 0) {
            var doc = app.activeDocument;
            var selection = doc.selection;

            // 選択されたオブジェクトが存在するか確認
            if (selection.length > 0) {
                // 元の選択を保存
                var originalSelection = selection.slice();

                for (var i = 0; i < originalSelection.length; i++) {
                    if (originalSelection[i].typename == "TextFrame") {
                        processTextFrame(originalSelection[i]);
                    }
                }
            } else {
                alert(_("ERR_NO_SELECTION"));
            }
        } else {
            alert(_("ERR_NO_DOC"));
        }
    }

    function processTextFrame(textFrame) {
        var textRange = textFrame.textRange;

        // テキスト情報を取得
        var content = textRange.contents;
        var fontName = textRange.characterAttributes.textFont.name;
        var fontSize = roundToTwoDecimals(textRange.characterAttributes.size);
        var leadingValue = roundToTwoDecimals(textRange.characterAttributes.leading);
        var kerningMethod = textRange.characterAttributes.kerningMethod;
        var proportionalMetrics = textRange.characterAttributes.proportionalMetrics;
        var trackingValue = textRange.characterAttributes.tracking;
        var orientation = textFrame.orientation == TextOrientation.VERTICAL ? "縦組み" : "横組み";

        // カーニングメソッドの文字列表現を取得
        var kerningMethodText;
        switch (kerningMethod) {
            case AutoKernType.AUTO:
                kerningMethodText = "メトリクス";
                break;
            case AutoKernType.METRICSROMANONLY:
                kerningMethodText = "和文等幅";
                break;
            case AutoKernType.OPTICAL:
                kerningMethodText = "オプティカル";
                break;
            default:
                kerningMethodText = "なし";
        }

        // プロポーショナルメトリクスの文字列表現を取得
        var proportionalMetricsText = proportionalMetrics ? "true" : "false";

        // 座標を取得（見た目＝geometricBounds 基準）
        // ※フォント/サイズ等が確定した後の bounds を使うのがポイント
        app.redraw(); // bounds が更新されない環境対策（不要な場合もあります）

        var bounds = textFrame.geometricBounds; // [left, top, right, bottom]
        var x = roundToTwoDecimals(bounds[0]);
        var y = roundToTwoDecimals(bounds[1]);
        var r = roundToTwoDecimals(bounds[2]);
        var b = roundToTwoDecimals(bounds[3]);

        // メモ用のテキストを作成
        var memoText = "文字列：\n" + content + "\n\n" +
                       "フォント：\n" + fontName + "\n\n" +
                       "フォントサイズ：\n" + fontSize + "\n\n" +
                       "行送り：\n" + leadingValue + "\n\n" +
                       "カーニング：\n" + kerningMethodText + "\n\n" +
                       "プロポーショナルメトリクス：\n" + proportionalMetricsText + "\n\n" +
                       "トラッキング：\n" + trackingValue + "\n\n" +
                       "組み方向：\n" + orientation + "\n\n" +
                       "座標（geometricBounds）：\nL = " + x + ", T = " + y + ", R = " + r + ", B = " + b;

        // 元の選択をクリアし、現在のテキストフレームを選択
        app.activeDocument.selection = null;
        textFrame.selected = true;

        // テキストをアウトライン化
        textFrame.createOutline();

        // アウトライン化したオブジェクトを取得
        var outlinedObject = app.activeDocument.selection[0];

        // メモに情報を入力
        outlinedObject.note = memoText;
    }

    function roundToTwoDecimals(value) {
        return Math.round(value * 100) / 100;
    }

    // スクリプトを実行

    main();

})();
