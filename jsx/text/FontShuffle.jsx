#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストの1文字ごとに、ランダムなフォントを適用します。
ダイアログでプレビューを確認しながら、［再実行］で抽選し直せます。

詳細は README を参照してください。
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontShuffle.md

### Overview

Applies a random font to each character of the selected text.
The dialog previews the result, and a Reshuffle button redraws the assignment.

See the README for details.
https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontShuffle.md

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "FontShuffle";                  /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.3";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-02-16";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-22";                   /* 更新日 / last updated */

var SCRIPT_README_JA = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FontShuffle.md"; /* README（日本語） */
var SCRIPT_README_EN = "https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/FontShuffle.md"; /* README (English) */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var MANIFESTO_TRACKING = 200;  /* 犯行声明文風の字送り（1/1000 em） / tracking for the manifesto style (1/1000 em) */
    var BG_RECT_PADDING    = 1;    /* 背景長方形を文字の外へ広げる量（pt） / padding around each character (pt) */
    var BG_RECT_JITTER     = 1;    /* 背景長方形の角を外へずらす最大量（pt） / max outward jitter of each corner (pt) */
    var BG_GRAY_MIN        = 10;   /* 背景のグレーの最小値（K%） / lightest background gray (K%) */
    var BG_GRAY_MAX        = 50;   /* 背景のグレーの最大値（K%） / darkest background gray (K%) */

    /* 和文フォントとみなす名前のキーワード / Name keywords treated as Japanese fonts */
    var JP_FONT_KEYWORDS = ["Pr6N", "Pr6", "AB-", "FOT", "-OTF"];

    /* 背景長方形を置くレイヤーとグループの名前 / Layer and group names for the background rectangles */
    var BG_LAYER_NAME = "__FontShuffle_BG__";
    var BG_GROUP_NAME = "__BGRects__";

    /* ダイアログ位置を覚える $.global のキー / $.global key that remembers the dialog position */
    var DIALOG_KEY = "__FontShuffle_Dialog__";

    // =========================================
    // レイアウト / Layout
    // =========================================
    var DIALOG_OPACITY      = 0.98;  /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_OFFSET_X     = 300;   /* 初回表示の横オフセット / initial horizontal offset */
    var DIALOG_OFFSET_Y     = 0;     /* 初回表示の縦オフセット / initial vertical offset */
    var INFO_CHARACTERS     = 52;    /* 説明文の幅（文字数） / width of the description (characters) */
    var BUTTON_SPACER_WIDTH = 20;    /* ボタン行の中央スペーサーの幅 / width of the button-row spacer */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /**
     * Illustrator の UI 言語から表示言語を判定する
     * @returns {string} "ja" または "en"
     */
    function detectUILanguage() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }

    var uiLang = detectUILanguage();

    /* 日英ラベル定義 / Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "フォントシャッフル", en: "Font Shuffle" },
            description: {
                ja: "選択したテキストに対して、1文字ごとにランダムなフォントを適用します。",
                en: "Applies a random font to each character in the selected text."
            }
        },
        checkbox: {
            limitToJPFonts: { ja: "和文フォントに限る", en: "Limit to JP fonts" },
            manifestoStyle: { ja: "犯行声明文風", en: "Manifesto style" }
        },
        tooltip: {
            limitToJPFonts: {
                ja: "名前に Pr6／Pr6N／AB-／FOT／-OTF を含むフォントだけから選びます。英数字以外の文字を選択していると最初からオンになります。",
                en: "Picks only fonts whose names contain Pr6, Pr6N, AB-, FOT or -OTF. Starts on when the selected text contains anything other than letters and digits."
            },
            manifestoStyle: {
                ja: "字送りを 200 にし、1文字ごとにランダムなグレー（K10〜K50）の長方形を背面に敷きます。",
                en: "Sets tracking to 200 and places a random gray (K10–K50) rectangle behind each character."
            },
            rerun: { ja: "フォントを抽選し直してプレビューします。", en: "Reshuffles the fonts and updates the preview." }
        },
        button: {
            rerun: { ja: "再実行", en: "Rerun" },
            ok: { ja: "OK", en: "OK" },
            cancel: { ja: "キャンセル", en: "Cancel" }
        },
        alert: {
            noDocument: { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noTextSelected: { ja: "テキストオブジェクトを選択してください。", en: "Please select at least one text object." },
            noJPFonts: { ja: "対象の和文フォント（Pr6 / Pr6N）が見つかりません。", en: "No target JP fonts (Pr6 / Pr6N) were found." }
        },
        history: {
            preview: { ja: "FontShuffle プレビュー", en: "FontShuffle Preview" }
        }
    };

    /**
     * ドット区切りのパスで表示言語のラベルを取得する
     * @param {string} labelPath - "button.ok" のようなドット区切りキー
     * @returns {string} 表示言語のラベル。見つからない場合はパスをそのまま返す
     */
    function getLabel(labelPath) {
        var pathKeys = labelPath.split(".");
        var labelNode = LABELS;
        for (var i = 0; i < pathKeys.length; i++) {
            labelNode = labelNode[pathKeys[i]];
            if (!labelNode) return labelPath;
        }
        if (labelNode[uiLang] != null) return labelNode[uiLang];
        return (labelNode.ja != null) ? labelNode.ja : labelPath;
    }

    // =========================================
    // 共通ユーティリティ / Shared utilities
    // =========================================
    // 他のスクリプトと同じ実装を $.global に載せて共有するため、名前と中身はそろえたまま
    // Shared with other scripts through $.global, so names and bodies are kept identical

    /* DialogPersist util (extractable)
     * ダイアログの不透明度・初期位置・位置記憶を共通化
     * 使い方:
     *   DialogPersist.setOpacity(dlg, 0.95);
     *   DialogPersist.restorePosition(dlg, "__YourDialogKey", offsetX, offsetY);
     *   DialogPersist.rememberOnMove(dlg, "__YourDialogKey");
     *   DialogPersist.savePosition(dlg, "__YourDialogKey");
     */
    (function(g){
      if (!g.DialogPersist) {
        g.DialogPersist = {
          setOpacity: function(dlg, v){ try{ dlg.opacity=v; }catch(e){} },
          _getSaved: function(key){ return g[key] && g[key].length===2 ? g[key] : null; },
          _setSaved: function(key, loc){ g[key] = [loc[0], loc[1]]; },
          _clampToScreen: function(loc){
            try{
              var vb = ($.screens && $.screens.length) ? $.screens[0].visibleBounds : [0,0,1920,1080];
              var x = Math.max(vb[0]+10, Math.min(loc[0], vb[2]-10));
              var y = Math.max(vb[1]+10, Math.min(loc[1], vb[3]-10));
              return [x,y];
            }catch(e){ return loc; }
          },
          restorePosition: function(dlg, key, offsetX, offsetY){
            var loc = this._getSaved(key);
            try{
              if (loc) dlg.location = this._clampToScreen(loc);
              else { var l = dlg.location; dlg.location = [l[0]+(offsetX|0), l[1]+(offsetY|0)]; }
            }catch(e){}
          },
          rememberOnMove: function(dlg, key){
            var self = this;
            dlg.onMove = function(){
              try{ self._setSaved(key, [dlg.location[0], dlg.location[1]]); }catch(e){}
            };
          },
          savePosition: function(dlg, key){
            try{ this._setSaved(key, [dlg.location[0], dlg.location[1]]); }catch(e){}
          }
        };
      }
    })($.global);

    /* =========================================
     * PreviewHistory util (extractable)
     * ヒストリーを残さないプレビューのための小さなユーティリティ。
     * 他スクリプトでもこのブロックをコピペすれば再利用できます。
     * 使い方:
     *   PreviewHistory.start();     // ダイアログ表示時などにカウンタ初期化
     *   PreviewHistory.bump();      // プレビュー描画ごとにカウント(+1)
     *   PreviewHistory.undo();      // 閉じる/キャンセル時に一括Undo
     *   PreviewHistory.cancelTask(t);// app.scheduleTaskのキャンセル補助
     * ========================================= */
    (function(g){
        if (!g.PreviewHistory) {
            g.PreviewHistory = {
                start: function(){ g.__previewUndoCount = 0; },
                bump:  function(){ g.__previewUndoCount = (g.__previewUndoCount | 0) + 1; },
                undo:  function(){
                    var n = g.__previewUndoCount | 0;
                    try { for (var i = 0; i < n; i++) app.executeMenuCommand('undo'); } catch (e) {}
                    g.__previewUndoCount = 0;
                },
                cancelTask: function(taskId){
                    try { if (taskId) app.cancelTask(taskId); } catch (e) {}
                }
            };
        }
    })($.global);

    // =========================================
    // 背景長方形 / Background rectangles
    // =========================================
    // 1文字ごとの位置はアウトライン化した複製から測る / Per-character bounds come from an outlined duplicate

    /**
     * 背景用レイヤーを探す
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer|null} 見つかったレイヤー。無ければ null
     */
    function findBgLayer(doc) {
        try {
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i] && doc.layers[i].name === BG_LAYER_NAME) return doc.layers[i];
            }
        } catch (e) { }
        return null;
    }

    /**
     * 背景用レイヤーを返す（無ければ作る）
     * @param {Document} doc - 対象ドキュメント
     * @returns {Layer} 背景用レイヤー
     */
    function ensureBgLayer(doc) {
        var bgLayer = findBgLayer(doc);
        if (!bgLayer) {
            bgLayer = doc.layers.add();
            bgLayer.name = BG_LAYER_NAME;
        }
        return bgLayer;
    }

    /**
     * 背景用レイヤーから背景長方形のグループを削除する
     * @param {Layer} bgLayer - 背景用レイヤー
     * @returns {void}
     */
    function removeBgGroups(bgLayer) {
        try {
            for (var i = bgLayer.groupItems.length - 1; i >= 0; i--) {
                if (bgLayer.groupItems[i].name === BG_GROUP_NAME) {
                    bgLayer.groupItems[i].remove();
                }
            }
        } catch (e) { }
    }

    /**
     * 背景長方形があれば削除する
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function clearBackgroundRects(doc) {
        var bgLayer = findBgLayer(doc);
        if (!bgLayer) return;
        removeBgGroups(bgLayer);
    }

    /**
     * min 以上 max 以下の整数を乱数で返す
     * @param {number} min - 下限
     * @param {number} max - 上限
     * @returns {number} 乱数の整数
     */
    function randomInt(min, max) {
        return Math.floor(Math.random() * (max - min + 1)) + min;
    }

    /**
     * 背景長方形を塗りだけのランダムなグレーにする
     * @param {PathItem} bgRect - 対象の長方形
     * @returns {void}
     */
    function applyRandomGrayFill(bgRect) {
        try {
            bgRect.stroked = false;
            bgRect.filled = true;
            /* GrayColor.gray は 0 が白、100 が黒 / GrayColor.gray: 0 = white, 100 = black */
            var grayColor = new GrayColor();
            grayColor.gray = randomInt(BG_GRAY_MIN, BG_GRAY_MAX);
            bgRect.fillColor = grayColor;
        } catch (e) { }
    }

    /**
     * 長方形の各アンカーを中心から外向きにランダムにずらす（手作り感を出す）
     * @param {PathItem} bgRect - 対象の長方形
     * @param {number} maxOffset - ずらす最大量（pt）
     * @returns {void}
     */
    function jitterRectOutward(bgRect, maxOffset) {
        if (!bgRect || bgRect.typename !== "PathItem") return;
        if (!bgRect.pathPoints || bgRect.pathPoints.length < 4) return;

        var rectBounds;
        try { rectBounds = bgRect.geometricBounds; } catch (e) { return; }
        /* rectBounds: [左, 上, 右, 下] / [L, T, R, B] */
        var centerX = (rectBounds[0] + rectBounds[2]) / 2;
        var centerY = (rectBounds[1] + rectBounds[3]) / 2;

        for (var i = 0; i < bgRect.pathPoints.length; i++) {
            try {
                var pathPoint = bgRect.pathPoints[i];
                var anchorX = pathPoint.anchor[0], anchorY = pathPoint.anchor[1];
                var dx = anchorX - centerX;
                var dy = anchorY - centerY;
                var distanceFromCenter = Math.sqrt(dx * dx + dy * dy);
                if (!distanceFromCenter) continue;

                /* 外向きのランダムな移動量 / random outward distance */
                var offsetDistance = Math.random() * maxOffset;
                var offsetX = dx / distanceFromCenter * offsetDistance;
                var offsetY = dy / distanceFromCenter * offsetDistance;

                /* 角の形を保つため、アンカーとハンドルを一緒に動かす / move anchor and handles together to keep the corner shape */
                pathPoint.anchor = [anchorX + offsetX, anchorY + offsetY];
                pathPoint.leftDirection = [pathPoint.leftDirection[0] + offsetX, pathPoint.leftDirection[1] + offsetY];
                pathPoint.rightDirection = [pathPoint.rightDirection[0] + offsetX, pathPoint.rightDirection[1] + offsetY];
            } catch (e) { }
        }
    }

    /**
     * 1文字の範囲に背景長方形を追加する
     * @param {GroupItem} bgGroup - 追加先のグループ
     * @param {number[]} charBounds - 文字の geometricBounds [左, 上, 右, 下]
     * @returns {PathItem|null} 作った長方形。幅か高さが 0 以下なら null
     */
    function addBackgroundRect(bgGroup, charBounds) {
        var left = charBounds[0] - BG_RECT_PADDING;
        var top = charBounds[1] + BG_RECT_PADDING;
        var right = charBounds[2] + BG_RECT_PADDING;
        var bottom = charBounds[3] - BG_RECT_PADDING;

        var rectWidth = right - left;
        var rectHeight = top - bottom;
        if (rectWidth <= 0 || rectHeight <= 0) return null;

        var bgRect = bgGroup.pathItems.rectangle(top, left, rectWidth, rectHeight);
        applyRandomGrayFill(bgRect);
        jitterRectOutward(bgRect, BG_RECT_JITTER);
        return bgRect;
    }

    /**
     * アウトライン化の結果から、1文字ずつのまとまりを集める
     * 典型的な構造は「outlinedGroup > 行・ブロックのグループ > 1文字のグループ」なので、
     * 2段目の groupItems を1文字として扱い、それが無い構造にはフォールバックする
     * @param {PageItem} outlinedGroup - createOutline() の結果
     * @param {PageItem[]} charUnits - 収集先の配列
     * @returns {void}
     */
    function collectCharGroupsFromOutlined(outlinedGroup, charUnits) {
        if (!outlinedGroup) return;

        /**
         * グループをたどって末端のパスを集める
         * @param {GroupItem} containerItem - たどるグループ
         * @returns {void}
         */
        function pushLeafUnits(containerItem) {
            try {
                if (!containerItem || !containerItem.pageItems) return;
                for (var k = 0; k < containerItem.pageItems.length; k++) {
                    var childItem = containerItem.pageItems[k];
                    if (!childItem) continue;
                    if (childItem.typename === "GroupItem") {
                        pushLeafUnits(childItem);
                    } else if (childItem.typename === "PathItem" || childItem.typename === "CompoundPathItem") {
                        charUnits.push(childItem);
                    }
                }
            } catch (e) { }
        }

        try {
            if (outlinedGroup.typename === "GroupItem") {
                var hasPushed = false;

                /* 2段目の groupItems を優先する / Prefer the second-level groupItems */
                if (outlinedGroup.groupItems && outlinedGroup.groupItems.length > 0) {
                    for (var i = 0; i < outlinedGroup.groupItems.length; i++) {
                        var lineGroup = outlinedGroup.groupItems[i];
                        if (!lineGroup || lineGroup.typename !== "GroupItem") continue;

                        if (lineGroup.groupItems && lineGroup.groupItems.length > 0) {
                            for (var j = 0; j < lineGroup.groupItems.length; j++) {
                                charUnits.push(lineGroup.groupItems[j]);
                                hasPushed = true;
                            }
                        } else {
                            /* その下にグループが無ければ、それ自体を1単位にする / Use the child itself when it has no further groups */
                            charUnits.push(lineGroup);
                            hasPushed = true;
                        }
                    }
                }

                if (hasPushed) return;

                /* フォールバック：末端のパスを単位にして、1つにまとまらないようにする / Fallback: leaf paths, so it does not collapse into one unit */
                pushLeafUnits(outlinedGroup);
                if (charUnits.length > 0) return;

                /* 最後の手段：アウトライン全体 / Last resort: the whole outline */
                charUnits.push(outlinedGroup);
                return;
            }

            charUnits.push(outlinedGroup);
        } catch (e) { }
    }

    /**
     * テキストフレームを背面に複製してアウトライン化する（元のテキストはそのまま）
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @returns {PageItem|null} アウトラインの結果。複製・アウトライン化に失敗したら null
     */
    function outlineDuplicateOf(textFrame) {
        var textDuplicate;
        try {
            textDuplicate = textFrame.duplicate();
            try { textDuplicate.move(textFrame, ElementPlacement.PLACEAFTER); } catch (e) { }
        } catch (e) {
            return null;
        }

        var outlinedGroup;
        try {
            outlinedGroup = textDuplicate.createOutline();
        } catch (e) {
            try { textDuplicate.remove(); } catch (eRemove) { }
            return null;
        }

        /* createOutline() は複製を消費するので、通常ここは例外になる / createOutline() consumes the duplicate, so this normally throws */
        try { textDuplicate.remove(); } catch (e) { }
        return outlinedGroup;
    }

    /**
     * 1文字ごとの背景長方形を作り直す（アウトライン化した複製で文字の位置を測る）
     * @param {Document} doc - 対象ドキュメント
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {void}
     */
    function createBackgroundRectsByOutlining(doc, textFrames) {
        if (!textFrames || textFrames.length === 0) return;

        var bgLayer = ensureBgLayer(doc);
        removeBgGroups(bgLayer);
        var bgGroup = bgLayer.groupItems.add();
        bgGroup.name = BG_GROUP_NAME;

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!textFrame || textFrame.typename !== "TextFrame") continue;

            var outlinedGroup = outlineDuplicateOf(textFrame);
            if (!outlinedGroup) continue;

            var charUnits = [];
            collectCharGroupsFromOutlined(outlinedGroup, charUnits);

            for (var j = 0; j < charUnits.length; j++) {
                try {
                    addBackgroundRect(bgGroup, charUnits[j].geometricBounds);
                } catch (e) { }
            }

            /* 測り終えたアウトラインは消す / Remove the outline once measured */
            try { outlinedGroup.remove(); } catch (e) { }
        }

        /* 背景は内容より背面に置く / Keep backgrounds behind content */
        try { bgGroup.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
        try { bgLayer.zOrder(ZOrderMethod.SENDTOBACK); } catch (e) { }
    }

    // =========================================
    // フォントの適用 / Font assignment
    // =========================================

    /**
     * 名前に和文フォントのキーワードを含むか判定する
     * @param {string} fontName - フォント名
     * @returns {boolean} 含んでいれば true
     */
    function hasJPFontKeyword(fontName) {
        if (!fontName) return false;
        for (var i = 0; i < JP_FONT_KEYWORDS.length; i++) {
            if (fontName.indexOf(JP_FONT_KEYWORDS[i]) !== -1) return true;
        }
        return false;
    }

    /**
     * インストール済みフォントから、名前で和文フォントと判定できるものを集める
     * @param {TextFonts} allFonts - app.textFonts
     * @returns {TextFont[]} 和文フォントの配列
     */
    function getJPFonts(allFonts) {
        var jpFonts = [];
        for (var i = 0; i < allFonts.length; i++) {
            try {
                var textFont = allFonts[i];
                if (!textFont) continue;

                var fullName = textFont.fullName ? String(textFont.fullName) : "";
                var familyName = textFont.name ? String(textFont.name) : "";
                var postScriptName = textFont.postScriptName ? String(textFont.postScriptName) : "";

                if (hasJPFontKeyword(fullName) || hasJPFontKeyword(familyName) || hasJPFontKeyword(postScriptName)) {
                    jpFonts.push(textFont);
                }
            } catch (e) { }
        }
        return jpFonts;
    }

    /**
     * 選択からテキストフレームだけを取り出す（グループの中はたどらない）
     * @param {Document} doc - 対象ドキュメント
     * @returns {TextFrame[]} 選択中のテキストフレーム
     */
    function getSelectedTextFrames(doc) {
        var currentSelection = doc.selection;
        if (!currentSelection || currentSelection.length === 0) return [];
        var textFrames = [];
        for (var i = 0; i < currentSelection.length; i++) {
            try {
                if (currentSelection[i] && currentSelection[i].typename === "TextFrame") textFrames.push(currentSelection[i]);
            } catch (e) { }
        }
        return textFrames;
    }

    /**
     * テキストに ASCII 英数字と改行以外の文字が含まれるか判定する
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @returns {boolean} 含まれていれば true
     */
    function containsNonAlphanumeric(textFrames) {
        if (!textFrames || textFrames.length === 0) return false;

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!textFrame || textFrame.typename !== "TextFrame") continue;

            var frameText;
            try { frameText = textFrame.contents; } catch (e) { continue; }
            for (var j = 0; j < frameText.length; j++) {
                var character = frameText.charAt(j);
                if (/[A-Za-z0-9]/.test(character)) continue;
                /* 改行は除外 / ignore line breaks */
                if (character === "\r" || character === "\n") continue;
                return true;
            }
        }
        return false;
    }

    /**
     * 各文字にランダムなフォントを適用する（改行と半角スペースは飛ばす）
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {TextFont[]|TextFonts} fontList - 抽選するフォント
     * @returns {void}
     */
    function applyRandomFontsToFrames(textFrames, fontList) {
        if (!textFrames || textFrames.length === 0) return;
        if (!fontList || fontList.length === 0) return;

        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!textFrame || textFrame.typename !== "TextFrame") continue;

            var characters = textFrame.textRange.characters;
            for (var j = 0; j < characters.length; j++) {
                /* 改行や空白文字はスキップ（エラー回避と見た目のため） / skip returns and spaces */
                if (characters[j].contents === "\r" || characters[j].contents === " ") continue;

                try {
                    var randomFontIndex = Math.floor(Math.random() * fontList.length);
                    characters[j].characterAttributes.textFont = fontList[randomFontIndex];
                } catch (e) {
                    /* 特定のフォントが適用できない場合は無視 / ignore fonts that cannot be applied */
                }
            }
        }

        app.redraw();
    }

    /**
     * テキストフレーム全体の字送りを設定する
     * @param {TextFrame[]} textFrames - 対象のテキストフレーム
     * @param {number} trackingValue - 字送り（1/1000 em）
     * @returns {void}
     */
    function setTrackingForFrames(textFrames, trackingValue) {
        if (!textFrames || textFrames.length === 0) return;
        for (var i = 0; i < textFrames.length; i++) {
            var textFrame = textFrames[i];
            if (!textFrame || textFrame.typename !== "TextFrame") continue;
            try {
                textFrame.textRange.characterAttributes.tracking = trackingValue;
            } catch (e) { }
        }
    }

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * ダイアログを組み立てる
     * @returns {{dialogWindow: Window, limitToJPFontsCheckbox: Checkbox, manifestoCheckbox: Checkbox, btnRerun: Button, btnCancel: Button, btnOK: Button}} ダイアログと操作するコントロール
     */
    function buildShuffleDialog() {
        var shuffleDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        DialogPersist.setOpacity(shuffleDialog, DIALOG_OPACITY);
        shuffleDialog.onShow = function () {
            DialogPersist.restorePosition(shuffleDialog, DIALOG_KEY, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        };
        DialogPersist.rememberOnMove(shuffleDialog, DIALOG_KEY);
        shuffleDialog.orientation = "column";
        shuffleDialog.alignChildren = ["fill", "top"];

        var descriptionText = shuffleDialog.add("statictext", undefined, getLabel("dialog.description"));
        descriptionText.characters = INFO_CHARACTERS;

        var limitToJPFontsCheckbox = shuffleDialog.add("checkbox", undefined, getLabel("checkbox.limitToJPFonts"));
        limitToJPFontsCheckbox.helpTip = getLabel("tooltip.limitToJPFonts");
        limitToJPFontsCheckbox.value = false;
        var manifestoCheckbox = shuffleDialog.add("checkbox", undefined, getLabel("checkbox.manifestoStyle"));
        manifestoCheckbox.helpTip = getLabel("tooltip.manifestoStyle");
        manifestoCheckbox.value = false;

        /* 下部ボタン行（左・スペーサー・右） / Bottom bar (left, spacer, right) */
        var btnRowGroup = shuffleDialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.alignChildren = ["fill", "center"];
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.orientation = "row";
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnRerun = btnLeftGroup.add("button", undefined, getLabel("button.rerun"));
        btnRerun.helpTip = getLabel("tooltip.rerun");

        var spacer = btnRowGroup.add("group");
        spacer.orientation = "row";
        spacer.alignment = ["fill", "center"];
        spacer.add("statictext", undefined, "");
        spacer.preferredSize.width = BUTTON_SPACER_WIDTH;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.orientation = "row";
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.alignment = ["right", "center"];
        var btnCancel = btnRightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOK = btnRightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });

        return {
            dialogWindow: shuffleDialog,
            limitToJPFontsCheckbox: limitToJPFontsCheckbox,
            manifestoCheckbox: manifestoCheckbox,
            btnRerun: btnRerun,
            btnCancel: btnCancel,
            btnOK: btnOK
        };
    }

    /**
     * ダイアログを表示し、プレビュー・再抽選・確定・取り消しを受け持つ
     * @param {Document} doc - 対象ドキュメント
     * @returns {void}
     */
    function showShuffleDialog(doc) {
        /* インストールされている全フォント（数が多いと処理に時間がかかる場合がある） / all installed fonts (can be slow when there are many) */
        var allFonts = app.textFonts;
        var dialogControls = buildShuffleDialog();
        var shuffleDialog = dialogControls.dialogWindow;
        var limitToJPFontsCheckbox = dialogControls.limitToJPFontsCheckbox;
        var manifestoCheckbox = dialogControls.manifestoCheckbox;

        /* プレビュー済みかどうか。閉じるときは PreviewHistory で一括 Undo する / whether a preview is applied; undone in one go on close */
        var isPreviewApplied = false;

        /**
         * プレビューを描く（フォントの抽選・字送り・背景長方形）
         * @param {boolean} makeRects - 犯行声明文風のとき背景長方形を作るなら true
         * @param {boolean} [randomizeFonts] - フォントを抽選し直すなら true（既定 true）
         * @param {boolean} [countHistory] - PreviewHistory の取り消し段数に数えるなら true（既定 true）
         * @returns {void}
         */
        function runPreview(makeRects, randomizeFonts, countHistory) {
            if (typeof randomizeFonts === "undefined") randomizeFonts = true;
            if (typeof countHistory === "undefined") countHistory = true;

            /**
             * プレビューの各段階を適用する（suspendHistory から文字列で呼べるよう名前付きで定義）
             * @returns {void}
             */
            function applyPreviewSteps() {
                var textFrames = getSelectedTextFrames(doc);
                if (textFrames.length === 0) {
                    alert(getLabel("alert.noTextSelected"));
                    return;
                }

                /* 抽選し直すときだけ、直前のプレビューを undo で戻す / undo the previous preview only when reshuffling */
                if (randomizeFonts && isPreviewApplied) {
                    try { app.undo(); } catch (e) { }
                }

                /* 1) フォントの抽選（任意） / random fonts (optional) */
                if (randomizeFonts) {
                    var fontList = allFonts;
                    if (limitToJPFontsCheckbox.value) {
                        fontList = getJPFonts(allFonts);
                        if (fontList.length === 0) {
                            alert(getLabel("alert.noJPFonts"));
                            return;
                        }
                    }
                    applyRandomFontsToFrames(textFrames, fontList);
                }

                /* 2) 犯行声明文風：字送り（抽選しない場合も適用） / manifesto style: tracking, also without reshuffling */
                if (manifestoCheckbox.value) {
                    setTrackingForFrames(textFrames, MANIFESTO_TRACKING);
                }

                /* 3) 背景長方形：要求があり、かつ犯行声明文風がオンのときだけ / background rectangles only when requested and the style is on */
                if (makeRects && manifestoCheckbox.value) {
                    createBackgroundRectsByOutlining(doc, textFrames);
                } else {
                    /* オフなら背景を付けない（既存があれば消す） / otherwise remove any existing background */
                    clearBackgroundRects(doc);
                }

                isPreviewApplied = true;
            }

            /* 取り消し1回で戻せるよう1段にまとめたい（Illustrator には suspendHistory が無いので通常は下へ進む）
               Try a single history step (Illustrator has no suspendHistory, so this normally falls through) */
            try {
                if (doc && doc.suspendHistory) {
                    doc.suspendHistory(getLabel("history.preview"), "applyPreviewSteps()");
                    if (countHistory) PreviewHistory.bump();
                    return;
                }
            } catch (e) { }

            applyPreviewSteps();
            if (countHistory) PreviewHistory.bump();
        }

        /* 犯行声明文風の切り替えは、プレビュー済みなら抽選せずに背景の有無だけ更新 / toggling the style keeps the fonts and only updates the background */
        manifestoCheckbox.onClick = function () {
            if (!isPreviewApplied) return;
            runPreview(manifestoCheckbox.value, false);
        };

        dialogControls.btnRerun.onClick = function () {
            runPreview(false, true);
        };

        dialogControls.btnOK.onClick = function () {
            /* プレビューの履歴を消してから最終適用 / Clear preview history before final apply */
            PreviewHistory.undo();
            if (!isPreviewApplied) {
                /* プレビュー未実行ならここで1回適用（犯行声明文風のときだけ長方形も作る） / apply once when never previewed */
                runPreview(true, true, false);
            } else if (manifestoCheckbox.value) {
                /* プレビュー済み：フォントは維持し、背景長方形だけ付ける / keep the fonts and add the background rectangles */
                runPreview(true, false, false);
            } else {
                /* オフなら長方形なし（既存があれば消す） / no rectangles when the style is off */
                clearBackgroundRects(doc);
            }
            DialogPersist.savePosition(shuffleDialog, DIALOG_KEY);
            shuffleDialog.close(1);
        };

        dialogControls.btnCancel.onClick = function () {
            /* プレビューを戻す（履歴を残さない） / Undo all previews */
            PreviewHistory.undo();
            DialogPersist.savePosition(shuffleDialog, DIALOG_KEY);
            shuffleDialog.close(0);
        };

        /* 英数字以外を含むテキストなら、和文フォントに限る をオンにして開く / start with JP fonts only for non-alphanumeric text */
        if (containsNonAlphanumeric(getSelectedTextFrames(doc))) {
            limitToJPFontsCheckbox.value = true;
        }

        PreviewHistory.start();
        shuffleDialog.center();
        shuffleDialog.show();
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントを確認してダイアログを開く
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        showShuffleDialog(app.activeDocument);
    }

    main();

})();
