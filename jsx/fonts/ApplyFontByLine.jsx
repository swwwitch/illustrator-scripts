#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

選択したテキストフレームの各行（段落）を「フォント名（＋サイズ・行送り）」の指定とみなし、行単位でフォントを適用します。
「ヒラギノ角ゴシック W3 12pt↓16pt」のようにサイズ・行送りを併記でき、併記のない行はフォントだけを適用します。

詳細は README を参照してください。

### Overview

Reads each line (paragraph) of the selected text frames as a "font name (plus size and leading)" spec and applies it line by line.
Sizes and leading can be written after the name, as in "Hiragino Kaku Gothic W3 12pt↓16pt"; a line without them only changes the font.

See the README for details.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ApplyFontByLine";              /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.2.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-06-06";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-09-03";                   /* 更新日 / last updated */

// README (Japanese)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/ApplyFontByLine.md
// README (English)
// https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/ApplyFontByLine.md
var SCRIPT_ARTICLE_URL = "https://note.com/dtp_tranist/n/nc1769b1640e7"; /* 紹介記事 / article URL */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // ユーザー設定 / User settings
    // =========================================

    /* 特定の文字列に強制的に割り当てるフォント（照合前に置き換える）/ Forced font per line text */
    var CUSTOM_FONT_MAP = {
        "Jenson": "Adobe Jenson Pro",
        "Garamond": "Adobe Garamond Pro",
        "Myriad": "Myriad Pro",
        "Frutiger": "Neue Frutiger World",
        "FF DIN": "DIN 2014",
        "Minion": "Minion Pro"
    };

    /* 複数スタイルが候補になったときの優先順位（先頭ほど優先。小文字で比較）/ Style preference order */
    var STYLE_PRIORITY = ["bold", "semibold", "medium", "regular"];

    var MARKER_LAYER_NAME = "// missing-fonts";  /* 未適用の目印を置くレイヤー名 / marker layer name */
    var MARKER_OPACITY    = 35;                  /* 目印の不透明度（％）/ marker opacity in percent */

    var ZOOM_FIT_RATIO = 0.6;   /* ピッカー表示時、対象が表示領域を占める割合 / zoom fit ratio */
    var ZOOM_MIN       = 0.03;  /* Illustrator のズーム下限（3%）/ minimum zoom of Illustrator */
    var ZOOM_MAX       = 64;    /* Illustrator のズーム上限（6400%）/ maximum zoom of Illustrator */

    // =========================================
    // レイアウト設定 / Layout settings
    // =========================================

    var DIALOG_MARGINS      = 15;             /* ダイアログの余白 / dialog margins */
    var BUTTON_BAR_MARGINS  = [0, 10, 0, 0];  /* ボタンバーの余白 / margins of the button bar */
    var PROGRESS_BAR_SIZE   = [320, 9];       /* プログレスバーの寸法 / size of the progress bar */
    var PROGRESS_TEXT_WIDTH = 320;            /* 進捗表示の幅 / width of the progress count label */
    var RESULT_FIELD_SIZE   = [380, 220];     /* 未適用一覧の寸法 / size of the unapplied list field */
    var PICKER_LABEL_WIDTH  = 70;             /* ピッカーの項目名の幅 / width of the picker row labels */
    var PICKER_FIELD_WIDTH  = 200;            /* ピッカーの入力欄・プルダウンの幅 / width of the picker fields */
    var PICKER_TARGET_WIDTH = 260;            /* 対象テキスト表示の幅 / width of the target text label */

    // =========================================
    // ローカライズ / Localization
    // =========================================

    var currentLanguage = (String(app.locale).indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        alert: {
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noSelection: {
                ja: "テキストオブジェクトを選択してください。",
                en: "Please select a text object."
            }
        },
        dialog: {
            progressTitle: {
                ja: "フォントを適用中…",
                en: "Applying fonts…"
            },
            resultTitle: {
                ja: "適用結果",
                en: "Apply results"
            },
            resultHeader: {
                ja: "フォントを適用できなかった文字列：",
                en: "Strings with no matching font:"
            },
            pickerTitle: {
                ja: "フォントを選択",
                en: "Choose font"
            }
        },
        panel: {
            applyFont: {
                ja: "適用するフォント",
                en: "Font to apply"
            }
        },
        fieldLabel: {
            targetText: {
                ja: "対象テキスト",
                en: "Target text"
            },
            search: {
                ja: "検索",
                en: "Search"
            },
            family: {
                ja: "フォント",
                en: "Font"
            },
            style: {
                ja: "スタイル",
                en: "Style"
            }
        },
        button: {
            copy: {
                ja: "クリップボードにコピー",
                en: "Copy to clipboard"
            },
            close: {
                ja: "閉じる",
                en: "Close"
            },
            apply: {
                ja: "適用",
                en: "Apply"
            },
            skip: {
                ja: "スキップ",
                en: "Skip"
            },
            quit: {
                ja: "終了",
                en: "Quit"
            }
        }
    };

    /* showFontPicker が「終了」で返す番兵（選択ループを打ち切る合図）/ Sentinel returned on Quit */
    var PICKER_QUIT = {};

    /**
     * ラベルのリーフ（{ ja, en }）から現在の言語の文字列を返す
     * 現在の言語が未定義なら英語、それも無ければ空文字へフォールバックする
     * @param {Object} labelNode - { ja, en } を持つラベル
     * @returns {string} 現在の言語の文字列
     */
    function getLabel(labelNode) {
        var text = "";
        if (labelNode) {
            text = labelNode[currentLanguage] || labelNode.en || "";
        }
        return String(text).replace(/\{slash\}/g, "/");
    }

    /**
     * 項目名に言語別のコロンを付けて返す（日本語は全角、英語は半角）
     * @param {Object} labelNode - { ja, en } を持つラベル
     * @returns {string} コロン付きの項目名
     */
    function labelText(labelNode) {
        var text = getLabel(labelNode);
        if (text === "") return "";
        return text + (currentLanguage === "ja" ? "：" : ":");
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    if (app.documents.length === 0) {
        alert(getLabel(LABELS.alert.noDocument));
        return;
    }

    var doc = app.activeDocument;
    var fontIndex = null; /* 全インストールフォントの索引（選択チェック後に作成）/ index of installed fonts */

    var selectedItems = doc.selection;
    if (!selectedItems || selectedItems.length === 0) {
        alert(getLabel(LABELS.alert.noSelection));
        return;
    }

    /* 選択物からテキストフレームを再帰収集（グループ内も対象）/ Collect text frames recursively */
    var targetTextFrames = [];
    collectTextFrames(selectedItems, targetTextFrames);

    if (targetTextFrames.length === 0) {
        alert(getLabel(LABELS.alert.noSelection));
        return;
    }

    /* 対象が確定してから索引化する（重い処理なので選択チェックの後）/ Build the index once the target is fixed */
    fontIndex = createFontIndex();

    /* フェーズ1：厳密一致だけ自動適用し、判断が要る行は保留にする / Phase 1: apply confident matches */
    var pendingLines = autoApplyFonts(targetTextFrames);

    /* フェーズ2：保留分を対話ピッカーで1件ずつ決める / Phase 2: resolve the queue interactively */
    var unapplied = resolvePendingLines(pendingLines);

    /* 未適用の行を含むフレームに目印を置く / Mark the frames that still have unapplied lines */
    markUnappliedFrames(unapplied.frames);

    /* 適用できなかった文字列を一覧表示し、要求があればクリップボードへコピー / Show and optionally copy */
    if (unapplied.texts.length > 0 && showUnappliedDialog(unapplied.texts)) {
        copyTextToClipboard(unapplied.texts.join("\n"));
    }

    // =========================================
    // 処理フロー / Processing flow
    // =========================================

    /**
     * 選択物（配列やコレクション）を再帰的にたどり、テキストフレームを集める
     * ロック・非表示のオブジェクトは無視する（グループならその中身ごとスキップ）
     * @param {Object} sourceItems - PageItem の配列またはコレクション
     * @param {TextFrame[]} collectedFrames - 収集先の配列
     * @returns {void}
     */
    function collectTextFrames(sourceItems, collectedFrames) {
        for (var k = 0; k < sourceItems.length; k++) {
            var currentItem = sourceItems[k];

            /* テキスト編集中の選択（TextRange）などは要素を取り出せないので無視 / Skip non page items */
            if (!currentItem || currentItem.locked || currentItem.hidden) continue;

            if (currentItem.typename === "TextFrame") {
                collectedFrames.push(currentItem);
            } else if (currentItem.typename === "GroupItem") {
                collectTextFrames(currentItem.pageItems, collectedFrames);
            }
        }
    }

    /**
     * フェーズ1：全フレームを走査し、厳密一致のフォントだけを自動適用する
     * 判断が要る行（あいまい一致・未一致・適用失敗）は保留リストに積んで返す
     * @param {TextFrame[]} textFrameList - 対象のテキストフレーム
     * @returns {Object[]} 保留行（{ frame, index, lineText, fontName, fontSize, fontLeading, initialFont }）
     */
    function autoApplyFonts(textFrameList) {
        var pendingLines = [];
        var progress = createProgressPalette(textFrameList.length);

        for (var i = 0; i < textFrameList.length; i++) {
            progress.update(i + 1);
            autoApplyLinesInFrame(textFrameList[i], pendingLines);
        }

        progress.window.close();
        return pendingLines;
    }

    /**
     * 1つのテキストフレームの各行を照合し、厳密一致は適用、それ以外は保留に積む
     * @param {TextFrame} textFrame - 対象のテキストフレーム
     * @param {Object[]} pendingLines - 保留行の追加先
     * @returns {void}
     */
    function autoApplyLinesInFrame(textFrame, pendingLines) {
        /* 段落コレクションはキャッシュしない。適用のたびに参照が無効化され Error 1302 になるため
           / Fetch paragraphs live: applying a font invalidates a cached collection */
        var paragraphCount = textFrame.paragraphs.length;

        /* 下から上へたどると、適用によるインデックスのズレを避けられる / Loop upwards to keep indexes valid */
        for (var j = paragraphCount - 1; j >= 0; j--) {
            var lineText = textFrame.paragraphs[j].contents.replace(/^[\s　]+|[\s　]+$/g, "");
            if (lineText === "") continue; /* 空行・空白だけの行は対象外 / Skip blank lines */

            /* 行末に併記されたサイズ・行送りを切り出す（照合は名前部分だけで行う）/ Split off size and leading */
            var fontSpec = parseFontSpec(lineText);
            var fontMatch = findFont(fontSpec.name);

            var isApplied = fontMatch.confident
                && applyFontSpecToLine(textFrame, j, fontMatch.font, fontSpec.size, fontSpec.leading);

            if (!isApplied) {
                pendingLines.push({
                    frame: textFrame,
                    index: j,
                    lineText: lineText,
                    fontName: fontSpec.name,
                    fontSize: fontSpec.size,
                    fontLeading: fontSpec.leading,
                    initialFont: fontMatch.font
                });
            }
        }
    }

    /**
     * フェーズ2：保留行を対話ピッカーで1件ずつ決める
     * 同じフォント名は実行中に一度決めたら再質問せず、その結果を使い回す
     * @param {Object[]} pendingLines - 保留行
     * @returns {Object} 未適用の記録（{ frames: TextFrame[], texts: string[] }）
     */
    function resolvePendingLines(pendingLines) {
        var unapplied = { frames: [], texts: [] };
        var decidedFonts = {}; /* フォント名 -> TextFont（適用）/ null（スキップ）/ decision per font name */
        var fontPicker = null; /* ピッカーは初回だけ生成して使い回す / Build the picker only once */

        for (var i = 0; i < pendingLines.length; i++) {
            var pendingLine = pendingLines[i];

            /* 決定済みのフォント名は再質問しない。サイズ・行送りは行ごとの値を適用する
               / Reuse the earlier decision; size and leading still come from this line */
            if (decidedFonts.hasOwnProperty(pendingLine.fontName)) {
                if (!applyPendingLine(pendingLine, decidedFonts[pendingLine.fontName])) {
                    recordUnapplied(unapplied, pendingLine);
                }
                continue;
            }

            if (!fontPicker) fontPicker = createFontPickerDialog(fontIndex.families);
            var pickedFont = showFontPicker(fontPicker, pendingLine);

            /* 「終了」：残りの保留分をすべて未適用として記録し、ループを打ち切る / Quit stops the queue */
            if (pickedFont === PICKER_QUIT) {
                for (var q = i; q < pendingLines.length; q++) {
                    recordUnapplied(unapplied, pendingLines[q]);
                }
                break;
            }

            decidedFonts[pendingLine.fontName] = pickedFont;
            if (!pickedFont) recordUnapplied(unapplied, pendingLine);
        }

        return unapplied;
    }

    /**
     * 未適用の行を含むフレームの背面に、赤・半透明の長方形を目印として置く
     * 実行のたびに古い目印レイヤーは削除し、作成後はレイヤーをロックする
     * @param {TextFrame[]} frameList - 未適用の行を含むフレーム
     * @returns {void}
     */
    function markUnappliedFrames(frameList) {
        removeLayerByName(MARKER_LAYER_NAME);
        if (frameList.length === 0) return;

        var previousActiveLayer = doc.activeLayer;
        var markerLayer = createMarkerLayer(MARKER_LAYER_NAME);

        for (var i = 0; i < frameList.length; i++) {
            createMarkerRect(frameList[i], markerLayer);
        }

        /* 目印を誤って動かさないようロックし、作業レイヤーは元へ戻す / Lock it and restore the active layer */
        markerLayer.locked = true;
        try { doc.activeLayer = previousActiveLayer; } catch (e) { }
    }

    /**
     * 未適用の行（スキップ・適用失敗）をフレーム・文字列として記録する
     * @param {Object} unapplied - 記録先（{ frames, texts }）
     * @param {Object} pendingLine - 対象の保留行
     * @returns {void}
     */
    function recordUnapplied(unapplied, pendingLine) {
        pushUnique(unapplied.texts, pendingLine.lineText);
        pushUnique(unapplied.frames, pendingLine.frame); /* 同一参照なので === で重複排除できる */
    }

    /**
     * 配列に未登録の値だけ追加する（重複防止）
     * @param {Array} targetList - 追加先の配列
     * @param {Object} newValue - 追加する値
     * @returns {void}
     */
    function pushUnique(targetList, newValue) {
        for (var i = 0; i < targetList.length; i++) {
            if (targetList[i] === newValue) return;
        }
        targetList.push(newValue);
    }

    // =========================================
    // フォントの適用 / Applying fonts
    // =========================================

    /**
     * 保留行に、決定したフォントとその行のサイズ・行送りを適用する
     * @param {Object} pendingLine - 対象の保留行
     * @param {TextFont} font - 適用するフォント（null なら適用しない）
     * @returns {boolean} 適用できたかどうか
     */
    function applyPendingLine(pendingLine, font) {
        return applyFontSpecToLine(pendingLine.frame, pendingLine.index, font,
            pendingLine.fontSize, pendingLine.fontLeading);
    }

    /**
     * 段落へフォント・サイズ・行送りを適用する（サイズ・行送りは指定があるものだけ）
     * 段落は使用直前にライブ取得する。保持した参照を使い回すと無効化され、
     * try/catch でも拾えないネイティブクラッシュを起こすため
     * 行送りは絶対値では入れず、行送り÷サイズの百分率を自動行送りに代入する
     * （例：12pt↓16pt → 16/12 ≒ 133.3%）
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {number} paragraphIndex - 段落インデックス
     * @param {TextFont} font - 適用するフォント（null なら適用しない）
     * @param {number} sizePt - 文字サイズ（ポイント。null なら変更しない）
     * @param {number} leadingPt - 行送り（ポイント。null なら変更しない）
     * @returns {boolean} 適用できたかどうか
     */
    function applyFontSpecToLine(frame, paragraphIndex, font, sizePt, leadingPt) {
        if (!font) return false;

        try {
            var paragraph = frame.paragraphs[paragraphIndex];
            var attributes = paragraph.characterAttributes;

            attributes.textFont = font;
            if (sizePt) attributes.size = sizePt;
            if (leadingPt && sizePt) {
                paragraph.paragraphAttributes.autoLeadingAmount = (leadingPt / sizePt) * 100;
                attributes.autoLeading = true;
            }
            return true;
        } catch (e) {
            return false;
        }
    }

    /**
     * 行末に併記されたフォントサイズ・行送りを切り出し、{ name, size, leading } で返す
     * 例）"ヒラギノ角ゴシック W3 12pt↓16pt" → { name: "ヒラギノ角ゴシック W3", size: 12, leading: 16 }
     * 　　"Helvetica Neue, 14"            → { name: "Helvetica Neue", size: 14, leading: null }
     * 　　"DIN 2014"                      → { name: "DIN 2014", size: null, leading: null }
     * サイズは「単位付き」か「、」「,」「/」区切りのどちらかで必ずアンカーされるため、
     * フォント名末尾の数字（"DIN 2014" など）はサイズと誤認されない
     * @param {string} rawText - 行の文字列
     * @returns {Object} { name: string, size: number, leading: number }（size / leading は未指定なら null）
     */
    function parseFontSpec(rawText) {
        var text = String(rawText).replace(/^[\s　]+|[\s　]+$/g, "");
        var fontSpec = { name: text, size: null, leading: null };

        /* 末尾の行送りは任意：[区切り（↓含む）→ 数値 →（任意で単位）] / Trailing leading is optional */
        var trailingLeading = "(?:[ \\t　、,/↓]+([0-9]+(?:\\.[0-9]+)?)(pt|px|q)?)?$";

        /* パターン1：サイズに単位（pt/px/Q）が付くケース / Size with an explicit unit */
        var sizeWithUnit = new RegExp(
            "^([\\s\\S]*?)[ \\t　、,/]+([0-9]+(?:\\.[0-9]+)?)(pt|px|q)" + trailingLeading, "i");

        /* パターン2：単位は無いが「、」「,」「/」で区切られているケース / Size after a list separator */
        var sizeWithListSeparator = new RegExp(
            "^([\\s\\S]*?)[ \\t　]*[、,/][ \\t　]*([0-9]+(?:\\.[0-9]+)?)(pt|px|q)?" + trailingLeading, "i");

        /* groups: 1=名前 / 2=サイズ数値 / 3=サイズ単位 / 4=行送り数値 / 5=行送り単位 */
        var matched = text.match(sizeWithUnit) || text.match(sizeWithListSeparator);
        if (!matched) return fontSpec;

        var name = String(matched[1]).replace(/[\s　]+$/g, "");
        if (name === "") return fontSpec; /* 名前が空ならサイズ扱いしない / Not a size when the name is empty */

        var sizePt = unitToPoints(matched[2], matched[3]);
        if (sizePt === null) return fontSpec;

        fontSpec.name = name;
        fontSpec.size = sizePt;

        /* サイズが確定したときだけ行送りを見る / Read the leading only once the size is settled */
        if (matched[4]) fontSpec.leading = unitToPoints(matched[4], matched[5]);

        return fontSpec;
    }

    /**
     * 数値文字列＋単位をポイント値へ換算する（1Q = 0.25mm、px/pt/単位なしは 1px = 1pt）
     * @param {string} numberText - 数値の文字列
     * @param {string} unitText - 単位（pt / px / q / 未指定）
     * @returns {number} ポイント値。正の値でなければ null
     */
    function unitToPoints(numberText, unitText) {
        var value = parseFloat(numberText);
        if (!(value > 0)) return null;
        if (unitText && String(unitText).toLowerCase() === "q") {
            return value * (72 / 25.4 / 4); /* 1Q = 0.25mm → pt */
        }
        return value;
    }

    // =========================================
    // フォントの照合 / Font matching
    // =========================================

    /**
     * インストール済みフォントを索引化する（照合と一覧表示の両方で使う）
     * @returns {Object} フォント索引
     */
    function createFontIndex() {
        var installedFonts = app.textFonts;
        var index = {
            families: [],                /* 表示用のファミリー名（重複なし・ソート済み）*/
            stylesByFamily: {},          /* ファミリー名 -> スタイル名の配列 */
            fontByFamilyStyle: {},       /* ファミリー名＋スタイル名 -> TextFont */
            fontsByNormalizedFamily: {}, /* 正規化ファミリー名 -> TextFont の配列 */
            fontsByNormalizedFull: {},   /* 正規化「ファミリー名＋スタイル名」-> TextFont の配列 */
            normalizedFonts: [],         /* 部分一致の走査用 */
            customFontByNormalized: {}   /* 正規化した CUSTOM_FONT_MAP のキー -> 置換後のフォント名 */
        };

        for (var i = 0; i < installedFonts.length; i++) {
            var currentFont = installedFonts[i];
            var family = currentFont.family;
            var style = currentFont.style;

            if (!index.stylesByFamily.hasOwnProperty(family)) {
                index.stylesByFamily[family] = [];
                index.families.push(family);
            }
            index.stylesByFamily[family].push(style);
            index.fontByFamilyStyle[makeFamilyStyleKey(family, style)] = currentFont;

            var normalizedFamily = normalize(family);
            var normalizedFull = normalize(family + " " + style);
            pushToBucket(index.fontsByNormalizedFamily, normalizedFamily, currentFont);
            pushToBucket(index.fontsByNormalizedFull, normalizedFull, currentFont);

            index.normalizedFonts.push({
                font: currentFont,
                family: normalizedFamily,
                full: normalizedFull,
                name: normalize(currentFont.name)
            });
        }

        index.families.sort();
        for (var familyName in index.stylesByFamily) {
            if (index.stylesByFamily.hasOwnProperty(familyName)) {
                index.stylesByFamily[familyName].sort();
            }
        }

        /* カスタム置換ルールも正規化しておき、大文字小文字やスペースのゆれを吸収する
           / Normalize the custom map so case and spacing differences still hit */
        for (var customKey in CUSTOM_FONT_MAP) {
            if (CUSTOM_FONT_MAP.hasOwnProperty(customKey)) {
                index.customFontByNormalized[normalize(customKey)] = CUSTOM_FONT_MAP[customKey];
            }
        }

        return index;
    }

    /**
     * 索引のバケット（キー -> 配列）へ値を追加する
     * @param {Object} bucketMap - バケットを保持するオブジェクト
     * @param {string} key - キー
     * @param {Object} value - 追加する値
     * @returns {void}
     */
    function pushToBucket(bucketMap, key, value) {
        if (!bucketMap.hasOwnProperty(key)) bucketMap[key] = [];
        bucketMap[key].push(value);
    }

    /**
     * ファミリー名＋スタイル名の索引用キーを作る
     * @param {string} family - ファミリー名
     * @param {string} style - スタイル名
     * @returns {string} 索引用のキー
     */
    function makeFamilyStyleKey(family, style) {
        return family + " " + style;
    }

    /**
     * 指定ファミリーに属するスタイル名の一覧を返す
     * @param {string} familyName - ファミリー名
     * @returns {string[]} スタイル名の配列（無ければ空配列）
     */
    function stylesForFamily(familyName) {
        return fontIndex.stylesByFamily[familyName] || [];
    }

    /**
     * ファミリー名＋スタイル名から TextFont を取得する
     * @param {string} familyName - ファミリー名
     * @param {string} styleName - スタイル名
     * @returns {TextFont} 該当するフォント（無ければ null）
     */
    function fontFor(familyName, styleName) {
        return fontIndex.fontByFamilyStyle[makeFamilyStyleKey(familyName, styleName)] || null;
    }

    /**
     * 文字列に対応するフォントを、厳密一致 → あいまい一致の順に探す
     * @param {string} fontName - 行から切り出したフォント名
     * @returns {Object} { font: TextFont, confident: boolean }
     *   confident=true … 厳密一致。自動適用してよい
     *   confident=false 且つ font!=null … あいまい一致。ピッカーの初期値に使う
     *   font=null … 未一致。ピッカーで一から選ばせる
     */
    function findFont(fontName) {
        var normalizedName = normalize(fontName);
        var targetName = fontIndex.customFontByNormalized.hasOwnProperty(normalizedName)
            ? fontIndex.customFontByNormalized[normalizedName]
            : fontName;

        var exactFont = findExactFont(targetName);
        if (exactFont) return { font: exactFont, confident: true };

        return { font: findFuzzyFont(targetName), confident: false };
    }

    /**
     * 厳密一致でフォントを探す（PostScript 名 → ファミリー名＋スタイル名 → ファミリー名）
     * 照合はスペース・ピリオドを除いた正規化文字列で行う
     * @param {string} fontName - 探すフォント名
     * @returns {TextFont} 該当するフォント（無ければ null）
     */
    function findExactFont(fontName) {
        try {
            return app.textFonts.getByName(fontName); /* PostScript 名の完全一致 */
        } catch (e) { }

        var query = normalize(fontName);

        var exactFull = fontIndex.fontsByNormalizedFull[query];
        if (exactFull) return getBestStyle(exactFull);

        var exactFamily = fontIndex.fontsByNormalizedFamily[query];
        if (exactFamily) return getBestStyle(exactFamily);

        return null;
    }

    /**
     * あいまい一致でフォントを探す（ファミリー名の部分一致 → フォント名全体の部分一致 → 先頭ワード）
     * 例）"Jenson" → "Adobe Jenson Pro"、"Myriad Pro Cond"（未インストール）→ "Myriad Pro"
     * @param {string} fontName - 探すフォント名
     * @returns {TextFont} 該当するフォント（無ければ null）
     */
    function findFuzzyFont(fontName) {
        var query = normalize(fontName);

        var partialFamily = collectIndexedFonts(function (fontInfo) {
            return fontInfo.family.indexOf(query) !== -1;
        });
        if (partialFamily.length > 0) return getBestStyle(partialFamily);

        var partialFull = collectIndexedFonts(function (fontInfo) {
            return fontInfo.full.indexOf(query) !== -1 || fontInfo.name.indexOf(query) !== -1;
        });
        if (partialFull.length > 0) return getBestStyle(partialFull);

        /* 最終手段：先頭ワードがファミリー名に含まれればOK。2文字以下は誤マッチ防止のため対象外
           / Last resort: match on the first word, ignoring words of two characters or fewer */
        var firstWord = String(fontName).toLowerCase()
            .replace(/[.　]+/g, " ").replace(/^\s+/, "").split(/\s+/)[0] || "";
        if (firstWord.length >= 3) {
            var looseFamily = collectIndexedFonts(function (fontInfo) {
                return fontInfo.family.indexOf(firstWord) !== -1;
            });
            if (looseFamily.length > 0) return getBestStyle(looseFamily);
        }

        return null;
    }

    /**
     * 正規化済みフォント索引から、条件に合うフォントを集めて返す
     * @param {function} isMatch - 判定関数（{ font, family, full, name } を受け取る）
     * @returns {TextFont[]} 条件に合ったフォント
     */
    function collectIndexedFonts(isMatch) {
        var matchedFonts = [];
        var normalizedFonts = fontIndex.normalizedFonts;
        for (var i = 0; i < normalizedFonts.length; i++) {
            if (isMatch(normalizedFonts[i])) {
                matchedFonts.push(normalizedFonts[i].font);
            }
        }
        return matchedFonts;
    }

    /**
     * フォント名照合用の正規化：小文字化し、空白（半角・全角）とピリオドを除去する
     * 「Bank Gothic」と「BankGothic」、「Mrs. Eaves」と「Mrs Eaves」のゆれを吸収する
     * @param {string} text - 対象の文字列
     * @returns {string} 正規化した文字列
     */
    function normalize(text) {
        return String(text)
            .toLowerCase()
            .replace(/[.\s　]+/g, "");
    }

    /**
     * 候補フォントから STYLE_PRIORITY の順で最適なスタイルを選ぶ
     * @param {TextFont[]} candidateFonts - 候補のフォント
     * @returns {TextFont} 選ばれたフォント（該当が無ければ候補の先頭）
     */
    function getBestStyle(candidateFonts) {
        for (var k = 0; k < STYLE_PRIORITY.length; k++) {
            for (var j = 0; j < candidateFonts.length; j++) {
                if (candidateFonts[j].style.toLowerCase() === STYLE_PRIORITY[k]) {
                    return candidateFonts[j];
                }
            }
        }
        return candidateFonts[0];
    }

    // =========================================
    // ダイアログ / Dialogs
    // =========================================

    /**
     * 適用中に表示するプログレスバー（パレット）を作成する
     * @param {number} totalCount - 対象の総数
     * @returns {Object} { window: Window, update: function }
     */
    function createProgressPalette(totalCount) {
        var progressWindow = new Window("palette", getLabel(LABELS.dialog.progressTitle) + " " + SCRIPT_VERSION);
        progressWindow.alignChildren = "fill";
        progressWindow.margins = DIALOG_MARGINS;

        var progressBar = progressWindow.add("progressbar", undefined, 0, totalCount);
        progressBar.preferredSize = PROGRESS_BAR_SIZE;

        var countLabel = progressWindow.add("statictext", undefined, "0 / " + totalCount);
        countLabel.preferredSize.width = PROGRESS_TEXT_WIDTH;

        progressWindow.show();
        progressWindow.update();

        return {
            window: progressWindow,
            update: function (doneCount) {
                progressBar.value = doneCount;
                countLabel.text = doneCount + " / " + totalCount;
                progressWindow.update();
            }
        };
    }

    /**
     * 適用できなかった文字列を一覧表示する
     * コピーはダイアログを閉じてから行う（モーダル表示中はドキュメントを触らない）
     * @param {string[]} unappliedTexts - 未適用の文字列
     * @returns {boolean} クリップボードへのコピーが要求されたか
     */
    function showUnappliedDialog(unappliedTexts) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.resultTitle) + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";
        dialog.margins = DIALOG_MARGINS;

        dialog.add("statictext", undefined, getLabel(LABELS.dialog.resultHeader));

        /* 一覧（読み取り専用・複数行・スクロール可）。手動で選択もできる / Read-only scrollable list */
        var listField = dialog.add("edittext", undefined, unappliedTexts.join("\n"),
            { multiline: true, scrolling: true, readonly: true });
        listField.preferredSize = RESULT_FIELD_SIZE;

        /* === ボタンエリア（左右分割：左=コピー／右=閉じる）=== */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnCopy = btnLeftGroup.add("button", undefined, getLabel(LABELS.button.copy));
        btnCopy.onClick = function () { dialog.close(2); }; /* 2 = コピーして閉じる / copy and close */

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.close), { name: "ok" });

        return dialog.show() === 2;
    }

    /**
     * 文字列をクリップボードへコピーする
     * ExtendScript には直接コピーする API が無いため、一時テキストフレーム経由でコピーする
     * app.copy() は黙って無視されることがあるので、再描画とメニューコマンドを挟む
     * @param {string} textToCopy - コピーする文字列
     * @returns {void}
     */
    function copyTextToClipboard(textToCopy) {
        var previousSelection = doc.selection; /* コピー後に元へ戻すため控える / Keep it to restore later */
        var editableLayer = getEditableLayer();
        if (!editableLayer) return;

        var tempFrame = null;
        try {
            doc.activeLayer = editableLayer;
            tempFrame = doc.textFrames.add();
            tempFrame.contents = textToCopy;
            tempFrame.position = [-100000, -100000]; /* 画面外に逃がす / Move it off-canvas */

            app.redraw(); /* 追加直後のフレームは再描画しないとコピー対象にならない */
            app.executeMenuCommand("deselectall");
            tempFrame.selected = true;
            app.redraw();
            app.executeMenuCommand("copy"); /* app.copy() は黙って無視されることがある */
            app.redraw(); /* コピー確定前に削除すると空になる */
        } catch (e) {
        } finally {
            if (tempFrame) {
                try { tempFrame.remove(); } catch (err) { }
            }
            try { doc.selection = previousSelection; } catch (err) { doc.selection = null; }
        }
    }

    /**
     * 一時オブジェクトを置ける（ロックも非表示もされていない）レイヤーを返す
     * @returns {Layer} 編集できるレイヤー（無ければ null）
     */
    function getEditableLayer() {
        if (!doc.activeLayer.locked && doc.activeLayer.visible) return doc.activeLayer;

        for (var i = 0; i < doc.layers.length; i++) {
            if (!doc.layers[i].locked && doc.layers[i].visible) return doc.layers[i];
        }
        return null;
    }

    // =========================================
    // フォントピッカー / Font picker
    // =========================================

    /**
     * 保留行のフォントを対話的に選ばせる
     * ダイアログは作り直さず使い回し、ここでは対象テキストと選択のリセットだけ行う
     * モーダル表示中はドキュメントを変更しない（ライブプレビューはクラッシュ要因のため廃止）
     * @param {Object} fontPicker - createFontPickerDialog が返したピッカー
     * @param {Object} pendingLine - 対象の保留行
     * @returns {TextFont} 適用したフォント／スキップは null／終了は PICKER_QUIT
     */
    function showFontPicker(fontPicker, pendingLine) {
        zoomToFrame(pendingLine.frame); /* 対象を画面にフィット（モーダル表示前に行う）*/

        fontPicker.targetLabel.text = pendingLine.lineText;
        fontPicker.searchField.text = "";
        filterFamilyDropdown(fontPicker, "");
        initializePickerSelection(fontPicker, pendingLine.initialFont);

        /* 「適用」=1 /「スキップ」=2（閉じるを含む）/「終了」=3 */
        var dialogResult = fontPicker.dialog.show();

        if (dialogResult === 1 && fontPicker.familyDropdown.selection && fontPicker.styleDropdown.selection) {
            var chosenFont = fontFor(fontPicker.familyDropdown.selection.text,
                fontPicker.styleDropdown.selection.text);
            /* 実適用はダイアログを閉じた後（＝モーダル表示外）なので安全 / Apply once the modal is closed */
            if (applyPendingLine(pendingLine, chosenFont)) return chosenFont;
        }

        return (dialogResult === 3) ? PICKER_QUIT : null;
    }

    /**
     * フォントピッカーの UI を生成する
     * ボタンは name:"ok"/"cancel" なので、判定は dialog.show() の戻り値で行う
     * @param {string[]} familyNames - ファミリー名の一覧
     * @returns {Object} ピッカー（{ dialog, familyDropdown, styleDropdown, searchField, targetLabel, familyNames }）
     */
    function createFontPickerDialog(familyNames) {
        var dialog = new Window("dialog", getLabel(LABELS.dialog.pickerTitle) + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";
        dialog.margins = DIALOG_MARGINS;

        /* 対象テキスト（パネルの外）。値はピックごとに差し替えるので空で作り、幅だけ確保する */
        var targetRow = dialog.add("group");
        targetRow.add("statictext", undefined, labelText(LABELS.fieldLabel.targetText));
        var targetLabel = targetRow.add("statictext", undefined, "", { truncate: "end" });
        targetLabel.preferredSize.width = PICKER_TARGET_WIDTH;

        var applyFontPanel = dialog.add("panel", undefined, getLabel(LABELS.panel.applyFont));
        applyFontPanel.orientation = "column";
        applyFontPanel.alignChildren = "left";
        applyFontPanel.margins = DIALOG_MARGINS;

        var searchField = addPickerRow(applyFontPanel, LABELS.fieldLabel.search, "edittext", "");

        /* 巨大配列での dropdownlist 生成を繰り返すと Illustrator が落ちるため、
           このダイアログ（と familyDropdown）は実行中に1回だけ生成して使い回す */
        var familyDropdown = addPickerRow(applyFontPanel, LABELS.fieldLabel.family, "dropdownlist", familyNames);
        var styleDropdown = addPickerRow(applyFontPanel, LABELS.fieldLabel.style, "dropdownlist", []);

        /* === ボタンエリア（左右分割：左=終了／右=スキップ・適用）=== */
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = BUTTON_BAR_MARGINS;
        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnQuit = btnLeftGroup.add("button", undefined, getLabel(LABELS.button.quit));
        btnQuit.onClick = function () { dialog.close(3); }; /* 3 = 終了（保留分を打ち切る）/ quit the queue */

        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.skip), { name: "cancel" });
        btnRightGroup.add("button", undefined, getLabel(LABELS.button.apply), { name: "ok" });

        var fontPicker = {
            dialog: dialog,
            familyDropdown: familyDropdown,
            styleDropdown: styleDropdown,
            searchField: searchField,
            targetLabel: targetLabel,
            familyNames: familyNames
        };

        bindFontPickerEvents(fontPicker); /* イベントも生成時に1回だけ接続する / Bind the events once too */
        return fontPicker;
    }

    /**
     * ピッカーのパネルへ「項目名＋コントロール」の行を追加する
     * @param {Panel} parentPanel - 追加先のパネル
     * @param {Object} labelNode - 項目名のラベル（{ ja, en }）
     * @param {string} controlType - コントロールの種類（"edittext" / "dropdownlist"）
     * @param {Object} initialValue - 初期値（edittext は文字列、dropdownlist は項目の配列）
     * @returns {Object} 追加したコントロール
     */
    function addPickerRow(parentPanel, labelNode, controlType, initialValue) {
        var row = parentPanel.add("group");
        var rowLabel = row.add("statictext", undefined, labelText(labelNode));
        rowLabel.preferredSize.width = PICKER_LABEL_WIDTH;

        var control = row.add(controlType, undefined, initialValue);
        control.preferredSize.width = PICKER_FIELD_WIDTH;
        return control;
    }

    /**
     * フォントピッカーの検索・ファミリー変更イベントを接続する
     * （ドキュメントは変更しない。実適用は「適用」確定後に行う）
     * @param {Object} fontPicker - 対象のピッカー
     * @returns {void}
     */
    function bindFontPickerEvents(fontPicker) {
        fontPicker.searchField.onChanging = function () {
            filterFamilyDropdown(fontPicker, fontPicker.searchField.text);
            populateStyleDropdown(fontPicker, null);
        };

        fontPicker.familyDropdown.onChange = function () {
            populateStyleDropdown(fontPicker, null);
        };
    }

    /**
     * ピッカーの初期選択（ファミリー・スタイル）を設定する
     * @param {Object} fontPicker - 対象のピッカー
     * @param {TextFont} initialFont - 初期値にするフォント（無ければ null）
     * @returns {void}
     */
    function initializePickerSelection(fontPicker, initialFont) {
        selectDropdownItemByText(fontPicker.familyDropdown,
            initialFont ? initialFont.family : fontPicker.familyNames[0]);
        if (!fontPicker.familyDropdown.selection) fontPicker.familyDropdown.selection = 0;

        populateStyleDropdown(fontPicker, initialFont ? initialFont.style : null);
    }

    /**
     * 選択中のファミリーに合わせてスタイル一覧を作り直し、選ぶべきスタイルを選択する
     * @param {Object} fontPicker - 対象のピッカー
     * @param {string} preferredStyleName - 優先して選ぶスタイル名（無ければ null）
     * @returns {void}
     */
    function populateStyleDropdown(fontPicker, preferredStyleName) {
        var styleDropdown = fontPicker.styleDropdown;
        styleDropdown.removeAll();
        if (!fontPicker.familyDropdown.selection) return;

        var styles = stylesForFamily(fontPicker.familyDropdown.selection.text);
        for (var i = 0; i < styles.length; i++) {
            styleDropdown.add("item", styles[i]);
        }

        if (preferredStyleName) selectDropdownItemByText(styleDropdown, preferredStyleName);
        if (!styleDropdown.selection && styleDropdown.items.length > 0) styleDropdown.selection = 0;
    }

    /**
     * 検索クエリ（部分一致・大文字小文字を無視）でファミリーのプルダウンを絞り込む
     * 絞り込み後も、可能なら直前に選択していたファミリーを選び直す
     * 空クエリで全件表示済みなら作り直さない（既存コントロールへの item 追加は安全だが、無駄を避ける）
     * @param {Object} fontPicker - 対象のピッカー
     * @param {string} searchQuery - 検索クエリ（空文字なら全件）
     * @returns {void}
     */
    function filterFamilyDropdown(fontPicker, searchQuery) {
        var familyDropdown = fontPicker.familyDropdown;
        var familyNames = fontPicker.familyNames;
        var needle = String(searchQuery).toLowerCase();

        if (needle === "" && familyDropdown.items.length === familyNames.length) return;

        var previousFamily = familyDropdown.selection ? familyDropdown.selection.text : null;
        familyDropdown.removeAll();
        for (var i = 0; i < familyNames.length; i++) {
            if (needle === "" || familyNames[i].toLowerCase().indexOf(needle) !== -1) {
                familyDropdown.add("item", familyNames[i]);
            }
        }

        if (previousFamily) selectDropdownItemByText(familyDropdown, previousFamily);
        if (!familyDropdown.selection && familyDropdown.items.length > 0) familyDropdown.selection = 0;
    }

    /**
     * プルダウンで指定テキストの項目を選択する（無ければ何もしない）
     * @param {DropDownList} dropdown - 対象のプルダウン
     * @param {string} itemText - 選択する項目のテキスト
     * @returns {void}
     */
    function selectDropdownItemByText(dropdown, itemText) {
        for (var i = 0; i < dropdown.items.length; i++) {
            if (dropdown.items[i].text === itemText) {
                dropdown.selection = i;
                return;
            }
        }
    }

    /**
     * 対象フレームが画面にフィットするようズーム＋センタリングする
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @returns {void}
     */
    function zoomToFrame(frame) {
        try {
            var view = doc.activeView;
            var bounds = frame.geometricBounds; /* [left, top, right, bottom] */
            var frameWidth = bounds[2] - bounds[0];
            var frameHeight = bounds[1] - bounds[3];
            if (frameWidth <= 0 || frameHeight <= 0) return;

            var viewBounds = view.bounds; /* 現在の表示範囲（ドキュメント座標）/ current view in document coordinates */
            var viewWidth = viewBounds[2] - viewBounds[0];
            var viewHeight = viewBounds[1] - viewBounds[3];

            /* 表示領域の ZOOM_FIT_RATIO に収まる倍率を現在ズームから算出し、上下限に収める */
            var fitZoom = Math.min(viewWidth / frameWidth, viewHeight / frameHeight) * view.zoom * ZOOM_FIT_RATIO;
            view.zoom = Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, fitZoom));
            view.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (e) { }
    }

    // =========================================
    // 目印レイヤー / Marker layer
    // =========================================

    /**
     * 目印レイヤーを新規作成し、テキストの背面に来るよう最背面へ送る
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 作成したレイヤー
     */
    function createMarkerLayer(layerName) {
        var markerLayer = doc.layers.add();
        markerLayer.name = layerName;

        try {
            markerLayer.move(doc.layers[doc.layers.length - 1], ElementPlacement.PLACEAFTER);
        } catch (e) { }

        return markerLayer;
    }

    /**
     * テキストフレームの背面（＝目印レイヤー上）に、赤・半透明の長方形を作る
     * @param {TextFrame} frame - 対象のテキストフレーム
     * @param {Layer} layer - 目印を置くレイヤー
     * @returns {void}
     */
    function createMarkerRect(frame, layer) {
        try {
            var bounds = frame.geometricBounds; /* [left, top, right, bottom]（線幅は含まない）*/
            var width = bounds[2] - bounds[0];
            var height = bounds[1] - bounds[3];
            if (width <= 0 || height <= 0) return;

            var markerRect = layer.pathItems.rectangle(bounds[1], bounds[0], width, height);
            markerRect.stroked = false;
            markerRect.filled = true;
            markerRect.fillColor = makeRedColor();
            markerRect.opacity = MARKER_OPACITY;
        } catch (e) { }
    }

    /**
     * ドキュメントのカラースペースに合わせた赤色を返す
     * @returns {Object} CMYKColor または RGBColor
     */
    function makeRedColor() {
        if (doc.documentColorSpace === DocumentColorSpace.CMYK) {
            var cmyk = new CMYKColor();
            cmyk.cyan = 0;
            cmyk.magenta = 100;
            cmyk.yellow = 100;
            cmyk.black = 0;
            return cmyk;
        }

        var rgb = new RGBColor();
        rgb.red = 255;
        rgb.green = 0;
        rgb.blue = 0;
        return rgb;
    }

    /**
     * 指定名のレイヤーがあれば削除する（ロックされていても解除してから削除する）
     * @param {string} layerName - レイヤー名
     * @returns {void}
     */
    function removeLayerByName(layerName) {
        var layer = getLayerByName(layerName);
        if (!layer) return;

        unlockContainer(layer);
        unlockPageItems(layer.pageItems);

        try {
            layer.remove();
        } catch (e) { }
    }

    /**
     * 指定名のレイヤーを取得する
     * @param {string} layerName - レイヤー名
     * @returns {Layer} 該当するレイヤー（無ければ null）
     */
    function getLayerByName(layerName) {
        try {
            return doc.layers.getByName(layerName);
        } catch (e) {
            return null;
        }
    }

    /**
     * レイヤー／ページアイテムのロックを解除し、表示状態に戻す
     * 表示のプロパティはレイヤーが visible、ページアイテムが hidden と異なる
     * @param {Object} target - Layer または PageItem
     * @returns {void}
     */
    function unlockContainer(target) {
        try {
            target.locked = false;
            if (target.typename === "Layer") {
                target.visible = true;
            } else {
                target.hidden = false;
            }
        } catch (e) { }
    }

    /**
     * ページアイテムを再帰的にロック解除・表示する
     * @param {Object} items - PageItems コレクション
     * @returns {void}
     */
    function unlockPageItems(items) {
        for (var i = 0; i < items.length; i++) {
            unlockContainer(items[i]);

            if (items[i].typename === "GroupItem") {
                unlockPageItems(items[i].pageItems);
            }
        }
    }
})();
