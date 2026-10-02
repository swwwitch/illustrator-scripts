#target illustrator
#targetengine "CenterAlignToArtboards"
app.preferences.setBooleanPreference("ShowExternalJSXWarning", false);

/*

### 概要

選択したオブジェクトをグループ化せず1つずつ、それぞれが載っているアートボードの天地中央・左右中央に整列します。
ダイアログで天地中央・左右中央を選べます。実行中だけ「字形の境界に整列」をONにし、終了時に元の状態へ戻します。

### Overview

Aligns each selected object, one at a time and without grouping, to the vertical and/or horizontal center of the artboard it sits on.
Choose the directions in the dialog. "Align to glyph bounds" is turned on only while it runs and restored afterwards.

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "CenterAlignToArtboards";       /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.0.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "2026-10-03";                   /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "2026-10-03";                   /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function() {

    // =========================================
    // ユーザー設定 / User Settings
    // =========================================
    var DEFAULT_ALIGN_OPTIONS = {
        vertical: true,   /* 天地中央の初期値 / vertical center on by default */
        horizontal: true  /* 左右中央の初期値 / horizontal center on by default */
    };

    // =========================================
    // レイアウト / Layout
    // =========================================

    // UIレイアウト（再利用パーツ） / UI layout (reusable)

    /* ウィンドウ・パネルの余白と間隔 / Window & panel margins and spacing */
    var WINDOW_MARGINS = 16;                 /* ウィンドウ外周の余白 / window margin */
    var WINDOW_SPACING = 12;                 /* ウィンドウ内の要素間隔 / window spacing */
    var PANEL_MARGINS  = [16, 20, 16, 12];   /* パネル余白 [左,上,右,下] / panel margins */
    var PANEL_SPACING  = 12;                 /* パネル内の要素間隔 / panel spacing */
    var COLUMN_SPACING = 12;                 /* 2カラムの間隔 / gap between columns */
    var TAB_MARGINS    = [15, 20, 5, 10];    /* タブ余白 [左,上,右,下] / tab margins */

    /**
     * ウィンドウの共通設定
     * @param {Window} targetWindow - 対象のウィンドウ
     * @param {number} [spacing] - 要素間隔（省略時は WINDOW_SPACING）
     * @returns {void}
     */
    function setupWindow(targetWindow, spacing) {
        targetWindow.orientation = "column";
        targetWindow.alignChildren = "fill";
        targetWindow.margins = WINDOW_MARGINS;
        targetWindow.spacing = (typeof spacing === "number") ? spacing : WINDOW_SPACING;
    }

    /**
     * パネルの共通設定（子は幅いっぱい。ボタンは alignment = "left" で広げない）
     * @param {Panel} targetPanel - 対象のパネル
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupPanel(targetPanel, spacing) {
        targetPanel.orientation = "column";
        targetPanel.alignChildren = ["fill", "top"];
        targetPanel.alignment = "fill";
        targetPanel.margins = PANEL_MARGINS;
        targetPanel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * タブの共通設定
     * @param {Tab} targetTab - 対象のタブ
     * @param {number} [spacing] - 要素間隔（省略時は変えない）
     * @returns {void}
     */
    function setupTab(targetTab, spacing) {
        targetTab.orientation = "column";
        targetTab.alignChildren = "fill";
        targetTab.margins = TAB_MARGINS;
        if (typeof spacing === "number") targetTab.spacing = spacing;
    }

    /**
     * 横並びの行グループの共通設定（ボタン列など）。
     * alignment と alignChildren を対で指定し、中のボタンが横に伸びたり天地がずれたりしないようにする
     * @param {Group} rowGroup - 対象のグループ
     * @param {string|string[]} [rowAlignment] - 横方向の alignment（省略時は "left"）。配列ならそのまま使う
     * @param {number} [spacing] - 要素間隔（省略時は PANEL_SPACING）
     * @returns {void}
     */
    function setupRow(rowGroup, rowAlignment, spacing) {
        rowGroup.orientation = "row";
        rowGroup.alignment = (rowAlignment instanceof Array) ? rowAlignment : [rowAlignment || "left", "center"];
        rowGroup.alignChildren = ["left", "center"];
        rowGroup.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /**
     * ボタンの高さを指定した px だけ詰める（レイアウトが決まったあとに呼ぶ）
     * @param {Button} targetButton - 対象のボタン
     * @param {number} trimPixels - 詰める量（px）
     * @returns {void}
     */
    function trimButtonHeight(targetButton, trimPixels) {
        /* レイアウト前は size が無い / size is not set until the layout runs */
        if (!targetButton.size) return;
        targetButton.size = [targetButton.size.width, targetButton.size.height - trimPixels];
    }

    // UIレイアウト（再利用パーツ）ここまで / End of the reusable UI layout

    // =========================================
    // アクション設定 / Action settings
    // =========================================
    var ACTION_SET_NAME = "CenterAlignToArtboards"; /* アクションセット名 / action set name */
    var ACTION_NAME     = "Center";                 /* アクション名 / action name */

    /* 整列パネルのコマンド名と値（.aia の parameter-1） / Align panel command names and values */
    var ALIGN_EVENT_SPECS = {
        horizontal: { commandName: "水平方向中央に整列", value: 2 },
        vertical:   { commandName: "垂直方向中央に整列", value: 5 }
    };

    /**
     * 整列パネルの1コマンド分のイベント定義の行を返す
     * @param {number} eventNumber - イベント番号（1始まり）
     * @param {Object} eventSpec - ALIGN_EVENT_SPECS の要素
     * @returns {string[]} イベント定義の行
     */
    function buildAlignEventLines(eventNumber, eventSpec) {
        return [].concat(
            ["\t/event-" + eventNumber + " {",
             "\t\t/useRulersIn1stQuadrant 0",
             "\t\t/internalName (ai_plugin_alignPalette)"],
            buildActionNameLines("\t\t", "整列", "localizedName"),
            ["\t\t/isOpen 1",
             "\t\t/isOn 1",
             "\t\t/hasDialog 0",
             "\t\t/parameterCount 1",
             "\t\t/parameter-1 {",
             "\t\t\t/key 1954115685",
             "\t\t\t/showInPalette 4294967295",
             "\t\t\t/type (enumerated)"],
            buildActionNameLines("\t\t\t", eventSpec.commandName),
            ["\t\t\t/value " + eventSpec.value,
             "\t\t}",
             "\t}"]
        );
    }

    /**
     * 選んだ方向の整列だけを含むアクション定義（.aia 形式）を作る
     * @param {{vertical: boolean, horizontal: boolean}} alignOptions - 整列する方向
     * @returns {string} アクション定義のテキスト
     */
    function buildAlignActionCode(alignOptions) {
        var eventSpecs = [];
        if (alignOptions.horizontal) eventSpecs.push(ALIGN_EVENT_SPECS.horizontal);
        if (alignOptions.vertical) eventSpecs.push(ALIGN_EVENT_SPECS.vertical);

        var actionLines = ["/version 3"].concat(
            buildActionNameLines("", ACTION_SET_NAME),
            ["/isOpen 1", "/actionCount 1", "/action-1 {"],
            buildActionNameLines("\t", ACTION_NAME),
            ["\t/keyIndex 0", "\t/colorIndex 0", "\t/isOpen 1", "\t/eventCount " + eventSpecs.length]
        );
        for (var i = 0; i < eventSpecs.length; i++) {
            actionLines = actionLines.concat(buildAlignEventLines(i + 1, eventSpecs[i]));
        }
        actionLines.push("}", "");
        return actionLines.join("\n");
    }

    // =========================================
    // 前回の設定 / Session memory
    // =========================================
    var ALIGN_OPTIONS_KEY = "__" + SCRIPT_NAME + "_AlignOptions"; /* $.global のキー / key on $.global */

    /**
     * 前回の整列方向を読む（無ければ DEFAULT_ALIGN_OPTIONS）
     * @returns {{vertical: boolean, horizontal: boolean}} 整列する方向
     */
    function loadAlignOptions() {
        var savedOptions = $.global[ALIGN_OPTIONS_KEY];
        var sourceOptions = savedOptions || DEFAULT_ALIGN_OPTIONS;
        return { vertical: sourceOptions.vertical === true, horizontal: sourceOptions.horizontal === true };
    }

    /**
     * 整列方向を次回のために覚える（Illustrator を終了するまで）
     * @param {{vertical: boolean, horizontal: boolean}} alignOptions - 整列する方向
     * @returns {void}
     */
    function saveAlignOptions(alignOptions) {
        $.global[ALIGN_OPTIONS_KEY] = { vertical: alignOptions.vertical, horizontal: alignOptions.horizontal };
    }

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

    /* カテゴリ分けした日英ラベル定義 / Categorized Japanese-English label definitions */
    var LABELS = {
        dialog: {
            title: { ja: "それぞれのアートボードに整列", en: "Align Each to Its Artboard" }
        },
        panel: {
            align: { ja: "整列", en: "Align" }
        },
        checkbox: {
            verticalCenter:   { ja: "天地中央", en: "Vertical Center" },
            horizontalCenter: { ja: "左右中央", en: "Horizontal Center" }
        },
        button: {
            cancel: { ja: "キャンセル", en: "Cancel" },
            ok:     { ja: "OK", en: "OK" }
        },
        alert: {
            noDocument:   { ja: "ドキュメントが開かれていません。", en: "No document is open." },
            noSelection:  { ja: "オブジェクトが選択されていません。", en: "No object is selected." },
            genericError: { ja: "エラーが発生しました：", en: "An error occurred: " }
        }
    };

    // =========================================
    // 環境設定 / Preferences
    // =========================================

    /**
     * 「字形の境界に整列」の現在の状態を取得する
     * @returns {{point: boolean, area: boolean}} ポイント文字・エリア内文字それぞれのON/OFF
     */
    function getGlyphBoundsAlign() {
        return {
            point: app.preferences.getBooleanPreference("EnableActualPointTextSpaceAlign") === true,
            area: app.preferences.getBooleanPreference("EnableActualAreaTextSpaceAlign") === true
        };
    }

    /**
     * 「字形の境界に整列」をポイント文字・エリア内文字それぞれに設定する
     * @param {{point: boolean, area: boolean}} glyphBoundsState - 設定するON/OFF
     * @returns {void}
     */
    function setGlyphBoundsAlign(glyphBoundsState) {
        app.preferences.setBooleanPreference("EnableActualPointTextSpaceAlign", glyphBoundsState.point);
        app.preferences.setBooleanPreference("EnableActualAreaTextSpaceAlign", glyphBoundsState.area);
    }

    // 一時アクション（再利用パーツ） / Temporary action (reusable)

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

    // 一時アクション（再利用パーツ）ここまで / End of the reusable temporary action

    // 選択の収集と境界（再利用パーツ） / Selection items and bounds (reusable)

    /* 座標を同じと見なす許容値（pt） / Tolerance for treating coordinates as equal, in points */
    var SELECTION_ITEMS_TOLERANCE = 0.001;

    /**
     * 選択やコレクションを、オブジェクトの配列にそろえる
     * TextRange・PathItem は length を持つので、typename で1個か集まりかを見分ける
     * @param {*} source - doc.selection、配列、DOM のコレクション、または単独のオブジェクト
     * @returns {Array} オブジェクトの配列（空なら []）
     */
    function normalizeSelectionItems(source) {
        var items = [];
        if (!source) return items;
        var typeName = "";
        try { typeName = source.typename || ""; } catch (e) { /* 読めない種類 / unreadable kind */ }
        /* 単数形の typename は1個（PageItems などのコレクションは s で終わる）
           A singular typename is one object (collections such as PageItems end in s) */
        if (typeName && !/s$/.test(typeName)) return [source];
        if (typeof source.length !== "number") return items;
        for (var i = 0; i < source.length; i++) items.push(source[i]);
        return items;
    }

    /**
     * 文字カーソルの選択（TextRange）を、それを含むテキストフレームに読み替える
     * @param {TextRange} textRange - 文字の範囲
     * @returns {TextFrame|null} テキストフレーム（たどれなければ null）
     */
    function resolveTextRangeFrame(textRange) {
        var current = textRange;
        /* parent をたどる（深さは念のため制限） / Walk up the parents, with a safety limit */
        for (var depth = 0; depth < 10 && current; depth++) {
            try {
                if (current.typename === "TextFrame") return current;
                current = current.parent;
            } catch (e) {
                break;
            }
        }
        /* ストーリーの先頭フレームで代用する / Fall back to the first frame of the story */
        try {
            var storyFrames = textRange.story.textFrames;
            if (storyFrames.length > 0) return storyFrames[0];
        } catch (e2) { /* ストーリーを持たない / no story */ }
        return null;
    }

    /**
     * 選択から条件に合うオブジェクトを集める（グループ・レイヤーを再帰でたどり、重複は除く）
     * 条件に合ったオブジェクトの中へは進まない
     * @param {*} source - doc.selection、配列、コレクション、または単独のオブジェクト
     * @param {Object} [options] - 収集の設定
     * @param {function(PageItem): boolean} [options.accept] - 集める条件（既定はグループ・レイヤー以外すべて）
     * @param {boolean} [options.enterGroups] - グループの中をたどる（既定 true）
     * @param {boolean} [options.enterClipGroups] - クリップグループの中をたどる（既定は enterGroups と同じ）
     * @param {boolean} [options.enterCompoundPaths] - 複合パスの中のパスをたどる（既定 false）
     * @param {boolean} [options.textRangeToFrame] - 文字の選択をテキストフレームに読み替える（既定 true）
     * @param {boolean} [options.skipLocked] - ロックされたものを中ごと外す（既定 false）
     * @param {boolean} [options.skipHidden] - 非表示のものを中ごと外す（既定 false）
     * @param {boolean} [options.skipClipMasks] - クリッピングマスクを外す（既定 false）
     * @param {boolean} [options.skipGuides] - ガイドを外す（既定 false）
     * @param {boolean} [options.unique] - 同じ参照を1回だけにする（既定 true。数千件で遅ければ false）
     * @returns {Array} 集めたオブジェクト（前面→背面の順）
     */
    function collectSelectionItems(source, options) {
        var opts = options || {};
        var enterGroups = (opts.enterGroups !== false);
        var enterClipGroups = (opts.enterClipGroups === undefined) ? enterGroups : (opts.enterClipGroups === true);
        var accept = opts.accept || function (item) {
            return item.typename !== "GroupItem" && item.typename !== "Layer";
        };
        var collected = [];

        /**
         * 集めた配列に加える（unique のときは同じ参照を足さない）
         * @param {PageItem} item - 加えるオブジェクト
         * @returns {void}
         */
        function pushItem(item) {
            if (opts.unique !== false) {
                for (var k = 0; k < collected.length; k++) {
                    if (collected[k] === item) return;
                }
            }
            collected.push(item);
        }

        /**
         * 設定に従って外すオブジェクトか判定する
         * @param {PageItem} item - 判定するオブジェクト
         * @returns {boolean} 外すなら true
         */
        function isSkipped(item) {
            try {
                if (item.typename === "Layer") {
                    if (opts.skipLocked && item.locked) return true;
                    if (opts.skipHidden && !item.visible) return true;
                    return false;
                }
                if (opts.skipLocked && item.locked) return true;
                if (opts.skipHidden && item.hidden) return true;
                if (opts.skipGuides && item.guides === true) return true;
                if (opts.skipClipMasks && isClipMaskItem(item)) return true;
            } catch (e) {
                /* 読めないプロパティは「外さない」に倒す / Unreadable properties do not exclude */
            }
            return false;
        }

        /**
         * 1件をたどって集める
         * @param {PageItem} item - 対象のオブジェクト
         * @returns {void}
         */
        function visit(item) {
            if (!item) return;
            var typeName = "";
            try { typeName = item.typename; } catch (e) { return; }

            if (typeName === "TextRange" || typeName === "InsertionPoint") {
                if (opts.textRangeToFrame === false) {
                    if (accept(item)) pushItem(item);
                    return;
                }
                visit(resolveTextRangeFrame(item));
                return;
            }
            if (isSkipped(item)) return;
            if (accept(item)) {
                pushItem(item);
                return;
            }

            var children = null;
            if (typeName === "GroupItem") {
                var isClipped = false;
                try { isClipped = (item.clipped === true); } catch (e2) { }
                if (isClipped ? enterClipGroups : enterGroups) children = item.pageItems;
            } else if (typeName === "CompoundPathItem") {
                if (opts.enterCompoundPaths) children = item.pathItems;
            } else if (typeName === "Layer") {
                /* 重なり順はサブレイヤーとページアイテムで別々なので、ページアイテム→サブレイヤーの順にする
                   Page items and sublayers stack separately; visit page items first, then sublayers */
                walk(item.pageItems);
                walk(item.layers);
                return;
            }
            if (children) walk(children);
        }

        /**
         * 集まりの各要素をたどる
         * @param {*} list - 配列またはコレクション
         * @returns {void}
         */
        function walk(list) {
            var listItems = normalizeSelectionItems(list);
            for (var i = 0; i < listItems.length; i++) visit(listItems[i]);
        }

        walk(source);
        return collected;
    }

    /**
     * テキストフレームの種類を "point" / "area" / "path" で返す
     * @param {TextFrame} textFrame - テキストフレーム
     * @returns {string} 種類のキー（判定できなければ ""）
     */
    function getTextFrameKindKey(textFrame) {
        try {
            if (textFrame.kind === TextType.POINTTEXT) return "point";
            if (textFrame.kind === TextType.AREATEXT) return "area";
            if (textFrame.kind === TextType.PATHTEXT) return "path";
        } catch (e) { /* kind を読めない / kind is unreadable */ }
        return "";
    }

    /**
     * 選択からテキストフレームを集める（グループの中・文字カーソルの選択を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string[]} [options.kinds] - 集める種類（"point" / "area" / "path"。既定はすべて）
     * @returns {TextFrame[]} テキストフレーム（前面→背面の順）
     */
    function collectSelectionTextFrames(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var kindFilter = null;
        if (opts.kinds && opts.kinds.length) {
            kindFilter = {};
            for (var i = 0; i < opts.kinds.length; i++) kindFilter[opts.kinds[i]] = true;
        }
        opts.accept = function (item) {
            if (item.typename !== "TextFrame") return false;
            return !kindFilter || kindFilter[getTextFrameKindKey(item)] === true;
        };
        /* 種類で外したテキストは中をたどらない（accept が false でも子は無い） / Text frames have no children to walk */
        return collectSelectionItems(source, opts);
    }

    /**
     * 選択からパスを集める（グループの中を含む）
     * @param {*} source - doc.selection など
     * @param {Object} [options] - collectSelectionItems と同じ設定に加えて次を受ける
     * @param {string} [options.compoundPaths] - 複合パスの扱い。"children"（中のパス、既定）/ "whole"（複合パスごと）/ "skip"（外す）
     * @returns {Array} PathItem（"whole" のときは CompoundPathItem も）の配列
     */
    function collectSelectionPathItems(source, options) {
        var opts = {};
        var sourceOptions = options || {};
        for (var key in sourceOptions) {
            if (sourceOptions.hasOwnProperty(key)) opts[key] = sourceOptions[key];
        }
        var compoundMode = opts.compoundPaths || "children";
        opts.enterCompoundPaths = (compoundMode === "children");
        opts.accept = function (item) {
            if (item.typename === "PathItem") return true;
            return compoundMode === "whole" && item.typename === "CompoundPathItem";
        };
        return collectSelectionItems(source, opts);
    }

    /**
     * クリッピングマスク（クリップグループの型）か判定する
     * パスは clipping、複合パスは中の先頭パスの clipping、テキストは clipping が無いので「クリップグループの先頭」で見る
     * @param {PageItem} item - 判定するオブジェクト
     * @returns {boolean} マスクなら true
     */
    function isClipMaskItem(item) {
        try {
            if (item.typename === "PathItem") return item.clipping === true;
            if (item.typename === "CompoundPathItem") {
                return item.pathItems.length > 0 && item.pathItems[0].clipping === true;
            }
            if (item.typename === "TextFrame") {
                var parentGroup = item.parent;
                return parentGroup.typename === "GroupItem" && parentGroup.clipped === true &&
                    parentGroup.pageItems.length > 0 && parentGroup.pageItems[0] === item;
            }
        } catch (e) { /* 読めない種類はマスクではない / unreadable kinds are not masks */ }
        return false;
    }

    /**
     * クリップグループの型（マスク）を返す
     * フラグで探し、見つからなければ先頭（pageItems[0]）を返す（型は常に最前面。テキストの型はフラグを持たない）
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {PageItem|null} マスク（クリップグループでなければ null）
     */
    function getClipMaskItem(groupItem) {
        try {
            if (!groupItem || groupItem.typename !== "GroupItem" || groupItem.clipped !== true) return null;
            var groupChildren = groupItem.pageItems;
            if (groupChildren.length === 0) return null;
            for (var i = 0; i < groupChildren.length; i++) {
                var childType = groupChildren[i].typename;
                if ((childType === "PathItem" || childType === "CompoundPathItem") && isClipMaskItem(groupChildren[i])) {
                    return groupChildren[i];
                }
            }
            return groupChildren[0];
        } catch (e) {
            return null;
        }
    }

    /**
     * グループの中（入れ子を含む）にクリップグループがあるか判定する
     * @param {GroupItem} groupItem - 対象のグループ
     * @returns {boolean} あれば true
     */
    function hasClippedDescendant(groupItem) {
        try {
            var groupChildren = groupItem.pageItems;
            for (var i = 0; i < groupChildren.length; i++) {
                if (groupChildren[i].typename !== "GroupItem") continue;
                if (groupChildren[i].clipped === true || hasClippedDescendant(groupChildren[i])) return true;
            }
        } catch (e) { /* 中を読めない / cannot read the children */ }
        return false;
    }

    /**
     * 環境設定の［プレビュー境界を使用］を読む
     * @returns {boolean} オンなら true（読めなければ false）
     */
    function readUsePreviewBoundsPreference() {
        try {
            return app.preferences.getBooleanPreference("includeStrokeInBounds");
        } catch (e) {
            return false;
        }
    }

    /**
     * 見た目どおりの境界を返す。クリップグループはマスクの境界、
     * 中にクリップグループを含むグループは子の境界を合わせたもの（隠れた部分を含めない）
     * @param {PageItem} item - 対象のオブジェクト
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下] の新しい配列（測れなければ null）
     */
    function getClipAwareBounds(item, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        try {
            var measuredItem = item;
            if (item.typename === "GroupItem") {
                var maskItem = getClipMaskItem(item);
                if (maskItem) {
                    measuredItem = maskItem;
                } else if (hasClippedDescendant(item)) {
                    /* グループ自体の効果（影など）の広がりは含まれなくなる
                       This leaves out the reach of effects applied to the group itself (drop shadows etc.) */
                    var childBounds = getClipAwareUnionBounds(filterMeasurableChildren(item.pageItems), usePreview);
                    if (childBounds) return childBounds;
                }
            }
            var bounds = usePreview ? measuredItem.visibleBounds : measuredItem.geometricBounds;
            return [bounds[0], bounds[1], bounds[2], bounds[3]];
        } catch (e) {
            return null;
        }
    }

    /**
     * 境界の計算に入れる子だけを残す（非表示とガイドを外す）
     * @param {*} childList - 子のコレクション
     * @returns {Array} 残した子
     */
    function filterMeasurableChildren(childList) {
        var childItems = normalizeSelectionItems(childList);
        var measurable = [];
        for (var i = 0; i < childItems.length; i++) {
            try {
                if (childItems[i].hidden === true || childItems[i].guides === true) continue;
            } catch (e) { /* 読めなければ残す / keep when unreadable */ }
            measurable.push(childItems[i]);
        }
        return measurable;
    }

    /**
     * 複数のオブジェクトを囲む外接範囲を返す（クリップグループはマスクで測る）
     * @param {*} items - オブジェクトの配列・コレクション・選択
     * @param {boolean} [usePreviewBounds] - true で visibleBounds、false で geometricBounds（省略時は環境設定に従う）
     * @returns {number[]|null} [左, 上, 右, 下]（測れるものが無ければ null）
     */
    function getClipAwareUnionBounds(items, usePreviewBounds) {
        var usePreview = (usePreviewBounds === undefined || usePreviewBounds === null) ?
            readUsePreviewBoundsPreference() : (usePreviewBounds === true);
        var itemList = normalizeSelectionItems(items);
        var unionBounds = null;
        for (var i = 0; i < itemList.length; i++) {
            var itemBounds = getClipAwareBounds(itemList[i], usePreview);
            if (!itemBounds) continue;
            if (!unionBounds) {
                unionBounds = itemBounds;
                continue;
            }
            if (itemBounds[0] < unionBounds[0]) unionBounds[0] = itemBounds[0];
            if (itemBounds[1] > unionBounds[1]) unionBounds[1] = itemBounds[1];
            if (itemBounds[2] > unionBounds[2]) unionBounds[2] = itemBounds[2];
            if (itemBounds[3] < unionBounds[3]) unionBounds[3] = itemBounds[3];
        }
        return unionBounds;
    }

    /**
     * 2つの座標を許容値つきで比べる
     * @param {number} valueA - 座標A（pt）
     * @param {number} valueB - 座標B（pt）
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 差が許容値以下なら true
     */
    function isNearlySameCoordinate(valueA, valueB, tolerance) {
        var limit = (typeof tolerance === "number") ? tolerance : SELECTION_ITEMS_TOLERANCE;
        return Math.abs(valueA - valueB) <= limit;
    }

    /**
     * 2つの境界を許容値つきで比べる
     * @param {number[]} boundsA - [左, 上, 右, 下]
     * @param {number[]} boundsB - [左, 上, 右, 下]
     * @param {number} [tolerance] - 許容値（pt、既定は SELECTION_ITEMS_TOLERANCE）
     * @returns {boolean} 4辺とも許容値以内なら true
     */
    function areBoundsNearlyEqual(boundsA, boundsB, tolerance) {
        if (!boundsA || !boundsB) return false;
        for (var i = 0; i < 4; i++) {
            if (!isNearlySameCoordinate(boundsA[i], boundsB[i], tolerance)) return false;
        }
        return true;
    }

    // 選択の収集と境界（再利用パーツ）ここまで / End of the reusable selection items and bounds

    // 字面の中央揃え（再利用パーツ） / Optical center kerning (reusable)

    /* 既定値。options で上書きできる / Defaults; override them with options */
    var OPTICAL_CENTER_KERNING_DEFAULTS = {
        strength: 35,   /* 補正の強さ（%）。100で字面の中心をぴったり揃える / correction strength (%); 100 centers the glyphs exactly */
        minKerning: 1   /* これ未満の補正は0にする（1/1000 em） / corrections smaller than this become 0 (1/1000 em) */
    };

    /**
     * 本文を改行で区切り、行ごとの文字番号の範囲を返す
     * @param {TextFrame} textFrame - 対象のテキスト
     * @returns {Array} [{ start: 行頭の文字番号, end: 改行の位置（含まない） }, …]
     */
    function getTextLineRanges(textFrame) {
        var lineTexts = textFrame.contents.split("\r");
        var ranges = [];
        var start = 0;
        for (var i = 0; i < lineTexts.length; i++) {
            ranges.push({ start: start, end: start + lineTexts[i].length });
            start += lineTexts[i].length + 1;
        }
        return ranges;
    }

    /**
     * 1行分だけを残した複製をアウトライン化し、字面の左右を測る
     * @param {TextFrame} textFrame - 対象のポイント文字
     * @param {number} lineStart - 行頭の文字番号
     * @param {number} lineEnd - 行末の次（改行の位置）の文字番号
     * @returns {Array|null} [左, 右]。字形が無ければ null
     */
    function measureLineInk(textFrame, lineStart, lineEnd) {
        var dup = textFrame.duplicate();

        /* 対象行以外の文字を右から消す / Remove the other lines' characters from the right */
        for (var i = dup.characters.length - 1; i >= lineEnd; i--) {
            dup.characters[i].remove();
        }
        for (var j = lineStart - 1; j >= 0; j--) {
            dup.characters[j].remove();
        }

        /* 補正前の状態で測る / Measure without the existing correction */
        try { dup.characters[0].kerning = 0; } catch (e) {}

        /* createOutline は複製を消費する / createOutline consumes the duplicate */
        var outline = dup.createOutline();
        var bounds = outline.geometricBounds;
        var hasGlyphs = outline.pageItems.length > 0;
        outline.remove();

        return hasGlyphs ? [bounds[0], bounds[2]] : null;
    }

    /**
     * 字面の中心を揃えの基準に寄せるカーニング値（1/1000 em）を求める
     * 中央揃えで行頭に k を入れると行幅が k 伸び、字面は k/2 右へ動く
     * @param {number} offset - 基準X − 字面の中心（pt）
     * @param {number} fontSize - 行頭の文字サイズ（pt）
     * @param {number} strength - 補正の強さ（%）
     * @returns {number} 強さを掛けて丸めたカーニング値
     */
    function calcOpticalCenterKerning(offset, fontSize, strength) {
        return Math.round(2 * offset / fontSize * 1000 * strength / 100);
    }

    /**
     * 中央揃えの行ごとに、行頭へ入れるカーニング値を求める（テキストは変えない）
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {Object} [options] - { strength, minKerning }。省いた値は OPTICAL_CENTER_KERNING_DEFAULTS
     * @returns {Array} [{ lineStart: 行頭の文字番号, kerning: 値 }, …]。対象外のテキストは []
     */
    function getOpticalCenterKernings(textFrame, options) {
        var results = [];
        if (textFrame.kind !== TextType.POINTTEXT) return results;

        var opts = options || {};
        var strength = (opts.strength != null) ? opts.strength : OPTICAL_CENTER_KERNING_DEFAULTS.strength;
        var minKerning = (opts.minKerning != null) ? opts.minKerning : OPTICAL_CENTER_KERNING_DEFAULTS.minKerning;

        var characters = textFrame.characters;
        var centerX = textFrame.anchor[0];
        var ranges = getTextLineRanges(textFrame);

        for (var i = 0; i < ranges.length; i++) {
            var range = ranges[i];

            /* 空行と中央揃え以外の行は飛ばす / Skip empty lines and lines that are not centered */
            if (range.end <= range.start) continue;
            if (characters[range.start].paragraphAttributes.justification !== Justification.CENTER) continue;

            var ink = measureLineInk(textFrame, range.start, range.end);
            if (!ink) continue;

            var fontSize = characters[range.start].characterAttributes.size;
            var kerning = calcOpticalCenterKerning(centerX - (ink[0] + ink[1]) / 2, fontSize, strength);
            results.push({ lineStart: range.start, kerning: (Math.abs(kerning) < minKerning) ? 0 : kerning });
        }
        return results;
    }

    /**
     * 中央揃えの行ごとに、行頭のカーニングで字面の中心を補正する
     * @param {TextFrame} textFrame - 対象のテキスト
     * @param {Object} [options] - getOpticalCenterKernings と同じ
     * @returns {number} 補正した行数
     */
    function applyOpticalCenterKerning(textFrame, options) {
        /* 先にすべての行を測ってから書き込む / Measure every line before writing */
        var kernings = getOpticalCenterKernings(textFrame, options);
        for (var i = 0; i < kernings.length; i++) {
            textFrame.characters[kernings[i].lineStart].kerning = kernings[i].kerning;
        }
        return kernings.length;
    }

    // 字面の中央揃え（再利用パーツ）ここまで / End of the reusable optical center kerning

    // =========================================
    // アートボード / Artboards
    // =========================================

    /**
     * 2つの矩形が重なっている面積を求める
     * @param {number[]} boundsA - [左, 上, 右, 下] の座標
     * @param {number[]} boundsB - [左, 上, 右, 下] の座標
     * @returns {number} 重なっている面積（重ならない場合は 0）
     */
    function getOverlapArea(boundsA, boundsB) {
        var overlapWidth = Math.min(boundsA[2], boundsB[2]) - Math.max(boundsA[0], boundsB[0]);
        var overlapHeight = Math.min(boundsA[1], boundsB[1]) - Math.max(boundsA[3], boundsB[3]);
        if (overlapWidth <= 0 || overlapHeight <= 0) {
            return 0;
        }
        return overlapWidth * overlapHeight;
    }

    /**
     * 選択範囲と最も広く重なるアートボードを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} selectionBounds - [左, 上, 右, 下] の座標
     * @param {number} currentIndex - 重なりが同じときに優先するアートボード番号
     * @returns {number} アートボード番号（どこにも重ならない場合は -1）
     */
    function findOverlappingArtboardIndex(doc, selectionBounds, currentIndex) {
        /* 現在のアートボードを先に見て、重なりが同じなら切り替えない / Check the current artboard first so ties keep it */
        var searchOrder = [currentIndex];
        for (var i = 0; i < doc.artboards.length; i++) {
            if (i !== currentIndex) {
                searchOrder.push(i);
            }
        }
        var largestOverlapIndex = -1;
        var largestOverlapArea = 0;
        for (var j = 0; j < searchOrder.length; j++) {
            var overlapArea = getOverlapArea(selectionBounds, doc.artboards[searchOrder[j]].artboardRect);
            if (overlapArea > largestOverlapArea) {
                largestOverlapArea = overlapArea;
                largestOverlapIndex = searchOrder[j];
            }
        }
        return largestOverlapIndex;
    }

    /**
     * 選択範囲の中心に最も近いアートボードを探す
     * @param {Document} doc - 対象ドキュメント
     * @param {number[]} selectionBounds - [左, 上, 右, 下] の座標
     * @returns {number} アートボード番号
     */
    function findNearestArtboardIndex(doc, selectionBounds) {
        var centerX = (selectionBounds[0] + selectionBounds[2]) / 2;
        var centerY = (selectionBounds[1] + selectionBounds[3]) / 2;
        var nearestIndex = 0;
        var nearestDistance = null;
        for (var i = 0; i < doc.artboards.length; i++) {
            var artboardRect = doc.artboards[i].artboardRect;
            var offsetX = centerX - (artboardRect[0] + artboardRect[2]) / 2;
            var offsetY = centerY - (artboardRect[1] + artboardRect[3]) / 2;
            var distance = offsetX * offsetX + offsetY * offsetY;
            if (nearestDistance === null || distance < nearestDistance) {
                nearestDistance = distance;
                nearestIndex = i;
            }
        }
        return nearestIndex;
    }

    /**
     * 選択が現在のアートボード上にないとき、選択を含むアートボードを現在のアートボードにする
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function activateArtboardForSelection(doc, selectedItems) {
        if (doc.artboards.length < 2) {
            return;
        }
        var selectionBounds = getClipAwareUnionBounds(selectedItems, true);
        if (!selectionBounds) return; /* 測れる選択が無い / Nothing measurable */
        var currentIndex = doc.artboards.getActiveArtboardIndex();
        var targetIndex = findOverlappingArtboardIndex(doc, selectionBounds, currentIndex);
        /* どのアートボードにも重ならないときは一番近いアートボードを使う / Fall back to the nearest artboard */
        if (targetIndex < 0) {
            targetIndex = findNearestArtboardIndex(doc, selectionBounds);
        }
        if (targetIndex !== currentIndex) {
            doc.artboards.setActiveArtboardIndex(targetIndex);
        }
    }

    // =========================================
    // 選択の検査 / Selection checks
    // =========================================

    /**
     * 文字を部分選択している場合に、その文字を含むテキストオブジェクトを選択し直す
     * @param {Document} doc - 対象ドキュメント
     * @returns {PageItem[]} 選択し直したあとの選択内容
     */
    function selectTextFrameFromTextRange(doc) {
        var storyFrames = doc.selection.story.textFrames;
        var targetFrames = [];
        for (var i = 0; i < storyFrames.length; i++) {
            targetFrames.push(storyFrames[i]);
        }
        /* 文字編集を抜けてからテキストオブジェクトを選択 / Leave text editing, then select the frames */
        app.executeMenuCommand("deselectall");
        for (var j = 0; j < targetFrames.length; j++) {
            targetFrames[j].selected = true;
        }
        return doc.selection;
    }

    // =========================================
    // テキスト / Text
    // =========================================

    /**
     * 選択が1行だけのテキストオブジェクト1つか判定する
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {boolean} 1行だけのテキストオブジェクト1つなら true
     */
    function isSingleLineTextFrame(selectedItems) {
        if (selectedItems.length !== 1 || selectedItems[0].typename !== "TextFrame") {
            return false;
        }
        /* 折り返しも含めた実際の行数で判定 / Count the rendered lines, wrapping included */
        return selectedItems[0].lines.length === 1;
    }

    /**
     * テキストオブジェクト全体の行揃えを中央揃えにする
     * @param {TextFrame} textFrame - 対象のテキストオブジェクト
     * @returns {void}
     */
    function setCenterJustification(textFrame) {
        textFrame.textRange.paragraphAttributes.justification = Justification.CENTER;
    }

    /**
     * 選択に含まれるポイント文字の中央揃えの行を、字面の中心に寄せる
     * @param {PageItem[]} selectedItems - 選択中のオブジェクト
     * @returns {void}
     */
    function applyOpticalCenterKerningToSelection(selectedItems) {
        var textFrames = collectSelectionTextFrames(selectedItems, { kinds: ["point"], skipLocked: true, skipHidden: true });
        for (var i = 0; i < textFrames.length; i++) {
            applyOpticalCenterKerning(textFrames[i]);
        }
    }

    // ダイアログの位置と不透明度（再利用パーツ） / Dialog position and opacity (reusable)

    var DIALOG_OPACITY = 0.98;       /* ダイアログの不透明度 / dialog opacity */
    var DIALOG_AVOID_MARGIN = 60;    /* 選択範囲の推定位置の両側に取る余裕（px）/ margin on each side of the estimated selection (px) */
    var DIALOG_AVOID_MAX_ITEMS = 100; /* 選択範囲を測るオブジェクトの上限 / max items measured for the selection bounds */

    /**
     * ダイアログの不透明度を設定し、前回閉じた位置で開いて、動かした位置を記録するようにする。
     * 開く位置が選択中のオブジェクトに重なりそうなときは、左右の反対側へずらす（Illustrator のみ）。
     * 既存の onShow / onMove / onClose は先に呼んでから、位置の復元・記録を行う。
     * @param {Window} dialog - 対象のダイアログ
     * @param {string} storageKey - 位置を覚えるキー（ふつうは SCRIPT_NAME）
     * @returns {void}
     */
    function prepareDialogWindow(dialog, storageKey) {
        /* 同じダイアログを開き直すときは、選択範囲を測り直すだけにする（ハンドラーを重ねない）
           When the same dialog is shown again, only re-measure the selection (don't stack handlers) */
        if (dialog.dialogWindowState) {
            dialog.dialogWindowState.selectionSpan = getSelectionViewSpan();
            dialog.dialogWindowState.avoidedLocation = null;
            return;
        }
        var locationKey = "__" + storageKey + "_DialogLocation";
        var previousOnShow = dialog.onShow;
        var previousOnMove = dialog.onMove;
        var previousOnClose = dialog.onClose;
        var windowState = {
            selectionSpan: getSelectionViewSpan(), /* 選択範囲は show() の前に測る / measured before show() */
            screenWidth: null,                     /* 最初に開いたときに推定する / estimated on the first show */
            avoidedLocation: null                  /* 避けるためにずらした位置（記録しない）/ location set to avoid the selection (not remembered) */
        };
        dialog.dialogWindowState = windowState;

        dialog.opacity = DIALOG_OPACITY;

        /* 今の位置を記録する / Remember the current location */
        function rememberDialogLocation() {
            var currentLocation = [dialog.location[0], dialog.location[1]];
            var avoidedLocation = windowState.avoidedLocation;
            if (avoidedLocation && currentLocation[0] === avoidedLocation[0] && currentLocation[1] === avoidedLocation[1]) return;
            $.global[locationKey] = currentLocation;
        }

        dialog.onShow = function () {
            /* 最初に開くときの既定の位置は画面の横中央なので、画面の幅を逆算できる。2回目からは前回の位置なので使い回す
               On the first show the default location is centered horizontally, which gives the screen width; reuse it afterwards */
            if (windowState.screenWidth === null) windowState.screenWidth = dialog.location[0] * 2 + dialog.bounds.width;
            if (previousOnShow) previousOnShow.apply(this, arguments);
            /* $.screens は実際の画面の大きさと合わない（Mac で 1280×524 など）ので、画面内かは判定しない
               $.screens does not match the real display (e.g. 1280x524 on a Mac), so no on-screen check */
            var savedLocation = $.global[locationKey];
            if (savedLocation) dialog.location = [savedLocation[0], savedLocation[1]];
            if (windowState.selectionSpan) {
                var avoidLeft = findDialogLeftAvoidingSelection(dialog.location[0], dialog.bounds.width, windowState.screenWidth, windowState.selectionSpan);
                if (avoidLeft !== null) {
                    dialog.location = [avoidLeft, dialog.location[1]];
                    /* 代入後の値で比べる（丸められることがある）/ Compare with the value after assignment, which may be rounded */
                    windowState.avoidedLocation = [dialog.location[0], dialog.location[1]];
                }
            }
        };
        dialog.onMove = function () {
            if (previousOnMove) previousOnMove.apply(this, arguments);
            rememberDialogLocation();
        };
        dialog.onClose = function () {
            rememberDialogLocation();
            /* false を返すと閉じるのを取りやめるので、戻り値は元の onClose のものを返す
               Returning false cancels the close, so pass the original onClose result through */
            if (previousOnClose) return previousOnClose.apply(this, arguments);
        };
    }

    /**
     * 選択中のオブジェクトが、ドキュメントの表示域の左端から画面上で何 px の範囲にあるかを返す。
     * @returns {{left: number, right: number, viewWidth: number}|null} 選択が無い・測れないときは null
     */
    function getSelectionViewSpan() {
        try {
            if (app.name !== "Adobe Illustrator" || !app.documents.length) return null;
            var targetDoc = app.activeDocument;
            var selectedItems = targetDoc.selection;
            /* 文字ツールで文字を選択しているときは TextRange が返り、[0] が無い / Selecting characters with the Type tool returns a TextRange, which has no [0] */
            if (!selectedItems || selectedItems.typename === "TextRange" || !selectedItems.length || !selectedItems[0].visibleBounds) return null;
            var itemCount = Math.min(selectedItems.length, DIALOG_AVOID_MAX_ITEMS);
            var spanLeft = Infinity;
            var spanRight = -Infinity;
            for (var i = 0; i < itemCount; i++) {
                var itemBounds = selectedItems[i].visibleBounds;
                if (itemBounds[0] < spanLeft) spanLeft = itemBounds[0];
                if (itemBounds[2] > spanRight) spanRight = itemBounds[2];
            }
            var activeView = targetDoc.activeView; /* 複数ウィンドウで開いていても今のウィンドウ / the current window even with multiple windows */
            var viewBounds = activeView.bounds;
            var zoom = activeView.zoom;
            var viewWidth = (viewBounds[2] - viewBounds[0]) * zoom;
            /* 表示域の外にはみ出した部分は数えない / Ignore the part outside the view */
            var left = Math.max(0, (spanLeft - viewBounds[0]) * zoom);
            var right = Math.min(viewWidth, (spanRight - viewBounds[0]) * zoom);
            if (right <= left) return null;
            return { left: left, right: right, viewWidth: viewWidth };
        } catch (e) {
            /* テキスト編集中など測れないときは避けない / Do not avoid when it cannot be measured, e.g. while editing text */
            return null;
        }
    }

    /**
     * ダイアログが選択範囲に重なるなら、重ならない左端の位置を返す。
     * 表示域は画面の横中央にあるとみなし、ずれは DIALOG_AVOID_MARGIN で吸収する。
     * @param {number} dialogLeft - 今のダイアログの左端
     * @param {number} dialogWidth - ダイアログの幅
     * @param {number} screenWidth - 画面の幅
     * @param {{left: number, right: number, viewWidth: number}} selectionSpan - getSelectionViewSpan() の結果
     * @returns {number|null} ずらした左端。重ならない・どちらにも収まらないときは null
     */
    function findDialogLeftAvoidingSelection(dialogLeft, dialogWidth, screenWidth, selectionSpan) {
        var viewLeft = (screenWidth - selectionSpan.viewWidth) / 2;
        var avoidLeft = viewLeft + selectionSpan.left - DIALOG_AVOID_MARGIN;
        var avoidRight = viewLeft + selectionSpan.right + DIALOG_AVOID_MARGIN;
        if (dialogLeft + dialogWidth <= avoidLeft || dialogLeft >= avoidRight) return null;

        var leftSideLeft = avoidLeft - dialogWidth;   /* 選択範囲の左に置くとき / placed left of the selection */
        var rightSideLeft = avoidRight;               /* 選択範囲の右に置くとき / placed right of the selection */
        var fitsLeft = leftSideLeft >= 0;
        var fitsRight = rightSideLeft + dialogWidth <= screenWidth;
        /* 選択範囲が画面の右寄りなら左へ、左寄りなら右へ逃がす / Move away from the side the selection leans to */
        var preferLeft = (avoidLeft + avoidRight) / 2 > screenWidth / 2;
        if (preferLeft && fitsLeft) return leftSideLeft;
        if (fitsRight) return rightSideLeft;
        if (fitsLeft) return leftSideLeft;
        return null;
    }

    // ダイアログの位置と不透明度（再利用パーツ）ここまで / End of the reusable dialog position and opacity

    // ボタン行（再利用パーツ） / Button row (reusable)

    var BUTTON_ROW_TOP_MARGIN = 5; /* ボタン行の上の余白 / top margin of the button row */
    var BUTTON_ROW_BOTTOM_MARGIN = 14; /* ボタン行の下の余白。ダイアログの下余白と合わせて約30px（Illustrator 標準のダイアログに合わせる） / bottom margin; with the dialog margin about 30px, like Illustrator's own dialogs */
    var BUTTON_ROW_SPACING = 10;   /* ボタンどうしの間隔 / spacing between buttons */
    var BUTTON_ROW_CENTER_MAX_WIDTH = 200; /* 右のボタンだけの行を中央に置く、ダイアログの内側の最大幅（px、左右の余白を除く）。広いダイアログは右揃え / max inner dialog width (px, margins excluded) that centers a right-only row; wider dialogs keep it right-aligned */

    /**
     * ダイアログ下部のボタン行を作る。
     * 通常は「左のグループ・伸びるスペーサー・右のグループ」、centered なら行そのものを左右中央に置く
     * @param {Window|Group|Panel} parent - 行を足す先（ふつうはダイアログ）
     * @param {Object} [rowOptions] - { centered: true } で左右中央に並べる
     * @returns {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} 行と左右のグループ（centered のときは左右が null）
     */
    function addButtonRow(parent, rowOptions) {
        var isCentered = !!(rowOptions && rowOptions.centered);
        var btnRowGroup = parent.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, BUTTON_ROW_TOP_MARGIN, 0, BUTTON_ROW_BOTTOM_MARGIN];
        btnRowGroup.spacing = BUTTON_ROW_SPACING;

        if (isCentered) {
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            return { rowGroup: btnRowGroup, leftGroup: null, rightGroup: null };
        }

        btnRowGroup.alignment = ["fill", "bottom"];

        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        btnLeftGroup.spacing = BUTTON_ROW_SPACING;

        /* 余りの幅を吸って、右のグループを右端に寄せる / Absorbs the extra width so the right group sits at the right edge */
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.spacing = BUTTON_ROW_SPACING;

        return { rowGroup: btnRowGroup, leftGroup: btnLeftGroup, rightGroup: btnRightGroup };
    }

    /**
     * 左のグループにボタンが無い（右のボタンだけの）行を、ダイアログの幅に合わせて揃える。
     * 内側の幅（左右の余白を除く）が BUTTON_ROW_CENTER_MAX_WIDTH 以下なら左右中央、それより広ければ右揃えのまま。
     * 幅はレイアウトが決まるまで分からないので、ダイアログを表示した時点（show イベント）で判定する。
     * ボタンをすべて足したあと、show() の前に呼ぶ。centered で作った行や、左にボタンがある行はそのまま
     * @param {{rowGroup: Group, leftGroup: Group|null, rightGroup: Group|null}} buttonRow - addButtonRow() の戻り値
     * @returns {void}
     */
    function alignRightOnlyButtonRow(buttonRow) {
        if (!buttonRow.leftGroup || buttonRow.leftGroup.children.length > 0) return;
        var dialogWindow = buttonRow.rowGroup.window;
        dialogWindow.addEventListener("show", function () {
            if (!buttonRow.leftGroup) return;
            var btnRowGroup = buttonRow.rowGroup;
            /* 行の幅＝ダイアログの内側の幅（左右の余白を除く）/ The row spans the dialog's inner width (margins excluded) */
            if (!btnRowGroup.size || btnRowGroup.size.width > BUTTON_ROW_CENTER_MAX_WIDTH) return;
            /* 左のグループとスペーサーを外し、右のグループだけを中央に置く / Drop the left group and the spacer so only the right group remains, centered */
            btnRowGroup.remove(buttonRow.leftGroup);
            btnRowGroup.remove(btnRowGroup.children[0]); /* 左のグループを外すと先頭はスペーサー / the spacer is first once the left group is gone */
            btnRowGroup.alignment = ["center", "bottom"];
            btnRowGroup.alignChildren = ["center", "center"];
            buttonRow.leftGroup = null;
            dialogWindow.layout.layout(true);
        });
    }

    // ボタン行（再利用パーツ）ここまで / End of the reusable button row

    // =========================================
    // ダイアログ / Dialog
    // =========================================

    /**
     * 整列する方向を選ぶダイアログを表示する
     * @param {{vertical: boolean, horizontal: boolean}} initialOptions - 初期値
     * @returns {{vertical: boolean, horizontal: boolean}|null} 選んだ方向（キャンセルなら null）
     */
    function showAlignDialog(initialOptions) {
        var alignDialog = new Window("dialog", getLabel("dialog.title") + " " + SCRIPT_VERSION);
        setupWindow(alignDialog);

        var alignPanel = alignDialog.add("panel", undefined, getLabel("panel.align"));
        setupPanel(alignPanel, 6);
        var verticalCheckbox = alignPanel.add("checkbox", undefined, getLabel("checkbox.verticalCenter"));
        var horizontalCheckbox = alignPanel.add("checkbox", undefined, getLabel("checkbox.horizontalCenter"));
        verticalCheckbox.value = initialOptions.vertical;
        horizontalCheckbox.value = initialOptions.horizontal;

        var buttonRow = addButtonRow(alignDialog);
        buttonRow.rightGroup.add("button", undefined, getLabel("button.cancel"), { name: "cancel" });
        var btnOk = buttonRow.rightGroup.add("button", undefined, getLabel("button.ok"), { name: "ok" });
        alignRightOnlyButtonRow(buttonRow);

        /* どちらもOFFなら OK を押せない。表示前の value は読み戻せないので初期値で判定する
           OK needs at least one direction; before show() the values cannot be read back, so use the initial options */
        btnOk.enabled = initialOptions.vertical || initialOptions.horizontal;
        verticalCheckbox.onClick = horizontalCheckbox.onClick = function () {
            btnOk.enabled = verticalCheckbox.value || horizontalCheckbox.value;
        };

        prepareDialogWindow(alignDialog, SCRIPT_NAME);
        if (alignDialog.show() !== 1) return null;
        return { vertical: verticalCheckbox.value, horizontal: horizontalCheckbox.value };
    }

    // =========================================
    // 整列 / Alignment
    // =========================================

    /**
     * 1つのオブジェクトだけを選択し、それが載っているアートボードを現在のアートボードにしてアクションを実行する
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem} targetItem - 整列するオブジェクト
     * @returns {void}
     */
    function alignItemToItsArtboard(doc, targetItem) {
        doc.selection = null;
        targetItem.selected = true;
        activateArtboardForSelection(doc, [targetItem]);
        app.doScript(ACTION_NAME, ACTION_SET_NAME);
    }

    /**
     * 選択を元のオブジェクトに戻す（消えたオブジェクトは飛ばす）
     * @param {Document} doc - 対象ドキュメント
     * @param {PageItem[]} targetItems - 選択し直すオブジェクト
     * @returns {void}
     */
    function restoreSelection(doc, targetItems) {
        doc.selection = null;
        for (var i = 0; i < targetItems.length; i++) {
            try { targetItems[i].selected = true; } catch (e) { /* 選択できない / cannot be selected */ }
        }
    }

    // =========================================
    // メイン処理 / Main
    // =========================================

    /**
     * ドキュメントと選択を確認し、ダイアログで方向を選んでから、オブジェクトごとにそれぞれのアートボードへ整列する
     * 左右中央のときは、1行だけのテキストオブジェクトの行揃えも中央揃えにする
     * @returns {void}
     */
    function main() {
        if (app.documents.length === 0) {
            alert(getLabel("alert.noDocument"));
            return;
        }
        var doc = app.activeDocument;
        var selectedItems = doc.selection;
        /* 文字を部分選択しているときは selection が TextRange になるため、テキストオブジェクトに置き換える
           A partial text selection comes back as a TextRange; promote it to the text object */
        if (selectedItems && !(selectedItems instanceof Array)) {
            selectedItems = selectTextFrameFromTextRange(doc);
        }
        if (!(selectedItems instanceof Array) || selectedItems.length === 0) {
            alert(getLabel("alert.noSelection"));
            return;
        }

        var alignOptions = showAlignDialog(loadAlignOptions());
        if (!alignOptions) return;
        saveAlignOptions(alignOptions);

        /* 選択し直すと doc.selection が変わるので先に控える / Copy first; reselecting changes doc.selection */
        var targetItems = [];
        for (var i = 0; i < selectedItems.length; i++) {
            targetItems.push(selectedItems[i]);
        }
        var previousArtboardIndex = doc.artboards.getActiveArtboardIndex();

        /* 実行中だけONにして、終了時に元の状態へ戻す / Turn on for this run only, then restore */
        var previousGlyphBounds = getGlyphBoundsAlign();
        try {
            setGlyphBoundsAlign({ point: true, area: true });
            if (alignOptions.horizontal) {
                /* 1行だけのテキストは行揃えも中央揃えにする / Single-line text objects get centered justification too */
                for (var j = 0; j < targetItems.length; j++) {
                    if (isSingleLineTextFrame([targetItems[j]])) {
                        setCenterJustification(targetItems[j]);
                    }
                }
                /* 中央揃えの行は、行末の約物などで寄って見える字面を行頭のカーニングで戻す
                   Pull centered lines back toward the center with kerning at the line start */
                applyOpticalCenterKerningToSelection(targetItems);
            }
            /* アクションは1回だけ読み込み、オブジェクトごとに実行する / Load the action once and run it per object */
            if (!loadTemporaryActionSet(buildAlignActionCode(alignOptions), ACTION_SET_NAME)) {
                throw new Error("Could not load the action \"" + ACTION_NAME + "\".");
            }
            for (var k = 0; k < targetItems.length; k++) {
                alignItemToItsArtboard(doc, targetItems[k]);
            }
        } catch (e) {
            alert(getLabel("alert.genericError") + e);
        } finally {
            unloadTemporaryActionSet(ACTION_SET_NAME);
            setGlyphBoundsAlign(previousGlyphBounds);
            restoreSelection(doc, targetItems);
            doc.artboards.setActiveArtboardIndex(previousArtboardIndex);
        }
    }

    main();

})();
