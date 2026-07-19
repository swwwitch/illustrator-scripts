#target illustrator
#targetengine "ImportAndApplyGraphicStyle"
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### 概要

ポイント文字・パス上文字・図形＋テキストを、見た目を保ったままエリア内文字へ変換する。

処理の流れ：

1. 実行時に常駐パレットを表示（大きさ調整の する/しない と 幅%・高さ%、グラフィックスタイル：元の見た目／読み込んだスタイル）。「変換」ボタンで実行し、パレットは開いたまま
2. 変換前に、適用するグラフィックスタイルを用意（「元の見た目」＝選択テキストの見た目を一時登録／読み込んだスタイル＝「読み込み」で選んだ AI ファイルから現書類へ取り込み）
3. ポイント文字 / パス上文字 → 計測した実寸（大きさ調整ONなら幅・高さに倍率）でエリア内文字へ変換
   テキスト＋長方形 → 長方形を複製してエリア内文字にし、テキストを流し込む
4. ボタン背景（テキストのみ・する ／ テキスト＋長方形。※外部スタイル未使用時のみ）→ New Fill を2枚追加＋長方形シェイプ効果
5. 変換後のエリア内文字に用意したグラフィックスタイルを適用（「元の見た目」の一時スタイルのみ削除し、取り込んだ外部スタイルは残す）
6. 行揃え・テキストの配置は常に中央

#### 補足

- テキスト＋図形のフレームに使える図形は長方形のみ（軸並行・直線コーナーの長方形）。長方形以外の閉じたパスのみが図形として選ばれている場合は警告して中止する。
- テキスト＋図形は 1 組（最初のテキスト＋最初の長方形）のみ処理する。
- フレームサイズは、複製→アピアランス分割→アウトラインで計測した実寸を基準に決定する。
- 縦方向の中央配置とグラフィックスタイル登録にはダイナミックアクションを使用する。実行時に一時的に読み込み、終了時に自動で破棄するため、アクションパネルに残骸は残らない。

#### 対応している選択

① テキストのみ
- ポイント文字 … そのまま計測してエリア内文字化
- パス上文字 … 内部で一度ポイント文字に分離（workerDetachPathText）してから同じ経路で処理
- 複数選択時はそれぞれ個別に変換

② テキスト＋長方形
- テキスト（ポイント文字／パス上文字）＋ 長方形（軸並行・直線コーナー）
- 長方形を複製してエリア内文字のフレームにし、テキストを流し込む
- 1組（最初のテキスト＋最初の長方形）のみ処理

#### オプションの効き方

| | 大きさ調整の倍率 | ボタン背景（塗り2枚＋長方形効果） |
|---|---|---|
| テキストのみ・する | ○（幅×1.2 / 高さ×1.6 等） | ○ |
| テキストのみ・しない | 等倍 | ✕ |
| テキスト＋長方形 | ―（長方形サイズ優先） | ○ |

- ボタン背景は New Fill を2枚追加し、長方形シェイプ効果で背景化する。塗り色は設定しない（追加した塗りは既定のまま）。
- グラフィックスタイルで「読み込んだスタイル」を選んだ場合、そのスタイルが見た目を定義するためボタン背景は付与しない。現書類に未登録なら、記憶した AI ファイルから取り込んで適用する。参照ファイルとスタイル名は Folder.userData（styles_for_TextWithShapeToAreaType.txt）に記憶する。

### Overview

Convert point text, path text, or text + shape into area type while preserving appearance.

Flow:

1. On launch, show the resident palette (size adjustment On/Off with width% / height%, and a graphic-style choice: original appearance / a loaded style). Run via the "Convert" button; the palette stays open.
2. Before converting, prepare the graphic style to apply ("original appearance" = register the selected text's appearance temporarily; a loaded style = import from the AI file picked via "Load").
3. Point / path text → area type at the measured real size (scaled by the width/height ratios when size adjustment is On).
   Text + rectangle → duplicate the rectangle into an area-type frame and fill it with the text.
4. Button background (text-only "On" / text + rectangle; only when no external style is used) → add two fills via New Fill + a rectangle shape effect.
5. Apply the prepared graphic style to the converted area type (remove only the temp "original appearance" style; keep imported external styles).
6. Justification and vertical alignment are always centered.

#### Notes

- Only a rectangle (axis-aligned, straight corners) can serve as the text + shape frame. If the only shape selected is a non-rectangle closed path, the script alerts and aborts.
- Text + shape processes a single pair only (the first text + the first rectangle).
- Frame size is based on the real size measured via duplicate → expand appearance → create outlines.
- Vertical centering and graphic-style registration use dynamic actions that are loaded temporarily and removed automatically on exit, so nothing is left behind in the Actions panel.

#### Supported selections

① Text only
- Point text … measured directly and converted to area type
- Path text … first detached into point text internally (workerDetachPathText), then converted through the same path
- With a multiple selection, each is converted individually

② Text + rectangle
- Text (point / path) + a rectangle (axis-aligned, straight corners)
- The rectangle is duplicated into the area-type frame and filled with the text
- Only a single pair is processed (the first text + the first rectangle)

#### How the options behave

| | Size ratio | Button background (2 fills + rectangle effect) |
|---|---|---|
| Text only · On | Yes (e.g. W×1.2 / H×1.6) | Yes |
| Text only · Off | 1× | No |
| Text + rectangle | — (rectangle size wins) | Yes |

- The button background adds two fills via New Fill and reshapes them with the rectangle shape effect. Fill colors are not set (the added fills keep their defaults).
- When a loaded style is chosen, the style defines the appearance, so no button background is added. If not registered in the current document, it is imported from the remembered AI file and applied. The source file and its style names are remembered in Folder.userData (styles_for_TextWithShapeToAreaType.txt).

### 更新履歴 / Changelog

- v1.3.0 (2026-07-01): モーダルダイアログを常駐パレット化（#targetengine ＋ $.global 単一インスタンスガード）。DOM 処理は BridgeTalk でメインエンジンへ委譲し、パレットは開いたまま実行。「変換」ボタンはエリア内文字オプションパネル内に配置。グラフィックスタイルの listbox はクリックした時点で選択オブジェクトへ即適用（「適用」ボタンなし）。閉じるボタンは廃止し、パレットをアクティブにして Esc キーで閉じる。worker と重複していた旧・非worker実装を一掃／ Converted the modal dialog into a resident palette (#targetengine + a $.global single-instance guard); DOM work is delegated to the main engine via BridgeTalk and runs while the palette stays open; the "Convert" button now sits inside the Area Type Options panel; clicking a graphic style in the listbox live-applies it to the current selection (no Apply button); the Close button was dropped in favor of pressing Esc while the palette is active; removed the legacy non-worker code duplicated by the workers
- v1.2.0 (2026-07-01): 固定パス（TARGET_FILE_PATH）と固定スタイル名（文字白抜き／枠のみ）を撤去。ダイアログに「スタイルの読み込み」ボタンを追加し、選んだ AI ファイルとスタイル名を Folder.userData に記憶。取り込んだスタイル名からラジオを自動生成する構成へ変更（ImportGraphicStyles v1.7.0 より移植）／ Removed the hardcoded path (TARGET_FILE_PATH) and fixed style names (white text / frame only); added a "Load Styles" button that remembers the picked AI file and style names in Folder.userData; radios are now generated automatically from the imported style names (ported from ImportGraphicStyles v1.7.0)
- v1.1.0 (2026-07-01): ダイアログにグラフィックスタイル選択（元の見た目／文字白抜き／枠のみ）を追加。外部 AI ファイル（TARGET_FILE_PATH）から未登録スタイルを取り込んで適用（ImportGraphicStyles より移植）。角丸オプションを削除／ Added a graphic-style choice (original appearance / white text / frame only); imports the chosen style from an external AI file when it is not yet registered (ported from ImportGraphicStyles). Removed the round-corners option
- v1.0.0 : 初期バージョン / Initial release

*/

// =========================================
// 基本情報 / Basic info
// =========================================
var SCRIPT_NAME     = "ImportAndApplyGraphicStyle";   /* スクリプト名 / script name */
var SCRIPT_VERSION  = "v1.3.0";                       /* バージョン / version */
var SCRIPT_AUTHOR   = "Masahiro Takano (@swwwitch)";  /* 作者 / author */
var SCRIPT_RELEASED = "";                             /* 最初のリリース日 / first release date */
var SCRIPT_UPDATED  = "";                             /* 更新日 / last updated */

// Released under the MIT license
// http://opensource.org/licenses/mit-license.php

(function () {

    // =========================================
    // 設定ファイル / Preferences file
    // =========================================

    /* 参照した AI ファイルとスタイル名を記憶する設定ファイル / Prefs file remembering the picked AI file and its style names */
    var PREFS_FILE_NAME = "styles_for_TextWithShapeToAreaType.txt";

    // =========================================
    // ローカライズ / Localization
    // =========================================

    /* 表示言語を判定 / Detect display language */
    function getCurrentLang() {
        return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
    }
    var currentLanguage = getCurrentLang();

    /* 日英ラベル定義（カテゴリ構造）/ Japanese-English label definitions (categorized) */
    var LABELS = {
        /* 警告メッセージ / Alerts */
        alert: {
            selectText: {
                ja: "ポイント文字・パス上文字、またはテキストと図形を選択してください。",
                en: "Please select point text, path text, or text and a shape."
            },
            noDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            rectangleOnly: {
                ja: "フレームに使える図形は長方形のみです。長方形を選択してください。",
                en: "Only a rectangle can serve as the frame. Please select a rectangle."
            },
            fileNotFound: {
                ja: "指定されたファイルが見つかりません：\n",
                en: "The specified file was not found:\n"
            },
            styleNotFound: {
                ja: "指定したグラフィックスタイルが見つかりません：\n",
                en: "The graphic style was not found:\n"
            }
        },
        /* パレットUI / Palette UI */
        ui: {
            dialogTitle: { ja: "エリア内文字に変換", en: "Convert to Area Type" },
            areaTypeOptionPanel: { ja: "エリア内文字オプション", en: "Area Type Options" },
            convertToAreaType: { ja: "エリア内文字に変換", en: "Convert to area type" },
            sizeLabel: { ja: "サイズ：", en: "Size:" },
            doAdjust: { ja: "調整する", en: "Adjust" },
            dontAdjust: { ja: "調整しない", en: "Don't adjust" },
            widthRatio: { ja: "幅：", en: "Width:" },
            heightRatio: { ja: "高さ：", en: "Height:" },
            stylePanel: { ja: "グラフィックスタイル", en: "Graphic style" },
            styleOriginal: { ja: "元の見た目", en: "Original appearance" },
            usedOnly: { ja: "ドキュメント内で使用しているもののみ", en: "Only those used in the document" },
            loadPanel: { ja: "スタイルの読み込み", en: "Load Styles" },
            loadButton: { ja: "読み込み", en: "Load" },
            reloadButton: { ja: "再読み込み", en: "Reload" },
            openStylePanel: { ja: "パネル", en: "Panel" },
            openStylePanelHint: {
                ja: "グラフィックスタイルパネルに移動",
                en: "Go to the Graphic Styles panel."
            },
            clearStyle: { ja: "クリア", en: "Clear" },
            clearStyleHint: {
                ja: "未使用のグラフィックスタイルを削除",
                en: "Delete unused graphic styles."
            },
            pickFile: { ja: "スタイルの AI ファイルを選択", en: "Select a style AI file" },
            noFileSelected: { ja: "ファイル未選択", en: "No file selected" },
            noStylesHint: {
                ja: "「読み込み」でスタイルの AI ファイルを選択してください。",
                en: "Click “Load” to choose a style AI file."
            },
            reloadHint: {
                ja: "記憶したファイルからスタイルを取り込み直します。",
                en: "Re-import styles from the remembered file."
            },
            runConvert: { ja: "変換", en: "Convert" }
        }
    };

    /* ドット区切りキーから LABELS のエントリを取得 / Resolve a LABELS entry from a dot-separated key */
    function getLabelEntry(key) {
        var parts = key.split(".");
        var node = LABELS;
        for (var i = 0; i < parts.length; i++) {
            if (node && typeof node[parts[i]] !== "undefined") { node = node[parts[i]]; }
            else { return null; }
        }
        return node;
    }

    /* キーからローカライズ文字列を取得 / Get a localized string by key */
    function L(key) {
        var entry = getLabelEntry(key);
        if (entry) {
            if (entry[currentLanguage]) return entry[currentLanguage];
            if (entry.en) return entry.en;
        }
        return key;
    }

    // =========================================
    // BridgeTalk 委譲基盤 / BridgeTalk delegation
    // 常駐パレットの app は表示中に DOM 接続を失うため、DOM を触る処理は worker 関数に集約し、
    // 押下のたびにメインエンジン（illustrator）へ同期委譲する。
    // A resident palette's app loses its DOM connection while shown, so DOM work lives in worker
    // functions and is delegated synchronously to the main engine (illustrator) on each action.
    // =========================================

    /* 戻り値マーカー / Result markers */
    var MARKER_OK = "OK";
    var MARKER_NODOC = "NODOC";
    var MARKER_NOSEL = "NOSEL";
    var MARKER_ERR = "ERR";

    /* 委譲する worker 関数を全登録（追加漏れ防止）。
       worker 関数内は「// 行コメント禁止・/* *\/ のみ・必ずセミコロンで終える」（toString が改行を消すため）。
       Register every worker function here (avoid omissions). Inside workers: no // comments, /* *\/ only,
       always end statements with a semicolon (toString strips newlines). */
    var WORKER_FUNCS = [];

    /* 再入防止ガード / Re-entrancy guard */
    var isBusy = false;

    /* 登録された worker 関数を連結して worker ソースを生成 / Concatenate registered workers into the worker source */
    function buildWorkerSource() {
        var source = "";
        for (var i = 0; i < WORKER_FUNCS.length; i++) {
            source += WORKER_FUNCS[i].toString() + "\n";
        }
        return source;
    }

    /* worker 呼び出し式をメインエンジンへ同期委譲し、戻り値の文字列を返す
       callExpr は文字列を返す式（例: 'workerOpenStylePanel()'）
       Delegate a worker call expression to the main engine synchronously; returns its string result.
       callExpr is an expression that returns a string (e.g. 'workerOpenStylePanel()'). */
    function delegate(callExpr) {
        if (isBusy) return MARKER_ERR + ":busy";
        isBusy = true;
        var holder = { result: null };
        try {
            var payload = buildWorkerSource() + "\n" + callExpr + ";";
            var bridge = new BridgeTalk();
            bridge.target = "illustrator";
            /* バックスラッシュ・多バイト・改行の破損を避けるため encodeURIComponent で包む
               Wrap via encodeURIComponent to avoid backslash / multibyte / newline corruption */
            bridge.body = "eval(decodeURIComponent(\"" + encodeURIComponent(payload) + "\"));";
            bridge.onResult = function (message) { holder.result = message.body; };
            bridge.onError = function (message) { holder.result = MARKER_ERR + ":" + message.body; };
            bridge.send(10); /* 同期送信 / Synchronous send */
        } catch (e) {
            holder.result = MARKER_ERR + ":" + e;
        } finally {
            isBusy = false;
        }
        return (holder.result === null) ? (MARKER_ERR + ":noresult") : holder.result;
    }

    /* 文字列を委譲用の JS 文字列リテラルへエスケープ / Escape a string as a JS string literal for the call expression */
    function jsStringLiteral(text) {
        text = String(text);
        var out = '"';
        for (var i = 0; i < text.length; i++) {
            var ch = text.charAt(i);
            if (ch === '\\') out += '\\\\';
            else if (ch === '"') out += '\\"';
            else if (ch === '\n') out += '\\n';
            else if (ch === '\r') out += '\\r';
            else if (ch === '\t') out += '\\t';
            else out += ch;
        }
        return out + '"';
    }

    /* worker の "OK\t<改行連結>" 形式を解析 / Parse the worker's "OK\t<newline-joined>" result */
    function parseMarkerList(result) {
        if (!result) return { ok: false, marker: MARKER_ERR, items: [] };
        var tabIndex = result.indexOf("\t");
        var marker = (tabIndex >= 0) ? result.substring(0, tabIndex) : result;
        if (marker !== MARKER_OK) return { ok: false, marker: marker, items: [] };
        var rest = (tabIndex >= 0) ? result.substring(tabIndex + 1) : "";
        var items = rest.length ? rest.split("\n") : [];
        return { ok: true, marker: MARKER_OK, items: items };
    }

    // ---- worker 関数 / Worker functions（メインエンジンで実行 / run in the main engine）----

    /* worker: グラフィックスタイルパネルを表示 / Show the Graphic Styles panel */
    function workerOpenStylePanel() {
        try {
            app.executeMenuCommand("Adobe Style Palette");
            return "OK";
        } catch (e) {
            return "ERR:" + e;
        }
    }
    WORKER_FUNCS.push(workerOpenStylePanel);

    /* worker: 文字列を16進へ / Convert a string to hex */
    function workerHexAscii(text) {
        var hex = "";
        var i;
        for (i = 0; i < text.length; i++) {
            var pair = text.charCodeAt(i).toString(16);
            if (pair.length < 2) { pair = "0" + pair; }
            hex += pair;
        }
        return hex;
    }
    WORKER_FUNCS.push(workerHexAscii);

    /* worker: メニューコマンド1件のイベントブロックを組み立て / Build one menu-command event block */
    function workerMenuEventBlock(eventIndex, internalName, localizedNameHex, commandNameHex, value, hasDialog) {
        var lines = [];
        lines.push("\t/event-" + eventIndex + " {");
        lines.push("\t\t/useRulersIn1stQuadrant 1");
        lines.push("\t\t/internalName (" + internalName + ")");
        lines.push("\t\t/localizedName [ " + (localizedNameHex.length / 2));
        lines.push("\t\t\t" + localizedNameHex);
        lines.push("\t\t]");
        lines.push("\t\t/isOpen 0");
        lines.push("\t\t/isOn 1");
        lines.push("\t\t/hasDialog " + (hasDialog ? "1" : "0"));
        if (hasDialog) { lines.push("\t\t/showDialog 0"); }
        lines.push("\t\t/parameterCount 1");
        lines.push("\t\t/parameter-1 {");
        lines.push("\t\t\t/key 1835363957");
        lines.push("\t\t\t/showInPalette 1");
        lines.push("\t\t\t/type (enumerated)");
        lines.push("\t\t\t/name [ " + (commandNameHex.length / 2));
        lines.push("\t\t\t\t" + commandNameHex);
        lines.push("\t\t\t]");
        lines.push("\t\t\t/value " + value);
        lines.push("\t\t}");
        lines.push("\t}");
        return lines.join("\n");
    }
    WORKER_FUNCS.push(workerMenuEventBlock);

    /* worker: グラフィックスタイル用「未使用をすべて選択 → 削除」の一時アクション定義を組み立て
       Build the "Select All Unused -> Delete" action for graphic styles */
    function workerBuildPruneAction(setName, actionName) {
        var internalName = "ai_plugin_styles";
        var localizedNameHex = "e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab";
        var selectValue = 14;
        var deleteValue = 3;
        var deleteNameHex = "44656c657465205374796c65";
        var selectAllUnusedHex = "53656c65637420416c6c20556e75736564";
        var parts = [];
        parts.push("/version 3");
        parts.push("/name [ " + setName.length);
        parts.push("\t" + workerHexAscii(setName));
        parts.push("]");
        parts.push("/isOpen 1");
        parts.push("/actionCount 1");
        parts.push("/action-1 {");
        parts.push("\t/name [ " + actionName.length);
        parts.push("\t\t" + workerHexAscii(actionName));
        parts.push("\t]");
        parts.push("\t/keyIndex 0");
        parts.push("\t/colorIndex 0");
        parts.push("\t/isOpen 1");
        parts.push("\t/eventCount 2");
        parts.push(workerMenuEventBlock(1, internalName, localizedNameHex, selectAllUnusedHex, selectValue, false));
        parts.push(workerMenuEventBlock(2, internalName, localizedNameHex, deleteNameHex, deleteValue, true));
        parts.push("}");
        return parts.join("\n");
    }
    WORKER_FUNCS.push(workerBuildPruneAction);

    /* worker: 失敗を無視する後始末 / Best-effort cleanup, ignoring failures */
    function workerIgnoringErrors(fn) {
        try { fn(); } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerIgnoringErrors);

    /* worker: アクションを一時ファイルに書き出して再生 / Write, load, and play the action */
    function workerPlayTempAction(actionSource, setName, actionName, fileName) {
        var actionFile = new File(fileName);
        try {
            actionFile.encoding = "UTF-8";
            if (!actionFile.open("w")) { return false; }
            actionFile.write(actionSource);
            actionFile.close();
            workerIgnoringErrors(function () { app.unloadAction(setName, ""); });
            app.loadAction(actionFile);
            app.doScript(actionName, setName);
        } catch (e) {
            /* ignore */
        }
        workerIgnoringErrors(function () { actionFile.close(); });
        workerIgnoringErrors(function () { app.unloadAction(setName, ""); });
        workerIgnoringErrors(function () { actionFile.remove(); });
        return true;
    }
    WORKER_FUNCS.push(workerPlayTempAction);

    /* worker: 未使用グラフィックスタイルをアクションで削除し、削除件数を返す / Prune unused graphic styles, return the count */
    function workerPruneUnusedCore(doc) {
        var countBefore = doc.graphicStyles.length;
        var source = workerBuildPruneAction("TemporaryActionSet", "TemporaryActionName");
        workerPlayTempAction(source, "TemporaryActionSet", "TemporaryActionName", Folder.temp + "/ImportAndApplyGraphicStyle_prune.aia");
        try { app.redraw(); } catch (e) { /* ignore */ }
        var removed = countBefore - doc.graphicStyles.length;
        if (removed < 0) { removed = 0; }
        return removed;
    }
    WORKER_FUNCS.push(workerPruneUnusedCore);

    /* worker: 未使用スタイル名の集合を取得（未使用削除→残りから判定→削除があれば undo で復元）
       Get the set of unused style names (delete unused, classify by what remains, undo only if deleted) */
    function workerComputeUnusedSet(doc) {
        var unused = {};
        var before = [];
        var i;
        for (i = 1; i < doc.graphicStyles.length; i++) { before.push(doc.graphicStyles[i].name); }
        if (before.length === 0) { return unused; }
        var lengthBefore = doc.graphicStyles.length;
        workerPruneUnusedCore(doc);
        var remaining = {};
        var j;
        for (j = 1; j < doc.graphicStyles.length; j++) { remaining[doc.graphicStyles[j].name] = true; }
        var k;
        for (k = 0; k < before.length; k++) {
            if (!remaining[before[k]]) { unused[before[k]] = true; }
        }
        if (doc.graphicStyles.length < lengthBefore) {
            try { app.undo(); app.redraw(); } catch (e) { /* ignore */ }
        }
        return unused;
    }
    WORKER_FUNCS.push(workerComputeUnusedSet);

    /* worker: 現書類のグラフィックスタイル名を返す（既定と temp_style を除外、usedOnly で未使用も除外）
       戻り値: "OK\t" + 名前を \n 連結 / "NODOC"
       Return the document's graphic-style names (skip default and temp_style; usedOnly also skips unused).
       Returns "OK\t" + names joined by \n, or "NODOC" */
    function workerGetStyleNames(usedOnly) {
        if (app.documents.length === 0) { return "NODOC"; }
        var doc = app.activeDocument;
        var unused = {};
        if (usedOnly) { unused = workerComputeUnusedSet(doc); }
        var names = [];
        var i;
        for (i = 1; i < doc.graphicStyles.length; i++) {
            var nm = doc.graphicStyles[i].name;
            if (nm === "temp_style") { continue; }
            if (usedOnly && unused[nm]) { continue; }
            names.push(nm);
        }
        return "OK\t" + names.join("\n");
    }
    WORKER_FUNCS.push(workerGetStyleNames);

    /* worker: 未使用グラフィックスタイルを削除。戻り値: "OK\t" + 件数 / "NODOC"
       Prune unused graphic styles. Returns "OK\t" + count, or "NODOC" */
    function workerPruneUnused() {
        if (app.documents.length === 0) { return "NODOC"; }
        var removed = workerPruneUnusedCore(app.activeDocument);
        return "OK\t" + removed;
    }
    WORKER_FUNCS.push(workerPruneUnused);

    /* worker: 指定 AI ファイルからグラフィックスタイルを現書類へ取り込む
       戻り値: "OK\t" + 実際に登録された名前を \n 連結 / "NOFILE" / "NODOC"
       Import graphic styles from the AI file into the current document.
       Returns "OK\t" + registered names joined by \n, or "NOFILE" / "NODOC" */
    function workerImportStyles(filePath) {
        var styleFile = new File(filePath);
        if (!styleFile.exists) { return "NOFILE"; }
        if (app.documents.length === 0) { return "NODOC"; }
        var destinationDoc = app.activeDocument;
        var styleSourceDoc = app.open(styleFile);
        var sourceStyleNames = [];
        var i;
        for (i = 1; i < styleSourceDoc.graphicStyles.length; i++) { sourceStyleNames.push(styleSourceDoc.graphicStyles[i].name); }
        app.executeMenuCommand("selectallinartboard");
        app.executeMenuCommand("copy");
        styleSourceDoc.close(SaveOptions.DONOTSAVECHANGES);
        app.activeDocument = destinationDoc;
        var importLayerName = "// _imported";
        var importLayer;
        try { importLayer = destinationDoc.layers.getByName(importLayerName); }
        catch (e) { importLayer = destinationDoc.layers.add(); importLayer.name = importLayerName; }
        importLayer.locked = false;
        importLayer.visible = true;
        destinationDoc.activeLayer = importLayer;
        app.executeMenuCommand("paste");
        try { importLayer.remove(); } catch (e2) { /* ignore */ }
        try { app.redraw(); } catch (e3) { /* ignore */ }
        var imported = [];
        var k;
        for (k = 0; k < sourceStyleNames.length; k++) {
            var exists = false;
            try { destinationDoc.graphicStyles.getByName(sourceStyleNames[k]); exists = true; } catch (e4) { exists = false; }
            if (exists) { imported.push(sourceStyleNames[k]); }
        }
        return "OK\t" + imported.join("\n");
    }
    WORKER_FUNCS.push(workerImportStyles);

    /* worker: 直前の操作を取り消す（★ドキュメント確認を app.undo() より先に）
       Undo the last operation (doc check MUST come before app.undo()) */
    function workerUndoLast() {
        if (app.documents.length === 0) { return "NODOC"; }
        try { app.undo(); app.redraw(); } catch (e) { return "ERR:" + e; }
        return "OK";
    }
    WORKER_FUNCS.push(workerUndoLast);

    /* worker: 選択オブジェクトに指定グラフィックスタイルを適用
       undoFirst=true なら直前のプレビューを app.undo() で取り消してから再適用（★ドキュメント確認を先に）
       戻り値: "OK" / "NODOC" / "NOSEL" / "NOSTYLE"
       Apply the named graphic style to the current selection. When undoFirst is true, undo the previous
       preview first, then re-apply (doc check comes before app.undo()). Returns OK/NODOC/NOSEL/NOSTYLE */
    function workerApplyStyleToSelection(styleName, undoFirst) {
        if (app.documents.length === 0) { return "NODOC"; }
        if (undoFirst) {
            try { app.undo(); app.redraw(); } catch (e) { /* ignore */ }
        }
        var doc = app.activeDocument;
        var sel = doc.selection;
        if (!sel || sel.length === 0) { return "NOSEL"; }
        var style = null;
        try { style = doc.graphicStyles.getByName(styleName); } catch (e2) { style = null; }
        if (!style) { return "NOSTYLE"; }
        var i;
        for (i = 0; i < sel.length; i++) {
            try { style.applyTo(sel[i]); } catch (e3) { /* ignore */ }
        }
        try { app.redraw(); } catch (e4) { /* ignore */ }
        return "OK";
    }
    WORKER_FUNCS.push(workerApplyStyleToSelection);

    // ---- エリア内文字変換 worker 群 / Area-type conversion workers ----

    /* worker: /name ブロック（ASCII, 大文字hex）/ /name block (ASCII, uppercase hex) */
    function workerNameBlockAscii(text) {
        return "/name [ " + text.length + " " + workerHexAscii(text).toUpperCase() + " ]";
    }
    WORKER_FUNCS.push(workerNameBlockAscii);

    /* worker: アクションセット定義(.aia)文字列を組み立て / Build an action set (.aia) string */
    function workerBuildActionSetAIA(setName, internalName, localizedNameHex, paramKeyInt, actionDefs) {
        var aia = "/version 3" + workerNameBlockAscii(setName) + "/isOpen 1" + "/actionCount " + actionDefs.length;
        var i;
        for (i = 0; i < actionDefs.length; i++) {
            var d = actionDefs[i];
            aia += "/action-" + (i + 1) + " {" +
                " " + workerNameBlockAscii(d.name) +
                " /keyIndex 0" +
                " /colorIndex 0" +
                " /isOpen 1" +
                " /eventCount 1" +
                " /event-1 {" +
                " /useRulersIn1stQuadrant 0" +
                " /internalName (" + internalName + ")" +
                (localizedNameHex ? (" /localizedName [ " + localizedNameHex + " ]") : "") +
                " /isOpen 0" +
                " /isOn 1" +
                " /hasDialog 0" +
                " /parameterCount 1" +
                " /parameter-1 {" +
                " /key " + paramKeyInt +
                " /showInPalette 4294967295" +
                " /type (integer)" +
                " /value " + d.value +
                " }" +
                " }" +
                "}";
        }
        return aia;
    }
    WORKER_FUNCS.push(workerBuildActionSetAIA);

    /* worker: アクションセットを読み込む（temp に .aia を書き出して loadAction）/ Load an action set */
    function workerLoadActionSet(setName, aiaString) {
        try {
            try { app.unloadAction(setName, ""); } catch (e0) { /* ignore */ }
            var actionFile = new File(Folder.temp + "/IAAGS_action_" + setName + ".aia");
            actionFile.open("w");
            actionFile.write(aiaString);
            actionFile.close();
            app.loadAction(actionFile);
            try { actionFile.remove(); } catch (e1) { /* ignore */ }
        } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerLoadActionSet);

    /* worker: アクションセットを破棄 / Unload an action set */
    function workerUnloadActionSet(setName) {
        try { app.unloadAction(setName, ""); } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerUnloadActionSet);

    /* worker: フレーム整列アクション（AlignTop/Center/Bottom/Justify）を読み込む / Load frame-alignment actions */
    function workerLoadAreaTextActions() {
        var aia = workerBuildActionSetAIA(
            "AreaText",
            "adobe_frameAlignment",
            "39 e382a8e383aae382a2e58685e69687e5ad97e381aee38395e383ace383bce383a0e695b4e58897",
            1717660782,
            [
                { name: "AlignTop", value: 0 },
                { name: "AlignCenter", value: 1 },
                { name: "AlignBottom", value: 2 },
                { name: "AlignJustify", value: 3 }
            ]
        );
        workerLoadActionSet("AreaText", aia);
    }
    WORKER_FUNCS.push(workerLoadAreaTextActions);

    /* worker: フレーム整列アクションを破棄 / Unload frame-alignment actions */
    function workerUnloadAreaTextActions() {
        workerUnloadActionSet("AreaText");
    }
    WORKER_FUNCS.push(workerUnloadAreaTextActions);

    /* worker: 縦方向の配置アクションを実行（0=上,1=中央,2=下,3=均等）/ Run vertical-placement action */
    function workerRunFrameAlignment(valueInt) {
        if (valueInt !== 0 && valueInt !== 1 && valueInt !== 2 && valueInt !== 3) { return; }
        var actionName = "AlignTop";
        if (valueInt === 1) { actionName = "AlignCenter"; }
        else if (valueInt === 2) { actionName = "AlignBottom"; }
        else if (valueInt === 3) { actionName = "AlignJustify"; }
        try { app.doScript(actionName, "AreaText", false); } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerRunFrameAlignment);

    /* worker: 指定フレームに縦方向の配置を適用 / Apply vertical placement to a frame */
    function workerApplyFrameAlignment(areaType, valueInt) {
        try {
            var doc = app.activeDocument;
            doc.selection = null;
            doc.selection = [areaType];
            app.redraw();
            workerRunFrameAlignment(valueInt);
        } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerApplyFrameAlignment);

    /* worker: TextFrame の見た目を temp_style として登録し、登録名を返す / Register the appearance as temp_style */
    function workerRegisterTempStyle(textFrame) {
        if (!textFrame) { return null; }
        var doc = app.activeDocument;
        if (!doc) { return null; }
        var graphicStyles = doc.graphicStyles;
        if (!graphicStyles) { return null; }
        try { graphicStyles.getByName("temp_style").remove(); } catch (e) { /* ignore */ }
        doc.selection = null;
        try { textFrame.selected = true; } catch (selErr) { return null; }
        var countBefore = graphicStyles.length;
        var aia = '/version 3 /name [ 12 477261706869635374796c65 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /value 1 } } }';
        workerLoadActionSet("GraphicStyle", aia);
        try { app.doScript("AddNewWithoutName", "GraphicStyle", false); } catch (runErr) { /* ignore */ }
        workerUnloadActionSet("GraphicStyle");
        if (graphicStyles.length <= countBefore) { return null; }
        graphicStyles[graphicStyles.length - 1].name = "temp_style";
        return "temp_style";
    }
    WORKER_FUNCS.push(workerRegisterTempStyle);

    /* worker: 名前でグラフィックスタイルを適用 / Apply a graphic style by name */
    function workerApplyStyleByName(name, item) {
        if (!name || !item) { return false; }
        try { app.activeDocument.graphicStyles.getByName(name).applyTo(item); return true; } catch (e) { return false; }
    }
    WORKER_FUNCS.push(workerApplyStyleByName);

    /* worker: 名前でグラフィックスタイルを削除 / Remove a graphic style by name */
    function workerRemoveStyleByName(name) {
        if (!name) { return; }
        try { app.activeDocument.graphicStyles.getByName(name).remove(); } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerRemoveStyleByName);

    /* worker: 選択の可視バウンディングボックスの和 / Union of visibleBounds over a selection */
    function workerSelectionVisibleBounds(sel) {
        if (!sel || !sel.length) { return null; }
        var left = null, top = null, right = null, bottom = null;
        var i;
        for (i = 0; i < sel.length; i++) {
            var b;
            try { b = sel[i].visibleBounds; } catch (e) { continue; }
            if (!b) { continue; }
            if (left === null || b[0] < left) { left = b[0]; }
            if (top === null || b[1] > top) { top = b[1]; }
            if (right === null || b[2] > right) { right = b[2]; }
            if (bottom === null || b[3] < bottom) { bottom = b[3]; }
        }
        if (left === null) { return null; }
        return [left, top, right, bottom];
    }
    WORKER_FUNCS.push(workerSelectionVisibleBounds);

    /* worker: 複製→アピアランス分割→アウトラインで正確な可視サイズを計測 / Measure accurate visible size */
    function workerMeasureAccurateBounds(item) {
        var doc = app.activeDocument;
        var savedSel = doc.selection;
        var measured = null;
        var duplicatedItem = null;
        try {
            duplicatedItem = item.duplicate();
            doc.selection = null;
            duplicatedItem.selected = true;
            app.redraw();
            try { app.executeMenuCommand("expandStyle"); } catch (e) { /* ignore */ }
            try { app.executeMenuCommand("outline"); } catch (e5) { /* ignore */ }
            app.redraw();
            var resultSel = doc.selection;
            var bounds = workerSelectionVisibleBounds(resultSel);
            if (bounds) {
                measured = { left: bounds[0], top: bounds[1], right: bounds[2], bottom: bounds[3], width: bounds[2] - bounds[0], height: bounds[1] - bounds[3] };
            }
            var d;
            for (d = resultSel.length - 1; d >= 0; d--) {
                try { resultSel[d].remove(); } catch (e2) { /* ignore */ }
            }
        } catch (e0) {
            if (duplicatedItem) { try { duplicatedItem.remove(); } catch (e3) { /* ignore */ } }
        }
        try { doc.selection = savedSel; } catch (e4) { /* ignore */ }
        return measured;
    }
    WORKER_FUNCS.push(workerMeasureAccurateBounds);

    /* worker: geometricBounds を {left, top, width, height} で返す / geometricBounds as an object */
    function workerGeometricBounds(item) {
        var b = item.geometricBounds;
        return { left: b[0], top: b[1], width: b[2] - b[0], height: b[1] - b[3] };
    }
    WORKER_FUNCS.push(workerGeometricBounds);

    /* worker: 軸並行・直線コーナーの長方形か判定 / Is this an axis-aligned straight-corner rectangle */
    function workerIsRectanglePath(item) {
        if (!item || item.typename !== "PathItem" || !item.closed) { return false; }
        var points;
        try { points = item.pathPoints; } catch (e) { return false; }
        if (!points || points.length !== 4) { return false; }
        var i;
        for (i = 0; i < points.length; i++) {
            var anchor = points[i].anchor, left = points[i].leftDirection, right = points[i].rightDirection;
            if (Math.abs(anchor[0] - left[0]) >= 0.01 || Math.abs(anchor[1] - left[1]) >= 0.01) { return false; }
            if (Math.abs(anchor[0] - right[0]) >= 0.01 || Math.abs(anchor[1] - right[1]) >= 0.01) { return false; }
        }
        var distinctX = [], distinctY = [];
        var j, k;
        for (j = 0; j < points.length; j++) {
            var ax = points[j].anchor[0], ay = points[j].anchor[1];
            var foundX = false, foundY = false;
            for (k = 0; k < distinctX.length; k++) { if (Math.abs(distinctX[k] - ax) < 0.01) { foundX = true; } }
            if (!foundX) { distinctX.push(ax); }
            for (k = 0; k < distinctY.length; k++) { if (Math.abs(distinctY[k] - ay) < 0.01) { foundY = true; } }
            if (!foundY) { distinctY.push(ay); }
        }
        return distinctX.length === 2 && distinctY.length === 2;
    }
    WORKER_FUNCS.push(workerIsRectanglePath);

    /* worker: 元テキストの自動カーニングと文字組みを取得 / Capture auto-kerning and mojikumi */
    function workerReadKerningMojikumi(sourceText) {
        var snap = { kerningMethod: null, mojikumi: null };
        try { snap.kerningMethod = sourceText.textRange.characterAttributes.kerningMethod; } catch (e) { /* ignore */ }
        try {
            if (sourceText.paragraphs.length > 0) { snap.mojikumi = sourceText.paragraphs[0].paragraphAttributes.mojikumi; }
        } catch (e2) { /* ignore */ }
        return snap;
    }
    WORKER_FUNCS.push(workerReadKerningMojikumi);

    /* worker: 自動カーニングと文字組みをエリア内文字へ適用 / Apply auto-kerning and mojikumi */
    function workerApplyKerningMojikumi(areaType, snap) {
        if (snap.kerningMethod !== null) {
            try { areaType.textRange.characterAttributes.kerningMethod = snap.kerningMethod; } catch (e) { /* ignore */ }
        }
        if (snap.mojikumi !== null && snap.mojikumi !== undefined) {
            var paragraphs = areaType.paragraphs;
            var i;
            for (i = 0; i < paragraphs.length; i++) {
                try { paragraphs[i].paragraphAttributes.mojikumi = snap.mojikumi; } catch (e2) { /* ignore */ }
            }
        }
    }
    WORKER_FUNCS.push(workerApplyKerningMojikumi);

    /* worker: 図形パスをエリア内文字にして内容・書式・スタイルを移す / Turn a path into area type carrying text */
    function workerFillAreaType(doc, framePath, sourceText, graphicStyleName) {
        var sourceFont = null, sourceSize = 0;
        try {
            var sa = sourceText.textRange.characterAttributes;
            sourceFont = sa.textFont;
            sourceSize = sa.size;
        } catch (e) { /* ignore */ }
        var snap = workerReadKerningMojikumi(sourceText);
        var areaType = doc.textFrames.areaText(framePath);
        areaType.contents = sourceText.contents;
        try {
            if (sourceFont) { areaType.textRange.characterAttributes.textFont = sourceFont; }
            if (sourceSize > 0) { areaType.textRange.characterAttributes.size = sourceSize; }
        } catch (e2) { /* ignore */ }
        workerApplyKerningMojikumi(areaType, snap);
        if (graphicStyleName) { workerApplyStyleByName(graphicStyleName, areaType); }
        return areaType;
    }
    WORKER_FUNCS.push(workerFillAreaType);

    /* worker: エリア内文字を水平・垂直とも中央に / Center area-type contents */
    function workerCenterAreaType(areaType) {
        try { areaType.textRange.paragraphAttributes.justification = Justification.CENTER; } catch (e) { /* ignore */ }
        workerApplyFrameAlignment(areaType, 1);
    }
    WORKER_FUNCS.push(workerCenterAreaType);

    /* worker: 塗り2枚＋長方形シェイプ効果でボタン状の背景に / Add two fills + rectangle shape effect */
    function workerApplyButtonShape(areaType) {
        try {
            app.activeDocument.selection = null;
            areaType.selected = true;
            app.redraw();
            app.executeMenuCommand("Adobe New Fill Shortcut");
            app.executeMenuCommand("Adobe New Fill Shortcut");
            areaType.applyEffect('<LiveEffect name="Adobe Shape Effects" isPre="1"><Dict data="U DisplayString Rectangle I Shape 0 R RelWidth 0 R RelHeight 0 R AbsWidth 0 R AbsHeight 0 R Absolute 0 R CornerRadius 9 "/></LiveEffect>');
        } catch (e) { /* ignore */ }
    }
    WORKER_FUNCS.push(workerApplyButtonShape);

    /* worker: パス上文字を字形を保ったままポイント文字へ分離 / Detach path text into point text */
    function workerDetachPathText(doc, pathTextFrames) {
        function safe(fn) { try { return fn(); } catch (e) { return undefined; } }
        var created = [];
        if (!doc || !pathTextFrames || !pathTextFrames.length) { return created; }
        safe(function () { doc.selection = null; });
        var j;
        for (j = pathTextFrames.length - 1; j >= 0; j--) {
            var pathText = pathTextFrames[j];
            if (!pathText || pathText.typename !== "TextFrame" || pathText.kind !== TextType.PATHTEXT) { continue; }
            var originalPath = null;
            safe(function () { originalPath = pathText.textPath; });
            if (!originalPath) { continue; }
            var charAttrs = [];
            var c;
            for (c = 0; c < pathText.characters.length; c++) {
                var ca = pathText.characters[c].characterAttributes;
                charAttrs.push({ font: ca.textFont, size: ca.size, fillColor: ca.fillColor, strokeColor: ca.strokeColor, strokeWeight: ca.strokeWeight, autoLeading: ca.autoLeading, leading: ca.leading });
            }
            var textContents = "";
            safe(function () { textContents = pathText.contents; });
            var justification = null;
            safe(function () {
                if (pathText.paragraphs && pathText.paragraphs.length > 0) { justification = pathText.paragraphs[0].paragraphAttributes.justification; }
            });
            var newText = doc.textFrames.add();
            var anchorPoint = null;
            safe(function () {
                if (originalPath.pathPoints && originalPath.pathPoints.length > 0) { anchorPoint = originalPath.pathPoints[0].anchor; }
            });
            if (anchorPoint) { newText.position = [anchorPoint[0], anchorPoint[1]]; }
            newText.contents = textContents;
            if (justification !== null && newText.paragraphs && newText.paragraphs.length > 0) {
                safe(function () { newText.paragraphs[0].paragraphAttributes.justification = justification; });
            }
            safe(function () {
                var noColor = new NoColor();
                newText.textRange.characterAttributes.strokeColor = noColor;
                newText.textRange.characterAttributes.strokeWeight = 0;
            });
            var restoreCount = Math.min(newText.characters.length, charAttrs.length);
            var k;
            for (k = 0; k < restoreCount; k++) {
                var targetAttr = newText.characters[k].characterAttributes;
                var sourceAttr = charAttrs[k];
                safe(function () { targetAttr.textFont = sourceAttr.font; });
                safe(function () { targetAttr.size = sourceAttr.size; });
                safe(function () { targetAttr.fillColor = sourceAttr.fillColor; });
                safe(function () {
                    var sourceStroke = sourceAttr.strokeColor;
                    targetAttr.strokeColor = sourceStroke;
                    targetAttr.strokeWeight = (sourceStroke && sourceStroke.typename === "NoColor") ? 0 : sourceAttr.strokeWeight;
                });
                safe(function () { targetAttr.baselineShift = 0; });
                safe(function () { targetAttr.horizontalScale = 100; });
                safe(function () { targetAttr.verticalScale = 100; });
                safe(function () { targetAttr.autoLeading = sourceAttr.autoLeading; });
                if (!sourceAttr.autoLeading) { safe(function () { targetAttr.leading = sourceAttr.leading; }); }
            }
            safe(function () { pathText.remove(); });
            safe(function () { newText.selected = true; });
            created.push(newText);
        }
        return created;
    }
    WORKER_FUNCS.push(workerDetachPathText);

    /* worker: 選択内のパス上文字をポイント文字へ置き換えた選択配列を返す / Replace path text with point text */
    function workerPreprocessPathText(doc, sel) {
        if (!doc || !sel || !sel.length) { return sel; }
        var pathTexts = [];
        var i;
        for (i = 0; i < sel.length; i++) {
            var item = sel[i];
            try {
                if (item && item.typename === "TextFrame" && item.kind === TextType.PATHTEXT) { pathTexts.push(item); }
            } catch (e0) { /* ignore */ }
        }
        if (!pathTexts.length) { return sel; }
        var newTexts = workerDetachPathText(doc, pathTexts);
        if (!newTexts.length) { return sel; }
        var replaced = [];
        var j;
        for (j = 0; j < sel.length; j++) {
            var keepItem = sel[j];
            try {
                if (keepItem && keepItem.typename === "TextFrame" && keepItem.kind === TextType.PATHTEXT) { /* skip old */ }
                else if (keepItem) { replaced.push(keepItem); }
            } catch (e1) { /* ignore */ }
        }
        var k;
        for (k = 0; k < newTexts.length; k++) { replaced.push(newTexts[k]); }
        try { doc.selection = replaced; } catch (e) { /* ignore */ }
        try { app.redraw(); } catch (e2) { /* ignore */ }
        return replaced;
    }
    WORKER_FUNCS.push(workerPreprocessPathText);

    /* worker: 適用スタイルを決定（外部スタイル or 元テキストの一時スタイル）/ Resolve the style to apply */
    function workerResolveStyle(externalStyleName, sourceText) {
        if (externalStyleName) { return { name: externalStyleName, isTemp: false }; }
        return { name: workerRegisterTempStyle(sourceText), isTemp: true };
    }
    WORKER_FUNCS.push(workerResolveStyle);

    /* worker: テキスト＋図形 → 図形を複製してエリア内文字に / Text + shape -> duplicate the shape into area type */
    function workerConvertTextIntoShape(doc, selection, externalStyleName) {
        var created = [];
        var sourceText = null, sourceShape = null;
        var i;
        for (i = 0; i < selection.length; i++) {
            var item = selection[i];
            if (!sourceText && item.typename === "TextFrame") { sourceText = item; }
            else if (!sourceShape && workerIsRectanglePath(item)) { sourceShape = item; }
        }
        if (!sourceText || !sourceShape) { return created; }
        var styleInfo = workerResolveStyle(externalStyleName, sourceText);
        try {
            var framePath = sourceShape.duplicate();
            framePath.filled = false;
            framePath.stroked = false;
            var areaType = workerFillAreaType(doc, framePath, sourceText, styleInfo.name);
            workerCenterAreaType(areaType);
            if (styleInfo.isTemp) { workerApplyButtonShape(areaType); }
            sourceText.remove();
            sourceShape.remove();
            created.push(areaType);
        } catch (e) {
            /* ignore */
        } finally {
            if (styleInfo.isTemp) { workerRemoveStyleByName(styleInfo.name); }
        }
        return created;
    }
    WORKER_FUNCS.push(workerConvertTextIntoShape);

    /* worker: ポイント文字のみ → 計測した実寸でフレームを作り中央配置 / Point text -> frame at measured size, centered */
    function workerConvertPointText(doc, selection, adjust, widthRatio, heightRatio, externalStyleName) {
        var created = [];
        var wRatio = adjust ? widthRatio : 1;
        var hRatio = adjust ? heightRatio : 1;
        var i;
        for (i = selection.length - 1; i >= 0; i--) {
            var sourceText = selection[i];
            if (!(sourceText.typename === "TextFrame" && sourceText.kind === TextType.POINTTEXT)) { continue; }
            var styleInfo = workerResolveStyle(externalStyleName, sourceText);
            try {
                var measuredBounds = workerMeasureAccurateBounds(sourceText);
                if (!measuredBounds) { measuredBounds = workerGeometricBounds(sourceText); }
                var frameWidth = measuredBounds.width * wRatio;
                var frameHeight = measuredBounds.height * hRatio;
                var centerX = measuredBounds.left + measuredBounds.width / 2;
                var centerY = measuredBounds.top - measuredBounds.height / 2;
                var frameLeft = centerX - frameWidth / 2;
                var frameTop = centerY + frameHeight / 2;
                var framePath = doc.pathItems.rectangle(frameTop, frameLeft, frameWidth, frameHeight);
                framePath.filled = false;
                framePath.stroked = false;
                var areaType = workerFillAreaType(doc, framePath, sourceText, styleInfo.name);
                workerCenterAreaType(areaType);
                if (adjust) {
                    if (styleInfo.isTemp) { workerApplyButtonShape(areaType); }
                }
                created.push(areaType);
                sourceText.remove();
            } catch (e) {
                /* ignore */
            } finally {
                if (styleInfo.isTemp) { workerRemoveStyleByName(styleInfo.name); }
            }
        }
        return created;
    }
    WORKER_FUNCS.push(workerConvertPointText);

    /* worker: 選択をエリア内文字へ変換（エントリ）
       戻り値: "OK" / "NODOC" / "NOSEL" / "NOTEXT" / "RECTONLY" / "ERR:..."
       Convert the selection to area type (entry). Returns markers. */
    function workerConvertSelection(adjust, widthRatio, heightRatio, externalStyleName, undoFirst) {
        if (app.documents.length === 0) { return "NODOC"; }
        if (undoFirst) {
            try { app.undo(); app.redraw(); } catch (eU) { /* ignore */ }
        }
        var doc = app.activeDocument;
        var sel = doc.selection;
        if (!sel || sel.length === 0) { return "NOSEL"; }
        workerLoadAreaTextActions();
        var resultMarker = "OK";
        try {
            sel = workerPreprocessPathText(doc, sel);
            var selection = doc.selection;
            if (!selection || selection.length === 0) { return "NOSEL"; }
            var hasSourceText = false, hasFrameShape = false, hasNonRectShape = false;
            var i;
            for (i = 0; i < selection.length; i++) {
                var item = selection[i];
                if (item.typename === "TextFrame" && (item.kind === TextType.POINTTEXT || item.kind === TextType.PATHTEXT)) { hasSourceText = true; }
                if (workerIsRectanglePath(item)) { hasFrameShape = true; }
                else if (item.typename === "PathItem" && item.closed) { hasNonRectShape = true; }
            }
            if (!hasSourceText) { return "NOTEXT"; }
            if (hasNonRectShape && !hasFrameShape) { return "RECTONLY"; }
            var created = hasFrameShape
                ? workerConvertTextIntoShape(doc, selection, externalStyleName)
                : workerConvertPointText(doc, selection, adjust, widthRatio, heightRatio, externalStyleName);
            if (created.length > 0) {
                try { doc.selection = created; app.redraw(); } catch (e) { /* ignore */ }
            }
        } catch (eMain) {
            resultMarker = "ERR:" + eMain;
        } finally {
            workerUnloadAreaTextActions();
        }
        return resultMarker;
    }
    WORKER_FUNCS.push(workerConvertSelection);

    // =========================================
    // 設定ファイルとファイル選択 / Prefs & file picking
    // =========================================

    /* パスから表示用ファイル名を取得 / Get a display-friendly filename from a path */
    function getDisplayFileName(filePath) {
        try { return decodeURI(new File(filePath).name); } catch (e) { return filePath; }
    }

    /* 設定ファイル（前回のスタイルファイルのパスとスタイル名を記憶）/ Prefs file remembering the last style file and names */
    function getPrefsFile() {
        return new File(Folder.userData + "/" + PREFS_FILE_NAME);
    }

    /* 記憶しているスタイルファイルのパスとスタイル名を読み込む / Load the remembered style-file path and names */
    function loadSavedStyleState() {
        var state = { filePath: "", styleNames: [] };
        var prefsFile = getPrefsFile();
        if (!prefsFile.exists) return state;
        try {
            prefsFile.encoding = "UTF-8";
            prefsFile.open("r");
            var content = prefsFile.read();
            prefsFile.close();
            var lines = content.split(/\r\n|\r|\n/);
            for (var i = 0; i < lines.length; i++) {
                var separatorIndex = lines[i].indexOf("=");
                if (separatorIndex < 0) continue;
                var key = lines[i].substring(0, separatorIndex);
                var value = lines[i].substring(separatorIndex + 1);
                if (key === "styleFilePath") state.filePath = value;
                else if (key === "styleNames") state.styleNames = value ? value.split("\t") : [];
            }
        } catch (e) { }
        return state;
    }

    /* スタイルファイルのパスとスタイル名を記憶する（key=value 形式）/ Remember the style-file path and names (key=value) */
    function saveStyleState(filePath, styleNames) {
        var prefsFile = getPrefsFile();
        try {
            prefsFile.encoding = "UTF-8";
            prefsFile.open("w");
            prefsFile.write("styleFilePath=" + filePath + "\n");
            prefsFile.write("styleNames=" + styleNames.join("\t") + "\n");
            prefsFile.close();
        } catch (e) { }
    }

    /* スタイル用 AI ファイルを選ばせる（キャンセルで空文字）/ Let the user pick a style AI file */
    function pickStyleFile() {
        var picked = File.openDialog(L("ui.pickFile"), function (candidate) {
            return (candidate instanceof Folder) || /\.ai$/i.test(candidate.name);
        });
        return picked ? picked.fsName : "";
    }

    // =========================================
    // パレット / Palette
    // =========================================

    /* 大きさ調整「する」の既定倍率（幅・高さ）/ Default size-adjustment ratios (width, height) */
    var BUTTON_WIDTH_RATIO = 1.2;   // 元の幅に対する倍率 / Ratio of original width
    var BUTTON_HEIGHT_RATIO = 1.6;  // 元の高さに対する倍率 / Ratio of original height

    /* パネルの余白と間隔 / Panel margins and spacing */
    var PANEL_MARGINS = [16, 20, 16, 12];
    var PANEL_SPACING = 8;

    /* ボタンの高さを指定 px 詰める（レイアウト確定後に呼ぶ）/ Trim a button's height by the given px (call after layout) */
    function trimButtonHeight(button, px) {
        try {
            button.size = [button.size.width, button.size.height - px];
        } catch (e) { }
    }

    /* パネルの共通設定 / Apply shared panel layout */
    function setupPanel(panel, spacing) {
        panel.orientation = "column";
        panel.alignChildren = ["fill", "top"];
        panel.alignment = "fill";
        panel.margins = PANEL_MARGINS;
        panel.spacing = (typeof spacing === "number") ? spacing : PANEL_SPACING;
    }

    /* 常駐パレットを構築して表示。「変換」ボタンで選択をエリア内文字化（DOM 処理はメインエンジンへ委譲）
       「読み込み」ボタンで別の AI ファイルを選ぶと、その場でリストを組み直す
       Build and show the resident palette. The "Convert" button converts the selection to area type
       (DOM work is delegated to the main engine). The "Load" button re-imports and rebuilds the list */
    function showPalette(savedStyleState) {
        // すでにパレットが開いていれば前面に出して終了 / If the palette already exists, bring it forward and return
        try {
            if ($.global.__importAndApplyGraphicStylePalette && $.global.__importAndApplyGraphicStylePalette.visible) {
                $.global.__importAndApplyGraphicStylePalette.show();
                return;
            }
        } catch (eExisting) { /* ignore */ }

        var styleState = {
            filePath: (savedStyleState && savedStyleState.filePath) || "",
            styleNames: (savedStyleState && savedStyleState.styleNames) || []
        };

        var paletteWindow = new Window("palette", L("ui.dialogTitle") + " " + SCRIPT_VERSION, undefined, { resizeable: false });
        $.global.__importAndApplyGraphicStylePalette = paletteWindow;
        paletteWindow.orientation = "column";
        paletteWindow.alignChildren = "fill";
        paletteWindow.margins = 15;

        var areaTypeOptionPanel = paletteWindow.add("panel", undefined, L("ui.areaTypeOptionPanel"));
        setupPanel(areaTypeOptionPanel, 6);

        // マスターチェック：オンのときだけ他パーツを有効化 / Master checkbox: enables the other parts only when checked
        var convertCheckbox = areaTypeOptionPanel.add("checkbox", undefined, L("ui.convertToAreaType"));
        convertCheckbox.value = false; // 既定はオフ / Default: off

        // サイズ：調整する / 調整しない / Size: adjust / don't adjust
        var LABEL_WIDTH = 44; // サイズ／幅／高さのラベル幅を統一 / Shared label width for Size/Width/Height
        var adjustModeGroup = areaTypeOptionPanel.add("group");
        var sizeLabel = adjustModeGroup.add("statictext", undefined, L("ui.sizeLabel"));
        sizeLabel.preferredSize.width = LABEL_WIDTH;
        var adjustOnRadio = adjustModeGroup.add("radiobutton", undefined, L("ui.doAdjust"));
        var adjustOffRadio = adjustModeGroup.add("radiobutton", undefined, L("ui.dontAdjust"));
        adjustOffRadio.value = true; // 既定は「しない」/ Default: off

        // 幅・高さの倍率（1行に横並び、百分率 % で入力）/ Width and height ratios (one row, entered as %)
        var ratioRow = areaTypeOptionPanel.add("group");
        var widthLabel = ratioRow.add("statictext", undefined, L("ui.widthRatio"));
        var widthInput = ratioRow.add("edittext", undefined, String(Math.round(BUTTON_WIDTH_RATIO * 100)));
        widthInput.characters = 4;
        widthInput.preferredSize.width = 40;
        ratioRow.add("statictext", undefined, "%");

        var heightLabel = ratioRow.add("statictext", undefined, L("ui.heightRatio"));
        var heightInput = ratioRow.add("edittext", undefined, String(Math.round(BUTTON_HEIGHT_RATIO * 100)));
        heightInput.characters = 4;
        heightInput.preferredSize.width = 40;
        ratioRow.add("statictext", undefined, "%");

        // 変換ボタン（このパネル内・主アクション）。パネル幅いっぱいに広げず右寄せ・自然幅に
        // onClick は各パネル構築後に割り当て
        // Convert button (primary action, inside this panel); right-aligned at its natural width (not full width)
        // onClick is assigned after all panels are built
        var convertButton = areaTypeOptionPanel.add("button", undefined, L("ui.runConvert"), { name: "ok" });
        convertButton.alignment = ["right", "center"]; // 親の fill を上書きして広げない / Override the panel's fill so it doesn't stretch

        function updateAreaTypeOptionEnabled() {
            var convertOn = convertCheckbox.value;
            // マスターがオフなら する/しない ラジオごとディム / Dim the on/off radios when the master is off
            adjustOnRadio.enabled = convertOn;
            adjustOffRadio.enabled = convertOn;
            // 幅・高さはマスターオン かつ「する」のときのみ / Width/height only when the master is on and "On" is chosen
            widthInput.enabled = convertOn && adjustOnRadio.value;
            heightInput.enabled = convertOn && adjustOnRadio.value;
        }
        convertCheckbox.onClick = updateAreaTypeOptionEnabled;
        adjustOnRadio.onClick = updateAreaTypeOptionEnabled;
        adjustOffRadio.onClick = updateAreaTypeOptionEnabled;
        updateAreaTypeOptionEnabled();

        // グラフィックスタイル（元の見た目／読み込んだスタイル）/ Graphic style (original appearance / loaded styles)
        var stylePanel = paletteWindow.add("panel", undefined, L("ui.stylePanel"));
        setupPanel(stylePanel, 6);

        // 使用中のみ表示するフィルタ / Filter to show only styles used in the document
        var usedOnlyCheckbox = stylePanel.add("checkbox", undefined, L("ui.usedOnly"));
        usedOnlyCheckbox.value = false;

        // スタイル選択リスト（先頭が「元の見た目」、以降が現書類のグラフィックスタイル）
        // Style list (first row = original appearance, then the document's graphic styles)
        var styleListbox = stylePanel.add("listbox", undefined, [], { multiselect: false });
        styleListbox.preferredSize.height = 120;
        var styleListValues = []; // 各行に対応する値（null=元の見た目、以降はスタイル名）/ Value per row (null = original, else style name)

        // 最下部：左右分割（左＝クリア／スペーサー／右＝パネル）/ Bottom row: split layout (left = Clear, spacer, right = Panel)
        var styleFooterRow = stylePanel.add("group");
        styleFooterRow.orientation = "row";
        styleFooterRow.alignment = ["fill", "bottom"];
        styleFooterRow.margins = [0, 5, 0, 0]; // ボタンエリア上部にマージン / Top margin above the button area

        // 左側グループ：未使用のグラフィックスタイルを削除 / Left group: delete unused graphic styles
        var styleFooterLeft = styleFooterRow.add("group");
        styleFooterLeft.alignChildren = ["left", "center"];
        var clearButton = styleFooterLeft.add("button", undefined, L("ui.clearStyle"));
        clearButton.helpTip = L("ui.clearStyleHint");
        clearButton.onClick = function () {
            // 未使用削除はメインエンジンへ委譲（ダイナミックアクション）/ Delegate the unused-prune to the main engine (dynamic action)
            delegate("workerPruneUnused()");
            rebuildStyleList();
        };

        // スペーサー（伸縮）/ Spacer (stretchable)
        var styleFooterSpacer = styleFooterRow.add("group");
        styleFooterSpacer.alignment = ["fill", "fill"];
        styleFooterSpacer.minimumSize.width = 0;

        // 右側グループ：グラフィックスタイルパネルへ移動 / Right group: jump to the Graphic Styles panel
        var styleFooterRight = styleFooterRow.add("group");
        styleFooterRight.alignChildren = ["right", "center"];
        var openStylePanelButton = styleFooterRight.add("button", undefined, L("ui.openStylePanel"));
        openStylePanelButton.helpTip = L("ui.openStylePanelHint");
        openStylePanelButton.onClick = function () {
            // DOM/メニュー操作はメインエンジンへ委譲 / Delegate the DOM/menu op to the main engine
            delegate("workerOpenStylePanel()");
        };

        // スタイルの読み込みパネル（ファイル名を上、ボタンを下に表示）/ Load-styles panel (filename on top, buttons below)
        var loadPanel = paletteWindow.add("panel", undefined, L("ui.loadPanel"));
        setupPanel(loadPanel, 6);
        var fileNameText = loadPanel.add("statictext", undefined, "", { truncate: "middle" });
        fileNameText.preferredSize.width = 240;
        // 読み込み / 再読み込みボタンを左寄せで横並び / Load & Reload buttons in a left-aligned row
        var loadButtonRow = loadPanel.add("group");
        loadButtonRow.alignment = "left";
        loadButtonRow.margins = [0, 7, 0, 0];
        var loadButton = loadButtonRow.add("button", undefined, L("ui.loadButton"));
        loadButton.helpTip = L("ui.noStylesHint"); // 使い方はツールチップで案内 / Usage hint shown as a tooltip
        var reloadButton = loadButtonRow.add("button", undefined, L("ui.reloadButton"));
        reloadButton.helpTip = L("ui.reloadHint"); // 記憶したファイルから再取り込み / Re-import from the remembered file

        /* 選択中のファイル名表示と再読み込みボタンの有効状態を更新 / Update the file-name label and Reload's enabled state */
        function refreshFileLabel() {
            fileNameText.text = styleState.filePath ? getDisplayFileName(styleState.filePath) : L("ui.noFileSelected");
            reloadButton.enabled = !!styleState.filePath; // 記憶したファイルが無ければ再読み込み不可 / Disable Reload without a remembered file
        }

        // リビルド中は onChange の即時適用を抑止 / Suppress live-apply during a rebuild
        var isRebuildingStyleList = false;

        /* 現在のグラフィックスタイルでリストを組み直す（一覧取得はメインエンジンへ委譲）
           Rebuild the list from the current graphic styles (names fetched via the main engine) */
        function rebuildStyleList() {
            styleListbox.removeAll();
            styleListValues = [];
            // 先頭に「元の見た目」/ First row = original appearance
            styleListbox.add("item", L("ui.styleOriginal"));
            styleListValues.push(null);

            // 「使用中のみ」ON のときだけ未使用を除外（判定は worker 内で実施）
            // Skip unused only when the filter is on (classification happens inside the worker)
            var parsed = parseMarkerList(delegate("workerGetStyleNames(" + (usedOnlyCheckbox.value ? "true" : "false") + ")"));
            for (var j = 0; j < parsed.items.length; j++) {
                styleListbox.add("item", parsed.items[j]);
                styleListValues.push(parsed.items[j]);
            }
            isRebuildingStyleList = true;
            styleListbox.selection = 0; // 既定は「元の見た目」/ Default: original appearance
            isRebuildingStyleList = false;
        }
        usedOnlyCheckbox.onClick = function () { rebuildStyleList(); };

        // listbox で選んだグラフィックスタイルを、その時点の選択オブジェクトへ即適用（「適用」ボタン無し）
        // 「元の見た目」行・リビルド時は適用しない。選択なし等はクリックごとに警告せず無言で無視
        // Live-apply the picked style to the current selection (no Apply button);
        // skip the "Original" row and rebuilds, and ignore no-selection silently
        styleListbox.onChange = function () {
            if (isRebuildingStyleList) return;
            if (!styleListbox.selection) return;
            var pickedStyleName = styleListValues[styleListbox.selection.index];
            if (!pickedStyleName) return; // 「元の見た目」は適用対象なし / "Original appearance" has nothing to apply
            delegate("workerApplyStyleToSelection(" + jsStringLiteral(pickedStyleName) + ",false)");
        };

        // onClick で連結（addEventListener は発火しない環境があるため）/ Use onClick, not addEventListener
        loadButton.onClick = function () {
            var pickedPath = pickStyleFile();
            if (!pickedPath) return;
            // 取り込みはメインエンジンへ委譲 / Delegate the import to the main engine
            var parsed = parseMarkerList(delegate("workerImportStyles(" + jsStringLiteral(pickedPath) + ")"));
            if (!parsed.ok) {
                if (parsed.marker === "NOFILE") alert(L("alert.fileNotFound") + getDisplayFileName(pickedPath));
                return;
            }
            styleState.filePath = pickedPath;
            styleState.styleNames = parsed.items;
            saveStyleState(pickedPath, parsed.items); // 次回以降このファイルを参照 / Remember for next runs
            refreshFileLabel();
            rebuildStyleList();
        };

        // 記憶したファイルを選び直さずに再取り込み（別ドキュメントでも同じファイルを再利用）
        // Re-import from the remembered file without re-picking (reuse the same file in another document)
        reloadButton.onClick = function () {
            if (!styleState.filePath) return;
            var parsed = parseMarkerList(delegate("workerImportStyles(" + jsStringLiteral(styleState.filePath) + ")"));
            if (!parsed.ok) {
                if (parsed.marker === "NOFILE") alert(L("alert.fileNotFound") + getDisplayFileName(styleState.filePath));
                return;
            }
            styleState.styleNames = parsed.items;
            saveStyleState(styleState.filePath, parsed.items); // スタイル名の変化を反映 / Reflect any style-name changes
            refreshFileLabel();
            rebuildStyleList();
        };

        refreshFileLabel();
        rebuildStyleList();

        // 「変換」：現在の UI 設定で選択をエリア内文字化。DOM 処理はメインエンジンへ委譲し、パレットは開いたまま
        // "Convert": convert the selection using the current UI settings. DOM work is delegated; the palette stays open.
        convertButton.onClick = function () {
            // 幅・高さは百分率 % 入力を倍率へ換算 / Width/height: convert the % input to a ratio
            var wPercent = parseFloat(widthInput.text);
            var hPercent = parseFloat(heightInput.text);
            var widthRatio = (isNaN(wPercent) || wPercent <= 0) ? BUTTON_WIDTH_RATIO : wPercent / 100;
            var heightRatio = (isNaN(hPercent) || hPercent <= 0) ? BUTTON_HEIGHT_RATIO : hPercent / 100;
            // 「元の見た目」なら null、スタイル選択時はその名前 / null for original, else the selected style name
            var externalStyleName = null;
            if (styleListbox.selection) { externalStyleName = styleListValues[styleListbox.selection.index]; }
            var extName = externalStyleName ? externalStyleName : "";

            if (convertCheckbox.value) {
                // エリア内文字変換はメインエンジンへ委譲（スタイルは読み込み済みで現書類に存在）
                // Delegate the area-type conversion to the main engine (the style is already imported in the doc)
                var convRes = delegate("workerConvertSelection(" +
                    (adjustOnRadio.value ? "true" : "false") + "," +
                    widthRatio + "," +
                    heightRatio + "," +
                    jsStringLiteral(extName) + ",false)");
                if (convRes === "NODOC") { alert(L("alert.noDocument")); }
                else if (convRes === "RECTONLY") { alert(L("alert.rectangleOnly")); }
                else if (convRes === "NOSEL" || convRes === "NOTEXT") { alert(L("alert.selectText")); }
            } else if (externalStyleName) {
                // 変換OFF：選択したグラフィックスタイルを選択オブジェクトへ適用（メインエンジンへ委譲）
                // Convert off: apply the chosen graphic style to the selection (delegated to the main engine)
                var applyRes = delegate("workerApplyStyleToSelection(" + jsStringLiteral(externalStyleName) + ",false)");
                if (applyRes === "NODOC") { alert(L("alert.noDocument")); }
                else if (applyRes === "NOSEL") { alert(L("alert.selectText")); }
            }
        };

        // パレットをアクティブにして Esc キーで閉じる / Close the palette with Esc while it is active
        paletteWindow.addEventListener("keydown", function (event) {
            if (event && event.keyName === "Escape") { paletteWindow.close(); }
        });

        // パネル内のボタンはレイアウト確定後に高さを 2px 詰める / Trim panel buttons' height by 2px after layout
        paletteWindow.onShow = function () {
            trimButtonHeight(loadButton, 2);
            trimButtonHeight(reloadButton, 2);
            trimButtonHeight(clearButton, 2);
            trimButtonHeight(openStylePanelButton, 2);
        };
        // 閉じたらグローバル参照を解放 / Release the global reference when closed
        paletteWindow.onClose = function () {
            $.global.__importAndApplyGraphicStylePalette = null;
        };

        paletteWindow.show();
    }

    // =========================================
    // エントリポイント / Entry point
    // =========================================
    if (app.documents.length > 0) {
        // 記憶した参照ファイルとスタイル名を読み込み、常駐パレットを表示
        // 変換・スタイル適用は「変換」ボタンから実行（DOM 処理はメインエンジンへ委譲）
        // Load the remembered file/style names, then show the resident palette.
        // Conversion / style application runs from the "Convert" button (DOM work delegated to the main engine).
        var savedStyleState = loadSavedStyleState();
        showPalette(savedStyleState);
    } else {
        alert(L("alert.noDocument"));
    }

})();
