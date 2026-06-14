#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### 概要 / Overview

- オブジェクトとテキストを同時に選択して実行すると、オブジェクトの見た目をグラフィックスタイルに登録し、テキストの文字列をそのスタイル名にするスクリプト。
- 「テキスト1点＋オブジェクト1点」の選択に加え、「グループ1点（中身がテキスト＋オブジェクト）」の選択にも対応。
- 複数のグループを選択した場合は、グループごとに登録を実行する。
- 同名スタイルが既に存在する場合は削除してから上書き登録するので、繰り返し実行しても重複しない。
- Registers the appearance of the selected object as a graphic style, using the selected text's content as the style name.
- Accepts a top-level selection of one TextFrame + one object, a single group containing one TextFrame + one object, or multiple groups (processed one by one).
- If a style with the same name already exists, it is removed first so repeated runs do not create duplicates.

### 処理の流れ / Process flow

1. 選択から登録ジョブ（テキスト＋スタイル対象オブジェクトの組）を集める。グループ1点はその中身、複数グループはグループごと（1件も無ければ終了） / Collect register jobs (text + style-target pairs) from the selection: a single group is unwrapped, multiple groups are processed one by one (exit if none).
2. 一時アクションを一度だけロード / Load the temporary action once.
3. ジョブごとに以下を実行 / For each job:
   1. テキストから文字列を取得（前後の空白を除去、空ならそのジョブはスキップ） / Get the text content (trimmed); skip the job if empty.
   2. 既存の同名スタイルを削除 / Remove the existing style with the same name, if any.
   3. スタイル対象オブジェクトだけを選択し直す（テキストの見た目は含めない） / Reselect only the style-target object (exclude the text's appearance).
   4. アクションを実行して無名のグラフィックスタイルを末尾に追加し、末尾を取得した文字列に改名 / Run the action to append an unnamed graphic style, then rename the last style to the text content.
4. アクションをアンロード / Unload the action.
5. 元の選択に戻す / Restore the original selection.

*/

// =========================================
// バージョンと設定 / Version & Settings
// =========================================

var SCRIPT_VERSION = "v1.1.0";

// =========================================
// ヘルパー / Helpers
// =========================================

/* 前後の空白を除去（ES3 には String.trim が無い） / Trim leading & trailing whitespace */
function trimText(value) {
    return String(value).replace(/^\s+/, '').replace(/\s+$/, '');
}

/* 選択からテキストとスタイル対象オブジェクトを取り出す
   Resolve the text frame and the style-target object from the selection.
   対応: (a) テキスト1点＋オブジェクト1点 / (b) グループ1点（中身がテキスト＋オブジェクト）
   見つからなければ null を返す / Returns { text: TextFrame, target: PageItem } or null. */
function resolveTextAndStyleTarget(selectedItems) {
    // グループ1点なら、その中身を仕分け対象にする / If a single group is selected, look inside it
    var candidates = selectedItems;
    if (selectedItems.length === 1 && selectedItems[0].typename === 'GroupItem') {
        candidates = selectedItems[0].pageItems;
    }

    if (candidates.length !== 2) {
        return null;
    }

    // テキストとスタイル対象オブジェクトに仕分け / Sort into text and style-target
    var textItem = null;
    var styleTarget = null;
    for (var i = 0; i < candidates.length; i++) {
        if (candidates[i].typename === 'TextFrame') {
            textItem = candidates[i];
        } else {
            styleTarget = candidates[i];
        }
    }
    if (!textItem || !styleTarget) {
        return null;
    }
    return { text: textItem, target: styleTarget };
}

/* 選択から登録ジョブ（テキスト＋スタイル対象オブジェクトの組）の配列を集める
   Collect register jobs (text + style-target pairs) from the selection.
   対応: (a) テキスト1点＋オブジェクト1点 / (b) グループ1点 / (c) 複数グループ（グループごと）
   Returns an array of { text, target }. */
function collectStyleJobs(selectedItems) {
    // 全てグループなら、グループごとに仕分け / If every selected item is a group, process each group
    var allGroups = selectedItems.length >= 1;
    for (var i = 0; i < selectedItems.length; i++) {
        if (selectedItems[i].typename !== 'GroupItem') {
            allGroups = false;
            break;
        }
    }

    var jobs = [];
    if (selectedItems.length >= 2 && allGroups) {
        for (var g = 0; g < selectedItems.length; g++) {
            var perGroup = resolveTextAndStyleTarget([selectedItems[g]]);
            if (perGroup) {
                jobs.push(perGroup);
            }
        }
    } else {
        var single = resolveTextAndStyleTarget(selectedItems);
        if (single) {
            jobs.push(single);
        }
    }
    return jobs;
}

// =========================================
// グラフィックスタイル関連 / Graphic Style Helpers
// =========================================

/* アクションを一時ファイルに書き出してロードする / Write & load the force-new-style action */
function loadForceNewGraphicStyleAction() {
    var actionData = '/version 3 /name [ 12 477261706869635374796c65 ] /isOpen 1 /actionCount 1 /action-1 { /name [ 17 4164644e6577576974686f75744e616d65 ] /keyIndex 0 /colorIndex 0 /isOpen 1 /eventCount 1 /event-1 { /useRulersIn1stQuadrant 0 /internalName (ai_plugin_styles) /localizedName [ 30 e382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382a4e383ab ] /isOpen 1 /isOn 1 /hasDialog 1 /showDialog 0 /parameterCount 1 /parameter-1 { /key 1835363957 /showInPalette 4294967295 /type (enumerated) /name [ 36 e696b0e8a68fe382b0e383a9e38395e382a3e38383e382afe382b9e382bfe382 a4e383ab ] /value 1 } } }';

    var actionFile = new File(Folder.temp.fsName + '/__tmp_register_style.aia');
    actionFile.open('w');
    actionFile.write(actionData);
    actionFile.close();
    app.loadAction(actionFile);
    actionFile.remove();
}

/* ロード済みアクションを実行（選択オブジェクトをスタイル登録） / Run the loaded action on current selection */
function runForceNewGraphicStyleAction() {
    app.doScript('AddNewWithoutName', 'GraphicStyle', false);
}

/* アクションセットをアンロード / Unload the action set */
function unloadForceNewGraphicStyleAction() {
    app.unloadAction('GraphicStyle', '');
}

/* 1ジョブ分の登録（アクションはあらかじめロード済みであること）
   Register one job (the action must already be loaded). */
function registerGraphicStyleFromJob(activeDoc, graphicStyles, job) {
    // テキストからスタイル名を取得（空なら何もしない） / Get the style name from the text
    var styleName = trimText(job.text.contents);
    if (styleName === '') {
        return;
    }

    // 既存の同名スタイルがあれば削除 / Remove the existing style with the same name, if any
    try { graphicStyles.getByName(styleName).remove(); }
    catch (e) { }

    // スタイル対象オブジェクトだけを選択（テキストの見た目を含めない）
    activeDoc.selection = [job.target];

    // アクションでスタイルが追加された場合のみ、末尾を取得した文字列に改名
    var beforeCount = graphicStyles.length;
    runForceNewGraphicStyleAction();
    if (graphicStyles.length > beforeCount) {
        graphicStyles[graphicStyles.length - 1].name = styleName;
    }
}

// =========================================
// メイン処理 / Main
// =========================================

(function () {
    if (app.documents.length === 0) {
        return;
    }

    var activeDoc = app.activeDocument;
    var graphicStyles = activeDoc.graphicStyles;

    var selectedItems = activeDoc.selection;

    // 選択から登録ジョブを集める（複数グループはグループごと）
    var jobs = collectStyleJobs(selectedItems);
    if (jobs.length === 0) {
        return;
    }

    // アクションをロード → ジョブごとに実行 → アンロード
    loadForceNewGraphicStyleAction();
    for (var j = 0; j < jobs.length; j++) {
        registerGraphicStyleFromJob(activeDoc, graphicStyles, jobs[j]);
    }
    unloadForceNewGraphicStyleAction();

    // 元の選択に戻す
    activeDoc.selection = selectedItems;
})();
