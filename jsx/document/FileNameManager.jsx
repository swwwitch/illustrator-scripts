#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*
### スクリプト名：

FileNameManager.jsx — Rename or SaveAs Current Document

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- アクティブドキュメントのファイル名を変更（実ファイルのリネーム）または別名で保存。
- 保存先は常に現在ファイルと同一フォルダ。未保存時は最初に保存先を確定。
- % エンコードされたファイル名を自動デコード。
- 「インクリメント」ON で末尾の数値を +1（桁数維持）。右側にプレビュー表示。
- UI は「オプション」「ファイル名」パネルで構成し、ダイアログ不透明度を 0.98 に設定。

### Main Features：

- Rename current file (remove old file after Save As to emulate true rename)
- Save As (keep original file)
- Mutually exclusive checkboxes: New Name / Increment (+ live preview)
- Auto-aligned label widths for neat layout

### 処理の流れ：

1) ドキュメント検証 → 現在名の取得とデコード
2) ダイアログ作成（createDialog）→ 入力・モード選択
3) 入力名のサニタイズ → インクリメント適用（任意）
4) 未保存時は一度保存してフォルダ確定 → 常に同フォルダへ保存
5) モードが「ファイル名の変更」のときは旧ファイルを削除（実質リネーム）

### 更新履歴：

- v1.0 (20250816) : 初期バージョン
*/


// ---- Helpers ----

function incrementLastNumber(s) {
    s = String(s);
    // scan from end to find the last contiguous digit run
    var i = s.length - 1;
    while (i >= 0 && !isDigit(s.charAt(i))) i--;
    if (i < 0) {
        // no digits; append 1
        return s + '1';
    }
    var end = i;
    while (i >= 0 && isDigit(s.charAt(i))) i--;
    var start = i + 1;
    var numStr = s.substring(start, end + 1);
    var width = numStr.length;
    var n = parseInt(numStr, 10);
    if (isNaN(n)) n = 0;
    n = n + 1;
    var newNumStr = padLeft(String(n), width);
    return s.substring(0, start) + newNumStr + s.substring(end + 1);
}

function isDigit(ch) {
    var c = ch.charCodeAt(0);
    return c >= 48 && c <= 57; // '0'..'9'
}

function padLeft(str, width) {
    while (str.length < width) str = '0' + str;
    return str;
}

function tryDecode(s) {
    s = String(s);
    if (s.indexOf('%') === -1) return s;
    try {
        return decodeURIComponent(s);
    } catch (e) {
        try {
            return decodeURI(s);
        } catch (e2) {
            return s; // fallback
        }
    }
}

function stripExtension(name) {
    var i = name.lastIndexOf('.');
    return (i > 0) ? name.substring(0, i) : name;
}

function sanitizeFilename(s) {
    // Remove illegal characters: \\/:*?"<>|
    s = String(s);
    var illegal = '\\/:*?"<>|'; // characters to strip
    var out = '';
    for (var i = 0; i < s.length; i++) {
        var ch = s.charAt(i);
        if (illegal.indexOf(ch) === -1) out += ch;
    }
    // trim leading/trailing whitespace
    return out.replace(/^\s+|\s+$/g, '');
}

function selectByText(dropdown, text) {
    if (!dropdown || !dropdown.items) return;
    text = String(text).toLowerCase();
    for (var i = 0; i < dropdown.items.length; i++) {
        if (String(dropdown.items[i].text).toLowerCase() === text) {
            dropdown.selection = dropdown.items[i];
            return;
        }
    }
}

function saveAsWithExtension(doc, fileObj, ext) {
    switch (String(ext).toLowerCase()) {
        case 'ai':
            var aiOpt = new IllustratorSaveOptions();
            // ここで必要ならオプションを設定 / set options if needed
            doc.saveAs(fileObj, aiOpt);
            break;
        default:
            throw new Error('未対応の拡張子 / Unsupported extension: ' + ext);
    }
}

function main() {
    if (app.documents.length === 0) {
        alert('ドキュメントが開かれていません');
        return;
    }

    var doc = app.activeDocument;
    var wasUnsaved = !doc.saved;
    var rawName = wasUnsaved ? doc.name : doc.fullName.name; // may be percent-encoded
    var currentName = tryDecode(rawName); // decode %E3%... style names
    var currentFolder = wasUnsaved ? null : doc.fullName.parent;
    var base = stripExtension(currentName);

    // ---- Dialog ----
    var ui = createDialog(base, currentName, currentFolder);
    var dlg = ui.dlg;
    var nameField = ui.nameField;
    var incrementCheckbox = ui.incrementCheckbox;
    var rbSaveAs = ui.rbSaveAs;

    if (dlg.show() !== 1) return; // canceled

    var newBase = sanitizeFilename(nameField.text);
    if (!newBase) {
        alert('ファイル名が空です');
        return;
    }
    if (incrementCheckbox && incrementCheckbox.value) {
        newBase = incrementLastNumber(newBase);
    }
    var newExt = 'ai';

    var destFile;
    var targetFolder = currentFolder;
    if (wasUnsaved) {
        // 未保存の場合は、まず現在のドキュメントを保存して保存先フォルダを確定させる
        var firstSave = File.saveDialog('現在のファイルの保存先を指定', '*.ai');
        if (!firstSave) return; // キャンセル
        var firstPath = File(firstSave.fsName);
        if (!/\.ai$/i.test(firstPath.name)) {
            firstPath = File(firstPath.parent.fsName + '/' + stripExtension(firstPath.name) + '.ai');
        }
        // 現在のドキュメントを一度保存
        var tmpOpts = new IllustratorSaveOptions();
        doc.saveAs(firstPath, tmpOpts);
        targetFolder = firstPath.parent; // フォルダを確定
    }

    // 保存先は常に「現在のファイル」と同じフォルダ
    destFile = File(targetFolder.fsName + '/' + newBase + '.ai');

    var oldFsPath = (!wasUnsaved && doc.fullName) ? doc.fullName.fsName : null;

    // 既存同名の確認（どのモードでも）
    if (destFile.exists && (!oldFsPath || destFile.fsName !== oldFsPath)) {
        var ow = confirm('同名ファイルが存在します。上書きしますか？\n\n' + destFile.fsName);
        if (!ow) return;
    }

    try {
        // 保存（これによりアクティブドキュメントの関連付けは新ファイルになる）
        saveAsWithExtension(doc, destFile, newExt);

        // モード: ファイル名の変更（rbSaveAsがOFF）→ 元のファイルを削除して“リネーム”扱いに
        var isRenameMode = !(rbSaveAs && rbSaveAs.value);
        if (isRenameMode && oldFsPath && oldFsPath !== destFile.fsName) {
            var oldFile = File(oldFsPath);
            if (oldFile.exists) {
                try { oldFile.remove(); } catch (er) { /* 削除できない場合は黙って継続 */ }
            }
        }
    } catch (e) {
        alert('保存に失敗しました\n' + e);
    }
}

function createDialog(base, currentName, currentFolder) {
    var dlg = new Window('dialog', 'ファイル名の変更');
    var dialogOpacity = 0.98;
    dlg.opacity = dialogOpacity;
    dlg.orientation = 'column';
    dlg.alignChildren = 'fill';

    // モード選択（最上部）
    var modePanel = dlg.add('panel', undefined, 'オプション');
    modePanel.orientation = 'column';
    modePanel.alignChildren = 'left';
    modePanel.margins = [15, 20, 15, 10];
    var rbRename = modePanel.add('radiobutton', undefined, 'ファイル名の変更');
    var rbSaveAs = modePanel.add('radiobutton', undefined, '別名で保存');
    var hasDigitsInBase = /[0-9]/.test(base);
    if (hasDigitsInBase) { rbRename.value = false; rbSaveAs.value = true; } else { rbRename.value = true; rbSaveAs.value = false; }

    var info = dlg.add('panel', undefined, 'ファイル名');
    info.orientation = 'column';
    info.alignChildren = 'left';
    info.margins = [15, 20, 15, 10];

    var currentSubGroup = info.add('group');
    currentSubGroup.orientation = 'row';
    var lblCurrent = currentSubGroup.add('statictext', undefined, '現在のファイル名');
    var currentNameText = currentSubGroup.add('statictext', undefined, currentName);

    var nameIncGroup = info.add('group');
    nameIncGroup.alignment = 'left';
    nameIncGroup.orientation = 'column';
    nameIncGroup.alignChildren = ['left', 'center'];
    nameIncGroup.margins = [0, 10, 0, 10];

    var nameSubGroup = nameIncGroup.add('group');
    nameSubGroup.orientation = 'row';
    var enabledCheckbox = nameSubGroup.add('checkbox', undefined, '新しい名前');
    var nameField = nameSubGroup.add('edittext', undefined, base);
    nameField.preferredSize.width = 200;

    var incSubGroup = nameIncGroup.add('group');
    incSubGroup.orientation = 'row';
    var incrementCheckbox = incSubGroup.add('checkbox', undefined, 'インクリメント');

    // ---- Normalize label widths (現在のファイル名 / 新しい名前 / インクリメント) ----
    var g = info.graphics;
    function measureW(s) { return g.measureString(s).width; }
    var labelW = Math.max(
        measureW('現在のファイル名'),
        measureW('新しい名前'),
        measureW('インクリメント')
    ) + 12; // padding
    if (lblCurrent) lblCurrent.preferredSize = [labelW, lblCurrent.preferredSize.height || 20];
    if (enabledCheckbox) enabledCheckbox.preferredSize = [labelW, enabledCheckbox.preferredSize.height || 20];
    if (incrementCheckbox) incrementCheckbox.preferredSize = [labelW, incrementCheckbox.preferredSize.height || 20];

    // 初期状態：数字があればインクリメントON、なければ新しい名前ON
    enabledCheckbox.value = !hasDigitsInBase;
    incrementCheckbox.value = hasDigitsInBase;
    var incPreview = incSubGroup.add('statictext', undefined, '');

    function updateIncPreview() {
        if (incrementCheckbox.value) {
            incPreview.text = incrementLastNumber(sanitizeFilename(nameField.text));
        } else {
            incPreview.text = '';
        }
    }
    enabledCheckbox.onClick = function () { if (enabledCheckbox.value) incrementCheckbox.value = false; updateIncPreview(); };
    incrementCheckbox.onClick = function () { if (incrementCheckbox.value) enabledCheckbox.value = false; updateIncPreview(); };
    nameField.onChanging = updateIncPreview;
    updateIncPreview();

    var btns = dlg.add('group');
    btns.alignment = 'center';
    var btnCancel = btns.add('button', undefined, 'キャンセル', { name: 'cancel' });
    var btnOk = btns.add('button', undefined, 'OK', { name: 'ok' });

    return {
        dlg: dlg,
        nameField: nameField,
        incrementCheckbox: incrementCheckbox,
        rbSaveAs: rbSaveAs,
        enabledCheckbox: enabledCheckbox
    };
}

// 実行
main();

/* 改善提案 / Improvement Suggestions
- 失敗時の詳細ログ（ファイル権限・ディスク容量など）を追加
- macOS では旧ファイル削除を「ゴミ箱へ移動」に変更（安全性向上）
- IllustratorSaveOptions の外出し（設定プリセット化）
- 入力検証強化：最大長・先頭ドット・予約語（CON/NUL 等 Windows 互換）
- UI 幅の可変対応（長いファイル名の折返し抑止やスクロール対応）
*/