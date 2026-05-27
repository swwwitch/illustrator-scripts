#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

### スクリプト名：

FileNameManager.jsx — 現在のドキュメントのファイル名を変更／別名で保存

### GitHub：

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/FileNameManager.jsx

### 概要：

- アクティブドキュメントのファイル名を変更（実ファイルのリネーム）または別名で保存します。
- 保存先は常に現在のファイルと同じフォルダ。未保存ドキュメントは最初に保存先を確定します。
- % エンコードされたファイル名を自動デコードします。
- 「インクリメント」ON で末尾の数値を +1（桁数維持）。右側にプレビュー表示します。
- UI は「オプション」「ファイル名」パネルで構成し、ダイアログ不透明度は 0.98 に設定しています。

### 主な機能：

- ファイル名の変更（Save As 後に旧ファイルを削除して実質的にリネーム）
- 別名で保存（元ファイルは保持）
- 「新しい名前」／「インクリメント」を排他的チェックボックスで切替（ライブプレビュー）
- ラベル幅の自動整列による見やすいレイアウト

### 処理の流れ：

1) ドキュメント検証 → 現在名の取得とデコード
2) ダイアログ作成（createDialog）→ 入力・モード選択
3) 入力名のサニタイズ → インクリメント適用（任意）
4) 未保存時は一度保存してフォルダ確定 → 常に同フォルダへ保存
5) 「ファイル名の変更」時は旧ファイルを削除（実質リネーム）

### 更新履歴：

- v1.0 (20250816) : 初期バージョン
- v1.1 (20250816) : ローカライズ、UIパネル構成、ラベル幅の自動整列、透明度設定、%デコード、インクリメント＆プレビュー、同一フォルダ保存、リネーム実装（旧ファイル削除）、関数分離（createDialog）

---

### Script name:

FileNameManager.jsx — Rename or Save As the Current Document

### GitHub:

https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/document/FileNameManager.jsx

### Overview:

- Rename the active document (true rename by deleting the old file after Save As) or Save As.
- Destination is always the same folder as the current file; for unsaved docs, determine folder first.
- Automatically decodes percent-encoded file names.
- With "Increment" ON, increments the last numeric run (+1, preserving digits) with live preview.
- UI is composed of “Options” and “File Name” panels; dialog opacity is set to 0.98.

### Key features:

- Rename (emulated via Save As + remove original)
- Save As (keep original)
- Mutually-exclusive checkboxes: New Name / Increment (live preview)
- Auto-aligned label widths for a clean layout

### Flow:

1) Validate doc → get & decode current name
2) Build dialog (createDialog) → input & mode selection
3) Sanitize input → apply increment (optional)
4) For unsaved docs, save once to determine folder → always save into that folder
5) In Rename mode, remove the old file (true rename)

### Changelog:

- v1.0 (2025-08-16): Initial release
- v1.1 (2025-08-16): Localization, panelized UI, label-width alignment, opacity, percent-decoding, increment + preview, same-folder saving, rename behavior (delete original), function split (createDialog)

*/

var SCRIPT_VERSION = "v1.1";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    // 1) Dialog Title
    dialogTitle: {
        ja: "ファイル名の管理",
        en: "File Name Manager" 
    },

    // 2) Option panel
    optionPanelTitle: {
        ja: "オプション",
        en: "Options"
    },
    modeRename: {
        ja: "ファイル名の変更",
        en: "Rename"
    },
    modeSaveAs: {
        ja: "別名で保存",
        en: "Save As"
    },

    // 3) Filename panel
    filenamePanelTitle: {
        ja: "ファイル名",
        en: "File Name"
    },
    currentNameLabel: {
        ja: "現在のファイル名",
        en: "Current Name"
    },
    newNameLabel: {
        ja: "新しい名前",
        en: "New Name"
    },
    incrementLabel: {
        ja: "インクリメント",
        en: "Increment"
    },

    // 4) Buttons
    ok: {
        ja: "リネーム",
        en: "Rename"
    },
    cancel: {
        ja: "キャンセル",
        en: "Cancel"
    },

    // 5) Messages
    noDoc: {
        ja: "ドキュメントが開かれていません",
        en: "No document is open."
    },
    emptyName: {
        ja: "ファイル名が空です",
        en: "File name is empty."
    },
    chooseSaveDestCurrent: {
        ja: "現在のファイルの保存先を指定",
        en: "Choose destination for the current file"
    },
    confirmOverwrite: {
        ja: "同名ファイルが存在します。上書きしますか？",
        en: "A file with the same name exists. Overwrite?"
    },
    saveFailed: {
        ja: "保存に失敗しました",
        en: "Failed to save"
    }
};

function L(key) {
    var entry = LABELS[key];
    if (!entry) return key;
    return (lang === 'ja') ? entry.ja : entry.en;
}

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
        alert(L('noDoc'));
        return;
    }

    var doc = app.activeDocument;

    // 事前保存：ダイアログを開く前に現在のファイルを保存
    if (!doc.saved) {
        try {
            // 既存ファイルに上書き保存（初回保存でない場合）
            if (doc.fullName) doc.save();
        } catch (e_save) {}
        // まだ未保存なら、保存先を指定して初回保存
        if (!doc.saved) {
            var firstSave = File.saveDialog(L('chooseSaveDestCurrent'), '*.ai');
            if (!firstSave) return; // キャンセル
            var firstPath = File(firstSave.fsName);
            if (!/\.ai$/i.test(firstPath.name)) {
                firstPath = File(firstPath.parent.fsName + '/' + stripExtension(firstPath.name) + '.ai');
            }
            var initOpts = new IllustratorSaveOptions();
            doc.saveAs(firstPath, initOpts);
        }
    }

    var wasUnsaved = !doc.saved; // after pre-save, typically false
    var rawName = (doc.fullName ? doc.fullName.name : doc.name); // may be percent-encoded
    var currentName = tryDecode(rawName);
    var currentFolder = doc.fullName ? doc.fullName.parent : null;

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
        alert(L('emptyName'));
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
        var firstSave = File.saveDialog(L('chooseSaveDestCurrent'), '*.ai');
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
        var ow = confirm(L('confirmOverwrite') + '\n\n' + destFile.fsName);
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
                try {
                    oldFile.remove();
                } catch (er) {
                    /* 削除できない場合は黙って継続 */ }
            }
        }
    } catch (e) {
        alert(L('saveFailed') + '\n' + e);
    }
}

function createDialog(base, currentName, currentFolder) {
    var dlg = new Window('dialog', L('dialogTitle') + ' ' + SCRIPT_VERSION);
    var dialogOpacity = 0.98;
    dlg.opacity = dialogOpacity;
    dlg.orientation = 'column';
    dlg.alignChildren = 'fill';

    // モード選択（最上部）
    var modePanel = dlg.add('panel', undefined, L('optionPanelTitle'));
    modePanel.orientation = 'column';
    modePanel.alignChildren = 'left';
    modePanel.margins = [15, 20, 15, 10];
    var rbRename = modePanel.add('radiobutton', undefined, L('modeRename'));
    var rbSaveAs = modePanel.add('radiobutton', undefined, L('modeSaveAs'));
    var hasDigitsInBase = /[0-9]/.test(base);
    if (hasDigitsInBase) {
        rbRename.value = false;
        rbSaveAs.value = true;
    } else {
        rbRename.value = true;
        rbSaveAs.value = false;
    }

    var info = dlg.add('panel', undefined, L('filenamePanelTitle'));
    info.orientation = 'column';
    info.alignChildren = 'left';
    info.margins = [15, 20, 15, 10];

    var currentSubGroup = info.add('group');
    currentSubGroup.orientation = 'row';
    var lblCurrent = currentSubGroup.add('statictext', undefined, L('currentNameLabel'));
    var currentNameText = currentSubGroup.add('statictext', undefined, currentName);

    var nameIncGroup = info.add('group');
    nameIncGroup.alignment = 'left';
    nameIncGroup.orientation = 'column';
    nameIncGroup.alignChildren = ['left', 'center'];
    nameIncGroup.margins = [0, 10, 0, 10];

    var nameSubGroup = nameIncGroup.add('group');
    nameSubGroup.orientation = 'row';
    var enabledCheckbox = nameSubGroup.add('checkbox', undefined, L('newNameLabel'));
    var nameField = nameSubGroup.add('edittext', undefined, base);
    nameField.preferredSize.width = 200;

    var incSubGroup = nameIncGroup.add('group');
    incSubGroup.orientation = 'row';
    var incrementCheckbox = incSubGroup.add('checkbox', undefined, L('incrementLabel'));

    // ---- Normalize label widths (現在のファイル名 / 新しい名前 / インクリメント) ----
    var g = info.graphics;

    function measureW(s) {
        return g.measureString(s).width;
    }
    var labelW = Math.max(
        measureW(L('currentNameLabel')),
        measureW(L('newNameLabel')),
        measureW(L('incrementLabel'))
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
    enabledCheckbox.onClick = function() {
        if (enabledCheckbox.value) incrementCheckbox.value = false;
        updateIncPreview();
    };
    incrementCheckbox.onClick = function() {
        if (incrementCheckbox.value) enabledCheckbox.value = false;
        updateIncPreview();
    };
    nameField.onChanging = updateIncPreview;
    updateIncPreview();

    var btns = dlg.add('group');
    btns.alignment = 'center';
    var btnCancel = btns.add('button', undefined, L('cancel'), {
        name: 'cancel'
    });
    var btnOk = btns.add('button', undefined, L('ok'), {
        name: 'ok'
    });

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