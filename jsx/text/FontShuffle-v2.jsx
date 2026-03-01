/*
  RandomFonts.jsx
  選択したテキストの1文字ごとにランダムなフォントを適用するスクリプト

  - ダイアログボックスを表示
  - ［再実行］ボタンでランダム適用（プレビュー更新）
*/
(function () {
    // ドキュメントが開かれているか確認
    if (app.documents.length === 0) {
        alert("ドキュメントが開かれていません。");
        return;
    }

    var doc = app.activeDocument;

    // インストールされている全フォントを取得（数が多いと処理に時間がかかる場合があります）
    var allFonts = app.textFonts;
    var fontCount = allFonts.length;

    function getMorisawaFonts() {
        var list = [];
        var i, f;

        function matchJAKeyword(s) {
            if (!s) return false;
            s = String(s);
            // 和文フォント判定キーワード
            // Pr6 / Pr6N / AB- / FOT を含むものを対象
            return (
                s.indexOf('Pr6N') !== -1 ||
                s.indexOf('Pr6') !== -1 ||
                s.indexOf('AB-') !== -1 ||
                s.indexOf('FOT') !== -1 ||
                s.indexOf('-OTF') !== -1
            );
        }

        for (i = 0; i < allFonts.length; i++) {
            try {
                f = allFonts[i];
                if (!f) continue;

                var n1 = f.fullName ? String(f.fullName) : '';
                var n2 = f.name ? String(f.name) : '';
                var n3 = (f.postScriptName) ? String(f.postScriptName) : '';

                if (matchJAKeyword(n1) || matchJAKeyword(n2) || matchJAKeyword(n3)) {
                    list.push(f);
                }
            } catch (_) { }
        }

        return list;
    }

    function getSelectionTextFrames() {
        var sel = doc.selection;
        if (!sel || sel.length === 0) return [];
        var frames = [];
        for (var i = 0; i < sel.length; i++) {
            try {
                if (sel[i] && sel[i].typename === "TextFrame") frames.push(sel[i]);
            } catch (_) { }
        }
        return frames;
    }

    function applyRandomFontsToFrames(frames, fontList) {
        if (!frames || frames.length === 0) return;
        if (!fontList || fontList.length === 0) return;

        for (var i = 0; i < frames.length; i++) {
            var item = frames[i];
            if (!item || item.typename !== "TextFrame") continue;

            var chars = item.textRange.characters;
            for (var j = 0; j < chars.length; j++) {
                // 改行や空白文字はスキップ（エラー回避と見た目のため）
                if (chars[j].contents === "\r" || chars[j].contents === " ") continue;

                try {
                    var randomFontIndex = Math.floor(Math.random() * fontList.length);
                    chars[j].characterAttributes.textFont = fontList[randomFontIndex];
                } catch (_) {
                    // 特定のフォントが適用できない場合は無視
                }
            }
        }

        app.redraw();
    }

    // --- Dialog ---
    var dlg = new Window('dialog', 'Random Fonts');
    dlg.orientation = 'column';
    dlg.alignChildren = ['fill', 'top'];

    var info = dlg.add('statictext', undefined, '選択したテキストに対して、1文字ごとにランダムなフォントを適用します。');
    info.characters = 52;

    var chkMorisawa = dlg.add('checkbox', undefined, '和文フォントに限る');
    chkMorisawa.value = false;

    // --- Bottom bar (3 columns) ---
    var bottom = dlg.add('group');
    bottom.orientation = 'row';
    bottom.alignChildren = ['fill', 'center'];

    // Left
    var colL = bottom.add('group');
    colL.orientation = 'row';
    colL.alignChildren = ['left', 'center'];
    var btnRerun = colL.add('button', undefined, '再実行');

    // Center spacer
    var colC = bottom.add('group');
    colC.orientation = 'row';
    colC.alignment = ['fill', 'center'];
    colC.add('statictext', undefined, '');

    // Right
    var colR = bottom.add('group');
    colR.orientation = 'row';
    colR.alignChildren = ['right', 'center'];
    var btnCancel = colR.add('button', undefined, 'キャンセル', { name: 'cancel' });
    var btnOK = colR.add('button', undefined, 'OK', { name: 'ok' });

    try { colC.preferredSize.width = 20; } catch (_) { }
    try { bottom.alignment = ['fill', 'bottom']; } catch (_) { }
    try { colR.alignment = ['right', 'center']; } catch (_) { }

    // Preview state: 直前のプレビューを undo で戻してから再適用する
    var _previewApplied = false;

    function runPreview() {
        var frames = getSelectionTextFrames();
        if (frames.length === 0) {
            alert('テキストオブジェクトを選択してください。');
            return;
        }

        if (_previewApplied) {
            try { app.undo(); } catch (_) { }
        }

        var fontList = null;
        if (chkMorisawa && chkMorisawa.value) {
            fontList = getMorisawaFonts();
            if (!fontList || fontList.length === 0) {
                alert('対象の和文フォント（Pr6 / Pr6N）が見つかりません。');
                return;
            }
        } else {
            fontList = allFonts;
        }

        applyRandomFontsToFrames(frames, fontList);
        _previewApplied = true;
    }


    btnRerun.onClick = function () {
        // プレビュー更新（同じ扱い）
        runPreview();
    };

    btnOK.onClick = function () {
        // プレビュー未実行ならここで1回適用
        if (!_previewApplied) {
            runPreview();
        }
        dlg.close(1);
    };

    btnCancel.onClick = function () {
        // プレビューを戻す
        if (_previewApplied) {
            try { app.undo(); } catch (_) { }
        }
        dlg.close(0);
    };

    dlg.center();
    dlg.show();
})();
