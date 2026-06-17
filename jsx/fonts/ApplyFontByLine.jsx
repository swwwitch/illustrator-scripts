#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

/*

ApplyFontByLine.jsx — 行ごとにフォントを適用するスクリプト

概要
選択したテキストフレーム（グループ内も再帰的に対象。ロック・非表示は除外）の
各行（段落）について、その行の文字列をフォント名とみなして検索し、行単位で適用する。

主な機能
・CUSTOM_MAP による特定文字列の強制割り当て
・PostScript 名、ファミリー名＋スタイル名、ファミリー名による厳密一致
・ファミリー名／フォント名全体／先頭ワードによるあいまい一致
・厳密一致は自動適用、あいまい一致・未一致は対話ピッカーで確認
・同じ文字列は実行中に再質問せず、選択結果を再利用
・未適用の行を含むフレームには「// missing-fonts」レイヤー上に赤い目印を作成
・未適用文字列を一覧表示し、クリップボードへコピー可能

履歴
v1.0.0  初期バージョン
v1.1.1  現行バージョン

*/

(function () {

    // ============================================================
    // 設定 / バージョン / ローカライズ（Settings / Version / Localization）
    // ============================================================
    var SCRIPT_VERSION = "v1.1.1";

    // クラッシュ箇所を特定するためのデバッグログ（true でデスクトップにログ出力）。
    // 各ステップで「開く→書く→閉じる」を行い、その都度ディスクへ確定させるので、
    // Illustrator が落ちてもログの最終行までは残る＝その次の処理が犯人。
    var DEBUG_LOG = true;

    // ============================================================
    // 【カスタム置換ルール】
    // 特定の行のテキストに対して、強制的に適用したいフォントを定義できます。
    // ============================================================
    var CUSTOM_MAP = {
        "Jenson": "Adobe Jenson Pro",
        "Garamond": "Adobe Garamond Pro",
        "Myriad": "Myriad Pro",
        "Frutiger": "Neue Frutiger World",
        "FF DIN": "DIN 2014",
        "Minion": "Minion Pro",
    };

    // 未適用の行を含むフレームの「目印」を置くレイヤー名と長方形の不透明度
    var MARKER_LAYER_NAME = "// missing-fonts";
    var MARKER_OPACITY = 35; // ％

    // 対話ピッカーで対象へズームするときの、表示領域に対する占有率（0〜1）
    var ZOOM_FIT_RATIO = 0.6;

    // ファミリー名一致などで複数スタイルが候補になったとき、どのスタイルを選ぶかの優先順位。
    // 先頭にあるスタイルほど優先される（小文字で比較。どれにも該当しなければ候補の先頭）。
    // Regular 優先にしたいときは ["regular", "medium", "semibold", "bold"] に並べ替える。
    var STYLE_PRIORITY = ["bold", "semibold", "medium", "regular"];

    // showFontPicker が「終了」されたときに返す番兵（選択ループを打ち切る合図）
    var PICKER_QUIT = {};

    var fontIndex;

    // ============================================================
    // ローカライズ
    // ============================================================

    var currentLanguage = (String(app.locale).indexOf("ja") === 0) ? "ja" : "en";

    var LABELS = {
        message: {
            noDoc: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            noSelection: {
                ja: "テキストオブジェクトを選択してください。",
                en: "Please select a text object."
            }
        },
        progress: {
            title: {
                ja: "フォントを適用中…",
                en: "Applying fonts…"
            }
        },
        result: {
            header: {
                ja: "フォントを適用できなかった文字列：",
                en: "Strings with no matching font:"
            }
        },
        button: {
            copy: {
                ja: "クリップボードにコピー",
                en: "Copy to clipboard"
            },
            copied: {
                ja: "コピーしました",
                en: "Copied"
            },
            close: {
                ja: "閉じる",
                en: "Close"
            }
        },
        picker: {
            title: {
                ja: "フォントを選択",
                en: "Choose font"
            },
            target: {
                ja: "対象テキスト：",
                en: "Target text:"
            },
            replace: {
                ja: "適用するフォント",
                en: "Font to apply"
            },
            search: {
                ja: "検索：",
                en: "Search:"
            },
            family: {
                ja: "フォント：",
                en: "Font:"
            },
            style: {
                ja: "スタイル：",
                en: "Style:"
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

    // ラベルのリーフ（{ ja, en }）を渡すと現在の言語の文字列を返す。
    // 現在の言語が未定義の場合は英語へ、それも無ければ空文字へフォールバックする。
    // 文中の {slash} は "/" に置換する。
    function L(labelNode) {
        var text = "";
        if (labelNode) {
            text = labelNode[currentLanguage] || labelNode.en || "";
        }
        return String(text).replace(/\{slash\}/g, "/");
    }

    if (app.documents.length === 0) {
        alert(L(LABELS.message.noDoc));
        return;
    }

    var doc = app.activeDocument;
    fontIndex = createFontIndex();
    var currentSelection = doc.selection;
    var selection = [];
    if (currentSelection) {
        for (var i = 0; i < currentSelection.length; i++) {
            selection.push(currentSelection[i]);
        }
    }

    if (selection.length === 0) {
        alert(L(LABELS.message.noSelection));
        return;
    }

    // 選択物からテキストフレームを再帰収集（グループ内も対象にする）
    var textFrames = [];
    collectTextFrames(selection, textFrames);

    if (textFrames.length === 0) {
        alert(L(LABELS.message.noSelection));
        return;
    }

    // フェーズ1：プログレスバー表示中は「厳密一致だけ自動適用」。
    // あいまい／未一致の行は保留キュー（pendingPicks）に貯め、確認は後回しにする。
    var progress = createProgress(textFrames.length);
    var pendingPicks = [];
    for (var i = 0; i < textFrames.length; i++) {
        progress.bar.value = i + 1;
        progress.countLabel.text = (i + 1) + " / " + textFrames.length;
        progress.window.update();
        autoApplyAndQueue(textFrames[i], pendingPicks);
    }
    progress.window.close();

    // フェーズ2：プログレスバー終了後に、保留分を対話ピッカーで1件ずつ決めさせる。
    // スキップされた行（＝未適用）はフレームと文字列を控える。
    // pickMemo は「同じ文字列をこの実行中に一度決めたら、以降は再質問せず結果を再利用」する。
    var unappliedFrames = [];
    var unappliedTexts = [];
    var pickMemo = {}; // lineText -> TextFont（適用）/ null（スキップ）
    var picker = null; // ピッカーは初回だけ生成して全ピックで使い回す（dropdownlist 再生成クラッシュ対策）
    dlog("=== run start === pendingPicks=" + pendingPicks.length);
    for (var i = 0; i < pendingPicks.length; i++) {
        var pick = pendingPicks[i];
        dlog("pick[" + i + "] index=" + pick.index + " text=[" + pick.lineText + "]");

        // 同じ文字列を既に決めていれば、その結果を再利用（ピッカーを出さない）
        if (pickMemo.hasOwnProperty(pick.lineText)) {
            var remembered = pickMemo[pick.lineText];
            var reapplied = false;
            if (remembered) {
                // 段落は保持参照を使い回さず、使用直前にライブ取得する（無効化クラッシュ対策）
                var rememberParagraph = getLiveParagraph(pick.frame, pick.index);
                if (rememberParagraph) {
                    try {
                        rememberParagraph.characterAttributes.textFont = remembered;
                        reapplied = true;
                    } catch (e) { }
                }
            }
            if (!reapplied) recordUnapplied(unappliedFrames, unappliedTexts, pick);
            continue;
        }

        // 初登場の文字列はピッカーで決めさせ、結果を記憶する
        dlog("  -> showFontPicker open");
        if (!picker) {
            dlog("  -> createFontPickerDialog (once, families=" + fontIndex.families.length + ")");
            picker = createFontPickerDialog(fontIndex.families);
            dlog("  -> createFontPickerDialog done");
        }
        var picked = showFontPicker(picker, pick.frame, pick.index, pick.lineText, pick.initialFont);
        dlog("  -> showFontPicker closed: " + (picked === PICKER_QUIT ? "QUIT" : (picked === null ? "SKIP" : "APPLIED")));
        if (picked === PICKER_QUIT) {
            // 「終了」：このピック以降の保留分をすべて未適用として記録し、選択を打ち切る
            for (var q = i; q < pendingPicks.length; q++) {
                recordUnapplied(unappliedFrames, unappliedTexts, pendingPicks[q]);
            }
            break;
        }
        pickMemo[pick.lineText] = picked; // TextFont or null
        if (picked === null) {
            recordUnapplied(unappliedFrames, unappliedTexts, pick);
        }
    }

    dlog("loop done. unappliedFrames=" + unappliedFrames.length + " unappliedTexts=" + unappliedTexts.length);

    // 未適用の行を含むフレームには、背面に赤・半透明（MARKER_OPACITY%）の長方形を置いて目印にする
    // （テキスト自体は変更しない非破壊マーカー）。実行のたびに古い目印レイヤーは削除する。
    dlog("before removeLayerByName");
    removeLayerByName(MARKER_LAYER_NAME);
    dlog("after removeLayerByName");

    // 未適用があれば、目印レイヤーを新規作成して長方形を置く
    if (unappliedFrames.length > 0) {
        dlog("before createMarkerLayer");
        var markerLayer = createMarkerLayer(MARKER_LAYER_NAME);
        dlog("after createMarkerLayer");
        for (var i = 0; i < unappliedFrames.length; i++) {
            dlog("  createMarkerRect[" + i + "]");
            createMarkerRect(unappliedFrames[i], markerLayer);
        }
        // 目印を誤って動かさないよう、作成後にレイヤーをロック
        dlog("before markerLayer.locked");
        markerLayer.locked = true;
        dlog("after markerLayer.locked");
    }

    // 適用できなかった文字列のみをダイアログで表示（コピー用ボタン付き）
    if (unappliedTexts.length > 0) {
        dlog("before showResultDialog");
        showResultDialog(unappliedTexts);
        dlog("after showResultDialog");
    }
    dlog("=== run end ===");

    // ============================================================
    // 関数（Functions）
    // ============================================================

    // デバッグログを1行追記する。書き込みごとに close() してディスクへ確定させる。
    function dlog(message) {
        if (!DEBUG_LOG) return;
        try {
            var logFile = new File(Folder.desktop + "/ApplyFontByLine-debug.log");
            logFile.encoding = "UTF-8";
            logFile.open("a");
            logFile.writeln(message);
            logFile.close();
        } catch (e) { }
    }

    // 適用中に表示するプログレスバー（パレット）を作成して返す
    function createProgress(total) {
        var progressWindow = new Window("palette", L(LABELS.progress.title) + " " + SCRIPT_VERSION);
        progressWindow.alignChildren = "fill";
        progressWindow.margins = 15;

        var progressBar = progressWindow.add("progressbar", undefined, 0, total);
        progressBar.preferredSize = [320, 9];

        var countLabel = progressWindow.add("statictext", undefined, "0 / " + total);
        countLabel.preferredSize.width = 320;

        progressWindow.show();
        progressWindow.update();
        return { window: progressWindow, bar: progressBar, countLabel: countLabel };
    }

    // 適用できなかった文字列を一覧表示するダイアログ（コピー用ボタン付き）
    function showResultDialog(lines) {
        var listText = lines.join("\n");

        var dialog = new Window("dialog", L(LABELS.result.header) + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";
        dialog.margins = 15;

        dialog.add("statictext", undefined, L(LABELS.result.header));

        // 一覧（読み取り専用・複数行・スクロール可）。手動選択もできる
        var listField = dialog.add("edittext", undefined, listText, { multiline: true, scrolling: true, readonly: true });
        listField.preferredSize = [380, 220];

        // === ボタン（Mac 規約：左にコピー、右に閉じる）===
        var buttonGroup = dialog.add("group");
        buttonGroup.alignment = "right";
        var copyButton = buttonGroup.add("button", undefined, L(LABELS.button.copy));
        buttonGroup.add("button", undefined, L(LABELS.button.close), { name: "ok" });

        copyButton.onClick = function () {
            copyTextToClipboard(listText);
            copyButton.text = L(LABELS.button.copied);
            copyButton.enabled = false;
        };

        dialog.show();
    }

    // 文字列をクリップボードへコピーする。
    // 一時テキストフレームを作って app.copy() する方式（pbcopy は非同期で失敗するため）
    function copyTextToClipboard(textToCopy) {
        var tempFrame = null;
        try {
            tempFrame = doc.textFrames.add();
            tempFrame.contents = textToCopy;
            tempFrame.position = [-100000, -100000]; // 画面外に逃がす
            doc.selection = null;
            tempFrame.selected = true;
            app.copy();
        } catch (e) {
        } finally {
            if (tempFrame) {
                try { tempFrame.remove(); } catch (e2) { }
            }
            doc.selection = null;
        }
    }

    // 候補が曖昧／未適用の行について、フォントを対話的に選ばせる。
    // 「適用」で選んだ TextFont を返し、「スキップ」／閉じるでは null を返す。
    // モーダルダイアログ表示中はドキュメントを一切変更しない（ライブプレビューは
    // Illustrator のハードクラッシュ要因のため廃止）。実適用は閉じた後にまとめて行う。
    // ダイアログ（picker）は作り直さず使い回す。ここでは対象テキストと選択をリセットするだけ。
    function showFontPicker(picker, targetFrame, paragraphIndex, lineText, initialFont) {
        // 対象テキストを画面にフィット（モーダル表示前に行う）
        dlog("    showFontPicker: before zoomToFrame");
        zoomToFrame(targetFrame);
        dlog("    showFontPicker: after zoomToFrame");

        // このピック用に UI をリセット（対象テキスト差し替え／検索クリア／一覧と選択の再設定）
        dlog("    showFontPicker: before reset");
        picker.targetLabel.text = lineText;
        picker.searchField.text = "";
        resetFamilyListToAll(picker.familyList, picker.families);
        initializeFontPickerSelection(picker.familyList, picker.styleList, picker.families, initialFont);
        dlog("    showFontPicker: after reset");

        // 「適用」=1 / 「スキップ」=2（閉じる含む）/ 「終了」=3
        dlog("    showFontPicker: before dialog.show()");
        var result = picker.dialog.show();
        dlog("    showFontPicker: after dialog.show() result=" + result);

        return handleFontPickerResult(result, targetFrame, paragraphIndex, picker.familyList, picker.styleList);
    }

    // 検索で絞り込まれている可能性のある familyList を全件に戻す。
    // 全件と同数なら何もしない（スキップ連打のような未検索ケースで再構築を避ける）。
    // ※ 既存コントロールへの item 追加は安全。落ちるのは「巨大配列での新規 dropdownlist 生成」だけ。
    function resetFamilyListToAll(familyList, families) {
        if (familyList.items.length === families.length) return;
        familyList.removeAll();
        for (var i = 0; i < families.length; i++) {
            familyList.add("item", families[i]);
        }
    }

    // frame と段落インデックスから段落（TextRange）を毎回ライブ取得する。
    // フォント適用後に保持参照を使い回すと無効化され、try/catch でも拾えない
    // ネイティブクラッシュ（Illustrator ごと落ちる）を起こすため、使用直前に都度取り直す。
    function getLiveParagraph(frame, index) {
        try {
            return frame.paragraphs[index];
        } catch (e) {
            return null;
        }
    }

    // フォントピッカーの結果を処理して返す。
    // 実適用はダイアログを閉じた後（＝モーダル表示外）に行うため安全。
    function handleFontPickerResult(result, targetFrame, paragraphIndex, familyList, styleList) {
        if (result === 1 && familyList.selection && styleList.selection) {
            var chosenFont = fontFor(familyList.selection.text, styleList.selection.text);
            if (chosenFont) {
                // 段落は保持参照を使い回さず、使用直前にライブ取得（無効化クラッシュ対策）
                var applyParagraph = getLiveParagraph(targetFrame, paragraphIndex);
                if (applyParagraph) {
                    try {
                        applyParagraph.characterAttributes.textFont = chosenFont;
                        return chosenFont;
                    } catch (e) { }
                }
            }
        }

        // スキップ／終了：ダイアログ中にドキュメントを変更していないので復元は不要。
        return (result === 3) ? PICKER_QUIT : null;
    }


    // フォントピッカーの UI を生成して { dialog, familyList, styleList } を返す。
    // ボタンは name:"ok"/"cancel" なので、判定は dialog.show() の戻り値で行う。
    function createFontPickerDialog(families) {
        var dialog = new Window("dialog", L(LABELS.picker.title) + " " + SCRIPT_VERSION);
        dialog.alignChildren = "fill";
        dialog.margins = 15;

        // 対象テキスト（パネルの外）。値はピックごとに差し替えるので空で作り、幅だけ確保する
        var targetRow = dialog.add("group");
        targetRow.add("statictext", undefined, L(LABELS.picker.target));
        var targetLabel = targetRow.add("statictext", undefined, "", { truncate: "end" });
        targetLabel.preferredSize.width = 260;

        // 置換パネル（フォント／スタイル）
        var replacePanel = dialog.add("panel", undefined, L(LABELS.picker.replace));
        replacePanel.orientation = "column";
        replacePanel.alignChildren = "left";
        replacePanel.margins = 15;

        // 各行のラベル幅をそろえる
        var labelWidth = 70;

        // 検索フィルター（入力した文字列でファミリーのプルダウンを絞り込む）
        var searchRow = replacePanel.add("group");
        var searchLabel = searchRow.add("statictext", undefined, L(LABELS.picker.search));
        searchLabel.preferredSize.width = labelWidth;
        var searchField = searchRow.add("edittext", undefined, "");
        searchField.preferredSize.width = 200;

        // フォント（ファミリー）プルダウン
        var familyRow = replacePanel.add("group");
        var familyLabel = familyRow.add("statictext", undefined, L(LABELS.picker.family));
        familyLabel.preferredSize.width = labelWidth;
        // 巨大配列での dropdownlist 生成を繰り返すと Illustrator が落ちるため、
        // このダイアログ（と familyList）は実行中に1回だけ生成して使い回す。
        dlog("      createFontPickerDialog: before family dropdownlist add");
        var familyList = familyRow.add("dropdownlist", undefined, families);
        dlog("      createFontPickerDialog: after family dropdownlist add");
        familyList.preferredSize.width = 200;

        // スタイルプルダウン
        var styleRow = replacePanel.add("group");
        var styleLabel = styleRow.add("statictext", undefined, L(LABELS.picker.style));
        styleLabel.preferredSize.width = labelWidth;
        var styleList = styleRow.add("dropdownlist", undefined, []);
        styleList.preferredSize.width = 200;

        // === ボタンエリア（左右分割：左=終了／右=スキップ・適用）===
        // メイングループ（横並び） / Main group (horizontal layout)
        var btnRowGroup = dialog.add("group");
        btnRowGroup.orientation = "row";
        btnRowGroup.margins = [0, 10, 0, 0];
        btnRowGroup.alignment = ["fill", "bottom"];

        // 左側グループ / Left-side button group
        var btnLeftGroup = btnRowGroup.add("group");
        btnLeftGroup.alignChildren = ["left", "center"];
        var btnQuit = btnLeftGroup.add("button", undefined, L(LABELS.picker.quit));
        btnQuit.onClick = function () { dialog.close(3); }; // 3 = 終了（選択ロジックを打ち切る）

        // スペーサー（伸縮）/ Spacer (stretchable)
        var spacer = btnRowGroup.add("group");
        spacer.alignment = ["fill", "fill"];
        spacer.minimumSize.width = 0;

        // 右側グループ / Right-side button group
        var btnRightGroup = btnRowGroup.add("group");
        btnRightGroup.alignChildren = ["right", "center"];
        btnRightGroup.add("button", undefined, L(LABELS.picker.skip), { name: "cancel" });
        btnRightGroup.add("button", undefined, L(LABELS.picker.apply), { name: "ok" });

        // イベントもダイアログ生成時に1回だけ接続する（使い回すため）
        bindFontPickerEvents(familyList, styleList, searchField, families);

        return {
            dialog: dialog,
            familyList: familyList,
            styleList: styleList,
            searchField: searchField,
            targetLabel: targetLabel,
            families: families
        };
    }

    // フォントピッカーの検索・ファミリー変更イベントを接続する。
    // （ドキュメントは変更しない。実適用は「適用」確定後にまとめて行う）
    function bindFontPickerEvents(familyList, styleList, searchField, families) {
        // 検索フィールドに入力するたびに、ファミリーのプルダウンを絞り込む
        searchField.onChanging = function () {
            filterFamilyList(familyList, families, searchField.text);
            updateFontPickerStyleSelection(familyList, styleList);
        };

        familyList.onChange = function () {
            updateFontPickerStyleSelection(familyList, styleList);
        };
    }

    // 選択中のファミリーに合わせてスタイル一覧を更新し、先頭スタイルを選択する
    function updateFontPickerStyleSelection(familyList, styleList) {
        populateFontPickerStyles(familyList, styleList);
        if (styleList.items.length > 0) styleList.selection = 0;
    }

    // フォントピッカーのスタイル一覧を、選択中のファミリーに合わせて更新する
    function populateFontPickerStyles(familyList, styleList) {
        styleList.removeAll();
        if (!familyList.selection) return;
        var styles = stylesForFamily(familyList.selection.text);
        for (var i = 0; i < styles.length; i++) {
            styleList.add("item", styles[i]);
        }
    }

    // フォントピッカーの初期ファミリー／スタイルを選択する
    function initializeFontPickerSelection(familyList, styleList, families, initialFont) {
        selectInList(familyList, initialFont ? initialFont.family : families[0]);
        if (!familyList.selection) familyList.selection = 0;

        populateFontPickerStyles(familyList, styleList);

        if (initialFont) selectInList(styleList, initialFont.style);
        if (!styleList.selection && styleList.items.length > 0) styleList.selection = 0;
    }

    // 検索クエリ（部分一致・大文字小文字を無視）でファミリーのドロップダウンを絞り込む。
    // 絞り込み後も、可能なら直前に選択していたファミリーを選び直す。
    // 一致が無くなった場合は空のまま（プレビュー側でガードしている）。
    function filterFamilyList(familyList, allFamilies, query) {
        var previousFamily = familyList.selection ? familyList.selection.text : null;
        var needle = String(query).toLowerCase();

        familyList.removeAll();
        for (var i = 0; i < allFamilies.length; i++) {
            if (needle === "" || allFamilies[i].toLowerCase().indexOf(needle) !== -1) {
                familyList.add("item", allFamilies[i]);
            }
        }

        if (previousFamily) selectInList(familyList, previousFamily);
        if (!familyList.selection && familyList.items.length > 0) familyList.selection = 0;
    }

    // dropdownlist で指定テキストの項目を選択する（無ければ何もしない）
    function selectInList(list, text) {
        for (var i = 0; i < list.items.length; i++) {
            if (list.items[i].text === text) {
                list.selection = i;
                return;
            }
        }
    }

    // 対象フレームが画面にフィットするようズーム＋センタリングする
    function zoomToFrame(frame) {
        try {
            var view = doc.activeView;
            var bounds = frame.geometricBounds; // [left, top, right, bottom]
            var frameWidth = bounds[2] - bounds[0];
            var frameHeight = bounds[1] - bounds[3];
            if (frameWidth <= 0 || frameHeight <= 0) return;

            var visible = view.bounds; // 現在の表示範囲（ドキュメント座標）
            var visibleWidth = visible[2] - visible[0];
            var visibleHeight = visible[1] - visible[3];

            // 画面の ZOOM_FIT_RATIO に収まる倍率を現在ズームから算出
            var fitZoom = Math.min(visibleWidth / frameWidth, visibleHeight / frameHeight) * view.zoom * ZOOM_FIT_RATIO;
            if (fitZoom > 0) view.zoom = fitZoom;

            view.centerPoint = [(bounds[0] + bounds[2]) / 2, (bounds[1] + bounds[3]) / 2];
        } catch (e) { }
    }

    // インストール済みフォントを索引化する
    function createFontIndex() {
        var fonts = app.textFonts;
        var index = {
            fonts: fonts,
            families: [],
            stylesByFamily: {},
            fontByFamilyStyle: {},
            normalizedFonts: []
        };

        for (var i = 0; i < fonts.length; i++) {
            var family = fonts[i].family;
            var style = fonts[i].style;

            if (!index.stylesByFamily.hasOwnProperty(family)) {
                index.stylesByFamily[family] = [];
                index.families.push(family);
            }

            index.stylesByFamily[family].push(style);
            index.fontByFamilyStyle[makeFamilyStyleKey(family, style)] = fonts[i];
            index.normalizedFonts.push({
                font: fonts[i],
                family: normalize(family),
                full: normalize(family + " " + style),
                name: normalize(fonts[i].name)
            });
        }

        index.families.sort();

        for (var familyName in index.stylesByFamily) {
            if (index.stylesByFamily.hasOwnProperty(familyName)) {
                index.stylesByFamily[familyName].sort();
            }
        }

        return index;
    }

    // ファミリー名＋スタイル名の索引用キーを作る
    function makeFamilyStyleKey(family, style) {
        return family + "\u0000" + style;
    }

    // 指定ファミリーに属するスタイル名の一覧
    function stylesForFamily(family) {
        return fontIndex.stylesByFamily[family] || [];
    }

    // ファミリー名＋スタイル名から TextFont を取得（無ければ null）
    function fontFor(family, style) {
        return fontIndex.fontByFamilyStyle[makeFamilyStyleKey(family, style)] || null;
    }

    // 選択物（配列）を再帰的にたどり、TextFrame を collected に集める
    function collectTextFrames(items, collected) {
        for (var k = 0; k < items.length; k++) {
            var item = items[k];

            // ロック・非表示のオブジェクトは無視（グループならその中身ごとスキップ）
            if (item.locked || item.hidden) continue;

            if (item.typename === "TextFrame") {
                collected.push(item);
            } else if (item.typename === "GroupItem") {
                // グループの中身（pageItems）をさらにたどる
                collectTextFrames(item.pageItems, collected);
            }
        }
    }

    // 1つのテキストフレームの各行（段落）について、厳密一致のフォントは自動適用し、
    // あいまい／未一致の行は { frame, index, lineText, initialFont } を pendingPicks に積む。
    // （対話ピッカーはプログレスバー終了後にまとめて出すため、ここでは出さない）
    function autoApplyAndQueue(textFrame, pendingPicks) {
        // 段落数だけ先に取得（コレクション自体はキャッシュしない）。
        // フォント適用でテキストが変更されると paragraphs コレクションの
        // 参照が無効化され、Error 1302 になるため毎回ライブ取得する。
        var paragraphCount = textFrame.paragraphs.length;

        // 下から上にループすると、フォント適用によるズレを防げて安全です
        for (var j = paragraphCount - 1; j >= 0; j--) {
            var paragraph = textFrame.paragraphs[j]; // 毎回ライブ取得

            // 改行コードや前後の空白（半角・全角スペース含む）を除外して、純粋なテキストを取得
            var lineText = paragraph.contents.replace(/^[\s　]+|[\s　]+$/g, "");

            // 空行・空白だけの行は無視
            if (lineText === "") continue;

            // フォントを検索（{ font, confident }）
            var match = findFont(lineText);

            var autoApplied = false;
            if (match.confident && match.font) {
                // 厳密一致 → 自動適用
                try {
                    paragraph.characterAttributes.textFont = match.font;
                    autoApplied = true;
                } catch (e) {
                    autoApplied = false;
                }
            }

            // あいまい一致／未一致／適用失敗 → 後で対話ピッカーに回す
            if (!autoApplied) {
                pendingPicks.push({
                    frame: textFrame,
                    index: j,
                    lineText: lineText,
                    initialFont: match.font
                });
            }
        }
    }

    // 配列に未登録の値だけ追加する（重複防止）
    function pushUnique(list, value) {
        for (var i = 0; i < list.length; i++) {
            if (list[i] === value) return;
        }
        list.push(value);
    }

    // 未適用の行（スキップ／適用失敗）をフレーム・文字列リストへ記録する
    function recordUnapplied(frames, texts, pick) {
        pushUnique(texts, pick.lineText);
        pushUnique(frames, pick.frame); // 同一フレーム参照なので === で重複排除できる
    }

    // 指定名のレイヤーがあれば削除する（無ければ何もしない）
    function removeLayerByName(name) {
        var layer = getLayerByName(name);
        if (!layer) return;

        unlockLayerTree(layer);

        try {
            layer.remove();
        } catch (e) { }
    }

    // 指定名のレイヤーを取得する（無ければ null）
    function getLayerByName(name) {
        try {
            return doc.layers.getByName(name);
        } catch (e) {
            return null;
        }
    }

    // レイヤーと配下アイテムを、可能な範囲でロック解除・表示する
    function unlockLayerTree(layer) {
        unlockContainer(layer);
        unlockPageItems(layer.pageItems);
    }

    // レイヤー／アイテム共通のロック解除・表示処理
    function unlockContainer(item) {
        try {
            item.locked = false;
        } catch (e) { }

        try {
            item.visible = true;
        } catch (e2) { }

        try {
            item.hidden = false;
        } catch (e3) { }
    }

    // PageItems コレクション内を再帰的にロック解除・表示する
    function unlockPageItems(items) {
        for (var i = 0; i < items.length; i++) {
            unlockContainer(items[i]);

            if (items[i].typename === "GroupItem") {
                unlockPageItems(items[i].pageItems);
            }
        }
    }

    // 目印レイヤーを新規作成し、テキストの背面に来るよう最背面へ送る
    function createMarkerLayer(name) {
        var layer = doc.layers.add();
        layer.name = name;

        var lastLayer = doc.layers[doc.layers.length - 1];
        if (lastLayer.name !== layer.name) {
            try {
                layer.move(lastLayer, ElementPlacement.PLACEAFTER);
            } catch (e) { }
        }
        return layer;
    }

    // テキストフレームの背面（＝目印レイヤー上）に、赤・半透明の長方形を作る
    function createMarkerRect(frame, layer) {
        try {
            // geometricBounds: [left, top, right, bottom]（線幅は含まない）
            var bounds = frame.geometricBounds;
            var left = bounds[0], top = bounds[1], right = bounds[2], bottom = bounds[3];
            var width = right - left;
            var height = top - bottom;
            if (width <= 0 || height <= 0) return;

            var markerRect = layer.pathItems.rectangle(top, left, width, height);
            markerRect.stroked = false;
            markerRect.filled = true;
            markerRect.fillColor = makeRedColor();
            markerRect.opacity = MARKER_OPACITY;
        } catch (e) { }
    }

    // ドキュメントのカラースペースに合わせた赤色を返す
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

    // 正規化済みフォント索引から条件に合うフォントを集めて配列で返す
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

    // 文字列に対応するフォントを段階的に（厳密→ゆるい）探す。
    // 戻り値は { font, confident }:
    //   confident=true … 厳密一致。自動適用してよい
    //   confident=false 且つ font!=null … あいまい一致。ピッカーの初期値に使う
    //   font=null … 未一致。ピッカーで一から選ばせる
    function findFont(fontName) {
        // 1. カスタム置換マップ（hasOwnProperty で prototype のメソッド名への誤ヒットを防ぐ）
        if (CUSTOM_MAP.hasOwnProperty(fontName)) {
            fontName = CUSTOM_MAP[fontName];
        }

        // 2. PostScript名での完全一致（厳密）
        try {
            return { font: fontIndex.fonts.getByName(fontName), confident: true };
        } catch (e) { }

        // 以降はあいまい照合。スペース・ピリオドを除去した文字列で比較する
        var query = normalize(fontName);

        // 3. ファミリー名＋スタイル名の完全一致（厳密）
        var exactFull = collectIndexedFonts(function (fontInfo) {
            return fontInfo.full === query;
        });
        if (exactFull.length > 0) return { font: getBestStyle(exactFull), confident: true };

        // 4. ファミリー名のみの完全一致（厳密。スタイルは Bold 優先で確定）
        var exactFamily = collectIndexedFonts(function (fontInfo) {
            return fontInfo.family === query;
        });
        if (exactFamily.length > 0) return { font: getBestStyle(exactFamily), confident: true };

        // 5. ファミリー名への部分一致（あいまい。例: "Jenson" → "Adobe Jenson Pro"）
        var partialFamily = collectIndexedFonts(function (fontInfo) {
            return fontInfo.family.indexOf(query) !== -1;
        });
        if (partialFamily.length > 0) return { font: getBestStyle(partialFamily), confident: false };

        // 6. フォント名全体への部分一致（あいまい）
        var partialFull = collectIndexedFonts(function (fontInfo) {
            return fontInfo.full.indexOf(query) !== -1
                || fontInfo.name.indexOf(query) !== -1;
        });
        if (partialFull.length > 0) return { font: getBestStyle(partialFull), confident: false };

        // 7. 最終手段（ゆるめ）：先頭ワードがファミリー名に含まれればOK（あいまい）。
        // 例: "Myriad Pro Cond" が未インストールでも "myriad" で "Myriad Pro" にフォールバック。
        // 短すぎる語（2文字以下）は誤マッチ防止のため対象外。
        var firstWord = String(fontName).toLowerCase()
            .replace(/[.　]+/g, " ").replace(/^\s+/, "").split(/\s+/)[0] || "";
        if (firstWord.length >= 3) {
            var looseFamily = collectIndexedFonts(function (fontInfo) {
                return fontInfo.family.indexOf(firstWord) !== -1;
            });
            if (looseFamily.length > 0) return { font: getBestStyle(looseFamily), confident: false };
        }

        return { font: null, confident: false };
    }

    // フォント名照合用の正規化：小文字化し、空白（半角・全角）とピリオドを
    // すべて除去して連結する。これにより「Bank Gothic」⇄「BankGothic」、
    // 「Mrs. Eaves」⇄「Mrs Eaves OT」のようなスペース・記号ゆれを吸収する。
    function normalize(text) {
        return String(text)
            .toLowerCase()
            .replace(/[.\s　]+/g, "");
    }

    // マッチしたフォント群から、最適なスタイルを選択して返す。
    // 優先順位は STYLE_PRIORITY の並び順に従う（いずれにも該当しなければ候補の先頭）。
    function getBestStyle(candidates) {
        var preferredStyles = STYLE_PRIORITY;
        for (var k = 0; k < preferredStyles.length; k++) {
            for (var j = 0; j < candidates.length; j++) {
                if (candidates[j].style.toLowerCase() === preferredStyles[k]) {
                    return candidates[j];
                }
            }
        }
        return candidates[0];
    }
})();