#target illustrator

    /*
    
    ### スクリプト名：
    
    FileNameManager.jsx — 現在のドキュメントのファイル名を変更／別名で保存

    ### 概要：

    - アクティブな Illustrator ドキュメントのファイル名を変更（実ファイルのリネーム）／別名で保存／コピーを保存します。
    - 保存先は常に現在のファイルと同じフォルダ。未保存ドキュメントは保存先フォルダを選択して保存します。
    - ファイル名を base / title / status / timestamp / version の各セグメントに分解し、UI から個別に編集できます。
      - base: 先頭の固定部分（UI には出さずロジック側で保持）
      - title: 「タイトル」ラジオで「なし / 親フォルダー / 指定（入力欄）」を選択。入力欄はデフォルト空白
      - status: 「ステータス」dropdown で制作ステータス（wip / draft / review / approved / outlined など）を挿入。`：` より前の値だけがファイル名に入る
      - timestamp: 「タイムスタンプ」ラジオで「なし / YYYYMMDD / YYYY-MM-DD」を選択（デフォルト YYYYMMDD）
      - version: 「バージョン番号」ラジオで「なし / -vN / -v0N」を選択（デフォルト -vN）。既存の v 番号は +1、無い場合は v2 / v02 を付与
    - 整形ルール：
      - 区切り記号: 「変更しない / `-` / `_`」。`-` または `_` を選ぶとピース内の `-` `_` `.`（FEATURE_DOT_NORMALIZE）を統一。`YYYY-MM-DD` のタイムスタンプは対象外（内部 `-` を保護）
      - スペースの扱い: 「変更しない / -に変更 / _に変更」で半角スペース（タブ・改行含む `\s+`）を置換
    - 構成要素順のカスタマイズ：
      - 上部「構成要素の順序」パネルで「カスタム順」を選び、［順序を編集...］ボタンから ↑↓ で並び替え
      - 設定はプリセットに保存され次回起動時に復元
    - 動作モード: 「元ファイルをリネーム」（新名で保存後に元ファイルを削除）／「別名で保存」（元ファイルは残り、作業中のドキュメントが新ファイルに切り替わる。デフォルト）／「コピーを保存」（元ファイルを上書き保存したうえで別名のコピーを作成。作業中のドキュメントは元ファイルのまま）
    - UI 構成: 上部「動作 + 構成要素の順序」の 2 カラム → 「ファイル名プレビュー」（現在 / 保存後の名前）→ 「ファイル名の設定」 → Cancel / OK（右寄せ）
    - トップ部の FEATURE_STATUS / FEATURE_SORT / FEATURE_SEPARATOR / FEATURE_SPACES / FEATURE_DOT_NORMALIZE で各機能を個別オフ可能。FEATURE_SEPARATOR は `'-'` / `'_'` の文字列で既定の区切りも指定

    ### 主な機能：

    - 元ファイルをリネーム（新名で保存後に元ファイルを削除して実質的にリネーム）
    - 別名で保存（元ファイルは保持。アクティブドキュメントは新ファイルに切り替わる。デフォルト）
    - コピーを保存（元ファイルに上書き保存後、別名のコピーを物理ファイルとして作成）
    - セグメント分解による独立編集（base / title / status / timestamp / version。status / date / version は自動検出）
    - ステータス dropdown（11 種類 + 区切り線、ファイル名に入るのは `：` より前のみ。既存ファイル名から自動検出）
    - 構成要素順のカスタマイズ（サブダイアログで ↑↓、プリセット保存）
    - タイムスタンプ／バージョン番号は「なし」で削除、形式切替（YYYYMMDD / YYYY-MM-DD / -vN / -v0N）
    - 区切り記号統一（`-` `_` `.`）／空白置換（半角スペース・タブ・改行を `-` or `_` に置換）／YYYY-MM-DD の内部 `-` 保護
    - 各機能を FEATURE_* 定数で個別にオフ可能

    ### 処理の流れ：

    1) ドキュメント検証 → gatherDocumentInfo() でファイル情報を取得
    2) parseFileName() でセグメント配列に分解（区切り文字も保持、status も検出）
    3) createDialog() でダイアログを構築（buildModePanel + buildSortPanel + buildFilenamePanel + buildOptionsPanel）
    4) buildFinalName() で UI 状態と segments、現在の SEGMENT_ORDER（カスタム順含む）から最終ファイル名を構築
    5) ensureTargetFolder() で未保存ドキュメントは保存先フォルダを取得
    6) confirmOverwriteIfExists() で既存同名の上書きを確認
    7) executeOutput() で選択モード（rename / saveAs / copy）に応じた出力を実行

    ### 更新履歴：

    - v1.0 (2026-05-27) : 初期バージョン
    - v1.1 (2026-06-01) : コロン付きラベル関数追加
    - (2026-05-27 追記) : ステータス dropdown、構成要素順カスタマイズ、スペース置換、YYYY-MM-DD タイムスタンプ、`.` 区切り正規化、整形機能の FEATURE スイッチ化、UI 整理（ベース行非表示、動作・構成要素 2 カラム、ボタン右寄せ）
    
    ---
    
    ### Script name:
    
    FileNameManager.jsx — Rename or Save As the Current Document
    
    ### Overview:

    - Renames the active Illustrator document (true rename), saves as a new file, or saves a copy.
    - Destination is always the same folder as the current file; for unsaved docs, a destination folder is chosen.
    - Decomposes the filename into base / title / status / timestamp / version segments and lets you edit them.
      - base: the leading fixed part; not shown in the UI but preserved by the logic
      - title: "Title" radio (None / Parent Folder / Custom input); input field is empty by default
      - status: "Status" dropdown (wip / draft / review / approved / outlined, etc.); only the text before `:` is written to the filename
      - timestamp: "Timestamp" radio selects "None / YYYYMMDD / YYYY-MM-DD" (default YYYYMMDD)
      - version: "Version" radio selects "None / -vN / -v0N" (default -vN). An existing v-number is bumped by +1; otherwise v2 / v02 is added
    - Formatting rules:
      - Separator: "Don't change / `-` / `_`". Selecting `-` or `_` unifies `-`, `_`, and `.` (FEATURE_DOT_NORMALIZE) inside each piece. `YYYY-MM-DD` timestamps are exempt (internal `-` is preserved)
      - Spaces: "Don't change / Change to - / Change to _" replaces ASCII whitespace (incl. tabs/newlines via `\s+`)
    - Custom segment order:
      - In the top "Segment Order" panel, choose "Custom" and click [Edit order...] to reorder with ↑↓
      - The order is persisted to prefs and restored on next launch
    - Modes: "Rename Original" (saves with the new name, then deletes the original), "Save As" (keeps the original; the active document switches to the new file; default), and "Save a Copy" (saves over the original, then creates a separate copy; the active document remains the original file).
    - UI layout: top row "Mode + Segment Order" (2 columns) → "File Name Preview" (Current / Saved Name) → "Filename Settings" → Cancel / OK (right-aligned)
    - Features can be toggled individually via FEATURE_STATUS / FEATURE_SORT / FEATURE_SEPARATOR / FEATURE_SPACES / FEATURE_DOT_NORMALIZE. FEATURE_SEPARATOR accepts `'-'` or `'_'` to also pick the default separator.

    ### Key features:

    - Rename Original (saves with the new name, then removes the original)
    - Save As (keeps the original; the active document switches to the new file; default)
    - Save a Copy (saves over the original, then creates a separate copy; the active document remains the original file)
    - Segment-based editing (base / title / status / timestamp / version; status / date / version auto-detected)
    - Status dropdown (11 entries + divider; only the text before `:` is used; detected from existing filenames)
    - Custom segment order via sub-dialog (↑↓; persisted to prefs)
    - Timestamp / version: "None" removes the segment; switch format (YYYYMMDD / YYYY-MM-DD / -vN / -v0N)
    - Separator unification (`-` `_` `.`) / space replacement / `-` protection inside YYYY-MM-DD
    - Each feature toggleable via FEATURE_* constants

    ### Flow:

    1) gatherDocumentInfo() collects file info
    2) parseFileName() splits the name into ordered segments (separators preserved; status detected)
    3) createDialog() composes the dialog (buildModePanel + buildSortPanel + buildFilenamePanel + buildOptionsPanel)
    4) buildFinalName() composes the final filename from UI state, segments, and the current SEGMENT_ORDER (custom order included)
    5) ensureTargetFolder() prompts for a destination folder if the doc is unsaved
    6) confirmOverwriteIfExists() confirms before overwriting an existing file
    7) executeOutput() runs the action for the selected mode (rename / saveAs / copy)

    ### Changelog:

    - v1.0 (2026-05-27): Initial release
    - (2026-05-27 update): Added status dropdown, custom segment order, space replacement, YYYY-MM-DD timestamp, `.` separator normalization, FEATURE switches, UI cleanup (no base row, 2-col top, right-aligned buttons)
    
    */

    (function () {

        // =========================================
        // バージョン / Version
        // =========================================

        var SCRIPT_VERSION = "v1.1.0";

        // =========================================
        // ユーザー設定 / User Settings
        // =========================================

        /* 整形機能のオン/オフ。false にすると該当 UI とロジックを丸ごとスキップ
           / Feature switches. Set to false to omit the UI and skip the logic */
        var FEATURE_STATUS = true;        // ステータス dropdown と検出 / Status dropdown + detection
        var FEATURE_SORT = true;          // ソートパネル + 並び順カスタマイズ / Sort panel + custom segment order
        var FEATURE_SEPARATOR = '-';      // 区切り記号統一: '-' / '_' で有効化＋既定値、false で無効 / '-' or '_' enables with that default; false disables
        var FEATURE_SPACES = '-';         // スペース置換: '-' / '_' で有効化＋既定値、false で無効 / '-' or '_' enables with that default; false disables
        var FEATURE_DOT_NORMALIZE = true; // "." を区切り記号にあわせて置換（要 FEATURE_SEPARATOR） / "." normalization with the chosen separator

        /* 出力時のセグメント順序。base / title / status / timestamp / version。FEATURE_STATUS=false なら status は除外
           / Output segment order; "status" is dropped when FEATURE_STATUS is false */
        var SEGMENT_ORDER = FEATURE_STATUS
            ? ['base', 'title', 'status', 'timestamp', 'version']
            : ['base', 'title', 'timestamp', 'version'];

        /* ステータス選択肢。value がファイル名に入り、ja / en が UI 表示。先頭は「なし」。
           ja が '---' の項目は dropdown 上の区切り線として表示（選択不可）
           / Status choices: `value` is what enters the filename; `ja` / `en` are UI labels.
           Items with `ja === '---'` render as an unselectable divider. */
        var STATUS_ITEMS = [
            { value: '', ja: 'なし', en: 'None' },
            { value: 'wip', ja: 'wip：作業中・仕掛かり中', en: 'wip: Work in progress' },
            { value: 'draft', ja: 'draft：下書き・ラフ・素案', en: 'draft: Draft / Rough' },
            { value: 'review', ja: 'review：レビュー待ち', en: 'review: Awaiting review' },
            { value: 'revised', ja: 'revised：改訂版・修正反映版', en: 'revised: Revised' },
            { value: 'updated', ja: 'updated：更新版', en: 'updated: Updated' },
            { value: 'fixed', ja: 'fixed：修正完了', en: 'fixed: Fix completed' },
            { value: 'approved', ja: 'approved：承認済み・確定', en: 'approved: Approved / Final' },
            { value: 'rejected', ja: 'rejected：ボツ・不採用案', en: 'rejected: Rejected' },
            { value: 'archived', ja: 'archived：保管用', en: 'archived: Archive' },
            { value: '', ja: '---', en: '---' },
            { value: 'flattened', ja: 'flattened：レイヤー結合済み（.psd）', en: 'flattened: Flattened (.psd)' },
            { value: 'outlined', ja: 'outlined：アウトライン済み（.ai）', en: 'outlined: Outlined (.ai)' }
        ];

        /* 区切り線エントリかどうか / Whether a STATUS_ITEMS entry is a divider */
        function isStatusDivider(item) {
            return item && item.ja === '---';
        }

        var NEW_NAME_FIELD_WIDTH = 250;
        var PANEL_MARGINS = [15, 20, 15, 10];
        var PANEL_SPACING = 8;

        var DIALOG_OPACITY = 0.98;

        // =========================================
        // ローカライズ / Localization
        // =========================================

        var lang = ($.locale.indexOf("ja") === 0) ? "ja" : "en";

        /* 日英ラベル定義 / Japanese-English label definitions */

        var LABELS = {
            dialog: {
                title: { ja: "ファイル名を変更して保存", en: "Rename and Save" }
            },
            panel: {
                mode: { ja: "動作", en: "Mode" },
                filename: { ja: "ファイル名プレビュー", en: "File Name Preview" },
                options: { ja: "ファイル名の設定", en: "Filename Settings" },
                sort: { ja: "構成要素の順序", en: "Segment Order" }
            },
            radio: {
                rename: { ja: "元ファイルをリネーム", en: "Rename Original" },
                saveAs: { ja: "別名で保存", en: "Save As" },
                saveCopy: { ja: "コピーを保存", en: "Save a Copy" },
                noChange: { ja: "変更しない", en: "No Change" },
                titleNone: { ja: "なし", en: "None" },
                titleParent: { ja: "親フォルダー", en: "Parent Folder" },
                titleCustom: { ja: "指定", en: "Custom" },
                timestampNone: { ja: "なし", en: "None" },
                timestampDate: { ja: "YYYYMMDD", en: "YYYYMMDD" },
                timestampDateDash: { ja: "YYYY-MM-DD", en: "YYYY-MM-DD" },
                versionNone: { ja: "なし", en: "None" },
                versionShort: { ja: "-vN", en: "-vN" },
                versionPadded: { ja: "-v0N", en: "-v0N" },
                changeToDash: { ja: "-に変更", en: "Change to -" },
                changeToUnderscore: { ja: "_に変更", en: "Change to _" },
                sortOff: { ja: "標準順", en: "Default" },
                sortOn: { ja: "カスタム順", en: "Custom" }
            },
            label: {
                currentName: { ja: "現在", en: "Current" },
                finalName: { ja: "保存後の名前", en: "Saved Name" },
                base: { ja: "ベース", en: "Base" },
                title: { ja: "タイトル", en: "Title" },
                status: { ja: "ステータス", en: "Status" },
                timestamp: { ja: "タイムスタンプ", en: "Timestamp" },
                version: { ja: "バージョン番号", en: "Version" },
                separator: { ja: "区切り記号", en: "Separator" },
                spaces: { ja: "スペースの扱い", en: "Spaces" }
            },
            tip: {
                rename: {
                    ja: "新しい名前で保存したあと、元ファイルを削除します（実質的にリネーム）。",
                    en: "Saves with the new name, then deletes the original file (effectively a rename)."
                },
                saveAs: {
                    ja: "新しい名前で保存します。元ファイルは残り、作業中のドキュメントが新ファイルに切り替わります。",
                    en: "Saves with the new name. The original file is kept, and the active document switches to the new file."
                },
                saveCopy: {
                    ja: "元ファイルに上書き保存したうえで、別名のコピーを作成します。作業中のドキュメントは元ファイルのまま残ります。",
                    en: "Saves the original, then creates a copy with the new name. The active document remains the original file."
                },
                title: {
                    ja: "ファイル名のタイトル部分の扱いを選択します。「指定」で入力欄の文字列を使用します。",
                    en: "Choose how to set the title part of the filename. With \"Custom\", the entered text is used."
                },
                status: {
                    ja: "ファイル名に挿入する制作ステータスを選択します。「：」より前の文字列が入ります。",
                    en: "Choose a production status to insert. Only the text before \":\" is used in the filename."
                },
                timestamp: {
                    ja: "タイムスタンプの形式を選択。「なし」で元の日付があっても削除します。",
                    en: "Choose timestamp format. \"None\" removes any existing date."
                },
                version: {
                    ja: "バージョン番号の形式。-vN はパディング無し、-v0N はゼロ埋め。既存の v 番号は +1、無い場合は v2 / v02 を付与。「なし」で削除。",
                    en: "Version format. -vN has no padding, -v0N is zero-padded. An existing v-number is bumped by +1; otherwise v2 / v02 is added. \"None\" removes."
                },
                separator: {
                    ja: "ファイル名全体の区切り記号の扱いを選択します。「-」「_」「.」が対象。YYYY-MM-DD のタイムスタンプは保護されます。",
                    en: "Choose how separators in the filename are handled. Targets `-`, `_`, and `.`. YYYY-MM-DD timestamps are preserved."
                },
                spaces: {
                    ja: "最終ファイル名に含まれるスペースの扱いを選択します。",
                    en: "Choose how spaces in the final filename are handled."
                },
                sort: {
                    ja: "「カスタム順」を選び［順序を編集］から並び順を編集します。",
                    en: "Choose \"Custom\" and click [Edit order] to edit the order."
                }
            },
            button: {
                cancel: { ja: "キャンセル", en: "Cancel" },
                sort: { ja: "順序を編集...", en: "Edit order..." }
            },
            message: {
                noDoc: { ja: "ドキュメントが開かれていません", en: "No document is open." },
                emptyName: { ja: "ファイル名が空です", en: "File name is empty." },
                chooseDestination: { ja: "保存先フォルダを指定", en: "Choose destination folder" },
                confirmOverwrite: {
                    ja: "同名ファイルが存在します。上書きしますか？",
                    en: "A file with the same name exists. Overwrite?"
                },
                saveFailed: { ja: "保存に失敗しました", en: "Failed to save" }
            }
        };

        /* ドット記法でローカライズ済み文字列を取得 / Get the localized string by dotted key (e.g. L('dialog.title')) */
        function L(path) {
            var parts = String(path).split('.');
            var entry = LABELS;
            for (var i = 0; i < parts.length; i++) {
                entry = entry && entry[parts[i]];
                if (!entry) return path;
            }
            return entry[lang] || entry.en || path;
        }

        /* コロン付きラベル（日本語は全角、英語は半角）/ Label with colon (full-width JA, half-width EN) */
        function labelText(path) {
            return L(path) + (lang === 'ja' ? '：' : ':');
        }

        // =========================================
        // ヘルパー / Helpers
        // =========================================

        var VERSION_TOKEN_RE = /^[vV]\d+$/;   // v123 / V123

        /* 月（1-12）と日（1-31）の妥当性 / Validate month (1-12) and day (1-31) */
        function isValidMonthDay(monthStr, dayStr) {
            var m = parseInt(monthStr, 10);
            var d = parseInt(dayStr, 10);
            return m >= 1 && m <= 12 && d >= 1 && d <= 31;
        }

        /* YYYYMMDD（8 桁）または YYMMDD（6 桁）の日付トークンか / Whether a token is a YYYYMMDD or YYMMDD date */
        function isDateToken(token) {
            var s = String(token);
            if (/^\d{8}$/.test(s)) return isValidMonthDay(s.substring(4, 6), s.substring(6, 8));
            if (/^\d{6}$/.test(s)) return isValidMonthDay(s.substring(2, 4), s.substring(4, 6));
            return false;
        }

        /* "v123" 形式のバージョントークンか / Whether a token is a v+digits version */
        function isVersionToken(token) {
            return VERSION_TOKEN_RE.test(String(token));
        }

        /* STATUS_ITEMS のいずれかの value と完全一致（大文字小文字無視）すれば true。一致した正規 value を返す。divider は除外
           / Match against STATUS_ITEMS values (case-insensitive); returns the canonical value or '' (skips dividers) */
        function matchStatusToken(token) {
            var s = String(token).toLowerCase();
            for (var i = 1; i < STATUS_ITEMS.length; i++) {
                if (isStatusDivider(STATUS_ITEMS[i])) continue;
                if (STATUS_ITEMS[i].value === s) return STATUS_ITEMS[i].value;
            }
            return '';
        }

        /* ファイル名を順序付きセグメント配列に分解。各 segment は前にあった区切り文字も保持
           / Decompose into ordered segments; each segment stores the preceding separator */
        function parseFileName(name) {
            var split = String(name).split(/([-_])/); // ["handout","-","Adobe","-","20260422"]
            var segments = [];
            var textBuffer = [];      // token と区切りを交互に蓄積（join('') で結合）
            var textLeadingSep = '';  // テキスト segment の直前に置く区切り
            var hasDate = false, hasVersion = false, hasStatus = false;
            var currentSep = '';

            function flushText() {
                if (!textBuffer.length) return;
                segments.push({
                    kind: 'text',
                    value: textBuffer.join(''),
                    sep: textLeadingSep
                });
                textBuffer = [];
                textLeadingSep = '';
            }

            var i = 0;
            while (i < split.length) {
                if (i % 2 === 1) {
                    currentSep = split[i];
                    i++;
                    continue;
                }
                var token = split[i];

                // YYYY-MM-DD / YYYY_MM_DD（3 トークン + 同じ区切り 2 つ）の日付パターン
                if (!hasDate && i > 0
                    && /^\d{4}$/.test(token)
                    && i + 4 < split.length
                    && (split[i + 1] === '-' || split[i + 1] === '_')
                    && /^\d{2}$/.test(split[i + 2])
                    && split[i + 3] === split[i + 1]
                    && /^\d{2}$/.test(split[i + 4])
                    && isValidMonthDay(split[i + 2], split[i + 4])) {
                    flushText();
                    var compositeDate = token + split[i + 1] + split[i + 2] + split[i + 3] + split[i + 4];
                    segments.push({ kind: 'date', value: compositeDate, sep: currentSep });
                    hasDate = true;
                    i += 5; // 3 tokens + 2 separators
                    continue;
                }

                if (i === 0) {
                    segments.push({ kind: 'base', value: token, sep: '' });
                } else if (!hasDate && isDateToken(token)) {
                    flushText();
                    segments.push({ kind: 'date', value: token, sep: currentSep });
                    hasDate = true;
                } else if (!hasVersion && isVersionToken(token)) {
                    flushText();
                    segments.push({ kind: 'version', value: token, sep: currentSep });
                    hasVersion = true;
                } else if (FEATURE_STATUS && !hasStatus && matchStatusToken(token)) {
                    flushText();
                    segments.push({ kind: 'status', value: matchStatusToken(token), sep: currentSep });
                    hasStatus = true;
                } else {
                    if (!textBuffer.length) {
                        textLeadingSep = currentSep;
                    } else {
                        textBuffer.push(currentSep);
                    }
                    textBuffer.push(token);
                }
                i++;
            }
            flushText();
            return segments;
        }

        /* segments から最初に出現する kind の value を取得 / Get the value of the first segment of the given kind */
        function getFirstSegmentValue(segments, kind) {
            for (var i = 0; i < segments.length; i++) {
                if (segments[i].kind === kind) return segments[i].value;
            }
            return '';
        }

        /* 文字列を左 0 パディング / Left-pad a string with zeros to the given width */
        function padLeft(str, width) {
            while (str.length < width) str = '0' + str;
            return str;
        }

        /* 今日の日付を返す。sep を渡すと YYYY{sep}MM{sep}DD、未指定で YYYYMMDD
           / Today's date; with `sep` returns YYYY{sep}MM{sep}DD, otherwise YYYYMMDD */
        function todayTimestamp(sep) {
            sep = sep || '';
            var d = new Date();
            return String(d.getFullYear()) + sep +
                padLeft(String(d.getMonth() + 1), 2) + sep +
                padLeft(String(d.getDate()), 2);
        }

        /* バージョン文字列を +1。mode='padded' でゼロ埋め（最低 2 桁）。元バージョンが無ければ新規付与（v2 / v02）
           / Bump the version string by +1. 'padded' zero-pads to at least 2 digits. Returns v2 / v02 if no original */
        function formatVersion(originalVersion, mode) {
            var match = String(originalVersion || '').match(/^([vV])(\d+)$/);
            var letter = match ? match[1] : 'v';
            var nextNum = match ? (parseInt(match[2], 10) + 1) : 2;
            if (mode === 'padded') {
                var width = match ? Math.max(match[2].length, 2) : 2;
                return letter + padLeft(String(nextNum), width);
            }
            return letter + String(nextNum);
        }

        /* % エンコードを 1 回デコード / Decode a percent-encoded string once (best-effort) */
        function decodePercentEncoded(str) {
            str = String(str);
            if (str.indexOf('%') === -1) return str;
            try {
                return decodeURIComponent(str);
            } catch (e) {
                return str;
            }
        }

        /* 拡張子を除去 / Strip a trailing file extension */
        function stripExtension(name) {
            var dot = name.lastIndexOf('.');
            return (dot > 0) ? name.substring(0, dot) : name;
        }

        /* ファイル名として無効な文字を除去し前後空白をトリム / Strip invalid filename chars and trim whitespace */
        function sanitizeFilename(str) {
            return String(str).replace(/[\\\/:*?"<>|]/g, '').replace(/^\s+|\s+$/g, '');
        }

        /* 保存先ダイアログ。.ai 拡張子を補正して File を返す。キャンセルは null / Show save dialog, normalize to .ai; null on cancel */
        function pickAiDestination(promptLabel) {
            var chosen = File.saveDialog(promptLabel, '*.ai');
            if (!chosen) return null;
            var file = File(chosen.fsName);
            if (!/\.ai$/i.test(file.name)) {
                file = File(file.parent.fsName + '/' + stripExtension(file.name) + '.ai');
            }
            return file;
        }

        /* プリセット保存用ファイル / Preferences file (key=value lines) */
        function getPrefsFile() {
            return File(Folder.userData.fsName + '/FileNameManager-prefs.txt');
        }

        /* 前回保存したプリセットを読み込み（無ければ空オブジェクト） / Load previously saved prefs (or empty if none) */
        function loadPrefs() {
            var file = getPrefsFile();
            if (!file.exists) return {};
            file.encoding = 'UTF-8';
            file.open('r');
            var raw = file.read();
            file.close();
            var prefs = {};
            var lines = String(raw).split('\n');
            for (var i = 0; i < lines.length; i++) {
                var line = lines[i];
                var eq = line.indexOf('=');
                if (eq > 0) prefs[line.substring(0, eq)] = line.substring(eq + 1);
            }
            return prefs;
        }

        /* プリセットを保存（失敗時は黙って継続） / Save prefs (silently ignore failure) */
        function savePrefs(prefs) {
            try {
                var file = getPrefsFile();
                file.encoding = 'UTF-8';
                file.open('w');
                var lines = [];
                for (var key in prefs) {
                    if (prefs.hasOwnProperty(key)) lines.push(key + '=' + prefs[key]);
                }
                file.write(lines.join('\n'));
                file.close();
            } catch (e) { /* 保存できなくても処理は止めない */ }
        }

        /* ドキュメントから現在のファイル情報を収集。Illustrator は未保存ドキュメントでも fullName を返すため、
           fullName.exists で保存済みかを判定する
           / Collect file info from the active document; in Illustrator fullName exists even when unsaved,
           so use fullName.exists to detect a saved-to-disk file */
        function gatherDocumentInfo(doc) {
            var fullName;
            var savedToDisk = false;
            try {
                fullName = doc.fullName;
                savedToDisk = !!(fullName && fullName.exists);
            } catch (e) { /* unsaved document */ }
            var rawName = (fullName && fullName.name) ? fullName.name : doc.name;
            var currentName = decodePercentEncoded(rawName);
            var parentFolderName = '';
            if (savedToDisk && fullName.parent) {
                parentFolderName = decodePercentEncoded(fullName.parent.name);
            }
            return {
                currentName: currentName,
                baseName: stripExtension(currentName),
                fsPath: savedToDisk ? fullName.fsName : null,
                folder: savedToDisk ? fullName.parent : null,
                parentFolderName: parentFolderName
            };
        }

        /* 保存先フォルダを確定（未保存なら saveDialog で取得）。キャンセル時は null / Resolve the target folder (prompt if needed); null on cancel */
        function ensureTargetFolder(folder) {
            if (folder) return folder;
            var picked = pickAiDestination(L('message.chooseDestination'));
            return picked ? picked.parent : null;
        }

        /* 既存同名（自分自身を除く）の上書き確認。OK なら true / Confirm overwrite for an existing file (excluding self); true if approved */
        function confirmOverwriteIfExists(destFile, originalFsPath) {
            if (!destFile.exists) return true;
            if (destFile.fsName === originalFsPath) return true;
            return confirm(L('message.confirmOverwrite') + '\n\n' + destFile.fsName);
        }

        /* 旧ファイルを削除（失敗時は黙って継続） / Remove the original file (silently ignore failure) */
        function removeOriginalFile(originalFsPath, destFsPath) {
            if (!originalFsPath || originalFsPath === destFsPath) return;
            var file = File(originalFsPath);
            if (!file.exists) return;
            try { file.remove(); } catch (e) { /* 削除できない場合は黙って継続 */ }
        }

        /* モード別の出力処理（rename / saveAs / copy） / Execute the output according to the selected mode */
        function executeOutput(doc, destFile, mode, originalFsPath) {
            if (mode === 'copy' && originalFsPath) {
                // 現在の変更を元ファイルへ保存してから、物理ファイルとしてコピー
                if (!doc.saved) doc.save();
                File(originalFsPath).copy(destFile);
                return;
            }
            // rename / saveAs / 未保存ドキュメントの copy: 新名で保存
            doc.saveAs(destFile, new IllustratorSaveOptions());
            if (mode === 'rename') {
                removeOriginalFile(originalFsPath, destFile.fsName);
            }
        }

        /* segments と UI 状態から最終ファイル名（拡張子なし）を構築。SEGMENT_ORDER に従う
           / Build the final filename from segments and UI state, following SEGMENT_ORDER */
        function buildFinalName(segments, uiState) {
            // 区切り記号: 明示選択があればそれを、無ければ元のファイル名で優勢な区切りを使う
            var defaultSep;
            if (uiState.separator === '-' || uiState.separator === '_') {
                defaultSep = uiState.separator;
            } else {
                defaultSep = dominantSeparator(segments);
            }

            function valueForKind(kind) {
                if (kind === 'base') {
                    return getFirstSegmentValue(segments, 'base');
                }
                if (kind === 'title') {
                    if (uiState.titleMode === 'none') return '';
                    if (uiState.titleMode === 'parent') return sanitizeFilename(uiState.parentFolderName);
                    return sanitizeFilename(uiState.titleText);
                }
                if (kind === 'status') {
                    return uiState.status || '';
                }
                if (kind === 'timestamp') {
                    if (uiState.timestamp === 'date') return todayTimestamp();
                    if (uiState.timestamp === 'dateDash') return todayTimestamp('-');
                    return '';
                }
                if (kind === 'version') {
                    if (uiState.version === 'none') return '';
                    return formatVersion(getFirstSegmentValue(segments, 'version'), uiState.version);
                }
                return '';
            }

            // ピースごとに区切り統一（YYYY-MM-DD のタイムスタンプだけは内部 "-" を保護）
            var order = (uiState.segmentOrder && uiState.segmentOrder.length) ? uiState.segmentOrder : SEGMENT_ORDER;
            var pieces = [];
            for (var i = 0; i < order.length; i++) {
                var kind = order[i];
                var value = valueForKind(kind);
                if (!value) continue;
                var isProtectedDate = (kind === 'timestamp' && uiState.timestamp === 'dateDash');
                if (!isProtectedDate && FEATURE_SEPARATOR) {
                    // 区切り記号統一。FEATURE_DOT_NORMALIZE が true なら "." も対象
                    if (uiState.separator === '-') {
                        value = FEATURE_DOT_NORMALIZE ? value.replace(/[_.]/g, '-') : value.replace(/_/g, '-');
                    } else if (uiState.separator === '_') {
                        value = FEATURE_DOT_NORMALIZE ? value.replace(/[-.]/g, '_') : value.replace(/-/g, '_');
                    }
                }
                pieces.push(value);
            }
            var result = pieces.join(defaultSep);

            // スペースの置換は区切り記号統一の後に行う（連続スペースは 1 文字に畳む）
            if (FEATURE_SPACES) {
                if (uiState.spaces === '-') {
                    result = result.replace(/\s+/g, '-');
                } else if (uiState.spaces === '_') {
                    result = result.replace(/\s+/g, '_');
                }
            }

            return result;
        }

        /* segments から優勢な区切り記号を返す。同数なら "-" / Dominant separator across segments (defaults to "-") */
        function dominantSeparator(segments) {
            var dashes = 0, underscores = 0;
            for (var i = 0; i < segments.length; i++) {
                if (segments[i].sep === '-') dashes++;
                else if (segments[i].sep === '_') underscores++;
            }
            return (underscores > dashes) ? '_' : '-';
        }

        // =========================================
        // ダイアログビルダー / Dialog builders
        // =========================================

        /* パネルの共通設定（向き・整列・余白・子要素間隔） / Apply shared panel settings (orientation, alignment, margins, spacing) */
        function setupPanel(panel, spacing) {
            panel.orientation = 'column';
            panel.alignChildren = ['fill', 'top'];
            panel.alignment = 'fill';
            panel.margins = PANEL_MARGINS;
            panel.spacing = (typeof spacing === 'number') ? spacing : PANEL_SPACING;
        }

        /* ソートパネルを構築（しない / する + [ソート] ボタン）。ボタンは「する」のとき有効
           / Build the sort panel: Off/On radios + [Sort] button (enabled only when "On") */
        function buildSortPanel(parent, prefs) {
            var panel = parent.add('panel', undefined, L('panel.sort'));
            setupPanel(panel);
            var sortOffRadio = panel.add('radiobutton', undefined, L('radio.sortOff'));
            sortOffRadio.helpTip = L('tip.sort');
            var sortOnRadio = panel.add('radiobutton', undefined, L('radio.sortOn'));
            sortOnRadio.helpTip = L('tip.sort');
            var initialSort = (prefs && prefs.sort === 'on') ? 'on' : 'off';
            sortOffRadio.value = (initialSort === 'off');
            sortOnRadio.value = (initialSort === 'on');
            var sortButton = panel.add('button', undefined, L('button.sort'));
            sortButton.enabled = (initialSort === 'on');
            function syncSortButtonEnabled() {
                sortButton.enabled = sortOnRadio.value;
            }
            // onClick は呼び出し側で wire（refreshPreviews と組み合わせるため）
            return {
                panel: panel,
                sortOffRadio: sortOffRadio,
                sortOnRadio: sortOnRadio,
                sortButton: sortButton,
                isSortOn: function () { return sortOnRadio.value; },
                syncSortButtonEnabled: syncSortButtonEnabled
            };
        }

        /* モード選択パネルを構築（リネーム / 別名で保存 / コピーを保存） / Build the mode panel (Rename / Save As / Save a Copy) */
        function buildModePanel(parent, prefs) {
            var panel = parent.add('panel', undefined, L('panel.mode'));
            setupPanel(panel);
            var renameRadio = panel.add('radiobutton', undefined, L('radio.rename'));
            renameRadio.helpTip = L('tip.rename');
            var saveAsRadio = panel.add('radiobutton', undefined, L('radio.saveAs'));
            saveAsRadio.helpTip = L('tip.saveAs');
            var saveCopyRadio = panel.add('radiobutton', undefined, L('radio.saveCopy'));
            saveCopyRadio.helpTip = L('tip.saveCopy');
            // 初期選択（プリセットがあれば優先、無ければ「別名で保存」）
            var initialMode = (prefs && (prefs.mode === 'rename' || prefs.mode === 'copy')) ? prefs.mode : 'saveAs';
            renameRadio.value = (initialMode === 'rename');
            saveAsRadio.value = (initialMode === 'saveAs');
            saveCopyRadio.value = (initialMode === 'copy');
            return {
                panel: panel,
                renameRadio: renameRadio,
                saveAsRadio: saveAsRadio,
                saveCopyRadio: saveCopyRadio,
                /* 現在選択中のモード（'rename' / 'saveAs' / 'copy'） / Currently selected mode */
                getMode: function () {
                    if (renameRadio.value) return 'rename';
                    if (saveCopyRadio.value) return 'copy';
                    return 'saveAs';
                }
            };
        }

        /* ファイル名パネルを構築（現在名・変更後名のみ） / Build the file-name panel (current / final only) */
        function buildFilenamePanel(parent, currentName) {
            var panel = parent.add('panel', undefined, L('panel.filename'));
            setupPanel(panel);

            var currentNameRow = panel.add('group');
            currentNameRow.orientation = 'row';
            var currentNameLabel = currentNameRow.add('statictext', undefined, labelText('label.currentName'));
            currentNameRow.add('statictext', undefined, currentName);

            var finalNameRow = panel.add('group');
            finalNameRow.orientation = 'row';
            var finalNameLabel = finalNameRow.add('statictext', undefined, labelText('label.finalName'));
            // 「変更後：」は statictext のためレイアウト後にサイズ固定。
            // 現在のファイル名と「入力フィールド + 余白」の大きい方を確保しておく
            var finalNameValue = finalNameRow.add('statictext', undefined, currentName + '.ai');
            var currentNameWidth = panel.graphics.measureString(currentName + '.ai').width;
            finalNameValue.preferredSize.width = Math.max(currentNameWidth + 20, NEW_NAME_FIELD_WIDTH + 150);

            alignLabelWidths(panel, [
                labelText('label.currentName'),
                labelText('label.finalName')
            ], [currentNameLabel, finalNameLabel]);

            return {
                panel: panel,
                finalNameValue: finalNameValue
            };
        }

        /* オプションパネルを構築（ベース表示・タイトル選択・タイムスタンプ・バージョン番号・区切り） / Build the options panel */
        function buildOptionsPanel(parent, segments, prefs, parentFolderName) {
            var panel = parent.add('panel', undefined, L('panel.options'));
            setupPanel(panel);

            // ベースは UI には出さず、ロジック内（buildFinalName）で segments から取得
            // タイトル入力欄はデフォルト空白（元ファイル名のテキスト部は自動投入しない）

            // タイトル: 1 行目 = ラベル + 3 ラジオ、2 行目 = 「指定」用の入力欄
            var titleSection = panel.add('group');
            titleSection.orientation = 'column';
            titleSection.alignChildren = ['fill', 'top'];
            titleSection.spacing = 4;

            var titleRow = titleSection.add('group');
            titleRow.orientation = 'row';
            titleRow.alignment = ['left', 'top'];
            var titleLabel = titleRow.add('statictext', undefined, labelText('label.title'));
            titleLabel.helpTip = L('tip.title');
            var titleNoneRadio = titleRow.add('radiobutton', undefined, L('radio.titleNone'));
            titleNoneRadio.helpTip = L('tip.title');
            var titleParentRadio = titleRow.add('radiobutton', undefined, L('radio.titleParent'));
            titleParentRadio.helpTip = parentFolderName
                ? L('tip.title') + ' (' + parentFolderName + ')'
                : L('tip.title');
            if (!parentFolderName) titleParentRadio.enabled = false;
            var titleCustomRadio = titleRow.add('radiobutton', undefined, L('radio.titleCustom'));
            titleCustomRadio.helpTip = L('tip.title');

            // 「指定」用の入力欄は次の行（ラベル列幅だけ左に余白を入れて radios に揃える）
            var titleFieldRow = titleSection.add('group');
            titleFieldRow.orientation = 'row';
            titleFieldRow.alignment = ['left', 'top'];
            var titleFieldSpacer = titleFieldRow.add('statictext', undefined, '');
            var titleField = titleFieldRow.add('edittext', undefined, '');
            titleField.preferredSize.width = NEW_NAME_FIELD_WIDTH;
            titleField.helpTip = L('tip.title');

            // 初期モード: プリセット優先、無ければ 'none'。parent 指定で親フォルダー名が無ければ 'none'
            var initialTitleMode;
            if (prefs && (prefs.titleMode === 'none' || prefs.titleMode === 'parent' || prefs.titleMode === 'custom')) {
                initialTitleMode = prefs.titleMode;
            } else {
                initialTitleMode = 'none';
            }
            if (initialTitleMode === 'parent' && !parentFolderName) {
                initialTitleMode = 'none';
            }
            titleNoneRadio.value = (initialTitleMode === 'none');
            titleParentRadio.value = (initialTitleMode === 'parent');
            titleCustomRadio.value = (initialTitleMode === 'custom');
            titleField.enabled = (initialTitleMode === 'custom');

            function syncTitleFieldEnabled() {
                titleField.enabled = titleCustomRadio.value;
            }

            // ステータス（dropdown。先頭は「なし」）
            var statusLabel = null, statusDropdown = null;
            if (FEATURE_STATUS) {
                var statusRow = panel.add('group');
                statusRow.orientation = 'row';
                statusLabel = statusRow.add('statictext', undefined, labelText('label.status'));
                statusLabel.helpTip = L('tip.status');
                statusDropdown = statusRow.add('dropdownlist', undefined, undefined);
                statusDropdown.helpTip = L('tip.status');
                var parsedStatus = getFirstSegmentValue(segments, 'status');
                var initialStatus = parsedStatus || ((prefs && prefs.status) ? prefs.status : '');
                var initialStatusIdx = -1;
                for (var siItem = 0; siItem < STATUS_ITEMS.length; siItem++) {
                    var sItem = STATUS_ITEMS[siItem];
                    var ddItem = statusDropdown.add('item', (lang === 'ja' ? sItem.ja : sItem.en));
                    if (isStatusDivider(sItem)) {
                        ddItem.enabled = false;
                        continue;
                    }
                    if (initialStatusIdx === -1 && sItem.value === initialStatus) {
                        initialStatusIdx = siItem;
                    }
                }
                if (initialStatusIdx === -1) initialStatusIdx = 0;
                statusDropdown.selection = initialStatusIdx;
            }

            // タイムスタンプ（なし / YYYYMMDD / YYYY-MM-DD。デフォルト YYYYMMDD）
            var timestampRow = panel.add('group');
            timestampRow.orientation = 'row';
            var timestampLabel = timestampRow.add('statictext', undefined, labelText('label.timestamp'));
            timestampLabel.helpTip = L('tip.timestamp');
            var timestampNoneRadio = timestampRow.add('radiobutton', undefined, L('radio.timestampNone'));
            timestampNoneRadio.helpTip = L('tip.timestamp');
            var timestampDateRadio = timestampRow.add('radiobutton', undefined, L('radio.timestampDate'));
            timestampDateRadio.helpTip = L('tip.timestamp');
            var timestampDateDashRadio = timestampRow.add('radiobutton', undefined, L('radio.timestampDateDash'));
            timestampDateDashRadio.helpTip = L('tip.timestamp');
            var initialTimestamp = (prefs && (prefs.timestamp === 'none' || prefs.timestamp === 'dateDash')) ? prefs.timestamp : 'date';
            timestampNoneRadio.value = (initialTimestamp === 'none');
            timestampDateRadio.value = (initialTimestamp === 'date');
            timestampDateDashRadio.value = (initialTimestamp === 'dateDash');

            // バージョン番号（なし / -vN / -v0N。デフォルト -vN）
            var versionRow = panel.add('group');
            versionRow.orientation = 'row';
            var versionLabel = versionRow.add('statictext', undefined, labelText('label.version'));
            versionLabel.helpTip = L('tip.version');
            var versionNoneRadio = versionRow.add('radiobutton', undefined, L('radio.versionNone'));
            versionNoneRadio.helpTip = L('tip.version');
            var versionShortRadio = versionRow.add('radiobutton', undefined, L('radio.versionShort'));
            versionShortRadio.helpTip = L('tip.version');
            var versionPaddedRadio = versionRow.add('radiobutton', undefined, L('radio.versionPadded'));
            versionPaddedRadio.helpTip = L('tip.version');
            var initialVersion = (prefs && (prefs.version === 'none' || prefs.version === 'padded')) ? prefs.version : 'short';
            versionNoneRadio.value = (initialVersion === 'none');
            versionShortRadio.value = (initialVersion === 'short');
            versionPaddedRadio.value = (initialVersion === 'padded');

            // 区切り記号（変更しない / - / _ 横並び。デフォルト "-"）
            var separatorLabel = null, noChangeRadio = null, dashRadio = null, underscoreRadio = null;
            if (FEATURE_SEPARATOR) {
                var separatorRow = panel.add('group');
                separatorRow.orientation = 'row';
                separatorLabel = separatorRow.add('statictext', undefined, labelText('label.separator'));
                separatorLabel.helpTip = L('tip.separator');
                noChangeRadio = separatorRow.add('radiobutton', undefined, L('radio.noChange'));
                noChangeRadio.helpTip = L('tip.separator');
                dashRadio = separatorRow.add('radiobutton', undefined, '-');
                dashRadio.helpTip = L('tip.separator');
                underscoreRadio = separatorRow.add('radiobutton', undefined, '_');
                underscoreRadio.helpTip = L('tip.separator');
                // prefs があればそれを優先、無ければ FEATURE_SEPARATOR の指定文字（'-' or '_'）を初期値に
                var defaultSeparator = (FEATURE_SEPARATOR === '_') ? '_' : '-';
                var initialSeparator = (prefs && (prefs.separator === '' || prefs.separator === '-' || prefs.separator === '_')) ? prefs.separator : defaultSeparator;
                noChangeRadio.value = (initialSeparator === '');
                dashRadio.value = (initialSeparator === '-');
                underscoreRadio.value = (initialSeparator === '_');
            }

            // スペースの扱い（変更しない / -に変更 / _に変更。デフォルト "変更しない"）
            var spacesLabel = null, spaceNoChangeRadio = null, spaceDashRadio = null, spaceUnderscoreRadio = null;
            if (FEATURE_SPACES) {
                var spacesRow = panel.add('group');
                spacesRow.orientation = 'row';
                spacesLabel = spacesRow.add('statictext', undefined, labelText('label.spaces'));
                spacesLabel.helpTip = L('tip.spaces');
                spaceNoChangeRadio = spacesRow.add('radiobutton', undefined, L('radio.noChange'));
                spaceNoChangeRadio.helpTip = L('tip.spaces');
                spaceDashRadio = spacesRow.add('radiobutton', undefined, L('radio.changeToDash'));
                spaceDashRadio.helpTip = L('tip.spaces');
                spaceUnderscoreRadio = spacesRow.add('radiobutton', undefined, L('radio.changeToUnderscore'));
                spaceUnderscoreRadio.helpTip = L('tip.spaces');
                // prefs があればそれを優先、無ければ FEATURE_SPACES の指定文字（'-' or '_'）を初期値に
                var defaultSpaces = (FEATURE_SPACES === '_') ? '_' : '-';
                var initialSpaces = (prefs && (prefs.spaces === '' || prefs.spaces === '-' || prefs.spaces === '_')) ? prefs.spaces : defaultSpaces;
                spaceNoChangeRadio.value = (initialSpaces === '');
                spaceDashRadio.value = (initialSpaces === '-');
                spaceUnderscoreRadio.value = (initialSpaces === '_');
            }

            // ラベル幅を統一（FEATURE で UI 非表示の行は除外）
            var labelTexts = [labelText('label.title')];
            var labelControls = [titleLabel];
            if (FEATURE_STATUS) { labelTexts.push(labelText('label.status')); labelControls.push(statusLabel); }
            labelTexts.push(labelText('label.timestamp')); labelControls.push(timestampLabel);
            labelTexts.push(labelText('label.version')); labelControls.push(versionLabel);
            if (FEATURE_SEPARATOR) { labelTexts.push(labelText('label.separator')); labelControls.push(separatorLabel); }
            if (FEATURE_SPACES) { labelTexts.push(labelText('label.spaces')); labelControls.push(spacesLabel); }
            alignLabelWidths(panel, labelTexts, labelControls);

            // タイトル 2 行目「指定」入力欄の左余白をラベル列幅と一致させる
            titleFieldSpacer.preferredSize = [titleLabel.preferredSize.width, 1];

            return {
                panel: panel,
                titleNoneRadio: titleNoneRadio,
                titleParentRadio: titleParentRadio,
                titleCustomRadio: titleCustomRadio,
                titleField: titleField,
                syncTitleFieldEnabled: syncTitleFieldEnabled,
                statusDropdown: statusDropdown,
                timestampNoneRadio: timestampNoneRadio,
                timestampDateRadio: timestampDateRadio,
                timestampDateDashRadio: timestampDateDashRadio,
                versionNoneRadio: versionNoneRadio,
                versionShortRadio: versionShortRadio,
                versionPaddedRadio: versionPaddedRadio,
                noChangeRadio: noChangeRadio,
                dashRadio: dashRadio,
                underscoreRadio: underscoreRadio,
                spaceNoChangeRadio: spaceNoChangeRadio,
                spaceDashRadio: spaceDashRadio,
                spaceUnderscoreRadio: spaceUnderscoreRadio,
                /* '' = 変更しない、'-' / '_' = 統一。FEATURE_SEPARATOR=false なら常に '' */
                getSeparator: function () {
                    if (!FEATURE_SEPARATOR || !noChangeRadio) return '';
                    if (noChangeRadio.value) return '';
                    if (dashRadio.value) return '-';
                    return '_';
                },
                /* '' = 変更しない、'-' / '_' = 半角スペースを置換。FEATURE_SPACES=false なら常に '' */
                getSpaces: function () {
                    if (!FEATURE_SPACES || !spaceNoChangeRadio) return '';
                    if (spaceNoChangeRadio.value) return '';
                    if (spaceDashRadio.value) return '-';
                    return '_';
                },
                /* 'none' / 'date' / 'dateDash' */
                getTimestamp: function () {
                    if (timestampNoneRadio.value) return 'none';
                    if (timestampDateDashRadio.value) return 'dateDash';
                    return 'date';
                },
                /* STATUS_ITEMS の value（'' = なし）。FEATURE_STATUS=false なら常に '' */
                getStatus: function () {
                    if (!FEATURE_STATUS || !statusDropdown) return '';
                    var idx = statusDropdown.selection ? statusDropdown.selection.index : 0;
                    var item = STATUS_ITEMS[idx];
                    if (!item || isStatusDivider(item)) return '';
                    return item.value;
                },
                /* 'none' / 'short' / 'padded' */
                getVersion: function () {
                    if (versionNoneRadio.value) return 'none';
                    if (versionShortRadio.value) return 'short';
                    return 'padded';
                },
                /* 'none' / 'parent' / 'custom' */
                getTitleMode: function () {
                    if (titleNoneRadio.value) return 'none';
                    if (titleParentRadio.value) return 'parent';
                    return 'custom';
                }
            };
        }

        /* 並び順を ↑↓ で編集するサブダイアログ。OK で新しい順序を返す。キャンセルで null
           / Sub-dialog to reorder segments with ↑↓; returns the new order, or null on cancel */
        function openSortDialog(initialOrder) {
            var dlg = new Window('dialog', L('sort.title'));
            dlg.opacity = DIALOG_OPACITY;
            dlg.orientation = 'column';
            dlg.alignChildren = 'fill';
            dlg.margins = 15;
            dlg.spacing = 10;

            var hint = dlg.add('statictext', undefined, L('sort.hint'));
            hint.alignment = 'left';

            var body = dlg.add('group');
            body.orientation = 'row';
            body.alignChildren = ['fill', 'fill'];

            var order = initialOrder.slice();
            var list = body.add('listbox', undefined, []);
            list.preferredSize = [200, 140];
            for (var i = 0; i < order.length; i++) {
                list.add('item', L('label.' + order[i]));
            }
            list.selection = 0;

            var btnCol = body.add('group');
            btnCol.orientation = 'column';
            btnCol.alignChildren = 'fill';
            var upBtn = btnCol.add('button', undefined, '↑');
            upBtn.preferredSize = [36, 24];
            var downBtn = btnCol.add('button', undefined, '↓');
            downBtn.preferredSize = [36, 24];

            function refreshList(newIdx) {
                for (var i = 0; i < order.length; i++) {
                    list.items[i].text = L('label.' + order[i]);
                }
                list.selection = newIdx;
            }
            upBtn.onClick = function () {
                var idx = list.selection ? list.selection.index : -1;
                if (idx <= 0) return;
                var tmp = order[idx - 1];
                order[idx - 1] = order[idx];
                order[idx] = tmp;
                refreshList(idx - 1);
            };
            downBtn.onClick = function () {
                var idx = list.selection ? list.selection.index : -1;
                if (idx < 0 || idx >= order.length - 1) return;
                var tmp = order[idx + 1];
                order[idx + 1] = order[idx];
                order[idx] = tmp;
                refreshList(idx + 1);
            };

            var buttons = dlg.add('group');
            buttons.alignment = ['center', 'top'];
            buttons.add('button', undefined, L('button.cancel'), { name: 'cancel' });
            buttons.add('button', undefined, 'OK', { name: 'ok' });

            if (dlg.show() !== 1) return null;
            return order;
        }

        /* ES3 互換: 配列に値が含まれるか / ES3-safe array contains check */
        function arrayContains(arr, val) {
            for (var i = 0; i < arr.length; i++) {
                if (arr[i] === val) return true;
            }
            return false;
        }

        /* 保存された並び順文字列を妥当性チェックして配列で返す。不正なら null
           / Parse a comma-separated segment order; returns null if invalid */
        function parseSegmentOrderPref(value) {
            if (!value) return null;
            var parts = String(value).split(',');
            if (parts.length !== SEGMENT_ORDER.length) return null;
            var seen = {};
            for (var i = 0; i < parts.length; i++) {
                var kind = parts[i];
                if (!arrayContains(SEGMENT_ORDER, kind) || seen[kind]) return null;
                seen[kind] = true;
            }
            return parts;
        }

        /* labelTexts の最大幅に controls の幅を揃える。checkbox/radiobutton が含まれていれば
           インジケーター分（+20px）を全コントロールに加算して右端を揃える
           / Align controls' widths to the widest label. If any control is a checkbox/radiobutton,
           add indicator width (+20px) to all so their right edges align */
        function alignLabelWidths(panel, labelTexts, controls) {
            var graphics = panel.graphics;
            var maxWidth = 0;
            for (var i = 0; i < labelTexts.length; i++) {
                var width = graphics.measureString(labelTexts[i]).width;
                if (width > maxWidth) maxWidth = width;
            }
            maxWidth += 12;
            var needsIndicator = false;
            for (var k = 0; k < controls.length; k++) {
                if (controls[k].type === 'checkbox' || controls[k].type === 'radiobutton') {
                    needsIndicator = true;
                    break;
                }
            }
            if (needsIndicator) maxWidth += 20;
            for (var j = 0; j < controls.length; j++) {
                controls[j].preferredSize = [maxWidth, controls[j].preferredSize.height || 20];
            }
        }

        /* ダイアログ全体を組み立て、イベント配線とプレビューを行う / Compose the full dialog, wire events, and run live preview */
        function createDialog(segments, currentName, prefs, parentFolderName) {
            var dialog = new Window('dialog', L('dialog.title') + ' ' + SCRIPT_VERSION);
            dialog.opacity = DIALOG_OPACITY;
            dialog.orientation = 'column';
            dialog.alignChildren = 'fill';

            // 上部 2 カラム: 動作パネル + ソートパネル（FEATURE_SORT=false ならソートパネル無し）
            var topRow = dialog.add('group');
            topRow.orientation = 'row';
            topRow.alignChildren = ['fill', 'top'];
            var mode = buildModePanel(topRow, prefs);
            var sort = FEATURE_SORT ? buildSortPanel(topRow, prefs) : null;

            // 並び順カスタム値（prefs に保存されたものを採用、不正・未保存ならデフォルト）
            var customOrder = FEATURE_SORT
                ? (parseSegmentOrderPref(prefs && prefs.segmentOrder) || SEGMENT_ORDER.slice())
                : SEGMENT_ORDER.slice();

            var filename = buildFilenamePanel(dialog, currentName);
            var options = buildOptionsPanel(dialog, segments, prefs, parentFolderName);

            // ---- ライブプレビュー ----
            function currentUIState() {
                return {
                    titleMode: options.getTitleMode(),
                    titleText: options.titleField.text,
                    parentFolderName: parentFolderName,
                    status: options.getStatus(),
                    timestamp: options.getTimestamp(),
                    version: options.getVersion(),
                    separator: options.getSeparator(),
                    spaces: options.getSpaces(),
                    sort: (FEATURE_SORT && sort && sort.isSortOn()) ? 'on' : 'off',
                    customSegmentOrder: customOrder.slice(),
                    segmentOrder: (FEATURE_SORT && sort && sort.isSortOn()) ? customOrder : SEGMENT_ORDER
                };
            }

            function refreshPreviews() {
                options.syncTitleFieldEnabled();
                var finalBase = buildFinalName(segments, currentUIState());
                filename.finalNameValue.text = finalBase + '.ai';
            }

            options.titleNoneRadio.onClick = refreshPreviews;
            options.titleParentRadio.onClick = refreshPreviews;
            options.titleCustomRadio.onClick = refreshPreviews;
            options.titleField.onChanging = refreshPreviews;
            if (options.statusDropdown) options.statusDropdown.onChange = refreshPreviews;
            options.timestampNoneRadio.onClick = refreshPreviews;
            options.timestampDateRadio.onClick = refreshPreviews;
            options.timestampDateDashRadio.onClick = refreshPreviews;
            options.versionNoneRadio.onClick = refreshPreviews;
            options.versionShortRadio.onClick = refreshPreviews;
            options.versionPaddedRadio.onClick = refreshPreviews;
            if (options.noChangeRadio) {
                options.noChangeRadio.onClick = refreshPreviews;
                options.dashRadio.onClick = refreshPreviews;
                options.underscoreRadio.onClick = refreshPreviews;
            }
            if (options.spaceNoChangeRadio) {
                options.spaceNoChangeRadio.onClick = refreshPreviews;
                options.spaceDashRadio.onClick = refreshPreviews;
                options.spaceUnderscoreRadio.onClick = refreshPreviews;
            }

            // ソートパネルの ON/OFF とサブダイアログ起動（FEATURE_SORT のとき）
            if (FEATURE_SORT && sort) {
                var onSortToggle = function () {
                    sort.syncSortButtonEnabled();
                    refreshPreviews();
                };
                sort.sortOffRadio.onClick = onSortToggle;
                sort.sortOnRadio.onClick = onSortToggle;
                sort.sortButton.onClick = function () {
                    var newOrder = openSortDialog(customOrder);
                    if (newOrder) {
                        customOrder = newOrder;
                        refreshPreviews();
                    }
                };
            }

            refreshPreviews();

            // ---- ボタン（右寄せ Cancel / OK） ----
            var buttonGroup = dialog.add('group');
            buttonGroup.alignment = ['right', 'top'];
            buttonGroup.add('button', undefined, L('button.cancel'), { name: 'cancel' });
            buttonGroup.add('button', undefined, 'OK', { name: 'ok' });

            return {
                dialog: dialog,
                getUIState: currentUIState,
                getMode: mode.getMode
            };
        }

        // =========================================
        // メイン / Main
        // =========================================

        /* エントリポイント。ダイアログを開き、選択モードに応じた出力を実行 / Entry point: open the dialog and execute the selected mode */
        function main() {
            if (app.documents.length === 0) {
                alert(L('message.noDoc'));
                return;
            }

            var doc = app.activeDocument;
            var info = gatherDocumentInfo(doc);
            var segments = parseFileName(info.baseName);
            var prefs = loadPrefs();

            var ui = createDialog(segments, info.currentName, prefs, info.parentFolderName);
            if (ui.dialog.show() !== 1) return; // キャンセル

            var uiState = ui.getUIState();
            var newBaseName = buildFinalName(segments, uiState);
            if (!newBaseName) {
                alert(L('message.emptyName'));
                return;
            }

            var targetFolder = ensureTargetFolder(info.folder);
            if (!targetFolder) return; // キャンセル

            var destFile = File(targetFolder.fsName + '/' + newBaseName + '.ai');
            if (!confirmOverwriteIfExists(destFile, info.fsPath)) return;

            try {
                executeOutput(doc, destFile, ui.getMode(), info.fsPath);
                // 成功したら今回の選択をプリセットとして保存
                var prefsToSave = {
                    mode: ui.getMode(),
                    titleMode: uiState.titleMode,
                    timestamp: uiState.timestamp,
                    version: uiState.version
                };
                if (FEATURE_STATUS) prefsToSave.status = uiState.status;
                if (FEATURE_SEPARATOR) prefsToSave.separator = uiState.separator;
                if (FEATURE_SPACES) prefsToSave.spaces = uiState.spaces;
                if (FEATURE_SORT) {
                    prefsToSave.sort = uiState.sort;
                    prefsToSave.segmentOrder = uiState.customSegmentOrder.join(',');
                }
                savePrefs(prefsToSave);
            } catch (e) {
                alert(L('message.saveFailed') + '\n' + e);
            }
        }

        main();

    })();
