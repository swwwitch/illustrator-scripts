#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

    (function () {

        var SCRIPT_VERSION = "v1.2.0";
        var COORDINATE_PRECISION_DIGITS = 3;

        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var lang = getCurrentLang();

        /* ------------------------------- */
        /* 日英ラベル定義 / Japanese-English label definitions */
        /* ------------------------ */

        var LABELS = {
            dialogTitle: {
                ja: "行列にリネーム " + SCRIPT_VERSION,
                en: "Rename by Row-Column " + SCRIPT_VERSION
            },
            cancel: { ja: "キャンセル", en: "Cancel" },
            noDocument: { ja: "ドキュメントを開いてください。", en: "Please open a document." },
            noArtboards: { ja: "アートボードが存在しません。", en: "No artboards found." },
            close: { ja: "キャンセル", en: "Cancel" },
            namingPanel: { ja: "行-列", en: "Row-Column" },
            namingEnable: { ja: "アートボード名を 行-列 形式に更新", en: "Update artboard names as row-column" },
            namingFromPosition: { ja: "位置から", en: "From position" },
            namingFromExisting: { ja: "既存名を再フォーマット", en: "Reformat existing names" },
            namingSeparator: { ja: "区切り文字：", en: "Separator:" },
            namingPadWidth: { ja: "ゼロ埋め桁数：", en: "Pad width:" }
        };

        function main() {
            if (app.documents.length === 0) {
                alert(LABELS.noDocument[lang]);
                return;
            }
            var doc = app.activeDocument;
            if (doc.artboards.length === 0) {
                alert(LABELS.noArtboards[lang]);
                return;
            }

            var context = buildContext(doc);
            var ui = buildDialogUI();

            bindEvents(doc, context, ui);

            ui.dialog.center();
            ui.dialog.show();
        }

        function buildContext(doc) {
            var entries = [];
            for (var i = 0; i < doc.artboards.length; i++) {
                entries.push({ artboardRect: doc.artboards[i].artboardRect });
            }
            var autoTol = calculateAutoTolerance(entries);
            return {
                tolerance: Math.round(autoTol * 1.1)
            };
        }

        function bindEvents(doc, context, ui) {
            ui.executeBtn.onClick = function () {
                executeApply(doc, context, ui);
            };

            ui.closeBtn.onClick = function () {
                ui.dialog.close(-1);
            };
        }

        function executeApply(doc, context, ui) {
            var separator = ui.sepUnderscore.value ? "_" : (ui.sepX.value ? "x" : "-");
            var padWidth = ui.padW3.value ? 3 : (ui.padW2.value ? 2 : 1);
            var tolerance = context.tolerance;

            if (ui.fromPositionRadio.value) {
                /* まず［アートボード］パネル順をソート結果通りに再構築してから命名 */
                var topLeftSorter = function (_artboards, dp) {
                    sortArtboardsTopLeftWithTolerance(_artboards, dp, tolerance);
                };
                rebuildArtboardsInSortedOrder(doc, topLeftSorter, COORDINATE_PRECISION_DIGITS);
                renameArtboardsFromPositions(doc, separator, padWidth, tolerance);
            } else {
                renameArtboardsFromExistingNames(doc, separator, padWidth);
            }
            ui.dialog.close(1);
        }

        function buildDialogUI() {
            var dialog = new Window("dialog", LABELS.dialogTitle[lang]);
            dialog.orientation = "column";
            dialog.alignChildren = ['fill', 'top'];
            dialog.spacing = 15;

            var mainColumn = dialog.add("group");
            mainColumn.orientation = "column";
            mainColumn.alignChildren = ['fill', 'top'];
            mainColumn.alignment = ['fill', 'top'];

            /* 行-列パネル / Row-Column panel */
            var namingGroup = mainColumn.add("panel", undefined, LABELS.namingPanel[lang]);
            namingGroup.orientation = "column";
            namingGroup.alignChildren = ['fill', 'top'];
            namingGroup.margins = [15, 20, 15, 10];

            var namingSettingsGroup = namingGroup.add("group");
            namingSettingsGroup.orientation = "column";
            namingSettingsGroup.alignChildren = ['fill', 'top'];

            /* 命名ソース / Naming source */
            var sourceRow = namingSettingsGroup.add("group");
            sourceRow.orientation = "row";
            sourceRow.alignChildren = ['left', 'center'];
            var fromPositionRadio = sourceRow.add("radiobutton", undefined, LABELS.namingFromPosition[lang]);
            var fromExistingRadio = sourceRow.add("radiobutton", undefined, LABELS.namingFromExisting[lang]);
            fromPositionRadio.value = true;

            /* 区切り文字 / Separator */
            var sepRow = namingSettingsGroup.add("group");
            sepRow.orientation = "row";
            sepRow.alignChildren = ['left', 'center'];
            var sepLabel = sepRow.add("statictext", undefined, LABELS.namingSeparator[lang]);
            sepLabel.preferredSize.width = 90;
            // sepLabel.justify = 'right';
            var sepHyphen = sepRow.add("radiobutton", undefined, "-");
            var sepUnderscore = sepRow.add("radiobutton", undefined, "_");
            var sepX = sepRow.add("radiobutton", undefined, "x");
            sepHyphen.value = true;

            /* ゼロ埋め桁数 / Zero-pad width */
            var padRow = namingSettingsGroup.add("group");
            padRow.orientation = "row";
            padRow.alignChildren = ['left', 'center'];
            var padLabel = padRow.add("statictext", undefined, LABELS.namingPadWidth[lang]);
            padLabel.preferredSize.width = 90;
            var padW1 = padRow.add("radiobutton", undefined, "0");
            var padW2 = padRow.add("radiobutton", undefined, "00");
            var padW3 = padRow.add("radiobutton", undefined, "000");
            padW1.value = true;

            /* 実行ボタン / Execute button */
            var executeBtn = namingGroup.add("button", undefined, LABELS.namingEnable[lang]);

            var buttonGroup = dialog.add("group");
            buttonGroup.orientation = "row";
            buttonGroup.alignment = ['center', 'bottom'];

            var closeBtn = buttonGroup.add('button', undefined, LABELS.close[lang]);

            return {
                dialog: dialog,
                fromPositionRadio: fromPositionRadio,
                sepUnderscore: sepUnderscore,
                sepX: sepX,
                padW1: padW1,
                padW2: padW2,
                padW3: padW3,
                closeBtn: closeBtn,
                executeBtn: executeBtn
            };
        }

        /* アートボードを並べ替えて［アートボード］パネル順を再構築
         * Rebuild Artboards panel order based on sorted result */
        function rebuildArtboardsInSortedOrder(doc, sorterFn, precisionDigits) {
            var decimalPlaces = Math.pow(10, precisionDigits || 3);

            var _artboards = [];
            var originalArtboards = [];
            for (var i = 0; i < doc.artboards.length; i++) {
                var a = doc.artboards[i];
                _artboards.push({
                    name: a.name,
                    artboardRect: a.artboardRect,
                    srcIndex: i
                });
                originalArtboards.push({
                    name: a.name,
                    artboardRect: a.artboardRect,
                    rulerOrigin: a.rulerOrigin,
                    rulerPAR: a.rulerPAR,
                    showCenter: a.showCenter,
                    showCrossHairs: a.showCrossHairs,
                    showSafeAreas: a.showSafeAreas
                });
            }

            sorterFn(_artboards, decimalPlaces);

            /* ソート結果順に末尾へ複製を追加 */
            for (var i = 0; i < _artboards.length; i++) {
                var s = _artboards[i];
                var src = originalArtboards[s.srcIndex];
                var b = doc.artboards.add(src.artboardRect);
                b.name = src.name;
                b.rulerOrigin = src.rulerOrigin;
                b.rulerPAR = src.rulerPAR;
                b.showCenter = src.showCenter;
                b.showCrossHairs = src.showCrossHairs;
                b.showSafeAreas = src.showSafeAreas;
            }

            /* 元のアートボードを後ろから削除 */
            for (var i = originalArtboards.length - 1; i >= 0; i--) {
                doc.artboards[i].remove();
            }
        }

        /* 配置から行に分割: (1)top で厳密ソート → (2)tolerance で行グループ化 → (3)各行内を left で厳密ソート
         * 非推移コンパレータを排除して結果を決定的にする */
        function groupArtboardsIntoRows(_artboards, decimalPlaces, tolerance) {
            if (_artboards.length === 0) return [];
            var dp = decimalPlaces || 1000;

            /* 座標系自動判定: top > bottom なら Y-UP（上ほど top が大きい）、そうでなければ Y-DOWN */
            var firstRect = _artboards[0].artboardRect;
            var isYUp = firstRect[1] > firstRect[3];

            /* (1) 厳密ソート: top で上→下、同 top は left 昇順 */
            _artboards.sort(function (a, b) {
                var ta = Math.round(a.artboardRect[1] * dp) / dp;
                var tb = Math.round(b.artboardRect[1] * dp) / dp;
                if (ta !== tb) return isYUp ? (tb - ta) : (ta - tb);
                return a.artboardRect[0] - b.artboardRect[0];
            });

            /* (2) 許容差で行グループ化 */
            var rows = [[_artboards[0]]];
            for (var j = 1; j < _artboards.length; j++) {
                var prevTop = Math.round(_artboards[j - 1].artboardRect[1] * dp) / dp;
                var currTop = Math.round(_artboards[j].artboardRect[1] * dp) / dp;
                if (Math.abs(prevTop - currTop) > tolerance) {
                    rows.push([_artboards[j]]);
                } else {
                    rows[rows.length - 1].push(_artboards[j]);
                }
            }

            /* (3) 各行内を left 昇順で厳密ソート */
            for (var r = 0; r < rows.length; r++) {
                rows[r].sort(function (a, b) {
                    return a.artboardRect[0] - b.artboardRect[0];
                });
            }

            return rows;
        }

        /* 配列を「行→列」順に in-place で並べ替える（rebuildArtboardsInSortedOrder 用） */
        function sortArtboardsTopLeftWithTolerance(_artboards, decimalPlaces, tolerance) {
            var rows = groupArtboardsIntoRows(_artboards, decimalPlaces, tolerance);
            var idx = 0;
            for (var r = 0; r < rows.length; r++) {
                for (var c = 0; c < rows[r].length; c++) {
                    _artboards[idx++] = rows[r][c];
                }
            }
        }

        /* 数値をゼロ埋め / Zero-pad a number */
        function padNumber(num, width) {
            var s = String(num);
            while (s.length < width) s = "0" + s;
            return s;
        }

        /* 行-列形式の名前を組み立てる / Format a row-column style name */
        function formatRowColName(rowNum, colNum, separator, padWidth) {
            return padNumber(rowNum, padWidth) + separator + padNumber(colNum, padWidth);
        }

        /* 物理的な配置から 行-列 名を割り当てる / Assign row-column names based on physical positions */
        function renameArtboardsFromPositions(doc, separator, padWidth, tolerance) {
            var dp = Math.pow(10, COORDINATE_PRECISION_DIGITS);

            var entries = [];
            for (var i = 0; i < doc.artboards.length; i++) {
                entries.push({
                    index: i,
                    artboardRect: doc.artboards[i].artboardRect
                });
            }

            var rows = groupArtboardsIntoRows(entries, dp, tolerance);

            for (var r = 0; r < rows.length; r++) {
                for (var c = 0; c < rows[r].length; c++) {
                    doc.artboards[rows[r][c].index].name = formatRowColName(r + 1, c + 1, separator, padWidth);
                }
            }
        }

        /* 既存名を再フォーマット / Reformat existing names */
        /* - 行-列 形式（例: 1-2）: 区切り文字とゼロ埋めを反映 */
        /* - それ以外（例: A-1, Item5）: 数字部分のみゼロ埋めを反映 */
        function renameArtboardsFromExistingNames(doc, separator, padWidth) {
            for (var i = 0; i < doc.artboards.length; i++) {
                var artboard = doc.artboards[i];
                var name = artboard.name;
                var match = name.match(/^(\d+)[-_x](\d+)(.*)$/i);
                if (match) {
                    var row = parseInt(match[1], 10);
                    var col = parseInt(match[2], 10);
                    var rest = match[3] || "";
                    artboard.name = formatRowColName(row, col, separator, padWidth) + rest;
                } else {
                    artboard.name = name.replace(/\d+/g, function (digits) {
                        return padNumber(parseInt(digits, 10), padWidth);
                    });
                }
            }
        }

        /* 許容差自動計算関数 / Auto calculate tolerance function */
        function calculateAutoTolerance(_artboards) {
            var tops = [];
            for (var i = 0; i < _artboards.length; i++) {
                tops.push(_artboards[i].artboardRect[1]);
            }
            tops.sort(function (a, b) {
                return b - a;
            }); /* 上から下に並べる / Sort top to bottom */

            var diffs = [];
            for (var i = 1; i < tops.length; i++) {
                var diff = Math.abs(tops[i] - tops[i - 1]);
                if (diff > 0) {
                    diffs.push(diff);
                }
            }

            if (diffs.length === 0) {
                return 5; /* 差が無ければデフォルト / Default if no difference */
            }

            var minDiff = Math.min.apply(null, diffs);
            return minDiff + 2; /* 少しマージンを加える / Add some margin */
        }
        main();
    })();