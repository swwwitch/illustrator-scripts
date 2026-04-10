#target illustrator

    /*
    概要:
    ・選択オブジェクトの visibleBounds を基準に、ロゴ用の補助線およびクリアスペースを生成するスクリプトです（選択範囲の境界に一致する線を太く強調表示する機能に対応）。
    ・本スクリプトは「文字形状からグリッド構造を抽出する」ことを目的としています。
    
    【横線】
    ・横伸張率、上下に線を追加、左方向に延長を設定できます。
    ・ラインの決め方（なし／自動判定／水平セグメント／均等分割）に対応しています。
    
    【縦線】
    ・縦伸張率、左右に線を追加、縦分割、上方向に延長を設定できます。
    ・垂直エレメント／斜線エレメントの抽出に対応しています。
    
    【オプション（縦線）】
    ・縦分割数、垂直エレメント、斜線エレメントを制御できます。
    
    【共通設定とクリアスペース】 
    ・作成レイヤー名、線幅、ガイド化、グループ化を設定できます。
    ・クリアスペースON時は、横線・縦線はパネルごと無効化されます。
    ・このとき、線幅・ガイド化は無効、グループ化はON固定になります。
    ・境界線を強調：境界に一致する線の線幅を強調（4倍）します。
    
    【プリセット】
    ・プリセットは配列定義から自動生成されます。
    ・以下のプリセットを選択できます：
        - 1x1
        - auto
        - element
        - left
        - up-3
        - clear space
    ・現在のUI状態をJSON（配列貼り付け形式）として書き出すことができます。
    
    【動作】
    ・左方向に延長ON：右側基準で横線を左に延長
    ・上方向に延長ON：下側基準で縦線を上に延長
    
    【出力】
    ・補助線およびクリアスペースは指定レイヤーに作成（未存在時は自動生成）
    ・プレビュー対応
    ・ガイド変換対応
    ・単位はIllustratorのstrokeUnitsに準拠
    ・日英ローカライズ対応
    
    更新日: 2026-04-10
    */

    (function () {
        // =========================================
        // バージョンとローカライズ
        // =========================================

        var SCRIPT_VERSION = "v1.4.0";

        function getCurrentLang() {
            return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
        }
        var lang = getCurrentLang();

        /* 日英ラベル定義 / Japanese-English label definitions */
        var LABELS = {
            dialogTitle: {
                ja: "ロゴグリッドの作成",
                en: "Logo Grid Maker"
            },
            labelPreset: {
                ja: "プリセット:",
                en: "Preset:"
            },
            panelClearSpace: {
                ja: "共通設定とクリアスペース",
                en: "Common Settings & Clear Space"
            },
            preset1x1: {
                ja: "1x1",
                en: "1x1"
            },
            presetAuto: {
                ja: "auto",
                en: "auto"
            },
            presetElement: {
                ja: "element",
                en: "element"
            },
            presetLeft: {
                ja: "left",
                en: "left"
            },
            presetUp3: {
                ja: "up-3",
                en: "up-3"
            },
            presetClearSpace: {
                ja: "clear space",
                en: "clear space"
            },
            errNoDocument: {
                ja: "ドキュメントが開かれていません。",
                en: "No document is open."
            },
            errNoSelection: {
                ja: "オブジェクトを選択してください。",
                en: "Please select an object."
            },
            errNoBounds: {
                ja: "選択範囲の境界を取得できませんでした。",
                en: "Could not get the bounds of the selection."
            },
            errInvalidSize: {
                ja: "選択オブジェクトのサイズを取得できませんでした。",
                en: "Could not get the size of the selected object."
            },
            errInvalidInput: {
                ja: "入力値が不正です。数値を正しく入力してください。",
                en: "Invalid input. Please enter valid numeric values."
            },
            defaultGuideLayerName: {
                ja: "// logo_guide",
                en: "// logo_guide"
            },
            labelGuideLayer: {
                ja: "作成レイヤー",
                en: "Target Layer"
            },
            panelHorizontal: {
                ja: "横線",
                en: "Horizontal Lines"
            },
            panelVertical: {
                ja: "縦線",
                en: "Vertical Lines"
            },
            labelWidthScale: {
                ja: "横伸張率",
                en: "Width Scale"
            },
            labelHeightScale: {
                ja: "縦伸張率",
                en: "Height Scale"
            },
            labelStrokeWidth: {
                ja: "線幅",
                en: "Stroke Width"
            },
            panelLineMethod: {
                ja: "ラインの決め方",
                en: "Line Generation Method"
            },
            radioLineNone: {
                ja: "なし",
                en: "None"
            },
            radioLineAuto: {
                ja: "自動判定",
                en: "Auto Detect"
            },
            radioLineSegment: {
                ja: "水平セグメント",
                en: "Horizontal Segments"
            },
            radioLineEvenDiv: {
                ja: "均等分割",
                en: "Even Divisions"
            },
            unitLines: {
                ja: "本",
                en: "lines"
            },
            checkHorizontalExtra: {
                ja: "上下に線を追加",
                en: "Add Top & Bottom Lines"
            },
            checkHorizontalLeftExtra: {
                ja: "左方向に延長",
                en: "Extend Leftward"
            },
            checkVerticalOuter: {
                ja: "左右に線を追加",
                en: "Add Left & Right Lines"
            },
            checkVerticalDiv: {
                ja: "均等分割",
                en: "Even Divisions"
            },
            checkVerticalTopSpace: {
                ja: "上方向に延長",
                en: "Extend Upward"
            },
            checkVerticalElements: {
                ja: "垂直エレメント",
                en: "Vertical Elements"
            },
            checkDiagonalElements: {
                ja: "斜線エレメント",
                en: "Diagonal Elements"
            },
            checkGuides: {
                ja: "ガイドに変換する",
                en: "Convert to Guides"
            },
            checkWeighting: {
                ja: "境界線を強調",
                en: "Highlight Bounds"
            },
            panelOptions: {
                ja: "オプション",
                en: "Options"
            },
            checkIsolationArea: {
                ja: "クリアスペース",
                en: "Clear Space"
            },
            labelIsolationDiv: {
                ja: "分割数",
                en: "Divisions"
            },
            checkGroupItems: {
                ja: "グループ化",
                en: "Group Items"
            },
            checkPreview: {
                ja: "プレビュー",
                en: "Preview"
            },
            btnExport: {
                ja: "書き出し",
                en: "Export"
            },
            btnCancel: {
                ja: "キャンセル",
                en: "Cancel"
            },
            btnOk: {
                ja: "OK",
                en: "OK"
            },
            layerGuideLines: {
                ja: "補助線",
                en: "Guides"
            },
            layerGuideLinesPreview: {
                ja: "補助線_プレビュー",
                en: "Guides (Preview)"
            },
            layerIsolationArea: {
                ja: "クリアスペース",
                en: "Clear Space"
            }
        };

        function L(key) {
            var entry = LABELS[key];
            return entry ? entry[lang] : key;
        }

        /* Live Corner Annotator 実行 / Run Live Corner Annotator */
        function runLiveCornerAnnotator() {
            try {
                if (typeof app === "undefined" || !app || typeof app.executeMenuCommand !== "function") {
                    return;
                }
                app.executeMenuCommand('Live Corner Annotator');
            } catch (e) { }
        }

        // =========================================
        // 単位ユーティリティ
        // =========================================

        /* 単位コードマップ / Unit code map */
        var unitMap = {
            0: "in",
            1: "mm",
            2: "pt",
            3: "pica",
            4: "cm",
            6: "px",
            7: "ft/in",
            8: "m",
            9: "yd",
            10: "ft"
        };

        /* 単位コードと設定キーからラベル取得 / Get unit label from code and preference key */
        function getUnitLabel(code, prefKey) {
            if (code === 5) {
                var hKeys = {
                    "text/asianunits": true,
                    "rulerType": true,
                    "strokeUnits": true
                };
                return hKeys[prefKey] ? "H" : "Q";
            }
            return unitMap[code] || "pt";
        }

        /* 単位コードをpt係数へ変換 / Convert unit code to pt factor */
        function getUnitFactor(code, prefKey) {
            switch (code) {
                case 0: return 72; /* in */
                case 1: return 72 / 25.4; /* mm */
                case 2: return 1; /* pt */
                case 3: return 12; /* pica */
                case 4: return 72 / 2.54; /* cm */
                case 5: return getUnitLabel(code, prefKey) === "H" ? (72 / 72) : (72 / 101.6); /* H or Q */
                case 6: return 1; /* px (Illustrator points base) */
                case 7: return 1; /* ft/in は複合入力未対応のため安全側でpt扱い / Treat as pt because compound ft/in input is unsupported */
                case 8: return 72 / 0.0254; /* m */
                case 9: return 72 * 36; /* yd */
                case 10: return 72 * 12; /* ft */
                default: return 1;
            }
        }

        /* 環境設定の単位コード取得 / Get preference unit code */
        function getPreferenceUnitCode(prefKey) {
            try {
                return app.preferences.getIntegerPreference(prefKey);
            } catch (e) {
                return 2;
            }
        }


        /* 現在の線単位ラベル取得 / Get current stroke unit label */
        function getCurrentStrokeUnitLabel() {
            var unitCode = getPreferenceUnitCode("strokeUnits");
            return getUnitLabel(unitCode, "strokeUnits");
        }

        /* 単位値をptへ変換 / Convert unit value to pt */
        function convertToPoints(value, prefKey) {
            var code = getPreferenceUnitCode(prefKey);
            return value * getUnitFactor(code, prefKey);
        }

        var doc;
        var sel;
        var previewGroup = null;
        main();

        // =========================================
        // 関数定義
        // =========================================

        function main() {
            try {
                runLiveCornerAnnotator();
            } catch (e) { }
            try {
                if (app.documents.length === 0) {
                    alert(L("errNoDocument"));
                    return;
                }

                doc = app.activeDocument;

                if (doc.selection.length === 0) {
                    alert(L("errNoSelection"));
                    return;
                }

                sel = doc.selection;

                previewGroup = null;
                var ui = createDialog();

                var context = buildContext(doc, sel, ui.guideLayerInput.text);
                if (!context) {
                    return;
                }

                bindArrowKeyHandlers(ui, context);
                bindEvents(ui, context);
                applyPreset(ui, context);
                updateVerticalTopSpaceState(ui);
                syncDivisionUIState(ui);
                updateGroupCheckEnabled(ui);

                ui.dialog.show();
            } finally {
                try {
                    runLiveCornerAnnotator();
                } catch (e) { }
            }
        }

        // =========================================
        // ライン解析 / Line analysis
        // =========================================

        function isStraightSegment(pt1, pt2) {
            var epsilon = 0.001;
            var rDx = Math.abs(pt1.rightDirection[0] - pt1.anchor[0]);
            var rDy = Math.abs(pt1.rightDirection[1] - pt1.anchor[1]);
            var lDx = Math.abs(pt2.leftDirection[0] - pt2.anchor[0]);
            var lDy = Math.abs(pt2.leftDirection[1] - pt2.anchor[1]);
            return (rDx < epsilon && rDy < epsilon && lDx < epsilon && lDy < epsilon);
        }

        function classifySegmentDirection(p1, p2) {
            var epsilon = 0.001;
            var dx = Math.abs(p1[0] - p2[0]);
            var dy = Math.abs(p1[1] - p2[1]);
            if (dy < epsilon) return "HORIZONTAL";
            if (dx < epsilon) return "VERTICAL";
            return "DIAGONAL";
        }

        function createLineAnalysisContext() {
            return {
                allPoints: [],
                horizSegments: [],
                vertSegments: [],
                minX: Infinity,
                maxX: -Infinity,
                globalMaxY: -Infinity,
                globalMinY: Infinity
            };
        }

        function collectGeometryFromItems(items, ctx) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (item.typename === "PathItem") {
                    collectGeometryFromPathItem(item, ctx);
                } else if (item.typename === "GroupItem") {
                    collectGeometryFromItems(item.pageItems, ctx);
                } else if (item.typename === "CompoundPathItem") {
                    collectGeometryFromItems(item.pathItems, ctx);
                }
            }
        }

        function collectGeometryFromPathItem(pathItem, ctx) {
            var pts = pathItem.pathPoints;
            var n = pts.length;
            for (var j = 0; j < n; j++) {
                var p1 = pts[j].anchor;
                var p2 = pts[(j + 1) % n].anchor;
                if (p1[1] > ctx.globalMaxY) ctx.globalMaxY = p1[1];
                if (p1[1] < ctx.globalMinY) ctx.globalMinY = p1[1];
                if (p1[0] < ctx.minX) ctx.minX = p1[0];
                if (p1[0] > ctx.maxX) ctx.maxX = p1[0];
                var isHoriz = (Math.abs(p1[1] - p2[1]) < 0.2 && Math.abs(p1[0] - p2[0]) > 0.5);
                ctx.allPoints.push({ x: p1[0], y: p1[1], isHoriz: isHoriz });

                var nextIdx = (j + 1) % n;
                if ((j < n - 1 || pathItem.closed) && isStraightSegment(pts[j], pts[nextIdx])) {
                    var segDir = classifySegmentDirection(p1, p2);
                    if (segDir === "HORIZONTAL") {
                        ctx.horizSegments.push({ y: p1[1], length: Math.abs(p1[0] - p2[0]) });
                    } else if (segDir === "VERTICAL") {
                        ctx.vertSegments.push({ x: p1[0], length: Math.abs(p1[1] - p2[1]) });
                    }
                }
            }
        }

        function analyzeAutoMode(ctx) {
            var approxHeight = ctx.globalMaxY - ctx.globalMinY;
            var tolerance = Math.max(1.0, Math.min(5.0, approxHeight * 0.03));
            var sorted = ctx.allPoints.slice().sort(function (a, b) { return b.y - a.y; });

            var clusters = [];
            var currentCluster = null;
            for (var i = 0; i < sorted.length; i++) {
                var pt = sorted[i];
                if (!currentCluster || Math.abs(currentCluster.baseY - pt.y) > tolerance) {
                    currentCluster = { baseY: pt.y, points: [] };
                    clusters.push(currentCluster);
                }
                currentCluster.points.push(pt);
            }

            var finalLines = [];
            for (var i = 0; i < clusters.length; i++) {
                var pts = clusters[i].points;
                var hPoints = [];
                var yFreq = {};
                for (var j = 0; j < pts.length; j++) {
                    var y = pts[j].y;
                    if (pts[j].isHoriz) hPoints.push(y);
                    var key = y.toFixed(3);
                    yFreq[key] = (yFreq[key] || 0) + 1;
                }
                var maxF = -1; var freqY = pts[0].y;
                for (var k in yFreq) { if (yFreq[k] > maxF) { maxF = yFreq[k]; freqY = parseFloat(k); } }
                finalLines.push({ y: freqY, count: pts.length, hasHoriz: hPoints.length > 0 });
            }

            var ascY = finalLines[0].y;
            var descY = finalLines[finalLines.length - 1].y;

            var candidates = [];
            for (var i = 1; i < finalLines.length - 1; i++) {
                var line = finalLines[i];
                if (Math.abs(ascY - line.y) > tolerance && Math.abs(line.y - descY) > tolerance) {
                    candidates.push(line);
                }
            }

            var midY = (ascY + descY) / 2;
            var meanCandidates = [];
            var baseCandidates = [];
            for (var i = 0; i < candidates.length; i++) {
                if (candidates[i].y > midY) {
                    meanCandidates.push(candidates[i]);
                } else {
                    baseCandidates.push(candidates[i]);
                }
            }

            function bestByScore(arr) {
                arr.sort(function (a, b) {
                    var scoreA = a.count + (a.hasHoriz ? 1000 : 0);
                    var scoreB = b.count + (b.hasHoriz ? 1000 : 0);
                    return scoreB - scoreA;
                });
                return arr[0];
            }

            var meanY, baseY;
            if (meanCandidates.length >= 1 && baseCandidates.length >= 1) {
                meanY = bestByScore(meanCandidates).y;
                baseY = bestByScore(baseCandidates).y;
            } else {
                meanY = ascY - (ascY - descY) * 0.3;
                baseY = descY + (ascY - descY) * 0.3;
            }

            var totalHeight = ascY - descY;
            var hasDescender = true;
            if (totalHeight > 0) {
                var belowBase = baseY - descY;
                if (belowBase / totalHeight < 0.15) {
                    hasDescender = false;
                }
            }

            return { ascY: ascY, descY: descY, meanY: meanY, baseY: baseY, hasDescender: hasDescender };
        }

        function analyzeHorizontalSegmentsMode(ctx) {
            var approxHeight = ctx.globalMaxY - ctx.globalMinY;
            var tolerance = Math.max(1.0, Math.min(5.0, approxHeight * 0.03));
            var sorted = ctx.horizSegments.slice().sort(function (a, b) { return b.y - a.y; });

            var clusters = [];
            var currentCluster = null;
            for (var i = 0; i < sorted.length; i++) {
                var seg = sorted[i];
                if (!currentCluster || Math.abs(currentCluster.baseY - seg.y) > tolerance) {
                    currentCluster = { baseY: seg.y, segs: [] };
                    clusters.push(currentCluster);
                }
                currentCluster.segs.push(seg);
            }

            var scored = [];
            for (var i = 0; i < clusters.length; i++) {
                var segs = clusters[i].segs;
                var totalLen = 0;
                var weightedY = 0;
                for (var j = 0; j < segs.length; j++) {
                    totalLen += segs[j].length;
                    weightedY += segs[j].y * segs[j].length;
                }
                scored.push({ y: weightedY / totalLen });
            }

            scored.sort(function (a, b) { return b.y - a.y; });
            var lines = [];
            for (var i = 0; i < scored.length; i++) {
                lines.push(scored[i].y);
            }
            return lines;
        }

        function getLineYPositions(sel, lineMethod, lineEvenDivCount) {
            var ctx = createLineAnalysisContext();
            collectGeometryFromItems(sel, ctx);
            if (ctx.allPoints.length === 0) return [];

            if (lineMethod === "none") {
                return [];
            } else if (lineMethod === "segment") {
                return analyzeHorizontalSegmentsMode(ctx);
            } else if (lineMethod === "even") {
                var top = ctx.globalMaxY;
                var bottom = ctx.globalMinY;
                var lines = [top];
                var step = (top - bottom) / (lineEvenDivCount + 1);
                for (var i = 1; i <= lineEvenDivCount; i++) {
                    lines.push(top - step * i);
                }
                lines.push(bottom);
                return lines;
            } else {
                /* auto */
                var r = analyzeAutoMode(ctx);
                var lines = [r.ascY, r.meanY];
                if (r.hasDescender) lines.push(r.baseY);
                lines.push(r.descY);
                return lines;
            }
        }

        function collectVerticalSegments(items, result) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (item.typename === "PathItem") {
                    collectVerticalSegmentsFromPath(item, result);
                } else if (item.typename === "GroupItem") {
                    collectVerticalSegments(item.pageItems, result);
                } else if (item.typename === "CompoundPathItem") {
                    collectVerticalSegments(item.pathItems, result);
                }
            }
        }

        function collectVerticalSegmentsFromPath(pathItem, result) {
            var pts = pathItem.pathPoints;
            var n = pts.length;
            for (var j = 0; j < n; j++) {
                var nextIdx = (j + 1) % n;
                if (j >= n - 1 && !pathItem.closed) continue;
                if (!isStraightSegment(pts[j], pts[nextIdx])) continue;
                var p1 = pts[j].anchor;
                var p2 = pts[nextIdx].anchor;
                if (classifySegmentDirection(p1, p2) !== "VERTICAL") continue;

                result.push({ x: p1[0], length: Math.abs(p1[1] - p2[1]) });
            }
        }

        function getVerticalLineXPositions(sel) {
            var vertSegs = [];
            collectVerticalSegments(sel, vertSegs);
            if (vertSegs.length === 0) return [];

            var minX = Infinity, maxX = -Infinity;
            var totalLen = 0;
            for (var i = 0; i < vertSegs.length; i++) {
                if (vertSegs[i].x < minX) minX = vertSegs[i].x;
                if (vertSegs[i].x > maxX) maxX = vertSegs[i].x;
                totalLen += vertSegs[i].length;
            }

            var minAcceptedLength = Math.max(2, totalLen * 0.03 / vertSegs.length);
            var filteredSegs = [];
            for (var i = 0; i < vertSegs.length; i++) {
                if (vertSegs[i].length >= minAcceptedLength) {
                    filteredSegs.push(vertSegs[i]);
                }
            }
            if (filteredSegs.length === 0) return [];

            var approxWidth = maxX - minX;
            var tolerance = Math.max(1.0, Math.min(5.0, approxWidth * 0.03));
            var sorted = filteredSegs.slice().sort(function (a, b) { return a.x - b.x; });

            var clusters = [];
            var currentCluster = null;
            for (var i = 0; i < sorted.length; i++) {
                var seg = sorted[i];
                if (!currentCluster || Math.abs(currentCluster.baseX - seg.x) > tolerance) {
                    currentCluster = { baseX: seg.x, segs: [] };
                    clusters.push(currentCluster);
                }
                currentCluster.segs.push(seg);
            }

            var scored = [];
            for (var i = 0; i < clusters.length; i++) {
                var segs = clusters[i].segs;
                var totalLen = 0;
                var weightedX = 0;
                for (var j = 0; j < segs.length; j++) {
                    totalLen += segs[j].length;
                    weightedX += segs[j].x * segs[j].length;
                }
                scored.push({ x: weightedX / totalLen });
            }

            scored.sort(function (a, b) { return a.x - b.x; });
            var lines = [];
            for (var i = 0; i < scored.length; i++) {
                lines.push(scored[i].x);
            }
            return lines;
        }

        /* 斜線エレメント用ヘルパー / Diagonal elements helpers */

        function collectDiagonalSegments(items, result) {
            for (var i = 0; i < items.length; i++) {
                var item = items[i];
                if (item.typename === "PathItem") {
                    collectDiagonalSegmentsFromPath(item, result);
                } else if (item.typename === "GroupItem") {
                    collectDiagonalSegments(item.pageItems, result);
                } else if (item.typename === "CompoundPathItem") {
                    collectDiagonalSegments(item.pathItems, result);
                }
            }
        }

        function collectDiagonalSegmentsFromPath(pathItem, result) {
            var pts = pathItem.pathPoints;
            var n = pts.length;
            for (var j = 0; j < n; j++) {
                var nextIdx = (j + 1) % n;
                if (j >= n - 1 && !pathItem.closed) continue;
                if (!isStraightSegment(pts[j], pts[nextIdx])) continue;
                var p1 = pts[j].anchor;
                var p2 = pts[nextIdx].anchor;
                if (classifySegmentDirection(p1, p2) !== "DIAGONAL") continue;

                var dx = p2[0] - p1[0];
                var dy = p2[1] - p1[1];
                var angle = Math.atan2(dy, dx) * 180 / Math.PI;
                if (angle < 0) angle += 180;
                if (angle >= 180) angle -= 180;

                /* 斜体の傾きとして使いやすい右上がりの斜線のみ対象 */
                if (angle <= 45 || angle >= 90) continue;

                var length = Math.sqrt(dx * dx + dy * dy);
                result.push({ p1: p1, p2: p2, angle: angle, length: length });
            }
        }

        function extendLineToBounds(p1, p2, left, top, right, bottom) {
            var x1 = p1[0], y1 = p1[1];
            var x2 = p2[0], y2 = p2[1];
            var intersections = [];
            var epsilon = 0.001;

            if (Math.abs(x1 - x2) < epsilon) {
                intersections.push([x1, top]);
                intersections.push([x1, bottom]);
            } else if (Math.abs(y1 - y2) < epsilon) {
                intersections.push([left, y1]);
                intersections.push([right, y1]);
            } else {
                var m = (y2 - y1) / (x2 - x1);
                var b = y1 - m * x1;

                var yAtLeft = m * left + b;
                if (yAtLeft <= top + epsilon && yAtLeft >= bottom - epsilon)
                    intersections.push([left, yAtLeft]);

                var yAtRight = m * right + b;
                if (yAtRight <= top + epsilon && yAtRight >= bottom - epsilon)
                    intersections.push([right, yAtRight]);

                var xAtTop = (top - b) / m;
                if (xAtTop >= left - epsilon && xAtTop <= right + epsilon)
                    intersections.push([xAtTop, top]);

                var xAtBottom = (bottom - b) / m;
                if (xAtBottom >= left - epsilon && xAtBottom <= right + epsilon)
                    intersections.push([xAtBottom, bottom]);
            }

            if (intersections.length >= 2) {
                var uniquePoints = [intersections[0]];
                for (var j = 1; j < intersections.length; j++) {
                    var ddx = intersections[j][0] - uniquePoints[0][0];
                    var ddy = intersections[j][1] - uniquePoints[0][1];
                    if (Math.sqrt(ddx * ddx + ddy * ddy) > epsilon) {
                        uniquePoints.push(intersections[j]);
                        break;
                    }
                }
                if (uniquePoints.length === 2) return uniquePoints;
            }
            return null;
        }

        function extendPointAtAngle(point, angleRad, left, top, right, bottom) {
            var far = 100000;
            var p2 = [point[0] + Math.cos(angleRad) * far, point[1] + Math.sin(angleRad) * far];
            return extendLineToBounds(point, p2, left, top, right, bottom);
        }

        function makeDiagLineKey(endpoints) {
            var a = endpoints[0];
            var b = endpoints[1];
            var ka = a[0].toFixed(1) + "," + a[1].toFixed(1);
            var kb = b[0].toFixed(1) + "," + b[1].toFixed(1);
            return (ka < kb) ? (ka + "|" + kb) : (kb + "|" + ka);
        }

        function getDiagonalElementLines(sel, left, top, right, bottom) {
            var diagSegs = [];
            collectDiagonalSegments(sel, diagSegs);
            if (diagSegs.length === 0) return [];

            var angleTolerance = 2.0;
            diagSegs.sort(function (a, b) { return a.angle - b.angle; });

            var clusters = [];
            var cur = null;
            for (var i = 0; i < diagSegs.length; i++) {
                var seg = diagSegs[i];
                if (!cur || Math.abs(cur.baseAngle - seg.angle) > angleTolerance) {
                    cur = { baseAngle: seg.angle, segs: [], totalLength: 0 };
                    clusters.push(cur);
                }
                cur.segs.push(seg);
                cur.totalLength += seg.length;
            }

            /* 0°/180° 付近の角度の巻き戻し処理 */
            if (clusters.length >= 2) {
                var first = clusters[0];
                var last = clusters[clusters.length - 1];
                if ((180 - last.baseAngle + first.baseAngle) < angleTolerance) {
                    for (var mi = 0; mi < last.segs.length; mi++) {
                        first.segs.push(last.segs[mi]);
                    }
                    first.totalLength += last.totalLength;
                    clusters.pop();
                }
            }

            /* 最頻出角度のクラスタを特定 */
            var best = clusters[0];
            for (var i = 1; i < clusters.length; i++) {
                if (clusters[i].totalLength > best.totalLength) best = clusters[i];
            }

            /* 加重平均角度 */
            var totalLen = 0;
            var weightedAngle = 0;
            for (var i = 0; i < best.segs.length; i++) {
                totalLen += best.segs[i].length;
                weightedAngle += best.segs[i].angle * best.segs[i].length;
            }
            var bestAngle = weightedAngle / totalLen;
            var bestRad = bestAngle * Math.PI / 180;

            var lines = [];
            var drawnKeys = {};
            var candidateMids = [];

            var minAcceptedLength = Math.max(2, totalLen * 0.03 / best.segs.length);
            for (var i = 0; i < best.segs.length; i++) {
                var seg = best.segs[i];
                if (seg.length < minAcceptedLength) {
                    continue;
                }

                /* 連続する2つのアンカーポイント（= 1本の斜線セグメント）の中点を対象にする */
                candidateMids.push({
                    point: [
                        (seg.p1[0] + seg.p2[0]) / 2,
                        (seg.p1[1] + seg.p2[1]) / 2
                    ],
                    length: seg.length
                });
            }

            if (candidateMids.length === 0) {
                return [];
            }

            var nx = -Math.sin(bestRad);
            var ny = Math.cos(bestRad);
            for (var i = 0; i < candidateMids.length; i++) {
                candidateMids[i].offset = candidateMids[i].point[0] * nx + candidateMids[i].point[1] * ny;
            }
            candidateMids.sort(function (a, b) { return a.offset - b.offset; });

            var mergeTolerance = Math.max(2, minAcceptedLength * 0.6);
            var mergedMids = [];
            var current = null;

            for (var i = 0; i < candidateMids.length; i++) {
                var item = candidateMids[i];
                if (!current || Math.abs(item.offset - current.baseOffset) > mergeTolerance) {
                    current = {
                        baseOffset: item.offset,
                        sumX: item.point[0] * item.length,
                        sumY: item.point[1] * item.length,
                        totalWeight: item.length
                    };
                    mergedMids.push(current);
                } else {
                    current.sumX += item.point[0] * item.length;
                    current.sumY += item.point[1] * item.length;
                    current.totalWeight += item.length;
                }
            }

            for (var i = 0; i < mergedMids.length; i++) {
                var mergedMid = [
                    mergedMids[i].sumX / mergedMids[i].totalWeight,
                    mergedMids[i].sumY / mergedMids[i].totalWeight
                ];

                /* 斜体の傾きに沿う、垂直に近い線のみ描画 */
                var diagLine = extendPointAtAngle(mergedMid, bestRad, left, top, right, bottom);
                if (diagLine) {
                    var key = makeDiagLineKey(diagLine);
                    if (!drawnKeys[key]) {
                        lines.push(diagLine);
                        drawnKeys[key] = true;
                    }
                }
            }

            return lines;
        }

        function buildContext(doc, sel, layerName) {
            var bounds = getSelectionBounds(sel);
            if (!bounds) {
                alert(L("errNoBounds"));
                return null;
            }

            var bLeft = bounds[0];
            var bTop = bounds[1];
            var bRight = bounds[2];
            var bBottom = bounds[3];

            var A = bTop - bBottom;
            var B = bRight - bLeft;

            if (A <= 0 || B <= 0) {
                alert(L("errInvalidSize"));
                return null;
            }

            var centerX = (bLeft + bRight) / 2;
            var centerY = (bTop + bBottom) / 2;

            return {
                doc: doc,
                bLeft: bLeft,
                bTop: bTop,
                bRight: bRight,
                bBottom: bBottom,
                A: A,
                B: B,
                centerX: centerX,
                centerY: centerY,
                guideLayer: getOrCreateLayer(doc, layerName)
            };
        }
        function findLayerByName(doc, name) {
            for (var i = 0; i < doc.layers.length; i++) {
                if (doc.layers[i].name === name) {
                    return doc.layers[i];
                }
            }
            return null;
        }

        function getOrCreateLayer(doc, name) {
            var layer = findLayerByName(doc, name);
            if (layer) {
                return layer;
            }
            layer = doc.layers.add();
            layer.name = name;
            return layer;
        }

        function setUiValueAndNotify(editText, value, ui, context, suppressPreview) {
            if (editText === ui.hScaleInput || editText === ui.vScaleInput) {
                editText.text = (Number(value) * 100).toFixed(1);
            } else {
                editText.text = String(value);
            }
            if (editText === ui.vScaleInput) {
                updateVerticalTopSpaceState(ui);
            }
            if (!suppressPreview && ui.previewCheck.value) {
                updatePreview(ui, context);
            }
        }

        function getPresetDefinitions() {
            return {
                "1x1": {
                    hScale: "140.0", hExtra: true, hExtraCount: "1", hLeftExtra: false,
                    lineMethod: "none", lineEvenDivCount: "4",
                    vScale: "200.0", vOuter: true, vOuterCount: "1", vTopSpace: false,
                    vDiv: false, vDivisions: "2", vElements: false, vDiagonalElements: false,
                    emphasizeBounds: false,
                    isolationArea: false, isolationAreaScale: "4"
                },
                "auto": {
                    hScale: "120.0", hExtra: true, hExtraCount: "2", hLeftExtra: false,
                    lineMethod: "auto", lineEvenDivCount: "4",
                    vScale: "200.0", vOuter: true, vOuterCount: "2", vTopSpace: false,
                    vDiv: false, vDivisions: "2", vElements: false, vDiagonalElements: false,
                    emphasizeBounds: true,
                    isolationArea: false, isolationAreaScale: "4"
                },
                "element": {
                    hScale: "120", hExtra: false, hExtraCount: "2", hLeftExtra: false,
                    lineMethod: "segment", lineEvenDivCount: "4",
                    vScale: "300", vOuter: false, vOuterCount: "2", vTopSpace: false,
                    vDiv: false, vDivisions: "3", vElements: true, vDiagonalElements: true,
                    emphasizeBounds: false,
                    isolationArea: false, isolationAreaScale: "5"
                },
                "element+": {
                    hScale: "120.0", hExtra: true, hExtraCount: "2", hLeftExtra: false,
                    lineMethod: "segment", lineEvenDivCount: "4",
                    vScale: "300.0", vOuter: true, vOuterCount: "2", vTopSpace: false,
                    vDiv: false, vDivisions: "3", vElements: true, vDiagonalElements: true,
                    emphasizeBounds: true,
                    isolationArea: false, isolationAreaScale: "5"
                },
                "left": {
                    hScale: "150.0", hExtra: true, hExtraCount: "2", hLeftExtra: true,
                    lineMethod: "even", lineEvenDivCount: "4",
                    vScale: "200.0", vOuter: true, vOuterCount: "2", vTopSpace: false,
                    vDiv: false, vDivisions: "2", vElements: false, vDiagonalElements: false,
                    emphasizeBounds: true,
                    isolationArea: false, isolationAreaScale: "5"
                },
                "up-3": {
                    hScale: "120", hExtra: true, hExtraCount: "2", hLeftExtra: false,
                    lineMethod: "even", lineEvenDivCount: "4",
                    vScale: "300", vOuter: true, vOuterCount: "2", vTopSpace: true,
                    vDiv: true, vDivisions: "3", vElements: false, vDiagonalElements: false,
                    emphasizeBounds: true,
                    isolationArea: false, isolationAreaScale: "5"
                },
                "clear space": {
                    hScale: "140.0", hExtra: true, hExtraCount: "1", hLeftExtra: false,
                    lineMethod: "none", lineEvenDivCount: "4",
                    vScale: "200.0", vOuter: true, vOuterCount: "1", vTopSpace: false,
                    vDiv: false, vDivisions: "2", vElements: false, vDiagonalElements: false,
                    emphasizeBounds: false,
                    isolationArea: true, isolationAreaScale: "4"
                }
            };
        }

        function exportPreset(ui) {
            var nameInput = prompt(lang === "ja" ? "プリセット名を入力してください:" : "Enter preset name:", "");
            if (!nameInput) return;

            var data = {
                name: nameInput,
                hScale: ui.hScaleInput.text,
                hExtra: ui.hExtraCheck.value,
                hExtraCount: ui.hExtraInput.text,
                hLeftExtra: ui.hLeftExtraCheck.value,
                lineMethod: ui.lineNoneRadio.value ? "none" : (ui.lineEvenDivRadio.value ? "even" : (ui.lineSegmentRadio.value ? "segment" : "auto")),
                lineEvenDivCount: ui.lineEvenDivInput.text,
                vScale: ui.vScaleInput.text,
                vOuter: ui.vOuterCheck.value,
                vOuterCount: ui.vOuterInput.text,
                vTopSpace: ui.vTopSpaceCheck.value,
                vDiv: ui.vDivCheck.value,
                vDivisions: ui.vDivInput.text,
                vElements: ui.vElementsCheck.value,
                vDiagonalElements: ui.vDiagonalCheck.value,
                emphasizeBounds: ui.weightingCheck.value,
                isolationArea: ui.isolationAreaCheck.value,
                isolationAreaScale: ui.isolationAreaInput.text
            };

            var lines = [];
            var row = function (k, v, isLast) {
                var val = (typeof v === "string") ? '"' + v + '"' : String(v);
                return k + ": " + val + (isLast ? "" : ", ");
            };

            lines.push('                "' + nameInput + '": {');
            lines.push("                    " +
                row("hScale", data.hScale) + row("hExtra", data.hExtra) +
                row("hExtraCount", data.hExtraCount) + row("hLeftExtra", data.hLeftExtra, true) + ",");
            lines.push("                    " +
                row("lineMethod", data.lineMethod) + row("lineEvenDivCount", data.lineEvenDivCount, true) + ",");
            lines.push("                    " +
                row("vScale", data.vScale) + row("vOuter", data.vOuter) +
                row("vOuterCount", data.vOuterCount) + row("vTopSpace", data.vTopSpace, true) + ",");
            lines.push("                    " +
                row("vDiv", data.vDiv) + row("vDivisions", data.vDivisions) +
                row("vElements", data.vElements) + row("vDiagonalElements", data.vDiagonalElements, true) + ",");
            lines.push("                    " +
                row("emphasizeBounds", data.emphasizeBounds, true) + ",");
            lines.push("                    " +
                row("isolationArea", data.isolationArea) + row("isolationAreaScale", data.isolationAreaScale, true));
            lines.push("                },");

            var json = lines.join("\n");

            var desktop = Folder.desktop;
            var fileName = nameInput.replace(new RegExp('[\\\\/:*?"<>|]', "g"), "_") + ".json";
            var file = new File(desktop + "/" + fileName);
            file.encoding = "UTF-8";
            if (file.open("w")) {
                file.write(json);
                file.close();
                alert((lang === "ja" ? "書き出しました（配列に貼り付け用）:\n" : "Exported (for array paste):\n") + file.fsName);
            }
        }

        function applyPreset(ui, context) {
            if (!ui.presetDropdown.selection) return;

            var presets = getPresetDefinitions();
            var selectedKey = ui.presetDropdown.selection._key;
            var p = presets[selectedKey];
            if (!p) return;

            var suppressPreview = true;

            setUiValueAndNotify(ui.hScaleInput, parseFloat(p.hScale) / 100, ui, context, suppressPreview);
            ui.hExtraCheck.value = p.hExtra;
            setUiValueAndNotify(ui.hExtraInput, p.hExtraCount, ui, context, suppressPreview);
            ui.hLeftExtraCheck.value = p.hLeftExtra;

            setUiValueAndNotify(ui.vScaleInput, parseFloat(p.vScale) / 100, ui, context, suppressPreview);
            ui.vOuterCheck.value = p.vOuter;
            setUiValueAndNotify(ui.vOuterInput, p.vOuterCount, ui, context, suppressPreview);
            ui.vTopSpaceCheck.value = p.vTopSpace;

            ui.vDivCheck.value = p.vDiv;
            setUiValueAndNotify(ui.vDivInput, p.vDivisions, ui, context, suppressPreview);

            ui.vElementsCheck.value = p.vElements;
            ui.vDiagonalCheck.value = p.vDiagonalElements;

            ui.weightingCheck.value = !!p.emphasizeBounds;

            ui.isolationAreaCheck.value = p.isolationArea;
            setUiValueAndNotify(ui.isolationAreaInput, p.isolationAreaScale, ui, context, suppressPreview);

            ui.lineNoneRadio.value = p.lineMethod === "none";
            ui.lineAutoRadio.value = p.lineMethod === "auto";
            ui.lineSegmentRadio.value = p.lineMethod === "segment";
            ui.lineEvenDivRadio.value = p.lineMethod === "even";
            setUiValueAndNotify(ui.lineEvenDivInput, p.lineEvenDivCount, ui, context, suppressPreview);

            updateVerticalTopSpaceState(ui);
            updateIsolationAreaEnabled(ui);
            syncDivisionUIState(ui);
            updateGroupCheckEnabled(ui);

            if (ui.previewCheck.value) {
                updatePreview(ui, context);
            }
        }

        function bindArrowKeyHandlers(ui, context) {
            changeValueByArrowKey(ui.hScaleInput, false, false, ui, context);
            changeValueByArrowKey(ui.vScaleInput, false, false, ui, context);
            changeValueByArrowKey(ui.vDivInput, false, true, ui, context);
            changeValueByArrowKey(ui.hExtraInput, false, true, ui, context);
            changeValueByArrowKey(ui.vOuterInput, false, true, ui, context);
            changeValueByArrowKey(ui.isolationAreaInput, false, false, ui, context);
            changeValueByArrowKey(ui.lineEvenDivInput, false, true, ui, context);
            changeValueByArrowKey(ui.swInput, false, false, ui, context);
        }
        function bindEvents(ui, context) {
            /* イベントハンドラ / Event handlers */
            ui.presetDropdown.onChange = function () {
                applyPreset(ui, context);
            };
            ui.exportBtn.onClick = function () {
                exportPreset(ui);
            };
            ui.previewCheck.onClick = function () {
                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                } else {
                    removePreview();
                }
            };

            ui.hScaleInput.onChanging = ui.swInput.onChanging =
                ui.vDivInput.onChanging = ui.hExtraInput.onChanging = ui.vOuterInput.onChanging = ui.isolationAreaInput.onChanging = function () {
                    if (ui.previewCheck.value) {
                        updatePreview(ui, context);
                    }
                };

            ui.lineEvenDivInput.onChanging = function () {
                syncDivisionUIState(ui);
                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                }
            };

            ui.vScaleInput.onChanging = function () {
                updateVerticalTopSpaceState(ui);
                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                }
            };

            ui.vScaleInput.onBlur = function () {
                var value = parseFloat(ui.vScaleInput.text);
                if (!isNaN(value) && isFinite(value)) {
                    ui.vScaleInput.text = value.toFixed(1);
                    updateVerticalTopSpaceState(ui);
                }
            };

            ui.guideLayerInput.onBlur = function () {
                ui.guideLayerInput.text = normalizeLayerNameText(ui.guideLayerInput.text);
            };

            ui.lineNoneRadio.onClick = ui.lineAutoRadio.onClick = ui.lineSegmentRadio.onClick = ui.lineEvenDivRadio.onClick = function () {
                ui.lineEvenDivInput.enabled = ui.lineEvenDivRadio.value;
                syncDivisionUIState(ui);
                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                }
            };

            ui.hExtraCheck.onClick = ui.hLeftExtraCheck.onClick =
                ui.vOuterCheck.onClick = ui.vDivCheck.onClick =
                ui.vTopSpaceCheck.onClick = ui.vElementsCheck.onClick =
                ui.vDiagonalCheck.onClick =
                ui.groupCheck.onClick = function () {
                    ui.hExtraInput.enabled = ui.hExtraCheck.enabled && ui.hExtraCheck.value;
                    ui.vOuterInput.enabled = ui.vOuterCheck.enabled && ui.vOuterCheck.value;
                    syncDivisionUIState(ui);
                    if (ui.previewCheck.value) {
                        updatePreview(ui, context);
                    }
                };

            ui.isolationAreaCheck.onClick = function () {
                updateIsolationAreaEnabled(ui);
                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                }
            };

            ui.weightingCheck.onClick = function () {
                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                }
            };

            ui.guideCheck.onClick = function () {
                updateGroupCheckEnabled(ui);
                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                }
            };

            ui.okBtn.onClick = function () {
                removePreview();
                var params = getParams(ui);
                if (!params) {
                    alert(L("errInvalidInput"));
                    return;
                }
                context.guideLayer = getOrCreateLayer(context.doc, params.guideLayerName);
                createLines(context, params, false);
                ui.dialog.close(1);
            };

            ui.cancelBtn.onClick = function () {
                removePreview();
                ui.dialog.close(2);
            };

            ui.dialog.onClose = function () {
                removePreview();
            };
        }

        function createDialog() {
            var dlg = new Window("dialog", L("dialogTitle") + " " + SCRIPT_VERSION);
            dlg.orientation = "column";
            dlg.alignChildren = ["fill", "top"];

            var presetWrap = dlg.add("group");
            presetWrap.alignment = ["center", "top"];

            var presetRow = presetWrap.add("group");
            presetRow.add("statictext", undefined, L("labelPreset"));
            var presetDropdown = presetRow.add("dropdownlist");

            function addPresetItem(valueKey) {
                var item = presetDropdown.add("item", valueKey);
                item._key = valueKey;
                return item;
            }

            var presetDefs = getPresetDefinitions();
            var autoIndex = 0;
            var presetIndex = 0;
            for (var presetKey in presetDefs) {
                if (!presetDefs.hasOwnProperty(presetKey)) continue;
                addPresetItem(presetKey);
                if (presetKey === "auto") {
                    autoIndex = presetIndex;
                }
                presetIndex++;
            }

            if (presetDropdown.items.length > 0) {
                presetDropdown.selection = autoIndex;
            }
            var exportBtn = presetRow.add("button", undefined, L("btnExport"));

            /* オプション / Options */
            var clearSpacePanel = dlg.add("panel", undefined, L("panelClearSpace"));
            clearSpacePanel.alignChildren = ["fill", "top"];
            clearSpacePanel.margins = [15, 20, 15, 10];

            var clearSpaceColumns = clearSpacePanel.add("group");
            clearSpaceColumns.orientation = "row";
            clearSpaceColumns.alignChildren = ["fill", "top"];

            var clearLeft = clearSpaceColumns.add("group");
            clearLeft.orientation = "column";
            clearLeft.alignChildren = ["left", "top"];

            var clearRight = clearSpaceColumns.add("group");
            clearRight.orientation = "column";
            clearRight.alignChildren = ["left", "top"];

            var clearSpaceRow = clearLeft.add("group");
            clearSpaceRow.alignChildren = ["left", "center"];
            clearSpaceRow.add("statictext", undefined, L("labelIsolationDiv"));
            var isolationAreaInput = clearSpaceRow.add("edittext", undefined, "4");
            isolationAreaInput.characters = 2;

            var csStrokeRow = clearRight.add("group");
            csStrokeRow.alignChildren = ["left", "center"];
            csStrokeRow.add("statictext", undefined, L("labelStrokeWidth"));
            var swInput = csStrokeRow.add("edittext", undefined, "0.25");
            swInput.characters = 4;
            csStrokeRow.add("statictext", undefined, getCurrentStrokeUnitLabel());

            var csWeightingRow = clearRight.add("group");
            csWeightingRow.alignChildren = ["left", "center"];
            var weightingCheck = csWeightingRow.add("checkbox", undefined, L("checkWeighting"));
            weightingCheck.value = false;

            var csIsolationCheckRow = clearLeft.add("group");
            csIsolationCheckRow.alignChildren = ["left", "center"];
            var isolationAreaCheck = csIsolationCheckRow.add("checkbox", undefined, L("checkIsolationArea"));
            isolationAreaCheck.value = false;

            var csGuideRow = clearLeft.add("group");
            csGuideRow.alignChildren = ["left", "center"];
            var guideCheck = csGuideRow.add("checkbox", undefined, L("checkGuides"));
            guideCheck.value = false;

            var csGroupRow = clearRight.add("group");
            csGroupRow.alignChildren = ["left", "center"];
            var groupCheck = csGroupRow.add("checkbox", undefined, L("checkGroupItems"));
            groupCheck.value = true;

            var guideLayerRow = clearSpacePanel.add("group");
            guideLayerRow.alignChildren = ["left", "center"];
            guideLayerRow.add("statictext", undefined, L("labelGuideLayer"));
            var guideLayerInput = guideLayerRow.add("edittext", undefined, L("defaultGuideLayerName"));
            guideLayerInput.characters = 16;


            var columns = dlg.add("group");
            columns.orientation = "row";
            columns.alignChildren = ["fill", "top"];

            var leftColumn = columns.add("group");
            leftColumn.orientation = "column";
            leftColumn.alignChildren = ["fill", "top"];

            var rightColumn = columns.add("group");
            rightColumn.orientation = "column";
            rightColumn.alignChildren = ["fill", "top"];

            /* 横線設定 / Horizontal line settings */
            var hPanel = leftColumn.add("panel", undefined, L("panelHorizontal"));
            hPanel.alignChildren = ["left", "center"];
            hPanel.margins = [15, 20, 15, 10];

            var hRow1 = hPanel.add("group");
            hRow1.add("statictext", undefined, L("labelWidthScale"));
            var hScaleInput = hRow1.add("edittext", undefined, "140");
            hScaleInput.characters = 6;
            hRow1.add("statictext", undefined, "%");

            var hRow3 = hPanel.add("group");
            var hExtraCheck = hRow3.add("checkbox", undefined, L("checkHorizontalExtra"));
            hExtraCheck.value = false;
            var hExtraInput = hRow3.add("edittext", undefined, "1");
            hExtraInput.characters = 2;

            var hRow4 = hPanel.add("group");
            var hLeftExtraCheck = hRow4.add("checkbox", undefined, L("checkHorizontalLeftExtra"));
            hLeftExtraCheck.value = false;

            /* ラインの決め方 / Line method */
            var lineMethodPanel = hPanel.add("panel", undefined, L("panelLineMethod"));
            lineMethodPanel.alignChildren = ["left", "top"];
            lineMethodPanel.margins = [15, 20, 15, 10];

            var lineNoneRadio = lineMethodPanel.add("radiobutton", undefined, L("radioLineNone"));
            var lineAutoRadio = lineMethodPanel.add("radiobutton", undefined, L("radioLineAuto"));
            lineAutoRadio.value = true;
            var lineSegmentRadio = lineMethodPanel.add("radiobutton", undefined, L("radioLineSegment"));
            var lineEvenDivGroup = lineMethodPanel.add("group");
            lineEvenDivGroup.alignChildren = ["left", "center"];
            var lineEvenDivRadio = lineEvenDivGroup.add("radiobutton", undefined, L("radioLineEvenDiv"));
            var lineEvenDivInput = lineEvenDivGroup.add("edittext", undefined, "4");
            lineEvenDivInput.characters = 2;
            lineEvenDivGroup.add("statictext", undefined, L("unitLines"));

            /* 縦線設定 / Vertical line settings */
            var vPanel = rightColumn.add("panel", undefined, L("panelVertical"));
            vPanel.alignChildren = ["left", "top"];
            vPanel.margins = [15, 20, 15, 10];

            var vRow1 = vPanel.add("group");
            vRow1.add("statictext", undefined, L("labelHeightScale"));
            var vScaleInput = vRow1.add("edittext", undefined, "200");
            vScaleInput.characters = 6;
            vRow1.add("statictext", undefined, "%");

            var vRow3 = vPanel.add("group");
            var vOuterCheck = vRow3.add("checkbox", undefined, L("checkVerticalOuter"));
            vOuterCheck.value = false;
            var vOuterInput = vRow3.add("edittext", undefined, "1");
            vOuterInput.characters = 2;

            var vRow4 = vPanel.add("group");
            var vTopSpaceCheck = vRow4.add("checkbox", undefined, L("checkVerticalTopSpace"));
            vTopSpaceCheck.value = false;
            vTopSpaceCheck.enabled = false;

            /* オプション / Options */
            var vOptionsPanel = vPanel.add("panel", undefined, L("panelOptions"));
            vOptionsPanel.alignChildren = ["left", "center"];
            vOptionsPanel.margins = [15, 20, 15, 10];

            var vRow2 = vOptionsPanel.add("group");
            var vDivCheck = vRow2.add("checkbox", undefined, L("checkVerticalDiv"));
            vDivCheck.value = false;
            var vDivInput = vRow2.add("edittext", undefined, "2");
            vDivInput.characters = 2;
            vDivInput.enabled = false;

            var vRow5 = vOptionsPanel.add("group");
            var vElementsCheck = vRow5.add("checkbox", undefined, L("checkVerticalElements"));
            vElementsCheck.value = false;

            var vRow6 = vOptionsPanel.add("group");
            var vDiagonalCheck = vRow6.add("checkbox", undefined, L("checkDiagonalElements"));
            vDiagonalCheck.value = false;

            /* 下部ボタンエリア / Bottom button area */
            var bottomRow = dlg.add("group");
            bottomRow.orientation = "row";
            bottomRow.alignChildren = ["fill", "center"];

            var leftButtons = bottomRow.add("group");
            leftButtons.orientation = "row";
            leftButtons.alignChildren = ["left", "center"];
            var previewCheck = leftButtons.add("checkbox", undefined, L("checkPreview"));
            previewCheck.value = false;

            var spacer = bottomRow.add("group");
            spacer.alignment = ["fill", "fill"];
            spacer.minimumSize.width = 0;

            var rightButtons = bottomRow.add("group");
            rightButtons.orientation = "row";
            rightButtons.alignChildren = ["right", "center"];
            var cancelBtn = rightButtons.add("button", undefined, L("btnCancel"), { name: "cancel" });
            var okBtn = rightButtons.add("button", undefined, L("btnOk"), { name: "ok" });

            return {
                dialog: dlg,
                presetDropdown: presetDropdown,
                exportBtn: exportBtn,
                hPanel: hPanel,
                vPanel: vPanel,
                hScaleInput: hScaleInput,
                hExtraCheck: hExtraCheck,
                hExtraInput: hExtraInput,
                hLeftExtraCheck: hLeftExtraCheck,
                lineNoneRadio: lineNoneRadio,
                lineAutoRadio: lineAutoRadio,
                lineSegmentRadio: lineSegmentRadio,
                lineEvenDivRadio: lineEvenDivRadio,
                lineEvenDivInput: lineEvenDivInput,
                vScaleInput: vScaleInput,
                vOuterCheck: vOuterCheck,
                vOuterInput: vOuterInput,
                vDivCheck: vDivCheck,
                vDivInput: vDivInput,
                vTopSpaceCheck: vTopSpaceCheck,
                vElementsCheck: vElementsCheck,
                vDiagonalCheck: vDiagonalCheck,
                isolationAreaCheck: isolationAreaCheck,
                isolationAreaInput: isolationAreaInput,
                guideLayerInput: guideLayerInput,
                swInput: swInput,
                guideCheck: guideCheck,
                groupCheck: groupCheck,
                weightingCheck: weightingCheck,
                previewCheck: previewCheck,
                cancelBtn: cancelBtn,
                okBtn: okBtn
            };
        }

        function updateGroupCheckEnabled(ui) {
            ui.groupCheck.enabled = !ui.guideCheck.value;
            if (ui.guideCheck.value) {
                ui.groupCheck.value = false;
            }
        }

        function updateIsolationAreaEnabled(ui) {
            var iso = ui.isolationAreaCheck.value;

            /* 横線・縦線は panel ごとディム */
            ui.hPanel.enabled = !iso;
            ui.vPanel.enabled = !iso;

            /* 横線パネル内の入力状態を補正 */
            ui.hExtraInput.enabled = !iso && ui.hExtraCheck.value;
            ui.lineEvenDivInput.enabled = !iso && ui.lineEvenDivRadio.value;

            /* 縦線パネル内の入力状態を補正 */
            ui.vOuterInput.enabled = !iso && ui.vOuterCheck.value;
            ui.vDivInput.enabled = !iso && ui.vDivCheck.value;

            /* 共通設定とクリアスペース: 線幅とガイド化もディム */
            ui.swInput.enabled = !iso;
            ui.guideCheck.enabled = !iso;
            if (iso) {
                ui.guideCheck.value = false;
                ui.groupCheck.value = true;
                ui.groupCheck.enabled = false;
            } else {
                updateGroupCheckEnabled(ui);
            }

            syncDivisionUIState(ui);
        }

        function updateVerticalTopSpaceState(ui) {
            ui.vTopSpaceCheck.enabled = true;
        }

        function updateVerticalDivisionEnabled(ui) {
            ui.vDivInput.enabled = ui.vDivCheck.enabled && ui.vDivCheck.value;
        }

        function syncDivisionUIState(ui) {
            updateVerticalDivisionEnabled(ui);
            updateIsolationDivisionInputEnabled(ui);
        }

        function updateIsolationDivisionInputEnabled(ui) {
            var evenMode = ui.lineEvenDivRadio.value;
            var evenCount = parseInt(ui.lineEvenDivInput.text, 10);

            if (evenMode && !isNaN(evenCount) && isFinite(evenCount) && evenCount >= 1) {
                ui.isolationAreaInput.text = String(evenCount + 1);
            }

            ui.isolationAreaInput.enabled = !evenMode;
        }

        function normalizeLayerNameText(text) {
            return String(text)
                .replace(/[\r\n\t]+/g, " ")
                .replace(/^\s+|\s+$/g, "");
        }

        function isStrictPositiveNumberText(text) {
            return /^\d+(?:\.\d+)?$/.test(text);
        }

        function readParams(ui) {
            return {
                guideLayerNameText: normalizeLayerNameText(ui.guideLayerInput.text),
                hScaleText: ui.hScaleInput.text,
                vScaleText: ui.vScaleInput.text,
                strokeWidthText: ui.swInput.text,
                hScale: parseFloat(ui.hScaleInput.text) / 100,
                vScale: parseFloat(ui.vScaleInput.text) / 100,
                strokeWidthValue: parseFloat(ui.swInput.text),
                vDivisionsText: ui.vDivInput.text,
                vDivisions: parseInt(ui.vDivInput.text, 10),
                lineMethod: ui.lineNoneRadio.value ? "none" : (ui.lineEvenDivRadio.value ? "even" : (ui.lineSegmentRadio.value ? "segment" : "auto")),
                lineEvenDivCountText: ui.lineEvenDivInput.text,
                lineEvenDivCount: parseInt(ui.lineEvenDivInput.text, 10),
                hExtra: ui.hExtraCheck.value,
                hExtraCountText: ui.hExtraInput.text,
                hExtraCount: parseInt(ui.hExtraInput.text, 10),
                hLeftExtra: ui.hLeftExtraCheck.value,
                makeGuides: ui.guideCheck.value,
                groupItems: ui.groupCheck.value,
                vOuter: ui.vOuterCheck.value,
                vOuterCountText: ui.vOuterInput.text,
                vOuterCount: parseInt(ui.vOuterInput.text, 10),
                vDiv: ui.vDivCheck.value,
                vTopSpace: ui.vTopSpaceCheck.value,
                vElements: ui.vElementsCheck.value,
                vDiagonalElements: ui.vDiagonalCheck.value,
                emphasizeBounds: ui.weightingCheck.value,
                isolationArea: ui.isolationAreaCheck.value,
                isolationAreaScaleText: ui.isolationAreaInput.text,
                isolationAreaScale: parseFloat(ui.isolationAreaInput.text)
            };
        }

        function validateParams(raw) {
            if (!raw.guideLayerNameText || /^\s*$/.test(raw.guideLayerNameText)) {
                return false;
            }
            if (!isStrictPositiveNumberText(raw.hScaleText) || isNaN(raw.hScale) || !isFinite(raw.hScale) || raw.hScale <= 0) {
                return false;
            }
            if (!isStrictPositiveNumberText(raw.vScaleText) || isNaN(raw.vScale) || !isFinite(raw.vScale) || raw.vScale <= 0) {
                return false;
            }
            if (!isStrictPositiveNumberText(raw.strokeWidthText) || isNaN(raw.strokeWidthValue) || !isFinite(raw.strokeWidthValue) || raw.strokeWidthValue <= 0) {
                return false;
            }
            if (raw.lineMethod === "even") {
                if (!/^\d+$/.test(raw.lineEvenDivCountText) || isNaN(raw.lineEvenDivCount) || !isFinite(raw.lineEvenDivCount) || raw.lineEvenDivCount < 1) {
                    return false;
                }
            }
            if (raw.hExtra) {
                if (!/^\d+$/.test(raw.hExtraCountText) || isNaN(raw.hExtraCount) || !isFinite(raw.hExtraCount) || raw.hExtraCount < 1) {
                    return false;
                }
            }
            if (raw.vOuter) {
                if (!/^\d+$/.test(raw.vOuterCountText) || isNaN(raw.vOuterCount) || !isFinite(raw.vOuterCount) || raw.vOuterCount < 1) {
                    return false;
                }
            }
            if (raw.vDiv) {
                if (!/^\d+$/.test(raw.vDivisionsText) || isNaN(raw.vDivisions) || !isFinite(raw.vDivisions) || raw.vDivisions < 1) {
                    return false;
                }
            }
            if (!isStrictPositiveNumberText(raw.isolationAreaScaleText) || isNaN(raw.isolationAreaScale) || !isFinite(raw.isolationAreaScale) || raw.isolationAreaScale <= 0) {
                return false;
            }
            return true;
        }

        function normalizeParams(raw) {
            var strokeWidth = convertToPoints(raw.strokeWidthValue, "strokeUnits");
            var vDivisions = raw.vDiv ? raw.vDivisions : 0;

            if (isNaN(strokeWidth) || !isFinite(strokeWidth) || strokeWidth <= 0) {
                return null;
            }

            return {
                guideLayerName: raw.guideLayerNameText,
                hScale: raw.hScale,
                vScale: raw.vScale,
                strokeWidth: strokeWidth,
                lineMethod: raw.lineMethod,
                lineEvenDivCount: raw.lineMethod === "even" ? raw.lineEvenDivCount : 0,
                hExtra: raw.hExtra,
                hExtraCount: raw.hExtra ? raw.hExtraCount : 0,
                hLeftExtra: raw.hLeftExtra,
                makeGuides: raw.makeGuides,
                groupItems: raw.groupItems,
                vOuter: raw.vOuter,
                vOuterCount: raw.vOuter ? raw.vOuterCount : 0,
                vDiv: raw.vDiv,
                vDivisions: vDivisions,
                vTopSpace: raw.vTopSpace,
                vElements: raw.vElements,
                vDiagonalElements: raw.vDiagonalElements,
                emphasizeBounds: raw.emphasizeBounds,
                isolationArea: raw.isolationArea,
                isolationAreaScale: raw.isolationAreaScale
            };
        }

        function getParams(ui) {
            var raw = readParams(ui);
            if (!validateParams(raw)) {
                return null;
            }
            return normalizeParams(raw);
        }

        function createLines(context, params, isPreview) {
            var C = context.A / params.isolationAreaScale;

            var hLineLength = context.B * params.hScale;

            var hCenter = context.centerX;
            var hHalf = hLineLength / 2;

            var hLeft = hCenter - hHalf;
            var hRight = hCenter + hHalf;

            if (params.hLeftExtra) {
                /* 右側を固定 / Fix the right side */
                var shift = hRight - (context.bRight + (context.A / 2));

                hRight -= shift;
                hLeft -= shift;
            }

            var vLineHeight = context.A * params.vScale;
            var vT, vB;
            if (params.vTopSpace) {
                vB = context.bBottom - (C * 2);
                vT = vB + vLineHeight;
            } else {
                vT = context.centerY + vLineHeight / 2;
                vB = context.centerY - vLineHeight / 2;
            }

            /* 最小値クランプ: 外側の線も含め必ず交差するよう自動補正 */
            if (!params.isolationArea) {
                var needLeft = context.bLeft;
                var needRight = context.bRight;
                var needTop = context.bTop;
                var needBottom = context.bBottom;

                if (params.vOuter && params.vOuterCount > 1) {
                    needLeft = context.bLeft - C * (params.vOuterCount - 1);
                    needRight = context.bRight + C * (params.vOuterCount - 1);
                }
                if (params.hExtra && params.hExtraCount > 1) {
                    needTop = context.bTop + C * (params.hExtraCount - 1);
                    needBottom = context.bBottom - C * (params.hExtraCount - 1);
                }

                if (hLeft > needLeft) hLeft = needLeft;
                if (hRight < needRight) hRight = needRight;
                if (vT < needTop) vT = needTop;
                if (vB > needBottom) vB = needBottom;
            }

            var strokeColor = new GrayColor();
            strokeColor.gray = 100;

            var group = context.guideLayer.groupItems.add();
            group.name = isPreview ? L("layerGuideLinesPreview") : L("layerGuideLines");

            /* 境界線を強調 / Highlight Bounds: 選択範囲の境界に一致する線の線幅を4倍にする */
            var sw = params.strokeWidth;
            var swBound = params.emphasizeBounds ? sw * 4 : sw;

            if (!params.isolationArea) {
                /* 横線 / Horizontal lines */
                var isoC = context.A / params.isolationAreaScale;
                if (params.hExtra) {
                    /* 境界線 (上下) / Boundary lines (top & bottom) */
                    drawLine(group, hLeft, context.bTop, hRight, context.bTop, strokeColor, swBound);
                    drawLine(group, hLeft, context.bBottom, hRight, context.bBottom, strokeColor, swBound);
                    /* 外側線 (hExtraCount - 1 本ずつ) / Outer lines */
                    for (var eT = 1; eT < params.hExtraCount; eT++) {
                        drawLine(group, hLeft, context.bTop + isoC * eT, hRight, context.bTop + isoC * eT, strokeColor, sw);
                    }
                    for (var eB = 1; eB < params.hExtraCount; eB++) {
                        drawLine(group, hLeft, context.bBottom - isoC * eB, hRight, context.bBottom - isoC * eB, strokeColor, sw);
                    }
                }

                /* ラインの決め方に基づく横線 / Horizontal lines based on line method */
                var lineYs = getLineYPositions(sel, params.lineMethod, params.lineEvenDivCount);
                for (var li = 0; li < lineYs.length; li++) {
                    var lySw = (params.emphasizeBounds && (lineYs[li] === context.bTop || lineYs[li] === context.bBottom)) ? swBound : sw;
                    drawLine(group, hLeft, lineYs[li], hRight, lineYs[li], strokeColor, lySw);
                }

                /* 縦線 / Vertical lines */
                if (params.vOuter) {
                    /* 境界線 / Boundary lines */
                    drawLine(group, context.bLeft, vT, context.bLeft, vB, strokeColor, swBound);
                    drawLine(group, context.bRight, vT, context.bRight, vB, strokeColor, swBound);
                    /* 外側線 (vOuterCount - 1 本ずつ) / Outer lines */
                    for (var oL = 1; oL < params.vOuterCount; oL++) {
                        drawLine(group, context.bLeft - isoC * oL, vT, context.bLeft - isoC * oL, vB, strokeColor, sw);
                    }
                    for (var oR = 1; oR < params.vOuterCount; oR++) {
                        drawLine(group, context.bRight + isoC * oR, vT, context.bRight + isoC * oR, vB, strokeColor, sw);
                    }
                }

                /* 縦分割線 / Vertical division lines */
                if (params.vDiv && params.vDivisions > 1) {
                    var vDivStep = (context.bRight - context.bLeft) / params.vDivisions;
                    for (var k = 1; k < params.vDivisions; k++) {
                        var xDiv = context.bLeft + vDivStep * k;
                        drawLine(group, xDiv, vT, xDiv, vB, strokeColor, sw);
                    }
                }

                /* 垂直エレメント / Vertical elements */
                if (params.vElements) {
                    var vertXs = getVerticalLineXPositions(sel);
                    for (var vi = 0; vi < vertXs.length; vi++) {
                        var vxSw = (params.emphasizeBounds && (vertXs[vi] === context.bLeft || vertXs[vi] === context.bRight)) ? swBound : sw;
                        drawLine(group, vertXs[vi], vT, vertXs[vi], vB, strokeColor, vxSw);
                    }
                }

                /* 斜線エレメント / Diagonal elements */
                if (params.vDiagonalElements) {
                    var diagLines = getDiagonalElementLines(sel, hLeft, vT, hRight, vB);
                    for (var di = 0; di < diagLines.length; di++) {
                        drawLine(group, diagLines[di][0][0], diagLines[di][0][1],
                            diagLines[di][1][0], diagLines[di][1][1], strokeColor, sw);
                    }
                }
            }

            /* アイソレーションエリア / Isolation area */
            if (params.isolationArea) {
                var isoC = context.A / params.isolationAreaScale;
                var isoLeft = context.bLeft - isoC;
                var isoRight = context.bRight + isoC;
                var isoTop = context.bTop + isoC;
                var isoOuterW = isoRight - isoLeft;

                var cyanColor = new CMYKColor();
                cyanColor.cyan = 60;
                cyanColor.magenta = 0;
                cyanColor.yellow = 0;
                cyanColor.black = 0;

                var isoGroup = group.groupItems.add();
                isoGroup.name = L("layerIsolationArea");
                /* 上 / Top */
                drawIsolationRect(isoGroup, isoLeft, isoTop, isoOuterW, isoC, cyanColor);
                /* 下 / Bottom */
                drawIsolationRect(isoGroup, isoLeft, context.bBottom, isoOuterW, isoC, cyanColor);
                var isoOuterH = isoTop - (context.bBottom - isoC);
                /* 左 / Left */
                drawIsolationRect(isoGroup, isoLeft, isoTop, isoC, isoOuterH, cyanColor);
                /* 右 / Right */
                drawIsolationRect(isoGroup, context.bRight, isoTop, isoC, isoOuterH, cyanColor);
            }

            /* ガイド化（本番のみ） / Convert to guides (final only) */
            if (!isPreview && params.makeGuides) {
                for (var j = 0; j < group.pathItems.length; j++) {
                    group.pathItems[j].guides = true;
                }
            }

            if (!params.groupItems && !params.makeGuides) {
                while (group.pageItems.length > 0) {
                    group.pageItems[0].move(context.guideLayer, ElementPlacement.PLACEATBEGINNING);
                }
                group.remove();
                return null;
            }

            return group;
        }
        function changeValueByArrowKey(editText, allowNegative, integerOnly, ui, context) {
            editText.addEventListener("keydown", function (event) {
                var keyName = event.keyName;
                if (keyName !== "Up" && keyName !== "Down") {
                    return;
                }

                var value = Number(editText.text);
                if (isNaN(value) || !isFinite(value)) {
                    return;
                }

                var keyboard = ScriptUI.environment.keyboardState;
                var delta = 1;

                if (keyboard.shiftKey) {
                    delta = 10;
                    if (keyName === "Up") {
                        value = Math.ceil((value + 1) / delta) * delta;
                    } else {
                        value = Math.floor((value - 1) / delta) * delta;
                    }
                } else if (keyboard.altKey && !integerOnly) {
                    delta = 0.1;
                    if (keyName === "Up") {
                        value += delta;
                    } else {
                        value -= delta;
                    }
                } else {
                    if (keyName === "Up") {
                        value += delta;
                    } else {
                        value -= delta;
                    }
                }

                if (!allowNegative && value < 0) {
                    value = 0;
                }

                if (keyboard.altKey && !integerOnly) {
                    value = Math.round(value * 10) / 10; /* 小数第1位まで / Round to 1 decimal */
                } else {
                    value = Math.round(value); /* 整数に丸め / Round to integer */
                }

                event.preventDefault();
                if (editText === ui.hScaleInput || editText === ui.vScaleInput) {
                    editText.text = value.toFixed(1);
                } else {
                    editText.text = String(value);
                }

                if (editText === ui.vScaleInput) {
                    updateVerticalTopSpaceState(ui);
                }
                if (editText === ui.lineEvenDivInput || editText === ui.vDivInput) {
                    syncDivisionUIState(ui);
                }

                if (ui.previewCheck.value) {
                    updatePreview(ui, context);
                }
            });
        }

        function updatePreview(ui, context) {
            removePreview();
            var params = getParams(ui);
            if (!params) return;
            context.guideLayer = getOrCreateLayer(context.doc, params.guideLayerName);
            previewGroup = createLines(context, params, true);
            app.redraw();
        }

        function removePreview() {
            if (previewGroup) {
                try {
                    previewGroup.remove();
                } catch (e) {
                    try {
                        $.writeln("[logo-grid-maker] Failed to remove preview: " + e);
                    } catch (logErr) { }
                }
                previewGroup = null;
                app.redraw();
            }
        }

        function drawIsolationRect(parent, left, top, width, height, color) {
            /* 塗り: 不透明度20% / Fill: 20% opacity */
            var fill = parent.pathItems.rectangle(top, left, width, height);
            fill.filled = true;
            fill.fillColor = color;
            fill.stroked = false;
            fill.opacity = 20;
            /* 線: 不透明度100% / Stroke: 100% opacity */
            var stroke = parent.pathItems.rectangle(top, left, width, height);
            stroke.filled = false;
            stroke.stroked = true;
            stroke.strokeColor = color;
            stroke.strokeWidth = 0.25;
        }

        function drawLine(parent, x1, y1, x2, y2, color, sw) {
            var line = parent.pathItems.add();
            line.setEntirePath([[x1, y1], [x2, y2]]);
            line.stroked = true;
            line.filled = false;
            line.strokeWidth = sw;
            line.strokeColor = color;
        }

        function getSelectionBounds(items) {
            var first = true;
            var l, t, r, b;

            for (var i = 0; i < items.length; i++) {
                var ib = getItemBounds(items[i]);
                if (!ib) continue;

                if (first) {
                    l = ib[0];
                    t = ib[1];
                    r = ib[2];
                    b = ib[3];
                    first = false;
                } else {
                    if (ib[0] < l) l = ib[0];
                    if (ib[1] > t) t = ib[1];
                    if (ib[2] > r) r = ib[2];
                    if (ib[3] < b) b = ib[3];
                }
            }

            if (first) return null;
            return [l, t, r, b];
        }

        function getItemBounds(item) {
            try {
                return item.visibleBounds;
            } catch (e) {
                return null;
            }
        }
    })();