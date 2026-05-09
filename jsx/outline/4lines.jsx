#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false);

(function () {
    if (app.documents.length === 0) {
        alert("ドキュメントがありません。");
        return;
    }

    var doc = app.activeDocument;
    if (doc.selection.length === 0) {
        alert("アウトライン化された文字を選択してください。");
        return;
    }

    var items = doc.selection;
    var collected = collectSelectionPaths(items);
    var pathItems = collected.pathItems;
    var compoundItems = collected.compoundItems; // CompoundPathItem単位で保持

    function collectSelectionPaths(selectionItems) {
        var result = {
            pathItems: [],
            compoundItems: []
        };

        for (var i = 0; i < selectionItems.length; i++) {
            collectPageItem(selectionItems[i], result);
        }

        return result;
    }

    function collectPageItem(item, result) {
        if (!item) return;

        if (item.typename === "PathItem") {
            result.pathItems.push(item);
            return;
        }

        if (item.typename === "CompoundPathItem") {
            result.compoundItems.push(item);
            for (var i = 0; i < item.pathItems.length; i++) {
                result.pathItems.push(item.pathItems[i]);
            }
            return;
        }

        if (item.typename === "GroupItem") {
            for (var j = 0; j < item.pageItems.length; j++) {
                collectPageItem(item.pageItems[j], result);
            }
        }
    }

    if (pathItems.length === 0) {
        alert("PathItem が見つかりません。");
        return;
    }

    function roundStep(v, step) {
        return Math.round(v / step) * step;
    }

    function analyzePathItems(pathItems, handleTolerance, anchorTolerance) {
        var stats = {
            horizontalSegmentYs: [],
            handleFlatYs: [],
            tops: [],
            bottoms: [],
            selLeft: null,
            selRight: null
        };

        for (var i = 0; i < pathItems.length; i++) {
            var p = pathItems[i];
            var gb = p.geometricBounds;
            var left = gb[0];
            var top = gb[1];
            var right = gb[2];
            var bottom = gb[3];

            stats.tops.push(top);
            stats.bottoms.push(bottom);

            if (stats.selLeft === null || left < stats.selLeft) stats.selLeft = left;
            if (stats.selRight === null || right > stats.selRight) stats.selRight = right;

            for (var k = 0; k < p.pathPoints.length; k++) {
                var pt = p.pathPoints[k];
                var nextIndex = k + 1;
                if (nextIndex >= p.pathPoints.length) {
                    if (!p.closed) break;
                    nextIndex = 0;
                }
                var nextPt = p.pathPoints[nextIndex];
                var ay = pt.anchor[1];
                var nextAy = nextPt.anchor[1];

                if (Math.abs(nextAy - ay) <= anchorTolerance) {
                    stats.horizontalSegmentYs.push(ay);
                    stats.horizontalSegmentYs.push(nextAy);
                    stats.horizontalSegmentYs.push((ay + nextAy) / 2);
                    stats.horizontalSegmentYs.push(ay);
                    continue;
                }

                var ldy = pt.leftDirection[1];
                var rdy = pt.rightDirection[1];
                if (Math.abs(ldy - ay) <= handleTolerance && Math.abs(rdy - ay) <= handleTolerance) {
                    stats.handleFlatYs.push(ay);
                }
            }
        }

        return stats;
    }

    function collectCompoundBounds(compoundItems) {
        var stats = {
            tops: [],
            bottoms: []
        };

        for (var i = 0; i < compoundItems.length; i++) {
            var gb = compoundItems[i].geometricBounds;
            stats.tops.push(gb[1]);
            stats.bottoms.push(gb[3]);
        }

        return stats;
    }

    function getMinValue(values) {
        var min = values[0];
        for (var i = 1; i < values.length; i++) {
            if (values[i] < min) min = values[i];
        }
        return min;
    }

    function getMaxValue(values) {
        var max = values[0];
        for (var i = 1; i < values.length; i++) {
            if (values[i] > max) max = values[i];
        }
        return max;
    }

    function clusterPeaks(hist, tolerance) {
        var sorted = hist.slice().sort(function (a, b) { return a.y - b.y; });
        var clusters = [];
        var cur = { sum: sorted[0].y * sorted[0].count, count: sorted[0].count, items: [sorted[0]] };
        for (var i = 1; i < sorted.length; i++) {
            var centroid = cur.sum / cur.count;
            if (sorted[i].y - centroid <= tolerance) {
                cur.sum += sorted[i].y * sorted[i].count;
                cur.count += sorted[i].count;
                cur.items.push(sorted[i]);
            } else {
                clusters.push(cur);
                cur = { sum: sorted[i].y * sorted[i].count, count: sorted[i].count, items: [sorted[i]] };
            }
        }
        clusters.push(cur);
        clusters.sort(function (a, b) { return b.count - a.count; });
        return clusters;
    }

    function snapToFlat(flatHist, estimate, tolerance, step) {
        var nearby = [];
        for (var i = 0; i < flatHist.length; i++) {
            if (Math.abs(flatHist[i].y - estimate) <= tolerance) {
                nearby.push(flatHist[i]);
            }
        }
        if (nearby.length > 0) {
            var cl = clusterPeaks(nearby, step * 3);
            return cl[0].sum / cl[0].count;
        }
        return estimate;
    }

    function snapToNearestFlat(flatHist, estimate, tolerance, step) {
        var nearby = [];
        var i;
        for (i = 0; i < flatHist.length; i++) {
            if (Math.abs(flatHist[i].y - estimate) <= tolerance) {
                nearby.push(flatHist[i]);
            }
        }
        if (nearby.length === 0) {
            return estimate;
        }

        var clusters = clusterPeaks(nearby, step * 3);
        var bestY = estimate;
        var bestDist = Infinity;
        var bestCount = -Infinity;
        for (i = 0; i < clusters.length; i++) {
            var cy = clusters[i].sum / clusters[i].count;
            var dist = Math.abs(cy - estimate);
            if (dist < bestDist || (dist === bestDist && clusters[i].count > bestCount)) {
                bestDist = dist;
                bestCount = clusters[i].count;
                bestY = cy;
            }
        }
        return bestY;
    }

    function estimateDescenderFromBottomClusters(bottomHist, flatHist, baseline, step) {
        var bottomClusters = clusterPeaks(bottomHist, step * 3);
        var descender = baseline;
        var minCount = Math.max(1, bottomClusters[0].count * 0.2);
        var bestGap = Infinity;
        var i;

        for (i = 0; i < bottomClusters.length; i++) {
            if (bottomClusters[i].count < minCount) continue;
            var cy = bottomClusters[i].sum / bottomClusters[i].count;
            if (cy >= baseline - step * 2) continue;
            var gap = baseline - cy;
            if (gap < bestGap) {
                bestGap = gap;
                descender = cy;
            }
        }

        if (descender !== baseline) {
            return snapToNearestFlat(flatHist, descender, step * 5, step);
        }

        return snapToNearestFlat(flatHist, baseline, step * 5, step);
    }

    var handleTol = 0.5; // ハンドルが水平とみなす許容差
    var anchorTol = 0.5; // 隣接アンカー間の水平セグメントとみなす許容差
    var pathStats = analyzePathItems(pathItems, handleTol, anchorTol);
    var selLeft = pathStats.selLeft;
    var selRight = pathStats.selRight;

    var step = 1.0;

    function histogram(values, step) {
        var map = {};
        for (var i = 0; i < values.length; i++) {
            var key = roundStep(values[i], step).toFixed(2);
            if (!map[key]) map[key] = 0;
            map[key]++;
        }
        var arr = [];
        for (var k in map) {
            arr.push({ y: parseFloat(k), count: map[k] });
        }
        arr.sort(function (a, b) { return a.y - b.y; });
        return arr;
    }

    function topPeaks(hist, minCount) {
        var out = [];
        for (var i = 0; i < hist.length; i++) {
            if (hist[i].count >= minCount) out.push(hist[i]);
        }
        out.sort(function (a, b) { return b.count - a.count; });
        return out;
    }

    function mergeWeightedYs(horizontalSegmentYs, handleFlatYs, horizontalWeight, handleWeight) {
        var merged = [];
        var i;
        var j;

        for (i = 0; i < horizontalSegmentYs.length; i++) {
            for (j = 0; j < horizontalWeight; j++) {
                merged.push(horizontalSegmentYs[i]);
            }
        }

        for (i = 0; i < handleFlatYs.length; i++) {
            for (j = 0; j < handleWeight; j++) {
                merged.push(handleFlatYs[i]);
            }
        }

        return merged;
    }

    function getBaselineCandidateFromBottomClusters(bottomHist, step) {
        var bottomClusters = clusterPeaks(bottomHist, step * 3);
        var baselineCandidate = bottomClusters[0].sum / bottomClusters[0].count;
        var minCount = Math.max(1, bottomClusters[0].count * 0.25);
        var i;

        for (i = 0; i < bottomClusters.length; i++) {
            if (bottomClusters[i].count < minCount) continue;
            var cy = bottomClusters[i].sum / bottomClusters[i].count;
            if (cy > baselineCandidate) {
                baselineCandidate = cy;
            }
        }

        return baselineCandidate;
    }

    function snapBaselineToFlat(baseline, baselineFlatHist, step, tolerance) {
        return snapToNearestFlat(baselineFlatHist, baseline, tolerance, step);
    }

    function estimateBaseline(bottomHist, baselineFlatHist, step) {
        var baselineCandidate = getBaselineCandidateFromBottomClusters(bottomHist, step);
        return snapBaselineToFlat(baselineCandidate, baselineFlatHist, step, step * 5);
    }

    function estimateTypographyLines(pathStats, compoundStats, step) {
        var horizontalSegmentYs = pathStats.horizontalSegmentYs;
        var handleFlatYs = pathStats.handleFlatYs;
        var horizontalWeight = 12;
        var handleWeight = 1;
        var baselineHorizontalWeight = 24;
        var baselineHandleWeight = 1;
        var flatAnchorsY = mergeWeightedYs(horizontalSegmentYs, handleFlatYs, horizontalWeight, handleWeight);
        var baselineFlatAnchorsY = mergeWeightedYs(horizontalSegmentYs, handleFlatYs, baselineHorizontalWeight, baselineHandleWeight);
        var tops = pathStats.tops;
        var bottoms = pathStats.bottoms;
        var cpTops = compoundStats.tops;
        var cpBottoms = compoundStats.bottoms;

        var topHist = histogram(tops, step);
        var bottomHist = histogram(bottoms, step);
        var flatHist = histogram(flatAnchorsY, step);
        var baselineFlatHist = histogram(baselineFlatAnchorsY, step);

        var topPeaksList = topPeaks(topHist, 2);
        var bottomPeaksList = topPeaks(bottomHist, 2);
        if (topPeaksList.length === 0 || bottomPeaksList.length === 0) {
            return null;
        }

        var baseline = null;
        var meanLine = null;

        if (cpBottoms.length >= 2) {
            var cpBottomHist = histogram(cpBottoms, step);
            var cpBottomClusters = clusterPeaks(cpBottomHist, step * 3);
            var cpbMax = -Infinity;
            var cpbMinCount = cpBottomClusters[0].count * 0.3;
            for (var i = 0; i < cpBottomClusters.length; i++) {
                if (cpBottomClusters[i].count >= cpbMinCount) {
                    var cy = cpBottomClusters[i].sum / cpBottomClusters[i].count;
                    if (cy > cpbMax) cpbMax = cy;
                }
            }
            if (cpbMax > -Infinity) baseline = cpbMax;
        }

        if (cpTops.length >= 2) {
            var cpTopHist = histogram(cpTops, step);
            var cpTopClusters = clusterPeaks(cpTopHist, step * 3);
            var cptMinCount = cpTopClusters[0].count * 0.3;
            var cptBestDist = Infinity;
            var refBase = baseline !== null ? baseline : (clusterPeaks(bottomHist, step * 3))[0].sum / (clusterPeaks(bottomHist, step * 3))[0].count;
            for (var i = 0; i < cpTopClusters.length; i++) {
                if (cpTopClusters[i].count >= cptMinCount) {
                    var cy = cpTopClusters[i].sum / cpTopClusters[i].count;
                    if (cy > refBase) {
                        var dist = cy - refBase;
                        if (dist < cptBestDist) {
                            cptBestDist = dist;
                            meanLine = cy;
                        }
                    }
                }
            }
        }

        if (baseline !== null) baseline = snapBaselineToFlat(baseline, baselineFlatHist, step, step * 5);
        if (meanLine !== null) meanLine = snapToFlat(flatHist, meanLine, step * 5, step);

        if (baseline === null) {
            baseline = estimateBaseline(bottomHist, baselineFlatHist, step);
        }
        baseline = snapBaselineToFlat(baseline, baselineFlatHist, step, step * 3);

        var descender = estimateDescenderFromBottomClusters(bottomHist, flatHist, baseline, step);

        var topClusters = clusterPeaks(topHist, step * 3);
        var ascender = getMaxValue(tops);

        if (meanLine === null) {
            var flatXHist = [];
            for (var i = 0; i < flatHist.length; i++) {
                var fy = flatHist[i].y;
                if (fy > baseline + step * 4 && fy < ascender - step * 4) {
                    flatXHist.push(flatHist[i]);
                }
            }
            if (flatXHist.length > 0) {
                var fxc = clusterPeaks(flatXHist, step * 3);
                var fxMin = fxc[0].count * 0.3;
                var fxBest = Infinity;
                for (var i = 0; i < fxc.length; i++) {
                    if (fxc[i].count >= fxMin) {
                        var cy = fxc[i].sum / fxc[i].count;
                        var dist = cy - baseline;
                        if (dist < fxBest) {
                            fxBest = dist;
                            meanLine = cy;
                        }
                    }
                }
            }
        }
        if (meanLine === null) {
            var xMinCount = topClusters[0].count * 0.2;
            var xBest = Infinity;
            for (var i = 0; i < topClusters.length; i++) {
                var cy = topClusters[i].sum / topClusters[i].count;
                if (topClusters[i].count >= xMinCount && cy > baseline + step * 4 && cy < ascender - step * 4) {
                    var dist = cy - baseline;
                    if (dist < xBest) {
                        xBest = dist;
                        meanLine = cy;
                    }
                }
            }
        }
        if (meanLine === null) {
            meanLine = baseline + (ascender - baseline) * 0.7;
        }

        return {
            descender: descender,
            baseline: baseline,
            meanLine: meanLine,
            ascender: ascender
        };
    }

    var compoundStats = collectCompoundBounds(compoundItems);
    var estimatedLines = estimateTypographyLines(pathStats, compoundStats, step);
    if (!estimatedLines) {
        alert("十分なピークが見つかりません。");
        return;
    }

    var descender = estimatedLines.descender;
    var baseline = estimatedLines.baseline;
    var meanLine = estimatedLines.meanLine;
    var ascender = estimatedLines.ascender;

    var layer = doc.activeLayer;

    var GUIDE_STROKE_WIDTH = 0.25;
    var GUIDE_EXTEND_X = 20;

    function createBlackColor() {
        var c = new RGBColor();
        c.red = 0;
        c.green = 0;
        c.blue = 0;
        return c;
    }

    function applyLineStyle(line, options) {
        line.stroked = true;
        line.filled = false;
        line.strokeWidth = options.strokeWidth;
        line.strokeColor = options.strokeColor;
        if (options.name) line.name = options.name;
        return line;
    }

    function createLine(layer, points, options) {
        var line = layer.pathItems.add();
        line.setEntirePath(points);
        return applyLineStyle(line, options);
    }

    function drawHorizontalGuideLine(layer, left, right, y, name) {
        return createLine(layer, [[left - GUIDE_EXTEND_X, y], [right + GUIDE_EXTEND_X, y]], {
            strokeWidth: GUIDE_STROKE_WIDTH,
            strokeColor: createBlackColor(),
            name: name
        });
    }

    var group = layer.groupItems.add();
    group.name = "タイポグラフィガイド";

    var l1 = drawHorizontalGuideLine(layer, selLeft, selRight, descender, "ディセンダーライン");
    var l2 = drawHorizontalGuideLine(layer, selLeft, selRight, baseline, "ベースライン");
    var l3 = drawHorizontalGuideLine(layer, selLeft, selRight, meanLine, "ミーンライン");
    var l4 = drawHorizontalGuideLine(layer, selLeft, selRight, ascender, "アセンダーライン");

    l1.move(group, ElementPlacement.PLACEATEND);
    l2.move(group, ElementPlacement.PLACEATEND);
    l3.move(group, ElementPlacement.PLACEATEND);
    l4.move(group, ElementPlacement.PLACEATEND);

})();
