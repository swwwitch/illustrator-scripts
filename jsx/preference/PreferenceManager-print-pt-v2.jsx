/*
  PreferenceManager-print-pt-v2.jsx
  Illustrator ExtendScript
  - 単位と数値インクリメントの設定を完全にExtendScriptで実装
*/

function main() {
    if (app.documents.length === 0) {
        alert("ドキュメントを開いてから実行してください。");
        return;
    }

    var versionParts = app.version.split(".");
    var aiMajor = parseInt(versionParts[0], 10);
    var aiMinor = parseInt(versionParts[1] || "0", 10);

    if (aiMajor < 25) {
        alert("このスクリプトはIllustrator 25以降でのみ動作します。");
        return;
    }

    // 設定値（必要に応じてUIや外部入力に置き換え可）
    var param = {
        units_general: "mm",
        units_stroke:  "pt",
        units_type:    "pt",
        units_asian:   "pt",
        num_key:       "0.1",
        num_r:         "1",
        num_size:      "1",
        num_baseline:  "0.1"
    };

    // 数値設定
    setPref("cursorKeyLength", param.units_general, param.num_key);
    setPref("ovalRadius", param.units_general, param.num_r);
    setPref("text/sizeIncrement", param.units_type, param.num_size);
    setPref("text/riseIncrement", param.units_asian, param.num_baseline);

    // 単位設定
    setUnits("rulerType", param.units_general);
    setUnits("strokeUnits", param.units_stroke);
    setUnits("text/units", param.units_type);
    setUnits("text/asianunits", param.units_asian);

    slideUnitsViaAppleScript();

    alert("Illustratorの単位とインクリメント設定を反映しました。");
}

function setPref(prefKey, dstUnits, num) {
    var value = calcWithUnit(num, dstUnits);
    if (dstUnits !== "pt") {
        value = convertUnit(value, dstUnits, "pt");
    }
    if (value !== undefined) {
        app.preferences.setRealPreference(prefKey, value);
    }
}

function setUnits(prefKey, dstUnits) {
    var db = {
        "pt": 2, "pc": 3, "in": 0, "mm": 1, "cm": 4, "Q/H": 5, "px": 6
    };

    if (db[dstUnits] !== undefined) {
        app.preferences.setIntegerPreference(prefKey, db[dstUnits]);
    } else {
        alert("未対応の単位: " + dstUnits);
    }
}

function convertUnit(value, srcUnit, dstUnit) {
    var qReg = /^(?:Q\/H|[qh])$/i;
    var mmUnit = "mm";

    if (qReg.test(srcUnit)) {
        value = Number(value) * 0.25;
        srcUnit = mmUnit;
    }
    if (qReg.test(dstUnit)) {
        value = Number(value) * 4;
        dstUnit = mmUnit;
    }

    var result;
    try {
        result = new UnitValue(value, srcUnit).as(dstUnit);
    } catch (e) {
        result = undefined;
    }
    return result;
}

function calcWithUnit(str, defaultUnits) {
    if (!defaultUnits) defaultUnits = "mm";
    var newStr = str.replace(/[　\s]+/g, " ").replace(",", "");
    var regUnitValue = /([0-9]+(?:\.[0-9]+)?)( ?[a-zA-Z\/]+)/g;
    newStr = newStr.replace(regUnitValue, function (m0, m1, m2) {
        return convertUnit(m1, m2.replace(/^\s+|\s+$/g, ""), defaultUnits);
    });

    var result;
    try {
        result = eval(newStr);
    } catch (e) {
        result = undefined;
    }
    return result;
}

function runAppleScript(scriptText) {
    try {
        // 一時ファイルにAppleScriptを書き出す
        var tempFile = new File(Folder.temp + "/temp_unit_slide.scpt");
        tempFile.encoding = "UTF-8";
        tempFile.open("w");
        tempFile.write(scriptText);
        tempFile.close();

        // osascript 経由で実行
        var command = "/usr/bin/osascript '" + tempFile.fsName + "'";
        var result = app.system(command);
        return result;
    } catch (e) {
        alert("AppleScript 実行エラー: " + e.message);
    }
}

function slideUnitsViaAppleScript() {
    var appleScript =
        'tell application "Adobe Illustrator"\n' +
        'activate\n' +
        'end tell\n' +
        'delay 0.3\n' +
        'tell application "System Events"\n' +
        'keystroke "u" using {command down, shift down}\n' +
        'delay 0.5\n' +
        'repeat 4 times\n' + // フィールド数に応じて調整
        '    keystroke tab\n' +
        'end repeat\n' +
        'keystroke return\n' +
        'end tell';

    runAppleScript(appleScript);
}

main();
