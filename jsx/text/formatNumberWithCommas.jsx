#target illustrator
app.preferences.setBooleanPreference('ShowExternalJSXWarning', false); 

/*
このスクリプトは、Illustrator ドキュメント内の数値を検出し、指定ルールに基づいて桁区切りのカンマを付与します。
郵便番号・電話番号・MAC アドレス・クレジットカード番号などは除外対象とし、
必要に応じて2段階ダイアログで対象確認を行えます。
ポイント文字・エリア内文字・パス上文字・グループやクリップグループ内のテキストにも対応します。
*/

### スクリプト名：

数値に桁区切りカンマを付与（除外条件付き）

### GitHub：

https://github.com/swwwitch/illustrator-scripts

### 概要：

- Illustrator ドキュメント内の数字に桁区切りのカンマを自動で付与
- 除外条件（郵便番号、電話番号、MACアドレスなど）を設定可能

### 主な機能：

- 選択範囲または全ドキュメントを対象に処理
- 除外条件のカスタマイズ（郵便番号／西暦／電話番号／記号付き数値など）
- 間違ったカンマ位置の修正
- プレビュー確認ダイアログ

### 処理の流れ：

1. 対象範囲のテキストを取得
2. 除外条件に該当する数値をフィルタリング
3. カンマ付与対象のプレビューを表示
4. 確定後にテキストを置換

### 更新履歴：

- v1.0 (20250812) : 初期バージョン

---

### Script Name:

Add Thousand Separators to Numbers (with Exclusions)

### GitHub:

https://github.com/swwwitch/illustrator-scripts

### Overview:

- Automatically add thousand separators to numbers in Illustrator documents
- Configurable exclusion rules (postal codes, phone numbers, MAC addresses, etc.)

### Main Features:

- Process selected objects or the entire document
- Customizable exclusion conditions (postal code / year / phone number / symbols)
- Fix incorrectly placed commas
- Preview confirmation dialog

### Process Flow:

1. Retrieve text from the target range
2. Filter out numbers matching exclusion conditions
3. Show preview list of numbers to be processed
4. Replace text after confirmation

### Update History:

- v1.0 (20250812) : Initial version

var SCRIPT_VERSION = "v1.0";

function getCurrentLang() {
    return ($.locale.indexOf("ja") === 0) ? "ja" : "en";
}
var lang = getCurrentLang();

/* 日英ラベル定義 / Japanese-English label definitions */

var LABELS = {
    mainTitle: { ja: "桁区切りのカンマ " + SCRIPT_VERSION, en: "Add Thousands Separators " + SCRIPT_VERSION },
    panelTarget: { ja: "対象", en: "Scope" },
    rbSelection: { ja: "選択したオブジェクトのみ", en: "Selection only" },
    rbDocument: { ja: "ドキュメントすべて", en: "Whole document" },
    panelExclude: { ja: "除外オプション", en: "Exclude options" },
    exYears: { ja: "西暦", en: "Years" },
    exPostal: { ja: "郵便番号", en: "Postal codes" },
    exSlash: { ja: "スラッシュ（／・/）", en: "Slash-adjacent" },
    exPhone: { ja: "電話番号／携帯番号", en: "Phone numbers" },
    exCC: { ja: "クレジットカード番号", en: "Credit card numbers" },
    exMAC: { ja: "MACアドレス", en: "MAC addresses" },
    exVehicle: { ja: "自動車ナンバー", en: "Vehicle plate (last-4)" },
    ok: { ja: "OK", en: "OK" },
    cancel: { ja: "キャンセル", en: "Cancel" },
    dlg2Title: { ja: "カンマ付与対象の確認", en: "Review Numbers to Add Commas" },
    dlg2Note: { ja: "カンマを付ける数値を選択してください", en: "Select numbers to add/fix commas" },
    colSelect: { ja: "選択", en: "Select" },
    colFrame: { ja: "#", en: "#" },
    colValue: { ja: "数値", en: "Value" },
    selectAll: { ja: "✓", en: "✓" },
};

// Insert commas into a string of digits (no floats, no sign)
function insertCommasToDigits(digits) {
    // digits: string of 0-9 only
    var out = '';
    var len = digits.length;
    for (var i = 0; i < len; i++) {
        if (i > 0 && ((len - i) % 3) === 0) out += ',';
        out += digits.charAt(i);
    }
    return out;
}

function formatNumberWithCommas(value) {
    // Accept string or number; operate on string to preserve original fractional digits
    var s = String(value);
    // Normalize full-width signs/commas/dots to half-width for consistent formatting
    s = s.replace(/[，]/g, ',').replace(/[．]/g, '.').replace(/[＋]/g, '+').replace(/[－]/g, '-');
    var sign = '';
    if (s.charAt(0) === '+' || s.charAt(0) === '-') {
        sign = s.charAt(0);
        s = s.substring(1);
    }
    var parts = s.split('.');
    var intPart = parts[0].replace(/[，,]/g, ''); // remove half/full-width commas if any
    var fracPart = parts.length > 1 ? ('.' + parts[1]) : '';
    return sign + insertCommasToDigits(intPart) + fracPart;
}

// ダイアログの位置シフト / Shift dialog position
function shiftDialogPosition(dlg, offsetX, offsetY) {
    dlg.onShow = function() {
        try {
            var currentX = dlg.location[0];
            var currentY = dlg.location[1];
            dlg.location = [currentX + offsetX, currentY + offsetY];
        } catch (e) {}
    };
}
// ダイアログの透過度設定 / Set dialog opacity
function setDialogOpacity(dlg, opacityValue) {
    try {
        dlg.opacity = opacityValue;
    } catch (e) {}
}
// 既定値 / Defaults
var DIALOG_OFFSET_X = 300;
var DIALOG_OFFSET_Y = 0;
var DIALOG_OPACITY = 0.97;
// 視覚順の上下判定フラグ / Order by visual top: desc=true means upper items first
var ORDER_TOP_DESC = true; // true: 上(大きいtop)→下, false: 下→上

// UI metrics (list sizing)
var PREVIEW_LIST_WIDTH = 260;           // list.preferredSize[0]
var PREVIEW_COL_WIDTHS = [40, 70, 150]; // [選択, #, 数値]
var LIST_HEADER_H = 24;                 // header height
var LIST_ROW_H = 20;                    // row height
var LIST_PADDING = 40;                  // padding/margins
var LIST_MAX_H = 520;                   // max height
var LIST_MIN_ROWS = 4;                  // min visible rows
var LIST_MAX_ROWS = 16;                 // max visible rows before scroll

/*
選択範囲からテキストフレームを収集（ポイント/エリア/パス上、グループ/クリップグループ内を含む）
Collect text frames from selection (point/area/path text, including inside groups/clip groups)
*/
function isTextFrame(obj) {
    return obj && obj.typename === "TextFrame";
}

function collectTextFramesFromItem(item, bucket) {
    if (!item) return;

    // 直接テキストフレームなら追加 / If this item itself is a TextFrame, add it
    if (isTextFrame(item)) {
        bucket.push(item);
        return;
    }

    // 直下のテキストフレームを収集（Group/ClipGroup などコンテナ想定）
    // Collect direct text frames if the item exposes a textFrames collection
    try {
        if (item.textFrames && item.textFrames.length) {
            for (var i = 0; i < item.textFrames.length; i++) {
                bucket.push(item.textFrames[i]);
            }
        }
    } catch (e) {}

    // ネストしたグループを再帰探索（GroupItem や ClipGroup を含む）
    // Recurse into nested groups if groupItems is available
    try {
        if (item.groupItems && item.groupItems.length) {
            for (var g = 0; g < item.groupItems.length; g++) {
                collectTextFramesFromItem(item.groupItems[g], bucket);
            }
        }
    } catch (e) {}
}

function collectTextFramesFromSelection(selectionArray) {
    var result = [];
    for (var i = 0; i < selectionArray.length; i++) {
        collectTextFramesFromItem(selectionArray[i], result);
    }
    return result;
}

function collectTextFramesFromDocument(doc) {
    var result = [];
    var pageItems = doc.pageItems;
    for (var i = 0; i < pageItems.length; i++) {
        collectTextFramesFromItem(pageItems[i], result);
    }
    return result;
}

/**
 * Check slash adjacency around a numeric token.
 * 数値トークンの直前直後にスラッシュ（／,/）がある場合は除外。
 * @param {string} match - numeric token (no commas)
 * @param {number} offset - start index in normalized string
 * @param {string} string - normalized string (no commas)
 * @param {boolean} enabled - whether this exclusion is enabled
 * @returns {boolean} true = OK (eligible), false = exclude
 */
function _checkSlashAdjacency(match, offset, string, enabled){
    if (!enabled) return true;
    var prev = (offset > 0) ? string[offset - 1] : '';
    var next = (offset + match.length < string.length) ? string[offset + match.length] : '';
    return !(prev === '／' || prev === '/' || next === '／' || next === '/');
}
/**
 * Exclude all-zero numbers like 0000 or 0.000
 * 全て0で構成された数値を除外。
 * @param {string} match
 * @returns {boolean}
 */
function _checkZeroOnly(match){
    return !/^[-+]?0+(?:\.0+)?$/.test(match);
}
/**
 * Context-aware 4-digit year exclusion.
 * 4桁(1000–2999)でも、直後が「年」や日付区切り（/.-）の場合のみ除外。
 * @param {string} match
 * @param {number} offset
 * @param {string} string
 * @param {boolean} enabled
 * @returns {boolean}
 */
function _checkYear(match, offset, string, enabled){
    if (!enabled) return true;
    var yearNum = parseInt(match, 10);
    if (!/^\d{4}$/.test(match) || yearNum < 1000 || yearNum > 2999) return true;
    var prevIdx = offset - 1;
    var nextIdx = offset + match.length;
    var prevChar = prevIdx >= 0 ? string.charAt(prevIdx) : '';
    var nextChar = nextIdx < string.length ? string.charAt(nextIdx) : '';
    var k = nextIdx; while (k < string.length && /[\s\t]/.test(string.charAt(k))) k++;
    var nextNonSpace = k < string.length ? string.charAt(k) : '';
    var j = prevIdx; while (j >= 0 && /[\s\t]/.test(string.charAt(j))) j--;
    var prevNonSpace = j >= 0 ? string.charAt(j) : '';
    // Accept full-width separators and check after skipping spaces
    var dateSepSet = /[\/\.\-－―／．]/;
    var likelyYear = (nextNonSpace === '年') || dateSepSet.test(nextNonSpace) || dateSepSet.test(prevNonSpace);
    // If there is no explicit context ("年" or a date separator nearby), treat as a plain 4-digit number (eligible)
    return !likelyYear;
}
/**
 * Expand [start,end) to include a token defined by a character class.
 * 文字クラスで定義されるトークン全体を左右に拡張して取得。
 * @param {string} str
 * @param {number} start
 * @param {number} end
 * @param {RegExp} re - single-char class like /[0-9\-\s]/
 * @returns {{start:number,end:number,token:string}}
 */
function _expandToken(str, start, end, re){
    while (start > 0 && re.test(str.charAt(start - 1))) start--;
    while (end < str.length && re.test(str.charAt(end))) end++;
    return {start:start, end:end, token:str.substring(start,end)};
}
/**
 * Normalize full-width characters commonly used in phone numbers to half-width.
 * 電話番号に現れがちな全角文字を半角へ（数字・記号・スペース・括弧・+）。
 */
function toHalfWidthForPhone(s){
    if (!s || !s.replace) return s;
    return s
        .replace(/[０-９]/g, function(ch){ var code = ch.charCodeAt(0); return String.fromCharCode(code - 0xFEE0); })
        .replace(/[－ー―–—]/g, '-')   // various dashes to hyphen
        .replace(/[（]/g, '(')
        .replace(/[）]/g, ')')
        .replace(/[＋]/g, '+')
        .replace(/[\u3000]/g, ' ');   // full-width space
}
/**
 * Normalize full-width characters for MAC addresses to half-width.
 * MAC アドレスの全角文字（数字・A-F・コロン・ハイフン）を半角へ。
 */
function toHalfWidthForMAC(s){
    if (!s || !s.replace) return s;
    return s
        .replace(/[０-９]/g, function(ch){ var code = ch.charCodeAt(0); return String.fromCharCode(code - 0xFEE0); })
        .replace(/[Ａ-Ｆ]/g, function(ch){ return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
        .replace(/[ａ-ｆ]/g, function(ch){ return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
        .replace(/[：]/g, ':')
        .replace(/[－ー―–—]/g, '-');
}
/**
 * Exclude phone numbers (JP + generic international).
 * 日本の固定/携帯/IP/フリーダイヤルと国際表記、括弧・スペース・内線を考慮して除外。
 * @returns {boolean}
 */
function _checkPhone(match, offset, string, enabled){
    if (!enabled) return true;
    var s = offset, e = offset + match.length;
    var ex = _expandToken(string, s, e, /[0-9０-９\-－ \u3000\(\)（）＋+]/); // allow half/full-width spaces only (no tab)
    // Include optional leading '+' (half or full width)
    if (ex.start > 0) {
        var ch = string.charAt(ex.start - 1);
        if (ch === '+' || ch === '＋') { ex.start--; ex.token = string.substring(ex.start, ex.end); }
    }
    var tokenNorm = toHalfWidthForPhone(ex.token);
    var phoneRegexes = [
        /^(?:\(\d{2,4}\)|\d{2,4})[-)]?\d{2,4}-\d{4}$/,
        /^0[789]0(?:-\d{4}-\d{4}|\d{8})$/,
        /^050(?:-\d{4}-\d{4}|\d{8})$/,
        /^0120(?:-\d{3}-\d{3}|\d{6})$/,
        /^\d{2,4}-\d{2,4}-\d{4}$/,
        /^0[5789]0-\d{4}-\d{4}$/,
        /^(050|070|080|090)-\d{4}-\d{4}$/,
        /^\+?\d+(?:[-\s]\d+){4}$/,
        /^\+?\d+(?:[-\s]\d+){3}$/,
        /^\+?\d+(?:[-\s]\d+){2}$/,
        /^\+?\(?\d+\)?(?:[-\s]\(?\d+\)?){1,4}$/,
        /^\(\d+\)\s*\d+(?:[-\s]\d+){1,3}$/,
        /^\+?\(?\d+\)?(?:[-\s]\(?\d+\)?){1,4}\s*(?:ext\.?|x|内線)\s*\d{1,5}$/
    ];
    for (var i=0;i<phoneRegexes.length;i++) if (phoneRegexes[i].test(tokenNorm)) return false;
    return true;
}
/**
 * Luhn checksum validator for digit strings (13–19 digits typical for cards).
 * Luhn チェックサム検証（クレジットカード判定の精度向上）。
 * @param {string} digits - numeric string only
 * @returns {boolean}
 */
function _luhnValid(digits){
    var len = digits.length;
    var sum = 0;
    var alt = false;
    for (var i = len - 1; i >= 0; i--) {
        var c = digits.charCodeAt(i) - 48; // '0' = 48
        if (c < 0 || c > 9) return false;
        if (alt) {
            c *= 2;
            if (c > 9) c -= 9;
        }
        sum += c;
        alt = !alt;
    }
    return (sum % 10) === 0;
}

/**
 * Exclude major credit card numbers using brand patterns and Luhn checksum.
 * 主要ブランド(VISA/Master/Discover/Diners/Amex/JCB)の桁・先頭パターン＋Luhnチェックサムで除外。
 * @returns {boolean}
 */
function _checkCreditCard(match, offset, string, enabled){
    if (!enabled) return true;
    var s = offset, e = offset + match.length;
    var ex = _expandToken(string, s, e, /[0-9\- \(\)]/); // allow half-space only (no tab)
    var digitsOnly = ex.token.replace(/\D+/g,'');
    // Typical card length range
    var plausibleLen = digitsOnly.length >= 13 && digitsOnly.length <= 19;
    var ccPatterns = [
        /^4\d{12}(?:\d{3})?$/,              // VISA
        /^5[1-5]\d{14}$/,                    // MasterCard
        /^6011\d{12}$/,                      // Discover (simplified)
        /^3(?:0[0-5]|[68]\d)\d{11}$/,       // Diners Club
        /^3[47]\d{13}$/,                     // Amex
        /^(?:2131|1800|35\d{3})\d{11}$/     // JCB
    ];
    for (var i=0;i<ccPatterns.length;i++) if (ccPatterns[i].test(digitsOnly)) return false;
    // Fallback: if length is plausible and Luhn passes, also exclude
    if (plausibleLen && _luhnValid(digitsOnly)) return false;
    return true;
}
/**
 * Exclude MAC addresses in colon or hyphen notation.
 * コロン/ハイフン区切りの MAC アドレスを除外。
 * @returns {boolean}
 */
function _checkMAC(match, offset, string, enabled){
    if (!enabled) return true;
    var s = offset, e = offset + match.length;
    var ex = _expandToken(string, s, e, /[0-9０-９A-Fa-fＡ-Ｆａ-ｆ:\-：－]/);
    var tokenNorm = toHalfWidthForMAC(ex.token);
    return !(/^(?:[0-9A-Fa-f]{2}:){5}[0-9A-Fa-f]{2}$/.test(tokenNorm) || /^(?:[0-9A-Fa-f]{2}-){5}[0-9A-Fa-f]{2}$/.test(tokenNorm));
}
/**
 * Exclude Japanese vehicle plate last-4 block patterns.
 * 自動車ナンバー末尾4桁パターン（・やハイフン許容）を除外。
 * @returns {boolean}
 */
function _checkVehicle(match, offset, string, enabled){
    if (!enabled) return true;
    var s = offset, e = offset + match.length;
    var ex = _expandToken(string, s, e, /[0-9・\-\sぁ-んA-Za-z]/);
    var vToken = ex.token.replace(/\s+/g,'');
    return !/^(?:・|\d){2}-?(?:・|\d)\d$/.test(vToken);
}

function isEligibleNumericMatch(match, offset, string, opts) {
    if (!_checkSlashAdjacency(match, offset, string, opts.excludeSlashAny)) return false;
    if (!_checkZeroOnly(match)) return false;
    // NOTE: Postal-code exclusion is handled by precomputed offsets in the caller.
    // NOTE: Plain 7-digit numbers are allowed even when postal-code exclusion is on.
    if (!_checkYear(match, offset, string, opts.excludeYears)) return false;
    if (!_checkPhone(match, offset, string, opts.excludePhoneNumbers)) return false;
    if (!_checkCreditCard(match, offset, string, opts.excludeCreditCards)) return false;
    if (!_checkMAC(match, offset, string, opts.excludeMAC)) return false;
    if (!_checkVehicle(match, offset, string, opts.excludeVehicle)) return false;
    return true;
}

/**
 * Normalize full-width digits and hyphen to half-width for fast regex.
 * 全角数字(０-９)と全角ハイフン(－)を半角へ正規化。
 */
function toHalfWidthDigitsHyphen(s) {
    if (!s || !s.replace) return s;
    return s.replace(/[０-９－]/g, function(ch){
        if (ch === '－') return '-';
        var code = ch.charCodeAt(0);
        // '０'..'９' (0xFF10..0xFF19) -> '0'..'9' (0x30..0x39)
        if (code >= 0xFF10 && code <= 0xFF19) return String.fromCharCode(code - 0xFEE0);
        return ch;
    });
}

// 正規化文字列から郵便番号(〒\s*ddd-dddd)の数値開始オフセットを収集
function buildPostalOffsetsFromNormalized(str) {
    var map = {};
    var norm = toHalfWidthDigitsHyphen(str);
    var re = /(〒\s*)?(\d{3})-(\d{4})/g;
    var m;
    while ((m = re.exec(norm)) !== null) {
        var lead = m[1] ? m[1].length : 0; // optional '〒' + spaces
        var start3 = m.index + lead; // 3桁部分の開始
        var start4 = start3 + 4; // 3桁+ハイフンの直後 = 4桁部分の開始
        map[start3] = true;
        map[start4] = true;
    }
    return map;
}

// 数値テキストのみカンマ区切りに変換 / Convert only purely-numeric texts with grouping separators
function formatTextFrameIfNumeric(tf, excludeYears, excludePostalCodes, excludeSlashAny, selectedOffsetSet, excludePhoneNumbers, excludeCreditCards,
    excludeMAC,
    excludeVehicle) {
    try {
        var raw = tf.contents;
        if (typeof raw !== 'string') return;
        var trimmed = raw.replace(/^\s+|\s+$/g, "");
        var normalized = trimmed.replace(/[，,]/g, "");
        var postalOffsetsForString = excludePostalCodes ? buildPostalOffsetsFromNormalized(normalized) : null;
        var opts = {
            excludeYears: excludeYears,
            excludePostalCodes: excludePostalCodes,
            excludeSlashAny: excludeSlashAny,
            excludePhoneNumbers: excludePhoneNumbers,
            excludeCreditCards: excludeCreditCards,
            excludeMAC: excludeMAC,
            excludeVehicle: excludeVehicle
        };

        // Find all numeric substrings and replace them with comma formatted versions or remove commas
        var replaced = normalized.replace(/[+\-＋－]?[0-9０-９]+(?:[\.．][0-9０-９]+)?/g, function(match, offset, string) {
            if (excludePostalCodes) {
                // オフセットマップで郵便番号の一部なら即除外
                if (postalOffsetsForString && postalOffsetsForString[offset]) {
                    return match;
                }
                // フォールバック：ローカル周辺を直接判定（〒 optional, ddd[-－]dddd, 半角・全角数字対応）
                var start = Math.max(0, offset - 5);
                var end = Math.min(string.length, offset + match.length + 6);
                var around = string.substring(start, end);
                var aroundNorm = toHalfWidthDigitsHyphen(around);
                if (/(?:^|[^\d])(?:〒\s*)?\d{3}-\d{4}(?!\d)/.test(aroundNorm)) {
                    return match;
                }
            }
            if (!isEligibleNumericMatch(match, offset, string, opts)) {
                return match;
            }
            // If a selection set is provided, only process matches whose offset is included or whose token key is included
            if (selectedOffsetSet) {
                var okByOffset = !!selectedOffsetSet[offset];
                var okByKey = selectedOffsetSet._keys ? !!selectedOffsetSet._keys[match] : false;
                if (!okByOffset && !okByKey) return match;
            }
            return formatNumberWithCommas(match);
        });
        tf.contents = replaced;
    } catch (e) {
        // 失敗時はスキップ / Skip on failure
    }
}

// 上から順の採番用ヘルパー / Helpers to rank frames from top to bottom
/**
 * Get top coordinate; return null if unavailable.
 * 上端座標を取得。取得できない場合は null を返す（最後尾へ回すため）。
 */
function getTop(tf) {
    try {
        var gb = tf && tf.geometricBounds;
        var v = (gb && gb.length > 1) ? gb[1] : null;
        return (typeof v === 'number' && isFinite(v)) ? v : null;
    } catch (e) {
        return null;
    }
}
/**
 * Get left coordinate; return null if unavailable.
 * 左端座標を取得。取得できない場合は null（同一行の並び比較が不可能なとき）。
 */
function getLeft(tf) {
    try {
        var gb = tf && tf.geometricBounds;
        var v = (gb && gb.length > 0) ? gb[0] : null;
        return (typeof v === 'number' && isFinite(v)) ? v : null;
    } catch (e) {
        return null;
    }
}

function main() {
    // ダイアログ作成 / Create dialog for target selection
    var dlg = new Window('dialog', LABELS.mainTitle[lang]);
    setDialogOpacity(dlg, DIALOG_OPACITY);
    shiftDialogPosition(dlg, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);

    var panel = dlg.add('panel', undefined, LABELS.panelTarget[lang]);
    panel.orientation = 'column';
    panel.alignChildren = 'left';
    panel.margins = [15, 20, 15, 10];

    var rbSelection = panel.add('radiobutton', undefined, LABELS.rbSelection[lang]);
    var rbDocument = panel.add('radiobutton', undefined, LABELS.rbDocument[lang]);
    rbSelection.value = true;

    // 新規パネル追加：除外オプション
    var excludePanel = dlg.add('panel', undefined, LABELS.panelExclude[lang]);
    excludePanel.orientation = 'column';
    excludePanel.alignChildren = 'left';
    excludePanel.margins = [15, 20, 15, 10];

    var cbExcludeYears = excludePanel.add('checkbox', undefined, LABELS.exYears[lang]);
    cbExcludeYears.value = true;
    var cbExcludePostalCodes = excludePanel.add('checkbox', undefined, LABELS.exPostal[lang]);
    cbExcludePostalCodes.value = true;
    var cbExcludeSlashAny = excludePanel.add('checkbox', undefined, LABELS.exSlash[lang]);
    cbExcludeSlashAny.value = true;
    var cbExcludePhone = excludePanel.add('checkbox', undefined, LABELS.exPhone[lang]);
    cbExcludePhone.value = true;
    var cbExcludeCC = excludePanel.add('checkbox', undefined, LABELS.exCC[lang]);
    cbExcludeCC.value = true;
    var cbExcludeMAC = excludePanel.add('checkbox', undefined, LABELS.exMAC[lang]);
    cbExcludeMAC.value = true;
    var cbExcludeVehicle = excludePanel.add('checkbox', undefined, LABELS.exVehicle[lang]);
    cbExcludeVehicle.value = true;

    dlg.alignChildren = 'fill';
    // Add a group to hold the buttons horizontally and center them
    var buttonGroup = dlg.add('group');
    buttonGroup.orientation = 'row';
    buttonGroup.alignment = 'center';
    // Add Cancel button first (left), then OK button (right)
    var cancelBtn = buttonGroup.add('button', undefined, LABELS.cancel[lang], {
        name: 'cancel'
    });
    var okBtn = buttonGroup.add('button', undefined, LABELS.ok[lang], {
        name: 'ok'
    });

    var targetMode = null;
    if (dlg.show() == 1) { // OK pressed
        targetMode = rbSelection.value ? 'selection' : 'document';
    } else {
        targetMode = null;
    }
    // 件数からリストの高さを概算 / Auto height for list by item count
    function calcListHeightByCount(n) {
        var rows = Math.max(LIST_MIN_ROWS, Math.min(n, LIST_MAX_ROWS));
        var h = LIST_HEADER_H + LIST_ROW_H * rows + LIST_PADDING;
        return Math.min(h, LIST_MAX_H);
    }

    // --- Preview & select target numbers ---
    var frames = [];
    if (targetMode === 'selection') {
        var hasSelection = app.activeDocument.selection && app.activeDocument.selection.length > 0;
        if (hasSelection) {
            frames = collectTextFramesFromSelection(app.activeDocument.selection);
        }
    } else if (targetMode === 'document') {
        frames = collectTextFramesFromDocument(app.activeDocument);
    }

    // フレームを上→下（同列は左→右）で並べ、ランクを付与 / Rank frames by visual order
    var rankByFrameIndex = {};
    if (frames && frames.length > 0) {
        var idxs = [];
        for (var ii = 0; ii < frames.length; ii++) idxs.push(ii);
        idxs.sort(function(a, b) {
            var ta = getTop(frames[a]);
            var tb = getTop(frames[b]);
            // If either top is missing, push that frame to the end
            if (ta === null && tb === null) {
                // fall through to left comparison
            } else if (ta === null) {
                return 1; // a to the end
            } else if (tb === null) {
                return -1; // b to the end
            }
            var topOrder = ORDER_TOP_DESC ? (tb - ta) : (ta - tb); // true: 上→下, false: 下→上
            if (topOrder !== 0) return topOrder;
            var la = getLeft(frames[a]);
            var lb = getLeft(frames[b]);
            if (la === null && lb === null) return 0;
            if (la === null) return 1;
            if (lb === null) return -1;
            return la - lb; // 同一行は左→右
        });
        for (var r = 0; r < idxs.length; r++) {
            rankByFrameIndex[idxs[r]] = r + 1; // 1 origin
        }
    }

    // --- Step 2: プレビュー検出の関数化 ---
    // Collect numeric candidates from a raw string for preview (token pass + optional plain fallback)
    function _collectPreviewCandidates(trimmed, opts, postalOffsetsPreview, runPlainIfNone) {
        var out = []; // {offRaw, offNorm, token, key}
        var seen = {};
        function pushUnique(offRaw, offNorm, token, key) {
            var k = offNorm + '|' + key;
            if (!seen[k]) {
                out.push({ offRaw: offRaw, offNorm: offNorm, token: token, key: key });
                seen[k] = true;
            }
        }
        var normalizedWhole = trimmed.replace(/[，,]/g, "");

        // Pass 1: token pattern (commas allowed, including malformed)
        var tokenRe = /[+\-＋－]?(?:[0-9０-９]{1,3}(?:[，,][0-9０-９]+)+|[0-9０-９]+)(?:[\.．][0-9０-９]+)?/g;
        var m;
        while ((m = tokenRe.exec(trimmed)) !== null) {
            var token = m[0];
            var offRaw = m.index;
            var commasBefore = (trimmed.substring(0, offRaw).match(/[，,]/g) || []).length;
            var offNorm = offRaw - commasBefore;
            var tokenNoComma = token.replace(/[，,]/g, "");

            // Postal exclusions (offset map + local fallback)
            if (opts.excludePostalCodes) {
                if (postalOffsetsPreview && postalOffsetsPreview[offNorm]) continue;
                var startRaw = Math.max(0, offRaw - 5);
                var endRaw = Math.min(trimmed.length, offRaw + token.length + 6);
                var aroundRaw = trimmed.substring(startRaw, endRaw);
                var aroundRawNorm = toHalfWidthDigitsHyphen(aroundRaw);
                if (/(?:^|[^\d])(?:〒\s*)?\d{3}-\d{4}(?!\d)/.test(aroundRawNorm)) continue;
            }
            if (!isEligibleNumericMatch(tokenNoComma, offNorm, normalizedWhole, opts)) continue;

            // Only include if it will change (add/correct commas) and integer part >= 4
            var t = token; var sign = '';
            if (t.charAt(0) === '+' || t.charAt(0) === '-' || t.charAt(0) === '＋' || t.charAt(0) === '－') { sign = t.charAt(0); t = t.substring(1); }
            var parts = t.split(/[\.．]/);
            var intPartRaw = parts[0].replace(/[，,]/g, '');
            if (intPartRaw.length < 4) continue;
            var fracPart = parts.length > 1 ? ('.' + parts[1]) : '';
            var expected = sign + insertCommasToDigits(intPartRaw) + fracPart;
            if (expected === token) continue; // already correct

            pushUnique(offRaw, offNorm, token, tokenNoComma);
        }

        if (runPlainIfNone) {
            // Pass 2: plain digits (fallback)
            var plainNumRe = /[0-9０-９]{4,}(?:[\.．][0-9０-９]+)?/g; var m2;
            while ((m2 = plainNumRe.exec(trimmed)) !== null) {
                var token2 = m2[0];
                var offRaw2 = m2.index;
                var commasBefore2 = (trimmed.substring(0, offRaw2).match(/[，,]/g) || []).length;
                var offNorm2 = offRaw2 - commasBefore2;
                var tokenNoComma2 = token2;
                if (opts.excludePostalCodes) {
                    if (postalOffsetsPreview && postalOffsetsPreview[offNorm2]) continue;
                    var s2 = Math.max(0, offRaw2 - 5);
                    var e2 = Math.min(trimmed.length, offRaw2 + token2.length + 6);
                    var around2 = trimmed.substring(s2, e2);
                    var around2Norm = toHalfWidthDigitsHyphen(around2);
                    if (/(?:^|[^\d])(?:〒\s*)?\d{3}-\d{4}(?!\d)/.test(around2Norm)) continue;
                }
                if (!isEligibleNumericMatch(tokenNoComma2, offNorm2, normalizedWhole, opts)) continue;
                var intLen2 = tokenNoComma2.split(/[\.．]/)[0].replace(/^[+\-＋－]/, '').length;
                if (intLen2 < 4) continue;
                pushUnique(offRaw2, offNorm2, token2, tokenNoComma2);
            }
        }
        return out;
    }

    // Build preview list of eligible numeric substrings (add-mode only)
    var previewEntries = []; // {frameIndex, offset, length, text, key}
    // 既存の opts を置き換え
    var opts = {
        excludeYears: cbExcludeYears.value,
        excludePostalCodes: cbExcludePostalCodes.value,
        excludeSlashAny: cbExcludeSlashAny.value,
        excludePhoneNumbers: cbExcludePhone.value,
        excludeCreditCards: cbExcludeCC.value,
        excludeMAC: cbExcludeMAC.value,
        excludeVehicle: cbExcludeVehicle.value
    };
    for (var fi = 0; fi < frames.length; fi++) {
        var raw = frames[fi].contents;
        if (typeof raw !== 'string') continue;
        var trimmed = raw.replace(/^\s+|\s+$/g, "");
        var normalizedWhole = trimmed.replace(/[，,]/g, "");
        var postalOffsetsPreview = cbExcludePostalCodes.value ? buildPostalOffsetsFromNormalized(normalizedWhole) : null;
        var cand = _collectPreviewCandidates(trimmed, opts, postalOffsetsPreview, /*runPlainIfNone*/ true);
        for (var ci = 0; ci < cand.length; ci++) {
            previewEntries.push({
                frameIndex: fi,
                offset: cand[ci].offNorm,
                length: cand[ci].key.length,
                text: cand[ci].token,
                key: cand[ci].key
            });
        }
    }

    var selectedOffsetsByFrame = {};
    if (previewEntries.length > 0) {
        var dlg2 = new Window('dialog', LABELS.dlg2Title[lang]);
        setDialogOpacity(dlg2, DIALOG_OPACITY);
        shiftDialogPosition(dlg2, DIALOG_OFFSET_X, DIALOG_OFFSET_Y);
        dlg2.orientation = 'column';
        dlg2.alignChildren = 'fill';

        var note = dlg2.add('statictext', undefined, LABELS.dlg2Note[lang]);
        note.alignment = 'fill';

        // Sort previewEntries by rank (visual order) ascending before populating the listbox
        var entriesWithRank = [];
        for (var pi = 0; pi < previewEntries.length; pi++) {
            var e = previewEntries[pi];
            var rank = rankByFrameIndex && rankByFrameIndex[e.frameIndex] ? rankByFrameIndex[e.frameIndex] : (e.frameIndex + 1);
            entriesWithRank.push({
                entry: e,
                rank: rank
            });
        }
        entriesWithRank.sort(function(a, b) {
            return a.rank - b.rank;
        });
        // --- Patch: previewUsed flag and framesInPreview map ---
        var previewUsed = entriesWithRank.length > 0;
        var framesInPreview = {};
        for (var t = 0; t < entriesWithRank.length; t++) {
            framesInPreview[entriesWithRank[t].entry.frameIndex] = true;
        }

        var list = dlg2.add('listbox', undefined, [], {
            multiselect: true,
            numberOfColumns: 3,
            showHeaders: true,
            columnTitles: [LABELS.colSelect[lang], LABELS.colFrame[lang], LABELS.colValue[lang]]
        });
        list.preferredSize = [PREVIEW_LIST_WIDTH, calcListHeightByCount(entriesWithRank.length)];
        list.columnWidths = PREVIEW_COL_WIDTHS;
        for (var pi = 0; pi < entriesWithRank.length; pi++) {
            var e = entriesWithRank[pi].entry;
            var rank = entriesWithRank[pi].rank;
            var row = list.add('item', '✓'); // first column: check mark
            row.subItems[0].text = '#' + rank; // second column: visual order number
            row.subItems[1].text = e.text; // third column: value
            row.helpTip = e.text; // show full value on hover
        }
        // Select all by default
        for (var si = 0; si < list.items.length; si++) list.items[si].selected = true;

  

        // Add onChange event to toggle check mark in first column
        list.onChange = function() {
            for (var i = 0; i < list.items.length; i++) {
                list.items[i].text = list.items[i].selected ? '✓' : '';
            }
        };

        var grp2 = dlg2.add('group');
        grp2.alignment = 'center';
        grp2.orientation = 'row';
        // Select All button (left of Cancel)
        var btnAll = grp2.add('button', undefined, LABELS.selectAll[lang]);
        btnAll.preferredSize = [24, 24]; // width, height
        btnAll.onClick = function(){
            for (var i = 0; i < list.items.length; i++) {
                list.items[i].selected = true;
                list.items[i].text = '✓';
            }
        };
        // Cancel and OK buttons
        var cancel2 = grp2.add('button', undefined, LABELS.cancel[lang], { name: 'cancel' });
        var ok2 = grp2.add('button', undefined, LABELS.ok[lang], { name: 'ok' });
        if (dlg2.show() != 1) {
            // User cancelled preview, abort processing
            frames = [];
        } else {
            for (var li = 0; li < list.items.length; li++) {
                if (list.items[li].selected) {
                    var rowData = entriesWithRank[li].entry;
                    if (!selectedOffsetsByFrame[rowData.frameIndex]) selectedOffsetsByFrame[rowData.frameIndex] = {};
                    if (!selectedOffsetsByFrame[rowData.frameIndex]._keys) selectedOffsetsByFrame[rowData.frameIndex]._keys = {};
                    selectedOffsetsByFrame[rowData.frameIndex][rowData.offset] = true;
                    selectedOffsetsByFrame[rowData.frameIndex]._keys[rowData.key] = true;
                }
            }
            // --- Patch: Block all numbers in frames that had preview entries but nothing was selected ---
            for (var fi in framesInPreview) {
                if (!selectedOffsetsByFrame[fi]) selectedOffsetsByFrame[fi] = {}; // empty set blocks all
            }
        }
    }

    // 実行：選択範囲がある場合のみ / Execute only when there is a selection or whole document based on dialog
    if (targetMode !== null && app.documents.length > 0) {
        for (var i = 0; i < frames.length; i++) {
            // Patch: if preview dialog was used, always pass a set (possibly empty) to block unselected, else allow all
            var setForFrame = previewEntries.length > 0 ? (selectedOffsetsByFrame[i] || {}) : null;
            formatTextFrameIfNumeric(
                frames[i],
                cbExcludeYears.value,
                cbExcludePostalCodes.value,
                cbExcludeSlashAny.value,
                setForFrame,
                cbExcludePhone.value,
                cbExcludeCC.value,
                cbExcludeMAC.value,
                cbExcludeVehicle.value
            );
        }
    }
}

main();