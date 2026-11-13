# Randomize touch type adjustments

[![Direct](https://img.shields.io/badge/Direct%20Link-AutoTouchType.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/AutoTouchType.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AutoTouchType.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

An "auto touch type" tool that randomizes the baseline shift, scale, rotation, kerning and tracking of each character in the selected text.

### Features

- A seeded RNG keeps the preview stable
- When character rotation is applied, the tracking is corrected automatically for the rotation angle
- A lightweight mode defers the zoom update until the drag finishes

### Usage

1. Select the text.
2. Run the script.
3. Set how much each property varies, check the preview, and click OK.

### Article (note)

https://note.com/dtp_tranist/n/ne6545c4717af

### Update History

- v1.4.1 (2026-11-13): "Exclude kana fonts" also leaves out the TypeBank Ajioka kana series (Komachi, Ryokan, Yukinari, Kodoken and Tsukiji, such as TB Komachi G Std and Komachi-GM)
- v1.4.0 (2026-11-08): Added "Exclude kana fonts" (on by default), which leaves out kana-only fonts without kanji (such as ML Chocolate, ML Tokatsuki Kana, Seiren Mincho Kana, Hommaru Kana and Seimei Kana). Fonts that were left out only because their names contain Kana, such as FOT-Kanae Reisho, now count as Japanese
- v1.3.12 (2026-11-08): Japanese font detection no longer counts the kana-only fonts SDUdai, SDTenga and SDSeimu
- v1.3.11 (2026-11-08): Japanese font detection no longer counts fonts whose names start with HG
- v1.3.10 (2026-11-08): Japanese font detection no longer counts Fontworks Chinese and Korean fonts, Fontworks non-Japanese scripts such as FWArabic, Korean fonts such as BM Dohyeon (배달의민족 도현), Fontworks emoji, kana-only fonts without kanji, Chinese fonts such as PingFang, LiHei and Hiragino Sans GB, Shizuka Gothic (靜呉竹), Tsukimi, or fonts whose names contain Thai or 體
- v1.3.9 (2026-11-05) Moved the version from the title bar to the bottom left of the dialog, small and dimmed (shared part added)
- v1.3.8 (2026-11-02): Each text now keeps its own ransom-note background, so running on another text no longer removes earlier ones. Fixed: characters selected with the Type tool were rejected, a scale amount of 100% or more caused an error, Cancel did not restore auto kerning, and the rotation tracking compensation overwrote the original tracking. "Japanese only" now turns on automatically only when the text contains Japanese, and Random / Reset no longer toggle it. Cancel now also undoes Reset (backgrounds are rebuilt or removed on OK). Ransom-note tracking is always added to the original value (it used to depend on the rotation compensation checkbox). Per-character font changes skip fonts missing from the system. Fixed the baseline default becoming 0 in mm or inches, where the slider now steps by 0.1. Tooltips and the no-Japanese-fonts message now match what the script does
- v1.3.7 (2026-10-04) Japanese labels now end with " :" (half-width space and colon) (shared part update)
- v1.3.6 (2026-10-01): Added space below the button row to match Illustrator's own dialogs
- v1.3.5 (2026-10-01): Unified the window and panel margins and spacing with the shared layout part
- v1.3.4 (2026-09-30): Fixed an error when running with characters selected by the Type tool
- v1.3.3 (2026-09-29): Dialog opacity changed to 98%
- v1.3.2 (2026-09-28): Clip groups are now measured by their mask
- v1.3.2 (2026-09-28): The button row is now built with the shared part
- v1.3.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.3.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.2.9 (2026-09-19): Added the article link and reorganized the internals (nested label definitions, clearer naming, merged duplicated code).
- v1.2.7 (2026-02-20)
