# Find and adjust Tate-Chu-Yoko

[![Direct](https://img.shields.io/badge/Direct%20Link-FindEditTCY.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/FindEditTCY.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/FindEditTCY.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Collects the Tate-Chu-Yoko characters in the selected text and lists each place in a dialog box.
- Selecting a line selects that place in the document, so you can rewrite its string and adjust its glyph width and tracking on the spot.

### Features

- Tate-Chu-Yoko Text (left)
  - **List**: consecutive Tate-Chu-Yoko characters count as one place, listed in the order found. Identical strings still get one line each. The count is shown above
  - **Selecting a place**: selecting a line selects those characters in the document, scrolling to the text frame when it is off screen
  - **Replace with**: the new string for the selected place. While you type, the line shows "12 → 10". Leave it empty to delete the place (Tate-Chu-Yoko is kept)
  - **Apply**: applies the selected place's rewrite right away. Rewrites not applied this way are applied on OK
- Glyph Width (upper right)
  - **Glyph width**: Default Form, Full Widths, Half Widths, Third Widths, Quarter Widths. Selecting a line checks its current form; clicking one applies it right away
  - **Auto Apply**: applies a glyph width to every place in the list by its length (1 → Default Form, 2 → Half Widths, 3 → Third Widths, 4 → Quarter Widths; 5 or more unchanged)
- Tracking (lower right)
  - **Tracking**: the up/down buttons and arrow keys step by 1, Shift-click snaps to multiples of 10. Changes apply right away
  - **Reset**: resets the selected place's tracking to 0
- Closing the dialog box restores the selection you had before running
- Japanese / English UI switches automatically

### How to use

1. Select text objects that contain Tate-Chu-Yoko (or select a range with the Type tool)
2. Run the script
3. Select a line and adjust the string, glyph width and tracking
4. Click OK

### Notes

- Glyph width, tracking, Auto Apply and Apply take effect immediately, so Cancel does not undo them. To revert, close the dialog box and use Undo
- When a rewrite changes the number of characters, the rewritten part takes the formatting of its first character
- Threaded text is checked as one story. Locked or hidden text is skipped
- The list shows the first 500 lines only

### Article

- [DTP Transit 別館 (Japanese)](https://note.com/dtp_tranist/n/n1144206e9fda)

### Update History

- v1.0.0 (2026-10-16) Initial release
