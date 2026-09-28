# Insert a line break after the specified characters

[![Direct](https://img.shields.io/badge/Direct%20Link-titlemaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/titlemaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/titlemaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

- Inserts a line break right after the specified characters in the selected text frames
- The break characters (、 。 〜) are chosen with checkboxes
- Consecutive line breaks are collapsed into one afterwards
- Character alignment is set to the Roman baseline

### Process Flow

1. Choose the break characters in the dialog
2. Walk the contents of each selected text frame
3. Insert a break after each target character and collapse duplicates
4. Apply the Roman baseline

### Script info

- Version: v1.1.2

### Update History

- v1.1.0 (2026-09-27): Added stepper buttons to the number fields. The arrow keys now share the steppers' logic (to the next whole number; Shift to the next multiple of ten)
- v1.1.1 (2026-09-28): The dialog now reopens where it was last closed and moves sideways to avoid covering the selection; opacity unified at 97%
- v1.1.2 (2026-09-29): Dialog opacity changed to 98%
