# SmartShapeMaker-v2

[![Direct](https://img.shields.io/badge/Direct%20Link-SmartShapeMaker--v2.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/shape/SmartShapeMaker-v2.jsx)

[![English](https://img.shields.io/badge/README-English-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-en/SmartShapeMaker-v2.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### 更新日 / Updated

- 20260319

Main Features:
- Specify number of sides (0 = Circle, 3/4/5/6/8, or custom with slider)
- Circle panel:
  - Superellipse option (only when sides = 0)
  - Superellipse shape control (exponent)
  - Anchor Points panel (2 / 3 / 4 / 5 / 6)
  - When Superellipse is ON:
    - Rotate is forced OFF
    - Live Shape is forced OFF
    - Anchor Points panel is dimmed
- Star panel:
  - Star option + Pentagram option (side-by-side)
  - Inner radius input + 0–100 slider
  - Inner radius controls are dimmed when Star is OFF
  - When Pentagram is ON, Rotate is forced OFF
- Triangle direction options (Left / Right / Down) when sides = 3
- Width (size) panel with unit display
- Rotation panel:
  - Auto angle is used when Rotate is OFF (Circle=45°, Polygon=360/(sides*2))
  - When sides = 3 and Rotate is enabled, Triangle direction defaults to “Down” (60°)
  - Arrow-key editing supported
- Reuleaux-style option (odd-sided polygons only)
  - Adjustable appearance amount (0–200%)
  - Amount resets to 100% when Reuleaux is enabled
- Options panel:
  - Live Shape conversion (Convert to Shape) on finalize
  - Split at Anchor Points (creates open stroked segments)
- Dialog opacity and position are restored within the current Illustrator session
- Preview does not pollute Undo history; final result can be undone in a single step
- View Zoom slider above OK / Cancel

Keyboard Shortcuts:
- E : Circle (0)
- A : Toggle Rotate
- S : Toggle Star
- P : Toggle Pentagram
- L : Triangle Left (also sets sides = 3)
- R : Triangle Right (also sets sides = 3)
- B : Triangle Down (also sets sides = 3)
- D : Toggle Split at Anchor Points

Usage Flow:
1. Set sides, width, star/circle options, rotation, and options in the dialog
2. Preview updates in real-time
3. Click OK to finalize the preview object at the artboard center

Original Idea: Seiji Miyazawa (Sankai Lab)
https://x.com/onthehead/status/2007350198721483172

コーナースムージング
黒野 真吾さん
https://note.com/shingokurono/n/n348a3e73a465

### スクリプト情報

- バージョン: v1.9
