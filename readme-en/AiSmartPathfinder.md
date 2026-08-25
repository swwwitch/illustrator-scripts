# AiSmartPathfinder

[![Direct](https://img.shields.io/badge/Direct%20Link-AiSmartPathfinder.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/fx/AiSmartPathfinder.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/AiSmartPathfinder.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

A persistent palette that applies Pathfinder operations to the selected objects.

Clicking an icon delegates the operation to the main engine and runs it immediately.

### Features

- Two tabs: Basic and Special
- Mode: exclusive radio buttons for the output mode (shortcuts P / C / F)
- Shape mode: Unite, Minus Front, Intersect and Exclude (Adobe Pathfinder commands 0–3)
- Operations run through BridgeTalk on the main engine

### Usage

1. Select the objects you want to combine.
2. Run the script to open the palette.
3. Choose the mode and click the operation icon.

### Notes

- The palette is a `#targetengine` resident script: after editing the code, **close the palette before running it again**, otherwise the old code keeps running.

### Update History

- v1.1.0
