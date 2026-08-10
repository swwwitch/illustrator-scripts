# DistributeDownFromTop

[![Direct](https://img.shields.io/badge/Direct%20Link-DistributeDownFromTop.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/alignment/DistributeDownFromTop.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/DistributeDownFromTop.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Adjusts leading and placement, deciding what to do from the selection in this order.

1. A single text object: adds the Size/Leading value to its leading.
2. Several text objects whose top edges line up: positions stay put and only the leading changes. When the leadings differ they are unified to their average; when they already match, every leading increases by the Size/Leading value.
3. Any other multiple selection (stacked vertically): the topmost object stays fixed and the rest are redistributed downwards at even intervals of the Size/Leading value.

### Notes

- The step comes from the Size/Leading increment preference (`text/sizeIncrement`), converted to points using the display unit.
- Leading is applied by deriving an auto-leading amount (`autoLeadingAmount`, in %) from the target leading, rather than by setting manual leading.
