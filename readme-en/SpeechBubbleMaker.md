# Draw an underline or box with a speech-bubble tail around text

[![Direct](https://img.shields.io/badge/Direct%20Link-SpeechBubbleMaker.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/SpeechBubbleMaker.jsx)

[![Japanese](https://img.shields.io/badge/README-Japanese-4b8bbe.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/readme-ja/SpeechBubbleMaker.md)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Overview

Draws an underline, an L shape, a bracket, a box or top and bottom lines with a speech-bubble tail (a V, a slash, a taper, a U, a loop, a horn or a thought bubble) around the selected text. The line can go on any side, with color, cap, rounded corners and a shadow. Settings are previewed in a dialog box and can be saved as presets.

### Features

The dialog box has Presets centered at the top, the Tail and Text Options panels on the left below it, and the Line/Frame panel on the right. Preview sits at the left end of the button row.

#### Presets

- Picking a preset switches the tail and line/frame settings at once.
- Twenty-four built-in presets are included (simple speech bubble, wide curly, underline with a V, a V down right, slashes down right and left; a slash down right with no gap; a U; a loop; bracket with a slash; rounded bracket with a V; curly bracket; L shape with a V down left; box with a slash; rounded box with a slash; capsule box; box with a tail at the bottom-right or bottom-left corner; box with a V and shadow; capsule box with gaps; capsule box with a dotted slash; thought bubble; rounded box with a horn; top and bottom lines with a slash down left). Built-in presets size themselves to the font size of the first character.
- The save icon stores the current settings under a name; the trash icon deletes the selected saved preset (built-in presets cannot be deleted, and their names cannot be reused). Saved presets keep their sizes as saved and survive quitting Illustrator.
- Changing a setting by hand clears the preset selection.

#### Tail

- Shape: seven choices laid out in three columns.
  - V shape bends the line into a tail (it stays a single path). Slash draws a slash out of the line. Taper narrows out of the line in curves to a point, like the middle of a curly bracket in AiCurlyBracketMaker.
  - U shape bends the line into a U (Height is its depth). Loop curves out of the line into a small loop (Height reaches the bottom of the loop; a wide Gap makes the line sweep like a wave).
  - Horn lays a filled tail over the line, wide at the root and pointed at the end, like a stroke width profile (the line is not cut). Gap is the root width.
  - Thought adds a small bump with rounded joins to the line and puts a circle beyond it, as in a thought bubble. Gap is the bump's width and Height reaches the bottom of the circle; the circle is sized to the bump.
- Direction: Straight, Right or Left (on the bottom side these look straight down, down right and down left).
  - For every shape but a slash, Straight puts the tip (the circle of a thought, the loop of a loop) at the middle of the base; Right puts it Width to the right of the base's left end; Left puts it Width to the left of the base's right end.
  - A slash leans to the right or left at Angle; Straight makes it leave the line at a right angle.
- Angle (slash with Direction Right or Left): measured from the line (default 45°, 5–90°; 90° is at a right angle to it). The slash's horizontal run follows Height and Angle.
- Root (slash with Direction Right or Left): Center starts the slash at the middle of the gap, apart from the line; Edge starts it at the edge of the gap, joined to the line.
- Nudge X / Nudge Y (detached slash: Root is Center, or Direction is Straight): shift the slash a little (the stepper moves by 0.5). Positive X moves it along the line (right, or down on the left or right side); positive Y moves it outward, the way the tail points, and negative toward the text.
- Position: from 0% (start of the line) to 100% (end), in the field or with the slider below it (hold Shift for 10% steps); default 50%. A tail that would run past the ends is pulled inward.
- Gap: how much of the line is cut away at the tail; for a V or taper, the width of its base. A slash with 0 keeps the line whole.
- Width (any shape but a slash, with Direction Right or Left): how far the tip moves sideways.
- Height: how far the tail reaches out from the line.
- Dot at the tip (V, slash, taper or U): puts a filled dot three times the weight across at the tip of the tail. Saved in presets too.

#### Line/Frame

- Shape: Underline draws a line below the glyphs. L shape turns up only the end nearer the tail toward the text (the start when Position is 50% or less). Bracket turns both ends of the line up toward the text. Box encloses the glyphs in a rectangle. Top & bottom draws lines on both sides of the glyphs, with the tail on the line on the chosen side. A bracket whose Corner radius equals its Ticks gets curly-bracket ends.
- Side: Top, Bottom, Left or Right, laid out around a cross. Each time the dialog box opens, Bottom is picked when the first text is horizontal and Right when it is vertical (the same happens when picking a preset that has no side of its own).
  - On Top, the tail points up (Right goes up to the right, Left up to the left).
  - On Left or Right the line starts at the top, Right leans toward the bottom and Left toward the top.
- Distance: from the glyphs to the line; the top and bottom padding of a box.
- Extension: how far the line runs past the glyphs (negative values shorten it); the side padding of a box.
- Ticks (L shape or bracket): length of the turned-up ends.
- Corner radius (L shape, bracket or box): 0 keeps the corners sharp. When the tail overlaps a bottom corner's rounding (a V pointing Left at 0%, for example), that corner stays sharp and the tail comes out of it (a capsule keeps its tail inside the rounding instead).
- Capsule (L shape, bracket or box): uses the largest corner radius the sides allow, so the ends of a box become half circles (a bracket gets the Ticks length as its radius). Corner radius is unused while it is on. Saved in presets too.
- Shadow (box): adds a second line this far out on the right and bottom (0 for none). Its ends line up with where the box's corner rounding starts; the tail's right edge is offset in parallel and stops at the tip's height, and the bottom line stops short of the tail's left edge.
- Side gaps (box): cuts this much out of the middle of the box's left and right sides (0 for none). On a capsule the half circles are cut at their middle.
- Weight: stroke weight of the line and tail.
- Color: Text (default) uses the fill of each text's first character (black when it has no fill); the other radio button (custom, shown with just a swatch) uses the swatch's color. Clicking the swatch opens Illustrator's Color Picker and switches to custom.
- Cap: Butt (default) or Round. Butt keeps the corners pointed; Round rounds the corners too.

#### Text Options

- Use the frame of area type (on by default; available when area type is selected): area type is drawn along its frame instead of its glyphs (Distance and Extension count as 0; the two fields are dimmed when every selected text is area type). Point type and type on a path are still measured by their glyphs. Saved in presets too.
- Center as area type (box only): turns the text into area type the size of the box, centered horizontally and vertically. Point type is converted; area type gets its frame resized and moved (type on a path is left alone).
- Group with text: groups the line and tail together with the text. When off, the line and tail are grouped on their own, right in front of the text.

#### Other

- Preview (left end of the button row): shows the result on the canvas (on by default).
- While the script runs, Hide Edges and Hide Bounding Box are toggled to make the preview easier to see, and toggled back at the end.
- Lengths use the General unit under Preferences > Units; the weight uses the Stroke unit.
- The first time, the sizes are derived from the font size of the first character. The settings used with OK (custom color included) are remembered until Illustrator quits and come back next time (only Side is picked again from the text each time).

### Usage

1. Select text (several objects are fine, including text inside groups).
2. Run the script, pick a preset or set the tail and line/frame shapes and sizes, and click OK.

### Notes

- One set of lines is drawn around the bounds of all the glyphs of the text (the frame, when area type is measured by its frame). With several lines, an underline goes below the last line (outside the outermost column on the left or right).
- The area type conversion and vertical centering run after OK, so the preview does not show them.
- Locked or hidden text, and text containing only spaces, is skipped.

### Article

[Draw underlines and boxes with a speech-bubble tail on text with an Illustrator script (Japanese)](https://note.com/dtp_tranist/n/n8849614c960c)

### Update History

- v1.0.0 (2026-11-02): Initial release
- v1.0.1 (2026-11-02): Added the built-in preset Simple speech bubble and grouped the built-in presets with separators (staple, underlines, brackets, boxes). Renamed the Options panel to Text Options and Center the text as area type to Center as area type. Moved Shadow below Corner radius and made the save icon slightly smaller. The position slider moves in 10% steps with Shift held. Fixed the Color Picker sometimes opening in grayscale
- v1.0.2 (2026-11-02): Added the built-in preset Wide curly (an underline with a taper and round caps)
- v1.1.0 (2026-11-03): Added Nudge X and Nudge Y for a detached slash, and Capsule below Corner radius; a tail overlapping a bottom corner's rounding now leaves that corner sharp, plus the built-in presets Capsule box, Underline, slash down right, no gap, and boxes with a tail at the bottom-right or bottom-left corner. Renamed Direction to Straight, Right and Left (V renamed V shape) and revised tooltips for the Side setting
- v1.1.1 (2026-11-03): Picking Direction (or Shape) for a slash now sets Nudge X to minus half the weight for Right and plus half for Left; built-in presets without a Nudge X do the same
- v1.1.2 (2026-11-03): The built-in presets Underline, slash down right and Underline, slash down left now use Round cap
- v1.2.0 (2026-11-03): Added the tail shapes U shape, Loop, Horn and Thought, and the line shapes L shape and Top & bottom (the tail shapes now sit in three columns). Added Dot at the tip for the tail and Side gaps for a box, plus the built-in presets Underline, U; Underline, loop; L shape, V down left; Capsule box with gaps; Capsule box, slash with a dot; Thought bubble; Rounded box, horn; and Top & bottom, slash down left
- v1.2.1 (2026-11-05) Moved the version from the title bar to the bottom left of the dialog, small and dimmed (shared part added)

