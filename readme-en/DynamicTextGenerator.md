# DynamicTextGenerator

[![Direct](https://img.shields.io/badge/Direct%20Link-DynamicTextGenerator.jsx-ffcc00.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/jsx/text/DynamicTextGenerator.jsx)

[![Direct](https://img.shields.io/badge/Back%20to%20home-All%20scripts-cccccc.svg)](https://github.com/swwwitch/illustrator-scripts/blob/master/README.md)

---

### Description

- Script that builds a path sized to the selected text and converts it into type on a path.
- The path can be an arch, a circle, or a bow. The generated path has no fill and no stroke (invisible), so only the text is visible.
- A Block mode is also available, which uses no path and instead scales each line's font size so every line matches the widest one.
- Point text, area type, and existing type on a path are all valid targets. The preview is always on, so the result can be checked while the dialog stays open.

### Modes

Chosen with icon buttons. The mode name appears under each icon, and the help tip reads "name: description". **No mode is selected when the dialog opens**, so nothing is converted until one is picked.

| Mode | What it does |
| --- | --- |
| **Block** | Creates no path; scales each line's font size so every line matches the widest one |
| **Circle** | Converts to a closed circular path and flows the text around it, centered at the top |
| **Arch** | Converts to a path that bulges upward |
| **Bow** | Converts to a path that bulges downward |

Options belonging to unselected modes are dimmed. Pressing OK without picking a mode keeps the dialog open and asks for one.

### Block Options

- **Line breaks:**
  - **Keep** (default): keeps the current line breaks and fits each line
  - **At punctuation**: drops the current line breaks and re-breaks after each punctuation mark (`、。，．`)
  - **Line count [3] lines**: drops the current line breaks and re-splits the text into the given number of lines, preferring punctuation within one third of a line from the target position
- With "At punctuation" and "Line count", the font size is **levelled to the largest size in the frame** before the lines are re-cut. This prevents sizes left over from an earlier Block pass from being mixed inside a single line; on uniformly sized text it changes nothing.
- **Leading:** when checked, switches leading to auto with the given ratio (default 100%)

### Arch and Circle Options

- **Curve:** slider (0-100, default 100) for the path depth. 0 is almost straight; 100 gives the roundest curve. In Circle mode it sets how much of the circumference the text occupies (25-95%), i.e. the size of the circle
- **Adjust:** None / Font size (default) / Letter spacing. Works on closed paths such as circles as well
  - Font size: grows the text until it overflows, then shrinks it back to the path ends. Scaling is done by ratio, so mixed character sizes keep their relative differences
  - Letter spacing: keep the font size and adjust the spacing with a coarse pass followed by a fine pass
- **Effect:** Type on a Path effect, picked from a popup menu (Rainbow (default) / Skew / 3D Ribbon / Stair Step / Gravity)
- **Remove line breaks:** on by default. Joins the text into a single line, since it flows along one path

### Common Options

Settings that apply to every mode. Both change the glyph widths, so they are **applied before the widths are fitted**.

- **Kerning:** Keep (default) / Metrics / Optical / Metrics - Roman Only
  - Keep: leaves the kerning settings untouched
  - Choosing Metrics also **turns proportional metrics on** (the other choices turn it off)
  - "Metrics - Roman Only" means metrics for Roman text and monospaced Japanese
- **Tracking:** when the checkbox is on, adds the given value to the existing tracking (-100 to 500; arrow keys adjust, Shift = 10, Option = 0.1). Turning it off resets it to 0. Dimmed while "Adjust: Letter spacing" is selected

### View and Buttons

- **Zoom to selection:** on by default. Moves the view only when the converted result falls outside the visible area; the view is left alone when it is already in sight. When the result does not fit, the view zooms out only — it never zooms in
- **Hidden Characters:** toggles the display of hidden characters, useful for checking where "Line count" or "At punctuation" placed the breaks
- Cancel discards the preview and closes

### Keeping the Text from Being Clipped

- When "Remove line breaks" is on, the path length is measured from the text **as a single joined line**. Measuring the original lines would only span the widest line, leaving the path too short and clipping the text.
- After generation, overset (text that does not fit the path) is detected and the font size is shrunk until it fits, then eased back. Closed paths such as circles are covered too. Nothing is changed when the text already fits.
- Shrinking scales the sizes **by ratio** rather than assigning an absolute size, so texts with mixed character sizes keep their relative differences.

### Workflow

**Arch / Circle / Bow**

1. Collect point text, area type, and type on a path from the selection (recursing into groups)
2. Copy anything that is not point text into plain point text and use that as the source. Lines that an arch or circle had pushed out of sight come back, and area type loses its frame wrapping
3. Measure the source via temporary outlines, create a straight path along the baseline, and bend it into an arc with its Bézier handles (Circle mode builds a closed circular path instead)
4. Create the type on a path, duplicate the character attributes, then apply line-break removal, center justification, kerning, tracking, and the effect
5. Fit to the path with the chosen method, then shrink whatever still does not fit (any path selected together with the text is removed)

**Block**

1. Convert anything that is not point text into point text
2. Re-cut the lines with the chosen method
3. Apply kerning and tracking, then measure every line by its outline width and scale each line to the widest one (fixed leading is scaled along with it)
4. Switch leading to auto when requested

### Not Supported

- No open document
- Selections containing no text
- Empty text, or text whose lines / text ranges cannot be read
- Block requires two or more lines (a single line has nothing to fit to)
- When OK applies nothing, the dialog stays open so the settings can be corrected
- Path fitting is skipped for locked, hidden, or non-editable text

### note

- Article: https://note.com/dtp_tranist/n/nbabfda4b7920
- Original idea: Toshiyuki Takahashi (@gautt) https://note.com/gautt/n/n92f6faeda048

### Update History

- v1.0.0 (20260811): Public release
