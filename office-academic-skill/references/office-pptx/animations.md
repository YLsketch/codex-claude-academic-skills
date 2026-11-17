# PowerPoint Transitions and Element Animations

This module adds native PowerPoint transitions and entrance animations to editable PPTX objects. It edits the OOXML package directly and does not render slides as video.

## Capabilities

- Slide transitions: fade, push, wipe, split, cover, dissolve, circle, diamond, plus, wedge, and cut.
- Object entrances: appear, fade, wipe, fly, and zoom.
- Targets: text boxes, AutoShapes, pictures, graphic frames, connectors, and top-level grouped shapes.
- Presenter control: grouped-clicks, on-click, after-previous, and with-previous.

## Important limitation: flattened slides

If a slide contains only one full-slide picture, the script adds the slide transition but skips element entrance animation. Text, diagrams, and charts baked into that picture cannot be targeted separately. Do not describe a whole-slide picture fade as per-element animation. Use an editable source deck or rebuild that slide with separate native objects for true element animation.

## Safe workflow

1. Keep the original PPTX unchanged and always write to a new output path.
2. Run a dry inspection:

   ```bash
   python references/office-pptx/scripts/add_pptx_motion.py input.pptx --dry-run
   ```

3. Review the listed object IDs/names and the object order. Avoid animating logos, headers, footers, page numbers, watermarks, and decorative elements.
4. Apply both slide transitions and object entrances:

   ```bash
   python references/office-pptx/scripts/add_pptx_motion.py input.pptx -o output-animated.pptx --transition fade --effect fade --trigger grouped-clicks --clicks-per-slide 2
   ```

5. For slides that should keep a static cover/divider/closing layout, use `--skip-element-slides 1,2,12`. The transition still applies to those slides.
6. Run `--dry-run` on the output and open it in desktop PowerPoint to verify playback and click order.

## Recommended defaults

| Setting | Default | Guidance |
|---|---:|---|
| Slide transition | fade, 0.4 s | Use one restrained family across the deck |
| Element entrance | fade, 0.45 s | Use for titles, text groups, images, charts, and diagrams |
| Trigger | grouped-clicks | One presenter click starts one content group |
| Click groups | 2 per content slide | Use 1–3 groups for most academic slides |
| Stagger | 0.12 s | Keep within-group cascade subtle |
| Auto advance | Off | Enable only when explicitly requested |
| Static chrome | Excluded | Do not animate logos, headers, footers, page numbers, or decoration |

## CLI options

- `--transition`: none, cut, fade, push, wipe, split, cover, dissolve, circle, diamond, plus, wedge.
- `--transition-duration`: transition duration in seconds; default 0.4.
- `--effect`: none, appear, fade, wipe, fly, zoom.
- `--duration`: object entrance duration; default 0.45 seconds.
- `--stagger`: gap between objects within a click group; default 0.12 seconds.
- `--trigger`: grouped-clicks, on-click, after-previous, with-previous.
- `--clicks-per-slide`: approximate number of click groups; default 2.
- `--group-size`: fixed number of objects per click group; 0 infers from click count.
- `--skip-element-slides`: comma-separated 1-based slide numbers to leave static for element animation.
- `--max-targets`: maximum animated objects per slide; default 12.
- `--dry-run`: list target objects without writing output.
- `--replace-existing`: replace existing transition/timing XML. Not used by default.
- `--allow-full-slide-image`: allow a whole-slide image entrance. This is not per-element animation.
- `--auto-advance`: optional slide auto-advance delay; off by default.
- `--overwrite-output`: allow replacing an existing output file.

Existing transition and timing XML is preserved by default. If a slide already has a transition, the script does not add a second one; if it already has timing, the script does not overwrite it. Use `--replace-existing` only with explicit user authorization.

## QA checklist

- The source PPTX is unchanged and the output is a separate file.
- Slide order/count and existing content are unchanged.
- The PPTX package passes ZIP integrity checks and edited slide XML remains well-formed.
- The intended native objects are the animation targets; flattened full-slide images are reported and skipped.
- Existing animations are preserved unless replacement was approved.
- The output opens in desktop PowerPoint without a repair prompt.
- Actual transition and click-by-click playback is checked in PowerPoint. XML validity alone does not prove playback compatibility.

Prefer fade, appear, and wipe for broad compatibility. WPS, Keynote, and LibreOffice may remap some effects; PowerPoint desktop is the reference playback target.
