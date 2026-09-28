# Reproducible comparison

Two newly authored, self-contained slides are included in `examples/`:

| Slide | This exporter MAE | dom-to-pptx 2.1.2 MAE | Observation |
| --- | ---: | ---: | --- |
| Dark radial gradient | 7.96 | 161.56 | This exporter preserves the dark background. |
| Light cards | 5.66 | 4.38 | The comparison tool is closer and uses fewer pictures. |

Run `npm run benchmark` after `npm run build`. It uses the same Chrome viewport (1280 × 720) and 10 × 5.625 inch PPTX size for both tools. LibreOffice renders PPTX to PDF; `pdftoppm` produces 1280 × 720 PNGs. RGB mean absolute error (MAE) is computed against Chrome's HTML screenshot. Lower is better; it is not a perceptual score or a percent fidelity claim.

The test also counts OOXML shape and picture objects and saves `public-metrics.json`. Shape count includes text boxes and decoration shapes; picture count includes ordinary images as well as rasterized effects. Those counts alone do not prove every element is editable.

The dark slide emphasizes a known strength, so the light slide is included to show a case where the comparison tool does better. This small benchmark does not establish a global ranking. Visual results can change across fonts, Chrome versions, LibreOffice versions, operating systems, and PowerPoint itself.
