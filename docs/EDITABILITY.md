# What the proof clip shows

The 9-second `editability-proof.mp4` uses a newly authored example. Its three frames are:

1. HTML rendered by Chrome.
2. PPTX generated from that HTML and rendered by LibreOffice.
3. The **same PPTX** after one native `<a:t>` text run in `ppt/slides/slide1.xml` was changed from `Keep the design.` to `Change the design.`, then rendered again.

This verifies that the title is stored as editable native PPTX text rather than burned into a slide image. It does not show a person clicking and typing in Microsoft PowerPoint; a real UI recording should be added after testing the published package on a PowerPoint installation.
