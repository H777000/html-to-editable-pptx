import { cropRasterImage } from './imageExport';

// Keeps browser-capture output independent from DOM isolation. Callers supply
// an already-rendered slide image and a geometry result, then this module owns
// the raster crop and the matching PptxGenJS image write.
export async function addCroppedBrowserCapture({
    pptSlide,
    fullSlideData,
    position,
    slideRect,
    document,
}) {
    const data = await cropRasterImage(fullSlideData, {
        sourceWidth: slideRect.width,
        sourceHeight: slideRect.height,
        cropX: position.visibleRect.left - slideRect.left,
        cropY: position.visibleRect.top - slideRect.top,
        width: position.visibleRect.width,
        height: position.visibleRect.height,
    }, document);
    pptSlide.addImage({ data, x: position.x, y: position.y, w: position.w, h: position.h });
}
