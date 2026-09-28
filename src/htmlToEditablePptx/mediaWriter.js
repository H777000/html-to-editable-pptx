import { renderBackgroundImage, serializeSvgToDataUri } from './imageExport';

export async function writeRasterImage({
    pptSlide,
    element,
    position,
    preserveTransform = false,
    helpers,
}) {
    if (!position) return false;
    const { getRotateDeg, hasComplexTransform, restoreUnrotatedPosition } = helpers;
    const style = element.ownerDocument.defaultView.getComputedStyle(element);
    const rotate = preserveTransform || position.clipped || hasComplexTransform(style)
        ? 0 : getRotateDeg(style);
    const finalPosition = rotate ? restoreUnrotatedPosition(element, position) : position;
    const inlineTransform = element.style?.transform;
    try {
        if (rotate && element.style) element.style.transform = 'none';
        const cropX = Math.max(0, finalPosition.visibleRect.left - finalPosition.rect.left);
        const cropY = Math.max(0, finalPosition.visibleRect.top - finalPosition.rect.top);
        const captureStyle = finalPosition.clipped ? {
            width: `${finalPosition.rect.width}px`, height: `${finalPosition.rect.height}px`, margin: '0',
            transform: `translate(${-cropX}px, ${-cropY}px)${style.transform && style.transform !== 'none' ? ` ${style.transform}` : ''}`,
            transformOrigin: style.transformOrigin || 'top left',
        } : undefined;
        const data = await renderBackgroundImage(element, {
            w: finalPosition.clipped ? finalPosition.visibleRect.width : finalPosition.rect.width,
            h: finalPosition.clipped ? finalPosition.visibleRect.height : finalPosition.rect.height,
        }, { backgroundColor: null, ...(captureStyle ? { captureStyle } : {}) });
        pptSlide.addImage({
            data, x: finalPosition.x, y: finalPosition.y, w: finalPosition.w, h: finalPosition.h,
            ...(rotate ? { rotate } : {}),
        });
        return true;
    } catch (error) {
        console.warn('[AiPPT] 图片导出失败，已跳过:', error);
        return false;
    } finally {
        if (rotate && element.style) element.style.transform = inlineTransform;
    }
}

export async function writeSvg({ pptSlide, svg, position, helpers }) {
    if (!position) return false;
    const {
        getElementPosition, getRotateDeg, getShapeType, shapeTypes,
        restoreUnrotatedPosition, writeRasterImage,
    } = helpers;
    const style = svg.ownerDocument.defaultView.getComputedStyle(svg);
    const parent = svg.parentElement;
    const parentPosition = parent && getElementPosition(parent);
    const parentStyle = parent && svg.ownerDocument.defaultView.getComputedStyle(parent);
    const isSmallCircularBadgeIcon = parentPosition && parentStyle
        && Math.max(position.rect.width, position.rect.height) <= 64
        && Math.max(parentPosition.rect.width, parentPosition.rect.height) <= 144
        && getShapeType(parent, parentStyle, parentPosition) === shapeTypes.ellipse;
    if (isSmallCircularBadgeIcon || position.clipped) {
        return writeRasterImage({ pptSlide, element: svg, position, preserveTransform: position.clipped, helpers });
    }
    const rotate = getRotateDeg(style);
    const finalPosition = rotate ? restoreUnrotatedPosition(svg, position) : position;
    const requiresRasterization = Boolean(svg.querySelector(
        'foreignObject, image, filter, mask, pattern, animate, animateTransform, set, use',
    ));
    if (requiresRasterization) return writeRasterImage({ pptSlide, element: svg, position, helpers });
    try {
        pptSlide.addImage({
            data: serializeSvgToDataUri(svg, { width: finalPosition.rect.width, height: finalPosition.rect.height }),
            x: finalPosition.x, y: finalPosition.y, w: finalPosition.w, h: finalPosition.h,
            ...(rotate ? { rotate } : {}),
        });
        return true;
    } catch (error) {
        console.warn('[AiPPT] SVG 原生导出失败，正在使用图片兜底:', error);
        return writeRasterImage({ pptSlide, element: svg, position, helpers });
    }
}
