import { renderBackgroundImage } from './imageExport';

export async function writeCssBackgroundLayer({ pptSlide, element, position, helpers }) {
    if (!position) return false;
    const { getBorderRadiusCss, getEffectiveOpacity, getRotateDeg, getShapeType,
        restoreUnrotatedPosition, shapeTypes } = helpers;
    const document = element.ownerDocument;
    const style = document.defaultView.getComputedStyle(element);
    const rotate = position.clipped ? 0 : getRotateDeg(style);
    const finalPosition = rotate ? restoreUnrotatedPosition(element, position) : position;
    const isEllipticalDecoration = getShapeType(element, style, finalPosition) === shapeTypes.ellipse;
    const layer = document.createElement('div');
    const width = Math.max(1, Math.round(finalPosition.clipped
        ? finalPosition.visibleRect.width : finalPosition.rect.width));
    const height = Math.max(1, Math.round(finalPosition.clipped
        ? finalPosition.visibleRect.height : finalPosition.rect.height));
    const cropX = Math.max(0, finalPosition.visibleRect.left - finalPosition.rect.left);
    const cropY = Math.max(0, finalPosition.visibleRect.top - finalPosition.rect.top);
    layer.style.cssText = [
        'position:absolute', 'left:0', 'top:0', `width:${width}px`, `height:${height}px`,
        `background-color:${style.backgroundColor}`, `background-image:${style.backgroundImage}`,
        `background-position:${style.backgroundPosition}`, `background-repeat:${style.backgroundRepeat}`,
        `background-size:${style.backgroundSize}`, `background-origin:${style.backgroundOrigin}`,
        `background-clip:${style.backgroundClip}`, `box-sizing:${style.boxSizing}`,
        `border:${style.border}`, `border-radius:${getBorderRadiusCss(style)}`,
        `box-shadow:${isEllipticalDecoration ? 'none' : style.boxShadow}`,
        `opacity:${getEffectiveOpacity(element)}`, 'pointer-events:none',
    ].join(';');
    if (cropX || cropY) {
        layer.style.backgroundPosition = `calc(${style.backgroundPositionX || '0px'} - ${cropX}px) calc(${style.backgroundPositionY || '0px'} - ${cropY}px)`;
    }
    document.body.appendChild(layer);
    try {
        const data = await renderBackgroundImage(layer, { w: width, h: height }, {
            backgroundColor: null, requireVisiblePixels: true,
        });
        pptSlide.addImage({
            data, x: finalPosition.x, y: finalPosition.y,
            w: finalPosition.w, h: finalPosition.h, ...(rotate ? { rotate } : {}),
        });
        return true;
    } catch (error) {
        console.warn('[AiPPT] CSS 背景图层导出失败，已跳过:', error);
        return false;
    } finally {
        layer.remove();
    }
}

export async function writeGradientPageBackground({ pptSlide, background, slideRect, layout }) {
    const { element } = background;
    const document = element.ownerDocument;
    const style = document.defaultView.getComputedStyle(element);
    const layer = document.createElement('div');
    const width = Math.round(slideRect.width || layout.SLIDE_PIXEL_WIDTH);
    const height = Math.round(slideRect.height || layout.SLIDE_PIXEL_HEIGHT);
    layer.style.cssText = [
        'position:absolute', 'left:0', 'top:0', `width:${width}px`, `height:${height}px`,
        `background-color:${style.backgroundColor}`, `background-image:${style.backgroundImage}`,
        `background-position:${style.backgroundPosition}`, `background-repeat:${style.backgroundRepeat}`,
        `background-size:${style.backgroundSize}`, `background-origin:${style.backgroundOrigin}`,
        `background-clip:${style.backgroundClip}`, `opacity:${background.opacity}`, 'pointer-events:none',
    ].join(';');
    document.body.appendChild(layer);
    try {
        const data = await renderBackgroundImage(layer, { w: width, h: height }, {
            backgroundColor: null, requireVisiblePixels: true,
        });
        pptSlide.addImage({ data, x: 0, y: 0, w: layout.PPT_WIDTH_IN, h: layout.PPT_HEIGHT_IN });
        return true;
    } catch (error) {
        console.warn('[AiPPT] 渐变页面背景导出失败，已降级为纯色:', error);
        return false;
    } finally {
        layer.remove();
    }
}
