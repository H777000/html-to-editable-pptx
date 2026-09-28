import { cropRasterImage, renderBackgroundImage } from './imageExport';
import { getTagName } from './elementTraversal';

export async function captureComplexVisualLayer({ pptSlide, element, position, options = {}, operations }) {
    const { omitShadow = false, preserveText = false, atomic = false, browserComposedRoot = null } = options;
    if (!position) return false;
    if (position.clipped && browserComposedRoot && await operations.captureBrowserComposedVisual({
        pptSlide, element, root: browserComposedRoot, ...operations.context,
        operations: operations.browserCaptureOperations,
    })) return true;
    const style = element.ownerDocument.defaultView.getComputedStyle(element);
    const layoutWidth = Number(element.offsetWidth) || position.rect.width;
    const layoutHeight = Number(element.offsetHeight) || position.rect.height;
    const transformedBounds = position.clipped
        ? operations.getTransformedLocalBounds(style, layoutWidth, layoutHeight) : null;
    const hasClippedTransform = Boolean(transformedBounds);
    const sourceWidth = (hasClippedTransform ? layoutWidth : position.rect.width) || position.rect.width;
    const sourceHeight = (hasClippedTransform ? layoutHeight : position.rect.height) || position.rect.height;
    const clone = element.cloneNode(true);
    const ownerDocument = element.ownerDocument;
    clone.setAttribute('aria-hidden', 'true');
    [['position', 'absolute'], ['width', `${Math.max(1, Math.round(sourceWidth))}px`],
        ['height', `${Math.max(1, Math.round(sourceHeight))}px`], ['margin', '0'], ['pointer-events', 'none']]
        .forEach(([property, value]) => clone.style.setProperty(property, value, 'important'));
    operations.inlineRasterizedVisualStyle(style, clone);
    const borderRadius = operations.getBorderRadiusCss(style);
    if (borderRadius) clone.style.borderRadius = borderRadius;
    if (omitShadow) clone.style.boxShadow = 'none';
    let restoreNativePseudoElements = () => {};
    if (atomic) {
        operations.inlineAtomicVisualSubtreeStyles(element, clone);
        if (operations.materializePseudoElements(element, clone)) {
            restoreNativePseudoElements = operations.suppressCloneNativePseudoElements(clone);
        }
    }
    if (operations.getShapeType(element, style, position) === operations.shapeTypes.ellipse) {
        clone.style.backgroundColor = style.backgroundColor;
        clone.style.backgroundImage = style.backgroundImage;
        if (style.clipPath && style.clipPath !== 'none') clone.style.clipPath = style.clipPath;
        if (style.webkitClipPath && style.webkitClipPath !== 'none') clone.style.webkitClipPath = style.webkitClipPath;
        if (['hidden', 'clip'].includes(String(style.overflow || '').toLowerCase())) clone.style.overflow = style.overflow;
    }
    if (!preserveText) {
        const svgVisualNodes = new Set([...Array.from(clone.querySelectorAll('svg')), ...Array.from(clone.querySelectorAll('svg *'))]);
        [clone, ...Array.from(clone.querySelectorAll('*'))].forEach(node => {
            if (svgVisualNodes.has(node) || node.hasAttribute?.('data-pptx-pseudo')) return;
            ['color', '-webkit-text-fill-color', '-webkit-text-stroke-color', 'text-decoration-color']
                .forEach(property => node.style.setProperty(property, 'transparent', 'important'));
            node.style.setProperty('text-shadow', 'none', 'important');
            const nodeStyle = ownerDocument.defaultView.getComputedStyle(node);
            if (/text/i.test(String(nodeStyle.backgroundClip || nodeStyle.webkitBackgroundClip || ''))) node.style.setProperty('background-image', 'none', 'important');
            if (['TEXT', 'TSPAN'].includes(getTagName(node))) {
                node.style.setProperty('fill', 'transparent', 'important');
                node.style.setProperty('stroke', 'transparent', 'important');
            }
        });
    }
    let captureTarget = clone;
    if (hasClippedTransform) {
        clone.style.setProperty('transform', style.transform, 'important');
        clone.style.setProperty('transform-origin', style.transformOrigin || '50% 50%', 'important');
        const viewport = ownerDocument.createElement('div');
        viewport.setAttribute('aria-hidden', 'true');
        viewport.style.cssText = [
            'position:relative!important', `width:${Math.max(1, Math.round(position.rect.width))}px!important`,
            `height:${Math.max(1, Math.round(position.rect.height))}px!important`, 'overflow:hidden!important',
            'margin:0!important', 'padding:0!important', 'background:transparent!important', 'pointer-events:none!important',
        ].join(';');
        clone.style.setProperty('left', `${-transformedBounds.left}px`, 'important');
        clone.style.setProperty('top', `${-transformedBounds.top}px`, 'important');
        viewport.appendChild(clone); captureTarget = viewport;
    } else {
        clone.style.setProperty('left', '0', 'important'); clone.style.setProperty('top', '0', 'important');
    }
    const mountedHost = atomic ? operations.mountCloneInIsolatedContext(element, captureTarget)
        : operations.mountCloneWithAncestorContext(element, captureTarget, { ignoreAncestorOverflow: position.clipped });
    try {
        const capture = operations.getRasterCaptureOptions(position, style);
        const fullData = await renderBackgroundImage(captureTarget, {
            w: position.clipped ? position.rect.width : capture.width,
            h: position.clipped ? position.rect.height : capture.height,
        }, position.clipped ? { backgroundColor: null, requireVisiblePixels: true } : capture.options);
        const data = position.clipped ? await cropRasterImage(fullData, {
            sourceWidth: position.rect.width, sourceHeight: position.rect.height,
            cropX: Math.max(0, position.visibleRect.left - position.rect.left),
            cropY: Math.max(0, position.visibleRect.top - position.rect.top),
            width: position.visibleRect.width, height: position.visibleRect.height,
        }, ownerDocument) : fullData;
        pptSlide.addImage({ data, x: position.x, y: position.y, w: position.w, h: position.h });
        return true;
    } catch (error) {
        console.warn('[AiPPT] 复杂 CSS 视觉层导出失败，已跳过:', error);
        return false;
    } finally {
        if (mountedHost.isConnected) mountedHost.remove();
        restoreNativePseudoElements();
    }
}
