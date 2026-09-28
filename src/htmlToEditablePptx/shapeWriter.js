export function writeShape(pptSlide, element, style, position, helpers) {
    if (!position) return false;
    const {
        arrowTypes, colorTransparency, firstGradientColor, getBorderOptions,
        getEffectiveOpacity, getRoundedRectRadius, getRotateDeg, getShadowOptions,
        getShapeType, isTransparentColor, normalizeCssColor, shapeTypes,
    } = helpers;
    const type = getShapeType(element, style, position);
    const effectiveOpacity = getEffectiveOpacity(element);
    const background = !isTransparentColor(style.backgroundColor)
        ? style.backgroundColor : firstGradientColor(style.backgroundImage);
    const fillColor = normalizeCssColor(background);
    const options = { x: position.x, y: position.y, w: position.w, h: position.h };
    const rotate = getRotateDeg(style);
    if (rotate) options.rotate = rotate;
    if (type === shapeTypes.line) {
        options.line = getBorderOptions(style, effectiveOpacity) || {
            color: fillColor || '666666', pt: Math.max(0.5, Math.min(position.h, position.w) * 72),
        };
        const endArrowType = element.getAttribute('data-pptx-arrow-end');
        const beginArrowType = element.getAttribute('data-pptx-arrow-begin');
        if (arrowTypes.has(endArrowType)) options.line.endArrowType = endArrowType;
        if (arrowTypes.has(beginArrowType)) options.line.beginArrowType = beginArrowType;
    } else {
        options.fill = fillColor
            ? { color: fillColor, transparency: colorTransparency(background, effectiveOpacity) }
            : { color: 'FFFFFF', transparency: 100 };
        const line = getBorderOptions(style, effectiveOpacity);
        if (line) options.line = line;
        if (type === shapeTypes.roundRect) {
            const rectRadius = getRoundedRectRadius(style, position);
            if (rectRadius !== undefined) options.rectRadius = rectRadius;
        }
        const shadow = getShadowOptions(style);
        if (shadow && type !== shapeTypes.ellipse) options.shadow = shadow;
    }
    try {
        pptSlide.addShape(type, options);
        return true;
    } catch (error) {
        console.warn('[AiPPT] 形状导出失败，已跳过:', error);
        return false;
    }
}

export function writeSlideRootFrame(pptSlide, root, style, position, helpers) {
    const { getBorderOptions, getEffectiveOpacity, getRoundedRectRadius, getShadowOptions,
        getShapeType, hasAsymmetricVisibleBorder, shapeTypes } = helpers;
    const border = getBorderOptions(style, getEffectiveOpacity(root));
    const shadow = getShadowOptions(style);
    if ((!border && !shadow) || hasAsymmetricVisibleBorder(style) || !position) return false;
    const type = getShapeType(root, style, position);
    if (type === shapeTypes.line || type === shapeTypes.ellipse) return false;
    const options = {
        x: position.x, y: position.y, w: position.w, h: position.h,
        fill: { color: 'FFFFFF', transparency: 100 },
        ...(border ? { line: border } : {}), ...(shadow ? { shadow } : {}),
    };
    if (type === shapeTypes.roundRect) {
        const rectRadius = getRoundedRectRadius(style, position);
        if (rectRadius !== undefined) options.rectRadius = rectRadius;
    }
    try {
        pptSlide.addShape(type, options);
        return true;
    } catch (error) {
        console.warn('[AiPPT] 幻灯片边框导出失败，已跳过:', error);
        return false;
    }
}
