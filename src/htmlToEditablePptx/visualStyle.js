import { clamp, normalizeCssColor } from './styleExtractor';

export function isTransparentColor(value) {
    return !value || value === 'transparent' || /rgba\([^)]*,\s*0\s*\)/i.test(value);
}

export function colorTransparency(value, opacity = 1) {
    const rgbaMatch = String(value || '').match(/rgba?\(([^)]+)\)/i);
    const rgbaAlpha = rgbaMatch && rgbaMatch[1].split(',').length > 3
        ? Number(rgbaMatch[1].split(',')[3].trim()) : 1;
    const alpha = clamp((Number.isFinite(rgbaAlpha) ? rgbaAlpha : 1) * opacity, 0, 1);
    return Math.round((1 - alpha) * 100);
}

export function firstGradientColor(value) {
    const match = String(value || '').match(/(#[0-9a-f]{3,8}|rgba?\([^)]*\))/i);
    return match ? match[1] : undefined;
}

export function hasCssBackgroundImage(style) {
    return Boolean(style?.backgroundImage && style.backgroundImage !== 'none');
}

export function isGradientText(style) {
    return /text/i.test(String(style?.backgroundClip || style?.webkitBackgroundClip || '')) &&
        /gradient\(/i.test(String(style?.backgroundImage || ''));
}

export function getEffectiveOpacity(element) {
    let opacity = 1;
    let current = element;
    while (current?.ownerDocument) {
        const value = Number.parseFloat(current.ownerDocument.defaultView.getComputedStyle(current).opacity);
        if (Number.isFinite(value)) opacity *= value;
        current = current.parentElement;
    }
    return clamp(opacity, 0, 1);
}

export function getBackgroundInfo(element) {
    if (!element) return null;
    const style = element.ownerDocument.defaultView.getComputedStyle(element);
    const opacity = getEffectiveOpacity(element);
    if (opacity <= 0) return null;
    const backgroundColor = normalizeCssColor(style.backgroundColor);
    const gradientColor = firstGradientColor(style.backgroundImage);
    const color = backgroundColor || normalizeCssColor(gradientColor);
    const hasBackgroundImage = hasCssBackgroundImage(style);
    if (!color && !hasBackgroundImage) return null;
    return {
        color, source: backgroundColor ? style.backgroundColor : gradientColor, opacity, element,
        hasGradient: /gradient\(/i.test(style.backgroundImage || ''), hasBackgroundImage,
    };
}

export function getPageBackground(root) {
    const document = root.ownerDocument;
    const rootBackground = getBackgroundInfo(root);
    if (rootBackground) return rootBackground;
    const candidates = [document.body, document.documentElement].map(getBackgroundInfo).filter(Boolean);
    return candidates.find(background => background.color !== 'FFFFFF') || candidates[0] || null;
}

export function hasComplexTransform(style) {
    const transform = String(style?.transform || '').trim();
    if (!transform || transform === 'none') return false;
    if (transform.startsWith('matrix3d')) return true;
    const match = transform.match(/^matrix\(([^)]+)\)$/);
    if (!match) return true;
    const [a, b, c, d] = match[1].split(',').map(value => Number.parseFloat(value.trim()));
    if (![a, b, c, d].every(Number.isFinite)) return true;
    const epsilon = 0.01;
    return Math.abs(Math.hypot(a, b) - 1) > epsilon || Math.abs(Math.hypot(c, d) - 1) > epsilon ||
        Math.abs((a * c) + (b * d)) > epsilon || Math.abs(c + b) > epsilon || Math.abs(d - a) > epsilon;
}

function resolveTransformOriginCoordinate(value, size) {
    const normalized = String(value || '').trim().toLowerCase();
    if (!normalized || normalized === 'center') return size / 2;
    if (normalized === 'left' || normalized === 'top') return 0;
    if (normalized === 'right' || normalized === 'bottom') return size;
    if (normalized.endsWith('%')) return size * (Number.parseFloat(normalized) || 0) / 100;
    return Number.isFinite(Number.parseFloat(normalized)) ? Number.parseFloat(normalized) : size / 2;
}

export function getTransformedLocalBounds(style, width, height) {
    const transform = String(style?.transform || '').trim();
    const match = transform.match(/^matrix\(([^)]+)\)$/);
    if (!match || width <= 0 || height <= 0) return null;
    const [a, b, c, d, e, f] = match[1].split(',').map(value => Number.parseFloat(value.trim()));
    if (![a, b, c, d, e, f].every(Number.isFinite)) return null;
    const originValues = String(style.transformOrigin || '50% 50%').trim().split(/\s+/);
    const origin = { x: resolveTransformOriginCoordinate(originValues[0], width), y: resolveTransformOriginCoordinate(originValues[1], height) };
    const points = [[0, 0], [width, 0], [0, height], [width, height]].map(([x, y]) => ({
        x: (a * (x - origin.x)) + (c * (y - origin.y)) + e + origin.x,
        y: (b * (x - origin.x)) + (d * (y - origin.y)) + f + origin.y,
    }));
    return { left: Math.min(...points.map(point => point.x)), top: Math.min(...points.map(point => point.y)),
        right: Math.max(...points.map(point => point.x)), bottom: Math.max(...points.map(point => point.y)) };
}

export function hasComplexCssEffect(style) {
    const hasValue = value => Boolean(value && value !== 'none' && value !== 'normal');
    return hasValue(style?.filter) || hasValue(style?.backdropFilter) || hasValue(style?.webkitBackdropFilter) ||
        hasValue(style?.maskImage) || hasValue(style?.webkitMaskImage) || hasValue(style?.clipPath) ||
        hasValue(style?.webkitClipPath) || hasValue(style?.mixBlendMode) || hasValue(style?.backgroundBlendMode);
}

export function isVisiblePseudoStyle(style) {
    if (!style || style.display === 'none' || style.visibility === 'hidden') return false;
    const opacity = Number.parseFloat(style.opacity);
    if (Number.isFinite(opacity) && opacity <= 0) return false;
    const content = String(style.content || '').trim();
    const generatesBox = Boolean(content && content !== 'none' && content !== 'normal');
    if (!generatesBox) return false;
    const hasContent = content !== "''" && content !== '""';
    const hasBackground = !isTransparentColor(style.backgroundColor) || hasCssBackgroundImage(style);
    const hasBorder = ['Top', 'Right', 'Bottom', 'Left'].some(side => Number.parseFloat(style[`border${side}Width`]) > 0);
    const hasArea = Number.parseFloat(style.width) > 0 && Number.parseFloat(style.height) > 0;
    return hasContent || hasBackground || hasBorder || hasArea;
}

export function hasVisiblePseudoElement(element) {
    const view = element.ownerDocument?.defaultView;
    if (!view) return false;
    return ['::before', '::after'].some(pseudo => {
        try { return isVisiblePseudoStyle(view.getComputedStyle(element, pseudo)); } catch (error) { return false; }
    });
}
