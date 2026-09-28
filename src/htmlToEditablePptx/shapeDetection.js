import { clamp, normalizeCssColor } from './styleExtractor';
import { colorTransparency, hasCssBackgroundImage, isTransparentColor } from './visualStyle';
import { getTagName, NON_VISUAL_TAGS } from './elementTraversal';

export const SHAPE_TYPES = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', line: 'line' };

export function getCornerRadiusValues(style) {
    const shorthand = String(style.borderRadius || '').split('/')[0].trim();
    const tokens = shorthand ? shorthand.split(/\s+/).filter(Boolean) : [
        style.borderTopLeftRadius, style.borderTopRightRadius,
        style.borderBottomRightRadius, style.borderBottomLeftRadius,
    ].map(value => String(value || '').trim());
    if (tokens.length === 1) return [tokens[0], tokens[0], tokens[0], tokens[0]];
    if (tokens.length === 2) return [tokens[0], tokens[1], tokens[0], tokens[1]];
    if (tokens.length === 3) return [tokens[0], tokens[1], tokens[2], tokens[1]];
    return tokens.slice(0, 4);
}

export function getBorderRadiusCss(style) {
    const shorthand = String(style.borderRadius || '').trim();
    if (shorthand) return shorthand;
    const values = getCornerRadiusValues(style);
    return values.length === 4 && values.every(Boolean) ? values.join(' ') : '';
}

function isEllipse(style, position) {
    const radii = getCornerRadiusValues(style);
    if (radii.length !== 4 || radii.some(value => !value)) return false;
    if (radii.every(value => String(value).endsWith('%') && (Number.parseFloat(value) || 0) >= 50)) return true;
    const { width, height } = position.rect;
    if (Math.abs(width - height) > Math.max(1, Math.min(width, height) * 0.025)) return false;
    const halfSide = (Math.min(width, height) / 2) - 1;
    return radii.every(value => (Number.parseFloat(value) || 0) >= halfSide);
}

export function getRoundedRectRadius(style, position) {
    const value = String(style.borderTopLeftRadius || '').trim();
    const radius = Number.parseFloat(value);
    if (!Number.isFinite(radius) || radius <= 0) return undefined;
    const shortestSide = Math.min(position.rect.width, position.rect.height);
    const radiusPx = value.endsWith('%') ? shortestSide * radius / 100 : radius;
    return clamp(radiusPx / Math.max(shortestSide / 2, 1), 0, 1);
}

export function getShadowOptions(style) {
    const value = String(style.boxShadow || '').trim();
    if (!value || value === 'none' || /\binset\b/i.test(value)) return undefined;
    const colorValue = value.match(/rgba?\([^)]*\)|#[0-9a-f]{3,8}/i)?.[0];
    const color = normalizeCssColor(colorValue);
    const values = value.replace(colorValue || '', '').match(/-?\d*\.?\d+px/g)?.map(Number.parseFloat) || [];
    if (!color || values.length < 3) return undefined;
    const [x, y, blur] = values;
    const alpha = Number.parseFloat(colorValue?.match(/rgba\([^,]+,[^,]+,[^,]+,\s*([\d.]+)\)/i)?.[1]);
    return { type: 'outer', color, opacity: clamp(Number.isFinite(alpha) ? alpha : 1, 0, 1),
        blur: Math.max(0, blur * 0.75), offset: Math.hypot(x, y) * 0.75,
        angle: (Math.atan2(y, x) * 180 / Math.PI + 360) % 360 };
}

export function getShapeType(element, style, position) {
    const requested = element.getAttribute('data-pptx-shape');
    if (requested && SHAPE_TYPES[requested]) return SHAPE_TYPES[requested];
    if (element.getAttribute('data-pptx-role') === 'line' || getTagName(element) === 'HR') return SHAPE_TYPES.line;
    if (element.getAttribute('data-pptx-role') === 'ellipse' || isEllipse(style, position)) return SHAPE_TYPES.ellipse;
    return (Number.parseFloat(style.borderTopLeftRadius) || 0) >= 2 ? SHAPE_TYPES.roundRect : SHAPE_TYPES.rect;
}

export function getBorderOptions(style, effectiveOpacity = Number.parseFloat(style.opacity) || 1) {
    const widthPx = Number.parseFloat(style.borderTopWidth) || 0;
    const color = normalizeCssColor(style.borderTopColor);
    if (!widthPx || !color || style.borderTopStyle === 'none') return undefined;
    return { color, transparency: colorTransparency(style.borderTopColor, effectiveOpacity), pt: Math.max(0.25, widthPx * 0.75) };
}

export function hasAsymmetricVisibleBorder(style) {
    const sides = ['Top', 'Right', 'Bottom', 'Left'].map(side => ({
        width: Number.parseFloat(style[`border${side}Width`]) || 0,
        style: String(style[`border${side}Style`] || 'none').toLowerCase(),
        color: normalizeCssColor(style[`border${side}Color`]),
    }));
    const visible = sides.filter(side => side.width > 0 && side.style !== 'none' && Boolean(side.color));
    if (!visible.length) return false;
    if (visible.length !== sides.length) return true;
    return visible.some(side => Math.abs(side.width - visible[0].width) > 0.01 || side.style !== visible[0].style || side.color !== visible[0].color);
}

export function hasAsymmetricCornerRadius(style) {
    const values = getCornerRadiusValues(style).map(value => String(value || '').trim());
    return values.length === 4 && values.every(Boolean) && values.some(value => (Number.parseFloat(value) || 0) > 0) && values.some(value => value !== values[0]);
}

export function isShapeCandidate(element, style) {
    const tagName = getTagName(element);
    if (NON_VISUAL_TAGS.has(tagName) || ['IMG', 'TABLE', 'SVG'].includes(tagName)) return false;
    const opacity = Number.parseFloat(style.opacity);
    if (style.display === 'none' || style.visibility === 'hidden' || (Number.isFinite(opacity) && opacity === 0)) return false;
    const role = element.getAttribute('data-pptx-role');
    if (['shape', 'line', 'ellipse'].includes(role) || element.getAttribute('data-pptx-shape')) return true;
    return !isTransparentColor(style.backgroundColor) || hasCssBackgroundImage(style) ||
        ['Top', 'Right', 'Bottom', 'Left'].some(side => Number.parseFloat(style[`border${side}Width`]) > 0) || tagName === 'HR';
}
