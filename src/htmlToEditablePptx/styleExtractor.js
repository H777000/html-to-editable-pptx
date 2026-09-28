import { AI_PPT_EXPORT_FONT_MAP, AI_PPT_EXPORT_LAYOUT } from '../../constants';

const SKIP_TEXT_TAGS = new Set(['SCRIPT', 'STYLE', 'LINK', 'META', 'HEAD', 'NOSCRIPT']);
const SANS_FALLBACK = 'Microsoft YaHei';
const SERIF_FALLBACK = 'SimSun';

export function stripHash(color) {
    return color ? color.replace('#', '').toUpperCase() : undefined;
}

export function normalizeCssColor(value) {
    if (!value || value === 'transparent') return undefined;
    if (value.startsWith('#')) return stripHash(value);

    const rgbaMatch = value.match(/rgba?\(([^)]+)\)/i);
    if (!rgbaMatch) return undefined;

    const parts = rgbaMatch[1].split(',').map(part => part.trim());
    const alpha = parts.length > 3 ? Number(parts[3]) : 1;
    if (Number.isFinite(alpha) && alpha <= 0) return undefined;

    const [red, green, blue] = parts.slice(0, 3).map(part => {
        const channel = Number(part);
        if (!Number.isFinite(channel)) return 0;
        return Math.max(0, Math.min(255, Math.round(channel)));
    });

    return [red, green, blue]
        .map(channel => channel.toString(16).padStart(2, '0'))
        .join('')
        .toUpperCase();
}

function getGradientTextColor(style) {
    const clipsText = /text/i.test(String(style.backgroundClip || style.webkitBackgroundClip || ''));
    if (!clipsText || !/gradient\(/i.test(String(style.backgroundImage || ''))) return undefined;
    const match = String(style.backgroundImage).match(/(#[0-9a-f]{3,8}|rgba?\([^)]*\))/i);
    return match?.[1];
}

function getColorAlpha(value) {
    const rgbaMatch = String(value || '').match(/rgba?\(([^)]+)\)/i);
    if (!rgbaMatch) return 1;

    const parts = rgbaMatch[1].split(',').map(part => part.trim());
    if (parts.length < 4) return 1;

    const alpha = Number(parts[3]);
    return Number.isFinite(alpha) ? clamp(alpha, 0, 1) : 1;
}

function getEffectiveTextOpacity(element, color) {
    let opacity = getColorAlpha(color);
    let current = element;

    while (current?.ownerDocument) {
        const value = Number.parseFloat(current.ownerDocument.defaultView.getComputedStyle(current).opacity);
        if (Number.isFinite(value)) opacity *= value;
        current = current.parentElement;
    }

    return clamp(opacity, 0, 1);
}

function getTextUnderline(style, textColor) {
    const line = String(style.textDecorationLine || style.textDecoration || '').toLowerCase();
    if (!/(^|\s)underline(\s|$)/.test(line)) return undefined;

    const decorationStyle = String(style.textDecorationStyle || '').toLowerCase();
    const underlineStyle = {
        dashed: 'dash',
        dotted: 'dotted',
        double: 'dbl',
        wavy: 'wavy',
    }[decorationStyle] || 'sng';
    const color = normalizeCssColor(style.textDecorationColor) || normalizeCssColor(textColor);

    return {
        style: underlineStyle,
        ...(color ? { color } : {}),
    };
}

function getExportPlatform(platform) {
    if (platform) return String(platform);
    if (typeof navigator === 'undefined') return '';
    return `${navigator.platform || ''} ${navigator.userAgent || ''}`;
}

export function mapFont(fontFamily, platform) {
    if (!fontFamily) return SANS_FALLBACK;

    const exportPlatform = getExportPlatform(platform);
    const isApplePlatform = /mac|iphone|ipad|ipod/i.test(exportPlatform);
    const parts = fontFamily
        .split(',')
        .map(part => part.replace(/["']/g, '').trim())
        .filter(Boolean);

    for (const part of parts) {
        // The browser preview on macOS uses PingFang SC. Replacing it with
        // Microsoft YaHei makes local WPS fall back to a different font,
        // changing glyph widths and baselines. Windows retains the existing
        // Microsoft YaHei mapping because PingFang is generally unavailable.
        if (part === 'PingFang SC' && isApplePlatform) {
            return 'PingFang SC';
        }
        if (AI_PPT_EXPORT_FONT_MAP[part]) {
            return AI_PPT_EXPORT_FONT_MAP[part];
        }
    }

    const normalized = fontFamily.toLowerCase();
    if (/serif/.test(normalized) && !/sans/.test(normalized)) {
        return SERIF_FALLBACK;
    }

    return parts[0] || SANS_FALLBACK;
}

export function applyTextTransform(text, transform) {
    if (transform === 'uppercase') return text.toUpperCase();
    if (transform === 'lowercase') return text.toLowerCase();
    return text;
}

export function isBoldWeight(weight) {
    if (!weight) return false;
    const parsed = Number.parseInt(weight, 10);
    return Number.isFinite(parsed) ? parsed >= 600 : /bold/i.test(weight);
}

export function alignToPptx(align) {
    if (align === 'center' || align === 'right' || align === 'justify') return align;
    if (align === 'end') return 'right';
    return 'left';
}

export function clamp(value, min, max) {
    return Math.max(min, Math.min(max, value));
}

export function isInlineDisplay(display) {
    return [
        'inline',
        'inline-block',
        'inline-flex',
        'inline-grid',
        'ruby',
        'contents',
    ].includes(display);
}

export function isBlockLikeChild(node) {
    if (node.nodeType !== Node.ELEMENT_NODE) return false;
    if (SKIP_TEXT_TAGS.has(node.tagName)) return false;

    const style = node.ownerDocument.defaultView.getComputedStyle(node);
    if (style.display === 'none') return false;
    if (!isInlineDisplay(style.display)) return true;

    const parent = node.parentElement;
    if (!parent) return false;

    const parentStyle = node.ownerDocument.defaultView.getComputedStyle(parent);
    return [
        'flex',
        'inline-flex',
        'grid',
        'inline-grid',
    ].includes(parentStyle.display);
}

export function extractRunStyle(element) {
    const style = element.ownerDocument.defaultView.getComputedStyle(element);
    const color = getGradientTextColor(style) || style.color;
    const textTransparency = Math.round((1 - getEffectiveTextOpacity(element, color)) * 100);
    const underline = getTextUnderline(style, color);
    const fontSizePx = Number.parseFloat(style.fontSize) || 14;
    const letterSpacingPx = (() => {
        if (!style.letterSpacing || style.letterSpacing === 'normal') return 0;
        const value = Number.parseFloat(style.letterSpacing);
        return Number.isFinite(value) ? value : 0;
    })();

    return {
        fontSizePt: Math.max(
            AI_PPT_EXPORT_LAYOUT.MIN_FONT_SIZE_PT,
            fontSizePx *
                ((AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN / AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH) * 72) *
                AI_PPT_EXPORT_LAYOUT.FONT_SIZE_SCALE,
        ),
        color,
        fontFamily: style.fontFamily,
        fontWeight: style.fontWeight,
        fontStyle: style.fontStyle,
        textTransform: style.textTransform,
        letterSpacingPt: letterSpacingPx *
            ((AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN / AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH) * 72) *
            AI_PPT_EXPORT_LAYOUT.FONT_SIZE_SCALE,
        ...(textTransparency > 0 ? { transparency: textTransparency } : {}),
        ...(underline ? { underline } : {}),
    };
}

export function getRotateDeg(style) {
    const transform = style.transform;
    if (!transform || !transform.startsWith('matrix')) return 0;

    const match = transform.match(/matrix\(([^)]+)\)/);
    if (!match) return 0;

    const values = match[1].split(',').map(value => Number.parseFloat(value.trim()));
    const [a, b] = values;
    const angle = Math.atan2(b, a) * (180 / Math.PI);
    if (Math.abs(angle) > 0.5) return angle;
    return 0;
}
