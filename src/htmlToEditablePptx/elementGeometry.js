import { AI_PPT_EXPORT_LAYOUT } from '../../constants';
import { clamp, getRotateDeg } from './styleExtractor';

function getBoxShadowCaptureInsets(style) {
    const value = String(style?.boxShadow || '').trim();
    if (!value || value === 'none' || /\binset\b/i.test(value)) {
        return { top: 0, right: 0, bottom: 0, left: 0 };
    }
    const values = value.replace(/rgba?\([^)]*\)|#[0-9a-f]{3,8}/ig, '')
        .match(/-?\d*\.?\d+px/g)?.slice(0, 4).map(Number.parseFloat) || [];
    if (values.length < 3) return { top: 0, right: 0, bottom: 0, left: 0 };
    const [x, y, blur, spread = 0] = values;
    const extent = Math.max(0, (blur * 1.5) + spread);
    return {
        top: Math.max(0, extent - y), right: Math.max(0, extent + x),
        bottom: Math.max(0, extent + y), left: Math.max(0, extent - x),
    };
}

export function getSlideBoundedPosition(element, slideRect, pxToInH, pxToInV, { includeShadow = false } = {}) {
    const rect = element.getBoundingClientRect();
    if (!rect || rect.width <= 0.5 || rect.height <= 0.5) return null;
    const right = rect.right ?? (rect.left + rect.width);
    const bottom = rect.bottom ?? (rect.top + rect.height);
    const slideRight = slideRect.right ?? (slideRect.left + slideRect.width);
    const slideBottom = slideRect.bottom ?? (slideRect.top + slideRect.height);
    const inset = includeShadow
        ? getBoxShadowCaptureInsets(element.ownerDocument.defaultView.getComputedStyle(element))
        : { top: 0, right: 0, bottom: 0, left: 0 };
    const left = Math.max(rect.left - inset.left, slideRect.left);
    const top = Math.max(rect.top - inset.top, slideRect.top);
    const boundedRight = Math.min(right + inset.right, slideRight);
    const boundedBottom = Math.min(bottom + inset.bottom, slideBottom);
    if (boundedRight <= left || boundedBottom <= top) return null;
    return {
        rect,
        visibleRect: { left, top, right: boundedRight, bottom: boundedBottom, width: boundedRight - left, height: boundedBottom - top },
        x: clamp(pxToInH(left - slideRect.left), 0, AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN),
        y: clamp(pxToInV(top - slideRect.top), 0, AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN),
        w: Math.min(pxToInH(boundedRight - left), AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN),
        h: Math.min(pxToInV(boundedBottom - top), AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN),
    };
}

export function getElementPosition(element, slideRect, pxToInH, pxToInV, helpers) {
    const rect = element.getBoundingClientRect();
    if (!rect || rect.width <= 0.5 || rect.height <= 0.5) return null;
    let left = Math.max(rect.left, slideRect.left);
    let top = Math.max(rect.top, slideRect.top);
    let right = Math.min(rect.right ?? (rect.left + rect.width), slideRect.right ?? (slideRect.left + slideRect.width));
    let bottom = Math.min(rect.bottom ?? (rect.top + rect.height), slideRect.bottom ?? (slideRect.top + slideRect.height));
    const slideVisibleRect = { left, top, right, bottom };
    let clippedByAncestor = false;
    const clipsOverflow = value => ['auto', 'clip', 'hidden', 'scroll'].includes(String(value || '').toLowerCase());
    let ancestor = element.parentElement;
    while (ancestor && ancestor !== element.ownerDocument.documentElement) {
        const style = element.ownerDocument.defaultView.getComputedStyle(ancestor);
        const overflowX = style.overflowX || style.overflow;
        const overflowY = style.overflowY || style.overflow;
        if ((clipsOverflow(overflowX) || clipsOverflow(overflowY)) && helpers.clipsOutOfFlowElement(element, ancestor)) {
            const ancestorRect = ancestor.getBoundingClientRect();
            if (clipsOverflow(overflowX)) {
                const previousLeft = left; const previousRight = right;
                left = Math.max(left, ancestorRect.left);
                right = Math.min(right, ancestorRect.right ?? (ancestorRect.left + ancestorRect.width));
                clippedByAncestor = clippedByAncestor || left !== previousLeft || right !== previousRight;
            }
            if (clipsOverflow(overflowY)) {
                const previousTop = top; const previousBottom = bottom;
                top = Math.max(top, ancestorRect.top);
                bottom = Math.min(bottom, ancestorRect.bottom ?? (ancestorRect.top + ancestorRect.height));
                clippedByAncestor = clippedByAncestor || top !== previousTop || bottom !== previousBottom;
            }
        }
        ancestor = ancestor.parentElement;
    }
    if (clippedByAncestor && helpers.hasVerifiedVisibleBounds(element, rect)) ({ left, top, right, bottom } = slideVisibleRect);
    if (right <= left || bottom <= top) return null;
    const x = pxToInH(left - slideRect.left); const y = pxToInV(top - slideRect.top);
    const w = pxToInH(right - left); const h = pxToInV(bottom - top);
    if (w <= 0 || h <= 0) return null;
    return {
        rect,
        visibleRect: { left, top, right, bottom, width: right - left, height: bottom - top },
        clipped: left !== rect.left || top !== rect.top ||
            right !== (rect.right ?? (rect.left + rect.width)) || bottom !== (rect.bottom ?? (rect.top + rect.height)),
        x: clamp(x, 0, AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN),
        y: clamp(y, 0, AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN),
        w: Math.min(w, AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN), h: Math.min(h, AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN),
    };
}

export function createsContainingBlock(style, { fixed = false } = {}) {
    const position = String(style?.position || '').toLowerCase();
    if (!fixed && position && position !== 'static') return true;
    const hasValue = value => {
        const normalized = String(value || '').trim().toLowerCase();
        return normalized && normalized !== 'none' && normalized !== 'normal' && normalized !== 'auto';
    };
    const contain = String(style?.contain || '').toLowerCase();
    return hasValue(style?.transform) || hasValue(style?.perspective) || hasValue(style?.filter) ||
        hasValue(style?.backdropFilter) || hasValue(style?.webkitBackdropFilter) ||
        /\b(layout|paint|strict|content)\b/.test(contain) ||
        /\b(transform|perspective|filter)\b/.test(String(style?.willChange || '').toLowerCase());
}

export function clipsOutOfFlowElement(element, ancestor) {
    const view = element.ownerDocument?.defaultView;
    const position = String(view?.getComputedStyle?.(element)?.position || '').toLowerCase();
    if (!['absolute', 'fixed'].includes(position)) return true;
    const fixed = position === 'fixed';
    let current = element.parentElement;
    while (current && current !== element.ownerDocument.documentElement) {
        const isContainingBlock = createsContainingBlock(view.getComputedStyle(current), { fixed });
        if (current === ancestor) return isContainingBlock;
        if (isContainingBlock) return true;
        current = current.parentElement;
    }
    return false;
}

export function hasVerifiedVisibleBounds(element, rect) {
    const document = element.ownerDocument;
    if (typeof document?.elementFromPoint !== 'function' && typeof document?.elementsFromPoint !== 'function') return false;
    const right = rect.right ?? (rect.left + rect.width);
    const bottom = rect.bottom ?? (rect.top + rect.height);
    return [[0.5, 0.5], [0.5, 0.16], [0.84, 0.5], [0.5, 0.84], [0.16, 0.5]].every(([xRatio, yRatio]) => {
        const x = rect.left + ((right - rect.left) * xRatio);
        const y = rect.top + ((bottom - rect.top) * yRatio);
        const getHits = () => {
            const hits = typeof document.elementsFromPoint === 'function' ? Array.from(document.elementsFromPoint(x, y) || []) : [];
            if (hits.length) return hits;
            const hit = typeof document.elementFromPoint === 'function' ? document.elementFromPoint(x, y) : null;
            return hit ? [hit] : [];
        };
        const contains = hits => hits.some(hit => hit === element || element.contains(hit));
        let hits = getHits();
        if (!contains(hits) && element.style && element.ownerDocument.defaultView.getComputedStyle(element).pointerEvents === 'none') {
            const previous = element.style.pointerEvents;
            element.style.pointerEvents = 'auto'; hits = getHits(); element.style.pointerEvents = previous;
        }
        return contains(hits);
    });
}

export function restoreUnrotatedPosition(element, position, slideRect, pxToInH, pxToInV) {
    if (!getRotateDeg(element.ownerDocument.defaultView.getComputedStyle(element)) || position.clipped) return position;
    const width = Number(element.offsetWidth); const height = Number(element.offsetHeight);
    if (!(width > 0 && height > 0)) return position;
    const right = position.rect.right ?? (position.rect.left + position.rect.width);
    const bottom = position.rect.bottom ?? (position.rect.top + position.rect.height);
    const centerX = (position.rect.left + right) / 2; const centerY = (position.rect.top + bottom) / 2;
    return {
        ...position, rect: { ...position.rect, width, height },
        x: clamp(pxToInH(centerX - (width / 2) - slideRect.left), 0, AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN),
        y: clamp(pxToInV(centerY - (height / 2) - slideRect.top), 0, AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN),
        w: Math.min(pxToInH(width), AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN),
        h: Math.min(pxToInV(height), AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN),
    };
}

function resolveTransformOriginCoordinate(value, size, axis) {
    const normalized = String(value || '').trim().toLowerCase();
    if (!normalized || normalized === 'center') return size / 2;
    if (normalized === 'left' || normalized === 'top') return 0;
    if (normalized === 'right' || normalized === 'bottom') return size;
    if (normalized.endsWith('%')) return size * (Number.parseFloat(normalized) || 0) / 100;
    const parsed = Number.parseFloat(normalized);
    return Number.isFinite(parsed) ? parsed : (axis === 'x' ? size / 2 : size / 2);
}

function getTransformOrigin(style, width, height) {
    const values = String(style?.transformOrigin || '50% 50%').trim().split(/\s+/);
    return {
        x: resolveTransformOriginCoordinate(values[0], width, 'x'),
        y: resolveTransformOriginCoordinate(values[1], height, 'y'),
    };
}

export function getTransformedLocalBounds(style, width, height) {
    const transform = String(style?.transform || '').trim();
    if (!transform || transform === 'none' || width <= 0 || height <= 0) return null;
    const matrixMatch = transform.match(/^matrix\(([^)]+)\)$/);
    const matrix3dMatch = transform.match(/^matrix3d\(([^)]+)\)$/);
    let mapPoint;
    if (matrixMatch) {
        const values = matrixMatch[1].split(',').map(value => Number.parseFloat(value.trim()));
        if (values.length !== 6 || values.some(value => !Number.isFinite(value))) return null;
        const [a, b, c, d, e, f] = values;
        mapPoint = (x, y) => ({ x: (a * x) + (c * y) + e, y: (b * x) + (d * y) + f });
    } else if (matrix3dMatch) {
        const values = matrix3dMatch[1].split(',').map(value => Number.parseFloat(value.trim()));
        if (values.length !== 16 || values.some(value => !Number.isFinite(value))) return null;
        mapPoint = (x, y) => {
            const denominator = (values[3] * x) + (values[7] * y) + values[15];
            if (Math.abs(denominator) < 0.00001) return null;
            return {
                x: ((values[0] * x) + (values[4] * y) + values[12]) / denominator,
                y: ((values[1] * x) + (values[5] * y) + values[13]) / denominator,
            };
        };
    } else return null;
    const origin = getTransformOrigin(style, width, height);
    const points = [[0, 0], [width, 0], [0, height], [width, height]].map(([x, y]) => {
        const mapped = mapPoint(x - origin.x, y - origin.y);
        return mapped && { x: mapped.x + origin.x, y: mapped.y + origin.y };
    });
    if (points.some(point => !point)) return null;
    return {
        left: Math.min(...points.map(point => point.x)), top: Math.min(...points.map(point => point.y)),
        right: Math.max(...points.map(point => point.x)), bottom: Math.max(...points.map(point => point.y)),
    };
}
