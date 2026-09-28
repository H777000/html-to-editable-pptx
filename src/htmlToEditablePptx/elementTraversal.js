export const NON_VISUAL_TAGS = new Set([
    'SCRIPT', 'STYLE', 'LINK', 'META', 'HEAD', 'NOSCRIPT', 'TEXT', 'TSPAN',
    'PATH', 'G', 'DEFS', 'CLIPPATH', 'MASK', 'USE',
]);

export function getTagName(element) {
    return String(element.tagName || '').toUpperCase();
}

function getZIndex(element) {
    const value = element.ownerDocument.defaultView.getComputedStyle(element).zIndex;
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) ? parsed : 0;
}

// Preserves DOM sibling order inside each z-index band, matching browser paint order.
export function collectElementsInPaintOrder(root) {
    const result = [];
    const visitChildren = parent => {
        const children = Array.from(parent.children || [])
            .map((element, index) => ({ element, index, zIndex: getZIndex(element) }))
            .sort((left, right) => left.zIndex - right.zIndex || left.index - right.index);
        children.forEach(({ element }) => {
            result.push(element);
            visitChildren(element);
        });
    };
    visitChildren(root);
    return result;
}
