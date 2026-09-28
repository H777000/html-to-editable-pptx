import { isVisiblePseudoStyle } from './visualStyle';

let pseudoSuppressionSequence = 0;

const RASTER_PROPERTIES = [
    ['display', 'display'], ['boxSizing', 'box-sizing'], ['alignItems', 'align-items'], ['justifyContent', 'justify-content'],
    ['flexDirection', 'flex-direction'], ['flexWrap', 'flex-wrap'], ['gap', 'gap'], ['padding', 'padding'],
    ['backgroundColor', 'background-color'], ['backgroundImage', 'background-image'], ['backgroundPosition', 'background-position'],
    ['backgroundSize', 'background-size'], ['backgroundRepeat', 'background-repeat'], ['backgroundOrigin', 'background-origin'],
    ['backgroundClip', 'background-clip'], ['borderTop', 'border-top'], ['borderRight', 'border-right'], ['borderBottom', 'border-bottom'],
    ['borderLeft', 'border-left'], ['borderRadius', 'border-radius'], ['boxShadow', 'box-shadow'], ['opacity', 'opacity'],
    ['overflow', 'overflow'], ['color', 'color'], ['fontFamily', 'font-family'], ['fontSize', 'font-size'],
    ['fontStyle', 'font-style'], ['fontWeight', 'font-weight'], ['lineHeight', 'line-height'], ['letterSpacing', 'letter-spacing'],
    ['textAlign', 'text-align'], ['textTransform', 'text-transform'], ['textShadow', 'text-shadow'], ['whiteSpace', 'white-space'],
    ['wordBreak', 'word-break'], ['writingMode', 'writing-mode'], ['direction', 'direction'],
];

const PSEUDO_PROPERTIES = [
    ['position', 'position'], ['display', 'display'], ['boxSizing', 'box-sizing'], ['top', 'top'], ['right', 'right'],
    ['bottom', 'bottom'], ['left', 'left'], ['inset', 'inset'], ['zIndex', 'z-index'], ['width', 'width'], ['height', 'height'],
    ['minWidth', 'min-width'], ['minHeight', 'min-height'], ['maxWidth', 'max-width'], ['maxHeight', 'max-height'],
    ['margin', 'margin'], ['padding', 'padding'], ...RASTER_PROPERTIES.slice(8, 23), ['transform', 'transform'],
    ['transformOrigin', 'transform-origin'], ...RASTER_PROPERTIES.slice(23),
];

export function inlineRasterizedVisualStyle(sourceStyle, clone) {
    RASTER_PROPERTIES.forEach(([property, cssProperty]) => {
        const value = sourceStyle?.[property] || sourceStyle?.getPropertyValue?.(cssProperty);
        if (value && value !== 'initial') clone.style.setProperty(cssProperty, value, 'important');
    });
}

function getPseudoContent(style) {
    const content = String(style?.content || '').trim();
    if (!content || ['none', 'normal', "''", '""'].includes(content)) return '';
    return (content[0] === '"' || content[0] === "'") && content.endsWith(content[0]) ? content.slice(1, -1) : content;
}

function inlineResolvedStyle(sourceStyle, target) {
    const properties = new Set(PSEUDO_PROPERTIES.map(([, cssProperty]) => cssProperty));
    const length = Number(sourceStyle?.length);
    if (Number.isFinite(length)) for (let index = 0; index < length; index += 1) properties.add(sourceStyle[index] || sourceStyle.item?.(index));
    properties.forEach(cssProperty => {
        if (!cssProperty || cssProperty === 'content') return;
        const property = PSEUDO_PROPERTIES.find(([, name]) => name === cssProperty)?.[0];
        const value = sourceStyle?.getPropertyValue?.(cssProperty) || sourceStyle?.[property];
        if (value && value !== 'initial') target.style.setProperty(cssProperty, value, 'important');
    });
}

function addPseudoContent(layer, style) {
    const content = getPseudoContent(style);
    if (!content) return;
    const imageMatch = content.match(/^url\(\s*['"]?([^'")]+)['"]?\s*\)$/i);
    if (imageMatch) {
        const image = layer.ownerDocument.createElement('img');
        image.setAttribute('src', imageMatch[1]); image.setAttribute('aria-hidden', 'true');
        ['display', 'width', 'height'].forEach((property, index) => image.style.setProperty(property, index ? '100%' : 'block', 'important'));
        layer.appendChild(image); return;
    }
    layer.textContent = content;
}

export function materializePseudoElements(element, host) {
    const view = element.ownerDocument?.defaultView;
    if (!view) return false;
    const pseudoStyles = ['::before', '::after'].map(pseudo => ({ pseudo, style: view.getComputedStyle(element, pseudo) }))
        .filter(({ style }) => isVisiblePseudoStyle(style));
    if (!pseudoStyles.length) return false;
    pseudoStyles.forEach(({ pseudo, style }) => {
        const layer = element.ownerDocument.createElement('span');
        layer.setAttribute('aria-hidden', 'true'); layer.setAttribute('data-pptx-pseudo', pseudo.slice(2));
        inlineResolvedStyle(style, layer); addPseudoContent(layer, style);
        if (pseudo === '::before') host.insertBefore(layer, host.firstChild); else host.appendChild(layer);
    });
    return true;
}

// A complex capture clone keeps the source class names so that its computed
// layout remains faithful. Once its generated pseudo paint has been copied to
// real child nodes, suppress only that clone's original pseudo paint; otherwise
// the browser renders the same decoration twice in the rasterized image.
export function suppressCloneNativePseudoElements(clone) {
    const doc = clone?.ownerDocument;
    if (!doc?.head) return () => {};
    const token = `ai-ppt-pseudo-${pseudoSuppressionSequence += 1}`;
    const style = doc.createElement('style');
    clone.setAttribute('data-pptx-pseudo-suppression', token);
    style.setAttribute('data-pptx-pseudo-suppression-style', token);
    style.textContent = `[data-pptx-pseudo-suppression="${token}"]::before,
      [data-pptx-pseudo-suppression="${token}"]::after {
        content: none !important;
        display: none !important;
      }`;
    doc.head.appendChild(style);
    return () => {
        style.remove();
        clone.removeAttribute('data-pptx-pseudo-suppression');
    };
}

export function inlineAtomicVisualSubtreeStyles(source, clone) {
    const sourceNodes = [source, ...Array.from(source.querySelectorAll('*'))];
    const cloneNodes = [clone, ...Array.from(clone.querySelectorAll('*'))];
    sourceNodes.forEach((sourceNode, index) => {
        const cloneNode = cloneNodes[index];
        if (cloneNode) inlineResolvedStyle(sourceNode.ownerDocument.defaultView.getComputedStyle(sourceNode), cloneNode);
    });
}

// Rasterizes only generated pseudo-element paint while keeping normal child
// elements available to the native export paths.
export function getRasterCaptureOptions(position, style, { includeTransform = true } = {}) {
    if (!position.clipped) {
        return { width: position.rect.width, height: position.rect.height,
            options: { backgroundColor: null, requireVisiblePixels: true } };
    }
    const cropX = Math.max(0, position.visibleRect.left - position.rect.left);
    const cropY = Math.max(0, position.visibleRect.top - position.rect.top);
    const sourceTransform = includeTransform && style?.transform && style.transform !== 'none' ? ` ${style.transform}` : '';
    return {
        width: position.visibleRect.width, height: position.visibleRect.height,
        options: { backgroundColor: null, requireVisiblePixels: true, captureStyle: {
            width: `${position.rect.width}px`, height: `${position.rect.height}px`, margin: '0',
            transform: `translate(${-cropX}px, ${-cropY}px)${sourceTransform}`,
            transformOrigin: style?.transformOrigin || 'top left',
        } },
    };
}

export function mountCloneWithAncestorContext(element, clone, { ignoreAncestorOverflow = false } = {}) {
    const { body, documentElement } = element.ownerDocument;
    let host = clone;
    let ancestor = element.parentElement;
    while (ancestor && ancestor !== body && ancestor !== documentElement) {
        const shell = ancestor.cloneNode(false);
        shell.setAttribute('aria-hidden', 'true');
        if (ignoreAncestorOverflow) {
            ['overflow', 'overflow-x', 'overflow-y'].forEach(property => shell.style.setProperty(property, 'visible', 'important'));
        }
        shell.appendChild(host); host = shell; ancestor = ancestor.parentElement;
    }
    body.appendChild(host);
    return host;
}

export function mountCloneInIsolatedContext(element, clone) {
    element.ownerDocument.body.appendChild(clone);
    return clone;
}

export function temporarilySetStyle(element, declarations, restorers) {
    const previousStyle = element.getAttribute('style');
    restorers.push(() => {
        if (previousStyle === null) element.removeAttribute('style');
        else element.setAttribute('style', previousStyle);
    });
    Object.entries(declarations).forEach(([property, value]) => element.style.setProperty(property, value, 'important'));
}

export function isInBrowserComposedCaptureTree(node, element) {
    return node === element || node.contains(element) || element.contains(node);
}

export function temporarilyHideCaptureText(element, restorers) {
    const view = element.ownerDocument.defaultView;
    const textParents = new Set();
    const walker = element.ownerDocument.createTreeWalker(element, view.NodeFilter.SHOW_TEXT);
    let textNode = walker.nextNode();
    while (textNode) {
        if (String(textNode.textContent || '').trim() && !textNode.parentElement?.closest('svg')) textParents.add(textNode.parentElement);
        textNode = walker.nextNode();
    }
    textParents.forEach(parent => {
        const style = view.getComputedStyle(parent);
        const declarations = {
            color: 'transparent', '-webkit-text-fill-color': 'transparent',
            '-webkit-text-stroke-color': 'transparent', 'text-decoration-color': 'transparent', 'text-shadow': 'none',
        };
        if (/text/i.test(String(style.backgroundClip || style.webkitBackgroundClip || ''))) declarations['background-image'] = 'none';
        temporarilySetStyle(parent, declarations, restorers);
    });
}

export function temporarilyPinCaptureColor(element, restorers) {
    const color = String(element.ownerDocument.defaultView.getComputedStyle(element).color || '').trim();
    if (color) temporarilySetStyle(element, { color }, restorers);
}

export async function capturePseudoElementLayer({ pptSlide, element, position, operations }) {
    if (!position) return false;
    const style = element.ownerDocument.defaultView.getComputedStyle(element);
    const clone = element.ownerDocument.createElement('pptx-pseudo-host');
    Array.from(element.childNodes || []).forEach(child => clone.appendChild(child.cloneNode(true)));
    clone.setAttribute('aria-hidden', 'true');
    operations.materializePseudoElements(element, clone);
    operations.inlineRasterizedVisualStyle(style, clone);
    clone.style.cssText = [
        clone.style.cssText, 'position:absolute!important', 'left:0!important', 'top:0!important',
        `width:${Math.max(1, Math.round(position.rect.width))}px!important`,
        `height:${Math.max(1, Math.round(position.rect.height))}px!important`, 'margin:0!important',
        'color:transparent!important', 'background:transparent!important', 'border-color:transparent!important',
        'box-shadow:none!important', 'transform:none!important', 'pointer-events:none!important',
    ].join(';');
    Array.from(clone.children || []).forEach(child => {
        if (!child.hasAttribute('data-pptx-pseudo')) child.style.setProperty('visibility', 'hidden', 'important');
    });
    const mountedHost = operations.mountCloneWithAncestorContext(element, clone);
    try {
        const capture = operations.getRasterCaptureOptions(position, style, { includeTransform: false });
        const data = await operations.renderBackgroundImage(clone, { w: capture.width, h: capture.height }, capture.options);
        pptSlide.addImage({ data, x: position.x, y: position.y, w: position.w, h: position.h });
        return true;
    } catch (error) {
        console.warn('[AiPPT] 伪元素装饰导出失败，已跳过:', error);
        return false;
    } finally {
        mountedHost.remove();
    }
}
