import { toPng } from 'html-to-image';
import { AI_PPT_EXPORT_LAYOUT } from '../../constants';

const SVG_STYLE_PROPERTIES = [
    'color',
    'display',
    'fill',
    'fill-opacity',
    'font-family',
    'font-size',
    'font-style',
    'font-weight',
    'opacity',
    'stroke',
    'stroke-linecap',
    'stroke-linejoin',
    'stroke-opacity',
    'stroke-width',
    'text-anchor',
    'visibility',
];

const SVG_CURRENT_COLOR_PROPERTIES = [
    'color',
    'fill',
    'flood-color',
    'lighting-color',
    'stop-color',
    'stroke',
];

function inlineSvgStyles(source, target) {
    const view = source.ownerDocument?.defaultView;
    if (!view?.getComputedStyle) return;

    const computedStyle = view.getComputedStyle(source);
    SVG_STYLE_PROPERTIES.forEach(property => {
        const value = computedStyle.getPropertyValue(property);
        if (value) target.style.setProperty(property, value);
    });

    Array.from(source.children || []).forEach((child, index) => {
        const targetChild = target.children?.[index];
        if (targetChild) inlineSvgStyles(child, targetChild);
    });
}

function getInheritedSvgColor(element) {
    const view = element.ownerDocument?.defaultView;
    let current = element;
    while (current) {
        const computedColor = view?.getComputedStyle?.(current)?.getPropertyValue('color');
        const inlineColor = current.style?.getPropertyValue?.('color');
        const color = String(computedColor || inlineColor || '').trim();
        if (color && color.toLowerCase() !== 'currentcolor') return color;
        current = current.parentElement;
    }
    return null;
}

function resolveSvgCurrentColor(source, target, inheritedColor = null) {
    const resolvedColor = getInheritedSvgColor(source) || inheritedColor;
    if (resolvedColor) {
        SVG_CURRENT_COLOR_PROPERTIES.forEach(property => {
            if (String(target.getAttribute(property) || '').trim().toLowerCase() === 'currentcolor') {
                target.setAttribute(property, resolvedColor);
            }
            if (String(target.style?.getPropertyValue(property) || '').trim().toLowerCase() === 'currentcolor') {
                target.style.setProperty(property, resolvedColor);
            }
        });
    }

    Array.from(source.children || []).forEach((child, index) => {
        const targetChild = target.children?.[index];
        if (targetChild) resolveSvgCurrentColor(child, targetChild, resolvedColor);
    });
}

const XML_ATTRIBUTE_NAME = /^[A-Za-z_][A-Za-z0-9_.:-]*$/;

export function sanitizeSvgMarkup(markup) {
    // HTML parsers can preserve a malformed attribute such as
    // `stroke-linecap="" "round"=""`. It renders in the browser but is not
    // valid XML and cannot be decoded from a data URI by the export library.
    return String(markup).replace(
        /\s+(?:"[^"]*"|'[^']*')\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/g,
        '',
    );
}

function removeInvalidSvgAttributes(element) {
    Array.from(element.attributes || []).forEach(attribute => {
        if (!XML_ATTRIBUTE_NAME.test(attribute.name)) {
            element.removeAttribute(attribute.name);
        }
    });
    Array.from(element.children || []).forEach(removeInvalidSvgAttributes);
}

function encodeBase64Utf8(value) {
    const binary = encodeURIComponent(value).replace(/%([0-9A-F]{2})/g, (_match, hex) => (
        String.fromCharCode(Number.parseInt(hex, 16))
    ));
    return window.btoa(binary);
}

export function serializeSvgToDataUri(svg, { width, height }) {
    const clone = svg.cloneNode(true);
    const safeWidth = Math.max(1, Math.round(width));
    const safeHeight = Math.max(1, Math.round(height));

    clone.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
    if (!clone.getAttribute('viewBox')) {
        clone.setAttribute('viewBox', `0 0 ${safeWidth} ${safeHeight}`);
    }
    clone.setAttribute('width', String(safeWidth));
    clone.setAttribute('height', String(safeHeight));
    inlineSvgStyles(svg, clone);
    // `currentColor` is resolved by the browser through CSS inheritance, but
    // Office/WPS SVG decoders do not consistently apply that inheritance.
    resolveSvgCurrentColor(svg, clone);
    // Browsers can display malformed inline SVG attributes, but PptxGenJS
    // later decodes this URI as XML. Remove only invalid XML attribute names.
    removeInvalidSvgAttributes(clone);

    const Serializer = svg.ownerDocument.defaultView?.XMLSerializer || window.XMLSerializer;
    const markup = Serializer
        ? new Serializer().serializeToString(clone)
        : clone.outerHTML;
    return `data:image/svg+xml;base64,${encodeBase64Utf8(sanitizeSvgMarkup(markup))}`;
}

async function loadRasterImage(dataUrl, document) {
    const view = document?.defaultView || window;
    return new Promise((resolve, reject) => {
        const target = new view.Image();
        const timer = view.setTimeout(() => reject(new Error('截图图像加载超时')), 3000);
        target.onload = () => {
            view.clearTimeout(timer);
            resolve(target);
        };
        target.onerror = () => {
            view.clearTimeout(timer);
            reject(new Error('截图结果无法解码'));
        };
        target.src = dataUrl;
    });
}

async function assertRasterHasVisiblePixels(dataUrl, document) {
    const image = await loadRasterImage(dataUrl, document);
    const canvas = document.createElement('canvas');
    canvas.width = Math.min(64, Math.max(1, image.naturalWidth || image.width || 1));
    canvas.height = Math.min(36, Math.max(1, image.naturalHeight || image.height || 1));
    const context = canvas.getContext('2d', { willReadFrequently: true });
    if (!context) throw new Error('浏览器不支持截图像素校验');
    context.drawImage(image, 0, 0, canvas.width, canvas.height);
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
    for (let index = 3; index < pixels.length; index += 4) {
        if (pixels[index] > 0) return;
    }
    throw new Error('截图结果完全透明');
}

export async function cropRasterImage(dataUrl, {
    sourceWidth,
    sourceHeight,
    cropX,
    cropY,
    width,
    height,
}, document) {
    const safeSourceWidth = Math.max(1, Number(sourceWidth) || 1);
    const safeSourceHeight = Math.max(1, Number(sourceHeight) || 1);
    const safeWidth = Math.max(1, Number(width) || 1);
    const safeHeight = Math.max(1, Number(height) || 1);
    const image = await loadRasterImage(dataUrl, document);
    const scaleX = (image.naturalWidth || image.width || safeSourceWidth) / safeSourceWidth;
    const scaleY = (image.naturalHeight || image.height || safeSourceHeight) / safeSourceHeight;
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(safeWidth * scaleX));
    canvas.height = Math.max(1, Math.round(safeHeight * scaleY));
    const context = canvas.getContext('2d');
    if (!context) throw new Error('浏览器不支持图片裁剪');

    context.drawImage(
        image,
        Math.max(0, cropX || 0) * scaleX,
        Math.max(0, cropY || 0) * scaleY,
        canvas.width,
        canvas.height,
        0,
        0,
        canvas.width,
        canvas.height,
    );
    return canvas.toDataURL('image/png');
}

export async function renderBackgroundImage(root, slideRect, {
    backgroundColor = '#ffffff',
    captureStyle,
    requireVisiblePixels = false,
} = {}) {
    try {
        const dataUrl = await toPng(root, {
            width: Math.round(slideRect.w || AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH),
            height: Math.round(slideRect.h || AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT),
            pixelRatio: AI_PPT_EXPORT_LAYOUT.BACKGROUND_PIXEL_RATIO,
            ...(backgroundColor ? { backgroundColor } : {}),
            ...(captureStyle ? { style: captureStyle } : {}),
            cacheBust: false,
        });
        if (requireVisiblePixels) {
            await assertRasterHasVisiblePixels(dataUrl, root.ownerDocument);
        }
        return dataUrl;
    } catch (error) {
        throw new Error(`HTML 背景图渲染失败：${error.message || error}`);
    }
}
