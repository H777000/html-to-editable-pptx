import { AI_PPT_EXPORT_LAYOUT } from '../../constants';

const TRUSTED_ASSET_HOST_PATTERNS = [];

const FRAME_ASSET_TIMEOUT_MS = 12000;
const REQUIRED_STABLE_LAYOUT_FRAMES = 3;
const FRAME_LAYOUT_TIMEOUT_MS = 5000;
const FRAME_LAYOUT_MAX_SAMPLES = 180;
const FRAME_ANIMATION_FRAME_FALLBACK_MS = 100;
let exportFrameSequence = 0;

function createExportAbortError() {
    const error = new Error('PPT 导出任务已取消');
    error.name = 'AbortError';
    return error;
}

export function isExportAbortError(error) {
    return error?.name === 'AbortError';
}

export function throwIfExportAborted(signal) {
    if (signal?.aborted) throw createExportAbortError();
}

function getRemainingTimeout(deadlineAt, fallbackTimeoutMs) {
    if (!deadlineAt) return fallbackTimeoutMs;
    const remaining = deadlineAt - Date.now();
    if (remaining <= 0) throw new Error('PPT 导出准备总超时');
    return Math.min(fallbackTimeoutMs, remaining);
}

function isTrustedAssetUrl(value) {
    const url = String(value || '').trim();
    if (!url) return true;
    if (/^(data:image\/|blob:|#)/i.test(url)) return true;
    try {
        const parsed = new URL(url, window.location?.origin || 'http://localhost');
        if (parsed.origin === window.location?.origin) return true;
        return TRUSTED_ASSET_HOST_PATTERNS.some(pattern => pattern.test(parsed.hostname));
    } catch (error) {
        return false;
    }
}

function removeUnsafeInlineUrls(element) {
    const style = element.getAttribute('style');
    if (!style || !/url\s*\(/i.test(style)) return;
    const cleaned = style.replace(/url\((['"]?)(.*?)\1\)/gi, (_match, _quote, url) => (
        isTrustedAssetUrl(url) ? `url("${url}")` : 'none'
    ));
    element.setAttribute('style', cleaned);
}

function sanitizeStyleText(cssText) {
    return String(cssText || '')
        .replace(/@import[^;]+;/gi, '')
        .replace(/url\((['"]?)(.*?)\1\)/gi, (_match, _quote, url) => (
            isTrustedAssetUrl(url) ? `url("${url}")` : 'none'
        ))
        .replace(/<\/style/gi, '<\\/style');
}

function serializeSafeRootAttributes(element) {
    if (!element) return '';
    return ['class', 'style', 'lang', 'dir']
        .map(name => {
            const value = element.getAttribute(name);
            if (!value) return '';
            const escaped = value
                .replace(/&/g, '&amp;')
                .replace(/"/g, '&quot;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;');
            return ` ${name}="${escaped}"`;
        })
        .join('');
}

function buildBaseStyle() {
    return `
  <style data-ai-ppt-export-style>
    html, body {
      margin: 0;
      padding: 0;
      width: ${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH}px;
      height: ${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT}px;
      overflow: hidden;
    }
</style>`;
}

function buildSlideRootNormalizationStyle() {
    // The preview iframe and export iframe must resolve the same slide root
    // geometry. Export also applies these values inline before extraction.
    return `
  <style data-ai-ppt-slide-root-style>
    .ppt-slide,
    [data-slide-root],
    [data-pptx="slide"],
    body > :first-child {
      position: absolute !important;
      left: 0 !important;
      top: 0 !important;
      width: ${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH}px !important;
      height: ${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT}px !important;
      margin: 0 !important;
      box-shadow: none !important;
      overflow: hidden !important;
    }
  </style>`;
}

function sanitizeAiPptDocumentForExport(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(normalizeHtmlDocument(html), 'text/html');

    doc.querySelectorAll('script, iframe, object, embed, meta[http-equiv], link[rel="stylesheet"]').forEach(node => {
        node.remove();
    });

    doc.querySelectorAll('*').forEach(element => {
        Array.from(element.attributes || []).forEach(attribute => {
            const name = attribute.name.toLowerCase();
            if (name.startsWith('on')) {
                element.removeAttribute(attribute.name);
            }
        });

        ['src', 'href', 'xlink:href', 'poster'].forEach(attributeName => {
            const value = element.getAttribute(attributeName);
            if (value && !isTrustedAssetUrl(value)) {
                element.removeAttribute(attributeName);
            }
        });
        removeUnsafeInlineUrls(element);
        if (element.tagName === 'IMG' && !element.getAttribute('src') && !element.getAttribute('srcset')) {
            element.remove();
        }
    });

    const headStyles = Array.from(doc.querySelectorAll('style'))
        .filter(styleNode => !styleNode.hasAttribute('data-ai-ppt-export-style'))
        .map(styleNode => `<style data-ai-ppt-preserved-style>${sanitizeStyleText(styleNode.textContent)}</style>`)
        .join('\n');

    doc.querySelectorAll('style').forEach(styleNode => {
        styleNode.remove();
    });

    return {
        bodyHtml: doc.body.innerHTML,
        headStyles,
        htmlAttributes: serializeSafeRootAttributes(doc.documentElement),
        bodyAttributes: serializeSafeRootAttributes(doc.body),
    };
}

export function sanitizeAiPptHtmlForExport(html) {
    return sanitizeAiPptDocumentForExport(html).bodyHtml;
}

function normalizeHtmlDocument(html) {
    const baseStyle = buildBaseStyle();
    const safeBody = html || '';

    if (/<html[\s>]/i.test(html || '')) {
        if (/<\/head>/i.test(html)) {
            return html.replace(/<\/head>/i, `${baseStyle}</head>`);
        }
        return html.replace(/<html([^\u003e]*)>/i, `<html$1><head>${baseStyle}</head>`);
    }

    return `
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  ${baseStyle}
</head>
<body>${safeBody}</body>
</html>`;
}

export function buildSafeAiPptFrameHtml(html, { headExtras = '' } = {}) {
    const {
        bodyHtml,
        headStyles,
        htmlAttributes,
        bodyAttributes,
    } = sanitizeAiPptDocumentForExport(html);
    return `
<!doctype html>
<html${htmlAttributes}>
<head>
  <meta charset="utf-8" />
  <meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src 'self' data: blob:; style-src 'unsafe-inline'; font-src 'self' data:; script-src 'none'; frame-src 'none'; object-src 'none'; base-uri 'none';">
  ${buildBaseStyle()}
  ${headStyles}
  ${buildSlideRootNormalizationStyle()}
  ${headExtras || ''}
</head>
<body${bodyAttributes}>${bodyHtml}</body>
</html>`;
}

export function createExportFrame(html, { signal, frameId } = {}) {
    throwIfExportAborted(signal);
    const frame = document.createElement('iframe');
    frame.id = frameId || `${AI_PPT_EXPORT_LAYOUT.EXPORT_FRAME_ID}-${exportFrameSequence += 1}`;
    frame.setAttribute('sandbox', 'allow-same-origin');
    frame.style.position = 'fixed';
    frame.style.left = '-9999px';
    frame.style.top = '-9999px';
    frame.style.width = `${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH}px`;
    frame.style.height = `${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT}px`;
    frame.style.pointerEvents = 'none';
    frame.style.border = '0';
    frame.srcdoc = buildSafeAiPptFrameHtml(html);
    document.body.appendChild(frame);

    return new Promise((resolve, reject) => {
        const cleanup = () => {
            window.clearTimeout(timer);
            frame.onload = null;
            if (signal) signal.removeEventListener('abort', onAbort);
        };
        const onAbort = () => {
            cleanup();
            frame.remove();
            reject(createExportAbortError());
        };
        const timer = window.setTimeout(() => {
            cleanup();
            frame.remove();
            reject(new Error('HTML 渲染超时'));
        }, 8000);
        frame.onload = () => {
            cleanup();
            resolve(frame);
        };
        if (signal) signal.addEventListener('abort', onAbort, { once: true });
        if (signal?.aborted) onAbort();
    });
}

export function getSlideRoot(doc) {
    return doc.querySelector('.ppt-slide') ||
        doc.querySelector('[data-slide-root]') ||
        doc.querySelector('[data-pptx="slide"]') ||
        doc.body.firstElementChild ||
        doc.body;
}

function withTimeout(view, promise, label, timeoutMs = FRAME_ASSET_TIMEOUT_MS, signal) {
    let timer = null;
    return new Promise((resolve, reject) => {
        const cleanup = () => {
            view.clearTimeout(timer);
            if (signal) signal.removeEventListener('abort', onAbort);
        };
        const onAbort = () => {
            cleanup();
            reject(createExportAbortError());
        };
        timer = view.setTimeout(() => {
            cleanup();
            reject(new Error(`${label}加载超时`));
        }, timeoutMs);
        Promise.resolve(promise).then(
            value => {
                cleanup();
                resolve(value);
            },
            error => {
                cleanup();
                reject(error);
            },
        );
        if (signal) signal.addEventListener('abort', onAbort, { once: true });
        if (signal?.aborted) onAbort();
    });
}

function waitForAnimationFrame(view, signal) {
    return new Promise(resolve => {
        let frameId = null;
        let timer = null;
        const cleanup = () => {
            if (frameId !== null && typeof view.cancelAnimationFrame === 'function') view.cancelAnimationFrame(frameId);
            view.clearTimeout(timer);
            if (signal) signal.removeEventListener('abort', onAbort);
        };
        const done = () => {
            cleanup();
            resolve();
        };
        const onAbort = () => {
            cleanup();
            resolve();
        };
        if (typeof view.requestAnimationFrame === 'function') {
            frameId = view.requestAnimationFrame(done);
        }
        // 后台标签页可能暂停 requestAnimationFrame；定时器确保稳定检测
        // 仍会在截止时间内结束，而不会永久卡住。
        timer = view.setTimeout(done, FRAME_ANIMATION_FRAME_FALLBACK_MS);
        if (signal) signal.addEventListener('abort', onAbort, { once: true });
        if (signal?.aborted) onAbort();
    });
}

function isVisibleLayoutElement(element, view) {
    const tagName = String(element?.tagName || '').toLowerCase();
    if (!element || ['style', 'script', 'link', 'meta'].includes(tagName)) return false;
    const style = view.getComputedStyle(element);
    return style.display !== 'none' && style.visibility !== 'hidden';
}

export function getLayoutFingerprint(root) {
    if (!root) return '';
    const view = root.ownerDocument?.defaultView || window;
    const elements = [root, ...Array.from(root.querySelectorAll('*'))];
    return elements
        .filter(element => isVisibleLayoutElement(element, view))
        .map((element, index) => {
            const rect = element.getBoundingClientRect();
            return [
                index,
                element.tagName,
                rect.left.toFixed(2),
                rect.top.toFixed(2),
                rect.width.toFixed(2),
                rect.height.toFixed(2),
            ].join(':');
        })
        .join('|');
}

export async function waitForStableLayout(doc, root, stableFrameCount = REQUIRED_STABLE_LAYOUT_FRAMES, {
    signal,
    deadlineAt,
    timeoutMs = FRAME_LAYOUT_TIMEOUT_MS,
    maxSamples = FRAME_LAYOUT_MAX_SAMPLES,
} = {}) {
    const view = doc.defaultView || window;
    const startedAt = Date.now();
    let previousFingerprint = null;
    let consecutiveFrames = 0;
    let samples = 0;

    while (consecutiveFrames < stableFrameCount && samples < maxSamples) {
        throwIfExportAborted(signal);
        if (Date.now() - startedAt >= timeoutMs || (deadlineAt && Date.now() >= deadlineAt)) break;
        await waitForAnimationFrame(view, signal);
        throwIfExportAborted(signal);
        const fingerprint = getLayoutFingerprint(root);
        consecutiveFrames = fingerprint === previousFingerprint
            ? consecutiveFrames + 1
            : 1;
        previousFingerprint = fingerprint;
        samples += 1;
    }

    return previousFingerprint || getLayoutFingerprint(root);
}

function waitForImageDecode(image, label) {
    if (image.complete) {
        if (!image.naturalWidth && !String(image.currentSrc || image.src || '').startsWith('data:image/')) {
            return Promise.reject(new Error(`${label}加载失败`));
        }
        return image.decode ? image.decode() : Promise.resolve();
    }

    return new Promise((resolve, reject) => {
        image.onload = () => {
            if (!image.naturalWidth && !String(image.currentSrc || image.src || '').startsWith('data:image/')) {
                reject(new Error(`${label}加载失败`));
                return;
            }
            if (!image.decode) {
                resolve();
                return;
            }
            image.decode().then(resolve, reject);
        };
        image.onerror = () => reject(new Error(`${label}加载失败`));
    });
}

function collectBackgroundUrls(doc) {
    const backgroundUrls = new Set();
    const urlPattern = /url\((['"]?)(.*?)\1\)/gi;
    const collectBackgroundUrlsFromStyle = style => {
        const values = [style?.backgroundImage, style?.maskImage, style?.webkitMaskImage];
        values.forEach(value => {
            let match;
            while ((match = urlPattern.exec(String(value || ''))) !== null) {
                if (match[2] && match[2] !== 'none') backgroundUrls.add(match[2]);
            }
            urlPattern.lastIndex = 0;
        });
    };

    Array.from(doc.querySelectorAll('*')).forEach(element => {
        collectBackgroundUrlsFromStyle(doc.defaultView.getComputedStyle(element));
        collectBackgroundUrlsFromStyle(doc.defaultView.getComputedStyle(element, '::before'));
        collectBackgroundUrlsFromStyle(doc.defaultView.getComputedStyle(element, '::after'));
    });
    return backgroundUrls;
}

function hasLoadableImageSource(image) {
    return Boolean(String(image?.getAttribute?.('src') || image?.getAttribute?.('srcset') || image?.currentSrc || '').trim());
}

function reportDegradedResources(results, label) {
    const failedCount = results.filter(result => result.status === 'rejected').length;
    if (failedCount) console.warn(`[AiPPT] ${failedCount} 个${label}未加载，已按当前可用内容继续导出。`);
}

export async function waitForFrameAssets(doc, root = getSlideRoot(doc), { signal, deadlineAt } = {}) {
    const view = doc.defaultView || window;
    throwIfExportAborted(signal);

    if (doc.fonts?.ready) {
        await withTimeout(view, doc.fonts.ready, '字体', getRemainingTimeout(deadlineAt, FRAME_ASSET_TIMEOUT_MS), signal);
    }

    const images = Array.from(doc.images || []).filter(hasLoadableImageSource);
    images.forEach(image => image.setAttribute('loading', 'eager'));
    const imageResults = await Promise.allSettled(images.map((image, index) => (
        withTimeout(view, waitForImageDecode(image, `第 ${index + 1} 张图片`), `第 ${index + 1} 张图片`, getRemainingTimeout(deadlineAt, FRAME_ASSET_TIMEOUT_MS), signal)
    )));
    if (signal?.aborted) throwIfExportAborted(signal);
    reportDegradedResources(imageResults, '图片');

    const backgroundUrls = collectBackgroundUrls(doc);
    const backgroundResults = await Promise.allSettled(Array.from(backgroundUrls).map((url, index) => {
        const image = new view.Image();
        image.src = url;
        return withTimeout(
            view,
            waitForImageDecode(image, `第 ${index + 1} 张背景图`),
            `第 ${index + 1} 张背景图`,
            getRemainingTimeout(deadlineAt, FRAME_ASSET_TIMEOUT_MS),
            signal,
        );
    }));
    if (signal?.aborted) throwIfExportAborted(signal);
    reportDegradedResources(backgroundResults, '背景图');

    return waitForStableLayout(doc, root, REQUIRED_STABLE_LAYOUT_FRAMES, { signal, deadlineAt });
}

export function normalizeSlideRoot(doc, root) {
    const rootStyle = root.style;
    rootStyle.position = 'absolute';
    rootStyle.left = '0';
    rootStyle.top = '0';
    rootStyle.width = `${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH}px`;
    rootStyle.height = `${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT}px`;
    rootStyle.margin = '0';
    rootStyle.boxShadow = 'none';
    rootStyle.overflow = 'hidden';

    doc.documentElement.style.margin = '0';
    doc.documentElement.style.padding = '0';
    doc.documentElement.style.width = `${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH}px`;
    doc.documentElement.style.height = `${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT}px`;
    doc.documentElement.style.overflow = 'hidden';

    doc.body.style.margin = '0';
    doc.body.style.padding = '0';
    doc.body.style.width = `${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH}px`;
    doc.body.style.height = `${AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT}px`;
    doc.body.style.overflow = 'hidden';
}
