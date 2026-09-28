import { AI_PPT_EXPORT_LAYOUT } from '../../constants';
import {
    createExportFrame,
    getSlideRoot,
    isExportAbortError,
    normalizeSlideRoot,
    throwIfExportAborted,
    waitForFrameAssets,
} from './domParser';
import { renderBackgroundImage } from './imageExport';
import { addNativeVisualElementsToSlide } from './elementMapper';
import {
    alignToPptx,
    applyTextTransform,
    clamp,
    extractRunStyle,
    getRotateDeg,
    isBlockLikeChild,
    isBoldWeight,
    mapFont,
    normalizeCssColor,
} from './styleExtractor';

const SKIP_TEXT_TAGS = new Set(['SCRIPT', 'STYLE', 'LINK', 'META', 'HEAD', 'NOSCRIPT']);
const PT_PER_PX = (AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN / AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH) * 72;
const SCALED_PT_PER_PX = PT_PER_PX * AI_PPT_EXPORT_LAYOUT.FONT_SIZE_SCALE;

function isInvisibleStyle(style) {
    const opacity = String(style.opacity || '').trim();
    return style.display === 'none' ||
        style.visibility === 'hidden' ||
        (opacity !== '' && Number(opacity) === 0);
}

function isWithinSkippedTextElement(node, skippedTextElements) {
    if (!skippedTextElements?.size) return false;
    let current = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
    while (current) {
        if (skippedTextElements.has(current)) return true;
        current = current.parentElement;
    }
    return false;
}

function collectLineFragments(doc, startNode, endNode, skippedTextElements) {
    const view = doc.defaultView;
    const textNodes = [];
    const visit = (node) => {
        if (isWithinSkippedTextElement(node, skippedTextElements)) return;
        if (node.nodeType === Node.TEXT_NODE) {
            if (node.textContent) textNodes.push(node);
            return;
        }
        if (node.nodeType !== Node.ELEMENT_NODE) return;
        if (SKIP_TEXT_TAGS.has(node.tagName)) return;
        const style = view.getComputedStyle(node);
        if (isInvisibleStyle(style)) return;
        Array.from(node.childNodes || []).forEach(visit);
    };
    let current = startNode;
    while (current) {
        visit(current);
        if (current === endNode) break;
        current = current.nextSibling;
    }

    const lines = [];
    for (const textNode of textNodes) {
        const styleElement = textNode.parentElement;
        const text = textNode.textContent;
        for (let i = 0; i < text.length; i++) {
            let rect;
            try {
                const range = doc.createRange();
                range.setStart(textNode, i);
                range.setEnd(textNode, i + 1);
                rect = range.getClientRects()[0];
            } catch (error) {
                rect = null;
            }
            if (!rect || (!rect.width && !rect.height)) continue;
            const last = lines[lines.length - 1];
            if (!last || (rect.top - last.anchorTop) > Math.max(3, (last.bottom - last.anchorTop) * 0.6)) {
                lines.push({
                    anchorTop: rect.top,
                    top: rect.top,
                    bottom: rect.bottom,
                    left: rect.left,
                    right: rect.right,
                    cells: [{ char: text[i], styleElement }],
                });
            } else {
                last.top = Math.min(last.top, rect.top);
                last.bottom = Math.max(last.bottom, rect.bottom);
                last.left = Math.min(last.left, rect.left);
                last.right = Math.max(last.right, rect.right);
                last.cells.push({ char: text[i], styleElement });
            }
        }
    }

    // FALLBACK: 逐字符 getClientRects 失败时，用 textNode 级别的 bounding rect 兜底
    if (!lines.length) {
        for (const textNode of textNodes) {
            const text = textNode.textContent;
            if (!text || !text.trim()) continue;
            try {
                const range = doc.createRange();
                range.selectNode(textNode);
                const rect = range.getBoundingClientRect();
                if (rect && rect.width > 0 && rect.height > 0) {
                    const styleElement = textNode.parentElement;
                    lines.push({
                        anchorTop: rect.top,
                        top: rect.top,
                        bottom: rect.bottom,
                        left: rect.left,
                        right: rect.right,
                        cells: text.split('').map(char => ({ char, styleElement })),
                    });
                }
            } catch (e) {
                // ignore
            }
        }
    }

    return lines.map(line => {
        const runs = [];
        for (const cell of line.cells) {
            const tail = runs[runs.length - 1];
            if (tail && tail.styleElement === cell.styleElement) {
                tail.text += cell.char;
            } else {
                runs.push({ text: cell.char, styleElement: cell.styleElement });
            }
        }
        let normalized = runs.map(run => ({ text: run.text.replace(/\s+/g, ' '), styleElement: run.styleElement }));
        if (normalized.length) normalized[0].text = normalized[0].text.replace(/^\s+/, '');
        if (normalized.length) normalized[normalized.length - 1].text = normalized[normalized.length - 1].text.replace(/\s+$/, '');
        normalized = normalized.filter(run => run.text.length > 0);
        return {
            left: line.left,
            top: line.top,
            width: line.right - line.left,
            height: line.bottom - line.top,
            runs: normalized,
        };
    }).filter(line => line.runs.length);
}

function collectVisibleTextFromRun(doc, startNode, endNode, skippedTextElements) {
    const view = doc.defaultView;
    const parts = [];

    function visit(node) {
        if (isWithinSkippedTextElement(node, skippedTextElements)) return;
        if (node.nodeType === Node.TEXT_NODE) {
            if (node.textContent) parts.push(node.textContent);
            return;
        }
        if (node.nodeType !== Node.ELEMENT_NODE) return;
        if (SKIP_TEXT_TAGS.has(node.tagName)) return;
        const style = view.getComputedStyle(node);
        if (isInvisibleStyle(style)) return;
        Array.from(node.childNodes || []).forEach(visit);
    }

    let current = startNode;
    while (current) {
        visit(current);
        if (current === endNode) break;
        current = current.nextSibling;
    }

    return parts.join('').replace(/\s+/g, ' ').trim();
}

function buildElementTextFallback(doc, parentElement, startNode, endNode, baseRect, skippedTextElements, root) {
    const text = collectVisibleTextFromRun(doc, startNode, endNode, skippedTextElements);
    if (!text) return null;

    const rect = parentElement.getBoundingClientRect();
    if (!rect || rect.width <= 0 || rect.height <= 0) return null;

    const style = doc.defaultView.getComputedStyle(parentElement);
    const fontPx = Number.parseFloat(style.fontSize) || 14;
    const lineHeightPx = (() => {
        if (style.lineHeight === 'normal') return fontPx * 1.2;
        const value = Number.parseFloat(style.lineHeight);
        return Number.isFinite(value) ? value : fontPx * 1.2;
    })();
    const numericBadgeBox = getCenteredNumericBadgeBox(parentElement, root, baseRect, text);
    const rotateDeg = getRotateDeg(style);
    const rotatedTextBox = getUnrotatedTextBox(parentElement, baseRect, rotateDeg);

    return {
        runs: [{
            text,
            ...extractRunStyle(parentElement),
        }],
        x: numericBadgeBox?.x ?? rotatedTextBox?.x ?? rect.left - baseRect.left,
        y: numericBadgeBox?.y ?? rotatedTextBox?.y ?? rect.top - baseRect.top,
        w: numericBadgeBox?.w ?? rotatedTextBox?.w ?? rect.width,
        h: numericBadgeBox?.h ?? rotatedTextBox?.h ?? rect.height,
        textAlign: numericBadgeBox ? 'center' : style.textAlign,
        lineSpacingPt: lineHeightPx * SCALED_PT_PER_PX,
        parentFontPx: fontPx,
        rotateDeg,
        singleLine: false,
        valign: numericBadgeBox ? 'mid' : 'top',
    };
}

function getUnrotatedTextBox(element, baseRect, rotateDeg) {
    if (!rotateDeg) return null;

    const width = Number(element.offsetWidth);
    const height = Number(element.offsetHeight);
    if (!(width > 0 && height > 0)) return null;

    const rect = element.getBoundingClientRect();
    const rectRight = rect.right ?? (rect.left + rect.width);
    const rectBottom = rect.bottom ?? (rect.top + rect.height);
    const centerX = (rect.left + rectRight) / 2;
    const centerY = (rect.top + rectBottom) / 2;

    return {
        x: centerX - (width / 2) - baseRect.left,
        y: centerY - (height / 2) - baseRect.top,
        w: width,
        h: height,
    };
}

function shouldExtractElementSeparately(element) {
    const style = element.ownerDocument.defaultView.getComputedStyle(element);
    return [
        'inline-block',
        'inline-flex',
        'inline-grid',
    ].includes(style.display);
}

function getCenteredNumericBadgeBox(parentElement, root, baseRect, text) {
    if (!/^\d{1,2}$/.test(String(text || '').trim())) return null;

    let current = parentElement;
    while (current && current !== root.parentElement) {
        const style = current.ownerDocument.defaultView.getComputedStyle(current);
        const rect = current.getBoundingClientRect();
        const isSquare = rect.width > 0 && rect.height > 0 &&
            Math.abs(rect.width - rect.height) <= Math.max(2, Math.min(rect.width, rect.height) * 0.12);
        const radius = String(style.borderRadius || style.borderTopLeftRadius || '');
        const usesCircularRadius = /(?:^|\s)50%(?:\s|$)/.test(radius);

        if (isSquare && rect.width <= 180 && usesCircularRadius) {
            // Centering a standalone number inside a circular badge fixes the
            // browser-to-PPT glyph-anchor mismatch. A badge with a subtitle
            // (for example "01" + "MODEL") is a two-line composition instead;
            // centering only its number over the whole circle makes it overlap
            // the measured subtitle line in Office/WPS.
            const visibleText = collectVisibleTextFromRun(
                current.ownerDocument,
                current,
                current,
                new Set(),
            );
            if (visibleText !== String(text).trim()) {
                if (current === root) break;
                current = current.parentElement;
                continue;
            }
            return {
                x: rect.left - baseRect.left,
                y: rect.top - baseRect.top,
                w: rect.width,
                h: rect.height,
            };
        }
        if (current === root) break;
        current = current.parentElement;
    }
    return null;
}

function extractTextsAndHide(root, { skippedTextElements = new Set(), hide = false } = {}) {
    const doc = root.ownerDocument;
    const baseRect = root.getBoundingClientRect();
    const texts = [];
    const transparentTargets = new Set();

    function emitInlineRun(parentElement, startNode, endNode) {
        const lines = collectLineFragments(doc, startNode, endNode, skippedTextElements);
        if (!lines.length) {
            const fallbackText = buildElementTextFallback(
                doc,
                parentElement,
                startNode,
                endNode,
                baseRect,
                skippedTextElements,
                root,
            );
            if (fallbackText) {
                texts.push(fallbackText);
                transparentTargets.add(parentElement);
            }
            return;
        }

        const parentStyle = doc.defaultView.getComputedStyle(parentElement);
        const parentFontPx = Number.parseFloat(parentStyle.fontSize) || 14;
        const lineHeightPx = (() => {
            if (parentStyle.lineHeight === 'normal') return parentFontPx * 1.2;
            const value = Number.parseFloat(parentStyle.lineHeight);
            return Number.isFinite(value) ? value : parentFontPx * 1.2;
        })();
        const textAlign = parentStyle.textAlign;
        const rotateDeg = getRotateDeg(parentStyle);
        const rotatedTextBox = lines.length === 1
            ? getUnrotatedTextBox(parentElement, baseRect, rotateDeg)
            : null;
        const styleElements = new Set();

        for (const line of lines) {
            const lineText = line.runs.map(run => run.text).join('');
            const numericBadgeBox = getCenteredNumericBadgeBox(parentElement, root, baseRect, lineText);
            texts.push({
                runs: line.runs.map(piece => ({
                    text: piece.text,
                    ...extractRunStyle(piece.styleElement || parentElement),
                })),
                x: numericBadgeBox?.x ?? rotatedTextBox?.x ?? line.left - baseRect.left,
                y: numericBadgeBox?.y ?? rotatedTextBox?.y ?? line.top - baseRect.top,
                w: numericBadgeBox?.w ?? rotatedTextBox?.w ?? line.width,
                h: numericBadgeBox?.h ?? rotatedTextBox?.h ?? line.height,
                textAlign: numericBadgeBox ? 'center' : textAlign,
                lineSpacingPt: lineHeightPx * SCALED_PT_PER_PX,
                parentFontPx,
                rotateDeg,
                singleLine: true,
                valign: numericBadgeBox ? 'mid' : 'top',
            });
            line.runs.forEach(piece => {
                if (piece.styleElement) styleElements.add(piece.styleElement);
            });
        }

        if (styleElements.size) {
            styleElements.forEach(element => transparentTargets.add(element));
        } else {
            transparentTargets.add(parentElement);
        }
    }

    function walk(element) {
        if (!element || element.nodeType !== 1) return;
        if (SKIP_TEXT_TAGS.has(element.tagName)) return;
        if (isWithinSkippedTextElement(element, skippedTextElements)) return;

        const style = doc.defaultView.getComputedStyle(element);
        if (isInvisibleStyle(style)) {
            return;
        }

        let runStart = null;
        let runEnd = null;
        const flush = () => {
            if (runStart) {
                emitInlineRun(element, runStart, runEnd);
                runStart = null;
                runEnd = null;
            }
        };

        Array.from(element.childNodes || []).forEach(child => {
            if (child.nodeType === Node.TEXT_NODE) {
                if (child.textContent && child.textContent.trim()) {
                    if (!runStart) runStart = child;
                    runEnd = child;
                } else if (runStart) {
                    runEnd = child;
                }
                return;
            }

            if (child.nodeType !== Node.ELEMENT_NODE) return;

            if (shouldExtractElementSeparately(child)) {
                flush();
                walk(child);
                return;
            }

            if (isBlockLikeChild(child)) {
                flush();
                walk(child);
            } else {
                if (!runStart) runStart = child;
                runEnd = child;
            }
        });

        flush();
    }

    walk(root);

    if (hide) {
        transparentTargets.forEach(element => {
            element.style.setProperty('color', 'transparent', 'important');
            element.style.setProperty('text-shadow', 'none', 'important');
            if (element.tagName?.toLowerCase() === 'text') {
                element.style.setProperty('fill', 'transparent', 'important');
            }
        });
    }

    return {
        slideRect: {
            x: baseRect.left,
            y: baseRect.top,
            w: baseRect.width,
            h: baseRect.height,
        },
        texts,
    };
}

function hasVisibleTextContent(root, skippedTextElements = new Set()) {
    const doc = root.ownerDocument;
    const view = doc.defaultView;
    let found = false;

    function walk(node) {
        if (found) return;
        if (isWithinSkippedTextElement(node, skippedTextElements)) return;
        if (node.nodeType === Node.TEXT_NODE) {
            if (node.textContent && node.textContent.trim()) {
                found = true;
            }
            return;
        }
        if (node.nodeType !== Node.ELEMENT_NODE) return;
        if (SKIP_TEXT_TAGS.has(node.tagName)) return;
        const style = view.getComputedStyle(node);
        if (isInvisibleStyle(style)) return;
        Array.from(node.childNodes || []).forEach(walk);
    }

    walk(root);
    return found;
}

function addExtractedTextsToSlide(pptSlide, texts, pxToInH, pxToInV) {
    let addedCount = 0;
    let failedCount = 0;

    for (const textBox of texts) {
        if (!textBox.runs?.length) continue;

        const maxFontPt = textBox.runs.reduce(
            (max, run) => Math.max(max, run.fontSizePt || 0),
            0,
        ) || 12;
        const isMultiLine = !textBox.singleLine
            && textBox.h > textBox.parentFontPx * AI_PPT_EXPORT_LAYOUT.MULTILINE_HEIGHT_FACTOR;
        const align = alignToPptx(textBox.textAlign);
        const textWidthIn = pxToInH(textBox.w);
        const widthFactor = isMultiLine
            ? AI_PPT_EXPORT_LAYOUT.MULTILINE_WIDTH_FACTOR
            : AI_PPT_EXPORT_LAYOUT.SINGLE_LINE_WIDTH_FACTOR;
        const padWidth = isMultiLine
            ? AI_PPT_EXPORT_LAYOUT.MULTILINE_PAD_IN
            : AI_PPT_EXPORT_LAYOUT.SINGLE_LINE_PAD_IN;
        const leftIn = pxToInH(textBox.x);
        const rightIn = pxToInH(textBox.x + textBox.w);
        const pageWidth = AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN;
        const boundedLeftIn = clamp(leftIn, 0, pageWidth);
        const boundedRightIn = clamp(rightIn, 0, pageWidth);
        const requestedWidthIn = Math.max(0.4, textWidthIn * widthFactor + padWidth);
        const centerIn = clamp((boundedLeftIn + boundedRightIn) / 2, 0, pageWidth);
        const maxWidthIn = align === 'right'
            ? boundedRightIn
            : align === 'center'
                ? Math.min(centerIn, pageWidth - centerIn) * 2
                : pageWidth - boundedLeftIn;
        // Do not let a width safety margin move a left-aligned text run to
        // x=0. The browser's measured start position is the visual anchor.
        const wIn = Math.max(0.4, Math.min(requestedWidthIn, Math.max(0.4, maxWidthIn)));
        const lineInchPad = (maxFontPt / 72) * 1.2;
        const hMultiplier = isMultiLine ? 1.35 : 1.0;
        const hIn = Math.max(0.22, pxToInV(textBox.h) * hMultiplier + (isMultiLine ? lineInchPad * 2 : 0.06));
        let xIn = leftIn;

        if (align === 'right') {
            xIn = rightIn - wIn;
        } else if (align === 'center') {
            xIn = leftIn + (textWidthIn / 2) - (wIn / 2);
        }

        const textOptions = {
            x: clamp(xIn, 0, Math.max(0, pageWidth - wIn)),
            y: clamp(pxToInV(textBox.y), 0, AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN - hIn),
            w: wIn,
            h: hIn,
            align,
            valign: textBox.valign || 'top',
            margin: 0,
            isTextBox: true,
            fit: 'shrink',
            wrap: isMultiLine,
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
        };
        if (textBox.lineSpacingPt) {
            textOptions.lineSpacing = textBox.lineSpacingPt;
        }
        if (textBox.rotateDeg) {
            textOptions.rotate = Math.round(textBox.rotateDeg);
        }

        const runSpec = textBox.runs.map(run => ({
            text: applyTextTransform(run.text, run.textTransform),
            options: {
                fontFace: mapFont(run.fontFamily),
                fontSize: run.fontSizePt,
                color: normalizeCssColor(run.color) || '333333',
                bold: isBoldWeight(run.fontWeight),
                italic: /italic/i.test(run.fontStyle || ''),
                charSpacing: run.letterSpacingPt || 0,
                breakLine: false,
                ...(run.transparency ? { transparency: run.transparency } : {}),
                ...(run.underline ? { underline: run.underline } : {}),
            },
        }));

        try {
            pptSlide.addText(runSpec, textOptions);
            addedCount += 1;
        } catch (error) {
            failedCount += 1;
            console.warn('[AiPPT] 文本框导出失败，已跳过:', error);
        }
    }

    return { addedCount, failedCount };
}

async function addVisualFallbackSlideImage(pptSlide, htmlSlide, { signal, deadlineAt } = {}) {
    const frame = await createExportFrame(htmlSlide.html, { signal });

    try {
        const doc = frame.contentDocument;
        const root = getSlideRoot(doc);
        normalizeSlideRoot(doc, root);
        await waitForFrameAssets(doc, root, { signal, deadlineAt });
        throwIfExportAborted(signal);

        const slideRect = root.getBoundingClientRect();
        const bgDataUrl = await renderBackgroundImage(doc.body, {
            x: slideRect.left,
            y: slideRect.top,
            w: AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH,
            h: AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT,
        }, { requireVisiblePixels: true });
        pptSlide.addImage({
            data: bgDataUrl,
            x: 0,
            y: 0,
            w: AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN,
            h: AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN,
        });
    } finally {
        frame.remove();
    }
}

// Visual-text export is the production-safe mode for arbitrary AI generated
// HTML. The browser is the only renderer that can faithfully compose every
// CSS radius, overflow clip, pseudo-element and transform. Keep that complete
// visual result as one slide image, then place the extracted browser-measured
// text back as native editable PowerPoint text. In particular, do not route
// decorations through the HTML-to-DrawingML shape guesser.
export async function addHtmlSlideToPptxWithEditableTextOverlay(pptSlide, htmlSlide, { signal, deadlineAt } = {}) {
    const frame = await createExportFrame(htmlSlide.html, { signal });
    const operationBuffer = createSlideOperationBuffer();
    const slideObjectCheckpoint = getSlideObjectCheckpoint(pptSlide);
    let needsVisualFallback = false;
    const fallbackReasons = [];
    let textExportResult = null;

    try {
        const doc = frame.contentDocument;
        const root = getSlideRoot(doc);
        normalizeSlideRoot(doc, root);
        await waitForFrameAssets(doc, root, { signal, deadlineAt });
        throwIfExportAborted(signal);

        const slideRect = root.getBoundingClientRect();
        const pxPerInH = (slideRect.width || AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH) / AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN;
        const pxPerInV = (slideRect.height || AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT) / AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN;
        const hadVisibleText = hasVisibleTextContent(root);
        const { texts } = extractTextsAndHide(root, { hide: true });

        if (hadVisibleText && !texts.length) {
            needsVisualFallback = true;
            fallbackReasons.push('text-not-extracted');
            console.warn(`[AiPPT] 第 ${htmlSlide.page || '?'} 页未能提取可编辑文字，将保留完整浏览器视觉页。`);
        } else {
            const visualData = await renderBackgroundImage(doc.body, {
                x: slideRect.left,
                y: slideRect.top,
                w: AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH,
                h: AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT,
            }, { requireVisiblePixels: true });
            operationBuffer.slide.addImage({
                data: visualData,
                x: 0,
                y: 0,
                w: AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN,
                h: AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN,
            });
            textExportResult = addExtractedTextsToSlide(
                operationBuffer.slide,
                texts,
                px => px / pxPerInH,
                px => px / pxPerInV,
            );
            if (textExportResult.failedCount > 0) {
                needsVisualFallback = true;
                fallbackReasons.push('text-write-failed');
            }
        }
    } catch (error) {
        if (isExportAbortError(error)) throw error;
        needsVisualFallback = true;
        fallbackReasons.push('visual-background-failed');
        console.warn(`[AiPPT] 第 ${htmlSlide.page || '?'} 页浏览器视觉背景导出失败，将按完整视觉页兜底:`, error);
    } finally {
        frame.remove();
    }

    if (!needsVisualFallback) {
        try {
            const { failedVisualOperations } = operationBuffer.replay(pptSlide);
            if (failedVisualOperations.length) {
                rollbackSlideObjects(pptSlide, slideObjectCheckpoint);
                needsVisualFallback = true;
                fallbackReasons.push('visual-background-write-failed');
            }
        } catch (error) {
            rollbackSlideObjects(pptSlide, slideObjectCheckpoint);
            needsVisualFallback = true;
            fallbackReasons.push('text-write-failed');
            console.warn(`[AiPPT] 第 ${htmlSlide.page || '?'} 页可编辑文字写入失败，将按完整视觉页兜底:`, error);
        }
    }

    if (needsVisualFallback) {
        rollbackSlideObjects(pptSlide, slideObjectCheckpoint);
        await addVisualFallbackSlideImage(pptSlide, htmlSlide, { signal, deadlineAt });
    }

    return {
        page: htmlSlide.page,
        nativeVisualCount: 0,
        editableTextCount: textExportResult?.addedCount || 0,
        rasterizedVisualCount: needsVisualFallback ? 0 : 1,
        visualFallback: needsVisualFallback,
        visualTextOverlay: !needsVisualFallback,
        fallbackReasons,
    };
}

function createSlideOperationBuffer() {
    const operations = [];
    const buffer = {};
    ['addImage', 'addShape', 'addTable', 'addChart', 'addText'].forEach(method => {
        buffer[method] = (...args) => {
            operations.push({ method, args });
        };
    });
    return {
        slide: buffer,
        replay(target) {
            const failedVisualOperations = [];
            for (const operation of operations) {
                try {
                    target[operation.method](...operation.args);
                } catch (error) {
                    error.aiPptOperation = operation.method;
                    // Keep the rest of the slide editable when one visual
                    // object is unsupported by the target PPT renderer.
                    // Text is intentionally stricter: losing it would break
                    // the editable-content contract for the whole slide.
                    if (operation.method === 'addText') throw error;
                    failedVisualOperations.push(operation.method);
                    console.warn(`[AiPPT] PPT ${operation.method} 写入失败，已跳过该视觉对象:`, error);
                }
            }
            return { failedVisualOperations };
        },
    };
}

function getSlideObjectCheckpoint(pptSlide) {
    return Array.isArray(pptSlide?._slideObjects) ? pptSlide._slideObjects.length : null;
}

function rollbackSlideObjects(pptSlide, checkpoint) {
    if (checkpoint === null || !Array.isArray(pptSlide?._slideObjects)) return;
    pptSlide._slideObjects.splice(checkpoint);
}

export async function addHtmlSlideToPptx(pptSlide, htmlSlide, { signal, deadlineAt } = {}) {
    const frame = await createExportFrame(htmlSlide.html, { signal });
    const operationBuffer = createSlideOperationBuffer();
    const slideObjectCheckpoint = getSlideObjectCheckpoint(pptSlide);
    let needsVisualFallback = false;
    const fallbackReasons = [];
    let nativeVisualResult = null;
    let textExportResult = null;

    try {
        const doc = frame.contentDocument;
        const root = getSlideRoot(doc);
        normalizeSlideRoot(doc, root);
        await waitForFrameAssets(doc, root, { signal, deadlineAt });
        throwIfExportAborted(signal);

        const slideRect = root.getBoundingClientRect();
        const pxPerInH = (slideRect.width || AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_WIDTH) / AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN;
        const pxPerInV = (slideRect.height || AI_PPT_EXPORT_LAYOUT.SLIDE_PIXEL_HEIGHT) / AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN;
        nativeVisualResult = await addNativeVisualElementsToSlide({
            pptSlide: operationBuffer.slide,
            root,
            slideRect,
            pxToInH: px => px / pxPerInH,
            pxToInV: px => px / pxPerInV,
        });
        const { texts } = extractTextsAndHide(root, {
            skippedTextElements: nativeVisualResult.skippedTextElements,
        });
        textExportResult = addExtractedTextsToSlide(
            operationBuffer.slide,
            texts,
            px => px / pxPerInH,
            px => px / pxPerInV,
        );

        const hasVisibleText = hasVisibleTextContent(root, nativeVisualResult.skippedTextElements);
        if (hasVisibleText && !texts.length) {
            fallbackReasons.push('text-not-extracted');
            console.warn(`[AiPPT] 第 ${htmlSlide.page || '?'} 页未能提取全部可编辑文本，将按整页视觉保真导出。`);
        }
        if (textExportResult.failedCount > 0) {
            fallbackReasons.push('text-write-failed');
            console.warn(`[AiPPT] 第 ${htmlSlide.page || '?'} 页部分可编辑文本导出失败，将按整页视觉保真导出。`);
        }
        if (nativeVisualResult.failures.length) {
            // Keep successfully mapped text and native objects editable. Failed
            // visual elements are handled as local degradations, never by
            // converting the whole page into an image.
            console.warn(`[AiPPT] 第 ${htmlSlide.page || '?'} 页存在未映射视觉元素，已保留可编辑文字和其他原生对象。`);
        }
        if (!nativeVisualResult.addedCount && !textExportResult.addedCount) {
            fallbackReasons.push('no-native-content');
        }
        if (fallbackReasons.length) {
            needsVisualFallback = true;
            console.warn(`[AiPPT] 第 ${htmlSlide.page || '?'} 页已按整页视觉兜底导出。`);
        }
    } finally {
        frame.remove();
    }

    if (!needsVisualFallback) {
        try {
            const { failedVisualOperations } = operationBuffer.replay(pptSlide);
            failedVisualOperations.forEach(operation => {
                fallbackReasons.push(`${operation}-write-failed`);
            });
        } catch (error) {
            rollbackSlideObjects(pptSlide, slideObjectCheckpoint);
            needsVisualFallback = true;
            const operation = error.aiPptOperation || 'unknown';
            fallbackReasons.push(operation === 'addText' ? 'text-write-failed' : `${operation}-write-failed`);
            console.warn(`[AiPPT] 第 ${htmlSlide.page || '?'} 页 PPT 对象写入失败，将按整页视觉兜底导出:`, error);
        }
    }

    if (needsVisualFallback) {
        rollbackSlideObjects(pptSlide, slideObjectCheckpoint);
        await addVisualFallbackSlideImage(pptSlide, htmlSlide, { signal, deadlineAt });
    }

    return {
        page: htmlSlide.page,
        nativeVisualCount: nativeVisualResult?.contentVisualCount || 0,
        editableTextCount: textExportResult?.addedCount || 0,
        rasterizedVisualCount: nativeVisualResult?.rasterizedVisualCount || 0,
        visualFallback: needsVisualFallback,
        fallbackReasons: [
            ...fallbackReasons,
            ...(nativeVisualResult?.degradations || []),
            ...(nativeVisualResult?.failures || []),
        ],
    };
}
