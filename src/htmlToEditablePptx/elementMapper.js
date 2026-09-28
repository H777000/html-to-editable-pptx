import { cropRasterImage, renderBackgroundImage } from './imageExport';
import {
    clipsOutOfFlowElement,
    getElementPosition as calculateElementPosition,
    getSlideBoundedPosition,
    getTransformedLocalBounds as calculateTransformedLocalBounds,
    hasVerifiedVisibleBounds,
    restoreUnrotatedPosition,
} from './elementGeometry';
import { addCroppedBrowserCapture } from './browserCaptureOutput';
import { captureBrowserComposedVisual } from './browserComposedCapture';
import {
    capturePseudoElementLayer,
    getRasterCaptureOptions as getPseudoRasterCaptureOptions,
    inlineAtomicVisualSubtreeStyles as inlinePseudoAtomicVisualSubtreeStyles,
    inlineRasterizedVisualStyle as inlinePseudoRasterizedVisualStyle,
    isInBrowserComposedCaptureTree as isInPseudoCaptureTree,
    mountCloneInIsolatedContext as mountPseudoCloneInIsolatedContext,
    mountCloneWithAncestorContext as mountPseudoCloneWithAncestorContext,
    materializePseudoElements as materializePseudoElementsForCapture,
    suppressCloneNativePseudoElements,
    temporarilyHideCaptureText as hidePseudoCaptureText,
    temporarilyPinCaptureColor as pinPseudoCaptureColor,
    temporarilySetStyle as setPseudoTemporaryStyle,
} from './pseudoElementCapture';
import { captureComplexVisualLayer } from './complexVisualCapture';
import { createExportStrategy } from './exportStrategy';
import { writeChart } from './chartWriter';
import { writeTable } from './tableWriter';
import { writeRasterImage, writeSvg } from './mediaWriter';
import { writeShape, writeSlideRootFrame } from './shapeWriter';
import { writeCssBackgroundLayer, writeGradientPageBackground } from './backgroundWriter';
import { collectElementsInPaintOrder } from './elementTraversal';
import {
    getBorderOptions,
    getBorderRadiusCss,
    getCornerRadiusValues,
    getRoundedRectRadius,
    getShadowOptions,
    getShapeType,
    hasAsymmetricCornerRadius,
    hasAsymmetricVisibleBorder,
    isShapeCandidate,
    SHAPE_TYPES,
} from './shapeDetection';
import {
    colorTransparency,
    firstGradientColor,
    getEffectiveOpacity,
    getPageBackground,
    hasComplexCssEffect,
    hasComplexTransform,
    hasCssBackgroundImage,
    hasVisiblePseudoElement,
    isGradientText,
    isTransparentColor,
} from './visualStyle';
import {
    clamp,
    getRotateDeg,
    normalizeCssColor,
} from './styleExtractor';
import { AI_PPT_EXPORT_LAYOUT } from '../../constants';

const ARROW_TYPES = new Set(['arrow', 'diamond', 'oval', 'stealth', 'triangle', 'none']);

const exportStrategy = createExportStrategy({
    getElementPosition,
    getShapeType,
    getRotateDeg,
    hasComplexCssEffect,
    hasComplexTransform,
    hasCssBackgroundImage,
    hasAsymmetricCornerRadius,
    hasAsymmetricVisibleBorder,
    hasVisiblePseudoElement,
    isGradientText,
    isShapeCandidate,
    isTransparentColor,
    getCornerRadiusValues,
    getTagName: element => element?.tagName?.toUpperCase?.() || '',
    shapeTypes: SHAPE_TYPES,
});

function getElementPosition(element, slideRect, pxToInH, pxToInV) {
    return calculateElementPosition(element, slideRect, pxToInH, pxToInV, {
        clipsOutOfFlowElement,
        hasVerifiedVisibleBounds,
    });
}

function addShape(pptSlide, element, slideRect, pxToInH, pxToInV) {
    const style = element.ownerDocument.defaultView.getComputedStyle(element);
    if (!isShapeCandidate(element, style)) return false;
    if (isGradientText(style)) return false;
    let position = getElementPosition(element, slideRect, pxToInH, pxToInV);
    if (!position) return false;
    position = restoreUnrotatedPosition(element, position, slideRect, pxToInH, pxToInV);
    return writeShape(pptSlide, element, style, position, {
        arrowTypes: ARROW_TYPES, colorTransparency, firstGradientColor, getBorderOptions,
        getEffectiveOpacity, getRoundedRectRadius, getRotateDeg, getShadowOptions,
        getShapeType, isTransparentColor, normalizeCssColor, shapeTypes: SHAPE_TYPES,
    });
}

function addSlideRootFrame(pptSlide, root, slideRect, pxToInH, pxToInV) {
    const style = root.ownerDocument.defaultView.getComputedStyle(root);
    const position = getElementPosition(root, slideRect, pxToInH, pxToInV);
    return writeSlideRootFrame(pptSlide, root, style, position, {
        getBorderOptions, getEffectiveOpacity, getRoundedRectRadius, getShadowOptions,
        getShapeType, hasAsymmetricVisibleBorder, shapeTypes: SHAPE_TYPES,
    });
}

function addTable(pptSlide, table, slideRect, pxToInH, pxToInV) {
    const position = getElementPosition(table, slideRect, pxToInH, pxToInV);
    return writeTable(pptSlide, table, position, pxToInH, pxToInV, {
        colorTransparency,
        getBorderOptions,
        getEffectiveOpacity,
    });
}

async function addImage(pptSlide, image, slideRect, pxToInH, pxToInV, { preserveTransform = false } = {}) {
    const position = getElementPosition(image, slideRect, pxToInH, pxToInV);
    return writeRasterImage({
        pptSlide, element: image, position, preserveTransform,
        helpers: {
            getRotateDeg,
            hasComplexTransform,
            restoreUnrotatedPosition: (target, targetPosition) => restoreUnrotatedPosition(
                target, targetPosition, slideRect, pxToInH, pxToInV,
            ),
        },
    });
}

function getRasterCaptureOptions(position, style, { includeTransform = true } = {}) {
    return getPseudoRasterCaptureOptions(position, style, { includeTransform });
}

function mountCloneWithAncestorContext(element, clone, { ignoreAncestorOverflow = false } = {}) {
    return mountPseudoCloneWithAncestorContext(element, clone, { ignoreAncestorOverflow });
}

function mountCloneInIsolatedContext(element, clone) {
    return mountPseudoCloneInIsolatedContext(element, clone);
}

function temporarilySetStyle(element, declarations, restorers) {
    return setPseudoTemporaryStyle(element, declarations, restorers);
}

function isInBrowserComposedCaptureTree(node, element) {
    return isInPseudoCaptureTree(node, element);
}

function temporarilyHideCaptureText(element, restorers) {
    return hidePseudoCaptureText(element, restorers);
}

function temporarilyPinCaptureColor(element, restorers) {
    return pinPseudoCaptureColor(element, restorers);
}

// A clipped ellipse cannot be represented by a native PPTX shape. More
// importantly, an overflow ancestor's layout box is not proof that the
// browser clipped the visual (absolute/fixed/containing-block rules can let
// it paint outside). Capture the real browser composition of this *one*
// element in its original slide context, then crop only to the slide canvas.
// Other page elements are temporarily hidden and are still exported natively.
async function addBrowserComposedClippedVisual(
    pptSlide,
    element,
    root,
    slideRect,
    pxToInH,
    pxToInV,
    { hideText = false, includeShadow = false } = {},
) {
    return captureBrowserComposedVisual({
        pptSlide, element, root, slideRect, pxToInH, pxToInV, hideText, includeShadow,
        operations: {
            getSlideBoundedPosition,
            isInCaptureTree: isInBrowserComposedCaptureTree,
            temporarilySetStyle,
            temporarilyPinCaptureColor,
            temporarilyHideCaptureText,
            renderBackgroundImage,
            addCroppedBrowserCapture,
        },
    });
}

async function addPseudoElementLayer(pptSlide, element, slideRect, pxToInH, pxToInV) {
    const position = getElementPosition(element, slideRect, pxToInH, pxToInV);
    return capturePseudoElementLayer({
        pptSlide, element, position,
        operations: {
            getRasterCaptureOptions,
            inlineRasterizedVisualStyle: inlinePseudoRasterizedVisualStyle,
            materializePseudoElements: materializePseudoElementsForCapture,
            mountCloneWithAncestorContext,
            renderBackgroundImage,
        },
    });
}

async function addComplexVisualLayer(
    pptSlide,
    element,
    slideRect,
    pxToInH,
    pxToInV,
    { omitShadow = false, preserveText = false, atomic = false, browserComposedRoot = null } = {},
) {
    const position = getElementPosition(element, slideRect, pxToInH, pxToInV);
    return captureComplexVisualLayer({
        pptSlide,
        element,
        position,
        options: { omitShadow, preserveText, atomic, browserComposedRoot },
        operations: {
            browserCaptureOperations: {
                getSlideBoundedPosition,
                isInCaptureTree: isInBrowserComposedCaptureTree,
                temporarilySetStyle,
                temporarilyPinCaptureColor,
                temporarilyHideCaptureText,
                renderBackgroundImage,
                addCroppedBrowserCapture,
            },
            captureBrowserComposedVisual,
            context: { slideRect, pxToInH, pxToInV },
            getBorderRadiusCss,
            getRasterCaptureOptions,
            getShapeType,
            getTransformedLocalBounds: calculateTransformedLocalBounds,
            inlineAtomicVisualSubtreeStyles: inlinePseudoAtomicVisualSubtreeStyles,
            inlineRasterizedVisualStyle: inlinePseudoRasterizedVisualStyle,
            materializePseudoElements: materializePseudoElementsForCapture,
            suppressCloneNativePseudoElements,
            mountCloneInIsolatedContext,
            mountCloneWithAncestorContext,
            shapeTypes: SHAPE_TYPES,
        },
    });
}

async function addCssBackgroundLayer(pptSlide, element, slideRect, pxToInH, pxToInV) {
    const position = getElementPosition(element, slideRect, pxToInH, pxToInV);
    return writeCssBackgroundLayer({
        pptSlide, element, position,
        helpers: {
            getBorderRadiusCss, getEffectiveOpacity, getRotateDeg, getShapeType,
            restoreUnrotatedPosition: (target, targetPosition) => restoreUnrotatedPosition(
                target, targetPosition, slideRect, pxToInH, pxToInV,
            ),
            shapeTypes: SHAPE_TYPES,
        },
    });
}

async function addSvg(pptSlide, svg, slideRect, pxToInH, pxToInV) {
    const position = getElementPosition(svg, slideRect, pxToInH, pxToInV);
    const helpers = {
        getElementPosition: target => getElementPosition(target, slideRect, pxToInH, pxToInV),
        getRotateDeg,
        getShapeType,
        shapeTypes: SHAPE_TYPES,
        restoreUnrotatedPosition: (target, targetPosition) => restoreUnrotatedPosition(
            target, targetPosition, slideRect, pxToInH, pxToInV,
        ),
        writeRasterImage,
        hasComplexTransform,
    };
    return writeSvg({ pptSlide, svg, position, helpers });
}

async function addGradientPageBackground(pptSlide, background, slideRect) {
    return writeGradientPageBackground({
        pptSlide, background, slideRect, layout: AI_PPT_EXPORT_LAYOUT,
    });
}

function addChart(pptSlide, chart, slideRect, pxToInH, pxToInV) {
    const position = getElementPosition(chart, slideRect, pxToInH, pxToInV);
    return writeChart(pptSlide, chart, position);
}

function isWithinHandledElement(element, handledElements) {
    let current = element.parentElement;
    while (current) {
        if (handledElements.has(current)) return true;
        current = current.parentElement;
    }
    return false;
}

/**
 * Converts visual DOM primitives into native PptxGenJS objects.  It returns
 * text roots which must not be exported again as text boxes (tables and SVGs).
 */
export async function addNativeVisualElementsToSlide({
    pptSlide,
    root,
    slideRect,
    pxToInH,
    pxToInV,
}) {
    const handledElements = new Set();
    const skippedTextElements = new Set();
    let addedCount = 0;
    let contentVisualCount = 0;
    let rasterizedVisualCount = 0;
    const failures = [];
    const degradations = [];

    const pageBackground = getPageBackground(root);
    if (pageBackground && (pageBackground.hasBackgroundImage || pageBackground.color !== 'FFFFFF')) {
        try {
            const gradientAdded = pageBackground.hasBackgroundImage && await addGradientPageBackground(
                pptSlide,
                pageBackground,
                slideRect,
            );
            if (!gradientAdded) {
                pptSlide.addShape(SHAPE_TYPES.rect, {
                    x: 0,
                    y: 0,
                    w: AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN,
                    h: AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN,
                    fill: {
                        color: pageBackground.color || 'FFFFFF',
                        transparency: colorTransparency(pageBackground.source, pageBackground.opacity),
                    },
                    line: { color: 'FFFFFF', transparency: 100 },
                });
                if (pageBackground.hasBackgroundImage) degradations.push('page-gradient');
            }
            addedCount += 1;
        } catch (error) {
            console.warn('[AiPPT] 页面背景导出失败，已跳过:', error);
        }
    }

    if (addSlideRootFrame(pptSlide, root, slideRect, pxToInH, pxToInV)) {
        addedCount += 1;
    }

    if (hasVisiblePseudoElement(root)) {
        if (await addPseudoElementLayer(pptSlide, root, slideRect, pxToInH, pxToInV)) {
            addedCount += 1;
            contentVisualCount += 1;
            rasterizedVisualCount += 1;
            degradations.push('pseudo-element');
        } else {
            failures.push('pseudo-element');
        }
    }

    const elements = collectElementsInPaintOrder(root);
    const state = {
        handledElements,
        skippedTextElements,
        addedCount,
        contentVisualCount,
        rasterizedVisualCount,
        failures,
        degradations,
    };
    const operations = {
        addBrowserComposedClippedVisual: (element, captureRoot, options) => addBrowserComposedClippedVisual(
            pptSlide, element, captureRoot, slideRect, pxToInH, pxToInV, options,
        ),
        addComplexVisualLayer: (element, options) => addComplexVisualLayer(
            pptSlide, element, slideRect, pxToInH, pxToInV, options,
        ),
        addCssBackgroundLayer: element => addCssBackgroundLayer(pptSlide, element, slideRect, pxToInH, pxToInV),
        addPseudoElementLayer: element => addPseudoElementLayer(pptSlide, element, slideRect, pxToInH, pxToInV),
        addShape: element => addShape(pptSlide, element, slideRect, pxToInH, pxToInV),
        addTable: element => addTable(pptSlide, element, slideRect, pxToInH, pxToInV),
        addImage: (element, options) => addImage(pptSlide, element, slideRect, pxToInH, pxToInV, options),
        addSvg: element => addSvg(pptSlide, element, slideRect, pxToInH, pxToInV),
        addChart: element => addChart(pptSlide, element, slideRect, pxToInH, pxToInV),
    };

    for (const element of elements) {
        if (isWithinHandledElement(element, handledElements)) continue;
        await exportStrategy.exportElement({
            element, root, slideRect, pxToInH, pxToInV, operations, state,
        });
    }

    return {
        addedCount: state.addedCount,
        contentVisualCount: state.contentVisualCount,
        rasterizedVisualCount: state.rasterizedVisualCount,
        failures,
        degradations,
        skippedTextElements,
    };
}
