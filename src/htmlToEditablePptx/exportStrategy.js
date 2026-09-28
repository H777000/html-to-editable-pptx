// Strategy predicates are intentionally dependency-injected: geometry and
// visual-style mappings remain single-source modules instead of being copied.
export function createExportStrategy(helpers) {
    const {
        getElementPosition, getShapeType, getRotateDeg, hasComplexCssEffect,
        hasComplexTransform, hasCssBackgroundImage, hasAsymmetricCornerRadius,
        hasAsymmetricVisibleBorder, hasVisiblePseudoElement, isGradientText,
        isShapeCandidate, isTransparentColor, getTagName, shapeTypes,
    } = helpers;

    function hasMeaningfulElementText(element) {
        return String(element?.textContent || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().length > 0;
    }

    function hasPillGeometry(style, position) {
        const radii = helpers.getCornerRadiusValues(style);
        if (radii.length !== 4 || radii.some(value => !value)) return false;
        const halfMinorAxis = Math.min(position.rect.width, position.rect.height) / 2;
        return radii.every(value => {
            const normalized = String(value).trim();
            return normalized.endsWith('%') ? (Number.parseFloat(normalized) || 0) >= 45
                : (Number.parseFloat(normalized) || 0) >= (halfMinorAxis - 1);
        });
    }

    function hasNativeLineArrow(element) {
        if (element.getAttribute('data-pptx-role') !== 'line') return false;
        const nativeArrowTypes = new Set(['arrow', 'diamond', 'oval', 'stealth', 'triangle']);
        return ['data-pptx-arrow-begin', 'data-pptx-arrow-end']
            .some(attribute => nativeArrowTypes.has(element.getAttribute(attribute)));
    }

    return {
        hasMeaningfulElementText,
        needsBrowserComposedDecorationCapture(element, style, slideRect, pxToInH, pxToInV) {
            if (!isShapeCandidate(element, style)) return false;
            if (element.getAttribute('data-pptx-shape') || element.getAttribute('data-pptx-role')) return false;
            const position = getElementPosition(element, slideRect, pxToInH, pxToInV);
            if (!position) return false;
            return getShapeType(element, style, position) === shapeTypes.ellipse || hasPillGeometry(style, position);
        },
        needsClippedShapeRasterization(element, style, slideRect, pxToInH, pxToInV) {
            if (!isShapeCandidate(element, style)) return false;
            const position = getElementPosition(element, slideRect, pxToInH, pxToInV);
            if (!position?.clipped) return false;
            const type = getShapeType(element, style, position);
            return type === shapeTypes.ellipse || type === shapeTypes.roundRect || Boolean(getRotateDeg(style));
        },
        shouldRasterizeGradientElement(element, style) {
            return hasCssBackgroundImage(style) && !isGradientText(style) && isShapeCandidate(element, style);
        },
        isCompactDecoratedText(element, style, slideRect, pxToInH, pxToInV) {
            const text = String(element.textContent || '').replace(/\s+/g, ' ').trim();
            if (!text || text.length > 48) return false;
            const position = getElementPosition(element, slideRect, pxToInH, pxToInV);
            if (!position || position.rect.width > 280 || position.rect.height > 64) return false;
            const inline = ['inline', 'inline-block', 'inline-flex', 'inline-grid'].includes(String(style.display || '').toLowerCase());
            const visual = !isTransparentColor(style.backgroundColor) || hasCssBackgroundImage(style) || hasComplexCssEffect(style);
            const rounded = (Number.parseFloat(style.borderTopLeftRadius) || 0) > 0 || /%/.test(String(style.borderRadius || ''));
            return visual && (inline || hasComplexCssEffect(style) || rounded);
        },
        isPillVisualAtom(element, style, slideRect, pxToInH, pxToInV) {
            const text = String(element.textContent || '').replace(/\s+/g, ' ').trim();
            if (!text || text.length > 64) return false;
            const position = getElementPosition(element, slideRect, pxToInH, pxToInV);
            if (!position || position.rect.height > 96 || position.rect.width > 960 || position.rect.width / Math.max(1, position.rect.height) < 1.6) return false;
            const visual = !isTransparentColor(style.backgroundColor) || hasCssBackgroundImage(style) || hasComplexCssEffect(style) ||
                ['Top', 'Right', 'Bottom', 'Left'].some(side => Number.parseFloat(style[`border${side}Width`]) > 0);
            return visual && hasPillGeometry(style, position);
        },
        isSmallCircularVisualBadge(element, style, slideRect, pxToInH, pxToInV) {
            if (helpers.getTagName(element) === 'SVG' || String(element.textContent || '').trim()) return false;
            const position = getElementPosition(element, slideRect, pxToInH, pxToInV);
            if (!position || Math.max(position.rect.width, position.rect.height) > 160) return false;
            if (getShapeType(element, style, position) !== shapeTypes.ellipse) return false;
            if (helpers.hasVisiblePseudoElement(element)) return true;
            return Array.from(element.querySelectorAll('*')).some(child => {
                if (['SVG', 'IMG'].includes(helpers.getTagName(child))) return true;
                const childPosition = getElementPosition(child, slideRect, pxToInH, pxToInV);
                if (!childPosition || Math.max(childPosition.rect.width, childPosition.rect.height) > 96) return false;
                const childStyle = child.ownerDocument.defaultView.getComputedStyle(child);
                return helpers.hasVisiblePseudoElement(child) || hasCssBackgroundImage(childStyle) ||
                    hasComplexCssEffect(childStyle) || !isTransparentColor(childStyle.backgroundColor);
            });
        },
        async exportElement({ element, root, slideRect, pxToInH, pxToInV, operations, state }) {
            const tagName = getTagName(element);
            const style = element.ownerDocument.defaultView.getComputedStyle(element);
            const hasPseudoElement = hasVisiblePseudoElement(element);
            const {
                addBrowserComposedClippedVisual, addComplexVisualLayer, addCssBackgroundLayer,
                addPseudoElementLayer, addShape, addTable, addImage, addSvg, addChart,
            } = operations;
            const addSuccess = ({ handled = false, skipText = false, rasterized = false, degradation } = {}) => {
                if (handled) state.handledElements.add(element);
                if (skipText) state.skippedTextElements.add(element);
                state.addedCount += 1;
                state.contentVisualCount += 1;
                if (rasterized) state.rasterizedVisualCount += 1;
                if (degradation && !state.degradations.includes(degradation)) state.degradations.push(degradation);
            };
            const addFailure = failure => state.failures.push(failure);

            if (this.needsBrowserComposedDecorationCapture(element, style, slideRect, pxToInH, pxToInV)) {
                if (await addBrowserComposedClippedVisual(element, root, { hideText: this.hasMeaningfulElementText(element) })) {
                    addSuccess({ handled: true, rasterized: true, degradation: 'browser-composed-decoration' });
                } else addFailure('browser-composed-decoration');
                return;
            }

            if (this.isSmallCircularVisualBadge(element, style, slideRect, pxToInH, pxToInV)) {
                const browserComposedAdded = await addBrowserComposedClippedVisual(element, root, { includeShadow: true });
                const added = browserComposedAdded || await addComplexVisualLayer(element, { omitShadow: true, atomic: true });
                if (added) addSuccess({ handled: true, skipText: true, rasterized: true, degradation: 'icon-badge' });
                else addFailure('icon-badge');
                return;
            }

            if (this.isPillVisualAtom(element, style, slideRect, pxToInH, pxToInV)) {
                const browserComposedAdded = await addBrowserComposedClippedVisual(element, root, { hideText: true, includeShadow: true });
                const added = browserComposedAdded || await addComplexVisualLayer(element, { preserveText: true, atomic: true });
                if (added) addSuccess({ handled: true, skipText: !browserComposedAdded, rasterized: true, degradation: 'pill-atom' });
                else addFailure('pill-atom');
                return;
            }

            if (this.isCompactDecoratedText(element, style, slideRect, pxToInH, pxToInV)) {
                if (await addComplexVisualLayer(element, { preserveText: true })) {
                    addSuccess({ handled: true, skipText: true, rasterized: true, degradation: 'compact-decorated-text' });
                } else addFailure('compact-decorated-text');
                return;
            }

            if (hasComplexTransform(style) || hasComplexCssEffect(style)) {
                const degradation = hasComplexTransform(style) ? 'complex-transform' : 'complex-css';
                const keepsEditableText = !['IMG', 'SVG', 'TABLE'].includes(tagName);
                const position = keepsEditableText && getElementPosition(element, slideRect, pxToInH, pxToInV);
                const omitShadow = position && getShapeType(element, style, position) === shapeTypes.ellipse;
                const added = keepsEditableText
                    ? await addComplexVisualLayer(element, { omitShadow })
                    : await addImage(element, { preserveTransform: true });
                if (added) addSuccess({ handled: true, skipText: !keepsEditableText, rasterized: true, degradation });
                else addFailure(degradation);
                return;
            }

            if (tagName === 'TABLE') {
                if (addTable(element)) addSuccess({ handled: true, skipText: true });
                else addFailure('table');
                return;
            }

            if (tagName === 'IMG') {
                if (await addImage(element)) addSuccess({ handled: true, rasterized: true });
                else addFailure('image');
                return;
            }

            if (tagName === 'SVG') {
                if (await addSvg(element)) {
                    addSuccess({
                        handled: true,
                        skipText: true,
                        rasterized: Boolean(element.querySelector('foreignObject, image, filter, mask, pattern, animate, animateTransform, set, use')),
                    });
                } else addFailure('svg');
                return;
            }

            if (element.getAttribute('data-pptx-role') === 'chart') {
                if (addChart(element)) addSuccess({ handled: true, skipText: true });
                else if (await addImage(element)) addSuccess({ handled: true, skipText: true, rasterized: true });
                else addFailure('chart');
                return;
            }

            if (this.needsClippedShapeRasterization(element, style, slideRect, pxToInH, pxToInV)) {
                const position = getElementPosition(element, slideRect, pxToInH, pxToInV);
                const type = position && getShapeType(element, style, position);
                if (await addComplexVisualLayer(element, { omitShadow: type === shapeTypes.ellipse, browserComposedRoot: root })) {
                    addSuccess({ handled: true, rasterized: true, degradation: 'clipped-shape' });
                } else addFailure('clipped-shape');
                return;
            }

            if (isShapeCandidate(element, style) && hasAsymmetricVisibleBorder(style)) {
                if (await addComplexVisualLayer(element)) addSuccess({ handled: true, rasterized: true, degradation: 'asymmetric-border' });
                else addFailure('asymmetric-border');
                return;
            }

            if (isShapeCandidate(element, style) && hasAsymmetricCornerRadius(style)) {
                if (await addComplexVisualLayer(element)) addSuccess({ handled: true, rasterized: true, degradation: 'asymmetric-radius' });
                else addFailure('asymmetric-radius');
                return;
            }

            if (this.shouldRasterizeGradientElement(element, style)) {
                if (await addCssBackgroundLayer(element)) {
                    addSuccess({ rasterized: true });
                    if (hasPseudoElement) {
                        if (await addPseudoElementLayer(element)) addSuccess({ rasterized: true, degradation: 'pseudo-element' });
                        else addFailure('pseudo-element');
                    }
                } else addFailure('gradient');
                return;
            }

            const shapeAdded = addShape(element);
            if (shapeAdded) addSuccess();
            // AI line markup explicitly asks for a native PPT arrowhead. Its CSS
            // pseudo element is the browser preview counterpart, not a second
            // decoration to export. Keep the pseudo path only as a fallback if
            // native shape export fails.
            if (hasPseudoElement && !(shapeAdded && hasNativeLineArrow(element))) {
                if (await addPseudoElementLayer(element)) addSuccess({ rasterized: true, degradation: 'pseudo-element' });
                else addFailure('pseudo-element');
            }
        },
    };
}
