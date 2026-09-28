// Captures a single browser-composed visual while temporarily hiding unrelated
// slide paint. The mapper supplies shared DOM-isolation helpers until their
// lower-level extraction is complete.
export async function captureBrowserComposedVisual({
    pptSlide, element, root, slideRect, pxToInH, pxToInV,
    hideText = false, includeShadow = false, operations,
}) {
    if (!root || !root.contains(element)) return false;
    const position = operations.getSlideBoundedPosition(element, slideRect, pxToInH, pxToInV, { includeShadow });
    if (!position) return false;
    const restorers = [];
    const allNodes = [root, ...Array.from(root.querySelectorAll('*'))];
    try {
        allNodes.forEach(node => {
            if (!operations.isInCaptureTree(node, element)) {
                operations.temporarilySetStyle(node, { visibility: 'hidden' }, restorers);
            }
        });
        operations.temporarilyPinCaptureColor(element, restorers);
        if (hideText) operations.temporarilyHideCaptureText(element, restorers);
        let ancestor = element.parentElement;
        while (ancestor && ancestor !== root.parentElement) {
            operations.temporarilySetStyle(ancestor, {
                'background-color': 'transparent', 'background-image': 'none',
                'border-color': 'transparent', 'box-shadow': 'none', color: 'transparent',
            }, restorers);
            if (ancestor === root) break;
            ancestor = ancestor.parentElement;
        }
        const fullSlideData = await operations.renderBackgroundImage(root, {
            w: slideRect.width, h: slideRect.height,
        }, { backgroundColor: null, requireVisiblePixels: true });
        await operations.addCroppedBrowserCapture({
            pptSlide, fullSlideData, position, slideRect, document: element.ownerDocument,
        });
        return true;
    } catch (error) {
        console.warn('[AiPPT] 浏览器合成裁剪形状导出失败，回退局部克隆:', error);
        return false;
    } finally {
        restorers.reverse().forEach(restore => restore());
    }
}
