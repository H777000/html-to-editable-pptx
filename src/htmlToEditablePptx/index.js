import pptxgen from 'pptxgenjs';
import { AI_PPT_EXPORT_LAYOUT } from '../../constants';
import {
    addHtmlSlideToPptx,
    addHtmlSlideToPptxWithEditableTextOverlay,
} from './slideBuilder';
import { downloadPptxBlob, repairPptxPackage } from './pptxPostProcessor';

const EXPORT_TOTAL_TIMEOUT_MS = 120000;

function throwIfExportStopped(signal, deadlineAt) {
    if (signal?.aborted) {
        const error = new Error('PPT 导出任务已取消');
        error.name = 'AbortError';
        throw error;
    }
    if (Date.now() >= deadlineAt) throw new Error('PPT 导出准备总超时');
}

export async function buildHtmlSlidesToEditablePptx({
    htmlSlides,
    fileName,
    visualFidelityPages = [],
    signal,
}) {
    const slides = (htmlSlides || [])
        .filter(item => item && item.html)
        .sort((left, right) => Number(left.page) - Number(right.page));

    if (!slides.length) {
        throw new Error('当前没有可导出的 HTML PPT 页面。');
    }

    const pptx = new pptxgen();
    pptx.defineLayout({
        name: 'AI_PPT_HTML_HYBRID_16_9',
        width: AI_PPT_EXPORT_LAYOUT.PPT_WIDTH_IN,
        height: AI_PPT_EXPORT_LAYOUT.PPT_HEIGHT_IN,
    });
    pptx.layout = 'AI_PPT_HTML_HYBRID_16_9';
    pptx.author = 'HTML to Editable PPTX';
    pptx.subject = 'HTML slide conversion';
    pptx.title = fileName || 'slides.pptx';
    pptx.company = '';

    const visualFidelityPageSet = new Set(
        (visualFidelityPages || [])
            .map(page => Number(page))
            .filter(page => Number.isInteger(page) && page > 0),
    );
    const slideReports = [];
    const deadlineAt = Date.now() + EXPORT_TOTAL_TIMEOUT_MS;
    for (const item of slides) {
        throwIfExportStopped(signal, deadlineAt);
        const slide = pptx.addSlide();
        // 指定页直接采用浏览器视觉底图，避免 Office 重新解释圆/椭圆时发生裁剪。
        // 页面中的文字由导出器再叠加为原生 PPT 文本，仍可编辑。
        const report = visualFidelityPageSet.has(Number(item.page))
            ? await addHtmlSlideToPptxWithEditableTextOverlay(slide, item, { signal, deadlineAt })
            : await addHtmlSlideToPptx(slide, item, { signal, deadlineAt });
        if (report) slideReports.push(report);
    }

    throwIfExportStopped(signal, deadlineAt);
    const finalFileName = fileName || `slides_${Date.now()}.pptx`;
    const rawPptxBlob = await pptx.write({ outputType: 'blob', compression: true });
    throwIfExportStopped(signal, deadlineAt);
    const repairedPptxBlob = await repairPptxPackage(rawPptxBlob);
    throwIfExportStopped(signal, deadlineAt);
    return {
        blob: repairedPptxBlob,
        fileName: finalFileName,
        slideReports,
        visualFallbackPages: slideReports
            .filter(report => report.visualFallback)
            .map(report => report.page),
        visualTextOverlayPages: slideReports
            .filter(report => report.visualTextOverlay)
            .map(report => report.page),
        degradedPages: slideReports
            .filter(report => !report.visualFallback && report.fallbackReasons?.length)
            .map(report => report.page),
    };
}

// 保留原调用方式，供旧调用方与测试继续使用。新流程会先调用
// buildHtmlSlidesToEditablePptx，再由用户真实点击原生下载链接。
export async function exportHtmlSlidesToEditablePptx(options) {
    const result = await buildHtmlSlidesToEditablePptx(options);
    downloadPptxBlob(result.blob, result.fileName);
    return result;
}
