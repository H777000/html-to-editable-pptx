import { buildHtmlSlidesToEditablePptx } from './htmlToEditablePptx';

window.HtmlToEditablePptx = {
    async convert(html, fileName = 'slides.pptx') {
        return buildHtmlSlidesToEditablePptx({
            htmlSlides: [{ page: 1, html }],
            fileName,
        });
    },
};
