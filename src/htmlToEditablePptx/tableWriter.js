import {
    alignToPptx,
    applyTextTransform,
    extractRunStyle,
    isBoldWeight,
    mapFont,
    normalizeCssColor,
} from './styleExtractor';

function getVisibleCellFill(cell, table, { colorTransparency, getEffectiveOpacity }) {
    let current = cell;
    while (current) {
        const style = current.ownerDocument.defaultView.getComputedStyle(current);
        const backgroundColor = style.backgroundColor;
        const color = normalizeCssColor(backgroundColor);
        if (color) {
            const transparency = colorTransparency(backgroundColor, getEffectiveOpacity(current));
            return { color, ...(transparency > 0 ? { transparency } : {}) };
        }
        if (current === table) break;
        current = current.parentElement;
    }
    // Avoid Office/WPS applying a built-in table style to unfilled cells.
    return { color: 'FFFFFF' };
}

function buildTableRows(table, dependencies) {
    const { getBorderOptions } = dependencies;
    return Array.from(table.rows || []).map(row => Array.from(row.cells || []).map(cell => {
        const style = cell.ownerDocument.defaultView.getComputedStyle(cell);
        const runStyle = extractRunStyle(cell);
        const fill = getVisibleCellFill(cell, table, dependencies);
        const border = getBorderOptions(style);
        return {
            text: applyTextTransform(cell.innerText || cell.textContent || '', runStyle.textTransform),
            options: {
                fontFace: mapFont(runStyle.fontFamily),
                fontSize: runStyle.fontSizePt,
                color: normalizeCssColor(runStyle.color) || '333333',
                bold: isBoldWeight(runStyle.fontWeight),
                italic: /italic/i.test(runStyle.fontStyle || ''),
                ...(runStyle.transparency ? { transparency: runStyle.transparency } : {}),
                align: alignToPptx(style.textAlign),
                valign: 'mid', margin: 0.04, fill,
                ...(border ? { border } : {}),
            },
        };
    })).filter(row => row.length);
}

// The mapper supplies shared visual-style helpers until that layer is moved
// out too. Keeping this dependency explicit prevents duplicate CSS mappings.
export function writeTable(pptSlide, table, position, pxToInH, pxToInV, dependencies) {
    const rows = buildTableRows(table, dependencies);
    if (!position || !rows.length) return false;
    const firstRow = Array.from(table.rows?.[0]?.cells || []);
    const colW = firstRow.length
        ? firstRow.map(cell => pxToInH(cell.getBoundingClientRect().width)) : undefined;
    const rowH = Array.from(table.rows || []).map(row => pxToInV(row.getBoundingClientRect().height));
    try {
        pptSlide.addTable(rows, {
            x: position.x, y: position.y, w: position.w, h: position.h,
            margin: 0.04,
            ...(colW?.length ? { colW } : {}),
            ...(rowH.length ? { rowH } : {}),
            border: { color: 'B7B7B7', pt: 0.75 },
        });
        return true;
    } catch (error) {
        console.warn('[AiPPT] 表格导出失败，已跳过:', error);
        return false;
    }
}
