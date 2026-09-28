import { normalizeCssColor } from './styleExtractor';

export function parseChartData(element) {
    const rawData = element.getAttribute('data-pptx-chart-data');
    if (!rawData) return null;
    try {
        const parsed = JSON.parse(rawData);
        const labels = Array.isArray(parsed.labels) ? parsed.labels.map(value => String(value)) : [];
        const data = (Array.isArray(parsed.series) ? parsed.series : []).map((item, index) => ({
            name: String(item?.name || `系列 ${index + 1}`),
            labels,
            values: Array.isArray(item?.values)
                ? item.values.map(value => Number(value)).filter(value => Number.isFinite(value)) : [],
        })).filter(item => item.values.length && item.values.length === labels.length);
        if (!labels.length || !data.length) return null;
        return {
            data,
            colors: Array.isArray(parsed.colors) ? parsed.colors.map(normalizeCssColor).filter(Boolean) : [],
            showLegend: parsed.showLegend === true || parsed.showLegend === 'true',
            showValue: parsed.showValue === true || parsed.showValue === 'true',
        };
    } catch (error) {
        console.warn('[AiPPT] 图表数据格式无效，已跳过:', error);
        return null;
    }
}
