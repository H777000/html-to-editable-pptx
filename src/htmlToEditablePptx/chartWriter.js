import { parseChartData } from './chartData';

const CHART_TYPES = new Set(['area', 'bar', 'doughnut', 'line', 'pie', 'radar']);

export function writeChart(pptSlide, chart, position) {
    const type = chart.getAttribute('data-pptx-chart-type') || 'bar';
    if (!CHART_TYPES.has(type)) return false;
    const chartData = parseChartData(chart);
    if (!position || !chartData) return false;
    try {
        pptSlide.addChart(type, chartData.data, {
            x: position.x, y: position.y, w: position.w, h: position.h,
            showLegend: chartData.showLegend, showValue: chartData.showValue,
            ...(chartData.colors.length ? { chartColors: chartData.colors } : {}),
            showTitle: false, showCatName: false,
            catAxisLabelFontFace: 'Microsoft YaHei', valAxisLabelFontFace: 'Microsoft YaHei',
            chartArea: { border: { color: 'FFFFFF', transparency: 100 } },
            plotArea: { border: { color: 'FFFFFF', transparency: 100 } },
        });
        return true;
    } catch (error) {
        console.warn('[AiPPT] 图表导出失败，已跳过:', error);
        return false;
    }
}
