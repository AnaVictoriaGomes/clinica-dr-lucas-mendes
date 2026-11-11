import * as XLSX from 'xlsx';
import { exportDataFile, renderExportModal } from './export';
import {
  renderGraficosTabContent,
  startEditingChart,
  renderStudioSvg,
  computeStudioChartData,
  getStandardChartDefs
} from './chartBuilder';

declare global {
  interface Window {
    XLSX: typeof XLSX;
    KommoExport: {
      exportDataFile: typeof exportDataFile;
      renderExportModal: typeof renderExportModal;
    };
    KommoGraf: {
      render: typeof renderGraficosTabContent;
      editChart: typeof startEditingChart;
      renderVisualChart: typeof renderStudioSvg;
      computeChartData: typeof computeStudioChartData;
      getStandardChartDefs: typeof getStandardChartDefs;
    };
  }
}

window.XLSX = XLSX;
window.KommoExport = {
  exportDataFile,
  renderExportModal
};
window.KommoGraf = {
  render: renderGraficosTabContent,
  editChart: startEditingChart,
  renderVisualChart: renderStudioSvg,
  computeChartData: computeStudioChartData,
  getStandardChartDefs
};

const grafEl = document.getElementById('graf');
if (grafEl && !grafEl.classList.contains('hide')) {
  try {
    renderGraficosTabContent();
  } catch (err) {
    console.error('Erro ao renderizar aba de gráficos:', err);
  }
}
