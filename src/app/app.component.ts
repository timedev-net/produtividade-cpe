import { CommonModule } from '@angular/common';
import { Component, ElementRef, QueryList, ViewChild, ViewChildren, AfterViewInit, OnDestroy } from '@angular/core';
import { FormsModule } from '@angular/forms';
import Chart, { ChartConfiguration, ChartDataset } from 'chart.js/auto';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import * as XLSX from 'xlsx';

type MonthKey = string;

interface ChartDataConfig {
  key: MonthKey;
  monthLabel: string;
  labels: string[];
  datasets: ChartDataset<'line'>[];
}

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [CommonModule, FormsModule],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent implements AfterViewInit, OnDestroy {
  title = 'processos-app';
  paginaAtiva: 'dashboard' | 'analise' = 'dashboard';
  carregado = false;
  fileName = '';
  erroLeitura = '';
  qlikFileName = '';
  eolisFileName = '';
  analiseErro = '';
  resultadoAnaliseHeaders: string[] = [];
  resultadoAnaliseRows: string[][] = [];
  totalProcessos = 0;
  servidores: string[] = [];
  meses: MonthKey[] = [];
  ajustes: Record<string, number> = {};
  diasUteisPorMes: Partial<Record<MonthKey, number>> = {};
  totaisPorMesEServidor: Partial<Record<MonthKey, Record<string, number>>> = {};
  mediasPorMesEServidor: Partial<Record<MonthKey, Record<string, number>>> = {};
  dadosPorDiaEMes: Partial<Record<MonthKey, Record<string, Record<string, number>>>> = {};
  chartConfigs: ChartDataConfig[] = [];

  @ViewChild('reportArea') reportArea?: ElementRef;
  @ViewChildren('monthChart') monthCharts?: QueryList<ElementRef<HTMLCanvasElement>>;

  private charts: Record<MonthKey, Chart> = {};
  private qlikValores = new Set<string>();
  private eolisValores = new Set<string>();
  private eolisLinhas: string[][] = [];
  private eolisHeaders: string[] = [];
  private readonly palette = [
    '#2563eb', '#10b981', '#f97316', '#8b5cf6', '#ef4444',
    '#0ea5e9', '#22c55e', '#f59e0b', '#a855f7', '#64748b',
    '#14b8a6', '#fb7185', '#3b82f6', '#84cc16', '#d946ef'
  ];

  ngAfterViewInit(): void {
    this.monthCharts?.changes.subscribe(() => this.desenharGraficos());
  }

  ngOnDestroy(): void {
    this.destruirGraficos();
  }

  get analisePronta(): boolean {
    return this.qlikValores.size > 0 && this.eolisLinhas.length > 0 && this.eolisValores.size > 0;
  }

  abrirPagina(pagina: 'dashboard' | 'analise'): void {
    this.paginaAtiva = pagina;
  }

  onFileSelected(event: Event): void {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) {
      return;
    }

    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      this.erroLeitura = 'Envie um arquivo .xlsx';
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const arrayBuffer = e.target?.result;
      if (!arrayBuffer) {
        this.erroLeitura = 'Não foi possível ler o arquivo.';
        return;
      }

      const workbook = XLSX.read(arrayBuffer as ArrayBuffer, { type: 'array' });
      this.tratarWorkbook(workbook);
      this.fileName = file.name;
    };
    reader.readAsArrayBuffer(file);
  }

  onQlikSelected(event: Event): void {
    this.processarArquivoAnalise(event, 'qlik');
  }

  onEolisSelected(event: Event): void {
    this.processarArquivoAnalise(event, 'eolis');
  }

  baixarResultadoAnalise(): void {
    if (!this.resultadoAnaliseRows.length) {
      return;
    }

    const dados = [this.resultadoAnaliseHeaders, ...this.resultadoAnaliseRows];
    const sheet = XLSX.utils.aoa_to_sheet(dados);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, 'Resultado');
    XLSX.writeFile(workbook, 'analise-paralisados.xlsx');
  }

  private tratarWorkbook(workbook: XLSX.WorkBook): void {
    this.limparDados();

    if (!workbook.SheetNames.length) {
      this.erroLeitura = 'A planilha está vazia.';
      this.carregado = false;
      return;
    }

    const primeiraAba = workbook.SheetNames[0];
    const sheet = workbook.Sheets[primeiraAba];
    const rows = XLSX.utils.sheet_to_json<Record<string, any>>(sheet, { defval: null });

    const registros: { servidor: string; processo: string; data: Date }[] = [];
    rows.forEach((row) => {
      const servidor = this.buscarCampo(row, 'Servidor');
      const processo = this.buscarCampo(row, 'Processo');
      const data = this.parseDate(this.buscarCampo(row, 'Data'));

      if (servidor && processo && data) {
        registros.push({
          servidor: servidor.toString().trim(),
          processo: processo.toString().trim(),
          data
        });
      }
    });

    if (!registros.length) {
      this.erroLeitura = 'Nenhum registro válido encontrado. Confira os cabeçalhos: Servidor, Processo, Data.';
      this.carregado = false;
      this.limparDados();
      return;
    }

    this.totalProcessos = registros.length;
    this.montarEstruturas(registros);
    this.erroLeitura = '';
    this.carregado = true;
  }

  private processarArquivoAnalise(event: Event, origem: 'qlik' | 'eolis'): void {
    const file = (event.target as HTMLInputElement).files?.[0];
    if (!file) {
      return;
    }

    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      this.analiseErro = 'Envie um arquivo .xlsx';
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      const arrayBuffer = e.target?.result;
      if (!arrayBuffer) {
        this.analiseErro = 'Nao foi possivel ler o arquivo.';
        return;
      }

      const workbook = XLSX.read(arrayBuffer as ArrayBuffer, { type: 'array' });
      const linhas = this.extrairLinhas(workbook);
      const { headers, dataRows } = this.prepararPlanilha(linhas);
      const valores = this.extrairValoresColunaB(dataRows);

      if (!dataRows.length) {
        this.analiseErro = 'Nao ha dados na planilha.';
        if (origem === 'qlik') {
          this.qlikValores = new Set();
          this.qlikFileName = '';
        } else {
          this.eolisValores = new Set();
          this.eolisFileName = '';
          this.eolisLinhas = [];
          this.eolisHeaders = [];
        }
        this.resultadoAnaliseRows = [];
        this.resultadoAnaliseHeaders = [];
        return;
      }

      if (!valores.length) {
        this.analiseErro = 'Nao ha dados na coluna B.';
        if (origem === 'qlik') {
          this.qlikValores = new Set();
          this.qlikFileName = '';
        } else {
          this.eolisValores = new Set();
          this.eolisFileName = '';
          this.eolisLinhas = [];
          this.eolisHeaders = [];
        }
        this.resultadoAnaliseRows = [];
        this.resultadoAnaliseHeaders = [];
        return;
      }

      if (origem === 'qlik') {
        this.qlikValores = new Set(valores);
        this.qlikFileName = file.name;
      } else {
        this.eolisValores = new Set(valores);
        this.eolisFileName = file.name;
        this.eolisLinhas = dataRows;
        this.eolisHeaders = headers;
      }

      this.analiseErro = '';
      this.atualizarAnalise();
    };
    reader.readAsArrayBuffer(file);
  }

  private extrairLinhas(workbook: XLSX.WorkBook): string[][] {
    if (!workbook.SheetNames.length) {
      return [];
    }

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json<any[]>(sheet, { header: 1, defval: '' });

    return rows.map((row) => {
      if (!Array.isArray(row)) {
        return [];
      }
      return row.map((value) => this.normalizarCelula(value));
    });
  }

  private prepararPlanilha(rows: string[][]): { headers: string[]; dataRows: string[][] } {
    if (!rows.length || rows.length < 2) {
      return { headers: [], dataRows: [] };
    }

    const headersRaw = rows[0] ?? [];
    const dataRowsRaw = rows.slice(1);
    const colCount = Math.max(headersRaw.length, ...dataRowsRaw.map((row) => row.length));

    const headers = this.padLinha(headersRaw, colCount).map((value, index) =>
      value.length > 0 ? value : this.nomearColuna(index)
    );
    const dataRows = dataRowsRaw.map((row) => this.padLinha(row, colCount));

    return { headers, dataRows };
  }

  private extrairValoresColunaB(rows: string[][]): string[] {
    return rows
      .map((row) => row[1] ?? '')
      .map((value) => value.trim())
      .filter((value) => value.length > 0);
  }

  private atualizarAnalise(): void {
    if (!this.analisePronta) {
      this.resultadoAnaliseRows = [];
      this.resultadoAnaliseHeaders = [];
      return;
    }

    const vistos = new Set<string>();
    const resultado: string[][] = [];

    this.eolisLinhas.forEach((row) => {
      const valorB = (row[1] ?? '').trim();
      if (!valorB || this.qlikValores.has(valorB) || vistos.has(valorB)) {
        return;
      }
      vistos.add(valorB);
      resultado.push(row);
    });

    this.resultadoAnaliseHeaders = this.eolisHeaders;
    this.resultadoAnaliseRows = resultado;
  }

  private normalizarCelula(value: any): string {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  }

  private padLinha(row: string[], colCount: number): string[] {
    const linha = row.slice(0, colCount);
    while (linha.length < colCount) {
      linha.push('');
    }
    return linha;
  }

  private nomearColuna(index: number): string {
    return `Coluna ${this.indiceParaColuna(index)}`;
  }

  private indiceParaColuna(index: number): string {
    let numero = index + 1;
    let letras = '';
    while (numero > 0) {
      const resto = (numero - 1) % 26;
      letras = String.fromCharCode(65 + resto) + letras;
      numero = Math.floor((numero - 1) / 26);
    }
    return letras;
  }

  private montarEstruturas(registros: { servidor: string; processo: string; data: Date }[]): void {
    const meses = new Set<MonthKey>();
    const servidores = new Set<string>();
    const totais: Record<MonthKey, Record<string, number>> = {};
    const dadosDiaMes: Record<MonthKey, Record<string, Record<string, number>>> = {};
    const vistos = new Set<string>(); // garante 1 processo por servidor/dia/processo
    let contagemUnica = 0;

    registros.forEach(({ servidor, processo, data }) => {
      const dateKey = this.toDateKey(data);
      const mesKey = dateKey.slice(0, 7);
      const chaveUnica = `${servidor}__${mesKey}__${dateKey}__${processo}`;

      if (vistos.has(chaveUnica)) {
        return;
      }
      vistos.add(chaveUnica);
      contagemUnica += 1;

      meses.add(mesKey);
      servidores.add(servidor);

      if (!totais[mesKey]) {
        totais[mesKey] = {};
      }
      totais[mesKey][servidor] = (totais[mesKey][servidor] ?? 0) + 1;

      if (!dadosDiaMes[mesKey]) {
        dadosDiaMes[mesKey] = {};
      }
      if (!dadosDiaMes[mesKey][dateKey]) {
        dadosDiaMes[mesKey][dateKey] = {};
      }
      dadosDiaMes[mesKey][dateKey][servidor] = (dadosDiaMes[mesKey][dateKey][servidor] ?? 0) + 1;
    });

    this.servidores = Array.from(servidores).sort((a, b) => a.localeCompare(b, 'pt-BR'));
    this.meses = Array.from(meses).sort();
    this.totaisPorMesEServidor = totais;
    this.dadosPorDiaEMes = dadosDiaMes;
    this.calcularDiasUteis(meses);
    this.totalProcessos = contagemUnica;

    this.servidores.forEach((servidor) => {
      if (this.ajustes[servidor] === undefined || this.ajustes[servidor] === null) {
        this.ajustes[servidor] = 0;
      }
    });

    this.recalcularMedias();
    this.montarChartConfigs();
  }

  onAjusteChange(): void {
    this.recalcularMedias();
  }

  private recalcularMedias(): void {
    const medias: Record<MonthKey, Record<string, number>> = {};

    this.meses.forEach((mes) => {
      const diasUteis = this.diasUteisPorMes[mes] ?? 0;
      medias[mes] = {};

      this.servidores.forEach((servidor) => {
        const total = this.totaisPorMesEServidor[mes]?.[servidor] ?? 0;
        const ajuste = Math.max(0, Number(this.ajustes[servidor]) || 0);
        const divisor = Math.max(1, diasUteis - ajuste);
        medias[mes][servidor] = Number((total / divisor).toFixed(2));
      });
    });

    this.mediasPorMesEServidor = medias;
  }

  private montarChartConfigs(): void {
    const configs: ChartDataConfig[] = this.meses.map((mesKey) => {
      const diasMap = this.dadosPorDiaEMes[mesKey] ?? {};
      const diaKeys = Object.keys(diasMap).sort();
      const labels = diaKeys.map((dia) => this.formatarDia(dia));

      const datasets: ChartDataset<'line'>[] = this.servidores.map((servidor, idx) => ({
        label: servidor,
        data: diaKeys.map((dia) => diasMap[dia]?.[servidor] ?? 0),
        borderColor: this.palette[idx % this.palette.length],
        backgroundColor: this.hexToRgba(this.palette[idx % this.palette.length], 0.15),
        tension: 0.25,
        fill: true,
        pointRadius: 3,
        pointHoverRadius: 5
      }));

      return {
        key: mesKey,
        monthLabel: this.formatarMes(mesKey),
        labels,
        datasets
      };
    });

    this.chartConfigs = configs;
    setTimeout(() => this.desenharGraficos());
  }

  private desenharGraficos(): void {
    if (!this.monthCharts) {
      return;
    }

    const ativos = new Set(this.chartConfigs.map((c) => c.key));
    Object.keys(this.charts).forEach((mes) => {
      if (!ativos.has(mes)) {
        this.charts[mes].destroy();
        delete this.charts[mes];
      }
    });

    this.monthCharts.forEach((canvasRef) => {
      const mes = canvasRef.nativeElement.dataset['month'] as MonthKey;
      const config = this.chartConfigs.find((c) => c.key === mes);
      if (!config) {
        return;
      }

      if (this.charts[mes]) {
        this.charts[mes].destroy();
      }

      const chartConfig: ChartConfiguration<'line'> = {
        type: 'line',
        data: {
          labels: config.labels,
          datasets: config.datasets
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          interaction: { mode: 'index', intersect: false },
          scales: {
            y: {
              beginAtZero: true,
              ticks: { precision: 0 }
            }
          },
          plugins: {
            legend: { position: 'top' },
            tooltip: { mode: 'index', intersect: false }
          }
        }
      };

      this.charts[mes] = new Chart(canvasRef.nativeElement, chartConfig);
    });
  }

  private destruirGraficos(): void {
    Object.values(this.charts).forEach((chart) => chart.destroy());
    this.charts = {};
  }

  async exportarPdf(): Promise<void> {
    if (!this.reportArea) {
      return;
    }

    const canvas = await html2canvas(this.reportArea.nativeElement, {
      scale: 2,
      backgroundColor: '#f8fafc'
    });

    const imgData = canvas.toDataURL('image/png');
    const pdf = new jsPDF('p', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
    const pageHeight = pdf.internal.pageSize.getHeight();
    let heightLeft = pdfHeight;
    let position = 0;

    pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
    heightLeft -= pageHeight;

    while (heightLeft > 0) {
      position = heightLeft - pdfHeight;
      pdf.addPage();
      pdf.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
      heightLeft -= pageHeight;
    }

    const nome = this.fileName ? this.fileName.replace(/\\.xlsx$/i, '') : 'relatorio-processos';
    pdf.save(`relatorio-${nome}.pdf`);
  }

  formatarMes(mesKey: MonthKey): string {
    const [ano, mes] = mesKey.split('-').map(Number);
    const nomes = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    return `${nomes[mes - 1]}/${ano}`;
  }

  totalDoMes(mesKey: MonthKey): number {
    return Object.values(this.totaisPorMesEServidor[mesKey] ?? {}).reduce((acc, val) => acc + val, 0);
  }

  formatarDia(diaKey: string): string {
    const [ano, mes, dia] = diaKey.split('-');
    return `${dia}/${mes}`;
  }

  private contarDiasUteisNoMes(ano: number, mesIndex: number): number {
    let dias = 0;
    const data = new Date(Date.UTC(ano, mesIndex, 1));
    while (data.getUTCMonth() === mesIndex) {
      const day = data.getUTCDay();
      if (day !== 0 && day !== 6) {
        dias += 1;
      }
      data.setUTCDate(data.getUTCDate() + 1);
    }
    return dias;
  }

  private calcularDiasUteis(meses: Set<MonthKey>): void {
    const dias: Record<MonthKey, number> = {};
    meses.forEach((mesKey) => {
      const [anoStr, mesStr] = mesKey.split('-');
      dias[mesKey] = this.contarDiasUteisNoMes(Number(anoStr), Number(mesStr) - 1);
    });
    this.diasUteisPorMes = dias;
  }

  private toDateKey(date: Date): string {
    const ano = date.getUTCFullYear();
    const mes = String(date.getUTCMonth() + 1).padStart(2, '0');
    const dia = String(date.getUTCDate()).padStart(2, '0');
    return `${ano}-${mes}-${dia}`;
  }

  private buscarCampo(row: Record<string, any>, alvo: string): any {
    const chave = Object.keys(row).find((k) => k.toLowerCase() === alvo.toLowerCase());
    return chave ? row[chave] : null;
  }

  private parseDate(value: any): Date | null {
    if (!value && value !== 0) {
      return null;
    }

    if (value instanceof Date) {
      return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
    }

    if (typeof value === 'number') {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) {
        return new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
      }
    }

    if (typeof value === 'string') {
      const trimmed = value.trim();
      const match = trimmed.match(/^(\d{1,4})[\\/.-](\d{1,2})[\\/.-](\d{1,4})$/);

      if (match) {
        const [_, p1, p2, p3] = match;
        let ano: number;
        let mes: number;
        let dia: number;

        if (p1.length === 4) {
          ano = Number(p1);
          mes = Number(p2);
          dia = Number(p3);
        } else {
          dia = Number(p1);
          mes = Number(p2);
          ano = p3.length === 2 ? Number(`20${p3}`) : Number(p3);
        }

        return new Date(Date.UTC(ano, mes - 1, dia));
      }

      const parsed = new Date(trimmed);
      if (!isNaN(parsed.getTime())) {
        return new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()));
      }
    }

    return null;
  }

  private hexToRgba(hex: string, alpha: number): string {
    const cleanHex = hex.replace('#', '');
    const bigint = parseInt(cleanHex, 16);
    const r = (bigint >> 16) & 255;
    const g = (bigint >> 8) & 255;
    const b = bigint & 255;
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
  }

  private limparDados(): void {
    this.servidores = [];
    this.meses = [];
    this.totaisPorMesEServidor = {};
    this.mediasPorMesEServidor = {};
    this.dadosPorDiaEMes = {};
    this.chartConfigs = [];
    this.diasUteisPorMes = {};
    this.totalProcessos = 0;
    this.carregado = false;
    this.destruirGraficos();
  }
}
