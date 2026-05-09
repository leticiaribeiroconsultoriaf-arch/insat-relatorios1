/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef, useEffect } from 'react';
import { 
  FileText, 
  FileSpreadsheet, 
  CheckCircle2, 
  AlertCircle, 
  Loader2, 
  Terminal,
  ArrowRight,
  Info
} from 'lucide-react';
import * as XLSX from 'xlsx';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import ImageModule from 'docxtemplater-image-module-free';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { Chart, registerables } from 'chart.js';
import ChartDataLabels from 'chartjs-plugin-datalabels';
import { motion, AnimatePresence } from 'motion/react';

Chart.register(...registerables, ChartDataLabels);

// --- Types ---
interface LogEntry {
  id: string;
  message: string;
  timestamp: string;
}

interface BufferedFile {
  name: string;
  data: ArrayBuffer;
}

// --- Dynamic Table Component ---
const CriticosTable = () => {
  const factors = [
    "Ambiente físico de trabalho e equipamentos inadequados",
    "Carga e/ou ritmo de trabalho excessivo",
    "Jornada de trabalho excessiva",
    "Contacto com o sofrimento humano",
    "Contato direto com pessoas agressivas",
    "Ausência de portal/site específico para denúncia de assédio",
    "Ausência de autonomia e controle de decisões",
    "Relacionamento interpessoal conflitante e/ou Isolamento",
    "Desequilíbrio entre vida profissional e familiar"
  ];

  const dados = [
    [null, null, null],  // Fator 1
    [0,    0,    0   ],  // Fator 2
    [null, null, null],  // Fator 3
    [0,    null, null],  // Fator 4
    [null, 0,    null],  // Fator 5
    [null, null, null],  // Fator 6
    [null, null, 1   ],  // Fator 7
    [null, null, 0   ],  // Fator 8
    [null, null, null]   // Fator 9
  ];

  const thresholds = [2, 2, 3];
  const sums = [0, 0, 0];
  
  // Calculate sums
  dados.forEach(row => {
    row.forEach((val, i) => {
      if (val === 1) sums[i]++;
    });
  });

  const headers = [
    { t: "CRÍTICA 1", s: "(Esgotamento profissional;\ncontato com o sofrimento)" },
    { t: "CRÍTICA 2", s: "(Burn-out;\ncontacto com usuários)" },
    { t: "CRÍTICA 3", s: "(Controle das demandas\ndo trabalho)" }
  ];

  return (
    <div className="w-full my-8 bg-white overflow-x-auto">
      <table className="w-full border-collapse border-2 border-[#333] table-fixed font-sans text-[11px]">
        <thead>
          <tr>
            <th className="border border-[#333] bg-white w-[55%] p-2.5 text-center align-middle">
              <span className="text-[#CC0000] font-bold text-[13px]">CENÁRIOS CRÍTICOS</span>
              <span className="text-[#CC0000] font-normal text-[13px] ml-1">por combinações predefinidas de fatores de exposição.</span>
            </th>
            {headers.map((h, i) => (
              <th key={i} className="border border-[#333] bg-[#FFFF00] w-[15%] text-center p-1.5 align-middle">
                <div className="flex flex-col items-center leading-tight">
                  <span className="font-bold text-[10px] text-black uppercase">{h.t}</span>
                  <span className="font-normal text-[10px] text-black whitespace-pre-line">{h.s}</span>
                </div>
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {factors.map((f, fIdx) => (
            <tr key={fIdx} className="h-[36px]">
              <td className="border border-[#333] bg-[#E8F0F8] p-2 text-left align-middle font-bold text-black break-words leading-tight">
                {f}
              </td>
              {[0, 1, 2].map(cIdx => {
                const valueToShow = dados[fIdx][cIdx];
                return (
                  <td 
                    key={cIdx} 
                    style={{
                      borderTop: valueToShow === null ? 'none' : '1px solid #333',
                      borderBottom: valueToShow === null ? 'none' : '1px solid #333',
                      borderLeft: '1px solid #333',
                      borderRight: '1px solid #333'
                    }}
                    className={`text-center align-middle font-bold text-[18px] transition-colors ${
                      valueToShow === 1 ? 'bg-[#9966FF]' : 'bg-[#E8F0F8]'
                    }`}
                  >
                    {valueToShow !== null ? valueToShow : ""}
                  </td>
                );
              })}
            </tr>
          ))}
          <tr>
            <td className="border border-[#333] bg-[#D3D3D3] p-2 align-middle">
              <div className="flex flex-col items-center">
                <span className="text-[#CC0000] font-bold text-[12px] w-full text-center">SITUAÇÃO CRÍTICA PRESENTE</span>
                <span className="text-black text-[11px] text-center mt-1 leading-tight max-w-[420px]">
                  A situação crítica específica é ativada quando todas as caixas acima mencionadas ficam roxas, para esse cenário crítico específico. Quando a situação crítica está presente, as caixas do lado direito também ficam roxas.
                </span>
              </div>
            </td>
            {sums.map((sum, i) => (
              <td 
                key={i} 
                className={`border border-[#333] text-center align-middle font-bold text-[20px] transition-colors ${
                  sum >= thresholds[i] ? 'bg-[#9966FF]' : 'bg-white'
                }`}
              >
                {sum}
              </td>
            ))}
          </tr>
        </tbody>
      </table>
    </div>
  );
};

// --- Risk Matrix Functions ---
function classificarDano(score: number) {
  return {
    AA248: (score <= 0.20) ? "X" : "",      // DESPREZÍVEL
    AA249: (score > 0.20 && score <= 0.52) ? "X" : "",  // BAIXO
    AA250: (score > 0.52 && score <= 0.80) ? "X" : "",  // MODERADO
    AA251: (score > 0.80) ? "X" : "",      // ALTO
    label: (score <= 0.20) ? "DESPREZÍVEL" : (score <= 0.52) ? "BAIXO" : (score <= 0.80) ? "MODERADO" : "ALTO"
  };
}

function classificarExposicao(score: number) {
  return {
    AA234: (score === 0) ? "X" : "",                         // AUSENTE
    AA239: (score > 0 && score <= 0.23) ? "X" : "",          // DESPREZÍVEL
    AA235: (score >= 0.231 && score <= 0.475) ? "X" : "",    // ACEITÁVEL
    AA236: (score >= 0.476 && score <= 0.735) ? "X" : "",    // MODERADO
    AA237: (score >= 0.736 && score <= 0.95) ? "X" : "",     // ALTO
    AA238: (score > 0.95) ? "X" : "",                        // CRÍTICO
    label: (score === 0) ? "AUSENTE" : (score <= 0.23) ? "DESPREZÍVEL" : (score <= 0.475) ? "ACEITÁVEL" : (score <= 0.735) ? "MODERADO" : (score <= 0.95) ? "ALTO" : "CRÍTICO"
  };
}

const RiskMatrix = ({ danoScore, exposicaoScore }: { danoScore: number, exposicaoScore: number }) => {
  const danoClass = classificarDano(danoScore);
  const expClass = classificarExposicao(exposicaoScore);

  const nivelDano = danoClass.label;
  const nivelExposicao = expClass.label;

  // EIXO Y — indicadores de linha
  const F322 = danoClass.AA251;  // ALTO
  const F323 = danoClass.AA250;  // MODERADO
  const F324 = danoClass.AA249;  // BAIXO
  const F325 = "";               // DESPREZÍVEL — sempre vazio

  // EIXO X — indicadores de coluna
  const H327 = expClass.AA239;   // DESPREZÍVEL
  const I327 = expClass.AA235;   // ACEITÁVEL
  const J327 = expClass.AA236;   // MODERADO
  const K327 = expClass.AA237;   // ALTO
  const L327 = expClass.AA238;   // CRÍTICO

  const rowLabels = ["ALTO", "MODERADO", "BAIXO", "DESPREZÍVEL"] as const;
  const colLabels = ["DESPREZÍVEL", "ACEITÁVEL", "MODERADO", "ALTO", "CRÍTICO"] as const;

  const rowIndicators = [F322, F323, F324, F325];
  const colIndicators = [H327, I327, J327, K327, L327];

  const staticColors = [
    ['#FFFF99', '#FFC000', '#FF0000', '#FF0000', '#C4A7FF'], // ALTO
    ['#FFFF97', '#FFFF97', '#FFD347', '#FF0000', '#C4A7FF'], // MODERADO
    ['#C9E7A7', '#FFFF97', '#FFD347', '#FFD347', '#FF0000'], // BAIXO
    ['#C9E7A7', '#C9E7A7', '#FFFF97', '#FFD347', '#FFC000']  // DESPREZÍVEL
  ];

  const getCellColor = (rIdx: number, cIdx: number, isActive: boolean) => {
    if (!isActive) return staticColors[rIdx][cIdx];
    const row = rowLabels[rIdx];
    const col = colLabels[cIdx];
    if (row === "ALTO" && col === "DESPREZÍVEL") return "#FF0000";
    if (row === "DESPREZÍVEL" && col === "CRÍTICO") return "#FF0000";
    return staticColors[rIdx][cIdx];
  };

  const getActiveColor = () => {
    if (nivelExposicao === 'AUSENTE' || nivelExposicao === "" || nivelDano === "") return '#FFFFFF';
    const rIdx = rowLabels.indexOf(nivelDano as any);
    const cIdx = colLabels.indexOf(nivelExposicao as any);
    if (rIdx === -1 || cIdx === -1) return '#FFFFFF';
    return getCellColor(rIdx, cIdx, true);
  };

  const activeColor = getActiveColor();
  
  const getRiskLabel = (color: string) => {
    if (nivelExposicao === 'AUSENTE') return 'NÃO IDENTIFICADO';
    if (color === '#C9E7A7' || color === '#B4E391') return 'BAIXO';
    if (color === '#FFFF97' || color === '#FFFF99') return 'ACEITÁVEL';
    if (color === '#FFD347') return 'MODERADO';
    if (color === '#FFC000') return 'ELEVADO';
    if (color === '#FF0000') return 'ALTO / CRÍTICO';
    if (color === '#C4A7FF' || color === '#9D90C9') return 'EXTREMO';
    return 'N/A';
  };

  return (
    <div className="w-fit h-fit mx-auto my-0 bg-white overflow-hidden border border-[#333]">
      {/* Title */}
      <div className="bg-[#FFFF00] p-3.5 px-4 text-center border-b border-[#333]">
        <h2 className="text-black font-bold text-[15px] leading-relaxed break-words normal-case">
          RISCO DE TRANSTORNOS MENTAIS (stress/depressão) RELACIONADOS AO TRABALHO
        </h2>
      </div>
      
      {/* Subtitle */}
      <div className="bg-[#D6E3BC] p-2 px-3 text-center border-b border-[#333]">
        <h3 className="text-black text-[12px] font-bold uppercase">
          MATRIZ DE RISCO: EXPOSIÇÃO (fatores de conteúdo e contexto: 'Probabilidade') x EVENTOS SENTINELA ('Danos')
        </h3>
      </div>

      <div className="p-0 flex justify-center">
        <table className="border-collapse border-none">
          <tbody>
            {rowLabels.map((rowLabel, rIdx) => (
              <tr key={rowLabel}>
                <td className="w-10 min-w-[40px] h-9 bg-[#EAF1DD] border border-[#333] text-black font-bold text-[14px] text-center">
                  {rowIndicators[rIdx]}
                </td>
                <td className="w-[90px] h-9 bg-[#3F3F3F] text-white font-bold text-[11px] border border-[#333]">
                  {rowLabel}
                </td>
                {colLabels.map((colLabel, cIdx) => {
                  const isActive = rowIndicators[rIdx] === "X" && colIndicators[cIdx] === "X";
                  return (
                    <td 
                      key={colLabel}
                      style={{ backgroundColor: getCellColor(rIdx, cIdx, isActive) }}
                      className="w-20 h-9 border border-[#333] text-black font-bold text-[14px] text-center"
                    >
                      {isActive ? "X" : ""}
                    </td>
                  );
                })}
              </tr>
            ))}
            <tr>
              <td className="bg-transparent border-none w-10 min-w-[40px]"></td>
              <td className="bg-transparent border-none"></td>
              {colLabels.map((colLabel, cIdx) => (
                <td key={colLabel} className="h-5 bg-[#D6E3BC] border border-[#333] text-black font-bold text-[14px]">
                  {colIndicators[cIdx]}
                </td>
              ))}
              <td className="bg-[#D6E3BC] border border-[#333] text-black text-[10px] p-1 font-bold">
                Exposição
              </td>
            </tr>
            <tr>
              <td className="bg-transparent border-none w-10 min-w-[40px]"></td>
              <td className="bg-transparent border-none"></td>
              {colLabels.map((colLabel) => (
                <td key={colLabel} className="h-9 bg-[#3F3F3F] text-white font-bold text-[10px] border border-[#333] uppercase px-1">
                  {colLabel}
                </td>
              ))}
              <td className="bg-transparent border-none"></td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
};

// --- Canvas Helpers ---
const wrapText = (ctx: CanvasRenderingContext2D, text: string, maxWidth: number, fontSize: number, bold = false) => {
  ctx.font = `${bold ? "900" : "400"} ${fontSize}px Arial`;
  const words = text.split(' ');
  const lines = [];
  let currentLine = words[0] || "";
  for (let j = 1; j < words.length; j++) {
      if (ctx.measureText(currentLine + " " + words[j]).width < maxWidth) {
          currentLine += " " + words[j];
      } else {
          lines.push(currentLine);
          currentLine = words[j];
      }
  }
  lines.push(currentLine);
  return lines;
};

export default function App() {
  const [molde, setMolde] = useState<BufferedFile | null>(null);
  const [excels, setExcels] = useState<BufferedFile[]>([]);
  const [includePlano, setIncludePlano] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [processingStatus, setProcessingStatus] = useState<{name: string, status: 'waiting' | 'processing' | 'completed' | 'error'}[]>([]);
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [dragActive, setDragActive] = useState<{ [key: string]: boolean }>({});
  const [modoGeracao, setModoGeracao] = useState<'individual' | 'consolidado'>('individual');
  
  // Risk Matrix Sentinel Events State
  const [eventosSentinela, setEventosSentinela] = useState({
    e1_morte: false,
    e2_hospital: false,
    e2_media: false,
    e3_afastamento: false,
    e4_aumento: false,
    e4_media: false,
    e5_aumento: false,
    e5_media: false,
  });

  const calcularScoreDano = (eventos: typeof eventosSentinela) => {
    let pontos = 0;
    if (eventos.e1_morte)    pontos += 4;
    if (eventos.e2_hospital) pontos += 3;
    if (eventos.e2_media)    pontos += 1;
    if (eventos.e3_afastamento) pontos += 2.5;
    if (eventos.e4_aumento)  pontos += 2;
    if (eventos.e4_media)    pontos += 1.5;
    if (eventos.e5_aumento)  pontos += 1;
    if (eventos.e5_media)    pontos += 0.5;
    return Math.min(pontos / 15.5, 1.0);
  };

  const [excelDanoScore, setExcelDanoScore] = useState<number | null>(null);
  const danoScore = excelDanoScore !== null ? excelDanoScore : calcularScoreDano(eventosSentinela);
  const [calculatedExposicao, setCalculatedExposicao] = useState(0);
  
  const canvasRef = useRef<HTMLCanvasElement>(null);
  const logEndRef = useRef<HTMLDivElement>(null);
  const chartInstance = useRef<Chart | null>(null);

  // Auto-scroll log disabled by user request
  // useEffect(() => {
  //   logEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  // }, [logs]);

  const addLog = (message: string) => {
    const timestamp = new Date().toLocaleTimeString();
    setLogs(prev => [...prev, { id: Math.random().toString(36).substr(2, 9), message, timestamp }]);
  };

  const clearLogs = () => {
    setLogs([]);
  };

  const handleDrag = (e: React.DragEvent, id: string, active: boolean) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(prev => ({ ...prev, [id]: active }));
  };

  const onDrop = (e: React.DragEvent, type: 'molde' | 'excels') => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(prev => ({ ...prev, [type]: false }));
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFiles(e.dataTransfer.files, type);
    }
  };

  const handleFiles = async (files: FileList | null, type: 'molde' | 'excels') => {
    if (!files) return;
    const fileList = Array.from(files);
    
    if (type === 'molde') {
      const docx = fileList.find(f => f.name.endsWith('.docx'));
      if (docx) {
        try {
          const buffer = await docx.arrayBuffer();
          setMolde({ name: docx.name, data: buffer });
          addLog(`Molde carregado: ${docx.name}`);
        } catch (err) {
          addLog(`Erro ao ler molde: ${docx.name}`);
        }
      } else {
        addLog(`Erro: O ficheiro deve ser .docx`);
      }
    } else {
      const xlsxList = fileList.filter(f => f.name.endsWith('.xlsx'));
      if (xlsxList.length > 0) {
        addLog(`Lendo ${xlsxList.length} ficheiro(s)...`);
        
        for (const f of xlsxList) {
          try {
            const buffer = await f.arrayBuffer();
            setExcels(prev => [...prev, { name: f.name, data: buffer }]);
          } catch (err) {
            addLog(`Erro ao ler: ${f.name}`);
          }
        }
        addLog(`${xlsxList.length} ficheiro(s) Excel pronto(s).`);
      }
    }
  };

  const generateCriticosBase64 = async (dadosExcel?: (number|null)[][], totaisExcel?: number[]): Promise<string> => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d');
    if (!ctx) return "";

    const factors = [
      "Ambiente físico de trabalho e equipamentos inadequados",
      "Carga e/ou ritmo de trabalho excessivo",
      "Jornada de trabalho excessiva",
      "Contacto com o sofrimento humano",
      "Contato direto com pessoas agressivas",
      "Ausência de portal/site específico para denúncia de assédio",
      "Ausência de autonomia e controle de decisões",
      "Relacionamento interpessoal conflitante e/ou Isolamento",
      "Desequilíbrio entre vida profissional e familiar"
    ];

    const dados = dadosExcel || [
      [null, null, null],
      [0,    0,    0   ],
      [null, null, null],
      [0,    null, null],
      [null, 0,    null],
      [null, null, null],
      [null, null, 1   ],
      [null, null, 0   ],
      [null, null, null]
    ];

    const colWidths = [450, 120, 120, 120]; 
    const totalWidth = colWidths.reduce((a, b) => a + b, 0);
    const headerRowHeight = 65;
    const dataRowHeight = 40;
    const footerHeight = 75;

    const startX = 20;
    const startY = 20;

    const totalTableHeight = headerRowHeight + (factors.length * dataRowHeight) + footerHeight;
    canvas.width = totalWidth + 40;
    canvas.height = totalTableHeight + 40;

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.lineWidth = 1;

    // --- ROW 1: COMBINED TITLE & HEADERS ---
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(startX, startY, colWidths[0], headerRowHeight);
    ctx.strokeStyle = "#333333";
    ctx.strokeRect(startX, startY, colWidths[0], headerRowHeight);
    
    ctx.textAlign = "center";
    ctx.fillStyle = "#CC0000";
    ctx.font = "bold 15px Arial";
    const titleText = "CENÁRIOS CRÍTICOS por combinações predefinidas de fatores de exposição.";
    const titleLines = wrapText(ctx, titleText, colWidths[0] - 20, 15, true);
    titleLines.forEach((line, li) => {
      ctx.fillText(line, startX + colWidths[0]/2, startY + (headerRowHeight/2 - (titleLines.length * 18)/2) + 14 + (li * 18));
    });

    const headers = [
      { t: "CRÍTICA 1", s: "(Esgotamento profissional;\ncontato com o sofrimento)" },
      { t: "CRÍTICA 2", s: "(Burn-out;\ncontacto com usuários)" },
      { t: "CRÍTICA 3", s: "(Controle das demandas\ndo trabalho)" }
    ];

    headers.forEach((h, i) => {
      const x = startX + colWidths[0] + (i * colWidths[i+1]);
      ctx.fillStyle = "#FFFF00";
      ctx.fillRect(x, startY, colWidths[i+1], headerRowHeight);
      ctx.strokeRect(x, startY, colWidths[i+1], headerRowHeight);
      
      ctx.fillStyle = "#000000";
      ctx.textAlign = "center";
      ctx.font = "900 12px Arial";
      ctx.fillText(h.t, x + colWidths[i+1]/2, startY + 22);
      
      const sublines = wrapText(ctx, h.s.replace('\n', ' '), colWidths[i+1] - 10, 10, false);
      sublines.forEach((line, li) => {
        ctx.fillText(line, x + colWidths[i+1]/2, startY + 36 + (li * 12));
      });
    });

    const sums = totaisExcel || [0, 0, 0];
    const thresholds = [2, 2, 3];
    
    // If not provided, calculate sums (for fallback)
    if (!totaisExcel) {
      dados.forEach(row => {
        row.forEach((v, i) => { if (Number(v) === 1) sums[i]++; });
      });
    }

    let currentY = startY + headerRowHeight;

    factors.forEach((f, idx) => {
      ctx.fillStyle = "#E8F0F8";
      ctx.fillRect(startX, currentY, colWidths[0], dataRowHeight);
      ctx.strokeRect(startX, currentY, colWidths[0], dataRowHeight);
      
      ctx.textAlign = "left";
      ctx.fillStyle = "#000000";
      ctx.font = "900 12px Arial";
      const fLines = wrapText(ctx, f, colWidths[0] - 20, 12, true);
      fLines.forEach((line, li) => {
        ctx.fillText(line, startX + 10, currentY + (dataRowHeight/2 - (fLines.length*15)/2) + 12 + (li * 15));
      });

      for (let c = 0; c < 3; c++) {
        const x = startX + colWidths[0] + (c * colWidths[c+1]);
        const val = dados[idx][c];

        // Specific styling: 1 = Purple, 0 = Pale Blue
        const numericVal = (val !== null && val !== undefined) ? Number(val) : null;

        if (numericVal === 1) {
          ctx.fillStyle = "#9966FF";
          ctx.fillRect(x, currentY, colWidths[c+1], dataRowHeight);
        } else if (numericVal === 0) { // displays "0" and adds light blue bg
          ctx.fillStyle = "#E8F0F8";
          ctx.fillRect(x, currentY, colWidths[c+1], dataRowHeight);
        }
        
        ctx.strokeStyle = "#333333";
        ctx.beginPath();
        ctx.moveTo(x, currentY);
        ctx.lineTo(x, currentY + dataRowHeight);
        ctx.moveTo(x + colWidths[c+1], currentY);
        ctx.lineTo(x + colWidths[c+1], currentY + dataRowHeight);
        
        // Horizontal borders only if value is NOT null/undefined
        if (numericVal !== null) {
          ctx.moveTo(x, currentY);
          ctx.lineTo(x + colWidths[c+1], currentY);
          ctx.moveTo(x, currentY + dataRowHeight);
          ctx.lineTo(x + colWidths[c+1], currentY + dataRowHeight);
        }
        ctx.stroke();

        if (numericVal !== null) {
          ctx.fillStyle = "#000000";
          ctx.textAlign = "center";
          ctx.font = "900 18px Arial";
          ctx.fillText(String(numericVal), x + colWidths[c+1]/2, currentY + dataRowHeight/2 + 7);
        }
      }
      currentY += dataRowHeight;
    });

    // FOOTER
    ctx.fillStyle = "#D3D3D3";
    ctx.fillRect(startX, currentY, colWidths[0], footerHeight);
    ctx.strokeRect(startX, currentY, colWidths[0], footerHeight);
    
    ctx.textAlign = "center";
    ctx.fillStyle = "#CC0000";
    ctx.font = "900 14px Arial";
    ctx.fillText("SITUAÇÃO CRÍTICA PRESENTE", startX + colWidths[0]/2, currentY + 22);

    ctx.fillStyle = "#000000";
    ctx.textAlign = "center";
    const footerText = "A situação crítica específica é ativada quando todas as caixas acima mencionadas ficam roxas, para esse cenário crítico específico. Quando a situação crítica está presente, as caixas do lado direito também ficam roxas.";
    const footerLines = wrapText(ctx, footerText, colWidths[0] - 40, 11, false);
    footerLines.forEach((line, li) => {
      ctx.fillText(line, startX + colWidths[0]/2, currentY + 44 + (li * 14));
    });

    for (let c = 0; c < 3; c++) {
      const x = startX + colWidths[0] + (c * colWidths[c+1]);
      ctx.fillStyle = (sums[c] >= thresholds[c]) ? "#9966FF" : "#FFFFFF";
      ctx.fillRect(x, currentY, colWidths[c+1], footerHeight);
      ctx.strokeRect(x, currentY, colWidths[c+1], footerHeight);
      
      ctx.fillStyle = "#000000";
      ctx.textAlign = "center";
      ctx.font = "900 20px Arial";
      ctx.fillText(String(sums[c]), x + colWidths[c+1]/2, currentY + footerHeight/2 + 8);
    }

    ctx.lineWidth = 2;
    ctx.strokeStyle = "#333333";
    ctx.strokeRect(startX, startY, totalWidth, totalTableHeight);

    return canvas.toDataURL('image/png').split(',')[1];
  };

  const generateHeaderBase64 = async (data: any): Promise<string> => {
    const canvas = document.createElement('canvas');
    canvas.width = 1600;
    canvas.height = 300; 
    const ctx = canvas.getContext('2d');
    if (!ctx) return "";

    // Professional Clean Background
    ctx.fillStyle = '#f8fafc';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    // --- GRID SYSTEM ---
    const drawCleanCard = (x: number, y: number, w: number, h: number, label: string, value: string) => {
      // White Background
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(x, y, w, h);
      
      // Crisp Thin Gray Border - Square Corner
      ctx.strokeStyle = '#7F7F7F'; 
      ctx.lineWidth = 1;
      ctx.strokeRect(x, y, w, h);

      // Label Styling
      ctx.textAlign = 'left';
      ctx.font = '900 15px "Inter", Arial'; 
      ctx.fillStyle = '#1e293b'; 
      ctx.fillText(label, x + 10, y + 22);

      // Value Styling
      ctx.font = '700 26px "Inter", Arial';
      ctx.fillStyle = '#0f172a';
      
      let displayValue = String(value || '---');
      const maxW = w - 20;
      
      if (ctx.measureText(displayValue).width > maxW) {
        ctx.font = '700 20px "Inter", Arial';
        if (ctx.measureText(displayValue).width > maxW) {
          displayValue = displayValue.substring(0, 45) + "...";
        }
      }
      ctx.fillText(displayValue, x + 10, y + 54);
    };

    const margin = 60;
    const gap = 8;
    const startY = 10;
    const cardH = 65; 
    const colW = (canvas.width - (2 * margin) - (2 * gap)) / 3;

    // Row 1
    drawCleanCard(margin, startY, colW * 2 + gap, cardH, 'EMPRESA:', data.EMPRESA);
    drawCleanCard(margin + (colW * 2) + (2 * gap), startY, colW, cardH, 'UNIDADE:', data.UNIDADE);

    // Row 2
    drawCleanCard(margin, startY + cardH + gap, colW * 1.5 + gap/2, cardH, 'SETOR:', data.SETOR);
    drawCleanCard(margin + (colW * 1.5) + (gap * 1.5), startY + cardH + gap, colW * 0.75 - gap/4, cardH, 'TOTAL DE FUNCIONÁRIOS:', String(data.FUNC_TOTAL));
    drawCleanCard(margin + (colW * 2.25) + (gap * 2), startY + cardH + gap, colW * 0.75, cardH, 'TOTAL DE PARTICIPANTES:', String(data.PARTIC_TOTAL));

    // Row 3
    drawCleanCard(margin, startY + (cardH + gap) * 2, colW, cardH, 'INDICE DE PARTICIPAÇÃO:', data.PERC_PARTIC);
    drawCleanCard(margin + colW + gap, startY + (cardH + gap) * 2, colW/2 - gap/2, cardH, 'HOMENS:', String(data.MASC_N));
    drawCleanCard(margin + colW * 1.5 + gap * 1.5, startY + (cardH + gap) * 2, colW/2 - gap/2, cardH, 'MULHERES:', String(data.FEM_N));
    drawCleanCard(margin + colW * 2 + gap * 2, startY + (cardH + gap) * 2, colW, cardH, 'DATA:', data.DATA);

    // Row 4
    drawCleanCard(margin, startY + (cardH + gap) * 3, canvas.width - (2 * margin), cardH, 'AVALIADOR:', data.AVALIADOR);

    return canvas.toDataURL('image/png').split(',')[1];
  };

  const generateMatrixBase64 = async (dScore: number, eScore: number): Promise<string> => {
    const canvas = document.createElement('canvas');
    canvas.width = 1000;
    canvas.height = 435;
    const ctx = canvas.getContext('2d');
    if (!ctx) return "";

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    const startX = 150;
    const startY = 95; // titleH (60) + subtitleH (35)
    const cw = 140;
    const rh = 70;

    const nDano = classificarDano(dScore);
    const nExposicao = classificarExposicao(eScore);

    // EIXO Y — indicadores de linha
    const F322 = nDano.AA251;  // ALTO
    const F323 = nDano.AA250;  // MODERADO
    const F324 = nDano.AA249;  // BAIXO
    const F325 = "";               // DESPREZÍVEL — sempre vazio

    // EIXO X — indicadores de coluna
    const H327 = nExposicao.AA239;   // DESPREZÍVEL
    const I327 = nExposicao.AA235;   // ACEITÁVEL
    const J327 = nExposicao.AA236;   // MODERADO
    const K327 = nExposicao.AA237;   // ALTO
    const L327 = nExposicao.AA238;   // CRÍTICO

    const rowIndicators = [F322, F323, F324, F325];
    const colIndicators = [H327, I327, J327, K327, L327];

    const rowLabels = ["ALTO", "MODERADO", "BAIXO", "DESPREZÍVEL"] as const;
    const colLabels = ["DESPREZÍVEL", "ACEITÁVEL", "MODERADO", "ALTO", "CRÍTICO"] as const;

    const staticColors = [
      ['#FFFF99', '#FFC000', '#FF0000', '#FF0000', '#C4A7FF'], // ALTO
      ['#FFFF97', '#FFFF97', '#FFD347', '#FF0000', '#C4A7FF'], // MODERADO
      ['#C9E7A7', '#FFFF97', '#FFD347', '#FFD347', '#FF0000'], // BAIXO
      ['#C9E7A7', '#C9E7A7', '#FFFF97', '#FFD347', '#FFC000']  // DESPREZÍVEL
    ];

    // Title Block
    ctx.fillStyle = "#FFFF00";
    ctx.fillRect(startX - 130, 0, cw * 5 + 130 + 90, 60);
    ctx.strokeStyle = "#333333";
    ctx.lineWidth = 1;
    ctx.strokeRect(startX - 130, 0, cw * 5 + 130 + 90, 60);
    
    ctx.fillStyle = "#000000";
    ctx.textAlign = "center";
    ctx.font = "bold 15px Arial";
    const titleText = "RISCO DE TRANSTORNOS MENTAIS (stress/depressão) RELACIONADOS AO TRABALHO";
    const titleLines = wrapText(ctx, titleText, cw * 5 + 130 + 70, 15, true);
    titleLines.forEach((line, li) => {
      ctx.fillText(line, startX - 130 + (cw * 5 + 130 + 90)/2, 26 + (li * 18));
    });

    // Subtitle Block
    ctx.fillStyle = "#D6E3BC";
    ctx.fillRect(startX - 130, 60, cw * 5 + 130 + 90, 35);
    ctx.strokeRect(startX - 130, 60, cw * 5 + 130 + 90, 35);
    
    ctx.fillStyle = "#000000";
    ctx.font = "bold 12px Arial";
    ctx.fillText("MATRIZ DE RISCO: EXPOSIÇÃO (fatores de conteúdo e contexto: 'Probabilidade') x EVENTOS SENTINELA ('Danos')", startX - 130 + (cw * 5 + 130 + 90)/2, 60 + 22);

    // Grid
    rowLabels.forEach((rl, r) => {
      // Indicator Column
      ctx.fillStyle = "#EAF1DD";
      ctx.fillRect(startX - 130, startY + (r * rh), 40, rh);
      ctx.strokeRect(startX - 130, startY + (r * rh), 40, rh);
      if (rowIndicators[r] === "X") {
        ctx.fillStyle = "#000000";
        ctx.font = "bold 20px Arial";
        ctx.fillText("x", startX - 110, startY + (r * rh) + rh/2 + 7);
      }

      // Label Row
      ctx.fillStyle = "#3F3F3F";
      ctx.fillRect(startX - 90, startY + (r * rh), 90, rh);
      ctx.strokeRect(startX - 90, startY + (r * rh), 90, rh);
      ctx.fillStyle = "#ffffff";
      ctx.font = "bold 14px Arial";
      ctx.fillText(rl, startX - 45, startY + (r * rh) + rh/2 + 6);

        colLabels.forEach((cl, c) => {
          const isActive = rowIndicators[r] === "X" && colIndicators[c] === "X";
          let color = staticColors[r][c];
          if (isActive) {
            if (rl === "ALTO" && cl === "DESPREZÍVEL") color = "#FF0000";
            if (rl === "DESPREZÍVEL" && cl === "CRÍTICO") color = "#FF0000";
          }

        ctx.fillStyle = color;
        ctx.fillRect(startX + (c * cw), startY + (r * rh), cw, rh);
        ctx.strokeRect(startX + (c * cw), startY + (r * rh), cw, rh);

        if (isActive) {
          ctx.fillStyle = "#000000";
          ctx.font = "bold 24px Arial";
          ctx.fillText("X", startX + (c * cw) + cw/2, startY + (r * rh) + rh/2 + 9);
        }
      });
    });

    // Indicator Row (X-axis)
    colLabels.forEach((cl, c) => {
      ctx.fillStyle = "#D6E3BC";
      ctx.fillRect(startX + (c * cw), startY + (4 * rh), cw, 25);
      ctx.strokeRect(startX + (c * cw), startY + (4 * rh), cw, 25);
      if (colIndicators[c] === "X") {
        ctx.fillStyle = "#000000";
        ctx.font = "bold 20px Arial";
        ctx.fillText("x", startX + (c * cw) + cw/2, startY + (4 * rh) + 18);
      }

      ctx.fillStyle = "#3F3F3F";
      ctx.fillRect(startX + (c * cw), startY + (4 * rh) + 25, cw, rh/2);
      ctx.strokeRect(startX + (c * cw), startY + (4 * rh) + 25, cw, rh/2);
      ctx.fillStyle = "#ffffff";
      ctx.font = "bold 13px Arial";
      ctx.fillText(cl, startX + (c * cw) + cw/2, startY + (4 * rh) + 25 + rh/4 + 5);
    });

    // Exposição Label at the end of indicator row
    ctx.fillStyle = "#D6E3BC";
    ctx.fillRect(startX + (5 * cw), startY + (4 * rh), 90, 25);
    ctx.strokeRect(startX + (5 * cw), startY + (4 * rh), 90, 25);
    ctx.fillStyle = "#000000";
    ctx.font = "bold 14px Arial";
    ctx.fillText("Exposição", startX + (5 * cw) + 45, startY + (4 * rh) + 18);

    return canvas.toDataURL('image/png').split(',')[1];
  };

  const generateExposureSummaryBase64 = async (intrinseca: number, sobrecarga: number): Promise<string> => {
    const canvas = document.createElement('canvas');
    const tableWidth = 900;
    const headerH = 45;
    const headerRowH = 30;
    const dataRowH = 30;
    const totalH = 335; // Matches 17.25cm x 6.46cm ratio with 900 width

    canvas.width = tableWidth;
    canvas.height = totalH;
    const ctx = canvas.getContext('2d');
    if (!ctx) return "";

    // Clear with same background as table
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    const startX = 0;
    let currentY = 0;

    // 1. Header Box (Yellow)
    ctx.fillStyle = "#ffff00";
    ctx.fillRect(startX, currentY, tableWidth, headerH);
    ctx.strokeStyle = "#000000";
    ctx.lineWidth = 2;
    ctx.strokeRect(startX, currentY, tableWidth, headerH);

    ctx.fillStyle = "#000000";
    ctx.textAlign = "center";
    ctx.font = "900 18px 'Inter', Arial";
    ctx.fillText("RESULTADOS DA AVALIAÇÃO DA EXPOSIÇÃO", startX + tableWidth / 2, currentY + 18);
    ctx.fillText("AOS FATORES DE RISCO PSICOSSOCIAIS", startX + tableWidth / 2, currentY + 38);

    currentY += headerH;

    // 2. Table Headers
    const col1W = 300;
    const col2W = 110;
    const col3W = 490;
    const rowH = dataRowH;

    ctx.font = "900 20px 'Inter', Arial";
    ctx.textAlign = "center";
    ctx.fillStyle = "#000000";
    
    // Header row
    ctx.strokeStyle = "#000000";
    ctx.lineWidth = 1.5;
    ctx.strokeRect(startX, currentY, col1W + col2W, rowH);
    ctx.fillText("EXPOSIÇÃO", startX + (col1W + col2W) / 2, currentY + 22);
    
    ctx.strokeRect(startX + col1W + col2W, currentY, col3W, rowH);
    ctx.fillText("AÇÕES", startX + col1W + col2W + col3W / 2, currentY + 22);

    currentY += rowH;

    const rows = [
      { label: "AUSENTE", perc: "0%", color: "#00ffff", action: "INCENTIVAR a manter as condições" },
      { label: "DESPREZÍVEL", perc: "1-23%", color: "#00ff00", action: "MANTER as condições de trabalho" },
      { label: "ACEITÁVEL", perc: "24-47,5%", color: "#ffff00", action: "APOIAR as condições e monitorar elementos críticos" },
      { label: "MODERADO", perc: "47,6-73,5%", color: "#ffa500", action: "MELHORAR, implementando ações preventivas e/ou corretivas" },
      { label: "ALTO", perc: "73,6-95%", color: "#ff0000", action: "RECUPERAR, implementar ações preventivas e/ou corretivas a curto prazo" },
      { label: "CRÍTICO", perc: "SUP 95 %", color: "#cc99ff", action: "AJUSTAR, introduzir urgentemente medidas preventivas e/ou corretivas" }
    ];

    rows.forEach((row) => {
      // Color Column
      ctx.fillStyle = row.color;
      ctx.fillRect(startX, currentY, col1W, rowH);
      ctx.strokeStyle = "#000000";
      ctx.lineWidth = 1;
      ctx.strokeRect(startX, currentY, col1W, rowH);
      
      ctx.fillStyle = "#000000";
      ctx.textAlign = "center";
      ctx.font = "900 18px 'Inter', Arial";
      ctx.fillText(row.label, startX + col1W/2, currentY + 22);

      // Percentage Column
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(startX + col1W, currentY, col2W, rowH);
      ctx.strokeRect(startX + col1W, currentY, col2W, rowH);
      ctx.fillStyle = "#000000";
      ctx.font = "900 15px 'Inter', Arial";
      ctx.fillText(row.perc, startX + col1W + col2W/2, currentY + 22);

      // Action Column
      ctx.strokeRect(startX + col1W + col2W, currentY, col3W, rowH);
      
      const wordsArr = row.action.split(' ');
      const fw = wordsArr[0];
      const rest = wordsArr.slice(1).join(' ');

      ctx.save();
      const colCenterX = startX + col1W + col2W + col3W / 2;

      if (row.action.length < 55) {
        // One line - Centered
        ctx.font = "900 14px 'Inter', Arial";
        const fwW = ctx.measureText(fw).width;
        const restW = ctx.measureText(" " + rest).width;
        const totalW = fwW + restW;

        ctx.textAlign = "left";
        const sx = colCenterX - totalW / 2;
        ctx.fillText(fw, sx, currentY + 21);
        ctx.fillText(" " + rest, sx + fwW, currentY + 21);
      } else {
        // Two lines - Both lines centered
        let splitIdx = 4;
        if (row.label === "MODERADO" || row.label === "ALTO" || row.label === "CRÍTICO") splitIdx = 2;

        const line1TxtArr = wordsArr.slice(0, splitIdx);
        const line2TxtArr = wordsArr.slice(splitIdx);

        const l1F = line1TxtArr[0];
        const l1R = line1TxtArr.slice(1).join(' ');

        // Line 1 centered
        ctx.font = "900 14px 'Inter', Arial";
        const fw1W = ctx.measureText(l1F).width;
        const rest1W = l1R ? ctx.measureText(" " + l1R).width : 0;
        const line1W = fw1W + rest1W;

        const l1sx = colCenterX - line1W / 2;
        ctx.textAlign = "left";
        ctx.fillText(l1F, l1sx, currentY + 14);
        if (l1R) {
          ctx.fillText(" " + l1R, l1sx + fw1W, currentY + 14);
        }

        // Line 2 centered
        ctx.textAlign = "center";
        ctx.fillText(line2TxtArr.join(' '), colCenterX, currentY + 27);
      }
      ctx.restore();

      currentY += rowH;
    });

    currentY += 15; // Gap before stat boxes

    // 4. Improved Stat Boxes
    const boxW = (tableWidth - 20) / 2;
    const boxH = 60;

    const drawStatBox = (x: number, y: number, label: string, value: string) => {
      ctx.fillStyle = "#f8fafc";
      ctx.beginPath();
      ctx.roundRect(x, y, boxW, boxH, 12);
      ctx.fill();
      ctx.strokeStyle = "#cbd5e1";
      ctx.lineWidth = 1;
      ctx.stroke();

      ctx.textAlign = "center";
      ctx.fillStyle = "#475569";
      ctx.font = "bold 13px 'Inter', Arial";
      ctx.fillText(label, x + boxW/2, y + 22);

      ctx.fillStyle = "#1e3a8a";
      ctx.font = "900 28px 'Inter', Arial";
      ctx.fillText(value, x + boxW/2, y + 52);
    };

    drawStatBox(0, currentY, "% Exposição Intrínseca", Math.round(intrinseca) + "%");
    drawStatBox(boxW + 20, currentY, "% Sobrecarga de Exposição", Math.round(sobrecarga) + "%");

    return canvas.toDataURL('image/png').split(',')[1];
  };

  const generateRadarBase64 = async (labels: string[], values: number[]): Promise<string> => {
    return new Promise((resolve) => {
      const canvas = document.createElement('canvas');
      // High resolution for clear presentation
      canvas.width = 1200;
      canvas.height = 1000; 
      const ctx = canvas.getContext('2d');
      if (!ctx) return resolve('');

      // Corporate Background
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);

      const limit = 30; 

      const chartInstance = new Chart(ctx, {
        type: 'radar',
        data: {
          labels: labels.map(l => {
             // Remove leading numbers/question identifiers (e.g. "Questão 1 - ", "1. ", etc)
             let clean = l.replace(/^\d+[\s.-]*/, "").replace(/^Questão\s+\d+[\s.-]*/i, "").trim();
             
             if (clean.length > 25) {
                const words = clean.split(' ');
                const lines = [];
                let current = "";
                words.forEach(w => {
                  if ((current + w).length > 22) {
                    lines.push(current.trim());
                    current = w + " ";
                  } else {
                    current += w + " ";
                  }
                });
                if (current) lines.push(current.trim());
                return lines;
             }
             return clean;
          }),
          datasets: [
            {
              label: 'Referência',
              data: labels.map(() => 100),
              backgroundColor: 'transparent',
              borderColor: '#cbd5e1',
              borderWidth: 1,
              borderDash: [5, 5],
              pointRadius: 0,
              order: 2
            },
            {
              label: 'Resultado',
              data: values,
              backgroundColor: (context: any) => {
                const chart = context.chart;
                const {ctx, chartArea} = chart;
                if (!chartArea) return 'rgba(37, 99, 235, 0.1)';
                const gradient = ctx.createRadialGradient(
                  (chartArea.left + chartArea.right) / 2,
                  (chartArea.top + chartArea.bottom) / 2,
                  0,
                  (chartArea.left + chartArea.right) / 2,
                  (chartArea.top + chartArea.bottom) / 2,
                  350
                );
                gradient.addColorStop(0, 'rgba(59, 130, 246, 0.4)');
                gradient.addColorStop(1, 'rgba(37, 99, 235, 0.05)');
                return gradient;
              },
              borderColor: '#2563eb',                    
              borderWidth: 5, 
              pointBackgroundColor: '#ffffff',
              pointBorderColor: '#2563eb',
              pointBorderWidth: 3, 
              pointRadius: 8,
              tension: 0,
              order: 1
            }
          ]
        },
        options: {
          responsive: false,
          animation: false,
          devicePixelRatio: 3, 
          layout: {
            padding: {
              top: 170,
              bottom: 80, 
              left: 100,
              right: 100
            }
          },
          plugins: {
            legend: { display: false },
            tooltip: { enabled: false },
            datalabels: { display: false } 
          },
          scales: {
            r: {
              beginAtZero: true,
              min: 0,
              max: limit,
              grid: {
                color: (context: any) => {
                  if (context.tick.value === 30) return 'rgba(37, 99, 235, 0.3)';
                  return '#e2e8f0';
                },
                lineWidth: 1
              },
              angleLines: {
                color: '#cbd5e1',
                lineWidth: 1
              },
              ticks: { 
                display: true, 
                stepSize: 5,
                font: { size: 11, weight: 600 }, 
                color: '#1e40af',
                backdropColor: 'rgba(255, 255, 255, 0.75)',
                callback: (val) => val + "%"
              },
              pointLabels: {
                font: { 
                  size: 14, 
                  weight: 700,
                  family: "'Inter', sans-serif" 
                }, 
                color: '#1e3a8a',
                padding: 10 
              }
            }
          }
        },
        plugins: [{
          id: 'vibrant-draw',
          beforeDraw(chart: any) {
            const { ctx, width, height } = chart;
            ctx.save();
            
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';
            
            ctx.font = 'bold 36px "Inter", Arial';
            ctx.fillStyle = '#1e3a8a';
            ctx.fillText('CENÁRIOS CRÍTICOS: combinações dos fatores de', width / 2, 60);
            ctx.fillText('exposição aos riscos psicossociais', width / 2, 105);

            ctx.font = '600 22px "Inter", Arial';
            ctx.fillStyle = '#3b82f6';
            ctx.fillText('Distribuição percentual por fator avaliado', width / 2, 155);

            // Legend removed to save space
            ctx.restore();
          }
        }, {
          id: 'vibrant-datalabels',
          afterDatasetsDraw(chart: any) {
            const { ctx, scales: { r } } = chart;
            ctx.textAlign = 'center';
            ctx.textBaseline = 'middle';

            chart.data.datasets[1].data.forEach((val: number, i: number) => {
               if (val === undefined || val === null || val <= 0) return;
               const point = r.getPointPositionForValue(i, val);
               
               const angle = r.getIndexAngle(i) - Math.PI / 2;
               const dist = 35;
               const lx = point.x + Math.cos(angle) * dist;
               const ly = point.y + Math.sin(angle) * dist;

               const txt = Math.round(val) + '%';
               ctx.font = 'bold 20px "Inter", Arial';
               const tw = ctx.measureText(txt).width;
               
               // Styled bubble
               ctx.fillStyle = '#2563eb';
               ctx.beginPath();
               ctx.roundRect(lx - (tw/2) - 10, ly - 16, tw + 20, 32, 8);
               ctx.fill();

               ctx.fillStyle = '#ffffff';
               ctx.fillText(txt, lx, ly);
            });
          }
        }, {
          id: 'chart-border',
          afterDraw(chart: any) {
            const { ctx, width, height } = chart;
            ctx.save();
            ctx.strokeStyle = "#7F7F7F";
            ctx.lineWidth = 1; 
            ctx.strokeRect(0.5, 0.5, width - 1, height - 1);
            ctx.restore();
          }
        }]
      });

      setTimeout(() => {
        const dataUrl = canvas.toDataURL('image/png').split(',')[1];
        resolve(dataUrl);
        chartInstance.destroy();
      }, 200);
    });
  };

  const generateReports = async () => {
    if (!molde || excels.length === 0) return;
    
    setIsProcessing(true);
    setProcessingStatus(excels.map(f => ({ name: f.name, status: 'waiting' })));
    addLog(`🚀 Iniciando processamento em lote (${excels.length} arquivos)...`);

    try {
      const moldeBuffer = molde.data;
      const meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"];
      
      const generatedFiles: { name: string; blob: Blob }[] = [];
      const allReportData: any[] = []; 

      for (let i = 0; i < excels.length; i++) {
        const file = excels[i];
        setProcessingStatus(prev => prev.map((s, idx) => idx === i ? { ...s, status: 'processing' } : s));
        
        try {
          addLog(`📄 Lendo arquivo ${i + 1}/${excels.length}: ${file.name}`);

          const data = file.data;
          const workbook = XLSX.read(data, { type: 'array', cellDates: true });
          
          const sheetName = workbook.SheetNames.find(n => n.toUpperCase().includes('PSICOSSOCIAL')) || workbook.SheetNames[0];
          const sheet = workbook.Sheets[sheetName];

          if (!sheet) {
            throw new Error(`Aba 'PSICOSSOCIAL' não encontrada no arquivo ${file.name}`);
          }

        // --- Helper Discovery Functions ---
        // Iterate in sorted order to ensure consistent "first match" behavior (top-to-bottom, left-to-right)
        const sortedKeys = Object.keys(sheet).sort((a, b) => {
          const ma = a.match(/([A-Z]+)(\d+)/);
          const mb = b.match(/([A-Z]+)(\d+)/);
          if (!ma || !mb) return 0;
          const rA = parseInt(ma[2]), rB = parseInt(mb[2]);
          if (rA !== rB) return rA - rB;
          return ma[1].localeCompare(mb[1]);
        });

        const findCellWith = (txt: string | string[], exactPrioritize = true) => {
          const targets = Array.isArray(txt) ? txt.map(t => t.toUpperCase().trim()) : [txt.toUpperCase().trim()];
          let partialMatch = null;
          
          for (let z of sortedKeys) {
            const cell = sheet[z];
            if (cell && cell.v) {
              const val = String(cell.v).trim().toUpperCase();
              if (targets.some(t => val === t)) {
                 const m = z.match(/([A-Z]+)(\d+)/);
                 if (m) return { c: m[1], r: parseInt(m[2]), v: cell.v };
              }
              if (targets.some(t => val.includes(t)) && !partialMatch) {
                 const m = z.match(/([A-Z]+)(\d+)/);
                 if (m) partialMatch = { c: m[1], r: parseInt(m[2]), v: cell.v };
              }
            }
          }
          return !exactPrioritize && partialMatch ? partialMatch : (partialMatch || null);
        };

        const smartParse = (val: any) => {
          if (val === undefined || val === null || val === "") return 0;
          if (typeof val === 'number') return val;
          
          let s = String(val).trim().replace(/\s/g, "");
          
          // Handle percentage strings like "75%", "75,0%"
          if (s.includes("%")) {
             let clean = s.replace(/\./g, "").replace(",", ".").replace(/[^-0-9.]/g, "");
             const n = parseFloat(clean);
             if (!isNaN(n)) return n / 100; // Return as decimal 0.75
          }

          // Handle BR format: 1.000,50 -> 1000.50
          let clean = s.replace(/\./g, "").replace(",", ".").replace(/[^-0-9.]/g, "");
          const res = parseFloat(clean);
          return isNaN(res) ? 0 : res;
        };

        const getValSmart = (labels: string | string[], type: 'any' | 'number' = 'any') => {
          const list = Array.isArray(labels) ? labels : [labels];
          for (const label of list) {
            const loc = findCellWith(label);
            if (!loc) continue;

            const currentStr = String(loc.v).toUpperCase().trim();
            // Handle "LABEL: VALUE" case - expanded to handle multiline or varied separators
            if (currentStr.includes(":") || currentStr.includes("-")) {
              const parts = currentStr.split(/[:\-]/);
              if (parts.length > 1 && parts[1].trim()) {
                 const val = parts.slice(1).join(":").trim();
                 if (type === 'number') {
                    const n = smartParse(val);
                    if (n !== 0 || val === "0") return n;
                 } else {
                    return val;
                 }
              }
            }

            // Neighbor search: expanded range significantly
            const searchOffsets = [
              { c: 0, r: 0 }, { c: 1, r: 0 }, { c: 2, r: 0 }, { c: 3, r: 0 }, { c: 4, r: 0 }, { c: 5, r: 0 }, { c: 6, r: 0 }, { c: 7, r: 0 }, { c: 8, r: 0 },
              { c: 0, r: 1 }, { c: 1, r: 1 }, { c: 2, r: 1 }, { c: 3, r: 1 }, { c: 4, r: 1 }, { c: 5, r: 1 },
              { c: 0, r: 2 }, { c: 1, r: 2 }, { c: 2, r: 2 },
              { c: 0, r: 3 }
            ];

            for (const offset of searchOffsets) {
              const col = XLSX.utils.encode_col(XLSX.utils.decode_col(loc.c) + offset.c);
              const row = loc.r + offset.r;
              const cell = sheet[`${col}${row}`];
              
              if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== "") {
                const cellStr = String(cell.v).trim();
                // If it's the exact same as label and we are at offset 0,0 skip
                if (offset.c === 0 && offset.r === 0 && cellStr.toUpperCase() === label.toUpperCase()) continue;
                
                if (type === 'number') {
                  const n = smartParse(cell.v);
                  if (n !== 0 || String(cell.v) === "0") return n;
                } else {
                  return cell.v;
                }
              }
            }
          }
          return null;
        };

        const shorten = (str: any, max: number = 42) => {
          const s = String(str || "").toUpperCase().trim();
          return s.length <= max ? s : s.substring(0, max - 3) + "...";
        };

        // --- Core Info ---
        const vEmpresaRaw = sheet['E12']?.v || getValSmart(["EMPRESA", "RAZÃO SOCIAL", "CLIENTE"]) || "NÃO IDENTIFICADO";
        const vSetorRaw = sheet['E16']?.v || getValSmart(["SETOR", "DEPARTAMENTO", "ÁREA"]) || "NÃO IDENTIFICADO";
        const vUnidadeRaw = sheet['N12']?.v || getValSmart(["UNIDADE", "LOCAL", "ESTABELECIMENTO"]) || "";
        const vCnpjRaw = sheet['F14']?.v || getValSmart(["CNPJ", "C.N.P.J."]) || "";
        const vAvaliadorRaw = sheet['E18']?.v || getValSmart(["AVALIADOR", "RESPONSÁVEL", "ELABORADO POR"]) || "";
        
        addLog(`>> Base: ${vEmpresaRaw} | Setor: ${vSetorRaw}`);

        let rawD = sheet['K18']?.v || sheet['K16']?.v || getValSmart(["DATA DA AVALIAÇÃO", "DATA DA AVALIACAO", "DATA DA COLETA", "DATA", "ATUALIZADO EM"]);
        let d = new Date();
        if (rawD instanceof Date) {
          d = rawD;
        } else if (typeof rawD === 'number') {
          // Excel serial number
          d = new Date(Math.round((rawD - 25569) * 86400 * 1000));
        } else if (rawD) {
          // Try regex to find date patterns in string (e.g. "São Paulo, 10 de maio de 2024" or "10/05/2024")
          const dateStr = String(rawD);
          const ddmmyyyy = dateStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
          if (ddmmyyyy) {
            const day = parseInt(ddmmyyyy[1]);
            const month = parseInt(ddmmyyyy[2]) - 1;
            let year = parseInt(ddmmyyyy[3]);
            if (year < 100) year += 2000;
            d = new Date(year, month, day);
          } else {
            const parsed = Date.parse(dateStr);
            if (!isNaN(parsed)) d = new Date(parsed);
          }
        }
        
        // Final verification for d
        if (!d || isNaN(d.getTime()) || d.getFullYear() < 1920) d = new Date(); 

        // --- Demografia ---
        const fTot = smartParse(sheet['L16']?.v || getValSmart(["TOTAL DE FUNCIONÁRIOS", "FUNC. TOTAL", "EFETIVOS"], 'number'));
        const pTot = smartParse(sheet['O16']?.v || getValSmart(["PARTICIPANTES", "PONTUADOS", "TOTAL PARTICIPANTES", "PESSOAS"], 'number'));
        
        let mN = smartParse(sheet['R16']?.v || getValSmart(["HOMENS", "MASCULINO"], 'number') || 0);
        let wN = smartParse(sheet['T16']?.v || getValSmart(["MULHERES", "FEMININO"], 'number') || 0);
        
        // If the mapping-specific cells are 0, try smart search
        if (mN === 0 && wN === 0) {
          const hLoc = findCellWith(["HOMENS", "MASCULINO"]);
          if (hLoc) {
            for (let i = 1; i <= 15; i++) {
              const val = sheet[`${XLSX.utils.encode_col(XLSX.utils.decode_col(hLoc.c) + i)}${hLoc.r}`]?.v;
              if (val !== undefined && val !== null && val !== "" && !isNaN(parseFloat(String(val)))) {
                mN = smartParse(val);
                break;
              }
            }
          }
          const mLocArr = findCellWith(["MULHERES", "FEMININO"]);
          if (mLocArr) {
            for (let i = 1; i <= 15; i++) {
              const val = sheet[`${XLSX.utils.encode_col(XLSX.utils.decode_col(mLocArr.c) + i)}${mLocArr.r}`]?.v;
              if (val !== undefined && val !== null && val !== "" && !isNaN(parseFloat(String(val)))) {
                wN = smartParse(val);
                break;
              }
            }
          }
        }
        
        const pParticLimit = fTot > 0 ? (pTot / fTot * 100) : 0;
        const totalGenero = mN + wN;
        
        // Sync participation if genres sum is higher
        const finalPTot = (pTot === 0 && totalGenero > 0) ? totalGenero : pTot;
        const finalPPartic = fTot > 0 ? (finalPTot / fTot * 100) : 0;

        const pMasc = totalGenero > 0 ? (mN / totalGenero * 100) : 0;
        const pFem = totalGenero > 0 ? (wN / totalGenero * 100) : 0;
        
        addLog(`>> Demografia: ${finalPTot}/${fTot} part., M:${mN}, F:${wN}`);
        if (totalGenero !== finalPTot && finalPTot > 0 && totalGenero > 0) {
          addLog(`⚠️ Divergência: Gêneros (${totalGenero}) ≠ Total Participantes (${finalPTot})`);
        }

    // --- Exposição Variables (Fórmulas do Excel) ---
    // PASSO 1: Score base de exposição (AA232)
    const AD301 = smartParse(sheet['AD301']?.v || getValSmart(["FATOR PASSADO", "AD301"], "number"));
    const AB26 = smartParse(sheet['AB26']?.v || getValSmart(["NORMALIZAÇÃO", "AB26"], "number"));
    const AA232_bruto = (AD301 === 1.0) ? 1.0 : AB26;

    // PASSO 2: Multiplicador de turno (BM248 -> AB40 da aba AVAL GERAL)
    const avalGeralSheetName = workbook.SheetNames.find(n => n.toUpperCase().includes('AVAL GERAL'));
    const avalSheet = avalGeralSheetName ? workbook.Sheets[avalGeralSheetName] : null;
    const BM248 = smartParse(avalSheet?.['AB40']?.v || sheet['BM248']?.v || getValSmart("BM248", "number") || 1.0);

    // PASSO 3 & 4: Score final de exposição para a matriz (AA233)
    const AA233_cell_v = (sheet['AA233'] !== undefined && sheet['AA233'].v !== null) ? sheet['AA233'].v : (sheet['P248'] !== undefined ? sheet['P248'].v : undefined);
    const C249 = String(sheet['C249']?.v || "").trim().toUpperCase();
    
    let currentAA233 = 0;
    if (AA233_cell_v !== undefined && AA233_cell_v !== null) {
      currentAA233 = smartParse(AA233_cell_v);
    } else {
      currentAA233 = (C249 === "X") ? 0 : (AA232_bruto * BM248);
    }
    
    setCalculatedExposicao(currentAA233);
    addLog(`>> Exposição Matriz (AA233): ${(currentAA233 * 100).toFixed(2)}%`);

    // PASSO 5: Score de Dano (AC229 / AA246)
    const AC229_cell_v = (sheet['AC229'] !== undefined && sheet['AC229'].v !== null) ? sheet['AC229'].v : (sheet['AA246'] !== undefined ? sheet['AA246'].v : undefined);
    let currentDanoImported = smartParse(AC229_cell_v);
    
    // Fallback: calcular se AC229 estiver zerado mas AC228/AD223 existirem
    if (currentDanoImported === 0) {
      const AC228 = smartParse(sheet['AC228']?.v);
      const AD223 = smartParse(sheet['AD223']?.v);
      if (AC228 > 0 && AD223 > 0) {
        currentDanoImported = AC228 / AD223;
        addLog(`>> Score Dano calculado via AC228/AD223: ${(currentDanoImported * 100).toFixed(2)}%`);
      }
    }

    let currentDanoScore = danoScore; // Default to manual/state
    if (currentDanoImported > 0) {
      currentDanoScore = currentDanoImported;
      setExcelDanoScore(currentDanoImported);
      addLog(`>> Score de Dano Importado: ${(currentDanoImported * 100).toFixed(2)}%`);
    } else {
      addLog(`>> Usando Score de Dano manual: ${(currentDanoScore * 100).toFixed(2)}%`);
    }

    const expIntrinseca = AA232_bruto * 100; // Para exibição no stat box
    const expSobrecarga = currentAA233 * 100; // Representação de sobrecarga

        // --- Factors / Radar Data ---
        const targetFactors = [
          { key: "Ambiente físico de trabalho e equipamentos", search: ["AMBIENTE FÍSICO", "AMBIENTE FISICO", "EQUIPAMENTOS"] },
          { key: "Carga e ritmo de trabalho", search: ["CARGA E RITMO", "CARGA E/OU RITMO"] },
          { key: "Jornada de trabalho", search: ["JORNADA", "DURAÇÃO DO TRABALHO"] },
          { key: "Contato com sofrimento humano", search: ["SOFRIMENTO HUMANO", "CONTATO COM O SOFRIMENTO"] },
          { key: "Contato com pessoas agressivas", search: ["CONTATO COM PESSOAS AGRESSIVAS", "PESSOAS AGRESSIVAS"] },
          { key: "Denúncia de assédio/violência", search: ["DENÚNCIA DE ASSÉDIO", "DENÚNCIA DE ASSEDIO", "ASSÉDIO", "VIOLÊNCIA"] },
          { key: "Autonomia de decisões e controle", search: ["AUTONOMIA", "CONTROLE DE DECISÕES"] },
          { key: "Relações interpessoais / Isolamento", search: ["INTERPESSOAL", "RELAÇÕES INTERPESSOAIS", "CONFLITANTE"] },
          { key: "Equilíbrio entre vida profissional e familiar", search: ["EQUILÍBRIO ENTRE VIDA PROFISSIONAL", "FAMILIAR", "VIDA PROFISSIONAL E FAMILIAR"] }
        ];
        let labels: string[] = [];
        let values: number[] = [];

        // Try mapping area first (H277-H285 / I277-I285 is standard for these percentages)
        const ranges = [
          { l: 'H277', v: 'I277' }, { l: 'H278', v: 'I278' }, { l: 'H279', v: 'I279' },
          { l: 'H280', v: 'I280' }, { l: 'H281', v: 'I281' }, { l: 'H282', v: 'I282' },
          { l: 'H283', v: 'I283' }, { l: 'H284', v: 'I284' }, { l: 'H285', v: 'I285' }
        ];

        ranges.forEach((range, idx) => {
          const lCell = sheet[range.l];
          const vCell = sheet[range.v];
          if (lCell && lCell.v) {
            labels.push(String(lCell.v).trim());
            let val = smartParse(vCell?.v);
            if (val > 0 && val <= 1.05) val *= 100;
            values.push(Math.round(val));
          } else {
            // Placeholder label if mapping fails but we want to keep the slot
            labels.push(targetFactors[idx]?.key || "Fator " + (idx + 1));
            values.push(0);
          }
        });

        // Fallback search ONLY if everything is zero or we missed labels
        if (values.every(v => v === 0)) {
          const currentLabels = [...labels];
          const currentValues = [...values];
          
          labels = [];
          values = [];
          const factAnchor = findCellWith(["PONTUAÇÃO POR FATOR", "QUADRO DE EXPOSIÇÃO", "FATOR DE EXPOSIÇÃO"]);
          const startR = factAnchor ? factAnchor.r : 200;
          
          for (const factor of targetFactors) {
            let foundVal = 0;
            let foundLabel = factor.key;
            let ok = false;
            
            for (let r = startR - 10; r < startR + 100; r++) {
              if (r < 1) continue;
              const cellA = String(sheet[`A${r}`]?.v || "").toUpperCase();
              const cellB = String(sheet[`B${r}`]?.v || "").toUpperCase();
              const cellC = String(sheet[`C${r}`]?.v || "").toUpperCase();
              const cellH = String(sheet[`H${r}`]?.v || "").toUpperCase();
              
              const match = (val: string) => factor.search.some(s => val.includes(s)) && val.length < 80;

              if (match(cellA) || match(cellB) || match(cellC) || match(cellH)) {
                foundLabel = String(sheet[`A${r}`]?.v || sheet[`B${r}`]?.v || sheet[`C${r}`]?.v || sheet[`H${r}`]?.v).trim();
                // Look for the percentage to the right
                const startCol = match(cellH) ? 'H' : (match(cellC) ? 'C' : (match(cellB) ? 'B' : 'A'));
                for (let i = 1; i <= 30; i++) {
                  const col = XLSX.utils.encode_col(XLSX.utils.decode_col(startCol) + i);
                  const v = sheet[`${col}${r}`]?.v;
                  if (v !== undefined && v !== null && v !== "" && String(v).toUpperCase() !== "X") {
                    let parsed = smartParse(v);
                    if (parsed > 0 && parsed <= 1.05) parsed *= 100;
                    foundVal = Math.round(parsed);
                    ok = true;
                    break;
                  }
                }
                if (ok) break;
              }
            }
            labels.push(foundLabel);
            values.push(foundVal);
          }

          // If fallback found nothing better than initial mapping, revert to mapping
          if (values.every(v => v === 0) && currentValues.length > 0) {
            labels = currentLabels;
            values = currentValues;
          }
        }
        
        addLog(`>> Fatores encontrados: ${labels.length}/9`);
        if (values.every(v => v === 0)) addLog("ℹ️ Radar zerado (conforme Excel).");

        const extraData: any = {};
        
        // --- 2. Risk Matrix ---
        let fRowIdx = -1, fColIdx = -1;
        const pRows = ["ALTO", "MODERADO", "BAIXO", "DESPREZÍVEL"];
        const pCols = ["DESPREZÍVEL", "ACEITÁVEL", "MODERADO", "ALTO", "CRÍTICO"];

        const matrixAnchor = findCellWith("MATRIZ DE RISCO") || findCellWith("EVENTO SENTINELA") || findCellWith("GRAVIDADE") || findCellWith("PROBABILIDADE");
        const startRMatrix = matrixAnchor ? matrixAnchor.r : 200;
        
        // Strategy: find the "X" in the matrix grid first
        let matrixX: { r: number, c: number } | null = null;
        for (let r = startRMatrix; r < startRMatrix + 100; r++) {
          for (let c = 0; c < 20; c++) {
            const col = XLSX.utils.encode_col(c);
            const val = String(sheet[`${col}${r}`]?.v || "").trim().toUpperCase();
            if (val === "X" || val === "V" || val === "1") {
              // Confirm this "X" is near matrix labels to avoid false positives
              let nearLabel = false;
              for (let dr = -10; dr <= 10; dr++) {
                const rowText = String(sheet[`A${r+dr}`]?.v || sheet[`B${r+dr}`]?.v || sheet[`C${r+dr}`]?.v || "").toUpperCase();
                if (pRows.some(pr => rowText.includes(pr)) || pCols.some(pc => rowText.includes(pc))) {
                  nearLabel = true;
                  break;
                }
              }
              if (nearLabel) {
                matrixX = { r, c };
                break;
              }
            }
          }
          if (matrixX) break;
        }

        if (matrixX) {
          // Find row label for this X
          for (let dr = -10; dr <= 10; dr++) {
            const rowText = String(sheet[`A${matrixX.r+dr}`]?.v || sheet[`B${matrixX.r+dr}`]?.v || sheet[`C${matrixX.r+dr}`]?.v || "").toUpperCase();
            for (let ri = 0; ri < pRows.length; ri++) {
              if (rowText.includes(pRows[ri])) { fRowIdx = ri + 1; break; }
            }
            if (fRowIdx !== -1) break;
          }
          // Find col label for this X
          for (let dr = -10; dr <= 10; dr++) {
            const rowText = String(sheet[`A${matrixX.r+dr}`]?.v || sheet[`B${matrixX.r+dr}`]?.v || sheet[`C${matrixX.r+dr}`]?.v || "").toUpperCase();
            for (let ci = 0; ci < pCols.length; ci++) {
              if (rowText.includes(pCols[ci])) { fColIdx = ci + 1; break; }
            }
            if (fColIdx !== -1) break;
          }
        }

        for(let r=1; r<=4; r++) for(let c=1; c<=5; c++) extraData[`M${r}${c}`] = "";
        if (fRowIdx !== -1 && fColIdx !== -1) {
          extraData[`M${fRowIdx}${fColIdx}`] = "X";
          addLog(`>> Matriz detectada: R:${fRowIdx} C:${fColIdx}`);
        } else {
          addLog(">> Matriz de risco não detectada automaticamente.");
        }

        // --- 1. Critical Scenarios ---
        const critAnchor = findCellWith("CENÁRIOS CRÍTICOS") || findCellWith("CRÍTICA 1") || findCellWith("CRITICA 1") || findCellWith("SITUAÇÃO CRÍTICA");
        if (critAnchor) {
          const startR = critAnchor.r;
          const baseColIdx = XLSX.utils.decode_col(critAnchor.c);
          
          let actualStart = startR;
          // Look for factor names to align rows
          for (let r = startR - 10; r < startR + 25; r++) {
             if (r < 1) continue;
             const rowVal = String(sheet[`A${r}`]?.v || sheet[`B${r}`]?.v || sheet[`C${r}`]?.v || "").toUpperCase();
             if (rowVal.includes("AMBIENTE") || rowVal.includes("FÍSICO") || rowVal.includes("CARGA") || rowVal.includes("RITMO")) { 
               actualStart = r - 1; 
               break; 
             }
          }

          // Find Column Indices for Critica 1, 2, 3 independently with wider range
          let colIdx1 = -1, colIdx2 = -1, colIdx3 = -1;
          for (let r = Math.max(1, startR - 20); r < startR + 30; r++) {
            for (let c = 0; c < 50; c++) {
              const head = String(sheet[`${XLSX.utils.encode_col(c)}${r}`]?.v || "").toUpperCase();
              if (head.includes("CRÍTICA 1") || head.includes("CRITICA 1")) colIdx1 = c;
              if (head.includes("CRÍTICA 2") || head.includes("CRITICA 2")) colIdx2 = c;
              if (head.includes("CRÍTICA 3") || head.includes("CRITICA 3")) colIdx3 = c;
            }
          }

          // Fallback column logic if headers not found via text
          if (colIdx1 === -1 && baseColIdx !== -1) colIdx1 = baseColIdx + 5; // common offset
          if (colIdx2 === -1 && colIdx1 !== -1) colIdx2 = colIdx1 + 1;
          if (colIdx3 === -1 && colIdx2 !== -1) colIdx3 = colIdx2 + 1;

          const targetCols = [colIdx1, colIdx2, colIdx3];
          if (colIdx1 !== -1) {
            for (let l = 1; l <= 9; l++) {
              const row = actualStart + l;
              for (let cIdx = 0; cIdx < 3; cIdx++) {
                 if (targetCols[cIdx] === -1) continue;
                 const cellValRaw = sheet[`${XLSX.utils.encode_col(targetCols[cIdx])}${row}`]?.v;
                 extraData[`C${cIdx+1}_L${l}`] = (cellValRaw !== undefined && cellValRaw !== null) ? String(cellValRaw).trim() : "";
              }
            }
            
            let sitRow = -1; 
            for(let sr = actualStart + 5; sr < actualStart + 100; sr++) {
              // Search columns A through J for the signature text
              for (let sc = 0; sc < 10; sc++) {
                const colLetter = XLSX.utils.encode_col(sc);
                const vText = String(sheet[`${colLetter}${sr}`]?.v || "").toUpperCase();
                if (vText.includes("SITUAÇÃO CRÍTICA") || vText.includes("SITUACAO CRITICA")) {
                  sitRow = sr;
                  break;
                }
              }
              if (sitRow !== -1) break;
            }
            if (sitRow !== -1) {
              for (let c = 0; c < 3; c++) {
                if (targetCols[c] === -1) continue;
                const sValRaw = sheet[`${XLSX.utils.encode_col(targetCols[c])}${sitRow}`]?.v;
                if (sValRaw !== undefined && sValRaw !== null && String(sValRaw).trim() !== "") {
                   extraData[`C${c+1}_SIT`] = String(sValRaw).trim().toUpperCase();
                }
              }
            }
          }
        }

        // --- Cenários Críticos (Dynamic Mapping) ---
        const dynamicCritMapping: (number | null)[][] = Array.from({ length: 9 }, () => [null, null, null]);
        for (let l = 1; l <= 9; l++) {
          for (let c = 1; c <= 3; c++) {
            const valStr = extraData[`C${c}_L${l}`];
            if (valStr === "1") dynamicCritMapping[l-1][c-1] = 1;
            else if (valStr === "0") dynamicCritMapping[l-1][c-1] = 0;
          }
        }

        const critMapping = dynamicCritMapping;

        const critTotais: number[] = [
          parseInt(extraData['C1_SIT']) || 0,
          parseInt(extraData['C2_SIT']) || 0,
          parseInt(extraData['C3_SIT']) || 0
        ];
        
        // Recalcular totais se vierem zerados
        if (critTotais[0] === 0 && critTotais[1] === 0 && critTotais[2] === 0) {
           for (let c = 0; c < 3; c++) {
              let s = 0;
              for (let r = 0; r < 9; r++) if (critMapping[r][c] === 1) s++;
              critTotais[c] = s;
           }
        }

        const critB64 = await generateCriticosBase64(critMapping, critTotais);
        const matrixB64 = await generateMatrixBase64(currentDanoScore, currentAA233);

        const dateStr = `${d.getDate().toString().padStart(2,'0')}/${(d.getMonth()+1).toString().padStart(2,'0')}/${d.getFullYear()}`;

        // --- Format Data ---
        const formatPerc = (val: any) => {
          let num = smartParse(val);
          if (num > 0 && num <= 1.05) num *= 100;
          return Math.round(num) + "%";
        };

        const radarB64 = labels.length > 0 ? await generateRadarBase64(labels, values) : null;
        
        // --- Pre-assemble report data to generate Header Image ---
        const tempReportData: any = {
          EMPRESA: String(vEmpresaRaw).toUpperCase(),
          UNIDADE: shorten(vUnidadeRaw, 80),
          SETOR: String(vSetorRaw).toUpperCase(),
          DATA: dateStr,
          AVALIADOR: vAvaliadorRaw || "NÃO IDENTIFICADO",
          FUNC_TOTAL: Math.round(fTot),
          PARTIC_TOTAL: Math.round(finalPTot),
          MASC_N: Math.round(mN),
          FEM_N: Math.round(wN),
          PERC_PARTIC: formatPerc(finalPPartic)
        };

        const headerImageB64 = await generateHeaderBase64(tempReportData);

        const reportData = {
          // Headers & Identificação
          headerImage: headerImageB64,
          INFO_HEADER: headerImageB64, // Tag to be used in Word as {%INFO_HEADER}

          // Top Level / Capa
          EMPRESA: tempReportData.EMPRESA,
          CNPJ: vCnpjRaw || "NÃO CONSTA",
          UNIDADE: tempReportData.UNIDADE,
          SETOR: tempReportData.SETOR,
          DATA: tempReportData.DATA,
          DATA_EXTENSA: `${d.getDate()} de ${meses[d.getMonth()]} de ${d.getFullYear()}`,
          MES_ANO_CAPA: `${meses[d.getMonth()].toUpperCase()} / ${d.getFullYear()}`,
          AVALIADOR: vAvaliadorRaw || "NÃO IDENTIFICADO",
          TOTAL_PAGINAS: includePlano ? "09" : "07",
          TEM_PLANO: includePlano,
          
          // Demografia
          FUNC_TOTAL: Math.round(fTot),
          PARTIC_TOTAL: Math.round(finalPTot),
          MASC_N: Math.round(mN),
          FEM_N: Math.round(wN),
          PERC_PARTIC: formatPerc(finalPPartic), 
          SOMA_P16: formatPerc(finalPPartic),
          PERC_EFETIVOS: formatPerc(finalPPartic),
          PERC_MASC: Math.round(pMasc) + "%",
          PERC_FEM: Math.round(pFem) + "%",
          
          // Exposição
          EXP_INTRINSECA: expIntrinseca.toFixed(1) + "%",
          EXP_SOBRECARGA: expSobrecarga.toFixed(1) + "%",
          
          // Conclusões
          CONCLUSOES_LISTA: labels.length > 0 ? labels.map((l, i) => `• ${l} = ${Math.round(values[i] || 0)}%`).join('\n') : "Fatores dentro da normalidade.",
          
          // Imagens
          radarImage: radarB64,
          GRAFICO: radarB64,
          criticosImage: critB64,
          TABELA_CRITICOS: critB64,
          matrixImage: matrixB64,
          MATRIZ_RISCO: matrixB64,
          TABELA_RESUMO_EXPOSICAO: await generateExposureSummaryBase64(expIntrinseca, expSobrecarga),
          TABELA_EXPOSICAO: await generateExposureSummaryBase64(expIntrinseca, expSobrecarga),
          TABELA_RESUMO: await generateExposureSummaryBase64(expIntrinseca, expSobrecarga),
          IMAGEM_EXPOSICAO: await generateExposureSummaryBase64(expIntrinseca, expSobrecarga),
          
          // Spread factor rows for backward compatibility
          ...extraData 
        };

        const fileName = `Relatório Psicossocial_${reportData.EMPRESA}_${reportData.UNIDADE}_${reportData.SETOR}_${d.getDate().toString().padStart(2,'0')}_${(d.getMonth()+1).toString().padStart(2,'0')}_${d.getFullYear()}`.replace(/[\/\\?%*:|"<>]/g, '_');
        
        allReportData.push(reportData); // guardar para modo consolidado
        
          if (modoGeracao === 'individual') {
            addLog(`🔨 Compilando documento Word: ${reportData.SETOR}`);
            const reportBlob = await renderDocument(moldeBuffer, reportData);
            generatedFiles.push({ name: `${fileName}.docx`, blob: reportBlob });
          }
          setProcessingStatus(prev => prev.map((s, idx) => idx === i ? { ...s, status: 'completed' } : s));
        } catch (err: any) {
          setProcessingStatus(prev => prev.map((s, idx) => idx === i ? { ...s, status: 'error' } : s));
          addLog(`⚠️ Erro ao processar arquivo ${file.name}: ${err.message}`);
          console.error(`Erro no arquivo ${file.name}:`, err);
          // Continua para o próximo arquivo no modo lote
        }
      }

      if (modoGeracao === 'consolidado' && allReportData.length > 0) {
        // Agrupar por empresa e gerar 1 docx por empresa
        addLog('Agrupando setores por empresa...');

        const grupos = new Map<string, any[]>();
        for (const rd of allReportData) {
          const emp = rd.EMPRESA || 'SEM_EMPRESA';
          if (!grupos.has(emp)) grupos.set(emp, []);
          grupos.get(emp)!.push(rd);
        }

        const consolidadoFiles: { name: string; blob: Blob }[] = [];

        for (const [empresa, setores] of grupos) {
          addLog(`Compilando ${setores.length} setor(es): ${empresa}`);
          const first = setores[0];

          const totalFunc   = setores.reduce((s: number, x: any) => s + (x.FUNC_TOTAL   || 0), 0);
          const totalPartic = setores.reduce((s: number, x: any) => s + (x.PARTIC_TOTAL || 0), 0);
          const totalMasc   = setores.reduce((s: number, x: any) => s + (x.MASC_N       || 0), 0);
          const totalFem    = setores.reduce((s: number, x: any) => s + (x.FEM_N        || 0), 0);
          const percPartic  = totalFunc > 0 ? Math.round(totalPartic / totalFunc * 100) + '%' : '0%';
          const percMasc    = (totalMasc + totalFem) > 0 ? Math.round(totalMasc / (totalMasc + totalFem) * 100) + '%' : '0%';
          const percFem     = (totalMasc + totalFem) > 0 ? Math.round(totalFem  / (totalMasc + totalFem) * 100) + '%' : '0%';

        const consolidadoData = {
          EMPRESA:          first.EMPRESA,
          CNPJ:             first.CNPJ,
          UNIDADE:          first.UNIDADE,
          SETORES_LISTA:    setores.map((s: any) => s.SETOR).join(' / '),
          SETOR:            setores.map((s: any) => s.SETOR).join(' / '),
          DATA:             first.DATA,
          DATA_EXTENSA:     first.DATA_EXTENSA,
          MES_ANO_CAPA:     first.MES_ANO_CAPA,
          AVALIADOR:        first.AVALIADOR,
          TOTAL_PAGINAS:    first.TOTAL_PAGINAS,
          TEM_PLANO:        first.TEM_PLANO,
          FUNC_TOTAL:       totalFunc,
          PARTIC_TOTAL:     totalPartic,
          MASC_N:           totalMasc,
          FEM_N:            totalFem,
          PERC_PARTIC:      percPartic,
          SOMA_P16:         percPartic,
          PERC_EFETIVOS:    percPartic,
          PERC_MASC:        percMasc,
          PERC_FEM:         percFem,
          EXP_INTRINSECA:   first.EXP_INTRINSECA,
          EXP_SOBRECARGA:   first.EXP_SOBRECARGA,
          CONCLUSOES_LISTA: first.CONCLUSOES_LISTA,
          setores:          setores.map(s => ({
            ...s,
            // Ensure all possible tag variations are included for maximum compatibility
            headerImage:  s.INFO_HEADER || s.headerImage,
            INFO_HEADER:  s.INFO_HEADER || s.headerImage,
            radarImage:   s.GRAFICO || s.radarImage || s.RADAR,
            GRAFICO:      s.GRAFICO || s.radarImage || s.RADAR,
            RADAR:        s.GRAFICO || s.radarImage || s.RADAR,
            matrixImage:  s.MATRIZ_RISCO || s.matrixImage,
            MATRIZ_RISCO: s.MATRIZ_RISCO || s.matrixImage,
            criticosImage: s.TABELA_CRITICOS || s.criticosImage,
            TABELA_CRITICOS: s.TABELA_CRITICOS || s.criticosImage
          })),
          // Also providing uppercase version of the array for versatility
          SETORES: setores.map(s => ({
            ...s,
            headerImage:  s.INFO_HEADER || s.headerImage,
            INFO_HEADER:  s.INFO_HEADER || s.headerImage,
            radarImage:   s.GRAFICO || s.radarImage || s.RADAR,
            GRAFICO:      s.GRAFICO || s.radarImage || s.RADAR,
            RADAR:        s.GRAFICO || s.radarImage || s.RADAR,
            matrixImage:  s.MATRIZ_RISCO || s.matrixImage,
            MATRIZ_RISCO: s.MATRIZ_RISCO || s.matrixImage,
            criticosImage: s.TABELA_CRITICOS || s.criticosImage,
            TABELA_CRITICOS: s.TABELA_CRITICOS || s.criticosImage
          }))
        };

        // Generate Consolidated Header Image (Aggregated)
        const consolidatedHeaderB64 = await generateHeaderBase64({
          EMPRESA: consolidadoData.EMPRESA,
          UNIDADE: consolidadoData.UNIDADE,
          SETOR: "CONSOLIDADO / TODOS OS SETORES",
          DATA: consolidadoData.DATA,
          AVALIADOR: consolidadoData.AVALIADOR,
          FUNC_TOTAL: consolidadoData.FUNC_TOTAL,
          PARTIC_TOTAL: consolidadoData.PARTIC_TOTAL,
          MASC_N: consolidadoData.MASC_N,
          FEM_N: consolidadoData.FEM_N,
          PERC_PARTIC: consolidadoData.PERC_PARTIC
        });

        Object.assign(consolidadoData, {
          headerImage: consolidatedHeaderB64,
          INFO_HEADER: consolidatedHeaderB64
        });

        const blob = await renderDocument(moldeBuffer, consolidadoData);
          const d = new Date();
          consolidadoFiles.push({
            name: `Relatório_Consolidado_${empresa}_${d.getDate().toString().padStart(2,'0')}_${(d.getMonth()+1).toString().padStart(2,'0')}_${d.getFullYear()}.docx`.replace(/[\/\\?%*:|"<>]/g, '_'),
            blob
          });
        }

        if (consolidadoFiles.length === 1) {
          saveAs(consolidadoFiles[0].blob, consolidadoFiles[0].name);
          addLog(`✅ Relatório consolidado salvo: ${consolidadoFiles[0].name}`);
        } else {
          const zip = new JSZip();
          consolidadoFiles.forEach(f => zip.file(f.name, f.blob));
          const zipBlob = await zip.generateAsync({ type: 'blob' });
          saveAs(zipBlob, `Relatorios_Consolidados_${new Date().getTime()}.zip`);
          addLog(`✅ ${consolidadoFiles.length} relatórios consolidados salvos em ZIP.`);
        }

      } else if (generatedFiles.length > 0) {
        if (generatedFiles.length === 1) {
          saveAs(generatedFiles[0].blob, generatedFiles[0].name);
          addLog(`✅ Relatório salvo: ${generatedFiles[0].name}`);
        } else {
          addLog(`📦 Empacotando ${generatedFiles.length} arquivos em um ZIP...`);
          const zip = new JSZip();
          generatedFiles.forEach(f => zip.file(f.name, f.blob));
          const zipBlob = await zip.generateAsync({ type: 'blob' });
          saveAs(zipBlob, `Relatorios_Psicossociais_Lote_${new Date().getTime()}.zip`);
          addLog(`✅ Pacote ZIP salvo com sucesso.`);
        }
      } else if (allReportData.length === 0) {
        addLog('⚠️ Nenhum dado válido foi extraído das planilhas.');
      }

      addLog('🏁 Processamento concluído.');
    } catch (err: any) {
      addLog(`❌ ERRO CRÍTICO: ${err.message}`);
      console.error(err);
    } finally {
      setIsProcessing(false);
    }
  };

  const renderDocument = async (moldeBuffer: ArrayBuffer, dataObj: any): Promise<Blob> => {
    const zip = new PizZip(moldeBuffer);
    
    const imageOptions = {
      centered: true,
      getImage: (tagValue: string) => {
        const binaryString = window.atob(tagValue);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) bytes[i] = binaryString.charCodeAt(i);
        return bytes.buffer;
      },
      getSize: (img: any, tagValue: string, tagName: string) => {
        // INFO_HEADER: ~17cm x 3.2cm
        if (tagName === 'INFO_HEADER' || tagName === 'headerImage') return [650, 122]; 
        // TABELA_CRITICOS: ~14.4cm x 9.7cm (from screenshot 14.36 x 9.72)
        if (tagName === 'TABELA_CRITICOS' || tagName === 'criticosImage') return [543, 367]; 
        // MATRIZ_RISCO: ~14.0cm x 6.5cm (from screenshot 13.99 x 6.54)
        if (tagName === 'MATRIZ_RISCO' || tagName === 'matrixImage') return [529, 247];
        // RADAR: Standard chart size
        if (tagName === 'chartImage' || tagName === 'RADAR' || tagName === 'GRAFICO' || tagName === 'radarImage') return [652, 473]; 
        // Exposure Tables
        if (tagName === 'TABELA_RESUMO_EXPOSICAO' || tagName === 'TABELA_EXPOSICAO' || tagName === 'TABELA_RESUMO' || tagName === 'IMAGEM_EXPOSICAO') return [652, 244]; 
        return [652, 500];
      }
    };

    const doc = new Docxtemplater(zip, {
      modules: [new ImageModule(imageOptions)],
      paragraphLoop: true,
      linebreaks: true,
      nullGetter: () => ""
    });

    try {
      doc.render(dataObj);
    } catch (error: any) {
      if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors.map((e: any) => e.message).join(' | ');
        throw new Error(`Erro de Tags no Word: ${errorMessages}`);
      }
      throw error;
    }
    
    addLog(`✅ Documento renderizado com sucesso.`);
    return doc.getZip().generate({ 
      type: "blob", 
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" 
    });
  };

  const removeExcel = (index: number) => {
    setExcels(prev => prev.filter((_, i) => i !== index));
    addLog('Arquivo Excel removido.');
  };

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-800 selection:bg-blue-100">
      
      <main className="max-w-5xl mx-auto py-12 px-6">
        
        {/* Header */}
        <header className="text-center mb-12">
          <motion.div 
            initial={{ scale: 0.8, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            className="inline-block p-4 bg-blue-600 rounded-3xl shadow-xl shadow-blue-200 mb-6"
          >
            <FileText className="w-10 h-10 text-white" />
          </motion.div>
          <h1 className="text-4xl font-extrabold text-slate-900 mb-3 tracking-tight">
            Gerador de Relatório Psicossocial
          </h1>
        </header>

        {/* Step settings */}
        <motion.div 
          initial={{ y: 20, opacity: 0 }}
          animate={{ y: 0, opacity: 1 }}
          className="flex justify-center mb-10"
        >
          <div 
            onClick={() => setIncludePlano(!includePlano)}
            className={`
              cursor-pointer group flex items-center gap-6 p-5 rounded-[2rem] border-2 transition-all w-full max-w-sm
              ${includePlano ? 'bg-blue-600 border-blue-600 shadow-lg shadow-blue-200' : 'bg-white border-slate-200 hover:border-blue-300'}
            `}
          >
            <div className={`p-4 rounded-2xl transition-colors ${includePlano ? 'bg-white/20' : 'bg-blue-50'}`}>
              <CheckCircle2 className={`w-8 h-8 ${includePlano ? 'text-white' : 'text-blue-600'}`} />
            </div>
            <div className="flex-1">
              <h3 className={`font-black text-lg ${includePlano ? 'text-white' : 'text-slate-800'}`}>Plano de Ação</h3>
              <p className={`text-xs ${includePlano ? 'text-blue-100' : 'text-slate-500'}`}>Incluir recomendações técnicas automáticas no relatório.</p>
            </div>
            <div className={`
              w-12 h-6 rounded-full relative transition-colors border-2
              ${includePlano ? 'bg-blue-400 border-white/30' : 'bg-slate-200 border-transparent'}
            `}>
              <motion.div 
                animate={{ x: includePlano ? 24 : 2 }}
                className="absolute top-0.5 w-4 h-4 rounded-full bg-white shadow-sm"
              />
            </div>
          </div>
        </motion.div>

        {/* Modo de Geração */}
        <motion.div
          initial={{ y: 20, opacity: 0 }}
          animate={{ y: 0, opacity: 1 }}
          className="grid grid-cols-2 gap-4 max-w-2xl mx-auto mb-10"
        >
          {([
            {
              id: 'individual',
              titulo: 'INDIVIDUAL',
              subtitulo: '1 Excel = 1 Relatório',
              descricao: 'Cada arquivo Excel gera seu próprio documento Word independente.'
            },
            {
              id: 'consolidado',
              titulo: 'CONSOLIDADO',
              subtitulo: 'Vários setores = 1 relatório por empresa',
              descricao: 'Agrupa todos os setores da mesma empresa em um único .docx, um abaixo do outro.'
            }
          ] as const).map(modo => (
            <button
              key={modo.id}
              onClick={() => setModoGeracao(modo.id)}
              className={`p-5 rounded-2xl border-2 text-left transition-all ${
                modoGeracao === modo.id
                  ? 'border-blue-500 bg-blue-50 shadow-md'
                  : 'border-slate-200 bg-white hover:border-slate-300'
              }`}
            >
              <div className="flex items-center gap-3 mb-2">
                <div className={`w-5 h-5 rounded-full border-2 flex items-center justify-center flex-shrink-0 ${modoGeracao === modo.id ? 'border-blue-500' : 'border-slate-300'}`}>
                  {modoGeracao === modo.id && <div className="w-2.5 h-2.5 rounded-full bg-blue-500" />}
                </div>
                <span className={`font-black text-sm ${modoGeracao === modo.id ? 'text-blue-700' : 'text-slate-700'}`}>{modo.titulo}</span>
              </div>
              <p className={`text-xs font-bold mb-1 ${modoGeracao === modo.id ? 'text-blue-600' : 'text-slate-500'}`}>{modo.subtitulo}</p>
              <p className="text-xs text-slate-500 leading-relaxed">{modo.descricao}</p>
            </button>
          ))}
        </motion.div>

        {/* Upload Grid */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 mb-12">
          
          {/* Step 1: Word Template */}
          <section className="bg-white p-8 rounded-[2rem] shadow-sm border border-slate-200 group transition-all">
            <div className="flex items-center justify-between mb-6">
              <h3 className="font-bold text-xl flex items-center text-blue-700">
                <span className="bg-blue-100 w-8 h-8 rounded-full flex items-center justify-center mr-3 text-sm font-black">1</span>
                Molde Word (.docx)
              </h3>
              {molde && <CheckCircle2 className="text-emerald-500 w-6 h-6" />}
            </div>
            
            <div 
              className={`
                relative h-48 border-2 border-dashed rounded-3xl flex flex-col items-center justify-center transition-all cursor-pointer
                ${dragActive['molde'] ? 'border-blue-500 bg-blue-50 scale-[0.98]' : 'border-slate-200 hover:border-blue-300 hover:bg-slate-50'}
              `}
              onDragEnter={(e) => handleDrag(e, 'molde', true)}
              onDragLeave={(e) => handleDrag(e, 'molde', false)}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => onDrop(e, 'molde')}
              onClick={() => document.getElementById('molde-input')?.click()}
            >
              <input 
                id="molde-input"
                type="file" 
                className="hidden" 
                accept=".docx" 
                onChange={(e) => handleFiles(e.target.files, 'molde')}
              />
              <FileText className={`w-10 h-10 mb-3 ${molde ? 'text-blue-600' : 'text-slate-300'}`} />
              <div className="text-center px-4">
                <p className={`font-semibold ${molde ? 'text-blue-700' : 'text-slate-500'}`}>
                  {molde ? molde.name : 'Arraste o arquivo molde aqui'}
                </p>
                {!molde && <p className="text-xs text-slate-400 mt-1">Clique para selecionar</p>}
              </div>
            </div>
          </section>

          {/* Step 2: Excels */}
          <section className="bg-white p-8 rounded-[2rem] shadow-sm border border-slate-200 group transition-all">
             <div className="flex items-center justify-between mb-6">
              <h3 className="font-bold text-xl flex items-center text-emerald-700">
                <span className="bg-emerald-100 w-8 h-8 rounded-full flex items-center justify-center mr-3 text-sm font-black">2</span>
                Excels (.xlsx)
              </h3>
              {excels.length > 0 && <span className="bg-emerald-100 text-emerald-700 px-3 py-1 rounded-full text-xs font-bold">{excels.length}</span>}
            </div>

            <div 
              className={`
                relative h-48 border-2 border-dashed rounded-3xl flex flex-col items-center justify-center transition-all cursor-pointer
                ${dragActive['excels'] ? 'border-emerald-500 bg-emerald-50 scale-[0.98]' : 'border-slate-200 hover:border-emerald-300 hover:bg-slate-50'}
              `}
              onDragEnter={(e) => handleDrag(e, 'excels', true)}
              onDragLeave={(e) => handleDrag(e, 'excels', false)}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => onDrop(e, 'excels')}
              onClick={() => document.getElementById('excel-input')?.click()}
            >
              <input 
                id="excel-input"
                type="file" 
                className="hidden" 
                accept=".xlsx" 
                multiple
                onChange={(e) => handleFiles(e.target.files, 'excels')}
              />
              <FileSpreadsheet className={`w-10 h-10 mb-3 ${excels.length > 0 ? 'text-emerald-600' : 'text-slate-300'}`} />
              <div className="text-center px-4">
                <p className={`font-semibold ${excels.length > 0 ? 'text-emerald-700' : 'text-slate-500'}`}>
                  {excels.length > 0 ? `${excels.length} arquivos prontos` : 'Selecione um ou vários Excels'}
                </p>
                {!excels.length && <p className="text-xs text-slate-400 mt-1">Os dados serão extraídos automaticamente</p>}
              </div>
            </div>

            {/* List of uploaded Excels */}
            {excels.length > 0 && (
              <div className="mt-4 space-y-2 max-h-40 overflow-y-auto pr-2 custom-scrollbar">
                {excels.map((file, idx) => (
                  <div key={idx} className="flex items-center justify-between bg-slate-50 p-2 rounded-xl border border-slate-100 group">
                    <div className="flex items-center gap-2 overflow-hidden">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-500 flex-shrink-0" />
                      <span className="text-xs font-medium truncate text-slate-600">{file.name}</span>
                    </div>
                    <button 
                      onClick={(e) => { e.stopPropagation(); removeExcel(idx); }}
                      className="text-[10px] font-bold text-slate-400 hover:text-rose-500 transition-colors uppercase px-2"
                    >
                      Remover
                    </button>
                  </div>
                ))}
              </div>
            )}
          </section>
        </div>

        {/* Generate Button Container */}
        <div className="relative group">
          <div className="absolute -inset-1 bg-gradient-to-r from-blue-600 to-emerald-500 rounded-[2.5rem] blur opacity-25 group-hover:opacity-50 transition-opacity duration-1000"></div>
          <div className="relative bg-white p-8 rounded-[2rem] shadow-xl border border-slate-100 flex flex-col items-center">
            <button 
              onClick={generateReports}
              disabled={!molde || excels.length === 0 || isProcessing}
              className={`
                w-full md:w-auto px-16 py-5 rounded-2xl font-black text-lg transition-all transform active:scale-95
                flex items-center justify-center gap-3 shadow-lg hover:shadow-blue-200
                ${(!molde || excels.length === 0) 
                  ? 'bg-slate-100 text-slate-400 cursor-not-allowed' 
                  : 'bg-blue-600 hover:bg-blue-700 text-white hover:-translate-y-1'}
              `}
            >
              {isProcessing ? (
                <>
                  <Loader2 className="animate-spin w-6 h-6" />
                  PROCESSANDO...
                </>
              ) : (
                <>
                  {modoGeracao === 'consolidado' ? 'GERAR RELATÓRIO CONSOLIDADO' : 'GERAR RELATÓRIOS INDIVIDUAIS'}
                  <ArrowRight className="w-6 h-6" />
                </>
              )}
            </button>
            <div className="mt-3 flex items-center gap-2">
              <div className={`w-2 h-2 rounded-full ${modoGeracao === 'consolidado' ? 'bg-amber-500 animate-pulse' : 'bg-emerald-500'}`}></div>
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-tighter">
                Modo selecionado: {modoGeracao === 'consolidado' ? 'Consolidação por Empresa' : 'Arquivos Individuais'}
              </span>
            </div>
            <AnimatePresence>
              {isProcessing && (
                <motion.div 
                  initial={{ height: 0, opacity: 0 }}
                  animate={{ height: 'auto', opacity: 1 }}
                  exit={{ height: 0, opacity: 0 }}
                  className="mt-8 overflow-hidden"
                >
                  {/* Dashboard Header - Minimalist & High Impact */}
                  <div className="bg-white rounded-3xl border border-slate-200 shadow-xl shadow-slate-200/50 p-6">
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
                      <div className="flex items-center gap-5">
                        <div className="relative">
                          <div className="w-16 h-16 bg-blue-600 rounded-2xl flex items-center justify-center shadow-lg shadow-blue-200 rotate-3">
                            <FileText className="w-8 h-8 text-white -rotate-3" />
                          </div>
                          <div className="absolute -bottom-1 -right-1 w-6 h-6 bg-emerald-500 rounded-lg border-2 border-white flex items-center justify-center">
                            <span className="text-[10px] font-black text-white">
                              {processingStatus.filter(s => s.status === 'completed').length}
                            </span>
                          </div>
                        </div>
                        <div>
                        <div className="flex items-center gap-2 mb-1">
                          <span className="px-2 py-0.5 bg-blue-50 text-blue-600 text-[9px] font-black uppercase tracking-widest rounded-md border border-blue-100 italic">
                            Live Engine
                          </span>
                        </div>
                        <h4 className="text-xl font-black text-slate-800 tracking-tight">
                          Gerando <span className="text-blue-600">
                            {processingStatus.filter(s => s.status === 'completed').length}
                          </span> de {processingStatus.length}
                        </h4>
                        <div className="flex items-center gap-2">
                          <Loader2 className="w-3 h-3 text-blue-500 animate-spin" />
                          <p className="text-xs text-slate-400 font-medium truncate max-w-[180px]">
                            {processingStatus.find(s => s.status === 'processing')?.name || "Processamento em Lote..."}
                          </p>
                        </div>
                      </div>
                    </div>

                    <div className="flex-1 max-w-[280px]">
                      <div className="flex justify-between items-end mb-2">
                        <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest">Progresso Total</span>
                        <span className="text-xs font-black text-blue-600 bg-blue-50 px-2 py-0.5 rounded-lg border border-blue-100">
                          {Math.round((processingStatus.filter(s => s.status === 'completed').length / processingStatus.length) * 100)}%
                        </span>
                      </div>
                      <div className="w-full h-3 bg-slate-100 rounded-full overflow-hidden p-0.5 border border-slate-50">
                        <motion.div 
                          initial={{ width: 0 }}
                          animate={{ 
                            width: `${(processingStatus.filter(s => s.status === 'completed').length / processingStatus.length) * 100}%` 
                          }}
                          className="h-full bg-gradient-to-r from-blue-600 to-indigo-600 rounded-full shadow-sm"
                        />
                      </div>
                    </div>
                  </div>

                  {/* High-End Grid View - Compact Particles */}
                  <div className="mt-8 flex flex-wrap gap-2 justify-center">
                    {processingStatus.map((item, idx) => (
                      <motion.div 
                        key={idx}
                        title={item.name}
                        initial={{ scale: 0.8, opacity: 0 }}
                        animate={{ scale: 1, opacity: 1 }}
                        transition={{ delay: Math.min(idx * 0.02, 0.5) }}
                        className={`
                          w-3 h-3 rounded-md transition-all duration-700
                          ${item.status === 'completed' ? 'bg-emerald-500 shadow-[0_0_8px_rgba(16,185,129,0.3)]' : 
                            item.status === 'processing' ? 'bg-blue-600 animate-pulse shadow-[0_0_12px_rgba(37,99,235,0.5)] scale-125 z-10' : 
                            item.status === 'error' ? 'bg-rose-500 shadow-[0_0_8px_rgba(244,63,94,0.3)]' : 
                            'bg-slate-200 border border-slate-300 opacity-40'}
                        `}
                      />
                    ))}
                  </div>
                </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </div>

        {/* Terminal Log */}
        <section className="mt-12 bg-slate-900 rounded-3xl p-8 text-emerald-400 font-mono shadow-2xl overflow-hidden border border-slate-800">
          <div className="flex items-center gap-3 mb-6 border-b border-slate-800 pb-4">
            <div className="flex gap-1.5">
              <div className="w-3 h-3 rounded-full bg-[#ff5f56]"></div>
              <div className="w-3 h-3 rounded-full bg-[#ffbd2e]"></div>
              <div className="w-3 h-3 rounded-full bg-[#27c93f]"></div>
            </div>
            <div className="flex items-center gap-2 text-slate-500 text-xs ml-4">
              <Terminal className="w-3 h-3" />
              <span>TERMINAL_INSAT_v2.0</span>
            </div>
            <button 
              onClick={clearLogs}
              className="ml-auto text-[10px] bg-slate-800 hover:bg-slate-700 text-slate-400 px-2 py-1 rounded transition-colors border border-slate-700"
            >
              LIMPAR TERMINAL
            </button>
          </div>
          
          <div className="max-h-64 overflow-y-auto scrollbar-thin scrollbar-thumb-slate-700 pr-2">
            <div className="space-y-1 text-sm">
              <p className="text-slate-500 opacity-70">[{new Date().toLocaleTimeString()}] Sistema inicializado. Aguardando arquivos...</p>
              {logs.length === 0 && (
                <p className="italic text-slate-600">Siga os passos 1 e 2 para começar.</p>
              )}
              {logs.map((log) => (
                <motion.div 
                  key={log.id} 
                  initial={{ x: -10, opacity: 0 }}
                  animate={{ x: 0, opacity: 1 }}
                  className="flex gap-4 group"
                >
                  <span className="text-slate-600 select-none group-hover:text-slate-400 transition-colors">[{log.timestamp}]</span>
                  <span className={log.message.startsWith('✅') ? 'text-emerald-300' : log.message.startsWith('⚠️') ? 'text-amber-300' : log.message.startsWith('🚀') ? 'text-blue-300 font-bold' : 'text-emerald-400/90'}>
                    {log.message}
                  </span>
                </motion.div>
              ))}
            </div>
          </div>
        </section>

        <footer className="mt-12 text-center text-slate-400 text-xs font-medium uppercase tracking-widest">
          Insat Web Engine &copy; {new Date().getFullYear()} • Edição Profissional
        </footer>


      </main>

      {/* Hidden Canvas Manager */}
      <div className="fixed -left-[10000px] -top-[10000px]" aria-hidden="true">
        <canvas ref={canvasRef} width="1200" height="1200" />
      </div>

  </div>
);
}
