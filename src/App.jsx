import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Html5QrcodeScanner } from 'html5-qrcode';
import { 
  Upload, Search, Mail, RefreshCw, CheckCircle2, 
  FileText, AlertCircle, Camera, X 
} from 'lucide-react';

/* =========================
   🔥 NORMALIZAÇÃO FORTE
========================= */
const normalizeKey = (key) =>
  String(key ?? '')
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ');

const normalizeData = (data) =>
  data.map(row => {
    const newRow = {};
    Object.keys(row).forEach(k => {
      newRow[normalizeKey(k)] = row[k];
    });
    return newRow;
  });

const clean = (v) =>
  String(v ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');

/* =========================
   🔥 COLUMN MAP DEFINITIVO
========================= */
const COLUMN_MAP = {
  ZSD036: {
    SZ: ['tappi number', 'tappi number '],
    LOTE: ['lote']
  },
  MB51: {
    LOTE: ['lote'],
    PESO: ['quantidade'],
    TRANSACAO: ['tipo de movimento'],
    DATA: ['data de lançamento'],
    USUARIO: ['nome do usuario', 'nome do usuário'],
    CABECALHO: ['texto cabeçalho documento']
  }
};

/* =========================
   🔥 PEGADOR SEGURO DE COLUNA
========================= */
const getValue = (row, keys) => {
  for (let k of keys) {
    if (row[k] !== undefined && row[k] !== '') return row[k];
  }
  return '';
};

export default function SAPBatchTracker() {
  const [dataZSD036, setDataZSD036] = useState([]);
  const [dataMB51, setDataMB51] = useState([]);
  const [szInput, setSzInput] = useState('');
  const [result, setResult] = useState(null);
  const [toast, setToast] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const scannerRef = useRef(null);

  const showToast = (msg, type = 'success') => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  };

  /* =========================
     📂 UPLOAD EXCEL
  ========================= */
  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });

        let data = XLSX.utils.sheet_to_json(
          wb.Sheets[wb.SheetNames[0]],
          { defval: '' }
        );

        data = normalizeData(data);

        if (type === 'ZSD036') setDataZSD036(data);
        else setDataMB51(data);

        showToast('Planilha carregada com sucesso');
      } catch (err) {
        showToast('Erro ao ler Excel', 'error');
      }
    };

    reader.readAsBinaryString(file);
  };

  /* =========================
     🔎 CONSULTA
  ========================= */
  const executeTrack = (value) => {
    const searchVal = clean(value || szInput);

    if (!searchVal) return showToast('Digite a SZ', 'error');
    if (!dataZSD036.length || !dataMB51.length)
      return showToast('Carregue as planilhas', 'error');

    setIsProcessing(true);

    try {
      /* 🔥 ZSD036 (SZ = Tappi Number) */
      const szRecord = dataZSD036.find(r =>
        clean(getValue(r, COLUMN_MAP.ZSD036.SZ)) === searchVal
      );

      if (!szRecord) throw new Error('SZ não encontrada na ZSD036');

      const batchId = clean(getValue(szRecord, COLUMN_MAP.ZSD036.LOTE));

      /* 🔥 MB51 */
      const movements = dataMB51.filter(r =>
        clean(getValue(r, COLUMN_MAP.MB51.LOTE)) === batchId
      );

      if (!movements.length) throw new Error('Lote não encontrado na MB51');

      const sorted = movements.sort(
        (a, b) =>
          new Date(getValue(b, COLUMN_MAP.MB51.DATA)) -
          new Date(getValue(a, COLUMN_MAP.MB51.DATA))
      );

      const latest = sorted[0];

      setResult({
        sz: searchVal,
        lote: batchId,
        peso: getValue(latest, COLUMN_MAP.MB51.PESO),
        transacao: getValue(latest, COLUMN_MAP.MB51.TRANSACAO),
        data: getValue(latest, COLUMN_MAP.MB51.DATA),
        usuario: getValue(latest, COLUMN_MAP.MB51.USUARIO),
        cabecalho: getValue(latest, COLUMN_MAP.MB51.CABECALHO)
      });

    } catch (err) {
      showToast(err.message, 'error');
    }

    setIsProcessing(false);
  };

  /* =========================
     🎯 UI (NÃO ALTERADO)
  ========================= */
  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 p-4 md:p-8 font-sans">

      <header className="max-w-6xl mx-auto mb-8 bg-white p-6 rounded-xl shadow-sm border-b-4 border-blue-600 flex justify-between items-center">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg text-white">
            <Search size={24}/>
          </div>
          <h1 className="text-2xl font-bold text-slate-700">
            Consulta aí <span className="text-blue-600">Basílio</span>
          </h1>
        </div>
        <span className="text-xs text-slate-400">SuzanLOKO V1.</span>
      </header>

      <main className="max-w-6xl mx-auto grid grid-cols-1 lg:grid-cols-12 gap-6">

        <div className="lg:col-span-4 space-y-6">

          <section className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="text-xs font-bold text-slate-500 uppercase mb-4 flex items-center gap-2">
              <FileText size={14}/> Dados SAP
            </h2>

            {['ZSD036','MB51'].map(type => (
              <div key={type}>
                <label className="block text-xs font-bold mb-1">
                  Planilha {type}
                </label>

                <div
                  onClick={() => document.getElementById(`f-${type}`).click()}
                  className="p-4 border-2 border-dashed rounded-lg cursor-pointer"
                >
                  <Upload size={20}/>
                  <span className="text-xs block text-center">
                    Clique para upload
                  </span>

                  <input
                    id={`f-${type}`}
                    type="file"
                    hidden
                    onChange={(e) => handleFileUpload(e, type)}
                  />
                </div>
              </div>
            ))}
          </section>

          <section className="bg-white p-6 rounded-xl shadow-sm border">
            <h2 className="text-xs font-bold mb-4">Rastreamento</h2>

            <div className="flex gap-2">
              <input
                className="flex-1 p-3 border rounded-lg"
                value={szInput}
                onChange={e => setSzInput(e.target.value)}
                onKeyDown={e => e.key === 'Enter' && executeTrack()}
              />

              <button
                onClick={() => executeTrack()}
                className="bg-blue-600 text-white px-4 rounded-lg"
              >
                {isProcessing ? <RefreshCw className="animate-spin"/> : 'OK'}
              </button>
            </div>
          </section>

        </div>

        <div className="lg:col-span-8 bg-white p-6 rounded-xl">
          {!result ? (
            <p>Aguardando consulta...</p>
          ) : (
            <>
              <p>SZ: {result.sz}</p>
              <p>Lote: {result.lote}</p>
              <p>Peso: {result.peso}</p>
              <p>Mov: {result.transacao}</p>
              <p>Data: {result.data}</p>

              <button
                onClick={() => {
                  setResult(null);
                  setSzInput('');
                }}
                className="mt-4 text-blue-600"
              >
                Nova Consulta
              </button>
            </>
          )}
        </div>

      </main>

      {toast && (
        <div className="fixed bottom-5 right-5 bg-green-600 text-white p-3 rounded-lg">
          {toast.msg}
        </div>
      )}

    </div>
  );
}
