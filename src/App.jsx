import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Html5QrcodeScanner } from 'html5-qrcode';
import { 
  Upload, Search, Mail, RefreshCw, CheckCircle2, 
  FileText, AlertCircle, Camera, X 
} from 'lucide-react';

// 🔥 NORMALIZAÇÃO REAL (resolve seu problema)
const normalizeKey = (key) =>
  String(key)
    .trim()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');

const normalizeData = (data) =>
  data.map(row => {
    const newRow = {};
    Object.keys(row).forEach(k => {
      newRow[normalizeKey(k)] = row[k];
    });
    return newRow;
  });

// 🔥 LIMPEZA DE VALOR (CRÍTICO PRA SZ)
const clean = (v) =>
  String(v ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');

// 🔥 DATA BR SEGURA (resolve bug invisível)
const parseDate = (val) => {
  if (!val) return 0;

  if (typeof val === 'string' && val.includes('/')) {
    const [d, m, y] = val.split('/');
    return new Date(`${y}-${m}-${d}`).getTime();
  }

  return new Date(val).getTime();
};

const COLUMN_MAP = {
  ZSD036: {
    SZ: 'tappi number',
    LOTE: 'lote'
  },
  MB51: {
    LOTE: 'lote',
    PESO: 'quantidade',
    TRANSACAO: 'tipo de movimento',
    DATA: 'data de lançamento',
    USUARIO: 'nome do usuário',
    CABECALHO: 'texto cabeçalho documento'
  }
};

export default function SAPBatchTracker() {
  const [dataZSD036, setDataZSD036] = useState([]);
  const [dataMB51, setDataMB51] = useState([]);
  const [szInput, setSzInput] = useState('');
  const [result, setResult] = useState(null);
  const [toast, setToast] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);

  const showToast = (msg, type='success') => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 3500);
  };

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

        showToast('Planilha carregada!');
      } catch {
        showToast('Erro ao ler planilha', 'error');
      }
    };

    reader.readAsBinaryString(file);
  };

  const executeTrack = (value) => {
    const searchVal = clean(value || szInput);

    if (!searchVal) return showToast('Digite a SZ', 'error');
    if (!dataZSD036.length || !dataMB51.length)
      return showToast('Carregue as planilhas', 'error');

    setIsProcessing(true);

    try {
      const szRecord = dataZSD036.find(r =>
        clean(r['tappi number']) === searchVal
      );

      if (!szRecord) throw new Error('SZ não encontrada');

      const batchId = clean(szRecord['lote']);

      const movements = dataMB51.filter(r =>
        clean(r['lote']) === batchId
      );

      if (!movements.length) throw new Error('Lote não encontrado');

      const sorted = movements.sort(
        (a,b) => parseDate(b['data de lançamento']) - parseDate(a['data de lançamento'])
      );

      const latest = sorted[0];

      setResult({
        sz: searchVal,
        lote: batchId,
        peso: latest['quantidade'],
        transacao: latest['tipo de movimento'],
        data: latest['data de lançamento'],
        usuario: latest['nome do usuário'],
        cabecalho: latest['texto cabeçalho documento']
      });

    } catch (err) {
      showToast(err.message, 'error');
    }

    setIsProcessing(false);
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 p-4 md:p-8 font-sans">

      {/* HEADER (INTACTO) */}
      <header className="max-w-6xl mx-auto mb-8 bg-white p-6 rounded-xl shadow-sm border-b-4 border-blue-600 flex justify-between items-center">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg text-white"><Search size={24}/></div>
          <h1 className="text-2xl font-bold">Consulta aí <span className="text-blue-600">Basílio</span></h1>
        </div>
        <span className="text-xs text-slate-400">SuzanLOKO V1.</span>
      </header>

      {/* INPUT + BUTTON */}
      <div className="max-w-6xl mx-auto flex gap-2 mb-6">
        <input
          className="flex-1 p-3 border rounded-lg"
          value={szInput}
          onChange={e => setSzInput(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && executeTrack()}
        />

        <button
          onClick={() => executeTrack()}
          disabled={isProcessing}
          className="bg-blue-600 text-white px-4 rounded-lg"
        >
          {isProcessing ? <RefreshCw className="animate-spin"/> : 'Consultar'}
        </button>

        <button
          onClick={() => {
            setResult(null);
            setSzInput('');
          }}
          className="bg-slate-200 px-4 rounded-lg"
        >
          Nova
        </button>
      </div>

      {/* UPLOAD */}
      <div className="max-w-6xl mx-auto flex gap-4 mb-6">
        <input type="file" onChange={(e)=>handleFileUpload(e,'ZSD036')} />
        <input type="file" onChange={(e)=>handleFileUpload(e,'MB51')} />
      </div>

      {/* RESULTADO */}
      <div className="max-w-6xl mx-auto bg-white p-6 rounded-xl">
        {result ? (
          <>
            <p>SZ: {result.sz}</p>
            <p>Lote: {result.lote}</p>
            <p>Peso: {result.peso}</p>
            <p>Mov: {result.transacao}</p>
            <p>Data: {result.data}</p>
          </>
        ) : (
          <p>Sem consulta ainda</p>
        )}
      </div>

      {/* TOAST */}
      {toast && (
        <div className="fixed bottom-5 right-5 bg-green-600 text-white p-3 rounded-lg">
          {toast.msg}
        </div>
      )}
    </div>
  );
}
