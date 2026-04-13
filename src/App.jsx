import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Html5QrcodeScanner } from 'html5-qrcode';
import { 
  Upload, Search, Mail, RefreshCw, CheckCircle2, 
  FileText, AlertCircle, Camera, X 
} from 'lucide-react';

// 🔥 NORMALIZAÇÃO (não muda layout)
const normalizeKeys = (data) => {
  return data.map(row => {
    const newRow = {};
    Object.keys(row).forEach(key => {
      const cleanKey = key.trim().toLowerCase();
      newRow[cleanKey] = row[key];
    });
    return newRow;
  });
};

// 🔥 AJUSTADO PARA SUAS PLANILHAS
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
  const [email, setEmail] = useState('');
  const [saveEmail, setSaveEmail] = useState(false);

  const [isScanning, setIsScanning] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [toast, setToast] = useState(null);
  const [uploads, setUploads] = useState({
    ZSD036: { progress: 0, status: 'idle', fileName: '' },
    MB51: { progress: 0, status: 'idle', fileName: '' },
  });

  const scannerRef = useRef(null);

  useEffect(() => {
    const storedEmail = localStorage.getItem('sap_tracker_email');
    if (storedEmail) {
      setEmail(storedEmail);
      setSaveEmail(true);
    }
  }, []);

  const showToast = (msg, type = 'success') => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 4000);
  };

  useEffect(() => {
    if (isScanning) {
      const timeout = setTimeout(() => {
        const scanner = new Html5QrcodeScanner("reader", { 
          fps: 10, 
          qrbox: { width: 250, height: 150 },
          aspectRatio: 1.0 
        });

        scanner.render((decodedText) => {
          setSzInput(decodedText);
          setIsScanning(false);
          scanner.clear();
          showToast("Código lido com sucesso!");
          setTimeout(() => executeTrack(decodedText), 500);
        }, () => {});

        scannerRef.current = scanner;
      }, 100);

      return () => {
        clearTimeout(timeout);
        if (scannerRef.current) scannerRef.current.clear();
      };
    }
  }, [isScanning]);

  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;

    setUploads(prev => ({ ...prev, [type]: { ...prev[type], status: 'loading', fileName: file.name, progress: 0 } }));

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });

        let data = XLSX.utils.sheet_to_json(
          wb.Sheets[wb.SheetNames[0]],
          { defval: '' }
        );

        data = normalizeKeys(data);

        if (type === 'ZSD036') setDataZSD036(data);
        else setDataMB51(data);

        setUploads(prev => ({ ...prev, [type]: { ...prev[type], status: 'complete', progress: 100 } }));
      } catch (err) {
        showToast("Erro ao ler Excel. Verifique o formato.", "error");
        setUploads(prev => ({ ...prev, [type]: { ...prev[type], status: 'error' } }));
      }
    };

    reader.readAsBinaryString(file);
  };

  const executeTrack = (value) => {
    const searchVal = value || szInput;

    if (!searchVal) return showToast("Digite ou escaneie a SZ", "error");

    if (dataZSD036.length === 0 || dataMB51.length === 0) {
      return showToast("Carregue as duas planilhas!", "error");
    }

    setIsProcessing(true);

    try {
      const szRecord = dataZSD036.find(row => 
        String(row[COLUMN_MAP.ZSD036.SZ] || '')
          .trim()
          .toLowerCase() === searchVal.trim().toLowerCase()
      );

      if (!szRecord) throw new Error("SZ não encontrada na ZSD036");

      const batchId = szRecord[COLUMN_MAP.ZSD036.LOTE];

      const movements = dataMB51.filter(row => 
        String(row[COLUMN_MAP.MB51.LOTE] || '').trim() === String(batchId || '').trim()
      );

      if (!movements.length) throw new Error("Lote não encontrado na MB51");

      const sorted = movements.sort((a, b) => {
        const d1 = new Date(a[COLUMN_MAP.MB51.DATA]);
        const d2 = new Date(b[COLUMN_MAP.MB51.DATA]);
        return d2 - d1;
      });

      const latest = sorted[0];

      setResult({
        sz: searchVal,
        lote: batchId,
        peso: latest[COLUMN_MAP.MB51.PESO],
        transacao: latest[COLUMN_MAP.MB51.TRANSACAO],
        data: latest[COLUMN_MAP.MB51.DATA],
        hora: '', // não existe na planilha
        usuario: latest[COLUMN_MAP.MB51.USUARIO],
        cabecalho: latest[COLUMN_MAP.MB51.CABECALHO]
      });

    } catch (err) {
      showToast(err.message, "error");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 p-4 md:p-8 font-sans">
      <header className="max-w-6xl mx-auto mb-8 bg-white p-6 rounded-xl shadow-sm border-b-4 border-blue-600 flex justify-between items-center">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg text-white"><Search size={24}/></div>
          <h1 className="text-2xl font-bold text-slate-700">Consulta aí <span className="text-blue-600">Basílio</span></h1>
        </div>
        <span className="text-xs text-slate-400 font-medium uppercase">SuzanLOKO V1.</span>
      </header>

      {/* RESTANTE DO JSX PERMANECE IGUAL */}
    </div>
  );
}

function Detail({ label, value, highlight = false }) {
  return (
    <div className="flex justify-between border-b pb-2">
      <span className="text-slate-500 text-sm">{label}:</span>
      <span className={`font-bold ${highlight ? 'text-blue-600' : 'text-slate-800'}`}>{value || '---'}</span>
    </div>
  );
}
