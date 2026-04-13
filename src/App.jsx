import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Html5QrcodeScanner } from 'html5-qrcode';
import { 
  Upload, Search, Mail, RefreshCw, CheckCircle2, 
  FileText, AlertCircle, Camera, X 
} from 'lucide-react';

// 🔥 NORMALIZAÇÃO FORTE
const normalizeKeys = (data) => {
  return data.map(row => {
    const newRow = {};
    Object.keys(row).forEach(key => {
      const cleanKey = key.trim().toLowerCase();
      newRow[cleanKey] = String(row[key] ?? '').trim();
    });
    return newRow;
  });
};

// 🔥 LIMPEZA DE VALOR (CRÍTICO)
const clean = (val) => {
  return String(val ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '');
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

  const handleFileUpload = (e, type) => {
    const file = e.target.files[0];
    if (!file) return;

    setUploads(prev => ({ ...prev, [type]: { ...prev[type], status: 'loading', fileName: file.name } }));

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'binary' });

        let data = XLSX.utils.sheet_to_json(
          wb.Sheets[wb.SheetNames[0]],
          { defval: '' }
        );

        data = normalizeKeys(data);

        console.log(type, data.slice(0,3)); // debug

        if (type === 'ZSD036') setDataZSD036(data);
        else setDataMB51(data);

        setUploads(prev => ({ ...prev, [type]: { ...prev[type], status: 'complete' } }));
      } catch (err) {
        showToast("Erro ao ler Excel", "error");
      }
    };

    reader.readAsBinaryString(file);
  };

  const executeTrack = (value) => {
    const searchVal = value || szInput;

    if (!searchVal) return showToast("Digite a SZ", "error");

    if (!dataZSD036.length || !dataMB51.length) {
      return showToast("Carregue as planilhas!", "error");
    }

    setIsProcessing(true);

    try {
      const szRecord = dataZSD036.find(row => 
        clean(row[COLUMN_MAP.ZSD036.SZ]) === clean(searchVal)
      );

      console.log("BUSCA:", searchVal);
      console.log("ENCONTRADO:", szRecord);

      if (!szRecord) throw new Error("SZ não encontrada");

      const batchId = szRecord[COLUMN_MAP.ZSD036.LOTE];

      const movements = dataMB51.filter(row => 
        clean(row[COLUMN_MAP.MB51.LOTE]) === clean(batchId)
      );

      if (!movements.length) throw new Error("Lote não encontrado");

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
    <div className="min-h-screen bg-slate-100 p-6">
      <input
        value={szInput}
        onChange={(e) => setSzInput(e.target.value)}
        onKeyDown={(e) => e.key === 'Enter' && executeTrack()}
        placeholder="Digite SZ"
      />

      <button onClick={() => executeTrack()}>
        CONSULTAR
      </button>

      <input type="file" onChange={(e) => handleFileUpload(e, 'ZSD036')} />
      <input type="file" onChange={(e) => handleFileUpload(e, 'MB51')} />

      {result && (
        <div>
          <p>SZ: {result.sz}</p>
          <p>Lote: {result.lote}</p>
          <p>Peso: {result.peso}</p>
        </div>
      )}
    </div>
  );
}
