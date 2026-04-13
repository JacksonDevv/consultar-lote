import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import { Html5QrcodeScanner } from 'html5-qrcode';
import { 
  Upload, Search, Mail, RefreshCw, CheckCircle2, 
  FileText, AlertCircle, Camera, X 
} from 'lucide-react';

// Configuração de colunas - Ajuste aqui se os nomes no Excel mudarem
const COLUMN_MAP = {
  ZSD036: { SZ: 'SZ', LOTE: 'Lote' },
  MB51: { 
    LOTE: 'Lote', PESO: 'Peso', TRANSACAO: 'Tipo Movimento', 
    DATA: 'Data Lançamento', HORA: 'Hora Lançamento', 
    USUARIO: 'Usuário', CABECALHO: 'Documento' 
  }
};

export default function SAPBatchTracker() {
  // Dados
  const [dataZSD036, setDataZSD036] = useState([]);
  const [dataMB51, setDataMB51] = useState([]);
  const [szInput, setSzInput] = useState('');
  const [result, setResult] = useState(null);
  const [email, setEmail] = useState('');
  const [saveEmail, setSaveEmail] = useState(false);

  // UI
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

  // --- LOGICA DO SCANNER (SÊNIOR APPROACH) ---
  useEffect(() => {
    if (isScanning) {
      // Pequeno delay para garantir que o HTML do modal foi renderizado
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
          // Dispara o rastreio automaticamente após ler
          setTimeout(() => executeTrack(decodedText), 500);
        }, (error) => {
          // Erros de scan são ignorados para evitar spam de logs
        });

        scannerRef.current = scanner;
      }, 100);

      return () => {
        clearTimeout(timeout);
        if (scannerRef.current) {
          scannerRef.current.clear();
        }
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
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        
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
    if (!dataZSD036.length || !dataMB51.length) return showToast("Carregue as duas planilhas!", "error");

    setIsProcessing(true);

    try {
      const szRecord = dataZSD036.find(row => 
        String(row[COLUMN_MAP.ZSD036.SZ] || '').trim().toLowerCase() === searchVal.trim().toLowerCase()
      );

      if (!szRecord) throw new Error("SZ não encontrada na ZSD036");

      const batchId = szRecord[COLUMN_MAP.ZSD036.LOTE];
      const movements = dataMB51.filter(row => 
        String(row[COLUMN_MAP.MB51.LOTE] || '').trim() === String(batchId || '').trim()
      );

      if (!movements.length) throw new Error("Lote não encontrado na MB51");

      const sorted = movements.sort((a, b) => {
        const d1 = new Date(`${a[COLUMN_MAP.MB51.DATA]} ${a[COLUMN_MAP.MB51.HORA]}`);
        const d2 = new Date(`${b[COLUMN_MAP.MB51.DATA]} ${b[COLUMN_MAP.MB51.HORA]}`);
        return d2 - d1;
      });

      const latest = sorted[0];
      setResult({
        sz: searchVal, lote: batchId, peso: latest[COLUMN_MAP.MB51.PESO],
        transacao: latest[COLUMN_MAP.MB51.TRANSACAO], data: latest[COLUMN_MAP.MB51.DATA],
        hora: latest[COLUMN_MAP.MB51.HORA], usuario: latest[COLUMN_MAP.MB51.USUARIO],
        cabecalho: latest[COLUMN_MAP.MB51.CABECHALHO]
      });
    } catch (err) {
      showToast(err.message, "error");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 p-4 md:p-8 font-sans">
      {/* HEADER */}
      <header className="max-w-6xl mx-auto mb-8 bg-white p-6 rounded-xl shadow-sm border-b-4 border-blue-600 flex justify-between items-center">
        <div className="flex items-center gap-3">
          <div className="bg-blue-600 p-2 rounded-lg text-white"><Search size={24}/></div>
          <h1 className="text-2xl font-bold text-slate-700">SAP <span className="text-blue-600">BatchTracker</span></h1>
        </div>
        <span className="text-xs text-slate-400 font-medium uppercase">v1.2 Stable</span>
      </header>

      <main className="max-w-6xl mx-auto grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* COLUNA ESQUERDA */}
        <div className="lg:col-span-4 space-y-6">
          <section className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="text-xs font-bold text-slate-500 uppercase mb-4 flex items-center gap-2"><FileText size={14}/> Dados SAP</h2>
            <div className="space-y-4">
              {['ZSD036', 'MB51'].map(type => (
                <div key={type}>
                  <label className="block text-xs font-bold text-slate-600 mb-1">{`Planilha ${type}`}</label>
                  <div 
                    onClick={() => document.getElementById(`f-${type}`).click()}
                    className={`p-4 border-2 border-dashed rounded-lg cursor-pointer transition-all flex flex-col items-center gap-2 
                    ${uploads[type].status === 'complete' ? 'border-green-500 bg-green-50' : 'border-slate-300 hover:bg-slate-50'}`}
                  >
                    <Upload size={20} className={uploads[type].status === 'complete' ? 'text-green-600' : 'text-slate-400'}/>
                    <span className="text-xs text-slate-500 truncate w-full text-center">
                      {uploads[type].fileName || 'Clique para Upload'}
                    </span>
                    <input id={`f-${type}`} type="file" className="hidden" accept=".xlsx, .csv" onChange={(e) => handleFileUpload(e, type)} />
                  </div>
                </div>
              ))}
            </div>
          </section>

          <section className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
            <h2 className="text-xs font-bold text-slate-500 uppercase mb-4">Rastreamento</h2>
            <div className="space-y-4">
              <div className="flex gap-2">
                <input 
                  className="flex-1 p-3 border rounded-lg outline-none focus:ring-2 focus:ring-blue-500" 
                  placeholder="Ex: SZ12345" 
                  value={szInput} 
                  onChange={e => setSzInput(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && executeTrack()}
                />
                <button onClick={() => setIsScanning(true)} className="p-3 bg-slate-100 rounded-lg hover:bg-slate-200 text-slate-600 transition-all"><Camera size={20}/></button>
              </div>
              <button 
                onClick={() => executeTrack()} 
                disabled={isProcessing}
                className="w-full bg-blue-600 text-white font-bold py-3 rounded-lg hover:bg-blue-700 disabled:opacity-50 transition-all flex justify-center gap-2"
              >
                {isProcessing ? <RefreshCw className="animate-spin" size={20}/> : <Search size={20}/>} RASTREAR
              </button>
            </div>
          </section>
        </div>

        {/* COLUNA DIREITA */}
        <div className="lg:col-span-8">
          {!result ? (
            <div className="h-full min-h-[400px] flex flex-col items-center justify-center bg-white rounded-xl border-2 border-dashed border-slate-300 p-12 text-center text-slate-400">
              <Search size={48} className="mb-4 opacity-20"/>
              <p>Aguardando entrada de dados para iniciar a busca...</p>
            </div>
          ) : (
            <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4">
              <div className="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
                <div className="bg-slate-50 p-4 border-b flex justify-between items-center">
                  <h3 className="font-bold text-slate-700 flex items-center gap-2"><CheckCircle2 className="text-blue-600" size={18}/> Resultado</h3>
                  <button onClick={() => {setResult(null); setSzInput('')}} className="text-xs text-blue-600 font-bold hover:underline">Nova Consulta</button>
                </div>
                <div className="p-6 grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-3">
                    <Detail label="SZ" value={result.sz} />
                    <Detail label="Lote" value={result.lote} highlight />
                    <Detail label="Peso" value={result.peso} />
                  </div>
                  <div className="space-y-3">
                    <Detail label="Transação" value={result.transacao} />
                    <Detail label="Data/Hora" value={`${result.data} ${result.hora}`} />
                    <Detail label="Usuário" value={result.usuario} />
                    <Detail label="Doc. Cabeçalho" value={result.cabecalho} />
                  </div>
                </div>
              </div>
              
              <div className="bg-white p-6 rounded-xl shadow-sm border border-slate-200">
                <h3 className="text-xs font-bold text-slate-500 uppercase mb-4 flex items-center gap-2"><Mail size={14}/> Enviar Notificação</h3>
                <div className="flex flex-col md:flex-row gap-4 items-end">
                  <input className="flex-1 p-3 border rounded-lg outline-none" placeholder="email@empresa.com" value={email} onChange={e => setEmail(e.target.value)} />
                  <div className="flex items-center gap-2 mb-3">
                    <input type="checkbox" checked={saveEmail} onChange={e => setSaveEmail(e.target.checked)} className="w-4 h-4" />
                    <label className="text-xs text-slate-500">Salvar e-mail</label>
                  </div>
                  <button onClick={() => {
                    if(!email) return showToast("Insira o e-mail", "error");
                    if(saveEmail) localStorage.setItem('sap_tracker_email', email);
                    window.location.href = `mailto:${email}?subject=Rastreio SAP ${result.sz}&body=SZ: ${result.sz}%0ALote: ${result.lote}`;
                    showToast("E-mail aberto!");
                  }} className="bg-slate-800 text-white px-6 py-3 rounded-lg font-bold hover:bg-slate-900 transition-all">Enviar</button>
                </div>
              </div>
            </div>
          )}
        </div>
      </main>

      {/* MODAL SCANNER */}
      {isScanning && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/80 backdrop-blur-sm p-4">
          <div className="bg-white rounded-2xl overflow-hidden w-full max-w-md relative">
            <div className="p-4 flex justify-between items-center border-b">
              <span className="font-bold text-slate-700 flex items-center gap-2"><Camera size={18}/> Scanner de SZ</span>
              <button onClick={() => setIsScanning(false)} className="p-1 hover:bg-slate-100 rounded-full"><X size={20}/></button>
            </div>
            <div className="p-6">
              <div id="reader" className="overflow-hidden rounded-lg"></div>
            </div>
          </div>
        </div>
      )}

      {/* TOAST */}
      {toast && (
        <div className={`fixed bottom-8 right-8 p-4 rounded-lg shadow-lg text-white flex items-center gap-3 z-50 animate-in slide-in-from-right-full ${toast.type === 'error' ? 'bg-red-500' : 'bg-green-600'}`}>
          {toast.type === 'error' ? <AlertCircle size={20}/> : <CheckCircle2 size={20}/>}
          <span className="font-medium">{toast.msg}</span>
        </div>
      )}
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
