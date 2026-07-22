import { useState, useRef, ChangeEvent, DragEvent } from 'react';
import ExcelJS from 'exceljs';
import './App.css';

type Pessoa = {
  nome: string;
  cpf: string;
};

function App() {
  const [pessoas, setPessoas] = useState<Pessoa[]>([{ nome: '', cpf: '' }]);
  const [sipra, setSipra] = useState('');
  const [endereco, setEndereco] = useState('São José do Capricho, em Flexeiras/AL');
  const [fotos, setFotos] = useState<File[]>([]);

  // NOVO ESTADO: Assinatura do Titular
  const [assinatura, setAssinatura] = useState<File | null>(null);

  const [isGenerating, setIsGenerating] = useState(false);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const signatureInputRef = useRef<HTMLInputElement>(null);

  // ================= LÓGICA DE PESSOAS =================
  const adicionarPessoa = () => {
    if (pessoas.length < 2) {
      setPessoas([...pessoas, { nome: '', cpf: '' }]);
    }
  };

  const removerPessoa = (index: number) => {
    setPessoas((prev) => prev.filter((_, i) => i !== index));
  };

  const atualizarPessoa = (index: number, campo: keyof Pessoa, valor: string) => {
    const novaLista = [...pessoas];
    novaLista[index] = { ...novaLista[index], [campo]: valor };
    setPessoas(novaLista);
  };

  // ================= MÁSCARA DE CPF =================
  const formatarCPF = (cpf: string) => {
    const cleaned = cpf.replace(/\D/g, '');
    if (cleaned.length <= 3) return cleaned;
    if (cleaned.length <= 6) return `${cleaned.slice(0, 3)}.${cleaned.slice(3)}`;
    if (cleaned.length <= 9) return `${cleaned.slice(0, 3)}.${cleaned.slice(3, 6)}.${cleaned.slice(6)}`;
    return `${cleaned.slice(0, 3)}.${cleaned.slice(3, 6)}.${cleaned.slice(6, 9)}-${cleaned.slice(9, 11)}`;
  };

  const handleCpfChange = (index: number, value: string) => {
    const onlyNumbers = value.replace(/\D/g, '');
    if (onlyNumbers.length <= 11) {
      atualizarPessoa(index, 'cpf', onlyNumbers);
    }
  };

  // ================= LÓGICA DE FOTOS =================
  const handleAddFiles = (newFiles: File[]) => {
    const remainingSlots = 8 - fotos.length;
    const filesToAdd = newFiles.slice(0, remainingSlots);
    setFotos((prev) => [...prev, ...filesToAdd]);
  };

  const onFileChange = (e: ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      handleAddFiles(Array.from(e.target.files));
      e.target.value = '';
    }
  };

  const onDrop = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    if (e.dataTransfer.files) {
      handleAddFiles(Array.from(e.dataTransfer.files));
    }
  };

  const onDragOver = (e: DragEvent<HTMLDivElement>) => {
    e.preventDefault();
  };

  const removeFoto = (index: number) => {
    setFotos((prev) => prev.filter((_, i) => i !== index));
  };

  // ================= VALIDAÇÃO E MODAL =================
  const onGenerateClick = () => {
    if (!pessoas[0].nome || !pessoas[0].cpf || !sipra) {
      alert('Por favor, preencha Nome, CPF do Titular e o SIPRA.');
      return;
    }
    setIsModalOpen(true);
  };

  const handleConfirmGeneration = () => {
    setIsModalOpen(false);
    handleGenerate();
  };

  // ================= LÓGICA DO EXCEL (GERAÇÃO) =================
  const handleGenerate = async () => {
    setIsGenerating(true);

    try {
      const response = await fetch('/modelo_base.xlsx');
      if (!response.ok) throw new Error('Arquivo modelo_base.xlsx não encontrado.');
      const buffer = await response.arrayBuffer();

      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      // ABA "0. Dados"
      const wsDados = workbook.getWorksheet('0. Dados');
      if (!wsDados) throw new Error('Aba "0. Dados" não encontrada.');

      wsDados.getCell('F7').value = endereco;
      wsDados.getCell('AA7').value = sipra;
      wsDados.getCell('D9').value = pessoas[0].nome;
      // ✨ CORREÇÃO: Inserir o CPF já formatado com máscara
      wsDados.getCell('Z9').value = formatarCPF(pessoas[0].cpf);
      wsDados.getCell('D11').value = pessoas[1]?.nome || '';
      wsDados.getCell('Z11').value = pessoas[1]?.cpf ? formatarCPF(pessoas[1].cpf) : '';

      // ABA "4.Reg. Fotografico" (Grid 4x2)
      const wsFotos = workbook.getWorksheet('4.Reg. Fotografico');
      if (!wsFotos) throw new Error('Aba "4.Reg. Fotografico" não encontrada.');

      const LARGURA_IMG = 6;
      const ALTURA_IMG = 6;
      const MAX_POR_LINHA = 4;

      for (let i = 0; i < fotos.length; i++) {
        const foto = fotos[i];
        const fotoBuffer = await foto.arrayBuffer();

        const imageId = workbook.addImage({
          buffer: fotoBuffer,
          extension: foto.name.split('.').pop() || 'jpg',
        });

        const colIndex = 1 + (i % MAX_POR_LINHA) * (LARGURA_IMG + 1);
        const rowIndex = 12 + Math.floor(i / MAX_POR_LINHA) * (ALTURA_IMG + 1);

        wsFotos.addImage(imageId, {
          tl: { col: colIndex, row: rowIndex },
          br: { col: colIndex + LARGURA_IMG, row: rowIndex + ALTURA_IMG },
        });
      }

      // ✨ NOVO: ABA "1. Proposta" - Inserir assinatura em D46
      if (assinatura) {
        const wsProposta = workbook.getWorksheet('1. Proposta');
        if (wsProposta) {
          const sigBuffer = await assinatura.arrayBuffer();
          const imageId = workbook.addImage({
            buffer: sigBuffer,
            extension: assinatura.name.split('.').pop() || 'png',
          });

          // D46 no Excel: Coluna D = índice 3, Linha 46 = índice 45
          // Tamanho: Ocupa colunas D até G (largura) e linhas 46 até 51 (altura)
          wsProposta.addImage(imageId, {
            tl: { col: 3, row: 45 },
            br: { col: 7, row: 51 },
          });
        }
      }

      // GERAR E BAIXAR
      const outBuffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([outBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);

      // Nome do arquivo
      const currentYear = new Date().getFullYear();
      let nomeBase = pessoas[0].nome.trim();
      if (pessoas.length > 1 && pessoas[1].nome.trim() !== '') {
        nomeBase += ` - ${pessoas[1].nome.trim()}`;
      }
      link.download = `${nomeBase}_${currentYear}.xlsx`;
      link.click();

      setSuccessMessage('✅ Planilha gerada com sucesso!');

      // Limpeza do formulário (incluindo a assinatura)
      setTimeout(() => {
        setPessoas([{ nome: '', cpf: '' }]);
        setSipra('');
        setEndereco('São José do Capricho, em Flexeiras/AL');
        setFotos([]);
        setAssinatura(null); // Limpa a assinatura
        setIsGenerating(false);
        setTimeout(() => setSuccessMessage(null), 3000);
      }, 1000);

    } catch (error: any) {
      console.error(error);
      alert('❌ Erro ao gerar: ' + error.message);
      setIsGenerating(false);
    }
  };

  // ================= RENDERIZAÇÃO =================
  return (
    <div className="container">
      {/* Toast de Sucesso */}
      {successMessage && (
        <div className="toast-success">
          {successMessage}
        </div>
      )}

      {/* Modal de Confirmação */}
      {isModalOpen && (
        <div className="modal-overlay">
          <div className="modal-content">
            <h3>📋 Confirme os dados</h3>
            <p><strong>Nome:</strong> {pessoas[0].nome}</p>
            <p><strong>CPF:</strong> {formatarCPF(pessoas[0].cpf)}</p>
            {pessoas[1] && <p><strong>Cônjuge:</strong> {pessoas[1].nome}</p>}
            <p><strong>SIPRA:</strong> {sipra}</p>
            <p><strong>Endereço:</strong> {endereco}</p>
            <p><strong>Total de fotos:</strong> {fotos.length} / 8</p>
            <p><strong>Assinatura anexada:</strong> {assinatura ? '✅ Sim' : '❌ Não'}</p>

            <div className="modal-actions">
              <button className="btn-cancel" onClick={() => setIsModalOpen(false)}>Cancelar</button>
              <button className="btn-confirm" onClick={handleConfirmGeneration}>Confirmar e Gerar</button>
            </div>
          </div>
        </div>
      )}

      {/* Header */}
      <header className="header">
        <div className="header-left">
          <div className="header-icon">📄</div>
          Gerador de Planilhas
        </div>
      </header>

      <h1 className="page-title">Nova Planilha</h1>

      <div className="content-grid">
        {/* Coluna Esquerda: Dados */}
        <div className="form-group">
          <div className="input-row">
            <div className="input-wrapper">
              <label>Endereço</label>
              <input value={endereco} onChange={(e) => setEndereco(e.target.value)} />
            </div>
            <div className="input-wrapper">
              <label>SIPRA</label>
              <input placeholder="Ex: AL018600000007" value={sipra} onChange={(e) => setSipra(e.target.value)} />
            </div>
          </div>

          {/* Pessoas */}
          {pessoas.map((pessoa, index) => (
            <div key={index} style={{ borderTop: '1px solid #e2e8f0', paddingTop: '1rem', marginTop: '1rem' }}>
              <div className="section-title" style={{ color: '#1a202c' }}>
                {index === 0 ? '👤 Titular' : '👥 Cônjuge'}
                {pessoas.length > 1 && (
                  <button onClick={() => removerPessoa(index)} style={{ marginLeft: 'auto', color: 'red', border: 'none', background: 'transparent', cursor: 'pointer' }}>✕</button>
                )}
              </div>
              <div className="input-wrapper">
                <label>Nome Completo</label>
                <input
                  placeholder="Nome"
                  value={pessoa.nome}
                  onChange={(e) => atualizarPessoa(index, 'nome', e.target.value)}
                />
              </div>
              <div className="input-wrapper">
                <label>CPF</label>
                <input
                  placeholder="000.000.000-00"
                  value={pessoa.cpf ? formatarCPF(pessoa.cpf) : ''}
                  onChange={(e) => handleCpfChange(index, e.target.value)}
                />
              </div>
            </div>
          ))}

          {pessoas.length < 2 && (
            <button onClick={adicionarPessoa} style={{ marginTop: '1rem', padding: '0.5rem', width: '100%', border: '1px dashed #6b46c1', background: 'transparent', borderRadius: '8px', cursor: 'pointer' }}>
              + Adicionar Cônjuge
            </button>
          )}

          {/* ✨ NOVO: Upload da Assinatura (Apenas Titular) */}
          <div style={{ marginTop: '1.5rem', borderTop: '1px solid #e2e8f0', paddingTop: '1rem' }}>
            <div className="section-title" style={{ color: '#1a202c' }}>
              ✍️ Assinatura do Titular
            </div>
            <div
              style={{
                border: '1px dashed #6b46c1',
                borderRadius: '8px',
                padding: '1rem',
                textAlign: 'center',
                backgroundColor: '#faf5ff',
                cursor: 'pointer',
                marginTop: '0.5rem'
              }}
              onClick={() => signatureInputRef.current?.click()}
            >
              {assinatura ? (
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.5rem' }}>
                  <span style={{ fontSize: '1.5rem' }}>✅</span>
                  <span>{assinatura.name}</span>
                  <button
                    onClick={(e) => { e.stopPropagation(); setAssinatura(null); }}
                    style={{ marginLeft: 'auto', color: 'red', border: 'none', background: 'transparent', cursor: 'pointer', fontSize: '1.2rem' }}
                  >✕</button>
                </div>
              ) : (
                <div>
                  <div style={{ fontSize: '1.5rem', marginBottom: '0.25rem' }}>🖊️</div>
                  <div>Clique para anexar a assinatura (sem fundo)</div>
                </div>
              )}
              <input
                type="file"
                ref={signatureInputRef}
                accept="image/*"
                style={{ display: 'none' }}
                onChange={(e) => {
                  if (e.target.files?.[0]) {
                    setAssinatura(e.target.files[0]);
                  }
                  e.target.value = '';
                }}
              />
            </div>
          </div>
        </div>

        {/* Coluna Direita: Fotos */}
        <div className="upload-section">
          <div className="section-title">📸 Registro Fotográfico</div>
          <div
            className="dropzone"
            onClick={() => fileInputRef.current?.click()}
            onDrop={onDrop}
            onDragOver={onDragOver}
          >
            <div className="dropzone-icon">☁️</div>
            <div className="dropzone-title">Arraste as fotos da vistoria aqui</div>
            <div className="dropzone-subtitle">Máximo 8 fotos</div>
            <input type="file" ref={fileInputRef} multiple accept="image/*" style={{ display: 'none' }} onChange={onFileChange} />
          </div>

          <div className="photos-header">
            <span>Fotos adicionadas ({fotos.length}/8)</span>
          </div>
          <div className="photos-grid">
            {Array.from({ length: 8 }).map((_, index) => {
              const foto = fotos[index];
              return (
                <div key={index} className={`photo-item ${foto ? '' : 'photo-empty'}`}>
                  {foto ? (
                    <>
                      <img src={URL.createObjectURL(foto)} alt="Preview" />
                      <button className="photo-remove" onClick={() => removeFoto(index)}>×</button>
                    </>
                  ) : null}
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* Botão Gerar */}
      <button className="btn-generate" onClick={onGenerateClick} disabled={isGenerating || isModalOpen}>
        {isGenerating ? 'Gerando...' : 'Gerar Planilha Excel'}
      </button>
    </div>
  );
}

export default App;
