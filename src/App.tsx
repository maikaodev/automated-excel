import { useState, useRef, ChangeEvent, DragEvent } from 'react';
import ExcelJS from 'exceljs';
import './App.css';

// Tipagem dos dados
type Pessoa = {
  nome: string;
  cpf: string;
  foto: File | null;
};

function App() {
  // Estados do formulário
  const [pessoas, setPessoas] = useState<Pessoa[]>([
    { nome: '', cpf: '', foto: null }, // Titular (pessoas[0])
  ]);
  const [sipra, setSipra] = useState('');
  const [endereco, setEndereco] = useState('São José do Capricho, em Flexeiras/AL');
  const [data, setData] = useState('');
  const [fotos, setFotos] = useState<File[]>([]);
  const [isGenerating, setIsGenerating] = useState(false);

  // Ref para o input de arquivo
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Lógica de Pessoas (Titular e Cônjuge)
  const adicionarPessoa = () => {
    if (pessoas.length < 2) {
      setPessoas([...pessoas, { nome: '', cpf: '', foto: null }]);
    }
  };

  const removerPessoa = (index: number) => {
    setPessoas((prev) => prev.filter((_, i) => i !== index));
  };

  const atualizarPessoa = (index: number, campo: keyof Pessoa, valor: any) => {
    const novaLista = [...pessoas];
    novaLista[index] = { ...novaLista[index], [campo]: valor };
    setPessoas(novaLista);
  };

  // Lógica de Fotos (Dropzone)
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

  // ==========================================
  // Lógica do EXCELJS (Mapeamento Exato)
  // ==========================================
  // const handleGenerate = async () => {
  //   if (!pessoas[0].nome || !pessoas[0].cpf || !sipra) {
  //     alert('Por favor, preencha Nome, CPF do Titular e o SIPRA.');
  //     return;
  //   }
  //   setIsGenerating(true);

  //   try {
  //     // 1. Buscar o modelo
  //     const response = await fetch('/modelo_base.xlsx');
  //     if (!response.ok) throw new Error('Arquivo modelo_base.xlsx não encontrado.');
  //     const buffer = await response.arrayBuffer();

  //     // 2. Carregar o modelo
  //     const workbook = new ExcelJS.Workbook();
  //     await workbook.xlsx.load(buffer);

  //     // 3. Acessar a aba "0. Dados" (não usar índice, usar o nome da aba)
  //     const worksheet = workbook.getWorksheet('0. Dados');
  //     if (!worksheet) {
  //       throw new Error('Aba "0. Dados" não encontrada no modelo.');
  //     }

  //     // ==========================================
  //     // PREENCHIMENTO DOS CAMPOS (Mapeamento Correto)
  //     // ==========================================

  //     // Endereço (Linha 7, Coluna F)
  //     worksheet.getCell('F7').value = endereco;

  //     // SIPRA (Linha 7, Coluna AA)
  //     worksheet.getCell('AA7').value = sipra;

  //     // Titular 1 (Linha 9)
  //     worksheet.getCell('D9').value = pessoas[0].nome;   // Nome
  //     worksheet.getCell('Z9').value = pessoas[0].cpf;    // CPF

  //     // Titular 2 - Cônjuge (Linha 11) - Se existir
  //     if (pessoas.length > 1) {
  //       worksheet.getCell('D11').value = pessoas[1].nome;
  //       worksheet.getCell('Z11').value = pessoas[1].cpf;
  //     } else {
  //       // Se não tiver cônjuge, limpa as células (opcional)
  //       worksheet.getCell('D11').value = '';
  //       worksheet.getCell('Z11').value = '';
  //     }

  //     // Data da Vistoria (Linha 13, Coluna X) - Sobrescreve a fórmula =TODAY()
  //     worksheet.getCell('X13').value = data || new Date().toISOString().split('T')[0];

  //     // ==========================================
  //     // INSERÇÃO DAS FOTOS (Sugestão de posicionamento)
  //     // Vamos inserir as fotos na aba "3. Edificação" que tem o título "RELATÓRIO FOTOGRÁFICO"
  //     // ou na aba "4.Reg. Fotografico". Vou usar a aba "3. Edificação" como exemplo.
  //     // ==========================================
  //     if (fotos.length > 0) {
  //       const worksheetFotos = workbook.getWorksheet('3. Edificação');
  //       // Se preferir, pode mudar para '4.Reg. Fotografico'

  //       if (worksheetFotos) {
  //         for (let i = 0; i < fotos.length; i++) {
  //           const foto = fotos[i];
  //           const fotoBuffer = await foto.arrayBuffer();

  //           const imageId = workbook.addImage({
  //             buffer: fotoBuffer,
  //             extension: foto.name.split('.').pop() || 'jpg',
  //           });

  //           // Posição: A partir da linha 10, coluna B (por exemplo)
  //           // Ajuste essas coordenadas (col, row) conforme a sua necessidade visual no modelo
  //           const linha = 10 + i;
  //           worksheetFotos.addImage(imageId, {
  //             tl: { col: 1, row: linha - 1 }, // Coluna B
  //             br: { col: 5, row: linha + 3 }, // Tamanho da imagem
  //           });
  //         }
  //       }
  //     }

  //     // 4. Gerar buffer e download
  //     const outBuffer = await workbook.xlsx.writeBuffer();
  //     const blob = new Blob([outBuffer], {
  //       type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  //     });
  //     const link = document.createElement('a');
  //     link.href = URL.createObjectURL(blob);
  //     link.download = `planilha_${pessoas[0].nome}.xlsx`;
  //     link.click();

  //   } catch (error: any) {
  //     console.error('Erro detalhado:', error);
  //     alert('❌ Erro ao gerar: ' + (error.message || 'Erro desconhecido'));
  //   } finally {
  //     setIsGenerating(false);
  //   }
  // };

  // ==========================================
  // Lógica do EXCELJS (Geração Completa)
  // ==========================================
  const handleGenerate = async () => {
    if (!pessoas[0].nome || !pessoas[0].cpf || !sipra) {
      alert('Por favor, preencha Nome, CPF do Titular e o SIPRA.');
      return;
    }
    setIsGenerating(true);

    try {
      // 1. Buscar o modelo base na pasta public/
      const response = await fetch('/modelo_base.xlsx');
      if (!response.ok) throw new Error('Arquivo modelo_base.xlsx não encontrado.');
      const buffer = await response.arrayBuffer();

      // 2. Carregar o modelo com o ExcelJS
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      // ==========================================
      // 3. PREENCHIMENTO DA ABA "0. Dados"
      // ==========================================
      const wsDados = workbook.getWorksheet('0. Dados');
      if (!wsDados) throw new Error('Aba "0. Dados" não encontrada no modelo.');

      // Endereço (F7)
      wsDados.getCell('F7').value = endereco;

      // SIPRA (AA7)
      wsDados.getCell('AA7').value = sipra;

      // Titular 1 (D9 e Z9)
      wsDados.getCell('D9').value = pessoas[0].nome;
      wsDados.getCell('Z9').value = pessoas[0].cpf;

      // Titular 2 - Cônjuge (D11 e Z11) - Se existir
      if (pessoas.length > 1) {
        wsDados.getCell('D11').value = pessoas[1].nome;
        wsDados.getCell('Z11').value = pessoas[1].cpf;
      } else {
        wsDados.getCell('D11').value = '';
        wsDados.getCell('Z11').value = '';
      }

      // Data da Vistoria (X13)
      wsDados.getCell('X13').value = data || new Date().toISOString().split('T')[0];

      // ==========================================
      // 4. INSERÇÃO DAS FOTOS NA ABA "4.Reg. Fotografico"
      // ==========================================
      const wsFotos = workbook.getWorksheet('4.Reg. Fotografico');
      if (!wsFotos) throw new Error('Aba "4.Reg. Fotografico" não encontrada.');

      // Tamanho padrão das imagens (ajuste esses números para aumentar/diminuir as fotos)
      const LARGURA_IMG = 6; // Ocupa 6 colunas de largura
      const ALTURA_IMG = 6;  // Ocupa 6 linhas de altura
      const MAX_POR_LINHA = 4; // Até 4 fotos por linha

      for (let i = 0; i < fotos.length; i++) {
        const foto = fotos[i];
        const fotoBuffer = await foto.arrayBuffer();

        const imageId = workbook.addImage({
          buffer: fotoBuffer,
          extension: foto.name.split('.').pop() || 'jpg',
        });

        // Distribuição em Grid 4x2
        // Colunas: B (1), H (7), N (13), T (19) - índice 0-based
        const colIndex = 1 + (i % MAX_POR_LINHA) * (LARGURA_IMG + 1);
        // Linhas: 13 (12), 19 (18), 25 (24), 31 (30) - índice 0-based
        const rowIndex = 12 + Math.floor(i / MAX_POR_LINHA) * (ALTURA_IMG + 1);

        wsFotos.addImage(imageId, {
          tl: { col: colIndex, row: rowIndex },
          br: { col: colIndex + LARGURA_IMG, row: rowIndex + ALTURA_IMG },
        });
      }

      // ==========================================
      // 5. GERAR O ARQUIVO E BAIXAR
      // ==========================================
      const outBuffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([outBuffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = `planilha_${pessoas[0].nome}.xlsx`;
      link.click();

    } catch (error: any) {
      console.error('Erro detalhado:', error);
      alert('❌ Erro ao gerar: ' + (error.message || 'Erro desconhecido'));
    } finally {
      setIsGenerating(false);
    }
  };

  // ==========================================
  // RENDERIZAÇÃO DO FORMULÁRIO (adaptado)
  // ==========================================
  return (
    <div className="container">
      <header className="header">
        <div className="header-left">
          <div className="header-icon">📄</div>
          Gerador de Planilhas
        </div>
      </header>

      <h1 className="page-title">Nova Planilha</h1>
      <p className="page-subtitle">Preencha os dados para gerar a planilha do INCRA.</p>

      <div className="content-grid">
        {/* Coluna Esquerda: Dados */}
        <div className="form-group">

          {/* Endereço e SIPRA */}
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

          <div className="input-wrapper">
            <label>Data da Vistoria</label>
            <input type="date" value={data} onChange={(e) => setData(e.target.value)} />
          </div>

          {/* Pessoas (Titular e Cônjuge) */}
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
                  value={pessoa.cpf}
                  onChange={(e) => atualizarPessoa(index, 'cpf', e.target.value)}
                />
              </div>
              <div className="input-wrapper">
                <label>Foto do Beneficiário</label>
                <input
                  type="file"
                  accept="image/*"
                  onChange={(e) => {
                    if (e.target.files?.[0]) {
                      atualizarPessoa(index, 'foto', e.target.files[0]);
                    }
                  }}
                />
              </div>
            </div>
          ))}

          {pessoas.length < 2 && (
            <button onClick={adicionarPessoa} style={{ marginTop: '1rem', padding: '0.5rem', width: '100%', border: '1px dashed #6b46c1', background: 'transparent', borderRadius: '8px', cursor: 'pointer' }}>
              + Adicionar Cônjuge
            </button>
          )}
        </div>

        {/* Coluna Direita: Registro Fotográfico Geral */}
        <div className="upload-section">
          <div className="section-title">📸 Registro Fotográfico (Geral)</div>
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
      <button className="btn-generate" onClick={handleGenerate} disabled={isGenerating}>
        {isGenerating ? 'Gerando...' : 'Gerar Planilha Excel'}
      </button>
    </div>
  );
}

export default App;
