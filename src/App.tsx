import { useState } from 'react';
import ExcelJS from 'exceljs';

function App() {
  // Estado dos dados
  const [pessoas, setPessoas] = useState([
    { nome: '', cpf: '', foto: null },
  ]);
  const [sipra, setSipra] = useState('');

  // Função para adicionar cônjuge (máximo 2)
  const adicionarPessoa = () => {
    if (pessoas.length < 2) {
      setPessoas([...pessoas, { nome: '', cpf: '', foto: null }]);
    }
  };

  const removerPessoa = (index: number) => {
    const novaLista = pessoas.filter((_, i) => i !== index);
    setPessoas(novaLista);
  };

  const handleFoto = (index: number, file: File) => {
    const novaLista = [...pessoas];
    novaLista[index].foto = file;
    setPessoas(novaLista);
  };

  // A MÁGICA DO EXCEL
  const gerarPlanilha = async () => {
    // 1. Buscar o modelo que está na pasta public/
    const response = await fetch('/modelo_base.xlsx');
    const buffer = await response.arrayBuffer();

    // 2. Carregar o modelo com ExcelJS
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.getWorksheet(1); // Primeira aba

    // 3. Preencher os dados (supondo que as células estejam organizadas assim)
    // Exemplo: Linha 1 = Titular, Linha 2 = Cônjuge (se existir)
    pessoas.forEach((pessoa, index) => {
      const linha = index + 1; // Linha 1 e Linha 2

      // Ajuste os nomes das células conforme seu modelo Excel!
      worksheet.getCell(`A${linha}`).value = pessoa.nome;
      worksheet.getCell(`B${linha}`).value = pessoa.cpf;

      // SIPRA vai apenas na primeira pessoa (titular), conforme você disse
      if (index === 0) {
        worksheet.getCell(`C${linha}`).value = sipra;
      }

      // Inserir a foto (se existir)
      if (pessoa.foto) {
        const fotoBuffer = await pessoa.foto.arrayBuffer();
        const imageId = workbook.addImage({
          buffer: fotoBuffer,
          extension: pessoa.foto.name.split('.').pop() || 'jpg',
        });

        // Posiciona a foto ao lado do nome (ou onde você quiser)
        // Exemplo: na coluna D, na mesma linha da pessoa
        worksheet.addImage(imageId, {
          tl: { col: 3, row: linha - 1 }, // 0-based, então linha 1 = row 0
          br: { col: 5, row: linha + 1 }, // Tamanho da imagem
        });
      }
    });

    // 4. Gerar o buffer final e baixar
    const outBuffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([outBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `planilha_${pessoas[0]?.nome || 'gerado'}.xlsx`;
    link.click();
  };

  return (
    <div style={{ padding: '2rem' }}>
      <h1>Gerador de Planilhas</h1>

      <div>
        <label>SIPRA (Titular):</label>
        <input value={sipra} onChange={(e) => setSipra(e.target.value)} />
      </div>

      {pessoas.map((p, index) => (
        <div key={index} style={{ border: '1px solid #ccc', margin: '1rem 0', padding: '1rem' }}>
          <h3>{index === 0 ? 'Titular' : 'Cônjuge'}</h3>
          <input
            placeholder="Nome"
            value={p.nome}
            onChange={(e) => {
              const nova = [...pessoas];
              nova[index].nome = e.target.value;
              setPessoas(nova);
            }}
          />
          <input
            placeholder="CPF"
            value={p.cpf}
            onChange={(e) => {
              const nova = [...pessoas];
              nova[index].cpf = e.target.value;
              setPessoas(nova);
            }}
          />
          <input
            type="file"
            accept="image/*"
            onChange={(e) => {
              if (e.target.files?.[0]) handleFoto(index, e.target.files[0]);
            }}
          />
          {pessoas.length > 1 && (
            <button onClick={() => removerPessoa(index)}>Remover</button>
          )}
        </div>
      ))}

      {pessoas.length < 2 && (
        <button onClick={adicionarPessoa}>Adicionar Cônjuge</button>
      )}

      <button onClick={gerarPlanilha} style={{ marginTop: '2rem', padding: '1rem 2rem' }}>
        Gerar e Baixar Excel
      </button>
    </div>
  );
}

export default App;
