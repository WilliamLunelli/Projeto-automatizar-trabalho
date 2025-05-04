const XLSX = require('xlsx');
const fs = require('fs');

// Função principal
function main() {
    try {
        console.log('=== SEPARADOR DE CÓDIGO E DESCRIÇÃO ===');

        // Obter argumentos da linha de comando
        const args = process.argv.slice(2);
        const arquivoEntrada = args[0] || 'dados_atuais.xlsx';
        const arquivoSaida = args[1] || 'dados_separados.xlsx';

        if (!fs.existsSync(arquivoEntrada)) {
            console.error(`Erro: O arquivo ${arquivoEntrada} não foi encontrado.`);
            console.log('Uso: node separador-codigo-descricao.js [arquivo_entrada.xlsx] [arquivo_saida.xlsx]');
            return;
        }

        console.log(`Lendo arquivo: ${arquivoEntrada}`);

        // Ler o arquivo Excel
        const workbook = XLSX.readFile(arquivoEntrada);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        // Converter para JSON
        const produtos = XLSX.utils.sheet_to_json(worksheet);
        console.log(`Lidos ${produtos.length} produtos.`);

        // Identificar a coluna de código
        const produtoAmostra = produtos[0];
        const colunas = Object.keys(produtoAmostra);

        console.log('Colunas encontradas:');
        colunas.forEach(col => console.log(`- ${col}`));

        // Tentar encontrar a coluna de código
        let colunaCodigo = null;
        const possiveisColunasCodigo = ['Código', 'Codigo', 'código', 'codigo', 'COD', 'CODIGO'];

        for (const col of possiveisColunasCodigo) {
            if (colunas.includes(col)) {
                colunaCodigo = col;
                break;
            }
        }

        if (!colunaCodigo) {
            console.error('Não foi possível identificar a coluna de código. Por favor, digite o nome da coluna:');
            // Em um ambiente interativo, aqui seria onde o usuário digitaria o nome da coluna
            colunaCodigo = colunas.find(col => col.toLowerCase().includes('cod'));

            if (!colunaCodigo) {
                console.error('Coluna de código não encontrada. Usando a primeira coluna como código.');
                colunaCodigo = colunas[0];
            }
        }

        console.log(`Usando coluna "${colunaCodigo}" como código.`);

        // Separar código e descrição e criar novos produtos
        const produtosSeparados = [];
        let codigosSemEspaco = 0;
        let codigosComEspaco = 0;

        for (let i = 0; i < produtos.length; i++) {
            const produto = produtos[i];
            const produtoNovo = { ...produto };

            if (produto[colunaCodigo]) {
                const codigo = String(produto[colunaCodigo]);

                // Verificar se o código contém espaços (o que indica que pode conter a descrição)
                if (codigo.includes(' ')) {
                    const partes = codigo.split(' ');
                    const codigoPuro = partes[0];
                    const descricao = partes.slice(1).join(' ');

                    produtoNovo[colunaCodigo] = codigoPuro;
                    produtoNovo['Descrição'] = descricao;
                    codigosComEspaco++;
                } else {
                    codigosSemEspaco++;
                }
            }

            produtosSeparados.push(produtoNovo);
        }

        console.log(`\nResultados:`);
        console.log(`- Produtos com código contendo espaços (descrição inferida): ${codigosComEspaco}`);
        console.log(`- Produtos sem espaços no código: ${codigosSemEspaco}`);

        // Criar nova planilha e salvar
        const novaPlanilha = XLSX.utils.json_to_sheet(produtosSeparados);
        const novoWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(novoWorkbook, novaPlanilha, 'Produtos');
        XLSX.writeFile(novoWorkbook, arquivoSaida);

        console.log(`\nArquivo salvo com sucesso em: ${arquivoSaida}`);
        console.log('Agora você pode usar o conversor original com este arquivo separado.');

    } catch (error) {
        console.error(`\nErro: ${error.message}`);
        console.error(error.stack);
    }
}

// Executar o programa
main();