#!/usr/bin/env node
// conversor-final-corrigido.js - Solução específica para o problema de descrição
const XLSX = require('xlsx');
const fs = require('fs');

/**
 * Função para normalizar strings (remover acentos, converter para minúsculo)
 * @param {string} texto - Texto a ser normalizado
 * @returns {string} - Texto normalizado
 */
function normalizar(texto) {
    if (!texto) return '';
    return String(texto)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim();
}

/**
 * Função para verificar se uma coluna existe com diferentes variações de nome
 * @param {Object} obj - Objeto onde procurar a coluna
 * @param {Array} possiveisNomes - Array de possíveis nomes da coluna
 * @returns {string|null} - Nome encontrado ou null
 */
function encontrarColuna(obj, possiveisNomes) {
    if (!obj || !possiveisNomes || !Array.isArray(possiveisNomes)) {
        return null;
    }

    // Verificar correspondência exata primeiro
    for (const nome of possiveisNomes) {
        if (obj.hasOwnProperty(nome)) {
            return nome;
        }
    }

    // Verificar por normalização
    const normalizados = possiveisNomes.map(n => normalizar(n));
    const todasChaves = Object.keys(obj);

    for (const chave of todasChaves) {
        const chaveNormalizada = normalizar(chave);
        const indice = normalizados.findIndex(n => n === chaveNormalizada);
        if (indice >= 0) {
            return chave;
        }
    }

    return null;
}

/**
 * Função para obter valor seguro de um objeto, buscando entre múltiplos nomes possíveis
 * @param {Object} obj - Objeto de onde obter o valor
 * @param {Array} possiveisNomes - Possíveis nomes da chave a buscar
 * @param {any} valorPadrao - Valor padrão caso não encontre
 * @returns {any} - Valor encontrado ou valor padrão
 */
function obterValorSeguro(obj, possiveisNomes, valorPadrao = '') {
    if (!obj || !possiveisNomes) {
        return valorPadrao;
    }

    // Se possiveisNomes não for array, trate como string única
    const nomes = Array.isArray(possiveisNomes) ? possiveisNomes : [possiveisNomes];

    const nomeEncontrado = encontrarColuna(obj, nomes);

    if (nomeEncontrado) {
        const valor = obj[nomeEncontrado];
        if (valor === undefined || valor === null) return valorPadrao;
        if (typeof valor === 'string' && valor.trim() === '') return valorPadrao;
        return valor;
    }

    return valorPadrao;
}

/**
 * Função para verificar se o valor está vazio
 * @param {any} valor - Valor a verificar
 * @returns {boolean} - true se estiver vazio, false caso contrário
 */
function estaVazio(valor) {
    if (valor === undefined || valor === null) return true;
    if (typeof valor === 'string' && valor.trim() === '') return true;
    return false;
}

/**
 * Função para converter números que podem estar em formatos diferentes
 * @param {any} valor - Valor a ser convertido
 * @param {number|null} valorPadrao - Valor padrão caso inválido
 * @returns {number|null} - Valor numérico ou null
 */
function parseNumero(valor, valorPadrao = null) {
    if (valor === undefined || valor === null || valor === '') return valorPadrao;
    try {
        if (typeof valor === 'number') return isNaN(valor) ? valorPadrao : valor;

        // Lidar com formato brasileiro (vírgula como separador decimal)
        const valorStr = String(valor).replace(/\./g, '').replace(',', '.');
        const valorConvertido = Number(valorStr);

        return isNaN(valorConvertido) ? valorPadrao : valorConvertido;
    } catch (error) {
        console.error(`Erro ao converter valor numérico: ${valor}`, error);
        return valorPadrao;
    }
}

/**
 * Função para converter um produto do formato atual para o novo formato
 * @param {Object} produtoAtual - Produto no formato atual
 * @param {Object} mapaColunasEncontradas - Mapa de todas as colunas encontradas
 * @param {number} index - Índice do produto para referência em logs
 * @returns {Object} - Produto no novo formato
 */
function converterProduto(produtoAtual, mapaColunasEncontradas, index) {
    try {
        // Garantir que mapaColunasEncontradas existe
        const mapa = mapaColunasEncontradas || {};

        // Extrair catálogo e código
        const catalogo = obterValorSeguro(produtoAtual, mapa.catalogo || ['Catálogo', 'Catalogo']);
        let codigo = obterValorSeguro(produtoAtual, mapa.codigo || ['Código', 'Codigo']);

        // Tratar código como string
        codigo = codigo ? String(codigo) : '';

        // Tentar obter descrição de diferentes maneiras
        let descricao = obterValorSeguro(produtoAtual, mapa.descricao || ['Descrição', 'Descricao', 'Descr']);

        // Se não houver descrição e o código parecer conter a descrição
        if (estaVazio(descricao) && !estaVazio(codigo) && codigo.includes(' ')) {
            const partes = codigo.split(' ');
            const codigoPuro = partes[0];
            descricao = partes.slice(1).join(' ');
            console.log(`\nInferindo descrição para produto ${index + 1} (Catálogo: ${catalogo}):`);
            console.log(`- Código original: "${codigo}"`);
            console.log(`- Código puro: "${codigoPuro}"`);
            console.log(`- Descrição inferida: "${descricao}"`);

            // Atualizar o código para usar apenas a parte numérica
            codigo = codigoPuro;
        }

        // Extrair outros campos com tratamento seguro
        const unidade = obterValorSeguro(produtoAtual, mapa.unidade || ['Unidade']);
        const ncm = obterValorSeguro(produtoAtual, mapa.ncm || ['NCM', 'Classificação Fiscal', 'Classificacao Fiscal']);
        const precoVarejo = obterValorSeguro(produtoAtual, mapa.precoVarejo || ['Preço Varejo', 'Preco Varejo']);
        const precoAtacado = obterValorSeguro(produtoAtual, mapa.precoAtacado || ['Preço Atacado', 'Preco Atacado']);
        const precoPromocao = obterValorSeguro(produtoAtual, mapa.precoPromocao || ['Preço Promoção', 'Preco Promocao']);
        const estoque = obterValorSeguro(produtoAtual, mapa.estoque || ['Saldo Estoque', 'Estoque']);
        const precoCompra = obterValorSeguro(produtoAtual, mapa.precoCompra || ['Preço Compra', 'Preco Compra']);
        const codigoOriginal = obterValorSeguro(produtoAtual, mapa.codigoOriginal || ['Código Original', 'Codigo Original']);
        const fornecedor = obterValorSeguro(produtoAtual, mapa.fornecedor || ['Fornecedor']);
        const endereco = obterValorSeguro(produtoAtual, mapa.endereco || ['Endereço', 'Endereco']);
        const endereco2 = obterValorSeguro(produtoAtual, mapa.endereco2 || ['Endereço 2', 'Endereco 2']);
        const garantia = obterValorSeguro(produtoAtual, mapa.garantia || ['Garantia']);
        const pendencia = obterValorSeguro(produtoAtual, mapa.pendencia || ['Pendência', 'Pendencia']);
        const linha = obterValorSeguro(produtoAtual, mapa.linha || ['Linha']);
        const grupo = obterValorSeguro(produtoAtual, mapa.grupo || ['Grupo']);

        // Criando observações com os preços especiais
        let observacoes = '';
        if (!estaVazio(precoAtacado)) {
            observacoes += `Preço atacado: ${precoAtacado}; `;
        }
        if (!estaVazio(precoPromocao)) {
            observacoes += `Preço promoção: ${precoPromocao}; `;
        }

        // Criando informações adicionais
        let infoAdicionais = '';
        if (!estaVazio(endereco2)) {
            infoAdicionais += `Endereço 2: ${endereco2}; `;
        }
        if (!estaVazio(pendencia)) {
            infoAdicionais += `Pendência: ${pendencia}; `;
        }

        // Mapeando o produto para o novo formato com valores seguros
        return {
            'ID': index + 1,
            'Código': codigo,
            'Descrição': descricao,
            'Unidade': unidade,
            'NCM': ncm,
            'Origem': '',
            'Preço': parseNumero(precoVarejo, 0),
            'Valor IPI fixo': null,
            'Observações': observacoes,
            'Situação': '',
            'Estoque': parseNumero(estoque, 0),
            'Preço de custo': parseNumero(precoCompra, 0),
            'Cód no fornecedor': codigoOriginal,
            'Fornecedor': fornecedor,
            'Localização': endereco,
            'Estoque maximo': null,
            'Estoque minimo': null,
            'Peso líquido (Kg)': null,
            'Peso bruto (Kg)': null,
            'GTIN/EAN': '',
            'GTIN/EAN da embalagem': '',
            'Largura do Produto': null,
            'Altura do Produto': null,
            'Profundidade do produto': null,
            'Data Validade': '',
            'Descrição do Produto no Fornecedor': '',
            'Descrição Complementar': '',
            'Itens p/ caixa': null,
            'Produto Variação': '',
            'Tipo Produção': '',
            'Classe de enquadramento do IPI': '',
            'Código da lista de serviços': '',
            'Tipo do item': '',
            'Grupo de Tags/Tags': '',
            'Tributos': '',
            'Código Pai': '',
            'Código Integração': '',
            'Grupo de produtos': linha,
            'Marca': '',
            'CEST': '',
            'Volumes': null,
            'Descrição Curta': '',
            'Cross-Docking': '',
            'URL Imagens Externas': '',
            'Link Externo': '',
            'Meses Garantia no Fornecedor': garantia,
            'Clonar dados do pai': '',
            'Condição do produto': '',
            'Frete Grátis': '',
            'Número FCI': '',
            'Vídeo': '',
            'Departamento': grupo,
            'Unidade de medida': '',
            'Preço de compra': parseNumero(precoCompra, 0),
            'Valor base ICMS ST para retenção': null,
            'Valor ICMS ST para retenção': null,
            'Valor ICMS próprio do substituto': null,
            'Categoria do produto': '',
            'Informações Adicionais': infoAdicionais
        };
    } catch (error) {
        throw new Error(`Erro ao converter produto ${index + 1}: ${error.message}`);
    }
}

/**
 * Função para fazer diagnóstico inicial das colunas encontradas no arquivo
 * @param {Array} produtos - Array de produtos lidos do Excel
 * @returns {Object} - Mapa de colunas encontradas
 */
function diagnosticarColunas(produtos) {
    if (!produtos || produtos.length === 0) {
        console.log("Nenhum produto encontrado para diagnóstico.");
        return {};
    }

    // Extrair todas as colunas
    const todasColunas = new Set();
    produtos.forEach(produto => {
        if (produto && typeof produto === 'object') {
            Object.keys(produto).forEach(chave => todasColunas.add(chave));
        }
    });

    console.log(`\nEncontradas ${todasColunas.size} colunas no arquivo:`);
    Array.from(todasColunas).forEach(coluna => {
        console.log(`- ${coluna}`);
    });

    // Se não há produtos ou o primeiro produto não é um objeto, retornar vazio
    if (produtos.length === 0 || !produtos[0] || typeof produtos[0] !== 'object') {
        return {};
    }

    // Mapeamento de colunas importantes
    const mapeamento = {
        catalogo: encontrarColuna(produtos[0], ['Catálogo', 'Catalogo']),
        codigo: encontrarColuna(produtos[0], ['Código', 'Codigo']),
        descricao: encontrarColuna(produtos[0], ['Descrição', 'Descricao']),
        unidade: encontrarColuna(produtos[0], ['Unidade']),
        ncm: encontrarColuna(produtos[0], ['NCM', 'Classificação Fiscal', 'Classificacao Fiscal']),
        precoVarejo: encontrarColuna(produtos[0], ['Preço Varejo', 'Preco Varejo']),
        precoAtacado: encontrarColuna(produtos[0], ['Preço Atacado', 'Preco Atacado']),
        precoPromocao: encontrarColuna(produtos[0], ['Preço Promoção', 'Preco Promocao']),
        estoque: encontrarColuna(produtos[0], ['Saldo Estoque', 'Estoque']),
        precoCompra: encontrarColuna(produtos[0], ['Preço Compra', 'Preco Compra']),
        endereco: encontrarColuna(produtos[0], ['Endereço', 'Endereco']),
        endereco2: encontrarColuna(produtos[0], ['Endereço 2', 'Endereco 2']),
        fornecedor: encontrarColuna(produtos[0], ['Fornecedor']),
        garantia: encontrarColuna(produtos[0], ['Garantia']),
        pendencia: encontrarColuna(produtos[0], ['Pendência', 'Pendencia']),
        linha: encontrarColuna(produtos[0], ['Linha']),
        grupo: encontrarColuna(produtos[0], ['Grupo'])
    };

    console.log("\nMapeamento de colunas encontradas:");
    Object.entries(mapeamento).forEach(([chave, valor]) => {
        console.log(`- ${chave}: ${valor || 'NÃO ENCONTRADO'}`);
    });

    // Verificar se o código contém a descrição
    if (mapeamento.codigo) {
        const colunaCodigo = mapeamento.codigo;
        const colunaDescricao = mapeamento.descricao;

        // Verificar se há descrições vazias ou se a coluna de descrição não existe
        const temProblemaDescricao = !colunaDescricao ||
            produtos.some(p => !p[colunaDescricao] && p && p[colunaCodigo]);

        if (temProblemaDescricao) {
            console.log("\nAnalisando campo Código para verificar se contém descrição...");

            // Pegar uma amostra de produtos não vazios
            const produtosValidos = produtos.filter(p => p && p[colunaCodigo]);
            const amostra = produtosValidos.slice(0, Math.min(5, produtosValidos.length));

            const codigosComEspaco = amostra.filter(p => {
                const codigo = p[colunaCodigo];
                return codigo &&
                    typeof codigo === 'string' &&
                    codigo.includes(' ');
            });

            if (codigosComEspaco.length > 0) {
                console.log(`\n✅ Encontrados ${codigosComEspaco.length} produtos (na amostra) com código contendo espaços.`);
                console.log("Exemplo de código contendo descrição:");

                const exemplo = codigosComEspaco[0];
                const codigo = exemplo[colunaCodigo];
                const partes = codigo.split(' ');

                console.log(`- Código original: "${codigo}"`);
                console.log(`- Possível código puro: "${partes[0]}"`);
                console.log(`- Possível descrição: "${partes.slice(1).join(' ')}"`);

                console.log("\n⚠️ Assumindo que o campo Código contém tanto o código quanto a descrição!");
                console.log("O conversor fará a separação automática.");
            }
        }
    }

    return mapeamento;
}

/**
 * Função principal que executa a conversão
 */
function iniciar() {
    try {
        console.log('===================================');
        console.log('  CONVERSOR DE TABELAS EXCEL');
        console.log('  Versão Final com Correções Específicas');
        console.log('===================================');
        console.log('\nEste script converte sua tabela do formato atual para o novo formato.\n');

        // Obter os argumentos da linha de comando ou usar valores padrão
        const args = process.argv.slice(2);
        const arquivoEntrada = args[0] || 'dados_atuais.xlsx';
        const arquivoSaida = args[1] || 'dados_convertidos.xlsx';
        const modoDebug = args.includes('--debug') || args.includes('-d');

        // Verificar se o arquivo de entrada existe
        if (!fs.existsSync(arquivoEntrada)) {
            console.error(`\nErro: O arquivo ${arquivoEntrada} não foi encontrado.`);
            console.log('\nUso: node conversor-final-corrigido.js [arquivo_entrada.xlsx] [arquivo_saida.xlsx] [--debug]');
            return;
        }

        console.log(`Arquivo de entrada: ${arquivoEntrada}`);
        console.log(`Arquivo de saída: ${arquivoSaida}`);
        console.log(`Modo debug: ${modoDebug ? 'Ativado' : 'Desativado'}`);
        console.log('\nIniciando conversão...');

        // Lendo o arquivo Excel de entrada com tratamento de erros
        let workbookEntrada;
        try {
            workbookEntrada = XLSX.readFile(arquivoEntrada, {
                cellStyles: true,
                cellDates: true,
                cellNF: true,
                raw: false, // Para ter um processamento mais confiável de texto e números
                type: 'binary'
            });
        } catch (error) {
            console.error(`\nErro ao ler o arquivo Excel: ${error.message}`);
            console.log('Tentando ler novamente com configurações alternativas...');

            workbookEntrada = XLSX.readFile(arquivoEntrada, {
                cellStyles: false,
                cellDates: false,
                cellNF: false,
                raw: true
            });
        }

        if (!workbookEntrada || !workbookEntrada.SheetNames || workbookEntrada.SheetNames.length === 0) {
            throw new Error('Não foi possível ler o arquivo Excel corretamente.');
        }

        const sheetNameEntrada = workbookEntrada.SheetNames[0];
        const worksheetEntrada = workbookEntrada.Sheets[sheetNameEntrada];

        if (!worksheetEntrada) {
            throw new Error(`Planilha '${sheetNameEntrada}' não encontrada no arquivo.`);
        }

        // Convertendo para JSON com tratamento de erros
        let produtosAtuaisRaw;
        try {
            produtosAtuaisRaw = XLSX.utils.sheet_to_json(worksheetEntrada, {
                raw: false,      // Obter valores formatados
                defval: '',      // Valor padrão para células vazias
                blankrows: false // Ignorar linhas em branco
            });
        } catch (error) {
            console.error(`\nErro ao converter planilha para JSON: ${error.message}`);
            console.log('Tentando método alternativo...');

            // Método alternativo: ler como array e converter manualmente
            const dadosRaw = XLSX.utils.sheet_to_json(worksheetEntrada, {
                header: 1,
                raw: true
            });

            if (!dadosRaw || dadosRaw.length <= 1) {
                throw new Error('A planilha não contém dados suficientes.');
            }

            const cabecalhos = dadosRaw[0];
            produtosAtuaisRaw = dadosRaw.slice(1).map(linha => {
                const produto = {};
                linha.forEach((valor, index) => {
                    if (index < cabecalhos.length && cabecalhos[index]) {
                        produto[cabecalhos[index]] = valor;
                    }
                });
                return produto;
            });
        }

        console.log(`\nLidos ${produtosAtuaisRaw.length} produtos do arquivo de entrada`);

        // Fazer diagnóstico das colunas encontradas
        const mapaColunasEncontradas = diagnosticarColunas(produtosAtuaisRaw);

        // Processando os produtos
        const produtosNovos = [];
        let sucessos = 0;
        let falhas = 0;
        const erros = [];
        const avisosDescricaoVazia = [];
        let descricoesInferidas = 0;

        for (const [index, produtoRaw] of produtosAtuaisRaw.entries()) {
            try {
                // Verificar se o produto é válido
                if (!produtoRaw || typeof produtoRaw !== 'object') {
                    throw new Error('Produto inválido ou vazio');
                }

                // Converter o produto
                const produtoNovo = converterProduto(produtoRaw, mapaColunasEncontradas, index);

                // Verificar se a descrição foi inferida do código
                const colunaCodigo = mapaColunasEncontradas.codigo;
                const colunaDescricao = mapaColunasEncontradas.descricao;

                if (colunaCodigo &&
                    (!colunaDescricao || estaVazio(produtoRaw[colunaDescricao])) &&
                    !estaVazio(produtoNovo.Descrição)) {
                    descricoesInferidas++;
                }

                // Verificar se ainda tem problema de descrição após conversão
                if (estaVazio(produtoNovo.Descrição)) {
                    avisosDescricaoVazia.push({
                        indice: index + 1,
                        codigo: produtoNovo.Código || `Item #${index + 1}`
                    });
                }

                // Adicionando à lista de produtos novos
                produtosNovos.push(produtoNovo);
                sucessos++;
            } catch (error) {
                console.error(`\nErro ao processar o produto ${index + 1}:`, error.message);
                erros.push(`Produto ${index + 1}: ${error.message}`);
                falhas++;
            }

            // Mostrar progresso a cada 100 produtos
            if ((index + 1) % 100 === 0 || index + 1 === produtosAtuaisRaw.length) {
                process.stdout.write(`\rProcessando: ${index + 1}/${produtosAtuaisRaw.length} produtos`);
            }
        }

        console.log('\n\nResultados da conversão:');
        console.log(`✅ ${sucessos} produtos processados com sucesso`);
        console.log(`❌ ${falhas} produtos com falhas durante o processamento`);
        console.log(`🔍 ${descricoesInferidas} descrições foram inferidas a partir do campo Código`);

        // Mostrar avisos de descrição vazia
        if (avisosDescricaoVazia.length > 0) {
            console.warn(`\n⚠️ Atenção: ${avisosDescricaoVazia.length} produtos ainda estão com o campo Descrição vazio!`);
            console.warn('Produtos afetados (primeiros 10): ' + avisosDescricaoVazia.slice(0, 10).map(item => item.codigo).join(', ') +
                (avisosDescricaoVazia.length > 10 ? ` e mais ${avisosDescricaoVazia.length - 10}...` : ''));
        } else {
            console.log('\n✅ Todos os produtos têm descrição válida!');
        }

        // Se houver erros, mostrar detalhes
        if (falhas > 0) {
            console.log('\nDetalhe dos erros (primeiros 10):');
            erros.slice(0, 10).forEach((erro, i) => {
                console.log(`${i + 1}. ${erro}`);
            });

            if (erros.length > 10) {
                console.log(`... e mais ${erros.length - 10} erros.`);
            }
        }

        // Verificar se há produtos para salvar
        if (produtosNovos.length === 0) {
            console.warn('\n⚠️ Atenção: Nenhum produto foi processado com sucesso para salvar!');
            return;
        }

        // Criando uma nova planilha com os dados convertidos
        try {
            const worksheetSaida = XLSX.utils.json_to_sheet(produtosNovos);
            const workbookSaida = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbookSaida, worksheetSaida, 'Produtos');

            // Salvando o arquivo de saída
            XLSX.writeFile(workbookSaida, arquivoSaida);
            console.log(`\n✨ Conversão concluída com sucesso! Arquivo salvo em: ${arquivoSaida}`);
        } catch (error) {
            console.error(`\nErro ao salvar o arquivo de saída: ${error.message}`);

            // Tentar salvar em outro formato
            try {
                console.log('Tentando salvar em formato alternativo (CSV)...');
                const csvSaida = arquivoSaida.replace(/\.xlsx?$/i, '.csv');
                const csvContent = XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(produtosNovos));
                fs.writeFileSync(csvSaida, csvContent, 'utf8');
                console.log(`Arquivo CSV salvo com sucesso em: ${csvSaida}`);
            } catch (csvError) {
                console.error(`Também não foi possível salvar como CSV: ${csvError.message}`);
                throw error; // Relancar erro original
            }
        }

        console.log('\n👋 Obrigado por usar o Conversor de Tabelas Excel!');

    } catch (error) {
        console.error('\n❌ Erro durante a execução do script:', error.message);
        if (error.stack) {
            console.error('Stack trace:', error.stack);
        }
    }
}

// Iniciar o script
iniciar();