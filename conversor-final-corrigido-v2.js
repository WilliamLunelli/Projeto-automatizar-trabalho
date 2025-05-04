#!/usr/bin/env node
// conversor-final-corrigido-v2.js - Versão aprimorada para resolver problemas de descrição vazia
const XLSX = require('xlsx');
const fs = require('fs');

/**
 * Função para normalizar strings (remover acentos, converter para minúsculo)
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
 * Função para encontrar coluna com múltiplas variações de nome
 */
function encontrarColuna(obj, possiveisNomes) {
    if (!obj || !possiveisNomes || !Array.isArray(possiveisNomes)) {
        return null;
    }

    // Primeiro, procura correspondência exata
    for (const nome of possiveisNomes) {
        if (obj.hasOwnProperty(nome)) {
            return nome;
        }
    }

    // Depois procura por normalização
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
 * Função para obter valor seguro de um objeto
 */
function obterValorSeguro(obj, possiveisNomes, valorPadrao = '') {
    if (!obj || !possiveisNomes) {
        return valorPadrao;
    }

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
 */
function estaVazio(valor) {
    if (valor === undefined || valor === null) return true;
    if (typeof valor === 'string' && valor.trim() === '') return true;
    return false;
}

/**
 * Nova função: tentar extrair descrição de diferentes maneiras mais agressivas
 */
function extrairDescricao(produtoAtual, mapaColunasEncontradas, index) {
    // Método 1: Tentar colunas de descrição
    const colunasDescricao = [
        'Descrição', 'Descricao', 'Descr', 'Desc', 'Description',
        'Nome', 'Nome do Produto', 'Produto', 'Denominação', 'Denominacao'
    ];

    for (const col of colunasDescricao) {
        const descricao = obterValorSeguro(produtoAtual, [col]);
        if (!estaVazio(descricao)) {
            return descricao;
        }
    }

    // Método 2: Procurar em TODAS as colunas por valores que podem ser descrição
    const todasColunas = Object.keys(produtoAtual);

    for (const col of todasColunas) {
        // Pula colunas que claramente não são descrição
        if (normalizar(col).includes('codigo') ||
            normalizar(col).includes('preco') ||
            normalizar(col).includes('valor') ||
            normalizar(col).includes('ncm') ||
            normalizar(col).includes('unidade')) {
            continue;
        }

        const valor = produtoAtual[col];

        // Verifica se o valor parece uma descrição
        if (!estaVazio(valor) && typeof valor === 'string') {
            // Considera descrição valores com pelo menos 3 caracteres e que contém letras
            if (valor.length >= 3 && /[a-zA-Z]/.test(valor)) {
                console.log(`Descrição encontrada em coluna "${col}" para produto ${index + 1}: "${valor}"`);
                return valor;
            }
        }
    }

    // Método 3: Tentar extrair do código
    const codigo = obterValorSeguro(produtoAtual, ['Código', 'Codigo']);
    if (!estaVazio(codigo) && typeof codigo === 'string' && codigo.includes(' ')) {
        const partes = codigo.split(' ');
        const descricaoInferida = partes.slice(1).join(' ');
        console.log(`Descrição inferida do código para produto ${index + 1}: "${descricaoInferida}"`);
        return descricaoInferida;
    }

    // Método 4: Tentar concatenar valores de múltiplas colunas
    const possiveisCamposDescricao = [];

    for (const col of todasColunas) {
        const valor = produtoAtual[col];
        if (!estaVazio(valor) && typeof valor === 'string' && valor.length > 1 && /[a-zA-Z]/.test(valor)) {
            possiveisCamposDescricao.push({ coluna: col, valor: valor });
        }
    }

    if (possiveisCamposDescricao.length > 0) {
        // Pega o valor mais longo como descrição
        const melhorDescricao = possiveisCamposDescricao.reduce((max, atual) =>
            atual.valor.length > max.valor.length ? atual : max
        );

        console.log(`Descrição extraída da coluna "${melhorDescricao.coluna}" para produto ${index + 1}: "${melhorDescricao.valor}"`);
        return melhorDescricao.valor;
    }

    return '';
}

/**
 * Função para converter um produto do formato atual para o novo formato
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

        // NOVA ABORDAGEM: Usar função de extração de descrição mais agressiva
        let descricao = extrairDescricao(produtoAtual, mapaColunasEncontradas, index);

        // Se ainda não encontrou descrição, tenta uma última estratégia
        if (estaVazio(descricao)) {
            // Cria uma descrição baseada no ID ou posição
            descricao = `Produto ${index + 1}`;
            console.warn(`⚠️ Usando descrição padrão para produto ${index + 1}: "${descricao}"`);
        }

        // Resto da conversão permanece igual...
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
 * Função para converter números que podem estar em formatos diferentes
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
 * Função expandida para diagnóstico de colunas
 */
function diagnosticarColunas(produtos) {
    if (!produtos || produtos.length === 0) {
        console.log("Nenhum produto encontrado para diagnóstico.");
        return {};
    }

    // Extrair todas as colunas
    const todasColunas = new Set();
    const contagemColunas = {};
    const exemplosColunas = {};

    produtos.forEach((produto, idx) => {
        if (produto && typeof produto === 'object') {
            Object.entries(produto).forEach(([chave, valor]) => {
                todasColunas.add(chave);

                // Conta quantas vezes cada coluna aparece com valor não vazio
                if (!estaVazio(valor)) {
                    contagemColunas[chave] = (contagemColunas[chave] || 0) + 1;

                    // Guarda exemplos de valores
                    if (!exemplosColunas[chave]) {
                        exemplosColunas[chave] = [];
                    }
                    if (exemplosColunas[chave].length < 3) {
                        exemplosColunas[chave].push(valor);
                    }
                }
            });
        }
    });

    console.log(`\n=== DIAGNÓSTICO DE COLUNAS ===`);
    console.log(`Total de produtos: ${produtos.length}`);
    console.log(`Total de colunas encontradas: ${todasColunas.size}`);

    console.log('\nDetalhes das colunas:');
    Array.from(todasColunas).forEach(coluna => {
        const contagem = contagemColunas[coluna] || 0;
        const percentual = Math.round((contagem / produtos.length) * 100);
        const exemplos = exemplosColunas[coluna] || [];

        console.log(`- "${coluna}": ${contagem} valores (${percentual}%)`);
        if (exemplos.length > 0) {
            console.log(`  Exemplos: ${exemplos.map(e => `"${e}"`).join(', ')}`);
        }
    });

    // Identificar possíveis colunas de descrição
    console.log('\n=== ANÁLISE DE POSSÍVEIS COLUNAS DE DESCRIÇÃO ===');
    const possiveisColunas = [];

    Array.from(todasColunas).forEach(coluna => {
        const normalizada = normalizar(coluna);
        const contagem = contagemColunas[coluna] || 0;
        const percentual = Math.round((contagem / produtos.length) * 100);
        const exemplos = exemplosColunas[coluna] || [];

        // Verifica se a coluna pode ser descrição por nome ou por conteúdo
        const pareceColunaDescricao = normalizada.includes('descr') ||
            normalizada.includes('desc') ||
            normalizada.includes('nome') ||
            normalizada.includes('produto') ||
            normalizada.includes('denominacao');

        const pareceConteudoDescricao = exemplos.some(exemplo =>
            exemplo &&
            typeof exemplo === 'string' &&
            exemplo.length > 3 &&
            /[a-zA-Z]/.test(exemplo)
        );

        if (pareceColunaDescricao || pareceConteudoDescricao) {
            possiveisColunas.push({
                coluna: coluna,
                contagem: contagem,
                percentual: percentual,
                exemplos: exemplos,
                scoring: pareceColunaDescricao ? 10 : 0 + (pareceConteudoDescricao ? 5 : 0) + (contagem / produtos.length) * 10
            });
        }
    });

    // Ordena por probabilidade de ser descrição
    possiveisColunas.sort((a, b) => b.scoring - a.scoring);

    console.log('\nPrincipais candidatas a coluna de descrição:');
    possiveisColunas.slice(0, 5).forEach((info, idx) => {
        console.log(`${idx + 1}. "${info.coluna}" - ${info.contagem} valores (${info.percentual}%)`);
        console.log(`   Exemplos: ${info.exemplos.slice(0, 2).map(e => `"${e}"`).join(', ')}`);
    });

    // Resto do diagnóstico permanece igual...
    const mapeamento = {
        catalogo: encontrarColuna(produtos[0], ['Catálogo', 'Catalogo']),
        codigo: encontrarColuna(produtos[0], ['Código', 'Codigo']),
        descricao: possiveisColunas[0] ? possiveisColunas[0].coluna : encontrarColuna(produtos[0], ['Descrição', 'Descricao']),
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

    console.log("\n=== MAPEAMENTO FINAL ===");
    Object.entries(mapeamento).forEach(([chave, valor]) => {
        console.log(`- ${chave}: ${valor || 'NÃO ENCONTRADO'}`);
    });

    return mapeamento;
}

/**
 * Função principal que executa a conversão
 */
function iniciar() {
    try {
        console.log('===================================');
        console.log('  CONVERSOR DE TABELAS EXCEL V2');
        console.log('  Resolução de problemas de descrição');
        console.log('===================================');
        console.log('\nEste script converte sua tabela com tratamento especial para descrições.\n');

        // Obter os argumentos da linha de comando ou usar valores padrão
        const args = process.argv.slice(2);
        const arquivoEntrada = args[0] || 'dados_atuais.xlsx';
        const arquivoSaida = args[1] || 'dados_convertidos.xlsx';
        const modoDebug = args.includes('--debug') || args.includes('-d');

        // Verificar se o arquivo de entrada existe
        if (!fs.existsSync(arquivoEntrada)) {
            console.error(`\nErro: O arquivo ${arquivoEntrada} não foi encontrado.`);
            console.log('\nUso: node conversor-final-corrigido-v2.js [arquivo_entrada.xlsx] [arquivo_saida.xlsx] [--debug]');
            return;
        }

        console.log(`Arquivo de entrada: ${arquivoEntrada}`);
        console.log(`Arquivo de saída: ${arquivoSaida}`);
        console.log(`Modo debug: ${modoDebug ? 'Ativado' : 'Desativado'}`);
        console.log('\nIniciando conversão...');

        // Lendo o arquivo Excel de entrada
        const workbookEntrada = XLSX.readFile(arquivoEntrada, {
            cellStyles: true,
            cellDates: true,
            cellNF: true,
            raw: false,
            type: 'binary'
        });

        if (!workbookEntrada || !workbookEntrada.SheetNames || workbookEntrada.SheetNames.length === 0) {
            throw new Error('Não foi possível ler o arquivo Excel corretamente.');
        }

        const sheetNameEntrada = workbookEntrada.SheetNames[0];
        const worksheetEntrada = workbookEntrada.Sheets[sheetNameEntrada];

        if (!worksheetEntrada) {
            throw new Error(`Planilha '${sheetNameEntrada}' não encontrada no arquivo.`);
        }

        // Convertendo para JSON
        const produtosAtuaisRaw = XLSX.utils.sheet_to_json(worksheetEntrada, {
            raw: false,
            defval: '',
            blankrows: false
        });

        console.log(`\nLidos ${produtosAtuaisRaw.length} produtos do arquivo de entrada`);

        // Fazer diagnóstico expandido das colunas encontradas
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

            // Mostrar progresso
            if ((index + 1) % 100 === 0 || index + 1 === produtosAtuaisRaw.length) {
                process.stdout.write(`\rProcessando: ${index + 1}/${produtosAtuaisRaw.length} produtos`);
            }
        }

        console.log('\n\n=== RESULTADOS DA CONVERSÃO ===');
        console.log(`✅ ${sucessos} produtos processados com sucesso`);
        console.log(`❌ ${falhas} produtos com falhas durante o processamento`);

        // Mostrar avisos de descrição vazia
        if (avisosDescricaoVazia.length > 0) {
            console.warn(`\n⚠️ Atenção: ${avisosDescricaoVazia.length} produtos ficaram com o campo Descrição vazio!`);

            if (avisosDescricaoVazia.length <= 20) {
                console.warn('Produtos afetados:');
                avisosDescricaoVazia.forEach(item => {
                    console.warn(`- ${item.codigo}`);
                });
            } else {
                console.warn('Produtos afetados (primeiros 20):');
                avisosDescricaoVazia.slice(0, 20).forEach(item => {
                    console.warn(`- ${item.codigo}`);
                });
                console.warn(`... e mais ${avisosDescricaoVazia.length - 20} produtos`);
            }
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
        const worksheetSaida = XLSX.utils.json_to_sheet(produtosNovos);
        const workbookSaida = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbookSaida, worksheetSaida, 'Produtos');

        // Salvando o arquivo de saída
        XLSX.writeFile(workbookSaida, arquivoSaida);
        console.log(`\n✨ Conversão concluída com sucesso! Arquivo salvo em: ${arquivoSaida}`);

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