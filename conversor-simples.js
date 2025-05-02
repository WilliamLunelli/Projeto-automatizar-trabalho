#!/usr/bin/env node
// conversor-corrigido.js - Versão com tratamento de erros aprimorado
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

/**
 * Função para converter valores numéricos que podem vir como strings
 * @param {any} valor - Valor a ser convertido
 * @param {number|null} valorPadrao - Valor padrão caso seja inválido
 * @returns {number|null} - Valor numérico ou null
 */
function parseNumero(valor, valorPadrao = null) {
    if (valor === undefined || valor === null || valor === '') return valorPadrao;
    try {
        if (typeof valor === 'number') return isNaN(valor) ? valorPadrao : valor;
        // Tentar converter para número
        const valorConvertido = Number(String(valor).replace(',', '.'));
        return isNaN(valorConvertido) ? valorPadrao : valorConvertido;
    } catch (error) {
        console.error(`Erro ao converter valor numérico: ${valor}`, error);
        return valorPadrao;
    }
}

/**
 * Função segura para obter valores de objetos, com valor padrão se não existir
 * @param {Object} obj - Objeto de onde obter o valor
 * @param {string} chave - Chave do valor a ser obtido
 * @param {any} valorPadrao - Valor padrão caso a chave não exista
 * @returns {any} - Valor obtido ou valor padrão
 */
function obterValorSeguro(obj, chave, valorPadrao = '') {
    if (!obj || obj[chave] === undefined || obj[chave] === null) return valorPadrao;
    return obj[chave];
}

/**
 * Função para converter um produto do formato atual para o novo formato
 * @param {Object} produtoAtual - Produto no formato atual
 * @returns {Object} - Produto no novo formato
 */
function converterProduto(produtoAtual) {
    if (!produtoAtual) {
        throw new Error('Produto inválido ou vazio');
    }

    try {
        // Verificar campos críticos
        if (!obterValorSeguro(produtoAtual, 'Código')) {
            throw new Error('Campo Código é obrigatório');
        }

        if (!obterValorSeguro(produtoAtual, 'Descrição')) {
            console.warn(`Produto ${obterValorSeguro(produtoAtual, 'Código')}: Campo Descrição está vazio`);
        }

        // Criando observações com os preços especiais
        let observacoes = '';
        const precoAtacado = obterValorSeguro(produtoAtual, 'Preço Atacado');
        if (precoAtacado) {
            observacoes += `Preço atacado: ${precoAtacado}; `;
        }

        const precoPromocao = obterValorSeguro(produtoAtual, 'Preço Promoção');
        if (precoPromocao) {
            observacoes += `Preço promoção: ${precoPromocao}; `;
        }

        // Criando informações adicionais
        let infoAdicionais = '';
        const endereco2 = obterValorSeguro(produtoAtual, 'Endereço 2');
        if (endereco2) {
            infoAdicionais += `Endereço 2: ${endereco2}; `;
        }

        const pendencia = obterValorSeguro(produtoAtual, 'Pendência');
        if (pendencia) {
            infoAdicionais += `Pendência: ${pendencia}; `;
        }

        // Mapeando o produto para o novo formato com valores seguros
        return {
            'ID': null, // Será preenchido depois sequencialmente
            'Código': obterValorSeguro(produtoAtual, 'Código'),
            'Descrição': obterValorSeguro(produtoAtual, 'Descrição'),
            'Unidade': obterValorSeguro(produtoAtual, 'Unidade'),
            'NCM': obterValorSeguro(produtoAtual, 'Classificação Fiscal'),
            'Origem': '',
            'Preço': parseNumero(produtoAtual['Preço Varejo'], 0),
            'Valor IPI fixo': null,
            'Observações': observacoes,
            'Situação': '',
            'Estoque': parseNumero(produtoAtual['Saldo Estoque'], 0),
            'Preço de custo': parseNumero(produtoAtual['Preço Compra'], 0),
            'Cód no fornecedor': obterValorSeguro(produtoAtual, 'Código Original'),
            'Fornecedor': obterValorSeguro(produtoAtual, 'Fornecedor'),
            'Localização': obterValorSeguro(produtoAtual, 'Endereço'),
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
            'Grupo de produtos': obterValorSeguro(produtoAtual, 'Linha'),
            'Marca': '',
            'CEST': '',
            'Volumes': null,
            'Descrição Curta': '',
            'Cross-Docking': '',
            'URL Imagens Externas': '',
            'Link Externo': '',
            'Meses Garantia no Fornecedor': obterValorSeguro(produtoAtual, 'Garantia'),
            'Clonar dados do pai': '',
            'Condição do produto': '',
            'Frete Grátis': '',
            'Número FCI': '',
            'Vídeo': '',
            'Departamento': obterValorSeguro(produtoAtual, 'Grupo'),
            'Unidade de medida': '',
            'Preço de compra': parseNumero(produtoAtual['Preço Compra'], 0),
            'Valor base ICMS ST para retenção': null,
            'Valor ICMS ST para retenção': null,
            'Valor ICMS próprio do substituto': null,
            'Categoria do produto': '',
            'Informações Adicionais': infoAdicionais
        };
    } catch (error) {
        // Adicionar informações ao erro para facilitar a depuração
        const codigo = obterValorSeguro(produtoAtual, 'Código', 'desconhecido');
        throw new Error(`Erro ao converter produto ${codigo}: ${error.message}`);
    }
}

/**
 * Função para verificar a estrutura do arquivo Excel
 * @param {Array} dados - Dados lidos do Excel
 * @returns {Object} - Resultado da verificação
 */
function verificarEstrutura(dados) {
    if (!dados || !Array.isArray(dados) || dados.length === 0) {
        return {
            valido: false,
            mensagem: 'Arquivo vazio ou sem dados'
        };
    }

    const primeiroItem = dados[0];
    const colunasEsperadas = [
        'Código', 'Descrição', 'Classificação Fiscal', 'Preço Compra',
        'Preço Varejo', 'Saldo Estoque'
    ];

    const colunasAusentes = colunasEsperadas.filter(coluna =>
        !primeiroItem.hasOwnProperty(coluna)
    );

    if (colunasAusentes.length > 0) {
        return {
            valido: false,
            mensagem: `Colunas obrigatórias ausentes: ${colunasAusentes.join(', ')}`,
            colunasAusentes
        };
    }

    return { valido: true };
}

/**
 * Função principal que executa a conversão
 */
function iniciar() {
    try {
        console.log('===================================');
        console.log('  CONVERSOR DE TABELAS EXCEL');
        console.log('  Versão Corrigida com Tratamento de Erros');
        console.log('===================================');
        console.log('\nEste script converte sua tabela do formato atual para o novo formato.\n');

        // Obter os argumentos da linha de comando ou usar valores padrão
        const args = process.argv.slice(2);
        const arquivoEntrada = args[0] || 'dados_atuais.xlsx';
        const arquivoSaida = args[1] || 'dados_convertidos.xlsx';

        // Verificar se o arquivo de entrada existe
        if (!fs.existsSync(arquivoEntrada)) {
            console.error(`\nErro: O arquivo ${arquivoEntrada} não foi encontrado.`);
            console.log('\nUso: node conversor-corrigido.js [arquivo_entrada.xlsx] [arquivo_saida.xlsx]');
            return;
        }

        console.log(`Arquivo de entrada: ${arquivoEntrada}`);
        console.log(`Arquivo de saída: ${arquivoSaida}`);
        console.log('\nIniciando conversão...');

        // Lendo o arquivo Excel de entrada
        const workbookEntrada = XLSX.readFile(arquivoEntrada, {
            cellStyles: true,
            cellDates: true,
            cellNF: true,
            raw: false // Para ter um processamento mais confiável de texto e números
        });
        const sheetNameEntrada = workbookEntrada.SheetNames[0];
        const worksheetEntrada = workbookEntrada.Sheets[sheetNameEntrada];

        // Convertendo para JSON com cabeçalhos explícitos
        const opcoes = {
            raw: false, // Para obter valores formatados
            defval: '', // Valor padrão para células vazias
            header: 'A' // Usar primeira linha como cabeçalhos
        };

        // Primeiro pegamos os cabeçalhos
        const cabecalhos = [];
        const range = XLSX.utils.decode_range(worksheetEntrada['!ref']);

        for (let C = range.s.c; C <= range.e.c; ++C) {
            const endereco = XLSX.utils.encode_cell({ r: range.s.r, c: C });
            if (!worksheetEntrada[endereco]) continue;
            cabecalhos[C] = worksheetEntrada[endereco].v;
        }

        // Verificamos se todas as colunas necessárias estão presentes
        const colunasObrigatorias = ['Código', 'Descrição', 'Classificação Fiscal'];
        const colunasAusentes = colunasObrigatorias.filter(col =>
            !cabecalhos.includes(col)
        );

        if (colunasAusentes.length > 0) {
            console.error(`\nAtenção: As seguintes colunas obrigatórias não foram encontradas:`);
            console.error(`- ${colunasAusentes.join('\n- ')}`);
            console.error(`\nVerifique se os nomes das colunas estão escritos exatamente como esperado.`);

            // Mostrar os cabeçalhos encontrados para ajudar a depurar
            console.log('\nCabeçalhos encontrados no arquivo:');
            cabecalhos.filter(Boolean).forEach(cab => console.log(`- "${cab}"`));

            const continuar = true; // Em produção, você poderia perguntar ao usuário
            if (!continuar) {
                return;
            }

            console.log('\nContinuando com os cabeçalhos encontrados, mas poderá haver erros...\n');
        }

        // Agora lemos os dados com os cabeçalhos adequados
        const produtosAtuaisRaw = XLSX.utils.sheet_to_json(worksheetEntrada);
        console.log(`\nLidos ${produtosAtuaisRaw.length} produtos do arquivo de entrada`);

        // Verificar estrutura dos dados
        const verificacaoEstrutura = verificarEstrutura(produtosAtuaisRaw);
        if (!verificacaoEstrutura.valido) {
            console.error(`\nErro na estrutura do arquivo: ${verificacaoEstrutura.mensagem}`);

            if (verificacaoEstrutura.colunasAusentes) {
                console.error('As seguintes colunas estão ausentes:');
                console.error(`- ${verificacaoEstrutura.colunasAusentes.join('\n- ')}`);

                // Mostrar todas as colunas encontradas
                if (produtosAtuaisRaw.length > 0) {
                    console.log('\nColunas encontradas:');
                    console.log(`- ${Object.keys(produtosAtuaisRaw[0]).join('\n- ')}`);
                }
            }

            // Em produção, você poderia perguntar se deseja continuar mesmo assim
            console.log('\nContinuando mesmo com estrutura incompleta. Alguns campos podem ficar vazios.\n');
        }

        // Processando os produtos
        const produtosNovos = [];
        let sucessos = 0;
        let falhas = 0;
        const erros = [];
        const avisosDescricaoVazia = [];

        for (const [index, produtoRaw] of produtosAtuaisRaw.entries()) {
            try {
                // Verificar campo descrição vazio
                if (!obterValorSeguro(produtoRaw, 'Descrição')) {
                    avisosDescricaoVazia.push(index + 1);
                }

                // Convertendo para o novo formato
                const produtoNovo = converterProduto(produtoRaw);

                // Adicionando ID sequencial
                produtoNovo.ID = index + 1;

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

        // Mostrar avisos de descrição vazia
        if (avisosDescricaoVazia.length > 0) {
            console.warn(`\n⚠️ Atenção: ${avisosDescricaoVazia.length} produtos estão com o campo Descrição vazio!`);
            console.warn('Produtos afetados (primeiros 10): ' + avisosDescricaoVazia.slice(0, 10).join(', ') +
                (avisosDescricaoVazia.length > 10 ? ` e mais ${avisosDescricaoVazia.length - 10}...` : ''));
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
        console.error('Stack trace:', error.stack);
    }
}

// Iniciar o script
iniciar();