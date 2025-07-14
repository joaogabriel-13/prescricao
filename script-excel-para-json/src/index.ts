import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Converte uma string de dicionário/lista Python para um objeto JSON.
 * Lida com None, True, False e aspas simples.
 * @param pyString A string no formato Python.
 * @returns Um objeto JSON ou a string original em caso de erro.
 */
function pythonStringToJson(pyString: string): any {
    if (!pyString || !pyString.trim()) {
        return ""; // Retorna string vazia se a entrada for vazia.
    }

    try {
        // 1. Substituições globais de palavras-chave Python para JSON
        let jsonString = pyString
            .replace(/\bNone\b/g, 'null')
            .replace(/\bTrue\b/g, 'true')
            .replace(/\bFalse\b/g, 'false');

        // 2. Substituição de aspas simples por duplas de forma segura
        // Esta regex visa substituir aspas simples que delimitam chaves e valores,
        // tentando ignorar aspas dentro de valores de string.
        // '([^']*)' captura o conteúdo entre aspas.
        // A substituição por `"$1"` envolve o conteúdo capturado com aspas duplas.
        // Isso é uma melhoria, mas ainda pode falhar em casos complexos (ex: aspas escapadas).
        jsonString = jsonString.replace(/'/g, '"');

        // 3. Tenta fazer o parse
        return JSON.parse(jsonString);
    } catch (e) {
        // Se o parse inicial falhar, pode ser por causa de aspas duplas dentro de strings.
        // Ex: "{'key': 'value with "quotes"'}" se torna '{"key": "value with "quotes""}' -> Inválido
        // Tentativa de corrigir aspas duplas internas que não foram escapadas.
        try {
            let correctedString = pyString
                .replace(/\bNone\b/g, 'null')
                .replace(/\bTrue\b/g, 'true')
                .replace(/\bFalse\b/g, 'false')
                .replace(/\\/g, '\\\\') // Escapa barras invertidas
                .replace(/'/g, '"') // Troca aspas simples por duplas
                .replace(/""/g, '\"'); // Tenta escapar aspas duplas que ficaram juntas

            return JSON.parse(correctedString);
        } catch (error) {
            console.error(`Falha ao converter a string Python para JSON após múltiplas tentativas. String original: ${pyString}`);
            // Retorna a string original ou vazia para não quebrar o resto do processo
            return pyString;
        }
    }
}


// __dirname aqui se refere ao diretório do arquivo JS compilado
// Ex: /workspaces/prescricao/script-excel-para-json/dist/

// Sobe um nível para chegar a /workspaces/prescricao/script-excel-para-json/
const baseDir = path.join(__dirname, '..');

const inputDir = path.join(baseDir, 'input'); // Deve resultar em /workspaces/prescricao/script-excel-para-json/input/
const outputDir = path.join(baseDir, 'output'); // Deve resultar em /workspaces/prescricao/script-excel-para-json/output/
const excelFileName = 'prescricoes.xlsx';
const excelFilePath = path.join(inputDir, excelFileName);

interface PlanilhaItem {
    [key: string]: any;
}

function converterExcelParaJson() {
    console.log(`Tentando ler o arquivo Excel de: ${excelFilePath}`);
    console.log(`Diretório de entrada configurado para: ${inputDir}`);
    console.log(`Diretório de saída configurado para: ${outputDir}`);
    console.log(`__dirname (diretório do script em execução): ${__dirname}`);
    console.log(`baseDir (calculado para input/output): ${baseDir}`);

    // Verifica se o diretório de entrada existe
    if (!fs.existsSync(inputDir)) {
        console.error(`ERRO: Diretório de entrada não encontrado: ${inputDir}`);
        return;
    }

    // Verifica se o arquivo Excel existe
    if (!fs.existsSync(excelFilePath)) {
        console.error(`ERRO: Arquivo Excel não encontrado: ${excelFilePath}`);
        return;
    }

    // Garante que o diretório de saída exista e seja um diretório
    if (fs.existsSync(outputDir)) {
        // Se existe, verifica se é um diretório
        if (!fs.statSync(outputDir).isDirectory()) {
            console.error(`ERRO CRÍTICO: O caminho de saída '${outputDir}' existe, mas NÃO é um diretório. Por favor, remova ou renomeie este item e tente novamente.`);
            return; // Interrompe a execução
        }
        console.log(`Diretório de saída '${outputDir}' já existe e é um diretório.`);
    } else {
        // Se não existe, cria o diretório
        try {
            fs.mkdirSync(outputDir, { recursive: true });
            console.log(`Diretório de saída '${outputDir}' criado com sucesso.`);
        } catch (error) {
            console.error(`ERRO: Não foi possível criar o diretório de saída '${outputDir}':`, error);
            return; // Interrompe a execução
        }
    }

    try {
        // Lê o arquivo Excel
        const workbook = XLSX.readFile(excelFilePath);
        const sheetNames = workbook.SheetNames;

        console.log(`Planilhas encontradas: ${sheetNames.join(', ')}`);

        sheetNames.forEach((sheetName: string) => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData: PlanilhaItem[] = XLSX.utils.sheet_to_json<PlanilhaItem>(worksheet, { defval: "" });

            // Itera sobre cada linha do JSON para processamento de dados
            jsonData.forEach(row => {
                // Converte a coluna 'isCalculable' de string para booleano
                if (row.hasOwnProperty('isCalculable')) {
                    const value = String(row.isCalculable).trim().toLowerCase();
                    if (value === 'true') {
                        row.isCalculable = true;
                    } else if (value === 'false') {
                        row.isCalculable = false;
                    }
                }

                // Processa a coluna 'PrescricoesPadronizadasJSON' para converter string Python-like em objeto JSON
                if (row.hasOwnProperty('PrescricoesPadronizadasJSON')) {
                    const pyJsonString = row.PrescricoesPadronizadasJSON;
                    if (typeof pyJsonString === 'string' && pyJsonString.trim()) {
                        // Usa a nova função robusta para conversão
                        row.PrescricoesPadronizadasJSON = pythonStringToJson(pyJsonString);
                    }
                }
            });

            if (jsonData.length > 0) {
                const allKeys = Object.keys(jsonData[0]);
                const keysToRemove: string[] = [];
                allKeys.forEach(key => {
                    const isColumnEmpty = jsonData.every(row => row[key] === "" || row[key] === null || typeof row[key] === 'undefined');
                    if (isColumnEmpty && key.startsWith('__EMPTY')) {
                        keysToRemove.push(key);
                    }
                });
                if (keysToRemove.length > 0) {
                    jsonData.forEach(row => {
                        keysToRemove.forEach(key => {
                            delete row[key];
                        });
                    });
                    console.log(`Colunas vazias removidas da planilha '${sheetName}': ${keysToRemove.join(', ')}`);
                }
            }

            const safeSheetName = sheetName.replace(/[^a-zA-Z0-9_.-]/g, '_').toLowerCase();
            const jsonFileName = `${safeSheetName}.json`;
            const jsonFilePath = path.join(outputDir, jsonFileName);

            try {
                fs.writeFileSync(jsonFilePath, JSON.stringify(jsonData, null, 2));
                console.log(`SUCESSO: Planilha '${sheetName}' convertida para '${jsonFilePath}'`);
            } catch (error) {
                console.error(`ERRO: Não foi possível escrever o arquivo JSON ${jsonFilePath}:`, error);
            }
        });

        console.log('Conversão concluída.');

    } catch (error) {
        console.error('ERRO durante o processo de conversão:', error);
    }
}

converterExcelParaJson();