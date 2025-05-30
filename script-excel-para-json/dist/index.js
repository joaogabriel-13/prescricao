"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
const XLSX = __importStar(require("xlsx"));
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
// __dirname aqui se refere ao diretório do arquivo JS compilado
// Ex: /workspaces/prescricao/script-excel-para-json/dist/
// Sobe um nível para chegar a /workspaces/prescricao/script-excel-para-json/
const baseDir = path.join(__dirname, '..');
const inputDir = path.join(baseDir, 'input'); // Deve resultar em /workspaces/prescricao/script-excel-para-json/input/
const outputDir = path.join(baseDir, 'output'); // Deve resultar em /workspaces/prescricao/script-excel-para-json/output/
const excelFileName = 'prescricoes.xlsx';
const excelFilePath = path.join(inputDir, excelFileName);
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
    }
    else {
        // Se não existe, cria o diretório
        try {
            fs.mkdirSync(outputDir, { recursive: true });
            console.log(`Diretório de saída '${outputDir}' criado com sucesso.`);
        }
        catch (error) {
            console.error(`ERRO: Não foi possível criar o diretório de saída '${outputDir}':`, error);
            return; // Interrompe a execução
        }
    }
    try {
        // Lê o arquivo Excel
        const workbook = XLSX.readFile(excelFilePath);
        const sheetNames = workbook.SheetNames;
        console.log(`Planilhas encontradas: ${sheetNames.join(', ')}`);
        sheetNames.forEach((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            if (jsonData.length > 0) {
                const allKeys = Object.keys(jsonData[0]);
                const keysToRemove = [];
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
            }
            catch (error) {
                console.error(`ERRO: Não foi possível escrever o arquivo JSON ${jsonFilePath}:`, error);
            }
        });
        console.log('Conversão concluída.');
    }
    catch (error) {
        console.error('ERRO durante o processo de conversão:', error);
    }
}
converterExcelParaJson();
