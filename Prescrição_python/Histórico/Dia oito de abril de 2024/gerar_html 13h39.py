# INÍCIO DO CÓDIGO COMPLETO (gerar_html.py) - VERSÃO FINAL CORRIGIDA
import pandas as pd
import html
import os
import re

# --- Configurações ---
PASTA_ATUAL = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_EXCEL = os.path.join(PASTA_ATUAL, 'prescricoes.xlsx')
ARQUIVO_HTML_SAIDA = os.path.join(PASTA_ATUAL, 'minhas_prescricoes.html')

COLUNAS_MEDICAMENTOS = ['NomeBusca', 'PrescricaoCompleta', 'Categoria', 'Doenca']
COLUNAS_GENERICAS = ['NomeBusca', 'ConteudoTexto']

CORES_ABAS = [
    '#a8dadc', '#f1faee', '#e63946', '#457b9d', '#1d3557',
    '#fca311', '#b7b7a4', '#d4a373', '#a2d2ff', '#ffafcc'
]
MAX_RECENTES = 7 # Número máximo de itens recentes

# --- Função Auxiliar ---
def sanitizar_nome(nome):
    # Função para criar IDs seguros para HTML/JS a partir dos nomes das planilhas
    nome = nome.lower(); nome = re.sub(r'[áàâãä]', 'a', nome); nome = re.sub(r'[éèêë]', 'e', nome); nome = re.sub(r'[íìîï]', 'i', nome); nome = re.sub(r'[óòôõö]', 'o', nome); nome = re.sub(r'[úùûü]', 'u', nome); nome = re.sub(r'[ç]', 'c', nome); nome = re.sub(r'[^a-z0-9\s-]', '', nome); nome = re.sub(r'\s+', '-', nome).strip('-'); return nome or "aba-generica"

# --- Templates HTML e JavaScript ---

HTML_INICIO = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Assistente Médico Rápido</title>
    <style>
        :root { /* Variáveis CSS */
            --primary-color: #2a6f97; --primary-hover: #1d5a7a; --success-color: #28a745;
            --light-bg: #f4f5f7; --dark-bg: #343a40; --card-bg: #fff; --card-border: #e9ecef;
            --text-color: #333; --text-muted: #6c757d; --border-color: #e9ecef;
            --shadow-sm: 0 2px 4px rgba(0,0,0,0.05); --shadow-md: 0 5px 15px rgba(0,0,0,0.07);
            --shadow-hover: 0 4px 8px rgba(0,0,0,0.1); --radius-sm: 4px; --radius-md: 8px; --radius-lg: 10px;
            --transition-speed: 0.3s;
        }
        [data-theme="dark"] { /* Tema Escuro */
            --light-bg: #212529; --dark-bg: #121416; --card-bg: #2a2d31; --card-border: #444;
            --text-color: #e9ecef; --text-muted: #adb5bd; --border-color: #444;
            --shadow-sm: 0 2px 4px rgba(0,0,0,0.2); --shadow-md: 0 5px 15px rgba(0,0,0,0.3); --shadow-hover: 0 4px 8px rgba(0,0,0,0.3);
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background-color: var(--light-bg); color: var(--text-color); transition: background-color var(--transition-speed), color var(--transition-speed); }
        .container { max-width: 950px; margin: 30px auto; background-color: var(--card-bg); padding: 30px; border-radius: var(--radius-lg); box-shadow: var(--shadow-md); transition: background-color var(--transition-speed), box-shadow var(--transition-speed); }
        .assinatura { text-align: center; width: 100%; order: -1; margin-bottom: 15px; font-size: 0.8em; color: var(--text-muted); padding-bottom: 15px; border-bottom: 1px solid var(--border-color); transition: color var(--transition-speed), border-color var(--transition-speed); }
        .assinatura a { color: var(--primary-color); text-decoration: none; transition: color var(--transition-speed); }
        .assinatura a:hover { text-decoration: underline; }
        .header-container { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; flex-wrap: wrap; gap: 10px;}
        h1 { flex-grow: 1; text-align: center; color: var(--text-color); margin-bottom: 10px; font-weight: 500; font-size: 1.8em; transition: color var(--transition-speed); }
        .controls { display: flex; gap: 10px; align-items: center;}
        .theme-toggle { background: none; border: 1px solid var(--border-color); border-radius: var(--radius-sm); padding: 8px 12px; cursor: pointer; color: var(--text-color); display: flex; align-items: center; gap: 5px; transition: all var(--transition-speed); }
        .theme-toggle:hover { background-color: rgba(0,0,0,0.05); }
        [data-theme="dark"] .theme-toggle:hover { background-color: rgba(255,255,255,0.05); }

        /* Abas */
        .tab-nav { list-style-type: none; padding: 0; margin: 20px 0 0 0; display: flex; flex-wrap: wrap; gap: 8px; border: none; width: 100%; }
        .tab-nav li { flex: 1; margin: 0; min-width: 130px; display: flex; }
        .tab-nav button { flex-grow: 1; padding: 14px 10px; font-size: 15px; font-weight: 500; cursor: pointer; border: none; background-color: #dce1e5; /* Cor base */ border-radius: var(--radius-md); text-align: center; transition: all var(--transition-speed); box-shadow: var(--shadow-sm); line-height: 1.2; }
        /* Cores aplicadas diretamente com !important */
        """ + "\n".join([f".tab-nav button.tab-color-{i} {{ background-color: {color} !important; }}" for i, color in enumerate(CORES_ABAS)]) + """
        /* Cor do texto com !important para garantir contraste */
        .tab-color-1, .tab-color-2, .tab-color-8 { color: #333 !important; }
        .tab-color-0, .tab-color-3, .tab-color-4, .tab-color-5, .tab-color-6, .tab-color-7, .tab-color-9 { color: #fff !important; }
        .tab-nav button:hover { opacity: 0.9; box-shadow: var(--shadow-hover); transform: translateY(-2px); }
        .tab-nav button.active { font-weight: 700; opacity: 1; filter: brightness(90%) saturate(120%); box-shadow: var(--shadow-sm); transform: translateY(0); position: relative; }
        .tab-nav button.active::after { content: ''; position: absolute; bottom: -8px; left: 50%; transform: translateX(-50%); width: 0; height: 0; border-left: 8px solid transparent; border-right: 8px solid transparent; border-top: 8px solid currentColor; opacity: 0.7; }

        /* Conteúdo das Abas */
        .tab-content { display: none; animation: fadeIn 0.4s; padding: 25px; background-color: var(--card-bg); border-radius: var(--radius-md); margin-top: 25px; transition: background-color var(--transition-speed); }
        .tab-content.active { display: block; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(5px); } to { opacity: 1; transform: translateY(0); } }

        /* Recentes e Favoritos */
        .quick-access-container { margin-bottom: 20px; display: none; padding: 15px; background-color: rgba(0,0,0,0.02); border-radius: var(--radius-md); transition: background-color var(--transition-speed); border: 1px solid var(--border-color); }
        [data-theme="dark"] .quick-access-container { background-color: rgba(255,255,255,0.02); }
        .quick-access-container.visible { display: block; }
        .quick-access-titulo { font-size: 1.1em; font-weight: 600; margin-bottom: 15px; color: var(--text-color); display: flex; align-items: center; gap: 8px; border-bottom: 1px solid var(--border-color); padding-bottom: 10px; }
        .quick-access-lista { display: flex; flex-wrap: wrap; gap: 10px; margin-bottom: 0px; }
        .quick-access-item { background-color: var(--card-bg); border: 1px solid var(--border-color); border-radius: var(--radius-sm); padding: 8px 12px; font-size: 0.9em; cursor: default; transition: all var(--transition-speed); display: inline-flex; align-items: center; gap: 8px; }
        .recente-item { border-left: 3px solid var(--primary-color); }
        .favorito-item { border-left: 3px solid #ffc107; }
        .quick-access-item .nome-link { flex-grow: 1; cursor: pointer; padding-right: 10px; color: var(--text-color); text-decoration: none; }
        .quick-access-item .nome-link:hover { color: var(--primary-color); text-decoration: underline; }
        .quick-access-item .icon { color: #ffc107; min-width: 16px; text-align: center;}
        .quick-access-item .recent-icon { color: var(--primary-color); min-width: 16px; text-align: center;}
        .quick-access-item .copy-icon { font-size: 1.1em; color: var(--text-muted); margin-left: auto; padding-left: 8px; transition: color var(--transition-speed); opacity: 0.7; cursor: pointer; }
        .quick-access-item .copy-icon:hover { color: var(--primary-color); opacity: 1;}
        .quick-access-item .copy-icon.copied { color: var(--success-color); }

        /* Campos de Busca */
        .busca-container { display: flex; flex-wrap: wrap; gap: 15px; margin-bottom: 15px; padding: 20px; background-color: rgba(0,0,0,0.03); border-radius: var(--radius-md); transition: background-color var(--transition-speed); }
        [data-theme="dark"] .busca-container { background-color: rgba(255,255,255,0.03); }
        .busca-container div { flex: 1; min-width: 200px; }
        .busca-container label { display: block; margin-bottom: 8px; font-weight: 600; font-size: 13px; color: var(--text-color); transition: color var(--transition-speed); }
        .busca-container input { width: 100%; padding: 12px; border: 1px solid var(--border-color); border-radius: var(--radius-sm); font-size: 14px; background-color: var(--card-bg); color: var(--text-color); transition: all var(--transition-speed); }
        .busca-container input:focus { outline: none; border-color: var(--primary-color); box-shadow: 0 0 0 2px rgba(42, 111, 151, 0.2); }
        .busca-container-single input { width: 100%; }
        .busca-controles-extra { display: flex; gap: 10px; margin-top: 15px; width: 100%; justify-content: flex-end; flex-basis: 100%;}
        .btn-limpar { font-size: 0.8em; padding: 5px 10px; background-color: #f8f9fa; color: var(--text-muted); border: 1px solid var(--border-color); border-radius: var(--radius-sm); cursor: pointer; transition: all var(--transition-speed); }
        .btn-limpar:hover { background-color: #e9ecef; color: var(--text-color); }
        [data-theme="dark"] .btn-limpar { background-color: #343a40; color: var(--text-muted); border-color: #444; }
        [data-theme="dark"] .btn-limpar:hover { background-color: #495057; color: var(--text-color); }

        /* Itens */
        .item { display: flex; gap: 15px; border: 1px solid var(--card-border); margin-bottom: 20px; padding: 20px; border-radius: var(--radius-md); background-color: var(--card-bg); box-shadow: var(--shadow-sm); transition: all var(--transition-speed); }
        .item:hover { box-shadow: var(--shadow-hover); border-color: var(--primary-color); }
        .item-selecionar { margin-top: 5px; accent-color: var(--primary-color); transform: scale(1.2); cursor: pointer; flex-shrink: 0; }
        .item-conteudo { flex-grow: 1; }
        .item-header { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 15px; padding-bottom: 10px; border-bottom: 1px dashed var(--border-color); flex-wrap: wrap; gap: 10px; transition: border-color var(--transition-speed); }
        .item-nome { font-weight: 600; font-size: 1.1em; color: var(--text-color); flex-basis: 60%; flex-grow: 1; transition: color var(--transition-speed); }
        .item-meta { font-size: 0.85em; color: var(--text-muted); text-align: right; flex-basis: 35%; flex-grow: 1; transition: color var(--transition-speed); }
        .item-meta span { display: block; margin-bottom: 3px; }
        .item-meta span strong { font-weight: 600; color: var(--text-color); transition: color var(--transition-speed); }
        .item pre { white-space: pre-wrap; word-wrap: break-word; margin: 0 0 15px 0; font-family: 'Consolas', 'Monaco', monospace; font-size: 14px; line-height: 1.6; color: var(--text-color); background-color: rgba(0,0,0,0.02); padding: 15px; border-radius: var(--radius-sm); border: 1px solid var(--border-color); transition: all var(--transition-speed); position: relative; }
        [data-theme="dark"] .item pre { background-color: rgba(255,255,255,0.02); }
        .char-counter { position: absolute; bottom: 5px; right: 10px; font-size: 11px; color: var(--text-muted); background-color: var(--card-bg); padding: 2px 5px; border-radius: 3px; opacity: 0.7; }
        .item-actions { margin-top: 10px; display: flex; gap: 8px; }
        .item button { padding: 8px 15px; border: none; border-radius: var(--radius-sm); cursor: pointer; font-size: 14px; transition: all var(--transition-speed); }
        .btn-copiar-item { background-color: var(--primary-color); color: white; }
        .btn-copiar-item:hover { background-color: var(--primary-hover); transform: translateY(-1px); }
        .btn-favorito-item { background-color: transparent; color: var(--primary-color) !important; border: 1px solid var(--primary-color) !important; }
        .btn-favorito-item:hover { transform: translateY(-1px); border-color: var(--primary-hover); color: var(--primary-hover) !important; }
        .btn-favorito-item.ativo { color: #ffc107 !important; border-color: #ffc107 !important; }
        .item button.copiado-feedback { background-color: var(--success-color); color: white; }

        .escondido { display: none; }

        /* Botão Copiar Selecionados */
        #copiarSelecionadosBtn { background-color: var(--success-color); color: white; border: none; border-radius: var(--radius-sm); padding: 8px 12px; cursor: pointer; display: none; /* Começa escondido */ align-items: center; gap: 5px; transition: all var(--transition-speed); }
        #copiarSelecionadosBtn:hover { filter: brightness(90%); box-shadow: var(--shadow-hover); }
        #copiarSelecionadosBtn.visivel { display: inline-flex !important; } /* Garante visibilidade */
        #copiarSelecionadosBtn.copiado-multi { background-color: #17a2b8; }

        /* Animações e Responsividade */
        @keyframes pulse { 0% { transform: scale(1); } 50% { transform: scale(1.05); } 100% { transform: scale(1); } }
        .pulse { animation: pulse 0.3s ease-in-out; }
        @media (max-width: 768px) { .container { margin: 15px; padding: 20px; } .tab-nav li { min-width: 100px; } .tab-nav button { padding: 10px 8px; font-size: 14px; } .item { padding: 15px; flex-direction: column; gap: 10px;} .item-selecionar { margin-top: 0; align-self: flex-start; } .item-conteudo{ width: 100%;} .item-header { flex-direction: column; align-items: flex-start; } .item-meta { text-align: left; margin-top: 5px; } .busca-container { padding: 15px; } .header-container { flex-direction: column; gap: 15px; } h1 { margin-bottom: 15px; } }
        @media (max-width: 480px) { .tab-nav li { min-width: 80px;} .item-actions { flex-direction: column; align-items: flex-start; } }

    </style>
</head>
<body>
    <div class="container">
        <p class="assinatura">
            Criado por: Joao Gabriel Andrade | E-mail: <a href="mailto:joaogabriel.pca@outlook.com">joaogabriel.pca@outlook.com</a>
        </p>
        <div class="header-container">
            <h1>Assistente Rápido</h1>
            <div class="controls">
                <button id="copiarSelecionadosBtn">
                    <span class="btn-icon">📋</span> Copiar Selecionados (<span id="contadorSelecionados">0</span>)
                </button>
                <button class="theme-toggle" id="themeToggle">
                    <span class="theme-icon">☀️</span>
                    <span class="theme-text">Modo Escuro</span>
                </button>
            </div>
        </div>

        <div class="quick-access-container" id="recentesContainer">
            <div class="quick-access-titulo">
                 <span class="recent-icon">🕒</span> Usados Recentemente
            </div>
            <div class="quick-access-lista" id="listaRecentes">
                 <span style="color: var(--text-muted); font-style: italic;">Nenhum item copiado recentemente.</span>
            </div>
        </div>

        <ul class="tab-nav" id="navAbas">
            @@@PLACEHOLDER_NAV@@@
        </ul>
        @@@PLACEHOLDER_CONTENT@@@
""" # Fim do HTML_INICIO

# TEMPLATES COM checkbox SEM onchange (CORRIGIDO)
ITEM_TEMPLATE_MEDICAMENTOS = """
        <div class="item item-medicamentos" data-nome="{nome_busca_lower}" data-categoria="{categoria_lower}" data-doenca="{doenca_lower}">
            <input type="checkbox" class="item-selecionar">
            <div class="item-conteudo">
                <div class="item-header"> <span class="item-nome">{nome_busca}</span> <span class="item-meta"> <span class="item-categoria"><strong>Cat:</strong> {categoria}</span> <span class="item-doenca"><strong>Ind:</strong> {doenca}</span> </span> </div>
                <pre>{conteudo_formatado}</pre>
                <div class="item-actions">
                    <button onclick="copiarTexto(this)" class="btn-copiar-item">Copiar</button>
                    <button class="btn-favorito-item" onclick="toggleFavorito(this)">Favoritar</button>
                </div>
            </div>
        </div>
"""
ITEM_TEMPLATE_GENERICO = """
        <div class="item item-{id_aba}" data-nome="{nome_busca_lower}">
             <input type="checkbox" class="item-selecionar">
             <div class="item-conteudo">
                <div class="item-header"> <span class="item-nome">{nome_busca}</span> <span class="item-meta"></span> </div>
                <pre>{conteudo_formatado}</pre>
                 <div class="item-actions">
                     <button onclick="copiarTexto(this)" class="btn-copiar-item">Copiar</button>
                     <button class="btn-favorito-item" onclick="toggleFavorito(this)">Favoritar</button>
                </div>
            </div>
        </div>
"""

# JAVASCRIPT_BLOCO contém TODAS as funções e correções
JAVASCRIPT_BLOCO = r"""
    <script>
        // Variáveis globais
        const STORAGE_KEY_THEME = 'assistenteMedicoTheme';
        const STORAGE_KEY_FAVORITOS = 'assistenteMedicoFavoritos';
        const STORAGE_KEY_RECENTES = 'assistenteMedicoRecentes';
        const MAX_RECENTES = """ + str(MAX_RECENTES) + """;
        let favoritos = {};
        let recentes = [];
        let copiarBtnMulti = null; // Definido no DOMContentLoaded

        // Inicialização
        document.addEventListener('DOMContentLoaded', function() {
             copiarBtnMulti = document.getElementById('copiarSelecionadosBtn');
             initTheme();
             loadFavoritos();
             loadRecentes();
             addCharCounters();
             setupCheckboxListeners();

             const navAbas = document.getElementById('navAbas');
             const primeiraAbaButton = navAbas ? navAbas.querySelector('button') : null;
             if(primeiraAbaButton) {
                 const match = primeiraAbaButton.getAttribute('onclick').match(/mostrarAba\('([^']+)'\)/);
                 if (match && match[1]) { mostrarAba(match[1]); }
                 else { console.error("Não foi possível encontrar o ID da primeira aba no botão."); }
             } else { console.error("Nenhum botão de aba encontrado para ativar."); }

             if(copiarBtnMulti) {
                 copiarBtnMulti.addEventListener('click', copiarSelecionados);
             } else {
                 console.error("Botão Copiar Selecionados (#copiarSelecionadosBtn) não encontrado!");
             }
             atualizarBotaoCopiarSelecionados(); // Garante estado inicial correto (escondido)
        });

        // Funções de tema
        function initTheme() { const savedTheme = localStorage.getItem(STORAGE_KEY_THEME); const themeToggle = document.getElementById('themeToggle'); if (!themeToggle) return; if (savedTheme) { document.documentElement.setAttribute('data-theme', savedTheme); updateThemeToggle(savedTheme === 'dark'); } else { const prefersDark = window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches; const initialTheme = prefersDark ? 'dark' : 'light'; document.documentElement.setAttribute('data-theme', initialTheme); updateThemeToggle(prefersDark); } themeToggle.addEventListener('click', toggleTheme); }
        function toggleTheme() { const currentTheme = document.documentElement.getAttribute('data-theme') || 'light'; const newTheme = currentTheme === 'dark' ? 'light' : 'dark'; document.documentElement.setAttribute('data-theme', newTheme); localStorage.setItem(STORAGE_KEY_THEME, newTheme); updateThemeToggle(newTheme === 'dark'); }
        function updateThemeToggle(isDark) { const themeToggle = document.getElementById('themeToggle'); if (!themeToggle) return; const themeIcon = themeToggle.querySelector('.theme-icon'); const themeText = themeToggle.querySelector('.theme-text'); if (!themeIcon || !themeText) return; if (isDark) { themeIcon.textContent = '🌙'; themeText.textContent = 'Modo Claro'; } else { themeIcon.textContent = '☀️'; themeText.textContent = 'Modo Escuro'; } }

        // Funções de navegação
        function mostrarAba(idAbaAlvo) {
            const navAbas = document.getElementById('navAbas');
            const conteudosAbas = document.querySelectorAll('.tab-content');
            if (!navAbas || !conteudosAbas) return;
            conteudosAbas.forEach(content => content.classList.remove('active'));
            navAbas.querySelectorAll('button').forEach(button => button.classList.remove('active'));
            const abaAlvo = document.getElementById(idAbaAlvo);
            const botaoAlvo = navAbas.querySelector(`button[onclick="mostrarAba('${idAbaAlvo}')"]`);
            if(abaAlvo) abaAlvo.classList.add('active');
            else console.error(`Conteúdo da aba não encontrado: ${idAbaAlvo}`);
            if(botaoAlvo) botaoAlvo.classList.add('active');
            else console.error(`Botão da aba não encontrado para: ${idAbaAlvo}`);
            atualizarFavoritos(idAbaAlvo);
            // Não desmarca mais checkboxes ao trocar de aba
            // desmarcarTodosCheckboxes(idAbaAlvo); 
            atualizarBotaoCopiarSelecionados(); // Apenas atualiza o estado do botão
        }

        // Funções de cópia
        function copiarTexto(button) {
             const itemDiv = button.closest('.item'); if (!itemDiv) return;
             const preElement = itemDiv.querySelector('pre');
             if (preElement) {
                 const nomeElement = itemDiv.querySelector('.item-nome');
                 const abaId = itemDiv.closest('.tab-content')?.id;
                 const nome = nomeElement ? nomeElement.textContent : 'Desconhecido';
                 const preClone = preElement.cloneNode(true);
                 const counter = preClone.querySelector('.char-counter');
                 if (counter) { preClone.removeChild(counter); }
                 const texto = (preClone.innerText || preClone.textContent).trim();
                 navigator.clipboard.writeText(texto).then(() => {
                     button.classList.add('copiado-feedback');
                     const originalText = button.textContent;
                     button.textContent = 'Copiado!';
                     setTimeout(() => { button.classList.remove('copiado-feedback'); button.textContent = originalText; }, 2000);
                     if(abaId && nome && texto) { registrarUsoRecente([{ nome: nome, texto: texto, abaId: abaId }]); }
                 }).catch(err => { console.error('Erro ao copiar texto: ', err); alert('Não foi possível copiar o texto.'); });
             } else { console.error('Elemento <pre> não encontrado.'); }
         }

        // Copia TODOS selecionados, ignorando filtro
        function copiarSelecionados() {
            console.log("[copiarSelecionados] Iniciando...");
            const abaAtiva = document.querySelector('.tab-content.active');
            const btnMulti = document.getElementById('copiarSelecionadosBtn');
            if (!abaAtiva || !btnMulti) { console.error("Cópia múltipla: Aba ou Botão não encontrados."); return; }
            const todosCheckboxesSelecionados = abaAtiva.querySelectorAll('.item-selecionar:checked');
            console.log(`[copiarSelecionados] Checkboxes selecionados encontrados: ${todosCheckboxesSelecionados.length}`);
            if (todosCheckboxesSelecionados.length === 0) { return; }
            let textoCombinado = []; let itensCopiadosInfo = []; let countCopiados = 0;
            todosCheckboxesSelecionados.forEach((checkbox, index) => {
                const itemDiv = checkbox.closest('.item');
                const preElement = itemDiv?.querySelector('pre');
                const nomeElement = itemDiv?.querySelector('.item-nome');
                const nomeLog = nomeElement ? nomeElement.textContent : `Item Index ${index}`;
                if (preElement && nomeElement) {
                    const preClone = preElement.cloneNode(true); const counter = preClone.querySelector('.char-counter'); if (counter) preClone.removeChild(counter);
                    const textoItem = (preClone.innerText || preClone.textContent).trim();
                    console.log(`[copiarSelecionados] -> Coletando item [${nomeLog}]`);
                    textoCombinado.push(textoItem);
                    if (nomeElement.textContent && textoItem && abaAtiva.id) { itensCopiadosInfo.push({ nome: nomeElement.textContent, texto: textoItem, abaId: abaAtiva.id }); }
                    countCopiados++;
                } else { console.warn(`[copiarSelecionados] Item selecionado [${nomeLog}] pulado.`); }
            });
            console.log(`[copiarSelecionados] ${countCopiados} textos coletados.`);
            if (countCopiados === 0) { alert("Não foi possível extrair texto."); atualizarBotaoCopiarSelecionados(); return; }
            const textoFinal = textoCombinado.join('\\r\\n\\r\\n');
            console.log("[copiarSelecionados] Texto final (início):", textoFinal.substring(0,100)+"...");
            navigator.clipboard.writeText(textoFinal).then(() => {
                console.log("[copiarSelecionados] Sucesso na cópia.");
                // Feedback Visual seguro
                btnMulti.classList.add('copiado-multi');
                const feedbackText = `✓ ${countCopiados} Iten(s) Copiado(s)!`;
                const originalHTML = `<span class="btn-icon">📋</span> Copiar Selecionados (<span id="contadorSelecionados">0</span>)`; // HTML Original Padrão
                btnMulti.textContent = feedbackText; // Muda texto para feedback

                registrarUsoRecente(itensCopiadosInfo);
                desmarcarTodosCheckboxes(abaAtiva.id);

                setTimeout(() => {
                    console.log("[copiarSelecionados] Timeout: Resetando botão.");
                    btnMulti.classList.remove('copiado-multi');
                    // Restaura HTML original e chama atualizar
                    btnMulti.innerHTML = originalHTML; 
                    atualizarBotaoCopiarSelecionados(); // Atualiza contador para 0 e esconde
                 }, 2500);
            }).catch(err => { console.error('[copiarSelecionados] Erro ao copiar: ', err); alert('Erro ao copiar.'); atualizarBotaoCopiarSelecionados(); });
        }

        // Funções de busca
         function filtrarLista(idAba) {
             const abaConteudo = document.getElementById(idAba); if (!abaConteudo) return;
             const itens = abaConteudo.querySelectorAll(`.item`);
             let mostrarItem;
             if (idAba === 'aba-medicamentos') {
                 const termoNome = document.getElementById('busca-medicamentos-nome').value.toLowerCase(); const termoCategoria = document.getElementById('busca-medicamentos-categoria').value.toLowerCase(); const termoDoenca = document.getElementById('busca-medicamentos-doenca').value.toLowerCase();
                 itens.forEach(item => { const nomeItem = item.getAttribute('data-nome') || ''; const categoriaItem = item.getAttribute('data-categoria') || ''; const doencaItem = item.getAttribute('data-doenca') || ''; const matchNome = (termoNome === '') || nomeItem.includes(termoNome); const matchCategoria = (termoCategoria === '') || categoriaItem.includes(termoCategoria); const matchDoenca = (termoDoenca === '') || doencaItem.includes(termoDoenca); mostrarItem = matchNome && matchCategoria && matchDoenca; item.classList.toggle('escondido', !mostrarItem); });
             } else {
                 const termoBusca = document.getElementById(`busca-${idAba}`).value.toLowerCase();
                 itens.forEach(item => { const nomeItem = item.getAttribute('data-nome') || ''; const preElement = item.querySelector('pre'); const textoConteudo = preElement ? preElement.textContent.toLowerCase() : ''; mostrarItem = (termoBusca === '') || nomeItem.includes(termoBusca) || textoConteudo.includes(termoBusca); item.classList.toggle('escondido', !mostrarItem); });
             }
             requestAnimationFrame(atualizarBotaoCopiarSelecionados);
        }

        // Funções de favoritos
        function loadFavoritos() { const savedFavoritos = localStorage.getItem(STORAGE_KEY_FAVORITOS); if (savedFavoritos) { try { favoritos = JSON.parse(savedFavoritos) || {}; } catch (e) { console.error("Erro ao carregar favoritos:", e); favoritos = {}; } } else { favoritos = {}; } document.querySelectorAll('.btn-favorito-item').forEach(button => { const item = button.closest('.item'); const aba = item.closest('.tab-content'); if (!item || !aba) return; const abaId = aba.id; const nome = item.querySelector('.item-nome')?.textContent; if (!nome) return; if (favoritos[abaId] && favoritos[abaId].includes(nome)) { button.classList.add('ativo'); button.textContent = '★ Favorito'; } else { button.classList.remove('ativo'); button.textContent = 'Favoritar'; } }); }
        function saveFavoritos() { localStorage.setItem(STORAGE_KEY_FAVORITOS, JSON.stringify(favoritos)); }
        function toggleFavorito(button) { const item = button.closest('.item'); const aba = item.closest('.tab-content'); if (!item || !aba) return; const abaId = aba.id; const nomeElement = item.querySelector('.item-nome'); if (!nomeElement) return; const nome = nomeElement.textContent; if (!favoritos[abaId]) { favoritos[abaId] = []; } const index = favoritos[abaId].indexOf(nome); if (index === -1) { favoritos[abaId].push(nome); button.classList.add('ativo'); button.textContent = '★ Favorito'; button.classList.add('pulse'); setTimeout(() => button.classList.remove('pulse'), 300); } else { favoritos[abaId].splice(index, 1); button.classList.remove('ativo'); button.textContent = 'Favoritar'; } saveFavoritos(); atualizarFavoritos(abaId); }
        function atualizarFavoritos(abaId) { const fContainer = document.getElementById(`favoritos-${abaId}`); const fList = document.getElementById(`lista-favoritos-${abaId}`); if (!fContainer || !fList) return; fList.innerHTML = ''; if (favoritos[abaId] && favoritos[abaId].length > 0) { fContainer.classList.add('visible'); const fOrdenados = [...favoritos[abaId]].sort((a, b) => a.localeCompare(b)); fOrdenados.forEach(nome => { const fItem = document.createElement('div'); fItem.className = 'quick-access-item favorito-item'; fItem.innerHTML = `<span class="icon favorito-icon">⭐</span><span class="nome-link">${nome}</span><span class="copy-icon" title="Copiar este favorito">📋</span>`; fItem.querySelector('.nome-link').onclick = () => { irParaItem(abaId, nome); }; fItem.querySelector('.copy-icon').onclick = (e) => { e.stopPropagation(); const items = document.querySelectorAll(`#${abaId} .item`); for (const item of items) { const nEl = item.querySelector('.item-nome'); if (nEl && nEl.textContent === nome) { const pre = item.querySelector('pre'); if(pre){ const preClone = pre.cloneNode(true); const counter = preClone.querySelector('.char-counter'); if (counter) preClone.removeChild(counter); const texto = (preClone.innerText || preClone.textContent).trim(); navigator.clipboard.writeText(texto).then(() => { const icon = e.target; icon.innerHTML = '✓'; icon.classList.add('copied'); setTimeout(() => { icon.innerHTML = '📋'; icon.classList.remove('copied');}, 1500); registrarUsoRecente([{nome: nome, texto: texto, abaId: abaId}]); }).catch(err => console.error("Erro copia fav:", err)); } break; } } }; fList.appendChild(fItem); }); } else { fContainer.classList.remove('visible'); } }

        // Funções de contador de caracteres
        function addCharCounters() { const preElements = document.querySelectorAll('pre'); preElements.forEach(pre => { const oldCounter = pre.querySelector('.char-counter'); if (oldCounter) pre.removeChild(oldCounter); const counter = document.createElement('span'); counter.className = 'char-counter'; counter.textContent = `${pre.textContent.trim().length} caracteres`; pre.appendChild(counter); }); }

        // --- Funções para Multi-Seleção e Limpar ---
        function setupCheckboxListeners() { const container = document.querySelector('.container'); if (container) { container.addEventListener('change', function(event) { if (event.target.matches('.item-selecionar')) { console.log("[Checkbox Change] Evento detectado:", event.target.checked); atualizarBotaoCopiarSelecionados(); } }); } else { console.error("Container principal não encontrado."); } }

        // Função ATUALIZADA para mostrar/esconder botão e atualizar contador (mais robusta)
        function atualizarBotaoCopiarSelecionados() {
            console.log("[atualizarBotao] Iniciando...");
            const abaAtiva = document.querySelector('.tab-content.active');
            const btnMulti = document.getElementById('copiarSelecionadosBtn');

            if (!abaAtiva || !btnMulti) {
                 console.warn("[atualizarBotao] Aba ativa ou Botão não encontrados.");
                 if(btnMulti) btnMulti.classList.remove('visivel');
                return;
            }

            // Busca o span DENTRO do botão CADA VEZ
            let spanContador = btnMulti.querySelector('#contadorSelecionados');

            // Conta TODOS os checkboxes checados na aba ativa
            const checkboxesSelecionados = abaAtiva.querySelectorAll('.item-selecionar:checked');
            let contagemItens = checkboxesSelecionados.length;
            console.log(`[atualizarBotao] Contagem na aba '${abaAtiva.id}': ${contagemItens}`);

            // Garante que o botão tenha a estrutura correta com o span ANTES de atualizar
            const textoBaseBotaoHTML = `<span class="btn-icon">📋</span> Copiar Selecionados (<span id="contadorSelecionados">0</span>)`;
            // Verifica se o span existe. Se não, ou se o texto indica estado "Copiado", reseta.
            if (!spanContador || !btnMulti.textContent.includes("Copiar Selecionados")) {
                console.warn("[atualizarBotao] Span não encontrado ou texto incorreto. Resetando innerHTML.");
                btnMulti.innerHTML = textoBaseBotaoHTML;
                // Re-busca o span após resetar
                spanContador = btnMulti.querySelector('#contadorSelecionados');
            }

            // Atualiza o texto do SPAN se ele foi encontrado ou recriado
            if (spanContador) {
                 spanContador.textContent = contagemItens;
                 console.log(`[atualizarBotao] Texto do span atualizado para: ${contagemItens}`);
            } else {
                 console.error("[atualizarBotao] ERRO FATAL: Span #contadorSelecionados não pôde ser encontrado ou recriado!");
            }

            // Mostra ou esconde o BOTÃO
            if (contagemItens > 0) {
                console.log("[atualizarBotao] CONDIÇÃO: contagemItens > 0. Tornando visível.");
                btnMulti.classList.add('visivel');
                btnMulti.classList.remove('copiado-multi');
            } else { // contagemItens é 0
                console.log("[atualizarBotao] CONDIÇÃO: contagemItens === 0. Escondendo.");
                btnMulti.classList.remove('visivel');
                btnMulti.classList.remove('copiado-multi');
                // Garante que o contador (se existir) mostra 0
                 if (spanContador) spanContador.textContent = '0';
            }
            console.log("[atualizarBotao] Classes do botão final:", btnMulti.classList);
            console.log("[atualizarBotao] Função CONCLUÍDA.");
        }


        function desmarcarTodosCheckboxes(abaId) { const aba = document.getElementById(abaId); if(aba) { const checkboxes = aba.querySelectorAll('.item-selecionar:checked'); checkboxes.forEach(cb => { cb.checked = false; }); } }
        function limparFiltros(idAba) { const aba = document.getElementById(idAba); if (!aba) return; const inputs = aba.querySelectorAll('.busca-container input[type="text"]'); inputs.forEach(input => input.value = ''); filtrarLista(idAba); }
        function limparSelecao(idAba) { desmarcarTodosCheckboxes(idAba); atualizarBotaoCopiarSelecionados(); }

        // --- Funções para Itens Recentes ---
        function loadRecentes() { const savedRecentes = localStorage.getItem(STORAGE_KEY_RECENTES); if (savedRecentes) { try { recentes = JSON.parse(savedRecentes) || []; } catch (e) { recentes = []; } } else { recentes = []; } atualizarRecentes(); }
        function saveRecentes() { localStorage.setItem(STORAGE_KEY_RECENTES, JSON.stringify(recentes)); }
        function registrarUsoRecente(itensInfo) { if (!Array.isArray(itensInfo)) itensInfo = [itensInfo]; const atuais = recentes || []; const novosRecentes = [...itensInfo, ...atuais]; const recentesUnicos = []; const vistos = new Set(); for (const item of novosRecentes) { const chave = `${item.abaId}|${item.nome}`; if (item.nome && item.texto && item.abaId && !vistos.has(chave)) { recentesUnicos.push(item); vistos.add(chave); if (recentesUnicos.length >= MAX_RECENTES) { break; } } } recentes = recentesUnicos; saveRecentes(); atualizarRecentes(); }
        function atualizarRecentes() { const rContainer = document.getElementById('recentesContainer'); const rList = document.getElementById('listaRecentes'); if (!rContainer || !rList) return; rList.innerHTML = ''; if (recentes && recentes.length > 0) { rContainer.classList.add('visible'); recentes.forEach(item => { const rItem = document.createElement('div'); rItem.className = 'quick-access-item recente-item'; const nomeEscapado = htmlEscape(item.nome); rItem.innerHTML = `<span class="recent-icon">🕒</span><span class="nome-link">${nomeEscapado}</span><span class="copy-icon" title="Copiar: ${nomeEscapado}">📋</span>`; rItem.querySelector('.nome-link').onclick = () => { irParaItem(item.abaId, item.nome); }; rItem.querySelector('.copy-icon').onclick = (e) => { e.stopPropagation(); navigator.clipboard.writeText(item.texto).then(() => { const icon = e.target; icon.innerHTML = '✓'; icon.classList.add('copied'); setTimeout(() => { icon.innerHTML = '📋'; icon.classList.remove('copied');}, 1500); registrarUsoRecente([item]); }).catch(err => console.error("Erro copia recente:", err)); }; rList.appendChild(rItem); }); } else { rContainer.classList.remove('visible'); rList.innerHTML = '<span style="color: var(--text-muted); font-style: italic;">Nenhum item copiado recentemente.</span>'; } }
        function irParaItem(abaId, nomeItem) { mostrarAba(abaId); setTimeout(() => { const abaC = document.getElementById(abaId); if (!abaC) return; const itensNaAba = abaC.querySelectorAll('.item'); for(const el of itemsNaAba) { const nEl = el.querySelector('.item-nome'); if(nEl && nEl.textContent === nomeItem) { el.scrollIntoView({ behavior: 'smooth', block: 'center' }); el.classList.add('pulse'); setTimeout(() => el.classList.remove('pulse'), 1200); break; } } }, 150); }
        function htmlEscape(str) { return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;'); }

    </script>
"""

HTML_FIM = """
</div> </body>
</html>
"""

# --- Lógica Principal ---
def gerar_html():
    print(f"Tentando ler TODAS as planilhas do arquivo Excel: {ARQUIVO_EXCEL}")
    try:
        sheets_data = pd.read_excel(ARQUIVO_EXCEL, sheet_name=None)
        print(f"Planilhas lidas com sucesso: {list(sheets_data.keys())}")
    except FileNotFoundError: print(f"ERRO: Arquivo Excel não encontrado em '{ARQUIVO_EXCEL}'"); return
    except Exception as e: print(f"ERRO ao ler o arquivo Excel: {e}"); return

    html_abas_nav_list = []; html_abas_conteudo_list = []; primeira_aba = True
    ordem_planilhas = ["Medicamentos", "ExameFisicos", "Procedimentos", "Orientacoes", "Outros"]
    planilhas_encontradas = list(sheets_data.keys())
    for sheet_name in planilhas_encontradas:
        if sheet_name not in ordem_planilhas: ordem_planilhas.append(sheet_name)

    for idx, sheet_name in enumerate(ordem_planilhas):
        if sheet_name not in sheets_data: print(f"AVISO: Planilha '{sheet_name}' não encontrada. Pulando."); continue
        df = sheets_data[sheet_name]; print(f"Processando planilha: '{sheet_name}'..."); id_aba = f"aba-{sanitizar_nome(sheet_name)}"
        is_medicamentos = False; template_item = ITEM_TEMPLATE_GENERICO; colunas_esperadas = COLUNAS_GENERICAS; coluna_conteudo = 'ConteudoTexto'
        if sheet_name.lower() == 'medicamentos': colunas_esperadas = COLUNAS_MEDICAMENTOS; template_item = ITEM_TEMPLATE_MEDICAMENTOS; is_medicamentos = True; coluna_conteudo = 'PrescricaoCompleta'
        colunas_faltantes = [col for col in colunas_esperadas if col not in df.columns]
        if colunas_faltantes: print(f"AVISO: Planilha '{sheet_name}' pulada. Colunas faltantes: {', '.join(colunas_faltantes)}."); continue
        cor_index = idx % len(CORES_ABAS); classe_cor = f"tab-color-{cor_index}"; active_class_nav = 'active' if primeira_aba else ''
        html_abas_nav_list.append(f'<li><button class="{active_class_nav} {classe_cor}" onclick="mostrarAba(\'{id_aba}\')">{sheet_name}</button></li>')
        active_class_content = 'active' if primeira_aba else ''; conteudo_atual_partes = [f'<div id="{id_aba}" class="tab-content {active_class_content}">']
        conteudo_atual_partes.append(f'<div class="quick-access-container favoritos-container" id="favoritos-{id_aba}"><div class="quick-access-titulo favoritos-titulo"><span class="icon favorito-icon">⭐</span> Favoritos</div><div class="quick-access-lista favoritos-lista" id="lista-favoritos-{id_aba}"></div></div>')
        busca_html = ['<div class="busca-container">']
        if is_medicamentos:
            busca_html.append(f'<div><label for="busca-medicamentos-nome">Buscar por Nome:</label><input type="text" id="busca-medicamentos-nome" onkeyup="filtrarLista(\'{id_aba}\')" placeholder="Ex: Dipirona..."></div>')
            busca_html.append(f'<div><label for="busca-medicamentos-categoria">Buscar por Categoria:</label><input type="text" id="busca-medicamentos-categoria" onkeyup="filtrarLista(\'{id_aba}\')" placeholder="Ex: Antibiótico..."></div>')
            busca_html.append(f'<div><label for="busca-medicamentos-doenca">Buscar por Doença:</label><input type="text" id="busca-medicamentos-doenca" onkeyup="filtrarLista(\'{id_aba}\')" placeholder="Ex: Dor, HAS..."></div>')
        else:
             busca_html.append(f'<div style="flex-basis: 100%;"><label for="busca-{id_aba}">Buscar por Nome ou Conteúdo:</label><input type="text" id="busca-{id_aba}" onkeyup="filtrarLista(\'{id_aba}\')" placeholder="Digite para buscar..."></div>')
        busca_html.append('<div class="busca-controles-extra" style="flex-basis: 100%;">')
        busca_html.append(f'<button class="btn-limpar btn-limpar-filtros" onclick="limparFiltros(\'{id_aba}\')">Limpar Filtros</button>')
        busca_html.append(f'<button class="btn-limpar btn-limpar-selecao" onclick="limparSelecao(\'{id_aba}\')">Limpar Seleção</button>')
        busca_html.append('</div>')
        busca_html.append('</div>')
        conteudo_atual_partes.extend(busca_html)
        # CORREÇÃO DE INDENTAÇÃO APLICADA AQUI
        itens_html_lista = [] # Inicializa lista AQUI com indentação correta
        for indice, linha in df.iterrows(): # Loop interno com indentação maior
            if 'NomeBusca' not in linha or pd.isna(linha['NomeBusca']): continue
            nome_busca = str(linha['NomeBusca']).strip()
            if coluna_conteudo not in linha or pd.isna(linha[coluna_conteudo]): continue
            conteudo_texto = str(linha[coluna_conteudo]).strip(); conteudo_formatado_escaped = html.escape(conteudo_texto)
            if is_medicamentos:
                 categoria = str(linha.get('Categoria', '')).strip(); doenca = str(linha.get('Doenca', '')).strip(); categoria_display = categoria if categoria else ''; doenca_display = doenca if doenca else ''
                 item_html = template_item.format(nome_busca=nome_busca, categoria=categoria_display, doenca=doenca_display, nome_busca_lower=nome_busca.lower(), categoria_lower=categoria.lower(), doenca_lower=doenca.lower(), conteudo_formatado=conteudo_formatado_escaped)
            else: item_html = template_item.format(id_aba=id_aba, nome_busca=nome_busca, nome_busca_lower=nome_busca.lower(), conteudo_formatado=conteudo_formatado_escaped)
            itens_html_lista.append(item_html) # Append indentado dentro do loop interno

        # Linhas seguintes alinhadas com itens_html_lista = []
        conteudo_atual_partes.append("\n".join(itens_html_lista))
        conteudo_atual_partes.append('</div>'); # Fecha tab-content
        html_abas_conteudo_list.append("\n".join(conteudo_atual_partes)) # Junta partes da aba
        primeira_aba = False

    # --- Montagem Final e Escrita (Incremental) ---
    print("Iniciando escrita incremental do HTML...")
    try:
        placeholder_nav = '@@@PLACEHOLDER_NAV@@@'; placeholder_conteudo = '@@@PLACEHOLDER_CONTENT@@@'
        idx_nav_placeholder = HTML_INICIO.find(placeholder_nav); idx_conteudo_placeholder = HTML_INICIO.find(placeholder_conteudo)
        if idx_nav_placeholder == -1 or idx_conteudo_placeholder == -1: print("ERRO CRÍTICO: Placeholders não encontrados!"); return
        with open(ARQUIVO_HTML_SAIDA, 'w', encoding='utf-8') as f:
            f.write(HTML_INICIO[:idx_nav_placeholder]) # Parte 1
            f.write("\n".join(html_abas_nav_list)) # Nav
            f.write(HTML_INICIO[idx_nav_placeholder + len(placeholder_nav) : idx_conteudo_placeholder]) # Parte 2
            for conteudo_aba in html_abas_conteudo_list: f.write(conteudo_aba + "\n") # Conteúdos
            f.write(HTML_INICIO[idx_conteudo_placeholder + len(placeholder_conteudo):]) # Parte 3
            f.write(JAVASCRIPT_BLOCO) # Script
            f.write(HTML_FIM) # Fim
        print(f"Arquivo HTML final gerado com sucesso em: {ARQUIVO_HTML_SAIDA}")
    except MemoryError: print(f"ERRO DE MEMÓRIA durante escrita.");
    except Exception as e: print(f"ERRO inesperado durante escrita: {e}");


# --- Execução ---
if __name__ == "__main__":
    print("--- Script iniciado ---")
    gerar_html()
    print("--- Chamada para gerar_html() concluída ---")

# FIM DO CÓDIGO COMPLETO