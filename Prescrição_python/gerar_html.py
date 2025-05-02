# INÍCIO DO CÓDIGO COMPLETO (gerar_html.py) - VERSÃO FINAL CORRIGIDA
import pandas as pd
import html
import os
import re
import json  # Importar json

# --- Configurações ---
PASTA_ATUAL = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_EXCEL = os.path.join(PASTA_ATUAL, 'prescricoes.xlsx')
ARQUIVO_HTML_SAIDA = os.path.join(PASTA_ATUAL, 'minhas_prescricoes.html')

COLUNAS_MEDICAMENTOS = ['NomeBusca', 'PrescricaoCompleta', 'Categoria', 'Doenca', 'OrdemPrioridade']
COLUNAS_GENERICAS = ['NomeBusca', 'ConteudoTexto']

CORES_ABAS = [
    '#a8dadc', '#f1faee', '#e63946', '#457b9d', '#1d3557',
    '#fca311', '#b7b7a4', '#d4a373', '#a2d2ff', '#ffafcc'
]
MAX_RECENTES = 7  # Número máximo de itens recentes

# --- Função Auxiliar ---
def sanitizar_nome(nome):
    # Função para criar IDs seguros para HTML/JS a partir dos nomes das planilhas
    nome = nome.lower()
    nome = re.sub(r'[áàâãä]', 'a', nome)
    nome = re.sub(r'[éèêë]', 'e', nome)
    nome = re.sub(r'[íìîï]', 'i', nome)
    nome = re.sub(r'[óòôõö]', 'o', nome)
    nome = re.sub(r'[úùûü]', 'u', nome)
    nome = re.sub(r'[ç]', 'c', nome)
    nome = re.sub(r'[^a-z0-9\s-]', '', nome)
    nome = re.sub(r'\s+', '-', nome).strip('-')
    return nome or "aba-generica"

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
            --primary-color: #457b9d; /* Azul mais próximo da imagem */
            --primary-hover: #3a6a8a;
            --success-color: #2a9d8f; /* Verde azulado */
            --light-bg: #f8f9fa; /* Fundo ligeiramente cinza */
            --dark-bg: #212529;
            --card-bg: #ffffff;
            --card-border: #dee2e6;
            --text-color: #212529;
            --text-muted: #6c757d;
            --border-color: #ced4da;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.05);
            --shadow-md: 0 4px 6px rgba(0,0,0,0.07);
            --shadow-hover: 0 6px 10px rgba(0,0,0,0.1);
            --radius-sm: 0.2rem;
            --radius-md: 0.375rem; /* 6px */
            --radius-lg: 0.5rem;  /* 8px */
            --transition-speed: 0.2s;
        }
        [data-theme="dark"] { /* Tema Escuro */
            --primary-color: #5fa8d3; /* Azul mais claro para contraste */
            --primary-hover: #7bbce0;
            --success-color: #52b788;
            --light-bg: #1a1a1a;
            --dark-bg: #121212;
            --card-bg: #2c2c2c;
            --card-border: #444;
            --text-color: #e9ecef;
            --text-muted: #adb5bd;
            --border-color: #555;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.3);
            --shadow-md: 0 4px 6px rgba(0,0,0,0.4);
            --shadow-hover: 0 6px 10px rgba(0,0,0,0.5);
        }
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; margin: 0; padding: 0; background-color: var(--light-bg); color: var(--text-color); transition: background-color var(--transition-speed), color var(--transition-speed); line-height: 1.5; }
        .container { max-width: 1000px; margin: 25px auto; background-color: var(--card-bg); padding: 25px; border-radius: var(--radius-lg); box-shadow: var(--shadow-md); transition: background-color var(--transition-speed), box-shadow var(--transition-speed); border: 1px solid var(--card-border); }
        .assinatura { text-align: center; width: 100%; order: 99; /* Coloca no final */ margin-top: 30px; font-size: 0.8em; color: var(--text-muted); padding-top: 15px; border-top: 1px solid var(--border-color); transition: color var(--transition-speed), border-color var(--transition-speed); }
        .assinatura a { color: var(--primary-color); text-decoration: none; transition: color var(--transition-speed); }
        .assinatura a:hover { text-decoration: underline; }

        /* Header Ajustado */
        .header-container { display: flex; justify-content: space-between; align-items: center; margin-bottom: 25px; flex-wrap: wrap; gap: 15px; padding-bottom: 15px; border-bottom: 1px solid var(--border-color); }
        h1 { flex-grow: 1; text-align: left; /* Alinha título à esquerda */ color: var(--text-color); margin: 0; font-weight: 600; font-size: 1.75em; transition: color var(--transition-speed); }
        .controls { display: flex; gap: 12px; align-items: center;}
        .theme-toggle { background: none; border: 1px solid var(--border-color); border-radius: var(--radius-md); padding: 8px 12px; cursor: pointer; color: var(--text-color); display: flex; align-items: center; gap: 6px; transition: all var(--transition-speed); font-size: 0.9em; }
        .theme-toggle:hover { background-color: rgba(0,0,0,0.05); border-color: var(--primary-color); }
        [data-theme="dark"] .theme-toggle:hover { background-color: rgba(255,255,255,0.1); border-color: var(--primary-color); }
        #copiarSelecionadosBtn { /* Estilo do botão copiar */
            background-color: #6c757d; /* Cinza escuro como nova cor base */
            color: white; border: none; padding: 8px 15px; border-radius: var(--radius-md); cursor: pointer; font-size: 0.9em; transition: all var(--transition-speed); display: none; /* Começa escondido */
            opacity: 0; visibility: hidden;
        }
        #copiarSelecionadosBtn.visivel { /* Classe adicionada por JS */
            display: inline-flex; align-items: center; gap: 6px;
            opacity: 1; visibility: visible;
        }
        #copiarSelecionadosBtn:hover { background-color: #5a6268; /* Cinza mais escuro no hover */ }
        #copiarSelecionadosBtn.copiado-multi { background-color: var(--success-color); } /* Verde sucesso mantido */
        #copiarSelecionadosBtn .btn-icon { margin-right: 4px; }

        /* Abas */
        .tab-nav { list-style-type: none; padding: 0; margin: 25px 0 0 0; display: flex; flex-wrap: wrap; gap: 10px; border: none; width: 100%; }
        .tab-nav li { flex: 1; margin: 0; min-width: 120px; display: flex; }
        .tab-nav button { flex-grow: 1; padding: 12px 10px; font-size: 0.95em; font-weight: 500; cursor: pointer; border: none; background-color: #e9ecef; /* Cor base mais clara */ border-radius: var(--radius-md); text-align: center; transition: all var(--transition-speed); box-shadow: var(--shadow-sm); line-height: 1.3; color: #333; /* Cor de texto padrão */ }
        /* Cores aplicadas diretamente com !important */
        """ + "\n".join([f".tab-nav button.tab-color-{i} {{ background-color: {color} !important; }}" for i, color in enumerate(CORES_ABAS)]) + """
        /* Cor do texto com !important para garantir contraste */
        .tab-color-1, .tab-color-8 { color: #333 !important; } /* Cores claras precisam de texto escuro */
        .tab-color-0, .tab-color-2, .tab-color-3, .tab-color-4, .tab-color-5, .tab-color-6, .tab-color-7, .tab-color-9 { color: #fff !important; } /* Cores escuras precisam de texto claro */
        .tab-nav button:hover { opacity: 0.9; box-shadow: var(--shadow-hover); transform: translateY(-1px); }
        .tab-nav button.active { font-weight: 600; opacity: 1; filter: brightness(100%) saturate(100%); /* Sem filtro extra */ box-shadow: inset 0 2px 4px rgba(0,0,0,0.1); transform: translateY(1px); position: relative; }
        /* Remove o triângulo ::after */
        /* .tab-nav button.active::after { content: none; } */

        /* Conteúdo das Abas */
        .tab-content { display: none; animation: fadeIn 0.3s ease-out; padding: 20px; background-color: transparent; /* Fundo transparente, container pai tem cor */ border: none; /* Sem borda extra */ border-radius: 0; margin-top: 20px; transition: none; }
        .tab-content.active { display: block; }
        @keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }

        /* Recentes e Favoritos (Layout de Chips) */
        .quick-access-container { margin-bottom: 25px; display: none; padding: 15px; background-color: var(--light-bg); border-radius: var(--radius-md); transition: background-color var(--transition-speed); border: 1px solid var(--border-color); }
        [data-theme="dark"] .quick-access-container { background-color: var(--dark-bg); }
        .quick-access-container.visible { display: block; }
        .quick-access-titulo { font-size: 1.05em; font-weight: 600; margin-bottom: 12px; color: var(--text-color); display: flex; align-items: center; gap: 8px; border-bottom: 1px solid var(--border-color); padding-bottom: 8px; }
        .quick-access-lista { display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 0px; }
        .quick-access-item {
            background-color: var(--card-bg); border: 1px solid var(--border-color); border-radius: var(--radius-md); /* Mais arredondado */
            padding: 5px 10px; /* Padding para chip */
            font-size: 0.85em; cursor: default; transition: all var(--transition-speed);
            display: inline-flex; align-items: center; gap: 6px; /* Espaço entre ícones/texto */
            box-shadow: var(--shadow-sm);
        }
        .recente-item { border-left: none; /* Remove borda esquerda */ }
        .favorito-item { border-left: none; /* Remove borda esquerda */ }
        .quick-access-item:hover { background-color: var(--light-bg); border-color: var(--primary-color); }
        [data-theme="dark"] .quick-access-item:hover { background-color: var(--dark-bg); }
        .quick-access-item .nome-link { flex-grow: 0; /* Não cresce */ cursor: pointer; padding-right: 5px; color: var(--text-color); text-decoration: none; white-space: nowrap; }
        .quick-access-item .nome-link:hover { color: var(--primary-color); text-decoration: underline; }
        .quick-access-item .icon { color: #ffc107; min-width: auto; } /* Ícone de favorito */
        .quick-access-item .recent-icon { color: var(--primary-color); min-width: auto; } /* Ícone de recente */
        .quick-access-item .copy-icon {
            font-size: 1em; color: var(--text-muted); margin-left: 5px; /* Espaço antes do ícone */
            padding: 2px; transition: color var(--transition-speed); opacity: 0.7; cursor: pointer;
            border-radius: 50%; /* Círculo sutil */
        }
        .quick-access-item .copy-icon:hover { color: var(--primary-color); opacity: 1; background-color: rgba(0,0,0,0.05); }
        [data-theme="dark"] .quick-access-item .copy-icon:hover { background-color: rgba(255,255,255,0.1); }
        .quick-access-item .copy-icon.copied { color: var(--success-color); }

        /* Container de Busca (Layout Horizontal para Medicamentos) */
        .busca-container {
            display: flex; flex-wrap: wrap; gap: 15px; /* Espaçamento entre campos/linhas */
            padding: 15px; background-color: var(--light-bg); border-radius: var(--radius-md);
            margin-bottom: 25px; border: 1px solid var(--border-color);
        }
        [data-theme="dark"] .busca-container { background-color: var(--dark-bg); }
        .busca-container > div { /* Divs que contêm label+input */
            flex: 1; /* Tenta ocupar espaço igual */
            min-width: 200px; /* Largura mínima antes de quebrar */
            display: flex; flex-direction: column; gap: 5px;
        }
        .busca-container label { font-size: 0.85em; color: var(--text-muted); font-weight: 500; }
        .busca-container input[type="text"] {
            padding: 8px 10px; border: 1px solid var(--border-color); border-radius: var(--radius-md);
            font-size: 0.95em; background-color: var(--card-bg); color: var(--text-color);
            transition: border-color var(--transition-speed), box-shadow var(--transition-speed);
        }
        .busca-container input[type="text"]:focus {
            border-color: var(--primary-color); outline: none;
            box-shadow: 0 0 0 2px rgba(var(--primary-color-rgb, 69, 123, 157), 0.25); /* Adiciona um brilho no foco */
        }
        /* Container para botões Limpar */
        .busca-controles-extra {
            flex-basis: 100%; /* Ocupa linha inteira */
            display: flex; gap: 10px; margin-top: 10px; /* Espaço acima dos botões */
            justify-content: flex-end; /* Alinha botões à direita */
        }
        .btn-limpar {
            background-color: transparent; color: var(--text-muted); border: 1px solid var(--border-color);
            padding: 6px 12px; border-radius: var(--radius-md); cursor: pointer; font-size: 0.85em;
            transition: all var(--transition-speed);
        }
        .btn-limpar:hover { background-color: var(--light-bg); border-color: var(--text-muted); color: var(--text-color); }
        [data-theme="dark"] .btn-limpar:hover { background-color: var(--dark-bg); }

        /* Estilo dos Itens */
        .item {
            background-color: var(--card-bg); border: 1px solid var(--border-color);
            border-radius: var(--radius-md); padding: 18px; /* Mais padding interno */
            margin-bottom: 15px; display: flex; gap: 15px; /* Espaço entre checkbox e conteúdo */
            box-shadow: var(--shadow-sm); transition: box-shadow var(--transition-speed), border-color var(--transition-speed);
        }
        .item:hover { border-color: var(--primary-color); box-shadow: var(--shadow-md); }
        .item.escondido { display: none; }
        .item-selecionar { margin-top: 4px; /* Alinha melhor com o texto */ height: 18px; width: 18px; accent-color: var(--primary-color); flex-shrink: 0; }
        .item-conteudo { flex-grow: 1; display: flex; flex-direction: column; gap: 8px; /* Espaço entre header, pre, actions */ }
        .item-header { display: flex; justify-content: space-between; align-items: flex-start; /* Alinha topo */ gap: 15px; flex-wrap: wrap; }
        .item-nome { font-weight: 600; font-size: 1.1em; color: var(--text-color); margin-right: auto; /* Empurra meta para a direita */ }
        .item-meta { font-size: 0.85em; color: var(--text-muted); text-align: right; white-space: nowrap; }
        .item-meta span { margin-left: 10px; } /* Espaço entre Cat e Ind */
        .item-meta strong { color: var(--text-color); font-weight: 500; }
        .item pre {
            white-space: pre-wrap; word-wrap: break-word; font-family: 'Consolas', 'Monaco', monospace;
            font-size: 0.9em; background-color: var(--light-bg); padding: 10px; border-radius: var(--radius-sm);
            color: var(--text-color); border: 1px solid var(--border-color); position: relative; /* Para o contador */
        }
        [data-theme="dark"] .item pre { background-color: var(--dark-bg); }
        .char-counter {
            position: absolute; bottom: 5px; right: 8px; font-size: 0.75em;
            color: var(--text-muted); background-color: rgba(255, 255, 255, 0.7);
            padding: 1px 4px; border-radius: var(--radius-sm);
        }
        [data-theme="dark"] .char-counter { background-color: rgba(0, 0, 0, 0.5); }
        .item-actions { display: flex; gap: 10px; margin-top: 5px; }
        /* Estilo Base Comum para botões de ação */
        .item-actions button {
            padding: 6px 12px; font-size: 0.85em; border-radius: var(--radius-md); cursor: pointer;
            border: 1px solid var(--border-color); transition: all var(--transition-speed);
        }
        /* Botão Copiar Item (Fundo Primário) */
        .item-actions .btn-copiar-item {
            background-color: var(--primary-color); color: white; border-color: var(--primary-color);
        }
        .item-actions .btn-copiar-item:hover { background-color: var(--primary-hover); border-color: var(--primary-hover); }
        .item-actions .btn-copiar-item.copiado-feedback { background-color: var(--success-color); color: white; border-color: var(--success-color); }

        /* Botão Favoritar Item (Estilo Sutil por Padrão) */
        .item-actions .btn-favorito-item {
            background-color: transparent; color: var(--text-muted); border-color: var(--border-color);
        }
        .item-actions .btn-favorito-item:hover { background-color: var(--light-bg); border-color: var(--text-muted); color: var(--text-color); }
        [data-theme="dark"] .item-actions .btn-favorito-item { background-color: transparent; color: var(--text-muted); border-color: var(--border-color); }
        [data-theme="dark"] .item-actions .btn-favorito-item:hover { background-color: var(--dark-bg); border-color: var(--text-muted); color: var(--text-color); }

        /* Botão Favoritar ATIVO (Estilo Amarelo) */
        .item-actions .btn-favorito-item.ativo {
            background-color: #ffc107; color: #333; border-color: #ffc107; font-weight: 500;
        }
        .item-actions .btn-favorito-item.ativo:hover { filter: brightness(95%); }

        /* Botão Voltar ao Topo (Retangular) */
        #scrollToTopBtn {
            display: none; position: fixed; bottom: 20px; right: 20px; z-index: 99;
            border: none; outline: none; background-color: var(--primary-color); color: white;
            cursor: pointer; padding: 8px 12px; /* Ajuste padding para retângulo */
            border-radius: var(--radius-sm); /* Pequeno arredondamento, não círculo */
            font-size: 16px; /* Pode ajustar tamanho do ícone/texto */
            box-shadow: var(--shadow-md);
            transition: background-color var(--transition-speed), opacity var(--transition-speed), transform var(--transition-speed);
            opacity: 0; transform: translateY(10px);
        }
        #scrollToTopBtn.visible { display: block; opacity: 0.8; transform: translateY(0); }
        #scrollToTopBtn:hover { background-color: var(--primary-hover); opacity: 1; }
        [data-theme="dark"] #scrollToTopBtn { background-color: var(--primary-hover); }
        [data-theme="dark"] #scrollToTopBtn:hover { background-color: var(--primary-color); }

        /* Botão Copiar Flutuante (Posição Ajustada) */
        .botao-copiar-flutuante {
          /* Herda estilos de #copiarSelecionadosBtn */
          position: fixed;
          bottom: 85px; /* Mais para cima */
          right: 50%; /* Centraliza horizontalmente */
          transform: translateX(50%); /* Corrige centralização */
          z-index: 1000;
          box-shadow: var(--shadow-md); /* Adiciona sombra para destaque */
        }

        /* Animações e Responsividade */
        @keyframes pulse { 0% { transform: scale(1); } 50% { transform: scale(1.05); } 100% { transform: scale(1); } }
        .pulse { animation: pulse 0.3s ease-in-out; }

        @media (max-width: 768px) {
            .container { margin: 15px; padding: 15px; }
            h1 { font-size: 1.5em; }
            .header-container { flex-direction: column; align-items: stretch; }
            .controls { justify-content: space-between; }
            .tab-nav li { min-width: 90px; }
            .tab-nav button { padding: 10px 8px; font-size: 0.9em; }
            .item { flex-direction: column; gap: 10px; padding: 12px; }
            .item-selecionar { align-self: flex-end; /* Checkbox no canto superior direito */ margin-top: 0; }
            .item-conteudo{ width: 100%; gap: 10px; }
            .item-header { gap: 5px; } /* Reduz gap no header */
            .item-meta { text-align: left; margin-top: 0; white-space: normal; } /* Meta abaixo em telas menores */
            .item-meta span { margin-left: 0; margin-right: 10px; }
            .busca-container { flex-direction: column; } /* Empilha busca em mobile */
            .busca-controles-extra { justify-content: center; } /* Centraliza botões limpar */
            #scrollToTopBtn { bottom: 15px; right: 15px; padding: 8px 12px; font-size: 16px; }
            .botao-copiar-flutuante {
                bottom: 75px; /* Ajusta posição em telas menores */
                right: 50%;
                transform: translateX(50%);
                padding: 10px 18px; /* Um pouco maior em mobile */
            }
        }
        @media (max-width: 480px) {
            .tab-nav li { min-width: 70px;}
            .item-actions { flex-direction: row; flex-wrap: wrap; } /* Botões podem quebrar linha */
            .quick-access-lista { gap: 6px; }
            .quick-access-item { padding: 4px 8px; font-size: 0.8em; }
            .botao-copiar-flutuante {
                 width: calc(100% - 30px); /* Ocupa mais largura em telas muito pequenas */
                 right: 15px;
                 transform: translateX(0); /* Remove centralização */
                 bottom: 65px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Header Movido para Cima -->
        <div class="header-container">
            <h1>Assistente Rápido</h1>
            <div class="controls">
                <!-- Botão Copiar Selecionados (JS controla visibilidade) -->
                <button id="copiarSelecionadosBtn">
                    <span class="btn-icon">📋</span> Copiar (<span id="contadorSelecionados">0</span>)
                </button>
                <button class="theme-toggle" id="themeToggle">
                    <span class="theme-icon">☀️</span>
                    <span class="theme-text">Modo Escuro</span>
                </button>
            </div>
        </div>

        <!-- Recentes -->
        <div class="quick-access-container" id="recentesContainer">
            <div class="quick-access-titulo">
                 <span class="recent-icon">🕒</span> Usados Recentemente
            </div>
            <div class="quick-access-lista" id="listaRecentes">
                 <span style="color: var(--text-muted); font-style: italic;">Nenhum item copiado recentemente.</span>
            </div>
        </div>

        <!-- Navegação e Conteúdo das Abas -->
        <ul class="tab-nav" id="navAbas">
            @@@PLACEHOLDER_NAV@@@
        </ul>
        @@@PLACEHOLDER_CONTENT@@@

        <!-- Botão Voltar ao Topo -->
        <button id="scrollToTopBtn" title="Voltar ao topo" class="visible">↑</button> <!-- Inicia visível para JS controlar -->

        <!-- Assinatura Movida para o Final -->
        <p class="assinatura">
            Criado por: Joao Gabriel Andrade | E-mail: <a href="mailto:joaogabriel.pca@outlook.com">joaogabriel.pca@outlook.com</a>
        </p>
    </div>
    """ # Fim do HTML_INICIO (o script JS vem depois no gerar_html)

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

JAVASCRIPT_BLOCO = r"""
    <script>
        // Variáveis globais
        const STORAGE_KEY_THEME = 'assistenteMedicoTheme';
        const STORAGE_KEY_FAVORITOS = 'assistenteMedicoFavoritos';
        const STORAGE_KEY_RECENTES = 'assistenteMedicoRecentes';
        const MAX_RECENTES = """ + f"{MAX_RECENTES}" + """;
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
             setupScrollToTop();

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
            const textoFinal = textoCombinado.join(`\r\n\r\n`); // Usar template literal JS
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

        // --- Funções de Favoritos (Refatoradas) ---
        function loadFavoritos() {
            const savedFavoritos = localStorage.getItem(STORAGE_KEY_FAVORITOS);
            if (savedFavoritos) {
                try {
                    favoritos = JSON.parse(savedFavoritos) || {};
                } catch (e) {
                    console.error("Erro ao carregar favoritos:", e);
                    favoritos = {};
                }
            } else {
                favoritos = {};
            }
            // Atualiza estado inicial dos botões e as listas de favoritos visíveis
            document.querySelectorAll('.btn-favorito-item').forEach(button => {
                atualizarEstadoBotaoFavorito(button);
            });
            document.querySelectorAll('.tab-content').forEach(aba => {
                atualizarFavoritos(aba.id); // Atualiza a lista de favoritos para cada aba
            });
        }

        function saveFavoritos() {
            localStorage.setItem(STORAGE_KEY_FAVORITOS, JSON.stringify(favoritos));
        }

        function atualizarEstadoBotaoFavorito(button) {
            const item = button.closest('.item');
            const aba = item?.closest('.tab-content');
            if (!item || !aba) return;
            const abaId = aba.id;
            const nomeElement = item.querySelector('.item-nome');
            if (!nomeElement) return;
            const nome = nomeElement.textContent;

            if (favoritos[abaId] && favoritos[abaId].includes(nome)) {
                button.classList.add('ativo');
                button.textContent = '★ Favorito';
            } else {
                button.classList.remove('ativo');
                button.textContent = 'Favoritar';
            }
        }

        function toggleFavorito(button) {
            const item = button.closest('.item');
            const aba = item?.closest('.tab-content');
            if (!item || !aba) return;
            const abaId = aba.id;
            const nomeElement = item.querySelector('.item-nome');
            if (!nomeElement) return;
            const nome = nomeElement.textContent;

            if (!favoritos[abaId]) {
                favoritos[abaId] = [];
            }
            const index = favoritos[abaId].indexOf(nome);

            if (index === -1) { // Adicionar favorito
                favoritos[abaId].push(nome);
                button.classList.add('pulse');
                setTimeout(() => button.classList.remove('pulse'), 300);
            } else { // Remover favorito
                favoritos[abaId].splice(index, 1);
            }

            saveFavoritos();
            atualizarEstadoBotaoFavorito(button); // Atualiza o botão clicado
            atualizarFavoritos(abaId); // Atualiza a lista de favoritos da aba
        }

        function atualizarFavoritos(abaId) {
            const fContainer = document.getElementById(`favoritos-${abaId}`);
            const fList = document.getElementById(`lista-favoritos-${abaId}`);
            if (!fContainer || !fList) {
                 // Não é um erro fatal se a aba não tiver container de favoritos
                 // console.warn(`Container/Lista de favoritos não encontrados para: ${abaId}`);
                 return;
            }

            fList.innerHTML = ''; // Limpa a lista atual

            if (favoritos[abaId] && favoritos[abaId].length > 0) {
                fContainer.classList.add('visible');
                // Ordena alfabeticamente para exibição
                const fOrdenados = [...favoritos[abaId]].sort((a, b) => a.localeCompare(b));

                fOrdenados.forEach(nome => {
                    // Encontra o item original para obter o texto (necessário para cópia)
                    const itemOriginal = document.querySelector(`#${abaId} .item[data-nome="${nome.toLowerCase()}"]`);
                    const preOriginal = itemOriginal?.querySelector('pre');
                    let textoOriginal = '';
                    if (preOriginal) {
                        const preClone = preOriginal.cloneNode(true);
                        const counter = preClone.querySelector('.char-counter');
                        if (counter) preClone.removeChild(counter);
                        textoOriginal = (preClone.innerText || preClone.textContent).trim();
                    } else {
                        console.warn(`Não foi possível encontrar o texto original para o favorito: ${nome} na aba ${abaId}`);
                    }

                    const fItem = document.createElement('div');
                    fItem.className = 'quick-access-item favorito-item';

                    // Escapa dados para HTML e atributos data-*
                    const nomeHtml = htmlEscape(nome);
                    const abaIdHtml = htmlEscape(abaId);
                    const textoData = textoOriginal; // Texto original para data attribute

                    // Cria HTML com data-* attributes
                    fItem.innerHTML = `<span class="icon favorito-icon">⭐</span><span class="nome-link" data-abaid="${abaIdHtml}" data-nome="${nomeHtml}">${nomeHtml}</span><span class="copy-icon" data-texto="${htmlEscape(textoData)}" data-nome="${nomeHtml}" data-abaid="${abaIdHtml}" title="Copiar: ${nomeHtml}">📋</span>`;

                    // Adiciona Listeners usando querySelector no fItem recém-criado
                    const nomeLink = fItem.querySelector('.nome-link');
                    const copyIcon = fItem.querySelector('.copy-icon');

                    if (nomeLink) {
                        nomeLink.addEventListener('click', (event) => {
                            const targetAbaId = event.target.dataset.abaid;
                            const targetNome = event.target.dataset.nome;
                            if(targetAbaId && targetNome) {
                                irParaItem(targetAbaId, targetNome);
                            } else {
                                console.error("Dados faltando no data attribute do nome-link (favoritos).");
                            }
                        });
                    }

                    if (copyIcon) {
                        copyIcon.addEventListener('click', (event) => {
                            event.stopPropagation();
                            const icon = event.target;
                            const targetTexto = icon.dataset.texto;
                            const targetNome = icon.dataset.nome;
                            const targetAbaId = icon.dataset.abaid;

                            if (targetTexto === undefined || targetNome === undefined || targetAbaId === undefined) {
                                console.error("Dados faltando no data attribute do copy-icon (favoritos).");
                                alert("Erro ao obter dados para cópia do favorito.");
                                return;
                            }

                            if (!targetTexto) {
                                console.warn(`Texto vazio para favorito: ${targetNome}. Cópia abortada.`);
                                alert(`Não foi possível encontrar o conteúdo para copiar "${targetNome}".`);
                                return;
                            }


                            navigator.clipboard.writeText(targetTexto).then(() => {
                                icon.innerHTML = '✓';
                                icon.classList.add('copied');
                                setTimeout(() => {
                                    if (icon) {
                                        icon.innerHTML = '📋';
                                        icon.classList.remove('copied');
                                    }
                                }, 1500);
                                const itemInfo = { nome: targetNome, texto: targetTexto, abaId: targetAbaId };
                                registrarUsoRecente([itemInfo]); // Registra como recente
                            }).catch(err => {
                                console.error("Erro ao copiar favorito:", err, targetNome);
                                alert("Erro ao copiar o favorito.");
                            });
                        });
                    }
                    fList.appendChild(fItem); // Adiciona o item à lista
                });
            } else {
                fContainer.classList.remove('visible'); // Esconde se não houver favoritos
            }
        }

        // Funções de contador de caracteres
        function addCharCounters() { const preElements = document.querySelectorAll('pre'); preElements.forEach(pre => { const oldCounter = pre.querySelector('.char-counter'); if (oldCounter) pre.removeChild(oldCounter); const counter = document.createElement('span'); counter.className = 'char-counter'; counter.textContent = `${pre.textContent.trim().length} caracteres`; pre.appendChild(counter); }); }

        // --- Funções para Multi-Seleção e Limpar ---
        function setupCheckboxListeners() { const container = document.querySelector('.container'); if (container) { container.addEventListener('change', function(event) { if (event.target.matches('.item-selecionar')) { console.log("[Checkbox Change] Evento detectado:", event.target.checked); atualizarBotaoCopiarSelecionados(); } }); } else { console.error("Container principal não encontrado."); } }

        // Função ATUALIZADA para mostrar/esconder botão, atualizar contador E controlar flutuação
        function atualizarBotaoCopiarSelecionados() {
            console.log("[atualizarBotao] Iniciando...");
            const abaAtiva = document.querySelector('.tab-content.active');
            const btnMulti = document.getElementById('copiarSelecionadosBtn');

            if (!abaAtiva || !btnMulti) {
                 console.warn("[atualizarBotao] Aba ativa ou Botão não encontrados.");
                 if(btnMulti) {
                     btnMulti.classList.remove('visivel', 'botao-copiar-flutuante'); // Esconde e remove flutuação se botão existe mas aba não
                 }
                return;
            }

            // Busca o span DENTRO do botão CADA VEZ
            let spanContador = btnMulti.querySelector('#contadorSelecionados');

            // Conta TODOS os checkboxes checados na aba ativa (usando a classe correta '.item-selecionar')
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

            // Mostra ou esconde o BOTÃO E CONTROLA FLUTUAÇÃO
            // Alterado para > 0 para mostrar com 1 ou mais selecionados
            if (contagemItens > 0) {
                console.log("[atualizarBotao] CONDIÇÃO: contagemItens > 0. Tornando visível E flutuante.");
                // Adiciona ambas as classes para mostrar e flutuar
                btnMulti.classList.add('visivel', 'botao-copiar-flutuante');
                btnMulti.classList.remove('copiado-multi'); // Remove feedback de cópia se houver
            } else { // contagemItens é 0
                console.log("[atualizarBotao] CONDIÇÃO: contagemItens === 0. Escondendo E removendo flutuação.");
                // Remove ambas as classes para esconder e parar de flutuar
                btnMulti.classList.remove('visivel', 'botao-copiar-flutuante');
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
        function atualizarRecentes() {
            const rContainer = document.getElementById('recentesContainer');
            const rList = document.getElementById('listaRecentes');
            if (!rContainer || !rList) {
                console.error("Elementos de recentes não encontrados.");
                return;
            }
            rList.innerHTML = ''; // Limpa a lista atual

            if (recentes && recentes.length > 0) {
                rContainer.classList.add('visible');
                recentes.forEach(item => {
                    // Validação básica do item
                    if (!item || typeof item.nome !== 'string' || typeof item.texto !== 'string' || typeof item.abaId !== 'string') {
                        console.warn("Item recente inválido encontrado:", item);
                        return; // Pula item inválido
                    }

                    const rItem = document.createElement('div');
                    rItem.className = 'quick-access-item recente-item';

                    // Escapa os dados para uso seguro em atributos HTML e texto visível
                    const nomeHtml = htmlEscape(item.nome);
                    const abaIdHtml = htmlEscape(item.abaId);
                    const textoData = item.texto; // Usar o texto original para o data attribute

                    // Cria a estrutura HTML básica com data-* attributes
                    rItem.innerHTML = `<span class="recent-icon">🕒</span><span class="nome-link" data-abaid="${abaIdHtml}" data-nome="${nomeHtml}">${nomeHtml}</span><span class="copy-icon" data-texto="${htmlEscape(textoData)}" data-nome="${nomeHtml}" data-abaid="${abaIdHtml}" title='Copiar: ${nomeHtml}'>📋</span>`;

                    // Encontra os elementos dentro do item recém-criado
                    const nomeLink = rItem.querySelector('.nome-link');
                    const copyIcon = rItem.querySelector('.copy-icon');

                    // Adiciona event listener para o link do nome
                    if (nomeLink) {
                        nomeLink.addEventListener('click', (event) => {
                            const targetAbaId = event.target.dataset.abaid;
                            const targetNome = event.target.dataset.nome;
                            if (targetAbaId && targetNome) {
                                irParaItem(targetAbaId, targetNome);
                            } else {
                                console.error("Dados faltando no data attribute do nome-link (recentes).");
                            }
                        });
                    } else {
                        console.warn("Elemento .nome-link não encontrado para item recente:", nomeHtml);
                    }

                    // Adiciona event listener para o ícone de cópia
                    if (copyIcon) {
                        copyIcon.addEventListener('click', (event) => {
                            event.stopPropagation();
                            const icon = event.target;
                            const targetTexto = icon.dataset.texto;
                            const targetNome = icon.dataset.nome;
                            const targetAbaId = icon.dataset.abaid;

                            if (targetTexto === undefined || targetNome === undefined || targetAbaId === undefined) {
                                console.error("Dados faltando no data attribute do copy-icon (recentes).");
                                alert("Erro ao obter dados para cópia.");
                                return;
                            }

                            navigator.clipboard.writeText(targetTexto).then(() => {
                                icon.innerHTML = '✓';
                                icon.classList.add('copied');
                                setTimeout(() => {
                                    if (icon) {
                                        icon.innerHTML = '📋';
                                        icon.classList.remove('copied');
                                    }
                                }, 1500);
                                const itemInfo = { nome: targetNome, texto: targetTexto, abaId: targetAbaId };
                                registrarUsoRecente([itemInfo]);
                            }).catch(err => {
                                console.error("Erro ao copiar item recente:", err, targetNome);
                                alert("Erro ao copiar o item.");
                            });
                        });
                    } else {
                         console.warn("Elemento .copy-icon não encontrado para item recente:", nomeHtml);
                    }
                    rList.appendChild(rItem);
                });
            } else {
                rContainer.classList.remove('visible');
                rList.innerHTML = '<span style="color: var(--text-muted); font-style: italic;">Nenhum item copiado recentemente.</span>';
            }
        }

        function irParaItem(abaId, nomeItem) { mostrarAba(abaId); setTimeout(() => { const abaC = document.getElementById(abaId); if (!abaC) return; const itensNaAba = abaC.querySelectorAll('.item'); for(const el of itensNaAba) { const nEl = el.querySelector('.item-nome'); if(nEl && nEl.textContent === nomeItem) { el.scrollIntoView({ behavior: 'smooth', block: 'center' }); el.classList.add('pulse'); setTimeout(() => el.classList.remove('pulse'), 1200); break; } } }, 150); }
        function htmlEscape(str) { return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;'); }

        // --- Funções Botão Voltar ao Topo ---
        function setupScrollToTop() {
            const scrollToTopBtn = document.getElementById("scrollToTopBtn");
            if (!scrollToTopBtn) {
                console.error("Botão Voltar ao Topo (#scrollToTopBtn) não encontrado!");
                return;
            }
            window.onscroll = function() { scrollFunction() };
            function scrollFunction() {
                if (document.body.scrollTop > 100 || document.documentElement.scrollTop > 100) {
                    scrollToTopBtn.style.display = "block";
                    setTimeout(() => { scrollToTopBtn.style.opacity = 0.8; }, 10);
                } else {
                    scrollToTopBtn.style.opacity = 0;
                    setTimeout(() => { scrollToTopBtn.style.display = "none"; }, 300);
                }
            }
            scrollToTopBtn.addEventListener("click", function() {
                window.scrollTo({top: 0, behavior: 'smooth'});
            });
        }

    </script>
"""

HTML_FIM = """
</div> </body>
</html>
"""

# --- Lógica Principal ---
def gerar_html():
    try:
        xls = pd.ExcelFile(ARQUIVO_EXCEL)
    except FileNotFoundError:
        print(f"ERRO: Arquivo Excel não encontrado em '{ARQUIVO_EXCEL}'.")
        return
    except Exception as e:
        print(f"ERRO ao abrir o arquivo Excel: {e}")
        return

    ordem_planilhas = xls.sheet_names # Ordem padrão do arquivo
    html_abas_nav_list = []
    html_abas_conteudo_list = []
    primeira_aba = True

    # Loop principal sobre as planilhas
    for sheet_name in ordem_planilhas:
        print(f"Processando planilha: '{sheet_name}'...")
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            # Remover linhas onde 'NomeBusca' é NaN ou vazio
            df = df.dropna(subset=['NomeBusca'])
            df = df[df['NomeBusca'].astype(str).str.strip() != '']
        except Exception as e:
            print(f"ERRO ao ler a planilha '{sheet_name}': {e}")
            continue # Pula para a próxima planilha

        id_aba = sanitizar_nome(sheet_name)

        # Determina se é planilha de medicamentos (verifica colunas)
        is_medicamentos = all(col in df.columns for col in COLUNAS_MEDICAMENTOS)

        if is_medicamentos:
            colunas_esperadas = COLUNAS_MEDICAMENTOS
            coluna_conteudo = 'PrescricaoCompleta'
        else:
            colunas_esperadas = COLUNAS_GENERICAS
            coluna_conteudo = 'ConteudoTexto'

        # Verifica colunas essenciais
        colunas_faltantes = [col for col in colunas_esperadas if col not in df.columns]
        if colunas_faltantes:
            print(f"AVISO: Planilha '{sheet_name}' pulada. Colunas essenciais faltantes: {', '.join(colunas_faltantes)}.")
            continue # Pula para a próxima planilha

        # --- Ordenação ---
        if is_medicamentos:
            print(f"Aplicando ordenação personalizada para '{sheet_name}'...")
            # Converter OrdemPrioridade para numérico, tratando erros e preenchendo NaN
            df['OrdemPrioridadeNumerica'] = pd.to_numeric(df['OrdemPrioridade'], errors='coerce').fillna(float('inf'))
            # Ordena: 1º Doenca (A-Z), 2º OrdemPrioridade (Menor primeiro, sem prioridade por último), 3º NomeBusca (A-Z)
            df = df.sort_values(by=['Doenca', 'OrdemPrioridadeNumerica', 'NomeBusca'], ascending=[True, True, True])
            # df = df.drop(columns=['OrdemPrioridadeNumerica']) # Opcional
            print(f"Ordenação para '{sheet_name}' concluída.")
        elif not df.empty and 'NomeBusca' in df.columns: # Ordenação padrão para outras abas
             df = df.sort_values(by=['NomeBusca'], ascending=True)

        # --- Geração de HTML para a Aba ---
        cor_index = ordem_planilhas.index(sheet_name) % len(CORES_ABAS)
        classe_cor = f"tab-color-{cor_index}"
        active_class_nav = 'active' if primeira_aba else ''
        html_abas_nav_list.append(f'<li><button class="{active_class_nav} {classe_cor}" onclick="mostrarAba(\'{id_aba}\')">{sheet_name}</button></li>')

        active_class_content = 'active' if primeira_aba else ''
        conteudo_atual_partes = [f'<div id="{id_aba}" class="tab-content {active_class_content}">']

        # Adiciona container de favoritos
        conteudo_atual_partes.append(f'<div class="quick-access-container favoritos-container" id="favoritos-{id_aba}"><div class="quick-access-titulo favoritos-titulo"><span class="icon favorito-icon">⭐</span> Favoritos</div><div class="quick-access-lista favoritos-lista" id="lista-favoritos-{id_aba}"></div></div>')

        # Adiciona container de busca
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
        busca_html.append('</div>') # Fecha busca-controles-extra
        busca_html.append('</div>') # Fecha busca-container
        conteudo_atual_partes.extend(busca_html)

        # --- Geração de HTML para os Itens ---
        itens_html_lista = []
        for indice, linha in df.iterrows():
            # Pula linha se coluna de conteúdo estiver vazia (NomeBusca já foi verificado)
            if pd.isna(linha[coluna_conteudo]): continue
            nome_busca = str(linha['NomeBusca']).strip()
            conteudo_texto = str(linha[coluna_conteudo]).strip()
            conteudo_formatado_escaped = html.escape(conteudo_texto)

            template_item = ITEM_TEMPLATE_MEDICAMENTOS if is_medicamentos else ITEM_TEMPLATE_GENERICO

            if is_medicamentos:
                 categoria = str(linha.get('Categoria', '')).strip()
                 doenca = str(linha.get('Doenca', '')).strip()
                 categoria_display = categoria if categoria else ''
                 doenca_display = doenca if doenca else ''
                 item_html = template_item.format(
                     nome_busca=html.escape(nome_busca),
                     categoria=html.escape(categoria_display),
                     doenca=html.escape(doenca_display),
                     nome_busca_lower=html.escape(nome_busca.lower()),
                     categoria_lower=html.escape(categoria.lower()),
                     doenca_lower=html.escape(doenca.lower()),
                     conteudo_formatado=conteudo_formatado_escaped
                 )
            else:
                 item_html = template_item.format(
                     id_aba=id_aba,
                     nome_busca=html.escape(nome_busca),
                     nome_busca_lower=html.escape(nome_busca.lower()),
                     conteudo_formatado=conteudo_formatado_escaped
                 )
            itens_html_lista.append(item_html)

        conteudo_atual_partes.append("\n".join(itens_html_lista))
        conteudo_atual_partes.append('</div>') # Fecha tab-content
        html_abas_conteudo_list.append("\n".join(conteudo_atual_partes))
        primeira_aba = False
    # Fim do loop principal sobre as planilhas

    # Combina todas as partes do HTML
    html_navegacao = "\n".join(html_abas_nav_list)
    html_conteudo = "\n".join(html_abas_conteudo_list)
    html_final = HTML_INICIO.replace('@@@PLACEHOLDER_NAV@@@', html_navegacao)
    html_final = html_final.replace('@@@PLACEHOLDER_CONTENT@@@', html_conteudo)
    html_final += JAVASCRIPT_BLOCO # Adiciona o bloco JavaScript
    html_final += HTML_FIM # Adiciona o final do HTML

    # Escreve o arquivo HTML
    try:
        with open(ARQUIVO_HTML_SAIDA, 'w', encoding='utf-8') as f:
            f.write(html_final)
        print(f"Arquivo HTML \'{ARQUIVO_HTML_SAIDA}\' gerado/atualizado com sucesso.")
    except Exception as e:
        print(f"ERRO ao escrever o arquivo HTML: {e}")


# --- Execução ---
if __name__ == "__main__":
    print("--- Script iniciado ---")
    gerar_html()
    print("--- Chamada para gerar_html() concluída ---")

# FIM DO CÓDIGO COMPLETO