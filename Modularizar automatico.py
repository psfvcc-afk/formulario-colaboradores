#!/usr/bin/env python3
"""
Script Automático de Modularização
===================================

Este script divide o código do Colaboradores web.py em módulos organizados.

COMO USAR:
----------
1. Colocar este script na mesma pasta que "Colaboradores web.py"
2. Executar: python3 modularizar_automatico.py
3. Testar: streamlit run main.py

AVISOS:
-------
⚠️ Faz BACKUP antes de executar!
⚠️ O script é uma ferramenta auxiliar, pode precisar ajustes manuais
⚠️ Revê sempre o resultado final
"""

import os
import re
import sys
from pathlib import Path

# Cores para output
class Colors:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BLUE = '\033[94m'
    END = '\033[0m'
    BOLD = '\033[1m'

def print_header(text):
    print(f"\n{Colors.BOLD}{Colors.BLUE}{'='*70}{Colors.END}")
    print(f"{Colors.BOLD}{Colors.BLUE}{text.center(70)}{Colors.END}")
    print(f"{Colors.BOLD}{Colors.BLUE}{'='*70}{Colors.END}\n")

def print_success(text):
    print(f"{Colors.GREEN}✓ {text}{Colors.END}")

def print_warning(text):
    print(f"{Colors.YELLOW}⚠ {text}{Colors.END}")

def print_error(text):
    print(f"{Colors.RED}✗ {text}{Colors.END}")

def print_info(text):
    print(f"{Colors.BLUE}ℹ {text}{Colors.END}")


class ModularizadorAutomatico:
    """
    Classe que automatiza a modularização do código
    """
    
    def __init__(self, arquivo_original="Colaboradores web.py"):
        self.arquivo_original = arquivo_original
        self.pasta_destino = "processamento_salarial"
        self.codigo_completo = ""
        
    def verificar_arquivo_original(self):
        """Verifica se o arquivo original existe"""
        if not os.path.exists(self.arquivo_original):
            print_error(f"Arquivo não encontrado: {self.arquivo_original}")
            print_info(f"Certifica-te que o arquivo '{self.arquivo_original}' está na mesma pasta")
            return False
        
        print_success(f"Arquivo encontrado: {self.arquivo_original}")
        return True
    
    def fazer_backup(self):
        """Cria backup do arquivo original"""
        backup_name = f"{self.arquivo_original}.backup"
        
        try:
            import shutil
            shutil.copy2(self.arquivo_original, backup_name)
            print_success(f"Backup criado: {backup_name}")
            return True
        except Exception as e:
            print_error(f"Erro ao criar backup: {e}")
            return False
    
    def ler_codigo_original(self):
        """Lê o código original"""
        try:
            with open(self.arquivo_original, 'r', encoding='utf-8') as f:
                self.codigo_completo = f.read()
            print_success(f"Código lido: {len(self.codigo_completo)} caracteres")
            return True
        except Exception as e:
            print_error(f"Erro ao ler código: {e}")
            return False
    
    def criar_estrutura_pastas(self):
        """Cria a estrutura de pastas"""
        print_header("CRIANDO ESTRUTURA DE PASTAS")
        
        pastas = [
            self.pasta_destino,
            f"{self.pasta_destino}/config",
            f"{self.pasta_destino}/auth",
            f"{self.pasta_destino}/database",
            f"{self.pasta_destino}/calculations",
            f"{self.pasta_destino}/business_logic",
            f"{self.pasta_destino}/ui",
            f"{self.pasta_destino}/utils"
        ]
        
        for pasta in pastas:
            try:
                os.makedirs(pasta, exist_ok=True)
                print_success(f"Pasta criada: {pasta}/")
            except Exception as e:
                print_error(f"Erro ao criar {pasta}: {e}")
                return False
        
        return True
    
    def criar_init_files(self):
        """Cria arquivos __init__.py em todas as pastas"""
        print_header("CRIANDO ARQUIVOS __init__.py")
        
        pastas = [
            "config", "auth", "database", "calculations", 
            "business_logic", "ui", "utils"
        ]
        
        for pasta in pastas:
            caminho = f"{self.pasta_destino}/{pasta}/__init__.py"
            try:
                with open(caminho, 'w', encoding='utf-8') as f:
                    f.write(f'"""\n{pasta} module\n"""\n')
                print_success(f"Criado: {caminho}")
            except Exception as e:
                print_error(f"Erro ao criar {caminho}: {e}")
                return False
        
        return True
    
    def extrair_secao(self, inicio_regex, fim_regex=None):
        """
        Extrai uma seção do código baseado em regex
        """
        match = re.search(inicio_regex, self.codigo_completo, re.MULTILINE)
        if not match:
            return None
        
        inicio = match.start()
        
        if fim_regex:
            match_fim = re.search(fim_regex, self.codigo_completo[inicio:], re.MULTILINE)
            if match_fim:
                fim = inicio + match_fim.start()
            else:
                fim = len(self.codigo_completo)
        else:
            fim = len(self.codigo_completo)
        
        return self.codigo_completo[inicio:fim]
    
    def criar_modulo_template(self, nome_arquivo, conteudo, pasta):
        """
        Cria um arquivo de módulo com template básico
        """
        caminho_completo = f"{self.pasta_destino}/{pasta}/{nome_arquivo}"
        
        # Template básico
        template = f'''"""
{nome_arquivo}
{'=' * len(nome_arquivo)}

Módulo gerado automaticamente pelo script de modularização.
Revê e ajusta conforme necessário.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, date
import numpy as np

# ==================== CÓDIGO EXTRAÍDO ====================

{conteudo}
'''
        
        try:
            with open(caminho_completo, 'w', encoding='utf-8') as f:
                f.write(template)
            return True
        except Exception as e:
            print_error(f"Erro ao criar {caminho_completo}: {e}")
            return False
    
    def modularizar(self):
        """
        Processo principal de modularização
        """
        print_header("PROCESSO DE MODULARIZAÇÃO AUTOMÁTICA")
        
        # Avisos
        print_warning("ATENÇÃO: Este é um processo automático que pode precisar ajustes manuais")
        print_warning("Certifica-te que tens backup do código original!")
        print()
        
        resposta = input("Continuar? (s/n): ").lower()
        if resposta != 's':
            print_info("Operação cancelada pelo utilizador")
            return False
        
        # Passo 1: Verificar arquivo
        if not self.verificar_arquivo_original():
            return False
        
        # Passo 2: Fazer backup
        if not self.fazer_backup():
            return False
        
        # Passo 3: Ler código
        if not self.ler_codigo_original():
            return False
        
        # Passo 4: Criar estrutura
        if not self.criar_estrutura_pastas():
            return False
        
        # Passo 5: Criar __init__.py
        if not self.criar_init_files():
            return False
        
        # Passo 6: Criar módulos básicos
        print_header("CRIANDO MÓDULOS")
        
        self.criar_modulo_guia()
        
        print_success("Estrutura básica criada!")
        print()
        print_header("PRÓXIMOS PASSOS MANUAIS")
        print_info("1. Abre o arquivo original 'Colaboradores web.py'")
        print_info("2. Identifica as secções principais do código:")
        print_info("   - Imports (início)")
        print_info("   - Configurações (EMPRESAS, FERIADOS, etc)")
        print_info("   - Funções de autenticação")
        print_info("   - Funções de cálculo")
        print_info("   - Funções de dados/Dropbox")
        print_info("   - Interface (cada página)")
        print_info("3. Copia cada secção para o módulo correspondente")
        print_info("4. Ajusta os imports em cada módulo")
        print_info("5. Cria o main.py seguindo o exemplo fornecido")
        print_info("6. Testa: streamlit run main.py")
        print()
        
        return True
    
    def criar_modulo_guia(self):
        """
        Cria um arquivo de guia com instruções
        """
        guia = '''# GUIA DE MODULARIZAÇÃO MANUAL
# ==============================

# Este arquivo contém instruções para completar a modularização manualmente.

# ESTRUTURA CRIADA:
# -----------------
# processamento_salarial/
#   ├── config/          ← Configurações
#   ├── auth/            ← Autenticação
#   ├── database/        ← Dropbox + Excel
#   ├── calculations/    ← Cálculos
#   ├── business_logic/  ← Lógica de negócio
#   ├── ui/              ← Interface
#   └── utils/           ← Utilitários

# PRÓXIMOS PASSOS:
# ----------------

# 1. CONFIG/SETTINGS.PY
# ---------------------
# Copiar para cá:
# - EMPRESAS = {...}
# - FERIADOS_NACIONAIS_2025 = [...]
# - SALARIO_MINIMO = ...
# - COLUNAS_SNAPSHOT = [...]
# - Todas as constantes e configurações

# 2. AUTH/AUTHENTICATION.PY
# --------------------------
# Copiar para cá:
# - def check_password():
# - Funções de login/logout

# 3. DATABASE/DROPBOX_MANAGER.PY
# -------------------------------
# Copiar para cá:
# - Funções get_dropbox_client()
# - Funções upload/download Dropbox
# - Gestão de ficheiros

# 4. CALCULATIONS/SALARY_CALCULATOR.PY
# -------------------------------------
# Copiar para cá:
# - Funções calcular_vencimento*
# - Funções calcular_subsidio*
# - Cálculos de salário

# 5. CALCULATIONS/IRS_CALCULATOR.PY
# ----------------------------------
# Copiar para cá:
# - Funções calcular_irs*
# - Tabela IRS
# - Lógica de retenção

# 6. UI/PAGE_CONFIGURACOES.PY
# ----------------------------
# Copiar para cá:
# - Código do elif menu == "⚙️ Configurações"
# - Criar função: def render_configuracoes()

# 7. UI/PAGE_PROCESSAR.PY
# ------------------------
# Copiar para cá:
# - Código do elif menu == "💼 Processar Salários"
# - Criar função: def render_processar()

# 8. UI/PAGE_OUTPUT.PY
# --------------------
# Copiar para cá:
# - Código do elif menu == "📊 Output"
# - Criar função: def render_output()

# 9. MAIN.PY (CRIAR NA RAIZ)
# ---------------------------
# Ver exemplo completo em EXEMPLO_main.py

# DICA: Consulta GUIA_MODULARIZACAO_COMPLETO.md para detalhes!
'''
        
        with open(f"{self.pasta_destino}/GUIA_MODULARIZACAO_MANUAL.txt", 'w', encoding='utf-8') as f:
            f.write(guia)
        
        print_success(f"Criado: {self.pasta_destino}/GUIA_MODULARIZACAO_MANUAL.txt")
    
    def criar_main_exemplo(self):
        """Cria um exemplo de main.py"""
        main_exemplo = '''"""
main.py - Ponto de entrada da aplicação
"""
import streamlit as st

# Imports (ajustar conforme módulos criados)
# from auth.authentication import check_password
# from ui.page_configuracoes import render_configuracoes
# from ui.page_processar import render_processar
# from ui.page_ftes import render_ftes
# from ui.page_output import render_output
# from ui.page_tabela_irs import render_tabela_irs

# Configuração da página
st.set_page_config(
    page_title="Processamento Salarial v3.5.2",
    page_icon="💰",
    layout="wide"
)

# TODO: Descomentar quando módulos estiverem prontos
# if not check_password():
#     st.stop()

st.title("💰 Processamento Salarial v3.5.2")

# Menu
menu = st.sidebar.radio(
    "Menu Principal",
    ["⚙️ Configurações", "💼 Processar Salários", 
     "👥 Visão FTEs/Secção", "📊 Output", "📈 Tabela IRS"]
)

# TODO: Implementar navegação
if menu == "⚙️ Configurações":
    st.write("TODO: Implementar render_configuracoes()")
    # render_configuracoes()
elif menu == "💼 Processar Salários":
    st.write("TODO: Implementar render_processar()")
    # render_processar()
elif menu == "👥 Visão FTEs/Secção":
    st.write("TODO: Implementar render_ftes()")
    # render_ftes()
elif menu == "📊 Output":
    st.write("TODO: Implementar render_output()")
    # render_output()
elif menu == "📈 Tabela IRS":
    st.write("TODO: Implementar render_tabela_irs()")
    # render_tabela_irs()
'''
        
        caminho = f"{self.pasta_destino}/EXEMPLO_main.py"
        with open(caminho, 'w', encoding='utf-8') as f:
            f.write(main_exemplo)
        
        print_success(f"Criado: {caminho}")


def main():
    """Função principal"""
    print_header("MODULARIZADOR AUTOMÁTICO DE CÓDIGO")
    print_info("Script para dividir Colaboradores web.py em módulos")
    print()
    
    # Verificar se estamos no diretório certo
    if not os.path.exists("Colaboradores web.py"):
        print_error("Arquivo 'Colaboradores web.py' não encontrado!")
        print_info("Certifica-te que estás na pasta correta")
        print_info("Ou ajusta o nome do arquivo no script")
        return 1
    
    # Criar instância do modularizador
    modularizador = ModularizadorAutomatico()
    
    # Executar modularização
    if modularizador.modularizar():
        print_header("SUCESSO!")
        print_success("Estrutura básica criada em: processamento_salarial/")
        print()
        print_info("Consulta os seguintes arquivos para continuar:")
        print_info("- processamento_salarial/GUIA_MODULARIZACAO_MANUAL.txt")
        print_info("- processamento_salarial/EXEMPLO_main.py")
        print_info("- GUIA_MODULARIZACAO_COMPLETO.md (documentação)")
        print()
        print_warning("IMPORTANTE: A modularização completa requer passos manuais")
        print_warning("Segue o guia criado para completar o processo")
        return 0
    else:
        print_error("Erro durante a modularização")
        return 1


if __name__ == "__main__":
    sys.exit(main())