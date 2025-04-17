#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Extrator de Dados do SISREG III para Google Sheets
Este script extrai informações de arquivos PDF do SISREG III e as envia para uma planilha do Google Sheets.
Versão otimizada para lidar com múltiplos formatos de PDF do SISREG III.
"""

import os
import re
import json
import subprocess
import pytesseract
from pdf2image import convert_from_path
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

def extrair_texto_pdf(pdf_path):
    """
    Extrai texto de um arquivo PDF usando múltiplos métodos.
    Tenta primeiro pdftotext, depois PyPDF2 e, por último, OCR com Tesseract.
    
    Args:
        pdf_path: Caminho para o arquivo PDF
        
    Returns:
        String contendo o texto extraído do PDF
    """
    # Tentativa 1: Usar pdftotext (mais rápido e preciso para PDFs digitais)
    try:
        resultado = subprocess.run(['pdftotext', '-layout', pdf_path, '-'], 
                                  capture_output=True, text=True)
        if resultado.returncode == 0 and resultado.stdout.strip():
            return resultado.stdout
    except Exception as e:
        print(f"Erro ao usar pdftotext: {e}")
    
    # Tentativa 2: Usar PyPDF2 (biblioteca Python pura, não requer dependências externas)
    try:
        import PyPDF2
        texto_total = ""
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page_num in range(len(reader.pages)):
                texto_total += reader.pages[page_num].extract_text() + "\n"
        if texto_total.strip():
            return texto_total
    except Exception as e:
        print(f"Erro ao usar PyPDF2: {e}")
    
    # Tentativa 3: Fallback para o método OCR se os outros métodos falharem
    try:
        imagens = convert_from_path(pdf_path)
        texto_total = ""
        for imagem in imagens:
            texto_total += pytesseract.image_to_string(imagem, lang='por')
        return texto_total
    except Exception as e:
        print(f"Erro ao usar OCR: {e}")
        return ""


def extrair_dados(texto, nome_arquivo=""):
    """
    Extrai dados específicos do texto do PDF usando expressões regulares.
    Otimizado para lidar com múltiplos formatos de PDF do SISREG III.
    
    Args:
        texto: Texto extraído do PDF
        nome_arquivo: Nome do arquivo para debug (opcional)
        
    Returns:
        Dicionário com os dados extraídos (nunca retorna None)
    """
    import os
    import re
    
    # Inicializar o dicionário de dados com valores padrão
    dados = {
        "codigo_solicitacao": "NÃO ENCONTRADO",
        "cns": "NÃO ENCONTRADO",
        "unidade_solicitante": "NÃO ENCONTRADO",
        "unidade_executante": "NÃO ENCONTRADO",
        "data_exame": "NÃO ENCONTRADO",
        "procedimento": "NÃO ENCONTRADO"
    }
    
    try:
        # Salvar o texto para debug
        if nome_arquivo:
            debug_file = f"texto_extraido_{os.path.basename(nome_arquivo)}.txt"
            with open(debug_file, "w", encoding="utf-8") as f:
                f.write(texto)
        
        # Extrair CNS - procurar especificamente por padrões de CNS
        cns_patterns = [
            r'CNS\s*:?\s*(\d{15})',  # CNS com 15 dígitos após "CNS:"
            r'CNS\s*:?\s*(\d+)',     # Qualquer número após "CNS:"
            r'DADOS\s+DO\s+PACIENTE[\s\S]*?CNS\s*:?\s*(\d+)',  # CNS na seção de dados do paciente
            r'Apelido\s*:?\s*(\d{15})'
        ]
        
        cns_encontrado = False
        for pattern in cns_patterns:
            match = re.search(pattern, texto, re.IGNORECASE)
            if match:
                dados["cns"] = match.group(1).strip()
                cns_encontrado = True
                break
        
        if not cns_encontrado:
            # Procurar por números de 15 dígitos (formato típico do CNS)
            match = re.search(r'\b(\d{15})\b', texto)
            if match:
                dados["cns"] = match.group(1).strip()
        
        # NOVA CORREÇÃO: Adicionar padrões específicos para o código junto com a data
        codigo_patterns = [
            # NOVOS PADRÕES: Código seguido imediatamente por data
            # Busca por "Vaga Solicitada:Vaga Consumida:" seguido por código de 9 dígitos começando com 5 e data
            r'Vaga\s*Solicitada\s*:\s*Vaga\s*Consumida\s*:\s*(5\d{8})(\d{2}/\d{2}/\d{4})',
            
            # Busca por "Consumida:" seguido por código de 9 dígitos começando com 5 e data
            r'Consumida\s*:\s*(5\d{8})(\d{2}/\d{2}/\d{4})',
            
            # Busca por qualquer ":" seguido por código de 9 dígitos começando com 5 e data
            r':\s*(5\d{8})(\d{2}/\d{2}/\d{4})',
            
            # Busca genérica por código de 9 dígitos começando com 5 seguido por data
            r'\b(5\d{8})(\d{2}/\d{2}/\d{4})\b',
            
            # PADRÕES ANTERIORES: Código que começa com 5
            # Busca por "Código da Solicitação" e depois procura por um número de 9 dígitos que começa com 5
            r'C[óo]digo\s*d[ae]\s*Solicita[çc][ãa]o\s*:?[\s\S]{0,100}?(5\d{8})\b',
            
            # Busca por "Código da Solicitação:" seguido de número que começa com 5
            r'C[óo]digo\s*d[ae]\s*Solicita[çc][ãa]o\s*:?\s*(5\d+)',
            
            # Busca por número de 9 dígitos que começa com 5 após "Vaga Solicitada" e "Vaga Consumida"
            r'Vaga\s+Solicitada\s*:?\s*Vaga\s+Consumida\s*:?[\s\S]{0,100}?(5\d{8})\b',
            
            # Busca por número de 9 dígitos que começa com 5 após "1ª Vez"
            r'1[ªa]\s+Vez[\s\S]{0,50}?(5\d{8})\b',
            
            # Busca por número de 9 dígitos que começa com 5 isolado no início de uma linha
            r'^\s*(5\d{8})\s*$',
            
            # Busca por número de 9 dígitos que começa com 5 em qualquer lugar
            r'\b(5\d{8})\b',
            
            # PADRÕES DE FALLBACK: Menos específicos
            r'C[óo]digo\s*d[ae]\s*Solicita[çc][ãa]o\s*:?[\s\S]{0,100}?(\d{9})\b',
            r'C[óo]digo\s*d[ae]\s*Solicita[çc][ãa]o\s*:?\s*(\d+)',
            r'Vaga\s+Solicitada\s*:?\s*Vaga\s+Consumida\s*:?[\s\S]{0,100}?(\d{9})\b',
            r'1[ªa]\s+Vez[\s\S]{0,50}?(\d{9})\b',
            r'^\s*(\d{9})\s*$',
            r'\b(\d{9})\b'
        ]
        
        # Tentar cada padrão para o código de solicitação
        codigo_encontrado = False
        for pattern in codigo_patterns:
            match = re.search(pattern, texto, re.IGNORECASE | re.MULTILINE)
            if match:
                codigo = match.group(1).strip()
                
                # Validação para garantir que o código:
                # 1. Tenha exatamente 9 dígitos
                # 2. Comece com o dígito 5 (se possível)
                # 3. Não seja igual ao CNS
                if (len(codigo) == 9 and 
                    codigo != dados.get("cns", "") and 
                    (codigo.startswith('5') or (not codigo_encontrado and not '5' in pattern))):
                    
                    # Se encontramos um código que começa com 5, usamos ele
                    # Se não, só aceitamos códigos que não começam com 5 se não tivermos encontrado nenhum outro
                    if codigo.startswith('5') or not codigo_encontrado:
                        dados["codigo_solicitacao"] = codigo
                        codigo_encontrado = True
                        print(f"Código de solicitação encontrado com o padrão: {pattern}")
                        print(f"Código encontrado: {codigo}")
                        
                        # Se encontramos um código que começa com 5, paramos a busca
                        if codigo.startswith('5'):
                            break
        
        # Extrair os demais campos
        outros_padroes = {
            "unidade_executante": r'UNIDADE\s*EXECUTANTE[\s\S]*?Nome\s*:\s*([^\r\n:]+)',
            "unidade_solicitante": r'UNIDADE\s*SOLICITANTE[\s\S]*?Nome\s*:?\s*([^\r\n:]+)',
            "data_exame": r'Data\s*e\s*Hor[áa]rio\s*de\s*Atendimento\s*:?\s*([^\r\n]+)',
            # MODIFICADO: Removido o padrão de procedimento, será tratado separadamente
        }
        
        # Aplicar cada padrão e armazenar os resultados
        for campo, padrao in outros_padroes.items():
            match = re.search(padrao, texto, re.IGNORECASE)
            if match:
                dados[campo] = match.group(1).strip()
        
        # NOVA FUNÇÃO: Limpar unidade executante para remover números e a palavra "CNES"
        def limpar_unidade_executante(texto):
            if not texto or texto == "NÃO ENCONTRADO":
                return texto
                
            # Remover a palavra "CNES" e variações, incluindo "Cod." e "Cod"
            texto = re.sub(r'CNES\s*:?', '', texto, flags=re.IGNORECASE)
            texto = re.sub(r'Cod\.?\s*CNES\s*:?', '', texto, flags=re.IGNORECASE)
            texto = re.sub(r'^Cod\.?\s+', '', texto, flags=re.IGNORECASE)
            
            # Extrair apenas palavras (sem números)
            palavras = re.findall(r'[A-Za-zÀ-ÖØ-öø-ÿ\s]+', texto)
            texto = ' '.join(palavras)
            
            # Normalizar espaços e remover espaços extras
            texto = re.sub(r'\s+', ' ', texto).strip()
            
            # Se o resultado for muito curto (menos de 3 caracteres), considerar como não encontrado
            if len(texto) < 3:
                return "NÃO ENCONTRADO"
                
            return texto
        
        # Aplicar limpeza à unidade executante
        dados["unidade_executante"] = limpar_unidade_executante(dados["unidade_executante"])
        
        # CORREÇÃO: Verificar se a unidade executante é "NÃO ENCONTRADO" ou está vazia
        if dados["unidade_executante"] == "NÃO ENCONTRADO" or not dados["unidade_executante"]:
            # Tentar padrões alternativos para unidade executante
            patterns = [
                # Busca por HOSPITAL seguido de texto
                r'HOSPITAL\s+([^\r\n:]+)',
                
                # Busca por nome após "UNIDADE EXECUTANTE" e "Nome:"
                r'UNIDADE\s*EXECUTANTE[\s\S]*?Nome\s*:\s*([A-Z][A-Z\s]+)',
                
                # Busca por nome após "EXECUTANTE" e "Nome:"
                r'EXECUTANTE[\s\S]*?Nome\s*:\s*([A-Z][A-Z\s]+)',
                
                # Busca por nome entre "Nome:" e "Endereço:"
                r'Nome\s*:\s*([A-Z][A-Z\s]+)(?:[\s\S]*?Endere[çc]o\s*:)',
                
                # Busca por nome após "UNIDADE EXECUTANTE"
                r'UNIDADE\s*EXECUTANTE[\s\S]*?([A-Z][A-Z\s]+(?:HOSPITAL|CLÍNICA|CENTRO|INSTITUTO)[^\r\n:]+)'
            ]
            
            for pattern in patterns:
                match = re.search(pattern, texto, re.IGNORECASE)
                if match:
                    resultado = match.group(1).strip()
                    # Limpar o resultado para remover números e a palavra "CNES"
                    resultado = limpar_unidade_executante(resultado)
                    if resultado and resultado != "NÃO ENCONTRADO":
                        dados["unidade_executante"] = resultado
                        break
                    
        # CORREÇÃO: Abordagens alternativas para unidade solicitante
        # Verificar se a unidade solicitante é "Cod. CNES" ou contém "CNES"
        if dados["unidade_solicitante"] == "Cod. CNES" or "CNES" in dados["unidade_solicitante"]:
            # Tentar padrões alternativos para unidade solicitante
            patterns = [
                # Busca por texto "COMPLEXO REGULADOR" após "Videofonista:"
                r'Videofonista\s*:\s*COMPLEXO\s+REGULADOR\s+ESTADUAL',
                
                # Busca por texto após "Videofonista:" limitado a palavras (sem números ou caracteres especiais)
                r'Videofonista\s*:\s*([A-Z][A-Z\s]+)(?:\d)',
                
                # Busca por texto após "Op. Videofonista:" limitado a palavras
                r'Op\.\s*Videofonista\s*:\s*([A-Z][A-Z\s]+)(?:\d)',
                
                # Busca por "COMPLEXO REGULADOR ESTADUAL" em qualquer lugar
                r'(COMPLEXO\s+REGULADOR\s+ESTADUAL)',
                
                # Busca por texto após "Videofonista:" até o próximo número ou caractere especial
                r'Videofonista\s*:\s*([^\r\n\d:]+)',
                
                # Fallback: busca por qualquer texto após "Videofonista:"
                r'Videofonista\s*:\s*([^\r\n:]+)'
            ]
            
            for pattern in patterns:
                match = re.search(pattern, texto, re.IGNORECASE)
                if match:
                    # Se o padrão contém grupo de captura, use-o
                    if '(' in pattern and ')' in pattern:
                        resultado = match.group(1).strip()
                    # Caso contrário, use "COMPLEXO REGULADOR ESTADUAL" como valor fixo
                    else:
                        resultado = "COMPLEXO REGULADOR ESTADUAL"
                        
                    if resultado and "Cod. CNES" not in resultado:
                        # CORREÇÃO: Armazenar em unidade_solicitante em vez de unidade_executante
                        dados["unidade_solicitante"] = resultado
                        print(f"Unidade solicitante encontrada após Videofonista: {resultado}")
                        break
                        
        # Abordagens alternativas para data do exame
        if dados["data_exame"] == "NÃO ENCONTRADO":
            patterns = [            
                # Busca por padrão de data e hora
                r'(\d{2}/\d{2}/\d{4}\s*\d{2}:\d{2})',
                
                # Busca por data após "Data de Atendimento:"
                r'Data\s*de\s*Atendimento\s*:?\s*([^\r\n]+)'
            ]
            
            for pattern in patterns:
                match = re.search(pattern, texto, re.IGNORECASE)
                if match:
                    data_encontrada = match.group(1).strip()
                    dados["data_exame"] = data_encontrada
                    break
        
        # Pós-processamento para garantir formato correto da data
        if dados["data_exame"] != "NÃO ENCONTRADO":
            # Verificar se a data está no formato DD/MM/AAAA
            data_match = re.match(r'(\d{2}/\d{2}/\d{4})', dados["data_exame"])
            if data_match:
                dados["data_exame"] = data_match.group(1)
            else:
                # Tentar extrair qualquer data no formato DD/MM/AAAA do texto encontrado
                data_match = re.search(r'(\d{2}/\d{2}/\d{4})', dados["data_exame"])
                if data_match:
                    dados["data_exame"] = data_match.group(1)

        # NOVA IMPLEMENTAÇÃO: Extração e limpeza do procedimento
        # Primeiro, extrair o procedimento bruto usando o padrão original
        procedimento_bruto = "NÃO ENCONTRADO"
        pattern = r'Procedimentos\s*Autorizados\s*:?[\s\S]*?([^\r\n]+?)(?:\s{2,}|\r|\n)'
        match = re.search(pattern, texto, re.IGNORECASE)
        if match:
            procedimento_bruto = match.group(1).strip()
            
        # NOVA FUNÇÃO: Limpar procedimento para remover códigos e números
        def limpar_procedimento(texto):
            if not texto or texto == "NÃO ENCONTRADO":
                return texto
                
            # Verificar se o texto contém "CONSULTA EM" e extrair o procedimento completo
            consulta_match = re.search(r'(CONSULTA\s+EM\s+[A-ZÀ-Úa-zà-ú\s\-]+)', texto, flags=re.IGNORECASE)
            if consulta_match:
                return consulta_match.group(1).strip()
                
            # Verificar outros tipos de procedimentos
            outros_procedimentos = [
                r'(TOMOGRAFIA\s+[A-ZÀ-Úa-zà-ú\s\-]+)',
                r'(RESSONANCIA\s+[A-ZÀ-Úa-zà-ú\s\-]+)',
                r'(EXAME\s+[A-ZÀ-Úa-zà-ú\s\-]+)'
            ]
            
            for pattern in outros_procedimentos:
                match = re.search(pattern, texto, flags=re.IGNORECASE)
                if match:
                    return match.group(1).strip()
            
            # Remover "Cod. Unificado:", "Cod. Interno:" e variações
            texto = re.sub(r'Cod\.\s*Unificado\s*:?', '', texto, flags=re.IGNORECASE)
            texto = re.sub(r'Cod\.\s*Interno\s*:?', '', texto, flags=re.IGNORECASE)
            
            # Extrair apenas palavras e hífens (sem números)
            palavras = re.findall(r'[A-ZÀ-Úa-zà-ú\s\-]+', texto)
            texto = ' '.join(palavras)
            
            # Normalizar espaços e remover espaços extras
            texto = re.sub(r'\s+', ' ', texto).strip()
            
            # Se o resultado for muito curto (menos de 3 caracteres), considerar como não encontrado
            if len(texto) < 3:
                return "NÃO ENCONTRADO"
                
            return texto
        
        # Aplicar a função de limpeza ao procedimento bruto
        dados["procedimento"] = limpar_procedimento(procedimento_bruto)
        
        # Se ainda não encontrou o procedimento, tentar padrões alternativos
        if dados["procedimento"] == "NÃO ENCONTRADO":
            patterns = [
                # Busca por CONSULTA EM seguido de texto
                r'CONSULTA\s+EM\s+([A-ZÀ-Úa-zà-ú\s\-]+)',
                
                # Busca por TOMOGRAFIA seguido de texto
                r'TOMOGRAFIA\s+([A-ZÀ-Úa-zà-ú\s\-]+)',
                
                # Busca por EXAME seguido de texto
                r'EXAME\s+([A-ZÀ-Úa-zà-ú\s\-]+)'
            ]
            
            for pattern in patterns:
                match = re.search(pattern, texto, re.IGNORECASE)
                if match:
                    # Capturar o texto completo, não apenas o grupo
                    inicio = match.start()
                    fim = match.end()
                    procedimento_completo = texto[inicio:fim].strip()
                    dados["procedimento"] = procedimento_completo
                    break
                    
        return dados
        
    except Exception as e:
        print(f"Erro ao extrair dados: {str(e)}")
        # Retornar o dicionário com valores padrão em caso de erro, mas nunca None
        return dados


def processar_pdfs(pasta):
    """
    Processa todos os arquivos PDF em uma pasta.
    
    Args:
        pasta: Caminho para a pasta contendo os PDFs
        
    Returns:
        Lista de dicionários com os dados extraídos de cada PDF
    """
    import os
    
    linhas = []
    sucessos = 0
    falhas = 0
    
    for nome_arquivo in os.listdir(pasta):
        if nome_arquivo.endswith(".pdf"):
            caminho_completo = os.path.join(pasta, nome_arquivo)
            try:
                print(f"Processando: {nome_arquivo}")
                texto = extrair_texto_pdf(caminho_completo)
                
                if not texto or not texto.strip():
                    print(f"AVISO: Não foi possível extrair texto do arquivo {nome_arquivo}")
                    falhas += 1
                    continue
                
                dados = extrair_dados(texto, nome_arquivo)
                
                # Verificar se dados é None antes de continuar
                if dados is None:
                    print(f"AVISO: Falha ao extrair dados de {nome_arquivo}")
                    falhas += 1
                    continue
                
                # Verificar se todos os campos obrigatórios foram encontrados
                campos_obrigatorios = ["codigo_solicitacao", "cns", "unidade_executante", 
                                     "unidade_solicitante", "data_exame", "procedimento"]
                campos_faltantes = [campo for campo in campos_obrigatorios 
                                  if campo not in dados or dados[campo] == "NÃO ENCONTRADO"]
                
                if campos_faltantes:
                    print(f"AVISO: Campos não encontrados em {nome_arquivo}: {', '.join(campos_faltantes)}")
                    # Adicionar mesmo com campos faltantes, mas marcar a falha
                    linhas.append(dados)
                    falhas += 1
                else:
                    linhas.append(dados)
                    sucessos += 1
                    print(f"Sucesso: Todos os campos extraídos de {nome_arquivo}")
            except Exception as e:
                print(f"ERRO ao processar {nome_arquivo}: {str(e)}")
                falhas += 1
    
    print(f"Processamento concluído: {sucessos} arquivos processados com sucesso, {falhas} falhas.")
    return linhas

def configurar_google_sheets(caminho_credenciais, nome_planilha, nome_aba):
    """
    Configura a conexão com o Google Sheets.
    
    Args:
        caminho_credenciais: Caminho para o arquivo de credenciais do Google
        nome_planilha: Nome da planilha do Google Sheets
        nome_aba: Nome da aba na planilha
        
    Returns:
        Objeto worksheet do gspread ou None em caso de erro
    """
    try:
        # Usar escopos mais amplos para garantir acesso
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        credentials = Credentials.from_service_account_file(caminho_credenciais, scopes=scopes)
        client = gspread.authorize(credentials)
        
        # Tentar abrir a planilha existente
        try:
            planilha = client.open(nome_planilha)
        except gspread.exceptions.SpreadsheetNotFound:
            print(f"Planilha '{nome_planilha}' não encontrada. Criando nova planilha...")
            planilha = client.create(nome_planilha)
            # Compartilhar com o email do usuário (opcional)
            # planilha.share('email_do_usuario@gmail.com', perm_type='user', role='writer')
        
        # Tentar abrir a aba existente ou criar nova
        try:
            aba = planilha.worksheet(nome_aba)
        except gspread.exceptions.WorksheetNotFound:
            print(f"Aba '{nome_aba}' não encontrada. Criando nova aba...")
            aba = planilha.add_worksheet(title=nome_aba, rows=1000, cols=20)
        
        return aba
    except Exception as e:
        print(f"Erro ao configurar Google Sheets: {str(e)}")
        return None


def enviar_para_sheets(dados, aba):
    """
    Versão corrigida da função enviar_para_sheets que garante que novos registros
    sejam adicionados após a última linha com dados, não após o cabeçalho.
    
    Args:
        dados: Lista de dicionários com os dados extraídos
        aba: Objeto worksheet do gspread
        
    Returns:
        Boolean indicando sucesso ou falha
    """
    if not dados:
        print("Sem dados para enviar.")
        return False
    
    if not aba:
        print("Problema com a planilha.")
        return False
    
    try:
        # Verificar se a planilha já tem cabeçalhos
        headers = aba.row_values(1)
        if not headers:
            print("ERRO: A planilha não possui cabeçalhos. Certifique-se de que a planilha tenha a estrutura correta.")
            return False
        
        # Mapear os índices das colunas que precisamos preencher
        colunas_necessarias = {
            "COD. SOLICITAÇÃO": "codigo_solicitacao",
            "CNS DO PACIENTE": "cns",
            "UNID. SOLICITANTE": "unidade_solicitante",
            "UNID. EXECUTANTE": "unidade_executante",
            "DATA DO EXAME/CONSULTA": "data_exame",
            "CONSULTA/EXAME/ESPECIALIDADE": "procedimento"
        }
        
        mapeamento_colunas = {}
        for i, header in enumerate(headers):
            if header in colunas_necessarias:
                mapeamento_colunas[colunas_necessarias[header]] = i
        
        # Verificar se todas as colunas necessárias foram encontradas
        colunas_faltantes = [col for col in colunas_necessarias.keys() 
                           if colunas_necessarias[col] not in mapeamento_colunas]
        if colunas_faltantes:
            print(f"AVISO: As seguintes colunas não foram encontradas na planilha: {', '.join(colunas_faltantes)}")
        
        # Obter todas as linhas existentes para verificar duplicatas e encontrar a última linha com dados
        todas_linhas = aba.get_all_values()
        codigos_existentes = set()
        
        # Encontrar a última linha com dados
        ultima_linha = 0  # Começar do início
        for i in range(len(todas_linhas)):
            # Verificar se a linha tem algum conteúdo
            if any(todas_linhas[i]):
                ultima_linha = i
        
        # Índice da coluna do código de solicitação
        if "codigo_solicitacao" in mapeamento_colunas:
            idx_codigo = mapeamento_colunas["codigo_solicitacao"]
            # Começar do índice 1 para pular o cabeçalho
            for i in range(1, len(todas_linhas)):
                if i < len(todas_linhas) and idx_codigo < len(todas_linhas[i]) and todas_linhas[i][idx_codigo]:
                    codigos_existentes.add(todas_linhas[i][idx_codigo])
        
        # Preparar os dados para adicionar à planilha
        novas_linhas = []
        
        for item in dados:
            # Verificar se o item é None ou não tem o campo codigo_solicitacao
            if item is None or "codigo_solicitacao" not in item:
                print(f"AVISO: Registro inválido encontrado. Pulando.")
                continue
                
            codigo = item.get("codigo_solicitacao", "")
            
            # Pular registros sem código de solicitação válido
            if codigo == "NÃO ENCONTRADO" or not codigo:
                print(f"AVISO: Registro sem código de solicitação válido. Pulando.")
                continue
            
            # Verificar se o código já existe na planilha
            if codigo in codigos_existentes:
                print(f"AVISO: Código de solicitação {codigo} já existe na planilha. Pulando.")
                continue
            
            # Criar uma nova linha com valores vazios
            nova_linha = [""] * len(headers)
            
            # Preencher apenas as colunas que precisamos
            for campo_dados, indice_coluna in mapeamento_colunas.items():
                if indice_coluna < len(nova_linha):
                    nova_linha[indice_coluna] = item.get(campo_dados, "")
            
            novas_linhas.append(nova_linha)
        
        # Adicionar as novas linhas à planilha
        if novas_linhas:
            # CORREÇÃO: Usar o método update para adicionar dados após a última linha com dados
            # Isso garante que os novos registros sejam adicionados após todos os dados existentes
            inicio_celula = f"A{ultima_linha + 2}"  # +2 porque: +1 para índice começar em 1, +1 para ir para a próxima linha
            aba.update(inicio_celula, novas_linhas, value_input_option="USER_ENTERED")
            print(f"Dados enviados com sucesso! {len(novas_linhas)} registros adicionados a partir da linha {ultima_linha + 2}.")
            return True
        else:
            print("Nenhum novo registro para adicionar à planilha.")
            return True
    except Exception as e:
        print(f"Erro ao enviar dados para o Google Sheets: {str(e)}")
        return False
    
# Método alternativo (caso o primeiro não funcione com sua versão do gspread)
def enviar_para_sheets_alternativo(dados, aba):
    """
    Método alternativo que identifica explicitamente a última linha preenchida
    e adiciona novos registros após essa linha.
    
    Args:
        dados: Lista de dicionários com os dados extraídos
        aba: Objeto worksheet do gspread
        
    Returns:
        Boolean indicando sucesso ou falha
    """
    if not dados:
        print("Sem dados para enviar.")
        return False
    
    if not aba:
        print("Problema com a planilha.")
        return False
    
    try:
        # Verificar se a planilha já tem cabeçalhos
        headers = aba.row_values(1)
        if not headers:
            print("ERRO: A planilha não possui cabeçalhos. Certifique-se de que a planilha tenha a estrutura correta.")
            return False
        
        # Mapear os índices das colunas que precisamos preencher
        colunas_necessarias = {
            "COD. SOLICITAÇÃO": "codigo_solicitacao",
            "CNS DO PACIENTE": "cns",
            "UNID. SOLICITANTE": "unidade_solicitante",
            "UNID. EXECUTANTE": "unidade_executante",
            "DATA DO EXAME/CONSULTA": "data_exame",
            "CONSULTA/EXAME/ESPECIALIDADE": "procedimento"
        }
        
        mapeamento_colunas = {}
        for i, header in enumerate(headers):
            if header in colunas_necessarias:
                mapeamento_colunas[colunas_necessarias[header]] = i
        
        # Verificar se todas as colunas necessárias foram encontradas
        colunas_faltantes = [col for col in colunas_necessarias.keys() 
                           if colunas_necessarias[col] not in mapeamento_colunas]
        if colunas_faltantes:
            print(f"AVISO: As seguintes colunas não foram encontradas na planilha: {', '.join(colunas_faltantes)}")
        
        # Obter todas as linhas existentes para verificar duplicatas
        todas_linhas = aba.get_all_values()
        codigos_existentes = set()
        
        # Encontrar a última linha preenchida
        ultima_linha = 1  # Começar após o cabeçalho
        for i in range(1, len(todas_linhas)):
            # Verificar se a linha tem algum conteúdo
            if any(todas_linhas[i]):
                ultima_linha = i + 1  # +1 porque queremos a próxima linha após a última preenchida
        
        # Índice da coluna do código de solicitação
        if "codigo_solicitacao" in mapeamento_colunas:
            idx_codigo = mapeamento_colunas["codigo_solicitacao"]
            # Começar do índice 1 para pular o cabeçalho
            for i in range(1, len(todas_linhas)):
                if i < len(todas_linhas) and idx_codigo < len(todas_linhas[i]) and todas_linhas[i][idx_codigo]:
                    codigos_existentes.add(todas_linhas[i][idx_codigo])
        
        # Preparar os dados para adicionar à planilha
        novas_linhas = []
        
        for item in dados:
            # Verificar se o item é None ou não tem o campo codigo_solicitacao
            if item is None or "codigo_solicitacao" not in item:
                print(f"AVISO: Registro inválido encontrado. Pulando.")
                continue
                
            codigo = item.get("codigo_solicitacao", "")
            
            # Pular registros sem código de solicitação válido
            if codigo == "NÃO ENCONTRADO" or not codigo:
                print(f"AVISO: Registro sem código de solicitação válido. Pulando.")
                continue
            
            # Verificar se o código já existe na planilha
            if codigo in codigos_existentes:
                print(f"AVISO: Código de solicitação {codigo} já existe na planilha. Pulando.")
                continue
            
            # Criar uma nova linha com valores vazios
            nova_linha = [""] * len(headers)
            
            # Preencher apenas as colunas que precisamos
            for campo_dados, indice_coluna in mapeamento_colunas.items():
                if indice_coluna < len(nova_linha):
                    nova_linha[indice_coluna] = item.get(campo_dados, "")
            
            novas_linhas.append(nova_linha)
        
        # Adicionar as novas linhas à planilha
        if novas_linhas:
            # CORREÇÃO: Usar o método update para adicionar dados a partir da última linha preenchida
            # Isso garante que não haverá sobreposição
            inicio_celula = f"A{ultima_linha + 1}"  # Exemplo: A2, A10, etc.
            aba.update(inicio_celula, novas_linhas, value_input_option="USER_ENTERED")
            print(f"Dados enviados com sucesso! {len(novas_linhas)} registros adicionados a partir da linha {ultima_linha + 1}.")
            return True
        else:
            print("Nenhum novo registro para adicionar à planilha.")
            return True
    except Exception as e:
        print(f"Erro ao enviar dados para o Google Sheets: {str(e)}")
        return False


def carregar_configuracao(arquivo_config='config.json'):
    """
    Carrega configurações do arquivo config.json.
    
    Args:
        arquivo_config: Caminho para o arquivo de configuração
        
    Returns:
        Dicionário com as configurações
    """
    # Valores padrão
    config = {
        "pasta_pdfs": "pdfs",
        "arquivo_credenciais": "credentials.json",
        "nome_planilha": "CENSO",
        "nome_aba": "geral"
    }
    
    # Tentar carregar do arquivo de configuração
    try:
        if os.path.exists(arquivo_config):
            with open(arquivo_config, 'r') as f:
                config_carregada = json.load(f)
                config.update(config_carregada)
    except Exception as e:
        print(f"Erro ao carregar configuração: {str(e)}")
    
    return config

def testar_extracao(pdf_path):
    """
    Função para testar a extração de um único arquivo PDF.
    
    Args:
        pdf_path: Caminho para o arquivo PDF de teste
        
    Returns:
        Dicionário com os dados extraídos
    """
    print(f"Testando extração do arquivo: {pdf_path}")
    texto = extrair_texto_pdf(pdf_path)
    
    if not texto.strip():
        print("ERRO: Não foi possível extrair texto do PDF.")
        return {}
    
    nome_arquivo = os.path.basename(pdf_path)
    dados = extrair_dados(texto, nome_arquivo)
    
    print("\nResultados da extração:")
    for campo, valor in dados.items():
        print(f"{campo}: {valor}")
    
    return dados

def verificar_dependencias():
    """
    Verifica se as dependências necessárias estão instaladas.
    
    Returns:
        Boolean indicando se todas as dependências estão disponíveis
    """
    dependencias_ok = True
    
    # Verificar PyPDF2
    try:
        import PyPDF2
        print("✓ PyPDF2 está instalado")
    except ImportError:
        print("✗ PyPDF2 não está instalado. Instalando...")
        try:
            subprocess.check_call(["pip", "install", "PyPDF2"])
            print("✓ PyPDF2 instalado com sucesso")
        except:
            print("✗ Falha ao instalar PyPDF2")
            dependencias_ok = False
    
    # Verificar pdftotext (poppler-utils)
    try:
        resultado = subprocess.run(["pdftotext", "-v"], capture_output=True, stderr=subprocess.STDOUT, text=True)
        if "pdftotext version" in resultado.stdout:
            print("✓ poppler-utils (pdftotext) está instalado")
        else:
            raise Exception("Versão não encontrada")
    except:
        print("✗ poppler-utils não está instalado ou não está no PATH")
        print("  Para instalar:")
        print("  - Ubuntu/Debian: sudo apt-get install poppler-utils")
        print("  - Windows: Baixe em https://github.com/oschwartz10612/poppler-windows/releases")
        dependencias_ok = False
    
    # Verificar Tesseract OCR
    try:
        resultado = subprocess.run(["tesseract", "--version"], capture_output=True, text=True)
        if resultado.returncode == 0:
            print("✓ Tesseract OCR está instalado")
        else:
            raise Exception("Comando retornou erro")
    except:
        print("✗ Tesseract OCR não está instalado ou não está no PATH")
        print("  Para instalar:")
        print("  - Ubuntu/Debian: sudo apt-get install tesseract-ocr tesseract-ocr-por")
        print("  - Windows: Baixe em https://github.com/UB-Mannheim/tesseract/wiki")
        dependencias_ok = False
    
    return dependencias_ok


from flask import Flask, jsonify
import os

app = Flask(__name__)

@app.route('/')
def index():
    # Executa o processamento ao acessar a rota
    config = carregar_configuracao()
    
    if not os.path.exists(config["pasta_pdfs"]):
        os.makedirs(config["pasta_pdfs"])
    
    dados_extraidos = processar_pdfs(config["pasta_pdfs"])
    
    if not dados_extraidos:
        return jsonify({"mensagem": "Nenhum dado extraído dos PDFs."})

    # Salvar localmente em CSV
    df = pd.DataFrame(dados_extraidos)
    df.to_csv("dados_extraidos.csv", index=False)
    
    if os.path.exists(config["arquivo_credenciais"]):
        aba = configurar_google_sheets(
            config["arquivo_credenciais"], 
            config["nome_planilha"], 
            config["nome_aba"]
        )
        if aba:
            enviar_para_sheets(dados_extraidos, aba)
    
    return jsonify({"mensagem": "Processamento concluído com sucesso!", "registros": dados_extraidos})

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
