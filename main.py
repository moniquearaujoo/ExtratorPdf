import os
import re
import PyPDF2
import pandas as pd
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
from flask_cors import CORS
from google_sheets_integration_fix import adicionar_dados_planilha

# Importar a função especializada para extrair município
try:
    from extract_municipio_module import extrair_municipio_residencia
except ImportError:
    # Função de fallback caso o módulo não seja encontrado
    def extrair_municipio_residencia(texto):
        # Implementação simplificada para garantir compatibilidade
        cidade_estado_pattern = r'([A-ZÀ-Ú\s]+[-–]\s*[A-Z]{2})'
        match = re.search(cidade_estado_pattern, texto)
        if match:
            return match.group(0).strip()
        return "NÃO ENCONTRADO"

app = Flask(__name__)
CORS(app)  # Habilita CORS para todas as rotas

# Configurações
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
MAX_FILE_SIZE = 2 * 1024 * 1024  # 2MB em bytes
MAX_FILES = 10

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE * MAX_FILES  # Limite total para todos os arquivos

# Criar pasta de uploads se não existir
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Verifica se o arquivo tem uma extensão permitida"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extrair_texto_pdf(pdf_path):
    """
    Extrai texto de um arquivo PDF usando apenas PyPDF2 com técnicas otimizadas.
    
    Args:
        pdf_path: Caminho para o arquivo PDF
        
    Returns:
        String contendo o texto extraído do PDF
    """
    try:
        # Método principal: PyPDF2 com configurações padrão
        texto_total = ""
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            for page_num in range(len(reader.pages)):
                texto_total += reader.pages[page_num].extract_text() + "\n"
        
        # Se conseguiu extrair texto, retorna
        if texto_total.strip():
            return texto_total
        
        # Se não conseguiu extrair texto, tenta técnicas alternativas
        texto_alternativo = ""
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            
            # Técnica 1: Extrair texto com diferentes parâmetros
            for page_num in range(len(reader.pages)):
                # Tenta extrair texto com diferentes configurações
                page = reader.pages[page_num]
                try:
                    # Tenta extrair texto ignorando LTFigures (pode ajudar em alguns PDFs)
                    texto_alternativo += page.extract_text(extraction_mode="layout", layout_mode_space_vertically=True) + "\n"
                except:
                    # Fallback para o método padrão
                    texto_alternativo += page.extract_text() + "\n"
        
        # Se conseguiu extrair texto com técnicas alternativas, retorna
        if texto_alternativo.strip():
            return texto_alternativo
        
        # Se ainda não conseguiu extrair texto, tenta extrair metadados
        with open(pdf_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            metadados = reader.metadata
            if metadados:
                texto_metadados = f"Título: {metadados.title or 'N/A'}\n"
                texto_metadados += f"Autor: {metadados.author or 'N/A'}\n"
                texto_metadados += f"Assunto: {metadados.subject or 'N/A'}\n"
                texto_metadados += f"Criador: {metadados.creator or 'N/A'}\n"
                texto_metadados += f"Produtor: {metadados.producer or 'N/A'}\n"
                return texto_metadados
        
        # Se nenhuma técnica funcionou, retorna string vazia
        return ""
    except Exception as e:
        print(f"Erro ao extrair texto com PyPDF2: {e}")
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
    import re
    import os
    
    # Inicializar o dicionário de dados com valores padrão
    dados = {
        "codigo_solicitacao": "NÃO ENCONTRADO",
        "cns": "NÃO ENCONTRADO",
        "unidade_solicitante": "NÃO ENCONTRADO",  # Manteremos o nome do campo para compatibilidade
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
        
        # Extrair código de solicitação
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
                        
                        # Se encontramos um código que começa com 5, paramos a busca
                        if codigo.startswith('5'):
                            break
        
        # Extrair unidade executante e data do exame
        outros_padroes = {
            "unidade_executante": r'UNIDADE\s*EXECUTANTE[\s\S]*?Nome\s*:\s*([^\r\n:]+)',
            "data_exame": r'Data\s*e\s*Hor[áa]rio\s*de\s*Atendimento\s*:?\s*([^\r\n]+)',
        }
        
        # Aplicar cada padrão e armazenar os resultados
        for campo, padrao in outros_padroes.items():
            match = re.search(padrao, texto, re.IGNORECASE)
            if match:
                dados[campo] = match.group(1).strip()
        
        # MODIFICADO: Extrair município de residência usando a função especializada
        municipio = extrair_municipio_residencia(texto)
        if municipio and municipio != "NÃO ENCONTRADO":
            dados["unidade_solicitante"] = municipio
        else:
            # Fallback para os padrões originais se a função especializada não encontrar o município
            municipio_patterns = [
                # Busca por padrão específico baseado na imagem do PDF
                r'Bairro\s*:\s*([A-ZÀ-Ú\s]+)Munic[íi]pio\s+de\s+Resid[êe]ncia\s*:\s*([A-ZÀ-Ú\s]+[-–]\s*[A-Z]{2})',
                
                # Busca por padrão específico baseado na imagem do PDF - versão alternativa
                r'Bairro\s*:([^:]+?)Munic[íi]pio\s+de\s+Resid[êe]ncia\s*:\s*([^C]+)CEP',
                
                # Busca por "Município de Residência:" seguido de texto até "CEP:"
                r'Munic[íi]pio\s+de\s+Resid[êe]ncia\s*:\s*([^C]+)CEP',
                
                # Busca por "Município de Residência:" seguido de texto até o próximo campo ou quebra de linha
                r'Munic[íi]pio\s+de\s+Resid[êe]ncia\s*:\s*([^\r\n:]+)',
                
                # Busca por "Município de Residência:" na seção de dados do paciente
                r'DADOS\s+DO\s+PACIENTE[\s\S]*?Munic[íi]pio\s+de\s+Resid[êe]ncia\s*:\s*([^\r\n:]+)',
                
                # Busca por texto entre "Município de Residência:" e o próximo campo
                r'Munic[íi]pio\s+de\s+Resid[êe]ncia\s*:\s*([^:]+?)(?:\s*\w+\s*:|$)',
                
                # Busca por texto após "Município:" que geralmente aparece próximo a CEP
                r'Munic[íi]pio\s*:\s*([^\r\n:]+?)(?:\s*CEP\s*:|$)'
            ]
            
            municipio_encontrado = False
            for pattern in municipio_patterns:
                match = re.search(pattern, texto, re.IGNORECASE)
                if match:
                    # Verificar se o padrão tem dois grupos (caso do padrão com Bairro)
                    if len(match.groups()) > 1 and 'Bairro' in pattern:
                        municipio = match.group(2).strip()
                    else:
                        municipio = match.group(1).strip()
                    
                    if municipio:
                        # Limpar o resultado para remover espaços extras e caracteres indesejados
                        municipio = re.sub(r'\s+', ' ', municipio).strip()
                        dados["unidade_solicitante"] = municipio
                        municipio_encontrado = True
                        break
            
            # Se não encontrou o município de residência, tentar padrões específicos da imagem do PDF
            if not municipio_encontrado:
                # Busca direta por "JOAO PESSOA - PB" ou variações
                cidade_estado_patterns = [
                    r'(JOAO\s+PESSOA\s*[-–]\s*PB)',
                    r'(JOÃO\s+PESSOA\s*[-–]\s*PB)',
                    r'(MANGABEIRA\s+JOAO\s+PESSOA\s*[-–]\s*PB)',
                    r'(MANGABEIRA\s+JOÃO\s+PESSOA\s*[-–]\s*PB)'
                ]
                
                for pattern in cidade_estado_patterns:
                    match = re.search(pattern, texto, re.IGNORECASE)
                    if match:
                        dados["unidade_solicitante"] = match.group(1).strip()
                        municipio_encontrado = True
                        break
                
                # Se ainda não encontrou, buscar por padrões genéricos de cidade-estado
                if not municipio_encontrado:
                    # Busca por padrões de cidade seguida de estado (ex: JOÃO PESSOA - PB)
                    cidade_estado_pattern = r'([A-ZÀ-Ú\s]+)\s*[-–]\s*([A-Z]{2})'
                    match = re.search(cidade_estado_pattern, texto, re.IGNORECASE)
                    if match:
                        cidade_estado = match.group(0).strip()  # Captura "JOÃO PESSOA - PB" completo
                        dados["unidade_solicitante"] = cidade_estado
                        
            # Verificação final: se ainda não encontrou, tentar extrair do texto completo
            if dados["unidade_solicitante"] == "NÃO ENCONTRADO" or not dados["unidade_solicitante"]:
                # Busca por "COMPLEXO REGULADOR" que pode ser a unidade solicitante
                match_complexo = re.search(r'(COMPLEXO\s+REGULADOR[A-ZÀ-Ú\s]+)', texto, re.IGNORECASE)
                if match_complexo:
                    dados["unidade_solicitante"] = match_complexo.group(1).strip()
        
        # FUNÇÃO CORRIGIDA: Limpar unidade executante preservando o nome "HOSPITAL"
        def limpar_unidade_executante(texto):
            if not texto or texto == "NÃO ENCONTRADO":
                return texto
                
            # Remover a palavra "CNES" e variações, incluindo "Cod." e "Cod"
            texto = re.sub(r'CNES\s*:?', '', texto, flags=re.IGNORECASE)
            texto = re.sub(r'Cod\.?\s*CNES\s*:?', '', texto, flags=re.IGNORECASE)
            texto = re.sub(r'^Cod\.?\s+', '', texto, flags=re.IGNORECASE)
            
            # Remover "Endereço" do final do texto
            texto = re.sub(r'\s*Endereço$', '', texto, flags=re.IGNORECASE)
            
            # Remover apenas números e caracteres especiais específicos
            texto = re.sub(r'[0-9]', '', texto)  # Remover números
            texto = re.sub(r'[^\w\sÀ-ÖØ-öø-ÿ]', '', texto)  # Manter letras, espaços e acentos
            
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
                r'UNIDADE\s*EXECUTANTE[\s\S]*?([A-Z][A-Z\s]+(?:HOSPITAL|CLÍNICA|CENTRO|INSTITUTO)[A-Z\s]+)'
            ]
            
            for pattern in patterns:
                match = re.search(pattern, texto, re.IGNORECASE)
                if match:
                    unidade = match.group(1).strip()
                    if unidade and len(unidade) > 3:
                        dados["unidade_executante"] = limpar_unidade_executante(unidade)
                        break
        
        # Extrair procedimento
        procedimento_patterns = [
            # Busca por procedimento após "Procedimentos Autorizados:" e "Cod. Unificado:"
            r'Procedimentos\s+Autorizados\s*:[\s\S]*?Cod\.\s*Unificado\s*:[\s\S]*?Cod\.\s*Interno\s*:\s*([^\r\n:]+)',
            
            # Busca por procedimento após "Procedimentos Autorizados:"
            r'Procedimentos\s+Autorizados\s*:[\s\S]*?([A-Z][A-Z\s\-]+)',
            
            # Busca por procedimento após "CONSULTA EM" ou "EXAME DE"
            r'(CONSULTA\s+EM\s+[A-Z\s\-]+|EXAME\s+DE\s+[A-Z\s\-]+)'
        ]
        
        for pattern in procedimento_patterns:
            match = re.search(pattern, texto, re.IGNORECASE)
            if match:
                procedimento = match.group(1).strip()
                if procedimento and len(procedimento) > 3:
                    dados["procedimento"] = procedimento
                    break
        
        return dados
        
    except Exception as e:
        print(f"Erro ao extrair dados: {e}")
        return dados

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    # Verificar se a requisição contém arquivos
    if 'files[]' not in request.files:
        return jsonify({"erro": "Nenhum arquivo enviado"}), 400
    
    files = request.files.getlist('files[]')
    
    # Verificar se há arquivos selecionados
    if not files or files[0].filename == '':
        return jsonify({"erro": "Nenhum arquivo selecionado"}), 400
    
    # Verificar se o número de arquivos não excede o limite
    if len(files) > MAX_FILES:
        return jsonify({"erro": f"Número máximo de arquivos excedido. Limite: {MAX_FILES}"}), 400
    
    # Obter o ID da planilha do Google Sheets
    id_planilha = request.form.get('id_planilha', '')
    if not id_planilha:
        return jsonify({"erro": "ID da planilha do Google Sheets não fornecido"}), 400
    
    # Processar cada arquivo
    resultados = []
    dados_extraidos = []
    
    for file in files:
        # Verificar se o arquivo tem uma extensão permitida
        if not allowed_file(file.filename):
            resultados.append({
                "arquivo": file.filename,
                "erro": "Tipo de arquivo não permitido. Apenas PDFs são aceitos."
            })
            continue
        
        # Salvar o arquivo
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        try:
            file.save(file_path)
            
            # Extrair texto do PDF
            texto = extrair_texto_pdf(file_path)
            
            if not texto:
                resultados.append({
                    "arquivo": file.filename,
                    "erro": "Não foi possível extrair texto do PDF."
                })
                continue
            
            # Extrair dados do texto
            dados = extrair_dados(texto, file.filename)
            
            # Adicionar à lista de dados extraídos
            dados_extraidos.append(dados)
            
            # Adicionar ao resultado
            resultados.append({
                "arquivo": file.filename,
                "dados": dados
            })
            
        except Exception as e:
            resultados.append({
                "arquivo": file.filename,
                "erro": f"Erro ao processar arquivo: {str(e)}"
            })
    
    # Adicionar dados à planilha do Google Sheets
    if dados_extraidos:
        try:
            resultado_planilha = adicionar_dados_planilha(id_planilha, dados_extraidos)
            
            # Verificar se houve erro na adição à planilha
            if "erro" in resultado_planilha:
                return jsonify({
                    "resultados": resultados,
                    "erro_planilha": resultado_planilha["erro"]
                }), 500
            
            # Adicionar resultado da planilha à resposta
            return jsonify({
                "resultados": resultados,
                "resultado_planilha": resultado_planilha
            })
            
        except Exception as e:
            return jsonify({
                "resultados": resultados,
                "erro_planilha": f"Erro ao adicionar dados à planilha: {str(e)}"
            }), 500
    
    return jsonify({"resultados": resultados})

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
