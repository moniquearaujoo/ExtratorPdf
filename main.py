import os
import re
import PyPDF2
import pandas as pd
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
from flask_cors import CORS
from google_sheets_integration_fix import adicionar_dados_planilha


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
    from cidades_paraiba import CidadesParaiba

    # Inicializar o dicionário de dados com valores padrão
    dados = {
        "codigo_solicitacao": "NÃO ENCONTRADO",
        "cns": "NÃO ENCONTRADO",
        "unidade_solicitante": "NÃO ENCONTRADO", 
        "unidade_executante": "NÃO ENCONTRADO",
        "data_exame": "NÃO ENCONTRADO",
        "procedimento": "NÃO ENCONTRADO"
    }

    print("\n" + "="*50)
    print(f"CONTEÚDO COMPLETO DO PDF: {nome_arquivo}")
    print("="*50)
    print(texto)
    print("="*50 + "\n")

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
            r'Código da Solicitação:\s*Situação Atual:\s*(5\d{8})',
            
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

            # Busca por número de 9 dígitos que começa com 5 após "Situação Atual"
            r'Situação+Atual*:?[\s\S]{0,100}?(5\d{8})\b',
            
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
            # Novo padrão para unidade_executante
            "unidade_executante": r'UNIDADE\s*EXECUTANTE[\s\S]*?Nome\s*:\s*([A-Z\sÀ-ÖØ-öø-ÿ]+?)(?:\s*Endereço|\s*C[óo]d\.\s*CNES|\s*Número|\s*Telefone|\s*Op\.\s*Autorizador|\s*Vaga\s*Consumida|$)',
            "data_exame": r'Data\s*e\s*Hor[áa]rio\s*de\s*Atendimento\s*:?\s*([^\r\n]+)',
        }
        
        # Aplicar cada padrão e armazenar os resultados
        for campo, padrao in outros_padroes.items():
            match = re.search(padrao, texto, re.IGNORECASE)
            if match:
                dados[campo] = match.group(1).strip()
        
        
        # FUNÇÃO CORRIGIDA: Limpar unidade executante preservando o nome "HOSPITAL"
        def limpar_unidade_executante(texto):
            if not texto or texto == "NÃO ENCONTRADO":
                return texto
            
            # 1. Encontrar o primeiro dígito no texto
            # Se um dígito for encontrado, desconsiderar tudo a partir dele.
            # Isso é baseado na premissa de que nomes de unidades executantes não contêm números.
            match_digito = re.search(r'\d', texto)
            if match_digito:
                texto = texto[:match_digito.start()]
            
            # 2. Remover qualquer caractere que não seja letra, espaço ou hífen.
            # Isso limpa pontuações, símbolos e outros caracteres especiais que possam ter sobrado.
            # Mantemos o hífen pois pode aparecer em nomes (ex: "Hospital São José-Maria").
            # O flag re.UNICODE é importante para lidar corretamente com caracteres acentuados.
            texto = re.sub(r'[^\w\sÀ-ÖØ-öø-ÿ-]', '', texto, flags=re.UNICODE)
            
            # 3. Normalizar espaços: substituir múltiplos espaços por um único e remover espaços nas extremidades.
            texto = re.sub(r'\s+', ' ', texto).strip()
            
            # 4. Validação final: garantir que o resultado seja um nome válido.
            # Se o texto resultante for muito curto ou não contiver nenhuma letra,
            # consideramos que a extração falhou e retornamos "NÃO ENCONTRADO".
            if len(texto) < 3 or not re.search(r'[A-ZÀ-ÖØ-öø-ÿ]', texto, flags=re.IGNORECASE):
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
                    # CORREÇÃO: Capturar o texto completo para preservar "HOSPITAL"
                    if "HOSPITAL" in pattern:
                        resultado = "HOSPITAL " + match.group(1).strip()
                    else:
                        resultado = match.group(1).strip()
                    
                    # Limpar o resultado para remover números e a palavra "CNES"
                    resultado = limpar_unidade_executante(resultado)
                    if resultado and resultado != "NÃO ENCONTRADO":
                        dados["unidade_executante"] = resultado
                        break
            
            # NOVA CORREÇÃO: Buscar diretamente por padrões de hospital no texto
            if dados["unidade_executante"] == "NÃO ENCONTRADO" or not dados["unidade_executante"]:
                hospital_patterns = [
                    r'HOSPITAL\s+([A-ZÀ-Úa-zà-ú\s]+)',
                    r'HOSPITAL\s+DE\s+([A-ZÀ-Úa-zà-ú\s]+)',
                    r'HOSPITAL\s+([A-ZÀ-Úa-zà-ú\s]+)(?:[\s\S]*?Endere[çc]o\s*:)'
                ]
                
                for pattern in hospital_patterns:
                    match = re.search(pattern, texto, re.IGNORECASE)
                    if match:
                        # Capturar o texto completo incluindo "HOSPITAL"
                        inicio = match.start()
                        fim = match.end()
                        hospital_completo = texto[inicio:fim].strip()
                        
                        # Limpar o resultado
                        hospital_completo = limpar_unidade_executante(hospital_completo)
                        if hospital_completo and hospital_completo != "NÃO ENCONTRADO":
                            dados["unidade_executante"] = hospital_completo
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


        # Extrair e limpar procedimento
        procedimento_bruto = "NÃO ENCONTRADO"
        pattern = r'Procedimentos\s*Autorizados\s*:?[\s\S]*?([^\r\n]+?)(?:\s{2,}|\r|\n)'
        match = re.search(pattern, texto, re.IGNORECASE)
        if match:
            procedimento_bruto = match.group(1).strip()
            
        # Limpar procedimento
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
                
                 # NOVO TRECHO: Pegar Município de Residência e colocar como Unidade Solicitante
                 # NOVO TRECHO: Pegar Município de Residência e colocar como Unidade Solicitante
        try:
            municipio_residencia = "" 
            
            # Padrão principal: Captura o texto após "Município de Residência:"
            # (mantido o padrão que funcionou para o caso anterior)
            padrao_municipio = r'Município\s*(?:de)?\s*Resid[êe]ncia\s*:\s*([^\r\n]+?)(?:\s*\d{5}-\d{3}|\s*Telefone\(s\):|\s*Laudo\s*/\s*Justificativa:|\s*DADOS\s*DA\s*SOLICITA[ÇC][ÃA]O|$)'
            
            match = re.search(padrao_municipio, texto, re.IGNORECASE)
            if match:
                municipio_residencia = match.group(1).strip()
                
                # REMOÇÃO DE "BRASILEIRA": Adicione esta linha
                municipio_residencia = re.sub(r'BRASILEIRA\s*', '', municipio_residencia, flags=re.IGNORECASE).strip()
                
                # Remover a palavra "BRASIL" (se ainda estiver presente por algum motivo)
                municipio_residencia = re.sub(r'BRASIL\s*', '', municipio_residencia, flags=re.IGNORECASE).strip()
                
                # Remover "PB" se estiver grudado no final (ex: "JOAO PESSOA - PB")
                municipio_residencia = re.sub(r'\s*-\s*PB\s*$', '', municipio_residencia, flags=re.IGNORECASE).strip()
                
                # Normalizar espaços após as remoções
                municipio_residencia = re.sub(r'\s+', ' ', municipio_residencia).strip()

            # Se o primeiro padrão não encontrar, tentar o padrão de fallback
            if not municipio_residencia: 
                trecho_pos_municipio = re.search(r'Município\s*de\s*Residência\s*:?(.*?)(?:\d{5}-\d{3}|Telefone\(s\):)', texto, re.IGNORECASE)
                if trecho_pos_municipio:
                    trecho = trecho_pos_municipio.group(1)
                    municipio = re.search(r'([A-ZÀ-Úa-zà-ú\s]+(?:\s*-\s*[A-Z]{2})?)', trecho)
                    if municipio:
                        municipio_residencia = municipio.group(1).strip()
                        # Aplicar as mesmas limpezas ao resultado do fallback
                        municipio_residencia = re.sub(r'BRASILEIRA\s*', '', municipio_residencia, flags=re.IGNORECASE).strip() # Adicionar aqui também
                        municipio_residencia = re.sub(r'BRASIL\s*', '', municipio_residencia, flags=re.IGNORECASE).strip()
                        municipio_residencia = re.sub(r'\s*-\s*PB\s*$', '', municipio_residencia, flags=re.IGNORECASE).strip()
                        municipio_residencia = re.sub(r'\s+', ' ', municipio_residencia).strip()
            
            # Depois de encontrar e limpar o municipio_residencia no texto:
            # Validar o município usando CidadesParaiba.validar_municipio
            municipio_validado = CidadesParaiba.validar_municipio(municipio_residencia)
            
            # Atribui o valor validado ou "NÃO ENCONTRADO" se a validação falhar
            dados["unidade_solicitante"] = municipio_validado if municipio_validado else "NÃO ENCONTRADO"

        except Exception as e:
            print(f"Erro ao extrair município para unidade solicitante: {str(e)}")
            dados["unidade_solicitante"] = "NÃO ENCONTRADO"


        # Mostrar os campos extraídos no console
        print("\nCAMPOS EXTRAÍDOS PARA PLANILHA:")
        print(f"Código de Solicitação: {dados['codigo_solicitacao']}")
        print(f"CNS do Paciente: {dados['cns']}")
        print(f"Unidade Executante: {dados['unidade_executante']}")
        print(f"Unidade Solicitante: {dados['unidade_solicitante']}")
        print(f"Data do Exame: {dados['data_exame']}")
        print(f"Procedimento: {dados['procedimento']}")
        print("\n")
    
        return dados
        
    except Exception as e:
        print(f"Erro ao extrair dados: {str(e)}")
        # Retornar o dicionário com valores padrão em caso de erro
        return dados

def processar_pdf(pdf_path):
    """
    Processa um único arquivo PDF.
    
    Args:
        pdf_path: Caminho para o arquivo PDF
        
    Returns:
        Dicionário com os dados extraídos ou None em caso de erro
    """
    try:
        nome_arquivo = os.path.basename(pdf_path)
        
        # Extrair texto do PDF
        texto = extrair_texto_pdf(pdf_path)
        
        # Verificar se conseguiu extrair texto
        if not texto.strip():
            return {"erro": f"Não foi possível extrair texto do PDF {nome_arquivo}", "arquivo": nome_arquivo}
        
        # Extrair dados do texto
        dados = extrair_dados(texto, nome_arquivo)
        
        # Adicionar o nome do arquivo aos dados
        dados["arquivo"] = nome_arquivo
        
        return dados
        
    except Exception as e:
        return {"erro": str(e), "arquivo": os.path.basename(pdf_path)}

@app.route('/')
def index():
    """Rota principal que renderiza a página de upload"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """
    Rota para processar o upload de arquivos PDF
    
    Recebe arquivos PDF via formulário, processa-os e retorna os dados extraídos
    """
    # Verificar se a requisição contém arquivos
    if 'files[]' not in request.files:
        return jsonify({"erro": "Nenhum arquivo enviado"}), 400
    
    files = request.files.getlist('files[]')
    
    # Verificar se há arquivos
    if not files or files[0].filename == '':
        return jsonify({"erro": "Nenhum arquivo selecionado"}), 400
    
    # Verificar o número de arquivos
    if len(files) > MAX_FILES:
        return jsonify({"erro": f"Número máximo de arquivos excedido. Limite: {MAX_FILES}"}), 400
    
    # Processar cada arquivo
    resultados = []
    sucessos = 0
    falhas = 0
    
    for file in files:
        # Verificar se é um arquivo permitido
        if not allowed_file(file.filename):
            resultados.append({"erro": f"Tipo de arquivo não permitido: {file.filename}", "arquivo": file.filename})
            falhas += 1
            continue
        
        # Verificar o tamanho do arquivo
        file.seek(0, os.SEEK_END)
        file_size = file.tell()
        file.seek(0)
        
        if file_size > MAX_FILE_SIZE:
            resultados.append({
                "erro": f"Tamanho do arquivo excede o limite de {MAX_FILE_SIZE/1024/1024:.1f}MB: {file.filename}", 
                "arquivo": file.filename
            })
            falhas += 1
            continue
        
        # Salvar o arquivo
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        
        # Processar o arquivo
        resultado = processar_pdf(file_path)
        
        # Verificar se houve erro no processamento
        if "erro" in resultado:
            falhas += 1
        else:
            sucessos += 1
        
        resultados.append(resultado)
    
    # Retornar os resultados
    return jsonify({
        "resultados": resultados,
        "estatisticas": {
            "total": len(files),
            "sucessos": sucessos,
            "falhas": falhas
        }
    })

@app.route('/download/csv', methods=['POST'])
def download_csv():
    """
    Rota para gerar e baixar um arquivo CSV com os dados extraídos
    
    Recebe os dados extraídos via JSON e retorna um arquivo CSV
    """
    # Obter os dados do corpo da requisição
    dados = request.json.get('dados', [])
    
    if not dados:
        return jsonify({"erro": "Nenhum dado fornecido"}), 400
    
    # Criar DataFrame com os dados
    df = pd.DataFrame(dados)
    
    # Salvar DataFrame como CSV
    csv_file = os.path.join(app.config['UPLOAD_FOLDER'], 'dados_extraidos.csv')
    df.to_csv(csv_file, index=False)
    
    # Retornar o arquivo CSV
    return send_from_directory(app.config['UPLOAD_FOLDER'], 'dados_extraidos.csv', as_attachment=True)

@app.route('/download/excel', methods=['POST'])
def download_excel():
    """
    Rota para gerar e baixar um arquivo Excel com os dados extraídos
    
    Recebe os dados extraídos via JSON e retorna um arquivo Excel
    """
    # Obter os dados do corpo da requisição
    dados = request.json.get('dados', [])
    
    if not dados:
        return jsonify({"erro": "Nenhum dado fornecido"}), 400
    
    # Criar DataFrame com os dados
    df = pd.DataFrame(dados)
    
    # Salvar DataFrame como Excel
    excel_file = os.path.join(app.config['UPLOAD_FOLDER'], 'dados_extraidos.xlsx')
    df.to_excel(excel_file, index=False)
    
    # Retornar o arquivo Excel
    return send_from_directory(app.config['UPLOAD_FOLDER'], 'dados_extraidos.xlsx', as_attachment=True)

@app.route('/planilha', methods=['POST'])
def adicionar_planilha():
      # Obter os dados do corpo da requisição
    dados = request.json.get('dados', [])
    id_planilha = request.json.get('id_planilha', '')
    
    # Usar a função de integração
    resultado = adicionar_dados_planilha(id_planilha, dados)
    
    # Verificar se houve erro
    if "erro" in resultado:
        return jsonify(resultado), 500
    
    return jsonify(resultado)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)