"""
Solução simplificada para integração com Google Sheets
Este arquivo contém uma implementação para adicionar dados a uma planilha do Google Sheets
com alertas simplificados para documentos duplicados
"""

import gspread
from google.oauth2.service_account import Credentials
import json
import traceback
import re
from datetime import datetime

def adicionar_dados_planilha(id_planilha, dados, arquivo_credenciais='credentials.json'):
    """
    Adiciona dados a uma planilha do Google Sheets com tratamento de erros robusto,
    mapeamento correto de colunas e alertas simplificados para documentos duplicados
    
    Args:
        id_planilha: ID da planilha do Google Sheets
        dados: Lista de dicionários com os dados a serem adicionados
        arquivo_credenciais: Caminho para o arquivo JSON de credenciais
        
    Returns:
        Dicionário com o resultado da operação
    """
    try:
        # Verificar se há dados para adicionar
        if not dados:
            return {
                "mensagem": "Nenhum dado fornecido para adicionar à planilha",
                "id_planilha": id_planilha,
                "registros_adicionados": 0
            }
        
        # Verificar se o ID da planilha foi fornecido
        if not id_planilha:
            return {"erro": "ID da planilha não fornecido"}
        
        print(f"Iniciando adição de {len(dados)} registros à planilha {id_planilha}")
        
        # Configurar as credenciais
        print("Configurando credenciais...")
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        try:
            credentials = Credentials.from_service_account_file(arquivo_credenciais, scopes=scope)
        except FileNotFoundError:
            return {"erro": f"Arquivo de credenciais não encontrado: {arquivo_credenciais}"}
        except json.JSONDecodeError:
            return {"erro": f"Arquivo de credenciais inválido: {arquivo_credenciais}"}
        
        # Obter o email da conta de serviço para referência
        try:
            with open(arquivo_credenciais, 'r') as f:
                creds_data = json.load(f)
                service_account_email = creds_data.get('client_email', 'Email não encontrado')
            print(f"Email da conta de serviço: {service_account_email}")
        except Exception as e:
            print(f"Aviso: Não foi possível ler o email da conta de serviço: {str(e)}")
        
        # Autorizar o cliente
        print("Autorizando cliente gspread...")
        client = gspread.authorize(credentials)
        
        # Abrir a planilha pelo ID
        try:
            print(f"Tentando abrir planilha com ID: {id_planilha}")
            planilha = client.open_by_key(id_planilha)
        except gspread.exceptions.APIError as e:
            if "not found" in str(e).lower():
                return {"erro": f"Planilha não encontrada com o ID: {id_planilha}"}
            elif "permission" in str(e).lower():
                return {"erro": f"Sem permissão para acessar a planilha. Certifique-se de compartilhar a planilha com {service_account_email}"}
            else:
                return {"erro": f"Erro ao abrir planilha: {str(e)}"}
        
        # Selecionar a primeira aba (índice 0)
        try:
            print("Selecionando primeira aba...")
            aba = planilha.get_worksheet(0)
            if not aba:
                print("Primeira aba não encontrada, criando nova aba...")
                aba = planilha.add_worksheet(title="Dados", rows=1000, cols=20)
        except Exception as e:
            return {"erro": f"Erro ao selecionar aba da planilha: {str(e)}"}
        
        # Verificar se a planilha está vazia ou tem cabeçalho
        try:
            print("Obtendo valores existentes...")
            todas_linhas = aba.get_all_values()
        except Exception as e:
            return {"erro": f"Erro ao ler dados da planilha: {str(e)}"}
        
        # Definir o mapeamento correto entre os campos dos dados e as colunas da planilha
        # Usando os nomes de colunas exatos da planilha do usuário
        mapeamento_colunas = {
            "DATA DO EXAME/CONSULTA": "data_exame",
            "COD. SOLICITAÇÃO": "codigo_solicitacao",
            "CNS DO PACIENTE": "cns",
            "UNID. SOLICITANTE": "unidade_solicitante",
            "UNID. EXECUTANTE": "unidade_executante",
            "CONSULTA/EXAME/ESPECIALIDADE": "procedimento"
        }
        
        # Se a planilha estiver vazia, adicionar cabeçalho
        if not todas_linhas:
            print("Planilha vazia, adicionando cabeçalho...")
            # Usar os nomes de colunas exatos da planilha do usuário
            cabecalho = ["DATA DA AUTORIZAÇÃO", "COD. SOLICITAÇÃO", "CNS DO PACIENTE", 
                         "UNID. SOLICITANTE", "UNID. EXECUTANTE", "PAES"]
            try:
                aba.append_row(cabecalho)
                todas_linhas = [cabecalho]
            except Exception as e:
                return {"erro": f"Erro ao adicionar cabeçalho: {str(e)}"}
        else:
            # Se já existir cabeçalho, usar o existente
            cabecalho = todas_linhas[0]
            print(f"Cabeçalho existente: {cabecalho}")
        
        # Criar um mapeamento de índices para saber em qual coluna cada dado deve ser inserido
        indices_colunas = {}
        for i, nome_coluna in enumerate(cabecalho):
            campo_dado = mapeamento_colunas.get(nome_coluna)
            if campo_dado:
                indices_colunas[campo_dado] = i
        
        print(f"Mapeamento de índices de colunas: {indices_colunas}")
        
        # Verificar se todas as colunas necessárias foram encontradas
        campos_necessarios = ["codigo_solicitacao", "cns", "unidade_solicitante", 
                             "unidade_executante", "data_exame", "procedimento"]
        campos_faltantes = [campo for campo in campos_necessarios if campo not in indices_colunas]
        
        if campos_faltantes:
            print(f"AVISO: Algumas colunas necessárias não foram encontradas: {campos_faltantes}")
            print("Continuando com as colunas disponíveis...")
        
        # Função para validar e formatar dados
        def validar_e_formatar_dados(dado):
            resultado = {}
            
            # Validar e formatar código de solicitação
            codigo = dado.get("codigo_solicitacao", "")
            if codigo and codigo != "NÃO ENCONTRADO":
                # Remover espaços e caracteres não numéricos
                codigo_limpo = re.sub(r'[^0-9]', '', codigo)
                if len(codigo_limpo) > 0:
                    resultado["codigo_solicitacao"] = codigo_limpo
                else:
                    resultado["codigo_solicitacao"] = ""
            else:
                resultado["codigo_solicitacao"] = ""
            
            # Validar e formatar CNS
            cns = dado.get("cns", "")
            if cns and cns != "NÃO ENCONTRADO":
                # Remover espaços e caracteres não numéricos
                cns_limpo = re.sub(r'[^0-9]', '', cns)
                if len(cns_limpo) > 0:
                    resultado["cns"] = cns_limpo
                else:
                    resultado["cns"] = ""
            else:
                resultado["cns"] = ""
            
            # Validar e formatar data
            data = dado.get("data_exame", "")
            if data and data != "NÃO ENCONTRADO":
                # Verificar se a data está no formato DD/MM/AAAA
                match = re.search(r'(\d{2}/\d{2}/\d{4})', data)
                if match:
                    resultado["data_exame"] = match.group(1)
                else:
                    # Tentar outros formatos comuns
                    try:
                        # Tentar interpretar a data em vários formatos
                        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%d.%m.%Y'):
                            try:
                                dt = datetime.strptime(data, fmt)
                                resultado["data_exame"] = dt.strftime('%d/%m/%Y')
                                break
                            except ValueError:
                                continue
                        else:
                            resultado["data_exame"] = data  # Manter o original se não conseguir converter
                    except:
                        resultado["data_exame"] = data  # Manter o original em caso de erro
            else:
                resultado["data_exame"] = ""
            
            # Copiar outros campos sem validação específica
            for campo in ["unidade_solicitante", "unidade_executante", "procedimento"]:
                valor = dado.get(campo, "")
                if valor and valor != "NÃO ENCONTRADO":
                    resultado[campo] = valor
                else:
                    resultado[campo] = ""
            
            return resultado
        
        # Preparar as linhas para adicionar
        print("Preparando novas linhas...")
        novas_linhas = []
        codigos_duplicados = []  # Lista para armazenar códigos de solicitação duplicados
        registros_invalidos = []
        
        # Determinar o índice da coluna do código de solicitação
        indice_codigo = indices_colunas.get("codigo_solicitacao")
        if indice_codigo is None:
            return {"erro": "Coluna para código de solicitação não encontrada na planilha"}
        
        # Verificar códigos existentes para evitar duplicatas
        codigos_existentes = []
        if todas_linhas and len(todas_linhas) > 1:  # Ignorar o cabeçalho
            for linha in todas_linhas[1:]:
                if linha and len(linha) > indice_codigo:
                    codigo = linha[indice_codigo]
                    if codigo:
                        codigos_existentes.append(codigo)
        
        for dado in dados:
            # Pular se for um erro
            if "erro" in dado:
                registros_invalidos.append({"erro": dado["erro"]})
                continue
            
            # Validar e formatar os dados
            dado_validado = validar_e_formatar_dados(dado)
            
            # Verificar se o código de solicitação é válido
            codigo_solicitacao = dado_validado.get("codigo_solicitacao", "")
            if not codigo_solicitacao:
                registros_invalidos.append({"erro": "Código de solicitação inválido ou não encontrado", "dados": dado})
                continue
                
            # Verificar se o código de solicitação já existe na planilha
            if codigo_solicitacao in codigos_existentes:
                codigos_duplicados.append(codigo_solicitacao)
                continue
            
            # Criar uma linha vazia com o mesmo número de colunas que o cabeçalho
            nova_linha = [""] * len(cabecalho)
            
            # Preencher apenas as colunas que temos mapeamento
            for campo, indice in indices_colunas.items():
                if indice < len(nova_linha):
                    nova_linha[indice] = dado_validado.get(campo, "")
            
            novas_linhas.append(nova_linha)
        
        # Se não houver novas linhas para adicionar, retornar mensagem simplificada
        if not novas_linhas and codigos_duplicados:
            print("Nenhuma nova linha para adicionar, apenas documentos duplicados.")
            
            # Criar mensagem simplificada conforme solicitado pelo usuário
            codigos_duplicados_str = ", ".join(codigos_duplicados)
            mensagem = f"Documento {codigos_duplicados_str} já se encontra na planilha. Registros adicionados: 0"
            
            return {
                "mensagem": mensagem,
                "id_planilha": id_planilha,
                "registros_adicionados": 0,
                "codigos_duplicados": codigos_duplicados
            }
        elif not novas_linhas:
            return {
                "mensagem": "Nenhum novo registro para adicionar à planilha",
                "id_planilha": id_planilha,
                "registros_adicionados": 0
            }
        
        # Encontrar a última linha preenchida
        ultima_linha = len(todas_linhas)
        print(f"Última linha preenchida: {ultima_linha}")
        
        # Adicionar as novas linhas após a última linha preenchida
        inicio_celula = f"A{ultima_linha + 1}"
        print(f"Adicionando {len(novas_linhas)} novas linhas a partir de {inicio_celula}...")
        
        try:
            aba.update(inicio_celula, novas_linhas, value_input_option="USER_ENTERED")
        except Exception as e:
            return {"erro": f"Erro ao adicionar dados à planilha: {str(e)}"}
        
        # Criar mensagem de resultado com formato simplificado
        mensagem = f"Registros adicionados: {len(novas_linhas)}"
        
        # Se houver documentos duplicados, adicionar à mensagem
        if codigos_duplicados:
            codigos_duplicados_str = ", ".join(codigos_duplicados)
            mensagem = f"Documento {codigos_duplicados_str} já se encontra na planilha. {mensagem}"
        
        print("Dados adicionados com sucesso!")
        return {
            "mensagem": mensagem,
            "id_planilha": id_planilha,
            "registros_adicionados": len(novas_linhas),
            "codigos_duplicados": codigos_duplicados
        }
        
    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print(f"ERRO DETALHADO: {erro_detalhado}")
        return {"erro": f"Erro ao adicionar dados à planilha: {str(e)}"}
