"""
Solução aprimorada para integração com Google Sheets
Este arquivo contém uma implementação para adicionar dados a uma planilha do Google Sheets
com logs detalhados e tratamento de erros robusto
"""

import gspread
from google.oauth2.service_account import Credentials
import json
import traceback
import re
import os
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
    # Criar arquivo de log para depuração
    log_file = "google_sheets_log.txt"
    with open(log_file, "a", encoding="utf-8") as log:
        log.write(f"\n\n--- NOVA EXECUÇÃO: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ---\n")
        log.write(f"ID da planilha: {id_planilha}\n")
        log.write(f"Número de registros: {len(dados) if dados else 0}\n")
        
        try:
            # Verificar se há dados para adicionar
            if not dados:
                log.write("ERRO: Nenhum dado fornecido para adicionar à planilha\n")
                return {
                    "mensagem": "Nenhum dado fornecido para adicionar à planilha",
                    "id_planilha": id_planilha,
                    "registros_adicionados": 0
                }
            
            # Verificar se o ID da planilha foi fornecido
            if not id_planilha:
                log.write("ERRO: ID da planilha não fornecido\n")
                return {"erro": "ID da planilha não fornecido"}
            
            log.write(f"Iniciando adição de {len(dados)} registros à planilha {id_planilha}\n")
            
            # Verificar se o arquivo de credenciais existe
            if not os.path.exists(arquivo_credenciais):
                log.write(f"ERRO: Arquivo de credenciais não encontrado: {arquivo_credenciais}\n")
                log.write(f"Diretório atual: {os.getcwd()}\n")
                log.write(f"Arquivos no diretório: {os.listdir()}\n")
                return {"erro": f"Arquivo de credenciais não encontrado: {arquivo_credenciais}. Verifique se o arquivo está no diretório correto."}
            
            # Configurar as credenciais
            log.write("Configurando credenciais...\n")
            scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            
            try:
                credentials = Credentials.from_service_account_file(arquivo_credenciais, scopes=scope)
                log.write("Credenciais carregadas com sucesso\n")
            except FileNotFoundError:
                log.write(f"ERRO: Arquivo de credenciais não encontrado: {arquivo_credenciais}\n")
                return {"erro": f"Arquivo de credenciais não encontrado: {arquivo_credenciais}"}
            except json.JSONDecodeError:
                log.write(f"ERRO: Arquivo de credenciais inválido (formato JSON inválido): {arquivo_credenciais}\n")
                return {"erro": f"Arquivo de credenciais inválido (formato JSON inválido): {arquivo_credenciais}"}
            except Exception as e:
                log.write(f"ERRO ao carregar credenciais: {str(e)}\n")
                return {"erro": f"Erro ao carregar credenciais: {str(e)}"}
            
            # Obter o email da conta de serviço para referência
            service_account_email = "Email não encontrado"
            try:
                with open(arquivo_credenciais, 'r') as f:
                    creds_data = json.load(f)
                    service_account_email = creds_data.get('client_email', 'Email não encontrado')
                log.write(f"Email da conta de serviço: {service_account_email}\n")
            except Exception as e:
                log.write(f"AVISO: Não foi possível ler o email da conta de serviço: {str(e)}\n")
            
            # Autorizar o cliente
            log.write("Autorizando cliente gspread...\n")
            try:
                client = gspread.authorize(credentials)
                log.write("Cliente gspread autorizado com sucesso\n")
            except Exception as e:
                log.write(f"ERRO ao autorizar cliente gspread: {str(e)}\n")
                return {"erro": f"Erro ao autorizar cliente gspread: {str(e)}"}
            
            # Abrir a planilha pelo ID
            try:
                log.write(f"Tentando abrir planilha com ID: {id_planilha}\n")
                planilha = client.open_by_key(id_planilha)
                log.write("Planilha aberta com sucesso\n")
            except gspread.exceptions.APIError as e:
                error_message = str(e)
                log.write(f"ERRO API do Google Sheets: {error_message}\n")
                
                if "not found" in error_message.lower():
                    log.write(f"Planilha não encontrada com o ID: {id_planilha}\n")
                    return {"erro": f"Planilha não encontrada com o ID: {id_planilha}. Verifique se o ID está correto."}
                elif "permission" in error_message.lower():
                    log.write(f"Sem permissão para acessar a planilha. Certifique-se de compartilhar a planilha com {service_account_email}\n")
                    return {"erro": f"Sem permissão para acessar a planilha. Certifique-se de compartilhar a planilha com {service_account_email} e dar permissão de edição."}
                else:
                    log.write(f"Erro desconhecido ao abrir planilha: {error_message}\n")
                    return {"erro": f"Erro ao abrir planilha: {error_message}"}
            except Exception as e:
                log.write(f"ERRO ao abrir planilha: {str(e)}\n")
                return {"erro": f"Erro ao abrir planilha: {str(e)}"}
            
            # Selecionar a primeira aba (índice 0)
            try:
                log.write("Selecionando primeira aba...\n")
                aba = planilha.get_worksheet(0)
                if not aba:
                    log.write("Primeira aba não encontrada, criando nova aba...\n")
                    aba = planilha.add_worksheet(title="Dados", rows=1000, cols=20)
                    log.write("Nova aba criada com sucesso\n")
                else:
                    log.write("Primeira aba selecionada com sucesso\n")
            except Exception as e:
                log.write(f"ERRO ao selecionar aba da planilha: {str(e)}\n")
                return {"erro": f"Erro ao selecionar aba da planilha: {str(e)}"}
            
            # Verificar se a planilha está vazia ou tem cabeçalho
            try:
                log.write("Obtendo valores existentes...\n")
                todas_linhas = aba.get_all_values()
                log.write(f"Obtidos {len(todas_linhas)} linhas da planilha\n")
            except Exception as e:
                log.write(f"ERRO ao ler dados da planilha: {str(e)}\n")
                return {"erro": f"Erro ao ler dados da planilha: {str(e)}"}
            
            # Definir o mapeamento correto entre os campos dos dados e as colunas da planilha
            # Usando os nomes de colunas exatos da planilha do usuário
            mapeamento_colunas = {
                "DATA DO EXAME/CONSULTA": "data_exame",
                "COD. SOLICITAÇÃO": "codigo_solicitacao",
                "CNS DO PACIENTE": "cns",
                "UNID. SOLICITANTE": "municipio_residencia",
                "UNID. EXECUTANTE": "unidade_executante",
                "CONSULTA/EXAME/ESPECIALIDADE": "procedimento",
                # Mapeamentos alternativos para maior compatibilidade
                "DATA DA AUTORIZAÇÃO": "data_exame",
                "CÓDIGO DE SOLICITAÇÃO": "codigo_solicitacao",
                "CÓDIGO SOLICITAÇÃO": "codigo_solicitacao",
                "CNS": "cns",
                "UNIDADE SOLICITANTE": "municipio_residencia",
                "UNIDADE EXECUTANTE": "unidade_executante",
                "PROCEDIMENTO": "procedimento",
                "ESPECIALIDADE": "procedimento",
                "PAES": "procedimento"
            }
            
            # Se a planilha estiver vazia, adicionar cabeçalho
            if not todas_linhas:
                log.write("Planilha vazia, adicionando cabeçalho...\n")
                # Usar os nomes de colunas exatos da planilha do usuário
                cabecalho = ["DATA DA AUTORIZAÇÃO", "COD. SOLICITAÇÃO", "CNS DO PACIENTE", 
                            "UNID. SOLICITANTE", "UNID. EXECUTANTE", "PAES"]
                try:
                    aba.append_row(cabecalho)
                    todas_linhas = [cabecalho]
                    log.write(f"Cabeçalho adicionado: {cabecalho}\n")
                except Exception as e:
                    log.write(f"ERRO ao adicionar cabeçalho: {str(e)}\n")
                    return {"erro": f"Erro ao adicionar cabeçalho: {str(e)}"}
            else:
                # Se já existir cabeçalho, usar o existente
                cabecalho = todas_linhas[0]
                log.write(f"Cabeçalho existente: {cabecalho}\n")
            
            # Criar um mapeamento de índices para saber em qual coluna cada dado deve ser inserido
            indices_colunas = {}
            for i, nome_coluna in enumerate(cabecalho):
                for key, campo_dado in mapeamento_colunas.items():
                    if nome_coluna.strip().upper() == key.strip().upper():
                        indices_colunas[campo_dado] = i
                        break
            
            log.write(f"Mapeamento de índices de colunas: {indices_colunas}\n")
            
            # Verificar se todas as colunas necessárias foram encontradas
            campos_necessarios = ["codigo_solicitacao", "cns", "unidade_solicitante", 
                                "unidade_executante", "data_exame", "procedimento"]
            campos_faltantes = [campo for campo in campos_necessarios if campo not in indices_colunas]
            
            if campos_faltantes:
                log.write(f"AVISO: Algumas colunas necessárias não foram encontradas: {campos_faltantes}\n")
                log.write("Continuando com as colunas disponíveis...\n")
            
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
            log.write("Preparando novas linhas...\n")
            novas_linhas = []
            codigos_duplicados = []  # Lista para armazenar códigos de solicitação duplicados
            registros_invalidos = []
            
            # Determinar o índice da coluna do código de solicitação
            indice_codigo = indices_colunas.get("codigo_solicitacao")
            if indice_codigo is None:
                log.write("ERRO: Coluna para código de solicitação não encontrada na planilha\n")
                return {"erro": "Coluna para código de solicitação não encontrada na planilha. Verifique se o cabeçalho da planilha contém uma coluna para o código de solicitação."}
            
            # Verificar códigos existentes para evitar duplicatas
            codigos_existentes = []
            if todas_linhas and len(todas_linhas) > 1:  # Ignorar o cabeçalho
                for linha in todas_linhas[1:]:
                    if linha and len(linha) > indice_codigo:
                        codigo = linha[indice_codigo]
                        if codigo:
                            codigos_existentes.append(codigo)
            
            log.write(f"Códigos existentes na planilha: {len(codigos_existentes)}\n")
            
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
            
            log.write(f"Novas linhas preparadas: {len(novas_linhas)}\n")
            log.write(f"Códigos duplicados: {len(codigos_duplicados)}\n")
            log.write(f"Registros inválidos: {len(registros_invalidos)}\n")
            
            # Se não houver novas linhas para adicionar, retornar mensagem simplificada
            if not novas_linhas and codigos_duplicados:
                log.write("Nenhuma nova linha para adicionar, apenas documentos duplicados.\n")
                
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
                log.write("Nenhum novo registro para adicionar à planilha\n")
                return {
                    "mensagem": "Nenhum novo registro para adicionar à planilha",
                    "id_planilha": id_planilha,
                    "registros_adicionados": 0
                }
            
            # Encontrar a última linha preenchida
            ultima_linha = len(todas_linhas)
            log.write(f"Última linha preenchida: {ultima_linha}\n")
            
            # Adicionar as novas linhas após a última linha preenchida
            inicio_celula = f"A{ultima_linha + 1}"
            log.write(f"Adicionando {len(novas_linhas)} novas linhas a partir de {inicio_celula}...\n")
            
            try:
                aba.update(inicio_celula, novas_linhas, value_input_option="USER_ENTERED")
                log.write("Dados adicionados com sucesso à planilha\n")
            except gspread.exceptions.APIError as e:
                error_message = str(e)
                log.write(f"ERRO API do Google Sheets ao adicionar dados: {error_message}\n")
                
                if "permission" in error_message.lower():
                    return {"erro": f"Sem permissão para editar a planilha. Certifique-se de compartilhar a planilha com {service_account_email} e dar permissão de edição."}
                else:
                    return {"erro": f"Erro ao adicionar dados à planilha: {error_message}"}
            except Exception as e:
                log.write(f"ERRO ao adicionar dados à planilha: {str(e)}\n")
                return {"erro": f"Erro ao adicionar dados à planilha: {str(e)}"}
            
            # Criar mensagem de resultado com formato simplificado
            mensagem = f"Registros adicionados: {len(novas_linhas)}"
            
            # Se houver documentos duplicados, adicionar à mensagem
            if codigos_duplicados:
                codigos_duplicados_str = ", ".join(codigos_duplicados)
                mensagem = f"Documento {codigos_duplicados_str} já se encontra na planilha. {mensagem}"
            
            log.write(f"Operação concluída com sucesso: {mensagem}\n")
            return {
                "mensagem": mensagem,
                "id_planilha": id_planilha,
                "registros_adicionados": len(novas_linhas),
                "codigos_duplicados": codigos_duplicados
            }
            
        except Exception as e:
            erro_detalhado = traceback.format_exc()
            log.write(f"ERRO DETALHADO: {erro_detalhado}\n")
            return {"erro": f"Erro ao adicionar dados à planilha: {str(e)}"}
    