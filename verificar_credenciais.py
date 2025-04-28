"""
Script para verificar a configuração das credenciais do Google API
"""

import os
import json
import sys
from datetime import datetime

def verificar_credenciais(arquivo_credenciais='credentials.json'):
    """
    Verifica se o arquivo de credenciais existe e está no formato correto
    """
    print(f"\n=== Verificando credenciais do Google API ===")
    print(f"Data e hora: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Diretório atual: {os.getcwd()}")
    
    # Verificar se o arquivo existe
    if not os.path.exists(arquivo_credenciais):
        print(f"ERRO: Arquivo de credenciais '{arquivo_credenciais}' não encontrado!")
        print(f"Arquivos no diretório atual: {os.listdir()}")
        return False
    
    print(f"✓ Arquivo de credenciais '{arquivo_credenciais}' encontrado")
    
    # Verificar se o arquivo está no formato JSON válido
    try:
        with open(arquivo_credenciais, 'r') as f:
            creds_data = json.load(f)
        print(f"✓ Arquivo de credenciais é um JSON válido")
    except json.JSONDecodeError:
        print(f"ERRO: Arquivo de credenciais não é um JSON válido!")
        return False
    except Exception as e:
        print(f"ERRO ao ler o arquivo de credenciais: {str(e)}")
        return False
    
    # Verificar se o arquivo contém as chaves necessárias
    required_keys = ['type', 'project_id', 'private_key_id', 'private_key', 
                     'client_email', 'client_id', 'auth_uri', 'token_uri']
    
    missing_keys = [key for key in required_keys if key not in creds_data]
    
    if missing_keys:
        print(f"ERRO: Arquivo de credenciais não contém todas as chaves necessárias!")
        print(f"Chaves faltando: {missing_keys}")
        return False
    
    print(f"✓ Arquivo de credenciais contém todas as chaves necessárias")
    
    # Exibir informações importantes
    print(f"\nInformações da conta de serviço:")
    print(f"- Tipo: {creds_data.get('type', 'Não especificado')}")
    print(f"- Email: {creds_data.get('client_email', 'Não especificado')}")
    print(f"- Projeto: {creds_data.get('project_id', 'Não especificado')}")
    
    print("\n=== IMPORTANTE ===")
    print(f"Para que a integração funcione, você DEVE compartilhar sua planilha do Google Sheets")
    print(f"com o email da conta de serviço: {creds_data.get('client_email')}")
    print(f"E dar permissão de EDIÇÃO (não apenas visualização)")
    
    return True

if __name__ == "__main__":
    # Usar o arquivo especificado como argumento ou o padrão
    arquivo = sys.argv[1] if len(sys.argv) > 1 else 'credentials.json'
    verificar_credenciais(arquivo)
