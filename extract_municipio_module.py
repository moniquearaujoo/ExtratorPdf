import re

def extrair_municipio_residencia(texto):
    """
    Função específica para extrair o município de residência do texto extraído do PDF.
    
    Args:
        texto: Texto extraído do PDF
        
    Returns:
        String contendo o município de residência extraído
    """
    # Lista de cidades conhecidas com seus estados
    cidades_conhecidas = [
        "JOAO PESSOA - PB", 
        "JOÃO PESSOA - PB",
        "SANTA RITA - PB",
        "CABEDELO - PB",
        "BAYEUX - PB",
        "CAMPINA GRANDE - PB"
    ]
    
    # Primeiro, tentar encontrar cidades conhecidas diretamente no texto
    for cidade in cidades_conhecidas:
        if cidade in texto:
            # Verificar se a cidade está concatenada com outro texto
            # Procurar por qualquer texto seguido imediatamente pela cidade
            pattern = r'([A-ZÀ-Ú]+)(' + re.escape(cidade) + r')'
            match = re.search(pattern, texto)
            if match:
                # Retorna apenas a cidade (grupo 2)
                return match.group(2).strip()
            else:
                # Se não está concatenada, retorna a cidade diretamente
                return cidade
    
    # Padrões para identificar municípios com formato "CIDADE - UF"
    municipio_patterns = [
        # Padrão específico para o formato encontrado no PDF - busca por cidade-estado após CEP
        r'CEP\s*:\s*([A-ZÀ-Ú]+)([A-ZÀ-Ú\s]+[-–]\s*[A-Z]{2})',
        
        # Padrão para capturar o município quando está concatenado com o bairro
        r'Bairro\s*:\s*Munic[íi]pio\s+de\s+Resid[êe]ncia\s*:\s*CEP\s*:\s*([A-ZÀ-Ú]+)([A-ZÀ-Ú\s]+[-–]\s*[A-Z]{2})',
        
        # Padrão para capturar o formato "CIDADE - UF" após "Município de Residência"
        r'Munic[íi]pio\s+de\s+Resid[êe]ncia\s*:\s*CEP\s*:\s*([A-ZÀ-Ú]+)([A-ZÀ-Ú\s]+[-–]\s*[A-Z]{2})'
    ]
    
    # Se não encontrou cidades conhecidas, tentar os padrões de regex
    for pattern in municipio_patterns:
        match = re.search(pattern, texto)
        if match:
            # Se o padrão tem dois grupos (caso do padrão com bairro concatenado)
            if len(match.groups()) > 1:
                # Verificar se o segundo grupo é um formato de cidade-estado
                cidade_estado = match.group(2).strip()
                if re.search(r'[A-ZÀ-Ú\s]+[-–]\s*[A-Z]{2}', cidade_estado):
                    return cidade_estado
                else:
                    # Se não for, tentar combinar o bairro com a cidade
                    bairro = match.group(1).strip()
                    # Verificar se o bairro está concatenado com a cidade
                    for cidade in cidades_conhecidas:
                        if cidade in bairro + cidade_estado:
                            return cidade
            else:
                # Caso contrário, retorna o grupo encontrado
                return match.group(1).strip()
    
    # Abordagem específica para o formato "BAIRROCIDADE - UF"
    # Tentar identificar padrões conhecidos de cidades no texto
    for cidade in cidades_conhecidas:
        # Extrair apenas a parte da cidade sem o estado
        cidade_sem_estado = cidade.split(" - ")[0]
        # Procurar por qualquer texto seguido pela cidade sem estado
        pattern = r'([A-ZÀ-Ú]+)(' + re.escape(cidade_sem_estado) + r')'
        match = re.search(pattern, texto)
        if match:
            # Retorna a cidade completa com estado
            return cidade
    
    # Tentar encontrar qualquer padrão de "CIDADE - UF" no texto
    cidade_uf_pattern = r'([A-ZÀ-Ú\s]+[-–]\s*[A-Z]{2})'
    matches = re.findall(cidade_uf_pattern, texto)
    if matches:
        # Filtrar apenas os que parecem ser cidades (não números ou códigos)
        cidades_possiveis = [m for m in matches if not re.search(r'\d', m)]
        if cidades_possiveis:
            return cidades_possiveis[0].strip()
    
    # Se nenhum dos métodos acima funcionar, retornar não encontrado
    return "NÃO ENCONTRADO"
