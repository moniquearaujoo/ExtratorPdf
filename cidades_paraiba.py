import unicodedata
import re

class CidadesParaiba:
    cidades = [
        "AGUIAR", "ALAGOA GRANDE", "ALAGOA NOVA", "ALAGOINHA", "ALHANDRA", "AMPARO", "APARECIDA", "ARAÇAGI",
        "ARARA", "ARARUNA", "AREIA", "AREIA DE BARAÚNAS", "AREIAL", "AROEIRAS", "ASSUNÇÃO", "BAÍA DA TRAIÇÃO",
        "BANANEIRAS", "BARAÚNA", "BARRA DE SANTA ROSA", "BARRA DE SANTANA", "BARRA DE SÃO MIGUEL",
        "BAYEUX", "BELÉM", "BELÉM DO BREJO DO CRUZ", "BERNARDINO BATISTA", "BOA VENTURA", "BOA VISTA",
        "BOM JESUS", "BOM SUCESSO", "BONITO DE SANTA FÉ", "BOQUEIRÃO", "BORBOREMA", "BREJO DO CRUZ",
        "BREJO DOS SANTOS", "CAAPORÃ", "CABACEIRAS", "CABEDELO", "CACHOEIRA DOS ÍNDIOS", "CACIMBA DE AREIA",
        "CACIMBA DE DENTRO", "CACIMBAS", "CAIÇARA", "CAJAZEIRAS", "CAJAZEIRINHAS", "CALDAS BRANDÃO",
        "CAMALAÚ", "CAMPINA GRANDE", "CARRAPATEIRA", "CATINGUEIRA", "CATINGUEIRA", "CATOLÉ DO ROCHA",
        "CEARÁ-MIRIM", "CONDADO", "CONDE", "CONGO", "COREMAS", "COXIXOLA", "CRUZ DO ESPÍRITO SANTO",
        "CUBATI", "CUITE", "CUITEGI", "CUITE DE MAMANGUAPE", "CURRAL DE CIMA", "CURRAL VELHO", "DAMIAO",
        "DESTERRO", "DIAMANTE", "DONA INÉS", "DUAS ESTRADAS", "EMAS", "ESPERANÇA", "FAGUNDES",
        "FREI MARTINHO", "GADO BRAVO", "GUARABIRA", "GURINHÉM", "GURJÃO", "IBIARA", "IGARACY",
        "IMACULADA", "INGA", "ITABAIANA", "ITAPORANGA", "ITAPOROROCA", "ITATUBA", "JACARAÚ", "JERICÓ",
        "JOÃO PESSOA", "JUAREZ TAVORA", "JUAZEIRINHO", "JUNCO DO SERIDO", "JURU", "LAGOA",
        "LAGOA DE DENTRO", "LAGOA SECA", "LASTRO", "LIVRAMENTO", "LOGRADOURO", "LUCENA", "MÃE D'ÁGUA",
        "MALTA", "MAMANGUAPE", "MANAÍRA", "MARCELÍA", "MARI", "MARIZÓPOLIS", "MASSARANDUBA", "MATARACA",
        "MATINHAS", "MATO GROSSO", "MATURÉIA", "MOGEIRO", "MONTADAS", "MONTE HOREBE", "MONTEIRO",
        "MULUNGU", "NATUBA", "NAZARÉ", "NOVA FLORESTA", "NOVA OLINDA", "NOVA PALMEIRA", "OLHO D'ÁGUA",
        "OLIVEDOS", "OURO VELHO", "PARARI", "PASSAGEM", "PATOS", "PAULISTA", "PEDRA BRANCA", "PEDRA LAVRADA",
        "PEDRAS DE FOGO", "PIANCÓ", "PICUÍ", "PILAR", "PILÕEZINHOS", "PIRPIRITUBA", "PITIMBU", "POCINHOS",
        "POLESTINA", "POMBAL", "PRATA", "PRINCESA ISABEL", "PUXINANÃ", "QUEIMADAS", "QUIXABA", "REMÍGIO", "RIACHÃO",
        "RIACHÃO DO BACAMARTE", "RIACHÃO DO POÇO", "RIACHO DE SANTO ANTÔNIO", "RIACHO DOS CAVALOS",
        "SALGADINHO", "SALGADO DE SÃO FÉLIX", "SANTA CECÍLIA", "SANTA CRUZ", "SANTA HELENA", "SANTA INÊS",
        "SANTA LUZIA", "SANTA RITA", "SANTA TERESINHA", "SANTANA DE MANGUEIRA", "SANTANA DOS GARROTES",
        "SANTO ANDRÉ", "SÃO BENTO", "SÃO BENTINHO", "SÃO DOMINGOS", "SÃO DOMINGOS DO CARIRI",
        "SÃO FRANCISCO", "SÃO JOÃO DO CARIRI", "SÃO JOÃO DO RIO DO PEIXE", "SÃO JOÃO DO TIGRE",
        "SÃO JOSÉ DA LAGOA TAPADA", "SÃO JOSÉ DE CAIANA", "SÃO JOSÉ DE ESPINHARAS",
        "SÃO JOSÉ DE PIRANHAS", "SÃO JOSÉ DE PRINCESA", "SÃO JOSÉ DO BONFIM", "SÃO JOSÉ DO BREJO DO CRUZ",
        "SÃO JOSÉ DO SABUGI", "SÃO JOSÉ DOS CORDEIROS", "SÃO JOSÉ DOS RAMOS", "SÃO MAMEDE",
        "SÃO MIGUEL DE TAIPU", "SÃO SEBASTIÃO DE LAGOA DE ROÇA", "SÃO SEBASTIÃO DO UMBUZEIRO",
        "SAPÉ", "SERIDÓ", "SERTÃOZINHO", "SOBRADO", "SOLÂNEA", "SOLEDADE", "SOUSA", "SUMÉ",
        "TAPEROÁ", "TAVARES", "TEIXEIRA", "TENÓRIO", "TRIUNFO", "UIRAÚna", "UMARI", "UMBURETAMA",
        "VÁRZEA", "VIEIRÓPOLIS", "VISTA SERRANA"
    ]

    @classmethod
    def remover_acentos(cls, texto):
        return ''.join(
            c for c in unicodedata.normalize('NFD', texto)
            if unicodedata.category(c) != 'Mn'
        )

    @classmethod
    def validar_municipio(cls, nome_extraido):
        nome_extraido = cls.remover_acentos(nome_extraido.upper().strip())

        # Corrige caso esteja grudado com o CEP, ex: "POMBAL - PB58840-000"
        nome_extraido = re.sub(r'(\s[A-Z]{2})(\d)', r'\1 \2', nome_extraido)

        for cidade in cls.cidades:
            cidade_sem_acentos = cls.remover_acentos(cidade.upper())
            if cidade_sem_acentos in nome_extraido:
                return cidade  # Retorna o nome com acento da lista original
        return "NÃO ENCONTRADO"
