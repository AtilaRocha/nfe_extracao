from extrair import extrair_xml2xlsx

# Chama a função com uma lista de arquivos XML fornecidos e o caminho de destino
arquivos = [
    'G:/Atila_Rocha/Programacao/asdasdasdasd/nfe_extracao/src/base/arquivo1.xml',
    'G:/Atila_Rocha/Programacao/asdasdasdasd/nfe_extracao/src/base/arquivo2.xml',
    ]
caminho_destino = 'G:/Atila_Rocha/Programacao/asdasdasdasd/nfe_extracao/db'

extrair_xml2xlsx(arquivos, caminho_destino)