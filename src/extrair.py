import xml.etree.ElementTree as ET
import openpyxl as xl

def extrair_xml2xlsx(arquivos, caminho_destino):
    try:
        for arquivo in arquivos:
            # Faz o parsing do arquivo XML
            tree = ET.parse(arquivo)
            root = tree.getroot()

            # Cria um novo arquivo Excel
            wb = xl.Workbook()
            ws = wb.active

            # Define o cabeçalho das colunas
            colunas = ['cProd', 'cEAN', 'xProd', 'NCM', 'CEST', 'qCom', 'vUnCom', 'vProd', 'cEANTrib', 'qTrib', 'vUnTrib']
            ws.append(colunas)

            # Obtém os dados do XML e popula o arquivo Excel
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            det_elements = root.findall('.//nfe:det', ns)

            for det in det_elements:
                dados = []
                prod = det.find('nfe:prod', ns)
                for coluna in colunas:
                    valor = prod.find(f'nfe:{coluna}', ns).text if prod.find(f'nfe:{coluna}', ns) is not None else ''
                    dados.append(valor)
                ws.append(dados)

            # Salva o arquivo Excel com um nome diferente no caminho de destino
            nome_arquivo = arquivo.split('/')[-1].replace('.xml', '.xlsx')
            caminho_arquivo = caminho_destino + '/' + nome_arquivo
            wb.save(caminho_arquivo)
            print(f"Arquivo '{caminho_arquivo}' criado com sucesso!")

    except Exception as e:
        print("Ocorreu um erro ao criar o arquivo:", str(e))

