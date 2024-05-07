import re
import getpass
import os
import pandas as pd
import win32com.client
import locale
from datetime import datetime
import xml.etree.ElementTree as ET

# Defina o locale para o Brasil (pt_BR)
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def trocar_descricao_kml(nome_arquivo_kml, nova_descricao):
    # Carregar o arquivo KML
    tree = ET.parse(nome_arquivo_kml)
    root = tree.getroot()

    # Definir namespaces que podem estar presentes no KML
    namespaces = {'kml': 'http://www.opengis.net/kml/2.2'}

    # Encontrar todos os elementos <Placemark> que contêm a descrição
    placemark_elementos = root.findall('.//kml:Placemark', namespaces)

    if placemark_elementos:
        for placemark_elemento in placemark_elementos:
            # Encontrar o elemento que contém a descrição dentro do <Placemark>
            descricao_elemento = placemark_elemento.find('.//kml:description', namespaces)

            # Se não houver um elemento de descrição, criar um novo e adicionar ao <Placemark>
            if descricao_elemento is None:
                # Criar um novo elemento de descrição
                descricao_elemento = ET.Element('{{{0}}}description'.format(namespaces['kml']))
                # Adicionar o texto da nova descrição ao elemento
                descricao_elemento.text = nova_descricao
                # Adicionar o novo elemento de descrição ao <Placemark>
                placemark_elemento.append(descricao_elemento)
            else:
                # Se o elemento de descrição existir, simplesmente atualizar seu texto
                descricao_elemento.text = nova_descricao
    else:
        print("Nenhum Placemark encontrado no arquivo KML.")

    # Salvar as mudanças de volta no arquivo KML
    tree.write(nome_arquivo_kml)

caminho_lnk = 'atalhos/CONTROLE DOS CONTRATOS - DOC I.lnk'
# Resolva o atalho para obter o caminho real do arquivo alvo
shell = win32com.client.Dispatch("WScript.Shell")
atalho = shell.CreateShortCut(caminho_lnk)
caminho_real = atalho.TargetPath

lista_de_abas = ['OBRAS_CONCLUÍDAS','LICITAÇÃO','OBRAS_PARALISADAS','SERVIÇOS_CONTÍNUOS','OBRAS_EM_EXECUÇÃO']

for nome in lista_de_abas:
    doc_planilha= pd.read_excel(caminho_real, sheet_name=nome)  
    doc_planilha.columns = doc_planilha.iloc[1]
    doc_planilha = doc_planilha.iloc[0:]
    doc_planilha = doc_planilha.dropna(subset=['PROCESSO'])
    doc_planilha = doc_planilha[2:]
    if nome == 'OBRAS_EM_EXECUÇÃO':
        doc_planilha_execucao = doc_planilha
    if nome == 'SERVIÇOS_CONTÍNUOS':
        doc_planilha_continuo = doc_planilha
    if nome =='OBRAS_PARALISADAS':
        doc_planilha_paralisada = doc_planilha
    if nome == 'LICITAÇÃO':
        doc_planilha_licitacao = doc_planilha
    if nome == 'OBRAS_CONCLUÍDAS':
        doc_planilha_concluida = doc_planilha


doc_planilha_execucao = doc_planilha_execucao.dropna(axis=1, how='all')
doc_planilha_execucao = doc_planilha_execucao.dropna(subset=['EMPRESA'])
doc_planilha_execucao['INÍCIO'] = pd.to_datetime(doc_planilha_execucao['INÍCIO'])
doc_planilha_execucao['TÉRMINO'] = pd.to_datetime(doc_planilha_execucao['TÉRMINO'])
doc_planilha_execucao['inicio_formatado'] = doc_planilha_execucao['INÍCIO'].dt.strftime('%d/%m/%Y')
doc_planilha_execucao['TÉRMINO_formatado'] = doc_planilha_execucao['TÉRMINO'].dt.strftime('%d/%m/%Y')
doc_planilha_execucao['Data base de Reajustamento'] = pd.to_datetime(doc_planilha_execucao['Data base de Reajustamento'])


doc_planilha_continuo['INÍCIO'] = pd.to_datetime(doc_planilha_continuo['INÍCIO'])
doc_planilha_continuo['TÉRMINO'] = pd.to_datetime(doc_planilha_continuo['TÉRMINO'])
doc_planilha_continuo['inicio_formatado'] = doc_planilha_continuo['INÍCIO'].dt.strftime('%d/%m/%Y')
doc_planilha_continuo['TÉRMINO_formatado'] = doc_planilha_continuo['TÉRMINO'].dt.strftime('%d/%m/%Y')
doc_planilha_continuo['ÚLTIMA MEDIÇÃO'] = pd.to_datetime(doc_planilha_continuo['ÚLTIMA MEDIÇÃO'])
doc_planilha_continuo['Data base de Reajustamento'] = pd.to_datetime(doc_planilha_continuo['Data base de Reajustamento'])

# print(doc_planilha_paralisada.columns)
doc_planilha_paralisada = doc_planilha_paralisada.dropna(axis=1, how='all')
doc_planilha_paralisada['INÍCIO'] = pd.to_datetime(doc_planilha_paralisada['INÍCIO'])
doc_planilha_paralisada['TÉRMINO'] = pd.to_datetime(doc_planilha_paralisada['TÉRMINO'])
doc_planilha_paralisada = doc_planilha_paralisada.dropna(axis=1)
# print(doc_planilha_paralisada.columns)
doc_planilha_paralisada['inicio_formatado'] = doc_planilha_paralisada['INÍCIO'].dt.strftime('%d/%m/%Y')
doc_planilha_paralisada['TÉRMINO_formatado'] = doc_planilha_paralisada['TÉRMINO'].dt.strftime('%d/%m/%Y')


doc_planilha_licitacao = doc_planilha_licitacao.dropna(axis=1, how='all')

doc_planilha_concluida = doc_planilha_concluida[doc_planilha_concluida['PROCESSO'] != 'PROCESSO']
doc_planilha_concluida = doc_planilha_concluida.dropna(subset = ['NATUREZA'])
doc_planilha_concluida['INÍCIO'] = pd.to_datetime(doc_planilha_concluida['INÍCIO'])
doc_planilha_concluida['TÉRMINO'] = pd.to_datetime(doc_planilha_concluida['TÉRMINO'])
doc_planilha_concluida['inicio_formatado'] = doc_planilha_concluida['INÍCIO'].dt.strftime('%d/%m/%Y')
doc_planilha_concluida['TÉRMINO_formatado'] = doc_planilha_concluida['TÉRMINO'].dt.strftime('%d/%m/%Y')
doc_planilha_concluida['ÚLTIMA MEDIÇÃO'] = pd.to_datetime(doc_planilha_concluida['ÚLTIMA MEDIÇÃO'])

# print(caminho_real)

lista_dataframe = [doc_planilha_execucao,doc_planilha_licitacao,doc_planilha_paralisada,doc_planilha_continuo,doc_planilha_concluida]
todos_processos = []
for dataframe in lista_dataframe:
    lista_de_processos = dataframe["PROCESSO"].tolist()
    for processo in lista_de_processos:
        processo_formatado = processo.replace('.', '-')
        processo_formatado = processo_formatado.replace('/', '-')
        processo_formatado = processo_formatado.replace(' ', '')  # Remover espaços
        todos_processos.append(processo_formatado)

pasta_kmz = "KML DOC's/KML DOC I"  

# Listar todos os arquivos KMZ em todas as subpastas
nomes_kml = []
for pasta_raiz, _, arquivos in os.walk(pasta_kmz):
    for arquivo in arquivos:
        if arquivo.endswith('.kml'):
            nomes_kml.append(os.path.join(pasta_raiz, arquivo))  # Adiciona o caminho completo do arquivo

relacionamento = {}

for nome_arquivo in nomes_kml:
    nome_arquivo_sem_extensao = os.path.splitext(os.path.basename(nome_arquivo))[0]
    for processo in todos_processos:  
        if processo in nome_arquivo_sem_extensao:
            relacionamento[processo] = nome_arquivo

print(relacionamento)

processos_fora = []
for processo0 in todos_processos:
    soma = 0
    for chave, valor in relacionamento.items():
        if chave == processo0:
            soma+=1
    if soma ==0:
        processos_fora.append(processo0)
print(processos_fora)
            
estilo_tabela = """
body {
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    background-color: #f7f7f7;
    margin: 0;
    padding: 0;
}

table {
    width: 100%;
    border-collapse: collapse;
    border-spacing: 0;
    margin-top: 20px;
    background: linear-gradient(to bottom right, #ffffff, #f2f2f2);
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
}

th, td {
    border: 1px solid #ddd;
    padding: 12px;
    text-align: left;
}

th {
    background-color: #4CAF50;
    color: white;
    font-weight: bold;
    text-transform: uppercase;
}

tr:nth-child(even) {
    background-color: #e7e6e7;
}

tr:hover {
    background-color: #f2f2f2;
}

a {
    color: #007bff;
    text-decoration: none;
    transition: color 0.3s;
}

a:hover {
    color: #0056b3;
}

.inner-table {
    width: 100%;
    border-collapse: collapse;
    border-spacing: 0;
}

.inner-table td {
    padding: 10px;
    border: none;
}

.inner-table tr:last-child td {
    border-bottom: 1px solid #ddd;
}

.inner-table td:first-child {
    font-weight: bold;
    color: #555;
}

.inner-table td:last-child {
    text-align: right;
}

"""

def html_exe_conti_para(processo, objeto_processo, contratada, investimento_formatado, contrato, inicio, previsao_termino, roc, fiscais, gestor, siafe, link, status):  
    global estilo_tabela
    html = f"""
<html xmlns:fo="http://www.w3.org/1999/XSL/Format" xmlns:msxsl="urn:schemas-microsoft-com:xslt">
<head>
    <meta charset="UTF-8">
    <title>Tabela de Contratos</title>
    <style>
        {estilo_tabela}
    </style>
</head>
<body>
    <table>
        <tr>
            <th>{processo}</th>
        </tr>
        <tr>
            <td>
                <table>
                    <tr>
                        <td>Objeto</td>
                        <td>{objeto_processo}</td>
                    </tr>
                    <tr>
                        <td>Contratada</td>
                        <td>{contratada}</td>
                    </tr>
                    <tr>
                        <td>Investimento:</td>
                        <td>{investimento_formatado}</td>
                    </tr>
                    <tr>
                        <td>Contrato</td>
                        <td>{contrato}</td>
                    </tr>
                    <tr>
                        <td>Início</td>
                        <td>{inicio}</td>
                    </tr>
                    <tr>
                        <td>Previsão de Término</td>
                        <td>{previsao_termino}</td>
                    </tr>
                    <tr>
                        <td>Residência Responsável</td>
                        <td>{roc}</td>
                    </tr>
                    <tr>
                        <td>Fiscais</td>
                        <td>{fiscais}</td>
                    </tr>
                    <tr>
                        <td>Gestor(a)</td>
                        <td>{gestor}</td>
                    </tr>
                    <tr>
                        <td>Siafe</td>
                        <td>{siafe}</td>
                    </tr>
                    <tr>
                        <td>Status</td>
                        <td><a href="{link}">PACTO RJ - {status}</a></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"""
    return html

def html_licitacao(processo,objeto_processo,roc,gestor,link,status):
    global estilo_tabela
    html = f"""
<html>
<head>
    <meta charset="UTF-8">
    <title>Detalhes da Licitação</title>
    <style>
        {estilo_tabela}
    </style>
</head>
<body>
    <table>
        <tr>
            <th>{processo}</th>
        </tr>
        <tr>
            <td>
                <table>
                    <tr>
                        <td>Objeto</td>
                        <td>{objeto_processo}</td>
                    </tr>
                    <tr>
                        <td>Residência Responsável</td>
                        <td>{roc}</td>
                    </tr>
                    <tr>
                        <td>Gestor(a)</td>
                        <td>{gestor}</td>
                    </tr>
                    <tr>
                        <td>Status</td>
                        <td><a href="{link}">PACTO RJ - {status}</a></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"""
    return html

def html_concluida(processo, objeto_processo, contratada, investimento_formatado, contrato, inicio, previsao_termino, roc, fiscais, gestor, link, status, siafe):
    global estilo_tabela
    html = f"""
<html>
<head>
    <meta charset="UTF-8">
    <title>Detalhes do Contrato Concluído</title>
    <style>
        {estilo_tabela}
    </style>
</head>
<body>
    <table>
        <tr>
            <th>{processo}</th>
        </tr>
        <tr>
            <td>
                <table>
                    <tr>
                        <td>Objeto</td>
                        <td>{objeto_processo}</td>
                    </tr>
                    <tr>
                        <td>Contratada</td>
                        <td>{contratada}</td>
                    </tr>
                    <tr>
                        <td>Investimento:</td>
                        <td>{investimento_formatado}</td>
                    </tr>
                    <tr>
                        <td>Contrato</td>
                        <td>{contrato}</td>
                    </tr>
                    <tr>
                        <td>Início</td>
                        <td>{inicio}</td>
                    </tr>
                    <tr>
                        <td>Término</td>
                        <td>{previsao_termino}</td>
                    </tr>
                    <tr>
                        <td>Residência Responsável</td>
                        <td>{roc}</td>
                    </tr>
                    <tr>
                        <td>Fiscais</td>
                        <td>{fiscais}</td>
                    </tr>
                    <tr>
                        <td>Gestor(a)</td>
                        <td>{gestor}</td>
                    </tr>
                    <tr>
                        <td>Siafe</td>
                        <td>{siafe}</td>
                    </tr>
                    <tr>
                        <td>Status</td>
                        <td><a href="{link}">PACTO RJ - {status}</a></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"""
    return html



all_processos = []
dict_html = []
for dataframe in lista_dataframe:

    lista_de_processos = dataframe["PROCESSO"].tolist()

    if dataframe.equals(doc_planilha_execucao) or dataframe.equals(doc_planilha_continuo) or dataframe.equals(doc_planilha_paralisada):
        if dataframe.equals(doc_planilha_execucao):
            status = 'Em execução'
        if dataframe.equals(doc_planilha_continuo):
            status = 'Serviço Contínuo'
        if dataframe.equals(doc_planilha_paralisada):
            status = 'Obra paralisada'
        for processo in lista_de_processos:
            objeto_processo = dataframe.loc[dataframe["PROCESSO"] == processo, "DESCRIÇÃO"].values[0]
            contratada = dataframe.loc[dataframe["PROCESSO"] == processo, "EMPRESA"].values[0]
            investimento = dataframe.loc[dataframe["PROCESSO"] == processo, "VALOR CONTRATO"].values[0]
            investimento_formatado = locale.currency(investimento, grouping=True, symbol=True)
            roc = dataframe.loc[dataframe["PROCESSO"] == processo, "ROC"].values[0] + ' ROC'
            inicio = dataframe.loc[dataframe["PROCESSO"] == processo, "inicio_formatado"].values[0]
            previsao_termino = dataframe.loc[dataframe["PROCESSO"] == processo, "TÉRMINO_formatado"].values[0]
            fiscal1 = dataframe.loc[dataframe["PROCESSO"] == processo, "FISCAL"].values[0][0]
            fiscal2 = dataframe.loc[dataframe["PROCESSO"] == processo, "FISCAL"].values[0][1]
            contrato = dataframe.loc[dataframe["PROCESSO"] == processo, "CONTRATO"].values[0]
            #CNPJ = doc_planilha_execucao.loc[doc_planilha_execucao["PROCESSO"] == processo, " "].values[0]
            gestor = "GLEICE D' LURDES GONÇALVES DE AMORIM"
            fiscais = f"{fiscal1} , {fiscal2}"
            link = ' aaaaaa'

            #ADICIONANDO NÚMERO AUTOMÁTICO
            siafe = dataframe.loc[dataframe["PROCESSO"] == processo, "número automatico siafe"].values[0]

            html_template0 = html_exe_conti_para(processo,objeto_processo,contratada,investimento_formatado,contrato,inicio,previsao_termino,roc,fiscais,gestor,siafe,link,status)    
            
            processo_formatado = processo.replace('.', '-')
            processo_formatado = processo_formatado.replace('/', '-')
            all_processos.append(processo_formatado)

            
            if dataframe.equals(doc_planilha_execucao):
                status = 'Em execução'
                with open(f"htmls/Obras_em_execução/{processo_formatado}.html", "w", encoding='utf-8') as html_file:
                    html_file.write(html_template0)

            if dataframe.equals(doc_planilha_continuo):
                status = 'Serviço Contínuo'
                with open(f"htmls/Serviços_contínuos/{processo_formatado}.html", "w", encoding='utf-8') as html_file:
                    html_file.write(html_template0)

            if dataframe.equals(doc_planilha_paralisada):
                status = 'Obra paralisada'
                with open(f"htmls/paralisadas/{processo_formatado}.html", "w", encoding='utf-8') as html_file:
                    html_file.write(html_template0)

            print(processo_formatado," ", status, ' ', processo)########################################################
            if relacionamento.get(processo_formatado) is not None:
                print('vim até aqui')
                trocar_descricao_kml(relacionamento.get(processo_formatado), html_template0)

        

    if dataframe.equals(doc_planilha_licitacao):
        lista_de_processos = dataframe["PROCESSO"].tolist()
        for processo in lista_de_processos:
            objeto_processo = dataframe.loc[dataframe["PROCESSO"] == processo, "DESCRIÇÃO"].values[0]
            roc = dataframe.loc[dataframe["PROCESSO"] == processo, "ROC"].values[0] + ' ROC'
            investimento = dataframe.loc[dataframe["PROCESSO"] == processo, "VALOR CONTRATO"].values[0]
            status = 'Licitação'

            # print(processo)
            # print(objeto_processo)
            html_template0 = html_licitacao(processo,objeto_processo,roc,gestor,link,status)
            
            processo_formatado = processo.replace('.', '-')
            processo_formatado = processo_formatado.replace('/', '-')
            all_processos.append(processo_formatado)

            # with open(f"htmls/Licitações/{processo_formatado}.html", "w", encoding='utf-8') as html_file:
            #     html_file.write(html_template0)
            
            print(processo_formatado," ", status, ' ', processo)##########################################################################
            if relacionamento.get(processo_formatado) is not None:
                print('vim até aqui')
                trocar_descricao_kml(relacionamento.get(processo_formatado), html_template0)
        
    if dataframe.equals(doc_planilha_concluida):
        lista_de_processos = dataframe["PROCESSO"].tolist()
        for processo in lista_de_processos:
            objeto_processo = dataframe.loc[dataframe["PROCESSO"] == processo, "DESCRIÇÃO"].values[0]
            investimento = dataframe.loc[dataframe["PROCESSO"] == processo, "VALOR CONTRATO"].values[0]
            if not dataframe.empty and processo in dataframe["PROCESSO"].values:
                objeto_processo = dataframe.loc[dataframe["PROCESSO"] == processo, "DESCRIÇÃO"].values[0]
                contratada = dataframe.loc[dataframe["PROCESSO"] == processo, "EMPRESA"].values[0]
                investimento = dataframe.loc[dataframe["PROCESSO"] == processo, "VALOR CONTRATO"].values[0]
            else:
                # Tratar o caso em que não há dados correspondentes ao processo especificado
                objeto_processo = None
                contratada = None
                investimento = None
            # objeto_processo = dataframe.loc[dataframe["PROCESSO"] == processo, "DESCRIÇÃO"].values[0]
            # contratada = dataframe.loc[dataframe["PROCESSO"] == processo, "EMPRESA"].values[0]
            # investimento = dataframe.loc[dataframe["PROCESSO"] == processo, "VALOR CONTRATO"].values[0]
            # aceite_provi = dataframe.loc[dataframe["PROCESSO"] == processo, "ACEITE PROVISÓRIO"].values[0]
            # aceite_defi = dataframe.loc[dataframe["PROCESSO"] == processo, "ACEITE DEFINITIVO"].values[0]
            investimento_formatado = locale.currency(investimento, grouping=True, symbol=True)
            roc_value = dataframe.loc[dataframe["PROCESSO"] == processo, "ROC"].values[0]
            roc = str(roc_value) + ' ROC' if not pd.isna(roc_value) else ''  # Convertendo o valor numérico em uma string
            inicio = dataframe.loc[dataframe["PROCESSO"] == processo, "inicio_formatado"].values[0]
            previsao_termino = dataframe.loc[dataframe["PROCESSO"] == processo, "TÉRMINO_formatado"].values[0]
            fiscal1 = dataframe.loc[dataframe["PROCESSO"] == processo, "FISCAL"].values[0][0]
            fiscal2 = dataframe.loc[dataframe["PROCESSO"] == processo, "FISCAL"].values[0][1]
            contrato = dataframe.loc[dataframe["PROCESSO"] == processo, "CONTRATO"].values[0]
            gestor = "GLEICE D' LURDES GONÇALVES DE AMORIM"
            fiscais = f"{fiscal1} , {fiscal2}"
            status = 'Obra Concluída'
            siafe = dataframe.loc[dataframe["PROCESSO"] == processo, "número automatico siafe"].values[0]

            html_template0 = html_concluida(processo,objeto_processo,contratada,investimento_formatado,contrato,inicio,previsao_termino,roc,fiscais,gestor,link,status,siafe)
            
            processo_formatado = processo.replace('.', '-')
            processo_formatado = processo_formatado.replace('/', '-')
            all_processos.append(processo_formatado)
            
            with open(f"htmls/Concluídas/{processo_formatado}.html", "w", encoding='utf-8') as html_file:
                html_file.write(html_template0)

            print(processo_formatado," ", status, ' ', processo)########################################################
            if relacionamento.get(processo_formatado) is not None:
                print('vim até aqui')
                print(processo_formatado+" "+relacionamento.get(processo_formatado))
                trocar_descricao_kml(relacionamento.get(processo_formatado), html_template0)



