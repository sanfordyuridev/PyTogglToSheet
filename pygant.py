import os
import gspread
import requests
from base64 import b64encode
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

NomeParaAtualizar = os.environ.get('NOME')

def conectarPlanilhaEspelho(nomeArquivo, codigoPlanilha, nomeFolha):
    gc = gspread.service_account(filename=nomeArquivo)
    sh = gc.open_by_key(codigoPlanilha)

    return sh.worksheet(nomeFolha)

ws = conectarPlanilhaEspelho(os.environ.get('NOME_ARQUIVO'), os.environ.get('CODIGO_PLANILHA'), os.environ.get('NOME_FOLHA'))

def converteSegundosParaHoras(duracaoEmSegundos):
    horaEmSegundos = 3600

    return round(float(duracaoEmSegundos / horaEmSegundos))

def autenticar(email, password):
    url = 'https://api.track.toggl.com/api/v9/me'
    response = requests.get(url, auth=(email, password))
    if response.status_code == 200:
        return True
    else:
        return False

def obterTarefas(email, password, data_inicio, data_fim):
    url = f"https://api.track.toggl.com/api/v9/me/time_entries?start_date={data_inicio}&end_date={data_fim}"
    headers = {
        'Content-Type': 'application/json',
        'Authorization': 'Basic %s' % b64encode(f"{email}:{password}".encode("ascii")).decode("ascii")
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()
    else:
        return None

def formatarData(data):
    date = datetime.strptime(data, '%Y-%m-%dT%H:%M:%S%z')
    return date.strftime('%d/%m')

def salvarNaPlanilha(planilha, tarefa, posicao):
    pos_ref_dias_inicio = 'D1'
    pos_ref_dias_fim = 'Z1'

    planilha.update('A' + str(posicao), tarefa['tag'])
    planilha.update('B' + str(posicao), tarefa['descricao'])
    planilha.update('C' + str(posicao), tarefa['duracao'])

    dias = tarefa['dias'].split()

    cell_range = planilha.range(pos_ref_dias_inicio + ':' + pos_ref_dias_fim)
    valores = [celula.value for celula in cell_range]

    for dia in dias:
        indice_coluna = next(indice for indice, valor in enumerate(valores) if valor == dia)
        letra_coluna = chr(ord(pos_ref_dias_inicio[0]) + indice_coluna)

        posicao_celula = letra_coluna + str(posicao)

        planilha.update(posicao_celula, NomeParaAtualizar)

        planilha.format(posicao_celula, {'backgroundColor': {'red': 0.949, 'green': 0.949, 'blue': 0.949}})


email = os.environ.get('EMAIL')
password = os.environ.get('SENHA')

autenticado = autenticar(email, password)

if autenticado:
    data_inicio = input("Digite a data de início (no formato YYYY-MM-DD): ")
    data_fim = input("Digite a data de fim (no formato YYYY-MM-DD): ")

    tarefas = obterTarefas(email, password, data_inicio, data_fim)

    if tarefas is not None:
        timesEntries = []

        for entrada in tarefas:
            tags = entrada.get('tags')
            descricao = entrada.get('description')
            duracao = entrada.get('duration', 0)

            dias = entrada.get('start')

            data_formatada = formatarData(dias)

            timeentry = {
                'tag': tags[0],
                'descricao': descricao,
                'duracao': duracao,
                'dias': data_formatada
            }

            for entrada_existente in timesEntries:
                if entrada_existente['descricao'] == timeentry['descricao']:
                    entrada_existente['duracao'] += timeentry['duracao']

                    if not (timeentry['dias'] in entrada_existente['dias']):
                        entrada_existente['dias'] += ' ' + timeentry['dias']
                    break
            else:
                timesEntries.append(timeentry)

        posicao = 3
        duracaoTotal = 0

        for te in timesEntries:
            duracao_em_horas = converteSegundosParaHoras(int(te['duracao'])) 
            duracaoTotal += duracao_em_horas 
            if duracao_em_horas > 0:
                te['duracao'] = duracao_em_horas
                salvarNaPlanilha(ws, te, posicao)
                posicao += 1

        print(' ')
        print(f'Foram salvas com sucesso {len(timesEntries)} tarefas')
        print(' ')
        print(f'Total de horas trabalhadas: {duracaoTotal} horas')
        print(f'Média de horas trabalhadas por demanda: {duracaoTotal / len(timesEntries)} horas')
        print(' ')
    else:
        print('Erro ao obter as tarefas.')
else:
    print('Erro na autenticação.')
