#!/usr/bin/env python3
"""Processa a fila de ditados e mantém o DWR do dia completo.

Uso:
    python3 processar_fila.py [pasta_do_servico] [--vigiar]

Pasta do serviço (padrão: a pasta onde está este script):
    fila/      arquivos .txt com o texto ditado (um por acionamento)
    feito/     .txt já processados
    erros/     .txt que falharam (+ um .erro.txt com o motivo)
    exemplo/   base.docx = DWR completo do serviço (cabeçalho, tabelas e assinaturas)
    estado.json  (criado automaticamente) modo, número e registros do dia

Os DWRs gerados ficam na própria pasta do serviço: "DWR Nº 05.docx".

Comandos falados: "nova ação" e "nova atividade" só trocam o modo.
Qualquer outro texto vira um registro no modo atual.
Regras: a data é o dia do ditado (MM/DD/AAAA); a cada novo dia o número
sobe de 1 e os registros recomeçam. O primeiro DWR usa o número do base.docx.
"""
import json
import os
import re
import shutil
import sys
import time
import unicodedata
from datetime import datetime

import gerar_dwr

COMANDOS = {'nova acao': 'acoes', 'nova atividade': 'atividades'}


def normalizar(txt):
    txt = unicodedata.normalize('NFD', txt.lower())
    txt = ''.join(c for c in txt if unicodedata.category(c) != 'Mn')
    return ' '.join(re.sub(r'[^a-z0-9 ]', ' ', txt).split())


def numero_do_base(base):
    import zipfile
    with zipfile.ZipFile(base) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    tbl0 = gerar_dwr.TBL.findall(xml)[0]
    celulas = gerar_dwr.CELL.findall(tbl0)
    return int(gerar_dwr.texto(celulas[1]).strip())


def carregar_estado(pasta, base):
    caminho = os.path.join(pasta, 'estado.json')
    if os.path.exists(caminho):
        with open(caminho, encoding='utf-8') as f:
            return json.load(f)
    return {'numero': numero_do_base(base) - 1, 'data': None,
            'modo': 'atividades', 'atividades': [], 'acoes': []}


def salvar_estado(pasta, estado):
    caminho = os.path.join(pasta, 'estado.json')
    tmp = caminho + '.tmp'
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)
    os.replace(tmp, caminho)


def mover(origem, destino_dir, motivo=None):
    os.makedirs(destino_dir, exist_ok=True)
    nome = os.path.basename(origem)
    destino = os.path.join(destino_dir, nome)
    if os.path.exists(destino):
        base, ext = os.path.splitext(nome)
        destino = os.path.join(destino_dir, f'{base}_{int(time.time())}{ext}')
    shutil.move(origem, destino)
    if motivo:
        with open(destino + '.erro.txt', 'w', encoding='utf-8') as f:
            f.write(motivo + '\n')


def processar_arquivo(pasta, caminho, estado, base):
    """Devolve o novo estado (sem alterar o antigo) e grava o DWR do dia."""
    with open(caminho, encoding='utf-8-sig') as f:
        texto = f.read().strip()
    if not texto:
        raise gerar_dwr.ErroDWR('arquivo vazio')

    novo = json.loads(json.dumps(estado))          # cópia
    dia = datetime.fromtimestamp(os.path.getmtime(caminho)).strftime('%m/%d/%Y')
    if novo['data'] != dia:                        # novo dia
        novo.update(numero=novo['numero'] + 1, data=dia, modo='atividades',
                    atividades=[], acoes=[])

    comando = COMANDOS.get(normalizar(texto))
    if comando:
        novo['modo'] = comando
    else:
        novo[novo['modo']].append(texto)

    saida = os.path.join(pasta, f'DWR Nº {novo["numero"]:02d}.docx')
    tmp = saida + '.tmp'
    gerar_dwr.gerar(base, tmp, novo['numero'], novo['data'],
                    novo['atividades'], novo['acoes'])
    os.replace(tmp, saida)                         # PermissionError se estiver aberto no Word
    return novo


def processar_fila(pasta):
    base = os.path.join(pasta, 'exemplo', 'base.docx')
    if not os.path.exists(base):
        print(f'ERRO: não achei {base}')
        return 0
    fila = os.path.join(pasta, 'fila')
    os.makedirs(fila, exist_ok=True)
    arquivos = sorted((os.path.join(fila, n) for n in os.listdir(fila)
                       if n.lower().endswith('.txt')),
                      key=lambda p: (os.path.getmtime(p), p))
    feitos = 0
    estado = carregar_estado(pasta, base)
    for caminho in arquivos:
        try:
            estado = processar_arquivo(pasta, caminho, estado, base)
        except PermissionError:
            print('O DWR está aberto em outro programa; feche-o e rode de novo.')
            break                                  # deixa na fila para tentar depois
        except Exception as e:                     # qualquer outro problema vai para erros
            mover(caminho, os.path.join(pasta, 'erros'), f'{type(e).__name__}: {e}')
            print(f'erro em {os.path.basename(caminho)}: {e}')
            continue
        salvar_estado(pasta, estado)
        mover(caminho, os.path.join(pasta, 'feito'))
        feitos += 1
        print(f'ok: {os.path.basename(caminho)}')
    return feitos


if __name__ == '__main__':
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    pasta = os.path.abspath(args[0]) if args else os.path.dirname(os.path.abspath(__file__))
    if '--vigiar' in sys.argv:
        print(f'Vigiando {pasta}/fila (Ctrl+C para parar)')
        while True:
            processar_fila(pasta)
            time.sleep(10)
    else:
        processar_fila(pasta)
