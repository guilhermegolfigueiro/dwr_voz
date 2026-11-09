#!/usr/bin/env python3
"""Gera um DWR completo a partir do documento base.

Uso:  python3 gerar_dwr.py base.docx saida.docx dados.json

dados.json:
{
  "numero": 5,
  "data": "08/16/2026",
  "atividades": ["texto 1", "texto 2"],
  "acoes": ["texto 1"]
}

Só usa a biblioteca padrão (zipfile, re), então roda no a-Shell sem instalar nada.
O documento base nunca é alterado: o resultado é gravado em outro arquivo.
"""
import json
import re
import sys
import zipfile
from xml.sax.saxutils import escape

RPR = ('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:bCs/>'
       '<w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US"/></w:rPr>')

TBL = re.compile(r'<w:tbl>.*?</w:tbl>', re.S)
ROW = re.compile(r'<w:tr[ >].*?</w:tr>', re.S)
CELL = re.compile(r'<w:tc>.*?</w:tc>', re.S)
TEXT = re.compile(r'<w:t(?: [^>]*)?>([^<]*)</w:t>')


class ErroDWR(Exception):
    pass


def texto(xml):
    return ''.join(TEXT.findall(xml))


def novo_paragrafo(txt, centralizado):
    ppr = '<w:pPr>' + ('<w:jc w:val="center"/>' if centralizado else '') + RPR + '</w:pPr>'
    run = (f'<w:r>{RPR}<w:t xml:space="preserve">{escape(txt)}</w:t></w:r>') if txt else ''
    return f'<w:p>{ppr}{run}</w:p>'


def nova_linha(item, desc):
    return ('<w:tr><w:tc><w:tcPr><w:tcW w:w="912" w:type="dxa"/></w:tcPr>'
            + novo_paragrafo(item, True) + '</w:tc>'
            '<w:tc><w:tcPr><w:tcW w:w="8727" w:type="dxa"/></w:tcPr>'
            + novo_paragrafo(desc, False) + '</w:tc></w:tr>')


def reconstruir_tabela(tbl, registros):
    """Mantém o cabeçalho da tabela e recria as linhas de dados:
    uma linha por registro, numeradas 01, 02..."""
    linhas = ROW.findall(tbl)
    if len(linhas) < 2:
        raise ErroDWR('tabela sem linhas suficientes')
    inicio = tbl.index(linhas[0])
    if tbl != tbl[:inicio] + ''.join(linhas) + '</w:tbl>':
        raise ErroDWR('estrutura inesperada dentro da tabela')
    cabecalho = linhas[0]
    if 'Item' not in texto(cabecalho) or 'Description' not in texto(cabecalho):
        raise ErroDWR('cabeçalho da tabela não é Item / Description')

    # Linhas de nota do base (ex.: '*Brazilian time') não são mantidas.
    corpo = [nova_linha(f'{i:02d}', t.strip()) for i, t in enumerate(registros, 1) if t.strip()]
    if not corpo:
        corpo = [nova_linha('', '')]          # linha padrão obrigatória
    return tbl[:inicio] + cabecalho + ''.join(corpo) + '</w:tbl>'


def trocar_texto_celula(celula, novo):
    """Põe `novo` no primeiro <w:t> da célula e esvazia os demais
    (o Word costuma quebrar '08/12/2026' em vários trechos)."""
    achados = list(TEXT.finditer(celula))
    if not achados:
        raise ErroDWR('célula sem texto para atualizar')
    partes, pos = [], 0
    for k, m in enumerate(achados):
        partes.append(celula[pos:m.start()])
        partes.append(f'<w:t xml:space="preserve">{escape(novo) if k == 0 else ""}</w:t>')
        pos = m.end()
    partes.append(celula[pos:])
    return ''.join(partes)


def atualizar_numero_data(tbl0, numero, data):
    celulas = CELL.findall(tbl0)
    if len(celulas) < 5:
        raise ErroDWR('tabela do número/data diferente do esperado')
    if not re.fullmatch(r'\s*\d+\s*', texto(celulas[1])):
        raise ErroDWR('não achei o número do DWR no documento base')
    if not re.fullmatch(r'\s*\d\d/\d\d/\d{4}\s*', texto(celulas[4])):
        raise ErroDWR('não achei a data no documento base')
    novas = list(celulas)
    novas[1] = trocar_texto_celula(celulas[1], f'{int(numero):02d}')
    novas[4] = trocar_texto_celula(celulas[4], data)
    saida, pos = [], 0
    for antiga, nova in zip(celulas, novas):
        i = tbl0.index(antiga, pos)
        saida.append(tbl0[pos:i])
        saida.append(nova)
        pos = i + len(antiga)
    saida.append(tbl0[pos:])
    return ''.join(saida)


def processar_documento(xml, numero, data, atividades, acoes):
    tabelas = TBL.findall(xml)
    if len(tabelas) != xml.count('</w:tbl>') or any('<w:tbl>' in t[8:] for t in tabelas):
        raise ErroDWR('o documento base tem tabelas aninhadas, não suportado')
    if len(tabelas) < 6:
        raise ErroDWR('o documento base não tem as 6 tabelas esperadas')
    plano = texto(xml)
    for titulo in ('LIST OF ACTIVITIES PERFORMED', 'ACTIONS / REQUESTS'):
        if titulo not in plano:
            raise ErroDWR(f'título "{titulo}" não encontrado no documento base')

    if 'DWR' not in texto(tabelas[0]):
        raise ErroDWR('a primeira tabela não é a do número/data')
    # tabelas de dados: as duas com cabeçalho Item / Description, na ordem
    dados = [t for t in tabelas if ROW.findall(t) and
             texto(ROW.findall(t)[0]).replace(' ', '') == 'ItemDescription']
    if len(dados) != 2:
        raise ErroDWR('não achei as duas tabelas Item / Description')

    substituicoes = [
        (tabelas[0], atualizar_numero_data(tabelas[0], numero, data)),
        (dados[0], reconstruir_tabela(dados[0], atividades)),
        (dados[1], reconstruir_tabela(dados[1], acoes)),
    ]
    for antiga, nova in substituicoes:
        if xml.count(antiga) != 1:
            raise ErroDWR('trecho do documento aparece mais de uma vez')
        xml = xml.replace(antiga, nova)
    return xml


def gerar(base, saida, numero, data, atividades, acoes):
    with zipfile.ZipFile(base) as zin:
        xml = zin.read('word/document.xml').decode('utf-8')
        novo = processar_documento(xml, numero, data, atividades, acoes)
        with zipfile.ZipFile(saida, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.is_dir():
                    continue
                conteudo = zin.read(item.filename)
                if item.filename == 'word/document.xml':
                    conteudo = novo.encode('utf-8')
                zout.writestr(item, conteudo)


if __name__ == '__main__':
    if len(sys.argv) != 4:
        sys.exit(__doc__)
    base, saida, arq = sys.argv[1:4]
    with open(arq, encoding='utf-8') as f:
        d = json.load(f)
    try:
        gerar(base, saida, d['numero'], d['data'], d.get('atividades', []), d.get('acoes', []))
    except ErroDWR as e:
        sys.exit(f'ERRO: {e}')
    print('ok:', saida)
