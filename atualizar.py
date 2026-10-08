#!/usr/bin/env python3
"""
atualizar.py
Lê dados.xlsx e gera index.html atualizado.
Uso: python atualizar.py
Requisitos: pip install openpyxl
"""

import openpyxl, json, re, sys, os
from collections import defaultdict

# Garante que o script trabalha sempre na sua própria pasta
os.chdir(os.path.dirname(os.path.abspath(__file__)))

XLSX = 'dados.xlsx'
TEMPLATE = 'index.html'
SAIDA = 'index.html'

# ── Lê a planilha ──────────────────────────────────────────────────────────
def ler_planilha(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb['Aproveitamentos']
    rows = list(ws.iter_rows(values_only=True))
    # Linha 1 = cabeçalho
    dados = []
    for r in rows[1:]:
        if not r[0]:  # linha vazia
            continue
        dados.append({
            'comp_ufrn':      str(r[0]).strip(),
            'nivel_vesp':     int(r[1]) if r[1] else None,
            'nivel_not':      int(r[2]) if r[2] else None,
            'instituicao':    str(r[3]).strip() if r[3] else '',
            'id_opcao':       int(r[4]) if r[4] else 0,
            'num_opcao':      int(r[5]) if r[5] else 1,
            'componentes_str':str(r[6]).strip() if r[6] else '',
            'requer_conjunto':str(r[7]).strip().lower() == 'sim' if r[7] else False,
            'resultado':      str(r[8]).strip().upper() if r[8] else 'DEFERIDO',
            'observacao':     str(r[9]).strip() if r[9] else '',
        })
    return dados

# ── Normaliza para JSON do simulador ──────────────────────────────────────
def normalizar(dados):
    # Agrupa por comp_ufrn
    estrutura = {}
    # Primeiro: coleta metadados por componente
    meta = {}
    for r in dados:
        c = r['comp_ufrn']
        if c not in meta:
            meta[c] = {'nivel_vespertino': r['nivel_vesp'], 'nivel_noturno': r['nivel_not']}

    # Agrupa opções por (comp_ufrn, instituicao, id_opcao)
    opcoes_map = defaultdict(lambda: defaultdict(dict))
    for r in dados:
        c   = r['comp_ufrn']
        inst= r['instituicao']
        iop = r['id_opcao']
        if iop not in opcoes_map[c][inst]:
            # Extrai componentes: conjuntos separados por " + "
            if r['requer_conjunto']:
                comps = [x.strip() for x in r['componentes_str'].split(' + ') if x.strip()]
            else:
                comps = [r['componentes_str']] if r['componentes_str'] else []

            opcoes_map[c][inst][iop] = {
                'id_opcao':        iop,
                'num_opcao':       r['num_opcao'],
                'componentes':     comps,
                'requer_conjunto': r['requer_conjunto'],
                'resultado':       r['resultado'],
            }

    # Monta estrutura final — inclui DEFERIDO e INDEFERIDO
    for c in sorted(opcoes_map.keys()):
        estrutura[c] = {
            'nivel_vespertino': meta[c]['nivel_vespertino'],
            'nivel_noturno':    meta[c]['nivel_noturno'],
            'instituicoes': {}
        }
        for inst in sorted(opcoes_map[c].keys()):
            ops = sorted(opcoes_map[c][inst].values(), key=lambda x: x['num_opcao'])
            # Inclui DEFERIDO e INDEFERIDO; exclui linhas sem instituição
            ops_visiveis = [o for o in ops if o['resultado'] in ('DEFERIDO', 'INDEFERIDO')]
            if ops_visiveis:
                estrutura[c]['instituicoes'][inst] = [
                    {k: v for k, v in o.items()}
                    for o in ops_visiveis
                ]

    return estrutura

# ── Gera HTML ─────────────────────────────────────────────────────────────
def gerar_html(estrutura, template_path, saida_path):
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    dados_json  = json.dumps(estrutura, ensure_ascii=False, separators=(',', ':'))
    comps_lista = sorted(estrutura.keys())
    comps_js    = json.dumps(comps_lista, ensure_ascii=False)
    opcoes_comp = '\n'.join([
        '<option value="' + c.replace('"', '&quot;') + '">' + c + '</option>'
        for c in comps_lista
    ])

    # Substitui var DADOS usando delimitadores exatos (evita regex DOTALL guloso)
    marcador_dados = 'var DADOS = {'
    if marcador_dados in html:
        pos_di = html.index(marcador_dados)
        # Encontra o fechamento correto: navega pelos {} para achar o par
        depth = 0
        pos_df = pos_di + len('var DADOS = ')
        i = pos_df
        while i < len(html):
            if html[i] == '{':
                depth += 1
            elif html[i] == '}':
                depth -= 1
                if depth == 0:
                    pos_df = i + 2  # inclui o };
                    break
            i += 1
        html = html[:pos_di] + 'var DADOS = ' + dados_json + ';' + html[pos_df:]
    else:
        print('AVISO: var DADOS nao encontrado no template.')

    # Substitui var COMPS_LISTA
    marcador_comps = 'var COMPS_LISTA = ['
    if marcador_comps in html:
        pos_ci = html.index(marcador_comps)
        pos_cf = html.index('];', pos_ci) + 2
        html = html[:pos_ci] + 'var COMPS_LISTA = ' + comps_js + ';' + html[pos_cf:]
    else:
        print('AVISO: var COMPS_LISTA nao encontrado no template.')

    # Substitui as options do datalist
    html = re.sub(
        r'(<datalist id="lista-comps-ufrn">).*?(</datalist>)',
        r'\g<1>\n' + opcoes_comp + r'\n\g<2>',
        html, count=1, flags=re.DOTALL
    )

    with open(saida_path, 'w', encoding='utf-8') as f:
        f.write(html)

# ── Main ──────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    if not os.path.exists(XLSX):
        print(f'Erro: {XLSX} não encontrado.')
        sys.exit(1)
    if not os.path.exists(TEMPLATE):
        print(f'Erro: {TEMPLATE} não encontrado.')
        sys.exit(1)

    print(f'Lendo {XLSX}...')
    dados = ler_planilha(XLSX)
    print(f'  {len(dados)} linhas lidas.')

    print('Normalizando...')
    estrutura = normalizar(dados)
    total_opts = sum(
        len(ops)
        for v in estrutura.values()
        for ops in v['instituicoes'].values()
    )
    print(f'  {len(estrutura)} componentes UFRN, {total_opts} opções de equivalência.')

    print(f'Gerando {SAIDA}...')
    gerar_html(estrutura, TEMPLATE, SAIDA)
    print(f'Concluído. Arquivo gerado: {SAIDA}')
    print()
    print('Próximo passo: suba o index.html no GitHub para publicar o site atualizado.')