#!/usr/bin/env python3
"""
atualizar_dashboard.py — Atualiza index.html do Indicador Ansell a partir de
uma planilha consolidada (Ansell + Hercules) exportada do relatorio 106
(Emissoes) do sistema Brudam.

Uso: python atualizar_dashboard.py <planilha.xlsx> [pasta_do_dashboard]
Padrao pasta_do_dashboard: diretorio deste script.
"""

import sys
import json
from datetime import datetime
from pathlib import Path

import pandas as pd

MES_ABREV = {1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr', 5: 'Mai', 6: 'Jun',
             7: 'Jul', 8: 'Ago', 9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'}

UF_REGIAO = {
    'AC': 'Norte', 'AP': 'Norte', 'AM': 'Norte', 'PA': 'Norte', 'RO': 'Norte', 'RR': 'Norte', 'TO': 'Norte',
    'AL': 'Nordeste', 'BA': 'Nordeste', 'CE': 'Nordeste', 'MA': 'Nordeste', 'PB': 'Nordeste',
    'PE': 'Nordeste', 'PI': 'Nordeste', 'RN': 'Nordeste', 'SE': 'Nordeste',
    'DF': 'Centro-Oeste', 'GO': 'Centro-Oeste', 'MT': 'Centro-Oeste', 'MS': 'Centro-Oeste',
    'ES': 'Sudeste', 'MG': 'Sudeste', 'RJ': 'Sudeste', 'SP': 'Sudeste',
    'PR': 'Sul', 'RS': 'Sul', 'SC': 'Sul',
}

# Centroides aproximados por UF (uso apenas decorativo no mapa do dashboard).
UF_CENTROID = {
    'AC': (-9.02, -70.81), 'AL': (-9.57, -36.78), 'AM': (-3.47, -65.10),
    'AP': (0.90, -52.00), 'BA': (-12.96, -41.70), 'CE': (-5.50, -39.32),
    'DF': (-15.78, -47.93), 'ES': (-19.19, -40.34), 'GO': (-15.83, -49.83),
    'MA': (-5.42, -45.44), 'MG': (-18.51, -44.55), 'MS': (-20.77, -54.79),
    'MT': (-12.64, -55.42), 'PA': (-3.42, -52.29), 'PB': (-7.06, -36.72),
    'PE': (-8.38, -37.86), 'PI': (-7.72, -42.73), 'PR': (-24.89, -51.55),
    'RJ': (-22.25, -42.66), 'RN': (-5.40, -36.95), 'RO': (-10.83, -63.34),
    'RR': (1.99, -61.33), 'RS': (-30.17, -53.50), 'SC': (-27.45, -50.95),
    'SE': (-10.57, -37.45), 'SP': (-22.25, -48.63), 'TO': (-10.25, -48.25),
}


def iso(d):
    if pd.isna(d):
        return ''
    return d.strftime('%Y-%m-%d')


def compute_status(tipo_emissao, data_entrega, prev_entrega, data_agendamento):
    if tipo_emissao == 'DEVOLUCAO':
        return 'DEVOLUCAO'
    if tipo_emissao == 'REENTREGA':
        return 'REENTREGA'
    if pd.isna(data_entrega):
        return 'AGENDADO' if not pd.isna(data_agendamento) else 'EM TRÂNSITO'
    efetivo = data_agendamento if not pd.isna(data_agendamento) else prev_entrega
    if pd.isna(efetivo):
        return 'EM ATRASO'
    return 'NO PRAZO' if data_entrega <= efetivo else 'EM ATRASO'


def find_tipo_emissao_col(df):
    # O cabeçalho "TIPO EMISSÃO" costuma vir com o "Ã" corrompido no Excel
    # exportado pelo portal (caractere de substituição real, não só exibição).
    candidates = [c for c in df.columns if c.startswith('TIPO EMISS')]
    if not candidates:
        raise ValueError('Coluna "TIPO EMISSAO" nao encontrada na planilha')
    return candidates[0]


def build_rows(df):
    tipo_col = find_tipo_emissao_col(df)
    rows = []
    for _, r in df.iterrows():
        data_emissao = r['DATA EMISSAO']
        data_entrega = r['DATA ENTREGA']
        prev_entrega = r['PREV. ENTREGA']
        data_agendamento = r['DATA DE AGENDAMENTO']
        tipo = r[tipo_col]
        eff_uf = r['UF ENTREGA'] if not pd.isna(r['UF ENTREGA']) else ''

        rows.append({
            'MES': data_emissao.strftime('%Y-%m'),
            'MES_NOME': f"{MES_ABREV[data_emissao.month]}/{data_emissao.year}",
            'STATUS': compute_status(tipo, data_entrega, prev_entrega, data_agendamento),
            'CLIENTE': r['CLIENTE'],
            'MINUTA': str(r['MINUTA']),
            'NF_DOC': str(r['NF/DOC']),
            'VOLUMES': int(r['VOLUMES']) if not pd.isna(r['VOLUMES']) else 0,
            'FRETE TOTAL': float(r['FRETE TOTAL']) if not pd.isna(r['FRETE TOTAL']) else 0.0,
            'NF VALOR': float(r['NF VALOR']) if not pd.isna(r['NF VALOR']) else 0.0,
            'TX. PEDAGIO': float(r['TX. PEDAGIO']) if not pd.isna(r['TX. PEDAGIO']) else 0.0,
            'VALOR ICMS': float(r['VALOR ICMS']) if not pd.isna(r['VALOR ICMS']) else 0.0,
            'TX. GRIS': float(r['TX. GRIS']) if not pd.isna(r['TX. GRIS']) else 0.0,
            'TX. FRETE PESO': float(r['TX. FRETE PESO']) if not pd.isna(r['TX. FRETE PESO']) else 0.0,
            'TX. OUTROS': float(r['TX. OUTROS']) if not pd.isna(r['TX. OUTROS']) else 0.0,
            'TX. NOTA': float(r['TX. NOTA']) if not pd.isna(r['TX. NOTA']) else 0.0,
            'TIPO EMISSÃO': tipo,
            # COTACAO no export bruto e um numero de cotacao (quando negociado
            # fora da tabela) ou vazio — o filtro Sim/Nao do dashboard espera
            # 'S'/'N', entao convertemos presenca/ausencia de valor.
            'COTACAO': 'S' if not pd.isna(r['COTACAO']) else 'N',
            'DATA EMISSAO': iso(data_emissao),
            'DATA ENTREGA': iso(data_entrega),
            'PREV. ENTREGA': iso(prev_entrega),
            'DATA DE AGENDAMENTO': iso(data_agendamento),
            'EFF_LOCAL': r['LOCAL ENTREGA'] if not pd.isna(r['LOCAL ENTREGA']) else '',
            'EFF_CIDADE': r['CIDADE ENTREGA'] if not pd.isna(r['CIDADE ENTREGA']) else '',
            'EFF_UF': eff_uf,
            'REGIAO': UF_REGIAO.get(eff_uf, ''),
            'LAT': UF_CENTROID[eff_uf][0] if eff_uf in UF_CENTROID else None,
            'LNG': UF_CENTROID[eff_uf][1] if eff_uf in UF_CENTROID else None,
        })
    return rows


def build_raw(rows, gerado_em=None):
    meses = sorted(set(r['MES'] for r in rows))
    tipos = sorted(set(r['TIPO EMISSÃO'] for r in rows))
    ufs = sorted(set(r['EFF_UF'] for r in rows if r['EFF_UF']))
    clientes = sorted(set(r['CLIENTE'] for r in rows))
    return {
        'meta': {
            'meses': meses,
            'tipos_emissao': tipos,
            'ufs': ufs,
            'clientes': clientes,
            'gerado_em': gerado_em or datetime.now().strftime('%d/%m/%Y %H:%M'),
        },
        'rows': rows,
    }


def read_json_blob(content, marker):
    pos = content.find(marker)
    if pos == -1:
        return None
    start = pos + len(marker)
    decoder = json.JSONDecoder()
    data, _ = decoder.raw_decode(content, start)
    return data


def replace_json_blob(content, marker, new_json):
    pos = content.find(marker)
    if pos == -1:
        raise ValueError(f'Marcador nao encontrado: {marker!r}')
    start = pos + len(marker)
    decoder = json.JSONDecoder()
    _, end = decoder.raw_decode(content, start)
    return content[:start] + new_json + content[end:]


def main():
    if len(sys.argv) < 2:
        print('Uso: python atualizar_dashboard.py <planilha.xlsx> [pasta_do_dashboard]')
        sys.exit(1)

    xlsx_path = Path(sys.argv[1])
    dash_dir = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).parent
    index_path = dash_dir / 'index.html'

    if not xlsx_path.exists():
        print(f'Planilha nao encontrada: {xlsx_path}')
        sys.exit(1)

    print(f'Lendo {xlsx_path}...')
    df = pd.read_excel(xlsx_path, sheet_name='Brudam')
    print(f'{len(df)} linhas carregadas.')

    rows = build_rows(df)
    raw = build_raw(rows)

    print(f"Periodo: {raw['meta']['meses'][0]} a {raw['meta']['meses'][-1]}")
    print(f"Clientes: {raw['meta']['clientes']}")
    print(f"Linhas RAW: {len(rows)}")

    index_content = index_path.read_text(encoding='utf-8')
    raw_json = json.dumps(raw, ensure_ascii=False)
    index_content = replace_json_blob(index_content, 'const RAW = ', raw_json)
    index_path.write_text(index_content, encoding='utf-8')
    print(f'Atualizado: {index_path}')


if __name__ == '__main__':
    main()
