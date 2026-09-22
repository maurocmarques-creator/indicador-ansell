#!/usr/bin/env python3
"""
sync_mural_para_dashboard.py — Aplica as observacoes do Mural (campo
"observacoes_transito" de cliente_config.json) dentro do bloco de dados
(const RAW = {...}) ja publicado em index.html, sem precisar re-rodar a
extracao completa do portal Brudam.

Roda tanto localmente (chamado a mao ou por uma rotina agendada) quanto
no GitHub Actions (.github/workflows/sync-mural.yml), disparado sempre
que o Mural grava uma mensagem nova direto em cliente_config.json via
API do GitHub.

Uso: python sync_mural_para_dashboard.py [pasta_do_dashboard]
Padrao pasta_do_dashboard: diretorio deste script.
"""

import sys
import json
from pathlib import Path


def main():
    dash_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent
    config_path = dash_dir / 'cliente_config.json'
    index_path = dash_dir / 'index.html'

    cfg = json.loads(config_path.read_text(encoding='utf-8'))
    observacoes = cfg.get('observacoes_transito', {})

    html = index_path.read_text(encoding='utf-8')
    marker = 'const RAW = '
    start = html.index(marker) + len(marker)
    decoder = json.JSONDecoder()
    raw, end = decoder.raw_decode(html, start)

    alterados = 0
    for row in raw['rows']:
        minuta = str(row.get('MINUTA', ''))
        novas = observacoes.get(minuta, [])
        if row.get('OBSERVACOES', []) != novas:
            row['OBSERVACOES'] = novas
            alterados += 1

    if alterados:
        new_json = json.dumps(raw, ensure_ascii=False)
        html = html[:start] + new_json + html[end:]
        index_path.write_text(html, encoding='utf-8')

    print(f'Linhas de RAW com OBSERVACOES atualizadas: {alterados}')


if __name__ == '__main__':
    main()
