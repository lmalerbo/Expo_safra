"""
upload_releases_lote.py — Sobe em lote os arquivos .dwg/.zip de projetos já finalizados
para o GitHub Releases, via o mesmo proxy (Cloudflare Worker) usado pelo formulario.html.

Varre pasta_destino_dwg e pasta_destino_zip (config.json), agrupa por COD_FAZ
(um release por fazenda, tag = COD_FAZ) e envia cada arquivo via POST
{RELEASE_PROXY_URL}/upload?tag=...&name=...&filename=....

Uso:
  python engine/upload_releases_lote.py            # sobe tudo
  python engine/upload_releases_lote.py --dry-run  # só lista o que seria enviado
  python engine/upload_releases_lote.py --only 10003,10017  # só esses COD_FAZ
"""

import os
import re
import sys
import json
import argparse

import requests

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

RELEASE_PROXY_URL = 'https://expo-safra-release-proxy.leonardo-malerbo.workers.dev'

RE_ARQ_PROJETO = re.compile(r'^(\d+)_.+_Exp\d+L\.(dwg|zip)$', re.IGNORECASE)

with open(os.path.join(_BASE_DIR, 'config.json'), 'r', encoding='utf-8') as f:
    _cfg = json.load(f)

PASTA_DWG = _cfg['pasta_destino_dwg']
PASTA_ZIP = _cfg['pasta_destino_zip']
PREFIXOS_EXCLUIR = _cfg['codfaz_excluir_prefixo']
if isinstance(PREFIXOS_EXCLUIR, str):
    PREFIXOS_EXCLUIR = [PREFIXOS_EXCLUIR]
PREFIXOS_EXCLUIR = tuple(PREFIXOS_EXCLUIR)


def coletar_arquivos(pasta):
    arquivos = []
    if not os.path.isdir(pasta):
        print(f'Aviso: pasta não encontrada: {pasta}')
        return arquivos
    for nome in os.listdir(pasta):
        caminho = os.path.join(pasta, nome)
        if not os.path.isfile(caminho):
            continue
        m = RE_ARQ_PROJETO.match(nome)
        if not m:
            continue
        cod_faz = m.group(1)
        if cod_faz.startswith(PREFIXOS_EXCLUIR):
            continue
        arquivos.append((cod_faz, nome, caminho))
    return arquivos


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dry-run', action='store_true', help='Apenas lista o que seria enviado')
    parser.add_argument('--only', help='Lista de COD_FAZ separados por vírgula (filtra)')
    parser.add_argument('--skip', help='Lista de COD_FAZ separados por vírgula (ignora)')
    args = parser.parse_args()

    only = set(c.strip() for c in args.only.split(',')) if args.only else None
    skip = set(c.strip() for c in args.skip.split(',')) if args.skip else set()

    arquivos = coletar_arquivos(PASTA_DWG) + coletar_arquivos(PASTA_ZIP)

    grupos = {}
    for cod_faz, nome, caminho in arquivos:
        if only and cod_faz not in only:
            continue
        if cod_faz in skip:
            continue
        grupos.setdefault(cod_faz, []).append((nome, caminho))

    if not grupos:
        print('Nenhum arquivo encontrado.')
        return

    total_arquivos = sum(len(v) for v in grupos.values())
    print(f'{len(grupos)} fazenda(s), {total_arquivos} arquivo(s) no total.')

    erros = []
    for cod_faz in sorted(grupos):
        for nome, caminho in sorted(grupos[cod_faz]):
            if args.dry_run:
                print(f'[dry-run] {cod_faz} -> {nome}')
                continue

            url = f'{RELEASE_PROXY_URL}/upload'
            params = {'tag': cod_faz, 'name': cod_faz, 'filename': nome}
            print(f'Enviando {nome} (release {cod_faz})...')
            with open(caminho, 'rb') as f:
                res = requests.post(url, params=params, data=f,
                                     headers={'Content-Type': 'application/octet-stream'})
            if res.ok:
                print(f'  ok')
            else:
                print(f'  ERRO {res.status_code}: {res.text}')
                erros.append((cod_faz, nome, res.status_code, res.text))

    if erros:
        print(f'\n{len(erros)} erro(s):')
        for cod_faz, nome, status, texto in erros:
            print(f'  {cod_faz}/{nome}: {status} {texto}')
        sys.exit(1)
    elif not args.dry_run:
        print('\nTodos os arquivos enviados com sucesso!')


if __name__ == '__main__':
    main()
