"""
upload_releases_fazendas.py — Varre a pasta FAZENDAS (ciclo 2026-2027), encontra os
projetos de colheita já finalizados em
  {FAZENDA}\\2026-2027\\03 EXPORTACOES\\02 COLHEITA\\{N}L\\
e sobe os .dwg/.zip para o GitHub Releases via o mesmo proxy usado pelo formulario.html
(um release por COD_FAZ, tag = COD_FAZ).

Para cada subpasta "{N}L" dentro de "02 COLHEITA" (a "Exp0*" na raiz de "02 COLHEITA"
é o dwg mestre e é ignorada):

  - Se já existe o .dwg e o .zip no padrão {COD_FAZ}_{NOME}_Exp{N}L.{dwg|zip}:
    pronto para subir.
  - Se existe o .dwg mas falta o .zip, e as 4 pastas (AgData, AgGPS, GS3_2630, Gen4)
    estão presentes: compacta essas 4 pastas em {COD_FAZ}_{NOME}_Exp{N}L.zip
    (mesmo padrão dos zips já existentes) e então sobe.
  - Qualquer outra situação (falta dwg, faltam pastas para compactar, nome fora do
    padrão, etc.): registrada em um log para análise manual, nada é enviado.

Antes de subir, verifica os assets já existentes no release (mesmo tag/nome/tamanho)
para não reenviar o que já está no GitHub.

Uso:
  python engine/upload_releases_fazendas.py --scan         # só analisa e gera o log, não altera nada
  python engine/upload_releases_fazendas.py --dry-run      # mostra o que seria compactado/enviado
  python engine/upload_releases_fazendas.py                # compacta o que falta e envia tudo
  python engine/upload_releases_fazendas.py --only 10003,10017
"""

import os
import re
import sys
import json
import time
import zipfile
import argparse

import requests

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR = os.path.dirname(_SCRIPT_DIR)

RELEASE_PROXY_URL = 'https://expo-safra-release-proxy.leonardo-malerbo.workers.dev'
GH_API_RELEASES = 'https://api.github.com/repos/lmalerbo/Expo_safra/releases'

CICLO = '2026-2027'
SUBPATH_COLHEITA = os.path.join('03 EXPORTACOES', '02 COLHEITA')

RE_FAZENDA_DIR = re.compile(r'^(\d+)\s+(.+)$')
RE_LINHA_DIR = re.compile(r'^(\d+)[Ll]$')
RE_ARQ_PROJETO = re.compile(r'^(\d+)_.+_Exp(\d+)[Ll]\.(dwg|zip)$', re.IGNORECASE)

PASTAS_PROJETO = {'agdata', 'aggps', 'gen4', 'gs3_2630'}

with open(os.path.join(_BASE_DIR, 'config.json'), 'r', encoding='utf-8') as f:
    _cfg = json.load(f)

PREFIXOS_EXCLUIR = _cfg['codfaz_excluir_prefixo']
if isinstance(PREFIXOS_EXCLUIR, str):
    PREFIXOS_EXCLUIR = [PREFIXOS_EXCLUIR]
PREFIXOS_EXCLUIR = tuple(PREFIXOS_EXCLUIR)

LOG_PATH = os.path.join(_SCRIPT_DIR, 'upload_releases_fazendas_manual.log')


def encontrar_arquivo(pasta, ext):
    for nome in os.listdir(pasta):
        caminho = os.path.join(pasta, nome)
        if not os.path.isfile(caminho):
            continue
        m = RE_ARQ_PROJETO.match(nome)
        if m and m.group(3).lower() == ext:
            return nome, m
    return None, None


def pastas_projeto_presentes(pasta):
    presentes = set()
    for nome in os.listdir(pasta):
        caminho = os.path.join(pasta, nome)
        if os.path.isdir(caminho):
            presentes.add(nome.lower())
    return PASTAS_PROJETO & presentes == PASTAS_PROJETO


def compactar_pastas_projeto(pasta, nome_zip, dry_run=False):
    """Cria {pasta}/{nome_zip} com o conteúdo das 4 pastas de projeto (AgData, AgGPS, GS3_2630, Gen4)."""
    destino = os.path.join(pasta, nome_zip)
    if dry_run:
        print(f'  [dry-run] compactaria -> {nome_zip}')
        return destino

    nomes_pastas = [n for n in os.listdir(pasta)
                    if os.path.isdir(os.path.join(pasta, n)) and n.lower() in PASTAS_PROJETO]

    with zipfile.ZipFile(destino, 'w', zipfile.ZIP_DEFLATED) as zf:
        for nome_pasta in nomes_pastas:
            raiz_pasta = os.path.join(pasta, nome_pasta)
            for dirpath, dirnames, filenames in os.walk(raiz_pasta):
                rel_dir = os.path.relpath(dirpath, pasta)
                if not filenames and not dirnames:
                    zf.write(dirpath, rel_dir.replace(os.sep, '/') + '/')
                    continue
                for fname in filenames:
                    caminho = os.path.join(dirpath, fname)
                    rel = os.path.join(rel_dir, fname)
                    zf.write(caminho, rel.replace(os.sep, '/'))
    print(f'  compactado -> {nome_zip}')
    return destino


def coletar_situacoes(root_fazendas):
    prontos = []     # (cod_faz, [(nome_arquivo, caminho), ...])
    a_compactar = []  # (cod_faz, pasta, nome_zip, dwg_nome, dwg_caminho)
    manuais = []      # mensagens de log

    for nome_fazenda in sorted(os.listdir(root_fazendas)):
        caminho_fazenda = os.path.join(root_fazendas, nome_fazenda)
        if not os.path.isdir(caminho_fazenda):
            continue
        m = RE_FAZENDA_DIR.match(nome_fazenda)
        if not m:
            continue
        cod_faz = m.group(1)
        if cod_faz.startswith(PREFIXOS_EXCLUIR):
            continue

        colheita = os.path.join(caminho_fazenda, CICLO, SUBPATH_COLHEITA)
        if not os.path.isdir(colheita):
            continue

        for nome_linha in sorted(os.listdir(colheita)):
            pasta_linha = os.path.join(colheita, nome_linha)
            if not os.path.isdir(pasta_linha):
                continue
            ml = RE_LINHA_DIR.match(nome_linha)
            if not ml:
                continue

            dwg_nome, dwg_match = encontrar_arquivo(pasta_linha, 'dwg')
            zip_nome, _ = encontrar_arquivo(pasta_linha, 'zip')

            ctx = f'{nome_fazenda}/{CICLO}/{SUBPATH_COLHEITA}/{nome_linha}'

            if not dwg_nome:
                manuais.append(f'{ctx}: sem .dwg no padrão esperado')
                continue

            if dwg_match.group(1) != cod_faz:
                manuais.append(f'{ctx}: COD_FAZ do arquivo ({dwg_match.group(1)}) difere da pasta ({cod_faz})')
                continue

            if zip_nome:
                prontos.append((cod_faz, [
                    (dwg_nome, os.path.join(pasta_linha, dwg_nome)),
                    (zip_nome, os.path.join(pasta_linha, zip_nome)),
                ]))
                continue

            if pastas_projeto_presentes(pasta_linha):
                nome_zip = re.sub(r'\.dwg$', '.zip', dwg_nome, flags=re.IGNORECASE)
                a_compactar.append((cod_faz, pasta_linha, nome_zip, dwg_nome, os.path.join(pasta_linha, dwg_nome)))
                continue

            manuais.append(f'{ctx}: falta .zip e faltam pastas de projeto (AgData/AgGPS/GS3_2630/Gen4) para compactar')

    return prontos, a_compactar, manuais


def carregar_assets_existentes():
    """Busca todos os releases (paginado) de uma vez, evitando o rate limit da API
    não-autenticada do GitHub (60 req/h) ao consultar tag por tag."""
    assets_por_tag = {}
    page = 1
    while True:
        res = requests.get(GH_API_RELEASES, params={'per_page': 100, 'page': page})
        res.raise_for_status()
        releases = res.json()
        if not releases:
            break
        for rel in releases:
            assets_por_tag[rel['tag_name']] = {a['name']: a['size'] for a in rel.get('assets', [])}
        page += 1
    return assets_por_tag


def upload_arquivo(tag, nome, caminho, cache_assets):
    tamanho_local = os.path.getsize(caminho)
    if cache_assets.get(tag, {}).get(nome) == tamanho_local:
        print(f'  já existe no release {tag}, pulando: {nome}')
        return True

    url = f'{RELEASE_PROXY_URL}/upload'
    params = {'tag': tag, 'name': tag, 'filename': nome}

    for tentativa in range(1, 4):
        print(f'  enviando {nome} (release {tag})...' + (f' [tentativa {tentativa}]' if tentativa > 1 else ''))
        try:
            with open(caminho, 'rb') as f:
                res = requests.post(url, params=params, data=f,
                                     headers={'Content-Type': 'application/octet-stream'})
        except requests.exceptions.RequestException as e:
            if tentativa < 3:
                print(f'    erro de conexão ({e}), tentando novamente...')
                time.sleep(3)
                continue
            print(f'    ERRO de conexão: {e}')
            return False
        if res.ok:
            print('    ok')
            cache_assets.setdefault(tag, {})[nome] = tamanho_local
            return True
        if 'already_exists' in res.text:
            if tentativa < 3:
                print('    asset ainda sendo substituído, tentando novamente...')
                time.sleep(3)
                continue
            # API de releases tem delay de consistência: o proxy não viu o asset
            # recém-criado para substituí-lo, mas o asset já está presente no
            # release (enviado em execução anterior) — objetivo atingido.
            print('    asset já presente no release (already_exists), considerando ok')
            cache_assets.setdefault(tag, {})[nome] = tamanho_local
            return True
        print(f'    ERRO {res.status_code}: {res.text}')
        return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--root', default=r'G:\GeoProc\GEOTECNOLOGIA\08 PROJETOS PILOTO AUTOMATICO\15 PROPRIO\FAZENDAS')
    parser.add_argument('--scan', action='store_true', help='Só analisa e gera o log, não compacta nem envia')
    parser.add_argument('--dry-run', action='store_true', help='Mostra o que seria compactado/enviado, sem executar')
    parser.add_argument('--only', help='Lista de COD_FAZ separados por vírgula (filtra)')
    args = parser.parse_args()

    only = set(c.strip() for c in args.only.split(',')) if args.only else None

    prontos, a_compactar, manuais = coletar_situacoes(args.root)

    if only:
        prontos = [(c, arqs) for c, arqs in prontos if c in only]
        a_compactar = [t for t in a_compactar if t[0] in only]

    print(f'Prontos para subir (dwg+zip): {len(prontos)}')
    print(f'A compactar (dwg + pastas, falta zip): {len(a_compactar)}')
    print(f'Para análise manual: {len(manuais)}')

    with open(LOG_PATH, 'w', encoding='utf-8') as f:
        for linha in manuais:
            f.write(linha + '\n')
    print(f'Log de pendências manuais: {LOG_PATH}')

    if args.scan:
        for cod_faz, arqs in prontos:
            print(f'  pronto {cod_faz}: ' + ', '.join(n for n, _ in arqs))
        for cod_faz, pasta, nome_zip, dwg_nome, _ in a_compactar:
            print(f'  compactar {cod_faz}: {dwg_nome} + (AgData/AgGPS/GS3_2630/Gen4) -> {nome_zip}')
        return

    print('Carregando assets já existentes nos releases...')
    try:
        cache_assets = carregar_assets_existentes()
    except requests.exceptions.HTTPError as e:
        print(f'  não foi possível consultar (seguindo sem cache): {e}')
        cache_assets = {}
    erros = 0

    for cod_faz, pasta, nome_zip, dwg_nome, dwg_caminho in a_compactar:
        zip_caminho = compactar_pastas_projeto(pasta, nome_zip, dry_run=args.dry_run)
        if args.dry_run:
            print(f'  [dry-run] enviaria {dwg_nome} e {nome_zip} (release {cod_faz})')
            continue
        if not upload_arquivo(cod_faz, dwg_nome, dwg_caminho, cache_assets):
            erros += 1
        if not upload_arquivo(cod_faz, nome_zip, zip_caminho, cache_assets):
            erros += 1

    for cod_faz, arqs in prontos:
        if args.dry_run:
            print(f'  [dry-run] enviaria {[n for n, _ in arqs]} (release {cod_faz})')
            continue
        for nome, caminho in arqs:
            if not upload_arquivo(cod_faz, nome, caminho, cache_assets):
                erros += 1

    if erros:
        print(f'\n{erros} erro(s) durante o envio.')
        sys.exit(1)
    elif not args.dry_run:
        print('\nConcluído.')


if __name__ == '__main__':
    main()
