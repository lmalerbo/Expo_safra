"""
atualizar_programacao.py
Uso: python atualizar_programacao.py  OU  duplo clique no ATUALIZAR.bat

Estrutura esperada:
  base_icol/*.xlsm          ← base ICOL (agendamento por frente)
  base_fazendas/*.xlsx      ← base mestre de talhões (área, estágio)
  supabase_config.json      ← { "url": "...", "secret_key": "sb_secret_..." }

Lê a base ICOL + base de fazendas, monta a programação por talhão e faz
upsert na tabela `programacao` do Supabase, preservando STATUS/TIPO_LINHA/
CICLO já preenchidos para LAYERs existentes.
"""

import os
import sys

# ── Utilitários compartilhados ────────────────────────────────────────────
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR   = os.path.dirname(_SCRIPT_DIR)   # sistema_preenchimento/
sys.path.insert(0, _SCRIPT_DIR)
from utils import layer_to_str, redirecionar_stdout, fechar_log

_log_fh = redirecionar_stdout(os.path.join(_BASE_DIR, 'logs', 'atualizar.log'))

import pandas as pd
import datetime, time, json
import glob as _glob
import requests

# ── Carrega configurações ─────────────────────────────────────────────────
_cfg_path = os.path.join(_BASE_DIR, 'config.json')
try:
    with open(_cfg_path, 'r', encoding='utf-8') as _f:
        _cfg = json.load(_f)
except Exception:
    _cfg = {}

CODFAZ_EXCLUIR_PREFIXO = _cfg.get('codfaz_excluir_prefixo', '20')

_config_path = os.path.join(_BASE_DIR, 'supabase_config.json')
if not os.path.exists(_config_path):
    print(f"ERRO: Arquivo nao encontrado → {_config_path}")
    print("  Crie esse arquivo com: { \"url\": \"https://xxxx.supabase.co\", \"secret_key\": \"sb_secret_...\" }")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)

with open(_config_path, 'r', encoding='utf-8') as _f:
    _sb_cfg = json.load(_f)

SUPABASE_URL = _sb_cfg['url'].rstrip('/')
SECRET_KEY   = _sb_cfg['secret_key']
SB_HEADERS = {
    'apikey': SECRET_KEY,
    'Authorization': f'Bearer {SECRET_KEY}',
    'Content-Type': 'application/json',
}

# ── Localiza o .xlsm em base_icol/ ───────────────────────────────────────
os.chdir(_BASE_DIR)
_xlsm_found = _glob.glob("base_icol/*.xlsm")
if not _xlsm_found:
    print("ERRO: Nenhum arquivo .xlsm encontrado em base_icol/")
    print("  Coloque a base ICOL na pasta base_icol/ e tente novamente.")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)
SOURCE_XLSM = _xlsm_found[0]
print(f"Base ICOL encontrada: {SOURCE_XLSM}")

# ── 1. Ler valores existentes no Supabase (preservar preenchimento) ──────
print("Lendo programação existente no Supabase...")
_res = requests.get(f"{SUPABASE_URL}/rest/v1/programacao?select=layer,status,tipo_linha,ciclo", headers=SB_HEADERS)
if not _res.ok:
    print(f"ERRO ao ler programacao: {_res.status_code} {_res.text}")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)
preserved = {}   # layer_str → (status, tipo, ciclo)
for row in _res.json():
    layer = layer_to_str(row.get('layer'))
    if layer:
        preserved[layer] = (row.get('status') or '', row.get('tipo_linha') or '', row.get('ciclo') or '')
print(f"  {len(preserved)} linhas existentes carregadas.\n")

# ── 2. Ler base_fazendas (mestre de talhões) ─────────────────────────────
print("Verificando base fazendas...")
_base_faz_files = _glob.glob("base_fazendas/*.xls*")
df_base = None
if not _base_faz_files:
    print("  AVISO: Nenhum arquivo em base_fazendas/ — usando apenas ICOL.\n")
else:
    SOURCE_BASE = _base_faz_files[0]
    print(f"  Base fazendas: {SOURCE_BASE}")
    df_base = pd.read_excel(SOURCE_BASE, engine='openpyxl')
    df_base = df_base.rename(columns={
        'SECAO':     'COD FAZ',
        'DESC_SECAO':'FAZENDA',
        'TALHAO':    'TALHOES',
        'AREA_PROD': 'AREA_HA',
    })
    df_base['COD FAZ'] = pd.to_numeric(df_base['COD FAZ'], errors='coerce')
    df_base['TALHOES'] = pd.to_numeric(df_base['TALHOES'], errors='coerce')
    df_base['AREA_HA'] = pd.to_numeric(df_base['AREA_HA'], errors='coerce')
    if 'ESTAGIO' not in df_base.columns:
        df_base['ESTAGIO'] = ''
    df_base = df_base.dropna(subset=['COD FAZ', 'TALHOES']).reset_index(drop=True)
    print(f"  {len(df_base)} talhões na base fazendas.\n")

# ── 3. Ler base ICOL (agendamento por fazenda) ───────────────────────────
print("Lendo base ICOL...")
df_raw = pd.read_excel(SOURCE_XLSM, sheet_name='BASE PARA PLANEJAMENTO', header=None, engine='openpyxl')

# Encontra primeira linha de dados (COD FAZ em AA=índice 26 deve ser numérico)
first_data_row = None
for idx, row_vals in df_raw.iterrows():
    try:
        val = float(str(row_vals.iloc[26]))
        if not pd.isna(val):
            first_data_row = idx
            break
    except (ValueError, TypeError):
        continue

if first_data_row is None:
    print("ERRO: Dados não encontrados na aba BASE PARA PLANEJAMENTO")
    print("  Verifique se a aba existe e se COD FAZENDA está na coluna AA.")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)

# H=7 FRENTE, I=8 PERÍODO OP (mês), AA=26 COD FAZ, AB=27 FAZENDA
df_data = df_raw.iloc[first_data_row:, [7, 8, 26, 27]].copy()
df_data.columns = ['FRENTE', 'PERIODO_OP', 'COD FAZ', 'FAZENDA']

df_icol = df_data.dropna(subset=['COD FAZ']).copy()
df_icol['COD FAZ']    = pd.to_numeric(df_icol['COD FAZ'],    errors='coerce')
df_icol['FRENTE']     = pd.to_numeric(df_icol['FRENTE'],     errors='coerce')
df_icol['PERIODO_OP'] = pd.to_numeric(df_icol['PERIODO_OP'], errors='coerce')
df_icol = df_icol.dropna(subset=['COD FAZ'])
n_icol_raw = len(df_icol)
df_icol = df_icol.drop_duplicates(subset=['COD FAZ'], keep='first').reset_index(drop=True)
n_dup_icol = n_icol_raw - len(df_icol)
print(f"  {len(df_icol)} fazendas únicas no ICOL{f' ({n_dup_icol} linhas duplicadas ignoradas)' if n_dup_icol else ''}.\n")

# ── 4. Merge: ICOL (fazenda) LEFT JOIN base_fazendas (talhões) ───────────
if df_base is not None:
    result = df_icol[['COD FAZ', 'FAZENDA', 'FRENTE', 'PERIODO_OP']].merge(
        df_base[['COD FAZ', 'TALHOES', 'AREA_HA', 'ESTAGIO']],
        on='COD FAZ',
        how='left'
    )
    sem_ha = result['AREA_HA'].isna().sum()
    print(f"  {len(result)} talhões expandidos do ICOL — {len(result)-sem_ha} com AREA_HA, {sem_ha} sem.\n")
    faz_sem_base = result[result['TALHOES'].isna()].drop_duplicates('COD FAZ')
    n_sem_base = len(faz_sem_base)
    if n_sem_base:
        print(f"  ⚠  {n_sem_base} fazenda(s) do ICOL sem talhões na base_fazendas (serão omitidas):")
        for _, r in faz_sem_base.head(10).iterrows():
            print(f"     COD FAZ {int(r['COD FAZ'])}: {r['FAZENDA']}")
        if n_sem_base > 10:
            print(f"     ... e mais {n_sem_base - 10}")
        print()
else:
    n_sem_base = 0
    result = df_icol.copy()
    result['TALHOES'] = None
    result['AREA_HA'] = None
    result['ESTAGIO'] = ''

result = result.sort_values(['FRENTE', 'PERIODO_OP']).reset_index(drop=True)

# Fazendas com COD FAZ iniciando no prefixo configurado são unidades administrativas
n_antes = result['COD FAZ'].nunique()
result   = result[~result['COD FAZ'].astype(str).str.startswith(CODFAZ_EXCLUIR_PREFIXO)].reset_index(drop=True)
n_excluidas = n_antes - result['COD FAZ'].nunique()
if n_excluidas:
    print(f"  Filtro administrativo (COD FAZ {CODFAZ_EXCLUIR_PREFIXO}x): {n_excluidas} fazenda(s) excluída(s).\n")

result = result.dropna(subset=['TALHOES']).reset_index(drop=True)

def make_layer(row):
    try: return int(f"{int(row['COD FAZ'])}{int(row['TALHOES']):03d}")
    except (ValueError, TypeError, KeyError): return None

result['LAYER'] = result.apply(make_layer, axis=1)
result = result.dropna(subset=['LAYER']).reset_index(drop=True)

print(f"  {len(result)} linhas para processar.\n")

# ── 5. Montar registros e fazer upsert no Supabase ────────────────────────
print("Montando registros para o Supabase...")
novos = 0
prog_rows = []
for _, row in result.iterrows():
    layer_val = int(row['LAYER'])
    ly_str = layer_to_str(layer_val)
    if ly_str in preserved:
        status, tipo, ciclo = preserved[ly_str]
    else:
        status, tipo, ciclo = '', '', ''
        novos += 1
    try:
        area_ha = round(float(row.get('AREA_HA', 0) or 0), 2)
    except (ValueError, TypeError):
        area_ha = 0
    prog_rows.append({
        'layer':      layer_val,
        'frente':     int(row['FRENTE']) if pd.notna(row.get('FRENTE')) else None,
        'periodo_op': int(row['PERIODO_OP']) if pd.notna(row.get('PERIODO_OP')) else None,
        'cod_faz':    int(row['COD FAZ']) if pd.notna(row['COD FAZ']) else None,
        'fazenda':    str(row['FAZENDA']) if pd.notna(row['FAZENDA']) else '',
        'talhao':     int(row['TALHOES']) if pd.notna(row.get('TALHOES')) else None,
        'status':     status,
        'tipo_linha': tipo,
        'ciclo':      ciclo,
        'area_ha':    area_ha,
        'estagio':    str(row.get('ESTAGIO', '') or '').strip(),
    })

print(f"Enviando {len(prog_rows)} linhas (upsert por LAYER)...")
HEADERS_UPSERT = dict(SB_HEADERS, Prefer='resolution=merge-duplicates,return=minimal')
BATCH = 500
for i in range(0, len(prog_rows), BATCH):
    chunk = prog_rows[i:i+BATCH]
    res = requests.post(f"{SUPABASE_URL}/rest/v1/programacao", headers=HEADERS_UPSERT, json=chunk)
    if not res.ok:
        print(f"ERRO ao enviar lote {i}-{i+len(chunk)}: {res.status_code} {res.text}")
        fechar_log(_log_fh)
        input("\nPressione Enter para sair...")
        sys.exit(1)
    print(f"  {min(i+BATCH, len(prog_rows))}/{len(prog_rows)}")

print(f"\n  Upsert concluído. Novas linhas: {novos}\n")

# ── Aviso: LAYERs com preenchimento removidos desta atualização ───────────
layers_novos_str = {layer_to_str(r['layer']) for r in prog_rows}
layers_com_preenchimento = {ly for ly, vals in preserved.items() if any(str(v).strip() for v in vals)}
removidos_icol = layers_com_preenchimento - layers_novos_str
if removidos_icol:
    print(f"  ⚠  ATENÇÃO: {len(removidos_icol)} LAYER(s) preenchidos não estão na nova base ICOL:")
    for ly in sorted(removidos_icol)[:10]:
        status, tipo, ciclo = preserved[ly]
        print(f"     LAYER {ly}: {status} | {tipo} | {ciclo}")
    if len(removidos_icol) > 10:
        print(f"     ... e mais {len(removidos_icol)-10}")
    print("     Esses dados continuam no Supabase — apenas saíram do ICOL atual.\n")

# ── Resumo ────────────────────────────────────────────────────────────────
print(f"{'='*50}")
print(f"  Atualizacao concluida!")
print(f"  Total talhoes : {len(result)}")
if df_base is not None:
    sem_ha = result['AREA_HA'].isna().sum()
    print(f"  Com AREA_HA   : {len(result) - sem_ha}")
    print(f"  Sem AREA_HA   : {sem_ha}")
print(f"  Preservados   : {len(result) - novos}")
print(f"{'='*50}")

fechar_log(_log_fh)
input("\nPressione Enter para fechar...")
