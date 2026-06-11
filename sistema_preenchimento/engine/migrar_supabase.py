"""
migrar_supabase.py — Migração única dos dados de programacao_frentes.xlsx + usuarios.json
para as tabelas do Supabase (programacao, log_exportacoes, usuarios).

Uso:
  1. Crie sistema_preenchimento/supabase_config.json (gitignored) com:
     { "url": "https://xxxx.supabase.co", "secret_key": "sb_secret_..." }
  2. Rode `supabase db push` antes (cria as tabelas via migration 0001_init.sql).
  3. python engine/migrar_supabase.py
"""

import os
import sys
import json

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR   = os.path.dirname(_SCRIPT_DIR)
sys.path.insert(0, _SCRIPT_DIR)
from utils import layer_to_str

import pandas as pd
import requests

MESTRE = os.path.join(_BASE_DIR, 'programacao_frentes.xlsx')
USUARIOS_JSON = os.path.join(_BASE_DIR, 'usuarios.json')
CONFIG_PATH = os.path.join(_BASE_DIR, 'supabase_config.json')

with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
    _cfg = json.load(f)

SUPABASE_URL = _cfg['url'].rstrip('/')
SECRET_KEY   = _cfg['secret_key']

HEADERS = {
    'apikey': SECRET_KEY,
    'Authorization': f'Bearer {SECRET_KEY}',
    'Content-Type': 'application/json',
    'Prefer': 'resolution=merge-duplicates,return=minimal',
}


def _safe_int(v):
    try:
        return int(v) if pd.notna(v) else None
    except (ValueError, TypeError):
        return None


def _safe_str(v):
    s = str(v or '').strip()
    return '' if s == 'nan' else s


def _safe_float(v):
    try:
        return round(float(v), 2) if pd.notna(v) else 0
    except (ValueError, TypeError):
        return 0


def _upsert(table, rows, batch_size=500):
    if not rows:
        print(f"  {table}: nada para enviar.")
        return
    total = len(rows)
    for i in range(0, total, batch_size):
        chunk = rows[i:i + batch_size]
        res = requests.post(f'{SUPABASE_URL}/rest/v1/{table}', headers=HEADERS, json=chunk)
        if not res.ok:
            raise RuntimeError(f"Erro ao enviar {table} ({i}-{i+len(chunk)}): {res.status_code} {res.text}")
        print(f"  {table}: {min(i+batch_size, total)}/{total}")


# ── 1. Programação ─────────────────────────────────────────────────────────
print("Lendo aba 'Programação'...")
df_prog = pd.read_excel(MESTRE)
prog_rows = []
for _, row in df_prog.iterrows():
    layer = _safe_int(row.get('LAYER'))
    if layer is None:
        continue
    prog_rows.append({
        'layer':      layer,
        'frente':     _safe_int(row.get('FRENTE')),
        'periodo_op': _safe_int(row.get('PERÍODO OP')),
        'cod_faz':    _safe_int(row.get('COD FAZ')),
        'fazenda':    _safe_str(row.get('FAZENDA')),
        'talhao':     _safe_int(row.get('TALHÕES')),
        'status':     _safe_str(row.get('EXPORTAÇÃO')),
        'tipo_linha': _safe_str(row.get('TIPO DE LINHA')),
        'ciclo':      _safe_str(row.get('CICLO')),
        'area_ha':    _safe_float(row.get('AREA_HA')),
        'estagio':    _safe_str(row.get('ESTÁGIO')),
    })
print(f"  {len(prog_rows)} talhões.")

layer_ha = {str(r['layer']): r['area_ha'] for r in prog_rows}

# ── 2. Log Exportações ──────────────────────────────────────────────────────
print("Lendo aba 'Log Exportações'...")
df_log = pd.read_excel(MESTRE, sheet_name='Log Exportações')
log_rows = []
for _, row in df_log.iterrows():
    layer = layer_to_str(row.get('LAYER'))
    if not layer:
        continue
    log_rows.append({
        'data_consolidacao': pd.to_datetime(row.get('DATA_CONSOLIDACAO'), dayfirst=True, errors='coerce').isoformat()
                              if pd.notna(row.get('DATA_CONSOLIDACAO')) else None,
        'registrado_em': pd.to_datetime(row.get('TIMESTAMP'), dayfirst=True, errors='coerce').isoformat()
                          if pd.notna(row.get('TIMESTAMP')) else None,
        'usuario':    _safe_str(row.get('USUARIO')),
        'layer':      int(layer),
        'fazenda':    _safe_str(row.get('FAZENDA')),
        'talhao':     _safe_int(row.get('TALHAO')),
        'tipo_linha': _safe_str(row.get('TIPO_LINHA')),
        'ciclo':      _safe_str(row.get('CICLO')),
        'status':     _safe_str(row.get('STATUS')),
    })
print(f"  {len(log_rows)} registros de log.")

# ── 3. Usuários (perfis de usuarios.json + ha calculado do log) ────────────
print("Lendo usuarios.json e calculando ha por usuário...")
usuarios_config = []
if os.path.exists(USUARIOS_JSON):
    with open(USUARIOS_JSON, 'r', encoding='utf-8') as f:
        usuarios_config = json.load(f)

df_log_dedup = df_log.drop_duplicates(subset='LAYER', keep='last')
usuario_ha = {}
for _, row in df_log_dedup.iterrows():
    usuario = _safe_str(row.get('USUARIO'))
    if not usuario:
        continue
    ly = layer_to_str(row.get('LAYER'))
    if ly:
        usuario_ha[usuario] = round(usuario_ha.get(usuario, 0) + layer_ha.get(ly, 0), 2)

usuarios_rows = []
nomes_vistos = set()
for u in usuarios_config:
    nome = u['nome']
    nomes_vistos.add(nome)
    usuarios_rows.append({'nome': nome, 'perfis': u.get('perfis', []), 'ha': usuario_ha.get(nome, 0)})
for nome, ha in usuario_ha.items():
    if nome not in nomes_vistos:
        usuarios_rows.append({'nome': nome, 'perfis': [], 'ha': ha})
print(f"  {len(usuarios_rows)} usuários.")

# ── 4. Envio ─────────────────────────────────────────────────────────────────
print("\nEnviando para o Supabase...")
_upsert('programacao', prog_rows)
_upsert('log_exportacoes', log_rows)
_upsert('usuarios', usuarios_rows)

print("\nMigração concluída.")
