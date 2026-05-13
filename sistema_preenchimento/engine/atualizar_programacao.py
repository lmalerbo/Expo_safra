"""
atualizar_programacao.py
Uso: python atualizar_programacao.py  OU  duplo clique no ATUALIZAR.bat

Estrutura esperada:
  base_icol/*.xlsm          ← base ICOL (agendamento por frente)
  base_fazendas/*.xlsx      ← base mestre de talhões (área, estágio)
  programacao_frentes.xlsx
  formulario.html
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import datetime, os, sys, time, json, re

DEST_FILE   = "programacao_frentes.xlsx"
HTML_FILE   = "formulario.html"

# Índices de coluna na aba "Programação" (1-based, conforme openpyxl)
COL_LAYER   = 1
COL_FRENTE  = 2
COL_PERIODO = 3
COL_DIA     = 4
COL_CODFAZ  = 5
COL_FAZENDA = 6
COL_TALHAO  = 7
COL_STATUS  = 8   # cabeçalho na planilha: EXPORTAÇÃO
COL_TIPO    = 9   # cabeçalho: TIPO DE LINHA
COL_CICLO   = 10
COL_AREA_HA = 11
COL_ESTAGIO = 12

# ── Localiza o .xlsm em base_icol/ ───────────────────────────────────────
import glob as _glob
_xlsm_found = _glob.glob("base_icol/*.xlsm")
if not _xlsm_found:
    print("ERRO: Nenhum arquivo .xlsm encontrado em base_icol/")
    print("  Coloque a base ICOL na pasta base_icol/ e tente novamente.")
    input("\nPressione Enter para sair...")
    sys.exit(1)
SOURCE_XLSM = _xlsm_found[0]
print(f"Base ICOL encontrada: {SOURCE_XLSM}")

# ── Verificações ──────────────────────────────────────────────────────────
for f in [DEST_FILE, HTML_FILE]:
    if not os.path.exists(f):
        print(f"ERRO: Arquivo nao encontrado → {f}")
        input("\nPressione Enter para sair...")
        sys.exit(1)

def arquivo_bloqueado(path):
    try:
        with open(path, 'a+b'): return False
    except (IOError, PermissionError): return True

print("Verificando se a planilha esta aberta...")
tentativas = 0
while arquivo_bloqueado(DEST_FILE):
    tentativas += 1
    if tentativas >= 3:
        print(f"ERRO: '{DEST_FILE}' está aberto. Feche e tente novamente.")
        input("\nPressione Enter para sair...")
        sys.exit(1)
    print(f"  Arquivo em uso, aguardando 10s... ({tentativas}/3)")
    time.sleep(10)
print("  OK.\n")

# ── 1. Ler valores existentes ─────────────────────────────────────────────
print("Lendo dados existentes na planilha...")
wb_ex = load_workbook(DEST_FILE, data_only=True)
ws_ex = wb_ex["Programação"]
preserved = {}
for row in ws_ex.iter_rows(min_row=2, values_only=True):
    layer = row[0]
    if layer:
        # índices 0-based → COL_STATUS-1, COL_TIPO-1, COL_CICLO-1
        preserved[int(layer)] = (row[COL_STATUS-1] or '', row[COL_TIPO-1] or '', row[COL_CICLO-1] or '')
wb_ex.close()
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
    # Fazendas do ICOL sem nenhum talhão na base_fazendas (não aparecerão no dashboard)
    faz_sem_base = result[result['TALHOES'].isna()].drop_duplicates('COD FAZ')
    n_sem_base = len(faz_sem_base)
    if n_sem_base:
        print(f"  ⚠  {n_sem_base} fazenda(s) do ICOL sem talhões na base_fazendas (serão omitidas do dashboard):")
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

# Fazendas com COD FAZ iniciando em 20 são unidades administrativas (sede, almoxarifado)
# que não têm colheita programada — devem ser excluídas da programação
CODFAZ_EXCLUIR_PREFIXO = '20'
n_antes_20x = result['COD FAZ'].nunique()
result = result[~result['COD FAZ'].astype(str).str.startswith(CODFAZ_EXCLUIR_PREFIXO)].reset_index(drop=True)
n_excluidas_20x = n_antes_20x - result['COD FAZ'].nunique()
if n_excluidas_20x:
    print(f"  Filtro administrativo (COD FAZ {CODFAZ_EXCLUIR_PREFIXO}x): {n_excluidas_20x} fazenda(s) excluída(s).\n")

def make_layer(row):
    try: return int(f"{int(row['COD FAZ'])}{int(row['TALHOES']):03d}")
    except (ValueError, TypeError, KeyError): return None

result['LAYER'] = result.apply(make_layer, axis=1)

# ha_lookup para consulta rápida (COD FAZ, TALHOES) → AREA_HA
ha_lookup = {}
if df_base is not None:
    for _, row in df_base.iterrows():
        try:
            key = (int(row['COD FAZ']), int(row['TALHOES']))
            if pd.notna(row.get('AREA_HA')):
                ha_lookup[key] = round(float(row['AREA_HA']), 2)
        except (ValueError, TypeError):
            pass

# Índice LAYER → ha (para usar na leitura do log)
layer_ha = {}
for _, row in result.iterrows():
    try:
        ly = int(row['LAYER']) if pd.notna(row['LAYER']) else None
        if ly is not None:
            layer_ha[ly] = ha_lookup.get((int(row['COD FAZ']), int(row['TALHOES'])), 0)
    except (ValueError, TypeError):
        pass

print(f"  {len(result)} linhas para processar.\n")

# ── 3. Atualizar planilha ─────────────────────────────────────────────────
print("Atualizando planilha...")
wb = load_workbook(DEST_FILE)
ws = wb["Programação"]
ws.delete_rows(2, ws.max_row)

frente_row_colors = {2:"DDEEFF",3:"E8F2FB",4:"D6E4F5",5:"EBF3FF",
                     6:"F2F8FD",8:"DDEEFF",9:"E8F2FB",10:"D6E4F5"}
thin   = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

today = datetime.date.today()

# Garante cabeçalhos para colunas que podem não existir em templates antigos
if ws.cell(1, COL_PERIODO).value not in ('PERÍODO OP', 'PERÍODO OPERACIONAL'):
    h = ws.cell(1, COL_PERIODO, value='PERÍODO OP')
    h.font      = Font(name="Arial", bold=True, size=10)
    h.border    = border
    h.alignment = Alignment(horizontal="center", vertical="center")
if ws.cell(1, COL_AREA_HA).value != 'AREA_HA':
    h = ws.cell(1, COL_AREA_HA, value='AREA_HA')
    h.font      = Font(name="Arial", bold=True, size=10)
    h.border    = border
    h.alignment = Alignment(horizontal="center", vertical="center")
if ws.cell(1, COL_ESTAGIO).value != 'ESTÁGIO':
    h = ws.cell(1, COL_ESTAGIO, value='ESTÁGIO')
    h.font      = Font(name="Arial", bold=True, size=10)
    h.border    = border
    h.alignment = Alignment(horizontal="center", vertical="center")

novos = 0
for i, row in result.iterrows():
    r      = i + 2
    frente = int(row['FRENTE']) if pd.notna(row.get('FRENTE')) else 0
    cor    = frente_row_colors.get(frente, "FFFFFF")
    fill   = PatternFill("solid", start_color=cor, end_color=cor)
    layer  = row['LAYER']
    if layer in preserved:
        exp, tipo, ciclo = preserved[layer]
    else:
        exp, tipo, ciclo = '', '', ''
        novos += 1
    try:
        ha_val = ha_lookup.get((int(row['COD FAZ']), int(row['TALHOES'])), 0)
    except (ValueError, TypeError):
        ha_val = 0
    estagio = str(row.get('ESTAGIO', '') or '').strip()
    periodo = int(row['PERIODO_OP']) if pd.notna(row.get('PERIODO_OP')) else ''
    vals = [
        layer if layer else '', frente,
        periodo,
        '',  # COL_DIA — sem data na nova estrutura
        int(row['COD FAZ']) if pd.notna(row['COD FAZ']) else '',
        str(row['FAZENDA']) if pd.notna(row['FAZENDA']) else '',
        int(row['TALHOES']) if pd.notna(row.get('TALHOES')) else '',
        exp, tipo, ciclo,
        ha_val,
        estagio,
    ]
    for c, val in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font      = Font(name="Arial", size=10)
        cell.fill      = fill
        cell.border    = border
        cell.alignment = Alignment(horizontal="left" if c==COL_FAZENDA else "center", vertical="center")
        if c == COL_AREA_HA:
            cell.number_format = '#,##0.00'

wb.save(DEST_FILE)
print(f"  Planilha atualizada. Novas linhas: {novos}\n")

# ── Aviso: LAYERs com preenchimento removidos desta atualização ───────────
layers_novos = set(result['LAYER'].dropna().astype(int))
layers_com_preenchimento = {ly for ly, vals in preserved.items() if any(str(v).strip() for v in vals)}
removidos_icol = layers_com_preenchimento - layers_novos
if removidos_icol:
    print(f"  ⚠  ATENÇÃO: {len(removidos_icol)} LAYER(s) preenchidos foram removidos desta base ICOL:")
    for ly in sorted(removidos_icol)[:10]:
        exp, tipo, ciclo = preserved[ly]
        print(f"     LAYER {ly}: {exp} | {tipo} | {ciclo}")
    if len(removidos_icol) > 10:
        print(f"     ... e mais {len(removidos_icol)-10}")
    print("     Esses dados NÃO foram perdidos — apenas saíram do ICOL atual.\n")

# ── 4. Gerar dados para o HTML ────────────────────────────────────────────
print("Gerando dados para o formulario...")
df_html = pd.read_excel(DEST_FILE)

fazendas_data = {}
for _, row in df_html.iterrows():
    faz  = str(row['FAZENDA']).strip()
    tal  = int(row['TALHÕES']) if pd.notna(row['TALHÕES']) else None
    layer = int(row['LAYER']) if pd.notna(row['LAYER']) else None
    cod  = int(row['COD FAZ']) if pd.notna(row['COD FAZ']) else None
    frente = int(row['FRENTE']) if pd.notna(row['FRENTE']) else None
    semana = int(row['PERÍODO OP']) if pd.notna(row.get('PERÍODO OP')) else None
    exp  = str(row['EXPORTAÇÃO']).strip() if pd.notna(row['EXPORTAÇÃO']) and str(row['EXPORTAÇÃO']) not in ['','nan'] else ''
    tipo = str(row['TIPO DE LINHA']).strip() if pd.notna(row['TIPO DE LINHA']) and str(row['TIPO DE LINHA']) not in ['','nan'] else ''
    ciclo = str(row['CICLO']).strip() if pd.notna(row['CICLO']) and str(row['CICLO']) not in ['','nan'] else ''

    if not faz or faz == 'nan' or tal is None: continue

    if faz not in fazendas_data:
        fazendas_data[faz] = {'cod': cod, 'talhoes': {}}

    tal_key = str(tal)
    if tal_key not in fazendas_data[faz]['talhoes']:
        ha = ha_lookup.get((cod, tal), 0)
        estagio_str = str(row.get('ESTÁGIO', '') or '').strip()
        fazendas_data[faz]['talhoes'][tal_key] = {
            'layer': layer, 'frente': frente, 'semana': semana,
            'exp': exp, 'tipo': tipo, 'ciclo': ciclo, 'ha': ha, 'dia': None,
            'estagio': estagio_str
        }

faz_json = json.dumps(fazendas_data, ensure_ascii=False)
n_html_faz = len(fazendas_data)
n_html_tal = sum(len(v['talhoes']) for v in fazendas_data.values())
print(f"  {n_html_faz} fazendas, {n_html_tal} talhões → HTML.\n")
print(f"  Resumo do funil de fazendas:")
print(f"    ICOL bruto                : {n_icol_raw}")
if n_dup_icol:    print(f"    (-) Duplicatas COD FAZ    : {n_dup_icol}")
if n_excluidas_20x: print(f"    (-) Administrativas (20x) : {n_excluidas_20x}")
if n_sem_base:    print(f"    (-) Sem talhão na base    : {n_sem_base}")
print(f"    (=) No dashboard          : {n_html_faz}\n")

# Ler Log Exportações → ha por usuário
print("Lendo log de exportacoes...")
usuario_ha = {}
LOG_SHEET = "Log Exportações"
try:
    df_log = pd.read_excel(DEST_FILE, sheet_name=LOG_SHEET)
    # 1 entrada por LAYER (última submissão), para não inflar ha do mesmo talhão
    df_log_dedup = df_log.drop_duplicates(subset='LAYER', keep='last')
    for _, row in df_log_dedup.iterrows():
        usuario = str(row.get('USUARIO', '') or '').strip()
        if not usuario or usuario in ('nan', ''): continue
        try:
            ly = int(float(str(row.get('LAYER', 0) or 0)))
            ha = layer_ha.get(ly, 0)
            usuario_ha[usuario] = round(usuario_ha.get(usuario, 0) + ha, 2)
        except (ValueError, TypeError):
            pass
    print(f"  {len(usuario_ha)} usuário(s) no log.\n")
except Exception as e:
    print(f"  AVISO: log nao encontrado ou vazio ({e})\n")

usuarios_json = json.dumps(usuario_ha, ensure_ascii=False)

# Ler configuração de usuários
_usuarios_config = []
if os.path.exists('usuarios.json'):
    try:
        with open('usuarios.json', 'r', encoding='utf-8') as f:
            _usuarios_config = json.load(f)
    except Exception as e:
        print(f"  AVISO: nao foi possivel ler usuarios.json ({e})\n")
usuarios_config_json = json.dumps(_usuarios_config, ensure_ascii=False)

# ── 5. Atualizar HTML ─────────────────────────────────────────────────────
print("Atualizando formulario.html...")
with open(HTML_FILE, 'r', encoding='utf-8') as f:
    html_content = f.read()

html_updated = re.sub(r'const FAZENDAS = \{.*?\};',
    f"const FAZENDAS = {faz_json};", html_content, flags=re.DOTALL)
html_updated = re.sub(r'const USUARIOS_HA = \{.*?\};',
    f"const USUARIOS_HA = {usuarios_json};", html_updated, flags=re.DOTALL)
html_updated = re.sub(r'const USUARIOS_CONFIG = \[.*?\];',
    f"const USUARIOS_CONFIG = {usuarios_config_json};", html_updated, flags=re.DOTALL)

if html_updated == html_content:
    print("  AVISO: marcador 'const FAZENDAS' nao encontrado no HTML.")
    print("  Verifique se o formulario.html e o correto.\n")
else:
    with open(HTML_FILE, 'w', encoding='utf-8') as f:
        f.write(html_updated)
    print("  formulario.html atualizado com sucesso.\n")

# ── Resumo ────────────────────────────────────────────────────────────────
print(f"{'='*50}")
print(f"  Atualizacao concluida!")
print(f"  Total talhoes : {len(result)}")
if df_base is not None:
    sem_ha = result['AREA_HA'].isna().sum()
    print(f"  Com AREA_HA   : {len(result) - sem_ha}")
    print(f"  Sem AREA_HA   : {sem_ha}")
print(f"  Preservados   : {len(result) - novos}")
print(f"  HTML regerado : sim")
print(f"{'='*50}")
input("\nPressione Enter para fechar...")
