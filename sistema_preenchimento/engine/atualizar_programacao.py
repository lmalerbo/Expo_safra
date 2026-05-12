"""
atualizar_programacao.py
Uso: python atualizar_programacao.py  OU  duplo clique no ATUALIZAR.bat

Estrutura esperada:
  base_icol/base_icol.xlsm   ← base ICOL atualizada
  programacao_frentes.xlsx
  formulario.html
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import datetime, os, sys, time, json, re

DEST_FILE   = "programacao_frentes.xlsx"
HTML_FILE   = "formulario.html"

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
        preserved[int(layer)] = (row[7] or '', row[8] or '', row[9] or '')
wb_ex.close()
print(f"  {len(preserved)} linhas existentes carregadas.\n")

# ── 2. Ler base ICOL ──────────────────────────────────────────────────────
print("Lendo base ICOL...")
df_prog = pd.read_excel(SOURCE_XLSM, sheet_name='Programacao', header=None, engine='openpyxl')
row7 = df_prog.iloc[7]
frente_groups = []
for col_idx, val in enumerate(row7):
    if str(val).strip() == 'Frente':
        frente_num = df_prog.iloc[7, col_idx + 1]
        if pd.notna(frente_num):
            frente_groups.append((col_idx, int(frente_num)))
print(f"  Frentes: {[f for _, f in frente_groups]}")

TARGET_NAMES = ['SEMANA - ANO', 'DIA', 'COD FAZ', 'FAZENDA', 'TALHOES', 'AREA_HA']
all_frames = []
for start_col, frente_num in frente_groups:
    n_cols = df_prog.shape[1]
    ha_col = start_col + 8
    if ha_col < n_cols:
        cols = [start_col + i for i in [0, 1, 2, 3, 4]] + [ha_col]
    else:
        cols = [start_col + i for i in range(5)]
        TARGET_NAMES = ['SEMANA - ANO', 'DIA', 'COD FAZ', 'FAZENDA', 'TALHOES']
    block = df_prog.iloc[9:, cols].copy()
    block.columns = TARGET_NAMES[:len(cols)]
    if 'AREA_HA' not in block.columns:
        block['AREA_HA'] = None
    block['FRENTE'] = frente_num
    block = block.dropna(subset=['COD FAZ'])
    all_frames.append(block)

result = pd.concat(all_frames, ignore_index=True)
result = result.sort_values(['FRENTE', 'DIA']).reset_index(drop=True)
result['DIA'] = pd.to_datetime(result['DIA'], errors='coerce')

def make_layer(row):
    try: return int(f"{int(row['COD FAZ'])}{int(row['TALHOES']):03d}")
    except: return None

result['LAYER'] = result.apply(make_layer, axis=1)

# Lookup de área (ha) por (COD FAZ, TALHAO) — usa primeira ocorrência não-nula
ha_lookup = {}
for _, row in result.iterrows():
    try:
        key = (int(row['COD FAZ']), int(row['TALHOES']))
        if key not in ha_lookup and pd.notna(row.get('AREA_HA')):
            ha_lookup[key] = round(float(row['AREA_HA']), 2)
    except (ValueError, TypeError):
        pass

print(f"  {len(result)} linhas lidas.\n")

# ── 3. Atualizar planilha ─────────────────────────────────────────────────
print("Atualizando planilha...")
wb = load_workbook(DEST_FILE)
ws = wb["Programação"]
ws.delete_rows(2, ws.max_row)

frente_row_colors = {2:"DDEEFF",3:"E8F2FB",4:"D6E4F5",5:"EBF3FF",
                     6:"F2F8FD",8:"DDEEFF",9:"E8F2FB",10:"D6E4F5"}
thin   = Side(style="thin", color="BFBFBF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

novos = 0
for i, row in result.iterrows():
    r      = i + 2
    frente = int(row['FRENTE'])
    cor    = frente_row_colors.get(frente, "FFFFFF")
    fill   = PatternFill("solid", start_color=cor, end_color=cor)
    layer  = row['LAYER']
    if layer in preserved:
        exp, tipo, ciclo = preserved[layer]
    else:
        exp, tipo, ciclo = '', '', ''
        novos += 1
    dia = row['DIA']
    dia = dia.date() if pd.notna(dia) else ''
    vals = [
        layer if layer else '', frente,
        int(row['SEMANA - ANO']) if pd.notna(row['SEMANA - ANO']) else '',
        dia,
        int(row['COD FAZ']) if pd.notna(row['COD FAZ']) else '',
        str(row['FAZENDA']) if pd.notna(row['FAZENDA']) else '',
        int(row['TALHOES']) if pd.notna(row['TALHOES']) else '',
        exp, tipo, ciclo,
    ]
    for c, val in enumerate(vals, 1):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font      = Font(name="Arial", size=10)
        cell.fill      = fill
        cell.border    = border
        cell.alignment = Alignment(horizontal="left" if c==6 else "center", vertical="center")
        if c == 4 and isinstance(val, datetime.date):
            cell.number_format = 'DD/MM/YYYY'

wb.save(DEST_FILE)
print(f"  Planilha atualizada. Novas linhas: {novos}\n")

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
    semana = int(row['SEMANA - ANO']) if pd.notna(row['SEMANA - ANO']) else None
    exp  = str(row['EXPORTAÇÃO']).strip() if pd.notna(row['EXPORTAÇÃO']) and str(row['EXPORTAÇÃO']) not in ['','nan'] else ''
    tipo = str(row['TIPO DE LINHA']).strip() if pd.notna(row['TIPO DE LINHA']) and str(row['TIPO DE LINHA']) not in ['','nan'] else ''
    ciclo = str(row['CICLO']).strip() if pd.notna(row['CICLO']) and str(row['CICLO']) not in ['','nan'] else ''

    if not faz or faz == 'nan' or tal is None: continue

    if faz not in fazendas_data:
        fazendas_data[faz] = {'cod': cod, 'talhoes': {}}

    tal_key = str(tal)
    if tal_key not in fazendas_data[faz]['talhoes']:
        ha = ha_lookup.get((cod, tal), 0)
        fazendas_data[faz]['talhoes'][tal_key] = {
            'layer': layer, 'frente': frente, 'semana': semana,
            'exp': exp, 'tipo': tipo, 'ciclo': ciclo, 'ha': ha
        }

faz_json = json.dumps(fazendas_data, ensure_ascii=False)
print(f"  {len(fazendas_data)} fazendas, {sum(len(v['talhoes']) for v in fazendas_data.values())} talhoes.\n")

# ── 5. Atualizar HTML ─────────────────────────────────────────────────────
print("Atualizando formulario.html...")
with open(HTML_FILE, 'r', encoding='utf-8') as f:
    html_content = f.read()

# Substitui o bloco de dados entre os marcadores
new_data_block = f"const FAZENDAS = {faz_json};"
html_updated = re.sub(
    r'const FAZENDAS = \{.*?\};',
    new_data_block,
    html_content,
    flags=re.DOTALL
)

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
print(f"  Total linhas  : {len(result)}")
print(f"  Novas         : {novos}")
print(f"  Preservadas   : {len(result) - novos}")
print(f"  HTML regerado : sim")
print(f"{'='*50}")
input("\nPressione Enter para fechar...")
