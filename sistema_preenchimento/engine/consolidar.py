"""
consolidar.py — Consolida os exports diários na planilha mestre

Uso:
  1. Coloque este script na mesma pasta que programacao_frentes.xlsm
  2. Crie uma subpasta chamada  exports/
  3. Mova todos os arquivos export_frentes_*.xlsx para essa pasta
  4. Execute: python consolidar.py

O script:
  - Lê todos os exports da pasta exports/
  - Para cada registro, sobrescreve as colunas EXPORTAÇÃO, TIPO DE LINHA, CICLO
    na linha correspondente ao LAYER na planilha mestre
  - Move os arquivos processados para exports/processados/
  - Gera um relatório resumido no terminal
"""

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import datetime
import os
import sys
import shutil
import glob

MESTRE    = "programacao_frentes.xlsx"
PASTA_EXP = "exports"
PASTA_OK  = os.path.join(PASTA_EXP, "processados")

# ── Verificações iniciais ─────────────────────────────────────────────────
for f in [MESTRE]:
    if not os.path.exists(f):
        print(f"ERRO: Arquivo nao encontrado → {f}")
        input("\nPressione Enter para sair...")
        sys.exit(1)

os.makedirs(PASTA_OK, exist_ok=True)

# ── Verifica se mestre está aberto ────────────────────────────────────────
def arquivo_bloqueado(path):
    try:
        with open(path, 'a+b'): return False
    except (IOError, PermissionError): return True

import time
tentativas = 0
while arquivo_bloqueado(MESTRE):
    tentativas += 1
    if tentativas >= 3:
        print(f"ERRO: '{MESTRE}' está aberto. Feche o arquivo e tente novamente.")
        input("\nPressione Enter para sair...")
        sys.exit(1)
    print(f"  Arquivo em uso, aguardando 10s... ({tentativas}/3)")
    time.sleep(10)

# ── Carrega exports ───────────────────────────────────────────────────────
exports = glob.glob(os.path.join(PASTA_EXP, "export_frentes_*.xlsx"))
if not exports:
    print(f"Nenhum arquivo export_frentes_*.xlsx encontrado em '{PASTA_EXP}/'")
    input("\nPressione Enter para sair...")
    sys.exit(0)

print(f"Encontrados {len(exports)} arquivo(s) de export:\n")
all_records = []
for path in exports:
    try:
        df = pd.read_excel(path)
        df['_source'] = os.path.basename(path)
        all_records.append(df)
        print(f"  ✓ {os.path.basename(path)} — {len(df)} registros")
    except Exception as e:
        print(f"  ✗ {os.path.basename(path)} — ERRO: {e}")

if not all_records:
    print("\nNenhum registro válido encontrado.")
    input("\nPressione Enter para sair...")
    sys.exit(0)

df_all = pd.concat(all_records, ignore_index=True)
print(f"\nTotal de registros para processar: {len(df_all)}")

# Em caso de LAYER duplicado nos exports, mantém o mais recente (última linha)
df_all['LAYER'] = df_all['LAYER'].astype(str).str.strip()
df_dedup = df_all.drop_duplicates(subset='LAYER', keep='last')
print(f"Após deduplicação por LAYER: {len(df_dedup)} registros únicos")

# Monta lookup layer → valores
lookup = df_dedup.set_index('LAYER')[['STATUS','TIPO_LINHA','CICLO','USUARIO','TIMESTAMP']].to_dict('index')

# ── Verifica conflitos (layers que já têm preenchimento na mestre) ─────────
print("\nVerificando conflitos na planilha mestre...")
wb = load_workbook(MESTRE)
ws = wb["Programação"]
lastRow = ws.max_row

conflitos = []   # (row, layer, val_mestre, val_export)
limpos    = []   # (row, layer) sem preenchimento atual

for row in range(2, lastRow + 1):
    layer_cell = ws.cell(row=row, column=1).value
    if layer_cell is None:
        continue
    layer_str = str(int(layer_cell)) if isinstance(layer_cell, float) else str(layer_cell).strip()
    if layer_str not in lookup:
        continue

    status_atual = ws.cell(row=row, column=8).value or ''
    tipo_atual   = ws.cell(row=row, column=9).value or ''
    ciclo_atual  = ws.cell(row=row, column=10).value or ''

    ja_preenchido = any([str(status_atual).strip(), str(tipo_atual).strip(), str(ciclo_atual).strip()])
    rec = lookup[layer_str]

    if ja_preenchido:
        conflitos.append({
            'row': row, 'layer': layer_str,
            'mestre': f"{status_atual} | {tipo_atual} | {ciclo_atual}",
            'export': f"{rec['STATUS']} | {rec['TIPO_LINHA']} | {rec['CICLO']}",
            'usuario': rec['USUARIO'], 'ts': rec['TIMESTAMP']
        })
    else:
        limpos.append({'row': row, 'layer': layer_str})

print(f"  Sem conflito (serão preenchidos): {len(limpos)}")
print(f"  Com conflito (já preenchidos)   : {len(conflitos)}")

# ── Resolve conflitos interativamente ─────────────────────────────────────
sobrescrever_ids = set()   # layers aprovados para sobrescrever

if conflitos:
    print(f"\n{'='*60}")
    print(f"  ATENÇÃO: {len(conflitos)} talhão(ões) já têm preenchimento.")
    print(f"{'='*60}")
    print("\nO que deseja fazer?\n")
    print("  1 - Ver cada conflito e decidir um por um")
    print("  2 - Sobrescrever TODOS automaticamente")
    print("  3 - Manter TODOS (não sobrescreve nenhum conflito)")
    print()
    escolha = input("Digite 1, 2 ou 3: ").strip()

    if escolha == '2':
        sobrescrever_ids = {c['layer'] for c in conflitos}
        print(f"\n  → Todos os {len(conflitos)} conflitos serão sobrescritos.")

    elif escolha == '1':
        print()
        for i, c in enumerate(conflitos, 1):
            print(f"  [{i}/{len(conflitos)}] LAYER {c['layer']}")
            print(f"    Mestre : {c['mestre']}")
            print(f"    Export : {c['export']}  (por {c['usuario']} em {c['ts']})")
            resp = input("    Sobrescrever? (s/N): ").strip().lower()
            if resp == 's':
                sobrescrever_ids.add(c['layer'])
        print(f"\n  → {len(sobrescrever_ids)} de {len(conflitos)} conflitos serão sobrescritos.")

    else:  # '3' ou qualquer outra coisa
        print(f"\n  → Nenhum conflito será sobrescrito. Apenas novos registros serão aplicados.")

# ── Aplica atualizações ───────────────────────────────────────────────────
print("\nAplicando atualizações...")
atualizados = 0
ignorados   = 0
nao_encontrados = []

for row in range(2, lastRow + 1):
    layer_cell = ws.cell(row=row, column=1).value
    if layer_cell is None:
        continue
    layer_str = str(int(layer_cell)) if isinstance(layer_cell, float) else str(layer_cell).strip()
    if layer_str not in lookup:
        continue

    # Verifica se é conflito e se foi aprovado
    is_conflito = any(c['layer'] == layer_str for c in conflitos)
    if is_conflito and layer_str not in sobrescrever_ids:
        ignorados += 1
        continue

    rec = lookup[layer_str]
    ws.cell(row=row, column=8).value  = str(rec['STATUS'])
    ws.cell(row=row, column=9).value  = str(rec['TIPO_LINHA'])
    ws.cell(row=row, column=10).value = str(rec['CICLO'])
    atualizados += 1

# Verifica layers não encontrados na mestre
layers_mestre = set()
for row in range(2, lastRow + 1):
    v = ws.cell(row=row, column=1).value
    if v:
        layers_mestre.add(str(int(v)) if isinstance(v, float) else str(v).strip())

for layer in lookup:
    if layer not in layers_mestre:
        nao_encontrados.append(layer)

# ── Aba de log histórico ──────────────────────────────────────────────────
LOG_SHEET = "Log Exportações"
if LOG_SHEET not in wb.sheetnames:
    ws_log = wb.create_sheet(LOG_SHEET)
    # Headers
    headers = ['DATA_CONSOLIDACAO','TIMESTAMP','USUARIO','LAYER','FAZENDA','TALHAO','TIPO_LINHA','CICLO','STATUS','ARQUIVO']
    thin = Side(style="thin", color="BFBFBF")
    brd  = Border(left=thin, right=thin, top=thin, bottom=thin)
    azul = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    for c, h in enumerate(headers, 1):
        cell = ws_log.cell(row=1, column=c, value=h)
        cell.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell.fill = azul
        cell.border = brd
        cell.alignment = Alignment(horizontal="center")
else:
    ws_log = wb[LOG_SHEET]

# Append records to log
last_log = ws_log.max_row
now_str  = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')
for _, rec in df_all.iterrows():
    last_log += 1
    ws_log.cell(row=last_log, column=1).value  = now_str
    ws_log.cell(row=last_log, column=2).value  = str(rec.get('TIMESTAMP',''))
    ws_log.cell(row=last_log, column=3).value  = str(rec.get('USUARIO',''))
    ws_log.cell(row=last_log, column=4).value  = str(rec.get('LAYER',''))
    ws_log.cell(row=last_log, column=5).value  = str(rec.get('FAZENDA',''))
    ws_log.cell(row=last_log, column=6).value  = str(rec.get('TALHAO',''))
    ws_log.cell(row=last_log, column=7).value  = str(rec.get('TIPO_LINHA',''))
    ws_log.cell(row=last_log, column=8).value  = str(rec.get('CICLO',''))
    ws_log.cell(row=last_log, column=9).value  = str(rec.get('STATUS',''))
    ws_log.cell(row=last_log, column=10).value = str(rec.get('_source',''))
    for c in range(1, 11):
        ws_log.cell(row=last_log, column=c).font = Font(name="Arial", size=10)

wb.save(MESTRE)

# ── Move exports processados ──────────────────────────────────────────────
for path in exports:
    shutil.move(path, os.path.join(PASTA_OK, os.path.basename(path)))

# ── Relatório final ───────────────────────────────────────────────────────
print(f"\n{'='*50}")
print(f"  Consolidação concluída!")
print(f"  Registros processados : {len(df_all)}")
print(f"  Linhas atualizadas    : {atualizados}")
print(f"  Conflitos ignorados   : {ignorados}")
print(f"  Arquivos movidos para : {PASTA_OK}/")
if nao_encontrados:
    print(f"\n  ⚠  {len(nao_encontrados)} LAYER(s) não encontrados na mestre:")
    for l in nao_encontrados[:10]:
        print(f"     - {l}")
    if len(nao_encontrados) > 10:
        print(f"     ... e mais {len(nao_encontrados)-10}")
print(f"  Log salvo na aba '{LOG_SHEET}'")
print(f"{'='*50}")
input("\nPressione Enter para fechar...")
