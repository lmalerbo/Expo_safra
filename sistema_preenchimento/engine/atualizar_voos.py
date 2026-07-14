"""
atualizar_voos.py
Uso: python atualizar_voos.py  OU  duplo clique no ATUALIZAR_VOOS.bat

Lê o status de voo (drone) por talhão do portal dronemgmt
(prd-dronemgmt.pedraagroindustrial.com.br/portal/flight-consult) e:

  1. Faz upsert em `voo_status` no Supabase — uma linha por voo de QUALQUER
     projeto (histórico completo, LAYER pode repetir: reagendamentos, etc).
  2. Atualiza `voo_control_status`/`voo_control_status_descricao` em `programacao` — só com
     os voos do projeto de colheita (Falha Soca), pegando o mais recente por LAYER (marcando
     "(Nx)" na descrição se houver repetição). Só atualiza LAYERs que já existem em
     `programacao` — nunca cria registro novo via esse fluxo.

Estrutura esperada:
  dronemgmt_config.json    ← { "base_url", "form_id", "unit_id", "usuario", "senha" }
                              (ver dronemgmt_config.example.json)
  supabase_config.json     ← { "url": "...", "key": "sb_publishable_..." }

O dronemgmt usa cookie de sessão (não token fixo), e o token interno por trás
dele expira rápido (minutos) mesmo que o cookie pareça válido por mais tempo —
por isso o script loga via Playwright (navegador headless) a cada execução,
em vez de depender de um cookie capturado manualmente. Requer, uma única vez:
  pip install -r engine/requirements.txt
  python -m playwright install chromium
"""

import os
import sys
import json
import unicodedata
from collections import defaultdict

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_BASE_DIR   = os.path.dirname(_SCRIPT_DIR)   # sistema_preenchimento/
sys.path.insert(0, _SCRIPT_DIR)
from utils import redirecionar_stdout, fechar_log

_log_fh = redirecionar_stdout(os.path.join(_BASE_DIR, 'logs', 'atualizar_voos.log'))

import requests
from playwright.sync_api import sync_playwright


def login_dronemgmt(base_url, usuario, senha, timeout_ms=30000):
    """Loga no portal (via Plataforma Coorporativa de Governança de Contratos,
    SSO da Pedra) usando Playwright headless e retorna (cookie_str, xsrf_token)
    prontos para as próximas chamadas via requests."""
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        try:
            context = browser.new_context()
            page = context.new_page()
            page.goto(f"{base_url}/portal/flight-consult", wait_until='networkidle', timeout=timeout_ms)
            page.get_by_placeholder("Digite seu e-mail").fill(usuario)
            page.get_by_placeholder("Digite sua senha").fill(senha)
            page.get_by_role("button", name="Entrar").click()
            # Espera o app autenticado de fato (não só a URL) — o login passa por
            # vários redirects (SSO -> callback -> app) antes do cookie ser gravado.
            page.get_by_text("Consulta Geral").wait_for(timeout=timeout_ms)
            page.wait_for_load_state('networkidle')

            dm_cookies = [c for c in context.cookies() if c['name'].startswith('DRONEMANAGEMENT-PORTAL')]
            cookie_str = '; '.join(f"{c['name']}={c['value']}" for c in dm_cookies)
            xsrf_token = next((c['value'] for c in dm_cookies if c['name'] == 'DRONEMANAGEMENT-PORTAL-XSRF-TOKEN'), None)
            return cookie_str, xsrf_token
        finally:
            browser.close()

# ── controlStatus → descrição. Preencher conforme for confirmado na tela. ──
# Códigos sem mapeamento aqui aparecem como "Status N" (não ficam em branco).
# 3/8/11 confirmados contra dados reais do projeto "Falhas Soca" (ver vooBucket() em
# formulario.html): 11 = na fila do dronemgmt, sem imagem/motivo ainda; 3/8 = imagem já
# processada mas rejeitada (motivo específico em reason_descricao, ex: "Porte Baixo",
# "Imagem não recebida"), precisa reprocessar/reagendar.
STATUS_LABELS = {
    1:  "Aguardar plantio",
    2:  "A voar",
    3:  "Imagem rejeitada (ver motivo)",
    5:  "Aguard. envio p/ processamento",
    8:  "Imagem rejeitada (ver motivo)",
    9:  "Processado, aguardando divulgação",
    10: "Relatório divulgado",
    11: "Aguardando voo",
}

# Projeto de voo relevante para a colheita (nome normalizado: sem acento, upper).
# Os demais projetos entram em voo_status mas não tocam programacao.voo_control_status.
COLHEITA_PROJETOS = {'FALHAS SOCA'}


def _fix_mojibake(s):
    """Corrige UTF-8 que foi lido como Latin-1 pelo backend do dronemgmt
    (ex: 'SistematizaÃ§Ã£o' → 'Sistematização')."""
    if not isinstance(s, str):
        return s
    try:
        return s.encode('latin-1').decode('utf-8')
    except (UnicodeDecodeError, UnicodeEncodeError):
        return s


def _norm(s):
    s = _fix_mojibake(s) or ''
    s = ''.join(c for c in unicodedata.normalize('NFKD', s) if not unicodedata.combining(c))
    return s.strip().upper()

# ── Carrega configs ─────────────────────────────────────────────────────
_dm_config_path = os.path.join(_BASE_DIR, 'dronemgmt_config.json')
if not os.path.exists(_dm_config_path):
    print(f"ERRO: Arquivo nao encontrado → {_dm_config_path}")
    print("  Copie dronemgmt_config.example.json para dronemgmt_config.json e preencha.")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)

with open(_dm_config_path, 'r', encoding='utf-8') as _f:
    _dm_cfg = json.load(_f)

DM_BASE_URL = _dm_cfg['base_url'].rstrip('/')
DM_FORM_ID  = _dm_cfg['form_id']
DM_UNIT_ID  = _dm_cfg['unit_id']
DM_USUARIO  = _dm_cfg.get('usuario', '')
DM_SENHA    = _dm_cfg.get('senha', '')

if not DM_USUARIO or not DM_SENHA:
    print("ERRO: usuario/senha vazios em dronemgmt_config.json.")
    print("  Preencha os dois campos (ver dronemgmt_config.example.json).")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)

print("Fazendo login no dronemgmt (Playwright)...")
try:
    DM_COOKIE, DM_XSRF = login_dronemgmt(DM_BASE_URL, DM_USUARIO, DM_SENHA)
except Exception as e:
    print(f"ERRO ao logar no dronemgmt: {e}")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)

if not DM_COOKIE or not DM_XSRF:
    print("ERRO: login não retornou cookie/xsrf válidos (usuário/senha incorretos?).")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)
print("  Login OK.\n")

_config_path = os.path.join(_BASE_DIR, 'supabase_config.json')
if not os.path.exists(_config_path):
    print(f"ERRO: Arquivo nao encontrado → {_config_path}")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)

with open(_config_path, 'r', encoding='utf-8') as _f:
    _sb_cfg = json.load(_f)

SUPABASE_URL = _sb_cfg['url'].rstrip('/')
SUPABASE_KEY = _sb_cfg['key']
SB_HEADERS = {
    'apikey': SUPABASE_KEY,
    'Authorization': f'Bearer {SUPABASE_KEY}',
    'Content-Type': 'application/json',
}

DM_HEADERS = {
    'Accept': '*/*',
    'Cookie': DM_COOKIE,
    'x-xsrf-token': DM_XSRF,
    'x-requested-with': 'XmlHttpRequest',
    'formid': DM_FORM_ID,
    'locale': 'pt-BR',
}

# ── 1. Busca paginada no dronemgmt ────────────────────────────────────────
print("Consultando dronemgmt (flight-consult)...")
_filter = json.dumps({"$and": [{"unitId": f"UUID('{DM_UNIT_ID}')"}, {}]})
_expand = "layer,flightProject,companyProcessing,reason"
_query_url = f"{DM_BASE_URL}/portal/api/v1/gateway/formbuilder/formdata/query"

PAGE_SIZE = 500
records = []
page = 1
total = None
while True:
    params = {
        'pageNumber': page,
        'pageSize': PAGE_SIZE,
        'filter': _filter,
        'expand': _expand,
    }
    res = requests.get(_query_url, headers=DM_HEADERS, params=params)
    if res.status_code in (401, 403):
        print(f"ERRO {res.status_code}: sessão do dronemgmt rejeitada logo após o login.")
        print("  Pode ser mudança no formid/unit_id ou no fluxo de login — verifique manualmente no navegador.")
        fechar_log(_log_fh)
        input("\nPressione Enter para sair...")
        sys.exit(1)
    if not res.ok:
        print(f"ERRO ao consultar dronemgmt: {res.status_code} {res.text[:500]}")
        fechar_log(_log_fh)
        input("\nPressione Enter para sair...")
        sys.exit(1)

    data = res.json()
    total = data.get('count', 0)
    page_records = data.get('value', [])
    records.extend(page_records)
    print(f"  página {page}: {len(page_records)} registros ({len(records)}/{total})")

    if not page_records or len(records) >= total:
        break
    page += 1

print(f"  Total obtido: {len(records)} registros.\n")

# ── 2. Monta linhas para o Supabase (voo_status: todos os projetos) ──────
print("Montando registros para o Supabase...")


def _int_or_none(v):
    try:
        return int(v)
    except (ValueError, TypeError):
        return None


voo_rows = []
for rec in records:
    layer_details = rec.get('layerDetails') or {}
    flight_project = rec.get('flightProjectDetails') or {}
    reason = rec.get('reasonDetails') or {}
    company = rec.get('companyProcessingDetails') or {}
    control_status = rec.get('controlStatus')

    voo_rows.append({
        'id':                         rec['id'],
        'layer':                      _int_or_none(layer_details.get('number')),
        'harvest':                    rec.get('harvest'),
        'section':                    rec.get('section'),
        'land_plot':                  rec.get('landPlot'),
        'flight_project_code':        flight_project.get('code'),
        'flight_project_descricao':   _fix_mojibake(flight_project.get('description')),
        'control_status':             control_status,
        'control_status_descricao':   STATUS_LABELS.get(control_status),
        'start_date_flight':          rec.get('startDateFlight'),
        'end_date_flight':            rec.get('endDateFlight'),
        'scheduled_date':             rec.get('scheduledDate'),
        'reason_descricao':           _fix_mojibake(reason.get('description')),
        'company_processing_nome':    _fix_mojibake(company.get('fantasyName')),
        'image_name':                 rec.get('imageName'),
        'image_download':             rec.get('imageDownload'),
        'date_dispatch_processing':   rec.get('dateDispatchProcessing'),
        'date_delivery_processing':   rec.get('dateDeliveryProcessing'),
        'date_disclosure_processing': rec.get('dateDisclosureProcessing'),
        'modified_utc':               rec.get('modifiedUtc'),
    })

# ── 3. Upsert no Supabase (voo_status, por id) ────────────────────────────
print(f"Enviando {len(voo_rows)} linhas para voo_status (upsert por id)...")
HEADERS_UPSERT = dict(SB_HEADERS, Prefer='resolution=merge-duplicates,return=minimal')
BATCH = 500
for i in range(0, len(voo_rows), BATCH):
    chunk = voo_rows[i:i + BATCH]
    res = requests.post(f"{SUPABASE_URL}/rest/v1/voo_status", headers=HEADERS_UPSERT, json=chunk)
    if not res.ok:
        print(f"ERRO ao enviar lote {i}-{i + len(chunk)}: {res.status_code} {res.text}")
        fechar_log(_log_fh)
        input("\nPressione Enter para sair...")
        sys.exit(1)
    print(f"  {min(i + BATCH, len(voo_rows))}/{len(voo_rows)}")

# ── 4. Agrega por LAYER (só o projeto de colheita) ────────────────────────
print("\nFiltrando projeto da colheita (Falha Soca)...")
por_layer = defaultdict(list)
for rec in records:
    flight_project = rec.get('flightProjectDetails') or {}
    if _norm(flight_project.get('description')) not in COLHEITA_PROJETOS:
        continue
    layer_val = _int_or_none((rec.get('layerDetails') or {}).get('number'))
    if layer_val is None:
        continue
    por_layer[layer_val].append(rec)

voo_updates = []
for layer_val, recs in por_layer.items():
    recs_ordenados = sorted(recs, key=lambda r: r.get('modifiedUtc') or '', reverse=True)
    mais_recente = recs_ordenados[0]
    status = mais_recente.get('controlStatus')
    texto = STATUS_LABELS.get(status, f"Status {status}")
    if len(recs_ordenados) > 1:
        texto = f"{texto} ({len(recs_ordenados)}x)"
    voo_updates.append({'layer': layer_val, 'voo_control_status': status, 'voo_control_status_descricao': texto})

print(f"  {len(voo_updates)} LAYER(s) com voo de Falha Soca.\n")

# ── 5. Atualiza programacao.voo_control_status — só para LAYERs que já existem ──
print("Verificando LAYERs existentes em programacao...")
_res_layers = requests.get(f"{SUPABASE_URL}/rest/v1/programacao?select=layer", headers=SB_HEADERS)
if not _res_layers.ok:
    print(f"ERRO ao ler layers de programacao: {_res_layers.status_code} {_res_layers.text}")
    fechar_log(_log_fh)
    input("\nPressione Enter para sair...")
    sys.exit(1)
existing_layers = {row['layer'] for row in _res_layers.json()}

voo_updates_validos = [r for r in voo_updates if r['layer'] in existing_layers]
n_fora = len(voo_updates) - len(voo_updates_validos)
if n_fora:
    print(f"  ⚠  {n_fora} LAYER(s) com voo na API mas fora da programação atual — ignorados (não cria registro novo).")

print(f"\nAtualizando programacao.voo_control_status para {len(voo_updates_validos)} LAYER(s)...")
for i in range(0, len(voo_updates_validos), BATCH):
    chunk = voo_updates_validos[i:i + BATCH]
    res = requests.post(f"{SUPABASE_URL}/rest/v1/programacao", headers=HEADERS_UPSERT, json=chunk)
    if not res.ok:
        print(f"ERRO ao atualizar voo (lote {i}-{i + len(chunk)}): {res.status_code} {res.text}")
        fechar_log(_log_fh)
        input("\nPressione Enter para sair...")
        sys.exit(1)
    print(f"  {min(i + BATCH, len(voo_updates_validos))}/{len(voo_updates_validos)}")

print(f"\n{'='*50}")
print("  Atualizacao de voos concluida!")
print(f"  Total em voo_status              : {len(voo_rows)}")
print(f"  LAYERs com voo de Falha Soca      : {len(voo_updates)}")
print(f"  programacao.voo_control_status   : {len(voo_updates_validos)}")
print(f"{'='*50}")

fechar_log(_log_fh)
input("\nPressione Enter para fechar...")
