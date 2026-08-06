// Proxy para upload de arquivos de projeto (.dwg/.zip) nas Releases do GitHub, e para os
// relatórios somente-leitura consumidos fora do app (Power BI, Sheets, scripts).
//
// Motivo do upload: o formulario.html é estático e público (GitHub Pages) — qualquer token
// do GitHub embutido nele é detectado e revogado automaticamente pelo secret scanning. Este
// Worker guarda o token como secret do Cloudflare (nunca commitado).
//
// Endpoints:
//   POST /upload?tag=&name=&filename=   (body = arquivo)  → usado no formulario.html
//   POST /trigger-voos                                     → dispara o workflow atualizar-voos.yml
//   GET  /voos-status                                      → status da última execução desse workflow
//   GET  /report/fazendas                                  → status consolidado por fazenda (relatórios)
//   GET  /report/talhoes                                   → dados de programação por talhão (relatórios)
//
// Os dois endpoints /report/* são somente leitura e usam a mesma chave anon pública já
// embutida no formulario.html — não expõem nada que o sistema não exponha hoje. CORS
// liberado (Access-Control-Allow-Origin: *) porque são consumidos por scripts/Power BI/
// Sheets fora do GitHub Pages, não só pelo formulario.html.
//
// Deploy (via dashboard do Cloudflare):
//   1. Workers & Pages → Create → Create Worker → cole este arquivo.
//   2. Settings → Variables → Add secret: GH_TOKEN = <PAT com permissão "Contents" read/write
//      E "Actions" read/write no repo lmalerbo/Expo_safra — o escopo Actions é necessário para
//      as rotas /trigger-voos e /voos-status (disparo/consulta do workflow atualizar-voos.yml)>.
//   3. Anote a URL do worker (https://<nome>.<conta>.workers.dev) e configure
//      RELEASE_PROXY_URL no formulario.html com esse valor.

const GH_OWNER = 'lmalerbo';
const GH_REPO = 'Expo_safra';
const ALLOWED_ORIGIN = 'https://lmalerbo.github.io';

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': ALLOWED_ORIGIN,
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
}

// CORS aberto, só para os endpoints /report/* (dados públicos, leitura, consumidos fora do
// GitHub Pages — Power BI, Sheets, scripts).
function corsHeadersOpen() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
}

// ── Relatórios (leitura direta do Supabase, mesma chave anon do formulario.html) ──
const SUPABASE_URL = 'https://wewicqysphguehqnyjdh.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_qAKXJVaKSMeD6oUvLoQ7EA_ae5t6XBQ';

// PostgREST limita a 1000 linhas por requisição — busca todas as páginas (igual
// _sbFetchAll em formulario.html; programacao já passa de 4500 linhas).
async function sbFetchAll(path) {
  const PAGE = 1000;
  let offset = 0;
  let all = [];
  while (true) {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
      headers: {
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${SUPABASE_ANON_KEY}`,
        'Range-Unit': 'items',
        Range: `${offset}-${offset + PAGE - 1}`,
      },
    });
    if (!res.ok) throw new Error(`supabase ${path}: ${res.status} ${await res.text()}`);
    const page = await res.json();
    all = all.concat(page);
    if (page.length < PAGE) break;
    offset += PAGE;
  }
  return all;
}

// Status por talhão — mesma lógica de statusEfetivo() em formulario.html: CONCLUIDO só
// conta quando há um usuário por trás (log_exportacoes); sem isso vira REAVALIAR.
function statusEfetivo(t) {
  if (t.exp === 'SEM LINHAS' || t.tipo === 'SEM LINHAS') return 'SEM LINHAS';
  if (t.exp === 'CONCLUIDO') return t.usuario ? 'CONCLUIDO' : 'REAVALIAR';
  return 'A FAZER';
}
function isTalDone(t) {
  const st = statusEfetivo(t);
  return st === 'CONCLUIDO' || st === 'SEM LINHAS';
}
function isTalReavaliar(t) {
  return statusEfetivo(t) === 'REAVALIAR';
}
// Status agregado por fazenda — mesma prioridade de statusAgregado() em formulario.html:
// A FAZER > REAVALIAR > CONCLUÍDO.
function statusAgregado(tals) {
  if (!tals.length) return '—';
  if (tals.some(t => statusEfetivo(t) === 'A FAZER')) return 'A FAZER';
  if (tals.some(isTalReavaliar)) return 'REAVALIAR';
  return 'CONCLUÍDO';
}

// Tag de porte (dias corridos desde o corte) — mesma lógica de tagPorte() em
// formulario.html. Safra atual: 03/2026 a 03/2027 (fim exclusivo).
const SAFRA_INICIO = new Date(2026, 2, 1);
const SAFRA_FIM = new Date(2027, 2, 1);
function tagPorte(t) {
  if (statusEfetivo(t) !== 'SEM LINHAS') return null;
  if (String(t.estagio || '').toUpperCase().trim() === 'A PLANTAR') return null;
  if (!t.data_corte) return null;
  const dataCorte = new Date(t.data_corte);
  if (isNaN(dataCorte) || dataCorte < SAFRA_INICIO || dataCorte >= SAFRA_FIM) return null;
  return Math.floor((Date.now() - dataCorte.getTime()) / 86400000);
}

// Bucket de cor da tag de porte, a partir do status de voo do dronemgmt (projeto "Falhas
// Soca") — mesma lógica de vooBucket() em formulario.html: control_status 3/8 = imagem já
// processada (rejeitada/reprocessar) = "voado"; 10 = "Relatório divulgado" = "divulgado";
// verify_flight_size === 5 ("Voo liberado") = "agendado"; senão "sem_agendamento".
function vooBucket(t) {
  const cod = t.voo_control_status;
  if (cod == null && t.voo_verify_flight_size == null) return 'sem_agendamento';
  if (cod === 10) return 'divulgado';
  if (cod === 3 || cod === 8) return 'voado';
  if (t.voo_verify_flight_size === 5) return 'agendado';
  return 'sem_agendamento';
}

async function buildRelatorioFazendas() {
  const [programacao, logs] = await Promise.all([
    sbFetchAll('programacao?select=*'),
    sbFetchAll('log_exportacoes?select=usuario,layer,data_consolidacao&order=data_consolidacao.asc'),
  ]);

  // Último usuário por LAYER (logs vêm em ordem ascendente — o último sobrescreve).
  const usuarioPorLayer = new Map();
  for (const row of logs) usuarioPorLayer.set(String(row.layer), row.usuario || '');

  // Agrupa por cod_faz (não por nome — duas fazendas distintas podem ter o mesmo nome).
  const porFazenda = new Map();
  for (const row of programacao) {
    if (!row.fazenda) continue;
    if (!porFazenda.has(row.cod_faz)) {
      porFazenda.set(row.cod_faz, { cod_faz: row.cod_faz, fazenda: row.fazenda, tals: [] });
    }
    porFazenda.get(row.cod_faz).tals.push({
      exp: row.status || '', tipo: row.tipo_linha || '', ciclo: row.ciclo || '',
      ha: row.area_ha || 0, frente: row.frente, estagio: row.estagio || '',
      data_corte: row.data_corte || null,
      voo_control_status: row.voo_control_status ?? null,
      voo_verify_flight_size: row.voo_verify_flight_size ?? null,
      usuario: usuarioPorLayer.get(String(row.layer)) || '',
    });
  }

  return [...porFazenda.values()].map(f => {
    const frentes = [...new Set(f.tals.map(t => t.frente).filter(Boolean))].sort((a, b) => a - b).join(', ');
    let porteMax = null, porteBucket = null;
    for (const t of f.tals) {
      const dias = tagPorte(t);
      if (dias !== null && (porteMax === null || dias > porteMax)) { porteMax = dias; porteBucket = vooBucket(t); }
    }
    return {
      cod_faz: f.cod_faz,
      fazenda: f.fazenda,
      status: statusAgregado(f.tals),
      frentes,
      ciclo: f.tals.find(t => t.ciclo)?.ciclo || '',
      area_ha: Math.round(f.tals.reduce((s, t) => s + (t.ha || 0), 0) * 100) / 100,
      talhoes: f.tals.length,
      talhoes_concluidos: f.tals.filter(isTalDone).length,
      porte_max_dias: porteMax,
      porte_bucket: porteBucket,
    };
  }).sort((a, b) => a.fazenda.localeCompare(b.fazenda));
}

async function buildRelatorioTalhoes() {
  const [programacao, logs] = await Promise.all([
    sbFetchAll('programacao?select=*'),
    sbFetchAll('log_exportacoes?select=usuario,layer,data_consolidacao&order=data_consolidacao.asc'),
  ]);

  const usuarioPorLayer = new Map();
  const dataConsPorLayer = new Map();
  for (const row of logs) {
    usuarioPorLayer.set(String(row.layer), row.usuario || '');
    dataConsPorLayer.set(String(row.layer), row.data_consolidacao || null);
  }

  return programacao.map(row => {
    const t = {
      exp: row.status || '', tipo: row.tipo_linha || '', estagio: row.estagio || '',
      data_corte: row.data_corte || null,
      voo_control_status: row.voo_control_status ?? null,
      voo_verify_flight_size: row.voo_verify_flight_size ?? null,
      usuario: usuarioPorLayer.get(String(row.layer)) || '',
    };
    return {
      layer: row.layer, cod_faz: row.cod_faz, fazenda: row.fazenda, talhao: row.talhao,
      frente: row.frente, periodo_op: row.periodo_op, area_ha: row.area_ha || 0,
      ciclo: row.ciclo || '', tipo_linha: row.tipo_linha || '',
      status: statusEfetivo(t), estagio: row.estagio || '', data_corte: row.data_corte || null,
      porte_dias: tagPorte(t),
      voo_control_status: row.voo_control_status ?? null,
      voo_verify_flight_size: row.voo_verify_flight_size ?? null,
      voo_bucket: vooBucket(t),
      usuario: t.usuario || null,
      data_consolidacao: dataConsPorLayer.get(String(row.layer)) || null,
    };
  }).sort((a, b) => (a.fazenda || '').localeCompare(b.fazenda || '') || (a.talhao || 0) - (b.talhao || 0));
}

function ghHeaders(env, extra) {
  return Object.assign({
    'Authorization': `Bearer ${env.GH_TOKEN}`,
    'Accept': 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28',
    'User-Agent': 'project-colheita-release-proxy',
  }, extra || {});
}

async function getOrCreateRelease(env, tag, name) {
  let res = await fetch(`https://api.github.com/repos/${GH_OWNER}/${GH_REPO}/releases/tags/${encodeURIComponent(tag)}`,
    { headers: ghHeaders(env) });
  if (res.status === 404) {
    res = await fetch(`https://api.github.com/repos/${GH_OWNER}/${GH_REPO}/releases`, {
      method: 'POST',
      headers: ghHeaders(env, { 'Content-Type': 'application/json' }),
      body: JSON.stringify({ tag_name: tag, name, target_commitish: 'main' }),
    });
  }
  if (!res.ok) throw new Error(`release ${tag}: ${res.status} ${await res.text()}`);
  return res.json();
}

// O GitHub substitui espaço por ponto no nome do asset ao salvá-lo (ex: "SANTA RITA 8"
// fica "SANTA.RITA.8") — sem normalizar aqui, a busca pelo asset existente nunca encontra
// nada (sempre teria espaço onde o GitHub já tem ponto), o delete nunca roda, e o upload
// novo colide com "already_exists" porque o GitHub aplica essa mesma normalização nele.
function nomeComoGitHubSalva(filename) {
  return filename.replace(/ /g, '.');
}

// Apaga o asset existente com esse nome, se houver. Não basta disparar o DELETE e seguir —
// se ele falhar (permissão, asset já removido por outra causa) o upload seguinte colide
// com "already_exists" sem nenhuma pista do motivo real. Retorna um diagnóstico (achou? em
// qual id? o delete deu qual status?) pra poder ser anexado na mensagem de erro final.
async function deleteAssetIfExists(env, release, filename) {
  const normalizado = nomeComoGitHubSalva(filename);
  const existente = (release.assets || []).find(a => a.name === filename || a.name === normalizado);
  if (!existente) return { found: false };
  const res = await fetch(`https://api.github.com/repos/${GH_OWNER}/${GH_REPO}/releases/assets/${existente.id}`,
    { method: 'DELETE', headers: ghHeaders(env) });
  return { found: true, id: existente.id, deleteStatus: res.status, deleteOk: res.ok || res.status === 404 };
}

async function uploadAsset(env, release, filename, contentType, body) {
  const uploadUrl = release.upload_url.replace('{?name,label}', '') + `?name=${encodeURIComponent(filename)}`;
  return fetch(uploadUrl, {
    method: 'POST',
    headers: ghHeaders(env, { 'Content-Type': contentType || 'application/octet-stream' }),
    body,
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const cors = url.pathname.startsWith('/report/') ? corsHeadersOpen() : corsHeaders();

    if (request.method === 'OPTIONS') {
      return new Response(null, { headers: cors });
    }

    try {
      if (url.pathname === '/_version') {
        return new Response('release-proxy v5 (relatorios /report/*)', { headers: cors });
      }

      if (url.pathname === '/report/fazendas' && request.method === 'GET') {
        const data = await buildRelatorioFazendas();
        return new Response(JSON.stringify(data), { headers: { ...cors, 'Content-Type': 'application/json' } });
      }

      if (url.pathname === '/report/talhoes' && request.method === 'GET') {
        const data = await buildRelatorioTalhoes();
        return new Response(JSON.stringify(data), { headers: { ...cors, 'Content-Type': 'application/json' } });
      }

      if (url.pathname === '/upload' && request.method === 'POST') {
        const tag = url.searchParams.get('tag');
        const name = url.searchParams.get('name') || tag;
        const filename = url.searchParams.get('filename');
        if (!tag || !filename) {
          return new Response('tag e filename são obrigatórios', { status: 400, headers: corsHeaders() });
        }

        const contentType = request.headers.get('Content-Type') || 'application/octet-stream';
        const body = await request.arrayBuffer();

        let release = await getOrCreateRelease(env, tag, name);
        let del1 = await deleteAssetIfExists(env, release, filename);

        let res = await uploadAsset(env, release, filename, contentType, body);

        // Se ainda colidir com "already_exists" (delete não propagou a tempo do lado do
        // GitHub), busca o release de novo, tenta apagar uma segunda vez e reenvia uma
        // única vez antes de desistir — em vez de falhar direto na primeira corrida.
        let del2 = null;
        if (res.status === 422 && /already_exists/.test(await res.clone().text())) {
          release = await getOrCreateRelease(env, tag, name);
          del2 = await deleteAssetIfExists(env, release, filename);
          res = await uploadAsset(env, release, filename, contentType, body);
        }
        if (!res.ok) {
          const diag = `delete1=${JSON.stringify(del1)} delete2=${JSON.stringify(del2)}`;
          throw new Error(`upload ${filename}: ${res.status} ${await res.text()} | ${diag}`);
        }

        return new Response(await res.text(), { headers: { ...corsHeaders(), 'Content-Type': 'application/json' } });
      }

      if (url.pathname === '/trigger-voos' && request.method === 'POST') {
        const res = await fetch(
          `https://api.github.com/repos/${GH_OWNER}/${GH_REPO}/actions/workflows/atualizar-voos.yml/dispatches`,
          { method: 'POST', headers: ghHeaders(env, { 'Content-Type': 'application/json' }), body: JSON.stringify({ ref: 'main' }) }
        );
        if (!res.ok) return new Response(`erro ao disparar: ${res.status} ${await res.text()}`, { status: res.status, headers: corsHeaders() });
        return new Response(JSON.stringify({ ok: true }), { headers: { ...corsHeaders(), 'Content-Type': 'application/json' } });
      }

      if (url.pathname === '/voos-status' && request.method === 'GET') {
        const res = await fetch(
          `https://api.github.com/repos/${GH_OWNER}/${GH_REPO}/actions/workflows/atualizar-voos.yml/runs?per_page=1`,
          { headers: ghHeaders(env) }
        );
        if (!res.ok) return new Response(`erro ao consultar status: ${res.status} ${await res.text()}`, { status: res.status, headers: corsHeaders() });
        return new Response(await res.text(), { headers: { ...corsHeaders(), 'Content-Type': 'application/json' } });
      }

      return new Response('Not found', { status: 404, headers: corsHeaders() });
    } catch (e) {
      return new Response(JSON.stringify({ error: e.message }), {
        status: 500, headers: { ...corsHeaders(), 'Content-Type': 'application/json' },
      });
    }
  },
};
