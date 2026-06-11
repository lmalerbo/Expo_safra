# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Current architecture (Supabase)

`formulario.html` is now a fully online, self-contained HTML file: it reads and writes the "database"
directly via the **Supabase REST API** (PostgREST), using a public `sb_publishable_...` key embedded in
the file (security comes from RLS policies, not from hiding the key — see `supabase/migrations/0001_init.sql`).

- `SUPABASE_URL` / `SUPABASE_KEY` consts near `RELEASE_PROXY_URL`.
- `carregarDadosSupabase()` (called from `window.onload`, which is now `async`) fetches `programacao`,
  `log_exportacoes` and `usuarios` and rebuilds `FAZENDAS` / `USUARIOS_HA` / `USUARIOS_CONFIG` in place.
  The values embedded as `const FAZENDAS = {...}` etc. (still injected by `atualizar_html()`, see legacy
  section below) are kept only as a **seed/fallback** in case the Supabase request fails.
- `consolidar()` writes directly to Supabase: `PATCH /rest/v1/programacao` (per LAYER), bulk
  `POST /rest/v1/log_exportacoes`, and `PATCH /rest/v1/usuarios` to update the user's accumulated `ha`.
  No file queue, no local Python process required.
- Tables and RLS policies: `supabase/migrations/0001_init.sql` (+ `0002_programacao_insert_anon.sql`,
  which lets `anon` `INSERT` into `programacao` so the browser-based base update can upsert new LAYERs).
  Apply with `supabase link --project-ref wewicqysphguehqnyjdh` + `supabase db push`.
- One-time data migration from the old `programacao_frentes.xlsx`: `engine/migrar_supabase.py` (needs
  `supabase_config.json`, gitignored, with `{ "url": ..., "secret_key": "sb_secret_..." }`).
- Refreshing the ICOL base — **primary path**: in `formulario.html`, "⚙ Gerenciar" → "Atualizar base
  ICOL / Fazendas" lets the user upload the `.xlsm` (ICOL, sheet `BASE PARA PLANEJAMENTO`) and `.xlsx`
  (base de fazendas) files directly in the browser. `processarAtualizacaoBase()` parses both with SheetJS
  (already loaded for `_exportarXLSX`), replicates the merge/preserve logic of
  `engine/atualizar_programacao.py` in JS, and upserts straight into `programacao` via the publishable
  key (needs the `programacao_insert_anon` policy from `supabase/migrations/0002_*.sql`). "ⓘ" buttons
  next to each file input explain the expected layout and offer a downloadable example via
  `_baixarModeloICOL()` / `_baixarModeloBaseFazendas()`.
- Refreshing the ICOL base — **legacy/local fallback**: `engine/atualizar_programacao.py` reads
  `base_icol/*.xlsm` + `base_fazendas/*.xlsx`, preserves existing STATUS/TIPO_LINHA/CICLO per LAYER, and
  upserts into the Supabase `programacao` table (needs Python + `supabase_config.json` with the secret
  key). Run via `ATUALIZAR.bat`.
- File upload (`.dwg`/`.zip`) to GitHub Releases at the end of consolidation: `enviarArquivosProjeto()`
  calls `_releaseUploadAsset()`, which posts to a **Cloudflare Worker proxy** (`RELEASE_PROXY_URL`,
  `cloudflare-worker/release-proxy.js`). The Worker holds the GitHub PAT as a Cloudflare secret (never
  committed) and creates/updates the release + uploads the asset. This indirection exists because
  `formulario.html` is public (GitHub Pages) and any `github_pat_...` embedded in it gets auto-revoked
  by GitHub's secret scanning as soon as it's pushed.

Run a single test:
```bat
python -m pytest tests/test_utils.py -v
```

## Legacy: file-queue architecture (`vigia.py` / Excel)

> Kept for now as a fallback/reference. Not part of the daily flow anymore — `formulario.html` no longer
> writes to `consolidar_queue/` or depends on `vigia.py` being online.

```bat
# Start the background watcher (runs 24h on the server PC)
INICIAR_VIGIA.bat         # launches vigia.py via pythonw (no console window)

# Manually trigger consolidation (outside the browser flow)
CONSOLIDAR.bat            # runs engine/consolidar.py interactively

# Update farm/user data embedded in formulario.html (legacy const seed)
ATUALIZAR.bat             # runs the legacy engine/atualizar_programacao.py flow
```

The legacy system has two loosely-coupled halves that communicate only through JSON files on a shared G: drive.

### Browser half (`formulario.html`, legacy bits)

- Old "Consolidar" flow wrote `consolidar_queue/req_{id}.json`, polled for `res_{id}.json`, then deleted it.
- Static data (`FAZENDAS`, `USUARIOS_HA`, `USUARIOS_CONFIG`) embedded directly in the HTML as JS values —
  injected/updated by `engine/utils.py:atualizar_html()` using line-prefix markers. Now only a fallback seed.
- File copy feature: uses File System Access API to scan `FAZENDAS\` for `.dwg`/`.zip` and copy to SharePoint folders.

IndexedDB (`pf-config` database, `handles` store) persists `FileSystemDirectoryHandle` objects for:

- `filesRootDir` — FAZENDAS root for project file scan
- `filesDwgDir` — SharePoint destination for `.dwg`
- `filesExpDir` — SharePoint destination for `.zip`

### Python half (`vigia.py` + `engine/`)

`vigia.py` polls `consolidar_queue/` every 2s (configurable). On finding `req_{id}.json`:

1. Saves records to `exports/export_frentes_{user}_{date}_{id}.xlsx`
2. Runs `engine/consolidar.py` (which updated `programacao_frentes.xlsx` and regenerated `formulario.html`)
3. Writes `res_{id}.json` with `{ok, arquivo, output}`

`engine/consolidar.py` — matches export rows to the master spreadsheet by LAYER string, overwrites STATUS/TIPO_LINHA/CICLO columns, moves processed exports to `exports/processados/`, then calls `atualizar_html()`.

`engine/utils.py` — shared helpers: `layer_to_str()`, `arquivo_bloqueado()`, `aguardar_arquivo_livre()`, `atualizar_html()`, stdout tee logging. `layer_to_str()` and `redirecionar_stdout`/`fechar_log` are still used by the Supabase scripts.

### Configuration
`config.json` — tuning knobs: poll interval, consolidation timeout, retry settings, column index map for the master spreadsheet (still used by the legacy `consolidar.py`).

`usuarios.json` — seed user list with `preenchimento`/`dashboard` profiles, used only by `engine/migrar_supabase.py` for the initial migration. The live source of truth is the Supabase `usuarios` table.

## TODO / Débito técnico conhecido

- ~~`GH_TOKEN` exposto em `formulario.html`~~ — **resolvido**: o repo é público (GitHub Pages), e
  qualquer `github_pat_...` embutido no HTML é detectado e revogado automaticamente pelo GitHub
  (secret scanning), independente de "allow" no push protection. A solução foi mover o upload de
  arquivos para um proxy (Cloudflare Worker, `cloudflare-worker/release-proxy.js`) que guarda o token
  como secret do Cloudflare — `formulario.html` só conhece a URL pública do Worker
  (`RELEASE_PROXY_URL`), sem nenhuma credencial.

## Key conventions

**File naming for project files:**
- Farms folder: `{COD_FAZ} {NOME_FAZENDA}` (space-separated, e.g. `10503 SANTA LUZIA 5`)
- Project files: `{COD_FAZ}_{NOME_FAZENDA}_Exp{N}L.{dwg|zip}` (underscore-separated, e.g. `10503_SANTA LUZIA 5_Exp1L.dwg`)
- Regex used in `_varrerFazendas`: `/^(\d+)_.+_Exp\d+L\.(dwg|zip)$/i`
- COD_FAZ extracted from file name: `name.split('_')[0]`
- COD_FAZ extracted from folder name: `name.split(' ')[0]`

**LAYER key:** always a string of digits without decimal point — use `layer_to_str()` whenever converting from spreadsheet values (handles `1001005.0` → `"1001005"`).

**`codfaz_excluir_prefixo`** in `config.json` (value `"20"`) — COD_FAZ values starting with this prefix are excluded from consolidation (test/administrative farms).

**consolidar_queue protocol:** browser writes `req_{timestamp}.json` → vigia processes → writes `res_{id}.json` → browser reads and deletes. The browser polls with a 1s interval and 180s timeout.
