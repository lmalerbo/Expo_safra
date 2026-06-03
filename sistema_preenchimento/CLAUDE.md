# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the system

```bat
# Start the background watcher (runs 24h on the server PC)
INICIAR_VIGIA.bat         # launches vigia.py via pythonw (no console window)

# Manually trigger consolidation (outside the browser flow)
CONSOLIDAR.bat            # runs engine/consolidar.py interactively

# Update farm/user data embedded in formulario.html
ATUALIZAR.bat             # runs engine/atualizar_programacao.py
```

> `vigia.py` also triggers `engine/consolidar.py` automatically after each browser request, so `CONSOLIDAR.bat` is rarely needed manually.
>
> `ATUALIZAR.bat` is also rarely needed manually — `engine/consolidar.py` calls `atualizar_html()` at the end of every consolidation run.

Run a single test:
```bat
python -m pytest tests/test_utils.py -v
```

## Architecture

The system has two loosely-coupled halves that communicate only through JSON files on a shared G: drive.

### Browser half (`formulario.html`)
A single self-contained HTML file with no build step. All JS is inline. Users open it directly in Chrome. Key responsibilities:
- Field registration UI (layers, fazendas, talhões, areas)
- Tabs: Registros, Consultas, Dashboards
- On "Consolidar": writes `consolidar_queue/req_{id}.json`, polls for `res_{id}.json`, then deletes it
- File copy feature: uses File System Access API to scan `FAZENDAS\` for `.dwg`/`.zip` and copy to SharePoint folders, running in parallel with Python via `Promise.all`

Static data (`FAZENDAS`, `USUARIOS_HA`, `USUARIOS_CONFIG`) is embedded directly in the HTML as JS `const` assignments — injected/updated by `engine/utils.py:atualizar_html()` using line-prefix markers.

IndexedDB (`pf-config` database, `handles` store) persists `FileSystemDirectoryHandle` objects for:
- `queueDir` — `consolidar_queue/` folder
- `filesRootDir` — FAZENDAS root for project file scan
- `filesDwgDir` — SharePoint destination for `.dwg`
- `filesExpDir` — SharePoint destination for `.zip`

### Python half (`vigia.py` + `engine/`)
`vigia.py` polls `consolidar_queue/` every 2s (configurable). On finding `req_{id}.json`:
1. Saves records to `exports/export_frentes_{user}_{date}_{id}.xlsx`
2. Runs `engine/consolidar.py` (which updates `programacao_frentes.xlsx` and regenerates `formulario.html`)
3. Writes `res_{id}.json` with `{ok, arquivo, output}`

`engine/consolidar.py` — matches export rows to the master spreadsheet by LAYER string, overwrites STATUS/TIPO_LINHA/CICLO columns, moves processed exports to `exports/processados/`, then calls `atualizar_html()`.

`engine/atualizar_programacao.py` — reads `programacao_frentes.xlsx` and `base_fazendas/base.xlsx` to rebuild the fazendas/talhões data, then calls `atualizar_html()`.

`engine/utils.py` — shared helpers: `layer_to_str()`, `arquivo_bloqueado()`, `aguardar_arquivo_livre()`, `atualizar_html()`, stdout tee logging.

### Configuration
`config.json` — tuning knobs: poll interval, consolidation timeout, retry settings, column index map for the master spreadsheet.

`usuarios.json` — user list with `preenchimento` (can register + consolidate) or `dashboard` (read-only) profiles.

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
