"""Utilitários compartilhados entre os scripts da engine."""

import json
import os
import datetime


# ── Normalização de LAYER ──────────────────────────────────────────────────

def layer_to_str(v):
    """Converte qualquer representação de LAYER para string inteira padronizada.

    Trata floats do pandas (1001005.0), strings ("1001005"), ints (1001005) e
    casos inválidos (None, '', 'nan') de forma uniforme.

    Exemplos:
        1001005.0  → "1001005"
        "1001005"  → "1001005"
        None       → ""
        "nan"      → ""
    """
    if v is None:
        return ''
    try:
        s = str(v).strip()
        if s in ('', 'nan'):
            return ''
        return str(int(float(s)))
    except (ValueError, TypeError):
        return ''


# ── Logging persistente ────────────────────────────────────────────────────

class _TeeWriter:
    """Encaminha escrita para múltiplos writers (console + arquivo)."""

    def __init__(self, *writers):
        self._writers = writers

    def write(self, text):
        for w in self._writers:
            try:
                w.write(text)
            except Exception:
                pass

    def flush(self):
        for w in self._writers:
            try:
                w.flush()
            except Exception:
                pass

    def isatty(self):
        return False


def redirecionar_stdout(log_path):
    """Redireciona sys.stdout para escrever simultaneamente no console e em log_path.

    Deve ser chamado no início do script. Retorna o handle do arquivo de log
    para que possa ser fechado ao final, se necessário.
    """
    import sys
    os.makedirs(os.path.dirname(os.path.abspath(log_path)), exist_ok=True)
    try:
        log_fh = open(log_path, 'a', encoding='utf-8')
    except Exception:
        return None

    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_fh.write(f'\n{"="*60}\n[{ts}] Sessão iniciada\n{"="*60}\n')
    log_fh.flush()

    sys.stdout = _TeeWriter(sys.__stdout__, log_fh)
    return log_fh


def fechar_log(log_fh):
    """Fecha o handle do log e restaura sys.stdout."""
    import sys
    if log_fh is None:
        return
    try:
        ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_fh.write(f'[{ts}] Sessão encerrada\n')
        log_fh.flush()
        log_fh.close()
    except Exception:
        pass
    finally:
        try:
            sys.stdout = sys.__stdout__
        except Exception:
            pass


# ── Verificação de arquivo bloqueado ──────────────────────────────────────

def arquivo_bloqueado(path):
    """Retorna True se o arquivo está aberto/bloqueado por outro processo."""
    try:
        with open(path, 'a+b'):
            return False
    except (IOError, PermissionError):
        return True


def aguardar_arquivo_livre(path, tentativas=3, espera_s=10):
    """Aguarda o arquivo ficar livre, com tentativas e pausa configuráveis.

    Retorna True se livre, False se ainda bloqueado após todas as tentativas.
    """
    import time
    for i in range(tentativas):
        if not arquivo_bloqueado(path):
            return True
        print(f"  Arquivo em uso, aguardando {espera_s}s... ({i+1}/{tentativas})")
        time.sleep(espera_s)
    return not arquivo_bloqueado(path)


# ── Atualização do formulario.html ────────────────────────────────────────

def atualizar_html(html_file, fazendas_data, usuario_ha, usuarios_config):
    """Substitui os blocos de dados embutidos no formulario.html.

    Usa substituição linha a linha buscando os marcadores 'const X ='
    no início de cada linha — mais robusto do que regex em JSON grande.

    Retorna (ok: bool, msg: str).
    """
    if not os.path.exists(html_file):
        return False, f"Arquivo não encontrado: {html_file}"

    faz_json  = json.dumps(fazendas_data,   ensure_ascii=False)
    uha_json  = json.dumps(usuario_ha,      ensure_ascii=False)
    ucfg_json = json.dumps(usuarios_config, ensure_ascii=False)

    try:
        with open(html_file, 'r', encoding='utf-8') as f:
            html_original = f.read()
    except Exception as e:
        return False, f"Erro ao ler HTML: {e}"

    lines  = html_original.splitlines(keepends=True)
    faz_ok = uha_ok = ucfg_ok = False

    for i, line in enumerate(lines):
        s    = line.lstrip()
        tail = line[len(line.rstrip('\r\n')):]
        if not faz_ok and s.startswith('const FAZENDAS ='):
            lines[i] = f'const FAZENDAS = {faz_json};{tail}'
            faz_ok   = True
        elif not uha_ok and s.startswith('const USUARIOS_HA ='):
            lines[i] = f'const USUARIOS_HA = {uha_json};{tail}'
            uha_ok   = True
        elif not ucfg_ok and s.startswith('const USUARIOS_CONFIG ='):
            lines[i] = f'const USUARIOS_CONFIG = {ucfg_json};{tail}'
            ucfg_ok  = True
        if faz_ok and uha_ok and ucfg_ok:
            break

    ausentes = [n for n, ok in [('FAZENDAS', faz_ok), ('USUARIOS_HA', uha_ok), ('USUARIOS_CONFIG', ucfg_ok)] if not ok]
    if ausentes:
        return False, f"Marcadores não encontrados: {', '.join(ausentes)}"

    html_novo = ''.join(lines)
    if html_novo == html_original:
        return True, "sem alterações"

    try:
        with open(html_file, 'w', encoding='utf-8') as f:
            f.write(html_novo)
    except Exception as e:
        return False, f"Erro ao salvar HTML: {e}"

    n_faz = len(fazendas_data)
    n_tal = sum(len(v.get('talhoes', {})) for v in fazendas_data.values())
    return True, f"{n_faz} fazendas, {n_tal} talhões"
