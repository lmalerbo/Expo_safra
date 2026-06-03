"""
Testes unitários para engine/utils.py

Uso:
  python tests/test_utils.py
"""

import os
import sys
import json
import tempfile

# Garante que engine/ está no path
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(_ROOT, 'engine'))
from utils import layer_to_str, atualizar_html

_ok = _fail = 0

def _assert(cond, desc):
    global _ok, _fail
    if cond:
        print(f"  OK  {desc}")
        _ok += 1
    else:
        print(f"  FALHOU  {desc}")
        _fail += 1


# ── layer_to_str ──────────────────────────────────────────────────────────

def test_layer_to_str():
    print("\nlayer_to_str:")
    _assert(layer_to_str(1001005)      == "1001005", "int direto")
    _assert(layer_to_str(1001005.0)    == "1001005", "float do pandas")
    _assert(layer_to_str("1001005")    == "1001005", "string inteira")
    _assert(layer_to_str("1001005.0")  == "1001005", "string float")
    _assert(layer_to_str("  1001005 ") == "1001005", "string com espaços")
    _assert(layer_to_str(0)            == "0",       "zero")
    _assert(layer_to_str(None)         == "",        "None")
    _assert(layer_to_str("")           == "",        "string vazia")
    _assert(layer_to_str("nan")        == "",        "string 'nan'")
    _assert(layer_to_str("abc")        == "",        "string inválida")


# ── atualizar_html ────────────────────────────────────────────────────────

_HTML_TEMPLATE = """\
<html>
<script>
// ── Data ──────────────────────────────────────────────────────────────────
const USUARIOS_HA = {"alice": 100.0};
const USUARIOS_CONFIG = [{"nome": "alice", "perfis": ["preenchimento"]}];
const FAZENDAS = {"Fazenda A": {"cod": 1001, "talhoes": {"1": {"layer": 1001001, "ha": 50.0}}}};
const FAZENDA_NAMES = Object.keys(FAZENDAS).sort();
</script>
</html>
"""

def test_atualizar_html():
    print("\natualizar_html:")

    fazendas = {"Fazenda B": {"cod": 2002, "talhoes": {"5": {"layer": 2002005, "ha": 120.0}}}}
    uha      = {"bob": 120.0}
    ucfg     = [{"nome": "bob", "perfis": ["preenchimento", "dashboard"]}]

    # Escreve HTML de teste em arquivo temporário
    fd, tmp_path = tempfile.mkstemp(suffix='.html')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(_HTML_TEMPLATE)

        ok, msg = atualizar_html(tmp_path, fazendas, uha, ucfg)
        _assert(ok, f"retornou ok (msg: {msg})")

        with open(tmp_path, 'r', encoding='utf-8') as f:
            content = f.read()

        _assert('"Fazenda B"' in content, "FAZENDAS atualizado com nova fazenda")
        _assert('"Fazenda A"' not in content, "FAZENDAS removeu fazenda antiga")
        _assert('"bob"' in content and '"alice"' not in content, "USUARIOS_HA atualizado")
        _assert('"dashboard"' in content, "USUARIOS_CONFIG atualizado")
        _assert('const FAZENDA_NAMES' in content, "linha seguinte preservada")

        # Chamada sem alteração → "sem alterações"
        ok2, msg2 = atualizar_html(tmp_path, fazendas, uha, ucfg)
        _assert(ok2 and "sem alterações" in msg2, "sem alterações quando dados iguais")

    finally:
        try: os.unlink(tmp_path)
        except Exception: pass


def test_atualizar_html_arquivo_inexistente():
    print("\natualizar_html (arquivo inexistente):")
    ok, msg = atualizar_html("nao_existe.html", {}, {}, [])
    _assert(not ok, "retorna False para arquivo inexistente")
    _assert("não encontrado" in msg, "mensagem de erro descritiva")


def test_atualizar_html_marcador_ausente():
    print("\natualizar_html (marcador ausente):")
    html_sem_marcador = "<html><script>var x = 1;</script></html>"
    fd, tmp_path = tempfile.mkstemp(suffix='.html')
    try:
        with os.fdopen(fd, 'w', encoding='utf-8') as f:
            f.write(html_sem_marcador)
        ok, msg = atualizar_html(tmp_path, {}, {}, [])
        _assert(not ok, "retorna False quando marcadores ausentes")
        _assert("FAZENDAS" in msg, "lista marcadores ausentes na mensagem")
    finally:
        try: os.unlink(tmp_path)
        except Exception: pass


# ── Runner ────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    print("=" * 50)
    print("  Testes — engine/utils.py")
    print("=" * 50)

    test_layer_to_str()
    test_atualizar_html()
    test_atualizar_html_arquivo_inexistente()
    test_atualizar_html_marcador_ausente()

    print(f"\n{'='*50}")
    print(f"  Resultado: {_ok} OK  |  {_fail} FALHOU")
    print(f"{'='*50}")
    sys.exit(0 if _fail == 0 else 1)
