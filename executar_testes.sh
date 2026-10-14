#!/usr/bin/env bash
# Roda as duas suites: o motor de limpeza e a camada web.
set -u
cd "$(dirname "$0")"
falhou=0
echo "== Motor de limpeza =="
python3 testar_correcoes.py -q 2>&1 | grep -E "^Ran|^OK|^FAILED|FAIL:|ERROR:" || falhou=1
echo "== Camada web =="
python3 testar_app.py -q 2>&1 | grep -v "ScriptRunContext\|bare mode" \
  | grep -E "^Ran|^OK|^FAILED|FAIL:|ERROR:" || falhou=1
exit $falhou
