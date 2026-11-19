#!/usr/bin/env bash
# Disponibiliza a planilha gerada por HTTP na rede local, so por um tempo limitado.
# Uso:  ./servir.sh              -> serve a descricaoAjustada_NN.xlsx mais recente
#       ./servir.sh arquivo.xlsx -> serve o arquivo indicado
#       PORTA=9000 TEMPO=1200 ./servir.sh
set -euo pipefail

PORTA="${PORTA:-8000}"
TEMPO="${TEMPO:-900}"            # segundos ate o servidor se encerrar sozinho (padrao: 15 min)
PASTA="$(cd "$(dirname "$0")" && pwd)"

if [ $# -ge 1 ]; then
    ARQUIVO="$1"
else
    ARQUIVO="$(ls -1 "$PASTA"/descricaoAjustada_*.xlsx 2>/dev/null | sort -V | tail -n 1 || true)"
fi

if [ -z "${ARQUIVO:-}" ] || [ ! -f "$ARQUIVO" ]; then
    echo "Nenhuma planilha encontrada. Rode o main.py antes ou informe o arquivo: ./servir.sh arquivo.xlsx" >&2
    exit 1
fi

# serve uma pasta temporaria com SO a planilha (nunca o projeto inteiro)
TMP="$(mktemp -d)"
trap 'kill "$PID" 2>/dev/null || true; rm -rf "$TMP"; echo; echo "Servidor encerrado."' EXIT INT TERM
cp "$ARQUIVO" "$TMP/"
NOME="$(basename "$ARQUIVO")"

# no container, passe o IP da maquina: IP_HOST=192.168.x.x (no Docker, o hostname -I mostra o IP interno)
IP="${IP_HOST:-$(hostname -I 2>/dev/null | awk '{print $1}')}"
IP="${IP:-IP-DA-MAQUINA}"

cd "$TMP"
python3 -m http.server "$PORTA" --bind 0.0.0.0 >/dev/null 2>&1 &
PID=$!
sleep 1
if ! kill -0 "$PID" 2>/dev/null; then
    echo "Nao consegui subir o servidor na porta $PORTA (porta em uso?). Tente: PORTA=9000 ./servir.sh" >&2
    exit 1
fi

echo "Servindo: $NOME"
echo "No Windows, abra no navegador ou use PowerShell:"
echo "  http://$IP:$PORTA/$NOME"
echo "  Invoke-WebRequest http://$IP:$PORTA/$NOME -OutFile $NOME"
echo "Encerra sozinho em $((TEMPO / 60)) min, ou aperte Ctrl+C."

sleep "$TEMPO" &
wait $! || true