#!/usr/bin/env bash

set -euo pipefail
shopt -s nullglob

MODEL="${MODEL:-gpt-6.1-sol}"   # een vast Codex-model voor iedereen, noteer het
RUNS=1
TIME_LIMIT_SECONDS=420   # maximaal 7 minuten per prompt
SAVE_AFTER_SECONDS=390   # spreadsheet veiligstellen na 6 minuten 30 seconden
START_FROM="${START_FROM:-}"   # optioneel: bijvoorbeeld student02
SUFFIX=$'\n\n---\nSla het eindresultaat op als model.xlsx in de huidige map. Sla eerst een bruikbare basisversie op, voordat je uitgebreid onderzoek of verdere verbeteringen uitvoert. Sla daarna elke verbetering direct op. Gebruik een tijdelijk .xlsx-bestand en vervang model.xlsx atomair na een geslaagde save. Bewaar alle benodigde code in de huidige map, niet uitsluitend in /tmp. Controleer na iedere stap of SAVE_NOW.txt bestaat; zo ja, stop onderzoek en uitbreidingen, sla de huidige spreadsheet onmiddellijk op en rond af. De save-deadline is 6 minuten 30 seconden; na 7 minuten wordt het proces gestopt.'

cd "$(dirname "$0")"
BASE_DIR="$PWD"

if ! command -v codex >/dev/null 2>&1; then
    echo "FOUT: codex is niet gevonden in PATH." >&2
    exit 1
fi
if ! command -v python3 >/dev/null 2>&1; then
    echo "FOUT: python3 is nodig om de tijdslimiet te bewaken." >&2
    exit 1
fi

# Bewaar checkpoints en stop na 7 minuten de hele procesgroep.
run_with_timeout() {
    python3 "$BASE_DIR/run_with_deadline.py" \
        "$out" "$SAVE_AFTER_SECONDS" "$TIME_LIMIT_SECONDS" "$@"
}

prompts=(prompts/*.txt)
if [ "${#prompts[@]}" -eq 0 ]; then
    echo "FOUT: geen prompts/*.txt gevonden." >&2
    exit 1
fi
if [ -n "$START_FROM" ] && [ ! -f "prompts/$START_FROM.txt" ]; then
    echo "FOUT: startprompt prompts/$START_FROM.txt niet gevonden." >&2
    exit 1
fi
mkdir -p runs
echo "Model: $MODEL; pogingen per prompt: $RUNS; save-checkpoint: 6:30; tijdslimiet: 7 minuten"

for p in "${prompts[@]}"; do
    name=$(basename "$p" .txt)
    if [ -n "$START_FROM" ]; then
        if [ "$name" != "$START_FROM" ]; then
            echo "SKIP  $name (voor startprompt $START_FROM)"
            continue
        fi
        START_FROM=""
    fi

    for ((i=1; i<=RUNS; i++)); do
        out="$BASE_DIR/runs/${name}_run${i}"

        if [ -s "$out/model.xlsx" ]; then
            echo "SKIP  $out bestaat al"
            continue
        fi

        mkdir -p "$out"
        echo ">> $name - poging $i"

        if ! {
            cat "$p"
            printf '%s' "$SUFFIX"
        } | run_with_timeout codex exec \
            --model "$MODEL" \
            -c model_reasoning_effort="medium" \
            --sandbox workspace-write \
            --skip-git-repo-check \
            --cd "$out" \
            --output-last-message "$out/laatste_antwoord.txt" \
            - > "$out/log.txt" 2>&1; then
            echo "MISLUKT  $name poging $i (zie $out/log.txt)" >&2
            tail -n 12 "$out/log.txt" >&2
            exit 1
        fi

        if [ ! -s "$out/model.xlsx" ]; then
            echo "MISLUKT  $name poging $i: geen model.xlsx gemaakt (zie $out/log.txt en $out/laatste_antwoord.txt)." >&2
            exit 1
        fi
    done
done

echo "Klaar"
