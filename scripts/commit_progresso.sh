#!/bin/bash
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# Script de commit do progresso — chamado pelo workflow
# Resolve o problema do git add condicional
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
set -e

git config user.name "EBE Bot"
git config user.email "bot@escolabiblicaepignosis.org"

# Sempre adicionar o estado
git add scripts/estado.json

# Só adicionar se existir
[ -f scripts/ultimo_batch.json ] && git add scripts/ultimo_batch.json || true
[ -d output ] && git add output/ || true

# Só commitar e pushar se houver mudanças
git diff --staged --quiet || git commit -m "🤖 Apostilas geradas — batch $(date +'%Y-%m-%d %H:%M')"
git push || true
