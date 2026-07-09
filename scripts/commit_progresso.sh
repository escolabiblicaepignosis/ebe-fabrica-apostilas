#!/bin/bash
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# commit_progresso.sh — Commit e push do progresso
# Chamado pelo workflow do GitHub Actions
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
set -euo pipefail

echo "📦 A preparar commit do progresso..."

git config user.name "EBE Bot"
git config user.email "bot@escolabiblicaepignosis.org"

# Estado é sempre necessário
if [ -f scripts/estado.json ]; then
  git add scripts/estado.json
  echo "  ✅ scripts/estado.json"
fi

# Batch é opcional
if [ -f scripts/ultimo_batch.json ]; then
  git add scripts/ultimo_batch.json
  echo "  ✅ scripts/ultimo_batch.json"
fi

# Output é opcional
if [ -d output ] && ls output/*.docx 1>/dev/null 2>&1; then
  git add output/
  echo "  ✅ output/"
fi

# Só commitar se houver mudanças
if git diff --staged --quiet; then
  echo "  ℹ️  Sem mudanças para commitar"
else
  git commit -m "🤖 Apostilas geradas — batch $(date +'%Y-%m-%d %H:%M')"
  echo "  ✅ Commit feito"
fi

# Push (não falhar o workflow se push falhar)
if git push; then
  echo "  ✅ Push feito"
else
  echo "  ⚠️  Push falhou (não crítico)"
fi

echo "✅ Progresso guardado"
