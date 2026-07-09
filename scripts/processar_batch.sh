#!/bin/bash
# Script para processar um batch: Gerar, Fazer Upload e Limpar

# 1. Gerar apostilas
echo "🚀 Iniciando geração de apostilas..."
python3 scripts/gerar_apostila.py

# 2. Upload para Google Drive (se as credenciais existirem)
if [ -n "$GOOGLE_DRIVE_CREDENTIALS" ]; then
    echo "☁️ Iniciando upload para Google Drive..."
    python3 scripts/upload_drive.py
    
    # 3. Limpar pasta output apenas se o upload teve sucesso
    if [ $? -eq 0 ]; then
        echo "🧹 Limpando pasta output para poupar espaço..."
        rm -rf output/*
    else
        echo "⚠️ Falha no upload. Mantendo ficheiros em output/."
    fi
else
    echo "⚠️ GOOGLE_DRIVE_CREDENTIALS não configurado. Mantendo ficheiros em output/."
fi
