# 🚀 ROADMAP: Fase 2 - Google Drive Integration

## 📋 Objetivo

Após gerar as apostilas no repositório, fazer upload automático para Google Drive.

## ⚙️ Arquitetura

```
ebe-fabrica-apostilas (GERADOR)
        ⬇️
   Apostilas (.docx)
        ⬇️
   Repositório Git
        ⬇️
  Action de Upload
        ⬇️
  Google Drive
```

## 📅 Timeline

| Fase | Descrição | Status |
|------|-----------|--------|
| **1** | Gerar apostilas em Git | ✅ Atual |
| **2** | Upload para Google Drive | ⏳ Próximo |
| **3** | Dashboard de progresso | 📋 Futuro |
| **4** | Backup em Dropbox | 📋 Futuro |

## 🔧 Configuração Google Drive (Fase 2)

### Passo 1: Criar Service Account

1. Ir para: https://console.cloud.google.com
2. Criar novo projeto: "EBE Apostilas"
3. Ativar "Google Drive API"
4. Criar "Service Account"
5. Gerar chave JSON

### Passo 2: Compartilhar Pasta

1. Criar pasta no Google Drive: "EBE Apostilas"
2. Ir para "Compartilhar"
3. Adicionar email da service account
4. Copiar folder ID da URL

### Passo 3: GitHub Secrets

```
GOOGLE_DRIVE_CREDENTIALS = [JSON da service account]
GOOGLE_DRIVE_FOLDER_ID = [ID da pasta]
```

### Passo 4: Criar Action

Arquivo: `.github/workflows/upload-google-drive.yml`

```yaml
name: "☁️ Upload para Google Drive"

on:
  push:
    branches: [main]
    paths:
      - "apostilas/**"

jobs:
  upload:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      - run: pip install google-api-python-client google-auth
      - env:
          GOOGLE_DRIVE_CREDENTIALS: ${{ secrets.GOOGLE_DRIVE_CREDENTIALS }}
          GOOGLE_DRIVE_FOLDER_ID: ${{ secrets.GOOGLE_DRIVE_FOLDER_ID }}
        run: python scripts/upload_google_drive.py
```

## 📊 Estimativa

| Métrica | Valor |
|---------|-------|
| **Tempo Fase 1** | 3 meses |
| **Tempo Fase 2 Setup** | 1-2 horas |
| **Tempo Fase 2 Automação** | Paralelo |
| **Total até distribuído** | 3 meses |

## 🎯 Próximos Passos

1. ✅ Completar Fase 1 (geração em 3 meses)
2. ⏳ Configurar Google Drive (1-2 horas)
3. ⏳ Implementar upload automático
4. ⏳ Criar dashboard

