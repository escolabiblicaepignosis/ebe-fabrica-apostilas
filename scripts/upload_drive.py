#!/usr/bin/env python3
"""
Upload de apostilas EBE para o Google Drive.
Organiza em pastas: Instituto > Escola > Curso > Módulo
"""

import json
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ──────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ──────────────────────────────────────────────────────────────
DRIVE_FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "")
CREDENTIALS_JSON = os.environ.get("GOOGLE_DRIVE_CREDENTIALS", "")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")
BATCH_PATH = os.path.join(os.path.dirname(__file__), "ultimo_batch.json")
ESTADO_PATH = os.path.join(os.path.dirname(__file__), "estado.json")

SCOPES = ["https://www.googleapis.com/auth/drive.file"]


# ──────────────────────────────────────────────────────────────
# AUTENTICAÇÃO
# ──────────────────────────────────────────────────────────────
def get_drive_service():
    """Autentica no Google Drive usando Service Account."""
    if not CREDENTIALS_JSON:
        raise RuntimeError("GOOGLE_DRIVE_CREDENTIALS não definido")

    # As credenciais vêm como JSON string (do GitHub Secret)
    creds_info = json.loads(CREDENTIALS_JSON)

    credentials = service_account.Credentials.from_service_account_info(
        creds_info, scopes=SCOPES
    )
    return build("drive", "v3", credentials=credentials)


# ──────────────────────────────────────────────────────────────
# GESTÃO DE PASTAS
# ──────────────────────────────────────────────────────────────
def sanitizar_nome(nome):
    """Remove caracteres problemáticos para nomes de pasta no Drive."""
    # Remover aspas, pipes, barras
    for char in ['"', "'", '/', '\\', '|', ':', '*']:
        nome = nome.replace(char, '')
    return nome.strip()[:100]


def encontrar_ou_criar_pasta(service, nome, parent_id):
    """Encontra uma pasta existente ou cria uma nova no Drive."""
    nome = sanitizar_nome(nome)

    # Buscar pasta existente
    query = (
        f"name = '{nome}' and "
        f"'{parent_id}' in parents and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )
    results = service.files().list(
        q=query,
        spaces='drive',
        fields='files(id, name)',
        pageSize=10
    ).execute()

    files = results.get('files', [])
    if files:
        return files[0]['id']

    # Criar nova pasta
    folder_metadata = {
        'name': nome,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [parent_id]
    }
    folder = service.files().create(
        body=folder_metadata,
        fields='id'
    ).execute()

    print(f"    📁 Pasta criada: {nome}")
    return folder['id']


def garantir_caminho(service, parent_id, partes):
    """
    Garante que toda a hierarquia de pastas existe.
    Retorna o ID da pasta final.
    """
    current_id = parent_id
    for parte in partes:
        if parte:
            current_id = encontrar_ou_criar_pasta(service, parte, current_id)
    return current_id


# ──────────────────────────────────────────────────────────────
# UPLOAD
# ──────────────────────────────────────────────────────────────
def upload_ficheiro(service, file_path, file_name, folder_id):
    """Faz upload de um ficheiro .docx para o Google Drive."""
    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaFileUpload(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        resumable=True
    )
    file = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, name, webViewLink'
    ).execute()

    return file


def processar_batch():
    """Faz upload de todos os ficheiros do último batch."""
    if not DRIVE_FOLDER_ID:
        print("❌ GOOGLE_DRIVE_FOLDER_ID não definido.")
        sys.exit(1)

    if not os.path.exists(BATCH_PATH):
        print("⚠️  Nenhum batch para fazer upload (ultimo_batch.json não encontrado)")
        return

    with open(BATCH_PATH, 'r', encoding='utf-8') as f:
        batch = json.load(f)

    if not batch:
        print("⚠️  Batch vazio.")
        return

    print("=" * 60)
    print("  Upload para Google Drive")
    print(f"  Pasta mãe: {DRIVE_FOLDER_ID}")
    print(f"  Ficheiros: {len(batch)}")
    print("=" * 60)

    service = get_drive_service()

    # Cache de pastas para evitar recriação
    pasta_cache = {}
    uploaded = []

    for idx, item in enumerate(batch, 1):
        fname = item["ficheiro"]
        file_path = os.path.join(OUTPUT_DIR, fname)

        if not os.path.exists(file_path):
            print(f"  [{idx}] ⚠️  Ficheiro não encontrado: {fname}")
            continue

        # Montar hierarquia: Instituto > Escola > Curso > Módulo
        partes = [
            sanitizar_nome(item.get("instituto", "Outros")),
            sanitizar_nome(item.get("escola", "Geral")),
            sanitizar_nome(item.get("curso", "Curso")),
            sanitizar_nome(f"Módulo {item.get('modulo', '1')}"),
        ]

        # Usar cache para evitar múltiplas buscas
        cache_key = " > ".join(partes)
        if cache_key in pasta_cache:
            folder_id = pasta_cache[cache_key]
        else:
            folder_id = garantir_caminho(service, DRIVE_FOLDER_ID, partes)
            pasta_cache[cache_key] = folder_id

        print(f"  [{idx}/{len(batch)}] Upload: {fname}")
        try:
            file = upload_ficheiro(service, file_path, fname, folder_id)
            link = file.get('webViewLink', '')
            print(f"    ✅ Enviado: {link}")
            uploaded.append({
                **item,
                "drive_id": file['id'],
                "drive_link": link,
            })
        except Exception as e:
            print(f"    ❌ Erro no upload: {e}")

    # Atualizar estado com links do Drive
    if uploaded:
        estado = {}
        if os.path.exists(ESTADO_PATH):
            with open(ESTADO_PATH, 'r', encoding='utf-8') as f:
                estado = json.load(f)

        estado.setdefault("uploads", []).extend(uploaded)
        with open(ESTADO_PATH, 'w', encoding='utf-8') as f:
            json.dump(estado, f, ensure_ascii=False, indent=2)

    print(f"\n{'=' * 60}")
    print(f"  ✅ Upload concluído: {len(uploaded)}/{len(batch)} ficheiros")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    processar_batch()
