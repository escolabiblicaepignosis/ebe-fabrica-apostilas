#!/usr/bin/env python3
"""
Upload de apostilas EBE para o Google Drive.
Organiza em pastas: Instituto > Escola > Curso > Módulo
"""

import json
import os
import sys

def log(msg):
    print(msg, flush=True)

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DRIVE_FOLDER_ID = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "")
CREDENTIALS_JSON = os.environ.get("GOOGLE_DRIVE_CREDENTIALS", "")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "..", "output")
BATCH_PATH = os.path.join(SCRIPT_DIR, "ultimo_batch.json")
ESTADO_PATH = os.path.join(SCRIPT_DIR, "estado.json")
SCOPES = ["https://www.googleapis.com/auth/drive.file"]


def get_drive_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    if not CREDENTIALS_JSON:
        raise RuntimeError("GOOGLE_DRIVE_CREDENTIALS não definido")
    creds_info = json.loads(CREDENTIALS_JSON)
    credentials = service_account.Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    return build("drive", "v3", credentials=credentials)


def sanitizar_nome(nome):
    for char in ['"', "'", '/', '\\', '|', ':', '*', '?', '<', '>']:
        nome = nome.replace(char, '')
    return nome.strip()[:100]


def encontrar_ou_criar_pasta(service, nome, parent_id):
    nome = sanitizar_nome(nome)
    query = (
        f"name = '{nome}' and '{parent_id}' in parents and "
        f"mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    )
    results = service.files().list(q=query, spaces='drive', fields='files(id, name)', pageSize=10).execute()
    files = results.get('files', [])
    if files:
        return files[0]['id']
    folder = service.files().create(
        body={'name': nome, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [parent_id]},
        fields='id'
    ).execute()
    log(f"    📁 Pasta criada: {nome}")
    return folder['id']


def garantir_caminho(service, parent_id, partes):
    current_id = parent_id
    for parte in partes:
        if parte:
            current_id = encontrar_ou_criar_pasta(service, parte, current_id)
    return current_id


def upload_ficheiro(service, file_path, file_name, folder_id):
    from googleapiclient.http import MediaFileUpload
    media = MediaFileUpload(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        resumable=True
    )
    return service.files().create(
        body={'name': file_name, 'parents': [folder_id]},
        media_body=media, fields='id, name, webViewLink'
    ).execute()


def processar_batch():
    log("=" * 60)
    log("  Upload para Google Drive")
    log("=" * 60)

    if not DRIVE_FOLDER_ID:
        log("❌ GOOGLE_DRIVE_FOLDER_ID não definido.")
        sys.exit(1)
    log(f"✅ Pasta mãe: {DRIVE_FOLDER_ID}")

    if not CREDENTIALS_JSON:
        log("❌ GOOGLE_DRIVE_CREDENTIALS não definido.")
        sys.exit(1)
    log(f"✅ Credenciais: {len(CREDENTIALS_JSON)} chars")

    log(f"📂 OUTPUT_DIR: {os.path.abspath(OUTPUT_DIR)}")
    log(f"📂 Existe: {os.path.exists(OUTPUT_DIR)}")
    if os.path.exists(OUTPUT_DIR):
        log(f"📂 Conteúdo: {os.listdir(OUTPUT_DIR)}")

    if not os.path.exists(BATCH_PATH):
        log("⚠️  Nenhum batch (ultimo_batch.json não encontrado)")
        return

    with open(BATCH_PATH, 'r', encoding='utf-8') as f:
        batch = json.load(f)

    if not batch:
        log("⚠️  Batch vazio.")
        return

    log(f"📋 Batch: {len(batch)} ficheiros")

    service = get_drive_service()
    log("✅ Conectado ao Google Drive")

    pasta_cache = {}
    uploaded = []

    for idx, item in enumerate(batch, 1):
        fname = item["ficheiro"]
        abs_path = os.path.abspath(os.path.join(OUTPUT_DIR, fname))

        log(f"\n  [{idx}/{len(batch)}] {fname}")
        log(f"    Path: {abs_path}")

        if not os.path.exists(abs_path):
            log(f"    ❌ Não encontrado!")
            if os.path.exists(OUTPUT_DIR):
                log(f"    Em output/: {os.listdir(OUTPUT_DIR)}")
            continue

        size_kb = os.path.getsize(abs_path) / 1024
        log(f"    Tamanho: {size_kb:.0f} KB")

        partes = [
            sanitizar_nome(item.get("instituto", "Outros")),
            sanitizar_nome(item.get("escola", "Geral")),
            sanitizar_nome(item.get("curso", "Curso")),
            sanitizar_nome(f"Módulo {item.get('modulo', '1')}"),
        ]
        log(f"    Destino: {' / '.join(partes)}")

        cache_key = " > ".join(partes)
        if cache_key in pasta_cache:
            folder_id = pasta_cache[cache_key]
        else:
            folder_id = garantir_caminho(service, DRIVE_FOLDER_ID, partes)
            pasta_cache[cache_key] = folder_id

        try:
            file = upload_ficheiro(service, abs_path, fname, folder_id)
            link = file.get('webViewLink', '')
            log(f"    ✅ Enviado! ID: {file['id']}")
            log(f"    🔗 {link}")
            uploaded.append({**item, "drive_id": file['id'], "drive_link": link})
        except Exception as e:
            log(f"    ❌ Erro: {e}")
            import traceback
            log(f"    {traceback.format_exc()}")

    if uploaded:
        estado = {}
        if os.path.exists(ESTADO_PATH):
            with open(ESTADO_PATH, 'r', encoding='utf-8') as f:
                estado = json.load(f)
        estado.setdefault("uploads", []).extend(uploaded)
        with open(ESTADO_PATH, 'w', encoding='utf-8') as f:
            json.dump(estado, f, ensure_ascii=False, indent=2)

    log(f"\n{'=' * 60}")
    log(f"  ✅ Upload: {len(uploaded)}/{len(batch)} ficheiros")
    log(f"{'=' * 60}")


if __name__ == "__main__":
    processar_batch()
