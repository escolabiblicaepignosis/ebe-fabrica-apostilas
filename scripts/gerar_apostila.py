#!/usr/bin/env python3
"""
Gerador automático de apostilas EBE usando Gemini AI.
Gera apostilas de 15-20 páginas seguindo o modelo institucional da EBE.
"""

import json
import os
import re
import sys
import time
import unicodedata
import traceback

def log(msg):
    """Log com flush imediato (visível no GitHub Actions)."""
    print(msg, flush=True)

# ──────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ──────────────────────────────────────────────────────────────
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
MAPA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mapa_apostilas.json")
ESTADO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "estado.json")
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output")
BATCH_SIZE = int(os.environ.get("BATCH_SIZE", "5"))
MODEL_NAME = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
DELAY_ENTRE_APOSTILAS = 90  # segundos

# Paleta EBE
from docx.shared import RGBColor, Pt, Cm
COR_PRIMARIA   = RGBColor(0x1B, 0x3A, 0x5C)
COR_SECUNDARIA = RGBColor(0x2E, 0x7D, 0x4F)
COR_TERCIARIA  = RGBColor(0xC9, 0xA1, 0x4B)
COR_TEXTO      = RGBColor(0x1A, 0x1A, 0x1A)
COR_CITACAO    = RGBColor(0x55, 0x55, 0x55)
HEX_PRIMARIA   = "1B3A5C"
HEX_SECUNDARIA = "2E7D4F"
FONTE          = "Garamond"

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src", "estilos", "_assets")
LOGO_PATH  = os.path.join(ASSETS_DIR, "logo_ebe.png")


# ──────────────────────────────────────────────────────────────
# DOCX UTILITIES
# ──────────────────────────────────────────────────────────────
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def _add_horizontal_line(paragraph, color=HEX_PRIMARIA, size=8):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def page_break(doc):
    p = doc.add_paragraph()
    run = p.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run._r.append(br)


def configurar_estilos_base(doc):
    style = doc.styles["Normal"]
    style.font.name = FONTE
    style.font.size = Pt(12)
    style.font.color.rgb = COR_TEXTO
    pf = style.paragraph_format
    pf.space_after = Pt(6)
    pf.space_before = Pt(0)
    pf.line_spacing = 1.4
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for section in doc.sections:
        section.page_height = Cm(29.7)
        section.page_width  = Cm(21.0)
        section.top_margin    = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin   = Cm(3.0)
        section.right_margin  = Cm(2.5)


def inserir_logo(doc, caminho=None, largura_cm=5.5):
    caminho = caminho or LOGO_PATH
    if not os.path.exists(caminho):
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    r.add_picture(caminho, width=Cm(largura_cm))


def add_texto(doc, texto, font_size=12, bold=False, italic=False,
              color=None, alignment=None, indent_cm=0):
    p = doc.add_paragraph()
    if alignment:
        p.alignment = alignment
    if indent_cm:
        p.paragraph_format.left_indent = Cm(indent_cm)
    r = p.add_run(texto)
    r.font.name = FONTE
    r.font.size = Pt(font_size)
    r.font.bold = bold
    r.font.italic = italic
    if color:
        r.font.color.rgb = color
    return p


def cabecalho_rodape(doc, titulo_doc, codigo_doc):
    for section in doc.sections:
        section.different_first_page_header_footer = True
        header = section.header
        ph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        ph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = ph.add_run(f"Escola Bíblica Epignósis  ·  {titulo_doc}")
        r.font.name = FONTE; r.font.size = Pt(9)
        r.font.italic = True; r.font.color.rgb = COR_SECUNDARIA
        footer = section.footer
        pf = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = pf.add_run(f"{codigo_doc}  ·  ")
        r.font.name = FONTE; r.font.size = Pt(9); r.font.color.rgb = COR_CITACAO
        run = pf.add_run()
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")
        instrText = OxmlElement("w:instrText")
        instrText.set(qn("xml:space"), "preserve")
        instrText.text = "PAGE"
        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "end")
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
        run.font.name = FONTE; run.font.size = Pt(9); run.font.color.rgb = COR_CITACAO


def slugify(texto):
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r'[^\w\s-]', '', texto).strip().lower()
    texto = re.sub(r'[\s_]+', '_', texto)
    return texto[:80]


def construir_docx(meta, conteudo_md, output_path):
    doc = Document()
    configurar_estilos_base(doc)
    codigo = meta.get("codigo", "EBE-APO-XXXX")
    cabecalho_rodape(doc, meta.get("titulo", "Apostila"), codigo)

    # CAPA
    doc.add_paragraph()
    inserir_logo(doc, LOGO_PATH, largura_cm=5.5)
    add_texto(doc, "ἐπίγνωσις  ·  Conhecer a Deus. Viver a Palavra. Manifestar o Reino.",
              font_size=11, italic=True, color=COR_SECUNDARIA, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    _add_horizontal_line(p, color=HEX_SECUNDARIA, size=8)
    for _ in range(2): doc.add_paragraph()
    if meta.get("supratitulo"):
        add_texto(doc, meta["supratitulo"].upper(), font_size=11, bold=True,
                  color=COR_SECUNDARIA, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_texto(doc, meta.get("titulo", ""), font_size=26, bold=True,
              color=COR_PRIMARIA, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    if meta.get("subtitulo"):
        add_texto(doc, meta["subtitulo"], font_size=13, italic=True,
                  color=COR_TEXTO, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(4): doc.add_paragraph()
    p = doc.add_paragraph()
    _add_horizontal_line(p, color=HEX_SECUNDARIA, size=4)
    add_texto(doc, meta.get("info_linha", f"Código {codigo} · 2026"),
              font_size=10, color=COR_CITACAO, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    page_break(doc)

    # MARCO FILOSÓFICO
    for _ in range(5): doc.add_paragraph()
    add_texto(doc, "MARCO FILOSÓFICO", font_size=12, bold=True,
              color=COR_SECUNDARIA, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    _add_horizontal_line(p, color=HEX_SECUNDARIA, size=4)
    doc.add_paragraph()
    add_texto(doc,
        "\u201cAcreditamos que o verdadeiro conhecimento de Deus transforma "
        "a mente pela verdade das Escrituras, o coração pela acção do "
        "Espírito Santo e a vida pelo compromisso de viver e anunciar "
        "o Evangelho de Jesus Cristo.\u201d",
        font_size=14, italic=True, color=COR_PRIMARIA,
        alignment=WD_ALIGN_PARAGRAPH.CENTER, indent_cm=2)
    doc.add_paragraph()
    add_texto(doc, "— Escola Bíblica Epignósis —", font_size=10,
              color=COR_CITACAO, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    page_break(doc)

    # CONTEÚDO
    for linha in conteudo_md.strip().split('\n'):
        linha = linha.rstrip()
        if not linha.strip():
            continue
        if linha.startswith('## '):
            texto = linha[3:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(texto)
            r.font.name = FONTE; r.font.size = Pt(13)
            r.font.bold = True; r.font.color.rgb = COR_PRIMARIA
        elif linha.startswith('### '):
            texto = linha[4:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            r = p.add_run(texto)
            r.font.name = FONTE; r.font.size = Pt(11.5)
            r.font.bold = True; r.font.color.rgb = COR_SECUNDARIA
        elif linha.startswith('> '):
            texto = linha[2:].strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(1.5)
            r = p.add_run(texto)
            r.font.name = FONTE; r.font.size = Pt(11)
            r.font.italic = True; r.font.color.rgb = COR_CITACAO
        elif linha.startswith('- ') or linha.startswith('* '):
            texto = re.sub(r'\*\*(.+?)\*\*', r'\1', linha[2:].strip())
            texto = re.sub(r'\*(.+?)\*', r'\1', texto)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(0.8)
            p.paragraph_format.first_line_indent = Cm(-0.5)
            r = p.add_run("•  ")
            r.font.name = FONTE; r.font.size = Pt(12)
            r.font.bold = True; r.font.color.rgb = COR_SECUNDARIA
            r2 = p.add_run(texto)
            r2.font.name = FONTE; r2.font.size = Pt(12)
        elif linha.startswith('# '):
            texto = linha[2:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(18)
            r = p.add_run(texto.upper())
            r.font.name = FONTE; r.font.size = Pt(15)
            r.font.bold = True; r.font.color.rgb = COR_PRIMARIA
            _add_horizontal_line(p, color=HEX_PRIMARIA, size=6)
        else:
            texto = re.sub(r'\*\*(.+?)\*\*', r'\1', linha.strip())
            texto = re.sub(r'\*(.+?)\*', r'\1', texto)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            r = p.add_run(texto)
            r.font.name = FONTE; r.font.size = Pt(12)

    # SELO FINAL
    doc.add_paragraph()
    p = doc.add_paragraph()
    _add_horizontal_line(p, color=HEX_SECUNDARIA, size=4)
    doc.add_paragraph()
    add_texto(doc, "ESCOLA BÍBLICA EPIGNÓSIS", font_size=11, bold=True,
              color=COR_PRIMARIA, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_texto(doc, "Conhecer a Deus. Viver a Palavra. Manifestar o Reino.",
              font_size=10, italic=True, color=COR_SECUNDARIA, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_texto(doc, "Soli Deo Gloria", font_size=9, italic=True,
              color=COR_CITACAO, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    doc.save(output_path)


# ──────────────────────────────────────────────────────────────
# GEMINI
# ──────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """Você é um teólogo, pedagogo e editor cristão sénior da Escola Bíblica Epignósis (EBE).
Redija o conteúdo integral de uma apostila académica bíblica em português europeu/Angola (pt-PT).

REGRAS:
1. Use APENAS a Almeida Revista e Corrigida (ARC).
2. 15-20 páginas (~6.000-8.000 palavras).
3. Estilo académico formal, acessível e devocional.
4. Todas as citações bíblicas com referência completa.
5. Doutrinariamente protestante reformado/evangélico.
6. Substância real em cada secção — sem placeholders.
7. Grego/hebraico quando relevante (com transliteração).
8. Escreva em Markdown (## títulos, **negrito**, > citações, - listas).

ESTRUTURA OBRIGATÓRIA:
## FICHA TÉCNICA
## SUMÁRIO
## APRESENTAÇÃO DA APOSTILA
## OBJECTIVOS DE APRENDIZAGEM (Conhecer 40%, Crer 20%, Viver 20%, Servir 20%)
## VERSÍCULO-CHAVE
## TEXTO-BASE PARA LEITURA
## 1. INTRODUÇÃO
## 2. DESENVOLVIMENTO DO CONCEITO CENTRAL
### 2.1 Fundamentos bíblicos
### 2.2-2.4 Desenvolvimento temático
### 2.5 Quadro de Destaque
## 3. APLICAÇÃO PRÁTICA
## 4. SÍNTESE E CONCLUSÃO
## EXERCÍCIOS DE REVISÃO
## ESTUDO BÍBLICO COMPLEMENTAR
## PARA A PRÓXIMA APOSTILA
## GLOSSÁRIO
## BIBLIOGRAFIA RECOMENDADA
## ANOTAÇÕES PESSOAIS
"""


def gerar_conteudo_gemini(meta):
    import google.generativeai as genai

    log(f"  📡 Configurando Gemini API...")
    genai.configure(api_key=GEMINI_API_KEY)

    log(f"  🤖 Criando modelo: {MODEL_NAME}")
    model = genai.GenerativeModel(
        model_name=MODEL_NAME,
        system_instruction=SYSTEM_PROMPT,
    )

    prompt = f"""Gere o conteúdo integral da seguinte apostila:

- Nível: {meta['nivel']}
- Instituto: {meta['instituto']}
- Escola: {meta['escola']}
- Curso: {meta['curso']} (Carga horária: {meta.get('carga_horaria', 'N/D')})
- Módulo: {meta['modulo']}
- Apostila N.º: {meta['numero_apostila']} — {meta['titulo']}
- Código: {meta['codigo']}

Gere em Markdown. Não inclua capa nem marco filosófico — apenas da FICHA TÉCNICA até ANOTAÇÕES PESSOAIS.
"""

    log(f"  📤 Enviando prompt para Gemini ({len(prompt)} chars)...")
    response = model.generate_content(
        prompt,
        generation_config=genai.types.GenerationConfig(
            temperature=0.7,
            max_output_tokens=16384,
            top_p=0.9,
        ),
    )

    # Verificar se a resposta é válida
    if not response:
        raise RuntimeError("Resposta vazia do Gemini")

    if hasattr(response, 'prompt_feedback') and response.prompt_feedback:
        log(f"  ⚠️  Feedback: {response.prompt_feedback}")

    try:
        texto = response.text
    except Exception as e:
        log(f"  ❌ Erro ao aceder a response.text: {e}")
        if hasattr(response, 'candidates'):
            log(f"  Candidates: {response.candidates}")
        raise RuntimeError(f"Resposta do Gemini sem texto: {e}")

    if not texto or len(texto) < 100:
        raise RuntimeError(f"Conteúdo muito curto ({len(texto) if texto else 0} chars)")

    # Limpar delimitadores markdown
    texto = re.sub(r'^```(?:markdown)?\s*\n?', '', texto, flags=re.MULTILINE)
    texto = re.sub(r'\n?```\s*$', '', texto, flags=re.MULTILINE)

    log(f"  ✅ Resposta recebida: {len(texto)} caracteres")
    return texto.strip()


# ──────────────────────────────────────────────────────────────
# ESTADO
# ──────────────────────────────────────────────────────────────
def carregar_estado():
    if os.path.exists(ESTADO_PATH):
        with open(ESTADO_PATH, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"concluidas": [], "falhadas": [], "total_gerado": 0}


def salvar_estado(estado):
    with open(ESTADO_PATH, 'w', encoding='utf-8') as f:
        json.dump(estado, f, ensure_ascii=False, indent=2)


def carregar_mapa():
    with open(MAPA_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def obter_proximas_apostilas(mapa, estado, batch_size):
    concluidas = set(estado.get("concluidas", []))
    pendentes = []
    for nivel in mapa["niveis"]:
        for inst in nivel["institutos"]:
            for escola in inst["escolas"]:
                for curso in escola["cursos"]:
                    carga = curso.get("carga_horaria", "N/D")
                    for modulo in curso["modulos"]:
                        for apo in modulo["apostilas"]:
                            if apo["id"] not in concluidas:
                                pendentes.append({
                                    "id": apo["id"],
                                    "titulo": apo["titulo"],
                                    "nivel": nivel["id"],
                                    "instituto": inst["nome"],
                                    "escola": escola["nome"],
                                    "curso": curso["nome"],
                                    "carga_horaria": carga,
                                    "modulo": f"{modulo['numero']} — {modulo['nome']}",
                                    "numero_apostila": apo["id"].split("-")[1].lstrip("0") or "1",
                                    "codigo": apo["id"],
                                })
    return pendentes[:batch_size]


# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
def main():
    log("=" * 60)
    log("  ESCOLA BÍBLICA EPIGNÓSIS — Gerador de Apostilas")
    log(f"  Modelo: {MODEL_NAME} | Batch: {BATCH_SIZE}")
    log("=" * 60)

    # Verificações iniciais
    if not GEMINI_API_KEY:
        log("❌ ERRO CRÍTICO: GEMINI_API_KEY não definida!")
        log(f"   Tamanho da chave: {len(GEMINI_API_KEY)}")
        sys.exit(1)
    else:
        log(f"✅ GEMINI_API_KEY configurada ({len(GEMINI_API_KEY)} chars)")

    if not os.path.exists(MAPA_PATH):
        log(f"❌ ERRO CRÍTICO: Mapa não encontrado: {MAPA_PATH}")
        sys.exit(1)
    log(f"✅ Mapa encontrado: {MAPA_PATH}")

    mapa = carregar_mapa()
    estado = carregar_estado()
    log(f"📊 Mapa: {mapa['total_apostilas']} apostilas | Estado: {len(estado.get('concluidas', []))} concluídas")

    proximas = obter_proximas_apostilas(mapa, estado, BATCH_SIZE)
    if not proximas:
        log("🎉 Todas as apostilas já foram geradas!")
        return

    log(f"📋 Próximas {len(proximas)} apostilas:")
    for apo in proximas:
        log(f"   • {apo['codigo']} — {apo['titulo']}")

    geradas = []
    falhadas = []

    for idx, meta in enumerate(proximas, 1):
        log(f"\n{'─' * 60}")
        log(f"  [{idx}/{len(proximas)}] {meta['codigo']} — {meta['titulo']}")
        log(f"  Instituto: {meta['instituto']}")
        log(f"  Escola: {meta['escola']}")
        log(f"  Curso: {meta['curso']} | Módulo: {meta['modulo']}")

        try:
            # Gerar conteúdo
            conteudo = gerar_conteudo_gemini(meta)

            if not conteudo or len(conteudo) < 500:
                raise RuntimeError(f"Conteúdo muito curto: {len(conteudo) if conteudo else 0} chars")

            log(f"  📄 Conteúdo: {len(conteudo)} caracteres")

            # Construir docx
            supratitulo = meta['instituto']
            subtitulo = f"{meta['escola']}  ·  Curso «{meta['curso']}»  ·  Módulo {meta['modulo']}"
            docx_meta = {
                "titulo": meta["titulo"],
                "supratitulo": supratitulo,
                "subtitulo": subtitulo,
                "codigo": meta["codigo"],
                "info_linha": f"Material didáctico oficial · Código {meta['codigo']} · 2026",
            }

            os.makedirs(OUTPUT_DIR, exist_ok=True)
            fname = f"{meta['codigo']}_{slugify(meta['titulo'])}.docx"
            output_path = os.path.join(OUTPUT_DIR, fname)

            log(f"  🏗️  Construindo .docx...")
            construir_docx(docx_meta, conteudo, output_path)

            size_kb = os.path.getsize(output_path) / 1024
            log(f"  ✅ Criado: {fname} ({size_kb:.0f} KB)")

            geradas.append({
                "id": meta["id"],
                "codigo": meta["codigo"],
                "titulo": meta["titulo"],
                "ficheiro": fname,
                "nivel": meta["nivel"],
                "instituto": meta["instituto"],
                "escola": meta["escola"],
                "curso": meta["curso"],
                "modulo": meta["modulo"],
            })

            estado["concluidas"].append(meta["id"])
            estado["total_gerado"] = len(estado["concluidas"])
            salvar_estado(estado)
            log(f"  💾 Estado actualizado: {estado['total_gerado']}/{mapa['total_apostilas']}")

        except Exception as e:
            log(f"  ❌ ERRO: {e}")
            log(f"  📍 Traceback:\n{traceback.format_exc()}")
            falhadas.append({"id": meta["id"], "erro": str(e)})
            estado.setdefault("falhadas", []).append(meta["id"])
            salvar_estado(estado)

        if idx < len(proximas):
            log(f"\n  ⏳ Aguardando {DELAY_ENTRE_APOSTILAS}s...")
            time.sleep(DELAY_ENTRE_APOSTILAS)

    # Resumo
    log(f"\n{'=' * 60}")
    log(f"  RESUMO: ✅ {len(geradas)} geradas | ❌ {len(falhadas)} falhadas")
    log(f"  Total: {len(estado.get('concluidas', []))}/{mapa['total_apostilas']}")
    log(f"{'=' * 60}")

    if geradas:
        batch_path = os.path.join(os.path.dirname(ESTADO_PATH), "ultimo_batch.json")
        with open(batch_path, 'w', encoding='utf-8') as f:
            json.dump(geradas, f, ensure_ascii=False, indent=2)
        log(f"💾 Batch salvo: {batch_path}")

    # Exit code baseado em resultados
    if not geradas and falhadas:
        log("❌ Nenhuma apostila gerada — a sair com erro")
        sys.exit(1)

    return geradas


if __name__ == "__main__":
    main()
