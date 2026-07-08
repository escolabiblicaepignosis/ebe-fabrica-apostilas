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
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ──────────────────────────────────────────────────────────────
# CONFIGURAÇÃO
# ──────────────────────────────────────────────────────────────
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
MAPA_PATH = os.path.join(os.path.dirname(__file__), "mapa_apostilas.json")
ESTADO_PATH = os.path.join(os.path.dirname(__file__), "estado.json")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")
BATCH_SIZE = int(os.environ.get("BATCH_SIZE", "5"))
MODEL_NAME = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
DELAY_ENTRE_APOSTILAS = 90  # segundos (respeitar limite gratuito ~15 RPM)

# Paleta EBE
COR_PRIMARIA   = RGBColor(0x1B, 0x3A, 0x5C)
COR_SECUNDARIA = RGBColor(0x2E, 0x7D, 0x4F)
COR_TERCIARIA  = RGBColor(0xC9, 0xA1, 0x4B)
COR_TEXTO      = RGBColor(0x1A, 0x1A, 0x1A)
COR_CITACAO    = RGBColor(0x55, 0x55, 0x55)
HEX_PRIMARIA   = "1B3A5C"
HEX_SECUNDARIA = "2E7D4F"
FONTE          = "Garamond"

# Diretório dos logos
ASSETS_DIR = os.path.join(os.path.dirname(__file__), "..", "src", "estilos", "_assets")
LOGO_PATH  = os.path.join(ASSETS_DIR, "logo_ebe.png")


# ──────────────────────────────────────────────────────────────
# UTILITÁRIOS DE ESTILO (compatíveis com _estilos.py)
# ──────────────────────────────────────────────────────────────
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
              color=None, alignment=None, indent_cm=0, space_before=0, space_after=6):
    p = doc.add_paragraph()
    if alignment:
        p.alignment = alignment
    if indent_cm:
        p.paragraph_format.left_indent = Cm(indent_cm)
    if space_before:
        p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
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


# ──────────────────────────────────────────────────────────────
# CONSTRUIR DOCUMENTO DOCX A PARTIR DO CONTEÚDO MARKDOWN
# ──────────────────────────────────────────────────────────────
def slugify(texto):
    """Gera um nome de ficheiro seguro."""
    texto = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")
    texto = re.sub(r'[^\w\s-]', '', texto).strip().lower()
    texto = re.sub(r'[\s_]+', '_', texto)
    return texto[:80]


def construir_docx(meta, conteudo_md, output_path):
    """
    Constrói um documento .docx a partir do conteúdo gerado pela IA em markdown.
    """
    doc = Document()
    configurar_estilos_base(doc)
    codigo = meta.get("codigo", "EBE-APO-XXXX")
    cabecalho_rodape(doc, meta.get("titulo", "Apostila"), codigo)

    # ── CAPA ──
    doc.add_paragraph()
    inserir_logo(doc, LOGO_PATH, largura_cm=5.5)

    add_texto(doc, "ἐπίγνωσις  ·  Conhecer a Deus. Viver a Palavra. Manifestar o Reino.",
              font_size=11, italic=True, color=COR_SECUNDARIA,
              alignment=WD_ALIGN_PARAGRAPH.CENTER)

    p = doc.add_paragraph()
    _add_horizontal_line(p, color=HEX_SECUNDARIA, size=8)

    for _ in range(2):
        doc.add_paragraph()

    # Supratítulo (instituto + escola + curso)
    supratitulo = meta.get("supratitulo", "")
    if supratitulo:
        add_texto(doc, supratitulo.upper(), font_size=11, bold=True,
                  color=COR_SECUNDARIA, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    # Título principal
    add_texto(doc, meta.get("titulo", ""), font_size=26, bold=True,
              color=COR_PRIMARIA, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    # Subtítulo
    subtitulo = meta.get("subtitulo", "")
    if subtitulo:
        add_texto(doc, subtitulo, font_size=13, italic=True,
                  color=COR_TEXTO, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    for _ in range(4):
        doc.add_paragraph()

    p = doc.add_paragraph()
    _add_horizontal_line(p, color=HEX_SECUNDARIA, size=4)

    info_linha = meta.get("info_linha", f"Material didáctico oficial · Código {codigo} · 2026")
    add_texto(doc, info_linha, font_size=10, color=COR_CITACAO,
              alignment=WD_ALIGN_PARAGRAPH.CENTER)

    page_break(doc)

    # ── MARCO FILOSÓFICO ──
    for _ in range(5):
        doc.add_paragraph()
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
    doc.add_paragraph()
    add_texto(doc,
        "\u201cAté que todos cheguemos à unidade da fé e ao pleno conhecimento "
        "(ἐπίγνωσις) do Filho de Deus, a homem perfeito, à medida da estatura "
        "completa de Cristo.\u201d\nEfésios 4.13",
        font_size=10, italic=True, color=COR_CITACAO,
        alignment=WD_ALIGN_PARAGRAPH.CENTER)
    page_break(doc)

    # ── CONTEÚDO PRINCIPAL (parse do markdown) ──
    linhas = conteudo_md.strip().split('\n')
    i = 0
    in_table = False
    table_rows = []

    while i < len(linhas):
        linha = linhas[i].rstrip()

        # Ignorar linhas vazias
        if not linha.strip():
            i += 1
            continue

        # Tabelas markdown
        if '|' in linha and linha.strip().startswith('|'):
            if not in_table:
                in_table = True
                table_rows = []
            # Ignorar separadores
            if re.match(r'^\|[\s\-:|]+\|$', linha.strip()):
                i += 1
                continue
            cells = [c.strip() for c in linha.split('|')[1:-1]]
            table_rows.append(cells)
            # Verificar se a próxima linha ainda é tabela
            if i + 1 < len(linhas) and '|' in linhas[i+1] and linhas[i+1].strip().startswith('|'):
                i += 1
                continue
            else:
                # Renderizar tabela
                if table_rows:
                    max_cols = max(len(r) for r in table_rows)
                    table = doc.add_table(rows=len(table_rows), cols=max_cols)
                    table.style = 'Table Grid'
                    for ri, row_data in enumerate(table_rows):
                        for ci, cell_text in enumerate(row_data):
                            if ci < max_cols:
                                cell = table.rows[ri].cells[ci]
                                cell.text = cell_text
                                for paragraph in cell.paragraphs:
                                    for run in paragraph.runs:
                                        run.font.name = FONTE
                                        run.font.size = Pt(10)
                    doc.add_paragraph()
                in_table = False
                table_rows = []
                i += 1
                continue

        # H1 → h1 estilo EBE
        if linha.startswith('# '):
            texto = linha[2:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(6)
            r = p.add_run(texto.upper())
            r.font.name = FONTE; r.font.size = Pt(15)
            r.font.bold = True; r.font.color.rgb = COR_PRIMARIA
            _add_horizontal_line(p, color=HEX_PRIMARIA, size=6)

        # H2
        elif linha.startswith('## '):
            texto = linha[3:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(texto)
            r.font.name = FONTE; r.font.size = Pt(13)
            r.font.bold = True; r.font.color.rgb = COR_PRIMARIA

        # H3
        elif linha.startswith('### '):
            texto = linha[4:].strip()
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(texto)
            r.font.name = FONTE; r.font.size = Pt(11.5)
            r.font.bold = True; r.font.color.rgb = COR_SECUNDARIA

        # Citação blockquote
        elif linha.startswith('> '):
            texto = linha[2:].strip()
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(1.5)
            p.paragraph_format.right_indent = Cm(1.0)
            r = p.add_run(texto)
            r.font.name = FONTE; r.font.size = Pt(11)
            r.font.italic = True; r.font.color.rgb = COR_CITACAO

        # Lista com bullet
        elif linha.startswith('- ') or linha.startswith('* '):
            texto = linha[2:].strip()
            # Remover markdown bold/italic
            texto = re.sub(r'\*\*(.+?)\*\*', r'\1', texto)
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

        # Lista numerada
        elif re.match(r'^\d+\.\s', linha):
            m = re.match(r'^(\d+)\.\s+(.+)', linha)
            if m:
                num = m.group(1)
                texto = m.group(2).strip()
                texto = re.sub(r'\*\*(.+?)\*\*', r'\1', texto)
                texto = re.sub(r'\*(.+?)\*', r'\1', texto)
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.left_indent = Cm(0.8)
                p.paragraph_format.first_line_indent = Cm(-0.5)
                r = p.add_run(f"{num}. ")
                r.font.name = FONTE; r.font.size = Pt(12)
                r.font.bold = True; r.font.color.rgb = COR_SECUNDARIA
                r2 = p.add_run(texto)
                r2.font.name = FONTE; r2.font.size = Pt(12)

        # Linha horizontal
        elif linha.strip() in ('---', '***', '___'):
            p = doc.add_paragraph()
            _add_horizontal_line(p, color=HEX_SECUNDARIA, size=4)

        # Parágrafo normal
        else:
            texto = linha.strip()
            # Processar inline markdown
            # Bold
            texto_clean = re.sub(r'\*\*(.+?)\*\*', r'\1', texto)
            texto_clean = re.sub(r'\*(.+?)\*', r'\1', texto_clean)

            has_bold = '**' in texto
            has_italic = '*' in texto and '**' not in texto

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # Se linha toda bold (tipo título de secção)
            if texto.startswith('**') and texto.endswith('**'):
                inner = texto[2:-2]
                r = p.add_run(inner)
                r.font.name = FONTE; r.font.size = Pt(12)
                r.font.bold = True; r.font.color.rgb = COR_PRIMARIA
            else:
                r = p.add_run(texto_clean)
                r.font.name = FONTE; r.font.size = Pt(12)

        i += 1

    # ── SELO FINAL ──
    doc.add_paragraph()
    p = doc.add_paragraph()
    _add_horizontal_line(p, color=HEX_SECUNDARIA, size=4)
    doc.add_paragraph()
    add_texto(doc, "ESCOLA BÍBLICA EPIGNÓSIS", font_size=11, bold=True,
              color=COR_PRIMARIA, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_texto(doc, "Conhecer a Deus. Viver a Palavra. Manifestar o Reino.",
              font_size=10, italic=True, color=COR_SECUNDARIA,
              alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_texto(doc, "Soli Deo Gloria", font_size=9, italic=True,
              color=COR_CITACAO, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    doc.save(output_path)
    return output_path


# ──────────────────────────────────────────────────────────────
# GEMINI — PROMPT E GERAÇÃO
# ──────────────────────────────────────────────────────────────
SYSTEM_PROMPT = """\
Você é um teólogo, pedagogo e editor cristão sénior da Escola Bíblica Epignósis (EBE).
A sua tarefa é redigir o conteúdo integral de uma apostila académica bíblica em português europeu/Angola (pt-PT).

REGRAS OBRIGATÓRIAS:
1. Use APENAS a versão bíblica Almeida Revista e Corrigida (ARC) para todas as citações.
2. O conteúdo deve ter entre 15 a 20 páginas (aproximadamente 6.000 a 8.000 palavras).
3. O estilo deve ser académico formal, mas acessível e devocional.
4. Todas as citações bíblicas devem incluir referência completa (livro, capítulo, versículo).
5. NÃO use linguagem inclusiva moderna — siga o padrão da ARC.
6. O conteúdo deve ser doutrinariamente consistente com o protestante reformado/evangélico.
7. Cada secção deve ter substância real — NÃO usar placeholder ou "[conteúdo aqui]".
8. O grego/hebraico pode ser usado quando relevante (com transliteração).
9. O conteúdo deve ser original e profundo, digno de um seminário teológico.
10. Escreva em Markdown formatado (## para títulos, ** para negrito, etc.).

ESTRUTURA OBRIGATÓRIA (nesta ordem):
- FICHA TÉCNICA (título, curso, módulo, escola, instituto, nível, autor, versão, código)
- SUMÁRIO
- APRESENTAÇÃO DA APOSTILA (1 parágrafo)
- OBJECTIVOS DE APRENDIZAGEM (nas 4 dimensões: Conhecer 40%, Crer 20%, Viver 20%, Servir 20%)
- VERSÍCULO-CHAVE (com referência)
- TEXTO-BASE PARA LEITURA (passagem completa com referência, ARC)
- 1. INTRODUÇÃO (2-3 parágrafos)
- 2. DESENVOLVIMENTO DO CONCEITO CENTRAL (com 2.1 a 2.5, cada um com parágrafos densos)
  - 2.1 Fundamentos bíblicos (mínimo 3 passagens com exegese)
  - 2.2-2.4 Desenvolvimento temático profundo (com sub-tópicos)
  - 2.5 Quadro de Destaque (tabela ou lista resumo)
- 3. APLICAÇÃO PRÁTICA (5 áreas: pessoal, família, igreja, sociedade, ministério)
- 4. SÍNTESE E CONCLUSÃO (2-3 parágrafos)
- EXERCÍCIOS DE REVISÃO (3 blocos: I — Compreensão 5 perguntas, II — Reflexão pessoal 3 perguntas, III — Ministério e serviço 2 perguntas)
- ESTUDO BÍBLICO COMPLEMENTAR (passagem completa + 5 perguntas exegéticas)
- PARA A PRÓXIMA APOSTILA (2 perguntas de antecipação)
- GLOSSÁRIO (5-10 termos com definição)
- BIBLIOGRAFIA RECOMENDADA (5-8 referências reais: ARC, FEE, STUART, BERKHOF, LOPES, etc.)
- ANOTAÇÕES PESSOAIS (secção em branco)
"""


def gerar_conteudo_gemini(meta):
    """Chama a API Gemini para gerar o conteúdo da apostila."""
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel(
        model_name=MODEL_NAME,
        system_instruction=SYSTEM_PROMPT,
    )

    prompt = f"""\
Gere o conteúdo integral da seguinte apostila da Escola Bíblica Epignósis:

- Nível: {meta['nivel']}
- Instituto: {meta['instituto']}
- Escola: {meta['escola']}
- Curso: {meta['curso']} (Carga horária: {meta.get('carga_horaria', 'N/D')})
- Módulo: {meta['modulo']}
- Apostila N.º: {meta['numero_apostila']} — {meta['titulo']}
- Código: {meta['codigo']}

O conteúdo deve ser EXAUSTIVO, PROFUNDO e SEGUINDO EXACTAMENTE a estrutura obrigatória descrita no sistema.
Gere em formato Markdown. Use ## para títulos principais, ### para subtítulos, - para listas, ** para negrito, > para citações.
Não inclua o cabeçalho da capa nem o marco filosófico — apenas o conteúdo a partir da FICHA TÉCNICA até ANOTAÇÕES PESSOAIS.
"""

    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = model.generate_content(
                prompt,
                generation_config=genai.GenerationConfig(
                    temperature=0.7,
                    max_output_tokens=16384,
                    top_p=0.9,
                ),
            )
            texto = response.text
            # Limpar delimitadores de código se existirem
            texto = re.sub(r'^```(?:markdown)?\s*\n?', '', texto, flags=re.MULTILINE)
            texto = re.sub(r'\n?```\s*$', '', texto, flags=re.MULTILINE)
            return texto.strip()
        except Exception as e:
            print(f"  ⚠️  Tentativa {attempt+1}/{max_retries} falhou: {e}")
            if attempt < max_retries - 1:
                wait = 30 * (attempt + 1)
                print(f"  ⏳ Aguardando {wait}s antes de tentar novamente...")
                time.sleep(wait)
            else:
                raise


# ──────────────────────────────────────────────────────────────
# ESTADO / PROGRESSO
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
    """Retorna as próximas apostilas a gerar."""
    concluidas = set(estado.get("concluidas", []))
    pendentes = []

    for nivel in mapa["niveis"]:
        nivel_nome = nivel["id"]
        for inst in nivel["institutos"]:
            inst_nome = inst["nome"]
            for escola in inst["escolas"]:
                esc_nome = escola["nome"]
                for curso in escola["cursos"]:
                    curso_nome = curso["nome"]
                    carga = curso.get("carga_horaria", "N/D")
                    for modulo in curso["modulos"]:
                        mod_nome = modulo["nome"]
                        mod_num = modulo["numero"]
                        for apo in modulo["apostilas"]:
                            apo_id = apo["id"]
                            if apo_id not in concluidas:
                                pendentes.append({
                                    "id": apo_id,
                                    "titulo": apo["titulo"],
                                    "nivel": nivel_nome,
                                    "instituto": inst_nome,
                                    "escola": esc_nome,
                                    "curso": curso_nome,
                                    "carga_horaria": carga,
                                    "modulo": f"{mod_num} — {mod_nome}",
                                    "numero_apostila": apo_id.split("-")[1].lstrip("0") or "1",
                                    "codigo": apo_id,
                                    "mod_num": mod_num,
                                    "nivel_nome_completo": nivel["nome"],
                                })

    return pendentes[:batch_size]


# ──────────────────────────────────────────────────────────────
# PIPELINE PRINCIPAL
# ──────────────────────────────────────────────────────────────
def gerar_apostila(meta):
    """Gera uma apostila completa: conteúdo IA → .docx."""
    print(f"\n  📝 Gerando conteúdo via Gemini ({MODEL_NAME})...")
    conteudo = gerar_conteudo_gemini(meta)

    if not conteudo or len(conteudo) < 500:
        raise RuntimeError("Conteúdo gerado muito curto ou vazio")

    print(f"  📄 Conteúdo gerado: {len(conteudo)} caracteres")

    # Montar metadados para o docx
    supratitulo = f"{meta['instituto']}"
    subtitulo = f"{meta['escola']}  ·  Curso «{meta['curso']}»  ·  Módulo {meta['modulo']}"
    info = f"Material didáctico oficial · Código {meta['codigo']} · 2026"

    docx_meta = {
        "titulo": meta["titulo"],
        "supratitulo": supratitulo,
        "subtitulo": subtitulo,
        "codigo": meta["codigo"],
        "info_linha": info,
    }

    # Criar pasta de output
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # Nome do ficheiro
    fname = f"{meta['codigo']}_{slugify(meta['titulo'])}.docx"
    output_path = os.path.join(OUTPUT_DIR, fname)

    print(f"  🏗️  Construindo .docx...")
    construir_docx(docx_meta, conteudo, output_path)

    # Verificar tamanho
    size_kb = os.path.getsize(output_path) / 1024
    print(f"  ✅ Ficheiro criado: {fname} ({size_kb:.0f} KB)")

    return output_path, fname


def main():
    if not GEMINI_API_KEY:
        print("❌ GEMINI_API_KEY não definida. Configure a variável de ambiente.")
        sys.exit(1)

    print("=" * 60)
    print("  ESCOLA BÍBLICA EPIGNÓSIS — Gerador de Apostilas")
    print(f"  Modelo: {MODEL_NAME} | Batch: {BATCH_SIZE}")
    print("=" * 60)

    mapa = carregar_mapa()
    estado = carregar_estado()

    total_mapa = mapa["total_apostilas"]
    total_feitas = len(estado.get("concluidas", []))
    print(f"\n  📊 Progresso: {total_feitas}/{total_mapa} apostilas geradas")

    proximas = obter_proximas_apostilas(mapa, estado, BATCH_SIZE)

    if not proximas:
        print("\n  🎉 Todas as apostilas já foram geradas!")
        return

    print(f"\n  📋 Próximas {len(proximas)} apostilas:")
    for apo in proximas:
        print(f"     • {apo['codigo']} — {apo['titulo']}")

    geradas = []
    falhadas = []

    for idx, meta in enumerate(proximas, 1):
        print(f"\n{'─' * 60}")
        print(f"  [{idx}/{len(proximas)}] {meta['codigo']} — {meta['titulo']}")
        print(f"  Nível: {meta['nivel']} | Instituto: {meta['instituto']}")
        print(f"  Escola: {meta['escola']}")
        print(f"  Curso: {meta['curso']} | Módulo: {meta['modulo']}")

        try:
            output_path, fname = gerar_apostila(meta)
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
                "caminho_drive": "",  # será preenchido pelo upload
            })

            # Atualizar estado incrementalmente
            estado["concluidas"].append(meta["id"])
            estado["total_gerado"] = len(estado["concluidas"])
            salvar_estado(estado)

        except Exception as e:
            print(f"  ❌ ERRO: {e}")
            falhadas.append({"id": meta["id"], "erro": str(e)})
            if meta["id"] not in estado.get("falhadas", []):
                estado.setdefault("falhadas", []).append(meta["id"])

        # Delay entre apostilas para respeitar rate limit
        if idx < len(proximas):
            print(f"\n  ⏳ Aguardando {DELAY_ENTRE_APOSTILAS}s (rate limit)...")
            time.sleep(DELAY_ENTRE_APOSTILAS)

    # Resumo final
    print(f"\n{'=' * 60}")
    print(f"  RESUMO DO BATCH")
    print(f"  ✅ Geradas com sucesso: {len(geradas)}")
    print(f"  ❌ Falhadas: {len(falhadas)}")
    print(f"  📊 Total acumulado: {len(estado.get('concluidas', []))}/{total_mapa}")
    print(f"{'=' * 60}")

    # Salvar lista de geradas para o upload
    if geradas:
        batch_path = os.path.join(os.path.dirname(ESTADO_PATH), "ultimo_batch.json")
        with open(batch_path, 'w', encoding='utf-8') as f:
            json.dump(geradas, f, ensure_ascii=False, indent=2)
        print(f"\n  💾 Lista do batch salva em: {batch_path}")

    return geradas


if __name__ == "__main__":
    main()
