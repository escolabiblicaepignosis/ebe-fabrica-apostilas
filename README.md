<p align="center">
  <img src="assets/logos/logo_ebe.png" alt="Escola Bíblica Epignósis" width="260">
</p>

<h1 align="center">Fábrica de Apostilas — EBE</h1>

<p align="center">
  <strong>Sistema automatizado de geração de documentos institucionais</strong><br>
  Escola Bíblica Epignósis (EBE)
</p>

<p align="center">
  <img src="https://img.shields.io/badge/idioma--PT--PT-blue" alt="pt-PT">
  <img src="https://img.shields.io/badge/versão_bíblica-ARC-green" alt="ARC">
  <img src="https://img.shields.io/badge/python-3.11+-yellow" alt="Python 3.11+">
  <img src="https://img.shields.io/badge/Gemini--AI-automatizado-purple" alt="Gemini AI">
  <img src="https://img.shields.io/badge/apostilas-1_029-orange" alt="1029 Apostilas">
</p>

---

## 📌 Sobre o projecto

Este repositório contém o **sistema documental completo** da Escola Bíblica Epignósis, com duas componentes:

1. **Documentos institucionais** — apostilas, certificados, formulários, manuais e compêndios gerados por scripts Python.
2. **Fábrica automatizada de apostilas** — sistema de geração em massa de **1.029 apostilas** usando **Gemini AI**, com upload automático para **Google Drive** via **GitHub Actions**.

Todos os documentos seguem a **identidade visual oficial** da EBE (tipografia Garamond, paleta azul-marinho `#1B3A5C` + verde `#2E7D4F` + dourado `#C9A14B`), são escritos em **português europeu/Angola (pt-PT)** e utilizam a **Almeida Revista e Corrigida (ARC)** como versão bíblica de referência.

---

## 📁 Estrutura do repositório

```
ebe-fabrica-apostilas/
│
├── 📂 src/                          # Código-fonte
│   ├── estilos/                     # Módulo de estilos comum (_estilos.py)
│   │   └── _assets/                 # Logotipos (symlinks → assets/logos)
│   ├── gerador/                     # Scripts individuais de geração de documentos
│   │   ├── apostila_piloto.py
│   │   ├── compendio.py
│   │   ├── doc1_missao_visao_valores.py ... doc8_pre_requisitos.py
│   │   ├── manual_aluno.py / manual_docente.py
│   │   └── modelo_*.py / modelos_certificados.py
│   └── utilizadores/                # (reservado)
│
├── 📂 scripts/                      # Sistema de automação
│   ├── gerar_apostila.py            # 🤖 Gerador com Gemini AI
│   ├── upload_drive.py              # ☁️ Upload para Google Drive
│   ├── mapa_apostilas.json          # 📋 Mapa das 1.029 apostilas
│   ├── estado.json                  # 📊 Progresso da geração
│   └── ultimo_batch.json            # Último batch gerado (temporário)
│
├── 📂 .github/workflows/
│   └── gerar-apostilas.yml          # ⚙️ GitHub Actions (automático + manual)
│
├── 📂 docs/                         # Documentos gerados
│   ├── institucionais/              # 8 documentos oficiais + compêndios
│   ├── formularios/                 # 6 formulários + histórico escolar
│   ├── modelos/                     # Modelos editáveis + certificados
│   └── Mapa_Completo_Apostilas.*    # Mapa curricular em .docx e .pdf
│
├── 📂 assets/logos/                 # Logotipos oficiais (4 variantes)
│
├── 📂 output/                       # Apostilas geradas (local, .gitignore)
│
├── 📄 README.md                     # Este ficheiro
├── 📄 LEIA-ME.md                    # Catálogo detalhado de documentos
├── 📄 requirements.txt              # Dependências Python
└── 📄 .gitignore
```

---

## 🤖 Sistema de Automação — Geração em Massa

### Como funciona

```
┌─────────────────┐     ┌──────────────┐     ┌──────────────┐
│  GitHub Actions  │────▶│  Gemini AI   │────▶│  Google Drive │
│  (cron 8h/8h)   │     │  (conteúdo)  │     │  (upload)     │
└────────┬────────┘     └──────────────┘     └──────────────┘
         │                                           │
         └──── commit automático ◀───────────────────┘
```

1. O **GitHub Actions** executa automaticamente a cada 8 horas.
2. O script `gerar_apostila.py` consulta o **mapa de 1.029 apostilas** e seleciona as próximas pendentes.
3. Para cada apostila, chama a **API Gemini AI** que gera 15-20 páginas de conteúdo teológico.
4. O conteúdo é formatado num `.docx` seguindo o padrão institucional EBE (capa, marco filosófico, sumário, corpo, glossário, etc.).
5. O script `upload_drive.py` faz upload para o **Google Drive**, organizando em pastas: **Instituto → Escola → Curso → Módulo**.
6. O progresso é guardado em `scripts/estado.json` e committedo automaticamente.

### Limites e configuração

| Recurso | Limite gratuito | Configuração actual |
|---|---|---|
| **GitHub Actions** (privado) | 2.000 min/mês | ~1.080 min/mês (✅) |
| **GitHub Actions** (público) | Ilimitado | — |
| **Gemini 2.5 Flash** | ~15 RPM, ~1M tokens/dia | ~225K tokens/dia (✅) |
| **Google Drive** | 15 GB | Cresce com as apostilas |

### Métricas do batch

| Parâmetro | Valor |
|---|---|
| Apostilas por execução | **3** (configurável) |
| Intervalo entre execuções | **8 horas** |
| Apostilas por dia | **~9** |
| Apostilas por mês | **~270** |
| Tempo total estimado (1.029) | **~4 meses** |
| Páginas por apostila | **15-20** |
| Tokens por apostila | **~25K** |

---

## 🚀 Como usar

### Pré-requisitos

- **Python 3.11** ou superior
- Conta **Google Cloud** com Service Account e Google Drive API activa
- **Gemini API Key** (plano gratuito ou superior)
- Repositório GitHub com os Secrets configurados

### 1. Instalação local

```bash
git clone https://github.com/escolabiblicaepignosis/ebe-fabrica-apostilas.git
cd ebe-fabrica-apostilas

python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### 2. Configurar Secrets no GitHub

Vá a **Settings → Secrets and variables → Actions** e adicione:

| Secret | Descrição |
|---|---|
| `GEMINI_API_KEY` | Chave da API Google Gemini (AI Studio) |
| `GOOGLE_DRIVE_CREDENTIALS` | JSON completo da Service Account do Google Cloud |
| `GOOGLE_DRIVE_FOLDER_ID` | ID da pasta mãe no Google Drive |

### 3. Gerar apostilas localmente (teste)

```bash
export GEMINI_API_KEY="sua-chave-aqui"
export BATCH_SIZE=1  # Apenas 1 para teste
python scripts/gerar_apostila.py
```

### 4. Primeiro teste no GitHub Actions

1. Vá a **Actions → Gerar Apostilas EBE → Run workflow**
2. Defina `batch_size` como **1** para o primeiro teste
3. Selecione o modelo **gemini-2.5-flash**
4. Clique em **Run workflow**
5. Acompanhe o progresso nos logs

### 5. Activar o cron automático

Após o primeiro teste bem-sucedido, o cron já estará activo:
- **`0 */8 * * *`** — executa às 00:00, 08:00 e 16:00 UTC
- Pode ajustar o intervalo editando `.github/workflows/gerar-apostilas.yml`

### 6. Gerar documentos institucionais (scripts individuais)

```bash
# Documentos oficiais
python3 src/gerador/doc1_missao_visao_valores.py
python3 src/gerador/doc2_declaracao_fe.py
# ... (ver lista completa em src/gerador/)

# Compêndio unificado
python3 src/gerador/compendio_merge.py

# Manuais
python3 src/gerador/manual_aluno.py
python3 src/gerador/manual_docente.py
```

---

## 📚 Catálogo de documentos

### 📖 Documentos Institucionais (EBE-DOC-001 a 008)

| Código | Documento | Conteúdo |
|---|---|---|
| EBE-DOC-001 | Identidade Institucional | Missão, Visão, Valores e Filosofia de Ensino |
| EBE-DOC-002 | Declaração de Fé | 16 artigos da fé institucional |
| EBE-DOC-003 | Projecto Pedagógico Oficial | PPO completo |
| EBE-DOC-004 | Regimento Acadêmico | Normas e regulamentos académicos |
| EBE-DOC-005 | Arquitectura Oficial | Estrutura de 7 níveis formativos |
| EBE-DOC-006 | Mapa Oficial de Cursos | Todos os cursos por Instituto |
| EBE-DOC-007 | Duração Oficial dos Cursos | Carga horária de cada curso |
| EBE-DOC-008 | Sistema de Pré-Requisitos | Regras de progressão académica |

### 📝 Formulários da Secretaria Acadêmica (EBE-FRM-001 a 006 + HIS-001)

| Código | Formulário | Descrição |
|---|---|---|
| EBE-FRM-001 | Ficha de Matrícula | Formulário oficial de inscrição (9 secções) |
| EBE-FRM-002 | Plano de Aula | Cronograma das 5 fases da aula + relatório pós-aula |
| EBE-FRM-003 | Declaração de Frequência | Declaração oficial em texto corrido |
| EBE-FRM-004 | Ficha de Avaliação | Avaliação em 4 dimensões (Conhecer·Crer·Viver·Servir) |
| EBE-FRM-005 | Lista de Presença | Chamada semanal para 30 alunos (A4 paisagem) |
| EBE-FRM-006 | Pauta de Avaliação | Lançamento de notas por dimensão e média ponderada |
| EBE-HIS-001 | Histórico Escolar | Registo detalhado por nível (A4 paisagem) |

### 🎓 Mapa das 1.029 Apostilas

O mapa curricular cobre:

| Nível | Nome | Institutos | Cursos |
|---|---|---|---|
| 1 | Discípulo (Conhecer) | Instituto de Formação Cristã, Instituto de Ciências Bíblicas | ~42 cursos |
| 2 | Crescimento (Ser) | Instituto de Ciências Teológicas, Instituto de Formação Espiritual | ~30 cursos |
| 3 | Servir (Ministério) | Instituto Ministerial, Instituto do Reino e Poder, Instituto dos Cinco Ministérios | ~35 cursos |
| 4 | Multiplicação (Reino) | Instituto de Missões, Instituto de Liderança e Multiplicação, Instituto de Pesquisa Bíblica e Excelência | ~24 cursos |

Cada curso contém 3 módulos, e cada módulo contém 3 apostilas (1.029 no total).

---

## 🎨 Identidade Visual

| Elemento | Valor |
|---|---|
| Cor primária | Azul-marinho `#1B3A5C` |
| Cor secundária | Verde `#2E7D4F` |
| Cor terciária (selos) | Dourado `#C9A14B` |
| Tipografia | Garamond |
| Língua | Português europeu/Angola (pt-PT) |
| Versão bíblica | Almeida Revista e Corrigida (ARC) |
| Estilo | Académico formal |

---

## ⚖️ Sistema avaliativo

Todos os documentos de avaliação da EBE seguem o modelo de **4 dimensões**:

| Dimensão | Peso | Descrição |
|---|---|---|
| **Conhecer** | 40% | Conhecimento bíblico e teológico |
| **Crer** | 20% | Convicção pessoal e confissão de fé |
| **Viver** | 20% | Aplicação prática na vida cristã |
| **Servir** | 20% | Exercício ministerial e serviço ao próximo |

---

## 🏛️ Arquitectura formativa

```
APOSTILA (1-3h)
  └─▶ MÓDULO (4-10h)
        └─▶ CURSO (20-60h)
              └─▶ ESCOLA (60-180h)
                    └─▶ INSTITUTO (200-600h)
                          └─▶ PROGRAMA DE FORMAÇÃO
                                └─▶ DIPLOMA COMPLETA (2.200-2.400h)
```

---

## 🤝 Contribuir

1. Faça um fork do repositório
2. Crie uma branch (`git checkout -b feature/nova-apostila`)
3. Faça as suas alterações
4. Teste localmente (`python3 script.py`)
5. Submeta um Pull Request

---

<p align="center">
  <em>"Acreditamos que o verdadeiro conhecimento de Deus transforma a mente pela
  verdade das Escrituras, o coração pela acção do Espírito Santo e a vida pelo
  compromisso de viver e anunciar o Evangelho de Jesus Cristo."</em>
  <br><br>
  — Marco Filosófico, Escola Bíblica Epignósis
  <br><br>
  <strong>Soli Deo Gloria.</strong>
</p>
