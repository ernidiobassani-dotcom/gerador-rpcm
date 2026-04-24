import streamlit as st
import requests
import re
import zipfile
import shutil
import copy
import io
import calendar
import tempfile
import os
from lxml import etree
from docx import Document
from docx.oxml.ns import qn

st.set_page_config(
    page_title="Gerador de RPCM",
    page_icon="📋",
    layout="centered"
)

st.title("📋 Gerador de RPCM")
st.markdown("**Relatório de Prestação de Contas Mensal — Contratos de Credenciamento**")
st.markdown("---")

MESES = {
    1:  ("JAN", "JANEIRO"),
    2:  ("FEV", "FEVEREIRO"),
    3:  ("MAR", "MARCO"),
    4:  ("ABR", "ABRIL"),
    5:  ("MAI", "MAIO"),
    6:  ("JUN", "JUNHO"),
    7:  ("JUL", "JULHO"),
    8:  ("AGO", "AGOSTO"),
    9:  ("SET", "SETEMBRO"),
    10: ("OUT", "OUTUBRO"),
    11: ("NOV", "NOVEMBRO"),
    12: ("DEZ", "DEZEMBRO"),
}
MESES_LISTA = [
    "Janeiro","Fevereiro","Março","Abril","Maio","Junho",
    "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"
]

# ─── Funções ────────────────────────────────────────────────────────────────

def extrair_cnpj(file_bytes):
    """Extrai CNPJ de qualquer arquivo Word (docx/dotx) sem depender do python-docx."""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            with z.open('word/document.xml') as f:
                content = f.read().decode('utf-8')
        text = re.sub(r'<[^>]+>', ' ', content)
        text = re.sub(r'\s+', ' ', text)
        match = re.search(r'\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}', text)
        return match.group() if match else None
    except Exception:
        return None

def limpar_cnpj(cnpj):
    return re.sub(r'[.\-/]', '', cnpj)

def get_id_interno(cnpj_limpo):
    """Busca o ID interno da empresa seguindo o redirect do Portal."""
    url = f'https://portaldatransparencia.gov.br/pessoa-juridica/{cnpj_limpo}'
    headers = {'User-Agent': 'Mozilla/5.0 (compatible; RPCM-Bot/1.0)'}
    try:
        r = requests.get(url, headers=headers, allow_redirects=True, timeout=20)
        match = re.search(r'[?&]id=(\d+)', r.url)
        if match:
            return match.group(1)
        # Tenta encontrar no corpo da resposta
        match = re.search(r'"id"\s*:\s*"?(\d+)"?', r.text)
        return match.group(1) if match else None
    except Exception:
        return None

def get_pagamentos(id_interno, mes_num, ano):
    """Busca pagamentos no Portal da Transparência."""
    ultimo_dia = calendar.monthrange(int(ano), mes_num)[1]
    di = f'01%2F{mes_num:02d}%2F{ano}'
    df = f'{ultimo_dia:02d}%2F{mes_num:02d}%2F{ano}'

    params = (
        f'paginacaoSimples=true&tamanhoPagina=100&offset=0&direcaoOrdenacao=asc'
        f'&colunasSelecionadas=data%2CdocumentoResumido%2Cvalor%2Cfavorecido'
        f'&de={di}&ate={df}'
        f'&favorecido={id_interno}&orgaos=UG167399&faseDespesa=3'
        f'&ordenarPor=data&direcao=asc'
    )

    headers_json = {
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'X-Requested-With': 'XMLHttpRequest',
        'Referer': 'https://portaldatransparencia.gov.br/',
    }
    headers_html = {
        'User-Agent': 'Mozilla/5.0',
        'Accept': 'text/html,application/xhtml+xml',
    }

    # 1. Tenta endpoint JSON (AJAX interno do portal)
    for endpoint in [
        f'https://portaldatransparencia.gov.br/despesas/favorecido/listar?{params}',
        f'https://portaldatransparencia.gov.br/api-de-dados/despesas/documentos?{params}',
    ]:
        try:
            r = requests.get(endpoint, headers=headers_json, timeout=20)
            if r.status_code == 200 and 'json' in r.headers.get('Content-Type', ''):
                resultado = parse_json(r.json())
                if resultado is not None:
                    return resultado
        except Exception:
            continue

    # 2. Fallback: HTML scraping
    url_html = f'https://portaldatransparencia.gov.br/despesas/favorecido?{params}'
    try:
        r = requests.get(url_html, headers=headers_html, timeout=20)
        resultado = parse_html(r.text)
        if resultado is not None:
            return resultado
    except Exception:
        pass

    return None

def parse_json(data):
    """Parseia resposta JSON do portal."""
    items = data if isinstance(data, list) else data.get('data', data.get('resultado', []))
    if not isinstance(items, list):
        return None
    pagamentos = []
    for item in items:
        doc_num  = item.get('documentoResumido') or item.get('documento', '')
        data_pgto = item.get('data', '')
        valor_raw = item.get('valor', 0)
        if isinstance(valor_raw, (int, float)):
            valor_fmt = formatar_valor(valor_raw)
            pagamentos.append((doc_num, data_pgto, valor_fmt, float(valor_raw)))
        else:
            # Tenta converter string
            try:
                v = float(str(valor_raw).replace('.','').replace(',','.'))
                pagamentos.append((doc_num, data_pgto, f'R$ {valor_raw}', v))
            except Exception:
                pass
    return pagamentos

def parse_html(html):
    """Extrai pagamentos do HTML do portal."""
    # Remove tags e normaliza
    text = re.sub(r'<[^>]+>', ' ', html)
    text = re.sub(r'\s+', ' ', text)

    pagamentos = []

    # Padrão: data + documento (OB ou DF) + valor no formato brasileiro
    pattern = r'(\d{2}/\d{2}/\d{4})\s+(20\d{2}(?:OB|DF)\d+).*?R\$?\s*([\d]+\.[\d]{3},\d{2}|[\d]+,\d{2})'
    matches = re.findall(pattern, text)

    for data_pgto, doc_num, valor_str in matches:
        try:
            valor_num = float(valor_str.replace('.', '').replace(',', '.'))
            pagamentos.append((doc_num, data_pgto, f'R$ {valor_str}', valor_num))
        except Exception:
            continue

    # Fallback: tenta padrão sem R$
    if not pagamentos:
        pattern2 = r'(\d{2}/\d{2}/\d{4})\s+(20\d{2}(?:OB|DF)\d+)[^0-9]+([\d]+\.?[\d]*,\d{2})'
        matches2 = re.findall(pattern2, text)
        for data_pgto, doc_num, valor_str in matches2:
            try:
                valor_num = float(valor_str.replace('.', '').replace(',', '.'))
                pagamentos.append((doc_num, data_pgto, f'R$ {valor_str}', valor_num))
            except Exception:
                continue

    return pagamentos if pagamentos else []

def formatar_valor(v):
    """Formata float para padrão brasileiro: R$ 1.234,56"""
    return 'R$ {:,.2f}'.format(v).replace(',', 'X').replace('.', ',').replace('X', '.')

def calcular_total(pagamentos):
    total = sum(v for _, _, _, v in pagamentos)
    return '{:,.2f}'.format(total).replace(',', 'X').replace('.', ',').replace('X', '.')

def abrir_documento(file_bytes, filename):
    """Abre .docx ou .dotx com python-docx, convertendo .dotx se necessário."""
    ext = os.path.splitext(filename)[1].lower()
    tmp = tempfile.NamedTemporaryFile(suffix=ext, delete=False)
    tmp.write(file_bytes)
    tmp.close()
    try:
        return Document(tmp.name)
    except ValueError:
        # .dotx: corrigir content type
        dst = tmp.name.replace(ext, '.docx')
        shutil.copy2(tmp.name, dst)
        with zipfile.ZipFile(dst, 'r') as zin:
            files = {n: zin.read(n) for n in zin.namelist()}
        ct = files['[Content_Types].xml'].decode('utf-8')
        ct = ct.replace(
            'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
        )
        files['[Content_Types].xml'] = ct.encode('utf-8')
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for n, d in files.items():
                zout.writestr(n, d)
        return Document(dst)

def atualizar_documento(doc, mes_abrev, ano, pagamentos, total_str):
    """Atualiza mês/ano no cabeçalho e substitui a tabela de pagamentos."""
    W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

    # 1. Cabeçalho
    for para in doc.paragraphs:
        for run in para.runs:
            if re.search(r'\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\b', run.text):
                run.text = re.sub(
                    r'\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\b',
                    mes_abrev, run.text)
            if re.search(r'\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)/\d{4}\b', run.text):
                run.text = re.sub(r'(?<=\/)\d{4}', ano, run.text)

    # 2. Tabela de pagamentos
    for table in doc.tables:
        if 'DOCUMENTO' not in table.rows[0].cells[0].text:
            continue

        tbl = table._tbl
        template_tr = copy.deepcopy(table.rows[1]._tr)   # linha modelo
        total_tr    = table.rows[-1]._tr                  # linha TOTAL PAGO

        # Remover linhas de dados existentes
        for row in list(table.rows[1:-1]):
            tbl.remove(row._tr)

        def make_row(doc_num, data, valor):
            new_tr = copy.deepcopy(template_tr)
            tcs = new_tr.findall(qn('w:tc'))
            for tc, texto in zip(tcs, [doc_num, data, valor]):
                para   = tc.find(qn('w:p'))
                old_r  = para.find(qn('w:r'))
                rPr    = old_r.find(qn('w:rPr')) if old_r is not None else None
                for r in para.findall(qn('w:r')):
                    para.remove(r)
                new_r = etree.SubElement(para, f'{{{W}}}r')
                if rPr is not None:
                    new_r.insert(0, copy.deepcopy(rPr))
                t = etree.SubElement(new_r, f'{{{W}}}t')
                t.text = texto
                if texto.startswith(' ') or texto.endswith(' '):
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return new_tr

        # Inserir linhas (mesmo que vazio, não insere nada — tabela só com cabeçalho+total)
        for doc_num, data, valor, _ in pagamentos:
            total_tr.addprevious(make_row(doc_num, data, valor))

        # Atualizar TOTAL PAGO
        merged_tc = total_tr.findall(qn('w:tc'))[1]
        for r in merged_tc.find(qn('w:p')).findall(qn('w:r')):
            t = r.find(qn('w:t'))
            if t is not None:
                t.text = f'R$ {total_str}'
        break

    return doc

def nome_saida(nome_entrada, mes_nome_arquivo):
    """Deriva nome do arquivo de saída substituindo o mês."""
    base = os.path.splitext(nome_entrada)[0]
    # Substituir qualquer nome de mês (em PT) no nome do arquivo
    meses_re = '|'.join([
        'JANEIRO','FEVEREIRO','MARCO','MARÇO','ABRIL','MAIO','JUNHO',
        'JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO',
        'JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'
    ])
    novo = re.sub(
        rf'\b({meses_re})\b', mes_nome_arquivo,
        base.upper(), flags=re.IGNORECASE
    )
    return novo + '.docx'

# ─── Interface ──────────────────────────────────────────────────────────────

uploaded = st.file_uploader(
    "📎 Selecione o documento base (.docx ou .dotx)",
    type=["docx", "dotx"]
)

col1, col2 = st.columns(2)
with col1:
    mes_selecionado = st.selectbox("Mês de referência", MESES_LISTA)
with col2:
    ano_input = st.text_input("Ano", value="2025", max_chars=4)

gerar = st.button(
    "📄 Gerar Relatório",
    type="primary",
    disabled=(uploaded is None or not ano_input.isdigit())
)

# ─── Lógica principal ───────────────────────────────────────────────────────

if gerar and uploaded:
    mes_num = MESES_LISTA.index(mes_selecionado) + 1
    mes_abrev, mes_nome_arq = MESES[mes_num]
    ano = ano_input.strip()

    progress = st.progress(0)
    status   = st.empty()
    file_bytes = uploaded.read()

    # PASSO 1 — CNPJ
    status.info("🔍 Lendo documento e extraindo CNPJ...")
    cnpj = extrair_cnpj(file_bytes)
    progress.progress(15)

    if not cnpj:
        st.error("❌ CNPJ não encontrado no documento. Verifique se o arquivo está no formato correto.")
        st.stop()

    st.success(f"✅ CNPJ: **{cnpj}**")

    # PASSO 2 — ID interno
    status.info("🌐 Localizando empresa no Portal da Transparência...")
    id_interno = get_id_interno(limpar_cnpj(cnpj))
    progress.progress(35)

    if not id_interno:
        st.error("❌ Empresa não encontrada no Portal da Transparência. Verifique a conexão ou tente mais tarde.")
        st.stop()

    # PASSO 3 — Pagamentos
    status.info(f"📊 Buscando pagamentos de {mes_selecionado}/{ano}...")
    pagamentos = get_pagamentos(id_interno, mes_num, ano)
    progress.progress(65)

    if pagamentos is None:
        st.error("❌ Erro ao buscar pagamentos no Portal da Transparência. Tente novamente em alguns instantes.")
        st.stop()

    if len(pagamentos) == 0:
        st.warning(f"⚠️ Nenhum pagamento encontrado para {mes_selecionado}/{ano}. O relatório será gerado com tabela vazia.")
        total_str = "0,00"
    else:
        total_str = calcular_total(pagamentos)
        st.info(f"📋 {len(pagamentos)} pagamento(s) | Total: **R$ {total_str}**")

    # PASSO 4 — Gerar documento
    status.info("📝 Atualizando documento...")
    doc = abrir_documento(file_bytes, uploaded.name)
    doc = atualizar_documento(doc, mes_abrev, ano, pagamentos, total_str)
    progress.progress(90)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    output_name = nome_saida(uploaded.name, mes_nome_arq)
    progress.progress(100)
    status.success("✅ Documento gerado com sucesso!")

    st.download_button(
        label=f"⬇️ Baixar {output_name}",
        data=buf.read(),
        file_name=output_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )

    # Exibir resumo dos pagamentos
    if pagamentos:
        st.markdown("### Pagamentos incluídos:")
        import pandas as pd
        rows = [(d, dt, v) for d, dt, v, _ in pagamentos]
        rows.append(("TOTAL PAGO", "", f"R$ {total_str}"))
        df = pd.DataFrame(rows, columns=["Documento", "Data", "Valor"])
        st.dataframe(df, hide_index=True, use_container_width=True)

st.markdown("---")
st.caption("Dados obtidos do Portal da Transparência do Governo Federal · UG 167399")
