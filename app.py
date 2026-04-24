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

def get_pagamentos(cnpj_limpo, mes_num, ano):
    """Busca pagamentos via API oficial do Portal da Transparência."""
    ultimo_dia = calendar.monthrange(int(ano), mes_num)[1]
    data_ini = f'01/{mes_num:02d}/{ano}'
    data_fim = f'{ultimo_dia:02d}/{mes_num:02d}/{ano}'

    api_key = st.secrets.get("TRANSPARENCIA_API_KEY", "")

    headers = {
        'chave-api-dados': api_key,
        'Accept': 'application/json',
    }

    pagamentos = []
    pagina = 1
    ultimo_status = None
    ultimo_erro = None

    while True:
        params = {
            'codigoPessoa': cnpj_limpo,
            'dataInicial': data_ini,
            'dataFinal': data_fim,
            'fase': 'PAG',
            'pagina': pagina,
        }
        try:
            r = requests.get(
                'https://api.portaldatransparencia.gov.br/api-de-dados/despesas/documentos-por-favorecido',
                params=params,
                headers=headers,
                timeout=30,
            )
            ultimo_status = r.status_code
            if r.status_code != 200:
                ultimo_erro = r.text[:300]
                break
            data = r.json()
            if not isinstance(data, list) or len(data) == 0:
                break
            for item in data:
                doc_num   = (item.get('documento') or {}).get('codigoResumido', '') or item.get('codigoDocumento', '')
                data_pgto = item.get('dataDocumento', item.get('data', ''))
                valor_raw = item.get('valorDocumento', item.get('valor', 0))
                try:
                    v = float(valor_raw)
                    pagamentos.append((doc_num, data_pgto, formatar_valor(v), v))
                except Exception:
                    pass
            if len(data) < 500:
                break
            pagina += 1
        except Exception as e:
            ultimo_erro = str(e)
            break

    return pagamentos, ultimo_status, ultimo_erro

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
        template_tr = copy.deepcopy(table.rows[1]._tr)
        total_tr    = table.rows[-1]._tr

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

        for doc_num, data, valor, _ in pagamentos:
            total_tr.addprevious(make_row(doc_num, data, valor))

        merged_tc = total_tr.findall(qn('w:tc'))[1]
        for r in merged_tc.find(qn('w:p')).findall(qn('w:r')):
            t = r.find(qn('w:t'))
            if t is not None:
                t.text = f"R$ {total_str}"
        break

    return doc

def nome_saida(nome_entrada, mes_nome_arquivo):
    """Deriva nome do arquivo de saída substituindo o mês."""
    base = os.path.splitext(nome_entrada)[0]
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

st.info(
    "📌 **Antes de enviar o documento, verifique:**\n\n"
    "O arquivo deve conter o **CNPJ da empresa** no formato `XX.XXX.XXX/XXXX-XX` "
    "(geralmente no campo *OCS* do relatório). "
    "O programa usa o CNPJ para buscar os pagamentos automaticamente no Portal da Transparência. "
    "Se o documento tiver apenas o nome da empresa, o relatório **não poderá ser gerado**."
)

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
        st.error(
            "❌ **CNPJ não encontrado no documento.**\n\n"
            "O programa precisa do CNPJ da empresa (formato `XX.XXX.XXX/XXXX-XX`) "
            "para buscar os pagamentos no Portal da Transparência. "
            "Adicione o CNPJ ao documento e tente novamente."
        )
        st.stop()

    st.success(f"✅ CNPJ: **{cnpj}**")

    # PASSO 2 — Pagamentos via API oficial
    status.info(f"🌐 Buscando pagamentos de {mes_selecionado}/{ano} no Portal da Transparência...")
    pagamentos, api_status, api_erro = get_pagamentos(limpar_cnpj(cnpj), mes_num, ano)
    progress.progress(65)

    if len(pagamentos) == 0:
        msg = f"⚠️ Nenhum pagamento encontrado para {mes_selecionado}/{ano}. O relatório será gerado com tabela vazia."
        if api_status and api_status != 200:
            msg += f"\n\n🔍 **Debug:** API retornou status **{api_status}**."
            if api_erro:
                msg += f" Resposta: `{api_erro}`"
        elif api_status is None:
            msg += "\n\n🔍 **Debug:** Não foi possível conectar à API (timeout ou erro de rede)."
        st.warning(msg)
        total_str = "0,00"
    else:
        total_str = calcular_total(pagamentos)
        st.info(f"📋 {len(pagamentos)} pagamento(s) | Total: **R$ {total_str}**")

    # PASSO 3 — Gerar documento
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

    if pagamentos:
        st.markdown("### Pagamentos incluídos:")
        import pandas as pd
        rows = [(d, dt, v) for d, dt, v, _ in pagamentos]
        rows.append(("TOTAL PAGO", "", f"R$ {total_str}"))
        df = pd.DataFrame(rows, columns=["Documento", "Data", "Valor"])
        st.dataframe(df, hide_index=True, use_container_width=True)

st.markdown("---")
st.caption("Dados obtidos do Portal da Transparência do Governo Federal · UG 167399")
