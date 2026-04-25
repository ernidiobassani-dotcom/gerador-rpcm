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
import subprocess
from datetime import date
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

# ODT namespaces
NS_TEXT  = 'urn:oasis:names:tc:opendocument:xmlns:text:1.0'
NS_TABLE = 'urn:oasis:names:tc:opendocument:xmlns:table:1.0'

# ─── Funções de extração de CNPJ ────────────────────────────────────────────

CNPJ_REGEX = re.compile(r'\d{2}\.\d{3}\.\d{3}\/\d{4}\s*-\s*\d{2}')

def _normalizar_cnpj(valor):
    return re.sub(r'\s+', '', valor)

def extrair_cnpj_zip(file_bytes):
    """Extrai CNPJ varrendo todos os XMLs de um arquivo ZIP (docx, dotx, odt)."""
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
            for nome in z.namelist():
                if not nome.endswith('.xml'):
                    continue
                try:
                    conteudo = z.read(nome).decode('utf-8', errors='ignore')
                    texto = re.sub(r'<[^>]+>', ' ', conteudo)
                    texto = re.sub(r'\s+', ' ', texto)
                    match = CNPJ_REGEX.search(texto)
                    if match:
                        return _normalizar_cnpj(match.group())
                except Exception:
                    continue
    except Exception:
        pass
    return None

def extrair_cnpj_doc(file_bytes):
    """Extrai CNPJ de arquivo .doc (binário legacy) buscando no conteúdo."""
    # Word .doc armazena texto internamente em UTF-16-LE
    try:
        texto = file_bytes.decode('utf-16-le', errors='ignore')
        match = CNPJ_REGEX.search(texto)
        if match:
            return _normalizar_cnpj(match.group())
    except Exception:
        pass
    # Fallback: latin-1
    try:
        texto = file_bytes.decode('latin-1', errors='ignore')
        match = CNPJ_REGEX.search(texto)
        if match:
            return _normalizar_cnpj(match.group())
    except Exception:
        pass
    # Fallback: strings via subprocess
    try:
        tmp = tempfile.NamedTemporaryFile(suffix='.doc', delete=False)
        tmp.write(file_bytes)
        tmp.close()
        result = subprocess.run(['strings', tmp.name], capture_output=True, text=True, timeout=10)
        match = CNPJ_REGEX.search(result.stdout)
        if match:
            return _normalizar_cnpj(match.group())
    except Exception:
        pass
    return None

def extrair_cnpj(file_bytes, ext='.docx'):
    """Extrai CNPJ do documento de acordo com o formato."""
    if ext in ('.docx', '.dotx', '.odt'):
        return extrair_cnpj_zip(file_bytes)
    elif ext == '.doc':
        return extrair_cnpj_doc(file_bytes)
    return None

def limpar_cnpj(cnpj):
    return re.sub(r'[.\-/]', '', cnpj)

# ─── API Portal da Transparência ────────────────────────────────────────────

def get_pagamentos(cnpj_limpo, mes_num, ano):
    """Busca pagamentos via API oficial do Portal da Transparência."""
    api_key = st.secrets.get("TRANSPARENCIA_API_KEY", "")

    headers = {
        'chave-api-dados': api_key,
        'Accept': 'application/json',
    }

    todos = []
    pagina = 1
    ultimo_status = None
    ultimo_erro = None

    while True:
        params = {
            'codigoPessoa': cnpj_limpo,
            'fase': 3,
            'ano': int(ano),
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
            todos.extend(data)
            if len(data) < 500:
                break
            pagina += 1
        except Exception as e:
            ultimo_erro = str(e)
            break

    pagamentos = []
    mes_str = f'{mes_num:02d}'
    for item in todos:
        data_pgto = item.get('data', item.get('dataDocumento', ''))
        data_str  = str(data_pgto)
        mes_ok = False
        if len(data_str) >= 7 and data_str[4:5] == '-' and data_str[5:7] == mes_str:
            mes_ok = True  # ISO: YYYY-MM-DD
        elif len(data_str) >= 5 and data_str[2:3] == '/' and data_str[3:5] == mes_str:
            mes_ok = True  # BR: DD/MM/YYYY
        elif mes_str in data_str:
            mes_ok = True  # fallback
        if not mes_ok:
            continue
        doc_num   = item.get('documentoResumido', '') or item.get('documento', '')
        valor_raw = item.get('valor', item.get('valorDocumento', '0'))
        try:
            if isinstance(valor_raw, str):
                v = float(valor_raw.replace('.', '').replace(',', '.'))
            else:
                v = float(valor_raw)
            pagamentos.append((doc_num, data_pgto, formatar_valor(v), v))
        except Exception:
            pass

    return pagamentos, ultimo_status, ultimo_erro, todos

def formatar_valor(v):
    """Formata float para padrão brasileiro: R$ 1.234,56"""
    return 'R$ {:,.2f}'.format(v).replace(',', 'X').replace('.', ',').replace('X', '.')

def calcular_total(pagamentos):
    total = sum(v for _, _, _, v in pagamentos)
    return '{:,.2f}'.format(total).replace(',', 'X').replace('.', ',').replace('X', '.')

# ─── DOCX / DOTX ────────────────────────────────────────────────────────────

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
    """Atualiza mês/ano no cabeçalho e substitui a tabela de pagamentos (DOCX)."""
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
                t.text = f'R$ {total_str}'
        break

    return doc

# ─── ODT ────────────────────────────────────────────────────────────────────

def _odt_cell_text(cell):
    """Extrai texto de uma célula ODT."""
    partes = []
    for node in cell.iter():
        if node.text:
            partes.append(node.text)
        if node.tail:
            partes.append(node.tail)
    return ''.join(partes).strip()

def _odt_set_cell_text(cell, texto):
    """Define texto de uma célula ODT, preservando atributos de estilo da célula."""
    NS_T = f'{{{NS_TEXT}}}'
    paragrafos = cell.findall(f'{NS_T}p')
    if paragrafos:
        p = paragrafos[0]
        for child in list(p):
            p.remove(child)
        p.text = texto
        for extra in paragrafos[1:]:
            cell.remove(extra)
    else:
        p = etree.SubElement(cell, f'{NS_T}p')
        p.text = texto

def _substituir_mes_ano_odt(tree, mes_abrev, ano):
    """Substitui mês/ano em todos os nós de texto da árvore XML do ODT."""
    meses_re = r'\b(JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\b'
    for node in tree.iter():
        if node.text and re.search(meses_re, node.text):
            node.text = re.sub(meses_re, mes_abrev, node.text)
            node.text = re.sub(r'(?<=/)\d{4}', ano, node.text)
        if node.tail and re.search(meses_re, node.tail):
            node.tail = re.sub(meses_re, mes_abrev, node.tail)
            node.tail = re.sub(r'(?<=/)\d{4}', ano, node.tail)

def atualizar_odt(file_bytes, mes_abrev, ano, pagamentos, total_str):
    """Atualiza ODT: substitui mês/ano e reconstrói tabela de pagamentos."""
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        files = {n: z.read(n) for n in z.namelist()}

    NS_T = f'{{{NS_TEXT}}}'
    NS_TB = f'{{{NS_TABLE}}}'

    content_xml = files.get('content.xml')
    if content_xml is None:
        raise ValueError("ODT inválido: content.xml não encontrado")

    tree = etree.fromstring(content_xml)
    _substituir_mes_ano_odt(tree, mes_abrev, ano)

    for table in tree.iter(f'{NS_TB}table'):
        rows = table.findall(f'{NS_TB}table-row')
        if len(rows) < 2:
            continue
        header_cells = rows[0].findall(f'{NS_TB}table-cell')
        if not header_cells:
            continue
        if 'DOCUMENTO' not in _odt_cell_text(header_cells[0]).upper():
            continue

        template_row = rows[1]
        total_row    = rows[-1]

        for row in rows[1:-1]:
            table.remove(row)

        for doc_num, data, valor, _ in pagamentos:
            new_row = copy.deepcopy(template_row)
            cells = new_row.findall(f'{NS_TB}table-cell')
            if len(cells) >= 3:
                _odt_set_cell_text(cells[0], str(doc_num))
                _odt_set_cell_text(cells[1], str(data))
                _odt_set_cell_text(cells[2], str(valor))
            total_row.addprevious(new_row)

        total_cells = total_row.findall(f'{NS_TB}table-cell')
        if len(total_cells) >= 2:
            _odt_set_cell_text(total_cells[1], f'R$ {total_str}')
        elif len(total_cells) == 1:
            _odt_set_cell_text(total_cells[0], f'R$ {total_str}')
        break

    files['content.xml'] = etree.tostring(
        tree, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    if 'styles.xml' in files:
        styles_tree = etree.fromstring(files['styles.xml'])
        _substituir_mes_ano_odt(styles_tree, mes_abrev, ano)
        files['styles.xml'] = etree.tostring(
            styles_tree, xml_declaration=True, encoding='UTF-8', standalone=True
        )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        if 'mimetype' in files:
            info = zipfile.ZipInfo('mimetype')
            info.compress_type = zipfile.ZIP_STORED
            zout.writestr(info, files['mimetype'])
        for n, d in files.items():
            if n == 'mimetype':
                continue
            zout.writestr(n, d)
    buf.seek(0)
    return buf.read()

# ─── DOC (legado) ───────────────────────────────────────────────────────────

def processar_doc_libreoffice(file_bytes, filename, mes_abrev, ano, pagamentos, total_str):
    """Tenta processar .doc via LibreOffice: converte para docx, processa, converte de volta."""
    tmp_dir = tempfile.mkdtemp()
    doc_path = os.path.join(tmp_dir, filename)
    with open(doc_path, 'wb') as f:
        f.write(file_bytes)

    result = subprocess.run(
        ['libreoffice', '--headless', '--convert-to', 'docx', '--outdir', tmp_dir, doc_path],
        capture_output=True, timeout=90
    )
    docx_path = os.path.join(tmp_dir, os.path.splitext(filename)[0] + '.docx')
    if not os.path.exists(docx_path):
        raise RuntimeError("LibreOffice não converteu o arquivo")

    with open(docx_path, 'rb') as f:
        docx_bytes = f.read()

    doc = abrir_documento(docx_bytes, docx_path)
    doc = atualizar_documento(doc, mes_abrev, ano, pagamentos, total_str)

    processed_path = os.path.join(tmp_dir, 'saida.docx')
    doc.save(processed_path)

    subprocess.run(
        ['libreoffice', '--headless', '--convert-to', 'doc', '--outdir', tmp_dir, processed_path],
        capture_output=True, timeout=90
    )
    out_doc = os.path.join(tmp_dir, 'saida.doc')
    if os.path.exists(out_doc):
        with open(out_doc, 'rb') as f:
            return f.read(), '.doc'

    with open(processed_path, 'rb') as f:
        return f.read(), '.docx'

# ─── Utilitário ─────────────────────────────────────────────────────────────

def nome_saida(nome_entrada, mes_nome_arquivo, ext_saida=None):
    """Deriva nome do arquivo de saída substituindo o mês."""
    base = os.path.splitext(nome_entrada)[0]
    ext  = ext_saida if ext_saida else os.path.splitext(nome_entrada)[1]
    meses_re = '|'.join([
        'JANEIRO','FEVEREIRO','MARCO','MARÇO','ABRIL','MAIO','JUNHO',
        'JULHO','AGOSTO','SETEMBRO','OUTUBRO','NOVEMBRO','DEZEMBRO',
        'JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'
    ])
    novo = re.sub(
        rf'\b({meses_re})\b', mes_nome_arquivo,
        base.upper(), flags=re.IGNORECASE
    )
    return novo + ext

# ─── Interface ──────────────────────────────────────────────────────────────

st.info(
    "📌 **Antes de enviar o documento, verifique:**\n\n"
    "O arquivo deve conter o **CNPJ da empresa** no formato `XX.XXX.XXX/XXXX-XX` "
    "(geralmente no campo *OCS* do relatório). "
    "O programa usa o CNPJ para buscar os pagamentos automaticamente no Portal da Transparência. "
    "Se o documento tiver apenas o nome da empresa, o relatório **não poderá ser gerado**."
)

uploaded = st.file_uploader(
    "📎 Selecione o documento base (.docx, .dotx, .odt ou .doc)",
    type=["docx", "dotx", "odt", "doc"]
)

ano_atual = date.today().year
anos_disponiveis = list(range(ano_atual, ano_atual - 7, -1))

col1, col2 = st.columns(2)
with col1:
    mes_selecionado = st.selectbox("Mês de referência", MESES_LISTA)
with col2:
    ano_selecionado = st.selectbox("Ano", anos_disponiveis)

gerar = st.button(
    "📄 Gerar Relatório",
    type="primary",
    disabled=(uploaded is None)
)

# ─── Lógica principal ───────────────────────────────────────────────────────

if gerar and uploaded:
    mes_num = MESES_LISTA.index(mes_selecionado) + 1
    mes_abrev, mes_nome_arq = MESES[mes_num]
    ano = str(ano_selecionado)
    ext_entrada = os.path.splitext(uploaded.name)[1].lower()

    progress = st.progress(0)
    status   = st.empty()
    file_bytes = uploaded.read()

    # PASSO 1 — CNPJ
    status.info("🔍 Lendo documento e extraindo CNPJ...")
    cnpj = extrair_cnpj(file_bytes, ext_entrada)
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
    pagamentos, api_status, api_erro, todos_brutos = get_pagamentos(limpar_cnpj(cnpj), mes_num, ano)
    progress.progress(65)

    # Debug: mostra dados brutos da API
    with st.expander("🔍 Debug — resposta bruta da API"):
        st.write(f"**Status HTTP:** {api_status}")
        st.write(f"**Total de registros retornados pela API (todos os meses):** {len(todos_brutos)}")
        if api_erro:
            st.write(f"**Erro:** {api_erro}")
        if todos_brutos:
            st.write("**Primeiros registros (raw):**")
            st.json(todos_brutos[:5])

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
    ext_saida = ext_entrada
    mime_type = 'application/octet-stream'

    if ext_entrada == '.odt':
        output_bytes = atualizar_odt(file_bytes, mes_abrev, ano, pagamentos, total_str)
        ext_saida = '.odt'
        mime_type = 'application/vnd.oasis.opendocument.text'

    elif ext_entrada == '.doc':
        try:
            output_bytes, ext_saida = processar_doc_libreoffice(
                file_bytes, uploaded.name, mes_abrev, ano, pagamentos, total_str
            )
            if ext_saida == '.docx':
                st.warning(
                    "⚠️ Não foi possível manter o formato `.doc`. "
                    "O arquivo foi salvo como `.docx`."
                )
            mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        except Exception:
            st.error(
                "❌ **Formato `.doc` não suportado para processamento automático.**\n\n"
                "O formato `.doc` (Word 97-2003) requer conversão prévia. "
                "Por favor, abra o arquivo no Word e salve como **`.docx`**, "
                "depois envie novamente."
            )
            st.stop()

    else:
        # DOCX / DOTX
        doc = abrir_documento(file_bytes, uploaded.name)
        doc = atualizar_documento(doc, mes_abrev, ano, pagamentos, total_str)
        buf = io.BytesIO()
        doc.save(buf)
        output_bytes = buf.getvalue()
        ext_saida = '.docx'
        mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

    progress.progress(90)
    output_name = nome_saida(uploaded.name, mes_nome_arq, ext_saida)
    progress.progress(100)
    status.success("✅ Documento gerado com sucesso!")

    st.download_button(
        label=f"⬇️ Baixar {output_name}",
        data=output_bytes,
        file_name=output_name,
        mime=mime_type,
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
