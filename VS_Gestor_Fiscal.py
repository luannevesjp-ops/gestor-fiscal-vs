# ============================================================================
# GESTOR FISCAL - LUATECH
# Sistema de Gestão Fiscal com Streamlit
# ============================================================================

import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, ColumnsAutoSizeMode, DataReturnMode
from io import BytesIO
import requests
import time
import os
import base64
import json
from pathlib import Path
from datetime import date

# Tkinter só existe em ambientes com display (execução local). No Streamlit Cloud não está disponível.
try:
    import tkinter as _tk_test
    _tk_test.Tk().destroy()
    _TKINTER_OK = True
    del _tk_test
except Exception:
    _TKINTER_OK = False

# ============================================================================
# CONFIGURAÇÕES INICIAIS
# ============================================================================

st.set_page_config(page_title="LUATECH-GESTÃO-VS", layout="wide")

if 'main_container' not in st.session_state:
    st.session_state.main_container = st.empty()

SHEET_ID         = "169PDNXNSa_0ybDg2wQZbRewAd2mWKRQgfexgUXM8cZQ"
GOOGLE_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=xlsx"
SHEET_EMPRESAS   = "GERAL"
SHEET_XML_DMS    = "Leitura Xml DMS"
SHEET_XML_REST   = "Leitura Xml REST"
SHEET_SEFAZ      = "SEFAZ"
SHEET_SEFAZ_INICIAL = "SEFAZ INICIAL"
SHEET_SEFAZ_FINAL    = "SEFAZ FINAL"
SHEET_SITUACAO_FISCAL = "SITUAÇÃO FISCAL"
SHEET_CERT_ABA   = "CERTIFICADOS"
SHEET_EMAIL_ABA  = "EMAIL"
SHEET_MSG_ABA    = "MENSAGEM"

CERT_DATA_FILE = Path(os.path.abspath(__file__)).parent / "cert_data.json"

# ============================================================================
# CSS E ESTILOS
# ============================================================================

st.markdown("""
<meta name="google" content="notranslate">
<meta name="googlebot" content="notranslate">
<script>
window.addEventListener('error', function(e) {
    if (e.message && e.message.includes('removeChild')) {
        e.preventDefault();
        console.warn('Erro removeChild suprimido:', e.message);
    }
});
</script>
""", unsafe_allow_html=True)

st.markdown("""
<style>
.header-class .ag-header-cell-label {
    color: white !important;
    font-weight: bold !important;
    background-color: #1d3f77 !important;
}
.sidebar-lt {
    background-color: #1d3f77;
    padding: 0;
    margin: 0;
}
.sidebar-lt img {
    width: 100%;
    display: block;
    border-radius: 0;
}
/* Botões do menu principal */
.menu-btn button {
    width: 100%;
    height: 110px;
    font-size: 22px !important;
    font-weight: bold !important;
    border-radius: 12px !important;
    border: 2px solid #1d3f77 !important;
    background-color: #1d3f77 !important;
    color: white !important;
    cursor: pointer;
    transition: background-color 0.2s;
}
.menu-btn button:hover {
    background-color: #163066 !important;
}
.ag-body-horizontal-scroll { display: block !important; }
.ag-body-horizontal-scroll-viewport { display: block !important; }
.ag-root-wrapper { overflow: visible !important; }
.ag-body-horizontal-scroll { opacity: 1 !important; height: 16px !important; }
.ag-body-horizontal-scroll-viewport { overflow-x: scroll !important; }
</style>
""", unsafe_allow_html=True)

grid_container = st.empty()

# ============================================================================
# FUNÇÕES AUXILIARES
# ============================================================================

@st.cache_data(ttl=600)
def le_planilha_google(url: str, aba: str):
    try:
        resp = requests.get(url)
        resp.raise_for_status()
        df = pd.read_excel(BytesIO(resp.content), sheet_name=aba, engine='openpyxl')
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Erro ao ler a planilha: {e}")
        return None


def exibe_aggrid(df, height=400, grid_key="grid", selection_mode='none', retorna_filtrado=False):
    # retorna_filtrado=True: a grade avisa o Python a cada filtro/ordenação feito
    # nas colunas, e o .data do retorno passa a trazer só as linhas visíveis
    # (usado pra exportar pro Excel exatamente o que está filtrado na tela).
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filter=True, sortable=True, editable=False, resizable=True,
                                 minWidth=110, wrapHeaderText=True, autoHeaderHeight=True)

    if selection_mode != 'none':
        gb.configure_selection(selection_mode=selection_mode, use_checkbox=True)

    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            gb.configure_column(col, filter="agNumberColumnFilter")
        else:
            gb.configure_column(col, filter="agTextColumnFilter")

    gb.configure_grid_options(
        domLayout="normal", floatingFilter=True, headerHeight=40, rowHeight=30,
        enableBrowserTooltips=True, enableCellTextSelection=True, suppressMenuHide=True,
        localeText={
            'filterOoo': 'Filtrar...', 'contains': 'Contém', 'notContains': 'Não contém',
            'equals': 'Igual', 'notEqual': 'Diferente', 'blank': 'Em branco',
            'notBlank': 'Não em branco', 'noRowsToShow': 'Nenhum registro para mostrar',
        }
    )

    grid_options = gb.build()
    update_on = ['selectionChanged'] if selection_mode != 'none' else []
    if retorna_filtrado:
        update_on += ['filterChanged', 'sortChanged']

    return AgGrid(df, gridOptions=grid_options, height=height, key=grid_key,
                  columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS,
                  enable_enterprise_modules=False,
                  data_return_mode=(DataReturnMode.FILTERED_AND_SORTED if retorna_filtrado
                                    else DataReturnMode.AS_INPUT),
                  update_on=update_on, allow_unsafe_jscode=True, reload_data=False)


def exibe_aggrid_com_oculta(df, height=400, grid_key="grid", selection_mode='none', colunas_ocultas=None):
    if colunas_ocultas is None:
        colunas_ocultas = []

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filter=True, sortable=True, editable=False, resizable=True,
                                 minWidth=110, wrapHeaderText=True, autoHeaderHeight=True)

    if selection_mode != 'none':
        gb.configure_selection(selection_mode=selection_mode, use_checkbox=True)

    for col in df.columns:
        if col in colunas_ocultas:
            gb.configure_column(col, hide=True)
        elif pd.api.types.is_numeric_dtype(df[col]):
            gb.configure_column(col, filter="agNumberColumnFilter")
        else:
            gb.configure_column(col, filter="agTextColumnFilter")

    gb.configure_grid_options(
        domLayout="normal", floatingFilter=True, headerHeight=40, rowHeight=30,
        enableBrowserTooltips=True, enableCellTextSelection=True, suppressMenuHide=True,
        localeText={'filterOoo': 'Filtrar...', 'noRowsToShow': 'Nenhum registro'}
    )

    grid_options = gb.build()
    update_on = ['selectionChanged'] if selection_mode != 'none' else []

    return AgGrid(df, gridOptions=grid_options, height=height, key=grid_key,
                  columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS,
                  enable_enterprise_modules=False,
                  update_on=update_on, allow_unsafe_jscode=True, reload_data=False)

# ============================================================================
# CERTIFICADO DIGITAL — PERSISTÊNCIA, LEITURA E ENVIO DE EMAIL
# ============================================================================

_MSG_PADRAO = {
    "vencendo": (
        "Prezado(a),\n\n"
        "Informamos que o certificado digital de {razao_social} (CNPJ: {cnpj}) "
        "vence em {dias} dia(s), no dia {validade}.\n\n"
        "Por favor, providencie a renovação com urgência.\n\n"
        "Atenciosamente,\nDepartamento Fiscal"
    ),
    "vencido": (
        "Prezado(a),\n\n"
        "Informamos que o certificado digital de {razao_social} (CNPJ: {cnpj}) "
        "venceu em {validade}.\n\n"
        "Por favor, providencie a renovação imediatamente.\n\n"
        "Atenciosamente,\nDepartamento Fiscal"
    ),
}

_COLS_CERT  = ["arquivo", "nome_arquivo", "senha", "razao_social", "cnpj", "validade", "validade_iso"]
_COLS_EMAIL = ["cnpj", "emails"]
_COLS_MSG   = ["tipo", "mensagem"]

# URL do Apps Script publicado como Web App na planilha Google
# (após publicar o script, cole a URL aqui)
APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzM9WHpnqFgr_EjJU4abUNMR2ivSuMqs1olP6eOpL-z7UsiXVfqGFFaJXzOz1GIJMrHFw/exec"


def _cert_carregar_dados():
    # ── Cache de sessão (evita re-download na mesma sessão) ───────────────────
    if "cert_dados" in st.session_state:
        return st.session_state["cert_dados"]

    # ── Leitura direta da planilha Google (abas CERTIFICADOS/EMAIL/MENSAGEM) ──
    try:
        resp = requests.get(GOOGLE_SHEET_URL, timeout=20)
        resp.raise_for_status()
        xls = resp.content

        try:
            df = pd.read_excel(BytesIO(xls), sheet_name=SHEET_CERT_ABA, engine="openpyxl")
            df.columns = df.columns.str.strip()
            certificados = []
            for r in df.to_dict("records"):
                if not any(str(v).strip() for v in r.values()):
                    continue
                c = {k: str(v) if v is not None else "" for k, v in r.items()}
                # Normaliza validade_iso: 'YYYY-MM-DD HH:MM:SS' → 'YYYY-MM-DD'
                vi = c.get("validade_iso", "").strip()
                if len(vi) > 10:
                    vi = vi[:10]
                c["validade_iso"] = vi
                # Normaliza validade: se vier como 'YYYY-MM-DD...' converte para 'DD/MM/YYYY'
                v = c.get("validade", "").strip()
                if len(v) >= 10 and v[4:5] == "-":
                    try:
                        from datetime import datetime as _dt
                        v = _dt.strptime(v[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
                    except Exception:
                        pass
                c["validade"] = v
                certificados.append(c)
        except Exception:
            certificados = []

        try:
            df = pd.read_excel(BytesIO(xls), sheet_name=SHEET_EMAIL_ABA, engine="openpyxl")
            df.columns = df.columns.str.strip()
            emails = {}
            for _, row in df.iterrows():
                cnpj = str(row.get("cnpj", "")).strip()
                lista = [e.strip() for e in str(row.get("emails", "")).split(";") if e.strip()]
                if cnpj and lista:
                    emails[cnpj] = lista
        except Exception:
            emails = {}

        try:
            df = pd.read_excel(BytesIO(xls), sheet_name=SHEET_MSG_ABA, engine="openpyxl")
            df.columns = df.columns.str.strip()
            mensagens = dict(_MSG_PADRAO)
            for _, row in df.iterrows():
                tipo = str(row.get("tipo", "")).strip()
                msg  = str(row.get("mensagem", "")).strip()
                if tipo and msg:
                    mensagens[tipo] = msg
        except Exception:
            mensagens = dict(_MSG_PADRAO)

        dados = {"certificados": certificados, "emails": emails, "mensagens": mensagens}
        st.session_state["cert_dados"] = dados
        return dados

    except Exception:
        pass

    # ── Fallback: JSON local ──────────────────────────────────────────────────
    if CERT_DATA_FILE.exists():
        try:
            dados = json.loads(CERT_DATA_FILE.read_text(encoding="utf-8"))
            st.session_state["cert_dados"] = dados
            return dados
        except Exception:
            pass

    return {"certificados": [], "emails": {}, "mensagens": dict(_MSG_PADRAO)}


def _cert_salvar_dados(data):
    st.session_state["cert_dados"] = data  # atualiza cache de sessão imediatamente

    # ── JSON local (backup offline) ───────────────────────────────────────────
    try:
        CERT_DATA_FILE.write_text(
            json.dumps(data, indent=2, ensure_ascii=False, default=str),
            encoding="utf-8",
        )
    except Exception:
        pass

    # ── Apps Script → grava nas abas da planilha Google ───────────────────────
    if not APPS_SCRIPT_URL:
        return

    try:
        payload = {
            "certificados": {
                "cabecalho": _COLS_CERT,
                "linhas": [[str(c.get(col, "")) for col in _COLS_CERT]
                           for c in data.get("certificados", [])],
            },
            "emails": {
                "cabecalho": _COLS_EMAIL,
                "linhas": [[cnpj, "; ".join(lista)]
                           for cnpj, lista in data.get("emails", {}).items()],
            },
            "mensagens": {
                "cabecalho": _COLS_MSG,
                "linhas": [[tipo, msg]
                           for tipo, msg in data.get("mensagens", {}).items()],
            },
        }
        requests.post(APPS_SCRIPT_URL, json=payload, timeout=30)
    except Exception as e:
        st.warning(f"Aviso: não foi possível salvar na planilha — {e}")


# Cópia dos arquivos .pfx na pasta CERTIFICADOS do Drive — Apps Script único
# para os 6 escritórios (PROGRAMA/SCRIPTS_COMPARTILHADOS/apps_script_certificados_drive.gs),
# MESMA URL em todos;
# só muda CERT_DRIVE_ESCRITORIO. O envio não pede token: quem tiver a URL só
# consegue colocar arquivo na pasta (ver/baixar exige o TOKEN_ADMIN do script).
# URL vazia = envio desligado.
CERT_DRIVE_URL = "https://script.google.com/macros/s/AKfycbw6cARz9B3vqKRIyaY9hSYVRLMXqV3on5dQT8kH18UESR1qtdF-OvlIsosksKgyDZ71kw/exec"
CERT_DRIVE_ESCRITORIO = "VS"


def _cert_enviar_drive(nome, conteudo, senha, cnpj, razao, validade_iso):
    """Retorna (True, "") se guardou, (False, erro) se falhou,
    (None, "") se o envio está desligado (CERT_DRIVE_URL vazio)."""
    if not CERT_DRIVE_URL:
        return None, ""
    try:
        resp = requests.post(CERT_DRIVE_URL, json={
            "acao": "upload",
            "escritorio": CERT_DRIVE_ESCRITORIO,
            "nome_arquivo": nome,
            "conteudo_b64": base64.b64encode(conteudo).decode("ascii"),
            "senha": senha,
            "cnpj": cnpj,
            "razao_social": razao,
            "validade_iso": validade_iso,
        }, timeout=60)
        resp.raise_for_status()
        ret = resp.json()
        if ret.get("status") == "ok":
            return True, ""
        return False, ret.get("mensagem", "resposta inesperada do Apps Script")
    except Exception as e:
        return False, str(e)


def _cert_enviar_drive_lote(itens):
    """Envia vários .pfx numa chamada só (bem mais rápido que um por vez).
    itens = [(nome, conteudo, senha, cnpj, razao, validade_iso), ...]
    Retorna (qtd_guardados, [erros], desligado). Se o Apps Script publicado
    ainda for a versão sem "upload_lote", cai no envio um por um."""
    if not CERT_DRIVE_URL:
        return 0, [], True
    try:
        resp = requests.post(CERT_DRIVE_URL, json={
            "acao": "upload_lote",
            "escritorio": CERT_DRIVE_ESCRITORIO,
            "itens": [{
                "nome_arquivo": nome,
                "conteudo_b64": base64.b64encode(conteudo).decode("ascii"),
                "senha": senha,
                "cnpj": cnpj,
                "razao_social": razao,
                "validade_iso": validade_iso,
            } for nome, conteudo, senha, cnpj, razao, validade_iso in itens],
        }, timeout=180)
        resp.raise_for_status()
        ret = resp.json()
    except Exception as e:
        return 0, [f"{nome}: {e}" for nome, *_ in itens], False

    if ret.get("status") == "ok":
        ok, erros = 0, []
        for (nome, *_), r in zip(itens, ret.get("itens", [])):
            if r.get("status") == "ok":
                ok += 1
            else:
                erros.append(f"{nome}: {r.get('mensagem', 'erro no Apps Script')}")
        return ok, erros, False

    if str(ret.get("mensagem", "")).startswith("Ação desconhecida"):   # script antigo
        ok, erros = 0, []
        for item in itens:
            sucesso, msg = _cert_enviar_drive(*item)
            if sucesso:
                ok += 1
            else:
                erros.append(f"{item[0]}: {msg}")
        return ok, erros, False
    return 0, [f"{nome}: {ret.get('mensagem', 'resposta inesperada do Apps Script')}"
               for nome, *_ in itens], False


def _cert_situacao(validade_iso: str):
    """Retorna (situação, dias) a partir de 'YYYY-MM-DD'."""
    try:
        venc = date.fromisoformat(validade_iso)
        dias = (venc - date.today()).days
        if dias < 0:
            return "VENCIDO", dias
        if dias <= 30:
            return "VENCENDO", dias
        return "NORMAL", dias
    except Exception:
        return "DESCONHECIDO", None


def _cert_ler_pfx(fonte, senha: str):
    """Lê um PFX (caminho no disco ou bytes) e retorna (razao_social, documento, validade_str, validade_iso).
    documento pode ser CNPJ (14 dígitos) ou CPF (11 dígitos).
    """
    import re as _re
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography import x509
    from datetime import timezone

    pfx_bytes = fonte if isinstance(fonte, (bytes, bytearray)) else Path(fonte).read_bytes()
    _, cert, _ = pkcs12.load_key_and_certificates(pfx_bytes, senha.encode("utf-8"))

    OID_ECNPJ = "2.16.76.1.3.3"
    OID_ECPF  = "2.16.76.1.3.1"
    documento = ""

    # 1ª tentativa: SubjectAlternativeName
    try:
        san = cert.extensions.get_extension_for_class(x509.SubjectAlternativeName).value
        for gn in san:
            if isinstance(gn, x509.OtherName):
                oid = gn.type_id.dotted_string
                txt = "".join(chr(b) for b in gn.value if 32 <= b < 127)
                if oid == OID_ECNPJ:
                    m = _re.search(r"\d{14}", txt)
                    if m:
                        documento = m.group(0)
                        break
                elif oid == OID_ECPF:
                    m = _re.search(r"\d{11}", txt)
                    if m:
                        documento = m.group(0)
                        break
    except Exception:
        pass

    # 2ª tentativa: CN no formato "NOME:DOCUMENTO"
    if not documento:
        try:
            cn = cert.subject.get_attributes_for_oid(x509.oid.NameOID.COMMON_NAME)[0].value
            if ":" in cn:
                d = _re.sub(r"\D", "", cn.split(":")[-1])
                if len(d) >= 14:
                    documento = d[-14:]
                elif len(d) >= 11:
                    documento = d[-11:]
        except Exception:
            pass

    razao = ""
    try:
        cn = cert.subject.get_attributes_for_oid(x509.oid.NameOID.COMMON_NAME)[0].value
        razao = cn.split(":")[0].strip()
    except Exception:
        pass

    try:
        venc = cert.not_valid_after_utc
    except AttributeError:
        venc = cert.not_valid_after.replace(tzinfo=timezone.utc)

    return razao, documento, venc.strftime("%d/%m/%Y"), venc.strftime("%Y-%m-%d")


def _picker_pasta_cert():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes("-topmost", True)
    root.lift()
    pasta = filedialog.askdirectory(title="Selecione a pasta com certificados .pfx")
    root.destroy()
    return pasta or ""


def _picker_arquivo_cert():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes("-topmost", True)
    root.lift()
    arquivo = filedialog.askopenfilename(
        title="Selecione o certificado .pfx",
        filetypes=[("Certificado Digital", "*.pfx"), ("Todos os arquivos", "*.*")],
    )
    root.destroy()
    return arquivo or ""


def _extrair_senha_nome(nome_arquivo: str) -> str:
    """Extrai senha do nome: 'senha' (case-insensitive) + espaços/traços opcionais + tudo até próximo espaço."""
    import re
    m = re.search(r'(?i)senha[\s\-]*([^\s]+)', Path(nome_arquivo).stem)
    return m.group(1) if m else ""


def _enviar_outlook_cert(para: list, assunto: str, corpo: str):
    try:
        import win32com.client
        ol = win32com.client.Dispatch("Outlook.Application")
        mail = ol.CreateItem(0)
        mail.To = "; ".join(para)
        mail.Subject = assunto
        mail.Body = corpo
        mail.Send()
        return True, ""
    except Exception as e:
        return False, str(e)


def _cert_excel_tabela(df):
    """Gera o .xlsx da lista de certificados formatado como Tabela do Excel
    (estilo azul TableStyleMedium2), com Validade como data de verdade e Dias
    como número, pra ordenar/filtrar certo dentro do Excel."""
    from datetime import datetime as _dt
    from openpyxl import Workbook
    from openpyxl.styles import Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = Workbook()
    ws = wb.active
    ws.title = "CERTIFICADOS"
    colunas = list(df.columns)
    ws.append(colunas)

    for reg in df.itertuples(index=False):
        linha = []
        for col, val in zip(colunas, reg):
            if col == "Validade":
                try:
                    val = _dt.strptime(str(val).strip(), "%d/%m/%Y")
                except ValueError:
                    pass
            elif col == "Dias":
                try:
                    val = int(float(val))
                except (TypeError, ValueError):
                    pass
            linha.append(val)
        ws.append(linha)

    n_linhas = ws.max_row
    for i, col in enumerate(colunas, start=1):
        letra = get_column_letter(i)
        largura = max([len(str(col))] + [len(str(c.value or "")) for c in ws[letra][1:]])
        ws.column_dimensions[letra].width = min(largura + 4, 60)
        if col == "Validade":
            for c in ws[letra][1:]:
                c.number_format = "DD/MM/YYYY"
        if col in ("Código", "Validade", "Dias", "Situação", "CPF/CNPJ", "Certificado"):
            for c in ws[letra][1:]:
                c.alignment = Alignment(horizontal="center")

    # Tabela exige ao menos 1 linha de dados além do cabeçalho
    if n_linhas >= 2:
        tabela = Table(displayName="Certificados",
                       ref=f"A1:{get_column_letter(len(colunas))}{n_linhas}")
        tabela.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tabela)
    ws.freeze_panes = "A2"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _cert_linhas_com_empresas(certs):
    """Cruza os certificados com as empresas ATIVAS da aba GERAL (menu EMPRESAS).
    Uma linha por empresa ativa (com o certificado dela, o da matriz — mesma
    raiz de CNPJ — ou SEM CERTIFICADO) + uma linha por certificado cujo CPF/CNPJ
    não está entre as empresas ativas ("Fora da base")."""
    def _doc(v):
        # CNPJ do certificado pode vir da planilha como número (perde o zero à esquerda)
        s = str(v).strip()
        if s.endswith(".0"):
            s = s[:-2]
        d = re.sub(r"\D", "", s)
        if not d:
            return ""
        return d.zfill(11) if len(d) <= 11 else _normaliza_cnpj(d)   # CPF x CNPJ

    def _linha_cert(c):
        sit, dias = _cert_situacao(c.get("validade_iso", ""))
        return c.get("validade", ""), (dias if dias is not None else "?"), sit

    # Por documento, fica o certificado de validade mais longa (renovado > antigo)
    cert_por_doc = {}
    for c in certs:
        d = _doc(c.get("cnpj", ""))
        if not d:
            continue
        atual = cert_por_doc.get(d)
        if atual is None or str(c.get("validade_iso", "")) > str(atual.get("validade_iso", "")):
            cert_por_doc[d] = c
    cert_por_raiz = {}
    for d, c in cert_por_doc.items():
        if len(d) == 14:
            atual = cert_por_raiz.get(d[:8])
            if atual is None or str(c.get("validade_iso", "")) > str(atual.get("validade_iso", "")):
                cert_por_raiz[d[:8]] = c

    rows, docs_empresas = [], set()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is not None and "Situação" in df.columns and "CNPJ" in df.columns:
        df_at = _sanitiza_df(df[df["Situação"].astype(str).str.upper() == "ATIVA"])
        for _, emp in df_at.iterrows():
            d = _doc(emp.get("CNPJ", ""))
            if not d:
                continue
            docs_empresas.add(d)
            c, origem = cert_por_doc.get(d), "Próprio"
            if c is None and len(d) == 14:
                c, origem = cert_por_raiz.get(d[:8]), "Da matriz"
            if c is not None:
                validade, dias, sit = _linha_cert(c)
            else:
                validade, dias, sit, origem = "", "", "SEM CERTIFICADO", "—"
            rows.append({
                "Código": str(emp.get("Código", "") or ""),
                "Razão Social": str(emp.get("Razão Social", "") or ""),
                "CPF/CNPJ": _formata_cnpj_mascara(d) if len(d) == 14 else d,
                "Validade": validade,
                "Dias": dias,
                "Situação": sit,
                "Certificado": origem,
                "_sit": sit,
                "_empresa": True,
            })
    else:
        st.warning("Não foi possível ler as empresas da aba GERAL — mostrando só os certificados.")

    for c in certs:
        d = _doc(c.get("cnpj", ""))
        if d and d in docs_empresas:
            continue
        validade, dias, sit = _linha_cert(c)
        rows.append({
            "Código": "",
            "Razão Social": c.get("razao_social", ""),
            "CPF/CNPJ": _formata_cnpj_mascara(d) if len(d) == 14 else d,
            "Validade": validade,
            "Dias": dias,
            "Situação": sit,
            "Certificado": "Fora da base" if docs_empresas else "Próprio",
            "_sit": sit,
            "_empresa": False,
        })
    return rows


# ── página CERTIFICADOS ────────────────────────────────────────────────────────
def pagina_certificados():
    st.markdown("<h2>CERTIFICADOS DIGITAIS</h2>", unsafe_allow_html=True)

    dados = _cert_carregar_dados()
    # Mesmo arquivo 2x na lista (ex.: Importar clicado duas vezes) → fica 1.
    # A planilha é limpa na próxima gravação (importar/remover).
    _vistos, _unicos = set(), []
    for _c in dados["certificados"]:
        if _c.get("arquivo") not in _vistos:
            _vistos.add(_c.get("arquivo"))
            _unicos.append(_c)
    dados["certificados"] = _unicos
    certs = dados["certificados"]

    # Resultado do último Importar (guardado antes do rerun, senão a mensagem some)
    for tipo, msg in st.session_state.pop("cert_import_msgs", []):
        (st.success if tipo == "ok" else st.error)(msg)

    # ── Importar certificados (um ou vários) ─────────────────────────────────
    with st.expander("➕ Adicionar Certificados", expanded=True):

        # Campo de senha padrão — sempre visível, antes do upload
        col_sp, col_info = st.columns([2, 3])
        with col_sp:
            senha_padrao_global = st.text_input(
                "🔑 Senha padrão:",
                type="password",
                key="cert_senha_padrao",
                placeholder="Usada para quem não tem senha no nome",
                help=(
                    "Esta senha é aplicada automaticamente aos certificados cujo "
                    "nome de arquivo NÃO contém 'SENHA xxxx'. "
                    "Para qualquer arquivo você pode digitar uma senha individual abaixo."
                ),
            )
        with col_info:
            st.markdown(
                "<small style='color:#555;'>"
                "Selecione um ou mais arquivos `.pfx`. "
                "Se o nome do arquivo já contiver a senha (ex.: <code>EMPRESA SENHA 1234.pfx</code>), "
                "ela será detectada automaticamente. "
                "Caso contrário, será usada a senha padrão à esquerda "
                "— ou você pode digitar uma senha individual por arquivo."
                "</small>",
                unsafe_allow_html=True,
            )

        arquivos_up = st.file_uploader(
            "Selecione os arquivos .pfx:",
            type=["pfx"],
            accept_multiple_files=True,
            key="upload_pfx_multiplos",
        )

        if arquivos_up:
            nomes_ja = {c["nome_arquivo"] for c in certs}
            novos = [f for f in arquivos_up if f.name not in nomes_ja]

            if not novos:
                st.info(f"{len(arquivos_up)} arquivo(s) selecionado(s) — todos já estão na lista.")
            else:
                senhas_novas = {}
                st.markdown(f"**{len(novos)} novo(s) certificado(s) — confira as senhas:**")
                st.markdown("<hr style='margin:6px 0'>", unsafe_allow_html=True)

                for f in novos:
                    senha_auto = _extrair_senha_nome(f.name)
                    c1, c2, c3 = st.columns([4, 2, 2])
                    with c1:
                        st.markdown(f"`{f.name}`")
                    with c2:
                        if senha_auto:
                            st.markdown(f"🔒 detectada: `{senha_auto}`")
                        elif senha_padrao_global:
                            st.markdown(f"🔑 usará padrão: `{'*' * len(senha_padrao_global)}`")
                        else:
                            st.markdown("⚠️ sem senha")
                    with c3:
                        override = st.text_input(
                            "Senha individual", type="password",
                            key=f"sn_{abs(hash(f.name))}",
                            label_visibility="collapsed",
                            placeholder="senha individual (opcional)",
                        )
                    # Prioridade: individual > detectada no nome > padrão global
                    senha_final = override or senha_auto or senha_padrao_global
                    senhas_novas[f.name] = (f, senha_final)

                st.markdown("<hr style='margin:6px 0'>", unsafe_allow_html=True)
                if st.button("✅ Importar", key="btn_importar_pasta", type="primary"):
                    adicionados, erros = 0, []
                    drive_ok, drive_erros, drive_desligado = 0, [], False
                    para_drive = []
                    for nome, (f_obj, senha) in senhas_novas.items():
                        if not senha:
                            erros.append(f"{nome}: senha não informada.")
                            continue
                        try:
                            conteudo = f_obj.read()
                            razao, cnpj, val_str, val_iso = _cert_ler_pfx(conteudo, senha)
                            certs.append({
                                "arquivo": nome,
                                "nome_arquivo": nome,
                                "senha": senha,
                                "razao_social": razao or Path(nome).stem,
                                "cnpj": cnpj,
                                "validade": val_str,
                                "validade_iso": val_iso,
                            })
                            adicionados += 1
                        except Exception as e:
                            erros.append(f"{nome}: {e}")
                            continue
                        para_drive.append((nome, conteudo, senha, cnpj, razao, val_iso))
                    # Cópia dos .pfx na pasta CERTIFICADOS do Drive — todos numa
                    # chamada só, em paralelo com a gravação na planilha (não
                    # impede a importação se falhar)
                    from concurrent.futures import ThreadPoolExecutor
                    with ThreadPoolExecutor(max_workers=1) as _pool:
                        fut_drive = _pool.submit(_cert_enviar_drive_lote, para_drive) if para_drive else None
                        dados["certificados"] = certs
                        _cert_salvar_dados(dados)
                        if fut_drive is not None:
                            drive_ok, drive_erros, drive_desligado = fut_drive.result()
                    msgs = []
                    if adicionados:
                        msgs.append(("ok", f"✅ {adicionados} certificado(s) importado(s)!"))
                    if drive_desligado:
                        msgs.append(("erro", "⚠️ Cópia no Drive desligada (CERT_DRIVE_URL vazio)."))
                    if drive_ok:
                        msgs.append(("ok", f"☁️ {drive_ok} arquivo(s) guardado(s) no Drive."))
                    for err in drive_erros:
                        msgs.append(("erro", f"⚠️ Importado, mas não foi para o Drive — {err}"))
                    for err in erros:
                        msgs.append(("erro", f"❌ Não importado — {err} "
                                             "Confira a senha e selecione o arquivo de novo."))
                    st.session_state["cert_import_msgs"] = msgs
                    st.rerun()

    st.divider()

    rows = _cert_linhas_com_empresas(certs)

    if not rows:
        st.info("Nenhum certificado cadastrado. Use as opções acima para importar.")
        return

    # ── Contadores ────────────────────────────────────────────────────────────
    n_vencidos = sum(1 for r in rows if r["_sit"] == "VENCIDO")
    n_vencendo = sum(1 for r in rows if r["_sit"] == "VENCENDO")
    n_normais  = sum(1 for r in rows if r["_sit"] == "NORMAL")
    n_sem_cert = sum(1 for r in rows if r["_sit"] == "SEM CERTIFICADO")
    n_empresas = sum(1 for r in rows if r["_empresa"])
    n_matriz   = sum(1 for r in rows if r["Certificado"] == "Da matriz")
    n_fora     = sum(1 for r in rows if r["Certificado"] == "Fora da base")

    st.markdown(
        f"<p style='text-align:right; font-size:18px;'><b>Empresas ativas:</b> {n_empresas}"
        f" | <b>Certificados:</b> {len(certs)}</p>",
        unsafe_allow_html=True,
    )

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdedec; border-radius:8px; border-left:4px solid #c0392b;'>"
            f"<span style='font-size:22px; font-weight:700; color:#c0392b;'>{n_vencidos}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Vencidos</span></div>",
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fef9e7; border-radius:8px; border-left:4px solid #f39c12;'>"
            f"<span style='font-size:22px; font-weight:700; color:#f39c12;'>{n_vencendo}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Vencendo em 30 dias</span></div>",
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#eafaf1; border-radius:8px; border-left:4px solid #27ae60;'>"
            f"<span style='font-size:22px; font-weight:700; color:#27ae60;'>{n_normais}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Normais</span></div>",
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#eef0f3; border-radius:8px; border-left:4px solid #5d6d7e;'>"
            f"<span style='font-size:22px; font-weight:700; color:#5d6d7e;'>{n_sem_cert}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Empresas sem certificado</span></div>",
            unsafe_allow_html=True,
        )

    obs = []
    if n_matriz:
        obs.append(f"{n_matriz} filial(is) usando o certificado da matriz (mesma raiz de CNPJ)")
    if n_fora:
        obs.append(f"{n_fora} certificado(s) de CPF/CNPJ que não está entre as empresas ativas")
    if obs:
        st.caption(" · ".join(obs) + " — veja a coluna *Certificado* da lista.")

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Filtro ────────────────────────────────────────────────────────────────
    filtro_cert = st.radio(
        "Filtrar por:",
        ["Todos", "Vencidos", "Vencendo em 30 dias", "Normais", "Sem certificado"],
        horizontal=True, key="filtro_cert",
    )

    mapa_sit = {"Todos": None, "Vencidos": "VENCIDO", "Vencendo em 30 dias": "VENCENDO",
                "Normais": "NORMAL", "Sem certificado": "SEM CERTIFICADO"}
    alvo_sit = mapa_sit[filtro_cert]
    rows_filtradas = [r for r in rows if alvo_sit is None or r["_sit"] == alvo_sit]

    if not rows_filtradas:
        st.info("Nenhum registro neste filtro.")
    else:
        df_cert = pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")} for r in rows_filtradas])
        cols_cert = list(df_cert.columns)  # antes do AgGrid (versões novas injetam colunas internas "::...")
        grid_cert = exibe_aggrid(df_cert, height=350, grid_key=f"grid_certs_{filtro_cert}",
                                 retorna_filtrado=True)

        # ── Baixar Excel: respeita o filtro acima (Vencidos/Vencendo/Normais)
        # e também os filtros/ordenação digitados nas colunas da grade. Sem
        # filtro nenhum ("Todos" e colunas limpas) sai a lista inteira.
        df_export = df_cert[cols_cert]
        try:
            df_grid = grid_cert.data
            if isinstance(df_grid, pd.DataFrame):
                if df_grid.empty:  # filtro das colunas não deixou nenhuma linha
                    df_export = df_export.iloc[0:0]
                elif set(cols_cert).issubset(df_grid.columns):
                    df_export = df_grid[cols_cert]
        except Exception:
            pass

        filtrado = filtro_cert != "Todos" or len(df_export) != len(df_cert)
        sufixo = filtro_cert.lower().replace(" ", "_") if filtro_cert != "Todos" else "todos"
        if len(df_export) != len(df_cert):
            sufixo += "_filtrado"
        st.download_button(
            f"📥 Baixar Excel ({len(df_export)} linha(s){' — filtrado' if filtrado else ''})",
            data=_cert_excel_tabela(df_export),
            file_name=f"certificados_{sufixo}_{date.today():%d-%m-%Y}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_excel_certs",
            disabled=df_export.empty,
        )

    # ── Remover certificado ───────────────────────────────────────────────────
    st.divider()
    with st.expander("🗑️ Remover Certificados da Lista", expanded=False):
        if not certs:
            st.info("Nenhum certificado na lista.")
        else:
            with st.form("form_del_cert"):
                st.markdown("Marque os certificados que deseja excluir e clique em **Excluir**:")
                checks = {
                    c["arquivo"]: st.checkbox(
                        f"{c.get('razao_social', '')} — "
                        f"{_formata_cnpj_mascara(c.get('cnpj',''))}  ·  {c['nome_arquivo']}",
                        key=f"chk_{i}_{abs(hash(c['arquivo']))}",   # índice: chave nunca repete
                    )
                    for i, c in enumerate(certs)
                }
                submitted = st.form_submit_button("🗑️ Excluir Selecionados", type="primary")
                if submitted:
                    para_remover = {arq for arq, marcado in checks.items() if marcado}
                    if not para_remover:
                        st.warning("Nenhum certificado marcado.")
                    else:
                        dados["certificados"] = [c for c in certs if c["arquivo"] not in para_remover]
                        _cert_salvar_dados(dados)
                        st.success(f"{len(para_remover)} certificado(s) removido(s).")
                        st.rerun()

    # ── Envio de email ────────────────────────────────────────────────────────
    certs_vencidos = [c for c in certs if _cert_situacao(c.get("validade_iso", ""))[0] == "VENCIDO"]
    certs_vencendo = [c for c in certs if _cert_situacao(c.get("validade_iso", ""))[0] == "VENCENDO"]

    def _bloco_envio(certs_alvo, sit_alvo, label, key_sfx):
        if not certs_alvo:
            return
        st.divider()
        with st.expander(f"📧 Enviar E-mail — {label} ({len(certs_alvo)})", expanded=False):
            opcoes_email = [
                f"{c.get('razao_social', '')} — {_formata_cnpj_mascara(c.get('cnpj',''))}"
                for c in certs_alvo
            ]
            selecionados = st.multiselect(
                "Selecione os certificados (todos marcados por padrão):",
                opcoes_email, default=opcoes_email, key=f"ms_email_{key_sfx}",
            )
            if st.button(f"📧 Enviar pelo Outlook", key=f"btn_email_{key_sfx}", type="primary"):
                d2 = _cert_carregar_dados()
                template = d2.get("mensagens", {}).get(
                    "vencido" if sit_alvo == "VENCIDO" else "vencendo", ""
                )
                emails_cfg = d2.get("emails", {})
                enviados, sem_email, erros = 0, 0, []
                for cert_sel in certs_alvo:
                    rotulo = f"{cert_sel.get('razao_social', '')} — {_formata_cnpj_mascara(cert_sel.get('cnpj',''))}"
                    if rotulo not in selecionados:
                        continue
                    cnpj_raw = cert_sel.get("cnpj", "")
                    enderecos = emails_cfg.get(cnpj_raw, [])
                    if not enderecos:
                        sem_email += 1
                        continue
                    _, dias_val = _cert_situacao(cert_sel.get("validade_iso", ""))
                    try:
                        corpo = template.format(
                            razao_social=cert_sel.get("razao_social", ""),
                            cnpj=_formata_cnpj_mascara(cnpj_raw),
                            dias=abs(dias_val) if dias_val is not None else "?",
                            validade=cert_sel.get("validade", ""),
                        )
                    except Exception:
                        corpo = template
                    assunto = (
                        f"CERTIFICADO DIGITAL VENCIDO — {cert_sel.get('razao_social', '')}"
                        if sit_alvo == "VENCIDO"
                        else f"CERTIFICADO DIGITAL VENCENDO — {cert_sel.get('razao_social', '')}"
                    )
                    ok, err = _enviar_outlook_cert(enderecos, assunto, corpo)
                    if ok:
                        enviados += 1
                    else:
                        erros.append(f"{cert_sel.get('razao_social', '')}: {err}")
                if enviados:
                    st.success(f"{enviados} e-mail(s) enviado(s) com sucesso!")
                if sem_email:
                    st.warning(
                        f"{sem_email} certificado(s) sem e-mail cadastrado. "
                        "Cadastre os endereços em 'ENDEREÇO DE EMAIL'."
                    )
                for err in erros:
                    st.error(err)

    _bloco_envio(certs_vencidos, "VENCIDO",  "Vencidos",          "vencidos")
    _bloco_envio(certs_vencendo, "VENCENDO", "Vencendo em 30 dias", "vencendo")


# ── página ENDEREÇO DE EMAIL ───────────────────────────────────────────────────
def pagina_emails_cnpj():
    import re as _re_em
    st.markdown("<h2>ENDEREÇO DE EMAIL POR CNPJ</h2>", unsafe_allow_html=True)

    dados  = _cert_carregar_dados()
    certs  = dados.get("certificados", [])
    emails = dados.get("emails", {})

    if not certs:
        st.info("Nenhum certificado cadastrado. Importe certificados primeiro em 'CERTIFICADOS'.")
        return

    # Deduplica por CNPJ mantendo ordem de inserção
    unicos = list({c["cnpj"]: c for c in certs if c.get("cnpj")}.values())

    # ── Download do modelo / Upload da planilha ───────────────────────────────
    col_dl, col_up = st.columns([1, 2])

    with col_dl:
        buf = BytesIO()
        pd.DataFrame([
            {
                "Razão Social": c.get("razao_social", ""),
                "CNPJ": _formata_cnpj_mascara(c["cnpj"]),
                "E-mails": "; ".join(emails.get(c["cnpj"], [])),
            }
            for c in unicos
        ] or [{"Razão Social": "", "CNPJ": "", "E-mails": ""}]).to_excel(
            buf, index=False, engine="openpyxl"
        )
        buf.seek(0)
        st.download_button(
            "⬇️ Baixar Modelo Excel",
            data=buf,
            file_name="emails_certificados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    with col_up:
        arq = st.file_uploader(
            "📤 Importar planilha preenchida (.xlsx)",
            type=["xlsx"],
            key="upload_emails_xlsx",
        )
        if arq is not None:
            try:
                df_imp = pd.read_excel(BytesIO(arq.read()), engine="openpyxl")
                df_imp.columns = df_imp.columns.str.strip()
                importados = 0
                for _, row in df_imp.iterrows():
                    cnpj_d = _re_em.sub(r'\D', '', str(row.get("CNPJ", "")))
                    lista  = [e.strip() for e in str(row.get("E-mails", "")).split(";")
                              if e.strip() and "@" in e]
                    if cnpj_d and lista:
                        emails[cnpj_d] = lista
                        importados += 1
                if importados:
                    dados["emails"] = emails
                    _cert_salvar_dados(dados)
                    st.success(f"E-mails importados para {importados} empresa(s) e salvos!")
                    st.rerun()
                else:
                    st.warning("Nenhum e-mail válido encontrado. Verifique a coluna 'E-mails'.")
            except Exception as e:
                st.error(f"Erro ao ler a planilha: {e}")

    st.divider()
    st.markdown(
        "Preencha ou ajuste os e-mails abaixo. "
        "Separe múltiplos endereços com **ponto e vírgula** (`;`)."
    )
    st.markdown("<br>", unsafe_allow_html=True)

    # Cabeçalho
    h1, h2, h3 = st.columns([3, 2, 5])
    with h1:
        st.markdown("**Razão Social**")
    with h2:
        st.markdown("**CPF/CNPJ**")
    with h3:
        st.markdown("**E-mails (separados por `;`)**")
    st.divider()

    novos_emails = {}
    for cert in unicos:
        cnpj_raw = cert["cnpj"]
        atual = "; ".join(emails.get(cnpj_raw, []))
        col_r, col_c, col_e = st.columns([3, 2, 5])
        with col_r:
            st.markdown(cert.get("razao_social", ""))
        with col_c:
            st.markdown(_formata_cnpj_mascara(cnpj_raw))
        with col_e:
            novos_emails[cnpj_raw] = st.text_input(
                "Emails",
                value=atual,
                key=f"email_input_{cnpj_raw}",
                label_visibility="collapsed",
                placeholder="email1@emp.com; email2@emp.com",
            )

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💾 Salvar E-mails", key="btn_salvar_emails", type="primary"):
        for cnpj_raw, valor in novos_emails.items():
            lista = [e.strip() for e in valor.split(";") if e.strip()]
            if lista:
                emails[cnpj_raw] = lista
            elif cnpj_raw in emails:
                del emails[cnpj_raw]
        dados["emails"] = emails
        _cert_salvar_dados(dados)
        st.success("E-mails salvos com sucesso!")
        st.rerun()


# ── página MENSAGENS DE EMAIL ─────────────────────────────────────────────────
def pagina_mensagens_email():
    st.markdown("<h2>MENSAGENS DE E-MAIL</h2>", unsafe_allow_html=True)
    st.markdown(
        "Personalize os modelos de mensagem. Variáveis disponíveis: "
        "`{razao_social}` `{cnpj}` `{dias}` `{validade}`"
    )
    st.markdown("<br>", unsafe_allow_html=True)

    dados = _cert_carregar_dados()
    msgs  = dados.get("mensagens", {})

    col_m1, col_m2 = st.columns(2)

    with col_m1:
        st.markdown("### Certificados Vencendo (em até 30 dias)")
        msg_vencendo = st.text_area(
            "Modelo — Vencendo:", value=msgs.get("vencendo", ""),
            height=220, key="ta_msg_vencendo", label_visibility="collapsed",
        )

    with col_m2:
        st.markdown("### Certificados Vencidos")
        msg_vencido = st.text_area(
            "Modelo — Vencido:", value=msgs.get("vencido", ""),
            height=220, key="ta_msg_vencido", label_visibility="collapsed",
        )

    if st.button("💾 Salvar Mensagens", key="btn_salvar_msgs", type="primary"):
        dados["mensagens"] = {"vencendo": msg_vencendo, "vencido": msg_vencido}
        _cert_salvar_dados(dados)
        st.success("Mensagens salvas com sucesso!")

    st.divider()
    st.markdown("### Pré-visualização com dados de exemplo")

    exemplo = {
        "razao_social": "EMPRESA EXEMPLO LTDA",
        "cnpj": "12.345.678/0001-90",
        "dias": 12,
        "validade": "09/07/2026",
    }
    col_p1, col_p2 = st.columns(2)
    with col_p1:
        st.markdown("**Vencendo:**")
        try:
            st.code(msg_vencendo.format(**exemplo), language=None)
        except Exception as e:
            st.warning(f"Erro na variável: {e}")
    with col_p2:
        st.markdown("**Vencido:**")
        try:
            st.code(msg_vencido.format(**exemplo), language=None)
        except Exception as e:
            st.warning(f"Erro na variável: {e}")


# ============================================================================
# AUTENTICAÇÃO
# ============================================================================

if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False

# Controle da área principal do menu
if "menu_area" not in st.session_state:
    st.session_state["menu_area"] = None

def tela_login():
    st.markdown("<h1 style='text-align:center; color:#0f4fa3;'>Gestão Fiscal</h1>", unsafe_allow_html=True)
    senha = st.text_input("Senha", type="password", max_chars=20)
    if st.button("Entrar"):
        if senha == "VS":
            st.session_state["autenticado"] = True
        else:
            st.error("Senha incorreta.")

if not st.session_state["autenticado"]:
    tela_login()
    st.stop()

# ============================================================================
# MENU PRINCIPAL (FISCAL / PARALEGAL / CONTÁBIL)
# ============================================================================

def tela_menu_principal():
    """Tela de seleção da área após o login"""
    st.sidebar.markdown("""
    <div class="sidebar-lt" >
        <img src="data:image/png;base64,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" style="width:100%; display:block;">
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("<h1 style='text-align:center; color:#1d3f77;'>Selecione a Área</h1>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.markdown('<div class="menu-btn">', unsafe_allow_html=True)
        if st.button("FISCAL", use_container_width=True, key="btn_fiscal"):
            st.session_state["menu_area"] = "FISCAL"
            st.session_state["pagina_atual"] = "EMPRESAS"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="menu-btn">', unsafe_allow_html=True)
        if st.button("DEPARTAMENTO PARALEGAL", use_container_width=True, key="btn_paralegal"):
            st.session_state["menu_area"] = "PARALEGAL"
            st.session_state["pagina_atual"] = "EMPRESAS"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col3:
        st.markdown('<div class="menu-btn">', unsafe_allow_html=True)
        if st.button("CONTÁBIL", use_container_width=True, key="btn_contabil"):
            st.session_state["menu_area"] = "CONTÁBIL"
            st.session_state["pagina_atual"] = "EMPRESAS"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col4:
        st.markdown('<div class="menu-btn">', unsafe_allow_html=True)
        if st.button("SITUAÇÃO FISCAL", use_container_width=True, key="btn_situacao_fiscal"):
            st.session_state["menu_area"] = "SITUAÇÃO FISCAL"
            st.session_state["pagina_atual"] = "DASHBOARD"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    with col5:
        st.markdown('<div class="menu-btn">', unsafe_allow_html=True)
        if st.button("CERTIFICADO DIGITAL", use_container_width=True, key="btn_cert"):
            st.session_state["menu_area"] = "CERTIFICADO DIGITAL"
            st.session_state["pagina_atual"] = "CERTIFICADOS"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

if st.session_state["menu_area"] is None:
    tela_menu_principal()
    st.stop()

# ============================================================================
# SIDEBAR
# ============================================================================

st.sidebar.markdown("""
<div class="sidebar-lt" >
    <img src="data:image/png;base64,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" style="width:100%; display:block;">
</div>
""", unsafe_allow_html=True)

# Exibe a área atual e botão de voltar
label_area = {"FISCAL": "FISCAL", "PARALEGAL": "DEPARTAMENTO PARALEGAL", "CONTÁBIL": "CONTÁBIL", "CERTIFICADO DIGITAL": "CERTIFICADO DIGITAL", "SITUAÇÃO FISCAL": "SITUAÇÃO FISCAL"}
st.sidebar.markdown(f"<p style='text-align:center; color:#1d3f77; font-weight:bold; margin-top:10px;'>{label_area.get(st.session_state['menu_area'], st.session_state['menu_area'])}</p>", unsafe_allow_html=True)

if st.sidebar.button("← DEPARTAMENTOS", use_container_width=True):
    st.session_state["menu_area"] = None
    st.session_state["pagina_atual"] = None
    st.rerun()

st.sidebar.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

# Define as páginas disponíveis por área
if st.session_state["menu_area"] == "FISCAL":
    paginas_disponiveis = ["EMPRESAS", "SIMPLES NACIONAL", "REINF", "DCTF WEB",
                           "DMS", "SERVIÇOS TOMADOS", "SEFAZ", "LEITURA XML DMS", "LEITURA XML REST","SEFAZ ALTERAÇÃO QUANTIDADE NOTAS"]

elif st.session_state["menu_area"] == "PARALEGAL":
    paginas_disponiveis = ["DASHBOARD", "EMPRESAS", "CND MUNICIPAL", "SEM ACESSO", "ALVARÁS"]

elif st.session_state["menu_area"] == "CONTÁBIL":
    paginas_disponiveis = ["EMPRESAS"]

elif st.session_state["menu_area"] == "CERTIFICADO DIGITAL":
    paginas_disponiveis = ["CERTIFICADOS", "ENDEREÇO DE EMAIL", "MENSAGENS DE EMAIL"]

elif st.session_state["menu_area"] == "SITUAÇÃO FISCAL":
    paginas_disponiveis = ["DASHBOARD", "EMPRESAS", "CAIXA POSTAL"]

else:
    paginas_disponiveis = ["EMPRESAS"]

pagina = st.sidebar.radio("Menu", paginas_disponiveis,
                          label_visibility="collapsed")

if "pagina_atual" not in st.session_state:
    st.session_state["pagina_atual"] = pagina

if st.session_state["pagina_atual"] != pagina:
    st.session_state["pagina_atual"] = pagina
    st.rerun()

# ============================================================================
# PÁGINAS
# ============================================================================

def pagina_empresas():
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    if "Situação" in df.columns:
        df_empresas = df[df["Situação"].astype(str).str.upper() == "ATIVA"]
    else:
        st.error("Coluna 'Situação' não encontrada.")
        return
    
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado", "Matriz / Filial", "Situação"]
    df_empresas = df_empresas[[c for c in colunas if c in df_empresas.columns]]
    df_empresas = _sanitiza_df(df_empresas)
    total_empresas = df_empresas.shape[0]

    # ── paleta de cores por regime ────────────────────────────────────────────
    _CORES_REGIME = [
        "#1d3f77", "#27ae60", "#e67e22", "#8e44ad",
        "#c0392b", "#2471a3", "#148f77", "#d35400",
        "#7f8c8d", "#b7950b",
    ]

    st.subheader("Empresas - Apenas ATIVAS")
    st.markdown(f"<p style='text-align:right; font-size:20px;'><b>Total:</b> {total_empresas} | <b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)

    if "Regime" in df_empresas.columns:
        regime_serie = df_empresas["Regime"].replace({"nan": "", "None": ""}).fillna("")
        regime_serie = regime_serie.apply(lambda v: "Em Branco" if str(v).strip() == "" else str(v).strip())
        contagem_regime = regime_serie.value_counts().to_dict()

        badges = ""
        for i, (regime, qtd) in enumerate(sorted(contagem_regime.items())):
            cor = _CORES_REGIME[i % len(_CORES_REGIME)]
            badges += (
                f"<span style='display:inline-block; margin:3px 6px 3px 0; padding:5px 14px; "
                f"background:{cor}; color:#fff; border-radius:20px; font-size:13px; font-weight:600;'>"
                f"{regime}: {qtd}</span>"
            )
        st.markdown(f"<div style='margin-bottom:10px;'>{badges}</div>", unsafe_allow_html=True)

    with st.container():
        df_empresas = _sanitiza_df(df_empresas)
        exibe_aggrid(df_empresas, height=400, grid_key="grid_empresas")
    
    output = BytesIO()
    df_empresas.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(), file_name="empresas.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


import re  # adicione no topo do arquivo se ainda não tiver

# ── helper CNPJ ─────────────────────────────────────────────────────────────
def _normaliza_cnpj(val):
    """Remove formatação e garante 14 dígitos — remove dígito extra à direita se vier com 15."""
    digits = re.sub(r'\D', '', str(val))
    if len(digits) == 15:
        digits = digits[:14]   # remove o dígito extra da direita
    return digits.zfill(14)


def _formata_cnpj_mascara(val):
    """Formata CPF (000.000.000-00) ou CNPJ (00.000.000/0000-00)."""
    digits = re.sub(r'\D', '', str(val))
    if len(digits) == 15:
        digits = digits[:14]
    if len(digits) == 11:
        return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:11]}"
    digits = digits.zfill(14)
    return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:14]}"

@st.dialog("Simples Nacional — Não Concluídas")
def _modal_simples_nao_concluidas(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) não concluída(s)**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Município",
                        "SIMPLES GERADO", "MOTIVO SITUAÇÃO DO DAS"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)



def pagina_simples():
    import plotly.graph_objects as go
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] \
        if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors="coerce").strftime("%m/%Y") \
        if competencia_raw else ""

    if "Situação" not in df.columns:
        st.error("Coluna 'Situação' não encontrada.")
        return

    df_ativas = df[
        (df["Situação"].astype(str).str.upper() == "ATIVA") &
        (df["Regime"].astype(str).str.upper() == "SIMPLES NACIONAL")
    ].copy()

    if df_ativas.empty:
        st.warning("Nenhuma empresa SIMPLES NACIONAL ATIVA encontrada.")
        return

    # ── detecta filiais pela coluna MATRIZ / FILIAL ───────────────────────────
    if "MATRIZ / FILIAL" in df_ativas.columns:
        mask_filial = df_ativas["MATRIZ / FILIAL"].astype(str).str.strip().str.upper() == "FILIAL"
    else:
        mask_filial = pd.Series([False] * len(df_ativas), index=df_ativas.index)

    df_filiais    = df_ativas[mask_filial].copy()
    df_nao_filial = df_ativas[~mask_filial].copy()

    # ── colunas para exibição ─────────────────────────────────────────────────
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado",
               "SIMPLES GERADO", "MOTIVO SITUAÇÃO DO DAS", "Situação"]

    df_nao_filial = df_nao_filial[[c for c in colunas if c in df_nao_filial.columns]].copy()
    df_filiais    = df_filiais[[c for c in colunas if c in df_filiais.columns]].copy()

    # ── CNPJ: 14 dígitos ─────────────────────────────────────────────────────
    for _df in [df_nao_filial, df_filiais]:
        if "CNPJ" in _df.columns:
            _df["CNPJ"] = _df["CNPJ"].apply(_normaliza_cnpj)

    # ── classificação das não-filiais ─────────────────────────────────────────
    def _classifica(val):
        v = str(val).strip().upper() \
            if pd.notna(val) and str(val).strip() not in ("", "NAN") else ""
        if "CONCLUÍDA" in v or "CONCLUIDA" in v:
            return "Concluída"
        return "Não Concluída"

    if "SIMPLES GERADO" in df_nao_filial.columns:
        df_nao_filial["SIMPLES GERADO"] = df_nao_filial["SIMPLES GERADO"].apply(_classifica)

    if "SIMPLES GERADO" in df_filiais.columns:
        df_filiais["SIMPLES GERADO"] = "Filial"

    # ── df final para a tabela ────────────────────────────────────────────────
    df_simples = pd.concat([df_nao_filial, df_filiais], ignore_index=True)

    # ── contagens ─────────────────────────────────────────────────────────────
    concluidas     = (df_nao_filial["SIMPLES GERADO"] == "Concluída").sum()
    nao_concluidas = (df_nao_filial["SIMPLES GERADO"] == "Não Concluída").sum()
    filiais        = len(df_filiais)
    total          = concluidas + nao_concluidas

    st.markdown("<h2>SIMPLES NACIONAL</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Concluídas:</b> {concluidas} &nbsp;|&nbsp; "
        f"<b>Não concluídas:</b> {nao_concluidas} &nbsp;|&nbsp; "
        f"<b>Filiais:</b> {filiais} &nbsp;|&nbsp; "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True,
    )

    # ── donut ─────────────────────────────────────────────────────────────────
    if "simples_chart_key" not in st.session_state:
        st.session_state["simples_chart_key"] = 0

    pct_c  = round(concluidas     / total * 100) if total else 0
    pct_nc = round(nao_concluidas / total * 100) if total else 0

    fig = go.Figure(data=[go.Pie(
        labels=["Concluídas", "Não Concluídas"],
        values=[int(concluidas), int(nao_concluidas)],
        hole=0.68,
        marker=dict(
            colors=["#27ae60", "#e74c3c"],
            line=dict(color="#ffffff", width=3),
        ),
        textinfo="none",
        hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
        direction="clockwise",
        sort=False,
    )])

    fig.update_layout(
        paper_bgcolor="white", plot_bgcolor="white",
        showlegend=False,
        margin=dict(t=20, b=20, l=20, r=20),
        height=300,
        annotations=[dict(
            text=f"<b>{total}</b><br><span style='font-size:11px'>empresas</span>",
            x=0.5, y=0.5,
            xanchor="center", yanchor="middle",
            showarrow=False,
            font=dict(size=22, color="#1d3f77"),
        )],
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"chart_simples_{st.session_state['simples_chart_key']}",
    )

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#f0faf4; "
            f"border-radius:8px; border-left:4px solid #27ae60;'>"
            f"<span style='font-size:22px; font-weight:700; color:#27ae60;'>{concluidas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Concluídas ({pct_c}%)</span></div>",
            unsafe_allow_html=True,
        )
    with col_r:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdf2f2; "
            f"border-radius:8px; border-left:4px solid #e74c3c;'>"
            f"<span style='font-size:22px; font-weight:700; color:#e74c3c;'>{nao_concluidas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Não Concluídas ({pct_nc}%)</span></div>",
            unsafe_allow_html=True,
        )
        if st.button("Ver empresas não concluídas", use_container_width=True,
                     key="btn_nao_concluidas"):
            df_nc = df_nao_filial[df_nao_filial["SIMPLES GERADO"] == "Não Concluída"]
            _modal_simples_nao_concluidas(df_nc)

    st.divider()

    # ── tabela principal ──────────────────────────────────────────────────────
    df_simples = _sanitiza_df(df_simples)
    exibe_aggrid(df_simples, height=400, grid_key="grid_simples")

    output = BytesIO()
    df_simples.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel", data=output.getvalue(),
        file_name="simples_nacional.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _limpa_codigo_sf(val):
    """Normaliza a coluna Código para casar linhas entre os blocos A-D e F+ da
    aba SITUAÇÃO FISCAL (mesmo padrão de _limpa_codigo usado em pagina_sefaz_comparacao)."""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s.upper().replace("NAN", "").strip()


def _extrai_mensagens_caixa_postal(valor):
    """Uma célula de CAIXA POSTAL pode trazer mais de um arquivo, separados por
    ' | ' (ex: '_461_Painel de Conformidade...png | _461_Opção pelo Simples...').
    Retorna a lista de assuntos já limpos, sem o prefixo _<código>_ nem a extensão."""
    import re
    if pd.isna(valor) or not str(valor).strip():
        return []
    assuntos = []
    for parte in str(valor).split("|"):
        assunto = re.sub(r"^_\d+_", "", parte.strip())
        assunto = re.sub(r"\.(png|pdf|jpe?g)$", "", assunto, flags=re.IGNORECASE).strip()
        # agrupa variações numeradas (ex: "Intimação nº 19986101") num único assunto,
        # senão cada número vira um botão de filtro diferente
        assunto = re.sub(r"\s*n[ºo°]\s*\d+\s*$", "", assunto, flags=re.IGNORECASE).strip()
        if assunto:
            assuntos.append(assunto)
    return assuntos


@st.dialog("SITUAÇÃO FISCAL — Não Concluídas")
def _modal_sf_nao_concluidas(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) não concluída(s)**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Status", "Detalhe"] if c in df_show.columns]
    df_exib = df_show[cols].copy()
    st.dataframe(df_exib.reset_index(drop=True), use_container_width=True, hide_index=True)


def _carrega_situacao_fiscal():
    """Lê a aba SITUAÇÃO FISCAL e faz o PROCV entre o bloco de empresas (A-D) e
    o bloco de leitura (F em diante, K/ARQUIVO fora) pelo Código — as duas partes
    não pertencem às mesmas linhas na planilha. Compartilhado pelas páginas
    DASHBOARD e EMPRESAS do departamento SITUAÇÃO FISCAL para não duplicar a
    lógica de junção. Retorna (df_merge, mes_leitura, cols_situacoes); em caso de
    erro retorna (None, "", [])."""
    df_sf = le_planilha_google(GOOGLE_SHEET_URL, SHEET_SITUACAO_FISCAL)
    if df_sf is None:
        return None, "", []

    cols_sf = df_sf.columns.tolist()
    if len(cols_sf) < 20:
        st.error("Estrutura da aba SITUAÇÃO FISCAL inesperada (menos de 20 colunas).")
        return None, "", []

    # ── posições fixas na aba: A-D = empresas, F-L = leitura (K oculta), ───────
    # ── M-T = Caixa Postal (menu separado), U em diante = situações ───────────
    col_cod_base, col_razao, col_regime, col_rodou = cols_sf[0], cols_sf[1], cols_sf[2], cols_sf[3]
    col_codigo, col_cnpj, col_mes, col_status, col_detalhe = (
        cols_sf[5], cols_sf[6], cols_sf[7], cols_sf[8], cols_sf[9]
    )
    # cols_sf[10] = ARQUIVO (K) — não exibir
    col_data_hora = cols_sf[11]
    cols_situacoes = cols_sf[20:]  # U em diante

    df_base = df_sf[[col_cod_base, col_razao, col_regime, col_rodou]].copy()
    df_base.columns = ["Código", "Razão Social", "Regime", "Rodou"]
    df_base = df_base[df_base["Código"].notna()].copy()
    df_base["Código"] = df_base["Código"].apply(_limpa_codigo_sf)

    if df_base.empty:
        st.warning("Nenhuma empresa encontrada na aba SITUAÇÃO FISCAL.")
        return None, "", []

    df_leitura = df_sf[[col_codigo, col_cnpj, col_mes, col_status, col_detalhe, col_data_hora]
                        + cols_situacoes].copy()
    df_leitura = df_leitura[df_leitura[col_codigo].notna()].copy()
    df_leitura[col_codigo] = df_leitura[col_codigo].apply(_limpa_codigo_sf)
    df_leitura.rename(columns={
        col_codigo: "Código", col_cnpj: "CNPJ", col_mes: "Mês Leitura",
        col_status: "Status", col_detalhe: "Detalhe", col_data_hora: "Data/Hora",
    }, inplace=True)

    mes_validos = df_leitura["Mês Leitura"].dropna()
    mes_leitura = pd.to_datetime(mes_validos.iloc[0], errors="coerce").strftime("%m/%Y") \
        if not mes_validos.empty else ""

    # ── PROCV: junta a leitura (colunas F em diante) às empresas (A-D) pelo Código ──
    df_merge = pd.merge(df_base, df_leitura, on="Código", how="left")

    if "CNPJ" in df_merge.columns:
        df_merge["CNPJ"] = df_merge["CNPJ"].apply(_formata_cnpj_mascara)

    def _classifica(val):
        v = str(val).strip().upper() if pd.notna(val) and str(val).strip() not in ("", "NAN") else ""
        return "Concluída" if "BAIXADO" in v else "Não Concluída"

    df_merge["Situação Leitura"] = df_merge["Rodou"].apply(_classifica)

    return df_merge, mes_leitura, cols_situacoes


def pagina_situacao_fiscal_dashboard():
    import plotly.graph_objects as go
    st.empty()

    df_merge, mes_leitura, _ = _carrega_situacao_fiscal()
    if df_merge is None:
        return

    concluidas     = (df_merge["Situação Leitura"] == "Concluída").sum()
    nao_concluidas = (df_merge["Situação Leitura"] == "Não Concluída").sum()
    total          = concluidas + nao_concluidas

    st.markdown("<h2>SITUAÇÃO FISCAL — DASHBOARD</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Concluídas:</b> {concluidas} &nbsp;|&nbsp; "
        f"<b>Não concluídas:</b> {nao_concluidas} &nbsp;|&nbsp; "
        f"<b>Mês de Leitura:</b> {mes_leitura}</p>",
        unsafe_allow_html=True,
    )

    # ── donut (mesmo padrão do menu SIMPLES NACIONAL) ──────────────────────────
    if "sf_chart_key" not in st.session_state:
        st.session_state["sf_chart_key"] = 0

    pct_c  = round(concluidas     / total * 100) if total else 0
    pct_nc = round(nao_concluidas / total * 100) if total else 0

    fig = go.Figure(data=[go.Pie(
        labels=["Concluídas", "Não Concluídas"],
        values=[int(concluidas), int(nao_concluidas)],
        hole=0.68,
        marker=dict(
            colors=["#27ae60", "#e74c3c"],
            line=dict(color="#ffffff", width=3),
        ),
        textinfo="none",
        hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
        direction="clockwise",
        sort=False,
    )])

    fig.update_layout(
        paper_bgcolor="white", plot_bgcolor="white",
        showlegend=False,
        margin=dict(t=20, b=20, l=20, r=20),
        height=300,
        annotations=[dict(
            text=f"<b>{total}</b><br><span style='font-size:11px'>empresas</span>",
            x=0.5, y=0.5,
            xanchor="center", yanchor="middle",
            showarrow=False,
            font=dict(size=22, color="#1d3f77"),
        )],
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"chart_sf_{st.session_state['sf_chart_key']}",
    )

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#f0faf4; "
            f"border-radius:8px; border-left:4px solid #27ae60;'>"
            f"<span style='font-size:22px; font-weight:700; color:#27ae60;'>{concluidas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Concluídas ({pct_c}%)</span></div>",
            unsafe_allow_html=True,
        )
    with col_r:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdf2f2; "
            f"border-radius:8px; border-left:4px solid #e74c3c;'>"
            f"<span style='font-size:22px; font-weight:700; color:#e74c3c;'>{nao_concluidas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Não Concluídas ({pct_nc}%)</span></div>",
            unsafe_allow_html=True,
        )
        if st.button("Ver empresas não concluídas", use_container_width=True,
                     key="btn_sf_nao_concluidas"):
            df_nc = df_merge[df_merge["Situação Leitura"] == "Não Concluída"]
            _modal_sf_nao_concluidas(df_nc)


# Apps Script (só leitura no Drive, não grava em nada) que devolve o link do
# PDF mais recente de SITUAÇÃO FISCAL de cada empresa, casando pelo Código
# embutido no nome do arquivo ("_<código>_SITUAÇÃO FISCAL MM-AAAA.pdf").
# COMPARTILHADO entre todos os escritórios (ver apps_script_pdf_situacao_fiscal.gs
# guardado em PROGRAMA/SCRIPTS_COMPARTILHADOS) - a URL só funciona de fato depois que o script
# compartilhado for republicado com a pasta do Drive da VS cadastrada em
# PASTAS_POR_ESCRITORIO; até lá a lupa simplesmente não aparece na tabela, a
# página não quebra.
APPS_SCRIPT_PDF_SITUACAO_FISCAL_URL = "https://script.google.com/macros/s/AKfycbz7FnVmU0-39_HszitoisrnNZ60fNpVVSXH55m3ufxoPmWEzj5uEy6Gsx9G87WiesXi/exec?escritorio=vs"


@st.cache_data(ttl=600)
def _busca_links_pdf_situacao_fiscal():
    """Busca no Apps Script {código: link do PDF} de SITUAÇÃO FISCAL. Retorna {}
    (silenciosamente) se a URL não estiver configurada ou a chamada falhar."""
    if not APPS_SCRIPT_PDF_SITUACAO_FISCAL_URL:
        return {}
    try:
        resp = requests.get(APPS_SCRIPT_PDF_SITUACAO_FISCAL_URL, timeout=20)
        resp.raise_for_status()
        dados = resp.json()
        return {str(k).strip(): str(v).strip() for k, v in dados.items()}
    except Exception:
        return {}


def pagina_situacao_fiscal_empresas():
    st.empty()

    df_merge, mes_leitura, cols_situacoes = _carrega_situacao_fiscal()
    if df_merge is None:
        return

    st.markdown("<h2>SITUAÇÃO FISCAL — EMPRESAS</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:16px;'>"
        f"<b>Empresas:</b> {len(df_merge)} &nbsp;|&nbsp; "
        f"<b>Mês de Leitura:</b> {mes_leitura}</p>",
        unsafe_allow_html=True,
    )

    # ── cores dos filtros (botões e listas) ─────────────────────────────────────
    CATEGORIAS_COR = {
        "OMISSÕES":      ("sf_cat_omissoes",      "#e67e22"),
        "PARCELAMENTOS": ("sf_cat_parcelamentos", "#2e86de"),
        "DÉBITOS":       ("sf_cat_debitos",        "#e74c3c"),
        "DEMAIS":        ("sf_cat_demais",         "#7f8c8d"),
    }
    regras_css = [
        # botões "Concluída"/"Não Concluída" — 1º verde, 2º vermelho
        ".st-key-sf_box_status div[data-testid='stButtonGroup'] [role='radio']:nth-of-type(1),"
        ".st-key-sf_box_status div[data-testid='stButtonGroup'] label:nth-of-type(1) {"
        " border-color:#27ae60 !important; color:#27ae60 !important; }",
        ".st-key-sf_box_status div[data-testid='stButtonGroup'] [role='radio']:nth-of-type(1)[aria-checked='true'],"
        ".st-key-sf_box_status div[data-testid='stButtonGroup'] label:nth-of-type(1)[aria-checked='true'] {"
        " background-color:#27ae60 !important; color:#fff !important; }",
        ".st-key-sf_box_status div[data-testid='stButtonGroup'] [role='radio']:nth-of-type(2),"
        ".st-key-sf_box_status div[data-testid='stButtonGroup'] label:nth-of-type(2) {"
        " border-color:#e74c3c !important; color:#e74c3c !important; }",
        ".st-key-sf_box_status div[data-testid='stButtonGroup'] [role='radio']:nth-of-type(2)[aria-checked='true'],"
        ".st-key-sf_box_status div[data-testid='stButtonGroup'] label:nth-of-type(2)[aria-checked='true'] {"
        " background-color:#e74c3c !important; color:#fff !important; }",
    ]
    for nome_cat, (key_cat, cor) in CATEGORIAS_COR.items():
        regras_css.append(
            f".st-key-{key_cat} {{ border-top:3px solid {cor} !important; border-radius:8px; padding:6px 8px 2px 8px; }}"
        )
        regras_css.append(
            f".st-key-{key_cat} span[data-baseweb='tag'] {{ background-color:{cor} !important; }}"
        )
    st.markdown(f"<style>{''.join(regras_css)}</style>", unsafe_allow_html=True)

    # ── filtros em botões (pills) — em vez de uma coluna por situação ──────────
    st.markdown("<p style='margin-bottom:2px;'><b>Status da leitura</b></p>", unsafe_allow_html=True)
    with st.container(key="sf_box_status"):
        status_sel = st.pills(
            "Status da leitura", ["Concluída", "Não Concluída"],
            selection_mode="single", label_visibility="collapsed", key="sf_pill_status",
        )

    contagens = {c: int((df_merge[c].astype(str).str.strip().str.upper() == "X").sum())
                 for c in cols_situacoes}

    # ── situações fiscais em 4 listas (tipo validação de dados), em vez de ────
    # ── um botão para cada uma das 31 colunas ──────────────────────────────────
    categorias = {"OMISSÕES": [], "PARCELAMENTOS": [], "DÉBITOS": [], "DEMAIS": []}
    for c in cols_situacoes:
        up = c.upper()
        if up.startswith("OMISSÃO"):
            categorias["OMISSÕES"].append(c)
        elif up.startswith("PARCELAMENTO"):
            categorias["PARCELAMENTOS"].append(c)
        elif up.startswith("DÉBITO"):
            categorias["DÉBITOS"].append(c)
        else:
            categorias["DEMAIS"].append(c)

    situacoes_sel = []
    cols_categorias = st.columns(4)
    for col_widget, (nome_cat, itens) in zip(cols_categorias, categorias.items()):
        with col_widget:
            key_cat, _ = CATEGORIAS_COR[nome_cat]
            with st.container(key=key_cat):
                opcoes = sorted((c for c in itens if contagens[c] > 0),
                                 key=lambda c: contagens[c], reverse=True)
                sel = st.multiselect(
                    nome_cat, opcoes,
                    format_func=lambda c: f"{c} ({contagens[c]})",
                    key=f"sf_ms_{nome_cat}",
                )
                situacoes_sel.extend(sel)

    df_filtrado = df_merge
    if status_sel:
        df_filtrado = df_filtrado[df_filtrado["Situação Leitura"] == status_sel]
    if situacoes_sel:
        mask = pd.Series(False, index=df_filtrado.index)
        for c in situacoes_sel:
            mask |= df_filtrado[c].astype(str).str.strip().str.upper() == "X"
        df_filtrado = df_filtrado[mask]

    # ── resume as situações marcadas de cada empresa numa única coluna ────────
    def _resume_situacoes(row):
        marcadas = [c for c in cols_situacoes if str(row.get(c, "")).strip().upper() == "X"]
        return " · ".join(marcadas)

    df_filtrado = df_filtrado.copy()
    df_filtrado["Situações"] = df_filtrado.apply(_resume_situacoes, axis=1)

    st.divider()
    st.caption(f"{len(df_filtrado)} empresa(s) exibida(s)")

    colunas_exibir = ["Código", "Razão Social", "Regime", "CNPJ", "Mês Leitura",
                       "Status", "Data/Hora", "Situação Leitura", "Situações"]
    colunas_exibir = [c for c in colunas_exibir if c in df_filtrado.columns]
    df_tabela = _sanitiza_df(df_filtrado[colunas_exibir])

    # ── lupa por empresa: PDF vem de uma pasta do Drive (sistema roda na nuvem do
    # ── Streamlit, sem acesso a disco/rede local — só dá pra abrir arquivo por URL) ──
    links_pdf = _busca_links_pdf_situacao_fiscal()
    df_tabela["PDF"] = df_tabela["Código"].map(links_pdf).fillna("") if links_pdf else ""

    # ── a chave do grid muda conforme o filtro — o AgGrid usa reload_data=False,
    # ── então com uma chave fixa ele ignora dado novo e mantém a lista antiga ──
    filtro_estado = "|".join([status_sel or "todos"] + sorted(situacoes_sel))
    grid_key = f"grid_situacao_fiscal_{abs(hash(filtro_estado))}"

    gb = GridOptionsBuilder.from_dataframe(df_tabela)
    gb.configure_default_column(filter=True, sortable=True, editable=False, resizable=True)
    for col in df_tabela.columns:
        if col == "PDF":
            continue
        gb.configure_column(col, filter="agTextColumnFilter")

    # cellRenderer baseado em CLASSE (init/getGui), não em função que retorna string:
    # o streamlit-aggrid usa ag-grid-react por baixo, e uma função que só devolve uma
    # string HTML é escapada como texto puro pelo React (apareceu literalmente "<a
    # href=..." cortado pela coluna estreita). Setando innerHTML manualmente dentro de
    # init() contorna esse escape.
    from st_aggrid import JsCode
    lupa_renderer = JsCode("""
        class LupaPdfRenderer {
            init(params) {
                this.eGui = document.createElement('span');
                if (params.value) {
                    this.eGui.innerHTML =
                        '<a href="' + params.value + '" target="_blank" rel="noopener" ' +
                        'style="font-size:18px; text-decoration:none;" title="Abrir PDF">🔎</a>';
                }
            }
            getGui() { return this.eGui; }
            refresh(params) { return false; }
        }
    """)
    gb.configure_column("PDF", header_name="", cellRenderer=lupa_renderer,
                         filter=False, sortable=False, resizable=False,
                         suppressSizeToFit=True, width=56, pinned="left",
                         cellStyle={"textAlign": "center"})

    gb.configure_grid_options(
        domLayout="normal", floatingFilter=True, headerHeight=40, rowHeight=30,
        enableBrowserTooltips=True, enableCellTextSelection=True, suppressMenuHide=True,
        localeText={
            'filterOoo': 'Filtrar...', 'contains': 'Contém', 'notContains': 'Não contém',
            'equals': 'Igual', 'notEqual': 'Diferente', 'blank': 'Em branco',
            'notBlank': 'Não em branco', 'noRowsToShow': 'Nenhum registro para mostrar',
        }
    )
    AgGrid(df_tabela, gridOptions=gb.build(), height=450, key=grid_key,
           fit_columns_on_grid_load=True, enable_enterprise_modules=False,
           allow_unsafe_jscode=True, reload_data=False)

    if not links_pdf:
        st.caption("🔎 Lupa de PDF ainda não configurada — falta publicar "
                   "apps_script_pdf_situacao_fiscal.gs e preencher "
                   "APPS_SCRIPT_PDF_SITUACAO_FISCAL_URL.")

    output = BytesIO()
    df_tabela.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel", data=output.getvalue(),
        file_name="situacao_fiscal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_situacao_fiscal",
    )


def pagina_caixa_postal():
    from collections import Counter
    st.empty()

    df_sf = le_planilha_google(GOOGLE_SHEET_URL, SHEET_SITUACAO_FISCAL)
    if df_sf is None:
        return

    cols_sf = df_sf.columns.tolist()
    if len(cols_sf) < 20:
        st.error("Estrutura da aba SITUAÇÃO FISCAL inesperada (menos de 20 colunas).")
        return

    col_cod_base, col_razao = cols_sf[0], cols_sf[1]
    col_codigo, col_cnpj = cols_sf[5], cols_sf[6]
    cols_caixa_postal = cols_sf[12:20]  # M a T

    df_base = df_sf[[col_cod_base, col_razao]].copy()
    df_base.columns = ["Código", "Razão Social"]
    df_base = df_base[df_base["Código"].notna()].copy()
    df_base["Código"] = df_base["Código"].apply(_limpa_codigo_sf)

    if df_base.empty:
        st.warning("Nenhuma empresa encontrada na aba SITUAÇÃO FISCAL.")
        return

    df_cp = df_sf[[col_codigo, col_cnpj] + cols_caixa_postal].copy()
    df_cp = df_cp[df_cp[col_codigo].notna()].copy()
    df_cp[col_codigo] = df_cp[col_codigo].apply(_limpa_codigo_sf)
    df_cp.rename(columns={col_codigo: "Código", col_cnpj: "CNPJ"}, inplace=True)

    df_merge = pd.merge(df_base, df_cp, on="Código", how="left")
    if "CNPJ" in df_merge.columns:
        df_merge["CNPJ"] = df_merge["CNPJ"].apply(_formata_cnpj_mascara)

    # ── extrai as mensagens de cada empresa a partir das 8 colunas CAIXA POSTAL ──
    def _mensagens_da_linha(row):
        assuntos = []
        for c in cols_caixa_postal:
            assuntos.extend(_extrai_mensagens_caixa_postal(row.get(c)))
        return assuntos

    df_merge["_mensagens"] = df_merge.apply(_mensagens_da_linha, axis=1)
    df_merge["Qtd. Mensagens"] = df_merge["_mensagens"].apply(len)
    df_merge["Mensagens"] = df_merge["_mensagens"].apply(lambda lst: " · ".join(lst))

    st.markdown("<h2>CAIXA POSTAL</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:16px;'><b>Empresas:</b> {len(df_merge)}</p>",
        unsafe_allow_html=True,
    )

    # ── filtro em botão — assunto da mensagem, um botão por assunto (seleção ──
    # ── única: marcar um desmarca automaticamente o anterior) ──────────────────
    st.markdown(
        "<style>"
        ".st-key-cp_box_assunto div[data-testid='stButtonGroup'] [role='radio'],"
        ".st-key-cp_box_assunto div[data-testid='stButtonGroup'] label {"
        " border-color:#16a085 !important; color:#16a085 !important; }"
        ".st-key-cp_box_assunto div[data-testid='stButtonGroup'] [role='radio'][aria-checked='true'],"
        ".st-key-cp_box_assunto div[data-testid='stButtonGroup'] label[aria-checked='true'] {"
        " background-color:#16a085 !important; color:#fff !important; }"
        "</style>",
        unsafe_allow_html=True,
    )

    contagem_assuntos = Counter(a for lst in df_merge["_mensagens"] for a in lst)
    opcoes_assunto = sorted(contagem_assuntos, key=contagem_assuntos.get, reverse=True)

    if opcoes_assunto:
        st.markdown("<p style='margin-bottom:2px;'><b>Assunto da mensagem</b></p>", unsafe_allow_html=True)
        with st.container(key="cp_box_assunto"):
            assunto_sel = st.pills(
                "Assunto da mensagem", opcoes_assunto,
                selection_mode="single", label_visibility="collapsed",
                format_func=lambda a: f"{a} ({contagem_assuntos[a]})", key="cp_pill_assunto",
            )
    else:
        assunto_sel = None
        st.caption("Nenhuma mensagem encontrada na Caixa Postal ainda.")

    df_filtrado = df_merge
    if assunto_sel:
        df_filtrado = df_filtrado[df_filtrado["_mensagens"].apply(lambda lst: assunto_sel in lst)]

    st.divider()
    st.caption(f"{len(df_filtrado)} empresa(s) exibida(s)")

    colunas_exibir = ["Código", "Razão Social", "CNPJ", "Qtd. Mensagens", "Mensagens"]
    df_tabela = _sanitiza_df(df_filtrado[colunas_exibir])

    # ── a chave do grid muda conforme o filtro — o AgGrid usa reload_data=False,
    # ── então com uma chave fixa ele ignora dado novo e mantém a lista antiga ──
    grid_key = f"grid_caixa_postal_{abs(hash(assunto_sel or 'todos'))}"
    exibe_aggrid(df_tabela, height=450, grid_key=grid_key)

    output = BytesIO()
    df_tabela.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel", data=output.getvalue(),
        file_name="caixa_postal.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="btn_download_caixa_postal",
    )


@st.dialog("REINF — Não Transmitidas")
def _modal_reinf_nao_transmitidas(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) não transmitida(s)**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Município",
                        "TRANSMISSÃO REINF", "MOTIVO SITUAÇÃO REINF"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.fragment
def pagina_reinf():
    import plotly.graph_objects as go
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] \
        if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors="coerce").strftime("%m/%Y") \
        if competencia_raw else ""

    if "Situação" not in df.columns:
        st.error("Coluna 'Situação' não encontrada.")
        return

    df_ativas = df[df["Situação"].astype(str).str.upper() == "ATIVA"].copy()

    if df_ativas.empty:
        st.warning("Nenhuma empresa ATIVA encontrada para REINF.")
        return

    # ── detecta filiais pela coluna MATRIZ / FILIAL ───────────────────────────
    if "MATRIZ / FILIAL" in df_ativas.columns:
        mask_filial = df_ativas["MATRIZ / FILIAL"].astype(str).str.strip().str.upper() == "FILIAL"
    else:
        mask_filial = pd.Series([False] * len(df_ativas), index=df_ativas.index)

    df_filiais    = df_ativas[mask_filial].copy()
    df_nao_filial = df_ativas[~mask_filial].copy()

    # ── colunas para exibição ─────────────────────────────────────────────────
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado",
               "TRANSMISSÃO REINF", "MOTIVO SITUAÇÃO REINF", "Situação"]

    df_nao_filial = df_nao_filial[[c for c in colunas if c in df_nao_filial.columns]].copy()
    df_filiais    = df_filiais[[c for c in colunas if c in df_filiais.columns]].copy()

    # ── CNPJ: 14 dígitos ─────────────────────────────────────────────────────
    for _df in [df_nao_filial, df_filiais]:
        if "CNPJ" in _df.columns:
            _df["CNPJ"] = _df["CNPJ"].apply(_normaliza_cnpj)

    # ── classificação das não-filiais ─────────────────────────────────────────
    def _classifica_reinf(val):
        v = str(val).strip().upper() \
            if pd.notna(val) and str(val).strip() not in ("", "NAN") else ""
        if "CONCLUÍDA" in v or "CONCLUIDA" in v:
            return "Transmitida"
        return "Não Transmitida"

    col_transm = "TRANSMISSÃO REINF"

    if col_transm in df_nao_filial.columns:
        df_nao_filial[col_transm] = df_nao_filial[col_transm].apply(_classifica_reinf)
    else:
        df_nao_filial[col_transm] = "Não Transmitida"

    if col_transm in df_filiais.columns:
        df_filiais[col_transm] = "Filial"
    else:
        df_filiais[col_transm] = "Filial"

    # ── df final para a tabela ────────────────────────────────────────────────
    df_reinf = pd.concat([df_nao_filial, df_filiais], ignore_index=True)

    # ── contagens ─────────────────────────────────────────────────────────────
    transmitidas     = (df_nao_filial[col_transm] == "Transmitida").sum()
    nao_transmitidas = (df_nao_filial[col_transm] == "Não Transmitida").sum()
    filiais          = len(df_filiais)
    total            = transmitidas + nao_transmitidas

    st.markdown("<h2>REINF</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Transmitidas:</b> {transmitidas} &nbsp;|&nbsp; "
        f"<b>Não transmitidas:</b> {nao_transmitidas} &nbsp;|&nbsp; "
        f"<b>Filiais:</b> {filiais} &nbsp;|&nbsp; "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True,
    )

    # ── donut ─────────────────────────────────────────────────────────────────
    if "reinf_chart_key" not in st.session_state:
        st.session_state["reinf_chart_key"] = 0

    pct_t  = round(transmitidas     / total * 100) if total else 0
    pct_nt = round(nao_transmitidas / total * 100) if total else 0

    fig = go.Figure(data=[go.Pie(
        labels=["Transmitidas", "Não Transmitidas"],
        values=[int(transmitidas), int(nao_transmitidas)],
        hole=0.68,
        marker=dict(
            colors=["#2980b9", "#e67e22"],
            line=dict(color="#ffffff", width=3),
        ),
        textinfo="none",
        hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
        direction="clockwise",
        sort=False,
    )])

    fig.update_layout(
        paper_bgcolor="white", plot_bgcolor="white",
        showlegend=False,
        margin=dict(t=20, b=20, l=20, r=20),
        height=300,
        annotations=[dict(
            text=f"<b>{total}</b><br><span style='font-size:11px'>empresas</span>",
            x=0.5, y=0.5,
            xanchor="center", yanchor="middle",
            showarrow=False,
            font=dict(size=22, color="#1d3f77"),
        )],
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"chart_reinf_{st.session_state['reinf_chart_key']}",
    )

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#eaf4fb; "
            f"border-radius:8px; border-left:4px solid #2980b9;'>"
            f"<span style='font-size:22px; font-weight:700; color:#2980b9;'>{transmitidas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Transmitidas ({pct_t}%)</span></div>",
            unsafe_allow_html=True,
        )
    with col_r:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdf3e7; "
            f"border-radius:8px; border-left:4px solid #e67e22;'>"
            f"<span style='font-size:22px; font-weight:700; color:#e67e22;'>{nao_transmitidas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Não Transmitidas ({pct_nt}%)</span></div>",
            unsafe_allow_html=True,
        )
        if st.button("Ver empresas não transmitidas", use_container_width=True,
                     key="btn_nao_transmitidas_reinf"):
            df_nt = df_nao_filial[df_nao_filial[col_transm] == "Não Transmitida"]
            _modal_reinf_nao_transmitidas(df_nt)

    st.divider()

    # ── tabela principal ──────────────────────────────────────────────────────
    df_reinf = _sanitiza_df(df_reinf)
    exibe_aggrid(df_reinf, height=400, grid_key="grid_reinf")

    output = BytesIO()
    df_reinf.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel", data=output.getvalue(),
        file_name="reinf.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@st.dialog("DCTF WEB — Sem Procuração")
def _modal_dctf_sem_procuracao(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) sem procuração**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Regime",
                        "SITUAÇÃO DCTF", "MATRIZ / FILIAL"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.dialog("DCTF WEB — Não Concluídas")
def _modal_dctf_nao_concluidas(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) não concluída(s)**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Regime",
                        "SITUAÇÃO DCTF", "MATRIZ / FILIAL"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.fragment
def pagina_dctf_web():
    import plotly.graph_objects as go
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None or df.empty:
        st.warning("Nenhum dado encontrado.")
        return

    df = df.fillna("")
    df = df[df["Situação"].astype(str).str.upper() == "ATIVA"]
    if df.empty:
        st.warning("Nenhuma empresa ATIVA encontrada.")
        return

    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] \
        if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    _dt_comp = pd.to_datetime(competencia_raw, errors='coerce')
    competencia = _dt_comp.strftime("%m/%Y") if not pd.isna(_dt_comp) else ""

    # ── colunas para exibição ─────────────────────────────────────────────────
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "PERÍODO", "ORIGEM",
               "TIPO", "SITUAÇÃO DCTF", "MATRIZ / FILIAL", "Situação"]
    df_dctf = df[[c for c in colunas if c in df.columns]].copy()

    if "PERÍODO" in df_dctf.columns:
        df_dctf["PERÍODO"] = pd.to_datetime(
            df_dctf["PERÍODO"], errors="coerce"
        ).dt.strftime("%m-%Y").fillna("")

    if "CNPJ" in df_dctf.columns:
        df_dctf["CNPJ"] = df_dctf["CNPJ"].apply(_normaliza_cnpj)

    # ── classificação ─────────────────────────────────────────────────────────
    def _classifica_dctf(val):
        v = str(val).strip().upper()
        if "CONCLUÍDA" in v or "CONCLUIDA" in v or v == "ATIVA":
            return "Concluída"
        if "PROCURA" in v:
            return "Sem Procuração"
        return "Não Concluída"

    col_sit = "SITUAÇÃO DCTF"
    df_dctf["_status"] = df_dctf[col_sit].apply(_classifica_dctf) \
        if col_sit in df_dctf.columns else "Não Concluída"

    if "MATRIZ / FILIAL" in df_dctf.columns:
        mask_filial = df_dctf["MATRIZ / FILIAL"].astype(str).str.strip().str.upper() == "FILIAL"
        df_dctf.loc[mask_filial, "_status"] = "Filial"

    # ── contagens ─────────────────────────────────────────────────────────────
    concluidas     = (df_dctf["_status"] == "Concluída").sum()
    sem_procuracao = (df_dctf["_status"] == "Sem Procuração").sum()
    nao_concluidas = (df_dctf["_status"] == "Não Concluída").sum()
    filiais        = (df_dctf["_status"] == "Filial").sum()
    total          = concluidas + sem_procuracao + nao_concluidas

    # ── donut ─────────────────────────────────────────────────────────────────
    if "dctf_chart_key" not in st.session_state:
        st.session_state["dctf_chart_key"] = 0

    pct_c  = round(concluidas     / total * 100) if total else 0
    pct_sp = round(sem_procuracao / total * 100) if total else 0
    pct_nc = round(nao_concluidas / total * 100) if total else 0

    fig = go.Figure(data=[go.Pie(
        labels=["Concluídas", "Sem Procuração", "Não Concluídas"],
        values=[int(concluidas), int(sem_procuracao), int(nao_concluidas)],
        hole=0.68,
        marker=dict(
            colors=["#27ae60", "#e67e22", "#e74c3c"],
            line=dict(color="#ffffff", width=3),
        ),
        textinfo="none",
        hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
        direction="clockwise",
        sort=False,
    )])
    fig.update_layout(
        paper_bgcolor="white", plot_bgcolor="white",
        showlegend=False,
        margin=dict(t=20, b=20, l=20, r=20),
        height=300,
        annotations=[dict(
            text=f"<b>{total}</b><br><span style='font-size:11px'>empresas</span>",
            x=0.5, y=0.5,
            xanchor="center", yanchor="middle",
            showarrow=False,
            font=dict(size=22, color="#1d3f77"),
        )],
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        key=f"chart_dctf_{st.session_state['dctf_chart_key']}",
    )

    col_l, col_m, col_r = st.columns(3)
    with col_l:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#f0faf4; "
            f"border-radius:8px; border-left:4px solid #27ae60;'>"
            f"<span style='font-size:22px; font-weight:700; color:#27ae60;'>{concluidas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Concluídas ({pct_c}%)</span></div>",
            unsafe_allow_html=True,
        )
    with col_m:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdf3e7; "
            f"border-radius:8px; border-left:4px solid #e67e22;'>"
            f"<span style='font-size:22px; font-weight:700; color:#e67e22;'>{sem_procuracao}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Sem Procuração ({pct_sp}%)</span></div>",
            unsafe_allow_html=True,
        )
        if st.button("Ver sem procuração", use_container_width=True,
                     key="btn_dctf_sem_proc"):
            df_sp = df_dctf[df_dctf["_status"] == "Sem Procuração"]
            _modal_dctf_sem_procuracao(df_sp)
    with col_r:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdf2f2; "
            f"border-radius:8px; border-left:4px solid #e74c3c;'>"
            f"<span style='font-size:22px; font-weight:700; color:#e74c3c;'>{nao_concluidas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Não Concluídas ({pct_nc}%)</span></div>",
            unsafe_allow_html=True,
        )
        if st.button("Ver não concluídas", use_container_width=True,
                     key="btn_dctf_nao_conc"):
            df_nc = df_dctf[df_dctf["_status"] == "Não Concluída"]
            _modal_dctf_nao_concluidas(df_nc)

    st.divider()

    # ── tabela principal ──────────────────────────────────────────────────────
    df_dctf = _sanitiza_df(df_dctf.drop(columns=["_status"]))
    exibe_aggrid(df_dctf, height=400, grid_key="grid_dctf")

    output = BytesIO()
    df_dctf.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel", data=output.getvalue(),
        file_name="dctf_web.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@st.dialog("DMS — Sem Acesso")
def _modal_dms_sem_acesso(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) sem acesso**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Município", "Estado", "DMS"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.dialog("GUIA ISS DMS — Com Imposto")
def _modal_dms_com_imposto(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) com imposto**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Município", "Estado",
                        "DMS", "GUIA ISS DMS"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.fragment
def pagina_dms():
    import plotly.graph_objects as go
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] \
        if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors="coerce").strftime("%m/%Y") \
        if competencia_raw else ""

    if "Situação" not in df.columns:
        st.error("Coluna 'Situação' não encontrada.")
        return

    df_dms = df[df["Situação"].astype(str).str.upper() == "ATIVA"].copy()

    if df_dms.empty:
        st.warning("Nenhuma empresa ATIVA encontrada para DMS.")
        return

    # ── colunas para exibição ─────────────────────────────────────────────────
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado",
               "DMS", "GUIA ISS DMS", "Situação"]
    df_dms = df_dms[[c for c in colunas if c in df_dms.columns]].copy()

    # ── CNPJ: 14 dígitos ─────────────────────────────────────────────────────
    if "CNPJ" in df_dms.columns:
        df_dms["CNPJ"] = df_dms["CNPJ"].apply(_normaliza_cnpj)

    # ── classificação DMS (coluna AC) ─────────────────────────────────────────
    def _classifica_dms(val):
        v = str(val).strip().upper() \
            if pd.notna(val) and str(val).strip() not in ("", "NAN") else ""
        if "SEM ACESSO" in v:
            return "Sem Acesso"
        return "Concluída"   # "Concluída" ou em branco

    if "DMS" in df_dms.columns:
        df_dms["DMS"] = df_dms["DMS"].apply(_classifica_dms)
    else:
        df_dms["DMS"] = "Concluída"

    # ── classificação GUIA ISS DMS (coluna AD) ────────────────────────────────
    def _classifica_guia(val):
        v = str(val).strip().upper() \
            if pd.notna(val) and str(val).strip() not in ("", "NAN") else ""
        if v == "SIM":
            return "Com Imposto"
        return "Sem Imposto"

    if "GUIA ISS DMS" in df_dms.columns:
        df_dms["GUIA ISS DMS"] = df_dms["GUIA ISS DMS"].apply(_classifica_guia)
    else:
        df_dms["GUIA ISS DMS"] = "Sem Imposto"

    # ── contagens DMS ─────────────────────────────────────────────────────────
    concluidas  = (df_dms["DMS"] == "Concluída").sum()
    sem_acesso  = (df_dms["DMS"] == "Sem Acesso").sum()
    total_dms   = concluidas + sem_acesso

    # ── contagens GUIA ISS ────────────────────────────────────────────────────
    com_imposto  = (df_dms["GUIA ISS DMS"] == "Com Imposto").sum()
    sem_imposto  = (df_dms["GUIA ISS DMS"] == "Sem Imposto").sum()
    total_guia   = com_imposto + sem_imposto

    st.markdown("<h2>DMS</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Concluídas:</b> {concluidas} &nbsp;|&nbsp; "
        f"<b>Sem Acesso:</b> {sem_acesso} &nbsp;|&nbsp; "
        f"<b>Com Imposto:</b> {com_imposto} &nbsp;|&nbsp; "
        f"<b>Sem Imposto:</b> {sem_imposto} &nbsp;|&nbsp; "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True,
    )

    # ── session keys ─────────────────────────────────────────────────────────
    if "dms_chart_key" not in st.session_state:
        st.session_state["dms_chart_key"] = 0
    if "guia_chart_key" not in st.session_state:
        st.session_state["guia_chart_key"] = 0

    pct_c  = round(concluidas  / total_dms  * 100) if total_dms  else 0
    pct_sa = round(sem_acesso  / total_dms  * 100) if total_dms  else 0
    pct_ci = round(com_imposto / total_guia * 100) if total_guia else 0
    pct_si = round(sem_imposto / total_guia * 100) if total_guia else 0

    # ── dois donuts lado a lado ───────────────────────────────────────────────
    col_d1, col_d2 = st.columns(2)

    # ── donut 1: DMS ─────────────────────────────────────────────────────────
    with col_d1:
        st.markdown("<h4 style='text-align:center; color:#1d3f77;'>DMS</h4>",
                    unsafe_allow_html=True)

        fig1 = go.Figure(data=[go.Pie(
            labels=["Concluídas", "Sem Acesso"],
            values=[int(concluidas), int(sem_acesso)],
            hole=0.68,
            marker=dict(
                colors=["#8e44ad", "#7f8c8d"],
                line=dict(color="#ffffff", width=3),
            ),
            textinfo="none",
            hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
            direction="clockwise",
            sort=False,
        )])
        fig1.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            showlegend=False,
            margin=dict(t=10, b=10, l=10, r=10),
            height=260,
            annotations=[dict(
                text=f"<b>{total_dms}</b><br><span style='font-size:11px'>empresas</span>",
                x=0.5, y=0.5,
                xanchor="center", yanchor="middle",
                showarrow=False,
                font=dict(size=20, color="#1d3f77"),
            )],
        )
        st.plotly_chart(fig1, use_container_width=True,
                        key=f"chart_dms_{st.session_state['dms_chart_key']}")

        cl1, cl2 = st.columns(2)
        with cl1:
            st.markdown(
                f"<div style='text-align:center; padding:8px; background:#f5eef8; "
                f"border-radius:8px; border-left:4px solid #8e44ad;'>"
                f"<span style='font-size:20px; font-weight:700; color:#8e44ad;'>{concluidas}</span><br>"
                f"<span style='font-size:12px; color:#555;'>Concluídas ({pct_c}%)</span></div>",
                unsafe_allow_html=True,
            )
        with cl2:
            st.markdown(
                f"<div style='text-align:center; padding:8px; background:#f2f3f4; "
                f"border-radius:8px; border-left:4px solid #7f8c8d;'>"
                f"<span style='font-size:20px; font-weight:700; color:#7f8c8d;'>{sem_acesso}</span><br>"
                f"<span style='font-size:12px; color:#555;'>Sem Acesso ({pct_sa}%)</span></div>",
                unsafe_allow_html=True,
            )
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Ver empresas sem acesso", use_container_width=True,
                     key="btn_dms_sem_acesso"):
            _modal_dms_sem_acesso(df_dms[df_dms["DMS"] == "Sem Acesso"])

    # ── donut 2: GUIA ISS DMS ────────────────────────────────────────────────
    with col_d2:
        st.markdown("<h4 style='text-align:center; color:#1d3f77;'>Guia ISS DMS</h4>",
                    unsafe_allow_html=True)

        fig2 = go.Figure(data=[go.Pie(
            labels=["Com Imposto", "Sem Imposto"],
            values=[int(com_imposto), int(sem_imposto)],
            hole=0.68,
            marker=dict(
                colors=["#16a085", "#bdc3c7"],
                line=dict(color="#ffffff", width=3),
            ),
            textinfo="none",
            hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
            direction="clockwise",
            sort=False,
        )])
        fig2.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            showlegend=False,
            margin=dict(t=10, b=10, l=10, r=10),
            height=260,
            annotations=[dict(
                text=f"<b>{total_guia}</b><br><span style='font-size:11px'>empresas</span>",
                x=0.5, y=0.5,
                xanchor="center", yanchor="middle",
                showarrow=False,
                font=dict(size=20, color="#1d3f77"),
            )],
        )
        st.plotly_chart(fig2, use_container_width=True,
                        key=f"chart_guia_{st.session_state['guia_chart_key']}")

        cg1, cg2 = st.columns(2)
        with cg1:
            st.markdown(
                f"<div style='text-align:center; padding:8px; background:#e8f8f5; "
                f"border-radius:8px; border-left:4px solid #16a085;'>"
                f"<span style='font-size:20px; font-weight:700; color:#16a085;'>{com_imposto}</span><br>"
                f"<span style='font-size:12px; color:#555;'>Com Imposto ({pct_ci}%)</span></div>",
                unsafe_allow_html=True,
            )
        with cg2:
            st.markdown(
                f"<div style='text-align:center; padding:8px; background:#f8f9f9; "
                f"border-radius:8px; border-left:4px solid #bdc3c7;'>"
                f"<span style='font-size:20px; font-weight:700; color:#7f8c8d;'>{sem_imposto}</span><br>"
                f"<span style='font-size:12px; color:#555;'>Sem Imposto ({pct_si}%)</span></div>",
                unsafe_allow_html=True,
            )
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Ver empresas com imposto", use_container_width=True,
                     key="btn_guia_com_imposto"):
            _modal_dms_com_imposto(df_dms[df_dms["GUIA ISS DMS"] == "Com Imposto"])

    st.divider()

    # ── tabela principal ──────────────────────────────────────────────────────
    df_dms = _sanitiza_df(df_dms)
    exibe_aggrid(df_dms, height=400, grid_key="grid_dms")

    output = BytesIO()
    df_dms.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel", data=output.getvalue(),
        file_name="dms.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@st.dialog("SERVIÇOS TOMADOS — Sem Acesso")
def _modal_rest_sem_acesso(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) sem acesso**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Município", "Estado", "REST"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)




@st.dialog("GUIA ISS REST — Com Imposto")
def _modal_rest_com_imposto(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) com imposto**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Município", "Estado",
                        "REST", "GUIA ISS REST"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.fragment
def pagina_rest():
    import plotly.graph_objects as go
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] \
        if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors="coerce").strftime("%m/%Y") \
        if competencia_raw else ""

    if "Situação" not in df.columns:
        st.error("Coluna 'Situação' não encontrada.")
        return

    df_rest = df[df["Situação"].astype(str).str.upper() == "ATIVA"].copy()

    if df_rest.empty:
        st.warning("Nenhuma empresa ATIVA encontrada para SERVIÇOS TOMADOS.")
        return

    # ── colunas para exibição ─────────────────────────────────────────────────
    colunas = ["Código", "Razão Social", "CNPJ", "Regime", "Município", "Estado",
               "REST", "GUIA ISS REST", "Situação"]
    df_rest = df_rest[[c for c in colunas if c in df_rest.columns]].copy()

    # ── CNPJ: 14 dígitos ─────────────────────────────────────────────────────
    if "CNPJ" in df_rest.columns:
        df_rest["CNPJ"] = df_rest["CNPJ"].apply(_normaliza_cnpj)

    # ── classificação REST (coluna AE) ────────────────────────────────────────
    def _classifica_rest(val):
        v = str(val).strip().upper() \
            if pd.notna(val) and str(val).strip() not in ("", "NAN") else ""
        if "SEM ACESSO" in v:
            return "Sem Acesso"
        return "Concluída"

    if "REST" in df_rest.columns:
        df_rest["REST"] = df_rest["REST"].apply(_classifica_rest)
    else:
        df_rest["REST"] = "Concluída"

    # ── classificação GUIA ISS REST (coluna AG) ───────────────────────────────
    def _classifica_guia_rest(val):
        v = str(val).strip().upper() \
            if pd.notna(val) and str(val).strip() not in ("", "NAN") else ""
        if v == "SIM":
            return "Com Imposto"
        return "Sem Imposto"

    if "GUIA ISS REST" in df_rest.columns:
        df_rest["GUIA ISS REST"] = df_rest["GUIA ISS REST"].apply(_classifica_guia_rest)
    else:
        df_rest["GUIA ISS REST"] = "Sem Imposto"

    # ── contagens REST ────────────────────────────────────────────────────────
    concluidas       = (df_rest["REST"] == "Concluída").sum()
    sem_acesso       = (df_rest["REST"] == "Sem Acesso").sum()
    total_rest       = concluidas + sem_acesso

    # ── contagens GUIA ISS ────────────────────────────────────────────────────
    com_imposto  = (df_rest["GUIA ISS REST"] == "Com Imposto").sum()
    sem_imposto  = (df_rest["GUIA ISS REST"] == "Sem Imposto").sum()
    total_guia   = com_imposto + sem_imposto

    st.markdown("<h2>SERVIÇOS TOMADOS</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Concluídas:</b> {concluidas} &nbsp;|&nbsp; "
        f"<b>Sem Acesso:</b> {sem_acesso} &nbsp;|&nbsp; "
        f"<b>Com Imposto:</b> {com_imposto} &nbsp;|&nbsp; "
        f"<b>Sem Imposto:</b> {sem_imposto} &nbsp;|&nbsp; "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True,
    )

    # ── session keys ──────────────────────────────────────────────────────────
    if "rest_chart_key" not in st.session_state:
        st.session_state["rest_chart_key"] = 0
    if "guia_rest_chart_key" not in st.session_state:
        st.session_state["guia_rest_chart_key"] = 0

    pct_c   = round(concluidas  / total_rest * 100) if total_rest else 0
    pct_sa  = round(sem_acesso  / total_rest * 100) if total_rest else 0
    pct_ci  = round(com_imposto / total_guia * 100) if total_guia else 0
    pct_si  = round(sem_imposto / total_guia * 100) if total_guia else 0

    # ── dois donuts lado a lado ───────────────────────────────────────────────
    col_d1, col_d2 = st.columns(2)

    # ── donut 1: REST ─────────────────────────────────────────────────────────
    with col_d1:
        st.markdown("<h4 style='text-align:center; color:#1d3f77;'>Serviços Tomados</h4>",
                    unsafe_allow_html=True)

        fig1 = go.Figure(data=[go.Pie(
            labels=["Concluídas", "Sem Acesso"],
            values=[int(concluidas), int(sem_acesso)],
            hole=0.68,
            marker=dict(
                colors=["#d35400", "#7f8c8d"],
                line=dict(color="#ffffff", width=3),
            ),
            textinfo="none",
            hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
            direction="clockwise",
            sort=False,
        )])
        fig1.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            showlegend=False,
            margin=dict(t=10, b=10, l=10, r=10),
            height=260,
            annotations=[dict(
                text=f"<b>{total_rest}</b><br><span style='font-size:11px'>empresas</span>",
                x=0.5, y=0.5,
                xanchor="center", yanchor="middle",
                showarrow=False,
                font=dict(size=20, color="#1d3f77"),
            )],
        )
        st.plotly_chart(fig1, use_container_width=True,
                        key=f"chart_rest_{st.session_state['rest_chart_key']}")

        cl1, cl2 = st.columns(2)
        with cl1:
            st.markdown(
                f"<div style='text-align:center; padding:8px; background:#fdf0e8; "
                f"border-radius:8px; border-left:4px solid #d35400;'>"
                f"<span style='font-size:18px; font-weight:700; color:#d35400;'>{concluidas}</span><br>"
                f"<span style='font-size:11px; color:#555;'>Concluídas ({pct_c}%)</span></div>",
                unsafe_allow_html=True,
            )
        with cl2:
            st.markdown(
                f"<div style='text-align:center; padding:8px; background:#f2f3f4; "
                f"border-radius:8px; border-left:4px solid #7f8c8d;'>"
                f"<span style='font-size:18px; font-weight:700; color:#7f8c8d;'>{sem_acesso}</span><br>"
                f"<span style='font-size:11px; color:#555;'>Sem Acesso ({pct_sa}%)</span></div>",
                unsafe_allow_html=True,
            )

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Ver sem acesso", use_container_width=True,
                     key="btn_rest_sem_acesso"):
            _modal_rest_sem_acesso(df_rest[df_rest["REST"] == "Sem Acesso"])

    # ── donut 2: GUIA ISS REST ────────────────────────────────────────────────
    with col_d2:
        st.markdown("<h4 style='text-align:center; color:#1d3f77;'>Guia ISS REST</h4>",
                    unsafe_allow_html=True)

        fig2 = go.Figure(data=[go.Pie(
            labels=["Com Imposto", "Sem Imposto"],
            values=[int(com_imposto), int(sem_imposto)],
            hole=0.68,
            marker=dict(
                colors=["#1abc9c", "#bdc3c7"],
                line=dict(color="#ffffff", width=3),
            ),
            textinfo="none",
            hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
            direction="clockwise",
            sort=False,
        )])
        fig2.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            showlegend=False,
            margin=dict(t=10, b=10, l=10, r=10),
            height=260,
            annotations=[dict(
                text=f"<b>{total_guia}</b><br><span style='font-size:11px'>empresas</span>",
                x=0.5, y=0.5,
                xanchor="center", yanchor="middle",
                showarrow=False,
                font=dict(size=20, color="#1d3f77"),
            )],
        )
        st.plotly_chart(fig2, use_container_width=True,
                        key=f"chart_guia_rest_{st.session_state['guia_rest_chart_key']}")

        cg1, cg2 = st.columns(2)
        with cg1:
            st.markdown(
                f"<div style='text-align:center; padding:8px; background:#e8faf5; "
                f"border-radius:8px; border-left:4px solid #1abc9c;'>"
                f"<span style='font-size:20px; font-weight:700; color:#1abc9c;'>{com_imposto}</span><br>"
                f"<span style='font-size:12px; color:#555;'>Com Imposto ({pct_ci}%)</span></div>",
                unsafe_allow_html=True,
            )
        with cg2:
            st.markdown(
                f"<div style='text-align:center; padding:8px; background:#f8f9f9; "
                f"border-radius:8px; border-left:4px solid #bdc3c7;'>"
                f"<span style='font-size:20px; font-weight:700; color:#7f8c8d;'>{sem_imposto}</span><br>"
                f"<span style='font-size:12px; color:#555;'>Sem Imposto ({pct_si}%)</span></div>",
                unsafe_allow_html=True,
            )

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Ver empresas com imposto", use_container_width=True,
                     key="btn_guia_rest_com_imposto"):
            _modal_rest_com_imposto(df_rest[df_rest["GUIA ISS REST"] == "Com Imposto"])

    st.divider()

    # ── tabela principal ──────────────────────────────────────────────────────
    df_rest = _sanitiza_df(df_rest)
    exibe_aggrid(df_rest, height=400, grid_key="grid_rest")

    output = BytesIO()
    df_rest.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel", data=output.getvalue(),
        file_name="servicos_tomados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@st.dialog("SEFAZ — Sem Acesso")
def _modal_sefaz_sem_acesso(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) sem acesso**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Estado", "Insc. Estadual", "IMPORTAÇÃO"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.dialog("SEFAZ — Sem Busca")
def _modal_sefaz_sem_busca(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) sem busca**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Estado", "Insc. Estadual", "IMPORTAÇÃO"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.dialog("SEFAZ — Sem Movimento")
def _modal_sefaz_sem_movimento(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) sem movimento**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Estado", "Insc. Estadual", "IMPORTAÇÃO"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.dialog("SEFAZ — Divergentes")
def _modal_sefaz_divergentes(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) com quantidade divergente**")
    cols = [c for c in ["Código", "Razão Social", "CNPJ", "Estado", "Insc. Estadual",
                        "ENTRADAS SEFAZ", "SAÍDAS SEFAZ", "TOTAL SEFAZ", "TOTAL DOMÍNIO", "MOTIVO"]
            if c in df_show.columns]
    df_exib = df_show[cols].copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_normaliza_cnpj)
    st.dataframe(df_exib.reset_index(drop=True),
                 use_container_width=True, hide_index=True)


@st.fragment
def pagina_sefaz():
    import plotly.graph_objects as go
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] \
        if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors="coerce").strftime("%m/%Y") \
        if competencia_raw else ""

    if "Situação" not in df.columns:
        st.error("Coluna 'Situação' não encontrada.")
        return

    df_sefaz = df[df["Situação"].astype(str).str.upper() == "ATIVA"].copy()

    if df_sefaz.empty:
        st.warning("Nenhuma empresa ATIVA encontrada para SEFAZ.")
        return

    # ── colunas para exibição ─────────────────────────────────────────────────
    colunas = ["Código", "Razão Social", "CNPJ", "Estado", "Insc. Estadual",
               "IMPORTAÇÃO", "TOTAL ENTRADA", "TOTAL SAÍDA", "EMISSÃO TERCEIRO",
               "PERCA", "TOTAL DOMÍNIO", "MOTIVO DIFERENÇA SEFAZ", "Situação"]
    df_sefaz = df_sefaz[[c for c in colunas if c in df_sefaz.columns]].copy()

    # ── CNPJ: 14 dígitos ─────────────────────────────────────────────────────
    if "CNPJ" in df_sefaz.columns:
        df_sefaz["CNPJ"] = df_sefaz["CNPJ"].apply(_normaliza_cnpj)

    # ── colunas numéricas ─────────────────────────────────────────────────────
    for col in ["TOTAL ENTRADA", "TOTAL SAÍDA", "TOTAL DOMÍNIO", "EMISSÃO TERCEIRO", "PERCA"]:
        if col in df_sefaz.columns:
            df_sefaz[col] = pd.to_numeric(df_sefaz[col], errors="coerce").fillna(0)

    # ── TOTAL SEFAZ = ENTRADAS + SAÍDAS - EMISSÃO TERCEIRO - PERCA ────────────
    # (no escritório VISÃO a regra é diferente: ENTRADAS + SAÍDAS - PERCA, sem
    # subtrair EMISSÃO TERCEIRO — ver LEIA-ME.md, "Diferenças conhecidas")
    for c in ["TOTAL ENTRADA", "TOTAL SAÍDA", "EMISSÃO TERCEIRO", "PERCA"]:
        if c not in df_sefaz.columns:
            df_sefaz[c] = 0
    df_sefaz["TOTAL SEFAZ"] = (
        df_sefaz["TOTAL ENTRADA"] + df_sefaz["TOTAL SAÍDA"]
        - df_sefaz["EMISSÃO TERCEIRO"] - df_sefaz["PERCA"]
    )

    # ── coluna Confronto: TOTAL SEFAZ (ajustado) x TOTAL DOMÍNIO ──────────────
    if "TOTAL DOMÍNIO" in df_sefaz.columns:
        total_dom    = df_sefaz["TOTAL DOMÍNIO"]
        total_sefaz  = df_sefaz["TOTAL SEFAZ"]
        df_sefaz["Confronto"] = [
            "Importação OK" if total_sefaz[i] == total_dom[i] else "Quantidade Diferente"
            for i in df_sefaz.index
        ]
    else:
        df_sefaz["Confronto"] = "Importação OK"

    # ── renomeia colunas para exibição ────────────────────────────────────────
    df_sefaz = df_sefaz.rename(columns={
        "TOTAL ENTRADA": "ENTRADAS SEFAZ",
        "TOTAL SAÍDA": "SAÍDAS SEFAZ",
        "MOTIVO DIFERENÇA SEFAZ": "MOTIVO",
    })

    # ── ordem final das colunas ───────────────────────────────────────────────
    ordem = ["Código", "Razão Social", "CNPJ", "Estado", "Insc. Estadual",
             "IMPORTAÇÃO", "ENTRADAS SEFAZ", "SAÍDAS SEFAZ", "EMISSÃO TERCEIRO",
             "TOTAL SEFAZ", "PERCA", "TOTAL DOMÍNIO", "Confronto", "MOTIVO",
             "Situação"]
    df_sefaz = df_sefaz[[c for c in ordem if c in df_sefaz.columns]]

    # ── classificação IMPORTAÇÃO (coluna BS) ──────────────────────────────────  ← CONTINUA IGUAL
    def _classifica_sefaz(val):
        v = str(val).strip().upper() \
            if pd.notna(val) and str(val).strip() not in ("", "NAN") else ""
        if "SEM ACESSO" in v:
            return "Sem Acesso"
        if "SEM BUSCA" in v:
            return "Sem Busca"
        if "SEM MOVIMENTO" in v:
            return "Sem Movimento"
        return "Com Movimento"   # COM MOVIMENTO ou qualquer outro valor

    if "IMPORTAÇÃO" in df_sefaz.columns:
        df_sefaz["IMPORTAÇÃO"] = df_sefaz["IMPORTAÇÃO"].apply(_classifica_sefaz)
    else:
        df_sefaz["IMPORTAÇÃO"] = "Com Movimento"

    # ── contagens ─────────────────────────────────────────────────────────────
    com_movimento  = (df_sefaz["IMPORTAÇÃO"] == "Com Movimento").sum()
    sem_acesso     = (df_sefaz["IMPORTAÇÃO"] == "Sem Acesso").sum()
    sem_busca      = (df_sefaz["IMPORTAÇÃO"] == "Sem Busca").sum()
    sem_movimento  = (df_sefaz["IMPORTAÇÃO"] == "Sem Movimento").sum()
    total          = com_movimento + sem_acesso + sem_busca + sem_movimento
    divergentes    = int((df_sefaz["Confronto"] == "Quantidade Diferente").sum()) \
        if "Confronto" in df_sefaz.columns else 0

    st.markdown("<h2>SEFAZ</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Com Movimento:</b> {com_movimento} &nbsp;|&nbsp; "
        f"<b>Sem Acesso:</b> {sem_acesso} &nbsp;|&nbsp; "
        f"<b>Sem Busca:</b> {sem_busca} &nbsp;|&nbsp; "
        f"<b>Sem Movimento:</b> {sem_movimento} &nbsp;|&nbsp; "
        f"<b>Divergentes:</b> {divergentes} &nbsp;|&nbsp; "
        f"<b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True,
    )

    # ── session key ───────────────────────────────────────────────────────────
    if "sefaz_chart_key" not in st.session_state:
        st.session_state["sefaz_chart_key"] = 0

    pct_cm = round(com_movimento / total * 100) if total else 0
    pct_sa = round(sem_acesso    / total * 100) if total else 0
    pct_sb = round(sem_busca     / total * 100) if total else 0
    pct_sm = round(sem_movimento / total * 100) if total else 0

    # ── donut centralizado ────────────────────────────────────────────────────
    col_esq, col_centro, col_dir = st.columns([1, 2, 1])
    with col_centro:
        st.markdown("<h4 style='text-align:center; color:#1d3f77;'>Importação SEFAZ</h4>",
                    unsafe_allow_html=True)

        fig = go.Figure(data=[go.Pie(
            labels=["Com Movimento", "Sem Acesso", "Sem Busca", "Sem Movimento"],
            values=[int(com_movimento), int(sem_acesso),
                    int(sem_busca),     int(sem_movimento)],
            hole=0.68,
            marker=dict(
                colors=["#2471a3", "#c0392b", "#e67e22", "#7f8c8d"],
                line=dict(color="#ffffff", width=3),
            ),
            textinfo="none",
            hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
            direction="clockwise",
            sort=False,
        )])
        fig.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            showlegend=False,
            margin=dict(t=10, b=10, l=10, r=10),
            height=300,
            annotations=[dict(
                text=f"<b>{total}</b><br><span style='font-size:11px'>empresas</span>",
                x=0.5, y=0.5,
                xanchor="center", yanchor="middle",
                showarrow=False,
                font=dict(size=22, color="#1d3f77"),
            )],
        )
        st.plotly_chart(fig, use_container_width=True,
                        key=f"chart_sefaz_{st.session_state['sefaz_chart_key']}")

    # ── cards com os 4 status ─────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#eaf4fb; "
            f"border-radius:8px; border-left:4px solid #2471a3;'>"
            f"<span style='font-size:20px; font-weight:700; color:#2471a3;'>{com_movimento}</span><br>"
            f"<span style='font-size:12px; color:#555;'>Com Movimento ({pct_cm}%)</span></div>",
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdedec; "
            f"border-radius:8px; border-left:4px solid #c0392b;'>"
            f"<span style='font-size:20px; font-weight:700; color:#c0392b;'>{sem_acesso}</span><br>"
            f"<span style='font-size:12px; color:#555;'>Sem Acesso ({pct_sa}%)</span></div>",
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdf3e7; "
            f"border-radius:8px; border-left:4px solid #e67e22;'>"
            f"<span style='font-size:20px; font-weight:700; color:#e67e22;'>{sem_busca}</span><br>"
            f"<span style='font-size:12px; color:#555;'>Sem Busca ({pct_sb}%)</span></div>",
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#f2f3f4; "
            f"border-radius:8px; border-left:4px solid #7f8c8d;'>"
            f"<span style='font-size:20px; font-weight:700; color:#7f8c8d;'>{sem_movimento}</span><br>"
            f"<span style='font-size:12px; color:#555;'>Sem Movimento ({pct_sm}%)</span></div>",
            unsafe_allow_html=True,
        )

    # ── card Confronto (divergentes SEFAZ x Domínio) ──────────────────────────
    pct_dv = round(divergentes / total * 100) if total else 0
    st.markdown("<br>", unsafe_allow_html=True)
    col_dv_esq, col_dv_centro, col_dv_dir = st.columns([1, 2, 1])
    with col_dv_centro:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdedec; "
            f"border-radius:8px; border-left:4px solid #c0392b;'>"
            f"<span style='font-size:20px; font-weight:700; color:#c0392b;'>{divergentes}</span><br>"
            f"<span style='font-size:12px; color:#555;'>Confronto Divergente ({pct_dv}%)</span></div>",
            unsafe_allow_html=True,
        )

    # ── botões das listas ─────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    b1, b2, b3, b4 = st.columns(4)
    with b1:
        if st.button("Ver sem acesso", use_container_width=True,
                     key="btn_sefaz_sem_acesso"):
            _modal_sefaz_sem_acesso(df_sefaz[df_sefaz["IMPORTAÇÃO"] == "Sem Acesso"])
    with b2:
        if st.button("Ver sem busca", use_container_width=True,
                     key="btn_sefaz_sem_busca"):
            _modal_sefaz_sem_busca(df_sefaz[df_sefaz["IMPORTAÇÃO"] == "Sem Busca"])
    with b3:
        if st.button("Ver sem movimento", use_container_width=True,
                     key="btn_sefaz_sem_movimento"):
            _modal_sefaz_sem_movimento(df_sefaz[df_sefaz["IMPORTAÇÃO"] == "Sem Movimento"])
    with b4:
        if st.button("Ver divergentes", use_container_width=True,
                     key="btn_sefaz_divergentes"):
            _modal_sefaz_divergentes(df_sefaz[df_sefaz["Confronto"] == "Quantidade Diferente"])

    st.divider()

    # ── tabela principal ──────────────────────────────────────────────────────
    df_sefaz = _sanitiza_df(df_sefaz)
    exibe_aggrid(df_sefaz, height=400, grid_key="grid_sefaz")

    output = BytesIO()
    df_sefaz.to_excel(output, index=False)
    st.download_button(
        "Baixar Excel", data=output.getvalue(),
        file_name="sefaz.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def pagina_cnd_municipal():
    import plotly.graph_objects as go
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    competencia_raw = df["PERÍODO DE COMPETÊNCIA"].iloc[0] \
        if "PERÍODO DE COMPETÊNCIA" in df.columns else ""
    competencia = pd.to_datetime(competencia_raw, errors="coerce").strftime("%m/%Y") \
        if competencia_raw else ""

    df_cnd = df[df["Situação"].astype(str).str.upper() == "ATIVA"].copy() \
        if "Situação" in df.columns else pd.DataFrame()
    if df_cnd.empty:
        st.warning("Nenhuma empresa ATIVA encontrada.")
        return

    colunas_solicitadas = ["Código", "Razão Social", "CNPJ", "Município", "Estado",
                           "SITUAÇÃO CND MUNICIPAL", "VALIDADE", "LINK CND MUNICIPAL", "Situação"]
    df_cnd = df_cnd[[c for c in colunas_solicitadas if c in df_cnd.columns]].copy()

    if "VALIDADE" in df_cnd.columns:
        df_cnd["VALIDADE"] = pd.to_datetime(df_cnd["VALIDADE"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")

    if "CNPJ" in df_cnd.columns:
        df_cnd["CNPJ"] = df_cnd["CNPJ"].apply(_normaliza_cnpj)

    # ── Contagens ─────────────────────────────────────────────────────────────
    if "SITUAÇÃO CND MUNICIPAL" in df_cnd.columns:
        sit_u = df_cnd["SITUAÇÃO CND MUNICIPAL"].fillna("").astype(str).str.strip().str.upper()
        positivas           = (sit_u == "POSITIVA").sum()
        negativas           = (sit_u == "NEGATIVA").sum()
        positiva_efeito_neg = (sit_u == "POSITIVA COM EFEITO NEGATIVA").sum()
        nao_geradas         = sit_u.isin(["", "NAN"]).sum()
    else:
        positivas = negativas = positiva_efeito_neg = nao_geradas = 0

    total_geral = df_cnd.shape[0]

    # ── Cabeçalho ─────────────────────────────────────────────────────────────
    st.markdown("<h2>CND MUNICIPAL</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Total:</b> {total_geral} &nbsp;|&nbsp; <b>Competência:</b> {competencia}</p>",
        unsafe_allow_html=True,
    )

    # ── Donut ─────────────────────────────────────────────────────────────────
    total_dash = int(positivas + negativas + positiva_efeito_neg + nao_geradas)
    if total_dash > 0:
        fig = go.Figure(data=[go.Pie(
            labels=["Positivas", "Negativas", "Positiva c/ Efeito Neg.", "Não Geradas"],
            values=[int(positivas), int(negativas), int(positiva_efeito_neg), int(nao_geradas)],
            hole=0.68,
            marker=dict(colors=["#e74c3c", "#27ae60", "#f39c12", "#bdc3c7"],
                        line=dict(color="#ffffff", width=3)),
            textinfo="none",
            hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
            direction="clockwise", sort=False,
        )])
        fig.update_layout(
            paper_bgcolor="white", plot_bgcolor="white", showlegend=False,
            margin=dict(t=20, b=20, l=20, r=20), height=300,
            annotations=[dict(
                text=f"<b>{total_geral}</b><br><span style='font-size:11px'>empresas</span>",
                x=0.5, y=0.5, xanchor="center", yanchor="middle",
                showarrow=False, font=dict(size=22, color="#1d3f77"),
            )],
        )
        st.plotly_chart(fig, use_container_width=True, key="chart_cnd_municipal_donut")

    # ── Cards ─────────────────────────────────────────────────────────────────
    c1, c2, c3, c4 = st.columns(4)

    _cols_modal = ["Código", "Razão Social", "CNPJ", "Município", "SITUAÇÃO CND MUNICIPAL", "VALIDADE"]

    with c1:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdf2f2; border-radius:8px; border-left:4px solid #e74c3c;'>"
            f"<span style='font-size:22px; font-weight:700; color:#e74c3c;'>{positivas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Positivas</span></div>",
            unsafe_allow_html=True,
        )
        if st.button("Ver Positivas", key="btn_cnd_pos", use_container_width=True):
            df_pos = df_cnd[df_cnd["SITUAÇÃO CND MUNICIPAL"].fillna("").astype(str).str.strip().str.upper() == "POSITIVA"]
            _modal_dashboard("Positivas", df_pos, _cols_modal)

    with c2:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#eafaf1; border-radius:8px; border-left:4px solid #27ae60;'>"
            f"<span style='font-size:22px; font-weight:700; color:#27ae60;'>{negativas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Negativas</span></div>",
            unsafe_allow_html=True,
        )
        if st.button("Ver Negativas", key="btn_cnd_neg", use_container_width=True):
            df_neg = df_cnd[df_cnd["SITUAÇÃO CND MUNICIPAL"].fillna("").astype(str).str.strip().str.upper() == "NEGATIVA"]
            _modal_dashboard("Negativas", df_neg, _cols_modal)

    with c3:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fef9e7; border-radius:8px; border-left:4px solid #f39c12;'>"
            f"<span style='font-size:22px; font-weight:700; color:#f39c12;'>{positiva_efeito_neg}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Positiva c/ Efeito Neg.</span></div>",
            unsafe_allow_html=True,
        )
        if st.button("Ver Positiva c/ Efeito Neg.", key="btn_cnd_pen", use_container_width=True):
            df_pen = df_cnd[df_cnd["SITUAÇÃO CND MUNICIPAL"].fillna("").astype(str).str.strip().str.upper() == "POSITIVA COM EFEITO NEGATIVA"]
            _modal_dashboard("Positiva c/ Efeito Negativa", df_pen, _cols_modal)

    with c4:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#f8f9fa; border-radius:8px; border-left:4px solid #bdc3c7;'>"
            f"<span style='font-size:22px; font-weight:700; color:#7f8c8d;'>{nao_geradas}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Não Geradas</span></div>",
            unsafe_allow_html=True,
        )
        if st.button("Ver Não Geradas", key="btn_cnd_ng", use_container_width=True):
            sit_u2 = df_cnd["SITUAÇÃO CND MUNICIPAL"].fillna("").astype(str).str.strip().str.upper()
            df_ng = df_cnd[sit_u2.isin(["", "NAN"])]
            _modal_dashboard("Não Geradas", df_ng, ["Código", "Razão Social", "CNPJ", "Município"])

    st.divider()

    # ── Visualizador de PDF / grade ───────────────────────────────────────────
    if "visualizando_pdf" not in st.session_state:
        st.session_state.visualizando_pdf = False
        st.session_state.pdf_selecionado = None

    if st.session_state.visualizando_pdf and st.session_state.pdf_selecionado:
        col1, col2 = st.columns([6, 1])
        with col1:
            if st.button("← Voltar para a lista", type="primary"):
                st.session_state.visualizando_pdf = False
                st.session_state.pdf_selecionado = None
                st.rerun()
        with col2:
            row = st.session_state.pdf_selecionado
            link_pdf = row.get("LINK CND MUNICIPAL", "")
            if link_pdf and "drive.google.com" in str(link_pdf) and "/file/d/" in str(link_pdf):
                file_id = link_pdf.split("/file/d/")[1].split("/")[0]
                st.markdown(
                    f'<a href="https://drive.google.com/uc?export=download&id={file_id}" target="_blank">'
                    f'<button style="background-color:#1d3f77;color:white;padding:8px 16px;'
                    f'border:none;border-radius:4px;cursor:pointer;font-size:14px;">📥 Baixar PDF</button></a>',
                    unsafe_allow_html=True,
                )
        st.divider()
        row = st.session_state.pdf_selecionado
        link_pdf = row.get("LINK CND MUNICIPAL", "")
        st.subheader(f"📄 {row.get('Razão Social', '')}")
        st.caption(f"CNPJ: {row.get('CNPJ', '')}")
        if link_pdf and str(link_pdf).strip():
            try:
                if "drive.google.com" in str(link_pdf) and "/file/d/" in str(link_pdf):
                    file_id = link_pdf.split("/file/d/")[1].split("/")[0]
                    embed_url = f"https://drive.google.com/file/d/{file_id}/preview"
                    st.markdown(f'<iframe src="{embed_url}" width="100%" height="800" frameborder="0"></iframe>',
                                unsafe_allow_html=True)
                else:
                    st.markdown(f'<iframe src="{link_pdf}" width="100%" height="800" frameborder="0"></iframe>',
                                unsafe_allow_html=True)
            except Exception as e:
                st.error(f"❌ Erro ao carregar PDF: {e}")
        else:
            st.error("❌ PDF não disponível")
            st.info("Link do PDF não foi encontrado na planilha (coluna LINK CND MUNICIPAL)")
    else:
        st.info("💡 Selecione uma linha na tabela para visualizar o PDF correspondente.")
        df_grid = _sanitiza_df(df_cnd)
        grid_response = exibe_aggrid_com_oculta(
            df_grid, height=400, grid_key="grid_cnd_municipal",
            selection_mode="none",
            colunas_ocultas=["Situação", "LINK CND MUNICIPAL"],
        )

        output = BytesIO()
        df_cnd.to_excel(output, index=False)
        st.download_button("📥 Baixar Excel", data=output.getvalue(), file_name="cnd_municipal.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============================================================================
# DASHBOARD PARALEGAL
# ============================================================================

_UF_NOMES = {
    "AC": "Acre", "AL": "Alagoas", "AP": "Amapá", "AM": "Amazonas", "BA": "Bahia",
    "CE": "Ceará", "DF": "Distrito Federal", "ES": "Espírito Santo", "GO": "Goiás",
    "MA": "Maranhão", "MT": "Mato Grosso", "MS": "Mato Grosso do Sul",
    "MG": "Minas Gerais", "PA": "Pará", "PB": "Paraíba", "PR": "Paraná",
    "PE": "Pernambuco", "PI": "Piauí", "RJ": "Rio de Janeiro",
    "RN": "Rio Grande do Norte", "RS": "Rio Grande do Sul", "RO": "Rondônia",
    "RR": "Roraima", "SC": "Santa Catarina", "SP": "São Paulo", "SE": "Sergipe",
    "TO": "Tocantins",
}

_DASH_AZUL       = "#1d3f77"
_DASH_AZUL_CLARO = "#4a90d9"


def _nome_uf(uf):
    """GO → "Goiás (GO)". O nome por extenso também evita o Chrome traduzir a sigla."""
    uf = str(uf).strip().upper()
    if uf in _UF_NOMES:
        return f"{_UF_NOMES[uf]} ({uf})"
    return "Não informado" if uf in ("", "NAN", "NONE", "N/I") else uf


def _texto_local(serie):
    """Estado/Município limpos pra agrupar: vazio/nan viram "Não informado"."""
    s = serie.fillna("").astype(str).str.strip()
    return s.where(~s.str.upper().isin(["", "NAN", "NONE"]), "Não informado")


def _fmt_pct(v):
    return f"{v:.1f}%".replace(".", ",")


@st.dialog("Detalhes das Empresas", width="large")
def _modal_dashboard(titulo, df_show, colunas):
    st.markdown(f"**{titulo}** — {df_show.shape[0]} empresa(s)")
    cols_ok = [c for c in colunas if c in df_show.columns]
    df_exib = df_show[cols_ok].reset_index(drop=True).copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_formata_cnpj_mascara)
    df_exib = _sanitiza_df(df_exib)
    st.dataframe(df_exib, use_container_width=True, hide_index=True)


def _dash_card(icone, titulo, valor, detalhe=""):
    return (
        "<div class='dp-card'>"
        f"<div class='dp-card-top'><span class='dp-card-ico'>{icone}</span>{titulo}</div>"
        f"<div class='dp-card-val'>{valor}</div>"
        f"<div class='dp-card-det'>{detalhe}</div>"
        "</div>"
    )


def _dash_barras(df_count, rotulo_col, altura_linha=34, key=None):
    """Barras horizontais (maior em cima) com "qtd · %" na ponta. Devolve o evento
    de seleção do st.plotly_chart (clique na barra)."""
    import plotly.graph_objects as go

    df_plot = df_count.iloc[::-1]   # plotly desenha de baixo pra cima
    maximo = df_count["Quantidade"].max()
    cores = [_DASH_AZUL if i == 0 else _DASH_AZUL_CLARO for i in range(len(df_count))][::-1]

    fig = go.Figure(go.Bar(
        x=df_plot["Quantidade"], y=df_plot[rotulo_col], orientation="h",
        marker=dict(color=cores, cornerradius=6),
        text=[f"<b>{q}</b>  ·  {_fmt_pct(p)}" for q, p in zip(df_plot["Quantidade"], df_plot["Pct"])],
        # barra grande: número dentro (branco); barra pequena: fora (cinza)
        textposition="auto", cliponaxis=False, insidetextanchor="end",
        insidetextfont=dict(size=12, color="#ffffff"),
        outsidetextfont=dict(size=12, color="#34495e"),
        customdata=df_plot[["Chave"]].values,
        hovertemplate="<b>%{y}</b><br>%{x} empresa(s)<extra>Clique para ver a lista</extra>",
    ))
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        showlegend=False, bargap=0.35,
        xaxis=dict(visible=False, range=[0, maximo * 1.05]),
        yaxis=dict(title="", tickfont=dict(size=13, color="#2c3e50"), showgrid=False,
                   ticksuffix="  "),
        margin=dict(t=6, r=45, b=6, l=6),
        height=max(160, len(df_count) * altura_linha + 20),
        clickmode="event+select",
        dragmode=False,
    )
    return st.plotly_chart(fig, use_container_width=True, on_select="rerun", key=key,
                           config={"displayModeBar": False})


@st.fragment
def pagina_dashboard_paralegal():
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    if "Situação" not in df.columns:
        st.error("Coluna 'Situação' não encontrada.")
        return

    df_ativas = df[df["Situação"].astype(str).str.upper() == "ATIVA"].copy()
    total_ativas = df_ativas.shape[0]
    tem_estado = "Estado" in df_ativas.columns
    tem_mun = "Município" in df_ativas.columns

    df_ativas["_UF"] = _texto_local(df_ativas["Estado"]).str.upper() if tem_estado else "Não informado"
    df_ativas["_UF"] = df_ativas["_UF"].replace({"NÃO INFORMADO": "Não informado"})
    df_ativas["_MUN"] = _texto_local(df_ativas["Município"]) if tem_mun else "Não informado"

    st.markdown("""
<style>
.dp-titulo { color:#1d3f77 !important; margin:0 0 2px 0; font-weight:700; }
.dp-sub    { color:#7f8c8d; font-size:15px; margin:0 0 18px 0; }
.dp-cards  { display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr));
             gap:14px; margin-bottom:22px; }
.dp-card   { background:#fff; border:1px solid #e3e9f2; border-left:5px solid #1d3f77;
             border-radius:12px; padding:14px 18px;
             box-shadow:0 2px 8px rgba(29,63,119,0.07); }
.dp-card-top { color:#7f8c8d; font-size:13px; font-weight:600; text-transform:uppercase;
               letter-spacing:.4px; }
.dp-card-ico { margin-right:6px; }
.dp-card-val { color:#1d3f77; font-size:30px; font-weight:800; line-height:1.25; margin-top:4px;
               white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
.dp-card-det { color:#95a5a6; font-size:13px; min-height:18px; }
.dp-sec      { color:#1d3f77; font-size:19px; font-weight:700; margin:0; }
.dp-sec-cap  { color:#95a5a6; font-size:13px; margin:0 0 6px 0; }
</style>
""", unsafe_allow_html=True)

    st.markdown("<h2 class='dp-titulo'>Dashboard — Departamento Paralegal</h2>"
                "<p class='dp-sub'>Distribuição das empresas ativas por Estado e Município</p>",
                unsafe_allow_html=True)

    if total_ativas == 0:
        st.info("Nenhuma empresa ativa encontrada.")
        return

    # ── contagens ────────────────────────────────────────────────────────────
    df_uf = df_ativas["_UF"].value_counts().reset_index()
    df_uf.columns = ["Chave", "Quantidade"]
    df_uf["Pct"] = df_uf["Quantidade"] / total_ativas * 100
    df_uf["Estado"] = df_uf["Chave"].apply(_nome_uf)

    df_mun = df_ativas["_MUN"].value_counts().reset_index()
    df_mun.columns = ["Chave", "Quantidade"]
    df_mun["Pct"] = df_mun["Quantidade"] / total_ativas * 100
    df_mun["Município"] = df_mun["Chave"]

    qtd_uf  = int((df_uf["Chave"] != "Não informado").sum())
    qtd_mun = int((df_mun["Chave"] != "Não informado").sum())
    top_mun = df_mun.iloc[0]
    top_uf  = df_uf.iloc[0]

    # ── cards ────────────────────────────────────────────────────────────────
    st.markdown(
        "<div class='dp-cards'>"
        + _dash_card("🏢", "Empresas ativas", total_ativas, "base da aba GERAL")
        + _dash_card("🗺️", "Estados", qtd_uf,
                     f"{_nome_uf(top_uf['Chave'])} concentra {_fmt_pct(top_uf['Pct'])}")
        + _dash_card("📍", "Municípios", qtd_mun, "com ao menos 1 empresa ativa")
        + _dash_card("⭐", "Principal município", top_mun["Chave"].title(),
                     f"{int(top_mun['Quantidade'])} empresas · {_fmt_pct(top_mun['Pct'])} do total")
        + "</div>",
        unsafe_allow_html=True,
    )

    for k in ["ult_estado", "ult_municipio"]:
        if k not in st.session_state:
            st.session_state[k] = None
    modal_abrir = None   # apenas UM modal por execução
    colunas_modal = ["Código", "Razão Social", "CNPJ", "Município", "Estado", "Regime"]

    col_uf, col_mun = st.columns([1, 1.2], gap="large")

    # ── POR ESTADO ───────────────────────────────────────────────────────────
    with col_uf:
        with st.container(border=True):
            st.markdown("<p class='dp-sec'>🗺️ Empresas por Estado</p>"
                        "<p class='dp-sec-cap'>Clique em uma barra para ver as empresas</p>",
                        unsafe_allow_html=True)
            if tem_estado:
                ev_uf = _dash_barras(df_uf, "Estado", altura_linha=40, key="chart_estado")
                if ev_uf and ev_uf.selection and ev_uf.selection.points:
                    sel = (ev_uf.selection.points[0].get("customdata") or [None])[0]
                    if sel and sel != st.session_state["ult_estado"]:
                        st.session_state["ult_estado"] = sel
                        modal_abrir = (f"Estado: {_nome_uf(sel)}",
                                       df_ativas[df_ativas["_UF"] == sel], colunas_modal)
            else:
                st.warning("Coluna 'Estado' não encontrada.")

    # ── POR MUNICÍPIO ────────────────────────────────────────────────────────
    with col_mun:
        with st.container(border=True):
            st.markdown("<p class='dp-sec'>📍 Empresas por Município</p>"
                        "<p class='dp-sec-cap'>Clique em uma barra para ver as empresas</p>",
                        unsafe_allow_html=True)
            if tem_mun:
                LIMITE = 10
                ver_todos = False
                if len(df_mun) > LIMITE:
                    ver_todos = st.toggle(f"Mostrar todos os {len(df_mun)} municípios",
                                          key="dash_mun_todos")
                df_mun_exib = df_mun if ver_todos else df_mun.head(LIMITE)
                ev_mun = _dash_barras(df_mun_exib, "Município", altura_linha=34,
                                      key="chart_municipio_todos" if ver_todos else "chart_municipio")
                if not ver_todos and len(df_mun) > LIMITE:
                    resto = df_mun.iloc[LIMITE:]
                    st.caption(f"Outros {len(resto)} municípios somam {int(resto['Quantidade'].sum())} "
                               f"empresa(s) ({_fmt_pct(resto['Pct'].sum())}) — ative \"Mostrar todos\" para ver.")
                if modal_abrir is None and ev_mun and ev_mun.selection and ev_mun.selection.points:
                    sel = (ev_mun.selection.points[0].get("customdata") or [None])[0]
                    if sel and sel != st.session_state["ult_municipio"]:
                        st.session_state["ult_municipio"] = sel
                        modal_abrir = (f"Município: {sel}",
                                       df_ativas[df_ativas["_MUN"] == sel], colunas_modal)
            else:
                st.warning("Coluna 'Município' não encontrada.")

    # ── CONSULTA POR LOCALIDADE ──────────────────────────────────────────────
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    with st.container(border=True):
        st.markdown("<p class='dp-sec'>🔎 Consultar empresas por localidade</p>"
                    "<p class='dp-sec-cap'>Filtre por Estado e Município para ver e baixar a lista</p>",
                    unsafe_allow_html=True)

        f1, f2 = st.columns(2)
        ufs = df_uf["Chave"].tolist()
        uf_sel = f1.selectbox("Estado", ["Todos"] + ufs,
                              format_func=lambda v: v if v == "Todos" else _nome_uf(v),
                              key="dash_filtro_uf")
        base = df_ativas if uf_sel == "Todos" else df_ativas[df_ativas["_UF"] == uf_sel]
        muns = base["_MUN"].value_counts().index.tolist()
        mun_sel = f2.selectbox("Município", ["Todos"] + sorted(muns), key="dash_filtro_mun")
        if mun_sel != "Todos":
            base = base[base["_MUN"] == mun_sel]

        # resumo Estado × Município da seleção
        resumo = (base.groupby(["_UF", "_MUN"]).size().reset_index(name="Empresas")
                  .sort_values(["Empresas", "_MUN"], ascending=[False, True]))
        resumo["Estado"] = resumo["_UF"].apply(_nome_uf)
        resumo["Município"] = resumo["_MUN"]
        resumo["% do total"] = (resumo["Empresas"] / total_ativas * 100).apply(_fmt_pct)

        cols_lista = [c for c in colunas_modal if c in base.columns]
        lista = base[cols_lista].copy()
        if "CNPJ" in lista.columns:
            lista["CNPJ"] = lista["CNPJ"].apply(_formata_cnpj_mascara)
        if "Razão Social" in lista.columns:
            lista = lista.sort_values("Razão Social", key=lambda c: c.astype(str).str.upper())
        lista = _sanitiza_df(lista.reset_index(drop=True))

        st.markdown(f"**{len(lista)} empresa(s)** em **{base['_MUN'].nunique()} município(s)**")
        aba_lista, aba_resumo = st.tabs(["📋 Empresas", "📊 Resumo por município"])
        with aba_lista:
            st.dataframe(lista, use_container_width=True, hide_index=True,
                         height=min(420, 38 + 35 * max(len(lista), 1)))
        with aba_resumo:
            st.dataframe(resumo[["Estado", "Município", "Empresas", "% do total"]],
                         use_container_width=True, hide_index=True,
                         height=min(420, 38 + 35 * max(len(resumo), 1)),
                         column_config={"Empresas": st.column_config.ProgressColumn(
                             "Empresas", format="%d", min_value=0,
                             max_value=int(resumo["Empresas"].max()) if len(resumo) else 1)})

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as wr:
            lista.to_excel(wr, index=False, sheet_name="Empresas")
            resumo[["Estado", "Município", "Empresas", "% do total"]].to_excel(
                wr, index=False, sheet_name="Resumo")
        st.download_button("📥 Baixar Excel", data=output.getvalue(),
                           file_name="empresas_por_localidade.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="dash_baixar_excel")

    # ── abre o modal (apenas um por execução) ────────────────────────────────
    if modal_abrir:
        _modal_dashboard(*modal_abrir)


COLUNAS_XML = [
    "Número da Nota", "Data de Emissão", "Situação",
    "Prestador Razão Social", "Prestador CNPJ/CPF",
    "Tomador Razão Social", "Tomador CNPJ/CPF",
    "Valor Serviço", "Base de Cálculo", "Valor ISS",
    "PIS", "COFINS", "CSLL", "IRRF", "INSS", "ISS Retido",
    "Federais Retidos", "Tipo Retenção Federal",
    "CNAE", "Código LC", "Descrição LC", "Observações",
    "IBS", "CBS",
]

COFINS_LABEL = "*COFINS"
IBS_LABEL    = "*IBS"
COLS_IMPOSTO = ["PIS", COFINS_LABEL, "CSLL", "IRRF", "INSS", "ISS Retido", IBS_LABEL, "CBS"]
COLS_CODIGO  = ["CNAE", "Código LC", "Descrição LC"]


import re
from datetime import datetime as _dt

def _corrige_codigo_lc(val):
    """Converte qualquer formato de data/número para 00.00"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    # Timestamp ou datetime
    if isinstance(val, (pd.Timestamp, _dt)):
        return f"{val.day:02d}.{val.month:02d}"
    # String
    s = str(val).strip()
    if not s or s.upper() in ("NAN", "NAT", "NONE", ""):
        return ""
    # DD/MM/AAAA
    m = re.match(r'^(\d{1,2})/(\d{1,2})/\d{2,4}$', s)
    if m:
        return f"{int(m.group(1)):02d}.{int(m.group(2)):02d}"
    # já está no formato DD.MM
    m2 = re.match(r'^(\d{1,2})\.(\d{1,2})$', s)
    if m2:
        return f"{int(m2.group(1)):02d}.{int(m2.group(2)):02d}"
    # número decimal ex: 14.1  →  14.01
    m3 = re.match(r'^(\d{1,2})\.(\d{1,2})(\d*)$', s)
    if m3:
        return f"{int(m3.group(1)):02d}.{int(m3.group(2)):02d}"
    return s


def _limpa_numero(val):
    """Converte para float aceitando vírgula decimal e R$."""
    if pd.isna(val):
        return 0.0
    s = str(val).strip().replace("R$", "").replace(" ", "")
    # se tiver vírgula como decimal: 1.234,56 → 1234.56
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


@st.cache_data(ttl=600)
def _carrega_xml_cache(sheet_name):
    try:
        resp = requests.get(GOOGLE_SHEET_URL)
        resp.raise_for_status()
        df = pd.read_excel(
            BytesIO(resp.content),
            sheet_name=sheet_name,
            engine="openpyxl",
            header=0,
        )
    except Exception as e:
        return None, str(e)
    df.columns = df.columns.str.strip()
    n_cols = min(len(df.columns), len(COLUNAS_XML))
    df = df.iloc[:, :n_cols].copy()
    df.columns = COLUNAS_XML[:n_cols]
    return df, None


def _carrega_xml(sheet_name, col_cnpj_filtro):
    df_geral = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df_geral is None:
        return None, None, None

    df_ativos = df_geral[df_geral["Situação"].astype(str).str.upper() == "ATIVA"].copy()
    df_ativos["_cnpj_norm"] = df_ativos["CNPJ"].apply(_normaliza_cnpj)
    cnpjs_ativos = set(df_ativos["_cnpj_norm"])

    # mapa CNPJ → Razão Social da planilha GERAL
    mapa_razao = dict(zip(df_ativos["_cnpj_norm"], df_ativos["Razão Social"]))

    df, erro = _carrega_xml_cache(sheet_name)
    if df is None:
        st.error(f"Erro ao ler aba '{sheet_name}': {erro}")
        return None, None, None

    df = df.copy()

    # ── renomeia COFINS e IBS para evitar tradução do Chrome ─────────────────
    if "COFINS" in df.columns:
        df = df.rename(columns={"COFINS": COFINS_LABEL})
    if "IBS" in df.columns:
        df = df.rename(columns={"IBS": IBS_LABEL})

    # ── Código LC ─────────────────────────────────────────────────────────────
    if "Código LC" in df.columns:
        df["Código LC"] = df["Código LC"].apply(_corrige_codigo_lc)

    # ── normaliza CNPJ e filtra ativos ────────────────────────────────────────
    if col_cnpj_filtro in df.columns:
        df[col_cnpj_filtro] = df[col_cnpj_filtro].apply(_normaliza_cnpj)
        # para DMS o Prestador pode ter CPF (11 dígitos) — compara só os CNPJs (14 dígitos)
        mask = df[col_cnpj_filtro].isin(cnpjs_ativos)
        # se não encontrar nada com 14 dígitos, tenta com os dígitos que tiver
        if mask.sum() == 0:
            cnpjs_flexivel = set(c.lstrip("0") for c in cnpjs_ativos)
            mask = df[col_cnpj_filtro].apply(
                lambda x: x.lstrip("0") in cnpjs_flexivel
            )
        df = df[mask].copy()

    # ── força string em colunas de CNPJ/CPF ──────────────────────────────────
    for col in ["Prestador CNPJ/CPF", "Tomador CNPJ/CPF"]:
        if col in df.columns:
            df[col] = df[col].astype(str)        

    # ── converte colunas monetárias para float ────────────────────────────────
    COLS_MONETARIAS = [
        "Valor Serviço", "Base de Cálculo", "Valor ISS",
        "PIS", COFINS_LABEL, "CSLL", "IRRF", "INSS",
        "Federais Retidos", IBS_LABEL, "CBS",
    ]
    for col in COLS_MONETARIAS:
        if col in df.columns:
            df[col] = df[col].apply(_limpa_numero)

    # ISS Retido é texto — mantém como string limpa
    if "ISS Retido" in df.columns:
        df["ISS Retido"] = df["ISS Retido"].fillna("").astype(str).str.strip()

    # ── formata data ──────────────────────────────────────────────────────────
    if "Data de Emissão" in df.columns:
        df["Data de Emissão"] = pd.to_datetime(
            df["Data de Emissão"], errors="coerce"
        ).dt.strftime("%d/%m/%Y").fillna("")

    return df, cnpjs_ativos, mapa_razao


def _filtros_xml(df, col_cnpj, mapa_razao, formata_cnpj=True, page_id="rest"):
    """
    col_cnpj     : coluna de CNPJ a filtrar (Tomador ou Prestador)
    mapa_razao   : dict CNPJ_14digitos → Razão Social da planilha GERAL
    formata_cnpj : True para Tomador (14 dígitos), False para Prestador (CPF misturado)
    page_id      : "dms" ou "rest" — garante keys únicas e comportamento distinto
    """
    st.markdown("""
    <div style='background:#f4f6fa; border-radius:10px; padding:16px 20px 8px 20px;
                border:1px solid #dce3f0; margin-bottom:16px;'>
    <span style='font-size:15px; font-weight:700; color:#1d3f77;'>Filtros</span>
    </div>
    """, unsafe_allow_html=True)

    # ── empresa + situação ────────────────────────────────────────────────────
    if col_cnpj in df.columns:
        cnpjs_unicos = sorted(df[col_cnpj].dropna().astype(str).unique().tolist())
        def _label(cnpj):
            razao = mapa_razao.get(cnpj, "")
            cnpj_fmt = cnpj.zfill(14) if formata_cnpj else cnpj
            return f"{razao} — {cnpj_fmt}" if razao else cnpj_fmt
        opcoes_emp = ["Todas as empresas"] + [_label(c) for c in cnpjs_unicos]
        mapa_label_cnpj = {"Todas as empresas": None}
        for c in cnpjs_unicos:
            mapa_label_cnpj[_label(c)] = c
    else:
        opcoes_emp = ["Todas as empresas"]
        mapa_label_cnpj = {"Todas as empresas": None}

    fc1, fc2 = st.columns([3, 1])
    with fc1:
        sel_emp = st.selectbox(
            "Empresa",
            options=opcoes_emp,
            key=f"filtro_emp_{page_id}",
        )
    with fc2:
        sel_sit = st.selectbox(
            "Situação da nota",
            options=["Todas", "Autorizada", "Cancelada"],
            key=f"filtro_sit_{page_id}",
        )

    # ── impostos ──────────────────────────────────────────────────────────────
    st.markdown(
        "<p style='font-size:13px; color:#555; margin:8px 0 4px;'>"
        "<b>Exibir apenas notas com valor &gt; 0 em:</b></p>",
        unsafe_allow_html=True,
    )
    fi_cols = st.columns(len(COLS_IMPOSTO))
    filtros_imposto = {}
    for i, col in enumerate(COLS_IMPOSTO):
        with fi_cols[i]:
            if col == "ISS Retido":
                label_exib = "ISS próprio" if page_id == "dms" else "ISS Retido"
            else:
                label_exib = col if col.strip() else "—"
            tem = col in df.columns
            filtros_imposto[col] = st.checkbox(
                label_exib,
                key=f"imp_{col}_{page_id}",
                disabled=not tem,
            )

    # ── códigos ───────────────────────────────────────────────────────────────
    st.markdown(
        "<p style='font-size:13px; color:#555; margin:8px 0 4px;'>"
        "<b>Filtrar por código:</b></p>",
        unsafe_allow_html=True,
    )
    fk1, fk2, fk3 = st.columns(3)
    filtros_codigo = {}
    for col, fc in zip(COLS_CODIGO, [fk1, fk2, fk3]):
        with fc:
            opcoes = sorted(
                df[col].dropna().astype(str)
                .replace("", pd.NA).dropna().unique().tolist()
            ) if col in df.columns else []
            filtros_codigo[col] = st.multiselect(
                col if col.strip() else "—",
                options=opcoes,
                placeholder="Todos...",
                key=f"cod_{col}_{page_id}",
            )

    # ── aplica filtros ────────────────────────────────────────────────────────
    df_f = df.copy()

    cnpj_sel = mapa_label_cnpj.get(sel_emp)
    if cnpj_sel and col_cnpj in df_f.columns:
        df_f = df_f[df_f[col_cnpj].astype(str) == cnpj_sel]

    if sel_sit != "Todas" and "Situação" in df_f.columns:
        df_f = df_f[df_f["Situação"].astype(str).str.strip().str.upper()
                    == sel_sit.upper()]

    for col, ativo in filtros_imposto.items():
        if not ativo or col not in df_f.columns:
            continue
        if col == "ISS Retido":
            if page_id == "dms":
                df_f = df_f[df_f[col].astype(str).str.strip().str.upper() == "NÃO"]
            else:
                df_f = df_f[df_f[col].astype(str).str.strip().str.upper() == "SIM"]
        else:
            df_f = df_f[pd.to_numeric(df_f[col], errors="coerce").fillna(0) > 0]

    for col, selecionados in filtros_codigo.items():
        if selecionados and col in df_f.columns:
            df_f = df_f[df_f[col].astype(str).isin(selecionados)]

    return df_f


COLS_TOTALIZADOR = [
    "Valor Serviço", "Base de Cálculo", "Valor ISS",
    "PIS", COFINS_LABEL, "CSLL", "IRRF", "INSS", IBS_LABEL, "CBS",
]

CORES_TOTAL = [
    "#1d3f77", "#2471a3", "#148f77",
    "#1e8449", "#b7950b", "#784212", "#922b21", "#6c3483",
    "#117a65", "#1a5276",
]

def _exibe_totalizador(df):
    colunas_presentes = [c for c in COLS_TOTALIZADOR if c in df.columns]
    if not colunas_presentes:
        return

    st.markdown(
        "<p style='font-size:13px; font-weight:700; color:#1d3f77; "
        "margin:12px 0 6px;'>Totais do filtro atual:</p>",
        unsafe_allow_html=True,
    )

    cols_ui = st.columns(len(colunas_presentes))
    for i, col in enumerate(colunas_presentes):
        cor = CORES_TOTAL[COLS_TOTALIZADOR.index(col)]
        nome_exib = "COFINS" if col == COFINS_LABEL else col

        total = pd.to_numeric(df[col], errors="coerce").fillna(0).sum()
        valor_fmt = (
            f"R$ {total:,.2f}"
            .replace(",", "X").replace(".", ",").replace("X", ".")
        )

        with cols_ui[i]:
            st.markdown(
                f"<div translate='no' style='text-align:center; padding:8px 4px; "
                f"background:{cor}; border-radius:8px;'>"
                f"<span style='font-size:11px; color:rgba(255,255,255,0.8);'>"
                f"{nome_exib}</span><br>"
                f"<span style='font-size:14px; font-weight:700; color:#ffffff;'>"
                f"{valor_fmt}</span></div>",
                unsafe_allow_html=True,
            )

def _exibe_grid_xml(df, grid_key):
    import hashlib
    hash_key = hashlib.md5(str(df.shape).encode() + str(df.index.tolist()).encode()).hexdigest()[:8]

    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(
        resizable=True, filter=True, sortable=True,
        minWidth=120, width=150,
    )

    colunas_fixas = ["Número da Nota", "Data de Emissão",
                     "Prestador Razão Social", "Prestador CNPJ/CPF",
                     "Tomador Razão Social", "Tomador CNPJ/CPF"]
    for col in colunas_fixas[:4]:
        if col in df.columns:
            gb.configure_column(col, pinned="left", width=160)

    gb.configure_grid_options(
        domLayout="normal",
        suppressHorizontalScroll=False,
        enableRangeSelection=True,
        suppressColumnVirtualisation=True,
        alwaysShowHorizontalScroll=True, 
    )

    AgGrid(
        df,
        gridOptions=gb.build(),
        height=500,
        key=f"{grid_key}_{hash_key}",
        fit_columns_on_grid_load=False,
        enable_enterprise_modules=False,
        update_mode=GridUpdateMode.NO_UPDATE,
        allow_unsafe_jscode=True,
    )


def pagina_leitura_xml_dms():
    st.empty()
    st.markdown("<h2>LEITURA XML DMS</h2>", unsafe_allow_html=True)

    col_cnpj = "Prestador CNPJ/CPF"

    df, _, mapa_razao = _carrega_xml(SHEET_XML_DMS, col_cnpj)
    if df is None:
        return


    total_empresas = df[col_cnpj].nunique() if col_cnpj in df.columns else 0
    total_notas    = len(df)
    st.markdown(
        f"<p style='font-size:18px;'>"
        f"<b>Empresas ativas com notas:</b> {total_empresas} &nbsp;|&nbsp; "
        f"<b>Total de notas:</b> {total_notas}</p>",
        unsafe_allow_html=True,
    )

    df_filtrado = _filtros_xml(df, col_cnpj, mapa_razao, formata_cnpj=False, page_id="dms")


    st.markdown(
        f"<p style='font-size:14px; color:#555;'>Exibindo <b>{len(df_filtrado)}</b> "
        f"nota(s) de <b>{df_filtrado[col_cnpj].nunique()}</b> empresa(s)</p>",
        unsafe_allow_html=True,
    )

    _exibe_totalizador(df_filtrado)

    df_filtrado = _sanitiza_df(df_filtrado)
    _exibe_grid_xml(df_filtrado, "grid_xml_dms")

    output = BytesIO()
    df_filtrado.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(),
                       file_name="leitura_xml_dms.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def pagina_leitura_xml_rest():
    st.empty()
    st.markdown("<h2>LEITURA XML REST</h2>", unsafe_allow_html=True)

    col_cnpj = "Tomador CNPJ/CPF"

    df, _, mapa_razao = _carrega_xml(SHEET_XML_REST, col_cnpj)
    if df is None:
        return

    total_empresas = df[col_cnpj].nunique() if col_cnpj in df.columns else 0
    total_notas    = len(df)
    st.markdown(
        f"<p style='font-size:18px;'>"
        f"<b>Empresas ativas com notas:</b> {total_empresas} &nbsp;|&nbsp; "
        f"<b>Total de notas:</b> {total_notas}</p>",
        unsafe_allow_html=True,
    )

    df_filtrado = _filtros_xml(df, col_cnpj, mapa_razao, formata_cnpj=True, page_id="rest")

    st.markdown(
        f"<p style='font-size:14px; color:#555;'>Exibindo <b>{len(df_filtrado)}</b> "
        f"nota(s) de <b>{df_filtrado[col_cnpj].nunique()}</b> empresa(s)</p>",
        unsafe_allow_html=True,
    )

    _exibe_totalizador(df_filtrado)

    df_filtrado = _sanitiza_df(df_filtrado)
    _exibe_grid_xml(df_filtrado, "grid_xml_rest")

    output = BytesIO()
    df_filtrado.to_excel(output, index=False)
    st.download_button("Baixar Excel", data=output.getvalue(),
                       file_name="leitura_xml_rest.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

def _sanitiza_df(df):
    """Converte todas as colunas para string — evita erro Arrow.
    Números inteiros (mesmo vindos como float, ex.: 3.0) não mostram
    o '.0' no final; valores realmente decimais (ex.: 3.5) são mantidos."""
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            def _fmt_num(v):
                if pd.isna(v):
                    return ""
                return str(int(v)) if float(v).is_integer() else str(v)
            df[col] = df[col].apply(_fmt_num)
        else:
            df[col] = df[col].astype(str).replace("nan", "").replace("None", "")
    return df

@st.dialog("Prefeitura — DMS Sem Acesso")
def _modal_sem_acesso_dms(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s)**")
    st.dataframe(df_show.reset_index(drop=True), use_container_width=True, hide_index=True)

@st.dialog("SEFAZ — Sem Acesso")
def _modal_sem_acesso_sefaz(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s)**")
    st.dataframe(df_show.reset_index(drop=True), use_container_width=True, hide_index=True)

@st.dialog("eCAC — Sem Procuração")
def _modal_sem_acesso_ecac(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s)**")
    st.dataframe(df_show.reset_index(drop=True), use_container_width=True, hide_index=True)


@st.fragment
def pagina_sem_acesso():
    import plotly.graph_objects as go
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    df_ativas = df[df["Situação"].astype(str).str.upper() == "ATIVA"].copy()
    if df_ativas.empty:
        st.warning("Nenhuma empresa ATIVA encontrada.")
        return

    # ── colunas base para exibição ────────────────────────────────────────────
    COLS_BASE = ["Código", "Razão Social", "CNPJ"]

    def _prepara(df_fil):
        cols = [c for c in COLS_BASE if c in df_fil.columns]
        d = df_fil[cols].copy()
        if "CNPJ" in d.columns:
            d["CNPJ"] = d["CNPJ"].apply(_formata_cnpj_mascara)
        return d.reset_index(drop=True)

    # ── PREFEITURA — DMS ──────────────────────────────────────────────────────
    col_dms = "DMS"
    if col_dms in df_ativas.columns:
        mask_dms = df_ativas[col_dms].astype(str).str.upper().str.contains("SEM ACESSO", na=False)
        df_dms_sa = _prepara(df_ativas[mask_dms])
    else:
        df_dms_sa = pd.DataFrame(columns=COLS_BASE)

    # ── SEFAZ ─────────────────────────────────────────────────────────────────
    col_sefaz = "IMPORTAÇÃO"
    if col_sefaz in df_ativas.columns:
        mask_sefaz = df_ativas[col_sefaz].astype(str).str.upper().str.contains("SEM ACESSO", na=False)
        df_sefaz_sa = _prepara(df_ativas[mask_sefaz])
    else:
        df_sefaz_sa = pd.DataFrame(columns=COLS_BASE)

    # ── eCAC — SIMPLES, REINF, DCTF WEB ──────────────────────────────────────
    ecac_masks = []
    for col in ["MOTIVO SITUAÇÃO DO DAS", "MOTIVO SITUAÇÃO REINF", "MOTIVO SITUAÇÃO DCTF WEB"]:
        if col in df_ativas.columns:
            ecac_masks.append(
                df_ativas[col].astype(str).str.upper().str.contains("PROCURA", na=False)
            )

    if ecac_masks:
        mask_ecac = ecac_masks[0]
        for m in ecac_masks[1:]:
            mask_ecac = mask_ecac | m
        df_ecac_sa = _prepara(df_ativas[mask_ecac])
    else:
        df_ecac_sa = pd.DataFrame(columns=COLS_BASE)

    # ── cabeçalho ─────────────────────────────────────────────────────────────
    st.markdown("<h2>SEM ACESSO</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:18px;'>"
        f"<b>Prefeitura:</b> {len(df_dms_sa)} &nbsp;|&nbsp; "
        f"<b>SEFAZ:</b> {len(df_sefaz_sa)} &nbsp;|&nbsp; "
        f"<b>eCAC:</b> {len(df_ecac_sa)}</p>",
        unsafe_allow_html=True,
    )

    # ── session keys ──────────────────────────────────────────────────────────
    for k in ["sa_chart_key"]:
        if k not in st.session_state:
            st.session_state[k] = 0

    # ── três donuts lado a lado ───────────────────────────────────────────────
    total_geral = len(df_ativas)
    col1, col2, col3 = st.columns(3)

    def _donut(titulo, qtd_sa, total, cor_sa, cor_ok):
        qtd_ok = max(total - qtd_sa, 0)
        fig = go.Figure(data=[go.Pie(
            labels=["Sem Acesso", "Com Acesso"],
            values=[int(qtd_sa), int(qtd_ok)],
            hole=0.68,
            marker=dict(colors=[cor_sa, cor_ok], line=dict(color="#ffffff", width=3)),
            textinfo="none",
            hovertemplate="<b>%{label}</b><br>%{value} empresa(s)<extra></extra>",
            direction="clockwise",
            sort=False,
        )])
        fig.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            showlegend=False,
            margin=dict(t=10, b=10, l=10, r=10),
            height=220,
            annotations=[dict(
                text=f"<b>{qtd_sa}</b><br><span style='font-size:10px'>sem acesso</span>",
                x=0.5, y=0.5, xanchor="center", yanchor="middle",
                showarrow=False,
                font=dict(size=18, color="#1d3f77"),
            )],
        )
        return fig

    with col1:
        st.markdown("<h4 style='text-align:center; color:#1d3f77;'>Prefeitura</h4>",
                    unsafe_allow_html=True)
        st.plotly_chart(_donut("Prefeitura", len(df_dms_sa), total_geral,
                               "#c0392b", "#bdc3c7"),
                        use_container_width=True,
                        key=f"chart_sa_dms_{st.session_state['sa_chart_key']}")
        pct = round(len(df_dms_sa) / total_geral * 100) if total_geral else 0
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdedec; "
            f"border-radius:8px; border-left:4px solid #c0392b;'>"
            f"<span style='font-size:20px; font-weight:700; color:#c0392b;'>{len(df_dms_sa)}</span><br>"
            f"<span style='font-size:12px; color:#555;'>DMS Sem Acesso ({pct}%)</span></div>",
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Ver empresas", key="btn_sa_dms", use_container_width=True):
            _modal_sem_acesso_dms(df_dms_sa)

    with col2:
        st.markdown("<h4 style='text-align:center; color:#1d3f77;'>SEFAZ</h4>",
                    unsafe_allow_html=True)
        st.plotly_chart(_donut("SEFAZ", len(df_sefaz_sa), total_geral,
                               "#e67e22", "#bdc3c7"),
                        use_container_width=True,
                        key=f"chart_sa_sefaz_{st.session_state['sa_chart_key']}")
        pct = round(len(df_sefaz_sa) / total_geral * 100) if total_geral else 0
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdf3e7; "
            f"border-radius:8px; border-left:4px solid #e67e22;'>"
            f"<span style='font-size:20px; font-weight:700; color:#e67e22;'>{len(df_sefaz_sa)}</span><br>"
            f"<span style='font-size:12px; color:#555;'>SEFAZ Sem Acesso ({pct}%)</span></div>",
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Ver empresas", key="btn_sa_sefaz", use_container_width=True):
            _modal_sem_acesso_sefaz(df_sefaz_sa)

    with col3:
        st.markdown("<h4 style='text-align:center; color:#1d3f77;'>eCAC</h4>",
                    unsafe_allow_html=True)
        st.plotly_chart(_donut("eCAC", len(df_ecac_sa), total_geral,
                               "#8e44ad", "#bdc3c7"),
                        use_container_width=True,
                        key=f"chart_sa_ecac_{st.session_state['sa_chart_key']}")
        pct = round(len(df_ecac_sa) / total_geral * 100) if total_geral else 0
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#f5eef8; "
            f"border-radius:8px; border-left:4px solid #8e44ad;'>"
            f"<span style='font-size:20px; font-weight:700; color:#8e44ad;'>{len(df_ecac_sa)}</span><br>"
            f"<span style='font-size:12px; color:#555;'>eCAC Sem Procuração ({pct}%)</span></div>",
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Ver empresas", key="btn_sa_ecac", use_container_width=True):
            _modal_sem_acesso_ecac(df_ecac_sa)

    st.divider()

    # ── lista geral ───────────────────────────────────────────────────────────
    st.markdown("### Lista Geral — Todas as pendências", unsafe_allow_html=True)

    df_dms_sa["Origem"] = "DMS — Prefeitura"
    df_sefaz_sa["Origem"] = "SEFAZ"
    df_ecac_sa["Origem"] = "eCAC"

    df_geral_sa = pd.concat([df_dms_sa, df_sefaz_sa, df_ecac_sa], ignore_index=True)

    if not df_geral_sa.empty:
        df_geral_sa = _sanitiza_df(df_geral_sa)
        exibe_aggrid(df_geral_sa, height=400, grid_key="grid_sem_acesso")

        output = BytesIO()
        df_geral_sa.to_excel(output, index=False)
        st.download_button(
            "Baixar Excel", data=output.getvalue(),
            file_name="sem_acesso.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.success("Nenhuma pendência encontrada!")


@st.dialog("SEFAZ ALTERAÇÃO QUANTIDADE NOTAS — Alteradas")
def _modal_sefaz_alteracao_quantidade(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) com quantidade alterada**")
    st.dataframe(df_show.reset_index(drop=True), use_container_width=True, hide_index=True)


@st.fragment
def pagina_sefaz_alteracao_quantidade_notas():
    import plotly.graph_objects as go
    st.empty()

    # ── carrega abas SEFAZ INICIAL e SEFAZ FINAL ──────────────────────────────
    try:
        resp = requests.get(GOOGLE_SHEET_URL)
        resp.raise_for_status()
        conteudo = BytesIO(resp.content)
        df_inicial = pd.read_excel(
            conteudo, sheet_name=SHEET_SEFAZ_INICIAL, engine="openpyxl", header=0,
        )
        conteudo.seek(0)
        df_final = pd.read_excel(
            conteudo, sheet_name=SHEET_SEFAZ_FINAL, engine="openpyxl", header=0,
        )
    except Exception as e:
        st.error(f"Erro ao ler abas SEFAZ INICIAL / SEFAZ FINAL: {e}")
        return

    df_inicial.columns = df_inicial.columns.str.strip()
    df_final.columns   = df_final.columns.str.strip()

    st.markdown("<h2>SEFAZ ALTERAÇÃO QUANTIDADE NOTAS</h2>", unsafe_allow_html=True)

    # ── identifica colunas por posição — A=0 (CÓDIGO), E=4 (QUANTIDADE) ───────
    cols_inicial = df_inicial.columns.tolist()
    cols_final   = df_final.columns.tolist()

    def _col(cols, idx):
        return cols[idx] if idx < len(cols) else None

    nome_cod_ini  = _col(cols_inicial, 0)
    nome_qtd_ini  = _col(cols_inicial, 4)
    nome_cod_fim  = _col(cols_final, 0)
    nome_qtd_fim  = _col(cols_final, 4)

    if not all([nome_cod_ini, nome_qtd_ini, nome_cod_fim, nome_qtd_fim]):
        st.error("Colunas A ou E não encontradas nas abas SEFAZ INICIAL / SEFAZ FINAL.")
        return

    # ── SEFAZ FINAL sem dados → NÃO DISPONÍVEL ────────────────────────────────
    if df_final[nome_qtd_fim].notna().sum() == 0:
        st.markdown(
            "<div style='text-align:center; padding:60px 20px; background:#fdf2e3; "
            "border-radius:12px; border-left:6px solid #e67e22;'>"
            "<span style='font-size:28px; font-weight:700; color:#e67e22;'>NÃO DISPONÍVEL</span><br>"
            "<span style='font-size:14px; color:#555;'>A aba SEFAZ FINAL ainda não possui dados.</span>"
            "</div>",
            unsafe_allow_html=True,
        )
        return

    # ── carrega aba GERAL para pegar Nome e CNPJ ──────────────────────────────
    df_geral = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df_geral is None:
        return

    df_geral = df_geral.copy()
    if "CNPJ" in df_geral.columns:
        df_geral["CNPJ_fmt"] = df_geral["CNPJ"].apply(_formata_cnpj_mascara)
    if "Código" in df_geral.columns:
        df_geral["Código"] = df_geral["Código"].astype(str).str.strip()

    mapa_empresa = {}
    for _, row in df_geral.iterrows():
        cod = str(row.get("Código", "")).strip()
        if cod:
            mapa_empresa[cod] = {
                "Razão Social": row.get("Razão Social", ""),
                "CNPJ": row.get("CNPJ_fmt", ""),
                "Insc. Estadual":  row.get("Insc. Estadual", ""),
            }

    # ── normaliza códigos e quantidades ───────────────────────────────────────
    def _limpa_codigo(val):
        s = str(val).strip()
        if s.endswith(".0"):
            s = s[:-2]
        return s.upper().replace("NAN", "").strip()

    df_inicial[nome_cod_ini] = df_inicial[nome_cod_ini].apply(_limpa_codigo)
    df_final[nome_cod_fim]   = df_final[nome_cod_fim].apply(_limpa_codigo)
    df_inicial[nome_qtd_ini] = df_inicial[nome_qtd_ini].apply(_limpa_numero)
    df_final[nome_qtd_fim]   = df_final[nome_qtd_fim].apply(_limpa_numero)

    # ── filtra linhas válidas de cada lado ────────────────────────────────────
    df_lado_ini = df_inicial[df_inicial[nome_cod_ini] != ""][
        [nome_cod_ini, nome_qtd_ini]
    ].copy()
    df_lado_ini.columns = ["Código", "Qtd_Inicial"]

    df_lado_fim = df_final[df_final[nome_cod_fim] != ""][
        [nome_cod_fim, nome_qtd_fim]
    ].copy()
    df_lado_fim.columns = ["Código", "Qtd_Final"]

    # ── agrupa por código (soma caso haja duplicatas) ─────────────────────────
    df_lado_ini = df_lado_ini.groupby("Código", as_index=False)["Qtd_Inicial"].sum()
    df_lado_fim = df_lado_fim.groupby("Código", as_index=False)["Qtd_Final"].sum()

    # ── junta pelos códigos ───────────────────────────────────────────────────
    df_merge = pd.merge(df_lado_ini, df_lado_fim, on="Código", how="outer").fillna(0)

    # ── compara ───────────────────────────────────────────────────────────────
    df_merge["Diferença"] = df_merge["Qtd_Final"] - df_merge["Qtd_Inicial"]
    df_merge["Status"] = df_merge["Diferença"].apply(
        lambda d: "MESMA QUANTIDADE" if d == 0 else "QUANTIDADE ALTERADA"
    )
    df_alterada = df_merge[df_merge["Diferença"] != 0].copy()
    df_mesma    = df_merge[df_merge["Diferença"] == 0].copy()

    # ── monta df de exibição com Nome e CNPJ da GERAL ─────────────────────────
    def _enriquece(df_in):
        rows = []
        for _, row in df_in.iterrows():
            cod = str(row["Código"]).strip()
            emp = mapa_empresa.get(cod, {})

            def _fmt_valor(v):
                try:
                    f = float(v)
                    if f == int(f):
                        return int(f)
                    return round(f, 2)
                except:
                    return v

            rows.append({
                "Código":          cod,
                "Razão Social":    emp.get("Razão Social", ""),
                "CNPJ":            emp.get("CNPJ", ""),
                "Insc. Estadual":  emp.get("Insc. Estadual", ""),
                "Quantidade Inicial": _fmt_valor(row["Qtd_Inicial"]),
                "Quantidade Final":   _fmt_valor(row["Qtd_Final"]),
                "Diferença":          _fmt_valor(row["Diferença"]),
                "Status":             row["Status"],
            })
        return pd.DataFrame(rows)

    df_result = _enriquece(df_alterada)

    total_empresas = len(df_merge)
    total_alterada = len(df_result)
    total_mesma    = len(df_mesma)

    # ── cabeçalho ─────────────────────────────────────────────────────────────
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Quantidade alterada:</b> {total_alterada} &nbsp;|&nbsp; "
        f"<b>Mesma quantidade:</b> {total_mesma} &nbsp;|&nbsp; "
        f"<b>Total:</b> {total_empresas}</p>",
        unsafe_allow_html=True,
    )

    # ── donut ─────────────────────────────────────────────────────────────────
    if "sefaz_altqtd_key" not in st.session_state:
        st.session_state["sefaz_altqtd_key"] = 0

    pct_alterada = round(total_alterada / total_empresas * 100) if total_empresas else 0
    pct_mesma    = round(total_mesma    / total_empresas * 100) if total_empresas else 0

    fig = go.Figure(data=[go.Pie(
        labels=["Quantidade Alterada", "Mesma Quantidade"],
        values=[int(total_alterada), int(total_mesma)],
        hole=0.68,
        marker=dict(
            colors=["#c0392b", "#27ae60"],
            line=dict(color="#ffffff", width=3),
        ),
        textinfo="none",
        hovertemplate="<b>%{label}</b><br>%{value} empresa(s) — %{percent}<extra></extra>",
        direction="clockwise",
        sort=False,
    )])
    fig.update_layout(
        paper_bgcolor="white", plot_bgcolor="white",
        showlegend=False,
        margin=dict(t=20, b=20, l=20, r=20),
        height=300,
        annotations=[dict(
            text=f"<b>{total_empresas}</b><br><span style='font-size:11px'>empresas</span>",
            x=0.5, y=0.5,
            xanchor="center", yanchor="middle",
            showarrow=False,
            font=dict(size=22, color="#1d3f77"),
        )],
    )

    col_esq, col_centro, col_dir = st.columns([1, 2, 1])
    with col_centro:
        st.plotly_chart(
            fig,
            use_container_width=True,
            key=f"chart_sefaz_altqtd_{st.session_state['sefaz_altqtd_key']}",
        )

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdedec; "
            f"border-radius:8px; border-left:4px solid #c0392b;'>"
            f"<span style='font-size:22px; font-weight:700; color:#c0392b;'>{total_alterada}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Quantidade Alterada ({pct_alterada}%)</span></div>",
            unsafe_allow_html=True,
        )
    with col_r:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#eafaf1; "
            f"border-radius:8px; border-left:4px solid #27ae60;'>"
            f"<span style='font-size:22px; font-weight:700; color:#27ae60;'>{total_mesma}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Mesma Quantidade ({pct_mesma}%)</span></div>",
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Ver empresas com quantidade alterada", use_container_width=True,
                 key="btn_sefaz_altqtd"):
        if not df_result.empty:
            _modal_sefaz_alteracao_quantidade(df_result)
        else:
            st.info("Nenhuma alteração de quantidade encontrada!")

    st.divider()

    # ── lista ─────────────────────────────────────────────────────────────────
    st.markdown("### Lista de Empresas com Quantidade Alterada", unsafe_allow_html=True)

    if not df_result.empty:
        df_exib = _sanitiza_df(df_result)
        exibe_aggrid(df_exib, height=400, grid_key="grid_sefaz_altqtd")

        output = BytesIO()
        df_result.to_excel(output, index=False)
        st.download_button(
            "Baixar Excel", data=output.getvalue(),
            file_name="sefaz_alteracao_quantidade_notas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.success("Nenhuma alteração de quantidade encontrada entre as abas!")

# ============================================================================
# ALVARÁS — DEPARTAMENTO PARALEGAL
# ============================================================================

# Aba ALVARA na PRÓPRIA planilha do escritório (a do GOOGLE_SHEET_URL). Se a
# aba ainda não existir, a lista nasce das empresas ATIVAS da GERAL e a aba é
# criada no primeiro "Salvar no Sheets". A gravação usa um Apps Script
# SEPARADO e COMPARTILHADO só de alvarás (apps_script_alvaras.gs, guardado em
# PROGRAMA/SCRIPTS_COMPARTILHADOS): MESMA URL nos 6 escritórios, só muda ALVARA_ESCRITORIO (chave em
# PLANILHAS no script). Não usa o script de certificados de propósito.
# Obs.: o VIDAL é diferente (planilha de alvarás separada + secrets).
_ABA_ALVARA = "ALVARA"
ALVARA_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzjBwCZLhIeiHGvdrfBUBqFCP87BwTk1ufuqmeT67qf3o7Fd18BMdWF-U7z81AmUfOO/exec"
ALVARA_ESCRITORIO = "VS"


def _normaliza_data_br(val):
    """Converte qualquer formato de data para DD/MM/AAAA. Retorna '' se inválido."""
    s = str(val).strip()
    if s in ("", "nan", "None", "NaT", "NaN"):
        return ""
    try:
        # Data de verdade na planilha chega como "AAAA-MM-DD 00:00:00": com
        # dayfirst=True o pandas trocava dia e mês (2027-09-03 virava 09/03/2027).
        iso = re.match(r"^\d{4}-\d{2}-\d{2}", s) is not None
        dt = pd.to_datetime(s, dayfirst=not iso, errors="coerce")
        if pd.isna(dt):
            dt = pd.to_datetime(s, dayfirst=iso, errors="coerce")
        return dt.strftime("%d/%m/%Y") if not pd.isna(dt) else ""
    except Exception:
        return ""


@st.cache_data(ttl=60)
def _alvara_carregar():
    try:
        resp = requests.get(GOOGLE_SHEET_URL, timeout=30)
        resp.raise_for_status()
        df = pd.read_excel(BytesIO(resp.content), sheet_name=_ABA_ALVARA,
                           engine="openpyxl", dtype=str)
        df.columns = df.columns.str.strip()
        for col in ["Vencimento Localização", "Vencimento Sanitário", "Vencimento Bombeiros",
                    "Vencimento Meio Ambiente"]:
            if col in df.columns:
                df[col] = df[col].fillna("").apply(_normaliza_data_br)
        return df
    except Exception:
        return pd.DataFrame()


def _alvara_salvar(df):
    """Regrava a aba ALVARA inteira via Apps Script de alvarás. Retorna (ok, mensagem)."""
    if not ALVARA_SCRIPT_URL:
        return False, ("O Apps Script de alvarás ainda não foi publicado "
                       "(ALVARA_SCRIPT_URL vazio — ver apps_script_alvaras.gs).")
    df_s = df.copy().fillna("").astype(str)
    payload = {"escritorio": ALVARA_ESCRITORIO,
               "alvaras": {"cabecalho": df_s.columns.tolist(), "linhas": df_s.values.tolist()}}
    try:
        resp = requests.post(ALVARA_SCRIPT_URL, json=payload, timeout=90).json()
    except Exception as e:
        return False, f"O Apps Script de alvarás não respondeu ({e})."
    if resp.get("status") != "ok" or _ABA_ALVARA not in (resp.get("gravadas") or []):
        return False, f"Erro do Apps Script: {resp.get('message') or resp}"
    return True, ""


def _classifica_vencimento_alvara(data_str):
    from datetime import date, timedelta
    today = date.today()
    vence30 = today + timedelta(days=30)
    try:
        dt = pd.to_datetime(data_str, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return "Sem Data"
        d = dt.date()
        if d < today:
            return "Vencido"
        elif d <= vence30:
            return "Vencendo"
        else:
            return "Válido"
    except Exception:
        return "Sem Data"


@st.dialog("Alvarás — Vencendo (até 30 dias)", width="large")
def _modal_alvara_vencendo(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) com alvará vencendo em até 30 dias**")
    st.dataframe(df_show.reset_index(drop=True), use_container_width=True, hide_index=True)


@st.dialog("Alvarás — Vencidos", width="large")
def _modal_alvara_vencidos(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) com alvará vencido**")
    st.dataframe(df_show.reset_index(drop=True), use_container_width=True, hide_index=True)


@st.dialog("Alvarás — Válidos", width="large")
def _modal_alvara_validos(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) com alvará válido**")
    st.dataframe(df_show.reset_index(drop=True), use_container_width=True, hide_index=True)


# ── Painel de Situação (modelo da planilha de alvarás do VIDAL) ─────────────
# Cada alvará cai em UM dos status abaixo, igual à planilha modelo do
# escritório, com duas diferenças: "Vencido" (tem alvará, data passou) e "Sem
# Alvará" (marcado NÃO) ficam separados — assim Válido + 30 dias + Vencido bate
# com o total "com alvará" dos donuts de cima — e "Não informado" é extra
# (linha ainda não preenchida no cadastro).
_ALV_TIPOS = [
    # (rótulo no painel, coluna situação, coluna vencimento, ícone)
    ("Bombeiros (Cercon)",        "Cert. Bombeiros",       "Vencimento Bombeiros",     "🚒"),
    ("Funcionamento",             "Alvará de Localização", "Vencimento Localização",   "🏢"),
    ("Sanitário",                 "Alvará Sanitário",      "Vencimento Sanitário",     "🩺"),
    ("Meio Ambiente",             "Meio Ambiente",         "Vencimento Meio Ambiente", "🌳"),
]
_ALV_STATUS = [
    # (status, cor forte, cor de fundo)
    ("Válido",               "#27ae60", "#eafaf1"),
    ("30 dias para vencer",  "#f39c12", "#fef5e7"),
    ("Em processo",          "#2e86de", "#eaf2fb"),
    ("Isento / Dispensado",  "#8e6bbf", "#f3eefa"),
    ("Vencido",              "#e74c3c", "#fdecea"),
    ("Sem Alvará",           "#5d6d7e", "#ebedef"),
    ("Não informado",        "#95a5a6", "#f2f4f4"),
]
_ALV_OPCOES = ["", "SIM", "NÃO", "ISENTO", "EM PROCESSO", "INDETERMINADO"]


def _alv_status_modelo(situacao, vencimento, data_ref):
    """Mesma regra da planilha modelo: data > ref+30 → Válido; data entre ref e
    ref+30 → 30 dias para vencer; data vencida → Vencido; NÃO → Sem Alvará;
    ISENTO → Isento / Dispensado; EM PROCESSO → Em processo; INDETERMINADO →
    Válido (alvará sem prazo de validade)."""
    from datetime import timedelta
    s = str(situacao).strip().upper()
    if s in ("", "NAN", "NONE"):
        return "Não informado"
    if s == "ISENTO":
        return "Isento / Dispensado"
    if s == "EM PROCESSO":
        return "Em processo"
    if s == "INDETERMINADO":
        return "Válido"
    if s == "NÃO":
        return "Sem Alvará"
    dt = pd.to_datetime(str(vencimento), dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return "Não informado"   # SIM sem data de vencimento
    d = dt.date()
    if d < data_ref:
        return "Vencido"
    if d <= data_ref + timedelta(days=30):
        return "30 dias para vencer"
    return "Válido"


def _alv_painel_situacao(df_work):
    from datetime import date
    st.markdown("### 📊 Painel de Situação dos Alvarás")

    c_ref, c_info = st.columns([1, 3])
    with c_ref:
        data_ref = st.date_input("Data de referência", value=date.today(),
                                 format="DD/MM/YYYY", key="alv_data_ref")
    with c_info:
        st.markdown(
            "<p style='font-size:12.5px; color:#666; margin-top:30px;'>"
            "Vence depois de 30 dias da data de referência = <b>Válido</b> · "
            "vence em até 30 dias = <b>30 dias para vencer</b> · já venceu = "
            "<b>Vencido</b> · marcado NÃO = <b>Sem Alvará</b>.</p>",
            unsafe_allow_html=True,
        )

    df_st = df_work[[c for c in ["Código", "Nome", "Município", "Estado"] if c in df_work.columns]].copy()
    for rotulo, col_sit, col_venc, _ in _ALV_TIPOS:
        df_st[rotulo] = [
            _alv_status_modelo(r.get(col_sit, ""), r.get(col_venc, ""), data_ref)
            for _, r in df_work.iterrows()
        ]
        df_st[f"Venc. {rotulo}"] = df_work[col_venc] if col_venc in df_work.columns else ""
    total = int(df_st.shape[0])

    # ── 4 cards (um por tipo de alvará) ───────────────────────────────────
    cards = []
    for rotulo, col_sit, _, icone in _ALV_TIPOS:
        cont = df_st[rotulo].value_counts()
        # mesmo critério do "com alvará" dos donuts de cima (SIM/INDETERMINADO)
        com_alvara = int(df_work[col_sit].astype(str).str.strip().str.upper()
                         .isin(["SIM", "INDETERMINADO"]).sum()) if col_sit in df_work.columns else 0
        em_dia = int(cont.get("Válido", 0) + cont.get("Isento / Dispensado", 0))
        pct_em_dia = (em_dia / total * 100) if total else 0
        barra = "".join(
            f"<div title='{s}: {int(cont.get(s, 0))}' style='width:{cont.get(s, 0) / total * 100 if total else 0:.2f}%;"
            f"background:{cor};'></div>"
            for s, cor, _ in _ALV_STATUS
        )
        linhas = "".join(
            f"<div style='display:flex; justify-content:space-between; align-items:center; "
            f"padding:3px 8px; margin:2px 0; border-radius:6px; background:{fundo};'>"
            f"<span style='font-size:12px; color:#333;'>"
            f"<span style='display:inline-block; width:9px; height:9px; border-radius:50%; "
            f"background:{cor}; margin-right:6px;'></span>{s}</span>"
            f"<span style='font-size:12px; color:{cor}; font-weight:700;'>"
            f"{int(cont.get(s, 0))} <span style='color:#888; font-weight:400;'>"
            f"({(cont.get(s, 0) / total * 100) if total else 0:.1f}%)</span></span></div>"
            for s, cor, fundo in _ALV_STATUS
        )
        cor_em_dia = "#27ae60" if pct_em_dia >= 70 else ("#f39c12" if pct_em_dia >= 40 else "#e74c3c")
        cards.append(
            f"<div style='flex:1 1 230px; background:white; border:1px solid #e3e8f0; "
            f"border-radius:14px; padding:14px 14px 10px; box-shadow:0 2px 8px rgba(29,63,119,.07);'>"
            f"<div style='display:flex; justify-content:space-between; align-items:flex-start;'>"
            f"<div><div style='font-size:14px; font-weight:700; color:#1d3f77;'>{icone} {rotulo}</div>"
            f"<div style='font-size:11px; color:#666; margin-top:2px;'>"
            f"<b style='color:#1d3f77;'>{com_alvara}</b> com alvará</div></div>"
            f"<div style='text-align:right;'><div style='font-size:22px; font-weight:800; "
            f"color:{cor_em_dia}; line-height:1;'>{pct_em_dia:.0f}%</div>"
            f"<div style='font-size:10px; color:#888;'>em dia</div></div></div>"
            f"<div style='display:flex; height:10px; border-radius:6px; overflow:hidden; "
            f"margin:10px 0 8px; background:#eef1f5;'>{barra}</div>"
            f"{linhas}</div>"
        )
    st.markdown(
        f"<div style='display:flex; flex-wrap:wrap; gap:14px; margin:6px 0 14px;'>{''.join(cards)}</div>"
        f"<p style='font-size:11.5px; color:#888; margin-top:-6px;'>"
        f"Percentuais sobre {total} empresa(s) do cadastro · <b>em dia</b> = Válido + Isento / Dispensado.</p>",
        unsafe_allow_html=True,
    )

    # ── Quadro resumo (igual ao topo da planilha modelo) ──────────────────
    with st.expander("📋 Quadro resumo (status × tipo de alvará)", expanded=False):
        th = "padding:7px 10px; background:#1d3f77; color:white; font-size:12.5px; text-align:center;"
        cab = f"<th style='{th} text-align:left;'>Status</th>" + "".join(
            f"<th style='{th}'>{icone} {rotulo}</th>" for rotulo, _, _, icone in _ALV_TIPOS
        )
        corpo = ""
        for s, cor, fundo in _ALV_STATUS:
            celulas = ""
            for rotulo, _, _, _ in _ALV_TIPOS:
                n = int((df_st[rotulo] == s).sum())
                pct = (n / total * 100) if total else 0
                celulas += (
                    f"<td style='padding:6px 10px; text-align:center; background:{fundo}; "
                    f"border-bottom:1px solid #fff;'><b style='color:{cor};'>{pct:.1f}%</b>"
                    f"<span style='color:#888; font-size:11px;'> ({n})</span></td>"
                )
            corpo += (
                f"<tr><td style='padding:6px 10px; font-size:12.5px; border-bottom:1px solid #eee;'>"
                f"<span style='display:inline-block; width:9px; height:9px; border-radius:50%; "
                f"background:{cor}; margin-right:6px;'></span>{s}</td>{celulas}</tr>"
            )
        corpo += (
            "<tr><td style='padding:6px 10px; font-weight:700; color:#1d3f77;'>Total</td>"
            + "".join(f"<td style='padding:6px 10px; text-align:center; font-weight:700; color:#1d3f77;'>"
                      f"100% ({total})</td>" for _ in _ALV_TIPOS)
            + "</tr>"
        )
        st.markdown(
            f"<div style='overflow-x:auto;'><table style='width:100%; border-collapse:collapse; "
            f"border-radius:10px; overflow:hidden;'><thead><tr>{cab}</tr></thead>"
            f"<tbody>{corpo}</tbody></table></div>",
            unsafe_allow_html=True,
        )

    # ── Situação por empresa (filtrável) ──────────────────────────────────
    st.markdown("#### 🔎 Situação por empresa")
    f1, f2, f3 = st.columns([1.2, 2, 1.5])
    with f1:
        tipo_sel = st.selectbox("Alvará", ["Todos"] + [t[0] for t in _ALV_TIPOS], key="alv_f_tipo")
    with f2:
        status_sel = st.multiselect("Situação", [s[0] for s in _ALV_STATUS], key="alv_f_status",
                                    placeholder="Todas")
    with f3:
        busca = st.text_input("Buscar empresa", key="alv_f_busca", placeholder="Nome ou código")

    tipos_filtro = [t[0] for t in _ALV_TIPOS] if tipo_sel == "Todos" else [tipo_sel]
    df_f = df_st.copy()
    if status_sel:
        df_f = df_f[df_f[tipos_filtro].isin(status_sel).any(axis=1)]
    if busca.strip():
        b = busca.strip().upper()
        df_f = df_f[df_f["Nome"].astype(str).str.upper().str.contains(b, regex=False)
                    | df_f["Código"].astype(str).str.contains(b, regex=False)]

    cols_base = [c for c in ["Código", "Nome", "Município"] if c in df_f.columns]
    if tipo_sel == "Todos":
        df_exib = df_f[cols_base + tipos_filtro].copy()
    else:
        df_exib = df_f[cols_base + [tipo_sel, f"Venc. {tipo_sel}"]].copy()
        df_exib = df_exib.rename(columns={tipo_sel: "Situação", f"Venc. {tipo_sel}": "Vencimento"})

        def _dias(v):
            dt = pd.to_datetime(str(v), dayfirst=True, errors="coerce")
            return "" if pd.isna(dt) else str((dt.date() - data_ref).days)
        df_exib["Dias p/ vencer"] = df_exib["Vencimento"].apply(_dias)

    cores = {s: (cor, fundo) for s, cor, fundo in _ALV_STATUS}

    def _pinta(v):
        if v in cores:
            cor, fundo = cores[v]
            return f"background-color:{fundo}; color:{cor}; font-weight:600;"
        return ""

    cols_status = tipos_filtro if tipo_sel == "Todos" else ["Situação"]
    st.caption(f"{df_exib.shape[0]} de {total} empresa(s)")
    st.dataframe(
        df_exib.reset_index(drop=True).style.map(_pinta, subset=cols_status),
        use_container_width=True, hide_index=True,
        height=min(38 + 35 * max(df_exib.shape[0], 1), 420),
    )


def pagina_alvaras():
    import plotly.graph_objects as go
    st.empty()

    st.markdown("<h2 style='color:#1d3f77;'>ALVARÁS</h2>", unsafe_allow_html=True)

    # ── Botões de controle ────────────────────────────────────────────────────
    col_rf, col_esp = st.columns([1, 4])
    with col_rf:
        if st.button("🔄 Atualizar do Sheets", key="btn_alvara_refresh", use_container_width=True):
            _alvara_carregar.clear()
            for k in ["alvara_df", "editor_alvaras"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()

    # ── Monta / carrega DataFrame de trabalho ─────────────────────────────────
    if "alvara_df" not in st.session_state:
        df_geral  = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
        df_sheets = _alvara_carregar()

        _COLS_EXTRAS = ["Usuário", "Senha",
                        "Alvará de Localização", "Vencimento Localização",
                        "Alvará Sanitário",       "Vencimento Sanitário",
                        "Cert. Bombeiros",         "Vencimento Bombeiros",
                        "Meio Ambiente",           "Vencimento Meio Ambiente",
                        "Taxa de Funcionamento"]

        if not df_sheets.empty and "Código" in df_sheets.columns:
            # Sheets já tem dados completos — usa diretamente
            df_base = df_sheets.copy()
            df_base["Código"] = df_base["Código"].astype(str).str.strip()
            for c in _COLS_EXTRAS:
                if c not in df_base.columns:
                    df_base[c] = ""

            # Adiciona empresas novas do GERAL que ainda não estão no Sheets
            if df_geral is not None and not df_geral.empty and "Situação" in df_geral.columns:
                mask_ativa = df_geral["Situação"].astype(str).str.upper() == "ATIVA"
                df_ativas = df_geral[mask_ativa].copy()
                df_ativas["Código"] = df_ativas["Código"].apply(
                    lambda v: str(v).strip().removesuffix(".0") if pd.notna(v) else ""
                )
                codigos_existentes = set(df_base["Código"])
                df_novos = df_ativas[~df_ativas["Código"].isin(codigos_existentes)][
                    [c for c in ["Código", "Razão Social", "CNPJ", "Município", "Estado"]
                     if c in df_ativas.columns]
                ].rename(columns={"Razão Social": "Nome"}).copy()
                if not df_novos.empty:
                    if "CNPJ" in df_novos.columns:
                        df_novos["CNPJ"] = df_novos["CNPJ"].apply(_formata_cnpj_mascara)
                    for c in _COLS_EXTRAS:
                        df_novos[c] = ""
                    df_base = pd.concat([df_base, df_novos], ignore_index=True)
        else:
            # Sheets vazio — constrói do GERAL com colunas em branco
            if df_geral is not None and not df_geral.empty and "Situação" in df_geral.columns:
                mask_ativa = df_geral["Situação"].astype(str).str.upper() == "ATIVA"
                df_base = df_geral[mask_ativa][
                    [c for c in ["Código", "Razão Social", "CNPJ", "Município", "Estado"]
                     if c in df_geral.columns]
                ].copy()
                df_base = df_base.rename(columns={"Razão Social": "Nome"})
                df_base["Código"] = df_base["Código"].apply(
                    lambda v: str(v).strip().removesuffix(".0") if pd.notna(v) else ""
                )
                if "CNPJ" in df_base.columns:
                    df_base["CNPJ"] = df_base["CNPJ"].apply(_formata_cnpj_mascara)
            else:
                df_base = pd.DataFrame(columns=["Código", "Nome", "CNPJ", "Município", "Estado"])
            for c in _COLS_EXTRAS:
                df_base[c] = ""

        df_base = df_base.reset_index(drop=True)
        st.session_state["alvara_df"] = df_base

    df_work = st.session_state["alvara_df"].copy()

    # ── Classificação por vencimento ──────────────────────────────────────────
    def _classifica_coluna(col_alvara, col_venc):
        resultado = []
        for _, row in df_work.iterrows():
            situacao = str(row.get(col_alvara, "")).strip().upper()
            if situacao == "INDETERMINADO":
                resultado.append("Válido")
                continue
            tem = situacao == "SIM"
            if not tem:
                resultado.append("Sem Alvará")
            else:
                resultado.append(_classifica_vencimento_alvara(str(row.get(col_venc, ""))))
        return resultado

    df_work["_st_loc"]  = _classifica_coluna("Alvará de Localização", "Vencimento Localização")
    df_work["_st_san"]  = _classifica_coluna("Alvará Sanitário",       "Vencimento Sanitário")
    df_work["_st_bomb"] = _classifica_coluna("Cert. Bombeiros",         "Vencimento Bombeiros")

    # ── Função de donut reutilizável ──────────────────────────────────────────
    def _donut(status_col, titulo, chart_key, col_venc, grp):
        serie = df_work[status_col]
        com_alvara = serie[serie != "Sem Alvará"]
        validos  = (com_alvara == "Válido").sum()
        vencendo = (com_alvara == "Vencendo").sum()
        vencidos = (com_alvara == "Vencido").sum()
        total    = int(com_alvara.shape[0])

        fig = go.Figure(data=[go.Pie(
            labels=["Válidos", "Vencendo", "Vencidos"],
            values=[int(validos), int(vencendo), int(vencidos)],
            hole=0.68,
            marker=dict(colors=["#27ae60", "#f39c12", "#e74c3c"],
                        line=dict(color="#ffffff", width=3)),
            textinfo="none",
            hovertemplate="<b>%{label}</b><br>%{value} empresa(s)<extra></extra>",
            direction="clockwise",
            sort=False,
        )])
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            showlegend=False,
            margin=dict(t=8, b=8, l=8, r=8),
            height=200,
            annotations=[dict(
                text=f"<b>{total}</b><br><span style='font-size:10px'>com alvará</span>",
                x=0.5, y=0.5, xanchor="center", yanchor="middle",
                showarrow=False, font=dict(size=18, color="#1d3f77"),
            )],
        )
        st.markdown(
            f"<h4 style='text-align:center; color:#1d3f77; margin:4px 0; font-size:14px;'>{titulo}</h4>",
            unsafe_allow_html=True,
        )
        st.plotly_chart(fig, use_container_width=True, key=chart_key)

        # Quadradinhos clicáveis: o próprio número abre a lista da empresa
        cols_modal = ["Código", "Nome", "CNPJ", "Município", col_venc]
        cartoes = [
            ("Válido",   validos,  "Válidos",  "ok",  _modal_alvara_validos),
            ("Vencendo", vencendo, "Vencendo", "ve",  _modal_alvara_vencendo),
            ("Vencido",  vencidos, "Vencidos", "vd",  _modal_alvara_vencidos),
        ]
        for coluna, (status, qtd, rotulo, tipo, modal) in zip(st.columns(3), cartoes):
            with coluna:
                with st.container(key=f"alvcard_{tipo}_{grp}"):
                    if st.button(f"**{qtd}**\n\n{rotulo}", key=f"btn_alv_{tipo}_{grp}",
                                 use_container_width=True,
                                 help=f"Clique para ver as empresas ({rotulo.lower()})"):
                        modal(df_work[df_work[status_col] == status][
                            [c for c in cols_modal if c in df_work.columns]
                        ].reset_index(drop=True))

    # ── Estilo dos quadradinhos + linha separando os 3 alvarás ────────────────
    _css_cards = ""
    for tipo, cor, fundo in [("ok", "#27ae60", "#eafaf1"),
                             ("ve", "#f39c12", "#fef9e7"),
                             ("vd", "#e74c3c", "#fdf2f2")]:
        _css_cards += (
            f"div[class*='st-key-alvcard_{tipo}_'] button {{"
            f"  background:{fundo} !important; border:1px solid {fundo} !important;"
            f"  border-left:3px solid {cor} !important; border-radius:8px !important;"
            f"  padding:9px 4px !important; min-height:0 !important; transition:all .15s; }}"
            f"div[class*='st-key-alvcard_{tipo}_'] button:hover {{"
            f"  border-color:{cor} !important; box-shadow:0 3px 10px rgba(0,0,0,.10);"
            f"  transform:translateY(-1px); }}"
            f"div[class*='st-key-alvcard_{tipo}_'] button p {{ margin:0 !important;"
            f"  font-size:11px !important; color:#333 !important; line-height:1.3 !important; }}"
            f"div[class*='st-key-alvcard_{tipo}_'] button p:first-child {{"
            f"  font-size:19px !important; color:{cor} !important; }}"
        )
    st.markdown(
        f"<style>{_css_cards}"
        ".st-key-alvgrp_loc, .st-key-alvgrp_san {"
        "  border-right:1px solid #d5dce8; padding-right:18px; }"
        "</style>",
        unsafe_allow_html=True,
    )

    # ── 3 Dashboards lado a lado ──────────────────────────────────────────────
    col_d1, col_d2, col_d3 = st.columns(3, gap="medium")

    with col_d1:
        with st.container(key="alvgrp_loc"):
            _donut("_st_loc", "Alvará de Localização e Funcionamento", "chart_alv_loc",
                   "Vencimento Localização", "loc")

    with col_d2:
        with st.container(key="alvgrp_san"):
            _donut("_st_san", "Alvará Sanitário", "chart_alv_san",
                   "Vencimento Sanitário", "san")

    with col_d3:
        with st.container(key="alvgrp_bomb"):
            _donut("_st_bomb", "Certificado do Corpo de Bombeiros", "chart_alv_bomb",
                   "Vencimento Bombeiros", "bomb")

    st.caption("Clique no número de cada quadradinho para ver a lista de empresas.")
    st.divider()

    # ── Total de empresas com cada alvará ─────────────────────────────────────
    total_loc  = (df_work["Alvará de Localização"].astype(str).str.upper() == "SIM").sum()
    total_san  = (df_work["Alvará Sanitário"].astype(str).str.upper() == "SIM").sum()
    total_bomb = (df_work["Cert. Bombeiros"].astype(str).str.upper() == "SIM").sum()
    total_amb  = (df_work["Meio Ambiente"].astype(str).str.upper() == "SIM").sum()

    st.markdown(
        f"<div style='background:#f4f6fa; border-radius:10px; padding:12px 16px; margin-bottom:12px;'>"
        f"<b style='color:#1d3f77;'>Total de empresas com cada alvará:</b> &nbsp;&nbsp;"
        f"<span style='color:#1d3f77; font-weight:600;'>Localização:</span> <b>{total_loc}</b>"
        f" &nbsp;|&nbsp; "
        f"<span style='color:#1d3f77; font-weight:600;'>Sanitário:</span> <b>{total_san}</b>"
        f" &nbsp;|&nbsp; "
        f"<span style='color:#1d3f77; font-weight:600;'>Bombeiros:</span> <b>{total_bomb}</b>"
        f" &nbsp;|&nbsp; "
        f"<span style='color:#1d3f77; font-weight:600;'>Meio Ambiente:</span> <b>{total_amb}</b>"
        f"</div>",
        unsafe_allow_html=True,
    )

    # ── Painel de Situação (modelo da planilha do escritório) ─────────────────
    _alv_painel_situacao(df_work)
    st.divider()

    # ── Tabela editável ───────────────────────────────────────────────────────
    st.markdown("### Cadastro de Alvarás")
    st.markdown(
        "<p style='font-size:13px; color:#666;'>"
        "Nas colunas de alvará use <b>SIM</b>, <b>NÃO</b>, <b>ISENTO</b>, <b>EM PROCESSO</b> "
        "ou <b>INDETERMINADO</b> (alvará sem prazo de validade). "
        "Informe datas no formato <b>DD/MM/AAAA</b>. "
        "Clique na célula para preencher; <b>Código</b> e <b>Nome</b> ficam fixos ao rolar para o lado. "
        "Clique em <b>Salvar no Sheets</b> para não perder os dados.</p>",
        unsafe_allow_html=True,
    )

    cols_exib = ["Código", "Nome", "CNPJ", "Município", "Estado",
                 "Usuário", "Senha",
                 "Alvará de Localização", "Vencimento Localização",
                 "Alvará Sanitário", "Vencimento Sanitário",
                 "Cert. Bombeiros", "Vencimento Bombeiros",
                 "Meio Ambiente", "Vencimento Meio Ambiente",
                 "Taxa de Funcionamento"]
    df_edit = df_work[[c for c in cols_exib if c in df_work.columns]].copy()

    # Garante strings puras — SelectboxColumn não aceita NaN/float
    for col in df_edit.columns:
        df_edit[col] = df_edit[col].fillna("").astype(str).replace({"nan": "", "None": "", "NaT": ""})
    for col in ["Alvará de Localização", "Alvará Sanitário", "Cert. Bombeiros", "Meio Ambiente"]:
        if col in df_edit.columns:
            df_edit[col] = df_edit[col].str.strip().str.upper().apply(
                lambda v: v if v in _ALV_OPCOES else "")
    for col in ["Vencimento Localização", "Vencimento Sanitário", "Vencimento Bombeiros",
                "Vencimento Meio Ambiente"]:
        if col in df_edit.columns:
            df_edit[col] = df_edit[col].apply(_normaliza_data_br)

    # Grade de preenchimento em AgGrid (e não st.data_editor) porque o
    # data_editor não deixa pintar linhas: aqui Código/Nome ficam congelados à
    # esquerda e as linhas alternam fundo claro/azulado com divisória mais
    # forte, pra não se perder na linha ao preencher.
    from st_aggrid import JsCode
    _cols_fixas = ["Código", "Nome", "CNPJ", "Município", "Estado"]
    _cabecalhos = {
        "Alvará de Localização": "Alvará Localização", "Vencimento Localização": "Vencto. Localização",
        "Vencimento Sanitário": "Vencto. Sanitário", "Vencimento Bombeiros": "Vencto. Bombeiros",
        "Vencimento Meio Ambiente": "Vencto. Meio Ambiente",
    }
    gb = GridOptionsBuilder.from_dataframe(df_edit)
    gb.configure_default_column(editable=True, resizable=True, sortable=True, filter=True,
                                minWidth=110, wrapHeaderText=True, autoHeaderHeight=True)
    for col in df_edit.columns:
        opcoes = dict(headerName=_cabecalhos.get(col, col))
        if col in _cols_fixas:
            opcoes.update(editable=False, cellStyle={"color": "#4a5568"})
        if col in ("Alvará de Localização", "Alvará Sanitário", "Cert. Bombeiros", "Meio Ambiente"):
            opcoes.update(cellEditor="agSelectCellEditor",
                          cellEditorParams={"values": _ALV_OPCOES},
                          cellStyle={"fontWeight": "600", "color": "#1d3f77"})
        gb.configure_column(col, **opcoes)
    gb.configure_column("Código", pinned="left", width=90, minWidth=80, filter="agTextColumnFilter")
    gb.configure_column("Nome", pinned="left", width=300, minWidth=200, filter="agTextColumnFilter")
    gb.configure_grid_options(
        domLayout="normal", floatingFilter=True, headerHeight=40, rowHeight=32,
        singleClickEdit=True, stopEditingWhenCellsLoseFocus=True,
        getRowStyle=JsCode(
            "function(p){ return (p.node.rowIndex % 2 === 0)"
            " ? {background:'#ffffff'} : {background:'#e3ebf7'}; }"),
        localeText={'filterOoo': 'Filtrar...', 'contains': 'Contém', 'equals': 'Igual',
                    'noRowsToShow': 'Nenhum registro para mostrar'},
    )
    _cols_edit = list(df_edit.columns)
    # cópia: o AgGrid 1.x acrescenta a coluna interna "::auto_unique_id::" no df recebido
    resp = AgGrid(
        df_edit.copy(), gridOptions=gb.build(), height=500, key="editor_alvaras",
        columns_auto_size_mode=ColumnsAutoSizeMode.FIT_CONTENTS,
        enable_enterprise_modules=False, allow_unsafe_jscode=True, reload_data=False,
        update_on=["cellValueChanged"], data_return_mode=DataReturnMode.AS_INPUT,
        custom_css={
            ".ag-row": {"border-bottom": "1px solid #9fb0c8 !important"},
            ".ag-row-hover": {"background-color": "#fff6d6 !important"},
            ".ag-pinned-left-cols-container": {"border-right": "2px solid #1d3f77 !important"},
            ".ag-pinned-left-header": {"border-right": "2px solid #1d3f77 !important"},
        },
    )
    df_editado = resp.data if resp is not None and resp.data is not None else df_edit
    df_editado = pd.DataFrame(df_editado).reindex(columns=_cols_edit).fillna("").astype(str)

    # ── Botão Salvar ──────────────────────────────────────────────────────────
    col_sv, _ = st.columns([1, 3])
    with col_sv:
        if st.button("💾 Salvar no Sheets", key="btn_alvara_salvar",
                     type="primary", use_container_width=True):
            st.session_state["alvara_df"] = df_editado.copy()
            if "editor_alvaras" in st.session_state:
                del st.session_state["editor_alvaras"]
            ok, msg = _alvara_salvar(df_editado)
            if ok:
                _alvara_carregar.clear()
                st.success("Dados salvos com sucesso no Google Sheets!")
            else:
                st.error(f"Não foi possível salvar no Sheets. {msg}")


# ============================================================================
# ROTEAMENTO
# ============================================================================

with st.session_state.main_container.container():
    if st.session_state["menu_area"] == "SITUAÇÃO FISCAL" and pagina == "DASHBOARD":
        pagina_situacao_fiscal_dashboard()
    elif st.session_state["menu_area"] == "SITUAÇÃO FISCAL" and pagina == "EMPRESAS":
        pagina_situacao_fiscal_empresas()
    elif st.session_state["menu_area"] == "SITUAÇÃO FISCAL" and pagina == "CAIXA POSTAL":
        pagina_caixa_postal()
    elif pagina == "DASHBOARD":
        pagina_dashboard_paralegal()
    elif pagina == "ALVARÁS":
        pagina_alvaras()
    elif pagina == "EMPRESAS":
        pagina_empresas()
    elif pagina == "SIMPLES NACIONAL":
        pagina_simples()
    elif pagina == "REINF":
        pagina_reinf()
    elif pagina == "DCTF WEB":
        pagina_dctf_web()
    elif pagina == "DMS":
        pagina_dms()
    elif pagina == "SERVIÇOS TOMADOS":
        pagina_rest()
    elif pagina == "SEFAZ":
        pagina_sefaz()
    elif pagina == "LEITURA XML DMS":
        pagina_leitura_xml_dms()
    elif pagina == "LEITURA XML REST":
        pagina_leitura_xml_rest()    
    elif pagina == "CND MUNICIPAL":
        pagina_cnd_municipal()
    elif pagina == "SEM ACESSO":
        pagina_sem_acesso()
    elif pagina == "SEFAZ ALTERAÇÃO QUANTIDADE NOTAS":
        pagina_sefaz_alteracao_quantidade_notas()
    elif pagina == "CERTIFICADOS":
        pagina_certificados()
    elif pagina == "ENDEREÇO DE EMAIL":
        pagina_emails_cnpj()
    elif pagina == "MENSAGENS DE EMAIL":
        pagina_mensagens_email()