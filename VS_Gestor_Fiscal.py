# ============================================================================
# GESTOR FISCAL - LUATECH
# Sistema de Gestão Fiscal com Streamlit
# ============================================================================

import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from io import BytesIO
import requests
import time
import os
import base64
import json
from pathlib import Path
from datetime import date

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


def exibe_aggrid(df, height=400, grid_key="grid", selection_mode='none'):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filter=True, sortable=True, editable=False, resizable=True)
    
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
    
    return AgGrid(df, gridOptions=grid_options, height=height, key=grid_key,
                  fit_columns_on_grid_load=True, enable_enterprise_modules=False,
                  update_on=update_on, allow_unsafe_jscode=True, reload_data=False)


def exibe_aggrid_com_oculta(df, height=400, grid_key="grid", selection_mode='none', colunas_ocultas=None):
    if colunas_ocultas is None:
        colunas_ocultas = []
    
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filter=True, sortable=True, editable=False, resizable=True)
    
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
                  fit_columns_on_grid_load=True, enable_enterprise_modules=False,
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


def _cert_ler_pfx(caminho: str, senha: str):
    """Lê um PFX e retorna (razao_social, documento, validade_str, validade_iso).
    documento pode ser CNPJ (14 dígitos) ou CPF (11 dígitos).
    """
    import re as _re
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography import x509
    from datetime import timezone

    pfx_bytes = Path(caminho).read_bytes()
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


# ── página CERTIFICADOS ────────────────────────────────────────────────────────
def pagina_certificados():
    st.markdown("<h2>CERTIFICADOS DIGITAIS</h2>", unsafe_allow_html=True)

    dados = _cert_carregar_dados()
    certs = dados["certificados"]

    # ── Importar da pasta ─────────────────────────────────────────────────────
    with st.expander("📁 Importar Certificados da Pasta", expanded=not certs):
        col_p1, col_p2 = st.columns([5, 1])
        with col_p1:
            pasta_digitada = st.text_input(
                "Caminho da pasta com os certificados (.pfx):",
                value=st.session_state.get("cert_pasta", ""),
                key="cert_pasta_input",
                placeholder="Ex.: C:\\Certificados",
            )
        with col_p2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📂 Selecionar", key="btn_pasta_picker", use_container_width=True):
                p = _picker_pasta_cert()
                if p:
                    st.session_state["cert_pasta"] = p
                    st.rerun()

        if pasta_digitada:
            st.session_state["cert_pasta"] = pasta_digitada
            pasta_path = Path(pasta_digitada)
            # Busca recursiva em todas as subpastas
            pfx_encontrados = sorted(pasta_path.rglob("*.pfx")) if pasta_path.exists() else []
            arquivos_ja = {c["arquivo"] for c in certs}
            novos = [p for p in pfx_encontrados if str(p) not in arquivos_ja]

            if not pfx_encontrados:
                st.warning("Nenhum arquivo .pfx encontrado nesta pasta (incluindo subpastas).")
            elif not novos:
                st.info(f"{len(pfx_encontrados)} certificado(s) encontrado(s) — todos já estão na lista.")
            else:
                # Classifica: com senha detectada no nome vs. sem senha
                com_senha = []   # (Path, senha_detectada)
                sem_senha = []   # Path
                for pfx in novos:
                    s = _extrair_senha_nome(pfx.name)
                    if s:
                        com_senha.append((pfx, s))
                    else:
                        sem_senha.append(pfx)

                st.markdown(
                    f"**{len(novos)} novo(s) certificado(s) encontrado(s) "
                    f"(pasta e subpastas):**"
                )

                senhas_novas = {}

                # Campo de senha padrão para arquivos sem senha no nome
                senha_padrao = ""
                if sem_senha:
                    senha_padrao = st.text_input(
                        f"🔑 Senha padrão para os {len(sem_senha)} certificado(s) sem senha no nome:",
                        type="password",
                        key="cert_senha_padrao",
                        placeholder="Digite a senha padrão",
                    )

                # Arquivos com senha detectada no nome
                if com_senha:
                    st.markdown("**Com senha detectada no nome do arquivo:**")
                    for pfx, senha_auto in com_senha:
                        try:
                            rel = pfx.relative_to(pasta_path)
                        except Exception:
                            rel = pfx.name
                        c1, c2, c3 = st.columns([4, 2, 2])
                        with c1:
                            st.markdown(f"`{rel}`")
                        with c2:
                            st.markdown(f"🔒 Detectada: `{senha_auto}`")
                        with c3:
                            override = st.text_input(
                                "Substituir", type="password",
                                key=f"sn_{abs(hash(str(pfx)))}",
                                label_visibility="collapsed",
                                placeholder="substituir (opcional)",
                            )
                        senhas_novas[str(pfx)] = override if override else senha_auto

                # Arquivos sem senha no nome
                if sem_senha:
                    st.markdown("**Sem senha no nome (usarão a senha padrão acima):**")
                    for pfx in sem_senha:
                        try:
                            rel = pfx.relative_to(pasta_path)
                        except Exception:
                            rel = pfx.name
                        c1, c2 = st.columns([4, 2])
                        with c1:
                            st.markdown(f"`{rel}`")
                        with c2:
                            override = st.text_input(
                                "Senha individual", type="password",
                                key=f"sn_{abs(hash(str(pfx)))}",
                                label_visibility="collapsed",
                                placeholder="ou senha individual",
                            )
                        senhas_novas[str(pfx)] = override if override else senha_padrao

                if st.button("✅ Importar todos", key="btn_importar_pasta"):
                    adicionados, erros = 0, []
                    for caminho, senha in senhas_novas.items():
                        if not senha:
                            erros.append(f"{Path(caminho).name}: senha não informada.")
                            continue
                        try:
                            razao, cnpj, val_str, val_iso = _cert_ler_pfx(caminho, senha)
                            certs.append({
                                "arquivo": caminho,
                                "nome_arquivo": Path(caminho).name,
                                "senha": senha,
                                "razao_social": razao or Path(caminho).stem,
                                "cnpj": cnpj,
                                "validade": val_str,
                                "validade_iso": val_iso,
                            })
                            adicionados += 1
                        except Exception as e:
                            erros.append(f"{Path(caminho).name}: {e}")
                    dados["certificados"] = certs
                    _cert_salvar_dados(dados)
                    if adicionados:
                        st.success(f"{adicionados} certificado(s) importado(s)!")
                    for err in erros:
                        st.error(err)
                    st.rerun()

    # ── Adicionar individual ──────────────────────────────────────────────────
    with st.expander("➕ Adicionar Certificado Individual", expanded=False):
        col_a1, col_a2 = st.columns([5, 1])
        with col_a1:
            add_caminho = st.text_input(
                "Caminho do arquivo .pfx:",
                value=st.session_state.get("add_cert_path", ""),
                key="add_cert_caminho",
                placeholder="Ex.: C:\\Certificados\\empresa.pfx",
            )
        with col_a2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("📂 Selecionar", key="btn_pick_arquivo", use_container_width=True):
                arq = _picker_arquivo_cert()
                if arq:
                    st.session_state["add_cert_path"] = arq
                    # Auto-detecta senha pelo nome do arquivo
                    senha_auto = _extrair_senha_nome(Path(arq).name)
                    if senha_auto:
                        st.session_state["add_cert_senha_auto"] = senha_auto
                    else:
                        st.session_state.pop("add_cert_senha_auto", None)
                    st.rerun()

        # Caminho efetivo (pode vir do picker ou digitação)
        add_caminho = add_caminho or st.session_state.get("add_cert_path", "")

        # Senha: auto-detectada (editável) ou campo em branco
        senha_auto = st.session_state.get("add_cert_senha_auto", "")
        col_s1, col_s2 = st.columns([3, 3])
        with col_s1:
            if senha_auto:
                st.markdown(f"🔒 Senha detectada no nome: `{senha_auto}`")
            else:
                st.markdown("Nenhuma senha detectada no nome do arquivo.")
        with col_s2:
            add_senha = st.text_input(
                "Senha:" if not senha_auto else "Substituir senha:",
                type="password",
                key="add_cert_senha",
                placeholder="senha detectada será usada" if senha_auto else "Digite a senha",
            )
        senha_final = add_senha if add_senha else senha_auto

        if st.button("✅ Testar e Adicionar", key="btn_add_individual", type="primary"):
            if not add_caminho:
                st.error("Selecione ou informe o caminho do arquivo .pfx.")
            elif not senha_final:
                st.error("Informe a senha do certificado.")
            elif not Path(add_caminho).exists():
                st.error("Arquivo não encontrado.")
            elif add_caminho in {c["arquivo"] for c in certs}:
                st.warning("Este certificado já está na lista.")
            else:
                try:
                    razao, cnpj, val_str, val_iso = _cert_ler_pfx(add_caminho, senha_final)
                    certs.append({
                        "arquivo": add_caminho,
                        "nome_arquivo": Path(add_caminho).name,
                        "senha": senha_final,
                        "razao_social": razao or Path(add_caminho).stem,
                        "cnpj": cnpj,
                        "validade": val_str,
                        "validade_iso": val_iso,
                    })
                    dados["certificados"] = certs
                    _cert_salvar_dados(dados)
                    st.session_state.pop("add_cert_path", None)
                    st.session_state.pop("add_cert_senha_auto", None)
                    st.success(f"Certificado adicionado: {razao or Path(add_caminho).stem}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao ler o certificado: {e}")

    st.divider()

    if not certs:
        st.info("Nenhum certificado cadastrado. Use as opções acima para importar.")
        return

    # ── Contadores ────────────────────────────────────────────────────────────
    rows = []
    for c in certs:
        sit, dias = _cert_situacao(c.get("validade_iso", ""))
        cnpj_fmt = _formata_cnpj_mascara(c["cnpj"]) if c.get("cnpj") else ""
        rows.append({
            "Razão Social": c.get("razao_social", ""),
            "CPF/CNPJ": cnpj_fmt,
            "Validade": c.get("validade", ""),
            "Dias": dias if dias is not None else "?",
            "Situação": sit,
            "_arquivo": c["arquivo"],
            "_sit": sit,
        })

    n_vencidos = sum(1 for r in rows if r["_sit"] == "VENCIDO")
    n_vencendo = sum(1 for r in rows if r["_sit"] == "VENCENDO")
    n_normais  = sum(1 for r in rows if r["_sit"] == "NORMAL")
    total_certs = len(rows)

    st.markdown(
        f"<p style='text-align:right; font-size:18px;'><b>Total:</b> {total_certs}</p>",
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns(3)
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

    st.markdown("<br>", unsafe_allow_html=True)

    # ── Filtro ────────────────────────────────────────────────────────────────
    filtro_cert = st.radio(
        "Filtrar por:",
        ["Todos", "Vencidos", "Vencendo em 30 dias", "Normais"],
        horizontal=True, key="filtro_cert",
    )

    mapa_sit = {"Todos": None, "Vencidos": "VENCIDO", "Vencendo em 30 dias": "VENCENDO", "Normais": "NORMAL"}
    alvo_sit = mapa_sit[filtro_cert]
    rows_filtradas = [r for r in rows if alvo_sit is None or r["_sit"] == alvo_sit]

    if not rows_filtradas:
        st.info("Nenhum certificado neste filtro.")
    else:
        df_cert = pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")} for r in rows_filtradas])
        exibe_aggrid(df_cert, height=350, grid_key=f"grid_certs_{filtro_cert}")

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
                        key=f"chk_{abs(hash(c['arquivo']))}",
                    )
                    for c in certs
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

    col1, col2, col3, col4 = st.columns(4)

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
label_area = {"FISCAL": "FISCAL", "PARALEGAL": "DEPARTAMENTO PARALEGAL", "CONTÁBIL": "CONTÁBIL", "CERTIFICADO DIGITAL": "CERTIFICADO DIGITAL"}
st.sidebar.markdown(f"<p style='text-align:center; color:#1d3f77; font-weight:bold; margin-top:10px;'>{label_area.get(st.session_state['menu_area'], st.session_state['menu_area'])}</p>", unsafe_allow_html=True)

if st.sidebar.button("← DEPARTAMENTOS", use_container_width=True):
    st.session_state["menu_area"] = None
    st.session_state["pagina_atual"] = None
    st.rerun()

st.sidebar.markdown("<hr style='margin: 8px 0;'>", unsafe_allow_html=True)

# Define as páginas disponíveis por área
if st.session_state["menu_area"] == "FISCAL":
    paginas_disponiveis = ["EMPRESAS", "SIMPLES NACIONAL", "REINF", "DCTF WEB",
                           "DMS", "SERVIÇOS TOMADOS", "SEFAZ", "LEITURA XML DMS", "LEITURA XML REST","SEFAZ COMPARAÇÃO"]

elif st.session_state["menu_area"] == "PARALEGAL":
    paginas_disponiveis = ["DASHBOARD", "EMPRESAS", "CND MUNICIPAL", "SEM ACESSO"]

elif st.session_state["menu_area"] == "CONTÁBIL":
    paginas_disponiveis = ["EMPRESAS"]

elif st.session_state["menu_area"] == "CERTIFICADO DIGITAL":
    paginas_disponiveis = ["CERTIFICADOS", "ENDEREÇO DE EMAIL", "MENSAGENS DE EMAIL"]

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
               "XML ENTRADA", "XML SAÍDA", "IMPORTAÇÃO",
               "TOTAL ENTRADA", "TOTAL SAÍDA", "TOTAL DOMÍNIO", "Situação"]
    df_sefaz = df_sefaz[[c for c in colunas if c in df_sefaz.columns]].copy()

    # ── CNPJ: 14 dígitos ─────────────────────────────────────────────────────
    if "CNPJ" in df_sefaz.columns:
        df_sefaz["CNPJ"] = df_sefaz["CNPJ"].apply(_normaliza_cnpj)

    # ── colunas numéricas ─────────────────────────────────────────────────────
    for col in ["TOTAL ENTRADA", "TOTAL SAÍDA", "TOTAL DOMÍNIO"]:
        if col in df_sefaz.columns:
            df_sefaz[col] = pd.to_numeric(df_sefaz[col], errors="coerce").fillna(0)

    # ── coluna Confronto ──────────────────────────────────────────────────────  ← NOVO
    if all(c in df_sefaz.columns for c in ["XML ENTRADA", "XML SAÍDA", "TOTAL DOMÍNIO"]):
        xml_entrada = pd.to_numeric(df_sefaz["XML ENTRADA"],   errors="coerce").fillna(0)
        xml_saida   = pd.to_numeric(df_sefaz["XML SAÍDA"],     errors="coerce").fillna(0)
        total_dom   = pd.to_numeric(df_sefaz["TOTAL DOMÍNIO"], errors="coerce").fillna(0)
        soma_xml    = xml_entrada + xml_saida

        def _confronto(idx):
            s = soma_xml[idx]
            d = total_dom[idx]
            if s == 0 and d == 0:
                return "Importação OK"
            if s == d:
                return "Importação OK"
            return "Quantidade Diferente"

        df_sefaz["Confronto"] = [_confronto(i) for i in df_sefaz.index]
    else:
        df_sefaz["Confronto"] = "Importação OK"

    cols = [c for c in df_sefaz.columns if c not in ("Confronto", "Situação")]
    cols_final = cols + ["Confronto"]
    if "Situação" in df_sefaz.columns:
        cols_final = cols + ["Confronto", "Situação"]
    df_sefaz = df_sefaz[cols_final]

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

    st.markdown("<h2>SEFAZ</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Com Movimento:</b> {com_movimento} &nbsp;|&nbsp; "
        f"<b>Sem Acesso:</b> {sem_acesso} &nbsp;|&nbsp; "
        f"<b>Sem Busca:</b> {sem_busca} &nbsp;|&nbsp; "
        f"<b>Sem Movimento:</b> {sem_movimento} &nbsp;|&nbsp; "
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

    # ── botões das listas ─────────────────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    b1, b2, b3 = st.columns(3)
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
    st.empty()
    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return
    
    competencia_raw = df.get("PERÍODO DE COMPETÊNCIA", [""])[0]
    competencia = pd.to_datetime(competencia_raw, errors='coerce').strftime("%m/%Y") if competencia_raw else ""
    
    df_cnd = df[df["Situação"].astype(str).str.upper() == "ATIVA"] if "Situação" in df.columns else pd.DataFrame()
    if df_cnd.empty:
        st.warning("Nenhuma empresa ATIVA encontrada.")
        return
    
    colunas_solicitadas = ["Código", "Razão Social", "CNPJ", "Município", "Estado", 
                           "SITUAÇÃO CND MUNICIPAL", "VALIDADE", "LINK CND MUNICIPAL", "Situação"]
    colunas_existentes = [c for c in colunas_solicitadas if c in df_cnd.columns]
    df_cnd = df_cnd[colunas_existentes].copy()
    
    if "VALIDADE" in df_cnd.columns:
        df_cnd["VALIDADE"] = pd.to_datetime(df_cnd["VALIDADE"], errors='coerce').dt.strftime("%d/%m/%Y").fillna("")
    
    def check_pdf_link(link):
        return "Disponível" if pd.notna(link) and str(link).strip() != "" else "Indisponível"
    
    if "LINK CND MUNICIPAL" in df_cnd.columns:
        df_cnd["PDF"] = df_cnd["LINK CND MUNICIPAL"].apply(check_pdf_link)
    else:
        df_cnd["PDF"] = "Indisponível"
    
    if "SITUAÇÃO CND MUNICIPAL" in df_cnd.columns:
        situacao_upper = df_cnd["SITUAÇÃO CND MUNICIPAL"].astype(str).str.upper().str.strip()
        positivas = (situacao_upper == "POSITIVA").sum()
        negativas = (situacao_upper == "NEGATIVA").sum()
        positiva_efeito_negativa = (situacao_upper == "POSITIVA COM EFEITO NEGATIVA").sum()
    if "SITUAÇÃO CND MUNICIPAL" in df_cnd.columns:
        situacao_upper = df_cnd["SITUAÇÃO CND MUNICIPAL"].astype(str).str.upper().str.strip()
        positivas = (situacao_upper == "POSITIVA").sum()
        negativas = (situacao_upper == "NEGATIVA").sum()
        positiva_efeito_negativa = (situacao_upper == "POSITIVA COM EFEITO NEGATIVA").sum()
        nao_geradas = ((situacao_upper == "") | (situacao_upper == "NAN") | df_cnd["SITUAÇÃO CND MUNICIPAL"].isna()).sum()
    else:
        positivas = negativas = positiva_efeito_negativa = nao_geradas = 0
    
    total_geral = df_cnd.shape[0]
    
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
            
            if link_pdf and "drive.google.com" in str(link_pdf):
                if "/file/d/" in link_pdf:
                    file_id = link_pdf.split("/file/d/")[1].split("/")[0]
                    download_url = f"https://drive.google.com/uc?export=download&id={file_id}"
                    
                    st.markdown(
                        f'<a href="{download_url}" target="_blank">'
                        f'<button style="background-color:#1d3f77; color:white; padding:8px 16px; '
                        f'border:none; border-radius:4px; cursor:pointer; font-size:14px;">'
                        f'📥 Baixar PDF</button></a>',
                        unsafe_allow_html=True
                    )
        
        st.divider()
        row = st.session_state.pdf_selecionado
        cnpj = row.get("CNPJ", "")
        razao = row.get("Razão Social", "")
        link_pdf = row.get("LINK CND MUNICIPAL", "")
        status_pdf = row.get("PDF", "Indisponível")
        
        st.subheader(f"📄 {razao}")
        st.caption(f"CNPJ: {cnpj}")
        
        if status_pdf == "Disponível" and link_pdf:
            try:
                if "drive.google.com" in str(link_pdf):
                    if "/file/d/" in link_pdf:
                        file_id = link_pdf.split("/file/d/")[1].split("/")[0]
                        embed_url = f"https://drive.google.com/file/d/{file_id}/preview"
                        st.markdown(f'<iframe src="{embed_url}" width="100%" height="800" frameborder="0"></iframe>',
                                    unsafe_allow_html=True)
                    else:
                        st.error("❌ Formato de link do Google Drive não reconhecido")
                        st.info(f"Link: {link_pdf}")
                else:
                    st.markdown(f'<iframe src="{link_pdf}" width="100%" height="800" frameborder="0"></iframe>',
                                unsafe_allow_html=True)
            except Exception as e:
                st.error(f"❌ Erro ao carregar PDF: {e}")
                st.info(f"Link: {link_pdf}")
        else:
            st.error("❌ PDF não disponível")
            st.info("Link do PDF não foi encontrado na planilha (coluna LINK CND MUNICIPAL)")
    else:
        st.markdown(f"<h2>CND Municipal</h2><p style='text-align:right; font-size:20px;'>"
                    f"<b>Positivas:</b> {positivas} | <b>Negativas:</b> {negativas} | "
                    f"<b>Positiva c/ efeito negativa:</b> {positiva_efeito_negativa} | "
                    f"<b>Não geradas:</b> {nao_geradas} | <b>Total:</b> {total_geral} | "
                    f"<b>Competência:</b> {competencia}</p>", unsafe_allow_html=True)
        
        st.info("💡 Selecione uma linha na tabela para visualizar o PDF correspondente.")
        
        with st.container():
            df_cnd = _sanitiza_df(df_cnd)
            grid_response = exibe_aggrid_com_oculta(df_cnd, height=400, grid_key="grid_cnd_municipal",
                                                     selection_mode='single',
                                                     colunas_ocultas=["Situação", "LINK CND MUNICIPAL"])
        
        selected_rows = grid_response.get('selected_rows', [])
        if selected_rows is not None and len(selected_rows) > 0:
            row = selected_rows.iloc[0].to_dict() if isinstance(selected_rows, pd.DataFrame) else selected_rows[0]
            st.session_state.pdf_selecionado = row
            st.session_state.visualizando_pdf = True
            st.rerun()
        
        output = BytesIO()
        df_cnd.to_excel(output, index=False)
        st.download_button("📥 Baixar Excel", data=output.getvalue(), file_name="cnd_municipal.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ============================================================================
# DASHBOARD PARALEGAL
# ============================================================================

def _formata_estado(s):
    """GO → G.O.  — impede Chrome de traduzir siglas"""
    s = str(s).strip()
    if len(s) == 2 and s.isalpha():
        return f"{s[0]}.{s[1]}."
    return s


def _estado_original(s):
    """G.O. → GO"""
    return s.replace(".", "")


@st.dialog("Detalhes das Empresas")
def _modal_dashboard(titulo, df_show, colunas):
    st.markdown(f"**{titulo}** — {df_show.shape[0]} empresa(s)")
    cols_ok = [c for c in colunas if c in df_show.columns]
    df_exib = df_show[cols_ok].reset_index(drop=True).copy()
    if "CNPJ" in df_exib.columns:
        df_exib["CNPJ"] = df_exib["CNPJ"].apply(_formata_cnpj_mascara)
    st.dataframe(df_exib, use_container_width=True, hide_index=True)


@st.fragment
def pagina_dashboard_paralegal():
    import plotly.express as px
    st.empty()

    df = le_planilha_google(GOOGLE_SHEET_URL, SHEET_EMPRESAS)
    if df is None:
        return

    if "Situação" not in df.columns:
        st.error("Coluna 'Situação' não encontrada.")
        return

    df_ativas = df[df["Situação"].astype(str).str.upper() == "ATIVA"].copy()
    total_ativas = df_ativas.shape[0]

    st.markdown("<h2 style='color:#1d3f77;'>Dashboard — Departamento Paralegal</h2>",
                unsafe_allow_html=True)
    st.markdown(f"<p style='font-size:18px;'><b>Total de empresas ativas:</b> {total_ativas}</p>",
                unsafe_allow_html=True)
    st.divider()

    # ── controle de modal: guarda último clique de cada gráfico
    for k in ["ult_estado", "ult_municipio", "ult_cnd"]:
        if k not in st.session_state:
            st.session_state[k] = None

    modal_abrir = None   # apenas UM modal por execução

    # ── EMPRESAS POR ESTADO ──────────────────────────────────────────────────
    st.markdown("### Empresas por Estado")
    st.caption("Clique em uma barra para ver as empresas")

    if "Estado" in df_ativas.columns:
        df_est = df_ativas["Estado"].fillna("N/I").astype(str).str.strip()
        df_est_count = df_est.value_counts().reset_index()
        df_est_count.columns = ["Estado_orig", "Quantidade"]
        df_est_count["Estado"] = df_est_count["Estado_orig"].apply(_formata_estado)

        df_est_count["Qtd_display"] = df_est_count["Quantidade"].apply(
            lambda x: max(x, df_est_count["Quantidade"].max() * 0.03)
        )

        fig_est = px.bar(
            df_est_count, x="Estado", y="Qtd_display",
            color="Qtd_display",
            color_continuous_scale=[[0, "#4a90d9"], [1, "#1d3f77"]],
            text="Quantidade",
            custom_data=["Estado_orig", "Quantidade"],
        )
        fig_est.update_traces(
            textposition="outside",
            hovertemplate="<b>%{customdata[0]}</b><br>Qtd: %{customdata[1]}<extra></extra>",
        )
        fig_est.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            coloraxis_showscale=False,
            xaxis=dict(title="", tickfont=dict(size=13), showgrid=False),
            yaxis=dict(title="", showgrid=False, zeroline=False),
            margin=dict(t=30, b=20, l=10, r=10),
            height=350,
            clickmode="event+select",
            bargap=0.1,
            bargroupgap=0.0,
        )

        ev_est = st.plotly_chart(fig_est, use_container_width=True,
                                  on_select="rerun", key="chart_estado")

        if ev_est and ev_est.selection and ev_est.selection.points:
            pt   = ev_est.selection.points[0]
            sel_raw = (pt.get("customdata") or [None])[0]
            if sel_raw is None:
                sel_raw = _estado_original(pt.get("x", ""))
            if sel_raw and sel_raw != st.session_state["ult_estado"]:
                st.session_state["ult_estado"] = sel_raw
                df_fil = df_ativas[
                    df_ativas["Estado"].fillna("N/I").astype(str).str.strip() == sel_raw
                ]
                modal_abrir = (f"Estado: {sel_raw}", df_fil,
                               ["Razão Social", "CNPJ"])
    else:
        st.warning("Coluna 'Estado' não encontrada.")

    st.divider()

    # ── EMPRESAS POR MUNICÍPIO ───────────────────────────────────────────────
    st.markdown("### Empresas por Município")
    st.caption("Clique em uma barra para ver as empresas")

    if "Município" in df_ativas.columns:
        df_mun_count = (df_ativas["Município"].fillna("N/I").astype(str).str.strip()
                        .value_counts().reset_index())
        df_mun_count.columns = ["Município", "Quantidade"]
        df_mun_count = df_mun_count.sort_values("Quantidade", ascending=True)

        altura_mun = max(400, len(df_mun_count) * 28)

        fig_mun = px.bar(
            df_mun_count, x="Quantidade", y="Município",
            orientation="h",
            color="Quantidade",
            color_continuous_scale=[[0, "#4a90d9"], [1, "#1d3f77"]],
            text="Quantidade",
        )
        fig_mun.update_traces(
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Qtd: %{x}<extra></extra>",
        )
        fig_mun.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            coloraxis_showscale=False,
            xaxis=dict(title="", showgrid=False, zeroline=False),
            yaxis=dict(title="", tickfont=dict(size=12), showgrid=False),
            margin=dict(t=20, r=60, b=20, l=10),
            height=altura_mun,
            clickmode="event+select",
        )

        ev_mun = st.plotly_chart(fig_mun, use_container_width=True,
                                  on_select="rerun", key="chart_municipio")

        if modal_abrir is None and ev_mun and ev_mun.selection and ev_mun.selection.points:
            sel_mun = ev_mun.selection.points[0].get("y")
            if sel_mun and sel_mun != st.session_state["ult_municipio"]:
                st.session_state["ult_municipio"] = sel_mun
                df_fil = df_ativas[
                    df_ativas["Município"].fillna("N/I").astype(str).str.strip() == sel_mun
                ]
                modal_abrir = (f"Município: {sel_mun}", df_fil,
                               ["Razão Social", "CNPJ"])
    else:
        st.warning("Coluna 'Município' não encontrada.")

    st.divider()

    # ── CND MUNICIPAL — SITUAÇÃO ─────────────────────────────────────────────
    st.markdown("### CND Municipal — Situação")
    st.caption("Clique em uma fatia para ver as empresas")

    if "SITUAÇÃO CND MUNICIPAL" in df_ativas.columns:
        df_cnd = df_ativas.copy()
        df_cnd["SIT"] = (df_cnd["SITUAÇÃO CND MUNICIPAL"]
                         .fillna("").astype(str).str.strip())
        df_cnd["CND_LABEL"] = df_cnd["SIT"].apply(
            lambda x: "Outros Municípios" if x == "" or x.upper() == "NAN" else x
        )

        df_cnd_count = df_cnd["CND_LABEL"].value_counts().reset_index()
        df_cnd_count.columns = ["Situação", "Quantidade"]

        color_map = {
            "NEGATIVA":                     "#27ae60",
            "POSITIVA":                     "#e74c3c",
            "POSITIVA COM EFEITO NEGATIVA": "#f39c12",
            "Outros Municípios":            "#bdc3c7",
        }

        # Ordena para barras menores no topo
        df_cnd_count = df_cnd_count.sort_values("Quantidade", ascending=True)

        fig_cnd = px.bar(
            df_cnd_count, x="Quantidade", y="Situação",
            orientation="h",
            text="Quantidade",
            color="Situação",
            color_discrete_map=color_map,
        )
        fig_cnd.update_traces(
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Qtd: %{x}<extra></extra>",
        )
        fig_cnd.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            showlegend=False,
            xaxis=dict(title="", showgrid=False, zeroline=False),
            yaxis=dict(title="", tickfont=dict(size=13), showgrid=False),
            margin=dict(t=20, r=80, b=20, l=10),
            height=max(200, len(df_cnd_count) * 60),
            clickmode="event+select",
        )

        ev_cnd = st.plotly_chart(fig_cnd, use_container_width=True,
                                  on_select="rerun", key="chart_cnd")

        if modal_abrir is None and ev_cnd and ev_cnd.selection and ev_cnd.selection.points:
            pt_cnd  = ev_cnd.selection.points[0]
            sel_cnd = pt_cnd.get("y")
            if sel_cnd and sel_cnd != st.session_state["ult_cnd"]:
                st.session_state["ult_cnd"] = sel_cnd
                df_fil = df_cnd[df_cnd["CND_LABEL"] == sel_cnd]
                modal_abrir = (f"Situação CND: {sel_cnd}", df_fil,
                               ["Razão Social", "CNPJ", "Município",
                                "SITUAÇÃO CND MUNICIPAL"])
    else:
        st.info("Coluna 'SITUAÇÃO CND MUNICIPAL' não encontrada.")

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
    """Converte todas as colunas para string — evita erro Arrow."""
    df = df.copy()
    for col in df.columns:
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


@st.dialog("SEFAZ COMPARAÇÃO — Divergências")
def _modal_sefaz_comparacao(df_show):
    st.markdown(f"**{df_show.shape[0]} empresa(s) com divergência**")
    st.dataframe(df_show.reset_index(drop=True), use_container_width=True, hide_index=True)


@st.fragment
def pagina_sefaz_comparacao():
    import plotly.graph_objects as go
    st.empty()

    # ── carrega aba SEFAZ ─────────────────────────────────────────────────────
    try:
        resp = requests.get(GOOGLE_SHEET_URL)
        resp.raise_for_status()
        df_sefaz = pd.read_excel(
            BytesIO(resp.content),
            sheet_name=SHEET_SEFAZ,
            engine="openpyxl",
            header=0,
        )
    except Exception as e:
        st.error(f"Erro ao ler aba SEFAZ: {e}")
        return

    df_sefaz.columns = df_sefaz.columns.str.strip()

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

    # ── identifica colunas por posição ───────────────────────────────────────
    # A=0, D=3, T=19, W=22
    cols = df_sefaz.columns.tolist()

    def _col(idx):
        return cols[idx] if idx < len(cols) else None

    nome_cod_a  = _col(0)   # A — CÓDIGO
    nome_said_d = _col(3)   # D — SAÍDAS
    nome_cod_t  = _col(19)  # T — CÓDIGO
    nome_said_w = _col(22)  # W — SAÍDAS

    if not all([nome_cod_a, nome_said_d, nome_cod_t, nome_said_w]):
        st.error("Colunas A, D, T ou W não encontradas na aba SEFAZ.")
        return

    # ── normaliza e converte ──────────────────────────────────────────────────
    # ── normaliza códigos — remove .0 ────────────────────────────────────────
    def _limpa_codigo(val):
        s = str(val).strip()
        if s.endswith(".0"):
            s = s[:-2]
        return s.upper().replace("NAN", "").strip()

    df_sefaz[nome_cod_a] = df_sefaz[nome_cod_a].apply(_limpa_codigo)
    df_sefaz[nome_cod_t] = df_sefaz[nome_cod_t].apply(_limpa_codigo)
    df_sefaz[nome_said_d] = df_sefaz[nome_said_d].apply(_limpa_numero)
    df_sefaz[nome_said_w] = df_sefaz[nome_said_w].apply(_limpa_numero)

    # ── filtra linhas válidas de cada lado ────────────────────────────────────
    df_lado_a = df_sefaz[df_sefaz[nome_cod_a] != ""][
        [nome_cod_a, nome_said_d]
    ].copy()
    df_lado_a.columns = ["Código", "Saídas_A"]

    df_lado_t = df_sefaz[df_sefaz[nome_cod_t] != ""][
        [nome_cod_t, nome_said_w]
    ].copy()
    df_lado_t.columns = ["Código", "Saídas_T"]

    # ── agrupa por código (soma caso haja duplicatas) ─────────────────────────
    df_lado_a = df_lado_a.groupby("Código", as_index=False)["Saídas_A"].sum()
    df_lado_t = df_lado_t.groupby("Código", as_index=False)["Saídas_T"].sum()

    # ── junta pelos códigos ───────────────────────────────────────────────────
    df_merge = pd.merge(df_lado_a, df_lado_t, on="Código", how="outer").fillna(0)

    # ── compara ───────────────────────────────────────────────────────────────
    df_merge["Diferença"] = df_merge["Saídas_A"] - df_merge["Saídas_T"]
    df_dif  = df_merge[df_merge["Diferença"] != 0].copy()
    df_ok   = df_merge[df_merge["Diferença"] == 0].copy()

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
                "Saídas Inicial":  _fmt_valor(row["Saídas_A"]),
                "Saídas Final":    _fmt_valor(row["Saídas_T"]),
                "Diferença":       _fmt_valor(row["Diferença"]),
            })
        return pd.DataFrame(rows)

    df_result = _enriquece(df_dif)

    total_empresas = len(df_merge)
    total_dif      = len(df_result)
    total_ok       = len(df_ok)

    # ── cabeçalho ─────────────────────────────────────────────────────────────
    st.markdown("<h2>SEFAZ COMPARAÇÃO</h2>", unsafe_allow_html=True)
    st.markdown(
        f"<p style='text-align:right; font-size:20px;'>"
        f"<b>Com divergência:</b> {total_dif} &nbsp;|&nbsp; "
        f"<b>Sem divergência:</b> {total_ok} &nbsp;|&nbsp; "
        f"<b>Total:</b> {total_empresas}</p>",
        unsafe_allow_html=True,
    )

    # ── donut ─────────────────────────────────────────────────────────────────
    if "sefaz_comp_key" not in st.session_state:
        st.session_state["sefaz_comp_key"] = 0

    pct_dif = round(total_dif / total_empresas * 100) if total_empresas else 0
    pct_ok  = round(total_ok  / total_empresas * 100) if total_empresas else 0

    fig = go.Figure(data=[go.Pie(
        labels=["Com Divergência", "Sem Divergência"],
        values=[int(total_dif), int(total_ok)],
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
            key=f"chart_sefaz_comp_{st.session_state['sefaz_comp_key']}",
        )

    col_l, col_r = st.columns(2)
    with col_l:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#fdedec; "
            f"border-radius:8px; border-left:4px solid #c0392b;'>"
            f"<span style='font-size:22px; font-weight:700; color:#c0392b;'>{total_dif}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Com Divergência ({pct_dif}%)</span></div>",
            unsafe_allow_html=True,
        )
    with col_r:
        st.markdown(
            f"<div style='text-align:center; padding:8px; background:#eafaf1; "
            f"border-radius:8px; border-left:4px solid #27ae60;'>"
            f"<span style='font-size:22px; font-weight:700; color:#27ae60;'>{total_ok}</span><br>"
            f"<span style='font-size:13px; color:#555;'>Sem Divergência ({pct_ok}%)</span></div>",
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("Ver empresas com divergência", use_container_width=True,
                 key="btn_sefaz_comp_dif"):
        if not df_result.empty:
            _modal_sefaz_comparacao(df_result)
        else:
            st.info("Nenhuma divergência encontrada!")

    st.divider()

    # ── lista ─────────────────────────────────────────────────────────────────
    st.markdown("### Lista de Divergências", unsafe_allow_html=True)

    if not df_result.empty:
        df_exib = _sanitiza_df(df_result)
        exibe_aggrid(df_exib, height=400, grid_key="grid_sefaz_comp")

        output = BytesIO()
        df_result.to_excel(output, index=False)
        st.download_button(
            "Baixar Excel", data=output.getvalue(),
            file_name="sefaz_comparacao.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.success("Nenhuma divergência encontrada entre as colunas!")

# ============================================================================
# ROTEAMENTO
# ============================================================================

with st.session_state.main_container.container():
    if pagina == "DASHBOARD":
        pagina_dashboard_paralegal()
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
    elif pagina == "SEFAZ COMPARAÇÃO":
        pagina_sefaz_comparacao()
    elif pagina == "CERTIFICADOS":
        pagina_certificados()
    elif pagina == "ENDEREÇO DE EMAIL":
        pagina_emails_cnpj()
    elif pagina == "MENSAGENS DE EMAIL":
        pagina_mensagens_email()