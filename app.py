import streamlit as st
import pandas as pd
import plotly.express as px
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime, date, time
import io

# Tentar importar gspread e google.oauth2
HAS_GSPREAD = False
try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_GSPREAD = True
except ImportError:
    HAS_GSPREAD = False

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA & TEMA
# ==============================================================================
st.set_page_config(
    page_title="RECOVERY — Gestão Financeira",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização CSS personalizada (Visual SaaS Moderno)
st.markdown("""
    <style>
    .main-header {
        font-size: 2.2rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 0px;
    }
    .sub-header {
        font-size: 1rem;
        color: #64748B;
        margin-bottom: 20px;
    }
    .preview-card {
        background-color: #F8FAFC;
        border-left: 4px solid #2563EB;
        padding: 15px;
        border-radius: 6px;
        margin-top: 10px;
    }
    .status-card {
        background-color: #F1F5F9;
        border: 1px solid #CBD5E1;
        padding: 12px 18px;
        border-radius: 8px;
        margin-bottom: 20px;
    }
    .user-badge {
        background-color: #E2E8F0;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        color: #1E293B;
        display: inline-block;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

SPREADSHEET_ID = "1aygRSlsXbsafrF-osiOq9S_BhNaIIDfKU4CbtVoemBA"

# ==============================================================================
# 2. INTEGRAÇÃO ROBUSTA COM GOOGLE SHEETS E DIAGNÓSTICO
# ==============================================================================
def get_gspread_client():
    """Obtém o cliente gspread autenticado usando os secrets do Streamlit."""
    if not HAS_GSPREAD:
        return None, "Biblioteca 'gspread' não instalada no requirements.txt."
    
    if "gcp_service_account" not in st.secrets:
        return None, "Configuração [gcp_service_account] não encontrada em Secrets do Streamlit."

    try:
        creds_dict = dict(st.secrets["gcp_service_account"])
        # Garante tratamento para quebras de linha na private_key
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        client = gspread.authorize(credentials)
        return client, None
    except Exception as e:
        return None, f"Erro de Autenticação nas Credenciais: {str(e)}"

def testar_conexao_google_sheets():
    """Testa a conexão e o acesso de escrita na planilha para exibir o diagnóstico no Dashboard."""
    client, err_msg = get_gspread_client()
    if err_msg:
        return False, err_msg, None
    
    try:
        sh = client.open_by_key(SPREADSHEET_ID)
        email_bot = st.secrets["gcp_service_account"].get("client_email", "Email N/A")
        return True, f"Conectado com SUCESSO à planilha '{sh.title}'", email_bot
    except Exception as e:
        return False, f"Erro ao acessar a planilha: {str(e)}", None

def sync_to_google_sheets(df_lancamentos):
    """Envia o DataFrame completo de Lançamentos para a aba 'Lançamentos' do Google Sheets."""
    client, err_msg = get_gspread_client()
    if err_msg:
        st.error(f"❌ Falha de Conexão com Google Sheets: {err_msg}")
        return False

    try:
        sh = client.open_by_key(SPREADSHEET_ID)
        
        try:
            ws = sh.worksheet("Lançamentos")
        except Exception:
            ws = sh.add_worksheet(title="Lançamentos", rows="1000", cols="30")

        # Limpa o conteúdo existente para regravar o histórico atualizado
        ws.clear()

        # Trata nulos e converte para lista de strings
        df_clean = df_lancamentos.copy().fillna("")
        for col in df_clean.columns:
            df_clean[col] = df_clean[col].astype(str)

        valores = [df_clean.columns.tolist()] + df_clean.values.tolist()
        ws.update("A1", valores)
        return True
    except Exception as e:
        st.error(f"❌ Erro ao gravar dados na planilha Google Sheets: {str(e)}")
        return False

# ==============================================================================
# 3. CÁLCULOS FINANCEIROS (CENTRALIZADOS)
# ==============================================================================
def to_decimal(val) -> Decimal:
    if val is None or val == "":
        return Decimal("0.00")
    if isinstance(val, Decimal):
        return val
    try:
        if isinstance(val, str):
            clean_str = val.replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
            return Decimal(clean_str)
        return Decimal(str(val))
    except Exception:
        return Decimal("0.00")

def quantize_money(amount: Decimal) -> Decimal:
    return amount.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

def calcular_lancamento(
    valor_bruto,
    taxa_pagamento_pct,
    taxa_pagamento_fixa=0,
    material_clinica=0,
    material_dra=0,
    material_adicional=0,
    aliquota_imposto_pct=0,
    custo_nota_fiscal=0,
    percentual_repasse_dra=50
):
    v_bruto = to_decimal(valor_bruto)
    taxa_pct = to_decimal(taxa_pagamento_pct)
    taxa_fixa = to_decimal(taxa_pagamento_fixa)
    mat_clinica = to_decimal(material_clinica)
    mat_dra = to_decimal(material_dra)
    mat_add = to_decimal(material_adicional)
    imposto_pct = to_decimal(aliquota_imposto_pct)
    c_nota = to_decimal(custo_nota_fiscal)
    repasse_pct = to_decimal(percentual_repasse_dra)

    # Taxa de Pagamento
    valor_taxa_pag = quantize_money((v_bruto * (taxa_pct / Decimal("100"))) + taxa_fixa)
    valor_apos_taxa = quantize_money(v_bruto - valor_taxa_pag)

    # Impostos (caso aplicável por alíquota)
    valor_imposto = quantize_money(v_bruto * (imposto_pct / Decimal("100")))
    cost_nf_total = c_nota if c_nota > 0 else valor_imposto

    # Material Total Clínica
    mat_clinica_total = quantize_money(mat_clinica + mat_add)

    # Lucro Líquido
    lucro_liquido = quantize_money(
        valor_apos_taxa - mat_clinica_total - mat_dra - cost_nf_total
    )

    # Divisão/Repasse
    repasse_dra_base = quantize_money(lucro_liquido * (repasse_pct / Decimal("100")))
    valor_final_dra = quantize_money(repasse_dra_base + mat_dra)
    valor_repassado_clinica = quantize_money(lucro_liquido - repasse_dra_base)
    valor_final_clinica_com_material = quantize_money(valor_repassado_clinica + mat_clinica_total)

    return {
        "valor_bruto": float(v_bruto),
        "taxa_pagamento_pct": float(taxa_pct),
        "valor_taxa_pagamento": float(valor_taxa_pag),
        "valor_apos_taxa": float(valor_apos_taxa),
        "material_clinica": float(mat_clinica),
        "material_adicional": float(mat_add),
        "material_clinica_total": float(mat_clinica_total),
        "material_dra": float(mat_dra),
        "aliquota_imposto_pct": float(imposto_pct),
        "custo_nota_fiscal": float(cost_nf_total),
        "lucro_liquido": float(lucro_liquido),
        "valor_repassado_clinica": float(valor_repassado_clinica),
        "valor_final_dra": float(valor_final_dra),
        "valor_final_clinica": float(valor_repassado_clinica),
        "valor_final_clinica_com_material": float(valor_final_clinica_com_material),
    }

# ==============================================================================
# 4. GESTÃO DE USUÁRIOS E AUTENTICAÇÃO HIERÁRQUICA
# ==============================================================================
USUARIOS = {
    "proprietario": {
        "senha": "123",
        "nome": "Dr. Proprietário",
        "perfil": "Proprietário",
        "acesso_financeiro": True,
        "pode_excluir_alterar": True,
        "pode_editar_parametros": True,
    },
    "pleno": {
        "senha": "123",
        "nome": "Auxiliar Pleno",
        "perfil": "Auxiliar Financeiro Pleno",
        "acesso_financeiro": True,
        "pode_excluir_alterar": True,
        "pode_editar_parametros": True,
    },
    "junior": {
        "senha": "123",
        "nome": "Auxiliar Júnior",
        "perfil": "Auxiliar Financeiro Júnior",
        "acesso_financeiro": True,
        "pode_excluir_alterar": False,
        "pode_editar_parametros": False,
    },
    "secretaria": {
        "senha": "123",
        "nome": "Secretária Recovery",
        "perfil": "Secretária",
        "acesso_financeiro": False,
        "pode_excluir_alterar": False,
        "pode_editar_parametros": False,
    }
}

def check_login():
    if "usuario_logado" not in st.session_state:
        st.session_state.usuario_logado = None

    if st.session_state.usuario_logado is None:
        st.title("🔐 RECOVERY — Portal de Acesso")
        st.caption("Insira suas credenciais para acessar a plataforma de gestão.")
        
        with st.form("form_login"):
            usuario_input = st.text_input("Usuário").strip().lower()
            senha_input = st.text_input("Senha", type="password")
            btn_entrar = st.form_submit_button("Entrar", use_container_width=True)

            if btn_entrar:
                if usuario_input in USUARIOS and USUARIOS[usuario_input]["senha"] == senha_input:
                    st.session_state.usuario_logado = USUARIOS[usuario_input]
                    st.success(f"Bem-vindo(a), {st.session_state.usuario_logado['nome']}!")
                    st.rerun()
                else:
                    st.error("Usuário ou senha incorretos.")
        return False
    return True

# ==============================================================================
# 5. INICIALIZAÇÃO DE ESTADO E DADOS
# ==============================================================================
def init_session_state():
    if "param_procedimentos" not in st.session_state:
        st.session_state.param_procedimentos = pd.DataFrame([
            {"Procedimento": "BIOIMPEDÂNCIA", "Custo Material": 150.0},
            {"Procedimento": "PROTOCOLO DE SOROTERAPIA - MÉDICA", "Custo Material": 100.0},
            {"Procedimento": "RETORNO BOTOX - ESTÉTICA", "Custo Material": 200.0},
            {"Procedimento": "CONSULTA - ESTÉTICA", "Custo Material": 0.0},
        ])

    if "param_taxas" not in st.session_state:
        st.session_state.param_taxas = pd.DataFrame([
            {"Forma de Pagamento": "PIX", "Taxa (%)": 0.0},
            {"Forma de Pagamento": "Boleto", "Taxa (%)": 0.0},
            {"Forma de Pagamento": "Visa Débito", "Taxa (%)": 0.79},
            {"Forma de Pagamento": "Visa Credito 1x", "Taxa (%)": 2.79},
            {"Forma de Pagamento": "Visa Credito 2x", "Taxa (%)": 4.08},
            {"Forma de Pagamento": "Visa Credito 12x", "Taxa (%)": 9.56},
            {"Forma de Pagamento": "Master Credito 1x", "Taxa (%)": 2.79},
        ])

    if "param_impostos" not in st.session_state:
        st.session_state.param_impostos = pd.DataFrame([
            {"Tipo": "IVA", "Taxa (%)": 10.0},
            {"Tipo": "ICBS", "Taxa (%)": 20.0},
            {"Tipo": "CB", "Taxa (%)": 30.0},
        ])

    if "param_origens" not in st.session_state:
        st.session_state.param_origens = ["Particular", "Instagram", "Indicação", "Google", "Médico", "Convênio", "Outros"]

    if "param_profissionais" not in st.session_state:
        st.session_state.param_profissionais = ["DRA. DENISSE", "DR. GABRIEL", "RECOVERY"]

    if "agendamentos" not in st.session_state:
        st.session_state.agendamentos = pd.DataFrame([
            {
                "ID": "AGD-001",
                "Data": "2026-08-25",
                "Horário": "09:00",
                "Paciente": "Ana Souza",
                "Procedimento": "BIOIMPEDÂNCIA",
                "Profissional": "DRA. DENISSE",
                "Telefone": "(47) 99999-8888",
                "Status": "Confirmado",
                "Observações": "Primeira consulta."
            }
        ])

    if "lancamentos" not in st.session_state:
        st.session_state.lancamentos = pd.DataFrame([
            {
                "ID": "LAN-001",
                "DATA ATENDIMENTO": "2026-08-18",
                "PACIENTE": "João Silva",
                "PROFISSIONAL": "DRA. DENISSE",
                "ORIGEM": "Instagram",
                "PROCEDIMENTO": "RETORNO BOTOX - ESTÉTICA",
                "Valor Bruto (R$)": 5000.0,
                "Forma de Pagamento": "Visa Credito 12x",
                "Taxa de Pagamento (%)": 9.56,
                "Valor da Taxa de Pagamento (R$)": 478.0,
                "Valor após Taxa de Pagamento (R$)": 4522.0,
                "Material da clínica (R$)": 200.0,
                "Custo Adicional Material (R$)": 50.0,
                "Material da Dra. Denise (R$)": 0.0,
                "Tipo de Imposto": "IVA",
                "Custo Nota Fiscal (R$)": 500.0,
                "Lucro Líquido (R$)": 3772.0,
                "Valor Repasse Clínica (R$)": 1886.0,
                "Valor Final Dra. Denise (R$)": 1886.0,
                "Valor Final Clínica/Com Material (R$)": 2136.0,
                "Observações": "Utilizado kit extra cirúrgico.",
                "CRIADO EM": "2026-08-18 14:30:00",
                "USUARIO CRIACAO": "Dr. Proprietário",
                "STATUS": "ATIVO",
                "EXCLUIDO EM": "",
                "USUARIO EXCLUSAO": ""
            }
        ])

def fmt_brl(val) -> str:
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ==============================================================================
# 6. EXECUÇÃO PRINCIPAL E NAVEGAÇÃO
# ==============================================================================
if check_login():
    init_session_state()
    user = st.session_state.usuario_logado

    # Sidebar com informações do perfil
    st.sidebar.title("RECOVERY")
    st.sidebar.caption("Gestão Financeira e Agendamentos")
    st.sidebar.markdown(f'<div class="user-badge">👤 {user["nome"]}<br><small>{user["perfil"]}</small></div>', unsafe_allow_html=True)
    
    if st.sidebar.button("🚪 Sair / Logout"):
        st.session_state.usuario_logado = None
        st.rerun()

    # Definir opções de menu conforme permissões do perfil
    if user["acesso_financeiro"]:
        menu_options = ["📊 Dashboard", "📅 Agendamentos", "➕ Novo Lançamento", "📋 Lançamentos", "⚙️ Parâmetros", "📈 Relatórios"]
    else:
        menu_options = ["📅 Agendamentos"] # Secretária acessa apenas Agendamentos

    menu = st.sidebar.radio("Navegação", menu_options)

    # ==========================================================================
    # MODULO: DASHBOARD (Com Card de Diagnóstico das Conexões)
    # ==========================================================================
    if menu == "📊 Dashboard":
        st.markdown('<div class="main-header">RECOVERY — Dashboard Financeiro</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Visão consolidada dos lançamentos ativos e indicadores da clínica.</div>', unsafe_allow_html=True)

        # CARD DE DIAGNÓSTICO DE CONEXÃO
        with st.expander("📡 Status de Conexão e Integrações (Google Drive / Google Sheets)", expanded=False):
            st_ok, msg_conexao, email_bot = testar_conexao_google_sheets()
            
            c_diag1, c_diag2 = st.columns(2)
            with c_diag1:
                st.write("**Biblioteca `gspread`:**", "🟢 Instalada" if HAS_GSPREAD else "🔴 Não instalada")
                st.write("**Secrets `gcp_service_account`:**", "🟢 Configurado" if "gcp_service_account" in st.secrets else "🔴 Ausente")
            
            with c_diag2:
                if st_ok:
                    st.success(f"🟢 **Google Sheets:** {msg_conexao}")
                    st.caption(f"E-mail da Service Account conectada: `{email_bot}`")
                else:
                    st.error(f"🔴 **Google Sheets:** {msg_conexao}")
                    st.caption("Verifique se você compartilhou a planilha como 'Editor' com o e-mail: `recoveryagora@recorybianca.iam.gserviceaccount.com`")

        # FILTROS
        df_lan_raw = st.session_state.lancamentos.copy()
        df_lan = df_lan_raw[df_lan_raw["STATUS"] == "ATIVO"] if "STATUS" in df_lan_raw.columns else df_lan_raw

        with st.expander("🔍 Filtros de Consulta", expanded=True):
            f_col1, f_col2, f_col3, f_col4 = st.columns(4)
            with f_col1:
                sel_prof = st.selectbox("Profissional", ["Todos"] + st.session_state.param_profissionais)
            with f_col2:
                sel_proc = st.selectbox("Procedimento", ["Todos"] + list(st.session_state.param_procedimentos["Procedimento"].unique()))
            with f_col3:
                sel_orig = st.selectbox("Origem", ["Todas"] + st.session_state.param_origens)
            with f_col4:
                sel_pag = st.selectbox("Forma de Pagamento", ["Todas"] + list(st.session_state.param_taxas["Forma de Pagamento"].unique()))

        if not df_lan.empty:
            if sel_prof != "Todos": df_lan = df_lan[df_lan["PROFISSIONAL"] == sel_prof]
            if sel_proc != "Todos": df_lan = df_lan[df_lan["PROCEDIMENTO"] == sel_proc]
            if sel_orig != "Todas": df_lan = df_lan[df_lan["ORIGEM"] == sel_orig]
            if sel_pag != "Todas": df_lan = df_lan[df_lan["Forma de Pagamento"] == sel_pag]

        v_bruto_tot = float(pd.to_numeric(df_lan["Valor Bruto (R$)"].values, errors='coerce').sum()) if not df_lan.empty else 0.0
        v_taxas_tot = float(pd.to_numeric(df_lan["Valor da Taxa de Pagamento (R$)"].values, errors='coerce').sum()) if not df_lan.empty else 0.0
        v_mat_clin = float(pd.to_numeric(df_lan["Material da clínica (R$)"].values, errors='coerce').sum()) if not df_lan.empty else 0.0
        v_mat_add = float(pd.to_numeric(df_lan["Custo Adicional Material (R$)"].values, errors='coerce').sum()) if not df_lan.empty else 0.0
        v_mat_dra = float(pd.to_numeric(df_lan["Material da Dra. Denise (R$)"].values, errors='coerce').sum()) if not df_lan.empty else 0.0
        v_impostos = float(pd.to_numeric(df_lan["Custo Nota Fiscal (R$)"].values, errors='coerce').sum()) if not df_lan.empty else 0.0
        v_lucro_liq = float(pd.to_numeric(df_lan["Lucro Líquido (R$)"].values, errors='coerce').sum()) if not df_lan.empty else 0.0
        v_final_dra = float(pd.to_numeric(df_lan["Valor Final Dra. Denise (R$)"].values, errors='coerce').sum()) if not df_lan.empty else 0.0
        v_final_clin = float(pd.to_numeric(df_lan["Valor Final Clínica/Com Material (R$)"].values, errors='coerce').sum()) if not df_lan.empty else 0.0

        kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
        kpi1.metric("Faturamento Bruto", fmt_brl(v_bruto_tot))
        kpi2.metric("Taxas Pagamento", fmt_brl(v_taxas_tot))
        kpi3.metric("Impostos/Notas", fmt_brl(v_impostos))
        kpi4.metric("Materiais Totais", fmt_brl(v_mat_clin + v_mat_add + v_mat_dra))
        kpi5.metric("Lucro Líquido", fmt_brl(v_lucro_liq))

        st.markdown("---")
        r_col1, r_col2 = st.columns(2)
        r_col1.info(f"### Repasse Profissional: **{fmt_brl(v_final_dra)}**")
        r_col2.success(f"### Total Clínica (c/ Material): **{fmt_brl(v_final_clin)}**")

        if not df_lan.empty:
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.subheader("Distribuição do Resultado")
                df_dist = pd.DataFrame({
                    "Categoria": ["Repasse Dra.", "Clínica (c/ Mat)", "Impostos", "Taxas Cartão"],
                    "Valor": [v_final_dra, v_final_clin, v_impostos, v_taxas_tot]
                })
                fig_pie = px.pie(df_dist, names="Categoria", values="Valor", hole=0.4)
                st.plotly_chart(fig_pie, use_container_width=True)

            with g_col2:
                st.subheader("Faturamento por Procedimento")
                df_proc_f = df_lan.groupby("PROCEDIMENTO")["Valor Bruto (R$)"].sum().reset_index()
                fig_bar = px.bar(df_proc_f, x="PROCEDIMENTO", y="Valor Bruto (R$)", text_auto=True)
                st.plotly_chart(fig_bar, use_container_width=True)

    # ==========================================================================
    # MODULO: AGENDAMENTOS
    # ==========================================================================
    elif menu == "📅 Agendamentos":
        st.markdown('<div class="main-header">RECOVERY — Agendamento de Atendimentos</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Gestão de horários, pacientes e procedimentos da clínica.</div>', unsafe_allow_html=True)

        tab_lista, tab_novo = st.tabs(["📋 Agendamentos Marcados", "➕ Novo Agendamento"])

        with tab_lista:
            df_agd = st.session_state.agendamentos
            if not df_agd.empty:
                st.dataframe(df_agd, use_container_width=True, hide_index=True)
            else:
                st.info("Nenhum agendamento cadastrado.")

        with tab_novo:
            with st.form("form_novo_agendamento", clear_on_submit=True):
                col1, col2, col3 = st.columns(3)
                with col1:
                    agd_data = st.date_input("Data do Atendimento", value=date.today())
                    agd_hora = st.time_input("Horário", value=time(9, 0))
                    agd_paciente = st.text_input("Nome do Paciente *")
                with col2:
                    procedimentos_lista = list(st.session_state.param_procedimentos["Procedimento"].unique())
                    agd_procedimento = st.selectbox("Procedimento *", procedimentos_lista)
                    agd_profissional = st.selectbox("Profissional", st.session_state.param_profissionais)
                with col3:
                    agd_telefone = st.text_input("Telefone / WhatsApp", placeholder="(00) 00000-0000")
                    agd_status = st.selectbox("Status Inicial", ["Agendado", "Confirmado", "Realizado", "Cancelado"])

                agd_obs = st.text_area("Observações do Agendamento", placeholder="Anotações gerais do paciente...")

                if st.form_submit_button("💾 Salvar Agendamento", use_container_width=True):
                    if not agd_paciente:
                        st.error("Por favor, preencha o nome do paciente.")
                    else:
                        novo_agd_id = f"AGD-{len(st.session_state.agendamentos) + 1:03d}"
                        novo_agd_rec = {
                            "ID": novo_agd_id,
                            "Data": str(agd_data),
                            "Horário": agd_hora.strftime("%H:%M"),
                            "Paciente": agd_paciente,
                            "Procedimento": agd_procedimento,
                            "Profissional": agd_profissional,
                            "Telefone": agd_telefone,
                            "Status": agd_status,
                            "Observações": agd_obs
                        }
                        st.session_state.agendamentos = pd.concat([
                            st.session_state.agendamentos, pd.DataFrame([novo_agd_rec])
                        ], ignore_index=True)
                        st.success(f"Agendamento {novo_agd_id} cadastrado com sucesso!")
                        st.rerun()

    # ==========================================================================
    # MODULO: NOVO LANÇAMENTO (Com Sincronização e Feedback)
    # ==========================================================================
    elif menu == "➕ Novo Lançamento":
        st.markdown('<div class="main-header">RECOVERY — Novo Lançamento Financeiro</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Registre lançamentos com prévia automatizada e sincronização na planilha.</div>', unsafe_allow_html=True)

        with st.form("form_novo_lancamento", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                data_atend = st.date_input("Data do Atendimento", value=date.today())
                paciente = st.text_input("Nome do Paciente *", placeholder="Ex: Maria Oliveira")
                profissional = st.selectbox("Profissional", st.session_state.param_profissionais)
            
            with col2:
                origem = st.selectbox("Origem do Paciente", st.session_state.param_origens)
                procedimentos_lista = list(st.session_state.param_procedimentos["Procedimento"].unique())
                procedimento_sel = st.selectbox("Procedimento *", procedimentos_lista)
                valor_bruto_inp = st.number_input("Valor Bruto (R$) *", min_value=0.0, step=50.0, value=500.0)

            with col3:
                pagamentos_lista = list(st.session_state.param_taxas["Forma de Pagamento"].unique())
                forma_pag_sel = st.selectbox("Forma de Pagamento *", pagamentos_lista)
                
                mat_default = st.session_state.param_procedimentos.loc[
                    st.session_state.param_procedimentos["Procedimento"] == procedimento_sel, "Custo Material"
                ].values
                mat_clinica_def = float(mat_default[0]) if len(mat_default) > 0 and pd.notnull(mat_default[0]) else 0.0

                material_clinica_inp = st.number_input("Material Padrão Clínica (R$)", min_value=0.0, value=mat_clinica_def)
                material_adicional_inp = st.number_input("Custo Adicional Material (R$)", min_value=0.0, value=0.0, help="Uso eventual de kits extras.")
                material_dra_inp = st.number_input("Material da Dra. Denise (R$)", min_value=0.0, value=0.0)
                imposto_sel = st.selectbox("Imposto Aplicável", ["Nenhum"] + list(st.session_state.param_impostos["Tipo"].unique()))

            observacoes_inp = st.text_area("Observações do Lançamento", placeholder="Detalhes opcionais sobre este atendimento...")

            # Cálculo de prévia
            taxa_pct_lookup = st.session_state.param_taxas.loc[
                st.session_state.param_taxas["Forma de Pagamento"] == forma_pag_sel, "Taxa (%)"
            ].values
            taxa_pct_val = float(taxa_pct_lookup[0]) if len(taxa_pct_lookup) > 0 else 0.0

            imposto_pct_val = 0.0
            if imposto_sel != "Nenhum":
                imp_lookup = st.session_state.param_impostos.loc[
                    st.session_state.param_impostos["Tipo"] == imposto_sel, "Taxa (%)"
                ].values
                imposto_pct_val = float(imp_lookup[0]) if len(imp_lookup) > 0 else 0.0

            calc = calcular_lancamento(
                valor_bruto=valor_bruto_inp,
                taxa_pagamento_pct=taxa_pct_val,
                material_clinica=material_clinica_inp,
                material_adicional=material_adicional_inp,
                material_dra=material_dra_inp,
                aliquota_imposto_pct=imposto_pct_val
            )

            st.markdown('<div class="preview-card">', unsafe_allow_html=True)
            st.markdown("### 📋 Prévia dos Cálculos")
            p_col1, p_col2, p_col3, p_col4 = st.columns(4)
            p_col1.write(f"**Taxa ({taxa_pct_val:.2f}%):** {fmt_brl(calc['valor_taxa_pagamento'])}")
            p_col2.write(f"**Pós Taxa:** {fmt_brl(calc['valor_apos_taxa'])}")
            p_col3.write(f"**Imposto:** {fmt_brl(calc['custo_nota_fiscal'])}")
            p_col4.write(f"**Lucro Líquido:** {fmt_brl(calc['lucro_liquido'])}")
            
            pr_col1, pr_col2 = st.columns(2)
            pr_col1.write(f"👉 **Repasse Profissional:** {fmt_brl(calc['valor_final_dra'])}")
            pr_col2.write(f"👉 **Clínica c/ Material Total:** {fmt_brl(calc['valor_final_clinica_com_material'])}")
            st.markdown('</div>', unsafe_allow_html=True)

            if st.form_submit_button("💾 Salvar e Sincronizar na Planilha", use_container_width=True):
                if not paciente:
                    st.error("Por favor, informe o nome do paciente.")
                elif valor_bruto_inp <= 0:
                    st.error("O valor bruto deve ser maior que zero.")
                else:
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    novo_id = f"LAN-{len(st.session_state.lancamentos) + 1:03d}"
                    
                    novo_registro = {
                        "ID": novo_id,
                        "DATA ATENDIMENTO": str(data_atend),
                        "PACIENTE": paciente,
                        "PROFISSIONAL": profissional,
                        "ORIGEM": origem,
                        "PROCEDIMENTO": procedimento_sel,
                        "Valor Bruto (R$)": calc["valor_bruto"],
                        "Forma de Pagamento": forma_pag_sel,
                        "Taxa de Pagamento (%)": calc["taxa_pagamento_pct"],
                        "Valor da Taxa de Pagamento (R$)": calc["valor_taxa_pagamento"],
                        "Valor após Taxa de Pagamento (R$)": calc["valor_apos_taxa"],
                        "Material da clínica (R$)": calc["material_clinica"],
                        "Custo Adicional Material (R$)": calc["material_adicional"],
                        "Material da Dra. Denise (R$)": calc["material_dra"],
                        "Tipo de Imposto": imposto_sel,
                        "Custo Nota Fiscal (R$)": calc["custo_nota_fiscal"],
                        "Lucro Líquido (R$)": calc["lucro_liquido"],
                        "Valor Repasse Clínica (R$)": calc["valor_repassado_clinica"],
                        "Valor Final Dra. Denise (R$)": calc["valor_final_dra"],
                        "Valor Final Clínica/Com Material (R$)": calc["valor_final_clinica_com_material"],
                        "Observações": observacoes_inp,
                        "CRIADO EM": now_str,
                        "USUARIO CRIACAO": user["nome"],
                        "STATUS": "ATIVO",
                        "EXCLUIDO EM": "",
                        "USUARIO EXCLUSAO": ""
                    }
                    
                    st.session_state.lancamentos = pd.concat([
                        st.session_state.lancamentos, pd.DataFrame([novo_registro])
                    ], ignore_index=True)

                    # Enviar para o Google Sheets
                    with st.spinner("Gravando na planilha do Google Sheets..."):
                        synced = sync_to_google_sheets(st.session_state.lancamentos)
                        if synced:
                            st.success(f"✅ Lançamento {novo_id} salvo localmente e GRAVADO com sucesso no Google Sheets!")

    # ==========================================================================
    # MODULO: LANÇAMENTOS
    # ==========================================================================
    elif menu == "📋 Lançamentos":
        st.markdown('<div class="main-header">RECOVERY — Histórico de Lançamentos</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Consulta completa com controle de auditoria e status do registro.</div>', unsafe_allow_html=True)

        df_lan = st.session_state.lancamentos

        if not df_lan.empty:
            ver_excluidos = st.checkbox("Exibir lançamentos marcados como EXCLUÍDO", value=True)
            if not ver_excluidos:
                df_exibir = df_lan[df_lan["STATUS"] == "ATIVO"]
            else:
                df_exibir = df_lan

            st.dataframe(df_exibir, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            if user["pode_excluir_alterar"]:
                st.subheader("🗑️ Cancelar / Marcar como Excluído")
                
                lan_ativos = list(df_lan[df_lan["STATUS"] == "ATIVO"]["ID"].unique())
                
                if lan_ativos:
                    sel_del = st.selectbox("Selecione o ID do Lançamento para marcar como EXCLUÍDO:", lan_ativos)
                    if st.button("Confirmar Exclusão e Atualizar Planilha", type="primary"):
                        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        
                        # Atualiza o status e salva o histórico em vez de apagar
                        st.session_state.lancamentos.loc[
                            st.session_state.lancamentos["ID"] == sel_del, ["STATUS", "EXCLUIDO EM", "USUARIO EXCLUSAO"]
                        ] = ["EXCLUÍDO", now_str, user["nome"]]
                        
                        with st.spinner("Atualizando histórico na planilha..."):
                            sync_to_google_sheets(st.session_state.lancamentos)
                            st.success(f"Lançamento {sel_del} marcado como EXCLUÍDO e atualizado no Google Sheets!")
                            st.rerun()
                else:
                    st.info("Não há lançamentos ativos para exclusão.")
            else:
                st.warning("⚠️ Seu perfil (Auxiliar Júnior) possui permissão apenas para visualização.")
        else:
            st.info("Nenhum lançamento cadastrado.")

    # ==========================================================================
    # MODULO: PARÂMETROS
    # ==========================================================================
    elif menu == "⚙️ Parâmetros":
        st.markdown('<div class="main-header">RECOVERY — Parâmetros do Sistema</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Configuração de taxas, procedimentos e impostos.</div>', unsafe_allow_html=True)

        if not user["pode_editar_parametros"]:
            st.warning("⚠️ Seu perfil de acesso permite apenas a visualização de parâmetros.")
            st.write("### Procedimentos Cadastrados")
            st.table(st.session_state.param_procedimentos)
            st.write("### Taxas de Pagamento")
            st.table(st.session_state.param_taxas)
        else:
            tab1, tab2, tab3 = st.tabs(["Procedimentos & Materiais", "Taxas de Pagamento", "Impostos"])

            with tab1:
                df_proc_ed = st.data_editor(st.session_state.param_procedimentos, num_rows="dynamic", use_container_width=True)
                if st.button("Salvar Alterações em Procedimentos"):
                    st.session_state.param_procedimentos = df_proc_ed
                    st.success("Procedimentos atualizados!")

            with tab2:
                df_taxas_ed = st.data_editor(st.session_state.param_taxas, num_rows="dynamic", use_container_width=True)
                if st.button("Salvar Alterações em Taxas"):
                    st.session_state.param_taxas = df_taxas_ed
                    st.success("Taxas atualizadas!")

            with tab3:
                df_imp_ed = st.data_editor(st.session_state.param_impostos, num_rows="dynamic", use_container_width=True)
                if st.button("Salvar Alterações em Impostos"):
                    st.session_state.param_impostos = df_imp_ed
                    st.success("Impostos atualizados!")

    # ==========================================================================
    # MODULO: RELATÓRIOS
    # ==========================================================================
    elif menu == "📈 Relatórios":
        st.markdown('<div class="main-header">RECOVERY — Relatórios Financeiros</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Exportação consolidada dos dados completos.</div>', unsafe_allow_html=True)

        df_lan = st.session_state.lancamentos
        if not df_lan.empty:
            col_exp1, col_exp2 = st.columns(2)
            
            csv_data = df_lan.to_csv(index=False).encode('utf-8')
            col_exp1.download_button("📥 Baixar Relatório CSV", data=csv_data, file_name=f"relatorio_recovery_{date.today()}.csv", mime="text/csv", use_container_width=True)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_lan.to_excel(writer, index=False, sheet_name='Lançamentos')
            col_exp2.download_button("📊 Baixar Relatório Excel", data=output.getvalue(), file_name=f"relatorio_recovery_{date.today()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        else:
            st.info("Não há dados para exportação.")
