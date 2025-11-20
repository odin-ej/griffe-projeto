# -*- coding: utf-8 -*-
"""
GRIFFE HUB - Revisor de Matrículas (Versão Standalone)
Sistema de revisão de formulários de matrícula
Arquivo atualizado: implementa normalização de colunas, tolerância a variações,
mensagens de aviso quando abas/colunas estão faltando e tratamento de planilhas vazias.
"""

import streamlit as st
import pandas as pd
from typing import Dict, List
from pathlib import Path

# ============================================================================	
# CONFIGURAÇÃO DA PÁGINA
# ============================================================================	

st.set_page_config(
    page_title="Revisor de Matrículas - Griffe Hub",
    page_icon="📋",
    layout="wide"
)

# ============================================================================	
# CLASSES E FUNÇÕES DO BACKEND (INCORPORADAS)
# ============================================================================	

class ExcelReader:
    """Classe para ler e processar dados da planilha de matrículas"""

    EXPECTED_SHEETS = ["Form_Matrícula", "Form_Inicial", "Form_Médico"]

    def __init__(self, excel_file_path: str):
        self.excel_file = excel_file_path
        self.df_matricula = pd.DataFrame()
        self.df_inicial = pd.DataFrame()
        self.df_medico = pd.DataFrame()
        self.students = []
        self.sheet_status = {}  # status de cada sheet lida
        self.missing_columns = {}  # possíveis colunas faltando por sheet

    def _normalize_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Padroniza colunas para evitar erros de KeyError (strip, upper, remove acentos)."""
        if df is None or df.shape[1] == 0:
            return df
        # Força string nas colunas e padroniza
        cols = (
            df.columns.astype(str)
              .str.strip()
              .str.upper()
              .str.normalize('NFKD')
              .str.encode('ascii', errors='ignore')
              .str.decode('utf-8')
        )
        df = df.copy()
        df.columns = cols
        return df

    def load_data(self) -> bool:
        """Carrega os dados das 3 sheets principais de forma tolerante.

        Mantém os nomes das abas conforme pedido (com acento).
        Se uma aba não existir, cria um DataFrame vazio para ela e registra o aviso.
        """
        try:
            xls = pd.ExcelFile(self.excel_file)
            available = xls.sheet_names
            # Ler cada sheet somente se existir
            for sheet in self.EXPECTED_SHEETS:
                if sheet in available:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    df = self._normalize_columns(df)
                    setattr(self, f"df_{sheet.split('_')[1].lower()}", df)
                    self.sheet_status[sheet] = "loaded"
                else:
                    # Mantém DataFrame vazio mas registra falta
                    setattr(self, f"df_{sheet.split('_')[1].lower()}", pd.DataFrame())
                    self.sheet_status[sheet] = "missing"

            # Após carregar, verifica colunas importantes e registra missing_columns
            for sheet_attr, sheet_name in [
                ("df_matrícula", "Form_Matrícula"),
                ("df_inicial", "Form_Inicial"),
                ("df_médico", "Form_Médico"),
            ]:
                df = getattr(self, sheet_attr.replace("í","i").replace("ê","e"), pd.DataFrame())
                # Note: attributes in object are df_matricula, df_inicial, df_medico
                # We'll map correctly:
            # Map to the actual attributes we set earlier
            self.df_matricula = getattr(self, "df_matricula", pd.DataFrame())
            self.df_inicial = getattr(self, "df_inicial", pd.DataFrame())
            self.df_medico = getattr(self, "df_medico", pd.DataFrame())

            # Check for key columns in each loaded sheet
            for sheet, df in [("Form_Matrícula", self.df_matricula),
                              ("Form_Inicial", self.df_inicial),
                              ("Form_Médico", self.df_medico)]:
                missing = []
                if df is None or df.shape[1] == 0:
                    # empty sheet
                    self.missing_columns[sheet] = ["(sheet empty or not loaded)"]
                    continue
                # Find approximations for Nome Completo and Email
                nome_col = next((c for c in df.columns if "NOME" in c and "COMPLETO" in c), None)
                email_col = next((c for c in df.columns if "EMAIL" in c), None)
                if not nome_col:
                    missing.append("NOME COMPLETO (approx)")
                if not email_col:
                    missing.append("EMAIL (approx)")
                self.missing_columns[sheet] = missing if missing else []

            # Build student list
            self._build_student_list()
            return True
        except Exception as e:
            st.error(f"Erro ao carregar dados: {e}")
            return False

    def _build_student_list(self):
        """Cria lista de estudantes únicos baseado em NOME COMPLETO e EMAIL (tolerante)"""
        students_set = set()

        for df, sheet in [(self.df_matricula, "Form_Matrícula"),
                          (self.df_inicial, "Form_Inicial"),
                          (self.df_medico, "Form_Médico")]:
            if df is None or df.shape[0] == 0:
                continue
            nome_col = next((c for c in df.columns if "NOME" in c and "COMPLETO" in c), None)
            email_col = next((c for c in df.columns if "EMAIL" in c), None)

            if not nome_col:
                # não conseguimos identificar coluna de nome nessa sheet
                continue

            for _, row in df.iterrows():
                nome = row.get(nome_col, "")
                email = row.get(email_col, "") if email_col else ""
                if pd.notna(nome) and str(nome).strip():
                    students_set.add((str(nome).strip(), str(email).strip() if pd.notna(email) and str(email).strip() else ""))

        self.students = sorted(list(students_set), key=lambda x: x[0])

    def get_students(self) -> List[Dict[str, str]]:
        """Retorna lista de estudantes"""
        return [{'nome': nome, 'email': email} for nome, email in self.students]

    def _match_row(self, df: pd.DataFrame, nome: str, email: str):
        """Retorna linha correspondente tolerante a variações de nomes de colunas"""
        if df is None or df.shape[0] == 0:
            return None

        nome_col = next((c for c in df.columns if "NOME" in c and "COMPLETO" in c), None)
        email_col = next((c for c in df.columns if "EMAIL" in c), None)
        if not nome_col:
            return None

        # Comparação tolerante
        mask = df[nome_col].astype(str).str.strip() == nome.strip()
        if email and email_col:
            mask = mask & (df[email_col].astype(str).str.strip() == email.strip())

        matches = df[mask]
        if len(matches) > 0:
            return matches.iloc[0].to_dict()
        return None

    def get_student_data(self, nome: str, email: str) -> Dict:
        """Busca dados de um estudante específico em todas as sheets"""
        return {
            'matricula': self._match_row(self.df_matricula, nome, email),
            'inicial': self._match_row(self.df_inicial, nome, email),
            'medico': self._match_row(self.df_medico, nome, email),
        }

# ============================================================================	
# Mapeamento de campos (mantido)
# ============================================================================	

FORM_MATRICULA_SECTIONS = {
    "Section 1 - Student Information": [
        "NOME DO ESTUDANTE", "SOBRENOME COMPLETO DO ESTUDANTE",
        "DATA DE NASCIMENTO DO ESTUDANTE", "SEXO DO ESTUDANTE",
        "PAIS DE NASCIMENTO DO ESTUDANTE", "EMAIL DO ESTUDANTE",
    ],
    "Section 2 - Passport Information": [
        "O ESTUDANTE POSSUI PASSAPORTE?",
        "SE SIM, O PASSAPORTE DO ESTUDANTE ESTA VALIDO?",
        "SE O ESTUDANTE TEM PASSAPORTE INFORME O NUMERO",
        "SE O ESTUDANTE TEM PASSAPORTE INFORME A DATA DE VALIDADE",
    ],
    "Section 3 - Parent One Information": [
        "NOME COMPLETO DA MAE", "NUMERO DE TELEFONE DA MAE (COM WHATSAPP)",
        "EMAIL DA MAE", "DATA DE NASCIMENTO DA SUA MAE",
    ],
    "Section 4 - Parent Two Information": [
        "NOME COMPLETO DO PAI", "NUMERO DE TELEFONE DO PAI (COM WHATSAPP)",
        "EMAIL DO PAI", "DATA DE NASCIMENTO DO SEU PAI",
    ],
    "Section 5 - Academic & Interests": [
        "VOCE GOSTA DE IR PARA ESCOLA?", "SUAS TRES MATERIAS FAVORITAS",
        "CONTE QUAIS SAO SEUS PLANOS PARA O FUTURO",
    ],
    "Section 6 - Homestay Preferences": [
        "VOCE PREFERE MORAR EM:", "VOCE GOSTA DE ANIMAIS DE ESTIMACAO?",
        "VOCE FUMA?",
    ],
}

FORM_INICIAL_SECTIONS = {
    "Section 1 - Student Information": [
        "NOME DO ESTUDANTE", "SOBRENOME COMPLETO DO ESTUDANTE",
        "NUMERO DO CPF DO ESTUDANTE", "NUMERO DO RG DO ESTUDANTE",
        "DATA DE NASCIMENTO DO ESTUDANTE", "SEXO DO ESTUDANTE",
    ],
    "Section 2 - Parent Information": [
        "NOME COMPLETO DA MAE", "NUMERO DE TELEFONE DA MAE (COM WHATSAPP)",
        "NOME COMPLETO DO PAI", "NUMERO DE TELEFONE DO PAI (COM WHATSAPP)",
    ],
    "Section 3 - Address": [
        "CEP DO ENDERECO DE RESIDENCIA DO ESTUDANTE",
        "ENDERECO COMPLETO DE RESIDENCIA DO ESTUDANTE",
    ],
}

FORM_MEDICO_SECTIONS = {
    "Section 1 - Health Conditions": [
        "VOCE TEM ALGUM PROBLEMA DE SAUDE?",
        "SE SIM, DESCREVA SUA(S) CONDICAO(OES) DE SAUDE:",
        "CONDICOES DE SAUDE (ATUAIS OU PASSADAS)",
    ],
    "Section 2 - Allergies": [
        "VOCE TEM ALGUM TIPO DE ALERGIA?",
        "CASO TENHA ALGUMA ALERGIA ASSINALADA, FAVOR DAR MAIS INFORMACOES",
    ],
    "Section 3 - Medications": [
        "VOCE FAZ USO DE ALGUM MEDICAMENTO DE FORMA CONTINUA (TODOS OS DIAS)?",
        "SE SIM, LISTE O MEDICAMENTO, DOSAGEM E FREQUENCIA:",
    ],
}

def get_field_label(field_name: str) -> str:
    """Converte nome de campo em label amigável"""
    label = field_name
    # Remove emojis
    for emoji in ['1️⃣', '2️⃣', '3️⃣', '4️⃣', '5️⃣', '6️⃣', '7️⃣', '8️⃣', '9️⃣', '🔟']:
        label = label.replace(emoji, '')
    label = label.replace('1️⃣1️⃣', '').replace('1️⃣2️⃣', '')
    return label.strip()

# ============================================================================	
# FUNÇÕES AUXILIARES DE INTERFACE
# ============================================================================	

def format_value(value):
    """Formata valor para exibição"""
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d/%m/%Y")
    return str(value)

def render_field(label, value, key):
    """Renderiza um campo com botão de copiar"""
    col1, col2 = st.columns([4, 1])

    with col1:
        st.markdown(f"**{label}**")
        formatted_value = format_value(value)

        if formatted_value:
            st.markdown(
                f'<div style="background-color: #f0f2f6; padding: 10px; '
                f'border-radius: 5px; margin-bottom: 10px;">{formatted_value}</div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                f'<div style="background-color: #fff3cd; padding: 10px; '
                f'border-radius: 5px; margin-bottom: 10px; color: #856404;">'
                f'<em>Não preenchido</em></div>',
                unsafe_allow_html=True
            )

    with col2:
        if formatted_value:
            if st.button("📋", key=f"copy_{key}", use_container_width=True, 
                        help="Clique para copiar"):
                st.code(formatted_value, language=None)

def render_section(section_title, fields, data, form_type):
    """Renderiza uma seção do formulário"""
    st.markdown(f"### {section_title}")
    st.markdown("---")

    for field in fields:
        value = data.get(field, "")
        key = f"{form_type}_{field}_{section_title}"
        label = get_field_label(field)
        render_field(label, value, key)

    st.markdown("<br>", unsafe_allow_html=True)

# ============================================================================	
# INTERFACE PRINCIPAL
# ============================================================================	

st.title("📋 Revisor de Matrículas")
st.markdown("Sistema de Revisão de Formulários de Matrícula")
st.markdown("---")

# SIDEBAR
with st.sidebar:
    st.markdown("### 🏠 Navegação")
    if st.button("← Voltar ao Hub", use_container_width=True):
        try:
            st.switch_page("streamlit_app.py")
        except Exception:
            # fallback: não quebrar caso não exista a página
            pass

    st.markdown("---")
    st.markdown("### ℹ️ Como Usar")
    st.markdown("""
    1. **Faça upload** da planilha Excel
    2. **Selecione um aluno** na lista
    3. **Visualize** os dados dos formulários
    4. **Copie** os campos necessários
    """)
    st.markdown("---")
    st.markdown("**Status:** ✅ Online")

# UPLOAD
st.header("📁 Upload da Planilha")

uploaded_file = st.file_uploader(
    "Selecione a planilha Excel com os dados",
    type=['xlsx', 'xls'],
    help="Arquivo deve conter: Form_Matrícula, Form_Inicial, Form_Médico"
)

if uploaded_file:
    import tempfile
    import os
    from io import BytesIO

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name

    try:
        with st.spinner("Carregando dados..."):
            reader = ExcelReader(tmp_path)
            ok = reader.load_data()
            # Mostrar avisos úteis sobre sheets e colunas
            missing_sheets = [s for s, status in reader.sheet_status.items() if status != "loaded"]
            if missing_sheets:
                st.warning(f"As seguintes abas não foram encontradas: {', '.join(missing_sheets)}. O app continuará, mas algumas informações podem ficar incompletas.")
            # Mostrar colunas faltando por sheet (se houver)
            for sheet, missing in reader.missing_columns.items():
                if missing:
                    st.info(f"Atenção - {sheet}: {', '.join(missing)}")
            if ok:
                st.success(f"✅ {len(reader.get_students())} estudantes encontrados")
                st.session_state['reader'] = reader
                st.session_state['students'] = reader.get_students()
            else:
                st.error("Falha ao carregar a planilha. Veja as mensagens acima.")
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

# VISUALIZAÇÃO
if 'students' in st.session_state and st.session_state['students']:
    st.markdown("---")
    st.header("👤 Seleção de Aluno")

    student_options = [
        f"{s['nome']} ({s['email']})" if s['email'] else s['nome']
        for s in st.session_state['students']
    ]

    selected_index = st.selectbox(
        "Selecione um aluno:",
        range(len(student_options)),
        format_func=lambda i: student_options[i]
    )

    if selected_index is not None:
        selected_student = st.session_state['students'][selected_index]
        reader = st.session_state['reader']
        student_data = reader.get_student_data(
            selected_student['nome'], 
            selected_student['email']
        )

        st.markdown("---")

        tab1, tab2, tab3 = st.tabs([
            "📝 Form Matrícula", 
            "📄 Form Inicial", 
            "🏥 Form Médico"
        ])

        with tab1:
            st.header("📝 Formulário de Matrícula")
            if student_data['matricula']:
                for section, fields in FORM_MATRICULA_SECTIONS.items():
                    with st.expander(section, expanded=False):
                        render_section(section, fields, 
                                     student_data['matricula'], 'matricula')
            else:
                st.warning("⚠️ Sem dados")

        with tab2:
            st.header("📄 Formulário Inicial")
            if student_data['inicial']:
                for section, fields in FORM_INICIAL_SECTIONS.items():
                    with st.expander(section, expanded=False):
                        render_section(section, fields, 
                                     student_data['inicial'], 'inicial')
            else:
                st.warning("⚠️ Sem dados")

        with tab3:
            st.header("🏥 Formulário Médico")
            if student_data['medico']:
                for section, fields in FORM_MEDICO_SECTIONS.items():
                    with st.expander(section, expanded=False):
                        render_section(section, fields, 
                                     student_data['medico'], 'medico')
            else:
                st.warning("⚠️ Sem dados")
else:
    st.info("👆 Faça upload de uma planilha para começar")
