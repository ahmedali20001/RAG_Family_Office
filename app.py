import streamlit as st
import os
import pandas as pd

# ── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Family Office Explorer",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium Look
st.markdown("""
<style>
    .main {
        background-color: #0e1117;
        color: #ffffff;
    }
    .stChatInput {
        border-radius: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ── LLM WRAPPERS ─────────────────────────────────────────────────────────────
def run_gemini(api_key: str, model_name: str, prompt: str) -> str:
    import google.generativeai as genai
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Error running Gemini API: {str(e)}"

def run_openai(api_key: str, model_name: str, prompt: str) -> str:
    from openai import OpenAI
    try:
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.2
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error running OpenAI API: {str(e)}"

# ── DATA LOADER ──────────────────────────────────────────────────────────────
@st.cache_data
def load_excel_data():
    """Loads the processed Excel file into a Pandas DataFrame."""
    excel_path = "family_offices_v2.xlsx"
    if not os.path.exists(excel_path):
        return None
    try:
        df = pd.read_excel(excel_path, sheet_name="Records")
        return df
    except Exception:
        return None

df = load_excel_data()

# ── SIDEBAR CONFIGURATION ─────────────────────────────────────────────────────
st.sidebar.image("https://img.icons8.com/color/96/000000/briefcase.png", width=60)
st.sidebar.title("AI Settings")

provider = st.sidebar.selectbox("LLM Provider", ["Google Gemini", "OpenAI"])
model_name = None

if provider == "Google Gemini":
    api_key = st.sidebar.text_input("Gemini API Key", type="password")
    
    if api_key:
        available_models = []
        try:
            import google.generativeai as genai
            genai.configure(api_key=api_key)
            models = genai.list_models()
            fetched_models = [m.name.replace("models/", "") for m in models if "generateContent" in m.supported_generation_methods]
            if fetched_models:
                available_models = fetched_models
        except Exception:
            available_models = ["gemini-1.5-flash-latest", "gemini-1.5-pro-latest", "gemini-pro"]
            
        if available_models:
            model_name = st.sidebar.selectbox("Model", available_models)
    else:
        st.sidebar.info("Please enter an API Key to load available models.")
else:
    api_key = st.sidebar.text_input("OpenAI API Key", type="password")
    if api_key:
        model_name = st.sidebar.selectbox("Model", ["gpt-4o-mini", "gpt-4o"])
    else:
        st.sidebar.info("Please enter an API Key to select a model.")

st.sidebar.markdown("---")
if df is not None:
    st.sidebar.markdown(f"""
    ### Document Scope
    - **Dataset Size:** {len(df)} Family Office Profiles
    - **Total Columns:** {len(df.columns)}
    - **Context Method:** Two-Step Column Filtering
    """)

# ── MAIN PANEL ───────────────────────────────────────────────────────────────
st.title("💼 Family Office Explorer")
st.markdown("Ask deep analytical questions across the entire dataset. The AI will dynamically select the relevant columns from the Excel file to provide hyper-accurate answers.")

if df is None:
    st.error("Could not load `family_offices_v2.xlsx`. Please run the extraction script first.")
    st.stop()

# Initialize Chat History
if "messages" not in st.session_state:
    st.session_state.messages = []

# Display Chat History
for message in st.session_state.messages:
    st.chat_message(message["role"]).markdown(message["content"])

# Query Input
query = st.chat_input("Ask a question (e.g., 'Who is the most experienced adviser in the dataset?')")

if query:
    # 1. Add user query to UI and history
    st.chat_message("user").markdown(query)
    st.session_state.messages.append({"role": "user", "content": query})
    
    if not api_key or not model_name:
        st.warning("⚠️ Please provide your API key in the sidebar to generate a response.")
        st.stop()
        
    all_columns = df.columns.tolist()
    
    # ── STEP 1: Determine Relevant Columns ──
    col_prompt = f"""You are a data router. Below are the available columns in a dataset of Family Office Advisers:
{', '.join(all_columns)}

Context for columns:
- "Total Disclosures" (number of disclosures)
- "Disclosure Details" (contains text describing all allegations, fines, penalties, and amounts)
- "Years of Experience", "Start Date", etc.

User Query: {query}

Which of these columns are strictly necessary to answer the user's query? 
If the user asks about penalties or disclosures, ensure you select both "Total Disclosures" and "Disclosure Details" to capture the full context.
Always include 'CRD#' and 'Full Name'.
Respond ONLY with a comma-separated list of exact column names from the provided list. Do not provide any other text or explanation."""

    with st.spinner("Determining required columns..."):
        if provider == "Google Gemini":
            col_response = run_gemini(api_key, model_name, col_prompt)
        else:
            col_response = run_openai(api_key, model_name, col_prompt)
            
    # Parse selected columns safely
    selected_cols = []
    for col in col_response.split(","):
        c = col.strip().strip("'").strip('"')
        if c in all_columns and c not in selected_cols:
            selected_cols.append(c)
            
    # Enforce primary keys
    if "CRD#" not in selected_cols:
        selected_cols.insert(0, "CRD#")
    if "Full Name" not in selected_cols:
        selected_cols.insert(1, "Full Name")
        
    # ── STEP 2: Answer the Query using the filtered data ──
    filtered_df = df[selected_cols]
    csv_data = filtered_df.to_csv(index=False)
    
    answer_prompt = f"""You are an expert institutional investment analyst reviewing regulatory profile data.
Below is the filtered dataset of {len(filtered_df)} Adviser Profiles in CSV format containing the exact columns needed. 
CRITICAL NOTE 1: ALL individuals in this dataset have already been verified as Family Offices from adviserinfo.sec.gov. 
CRITICAL NOTE 2: When asked for the "biggest", "most", or any comparison, YOU MUST READ THROUGH ALL ROWS IN THE CSV and compare their values before answering. Do not stop at the first match.
For example, if looking for the biggest penalty sum, extract the fine amounts from the disclosure text for EVERY person, sum them up per person, and THEN find the maximum.

User Query: {query}

--- CONTEXT DATA (CSV) ---
{csv_data}
"""

    with st.spinner("Analyzing data and generating answer..."):
        if provider == "Google Gemini":
            raw_answer = run_gemini(api_key, model_name, answer_prompt)
        else:
            raw_answer = run_openai(api_key, model_name, answer_prompt)
            
        # Combine routing context and answer into a single message
        full_answer = f"**🔍 AI analyzed these columns:** {', '.join(selected_cols)}\n\n{raw_answer}"
        
        st.chat_message("assistant").markdown(full_answer)
        st.session_state.messages.append({"role": "assistant", "content": full_answer})
