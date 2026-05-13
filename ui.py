import streamlit as st
import os
from app import get_template, extract_content_with_llm, generate_ppt, TEMPLATES

st.set_page_config(page_title="PPT Agent", page_icon="📊", layout="centered")
st.title("📊 AI PPT Generator")
st.caption("Describe the presentation you need and download the PPT instantly.")

# ✅ Check BOTH Streamlit secrets AND .env
api_key = st.secrets.get("GOOGLE_API_KEY") if "OPENAI_API_KEY" in st.secrets else os.getenv("OPENAI_API_KEY")

if not api_key:
    st.error("❌ GOOGLE_API_KEY not found! Add it in Streamlit Secrets.")
    st.stop()

# --- Initialize session state ---
if "extracted" not in st.session_state:
    st.session_state.extracted = None
if "template_path" not in st.session_state:
    st.session_state.template_path = None
if "template_name" not in st.session_state:
    st.session_state.template_name = None

# --- User Input ---
user_input = st.text_area(
    "What presentation do you need?",
    placeholder="e.g. Create a sprint review for Q2, focus on velocity improvements...",
    height=120
)

if st.button("🤖 Analyze Request"):
    if not user_input.strip():
        st.warning("Please enter a request.")
    else:
        template_path, template_name = get_template(user_input)

        if template_path:
            with st.spinner("🧠 Understanding your request..."):
                try:
                    st.session_state.extracted = extract_content_with_llm(user_input, template_name)
                    st.session_state.template_path = template_path
                    st.session_state.template_name = template_name
                except Exception as e:
                    st.error(f"❌ LLM extraction failed: {e}")
        else:
            st.error("❌ No matching template found.")
            st.info("**Available templates:**\n" + "\n".join([f"- `{k}`" for k in TEMPLATES.keys()]))

# --- Show Extracted Content ---
if st.session_state.extracted:
    st.subheader("📋 Extracted Content — Review & Edit")

    col1, col2 = st.columns(2)
    with col1:
        title     = st.text_input("Title",     value=st.session_state.extracted.get("title", ""))
        objective = st.text_input("Objective", value=st.session_state.extracted.get("objective", ""))
        agenda    = st.text_input("Agenda",    value=st.session_state.extracted.get("agenda", ""))
    with col2:
        risks = st.text_input("Risks", value=st.session_state.extracted.get("risks", ""))
        team  = st.text_input("Team",  value=st.session_state.extracted.get("team", ""))
        date  = st.text_input("Date",  value=st.session_state.extracted.get("date", ""))

    if st.button("✅ Generate & Download PPT"):
        data = {
            "{{TITLE}}":     title,
            "{{OBJECTIVE}}": objective,
            "{{AGENDA}}":    agenda,
            "{{RISKS}}":     risks,
            "{{TEAM}}":      team,
            "{{DATE}}":      date,
        }

        with st.spinner("📊 Building your presentation..."):
            try:
                output = generate_ppt(st.session_state.template_path, data)
            except Exception as e:
                st.error(f"❌ PPT generation failed: {e}")
                st.stop()

        st.success("✅ PPT Ready!")
        with open(output, "rb") as f:
            st.download_button(
                label="⬇️ Download Presentation",
                data=f,
                file_name=f"{st.session_state.template_name.replace(' ', '_')}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )