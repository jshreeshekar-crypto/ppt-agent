import streamlit as st
import os

from app import (
    get_template,
    extract_content_with_llm,
    generate_ppt,
    TEMPLATES
)

# 🔹 Page config
st.set_page_config(
    page_title="PPT Agent",
    page_icon="📊",
    layout="centered"
)

# 🔹 Title
st.title("📊 AI PPT Generator")
st.caption("Describe the presentation you need and download the PPT instantly.")

# 🔹 API Key Check
if not os.getenv("OPENAI_API_KEY"):
    st.error("OPENAI_API_KEY not found")
    st.stop()

# 🔹 Session state
if "extracted" not in st.session_state:
    st.session_state.extracted = None

if "template_path" not in st.session_state:
    st.session_state.template_path = None

if "template_name" not in st.session_state:
    st.session_state.template_name = None

# 🔹 User input
user_input = st.text_area(
    "What presentation do you need?",
    placeholder="Example: Create a sprint review presentation for Team Alpha with delivery highlights and risks.",
    height=150
)

# 🔹 Generate button
if st.button("🤖 Generate PPT"):

    if not user_input.strip():
        st.warning("Please enter a request.")

    else:

        template_path, template_name = get_template(user_input)

        if not template_path:
            st.error("No matching template found.")

            st.info(
                "Available templates:\n\n"
                + "\n".join([f"- {k}" for k in TEMPLATES.keys()])
            )

        else:

            with st.spinner("Understanding your request..."):

                try:

                    extracted = extract_content_with_llm(
                        user_input,
                        template_name
                    )

                    st.session_state.extracted = extracted
                    st.session_state.template_path = template_path
                    st.session_state.template_name = template_name

                except Exception as e:
                    st.error(f"LLM extraction failed: {e}")

# 🔹 Review section
if st.session_state.extracted:

    st.subheader("📋 Review & Edit Content")

    extracted = st.session_state.extracted

    col1, col2 = st.columns(2)

    with col1:
        title = st.text_input(
            "Title",
            value=extracted.get("title", "")
        )

        objective = st.text_input(
            "Objective",
            value=extracted.get("objective", "")
        )

        agenda = st.text_input(
            "Agenda",
            value=extracted.get("agenda", "")
        )

    with col2:
        risks = st.text_input(
            "Risks",
            value=extracted.get("risks", "")
        )

        team = st.text_input(
            "Team",
            value=extracted.get("team", "")
        )

        date = st.text_input(
            "Date",
            value=extracted.get("date", "")
        )

    # 🔹 Confirm button
    if st.button("✅ Confirm & Build PPT"):

        data = {
            "{{TITLE}}": title,
            "{{OBJECTIVE}}": objective,
            "{{AGENDA}}": agenda,
            "{{RISKS}}": risks,
            "{{TEAM}}": team,
            "{{DATE}}": date,
        }

        with st.spinner("Building presentation..."):

            try:

                output_path = generate_ppt(
                    st.session_state.template_path,
                    data
                )

                st.success("✅ Presentation generated successfully!")

                with open(output_path, "rb") as file:

                    st.download_button(
                        label="⬇️ Download PPT",
                        data=file,
                        file_name="generated_presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

            except Exception as e:
                st.error(f"PPT generation failed: {e}")