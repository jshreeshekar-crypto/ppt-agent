from pptx import Presentation
from pptx.util import Pt
from openai import OpenAI
from dotenv import load_dotenv
import json, os

load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# 🔹 STEP 1: Templates
TEMPLATES = {
    "pi planning": "templates/pi_planning.pptx",
    "sprint review": "templates/sprint_review.pptx",
    "retrospective": "templates/retrospective.pptx",
    "daily status": "templates/daily_status.pptx",
    "system demo": "templates/System_Demo.pptx",
}

# 🔹 STEP 2: Template selection
def get_template(user_input):
    user_input = user_input.lower()
    for key, path in TEMPLATES.items():
        if key in user_input:
            return path, key
    return None, None

# 🔹 STEP 3: LLM extracts content from user input
def extract_content_with_llm(user_input, template_name):
    prompt = f"""
You are a helpful assistant that extracts presentation content from user requests.

The user wants a "{template_name}" presentation.
User request: "{user_input}"

Extract and return a JSON object with these fields:
- title: presentation title
- objective: main objective or goal
- agenda: key agenda items (comma separated)
- risks: any risks or concerns mentioned (or "To be identified" if none)
- team: team name if mentioned (or "Team")
- date: date if mentioned (or "Today")

Return ONLY valid JSON, no extra text.
"""
    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.3,
    )
    raw = response.choices[0].message.content.strip()
    return json.loads(raw)

# 🔹 STEP 4: PPT generation
def generate_ppt(template_path, data):
    prs = Presentation(template_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        for key, value in data.items():
                            if key in run.text:
                                run.text = run.text.replace(key, value)

    os.makedirs("Output", exist_ok=True)
    output_path = "Output/generated_ppt.pptx"
    prs.save(output_path)
    return output_path

import streamlit as st

st.title("AI PPT Generator")

user_input = st.text_area(
    "Describe your presentation",
    placeholder="Create a PI Planning deck for Team Alpha..."
)

if st.button("Generate PPT"):
    template_path, template_name = get_template(user_input)

    if not template_path:
        st.error("No matching template found.")
    else:
        try:
            data = extract_content_with_llm(user_input, template_name)
            ppt_path = generate_ppt(template_path, data)

            st.success("Presentation generated successfully!")

            with open(ppt_path, "rb") as file:
                st.download_button(
                    label="Download PPT",
                    data=file,
                    file_name="generated_ppt.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

        except Exception as e:
            st.error(f"Error: {e}")