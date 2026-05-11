from pptx import Presentation
from openai import OpenAI
from dotenv import load_dotenv
import json
import os

# Load environment variables
load_dotenv()

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# 🔹 Templates
TEMPLATES = {
    "pi planning": "templates/pi_planning.pptx",
    "sprint review": "templates/sprint_review.pptx",
    "retrospective": "templates/retrospective.pptx",
    "daily status": "templates/daily_status.pptx",
    "system demo": "templates/System_Demo.pptx",
}

# 🔹 Template selection
def get_template(user_input):
    user_input = user_input.lower()

    for key, path in TEMPLATES.items():
        if key in user_input:
            return path, key

    return None, None

# 🔹 Extract content using OpenAI
def extract_content_with_llm(user_input, template_name):

    prompt = f"""
You are a helpful assistant that extracts presentation content from user requests.

The user wants a "{template_name}" presentation.

User request:
"{user_input}"

Return ONLY valid JSON with these fields:
- title
- objective
- agenda
- risks
- team
- date
"""

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "user", "content": prompt}
        ],
        temperature=0.3
    )

    raw = response.choices[0].message.content.strip()

    return json.loads(raw)

# 🔹 Generate PPT
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