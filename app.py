from pptx import Presentation
from openai import OpenAI
from dotenv import load_dotenv
import json
import os
import re

# =========================================================
# Load Environment Variables
# =========================================================
load_dotenv()

api_key = os.getenv("OPENAI_API_KEY")

if not api_key:
    raise ValueError("OPENAI_API_KEY not found")

client = OpenAI(api_key=api_key)

# =========================================================
# PPT Templates
# =========================================================
TEMPLATES = {
    "pi planning": "templates/pi_planning.pptx",
    "sprint review": "templates/sprint_review.pptx",
    "retrospective": "templates/retrospective.pptx",
    "daily status": "templates/daily_status.pptx",
    "system demo": "templates/System_Demo.pptx",
}

# =========================================================
# Select Template
# =========================================================
def get_template(user_input):

    user_input = user_input.lower()

    for key, path in TEMPLATES.items():
        if key in user_input:
            return path, key

    return None, None

# =========================================================
# Clean JSON response
# =========================================================
def clean_json_response(raw_text):

    raw_text = raw_text.strip()

    raw_text = re.sub(r"```json", "", raw_text)
    raw_text = re.sub(r"```", "", raw_text)

    return raw_text.strip()

# =========================================================
# AI Content Extraction
# =========================================================
def extract_content_with_llm(user_input, template_name):

    prompt = f"""
You are an enterprise Agile Release Train presentation assistant.

Your responsibility is to generate high-quality structured presentation content
for PowerPoint decks used in leadership meetings, PI planning sessions,
sprint reviews, retrospectives, and system demos.

=========================================================
TEMPLATE SELECTED:
{template_name}
=========================================================

USER REQUEST:
{user_input}

=========================================================
INSTRUCTIONS:
=========================================================

1. Understand the business context carefully.
2. Generate realistic enterprise-level Agile content.
3. Use professional language suitable for leadership reviews.
4. Infer missing details intelligently.
5. Use SAFe / Agile terminology where appropriate.
6. Keep content concise and presentation-ready.
7. Return ONLY valid JSON.
8. Do NOT return markdown.
9. Do NOT add explanations.

=========================================================
RETURN JSON FORMAT:
=========================================================

{{
    "title": "Presentation title",
    "objective": "Business objective",
    "agenda": "Agenda items separated by commas",
    "summary": "Executive summary",
    "highlights": "Key achievements or accomplishments",
    "risks": "Major risks or blockers",
    "dependencies": "Dependencies or coordination items",
    "metrics": "KPIs, velocity, predictability, delivery metrics",
    "team": "Team or ART name",
    "date": "PI number or date"
}}

=========================================================
EXAMPLE:
=========================================================

If user says:
'Create a PI Planning deck for PI 26.1 focusing on dependency management and feature readiness'

Then generate content related to:
- feature readiness
- dependency alignment
- planning risks
- ART coordination
- capacity planning
- milestones
- execution confidence

Return ONLY JSON.
"""

    try:

        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {
                    "role": "system",
                    "content": "You are a professional enterprise presentation assistant."
                },
                {
                    "role": "user",
                    "content": prompt
                }
            ],
            temperature=0.4
        )

        raw = response.choices[0].message.content

        cleaned = clean_json_response(raw)

        return json.loads(cleaned)

    except Exception as e:
        raise Exception(f"LLM extraction failed: {e}")

# =========================================================
# Replace text inside PPT
# =========================================================
def replace_text(shape, data):

    if not shape.has_text_frame:
        return

    for paragraph in shape.text_frame.paragraphs:

        full_text = ""

        for run in paragraph.runs:
            full_text += run.text

        updated_text = full_text

        for key, value in data.items():

            updated_text = updated_text.replace(key, str(value))

        if updated_text != full_text:

            paragraph.clear()

            paragraph.text = updated_text

# =========================================================
# Generate PPT
# =========================================================
def generate_ppt(template_path, data):

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template not found: {template_path}")

    prs = Presentation(template_path)

    for slide in prs.slides:

        for shape in slide.shapes:

            replace_text(shape, data)

    os.makedirs("Output", exist_ok=True)

    output_path = "Output/generated_ppt.pptx"

    prs.save(output_path)

    return output_path