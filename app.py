from pptx import Presentation

# 🔹 STEP 1: Add all templates here
TEMPLATES = {
    "pi planning": "templates/pi_planning.pptx",
    "sprint review": "templates/sprint_review.pptx",
    "retrospective": "templates/retrospective.pptx",
    "daily status": "templates/daily_status.pptx",
   "system demo": "templates/System_Demo.pptx",
}


# 🔹 STEP 2: Template selection (improved)
def get_template(user_input):
    user_input = user_input.lower()

    for key, path in TEMPLATES.items():
        if key in user_input:
            return path

    return None


# 🔹 STEP 3: PPT generation
def generate_ppt(template_path, data):
    prs = Presentation(template_path)

    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                text = shape.text
                for key, value in data.items():
                    if key in text:
                        text = text.replace(key, value)
                shape.text = text

    output_path = "Output/generated_ppt.pptx"
    prs.save(output_path)
    return output_path


# 🔹 STEP 4: Main
if __name__ == "__main__":
    user_input = input("Enter your request: ")

    template = get_template(user_input)

    if template:
        # You can later customize this per template
        data = {
            "{{TITLE}}": "Auto Generated Deck",
            "{{OBJECTIVE}}": "Generated based on user input",
            "{{AGENDA}}": "Discussion points",
            "{{RISKS}}": "To be identified"
        }

        output = generate_ppt(template, data)
        print(f"PPT generated: {output}")

    else:
        print("\n❌ Template not found.")
        print("Available templates:")
        for key in TEMPLATES.keys():
            print(f"- {key}")