import os
from docx import Document
from jinja2 import Template
from datetime import datetime

ascii_art ="""

   ___ _                      _ 
  / __| |___ ______ __ _ _ __(_)
 | (_ | / _ (_-<_-</ _` | '_ \ |
  \___|_\___/__/__/\__,_| .__/_|
                        |_|     

"""

print(ascii_art)


def extract_content(doc_path):
    """Extracts content from a .docx file and structures it into a dictionary."""
    doc = Document(doc_path)
    content = {"title": "", "sections": []}
    
    for para in doc.paragraphs:
        style_name = getattr(para.style, "name", "")
        if style_name.startswith("Heading"):
            content["sections"].append({"heading": para.text.strip(), "content": ""})
        elif content["sections"]:
            content["sections"][-1]["content"] += f"<p>{para.text.strip()}</p>"
        elif not content["title"] and para.text.strip():
            content["title"] = para.text.strip()
    
    return content

def generate_html(content, template_path):
    """Generates HTML from extracted content using a Jinja2 template."""
    with open(template_path, "r", encoding="utf-8") as file:
        template = Template(file.read())
    
    html = template.render(
        title=content["title"],
        date=datetime.now().strftime("%d-%m-%Y"),
        sections=content["sections"]
    )
    return html

if __name__ == "__main__":
    print (ascii_art)
    print("Welcome to HavenThoughtaName Type 'exit' at any prompt to quit.\n")
    
    while True:  # Infinite loop until the user types 'exit'
        try:
            # Prompt for folder path
            base_path = input("Enter the folder path the default is the current folder: ").strip()
            if base_path.lower() == "exit":
                print("Exiting the program. Goodbye!")
                break
            
            base_path = os.path.abspath(base_path) if base_path else os.getcwd()  # Default to current directory
            
            # Prompt for Word document name
            doc_filename = input("Enter the name of the Word document (e.g., newsletter.docx): ").strip()
            if doc_filename.lower() == "exit":
                print("Exiting the program. Goodbye!")
                break
            
            # Prompt for HTML template name
            template_filename = input("Enter the name of the HTML template file (default: template.html): ").strip()
            if template_filename.lower() == "exit":
                print("Exiting the program. Goodbye!")
                break
            
            template_filename = template_filename if template_filename else "template.html"

            # Constructing file paths
            doc_path = os.path.join(base_path, doc_filename)
            template_path = os.path.join(base_path, template_filename)
            
            # Checking file existence
            if not os.path.exists(doc_path):
                print(f"Error: The document file '{doc_path}' does not exist. Please try again.\n")
                continue
            if not os.path.exists(template_path):
                print(f"Error: The template file '{template_path}' does not exist. Please try again.\n")
                continue

            # Extract content from the .docx file using the docx library
            print("Extracting content from the Word document...(Hopefully)\n")
            content = extract_content(doc_path)
            
            # Generate HTML using the Jinja2 template funny because i am just using a library i did not create
            print("Doing really cool computer stuff...\nGenerating HTML file...\n")
            html = generate_html(content, template_path)
            
            # Saving the output HTML file
            output_path = os.path.join(base_path, "output.html")
            with open(output_path, "w", encoding="utf-8") as file:
                file.write(html)
            
            print(f"It probably worked :-) : '{output_path}'\n")
        
        except Exception as e:
            print(f"An error occurred (99% it was mine ): {e}. Please try again.\n")
