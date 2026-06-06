import os
from docxtpl import DocxTemplate

def generate_from_word(template_path, output_path, data):
    """
    Generates a document from a Word template using docxtpl.
    """
    doc = DocxTemplate(template_path)
    doc.render(data)
    doc.save(output_path)
    return output_path

# Example Usage (for testing)
if __name__ == "__main__":
    print("Word generator logic ready.")
