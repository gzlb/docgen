from pathlib import Path
import pandas as pd
from docx import Document
import os

def read_parameters(file_path: Path) -> dict[str, str]:
    """Read parameters from an Excel file."""
    df = pd.read_excel(file_path)
    return dict(zip(df['Parameter Name'], df['Value']))

def update_word_template(template_path: Path, output_path: Path, parameters: dict[str, str]) -> None:
    """Replace placeholders in a Word template with parameter values."""
    doc = Document(template_path)
    for paragraph in doc.paragraphs:
        for key, value in parameters.items():
            if f"{{{key}}}" in paragraph.text:
                paragraph.text = paragraph.text.replace(f"{{{key}}}", str(value))
    doc.save(output_path)

def make_file_read_only(file_path: Path) -> None:
    """Set file to read-only."""
    os.chmod(file_path, 0o444)

def main() -> None:
    """Main script function."""
    # Define file paths
    base_dir = Path(__file__).resolve().parent.parent
    template_path = base_dir / "template/template.docx"
    excel_path = base_dir / "data/parameters.xlsx"
    output_path = base_dir / "template/output.docx"
    
    # Process files
    parameters = read_parameters(excel_path)
    update_word_template(template_path, output_path, parameters)
    make_file_read_only(output_path)
    print(f"Document generated and saved to {output_path}")

if __name__ == "__main__":
    main()
  