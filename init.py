import os
from docx import Document

def create_products_folder():
    """
    Create a 'products' folder if it doesn't exist
    """
    if not os.path.exists("products"):
        os.makedirs("products")

def main():
    # Define the path for the input folder
    input_folder = "docs-in"

    # Iterate over each file in the input folder
    for filename in os.listdir(input_folder):
        # Check if the file is a .docx file
        if filename.endswith('.docx'):
            # Create the full path for the input file
            input_file_path = os.path.join(input_folder, filename)

            # Read the original document
            original_doc = Document(input_file_path)

            # Find highlight colors
            highlight_colors = set(run.font.highlight_color for paragraph in original_doc.paragraphs for run in paragraph.runs if run.font.highlight_color is not None)

            # Iterate over each highlight color found
            for color in highlight_colors:
                # Create a new document for each color
                new_doc = Document()

                # Iterate over each paragraph in the original document
                for original_paragraph in original_doc.paragraphs:
                    # Create a new paragraph in the new document
                    new_paragraph = new_doc.add_paragraph()

                    # Iterate over each run in the original paragraph
                    for original_run in original_paragraph.runs:
                        # Add the text from the original run to the new paragraph
                        new_run = new_paragraph.add_run(original_run.text)

                        # Apply highlighting only if the original run is highlighted with the current color
                        if original_run.font.highlight_color == color:
                            new_run.font.highlight_color = color

                # Construct the output file path
                output_file_name = f"{os.path.splitext(filename)[0]}-{color.name}.docx"
                output_file_path = os.path.join("products", output_file_name)

                # Save the duplicate document
                create_products_folder()
                new_doc.save(output_file_path)

                # Print a message indicating the duplicate document has been created
                print(f"Duplicate document created: {output_file_path}")

if __name__ == "__main__":
    main()
