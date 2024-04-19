import os
import shutil

def main():
    # Define the paths for the input and output folders
    input_folder = "docs-in"
    output_folder = "products"

    # Iterate over each file in the input folder
    for filename in os.listdir(input_folder):
        # Check if the file is a .docx file
        if filename.endswith('.docx'):
            # Create the full paths for the input and output files
            input_file_path = os.path.join(input_folder, filename)
            output_file_path = os.path.join(output_folder, filename)

            # Copy the file from the input folder to the output folder
            shutil.copy(input_file_path, output_file_path)

            # Print a message indicating the file has been duplicated
            print(f"File duplicated: {filename}")

if __name__ == "__main__":
    main()