import os
import cv2
import pandas as pd

def delete_old_data(folder):
    """Delete previously generated certificate images."""
    # Check if the output folder exists; if not, create it
    if not os.path.exists(folder):
        os.makedirs(folder)
    else:
        # If the folder exists, remove all files in it
        for file in os.listdir(folder):
            os.remove(os.path.join(folder, file))

def load_names_from_excel(file_path):
    """Load names from an Excel file, skipping the first row."""
    # Read the Excel file and skip the first row
    df = pd.read_excel(file_path, header=None, skiprows=1)
    # Extract names from the first column and convert them to a list
    return df.iloc[:, 0].tolist()

def generate_certificate(name, template_path, output_folder):
    """Generate a certificate image with the given name."""
    # Load the certificate template image
    certificate_template_image = cv2.imread(template_path)

    # Load the font
    font = cv2.FONT_HERSHEY_SIMPLEX

    # Set the font scale and thickness
    font_scale = 4.0
    font_thickness = 9  # Boldness

    # Get the text size
    (text_width, text_height), _ = cv2.getTextSize(name, font, font_scale, font_thickness)

    # Calculate the position to center align the text
    text_x = 1000 - text_width // 2
    text_y = 515 + text_height // 2

    # Set the font color (teal: #006C89)
    font_color = (137, 108, 0)

    # Put the text on the image
    cv2.putText(certificate_template_image, name, (text_x, text_y), font, font_scale, font_color, font_thickness, cv2.LINE_AA)

    # Save the generated certificate image
    output_path = os.path.join(output_folder, f"{name}.png")
    cv2.imwrite(output_path, certificate_template_image)
    print(f"Certificate generated for {name}")

def generate_certificates(names, template_path, output_folder):
    """Generate certificates for a list of names."""
    # Iterate over each name and generate a certificate for it
    for index, name in enumerate(names):
        generate_certificate(name, template_path, output_folder)
        print(f"Processing {index + 1} / {len(names)}")

def main():
    # Paths and settings
    excel_file = "registered.xlsx"  # Path to the Excel file containing names
    certificate_template = "certificate ko template.png"  # Path to the certificate template image
    output_folder = "Certificate Banyo/"  # Output folder for generated certificate images

    # Delete old certificate images (if any) from the output folder
    delete_old_data(output_folder)

    # Load names from Excel file
    names = load_names_from_excel(excel_file)

    # Generate certificates for the loaded names
    generate_certificates(names, certificate_template, output_folder)

if __name__ == '__main__':
    main()
