python get-pip.py
# if isinstance(date, datetime):
#         formatted_date = date.strftime("%d-%m-%Y")  # Convert to dd-mm-yyyy
#     else:
#         # Attempt to parse the date if it's not a datetime object
#         try:
#             parsed_date = datetime.strptime(str(date), "%Y-%m-%d")  # Adjust parsing if needed
#             formatted_date = parsed_date.strftime("%d-%m-%Y")
#         except ValueError:
#             formatted_date = str(date)  # Fallback to raw string if parsing fails


# pip install pandas pillow openpyxl

# import os
# from PIL import Image, ImageDraw, ImageFont
# import pandas as pd

# # Paths to resources
# CERTIFICATE_TEMPLATE_PATH = r"C:\Users\Dell\Desktop\Hexamad\Sample certificate.jpg"
# EXCEL_FILE_PATH = r"C:\Users\Dell\Desktop\Hexamad\hexamad.xlsx"
# OUTPUT_FOLDER_PATH = r"C:\Users\Dell\Desktop\Hexamad"

# # Create output directory if it doesn't exist
# os.makedirs(OUTPUT_FOLDER_PATH, exist_ok=True)

# def generate_certificates(template_path, excel_path, output_path):
#     # Load data from Excel
#     try:
#         data = pd.read_excel(excel_path)
#     except Exception as e:
#         print(f"Error loading Excel file: {e}")
#         return

#     if 'Name' not in data.columns or 'Date' not in data.columns:
#         print("Error: Excel file must contain 'Name' and 'Date' columns.")
#         return

#     # Open the certificate template
#     try:
#         template = Image.open(template_path)
#     except Exception as e:
#         print(f"Error loading template image: {e}")
#         return

#     # Define font settings
#     try:
#         font_name = ImageFont.truetype("arial.ttf", 40)
#         font_date = ImageFont.truetype("arial.ttf", 30)
#     except Exception as e:
#         print(f"Error loading font: {e}")
#         return

#     for index, row in data.iterrows():
#         name = row['Name']
#         date = row['Date']

#         # Create a copy of the template
#         certificate = template.copy()
#         draw = ImageDraw.Draw(certificate)

#         # Define positions for name and date
#         name_position = (400, 300)  # Adjust this to match your template
#         date_position = (400, 400)  # Adjust this to match your template

#         # Add text to the certificate
#         draw.text(name_position, name, fill="black", font=font_name)
#         draw.text(date_position, str(date), fill="black", font=font_date)

#         # Save the generated certificate
#         output_file_path = os.path.join(output_path, f"{name}_certificate.png")
#         certificate.save(output_file_path)
#         print(f"Certificate saved for {name} at {output_file_path}")

# if __name__ == "__main__":
#     generate_certificates(CERTIFICATE_TEMPLATE_PATH, EXCEL_FILE_PATH, OUTPUT_FOLDER_PATH)
