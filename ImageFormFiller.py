import os
import shutil
from zipfile import ZipFile

# Paths
candidates_base_folder = r"C:\Users\Ansh Gupta\temp_project"
final_zip_path = r"C:\Users\Ansh Gupta\final_candidates.zip"
output_folder = r"C:\Users\Ansh Gupta\temp_filled_candidates"

os.makedirs(output_folder, exist_ok=True)

for candidate_name in os.listdir(candidates_base_folder):
    candidate_path = os.path.join(candidates_base_folder, candidate_name)
    if os.path.isdir(candidate_path):
        # Candidate folder inside temp output
        candidate_output_folder = os.path.join(output_folder, candidate_name)
        os.makedirs(candidate_output_folder, exist_ok=True)

        # --- Fill form ---
        photo_path = os.path.join(candidate_path, "photo.jpg")
        candidate_data = {
            "Name": candidate_name,
            # Add other fields if needed
        }

        form_filler = ImageFormFiller(template_image, mapping_data, font_path="arial.ttf")
        filled_pdf_path = form_filler.fill_and_save_pdf(
            output_folder=candidate_output_folder,
            candidate_data=candidate_data,
            candidate_srno=candidate_name,
            candidate_name=candidate_name,
            photo_path=photo_path
        )

        # --- Copy all documents including the photo ---
        for file_name in os.listdir(candidate_path):
            file_path = os.path.join(candidate_path, file_name)
            if os.path.isfile(file_path):
                shutil.copy(file_path, os.path.join(candidate_output_folder, file_name))

# --- Create single ZIP with all candidates ---
with ZipFile(final_zip_path, 'w') as zipf:
    for root, _, files in os.walk(output_folder):
        for file in files:
            file_path = os.path.join(root, file)
            arcname = os.path.relpath(file_path, output_folder)
            zipf.write(file_path, arcname)

print("ZIP created successfully at:", final_zip_path)
