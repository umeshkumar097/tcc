import os
import zipfile
import pandas as pd
from io import BytesIO
from PIL import Image
import shutil

def unzip_and_organize_files(zip_file_path: str, destination_dir: str):
    """
    Unzips the provided ZIP file into the destination directory.
    Returns the path to the unzipped content.
    """
    os.makedirs(destination_dir, exist_ok=True)
    with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
        zip_ref.extractall(destination_dir)
    return destination_dir

def create_output_zip(source_dir: str, output_zip_path: str):
    """
    Creates a ZIP file from the contents of the source directory.
    """
    shutil.make_archive(os.path.splitext(output_zip_path)[0], 'zip', source_dir)
    # make_archive creates file with .zip extension, so just return output_zip_path
    return output_zip_path

def clean_temp_dirs(directory: str):
    """
    Removes the specified directory and all its contents.
    """
    if os.path.exists(directory) and os.path.isdir(directory):
        try:
            shutil.rmtree(directory)
            print(f"Cleaned up temporary directory: {directory}")
        except Exception as e:
            print(f"Error cleaning up directory {directory}: {e}")

def get_image_from_bytes(image_bytes: bytes) -> Image.Image:
    """Loads a PIL Image from bytes."""
    return Image.open(BytesIO(image_bytes))

def get_excel_df(excel_file_buffer) -> pd.DataFrame:
    """Reads an Excel file buffer into a pandas DataFrame."""
    return pd.read_excel(excel_file_buffer)
