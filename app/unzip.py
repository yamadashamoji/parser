import zipfile
import os
import logging

logging.basicConfig(filename='unzip.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def extract_zip(zip_path, extract_path):
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_path)
        logging.info(f"Successfully extracted {zip_path} to {extract_path}")
    except zipfile.BadZipFile:
        logging.error(f"The file {zip_path} is not a valid ZIP file")
        raise
    except Exception as e:
        logging.error(f"An error occurred while extracting {zip_path}: {str(e)}")
        raise