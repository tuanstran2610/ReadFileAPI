import os
import re
import tempfile
import cv2
import pytesseract as pyt
import fitz
from docx import Document
from docx.oxml.ns import qn
from flask import Flask, request, jsonify
from pdf2image import convert_from_path
from openpyxl import load_workbook
from pptx import Presentation
from PIL import Image
import io
from zipfile import ZipFile

app = Flask(__name__)

# Constants
TESSERACT_PATH = r"C:/Program Files/Tesseract-OCR/tesseract.exe"
SUPPORTED_IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg'}
TEXT_EXTENSIONS = {'.txt', '.docx', '.doc', '.pdf', '.xlsx', '.xls', '.pptx', '.ppt'}

# Configure Tesseract
pyt.pytesseract.tesseract_cmd = TESSERACT_PATH


def is_image_file(filepath: str) -> bool:
    """Check if a PDF file contains only images (no selectable text)."""
    doc = fitz.open(filepath)
    for page in doc:
        if page.get_text().strip():
            return False
    return True


def clean_text(raw_text: str) -> str:
    """Remove stray newlines and extra whitespace from text."""
    return re.sub(r'(?<!\n)\n(?!\n)', ' ', raw_text.strip())


def extract_text_with_ocr(file_path: str) -> str:
    """Extract text from image-based PDFs or image files using OCR."""
    text = ""
    if file_path.lower().endswith('.pdf'):
        images = convert_from_path(file_path)
        for image in images:
            temp_path = tempfile.mktemp(suffix='.png')
            image.save(temp_path, 'PNG')
            img = cv2.imread(temp_path)
            text += pyt.image_to_string(img, lang="eng") + "\n"
            os.unlink(temp_path)
    elif file_path.lower().endswith(tuple(SUPPORTED_IMAGE_EXTENSIONS)):
        img = cv2.imread(file_path)
        text = pyt.image_to_string(img, lang="eng")
    return text.strip()


def extract_images_from_docx(file_path: str) -> list:
    """Extract images from a DOCX file."""
    images = []
    try:
        with ZipFile(file_path) as docx_zip:
            for file_info in docx_zip.infolist():
                if file_info.filename.startswith('word/media/'):
                    with docx_zip.open(file_info) as file:
                        image_data = file.read()
                        image = Image.open(io.BytesIO(image_data))
                        if image.format in ['PNG', 'JPEG']:
                            temp_path = tempfile.mktemp(suffix='.png')
                            image.save(temp_path, 'PNG')
                            images.append(temp_path)
    except Exception as e:
        print(f"Error extracting images from DOCX: {str(e)}")
    return images


def extract_images_from_pptx(file_path: str) -> list:
    """Extract images from a PPTX file."""
    images = []
    try:
        presentation = Presentation(file_path)
        for slide in presentation.slides:
            for shape in slide.shapes:
                if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                    image = shape.image
                    image_data = image.blob
                    image_file = Image.open(io.BytesIO(image_data))
                    if image_file.format in ['PNG', 'JPEG']:
                        temp_path = tempfile.mktemp(suffix='.png')
                        image_file.save(temp_path, 'PNG')
                        images.append(temp_path)
    except Exception as e:
        print(f"Error extracting images from PPTX: {str(e)}")
    return images


def read_docx(file_path: str) -> str:
    """Read text and extract text from images in a DOCX file."""
    try:
        # Read text
        doc = Document(file_path)
        text = "\n".join(para.text for para in doc.paragraphs)

        # Extract text from images
        image_paths = extract_images_from_docx(file_path)
        image_text = ""
        for image_path in image_paths:
            img = cv2.imread(image_path)
            image_text += pyt.image_to_string(img, lang="eng") + "\n"
            os.unlink(image_path)

        return "\n".join([text, image_text]).strip()
    except Exception as e:
        return f"Error reading DOCX file: {str(e)}"


def read_text_file(file_path: str) -> str:
    """Read text from a plain text file."""
    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            return file.read()
    except Exception as e:
        return f"Error reading text file: {str(e)}"


def read_pdf_text(file_path: str) -> str:
    """Read text from a text-based PDF file."""
    try:
        doc = fitz.open(file_path)
        text = "\n".join(page.get_text() for page in doc)
        return text
    except Exception as e:
        return f"Error reading PDF file: {str(e)}"


def read_xlsx(file_path: str) -> str:
    """Read text from an Excel (.xlsx) file."""
    try:
        workbook = load_workbook(file_path, read_only=True)
        text = []
        for sheet in workbook:
            for row in sheet.rows:
                for cell in row:
                    if cell.value:
                        text.append(str(cell.value))
        return "\n".join(text)
    except Exception as e:
        return f"Error reading XLSX file: {str(e)}"


def read_pptx(file_path: str) -> str:
    """Read text and extract text from images in a PowerPoint (.pptx) file."""
    try:
        presentation = Presentation(file_path)
        text = []
        for slide in presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text:
                    text.append(shape.text)

        # Extract text from images
        image_paths = extract_images_from_pptx(file_path)
        image_text = ""
        for image_path in image_paths:
            img = cv2.imread(image_path)
            image_text += pyt.image_to_string(img, lang="eng") + "\n"
            os.unlink(image_path)

        return "\n".join([*text, image_text]).strip()
    except Exception as e:
        return f"Error reading PPTX file: {str(e)}"


@app.route('/read-file', methods=['POST'])
def read_file():
    """API endpoint to read content from various file types."""
    data = request.get_json()
    file_path = data.get('filePath')

    if not file_path or not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 400

    file_ext = os.path.splitext(file_path)[1].lower()

    try:
        if file_ext in SUPPORTED_IMAGE_EXTENSIONS or (file_ext == '.pdf' and is_image_file(file_path)):
            content = clean_text(extract_text_with_ocr(file_path))
        elif file_ext == '.docx':
            content = clean_text(read_docx(file_path))
        elif file_ext == '.txt':
            content = clean_text(read_text_file(file_path))
        elif file_ext == '.pdf':
            content = clean_text(read_pdf_text(file_path))
        elif file_ext == '.xlsx':
            content = clean_text(read_xlsx(file_path))
        elif file_ext == '.pptx':
            content = clean_text(read_pptx(file_path))
        else:
            return jsonify({"error": "Unsupported file type"}), 400

        return jsonify({"file_content": content}), 200

    except Exception as e:
        return jsonify({"error": f"Error processing file: {str(e)}"}), 500


if __name__ == '__main__':
    app.run(debug=True)
