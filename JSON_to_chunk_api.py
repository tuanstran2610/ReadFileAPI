# import os
# import json
# import fitz
# import re
# from flask import Flask, request, jsonify
# from langchain.text_splitter import RecursiveCharacterTextSplitter
# from langchain_experimental.text_splitter import SemanticChunker
# from langchain_community.embeddings import HuggingFaceEmbeddings
# from qdrant_client import QdrantClient
# from qdrant_client.http.models import PointStruct, VectorParams, Distance
# import uuid
# from docx import Document
# from pdf2image import convert_from_path
# import easyocr
# import tempfile
#
# app = Flask(__name__)
#
# FILE_EXTENSIONS = [".pdf", ".docx", ".txt", ".jpg", ".png", ".jpeg"]
# GENERAL_COLLECTION_NAME = "general_documents"
#
#
# def check_image(filepath):
#     doc = fitz.open(filepath)
#     for page in doc:
#         text = page.get_text()
#         if text.strip():
#             doc.close()
#             return False
#     doc.close()
#     return True
#
#
# def clean_text(raw_text):
#     return re.sub(r'(?<!\n)\n(?!\n)', ' ', raw_text.strip())
#
#
# def preprocess_text(text):
#     text = re.sub(r'(?:Page|Trang)?\s*-?\s*\d+\s*-?', '', text, flags=re.IGNORECASE)
#     text = re.sub(r"[^\w\s.,!?%\-–()]", "", text)
#     text = re.sub(r'\n{2,}', '\n', text)
#     text = re.sub(r'[ \t]+', ' ', text)
#     text = re.sub(r' +\n', '\n', text)
#     return text.strip()
#
#
# def extract_text_with_ocr(file_path):
#     reader = easyocr.Reader(['vi', 'en'])
#     text = ""
#     if file_path.lower().endswith('.pdf'):
#         images = convert_from_path(file_path)
#         for image in images:
#             temp_path = tempfile.mktemp(suffix='.png')
#             image.save(temp_path, 'PNG')
#             results = reader.readtext(temp_path, detail=0)
#             text += "\n".join(results) + "\n"
#             os.unlink(temp_path)
#     elif file_path.lower().endswith(('.png', '.jpg', '.jpeg')):
#         results = reader.readtext(file_path, detail=0)
#         text += "\n".join(results)
#     return preprocess_text(clean_text(text))
#
#
# def extract_text(file_path):
#     file_name = os.path.basename(file_path)
#     if file_path.lower().endswith(('.jpg', '.png', '.jpeg')) or (
#             file_path.lower().endswith('.pdf') and check_image(file_path)):
#         return extract_text_with_ocr(file_path), file_name
#     elif file_path.lower().endswith('.pdf'):
#         doc = fitz.open(file_path)
#         text = ""
#         for page in doc:
#             text += page.get_text()
#         doc.close()
#         return preprocess_text(clean_text(text)), file_name
#     elif file_path.lower().endswith('.txt'):
#         with open(file_path, 'r', encoding='utf-8') as f:
#             text = f.read()
#         return preprocess_text(clean_text(text)), file_name
#     elif file_path.lower().endswith('.docx'):
#         doc = Document(file_path)
#         text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
#         return preprocess_text(clean_text(text)), file_name
#     else:
#         raise ValueError(f"Unsupported file type: {file_path}")
#
#
# def filter_invalid_chunks(chunks, min_length=30):
#     filtered = []
#     for chunk in chunks:
#         cleaned = chunk.strip()
#         if len(cleaned) >= min_length and not re.fullmatch(r"[.?!,:;\"']+", cleaned):
#             filtered.append(cleaned)
#     return filtered
#
#
# def semantic_chunking(text, embed_model):
#     try:
#         base_splitter = RecursiveCharacterTextSplitter(
#             chunk_size=2000,
#             chunk_overlap=400,
#             separators=["\n\n", "\n", ".", "!", "?", ",", " ", ""]
#         )
#         base_chunks = base_splitter.split_text(text)
#         semantic_splitter = SemanticChunker(
#             embeddings=embed_model,
#             breakpoint_threshold_type="percentile",
#             breakpoint_threshold_amount=90
#         )
#         final_chunks = []
#         for chunk in base_chunks:
#             try:
#                 semantic_chunks = semantic_splitter.split_text(chunk)
#                 final_chunks.extend(semantic_chunks)
#             except Exception as e:
#                 print(f"Semantic split failed on chunk: {e}")
#                 final_chunks.append(chunk)
#         return filter_invalid_chunks(final_chunks)
#     except Exception as e:
#         print(f"Error during semantic chunking: {e}")
#         return []
#
#
# def store_in_qdrant(chunks, file_name, collection_name, form_data, embed_model, client):
#     try:
#         client.get_collection(collection_name)
#     except:
#         client.create_collection(
#             collection_name=collection_name,
#             vectors_config=VectorParams(
#                 size=embed_model.client.get_sentence_embedding_dimension(),
#                 distance=Distance.COSINE
#             )
#         )
#
#     embeddings = embed_model.embed_documents(chunks)
#     points = [
#         PointStruct(
#             id=str(uuid.uuid4()),
#             vector=embedding,
#             payload={
#                 "file_name": file_name,
#                 "chunk_id": i,
#                 "text": chunk,
#                 **(form_data or {})
#             }
#         )
#         for i, (chunk, embedding) in enumerate(zip(chunks, embeddings))
#     ]
#     client.upsert(collection_name=collection_name, points=points)
#
#
# @app.route('/store-documents', methods=['POST'])
# def store_documents():
#     try:
#         data = request.get_json()
#         if not data:
#             return jsonify({"error": "No JSON data provided"}), 400
#
#         loai_phieu = data.get('loai_phieu')
#         form_data = data.get('formData', {})
#         files = data.get('files', [])
#
#         if not loai_phieu:
#             return jsonify({"error": "Missing loai_phieu"}), 400
#         if not files:
#             return jsonify({"error": "No files provided"}), 400
#
#         embed_model = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
#         client = QdrantClient(url="http://localhost:6333")
#         results = []
#
#         for file_info in files:
#             file_path = file_info.get('path')
#             file_name = file_info.get('file_name')
#             file_type = file_info.get('file_type')
#
#             if not file_path or not file_name or not file_type:
#                 results.append({
#                     "file_name": file_name or "unknown",
#                     "status": "error",
#                     "message": "Missing file information"
#                 })
#                 continue
#
#             if not os.path.exists(file_path):
#                 results.append({
#                     "file_name": file_name,
#                     "status": "error",
#                     "message": f"File not found: {file_path}"
#                 })
#                 continue
#
#             if not any(file_path.lower().endswith(ext) for ext in FILE_EXTENSIONS):
#                 results.append({
#                     "file_name": file_name,
#                     "status": "error",
#                     "message": f"Unsupported file type. Supported extensions: {', '.join(FILE_EXTENSIONS)}"
#                 })
#                 continue
#
#             try:
#                 text, extracted_file_name = extract_text(file_path)
#                 if not text:
#                     results.append({
#                         "file_name": file_name,
#                         "status": "error",
#                         "message": "No text extracted"
#                     })
#                     continue
#
#                 chunks = semantic_chunking(text, embed_model)
#                 if not chunks:
#                     results.append({
#                         "file_name": file_name,
#                         "status": "error",
#                         "message": "No chunks created"
#                     })
#                     continue
#
#                 # Store in specific collection
#                 store_in_qdrant(chunks, file_name, loai_phieu, form_data, embed_model, client)
#                 # Store in general collection
#                 store_in_qdrant(chunks, file_name, GENERAL_COLLECTION_NAME, form_data, embed_model, client)
#
#                 results.append({
#                     "file_name": file_name,
#                     "status": "success",
#                     "message": f"Processed: {len(chunks)} chunks created and stored",
#                     "content": chunks
#                 })
#
#             except Exception as e:
#                 results.append({
#                     "file_name": file_name,
#                     "status": "error",
#                     "message": f"Error processing file: {str(e)}"
#                 })
#
#         return jsonify({
#             "status": "completed",
#             "results": results
#         }), 200
#
#     except Exception as e:
#         return jsonify({"error": f"Server error: {str(e)}"}), 500
#
#
# if __name__ == '__main__':
#     app.run(debug=True, host='0.0.0.0', port=5000)


import os
import json
import fitz
import re
from flask import Flask, request, jsonify
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_huggingface import HuggingFaceEmbeddings
from qdrant_client import QdrantClient
from qdrant_client.http.models import PointStruct, VectorParams, Distance
import uuid
from docx import Document
from pdf2image import convert_from_path
import easyocr
import tempfile
import torch
from concurrent.futures import ThreadPoolExecutor


app = Flask(__name__)

FILE_EXTENSIONS = [".pdf", ".docx", ".txt", ".jpg", ".png", ".jpeg"]
GENERAL_COLLECTION_NAME = "general_documents"
reader = easyocr.Reader(['vi', 'en'], gpu=False)  # Global easyocr reader, CPU
embed_model = HuggingFaceEmbeddings(
    model_name="sentence-transformers/all-MiniLM-L6-v2",
    model_kwargs={"device": "cpu"}
)
client = QdrantClient(url="http://localhost:6333")
created_collections = set()  # Cache for collection existence


def check_image(filepath):
    doc = fitz.open(filepath)
    for page in doc:
        text = page.get_text()
        if text.strip():
            doc.close()
            return False
    doc.close()
    return True


def clean_text(raw_text):
    return re.sub(r'(?<!\n)\n(?!\n)', ' ', raw_text.strip())


def preprocess_text(text):
    text = re.sub(r'(?:Page|Trang)?\s*-?\s*\d+\s*-?', '', text, flags=re.IGNORECASE)
    text = re.sub(r"[^\w\s.,!?%\-–()]", "", text)
    text = re.sub(r'\n{2,}', '\n', text)
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r' +\n', '\n', text)
    return text.strip()


def extract_text_with_ocr(file_path):
    text = ""
    if file_path.lower().endswith('.pdf'):
        images = convert_from_path(file_path, dpi=150, grayscale=True)
        for image in images:
            temp_path = tempfile.mktemp(suffix='.png')
            image.save(temp_path, 'PNG')
            results = reader.readtext(temp_path, detail=0, low_text=0.3)
            text += "\n".join(results) + "\n"
            os.unlink(temp_path)
    elif file_path.lower().endswith(('.png', '.jpg', '.jpeg')):
        results = reader.readtext(file_path, detail=0, low_text=0.3)
        text += "\n".join(results)
    return preprocess_text(clean_text(text))


def extract_text(file_path):
    file_name = os.path.basename(file_path)
    if file_path.lower().endswith(('.jpg', '.png', '.jpeg')) or (
            file_path.lower().endswith('.pdf') and check_image(file_path)):
        return extract_text_with_ocr(file_path), file_name
    elif file_path.lower().endswith('.pdf'):
        doc = fitz.open(file_path)
        text = ""
        for page in doc:
            text += page.get_text()
        doc.close()
        return preprocess_text(clean_text(text)), file_name
    elif file_path.lower().endswith('.txt'):
        with open(file_path, 'r', encoding='utf-8') as f:
            text = f.read()
        return preprocess_text(clean_text(text)), file_name
    elif file_path.lower().endswith('.docx'):
        doc = Document(file_path)
        text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        return preprocess_text(clean_text(text)), file_name
    else:
        raise ValueError(f"Unsupported file type: {file_path}")


def filter_invalid_chunks(chunks, min_length=30):
    filtered = []
    for chunk in chunks:
        cleaned = chunk.strip()
        if len(cleaned) >= min_length and not re.fullmatch(r"[.?!,:;\"']+", cleaned):
            filtered.append(cleaned)
    return filtered


def semantic_chunking(text, embed_model):
    try:
        base_splitter = RecursiveCharacterTextSplitter(
            chunk_size=500,
            chunk_overlap=100,
            separators=["\n\n", "\n", ".", "!", "?", ",", " ", ""]
        )
        chunks = base_splitter.split_text(text)
        return filter_invalid_chunks(chunks, min_length=30)
    except Exception as e:
        print(f"Error during chunking: {e}")
        return []


def ensure_collection(client, collection_name, embed_model):
    if collection_name not in created_collections:
        try:
            client.get_collection(collection_name)
            created_collections.add(collection_name)
        except:
            # Get embedding dimension by embedding a dummy text
            embedding = embed_model.embed_query("test")
            client.create_collection(
                collection_name=collection_name,
                vectors_config=VectorParams(
                    size=len(embedding),
                    distance=Distance.COSINE
                )
            )
            created_collections.add(collection_name)


def store_in_qdrant(chunks, file_name, collection_name, form_data, embed_model, client):
    batch_size = 16
    embeddings = []
    for i in range(0, len(chunks), batch_size):
        batch = chunks[i:i + batch_size]
        embeddings.extend(embed_model.embed_documents(batch))

    points = [
        PointStruct(
            id=str(uuid.uuid4()),
            vector=embedding,
            payload={
                "file_name": file_name,
                "chunk_id": i,
                "text": chunk,
                **(form_data or {})
            }
        )
        for i, (chunk, embedding) in enumerate(zip(chunks, embeddings))
    ]
    client.upsert(collection_name=collection_name, points=points)


def process_single_file(file_info, loai_phieu, form_data, embed_model, client):
    file_path = file_info.get('path')
    file_name = file_info.get('file_name')
    file_type = file_info.get('file_type')

    if not file_path or not file_name or not file_type:
        return {
            "file_name": file_name or "unknown",
            "status": "error",
            "message": "Missing file information"
        }

    if not os.path.exists(file_path):
        return {
            "file_name": file_name,
            "status": "error",
            "message": f"File not found: {file_path}"
        }

    if not any(file_path.lower().endswith(ext) for ext in FILE_EXTENSIONS):
        return {
            "file_name": file_name,
            "status": "error",
            "message": f"Unsupported file type. Supported extensions: {', '.join(FILE_EXTENSIONS)}"
        }

    try:
        text, extracted_file_name = extract_text(file_path)
        if not text:
            return {
                "file_name": file_name,
                "status": "error",
                "message": "No text extracted"
            }

        chunks = semantic_chunking(text, embed_model)
        if not chunks:
            return {
                "file_name": file_name,
                "status": "error",
                "message": "No chunks created"
            }

        ensure_collection(client, loai_phieu, embed_model)
        ensure_collection(client, GENERAL_COLLECTION_NAME, embed_model)
        store_in_qdrant(chunks, file_name, loai_phieu, form_data, embed_model, client)
        store_in_qdrant(chunks, file_name, GENERAL_COLLECTION_NAME, form_data, embed_model, client)

        return {
            "file_name": file_name,
            "status": "success",
            "message": f"Processed: {len(chunks)} chunks created and stored",
            "content": chunks
        }
    except Exception as e:
        return {
            "file_name": file_name,
            "status": "error",
            "message": f"Error processing file: {str(e)}"
        }


@app.route('/store-documents', methods=['POST'])
def store_documents():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400

        loai_phieu = data.get('loai_phieu')
        form_data = data.get('formData', {})
        files = data.get('files', [])

        if not loai_phieu:
            return jsonify({"error": "Missing loai_phieu"}), 400
        if not files:
            return jsonify({"error": "No files provided"}), 400

        results = []
        for file_info in files:
            result = process_single_file(file_info, loai_phieu, form_data, embed_model, client)
            results.append(result)

        return jsonify({
            "status": "completed",
            "results": results
        }), 200

    except Exception as e:
        return jsonify({"error": f"Server error: {str(e)}"}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)


