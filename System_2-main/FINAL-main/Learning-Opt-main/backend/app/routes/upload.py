# backend/app/routes/upload.py
from flask import Blueprint, request, jsonify
import os
import json

bp = Blueprint('upload', __name__, url_prefix='/upload')

@bp.route('/excel', methods=['POST'])
def upload_excel():
    data = request.get_json()
    rows = data.get('rows', [])

    upload_path = os.path.join(os.path.dirname(__file__), '..', 'static', 'excel')
    os.makedirs(upload_path, exist_ok=True)
    file_path = os.path.join(upload_path, 'uploaded_data.json')

    if os.path.exists(file_path):
        os.remove(file_path)

    with open(file_path, 'w', encoding='utf-8') as f:
        json.dump(rows, f, indent=2)

    return jsonify({'message': 'Excel data uploaded successfully', 'rowCount': len(rows)})


@bp.route('/excel', methods=['GET'])
def get_uploaded_excel():
    file_path = os.path.join(os.path.dirname(__file__), '..', 'static', 'excel', 'uploaded_data.json')

    if not os.path.exists(file_path):
        return jsonify({'rows': []})

    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    return jsonify({'rows': data})
