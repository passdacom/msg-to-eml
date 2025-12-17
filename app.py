#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MSG to EML Converter - Web Application
Flask 기반 웹 UI로 MSG 파일을 EML로 변환합니다.

실행: python app.py
접속: http://localhost:5000
"""

import os
import uuid
import zipfile
import tempfile
import shutil
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename

# 기존 변환기 import
from msg_to_eml import MSGtoEMLConverter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 제한

# 임시 파일 저장 디렉토리
UPLOAD_FOLDER = Path(tempfile.gettempdir()) / 'msg_to_eml_uploads'
CONVERTED_FOLDER = Path(tempfile.gettempdir()) / 'msg_to_eml_converted'

# 폴더 생성
UPLOAD_FOLDER.mkdir(exist_ok=True)
CONVERTED_FOLDER.mkdir(exist_ok=True)

# 변환 세션 저장 (메모리)
conversion_sessions = {}


def cleanup_old_files():
    """1시간 이상 된 파일 정리"""
    import time
    current_time = time.time()
    
    for folder in [UPLOAD_FOLDER, CONVERTED_FOLDER]:
        for file_path in folder.iterdir():
            if file_path.is_file():
                file_age = current_time - file_path.stat().st_mtime
                if file_age > 3600:  # 1시간
                    try:
                        file_path.unlink()
                    except:
                        pass


@app.route('/')
def index():
    """메인 페이지"""
    cleanup_old_files()
    return render_template('index.html')


@app.route('/api/convert', methods=['POST'])
def convert():
    """MSG 파일을 EML로 변환"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '파일이 없습니다'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'success': False, 'error': '파일명이 없습니다'}), 400
    
    if not file.filename.lower().endswith('.msg'):
        return jsonify({'success': False, 'error': 'MSG 파일만 지원됩니다'}), 400
    
    try:
        # 고유 ID 생성
        file_id = str(uuid.uuid4())
        
        # 파일 저장
        original_filename = secure_filename(file.filename)
        msg_path = UPLOAD_FOLDER / f"{file_id}_{original_filename}"
        file.save(str(msg_path))
        
        # 변환
        converter = MSGtoEMLConverter(verbose=False)
        eml_filename = original_filename.rsplit('.', 1)[0] + '.eml'
        eml_path = CONVERTED_FOLDER / f"{file_id}_{eml_filename}"
        
        converter.convert_file(str(msg_path), str(eml_path))
        
        # 원본 삭제
        msg_path.unlink()
        
        # 세션에 저장
        conversion_sessions[file_id] = {
            'original_name': original_filename,
            'eml_name': eml_filename,
            'eml_path': str(eml_path)
        }
        
        return jsonify({
            'success': True,
            'file_id': file_id,
            'original_name': original_filename,
            'eml_name': eml_filename
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/download/<file_id>')
def download(file_id):
    """변환된 EML 파일 다운로드"""
    if file_id not in conversion_sessions:
        return jsonify({'success': False, 'error': '파일을 찾을 수 없습니다'}), 404
    
    session = conversion_sessions[file_id]
    eml_path = Path(session['eml_path'])
    
    if not eml_path.exists():
        return jsonify({'success': False, 'error': '파일이 만료되었습니다'}), 404
    
    return send_file(
        str(eml_path),
        as_attachment=True,
        download_name=session['eml_name'],
        mimetype='message/rfc822'
    )


@app.route('/api/download-all', methods=['POST'])
def download_all():
    """여러 파일을 ZIP으로 다운로드"""
    data = request.get_json()
    file_ids = data.get('file_ids', [])
    
    if not file_ids:
        return jsonify({'success': False, 'error': '파일이 없습니다'}), 400
    
    try:
        # ZIP 파일 생성
        zip_id = str(uuid.uuid4())
        zip_path = CONVERTED_FOLDER / f"{zip_id}_converted_emails.zip"
        
        with zipfile.ZipFile(str(zip_path), 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_id in file_ids:
                if file_id in conversion_sessions:
                    session = conversion_sessions[file_id]
                    eml_path = Path(session['eml_path'])
                    if eml_path.exists():
                        zipf.write(str(eml_path), session['eml_name'])
        
        return send_file(
            str(zip_path),
            as_attachment=True,
            download_name='converted_emails.zip',
            mimetype='application/zip'
        )
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/clear', methods=['POST'])
def clear_session():
    """세션 초기화"""
    data = request.get_json()
    file_ids = data.get('file_ids', [])
    
    for file_id in file_ids:
        if file_id in conversion_sessions:
            session = conversion_sessions[file_id]
            try:
                Path(session['eml_path']).unlink(missing_ok=True)
            except:
                pass
            del conversion_sessions[file_id]
    
    return jsonify({'success': True})


if __name__ == '__main__':
    print("\n" + "="*50)
    print("  MSG to EML Converter")
    print("  http://localhost:5001 에서 접속하세요")
    print("="*50 + "\n")
    app.run(debug=True, host='0.0.0.0', port=5001)
