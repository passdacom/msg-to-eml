#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EML to MSG Converter
표준 .eml 파일을 Microsoft Outlook .msg 형식으로 변환

기술: Windows에서 pywin32를 통한 Outlook COM 사용
요구사항: Windows + Microsoft Outlook 설치

참고: COM 객체는 스레드별로 초기화해야 함
"""

import os
import sys
import email
import tempfile
from email import policy
from email.parser import BytesParser
from datetime import datetime
from pathlib import Path
from typing import Optional, List
import logging

logger = logging.getLogger(__name__)

# Windows/Outlook 체크
OUTLOOK_AVAILABLE = False
OUTLOOK_ERROR = None

try:
    import win32com.client
    import pythoncom
    OUTLOOK_AVAILABLE = True
except ImportError as e:
    OUTLOOK_ERROR = "pywin32가 설치되지 않았습니다. Windows에서만 지원됩니다."
except Exception as e:
    OUTLOOK_ERROR = str(e)


def check_outlook_available() -> tuple:
    """Outlook 사용 가능 여부 확인"""
    if not OUTLOOK_AVAILABLE:
        return False, OUTLOOK_ERROR
    
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        pythoncom.CoUninitialize()
        return True, None
    except Exception as e:
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return False, f"Outlook을 시작할 수 없습니다: {e}"


class EMLtoMSGConverter:
    """EML 파일을 MSG 형식으로 변환하는 클래스"""
    
    def __init__(self, verbose: bool = False):
        self.verbose = verbose
        # 초기 가용성 확인 (메인 스레드에서)
        self.available, self.error = check_outlook_available()
        if self.available:
            self.log("Outlook 연결 성공")
    
    def log(self, message: str):
        """상세 모드에서만 메시지 출력"""
        if self.verbose:
            print(f"  [정보] {message}")
        logger.debug(message)
    
    def is_available(self) -> tuple:
        """변환 가능 여부 확인"""
        return self.available, self.error
    
    def convert_file(self, eml_path: str, output_path: str = None) -> str:
        """
        단일 EML 파일을 MSG로 변환
        
        Args:
            eml_path: 입력 .eml 파일 경로
            output_path: 출력 .msg 파일 경로 (None이면 자동 생성)
            
        Returns:
            생성된 .msg 파일 경로
        """
        if not OUTLOOK_AVAILABLE:
            raise RuntimeError(f"EML→MSG 변환을 사용할 수 없습니다: {OUTLOOK_ERROR}")
        
        eml_path = Path(eml_path)
        
        if not eml_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {eml_path}")
        
        if not eml_path.suffix.lower() == '.eml':
            raise ValueError(f"EML 파일이 아닙니다: {eml_path}")
        
        # 출력 경로 결정
        if output_path is None:
            output_path = eml_path.with_suffix('.msg')
        else:
            output_path = Path(output_path)
        
        self.log(f"변환 중: {eml_path.name}")
        
        # EML 파일 파싱
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        
        # COM 초기화 (현재 스레드에서)
        pythoncom.CoInitialize()
        
        try:
            # Outlook 인스턴스 생성 (현재 스레드에서)
            outlook = win32com.client.Dispatch("Outlook.Application")
            
            # MSG 파일 생성
            self._create_msg_via_outlook(outlook, msg, output_path)
            
            self.log(f"저장됨: {output_path.name}")
            
            return str(output_path)
            
        finally:
            # COM 정리 (현재 스레드에서)
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _create_msg_via_outlook(self, outlook, email_msg, output_path: Path):
        """Outlook COM을 사용하여 MSG 파일 생성"""
        
        # 새 메일 아이템 생성
        mail_item = outlook.CreateItem(0)  # 0 = olMailItem
        
        # 제목 설정
        subject = email_msg.get('Subject', '')
        mail_item.Subject = subject
        
        # 본문 설정
        body = ""
        html_body = ""
        
        if email_msg.is_multipart():
            for part in email_msg.walk():
                content_type = part.get_content_type()
                content_disposition = part.get_content_disposition()
                
                if content_disposition != 'attachment':
                    if content_type == 'text/plain' and not body:
                        try:
                            body = part.get_content()
                        except:
                            payload = part.get_payload(decode=True)
                            if payload:
                                body = payload.decode('utf-8', errors='replace')
                    elif content_type == 'text/html' and not html_body:
                        try:
                            html_body = part.get_content()
                        except:
                            payload = part.get_payload(decode=True)
                            if payload:
                                html_body = payload.decode('utf-8', errors='replace')
        else:
            try:
                if email_msg.get_content_type() == 'text/html':
                    html_body = email_msg.get_content()
                else:
                    body = email_msg.get_content()
            except:
                payload = email_msg.get_payload(decode=True)
                if payload:
                    body = payload.decode('utf-8', errors='replace')
        
        if html_body:
            mail_item.HTMLBody = html_body
        else:
            mail_item.Body = body or ""
        
        # 수신자 설정 (To)
        to_addr = email_msg.get('To', '')
        if to_addr:
            for addr in self._split_addresses(to_addr):
                if addr:
                    try:
                        recipient = mail_item.Recipients.Add(addr)
                        recipient.Type = 1  # olTo
                    except:
                        pass
        
        # CC 설정
        cc_addr = email_msg.get('Cc', '')
        if cc_addr:
            for addr in self._split_addresses(cc_addr):
                if addr:
                    try:
                        recipient = mail_item.Recipients.Add(addr)
                        recipient.Type = 2  # olCC
                    except:
                        pass
        
        # 첨부파일 처리
        if email_msg.is_multipart():
            for part in email_msg.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        payload = part.get_payload(decode=True)
                        if payload:
                            temp_dir = tempfile.mkdtemp()
                            temp_path = os.path.join(temp_dir, filename)
                            try:
                                with open(temp_path, 'wb') as f:
                                    f.write(payload)
                                mail_item.Attachments.Add(temp_path)
                            finally:
                                try:
                                    os.remove(temp_path)
                                    os.rmdir(temp_dir)
                                except:
                                    pass
        
        # MSG 파일로 저장
        mail_item.SaveAs(str(output_path), 3)  # 3 = olMSG
        
        self.log(f"MSG 파일 생성 완료: {output_path}")
    
    def _split_addresses(self, addr_str: str) -> List[str]:
        """여러 이메일 주소 분리"""
        if not addr_str:
            return []
        
        addresses = []
        current = ""
        in_quotes = False
        in_brackets = False
        
        for char in addr_str:
            if char == '"':
                in_quotes = not in_quotes
            elif char == '<':
                in_brackets = True
            elif char == '>':
                in_brackets = False
            elif char in ',;' and not in_quotes and not in_brackets:
                if current.strip():
                    addresses.append(current.strip())
                current = ""
                continue
            current += char
        
        if current.strip():
            addresses.append(current.strip())
        
        return addresses
    
    def convert_directory(self, directory: str, recursive: bool = False, 
                         output_dir: str = None) -> list:
        """디렉토리 내의 모든 EML 파일을 변환"""
        if not OUTLOOK_AVAILABLE:
            raise RuntimeError(f"EML→MSG 변환을 사용할 수 없습니다: {OUTLOOK_ERROR}")
        
        directory = Path(directory)
        
        if not directory.exists():
            raise FileNotFoundError(f"디렉토리를 찾을 수 없습니다: {directory}")
        
        if not directory.is_dir():
            raise ValueError(f"디렉토리가 아닙니다: {directory}")
        
        # EML 파일 검색
        if recursive:
            eml_files = list(directory.rglob('*.eml')) + list(directory.rglob('*.EML'))
        else:
            eml_files = list(directory.glob('*.eml')) + list(directory.glob('*.EML'))
        
        eml_files = list(set(eml_files))
        
        if not eml_files:
            print(f"EML 파일을 찾을 수 없습니다: {directory}")
            return []
        
        print(f"총 {len(eml_files)}개의 EML 파일을 발견했습니다.")
        
        converted = []
        errors = []
        
        for eml_file in eml_files:
            try:
                if output_dir:
                    output_path = Path(output_dir) / eml_file.with_suffix('.msg').name
                else:
                    output_path = eml_file.with_suffix('.msg')
                
                result = self.convert_file(str(eml_file), str(output_path))
                converted.append(result)
                print(f"✓ {eml_file.name}")
                
            except Exception as e:
                errors.append((str(eml_file), str(e)))
                print(f"✗ {eml_file.name}: {e}")
        
        print(f"\n변환 완료: {len(converted)}개 성공, {len(errors)}개 실패")
        
        return converted
