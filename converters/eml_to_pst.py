#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EML to PST Converter
여러 .eml 파일을 하나의 Microsoft Outlook .pst 파일로 변환

기술: pywin32를 통한 Outlook COM 자동화
요구사항: Windows + Microsoft Outlook 설치
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
    OUTLOOK_ERROR = "pywin32가 설치되지 않았습니다. pip install pywin32"
except Exception as e:
    OUTLOOK_ERROR = str(e)


def check_outlook_available() -> tuple:
    """
    Outlook 사용 가능 여부 확인
    
    Returns:
        (available: bool, error_message: str or None)
    """
    if not OUTLOOK_AVAILABLE:
        return False, OUTLOOK_ERROR
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return True, None
    except Exception as e:
        return False, f"Outlook을 시작할 수 없습니다: {e}"


class EMLtoPSTConverter:
    """EML 파일들을 PST 파일로 변환하는 클래스"""
    
    def __init__(self, verbose: bool = False):
        self.verbose = verbose
        self.outlook = None
        self.namespace = None
    
    def log(self, message: str):
        """상세 모드에서만 메시지 출력"""
        if self.verbose:
            print(f"  [정보] {message}")
        logger.debug(message)
    
    def is_available(self) -> tuple:
        """Outlook 사용 가능 여부 확인"""
        return check_outlook_available()
    
    def _init_outlook(self):
        """Outlook 초기화"""
        if self.outlook is None:
            try:
                pythoncom.CoInitialize()
                self.outlook = win32com.client.Dispatch("Outlook.Application")
                self.namespace = self.outlook.GetNamespace("MAPI")
                self.log("Outlook 연결 성공")
            except Exception as e:
                raise RuntimeError(f"Outlook 초기화 실패: {e}")
    
    def _cleanup_outlook(self):
        """Outlook 정리"""
        try:
            if self.namespace:
                self.namespace = None
            if self.outlook:
                self.outlook = None
            pythoncom.CoUninitialize()
        except:
            pass
    
    def convert_files(self, eml_paths: List[str], pst_path: str, 
                      folder_name: str = "Imported Emails") -> str:
        """
        여러 EML 파일을 하나의 PST 파일로 변환
        
        Args:
            eml_paths: EML 파일 경로 리스트
            pst_path: 생성할 PST 파일 경로
            folder_name: PST 내 폴더 이름
            
        Returns:
            생성된 PST 파일 경로
        """
        available, error = self.is_available()
        if not available:
            raise RuntimeError(f"Outlook을 사용할 수 없습니다: {error}")
        
        pst_path = Path(pst_path)
        
        try:
            self._init_outlook()
            
            # PST 파일 생성
            self.log(f"PST 파일 생성: {pst_path}")
            
            # 기존 파일 삭제
            if pst_path.exists():
                os.remove(pst_path)
            
            # 새 PST 스토어 추가
            self.namespace.AddStoreEx(str(pst_path), 1)  # 1 = olStoreDefault
            
            # 새로 추가된 PST 스토어 찾기
            pst_store = None
            for store in self.namespace.Stores:
                if str(pst_path).lower() in store.FilePath.lower():
                    pst_store = store
                    break
            
            if not pst_store:
                raise RuntimeError("PST 스토어를 찾을 수 없습니다")
            
            # 루트 폴더 가져오기
            root_folder = pst_store.GetRootFolder()
            
            # 새 폴더 생성
            try:
                import_folder = root_folder.Folders.Add(folder_name)
            except:
                import_folder = root_folder.Folders[folder_name]
            
            self.log(f"폴더 생성: {folder_name}")
            
            # 각 EML 파일 처리
            success_count = 0
            error_count = 0
            
            for eml_path in eml_paths:
                try:
                    self._import_eml_to_folder(eml_path, import_folder)
                    success_count += 1
                    self.log(f"추가됨: {Path(eml_path).name}")
                except Exception as e:
                    error_count += 1
                    logger.error(f"EML 추가 실패 {eml_path}: {e}")
            
            self.log(f"변환 완료: {success_count}개 성공, {error_count}개 실패")
            
            # PST 스토어 분리 (선택적)
            # self.namespace.RemoveStore(root_folder)
            
            return str(pst_path)
            
        finally:
            self._cleanup_outlook()
    
    def _import_eml_to_folder(self, eml_path: str, folder):
        """
        단일 EML 파일을 Outlook 폴더로 가져오기
        """
        eml_path = Path(eml_path)
        
        # EML 파일 파싱
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        
        # 새 메일 아이템 생성
        mail_item = self.outlook.CreateItem(0)  # 0 = olMailItem
        
        # 속성 설정
        mail_item.Subject = msg.get('Subject', '')
        
        # 본문 설정
        body = ""
        html_body = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == 'text/plain' and not body:
                    try:
                        body = part.get_content()
                    except:
                        body = part.get_payload(decode=True).decode('utf-8', errors='replace')
                elif content_type == 'text/html' and not html_body:
                    try:
                        html_body = part.get_content()
                    except:
                        html_body = part.get_payload(decode=True).decode('utf-8', errors='replace')
        else:
            try:
                if msg.get_content_type() == 'text/html':
                    html_body = msg.get_content()
                else:
                    body = msg.get_content()
            except:
                payload = msg.get_payload(decode=True)
                if payload:
                    body = payload.decode('utf-8', errors='replace')
        
        if html_body:
            mail_item.HTMLBody = html_body
        else:
            mail_item.Body = body or ""
        
        # 수신자 설정 (To)
        to_addr = msg.get('To', '')
        if to_addr:
            for addr in to_addr.split(','):
                addr = addr.strip()
                if addr:
                    recipient = mail_item.Recipients.Add(addr)
                    recipient.Type = 1  # olTo
        
        # CC 설정
        cc_addr = msg.get('Cc', '')
        if cc_addr:
            for addr in cc_addr.split(','):
                addr = addr.strip()
                if addr:
                    recipient = mail_item.Recipients.Add(addr)
                    recipient.Type = 2  # olCC
        
        # 첨부파일 처리
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        # 임시 파일로 저장 후 첨부
                        payload = part.get_payload(decode=True)
                        if payload:
                            temp_dir = tempfile.mkdtemp()
                            temp_path = os.path.join(temp_dir, filename)
                            with open(temp_path, 'wb') as f:
                                f.write(payload)
                            mail_item.Attachments.Add(temp_path)
                            os.remove(temp_path)
                            os.rmdir(temp_dir)
        
        # 폴더로 이동
        mail_item.Move(folder)
    
    def convert_directory(self, directory: str, pst_path: str, 
                         recursive: bool = False,
                         folder_name: str = "Imported Emails") -> str:
        """
        디렉토리 내의 모든 EML 파일을 PST로 변환
        """
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
            return None
        
        print(f"총 {len(eml_files)}개의 EML 파일을 발견했습니다.")
        
        return self.convert_files(
            [str(f) for f in eml_files], 
            pst_path, 
            folder_name
        )


# macOS/Linux용 대체 구현 (MBOX)
class EMLtoMBOXConverter:
    """EML 파일들을 MBOX 파일로 변환 (크로스 플랫폼 대안)"""
    
    def __init__(self, verbose: bool = False):
        self.verbose = verbose
    
    def log(self, message: str):
        if self.verbose:
            print(f"  [정보] {message}")
        logger.debug(message)
    
    def convert_files(self, eml_paths: List[str], mbox_path: str) -> str:
        """
        여러 EML 파일을 하나의 MBOX 파일로 변환
        """
        import mailbox
        
        mbox_path = Path(mbox_path)
        
        # MBOX 파일 생성
        mbox = mailbox.mbox(str(mbox_path))
        
        try:
            mbox.lock()
            
            for eml_path in eml_paths:
                try:
                    with open(eml_path, 'rb') as f:
                        msg = BytesParser(policy=policy.default).parse(f)
                    
                    # MBOX 메시지로 변환
                    mbox_msg = mailbox.mboxMessage(msg)
                    mbox.add(mbox_msg)
                    
                    self.log(f"추가됨: {Path(eml_path).name}")
                    
                except Exception as e:
                    logger.error(f"EML 추가 실패 {eml_path}: {e}")
            
            mbox.flush()
            
        finally:
            mbox.unlock()
            mbox.close()
        
        self.log(f"MBOX 저장됨: {mbox_path}")
        
        return str(mbox_path)
