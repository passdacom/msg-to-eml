#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EML to PST Converter
여러 .eml 파일을 하나의 Microsoft Outlook .pst 파일로 변환

기술: pywin32를 통한 Outlook COM 자동화
요구사항: Windows + Microsoft Outlook 설치

수정사항 v2.1:
- 메일이 '받은 편지함' 형태로 표시되도록 수정 (초안 X)
- 폴더 이름에 날짜/시간 포함
- 폴더 이름 커스터마이즈 가능
"""

import os
import sys
import email
import tempfile
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from datetime import datetime
from pathlib import Path
from typing import Optional, List
import logging
import re

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


def generate_folder_name() -> str:
    """변환 폴더 이름 생성 (날짜/시간 포함)"""
    now = datetime.now()
    return f"Converted Mails ({now.strftime('%Y-%m-%d %H.%M')})"


class EMLtoPSTConverter:
    """EML 파일들을 PST 파일로 변환하는 클래스"""
    
    # MAPI 속성 태그
    PR_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
    PR_MESSAGE_DELIVERY_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"
    PR_CLIENT_SUBMIT_TIME = "http://schemas.microsoft.com/mapi/proptag/0x00390040"
    PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
    PR_SENDER_EMAIL_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F"
    
    # 메시지 플래그
    MSGFLAG_READ = 0x0001
    MSGFLAG_UNSENT = 0x0008  # 이 플래그가 있으면 초안/작성중
    
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
        """Outlook 초기화 (현재 스레드에서)"""
        try:
            # 항상 현재 스레드에서 COM 초기화
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.log("Outlook 연결 성공")
        except Exception as e:
            raise RuntimeError(f"Outlook 초기화 실패: {e}")
    
    def _cleanup_outlook(self):
        """Outlook 정리 (현재 스레드에서)"""
        try:
            self.namespace = None
            self.outlook = None
            pythoncom.CoUninitialize()
        except:
            pass
    
    def convert_files(self, eml_paths: List[str], pst_path: str, 
                      folder_name: str = None) -> str:
        """
        여러 EML 파일을 하나의 PST 파일로 변환
        
        Args:
            eml_paths: EML 파일 경로 리스트
            pst_path: 생성할 PST 파일 경로
            folder_name: PST 내 폴더 이름 (None이면 자동 생성)
            
        Returns:
            생성된 PST 파일 경로
        """
        available, error = self.is_available()
        if not available:
            raise RuntimeError(f"Outlook을 사용할 수 없습니다: {error}")
        
        pst_path = Path(pst_path)
        
        # 폴더 이름 자동 생성
        if folder_name is None:
            folder_name = generate_folder_name()
        
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
            
            # 루트 폴더 이름 변경 (PST 파일명의 표시 이름)
            try:
                pst_display_name = f"Imported ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
                root_folder.Name = pst_display_name
            except Exception as e:
                self.log(f"루트 폴더 이름 변경 실패: {e}")
            
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
                    self._import_eml_as_received(eml_path, import_folder)
                    success_count += 1
                    self.log(f"추가됨: {Path(eml_path).name}")
                except Exception as e:
                    error_count += 1
                    logger.error(f"EML 추가 실패 {eml_path}: {e}")
                    import traceback
                    logger.error(traceback.format_exc())
            
            self.log(f"변환 완료: {success_count}개 성공, {error_count}개 실패")
            
            return str(pst_path)
            
        finally:
            self._cleanup_outlook()
    
    def _import_eml_as_received(self, eml_path: str, folder):
        """
        EML 파일을 '받은 메일'로 가져오기 (초안 아님)
        
        핵심: PostItem을 사용하거나, MailItem 저장 후 MAPI 속성 수정
        """
        eml_path = Path(eml_path)
        
        # EML 파일 파싱
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        
        # 발신자/수신자 정보 추출
        from_addr = msg.get('From', '')
        to_addr = msg.get('To', '')
        cc_addr = msg.get('Cc', '')
        subject = msg.get('Subject', '(제목 없음)')
        
        # 날짜 파싱
        date_str = msg.get('Date', '')
        received_time = None
        if date_str:
            try:
                received_time = parsedate_to_datetime(date_str)
            except:
                pass
        
        # 본문 추출
        body = ""
        html_body = ""
        
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = part.get_content_disposition()
                
                # 첨부파일이 아닌 경우만 본문으로 처리
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
                if msg.get_content_type() == 'text/html':
                    html_body = msg.get_content()
                else:
                    body = msg.get_content()
            except:
                payload = msg.get_payload(decode=True)
                if payload:
                    body = payload.decode('utf-8', errors='replace')
        
        # 방법 1: PostItem 사용 (받은 메일처럼 표시)
        # PostItem은 저장 시 초안으로 표시되지 않음
        try:
            # 대상 폴더에 직접 아이템 생성
            mail_item = folder.Items.Add("IPM.Note")
        except:
            # 폴백: 일반 MailItem 생성
            mail_item = self.outlook.CreateItem(0)
        
        # 기본 속성 설정
        mail_item.Subject = subject
        
        if html_body:
            mail_item.HTMLBody = html_body
        else:
            mail_item.Body = body or ""
        
        # 발신자 정보 설정 (SenderName, SenderEmailAddress)
        try:
            # PropertyAccessor를 사용하여 발신자 정보 설정
            prop_accessor = mail_item.PropertyAccessor
            
            # 발신자 이름과 이메일 파싱
            sender_name, sender_email = self._parse_email_address(from_addr)
            
            if sender_name:
                try:
                    prop_accessor.SetProperty(self.PR_SENDER_NAME, sender_name)
                except:
                    pass
            if sender_email:
                try:
                    prop_accessor.SetProperty(self.PR_SENDER_EMAIL_ADDRESS, sender_email)
                except:
                    pass
        except Exception as e:
            self.log(f"발신자 설정 실패: {e}")
        
        # 수신자 추가 (To)
        if to_addr:
            for addr in self._split_addresses(to_addr):
                if addr:
                    try:
                        recipient = mail_item.Recipients.Add(addr)
                        recipient.Type = 1  # olTo
                    except:
                        pass
        
        # CC 추가
        if cc_addr:
            for addr in self._split_addresses(cc_addr):
                if addr:
                    try:
                        recipient = mail_item.Recipients.Add(addr)
                        recipient.Type = 2  # olCC
                    except:
                        pass
        
        # 수신자 확인
        try:
            mail_item.Recipients.ResolveAll()
        except:
            pass
        
        # 첨부파일 처리
        if msg.is_multipart():
            for part in msg.walk():
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
        
        # 먼저 저장하여 메시지 생성
        mail_item.Save()
        
        # MAPI 속성 수정하여 '받은 메일'로 표시
        # 핵심: MSGFLAG_UNSENT 비트를 클리어해야 함
        try:
            prop_accessor = mail_item.PropertyAccessor
            
            # 현재 메시지 플래그 읽기
            try:
                current_flags = prop_accessor.GetProperty(self.PR_MESSAGE_FLAGS)
            except:
                current_flags = self.MSGFLAG_UNSENT  # 기본값: 미발송
            
            # MSGFLAG_UNSENT (0x0008) 비트 클리어
            # 비트 AND NOT 연산으로 해당 비트만 제거
            new_flags = current_flags & (~self.MSGFLAG_UNSENT)
            
            # MSGFLAG_READ (0x0001) 비트 추가 (읽음 상태)
            new_flags = new_flags | self.MSGFLAG_READ
            
            self.log(f"플래그 변경: 0x{current_flags:08X} -> 0x{new_flags:08X}")
            
            # 새 플래그 설정
            prop_accessor.SetProperty(self.PR_MESSAGE_FLAGS, new_flags)
            
            # 받은 시간 설정
            if received_time:
                try:
                    prop_accessor.SetProperty(self.PR_MESSAGE_DELIVERY_TIME, received_time)
                    prop_accessor.SetProperty(self.PR_CLIENT_SUBMIT_TIME, received_time)
                except Exception as e:
                    self.log(f"시간 설정 실패: {e}")
            
            # 변경사항 저장
            mail_item.Save()
            
        except Exception as e:
            self.log(f"MAPI 속성 설정 실패: {e}")
            import traceback
            logger.error(traceback.format_exc())
        
        # 아이템이 다른 폴더에 있으면 대상 폴더로 이동
        try:
            if mail_item.Parent.EntryID != folder.EntryID:
                mail_item.Move(folder)
        except:
            pass
    
    def _parse_email_address(self, addr_str: str) -> tuple:
        """
        이메일 주소 문자열 파싱
        "Name <email@example.com>" -> ("Name", "email@example.com")
        """
        if not addr_str:
            return "", ""
        
        # "Name <email>" 형식
        match = re.match(r'^"?([^"<]*)"?\s*<([^>]+)>$', addr_str.strip())
        if match:
            return match.group(1).strip(), match.group(2).strip()
        
        # 이메일만 있는 경우
        match = re.match(r'^([^@\s]+@[^@\s]+)$', addr_str.strip())
        if match:
            return "", match.group(1)
        
        return "", addr_str.strip()
    
    def _split_addresses(self, addr_str: str) -> List[str]:
        """여러 이메일 주소 분리"""
        if not addr_str:
            return []
        
        # 간단한 분리 (쉼표/세미콜론)
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
    
    def convert_directory(self, directory: str, pst_path: str, 
                         recursive: bool = False,
                         folder_name: str = None) -> str:
        """디렉토리 내의 모든 EML 파일을 PST로 변환"""
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
        
        return self.convert_files([str(f) for f in eml_files], pst_path, folder_name)


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
        """여러 EML 파일을 하나의 MBOX 파일로 변환"""
        import mailbox
        
        mbox_path = Path(mbox_path)
        mbox = mailbox.mbox(str(mbox_path))
        
        try:
            mbox.lock()
            
            for eml_path in eml_paths:
                try:
                    with open(eml_path, 'rb') as f:
                        msg = BytesParser(policy=policy.default).parse(f)
                    
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
