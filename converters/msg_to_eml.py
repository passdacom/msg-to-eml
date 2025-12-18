#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MSG to EML Converter
Microsoft Outlook .msg 파일을 표준 .eml 형식으로 변환하는 프로그램

사용법:
    python msg_to_eml.py <input.msg>                    # 단일 파일 변환
    python msg_to_eml.py <input.msg> -o <output.eml>    # 출력 파일명 지정
    python msg_to_eml.py <directory>                    # 폴더 내 모든 .msg 파일 변환
    python msg_to_eml.py <directory> -r                 # 하위 폴더 포함 변환

필요한 라이브러리:
    pip install extract-msg
"""

import os
import sys
import argparse
import email
from email import policy
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email.utils import formataddr, formatdate, parseaddr
from email import encoders
from datetime import datetime
from pathlib import Path

try:
    import extract_msg
except ImportError:
    print("=" * 60)
    print("오류: 'extract-msg' 라이브러리가 설치되지 않았습니다.")
    print("다음 명령어로 설치해주세요:")
    print("    pip install extract-msg")
    print("=" * 60)
    sys.exit(1)


class MSGtoEMLConverter:
    """MSG 파일을 EML 형식으로 변환하는 클래스"""
    
    def __init__(self, verbose: bool = False):
        self.verbose = verbose
    
    def log(self, message: str):
        """상세 모드에서만 메시지 출력"""
        if self.verbose:
            print(f"  [정보] {message}")
    
    def convert_file(self, msg_path: str, output_path: str = None) -> str:
        """
        단일 MSG 파일을 EML로 변환
        
        Args:
            msg_path: 입력 .msg 파일 경로
            output_path: 출력 .eml 파일 경로 (None이면 자동 생성)
            
        Returns:
            생성된 .eml 파일 경로
        """
        msg_path = Path(msg_path)
        
        if not msg_path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {msg_path}")
        
        if not msg_path.suffix.lower() == '.msg':
            raise ValueError(f"MSG 파일이 아닙니다: {msg_path}")
        
        # 출력 경로 결정
        if output_path is None:
            output_path = msg_path.with_suffix('.eml')
        else:
            output_path = Path(output_path)
        
        self.log(f"변환 중: {msg_path.name}")
        
        # MSG 파일 열기 (유효성 검사 포함)
        try:
            msg = extract_msg.Message(str(msg_path))
        except Exception as e:
            error_str = str(e).lower()
            if 'ole2' in error_str or 'not an ole' in error_str or 'olefileerror' in error_str:
                raise ValueError(
                    f"유효한 MSG 파일이 아닙니다: {msg_path.name}\n"
                    f"이 파일은 Outlook MSG 형식(.msg)이 아닙니다.\n"
                    f"파일 확장자만 .msg로 변경된 다른 형식의 파일일 수 있습니다."
                )
            else:
                raise ValueError(f"MSG 파일을 열 수 없습니다: {msg_path.name} - {e}")
        
        try:
            # EML 메시지 생성
            eml_message = self._create_eml_message(msg)
            
            # EML 파일로 저장
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(eml_message.as_string())
            
            self.log(f"저장됨: {output_path.name}")
            
            return str(output_path)
            
        finally:
            msg.close()
    
    def _create_eml_message(self, msg) -> email.message.EmailMessage:
        """MSG 객체로부터 EML 메시지 객체 생성"""
        
        # 본문과 첨부파일 여부에 따라 메시지 타입 결정
        has_html = msg.htmlBody is not None
        has_plain = msg.body is not None
        has_attachments = len(msg.attachments) > 0 if msg.attachments else False
        
        if has_attachments or (has_html and has_plain):
            # multipart 메시지
            if has_attachments:
                eml = MIMEMultipart('mixed')
                
                # 본문 파트
                if has_html and has_plain:
                    body_part = MIMEMultipart('alternative')
                    body_part.attach(MIMEText(msg.body or '', 'plain', 'utf-8'))
                    body_part.attach(MIMEText(msg.htmlBody or '', 'html', 'utf-8'))
                    eml.attach(body_part)
                elif has_html:
                    eml.attach(MIMEText(msg.htmlBody, 'html', 'utf-8'))
                elif has_plain:
                    eml.attach(MIMEText(msg.body, 'plain', 'utf-8'))
            else:
                # 첨부파일 없이 HTML과 Plain text만 있는 경우
                eml = MIMEMultipart('alternative')
                eml.attach(MIMEText(msg.body or '', 'plain', 'utf-8'))
                eml.attach(MIMEText(msg.htmlBody or '', 'html', 'utf-8'))
        elif has_html:
            eml = MIMEText(msg.htmlBody, 'html', 'utf-8')
        else:
            eml = MIMEText(msg.body or '', 'plain', 'utf-8')
        
        # 헤더 설정
        self._set_headers(eml, msg)
        
        # 첨부파일 추가
        if has_attachments:
            self._add_attachments(eml, msg)
        
        return eml
    
    def _set_headers(self, eml, msg):
        """이메일 헤더 설정"""
        
        # 발신자 - sender가 이미 "이름 <이메일>" 형식의 문자열
        if msg.sender:
            eml['From'] = msg.sender
        
        # 수신자
        if msg.to:
            eml['To'] = msg.to
        
        # 참조
        if msg.cc:
            eml['Cc'] = msg.cc
        
        # 숨은 참조
        if msg.bcc:
            eml['Bcc'] = msg.bcc
        
        # 제목
        if msg.subject:
            eml['Subject'] = msg.subject
        
        # 날짜
        if msg.date:
            try:
                if isinstance(msg.date, datetime):
                    eml['Date'] = formatdate(msg.date.timestamp(), localtime=True)
                elif isinstance(msg.date, str):
                    eml['Date'] = msg.date
                else:
                    eml['Date'] = str(msg.date)
            except Exception:
                eml['Date'] = formatdate(localtime=True)
        else:
            eml['Date'] = formatdate(localtime=True)
        
        # Message-ID
        if hasattr(msg, 'messageId') and msg.messageId:
            eml['Message-ID'] = msg.messageId
        
        # 우선순위 (있는 경우)
        if hasattr(msg, 'importance') and msg.importance:
            importance_map = {0: 'low', 1: 'normal', 2: 'high'}
            priority = importance_map.get(msg.importance, 'normal')
            if priority != 'normal':
                eml['X-Priority'] = {'low': '5', 'high': '1'}.get(priority, '3')
                eml['Importance'] = priority.capitalize()
    
    def _add_attachments(self, eml: MIMEMultipart, msg):
        """첨부파일 추가"""
        
        if not msg.attachments:
            return
        
        for attachment in msg.attachments:
            try:
                # 첨부파일 데이터 가져오기
                if hasattr(attachment, 'data') and attachment.data:
                    data = attachment.data
                elif hasattr(attachment, '_data') and attachment._data:
                    data = attachment._data
                else:
                    self.log(f"첨부파일 데이터 없음: {getattr(attachment, 'longFilename', 'unknown')}")
                    continue
                
                # 파일명 결정
                filename = getattr(attachment, 'longFilename', None) or \
                          getattr(attachment, 'shortFilename', None) or \
                          getattr(attachment, 'name', None) or \
                          'attachment'
                
                # MIME 타입 결정
                content_type = getattr(attachment, 'mimetype', None) or 'application/octet-stream'
                
                # 첨부파일 MIME 파트 생성
                maintype, subtype = content_type.split('/', 1) if '/' in content_type else ('application', 'octet-stream')
                
                if maintype == 'text':
                    part = MIMEText(data.decode('utf-8', errors='replace'), subtype, 'utf-8')
                else:
                    part = MIMEBase(maintype, subtype)
                    part.set_payload(data)
                    encoders.encode_base64(part)
                
                # Content-Disposition 헤더 설정
                part.add_header('Content-Disposition', 'attachment', filename=filename)
                
                eml.attach(part)
                self.log(f"첨부파일 추가: {filename}")
                
            except Exception as e:
                self.log(f"첨부파일 처리 오류: {e}")
                continue
    
    def convert_directory(self, directory: str, recursive: bool = False, output_dir: str = None) -> list:
        """
        디렉토리 내의 모든 MSG 파일을 변환
        
        Args:
            directory: 입력 디렉토리 경로
            recursive: 하위 디렉토리 포함 여부
            output_dir: 출력 디렉토리 (None이면 원본과 동일 위치)
            
        Returns:
            변환된 파일 경로 리스트
        """
        directory = Path(directory)
        
        if not directory.exists():
            raise FileNotFoundError(f"디렉토리를 찾을 수 없습니다: {directory}")
        
        if not directory.is_dir():
            raise ValueError(f"디렉토리가 아닙니다: {directory}")
        
        # MSG 파일 검색
        if recursive:
            msg_files = list(directory.rglob('*.msg')) + list(directory.rglob('*.MSG'))
        else:
            msg_files = list(directory.glob('*.msg')) + list(directory.glob('*.MSG'))
        
        # 중복 제거
        msg_files = list(set(msg_files))
        
        if not msg_files:
            print(f"MSG 파일을 찾을 수 없습니다: {directory}")
            return []
        
        print(f"총 {len(msg_files)}개의 MSG 파일을 발견했습니다.")
        
        converted = []
        errors = []
        
        for msg_file in msg_files:
            try:
                # 출력 경로 결정
                if output_dir:
                    output_path = Path(output_dir) / msg_file.with_suffix('.eml').name
                else:
                    output_path = msg_file.with_suffix('.eml')
                
                result = self.convert_file(str(msg_file), str(output_path))
                converted.append(result)
                print(f"✓ {msg_file.name}")
                
            except Exception as e:
                errors.append((str(msg_file), str(e)))
                print(f"✗ {msg_file.name}: {e}")
        
        print(f"\n변환 완료: {len(converted)}개 성공, {len(errors)}개 실패")
        
        return converted


def main():
    parser = argparse.ArgumentParser(
        description='MSG to EML Converter - Microsoft Outlook .msg 파일을 .eml 형식으로 변환',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
예시:
    %(prog)s input.msg                     단일 파일 변환
    %(prog)s input.msg -o output.eml       출력 파일명 지정
    %(prog)s ./emails/                     폴더 내 모든 MSG 파일 변환
    %(prog)s ./emails/ -r                  하위 폴더 포함 변환
    %(prog)s ./emails/ -o ./converted/     변환된 파일을 다른 폴더에 저장
        '''
    )
    
    parser.add_argument(
        'input',
        help='변환할 MSG 파일 또는 디렉토리 경로'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='출력 파일 또는 디렉토리 경로'
    )
    
    parser.add_argument(
        '-r', '--recursive',
        action='store_true',
        help='하위 디렉토리를 포함하여 검색'
    )
    
    parser.add_argument(
        '-v', '--verbose',
        action='store_true',
        help='상세 정보 출력'
    )
    
    args = parser.parse_args()
    
    converter = MSGtoEMLConverter(verbose=args.verbose)
    input_path = Path(args.input)
    
    try:
        if input_path.is_file():
            # 단일 파일 변환
            result = converter.convert_file(str(input_path), args.output)
            print(f"✓ 변환 완료: {result}")
            
        elif input_path.is_dir():
            # 디렉토리 변환
            if args.output:
                output_dir = Path(args.output)
                output_dir.mkdir(parents=True, exist_ok=True)
            else:
                output_dir = None
            
            converter.convert_directory(
                str(input_path),
                recursive=args.recursive,
                output_dir=str(output_dir) if output_dir else None
            )
        else:
            print(f"오류: 경로를 찾을 수 없습니다: {input_path}")
            sys.exit(1)
            
    except Exception as e:
        print(f"오류: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()
