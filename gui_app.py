#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Email Format Converter - Desktop GUI Application
다양한 이메일 형식 변환을 지원하는 CustomTkinter 기반 데스크톱 앱

지원 변환:
- MSG → EML
- EML → MSG
- EML → PST (Windows + Outlook 필요)

실행: python gui_app.py
"""

import os
import sys
import threading
import logging
import traceback
import queue
import platform
from datetime import datetime
from pathlib import Path
from tkinter import filedialog
import customtkinter as ctk

# 로그 설정
def setup_logging():
    """로그 설정"""
    if getattr(sys, 'frozen', False):
        log_dir = Path(sys.executable).parent
    else:
        log_dir = Path(__file__).parent
    
    log_file = log_dir / "converter.log"
    log_format = "%(asctime)s [%(levelname)s] %(message)s"
    
    logging.basicConfig(
        level=logging.DEBUG,
        format=log_format,
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8', mode='a'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info("=" * 60)
    logger.info(f"Email Format Converter 시작 - {datetime.now()}")
    logger.info(f"Python 버전: {sys.version}")
    logger.info(f"OS: {platform.system()}")
    logger.info("=" * 60)
    
    return logger

logger = setup_logging()

# 변환기 import
try:
    from converters.msg_to_eml import MSGtoEMLConverter
    logger.info("MSGtoEMLConverter 로드 성공")
except Exception as e:
    logger.error(f"MSGtoEMLConverter 로드 실패: {e}")
    MSGtoEMLConverter = None

try:
    from converters.eml_to_msg import EMLtoMSGConverter
    logger.info("EMLtoMSGConverter 로드 성공")
except Exception as e:
    logger.error(f"EMLtoMSGConverter 로드 실패: {e}")
    EMLtoMSGConverter = None

try:
    from converters.eml_to_pst import EMLtoPSTConverter, check_outlook_available, EMLtoMBOXConverter
    logger.info("EMLtoPSTConverter 로드 성공")
except Exception as e:
    logger.error(f"EMLtoPSTConverter 로드 실패: {e}")
    EMLtoPSTConverter = None
    check_outlook_available = lambda: (False, "모듈 로드 실패")
    EMLtoMBOXConverter = None


class ConverterTab(ctk.CTkFrame):
    """변환기 탭의 기본 클래스"""
    
    def __init__(self, parent, app, source_ext: str, target_ext: str, 
                 converter_class, combine_output: bool = False):
        super().__init__(parent, fg_color="transparent")
        
        self.app = app
        self.source_ext = source_ext.lower()
        self.target_ext = target_ext.lower()
        self.converter_class = converter_class
        self.combine_output = combine_output  # PST처럼 여러 파일을 하나로 합치는 경우
        
        self.files = []  # [(path, status, output_path), ...]
        self.output_folder = None
        
        # 변환기 인스턴스
        if converter_class:
            try:
                self.converter = converter_class(verbose=True)
                logger.info(f"{converter_class.__name__} 인스턴스 생성")
            except Exception as e:
                logger.error(f"{converter_class.__name__} 인스턴스 생성 실패: {e}")
                self.converter = None
        else:
            self.converter = None
        
        self._create_ui()
    
    def _create_ui(self):
        """UI 생성"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # ===== 버튼 영역 =====
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")
        
        self.select_files_btn = ctk.CTkButton(
            button_frame,
            text=f"📁 {self.source_ext.upper()} 파일 선택",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40,
            command=self._select_files
        )
        self.select_files_btn.pack(side="left", padx=(0, 8))
        
        self.select_folder_btn = ctk.CTkButton(
            button_frame,
            text="📂 폴더 선택",
            font=ctk.CTkFont(size=13, weight="bold"),
            height=40,
            fg_color="#2d5a27",
            hover_color="#3d7a37",
            command=self._select_folder
        )
        self.select_folder_btn.pack(side="left", padx=(0, 8))
        
        self.clear_btn = ctk.CTkButton(
            button_frame,
            text="🗑️",
            font=ctk.CTkFont(size=13),
            height=40,
            width=40,
            fg_color="#555555",
            hover_color="#666666",
            command=self._clear_files
        )
        self.clear_btn.pack(side="right")
        
        # ===== 파일 목록 =====
        list_frame = ctk.CTkFrame(self)
        list_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)
        
        list_header = ctk.CTkFrame(list_frame, fg_color="transparent")
        list_header.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")
        
        self.file_count_label = ctk.CTkLabel(
            list_header,
            text=f"{self.source_ext.upper()} 파일 목록 (0개)",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        self.file_count_label.pack(side="left")
        
        self.file_list_frame = ctk.CTkScrollableFrame(list_frame, fg_color="transparent")
        self.file_list_frame.grid(row=1, column=0, padx=10, pady=(5, 15), sticky="nsew")
        self.file_list_frame.grid_columnconfigure(0, weight=1)
        
        self._show_empty_message()
        
        # ===== 출력 설정 =====
        output_frame = ctk.CTkFrame(self, fg_color="transparent")
        output_frame.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        output_frame.grid_columnconfigure(1, weight=1)
        
        if self.combine_output:
            output_label = ctk.CTkLabel(
                output_frame,
                text=f"출력 {self.target_ext.upper()} 파일:",
                font=ctk.CTkFont(size=12)
            )
        else:
            output_label = ctk.CTkLabel(
                output_frame,
                text="출력 폴더:",
                font=ctk.CTkFont(size=12)
            )
        output_label.grid(row=0, column=0, padx=(0, 8))
        
        default_text = "지정하려면 클릭" if self.combine_output else "원본 파일과 같은 위치"
        self.output_path_var = ctk.StringVar(value=default_text)
        output_entry = ctk.CTkEntry(
            output_frame,
            textvariable=self.output_path_var,
            font=ctk.CTkFont(size=11),
            state="readonly"
        )
        output_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        
        output_btn = ctk.CTkButton(
            output_frame,
            text="변경",
            width=50,
            height=28,
            command=self._select_output
        )
        output_btn.grid(row=0, column=2)
        
        # ===== 변환 버튼 =====
        self.convert_btn = ctk.CTkButton(
            self,
            text=f"🔄 {self.target_ext.upper()}로 변환",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=45,
            command=self._start_conversion
        )
        self.convert_btn.grid(row=3, column=0, padx=20, pady=(10, 5), sticky="ew")
        
        # ===== 진행률 =====
        self.progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.progress_frame.grid(row=4, column=0, padx=20, pady=(5, 15), sticky="ew")
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(fill="x", pady=(0, 3))
        self.progress_bar.set(0)
        
        self.status_label = ctk.CTkLabel(
            self.progress_frame,
            text="대기 중...",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        self.status_label.pack()
        self.progress_frame.grid_remove()
    
    def _show_empty_message(self):
        """빈 목록 메시지 표시"""
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()
        
        label = ctk.CTkLabel(
            self.file_list_frame,
            text=f"📭 {self.source_ext.upper()} 파일을 선택하세요\n\n파일 또는 폴더를 선택하면 여기에 표시됩니다",
            font=ctk.CTkFont(size=13),
            text_color="gray"
        )
        label.pack(expand=True, pady=40)
    
    def _select_files(self):
        """파일 선택"""
        logger.info(f"파일 선택 다이얼로그 ({self.source_ext})")
        
        filetypes = [(f"{self.source_ext.upper()} 파일", f"*.{self.source_ext}")]
        files = filedialog.askopenfilenames(
            title=f"{self.source_ext.upper()} 파일 선택",
            filetypes=filetypes
        )
        
        if files:
            for f in files:
                if f.lower().endswith(f'.{self.source_ext}'):
                    self._add_file(f)
            self._update_file_list()
    
    def _select_folder(self):
        """폴더 선택"""
        logger.info(f"폴더 선택 다이얼로그 ({self.source_ext})")
        
        folder = filedialog.askdirectory(title=f"{self.source_ext.upper()} 파일이 있는 폴더 선택")
        
        if folder:
            folder_path = Path(folder)
            files = list(folder_path.glob(f"*.{self.source_ext}"))
            files += list(folder_path.glob(f"*.{self.source_ext.upper()}"))
            
            for f in files:
                self._add_file(str(f))
            
            if files:
                self._update_file_list()
            else:
                self._show_message("알림", f"{self.source_ext.upper()} 파일이 없습니다.")
    
    def _select_output(self):
        """출력 경로 선택"""
        if self.combine_output:
            file_path = filedialog.asksaveasfilename(
                title=f"출력 {self.target_ext.upper()} 파일 저장",
                defaultextension=f".{self.target_ext}",
                filetypes=[(f"{self.target_ext.upper()} 파일", f"*.{self.target_ext}")]
            )
            if file_path:
                self.output_folder = file_path
                self.output_path_var.set(Path(file_path).name)
        else:
            folder = filedialog.askdirectory(title="출력 폴더 선택")
            if folder:
                self.output_folder = folder
                self.output_path_var.set(folder)
            else:
                self.output_folder = None
                self.output_path_var.set("원본 파일과 같은 위치")
    
    def _add_file(self, file_path: str):
        """파일 추가"""
        for f, _, _ in self.files:
            if f == file_path:
                return
        self.files.append((file_path, "pending", None))
    
    def _clear_files(self):
        """파일 목록 초기화"""
        self.files = []
        self._update_file_list()
    
    def _update_file_list(self):
        """파일 목록 UI 업데이트"""
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()
        
        self.file_count_label.configure(
            text=f"{self.source_ext.upper()} 파일 목록 ({len(self.files)}개)"
        )
        
        if not self.files:
            self._show_empty_message()
            return
        
        for i, (file_path, status, output_path) in enumerate(self.files):
            self._create_file_item(i, file_path, status, output_path)
    
    def _create_file_item(self, index: int, file_path: str, status: str, output_path: str):
        """파일 아이템 위젯"""
        path = Path(file_path)
        
        item_frame = ctk.CTkFrame(self.file_list_frame)
        item_frame.grid(row=index, column=0, sticky="ew", pady=2, padx=3)
        item_frame.grid_columnconfigure(1, weight=1)
        
        status_icons = {
            "pending": ("⏳", "gray"),
            "converting": ("🔄", "#f59e0b"),
            "success": ("✅", "#10b981"),
            "error": ("❌", "#ef4444")
        }
        icon, _ = status_icons.get(status, ("⏳", "gray"))
        
        icon_label = ctk.CTkLabel(item_frame, text=icon, font=ctk.CTkFont(size=14), width=25)
        icon_label.grid(row=0, column=0, padx=(8, 4), pady=8)
        
        name_label = ctk.CTkLabel(
            item_frame,
            text=path.name,
            font=ctk.CTkFont(size=12),
            anchor="w"
        )
        name_label.grid(row=0, column=1, sticky="w", pady=8)
        
        if status == "pending":
            btn = ctk.CTkButton(
                item_frame, text="✕", width=25, height=25,
                fg_color="transparent", hover_color="#444444",
                command=lambda idx=index: self._remove_file(idx)
            )
            btn.grid(row=0, column=2, padx=8)
    
    def _remove_file(self, index: int):
        """파일 제거"""
        if 0 <= index < len(self.files):
            del self.files[index]
            self._update_file_list()
    
    def _start_conversion(self):
        """변환 시작"""
        if not self.converter:
            self._show_message("오류", "변환기를 사용할 수 없습니다.")
            return
        
        pending = [(i, f, s, o) for i, (f, s, o) in enumerate(self.files) if s == "pending"]
        
        if not pending:
            self._show_message("알림", "변환할 파일이 없습니다.")
            return
        
        # PST 변환 시 출력 파일 필수
        if self.combine_output and not self.output_folder:
            self._show_message("알림", f"출력 {self.target_ext.upper()} 파일을 지정하세요.")
            return
        
        self.convert_btn.configure(state="disabled", text="변환 중...")
        self.select_files_btn.configure(state="disabled")
        self.select_folder_btn.configure(state="disabled")
        self.progress_frame.grid()
        self.progress_bar.set(0)
        
        thread = threading.Thread(target=self._convert_files, args=(pending,))
        thread.daemon = True
        thread.start()
    
    def _convert_files(self, pending_files):
        """변환 실행 (백그라운드)"""
        total = len(pending_files)
        success = 0
        errors = 0
        
        if self.combine_output:
            # PST처럼 여러 파일을 하나로 합치는 경우
            file_paths = [f for _, f, _, _ in pending_files]
            
            for i, (list_index, file_path, _, _) in enumerate(pending_files):
                self.files[list_index] = (file_path, "converting", None)
            
            self.app._schedule_update(self._update_file_list)
            self.app._schedule_update(lambda: self.status_label.configure(
                text=f"변환 중... {total}개 파일"
            ))
            
            try:
                result = self.converter.convert_files(file_paths, self.output_folder)
                
                for i, (list_index, file_path, _, _) in enumerate(pending_files):
                    self.files[list_index] = (file_path, "success", result)
                    success += 1
                
            except Exception as e:
                logger.error(f"변환 실패: {e}")
                logger.error(traceback.format_exc())
                
                for i, (list_index, file_path, _, _) in enumerate(pending_files):
                    self.files[list_index] = (file_path, "error", str(e))
                    errors += 1
        else:
            # 개별 파일 변환
            for idx, (list_index, file_path, _, _) in enumerate(pending_files):
                self.files[list_index] = (file_path, "converting", None)
                self.app._schedule_update(self._update_file_list)
                self.app._schedule_update(lambda i=idx, t=total, f=file_path: 
                    self.status_label.configure(text=f"변환 중... ({i+1}/{t}) {Path(f).name}")
                )
                
                try:
                    input_path = Path(file_path)
                    if self.output_folder:
                        output_path = Path(self.output_folder) / input_path.with_suffix(f'.{self.target_ext}').name
                    else:
                        output_path = input_path.with_suffix(f'.{self.target_ext}')
                    
                    result = self.converter.convert_file(str(input_path), str(output_path))
                    self.files[list_index] = (file_path, "success", str(output_path))
                    success += 1
                    
                except Exception as e:
                    logger.error(f"변환 실패 {file_path}: {e}")
                    logger.error(traceback.format_exc())
                    self.files[list_index] = (file_path, "error", str(e))
                    errors += 1
                
                self.app._schedule_update(self._update_file_list)
                self.app._schedule_update(lambda i=idx+1, t=total: self.progress_bar.set(i/t))
        
        self.app._schedule_update(lambda: self._conversion_complete(success, errors))
    
    def _conversion_complete(self, success: int, errors: int):
        """변환 완료"""
        self.convert_btn.configure(state="normal", text=f"🔄 {self.target_ext.upper()}로 변환")
        self.select_files_btn.configure(state="normal")
        self.select_folder_btn.configure(state="normal")
        self.progress_bar.set(1)
        self.status_label.configure(text=f"완료! 성공: {success}, 실패: {errors}")
        
        if errors == 0:
            self._show_message("완료", f"✅ {success}개 파일 변환 완료!")
        else:
            self._show_message("완료", f"✅ 성공: {success}개\n❌ 실패: {errors}개")
    
    def _show_message(self, title: str, message: str):
        """메시지 다이얼로그"""
        dialog = ctk.CTkToplevel(self.app)
        dialog.title(title)
        dialog.geometry("320x140")
        dialog.transient(self.app)
        dialog.grab_set()
        
        dialog.update_idletasks()
        x = self.app.winfo_x() + (self.app.winfo_width() - 320) // 2
        y = self.app.winfo_y() + (self.app.winfo_height() - 140) // 2
        dialog.geometry(f"+{x}+{y}")
        
        label = ctk.CTkLabel(dialog, text=message, font=ctk.CTkFont(size=13), wraplength=280)
        label.pack(expand=True, pady=15)
        
        btn = ctk.CTkButton(dialog, text="확인", command=dialog.destroy)
        btn.pack(pady=(0, 15))


class EmailConverterApp(ctk.CTk):
    """이메일 형식 변환 앱"""
    
    def __init__(self):
        logger.info("앱 초기화 시작")
        super().__init__()
        
        self.title("Email Format Converter")
        self.geometry("750x650")
        self.minsize(650, 550)
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.update_queue = queue.Queue()
        
        self._create_ui()
        self._poll_queue()
        
        logger.info("앱 초기화 완료")
    
    def _poll_queue(self):
        """큐 폴링"""
        try:
            while True:
                try:
                    callback = self.update_queue.get_nowait()
                    callback()
                except queue.Empty:
                    break
        except Exception as e:
            logger.error(f"큐 처리 오류: {e}")
        self.after(100, self._poll_queue)
    
    def _schedule_update(self, callback):
        """스레드 안전 UI 업데이트"""
        self.update_queue.put(callback)
    
    def _create_ui(self):
        """UI 생성"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # 헤더
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=0, column=0, padx=25, pady=(25, 10), sticky="ew")
        
        title = ctk.CTkLabel(
            header,
            text="📧 Email Format Converter",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.pack(anchor="w")
        
        subtitle = ctk.CTkLabel(
            header,
            text="MSG, EML, PST 형식 간 변환",
            font=ctk.CTkFont(size=13),
            text_color="gray"
        )
        subtitle.pack(anchor="w", pady=(3, 0))
        
        # 탭 뷰
        self.tabview = ctk.CTkTabview(self, height=450)
        self.tabview.grid(row=1, column=0, padx=25, pady=10, sticky="nsew")
        
        # 탭 1: MSG → EML
        tab1 = self.tabview.add("MSG → EML")
        self.msg_to_eml_tab = ConverterTab(
            tab1, self, "msg", "eml", MSGtoEMLConverter
        )
        self.msg_to_eml_tab.pack(fill="both", expand=True)
        
        # 탭 2: EML → MSG
        tab2 = self.tabview.add("EML → MSG")
        self.eml_to_msg_tab = ConverterTab(
            tab2, self, "eml", "msg", EMLtoMSGConverter
        )
        self.eml_to_msg_tab.pack(fill="both", expand=True)
        
        # 탭 3: EML → PST
        tab3 = self.tabview.add("EML → PST")
        
        # PST 변환 가능 여부 확인
        if platform.system() == "Windows" and EMLtoPSTConverter:
            available, error = check_outlook_available()
            if available:
                self.eml_to_pst_tab = ConverterTab(
                    tab3, self, "eml", "pst", EMLtoPSTConverter, combine_output=True
                )
                self.eml_to_pst_tab.pack(fill="both", expand=True)
            else:
                self._show_pst_unavailable(tab3, f"Outlook 필요: {error}")
        else:
            self._show_pst_unavailable(tab3, "Windows + Outlook 필요")
        
        # 푸터
        footer = ctk.CTkLabel(
            self,
            text="오프라인에서 작동 • 파일은 저장되지 않음",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        footer.grid(row=2, column=0, pady=(5, 15))
    
    def _show_pst_unavailable(self, parent, message: str):
        """PST 변환 불가 메시지"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="both", expand=True)
        
        label = ctk.CTkLabel(
            frame,
            text="⚠️ PST 변환 불가",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color="#f59e0b"
        )
        label.pack(pady=(80, 10))
        
        desc = ctk.CTkLabel(
            frame,
            text=message,
            font=ctk.CTkFont(size=13),
            text_color="gray"
        )
        desc.pack()
        
        info = ctk.CTkLabel(
            frame,
            text="PST 변환은 Windows에서 Microsoft Outlook이\n설치된 환경에서만 사용할 수 있습니다.",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        info.pack(pady=(20, 0))


def main():
    logger.info("main() 시작")
    try:
        app = EmailConverterApp()
        logger.info("메인 루프 시작")
        app.mainloop()
        logger.info("메인 루프 종료")
    except Exception as e:
        logger.error(f"앱 실행 오류: {e}")
        logger.error(traceback.format_exc())
        raise


if __name__ == '__main__':
    main()
