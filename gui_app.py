#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MSG to EML Converter - Desktop GUI Application
CustomTkinter 기반의 모던 데스크톱 앱

실행: python gui_app.py
패키징: pyinstaller --onefile --windowed gui_app.py
"""

import os
import sys
import threading
from pathlib import Path
from tkinter import filedialog
import customtkinter as ctk

# 기존 변환기 import
from msg_to_eml import MSGtoEMLConverter


class MSGtoEMLApp(ctk.CTk):
    """MSG to EML 변환기 데스크톱 앱"""
    
    def __init__(self):
        super().__init__()
        
        # 앱 설정
        self.title("MSG to EML Converter")
        self.geometry("700x600")
        self.minsize(600, 500)
        
        # 테마 설정
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # 변환기 인스턴스
        self.converter = MSGtoEMLConverter(verbose=False)
        
        # 파일 목록
        self.files = []  # [(path, status, output_path), ...]
        
        # UI 생성
        self._create_ui()
    
    def _create_ui(self):
        """UI 컴포넌트 생성"""
        
        # 메인 컨테이너
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)
        
        # ===== 헤더 =====
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=30, pady=(30, 10), sticky="ew")
        
        title_label = ctk.CTkLabel(
            header_frame,
            text="📧 MSG to EML Converter",
            font=ctk.CTkFont(size=28, weight="bold")
        )
        title_label.pack(anchor="w")
        
        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Outlook MSG 파일을 표준 EML 형식으로 변환합니다",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        subtitle_label.pack(anchor="w", pady=(5, 0))
        
        # ===== 파일 선택 버튼 영역 =====
        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=1, column=0, padx=30, pady=15, sticky="ew")
        
        self.select_files_btn = ctk.CTkButton(
            button_frame,
            text="📁 파일 선택",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=45,
            command=self._select_files
        )
        self.select_files_btn.pack(side="left", padx=(0, 10))
        
        self.select_folder_btn = ctk.CTkButton(
            button_frame,
            text="📂 폴더 선택",
            font=ctk.CTkFont(size=14, weight="bold"),
            height=45,
            fg_color="#2d5a27",
            hover_color="#3d7a37",
            command=self._select_folder
        )
        self.select_folder_btn.pack(side="left", padx=(0, 10))
        
        self.clear_btn = ctk.CTkButton(
            button_frame,
            text="🗑️ 초기화",
            font=ctk.CTkFont(size=14),
            height=45,
            fg_color="#555555",
            hover_color="#666666",
            width=100,
            command=self._clear_files
        )
        self.clear_btn.pack(side="right")
        
        # ===== 파일 목록 =====
        list_frame = ctk.CTkFrame(self)
        list_frame.grid(row=2, column=0, padx=30, pady=10, sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(1, weight=1)
        
        # 목록 헤더
        list_header = ctk.CTkFrame(list_frame, fg_color="transparent")
        list_header.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")
        
        self.file_count_label = ctk.CTkLabel(
            list_header,
            text="파일 목록 (0개)",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.file_count_label.pack(side="left")
        
        # 스크롤 가능한 파일 목록
        self.file_list_frame = ctk.CTkScrollableFrame(
            list_frame,
            fg_color="transparent"
        )
        self.file_list_frame.grid(row=1, column=0, padx=10, pady=(5, 15), sticky="nsew")
        self.file_list_frame.grid_columnconfigure(0, weight=1)
        
        # 빈 목록 안내
        self.empty_label = ctk.CTkLabel(
            self.file_list_frame,
            text="📭 변환할 MSG 파일을 선택하세요\n\n파일 또는 폴더를 선택하면 여기에 표시됩니다",
            font=ctk.CTkFont(size=14),
            text_color="gray"
        )
        self.empty_label.pack(expand=True, pady=50)
        
        # ===== 출력 폴더 설정 =====
        output_frame = ctk.CTkFrame(self, fg_color="transparent")
        output_frame.grid(row=3, column=0, padx=30, pady=10, sticky="ew")
        output_frame.grid_columnconfigure(1, weight=1)
        
        output_label = ctk.CTkLabel(
            output_frame,
            text="출력 폴더:",
            font=ctk.CTkFont(size=13)
        )
        output_label.grid(row=0, column=0, padx=(0, 10))
        
        self.output_path_var = ctk.StringVar(value="원본 파일과 같은 위치")
        output_entry = ctk.CTkEntry(
            output_frame,
            textvariable=self.output_path_var,
            font=ctk.CTkFont(size=12),
            state="readonly"
        )
        output_entry.grid(row=0, column=1, sticky="ew", padx=(0, 10))
        
        output_btn = ctk.CTkButton(
            output_frame,
            text="변경",
            width=60,
            height=30,
            command=self._select_output_folder
        )
        output_btn.grid(row=0, column=2)
        
        self.output_folder = None  # None이면 원본 위치
        
        # ===== 변환 버튼 =====
        self.convert_btn = ctk.CTkButton(
            self,
            text="🔄 변환 시작",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            command=self._start_conversion
        )
        self.convert_btn.grid(row=4, column=0, padx=30, pady=(10, 20), sticky="ew")
        
        # ===== 진행률 바 =====
        self.progress_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.progress_frame.grid(row=5, column=0, padx=30, pady=(0, 20), sticky="ew")
        
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame)
        self.progress_bar.pack(fill="x", pady=(0, 5))
        self.progress_bar.set(0)
        
        self.status_label = ctk.CTkLabel(
            self.progress_frame,
            text="대기 중...",
            font=ctk.CTkFont(size=12),
            text_color="gray"
        )
        self.status_label.pack()
        
        self.progress_frame.grid_remove()  # 초기에는 숨김
    
    def _select_files(self):
        """파일 선택 다이얼로그"""
        files = filedialog.askopenfilenames(
            title="MSG 파일 선택",
            filetypes=[("MSG 파일", "*.msg"), ("모든 파일", "*.*")]
        )
        
        if files:
            for f in files:
                if f.lower().endswith('.msg'):
                    self._add_file(f)
            self._update_file_list()
    
    def _select_folder(self):
        """폴더 선택 다이얼로그"""
        folder = filedialog.askdirectory(title="MSG 파일이 있는 폴더 선택")
        
        if folder:
            folder_path = Path(folder)
            msg_files = list(folder_path.glob("*.msg")) + list(folder_path.glob("*.MSG"))
            
            for f in msg_files:
                self._add_file(str(f))
            
            if msg_files:
                self._update_file_list()
            else:
                self._show_message("알림", "선택한 폴더에 MSG 파일이 없습니다.")
    
    def _select_output_folder(self):
        """출력 폴더 선택"""
        folder = filedialog.askdirectory(title="변환된 파일을 저장할 폴더 선택")
        
        if folder:
            self.output_folder = folder
            self.output_path_var.set(folder)
        else:
            self.output_folder = None
            self.output_path_var.set("원본 파일과 같은 위치")
    
    def _add_file(self, file_path: str):
        """파일 목록에 추가"""
        # 중복 체크
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
        # 기존 위젯 제거
        for widget in self.file_list_frame.winfo_children():
            widget.destroy()
        
        self.file_count_label.configure(text=f"파일 목록 ({len(self.files)}개)")
        
        if not self.files:
            self.empty_label = ctk.CTkLabel(
                self.file_list_frame,
                text="📭 변환할 MSG 파일을 선택하세요\n\n파일 또는 폴더를 선택하면 여기에 표시됩니다",
                font=ctk.CTkFont(size=14),
                text_color="gray"
            )
            self.empty_label.pack(expand=True, pady=50)
            return
        
        for i, (file_path, status, output_path) in enumerate(self.files):
            self._create_file_item(i, file_path, status, output_path)
    
    def _create_file_item(self, index: int, file_path: str, status: str, output_path: str):
        """파일 아이템 위젯 생성"""
        path = Path(file_path)
        
        item_frame = ctk.CTkFrame(self.file_list_frame)
        item_frame.grid(row=index, column=0, sticky="ew", pady=3, padx=5)
        item_frame.grid_columnconfigure(1, weight=1)
        
        # 상태 아이콘
        status_icons = {
            "pending": ("⏳", "gray"),
            "converting": ("🔄", "#f59e0b"),
            "success": ("✅", "#10b981"),
            "error": ("❌", "#ef4444")
        }
        icon, color = status_icons.get(status, ("⏳", "gray"))
        
        icon_label = ctk.CTkLabel(
            item_frame,
            text=icon,
            font=ctk.CTkFont(size=16),
            width=30
        )
        icon_label.grid(row=0, column=0, padx=(10, 5), pady=10)
        
        # 파일 정보
        info_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
        info_frame.grid(row=0, column=1, sticky="ew", pady=5)
        info_frame.grid_columnconfigure(0, weight=1)
        
        name_label = ctk.CTkLabel(
            info_frame,
            text=path.name,
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w"
        )
        name_label.grid(row=0, column=0, sticky="w")
        
        # 파일 크기
        try:
            size = path.stat().st_size
            size_str = self._format_size(size)
        except:
            size_str = ""
        
        detail_text = size_str
        if status == "success" and output_path:
            detail_text = f"{size_str} → {Path(output_path).name}"
        elif status == "error":
            detail_text = f"{size_str} - 변환 실패"
        
        detail_label = ctk.CTkLabel(
            info_frame,
            text=detail_text,
            font=ctk.CTkFont(size=11),
            text_color="gray",
            anchor="w"
        )
        detail_label.grid(row=1, column=0, sticky="w")
        
        # 삭제 버튼
        if status == "pending":
            remove_btn = ctk.CTkButton(
                item_frame,
                text="✕",
                width=30,
                height=30,
                fg_color="transparent",
                hover_color="#444444",
                command=lambda idx=index: self._remove_file(idx)
            )
            remove_btn.grid(row=0, column=2, padx=10)
        
        # 열기 버튼 (성공 시)
        if status == "success" and output_path:
            open_btn = ctk.CTkButton(
                item_frame,
                text="📂",
                width=30,
                height=30,
                fg_color="transparent",
                hover_color="#444444",
                command=lambda p=output_path: self._open_file_location(p)
            )
            open_btn.grid(row=0, column=2, padx=10)
    
    def _remove_file(self, index: int):
        """파일 목록에서 제거"""
        if 0 <= index < len(self.files):
            del self.files[index]
            self._update_file_list()
    
    def _format_size(self, size: int) -> str:
        """파일 크기 포맷"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"
    
    def _open_file_location(self, file_path: str):
        """파일 위치 열기"""
        import subprocess
        import platform
        
        folder = str(Path(file_path).parent)
        
        if platform.system() == "Darwin":  # macOS
            subprocess.run(["open", folder])
        elif platform.system() == "Windows":
            subprocess.run(["explorer", folder])
        else:  # Linux
            subprocess.run(["xdg-open", folder])
    
    def _start_conversion(self):
        """변환 시작"""
        pending_files = [(i, f, s, o) for i, (f, s, o) in enumerate(self.files) if s == "pending"]
        
        if not pending_files:
            self._show_message("알림", "변환할 파일이 없습니다.")
            return
        
        # UI 상태 변경
        self.convert_btn.configure(state="disabled", text="변환 중...")
        self.select_files_btn.configure(state="disabled")
        self.select_folder_btn.configure(state="disabled")
        self.progress_frame.grid()
        self.progress_bar.set(0)
        
        # 백그라운드 스레드에서 변환
        thread = threading.Thread(target=self._convert_files, args=(pending_files,))
        thread.daemon = True
        thread.start()
    
    def _convert_files(self, pending_files):
        """파일 변환 (백그라운드)"""
        total = len(pending_files)
        success_count = 0
        error_count = 0
        
        for idx, (list_index, file_path, _, _) in enumerate(pending_files):
            # 상태 업데이트
            self.files[list_index] = (file_path, "converting", None)
            self.after(0, self._update_file_list)
            self.after(0, lambda i=idx, t=total: self._update_progress(i, t, file_path))
            
            try:
                # 출력 경로 결정
                input_path = Path(file_path)
                if self.output_folder:
                    output_path = Path(self.output_folder) / input_path.with_suffix('.eml').name
                else:
                    output_path = input_path.with_suffix('.eml')
                
                # 변환
                self.converter.convert_file(str(input_path), str(output_path))
                
                self.files[list_index] = (file_path, "success", str(output_path))
                success_count += 1
                
            except Exception as e:
                self.files[list_index] = (file_path, "error", str(e))
                error_count += 1
            
            self.after(0, self._update_file_list)
            self.after(0, lambda i=idx+1, t=total: self.progress_bar.set(i/t))
        
        # 완료
        self.after(0, lambda: self._conversion_complete(success_count, error_count))
    
    def _update_progress(self, current: int, total: int, file_name: str):
        """진행률 업데이트"""
        self.status_label.configure(text=f"변환 중... ({current + 1}/{total}) {Path(file_name).name}")
    
    def _conversion_complete(self, success: int, error: int):
        """변환 완료 처리"""
        self.convert_btn.configure(state="normal", text="🔄 변환 시작")
        self.select_files_btn.configure(state="normal")
        self.select_folder_btn.configure(state="normal")
        self.progress_bar.set(1)
        self.status_label.configure(text=f"완료! 성공: {success}개, 실패: {error}개")
        
        if error == 0:
            self._show_message("완료", f"✅ {success}개 파일이 성공적으로 변환되었습니다!")
        else:
            self._show_message("완료", f"✅ 성공: {success}개\n❌ 실패: {error}개")
    
    def _show_message(self, title: str, message: str):
        """메시지 다이얼로그"""
        dialog = ctk.CTkToplevel(self)
        dialog.title(title)
        dialog.geometry("350x150")
        dialog.transient(self)
        dialog.grab_set()
        
        # 화면 중앙에 위치
        dialog.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() - 350) // 2
        y = self.winfo_y() + (self.winfo_height() - 150) // 2
        dialog.geometry(f"+{x}+{y}")
        
        label = ctk.CTkLabel(
            dialog,
            text=message,
            font=ctk.CTkFont(size=14),
            wraplength=300
        )
        label.pack(expand=True, pady=20)
        
        btn = ctk.CTkButton(
            dialog,
            text="확인",
            command=dialog.destroy
        )
        btn.pack(pady=(0, 20))


def main():
    app = MSGtoEMLApp()
    app.mainloop()


if __name__ == '__main__':
    main()
