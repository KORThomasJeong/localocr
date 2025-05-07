import os
import sys
import json
import glob
import threading
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QLineEdit, QPushButton, QTextEdit, QFileDialog, 
                            QTabWidget, QComboBox, QMessageBox, QProgressBar, QGroupBox,
                            QRadioButton, QButtonGroup, QListWidget, QListWidgetItem, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize
from PyQt5.QtGui import QIcon, QPixmap, QFont

from config import Config
import gemini

class WorkerThread(QThread):
    """백그라운드에서 OCR 처리를 수행하는 스레드"""
    progress_signal = pyqtSignal(int, int)  # 현재 처리 중인 이미지 번호, 총 이미지 수
    result_signal = pyqtSignal(list)  # 처리 결과
    error_signal = pyqtSignal(str)  # 오류 메시지
    complete_signal = pyqtSignal(str)  # 완료 메시지

    def __init__(self, api_key, model, image_paths, output_path, custom_prompt):
        super().__init__()
        self.api_key = api_key
        self.model = model
        self.image_paths = image_paths
        self.output_path = output_path
        self.custom_prompt = custom_prompt
        self.results = []

    def run(self):
        try:
            total_images = len(self.image_paths)
            
            for i, image_path in enumerate(self.image_paths):
                # 진행 상황 업데이트
                self.progress_signal.emit(i + 1, total_images)
                
                # 이미지 처리
                result = gemini.process_image(image_path, self.api_key, self.model, self.custom_prompt)
                
                # 결과에 이미지 파일 이름 추가
                if isinstance(result, dict):
                    result['image_file'] = os.path.basename(image_path)
                    self.results.append(result)
                else:
                    self.results.append({
                        'image_file': os.path.basename(image_path),
                        'error': 'Unexpected result format'
                    })
            
            # 결과를 DataFrame으로 변환하고 Excel로 저장
            import pandas as pd
            df = pd.json_normalize(self.results)
            
            # 출력 디렉토리가 없으면 생성
            output_dir = os.path.dirname(self.output_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Excel로 저장
            df.to_excel(self.output_path, index=False)
            
            # 결과 신호 발생
            self.result_signal.emit(self.results)
            self.complete_signal.emit(f"처리가 완료되었습니다. 결과가 {self.output_path}에 저장되었습니다.")
            
        except Exception as e:
            self.error_signal.emit(f"오류 발생: {str(e)}")


class GeminiOCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = Config()
        self.image_paths = []
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Gemini OCR 애플리케이션")
        self.setGeometry(100, 100, 900, 700)
        
        # 메인 위젯 및 레이아웃
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # 탭 위젯 생성
        tabs = QTabWidget()
        main_layout.addWidget(tabs)
        
        # 메인 탭
        main_tab = QWidget()
        main_tab_layout = QVBoxLayout()
        main_tab.setLayout(main_tab_layout)
        tabs.addTab(main_tab, "OCR 처리")
        
        # 설정 탭
        settings_tab = QWidget()
        settings_tab_layout = QVBoxLayout()
        settings_tab.setLayout(settings_tab_layout)
        tabs.addTab(settings_tab, "설정")
        
        # 도움말 탭
        help_tab = QWidget()
        help_tab_layout = QVBoxLayout()
        help_tab.setLayout(help_tab_layout)
        tabs.addTab(help_tab, "도움말")
        
        # ===== 메인 탭 내용 =====
        # API 키 입력 그룹
        api_group = QGroupBox("API 키")
        api_layout = QHBoxLayout()
        api_group.setLayout(api_layout)
        
        self.api_key_input = QLineEdit()
        self.api_key_input.setEchoMode(QLineEdit.Password)
        self.api_key_input.setText(self.config.get_api_key())
        api_layout.addWidget(QLabel("Gemini API 키:"))
        api_layout.addWidget(self.api_key_input)
        
        self.save_api_key_btn = QPushButton("저장")
        self.save_api_key_btn.clicked.connect(self.save_api_key)
        api_layout.addWidget(self.save_api_key_btn)
        
        main_tab_layout.addWidget(api_group)
        
        # 모델 선택 그룹
        model_group = QGroupBox("모델 선택")
        model_layout = QHBoxLayout()
        model_group.setLayout(model_layout)
        
        model_layout.addWidget(QLabel("Gemini 모델:"))
        self.model_combo = QComboBox()
        self.model_combo.addItems(["gemini-2.0-flash", "gemini-2.0-pro", "gemini-1.5-flash", "gemini-1.5-pro"])
        self.model_combo.setCurrentText(self.config.get_model())
        model_layout.addWidget(self.model_combo)
        
        main_tab_layout.addWidget(model_group)
        
        # 이미지 선택 그룹
        image_group = QGroupBox("이미지 선택")
        image_layout = QVBoxLayout()
        image_group.setLayout(image_layout)
        
        image_select_layout = QHBoxLayout()
        
        self.image_select_radio = QRadioButton("이미지 파일")
        self.folder_select_radio = QRadioButton("폴더")
        self.image_select_radio.setChecked(True)
        
        image_select_layout.addWidget(self.image_select_radio)
        image_select_layout.addWidget(self.folder_select_radio)
        
        image_layout.addLayout(image_select_layout)
        
        file_select_layout = QHBoxLayout()
        self.file_path_input = QLineEdit()
        self.file_path_input.setReadOnly(True)
        self.browse_btn = QPushButton("찾아보기")
        self.browse_btn.clicked.connect(self.browse_files)
        
        file_select_layout.addWidget(self.file_path_input)
        file_select_layout.addWidget(self.browse_btn)
        
        image_layout.addLayout(file_select_layout)
        
        # 선택된 이미지 목록
        self.image_list = QListWidget()
        image_layout.addWidget(QLabel("선택된 이미지:"))
        image_layout.addWidget(self.image_list)
        
        # 이미지 목록 관리 버튼
        image_btn_layout = QHBoxLayout()
        self.clear_images_btn = QPushButton("모두 지우기")
        self.clear_images_btn.clicked.connect(self.clear_images)
        self.remove_image_btn = QPushButton("선택 항목 제거")
        self.remove_image_btn.clicked.connect(self.remove_selected_image)
        
        image_btn_layout.addWidget(self.clear_images_btn)
        image_btn_layout.addWidget(self.remove_image_btn)
        image_layout.addLayout(image_btn_layout)
        
        main_tab_layout.addWidget(image_group)
        
        # 출력 설정 그룹
        output_group = QGroupBox("출력 설정")
        output_layout = QVBoxLayout()
        output_group.setLayout(output_layout)
        
        output_path_layout = QHBoxLayout()
        output_path_layout.addWidget(QLabel("출력 파일:"))
        self.output_path_input = QLineEdit()
        self.output_path_input.setText(self.config.get_last_output_path() or "output.xlsx")
        self.output_browse_btn = QPushButton("찾아보기")
        self.output_browse_btn.clicked.connect(self.browse_output_path)
        
        output_path_layout.addWidget(self.output_path_input)
        output_path_layout.addWidget(self.output_browse_btn)
        
        output_layout.addLayout(output_path_layout)
        
        main_tab_layout.addWidget(output_group)
        
        # 프롬프트 설정 그룹
        prompt_group = QGroupBox("프롬프트 설정")
        prompt_layout = QVBoxLayout()
        prompt_group.setLayout(prompt_layout)
        
        prompt_type_layout = QHBoxLayout()
        self.prompt_text_radio = QRadioButton("직접 입력")
        self.prompt_file_radio = QRadioButton("파일에서 로드")
        self.prompt_text_radio.setChecked(True)
        
        prompt_type_layout.addWidget(self.prompt_text_radio)
        prompt_type_layout.addWidget(self.prompt_file_radio)
        
        prompt_layout.addLayout(prompt_type_layout)
        
        # 프롬프트 직접 입력
        self.prompt_text_edit = QTextEdit()
        self.prompt_text_edit.setPlaceholderText("OCR 처리를 위한 프롬프트를 입력하세요...")
        self.prompt_text_edit.setText("Extract all text from the image and organize it into structured data.")
        prompt_layout.addWidget(self.prompt_text_edit)
        
        # 프롬프트 파일 선택
        prompt_file_layout = QHBoxLayout()
        self.prompt_file_input = QLineEdit()
        self.prompt_file_input.setReadOnly(True)
        self.prompt_file_input.setText(self.config.get_last_prompt_file())
        self.prompt_file_btn = QPushButton("찾아보기")
        self.prompt_file_btn.clicked.connect(self.browse_prompt_file)
        
        prompt_file_layout.addWidget(self.prompt_file_input)
        prompt_file_layout.addWidget(self.prompt_file_btn)
        
        prompt_layout.addLayout(prompt_file_layout)
        
        # 프롬프트 라디오 버튼 연결
        self.prompt_text_radio.toggled.connect(self.toggle_prompt_input)
        self.prompt_file_radio.toggled.connect(self.toggle_prompt_input)
        
        main_tab_layout.addWidget(prompt_group)
        
        # 진행 상황 표시
        progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(QLabel("진행 상황:"))
        progress_layout.addWidget(self.progress_bar)
        
        main_tab_layout.addLayout(progress_layout)
        
        # 실행 버튼
        self.run_btn = QPushButton("OCR 처리 시작")
        self.run_btn.setMinimumHeight(40)
        self.run_btn.clicked.connect(self.run_ocr)
        main_tab_layout.addWidget(self.run_btn)
        
        # ===== 설정 탭 내용 =====
        # API 키 관리
        settings_api_group = QGroupBox("API 키 관리")
        settings_api_layout = QVBoxLayout()
        settings_api_group.setLayout(settings_api_layout)
        
        settings_api_key_layout = QHBoxLayout()
        self.settings_api_key_input = QLineEdit()
        self.settings_api_key_input.setEchoMode(QLineEdit.Password)
        self.settings_api_key_input.setText(self.config.get_api_key())
        settings_api_key_layout.addWidget(QLabel("Gemini API 키:"))
        settings_api_key_layout.addWidget(self.settings_api_key_input)
        
        self.settings_save_api_key_btn = QPushButton("저장")
        self.settings_save_api_key_btn.clicked.connect(self.save_settings_api_key)
        settings_api_key_layout.addWidget(self.settings_save_api_key_btn)
        
        settings_api_layout.addLayout(settings_api_key_layout)
        
        # API 키 표시 체크박스
        self.show_api_key_check = QCheckBox("API 키 표시")
        self.show_api_key_check.toggled.connect(self.toggle_api_key_visibility)
        settings_api_layout.addWidget(self.show_api_key_check)
        
        settings_tab_layout.addWidget(settings_api_group)
        
        # 기본 모델 설정
        model_settings_group = QGroupBox("기본 모델 설정")
        model_settings_layout = QHBoxLayout()
        model_settings_group.setLayout(model_settings_layout)
        
        model_settings_layout.addWidget(QLabel("기본 Gemini 모델:"))
        self.settings_model_combo = QComboBox()
        self.settings_model_combo.addItems(["gemini-2.0-flash", "gemini-2.0-pro", "gemini-1.5-flash", "gemini-1.5-pro"])
        self.settings_model_combo.setCurrentText(self.config.get_model())
        model_settings_layout.addWidget(self.settings_model_combo)
        
        self.settings_save_model_btn = QPushButton("저장")
        self.settings_save_model_btn.clicked.connect(self.save_settings_model)
        model_settings_layout.addWidget(self.settings_save_model_btn)
        
        settings_tab_layout.addWidget(model_settings_group)
        
        # 기타 설정
        other_settings_group = QGroupBox("기타 설정")
        other_settings_layout = QVBoxLayout()
        other_settings_group.setLayout(other_settings_layout)
        
        # 여기에 추가 설정 옵션을 넣을 수 있습니다
        
        settings_tab_layout.addWidget(other_settings_group)
        settings_tab_layout.addStretch()
        
        # ===== 도움말 탭 내용 =====
        help_text = """
        <h2>Gemini OCR 애플리케이션 사용 설명서</h2>
        
        <h3>1. 시작하기</h3>
        <p>이 애플리케이션은 Google의 Gemini API를 사용하여 이미지에서 텍스트를 추출하고 구조화된 데이터로 변환합니다.</p>
        <p>시작하려면 Google AI Studio에서 발급받은 Gemini API 키가 필요합니다.</p>
        
        <h3>2. API 키 설정</h3>
        <p>- 'OCR 처리' 탭 또는 '설정' 탭에서 API 키를 입력하고 '저장' 버튼을 클릭합니다.</p>
        <p>- API 키는 안전하게 로컬 시스템에 저장됩니다.</p>
        
        <h3>3. 이미지 선택</h3>
        <p>- '이미지 파일' 옵션을 선택하여 개별 이미지를 선택하거나, '폴더' 옵션을 선택하여 폴더 내의 모든 이미지를 처리할 수 있습니다.</p>
        <p>- '찾아보기' 버튼을 클릭하여 이미지 파일이나 폴더를 선택합니다.</p>
        <p>- 선택된 이미지는 목록에 표시됩니다. '모두 지우기' 또는 '선택 항목 제거' 버튼을 사용하여 목록을 관리할 수 있습니다.</p>
        
        <h3>4. 출력 설정</h3>
        <p>- 결과를 저장할 Excel 파일의 경로를 지정합니다.</p>
        <p>- 기본값은 'output.xlsx'입니다.</p>
        
        <h3>5. 프롬프트 설정</h3>
        <p>- '직접 입력' 옵션을 선택하여 텍스트 영역에 프롬프트를 입력하거나, '파일에서 로드' 옵션을 선택하여 프롬프트 파일을 로드할 수 있습니다.</p>
        <p>- 프롬프트는 Gemini API에게 이미지에서 어떤 정보를 추출할지 지시하는 역할을 합니다.</p>
        
        <h3>6. OCR 처리 실행</h3>
        <p>- 모든 설정을 완료한 후 'OCR 처리 시작' 버튼을 클릭하여 처리를 시작합니다.</p>
        <p>- 진행 상황이 진행 표시줄에 표시됩니다.</p>
        <p>- 처리가 완료되면 결과가 지정된 Excel 파일에 저장됩니다.</p>
        
        <h3>7. 문제 해결</h3>
        <p>- API 키가 올바르지 않은 경우: API 키를 다시 확인하고 올바르게 입력했는지 확인하세요.</p>
        <p>- 이미지 처리 오류: 지원되는 이미지 형식(JPG, JPEG, PNG, BMP, GIF)인지 확인하세요.</p>
        <p>- 결과가 예상과 다른 경우: 프롬프트를 더 구체적으로 작성하여 Gemini API에게 명확한 지시를 제공하세요.</p>
        """
        
        help_text_edit = QTextEdit()
        help_text_edit.setReadOnly(True)
        help_text_edit.setHtml(help_text)
        help_tab_layout.addWidget(help_text_edit)
        
        # 초기 상태 설정
        self.toggle_prompt_input()
        
    def toggle_api_key_visibility(self, checked):
        """API 키 표시 여부를 토글합니다."""
        if checked:
            self.settings_api_key_input.setEchoMode(QLineEdit.Normal)
            self.api_key_input.setEchoMode(QLineEdit.Normal)
        else:
            self.settings_api_key_input.setEchoMode(QLineEdit.Password)
            self.api_key_input.setEchoMode(QLineEdit.Password)
    
    def toggle_prompt_input(self):
        """프롬프트 입력 방식에 따라 UI를 업데이트합니다."""
        if self.prompt_text_radio.isChecked():
            self.prompt_text_edit.setEnabled(True)
            self.prompt_file_input.setEnabled(False)
            self.prompt_file_btn.setEnabled(False)
        else:
            self.prompt_text_edit.setEnabled(False)
            self.prompt_file_input.setEnabled(True)
            self.prompt_file_btn.setEnabled(True)
    
    def save_api_key(self):
        """API 키를 저장합니다."""
        api_key = self.api_key_input.text().strip()
        self.config.set_api_key(api_key)
        self.settings_api_key_input.setText(api_key)
        QMessageBox.information(self, "정보", "API 키가 저장되었습니다.")
    
    def save_settings_api_key(self):
        """설정 탭에서 API 키를 저장합니다."""
        api_key = self.settings_api_key_input.text().strip()
        self.config.set_api_key(api_key)
        self.api_key_input.setText(api_key)
        QMessageBox.information(self, "정보", "API 키가 저장되었습니다.")
    
    def save_settings_model(self):
        """설정 탭에서 모델 설정을 저장합니다."""
        model = self.settings_model_combo.currentText()
        self.config.set_model(model)
        self.model_combo.setCurrentText(model)
        QMessageBox.information(self, "정보", "기본 모델 설정이 저장되었습니다.")
    
    def browse_files(self):
        """이미지 파일 또는 폴더를 선택합니다."""
        if self.image_select_radio.isChecked():
            # 이미지 파일 선택
            files, _ = QFileDialog.getOpenFileNames(
                self, "이미지 파일 선택", 
                self.config.get_last_photo_dir(),
                "이미지 파일 (*.jpg *.jpeg *.png *.bmp *.gif)"
            )
            
            if files:
                # 마지막 디렉토리 저장
                self.config.set_last_photo_dir(os.path.dirname(files[0]))
                
                # 이미지 목록에 추가
                for file_path in files:
                    if file_path not in self.image_paths:
                        self.image_paths.append(file_path)
                        self.image_list.addItem(os.path.basename(file_path))
                
                # 파일 경로 표시
                if len(files) == 1:
                    self.file_path_input.setText(files[0])
                else:
                    self.file_path_input.setText(f"{len(files)}개의 파일이 선택됨")
        else:
            # 폴더 선택
            folder = QFileDialog.getExistingDirectory(
                self, "이미지 폴더 선택", 
                self.config.get_last_photo_dir()
            )
            
            if folder:
                self.config.set_last_photo_dir(folder)
                self.file_path_input.setText(folder)
                
                # 폴더 내 이미지 파일 찾기
                image_extensions = ['*.jpg', '*.jpeg', '*.png', '*.bmp', '*.gif']
                image_files = []
                for ext in image_extensions:
                    image_files.extend(glob.glob(os.path.join(folder, ext)))
                
                # 이미지 목록에 추가
                for file_path in image_files:
                    if file_path not in self.image_paths:
                        self.image_paths.append(file_path)
                        self.image_list.addItem(os.path.basename(file_path))
                
                if not image_files:
                    QMessageBox.warning(self, "경고", "선택한 폴더에 이미지 파일이 없습니다.")
    
    def clear_images(self):
        """이미지 목록을 모두 지웁니다."""
        self.image_paths = []
        self.image_list.clear()
        self.file_path_input.clear()
    
    def remove_selected_image(self):
        """선택한 이미지를 목록에서 제거합니다."""
        selected_items = self.image_list.selectedItems()
        if not selected_items:
            return
        
        for item in selected_items:
            file_name = item.text()
            # 파일 이름으로 전체 경로 찾기
            for path in self.image_paths[:]:
                if os.path.basename(path) == file_name:
                    self.image_paths.remove(path)
                    break
            
            # 목록에서 항목 제거
            self.image_list.takeItem(self.image_list.row(item))
    
    def browse_output_path(self):
        """출력 파일 경로를 선택합니다."""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "출력 파일 저장", 
            self.config.get_last_output_path() or "output.xlsx",
            "Excel 파일 (*.xlsx);;모든 파일 (*.*)"
        )
        
        if file_path:
            self.output_path_input.setText(file_path)
            self.config.set_last_output_path(file_path)
    
    def browse_prompt_file(self):
        """프롬프트 파일을 선택합니다."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "프롬프트 파일 선택", 
            self.config.get_last_prompt_file(),
            "텍스트 파일 (*.txt);;모든 파일 (*.*)"
        )
        
        if file_path:
            self.prompt_file_input.setText(file_path)
            self.config.set_last_prompt_file(file_path)
    
    def get_prompt(self):
        """현재 설정에 따라 프롬프트를 가져옵니다."""
        if self.prompt_text_radio.isChecked():
            return self.prompt_text_edit.toPlainText().strip()
        else:
            prompt_file = self.prompt_file_input.text()
            if not prompt_file:
                return "Extract all text from the image and organize it into structured data."
            
            try:
                with open(prompt_file, 'r', encoding='utf-8') as f:
                    return f.read().strip()
            except Exception as e:
                QMessageBox.warning(self, "경고", f"프롬프트 파일을 읽는 중 오류가 발생했습니다: {str(e)}")
                return "Extract all text from the image and organize it into structured data."
    
    def run_ocr(self):
        """OCR 처리를 시작합니다."""
        # API 키 확인
        api_key = self.api_key_input.text().strip()
        if not api_key:
            QMessageBox.warning(self, "경고", "API 키를 입력하세요.")
            return
        
        # 이미지 파일 확인
        if not self.image_paths:
            QMessageBox.warning(self, "경고", "처리할 이미지 파일을 선택하세요.")
            return
        
        # 출력 경로 확인
        output_path = self.output_path_input.text().strip()
        if not output_path:
            QMessageBox.warning(self, "경고", "출력 파일 경로를 지정하세요.")
            return
        
        # 모델 이름 가져오기
        model = self.model_combo.currentText()
        
        # 프롬프트 가져오기
        custom_prompt = self.get_prompt()
        
        # 설정 저장
        self.config.set_api_key(api_key)
        self.config.set_model(model)
        self.config.set_last_output_path(output_path)
        
        # 진행 표시줄 초기화
        self.progress_bar.setValue(0)
        
        # 워커 스레드 생성 및 시작
        self.worker = WorkerThread(api_key, model, self.image_paths, output_path, custom_prompt)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.result_signal.connect(self.process_results)
        self.worker.error_signal.connect(self.show_error)
        self.worker.complete_signal.connect(self.show_completion)
        
        # UI 비활성화
        self.run_btn.setEnabled(False)
        self.run_btn.setText("처리 중...")
        
        # 스레드 시작
        self.worker.start()
    
    def update_progress(self, current, total):
        """진행 상황을 업데이트합니다."""
        progress = int((current / total) * 100)
        self.progress_bar.setValue(progress)
    
    def process_results(self, results):
        """처리 결과를 처리합니다."""
        # 여기에서 결과를 처리하는 추가 로직을 구현할 수 있습니다
        pass
    
    def show_error(self, error_message):
        """오류 메시지를 표시합니다."""
        QMessageBox.critical(self, "오류", error_message)
        self.run_btn.setEnabled(True)
        self.run_btn.setText("OCR 처리 시작")
    
    def show_completion(self, message):
        """완료 메시지를 표시합니다."""
        QMessageBox.information(self, "완료", message)
        self.run_btn.setEnabled(True)
        self.run_btn.setText("OCR 처리 시작")


def main():
    app = QApplication(sys.argv)
    window = GeminiOCRApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
