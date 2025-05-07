import os
import json
import configparser
from pathlib import Path

class Config:
    def __init__(self):
        self.config_dir = os.path.join(os.path.expanduser("~"), ".gemini_ocr")
        self.config_file = os.path.join(self.config_dir, "config.ini")
        self.ensure_config_dir()
        self.config = configparser.ConfigParser()
        self.load_config()

    def ensure_config_dir(self):
        """설정 디렉토리가 존재하는지 확인하고, 없으면 생성합니다."""
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)

    def load_config(self):
        """설정 파일을 로드합니다."""
        if os.path.exists(self.config_file):
            self.config.read(self.config_file, encoding='utf-8')
        else:
            # 기본 설정 생성
            self.config['API'] = {'api_key': ''}
            self.config['SETTINGS'] = {
                'model': 'gemini-2.0-flash',
                'last_photo_dir': '',
                'last_output_path': '',
                'last_prompt_file': ''
            }
            self.save_config()

    def save_config(self):
        """설정 파일을 저장합니다."""
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)

    def get_api_key(self):
        """API 키를 가져옵니다."""
        return self.config.get('API', 'api_key', fallback='')

    def set_api_key(self, api_key):
        """API 키를 설정합니다."""
        self.config['API']['api_key'] = api_key
        self.save_config()

    def get_model(self):
        """모델 이름을 가져옵니다."""
        return self.config.get('SETTINGS', 'model', fallback='gemini-2.0-flash')

    def set_model(self, model):
        """모델 이름을 설정합니다."""
        self.config['SETTINGS']['model'] = model
        self.save_config()

    def get_last_photo_dir(self):
        """마지막으로 사용한 사진 디렉토리를 가져옵니다."""
        return self.config.get('SETTINGS', 'last_photo_dir', fallback='')

    def set_last_photo_dir(self, directory):
        """마지막으로 사용한 사진 디렉토리를 설정합니다."""
        self.config['SETTINGS']['last_photo_dir'] = directory
        self.save_config()

    def get_last_output_path(self):
        """마지막으로 사용한 출력 경로를 가져옵니다."""
        return self.config.get('SETTINGS', 'last_output_path', fallback='')

    def set_last_output_path(self, path):
        """마지막으로 사용한 출력 경로를 설정합니다."""
        self.config['SETTINGS']['last_output_path'] = path
        self.save_config()

    def get_last_prompt_file(self):
        """마지막으로 사용한 프롬프트 파일을 가져옵니다."""
        return self.config.get('SETTINGS', 'last_prompt_file', fallback='')

    def set_last_prompt_file(self, path):
        """마지막으로 사용한 프롬프트 파일을 설정합니다."""
        self.config['SETTINGS']['last_prompt_file'] = path
        self.save_config()
