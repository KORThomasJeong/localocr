import os
import sys
import shutil
import subprocess
from pathlib import Path

def build_exe():
    """
    PyInstaller를 사용하여 애플리케이션을 실행 파일로 빌드합니다.
    """
    print("Gemini OCR 애플리케이션 빌드를 시작합니다...")
    
    # 빌드 디렉토리 생성
    build_dir = "build"
    dist_dir = "dist"
    
    # 기존 빌드 디렉토리 정리
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)
    if os.path.exists(dist_dir):
        shutil.rmtree(dist_dir)
    
    # PyInstaller 명령 구성
    pyinstaller_cmd = [
        "pyinstaller",
        "--name=GeminiOCR",
        "--windowed",  # GUI 애플리케이션
        "--onefile",   # 단일 실행 파일로 패키징
        "--icon=NONE", # 아이콘 없음 (필요시 아이콘 파일 경로로 변경)
        "--add-data=README.md;.",  # 사용설명서 포함
        "gemini_gui.py"  # 메인 스크립트
    ]
    
    # PyInstaller 실행
    try:
        subprocess.run(pyinstaller_cmd, check=True)
        print("빌드가 성공적으로 완료되었습니다!")
        print(f"실행 파일 위치: {os.path.abspath(os.path.join(dist_dir, 'GeminiOCR.exe'))}")
    except subprocess.CalledProcessError as e:
        print(f"빌드 중 오류가 발생했습니다: {e}")
        return False
    
    return True

if __name__ == "__main__":
    build_exe()
