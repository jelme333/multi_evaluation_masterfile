import flet as ft
import os
import sys

# 현재 파일(main.py)이 위치한 폴더(src)와 그 상위 폴더(root) 경로 설정
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(current_dir)
sys.path.insert(0, project_root)

# [수정 완벽 해결] 같은 src 폴더 내에 있으므로 'src.' 을 빼고 직접 파일명만 부릅니다.
from m_evaluation_logic import EvaluationLogic
from m_evaluation_ui import EvaluationExtractorUI

def main(page: ft.Page):
    # 1. 경로 설정 (데이터가 있는 최상위 폴더 기준)
    if getattr(sys, 'frozen', False):
        # .exe로 실행 중일 때
        base_path = os.path.dirname(sys.executable)
    else:
        # 파이썬으로 실행 중일 때
        base_path = project_root
        
    mapping_file_path = os.path.join(base_path, "세코닉스 리더 역량행동지표.xlsx")

    if not os.path.exists(mapping_file_path):
        page.add(ft.Text(f"치명적 오류: 매핑 파일을 찾을 수 없습니다.\n예상 경로: {mapping_file_path}", color="red"))
        return

    # 2. 핵심 로직 엔진 초기화
    try:
        logic = EvaluationLogic(base_path, mapping_file_path)
    except Exception as e:
        page.add(ft.Text(f"엔진 초기화 실패: {e}", color="red"))
        return

    # 3. 새로운 추출 전용 UI 엔진 초기화 및 렌더링
    ui = EvaluationExtractorUI(logic)
    page.add(ui.get_main_layout(page))

if __name__ == "__main__":
    # 브라우저 대신 데스크탑 앱 형태로 실행
    ft.app(target=main)