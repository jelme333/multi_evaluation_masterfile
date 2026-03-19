import flet as ft
import os
import sys

# src 폴더 내부의 모듈들을 원활하게 불러오기 위해 경로 추가
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from m_evaluation_logic import EvaluationLogic
from m_evaluation_ui import EvaluationExtractorUI

def main(page: ft.Page):
    # 경로 설정 (데이터가 있는 최상위 폴더 기준)
    if getattr(sys, 'frozen', False):
        # .exe로 실행 중일 때
        base_path = os.path.dirname(sys.executable)
    else:
        # 파이썬으로 실행 중일 때
        current_dir = os.path.dirname(os.path.abspath(__file__))
        if os.path.basename(current_dir).lower() == 'src':
            base_path = os.path.dirname(current_dir)
        else:
            base_path = current_dir

    # [A안 적용] 매핑 파일은 추출 시점에 연도별로 동적 탐색하므로
    #   앱 시작 시 고정 경로 체크 불필요 → base_path 만 전달
    try:
        logic = EvaluationLogic(base_path)
    except Exception as e:
        page.add(ft.Text(f"엔진 초기화 실패: {e}", color="red"))
        return

    # 추출 전용 UI 엔진 초기화 및 렌더링
    ui = EvaluationExtractorUI(logic)
    page.add(ui.get_main_layout(page))

if __name__ == "__main__":
    # 브라우저 대신 데스크탑 앱 형태로 실행
    ft.app(target=main)
