import flet as ft
import os
import traceback

class EvaluationExtractorUI:
    def __init__(self, logic_instance):
        self.logic = logic_instance
        self.years = ["2026", "2025", "2024", "2023"] 
        self.current_year = "2025" if "2025" in self.years else (self.years[0] if self.years else "2025")
        
        # 입력 필드들을 저장할 딕셔너리
        self.inputs = {'임원': {}, '팀장': {}}

    def get_main_layout(self, page: ft.Page):
        self.page = page
        page.title = "다면평가 마스터 데이터 추출 엔진"
        page.theme_mode = ft.ThemeMode.LIGHT
        page.bgcolor = "#f1f3f5"
        page.padding = 40
        page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

        self.loading_ring = ft.ProgressRing(visible=False)
        self.status_text = ft.Text("대기 중...", color="grey")

        # 1. 연도 선택
        year_dropdown = ft.Dropdown(
            label="추출 대상 연도", value=self.current_year, width=200,
            options=[ft.dropdown.Option(y) for y in self.years],
            on_change=self.change_year,
            border_color=ft.colors.BLUE_700
        )

        # 2. 가중치 설정 패널 생성 (예외 가중치 삭제, 표준 가중치만 남김)
        exec_panel = self._create_weight_panel("임원", default_std=[50, 30, 20])
        team_panel = self._create_weight_panel("팀장", default_std=[50, 30, 20])

        # 3. 추출 버튼
        extract_btn = ft.ElevatedButton(
            text="마스터 데이터 엑셀 생성", 
            icon=ft.icons.DATA_SAVER_ON,
            style=ft.ButtonStyle(
                bgcolor=ft.colors.BLUE_800, color=ft.colors.WHITE, 
                padding=20, shape=ft.RoundedRectangleBorder(radius=8)
            ),
            on_click=self.handle_extract
        )

        # 전체 레이아웃 조합
        header = ft.Column([
            ft.Text("다면평가 마스터 데이터 추출 엔진", size=28, weight="bold", color=ft.colors.BLUE_900),
            ft.Text("Reactive Viewer에 주입할 정답지(Master Excel)를 생성합니다.", color="grey"),
            ft.Text("※ 특정 평가자 그룹이 누락된 경우 엔진이 설정된 표준 비율에 맞춰 자동으로 100% 팽창 분배합니다.", color=ft.colors.BLUE_700, size=12),
        ], spacing=5)

        main_card = ft.Container(
            content=ft.Column([
                ft.Row([year_dropdown], alignment=ft.MainAxisAlignment.END),
                ft.Divider(height=20, color="transparent"),
                ft.Row([exec_panel, team_panel], alignment=ft.MainAxisAlignment.CENTER, spacing=30),
                ft.Divider(height=30, color="transparent"),
                ft.Row([extract_btn, self.loading_ring, self.status_text], alignment=ft.MainAxisAlignment.CENTER, spacing=20)
            ]),
            bgcolor="white", padding=40, border_radius=15, width=900,
            shadow=ft.BoxShadow(blur_radius=15, color=ft.colors.with_opacity(0.05, "black"))
        )

        return ft.Column([header, ft.Container(height=20), main_card], alignment=ft.MainAxisAlignment.START, horizontal_alignment=ft.CrossAxisAlignment.CENTER)

    def _create_weight_panel(self, group_name, default_std):
        def _input(label, default_val):
            field = ft.TextField(label=label, value=str(default_val), width=85, text_align="right", suffix_text="%")
            return field

        # 표준 가중치 필드
        std_peer = _input("동료", default_std[0])
        std_sub = _input("부하", default_std[1])
        std_sup = _input("상사", default_std[2])
        
        # 저장
        self.inputs[group_name] = {
            'std_peer': std_peer, 'std_sub': std_sub, 'std_sup': std_sup
        }

        std_row = ft.Row([std_peer, std_sub, std_sup], spacing=15, alignment=ft.MainAxisAlignment.CENTER)

        panel = ft.Container(
            content=ft.Column([
                ft.Text(f"{group_name} 가중치 설정", size=18, weight="bold", color=ft.colors.BLUE_GREY_900),
                ft.Container(height=10),
                std_row,
            ], horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            border=ft.border.all(1, "#dee2e6"), padding=25, border_radius=10, width=380, bgcolor="#f8f9fa"
        )
        return panel

    def change_year(self, e):
        self.current_year = e.control.value

    def _parse_weights(self):
        # UI에서 입력한 값을 Logic이 이해할 수 있는 소수점(0.0~1.0) 딕셔너리로 변환 및 100% 검증
        configs = {}
        for group in ['임원', '팀장']:
            fields = self.inputs[group]
            try:
                std_p = float(fields['std_peer'].value)
                std_sub = float(fields['std_sub'].value)
                std_sup = float(fields['std_sup'].value)
            except ValueError:
                raise ValueError(f"{group} 그룹에 숫자가 아닌 값이 입력되었습니다.")

            if (std_p + std_sub + std_sup) != 100.0:
                raise ValueError(f"{group}의 가중치 합계가 100%가 아닙니다.")

            configs[group] = {
                'standard': {'peer': std_p/100, 'subordinate': std_sub/100, 'superior': std_sup/100}
            }
        return configs

    def handle_extract(self, e):
        try:
            # 1. 가중치 파싱 및 검증
            weight_configs = self._parse_weights()
            
            self.loading_ring.visible = True
            self.status_text.value = "데이터 연산 및 병합 중..."
            self.status_text.color = "blue"
            self.page.update()

            # 2. 경로 설정
            # [버그 픽스 1] exe 임시 폴더 함정 회피 -> Logic이 잡고 있는 '진짜 최상위 경로'를 참조하도록 변경
            output_dir = os.path.join(self.logic.base_path, "Excel_Reports")
            os.makedirs(output_dir, exist_ok=True)
            file_name = f"{self.current_year}년_다면평가_마스터데이터.xlsx"
            save_path = os.path.join(output_dir, file_name)
            
            # [추가 방어막] 절대 경로로 변환하여 정확한 위치 추적 및 기존 찌꺼기 파일 삭제
            abs_save_path = os.path.abspath(save_path)
            if os.path.exists(abs_save_path):
                try:
                    os.remove(abs_save_path)
                except:
                    pass

            # 3. 엔진 구동 (가중치 설정값 주입)
            self.logic.export_master_excel(self.current_year, abs_save_path, weight_configs)
            
            # [버그 픽스 2] 로데이터가 없어서 파일이 안 만들어졌을 경우의 방어 로직 (Silent Pass 차단)
            if not os.path.exists(abs_save_path):
                raise Exception(f"[{self.current_year}년 개인별 평가_raw data] 폴더를 찾을 수 없거나 데이터가 비어 있습니다.")
            
            # [핵심] 성공 시, 헷갈리지 않게 C:\ 부터 시작하는 절대경로를 그대로 화면에 출력!
            self.status_text.value = f"추출 완료!\n저장위치: {abs_save_path}"
            self.status_text.color = "green"
            
            self.page.snack_bar = ft.SnackBar(ft.Text("마스터 데이터 엑셀 생성이 완료되었습니다."), bgcolor=ft.colors.GREEN_700)
            self.page.snack_bar.open = True
            
        except ValueError as ve:
            self.status_text.value = "입력 오류 발생"
            self.status_text.color = "red"
            self.page.snack_bar = ft.SnackBar(ft.Text(str(ve)), bgcolor=ft.colors.RED_700)
            self.page.snack_bar.open = True
        except Exception as ex:
            print(traceback.format_exc())
            self.status_text.value = "오류 발생"
            self.status_text.color = "red"
            self.page.snack_bar = ft.SnackBar(ft.Text(f"오류: {ex}"), bgcolor=ft.colors.RED_700)
            self.page.snack_bar.open = True
        finally:
            self.loading_ring.visible = False
            self.page.update()
