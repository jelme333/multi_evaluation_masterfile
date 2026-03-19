import os
import glob
import pandas as pd
import re
import numpy as np

class EvaluationLogic:
    # ================================================================
    # [A안 적용] __init__ 에서 매핑 파일 고정 로드 제거
    #   → export_master_excel(year, ...) 호출 시점에 연도에 맞는
    #     매핑 파일을 동적으로 찾아 로드하는 구조로 변경
    # ================================================================
    def __init__(self, base_path):
        self.base_path = base_path
        self.mapping_df = pd.DataFrame()   # 연도 확정 전 빈 상태
        self.area_order = []               # 매핑 파일 로드 시 동적으로 채워짐
        
    # ================================================================
    # 연도별 매핑 파일 탐색
    #   규칙: "{year} 세코닉스 리더 역량행동지표.xlsx (또는 .csv)"
    #   예시: "2025 세코닉스 리더 역량행동지표.xlsx"
    # ================================================================
    def _find_mapping_file(self, year):
        """
        base_path 에서 '{year} 세코닉스 리더 역량행동지표' 파일을 탐색.
        xlsx → csv 순으로 시도하며, 없으면 None 반환.
        """
        for ext in ['.xlsx', '.csv']:
            candidate = os.path.join(
                self.base_path,
                f"{year} 세코닉스 리더 역량행동지표{ext}"
            )
            if os.path.exists(candidate):
                return candidate
        return None

    # ================================================================
    # 매핑 파일 로드 (area_order 도 동적 추출)
    # ================================================================
    def _load_mapping(self, path):
        try:
            if path.endswith('.csv'):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path)

            df.columns = [str(c).strip() for c in df.columns]
            if '평가항목' in df.columns:
                df.rename(columns={'평가항목': '역량행동지표'}, inplace=True)

            df['clean_indicator'] = df['역량행동지표'].apply(self._normalize_header_for_matching)
            df['영역'] = df['영역'].astype(str).apply(
                lambda x: re.sub(' +', ' ', x.replace('\n', ' ')).strip()
            )

            # [A안 핵심] area_order: 매핑 파일의 영역 등장 순서를 그대로 추출
            #   → 연도별 영역 순서가 달라져도 자동 대응
            seen = []
            for area in df['영역']:
                if area not in seen:
                    seen.append(area)
            self.area_order = seen

            return df[['영역', '역량행동지표', 'clean_indicator']]
        except Exception as e:
            print(f"매핑 파일 로드 실패: {e}")
            return pd.DataFrame()

    # ================================================================
    # 연도별 매핑 로드 공개 메서드 (export_master_excel 에서 호출)
    # ================================================================
    def load_mapping_for_year(self, year):
        """
        지정 연도의 매핑 파일을 찾아 self.mapping_df / self.area_order 갱신.
        파일이 없으면 FileNotFoundError 를 발생시켜 UI 에서 명확한 오류 메시지를 받게 함.
        """
        mapping_path = self._find_mapping_file(year)
        if not mapping_path:
            raise FileNotFoundError(
                f"[{year}년] 역량행동지표 매핑 파일을 찾을 수 없습니다.\n"
                f"'{year} 세코닉스 리더 역량행동지표.xlsx' 파일이\n"
                f"아래 경로에 있는지 확인해 주세요:\n{self.base_path}"
            )
        self.mapping_df = self._load_mapping(mapping_path)
        print(f"[매핑 로드 완료] {os.path.basename(mapping_path)}  |  영역 순서: {self.area_order}")

    # ================================================================
    # 이하 기존 로직 코드 — 일체 변경 없음
    # ================================================================

    def _normalize_text(self, text):
        if pd.isna(text): return ""
        text = str(text)
        text = re.sub(r'[0-9\.\,"\'\-\(\)\[\]]', '', text)
        text = re.sub(' +', ' ', text).strip()
        text = text.replace('\n', ' ').replace('\t', ' ')
        return text.strip()

    def _normalize_header_for_matching(self, text):
        if pd.isna(text): return ""
        text = str(text)
        text = re.sub(r'[0-9\.\,"\'\-\(\)\[\]]', '', text)
        text = text.replace(' ', '').replace('\n', '').replace('\t', '').strip()
        return text

    def get_file_list(self, year, group):
        target_path = os.path.join(self.base_path, f"{year} 개인별 평가_raw data", group)
        if not os.path.exists(target_path): return {}

        files = glob.glob(os.path.join(target_path, "*.xlsx"))
        people_files = {}
        
        for file_path in files:
            filename = os.path.basename(file_path)
            rater_type = None
            if filename.startswith("□"): rater_type = 'peer'
            elif filename.startswith("△"): rater_type = 'subordinate'
            elif filename.startswith("☆"): rater_type = 'superior'
            
            if rater_type:
                # 스마트 파싱: 파일명에서 불필요한 기호, 연도, 단어를 모두 지우고 이름만 추출
                name_part = filename
                name_part = re.sub(r'[□△☆]', '', name_part) 
                name_part = name_part.replace(str(year), '') 
                name_part = name_part.replace('.xlsx', '') 
                name_part = name_part.replace('_', ' ').replace('-', ' ') 
                
                for word in ['다면평가', '평가', '리더십', '결과', '개인별', 'raw', 'data', '임원', '팀장']:
                    name_part = name_part.replace(word, '')
                    
                name_part = name_part.strip()
                
                if name_part.startswith('년 '):
                    name_part = name_part[2:].strip()
                elif name_part.startswith('년'):
                    name_part = name_part[1:].strip()
                    
                name = name_part.replace(' ', '')
                
                if name:
                    if name not in people_files:
                        people_files[name] = {'peer': None, 'subordinate': None, 'superior': None}
                    people_files[name][rater_type] = file_path

        return people_files

    def _clean_score(self, value):
        if pd.isna(value): return None
        val_str = str(value).strip()
        try: return float(val_str)
        except: pass
        match = re.search(r'(\d)(\.\d+)?', val_str)
        return float(match.group(0)) if match else None

    def _is_meaningful_feedback(self, text):
        if pd.isna(text) or not str(text).strip(): return False
        clean_text = re.sub(r'[\s\.\,\-\~_!@#$%^&*()\[\]\?]', '', str(text))
        if not clean_text: return False
        meaningless_words = [
            '없음', '없습니다', '없슴', '없슴다', '없어요', '없어용', '없음요',
            '의견없음', '의견없습니다', '의견없슴', '특별한의견없음', '특별한의견없습니다',
            '특이사항없음', '특이사항없습니다', '특이사항없슴',
            '해당없음', '해당없습니다', '해당사항없음', '해당사항없습니다',
            'na', 'none', 'null', '무', '잘모르겠습니다', '모름', '모르겠습니다'
        ]
        if clean_text.lower() in meaningless_words: return False
        return True

    def calculate_person_score(self, files_dict, weight_config=None):
        if weight_config and 'standard' in weight_config:
            base_weights = weight_config['standard']
        else:
            base_weights = {'peer': 0.5, 'subordinate': 0.3, 'superior': 0.2}

        loaded_data = {'peer': None, 'subordinate': None, 'superior': None}
        raw_means = {'peer': np.nan, 'subordinate': np.nan, 'superior': np.nan}
        feedback_list = []
        all_items = set(self.mapping_df['clean_indicator'].tolist())

        for rater_type, file_path in files_dict.items():
            if file_path and os.path.exists(file_path):
                try:
                    df = pd.read_excel(file_path)
                    if not df.empty and df.shape[1] >= 2:
                        df = df.drop_duplicates(keep='first').sort_index()

                        feedbacks = df.iloc[:, -1].dropna().tolist()
                        for fb in feedbacks:
                            cleaned_fb = str(fb).strip()
                            if self._is_meaningful_feedback(cleaned_fb):
                                rtype_kor = '동료' if rater_type=='peer' else ('부하' if rater_type=='subordinate' else '상사')
                                feedback_list.append({'type': rtype_kor, 'content': cleaned_fb})

                        df_score = df.iloc[:, 1:-1].copy()
                        df_score.columns = [self._normalize_header_for_matching(c) for c in df_score.columns]
                        valid_cols = [c for c in df_score.columns if c in all_items]
                        df_score = df_score[valid_cols].map(self._clean_score)

                        if not df_score.empty:
                            def check_all_six(row):
                                valid_values = row.dropna()
                                if valid_values.empty: return False
                                return np.all(np.isclose(valid_values, 6.0))
                            def check_all_one(row):
                                valid_values = row.dropna()
                                if valid_values.empty: return False
                                return np.all(np.isclose(valid_values, 1.0))

                            is_all_six = df_score.apply(check_all_six, axis=1)
                            six_indices = df_score.index[is_all_six].tolist()
                            is_all_one = df_score.apply(check_all_one, axis=1)
                            one_indices = df_score.index[is_all_one].tolist()
                            
                            if len(six_indices) == 1: df_score = df_score.drop(six_indices) 
                            elif len(six_indices) > 1: df_score = df_score.drop(six_indices[1:]) 
                            if len(one_indices) == 1: df_score = df_score.drop(one_indices) 
                            elif len(one_indices) > 1: df_score = df_score.drop(one_indices[1:]) 

                        if df_score.empty: continue

                        mean_series = df_score.mean(numeric_only=True)
                        loaded_data[rater_type] = mean_series
                        total_mean = df_score.stack().mean()
                        raw_means[rater_type] = total_mean if pd.notna(total_mean) else np.nan
                except Exception as e:
                    print(f"데이터 처리 중 오류 발생: {e}")

        # 가중치 자동 팽창 (Auto-balancing)
        has_peer = pd.notna(raw_means['peer'])
        has_subordinate = pd.notna(raw_means['subordinate'])
        has_superior = pd.notna(raw_means['superior'])

        active_w_sum = 0.0
        if has_peer: active_w_sum += base_weights.get('peer', 0)
        if has_subordinate: active_w_sum += base_weights.get('subordinate', 0)
        if has_superior: active_w_sum += base_weights.get('superior', 0)

        if active_w_sum > 0:
            current_weights = {
                'peer': (base_weights.get('peer', 0) / active_w_sum) if has_peer else 0.0,
                'subordinate': (base_weights.get('subordinate', 0) / active_w_sum) if has_subordinate else 0.0,
                'superior': (base_weights.get('superior', 0) / active_w_sum) if has_superior else 0.0
            }
            
            w_strs = []
            if has_peer: w_strs.append(f"동료 {current_weights['peer']*100:.1f}%")
            if has_subordinate: w_strs.append(f"부하 {current_weights['subordinate']*100:.1f}%")
            if has_superior: w_strs.append(f"상사 {current_weights['superior']*100:.1f}%")
            
            desc_str = " + ".join(w_strs).replace(".0%", "%") 
            
            if np.isclose(active_w_sum, 1.0):
                weight_desc = f"표준 가중치 ({desc_str})"
            else:
                weight_desc = f"자동 가중치 재분배 ({desc_str})"
        else:
            current_weights = {'peer': 0.0, 'subordinate': 0.0, 'superior': 0.0}
            weight_desc = "유효 평가 데이터 없음"

        indicator_results = []
        for _, row in self.mapping_df.iterrows():
            indicator_clean = row['clean_indicator']
            area = row['영역']

            p_score = loaded_data['peer'].get(indicator_clean, np.nan) if loaded_data['peer'] is not None else np.nan
            sub_score = loaded_data['subordinate'].get(indicator_clean, np.nan) if loaded_data['subordinate'] is not None else np.nan
            sup_score = loaded_data['superior'].get(indicator_clean, np.nan) if loaded_data['superior'] is not None else np.nan
            
            weighted_sum = 0
            weight_total = 0

            if pd.notna(p_score):
                weighted_sum += p_score * current_weights['peer']
                weight_total += current_weights['peer']
            if pd.notna(sub_score):
                weighted_sum += sub_score * current_weights['subordinate']
                weight_total += current_weights['subordinate']
            if pd.notna(sup_score):
                weighted_sum += sup_score * current_weights['superior']
                weight_total += current_weights['superior']

            total_weighted_item_score = (weighted_sum / weight_total) if weight_total > 0 else np.nan

            indicator_results.append({
                '영역': area, '평가항목': row['역량행동지표'], '합산': total_weighted_item_score,
                '동료': p_score, '부하': sub_score, '상사': sup_score
            })

        result_df = pd.DataFrame(indicator_results)

        if not result_df.empty:
            area_scores = result_df.groupby('영역')[['합산', '동료', '부하', '상사']].mean().round(2).reset_index()
            area_scores['영역'] = pd.Categorical(area_scores['영역'], categories=self.area_order, ordered=True)
            area_scores = area_scores.sort_values('영역').dropna(subset=['영역'])
        else:
            area_scores = pd.DataFrame(columns=['영역', '합산', '동료', '부하', '상사'])

        final_w_sum = 0
        final_score_sum = 0
        
        if pd.notna(raw_means['peer']):
            final_score_sum += raw_means['peer'] * current_weights['peer']
            final_w_sum += current_weights['peer']
        if pd.notna(raw_means['subordinate']):
            final_score_sum += raw_means['subordinate'] * current_weights['subordinate']
            final_w_sum += current_weights['subordinate']
        if pd.notna(raw_means['superior']):
            final_score_sum += raw_means['superior'] * current_weights['superior']
            final_w_sum += current_weights['superior']
            
        unrounded_total_score = (final_score_sum / final_w_sum) if final_w_sum > 0 else 0.0
        total_final_score = round(unrounded_total_score, 2)
        
        rounded_raw_scores = {k: round(v, 2) if pd.notna(v) else np.nan for k, v in raw_means.items()}
        unrounded_raw_scores = {k: v if pd.notna(v) else np.nan for k, v in raw_means.items()}
        
        has_superior = has_superior  # 변수명 유지 (기존 반환 키 호환)

        # [신규 기능] 평가자 그룹별 강/약점 추출 (PPT 보고서 대응)
        def get_top_bottom(df, col_name):
            valid_df = df[df[col_name].notna()]
            if valid_df.empty:
                return [], []
            sorted_v = valid_df.sort_values(by=col_name, ascending=False)
            t2 = sorted_v.head(2).to_dict('records')
            b2 = sorted_v.tail(2).sort_values(by=col_name, ascending=True).to_dict('records')
            return t2, b2

        if not result_df.empty:
            sorted_df = result_df.sort_values(by='합산', ascending=False)
            top_2 = sorted_df.head(2).to_dict('records')
            bottom_2 = sorted_df.tail(2).sort_values(by='합산', ascending=True).to_dict('records')
            
            peer_strong, peer_weak = get_top_bottom(result_df, '동료')
            sub_strong, sub_weak = get_top_bottom(result_df, '부하')
            sup_strong, sup_weak = get_top_bottom(result_df, '상사')
        else:
            top_2 = []; bottom_2 = []
            peer_strong = []; peer_weak = []
            sub_strong = []; sub_weak = []
            sup_strong = []; sup_weak = []

        return {
            'area_scores': area_scores, 'total_score': total_final_score,
            'raw_scores': rounded_raw_scores, 'unrounded_total_score': unrounded_total_score,
            'unrounded_raw_scores': unrounded_raw_scores, 'weight_desc': weight_desc,
            'has_superior': has_superior, 'top_strong': top_2, 'top_weak': bottom_2,
            'peer_strong': peer_strong, 'peer_weak': peer_weak,
            'sub_strong': sub_strong, 'sub_weak': sub_weak,
            'sup_strong': sup_strong, 'sup_weak': sup_weak,
            'detail_df': result_df, 'feedbacks': feedback_list
        }

    def get_ranking_data(self, year, group, weight_config=None):
        files = self.get_file_list(year, group)
        ranking_list = []
        for name, file_paths in files.items():
            try:
                res = self.calculate_person_score(file_paths, weight_config)
                if res['total_score'] > 0:
                    ranking_list.append({
                        '이름': name, '합산': res['total_score'],
                        '합산_unrounded': res.get('unrounded_total_score', res['total_score']),
                        '동료_unrounded': res.get('unrounded_raw_scores', {}).get('peer', res['raw_scores']['peer']),
                        '부하_unrounded': res.get('unrounded_raw_scores', {}).get('subordinate', res['raw_scores']['subordinate']),
                        '상사_unrounded': res.get('unrounded_raw_scores', {}).get('superior', res['raw_scores']['superior'])
                    })
            except: continue
        
        df = pd.DataFrame(ranking_list)
        if not df.empty:
            df['Rank'] = df['합산_unrounded'].rank(ascending=False, method='first')
            df['동료_Rank'] = df['동료_unrounded'].rank(ascending=False, method='first')
            df['부하_Rank'] = df['부하_unrounded'].rank(ascending=False, method='first')
            df['상사_Rank'] = df['상사_unrounded'].rank(ascending=False, method='first')
            df = df.sort_values('Rank')
        return df

    def export_master_excel(self, year, output_path, weight_configs=None):
        # ── [A안 핵심] 추출 시작 전 해당 연도 매핑 파일을 동적으로 로드 ──
        #    파일이 없으면 FileNotFoundError 발생 → UI SnackBar 에 표시됨
        self.load_mapping_for_year(year)

        export_data = []
        for group in ['임원', '팀장']:
            group_weight = weight_configs.get(group) if weight_configs else None
            
            files = self.get_file_list(year, group)
            ranking_df = self.get_ranking_data(year, group, group_weight)
            
            for name, file_paths in files.items():
                try:
                    res = self.calculate_person_score(file_paths, group_weight)
                    if res['total_score'] == 0: continue
                        
                    row_data = {
                        '연도': year, '그룹': group, '이름': name,
                        '종합점수': res['total_score'], '동료평균': res['raw_scores']['peer'],
                        '부하평균': res['raw_scores']['subordinate'], '상사평균': res['raw_scores']['superior'],
                    }
                    
                    if not ranking_df.empty:
                        person_rank = ranking_df[ranking_df['이름'] == name]
                        if not person_rank.empty:
                            row_data['종합순위'] = person_rank.iloc[0]['Rank']
                            row_data['동료순위'] = person_rank.iloc[0]['동료_Rank']
                            row_data['부하순위'] = person_rank.iloc[0]['부하_Rank']
                            row_data['상사순위'] = person_rank.iloc[0]['상사_Rank']
                    
                    area_df = res['area_scores']
                    for _, a_row in area_df.iterrows():
                        area_name = a_row['영역']
                        row_data[f'{area_name}_합산'] = a_row['합산']
                        row_data[f'{area_name}_동료'] = a_row['동료']
                        row_data[f'{area_name}_부하'] = a_row['부하']
                        row_data[f'{area_name}_상사'] = a_row['상사']
                    
                    # 엑셀 열 추가 함수화
                    def fill_strong_weak_cols(row_dict, prefix, strong_list, weak_list, score_key):
                        for i in range(2):
                            if i < len(strong_list):
                                row_dict[f'{prefix}강점{i+1}_영역'] = strong_list[i]['영역']
                                row_dict[f'{prefix}강점{i+1}_항목'] = strong_list[i]['평가항목']
                                row_dict[f'{prefix}강점{i+1}_점수'] = strong_list[i][score_key]
                            else:
                                row_dict[f'{prefix}강점{i+1}_영역'] = ""
                                row_dict[f'{prefix}강점{i+1}_항목'] = ""
                                row_dict[f'{prefix}강점{i+1}_점수'] = ""
                            if i < len(weak_list):
                                row_dict[f'{prefix}약점{i+1}_영역'] = weak_list[i]['영역']
                                row_dict[f'{prefix}약점{i+1}_항목'] = weak_list[i]['평가항목']
                                row_dict[f'{prefix}약점{i+1}_점수'] = weak_list[i][score_key]
                            else:
                                row_dict[f'{prefix}약점{i+1}_영역'] = ""
                                row_dict[f'{prefix}약점{i+1}_항목'] = ""
                                row_dict[f'{prefix}약점{i+1}_점수'] = ""

                    fill_strong_weak_cols(row_data, "", res.get('top_strong', []), res.get('top_weak', []), '합산')
                    fill_strong_weak_cols(row_data, "동료_", res.get('peer_strong', []), res.get('peer_weak', []), '동료')
                    fill_strong_weak_cols(row_data, "부하_", res.get('sub_strong', []), res.get('sub_weak', []), '부하')
                    fill_strong_weak_cols(row_data, "상사_", res.get('sup_strong', []), res.get('sup_weak', []), '상사')

                    # ── [신규] 항목별 원점수를 컬럼으로 추가 (PPT 강·약 항목 재현용)
                    detail_df = res.get('detail_df')
                    if detail_df is not None and not detail_df.empty:
                        area_counter = {}
                        for _, item_row in detail_df.iterrows():
                            area = item_row['영역']
                            area_counter[area] = area_counter.get(area, 0) + 1
                            n = area_counter[area]
                            prefix = f'항목_{area}_{n}'
                            row_data[f'{prefix}_명']  = item_row['평가항목']
                            row_data[f'{prefix}_합산'] = round(item_row['합산'], 4) if pd.notna(item_row['합산']) else ''
                            row_data[f'{prefix}_동료'] = round(item_row['동료'], 4) if pd.notna(item_row['동료']) else ''
                            row_data[f'{prefix}_부하'] = round(item_row['부하'], 4) if pd.notna(item_row['부하']) else ''
                            row_data[f'{prefix}_상사'] = round(item_row['상사'], 4) if pd.notna(item_row['상사']) else ''

                    peer_fbs = [fb['content'] for fb in res['feedbacks'] if fb['type'] == '동료']
                    sub_fbs = [fb['content'] for fb in res['feedbacks'] if fb['type'] == '부하']
                    sup_fbs = [fb['content'] for fb in res['feedbacks'] if fb['type'] == '상사']
                    
                    row_data['동료_피드백'] = "\n".join(f"-- {fb}" for fb in peer_fbs)
                    row_data['부하_피드백'] = "\n".join(f"-- {fb}" for fb in sub_fbs)
                    row_data['상사_피드백'] = "\n".join(f"-- {fb}" for fb in sup_fbs)
                    
                    export_data.append(row_data)
                except Exception as e:
                    print(f"Excel Export Error for {name}: {e}")
                    continue
        
        if export_data:
            df_export = pd.DataFrame(export_data)
            df_export.to_excel(output_path, index=False, engine='openpyxl')
        else:
            raise Exception(
                f"'{year} 개인별 평가_raw data' 폴더 안에 유효한 데이터가 전혀 없습니다. "
                f"(폴더명이나 위치를 확인해 주세요)"
            )
