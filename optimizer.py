# -*- coding: utf-8 -*-
"""
最適化エンジンモジュール
PuLPを使用してサンプリング業務の人員配置を最適化
"""

import pulp
import pandas as pd
from typing import Dict, List, Tuple, Any, Optional
from data_loader import (
    load_all_data, 
    get_schedule_by_date, 
    match_task_to_database
)


class SamplingOptimizer:
    """サンプリング業務人員配置最適化クラス"""
    
    def __init__(self, persons_df: pd.DataFrame, tasks_df: pd.DataFrame):
        self.persons = persons_df
        self.tasks = tasks_df
        self.person_names = persons_df['name'].tolist()
        
    def calculate_match_score(self, person: pd.Series, task: Dict) -> float:
        """
        作業者と業務のマッチングスコアを計算
        スコアが高いほど適任
        """
        score = 0.0
        
        # 技量マッチング (最重要)
        skill_diff = person['skill'] - task.get('required_skill', 3)
        if skill_diff >= 0:
            score += 30 + skill_diff * 5  # 要件を満たしていればボーナス
        else:
            score -= abs(skill_diff) * 20  # 不足はペナルティ
        
        # 体力マッチング
        strength_diff = person['strength'] - task.get('required_strength', 3)
        if strength_diff >= 0:
            score += 20 + strength_diff * 3
        else:
            score -= abs(strength_diff) * 15
        
        # 船上作業能力
        if task.get('ship_work', 1) >= 3:  # 船上作業が必要
            ship_diff = person['ship'] - task.get('ship_work', 1)
            if ship_diff >= 0:
                score += 15 + ship_diff * 2
            else:
                score -= abs(ship_diff) * 10
        
        # 操船能力
        if task.get('navigation_required', 1) >= 3:  # 操船が必要
            nav_ability = person['navigation']
            if nav_ability > 0:
                score += 10 + nav_ability * 2
            else:
                score -= 30  # 操船できない場合は大きなペナルティ
        
        return score
    
    def is_person_available(self, person: pd.Series, date_str: str) -> bool:
        """
        作業者が指定日に作業可能かチェック
        """
        notes = str(person.get('notes', ''))
        
        # 「分析優先」の作業者は割り当て可能だが優先度を下げる
        # ここでは利用可能として返す
        
        # 「月火のみ可能」のチェック
        if '月火のみ' in notes:
            from datetime import datetime
            date = datetime.strptime(date_str, '%Y-%m-%d')
            weekday = date.weekday()
            if weekday not in [0, 1]:  # 月曜=0, 火曜=1
                return False
        
        return True
    
    def get_priority_penalty(self, person: pd.Series) -> float:
        """
        作業者の優先度に基づくペナルティを取得
        分析優先の作業者にはペナルティを付与
        """
        notes = str(person.get('notes', ''))
        if '分析優先' in notes:
            return -20  # 分析優先の作業者はサンプリングへの割り当てを減らす
        return 0
    
    def optimize_day(self, date_str: str, task_list: List[str]) -> Dict[str, List[str]]:
        """
        1日分の最適化を実行
        
        Args:
            date_str: 対象日付 (YYYY-MM-DD形式)
            task_list: その日に実行する業務名のリスト
            
        Returns:
            Dict[task_name, List[person_name]]: 各業務に割り当てられた作業者
        """
        # 業務情報を取得
        tasks_info = []
        for task_name in task_list:
            task_info = match_task_to_database(task_name, self.tasks)
            task_info['original_name'] = task_name
            tasks_info.append(task_info)
        
        # 利用可能な作業者を取得
        available_persons = []
        for _, person in self.persons.iterrows():
            if self.is_person_available(person, date_str):
                available_persons.append(person)
        
        if not available_persons:
            return {task['original_name']: [] for task in tasks_info}
        
        # エリア別にタスクをグループ化
        area_tasks = {}
        for t_idx, task in enumerate(tasks_info):
            area = task.get('area', '不明')
            if area not in area_tasks:
                area_tasks[area] = []
            area_tasks[area].append(t_idx)
        
        # 最適化問題の定義
        prob = pulp.LpProblem(f"Sampling_Assignment_{date_str}", pulp.LpMaximize)
        
        # 決定変数: x[p,t] = 作業者pが業務tに割り当てられる場合1
        x = {}
        for p_idx, person in enumerate(available_persons):
            for t_idx, task in enumerate(tasks_info):
                x[p_idx, t_idx] = pulp.LpVariable(
                    f"x_{p_idx}_{t_idx}", 
                    cat=pulp.LpBinary
                )
        
        # エリア用変数: y[p,area] = 作業者pがエリアareaの業務に割り当てられる場合1
        y = {}
        for p_idx in range(len(available_persons)):
            for area in area_tasks.keys():
                area_safe = str(area).replace(' ', '_').replace('・', '_').replace('　', '_')
                y[p_idx, area] = pulp.LpVariable(
                    f"y_{p_idx}_{area_safe}",
                    cat=pulp.LpBinary
                )
        
        # 目的関数: マッチングスコアの最大化
        objective = []
        for p_idx, person in enumerate(available_persons):
            for t_idx, task in enumerate(tasks_info):
                score = self.calculate_match_score(person, task)
                score += self.get_priority_penalty(person)
                objective.append(score * x[p_idx, t_idx])
        
        prob += pulp.lpSum(objective)
        
        # 制約1: 各作業者は1つのエリアの業務のみに従事可能
        for p_idx in range(len(available_persons)):
            prob += pulp.lpSum([y[p_idx, area] for area in area_tasks.keys()]) <= 1
        
        # 制約2: 作業者がエリアに割り当てられた場合のみ、そのエリアの業務に従事可能
        for p_idx in range(len(available_persons)):
            for area, t_indices in area_tasks.items():
                for t_idx in t_indices:
                    prob += x[p_idx, t_idx] <= y[p_idx, area]
        
        # 制約3: 各業務に必要人数を割り当て（可能な限り）
        for t_idx, task in enumerate(tasks_info):
            required = task.get('required_workers', 1)  # デフォルト1名
            prob += pulp.lpSum([x[p_idx, t_idx] for p_idx in range(len(available_persons))]) <= required
        
        # 求解
        prob.solve(pulp.PULP_CBC_CMD(msg=0))
        
        # 結果の取得
        result = {task['original_name']: [] for task in tasks_info}
        
        if prob.status == pulp.LpStatusOptimal:
            for p_idx, person in enumerate(available_persons):
                for t_idx, task in enumerate(tasks_info):
                    if x[p_idx, t_idx].value() > 0.5:
                        result[task['original_name']].append(person['name'])
        
        return result
    
    def optimize_schedule(self, schedule_df: pd.DataFrame) -> Dict[str, Dict[str, List[str]]]:
        """
        全スケジュールの最適化を実行
        
        Returns:
            Dict[date_str, Dict[task_name, List[person_name]]]
        """
        schedule_by_date = get_schedule_by_date(schedule_df)
        
        all_results = {}
        for date_str, task_list in schedule_by_date.items():
            day_result = self.optimize_day(date_str, task_list)
            all_results[date_str] = day_result
        
        return all_results


def run_optimization(optimization_file: str, schedule_file: str) -> Dict:
    """
    最適化を実行するメイン関数
    """
    # データ読み込み
    persons_df, tasks_df, schedule_df = load_all_data(optimization_file, schedule_file)
    
    # 最適化実行
    optimizer = SamplingOptimizer(persons_df, tasks_df)
    results = optimizer.optimize_schedule(schedule_df)
    
    return {
        'results': results,
        'persons': persons_df.to_dict('records'),
        'schedule_dates': sorted(results.keys())
    }


if __name__ == '__main__':
    import sys
    import json
    sys.stdout.reconfigure(encoding='utf-8')
    
    result = run_optimization(
        'Mathematical Optimization_sampling_2026.xlsx',
        'sample.xlsx'
    )
    
    print("=== 最適化結果 ===")
    for date in result['schedule_dates'][:5]:
        print(f"\n{date}:")
        for task_name, persons in result['results'][date].items():
            if persons:
                print(f"  {task_name}: {', '.join(persons)}")
            else:
                print(f"  {task_name}: (未割当)")
