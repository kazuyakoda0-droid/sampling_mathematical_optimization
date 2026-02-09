# -*- coding: utf-8 -*-
"""
サンプルデータ作成スクリプト
作業者名と業務名を匿名化したデモ用データを生成
"""

import pandas as pd
from datetime import datetime, timedelta

# 匿名化した作業者データ
persons_data = {
    'name': ['山田太郎', '鈴木一郎', '田中花子', '佐藤次郎', '高橋三郎', 
             '伊藤美咲', '渡辺健太', '中村優子', '小林誠', '加藤裕子',
             '吉田直樹', '山本愛', '松本大輔', '井上恵'],
    'priority': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14],
    'skill': [5, 4, 4, 3, 5, 3, 4, 3, 5, 4, 3, 4, 3, 2],
    'trouble': [1, 2, 1, 3, 1, 2, 1, 2, 1, 1, 2, 1, 2, 3],
    'personality': [5, 4, 5, 3, 4, 4, 5, 4, 5, 4, 3, 5, 4, 3],
    'strength': [5, 4, 3, 4, 5, 3, 4, 3, 5, 3, 4, 3, 4, 2],
    'ship': [1, 1, 0, 1, 1, 0, 1, 0, 1, 0, 1, 0, 1, 0],
    'driving': [1, 1, 1, 1, 1, 1, 1, 0, 1, 1, 1, 1, 1, 0],
    'navigation': [1, 0, 0, 1, 1, 0, 0, 0, 1, 0, 1, 0, 0, 0],
    'notes': ['', '', '', '', '', '分析優先', '', '月火のみ', '', '', '', '', '', '']
}

# 匿名化した業務データ
tasks_data = {
    'task_id': list(range(1, 21)),
    'task_name': [
        '東地区 水質調査A', '東地区 水質調査B', '西地区 土壌サンプリング',
        '北地区 大気測定', '南地区 河川調査', '中央 排水検査',
        '工業地区 煙突測定', '港湾 海水サンプリング', '森林 植生調査',
        '湖沼 底質調査', '処理場 汚泥分析', '農地 土壌検査',
        '住宅地 騒音測定', '商業地 排気測定', 'ダム 水質モニタリング',
        '浄水場 検体採取', '下水処理 放流水検査', '工場 廃水分析',
        '海岸 砂浜調査', '山間部 湧水採取'
    ],
    'area': ['東地区', '東地区', '西地区', '北地区', '南地区', '中央',
             '工業地区', '港湾', '森林', '湖沼', '処理場', '農地',
             '住宅地', '商業地', 'ダム', '浄水場', '下水処理', '工場', '海岸', '山間部'],
    'required_workers': [2, 2, 2, 1, 2, 2, 2, 3, 2, 2, 2, 1, 1, 1, 2, 2, 2, 2, 2, 1],
    'required_skill': [4, 3, 3, 4, 3, 4, 5, 4, 3, 4, 4, 3, 2, 2, 4, 3, 3, 4, 3, 3],
    'required_strength': [3, 3, 4, 3, 4, 3, 4, 5, 4, 4, 3, 3, 2, 2, 3, 3, 3, 3, 4, 4],
    'urgency': [3, 3, 3, 4, 3, 4, 3, 3, 2, 3, 4, 2, 2, 2, 4, 3, 3, 4, 2, 2],
    'ship_work': [0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    'navigation_required': [0, 0, 0, 0, 0, 0, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    'duration': [1.0, 1.0, 1.5, 0.5, 2.0, 1.0, 1.5, 3.0, 2.0, 2.0, 1.0, 1.0, 0.5, 0.5, 1.5, 1.0, 1.0, 1.5, 2.0, 1.5]
}

# サンプルスケジュール（2025年4月）
schedule_dates = []
schedule_tasks = []

# 4月の平日に業務を割り当て
base_date = datetime(2025, 4, 1)
task_names = tasks_data['task_name']
task_idx = 0

for day in range(30):
    current_date = base_date + timedelta(days=day)
    # 土日はスキップ
    if current_date.weekday() >= 5:
        continue
    
    # 1日に2-4件の業務を割り当て
    num_tasks = (day % 3) + 2
    for _ in range(num_tasks):
        schedule_dates.append(current_date)
        schedule_tasks.append(task_names[task_idx % len(task_names)])
        task_idx += 1

schedule_data = {
    'date': schedule_dates,
    'task_name': schedule_tasks
}

# Excelファイルとして保存
with pd.ExcelWriter('sample_data_demo.xlsx', engine='openpyxl') as writer:
    pd.DataFrame(persons_data).to_excel(writer, sheet_name='person', index=False)
    pd.DataFrame(tasks_data).to_excel(writer, sheet_name='database', index=False)

# スケジュールファイル
schedule_df = pd.DataFrame(schedule_data)
schedule_df['date_serial'] = schedule_df['date'].apply(
    lambda x: (x - datetime(1899, 12, 31)).days + 366
)
schedule_df[['date_serial', 'task_name']].to_excel('sample_schedule_demo.xlsx', sheet_name='Sheet1', index=False, header=False)

print("サンプルデータを作成しました:")
print("  - sample_data_demo.xlsx (作業者・業務マスタ)")
print("  - sample_schedule_demo.xlsx (スケジュール)")
print("\nこれらのファイルを使用する場合は app.py の設定を変更してください。")
print("\nこれらのファイルを使用する場合は app.py の設定を変更してください。")
