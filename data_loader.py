# -*- coding: utf-8 -*-
"""
データローダーモジュール
Excelファイルからサンプリング業務データを読み込む
"""

import pandas as pd
from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Any


def excel_serial_to_date(serial: int) -> datetime:
    """Excelのシリアル値を日付に変換"""
    # sample.xlsxのシリアル値46113 = 2025年4月1日に合わせて調整
    base_date = datetime(1899, 12, 31)
    return base_date + timedelta(days=serial - 366)


def load_person_data(filepath: str) -> pd.DataFrame:
    """
    作業者データを読み込む
    
    Returns:
        DataFrame with columns: name, priority, skill, trouble, personality, 
                               strength, ship, driving, navigation, notes
    """
    df = pd.read_excel(filepath, sheet_name='person')
    
    # 列名を英語に変換
    column_mapping = {
        'Unnamed: 0': 'name',
        '優先順位': 'priority',
        '技量': 'skill',
        'トラブル': 'trouble',
        '性情': 'personality',
        '力': 'strength',
        '船上': 'ship',
        '運転': 'driving',
        '操縦': 'navigation',
        '備考': 'notes'
    }
    df = df.rename(columns=column_mapping)
    
    # NaNを空文字に置換
    df['notes'] = df['notes'].fillna('')
    
    return df


def load_task_database(filepath: str) -> pd.DataFrame:
    """
    業務データベースを読み込む
    
    Returns:
        DataFrame with task information
    """
    df = pd.read_excel(filepath, sheet_name='database')
    
    # 列名を英語に変換
    column_mapping = {
        '番号': 'task_id',
        '業務名（略）': 'task_name',
        '地区': 'area',
        '人工': 'required_workers',
        '技量': 'required_skill',
        '体力': 'required_strength',
        '緊急対応': 'urgency',
        '船上': 'ship_work',
        '操船': 'navigation_required',
        '所要時間': 'duration'
    }
    df = df.rename(columns=column_mapping)
    
    # 所要時間の範囲表記を数値に変換（例: "0.5～3" -> 1.75）
    def parse_duration(val):
        if pd.isna(val):
            return 1.0
        if isinstance(val, str):
            if '～' in val or '~' in val:
                parts = val.replace('～', '~').split('~')
                try:
                    return (float(parts[0]) + float(parts[1])) / 2
                except:
                    return 1.0
            try:
                return float(val)
            except:
                return 1.0
        return float(val)
    
    df['duration'] = df['duration'].apply(parse_duration)
    
    return df


def load_schedule(filepath: str) -> pd.DataFrame:
    """
    スケジュールデータを読み込む
    
    Returns:
        DataFrame with columns: date, task_name
    """
    df = pd.read_excel(filepath, sheet_name='Sheet1', header=None)
    df.columns = ['date_serial', 'task_name']
    
    # 日付に変換
    df['date'] = df['date_serial'].apply(excel_serial_to_date)
    df['date_str'] = df['date'].dt.strftime('%Y-%m-%d')
    
    return df


def get_schedule_by_date(schedule_df: pd.DataFrame) -> Dict[str, List[str]]:
    """
    日付ごとにタスクをグループ化
    
    Returns:
        Dict[date_str, List[task_name]]
    """
    result = {}
    for _, row in schedule_df.iterrows():
        date_str = row['date_str']
        if date_str not in result:
            result[date_str] = []
        result[date_str].append(row['task_name'])
    return result


def match_task_to_database(task_name: str, database_df: pd.DataFrame) -> Dict[str, Any]:
    """
    スケジュールのタスク名をデータベースとマッチング
    部分一致で最も類似度の高いものを返す
    """
    task_name_clean = task_name.strip()
    
    # 完全一致を試行
    exact_match = database_df[database_df['task_name'] == task_name_clean]
    if len(exact_match) > 0:
        return exact_match.iloc[0].to_dict()
    
    # 部分一致を試行
    for _, row in database_df.iterrows():
        db_name = str(row['task_name']).strip()
        if db_name in task_name_clean or task_name_clean in db_name:
            return row.to_dict()
    
    # マッチしない場合はデフォルト値を返す（マスタ未登録は1名配置）
    return {
        'task_id': 0,
        'task_name': task_name_clean,
        'area': '不明',
        'required_workers': 1,  # マスタ未登録業務は1名
        'required_skill': 3,
        'required_strength': 3,
        'urgency': 3,
        'ship_work': 1,
        'navigation_required': 1,
        'duration': 1.0,
        'is_unregistered': True  # マスタ未登録フラグ
    }


def load_all_data(optimization_file: str, schedule_file: str) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    全てのデータを読み込む
    
    Returns:
        (persons_df, tasks_df, schedule_df)
    """
    persons_df = load_person_data(optimization_file)
    tasks_df = load_task_database(optimization_file)
    schedule_df = load_schedule(schedule_file)
    
    return persons_df, tasks_df, schedule_df


if __name__ == '__main__':
    # テスト実行
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    
    persons, tasks, schedule = load_all_data(
        'Mathematical Optimization_sampling_2026.xlsx',
        'sample.xlsx'
    )
    
    print("=== 作業者データ ===")
    print(persons.to_string())
    
    print("\n=== 業務データ (先頭10件) ===")
    print(tasks.head(10).to_string())
    
    print("\n=== スケジュール (先頭10件) ===")
    print(schedule.head(10).to_string())
    
    print("\n=== 日付別タスク ===")
    by_date = get_schedule_by_date(schedule)
    for date, task_list in list(by_date.items())[:3]:
        print(f"{date}: {len(task_list)}件")
