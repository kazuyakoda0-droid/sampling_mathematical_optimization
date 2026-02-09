# -*- coding: utf-8 -*-
"""
Flask Webアプリケーション
サンプリング業務人員配置最適化システム
"""

from flask import Flask, render_template, jsonify, request, send_file
import json
import io
from datetime import datetime
import pandas as pd
from optimizer import run_optimization, SamplingOptimizer
from data_loader import load_all_data, get_schedule_by_date

app = Flask(__name__)

# グローバル変数でデータをキャッシュ
# 本番データ
# OPTIMIZATION_FILE = 'Mathematical Optimization_sampling_2026.xlsx'
# SCHEDULE_FILE = 'sample.xlsx'

# デモ用サンプルデータ（匿名化済み）
OPTIMIZATION_FILE = 'sample_data_demo.xlsx'
SCHEDULE_FILE = 'sample_schedule_demo.xlsx'

cached_data = None
cached_results = None


def get_cached_data():
    """データをキャッシュから取得または読み込み"""
    global cached_data
    if cached_data is None:
        persons_df, tasks_df, schedule_df = load_all_data(OPTIMIZATION_FILE, SCHEDULE_FILE)
        cached_data = {
            'persons': persons_df,
            'tasks': tasks_df,
            'schedule': schedule_df
        }
    return cached_data


@app.route('/')
def index():
    """メインページ"""
    return render_template('index.html')


@app.route('/api/data')
def get_data():
    """初期データを取得"""
    data = get_cached_data()
    schedule_by_date = get_schedule_by_date(data['schedule'])
    
    dates = sorted(schedule_by_date.keys())
    persons = data['persons'][['name', 'skill', 'notes']].to_dict('records')
    
    date_summary = [
        {
            'date': date,
            'task_count': len(tasks),
            'tasks': tasks
        }
        for date, tasks in sorted(schedule_by_date.items())
    ]
    
    return jsonify({
        'dates': dates,
        'persons': persons,
        'schedule': date_summary
    })


@app.route('/api/optimize', methods=['POST'])
def optimize():
    """最適化を実行"""
    global cached_results
    
    try:
        result = run_optimization(OPTIMIZATION_FILE, SCHEDULE_FILE)
        cached_results = result['results']
        
        return jsonify({
            'success': True,
            'results': cached_results,
            'dates': result['schedule_dates']
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 500


@app.route('/api/optimize/day/<date_str>', methods=['POST'])
def optimize_day(date_str):
    """特定の日の最適化を実行"""
    data = get_cached_data()
    schedule_by_date = get_schedule_by_date(data['schedule'])
    
    if date_str not in schedule_by_date:
        return jsonify({
            'success': False,
            'error': f'Date {date_str} not found in schedule'
        }), 404
    
    optimizer = SamplingOptimizer(data['persons'], data['tasks'])
    result = optimizer.optimize_day(date_str, schedule_by_date[date_str])
    
    return jsonify({
        'success': True,
        'date': date_str,
        'results': result
    })


@app.route('/api/schedule/<date_str>')
def get_schedule_for_date(date_str):
    """特定の日のスケジュールを取得"""
    global cached_results
    
    data = get_cached_data()
    schedule_by_date = get_schedule_by_date(data['schedule'])
    
    if date_str not in schedule_by_date:
        return jsonify({
            'success': False,
            'error': f'Date {date_str} not found'
        }), 404
    
    tasks = schedule_by_date[date_str]
    assignments = {}
    if cached_results and date_str in cached_results:
        assignments = cached_results[date_str]
    
    return jsonify({
        'success': True,
        'date': date_str,
        'tasks': tasks,
        'assignments': assignments
    })


@app.route('/api/persons')
def get_persons():
    """作業者リストを取得"""
    data = get_cached_data()
    persons = data['persons'].to_dict('records')
    return jsonify({
        'success': True,
        'persons': persons
    })


@app.route('/api/download')
def download_results():
    """最適化結果をExcelファイルでダウンロード"""
    global cached_results
    
    if not cached_results:
        return jsonify({
            'success': False,
            'error': '最適化が実行されていません。先に最適化を実行してください。'
        }), 400
    
    # データを整形
    rows = []
    for date_str in sorted(cached_results.keys()):
        date_obj = datetime.strptime(date_str, '%Y-%m-%d')
        weekday_names = ['月', '火', '水', '木', '金', '土', '日']
        weekday = weekday_names[date_obj.weekday()]
        
        for task_name, persons in cached_results[date_str].items():
            rows.append({
                '日付': date_str,
                '曜日': weekday,
                '業務名': task_name,
                '担当者1': persons[0] if len(persons) > 0 else '',
                '担当者2': persons[1] if len(persons) > 1 else '',
                '担当者3': persons[2] if len(persons) > 2 else '',
                '担当者4': persons[3] if len(persons) > 3 else '',
            })
    
    df = pd.DataFrame(rows)
    
    # Excelファイルをメモリに書き込み
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='人員配置表', index=False)
    output.seek(0)
    
    filename = f'人員配置表_2025年4月_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


if __name__ == '__main__':
    print("サンプリング業務人員配置最適化システム")
    print("http://localhost:5000 でアクセスしてください")
    app.run(debug=True, host='0.0.0.0', port=5000)
