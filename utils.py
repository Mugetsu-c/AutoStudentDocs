import requests
import os
import pandas as pd
from datetime import datetime, timedelta
from jinja2 import Environment, FileSystemLoader
from docx import Document
from docxtpl import DocxTemplate


def download_csv(url, save_path):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            if os.path.exists(save_path):
                os.remove(save_path)
            with open(save_path, 'wb') as f:
                f.write(response.content)
    except Exception as e:
        print(f"An error occurred: {e}")


def create_certificate_jinja(template_name, save_path, data):
    template_path = f"./templates/{template_name}.docx"
    template = DocxTemplate(template_path)
    template.render(data)
    template.save(save_path)

    
def read_filtered_data(csv_path, time_period, start_date=None, end_date=None):
    df = pd.read_csv(csv_path)
    now = datetime.now()

    if time_period == "Последние 24 часа":
        delta = timedelta(hours=24)
        min_time = now - delta
        df['Отметка времени'] = pd.to_datetime(df['Отметка времени'], format='%d.%m.%Y')
        filtered_df = df[df['Отметка времени'] >= min_time]
    elif time_period == "Последние 48 часов":
        delta = timedelta(hours=48)
        min_time = now - delta
        df['Отметка времени'] = pd.to_datetime(df['Отметка времени'], format='%d.%m.%Y')
        filtered_df = df[df['Отметка времени'] >= min_time]
    elif time_period == "Последние 72 часа":
        delta = timedelta(hours=72)
        min_time = now - delta
        df['Отметка времени'] = pd.to_datetime(df['Отметка времени'], format='%d.%m.%Y')
        filtered_df = df[df['Отметка времени'] >= min_time]
    elif time_period == "Произвольная дата":
        df['Отметка времени'] = pd.to_datetime(df['Отметка времени'], format='%d.%m.%Y')
        start_date = pd.to_datetime(start_date)
        end_date = pd.to_datetime(end_date)
        filtered_df = df[(df['Отметка времени'] >= start_date) & (df['Отметка времени'] <= end_date)]
    else:
        return []

    formatted_records = []
    for record in filtered_df.to_dict('records'):
        formatted_record = [
            record['Отметка времени'].strftime('%d.%m.%Y'),
            record['ФИО (полностью)'],
            record['Группа'],
            str(record['На каком курсе вы обучаетесь?']),  # Добавьте str() здесь
            record['Куда нужна справка?'],
            str(record['В каком количестве нужна справка ?'])
        ]
        formatted_records.append(', '.join(formatted_record))
    
    return formatted_records
