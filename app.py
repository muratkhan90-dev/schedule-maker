import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Автоматическое расписание", page_icon="📅", layout="wide")

st.title("📅 Автоматическое составление расписания")
st.markdown("Загрузите файл **сағат аты жөні сынып.xlsx** и получите готовое расписание без накладок")

# Функция для обработки нового формата (Класс, Предмет, Учитель и часы)
def parse_schedule_file(df):
    classes = []
    class_teachers = {}
    subjects_data = {}

    for index, row in df.iterrows():
        cls = str(row.get('Класс', '')).strip()
        subject = str(row.get('Предмет', '')).strip()
        value = str(row.get('Учитель и часы', '')).strip()

        if not cls or cls == 'nan' or not subject or subject == 'nan':
            continue

        if cls not in classes:
            classes.append(cls)

        if value and value != 'nan':
            parts = value.split('/')
            for part in parts:
                part = part.strip()
                if not part:
                    continue
                
                tokens = part.split()
                if len(tokens) >= 2:
                    teacher_name = " ".join(tokens[:-1])
                    hours = tokens[-1]
                    try:
                        hours = float(hours.replace(',', '.'))
                    except ValueError:
                        hours = 0
                    
                    if subject not in subjects_data:
                        subjects_data[subject] = {}
                    if teacher_name not in subjects_data[subject]:
                        subjects_data[subject][teacher_name] = 0
                    subjects_data[subject][teacher_name] += hours
                    
                    if cls not in class_teachers:
                        class_teachers[cls] = {}
                    if teacher_name not in class_teachers[cls]:
                        class_teachers[cls][teacher_name] = 0
                    class_teachers[cls][teacher_name] += hours
                
                else:
                    teacher_name = part
                    if subject not in subjects_data:
                        subjects_data[subject] = {}
                    if teacher_name not in subjects_data[subject]:
                        subjects_data[subject][teacher_name] = 0
                    subjects_data[subject][teacher_name] += 0
                    
                    if cls not in class_teachers:
                        class_teachers[cls] = {}
                    if teacher_name not in class_teachers[cls]:
                        class_teachers[cls][teacher_name] = 0
                    class_teachers[cls][teacher_name] += 0

    return classes, class_teachers, subjects_data

# Функция для создания итогового листа (упрощенная версия)
def create_load_sheet(classes, class_teachers, subjects_data):
    rows = []
    for cls in classes:
        for subject, teachers in subjects_data.items():
            for teacher, hours in teachers.items():
                if teacher in class_teachers.get(cls, {}):
                    rows.append({
                        'Класс': cls,
                        'Предмет': subject,
                        'Учитель': teacher,
                        'Часы': hours
                    })
    return pd.DataFrame(rows)

# Интерфейс загрузки файла
uploaded_file = st.file_uploader("📁 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    with st.spinner("⏳ Обработка файла..."):
        # Читаем файл
        df = pd.read_excel(uploaded_file)
        
        # Автоматически переименовываем колонки
        df.columns = [str(c).strip() for c in df.columns]
        rename_map = {}
        for col in df.columns:
            low_col = col.lower()
            if 'класс' in low_col:
                rename_map[col] = 'Класс'
            elif 'предмет' in low_col:
                rename_map[col] = 'Предмет'
            elif 'учитель' in low_col or 'часы' in low_col:
                rename_map[col] = 'Учитель и часы'
        df.rename(columns=rename_map, inplace=True)

        # Если не нашли по названиям, берем первые 3 колонки
        if 'Класс' not in df.columns:
            df.rename(columns={df.columns[0]: 'Класс'}, inplace=True)
        if 'Предмет' not in df.columns:
            df.rename(columns={df.columns[1]: 'Предмет'}, inplace=True)
        if 'Учитель и часы' not in df.columns:
            df.rename(columns={df.columns[2]: 'Учитель и часы'}, inplace=True)

        classes, class_teachers, subjects_data = parse_schedule_file(df)
        
        st.success(f"✅ Найдено {len(classes)} классов")

        # Отображаем данные
        st.subheader("📊 Данные по классам")
        load_df = create_load_sheet(classes, class_teachers, subjects_data)
        st.dataframe(load_df, use_container_width=True)

        # Кнопка скачивания
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            load_df.to_excel(writer, index=False, sheet_name='Нагрузка')
        processed_data = output.getvalue()

        st.download_button(
            label="📥 Скачать готовый файл Excel",
            data=processed_data,
            file_name="Расписание_нагрузка.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
