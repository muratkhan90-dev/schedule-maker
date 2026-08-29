import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Автоматическое расписание", page_icon="📅", layout="wide")

st.title("📅 Автоматическое составление расписания")
st.markdown("Загрузите файл **сағат аты жөні сынып.xlsx** и получите готовое расписание без накладок")

# --- Функция для чтения нового формата (3 колонки) ---
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

# --- Функция для создания расписания ---
def generate_schedule(classes, class_teachers, subjects_data, days, periods):
    # Простейший алгоритм: на каждый день/урок ставим предмет, у которого есть часы
    schedule = {}
    for cls in classes:
        schedule[cls] = {}
        for day in days:
            schedule[cls][day] = {}
            for period in range(1, periods + 1):
                schedule[cls][day][period] = None
                
        # Берем предметы для класса
        items = []
        for subject, teachers in subjects_data.items():
            for teacher, hours in teachers.items():
                if teacher in class_teachers.get(cls, {}):
                    items.append((subject, teacher, int(class_teachers[cls][teacher])))
        
        # Распределяем предметы по дням и урокам (очень упрощенно)
        idx = 0
        for day in days:
            for period in range(1, periods + 1):
                if idx < len(items):
                    schedule[cls][day][period] = {
                        'subject': items[idx][0],
                        'teacher': items[idx][1]
                    }
                    idx += 1
    return schedule

# --- Функция для проверки конфликтов ---
def check_conflicts(schedule, days, periods, classes):
    conflicts = []
    for day in days:
        for period in range(1, periods + 1):
            teachers_today = {}
            for cls in classes:
                item = schedule[cls][day][period]
                if item:
                    teacher = item['teacher']
                    if teacher not in teachers_today:
                        teachers_today[teacher] = []
                    teachers_today[teacher].append(cls)
            
            for teacher, cls_list in teachers_today.items():
                if len(cls_list) > 1:
                    conflicts.append(f"{day} {period}-урок: {teacher} занят в {', '.join(cls_list)}")
    return conflicts

# --- Интерфейс загрузки файла ---
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

        if 'Класс' not in df.columns:
            df.rename(columns={df.columns[0]: 'Класс'}, inplace=True)
        if 'Предмет' not in df.columns:
            df.rename(columns={df.columns[1]: 'Предмет'}, inplace=True)
        if 'Учитель и часы' not in df.columns:
            df.rename(columns={df.columns[2]: 'Учитель и часы'}, inplace=True)

        classes, class_teachers, subjects_data = parse_schedule_file(df)
        
        st.success(f"✅ Найдено {len(classes)} классов")

        # Создаем расписание (5 дней, 7 уроков)
        days = ['Пн', 'Вт', 'Ср', 'Чт', 'Пт']
        periods = 7
        
        schedule = generate_schedule(classes, class_teachers, subjects_data, days, periods)
        conflicts = check_conflicts(schedule, days, periods, classes)

        # Показываем расписание
        st.subheader("📅 Расписание уроков")
        for cls in classes:
            with st.expander(f"Класс {cls}"):
                data = []
                for day in days:
                    row = {'День': day}
                    for period in range(1, periods + 1):
                        item = schedule[cls][day][period]
                        if item:
                            row[f"{period}-урок"] = f"{item['subject']} ({item['teacher']})"
                        else:
                            row[f"{period}-урок"] = ""
                    data.append(row)
                st.table(pd.DataFrame(data))

        # Проверка конфликтов
        st.subheader("⚠️ Проверка конфликтов")
        if conflicts:
            st.error(f"Найдено {len(conflicts)} конфликтов")
            for c in conflicts:
                st.write(f"❌ {c}")
        else:
            st.success("✅ Конфликтов нет! Расписание составлено корректно.")

        # Кнопка скачивания
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Записываем расписание в Excel
            for cls in classes:
                data = []
                for day in days:
                    row = {'День': day}
                    for period in range(1, periods + 1):
                        item = schedule[cls][day][period]
                        if item:
                            row[f"{period}-урок"] = f"{item['subject']} ({item['teacher']})"
                        else:
                            row[f"{period}-урок"] = ""
                    data.append(row)
                pd.DataFrame(data).to_excel(writer, index=False, sheet_name=cls)
            
            # Записываем нагрузку
            load_rows = []
            for cls in classes:
                for subject, teachers in subjects_data.items():
                    for teacher, hours in teachers.items():
                        if teacher in class_teachers.get(cls, {}):
                            load_rows.append({
                                'Класс': cls,
                                'Предмет': subject,
                                'Учитель': teacher,
                                'Часы': hours
                            })
            pd.DataFrame(load_rows).to_excel(writer, index=False, sheet_name='Нагрузка')
            
        processed_data = output.getvalue()

        st.download_button(
            label="📥 Скачать готовый файл Excel",
            data=processed_data,
            file_name="Расписание_нагрузка.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
