import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(page_title="Автоматическое расписание", page_icon="📅", layout="wide")

st.title("📅 Автоматическое составление расписания")
st.markdown("Загрузите файл **сағат аты жөні сынып.xlsx** и получите готовое расписание без накладок")

def parse_schedule_file(df):
    classes = []
    class_teachers = {}
    
    for col in range(1, len(df.columns)):
        val = str(df.iloc[0, col])
        if any(x in val for x in ['а,', 'ә,', 'б,']):
            parts = val.split(',')
            if len(parts) >= 1:
                cls = parts[0].strip()
                if cls and cls[0].isdigit():
                    classes.append(cls)
                    if len(parts) > 1:
                        teacher = parts[1].strip().split()[0] if parts[1].strip() else 'Вакант'
                        class_teachers[cls] = teacher
                    else:
                        class_teachers[cls] = 'Вакант'
    
    subjects_data = {}
    for row in range(1, len(df)):
        subject = str(df.iloc[row, 0]).strip()
        if not subject or subject == 'nan':
            continue
            
        subjects_data[subject] = {}
        for idx, cls in enumerate(classes):
            col = idx + 1
            if col < len(df.columns):
                val = str(df.iloc[row, col]).strip()
                if val and val != 'nan':
                    teachers = []
                    hours = 0
                    
                    hour_match = re.search(r'(\d+[.,]?\d*)$', val)
                    if hour_match:
                        hours = float(hour_match.group(1).replace(',', '.'))
                        val = val[:hour_match.start()].strip()
                    
                    if '/' in val:
                        teachers = [t.strip() for t in val.split('/') if t.strip() and t.strip() != 'вакант']
                    else:
                        if val and val != 'вакант':
                            teachers = [val.strip()]
                    
                    if teachers and hours > 0:
                        subjects_data[subject][cls] = {'teachers': teachers, 'hours': hours}
    
    return classes, class_teachers, subjects_data

def create_load_sheet(classes, class_teachers, subjects_data):
    rows = []
    for subject, classes_data in subjects_data.items():
        for cls, info in classes_data.items():
            for teacher in info['teachers']:
                if teacher and teacher != 'вакант':
                    rows.append({
                        'Класс': cls,
                        'Предмет': subject,
                        'Учитель': teacher,
                        'Часов': info['hours']
                    })
    return pd.DataFrame(rows)

def generate_schedule(load_df, class_teachers):
    days = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ']
    classes = sorted(load_df['Класс'].unique())
    max_lessons = {5: 6, 6: 6, 7: 6, 8: 6, 9: 7, 10: 7, 11: 7}
    
    class_subjects = {}
    for cls in classes:
        df_cls = load_df[load_df['Класс'] == cls]
        subjects = []
        for _, row in df_cls.iterrows():
            for _ in range(int(row['Часов'])):
                subjects.append(f"{row['Предмет']} ({row['Учитель']})")
        class_subjects[cls] = subjects
    
    schedule = {cls: {day: [] for day in days} for cls in classes}
    teacher_busy = {day: {i: set() for i in range(1, 9)} for day in days}
    
    for cls in classes:
        cls_num = int(cls[0])
        max_day = max_lessons.get(cls_num, 6)
        class_teacher = class_teachers.get(cls, '')
        
        if class_teacher and class_teacher != 'Вакант':
            schedule[cls]['ПН'].append(f"КЧ ({class_teacher})")
            teacher_busy['ПН'][1].add(class_teacher)
        else:
            schedule[cls]['ПН'].append("КЧ (Вакант)")
        
        if class_teacher and class_teacher != 'Вакант' and class_subjects.get(cls):
            for i, subj in enumerate(class_subjects[cls]):
                if class_teacher in subj:
                    schedule[cls]['ПН'].append(subj)
                    teacher_busy['ПН'][2].add(class_teacher)
                    class_subjects[cls].pop(i)
                    break
            else:
                if class_subjects[cls]:
                    subj = class_subjects[cls].pop(0)
                    schedule[cls]['ПН'].append(subj)
                    teacher = subj.split('(')[-1].replace(')', '').strip()
                    if teacher:
                        teacher_busy['ПН'][2].add(teacher)
        else:
            if class_subjects.get(cls):
                subj = class_subjects[cls].pop(0)
                schedule[cls]['ПН'].append(subj)
                teacher = subj.split('(')[-1].replace(')', '').strip()
                if teacher:
                    teacher_busy['ПН'][2].add(teacher)
    
    for day in days:
        for cls in classes:
            cls_num = int(cls[0])
            max_day = max_lessons.get(cls_num, 6)
            
            for lesson_num in range(len(schedule[cls][day]) + 1, max_day + 1):
                if class_subjects.get(cls):
                    added = False
                    for i, subj in enumerate(class_subjects[cls]):
                        teacher = subj.split('(')[-1].replace(')', '').strip() if '(' in subj else ''
                        if teacher and teacher not in teacher_busy[day][lesson_num]:
                            schedule[cls][day].append(subj)
                            teacher_busy[day][lesson_num].add(teacher)
                            class_subjects[cls].pop(i)
                            added = True
                            break
                    if not added and class_subjects[cls]:
                        subj = class_subjects[cls].pop(0)
                        schedule[cls][day].append(subj)
                        teacher = subj.split('(')[-1].replace(')', '').strip() if '(' in subj else ''
                        if teacher:
                            teacher_busy[day][lesson_num].add(teacher)
                else:
                    schedule[cls][day].append("")
    
    return schedule

def check_conflicts(schedule, days, classes):
    conflicts = []
    for day in days:
        busy = {}
        for cls in classes:
            for lesson_num, lesson in enumerate(schedule[cls][day], 1):
                if lesson and '(' in lesson:
                    teacher = lesson.split('(')[-1].replace(')', '').strip()
                    if teacher and teacher != 'Вакант':
                        key = (teacher, lesson_num)
                        if key in busy:
                            conflicts.append(f"{day} {lesson_num} урок: {teacher} в {cls} и {busy[key]}")
                        else:
                            busy[key] = cls
    return conflicts

uploaded_file = st.file_uploader("📤 Загрузите Excel-файл", type=["xlsx"])

if uploaded_file:
    with st.spinner("⏳ Обработка файла..."):
        df = pd.read_excel(uploaded_file, header=None)
        classes, class_teachers, subjects_data = parse_schedule_file(df)
        
        st.success(f"✅ Найдено {len(classes)} классов")
        
        with st.expander("📋 Классные руководители"):
            st.json(class_teachers)
        
        load_df = create_load_sheet(classes, class_teachers, subjects_data)
        
        st.subheader("📊 Лист 'Нагрузка'")
        st.dataframe(load_df, use_container_width=True)
        
        schedule = generate_schedule(load_df, class_teachers)
        
        conflicts = check_conflicts(schedule, ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ'], classes)
        
        if conflicts:
            st.error(f"⚠️ Найдено {len(conflicts)} конфликтов:")
            for c in conflicts:
                st.write(f"- {c}")
        else:
            st.success("✅ Конфликтов нет! Расписание составлено корректно.")
        
        st.subheader("📅 Расписание")
        tabs = st.tabs(['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ'])
        
        for tab, day in zip(tabs, ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ']):
            with tab:
                day_data = {}
                for cls in classes:
                    day_data[cls] = schedule[cls][day]
                df_day = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in day_data.items()]))
                df_day.index = [f"{i+1}" for i in range(len(df_day))]
                st.dataframe(df_day, use_container_width=True)
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            load_df.to_excel(writer, sheet_name='Нагрузка', index=False)
            
            for day in ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ']:
                day_data = {}
                for cls in classes:
                    day_data[cls] = schedule[cls][day]
                df_day = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in day_data.items()]))
                df_day.index = [f"{i+1}" for i in range(len(df_day))]
                df_day.to_excel(writer, sheet_name=day, index_label='Урок')
            
            teacher_summary = {}
            for day in ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ']:
                for cls in classes:
                    for lesson in schedule[cls][day]:
                        if lesson and '(' in lesson:
                            teacher = lesson.split('(')[-1].replace(')', '').strip()
                            if teacher and teacher != 'Вакант':
                                if teacher not in teacher_summary:
                                    teacher_summary[teacher] = []
                                teacher_summary[teacher].append(f"{day} {cls}")
            
            summary_rows = []
            for teacher, lessons in teacher_summary.items():
                summary_rows.append({
                    'Учитель': teacher,
                    'Количество уроков': len(lessons),
                    'Классы': '; '.join(set(lessons))
                })
            df_summary = pd.DataFrame(summary_rows)
            df_summary.to_excel(writer, sheet_name='Сводка_учителей', index=False)
        
        output.seek(0)
        
        st.download_button(
            label="📥 Скачать расписание (Excel)",
            data=output,
            file_name="расписание_автомат.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Загрузите Excel-файл с нагрузкой, чтобы начать")