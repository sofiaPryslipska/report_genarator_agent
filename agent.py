import streamlit as st
import requests
import json
import pypdf
import docx
import os
import re
from dotenv import load_dotenv
from docx.shared import Pt

load_dotenv()

API_KEY = os.getenv('API_KEY')
tools = [
    {
        "type": "function",
        "function": {
            "name": "extract_text_from_file",
            "description": "Читає текст з файлу методички (PDF або DOCX).",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Шлях до файлу"}
                },
                "required": ["file_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "generate_report",
            "description": "Генерує фінальний структурований звіт у форматі DOCX.",
            "parameters": {
                "type": "object",
                "properties": {
                    "course_name": {"type": "string", "description": "Повна назва навчальної дисципліни"},
                    "lab_number": {"type": "string", "description": "Номер лабораторної роботи"},
                    "student_name": {"type": "string", "description": "Прізвище та ініціали студента"},
                    "group_name": {"type": "string", "description": "Шифр групи"},
                    "teacher_name": {"type": "string", "description": "Прізвище та ініціали викладача"},
                    "topic": {"type": "string", "description": "Тема лабораторної роботи (без слова 'Тема роботи')"},
                    "goal": {"type": "string", "description": "Мета лабораторної роботи (без слова 'Мета роботи')"},
                    "tasks": {"type": "string", "description": "Перелік завдань лабораторної роботи"},
                    "workflow": {"type": "string", "description": "Детальний покроковий опис ходу роботи, бібліотек та логіки коду"},
                    "program_code": {"type": "string", "description": "Повний чистий лістинг коду програми без обгорток ```"},
                    "conclusions": {"type": "string", "description": "Висновки до лабораторної роботи"}
                },
                "required": [
                    "course_name", "lab_number", "student_name", "group_name", "teacher_name",
                    "topic", "goal", "tasks", "workflow", "program_code", "conclusions"
                ]
            }
        }
    }
]


def extract_text_from_file(file_path):
    if not os.path.exists(file_path):
        return "Помилка: Файл не знайдено."
    ext = file_path.lower().split('.')[-1]
    text = ""
    try:
        if ext == "pdf":
            reader = pypdf.PdfReader(file_path)
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        elif ext == "docx":
            doc = docx.Document(file_path)
            for para in doc.paragraphs:
                text += para.text + "\n"
        return text
    except Exception as e:
        return f"Помилка при читанні файлу: {str(e)}"


def _add_formatted_paragraph(target_paragraph, text):
    clean_line = text.strip().replace('`', '')
    if not clean_line:
        return
    new_p = target_paragraph.insert_paragraph_before()
    new_p.paragraph_format.space_after = Pt(6)
    new_p.paragraph_format.line_spacing = 1.15
    parts = re.split(r'(\*\*.*?\*\*)', clean_line)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            run = new_p.add_run(part[2:-2])
            run.bold = True
        else:
            run = new_p.add_run(part.replace('*', ''))
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)


def _add_section_heading(target_paragraph, title):
    new_p = target_paragraph.insert_paragraph_before()
    new_p.paragraph_format.space_before = Pt(12)
    new_p.paragraph_format.space_after = Pt(4)
    run = new_p.add_run(title)
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)


def generate_report(course_name, lab_number, student_name, group_name, teacher_name,
                    topic, goal, tasks, workflow, program_code, conclusions):
    template_path = 'template.docx'
    if not os.path.exists(template_path):
        return f"Помилка: Шаблон '{template_path}' не знайдено."
    try:
        doc = docx.Document(template_path)
        replacements = {
            "{{COURSE_NAME}}": course_name,
            "{{LAB_NUMBER}}": str(lab_number),
            "{{STUDENT_NAME}}": student_name,
            "{{GROUP_NAME}}": group_name,
            "{{TEACHER_NAME}}": teacher_name
        }
        for paragraph in doc.paragraphs:
            for tag, actual_text in replacements.items():
                if tag in paragraph.text:
                    paragraph.text = paragraph.text.replace(tag, actual_text)

            if "{{REPORT_BODY}}" in paragraph.text:
                sections = [
                    ("1. Тема роботи", topic),
                    ("2. Мета роботи", goal),
                    ("3. Завдання", tasks),
                    ("4. Хід роботи", workflow),
                ]
                for title, content in sections:
                    _add_section_heading(paragraph, title)
                    for line in str(content).split('\n'):
                        _add_formatted_paragraph(paragraph, line)

                _add_section_heading(paragraph, "5. Код програми")
                code_lines = str(program_code).strip().split('\n')
                for line in code_lines:
                    if line.strip().startswith('```'):
                        continue
                    new_p = paragraph.insert_paragraph_before()
                    new_p.paragraph_format.space_before = Pt(0)
                    new_p.paragraph_format.space_after = Pt(0)
                    run = new_p.add_run(line)
                    run.font.name = 'Consolas'
                    run.font.size = Pt(10)

                _add_section_heading(paragraph, "6. Висновки")
                for line in str(conclusions).split('\n'):
                    _add_formatted_paragraph(paragraph, line)

                paragraph._element.getparent().remove(paragraph._element)

        output_filename = f"Lab_{lab_number}_{student_name.replace(' ', '_')}.docx"
        doc.save(output_filename)
        return f"Успіх! Звіт збережено як '{output_filename}'."
    except Exception as e:
        return f"Помилка при генерації: {str(e)}"


def send_to_llm(messages):
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    request = {
        'model': 'gemini-2.5-flash',
        'messages': messages,
        'tools': tools,
        'reasoning_effort': 'low'
    }
    response = requests.post(
        'https://generativelanguage.googleapis.com/v1beta/openai/chat/completions',
        headers=headers,
        json=request
    )
    if response.status_code != 200:
        st.error(f"API ERROR: {response.json()}")
        return None

    data = response.json()
    choice = data['choices'][0]
    msg = choice['message']

    if not msg.get('content') and not msg.get('tool_calls'):
        st.warning(f"Модель повернула порожню відповідь (finish_reason: {choice.get('finish_reason')}). Напишіть «Продовжуй» або повторіть запит.")
        return None

    return msg


# Налаштування інтерфейсу
st.set_page_config(page_title="Агент звітів", page_icon="📝", layout="centered")
st.title("Автоматичний генератор звітів 📝")

if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "system",
            "content": (
                "Ти академічний агент-помічник для створення звітів з лабораторних робіт. "
                "Правила роботи:\n"
                "1. Якщо завантажено методичку, прочитай її через extract_text_from_file, щоб взяти Тему, Мету та Завдання.\n"
                "2. Запитай у користувача його Прізвище та ініціали, групу, Прізвище та ініціали викладача, повну назву предмету, номер лаби та код програми, якщо чогось не вистачає.\n"
                "3. Коли всі дані зібрано, виклич generate_report і заповни всі поля українською мовою в академічному стилі. У поле program_code передай чистий код користувача без ```."
            )
        }
    ]

uploaded_file = st.file_uploader("Завантажте методичку (PDF/DOCX)", type=["pdf", "docx"])
if uploaded_file and "file_uploaded" not in st.session_state:
    file_path = uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.session_state.file_uploaded = True
    system_msg = f"[Системне повідомлення: Завантажено методичку '{file_path}'. Прочитай її за допомогою інструменту.]"
    st.session_state.messages.append({"role": "user", "content": system_msg})

for msg in st.session_state.messages:
    if msg["role"] in ["user", "assistant"] and not msg.get("content", "").startswith("[Системне"):
        if msg.get("content"):
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

if user_message := st.chat_input("Введіть повідомлення або вставте код..."):
    st.session_state.messages.append({"role": "user", "content": user_message})
    with st.chat_message("user"):
        st.markdown(user_message)

    with st.spinner("Агент думає..."):
        llm_msg = send_to_llm(st.session_state.messages)
        if llm_msg:
            st.session_state.messages.append(llm_msg)

        steps = 0
        while llm_msg and llm_msg.get('tool_calls') and steps < 4:
            steps += 1
            for tool in llm_msg['tool_calls']:
                tool_name = tool['function']['name']
                tool_id = tool['id']
                args = json.loads(tool['function']['arguments'])

                if tool_name == 'extract_text_from_file':
                    result = extract_text_from_file(args.get('file_path'))
                elif tool_name == 'generate_report':
                    result = generate_report(
                        args.get('course_name'),
                        args.get('lab_number'),
                        args.get('student_name'),
                        args.get('group_name'),
                        args.get('teacher_name'),
                        args.get('topic'),
                        args.get('goal'),
                        args.get('tasks'),
                        args.get('workflow'),
                        args.get('program_code'),
                        args.get('conclusions')
                    )
                    st.success(result)

                    output_filename = f"Lab_{args.get('lab_number')}_{args.get('student_name').replace(' ', '_')}.docx"
                    if os.path.exists(output_filename):
                        st.session_state.generated_report = output_filename
                else:
                    result = f'Інструмент {tool_name} не існує'

                st.session_state.messages.append({
                    "role": "tool",
                    "tool_call_id": tool_id,
                    "name": tool_name,
                    "content": result
                })

            llm_msg = send_to_llm(st.session_state.messages)
            if llm_msg:
                st.session_state.messages.append(llm_msg)

        if llm_msg and llm_msg.get('content'):
            with st.chat_message("assistant"):
                st.markdown(llm_msg['content'])

if "generated_report" in st.session_state and os.path.exists(st.session_state.generated_report):
    with open(st.session_state.generated_report, "rb") as file:
        st.download_button(
            label="📥 Завантажити готовий звіт",
            data=file.read(),
            file_name=st.session_state.generated_report,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
