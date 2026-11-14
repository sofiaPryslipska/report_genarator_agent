import streamlit as st
import requests
import json
import pypdf
import docx
import os
from dotenv import load_dotenv
import re
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
                    "course_name": {"type": "string"},
                    "lab_number": {"type": "string"},
                    "student_name": {"type": "string"},
                    "group_name": {"type": "string"},
                    "teacher_name": {"type": "string"},
                    "report_body": {
                        "type": "string",
                        "description": (
                            "Повний структурований текст звіту українською мовою в академічному стилі. "
                            "ОБОВ'ЯЗКОВО має містити такі 6 розділів із заголовками:\n"
                            "1. Тема роботи (взяти з методички або сформулювати за кодом).\n"
                            "2. Мета роботи (взяти з методички або сформулювати відповідно до завдання).\n"
                            "3. Завдання (опис завдання лабораторної роботи з методички).\n"
                            "4. Хід роботи (детальний покроковий опис алгоритму, використаних бібліотек, функцій та логіки програми).\n"
                            "5. Код програми (повний лістинг коду користувача без скорочень).\n"
                            "6. Висновки (підсумок виконаної роботи, чого навчилися та які результати отримано)."
                            "Кожен із 6 заголовків розділів обов'язково починай із символу # (наприклад: # 1. Тема роботи)."
                        )
                    }
                },
                "required": ["course_name", "lab_number", "student_name", "group_name", "teacher_name", "report_body"]
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


def generate_report(course_name, lab_number, student_name, group_name, teacher_name, report_body):
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
                in_code_block = False
                for line in report_body.split('\n'):
                    stripped = line.strip()

                    # Перемикаємо режим коду при зустрічі ```
                    if stripped.startswith('```'):
                        in_code_block = not in_code_block
                        continue

                    if in_code_block:
                        new_p = paragraph.insert_paragraph_before()
                        new_p.paragraph_format.space_before = Pt(0)
                        new_p.paragraph_format.space_after = Pt(0)
                        run = new_p.add_run(line)
                        run.font.name = 'Consolas'
                        run.font.size = Pt(10)
                    else:
                        if not stripped:
                            continue

                        clean_line = stripped.replace('`', '')

                        # Заголовки розділів (# Тема роботи, # Мета роботи...)
                        if clean_line.startswith('#'):
                            heading_text = re.sub(r'^#+\s*', '', clean_line).replace('**', '')
                            new_p = paragraph.insert_paragraph_before()
                            new_p.paragraph_format.space_before = Pt(12)
                            new_p.paragraph_format.space_after = Pt(4)
                            run = new_p.add_run(heading_text)
                            run.bold = True
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(14)
                        else:
                            # Звичайні абзаци та списки з підтримкою жирного тексту (**...**)
                            new_p = paragraph.insert_paragraph_before()
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

                # Видаляємо сам тег {{REPORT_BODY}} після вставки тексту
                paragraph._element.getparent().remove(paragraph._element)

        output_filename = f"Lab_{lab_number}_{student_name.replace(' ', '_')}.docx"
        doc.save(output_filename)
        return f"Успіх! Звіт збережено як '{output_filename}'."
    except Exception as e:
        return f"Помилка при генерації: {str(e)}"


def send_to_llm(messages):
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    request = {'model': 'gemini-2.5-flash', 'messages': messages, 'tools': tools}
    response = requests.post('https://generativelanguage.googleapis.com/v1beta/openai/chat/completions',
                             headers=headers, json=request)
    if response.status_code != 200:
        st.error(f"API ERROR: {response.json()}")
        return None
    return response.json()['choices'][0]['message']


# Налаштування інтерфейсу
st.set_page_config(page_title="Агент звітів", page_icon="📝", layout="centered")
st.title("Автоматичний генератор звітів 📝")

# Ініціалізація історії повідомлень із системним промптом
if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "system",
            "content": (
                "Ти академічний агент-помічник для створення звітів з лабораторних робіт. "
                "Правила роботи:\n"
                "1. Якщо завантажено методичку, обов'язково прочитай її через extract_text_from_file, щоб взяти звідти точну Тему, Мету та Завдання.\n"
                "2. Запитай у користувача його Прізвище та ініціали, групу,  Прізвище та ініціали викладача, повну назву предмету, номер лаби та код програми, якщо чогось не вистачає.\n"
                "3. Коли всі дані зібрано, НЕ вставляй у report_body лише сирий код. Самостійно напиши повноцінний звіт із 6 розділів: Тема роботи, Мета роботи, Завдання, Хід роботи (детальний розбір коду та логіки), Код програми (повністю) та Висновки.\n"
                "4. Виклич generate_report і передай цей структурований текст у report_body."
            )
        }
    ]

# Обробка завантаження файлу
uploaded_file = st.file_uploader("Завантажте методичку (PDF/DOCX)", type=["pdf", "docx"])
if uploaded_file and "file_uploaded" not in st.session_state:
    file_path = uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    st.session_state.file_uploaded = True
    system_msg = f"[Системне повідомлення: Завантажено методичку '{file_path}'. Прочитай її за допомогою інструменту.]"
    st.session_state.messages.append({"role": "user", "content": system_msg})

# Відображення історії чату (приховуємо системні та tool повідомлення для краси)
for msg in st.session_state.messages:
    if msg["role"] in ["user", "assistant"] and not msg.get("content", "").startswith("[Системне"):
        if msg.get("content"):
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

# Поле вводу
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
                        args.get('course_name'), args.get('lab_number'),
                        args.get('student_name'), args.get('group_name'),
                        args.get('teacher_name'), args.get('report_body')
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
