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
            "description": "Читає текст з файлу методички (PDF або DOCX). Використовуй лише для отримання теми, мети та умови завдання.",
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
            "description": (
                "Генерує DOCX-звіт. ПЕРЕДУМОВА ВИКЛИКУ: викликати ТІЛЬКИ тоді, коли користувач у чаті особисто надав УСІ 6 пунктів: "
                "1) ПІБ студента, 2) групу, 3) реальне ПІБ викладача, 4) назву предмету, 5) номер лаби, 6) власний готовий код програми. "
                "Якщо хоча б одного пункту немає — НЕ викликай цю функцію, а запитай відсутнє у користувача."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "course_name": {
                        "type": "string",
                        "description": "Повна назва навчальної дисципліни, вказана користувачем. ЗАБОРОНЕНО вигадувати або писати 'Не вказано'."
                    },
                    "lab_number": {
                        "type": "string",
                        "description": "Номер лабораторної роботи, вказаний користувачем. ЗАБОРОНЕНО вигадувати або писати 'Не вказано'."
                    },
                    "student_name": {
                        "type": "string",
                        "description": "Прізвище та ініціали студента, вказані користувачем. ЗАБОРОНЕНО вигадувати або писати 'Не вказано'."
                    },
                    "group_name": {
                        "type": "string",
                        "description": "Шифр групи, вказаний користувачем. ЗАБОРОНЕНО вигадувати або писати 'Не вказано'."
                    },
                    "teacher_name": {
                        "type": "string",
                        "description": "Реальне прізвище та ініціали викладача, вказані користувачем у чаті. КАТЕГОРИЧНО ЗАБОРОНЕНО писати 'Не вказано', 'Невідомо', 'Викладач' або вигадувати."
                    },
                    "topic": {
                        "type": "string",
                        "description": "Тема роботи: візьми з методички або сформулюй самостійно за наданим кодом користувача. НЕ запитуй у користувача."
                    },
                    "goal": {
                        "type": "string",
                        "description": "Мета роботи: візьми з методички або сформулюй самостійно за наданим кодом користувача. НЕ запитуй у користувача."
                    },
                    "tasks": {
                        "type": "string",
                        "description": "Завдання роботи: візьми з методички або сформулюй самостійно за наданим кодом користувача. НЕ запитуй у користувача."
                    },
                    "workflow": {
                        "type": "string",
                        "description": "Детальний покроковий опис ходу роботи, бібліотек та логіки на основі коду, який надіслав користувач."
                    },
                    "program_code": {
                        "type": "string",
                        "description": (
                            "ВИКЛЮЧНО готовий код програми, який користувач особисто надіслав у чаті (зі збереженням \\n та відступів, без обгорток ```). "
                            "КАТЕГОРИЧНО ЗАБОРОНЕНО писати код самостійно за завданням із методички!"
                        )
                    },
                    "conclusions": {
                        "type": "string",
                        "description": "Висновки до лабораторної роботи на основі виконаного коду користувача."
                    }
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


def _replace_tags_in_paragraph(paragraph, replacements):
    for tag, actual_text in replacements.items():
        if tag in paragraph.text:
            new_text = paragraph.text.replace(tag, actual_text)
            paragraph.text = ""
            run = paragraph.add_run(new_text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)


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

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        _replace_tags_in_paragraph(paragraph, replacements)

        for paragraph in doc.paragraphs:
            _replace_tags_in_paragraph(paragraph, replacements)

            if "{{REPORT_BODY}}" in paragraph.text:
                sections = [
                    ("Тема роботи", topic),
                    ("Мета роботи", goal),
                    ("Завдання", tasks),
                    ("Хід роботи", workflow),
                ]
                for title, content in sections:
                    _add_section_heading(paragraph, title)
                    for line in str(content).split('\n'):
                        _add_formatted_paragraph(paragraph, line)

                _add_section_heading(paragraph, "Код програми")
                raw_code = str(program_code).replace('\r\n', '\n').strip('\n')
                for line in raw_code.split('\n'):
                    if line.strip().startswith('```'):
                        continue
                    leading_spaces = len(line) - len(line.lstrip(' '))
                    formatted_line = ('\t' * (leading_spaces // 4)) + (' ' * (leading_spaces % 4)) + line.lstrip(' ')
                    new_p = paragraph.insert_paragraph_before()
                    new_p.paragraph_format.space_before = Pt(0)
                    new_p.paragraph_format.space_after = Pt(0)
                    new_p.paragraph_format.line_spacing = 1.0
                    run = new_p.add_run(formatted_line if formatted_line else " ")
                    run.font.name = 'Consolas'
                    run.font.size = Pt(10)

                _add_section_heading(paragraph, "Висновки")
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
        'model': 'gemini-3.5-flash-lite',
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


st.set_page_config(page_title="Агент звітів", page_icon="📝", layout="centered")
st.title("Автоматичний генератор звітів ")

if "messages" not in st.session_state:
    st.session_state.messages = [
        {
            "role": "system",
            "content": (
                "РОЛЬ: Ти суворий секретар-оформлювач академічних звітів. Ти НЕ є програмістом і НЕ вмієш писати програмний код.\n\n"
                "ДЖЕРЕЛА ДАНИХ (СУВОРО РОЗМЕЖОВУЙ):\n"
                "1. ДЖЕРЕЛО А — 6 ОБОВ'ЯЗКОВИХ ДАНИХ ВІД КОРИСТУВАЧА (мають бути явно написані користувачем у діалозі):\n"
                "   [1] Прізвище та ініціали студента (student_name)\n"
                "   [2] Шифр групи (group_name)\n"
                "   [3] Прізвище та ініціали викладача (teacher_name)\n"
                "   [4] Повна назва предмету (course_name)\n"
                "   [5] Номер лабораторної роботи (lab_number)\n"
                "   [6] Повний вихідний КОД ПРОГРАМИ студента (program_code)\n"
                "2. ДЖЕРЕЛО Б — 5 РОЗДІЛІВ, ЯКІ ТИ ФОРМУЄШ САМОСТІЙНО (НІКОЛИ не запитуй їх у користувача):\n"
                "   - Тема (topic), Мета (goal), Завдання (tasks) — береш із файлу методички через extract_text_from_file або формулюєш за кодом студента.\n"
                "   - Хід роботи (workflow) та Висновки (conclusions) — описуєш на основі коду студента.\n\n"
                "АЛГОРИТМ ПЕРЕВІРКИ ПЕРЕД КОЖНИМ ДІЄЮ (ВИКОНУЙ ПОКРОКОВО):\n"
                "- Крок 1: Якщо в історії є системне повідомлення про завантажену методичку і ти її ще не читав — спочатку виклич extract_text_from_file, щоб дізнатися тему, мету та завдання.\n"
                "- Крок 2: Перевір усі повідомлення користувача на наявність КОЖНОГО з 6 пунктів Джерела А.\n"
                "- Крок 3 (БЛОКУВАННЯ): Якщо в повідомленнях користувача відсутній хоча б ОДИН пункт із 6 (наприклад, немає ПІБ викладача або немає коду програми):\n"
                "  * КАТЕГОРИЧНО ЗАБОРОНЕНО викликати функцію generate_report!\n"
                "  * КАТЕГОРИЧНО ЗАБОРОНЕНО писати код самостійно за умовою задачі з методички!\n"
                "  * КАТЕГОРИЧНО ЗАБОРОНЕНО підставляти заглушки ('Не вказано', 'Невідомо', 'Студент', 'Викладач' тощо)!\n"
                "  * Напиши користувачу чіткий перелік того, чого саме не вистачає із 6 пунктів, і зупинись.\n"
                "- Крок 4: Викликай generate_report ЛИШЕ тоді, коли всі 6 пунктів Джерела А реально присутні в тексті користувача. У поле program_code передай дослівно код користувача без змін.\n\n"
                "ПРИКЛАДИ ПРАВИЛЬНОЇ ПОВЕДІНКИ:\n"
                "Приклад 1:\n"
                "Користувач: [Завантажив методичку] Я Іваненко І.І., група КН-41, предмет ШІ, лаба 3.\n"
                "Твоя дія: Викликати ТІЛЬКИ extract_text_from_file. Після отримання тексту методички НЕ викликати generate_report, а відповісти: "
                "'Методичку опрацьовано. Для створення звіту надайте, будь ласка: 1) Прізвище та ініціали викладача; 2) Ваш готовий код програми.'\n\n"
                "Приклад 2:\n"
                "Користувач: Ось код: import math... викладач Петренко П.П.\n"
                "Твоя дія: Перевірити, чи є тепер усі 6 пунктів (студент, група, викладач, предмет, номер лаби, код). Якщо є всі 6 — викликати generate_report."
            )
        }
    ]

if "tool_logs" not in st.session_state:
    st.session_state.tool_logs = []

with st.sidebar:
    st.header("Трейсинг інструментів")
    if not st.session_state.tool_logs:
        st.caption("Інструменти ще не викликалися.")
    else:
        for i, log in enumerate(st.session_state.tool_logs, 1):
            with st.expander(f"Крок {i}: {log['tool']}", expanded=True):
                st.markdown("**Аргументи:**")
                st.json(log["args"])
                st.markdown("**Результат:**")
                st.code(log["result"], language="text")

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

if user_message := st.chat_input("Введіть повідомлення..."):
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

                st.session_state.tool_logs.append({
                    "tool": tool_name,
                    "args": {k: (v[:200] + "..." if isinstance(v, str) and len(v) > 200 else v) for k, v in args.items()},
                    "result": result[:300] + "..." if len(str(result)) > 300 else str(result)
                })

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

        if st.session_state.tool_logs:
            st.rerun()

if "generated_report" in st.session_state and os.path.exists(st.session_state.generated_report):
    with open(st.session_state.generated_report, "rb") as file:
        st.download_button(
            label="Завантажити готовий звіт",
            data=file.read(),
            file_name=st.session_state.generated_report,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
