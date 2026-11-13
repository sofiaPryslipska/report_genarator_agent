import streamlit as st
import requests
import json
import pypdf
import docx
import os
from dotenv import load_dotenv

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
            "description": "Генерує фінальний звіт у форматі DOCX.",
            "parameters": {
                "type": "object",
                "properties": {
                    "course_name": {"type": "string"},
                    "lab_number": {"type": "string"},
                    "student_name": {"type": "string"},
                    "group_name": {"type": "string"},
                    "teacher_name": {"type": "string"},
                    "report_body": {"type": "string"}
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
                if extracted: text += extracted + "\n"
        elif ext == "docx":
            doc = docx.Document(file_path)
            for para in doc.paragraphs: text += para.text + "\n"
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
            "{{TEACHER_NAME}}": teacher_name,
            "{{REPORT_BODY}}": report_body
        }
        for paragraph in doc.paragraphs:
            for tag, actual_text in replacements.items():
                if tag in paragraph.text:
                    paragraph.text = paragraph.text.replace(tag, actual_text)

        output_filename = f"Lab_{lab_number}_{student_name.replace(' ', '_')}.docx"
        doc.save(output_filename)
        return f"Успіх! Звіт збережено як '{output_filename}'."
    except Exception as e:
        return f"Помилка при генерації: {str(e)}"


def send_to_llm(messages):
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    request = {'model': 'google/gemma-4-26b-a4b-it:free', 'messages': messages, 'tools': tools}
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
        {"role": "system",
         "content": "Ти агент-помічник. Твоя ціль — створити звіт. Обов'язково запитай у користувача його ім'я, групу, викладача, предмет, номер лаби та код програми. Тільки після отримання всіх даних використовуй generate_report."}
    ]

# Обробка завантаження файлу
uploaded_file = st.file_uploader("Завантажте методичку (PDF/DOCX)", type=["pdf", "docx"])
if uploaded_file and "file_uploaded" not in st.session_state:
    file_path = uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())  # Фізичне збереження на диск

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

            # Якщо модель викликала інструмент
            if llm_msg.get('tool_calls'):
                for tool in llm_msg['tool_calls']:
                    tool_name = tool['function']['name']
                    tool_id = tool['id']
                    args = json.loads(tool['function']['arguments'])  # Дістаємо аргументи

                    if tool_name == 'extract_text_from_file':
                        result = extract_text_from_file(args.get('file_path'))
                    elif tool_name == 'generate_report':
                        result = generate_report(
                            args.get('course_name'), args.get('lab_number'),
                            args.get('student_name'), args.get('group_name'),
                            args.get('teacher_name'), args.get('report_body')
                        )
                        st.success(result)  # Виводимо зелену плашку про успіх в UI
                        
                        output_filename = f"Lab_{args.get('lab_number')}_{args.get('student_name').replace(' ', '_')}.docx"
                        if os.path.exists(output_filename):
                            with open(output_filename, "rb") as file:
                                st.download_button(
                                    label="📥 Завантажити готовий звіт",
                                    data=file,
                                    file_name=output_filename,
                                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                                )
                    else:
                        result = f'Інструмент {tool_name} не існує'

                    st.session_state.messages.append({
                        "role": "tool",
                        "tool_call_id": tool_id,
                        "name": tool_name,
                        "content": result
                    })

                # Відправляємо результати інструменту назад у модель для фінальної відповіді
                llm_msg = send_to_llm(st.session_state.messages)
                if llm_msg:
                    st.session_state.messages.append(llm_msg)

            # Вивід відповіді агента
            if llm_msg and llm_msg.get('content'):
                with st.chat_message("assistant"):
                    st.markdown(llm_msg['content'])
