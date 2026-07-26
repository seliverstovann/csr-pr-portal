import streamlit as st
from openai import OpenAI
import PyPDF2
from pptx import Presentation
import io
import docx
import pandas as pd

# 1. Настройка страницы
st.set_page_config(page_title="ЦСР PR-портал", page_icon="⚡️", layout="wide", initial_sidebar_state="expanded")

API_KEY = st.secrets["OPENAI_API_KEY"]
client = OpenAI(api_key=API_KEY)

# Инициализация памяти (Session State)
if "current_text" not in st.session_state:
    st.session_state.current_text = ""
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "raw_source_data" not in st.session_state:
    st.session_state.raw_source_data = ""

# Вспомогательные функции
def extract_text_from_file(uploaded_file):
    text = ""
    if uploaded_file.name.endswith('.pdf'):
        reader = PyPDF2.PdfReader(uploaded_file)
        for page in reader.pages:
            if page.extract_text(): text += page.extract_text() + "\n"
    elif uploaded_file.name.endswith('.pptx'):
        prs = Presentation(uploaded_file)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"): text += shape.text + "\n"
    elif uploaded_file.name.endswith('.docx'):
        doc = docx.Document(uploaded_file)
        for para in doc.paragraphs: text += para.text + "\n"
    elif uploaded_file.name.endswith('.xlsx'):
        xls = pd.read_excel(uploaded_file, sheet_name=None)
        for sheet_name, df in xls.items():
            text += f"--- Лист: {sheet_name} ---\n"
            text += df.to_string(index=False) + "\n\n"
    return text

def create_docx(text):
    doc = docx.Document()
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# 2. Боковая панель
with st.sidebar:
    st.title("ЦСР PR-портал")
    st.caption("Внутренний ИИ-ассистент пресс-службы")
    st.divider()
    
    task = st.selectbox("Выберите задачу:", ["Написать пост для Telegram-канала"])
    post_category = st.selectbox("Рубрика поста:", [
        "1. Анонс мероприятия", "2. Исследование", "3. Пост после мероприятия (Итоги)",
        "4. Публикация в СМИ (Колонка / комментарий)", "5. Тематический пост",
        "6. Индексы", "7. Публикация каталогов", "8. Еженедельная статистика (Коротко о ценах)",
        "9. Рубрика: Цифры и факты"
    ])
    
    st.divider()
    st.subheader("Настройки текста")
    text_length = st.select_slider("Объем:", options=["Короткий (до 1000 зн.)", "Стандартный", "Развернутый (лонгрид)"], value="Стандартный")
    tone = st.selectbox("Тональность:", ["Строгий (сухие факты, официально)", "Корпоративно-информационный (стандарт)", "Живой (вовлекающий, для соцсетей)"])

# 3. Основной экран
left_column, right_column = st.columns([1.1, 0.9], gap="large")

with left_column:
    st.subheader("Исходные данные")
    with st.container(border=True):
        context = st.text_input("Контекст", placeholder="Мероприятие, исследование или инфоповод")
        speaker = st.text_input("Спикер", placeholder="Имя, фамилия и должность")
        materials_text = st.text_area("Факты, тезисы и исходный текст", height=180)
        uploaded_file = st.file_uploader("Загрузите файл (PDF, PPTX, DOCX, XLSX)", type=["pdf", "pptx", "docx", "xlsx"])
        
        if st.button("Сформировать материал", type="primary", use_container_width=True):
            extracted_file_text = extract_text_from_file(uploaded_file) if uploaded_file else ""
            final_materials = materials_text + "\n" + extracted_file_text
            st.session_state.raw_source_data = f"[КОНТЕКСТ]: {context}\n[СПИКЕР]: {speaker}\n[МАТЕРИАЛЫ]: {final_materials}"
            
            system_rules = f"""
            Ты — PR-редактор. Пиши пост для Telegram-канала. Рубрика: {post_category}.
            Тон: {tone}. Объем: {text_length}.
            Заголовок: Первая строка жирным шрифтом. Подвал: всегда ⭐️ подписаться на канал.
            """
            
            st.session_state.chat_history = [
                {"role": "system", "content": system_rules},
                {"role": "user", "content": st.session_state.raw_source_data}
            ]
            
            with st.spinner('Пишу текст...'):
                response = client.chat.completions.create(model="gpt-4o", messages=st.session_state.chat_history, temperature=0.4)
                st.session_state.current_text = response.choices[0].message.content
                st.session_state.chat_history.append({"role": "assistant", "content": st.session_state.current_text})

with right_column:
    st.subheader("Результат")
    with st.container(border=True):
        text_tab, fact_tab = st.tabs(["Готовый текст", "Проверка фактов и цифр"])
        
        with text_tab:
            if st.session_state.current_text:
                # Используем блок кода для удобного копирования в 1 клик (иконка в правом верхнем углу блока)
                st.code(st.session_state.current_text, language="markdown")
                
                # Кнопка скачивания Word
                docx_data = create_docx(st.session_state.current_text)
                st.download_button(label="📥 Скачать в формате Word (.docx)", data=docx_data, file_name="post.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                
                st.divider()
                st.write("🔄 **Внести правки:**")
                refine_prompt = st.text_input("Что нужно изменить? (например: 'сделай короче', 'убери эмодзи')")
                if st.button("Уточнить текст"):
                    with st.spinner('Переписываю...'):
                        st.session_state.chat_history.append({"role": "user", "content": refine_prompt})
                        response = client.chat.completions.create(model="gpt-4o", messages=st.session_state.chat_history, temperature=0.3)
                        st.session_state.current_text = response.choices[0].message.content
                        st.session_state.chat_history.append({"role": "assistant", "content": st.session_state.current_text})
                        st.rerun() # Перезагружаем интерфейс для отображения нового текста
            else:
                st.info("👈 Заполните данные слева и нажмите «Сформировать материал»")

        with fact_tab:
            if st.session_state.current_text:
                st.write("Здесь ИИ сверяет готовый текст с исходниками на предмет искажений.")
                if st.button("Запустить проверку", type="primary"):
                    fact_prompt = f"""
                    Ты строгий фактчекер. Сравни ИСХОДНИКИ и ГОТОВЫЙ ТЕКСТ.
                    1. Убедись, что все цифры, даты и фамилии перенесены без искажений.
                    2. Найди дублирующиеся блоки статистики и подсвети их как ошибку.
                    3. Жесткая проверка корпоративных стандартов: слово «муниципальные» написано корректно (опечатка «нумиципальный» недопустима), а при упоминании медикаментов выведены ТОЛЬКО проценты изменения без названий самих препаратов.
                    
                    ИСХОДНИКИ: {st.session_state.raw_source_data}
                    ГОТОВЫЙ ТЕКСТ: {st.session_state.current_text}
                    """
                    with st.spinner('Сверяю данные...'):
                        fact_response = client.chat.completions.create(
                            model="gpt-4o", 
                            messages=[{"role": "user", "content": fact_prompt}],
                            temperature=0.1
                        )
                        st.success("Проверка завершена!")
                        st.write(fact_response.choices[0].message.content)
            else:
                st.info("Сначала сгенерируйте текст.")