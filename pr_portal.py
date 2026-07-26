import streamlit as st
from openai import OpenAI
import PyPDF2
from pptx import Presentation
import io
import docx
import pandas as pd

# 1. Настройка страницы (Широкий формат)
st.set_page_config(
    page_title="ЦСР PR-портал",
    page_icon="⚡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

API_KEY = st.secrets["OPENAI_API_KEY"]

# Функция для извлечения текста
def extract_text_from_file(uploaded_file):
    text = ""
    if uploaded_file.name.endswith('.pdf'):
        reader = PyPDF2.PdfReader(uploaded_file)
        for page in reader.pages:
            if page.extract_text():
                text += page.extract_text() + "\n"
    elif uploaded_file.name.endswith('.pptx'):
        prs = Presentation(uploaded_file)
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text += shape.text + "\n"
    elif uploaded_file.name.endswith('.docx'):
        doc = docx.Document(uploaded_file)
        for para in doc.paragraphs:
            text += para.text + "\n"
    elif uploaded_file.name.endswith('.xlsx'):
        xls = pd.read_excel(uploaded_file, sheet_name=None)
        for sheet_name, df in xls.items():
            text += f"--- Лист: {sheet_name} ---\n"
            text += df.to_string(index=False) + "\n\n"
    return text

# 2. Боковая панель (Настройки)
with st.sidebar:
    st.title("ЦСР PR-портал")
    st.caption("Внутренний ИИ-ассистент пресс-службы")
    st.divider()
    
    task = st.selectbox("Выберите задачу:", [
        "Написать пост для Telegram-канала",
        "Очистить статистику (для инфографики)"
    ])
    
    post_category = None
    if task == "Написать пост для Telegram-канала":
        post_category = st.selectbox("Рубрика поста:", [
            "1. Анонс мероприятия",
            "2. Исследование",
            "3. Пост после мероприятия (Итоги)",
            "4. Публикация в СМИ (Колонка / комментарий)",
            "5. Тематический пост (экономика, климат, транспорт и др.)",
            "6. Индексы (Цен. доступность, Пасха, Оливье, Шашлык и др.)",
            "7. Публикация каталогов",
            "8. Еженедельная статистика (Коротко о ценах)",
            "9. Рубрика: Цифры и факты"
        ])

# 3. Основной экран (Колонки)
left_column, right_column = st.columns([1.1, 0.9], gap="large")

with left_column:
    st.subheader("Исходные данные")
    
    with st.container(border=True):
        if task == "Написать пост для Telegram-канала":
            context = st.text_input("Контекст", placeholder="Мероприятие, исследование или инфоповод")
            speaker = st.text_input("Спикер", placeholder="Имя, фамилия и должность")
            materials_text = st.text_area("Факты, тезисы и исходный текст", height=180, placeholder="Вставьте сырой текст...")
            uploaded_file = st.file_uploader("Загрузите файл (PDF, PPTX, DOCX, XLSX)", type=["pdf", "pptx", "docx", "xlsx"])
        else:
            raw_text_clean = st.text_area("Сырая статистика", height=280, placeholder="Введи текст с ошибками и дублями для очистки...")

        generate = st.button("Сформировать материал", type="primary", use_container_width=True)

with right_column:
    st.subheader("Результат")
    
    with st.container(border=True):
        text_tab, info_tab = st.tabs(["Готовый текст", "Служебная информация"])
        
        with info_tab:
            st.write("Здесь будут отображаться технические статусы и прочитанные данные.")
            
        with text_tab:
            if not generate:
                st.info("👈 Заполните данные слева и нажмите «Сформировать материал»")
            else:
                # === ЛОГИКА ГЕНЕРАЦИИ ===
                client = OpenAI(api_key=API_KEY)
                
                if task == "Написать пост для Telegram-канала":
                    extracted_file_text = ""
                    if uploaded_file is not None:
                        extracted_file_text = extract_text_from_file(uploaded_file)
                        st.success(f"Файл прочитан!")
                    
                    final_materials = materials_text + "\n" + extracted_file_text
                    raw_text = f"[КОНТЕКСТ]: {context}\n[СПИКЕР]: {speaker}\n[МАТЕРИАЛЫ]: {final_materials}"
                    
                    # Правила ИИ
                    base_rules = """
                    Ты — строгий редактор пресс-службы Центра стратегических разработок (ЦСР). 
                    Пиши пост для Telegram-канала. Тон: объективный, аналитический, в третьем лице.
                    Заголовок: Первая строка — четкий заголовок, ОБЯЗАТЕЛЬНО выделенный жирным шрифтом.
                    Оформление цитат: Используй знак > в начале каждого абзаца цитаты. 
                    Концовка: Добавь тематические хештеги (всегда начинай с #ЦСР #новости).
                    Подвал: Последней строкой АБСОЛЮТНО ВСЕГДА должна стоять фраза: ⭐️ подписаться на канал
                    """
                    
                    category_rules = ""
                    if post_category == "8. Еженедельная статистика (Коротко о ценах)":
                        category_rules = """
                        СПЕЦИФИКА РУБРИКИ: Заголовок всегда "**Коротко о том, что происходило с ценами в предыдущую неделю и заслуживает внимания**".
                        ЖЕСТКИЙ ФИЛЬТР ДАННЫХ: ПРОИГНОРИРУЙ всё лишнее и вытащи строго:
                        1. Продовольственные товары (эмодзи 📈).
                        2. Для автолюбителей (эмодзи ⛽️).
                        3. Для путешественников (эмодзи 🧳).
                        4. Медикаменты / Лекарства — ПРАВИЛО: выводи только конкретный процент изменения (например, +1.3%), строго без указания названий (не пиши "активированный уголь").
                        В конце отдельной строкой укажи инфляцию.
                        """
                    elif post_category == "9. Рубрика: Цифры и факты":
                        category_rules = """
                        СПЕЦИФИКА РУБРИКИ: Заголовок всегда "**Цифры и факты**". 
                        Используй цифры-эмодзи или значок ‼️ для каждого пункта.
                        """
                    elif post_category == "6. Индексы (Цен. доступность, Пасха, Оливье, Шашлык и др.)":
                        category_rules = """
                        СПЕЦИФИКА РУБРИКИ: 
                        Если это «Индекс ценовой доступности...»: выдели индекс жирным, блок ‼️**Основные факторы:** маркированным списком 🔘.
                        Если это сезонные гастрономические индексы: выдели общую стоимость корзины, выжимка продуктов (📈 подорожали, 📉 подешевели).
                        """
                    elif post_category == "1. Анонс мероприятия":
                        category_rules = "СПЕЦИФИКА РУБРИКИ: Текст краткий, дата/время/место, эмодзи 📢, 🗓️, ⚡️."
                    elif post_category == "3. Пост после мероприятия (Итоги)":
                        category_rules = "СПЕЦИФИКА РУБРИКИ: Ключевые итоги, соглашения. Развернутая цитата спикера."
                    
                    system_rules = base_rules + "\n" + category_rules
                    prompt_text = raw_text
                    
                else:
                    system_rules = """
                    Ты строгий PR-редактор. Очисти сырой текст для корпоративной инфографики.
                    Всегда используй правильное написание «муниципальные» (никогда не пиши «нумиципальный») и следи за правильным написанием слова «влияние».
                    Удаляй дублирующиеся блоки статистики. Выдавай только структурированные цифры (например, +1.3%).
                    """
                    prompt_text = raw_text_clean

                # Запрос к нейросети
                with st.spinner('Готовлю материал...'):
                    try:
                        response = client.chat.completions.create(
                            model="gpt-4o-mini",
                            messages=[
                                {"role": "system", "content": system_rules},
                                {"role": "user", "content": prompt_text}
                            ],
                            temperature=0.3
                        )
                        st.write(response.choices[0].message.content)
                    except Exception as e:
                        st.error(f"Произошла ошибка: {e}")