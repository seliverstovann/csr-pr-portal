import json
import os
import re
import uuid
from datetime import datetime

import docx
import pandas as pd
import PyPDF2
import streamlit as st
from bs4 import BeautifulSoup
from openai import OpenAI
from pptx import Presentation

# ============================================================
# НАСТРОЙКИ
# ============================================================
st.set_page_config(page_title="ЦСР PR-портал", page_icon="⚡️", layout="wide", initial_sidebar_state="expanded")

MODEL = "gpt-4o"
TEMP_GENERATE = 0.4
TEMP_REFINE = 0.3
TEMP_FACTCHECK = 0.1

STYLE_LIBRARY_FILE = "style_library.json"   # хранится на диске приложения
MAX_STYLE_EXAMPLES_IN_PROMPT = 5            # сколько примеров максимум уходит в промпт
MAX_CHARS_PER_EXAMPLE = 1200                # обрезаем длинные примеры, чтобы не раздувать промпт
HISTORY_KEEP_LAST = 6                       # сколько последних сообщений диалога отправляем при уточнении

# Поля формы под каждый формат поста
CATEGORY_FIELDS = {
    "1. Анонс мероприятия": [
        {"label": "Дата мероприятия", "key": "date", "area": False},
        {"label": "Название мероприятия", "key": "event_name", "area": False},
        {"label": "Участники / спикеры", "key": "participants", "area": True},
        {"label": "Темы для обсуждения", "key": "topics", "area": True},
        {"label": "Место проведения", "key": "location", "area": False},
    ],
    "2. Исследование": [
        {"label": "Название исследования", "key": "research_name", "area": False},
        {"label": "Источник / автор", "key": "source", "area": False},
        {"label": "Ключевые цифры и выводы", "key": "key_findings", "area": True},
        {"label": "Методология (кратко, необязательно)", "key": "methodology", "area": True},
    ],
    "3. Пост после мероприятия (Итоги)": [
        {"label": "Название мероприятия", "key": "event_name", "area": False},
        {"label": "Дата", "key": "date", "area": False},
        {"label": "Спикеры", "key": "participants", "area": True},
        {"label": "Основные итоги / тезисы", "key": "outcomes", "area": True},
    ],
    "4. Публикация в СМИ (Колонка / комментарий)": [
        {"label": "Название СМИ", "key": "media_name", "area": False},
        {"label": "Автор / спикер", "key": "author", "area": False},
        {"label": "Тема комментария", "key": "topic", "area": True},
        {"label": "Ссылка на публикацию", "key": "link", "area": False},
    ],
    "5. Тематический пост": [
        {"label": "Тема", "key": "topic", "area": False},
        {"label": "Ключевые тезисы", "key": "key_points", "area": True},
        {"label": "Источник (если есть)", "key": "source", "area": False},
    ],
    "6. Индексы": [
        {"label": "Название индекса", "key": "index_name", "area": False},
        {"label": "Значение / показатель", "key": "value", "area": False},
        {"label": "Динамика (рост/падение)", "key": "dynamics", "area": False},
        {"label": "Источник", "key": "source", "area": False},
    ],
    "7. Публикация каталогов": [
        {"label": "Название каталога", "key": "catalog_name", "area": False},
        {"label": "Ссылка / файл", "key": "link", "area": False},
        {"label": "Краткое описание содержания", "key": "description", "area": True},
    ],
    "8. Еженедельная статистика (Коротко о ценах)": [
        {"label": "Период", "key": "period", "area": False},
        {"label": "Категории и цифры", "key": "figures", "area": True},
        {"label": "Источник данных", "key": "source", "area": False},
    ],
    "9. Рубрика: Цифры и факты": [
        {"label": "Список фактов (цифра — пояснение, по одному на строку)", "key": "facts_list", "area": True},
        {"label": "Источники (по каждому факту)", "key": "sources", "area": True},
    ],
}

# ============================================================
# (ОПЦИОНАЛЬНАЯ) ЗАЩИТА ПАРОЛЕМ
# Если в secrets.toml задан APP_PASSWORD — попросим ввести пароль.
# Если секрет не задан, приложение работает без защиты, как раньше.
# ============================================================
try:
    APP_PASSWORD = st.secrets.get("APP_PASSWORD")
except Exception:
    APP_PASSWORD = None

if APP_PASSWORD:
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        st.title("ЦСР PR-портал")
        pwd = st.text_input("Введите пароль доступа", type="password")
        if st.button("Войти", type="primary"):
            if pwd == APP_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Неверный пароль")
        st.stop()

try:
    API_KEY = st.secrets["OPENAI_API_KEY"]
except Exception:
    st.error("Не найден OPENAI_API_KEY в secrets. Добавьте его в настройках приложения.")
    st.stop()

client = OpenAI(api_key=API_KEY)

# ============================================================
# ИНИЦИАЛИЗАЦИЯ SESSION STATE
# ============================================================
defaults = {
    "current_text": "",
    "chat_history": [],
    "raw_source_data": "",
    "versions_history": [],   # список dict: {"text": ..., "label": ...}
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

# ============================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================

def extract_text_from_file(uploaded_file):
    """Извлекает текст из одного файла (pdf, pptx, docx, xlsx, txt, html-экспорт Telegram)."""
    name = uploaded_file.name.lower()
    text = ""
    try:
        if name.endswith(".pdf"):
            reader = PyPDF2.PdfReader(uploaded_file)
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        elif name.endswith(".pptx"):
            prs = Presentation(uploaded_file)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        text += shape.text + "\n"
        elif name.endswith(".docx"):
            doc = docx.Document(uploaded_file)
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif name.endswith(".xlsx"):
            xls = pd.read_excel(uploaded_file, sheet_name=None)
            for sheet_name, df in xls.items():
                text += f"--- Лист: {sheet_name} ---\n"
                text += df.to_string(index=False) + "\n\n"
        elif name.endswith(".txt"):
            text = uploaded_file.read().decode("utf-8", errors="ignore")
        elif name.endswith(".html") or name.endswith(".htm"):
            # экспорт переписки Telegram и подобные HTML-файлы
            messages = extract_telegram_messages(uploaded_file.read())
            text = "\n\n---\n\n".join(messages)
        else:
            st.warning(f"Формат файла «{uploaded_file.name}» не поддерживается и был пропущен.")
    except Exception as e:
        st.warning(f"Не удалось прочитать файл «{uploaded_file.name}»: {e}")
    return text


def extract_text_from_files(uploaded_files):
    """Извлекает и объединяет текст из списка файлов."""
    if not uploaded_files:
        return ""
    chunks = []
    for f in uploaded_files:
        t = extract_text_from_file(f)
        if t.strip():
            chunks.append(f"### Файл: {f.name}\n{t}")
    return "\n\n".join(chunks)


def extract_telegram_messages(html_bytes):
    """Парсит HTML-экспорт Telegram и возвращает список текстов отдельных сообщений."""
    soup = BeautifulSoup(html_bytes, "html.parser")
    messages = []
    for div in soup.find_all("div", class_="text"):
        t = div.get_text(separator="\n").strip()
        t = re.sub(r"\n{3,}", "\n\n", t)
        if len(t) > 30:  # отсекаем служебные короткие блоки (имена, даты и т.п.)
            messages.append(t)
    return messages


def create_docx(text):
    """Создаёт .docx, поддерживая **жирный** markdown-синтаксис построчно."""
    doc = docx.Document()
    for line in text.split("\n"):
        if line.strip() == "":
            doc.add_paragraph("")
            continue
        p = doc.add_paragraph()
        parts = re.split(r"(\*\*.*?\*\*)", line)
        for part in parts:
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                run = p.add_run(part[2:-2])
                run.bold = True
            elif part:
                p.add_run(part)
    bio = __import__("io").BytesIO()
    doc.save(bio)
    return bio.getvalue()


def call_openai(messages, temperature):
    """Обёртка над вызовом OpenAI API с обработкой ошибок."""
    try:
        response = client.chat.completions.create(model=MODEL, messages=messages, temperature=temperature)
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"Ошибка обращения к OpenAI API: {e}")
        return None


def trim_history(history):
    """Ограничивает историю диалога: системный промпт + первый запрос + последние N сообщений."""
    if len(history) <= HISTORY_KEEP_LAST + 2:
        return history
    return history[:2] + history[-HISTORY_KEEP_LAST:]


# --- Библиотека эталонных примеров стиля ---

def load_style_library():
    if os.path.exists(STYLE_LIBRARY_FILE):
        try:
            with open(STYLE_LIBRARY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def save_style_library(library):
    with open(STYLE_LIBRARY_FILE, "w", encoding="utf-8") as f:
        json.dump(library, f, ensure_ascii=False, indent=2)


def add_style_examples(texts):
    library = load_style_library()
    for t in texts:
        t = t.strip()
        if not t:
            continue
        library.append({
            "id": str(uuid.uuid4()),
            "preview": (t[:70] + "…") if len(t) > 70 else t,
            "text": t,
            "added": datetime.now().strftime("%Y-%m-%d %H:%M"),
        })
    save_style_library(library)


def delete_style_example(example_id):
    library = load_style_library()
    library = [e for e in library if e["id"] != example_id]
    save_style_library(library)


def build_style_block(selected_ids):
    if not selected_ids:
        return ""
    library = load_style_library()
    chosen = [e for e in library if e["id"] in selected_ids][:MAX_STYLE_EXAMPLES_IN_PROMPT]
    if not chosen:
        return ""
    parts = []
    for e in chosen:
        snippet = e["text"][:MAX_CHARS_PER_EXAMPLE]
        parts.append(snippet)
    block = "\n\n---\n\n".join(parts)
    return f"""
Ниже приведены реальные примеры постов канала — ориентируйся на них как на эталон стиля,
подачи, длины предложений, использования эмодзи и структуры. Не копируй их содержание,
используй только как образец стиля:

{block}
"""


def build_dynamic_fields(category):
    """Рисует поля формы под выбранную рубрику и возвращает словарь значений."""
    values = {}
    fields = CATEGORY_FIELDS.get(category, [])
    for field in fields:
        if field["area"]:
            values[field["key"]] = st.text_area(field["label"], key=f"field_{field['key']}")
        else:
            values[field["key"]] = st.text_input(field["label"], key=f"field_{field['key']}")
    return values


def format_fields_as_text(category, values):
    fields = CATEGORY_FIELDS.get(category, [])
    lines = [f"Рубрика: {category}"]
    for field in fields:
        val = values.get(field["key"], "")
        if val:
            lines.append(f"{field['label']}: {val}")
    return "\n".join(lines)


# ============================================================
# БОКОВАЯ ПАНЕЛЬ
# ============================================================
with st.sidebar:
    st.title("ЦСР PR-портал")
    st.caption("Внутренний ИИ-ассистент пресс-службы")
    st.divider()

    task = st.selectbox("Выберите задачу:", ["Написать пост для Telegram-канала"])
    post_category = st.selectbox("Рубрика поста:", list(CATEGORY_FIELDS.keys()))

    st.divider()
    st.subheader("Настройки текста")
    text_length = st.select_slider(
        "Объем:", options=["Короткий (до 1000 зн.)", "Стандартный", "Развернутый (лонгрид)"], value="Стандартный"
    )
    tone = st.selectbox(
        "Тональность:",
        ["Строгий (сухие факты, официально)", "Корпоративно-информационный (стандарт)", "Живой (вовлекающий, для соцсетей)"],
    )

    st.divider()
    st.subheader("📚 Эталонные примеры стиля")
    st.caption("Добавьте примеры постов канала — модель будет ориентироваться на них при генерации.")

    with st.expander("Добавить примеры"):
        style_files = st.file_uploader(
            "Файлы с примерами (можно несколько; поддерживается HTML-экспорт Telegram)",
            type=["txt", "pdf", "docx", "html", "htm"],
            accept_multiple_files=True,
            key="style_files_uploader",
        )
        pasted_examples = st.text_area(
            "Или вставьте текст примеров (разделяйте несколько постов строкой ---)",
            key="style_paste_area",
        )
        if st.button("Добавить в библиотеку", use_container_width=True):
            new_texts = []
            for f in style_files or []:
                name = f.name.lower()
                if name.endswith((".html", ".htm")):
                    new_texts.extend(extract_telegram_messages(f.read()))
                else:
                    t = extract_text_from_file(f)
                    if t.strip():
                        new_texts.append(t)
            if pasted_examples.strip():
                new_texts.extend([chunk.strip() for chunk in pasted_examples.split("\n---\n") if chunk.strip()])
            if new_texts:
                add_style_examples(new_texts)
                st.success(f"Добавлено примеров: {len(new_texts)}")
                st.rerun()
            else:
                st.warning("Не найдено текста для добавления.")

    library = load_style_library()
    if library:
        st.caption(f"В библиотеке: {len(library)} примеров. Выберите, какие использовать при генерации:")
        if "selected_style_ids" not in st.session_state:
            # по умолчанию берём последние несколько примеров
            st.session_state.selected_style_ids = [e["id"] for e in library[-MAX_STYLE_EXAMPLES_IN_PROMPT:]]

        with st.expander(f"Управление примерами ({len(library)})"):
            for e in reversed(library):
                col1, col2 = st.columns([4, 1])
                with col1:
                    checked = st.checkbox(
                        e["preview"], value=e["id"] in st.session_state.selected_style_ids, key=f"chk_{e['id']}"
                    )
                    if checked and e["id"] not in st.session_state.selected_style_ids:
                        st.session_state.selected_style_ids.append(e["id"])
                    elif not checked and e["id"] in st.session_state.selected_style_ids:
                        st.session_state.selected_style_ids.remove(e["id"])
                with col2:
                    if st.button("🗑️", key=f"del_{e['id']}"):
                        delete_style_example(e["id"])
                        st.rerun()
    else:
        st.caption("Библиотека пуста — добавьте примеры выше.")
        st.session_state.selected_style_ids = []

    st.divider()
    if st.button("🔄 Начать заново (очистить всё)", use_container_width=True):
        for key in ["current_text", "chat_history", "raw_source_data", "versions_history"]:
            st.session_state[key] = defaults[key]
        st.rerun()

# ============================================================
# ОСНОВНОЙ ЭКРАН
# ============================================================
left_column, right_column = st.columns([1.1, 0.9], gap="large")

with left_column:
    st.subheader("Исходные данные")
    with st.container(border=True):
        st.caption(f"Поля для рубрики «{post_category}»")
        field_values = build_dynamic_fields(post_category)

        speaker = st.text_input("Спикер (если не указан в полях выше)", placeholder="Имя, фамилия и должность")
        materials_text = st.text_area("Дополнительные материалы / тезисы", height=120)
        uploaded_files = st.file_uploader(
            "Загрузите файлы (можно несколько): PDF, PPTX, DOCX, XLSX, TXT",
            type=["pdf", "pptx", "docx", "xlsx", "txt"],
            accept_multiple_files=True,
        )

        if st.button("Сформировать материал", type="primary", use_container_width=True):
            has_field_data = any(v.strip() for v in field_values.values() if v)
            if not has_field_data and not materials_text.strip() and not uploaded_files:
                st.warning("Заполните хотя бы одно поле, добавьте тезисы или загрузите файл.")
            else:
                extracted_file_text = extract_text_from_files(uploaded_files)
                fields_text = format_fields_as_text(post_category, field_values)

                final_materials = "\n".join(filter(None, [
                    fields_text,
                    f"Спикер: {speaker}" if speaker else "",
                    materials_text,
                    extracted_file_text,
                ]))
                st.session_state.raw_source_data = final_materials

                style_block = build_style_block(st.session_state.get("selected_style_ids", []))

                system_rules = f"""
Ты — PR-редактор. Пиши пост для Telegram-канала. Рубрика: {post_category}.
Тон: {tone}. Объем: {text_length}.
Заголовок: первая строка — заголовок, оформи её через **жирный текст** (markdown-нотация).
Подвал: всегда добавляй "⭐️ подписаться на канал".
{style_block}
"""
                st.session_state.chat_history = [
                    {"role": "system", "content": system_rules},
                    {"role": "user", "content": st.session_state.raw_source_data},
                ]

                with st.spinner("Пишу текст..."):
                    result = call_openai(st.session_state.chat_history, TEMP_GENERATE)
                    if result:
                        st.session_state.current_text = result
                        st.session_state.chat_history.append({"role": "assistant", "content": result})
                        st.session_state.versions_history.append({"text": result, "label": "Первая версия"})
                        st.rerun()

with right_column:
    st.subheader("Результат")
    with st.container(border=True):
        text_tab, fact_tab, history_tab = st.tabs(["Готовый текст", "Проверка фактов", "История версий"])

        with text_tab:
            if st.session_state.current_text:
                st.code(st.session_state.current_text, language="markdown")

                docx_data = create_docx(st.session_state.current_text)
                st.download_button(
                    label="📥 Скачать в формате Word (.docx)",
                    data=docx_data,
                    file_name="post.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                st.divider()
                st.write("🔄 **Внести правки:**")
                with st.form("refine_form", clear_on_submit=True):
                    refine_prompt = st.text_input("Что нужно изменить? (например: 'сделай короче', 'убери эмодзи')")
                    submitted = st.form_submit_button("Уточнить текст")
                    if submitted and refine_prompt.strip():
                        st.session_state.chat_history.append({"role": "user", "content": refine_prompt})
                        api_messages = trim_history(st.session_state.chat_history)
                        with st.spinner("Переписываю..."):
                            result = call_openai(api_messages, TEMP_REFINE)
                            if result:
                                st.session_state.current_text = result
                                st.session_state.chat_history.append({"role": "assistant", "content": result})
                                st.session_state.versions_history.append({"text": result, "label": refine_prompt})
                                st.rerun()
                    elif submitted:
                        st.warning("Опишите, что нужно изменить.")
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
3. Жесткая проверка корпоративных стандартов: слово «муниципальные» написано корректно
   (опечатка «нумиципальный» недопустима), а при упоминании медикаментов выведены ТОЛЬКО
   проценты изменения без названий самих препаратов.

ИСХОДНИКИ: {st.session_state.raw_source_data}
ГОТОВЫЙ ТЕКСТ: {st.session_state.current_text}
"""
                    with st.spinner("Сверяю данные..."):
                        result = call_openai([{"role": "user", "content": fact_prompt}], TEMP_FACTCHECK)
                        if result:
                            st.success("Проверка завершена!")
                            st.write(result)
            else:
                st.info("Сначала сгенерируйте текст.")

        with history_tab:
            if not st.session_state.versions_history:
                st.write("Здесь будут сохраняться все версии текстов (от первой генерации до последних правок).")
            else:
                for i, version in enumerate(reversed(st.session_state.versions_history)):
                    version_number = len(st.session_state.versions_history) - i
                    col1, col2 = st.columns([4, 1])
                    with col1:
                        st.markdown(f"**Версия {version_number}** · _{version['label']}_")
                    with col2:
                        if st.button("Восстановить", key=f"restore_{version_number}"):
                            st.session_state.current_text = version["text"]
                            st.rerun()
                    st.info(version["text"])