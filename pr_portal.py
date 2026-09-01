"""
ЦСР PR-ПОРТАЛ 3.0
Премиальный AI-ассистент для пресс-службы ЦСР
- Генерация постов с GPT-4o
- Мониторинг новостей (News API + Google RSS)
- Интеграция с Яндекс.Диском
- Аналитика и метрики
- История постов
"""

import json
import os
import re
import uuid
from datetime import datetime, timedelta
import requests
import feedparser
import urllib.parse

import docx
import pandas as pd
import PyPDF2
import streamlit as st
from bs4 import BeautifulSoup
from openai import OpenAI
from pptx import Presentation

# ============================================================
# ФИРМЕННАЯ ТЕМА ЦСР
# ============================================================

CSR_THEME_CSS = """
<style>
:root {
    --csr-bg: #F6F5FB;
    --csr-bg-soft: #FBFAFE;
    --csr-surface: #FFFFFF;
    --csr-surface-2: #F1F0F8;
    --csr-ink: #1D1D2B;
    --csr-ink-soft: #56556B;
    --csr-muted: #8B899D;
    --csr-primary: #3B376E;
    --csr-primary-2: #5E5998;
    --csr-lilac: #A4A3CE;
    --csr-border: rgba(59, 55, 110, 0.10);
    --csr-border-strong: rgba(59, 55, 110, 0.18);
    --csr-shadow: 0 14px 40px rgba(35, 33, 66, 0.07);
    --csr-shadow-soft: 0 6px 20px rgba(35, 33, 66, 0.05);
}

* { box-sizing: border-box; }

html, body, [data-testid="stAppViewContainer"], .stApp {
    color: var(--csr-ink);
    background:
        radial-gradient(circle at 86% 0%, rgba(164,163,206,0.16), transparent 26%),
        linear-gradient(180deg, #FCFBFE 0%, var(--csr-bg) 100%) !important;
}

[data-testid="stHeader"] {
    background: rgba(252,251,254,0.84) !important;
    backdrop-filter: blur(14px);
    border-bottom: 1px solid rgba(59,55,110,0.05);
}

.block-container {
    max-width: 1450px;
    padding-top: 2.15rem !important;
    padding-bottom: 3rem !important;
    padding-left: 2.3rem !important;
    padding-right: 2.3rem !important;
}

h1, h2, h3, h4, h5, h6 {
    color: var(--csr-ink) !important;
    font-weight: 750 !important;
    letter-spacing: -0.025em !important;
}
h1 { font-size: 2.2rem !important; }
h2 { font-size: 1.38rem !important; }
h3 { font-size: 1.08rem !important; }

/* Sidebar */
[data-testid="stSidebar"] {
    background: rgba(255,255,255,0.92) !important;
    border-right: 1px solid var(--csr-border) !important;
    box-shadow: 10px 0 32px rgba(35,33,66,0.025);
}
[data-testid="stSidebar"] .block-container {
    padding: 1.35rem 1.15rem 1.5rem !important;
}
[data-testid="stSidebar"] [data-testid="stImage"] {
    margin-bottom: 0.6rem;
}
.sidebar-product-name {
    font-size: 1.07rem;
    font-weight: 800;
    letter-spacing: -0.02em;
    color: var(--csr-ink);
    margin-top: 0.2rem;
}
.sidebar-product-caption {
    color: var(--csr-muted);
    font-size: 0.82rem;
    line-height: 1.45;
    margin-top: 0.25rem;
    margin-bottom: 0.7rem;
}
.sidebar-status {
    display: inline-flex;
    align-items: center;
    gap: .45rem;
    font-size: .76rem;
    font-weight: 650;
    color: #5D5A75;
    background: #F5F4F9;
    border: 1px solid var(--csr-border);
    border-radius: 999px;
    padding: .42rem .66rem;
    margin: .2rem 0 .4rem;
}
.status-dot { width: 7px; height: 7px; border-radius: 50%; background: #5E9B74; display: inline-block; }

/* Hero */
.csr-hero {
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, rgba(255,255,255,0.98), rgba(245,244,251,0.98));
    border: 1px solid var(--csr-border);
    border-radius: 28px;
    padding: 30px 34px 28px;
    box-shadow: var(--csr-shadow);
    margin-bottom: 1.5rem;
}
.csr-hero::after {
    content: "";
    position: absolute;
    width: 250px;
    height: 250px;
    right: -90px;
    top: -120px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(94,89,152,0.18), rgba(94,89,152,0));
}
.csr-kicker {
    display: inline-block;
    color: var(--csr-primary);
    background: rgba(94,89,152,0.08);
    border: 1px solid rgba(94,89,152,0.10);
    border-radius: 999px;
    padding: 6px 10px;
    font-size: .74rem;
    font-weight: 750;
    letter-spacing: .04em;
    text-transform: uppercase;
    margin-bottom: 12px;
}
.csr-hero-title {
    color: var(--csr-ink);
    font-size: 2.05rem;
    line-height: 1.08;
    font-weight: 820;
    letter-spacing: -0.04em;
    margin: 0 0 8px;
}
.csr-hero-subtitle {
    color: var(--csr-ink-soft);
    max-width: 800px;
    font-size: .98rem;
    line-height: 1.6;
    margin: 0;
}

/* Cards / containers */
[data-testid="stVerticalBlockBorderWrapper"] {
    background: rgba(255,255,255,0.94) !important;
    border: 1px solid var(--csr-border) !important;
    border-radius: 22px !important;
    box-shadow: var(--csr-shadow-soft) !important;
}

.metric-card {
    background: linear-gradient(180deg, #FFFFFF 0%, #F9F8FC 100%);
    border: 1px solid var(--csr-border);
    border-radius: 22px;
    padding: 21px 20px 19px;
    text-align: left;
    box-shadow: var(--csr-shadow-soft);
    min-height: 126px;
    transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease;
}
.metric-card:hover {
    transform: translateY(-2px);
    border-color: var(--csr-border-strong);
    box-shadow: var(--csr-shadow);
}
.metric-label {
    font-size: .72rem;
    font-weight: 750;
    color: var(--csr-muted);
    text-transform: uppercase;
    letter-spacing: .085em;
    margin-bottom: 15px;
}
.metric-value {
    font-size: 2.5rem;
    font-weight: 820;
    color: var(--csr-primary);
    line-height: 1;
    letter-spacing: -0.04em;
}

/* Inputs */
input, textarea,
[data-baseweb="input"], [data-baseweb="textarea"], [data-baseweb="select"] > div {
    background: #FFFFFF !important;
    color: var(--csr-ink) !important;
    border-color: var(--csr-border) !important;
    border-radius: 13px !important;
}
[data-baseweb="select"] > div:hover,
input:hover, textarea:hover { border-color: var(--csr-border-strong) !important; }
input:focus, textarea:focus {
    border-color: rgba(59,55,110,.35) !important;
    box-shadow: 0 0 0 3px rgba(94,89,152,.08) !important;
}
input::placeholder, textarea::placeholder { color: #A6A4B2 !important; }

/* Buttons */
.stButton > button, .stDownloadButton > button {
    min-height: 42px !important;
    border-radius: 13px !important;
    font-weight: 680 !important;
    border: 1px solid var(--csr-border) !important;
    background: #FFFFFF !important;
    color: var(--csr-primary) !important;
    box-shadow: none !important;
    transition: all .18s ease !important;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    transform: translateY(-1px);
    border-color: var(--csr-border-strong) !important;
    box-shadow: var(--csr-shadow-soft) !important;
}
button[kind="primary"] {
    background: linear-gradient(135deg, #3B376E 0%, #5E5998 100%) !important;
    color: #FFFFFF !important;
    border: none !important;
    box-shadow: 0 10px 24px rgba(59,55,110,.20) !important;
}
button[kind="primary"]:hover { box-shadow: 0 13px 30px rgba(59,55,110,.27) !important; }

/* Tabs */
[role="tablist"] {
    gap: 7px;
    border-bottom: 1px solid var(--csr-border);
}
[role="tab"] {
    color: var(--csr-muted) !important;
    font-weight: 650 !important;
    border-radius: 10px 10px 0 0 !important;
    padding-left: 12px !important;
    padding-right: 12px !important;
}
[role="tab"][aria-selected="true"] {
    color: var(--csr-primary) !important;
    background: rgba(94,89,152,.055) !important;
    border-bottom-color: var(--csr-primary) !important;
}

/* Supporting components */
[data-testid="stAlert"],
[data-testid="stExpander"],
[data-testid="stFileUploaderDropzone"],
[data-testid="stCodeBlock"], pre {
    background: rgba(255,255,255,.92) !important;
    border: 1px solid var(--csr-border) !important;
    border-radius: 16px !important;
    box-shadow: none !important;
}
[data-testid="stExpander"] summary { color: var(--csr-ink) !important; font-weight: 680 !important; }
[data-testid="stFileUploaderDropzone"] { border-style: dashed !important; }
[data-testid="stCodeBlock"] code, pre code, code { color: #353449 !important; background: transparent !important; }

[data-testid="stMarkdownContainer"] p { color: var(--csr-ink-soft) !important; line-height: 1.62 !important; }
[data-testid="stCaptionContainer"], .stCaption, small { color: var(--csr-muted) !important; }
label { color: #45435B !important; font-weight: 590 !important; }
a { color: var(--csr-primary) !important; font-weight: 650; text-decoration: none !important; }
a:hover { text-decoration: underline !important; }
hr {
    border: none !important;
    height: 1px !important;
    background: linear-gradient(90deg, transparent, rgba(59,55,110,.11), transparent) !important;
    margin: 1.15rem 0 !important;
}
footer, #MainMenu { visibility: hidden; }
</style>
"""

# ============================================================
# КОНФИГ
# ============================================================
st.set_page_config(
    page_title="ЦСР PR-портал 3.0", 
    page_icon="⚡️", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

st.markdown(CSR_THEME_CSS, unsafe_allow_html=True)

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "csr_logo.png")

MODEL = "gpt-4o"
TEMP_GENERATE = 0.4
TEMP_REFINE = 0.3
TEMP_FACTCHECK = 0.1

STYLE_LIBRARY_FILE = "style_library.json"
MAX_STYLE_EXAMPLES_IN_PROMPT = 5
MAX_CHARS_PER_EXAMPLE = 1200
HISTORY_KEEP_LAST = 6

# Путь к папке базы знаний на Яндекс.Диске.
# Если папка называется иначе — поменяйте здесь одну строку.
# Точное имя можно посмотреть в разделе «📁» (кнопка «Показать корень диска»).
YANDEX_BASE_PATH = "/ЦСР - база знаний"

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
# СЕКРЕТЫ И КЛЮЧИ
# ============================================================
try:
    APP_PASSWORD = st.secrets.get("APP_PASSWORD")
except:
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
                st.error("❌ Неверный пароль")
        st.stop()

try:
    API_KEY = st.secrets["OPENAI_API_KEY"]
except:
    st.error("⚠️ Не найден OPENAI_API_KEY в secrets.")
    st.stop()

try:
    YANDEX_TOKEN = st.secrets.get("YANDEX_DISK_TOKEN")
except:
    YANDEX_TOKEN = None

try:
    NEWS_API_KEY = st.secrets.get("NEWS_API_KEY")
except:
    NEWS_API_KEY = None

client = OpenAI(api_key=API_KEY)

# ============================================================
# ЯНДЕКС.ДИСК API
# ============================================================
class YandexDiskAPI:
    def __init__(self, token):
        self.token = token
        self.headers = {"Authorization": f"OAuth {token}"}
        self.base_url = "https://cloud-api.yandex.net/v1/disk"
    
    def list_files(self, path="/"):
        url = f"{self.base_url}/resources"
        params = {"path": path, "limit": 100}
        try:
            response = requests.get(url, headers=self.headers, params=params, timeout=10)
            if response.status_code == 200:
                items = response.json().get("_embedded", {}).get("items", [])
                return sorted(items, key=lambda x: x.get("name", ""))
            return []
        except:
            return []
    
    def get_download_url(self, path):
        url = f"{self.base_url}/resources/download"
        params = {"path": path}
        try:
            response = requests.get(url, headers=self.headers, params=params, timeout=10)
            if response.status_code == 200:
                return response.json().get("href")
        except:
            pass
        return None
    
    def download_file(self, path):
        download_url = self.get_download_url(path)
        if download_url:
            try:
                response = requests.get(download_url, timeout=10)
                if response.status_code == 200:
                    return response.content
            except:
                pass
        return None

disk_api = None
if YANDEX_TOKEN:
    disk_api = YandexDiskAPI(YANDEX_TOKEN)


# ============================================================
# ПОСТОЯННОЕ ХРАНИЛИЩЕ (Яндекс.Диск)
# ------------------------------------------------------------
# Файловая система Streamlit Cloud временная: при перезапуске
# приложения всё, что лежит рядом с кодом, стирается. Поэтому
# история постов, избранные новости и библиотека стилей хранятся
# в JSON-файлах на Яндекс.Диске, а локально держится только кэш.
# ============================================================

# Папка для служебных данных приложения. Создаётся автоматически.
APP_DATA_PATH = "/PR-портал ЦСР — данные"


class YandexStorage:
    def __init__(self, token, folder=APP_DATA_PATH):
        self.token = token
        self.folder = folder
        self.headers = {"Authorization": f"OAuth {token}"}
        self.base_url = "https://cloud-api.yandex.net/v1/disk"
        self.writable = None      # None = ещё не проверяли
        self.last_error = ""

    def _ensure_folder(self):
        """Создаёт папку приложения. 409 = уже существует, это норма."""
        try:
            r = requests.put(
                f"{self.base_url}/resources",
                headers=self.headers,
                params={"path": self.folder},
                timeout=10,
            )
            return r.status_code in (201, 409)
        except Exception as e:
            self.last_error = str(e)
            return False

    def load(self, name, default):
        """Читает JSON с Диска. При любой проблеме возвращает default."""
        if not self.token:
            return default
        try:
            r = requests.get(
                f"{self.base_url}/resources/download",
                headers=self.headers,
                params={"path": f"{self.folder}/{name}"},
                timeout=10,
            )
            if r.status_code != 200:
                return default          # файла ещё нет — первый запуск
            href = r.json().get("href")
            data = requests.get(href, timeout=15)
            if data.status_code != 200:
                return default
            return json.loads(data.content.decode("utf-8"))
        except Exception as e:
            self.last_error = str(e)
            return default

    def save(self, name, data):
        """Записывает JSON на Диск. Возвращает True при успехе."""
        if not self.token:
            self.writable = False
            return False
        try:
            self._ensure_folder()
            r = requests.get(
                f"{self.base_url}/resources/upload",
                headers=self.headers,
                params={"path": f"{self.folder}/{name}", "overwrite": "true"},
                timeout=10,
            )
            if r.status_code != 200:
                self.writable = False
                self.last_error = (
                    f"{r.status_code}: {r.json().get('message', '')}"
                )
                return False
            href = r.json().get("href")
            payload = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
            up = requests.put(href, data=payload, timeout=30)
            ok = up.status_code in (201, 202)
            self.writable = ok
            return ok
        except Exception as e:
            self.writable = False
            self.last_error = str(e)
            return False


storage = YandexStorage(YANDEX_TOKEN) if YANDEX_TOKEN else None


def storage_load(name, default):
    """Читает данные: с Диска, если он подключён, иначе из локального файла."""
    if storage:
        return storage.load(name, default)
    if os.path.exists(name):
        try:
            with open(name, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return default
    return default


def storage_save(name, data):
    """Пишет данные на Диск и одновременно в локальный файл (быстрый кэш)."""
    try:
        with open(name, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
    if storage:
        return storage.save(name, data)
    return False

# ============================================================
# МОНИТОРИНГ НОВОСТЕЙ
# ============================================================
class NewsMonitor:
    def __init__(self, news_api_key=None):
        self.news_api_key = news_api_key
        self.news_api_url = "https://newsapi.org/v2/everything"
        self.init_db()
    
    FAVORITES_FILE = "news_favorites.json"

    def init_db(self):
        pass  # хранилище файловое, инициализация не нужна

    def save_favorite(self, title, source, url, published, description):
        items = storage_load(self.FAVORITES_FILE, [])
        if not isinstance(items, list):
            items = []
        if any(i.get("title") == title for i in items):
            return
        items.insert(0, {
            "title": title,
            "source": source,
            "url": url,
            "published": published,
            "description": description,
            "saved_at": datetime.now().isoformat(timespec="seconds"),
        })
        storage_save(self.FAVORITES_FILE, items)

    def get_favorites(self):
        items = storage_load(self.FAVORITES_FILE, [])
        if not isinstance(items, list):
            return []
        return [
            (i.get("title", ""), i.get("source", ""), i.get("url", ""),
             i.get("published", ""), i.get("description", ""))
            for i in items
        ]

    def delete_favorite(self, title):
        items = storage_load(self.FAVORITES_FILE, [])
        if not isinstance(items, list):
            return
        items = [i for i in items if i.get("title") != title]
        storage_save(self.FAVORITES_FILE, items)

    def fetch_news_api(self, query="ЦСР", language="ru", days=7):
        if not self.news_api_key:
            return []
        try:
            params = {
                "q": query,
                "language": language,
                "sortBy": "publishedAt",
                "apiKey": self.news_api_key,
                "pageSize": 25,
                "from": (datetime.now() - timedelta(days=days)).strftime("%Y-%m-%d"),
            }
            response = requests.get(self.news_api_url, params=params, timeout=10)
            if response.status_code != 200:
                st.warning(
                    f"News API вернул {response.status_code}: "
                    f"{response.json().get('message', '')}"
                )
                return []
            if response.status_code == 200:
                data = response.json()
                news_list = []
                for article in data.get('articles', []):
                    try:
                        news_list.append({
                            'title': article.get('title', 'N/A'),
                            'source': article.get('source', {}).get('name', 'News API'),
                            'url': article.get('url', ''),
                            'published': article.get('publishedAt', ''),
                            'description': article.get('description', ''),
                            'type': 'news_api'
                        })
                    except:
                        continue
                return news_list
            return []
        except:
            return []
    
    def fetch_google_news(self, query="ЦСР OR \"Центр стратегических разработок\"", days=7):
        try:
            # ВАЖНО: запрос обязательно кодируем, иначе пробелы ломают URL
            q = f"{query} when:{days}d"
            url = (
                "https://news.google.com/rss/search?q="
                + urllib.parse.quote(q)
                + "&hl=ru&gl=RU&ceid=RU:ru"
            )
            feed = feedparser.parse(url)
            news_list = []
            for entry in feed.entries[:40]:
                try:
                    title = entry.get('title', 'N/A')
                    link = entry.get('link', '')
                    published = entry.get('published', '')
                    summary = entry.get('summary', '')
                    if summary:
                        summary = BeautifulSoup(summary, 'html.parser').get_text()[:250]
                    source = entry.get('source', {}).get('title') if entry.get('source') else None
                    if not source and ' - ' in title:
                        title, source = title.rsplit(' - ', 1)
                    news_list.append({
                        'title': title.strip(),
                        'source': (source or 'Google News').strip(),
                        'url': link,
                        'published': published,
                        'description': summary,
                        'type': 'google_news'
                    })
                except Exception:
                    continue
            return news_list
        except Exception as e:
            st.warning(f"Google News недоступен: {e}")
            return []
    
    def fetch_combined(self, query="ЦСР", use_news_api=True, use_google_news=True, days=7):
        all_news = []
        if use_news_api and self.news_api_key:
            api_news = self.fetch_news_api(query, days=days)
            all_news.extend(api_news)
        if use_google_news:
            google_news = self.fetch_google_news(query, days=days)
            all_news.extend(google_news)
        seen_titles = set()
        unique_news = []
        for news in all_news:
            title = news['title']
            if title not in seen_titles:
                seen_titles.add(title)
                unique_news.append(news)
        return sorted(unique_news, key=lambda x: x.get('published', ''), reverse=True)

news_monitor = NewsMonitor(news_api_key=NEWS_API_KEY)


# Кэш на 10 минут: без него новости перезапрашиваются при КАЖДОМ клике
# в интерфейсе и дневная квота News API (100 запросов) сгорает за пару минут.
@st.cache_data(ttl=600, show_spinner=False)
def cached_fetch_news(query, use_api, use_google, days, _cache_buster=0):
    return news_monitor.fetch_combined(
        query, use_news_api=use_api, use_google_news=use_google, days=days
    )


@st.cache_data(ttl=300, show_spinner=False)
def cached_list_disk(path, _cache_buster=0):
    if not disk_api:
        return []
    return disk_api.list_files(path)

# ============================================================
# ИСТОРИЯ ПОСТОВ (хранится на Яндекс.Диске)
# ============================================================
POSTS_FILE = "posts_history.json"


@st.cache_data(ttl=60, show_spinner=False)
def _load_posts(_buster=0):
    data = storage_load(POSTS_FILE, [])
    return data if isinstance(data, list) else []


def _bump_posts_cache():
    st.session_state.posts_buster = st.session_state.get("posts_buster", 0) + 1


def save_post(category, tone, length, text, raw_source, has_citations=0):
    """Дописывает пост в общий журнал. Перед записью перечитывает файл,
    чтобы не затереть посты, добавленные коллегой с другого компьютера."""
    posts = storage_load(POSTS_FILE, [])
    if not isinstance(posts, list):
        posts = []
    post_id = str(uuid.uuid4())
    posts.append({
        "id": post_id,
        "category": category,
        "tone": tone,
        "length": length,
        "text": text,
        "raw_source": raw_source[:4000],
        "has_citations": int(has_citations),
        "created_at": datetime.now().isoformat(timespec="seconds"),
    })
    storage_save(POSTS_FILE, posts)
    _bump_posts_cache()
    return post_id


def get_posts_history(limit=100):
    posts = _load_posts(st.session_state.get("posts_buster", 0))
    posts = sorted(posts, key=lambda p: p.get("created_at", ""), reverse=True)
    return [
        (p.get("id"), p.get("category"), p.get("tone"),
         p.get("text", ""), p.get("created_at", ""), p.get("has_citations", 0))
        for p in posts[:limit]
    ]


def get_metrics():
    posts = _load_posts(st.session_state.get("posts_buster", 0))
    now = datetime.now()

    def parse(p):
        try:
            return datetime.fromisoformat(p.get("created_at", ""))
        except Exception:
            return None

    week_ago = now - timedelta(days=7)
    month_ago = now - timedelta(days=30)
    dates = [d for d in (parse(p) for p in posts) if d]

    def group(field):
        counts = {}
        for p in posts:
            key = p.get(field) or "—"
            counts[key] = counts.get(key, 0) + 1
        return sorted(counts.items(), key=lambda x: x[1], reverse=True)

    with_citations = sum(1 for p in posts if p.get("has_citations"))

    return {
        "total_posts": len(posts),
        "posts_week": sum(1 for d in dates if d > week_ago),
        "posts_month": sum(1 for d in dates if d > month_ago),
        "by_category": group("category"),
        "by_tone": group("tone"),
        "by_length": group("length"),
        "with_citations": with_citations,
        "without_citations": len(posts) - with_citations,
    }

# ============================================================
# ФУНКЦИИ
# ============================================================

def extract_text_from_file(uploaded_file):
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
    except:
        pass
    return text

def extract_text_from_files(uploaded_files):
    if not uploaded_files:
        return ""
    chunks = []
    for f in uploaded_files:
        t = extract_text_from_file(f)
        if t.strip():
            chunks.append(f"### {f.name}\n{t}")
    return "\n\n".join(chunks)

def extract_yandex_file(disk_api, file_path):
    content = disk_api.download_file(file_path)
    if not content:
        return ""
    name = file_path.lower()
    text = ""
    try:
        if name.endswith(".pdf"):
            from io import BytesIO
            reader = PyPDF2.PdfReader(BytesIO(content))
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        elif name.endswith(".docx"):
            from io import BytesIO
            doc = docx.Document(BytesIO(content))
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif name.endswith(".txt"):
            text = content.decode("utf-8", errors="ignore")
        elif name.endswith(".xlsx"):
            from io import BytesIO
            xls = pd.read_excel(BytesIO(content), sheet_name=None)
            for sheet_name, df in xls.items():
                text += f"--- Лист: {sheet_name} ---\n"
                text += df.to_string(index=False) + "\n\n"
        elif name.endswith(".pptx"):
            from io import BytesIO
            prs = Presentation(BytesIO(content))
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        text += shape.text + "\n"
    except Exception as e:
        st.warning(f"Не удалось прочитать {file_path}: {e}")
    return text

def create_docx(text):
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
    try:
        response = client.chat.completions.create(
            model=MODEL, 
            messages=messages, 
            temperature=temperature
        )
        return response.choices[0].message.content
    except Exception as e:
        st.error(f"❌ Ошибка OpenAI: {e}")
        return None

def trim_history(history):
    if len(history) <= HISTORY_KEEP_LAST + 2:
        return history
    return history[:2] + history[-HISTORY_KEEP_LAST:]

def load_style_library():
    lib = storage_load(STYLE_LIBRARY_FILE, [])
    return lib if isinstance(lib, list) else []


def save_style_library(library):
    storage_save(STYLE_LIBRARY_FILE, library)


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
    parts = [e["text"][:MAX_CHARS_PER_EXAMPLE] for e in chosen]
    block = "\n\n---\n\n".join(parts)
    return f"Примеры постов для ориентира:\n\n{block}"

def get_enhanced_system_rules(category, tone, text_length, use_citations):
    citation_rule = "Пересказывай информацию своими словами без цитат" if not use_citations else "Используй цитаты из источников (максимум 2-3 на пост)"
    
    return f"""Ты — PR-редактор ЦСР. Основные правила:

🎯 СТИЛЬ:
• Аналитически и структурированно
• Заголовок первой строкой **жирным**
• Логические переходы между абзацами
• Четкая логика и причинно-следственные связи

📝 КОНТЕНТ:
• БЕЗ буллитов — только прозаический текст
• {citation_rule}
• Конкретные цифры с контекстом
• Завершающий вывод
• Язык живой, но серьёзный

🚫 ИЗБЕГАЙ:
• "и т.д.", "и т.п."
• Клише типа "как известно"
• Излишних эмодзи
• Коротких рубленых предложений

✅ ДЕЛАЙ:
• Связки: "Это означает...", "Поэтому..."
• Анализ, а не перечисление
• Подвал: "⭐️ подписаться на канал"

Рубрика: {category}
Тон: {tone}
Объем: {text_length}
"""

def build_dynamic_fields(category):
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

def render_page_hero(title, subtitle, kicker="ЦСР · PR WORKSPACE"):
    st.markdown(
        f"""
        <div class="csr-hero">
            <div class="csr-kicker">{kicker}</div>
            <div class="csr-hero-title">{title}</div>
            <p class="csr-hero-subtitle">{subtitle}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# SESSION STATE
# ============================================================
defaults = {
    "current_text": "",
    "chat_history": [],
    "raw_source_data": "",
    "versions_history": [],
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

if "show_metrics" not in st.session_state:
    st.session_state.show_metrics = False
if "show_yandex" not in st.session_state:
    st.session_state.show_yandex = False
if "show_news" not in st.session_state:
    st.session_state.show_news = False
if "news_for_post" not in st.session_state:
    st.session_state.news_for_post = None

# ============================================================
# БОКОВАЯ ПАНЕЛЬ
# ============================================================
with st.sidebar:
    if os.path.exists(LOGO_PATH):
        st.image(LOGO_PATH, width=168)
    else:
        st.markdown("### ЦСР")

    st.markdown(
        """
        <div class="sidebar-product-name">PR-портал</div>
        <div class="sidebar-product-caption">Рабочее пространство пресс-службы</div>
        """,
        unsafe_allow_html=True,
    )

    if storage is None:
        st.warning("Хранилище не подключено")
    elif storage.writable is False:
        st.warning("Яндекс.Диск подключён без записи")
    else:
        st.markdown(
            '<div class="sidebar-status"><span class="status-dot"></span>Данные сохраняются на Яндекс.Диск</div>',
            unsafe_allow_html=True,
        )

    st.divider()

    task = st.selectbox("Задача", ["Написать пост для Telegram-канала"])
    post_category = st.selectbox("Рубрика", list(CATEGORY_FIELDS.keys()))

    with st.expander("Настройки текста", expanded=True):
        text_length = st.select_slider(
            "Объем",
            options=["Короткий (до 1000 зн.)", "Стандартный", "Развернутый (лонгрид)"],
            value="Стандартный",
        )
        tone = st.selectbox(
            "Тональность",
            ["Строгий (сухие факты)", "Стандарт (информационный)", "Живой (для соцсетей)"],
        )
        use_citations = st.checkbox("Использовать цитаты", value=False)

    with st.expander("Примеры стиля"):
        style_files = st.file_uploader(
            "Добавить файлы",
            type=["txt", "pdf", "docx", "html", "htm"],
            accept_multiple_files=True,
            key="style_files_uploader",
        )
        pasted_examples = st.text_area(
            "Или вставьте текст",
            key="style_paste_area",
            height=80,
        )
        if st.button("Добавить примеры", use_container_width=True):
            new_texts = []
            for f in style_files or []:
                t = extract_text_from_file(f)
                if t.strip():
                    new_texts.append(t)
            if pasted_examples.strip():
                new_texts.extend(
                    [chunk.strip() for chunk in pasted_examples.split("\n---\n") if chunk.strip()]
                )
            if new_texts:
                add_style_examples(new_texts)
                st.success(f"Добавлено: {len(new_texts)}")
                st.rerun()

        library = load_style_library()
        if library:
            if "selected_style_ids" not in st.session_state:
                st.session_state.selected_style_ids = [
                    e["id"] for e in library[-MAX_STYLE_EXAMPLES_IN_PROMPT:]
                ]

            st.caption(f"В библиотеке: {len(library)}")
            for e in reversed(library):
                col1, col2 = st.columns([5, 1])
                with col1:
                    checked = st.checkbox(
                        e["preview"][:42],
                        value=e["id"] in st.session_state.selected_style_ids,
                        key=f"chk_{e['id']}",
                    )
                    if checked and e["id"] not in st.session_state.selected_style_ids:
                        st.session_state.selected_style_ids.append(e["id"])
                    elif not checked and e["id"] in st.session_state.selected_style_ids:
                        st.session_state.selected_style_ids.remove(e["id"])
                with col2:
                    if st.button("×", key=f"del_{e['id']}", help="Удалить", use_container_width=True):
                        delete_style_example(e["id"])
                        st.rerun()

    st.divider()
    st.caption("РАЗДЕЛЫ")

    if st.button("Новый пост", use_container_width=True):
        for key in defaults:
            st.session_state[key] = defaults[key]
        st.session_state.show_metrics = False
        st.session_state.show_yandex = False
        st.session_state.show_news = False
        st.rerun()

    nav1, nav2 = st.columns(2)
    with nav1:
        if st.button("Аналитика", use_container_width=True):
            st.session_state.show_metrics = True
            st.session_state.show_yandex = False
            st.session_state.show_news = False
            st.rerun()
    with nav2:
        if st.button("Новости", use_container_width=True):
            st.session_state.show_metrics = False
            st.session_state.show_yandex = False
            st.session_state.show_news = True
            st.rerun()

    if st.button("База знаний", use_container_width=True):
        st.session_state.show_metrics = False
        st.session_state.show_yandex = True
        st.session_state.show_news = False
        st.rerun()

# ============================================================
# ОСНОВНОЙ ЭКРАН
# ============================================================
if not st.session_state.show_metrics and not st.session_state.show_yandex and not st.session_state.show_news:
    render_page_hero(
        "PR-портал ЦСР",
        "Подготовка публикаций, работа с базой знаний, проверка фактуры и аналитика — в одном рабочем пространстве.",
        "ЦСР · CONTENT INTELLIGENCE",
    )
    left_column, right_column = st.columns([1.08, 0.92], gap="large")

    with left_column:
        st.subheader("Исходные данные")
        with st.container(border=True):
            st.caption(f"Рубрика: {post_category}")
            field_values = build_dynamic_fields(post_category)

            speaker = st.text_input("👤 Спикер", placeholder="Имя, фамилия, должность")
            materials_text = st.text_area("📝 Материалы / тезисы", height=100)

            # Новость, выбранная в разделе «Мониторинг» кнопкой 📋
            news_block = ""
            if st.session_state.news_for_post:
                n = st.session_state.news_for_post
                with st.container(border=True):
                    c1, c2 = st.columns([5, 1])
                    with c1:
                        st.caption("📰 Новость как источник")
                        st.markdown(f"**{n['title']}**")
                        st.caption(f"{n['source']} • {n.get('published', '')[:10]}")
                    with c2:
                        if st.button("✕", key="drop_news", use_container_width=True):
                            st.session_state.news_for_post = None
                            st.rerun()
                news_block = (
                    f"Новость-источник: {n['title']}\n"
                    f"Издание: {n['source']}\n"
                    f"Дата: {n.get('published', '')}\n"
                    f"Ссылка: {n['url']}\n"
                    f"Краткое содержание: {n.get('description', '')}"
                )
            
            source_tab1, source_tab2 = st.tabs(["Загрузить файлы", "Яндекс.Диск"])
            
            uploaded_files = None
            with source_tab1:
                uploaded_files = st.file_uploader(
                    "📁 Файлы (PDF, PPTX, DOCX, XLSX, TXT)",
                    type=["pdf", "pptx", "docx", "xlsx", "txt"],
                    accept_multiple_files=True,
                    key="main_uploader"
                )
            
            with source_tab2:
                if disk_api:
                    st.info("📂 Файлы из базы знаний ЦСР")
                    base_path = YANDEX_BASE_PATH
                    folders_data = cached_list_disk(base_path)
                    
                    if folders_data:
                        folders = [f for f in folders_data if f.get("type") == "dir"]
                        files = [f for f in folders_data if f.get("type") == "file"]
                        
                        folder_names = ["📁 Все файлы"] + [f"📁 {f['name']}" for f in folders]
                        selected_folder_display = st.selectbox(
                            "Выберите тему:",
                            folder_names,
                            key="disk_folder_select"
                        )
                        
                        if selected_folder_display == "📁 Все файлы":
                            files_to_show = files
                            current_path = base_path
                        else:
                            folder_name = selected_folder_display.replace("📁 ", "")
                            folder_obj = next((f for f in folders if f['name'] == folder_name), None)
                            if folder_obj:
                                files_to_show = cached_list_disk(folder_obj.get("path"))
                                files_to_show = [f for f in files_to_show if f.get("type") == "file"]
                                current_path = folder_obj.get("path")
                            else:
                                files_to_show = []
                        
                        if files_to_show:
                            st.write(f"**Файлы:** {len(files_to_show)}")
                            selected_disk_files = []
                            
                            for f in files_to_show:
                                ext = f['name'].lower().split('.')[-1]
                                if ext in ['pdf', 'docx', 'xlsx', 'txt', 'pptx', 'doc', 'xls']:
                                    file_icon = "📄" if ext == "pdf" else "📝" if ext in ["txt", "docx", "doc"] else "📊" if ext in ["xlsx", "xls"] else "🎬"
                                    if st.checkbox(f"{file_icon} {f['name']}", key=f"disk_{f.get('path')}"):
                                        selected_disk_files.append(f)
                            
                            if selected_disk_files:
                                if st.button("✓ Добавить с диска", use_container_width=True):
                                    disk_content = ""
                                    for f in selected_disk_files:
                                        path = f.get("path")
                                        content = extract_yandex_file(disk_api, path)
                                        if content:
                                            disk_content += f"### {f['name']}\n{content}\n\n"
                                    if disk_content:
                                        st.session_state.disk_content = disk_content
                                        st.success(f"✅ Загружено {len(selected_disk_files)} файлов с диска")
                        else:
                            st.info("📁 В выбранной папке нет поддерживаемых файлов")
                    else:
                        st.warning("❌ База знаний не найдена")
                else:
                    st.warning("⚠️ Яндекс.Диск не подключен")

            if st.button("▶️ Сформировать", type="primary", use_container_width=True):
                has_field_data = any(v.strip() for v in field_values.values() if v)
                disk_content = st.session_state.get("disk_content", "")
                
                if not has_field_data and not materials_text.strip() and not uploaded_files and not disk_content and not news_block:
                    st.warning("⚠️ Добавьте данные")
                else:
                    extracted_file_text = extract_text_from_files(uploaded_files) if uploaded_files else ""
                    fields_text = format_fields_as_text(post_category, field_values)

                    final_materials = "\n".join(filter(None, [
                        fields_text,
                        f"Спикер: {speaker}" if speaker else "",
                        materials_text,
                        news_block,
                        extracted_file_text,
                        disk_content
                    ]))
                    st.session_state.raw_source_data = final_materials

                    style_block = build_style_block(st.session_state.get("selected_style_ids", []))
                    system_rules = get_enhanced_system_rules(post_category, tone, text_length, use_citations)
                    if style_block:
                        system_rules += "\n" + style_block

                    st.session_state.chat_history = [
                        {"role": "system", "content": system_rules},
                        {"role": "user", "content": st.session_state.raw_source_data},
                    ]

                    with st.spinner("✨ Создаю пост..."):
                        result = call_openai(st.session_state.chat_history, TEMP_GENERATE)
                        if result:
                            st.session_state.current_text = result
                            st.session_state.chat_history.append({"role": "assistant", "content": result})
                            st.session_state.versions_history.append({"text": result, "label": "v1"})
                            save_post(post_category, tone, text_length, result, 
                                    st.session_state.raw_source_data, 1 if use_citations else 0)
                            st.session_state.disk_content = ""
                            st.rerun()

    with right_column:
        st.subheader("Результат")
        with st.container(border=True):
            text_tab, fact_tab, history_tab, export_tab = st.tabs(["Текст", "Проверка", "История", "Экспорт"])

            with text_tab:
                if st.session_state.current_text:
                    st.code(st.session_state.current_text, language="markdown")

                    st.divider()
                    st.write("🔄 **Правки:**")
                    with st.form("refine_form", clear_on_submit=True):
                        refine_prompt = st.text_input("Что изменить?", placeholder="Пример: сделай короче")
                        submitted = st.form_submit_button("✏️ Уточнить")
                        if submitted and refine_prompt.strip():
                            st.session_state.chat_history.append({"role": "user", "content": refine_prompt})
                            api_messages = trim_history(st.session_state.chat_history)
                            with st.spinner("✏️ Переписываю..."):
                                result = call_openai(api_messages, TEMP_REFINE)
                                if result:
                                    st.session_state.current_text = result
                                    st.session_state.chat_history.append({"role": "assistant", "content": result})
                                    v_num = len(st.session_state.versions_history) + 1
                                    st.session_state.versions_history.append({"text": result, "label": f"v{v_num}"})
                                    st.rerun()
                else:
                    st.info("👈 Заполните данные слева")

            with fact_tab:
                if st.session_state.current_text:
                    if st.button("🔍 Проверить", type="primary", use_container_width=True):
                        fact_prompt = f"""Проверь текст на ошибки:
1. Цифры, даты, фамилии — без искажений
2. Нет дублирующихся блоков
3. Правильность названий

ИСХОДНИКИ: {st.session_state.raw_source_data[:1000]}
ТЕКСТ: {st.session_state.current_text[:1000]}"""
                        with st.spinner("🔍 Проверяю..."):
                            result = call_openai([{"role": "user", "content": fact_prompt}], TEMP_FACTCHECK)
                            if result:
                                st.success("✅ Готово!")
                                st.write(result)
                else:
                    st.info("Сначала создайте текст")

            with history_tab:
                if st.session_state.versions_history:
                    for i, version in enumerate(reversed(st.session_state.versions_history)):
                        col1, col2 = st.columns([4, 1])
                        with col1:
                            st.markdown(f"**{version['label']}**")
                        with col2:
                            if st.button("✓", key=f"v_{i}", use_container_width=True):
                                st.session_state.current_text = version["text"]
                                st.rerun()
                        st.divider()
                else:
                    st.info("История версий")

            with export_tab:
                if st.session_state.current_text:
                    docx_data = create_docx(st.session_state.current_text)
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.download_button(
                            "📄 Word",
                            data=docx_data,
                            file_name="post.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    with col2:
                        st.download_button(
                            "📝 Markdown",
                            data=st.session_state.current_text,
                            file_name="post.md",
                            mime="text/markdown"
                        )
                    with col3:
                        st.download_button(
                            "📋 TXT",
                            data=st.session_state.current_text.encode('utf-8'),
                            file_name="post.txt",
                            mime="text/plain"
                        )
                else:
                    st.info("Создайте текст")

# ============================================================
# DASHBOARD МЕТРИК
# ============================================================
elif st.session_state.show_metrics:
    render_page_hero("Аналитика", "Статистика публикаций, распределение по рубрикам, тону и объему.", "ЦСР · DASHBOARD")
    
    if st.button("← Вернуться"):
        st.session_state.show_metrics = False
        st.rerun()
    
    st.divider()
    
    metrics = get_metrics()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">Всего постов</div>
            <div class="metric-value">{metrics['total_posts']}</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">За неделю</div>
            <div class="metric-value">{metrics['posts_week']}</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">За месяц</div>
            <div class="metric-value">{metrics['posts_month']}</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">С цитатами</div>
            <div class="metric-value">{metrics['with_citations']}</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.divider()
    
    col1, col2 = st.columns(2)
    
    def chart_from_pairs(pairs, label, strip_numbers=False):
        """Строит график с ЧИТАЕМЫМИ подписями по оси X."""
        if not pairs:
            st.info("Нет данных")
            return
        names = []
        for name, _ in pairs:
            n = name or "—"
            if strip_numbers:
                n = re.sub(r"^\d+\.\s*", "", n)
            names.append(n[:28])
        counts = [c for _, c in pairs]
        df = pd.DataFrame({label: counts}, index=names)
        st.bar_chart(df, use_container_width=True)

    with col1:
        st.subheader("По рубрикам")
        chart_from_pairs(metrics['by_category'], "Постов", strip_numbers=True)

    with col2:
        st.subheader("По тону")
        chart_from_pairs(metrics['by_tone'], "Постов")

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("По объему")
        chart_from_pairs(metrics['by_length'], "Постов")

    with col2:
        st.subheader("Использование цитат")
        # Раньше здесь передавался словарь из чисел — Streamlit падал с ошибкой
        chart_from_pairs(
            [("С цитатами", metrics['with_citations']),
             ("Без цитат", metrics['without_citations'])],
            "Постов"
        )
    
    st.divider()
    
    st.subheader("История последних постов")
    posts = get_posts_history(10)
    if posts:
        for post_id, category, tone, text, created_at, has_citations in posts:
            with st.expander(f"{category} • {created_at[:10]} • {'💬' if has_citations else ''}"):
                st.write(text[:200] + "...")
    else:
        st.info("Нет постов")

# ============================================================
# ЯНДЕКС.ДИСК БРАУЗЕР
# ============================================================
elif st.session_state.show_yandex:
    render_page_hero("База знаний", "Материалы ЦСР на Яндекс.Диске — для быстрого поиска и использования в публикациях.", "ЦСР · KNOWLEDGE BASE")
    
    if st.button("← Вернуться"):
        st.session_state.show_yandex = False
        st.rerun()
    
    st.divider()
    
    if not disk_api:
        st.error("⚠️ Яндекс.Диск не подключен")
        st.info("Добавьте YANDEX_DISK_TOKEN в Streamlit Cloud secrets")
    else:
        base_path = YANDEX_BASE_PATH
        
        col1, col2, col3 = st.columns([2, 1, 1])
        with col1:
            st.write(f"**📂 {base_path}**")
        with col2:
            pass
        with col3:
            if st.button("🔄 Обновить", use_container_width=True):
                st.rerun()
        
        st.divider()
        
        items = cached_list_disk(base_path)

        if not items:
            st.error(f"Папка «{base_path}» не найдена или пуста.")
            st.caption(
                "Имя папки на Диске должно совпадать с константой YANDEX_BASE_PATH "
                "в начале файла. Ниже — что реально лежит в корне вашего Диска:"
            )
            root = cached_list_disk("/")
            if root:
                for r in root:
                    kind = "📁" if r.get("type") == "dir" else "📄"
                    st.code(f'{kind}  {r.get("path")}')
                st.caption(
                    "Скопируйте нужный путь и подставьте его в YANDEX_BASE_PATH."
                )
            else:
                st.warning(
                    "Корень Диска тоже не читается — скорее всего, неверный "
                    "YANDEX_DISK_TOKEN или у токена нет прав на чтение Диска."
                )

        if items:
            folders = [f for f in items if f.get("type") == "dir"]
            files = [f for f in items if f.get("type") == "file"]
            
            st.write(f"**📊 Статистика:** {len(folders)} папок, {len(files)} файлов")
            st.divider()
            
            if folders:
                st.subheader("Тематические папки")
                for folder in sorted(folders, key=lambda x: x.get("name", "")):
                    folder_name = folder.get("name", "")
                    with st.expander(f"📁 {folder_name}"):
                        sub_items = cached_list_disk(folder.get("path"))
                        sub_files = [f for f in sub_items if f.get("type") == "file"]
                        
                        if sub_files:
                            st.write(f"**{len(sub_files)} файлов:**")
                            for f in sorted(sub_files, key=lambda x: x.get("name", "")):
                                file_name = f.get("name", "")
                                file_size = f.get("size", 0)
                                size_mb = file_size / (1024 * 1024) if file_size else 0
                                
                                ext = file_name.lower().split('.')[-1]
                                icon = "📄" if ext == "pdf" else "📝" if ext in ["txt", "docx", "doc"] else "📊" if ext in ["xlsx", "xls"] else "🎬" if ext == "pptx" else "🎨" if ext in ["ai", "psd"] else "📑"
                                
                                col1, col2, col3 = st.columns([2.5, 0.5, 0.5])
                                with col1:
                                    st.write(f"{icon} {file_name}")
                                with col2:
                                    st.caption(f"{size_mb:.1f} MB")
                                with col3:
                                    # download_button со ссылкой на диск: файл
                                    # подгружается только при самом нажатии
                                    st.download_button(
                                        "📥",
                                        data=lambda p=f.get("path"): (disk_api.download_file(p) or b""),
                                        file_name=file_name,
                                        key=f"dl_{f.get('path')}",
                                        use_container_width=True,
                                    )
                        else:
                            st.info("Папка пуста")
            
            if files:
                st.divider()
                st.subheader("Файлы в корне")
                for f in sorted(files, key=lambda x: x.get("name", "")):
                    file_name = f.get("name", "")
                    file_size = f.get("size", 0)
                    size_mb = file_size / (1024 * 1024) if file_size else 0
                    
                    ext = file_name.lower().split('.')[-1]
                    icon = "📄" if ext == "pdf" else "📝" if ext in ["txt", "docx"] else "📊" if ext in ["xlsx", "xls"] else "🎨" if ext in ["ai", "psd"] else "📑"
                    
                    col1, col2, col3 = st.columns([2.5, 0.5, 0.5])
                    with col1:
                        st.write(f"{icon} {file_name}")
                    with col2:
                        st.caption(f"{size_mb:.1f} MB")
                    with col3:
                        st.download_button(
                            "📥",
                            data=lambda p=f.get("path"): (disk_api.download_file(p) or b""),
                            file_name=file_name,
                            key=f"dl_root_{f.get('path')}",
                            use_container_width=True,
                        )
        else:
            st.warning(f"❌ Не удалось загрузить содержимое базы знаний")

# ============================================================
# МОНИТОРИНГ НОВОСТЕЙ
# ============================================================
elif st.session_state.show_news:
    render_page_hero("Мониторинг новостей", "Поиск упоминаний ЦСР и релевантных инфоповодов с возможностью сразу использовать материал в посте.", "ЦСР · MEDIA MONITOR")
    
    if st.button("← Вернуться"):
        st.session_state.show_news = False
        st.rerun()
    
    st.divider()
    
    col1, col2, col3 = st.columns([2, 1, 1])
    
    with col1:
        query = st.text_input(
            "🔍 Поиск:",
            value="ЦСР OR 'Центр стратегических разработок'",
            placeholder="Поисковый запрос"
        )
    
    with col2:
        days = st.selectbox("📅 Период:", [7, 14, 30], index=0)
    
    with col3:
        st.write("")
        if st.button("🔄 Обновить", use_container_width=True):
            st.session_state.news_cache_buster = (
                st.session_state.get("news_cache_buster", 0) + 1
            )
            st.rerun()
    
    st.divider()
    
    tab1, tab2, tab3 = st.tabs(["🔍 Новости", "⭐ Избранное", "📊 Статистика"])
    
    with tab1:
        st.write("**Последние новости о ЦСР**")
        
        col_google, col_api = st.columns(2)
        with col_google:
            use_google = st.checkbox("📰 Google News (бесплатно)", value=True)
        with col_api:
            use_api = st.checkbox("🔐 News API", value=NEWS_API_KEY is not None, disabled=NEWS_API_KEY is None)
        
        with st.spinner("📰 Загружаю новости..."):
            news_list = cached_fetch_news(
                query, use_api, use_google, days,
                st.session_state.get("news_cache_buster", 0)
            )
        
        if news_list:
            st.write(f"**Найдено: {len(news_list)} новостей**")
            st.divider()
            
            for i, news in enumerate(news_list):
                title = news.get('title', 'N/A')
                source = news.get('source', 'Unknown')
                published = news.get('published', 'N/A')
                description = news.get('description', '')
                url = news.get('url', '')
                
                with st.container(border=True):
                    col1, col2 = st.columns([5, 1])
                    
                    with col1:
                        st.markdown(f"**[{title}]({url})**")
                        st.caption(f"📰 {source} • {published[:10]}")
                        
                        if description:
                            st.write(description[:300] + "..." if len(description) > 300 else description)
                    
                    with col2:
                        col_a, col_b = st.columns(2)
                        with col_a:
                            if st.button("⭐", key=f"fav_{i}", help="В избранное", use_container_width=True):
                                news_monitor.save_favorite(title, source, url, published, description)
                                st.success("✅")
                        
                        with col_b:
                            if st.button("📋", key=f"use_{i}", help="В пост", use_container_width=True):
                                st.session_state.news_for_post = {
                                    'title': title,
                                    'source': source,
                                    'url': url,
                                    'description': description,
                                    'published': published
                                }
                                st.session_state.show_news = False
                                st.info("✅ Новость добавлена как источник!")
                
                if i < len(news_list) - 1:
                    st.divider()
        else:
            st.info("📭 Новостей не найдено. Попробуйте другой запрос.")
    
    with tab2:
        st.write("**Избранные новости**")
        
        favorites = news_monitor.get_favorites()
        
        if favorites:
            st.write(f"**Всего: {len(favorites)} новостей**")
            st.divider()
            
            for title, source, url, published, description in favorites:
                with st.container(border=True):
                    col1, col2 = st.columns([5, 1])
                    
                    with col1:
                        st.markdown(f"**[{title}]({url})**")
                        st.caption(f"📰 {source} • {published[:10]}")
                        
                        if description:
                            st.write(description[:200] + "...")
                    
                    with col2:
                        if st.button("🗑️", key=f"del_fav_{title[:20]}", use_container_width=True):
                            news_monitor.delete_favorite(title)
                            st.rerun()
                
                st.divider()
        else:
            st.info("⭐ Избранные новости пока пусты")
    
    with tab3:
        st.write("**Статистика мониторинга**")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">Найдено новостей</div>
                <div class="metric-value">{len(news_list) if 'news_list' in locals() else 0}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            favorites = news_monitor.get_favorites()
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">В избранном</div>
                <div class="metric-value">{len(favorites)}</div>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            status = "🟢" if NEWS_API_KEY else "🔴"
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">News API</div>
                <div class="metric-value">{status}</div>
            </div>
            """, unsafe_allow_html=True)
        
        st.divider()
        st.info("""
        📰 **Источники новостей:**
        - **News API** - полнотекстовый поиск (если подключен)
        - **Google News** - RSS парсер (всегда доступен)
        
        💡 **Как использовать:**
        1. Введите поисковый запрос
        2. Выберите источники
        3. Нажмите "Обновить" для поиска
        4. ⭐ Сохраняйте интересные новости
        5. 📋 Добавляйте источники в пост
        """)
