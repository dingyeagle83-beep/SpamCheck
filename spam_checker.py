import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import queue
import time
import os
import sys
import re
import traceback

# Импорты для Selenium (явные, чтобы PyInstaller их точно включил)
try:
    import selenium
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options as ChromeOptions
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from openpyxl import Workbook
except ImportError as e:
    print(f"Ошибка импорта: {e}")
    input("Установите зависимости: pip install selenium webdriver-manager openpyxl customtkinter")
    sys.exit(1)

# Определяем путь к ресурсам (для EXE)
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
    bundle_dir = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))
    bundle_dir = application_path

# Путь к иконке
ICON_PATH = os.path.join(bundle_dir, "icon.ico")

# ------------------------------------------------------------
# Настройки
# ------------------------------------------------------------
APP_NAME = "Проверка номеров на спам"
APP_VERSION = "1.0"
APP_AUTHOR = "Сергей Коротько"

YANDEX_URL = "https://ya.ru"
TBANK_URL = "https://www.tbank.ru/oleg/who-called/"

SPAM_KEYWORDS = [
    "нежелательный",
    "есть жалобы",
    "спам",
    "мошенники",
    "реклама",
    "обзвон"
]

DELAY_BETWEEN_SERVICES = 5
DELAY_BETWEEN_NUMBERS = 5
PAGE_LOAD_TIMEOUT = 10

# Цвета для тёмной темы (журнал)
COLOR_SPAM_DARK = "#FF5555"
COLOR_OK_DARK = "#4ADE80"
COLOR_ERROR_DARK = "#FBBF24"
COLOR_INFO_DARK = "#60A5FA"
COLOR_HEADER_DARK = "#E5E7EB"

# Цвета для светлой темы (журнал)
COLOR_SPAM_LIGHT = "#DC2626"
COLOR_OK_LIGHT = "#15803D"
COLOR_ERROR_LIGHT = "#B45309"
COLOR_INFO_LIGHT = "#1D4ED8"
COLOR_HEADER_LIGHT = "#1F2937"

# Фоны журнала
LOG_BG_DARK = "#1a1a1a"
LOG_FG_DARK = "#E5E7EB"
LOG_BG_LIGHT = "#f5f5f5"
LOG_FG_LIGHT = "#1a1a1a"

# Цвета фонов (тема-зависимые)
# (светлая тема, тёмная тема) — CustomTkinter принимает кортежи
BG_TAB = ("gray90", "gray15")
BG_FRAME = ("gray90", "gray15")


# ------------------------------------------------------------
# Чтение номеров из разных форматов файлов
# ------------------------------------------------------------
PHONE_PATTERN = re.compile(
    r'[\+]?[78][\s\-\(\)]*\d{3}[\s\-\(\)]*\d{3}[\s\-\(\)]*\d{2}[\s\-\(\)]*\d{2}'
)


def normalize_number(raw):
    """Приводит номер к единому виду +7XXXXXXXXXX."""
    if raw is None:
        return None
    
    if isinstance(raw, float):
        raw = str(int(raw))
    elif isinstance(raw, int):
        raw = str(raw)
    elif not isinstance(raw, str):
        raw = str(raw)
    
    raw = raw.strip()
    if not raw:
        return None
    
    digits = re.sub(r'\D', '', raw)
    
    if len(digits) == 11:
        if digits[0] in ('7', '8'):
            return '+7' + digits[1:]
        else:
            return '+' + digits
    elif len(digits) == 10:
        return '+7' + digits
    elif len(digits) == 12 and digits.startswith('7'):
        return '+' + digits
    else:
        return raw if raw else None


def extract_numbers_from_text(text):
    """Извлекает все номера из текста."""
    if not text:
        return []
    matches = PHONE_PATTERN.findall(text)
    result = []
    for m in matches:
        normalized = normalize_number(m)
        if normalized:
            result.append(normalized)
    return result


def read_numbers_from_txt(file_path):
    """Читает номера из txt-файла."""
    numbers = []
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            extracted = extract_numbers_from_text(line)
            if extracted:
                numbers.extend(extracted)
            else:
                normalized = normalize_number(line)
                if normalized:
                    numbers.append(normalized)
    return numbers


def read_numbers_from_excel(file_path):
    """Читает номера из Excel-файла (.xlsx или .xls)."""
    numbers = []
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == ".xlsx":
        from openpyxl import load_workbook
        wb = load_workbook(file_path, read_only=True, data_only=True)
        try:
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    for cell in row:
                        if cell is None:
                            continue
                        normalized = normalize_number(cell)
                        if normalized and len(re.sub(r'\D', '', normalized)) >= 10:
                            numbers.append(normalized)
                        else:
                            extracted = extract_numbers_from_text(str(cell))
                            numbers.extend(extracted)
        finally:
            wb.close()
    
    elif ext == ".xls":
        try:
            import xlrd
        except ImportError:
            raise Exception(
                "Для чтения .xls-файлов установите библиотеку xlrd:\n"
                "pip install xlrd"
            )
        
        wb = xlrd.open_workbook(file_path)
        for sheet in wb.sheets():
            for row_idx in range(sheet.nrows):
                for col_idx in range(sheet.ncols):
                    cell = sheet.cell_value(row_idx, col_idx)
                    if cell is None or cell == "":
                        continue
                    normalized = normalize_number(cell)
                    if normalized and len(re.sub(r'\D', '', normalized)) >= 10:
                        numbers.append(normalized)
                    else:
                        extracted = extract_numbers_from_text(str(cell))
                        numbers.extend(extracted)
    
    else:
        raise Exception(f"Неподдерживаемый формат Excel: {ext}")
    
    return numbers


def read_numbers_from_file(file_path):
    """Универсальная функция чтения номеров (определяет формат по расширению)."""
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext == ".txt":
        numbers = read_numbers_from_txt(file_path)
    elif ext in (".xlsx", ".xls"):
        numbers = read_numbers_from_excel(file_path)
    else:
        raise Exception(
            f"Неподдерживаемый формат файла: {ext}\n"
            "Поддерживаются: .txt, .xlsx, .xls"
        )
    
    seen = set()
    unique_numbers = []
    for n in numbers:
        if n not in seen:
            seen.add(n)
            unique_numbers.append(n)
    
    return unique_numbers


# ------------------------------------------------------------
# Логика проверки
# ------------------------------------------------------------
def get_chrome_driver():
    """Создает и возвращает драйвер Chrome со свёрнутым окном
    и игнорированием ошибок сертификатов."""
    options = ChromeOptions()
    
    # Основные опции для обхода автоматизации
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--incognito")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    # Опции для игнорирования SSL-ошибок и "небезопасного подключения"
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--ignore-ssl-errors")
    options.add_argument("--ignore-certificate-errors-spki-list")
    options.add_argument("--allow-running-insecure-content")
    options.add_argument("--disable-web-security")
    options.add_argument("--allow-insecure-localhost")
    options.add_argument("--disable-features=BlockInsecurePrivateNetworkRequests")
    options.add_argument("--disable-features=InsecureDownloadWarnings")
    options.add_argument("--disable-features=SSLVersionMin")
    options.add_argument("--safebrowsing-disable-download-protection")
    
    # Отключаем проверку сертификатов через capabilities
    options.set_capability("acceptInsecureCerts", True)
    
    # Свёрнутый режим
    options.add_argument("--start-minimized")
    options.add_argument("--window-position=-2400,-2400")
    options.add_argument("--window-size=800,600")
    
    from webdriver_manager.chrome import ChromeDriverManager
    driver_path = ChromeDriverManager().install()
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=options)
    return driver


def handle_security_warning(driver):
    """
    Если Chrome показал страницу 'Подключение не защищено',
    пытается нажать 'Дополнительно' → 'Перейти на сайт (небезопасно)'.
    Возвращает True, если удалось обойти.
    """
    try:
        title = driver.title.lower()
        page_source = driver.page_source.lower()
        
        markers = [
            "не защищено",
            "not secure",
            "privacy error",
            "your connection is not private",
            "подключение не защищено",
        ]
        
        is_warning_page = any(m in title for m in markers) or \
                          any(m in page_source for m in markers)
        
        if not is_warning_page:
            return False
        
        # Вариант 1: через JS-команду (Chrome)
        try:
            driver.execute_script("sendCommand('continue')")
            return True
        except:
            pass
        
        # Вариант 2: ищем кнопку "Дополнительно"
        try:
            details_btn = driver.find_element(By.ID, "details-button")
            details_btn.click()
            time.sleep(1)
            
            proceed_btn = driver.find_element(By.ID, "proceed-link")
            proceed_btn.click()
            time.sleep(2)
            return True
        except:
            pass
        
        # Вариант 3: ищем ссылку с текстом "proceed" / "перейти"
        try:
            links = driver.find_elements(By.TAG_NAME, "a")
            for link in links:
                text = link.text.lower()
                if "proceed" in text or "перейти" in text or "продолжить" in text:
                    link.click()
                    time.sleep(2)
                    return True
        except:
            pass
        
        return False
    except Exception:
        return False


def is_spam_text(text):
    if not text:
        return False
    text_lower = text.lower()
    return any(keyword in text_lower for keyword in SPAM_KEYWORDS)


def find_spam_keywords(text):
    if not text:
        return []
    text_lower = text.lower()
    return [kw for kw in SPAM_KEYWORDS if kw in text_lower]


def check_yandex(driver, number):
    """Проверяет номер через поиск Яндекса."""
    try:
        driver.get(YANDEX_URL)
        time.sleep(1)
        
        # Обходим предупреждение о безопасности, если оно появилось
        if handle_security_warning(driver):
            time.sleep(1)
        
        search_box = WebDriverWait(driver, PAGE_LOAD_TIMEOUT).until(
            EC.presence_of_element_located((By.NAME, "text"))
        )
        search_box.clear()
        search_box.send_keys(number)
        search_box.send_keys(Keys.RETURN)

        time.sleep(5)
        
        result_text = ""
        
        try:
            aon_blocks = driver.find_elements(By.CSS_SELECTOR, ".Aon")
            if aon_blocks:
                for block in aon_blocks:
                    result_text += block.text.strip() + " "
            
            if not result_text:
                header_blocks = driver.find_elements(By.CSS_SELECTOR, ".Aon-Header")
                if header_blocks:
                    for header in header_blocks:
                        result_text += header.text.strip() + " "
                
                if not result_text:
                    h3_elements = driver.find_elements(By.CSS_SELECTOR, "h3.AonProvider-Title")
                    if h3_elements:
                        for h3 in h3_elements:
                            result_text += h3.text.strip() + " "
        except:
            pass
        
        if not result_text:
            try:
                aon_elements = driver.find_elements(By.CSS_SELECTOR, "[data-fast-name='aon']")
                if aon_elements:
                    for element in aon_elements:
                        result_text += element.text.strip() + " "
            except:
                pass
        
        if not result_text:
            try:
                clean_number = number.replace("+", "").replace(" ", "").replace("-", "")
                elements_with_number = driver.find_elements(
                    By.XPATH,
                    f"//*[contains(text(), '{clean_number}') or contains(text(), '{number}')]"
                )
                for element in elements_with_number[:5]:
                    if element.text.strip():
                        result_text += element.text.strip() + " "
            except:
                pass
        
        spam_found = is_spam_text(result_text)
        found_keywords = find_spam_keywords(result_text) if spam_found else []
        
        return spam_found, None, found_keywords
        
    except Exception as e:
        return False, str(e), []


def check_tbank(driver, number):
    """Проверяет номер через веб-интерфейс Т-Банка."""
    try:
        driver.get(TBANK_URL)
        time.sleep(2)
        
        # Обходим предупреждение о безопасности, если оно появилось
        if handle_security_warning(driver):
            time.sleep(2)
        
        time.sleep(1)
        
        input_field = None
        selectors_to_try = [
            "input[type='tel']",
            "input[type='phone']",
            "input[placeholder*='телефон']",
            "input[placeholder*='номер']",
            "input[name*='phone']",
            "input[name*='tel']",
            "input[class*='phone']",
            "input[class*='tel']",
            "input[class*='number']"
        ]
        
        for selector in selectors_to_try:
            try:
                input_field = WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                if input_field:
                    break
            except:
                continue
        
        if not input_field:
            try:
                inputs = driver.find_elements(By.TAG_NAME, "input")
                if inputs:
                    input_field = inputs[0]
            except:
                pass
        
        if not input_field:
            raise Exception("Не удалось найти поле ввода на странице Т-Банка")
        
        input_field.clear()
        input_field.send_keys(number)
        input_field.send_keys(Keys.RETURN)
        time.sleep(5)
        
        result_text = ""
        
        try:
            main_blocks = driver.find_elements(By.CSS_SELECTOR, ".eb2YDDZ1O")
            if main_blocks:
                for block in main_blocks:
                    result_text += block.text.strip() + " "
            
            if not result_text:
                headers = driver.find_elements(By.CSS_SELECTOR, "[data-qa-type='mvno/numberInfoHeader']")
                if headers:
                    for header in headers:
                        result_text += header.text.strip() + " "
            
            if not result_text:
                categories = driver.find_elements(By.CSS_SELECTOR, ".ibshbJCH_")
                if categories:
                    for category in categories:
                        result_text += category.text.strip() + " "
            
            if not result_text:
                info_blocks = driver.find_elements(By.CSS_SELECTOR, ".abshbJCH_")
                if info_blocks:
                    for block in info_blocks:
                        result_text += block.text.strip() + " "
        except:
            pass
        
        if not result_text:
            try:
                number_elements = driver.find_elements(By.CSS_SELECTOR, "[data-qa-type*='number']")
                if number_elements:
                    for element in number_elements:
                        result_text += element.text.strip() + " "
            except:
                pass
        
        if not result_text:
            try:
                clean_number = number.replace("+", "").replace(" ", "").replace("-", "")
                elements_with_number = driver.find_elements(
                    By.XPATH,
                    f"//*[contains(text(), '{clean_number}') or contains(text(), '{number}')]"
                )
                for element in elements_with_number[:3]:
                    try:
                        parent = element.find_element(By.XPATH, "..")
                        parent_text = parent.text.strip()
                        if parent_text and len(parent_text) > 10:
                            result_text += parent_text + " "
                    except:
                        if element.text.strip():
                            result_text += element.text.strip() + " "
            except:
                pass
        
        spam_found = is_spam_text(result_text)
        found_keywords = find_spam_keywords(result_text) if spam_found else []
        
        return spam_found, None, found_keywords
        
    except Exception as e:
        return False, str(e), []


def save_results(results, output_dir):
    """Сохраняет результаты в файлы."""
    report_path = os.path.join(output_dir, "report.txt")
    results_spam_path = os.path.join(output_dir, "results_spam.xlsx")
    results_nonspam_path = os.path.join(output_dir, "results_nonspam.xlsx")
    
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("Отчёт о проверке номеров на спам\n")
        f.write("=" * 50 + "\n")
        f.write(f"Ключевые слова: {', '.join(SPAM_KEYWORDS)}\n")
        f.write("=" * 50 + "\n\n")
        
        for r in results:
            f.write(f"Номер: {r['number']}\n")
            f.write(f"  Яндекс: {'СПАМ' if r['yandex_spam'] else 'не спам'}")
            if r['yandex_error']:
                f.write(f" (ошибка: {r['yandex_error']})")
            f.write("\n")
            f.write(f"  Т-Банк: {'СПАМ' if r['tbank_spam'] else 'не спам'}")
            if r['tbank_error']:
                f.write(f" (ошибка: {r['tbank_error']})")
            f.write("\n")
            f.write(f"  Итог: {'СПАМ' if r['is_spam'] else 'Не спам'}")
            if r['spam_services']:
                f.write(f" (сервисы: {r['spam_services']})")
            f.write("\n\n")
    
    headers = ["Номер", "Яндекс спам", "Яндекс ошибка", "Т-Банк спам",
               "Т-Банк ошибка", "Итог спам", "Сервисы спама"]
    
    wb_spam = Workbook()
    ws_spam = wb_spam.active
    ws_spam.title = "Спам номера"
    ws_spam.append(headers)
    for r in results:
        if r["is_spam"]:
            ws_spam.append([
                r["number"],
                "Да" if r["yandex_spam"] else "Нет",
                r["yandex_error"] or "",
                "Да" if r["tbank_spam"] else "Нет",
                r["tbank_error"] or "",
                "Да" if r["is_spam"] else "Нет",
                r["spam_services"]
            ])
    wb_spam.save(results_spam_path)
    
    wb_nonspam = Workbook()
    ws_nonspam = wb_nonspam.active
    ws_nonspam.title = "Не спам номера"
    ws_nonspam.append(headers)
    for r in results:
        if not r["is_spam"]:
            ws_nonspam.append([
                r["number"],
                "Да" if r["yandex_spam"] else "Нет",
                r["yandex_error"] or "",
                "Да" if r["tbank_spam"] else "Нет",
                r["tbank_error"] or "",
                "Да" if r["is_spam"] else "Нет",
                r["spam_services"]
            ])
    wb_nonspam.save(results_nonspam_path)
    
    return report_path, results_spam_path, results_nonspam_path


# ------------------------------------------------------------
# Графический интерфейс на CustomTkinter
# ------------------------------------------------------------
class SpamCheckerGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Настройки окна
        self.title(APP_NAME)
        self.geometry("1000x800")
        self.minsize(850, 700)
        
        # Установка иконки окна
        try:
            if os.path.exists(ICON_PATH):
                self.iconbitmap(ICON_PATH)
            else:
                alt_icon = os.path.join(application_path, "icon.ico")
                if os.path.exists(alt_icon):
                    self.iconbitmap(alt_icon)
        except Exception as e:
            print(f"Не удалось установить иконку: {e}")
        
        # Тема
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        self.current_theme = "dark"
        
        # Переменные
        self.numbers_file = ctk.StringVar(value=os.path.join(application_path, "numbers.txt"))
        self.output_dir = ctk.StringVar(value=application_path)
        self.status_text = ctk.StringVar(value="Готов к работе")
        self.total_text = ctk.StringVar(value="0")
        self.checked_text = ctk.StringVar(value="0")
        self.spam_text = ctk.StringVar(value="0")
        self.errors_text = ctk.StringVar(value="0")
        
        # Флаги и очереди
        self.stop_flag = threading.Event()
        self.worker_thread = None
        self.message_queue = queue.Queue()
        self.is_running = False
        
        self.setup_ui()
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.process_queue()
    
    def apply_log_tags(self):
        """Применяет цвета тегов журнала в зависимости от текущей темы."""
        if self.current_theme == "dark":
            self.log_text.tag_config("spam", foreground=COLOR_SPAM_DARK,
                                     font=("Consolas", 10, "bold"))
            self.log_text.tag_config("ok", foreground=COLOR_OK_DARK)
            self.log_text.tag_config("error", foreground=COLOR_ERROR_DARK)
            self.log_text.tag_config("info", foreground=COLOR_INFO_DARK)
            self.log_text.tag_config("header", foreground=COLOR_HEADER_DARK,
                                     font=("Consolas", 10, "bold"))
            self.log_text.tag_config("header_green",
                                     foreground=COLOR_OK_DARK,
                                     font=("Consolas", 11, "bold"))
        else:
            self.log_text.tag_config("spam", foreground=COLOR_SPAM_LIGHT,
                                     font=("Consolas", 10, "bold"))
            self.log_text.tag_config("ok", foreground=COLOR_OK_LIGHT)
            self.log_text.tag_config("error", foreground=COLOR_ERROR_LIGHT)
            self.log_text.tag_config("info", foreground=COLOR_INFO_LIGHT)
            self.log_text.tag_config("header", foreground=COLOR_HEADER_LIGHT,
                                     font=("Consolas", 10, "bold"))
            self.log_text.tag_config("header_green",
                                     foreground=COLOR_OK_LIGHT,
                                     font=("Consolas", 11, "bold"))
    
    def toggle_theme(self):
        """Переключает тему между светлой и тёмной."""
        if self.current_theme == "dark":
            self.current_theme = "light"
            ctk.set_appearance_mode("light")
            self.theme_button.configure(text="🌙 Тёмная тема")
            
            self.log_text.configure(
                bg=LOG_BG_LIGHT,
                fg=LOG_FG_LIGHT,
                insertbackground=LOG_FG_LIGHT
            )
        else:
            self.current_theme = "dark"
            ctk.set_appearance_mode("dark")
            self.theme_button.configure(text="☀️ Светлая тема")
            
            self.log_text.configure(
                bg=LOG_BG_DARK,
                fg=LOG_FG_DARK,
                insertbackground=LOG_FG_DARK
            )
        
        # Обновляем цвета всех тегов
        self.apply_log_tags()
    
    def setup_ui(self):
        """Создает все элементы интерфейса."""
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # ===== Заголовок с кнопкой темы =====
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, padx=20, pady=(20, 5), sticky="ew")
        
        title = ctk.CTkLabel(
            header_frame,
            text="🔍 Проверка номеров на спам",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title.pack(side="left")
        
        subtitle = ctk.CTkLabel(
            header_frame,
            text="Яндекс + Т-Банк",
            font=ctk.CTkFont(size=13),
            text_color="gray"
        )
        subtitle.pack(side="left", padx=(10, 0), pady=(6, 0))
        
        self.theme_button = ctk.CTkButton(
            header_frame,
            text="☀️ Светлая тема",
            width=140,
            height=32,
            fg_color=("gray75", "gray30"),
            hover_color=("gray65", "gray40"),
            command=self.toggle_theme
        )
        self.theme_button.pack(side="right")
        
        # ===== Вкладки =====
        self.tabview = ctk.CTkTabview(
            self,
            fg_color=BG_TAB,
            segmented_button_selected_color=("#3B82F6", "#2563EB"),
            segmented_button_selected_hover_color=("#2563EB", "#1D4ED8"),
        )
        self.tabview.grid(row=1, column=0, padx=20, pady=(5, 20), sticky="nsew")
        
        self.tabview.add("Проверка")
        self.tabview.add("О программе")
        
        tab_check = self.tabview.tab("Проверка")
        tab_check.grid_columnconfigure(0, weight=1)
        tab_check.grid_rowconfigure(3, weight=1)
        
        tab_about = self.tabview.tab("О программе")
        tab_about.grid_columnconfigure(0, weight=1)
        tab_about.grid_rowconfigure(0, weight=1)
        
        # ===== ВКЛАДКА "ПРОВЕРКА" =====
        
        files_frame = ctk.CTkFrame(tab_check)
        files_frame.grid(row=0, column=0, padx=15, pady=(15, 10), sticky="ew")
        files_frame.grid_columnconfigure(1, weight=1)
        
        ctk.CTkLabel(
            files_frame, text="📄 Файл (txt / xlsx / xls):",
            font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=0, column=0, padx=(15, 10), pady=(15, 8), sticky="w")
        
        ctk.CTkEntry(
            files_frame, textvariable=self.numbers_file,
            height=32
        ).grid(row=0, column=1, padx=(0, 10), pady=(15, 8), sticky="ew")
        
        ctk.CTkButton(
            files_frame, text="Обзор", width=90, height=32,
            command=self.browse_numbers_file
        ).grid(row=0, column=2, padx=(0, 15), pady=(15, 8))
        
        ctk.CTkLabel(
            files_frame, text="📁 Папка результатов:",
            font=ctk.CTkFont(size=13, weight="bold")
        ).grid(row=1, column=0, padx=(15, 10), pady=(8, 15), sticky="w")
        
        ctk.CTkEntry(
            files_frame, textvariable=self.output_dir,
            height=32
        ).grid(row=1, column=1, padx=(0, 10), pady=(8, 15), sticky="ew")
        
        ctk.CTkButton(
            files_frame, text="Обзор", width=90, height=32,
            command=self.browse_output_dir
        ).grid(row=1, column=2, padx=(0, 15), pady=(8, 15))
        
        buttons_frame = ctk.CTkFrame(tab_check, fg_color="transparent")
        buttons_frame.grid(row=1, column=0, padx=15, pady=5, sticky="ew")
        
        self.start_button = ctk.CTkButton(
            buttons_frame,
            text="▶  НАЧАТЬ ПРОВЕРКУ",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=45,
            fg_color="#16A34A",
            hover_color="#15803D",
            command=self.start_checking
        )
        self.start_button.pack(side="left", padx=(0, 10), fill="x", expand=True)
        
        self.stop_button = ctk.CTkButton(
            buttons_frame,
            text="■  ОСТАНОВИТЬ",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=45,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            state="disabled",
            command=self.stop_checking
        )
        self.stop_button.pack(side="left", padx=(0, 0), fill="x", expand=True)
        
        progress_frame = ctk.CTkFrame(tab_check)
        progress_frame.grid(row=2, column=0, padx=15, pady=10, sticky="ew")
        progress_frame.grid_columnconfigure(0, weight=1)
        
        self.progress_bar = ctk.CTkProgressBar(
            progress_frame, height=12,
            progress_color="#3B82F6"
        )
        self.progress_bar.grid(row=0, column=0, columnspan=4, padx=15, pady=(15, 5), sticky="ew")
        self.progress_bar.set(0)
        
        ctk.CTkLabel(
            progress_frame,
            textvariable=self.status_text,
            font=ctk.CTkFont(size=12),
            anchor="w"
        ).grid(row=1, column=0, columnspan=4, padx=15, pady=(0, 10), sticky="ew")
        
        ctk.CTkFrame(progress_frame, height=1, fg_color=("gray70", "gray30")).grid(
            row=2, column=0, columnspan=4, padx=15, pady=5, sticky="ew"
        )
        
        stats_data = [
            ("Всего", self.total_text, "#6B7280"),
            ("Проверено", self.checked_text, "#3B82F6"),
            ("Спам", self.spam_text, "#EF4444"),
            ("Ошибок", self.errors_text, "#F59E0B"),
        ]
        
        for i, (label, var, color) in enumerate(stats_data):
            tile = ctk.CTkFrame(progress_frame, fg_color="transparent")
            tile.grid(row=3, column=i, padx=10, pady=(5, 15), sticky="ew")
            progress_frame.grid_columnconfigure(i, weight=1)
            
            ctk.CTkLabel(
                tile, text=label,
                font=ctk.CTkFont(size=11),
                text_color="gray"
            ).pack()
            
            ctk.CTkLabel(
                tile, textvariable=var,
                font=ctk.CTkFont(size=22, weight="bold"),
                text_color=color
            ).pack()
        
        log_frame = ctk.CTkFrame(tab_check)
        log_frame.grid(row=3, column=0, padx=15, pady=(5, 15), sticky="nsew")
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)
        
        log_header = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_header.grid(row=0, column=0, padx=15, pady=(15, 5), sticky="ew")
        
        ctk.CTkLabel(
            log_header,
            text="📋 Журнал проверки",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(side="left")
        
        ctk.CTkButton(
            log_header,
            text="Очистить",
            width=90,
            height=28,
            fg_color=("gray75", "gray30"),
            hover_color=("gray65", "gray40"),
            command=self.clear_log
        ).pack(side="right")
        
        text_container = ctk.CTkFrame(log_frame, fg_color="transparent")
        text_container.grid(row=1, column=0, padx=15, pady=(5, 15), sticky="nsew")
        text_container.grid_columnconfigure(0, weight=1)
        text_container.grid_rowconfigure(0, weight=1)
        
        self.log_text = tk.Text(
            text_container,
            wrap=tk.WORD,
            font=("Consolas", 10),
            bg=LOG_BG_DARK,
            fg=LOG_FG_DARK,
            insertbackground=LOG_FG_DARK,
            relief="flat",
            borderwidth=0,
            padx=10,
            pady=10,
            highlightthickness=0
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        
        scrollbar = ctk.CTkScrollbar(text_container, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.apply_log_tags()
        
        self.log_text.configure(state="disabled")
        
        # ===== ВКЛАДКА "О ПРОГРАММЕ" =====
        self.setup_about_tab(tab_about)
    
    def setup_about_tab(self, parent):
        """Создаёт содержимое вкладки 'О программе'."""
        
        # Используем обычный Frame вместо CTkScrollableFrame,
        # чтобы избежать проблем с фоновыми цветами CTkLabel
        container = ctk.CTkScrollableFrame(
            parent,
            fg_color=BG_FRAME,
            scrollbar_button_color=("gray75", "gray30"),
            scrollbar_button_hover_color=("gray65", "gray40"),
        )
        container.grid(row=0, column=0, padx=15, pady=15, sticky="nsew")
        container.grid_columnconfigure(0, weight=1)
        
        # Иконка-эмодзи
        ctk.CTkLabel(
            container,
            text="🔍",
            font=ctk.CTkFont(size=72),
            fg_color=BG_FRAME,
        ).grid(row=0, column=0, pady=(20, 10))
        
        # Название
        ctk.CTkLabel(
            container,
            text=APP_NAME,
            font=ctk.CTkFont(size=24, weight="bold"),
            fg_color=BG_FRAME,
        ).grid(row=1, column=0, pady=(0, 5))
        
        # Версия
        ctk.CTkLabel(
            container,
            text=f"Версия {APP_VERSION}",
            font=ctk.CTkFont(size=13),
            text_color="gray",
            fg_color=BG_FRAME,
        ).grid(row=2, column=0, pady=(0, 20))
        
        # Разделитель
        ctk.CTkFrame(container, height=1, fg_color=("gray70", "gray40")).grid(
            row=3, column=0, padx=40, pady=10, sticky="ew"
        )
        
        # ===== О программе =====
        ctk.CTkLabel(
            container,
            text="📖 О программе",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=BG_FRAME,
        ).grid(row=4, column=0, pady=(10, 8), sticky="w", padx=40)
        
        description_text = (
            "Программа предназначена для автоматической проверки телефонных "
            "номеров на наличие спама, мошенничества и нежелательных звонков.\n\n"
            "Проверка выполняется через два популярных сервиса:\n"
            "    • Яндекс (определитель номера)\n"
            "    • Т-Банк (сервис «Кто звонил»)\n\n"
            "Номера считаются спамом, если хотя бы в одном из сервисов "
            "найдено одно из следующих определений: «нежелательный», "
            "«есть жалобы», «спам», «мошенники», «реклама», «обзвон»."
        )
        
        ctk.CTkLabel(
            container,
            text=description_text,
            font=ctk.CTkFont(size=13),
            justify="left",
            anchor="w",
            wraplength=700,
            fg_color=BG_FRAME,
        ).grid(row=5, column=0, padx=40, pady=(0, 15), sticky="ew")
        
        # Разделитель
        ctk.CTkFrame(container, height=1, fg_color=("gray70", "gray40")).grid(
            row=6, column=0, padx=40, pady=10, sticky="ew"
        )
        
        # ===== Возможности =====
        ctk.CTkLabel(
            container,
            text="⚙️ Возможности",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=BG_FRAME,
        ).grid(row=7, column=0, pady=(10, 8), sticky="w", padx=40)
        
        features_text = (
            "✓ Чтение номеров из файлов .txt, .xlsx, .xls\n"
            "✓ Автоматическое извлечение номеров из любого текста\n"
            "✓ Приведение номеров к единому формату +7XXXXXXXXXX\n"
            "✓ Проверка в двух сервисах одновременно\n"
            "✓ Цветной журнал проверки в реальном времени\n"
            "✓ Прогресс-бар и статистика\n"
            "✓ Светлая и тёмная темы\n"
            "✓ Экспорт результатов в Excel (спам и не спам отдельно)\n"
            "✓ Возможность остановить проверку в любой момент"
        )
        
        ctk.CTkLabel(
            container,
            text=features_text,
            font=ctk.CTkFont(size=13),
            justify="left",
            anchor="w",
            wraplength=700,
            fg_color=BG_FRAME,
        ).grid(row=8, column=0, padx=40, pady=(0, 15), sticky="ew")
        
        # Разделитель
        ctk.CTkFrame(container, height=1, fg_color=("gray70", "gray40")).grid(
            row=9, column=0, padx=40, pady=10, sticky="ew"
        )
        
        # ===== Как использовать =====
        ctk.CTkLabel(
            container,
            text="📌 Как использовать",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=BG_FRAME,
        ).grid(row=10, column=0, pady=(10, 8), sticky="w", padx=40)
        
        how_text = (
            "1. Подготовьте файл с номерами (.txt, .xlsx или .xls)\n"
            "2. Выберите файл в поле «Файл с номерами»\n"
            "3. Выберите папку, куда сохранить результаты\n"
            "4. Нажмите кнопку «НАЧАТЬ ПРОВЕРКУ»\n"
            "5. Дождитесь завершения — результаты сохранятся автоматически"
        )
        
        ctk.CTkLabel(
            container,
            text=how_text,
            font=ctk.CTkFont(size=13),
            justify="left",
            anchor="w",
            wraplength=700,
            fg_color=BG_FRAME,
        ).grid(row=11, column=0, padx=40, pady=(0, 15), sticky="ew")
        
        # Разделитель
        ctk.CTkFrame(container, height=1, fg_color=("gray70", "gray40")).grid(
            row=12, column=0, padx=40, pady=10, sticky="ew"
        )
        
        # ===== Автор =====
        ctk.CTkLabel(
            container,
            text="👤 Автор",
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=BG_FRAME,
        ).grid(row=13, column=0, pady=(10, 5), sticky="w", padx=40)
        
        ctk.CTkLabel(
            container,
            text=APP_AUTHOR,
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=("#2563EB", "#60A5FA"),
            fg_color=BG_FRAME,
        ).grid(row=14, column=0, pady=(5, 5), sticky="w", padx=40)
        
        # Разделитель
        ctk.CTkFrame(container, height=1, fg_color=("gray70", "gray40")).grid(
            row=15, column=0, padx=40, pady=10, sticky="ew"
        )
        
        # ===== Копирайт =====
        ctk.CTkLabel(
            container,
            text=f"© 2025 {APP_AUTHOR}",
            font=ctk.CTkFont(size=11),
            text_color="gray",
            fg_color=BG_FRAME,
        ).grid(row=16, column=0, pady=(10, 20))
    
    def clear_log(self):
        """Очищает журнал."""
        self.log_text.configure(state="normal")
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state="disabled")
    
    def browse_numbers_file(self):
        """Выбор файла с номерами."""
        filename = filedialog.askopenfilename(
            title="Выберите файл с номерами",
            filetypes=[
                ("Все поддерживаемые", "*.txt *.xlsx *.xls"),
                ("Текстовые файлы", "*.txt"),
                ("Excel файлы", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("Все файлы", "*.*"),
            ],
            initialdir=application_path
        )
        if filename:
            self.numbers_file.set(filename)
    
    def browse_output_dir(self):
        """Выбор папки для результатов."""
        dirname = filedialog.askdirectory(
            title="Выберите папку для результатов",
            initialdir=application_path
        )
        if dirname:
            self.output_dir.set(dirname)
    
    def log(self, message, tag=None):
        """Добавляет сообщение в лог."""
        try:
            self.log_text.configure(state="normal")
            self.log_text.insert(tk.END, message + "\n", tag)
            self.log_text.see(tk.END)
            self.log_text.configure(state="disabled")
        except Exception:
            pass
    
    def start_checking(self):
        """Запускает проверку."""
        if self.is_running:
            return
        
        numbers_file = self.numbers_file.get()
        if not os.path.exists(numbers_file):
            messagebox.showerror("Ошибка", f"Файл не найден:\n{numbers_file}")
            return
        
        try:
            numbers = read_numbers_from_file(numbers_file)
        except Exception as e:
            messagebox.showerror(
                "Ошибка чтения файла",
                f"Не удалось прочитать файл:\n{numbers_file}\n\n{e}"
            )
            return
        
        if not numbers:
            messagebox.showwarning(
                "Внимание",
                "В файле не найдено ни одного номера телефона."
            )
            return
        
        self.tabview.set("Проверка")
        self.clear_log()
        
        self.total_text.set(str(len(numbers)))
        self.checked_text.set("0")
        self.spam_text.set("0")
        self.errors_text.set("0")
        
        self.progress_bar.set(0)
        
        self.is_running = True
        self.stop_flag.clear()
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        
        self.worker_thread = threading.Thread(
            target=self.worker,
            args=(numbers, self.output_dir.get()),
            daemon=True
        )
        self.worker_thread.start()
    
    def stop_checking(self):
        """Останавливает проверку."""
        if self.is_running:
            self.stop_flag.set()
            self.log("", None)
            self.log("⚠  Запрошена остановка... Дождитесь завершения текущего номера.",
                     "error")
            self.stop_button.configure(state="disabled")
    
    def worker(self, numbers, output_dir):
        """Рабочий поток."""
        driver = None
        results = []
        total = len(numbers)
        checked = 0
        spam_count = 0
        errors_count = 0
        
        try:
            self.message_queue.put(("log", "", None))
            self.message_queue.put(("log", "🌐 Инициализация браузера Chrome...", "info"))
            self.message_queue.put(("log", "📥 Поиск или скачивание ChromeDriver...", "info"))
            
            try:
                driver = get_chrome_driver()
                self.message_queue.put(("log", "✓  Браузер Chrome успешно запущен!", "ok"))
            except Exception as e:
                self.message_queue.put(("log", f"✗  Не удалось запустить Chrome: {e}", "error"))
                self.message_queue.put(("finished", None))
                return
            
            self.message_queue.put(("log", "", None))
            self.message_queue.put(("log", "=" * 65, "header"))
            self.message_queue.put(("log", f"🚀 Начинаю проверку {total} номеров", "header_green"))
            self.message_queue.put(("log", "=" * 65, "header"))
            self.message_queue.put(("log", "", None))
            
            for idx, number in enumerate(numbers, start=1):
                if self.stop_flag.is_set():
                    self.message_queue.put(
                        ("log", f"\n⚠  Остановлено на номере {idx} из {total}", "error")
                    )
                    break
                
                self.message_queue.put(("status", f"[{idx}/{total}] Проверяю номер: {number}"))
                self.message_queue.put(("log", f"┌─ [{idx}/{total}] Номер: {number}", "header"))
                
                yandex_spam, yandex_error, yandex_keywords = check_yandex(driver, number)
                
                if yandex_error:
                    self.message_queue.put(("log", f"│  🔍 Яндекс: ошибка — {yandex_error}", "error"))
                    errors_count += 1
                elif yandex_spam:
                    kw_str = ", ".join(yandex_keywords)
                    self.message_queue.put(
                        ("log", f"│  🔍 Яндекс: 🚫 СПАМ (найдено: {kw_str})", "spam")
                    )
                else:
                    self.message_queue.put(("log", "│  🔍 Яндекс: ✓ не спам", "ok"))
                
                if self.stop_flag.is_set():
                    self.message_queue.put(("log", "└─ ⚠  Остановлено пользователем", "error"))
                    break
                
                for _ in range(DELAY_BETWEEN_SERVICES):
                    if self.stop_flag.is_set():
                        break
                    time.sleep(1)
                
                if self.stop_flag.is_set():
                    self.message_queue.put(("log", "└─ ⚠  Остановлено пользователем", "error"))
                    break
                
                tbank_spam, tbank_error, tbank_keywords = check_tbank(driver, number)
                
                if tbank_error:
                    self.message_queue.put(("log", f"│  🏦 Т-Банк: ошибка — {tbank_error}", "error"))
                    errors_count += 1
                elif tbank_spam:
                    kw_str = ", ".join(tbank_keywords)
                    self.message_queue.put(
                        ("log", f"│  🏦 Т-Банк: 🚫 СПАМ (найдено: {kw_str})", "spam")
                    )
                else:
                    self.message_queue.put(("log", "│  🏦 Т-Банк: ✓ не спам", "ok"))
                
                spam_services = []
                if yandex_spam:
                    spam_services.append("Яндекс")
                if tbank_spam:
                    spam_services.append("Т-Банк")
                
                is_spam = bool(spam_services)
                
                if is_spam:
                    spam_count += 1
                    self.message_queue.put(
                        ("log",
                         f"└─ 📊 ИТОГ: 🚫 СПАМ в сервисах: {', '.join(spam_services)}",
                         "spam")
                    )
                else:
                    self.message_queue.put(
                        ("log", "└─ 📊 ИТОГ: ✓ спам не обнаружен", "ok")
                    )
                
                self.message_queue.put(("log", "", None))
                
                results.append({
                    "number": number,
                    "yandex_spam": yandex_spam,
                    "yandex_error": yandex_error,
                    "tbank_spam": tbank_spam,
                    "tbank_error": tbank_error,
                    "is_spam": is_spam,
                    "spam_services": ", ".join(spam_services) if spam_services else ""
                })
                
                checked += 1
                
                self.message_queue.put(("stats", (checked, spam_count, errors_count, total)))
                self.message_queue.put(("progress", idx / total))
                
                if idx < total and not self.stop_flag.is_set():
                    self.message_queue.put(
                        ("status", f"Ожидание перед следующим номером... ({DELAY_BETWEEN_NUMBERS} сек)")
                    )
                    for _ in range(DELAY_BETWEEN_NUMBERS):
                        if self.stop_flag.is_set():
                            break
                        time.sleep(1)
        
        except Exception as e:
            self.message_queue.put(("log", f"\n✗  Критическая ошибка: {e}", "error"))
            self.message_queue.put(("log", traceback.format_exc(), "error"))
        
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass
            
            if results:
                self.message_queue.put(("log", "", None))
                self.message_queue.put(("log", "=" * 65, "header"))
                self.message_queue.put(("log", "💾 Сохранение результатов...", "info"))
                try:
                    report_path, spam_path, nonspam_path = save_results(results, output_dir)
                    self.message_queue.put(("log", "✓  Сохранено:", "ok"))
                    self.message_queue.put(("log", f"   • {os.path.basename(report_path)} (отчёт)", "ok"))
                    self.message_queue.put(("log", f"   • {os.path.basename(spam_path)} (только спам)", "ok"))
                    self.message_queue.put(("log", f"   • {os.path.basename(nonspam_path)} (не спам)", "ok"))
                    self.message_queue.put(("log", "", None))
                    self.message_queue.put(("log", f"📁 Папка: {output_dir}", "info"))
                except Exception as e:
                    self.message_queue.put(("log", f"✗  Ошибка сохранения: {e}", "error"))
            
            self.message_queue.put(("finished", None))
    
    def process_queue(self):
        """Обрабатывает сообщения из очереди."""
        try:
            while True:
                try:
                    msg = self.message_queue.get_nowait()
                except queue.Empty:
                    break
                
                try:
                    msg_type = msg[0]
                    
                    if msg_type == "log":
                        _, text, tag = msg
                        self.log(text, tag)
                    
                    elif msg_type == "status":
                        self.status_text.set(msg[1])
                    
                    elif msg_type == "progress":
                        self.progress_bar.set(msg[1])
                    
                    elif msg_type == "stats":
                        stats_data = msg[1]
                        checked, spam_count, errors_count, total = stats_data
                        self.checked_text.set(str(checked))
                        self.spam_text.set(str(spam_count))
                        self.errors_text.set(str(errors_count))
                    
                    elif msg_type == "finished":
                        self.on_finished()
                
                except Exception as e:
                    print(f"Ошибка обработки сообщения: {e}")
                    traceback.print_exc()
        
        except Exception as e:
            print(f"Критическая ошибка в process_queue: {e}")
        
        self.after(80, self.process_queue)
    
    def on_finished(self):
        """По завершении проверки."""
        self.is_running = False
        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.status_text.set("✓ Завершено")
        self.log("", None)
        self.log("═" * 65, "header")
        self.log("🎉 Работа программы завершена.", "header_green")
        self.log("═" * 65, "header")
    
    def on_closing(self):
        """Обработка закрытия окна."""
        if self.is_running:
            if not messagebox.askyesno(
                "Подтверждение",
                "Проверка ещё выполняется. Вы уверены, что хотите закрыть программу?"
            ):
                return
            self.stop_flag.set()
        
        self.destroy()


# ------------------------------------------------------------
# Точка входа
# ------------------------------------------------------------
def main():
    app = SpamCheckerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()