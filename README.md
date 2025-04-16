# File Cleaner Documentation / Документация очистителя файлов

## Table of Contents / Содержание

- [Overview / Обзор](#overview--обзор)
- [Features / Функциональность](#features--функциональность)
- [Dependencies / Зависимости](#dependencies--зависимости)
- [Code Structure / Структура кода](#code-structure--структура-кода)
- [Usage / Использование](#usage--использование)
- [Configuration / Конфигурация](#configuration--конфигурация)
- [Error Handling & Backup / Обработка ошибок и резервное копирование](#error-handling--backup--обработка-ошибок-и-резервное-копирование)
---

## Overview / Обзор

**English:**  
The File Cleaner application is designed to remove metadata from various file formats including Microsoft Office documents (Word, Excel, PowerPoint), images (PNG, JPEG), audio files (MP3, WAV, FLAC), and PDFs. The application uses a combination of Python libraries to handle file parsing, metadata removal, and UI interactions through PyQt5.

**Русский:**  
Приложение «Очиститель файлов» предназначено для удаления метаданных из различных форматов файлов: документов Microsoft Office (Word, Excel, PowerPoint), изображений (PNG, JPEG), аудиофайлов (MP3, WAV, FLAC) и PDF. Для разбора файлов, удаления метаданных и взаимодействия с пользователем через интерфейс PyQt5, приложение использует набор библиотек Python.

---

## Features / Функциональность

**English:**
- **File Type Detection:** Automatically determines file types (PDF, PNG, JPEG, Office documents, audio files, etc.) based on magic-bytes.
- **Metadata Cleaning:** Clears metadata for various file formats using dedicated functions.
- **Extended Properties Cleanup:** For Office formats, additional XML-based extended properties are cleaned.
- **Backup & Rollback:** Creates a backup copy before modifying a file, with error handling to revert changes if needed.
- **GUI:** Provides a drag-and-drop interface with language support (English and Russian).
- **User Guide:** Built-in user guide is available via the interface.

**Русский:**  
- **Определение типа файла:** Автоматическое определение типа файла (PDF, PNG, JPEG, документы Office, аудиофайлы и т. д.) на основе magic-байтов.
- **Очистка метаданных:** Удаление метаданных для различных форматов файлов с использованием специализированных функций.
- **Очистка расширенных свойств:** Для форматов Office осуществляется очистка дополнительных XML-свойств.
- **Резервное копирование и откат:** Создание резервной копии перед изменением файла с возможностью отката при ошибках.
- **Графический интерфейс:** Поддерживается drag-and-drop интерфейс с переключением языка (английский и русский).
- **Руководство пользователя:** Встроенное руководство пользователя доступно через интерфейс.

---

## Dependencies / Зависимости

**English:**  
Make sure to install the following dependencies before running the application:

- Python 3.6+
- [python-docx](https://python-docx.readthedocs.io/en/latest/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [python-pptx](https://python-pptx.readthedocs.io/en/latest/)
- [Pillow](https://pillow.readthedocs.io/en/stable/)
- [pydub](https://github.com/jiaaro/pydub)
- [pikepdf](https://pikepdf.readthedocs.io/en/latest/)
- [lxml](https://lxml.de/)
- [PyQt5](https://doc.qt.io/)

You can install these dependencies via pip:

```bash
pip install -r requirements.txt
```

**Русский:**  
Перед запуском приложения убедитесь, что установлены следующие зависимости:

- Python 3.6+
- [python-docx](https://python-docx.readthedocs.io/en/latest/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [python-pptx](https://python-pptx.readthedocs.io/en/latest/)
- [Pillow](https://pillow.readthedocs.io/en/stable/)
- [pydub](https://github.com/jiaaro/pydub)
- [pikepdf](https://pikepdf.readthedocs.io/en/latest/)
- [lxml](https://lxml.de/)
- [PyQt5](https://doc.qt.io/)

Установить зависимости можно через pip:

```bash
pip install -r requirements.txt
```

---

## Code Structure / Структура кода

**English:**  
The project is organized into several functional areas:

- **File Type Detection:**  
  - `get_file_type(file_path)` reads file headers to determine the file format.
  
- **Metadata Cleaning Functions:**  
  - **Office Documents:**  
    - `clean_word_document(file_path)`
    - `clean_excel_document(file_path)`
    - `clean_powerpoint_document(file_path)`  
  - **Extended Properties:**  
    - `clean_ooxml_extended_properties(file_path)` cleans XML-based extended properties for Office documents.
  - **Images:**  
    - `clean_image(file_path)` creates a new image from the pixel data to remove metadata.
  - **Audio Files:**  
    - `clean_audio_file(file_path)` re-exports audio to remove metadata.
  - **PDF Files:**  
    - `clean_pdf_file(file_path)` uses pikepdf to remove document metadata.
  
- **File Cleaning Orchestration:**  
  - `clean_file(file_path)` creates a backup, determines the file type, calls the appropriate cleaning function, and handles errors with a rollback mechanism.
  
- **Settings Management:**  
  - `load_settings()` and `save_settings(settings)` manage a JSON configuration file for saving preferences (such as the language).
  
- **Graphical User Interface:**  
  - `FileListWidget`: A custom widget that handles drag and drop of file paths.
  - `FileCleanerUI`: The main application window built with PyQt5, including file list management, progress bar, language switching, and actions for cleaning files.
  
**Русский:**  
Проект разбит на несколько функциональных частей:

- **Определение типа файла:**  
  - Функция `get_file_type(file_path)` считывает заголовки файла для определения его формата.
  
- **Функции очистки метаданных:**  
  - **Документы Office:**  
    - `clean_word_document(file_path)`
    - `clean_excel_document(file_path)`
    - `clean_powerpoint_document(file_path)`  
  - **Очистка расширенных свойств:**  
    - `clean_ooxml_extended_properties(file_path)` очищает XML-свойства документов Office.
  - **Изображения:**  
    - `clean_image(file_path)` создает новое изображение из пиксельных данных для удаления метаданных.
  - **Аудиофайлы:**  
    - `clean_audio_file(file_path)` повторно экспортирует аудио для очистки метаданных.
  - **PDF-файлы:**  
    - `clean_pdf_file(file_path)` использует pikepdf для удаления метаданных из PDF.
  
- **Оркестрация очистки файлов:**  
  - Функция `clean_file(file_path)` создаёт резервную копию, определяет тип файла, вызывает соответствующую функцию очистки и реализует механизм отката при ошибках.
  
- **Управление настройками:**  
  - Функции `load_settings()` и `save_settings(settings)` работают с JSON-файлом настроек для сохранения предпочтений (например, языка интерфейса).
  
- **Графический интерфейс:**  
  - `FileListWidget`: Пользовательский виджет для поддержки перетаскивания файлов.
  - `FileCleanerUI`: Главное окно приложения на основе PyQt5, включающее управление списком файлов, индикатор выполнения, переключение языка и функции очистки файлов.

---

## Usage / Использование

**English:**  
1. **Starting the Application:**  
   Run the main script:
   ```bash
   python main.py
   ```
   The application window will launch with language switching options.

2. **Adding Files:**  
   Use the "Add Files" button or drag and drop files onto the list widget.

3. **Cleaning Files:**  
   Click on "Clean Files" to start the metadata cleaning process. The application will:
   - Create a backup of each file (with `.bak` extension).
   - Determine the file format.
   - Apply the appropriate cleaning function.
   - Delete the backup if cleaning is successful or revert on error.

4. **User Guidance:**  
   A built-in user guide is available by clicking the "User Guide" button.

**Русский:**  
1. **Запуск приложения:**  
   Запустите основной скрипт:
   ```bash
   python main.py
   ```
   Откроется окно приложения с возможностью переключения языка.

2. **Добавление файлов:**  
   Используйте кнопку «Добавить файлы» или перетащите файлы в область списка.

3. **Очистка файлов:**  
   Нажмите «Очистить файлы» для начала процесса удаления метаданных. Приложение:
   - Создает резервную копию каждого файла (с расширением `.bak`).
   - Определяет тип файла.
   - Вызывает соответствующую функцию очистки.
   - Удаляет резервную копию при успешном выполнении или восстанавливает оригинал в случае ошибки.

4. **Руководство пользователя:**  
   Встроенное руководство пользователя доступно по кнопке «Руководство пользователя».

---

## Configuration / Конфигурация

**English:**  
- **Settings File:**  
  The application uses a JSON file (located at `~/.file_cleaner_settings.json`) to store user settings such as the language preference.  
- **Modifying Defaults:**  
  To change default metadata values for Word or Excel, refer to the `WORD_DEFAULTS` and `EXCEL_DEFAULTS` dictionaries in the code.

**Русский:**  
- **Файл настроек:**  
  Приложение использует JSON-файл (расположенный по пути `~/.file_cleaner_settings.json`) для хранения настроек, таких как предпочтительный язык.
- **Изменение значений по умолчанию:**  
  Для изменения значений метаданных по умолчанию для Word или Excel, обратитесь к словарям `WORD_DEFAULTS` и `EXCEL_DEFAULTS` в коде.

---

## Error Handling & Backup / Обработка ошибок и резервное копирование

**English:**  
- Before modifying any file, the application creates a backup (file path appended with `.bak`).
- In case of any error during cleaning, the app attempts to revert changes using the backup.
- Exceptions are raised with detailed error messages for troubleshooting.

**Русский:**  
- Перед изменением файла приложение создает резервную копию (с расширением `.bak`).
- В случае возникновения ошибок приложение пытается восстановить исходный файл с помощью резервной копии.
