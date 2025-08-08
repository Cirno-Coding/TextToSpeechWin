import sys
import os
import win32com.client
from datetime import datetime
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QDialog, QVBoxLayout, QLabel, \
    QPushButton, QHBoxLayout, QInputDialog, QLineEdit, QTextEdit
from PyQt6.QtCore import Qt, QTimer, QUrl
from PyQt6.QtGui import QTextCursor, QTextCharFormat, QColor, QFont, QDesktopServices, QStandardItem, QStandardItemModel
from ui.MainWindow import Ui_MainWindow
from version import VERSION, VERSION_NAME, BUILD_DATE, AUTHOR, GITHUB_URL
from database import DatabaseManager, Category, Text


class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("О программе")
        self.setFixedSize(500, 400)
        self.setModal(True)
        
        # Настройка стиля окна
        self.setStyleSheet("""
            QDialog {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #0a0a0a, stop:0.3 #1a1a2e, stop:0.7 #16213e, stop:1 #0f3460);
                color: #ffffff;
            }
            QLabel {
                color: #ffffff;
                background: transparent;
            }
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #1a1a2e, stop:1 #16213e);
                border: 2px solid #00ffff;
                border-radius: 10px;
                padding: 8px 16px;
                color: #ffffff;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #00ffff, stop:1 #0080ff);
                color: #000000;
                box-shadow: 0 0 15px #00ffff;
            }
        """)
        
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout()
        
        # Название приложения (большими буквами и жирным шрифтом)
        title_font = QFont()
        title_font.setPointSize(18)
        title_font.setBold(True)
        
        title_label = QLabel("TEXT-TO-SPEECH WINDOWS APPLICATION")
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setStyleSheet("color: #00ffff; margin: 20px; font-size: 20px;")
        layout.addWidget(title_label)
        
        # Версия приложения (мелким шрифтом)
        version_font = QFont()
        version_font.setPointSize(10)
        
        version_text = f"{VERSION_NAME} - {VERSION} - {BUILD_DATE} : {AUTHOR}"
        version_label = QLabel(version_text)
        version_label.setFont(version_font)
        version_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version_label.setStyleSheet("color: #cccccc; margin: 10px;")
        layout.addWidget(version_label)
        
        # Описание функционала
        description = """
        Приложение для преобразования текста в речь с использованием Windows SAPI.
        
        Основные возможности:
        • Преобразование текста в речь с русскими голосами
        • Управление воспроизведением (воспроизведение, пауза, остановка)
        • Навигация по тексту (переход между предложениями)
        • Визуальное выделение текущего воспроизводимого текста
        • Регулировка скорости воспроизведения
        • Работа с файлами (сохранение и загрузка)
        • Современный интерфейс с темной темой
        """
        
        desc_label = QLabel(description)
        desc_label.setWordWrap(True)
        desc_label.setAlignment(Qt.AlignmentFlag.AlignLeft)
        desc_label.setStyleSheet("color: #ffffff; margin: 20px; line-height: 1.5;")
        layout.addWidget(desc_label)
        
        # Кнопки
        button_layout = QHBoxLayout()
        
        # Кнопка GitHub
        github_btn = QPushButton("GitHub")
        github_btn.clicked.connect(self.open_github)
        button_layout.addWidget(github_btn)
        
        button_layout.addStretch()
        
        # Кнопка закрытия
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        
        layout.addLayout(button_layout)
        self.setLayout(layout)
    
    def open_github(self):
        """Открытие ссылки на GitHub"""
        QDesktopServices.openUrl(QUrl(GITHUB_URL))


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        
        # Устанавливаем фиксированный размер окна
        self.setFixedSize(800, 600)

        self.voice_list = []
        self.speaker = None
        self.is_playing = False
        self.is_pause = False
        
        # Новые переменные для управления воспроизведением
        self.sentences = []
        self.sentence_positions = []  # Позиции предложений в тексте
        self.current_sentence_index = 0
        self.current_text = ""
        self.playback_timer = QTimer()
        self.playback_timer.timeout.connect(self.check_playback_status)
        
        # Форматы для выделения текста
        self.highlight_format = QTextCharFormat()
        self.highlight_format.setBackground(QColor(0, 255, 255, 100))  # Полупрозрачный голубой
        self.highlight_format.setForeground(QColor(0, 0, 0))  # Черный текст
        
        self.normal_format = QTextCharFormat()
        self.normal_format.setBackground(QColor(0, 0, 0, 0))  # Прозрачный фон
        self.normal_format.setForeground(QColor(255, 255, 255))  # Белый текст

        # Инициализация объектов
        self.setup_voices()
        self.setup_connections()
        self.update_speed_label()
        self.update_button_states()

        # Инициализация базы данных с обработкой ошибок
        try:
            self.db = DatabaseManager()
        except RuntimeError as e:
            self.statusbar.showMessage(str(e), 10000)
            self.db = DatabaseManager()

        # Загрузка категорий
        self.load_categories()

        self.textBrowser.setAcceptRichText(False)
        self.textBrowser.focusOutEvent = self.save_current_text
        self.current_text_id = None

    def save_current_text(self, event):
        """Сохранение текста при потере фокуса"""
        if self.current_text_id is None:
            return

        current_content = self.textBrowser.toPlainText()
        try:
            # Получаем заголовок из списка
            title = self.textsList.model().itemFromIndex(self.textsList.currentIndex()).text()
            self.db.update_text(self.current_text_id, title, current_content)
            self.statusbar.showMessage("Текст успешно сохранён", 3000)
        except Exception as e:
            self.statusbar.showMessage(f"Ошибка сохранения текста: {str(e)}", 5000)

        # Вызываем оригинальный обработчик события
        super(QTextEdit, self.textBrowser).focusOutEvent(event)

    def update_button_states(self):
        """
        Обновление состояния кнопок в зависимости от статуса воспроизведения
        """
        has_text = bool(self.textBrowser.toPlainText().strip())
        can_control = has_text and (self.is_playing or self.is_pause)
        
        # Кнопка остановки активна только при наличии текста и активном воспроизведении/паузе
        self.BtnStop.setEnabled(can_control)
        
        # Кнопки навигации активны только при наличии текста и активном воспроизведении/паузе
        # и только если можно перейти к предыдущему/следующему предложению
        can_go_previous = can_control and self.sentences and self.current_sentence_index > 0
        can_go_next = can_control and self.sentences and self.current_sentence_index < len(self.sentences) - 1
        
        self.BtnPrevious.setEnabled(can_go_previous)
        self.BtnNext.setEnabled(can_go_next)
        
        # Кнопка воспроизведения активна только при наличии текста
        self.BtnPausePlay.setEnabled(has_text)

    def setup_voices(self):
        """
        Получение голосов из SAPI и добавление их в список
        """
        try:
            self.speaker = win32com.client.Dispatch("SAPI.SpVoice")
            voices = self.speaker.GetVoices()

            self.voice_list.clear()
            self.VoicesList.clear()

            for i in voices:
                voice_name = i.GetDescription()
                # Проверка на русские символы в имени голоса
                if any(keyword in voice_name.lower() for keyword in ["рус", "russian", "rus"]):
                    self.VoicesList.addItem(voice_name)
                    self.voice_list.append(i)

        except Exception as e:
            print(f"Ошибка при получении голосов: {e}")
            self.VoicesList.addItem("Не найдено русских голосов")

    def get_selected_voice(self):
        """
        Получение выбранного голоса из списка
        """
        ind = self.VoicesList.currentIndex()
        if 0 <= ind <= len(self.voice_list):
            return self.voice_list[ind]
        return None

    def on_category_changed(self, index):
        """Обработчик изменения выбранной категории"""
        if index >= 0:
            category_id = self.catList.itemData(index)
            self.load_texts_for_category(category_id)
            if self.textsList.model().rowCount() > 0:
                self.textsList.setCurrentIndex(self.textsList.model().index(0, 0))
                self.on_text_selected(self.textsList.currentIndex())

    def add_new_category(self):
        """Добавление новой категории"""
        text, ok = QInputDialog.getText(
            self,
            "Новая категория",
            "Введите название категории:",
            QLineEdit.EchoMode.Normal,
            ""
        )
        if ok and text:
            try:
                cat_id = self.db.add_category(text)
                self.catList.addItem(text, cat_id)
                self.catList.setCurrentText(self.catList.count() - 1)
            except Exception as e:
                self.statusbar.showMessage(f"Ошибка создания категории: {str(e)}", 5000)

    def setup_connections(self):
        """
        Установка соединений между элементами интерфейса
        """
        self.ValueSpeed.valueChanged.connect(self.update_speed_label)
        self.catList.currentIndexChanged.connect(self.on_category_changed)

        self.BtnPausePlay.clicked.connect(self.toggle_play_pause)
        self.BtnStop.clicked.connect(self.stop_playback)
        self.BtnPrevious.clicked.connect(self.previous_phrase)
        self.BtnNext.clicked.connect(self.next_phrase)

        self.newCat.clicked.connect(self.add_new_category)
        
        # Обработчик изменения текста
        self.textBrowser.textChanged.connect(self.update_button_states)
        
        # Подключение действий меню
        self.ActAbout.triggered.connect(self.show_about_dialog)

    def update_speed_label(self):
        """
        Обновление метки скорости воспроизведения
        """
        speed_value = self.ValueSpeed.value() / 10
        self.PrintValueSpeed.setText(f"{speed_value:.1f}")

        if self.speaker:
            try:
                self.speaker.Rate = int((speed_value - 1) * 10)
            except Exception as e:
                print(f"Ошибка: {e}")

    def split_text_into_sentences(self, text):
        """
        Разбиение текста на предложения с отслеживанием позиций
        """
        sentences = []
        positions = []
        current_sentence = ""
        current_start = 0
        
        for i, char in enumerate(text):
            current_sentence += char
            if char in '.!?,:;':
                sentences.append(current_sentence.strip())
                positions.append((current_start, i + 1))
                current_sentence = ""
                current_start = i + 1
        
        # Добавляем оставшийся текст, если он есть
        if current_sentence.strip():
            sentences.append(current_sentence.strip())
            positions.append((current_start, len(text)))
        
        # Убираем пустые строки
        filtered_sentences = []
        filtered_positions = []
        for sentence, pos in zip(sentences, positions):
            if sentence:
                filtered_sentences.append(sentence)
                filtered_positions.append(pos)
        
        return filtered_sentences, filtered_positions

    def highlight_current_sentence(self):
        """
        Выделение текущего предложения в тексте
        """
        if not self.sentences or self.current_sentence_index >= len(self.sentence_positions):
            return
            
        # Сначала убираем все выделения
        self.clear_highlights()
        
        # Выделяем текущее предложение
        start_pos, end_pos = self.sentence_positions[self.current_sentence_index]
        
        cursor = self.textBrowser.textCursor()
        cursor.setPosition(start_pos)
        cursor.setPosition(end_pos, QTextCursor.MoveMode.KeepAnchor)
        
        # Применяем формат выделения
        cursor.mergeCharFormat(self.highlight_format)
        
        # Прокручиваем к выделенному тексту
        self.textBrowser.setTextCursor(cursor)
        self.textBrowser.ensureCursorVisible()

    def clear_highlights(self):
        """
        Убирает все выделения из текста
        """
        cursor = self.textBrowser.textCursor()
        cursor.select(QTextCursor.SelectionType.Document)
        cursor.mergeCharFormat(self.normal_format)
        
        # Возвращаем курсор в начало
        cursor.setPosition(0)
        self.textBrowser.setTextCursor(cursor)

    def toggle_play_pause(self):
        """
        Переключение воспроизведения/паузы
        """
        if not self.speaker:
            print("SAPI не инициализирован")
            return
        try:
            if not self.is_playing and not self.is_pause:
                self.start_playback()
            elif self.is_playing:
                self.pause_playback()
            elif self.is_pause:
                self.resume_playback()
        except Exception as e:
            print(f"Ошибка при воспроизведении: {e}")

    def start_playback(self):
        """
        Воспроизведение текста
        """
        try:
            text = self.textBrowser.toPlainText().strip()

            if not text:
                print("Нет текста для воспроизведения")
                return

            selected_voice = self.get_selected_voice()
            if not selected_voice:
                print("Голос не выбран")
                return

            # Разбиваем текст на предложения с позициями
            self.sentences, self.sentence_positions = self.split_text_into_sentences(text)
            self.current_sentence_index = 0
            self.current_text = text

            if not self.sentences:
                print("Нет предложений для воспроизведения")
                return

            self.speaker.Voice = selected_voice
            speed_value = self.ValueSpeed.value() / 10
            self.speaker.Rate = int((speed_value - 1) * 10)

            # Начинаем воспроизведение с первого предложения
            self.play_current_sentence()

            self.is_playing = True
            self.BtnPausePlay.setText("⏸️")
            
            # Запускаем таймер для отслеживания статуса воспроизведения
            self.playback_timer.start(100)  # Проверяем каждые 100мс
            
            # Обновляем состояние кнопок
            self.update_button_states()

        except Exception as e:
            print(f"Ошибка при воспроизведении {e}")

    def play_current_sentence(self):
        """
        Воспроизведение текущего предложения
        """
        if 0 <= self.current_sentence_index < len(self.sentences):
            sentence = self.sentences[self.current_sentence_index]
            self.speaker.Speak(sentence, 1)
            # Выделяем текущее предложение
            self.highlight_current_sentence()

    def check_playback_status(self):
        """
        Проверка статуса воспроизведения и переход к следующему предложению
        """
        if self.is_playing and not self.is_pause:
            try:
                # Проверяем, завершилось ли воспроизведение текущего предложения
                status = self.speaker.Status
                if hasattr(status, 'RunningState') and status.RunningState == 1:  # 1 = не воспроизводится
                    # Переходим к следующему предложению
                    self.current_sentence_index += 1
                    
                    if self.current_sentence_index < len(self.sentences):
                        # Воспроизводим следующее предложение
                        self.play_current_sentence()
                    else:
                        # Воспроизведение завершено
                        self.stop_playback()
                        
            except Exception as e:
                print(f"Ошибка при проверке статуса: {e}")

    def pause_playback(self):
        """
        Пауза при воспроизведении
        """
        try:
            if self.speaker and self.is_playing:
                self.speaker.Speak("", 3)  # Останавливаем текущее воспроизведение
                self.is_playing = False
                self.is_pause = True
                self.BtnPausePlay.setText("▶️")
                # Обновляем состояние кнопок
                self.update_button_states()
        except Exception as e:
            print(f"Ошибка паузы: {e}")

    def resume_playback(self):
        """
        Возобновление воспроизведения с того места, где остановились
        """
        try:
            if self.speaker and self.is_pause:
                if self.current_sentence_index < len(self.sentences):
                    # Возобновляем с текущего предложения
                    self.play_current_sentence()
                    self.is_playing = True
                    self.is_pause = False
                    self.BtnPausePlay.setText("⏸️")
                else:
                    # Если дошли до конца, начинаем сначала
                    self.current_sentence_index = 0
                    self.play_current_sentence()
                    self.is_playing = True
                    self.is_pause = False
                    self.BtnPausePlay.setText("⏸️")
                # Обновляем состояние кнопок
                self.update_button_states()
        except Exception as e:
            print(f"Ошибка возобновления воспроизведения: {e}")

    def stop_playback(self):
        """
        Остановка воспроизведения
        """
        try:
            if self.speaker:
                self.speaker.Speak("", 3)
                self.is_playing = False
                self.is_pause = False
                self.current_sentence_index = 0
                self.BtnPausePlay.setText("⏯️")
                self.playback_timer.stop()
                # Убираем выделение
                self.clear_highlights()
                # Обновляем состояние кнопок
                self.update_button_states()
        except Exception as e:
            print(f"Ошибка при остановке воспроизведения {e}")

    def previous_phrase(self):
        """
        Переход к предыдущей фразе
        """
        if not self.sentences or self.current_sentence_index <= 0:
            return
            
        # Останавливаем текущее воспроизведение
        if self.speaker and self.is_playing:
            self.speaker.Speak("", 3)
        
        # Переходим к предыдущему предложению
        self.current_sentence_index -= 1
        
        # Выделяем новое предложение
        self.highlight_current_sentence()
        
        # Если воспроизведение было активно, продолжаем с нового предложения
        if self.is_playing:
            self.play_current_sentence()
        elif self.is_pause:
            # Если была пауза, остаемся в состоянии паузы
            self.is_playing = False
            self.is_pause = True
        
        # Обновляем состояние кнопок
        self.update_button_states()

    def next_phrase(self):
        """
        Переход к следующей фразе
        """
        if not self.sentences or self.current_sentence_index >= len(self.sentences) - 1:
            return
            
        # Останавливаем текущее воспроизведение
        if self.speaker and self.is_playing:
            self.speaker.Speak("", 3)
        
        # Переходим к следующему предложению
        self.current_sentence_index += 1
        
        # Выделяем новое предложение
        self.highlight_current_sentence()
        
        # Если воспроизведение было активно, продолжаем с нового предложения
        if self.is_playing:
            self.play_current_sentence()
        elif self.is_pause:
            # Если была пауза, остаемся в состоянии паузы
            self.is_playing = False
            self.is_pause = True
        
        # Обновляем состояние кнопок
        self.update_button_states()

    def show_about_dialog(self):
        """
        Показывает диалог "О программе"
        """
        about_dialog = AboutDialog(self)
        about_dialog.exec()

    def load_categories(self):
        """Загрузка категорий из базы данных"""
        try:
            categories = self.db.get_all_categories()
            self.catList.clear()
            for cat_id, name in categories:
                self.catList.addItem(name, cat_id)
            if categories:
                self.load_texts_for_category(categories[0][0])
                # Выбираем последний элемент ("Новый текст") если список пуст
                if self.textsList.model().rowCount() > 0:
                    self.textsList.setCurrentIndex(self.textsList.model().index(0, 0))
                    self.on_text_selected(self.textsList.currentIndex())
        except Exception as e:
            self.statusbar.showMessage(f"Ошибка загрузки категорий: {str(e)}", 5000)

    def load_texts_for_category(self, category_id):
        """Загрузка текстов для выбранной категории"""
        try:
            texts = self.db.get_texts_by_category(category_id)
            model = QStandardItemModel()
            # Добавляем существующие тексты
            for text_id, _, title, _ in texts:
                item = QStandardItem(title)
                item.setData(text_id, Qt.ItemDataRole.UserRole)
                item.setEditable(False)
                model.appendRow(item)
            
            # Добавляем специальный элемент для создания нового текста
            new_item = QStandardItem("🖊️ Новый текст")
            new_item.setData(-1, Qt.ItemDataRole.UserRole)
            new_item.setForeground(QColor(0, 255, 255))  # Голубой цвет
            model.appendRow(new_item)

            self.textsList.setModel(model)
            self.textsList.clicked.connect(self.on_text_selected)
        except Exception as e:
            self.statusbar.showMessage(f"Ошибка загрузки текстов: {str(e)}", 5000)

    def on_text_selected(self, index):
        """Обработчик выбора текста в списке"""
        try:
            model = self.textsList.model()
            text_id = model.data(index, Qt.ItemDataRole.UserRole)

            if text_id == -1:
                text, ok = QInputDialog(
                    self,
                    "Новый текст",
                    "Введите название текста:",
                    QLineEdit.EchoMode.Normal,
                    ""
                )
                if ok and text:
                    try:
                        # Получаем текущую категорию
                        cat_index = self.catList.currentIndex()
                        category_id = self.catList.itemData(cat_index)

                        # Создаём новый текст в БД
                        new_id = self.db.save_text(category_id, text_id, "")
                        self.current_text_id = new_id
                        # Обновляем список текстов
                        self.load_texts_for_category(category_id)
                        # Выбираем новый текст в списке
                        self.textsList.setCurrentIndex(model.index(model.rowCount() - 2, 0))
                        self.textBrowser.setFocus()
                    except Exception as e:
                        self.statusbar.showMessage(f"Ошибка создания текста: {str(e)}", 5000)
                return

            text_content = self.db.get_text_content(text_id)
            self.current_text_id = text_id  # Сохраняем ID текущего текста
            self.textBrowser.setPlainText(text_content)
            self.textBrowser.setFocus()
        except Exception as e:
            self.statusbar.showMessage(f"Ошибка загрузки текста: {str(e)}", 5000)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
