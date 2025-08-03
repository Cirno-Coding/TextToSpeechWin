import sys
import os
import win32com.client
from datetime import datetime
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QTextCursor, QTextCharFormat, QColor
from ui.MainWindow import Ui_MainWindow


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

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

        self.textBrowser.setAcceptRichText(False)

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

    def setup_connections(self):
        """
        Установка соединений между элементами интерфейса
        """
        self.ValueSpeed.valueChanged.connect(self.update_speed_label)

        self.BtnPausePlay.clicked.connect(self.toggle_play_pause)
        self.BtnStop.clicked.connect(self.stop_playback)
        self.BtnPrevious.clicked.connect(self.previous_phrase)
        self.BtnNext.clicked.connect(self.next_phrase)
        
        # Обработчик изменения текста
        self.textBrowser.textChanged.connect(self.update_button_states)
        
        # Подключение действий меню
        self.ActOpen.triggered.connect(self.open_file)
        self.ActSave.triggered.connect(self.save_file)
        self.ActExit.triggered.connect(self.exit_application)

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

    def save_file(self):
        """
        Сохранение текста в файл
        """
        try:
            text_content = self.textBrowser.toPlainText()
            if text_content:
                # Создаем папку texts, если её нет
                texts_dir = os.path.join(os.getcwd(), "texts")
                if not os.path.exists(texts_dir):
                    os.makedirs(texts_dir)

                # Генерируем имя файла по умолчанию
                current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                default_filename = f"Текст от {current_time}.txt"
                default_path = os.path.join(texts_dir, default_filename)

                # Открываем диалог сохранения
                file_path, _ = QFileDialog.getSaveFileName(
                    self,
                    "Сохранить текст",
                    default_path,
                    "Текстовые файлы (*.txt)"
                )

                if file_path:
                    # Сохраняем содержимое QTextEdit в файл
                    with open(file_path, 'w', encoding='utf-8') as file:
                        file.write(text_content)

                    # Очищаем QTextEdit
                    self.textBrowser.clear()

                    # Останавливаем воспроизведение, если оно активно
                    if self.is_playing or self.is_pause:
                        self.stop_playback()
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при сохранении файла: {e}")

    def open_file(self):
        """
        Открытие текстового файла
        """
        try:
            # Проверяем, есть ли текст в QTextEdit
            current_text = self.textBrowser.toPlainText().strip()
            
            if current_text:
                # Спрашиваем пользователя о сохранении текущего текста
                reply = QMessageBox.question(
                    self,
                    "Сохранить текст",
                    "Хотите сохранить текущий текст?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.Yes
                )
                
                if reply == QMessageBox.StandardButton.Yes:
                    # Сохраняем текущий текст
                    self.save_file()
                    # Если пользователь отменил сохранение, не открываем новый файл
                    if self.textBrowser.toPlainText().strip():
                        return
                else:
                    # Очищаем QTextEdit
                    self.textBrowser.clear()
                    # Останавливаем воспроизведение, если оно активно
                    if self.is_playing or self.is_pause:
                        self.stop_playback()
            
            # Открываем диалог выбора файла
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Открыть текстовый файл",
                os.path.join(os.getcwd(), "texts"),
                "Текстовые файлы (*.txt)"
            )
            
            if file_path:
                # Читаем содержимое файла
                with open(file_path, 'r', encoding='utf-8') as file:
                    text_content = file.read()
                
                # Загружаем текст в QTextEdit
                self.textBrowser.setPlainText(text_content)
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка при открытии файла: {e}")

    def exit_application(self):
        """
        Выход из приложения
        """
        # Останавливаем воспроизведение, если оно активно
        if self.is_playing or self.is_pause:
            self.stop_playback()
        
        # Закрываем приложение
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
