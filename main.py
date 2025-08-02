import sys
import win32com.client
from PyQt6.QtWidgets import QApplication, QMainWindow
from PyQt6.QtCore import Qt
from ui.MainWindow import Ui_MainWindow


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.voice_list = []
        self.speaker = None
        self.is_playing = False
        self.is_pause = False
        self.pause_position = 0

        # Инициализация объектов
        self.setup_voices()
        self.setup_connections()
        self.update_speed_label()

        self.textBrowser.setAcceptRichText(False)

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
                self.plause_playback()
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

            self.speaker.Voice = selected_voice
            speed_value = self.ValueSpeed.value() / 10
            self.speaker.Rate = int((speed_value - 1) * 10)
            self.speaker.Speak(text, 1)

            self.is_playing = True

            self.BtnPausePlay.setText("⏸️")
        except Exception as e:
            print(f"Ошибка при воспроизведении {e}")

    def plause_playback(self):
        """
        Пауза при воспроизведении
        """
        try:
            if self.speaker and self.is_playing:
                self.speaker.Speak("", 3)
                self.pause_position = 0
                self.is_playing = False
                self.is_pause = True
                self.BtnPausePlay.setText("▶️")
        except Exception as e:
            print(f"Ошибка паузы: {e}")

    def resume_playback(self):
        """
        Возобновление воспроизведении
        """
        try:
            if self.speaker and self.is_pause:
                text = self.textBrowser.toPlainText().strip()

                if not text:
                    print("Нет текста")
                    return

                self.speaker.Speak(text, 1)
                self.is_playing = True
                self.is_pause = False
                self.BtnPausePlay.setText("⏸️")
        except Exception as e:
            print(f"Ошибка возобновлении воспроизведения: {e}")

    def stop_playback(self):
        """
        Остановка воспроизведения
        """
        try:
            if self.speaker and self.is_playing:
                self.speaker.Speak("", 3)
                self.is_playing = False
                self.BtnPausePlay.setText("⏯️")
        except Exception as e:
            print(f"Ошибка при остановки воспроизведения {e}")

    def previous_phrase(self):
        """
        Переход к предыдущей фразе
        """
        # TODO: Реализовать логику воспроизведения
        print('Назад нажата')

    def next_phrase(self):
        """
        Переход к следующей фразе
        """
        # TODO: Реализовать логику воспроизведения
        print('Далее нажата')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
