# main.py
import os
import threading
import time
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.gridlayout import GridLayout
from kivy.utils import platform
from kivy.metrics import dp
from kivy.clock import Clock
from kivy.core.window import Window

Window.size = (360, 640)

IS_ANDROID = (platform == 'android')

if IS_ANDROID:
    from jnius import autoclass
    TextToSpeech = autoclass('android.speech.tts.TextToSpeech')
    Locale = autoclass('java.util.Locale')
    PythonActivity = autoclass('org.kivy.android.PythonActivity')
    tts = TextToSpeech(PythonActivity.mActivity, None)
    tts.setLanguage(Locale.US)
    voices = tts.getVoices()
    male_voice = None
    female_voice = None
    for v in voices.toArray():
        name = v.getName()
        if "Голос 2" in name or "Voice 2" in name:
            male_voice = v
        elif "Голос 1" in name or "Voice 1" in name:
            female_voice = v
else:
    import win32com.client
    speaker = win32com.client.Dispatch("SAPI.SpVoice")


class PhraseApp(App):
    def build(self):
        self.current_voice = "female"
        self.phrases = []
        self.index = 0
        self.repeat = 0
        self.playing = False
        self.paused = False
        self.stop_flag = True
        self.thread = None
        self.phrase_buttons = []
        self.lock = threading.Lock()
        self.locked_theme = False

        root = BoxLayout(orientation='vertical', spacing=dp(1), padding=dp(2))

        voice_top = BoxLayout(size_hint_y=0.07, spacing=dp(4))
        voice_top.add_widget(Label(text="Голос:", size_hint_x=0.2, font_size='11sp'))
        bf = Button(text="Женский", font_size='10sp')
        bf.bind(on_press=lambda x: self.set_voice("female"))
        voice_top.add_widget(bf)
        bm = Button(text="Мужской", font_size='10sp')
        bm.bind(on_press=lambda x: self.set_voice("male"))
        voice_top.add_widget(bm)
        root.add_widget(voice_top)

        top = BoxLayout(size_hint_y=0.42, spacing=dp(1))

        themes_box = BoxLayout(orientation='vertical', size_hint_x=0.35)
        themes_box.add_widget(Label(text="Темы:", size_hint_y=0.12, font_size='10sp'))
        self.themes_layout = GridLayout(cols=1, size_hint_y=None, spacing=dp(1))
        self.themes_layout.bind(minimum_height=self.themes_layout.setter('height'))
        ts = ScrollView()
        ts.add_widget(self.themes_layout)
        themes_box.add_widget(ts)
        top.add_widget(themes_box)

        phrases_box = BoxLayout(orientation='vertical', size_hint_x=0.65)
        phrases_box.add_widget(Label(text="Фразы:", size_hint_y=0.12, font_size='10sp'))
        self.phrases_layout = GridLayout(cols=1, size_hint_y=None, spacing=dp(1))
        self.phrases_layout.bind(minimum_height=self.phrases_layout.setter('height'))
        ps = ScrollView()
        ps.add_widget(self.phrases_layout)
        phrases_box.add_widget(ps)
        top.add_widget(phrases_box)

        root.add_widget(top)

        root.add_widget(Label(text="Перевод:", size_hint_y=0.05, font_size='10sp'))
        self.eng_label = Label(text="", color=(1, 0, 0, 1), size_hint_y=0.14, font_size='20sp')
        root.add_widget(self.eng_label)
        self.rus_label = Label(text="", color=(0, 1, 0, 1), size_hint_y=0.09, font_size='13sp')
        root.add_widget(self.rus_label)

        controls = BoxLayout(size_hint_y=0.1, spacing=dp(1))
        b_pause = Button(text="Пауза", font_size='10sp')
        b_pause.bind(on_press=self.pause_play)
        b_resume = Button(text="Продолжить", font_size='10sp')
        b_resume.bind(on_press=self.resume_play)
        b_stop = Button(text="Стоп", font_size='10sp')
        b_stop.bind(on_press=self.stop_play)
        controls.add_widget(b_pause)
        controls.add_widget(b_resume)
        controls.add_widget(b_stop)
        root.add_widget(controls)

        self.status_label = Label(text="Выберите тему", size_hint_y=0.05, font_size='10sp')
        root.add_widget(self.status_label)

        self.load_phrases()
        return root

    def load_phrases(self):
        themes = {}
        path = os.path.join(os.path.dirname(__file__), "phrases.txt")
        with open(path, "r", encoding="utf-8-sig") as f:
            for line in f:
                parts = line.strip().split("|")
                if len(parts) == 3:
                    theme, rus, eng = parts
                    theme = theme.strip()
                    if theme not in themes:
                        themes[theme] = []
                    pair = (rus, eng)
                    if pair not in themes[theme]:
                        themes[theme].append(pair)
        self.themes = themes
        for theme in themes:
            btn = Button(text=theme, size_hint_y=None, height=dp(26), font_size='10sp')
            btn.bind(on_press=lambda x, t=theme: self.select_theme(t))
            self.themes_layout.add_widget(btn)

    def select_theme(self, theme):
        if self.locked_theme:
            Clock.schedule_once(lambda dt: setattr(self.status_label, 'text', 'Сначала Стоп'))
            return

        self.locked_theme = True

        self.phrases = self.themes[theme]
        self.index = 0
        self.repeat = 0

        self.phrases_layout.clear_widgets()
        self.phrase_buttons = []
        for i, (rus, eng) in enumerate(self.phrases):
            btn = Button(text=rus, size_hint_y=None, height=dp(26), font_size='10sp')
            btn.background_color = (0.3, 0.3, 0.3, 1)
            btn.bind(on_press=lambda x, r=rus, e=eng: self.speak_once(r, e))
            self.phrases_layout.add_widget(btn)
            self.phrase_buttons.append(btn)

        self.playing = True
        self.paused = False
        self.stop_flag = False
        Clock.schedule_once(lambda dt: setattr(self.status_label, 'text', 'Читается'))
        self.thread = threading.Thread(target=self._play_thread, daemon=True)
        self.thread.start()

    def _highlight(self, idx):
        for i, btn in enumerate(self.phrase_buttons):
            if i == idx:
                btn.background_color = (0.2, 0.6, 0.2, 1)
            else:
                btn.background_color = (0.3, 0.3, 0.3, 1)

    def speak_once(self, rus, eng):
        Clock.schedule_once(lambda dt: setattr(self.eng_label, 'text', eng))
        Clock.schedule_once(lambda dt: setattr(self.rus_label, 'text', rus))
        Clock.schedule_once(lambda dt, t=eng: self._speak_sync(t))

    def _speak_sync(self, text):
        if IS_ANDROID:
            tts.speak(text, TextToSpeech.QUEUE_FLUSH, None, None)
        else:
            try:
                speaker.Speak(text)
            except Exception as e:
                print("Ошибка озвучки:", e)

    def _play_thread(self):
        while True:
            with self.lock:
                if self.stop_flag or not self.playing:
                    return
                if not self.paused:
                    if self.index >= len(self.phrases):
                        self.index = 0
                    rus, eng = self.phrases[self.index]
                    cur = self.index
                    Clock.schedule_once(lambda dt, e=eng: setattr(self.eng_label, 'text', e))
                    Clock.schedule_once(lambda dt, r=rus: setattr(self.rus_label, 'text', r))
                    Clock.schedule_once(lambda dt, i=cur: self._highlight(i))
                    Clock.schedule_once(lambda dt, t=eng: self._speak_sync(t))
                    time.sleep(len(eng) * 0.08 + 0.5)
                    self.repeat += 1
                    if self.repeat >= 3:
                        self.repeat = 0
                        self.index += 1
            if self.paused:
                time.sleep(0.1)
                continue
            time.sleep(0.05)

    def pause_play(self, instance):
        with self.lock:
            if self.playing and not self.paused:
                self.paused = True
        Clock.schedule_once(lambda dt: setattr(self.status_label, 'text', 'Пауза'))

    def resume_play(self, instance):
        with self.lock:
            if self.paused:
                self.paused = False
        Clock.schedule_once(lambda dt: setattr(self.status_label, 'text', 'Читается'))

    def stop_play(self, instance):
        with self.lock:
            self.playing = False
            self.paused = False
            self.stop_flag = True
        self.index = 0
        self.repeat = 0
        self.locked_theme = False
        Clock.schedule_once(lambda dt: setattr(self.status_label, 'text', 'Стоп. Выберите тему'))

    def set_voice(self, voice):
        self.current_voice = voice
        if IS_ANDROID:
            try:
                if voice == "male" and male_voice:
                    tts.setVoice(male_voice)
                elif voice == "female" and female_voice:
                    tts.setVoice(female_voice)
            except Exception as e:
                print("Ошибка смены голоса:", e)
        else:
            try:
                vs = speaker.GetVoices()
                found = None
                # Мужской: David, Mark, Male, Guy
                # Женский: Zira, Anna, Irina, Female
                male_keys = ["david", "mark", "male", "guy"]
                female_keys = ["zira", "anna", "irina", "female"]
                keys = male_keys if voice == "male" else female_keys
                for i in range(vs.Count):
                    desc = vs.Item(i).GetDescription().lower()
                    for k in keys:
                        if k in desc:
                            found = vs.Item(i)
                            break
                    if found:
                        break
                if found:
                    speaker.Voice = found
                else:
                    # Если не нашли — берём по индексу
                    if voice == "male" and vs.Count > 0:
                        speaker.Voice = vs.Item(0)
                    elif voice == "female" and vs.Count > 1:
                        speaker.Voice = vs.Item(1)
            except Exception as e:
                print("Ошибка смены голоса:", e)


if __name__ == "__main__":
    PhraseApp().run()