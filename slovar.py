import os
import time
import win32com.client
import tkinter as tk
from tkinter import messagebox

PHRASES_FILE = "phrases.txt"

is_paused = False
stop_flag = False
current_phrase_index = 0
saved_theme_phrases = []

def get_voice(gender):
    tts = win32com.client.Dispatch('SAPI.SpVoice')
    voices = tts.GetVoices()
    for i in range(voices.Count):
        desc = voices.Item(i).GetDescription()
        if gender == "male" and "Guy" in desc:
            return voices.Item(i)
        if gender == "female" and "Anna" in desc:
            return voices.Item(i)
    return None

def speak(text):
    tts = win32com.client.Dispatch('SAPI.SpVoice')
    voice = get_voice(voice_var.get())
    if voice:
        tts.Voice = voice
    tts.Rate = 0
    tts.Speak(text)

def load_phrases(filename):
    phrases = []
    with open(filename, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.strip()
            if not line or '|' not in line:
                continue
            parts = line.split('|')
            if len(parts) >= 3:
                theme, ru_text, en_text = parts[0].strip(), parts[1].strip(), parts[2].strip()
                theme = theme.replace('\ufeff', '').replace('\u200b', '').strip()
                phrases.append((theme, ru_text, en_text))
    return phrases

def read_next_phrase():
    global current_phrase_index, is_paused, stop_flag, saved_theme_phrases
    if stop_flag or is_paused:
        return
    if current_phrase_index >= len(saved_theme_phrases):
        status_label.config(text="Готово")
        return

    theme, ru, en = saved_theme_phrases[current_phrase_index]
    update_current_phrase(en, ru)
    root.update()
    for _ in range(3):
        if stop_flag or is_paused:
            return
        speak(en)
    current_phrase_index += 1
    root.after(500, read_next_phrase)

def update_current_phrase(en, ru):
    current_en_label.config(text=en)
    current_ru_label.config(text=ru)
    for i in range(phrase_listbox.size()):
        if phrase_listbox.get(i) == ru:
            phrase_listbox.selection_clear(0, tk.END)
            phrase_listbox.selection_set(i)
            phrase_listbox.see(i)
            break

def start_reading():
    global current_phrase_index, is_paused, stop_flag, saved_theme_phrases
    if not theme_listbox.curselection():
        messagebox.showwarning("Внимание", "Выбери тему.")
        return
    theme_name = theme_listbox.get(theme_listbox.curselection())
    saved_theme_phrases = [p for p in phrases if p[0] == theme_name]
    if not saved_theme_phrases:
        messagebox.showwarning("Внимание", "В этой теме нет фраз.")
        return
    current_phrase_index = 0
    is_paused = False
    stop_flag = False
    status_label.config(text="Чтение...")
    read_next_phrase()

def pause_reading():
    global is_paused
    is_paused = True
    status_label.config(text="Пауза")

def resume_reading():
    global is_paused
    if not saved_theme_phrases:
        return
    is_paused = False
    status_label.config(text="Чтение...")
    read_next_phrase()

def stop_reading():
    global stop_flag, is_paused, current_phrase_index, saved_theme_phrases
    stop_flag = True
    is_paused = False
    current_phrase_index = 0
    saved_theme_phrases = []
    status_label.config(text="Остановлено")

def on_theme_select(event):
    if not theme_listbox.curselection():
        return
    theme = theme_listbox.get(theme_listbox.curselection())
    phrase_listbox.delete(0, tk.END)
    english_label.config(text="")
    current_ru_label.config(text="")
    current_en_label.config(text="")
    for p in phrases:
        if p[0] == theme:
            phrase_listbox.insert(tk.END, p[1])

def on_phrase_select(event):
    if not phrase_listbox.curselection():
        return
    ru_text = phrase_listbox.get(phrase_listbox.curselection())
    for p in phrases:
        if p[1] == ru_text:
            english_label.config(text=p[2])
            break

def main():
    global root, theme_listbox, phrase_listbox, english_label, current_ru_label, current_en_label, status_label, voice_var, phrases

    if not os.path.exists(PHRASES_FILE):
        messagebox.showerror("Ошибка", f"Нет файла {PHRASES_FILE}")
        return

    phrases = load_phrases(PHRASES_FILE)
    if not phrases:
        messagebox.showerror("Ошибка", "Файл пуст.")
        return

    themes = {}
    for theme, ru, en in phrases:
        if theme not in themes:
            themes[theme] = []
        themes[theme].append(ru)

    root = tk.Tk()
    root.title("Словарь (оффлайн)")
    root.geometry("800x650")

    top_frame = tk.Frame(root)
    top_frame.pack(fill=tk.X, padx=10, pady=5)
    tk.Label(top_frame, text="Голос:").pack(side=tk.LEFT)
    voice_var = tk.StringVar(value="male")
    tk.Radiobutton(top_frame, text="Женский", variable=voice_var, value="female").pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(top_frame, text="Мужской", variable=voice_var, value="male").pack(side=tk.LEFT, padx=5)

    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    left_frame = tk.Frame(frame)
    left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    right_frame = tk.Frame(frame)
    right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

    tk.Label(left_frame, text="Темы:").pack(anchor=tk.W)
    theme_listbox = tk.Listbox(left_frame, height=20)
    theme_listbox.pack(fill=tk.BOTH, expand=True)
    for theme in themes.keys():
        theme_listbox.insert(tk.END, theme)
    theme_listbox.bind('<<ListboxSelect>>', on_theme_select)

    tk.Label(right_frame, text="Фразы:").pack(anchor=tk.W)
    phrase_listbox = tk.Listbox(right_frame, height=20)
    phrase_listbox.pack(fill=tk.BOTH, expand=True)
    phrase_listbox.bind('<<ListboxSelect>>', on_phrase_select)

    bottom_frame = tk.Frame(root)
    bottom_frame.pack(fill=tk.X, padx=10, pady=5)
    tk.Label(bottom_frame, text="Перевод:").pack(anchor=tk.W)
    english_label = tk.Label(bottom_frame, text="", font=("Arial", 14), fg="blue")
    english_label.pack(anchor=tk.W, pady=5)

    tk.Label(bottom_frame, text="Сейчас читается:").pack(anchor=tk.W)
    current_en_label = tk.Label(bottom_frame, text="", font=("Arial", 12), fg="red")
    current_en_label.pack(anchor=tk.W, pady=2)
    current_ru_label = tk.Label(bottom_frame, text="", font=("Arial", 12), fg="green")
    current_ru_label.pack(anchor=tk.W, pady=2)

    control_frame = tk.Frame(root)
    control_frame.pack(fill=tk.X, padx=10, pady=5)
    tk.Button(control_frame, text="Читать", command=start_reading).pack(side=tk.LEFT, padx=5)
    tk.Button(control_frame, text="Пауза", command=pause_reading).pack(side=tk.LEFT, padx=5)
    tk.Button(control_frame, text="Продолжить", command=resume_reading).pack(side=tk.LEFT, padx=5)
    tk.Button(control_frame, text="Стоп", command=stop_reading).pack(side=tk.LEFT, padx=5)

    status_label = tk.Label(root, text="Готов", fg="grey")
    status_label.pack(side=tk.BOTTOM, anchor=tk.W, padx=10)

    root.mainloop()

if __name__ == "__main__":
    main()