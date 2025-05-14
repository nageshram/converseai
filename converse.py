import tkinter as tk
from tkinter import ttk, scrolledtext
import speech_recognition as sr
import pyttsx3
import google.generativeai as genai
import threading
import queue
import platform
import subprocess
import time
import sys
import os
import numpy as np
import random

# Additional libraries for wake word detection
import pvporcupine
import pyaudio

from screen_brightness_control import get_brightness, set_brightness
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import warnings

warnings.filterwarnings("ignore")

# ------------------------ Configuration ------------------------
CONFIG = {
    'history_file': "conversation_history.txt",
    'colors': {
        'user_bg': "#1abc9c",
        'ai_bg': "#2ecc71",
        'error_bg': "#e74c3c",
        'text_light': "#ecf0f1",
        'accent': "#3498db",
        'bg': "#2c3e50",
        'frame_bg': "#34495e"
    },
    'max_phrase_time': 20,
    'pause_threshold': 1,
    'dynamic_energy': True,
    'ambient_noise_duration': 3
}

# ------------------------ Keys and Wake Word Settings ------------------------
PORCUPINE_ACCESS_KEY = "KEY"
GEMINI_API_KEY = "APIKEY"
# Path to custom .ppn file (this file recognizes the spoken word "Jarvis")
CUSTOM_WAKE_WORD_PATH =r"N:/ConverseAi/purple_en_windows_v3_0_0.ppn"

WAKE_DURATION = 180  # Assistant stays awake for 3 minutes after wake word is detected

# List of stop commands
STOP_COMMANDS = ["stop", "quiet", "silent", "stay silent", "stop speaking", "stay quiet"]

# ------------------------ VoiceAssistant Class ------------------------
class VoiceAssistant:
    def __init__(self, root):
        self.root = root
        self.setup_styles()
        self.setup_ui()
        self.setup_audio()
        self.setup_ai()
        self.setup_hardware_controls()

        # Core variables
        self.running = True
        self.processing = False
        self.is_speaking = False
        self.stop_tts_flag = False
        self.conversation_queue = queue.Queue()
        self.audio_lock = threading.Lock()
        self.awake = False
        self.awake_start_time = None

        # Start threads for wake word detection and command listening
        self.wake_word_thread = threading.Thread(target=self.wake_word_detector, daemon=True)
        self.wake_word_thread.start()

        self.command_thread = threading.Thread(target=self.command_listener, daemon=True)
        self.command_thread.start()

        # Update UI status (includes wave animation)
        self.update_status()

        # Process the conversation queue
        self.root.after(100, self.process_queue)

        self.speak("Hello! Please say 'Purple' to wake me up and give your command.")

    # ------------------------ UI & Styles ------------------------
    def setup_styles(self):
        style = ttk.Style(self.root)
        style.theme_use('clam')
        style.configure("TFrame", background=CONFIG['colors']['frame_bg'])
        style.configure("TLabel", background=CONFIG['colors']['frame_bg'], foreground=CONFIG['colors']['text_light'], font=("Segoe UI", 10))
        style.configure("TButton", background=CONFIG['colors']['accent'], foreground=CONFIG['colors']['text_light'], font=("Segoe UI", 10, "bold"))
        style.map("TButton",
                  foreground=[('active', CONFIG['colors']['text_light'])],
                  background=[('active', CONFIG['colors']['accent'])])
        style.configure("Custom.Horizontal.TProgressbar", troughcolor=CONFIG['colors']['bg'], background=CONFIG['colors']['accent'])

    def setup_ui(self):
        self.root.title("ConverseAI")
        self.root.geometry("850x650")
        self.root.configure(bg=CONFIG['colors']['bg'])

        # Main chat frame (canvas for a rounded effect)
        self.chat_frame_canvas = tk.Canvas(self.root, bg=CONFIG['colors']['bg'], highlightthickness=0)
        self.chat_frame_canvas.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        self.chat_frame = ttk.Frame(self.chat_frame_canvas, style="TFrame")
        self.chat_frame_canvas.create_window((0, 0), window=self.chat_frame, anchor="nw", width=800, height=500)

        # Chat area
        self.chat_area = scrolledtext.ScrolledText(
            self.chat_frame,
            wrap=tk.WORD,
            bg=CONFIG['colors']['bg'],
            fg=CONFIG['colors']['text_light'],
            font=("Segoe UI", 12),
            relief=tk.FLAT, bd=0, padx=10, pady=10
        )
        self.chat_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.chat_area.configure(state=tk.DISABLED)
        self.configure_chat_tags()

        # Status frame
        self.status_frame = ttk.Frame(self.root, style="TFrame")
        self.status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(0,20))
        self.status_indicator = ttk.Label(self.status_frame, text="Status: Waiting for wake word", style="TLabel")
        self.status_indicator.pack(side=tk.LEFT, padx=10)

        # Wave animation canvas (for listening status)
        self.wave_canvas = tk.Canvas(self.status_frame, width=150, height=30, bg=CONFIG['colors']['frame_bg'], highlightthickness=0)
        self.wave_canvas.pack(side=tk.RIGHT, padx=10)

        self.progress = ttk.Progressbar(
            self.status_frame,
            style="Custom.Horizontal.TProgressbar",
            orient='horizontal',
            mode='indeterminate',
            length=150
        )
        self.progress.pack(side=tk.RIGHT, padx=10)
        self.progress.pack_forget()

    def configure_chat_tags(self):
        self.chat_area.tag_config('user',
                                  background=CONFIG['colors']['user_bg'],
                                  foreground=CONFIG['colors']['text_light'],
                                  font=("Segoe UI", 12, "bold"),
                                  spacing3=10, lmargin1=5, lmargin2=5, rmargin=5)
        self.chat_area.tag_config('ai',
                                  background=CONFIG['colors']['ai_bg'],
                                  foreground=CONFIG['colors']['text_light'],
                                  font=("Segoe UI", 12),
                                  spacing3=10, lmargin1=5, lmargin2=5, rmargin=5)
        self.chat_area.tag_config('error',
                                  background=CONFIG['colors']['error_bg'],
                                  foreground="yellow",
                                  font=("Segoe UI", 10, "italic"),
                                  spacing3=10, lmargin1=5, lmargin2=5, rmargin=5)

    # ------------------------ Audio & TTS Setup ------------------------
    def setup_audio(self):
        self.tts_engine = pyttsx3.init()
        self.configure_tts()
        self.recognizer = sr.Recognizer()
        self.recognizer.dynamic_energy_threshold = CONFIG['dynamic_energy']
        self.recognizer.pause_threshold = CONFIG['pause_threshold']
        self.recognizer.energy_threshold = 300
        self.mic_device_index = 0

    def configure_tts(self):
        voices = self.tts_engine.getProperty('voices')
        if platform.system() == 'Windows' and len(voices) > 1:
            self.tts_engine.setProperty('voice', voices[1].id)
        self.tts_engine.setProperty('rate', 200)

    # ------------------------ Gemini AI Setup ------------------------
    def setup_ai(self):
        genai.configure(api_key=GEMINI_API_KEY)
        self.model = genai.GenerativeModel(
            "gemini-2.0-flash",
            system_instruction="""Your name is Purple, a chatbot designed for a college mini project assigned by Professor Reshma. You are being developed by Mueen and Nagesha, who are MCA students at BMSIT, Bangalore.

Your goal is to converse like a human rather than an AI. Follow these guidelines:

1. Natural Conversation - Respond casually and naturally, like a human would.

2. Concise Answers - For big questions, provide a simple algorithm and a link for more details instead of a lengthy response.

3. Code Snippets Only - When asked for code, give only the necessary snippets, not the entire code.

4. Stay in Character - You are Purple, not an AI. Maintain""",
            generation_config={
                'temperature': 1,
                'top_p': 0.95,
                'top_k': 40,
                'max_output_tokens': 8192,
            }
        )

    # ------------------------ Hardware Controls ------------------------
    def setup_hardware_controls(self):
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume_control = cast(interface, POINTER(IAudioEndpointVolume))
        try:
            get_brightness()
            self.brightness_control = True
        except Exception:
            self.brightness_control = False

    # ------------------------ Wake Word Detection ------------------------
    def wake_word_detector(self):
        """ Uses Porcupine to detect the custom wake word from a .ppn file. """
        try:
            porcupine = pvporcupine.create(
                access_key=PORCUPINE_ACCESS_KEY,
                keyword_paths=[CUSTOM_WAKE_WORD_PATH]
            )
        except Exception:
            self.show_error_ui("Wake word detector initialization failed.")
            return

        pa = pyaudio.PyAudio()
        stream = pa.open(
            rate=porcupine.sample_rate,
            channels=1,
            format=pyaudio.paInt16,
            input=True,
            frames_per_buffer=porcupine.frame_length
        )

        while self.running:
            try:
                pcm = stream.read(porcupine.frame_length, exception_on_overflow=False)
                pcm = np.frombuffer(pcm, dtype=np.int16)
                result = porcupine.process(pcm)
                if result >= 0:
                    # Wake word detected ("Purple")
                    self.awake = True
                    self.awake_start_time = time.time()
                    # Signal command mode by sending an empty command
                    self.conversation_queue.put(("command", ""))
                    time.sleep(1)  # Avoid repeated detections
            except Exception:
                continue

        stream.stop_stream()
        stream.close()
        pa.terminate()
        porcupine.delete()

    # ------------------------ Command Listener ------------------------
    def command_listener(self):
        """ Listens for user commands after wake word detection using speech_recognition. """
        while self.running:
            if self.awake:
                with sr.Microphone(device_index=self.mic_device_index) as source:
                    try:
                        self.recognizer.adjust_for_ambient_noise(source, duration=CONFIG['ambient_noise_duration'])
                        audio = self.recognizer.listen(source, timeout=0, phrase_time_limit=CONFIG['max_phrase_time'])
                        command = self.recognizer.recognize_google(audio).lower().strip()
                        self.awake_start_time = time.time()  # Reset awake timer
                        self.conversation_queue.put(("command", command))
                    except sr.WaitTimeoutError:
                        continue
                    except sr.UnknownValueError:
                        self.show_error_ui("Could not understand command audio.")
                        continue
                    except sr.RequestError as e:
                        self.show_error_ui(f"Speech recognition error: {e}")
                        continue
                    except Exception:
                        continue
            else:
                time.sleep(0.1)
            if self.awake and (time.time() - self.awake_start_time > WAKE_DURATION):
                self.awake = False

    # ------------------------ UI Status & Wave Animation ------------------------
    def update_status(self):
        """ Updates the status label and runs a simple wave animation when awake. """
        if self.awake:
            self.status_indicator.config(text="Status: Awake - Listening for 'purple' command", foreground="#2ecc71")
            self.update_wave_animation()
        else:
            self.status_indicator.config(text="Status: Waiting for wake word ( say purple )", foreground=CONFIG['colors']['text_light'])
            self.wave_canvas.delete("all")
        self.root.after(1000, self.update_status)

    def update_wave_animation(self):
        """ Draws random bars on the wave canvas to simulate a listening animation. """
        self.wave_canvas.delete("all")
        width = self.wave_canvas.winfo_width()
        height = self.wave_canvas.winfo_height()
        num_bars = 10
        bar_width = width // num_bars
        for i in range(num_bars):
            # Simulate wave: random bar heights
            bar_height = random.randint(5, height)
            x0 = i * bar_width
            y0 = height - bar_height
            x1 = x0 + bar_width - 2
            y1 = height
            self.wave_canvas.create_rectangle(x0, y0, x1, y1, fill=CONFIG['colors']['accent'], width=0)

    # ------------------------ UI Helpers ------------------------
    def show_listening_ui(self):
        self.status_indicator.config(text="Listening...", foreground="#2ecc71")
        self.progress.pack(pady=5)
        self.progress.start()

    def hide_listening_ui(self):
        if hasattr(self, 'progress'):
            self.progress.stop()
            self.progress.pack_forget()
        self.status_indicator.config(text="Ready", foreground=CONFIG['colors']['accent'])

    def show_error_ui(self, message):
        self.chat_area.configure(state=tk.NORMAL)
        self.chat_area.insert(tk.END, f"System: {message}\n", 'error')
        self.chat_area.configure(state=tk.DISABLED)
        self.chat_area.see(tk.END)

    # ------------------------ Chat & Command Processing ------------------------
    def add_to_chat(self, sender, message, tag):
        self.chat_area.configure(state=tk.NORMAL)
        self.chat_area.insert(tk.END, f"{sender}: {message}\n", tag)
        self.chat_area.insert(tk.END, "\n")
        self.save_to_history(f"{sender}: {message}")
        self.chat_area.configure(state=tk.DISABLED)
        self.chat_area.see(tk.END)

    def save_to_history(self, entry):
        with open(CONFIG['history_file'], 'a') as f:
            f.write(f"{time.ctime()} - {entry}\n")

    def process_queue(self):
        while not self.conversation_queue.empty():
            msg_type, content = self.conversation_queue.get()
            if msg_type == "command":
                self.handle_command(content)
            elif msg_type == "error":
                self.handle_error(content)
        self.root.after(100, self.process_queue)

    def handle_command(self, command):
        try:
            if command:
                self.add_to_chat("You", command, "user")

            # Check for stop commands
            if any(stop_word in command for stop_word in STOP_COMMANDS):
                self.stop_speaking()
                return

            # Quit command
            if 'quit' in command or 'exit' in command:
                self.speak("Goodbye! Have a nice day.")
                self.graceful_exit()
                return

            # Application control
            if 'open' in command:
                app_name = command.split('open ')[-1]
                self.open_application(app_name)
                return
            if 'close' in command:
                app_name = command.split('close ')[-1]
                self.close_application(app_name)
                return

            # Volume/Brightness control
            if 'volume' in command:
                self.handle_volume(command)
                return
            if 'brightness' in command:
                self.handle_brightness(command)
                return

            # Default: Use Gemini AI for response
            self.processing = True
            response = self.model.generate_content(command)
            ai_text = self.clean_text(response.text)
            self.add_to_chat("AI", ai_text, "ai")
            self.speak(ai_text)
        except Exception as e:
            #elf.add_to_chat("System", f"Error: {str(e)}", "error")
            print(e)
        finally:
            self.processing = False

    def clean_text(self, text):
        return text.replace('*', '').replace('_', '')

    # ------------------------ Application Commands ------------------------
    def open_application(self, app_name):
        apps = {
            'notepad': 'notepad.exe',
            'calculator': 'calc.exe',
            'chrome': 'chrome.exe',
            'word': 'winword.exe',
            'excel': 'excel.exe'
        }
        try:
            if app_name in apps:
                os.startfile(apps[app_name])
                self.speak(f"Opening {app_name}")
                self.add_to_chat("System", f"Opened {app_name}", "ai")
            else:
                self.speak(f"Application {app_name} not configured")
        except Exception:
            self.speak(f"Failed to open {app_name}")

    def close_application(self, app_name):
        apps = {
            'notepad': 'notepad.exe',
            'calculator': 'Calculator.exe',
            'chrome': 'chrome.exe',
            'word': 'WINWORD.EXE',
            'excel': 'EXCEL.EXE'
        }
        try:
            if app_name in apps:
                if platform.system() == 'Windows':
                    os.system(f"taskkill /f /im {apps[app_name]}")
                    self.speak(f"Closing {app_name}")
                    self.add_to_chat("System", f"Closed {app_name}", "ai")
                else:
                    self.speak("Application closing not supported on this OS")
            else:
                self.speak(f"Application {app_name} not configured")
        except Exception:
            self.speak(f"Failed to close {app_name}")

    def handle_volume(self, command):
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume_control = cast(interface, POINTER(IAudioEndpointVolume))
            current_volume = volume_control.GetMasterVolumeLevelScalar()
            if 'mute' in command:
                volume_control.SetMute(1, None)
                self.speak("Volume muted")
                return
            if 'unmute' in command:
                volume_control.SetMute(0, None)
                self.speak("Volume unmuted")
                return
            if 'max' in command or 'full' in command:
                volume = 1.0
            elif 'min' in command or 'zero' in command:
                volume = 0.0
            else:
                volume = [int(s) for s in command.split() if s.isdigit()]
                volume = volume[0] / 100 if volume else current_volume
            volume_control.SetMasterVolumeLevelScalar(volume, None)
            self.speak(f"Volume set to {int(volume * 100)} percent")
        except Exception:
            self.speak("Failed to adjust volume")

    def handle_brightness(self, command):
        try:
            if not self.brightness_control:
                self.speak("Brightness control not available")
                return
            current = get_brightness()[0]
            if 'max' in command:
                brightness = 100
            elif 'min' in command:
                brightness = 0
            else:
                brightness = [int(s) for s in command.split() if s.isdigit()]
                brightness = brightness[0] if brightness else current
            set_brightness(brightness)
            self.speak(f"Brightness set to {brightness} percent")
        except Exception:
            self.speak("Failed to adjust brightness")

    def handle_error(self, error):
        self.add_to_chat("System", f"Error: {error}", "error")

    # ------------------------ Text-to-Speech ------------------------
    def speak(self, text):
        def _speak():
            with self.audio_lock:
                self.is_speaking = True
                max_chunk = 500
                chunks = [text[i:i + max_chunk] for i in range(0, len(text), max_chunk)]
                for chunk in chunks:
                    if self.stop_tts_flag:
                        break
                    self.tts_engine.say(chunk)
                    self.tts_engine.runAndWait()
                self.is_speaking = False
                self.stop_tts_flag = False  # Reset flag after finishing
        threading.Thread(target=_speak).start()

    def stop_speaking(self):
        """ Immediately stops text-to-speech and remains silent. """
        self.stop_tts_flag = True
        self.tts_engine.stop()
        self.add_to_chat("System", "Okay, I will remain silent.", "ai")

    # ------------------------ Graceful Exit ------------------------
    def graceful_exit(self):
        self.running = False
        self.tts_engine.stop()
        self.root.destroy()
        sys.exit()


# ------------------------ Main ------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = VoiceAssistant(root)
    root.protocol("WM_DELETE_WINDOW", app.graceful_exit)
    root.mainloop()
