import pyttsx3
import datetime
import speech_recognition as sr
import webbrowser as wb
import os
import subprocess
import sounddevice as sd
import numpy as np
import scipy.io.wavfile as wav
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import cv2
from PIL import ImageGrab
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from urllib.parse import quote
import tkinter as tk
import threading
import queue
import csv
import google.generativeai as genai  # Import the Gemini API

class JarvisAssistant:
    def __init__(self, api_key):
        self.engine = pyttsx3.init()
        self.command_queue = queue.Queue()
        self.gui = None
        self.running = True
        
        # Initialize Gemini API
        self.configure_gemini(api_key)
        
        # Set up CSV logging
        self.log_dir = os.path.join(os.getcwd(), "speech_logs")
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        
        # Create a new log file with timestamp
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.log_file = os.path.join(self.log_dir, f"jarvis_speech_log_{timestamp}.csv")
        
        # Initialize CSV file with headers
        with open(self.log_file, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(['Timestamp', 'User Speech', 'Response', 'Status'])
        
        print(f"Speech log file created at: {self.log_file}")

    def configure_gemini(self, api_key):
        """Configure the Gemini API with the provided API key."""
        try:
            # Configure the API with the key
            genai.configure(api_key=api_key)

            # List available models to debug and ensure we're using the right one
            models = genai.list_models()
            gemini_models = [model.name for model in models if "gemini" in model.name.lower()]
            print("Available Gemini models:", gemini_models)

            # Choose the most appropriate Gemini model from the available ones
            # Updated to prefer newer Gemini models (1.5 series) instead of deprecated 1.0 Pro Vision
            preferred_models = [
                "gemini-1.5-pro",
                "gemini-1.5-flash",
                "gemini-1.5-turbo"
            ]

            # Try to find a preferred model first
            model_name = None
            for preferred in preferred_models:
                matching = [m for m in gemini_models if preferred in m.lower()]
                if matching:
                    model_name = matching[0]
                    break

            # If no preferred model found, use any available Gemini model
            if not model_name and gemini_models:
                model_name = gemini_models[0]

            if model_name:
                print(f"Using Gemini model: {model_name}")

                # Set up generation config
                generation_config = {
                    "temperature": 0.7,
                    "top_p": 0.95,
                    "top_k": 40,
                    "max_output_tokens": 1024,
                }

                # Initialize the model with the correct name
                self.model = genai.GenerativeModel(
                    model_name=model_name,
                    generation_config=generation_config
                )

                # Initialize chat 
                self.safety_settings = [
                    {
                        "category": "HARM_CATEGORY_HARASSMENT",
                        "threshold": "BLOCK_MEDIUM_AND_ABOVE"
                    },
                    {
                        "category": "HARM_CATEGORY_HATE_SPEECH",
                        "threshold": "BLOCK_MEDIUM_AND_ABOVE"
                    },
                    {
                        "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
                        "threshold": "BLOCK_MEDIUM_AND_ABOVE"
                    },
                    {
                        "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
                        "threshold": "BLOCK_MEDIUM_AND_ABOVE"
                    }
                ]

                try:
                    self.chat = self.model.start_chat(
                        history=[
                            {
                                "role": "user",
                                "parts": ["You are JARVIS, an AI assistant. You assist users with tasks on their computer and answer questions. Keep responses concise."]
                            },
                            {
                                "role": "model",
                                "parts": ["I am JARVIS, ready to assist you with computer tasks and questions. I'll keep my responses concise and helpful."]
                            }
                        ]
                    )
                    print("Chat session initialized")
                except Exception as e:
                    print(f"Could not initialize chat: {e}. Will use direct generation.")
                    self.chat = None

            else:
                print("No Gemini models found. Please check your API key and permissions.")
                self.model = None
                self.chat = None

        except Exception as e:
            print(f"Error configuring Gemini API: {e}")
            self.model = None
            self.chat = None

    def speak(self, audio):
        """Convert text to speech."""
        self.engine.say(audio)
        self.engine.runAndWait()

    def get_time(self):
        """Get the current time."""
        Time = datetime.datetime.now().strftime("%I:%M %p")
        self.speak("The current time is")
        self.speak(Time)
        print("The current time is", Time)
        return f"The current time is {Time}"

    def get_date(self):
        """Get the current date."""
        day = datetime.datetime.now().day
        month = datetime.datetime.now().month
        year = datetime.datetime.now().year
        self.speak("The current date is")
        self.speak(f"{day} {month} {year}")
        print(f"The current date is {day}/{month}/{year}")
        return f"The current date is {day}/{month}/{year}"

    def wishme(self):
        """Greet the user."""
        self.speak("Welcome back, sir!")
        hour = datetime.datetime.now().hour
        greeting = ""
        if 4 <= hour < 12:
            greeting = "Good Morning Sir!"
            self.speak(greeting)
        elif 12 <= hour < 16:
            greeting = "Good Afternoon Sir!"
            self.speak(greeting)
        elif 16 <= hour < 24:
            greeting = "Good Evening Sir!"
            self.speak(greeting)
        else:
            greeting = "Good Night Sir, See You Tomorrow"
            self.speak(greeting)
        
        # Different greeting based on Gemini API status
        if self.model:
            self.speak("Jarvis at your service. Please tell me how may I assist you.")
            status = "Running"
        else:
            self.speak("Jarvis at your service. Running in basic mode only.")
            status = "Initializing in basic mode"
        
        # Log the initial greeting
        self.log_to_csv("System Startup", greeting, status)

    def record_audio(self, duration=5, samplerate=44100):
        """Record audio using the sounddevice module."""
        self.speak("Listening...")
        audio_data = sd.rec(int(duration * samplerate), samplerate=samplerate, channels=1, dtype='int16')
        sd.wait()
        audio_data = np.int16(audio_data / np.max(np.abs(audio_data)) * 32767)
        return audio_data, samplerate

    def takecommand(self):
        """Listen to a voice command and convert it to text."""
        try:
            audio_data, samplerate = self.record_audio()
            audio_path = "temp_audio.wav"
            wav.write(audio_path, samplerate, audio_data)

            r = sr.Recognizer()
            with sr.AudioFile(audio_path) as source:
                audio = r.record(source)

            query = r.recognize_google(audio, language="en-in")
            print(f"Command: {query}")
            os.remove(audio_path)
            return query.lower()

        except sr.UnknownValueError:
            self.speak("Sorry, I didn't catch that. Please repeat.")
            return "none"
        except sr.RequestError:
            self.speak("Speech service is currently unavailable.")
            return "none"
        except Exception as e:
            print(f"Error: {e}")
            self.speak("There was an error. Please try again.")
            return "none"

    def set_gui(self, gui):
        self.gui = gui

    def log_to_csv(self, user_speech, response, status):
        """Log user speech and system response to CSV file."""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            with open(self.log_file, 'a', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow([timestamp, user_speech, response, status])
            print(f"Logged interaction to CSV: {user_speech}")
        except Exception as e:
            print(f"Error logging to CSV: {e}")

    def process_command(self, query):
        # Update GUI with the command
        if self.gui:
            self.gui.add_command_to_history(query)
            self.gui.update_status("Processing command...")
        
        response = "Command processed"
        status = "Success"
            
        try:
            # Check for direct system commands first
            if "time" in query:
                response = self.get_time()
                
            elif "date" in query:
                response = self.get_date()
                
            elif "open youtube" in query:
                response = "Opening YouTube"
                self.open_website("YouTube", "https://www.youtube.com")
                
            elif "open google" in query:
                response = "Opening Google"
                self.open_website("Google", "https://www.google.com")
                
            elif "camera" in query:
                response = "Camera activated"
                if self.gui:
                    self.gui.update_status("Camera active - Press 'q' to close")
                self.open_camera()
                
            elif "screenshot" in query:
                filepath = self.screenshot()
                response = f"Screenshot saved at: {filepath}"
                
            elif "volume" in query:
                response = self.control_volume(query)
                
            elif "offline" in query:
                response = "Shutting down Jarvis"
                if self.gui:
                    self.gui.update_status("Shutting down...")
                self.speak("Goodbye, sir!")
                self.running = False
                
            elif "open word" in query:
                response = "Opening Microsoft Word"
                self.open_office_app("word")
                
            elif "open excel" in query:
                response = "Opening Microsoft Excel"
                self.open_office_app("excel")
                
            elif "open powerpoint" in query:
                response = "Opening Microsoft PowerPoint" 
                self.open_office_app("powerpoint")
                
            elif "search for" in query or "google" in query and "open google" not in query:
                search_term = query.replace("search for", "").replace("google", "").strip()
                response = f"Searching for: {search_term}"
                self.google_search(query)
                
            else:
                # Use Gemini API for other queries if available
                if self.model:
                    ai_response = self.get_gemini_response(query)
                    self.speak(ai_response)
                    response = ai_response
                else:
                    response = "I'm sorry for now. I can only perform basic commands."
                    self.speak(response)
                
        except Exception as e:
            response = f"Error processing command: {str(e)}"
            status = "Error"
            print(f"Error processing command: {e}")
            
        # Log the interaction to CSV
        self.log_to_csv(query, response, status)
            
        # Reset status after command processing
        if self.gui:
            self.gui.update_status("Ready")

    def get_gemini_response(self, query):
        """Get a response from the Gemini API and display it in terminal."""
        try:
            # Try to use chat if available
            if self.chat:
                try:
                    # Send the query to Gemini chat
                    response = self.chat.send_message(query)
                    text_response = response.text
                except Exception as chat_error:
                    print(f"Chat error: {chat_error}, falling back to direct generation")
                    # Fall back to direct generation if chat fails
                    response = self.model.generate_content(query, safety_settings=self.safety_settings)
                    text_response = response.text
            else:
                # Use direct generation if chat isn't available
                response = self.model.generate_content(query, safety_settings=self.safety_settings)
                text_response = response.text
            
            # Display the response in terminal with formatting
            print("\n" + "-"*50)
            print(f"JARVIS Response to: '{query}'")
            print("-"*50)
            print(text_response)
            print("-"*50 + "\n")
            
            # If the response is too long, truncate it for speech
            if len(text_response) > 200:
                # Find the last sentence end within the first 200 characters
                end_pos = text_response[:200].rfind('.')
                if end_pos == -1:
                    end_pos = 200
                speech_response = text_response[:end_pos+1]
                # For logging and display, keep more text
                display_text = text_response[:500] + ("..." if len(text_response) > 500 else "")
                # Speak shorter version, return longer for display
                self.speak(speech_response)
                return display_text
            
            return text_response
        except Exception as e:
            error_msg = f"Error getting Gemini response: {e}"
            print(error_msg)
            return error_msg

    def open_website(self, site_name, url):
        """Open a specific website."""
        self.speak(f"Opening {site_name}.")
        wb.open(url)

    def screenshot(self):
        try:
            screenshots_dir = os.path.join(os.getcwd(), "screenshots")
            if not os.path.exists(screenshots_dir):
                os.makedirs(screenshots_dir)
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"screenshot_{timestamp}.png"
            filepath = os.path.join(screenshots_dir, filename)
            screenshot = ImageGrab.grab()
            screenshot.save(filepath)
            print(f"Screenshot saved successfully at: {filepath}")
            return filepath
        except Exception as e:
            print(f"Error taking screenshot: {e}")
            return None

    def control_volume(self, command):
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            
            result = "Volume adjusted"
            
            if "set volume to" in command.lower():
                try:
                    target = int(''.join(filter(str.isdigit, command)))
                    if 0 <= target <= 100:
                        volume.SetMasterVolumeLevelScalar(target/100, None)
                        print(f"Volume set to {target}%")
                        result = f"Volume set to {target}%"
                except:
                    print("Could not determine volume level from command")
                    result = "Could not set volume level"
            elif "increase" in command.lower() or "up" in command.lower():
                current = volume.GetMasterVolumeLevelScalar() * 100
                new_volume = min(100, current + 10)
                volume.SetMasterVolumeLevelScalar(new_volume/100, None)
                print(f"Volume increased to {int(new_volume)}%")
                result = f"Volume increased to {int(new_volume)}%"
            elif "decrease" in command.lower() or "down" in command.lower():
                current = volume.GetMasterVolumeLevelScalar() * 100
                new_volume = max(0, current - 10)
                volume.SetMasterVolumeLevelScalar(new_volume/100, None)
                print(f"Volume decreased to {int(new_volume)}%")
                result = f"Volume decreased to {int(new_volume)}%"
            
            return result
            
        except Exception as e:
            print(f"Error controlling volume: {e}")
            return f"Error controlling volume: {str(e)}"

    def open_camera(self):
        try:
            cap = cv2.VideoCapture(0)
            if not cap.isOpened():
                print("Error: Could not open camera")
                return
                
            while True:
                ret, frame = cap.read()
                if ret:
                    cv2.imshow('Camera', frame)
                    if cv2.waitKey(1) & 0xFF == ord('q'):
                        break
                else:
                    print("Error: Could not read frame")
                    break
                    
            cap.release()
            cv2.destroyAllWindows()
        except Exception as e:
            print(f"Error accessing camera: {e}")

    def open_office_app(self, app_name):
        try:
            office_paths = {
                "word": r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                "excel": r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
                "powerpoint": r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
            }
            
            office_paths_alt = {
                "word": r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
                "excel": r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
                "powerpoint": r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE"
            }

            app_name = app_name.lower()
            
            if app_name in office_paths:
                if os.path.exists(office_paths[app_name]):
                    subprocess.Popen(office_paths[app_name])
                    print(f"Opening {app_name.capitalize()}...")
                elif os.path.exists(office_paths_alt[app_name]):
                    subprocess.Popen(office_paths_alt[app_name])
                    print(f"Opening {app_name.capitalize()}...")
                else:
                    subprocess.Popen(f"start {app_name}.exe", shell=True)
                    print(f"Opening {app_name.capitalize()}...")
            else:
                print("Invalid application name. Please use 'word', 'excel', or 'powerpoint'")
        except Exception as e:
            print(f"Error opening {app_name}: {e}")

    def google_search(self, query):
        try:
            if "search for" in query.lower():
                search_query = query.lower().split("search for")[1].strip()
            elif "google" in query.lower():
                search_query = query.lower().split("google")[1].strip()
            else:
                search_query = query
                
            search_url = f"https://www.google.com/search?q={quote(search_query)}"
            wb.open(search_url)
            print(f"Searching Google for: {search_query}")
        except Exception as e:
            print(f"Error performing search: {e}")

    def run(self):
        self.wishme()
        while self.running:
            query = self.takecommand()
            if query != "none":
                self.process_command(query)

class JarvisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("JARVIS Assistant AI")
        self.root.geometry("600x400")
        self.root.configure(bg='#2C2F33')  # Dark theme background
        
        # Create main frame
        self.main_frame = tk.Frame(self.root, bg='#2C2F33')
        self.main_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # Title label
        title_label = tk.Label(self.main_frame, 
                             text="JARVIS Voice Assistant",
                             font=('Helvetica', 16, 'bold'),
                             fg='#FFFFFF',
                             bg='#2C2F33')
        title_label.pack(pady=5)
        
        # Status frame
        status_frame = tk.Frame(self.main_frame, bg='#2C2F33')
        status_frame.pack(fill=tk.X, pady=5)
        
        # Status label
        self.status_label = tk.Label(status_frame,
                                   text="Status: Ready",
                                   font=('Helvetica', 10),
                                   fg='#00FF00',
                                   bg='#2C2F33')
        self.status_label.pack(side=tk.LEFT)
        
        # Time label
        self.time_label = tk.Label(status_frame,
                                 text="",
                                 font=('Helvetica', 10),
                                 fg='#FFFFFF',
                                 bg='#2C2F33')
        self.time_label.pack(side=tk.RIGHT)
        
        # Create visualization frame
        viz_frame = tk.Frame(self.main_frame, bg='#2C2F33')
        viz_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create canvas for wave visualization
        self.canvas = tk.Canvas(viz_frame,
                              bg='#23272A',
                              height=150,
                              highlightthickness=1,
                              highlightbackground='#99AAB5')
        self.canvas.pack(fill=tk.BOTH, expand=True, padx=5)
        
        # Command history frame
        history_frame = tk.Frame(self.main_frame, bg='#2C2F33')
        history_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Command history label
        history_label = tk.Label(history_frame,
                               text="Command History",
                               font=('Helvetica', 10, 'bold'),
                               fg='#FFFFFF',
                               bg='#2C2F33')
        history_label.pack()
        
        # Command history text
        self.history_text = tk.Text(history_frame,
                                  height=6,
                                  bg='#23272A',
                                  fg='#FFFFFF',
                                  font=('Courier', 9),
                                  wrap=tk.WORD)
        self.history_text.pack(fill=tk.BOTH, expand=True)
        
        # View logs button
        self.logs_button = tk.Button(self.main_frame,
                                   text="View Speech Logs",
                                   command=self.open_logs_folder,
                                   bg='#7289DA',
                                   fg='#FFFFFF')
        self.logs_button.pack(pady=5)
        
        # Help text
        help_text = "Press 'q' to close camera when active\nSay 'offline' to exit"
        help_label = tk.Label(self.main_frame,
                            text=help_text,
                            font=('Helvetica', 8),
                            fg='#99AAB5',
                            bg='#2C2F33')
        help_label.pack(pady=5)
        
        # Initialize wave parameters
        self.wave_points = []
        self.amplitude = 30
        self.frequency = 0.1
        self.phase = 0
        
        # Start animation and clock
        self.animate()
        self.update_clock()
        
    def animate(self):
        if self.root.winfo_exists():
            self.canvas.delete("wave")
            self.draw_wave()
            self.phase += 0.1
            self.root.after(50, self.animate)
            
    def draw_wave(self):
        width = self.canvas.winfo_width()
        height = self.canvas.winfo_height()
        center_y = height / 2
        
        points = []
        for x in range(0, width, 5):
            y = center_y + self.amplitude * np.sin(self.frequency * x + self.phase)
            points.extend([x, y])
            
        if len(points) >= 4:
            self.canvas.create_line(points, fill='#7289DA', smooth=True, width=2, tags="wave")
            
    def update_clock(self):
        if self.root.winfo_exists():
            current_time = datetime.datetime.now().strftime("%I:%M:%S %p")
            self.time_label.config(text=current_time)
            self.root.after(1000, self.update_clock)
            
    def update_status(self, status, is_error=False):
        color = '#FF0000' if is_error else '#00FF00'
        self.status_label.config(text=f"Status: {status}", fg=color)
        
    def add_command_to_history(self, command):
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.history_text.insert('1.0', f"[{timestamp}] {command}\n")
        self.history_text.see('1.0')
        
    def open_logs_folder(self):
        """Open the folder containing speech logs."""
        try:
            log_dir = os.path.join(os.getcwd(), "speech_logs")
            if os.path.exists(log_dir):
                if os.name == 'nt':  # Windows
                    os.startfile(log_dir)
                elif os.name == 'posix':  # macOS or Linux
                    subprocess.call(['open', log_dir])
            else:
                print("Logs directory does not exist yet")
        except Exception as e:
            print(f"Error opening logs folder: {e}")

def main():
    # Get Gemini API key from environment variable or user input
    api_key = "AIzaSyAvUGKpo9BFOM7M1t8CN7u5WW65ENGEV20"
    
    # If API key is not set in environment, ask for it
    if not api_key:
        print("Please enter your Gemini API key:")
        api_key = input()
    
    # Create and start the voice assistant with Gemini integration
    jarvis = JarvisAssistant(api_key)
    
    # Create and start the GUI
    root = tk.Tk()
    gui = JarvisGUI(root)
    
    # Connect GUI to assistant
    jarvis.set_gui(gui)
    
    # Start assistant thread
    assistant_thread = threading.Thread(target=jarvis.run)
    assistant_thread.daemon = True
    assistant_thread.start()
    
    # Start GUI main loop
    root.mainloop()

if __name__ == "__main__":
    main()