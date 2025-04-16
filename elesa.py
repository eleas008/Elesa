import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pywhatkit
import speech_recognition as sr
import pyttsx3
import datetime
import threading
import time
import json
import os
import requests
from googlesearch import search
import wikipedia
import webbrowser
import random
import re

# === Setup for Speech Engine ===
listener = sr.Recognizer()
engine = pyttsx3.init()

# Check for settings file or create default
SETTINGS_FILE = "elisa_settings.json"
if os.path.exists(SETTINGS_FILE):
    with open(SETTINGS_FILE, 'r') as f:
        settings = json.load(f)
else:
    # Default settings
    settings = {
        "voice_id": 0,
        "speech_rate": 150,
        "wake_words": ["elisa", "elesa", "aleesa"],
        "theme": "dark",
        "volume": 1.0  # Add default volume
    }
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f)

# Apply voice settings
voices = engine.getProperty('voices')
if len(voices) > settings["voice_id"]:
    engine.setProperty('voice', voices[settings["voice_id"]].id)
engine.setProperty('rate', settings["speech_rate"])
engine.setProperty('volume', settings.get("volume", 1.0))  # Apply volume setting

# Command history and global flags
command_history = []
stop_requested = False  # Flag to indicate if a stop is requested

def show_voice_settings():
    """Display voice settings window with testing options"""
    settings_window = tk.Toplevel(app)
    settings_window.title("Voice Settings")
    settings_window.geometry("600x500")
    settings_window.config(bg="#1e1e2e")
    
    # Create frame for voice settings
    voice_frame = tk.Frame(settings_window, bg="#1e1e2e")
    voice_frame.pack(padx=20, pady=20, fill=tk.BOTH)
    
    # Voice selection
    tk.Label(voice_frame, text="Select Voice:", bg="#1e1e2e", fg="#ffffff", font=text_font).grid(row=0, column=0, sticky="w", pady=10)
    
    # Get all available voices
    voices = engine.getProperty('voices')
    voice_names = [f"{i}: {v.name} ({v.id})" for i, v in enumerate(voices)]
    
    voice_var = tk.StringVar(value=voice_names[settings["voice_id"]] if voices and settings["voice_id"] < len(voices) else "No voices found")
    voice_menu = ttk.Combobox(voice_frame, textvariable=voice_var, values=voice_names, state="readonly", width=40)
    voice_menu.grid(row=0, column=1, sticky="w", pady=10)
    
    # Speech rate control
    tk.Label(voice_frame, text="Speech Rate:", bg="#1e1e2e", fg="#ffffff", font=text_font).grid(row=1, column=0, sticky="w", pady=10)
    
    rate_var = tk.IntVar(value=settings["speech_rate"])
    rate_scale = tk.Scale(voice_frame, from_=50, to=300, orient=tk.HORIZONTAL, variable=rate_var,
                         bg="#2e2e3e", fg="#ffffff", length=300, showvalue=True)
    rate_scale.grid(row=1, column=1, sticky="w", pady=10)
    
    # Volume control (new setting)
    tk.Label(voice_frame, text="Volume:", bg="#1e1e2e", fg="#ffffff", font=text_font).grid(row=2, column=0, sticky="w", pady=10)
    
    # Get current volume or default to 1.0
    current_volume = settings.get("volume", 1.0)
    volume_var = tk.DoubleVar(value=current_volume)
    volume_scale = tk.Scale(voice_frame, from_=0.1, to=1.0, orient=tk.HORIZONTAL, resolution=0.1,
                           variable=volume_var, bg="#2e2e3e", fg="#ffffff", length=300)
    volume_scale.grid(row=2, column=1, sticky="w", pady=10)
    
    # Test area
    test_frame = tk.LabelFrame(settings_window, text="Test Voice", bg="#1e1e2e", fg="#ffffff", font=text_font)
    test_frame.pack(padx=20, pady=10, fill=tk.BOTH)
    
    test_text = tk.Text(test_frame, height=3, font=text_font, wrap=tk.WORD, bg="#2e2e3e", fg="#ffffff")
    test_text.pack(padx=10, pady=10, fill=tk.BOTH)
    test_text.insert(tk.END, "This is a test of the voice output. If you can hear this, the voice settings are working correctly.")
    
    # Status label
    status_var = tk.StringVar(value="Ready to test")
    status_label = tk.Label(test_frame, textvariable=status_var, bg="#1e1e2e", fg="#ffff40", font=text_font)
    status_label.pack(pady=5)
    
    # Buttons frame
    buttons_frame = tk.Frame(settings_window, bg="#1e1e2e")
    buttons_frame.pack(padx=20, pady=20, fill=tk.X)
    
    # Function to test the selected voice
    def test_selected_voice():
        status_var.set("Testing voice...")
        settings_window.update()
        
        # Get selected voice index
        selected = voice_menu.current()
        if selected == -1 and voices:  # No selection but voices exist
            selected = 0
        
        # Get test rate and volume
        test_rate = rate_var.get()
        test_volume = volume_var.get()
        
        # Configure engine for test
        if selected >= 0 and selected < len(voices):
            engine.setProperty('voice', voices[selected].id)
        engine.setProperty('rate', test_rate)
        engine.setProperty('volume', test_volume)
        
        # Get test text
        test_message = test_text.get("1.0", tk.END).strip()
        
        # Visual feedback
        status_var.set("Speaking now...")
        settings_window.update()
        
        # Log test information
        print(f"Testing voice: {selected}, Rate: {test_rate}, Volume: {test_volume}")
        print(f"Voice ID: {voices[selected].id if selected >= 0 and selected < len(voices) else 'None'}")
        
        try:
            # Try using the main engine first
            engine.say(test_message)
            engine.runAndWait()
            status_var.set("Test complete! Did you hear the voice?")
        except Exception as e:
            status_var.set(f"Main engine error: {str(e)}. Trying alternative method...")
            settings_window.update()
            
            try:
                # Fallback: Create a new engine instance for the test
                test_engine = pyttsx3.init()
                if selected >= 0 and selected < len(voices):
                    test_engine.setProperty('voice', voices[selected].id)
                test_engine.setProperty('rate', test_rate)
                test_engine.setProperty('volume', test_volume)
                test_engine.say(test_message)
                test_engine.runAndWait()
                status_var.set("Alternative test complete! Did you hear the voice?")
            except Exception as e2:
                status_var.set(f"Error: {str(e2)}")
                print(f"Voice test error: {e2}")
    
    def try_windows_sapi():
        """Attempt to use Windows SAPI5 directly as a fallback"""
        status_var.set("Trying Windows SAPI5 directly...")
        settings_window.update()
        
        try:
            import win32com.client
            speaker = win32com.client.Dispatch("SAPI.SpVoice")
            speaker.Volume = int(volume_var.get() * 100)
            speaker.Rate = int((rate_var.get() - 150) / 25)  # Convert to SAPI rate (-10 to 10)
            
            test_message = test_text.get("1.0", tk.END).strip()
            speaker.Speak(test_message)
            status_var.set("Windows SAPI5 test complete! Did you hear the voice?")
        except Exception as e:
            status_var.set(f"Windows SAPI5 error: {str(e)}")
            print(f"Windows SAPI5 error: {e}")
    
    def save_settings():
        # Save the new settings
        selected = voice_menu.current()
        if selected >= 0:
            settings["voice_id"] = selected
            
        settings["speech_rate"] = rate_var.get()
        settings["volume"] = volume_var.get()
        
        # Apply settings to the engine
        if selected >= 0 and selected < len(voices):
            engine.setProperty('voice', voices[selected].id)
        engine.setProperty('rate', settings["speech_rate"])
        engine.setProperty('volume', settings["volume"])
        
        # Save to file
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f)
            
        status_var.set("Settings saved!")
        
        # Also show in main window
        update_response("Voice settings updated! Try speaking a command to test.")
    
    # Buttons
    test_button = ttk.Button(buttons_frame, text="Test Voice", command=test_selected_voice)
    test_button.pack(side=tk.LEFT, padx=5)
    
    # Add Windows SAPI test button (Windows-specific fallback)
    if os.name == 'nt':  # Windows systems only
        windows_test_button = ttk.Button(buttons_frame, text="Try Windows SAPI", command=try_windows_sapi)
        windows_test_button.pack(side=tk.LEFT, padx=5)
    
    save_button = ttk.Button(buttons_frame, text="Save Settings", command=save_settings)
    save_button.pack(side=tk.LEFT, padx=5)
    
    close_button = ttk.Button(buttons_frame, text="Close", command=settings_window.destroy)
    close_button.pack(side=tk.RIGHT, padx=5)
    
    # Diagnostic info section
    diag_frame = tk.LabelFrame(settings_window, text="Voice Diagnostics", bg="#1e1e2e", fg="#ffffff", font=text_font)
    diag_frame.pack(padx=20, pady=10, fill=tk.BOTH)
    
    diag_text = tk.Text(diag_frame, height=6, font=("Consolas", 10), wrap=tk.WORD, bg="#2e2e3e", fg="#ffffff")
    diag_text.pack(padx=10, pady=10, fill=tk.BOTH)
    
    # Add diagnostic info
    diag_info = f"System: {os.name} ({sys.platform if 'sys' in globals() else 'unknown'})\n"
    diag_info += f"Python version: {sys.version.split()[0] if 'sys' in globals() else 'unknown'}\n"
    diag_info += f"pyttsx3 version: {pyttsx3.__version__ if hasattr(pyttsx3, '__version__') else 'unknown'}\n"
    diag_info += f"Available voices: {len(voices)}\n"
    
    if len(voices) > 0:
        diag_info += f"Current voice: {settings['voice_id']} - "
        if settings['voice_id'] < len(voices):
            diag_info += f"{voices[settings['voice_id']].name}\n"
        else:
            diag_info += "Invalid voice ID!\n"
    else:
        diag_info += "No voices detected! TTS may not work.\n"
        
    diag_info += f"Engine driver: {engine.getProperty('driver') if hasattr(engine, 'getProperty') and callable(engine.getProperty) else 'unknown'}"
    
    diag_text.insert(tk.END, diag_info)
    diag_text.config(state='disabled')

def stop_all_processes():
    """Stop all ongoing processes"""
    global stop_requested
    stop_requested = True
    
    # Try to stop the text-to-speech engine
    try:
        engine.stop()
    except:
        # Some implementations of pyttsx3 don't support stop
        pass
    
    # Update UI to show we've stopped
    update_response("All processes stopped.")
    
    # Reset buttons state
    app.after(0, lambda: speak_button.config(state='normal', bg="#1e1e2e"))
    app.after(0, lambda: send_button.config(state='normal'))
    app.after(0, lambda: stop_button.config(state='disabled'))
    
    # Wait a bit and reset the flag
    app.after(1000, reset_stop_flag)

def reset_stop_flag():
    """Reset the stop flag after stopping processes"""
    global stop_requested
    stop_requested = False
    app.after(0, lambda: stop_button.config(state='normal'))

def talk(text):
    """Convert text to speech with enhanced error handling"""
    global stop_requested
    
    if stop_requested:
        return
        
    try:
        # Visual feedback that speech is happening
        update_response(f"Speaking: {text}", is_status=True)
        
        # Set volume (new property)
        volume = settings.get("volume", 1.0)
        engine.setProperty('volume', volume)
        
        # Speak the text
        engine.say(text)
        engine.runAndWait()
        
        # Reset the response after speaking
        update_response(text)
    except Exception as e:
        error_msg = f"Error in text-to-speech: {e}"
        print(error_msg)
        update_response(f"⚠️ Voice output error: {e}\nTry the 'Voice Settings' button to fix.")
        
        # Try alternate method using Windows SAPI directly
        if os.name == 'nt':  # Windows only
            try:
                import win32com.client
                speaker = win32com.client.Dispatch("SAPI.SpVoice")
                speaker.Volume = int(settings.get("volume", 1.0) * 100)
                speaker.Rate = int((settings.get("speech_rate", 150) - 150) / 25)
                speaker.Speak(text)
                print("Used Windows SAPI directly as fallback")
                update_response(text)  # Reset display
            except Exception as e2:
                print(f"Windows SAPI fallback failed: {e2}")

def check_microphone():
    """Check if a microphone is available"""
    try:
        with sr.Microphone() as source:
            return True
    except OSError:
        return False

def take_command():
    """Listen for voice commands"""
    global stop_requested
    
    try:
        with sr.Microphone() as source:
            update_response("Listening...", is_status=True)
            # Visual feedback for listening
            app.after(0, lambda: speak_button.config(bg="#ff4040"))
            
            listener.adjust_for_ambient_noise(source, duration=0.5)
            
            # Enable stop button during listening
            app.after(0, lambda: stop_button.config(state='normal'))
            
            voice = listener.listen(source, timeout=5)
            
            if stop_requested:
                return ""
                
            update_response("Processing your request...", is_status=True)
            # Visual feedback for processing
            app.after(0, lambda: speak_button.config(bg="#ff9940"))
            
            command = listener.recognize_google(voice)
            command = command.lower()
            
            # Check for wake words
            if any(word in command for word in settings["wake_words"]):
                # Remove the wake word
                for word in settings["wake_words"]:
                    command = command.replace(word, '')
                command = command.strip()
                update_response(f"You said: {command}")
                add_to_history("You", command)
                return command
            else:
                update_response("Wake word not detected. Please say 'Elisa' followed by your command.")
    except sr.WaitTimeoutError:
        update_response("Listening timed out. Please try again.")
    except sr.UnknownValueError:
        update_response("Could not understand audio")
    except sr.RequestError as e:
        update_response(f"Speech Recognition request error: {e}")
    except Exception as e:
        update_response(f"Error: {e}")
    return ""

def add_to_history(speaker, text):
    """Add a command or response to history"""
    timestamp = datetime.datetime.now().strftime("%H:%M:%S")
    command_history.append({"time": timestamp, "speaker": speaker, "text": text})
    # Keep only the last 20 items
    if len(command_history) > 20:
        command_history.pop(0)

def open_url(url):
    """Open URL in web browser"""
    try:
        webbrowser.open(url)
    except Exception as e:
        update_response(f"Error opening URL: {e}")

def get_weather(city):
    """Get weather information for a city"""
    global stop_requested
    
    if stop_requested:
        return "Weather request stopped."
        
    try:
        # Using a free API (no key required)
        url = f"https://wttr.in/{city}?format=j1"
        response = requests.get(url)
        data = response.json()
        current = data['current_condition'][0]
        temp_c = current['temp_C']
        temp_f = current['temp_F']
        desc = current['weatherDesc'][0]['value']
        humidity = current['humidity']
        
        weather_info = f"Weather in {city}: {desc}, {temp_c}°C ({temp_f}°F), Humidity: {humidity}%"
        return weather_info
    except Exception as e:
        return f"Could not get weather information: {str(e)}"

def run_elisa():
    """Main function to handle commands"""
    global stop_requested
    
    # Enable stop button when processing starts
    app.after(0, lambda: stop_button.config(state='normal'))
    
    if not check_microphone():
        update_response("No microphone detected. Please connect a microphone and try again.")
        talk("No microphone detected")
        return
        
    command = take_command()
    if not command or stop_requested:
        if stop_requested:
            update_response("Command processing stopped.")
        return
    
    response = None
        
    # Process commands
    if 'play' in command:
        if stop_requested:
            return
        song = command.replace('play', '').strip()
        response = f"Playing {song} on YouTube"
        talk(response)
        update_response(response)
        add_to_history("Elisa", response)
        pywhatkit.playonyt(song)

    elif any(time_word in command for time_word in ['time', 'what time', 'current time']):
        if stop_requested:
            return
        current_time = datetime.datetime.now().strftime("%I:%M %p")
        response = f"Current time is {current_time}"
        talk(response)
        update_response(response)
        add_to_history("Elisa", response)

    elif 'date' in command or 'today' in command:
        if stop_requested:
            return
        current_date = datetime.datetime.now().strftime("%A, %B %d, %Y")
        response = f"Today is {current_date}"
        talk(response)
        update_response(response)
        add_to_history("Elisa", response)

    elif 'search' in command:
        if stop_requested:
            return
        query = command.replace('search', '').strip()
        update_response(f"Searching for: {query}")
        try:
            num_result = 5
            # Fixed parameter name for googlesearch-python package
            search_result = search(query, num_results=num_result)
            
            if stop_requested:
                update_response("Search stopped.")
                return
                
            talk("Here are the search results. You can click any link to open it.")
            
            # Format results with explicit clickable links
            results = "Here are your search results:\n\n"
            for i, url in enumerate(search_result):
                if stop_requested:
                    update_response("Search stopped.")
                    return
                results += f"Result {i+1}: {url}\n\n"
                
            update_response(results)
            response = "Click any link above to open it in your browser."
            add_to_history("Elisa", results + "\n\n" + response)
        except Exception as e:
            response = f"Search error: {e}"
            update_response(response)
            talk("I encountered an error while searching")
            add_to_history("Elisa", response)

    elif 'weather' in command or 'temperature' in command:
        if stop_requested:
            return
        # Extract city name
        if 'in' in command:
            city = command.split('in', 1)[1].strip()
        else:
            city = command.replace('weather', '').replace('temperature', '').strip()
            if not city:
                city = "London"  # Default city
                
        weather_info = get_weather(city)
        if stop_requested:
            return
        talk(weather_info)
        update_response(weather_info)
        response = weather_info
        add_to_history("Elisa", response)

    elif 'open' in command:
        if stop_requested:
            return
        # Handle opening websites
        if 'website' in command or '.com' in command or '.org' in command:
            site = command.replace('open', '').replace('website', '').strip()
            if not (site.startswith('http://') or site.startswith('https://')):
                if not ('.' in site):
                    site = f"https://www.{site}.com"
                else:
                    site = f"https://{site}"
            response = f"Opening {site}"
            talk(response)
            update_response(response)
            add_to_history("Elisa", response)
            open_url(site)
        else:
            app_name = command.replace('open', '').strip()
            response = f"I don't know how to open {app_name} yet"
            talk(response)
            update_response(response)
            add_to_history("Elisa", response)

    elif 'tell me about' in command or 'who is' in command or 'what is' in command:
        if stop_requested:
            return
        # Extract the topic from different command formats
        if 'tell me about' in command:
            topic = command.split('tell me about', 1)[-1].strip()
        elif 'who is' in command:
            topic = command.split('who is', 1)[-1].strip()
        elif 'what is' in command:
            topic = command.split('what is', 1)[-1].strip()
            
        update_response(f"Searching Wikipedia for: {topic}")
        try:
            num_results = 5
            # Fixed parameter name for googlesearch-python package
            search_results = search(topic, num_results=num_results)
            
            if stop_requested:
                update_response("Wikipedia search stopped.")
                return
                
            wiki_url = next((res for res in search_results if 'wikipedia.org' in res), None)

            if wiki_url:
                title = wiki_url.split('/')[-1]
                try:
                    page = wikipedia.page(title)
                    
                    if stop_requested:
                        update_response("Wikipedia search stopped.")
                        return
                        
                    summary = page.summary.split('\n')[0]  # first paragraph
                    full_response = f"{summary}\n\nRead more: {wiki_url}"
                    talk(summary)
                    update_response(full_response)
                    add_to_history("Elisa", full_response)
                except wikipedia.exceptions.PageError:
                    response = "Wikipedia page not found."
                    talk("Couldn't find the page.")
                    update_response(response)
                    add_to_history("Elisa", response)
            else:
                response = "No Wikipedia page found."
                talk(response)
                update_response(response)
                add_to_history("Elisa", response)

        except Exception as e:
            response = f"Error: {e}"
            update_response(response)
            talk("Something went wrong.")
            add_to_history("Elisa", response)
    
    elif 'thank you' in command or 'thanks' in command:
        if stop_requested:
            return
        responses = ["You're welcome!", "Happy to help!", "Anytime!", "No problem!"]
        response = random.choice(responses)
        talk(response)
        update_response(response)
        add_to_history("Elisa", response)
        
    elif 'help' in command:
        if stop_requested:
            return
        help_text = """
I can help you with the following commands:
- "play [song name]" - Play a song on YouTube
- "what time is it" - Tell the current time
- "what's today's date" - Tell the current date
- "search for [query]" - Search the web
- "tell me about [topic]" - Get information from Wikipedia
- "what is [topic]" - Get information from Wikipedia
- "who is [person]" - Get information about a person
- "weather in [city]" - Get current weather
- "open [website]" - Open a website
- "thank you" - Express gratitude
- "help" - Show this help message
- "stop" - Stop any ongoing process
"""
        talk("Here are the things I can help you with")
        update_response(help_text)
        response = help_text
        add_to_history("Elisa", response)
    
    # Added voice test command
    elif 'test voice' in command:
        if stop_requested:
            return
        test_message = "This is a test of the voice system. If you can hear this, the voice is working correctly."
        talk(test_message)
        update_response(test_message)
        add_to_history("Elisa", test_message)
    
    # Add a direct command to stop
    elif 'stop' in command:
        stop_all_processes()
    
    # Voice settings command
    elif 'voice settings' in command or 'change voice' in command:
        response = "Opening voice settings panel."
        update_response(response) 
        add_to_history("Elisa", response)
        show_voice_settings()
    
    else:
        if stop_requested:
            return
        response = "I'm not sure how to help with that yet. Try asking for help to see what I can do."
        talk(response)
        update_response(response)
        add_to_history("Elisa", response)

def process_text_command():
    """Process command from text input"""
    command = text_input.get().strip().lower()
    if command:
        text_input.delete(0, tk.END)
        add_to_history("You", command)
        update_response(f"You typed: {command}")
        
        # Direct stop command handling
        if command == 'stop':
            stop_all_processes()
            return
        
        # Direct voice settings command
        if command == 'voice settings':
            show_voice_settings()
            return
        
        # Process the text command
        if any(word in command for word in settings["wake_words"]):
            # Remove the wake word
            for word in settings["wake_words"]:
                command = command.replace(word, '')
            command = command.strip()
        
        # Create a thread to process the command
        threading.Thread(target=lambda: process_command_thread(command)).start()

def process_command_thread(command):
    """Process a command in a separate thread"""
    global stop_requested
    
    # Enable stop button when processing starts
    app.after(0, lambda: stop_button.config(state='normal'))
    
    # Disable buttons during processing
    app.after(0, lambda: speak_button.config(state='disabled'))
    app.after(0, lambda: send_button.config(state='disabled'))
    
    try:
        response = None
        
        # Process commands (duplicating the logic from run_elisa)
        if 'play' in command:
            if stop_requested:
                return
            song = command.replace('play', '').strip()
            response = f"Playing {song} on YouTube"
            talk(response)
            update_response(response)
            add_to_history("Elisa", response)
            pywhatkit.playonyt(song)

        elif any(time_word in command for time_word in ['time', 'what time', 'current time']):
            if stop_requested:
                return
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            response = f"Current time is {current_time}"
            talk(response)
            update_response(response)
            add_to_history("Elisa", response)

        elif 'date' in command or 'today' in command:
            if stop_requested:
                return
            current_date = datetime.datetime.now().strftime("%A, %B %d, %Y")
            response = f"Today is {current_date}"
            talk(response)
            update_response(response)
            add_to_history("Elisa", response)

        elif 'search' in command:
            if stop_requested:
                return
            query = command.replace('search', '').strip()
            update_response(f"Searching for: {query}")
            try:
                num_result = 5
                # Fixed parameter name for googlesearch-python package
                search_result = search(query, num_results=num_result)
                
                if stop_requested:
                    update_response("Search stopped.")
                    return
                    
                talk("Here are the search results. You can click any link to open it.")
                
                # Format results with explicit clickable links
                results = "Here are your search results:\n\n"
                for i, url in enumerate(search_result):
                    if stop_requested:
                        update_response("Search stopped.")
                        return
                    results += f"Result {i+1}: {url}\n\n"
                    
                update_response(results)
                response = "Click any link above to open it in your browser."
                add_to_history("Elisa", results + "\n\n" + response)
            except Exception as e:
                response = f"Search error: {e}"
                update_response(response)
                talk("I encountered an error while searching")
                add_to_history("Elisa", response)

        elif 'weather' in command or 'temperature' in command:
            if stop_requested:
                return
            # Extract city name
            if 'in' in command:
                city = command.split('in', 1)[1].strip()
            else:
                city = command.replace('weather', '').replace('temperature', '').strip()
                if not city:
                    city = "London"  # Default city
                    
            weather_info = get_weather(city)
            if stop_requested:
                return
            talk(weather_info)
            update_response(weather_info)
            add_to_history("Elisa", weather_info)

        elif 'open' in command:
            if stop_requested:
                return
            # Handle opening websites
            if 'website' in command or '.com' in command or '.org' in command:
                site = command.replace('open', '').replace('website', '').strip()
                if not (site.startswith('http://') or site.startswith('https://')):
                    if not ('.' in site):
                        site = f"https://www.{site}.com"
                    else:
                        site = f"https://{site}"
                response = f"Opening {site}"
                talk(response)
                update_response(response)
                add_to_history("Elisa", response)
                open_url(site)
            else:
                app_name = command.replace('open', '').strip()
                response = f"I don't know how to open {app_name} yet"
                talk(response)
                update_response(response)
                add_to_history("Elisa", response)

        elif 'tell me about' in command or 'who is' in command or 'what is' in command:
            if stop_requested:
                return
            # Extract the topic from different command formats
            if 'tell me about' in command:
                topic = command.split('tell me about', 1)[-1].strip()
            elif 'who is' in command:
                topic = command.split('who is', 1)[-1].strip()
            elif 'what is' in command:
                topic = command.split('what is', 1)[-1].strip()
                
            update_response(f"Searching Wikipedia for: {topic}")
            try:
                num_results = 5
                # Fixed parameter name for googlesearch-python package
                search_results = search(topic, num_results=num_results)
                
                if stop_requested:
                    update_response("Wikipedia search stopped.")
                    return
                    
                wiki_url = next((res for res in search_results if 'wikipedia.org' in res), None)

                if wiki_url:
                    title = wiki_url.split('/')[-1]
                    try:
                        page = wikipedia.page(title)
                        
                        if stop_requested:
                            update_response("Wikipedia search stopped.")
                            return
                            
                        summary = page.summary.split('\n')[0]  # first paragraph
                        full_response = f"{summary}\n\nRead more: {wiki_url}"
                        talk(summary)
                        update_response(full_response)
                        add_to_history("Elisa", full_response)
                    except wikipedia.exceptions.PageError:
                        response = "Wikipedia page not found."
                        talk("Couldn't find the page.")
                        update_response(response)
                        add_to_history("Elisa", response)
                else:
                    response = "No Wikipedia page found."
                    talk(response)
                    update_response(response)
                    add_to_history("Elisa", response)

            except Exception as e:
                response = f"Error: {e}"
                update_response(response)
                talk("Something went wrong.")
                add_to_history("Elisa", response)
        
        elif 'thank you' in command or 'thanks' in command:
            if stop_requested:
                return
            responses = ["You're welcome!", "Happy to help!", "Anytime!", "No problem!"]
            response = random.choice(responses)
            talk(response)
            update_response(response)
            add_to_history("Elisa", response)
            
        elif 'help' in command:
            if stop_requested:
                return
            help_text = """
I can help you with the following commands:
- "play [song name]" - Play a song on YouTube
- "what time is it" - Tell the current time
- "what's today's date" - Tell the current date
- "search for [query]" - Search the web
- "tell me about [topic]" - Get information from Wikipedia
- "what is [topic]" - Get information from Wikipedia
- "who is [person]" - Get information about a person
- "weather in [city]" - Get current weather
- "open [website]" - Open a website
- "thank you" - Express gratitude
- "help" - Show this help message
- "test voice" - Test the voice output
- "voice settings" - Change voice settings
- "stop" - Stop any ongoing process
"""
            talk("Here are the things I can help you with")
            update_response(help_text)
            add_to_history("Elisa", help_text)
        
        # Added voice test command
        elif 'test voice' in command:
            if stop_requested:
                return
            test_message = "This is a test of the voice system. If you can hear this, the voice is working correctly."
            talk(test_message)
            update_response(test_message)
            add_to_history("Elisa", test_message)
        
        # Voice settings command
        elif 'voice settings' in command or 'change voice' in command:
            response = "Opening voice settings panel."
            update_response(response) 
            add_to_history("Elisa", response)
            show_voice_settings()
        
        # Add a direct command to stop
        elif 'stop' in command:
            stop_all_processes()
        
        else:
            if stop_requested:
                return
            response = "I'm not sure how to help with that yet. Try asking for help to see what I can do."
            talk(response)
            update_response(response)
            add_to_history("Elisa", response)
    
    finally:
        # Re-enable buttons and reset color if not stopped
        if not stop_requested:
            app.after(0, lambda: speak_button.config(state='normal', bg="#1e1e2e"))
            app.after(0, lambda: send_button.config(state='normal'))
            app.after(0, lambda: stop_button.config(state='disabled'))

def show_history():
    """Display command history in a new window"""
    history_window = tk.Toplevel(app)
    history_window.title("Conversation History")
    history_window.geometry("500x400")
    history_window.config(bg="#1e1e2e")
    
    # Create a scrolled text widget
    history_text = scrolledtext.ScrolledText(history_window, font=text_font, bg="#2e2e3e", fg="#ffffff")
    history_text.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
    
    # Populate with history
    for item in command_history:
        timestamp = item["time"]
        speaker = item["speaker"]
        text = item["text"]
        
        if speaker == "You":
            history_text.insert(tk.END, f"{timestamp} - You: ", "user")
            history_text.insert(tk.END, f"{text}\n\n", "user_text")
        else:
            history_text.insert(tk.END, f"{timestamp} - Elisa: ", "assistant")
            # Check for links in the text
            parts = re.split(r'(https?://\S+)', text)
            for i, part in enumerate(parts):
                if i % 2 == 0:  # Regular text
                    history_text.insert(tk.END, part, "assistant_text")
                else:  # URL
                    history_text.insert(tk.END, part, "hyperlink")
            history_text.insert(tk.END, "\n\n")
    
    # Add tags for styling
    history_text.tag_config("user", foreground="#ff9940")
    history_text.tag_config("user_text", foreground="#ffffff")
    history_text.tag_config("assistant", foreground="#40a0ff")
    history_text.tag_config("assistant_text", foreground="#ffffff")
    history_text.tag_config("hyperlink", foreground="#4287f5", underline=1)
    history_text.tag_bind("hyperlink", "<Button-1>", lambda e: open_url_from_history(history_text))
    
    # Clear button
    clear_button = ttk.Button(history_window, text="Clear History", 
                             command=lambda: [command_history.clear(), history_window.destroy()])
    clear_button.pack(pady=10)

def open_url_from_history(text_widget):
    """Open URL that was clicked in the history text"""
    try:
        # Get text near cursor
        text = text_widget.get("insert linestart", "insert lineend")
        # Extract URL using regex
        match = re.search(r'(https?://\S+)', text)
        if match:
            url = match.group(1)
            open_url(url)
    except Exception as e:
        print(f"Error opening URL from history: {e}")

# === GUI Functions ===
def start_voice_command():
    global stop_requested
    stop_requested = False  # Reset stop flag on new command
    speak_button.config(state='disabled')
    threading.Thread(target=run_voice_command_with_reset).start()

def run_voice_command_with_reset():
    try:
        run_elisa()
    finally:
        # Reset the button state and color if not stopped
        if not stop_requested:
            app.after(0, lambda: speak_button.config(state='normal', bg="#1e1e2e"))
            app.after(0, lambda: stop_button.config(state='disabled'))

def update_response(text, is_status=False):
    """Display text in the response area with clickable links"""
    response_text.config(state='normal')
    
    if is_status:
        # For status messages, replace the last line only
        response_text.delete(1.0, tk.END)
        response_text.insert(tk.END, text)
    else:
        # For regular responses, clear and insert new text
        response_text.delete(1.0, tk.END)
        
        # Check for URLs in the text
        parts = re.split(r'(https?://\S+)', text)
        
        for i, part in enumerate(parts):
            if i % 2 == 0:  # Regular text
                response_text.insert(tk.END, part)
            else:  # URL - make clickable
                response_text.insert(tk.END, part, "hyperlink")
    
    # Configure the hyperlink tag for styling and behavior
    response_text.tag_config("hyperlink", foreground="#4287f5", underline=1)
    response_text.tag_bind("hyperlink", "<Button-1>", lambda e: open_url_from_text(response_text))
    
    response_text.config(state='disabled')
    # Force update the display
    app.update_idletasks()

def open_url_from_text(text_widget):
    """Open URL that was clicked in the response text"""
    try:
        # Get current text near the cursor
        current_pos = text_widget.index(tk.CURRENT)
        line_start = text_widget.index(f"{current_pos} linestart")
        line_end = text_widget.index(f"{current_pos} lineend")
        line_text = text_widget.get(line_start, line_end)
        
        # Extract URL using regex
        match = re.search(r'(https?://\S+)', line_text)
        if match:
            url = match.group(1)
            open_url(url)
    except Exception as e:
        print(f"Error opening URL from text: {e}")

def test_voice():
    """Test the voice output"""
    global stop_requested
    stop_requested = False  # Reset stop flag on new command
    app.after(0, lambda: stop_button.config(state='normal'))
    
    threading.Thread(target=lambda: [
        update_response("Testing voice output..."),
        talk("This is a test of the voice system. If you can hear this, the voice is working correctly."),
        update_response("Voice test complete. Did you hear the voice?"),
        app.after(0, lambda: stop_button.config(state='disabled'))
    ]).start()

# === GUI Design ===
app = tk.Tk()
app.title("Elisa - Your Personal Assistant")
app.geometry("1500x800")
app.config(bg="#1e1e2e")

# Import sys for system info
import sys

title_font = ("Segoe UI", 20, "bold")
text_font = ("Segoe UI", 12)

title_label = tk.Label(app, text="👩‍💻 Elisa - Your Personal Voice Assistant", font=title_font, bg="#1e1e2e", fg="#ffffff")
title_label.pack(pady=20)

# Response Box - now supports clickable links
response_text = tk.Text(app, height=10, font=text_font, wrap=tk.WORD, bg="#2e2e3e", fg="#ffffff", insertbackground='white', cursor="arrow")
response_text.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)
response_text.config(state='disabled')

input_frame = tk.Frame(app, bg="#1e1e2e")
input_frame.pack(fill=tk.X, padx=20, pady=10)

text_input = ttk.Entry(input_frame, font=text_font, width=40)
text_input.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
text_input.bind("<Return>", lambda event: process_text_command())

send_button = ttk.Button(input_frame, text="Send", command=process_text_command)
send_button.pack(side=tk.RIGHT)

button_frame = tk.Frame(app, bg="#1e1e2e")
button_frame.pack(fill=tk.X, padx=20, pady=10)

speak_button = tk.Button(button_frame, text="🎙️ Speak to Elisa", command=start_voice_command, 
                      bg="#1e1e2e", fg="#ffffff", font=text_font, relief=tk.RAISED, borderwidth=2)
speak_button.pack(side=tk.LEFT, padx=5)

history_button = ttk.Button(button_frame, text="📜 History", command=show_history)
history_button.pack(side=tk.LEFT, padx=5)

# Voice Settings Button - NEW
voice_settings_button = ttk.Button(button_frame, text="🔈 Voice Settings", command=show_voice_settings)
voice_settings_button.pack(side=tk.LEFT, padx=5)

# Test Voice Button
test_voice_button = ttk.Button(button_frame, text="🔊 Test Voice", command=test_voice)
test_voice_button.pack(side=tk.LEFT, padx=5)

# Stop Button
stop_button = tk.Button(button_frame, text="🛑 Stop", command=stop_all_processes,
                     bg="#ff4040", fg="#ffffff", font=text_font, relief=tk.RAISED, borderwidth=2)
stop_button.pack(side=tk.LEFT, padx=5)
stop_button.config(state='disabled')  # Initially disabled until a process starts

help_button = ttk.Button(button_frame, text="❓ Help", 
                         command=lambda: [update_response("""I can help you with:
- "play [song name]" - Play a song on YouTube
- "what time is it" - Tell the current time
- "what's today's date" - Tell the current date
- "search for [query]" - Search the web
- "tell me about [topic]" - Get information from Wikipedia
- "what is [topic]" - Get information from Wikipedia
- "who is [person]" - Get information about a person
- "weather in [city]" - Get current weather
- "open [website]" - Open a website
- "test voice" - Test the voice output system
- "voice settings" - Change voice settings
- "stop" - Stop any ongoing process
""")])
help_button.pack(side=tk.LEFT, padx=5)

update_response("Hello! I'm Elisa, your personal assistant. Speak to me or type a command.")

# Check microphone
mic_available = check_microphone()
if not mic_available:
    messagebox.warning("Microphone Not Found", "No microphone detected. Please connect a microphone to use voice commands.")

# Test text-to-speech at startup and show diagnostic
try:
    voice_count = len(engine.getProperty('voices'))
    current_voice = settings["voice_id"]
    test_result = f"Found {voice_count} voices. Using voice #{current_voice}."
    
    print("TTS Diagnostic:")
    print(test_result)
    print(f"Speech rate: {settings['speech_rate']}")
    print(f"Volume: {settings.get('volume', 1.0)}")
    
    if voice_count == 0:
        messagebox.showwarning("Voice Issue", "No text-to-speech voices found. The assistant might not be able to speak. Try using the Voice Settings button.")
except Exception as e:
    print(f"TTS system error: {e}")
    messagebox.showwarning("Voice System Issue", 
                         f"There might be a problem with the text-to-speech system: {e}\nTry using the Voice Settings button.")

app.mainloop()