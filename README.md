Elisa is a desktop voice assistant application built in Python that provides a conversational interface for performing various tasks and retrieving information.
The program features both voice and text input methods with a modern, user-friendly interface.

Core Features
Voice Interaction: Recognizes voice commands using wake words ("Elisa", "Elesa", "Aleesa")
Text Input: Alternative method for entering commands through typing
Customizable Voice: Adjustable voice type, speech rate, and volume settings
Web Integration: Search capabilities, Wikipedia lookups, YouTube video playback
Information Services: Weather reports, date/time information
Website Navigation: Can open websites via voice or text commands
Conversation History: Records and displays past interactions

Technical Components
Speech Processing:

Voice recognition using Google's speech recognition API
Text-to-speech with pyttsx3 and Windows SAPI fallback
Multiple error handling layers for reliability
User Interface:

Clean, dark-themed Tkinter GUI
Real-time visual feedback during command processing
Clickable hyperlinks in responses
Responsive button states during operations
Command Handling:

Multi-threaded processing to keep UI responsive
Interruptible operations via stop button
Structured command parsing with keyword detection

Supported Commands
The assistant responds to commands like:

"Play [song name]" - YouTube playback
"What time is it" - Current time
"Today's date" - Current date
"Search for [query]" - Web search results
"Tell me about [topic]" - Wikipedia information
"Weather in [city]" - Weather conditions
"Open [website]" - Browser navigation
"Voice settings" - Customize speech output
"Test voice" - Verify speech functionality
"Help" - List available commands
The application includes robust error handling and diagnostic capabilities, making it suitable for everyday use while providing troubleshooting options when needed.
