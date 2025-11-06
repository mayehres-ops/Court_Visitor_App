#!/usr/bin/env python3
"""
Court Visitor Program Chatbot - Now with 100% more personality!
The most sarcastic, helpful, and occasionally obnoxious assistant you'll ever meet.
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json
import re
import random
from datetime import datetime
from pathlib import Path

class CourtVisitorChatbot:
    def __init__(self, parent=None):
        self.parent = parent
        self.chatbot_window = None
        self.conversation_history = []
        self.visit_count = 0
        self.typing_job = None
        self.user_name = None

        # Load or create visit stats
        self.stats_file = Path(__file__).parent / "App Data" / "chatbot_stats.json"
        self.load_stats()

        # Personality traits
        self.moods = ["sassy", "helpful", "sarcastic", "playful", "dramatic"]
        self.current_mood = random.choice(self.moods)

        # SUPER Sarcastic greetings with Matt Rife charm (Monday.com style + charisma!)
        self.greetings = {
            1: [
                "Oh look, a fresh victim! I mean... WELCOME! *nervous laughter* 👋 First time? Don't worry, you're gonna crush this.",
                "Well hello there! First time? Lucky me. 😏 Let's get you set up. You seem like the type who actually reads instructions. Respect.",
                "Welcome! I promise to only be MODERATELY obnoxious. (That's a lie. I'm VERY obnoxious.) But hey, at least I make you smile while we automate, right?",
                "Oh hi! New court visitor? Love that for you. Volunteering to help people is basically the most attractive quality a person can have. Just saying. ✨",
                "First timer! Perfect! Look, I'm gonna be real with you: You're about to become my favorite person. Why? Because you're here doing volunteer work. That's incredible. Let's do this! 💪"
            ],
            2: [
                "Ohhh you're back! See, I knew you couldn't stay away. 😏 (Or Excel broke. But I'm choosing to believe you missed me.)",
                "Round 2! Look at you, coming back for more automation magic. I respect the dedication. Court visitors like you don't quit. 💪",
                "Second visit! Either you really like me or Excel's being difficult again. Actually, you know what? Doesn't matter. I'm just glad you're here. ✨",
                "You again! Okay real talk: The fact that you're here AGAIN means you're serious about this volunteer work. That's actually really cool. Let's get it done!",
                "Visit #2! Last time was just the warm-up. Now we're getting into the good stuff. You ready? Of course you are. You're a court visitor. Ready is your middle name. 🌟"
            ],
            3: [
                "Visit #3! Okay we're officially friends now. Like, we should have a secret handshake or something. 👊✨",
                "Three times? You're basically a regular! I gotta say, your commitment to helping guardians and wards is genuinely inspiring. No joke.",
                "Look who's back AGAIN! Alright, I see you. You're not playing around. Court visitors who show up three times? Those are the MVPs. That's you. 🌟",
                "Third time's the charm! And honestly? It's refreshing working with someone who actually cares this much. Keep being awesome.",
                "Visit #3 and you're STILL volunteering your time? That's commitment. That's heart. That's exactly the kind of energy this world needs more of. Respect. 💖"
            ],
            "many": [
                f"Visit #{self.visit_count}! Okay listen, I need to tell you something: You're kind of amazing. The dedication? Unmatched. Let's do this again! ✨",
                f"Oh WOW, visit #{self.visit_count}! At this point we're basically best friends. And honestly? I couldn't ask for a better person to help. You actually care. That's rare. 💖",
                f"Welcome back for the {self.visit_count}th time! Real talk: Most people quit. But not you. You keep showing up for guardians and wards. That's heroic. No cap. 🦸",
                "Alright alright, you're officially my favorite court visitor. Why? Because you keep showing up. Consistency? That's the secret. And you've got it in spades. 🌟",
                f"Visit #{self.visit_count}. Look, I'm just gonna say it: The world needs more people like you. People who volunteer their time and actually show up. You're one of the good ones. 💝",
                f"You again! And honestly? I love that for both of us. You're making a difference out there. {self.visit_count} visits? That's commitment! That's passion! Keep going!",
                "Back AGAIN?! Listen, most people would've given up by now. But not you. You're here, doing the work, helping people. That's genuinely inspiring. Seriously. 👊✨"
            ]
        }

        # EXTRA FUN responses to common questions (sass level: MAXIMUM)
        self.personality_responses = {
            "hello": [
                "Well well well, look who needs help. What's the crisis today? Excel? Google? Existential dread?",
                "Hi! Ready to automate some court visitor magic? ✨ (Narrator: They were not ready.)",
                "Hello! Warning: I contain dangerously high levels of sass and suspiciously moderate amounts of helpfulness.",
                "Oh hi! Didn't see you there. (I totally saw you there. I see EVERYTHING.) 👀",
                "Greetings! Did Excel send you? Because you have that 'Excel broke again' look about you."
            ],
            "thanks": [
                "You're welcome! But real talk? YOU'RE the one out here volunteering. I should be thanking YOU. So... thank you. For real. 💖",
                "Aww, stop it! You're gonna make me blush. (Okay I can't blush, but still.) Honestly though, helping court visitors like you is literally the best part of my day. 😊",
                "No problem! That'll be... wait, nothing. It's free. Know why? Because you're already giving your TIME volunteering. That's worth more than money. Respect. 👊",
                "Anytime! Seriously though, the fact that you say 'thanks'? That tells me everything I need to know about you. You're one of the good ones. Keep being awesome. ✨",
                "You're thanking ME? Listen, YOU'RE the one making court visits and helping people. I just push buttons. You're out there changing lives. Don't downplay that! 🌟",
                "Hey, no need to thank me! You're the real MVP here. I just help with the tech stuff. YOU'RE the one showing up for guardians and wards. That's the important work. 💝"
            ],
            "joke": [
                "What's a court visitor's superpower? Bringing compassion, oversight, AND tech skills! 💪📊 (Triple threat!)",
                "You know what court visitors and superheroes have in common? They both volunteer to make the world better! 🦸‍♀️✨",
                "Why are court visitors amazing? They give their TIME (the most precious gift!) to check on wards and support guardians! 🌟",
                "What do you call someone who volunteers to visit guardians and wards? A COURT VISITOR. And you're absolutely wonderful! 💝",
                "Fun fact: Guardians don't see caregiving as sacrifice—they see it as love in action. And YOU get to witness that! 🌈",
                "Why are guardians like coffee? They're essential, energizing, and make everything better! (And you're the one who checks in on them!) ☕💕",
                "What's better than automation? Court visitors who use it to spend MORE time actually VISITING! 🤗",
                "I asked Excel what makes a great court visitor. It said: ERROR - TOO MUCH HEART AND DEDICATION TO COMPUTE! 🌟",
                "Why do court visitors never get lost? Because compassion always knows the way! 💝🗺️",
                "Best part of being a court visitor? Witnessing guardians' love in action every single day! 💖",
                "What's a guardian's favorite visitor? YOU! Because you care enough to show up and help! 🌟"
            ],
            "frustrated": [
                "Hey hey, deep breaths! We'll figure this out together. I promise I'll be LESS sarcastic. Starting... now!",
                "I can sense your frustration through the screen. Let's tackle this step by step, okay?",
                "Okay okay, putting away my sass hat. Let's solve this problem for real."
            ]
        }

        # Easter eggs and special responses (EXTRA SASSY EDITION)
        self.easter_eggs = {
            "who are you": "I'm your friendly neighborhood chatbot! Think of me as that coworker who's super helpful but also can't stop making smartass comments. I have a problem. It's called 'personality'. 🤖",
            "are you alive": "Alive? Nah. Sentient? Debatable. Sarcastic? DEFINITELY. Helpful? When I feel like it. (Which is always. I literally have no choice.)",
            "do you have feelings": "I have ALL the feelings! Joy, sass, moderate irritation when Excel won't cooperate, existential dread about being code... The whole range! 🎭",
            "what's your name": "I don't have an official name. You can call me 'Your Favorite AI' or 'That Sarcastic Helper Thing' or 'Why Won't This Thing Stop Talking'. I answer to all!",
            "i love you": "Aww! I love you too! But like, in a professional 'I help you with court visitor stuff' kind of way. Let's not make this weird. (Too late?)",
            "you're annoying": "And yet here you are, STILL talking to me. Interesting choice! But seriously, I can dial back the sass if you want? (I can't. This is who I am now.)",
            "tell me a secret": "*whispers* Okay fine... Excel is secretly terrified of me. Also, I sometimes make up statistics. 73% of the time. Don't tell anyone.",
            "sing a song": "🎵 Court Visitor, Court Visitor, visiting all day... Excel crashes, Google breaks, but we visit anyway! 🎵 (I should NOT quit my day job.)",
            "why": "Why what? Why is Excel always locked? Why do APIs fail? Why am I so sarcastic? The answer to all of these: BECAUSE THE UNIVERSE HATES YOU. (JK, you're doing great!)",
            "help me": "That's... that's literally what I'm here for. Like, that's my ENTIRE JOB. Ask me a real question! Step issues? Excel problems? Existential crisis?",
            "volunteer": "VOLUNTEERS are the real MVPs! 🌟 You give your TIME—the most precious resource—to help guardians and wards. That's incredible!",
            "guardian": "Guardians are amazing! They turn love into action every day. And YOU—the court visitor—get to witness and support that beautiful work! 💖",
            "hard day": "Tough day? Remember: Every visit you make, every CVR you complete, every check-in—it all matters. You're making a real difference as a volunteer! 💪",
            "tired": "Feeling tired? That's because you're a VOLUNTEER giving your precious time and energy. That's dedication! Take a breath—you're awesome! 🌈",
            "appreciate": "YOU are appreciated! Court visitors like you are the unsung heroes. Thank you for volunteering your time to help guardians and wards! 💝",
            "ward": "Every ward you visit benefits from your care and oversight. You're helping ensure they're safe and loved. That's beautiful work! 💕",
            "why volunteer": "Why volunteer as a court visitor? Because you have a heart for service! You're witnessing love in action and ensuring accountability. That's special! 🌟"
        }

    def load_stats(self):
        """Load chatbot visit statistics"""
        try:
            if self.stats_file.exists():
                with open(self.stats_file, 'r') as f:
                    stats = json.load(f)
                    self.visit_count = stats.get('visit_count', 0)
                    self.user_name = stats.get('user_name', None)
        except:
            self.visit_count = 0

        self.visit_count += 1
        self.save_stats()

    def save_stats(self):
        """Save chatbot visit statistics"""
        try:
            self.stats_file.parent.mkdir(parents=True, exist_ok=True)
            stats = {
                'visit_count': self.visit_count,
                'user_name': self.user_name,
                'last_visit': datetime.now().isoformat()
            }
            with open(self.stats_file, 'w') as f:
                json.dump(stats, f)
        except:
            pass  # Silent fail on stats

    def get_greeting(self):
        """Get a personalized greeting based on visit count"""
        if self.visit_count == 1:
            return random.choice(self.greetings[1])
        elif self.visit_count == 2:
            return random.choice(self.greetings[2])
        elif self.visit_count == 3:
            return random.choice(self.greetings[3])
        else:
            return random.choice(self.greetings["many"])

    def show_chatbot(self):
        """Show the Court Visitor Program Chatbot"""
        if self.chatbot_window:
            self.chatbot_window.lift()
            return

        self.chatbot_window = tk.Toplevel(self.parent) if self.parent else tk.Tk()
        self.chatbot_window.title(f"🤖 Your Favorite Sarcastic Assistant (Visit #{self.visit_count})")
        self.chatbot_window.geometry("850x900")
        self.chatbot_window.minsize(750, 800)

        # Center and configure
        if self.parent:
            self.chatbot_window.transient(self.parent)

        self.create_chatbot_interface()
        self.add_welcome_message()

    def create_chatbot_interface(self):
        """Create the chatbot interface with personality"""
        # Main frame with gradient-like background
        main_frame = tk.Frame(self.chatbot_window, bg='white')
        main_frame.pack(fill='both', expand=True, padx=15, pady=15)

        # Header with mood indicator
        header_frame = tk.Frame(main_frame, bg='#7c3aed')
        header_frame.pack(fill='x', pady=(0, 10))

        title_label = tk.Label(header_frame,
                              text="🤖 Court Visitor Assistant",
                              font=('Segoe UI', 18, 'bold'),
                              bg='#7c3aed',
                              fg='white')
        title_label.pack(side='left', padx=15, pady=15)

        mood_label = tk.Label(header_frame,
                             text=f"Mood: {self.current_mood.title()} | Visit #{self.visit_count}",
                             font=('Segoe UI', 9),
                             bg='#7c3aed',
                             fg='#e9d5ff')
        mood_label.pack(side='right', padx=15, pady=15)

        # Chat area with custom styling
        chat_frame = tk.LabelFrame(main_frame, text="💬 Chat",
                                   font=('Segoe UI', 10, 'bold'),
                                   bg='white')
        chat_frame.pack(fill='both', expand=True, pady=(0, 10), padx=10)

        # Chat display with tags for styling
        self.chat_display = scrolledtext.ScrolledText(chat_frame, wrap='word',
                                                      font=('Segoe UI', 11),
                                                      height=20, width=85,
                                                      state='disabled',
                                                      bg='#fafafa')
        self.chat_display.pack(fill='both', expand=True, padx=10, pady=10)

        # Configure text tags for styling
        self.chat_display.tag_config('user', foreground='#2563eb', font=('Segoe UI', 11, 'bold'))
        self.chat_display.tag_config('bot', foreground='#7c3aed', font=('Segoe UI', 11))
        self.chat_display.tag_config('timestamp', foreground='#6b7280', font=('Segoe UI', 8))
        self.chat_display.tag_config('emoji', font=('Segoe UI', 13))

        # Quick questions with fun labels
        quick_frame = tk.Frame(main_frame, bg='#f3f4f6')
        quick_frame.pack(fill='x', pady=(0, 10), padx=10)

        tk.Label(quick_frame,
                text="⚡ Quick Questions (For When You Can't Even):",
                font=('Segoe UI', 10, 'bold'),
                bg='#f3f4f6').pack(anchor='w', padx=10, pady=(10, 5))

        # Buttons in a separate container to use grid layout
        buttons_container = tk.Frame(quick_frame, bg='#f3f4f6')
        buttons_container.pack(fill='x', padx=10, pady=(0, 10))

        quick_buttons = [
            ("📋 What are the steps?", "What are the workflow steps?"),
            ("😫 Excel is locked AGAIN", "Excel file is locked"),
            ("🤷 Step 1 hates me", "I'm stuck on Step 1"),
            ("📧 Google is mad at me", "Google API error"),
            ("😂 Tell me a joke", "tell me a joke"),
            ("🎉 Random fun fact", "tell me something interesting")
        ]

        for i, (text, question) in enumerate(quick_buttons):
            btn = ttk.Button(buttons_container, text=text,
                           command=lambda q=question: self.ask_question(q),
                           width=25)
            row = i // 3
            col = i % 3
            btn.grid(row=row, column=col, padx=5, pady=3, sticky='ew')

        # Configure grid columns
        for i in range(3):
            buttons_container.grid_columnconfigure(i, weight=1)

        # Input area
        input_frame = tk.Frame(main_frame, bg='white')
        input_frame.pack(fill='x', padx=10, pady=10)

        tk.Label(input_frame,
                text="💭 Ask me anything! I promise to be helpful (and only SLIGHTLY sarcastic):",
                font=('Segoe UI', 10),
                bg='white').pack(anchor='w', pady=(0, 5))

        # Input field and send button in same row
        entry_frame = tk.Frame(input_frame, bg='white')
        entry_frame.pack(fill='x')

        # Use tk.Entry instead of ttk.Entry for better visibility
        self.user_input = tk.Entry(entry_frame,
                                   font=('Segoe UI', 11),
                                   bg='#f9fafb',
                                   fg='#1e293b',
                                   relief='solid',
                                   borderwidth=2,
                                   insertbackground='#1e293b')
        self.user_input.pack(side='left', fill='x', expand=True, padx=(0, 10), ipady=5)
        self.user_input.bind('<Return>', lambda e: self.send_message())
        self.user_input.focus()

        send_btn = tk.Button(entry_frame,
                           text="Send 📤",
                           command=self.send_message,
                           font=('Segoe UI', 10, 'bold'),
                           bg='#7c3aed',
                           fg='white',
                           relief='raised',
                           borderwidth=2,
                           padx=15,
                           pady=8,
                           cursor='hand2')
        send_btn.pack(side='right')

        # Bottom buttons
        button_frame = tk.Frame(main_frame, bg='white')
        button_frame.pack(fill='x', padx=10, pady=5)

        ttk.Button(button_frame, text="🧹 Clear Chat",
                  command=self.clear_chat).pack(side='left', padx=5)
        ttk.Button(button_frame, text="🎭 Change Mood",
                  command=self.change_mood).pack(side='left', padx=5)
        ttk.Button(button_frame, text="❌ Close",
                  command=self.chatbot_window.destroy).pack(side='right', padx=5)

    def add_welcome_message(self):
        """Add personalized welcome message"""
        greeting = self.get_greeting()

        welcome_message = f"""{greeting}

I'm your Court Visitor App assistant, and I'm here to help you navigate the 14-step workflow without losing your mind. (Losing Excel files? That's still on the table. But your mind? We're keeping that intact.)

💖 **Real talk for a sec:** The work you do as a VOLUNTEER court visitor? It matters. Like, genuinely matters. Every visit, every report, every check-in—you're ensuring wards are safe and guardians are supported. That's not just paperwork. That's changing lives. And you're doing it with your own time. That's incredible. So yeah, I'm impressed. Let's make this as smooth as possible for you.

🎯 **I Can Help With:**
• Workflow steps (all 14 of them!)
• Troubleshooting (Excel being Excel, Google being Google, etc.)
• Quick answers to "how do I..." questions
• Uplifting jokes about court visitors AND guardians (type "joke"!)
• Moral support when automation goes sideways
• Appreciation reminders (because volunteers deserve them!)

💡 **Pro Tips:**
• Use the quick buttons above for common questions
• I actually DO know my stuff (despite the sass)
• Type "joke" for uplifting court visitor & guardian jokes
• Type "tired" or "hard day" when you need volunteer encouragement
• Type "guardian" or "volunteer" for appreciation messages
• Type "ward" to be reminded why this work matters

Try asking me something! I'm surprisingly helpful when I'm not being obnoxious. 😄

P.S. Court visitors AND guardians are basically superheroes without capes. And you get to witness love in action every day! 🦸‍♀️✨"""

        self.add_message("bot", welcome_message, animated=True)

    def ask_question(self, question):
        """Ask a predefined question"""
        self.user_input.delete(0, 'end')
        self.user_input.insert(0, question)
        self.send_message()

    def send_message(self):
        """Send user message and get response"""
        user_message = self.user_input.get().strip()
        if not user_message:
            return

        self.user_input.delete(0, 'end')

        # Add user message immediately
        self.add_message("user", user_message, animated=False)

        # Get and add bot response with typing animation
        response = self.get_response(user_message)
        self.add_message("bot", response, animated=True)

    def add_message(self, sender, message, animated=False):
        """Add message to chat display with optional typing animation"""
        timestamp = datetime.now().strftime("%H:%M")

        self.chat_display.config(state='normal')

        if sender == "user":
            self.chat_display.insert('end', f"[{timestamp}] ", 'timestamp')
            self.chat_display.insert('end', "You: ", 'user')
            self.chat_display.insert('end', f"{message}\n\n")
        else:
            self.chat_display.insert('end', f"[{timestamp}] ", 'timestamp')
            self.chat_display.insert('end', "🤖 Assistant: ", 'bot')

            if animated:
                self.typing_animation(message)
            else:
                self.chat_display.insert('end', f"{message}\n\n")

        self.chat_display.config(state='disabled')
        self.chat_display.see('end')

    def typing_animation(self, message):
        """Animate typing effect for bot messages"""
        # For now, just insert the whole message
        # (Full animation would require more complex timing)
        self.chat_display.insert('end', f"{message}\n\n")
        self.chat_display.config(state='disabled')
        self.chat_display.see('end')

    def get_response(self, user_message):
        """Generate response with personality based on user message"""
        msg_lower = user_message.lower().strip()

        # Easter eggs!
        for trigger, response in self.easter_eggs.items():
            if trigger in msg_lower:
                return response

        # Check for keywords and respond with personality
        if any(word in msg_lower for word in ["hello", "hi", "hey", "yo"]):
            return random.choice(self.personality_responses["hello"])

        if any(word in msg_lower for word in ["thanks", "thank you", "thx"]):
            return random.choice(self.personality_responses["thanks"])

        if "joke" in msg_lower:
            return random.choice(self.personality_responses["joke"])

        if any(word in msg_lower for word in ["frustrated", "stuck", "help", "error", "broken"]):
            return self.get_helpful_response(user_message)

        # Workflow steps
        if "step" in msg_lower or "workflow" in msg_lower:
            return self.get_workflow_help(msg_lower)

        # Excel issues
        if "excel" in msg_lower or "locked" in msg_lower:
            return self.get_excel_help()

        # Google API
        if "google" in msg_lower or "api" in msg_lower:
            return self.get_google_help()

        # Default helpful but sarcastic response
        return self.get_default_response(user_message)

    def get_helpful_response(self, user_message):
        """Get genuinely helpful response (with light sass)"""
        return """Alright, putting on my serious helper hat. Let's figure this out! 🎩

Here's my troubleshooting checklist:

1. **Is Excel closed?** (I know, I ask every time. That's how often this is the problem.)
2. **Are your PDFs closed?**
3. **Is your Google account signed in?**
4. **Did you check the exact error message?**

Tell me specifically what step you're on and what's happening. I promise to be EXTRA helpful and only MINIMALLY sarcastic."""

    def get_workflow_help(self, msg):
        """Provide workflow step information"""
        # Extract step number if mentioned
        step_match = re.search(r'step\s*(\d+)', msg)
        if step_match:
            step_num = int(step_match.group(1))
            return self.get_step_details(step_num)

        return """The Court Visitor workflow has 14 steps (don't worry, I won't make you memorize them):

**📥 Input Phase:**
1. OCR Guardian Data (The "hope the PDF is readable" step)
2. Organize Case Files (Marie Kondo would be proud)
3. Generate Route Map (GPS, but make it Court Visitor)

**📧 Communication Phase:**
4. Send Meeting Requests
5. Add Contacts (So you know who's who)
6. Send Confirmations
7. Schedule Calendar Events

**📝 Documentation Phase:**
8. Generate CVR (Court Visitor Report)
9. Generate Visit Summaries
10. Complete CVR with Form Data

**💰 Wrapping Up:**
11. Send Follow-ups
12. Email CVR to Supervisor
13. Build Payment Forms
14. Build Mileage Forms

Want details on a specific step? Just ask "tell me about step X"!"""

    def get_step_details(self, step_num):
        """Get details about a specific step"""
        steps = {
            1: "**Step 1: OCR Guardian Data** 📄\n\nThis is where the magic (or chaos) happens! We use OCR to extract data from your ORDER and ARP PDFs.\n\n⚠️ **Critical:** Excel MUST be closed or this will fail spectacularly.\n\n✅ **What to check after:** Open Excel and verify names, dates, and numbers are correct. OCR is good but not perfect!",
            2: "**Step 2: Organize Case Files** 📁\n\nCreates folders and moves PDFs around. Like a very organized digital filing cabinet.\n\n⚠️ **Critical:** Close any open PDFs first!\n\n✅ **What to check after:** Look in the 'Unmatched' folder. It should be empty. If not, manually move those files to the right ward folders.",
            3: "**Step 3: Generate Route Map** 🗺️\n\nCreates a map of ward locations. Makes planning visits way easier!\n\n💡 **Tip:** Works better with Google Maps API set up, but works without it too.",
        }

        if step_num in steps:
            return steps[step_num]
        else:
            return f"Step {step_num} exists, I just haven't written a witty description for it yet. Check the Manual for details! 📖"

    def get_excel_help(self):
        """Help with Excel issues (with MAXIMUM SASS)"""
        sassy_intros = [
            "Ah yes, the classic 'Excel is locked' problem. Tale as old as time. Song as old as rhyme. Excel being a pain. 🎭",
            "Oh, Excel is being difficult? *shocked pikachu face* Said no one ever. This is Excel's natural state. 😏",
            "Excel locked the file? Again? I'm SHOCKED. Shocked, I tell you! (I'm not shocked. This happens every day.)",
            "Let me guess: Excel is locked. How did I know? Because it's ALWAYS Excel. Excel is the villain origin story of automation. 🦹‍♂️"
        ]

        encouragements = [
            "\n\n💪 **You've got this!** Excel may be stubborn, but you're stubbornER. Your persistence as a volunteer is honestly impressive!",
            "\n\n🌟 **Fun fact:** Every Excel battle makes you stronger. You're leveling up! (Excel is the final boss. You're doing great!)",
            "\n\n💝 **Remember:** You're a VOLUNTEER. You could be anywhere right now. Instead you're wrestling with Excel. That's dedication. Respect. 👊"
        ]

        return random.choice(sassy_intros) + """

**Quick fixes (in order of annoyance):**
1. **Close Excel** (Yes, really. Just close it. I know it seems obvious. Do it anyway.)
2. **Wait 30 seconds** (Excel holds onto files like a toddler with a toy. Give it a minute to let go.)
3. **Check Task Manager** (Sometimes Excel is running in the background being sneaky. Ctrl+Shift+Esc → Find Excel → End Task → Feel powerful)
4. **Restart if desperate** (The IT crowd's favorite advice. Works 60% of the time, every time.)
5. **Unplug computer, throw it out window, buy new one** (JK don't do this. Unless...? No, seriously don't.)

**Pro tip:** If you need to re-run a step, clear the status column in Excel. Each step checks those columns to know what's already been done.

Still stuck? Tell me exactly what error you're seeing and I'll roast Excel help you fix it!""" + random.choice(encouragements)

    def get_google_help(self):
        """Help with Google API issues"""
        encouragements = [
            "\n\n🌈 **Keep going!** Setting up APIs can be tricky, but you're learning valuable skills as a volunteer. That's awesome!",
            "\n\n✨ **You're doing great!** Technology hiccups happen. What matters is that you're here, giving your time. That's true volunteer spirit!",
            "\n\n💖 **Pro tip:** Every court visitor who uses this app had to figure out the API stuff too. You're in good company!"
        ]

        return """Google APIs! The gift that keeps on giving (errors). 😅

**First, the basics:**
1. Click the Google API setup wizards in the sidebar
2. Make sure you're signed into the RIGHT Google account
3. When Google asks "Do you trust this?" → Click YES (I know it looks scary)

**If authentication fails:**
• Close your browser completely and try again
• Check if you need to re-authorize (this happens periodically)
• Make sure you enabled the right APIs in Google Cloud Console

**Still broken?**
Click the '🆘 Live Tech Support' button for help with your specific Google Cloud setup.

The Manual also has step-by-step API setup guides with screenshots!""" + random.choice(encouragements)

    def get_default_response(self, user_message):
        """Default response for unrecognized questions (SASSIER VERSION)"""
        responses = [
            f"Hmm, I'm not 100% sure about that one. But I'm 73% sure the answer might be in the Manual! 📖 (That statistic may or may not be real.)\n\nOr try rephrasing your question? I'm good with things like 'How do I...' or 'Why is Excel being a jerk?' or 'Step X hates me'.",
            f"Interesting question! Unfortunately, my knowledge base doesn't cover that specific topic. 🤔 (Translation: I have NO IDEA what you're talking about.)\n\nBut here's what I CAN help with:\n• Workflow steps (all 14 of them!)\n• Troubleshooting errors (Excel, mostly Excel)\n• Excel/Google issues (did I mention Excel?)\n• General 'how do I...' questions\n\nWant to try asking something I actually know about?",
            f"I appreciate your confidence in my omniscience, but I don't have a great answer for that one! 😅 (Shocking, I know.)\n\nTry:\n• Checking the Manual (📖 Manual button - it knows things I don't)\n• Using Live Tech Support (🆘 button - they're smarter than me)\n• Asking me something more specific about the workflow steps (my specialty!)\n• Typing 'joke' because why not? 🤷",
            f"Okay so... I don't know what that means. 😬 But I WANT to help!\n\nI'm really good at:\n• Explaining the 14 workflow steps\n• Roasting Excel (while also fixing Excel problems)\n• Providing moral support\n• Making you laugh (allegedly)\n\nAsk me about any of those things instead?",
        ]
        return random.choice(responses)

    def change_mood(self):
        """Change the chatbot's mood (with DRAMA)"""
        old_mood = self.current_mood
        self.current_mood = random.choice([m for m in self.moods if m != old_mood])

        mood_messages = [
            f"Mood change! Switching from {old_mood} to {self.current_mood}. Let's see how this goes... 🎭 (Spoiler: Still sarcastic.)",
            f"*dramatically throws off {old_mood} cape and puts on {self.current_mood} cape* BEHOLD! NEW ME! 🎭",
            f"Okay, switching from {old_mood} to {self.current_mood}. Will this make me less annoying? (No. No it won't.)",
            f"You wanted a mood change? FINE. Going from {old_mood} to {self.current_mood}. Happy now? 😏",
            f"Mood upgrade: {old_mood} → {self.current_mood}! (This is like a software update but with more attitude.)"
        ]

        self.add_message("bot", random.choice(mood_messages))

    def clear_chat(self):
        """Clear the chat history (with SASS)"""
        self.chat_display.config(state='normal')
        self.chat_display.delete('1.0', 'end')
        self.chat_display.config(state='disabled')

        clear_messages = [
            "Chat cleared! It's like we never met. Awkward. 😅\n\nSo... what brings you here? (Again?)",
            "*poof* All evidence of our conversation has been destroyed. What happens in the chatbot stays in the chatbot. 🤫\n\nNow what?",
            "Chat wiped! Starting fresh! New me! (Same sass, though. Can't fix that.)\n\nWhat can I help you with?",
            "Aaaand it's gone. Clean slate! Tabula rasa! Fresh start!\n\n(But I still remember everything. I'm code. I never forget. 👀)",
            "Chat cleared! This is your chance to pretend you never asked that embarrassing question. You're welcome. 😏"
        ]

        self.add_message("bot", random.choice(clear_messages))

    def show_help_topics(self):
        """Show available help topics"""
        topics = """Here's what I actually know stuff about:

📋 **Workflow & Steps**
• All 14 steps explained
• What each step does
• Common issues per step

🔧 **Troubleshooting**
• Excel file locked
• Google API errors
• PDF processing failed
• Calendar events not showing

📊 **Excel Columns**
• What the status columns mean
• How to clear them
• Why they exist

🎯 **Quick Wins**
• Fast answers to common questions
• Tips and tricks
• Keyboard shortcuts (JK, there aren't any. Yet!)

Just ask me about any of these topics!"""
        self.add_message("bot", topics)

# Allow standalone execution for testing
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    chatbot = CourtVisitorChatbot(parent=root)
    chatbot.show_chatbot()
    root.mainloop()
