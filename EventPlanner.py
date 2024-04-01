import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import Calendar
import datetime
import os
import time
import base64
from email.mime.text import MIMEText
import threading
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from requests import HTTPError
import os.path
import json
import requests
import sys
import webbrowser

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class Event:
    def __init__(self, title, description, date, emails, notify_date = None, sent = False):
        self.title = title
        self.description = description
        self.date = date
        self.emails = emails
        self.notify_date = notify_date  # Add a field to store the date and time of the notification
        self.sent = sent
        
    def serialize(self):
        # Converting the event object to a dictionary
        event_dict = {
            "title": self.title,
            "description": self.description,
            "date": self.date.strftime("%Y-%m-%d %H:%M:%S"),  # Convert date to string
            "emails": self.emails,
            "notify_date": self.notify_date.strftime("%Y-%m-%d %H:%M:%S") if self.notify_date else None,  # Convert the notification date to a string
            "sent": self.sent
        }
        return event_dict
    
    @classmethod
    def deserialize(cls, event_dict):
        # Create an event object from the dictionary
        date = datetime.datetime.strptime(event_dict["date"], "%Y-%m-%d %H:%M:%S")
        notify_date = datetime.datetime.strptime(event_dict["notify_date"], "%Y-%m-%d %H:%M:%S") if event_dict["notify_date"] else None
        return cls(event_dict["title"], event_dict["description"], date, event_dict["emails"], notify_date, event_dict["sent"])
    
"""
This is the official English localization.
"""

class EventPlannerApp:
    def __init__(self, root):
        self.load_fonts()
        self.run_notification_loop()
        
        self.root = root
        self.root.title("Event Planner")
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        # Bind the event save function to the window close event
        self.root.protocol("WM_DELETE_WINDOW", self.close_application)
        
        # Check if theme file exists
        if os.path.exists("current_theme.txt"):
            with open("current_theme.txt", "r") as f:
                self.current_theme = f.read()
        else:
            # Set the default theme if the file does not exist
            self.current_theme = "light"
        
        self.style.configure("Yellow.TButton",
                        foreground="#323232",
                        background="#FCF7C9",
                        font=("Segoe UI Semibold", 12),
                        padding=10,
                        )
        self.style.map("Yellow.TButton",
                foreground=[("active", "white")],
                background=[("active", "#323232")],
                )

        self.main_frame = tk.Frame(self.root, bg='white')
        self.main_frame.pack(expand=True, fill="both")
        
        self.toolbar_frame = tk.Frame(self.main_frame, bg='white')
        self.toolbar_frame.pack(side="top", fill="x")

        self.new_event_button = ttk.Button(self.toolbar_frame, text="New event", style="Yellow.TButton", command=self.create_event_window)
        self.new_event_button.pack(side="left")

        self.settings_button = ttk.Button(self.toolbar_frame, text="Settings", style="Yellow.TButton", command=self.open_settings)
        self.settings_button.pack(side="right")
        
        self.save_button = ttk.Button(self.toolbar_frame, text="Save", style="Yellow.TButton", command=self.save_events_to_file)
        self.save_button.pack(side="right", padx=5)  # Add a small gap between the buttons
        
        ttk.Separator(self.main_frame, orient="horizontal").pack(fill="x")

        self.paned_window = ttk.PanedWindow(self.main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(expand=True, fill="both")

        self.events_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(self.events_frame, weight=1)

        self.events_listbox = tk.Listbox(self.events_frame, selectmode="single")
        self.events_listbox.pack(expand=True, fill="both")

        self.events_listbox.bind("<Button-3>", self.show_context_menu)
        self.events_listbox.bind("<<ListboxSelect>>", self.on_event_selected)

        self.context_menu = tk.Menu(self.root, tearoff=0, bg='white')
        self.context_menu.add_command(label="Edit", command=self.edit_event)
        self.context_menu.add_command(label="Delete", command=self.delete_event)
        
        ttk.Separator(self.paned_window, orient="vertical").pack(side="left", fill="y")

        self.details_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(self.details_frame, weight=2)

        self.details_text = tk.Text(self.details_frame, wrap="word", font=("Segoe UI", 12), spacing1=8, spacing2=8, spacing3=8, bg='white')
        self.details_text.config(foreground="#262626")

        scrollbar = tk.Scrollbar(self.details_frame, orient="vertical", command=self.details_text.yview)
        scrollbar.pack(side="right", fill="y")

        self.details_text.config(yscrollcommand=scrollbar.set)

        self.details_text.pack(expand=True, fill="both")

        self.paned_window.bind("<B1-Motion>")
        
        # Load events from a file if the file exists
        if os.path.exists("events.json") and os.path.getsize("events.json") > 0:
            with open("events.json", "r") as f:
                events_data = json.load(f)
            self.events = [Event.deserialize(data) for data in events_data]
            self.update_events_listbox()
        else:
            self.events = []

        if not self.events:
            self.show_event_details(-1)
            
        if self.current_theme == "dark":
            self.change_theme()
            
        if os.path.exists("current_scaling.txt"):
            with open("current_scaling.txt", "r") as f:
                self.current_scaling = float(f.read())
        else:
            self.current_scaling = 1.33
            
        # Bind zoom functions to keypress events
        self.root.bind("<Control-equal>", self.zoom_in)
        self.root.bind("<Control-minus>", self.zoom_out)
            
    def load_fonts(self):
        font_folder = resource_path("fonts")
        if not os.path.exists(font_folder):
            os.makedirs(font_folder)
        
        font_files = {
            "Segoe UI Black": "SEGUIBL.TTF",
            "Segoe UI": "SEGOEUI.TTF",
            "Segoe UI Semibold": "SEGUISB.TTF",
            "Segoe UI Italic": "SEGOEUII.TTF",
            "Segoe UI Semibold Italic": "SEGUISBI.TTF"
        }

        self.loaded_fonts = {}
        for font_name, font_file in font_files.items():
            font_path = os.path.join(font_folder, font_file)
            if os.path.exists(font_path):
                self.loaded_fonts[font_name] = font_path
                
    def is_google_account_authenticated(self):
        return os.path.exists('credentials.json')
    
    def save_events_to_file(self):
        with open("events.json", "w") as f:
            json.dump([event.serialize() for event in self.events], f)
            
    def close_application(self):
        # Saving events before closing the application
        self.save_events_to_file()
        # Close the application
        self.root.destroy()
        
    def zoom_in(self, event):
        response = messagebox.askquestion("Warning", "You want to zoom in on the interface. To do this, the application will be restarted. Continue?")
        if response == 'yes':                
            new_scaling = self.current_scaling + 0.33
            
            # Save the current theme to a file
            with open("current_scaling.txt", "w") as f:
                f.write(str(new_scaling))
                
            self.save_events_to_file
                
            # Restart application
            python = sys.executable
            os.execl(python, python, *sys.argv)
    def zoom_out(self, event):
        response = messagebox.askquestion("Warning", "You want to zoom out on the interface. To do this, the application will be restarted. Continue?")
        if response == 'yes':
            new_scaling = self.current_scaling - 0.33
            
            # Save the current theme to a file
            with open("current_scaling.txt", "w") as f:
                f.write(str(new_scaling))
                
            self.save_events_to_file
                
            # Restart application
            python = sys.executable
            os.execl(python, python, *sys.argv)

    def create_event_window(self):
        create_window = tk.Toplevel(self.root)
        create_window.title("Creation of event")
        create_window.geometry(f"{int(800/(1.33/self.current_scaling))}x{int(600/(1.33/self.current_scaling))}")
        create_window.iconphoto(False, tk.PhotoImage(file=resource_path('EP_cover.png')))

        padx = 10
        pady = 5

        if self.current_theme == 'light':
            canvas = tk.Canvas(create_window, bg='white')
        elif self.current_theme == 'dark':
            canvas = tk.Canvas(create_window, bg='#323232')
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(create_window, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        if self.current_theme == 'light':
            content_frame = tk.Frame(canvas, bg='white')
        elif self.current_theme == 'dark':
            content_frame = tk.Frame(canvas, bg='#323232')
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        
        def update_scroll_region(event):
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
            
        def on_mousewheel(event):
            if event.widget == description_entry:
                return
            canvas.yview_scroll(-1*(event.delta//120), "units")

        content_frame.bind("<Configure>", update_scroll_region)
        create_window.bind("<MouseWheel>", on_mousewheel)

        ttk.Label(content_frame, text="Event name:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
        if self.current_theme == 'light':
            title_entry = tk.Entry(content_frame, font=("Segoe UI Black", 16), bg='white', fg='black')
        elif self.current_theme == 'dark':
            title_entry = tk.Entry(content_frame, font=("Segoe UI Black", 16), bg='#414141', fg='white')
        title_entry.pack(pady=pady, padx=padx, fill="x")

        ttk.Label(content_frame, text="Event description:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")

        # Creating a frame for the text field and scrollbar
        if self.current_theme == 'light':
            description_frame = tk.Frame(content_frame, bg='white')
        elif self.current_theme == 'dark':
            description_frame = tk.Frame(content_frame, bg='#323232')
        description_frame.pack(fill="both", expand=True)

        # Create a field to describe the event
        if self.current_theme == 'light':
            description_entry = tk.Text(description_frame, wrap="word", height=10, font=("Segoe UI", 12), bg='white', fg='black')
        elif self.current_theme == 'dark':
            description_entry = tk.Text(description_frame, wrap="word", height=10, font=("Segoe UI", 12), bg='#414141', fg='white')

        # Adding a vertical scrollbar
        scrollbar = ttk.Scrollbar(description_frame, orient="vertical", command=description_entry.yview)
        scrollbar.pack(side="right", fill="y")

        # Linking a scrollbar to a text field
        description_entry.config(yscrollcommand=scrollbar.set)

        description_entry.pack(pady=pady, padx=padx, fill="both", expand=True)

        ttk.Label(content_frame, text="Event date and time:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
        if self.current_theme == 'light':
            date_time_frame = tk.Frame(content_frame, bg='white')
        elif self.current_theme == 'dark':
            date_time_frame = tk.Frame(content_frame, bg='#414141')
        date_time_frame.pack(pady=pady, padx=padx, fill="x")

        calendar = Calendar(date_time_frame, selectmode="day", year=datetime.datetime.now().year, month=datetime.datetime.now().month, day=datetime.datetime.now().day)
        calendar.pack(side="left", padx=(0, 10))

        if self.current_theme == 'light':
            time_frame = tk.Frame(date_time_frame, bg='white')
        elif self.current_theme == 'dark':
            time_frame = tk.Frame(date_time_frame, bg='#414141')
        time_frame.pack(side="left")

        hour_var = tk.StringVar(value=str(datetime.datetime.now().hour).zfill(2))
        minute_var = tk.StringVar(value=str(datetime.datetime.now().minute).zfill(2))

        ttk.Label(time_frame, text="Hours:", font=("Segoe UI", 14)).pack(anchor="w")
        hour_spinbox = ttk.Spinbox(time_frame, from_=0, to=23, textvariable=hour_var, width=2, font=("Segoe UI", 14))
        hour_spinbox.pack(anchor="w", pady=(0, 5))

        ttk.Label(time_frame, text="Minutes:", font=("Segoe UI", 14)).pack(anchor="w")
        minute_spinbox = ttk.Spinbox(time_frame, from_=0, to=59, textvariable=minute_var, width=2, font=("Segoe UI", 14))
        minute_spinbox.pack(anchor="w")

        ttk.Label(content_frame, text="List of emails (comma-separated):", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
        if self.current_theme == 'light':
            emails_entry = tk.Entry(content_frame, font=("Segoe UI Semibold", 12), bg='white', fg='black')
        elif self.current_theme == 'dark':
            emails_entry = tk.Entry(content_frame, font=("Segoe UI Semibold", 12), bg='#414141', fg='white')
        emails_entry.pack(pady=pady, padx=padx, fill="x")

        notify_var = tk.BooleanVar(value=False)

        def toggle_notify():
            if notify_var.get():
                notify_frame.pack(side="top", fill="x", padx=padx, anchor="w", pady=(10, 10))
            else:
                notify_frame.pack_forget()
                
            update_scroll_region(None)

        notify_checkbox = ttk.Checkbutton(content_frame, text="Notify", variable=notify_var, command=toggle_notify)
        notify_checkbox.pack(pady=pady, padx=padx, anchor="w")

        if self.current_theme == 'light':
            notify_frame = tk.Frame(content_frame, bg='white')
        elif self.current_theme == 'dark':
            notify_frame = tk.Frame(content_frame, bg='#414141')
        notify_frame.pack(pady=pady, padx=padx, fill="x", anchor="w")

        # Check if the user is signed in to a Google account
        if self.is_google_account_authenticated():
            # If you are logged in, then we allow the use of notifications
            notify_checkbox.config(state="normal")
        else:
            # If you are not logged in, then make the widget inactive
            notify_checkbox.config(state="disabled")

            # Create a label indicating that you need to sign in to your Google account
            notify_label = ttk.Label(content_frame, text="To use notifications, please sign in to your Google Account.", foreground="red")
            notify_label.pack(pady=(0, 10), padx=padx, anchor="w")

        notify_calendar = Calendar(notify_frame, selectmode="day", year=datetime.datetime.now().year, month=datetime.datetime.now().month, day=datetime.datetime.now().day)
        notify_calendar.pack(side="left", padx=(0, 10))

        if self.current_theme == 'light':
            notify_time_frame = tk.Frame(notify_frame, bg='white')
        elif self.current_theme == 'dark':
            notify_time_frame = tk.Frame(notify_frame, bg='#414141')
        notify_time_frame.pack(side="left")

        notify_hour_var = tk.StringVar(value=str(datetime.datetime.now().hour).zfill(2))
        notify_minute_var = tk.StringVar(value=str(datetime.datetime.now().minute).zfill(2))

        ttk.Label(notify_time_frame, text="Hours:", font=("Segoe UI", 14)).pack(anchor="w")
        notify_hour_spinbox = ttk.Spinbox(notify_time_frame, from_=0, to=23, textvariable=notify_hour_var, width=2, font=("Segoe UI", 14))
        notify_hour_spinbox.pack(anchor="w", pady=(0, 5))

        ttk.Label(notify_time_frame, text="Minutes:", font=("Segoe UI", 14)).pack(anchor="w")
        notify_minute_spinbox = ttk.Spinbox(notify_time_frame, from_=0, to=59, textvariable=notify_minute_var, width=2, font=("Segoe UI", 14))
        notify_minute_spinbox.pack(anchor="w")

        toggle_notify()  # Hide/show notification widgets depending on initial state

        def save_event():
            try:
                title = title_entry.get()
                description = description_entry.get("1.0", tk.END).strip()
                selected_date = calendar.get_date()
                hour = int(hour_var.get())
                minute = int(minute_var.get())
                date = datetime.datetime.strptime(selected_date, '%m/%d/%y').replace(hour=hour, minute=minute)
                emails = [email.strip() for email in emails_entry.get().split(',')]
                notify_date = None
                if notify_var.get():
                    notify_selected_date = notify_calendar.get_date()
                    notify_hour = int(notify_hour_var.get())
                    notify_minute = int(notify_minute_var.get())
                    notify_date = datetime.datetime.strptime(notify_selected_date, '%m/%d/%y').replace(hour=notify_hour, minute=notify_minute)

                new_event = Event(title, description, date, emails)
                new_event.notify_date = notify_date  # Save the notification date

                self.events.append(new_event)

                self.update_events_listbox()

                create_window.destroy()

                self.show_event_details(len(self.events) - 1)
            except ValueError as e:
                messagebox.showerror("Error", f"Incorrect data: {e}")

        save_button = ttk.Button(content_frame, style="Yellow.TButton", text="Save", command=save_event)
        save_button.pack(pady=10, padx=padx, side="right")
  
    def create_event_widgets(self, index, formatted_event):
        label = tk.Label(self.events_listbox, text=formatted_event, bg='#FCF7C9', fg='#323232', font=("Segoe UI Semibold", 12))
        label.pack(fill=tk.X)

        label.bind("<Button-1>", lambda e, idx=index: self.show_event_details(idx))
        
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Edit", command=lambda idx=index: self.edit_event(idx))
        context_menu.add_command(label="Delete", command=lambda idx=index: self.delete_event(idx))
        label.bind("<Button-3>", lambda e, menu=context_menu: menu.post(e.x_root, e.y_root))

    def update_events_listbox(self):
        for widget in self.events_listbox.winfo_children():
            widget.destroy()

        for index, event in enumerate(self.events):
            formatted_event = f"{event.title} - {event.date.strftime('%Y-%m-%d %H:%M:%S')}"
            self.create_event_widgets(index, formatted_event)

    def show_event_details(self, event_index):
        if event_index is not None and 0 <= event_index < len(self.events):
            selected_event = self.events[event_index]
            
            self.details_text.tag_configure("date", font=("Segoe UI Semibold", 14), foreground="red")
            if self.current_theme == 'light':
                self.details_text.tag_configure("title", font=("Segoe UI Black", 16), foreground="black")
                self.details_text.tag_configure("description", font=("Segoe UI", 12), foreground="#262626")
                self.details_text.tag_configure("emails", font=("Segoe UI Semibold", 12), foreground="#262626")
                self.details_text.tag_configure("status_not_sent_ok", font=("Segoe UI Black", 10), foreground="black")
            elif self.current_theme == 'dark':
                self.details_text.tag_configure("title", font=("Segoe UI Black", 16), foreground="white")
                self.details_text.tag_configure("description", font=("Segoe UI", 12), foreground="white")
                self.details_text.tag_configure("emails", font=("Segoe UI Semibold", 12), foreground="white")
                self.details_text.tag_configure("status_not_sent_ok", font=("Segoe UI Black", 10), foreground="white")
            
            self.details_text.tag_configure("status_sent", font=("Segoe UI Black", 10), foreground="green")
            self.details_text.tag_configure("status_not_sent_bad", font=("Segoe UI Black", 10), foreground="red")

            details_text_lines = [
                f"Name: {selected_event.title}",
                f"Description: {selected_event.description}",
                f"Date: {selected_event.date.strftime('%Y-%m-%d %H:%M:%S')}",
                f"Emails: {', '.join(selected_event.emails)}"
            ]
            
            if selected_event.notify_date:  # Check if notification date exists
                details_text_lines.append(f"Notification date: {selected_event.notify_date.strftime('%Y-%m-%d %H:%M:%S')}")
                
                # Checking the notification sending status
                if selected_event.sent:
                    details_text_lines.append("The notification was sent successfully!")
                else:
                    if datetime.datetime.now() < selected_event.notify_date + datetime.timedelta(minutes=1):
                        details_text_lines.append("Notification will be sent when notification time arrives.")
                    else:
                        details_text_lines.append("The notification was not sent because an error occurred. "
                                                "Check your Internet connection and the correctness of the entered data. "
                                                "A notification will be sent as soon as the issue is resolved.")

            self.details_text.config(state="normal")
            self.details_text.delete("1.0", tk.END)

            for line in details_text_lines:
                self.details_text.insert(tk.END, line + "\n")

                if line.startswith("Name:"):
                    start_index = self.details_text.search("Name:", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("title", start_index, end_index)
                elif line.startswith("Date:"):
                    start_index = self.details_text.search("Date:", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("date", start_index, end_index)
                elif line.startswith("Emails:"):
                    start_index = self.details_text.search("Emails:", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("emails", start_index, end_index)
                elif line.startswith("Notification date:"):  # Check if a string matches the notification date
                    start_index = self.details_text.search("Notification date:", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("date", start_index, end_index)
                    
                elif line.startswith("The notification was sent"):
                    start_index = self.details_text.search("The notification was sent", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("status_sent", start_index, end_index)
                elif line.startswith("Notification will"):
                    start_index = self.details_text.search("Notification will", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("status_not_sent_ok", start_index, end_index)
                elif line.startswith("The notification was not"):
                    start_index = self.details_text.search("The notification was not", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("status_not_sent_bad", start_index, end_index)
                    
                else:
                    start_index = self.details_text.search(line, "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("description", start_index, end_index)

            self.details_text.config(state="disabled")

            for widget in self.events_listbox.winfo_children():
                if self.current_theme == 'light':
                    widget.config(bg='#FCF7C9', fg='#323232')
                elif self.current_theme == 'dark':
                    widget.config(bg='#323232', fg='white')
            if self.current_theme == 'light':
                self.events_listbox.winfo_children()[event_index].config(bg='#323232', fg='white')
            elif self.current_theme == 'dark':
                self.events_listbox.winfo_children()[event_index].config(bg='#FCF7C9', fg='#323232')
        else:
            self.details_text.config(state="normal")
            self.details_text.delete("1.0", tk.END)
            self.details_text.insert(tk.END, "Select an event to display its details.")
            self.details_text.tag_add("center", "1.0", "end")
            self.details_text.config(state="disabled")
            
    def on_event_selected(self, event):
        selected_index = self.events_listbox.curselection()
        if selected_index:
            self.show_event_details(selected_index[0])
            
    def edit_event(self, event_index=None):
        if event_index is not None and 0 <= event_index < len(self.events):
            selected_event = self.events[event_index]

            # Creating an editing window
            edit_window = tk.Toplevel(self.root)
            edit_window.title("Editing of event")
            edit_window.geometry(f"{int(800/(1.33/self.current_scaling))}x{int(600/(1.33/self.current_scaling))}")
            edit_window.iconphoto(False, tk.PhotoImage(file=resource_path('EP_cover.png')))

            # Form elements for editing
            padx = 10
            pady = 5

            if self.current_theme == 'light':
                canvas = tk.Canvas(edit_window, bg='white')
            elif self.current_theme == 'dark':
                canvas = tk.Canvas(edit_window, bg='#323232')
            canvas.pack(side="left", fill="both", expand=True)

            scrollbar = tk.Scrollbar(edit_window, orient="vertical", command=canvas.yview)
            scrollbar.pack(side="right", fill="y")

            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

            if self.current_theme == 'light':
                content_frame = tk.Frame(canvas, bg='white')
            elif self.current_theme == 'dark':
                content_frame = tk.Frame(canvas, bg='#323232')
            canvas.create_window((0, 0), window=content_frame, anchor="nw")

            def update_scroll_region(event):
                canvas.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))
                
            def on_mousewheel(event):
                if event.widget == description_entry:
                    return
                canvas.yview_scroll(-1*(event.delta//120), "units")

            content_frame.bind("<Configure>", update_scroll_region)
            edit_window.bind("<MouseWheel>", on_mousewheel)

            ttk.Label(content_frame, text="Event name:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
            if self.current_theme == 'light':
                title_entry = tk.Entry(content_frame, font=("Segoe UI Black", 16), bg='white', fg='black')
            elif self.current_theme == 'dark':
                title_entry = tk.Entry(content_frame, font=("Segoe UI Black", 16), bg='#414141', fg='white')
            title_entry.insert(0, selected_event.title)
            title_entry.pack(pady=pady, padx=padx, fill="x")

            ttk.Label(content_frame, text="Event description:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")

            # Creating a frame for the text field and scrollbar
            if self.current_theme == 'light':
                description_frame = tk.Frame(content_frame, bg='white')
            elif self.current_theme == 'dark':
                description_frame = tk.Frame(content_frame, bg='#323232')
            description_frame.pack(fill="both", expand=True)

            # Create a field to describe the event
            if self.current_theme == 'light':
                description_entry = tk.Text(description_frame, wrap="word", height=10, font=("Segoe UI", 12), bg='white', fg='black')
            elif self.current_theme == 'dark':
                description_entry = tk.Text(description_frame, wrap="word", height=10, font=("Segoe UI", 12), bg='#414141', fg='white')
            description_entry.insert("1.0", selected_event.description)

            # Adding a vertical scrollbar
            scrollbar = ttk.Scrollbar(description_frame, orient="vertical", command=description_entry.yview)
            scrollbar.pack(side="right", fill="y")

            # Linking a scrollbar to a text field
            description_entry.config(yscrollcommand=scrollbar.set)

            description_entry.pack(pady=pady, padx=padx, fill="both", expand=True)


            ttk.Label(content_frame, text="Event date and time:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
            if self.current_theme == 'light':
                date_time_frame = tk.Frame(content_frame, bg='white')
            elif self.current_theme == 'dark':
                date_time_frame = tk.Frame(content_frame, bg='#414141')
            date_time_frame.pack(pady=pady, padx=padx, fill="x")

            calendar = Calendar(date_time_frame, selectmode="day", year=selected_event.date.year, month=selected_event.date.month, day=selected_event.date.day)
            calendar.pack(side="left", padx=(0, 10))

            if self.current_theme == 'light':
                time_frame = tk.Frame(date_time_frame, bg='white')
            elif self.current_theme == 'dark':
                time_frame = tk.Frame(date_time_frame, bg='#414141')
            time_frame.pack(side="left")

            hour_var = tk.StringVar(value=str(selected_event.date.hour).zfill(2))
            minute_var = tk.StringVar(value=str(selected_event.date.minute).zfill(2))

            ttk.Label(time_frame, text="Hours:", font=("Segoe UI", 14)).pack(anchor="w")
            hour_spinbox = ttk.Spinbox(time_frame, from_=0, to=23, textvariable=hour_var, width=2, font=("Segoe UI", 14))
            hour_spinbox.pack(anchor="w", pady=(0, 5))

            ttk.Label(time_frame, text="Minutes:", font=("Segoe UI", 14)).pack(anchor="w")
            minute_spinbox = ttk.Spinbox(time_frame, from_=0, to=59, textvariable=minute_var, width=2, font=("Segoe UI", 14))
            minute_spinbox.pack(anchor="w")

            ttk.Label(content_frame, text="List of emails (comma-separated):", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
            if self.current_theme == 'light':
                emails_entry = tk.Entry(content_frame, font=("Segoe UI Semibold", 12), bg='white', fg='black')
            elif self.current_theme == 'dark':
                emails_entry = tk.Entry(content_frame, font=("Segoe UI Semibold", 12), bg='#414141', fg='white')
            emails_entry.insert(0, ', '.join(selected_event.emails))
            emails_entry.pack(pady=pady, padx=padx, fill="x")
            
            notify_var = tk.BooleanVar(value=selected_event.notify_date is not None)  # Set the checkbox value depending on the presence of a notification

            def toggle_notify():
                if notify_var.get():
                    notify_frame.pack(side="top", fill="x", padx=padx, anchor="w", pady=(10, 10))
                else:
                    notify_frame.pack_forget()

                update_scroll_region(None)

            notify_checkbox = ttk.Checkbutton(content_frame, text="Notify", variable=notify_var, command=toggle_notify)
            notify_checkbox.pack(pady=pady, padx=padx, anchor="w")

            if self.current_theme == 'light':
                notify_frame = tk.Frame(content_frame, bg='white')
            elif self.current_theme == 'dark':
                notify_frame = tk.Frame(content_frame, bg='#414141')
            notify_frame.pack(pady=pady, padx=padx, fill="x", anchor="w")

            # Check if the user is signed in to a Google account
            if self.is_google_account_authenticated():
                # If you are logged in, then we allow the use of notifications
                notify_checkbox.config(state="normal")
            else:
                # If you are not logged in, then make the widget inactive
                notify_checkbox.config(state="disabled")

                # Create a label indicating that you need to sign in to your Google account
                notify_label = ttk.Label(content_frame, text="To use notifications, please sign in to your Google Account.", foreground="red")
                notify_label.pack(pady=(0, 10), padx=padx, anchor="w")

            if selected_event.notify_date:
                notify_calendar = Calendar(notify_frame, selectmode="day", year=selected_event.notify_date.year, month=selected_event.notify_date.month, day=selected_event.notify_date.day)
            else:
                notify_calendar = Calendar(notify_frame, selectmode="day", year=datetime.datetime.now().year, month=datetime.datetime.now().month, day=datetime.datetime.now().day)
            notify_calendar.pack(side="left", padx=(0, 10))

            if self.current_theme == 'light':
                notify_time_frame = tk.Frame(notify_frame, bg='white')
            elif self.current_theme == 'dark':
                notify_time_frame = tk.Frame(notify_frame, bg='#414141')
            notify_time_frame.pack(side="left")

            notify_hour_var = tk.StringVar(value=str(selected_event.notify_date.hour).zfill(2) if selected_event.notify_date else str(datetime.datetime.now().hour).zfill(2))
            notify_minute_var = tk.StringVar(value=str(selected_event.notify_date.minute).zfill(2) if selected_event.notify_date else str(datetime.datetime.now().minute).zfill(2))

            ttk.Label(notify_time_frame, text="Hours:", font=("Segoe UI", 14)).pack(anchor="w")
            notify_hour_spinbox = ttk.Spinbox(notify_time_frame, from_=0, to=23, textvariable=notify_hour_var, width=2, font=("Segoe UI", 14))
            notify_hour_spinbox.pack(anchor="w", pady=(0, 5))

            ttk.Label(notify_time_frame, text="Minutes:", font=("Segoe UI", 14)).pack(anchor="w")
            notify_minute_spinbox = ttk.Spinbox(notify_time_frame, from_=0, to=59, textvariable=notify_minute_var, width=2, font=("Segoe UI", 14))
            notify_minute_spinbox.pack(anchor="w")

            toggle_notify()  # Hide/show notification widgets depending on the initial state of the checkbox

            def save_changes():
                try:
                    edited_date_str = calendar.get_date()
                    edited_date = datetime.datetime.strptime(edited_date_str, "%m/%d/%y")
                    edited_date = edited_date.replace(hour=int(hour_var.get()), minute=int(minute_var.get()))
                    edited_emails = [email.strip() for email in emails_entry.get().split(',')]
                    edited_event = Event(title_entry.get(), description_entry.get("1.0", tk.END).strip(), edited_date, edited_emails)

                    edited_event.notify_date = None  # Set the notification date to None by default

                    if notify_var.get():
                        notify_selected_date = notify_calendar.get_date()
                        notify_hour = int(notify_hour_var.get())
                        notify_minute = int(notify_minute_var.get())
                        edited_event.notify_date = datetime.datetime.strptime(notify_selected_date, '%m/%d/%y').replace(hour=notify_hour, minute=notify_minute)

                    self.events[event_index] = edited_event

                    self.update_events_listbox()

                    edit_window.destroy()

                    self.show_event_details(event_index)
                except ValueError as e:
                    messagebox.showerror("Error", f"Incorrect data: {e}")                    
        # Create a button to save changes
            save_button = ttk.Button(content_frame, style="Yellow.TButton", text="Save changes", command=save_changes)
            save_button.pack(pady=10, padx=padx, side="right")
        else:
            messagebox.showinfo("Error", "Select an event to edit.")

    def delete_event(self, event_index=None):
        if event_index is not None and 0 <= event_index < len(self.events):
            del self.events[event_index]

            self.update_events_listbox()

            self.show_event_details(-1)
        else:
            messagebox.showinfo("Error", "Select an event to delete.")

    def show_context_menu(self, event):
        self.context_menu.post(event.x_root, event.y_root)
        
    def send_notifications(self):
        try:
            # Loading credentials from a file
            with open('credentials.json', 'r') as file:
                creds_data = json.load(file)
            # Converting a string to a dictionary
            creds_data = json.loads(creds_data)
            creds = Credentials(
                token=creds_data['token'],
                refresh_token=creds_data['refresh_token'],
                token_uri=creds_data['token_uri'],
                client_id=creds_data['client_id'],
                client_secret=creds_data['client_secret']
            )
            
            # Creating a Credential Object
            name_credentials = Credentials.from_authorized_user_info(creds_data)
            # Checking if the access token needs updating
            if name_credentials.expired and name_credentials.refresh_token:
                name_credentials.refresh(Request())
            
            service = build('gmail', 'v1', credentials=creds)

            # Checking all user events
            for event in self.events:
                # Checking whether the event has a notification and whether its time has come
                if event.notify_date and datetime.datetime.now() >= event.notify_date and not event.sent:
                    url = 'https://www.googleapis.com/oauth2/v3/userinfo'
                    headers = {'Authorization': f'Bearer {name_credentials.token}'}
                    response = requests.get(url, headers=headers)
                    if response.status_code == 200:
                        data = response.json()
                        name = data.get('name')
                    
                    # Создаем сообщение
                    message = MIMEText(f"{event.description}\nFrom {name} using Event Planner.")
                    message['to'] = ", ".join(event.emails)
                    message['subject'] = f"Reminder that event {event.title} will start on {event.date}!"

                    create_message = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

                    try:
                        message = (service.users().messages().send(userId="me", body=create_message).execute())
                        print(F'sent message to {message} Message Id: {message["id"]}')
                        
                        # Mark the event as sent
                        event.sent = True
                    except HTTPError as error:
                        print(F'An error occurred: {error}')
                        message = None

        except Exception as e:
            print(f"Error sending notifications: {e}")
        
    def run_notification_loop(self):
        def notification_thread():
            while True:
                self.send_notifications()
                time.sleep(60)

        thread = threading.Thread(target=notification_thread)
        thread.daemon = True
        thread.start()

    def open_settings(self):
        settings_window = SettingsWindow(self, self.current_theme)
        
    def change_theme(self):
        # Changing the colors of controls depending on the selected theme
        if self.current_theme == "dark":  # Dark theme
            # Colors for dark theme
            background_color = "#414141"
            details_text_color = 'white'
            listbox_background_color = "#323232"
            listbox_item_background_color = "#323232"
            listbox_item_foreground_color = "white"
            button_background_color = "#323232"
            button_foreground_color = "white"
            button_active_background_color = "#FCF7C9"
            button_active_foreground_color = "#323232"
        else: # Light theme
            # Light Theme Colors
            background_color = "white"
            details_text_color = '#262626'
            listbox_background_color = "white"
            listbox_item_background_color = "#FCF7C9"
            listbox_item_foreground_color = "#323232"
            button_background_color = "#FCF7C9"
            button_foreground_color = "#323232"
            button_active_background_color = "#323232"
            button_active_foreground_color = "white"

        # Applying colors to controls
        self.root.configure(bg=background_color)
        self.toolbar_frame.configure(bg=background_color)
        self.events_listbox.configure(bg=listbox_background_color)
        for widget in self.events_listbox.winfo_children():
                widget.config(bg=listbox_item_background_color, fg=listbox_item_foreground_color)
        self.context_menu.configure(bg=background_color, fg=button_foreground_color)
        self.details_text.configure(bg=background_color, fg=details_text_color)
        self.style.configure("Yellow.TButton", background=button_background_color, foreground=button_foreground_color)
        self.style.map("Yellow.TButton", background=[("active", button_active_background_color)], foreground=[("active", button_active_foreground_color)])

    def run(self):
        self.root.mainloop()
        
class SettingsWindow:
    def __init__(self, parent, current_theme):
        self.parent = parent
        self.settings_window = tk.Toplevel(parent.root)
        self.settings_window.title("Settings")
        self.settings_window.iconphoto(False, tk.PhotoImage(file=resource_path('EP_cover.png')))
        
        self.current_theme = current_theme
        
        if self.current_theme == 'light':
            self.settings_window.configure(bg='white')
        elif self.current_theme == 'dark':
            self.settings_window.configure(bg='#414141')
        
        self.credentials = None
        
        self.create_widgets()

    def create_widgets(self):
        # Adding Controls
        if self.current_theme == 'light':
            self.settings_frame = tk.Frame(self.settings_window, bg='white')
        elif self.current_theme == 'dark':
            self.settings_frame = tk.Frame(self.settings_window, bg='#414141')
        self.settings_frame.pack(pady=5, padx=10)

        # Google Account Zone
        if self.current_theme == 'light':
            self.google_account_frame = tk.Frame(self.settings_frame, bg='white')
        elif self.current_theme == 'dark':
            self.google_account_frame = tk.Frame(self.settings_frame, bg='#414141')
        self.google_account_frame.pack(side="left", padx=10)

        tk.Label(self.google_account_frame, text="Google account settings:", font=("Segoe UI Semibold", 12)).pack(pady=10, anchor="w")

        # Checking if Google Account data is saved
        if os.path.exists('credentials.json'):
            # Loading credentials from a file
            with open('credentials.json', 'r') as file:
                creds_data = json.load(file)
            # Converting a string to a dictionary
            creds_data = json.loads(creds_data)
            # Creating a Credential Object
            self.credentials = Credentials.from_authorized_user_info(creds_data)
            # Checking if credentials are valid
            if not self.credentials.valid:
                # Checking if the access token needs updating
                if self.credentials.expired and self.credentials.refresh_token:
                    self.credentials.refresh(Request())
                    # Display Google account information
                    email, name = self.get_user_info()
                    if email and name:
                        ttk.Label(self.google_account_frame, text="Email:", font=("Segoe UI", 12)).pack(anchor="w")
                        ttk.Label(self.google_account_frame, text=email, font=("Segoe UI", 12)).pack(anchor="w")
                        ttk.Label(self.google_account_frame, text="Username:", font=("Segoe UI", 12)).pack(anchor="w")
                        ttk.Label(self.google_account_frame, text=name, font=("Segoe UI", 12)).pack(anchor="w")
                        # Logout button
                        ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Log out of your account", command=self.logout_google_account).pack(pady=10, anchor="w")
                    else:
                        # If we were unable to obtain information, we display the login button
                        ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Sign in to Google Account", command=self.login_google_account).pack(pady=10, anchor="w")
                else:
                    # Credentials are invalid, you need to log in again
                    self.login_google_account()
            else:
                # Display Google account information
                email, name = self.get_user_info()
                if email and name:
                    ttk.Label(self.google_account_frame, text="Email:", font=("Segoe UI", 12)).pack(anchor="w")
                    ttk.Label(self.google_account_frame, text=email, font=("Segoe UI", 12)).pack(anchor="w")
                    ttk.Label(self.google_account_frame, text="Username:", font=("Segoe UI", 12)).pack(anchor="w")
                    ttk.Label(self.google_account_frame, text=name, font=("Segoe UI", 12)).pack(anchor="w")
                    # Logout button
                    ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Log out of your account", command=self.logout_google_account).pack(pady=10, anchor="w")
                else:
                    # If we were unable to obtain information, we display the login button
                    ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Sign in to Google Account", command=self.login_google_account).pack(pady=10, anchor="w")
        else:
            # If there are no credentials, we display the login button
            ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Sign in to Google Account", command=self.login_google_account).pack(pady=10, anchor="w")

        # Zone of buttons for changing theme and language
        if self.current_theme == 'light':
            self.theme_language_frame = tk.Frame(self.settings_frame, bg='white')
        elif self.current_theme == 'dark':            
            self.theme_language_frame = tk.Frame(self.settings_frame, bg='#414141')
        self.theme_language_frame.pack(side="right", padx=10)
        
        # Adding help button centered in the zone
        ttk.Button(self.theme_language_frame, style="Yellow.TButton", text="Help", command=lambda: webbrowser.open("https://docs.google.com/document/d/1yQYKMG--Q4hUG8daiSD0xcQOntGQ_f_nzOiG34x_KQE/edit?usp=sharing")).pack(side="top", padx=5)

        # Adding buttons to change theme and language
        ttk.Button(self.theme_language_frame, style="Yellow.TButton", text="Change color theme", command=self.save_and_change_theme).pack(side="left", padx=5)
        ttk.Button(self.theme_language_frame, style="Yellow.TButton", text="Change language", command=self.change_language).pack(side="bottom", pady=10)
        
        tk.Label(self.settings_window, text="© 2024 Hlib Ishchenko. All rights reserved.", font=("Segoe UI", 12)).pack(pady=10, side="bottom")

    def get_user_info(self):
        url = 'https://www.googleapis.com/oauth2/v3/userinfo'
        headers = {'Authorization': f'Bearer {self.credentials.token}'}
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            email = data.get('email')
            name = data.get('name')
            return email, name
        else:
            print('Error:', response.status_code)
            return None, None

    def login_google_account(self):
        os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'
        # Creating OAuth2 credentials to access the Gmail API and retrieve user information
        scopes = [
            'https://www.googleapis.com/auth/gmail.send',
            'https://www.googleapis.com/auth/userinfo.email',
            'https://www.googleapis.com/auth/userinfo.profile'
        ]
        flow = InstalledAppFlow.from_client_secrets_file(resource_path('your_client_secret.json'), scopes=scopes)
        self.credentials = flow.run_local_server(port=0)

        # Saving credentials to a file
        with open('credentials.json', 'w') as file:
            json.dump(self.credentials.to_json(), file)
        
        # Rebooting the settings window
        self.settings_window.destroy()
        self.__init__(self.parent)

    def logout_google_account(self):
        # Deleting files with credentials
        if os.path.exists('credentials.json'):
            os.remove('credentials.json')

        # Rebooting the settings window
        self.settings_window.destroy()
        self.__init__(self.parent)
        
    def save_and_change_theme(self):
        response = messagebox.askquestion("Warning", "The application will be restarted to change the theme. Continue?")
        if response == 'yes':
            new_theme = "dark" if self.current_theme == "light" else "light"
            
            # Save the current theme to a file
            with open("current_theme.txt", "w") as f:
                f.write(new_theme)
                
            self.parent.save_events_to_file()
            # Restart application
            python = sys.executable
            os.execl(python, python, *sys.argv)

    def show(self):
        self.settings_window.deiconify()

    def hide(self):
        self.settings_window.withdraw()
    
    def change_language(self):
        response = messagebox.askquestion("Warning", "The language will change to Ukrainian. The application will be restarted to change the language. Continue?")
        if response == 'yes':
            if os.path.exists("current_language.txt"):
                with open("current_language.txt", "r") as f:
                    current_language = f.read()
            else:
                current_language = "english"
                
            new_language = "ukrainian" if current_language == "english" else "english"
            
            # Save the current theme to a file
            with open("current_language.txt", "w") as f:
                f.write(new_language)
                
            self.parent.save_events_to_file()
            # Restart application
            python = sys.executable
            os.execl(python, python, *sys.argv)
            
"""
This is the official Ukrainian localization.
"""
            
class EventPlannerAppUKR:
    def __init__(self, root):
        self.load_fonts()
        self.run_notification_loop()
        
        self.root = root
        self.root.title("Event Planner")
        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        self.root.protocol("WM_DELETE_WINDOW", self.close_application)
        
        if os.path.exists("current_theme.txt"):
            with open("current_theme.txt", "r") as f:
                self.current_theme = f.read()
        else:
            self.current_theme = "light"
        
        self.style.configure("Yellow.TButton",
                        foreground="#323232",
                        background="#FCF7C9",
                        font=("Segoe UI Semibold", 12),
                        padding=10,
                        )
        self.style.map("Yellow.TButton",
                foreground=[("active", "white")],
                background=[("active", "#323232")],
                )

        self.main_frame = tk.Frame(self.root, bg='white')
        self.main_frame.pack(expand=True, fill="both")
        
        self.toolbar_frame = tk.Frame(self.main_frame, bg='white')
        self.toolbar_frame.pack(side="top", fill="x")

        self.new_event_button = ttk.Button(self.toolbar_frame, text="Нова подія", style="Yellow.TButton", command=self.create_event_window)
        self.new_event_button.pack(side="left")

        self.settings_button = ttk.Button(self.toolbar_frame, text="Налаштування", style="Yellow.TButton", command=self.open_settings)
        self.settings_button.pack(side="right")
        
        self.save_button = ttk.Button(self.toolbar_frame, text="Зберегти", style="Yellow.TButton", command=self.save_events_to_file)
        self.save_button.pack(side="right", padx=5)  # Add a small gap between the buttons
        
        ttk.Separator(self.main_frame, orient="horizontal").pack(fill="x")

        self.paned_window = ttk.PanedWindow(self.main_frame, orient=tk.HORIZONTAL)
        self.paned_window.pack(expand=True, fill="both")

        self.events_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(self.events_frame, weight=1)

        self.events_listbox = tk.Listbox(self.events_frame, selectmode="single")
        self.events_listbox.pack(expand=True, fill="both")

        self.events_listbox.bind("<Button-3>", self.show_context_menu)
        self.events_listbox.bind("<<ListboxSelect>>", self.on_event_selected)

        self.context_menu = tk.Menu(self.root, tearoff=0, bg='white')
        self.context_menu.add_command(label="Редагувати", command=self.edit_event)
        self.context_menu.add_command(label="Видалити", command=self.delete_event)
        
        ttk.Separator(self.paned_window, orient="vertical").pack(side="left", fill="y")

        self.details_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(self.details_frame, weight=2)

        self.details_text = tk.Text(self.details_frame, wrap="word", font=("Segoe UI", 12), spacing1=8, spacing2=8, spacing3=8, bg='white')
        self.details_text.config(foreground="#262626")

        scrollbar = tk.Scrollbar(self.details_frame, orient="vertical", command=self.details_text.yview)
        scrollbar.pack(side="right", fill="y")

        self.details_text.config(yscrollcommand=scrollbar.set)

        self.details_text.pack(expand=True, fill="both")

        self.paned_window.bind("<B1-Motion>")
        
        if os.path.exists("events.json") and os.path.getsize("events.json") > 0:
            with open("events.json", "r") as f:
                events_data = json.load(f)
            self.events = [Event.deserialize(data) for data in events_data]
            self.update_events_listbox()
        else:
            self.events = []

        if not self.events:
            self.show_event_details(-1)
            
        if self.current_theme == "dark":
            self.change_theme()
            
        if os.path.exists("current_scaling.txt"):
            with open("current_scaling.txt", "r") as f:
                self.current_scaling = float(f.read())
        else:
            self.current_scaling = 1.33
            
        # Bind zoom functions to keypress events
        self.root.bind("<Control-equal>", self.zoom_in)
        self.root.bind("<Control-minus>", self.zoom_out)
            
    def load_fonts(self):
        font_folder = "fonts"
        if not os.path.exists(font_folder):
            os.makedirs(font_folder)
        
        font_files = {
            "Segoe UI Black": "SEGUIBL.TTF",
            "Segoe UI": "SEGOEUI.TTF",
            "Segoe UI Semibold": "SEGUISB.TTF",
            "Segoe UI Italic": "SEGOEUII.TTF",
            "Segoe UI Semibold Italic": "SEGUISBI.TTF"
        }

        self.loaded_fonts = {}
        for font_name, font_file in font_files.items():
            font_path = os.path.join(font_folder, font_file)
            if os.path.exists(font_path):
                self.loaded_fonts[font_name] = font_path
                
    def is_google_account_authenticated(self):
        return os.path.exists('credentials.json')
    
    def save_events_to_file(self):
        with open("events.json", "w") as f:
            json.dump([event.serialize() for event in self.events], f)
            
    def close_application(self):
        self.save_events_to_file()
        self.root.destroy()
        
    def zoom_in(self, event):
        response = messagebox.askquestion("Увага", "Ви хочете збільшити масштаб інтерфейсу. Для цього програму буде перезапущено. Продовжити?")
        if response == 'yes':                
            new_scaling = self.current_scaling + 0.33
            
            # Save the current theme to a file
            with open("current_scaling.txt", "w") as f:
                f.write(str(new_scaling))
                
            self.save_events_to_file
                
            # Restart application
            python = sys.executable
            os.execl(python, python, *sys.argv)
    def zoom_out(self, event):
        response = messagebox.askquestion("Увага", "Ви хочете зменшити масштаб інтерфейсу. Для цього програму буде перезапущено. Продовжити?")
        if response == 'yes':
            new_scaling = self.current_scaling - 0.33
            
            # Save the current theme to a file
            with open("current_scaling.txt", "w") as f:
                f.write(str(new_scaling))
                
            self.save_events_to_file
                
            # Restart application
            python = sys.executable
            os.execl(python, python, *sys.argv)

    def create_event_window(self):
        create_window = tk.Toplevel(self.root)
        create_window.title("Створення події")
        create_window.geometry(f"{int(800/(1.33/self.current_scaling))}x{int(600/(1.33/self.current_scaling))}")
        create_window.iconphoto(False, tk.PhotoImage(file=resource_path('EP_cover.png')))

        padx = 10
        pady = 5

        if self.current_theme == 'light':
            canvas = tk.Canvas(create_window, bg='white')
        elif self.current_theme == 'dark':
            canvas = tk.Canvas(create_window, bg='#323232')
        canvas.pack(side="left", fill="both", expand=True)

        scrollbar = tk.Scrollbar(create_window, orient="vertical", command=canvas.yview)
        scrollbar.pack(side="right", fill="y")

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        if self.current_theme == 'light':
            content_frame = tk.Frame(canvas, bg='white')
        elif self.current_theme == 'dark':
            content_frame = tk.Frame(canvas, bg='#323232')
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        
        def update_scroll_region(event):
            canvas.update_idletasks()
            canvas.configure(scrollregion=canvas.bbox("all"))
            
        def on_mousewheel(event):
            if event.widget == description_entry:
                return
            canvas.yview_scroll(-1*(event.delta//120), "units")

        content_frame.bind("<Configure>", update_scroll_region)
        create_window.bind("<MouseWheel>", on_mousewheel)

        ttk.Label(content_frame, text="Назва події:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
        if self.current_theme == 'light':
            title_entry = tk.Entry(content_frame, font=("Segoe UI Black", 16), bg='white', fg='black')
        elif self.current_theme == 'dark':
            title_entry = tk.Entry(content_frame, font=("Segoe UI Black", 16), bg='#414141', fg='white')
        title_entry.pack(pady=pady, padx=padx, fill="x")

        ttk.Label(content_frame, text="Опис події:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")

        # Creating a frame for the text field and scrollbar
        if self.current_theme == 'light':
            description_frame = tk.Frame(content_frame, bg='white')
        elif self.current_theme == 'dark':
            description_frame = tk.Frame(content_frame, bg='#323232')
        description_frame.pack(fill="both", expand=True)

        # Create a field to describe the event
        if self.current_theme == 'light':
            description_entry = tk.Text(description_frame, wrap="word", height=10, font=("Segoe UI", 12), bg='white', fg='black')
        elif self.current_theme == 'dark':
            description_entry = tk.Text(description_frame, wrap="word", height=10, font=("Segoe UI", 12), bg='#414141', fg='white')

        # Adding a vertical scrollbar
        scrollbar = ttk.Scrollbar(description_frame, orient="vertical", command=description_entry.yview)
        scrollbar.pack(side="right", fill="y")

        # Linking a scrollbar to a text field
        description_entry.config(yscrollcommand=scrollbar.set)

        description_entry.pack(pady=pady, padx=padx, fill="both", expand=True)

        ttk.Label(content_frame, text="Дата та час події:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
        if self.current_theme == 'light':
            date_time_frame = tk.Frame(content_frame, bg='white')
        elif self.current_theme == 'dark':
            date_time_frame = tk.Frame(content_frame, bg='#414141')
        date_time_frame.pack(pady=pady, padx=padx, fill="x")

        calendar = Calendar(date_time_frame, selectmode="day", year=datetime.datetime.now().year, month=datetime.datetime.now().month, day=datetime.datetime.now().day)
        calendar.pack(side="left", padx=(0, 10))

        if self.current_theme == 'light':
            time_frame = tk.Frame(date_time_frame, bg='white')
        elif self.current_theme == 'dark':
            time_frame = tk.Frame(date_time_frame, bg='#414141')
        time_frame.pack(side="left")

        hour_var = tk.StringVar(value=str(datetime.datetime.now().hour).zfill(2))
        minute_var = tk.StringVar(value=str(datetime.datetime.now().minute).zfill(2))

        ttk.Label(time_frame, text="Години:", font=("Segoe UI", 14)).pack(anchor="w")
        hour_spinbox = ttk.Spinbox(time_frame, from_=0, to=23, textvariable=hour_var, width=2, font=("Segoe UI", 14))
        hour_spinbox.pack(anchor="w", pady=(0, 5))

        ttk.Label(time_frame, text="Хвилини:", font=("Segoe UI", 14)).pack(anchor="w")
        minute_spinbox = ttk.Spinbox(time_frame, from_=0, to=59, textvariable=minute_var, width=2, font=("Segoe UI", 14))
        minute_spinbox.pack(anchor="w")

        ttk.Label(content_frame, text="Список електронних адрес (через кому):", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
        if self.current_theme == 'light':
            emails_entry = tk.Entry(content_frame, font=("Segoe UI Semibold", 12), bg='white', fg='black')
        elif self.current_theme == 'dark':
            emails_entry = tk.Entry(content_frame, font=("Segoe UI Semibold", 12), bg='#414141', fg='white')
        emails_entry.pack(pady=pady, padx=padx, fill="x")

        notify_var = tk.BooleanVar(value=False)

        def toggle_notify():
            if notify_var.get():
                notify_frame.pack(side="top", fill="x", padx=padx, anchor="w", pady=(10, 10))
            else:
                notify_frame.pack_forget()
                
            update_scroll_region(None)

        notify_checkbox = ttk.Checkbutton(content_frame, text="Нагадати", variable=notify_var, command=toggle_notify)
        notify_checkbox.pack(pady=pady, padx=padx, anchor="w")

        if self.current_theme == 'light':
            notify_frame = tk.Frame(content_frame, bg='white')
        elif self.current_theme == 'dark':
            notify_frame = tk.Frame(content_frame, bg='#414141')
        notify_frame.pack(pady=pady, padx=padx, fill="x", anchor="w")

        # Check if the user is signed in to a Google account
        if self.is_google_account_authenticated():
            # If you are logged in, then we allow the use of notifications
            notify_checkbox.config(state="normal")
        else:
            # If you are not logged in, then make the widget inactive
            notify_checkbox.config(state="disabled")

            # Create a label indicating that you need to sign in to your Google account
            notify_label = ttk.Label(content_frame, text="Щоб користуватися нагадуваннями, увійдіть у свій обліковий запис Google.", foreground="red")
            notify_label.pack(pady=(0, 10), padx=padx, anchor="w")

        notify_calendar = Calendar(notify_frame, selectmode="day", year=datetime.datetime.now().year, month=datetime.datetime.now().month, day=datetime.datetime.now().day)
        notify_calendar.pack(side="left", padx=(0, 10))

        if self.current_theme == 'light':
            notify_time_frame = tk.Frame(notify_frame, bg='white')
        elif self.current_theme == 'dark':
            notify_time_frame = tk.Frame(notify_frame, bg='#414141')
        notify_time_frame.pack(side="left")

        notify_hour_var = tk.StringVar(value=str(datetime.datetime.now().hour).zfill(2))
        notify_minute_var = tk.StringVar(value=str(datetime.datetime.now().minute).zfill(2))

        ttk.Label(notify_time_frame, text="Години:", font=("Segoe UI", 14)).pack(anchor="w")
        notify_hour_spinbox = ttk.Spinbox(notify_time_frame, from_=0, to=23, textvariable=notify_hour_var, width=2, font=("Segoe UI", 14))
        notify_hour_spinbox.pack(anchor="w", pady=(0, 5))

        ttk.Label(notify_time_frame, text="Хвилини:", font=("Segoe UI", 14)).pack(anchor="w")
        notify_minute_spinbox = ttk.Spinbox(notify_time_frame, from_=0, to=59, textvariable=notify_minute_var, width=2, font=("Segoe UI", 14))
        notify_minute_spinbox.pack(anchor="w")

        toggle_notify()  # Hide/show notification widgets depending on initial state

        def save_event():
            try:
                title = title_entry.get()
                description = description_entry.get("1.0", tk.END).strip()
                selected_date = calendar.get_date()
                hour = int(hour_var.get())
                minute = int(minute_var.get())
                date = datetime.datetime.strptime(selected_date, '%m/%d/%y').replace(hour=hour, minute=minute)
                emails = [email.strip() for email in emails_entry.get().split(',')]
                notify_date = None
                if notify_var.get():
                    notify_selected_date = notify_calendar.get_date()
                    notify_hour = int(notify_hour_var.get())
                    notify_minute = int(notify_minute_var.get())
                    notify_date = datetime.datetime.strptime(notify_selected_date, '%m/%d/%y').replace(hour=notify_hour, minute=notify_minute)

                new_event = Event(title, description, date, emails)
                new_event.notify_date = notify_date  # Save the notification date

                self.events.append(new_event)

                self.update_events_listbox()

                create_window.destroy()

                self.show_event_details(len(self.events) - 1)
            except ValueError as e:
                messagebox.showerror("Помилка", f"Невірні дані: {e}")

        save_button = ttk.Button(content_frame, style="Yellow.TButton", text="Зберегти", command=save_event)
        save_button.pack(pady=10, padx=padx, side="right")
  
    def create_event_widgets(self, index, formatted_event):
        label = tk.Label(self.events_listbox, text=formatted_event, bg='#FCF7C9', fg='#323232', font=("Segoe UI Semibold", 12))
        label.pack(fill=tk.X)

        label.bind("<Button-1>", lambda e, idx=index: self.show_event_details(idx))
        
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="Редагувати", command=lambda idx=index: self.edit_event(idx))
        context_menu.add_command(label="Видалити", command=lambda idx=index: self.delete_event(idx))
        label.bind("<Button-3>", lambda e, menu=context_menu: menu.post(e.x_root, e.y_root))

    def update_events_listbox(self):
        for widget in self.events_listbox.winfo_children():
            widget.destroy()

        for index, event in enumerate(self.events):
            formatted_event = f"{event.title} - {event.date.strftime('%Y-%m-%d %H:%M:%S')}"
            self.create_event_widgets(index, formatted_event)

    def show_event_details(self, event_index):
        if event_index is not None and 0 <= event_index < len(self.events):
            selected_event = self.events[event_index]
            
            self.details_text.tag_configure("date", font=("Segoe UI Semibold", 14), foreground="red")
            if self.current_theme == 'light':
                self.details_text.tag_configure("title", font=("Segoe UI Black", 16), foreground="black")
                self.details_text.tag_configure("description", font=("Segoe UI", 12), foreground="#262626")
                self.details_text.tag_configure("emails", font=("Segoe UI Semibold", 12), foreground="#262626")
                self.details_text.tag_configure("status_not_sent_ok", font=("Segoe UI Black", 10), foreground="black")
            elif self.current_theme == 'dark':
                self.details_text.tag_configure("title", font=("Segoe UI Black", 16), foreground="white")
                self.details_text.tag_configure("description", font=("Segoe UI", 12), foreground="white")
                self.details_text.tag_configure("emails", font=("Segoe UI Semibold", 12), foreground="white")
                self.details_text.tag_configure("status_not_sent_ok", font=("Segoe UI Black", 10), foreground="white")
            
            self.details_text.tag_configure("status_sent", font=("Segoe UI Black", 10), foreground="green")
            self.details_text.tag_configure("status_not_sent_bad", font=("Segoe UI Black", 10), foreground="red")

            details_text_lines = [
                f"Назва: {selected_event.title}",
                f"Опис: {selected_event.description}",
                f"Дата: {selected_event.date.strftime('%Y-%m-%d %H:%M:%S')}",
                f"Електронні адреси: {', '.join(selected_event.emails)}"
            ]
            
            if selected_event.notify_date:  # Check if notification date exists
                details_text_lines.append(f"Дата нагадування: {selected_event.notify_date.strftime('%Y-%m-%d %H:%M:%S')}")
                
                # Checking the notification sending status
                if selected_event.sent:
                    details_text_lines.append("Нагадування успішно надіслано!")
                else:
                    if datetime.datetime.now() < selected_event.notify_date + datetime.timedelta(minutes=1):
                        details_text_lines.append("Нагадування буде надіслано, коли настане час сповіщення.")
                    else:
                        details_text_lines.append("Нагадування не надіслано через помилку. "
                                                "Перевірте підключення до Інтернету та правильність введених даних. "
                                                "Нагадування буде надіслано, щойно проблему буде вирішено.")

            self.details_text.config(state="normal")
            self.details_text.delete("1.0", tk.END)

            for line in details_text_lines:
                self.details_text.insert(tk.END, line + "\n")

                if line.startswith("Назва:"):
                    start_index = self.details_text.search("Назва:", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("title", start_index, end_index)
                elif line.startswith("Дата:"):
                    start_index = self.details_text.search("Дата:", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("date", start_index, end_index)
                elif line.startswith("Електронні адреси:"):
                    start_index = self.details_text.search("Електронні адреси:", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("emails", start_index, end_index)
                elif line.startswith("Дата нагадування:"):  # Check if a string matches the notification date
                    start_index = self.details_text.search("Дата нагадування:", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("date", start_index, end_index)
                    
                elif line.startswith("Нагадуваання успішно"):
                    start_index = self.details_text.search("Нагадуваання успішно", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("status_sent", start_index, end_index)
                elif line.startswith("Нагадування буде"):
                    start_index = self.details_text.search("Нагадування буде", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("status_not_sent_ok", start_index, end_index)
                elif line.startswith("Нагадування не"):
                    start_index = self.details_text.search("Нагадування не", "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("status_not_sent_bad", start_index, end_index)
                    
                else:
                    start_index = self.details_text.search(line, "1.0", tk.END)
                    end_index = self.details_text.index(f"{start_index} lineend")
                    self.details_text.tag_add("description", start_index, end_index)

            self.details_text.config(state="disabled")

            for widget in self.events_listbox.winfo_children():
                if self.current_theme == 'light':
                    widget.config(bg='#FCF7C9', fg='#323232')
                elif self.current_theme == 'dark':
                    widget.config(bg='#323232', fg='white')
            if self.current_theme == 'light':
                self.events_listbox.winfo_children()[event_index].config(bg='#323232', fg='white')
            elif self.current_theme == 'dark':
                self.events_listbox.winfo_children()[event_index].config(bg='#FCF7C9', fg='#323232')
        else:
            self.details_text.config(state="normal")
            self.details_text.delete("1.0", tk.END)
            self.details_text.insert(tk.END, "Виберіть подію, щоб відобразити її деталі.")
            self.details_text.tag_add("center", "1.0", "end")
            self.details_text.config(state="disabled")
            
    def on_event_selected(self, event):
        selected_index = self.events_listbox.curselection()
        if selected_index:
            self.show_event_details(selected_index[0])
            
    def edit_event(self, event_index=None):
        if event_index is not None and 0 <= event_index < len(self.events):
            selected_event = self.events[event_index]

            # Creating an editing window
            edit_window = tk.Toplevel(self.root)
            edit_window.title("Редагування події")
            edit_window.geometry(f"{int(800/(1.33/self.current_scaling))}x{int(600/(1.33/self.current_scaling))}")
            edit_window.iconphoto(False, tk.PhotoImage(file=resource_path('EP_cover.png')))

            # Form elements for editing
            padx = 10
            pady = 5

            if self.current_theme == 'light':
                canvas = tk.Canvas(edit_window, bg='white')
            elif self.current_theme == 'dark':
                canvas = tk.Canvas(edit_window, bg='#323232')
            canvas.pack(side="left", fill="both", expand=True)

            scrollbar = tk.Scrollbar(edit_window, orient="vertical", command=canvas.yview)
            scrollbar.pack(side="right", fill="y")

            canvas.configure(yscrollcommand=scrollbar.set)
            canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

            if self.current_theme == 'light':
                content_frame = tk.Frame(canvas, bg='white')
            elif self.current_theme == 'dark':
                content_frame = tk.Frame(canvas, bg='#323232')
            canvas.create_window((0, 0), window=content_frame, anchor="nw")

            def update_scroll_region(event):
                canvas.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))
                
            def on_mousewheel(event):
                if event.widget == description_entry:
                    return
                canvas.yview_scroll(-1*(event.delta//120), "units")

            content_frame.bind("<Configure>", update_scroll_region)
            edit_window.bind("<MouseWheel>", on_mousewheel)

            ttk.Label(content_frame, text="Назва події:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
            if self.current_theme == 'light':
                title_entry = tk.Entry(content_frame, font=("Segoe UI Black", 16), bg='white', fg='black')
            elif self.current_theme == 'dark':
                title_entry = tk.Entry(content_frame, font=("Segoe UI Black", 16), bg='#414141', fg='white')
            title_entry.insert(0, selected_event.title)
            title_entry.pack(pady=pady, padx=padx, fill="x")

            ttk.Label(content_frame, text="Опис події:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")

            # Creating a frame for the text field and scrollbar
            if self.current_theme == 'light':
                description_frame = tk.Frame(content_frame, bg='white')
            elif self.current_theme == 'dark':
                description_frame = tk.Frame(content_frame, bg='#323232')
            description_frame.pack(fill="both", expand=True)

            # Create a field to describe the event
            if self.current_theme == 'light':
                description_entry = tk.Text(description_frame, wrap="word", height=10, font=("Segoe UI", 12), bg='white', fg='black')
            elif self.current_theme == 'dark':
                description_entry = tk.Text(description_frame, wrap="word", height=10, font=("Segoe UI", 12), bg='#414141', fg='white')
            description_entry.insert("1.0", selected_event.description)

            # Adding a vertical scrollbar
            scrollbar = ttk.Scrollbar(description_frame, orient="vertical", command=description_entry.yview)
            scrollbar.pack(side="right", fill="y")

            # Linking a scrollbar to a text field
            description_entry.config(yscrollcommand=scrollbar.set)

            description_entry.pack(pady=pady, padx=padx, fill="both", expand=True)


            ttk.Label(content_frame, text="Дата і час події:", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
            if self.current_theme == 'light':
                date_time_frame = tk.Frame(content_frame, bg='white')
            elif self.current_theme == 'dark':
                date_time_frame = tk.Frame(content_frame, bg='#414141')
            date_time_frame.pack(pady=pady, padx=padx, fill="x")

            calendar = Calendar(date_time_frame, selectmode="day", year=selected_event.date.year, month=selected_event.date.month, day=selected_event.date.day)
            calendar.pack(side="left", padx=(0, 10))

            if self.current_theme == 'light':
                time_frame = tk.Frame(date_time_frame, bg='white')
            elif self.current_theme == 'dark':
                time_frame = tk.Frame(date_time_frame, bg='#414141')
            time_frame.pack(side="left")

            hour_var = tk.StringVar(value=str(selected_event.date.hour).zfill(2))
            minute_var = tk.StringVar(value=str(selected_event.date.minute).zfill(2))

            ttk.Label(time_frame, text="Години:", font=("Segoe UI", 14)).pack(anchor="w")
            hour_spinbox = ttk.Spinbox(time_frame, from_=0, to=23, textvariable=hour_var, width=2, font=("Segoe UI", 14))
            hour_spinbox.pack(anchor="w", pady=(0, 5))

            ttk.Label(time_frame, text="Хвилини:", font=("Segoe UI", 14)).pack(anchor="w")
            minute_spinbox = ttk.Spinbox(time_frame, from_=0, to=59, textvariable=minute_var, width=2, font=("Segoe UI", 14))
            minute_spinbox.pack(anchor="w")

            ttk.Label(content_frame, text="Список електронних адрес (через кому):", font=("Segoe UI", 12)).pack(pady=pady, padx=padx, anchor="w")
            if self.current_theme == 'light':
                emails_entry = tk.Entry(content_frame, font=("Segoe UI Semibold", 12), bg='white', fg='black')
            elif self.current_theme == 'dark':
                emails_entry = tk.Entry(content_frame, font=("Segoe UI Semibold", 12), bg='#414141', fg='white')
            emails_entry.insert(0, ', '.join(selected_event.emails))
            emails_entry.pack(pady=pady, padx=padx, fill="x")
            
            notify_var = tk.BooleanVar(value=selected_event.notify_date is not None)  # Set the checkbox value depending on the presence of a notification

            def toggle_notify():
                if notify_var.get():
                    notify_frame.pack(side="top", fill="x", padx=padx, anchor="w", pady=(10, 10))
                else:
                    notify_frame.pack_forget()

                update_scroll_region(None)

            notify_checkbox = ttk.Checkbutton(content_frame, text="Нагадати", variable=notify_var, command=toggle_notify)
            notify_checkbox.pack(pady=pady, padx=padx, anchor="w")

            if self.current_theme == 'light':
                notify_frame = tk.Frame(content_frame, bg='white')
            elif self.current_theme == 'dark':
                notify_frame = tk.Frame(content_frame, bg='#414141')
            notify_frame.pack(pady=pady, padx=padx, fill="x", anchor="w")

            # Check if the user is signed in to a Google account
            if self.is_google_account_authenticated():
                # If you are logged in, then we allow the use of notifications
                notify_checkbox.config(state="normal")
            else:
                # If you are not logged in, then make the widget inactive
                notify_checkbox.config(state="disabled")

                # Create a label indicating that you need to sign in to your Google account
                notify_label = ttk.Label(content_frame, text="Щоб користуватися нагадуваннями, увійдіть у свій обліковий запис Google.", foreground="red")
                notify_label.pack(pady=(0, 10), padx=padx, anchor="w")

            if selected_event.notify_date:
                notify_calendar = Calendar(notify_frame, selectmode="day", year=selected_event.notify_date.year, month=selected_event.notify_date.month, day=selected_event.notify_date.day)
            else:
                notify_calendar = Calendar(notify_frame, selectmode="day", year=datetime.datetime.now().year, month=datetime.datetime.now().month, day=datetime.datetime.now().day)
            notify_calendar.pack(side="left", padx=(0, 10))

            if self.current_theme == 'light':
                notify_time_frame = tk.Frame(notify_frame, bg='white')
            elif self.current_theme == 'dark':
                notify_time_frame = tk.Frame(notify_frame, bg='#414141')
            notify_time_frame.pack(side="left")

            notify_hour_var = tk.StringVar(value=str(selected_event.notify_date.hour).zfill(2) if selected_event.notify_date else str(datetime.datetime.now().hour).zfill(2))
            notify_minute_var = tk.StringVar(value=str(selected_event.notify_date.minute).zfill(2) if selected_event.notify_date else str(datetime.datetime.now().minute).zfill(2))

            ttk.Label(notify_time_frame, text="Години:", font=("Segoe UI", 14)).pack(anchor="w")
            notify_hour_spinbox = ttk.Spinbox(notify_time_frame, from_=0, to=23, textvariable=notify_hour_var, width=2, font=("Segoe UI", 14))
            notify_hour_spinbox.pack(anchor="w", pady=(0, 5))

            ttk.Label(notify_time_frame, text="Хвилини:", font=("Segoe UI", 14)).pack(anchor="w")
            notify_minute_spinbox = ttk.Spinbox(notify_time_frame, from_=0, to=59, textvariable=notify_minute_var, width=2, font=("Segoe UI", 14))
            notify_minute_spinbox.pack(anchor="w")

            toggle_notify()  # Hide/show notification widgets depending on the initial state of the checkbox

            def save_changes():
                try:
                    edited_date_str = calendar.get_date()
                    edited_date = datetime.datetime.strptime(edited_date_str, "%m/%d/%y")
                    edited_date = edited_date.replace(hour=int(hour_var.get()), minute=int(minute_var.get()))
                    edited_emails = [email.strip() for email in emails_entry.get().split(',')]
                    edited_event = Event(title_entry.get(), description_entry.get("1.0", tk.END).strip(), edited_date, edited_emails)

                    edited_event.notify_date = None  # Set the notification date to None by default

                    if notify_var.get():
                        notify_selected_date = notify_calendar.get_date()
                        notify_hour = int(notify_hour_var.get())
                        notify_minute = int(notify_minute_var.get())
                        edited_event.notify_date = datetime.datetime.strptime(notify_selected_date, '%m/%d/%y').replace(hour=notify_hour, minute=notify_minute)

                    self.events[event_index] = edited_event

                    self.update_events_listbox()

                    edit_window.destroy()

                    self.show_event_details(event_index)
                except ValueError as e:
                    messagebox.showerror("Помилка", f"Невірні дані: {e}")                    
        # Create a button to save changes
            save_button = ttk.Button(content_frame, style="Yellow.TButton", text="Зберегти зміни", command=save_changes)
            save_button.pack(pady=10, padx=padx, side="right")
        else:
            messagebox.showinfo("Помилка", "Виберіть подію для редагування.")

    def delete_event(self, event_index=None):
        if event_index is not None and 0 <= event_index < len(self.events):
            del self.events[event_index]

            self.update_events_listbox()

            self.show_event_details(-1)
        else:
            messagebox.showinfo("Помилка", "Виберіть подію для видалення.")

    def show_context_menu(self, event):
        self.context_menu.post(event.x_root, event.y_root)
        
    def send_notifications(self):
        try:
            # Loading credentials from a file
            with open('credentials.json', 'r') as file:
                creds_data = json.load(file)
            # Converting a string to a dictionary
            creds_data = json.loads(creds_data)
            creds = Credentials(
                token=creds_data['token'],
                refresh_token=creds_data['refresh_token'],
                token_uri=creds_data['token_uri'],
                client_id=creds_data['client_id'],
                client_secret=creds_data['client_secret']
            )
            
            # Creating a Credential Object
            name_credentials = Credentials.from_authorized_user_info(creds_data)
            # Checking if the access token needs updating
            if name_credentials.expired and name_credentials.refresh_token:
                name_credentials.refresh(Request())
            
            service = build('gmail', 'v1', credentials=creds)

            # Checking all user events
            for event in self.events:
                # Checking whether the event has a notification and whether its time has come
                if event.notify_date and datetime.datetime.now() >= event.notify_date and not event.sent:
                    url = 'https://www.googleapis.com/oauth2/v3/userinfo'
                    headers = {'Authorization': f'Bearer {name_credentials.token}'}
                    response = requests.get(url, headers=headers)
                    if response.status_code == 200:
                        data = response.json()
                        name = data.get('name')
                    
                    # Создаем сообщение
                    message = MIMEText(f"{event.description}\nВід {name} за допомогою Event Planner.")
                    message['to'] = ", ".join(event.emails)
                    message['subject'] = f"Нагадуємо, що подія {event.title} розпочнеться {event.date}!"

                    create_message = {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode()}

                    try:
                        message = (service.users().messages().send(userId="me", body=create_message).execute())
                        print(F'sent message to {message} Message Id: {message["id"]}')
                        
                        # Mark the event as sent
                        event.sent = True
                    except HTTPError as error:
                        print(F'An error occurred: {error}')
                        message = None

        except Exception as e:
            print(f"Error sending notifications: {e}")
        
    def run_notification_loop(self):
        def notification_thread():
            while True:
                self.send_notifications()
                time.sleep(60)

        thread = threading.Thread(target=notification_thread)
        thread.daemon = True
        thread.start()

    def open_settings(self):
        settings_window = SettingsWindowUKR(self, self.current_theme)
        
    def change_theme(self):
        # Changing the colors of controls depending on the selected theme
        if self.current_theme == "dark":  # Dark theme
            # Colors for dark theme
            background_color = "#414141"
            details_text_color = 'white'
            listbox_background_color = "#323232"
            listbox_item_background_color = "#323232"
            listbox_item_foreground_color = "white"
            button_background_color = "#323232"
            button_foreground_color = "white"
            button_active_background_color = "#FCF7C9"
            button_active_foreground_color = "#323232"
        else: # Light theme
            # Light Theme Colors
            background_color = "white"
            details_text_color = '#262626'
            listbox_background_color = "white"
            listbox_item_background_color = "#FCF7C9"
            listbox_item_foreground_color = "#323232"
            button_background_color = "#FCF7C9"
            button_foreground_color = "#323232"
            button_active_background_color = "#323232"
            button_active_foreground_color = "white"

        # Applying colors to controls
        self.root.configure(bg=background_color)
        self.toolbar_frame.configure(bg=background_color)
        self.events_listbox.configure(bg=listbox_background_color)
        for widget in self.events_listbox.winfo_children():
                widget.config(bg=listbox_item_background_color, fg=listbox_item_foreground_color)
        self.context_menu.configure(bg=background_color, fg=button_foreground_color)
        self.details_text.configure(bg=background_color, fg=details_text_color)
        self.style.configure("Yellow.TButton", background=button_background_color, foreground=button_foreground_color)
        self.style.map("Yellow.TButton", background=[("active", button_active_background_color)], foreground=[("active", button_active_foreground_color)])

    def run(self):
        self.root.mainloop()
        
class SettingsWindowUKR:
    def __init__(self, parent, current_theme):
        self.parent = parent
        self.settings_window = tk.Toplevel(parent.root)
        self.settings_window.title("Налаштування")
        self.settings_window.iconphoto(False, tk.PhotoImage(file=resource_path('EP_cover.png')))
        
        self.current_theme = current_theme
        
        if self.current_theme == 'light':
            self.settings_window.configure(bg='white')
        elif self.current_theme == 'dark':
            self.settings_window.configure(bg='#414141')
        
        self.credentials = None
        
        self.create_widgets()

    def create_widgets(self):
        # Adding Controls
        if self.current_theme == 'light':
            self.settings_frame = tk.Frame(self.settings_window, bg='white')
        elif self.current_theme == 'dark':
            self.settings_frame = tk.Frame(self.settings_window, bg='#414141')
        self.settings_frame.pack(pady=5, padx=10)

        # Google Account Zone
        if self.current_theme == 'light':
            self.google_account_frame = tk.Frame(self.settings_frame, bg='white')
        elif self.current_theme == 'dark':
            self.google_account_frame = tk.Frame(self.settings_frame, bg='#414141')
        self.google_account_frame.pack(side="left", padx=10)

        tk.Label(self.google_account_frame, text="Налаштування облікового запису Google:", font=("Segoe UI Semibold", 12)).pack(pady=10, anchor="w")

        # Checking if Google Account data is saved
        if os.path.exists('credentials.json'):
            # Loading credentials from a file
            with open('credentials.json', 'r') as file:
                creds_data = json.load(file)
            # Converting a string to a dictionary
            creds_data = json.loads(creds_data)
            # Creating a Credential Object
            self.credentials = Credentials.from_authorized_user_info(creds_data)
            # Checking if credentials are valid
            if not self.credentials.valid:
                # Checking if the access token needs updating
                if self.credentials.expired and self.credentials.refresh_token:
                    self.credentials.refresh(Request())
                    # Display Google account information
                    email, name = self.get_user_info()
                    if email and name:
                        ttk.Label(self.google_account_frame, text="Електронна пошта:", font=("Segoe UI", 12)).pack(anchor="w")
                        ttk.Label(self.google_account_frame, text=email, font=("Segoe UI", 12)).pack(anchor="w")
                        ttk.Label(self.google_account_frame, text="Ім'я користувача:", font=("Segoe UI", 12)).pack(anchor="w")
                        ttk.Label(self.google_account_frame, text=name, font=("Segoe UI", 12)).pack(anchor="w")
                        # Logout button
                        ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Вийти зі свого облікового запису", command=self.logout_google_account).pack(pady=10, anchor="w")
                    else:
                        # If we were unable to obtain information, we display the login button
                        ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Увійти в обліковий запис Google", command=self.login_google_account).pack(pady=10, anchor="w")
                else:
                    # Credentials are invalid, you need to log in again
                    self.login_google_account()
            else:
                # Display Google account information
                email, name = self.get_user_info()
                if email and name:
                    ttk.Label(self.google_account_frame, text="Електронна пошта:", font=("Segoe UI", 12)).pack(anchor="w")
                    ttk.Label(self.google_account_frame, text=email, font=("Segoe UI", 12)).pack(anchor="w")
                    ttk.Label(self.google_account_frame, text="Ім'я користувача:", font=("Segoe UI", 12)).pack(anchor="w")
                    ttk.Label(self.google_account_frame, text=name, font=("Segoe UI", 12)).pack(anchor="w")
                    # Logout button
                    ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Вийти зі свого облікового запису", command=self.logout_google_account).pack(pady=10, anchor="w")
                else:
                    # If we were unable to obtain information, we display the login button
                    ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Увійти в обліковий запис Google", command=self.login_google_account).pack(pady=10, anchor="w")
        else:
            # If there are no credentials, we display the login button
            ttk.Button(self.google_account_frame, style="Yellow.TButton", text="Увійти в обліковий запис Google", command=self.login_google_account).pack(pady=10, anchor="w")

        # Zone of buttons for changing theme and language
        if self.current_theme == 'light':
            self.theme_language_frame = tk.Frame(self.settings_frame, bg='white')
        elif self.current_theme == 'dark':            
            self.theme_language_frame = tk.Frame(self.settings_frame, bg='#414141')
        self.theme_language_frame.pack(side="right", padx=10)
        
        # Adding help button centered in the zone
        ttk.Button(self.theme_language_frame, style="Yellow.TButton", text="Довідка", command=lambda: webbrowser.open("https://docs.google.com/document/d/1yQYKMG--Q4hUG8daiSD0xcQOntGQ_f_nzOiG34x_KQE/edit?usp=sharing")).pack(side="top", padx=5)

        # Adding buttons to change theme and language
        ttk.Button(self.theme_language_frame, style="Yellow.TButton", text="Змінити кольорову тему", command=self.save_and_change_theme).pack(side="left", padx=5)
        ttk.Button(self.theme_language_frame, style="Yellow.TButton", text="Змінити мову", command=self.change_language).pack(side="bottom", pady=1)
        
        tk.Label(self.settings_window, text="© 2024 Hlib Ishchenko. All rights reserved.", font=("Segoe UI", 12)).pack(pady=10, side="bottom")

    def get_user_info(self):
        url = 'https://www.googleapis.com/oauth2/v3/userinfo'
        headers = {'Authorization': f'Bearer {self.credentials.token}'}
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            email = data.get('email')
            name = data.get('name')
            return email, name
        else:
            print('Error:', response.status_code)
            return None, None

    def login_google_account(self):
        os.environ['OAUTHLIB_RELAX_TOKEN_SCOPE'] = '1'
        # Creating OAuth2 credentials to access the Gmail API and retrieve user information
        scopes = [
            'https://www.googleapis.com/auth/gmail.send',
            'https://www.googleapis.com/auth/userinfo.email',
            'https://www.googleapis.com/auth/userinfo.profile'
        ]
        flow = InstalledAppFlow.from_client_secrets_file(resource_path('your_client_secret.json'), scopes=scopes)
        self.credentials = flow.run_local_server(port=0)

        # Saving credentials to a file
        with open('credentials.json', 'w') as file:
            json.dump(self.credentials.to_json(), file)
        
        # Rebooting the settings window
        self.settings_window.destroy()
        self.__init__(self.parent)

    def logout_google_account(self):
        # Deleting files with credentials
        if os.path.exists('credentials.json'):
            os.remove('credentials.json')

        # Rebooting the settings window
        self.settings_window.destroy()
        self.__init__(self.parent)
        
    def save_and_change_theme(self):
        response = messagebox.askquestion("Увага", "Програму буде перезапущено, щоб змінити тему. Продовжити?")
        if response == 'yes':
            new_theme = "dark" if self.current_theme == "light" else "light"
            
            # Save the current theme to a file
            with open("current_theme.txt", "w") as f:
                f.write(new_theme)
                
            self.parent.save_events_to_file()
            # Restart application
            python = sys.executable
            os.execl(python, python, *sys.argv)

    def show(self):
        self.settings_window.deiconify()

    def hide(self):
        self.settings_window.withdraw()
    
    def change_language(self):
        response = messagebox.askquestion("Увага", "Мова зміниться на англійську. Програма буде перезапущена, щоб змінити мову. Продовжити?")
        if response == 'yes':
            if os.path.exists("current_language.txt"):
                with open("current_language.txt", "r") as f:
                    current_language = f.read()
            else:
                current_language = "english"
                
            new_language = "ukrainian" if current_language == "english" else "english"
            
            # Save the current theme to a file
            with open("current_language.txt", "w") as f:
                f.write(new_language)
                
            self.parent.save_events_to_file()
            # Restart application
            python = sys.executable
            os.execl(python, python, *sys.argv)  

if __name__ == "__main__":
    root = tk.Tk()
    
    root.iconphoto(False, tk.PhotoImage(file=resource_path('EP_cover.png')))
    
    if os.path.exists("current_scaling.txt"):
        with open("current_scaling.txt", "r") as f:
            current_scaling = float(f.read())
    else:
        current_scaling = 1.33
    root.tk.call('tk', 'scaling', current_scaling)
    
    if os.path.exists("current_language.txt"):
        with open("current_language.txt", "r") as f:
            current_language = f.read()
    else:
        current_language = "english"
    if current_language == "english":
        app = EventPlannerApp(root)
    elif current_language == "ukrainian":
        app = EventPlannerAppUKR(root)
    
    app.root.geometry(f"{int(800/(1.33/current_scaling))}x{int(600/(1.33/current_scaling))}")
    app.run()