import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import csv
import os
import threading
import subprocess
import psutil
import pygetwindow as gw
import time
import win32gui
import win32process
from datetime import datetime
from PIL import Image, ImageTk  # To handle app icons
from pynput import mouse, keyboard

# CSV file to store tasks
tasks_file_path = "tasks.csv"
activity_file_path = "active_window_log.csv"
apps_file_path = "frequent_apps.csv"

# Variables for tracking activity
window_start_time = None
mouse_clicks = 0
keypresses = 0
selected_task = None
selected_apps = []  # Store selected apps

# Load tasks from CSV file
def load_tasks():
    tasks = []
    if os.path.exists(tasks_file_path):
        with open(tasks_file_path, newline='', encoding='utf-8') as file:
            reader = csv.reader(file)
            tasks = [row[0] for row in reader]
    return tasks

# Save task to CSV file
def save_task(task):
    with open(tasks_file_path, mode="a", newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([task])

# Remove task from CSV file
def remove_task(task_to_remove):
    tasks = load_tasks()
    tasks = [task for task in tasks if task != task_to_remove]
    with open(tasks_file_path, mode="w", newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for task in tasks:
            writer.writerow([task])

# Load frequent apps from CSV file
def load_apps():
    apps = []
    if os.path.exists(apps_file_path):
        with open(apps_file_path, newline='', encoding='utf-8') as file:
            reader = csv.reader(file)
            apps = [row for row in reader]
    return apps

# Save app to CSV file
def save_app(app_name, app_path):
    with open(apps_file_path, mode="a", newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([app_name, app_path])

# Function to track mouse clicks
def on_click(x, y, button, pressed):
    global mouse_clicks
    if pressed:
        mouse_clicks += 1

# Function to track key presses
def on_press(key):
    global keypresses
    keypresses += 1

# Function to get active window title
def get_active_window_title():
    window = gw.getActiveWindow()
    if window is not None:
        return window.title
    return None

# Function to get active process name
def get_active_process():
    active_window = win32gui.GetForegroundWindow()
    thread_id, process_id = win32process.GetWindowThreadProcessId(active_window)
    process = psutil.Process(process_id)
    return process.name()

# Function to split the window title and extract activeWindow, activeObject, activeWork
def parse_window_title(window_title):
    parts = window_title.split(" - ")
    
    if len(parts) >= 2:
        active_work = parts[0]  # First part is activeWork
        active_window = parts[-1]  # Last part is activeWindow
        active_object = " - ".join(parts[1:-1])  # Everything in between is activeObject
    else:
        active_work = window_title
        active_object = window_title
        active_window = window_title
    
    return active_window, active_object, active_work

# Function to calculate the time spent on the window (returning whole seconds)
def get_time_spent(start_time):
    return int((datetime.now() - start_time).total_seconds())

# Function to format datetime in 'yyyy-MM-dd HH:mm:ss'
def format_time(dt):
    return dt.strftime('%Y-%m-%d %H:%M:%S')

# Function to track the active window and save the output to CSV
def track_active_window():
    global window_start_time, mouse_clicks, keypresses, selected_task

    current_window = None
    window_start_time = datetime.now()

    # Initialize mouse and keyboard listeners
    mouse_listener = mouse.Listener(on_click=on_click)
    keyboard_listener = keyboard.Listener(on_press=on_press)
    mouse_listener.start()
    keyboard_listener.start()

    # Check if the file exists
    file_exists = os.path.exists(activity_file_path)
    
    # Open the file in append mode if it exists, or write mode if it doesn't (with utf-8 encoding)
    with open(activity_file_path, mode="a" if file_exists else "w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        # Write header if the file is new
        if not file_exists:
            writer.writerow(["activeWindow", "activeObject", "activeWork", "startTime", "endTime", "timeSpent", "mouseClicks", "keyPresses", "Task"])

        while True:
            window_title = get_active_window_title()

            if window_title and (current_window != window_title):
                # Log the previous window's data
                if current_window:
                    end_time = datetime.now()
                    time_spent = get_time_spent(window_start_time)
                    writer.writerow([active_window, active_object, active_work, format_time(window_start_time), format_time(end_time), time_spent, mouse_clicks, keypresses, selected_task])
                    print(f"Window: {active_window} > {active_object} > {active_work}, Task: {selected_task}, Time spent: {time_spent} seconds, Mouse clicks: {mouse_clicks}, Key presses: {keypresses}")

                # Reset event counters
                mouse_clicks = 0
                keypresses = 0
                
                # Parse the new window's details
                active_window, active_object, active_work = parse_window_title(window_title)

                # Update the current window
                current_window = window_title
                window_start_time = datetime.now()

            file.flush()  # Ensure the data is written in real-time
            time.sleep(1)  # Check the active window every second

# Function to show the dashboard by running "streamlit run app.py"
def show_dashboard():
    try:
        # Run the streamlit command to open the dashboard
        subprocess.Popen(["streamlit", "run", "webApp.py"], shell=True)
        print("Dashboard is running...")
    except Exception as e:
        print(f"Error opening the dashboard: {e}")

# Function to open an application
def open_application(app_path):
    try:
        subprocess.Popen(app_path, shell=True)
    except Exception as e:
        print(f"Error opening the application: {e}")

# Tkinter GUI to manage tasks and apps
def task_manager_gui():
    global selected_task

    # Create the main window
    root = tk.Tk()
    root.title("RAFTASK-V1")
    root.geometry("500x400")
    root.configure(bg="#fffde7")

    root.iconbitmap("app_icon.ico")

    # Allow the window to resize
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    # Function to add a new task
    def add_task():
        task = task_entry.get()
        if task:
            save_task(task)
            task_entry.delete(0, tk.END)
            task_combobox['values'] = load_tasks()  # Update the dropdown with the new task
        else:
            messagebox.showwarning("Input Error", "Please enter a task")

    def delete_task():
        global selected_task
        task = task_combobox.get()
        if task:
            remove_task(task)
            task_combobox['values'] = load_tasks()  # Update the dropdown after removal
            task_combobox.set('')
            selected_task = "Removed Task"
            messagebox.showinfo("RAFTASK", f"Task '{task}' has been removed successfully.")
        else:
            messagebox.showwarning("Selection Error", "Please select a task to remove")

    # Function to select frequently used apps
    def select_app():
        app_path = filedialog.askopenfilename(title="Select Application", filetypes=[("Executable Files", "*.exe"), ("Shortcut Files", "*.lnk")])
        if app_path:
            app_name = os.path.splitext(os.path.basename(app_path))[0]  # Get app name without extension
            save_app(app_name, app_path)
            update_sidebar()

    # Load existing tasks and create a dropdown
    tasks = load_tasks()

    # Load existing apps
    selected_apps = load_apps()

    # Styles for Entry and Button
    style = ttk.Style()
    style.configure("TEntry", padding=5, relief="solid", font=("Arial", 12))
    style.configure("TButton", padding=5, relief="solid", font=("Arial", 12))
    style.map("TButton",
              foreground=[('active', '#ffffff'), ('!disabled', '#000000')],
              background=[('active', '#0056b3'), ('!disabled', '#007bff')])
    style.configure("TCombobox", font=("Arial", 12), relief="solid", padding=5)

    # Frame to hold the task entry and buttons
    entry_frame = tk.Frame(root, bg="#b2ebf2")
    entry_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

    # Resizable Entry box for new task
    task_entry = tk.Entry(entry_frame, font=("Arial", 12))
    task_entry.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

    # Add Task Button next to the entry field
    add_task_button = tk.Button(entry_frame, text="Add Task", command=add_task, font=("Arial", 12), bg="#007bff", fg="white", relief="raised", activebackground="#0056b3", activeforeground="#ffffff")
    add_task_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

    # Show Dashboard Button next to Add Task Button
    dashboard_button = tk.Button(entry_frame, text="Dashboard", command=show_dashboard, font=("Arial", 12), bg="#28a745", fg="white", relief="raised", activebackground="#218838", activeforeground="#ffffff")
    dashboard_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

    # Frame to hold the dropdown selection and label
    dropdown_frame = tk.Frame(root, bg="#c8e6c9")
    dropdown_frame.grid(row=2, column=1, padx=5, pady=5, sticky="nsew")

    # Label for Dropdown
    dropdown_label = tk.Label(dropdown_frame, text="Select Current Task", font=("Arial", 12, "bold"), bg="#c8e6c9")
    dropdown_label.pack(anchor='w', padx=5, pady=5)

    # Dropdown (Combobox) for selecting a task
    task_combobox = ttk.Combobox(dropdown_frame, values=tasks, font=("Arial", 10), style="TCombobox")
    task_combobox.pack(fill=tk.X, padx=5, pady=5, expand=True)

    # Delete Task Button below Dropdown
    delete_task_button = tk.Button(dropdown_frame, text="Delete Task", command=delete_task, font=("Arial", 12), bg="#dc3545", fg="white", relief="raised", activebackground="#c82333", activeforeground="#ffffff")
    delete_task_button.pack(anchor='w', padx=5, pady=5)

    # Frame to hold app selection button and sidebar
    app_frame = tk.Frame(root, bg="#e1f5fe")
    app_frame.grid(row=0, column=0, rowspan=3, padx=5, pady=5, sticky="ns")

    # Select Apps Button
    select_app_button = tk.Button(app_frame, text="Select Apps", command=select_app, font=("Arial", 12), bg="#007bff", fg="white", relief="raised", activebackground="#0056b3", activeforeground="#ffffff")
    select_app_button.pack(padx=5, pady=5, anchor='n')

    # Sidebar to hold app icons
    sidebar = tk.Frame(app_frame, bg="#e1f5fe")
    sidebar.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    # Function to update the sidebar with app icons
    def update_sidebar():
        # Clear existing widgets
        for widget in sidebar.winfo_children():
            widget.destroy()

        # Load apps and create buttons
        apps = load_apps()
        for app_name, app_path in apps:
            try:
                icon_path = app_path  # Use the same path to load icon for shortcut files
                icon = Image.open(icon_path)
                icon = icon.resize((24, 24), Image.ANTIALIAS)
                icon_img = ImageTk.PhotoImage(icon)
                app_button = tk.Button(sidebar, text=app_name, font=("Arial", 10), image=icon_img, compound=tk.LEFT, command=lambda p=app_path: open_application(p))
                app_button.image = icon_img  # Keep a reference to prevent garbage collection
                app_button.pack(fill=tk.X, pady=2)
            except Exception as e:
                app_button = tk.Button(sidebar, text=app_name, font=("Arial", 10), command=lambda p=app_path: open_application(p))
                app_button.pack(fill=tk.X, pady=2)

    # Update the sidebar initially
    update_sidebar()

    # Frame to display active window and selected task
    status_frame = tk.Frame(root, bg="#b2ebf2")
    status_frame.grid(row=3, column=1, padx=5, pady=5, sticky="nsew")

    # Label to display the active window and selected task
    status_title = tk.Label(status_frame, text="Active Window & Task", font=("Arial", 12, "bold"), bg="#b2ebf2")
    status_title.pack(anchor='w', padx=5, pady=5)

    status_label = tk.Label(status_frame, text="", font=("Arial", 10), bg="#b2ebf2", padx=5, pady=5, justify="left", anchor='w')
    status_label.pack(fill=tk.X, padx=5, pady=5, expand=True)

    # Function to update the status label
    def update_status_label():
        window_title = get_active_window_title()
        active_window, active_object, active_work = parse_window_title(window_title)
        status_text = f"Window: {active_window} > {active_work}\nTask: {selected_task}"
        status_label.config(text=status_text)
        root.after(1000, update_status_label)

    # Function to update the selected task from dropdown
    def update_selected_task(event):
        global selected_task
        selected_task = task_combobox.get()

    task_combobox.bind("<<ComboboxSelected>>", update_selected_task)

    # Start updating the status label
    update_status_label()

    # Start the main loop of the Tkinter GUI
    root.mainloop()

# Run the task manager GUI in a separate thread and start tracking windows
if __name__ == "__main__":
    # Start tracking the active window in a separate thread
    tracking_thread = threading.Thread(target=track_active_window)
    tracking_thread.daemon = True
    tracking_thread.start()

    # Run the Tkinter GUI for task management
    task_manager_gui()