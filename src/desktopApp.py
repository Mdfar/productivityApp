import tkinter as tk
from tkinter import ttk
import csv
import threading
import subprocess
import psutil
import pygetwindow as gw
import time
import win32gui
import win32process
from datetime import datetime, timedelta
from PIL import Image, ImageTk  # To handle app icons
from pynput import mouse, keyboard
import pandas as pd

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.tableview import Tableview
from tkinter import messagebox, filedialog
import os

ROOT_DIR  = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def find_file(filename, search_path):
    for root, dirs, files in os.walk(search_path):
        if filename in files:
            return os.path.join(root, dirs, filename)
    return None

# Ensure that the file paths are found, otherwise raise an error
def get_file_path(filename, dir):
    file_path = find_file(filename, ROOT_DIR)
    if file_path is None:
        raise FileNotFoundError(f"{filename} not found.")
    return file_path

# Update paths
tasks_file_path = os.path.join(ROOT_DIR, 'data', "tasks.csv")
goals_file_path = os.path.join(ROOT_DIR, 'data', "goals.csv")
activity_file_path = os.path.join(ROOT_DIR, 'data', "active_window_log.csv")
apps_file_path = os.path.join(ROOT_DIR, 'data', "frequent_apps.csv")

# Variables for tracking activity
window_start_time = None
mouse_clicks = 0
keypresses = 0
selected_task = None
selected_goal = None
table = None
selected_apps = []  # Store selected apps

# Load tasks from CSV file
def load_tasks():
    tasks = []
    if os.path.exists(tasks_file_path):
        with open(tasks_file_path, newline='', encoding='utf-8') as file:
            reader = csv.reader(file)
            tasks = [row[0] for row in reader]
    return tasks

# Load goals from CSV file
def load_goals():
    goals = []
    if os.path.exists(goals_file_path):
        with open(goals_file_path, newline='', encoding='utf-8') as file:
            reader = csv.reader(file)
            goals = [row[0] for row in reader]
    return goals

# Save task to CSV file
def save_task(task):
    with open(tasks_file_path, mode="a", newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([task])

# Save goal to CSV file
def save_goal(goal):
    with open(goals_file_path, mode="a", newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow([goal])

# Remove task from CSV file
def remove_task(task_to_remove):
    tasks = load_tasks()
    tasks = [task for task in tasks if task != task_to_remove]
    with open(tasks_file_path, mode="w", newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for task in tasks:
            writer.writerow([task])

# Remove goal from CSV file
def remove_goal(goal_to_remove):
    goals = load_goals()
    goals = [goal for goal in goals if goal != goal_to_remove]
    with open(goals_file_path, mode="w", newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for goal in goals:
            writer.writerow([goal])

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

# Function to parse the window title and extract activeWindow, activeObject, activeWork
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

# Function to calculate totals for all task/goal combinations from the CSV
def load_totals_from_csv():
    totals = {}

    # Read from active_window_log.csv
    if os.path.exists(activity_file_path):
        with open(activity_file_path, newline='', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                # Get the task and goal from each row
                task = row.get("task")
                goal = row.get("goal")
                if not task or not goal:
                    continue

                key = (goal, task)

                # Initialize the total for this goal/task combination if not already in the dictionary
                if key not in totals:
                    totals[key] = {"time_spent": 0, "mouse_clicks": 0, "key_presses": 0}

                # Add the time spent, mouse clicks, and key presses for the current row
                totals[key]["time_spent"] += int(row["timeSpent"])
                totals[key]["mouse_clicks"] += int(row["mouseClicks"])
                totals[key]["key_presses"] += int(row["keyPresses"])

    return totals

# Function to calculate totals based on the selected goal and task
def calculate_totals(goal, task):
    total_time_spent = 0
    total_mouse_clicks = 0
    total_key_presses = 0

    # Read from active_window_log.csv
    if os.path.exists(activity_file_path):
        with open(activity_file_path, newline='', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            for row in reader:
                # Filter rows by goal and task
                if row["task"] == task and row["goal"] == goal:
                    total_time_spent += int(row["timeSpent"])
                    total_mouse_clicks += int(row["mouseClicks"])
                    total_key_presses += int(row["keyPresses"])

    # Convert total time from seconds to HH:MM:SS
    total_time_spent = str(timedelta(seconds=total_time_spent))
    
    return total_time_spent, total_mouse_clicks, total_key_presses

# Function to update the table with all task/goal combinations found in the CSV
def update_table(goal=None, task=None):
    global table
    # Clear existing rows in the table
    for row in table.get_children():
        table.delete(row)

    # Get all the totals from the CSV file
    totals = load_totals_from_csv()

    # If a goal and task are selected, place them at the top
    if goal and task:
        total_time, total_clicks, total_keys = calculate_totals(goal, task)
        table.insert("", "end", values=(goal, task, total_time, total_clicks, total_keys))

    # Insert all task/goal combinations grouped by Goal and Task
    for (g, t), values in sorted(totals.items()):
        total_time_spent = str(timedelta(seconds=values["time_spent"]))
        table.insert("", "end", values=(g, t, total_time_spent, values["mouse_clicks"], values["key_presses"]))

# Function to track the active window and save the output to CSV
def track_active_window():
    global window_start_time, mouse_clicks, keypresses, selected_task, selected_goal

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
            writer.writerow(["activeWindow", "activeObject", "activeWork", "startTime", "endTime", "timeSpent", "mouseClicks", "keyPresses", "Task", "Goal"])

        while True:
            window_title = get_active_window_title()

            if window_title and (current_window != window_title):
                # Log the previous window's data
                if current_window:
                    end_time = datetime.now()
                    time_spent = get_time_spent(window_start_time)
                    writer.writerow([active_window, active_object, active_work, format_time(window_start_time), format_time(end_time), time_spent, mouse_clicks, keypresses, selected_task, selected_goal])
                    print(f"Window: {active_window} > {active_object} > {active_work}, Task: {selected_task}, Goal: {selected_goal}, Time spent: {time_spent} seconds, Mouse clicks: {mouse_clicks}, Key presses: {keypresses}")

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
    global selected_task, selected_goal, table

    # Create the main window using ttkbootstrap
    root = ttk.Window(title="RAFtask", size=(800,650))

    root.minsize(600, 400)

    style = ttk.Style("yeti")  # Set initial theme

    colors = root.style.colors
    
    # Allow the window to resize
    # Column Configuration: Set column 1 to be flexible
    root.columnconfigure(0, weight=0)  # Fixed width for column 0
    root.columnconfigure(1, weight=1)  # Flexible width for column 1 (middle)
    root.columnconfigure(2, weight=0)  # Fixed width for column 2

    root.rowconfigure(0, weight=0)  # Fixed height for entry_frame (row 0)
    root.rowconfigure(1, weight=0)  # Fixed height for dropdown_frame (row 1)
    root.rowconfigure(2, weight=0)  # Flexible height for table_frame (row 2)
    root.rowconfigure(3, weight=1)  # Flexible height for table_frame (row 3)

    # ---- Column 0: Left Sidebar (app_frame) ----
    app_frame = ttk.Frame(root, width=150, padding=5, relief="flat", style="primary")  # Fixed width
    app_frame.grid(row=0, column=0, rowspan=4, sticky="nsew", padx=5, pady=5)
    app_frame.grid_propagate(False)  # Prevent resizing of the frame

    # ---- Column 1: Middle Section ----

    # Frame 1: Entry Frame (fixed height)
    entry_frame = ttk.Frame(root, height=80, padding=5)
    entry_frame.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
    entry_frame.grid_propagate(False)  # Prevent vertical resizing

    # Frame 2: Dropdown Frame (fixed height)
    dropdown_frame = ttk.Frame(root, height=80, padding=5)
    dropdown_frame.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
    dropdown_frame.grid_propagate(False)  # Prevent vertical resizing

    # Frame 3: Table Frame (flexible height, stretches with window)
    table_frame = ttk.Frame(root, padding=5, height=80)
    table_frame.grid(row=2, column=1, padx=5, pady=5, sticky="ew")

    # Frame 4: Detail Frame (flexible height, stretches with window)
    detail_frame = ttk.Frame(root,padding=5, relief="groove")
    detail_frame.grid(row=3, column=1, padx=5, pady=5, sticky="nsew")

    # ---- Column 2: Right Sidebar (goal_frame) ----
    goal_frame = ttk.Frame(root, width=50, padding=5, bootstyle="primary")  # Fixed width with resize cursor
    goal_frame.grid(row=0, column=2, rowspan=4, sticky="nsew", padx=5, pady=5)
    goal_frame.grid_propagate(False)  # Prevent resizing of the frame
    # Load goals initially
    goals = load_goals()

    # Function to add a new task
    def add_task():
        task = task_entry.get()
        if task:
            save_task(task)
            task_entry.delete(0, tk.END)
            task_combobox['values'] = load_tasks()  # Update the dropdown with the new task
        else:
            messagebox.showwarning("Input Error", "Please enter a task")

    # Function to add a new goal
    def add_goal():
        goal = task_entry.get()
        if goal:
            save_goal(goal)
            task_entry.delete(0, tk.END)
            goals.append(goal)  # Add the goal to the list
            update_goal_list()  # Update the UI to display the goal
        else:
            messagebox.showwarning("Input Error", "Please enter a goal")

    # Function to delete a goal
    def delete_goal(goal):
        remove_goal(goal)  # Remove the goal from the CSV file
        goals.remove(goal)  # Remove the goal from the list
        update_goal_list()  # Update the UI to reflect changes

    # Function to update the list of goals dynamically
    def update_goal_list():
        global selected_goal
        selected_goal_var = tk.StringVar()  # Create a StringVar to track the selected goal

        # Clear existing goal widgets
        for widget in goal_frame.winfo_children():
            widget.destroy()

        goal_label = ttk.Label(goal_frame, text="All Goals", font=("Arial", 12, "bold"), bootstyle="primary-inverse")  # Match with 'primary' style
        goal_label.pack(anchor='center', padx=5, pady=5)

        # Create checkboxes for each goal
        for goal in goals:
            goal_item_frame = ttk.Frame(goal_frame)
            goal_item_frame.pack(fill="x", pady=2)

            # Define a function to handle goal selection (updates selected_goal when clicked)
            def on_goal_selected(goal):
                global selected_goal
                selected_goal = goal
                update_table(selected_goal)  # Update table when a goal is selected

            # Create a radio button for each goal
            goal_checkbox = ttk.Radiobutton(
                goal_item_frame,
                text=goal,
                variable=selected_goal_var,
                value=goal,
                command=lambda g=goal: on_goal_selected(g)  # Call the goal selection handler
            )
            goal_checkbox.pack(side="left", padx=5)

            # Delete button for the goal
            delete_button = ttk.Button(goal_item_frame, text="X", bootstyle="danger", command=lambda g=goal: delete_goal(g))
            delete_button.pack(side="right", padx=0)

    # Function to delete a task
    def delete_task():
        global selected_task
        task = task_combobox.get()
        if task:
            remove_task(task)
            task_combobox['values'] = load_tasks()  # Update the dropdown after removal
            task_combobox.set('')
            selected_task = None
            messagebox.showinfo("RAFTASK", f"Task '{task}' has been removed successfully.")
        else:
            messagebox.showwarning("Selection Error", "Please select a task to remove")

    # Load existing tasks and create a dropdown
    tasks = load_tasks()

    # Load existing apps
    selected_apps = load_apps()

    # Define a function to change the theme
    def change_theme(event):
        selected_theme = theme_combobox.get()
        style.theme_use(selected_theme)
    # Resizable Entry box for new task or goal
    task_entry = ttk.Entry(entry_frame, font=("Arial", 12))
    task_entry.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

    # Add Task Button next to the entry field
    add_task_button = ttk.Button(entry_frame, text="Add Task", command=add_task, bootstyle="primary")
    add_task_button.grid(row=1, column=0, padx=5, pady=5, sticky="ew")

    # Add Goal Button next to Add Task Button
    add_goal_button = ttk.Button(entry_frame, text="Add Goal", command=add_goal, bootstyle="success")
    add_goal_button.grid(row=1, column=1, padx=5, pady=5, sticky="ew")

    # Theme Selector (Combobox)
    label = ttk.Label(entry_frame, text="Theme:", font=("Arial", 10))
    label.grid(row=0, column=2, padx=5, pady=5, sticky="e")  # Adjust label to fit in the same row as task_entry

    # Create a combobox for theme selection
    themes = style.theme_names()  # Get available theme names
    theme_combobox = ttk.Combobox(entry_frame, values=themes)
    theme_combobox.set("flatly")  # Set default theme
    theme_combobox.grid(row=0, column=3, padx=0, pady=5, sticky="ew")  # Place the combobox next to the label

    # Bind the theme combobox to the theme change function
    theme_combobox.bind("<<ComboboxSelected>>", change_theme)

    # First, we will filter rows where 'goal' is not blank, and then group by 'goal' and 'task' to create the summary table.
    active_window_data = pd.read_csv("active_window_log.csv")
    filtered_data = active_window_data.dropna(subset=['goal'])

    global summary_table
    # Group by 'goal' and 'task', then aggregate the data to calculate the sum for 'timeSpent', 'mouseClicks', and 'keyPresses'
    summary_table = filtered_data.groupby(['goal', 'task']).agg({
        'timeSpent': 'sum',
        'mouseClicks': 'sum',
        'keyPresses': 'sum'
    }).reset_index()

    # Renaming columns for clarity
    summary_table.columns = ['Goal', 'Task', 'Net TimeSpent', 'Net MouseClicks', 'Net KeyPresses']

    # Define column data for Tableview
    coldata = [
        {"text": "Goal", "stretch": False},
        {"text": "Task", "stretch": False},  # This column will stretch
        {"text": "NTS (Net Time Spent)", "stretch": True},
        {"text": "NMC (Net Mouse Clicks)", "stretch": True},
        {"text": "NKP (Net Key Presses)", "stretch": True},
    ]
    
    # Create the Treeview widget (table)

    rowdata = summary_table.apply(tuple, axis=1).tolist()

    table = Tableview(
        master=table_frame,
        coldata=coldata,
        rowdata=rowdata,
        pagesize = 10,
        paginated=True,         # Adds pagination
        searchable=True,        # Enables searching in the table
        bootstyle=PRIMARY,      # ttkbootstrap style
        stripecolor=(colors.light, None),  # Set stripe color for alternating rows
    )

    table.autofit_columns()

    # Pack the tableview into the window
    table.pack(fill=BOTH, expand=YES, padx=10, pady=10)

    def update_table(goal=None):
        global table, summary_table

        # Prepare new rowdata based on the selected goal
        new_rowdata = []

        # Filter the summary_table to include only the selected goal
        if goal:
            selected_rows = summary_table[summary_table['Goal'] == goal]
            if not selected_rows.empty:
                new_rowdata = selected_rows.values.tolist()

        # Clear existing rows and insert new ones
        table.delete_rows()
        table.insert_rows('end', new_rowdata)

        # Reload the table data to reflect changes
        table.load_table_data()
        table.autofit_columns()

    # Initially update the goal list
    update_goal_list()
    # Load existing tasks and goals
    tasks = load_tasks()
    goals = load_goals()

    # Label for Dropdown
    dropdown_label = ttk.Label(dropdown_frame, text="Select Current Task", font=("Arial", 12, "bold"))
    dropdown_label.pack(anchor='w', padx=5, pady=5)

    # Dropdown (Combobox) for selecting a task
    task_combobox = ttk.Combobox(dropdown_frame, values=tasks, font=("Arial", 10))
    task_combobox.pack(fill="x", padx=5, pady=5, expand=False)

    # Delete Task Button below Dropdown
    delete_task_button = ttk.Button(dropdown_frame, text="Delete Task", command=delete_task, bootstyle="danger")
    delete_task_button.pack(anchor='w', padx=5, pady=5)

    # Function to select frequently used apps
    def select_app():
        app_path = filedialog.askopenfilename(title="Select Application", filetypes=[("Executable Files", "*.exe"), ("Shortcut Files", "*.lnk")])
        if app_path:
            app_name = os.path.splitext(os.path.basename(app_path))[0]  # Get app name without extension
            save_app(app_name, app_path)
            update_sidebar()

    # Sidebar to hold app icons
    sidebar = ttk.Frame(app_frame, bootstyle="primary")
    sidebar.pack(fill="both", expand=True, padx=5, pady=5)

    # Select Apps Button inside the sidebar at the top
    select_app_button = ttk.Button(sidebar, text="Add Apps", command=select_app, bootstyle="dark", cursor="hand2")
    select_app_button.pack(side="top", fill="x", padx=5, pady=5)

    def remove_app_button(event, button, app_name, app_path, popup):
        button.destroy()  # Remove the button

        # Load all apps from the CSV file
        apps = load_apps()

        # Remove the selected app from the list
        apps = [app for app in apps if app != [app_name, app_path]]

        # Write the updated list back to the CSV file
        with open(apps_file_path, mode="w", newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerows(apps)
        
        popup.destroy()

    def close_popup(event, popup):
        if not popup.winfo_containing(event.x_root, event.y_root):  # Check if the click is outside the popup
            popup.destroy()
            root.unbind("<Button-1>")  # Unbind the event once the popup is closed

    def show_popup(event, button, app_name, app_path):
        popup = tk.Toplevel(root)
        popup.overrideredirect(True)  # Remove title bar and frame
        popup.geometry(f"150x40+{event.x_root}+{event.y_root}")  # Position the window near the click

        # Bind the root window to detect clicks outside the popup
        root.bind("<Button-1>", lambda e: close_popup(e, popup))

        delete_button = ttk.Button(popup, text="Remove this app", command=lambda: remove_app_button(event, button, app_name, app_path, popup), bootstyle="light")
        delete_button.pack(fill="both", expand=True, pady=5, padx=5)

        # Focus on the popup and dismiss it when losing focus
        popup.focus_force()
        popup.bind("<FocusOut>", lambda e: popup.withdraw())  # Closes when focus is lost

    def update_sidebar():
        for widget in sidebar.winfo_children():
            if widget != select_app_button:  # Do not destroy the "Add Apps" button
                widget.destroy()

        apps = load_apps()  # Loads apps dynamically
        for app_name, app_path in apps:
            app_button = ttk.Button(sidebar, text=app_name, bootstyle="primary")
            app_button.pack(fill="x", pady=2)

            # Bind right-click to show popup window, passing app_name and app_path
            app_button.bind("<Button-3>", lambda e, b=app_button, n=app_name, p=app_path: show_popup(e, b, n, p))


    update_sidebar()  # Update the sidebar initially

    
    # Dashboard Button at the bottom of the app_frame
    dashboard_button = ttk.Button(app_frame, text="Dashboard", command=show_dashboard, bootstyle="success", cursor="hand2", )
    dashboard_button.pack(side="bottom", fill="x", padx=5, pady=5)  # Added side='bottom' to position at bottom

    def update_selected_task(event):
        global selected_task
        selected_task = task_combobox.get()

    task_combobox.bind("<<ComboboxSelected>>", update_selected_task)  # Update the selected task from dropdown


    # Function to update the status label
    def update_status_label():
        window_title = get_active_window_title()
        active_window, active_object, active_work = parse_window_title(window_title)
        status_text = f"Window: {active_window} > {active_work}\nTask: {selected_task}"
        status_label.config(text=status_text)
        root.after(1000, update_status_label)

    # Label to display the active window and selected task
    status_title = tk.Label(detail_frame, text="Active Window & Task", font=("Arial", 12, "bold"), bg="#b2ebf2")
    status_title.pack(anchor='w', padx=5, pady=5)

    status_label = tk.Label(detail_frame, text="", font=("Arial", 10), bg="#b2ebf2", padx=5, pady=5, justify="left", anchor='w')
    status_label.pack(fill=tk.X, padx=5, pady=5, expand=True)

    # Begin updating the status label
    update_status_label() 

    root.mainloop()  # Start the Tkinter event loop to keep the window running

# Run the task manager GUI in a separate thread and start tracking windows
if __name__ == "__main__":
    # Start tracking the active window in a separate thread
    tracking_thread = threading.Thread(target=track_active_window)
    tracking_thread.daemon = True
    tracking_thread.start()

    # Run the Tkinter GUI for task management
    task_manager_gui()