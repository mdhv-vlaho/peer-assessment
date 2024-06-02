import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import glob
import logging
from datetime import datetime
from forms_data_sorter import classlist_creation, assessment_data, email_prep, makecsv_tutpartic, feedback_data_shuttle

# Set up logging
logging.basicConfig(filename='email_log.log', level=logging.INFO, format='%(asctime)s - %(message)s')

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def update_progress(status, progress):
    status_label.config(text=status)
    progress_bar['value'] = progress
    root.update_idletasks()

def toggle_test_email():
    if test_run_var.get():
        testrunemail_label.grid(row=4, column=0)
        testrunemail_entry.grid(row=4, column=1)
    else:
        testrunemail_label.grid_remove()
        testrunemail_entry.grid_remove()

def reset_test_run():
    test_run_var.set(False)
    toggle_test_email()
    status_label.config(text="")
    progress_bar['value'] = 0

def process_feedback():
    try:
        print("Button clicked, starting process...")
        update_progress("Accessing files...", 0)
        tutorial_date = tutorial_date_entry.get()
        course_number = course_number_entry.get()
        is_test_run = test_run_var.get()
        test_run_email = testrunemail_entry.get() if is_test_run else None
        file_path = file_entry.get()

        if not (tutorial_date and course_number and file_path):
            raise ValueError("Tutorial date, course number, and file path are required")
        if is_test_run and not test_run_email:
            raise ValueError("Test run email is required for test run")

        update_progress("Processing started...", 10)
        print("Processing started...")

        parsed_date = datetime.strptime(tutorial_date, '%b %d')
        formatted_tutorial_date = parsed_date.strftime('%m-%d')
        term = '2024s'
        emailfrom = 'chem2X2.chemistry@mcgill.ca'
        server = 'mailhost.mcgill.ca'
        port = '25'

        update_progress("Assessing data...", 20)
        print("Assessing data...")
        feedback_dict, participation_grade = assessment_data(term, course_number, [test_run_email] if test_run_email else [], file_path)

        update_progress("Preparing emails...", 50)
        print("Preparing emails...")
        email_prep(feedback_dict, tutorial_date, emailfrom, course_number, test_run_email, server, port, testrun=is_test_run)

        update_progress("Creating CSV for participation...", 70)
        print("Creating CSV for participation...")
        makecsv_tutpartic(participation_grade, tutorial_date, course_number)

        if not is_test_run:
            newpath = f'/Users/deevlaho/Library/CloudStorage/OneDrive-McGillUniversity/Documents/PycharmProjects/peer-assessment/classlists/archive/{term}-{course_number}-peer_feedback-{formatted_tutorial_date}.xlsx'
            update_progress("Archiving feedback data...", 90)
            print("Archiving feedback data...")
            feedback_data_shuttle(file_path, newpath)

        update_progress("Process completed successfully!", 100)
        print("Process completed successfully!")
        messagebox.showinfo("Success", "Feedback processed and emails sent successfully!")
        reset_test_run()
    except Exception as e:
        update_progress(f"Error: {str(e)}", 0)
        print(f"Error: {str(e)}")
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Feedback Processor")
root.geometry('500x300')  # Set the initial size of the window

tk.Label(root, text="Tutorial Date (e.g., Jan 01)").grid(row=0, column=0)
tutorial_date_entry = tk.Entry(root)
tutorial_date_entry.grid(row=0, column=1)

tk.Label(root, text="Course Number").grid(row=1, column=0)
course_number_entry = tk.Entry(root)
course_number_entry.grid(row=1, column=1)

tk.Label(root, text="Is Test Run?").grid(row=2, column=0)
test_run_var = tk.BooleanVar()
tk.Checkbutton(root, variable=test_run_var, command=toggle_test_email).grid(row=2, column=1)

testrunemail_label = tk.Label(root, text="Test Run Email")
testrunemail_entry = tk.Entry(root)

tk.Label(root, text="Select XLSX File").grid(row=3, column=0)
file_entry = tk.Entry(root)
file_entry.grid(row=3, column=1)
tk.Button(root, text="Browse", command=browse_file).grid(row=3, column=2)

tk.Button(root, text="Process Feedback", command=process_feedback).grid(row=5, column=0, columnspan=3)

status_label = tk.Label(root, text="", fg="black")
status_label.grid(row=6, column=0, columnspan=3)

progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress_bar.grid(row=7, column=0, columnspan=3)

root.mainloop()
