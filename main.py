import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os

from load_files import load_data
from generate_reports_deliverable1 import create_excel_report
from generate_reports_deliverable2 import create_deliverable_2
from generate_pp_derivable3 import analyze_market_section, analyze_geography, generate_charts, generate_presentation

def run_gui():
    def select_folder():
        folder_selected = filedialog.askdirectory()
        folder_path.set(folder_selected)

    def execute_exercise3(final_df):
        # Perform analysis and generate charts
        update_progress("Performing analysis...")
        revenue_by_section = analyze_market_section(final_df)
        revenue_by_country = analyze_geography(final_df)

        # Save the charts in the Downloads folder
        output_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
        update_progress("Generating charts...")
        revenue_by_section_img, revenue_by_country_img = generate_charts(revenue_by_section, revenue_by_country, output_dir)
        update_progress(f"Charts generated in {output_dir}")

        # Generate the PowerPoint presentation
        update_progress("Generating PowerPoint presentation...")
        generate_presentation(output_dir, revenue_by_section_img, revenue_by_country_img, revenue_by_section, revenue_by_country)
        update_progress(f"Presentation generated in {output_dir}")

    def execute_actions():
        global data_path 
        data_path = folder_path.get()
        if not data_path:
            messagebox.showwarning("Warning", "Please select a data folder.")
            return

        # Check which options have been selected
        selected_options = []
        if var_ex1.get():
            selected_options.append(1)
        if var_ex2.get():
            selected_options.append(2)
        if var_ex3.get():
            selected_options.append(3)

        if not selected_options:
            messagebox.showwarning("Warning", "Please select at least one action.")
            return

        # Load data
        try:
            update_progress("Loading data...")
            final_df = load_data(data_path)
            update_progress("Data loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while loading the data: {e}")
            return

        # Execute the selected actions
        try:
            if 1 in selected_options:
                update_progress("Generating Excel Report (Deliverable 1)...")
                create_excel_report(final_df)
                update_progress("Excel report created successfully.")

            if 2 in selected_options:
                update_progress("Creating Deliverable 2...")
                create_deliverable_2(final_df)
                update_progress("Deliverable 2 completed.")

            if 3 in selected_options:
                update_progress("Executing Exercise 3...")
                execute_exercise3(final_df)
                update_progress("Exercise 3 completed.")
            
            messagebox.showinfo("Success", "The selected actions were completed successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while executing the actions: {e}")

    def update_progress(message):
        progress_label.config(text=message)
        root.update_idletasks()

    root = tk.Tk()
    root.title("FP&A Report Generator")
    root.geometry("700x500")
    root.resizable(False, False)

    # GUI variables
    folder_path = tk.StringVar()
    var_ex1 = tk.IntVar()
    var_ex2 = tk.IntVar()
    var_ex3 = tk.IntVar()

    # Main frame
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill='both', expand=True)

    # Frame for folder selection
    folder_frame = ttk.LabelFrame(main_frame, text="Select Data Folder", padding="10")
    folder_frame.pack(fill='x', padx=10, pady=10)
    
    ttk.Entry(folder_frame, textvariable=folder_path, width=50).pack(side='left', padx=10)
    ttk.Button(folder_frame, text="Browse", command=select_folder).pack(side='left', padx=10)

    # Frame for action selection
    action_frame = ttk.LabelFrame(main_frame, text="Select Actions to Perform", padding="10")
    action_frame.pack(fill='x', padx=10, pady=10)

    # Use tk.Checkbutton for wrapping text
    tk.Checkbutton(action_frame, text="Generate Excel file with Total transactions per client, per currency and Totals by client, in USD.",
                    variable=var_ex1, wraplength=550).pack(anchor='w', padx=10, pady=2)
    tk.Checkbutton(action_frame, text="Generate Excel file filtered for the Latin American leaders with the transactions on Latin American countries and A list of all clients that had transactions with Russia, Cuba, China or Venezuela for the last quarter of 2024.",
                    variable=var_ex2, wraplength=550).pack(anchor='w', padx=10, pady=2)
    tk.Checkbutton(action_frame, text="Executive Deck (PowerPoint).",
                    variable=var_ex3, wraplength=550).pack(anchor='w', padx=10, pady=2)

    # Frame for execution and progress
    exec_frame = ttk.Frame(main_frame, padding="10")
    exec_frame.pack(fill='x', padx=10, pady=10)

    ttk.Button(exec_frame, text="Execute", command=execute_actions).pack(pady=10)
    progress_label = ttk.Label(exec_frame, text="Waiting for action...", foreground="blue")
    progress_label.pack()

    root.mainloop()

def main():
    run_gui()

if __name__ == "__main__":
    main()
