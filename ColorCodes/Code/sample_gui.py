import tkinter as tk
from tkinter import filedialog, ttk
import finalcode as fc
import threading
from tkinter import messagebox
from tqdm import tqdm
input_path = ""
preprocess_path = ""
output_path = ""
progress_bar_tk = None
def input_folder_path(entry):
    global input_path
    input_path = filedialog.askdirectory()
    print("input:" + str(input_folder_path))
    entry.delete(0, tk.END)
    entry.insert(0, input_path)

def preprocess_folder_path(entry):
    global preprocess_path
    preprocess_path = filedialog.askdirectory()
    print("preprocess:" + str(preprocess_path))
    entry.delete(0, tk.END)
    entry.insert(0, preprocess_path)

def output_folder_path(entry):
    global output_path
    output_path = filedialog.askdirectory()
    print("output:" + str(output_path))
    entry.delete(0, tk.END)
    entry.insert(0, output_path)

def process_data():
    global input_path, preprocess_path, output_path

    # Start progress bar (0% complete)
    progress_bar = tqdm(total=100, desc='Processing Data', unit='%')

    process_status = fc.convert_pdf_to_images(input_path, preprocess_path)

    # Update progress bar (50% complete)
    progress_bar.update(50)

    if process_status == True:
        status = fc.process_folder(preprocess_path, output_path)
        
        # Update progress bar (50% complete)
        progress_bar.update(50)
        # progress_bar_tk.start()
        # progress_bar_tk["value"] = 50  # Update progress to 50%

        if status == True:
            # progress_bar_tk["value"] = 100  # Update progress to 100%
            messagebox.showinfo('Process Finished', 'Process finished!')
    
    # Close progress bar
    progress_bar.close()
    # progress_bar_tk.stop()

# def process_data():
#     global input_path, preprocess_path, output_path
#     process_status = fc.convert_pdf_to_images(input_path, preprocess_path)
#     if process_status == True:
#         fc.process_folder(preprocess_path, output_path)

def compute():
    computation_thread = threading.Thread(target=process_data)
    computation_thread.start()

def generate_gui():
    root = tk.Tk()
     # Main folder path
    main_folder_label = tk.Label(root, text="Enter main folder path:")
    main_folder_label.grid(row=0, column=0, padx=10, pady=10)
    main_folder_entry = tk.Entry(root)
    main_folder_entry.grid(row=0, column=1, padx=10, pady=10)
    main_folder_button = tk.Button(root, text="Browse", command=lambda: input_folder_path(main_folder_entry))
    main_folder_button.grid(row=0, column=2, padx=10, pady=10)
     # Temp folder path
    temp_folder_label = tk.Label(root, text="Enter temp folder path:")
    temp_folder_label.grid(row=1, column=0, padx=10, pady=10)
    temp_folder_entry = tk.Entry(root)
    temp_folder_entry.grid(row=1, column=1, padx=10, pady=10)
    temp_folder_button = tk.Button(root, text="Browse", command=lambda: preprocess_folder_path(temp_folder_entry))
    temp_folder_button.grid(row=1, column=2, padx=10, pady=10)
     # Output folder path
    output_folder_label = tk.Label(root, text="Enter output folder path:")
    output_folder_label.grid(row=2, column=0, padx=10, pady=10)
    output_folder_entry = tk.Entry(root)
    output_folder_entry.grid(row=2, column=1, padx=10, pady=10)
    output_folder_button = tk.Button(root, text="Browse", command=lambda: output_folder_path(output_folder_entry))
    output_folder_button.grid(row=2, column=2, padx=10, pady=10)
    # button to start processing
    process_button = tk.Button(root, text="Process",width= 50, command= compute)
    process_button.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

    # progress bar
    # global progress_bar_tk
    # progress_bar_tk = ttk.Progressbar(root, orient="horizontal", length = 350, mode="determinate")
    # progress_bar_tk.grid(row=4, column=0, columnspan=3, padx=10, pady=10)
    
    root.mainloop()
generate_gui()

