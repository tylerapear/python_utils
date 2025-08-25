import os
import shutil
import subprocess
import platform
import threading
from tkinter import Tk, Label, Button, filedialog, messagebox, StringVar, Frame
from docx2pdf import convert
from PyPDF2 import PdfMerger

dot_count = 0
running = False

def run_conversion(input_folder, output_folder):
    global running
    status_percent.config(text=f"0%")

    home_dir = os.path.expanduser("~")
    temp_dir = os.path.join(home_dir, "temp_dir")
    temp_word_dir = os.path.join(home_dir, "temp_word_dir")
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)
    if os.path.exists(temp_word_dir):
        shutil.rmtree(temp_word_dir)
    os.makedirs(temp_word_dir)

    word_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".docx")]
    total_files = len(word_files)

    for i, file in enumerate(word_files, start=1):
      if file.lower().endswith(".docx"):
          
          original_src = os.path.join(input_folder, file)
          temp_dst = os.path.join(temp_word_dir, file)
          shutil.copy2(original_src, temp_dst)

          pdf_path = os.path.join(temp_dir, file.replace(".docx", ".pdf"))
          try:
              convert(temp_dst, pdf_path)
          except Exception as e:
              messagebox.showerror("Error", f"Failed to convert {file}: {e}")
              continue
            
      # Update progress
      percent = int((i / total_files) * 100)
      status_percent.config(text=f"{percent}%")
      root.update_idletasks()  # ensures GUI updates

    merger = PdfMerger()
    for file in sorted(os.listdir(temp_dir)):
        if file.lower().endswith(".pdf"):
            merger.append(os.path.join(temp_dir, file))
    merger.write(os.path.join(output_folder, "CombinedPDF.pdf"))
    merger.close()

    shutil.rmtree(temp_word_dir)
    shutil.rmtree(temp_dir)

    running = False  # stops animation
    status_label.config(text="Done!")
    show_done_message(output_folder)

def animate_loading():
    global status_label
    global dot_count
    if not running:
        return

    dots = '.' * (dot_count % 4)  # cycles 0..3 dots
    status_label.config(text=f"Processing{dots}")
    dot_count += 1
    root.after(500, animate_loading)  # repeat every 500ms

def center_window(root, width=500, height=300):
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

def select_input_folder():
    folder = filedialog.askdirectory()
    if folder:
        input_folder_var.set(folder)

def select_output_folder():
    folder = filedialog.askdirectory()
    if folder:
        output_folder_var.set(folder)

def show_done_message(output_folder):
    if messagebox.askyesno("Done", "All Microsoft Word files have been converted to PDF.\n\nFile 'CombinedPDF.pdf Created.'\n\nOpen selected output folder?"):
        open_folder(output_folder)

def open_folder(path):
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", path])
    else:  # Linux
        subprocess.run(["xdg-open", path])

def convert_files():
    global running
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()

    home_dir = os.path.expanduser("~")
    temp_dir = os.path.join(home_dir, "temp_dir")
    if not input_folder:
        messagebox.showwarning("No folder selected", "Please select a folder with Word files.")
        return
    
    # Start animation
    running = True
    animate_loading()
    root.update()  # make sure the first frame shows
    
    threading.Thread(target=run_conversion, args=(input_folder, output_folder), daemon=True).start()

# --- Tkinter GUI setup ---
root = Tk()
root.title("Microsoft Word to PDF Converter")
center_window(root, 500, 300)

input_folder_var = StringVar()
output_folder_var = StringVar()

input_frame_title = Frame(root)
input_frame_input = Frame(root)
input_frame_output = Frame(root)
input_frame_create = Frame(root)
input_frame_status = Frame(root)
input_frame_title.grid(row=0, column=0, padx=10, pady=5, sticky="w")
input_frame_input.grid(row=1, column=0, padx=10, pady=5, sticky="w")
input_frame_output.grid(row=2, column=0, padx=10, pady=5, sticky="w")
input_frame_create.grid(row=3, column=0, padx=10, pady=5, sticky="w")
input_frame_status.grid(row=4, column=0, padx=10, pady=5, sticky="w")

status_label = Label(input_frame_status, text="Select an input and output folder, then click 'Create Combined PDF'.", anchor="w")
status_percent = Label(input_frame_status, text="", anchor="w")

Label(
    input_frame_title, 
    text="Select Folder with Microsoft Word Files:", 
    font=("Arial", 12, "bold")
).pack(side="left")

Button(
    input_frame_input, 
    text="Select Input Folder", 
    command=select_input_folder
).pack(side="left")

Label(
    input_frame_input, 
    textvariable=input_folder_var, 
    fg="blue", 
    wraplength=400
).pack(side="left")

Button(
    input_frame_output, 
    text="Select Output Folder", 
    command=select_output_folder
).pack(side="left")

Label(
    input_frame_output, 
    textvariable=output_folder_var, 
    fg="blue", 
    wraplength=400
).pack(side="left")

Button(
    input_frame_create, 
    text="Create Combined PDF", 
    command=convert_files
).pack(side="left")

status_percent.pack(side="left")
status_label.pack(side="left")

root.mainloop()