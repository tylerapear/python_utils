import sys, os, shutil, subprocess, platform, threading, traceback, win32com.client
from tkinter import Tk, Label, Button, filedialog, messagebox, StringVar, Frame, Entry
from docx2pdf import convert
from PyPDF2 import PdfMerger

dot_count = 0
running = False
output_file_name = ""

def run_conversion(input_folder, output_folder):
    
    # RETRIEVE GLOBAL RUNNING VAR #
    global running
    
    # SET STATUS PERCENT DISPLAY TO 0 %
    status_percent.config(text=f"0%")

    # CREATE TEMP PDF DIRECTORY #
    home_dir = os.path.expanduser("~")
    temp_dir = os.path.join(home_dir, "temp_dir")
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    # CREATE ARRAY OF NON-TEMP .DOCX FILES #
    word_files = [
      f for f in os.listdir(input_folder) 
      if f.lower().endswith(".docx") and not f.startswith("~$")
    ]
    total_files = len(word_files)

    # START AN INSTANCE OF MICROSOFT WORD #
    word_app = win32com.client.Dispatch("Word.Application")

    # CONVERT EACH WORD FILE TO A TEMP PDF
    for i, file in enumerate(word_files, start=1):
      if file.lower().endswith(".docx"):
          
        # Remove "Protected View" from file
        unblock_file(os.path.join(input_folder, file))
        
        # Define filepaths
        original_src = os.path.abspath(os.path.join(input_folder, file))
        pdf_path = os.path.join(temp_dir, file.replace(".docx", ".pdf"))
        
        # Try to convert file
        try:
          print("opening")
          doc = word_app.Documents.Open(original_src)
          print('opened')
          try:
            doc.ExportAsFixedFormat(pdf_path, 17)
          except Exception as e:
            print(f"Word export failed for {file}")
            print(e)
            traceback.print_exc()
            raise
          doc.Close(False)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert {file}: {e}")
            continue
            
      # Update progress display
      percent = int((i / total_files) * 100)
      status_percent.config(text=f"{percent}%")
      root.update_idletasks()  # ensures GUI updates

    # MERGE TEMP PDFS INTO COMBINED PDF #
    merger = PdfMerger()
    for file in sorted(os.listdir(temp_dir)):
        if file.lower().endswith(".pdf"):
            merger.append(os.path.join(temp_dir, file))
    merger.write(os.path.join(output_folder, f"{output_file_var.get().strip()}.pdf"))
    merger.close()

    # REMOVE TEMP PDF DIRECTORY
    shutil.rmtree(temp_dir)

    running = False  # stops animation
    status_label.config(text="Done!")
    
    # Clear input
    input_folder_var.set("")
    output_folder_var.set("")
    output_file_var.set("")
    
    input_button.config(state="normal")
    output_button.config(state="normal")
    filename_input.config(state="normal")
    
    show_done_message(output_folder)
    
def unblock_file(filepath):
  if os.name == 'nt':
    try:
      zone_file = filepath + ":Zone.Identifier"
      if os.path.exists(zone_file):
        os.remove(zone_file)
    except Exception as e:
      print(f"Failed to unblock {filepath}: {e}")

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
    if messagebox.askyesno("Done", f"All Microsoft Word files have been combined into a PDF.\n\nFile '{output_file_name}.pdf Created.'\n\nOpen file location?"):
        open_folder(output_folder)

def open_folder(path):
    if platform.system() == "Windows":
        os.startfile(path)
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", path])
    else:  # Linux
        subprocess.run(["xdg-open", path])

def update_create_button(*args):
  if input_folder_var.get().strip() and output_folder_var.get().strip() and output_file_var.get().strip():
    create_button.config(state="normal")
  else:
    create_button.config(state="disabled")

def convert_files():
    global running
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    global output_file_name
    output_file_name = output_file_var.get().strip()

    home_dir = os.path.expanduser("~")
    temp_dir = os.path.join(home_dir, "temp_dir")
    if not input_folder or not output_folder or not output_file_name:
        messagebox.showwarning("Missing Input", "Please select input/output folders and provide a file name.")
        return
      
    combined_pdf_path = os.path.join(output_folder, output_file_name + ".pdf")
    
    if os.path.exists(combined_pdf_path):
      overwrite = messagebox.askyesno(
        "File exists",
        f"The file '{output_file_name}.pdf' already exists in the ouput folder.\nDo you want to overwrite it?"
      )
      if not overwrite:
        return
      
    # Disable inputs while running
    create_button.config(state="disabled")
    input_button.config(state="disabled")
    output_button.config(state="disabled")
    filename_input.config(state="disabled")
    
    # Start animation
    running = True
    animate_loading()
    root.update()  # make sure the first frame shows
    
    threading.Thread(target=run_conversion, args=(input_folder, output_folder), daemon=True).start()

# --- Tkinter GUI setup ---
root = Tk()
root.title("Microsoft Word to PDF Converter")

if getattr(sys, 'frozen', False):
    # Running as EXE
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

icon_path = os.path.join(base_path, "app_icon.ico")
try:
    root.iconbitmap(icon_path)
except Exception as e:
    print(f"Failed to set icon: {e}")

center_window(root, 500, 300)

input_folder_var = StringVar()
output_folder_var = StringVar()
output_file_var = StringVar(value="")

output_file_var.trace_add("write", update_create_button)
output_folder_var.trace_add("write", update_create_button)
input_folder_var.trace_add("write", update_create_button)

input_frame_title = Frame(root)
input_frame_input = Frame(root)
input_frame_output = Frame(root)
input_frame_name_label = Frame(root)
input_frame_name = Frame(root)
input_frame_create = Frame(root)
input_frame_status = Frame(root)
input_frame_title.grid(row=0, column=0, padx=10, pady=5, sticky="w")
input_frame_input.grid(row=1, column=0, padx=10, pady=5, sticky="w")
input_frame_output.grid(row=2, column=0, padx=10, pady=5, sticky="w")
input_frame_name_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
input_frame_name.grid(row=4, column=0, padx=10, pady=5, sticky="w")
input_frame_create.grid(row=5, column=0, padx=10, pady=5, sticky="w")
input_frame_status.grid(row=6, column=0, padx=10, pady=5, sticky="w")

status_label = Label(input_frame_status, text="Select an input folder, output folder, and filename, then click 'Create Combined PDF'.", anchor="w")
status_percent = Label(input_frame_status, text="", anchor="w")
create_button = Button(
  input_frame_create, 
  text="Create Combined PDF", 
  command=convert_files
)
create_button.config(state="disabled")

input_button = Button(
  input_frame_input, 
  text="Select Input Folder", 
  command=select_input_folder
)

output_button = Button(
  input_frame_output, 
  text="Select Output Folder", 
  command=select_output_folder
)

filename_input = Entry(
  input_frame_name,
  textvariable=output_file_var, 
  width=30
)

Label(
    input_frame_title, 
    text="Select Folder with Microsoft Word Files:", 
    font=("Arial", 12, "bold")
).pack(side="left")

input_button.pack(side="left")

Label(
    input_frame_input, 
    textvariable=input_folder_var, 
    fg="blue", 
    wraplength=400
).pack(side="left")

output_button.pack(side="left")

Label(
    input_frame_output, 
    textvariable=output_folder_var, 
    fg="blue", 
    wraplength=400
).pack(side="left")

Label(
    input_frame_name_label, 
    text="What would you like to name the combined PDF?:", 
    wraplength=400
).pack(side="left")

filename_input.pack(side="left")

Label(
  input_frame_name,
  text=".pdf"
).pack(side="left")

create_button.pack(side="left")

status_percent.pack(side="left")
status_label.pack(side="left")

root.mainloop()