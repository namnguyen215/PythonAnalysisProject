from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk
import tkinter.font as font
import threading
from PIL import Image, ImageTk
import sys
import os



def read_help_file(file_name):
    
    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    path_to_help = os.path.abspath(os.path.join(bundle_dir,file_name))
    return path_to_help

def insert_text(ele,txt):
    ele.configure(state='normal')
    ele.delete("1.0",END)
    ele.insert(END, input)
    ele.configure(state='disabled')
    
def import_file():
    global is_correct_type  
    global input
    filename = filedialog.askopenfilename()
    if filename.endswith('.ASC'):
        is_correct_type = True
        input=filename
        insert_text(path_text, input)
        run_btn["state"]="normal"
    elif len(filename) == 0:
        is_correct_type = False
    else:
        is_correct_type = False
        messagebox.showerror("Error",message="File đầu vào phải có đuôi .asc")
    
    print('Selected:', filename)  

def disable_button():
    add_btn["state"] = "disabled"
    run_btn["state"] = "disabled"

def enable_button():
    add_btn["state"] = "normal"
    run_btn["state"] = "normal"

def choose_save_file():
    files = [('PDF', '*.pdf')]
    file = filedialog.asksaveasfilename(filetypes = files, defaultextension = files)
    # file = filedialog.FileDialog(app)
    return file

def run_execute(file_type, open_path, save_path):    
    if file_type == "COG":
        import cog_code
        cog_code.set_path(open_path, save_path)
    elif file_type == "LVT":
        import lvt_code
        lvt_code.set_path(open_path, save_path)
    elif file_type == "SMK":
        import smk_code
        smk_code.set_path(open_path, save_path)
    elif file_type == "Signal":
        import signal_code
        signal_code.set_path(open_path, save_path)
    elif file_type == "RT":
        import rt_code
        rt_code.set_path(open_path, save_path)
    elif file_type == "DT":
        import dt_code
        dt_code.set_path(open_path, save_path)
    else:
        print("Hic!")

def execute():
    file_path = choose_save_file()
    if len(file_path) != 0:
        app.config(cursor='wait')
        disable_button()
        file_type=cb1.get()
        if is_correct_type is True:
            try:
                run_execute(file_type, input, file_path)
            except:
                messagebox.showerror("Error",message="Kiểm tra lại file đầu vào")
            else:
                messagebox.showinfo("Run",message="Lưu thành công tại vị trí "+file_path)
        else:
            messagebox.showerror("Error",message="Kiểm tra lại file đầu vào")
        app.config(cursor='')
        enable_button()

def do_start():
    worker = threading.Thread(target=execute)
    worker.start()
        

    
    
# Create window object
app = Tk()
buttonFont = font.Font(size="10", family="Arial", weight="bold")
labelFont = font.Font(size="10", family="Arial")
#Title Phần mềm chuẩn hóa, đánh giá kết quả kiểm tra tâm sinh lý

title=Label(app, text="PHẦN MỀM CHUẨN HÓA, ĐÁNH GIÁ KẾT QUẢ TÂM SINH LÝ", font=('bold', 12))
title.config(background="#e6e5ed")
title.grid(row=0, column=0, columnspan=4, pady=20)

#Choose file
options = [
    "COG",
    "LVT",
    "SMK",
    "DT",
    "RT",
    "Signal"
]
# choose_type_txt = StringVar()
choose_type_txt = Label(app, text="Loại báo cáo", font=labelFont)
choose_type_txt.config(background="#e6e5ed")
choose_type_txt.grid(row=1, column=0, pady=5)

cb1 = ttk.Combobox(app, values=options,width=20, state="readonly")
cb1.grid(row=1,column=1,pady=5, sticky="W")
cb1.set("COG")


#Import file label
import_label = Label(app, text='Nhập File', font=labelFont)
import_label.config(background="#e6e5ed")
import_label.grid(row=2, column=0, sticky=W, pady=5)

xscrollbar = Scrollbar(app, orient=HORIZONTAL)
xscrollbar.grid(row=3, column=1, sticky=N+E+W)

#Path imported
path_text = Text(app, height = 1, width = 35,wrap=NONE, xscrollcommand=xscrollbar.set)
path_text.configure(state='disabled')
path_text.grid(row=2, column=1, pady=5)

xscrollbar.config(command=path_text.xview)

#Import file buttons
add_btn = Button(app, text='Chọn file', width=16,height=1, font=buttonFont, bg="#20bebe", fg="white", command=import_file)
add_btn.grid(row=2, column=2, sticky=W, pady=5, padx=5)



#execute file buttons
run_btn = Button(app, text='Phân tích', width=22,height=2,
                 font=buttonFont, bg="#20bebe", fg="white", command=do_start)
run_btn.grid(row=3, column=1, pady=20)


#Logo text
logo_lable = Label(app, text='Được phát triển bởi\n công ty cổ phần công nghệ\n One Mount Analytics',
                   font=("light",8), pady=20)
logo_lable.config(background="#e6e5ed")
logo_lable.grid(row=4, column=2)

#Logo
logo_path = read_help_file("logo.jpg")
load = Image.open(logo_path)
resized_image= load.resize((50,50), Image.ANTIALIAS)
render = ImageTk.PhotoImage(resized_image)
logo = Label(app, image=render)
logo.grid(row=4, column=3)     


if __name__ == '__main__':
    app.title('Phần mềm')
    app.geometry('600x305')
    app.resizable(False, False)
    run_btn["state"]="disabled"
    app.configure(background='#e6e5ed')
    # Start program
    app.mainloop()
    # pyinstaller --onefile --windowed --add-data "resource/*;." --icon=resource/logo.jpg app.py