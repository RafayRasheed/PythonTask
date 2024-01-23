import tkinter as tk
from tkinter import ttk, filedialog
import pandas as pd
import os
import json
import datetime
import time

width  = 1000
height = 500
def myWidth(val):
    global width
    return (val*width)/100

def myHeight(val):
    global height
    return (val*height)/100
def showHideProgressBar(show):

    return
def resett():
    result_label.config(text="", background="#FFFFFF")
    button_select_file.config(text="Select Excel File")
    progress_var.set(0)
    percentage_label.config(text=f"0%")

def get_default_download_path():
    # Get the user's home directory
    home_dir = os.path.expanduser("~")

    return os.path.join(home_dir, 'Downloads')
    
def load_data():
    databhej = None
    try:
        # Load data from the JSON file
        with open('data.json', 'r') as file:
            data = json.load(file)
            databhej = data.get('data', '')

    except FileNotFoundError:
        print("No data found.")
    return databhej
def save_data(data_to_save):
    # Save data to a JSON file
    with open('data.json', 'w') as file:
        json.dump({'data': data_to_save}, file)
   
def do_from_excel(file_path):
    df = pd.read_excel(file_path)
    length = len(df)
    selected_c=df['First Name']
    selected_Id=df['Id']
    divvv = 0
    
    def loop(index):
        nonlocal divvv
        nonlocal file_path
        if index < length:

            if selected_Id[index] > 3000:
                selected_c[index] = 'Yes'
            val = round(((index + 1) / length) * 100)
            update_progress(val)
            if(index<length-1):
                bata = int(val/5)
                if(bata>divvv  ):
                    divvv = bata
                    print(val)
                    app.after(1, loop, index + 1)
                else:
                    app.after(0, loop, index + 1)

                      # Schedule the next iteration with a delay
            else:
                destination_folder = label_Location.cget('text')
                file_name ='Ex'+'_'+datetime.datetime.now().strftime("%S%M%H%d%m%Y")+'.xlsx'
                file_path_update = os.path.join(destination_folder, file_name)
                df.to_excel(file_path_update, index=False)
                result_label.config(text=f"Save {file_name} to '{destination_folder}'", background="#2c3e50")



    loop(0)  # Start the loop

    # # Wait for the loop to complete before saving the DataFrame to Excel
    # app.after(1, lambda: df.to_excel('ttt.xlsx', index=False))

    # for i in range(0,length):
    #     # update_progress(val)
    #     if(selected_Id[i]>3000):
    #         selected_c[i] = 'Yes'

    #     val= round(((i+1)/length)*100)
    #     if(val==50):
    #         app.after(1)
    
    return
def on_button_click():
    if(progress_var.get()==0):
        file_path = ask_for_excel_file()
        if file_path:
            result_label.config(text=f"File : {file_path.split("/")[-1]}", background="#2c3e50")
            button_select_file.config(text="CLEAR")
            do_from_excel(file_path)
    else:
        resett()

def ask_for_excel_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls")],
        title="Select Excel File"
    )
    return file_path

def update_progress(new_value):
    progress_var.set(new_value)
    percentage_label.config(text=f"{new_value}%")
    
def update_from_save_data():
    daaaata = load_data()
    if(daaaata):
        label_Location.config(text=daaaata['download'])
    else:
        doPath = get_default_download_path()
        save_data({'download':doPath})
        label_Location.config(text=doPath)

def on_loc_button():
    folder_selected = filedialog.askdirectory(title="Select Destination Folder")
    if(folder_selected):
        daaaata = load_data()
        daaaata['download'] = folder_selected
        label_Location.config(text=folder_selected)
        save_data(daaaata)


showProgess = False
app = tk.Tk()
app.minsize(width=width, height=height)
app.maxsize(width=width, height=height)

screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
x_position = (screen_width - width) // 2
y_position = ((screen_height - height) // 2)-35

app.title("Excel File Selector")
app.geometry(f"{width}x{height}+{x_position}+{y_position}")  # Set the initial size of the window
style = ttk.Style()
style.configure("TButton", padding=10, relief="flat", background="#3498db", foreground="black")
style.configure("TLabel",padding=10, font=("Helvetica", 12), background="#ffffff", foreground="white")

progress_var = tk.IntVar()
progress_show =  tk.BooleanVar(value=False)
progress_bar = ttk.Progressbar(app, orient='horizontal',
                                mode='determinate', 
                                variable=progress_var,length=myWidth(70),
)
progress_bar.pack(pady=[myHeight(5),0])

# progress_bar.place_configure(x=10,y=50)
# Create a label to display the percentage
percentage_label = tk.Label(app, text="0%")
percentage_label.pack()
# Start the progress animation
# save_data({'data':'ga'})
# Create and pack widgets with ttk styling
button_select_file = ttk.Button(app, text="Select Excel File", command=on_button_click, style="TButton")
button_select_file.pack(pady=myHeight(5))


result_label = ttk.Label(app, text="",wraplength=myWidth(85), style="TLabel")
result_label.pack()

label_text_Download = "Download Path:"
label_Download = tk.Label(app, text=label_text_Download, )


label_text_Location =""
label_Location = tk.Label(app, text=label_text_Location, )
# Create Input Field (Entry) - Read-only


# Create Button
button = tk.Button(app, text="Change",command=on_loc_button )

# Pack widgets in a single row
label_Download.pack(side=tk.LEFT, padx=[myWidth(4),0])
label_Location.pack(side=tk.LEFT, padx=[0,5])
button.pack(side=tk.LEFT, padx=5)
update_from_save_data()

# Start the main event loop
app.mainloop()












