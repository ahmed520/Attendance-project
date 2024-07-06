import customtkinter as ctk
import os
import subprocess
import tkinter as tk
import tkinter.messagebox as msgbox


ctk.set_appearance_mode("Dark")
def terminate_process(process_name):
    answer = msgbox.askquestion("Exit", "Are you sure you want to exit?\n IF you exit and restart the program all data will be REMOVED...!")
    if answer == "yes":
        os.system(f"taskkill /f /im {process_name}.exe")
        window.destroy()

#window
window = ctk.CTk()
window.title("Attendance Application")
window.geometry('380x400')


#frame
frame = ctk.CTkFrame(window ,width = 500 , height = 700)
frame.pack(padx = 50 , pady = 30)

#label
label = ctk.CTkLabel(frame , text = 'Attendance' , text_color = 'white' , fg_color = 'transparent',font = ('Harlow Solid Italic' , 50))
label.pack()


#Button1
button_1 = ctk.CTkButton(frame , text = 'Start the program' , text_color = 'white' ,  hover_color = 'dodger blue', command = lambda:subprocess.call("C:\Attendance\Code\dist\Start.exe"))
button_1.pack(pady=20)
#Button2
button_2 = ctk.CTkButton(frame , text = 'Browse' , text_color = 'white' ,  hover_color = 'dodger blue', command =  lambda:subprocess.call("C:\Attendance\Code\dist\Browse\Browse.exe"))
button_2.pack(pady=20)

#Button2
button_2 = ctk.CTkButton(frame , text = 'Exit' , text_color = 'red' ,  hover_color = 'dodger blue', command = lambda:terminate_process("WNetWatcher"), font = ('Arial Black',14))
button_2.pack(pady=20)
window.mainloop()


#button_2 = ctk.CTkButton(frame , text = 'Browse' , text_color = 'white' ,  hover_color = 'dodger blue', command =  lambda:subprocess.call("C:\Attendance\Code\Browse.pyw"))
