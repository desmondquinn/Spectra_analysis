import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import savgol_filter
from BaselineRemoval import BaselineRemoval
import tkinter as tk
from tkinter import filedialog

#initialise window
window = tk.Tk() 
window.title("Spectra analysis")

#method for browsing and selecting the path for the data to be imported
def browseFiles():
    global filename
    filename = filedialog.askopenfilename(initialdir = "/", title = "Select a File",
                                          filetypes = (("Excel file", "*.xlsx*"), ("all files", "*.*")))
    label.configure(text="File Opened: "+filename)
    browseFiles.name = filename 

label = tk.Label(window, text = "Hello")
window.geometry('600x180') 
button_exit = tk.Button(window, text = "exit", command = window.destroy) 
label.grid(column = 1, row = 1) 
button_exit.grid(column = 1,row = 13)
browseFiles()
path = browseFiles.name

#creates input fields to getting parameters from the user
tk.Label(window, text = 'Lower limit').grid(row=7) 
tk.Label(window, text = 'Upper limit').grid(row=9) 
tk.Label(window, text = 'No. of points').grid(row=11)
x1 = tk.Entry(window, bd = 5) 
x2 = tk.Entry(window, bd = 5)
x3 = tk.Entry(window, bd = 5)
x1.grid(row=7, column=1) 
x2.grid(row=9, column=1)
x3.grid(row=11, column=1)

def show_entry_fields():
    global y1,y2,y3
    y1 = x1.get()
    y2 = x2.get()
    y3 = x3.get()

tk.Button(window, text='confirm', command=show_entry_fields).grid(row=13, column=0, sticky=tk.W, pady=4)

#checkbox for smoothing operation 
z = 0
def show_checkbox_value():
    global z
    z = Checkbuttonvalue.get()

Checkbuttonvalue = tk.IntVar() 
Button2 = tk.Checkbutton(window, text = "Smoothing", command = show_checkbox_value,
                         variable = Checkbuttonvalue).grid(row = 15) 


window.mainloop()

a = int(y1)
b = int(y2)
df = pd.read_excel(path)
#importing the x and y values from the dataframe
lambda_values = df.iloc[:,0:1].values
input_array = df.iloc[:,1:2].values
#the input array has to be in the form of a list in order for the methods to work
input_array = input_array.transpose()
input_array = input_array.tolist()
input_array = input_array[0]

#to check and apply smoothing if required
if y3 == '' or z == 0:
    to_baseline_correct = input_array
else:
    s = int(y3)
    to_baseline_correct = savgol_filter(input_array, s, 2)

#baseline removal
baseObj=BaselineRemoval(to_baseline_correct)
Output=baseObj.ZhangFit()

#plotting the spectra
plt.title('Output spectra')
plt.plot(lambda_values, Output)
plt.xlim([a,b])
