import os
from tkinter import *
import pandas


root = Tk()

empty_link = "https://iris.noncd.db.de/wbt/js/index.html?bhf="

def check(e):
    global data
    v= entry.get()
    if v=='':
        data = train_stations
    else:
        data=[]
        for item in train_stations:
            if v.lower() in item.lower():
                data.append(item)
    update(data)


def update(data):
    # Clear the Combobox
    menu.delete(0, END)
    # Add values to the combobox
    for value in data:
        menu.insert(END,value)

 
def read_data():
    global ibnr, train_stations 
    df = pandas.read_excel(f"ibnr.xlsx") # The options of that method are quite neat; Stores to a pandas.DataFrame object
    idx_of_ibnr = 1-1 # in case the column of interest is the 1st in Excel
    ibnr = list(df.iloc[:,idx_of_ibnr])
    idx_of_train_station = 2-1
    train_stations = list(df.iloc[:,idx_of_train_station])


def open_link():
    c = menu.get(menu.curselection())
    menu.delete(0, END)
    data= train_stations
    index_temp = data.index(c)
    link = empty_link + str(ibnr[index_temp])
    for value in data:
        menu.insert(END,value)
    os.system("start " + link)

read_data()
label = Label(root, text="IBNR-Finder")
label.grid(row=0, column=0, columnspan=2)
entry = Entry(root, width=35)
entry.bind("<KeyRelease>", check)
entry.grid(row=1, column=0, columnspan=2, padx=5, pady=5)
menu = Listbox(root)
menu.grid(row=2, column=0, columnspan=2)
b_submit = Button(root, text="Submit", command=open_link)
b_submit.grid(row=3, column=0, columnspan=2, pady=5)

for value in train_stations:
    menu.insert(END,value)


root.mainloop()