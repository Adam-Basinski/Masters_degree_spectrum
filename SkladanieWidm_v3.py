import pandas as pd
import tkinter as tk
from tkinter import filedialog
from matplotlib import pyplot as plt

path_for_coefficients = "C:\\Users\\adamb\OneDrive\\Praca dyplomowa - II stopień\\Widma\\12.04\\coefficients.xlsx"
path_for_background = "C:\\Users\\adamb\OneDrive\\Praca dyplomowa - II stopień\\Widma\\23.02\\Background_23.02.xlsx"

Excel_sheet_names = ["2007254U5_01", "2007255U5_01",
                     "2007256U5_01", "2007257U5_01", "2007258U5_01"]
Multipliers = [0.83164, 1.62584, 1.0, 0.32186, 0.52580]

# Getting filepath command


def get_file_path():
    global filename
    filetypes = [
        ('Excel files', '*.xlsx'),
        ('CSV files', '*.csv'),
        ('all files', '*.*')
    ]
    filename = filedialog.askopenfilename(
        title='Open CSV file',
        filetypes=filetypes,
        initialdir='C:\\Users\\adamb\\OneDrive\\Praca dyplomowa - II stopień\\Widma'
    )
    window.destroy()


# spectrum prepare function
def Spectrum_prepare(filename: str, IsBackGround: bool = False):
    Spectrum_list = []
    for i in range(len(Excel_sheet_names)):
        Spectrum_list.append(pd.DataFrame)

    for i in range(len(Excel_sheet_names)):
        Spectrum_list[i] = pd.read_excel(filename, Excel_sheet_names[i])
        Spectrum_list[i].columns = ["WaveLength", "I[-]"]
        Spectrum_list[i].drop(index=Spectrum_list[i].index[:5], inplace=True)
        Spectrum_list[i]["I[-]"] = Spectrum_list[i]["I[-]"] * Multipliers[i]

        if i == 0:
            Spectrum_list[i].drop(
                index=Spectrum_list[i].index[-20:], inplace=True)
        elif i == 1:
            Spectrum_list[i].drop(
                index=Spectrum_list[i].index[:21], inplace=True)
            Spectrum_list[i].drop(
                index=Spectrum_list[i].index[-17:], inplace=True)
        elif i == 2:
            Spectrum_list[i].drop(
                index=Spectrum_list[i].index[:25], inplace=True)
            Spectrum_list[i].drop(
                index=Spectrum_list[i].index[-18:], inplace=True)
        elif i == 3:
            Spectrum_list[i].drop(
                index=Spectrum_list[i].index[:118], inplace=True)
            Spectrum_list[i].drop(
                index=Spectrum_list[i].index[-20:], inplace=True)
        elif i == 4:
            Spectrum_list[i].drop(
                index=Spectrum_list[i].index[:60], inplace=True)
        else:
            print("ERROR")
    if IsBackGround:
        return Spectrum_list
    elif IsBackGround == False:
        for i in range(len(Excel_sheet_names)):
            Spectrum_list[i]['I[-]'] = Spectrum_list[i]['I[-]'] - \
                Full_excel_background[i]['I[-]']
            SpecMinimum = Spectrum_list[i]['I[-]'].min()
            if Excel_sheet_names[i] == Excel_sheet_names[1]:
                SpecMinimum += 20
            Spectrum_list[i].loc[Spectrum_list[i]['I[-]'] < 0, 'I[-]'] -= SpecMinimum

    Spectrum = pd.concat(Spectrum_list)
    Spectrum.columns = ["WaveLength", "I[-]"]
    Spectrum.sort_values(by="WaveLength", ascending=True, inplace=True)
    Spectrum.set_index("WaveLength")

    return Spectrum


# This will create window to choose the file
window = tk.Tk()
window.title('')
window.resizable(True, True)
window['background'] = 'blue'
window.geometry('300x150')

open_button = tk.Button(
    window,
    text='Choose CSV file',
    activebackground='blue',
    activeforeground='green',
    command=get_file_path
)
open_button.pack(expand=True)
window.mainloop()

# BACKGROUND

Full_excel_background = Spectrum_prepare(
    path_for_background, IsBackGround=True)

# Creating excel
Excel = Spectrum_prepare(filename=filename)

plt.plot(Excel['WaveLength'], Excel['I[-]'])
plt.show()

# Saving to excel
Excel.to_excel(filename[:-5] + "_merged_calibrated" + ".xlsx", index=False)
