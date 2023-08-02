import pandas as pd
import tkinter as tk
from tkinter import filedialog
from matplotlib import pyplot as plt


path_for_background = "C:\\Users\\adamb\OneDrive\\Praca dyplomowa - II stopień\\Widma\\23.02\\Background_23.02.xlsx"

Excel_sheet_names = ["2007254U5_01", "2007255U5_01",
                     "2007256U5_01", "2007257U5_01", "2007258U5_01", ]
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


#
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

    """if IsBackGround:
        return Spectrum_list
    elif IsBackGround == False:
        for i in range(len(Excel_sheet_names)):
            Spectrum_list[i]['I[-]'] = Spectrum_list[i]['I[-]'] - \
                Full_excel_background[i]['I[-]']
            Spectrum_list[i].loc[Spectrum_list[i]['I[-]'] < 0, 'I[-]'] = 0"""
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

"""Full_excel_background = []
Excel_background_mean = []
for i in range(len(Full_excel_names)):
    Full_excel_background.append(pd.DataFrame)

for i in range(len(Full_excel_background)):
    Full_excel_background[i] = pd.read_excel(
        path_for_background, Full_excel_names[i])
    Full_excel_background[i].columns = ["WaveLength", "I[-]"]
    Full_excel_background[i].drop(
        index=Full_excel_background[i].index[:5], inplace=True)
    Excel_background_mean.append(Full_excel_background[i]['I[-]'].mean())"""

Full_excel_background = Spectrum_prepare(
    path_for_background, IsBackGround=True)


# SPECTRUM

"""Full_excel = []
for i in range(len(Full_excel_names)):
    Full_excel.append(pd.DataFrame)

for i in range(len(Full_excel)):
    Full_excel[i] = pd.read_excel(filename, Full_excel_names[i])
    Full_excel[i].columns = ["WaveLength", "I[-]"]
    Full_excel[i].drop(index=Full_excel[i].index[:5], inplace=True)
    Full_excel[i]['I[-]'] = Full_excel[i]['I[-]'] - Full_excel_background[i]['I[-]']
    Full_excel[i].loc[Full_excel[i]['I[-]'] < 0, 'I[-]'] = 0

# Merging spectrum pieces
Excel = pd.concat(Full_excel)
Excel.columns = ["WaveLength", "I[-]"]
Excel.sort_values(by="WaveLength", ascending=True, inplace=True)
Excel.set_index("WaveLength")"""

Excel = Spectrum_prepare(filename=filename, IsBackGround=False)

#plt.plot(Excel['WaveLength'], Excel['I[-]'])
# plt.show()

# Saving to excel
Excel.to_excel(filename[:-5] + "_merged_calibrated" + ".xlsx", index=False)
