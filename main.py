import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
from openpyxl.workbook import Workbook
import re

root = tk.Tk()

canvas1 = tk.Canvas(root, width=300, height=300, bg='lightsteelblue2', relief='raised')
canvas1.pack()

label1 = tk.Label(root, text='SKU MAPPING GUI', bg='lightsteelblue2')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)


def getExcel():
    global read_file
    global import_file_path
    global exported_csv_file_path

    import_file_path = filedialog.askopenfilename()
    read_file = pd.read_excel(import_file_path)
    exported_csv_file_path = filedialog.asksaveasfilename(defaultextension='.csv')
    convert_to_csv = read_file.to_csv(exported_csv_file_path, index=0, header=True)


browseButton_Excel = tk.Button(text="  Unmapped Excel File", command=getExcel, bg='green', fg='white',
                               font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_Excel)


def getMappedSKU():
    global read_file
    global sku
    global exported_csv_file_path

    sku = pd.read_csv(exported_csv_file_path)

    x = sku['Distributor Item code']
    y = sku['Company Dscription']

    x = x.dropna()
    x = x.reset_index(drop=True)

    y = y.dropna()
    y = y.reset_index(drop=True)

    result = []
    item_code = []
    s = []
    for j in range(0, 10):

        dis_item_code = x[j]
        item_code.append(dis_item_code)

        for_search = dis_item_code.split()
        # result.append([])
        for i in range(len(y)):
            company_desc = y[i].strip()
            search_result = (re.split('[ ]|[-]', company_desc))

            if len(search_result) <= 1:
                continue
            elif len(search_result) > 1 and for_search[-1].lower() == search_result[1].lower():
                s.append(dis_item_code)
                result.append(company_desc)
                # result[j].append(company_desc)
    data = {'Distributor Item code': s, 'Company Dscription': result}
    final_mapped_sku = pd.DataFrame(data, columns=['Distributor Item code', 'Company Dscription'])
    export_file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    final_mapped_sku.to_excel(export_file_path, index=0, header=True)


mapped_sku_excel = tk.Button(text='Download mapped Excel file', command=getMappedSKU, bg='green', fg='white',
                             font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=mapped_sku_excel)


def exitApplication():
    MsgBox = tk.messagebox.askquestion('Exit Application', 'Are you sure you want to exit the application',
                                       icon='warning')
    if MsgBox == 'yes':
        root.destroy()


exitButton = tk.Button(root, text='       Exit Application     ', command=exitApplication, bg='brown', fg='white',
                       font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=exitButton)

root.mainloop()
