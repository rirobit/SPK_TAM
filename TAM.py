import tkinter as tk
from tkinter import *
from tkinter import filedialog
import pandas as pd
from win32com.client import Dispatch
import pathlib
import time
from openpyxl import Workbook


def main():
    m = Tk()
    m.winfo_toplevel().title("Manajemen TAM")
    # list of menu options
    v = tk.IntVar()
    # create a list of Radio buttons
    tk.Radiobutton(m, text='INPUT DATA', variable=v, value=0, indicatoron=0, width=40, bg='grey', fg='black',
                   command=lambda: import_data()).pack()
    tk.Radiobutton(m, text='SELEKSI CRITICAL EQUIPMENT', variable=v, value=1, indicatoron=0, width=40, bg='grey',
                   fg='black', command=lambda: critical_selection()).pack()
    tk.Radiobutton(m, text='SELEKSI RE', variable=v, value=2, indicatoron=0, width=40, bg='grey', fg='black',
                   command=lambda: cek_re()).pack()
    tk.Radiobutton(m, text='SELEKSI NON TA TASK', variable=v, value=3, indicatoron=0, width=40, bg='grey', fg='black',
                   command=lambda: pm_selection()).pack()
    tk.Radiobutton(m, text='FORMAT INPUT', variable=v, value=4, indicatoron=0, width=40, bg='grey', fg='black',
                   command=lambda: contoh()).pack()
    tk.Radiobutton(m, text='PERBARUI PENGETAHUAN DATA KRITIKAL', variable=v, value=5, indicatoron=0, width=40,
                   bg='grey', fg='black', command=lambda: update_critical()).pack()
    tk.Radiobutton(m, text='FORMAT PENGETAHUAN DATA KRITIKAL', variable=v, value=6, indicatoron=0, width=40, bg='grey',
                   fg='black', command=lambda: contoh_critical_equipment()).pack()
    tk.Radiobutton(m, text='PERBARUI PENGETAHUAN DATA PM', variable=v, value=7, indicatoron=0, width=40,
                   bg='grey', fg='black', command=lambda: update_pm()).pack()
    tk.Radiobutton(m, text='FORMAT PENGETAHUAN DATA PM', variable=v, value=8, indicatoron=0, width=40,
                   bg='grey', fg='black', command=lambda: contoh_pm()).pack()
    tk.Radiobutton(m, text='LIHAT DATA', variable=v, value=9, indicatoron=0, width=40, bg='grey', fg='black',
                   command=lambda: view_data()).pack()
    tk.Radiobutton(m, text='SIMPAN', variable=v, value=10, indicatoron=0, width=40, bg='grey', fg='black',
                   command=lambda: save()).pack()
    tk.Radiobutton(m, text='KELUAR', variable=v, value=11, indicatoron=0, width=40, bg='grey', fg='black',
                   command=lambda: m.destroy()).pack()
    m.geometry('350x310')
    m.configure(background='LightYellow2')
    m.mainloop()
    return v.get()


def import_data():
    # select a file to import
    import_filename = filedialog.askopenfilename(title="Select a file", filetypes=(("excel files", "*.xlsx"),
                                                                                   ("all files", "*.*")))
    global dataTA
    dataTA = pd.read_excel(import_filename, index_col='No')  # import file
    global backupTA

    if len(dataTA) == 0:
        dataTA = pd.read_excel(import_filename, index_col='No', sheet_name='Data TA')
        backupTA = pd.read_excel(import_filename, index_col='No', sheet_name='Data TA')
    else:
        backupTA = pd.read_excel(import_filename, index_col='No')


def critical_selection():
    noncritical.clear()
    index = dataTA.index
    unique = set()
    total = len(index)
    start = time.time()
    for i in index:
        if dataTA["Equipment"][i] in critical.keys():
            continue
        else:
            noncritical.append(i)
            unique.add(dataTA["Equipment"][i])
    stop = time.time()
    durasi = stop - start
    lunique = "Equipment unik: " + str(len(unique))
    label = "Data non critical: " + str(len(noncritical))
    title = "Data Non Critical Equipment"
    view_data_select(noncritical, label, title, total, lunique, durasi)


def cek_re():
    nonrc.clear()
    index = dataTA.index
    unique = set()
    total = len(index)
    start = time.time()
    for i in index:
        if dataTA["Equipment"][i] in critical.keys():
            if dataTA["RE"][i] > critical[dataTA["Equipment"][i]]:
                continue
            else:
                nonrc.append(i)
                unique.add(dataTA["Equipment"][i])
        else:
            nonrc.append(i)
    stop = time.time()
    durasi = stop - start
    lunique = "Equipment unik: " + str(len(unique))
    label = "Data non RE < RC: " + str(len(nonrc))
    title = "Data RE < RC"
    view_data_select(nonrc, label, title, total, lunique, durasi)


def pm_selection():
    pmtas.clear()
    index = dataTA.index
    unique = set()
    total = len(index)
    start = time.time()
    for i in index:
        eq = ""
        op = ""
        ma = ""
        if pd.isna(dataTA["Equipment"][i]):
            eq = ""
        else:
            eq = (dataTA["Equipment"][i])

        if pd.isna(dataTA["Oper.WorkCenter"][i]):
            op = ""
        else:
            op = dataTA["Oper.WorkCenter"][i]

        if pd.isna(dataTA["MaintActivType"][i]):
            ma = ""
        else:
            ma = str(dataTA["MaintActivType"][i])

        unik = str(eq) + str(op) + str(ma)

        if unik in dataPM:
            pmtas.append(i)
            unique.add(unik)

    stop = time.time()
    durasi = stop - start
    lunique = "Equipment unik: " + str(len(unique))
    label = "Data non TA: " + str(len(pmtas))
    title = "Data Non TA"
    view_data_select(pmtas, label, title, total, lunique, durasi)


def view_data_select(li, labell, titl, total, lunique, durasi):
    head = list(dataTA.columns)
    m2 = Tk()
    m2.winfo_toplevel().title(titl)
    lhead = [14, 16, 19, 18, 30, 50, 50, 12, 10, 20, 8, 21, 9, 27]
    header = ""
    for h in range(10):
        header += head[h].ljust(lhead[h])

    lheader = Label(
        m2,
        text=header
    )
    lheader.pack()

    yscrollbar = Scrollbar(m2)
    yscrollbar.pack(side=RIGHT, fill=Y)
    lis = Listbox(m2, selectmode="multiple", yscrollcommand=yscrollbar.set, width=268)
    lis.pack(expand=YES, fill="both")

    lhead = [15, 20, 30, 21, 30, 50, 50, 10, 8, 17, 6, 19, 7, 25]

    for i in range(len(li)):
        print(i)
        tex = ""
        for h in range(len(head)):
            print(len(head))
            tex += str(dataTA[head[h]][li[i]]).ljust(lhead[h])
        lis.insert(END, tex)
        # coloring alternative lines of listbox
        lis.itemconfig(i, bg="white" if i % 2 == 0 else "cyan")

    total = Label(
        m2,
        text="Total data: " + str(total)
    )
    total.pack()

    seleksi = Label(
        m2,
        text=labell
    )
    seleksi.pack()

    unik = Label(
        m2,
        text=lunique
    )
    unik.pack()

    durasi = Label(
        m2,
        text="Durasi: " + str("{:.4f}".format(durasi)) + " detik"
    )
    durasi.pack()

    deleteall = Button(
        m2,
        text='HAPUS SEMUA',
        command=lambda: delete_all(li, m2)
        )
    deleteall.pack()

    deletepick = Button(
        m2,
        text='HAPUS YANG DIPILIH',
        command=lambda: delete_pick(li, lis, m2)
        )
    deletepick.pack()

    deletenotpick = Button(
        m2,
        text='HAPUS YANG TIDAK DIPILIH',
        command=lambda: delete_notpick(li, lis, m2)
        )
    deletenotpick.pack()

    kembali = Button(
        m2,
        text='KEMBALI',
        command=lambda: m2.destroy()
        )
    kembali.pack()

    m2.mainloop()


def delete_all(li, mm):
    for i in li:
        dataTA.drop(index=i, inplace=True)
    mm.destroy()


def delete_pick(li, lis, mm):
    select = lis.curselection()

    for i in select:
        dataTA.drop(index=li[i], inplace=True)

    mm.destroy()


def delete_notpick(li, lis, mm):
    select = lis.curselection()

    for i in select:
        del li[i]

    for i in li:
        dataTA.drop(index=i, inplace=True)

    mm.destroy()


def view_data():
    head = list(dataTA.columns)
    li = dataTA.index
    m2 = Tk()
    m2.winfo_toplevel().title("DATA TAM MANAJEMEN")
    lhead = [14, 16, 19, 18, 30, 50, 50, 12, 10, 20, 8, 21, 9, 27]
    header = ""
    for h in range(10):
        header += head[h].ljust(lhead[h])

    lheader = Label(
        m2,
        text=header
    )
    lheader.pack()

    yscrollbar = Scrollbar(m2)
    yscrollbar.pack(side=RIGHT, fill=Y)
    lis = Listbox(m2, selectmode="multiple", yscrollcommand=yscrollbar.set, width=268)
    lis.pack(expand=YES, fill="both")

    lhead = [15, 20, 30, 21, 30, 50, 50, 10, 8, 17, 6, 19, 7, 25]
    for i in range(len(li)):
        tex = ""
        for h in range(len(head)):
            tex += str(dataTA[head[h]][li[i]]).ljust(lhead[h])
        lis.insert(END, tex)
        # coloring alternative lines of listbox
        lis.itemconfig(i, bg="white" if i % 2 == 0 else "cyan")

    total = Label(
        m2,
        text="Total data: " + str(len(dataTA))
    )
    total.pack()

    kembali = Button(
        m2,
        text='KEMBALI',
        command=lambda: m2.destroy()
        )
    kembali.pack()

    m2.mainloop()


def contoh():
    path = str(pathlib.Path().resolve()) + r"\Format TAM.xlsx"
    writer = pd.ExcelWriter(path)
    header = ['No', 'Order Type', 'Equipment', 'Main WorkCtr', 'Oper.WorkCenter', 'MaintActivType', 'Description',
              'Opr. short text', 'Priority', 'Number', 'Normal duration', 'Normal duration', 'Work', 'Tahun TA', 'RE']

    datacontoh = pd.DataFrame(columns=header)

    datacontoh.to_excel(writer, index=False)
    x1 = Dispatch("Excel.Application")
    x1.Visible = True
    writer.save()
    x1.Workbooks.Open(path)


def contoh_pm():
    path = str(pathlib.Path().resolve()) + r"\Format PM.xlsx"
    writer = pd.ExcelWriter(path)
    header = ['Equipment', 'Oper.WorkCenter', 'MaintActivType']

    datacontoh = pd.DataFrame(columns=header)

    datacontoh.to_excel(writer, index=False)
    x1 = Dispatch("Excel.Application")
    x1.Visible = True
    writer.save()
    x1.Workbooks.Open(path)


def contoh_critical_equipment():
    path = str(pathlib.Path().resolve()) + r"\Format Critical Analysis.xlsx"
    writer = pd.ExcelWriter(path)
    header = ['Equipment', 'ECR']
    datacontoh = pd.DataFrame(columns=header)

    datacontoh.to_excel(writer, index=False)
    x1 = Dispatch("Excel.Application")
    x1.Visible = True
    writer.save()
    x1.Workbooks.Open(path)


def save():
    path = str(pathlib.Path().resolve()) + r"\result TAM.xlsx"
    writer = pd.ExcelWriter(path)
    header = list(dataTA.columns)
    datanoncritical = pd.DataFrame(columns=header)
    datanonre = pd.DataFrame(columns=header)
    datanonta = pd.DataFrame(columns=header)

    if len(nonrc) != 0:
        for i in nonrc:
            p = len(datanonre)
            datanonre.loc[p] = list(backupTA.loc[i]).copy()

    if len(pmtas) != 0:
        for i in pmtas:
            p = len(datanonta)
            datanonta.loc[p] = list(backupTA.loc[i]).copy()

    if len(noncritical) != 0:
        for i in noncritical:
            p = len(datanoncritical)
            datanoncritical.loc[p] = list(backupTA.loc[i]).copy()

    dataTA.to_excel(writer, 'data TA')
    datanoncritical.to_excel(writer, 'non critical')
    datanonre.to_excel(writer, 'data non RE')
    datanonta.to_excel(writer, 'data no TA')
    x1 = Dispatch("Excel.Application")
    x1.Visible = True
    writer.save()
    x1.Workbooks.Open(path)


def update_pm():
    importfilename = filedialog.askopenfilename(
        title="Select a file",
        filetypes=(("excel files", "*.xlsx"), ("all files", "*.*"))
    )
    df = pd.read_excel(importfilename)
    if df.empty:
        pass
    else:
        print("PM: ", dataPM)
        dataPM.clear()
        print("PM: ", dataPM)

        for i in range(len(df)):
            eq = ""
            op = ""
            ma = ""
            if pd.isna(df["Equipment"][i]):
                eq = ""
            else:
                eq = (df["Equipment"][i])

            if pd.isna(df["Oper.WorkCenter"][i]):
                op = ""
            else:
                op = df["Oper.WorkCenter"][i]

            if pd.isna(df["MaintActivType"][i]):
                ma = ""
            else:
                ma = str(df["MaintActivType"][i])

            unik = str(eq) + str(op) + str(ma)
            dataPM.add(unik)
        print(len(dataPM))
        print("PM: ", dataPM)



def update_critical():
    import_filename = filedialog.askopenfilename(title="Select a file", filetypes=(("excel files", "*.xlsx"),
                                                                                   ("all files", "*.*")))
    df = pd.read_excel(import_filename, sheet_name="Sheet1", header=[0, 1, 2])

    if df.empty:
        print('kosong')
    else:
        print("critical = ", len(critical))
        critical.clear()
        print("critical = ",  len(critical))
        for i in range(len(df)):
            critical[df['Equipment'][i]] = df['ECR'][i]
        print("critical = ", len(critical))


dataPM = set(['D1-FV-404-K3D0123106', 'D1-FSLL-111-K3D0123106', 'D1-C-601-K3D0123129', 'D1-LT-821B-K3D0123106', 'D1-LT-502-K3D0123106', 'D1-PSLL-503-K3D0123106', 'D1-PT-108-K3D0123106', 'D0133129', 'D1-XV-105-K3D0123106', 'D1-LSH-208-K3D0123106', 'D1-LV-308-K3D0123106', 'D1-P-301A-K3D0149106', 'D1-PM-301A-K3D0133106', 'D1-PT-429-K3D0123106', 'D1-HV-502-K3D0123106', 'D1-PSH-106-K3D0123106', 'D1-LV-204-K3D0123106', 'D1-XV-301-K3D0123106', 'D1-LSH-403-K3D0123106', 'D1-TV-535-K3D0123106', 'D1-LSL-202-K3D0123106', 'D1-TSH-126-K3D0123106', 'D1-TSH-318-K3D0123106', 'D1-FT-402-K3D0123106', 'D1-FV-502-K3D0123106', 'D1-PM-502B-K3D0133106', 'D1-FT-101-K3D0123106', 'D1-P-602B-K3D0149106', 'D1-LT-309-K3D0123106', 'D1-FSL-102-K3D0123106', 'D1-FT-303-K3D0123106', 'D1-LSH-311-K3D0123106', 'D0149129', 'D1-FV-405-K3D0123106', 'D1-PT-503-K3D0123106', 'D1-PM-501A-K3D0133106', 'D1-FSLL-303-K3D0123106', 'D1-PV-509-K3D0123106', 'D1-PT-207-K3D0123106', 'D1-TSL-104-K3D0123106', 'D1-K-405-P1A-K3D0149106', 'D1-PSL-902-K3D0123106', 'D1-PSH-116-K3D0123106', 'D1-TV-326-K3D0123106', 'D1-PSLL-809B-K3D0123106', 'D1-LV-311-K3D0123106', 'D1-LSH-203-K3D0123106', 'D1-FV-408-K3D0123106', 'D1-LT-831A-K3D0123106', 'D-OHC-20T-1-X-401D0101106', 'D1-FT-406-K3D0123106', 'D1-LT-311-K3D0123106', 'D1-XV-302-K3D0123106', 'D1-K-402P1A-K3D0149106', 'D1-PT-412-K3D0123106', 'D1-PT-104B-K3D0123106', 'D1-TSH-208-K3D0123106', 'D1-P-303-K3D0149106', 'D1-FT-117-K3D0123106', 'D1-LSHH-416-K3D0123106', 'D1-LSH-831B-K3D0123106', 'D1-LV-301-K3D0123106', 'D1-FV-303-K3D0123106', 'D1-PM-503B-K3D0133106', 'D1-LSH-402-K3D0123106', 'D1-GM-301-K3D0133106', 'D1-FV-304-K3D0123106', 'D1-P-502B-K3D0149106', 'D1-PV-120-K3D0123106', 'D1-TT-104-K3D0123106', 'D1-TSH-401-K3D0123106', 'D1-FT-107-K3D0123106', 'D1-XV-104-K3D0123106', 'D1-LSL-308-K3D0123106', 'DPIPING3-101-D0109D0123129', 'D1-FT-121-K3D0123106', 'D1-TT-310-K3D0123106', 'D1-P-602A-K3D0149106', 'D1-TSHH-408-K3D0123106', 'D1-TSHH-404-K3D0123106', 'D1-TV-423-K3D0123106', 'D1-LSH-103-K3D0123106', 'D1-PDSH-413-K3D0123106', 'D1-TSH-315-K3D0123106', 'D1-FSL-108-K3D0123106', 'D1-TSHH-426-K3D0123106', 'D1-FSL-107-K3D0123106', 'D1-P-503B-K3D0149106', 'D1-LT-201-K3D0123106', 'D1-TV-506-K3D0123106', 'D1-PT-407-K3D0123106', 'D1-LSHH-418-K3D0123106', 'D1-TV-104-K3D0123106', 'D1-G-204-K3D0149106', 'D1-PT-809-K3D0123106', 'D1-TSH-104-K3D0123106', 'D1-PT-118-K3D0123106', 'D1-LSH-309-K3D0123106', 'D1-PSHH-966B-K3D0123106', 'D1-PM-501B-K3D0133106', 'D1-FT-308-K3D0123106', 'D1-PT-418-K3D0123106', 'D1-XV-501-K3D0123106', 'D1-E-424-K3D0123129', 'D1-P-305-K3D0149106', 'D1-P-301B-K3D0149106', 'D1-PT-120-K3D0123106', 'D-OHC-20T-1-X-401D0102106', 'D1-PT-514-K3D0123106', 'D2-R-201-K3D0133129', 'D1-FV-203-K3D0123106', 'D1-PM-303-K3D0133106', 'D2-R-201-K3D0123129', 'D1-LT-202-K3D0123106', 'D1-PM-304B-K3D0133106', 'D1-PT-403-K3D0123106', 'D1-P-302B-K3D0149106', 'D1-TS-601-K3D0149106', 'D1-K-405-P1B-K3D0149106', 'D1-TSHH-428-K3D0123106', 'D1-P-204A-K3D0149106', 'D1-PSL-905-K3D0123106', 'D1-LV-309A-K3D0123106', 'D1-PSL-514-K3D0123106', 'D1-FT-103-K3D0123106', 'D1-LV-507-K3D0123106', 'D1-K-403-P1A-K3D0149106', 'D1-K-403-P2A-K3D0149106', 'D1-FSLL-110-K3D0123106', 'D1-PCV-505-K3D0123106', 'D1-FV-103-K3D0123106', 'D1-PT-408-K3D0123106', 'D1-P-502A-K3D0149106', 'D1-FT-405-K3D0123106', 'D1-LSHH-421-K3D0123106', 'D1-TSL-861-K3D0123106', 'D1-LT-204-K3D0123106', 'D1-FT-602-K3D0123106', 'D1-K-201-K3D0149106', 'D1-PT-905-K3D0123106', 'D1-LSLL-821-K3D0123106', 'D1-FV-308-K3D0123106', 'D1-PSH-902-K3D0123106', 'D1-LV-503A-K3D0123106', 'D1-XV-201-K3D0123106', 'D1-PM-502A-K3D0133106', 'D1-LSH-401-K3D0123106', 'D1-PSHH-964C-K3D0123106', 'D1-TV-130-K3D0123106', 'D1-FV-406-K3D0123106', 'D1-TSL-208-K3D0123106', 'D1-PSH-120-K3D0123106', 'D1-PM-305-K3D0133106', 'D1-LSL-304-K3D0123106', 'D1-TSH-129-K3D0123106', 'D1-FT-403-K3D0123106', 'D1-LSHH-205-K3D0123106', 'D1-FSL-304-K3D0123106', 'D1-LT-308-K3D0123106', 'D1-FT-204-K3D0123106', 'D1-PDSHH-410B-K3D0123106', 'D1-LV-203-K3D0123106', 'D1-LSHH-419-K3D0123106', 'D1-TSH-219-K3D0123106', 'D1-LV-503B-K3D0123106', 'D1-FT-102-K3D0123106', 'D1-LSL-309-K3D0123106', 'D1-PSL-105-K3D0123106', 'D1-FSLL-102-K3D0123106', 'D1-TSH-321-K3D0123106', 'D1-TSHH-402-K3D0123106', 'D1-XV-102-K3D0123106', 'D1-FSLL-304-K3D0123106', 'D1-E-423-K3D0123129', 'D1-FV-503-K3D0123106', 'D1-FV-602-K3D0123106', 'D1-PSL-106-K3D0123106', 'D1-PSHH-866B-K3D0123106', 'D1-LSH-304-K3D0123106', 'D1-LV-401-K3D0123106', 'D1-PSLL-119-K3D0123106', 'D1-TV-308-K3D0123106', 'D1-LSL-503-K3D0123106', 'D1-TSHH-311-K3D0123106', 'D1-PM-301B-K3D0133106', 'D2-E-201-K3D0123129', 'D1-LSH-413-K3D0123106', 'D1-HV-201-K3D0123106', 'D1-PM-602B-K3D0133106', 'D1-PT-409-K3D0123106', 'D1-LSLL-105-K3D0123106', 'D1-PM-503A-K3D0133106', 'D1-P-601B-K3D0149106', 'D1-TV-208-K3D0123106', 'D1-PM-203A-K3D0133106', 'D1-GM-204-K3D0133106', 'D1-PT-106-K3D0123106', 'D1-LV-304B-K3D0123106', 'D1-PM-302A-K3D0133106', 'D1-K-403-P2B-K3D0149106', 'D1-LT-416-K3D0123106', 'D1-PV-101-K3D0123106', 'D1-PV-106-K3D0123106', 'D1-FRSLL-112-K3D0123106', 'D1-LT-511-K3D0123106', 'D1-P-302A-K3D0149106', 'D1-PSH-905-K3D0123106', 'D1-FV-107-K3D0123106', 'D1-TX-301-K3D0149106', 'D1-LSH-817A-K3D0123106', 'D1-LV-309B-K3D0123106', 'D1-TSL-961-K3D0123106', 'D1-LT-313-K3D0123106', 'D1-TT-308-K3D0123106', 'D1-LT-817-K3D0123106', 'D1-LSL-211-K3D0123106', 'D1-PSH-904-K3D0123106', 'D1-TV-422-K3D0123106', 'D1-P-201A-K3D0149106', 'D2-E-201-K3D0149129', 'D1-LSH-312-K3D0123106', 'D1-LSL-301-K3D0123106', 'D1-PV-504-K3D0123106', 'D2-R-201-K3D0149129', 'D1-TSHH-320-K3D0123106', 'D1-PM-201B-K3D0133106', 'D1-P-601A-K3D0149106', 'D1-PV-403-K3D0123106', 'D1-PSL-120-K3D0123106', 'D1-FT-404-K3D0123106', 'D1-P-203A-K3D0149106', 'D1-LT-304-K3D0123106', 'D1-PSL-104-K3D0123106', 'D1-LSH-306-K3D0123106', 'D1-LSH-202-K3D0123106', 'D1-LSH-412-K3D0123106', 'D1-LV-304A-K3D0123106', 'D1-LT-203-K3D0123106', 'D1-PSH-104-K3D0123106', 'D1-AT-101-K3D0123106', 'D1-P-203B-K3D0149106', 'D1-LSH-831A-K3D0123106', 'D1-PSHH-108-K3D0123106', 'D1-HV-309-K3D0123106', 'D1-FT-108-K3D0123106', 'D1-PM-602A-K3D0133106', 'D1-PM-204B-K3D0133106', 'D1-LV-202-K3D0123106', 'D1-LV-509-K3D0123106', 'D1-TSH-405-K3D0123106', 'D1-P-201B-K3D0149106', 'D1-TSHH-406-K3D0123106', 'D1-FT-112-1-K3D0123106', 'D1-LSH-503-K3D0123106', 'D1-K-201-P1-K3D0149106', 'D1-XV-103-K3D0123106', 'D1-FT-401-K3D0123106', 'D1-TSHH-319-K3D0123106', 'D1-LSL-210-K3D0123106', 'D1-PV-203-K3D0123106', 'D1-G-203-K3D0149106', 'D1-PSLL-107-K3D0123106', 'D1-PV-105-K3D0123106', 'D1-PSL-811-K3D0123106', 'D1-FV-108-K3D0123106', 'D1-LSL-203-K3D0123106', 'D1-LSL-306-K3D0123106', 'D1-KM-201-K3D0133106', 'D1-P-503A-K3D0149106', 'D1-XV-101-K3D0123106', 'D1-TV-106-K3D0123106', 'D1-P-201A-P1A-K3D0149106', 'D1-TT-132-K3D0123106', 'D0179106', 'D1-LT-401-K3D0123106', 'D1-LV-201-K3D0123106', 'D1-PT-411-K3D0123106', 'D1-PDSH-401-K3D0123106', 'D1-PSL-809A-K3D0123106', 'D1-PV-118-K3D0123106', 'D1-K-402-K3D0133115', 'D1-P-304B-K3D0149106', 'D1-PDSLL-502-K3D0123106', 'D1-TSHH-412-K3D0123106', 'D1-PT-116-K3D0123106', 'D1-FV-202-K3D0123106', 'D1-PSHH-864C-K3D0123106', 'D1-TSH-425-K3D0123106', 'D1-LSH-308-K3D0123106', 'D1-LSL-311-K3D0123106', 'D1-HV-501-K3D0123106', 'D1-LSLL-817-K3D0123106', 'D1-LSHH-408-K3D0123106', 'D1-LSL-821B-K3D0123106', 'D1-LV-817B-K3D0123106', 'D1-TSL-130-K3D0123106', 'D1-K-403-P1B-K3D0149106', 'D1-TSH-427-K3D0123106', 'D1-E-422-K3D0123129', 'D1-TV-421-K3D0123106', 'D1-PV-207-K3D0123106', 'D0123129', 'D1-G-301-K3D0149106', 'D1-PT-902-K3D0123106', 'D1-LSH-204-K3D0123106', 'D1-P-204B-K3D0149106', 'D1-TT-326-K3D0123106', 'D1-FV-401-K3D0123106', 'D1-PM-204A-K3D0133106', 'D1-P-304A-K3D0149106', 'D1-FRSLL-119-K3D0123106', 'D1-PT-101-K3D0123106', 'D1-LSHH-502-K3D0123106', 'D1-PM-302B-K3D0133106', 'D1-FT-407-K3D0123106', 'D1-LV-817A-K3D0123106', 'D1-LSH-411-K3D0123106', 'D1-LSHH-420-K3D0123106', 'D0175115', 'D1-HV-256-K3D0123106', 'D1-P-501B-K3D0149106', 'D1-ZSHH-841-K3D0123106', 'DPIPING3-102-D0107D0123129', 'D1-FV-402-K3D0123106', 'D1-XV-401-K3D0123106', 'D1-PDT-502-K3D0123106', 'D1-LSH-410-K3D0123106', 'D1-LT-408-K3D0123106', 'DPIPING3-101-D0108D0123129', 'D1-FV-407-K3D0123106', 'D1-TSHH-414-K3D0123106', 'D1-XV-402-K3D0123106', 'D1-K-402P1B-K3D0149106', 'D1-PV-116-K3D0123106', 'D1-PM-304A-K3D0133106', 'D1-FV-114-K3D0123106', 'D1-FT-203-K3D0123106', 'D1-LSHH-511-K3D0123106', 'D1-PSL-910A-K3D0123106', 'D1-TSH-403-K3D0123106', 'D1-PCV-961-K3D0123106', 'D1-PCV-861-K3D0123106', 'D1-LT-301-K3D0123106', 'D1-FT-304-K3D0123106', 'D1-P-201A-P1B-K3D0149106', 'D1-PT-402-K3D0123106', 'D1-TSHH-127-K3D0123106', 'D1-PSL-118-K3D0123106', 'D1-TSH-130-K3D0123106', 'D1-LSL-204-K3D0123106', 'D1-FRSL-112-K3D0123106', 'D1-LSHH-422-K3D0123106', 'D1-TV-310-K3D0123106', 'D1-PM-203B-K3D0133106', 'D1-PSL-116-K3D0123106', 'D1-LV-505-K3D0123106', 'D1-LSH-301-K3D0123106', 'D1-P-501A-K3D0149106', 'D1-PSH-118-K3D0123106', 'D1-LT-312-K3D0123106', 'D1-PT-910-K3D0123106', 'D1-LSL-312-K3D0123106', 'D1-PSLL-117-K3D0123106', 'D1-FRSL-119-K3D0123106', 'D1-LSL-817B-K3D0123106', 'D1-FT-102-1-K3D0123106', 'D1-FV-204-K3D0123106', 'D1-LV-506-K3D0123106', 'D1-TSH-117-K3D0123106', 'D1-FV-201-K3D0123106', 'D1-FSL-303-K3D0123106', 'D1-LSHH-417-K3D0123106', 'D1-LSLL-313-K3D0123106', 'D1-LSHH-414-K3D0123106', 'D17-H-201-K3D0123129', 'D1-PT-203-K3D0123106', 'D1-LT-831B-K3D0123106'])
critical = {'D1-E-501-K3': 96.5, 'D1-E-502-K3': 81.5, 'D1-R-102-K3': 80.5, 'D1-E-108-K3': 78.5, 'D1-E-109-K3': 78.5, 'D1-R-201-K3': 78.5, 'D1-R-202-K3': 78.5, 'D1-C-301-K3': 78.5, 'D1-C-302-K3': 78.5, 'D1-E-201-K3': 77.5, 'D1-E-202-K3': 77.5, 'D1-E-203-K3': 77.5, 'D1-E-209-K3': 77.5, 'D1-X-101-K3': 73.5, 'D1-S-102-K3': 71.5, 'D1-E-206-K3': 68.5, 'D1-E-301-K3': 68.5, 'D1-E-302-K3': 68.5, 'D1-E-303-K3': 68.5, 'D1-R-501-K3': 61.5, 'D1-K-405-K3': 61.5, 'D1-H-101-K3': 58.5, 'D1-E-306A-K3': 58.5, 'D1-E-306B-K3': 58.5, 'D1-E-307-K3': 58.5, 'D1-R-301-K3': 58.5, 'D1-K-403-K3': 58.5, 'D1-K-403HP-K3': 58.5, 'D1-K-403LP-K3': 58.5, 'D1-K-404-K3': 58.5, 'D1-E-509-K3': 58.5, 'D1-E-510A-K3': 58.5, 'D1-E-510B-K3': 58.5, 'D1-E-511-K3': 58.5, 'D1-E-513-K3': 58.5, 'D1-K-405HP-K3': 58.5, 'D1-K-405LP-K3': 58.5, 'D1-E-431-K3': 57.5, 'D1-E-432-K3': 57.5, 'D1-E-433-K3': 57.5, 'D1-E-434-K3': 57.5, 'D1-E-441-K3': 57.5, 'D1-E-503-K3': 57.5, 'D1-E-504-K3': 57.5, 'D1-E-505-K3': 57.5, 'D1-E-506-K3': 57.5, 'D1-E-507-K3': 57.5, 'D1-E-508-K3': 57.5, 'D1-E-512-K3': 57.5, 'D1-E-514-K3': 57.5, 'D1-S-431-K3': 57.5, 'D1-S-432-K3': 57.5, 'D1-S-433-K3': 57.5, 'D1-S-434-K3': 57.5, 'D1-S-501-K3': 57.5, 'D1-K-402P1A-K3': 48.5, 'D1-K-402-K3': 46.5, 'D1-K-402F1B-K3': 46.5, 'D1-TS-402-K3': 46.5, 'D1-E-104A-K3': 42.5, 'D1-E-104B-K3': 42.5, 'D1-E-402-K3': 39.5, 'D1-E-101-K3': 38.5, 'D1-E-102A-K3': 38.5, 'D1-E-102B-K3': 38.5, 'D1-E-103-K3': 38.5, 'D1-E-105-K3': 38.5, 'D1-E-106-K3': 38.5, 'D1-E-107-K3': 38.5, 'D1-V-101-K3': 38.5, 'D1-E-421-K3': 37.5, 'D1-E-422-K3': 37.5, 'D1-E-423-K3': 37.5, 'D1-E-424-K3': 37.5, 'D1-S-421-K3': 37.5, 'D1-S-422-K3': 37.5, 'D1-S-423-K3': 37.5, 'D1-S-424-K3': 37.5, 'D1-K-402-E3-K3': 37.5, 'D1-P-301A-K3': 36.5, 'D1-P-301B-K3': 36.5, 'D1-S-101-K3': 36.5, 'D1-TS-403-K3': 36.5, 'D1-V-501-K3': 36.5, 'D1-V-502-K3': 36.5, 'D1-E-304-K3': 33.5, 'D1-E-305-K3': 33.5, 'D1-PM-301A-K3': 33.5, 'D1-PM-301B-K3': 33.5, 'D1-K-403-P1B-K3': 33.5, 'D1-K-403-PM1-K3': 33.5, 'D1-K-403PM1B-K3': 33.5, 'D1-K-403-P2B-K3': 33.5, 'D1-K-403-PM2-K3': 33.5, 'D1-K-403E1-K3': 33.5, 'D1-TS-405-K3': 33.5, 'D1-E-204-K3': 32.5, 'D1-K-403-T4-K3': 32.5, 'D1-K-403-T5-K3': 32.5, 'D1-K-403-T6A-K3': 32.5, 'D1-K-403-T6B-K3': 32.5, 'D1-K-403-T7A-K3': 32.5, 'D1-K-403-T7B-K3': 32.5, 'D1-K-405-P1B-K3': 32.5, 'D1-K-405-PM1-K3': 32.5, 'D1-E-405-K3': 32.5, 'D1-K-405-E2-AFTCON': 32.5, 'D1-K-405-E2-INTCON': 32.5, 'D1-K-405-P2-K3': 32.5, 'D1-K-405-PM2-K3': 32.5, 'D1-P-201B-K3': 31.5, 'D2-R-201-K3': 80.5, 'D2-E-202-K3': 68.5, 'D2-E-201-K3': 68.5, 'D2-E-203-K3': 68.5, 'D2-J-201-K3': 62.5, 'D2-E-104A-K3': 48.5, 'D2-E-104B-K3': 48.5, 'D2-TS-102-K3': 46.5, 'D2-K-102-K3': 45.5, 'D2-GEAR-K3': 45.5, 'D2-E-102-K3': 43.5, 'D2-E-107-K3': 43.5, 'D2-E-121-K3': 43.5, 'D2-E-122-K3': 43.5, 'D2-E-123-K3': 43.5, 'D2-K-102-E2-K3': 43.5, 'D2-E-302-K3': 43.5, 'D2-E-303-K3': 43.5, 'D2-E-308-K3': 43.5, 'D2-R-101-K3': 42.5, 'D2-S-101-K3': 42.5, 'D2-V-101-K3': 37.5, 'D2-V-102-K3': 37.5, 'D2-V-103-K3': 37.5, 'D2-E-802-K3': 36.5, 'D2-E-803A-K3': 36.5, 'D2-E-803B-K3': 36.5, 'D2-E-804-K3': 36.5, 'D2-C-801-K3': 35.5, 'D2-C-802-K3': 35.5, 'D2-C-803-K3': 35.5, 'D14-J-001A/AR-K3': 42.5, 'D14-J-001B/BR-K3': 42.5, 'D14-J-002A/AR-K3': 42.5, 'D14-J-002B/BR-K3': 42.5, 'D14-P-001A-K3': 42.5, 'D14-PM-001A-K3': 42.5, 'D14-P-001B-K3': 42.5, 'D14-PM-001B-K3': 42.5, 'D14-P-002A-K3': 42.5, 'D14-PM-002A-K3': 42.5, 'D14-P-002B-K3': 42.5, 'D14-PM-002B-K3': 42.5, 'D14-P-003A-K3': 42.5, 'D14-PM-003A-K3': 42.5, 'D14-P-003B-K3': 42.5, 'D14-PM-003B-K3': 42.5, 'D14-P-004A-K3': 42.5, 'D14-PM-004A-K3': 42.5, 'D14-P-004B-K3': 42.5, 'D14-PM-004B-K3': 42.5, 'D17-ECONOMISER-K3': 34.5, 'D17-EVAPORATOR-K3': 34.5, 'D17-E-101-K3': 33.5, 'D17-BURNERS-K3': 33.5, 'D17-E-201-K3': 33.5, 'D1-P-502A-K3': 27.5, 'D1-P-502B-K3': 27.5, 'D1-P-501A-K3': 27.5, 'D1-P-501B-K3': 27.5, 'D1-TX-301-K3': 26.5, 'D1-KM-101A-K3': 25.0, 'D1-KM-101B-K3': 25.0, 'D1-K-101A-K3': 21.5, 'D1-K-101B-K3': 21.5, 'D2-P-102A-K3': 28.5, 'D2-P-102B-K3': 28.5, 'D2-E-101-K3': 27.5, 'D2-P-201A-K3': 27.5, 'D2-P-201B-K3': 27.5, 'D2-G-602A-K3': 27.5, 'D2-G-602B-K3': 27.5, 'D2-P-301A-K3': 25.5, 'D2-P-301B-K3': 25.5, 'D2-B-604A-K3': 23.0, 'D2-B-604B-K3': 23.0, 'D2-E-801-K3': 22.5, 'D2-C-305-K3': 21.5, 'D2-J-604-K3': 20.5, 'D2-BM-604A-K3': 20.0, 'D2-BM-604B-K3': 20.0, 'D2-C-303-K3': 19.0, 'D2-X-602-K3': 17.5, 'D15-V-103-K3': 28.0, 'D15-V-202-K3': 28.0, 'D15-V-203-K3': 28.0, 'D15-V-204-K3': 28.0, 'D15-V-205-K3': 28.0, 'D15-V-101A-K3': 26.5, 'D15-V-101B-K3': 26.5, 'D14-E-001A-K3': 22.5, 'D14-E-001B-K3': 22.5}
noncritical = []
pmtas = []
nonrc = []

main()
