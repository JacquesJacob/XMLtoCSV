from bs4 import BeautifulSoup
import csv
import os
from os import popen
from tkinter import *
import tkinter as tk
from tkinter import ttk
from PIL import ImageTk, Image
from tkinter.filedialog import askopenfile
from tkinter import filedialog
from tkinter import messagebox
from tkinter.ttk import Progressbar
import numpy as np
import pandas as pd


def save_csv():
    global save_csv_file
    save_csv_file = filedialog.asksaveasfilename(initialdir="/", title="Select file", filetypes=(("csv files", "*.csv"),
                                                                                                 ("all files", "*.*")))
    CSV_field.delete(0, END)
    CSV_field.insert(0, save_csv_file)


def selectxml():
    global selectxmlfile
    selectxmlfile = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(("xml files", "*.xml"),
                                                                                               ("all files", "*.*")))
    XML_field.delete(0, END)
    XML_field.insert(0, selectxmlfile)


def convert_data_alert():
    try:
        f = open(save_csv_file, 'w', newline='')
        csvwriter = csv.writer(f)
        bar()
        entity_record = "Match Number"
        search_name = "Batch Run Id"
        search_entity = "Full Name Search"
        search_origin = "Entity Type"
        search_listname = "Predefined Search Name"
        alertstatus = "Origin"
        match_resultid = "Alert"
        match_number = "Best Name Match"
        entity_name_match = "Best Name Score"
        bestname_score = "Reason"
        auto_false_positive = "Auto False Positive"
        head = [entity_record, search_name, search_entity, search_origin, search_listname, alertstatus, match_resultid,
                match_number, entity_name_match, bestname_score, auto_false_positive]
        csvwriter.writerow(head)

        infile = open(selectxmlfile, "r", encoding="utf-8")
        contents = infile.read()
        soup = BeautifulSoup(contents, 'xml')
        Indice = soup.find_all('ResultRecord')
        numero = 0
        for Lista in Indice:
            numero = numero + 1
            batchrun = Lista.find('BatchRunId').text
            namefull = Lista.find('Name').find('Full').text
            entitytype = Lista.find('InputEntity').find('EntityType').text
            pds = Lista.find('PredefinedSearchName').text
            origin = Lista.find('Origin').text
            alertstate = Lista.find('AlertState').text

            links = Lista.find_all('Match')
            for link in links:
                entity_name_match_result = link.find("BestNameMatch").text
                bestname_score_result = link.find("BestNameScore").text
                reason_match_result = link.find("Reason").text
                false_positive = link.find("IsAutomaticFalsePositive").text

                row = [numero, batchrun, namefull, entitytype, pds, origin, alertstate, entity_name_match_result,
                       bestname_score_result, reason_match_result, false_positive]
                csvwriter.writerow(row)

        messagebox.showinfo("Info", "Success!")

    except:
        messagebox.showerror("Error", "Check the file and try again.")


def convert_xlsx():
    try:
        file_format = save_csv_file.split(".csv")[0]
        file_xls = file_format + ".xlsx"
        read_file = pd.read_csv(r'%s' % save_csv_file, encoding="ANSI")
        read_file.to_excel(r'%s' % file_xls, index=None, header=True)
    except:
        messagebox.showerror("Error", "Cannot convert file to XLSX. Check the files and try again.")


def convert_data_batch():
    try:
        f = open(save_csv_file, 'w', newline='')
        csvwriter = csv.writer(f)
        bar()
        entity_record = "Entity Record"
        search_name = "Search Name"
        search_entity = "Search Entity"
        search_origin = "Search Origin"
        search_listname = "Search List Name"
        alertstatus = "Alert Status"
        match_resultid = "Match ResultID"
        match_number = "Match Number"
        entity_name_match = "Entity Name Match"
        bestname_score = "Best Name Score Match"
        reason_match = "Reason Match"
        head = [entity_record, search_name, search_entity, search_origin, search_listname, alertstatus, match_resultid,
                match_number, entity_name_match, bestname_score, reason_match]
        csvwriter.writerow(head)

        infile = open(selectxmlfile, "r", encoding="utf-8")
        contents = infile.read()
        soup = BeautifulSoup(contents, 'xml')
        Indice = soup.find_all('Entity')

        for Lista in Indice:
            response = Lista.get('Record')

            if response is None:
                continue
            ent_record_result = response
            Infos = soup.find("Entity", {"Record": "%s" % response})
            search_name_result = Infos.find("FullName").text
            search_entity_result = Infos.find("EntityType").text
            search_origin_result = Infos.find("Origin").text

            if Infos.find("FileName") is None:
                search_listname_result = "NA"
            else:
                search_listname_result = Infos.find("FileName").text

            if Infos.find("Status") is None:
                alert_status_result = "NA"
            else:
                alert_status_result = Infos.find("Status").text

            if Infos.find("Match") is None:
                match_resultid_result = "NA"

                row = [ent_record_result, search_name_result, search_entity_result, search_origin_result,
                       search_listname_result, alert_status_result, match_resultid_result]
                csvwriter.writerow(row)

            else:
                match_resultid_result = Infos.attrs['ResultID']
                links = Infos.find_all('Match')
                for link in links:
                    match_number_result = link.attrs['ID']
                    entity_name_match_result = link.find("EntityName").text
                    bestname_score_result = link.find("BestNameScore").text
                    reason_match_result = link.find("Reason").text

                    row = [ent_record_result, search_name_result, search_entity_result, search_origin_result,
                           search_listname_result, alert_status_result, match_resultid_result, match_number_result,
                           entity_name_match_result, bestname_score_result, reason_match_result]
                    csvwriter.writerow(row)

        messagebox.showinfo("Info", "Success!")

    except:
        messagebox.showerror("Error", "Check the file and try again.")


def chose_menu():
    xls_option = xls.get()
    selected = selection.get()
    if selected == "Manual Batch XML" or selected == "Alert Report XML":
        convert_data_alert()
        if xls_option == 1:
            convert_xlsx()
    else:
        convert_data_batch()
        if xls_option == 1:
            convert_xlsx()


def clear_fields():
    CSV_field.delete(0, END)
    XML_field.delete(0, END)
    global selectxmlfile
    global save_csv_file
    selectxmlfile = []
    save_csv_file = []


def about_info():
    messagebox.showinfo("Info", "Developed by: Jacques Jacob\n"
                                "Contact: jacques.jacob@gmail.com\n"
                                "Release Date: 06/01/2020\n"
                                "Version 1.0")


root = tk.Tk()
root.title("XML to CSV - Custom Conversion File")


def bar():
    progressing = tk.Tk()
    progressing.overrideredirect(True)

    winWidth = progressing.winfo_reqwidth()
    winHeight = progressing.winfo_reqheight()
    positRight = int(progressing.winfo_screenwidth() / 2 - winWidth / 2)
    positDown = int(progressing.winfo_screenheight() / 2 - winHeight / 2)
    progressing.geometry("+{}+{}".format(positRight, positDown))

    progress = Progressbar(progressing, orient=HORIZONTAL,
                           length=100, mode='determinate')
    progress.pack(pady=10)
    import time
    progress['value'] = 20
    root.update_idletasks()
    time.sleep(0.1)

    progressing.update()

    progress['value'] = 40
    root.update_idletasks()
    time.sleep(0.1)

    progressing.update()

    progress['value'] = 50
    root.update_idletasks()
    time.sleep(0.1)

    progressing.update()

    progress['value'] = 60
    root.update_idletasks()
    time.sleep(0.1)

    progressing.update()

    progress['value'] = 80
    root.update_idletasks()
    time.sleep(0.1)
    progress['value'] = 100

    progressing.update()

    progressing.destroy()


HEIGHT = 100
WIDTH = 415

labelframe = LabelFrame(root, text="Select file to convert:")
labelframe.pack(fill="both", expand="yes")

XML_text = Label(root, text="XML file:")
XML_text.place(relx=0.17, rely=0.22)

XML_btn = Button(root, text=' XML File ', command=lambda: selectxml())
XML_btn.place(relx=0.71, rely=0.22)

XML_field = tk.Entry(root)
XML_field.place(relwidth=0.395, relheight=0.18, relx=0.3, rely=0.22)

CSV_text = Label(root, text="Save as:")
CSV_text.place(relx=0.183, rely=0.45)

CSV_btn = Button(root, text='  CSV File ', command=lambda: save_csv())
CSV_btn.place(relx=0.71, rely=0.45)

CSV_field = tk.Entry(root)
CSV_field.place(relwidth=0.395, relheight=0.18, relx=0.3, rely=0.45)

CONVERT_btn = Button(root, text=' Convert to CSV ', command=lambda: chose_menu())
CONVERT_btn.place(relx=0.71, rely=0.70)

button_clear = tk.Button(root, text='Clear', command=lambda: clear_fields())
button_clear.place(relx=0.565, rely=0.70)

selection = StringVar(root, value="Manual Batch XML")
selection.set("Manual Batch XML")
option_menu = tk.OptionMenu(root, selection, "Manual Batch XML", "Auto Batch XML", "Alert Report XML")
option_menu.place(relx=0.20, rely=0.68)

xls = IntVar()
checkbut = Checkbutton(root, text="XLS", onvalue=1, offvalue=0, variable=xls)
checkbut.place(relx=0.87, rely=0.45)

lexisnexisimage = PhotoImage(file=r"image.PNG")
about_button = tk.Button(root, text="About ", image=lexisnexisimage, compound=LEFT, command=lambda: about_info())
about_button.place(relx=0.01, rely=0.69)

windowWidth = root.winfo_reqwidth()
windowHeight = root.winfo_reqheight()
positionRight = int(root.winfo_screenwidth()/2 - windowWidth/2)
positionDown = int(root.winfo_screenheight()/2 - windowHeight/2)
root.geometry("+{}+{}".format(positionRight, positionDown))

canvas = tk.Canvas(labelframe, height=HEIGHT, width=WIDTH)
canvas.pack()
root.resizable(False, False)

root.mainloop()
