import PyPDF2
import os
import re
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox


# Create Tk window to ask for directory
root = tk.Tk()

icon_dir = os.path.join(os.getcwd())
try:
    root.iconbitmap(icon_dir + '\\icon\\pdf_icon.ico')
except Exception as e:
    pass

root.withdraw()


def user_cancelled():
    result = messagebox.askyesno('Cancel?', 'Do you want to cancel?')
    if result is True:
        print('Exiting...')
        exit()
    return True


def no_pdf_files():
    messagebox.showwarning('No PDF', 'No PDF files found in this directory.')
    exit()


i = 1  # Increment i when there are duplicate names; used later

initial_dir = os.getcwd
title = 'Please select a directory'

open_dir = ''

# While the selected directory is empty, just keep looping
while not open_dir != '':
    open_dir = filedialog.askdirectory(initialdir=initial_dir, title=title)
    if open_dir == '':
        user_cancelled()

open_dir = open_dir + '/'

for pdf in os.listdir(open_dir):
    full_path = open_dir + pdf
    with open(full_path, 'rb') as pdf_fileobj:
        try:
            pdf_reader = PyPDF2.PdfFileReader(pdf_fileobj)
        except Exception as e:
            no_pdf_files()

        pdf_reader.numPages

        page_obj = pdf_reader.getPage(0)
        extracted = page_obj.extractText()

        # Extract index numbers from pdfs
        # to search and use for new_filename
        date_actioned = re.compile('Date Actioned:')
        date_actioned = date_actioned.search(extracted)
        time_actioned = re.compile('Time Actioned:')
        time_actioned = time_actioned.search(extracted)
        time_corrected = extracted[date_actioned.start()
                                   + 14:time_actioned.start()].replace(':', '-')
        trace_id = re.compile('Trace ID:')
        trace_id = trace_id.search(extracted)
        trace = extracted[time_actioned.start()
                          + 14:trace_id.start()].replace(':', '-')
        pmt_name = re.compile('Name: ')
        pmt_name = pmt_name.search(extracted)
        bank = re.compile('Bank:')
        bank = bank.search(extracted)

        new_filename = open_dir + extracted[pmt_name.start() + 6:bank.start()]
        new_filename = new_filename + extracted[date_actioned.start()
                                                + 14:time_actioned.start()]
        new_filename = new_filename + trace

        pdf_fileobj.close()

        try:
            os.rename(full_path, new_filename + '.pdf')
        except FileExistsError:
            try:
                os.rename(full_path, new_filename + ' (' + str(i) + ')' + '.pdf')
                i += 1
            except FileExistsError:
                continue

print('Done')
