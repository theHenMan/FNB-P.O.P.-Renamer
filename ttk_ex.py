"""
Open a file dialog window in tkinter using the filedialog method.
Tkinter has a prebuilt dialog window to access files.
This example is designed to show how you might use a file dialog
askopenfilename and use it in a program.
"""

from tkinter import ttk
import PyPDF2
import os
import sys
import re
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

root = tk.Tk()

ft = ttk.Frame()
ft.pack(expand=True, fill=tk.BOTH, side=tk.BOTTOM)

pb1 = ttk.Progressbar(ft, orient='horizontal', mode='indeterminate')
pb1.pack(expand=True, fill=tk.BOTH, side=tk.BOTTOM, padx=65, pady=5)

initial_dir = os.getcwd


def user_cancelled():
    result = messagebox.askyesno('Cancel?', 'Do you want to cancel?')
    if result is True:
        messagebox.showinfo('Exiting', 'The program will now exit')
        sys.exit()
    return True


def no_pdf_files():
    messagebox.showwarning('No PDF', 'No PDF files found in this directory.')
#     sys.exit()


# This is where we lauch the file manager bar.
def OpenFile():
    global pb1
    counter = 0
    i = 1  # Increment i when there are duplicate names; used later

    open_dir = ''
    while not open_dir != '':
        open_dir = filedialog.askdirectory(initialdir=initial_dir,
                                           title="Choose a folder.")
        if open_dir == '':
            if user_cancelled() is False:
                break

    open_dir = open_dir + '/'
    pdf_reader = ''

    for pdf in os.listdir(open_dir):
        pb1.start(50)
        if pdf == 'Thumbs.db':
            continue
        full_path = open_dir + pdf
        with open(full_path, 'rb') as pdf_fileobj:
            try:
                pdf_reader = PyPDF2.PdfFileReader(pdf_fileobj)
                counter += 1
            except Exception as e:
                if counter == 0:
                    no_pdf_files()
                else:
                    sys.exit()

            pdf_reader.numPages

            page_obj = pdf_reader.getPage(0)
            extracted = page_obj.extractText()

            # Extract index numbers from pdfs
            # to search and use for new_filename
            # The variable names are those that appear in the
            # extracted text
            date_actioned = re.compile('Date Actioned:')
            date_actioned = date_actioned.search(extracted)
            time_actioned = re.compile('Time Actioned:')
            time_actioned = time_actioned.search(extracted)
            date_corrected = extracted[date_actioned.start()
                                       + 14:time_actioned.start()].replace(':',
                                                                           '-')

            trace_id = re.compile('Trace ID:')
            trace_id = trace_id.search(extracted)

            pmt_name = re.compile('Name: ')
            pmt_name = pmt_name.search(extracted)
            bank = re.compile('Bank:')
            bank = bank.search(extracted)

            ref = re.compile('Reference: ')
            ref = ref.search(extracted)
            channel = re.compile('Channel: ')
            channel = channel.search(extracted)
            ref = extracted[ref.start() + 11:channel.start()]

            amount = re.compile('Amount: ')
            amount = amount.search(extracted)
            payee = re.compile('Payee')
            payee = payee.search(extracted)
            amount = extracted[amount.start() + 8:payee.start()]

            cleaned_up_name = extracted[pmt_name.start() + 6:bank.start()]
            cleaned_up_name = cleaned_up_name.replace('/', ' ')

            new_filename = open_dir + cleaned_up_name
            new_filename = new_filename + date_corrected + ' '
            new_filename = new_filename + ref + ' '
            new_filename = new_filename + 'R' + amount

            pdf_fileobj.close()

            try:
                os.rename(full_path, new_filename + '.pdf')
            except FileNotFoundError:
                pass
            except FileExistsError:
                try:
                    os.rename(full_path, new_filename
                              + ' (' + str(i) + ')' + '.pdf')
                    i += 1
                except FileExistsError:
                    continue
    pb1.stop()


Title = root.title("Rename FNB POPs")
label_text = '''PDF files inside the selected folder will be renamed as follows:
\nName
Date
Reference number
Amount'''
label = ttk.Label(root, text=label_text, foreground="black",
                  font=("Helvetica", 11))
label.pack(padx=30, pady=20)
select_button = ttk.Button(root, text='Select folder', command=OpenFile)
select_button.pack(padx=10, pady=10)

# Menu Bar
menu = tk.Menu(root)
root.config(menu=menu)

file = tk.Menu(menu, tearoff=0)
file.add_command(label='Open', command=OpenFile)
file.add_command(label='Exit', command=lambda: sys.exit())

menu.add_cascade(label='File', menu=file)

root.mainloop()
