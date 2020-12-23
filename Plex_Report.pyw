#!/usr/bin/env python
# -*- coding: utf-8 -*-

import traceback
from plexapi.myplex import MyPlexAccount
from tkinter import *
from tkinter import messagebox, filedialog
from datetime import date
import openpyxl
import pandas as pd
from styleframe import StyleFrame

def show_error(self, *args):
    err = traceback.format_exception(*args)
    messagebox.showerror('Exception', err)

    usernameLabel = Label(root, text='Enter Username: ', padx=10, width=25, anchor=W)
    passwordLabel = Label(root, text='Enter Password:', padx=10, width=25, anchor=W)
    serverLabel = Label(root, text='Enter Server Name: ', padx=10, width=25, anchor=W)
    libraryLabel = Label(root, text='Enter Library Name: ', padx=10, width=25, anchor=W)

    username = Entry(root, width=30)
    password = Entry(root, show='*', width=30)
    server = Entry(root, width=30)
    library = Entry(root, width=30)

    usernameLabel.grid(row=0, column=0)
    passwordLabel.grid(row=1, column=0)
    serverLabel.grid(row=2, column=0)
    libraryLabel.grid(row=3, column=0)

    username.grid(row=0, column=1)
    password.grid(row=1, column=1)
    server.grid(row=2, column=1)
    library.grid(row=3, column=1)

    Label(root).grid(row=4, column=0, columnspan=2)
    submitButton = Button(root, text='Create Notepad File', command=lambda: notepad_submit(username.get(), password.get(), server.get(), library.get()), padx=10, width=25)
    submitButton.grid(row=5, column=0, columnspan=2)

    submitButton = Button(root, text='Create Excel File', command=lambda: excel_submit(username.get(), password.get(), server.get(), library.get()), padx=10, width=25)
    submitButton.grid(row=6, column=0, columnspan=2)
    Label(root).grid(row=7, column=0, columnspan=2)

    root.update()

def excel_submit(username, password, server, library):
    global show

    output = []
    account = MyPlexAccount(username, password)
    plex = account.resource(server).connect()
    movies = plex.library.section(library)
    for video in movies.search():
        output.append(str(video.title))

    columns = [library]
    export_file_path = filedialog.asksaveasfilename(initialfile=f'{username}_{server}_{library}_{str(date.today())}', defaultextension='.xlsx')
    df = pd.DataFrame(data=output, columns=columns)
    excel_writer = StyleFrame.ExcelWriter(export_file_path)
    sf = StyleFrame(df)
    sf.to_excel(
        excel_writer=excel_writer, 
        best_fit=columns,
        columns_and_rows_to_freeze='A2', 
        row_to_add_filters=0,
    )
    excel_writer.save()

    messagebox.showinfo(title='Finished!', message='Time to party!')

def notepad_submit(username, password, server, library):
    global show

    account = MyPlexAccount(username, password)
    plex = account.resource(server).connect()
    movies = plex.library.section(library)
    export_file_path = filedialog.asksaveasfilename(initialfile=f'{username}_{server}_{library}_{str(date.today())}', defaultextension='.txt')
    with open(export_file_path, 'w') as fh:
        for video in movies.search():
            fh.write('\n')
            string_nonASCII = video.title
            string_encode = string_nonASCII.encode("ascii", "ignore")
            string_decode = string_encode.decode()
            fh.write(string_decode)

    messagebox.showinfo(title='Finished!', message='Time to party!')

if __name__ == '__main__':
    root = Tk()
    root.title('Plex Report')
    Tk.report_callback_exception = show_error


    usernameLabel = Label(root, text='Enter Username: ', padx=10, width=25, anchor=W)
    passwordLabel = Label(root, text='Enter Password:', padx=10, width=25, anchor=W)
    serverLabel = Label(root, text='Enter Server Name: ', padx=10, width=25, anchor=W)
    libraryLabel = Label(root, text='Enter Library Name: ', padx=10, width=25, anchor=W)

    username = Entry(root, width=30)
    password = Entry(root, show='*', width=30)
    server = Entry(root, width=30)
    library = Entry(root, width=30)

    usernameLabel.grid(row=0, column=0)
    passwordLabel.grid(row=1, column=0)
    serverLabel.grid(row=2, column=0)
    libraryLabel.grid(row=3, column=0)

    username.grid(row=0, column=1)
    password.grid(row=1, column=1)
    server.grid(row=2, column=1)
    library.grid(row=3, column=1)

    Label(root).grid(row=4, column=0, columnspan=2)
    submitButton = Button(root, text='Create Notepad File', command=lambda: notepad_submit(username.get(), password.get(), server.get(), library.get()), padx=10, width=25)
    submitButton.grid(row=5, column=0, columnspan=2)

    submitButton = Button(root, text='Create Excel File', command=lambda: excel_submit(username.get(), password.get(), server.get(), library.get()), padx=10, width=25)
    submitButton.grid(row=6, column=0, columnspan=2)
    Label(root).grid(row=7, column=0, columnspan=2)

    root.mainloop()
