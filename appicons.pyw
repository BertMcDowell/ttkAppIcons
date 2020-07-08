import os
import ast
import configparser
import base64
import zlib
import sys

try:
    from PIL import Image as PImage
except Exception as exception:
    print('Failed to import : '+str(exception))

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import scrolledtext
from tkinter import ttk

window = None

content = None

homePath = os.path.expanduser("~")

appName = "App Icons"
appID = "bertmcdowell."+appName.lower().replace(" ", "")+".01"
appDataPath = os.path.join(homePath, "."+appName.lower().replace(" ", ""))
appConfigPath = os.path.join(appDataPath, "config.ini")
appContentPath = os.path.join(appDataPath, "content.dat")

if "win32" in sys.platform:
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(appID)

def on_load():
    global window
    global browserPath
    global resultsBox

    if os.path.isfile(appConfigPath):
        config = configparser.ConfigParser()
        config.read(appConfigPath)

        if config.has_option("window", "geometry"):
            window.geometry(config.get("window", "geometry"))

        if config.has_option("browser", "terms"):
            try:
                browserPath['values'] = ast.literal_eval(config.get("browser", "terms"))
            except Exception as exception:
                print(exception)
            if len(browserPath['values']) > 0:
                browserPath.current(0)
        if config.has_option("browser", "selected"):
            try:
                browserPath.current(config.getint("browser", "selected"))
            except Exception as exception:
                print(exception)

    if os.path.isfile(appContentPath):
        resultsBox.delete("1.0", tk.END)
        content = open(appContentPath, 'r')
        resultsBox.insert(tk.END, content.read())
        content.close()

def on_save():
    global window
    global browserPath
    global resultsBox

    config = configparser.ConfigParser()
    config.add_section("window")
    config.set("window", "geometry", window.geometry())

    config.add_section("browser")
    config.set("browser", "selected", str(browserPath.current()))
    config.set("browser", "terms", str(browserPath['values']))

    if not os.path.exists(appDataPath):
        os.makedirs(appDataPath)

    configFile = open(appConfigPath, 'w')
    config.write(configFile)
    configFile.close()


def on_close():
    global window

    if messagebox.askokcancel("Quit", "Do you want to quit?"):
        on_save()

        if content != None:
            content.close()

        window.destroy()

def get_path():
    global browserPath
    path = browserPath.get()
    initialdir = ''
    if len(path) > 0:
        if os.path.isfile(path):
            initialdir = os.path.split(path)[0]
        elif os.path.isdir(path):
            initialdir = path
    return initialdir

def set_path(path):
    global browserPath
    if len(path) > 0:
        termsArray = list(browserPath['values'])
        if not (path in termsArray):
            termsArray.insert(0, path)
            if len(termsArray) > 20:
                termsArray.pop()
            browserPath['values'] = termsArray
            browserPath.current(0)
        else:
            browserPath.current(termsArray.index(path))

def do_open_file():
    browserFilePath = filedialog.askopenfilename(initialdir=get_path(), title = "Select file",filetypes = (("png files","*.png"),("bmp files","*.bmp")))
    set_path(browserFilePath)

def do_export_ico():
    exportfile = filedialog.asksaveasfilename(initialdir=get_path(), title="Save file", filetypes=[("ico files","*.ico")], defaultextension='.ico')
    if exportfile:
        try:
            img_file = browserPath.get()
            icon_sizes = [(16,16), (32, 32), (48, 48), (64,64)]
            img = PImage.open(img_file)
            img.save(exportfile, sizes=icon_sizes)
        except Exception as exception:
            messagebox.showerror(title='Error', message='export failed : '+str(exception))

def on_combobox_focusout(event):
    if isinstance(event.widget, ttk.Combobox):
        term = event.widget.get()
        if len(term) > 0:
            termsArray = list(event.widget['values'])
            if not (term in termsArray):
                termsArray.insert(0, term)
                if len(termsArray) > 20:
                    termsArray.pop()
                event.widget['values'] = termsArray

def on_convert_button():
    global browserPath
    global resultsBox

    path = browserPath.get()

    if os.path.isfile(path):
        icon = open(path, 'rb')
        data = icon.read()
        icon.close()

        data = zlib.compress(data, level=9)
        #data = base64.b64encode(data)
        data = base64.b85encode(data)

        resultsBox.delete('1.0', tk.END)
        resultsBox.insert(tk.END, str(data))

        if not os.path.exists(appDataPath):
            os.makedirs(appDataPath)

        content = open(appContentPath, 'w')
        content.write(str(data))
        content.close()

window = tk.Tk()

style = ttk.Style(window)
style.theme_use("clam")

# Contents
browserFrame = tk.Frame(master=window)
browserFrame.pack(fill=tk.X, pady=2, padx=2)

browserLabel = tk.Label(master=browserFrame, text="Path", width=6, anchor=tk.W)
browserPath = ttk.Combobox(master=browserFrame, values=[])
browserPath.bind("<FocusOut>", on_combobox_focusout)
browserButton = tk.Button(master=browserFrame, text="Convert", command=on_convert_button, width=30)

browserLabel.pack(side=tk.LEFT, padx=5, pady=5)
browserButton.pack(side=tk.RIGHT, padx=5, pady=5)
browserPath.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

resultsFrame = tk.Frame(master=window)
resultsFrame.pack(fill=tk.BOTH, expand=True, pady=2, padx=2)

resultsBox = scrolledtext.ScrolledText(master=resultsFrame, wrap=tk.WORD, width=20, height=10)
resultsBox.pack(fill=tk.BOTH, expand=True, pady=2, padx=2)

# Top Menus
menu = tk.Menu(master=window, tearoff=0)
filemenu = tk.Menu(master=menu, tearoff=0)
exportmenu = tk.Menu(master=menu, tearoff=0)

menu.add_cascade(label="File", menu=filemenu)

# export
exportmenu.add_command(label="ico", command=do_export_ico)

# File Menu
filemenu.add_command(label="Open File", command=do_open_file)
filemenu.add_separator()
filemenu.add_cascade(label="Export", menu=exportmenu)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=on_close)

window.config(menu=menu)
window.grid_columnconfigure(0, weight=1)
window.title(appName)
window.geometry('1080x1080')
window.protocol("WM_DELETE_WINDOW", on_close)

on_load()

window.mainloop()