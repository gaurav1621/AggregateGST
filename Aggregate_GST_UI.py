# -*- coding: utf-8 -*-
"""
Created on Sun Jun 13 14:34:53 2021

@author: gouna
"""

import PySimpleGUI as sg
import os.path

file_list_column = [
    [
        sg.Text("Choose Source Excel"),
        sg.In(size=(25, 1), enable_events=True, key="-SOURCEEXCEL-"),
        sg.FileBrowse(),
    ],
    [
        sg.Text("Choose Destination Folder"),
        sg.In(size=(25, 1), enable_events=True, key="-OUTPUTFOLDER-"),
        sg.FolderBrowse(),
    ],
     [sg.Button("Generate Excel", key="-GENERATEEXCEL-")],
     # [sg.Text("", key="-RESULTS-")],
]

# ----- Full layout -----
layout = [
    [
        sg.Column(file_list_column),
        # sg.VSeperator(),
        # sg.Column(image_viewer_column),
    ]
]

window = sg.Window("Excel Generator", layout)

while True:
    event, values = window.read()
    if event == "Exit" or event == sg.WIN_CLOSED:
        break
    if event == "-GENERATEEXCEL-":
        source_excel = values["-SOURCEEXCEL-"]
        destination_folder = values["-OUTPUTFOLDER-"]
        print ("source_excel: ", source_excel)
        print ("destination_folder : ", destination_folder )
        sg.popup('Excel generation process done...')
        # window["-RESULTS-"].Update(value="Generation process done...")
        #window.close()
        #break
    