import sqlite3, pandas as pd
from xlsxwriter.workbook import Workbook
import PySimpleGUI as sg


########################
# Main Logic - Start
#######################

def generate_excel(out_file_path, sour_file_path):
    try:
        output_excel_path = out_file_path+'/output.xlsx'
        source_file_path=sour_file_path
        
        sql_query="SELECT BillNo,Date,Name,GSTNo,TaxValue,CGSTPerc,CGST,SGSTPerc,SGST,SUM(TotalAmount) FROM Table1 group by BillNo,Date,Name,CGSTPerc"
        
        workbook = Workbook(output_excel_path)
        worksheet = workbook.add_worksheet()
        
        conn = sqlite3.connect(':memory:')
        cur = conn.cursor()
        
        excel1 = pd.read_excel(source_file_path, names=["BillNo","Date","Name","GSTNo","TaxValue","CGSTPerc","CGST","SGSTPerc","SGST","TotalAmount"], usecols="A:D,L:P,S", skiprows=[1])
        # print ("excel1: ", excel1)
        excel1.to_sql(name='Table1', con=conn, if_exists='append')
        
        
        mysel = cur.execute(sql_query)
        #names = list(map(lambda x: x[0], cur.description)) #Returns the column names
        # print(names)
        # for row in cur:
            # print(row)
        
        for i, row in enumerate(cur.description):
            worksheet.write(0,i,row[0])
        
        
        for i, row in enumerate(mysel):
            for j, value in enumerate(row):
                worksheet.write(i+1, j, value)
        workbook.close()
        
        cur.close()
        return "Successfully executed"
    # except Exception as m:
    except Exception:
        # return str(m)+" Failed"
        return "Failed"

########################
# Main Logic - End
#######################


########################
# Frontend UI - Start
#######################
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
        result = generate_excel (destination_folder, source_excel)
        sg.popup('Excel generation process done...', result)
        #window.close()
        #break
    

########################
# Frontend UI - End
#######################
