# Only for Windows OS
import os
from docxtpl import DocxTemplate
import pandas as pd
import time

from win32com import client
# data_frame = pd.read_csv('')
word_app = client.Dispatch("Word.Application")
data_frame = pd.read_excel('excel-final.xlsx')
data_frame = data_frame.fillna('')

for r_index, row in data_frame.iterrows():
    name = row['SNO'] # used for unique file names
    print('Creating ' +str(r_index+1) +' of ' + str(len(data_frame))) # status of the progress

    tpl = DocxTemplate("template.docx")
    df_to_doct = data_frame.to_dict()
    x = data_frame.to_dict(orient='records')
    context = x
    tpl.render(context[r_index])
    tpl.save('Doc\\'+name+".docx")

    # converting the Docs to PDFs
    time.sleep(1) # for safe PDF creation

    ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
    
    doc = word_app.Documents.Open(ROOT_DIR+'\\Doc\\'+name+ '.docx')
    doc.SaveAs(ROOT_DIR+'\\PDF\\' + name + '.pdf', FileFormat=17)

    print('Completed ' +str(r_index+1) +' of ' + str(len(data_frame))) # status of the progress
word_app.Quit()