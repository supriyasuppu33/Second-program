# Second-program
Trying another time


import subprocess
import tkinter as tk
import xml.etree.ElementTree as ET
import pandas as pd
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import xml.etree.ElementTree as ET
import pandas as pd


class s4:


    def convert(self,root_f):
        tree = ET.parse(root_f)

        # getting the parent tag of
        # the xml document
        root = tree.getroot()
        A = []
        for ele in root:
            B = {}
            for i in list(ele):
                B.update({i.tag: i.text})
                A.append(B)
        df = pd.DataFrame(A)
        df.drop_duplicates(keep='first', inplace=True)
        df.reset_index(drop=True, inplace=True)
        writer = pd.ExcelWriter("D:/testresult4.xlsx", engine='xlsxwriter')
        df.to_excel(writer, sheet_name='sheet1')
        worksheet = writer.sheets['sheet1']
        worksheet.set_column('B:Z', 30)
        writer.save()
        print("xml converted into Excel successfully")

    def select_files(self):
        filetypes = (
            ('text files', '*.xml'),
            ('All files', '*.*')
        )

        filenames = fd.askopenfilenames(
            title='Open files',
            initialdir='/',
            filetypes=filetypes)
        for filename in filenames:

            self.convert(filename)


if __name__=='__main__':
    p=s4()
    p.select_files()
