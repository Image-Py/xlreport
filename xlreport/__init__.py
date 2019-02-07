from xlreport import *
from openpyxl import load_workbook as load
import numpy as np

def test():
    wb = load('Personal Information.xlsx')
    info, keys = parse(wb)
    
    img = np.random.randint(0, 255, (100,100), dtype=np.uint8)
    data = {'Name':'YX Dragon',
            'Photo':img,
            'Sex':'Male',
            'Age':'30',
            'Height':170,
            'Weight':75.0,
            'Like_Sport':True}
    
    fill_value(wb, info, data)
    repair(wb)
    
    from tkinter import filedialog
    import os.path as osp
    direc = filedialog.askdirectory()
    wb.save(osp.join(direc, 'person.xlsx'))
