import openpyxl
import load
c1h=0
c2h=0
c3l=0
c4l=0
candel=False
candelnext=False
path = load.path
wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active
for i in range(2,sheet_obj.max_row-5):
    if(float(sheet_obj["B"+str(i)].value)<float(sheet_obj["E"+str(i)].value)):
        candel=True #green
    else :
        candel=False   #red
    if(float(sheet_obj["B"+str(i+1)].value)>float(sheet_obj["E"+str(i+1)].value)):
        candelnext=True#red
    else:
        candelnext=False#green
    if(candel):
        if(candelnext):
            c1l=float(sheet_obj["D"+str(i)].value)
            c2l=float(sheet_obj["D"+str(i+1)].value)
            c3h=float(sheet_obj["C"+str(i+2)].value)
            c4h=float(sheet_obj["C"+str(i+3)].value)
            if(c1l > c3h):
                sheet_obj["J"+str(i)].value=1
            elif(c2l > c4h):
                sheet_obj["J"+str(i+1)].value=1
    
wb_obj.save(path)
            
        
            
    
