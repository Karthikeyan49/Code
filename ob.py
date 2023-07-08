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
        candel=False #green
    else :
        candel=True   #red
    if(float(sheet_obj["B"+str(i+1)].value)>float(sheet_obj["E"+str(i+1)].value)):
        candelnext=False#red
    else:
        candelnext=True#green
    if(candel):
        if(candelnext):
            c1h=float(sheet_obj["C"+str(i)].value)
            c2h=float(sheet_obj["C"+str(i+1)].value)
            c3l=float(sheet_obj["D"+str(i+2)].value)
            c4l=float(sheet_obj["D"+str(i+3)].value)
            if(c1h < c3l):
                sheet_obj["J"+str(i)].value=1
            elif(c2h<c4l):
                sheet_obj["J"+str(i+1)].value=1
            else:
                sheet_obj["J"+str(i)].value=0
        else:
            sheet_obj["J"+str(i)].value=0
    
wb_obj.save(path)
            
        
            
    
