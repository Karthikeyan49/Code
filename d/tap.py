import load 
path = load.path
wb_obj = load.wb_obj
sheet_obj = load.sheet_obj
for i in range(2,sheet_obj.max_row-3):
    if((str(sheet_obj["J"+str(i)].value))==str(1)):
        for k in range(i+3,sheet_obj.max_row-3):
            if(float(sheet_obj["C"+str(i)].value)>=float(sheet_obj["C"+str(k)].value) and float(sheet_obj["D"+str(i)].value)<=float(sheet_obj["C"+str(k)].value)):
                sheet_obj["P"+str(k)].value="tap"+str(i)
                for j in range(k+1,sheet_obj.max_row-3):
                    if((float(sheet_obj["C"+str(i)].value)+1.5)<float(sheet_obj["C"+str(j)].value)):
                        # if(str(sheet_obj["A"+str(i)].value)[11:19]==str("09:15:00")):
                        #     if((float(sheet_obj["D"+str(i)].value)-1.5)>float(sheet_obj["E"+str(j)].value)):
                        #         continue
                        # else:
                            sheet_obj["L"+str(i)].value=1
                            break
                    elif(float(sheet_obj["C"+str(j)].value)<=((float(sheet_obj["D"+str(i)].value)))-((float(sheet_obj["C"+str(i)].value)-(float(sheet_obj["D"+str(i)].value)))*3)):
                        sheet_obj["K"+str(i)].value=1
                        break
                break
wb_obj.save(path)