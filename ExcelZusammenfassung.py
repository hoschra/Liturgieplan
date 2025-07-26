 #-*- coding: utf-8 -*-
#if it's not working: 
#    del /s /q c:\Users\user\AppData\Local\Temp\gen_py 
#and restart kernel or restart spyder

"""
@author: Volker
- Improvment set automatic over the linne if date is differt 1
+ add in seperator colum count of nominations --> implemented
- check if person is allocated more than once
- generate a pdf
"""

import win32com.client
from win32com.client import constants as constants
import os 
from datetime import datetime
import locale
from fpdf import FPDF


# (0) Setup
dir_path = os.path.dirname(os.path.realpath(__file__))
print("Dirpath:{}#".format(dir_path))

filename=dir_path + "\\" + "Liturgieplan.xlsx"
pdfname=dir_path + "\\" + "Liturgieplan_Aug_Sep_25.pdf"
sheet="Aug25"
content=[]
wb=None;
ws=None;

def ExcelSummary():
    global filename, sheet, content,wb,ws
    print("ExcelZusammenfassung")
    #excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
    excel = win32com.client.Dispatch('Excel.Application')
    
    #excel = win32com.client.dynamic.Dispatch('Excel.Application')
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(filename)
    colstart=3
    line=4
    if wb is None:
        print("Workbook not found for {}".format(filename))
        #return False
    print("sheet:{}".format(sheet))
    ws = wb.Worksheets(sheet)
    if ws == None:
        error=True
        print ("Sheet not found")
    lektoren={}
    locale.setlocale(locale.LC_ALL, 'de_DE.UTF-8')
    col=colstart
    while ((ws.Cells(line,col).Value) is not None):
        if ("Spalte" in ws.Cells(line,col).Value):
            break
        print(ws.Cells(line,col).Value)
        lektoren[col]=ws.Cells(line,col).Value
        col+=1
    seperatorcol1=col
    col+=1
    kommunionhelfer={}
    while ((ws.Cells(line,col).Value) is not None):
        if ("Spalte" in ws.Cells(line,col).Value):
            break
        #print(ws.Cells(line,col).Value)
        kommunionhelfer[col]=ws.Cells(line,col).Value
        col+=1
    #last valid col
    seperatorcol2=col

    col=0    
    line+=1
    resultline=40
    #write header
    ws.Cells(resultline,2).Value="LektorInnen"    
    ws.Cells(resultline,4).Value="KommunionhelferInnen"
    resultline+=1        
    #find colored 
    error=False
    lastdate=None
    while ((ws.Cells(line,1).Value) is not None):
        print("line:{}".format(line))
        if len(ws.Cells(line,1).Value)==0:
            break
        date=ws.Cells(line,1).Value+"_"+ws.Cells(line,2).Value
        

        if lastdate is not None and lastdate!=date:
            #difference in date
            differencedate = datetime.strptime(date, "%Y-%m-%d_%H:%M") - datetime.strptime(lastdate, "%Y-%m-%d_%H:%M")
            if differencedate.days>1:
                print("Date {} is not continuous in line {}".format(date,line))
                #build a excel range
                #print(constants.xlEdgeTop)
                #ws.Range(ws.Cells(line,1),ws.Cells(line,seperatorcol2)).Borders(constants.xlEdgeTop).ColorIndex = 1
                #ws.Range(ws.Cells(line,1),ws.Cells(line,seperatorcol2)).Borders(constants.xlEdgeTop).Weight = 2
                ws.Range(ws.Cells(line,1),ws.Cells(line,seperatorcol2)).Borders(7).ColorIndex = 1
                ws.Range(ws.Cells(line,1),ws.Cells(line,seperatorcol2)).Borders(7).Weight = 2
        lastdate=date

            
        date_object = datetime.strptime(date, "%Y-%m-%d_%H:%M")
        datestr=date_object.strftime("%a, %d.%b %H:%M")
        # print(datestr)
        lektor=None
        lektor2=None
        kommunion=None
        nomination1=0
        nomination2=0
        for col2 in range(colstart,seperatorcol2):
            if ws.Cells(line,col2).Interior.ColorIndex==6:
                if col2 in lektoren.keys():
                    if lektor==None :
                        #print("Lektor {} {}".format(datestr,lektoren[col2]))
                        lektor=lektoren[col2]
                        nomination1+=1
                    elif lektor2==None and datestr.find("8:15"):
                        lektor2=lektoren[col2]
                        nomination1+=1
                    else:    
                        print ("ERROR triple Lektor in line {}".format(line))
                        nomination1=0
                        error=True
                        lektor=None
                        lektor2=None
                        break
                if col2 in kommunionhelfer.keys():
                    if kommunion==None:
                        #print("Kommunion {} {}".format(datestr,kommunionhelfer[col2]))
                        kommunion=kommunionhelfer[col2]
                        nomination2+=1
                    else:    
                        print ("ERROR duplicate Kommunionhelfer in line {}".format(line))
                        nomination2=0
                        kommunion=None
                        error=True
                        break    
        if lektor:
            ws.Cells(resultline,1).Value=datestr
            ws.Cells(resultline,2).Value=lektor
            #print("write result {}".format(resultline))
        else:    
            print("Error for lektor in line {} on date:{}".format(line,datestr))    
            error=True
            break
        if lektor2:
            ws.Cells(resultline,3).Value=lektor2
        else:
            lektor2=""
        if kommunion:
            ws.Cells(resultline,4).Value=kommunion
        else:    
            kommunion=""
        print ("{}\t{};\t{}#\t{}".format(datestr,lektor,lektor2,kommunion)) 
        #for pdf
        row_data = [
            datestr,
            lektor,
            lektor2,
            kommunion]
        content.append(row_data)

        ws.Cells(line,seperatorcol1).Value=nomination1
        ws.Cells(line,seperatorcol2).Value=nomination2
        line+=1
        resultline+=1    
        if error==False:
            print("Normales Ende in line {}".format(line))
        else:
            break    
    return not error

def create_pdf(header,conent,outfile):
    global dir_path
    print("create_pdf for {}".format(outfile))
    # Create instance of FPDF class
    pdf = FPDF()

    # Add a page
    pdf.add_page()

    # Add a title
    pdf.set_font("Arial", size=16, style="B")
    pdf.cell(125, 10, txt = "Liturgieplan "+header, align = 'L')
    pdf.set_font("Arial", size=12)
    current_date = datetime.now().strftime("%d.%m.%Y")
    #pdf.cell(200, 10, txt = f"            Stand:  {current_date}", ln = True)
    scaler=10
    pdf.image(dir_path+"\\cropped-Logo_St-Franziskus-Coloman_RGB_2024.PNG",  w = 577/scaler, h=130/scaler)
    pdf.ln()
    
    # Add a line break
    #pdf.ln()

    # Add some text
    #pdf.cell(200, 10, txt = "Tabelle.", ln = True)
    
    # Define table column headers
    headers = ["Datum/Uhrzeit", "LektorInnen", "LektorInnen", "KommunionhelferInnen "]
    col_widths = [40, 45, 45, 50]
    pdf.set_font("Arial", size=10)
    # Print headers
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 10, header, border=1, align='C')
    pdf.ln()

    # Add rows
    high=12
    for row_data in conent:
        for i, item in enumerate(row_data):
            pdf.cell(col_widths[i], high, item, border=1, align='C')
        pdf.ln()
    # Save the PDF to a file
    pdf_file_path = outfile
    pdf.output(pdf_file_path)


if __name__ == "__main__":
    if ExcelSummary():
        create_pdf("August September 2025",content,pdfname)
        print("PDF created successfully.")






    
    