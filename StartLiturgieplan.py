
"""
version 1.1 sort filelist
@author: Volker.Erben@gmail.com
"""
import docx2txt
import re
from datetime import datetime
import locale

def readWordFile():
    text = docx2txt.process("Terminplan für Lektoren_Kommunionhelfer_08_09_25.docx")
    textarr = text.split("\n")
    locale.setlocale(locale.LC_ALL, 'de_DE.UTF-8')
    datestr=None
    datestr_save=None
    resultline=None
    count=0
    output_file = open("input.txt", "w", encoding="utf-8")
    for line in textarr:
        #Sonntag, 1. Dezember 2024
        datematch1=re.search(r'(\d+\.\s\w+\s20\d\d)',line)
        datematch2=re.search(r'\s(\d+\.\d+\.20\d\d)',line)
        #Samstag, 01.02.2025
        if datematch1:
            #print(datematch1.group(0))
            date_object = datetime.strptime(datematch1.group(0), "%d. %B %Y")
            datestr=date_object.strftime("%Y-%m-%d")
            if datestr is datestr_save:
                datestr=None
                #print ("same2")
            else:
                datestr_save=datestr
                print (resultline)
                if resultline:
                    output_file.write(resultline + "\n")

            #print (count)
            resultline=""
        elif datematch2:
            #print(datematch2.group(0))
            date_object = datetime.strptime(datematch2.group(1), "%d.%m.%Y")
            datestr=date_object.strftime("%Y-%m-%d")
            if datestr == datestr_save:
                datestr=None
                #print ("same1")
            else:
                datestr_save=datestr
                #print ("diff "+ datestr_save)
                print (resultline)
                if resultline:
                    output_file.write(resultline + "\n")
            #print (count)
            
        timematch1=re.search(r"(\d+)\.(\d+) Uhr", line)
        timematch2=re.search(r"(\d+):(\d+)", line)
        if timematch1:
            timestr=timematch1.group(1)+":"+timematch1.group(2)
            count+=1
            
            if datestr:
                resultline=datestr+";"+timestr
                #print (line)
                datestr=None
            else:
                if len(resultline):
                    resultline+=";"+timestr
        elif timematch2:
            timestr=timematch2.group(1)+":"+timematch2.group(2)
            count+=1
            
            if datestr:
                resultline=datestr+";"+timestr
                #print (line)
                datestr=None
            else:
                if len(resultline):
                    resultline+=";"+timestr            
                    
    #last entry
    if resultline:
        print (resultline)
        output_file.write(resultline + "\n")

    print ("Check: Serach (:) - pages: {} and check how many pages".format(count))
    output_file.close()


if __name__ == "__main__":
    readWordFile()
    #todo start ahk