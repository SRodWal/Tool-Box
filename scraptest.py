# -*- coding: utf-8 -*-
from urllib.request import urlopen
from time import sleep
import re, os, http.client
import webbrowser
ODSurl = ["https://www.ods.org.hn/index.php/informes/costes-marginales/costosmarginales2019"
          ]
#List of webpages to search
Meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
#This will order the month list to correspond the ODS links.
meses = []
for l in ODSurl:
    for m in Meses:
        if m.lower() in l:
            meses.append(m)

#Defines download, repository, and browser's directory
brow = r"C:\\Users\\serw1\\AppData\\Local\\Programs\\Opera\\launcher.exe"
downpath = "C:/Users/serw1/Downloads"
scrippath = os.getcwd().replace("\\","/")
pdfpath = scrippath+"/Data/DailyReports"
excelpath = scrippath+"/Data"
filetype = ".xlsx"

# Location of download files and storage 
for url, mes in zip(ODSurl,meses):
    page = urlopen(url)
    html_bytes = page.read()
    html = html_bytes.decode("utf-8")
    data = re.findall("href=.*?"+filetype,html)
    data[0] = data[0][len(data[0])-len(data[1]):len(data[0])] 
    webbrowser.register("opera", None,webbrowser.BackgroundBrowser(brow))
    op  =  webbrowser.get("opera")
    print("Iniciando mes de "+mes+". Total de archivos: "+str(len(data)))
    for l in data:
        date = l[len(l)-11:-5]
        day= date[0:2]
        yr = date[4:6]
        links = (l[l.find("http"):len(l)])
        fileloc = downpath+"/"+links[links.find("Predespacho"):len(links)]
        fileout = "CM_"+mes+str(yr)+"_S"+str(day)+filetype
        outloc = excelpath+"/"+fileout
        print("Downloading file : "+links[links.find("Predespacho"):len(links)])
        op.open_new_tab(links)
        sleep(9)
        print("Moving and renaming : "+fileout)
        try:
            os.rename(fileloc,outloc)
        except FileNotFoundError:
            print("File not found in downloads")
            pass
        except FileExistsError:
            print("File already on directory - deleting file.")
            os.remove(fileloc)     
        print("Operation Completed")