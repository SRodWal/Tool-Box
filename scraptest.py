# -*- coding: utf-8 -*-
"""
Created on Tue Dec 29 09:25:14 2020

@author: serw1
"""

ODSurl = ["https://www.ods.org.hn/index.php/informes/costes-marginales/2020/mayo",
          "https://www.ods.org.hn/index.php/informes/costes-marginales/2020/junio",
          "https://www.ods.org.hn/index.php/informes/costes-marginales/2020/julio",
          "https://www.ods.org.hn/index.php/informes/costes-marginales/2020/agosto",
          "https://www.ods.org.hn/index.php/informes/costes-marginales/2020/septiembre",
          "https://www.ods.org.hn/index.php/informes/costes-marginales/2020/octubre-cm20",
          "https://www.ods.org.hn/index.php/informes/costes-marginales/2020/noviembre-cm20",
          "https://www.ods.org.hn/index.php/informes/costes-marginales/2020/diciembrecostosm2020"]
meses = ["May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
exurl = "http://www.pythonscraping.com/exercises/exercise1.html"
brow = r"C:\\Users\\serw1\\AppData\\Local\\Programs\\Opera\\launcher.exe"
downpath = "C:/Users/serw1/Downloads"
datapath = "C:/Users/serw1/Desktop/Documents/Work/ODS Price analysis\Data"
#Desired urls
from urllib.request import urlopen
from urllib.request import urlretrieve, URLopener
from time import sleep
import re, os, requests, http.client, shutil
import webbrowser

http.client.HTTPConnection._http_vsn = 10
http.client.HTTPConnection._http_vsn_str = 'HTTP/1.0'
headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    }
for url, mes in zip(ODSurl, meses):
    
    page = urlopen(url)
    html_bytes = page.read()
    html = html_bytes.decode("utf-8")
    data = re.findall("href=.*?.xlsx",html)
    webbrowser.register("opera", None,webbrowser.BackgroundBrowser(brow))
    op  =  webbrowser.get("opera")
    k = 1
    print("Iniciando mes de "+mes+". Total de archivos: "+str(len(data)))
    for l in data:
        links = (l[l.find("http"):len(l)])
        fileloc = downpath+"/"+links[links.find("Precios"):len(links)]
        fileout = "CM_"+mes+"20_S"+str(k)+".xlsx"
        outloc = datapath+"/"+fileout
        print("Downloading file : "+links[links.find("Precios"):len(links)])
        op.open_new_tab(links)
        sleep(3)
        print("Moving and renaming : "+fileout)
        os.rename(fileloc,outloc)
        sleep(2)
        print("File Completed")
        k=k+1