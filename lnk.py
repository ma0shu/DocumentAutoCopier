# OfficeAutoCopier
# Author:MyySDSY2021
# Version:0.9.20211126 (Unstable)

import sys, win32com.client, shutil, os.path, datetime, time

base_dir=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\\")
index_dat=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\index.dat")
l_old = "0"

while True:
    if os.path.exists(index_dat) == True:
        os.remove(index_dat)
    l=os.listdir(base_dir)
    l.sort(key=lambda fn: os.path.getmtime(base_dir+fn) if not os.path.isdir(base_dir+fn) else 0)
    shell = win32com.client.Dispatch("WScript.Shell")
    if l_old != l[-1] :
        print('已更新：'+l[-1])
        shortcut = shell.CreateShortCut(base_dir+l[-1])
        l_old = l[-1]
        shutil.copy(shortcut.Targetpath, 'D:\\PPTS')
    time.sleep(60)    
 
