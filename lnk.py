# OfficeAutoCopier
# Author:MyySDSY2021
# Version:0.99.20211127 (Unstable)

import sys, win32com.client, shutil, os.path, datetime, time

base_dir=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\\")
target_dir=os.path.join(os.environ['USERPROFILE'],"Desktop\PPTS\\")
index_dat=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\index.dat")
l_oldl=os.listdir(base_dir)
l_oldl.sort(key=lambda fn: os.path.getmtime(base_dir+fn) if not os.path.isdir(base_dir+fn) else 0)
l_old=l_oldl[-1]

while True:
    if os.path.exists(index_dat) == True:
        os.remove(index_dat)
    l=os.listdir(base_dir)
    l.sort(key=lambda fn: os.path.getmtime(base_dir+fn) if not os.path.isdir(base_dir+fn) else 0)
    shell = win32com.client.Dispatch("WScript.Shell")
    if l_old != l[-1] :
        shortcut = shell.CreateShortCut(base_dir+l[-1])
        l_old = l[-1]
        shutil.copy(shortcut.Targetpath, target_dir)
    time.sleep(60)
