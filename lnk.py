'''
OfficeAutoCopier
Author:Myy-ShanDongExpHighSchool2021
Ver:0.99.20220227 (Unstable-Beta)
'''
import sys, win32com.client, shutil, os.path, datetime, time
from webdav4.client import Client
shell = win32com.client.Dispatch("WScript.Shell")

#【Remote Path】
# Webdav needed, "Nextcloud" or "Kodbox" recommend.
client = Client(base_url='(InputYourWebdavAddressHere!)', auth=('(InputYourUsernameHere!)', '(InputYourPasswordHere!)'))

#【Local Path】
# File will be placed at "~/Desktop/PPTS".
BaseDir=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\\")
IndexDat=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\index.dat")
TargetDir=os.path.join(os.environ['USERPROFILE'],"Desktop\PPTS\\")
if not os.path.exists(TargetDir):
    os.makedirs(TargetDir)

#【Loop Running】
while True:
    # Some version of MSOffice will automatically create "index.dat" log in our recent folder, which lead to errors. Delete it.
    if os.path.exists(IndexDat) == True:
        os.remove(IndexDat)  
    #【Sort Files】
    link=os.listdir(BaseDir)
    link.sort(key=lambda fn: os.path.getmtime(BaseDir+fn) if not os.path.isdir(BaseDir+fn) else 0)
    filename=os.path.splitext(link[-1])[0]
    #【Check Existence】
    if not os.path.exists(os.path.join(TargetDir,filename)):
        #【Locate&Copy】
        shortcut = shell.CreateShortCut(BaseDir+link[-1])
        shutil.copy(shortcut.Targetpath, TargetDir)
        #【Upload2Webdisk】
        client.upload_file(from_path=os.path.join(TargetDir,filename), to_path=os.path.join('/YourPath',filename), overwrite=True)
    '''
    Check new file per minute in order to save perf. 
    This may cause only the last one to be copied when opening two files in a short time, 
    so make your own choices.
    '''
    time.sleep(60)
