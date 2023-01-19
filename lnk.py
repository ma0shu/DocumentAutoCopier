'''
OfficeAutoCopier
Author:Myy-ShanDongExpHighSchool2021
Ver:0.99.20230113 (Beta)
'''
import sys, win32com.client, shutil, os.path, datetime, time
from webdav4.client import Client
shell = win32com.client.Dispatch("WScript.Shell")

#【Remote Path】
# Webdav needed, "Nextcloud" or "Kodbox" recommend.
client = Client(base_url='(InputYourWebdavAddressHere)', auth=('(InputYourUsernameHere)', '(InputYourPasswordHere)'))

#【Local Path】
# File will be placed at "~/Desktop/Document".
BaseDir=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\\")
IndexDat=os.path.join(os.environ['USERPROFILE'],"AppData\Roaming\Microsoft\Office\Recent\index.dat")
TargetDir=os.path.join(os.environ['USERPROFILE'],"Desktop\Document\\")
if not os.path.exists(TargetDir):
    os.makedirs(TargetDir)

#【Loop Running】
while True:
    '''
    Some version of MSOffice will automatically create "index.dat" log file in user's recent folder, leading to errors. Try deleting it.
    '''
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
        client.upload_file(from_path=os.path.join(TargetDir,filename), to_path=os.path.join('/Documents',filename), overwrite=True)
    '''
    Check new file every 30 seconds in order to save perf. 
    This may cause only the last file copied when opening two files in a short time, 
    so if you want to balance it by making your own choices, edit it.
    '''
    time.sleep(30)
