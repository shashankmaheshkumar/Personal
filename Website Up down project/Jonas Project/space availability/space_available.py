import os.path
import wmi
import shutil
from skpy import Skype


def testdrive(tested_drive):
    if shutil.disk_usage(tested_drive).free // (2 ** 40) <= 1 and tested_drive:
        for skype_contact in skype_contacts:
            Skype_account.contacts[skype_contact].chat.sendMsg("Drive " + tested_drive + " is running low on space. "
                                                                                  "\nCurrently available space: " +
                                                               str(shutil.disk_usage(tested_drive).free // (2 ** 30)) + "GB")


if __name__ == '__main__':
    Skype_account = Skype("support@educube.net", "8a9znalb", "tokenfile")
    skype_contacts = ["chetan.t.v", "shashank.globals"]
    while not os.path.exists("./tokenfile"):
        open("tokenfile", "w").close()
    drives = wmi.WMI().Win32_LogicalDisk(DriveType=3)
    for drive in drives:
        try:
            testdrive(drive.Caption)
        except FileNotFoundError:
            drives.remove(drive)
