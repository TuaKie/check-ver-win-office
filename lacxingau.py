import os
import time

print("Check System Win & Office")
print("Type -h to know more !!!")
def act():
    a = input("> ")
    if a == "w" or a == "W":
        print("Windows Checking System")
        print("--------------Checking...!--------------")
        time.sleep(3)
        os.system("cscript %windir%\system32\slmgr.vbs /dli & cscript %windir%\system32\slmgr.vbs /xpr")
        time.sleep(0.6)
        print("--------------DONE!--------------")
    elif a == "o" or a == "O":
        print("Office Checking System")
        print("--------------Checking...!--------------")
        time.sleep(3)
        os.system(""" for %a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%a") else (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%a")) """)
        time.sleep(3)
        os.system("""cscript %windir%\system32\slmgr.vbs /dli & cscript %windir%\system32\slmgr.vbs /xpr & cscript ospp.vbs /dstatus """)
        time.sleep(0.6)
        print("--------------DONE!--------------")
    elif a == "-h":
            print("""Type w or W to check Windows System.\nType o or-h O to check Office System.""")  
    else:
        print("Error !!! Try Again")
        
act()

