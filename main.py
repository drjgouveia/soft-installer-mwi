import subprocess
from tkinter import *
import os
import requests
import winget_export
import ctypes, sys
import getpass


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


USER_NAME = getpass.getuser()
# Format: (NAME, PACKAGE NAME IN WINGET, PACKAGE NAME IN CHOCOLATEY)
software = [("Anydesk", "AnyDeskSoftwareGmbH.AnyDeskMSI", "anydesk"),
            ("Firefox", "Mozilla.FirefoxESR", "firefox"),
            ("Google Chrome", "Google.Chrome", "googlechrome"),
            ("7-Zip", "7zip.7zip", "7zip"),
            ("K-Lite Codecs", "CodecGuide.K-LiteCodecPackStandard", "k-litecodecpack-standard"),
            ("VLC", "VideoLAN.VLC", "vlc"),
            ("Zoom", "Zoom.Zoom", "zoom"),
            ("Adobe Reader DC", "Adobe.AdobeAcrobatReaderDC", "adobereader"),
            ("LibreOffice", "LibreOffice.LibreOffice", "libreoffice-fresh"),
            ("OpenJDK 14", "AdoptOpenJDK.OpenJDK.14", "adoptopenjdk14"),
            ("OpenJDK 15", "AdoptOpenJDK.OpenJDK.15", "adoptopenjdk15openj9"),
            ("OpenJDK 16", "AdoptOpenJDK.OpenJDK.16", "adoptopenjdkopenj9")]

soft = []


class MyFirstGUI:
    def __init__(self, master):
        global soft
        self.master = master
        self.master.title("Software MWI installer")
        self.top = Frame(self.master)
        self.mid = Frame(self.master)
        self.mid2 = Frame(self.master)
        self.bottom = Frame(self.master)
        self.top.pack(side=TOP)
        self.mid.pack(side=TOP)
        self.mid2.pack(side=TOP, ipady=15)
        self.bottom.pack(side=BOTTOM, fill=BOTH, expand=True, ipady=10, pady=10)

        self.label = Label(master, text="Elija el software")
        self.label.pack(in_=self.top, side=TOP, fill=X, ipady=10)
        self.label.config(font=("Arial", 16), anchor=CENTER)

        #        i = 1
        #        u = 0
        k = 0
        #        self.label.grid(row=0, column=0, columnspan=2)

        soft = []
        for name, package, choco in software:
            string = str(name) + "," + str(package) + "," + str(choco)
            soft.append(Variable(value=(0, string)))
            self.option_button = Checkbutton(master, width=75, text=name, onvalue=(1, string), offvalue=(0, string),
                                             variable=soft[k])
            #            self.option_button.grid(row=i, column=u)
            self.option_button.pack(in_=self.mid, fill=X)
            self.option_button.select()
            k += 1

        self.install_button = Button(master, text="Install", command=self.install)
        self.install_button.config(anchor=CENTER)
        self.install_button.pack(in_=self.mid2, side=LEFT, ipadx=50, padx=50)

        self.update_button = Button(master, text="Update Windows", command=self.update)
        self.update_button.config(anchor=CENTER)
        self.update_button.pack(in_=self.mid2, side=RIGHT, ipadx=20, padx=20)

        self.close_button = Button(master, text="Close", command=master.quit)
        self.close_button.config(anchor=CENTER)
        self.close_button.pack(in_=self.bottom, side=BOTTOM, ipadx=50, padx=50)

    def install(self):
        total = 0
        for s in soft:
            active = int(s.get()[0])
            if active == 1:
                total += 1

        ini = "winget install --id="
        line = ini
        x = 0

        for s in soft:
            active = int(s.get()[0])
            if active == 1:
                try:
                    name = str(s.get()).split("'")[1].replace("{", "").replace("}", "").split(",")[0].replace("1 ", "")
                    winget = str(s.get()).split("'")[1].replace("{", "").replace("}", "").split(",")[1]
                    choco = str(s.get()).split("'")[1].replace("{", "").replace("}", "").split(",")[2]
                except IndexError:
                    name = str(s.get()).split("'")[0].replace("{", "").replace("}", "").split(",")[0].replace("1 ", "")
                    winget = str(s.get()).split("'")[0].replace("{", "").replace("}", "").split(",")[1]
                    choco = str(s.get()).split("'")[0].replace("{", "").replace("}", "").split(",")[2]

                line = line + winget + " -e --silent"
                return_code = os.system(line)

                if return_code != 0:
                    print("Winget dio un error. Empiezando con Chocolatey.")
                    return_code = subprocess.run(["powershell", "-Command", "choco install " + str(choco) + " -y"], capture_output=True)

                    if return_code != 0:
                        winget_export.export()
                        self.popupmsg(name + " could not be installed. Please install the winget manually (file in the same place as the executable)\nand click okay for the program to restart.")
                        break

                line = ini

    def restartProg(self):
        winget_export.delete()
        self.popup.destroy()
        self.master.destroy()
        os.system('cls')
        main()

    def popupmsg(self, msg):
        self.popup = Tk()
        self.popup.title("Error - Package not installed")
        label = Label(self.popup, text=msg)
        label.grid(column=0, row=0, pady=15, padx=15, ipadx=15)

        B1 = Button(self.popup, text="Okay", command=self.restartProg)
        B1.grid(column=0, row=1, pady=15)
        self.popup.update()

    def update(self):
        beforePolicy = subprocess.run(["powershell", "-Command", "Get-ExecutionPolicy"], capture_output=True).stdout.decode("utf-8")
        subprocess.run(["powershell", "-Command", "Set-ExecutionPolicy Unrestricted -Force"], capture_output=True)

        print("Instalando el Updater por la CMD.")
        subprocess.run(["powershell", "-Command", "Install-Module PSWindowsUpdate -Force"], capture_output=True)

        print("Empiezando a procurar atualizaciones.")
        procurandoUpdate = subprocess.run(["powershell", "-Command", "Import-Module PSWindowsUpdate; Get-WindowsUpdate; Install-WindowsUpdate -ForceDownload -ForceInstall -AcceptAll -IgnoreReboot"],
                                       capture_output=True)
        print("Actualizado.")
        subprocess.run(["powershell", "-Command", "Set-ExecutionPolicy" + beforePolicy + " -Force"], capture_output=True)


def main():
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        exit()

    url = "http://www.example.com"
    timeout = 5
    try:
        request = requests.get(url, timeout=timeout)
        print("Conectado a la internet.\n")
    except (requests.ConnectionError, requests.Timeout) as exception:
        print("No hay conexion a la internet.")
        if input("Te gustaria continuar? (Y/n) ") in ("N", "n"):
            exit()

    try:
        print("Instalando WinGet...")
        winget_export.export()
        cmd = "Add-AppxPackage \".\\winget.appxbundle\""
        completed = subprocess.run(["powershell", "-Command", cmd], capture_output=True)
        winget_export.delete()
        print("Instalado.\n")

        print("Instalando Chocolatey...")
        cmd = "Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))"
        subprocess.run(["powershell", "-Command", cmd], capture_output=True)
        subprocess.run(["powershell", "-Command", "choco upgrade chocolatey"], capture_output=True)
        print("Instalado.\n")
    except Exception as e:
        pass

    print("Empienzando la GUI.")
    root = Tk()
    x = 100 + 30 * len(software)
    root.geometry("400x" + str(x))
    my_gui = MyFirstGUI(root)
    root.mainloop()
    print("Bye!")


if __name__ == '__main__':
    main()