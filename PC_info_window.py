import PySimpleGUI as sg
import psutil
import subprocess
import wmi

#Definition of WMI classes
c = wmi.WMI()   
my_system = c.Win32_ComputerSystem()[0]
sku_system = c.Win32_ComputerSystemProduct()[0]
bios_system = c.Win32_Bios()[0]
os_system = c.Win32_OperatingSystem()[0]
ram_system = c.Win32_PhysicalMemoryArray()[0]
cpu_system = c.Win32_Processor()[0]
gpu_system = c.Win32_VideoController()[0]
ram_cimv = c.CIM_PhysicalMemory()[0]
disk_cimv = c.CIM_DiskDrive()[0]
bios_cimv = c.CIM_BIOSElement()[0]

#Theme layout
sg.theme('DarkAmber')

#Psutil information
def get_ram_amount():
    return f"{psutil.virtual_memory().total / (1024.0 ** 3):.2f} GB"
def get_hdd_amount():
    return f"{psutil.disk_usage('/').total / (1024.0 ** 3):.2f} GB"
def hostname():
  return subprocess.check_output(f'wmic bios get SerialNumber').decode('utf-8').split('\n')[1:]


#Main window
layout = [[sg.Text("Serialnumber", size=(15,1)), sg.Text(f"{bios_system.SerialNumber}", font=("tahoma", 23))],
          [sg.Text("UUID", size=(15,1)), sg.Text(f"{sku_system.UUID}", size=(68))],
          [sg.Text("Hostname", size=(15,1)), sg.Text(f"{my_system.Caption}", font=("tahoma", 23))],
#          [sg.Text("CPU info", size=(15,1)), sg.Text(f"{cpu_system.Name}", size=(68))],
          [sg.Text("GPU info", size=(15,1)), sg.Text(f"{gpu_system.Description}", size=(68))],
          [sg.Text("GPU info", size=(15,1)), sg.Text(f"{gpu_system.VideoProcessor}", size=(68))],
          [sg.Text("RAM", size=(15,1)), sg.Text(get_ram_amount(), size=(11))],
          [sg.Text("RAM modules", size=(15,1)), sg.Text(f"{ram_system.MemoryDevices}", size=(68))],
          [sg.Text("RAM speed", size=(15,1)), sg.Text(f"{ram_cimv.Speed} MHz", size=(68))],
          [sg.Text("RAM pn", size=(15,1)), sg.Text(f"{ram_cimv.PartNumber}", size=(68))], 
          [sg.Text("Harddrive SN:", size=(15,1)), sg.Text(f"{disk_cimv.SerialNumber}", font=("tahoma", 11))],
          [sg.Text("BIOS", size=(15,1)), sg.Text(f"{bios_cimv.Caption}", font=("tahoma", 11))],

          [sg.Button('Exit', key='Exit')]]

window = sg.Window('SMO CZ IT - Computer Information', layout, size=(800,450))

while True:
    event, values = window.read()
    if event in (None, 'Exit'):
        break

window.close()
