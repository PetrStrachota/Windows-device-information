import PySimpleGUI as sg
import psutil
import subprocess
import wmi
import datetime

#Definition of Version print
version = "Version v1.0 29-1-2023"

#Definition of Date and time print
now = datetime.datetime.now()
my_date = now.strftime("%Y-%m-%d %H:%M")

#Definition of WMI classes
#http://timgolden.me.uk/python/wmi/tutorial.html

c = wmi.WMI()   
my_system = c.Win32_ComputerSystem()[0]
sku_system = c.Win32_ComputerSystemProduct()[0]
bios_system = c.Win32_Bios()[0]
os_system = c.Win32_OperatingSystem()[0]
ram_system = c.Win32_PhysicalMemoryArray()[0]
ramdev_system = c.Win32_PhysicalMemory()[0]
cpu_system = c.Win32_Processor()[0]
gpu_system = c.Win32_VideoController()[0]
ram_cimv = c.CIM_PhysicalMemory()[0]
ram_cimv2 = c.CIM_PhysicalMemory()[1]
disk_cimv = c.CIM_DiskDrive()[0]
bios_cimv = c.CIM_BIOSElement()[0]

#Memory type function
smbiosramt = int(f"{ramdev_system.SMbiosmemorytype}")
def ram_type(smbiosramt):
    if smbiosramt == 10:
        return "ROM"
    elif smbiosramt == 11:
        return "Flash"
    elif smbiosramt == 12:
        return "EEPROM"
    elif smbiosramt == 13:
        return "FEPROM"
    elif smbiosramt == 14:
        return "EPROM"
    elif smbiosramt == 21:
        return "DDR2"
    elif smbiosramt == 22:
        return "DDR2 FB-DIMM"
    elif smbiosramt == 24:
        return "DDR3"
    elif smbiosramt == 25:
        return "FBD2"
    elif smbiosramt == 26:
        return "DDR4"
    else:
        return "N/A"

ramrestype = (ram_type(smbiosramt))

#Boot time definition
boot_time = datetime.datetime.fromtimestamp(psutil.boot_time()).strftime("%Y-%m-%d %H:%M:%S")

#hdd size
hdd_dev = (int(f"{disk_cimv.Size}"))
hdd_size = str(round(((hdd_dev)/(1024.0 ** 3)), 2))

#Psutil information
def get_ram_amount():
    return f"{psutil.virtual_memory().total / (1024.0 ** 3):.2f} GB"
def get_hdd_amount():
    return f"{psutil.disk_usage('/').total / (1024.0 ** 3):.2f} GB"
def hostname():
  return subprocess.check_output(f'wmic bios get SerialNumber').decode('utf-8').split('\n')[1:]

#Theme for dashboard theme_dict
theme_dict = {'BACKGROUND': '#2B475D',
                'TEXT': '#FFFFFF',
                'INPUT': '#F2EFE8',
                'TEXT_INPUT': '#000000',
                'SCROLL': '#F2EFE8',
                'BUTTON': ('#000000', '#C2D4D8'),
                'PROGRESS': ('#FFFFFF', '#C7D5E0'),
                'BORDER': 1,'SLIDER_DEPTH': 0, 'PROGRESS_DEPTH': 0}

# sg.theme_add_new('Dashboard', theme_dict)     # if using 4.20.0.1+
sg.LOOK_AND_FEEL_TABLE['Dashboard'] = theme_dict
sg.theme('Dashboard')

BORDER_COLOR = '#C7D5E0'
DARK_HEADER_COLOR = '#1B2838'
BPAD_TOP = ((20,20), (20, 10))
BPAD_LEFT = ((20,10), (0, 10))
BPAD_LEFT_INSIDE = (0, 10)
BPAD_RIGHT_INSIDE = ((0,10), (0,10))
BPAD_RIGHT = ((10,20), (10, 20))


#Theme layout
sg.theme('Dashboard')

#Main window

top_banner = [[sg.Text('SMO-IT', font='Any 14', background_color=DARK_HEADER_COLOR),
                sg.Text(my_date, font='Any 14', background_color=DARK_HEADER_COLOR)]]

top  = [[sg.Text('SMO CZ IT - Installation center', size=(50,1), justification='c', pad=BPAD_TOP, font='Any 20')],
            [sg.T(version)],]

block_1 = [[sg.Text('Identification', font='Any 16')],
          [sg.Text("Hostname", size=(15,1)), sg.Text(f"{my_system.Caption}", font=("tahoma", 23))],
          [sg.Text("Serialnumber", size=(15,1)), sg.Text(f"{bios_system.SerialNumber}", font=("tahoma", 23))],
          [sg.Text("UUID", size=(15,1)), sg.Text(f"{sku_system.UUID}", size=(68))],
          [sg.Text("Model", size=(15,1)), sg.Text(f"{my_system.Model}", size=(68))],]

block_2 = [[sg.Text('System information', font='Any 16')],
          [sg.Text("OS", size=(15,1)), sg.Text(f"{os_system.Caption} / build.{os_system.Version}", size=(68))],
          [sg.Text("OS last boot", size=(15,1)), sg.Text(boot_time, size=(68))],
          [sg.Text("Domain", size=(15,1)), sg.Text(f"{my_system.Domain}", size=(24))],
             ]

block_3 = [[sg.Text('Hardware specs', font='Any 16')],

          [sg.Text("RAM info", size=(15,1)), sg.Text(get_ram_amount()),sg.Text(f" in {ram_system.MemoryDevices} slot(s)"), sg.Text(ramrestype) ,sg.Text(f"{ram_cimv.Speed} MHz", size=(68))],
          [sg.Text("RAM PartNumber", size=(15,1)),sg.Text(f"{ram_cimv.partnumber}{ram_cimv.Serialnumber}")],[sg.Text("", size=(15,1)),sg.Text(f"{ram_cimv2.partnumber}{ram_cimv2.Serialnumber}")],
          [sg.Text("GPU info", size=(10,1)), sg.Text(f"{gpu_system.Description}", size=(68))],
          [sg.Text("GPU core", size=(10,1)), sg.Text(f"{gpu_system.VideoProcessor}", size=(68))],
          [sg.Text("Harddrive", size=(10,1)), sg.Text(f"{disk_cimv.Model}"), sg.Text(hdd_size),sg.Text("GiB")],]


block_4 = [[sg.Text('Bios', font='Any 16')],
          [sg.Text("BIOS", size=(10,1)), sg.Text(f"{bios_cimv.Caption}", font=("tahoma", 11))],]


#Blocks layout
layout = [[sg.Column(top_banner, size=(960, 60), pad=(0,0), background_color=DARK_HEADER_COLOR)],
          [sg.Column(top, size=(920, 90), pad=BPAD_TOP)],
          
          [sg.Column([[sg.Column(block_1, size=(450,200), pad=BPAD_LEFT_INSIDE)],
                      [sg.Column(block_2, size=(450,150),  pad=BPAD_LEFT_INSIDE)]], pad=BPAD_LEFT, background_color=BORDER_COLOR),
           
           sg.Column([[sg.Column(block_3, size=(450,215), pad=BPAD_LEFT_INSIDE)],
                      [sg.Column(block_4, size=(450,135),  pad=BPAD_LEFT_INSIDE)]], pad=BPAD_RIGHT, background_color=BORDER_COLOR)],     
        
           
#input_box         [sg.Text('Enter something on Row 2'), sg.InputText()],

          [sg.Button('Exit', key='Exit', pad=(460,15))]

#window = sg.Window('SMO CZ IT - Computer Information', layout, size=(800,450))
window = sg.Window('SMO CZ IT - Computer Information', layout, margins=(0,0), background_color=BORDER_COLOR, no_titlebar=True, grab_anywhere=True)



while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

window.close()
