import wmi


def get_info():
	wmi_conn = wmi.WMI()   
	my_system = wmi_conn.Win32_ComputerSystem()[0]
	bios = wmi_conn.Win32_Bios()[0]
	os_system = wmi_conn.Win32_OperatingSystem()[0]
	return my_system, bios, os_system


if __name__ == "__main__":
	my_system, bios, os_system = get_info()
	info = []
	
	info.append(f"SerialNumber: \t\t{bios.SerialNumber}")
	info.append(f"HostName: \t\t{my_system.Name}")
	info.append(f"Manufacturer: \t\t{my_system.Manufacturer}")
	info.append(f"Model: \t\t\t{my_system.Model}")
	info.append(f"NumberOfProcessors: \t{my_system.NumberOfProcessors}")
	info.append(f"SystemType: \t\t{my_system.SystemType}")
	info.append(f"Windows: \t\t{os_system.BuildType}")
	info.append(f"\t\t\t{os_system.InstallDate}")
	info.append(f"\t\t\t{os_system.Version}")
	info.append(f"\t\t\t{os_system.Name}")
		
	for item in info:
		print(item)