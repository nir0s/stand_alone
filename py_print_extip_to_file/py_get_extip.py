import requests
import socket
import csv
import re
from datetime import datetime
import os
import wmi

def get_cpu_load():
    """ Returns a list CPU Loads"""
    result = []
    cmd = "WMIC CPU GET LoadPercentage "
    response = os.popen(cmd + ' 2>&1','r').read().strip().split("\r\n")
    for load in response[1:]:
       result.append(int(load))
    return result

now = datetime.now()
time_stamp = now.strftime('%Y_%m_%d-%H%M')

hostname = socket.gethostname()
int_ip = socket.gethostbyname(hostname)
cpu_load = str(get_cpu_load())


path_ip_prov_list = 'D:\Dropbox\scripts\standalone\py_print_extip_to_file\ip_prov_urls.txt'
path_extip_file = 'd:\\dropbox\\system_status_report\\' + hostname + '.log'

with open(path_ip_prov_list, 'rb') as f:
	ip_prov_list = csv.reader(f, delimiter=',', quoting=csv.QUOTE_NONE)
	for ip_prov in ip_prov_list:
		ip_url = ip_prov[0]
		r = requests.get(ip_url)
		ext_ip = str(re.findall( r'[0-9]+(?:\.[0-9]+){3}', r.text ))
		if not r.status_code == 200:
			continue	
		else:
			with open(path_extip_file, 'a') as f:
				f.write(time_stamp + '\n')
				f.write('ext_ip:' + ip + '\n')
				f.write('int_ip:' + int_ip + '\n')
				f.write('cpu_load:' + cpu_load + '\n')
				f.write('\n')
			break