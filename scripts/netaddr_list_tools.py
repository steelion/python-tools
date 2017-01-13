# Python
#   netaddr_list_tools.py
#
# Das Script eliminiert die Überlappenden Konfigurationen
# und gibt die neue Liste der IP-Adressen und IP-Ranges auf der Console aus
#

from netaddr import *
import re
import pprint


# ip_addr: Hier die ausgehende Liste der IP-Ranges eintragen
ip_addr = "10.0.65.112/32 10.0.65.161/32 10.0.65.162/32 10.0.65.164/32 10.0.84.0/24 10.0.85.0/24 10.153.129.99/32 10.2.224.0/22 10.2.228.0/24 10.2.229.0/24 10.2.230.0/24 10.2.231.0/24 10.3.129.127/32 10.3.129.32/27 40.40.0.216/29 10.3.128.0/23 40.40.32.33 40.40.32.5 40.40.32.32 40.40.32.34 40.40.32.32 40.40.32.21 40.40.32.25 40.40.32.23 10.21.1.136"

ip_list = []
ip_ranges_list = ip_addr.split()
for ip_entry in ip_ranges_list:
    regex_pattern_range = re.compile("(\d+\.\d+\.\d+\.\d+)/(\d+)")
    regex_match_range = regex_pattern_range.match( ip_entry )
    regex_pattern_address = re.compile("(\d+\.\d+\.\d+\.\d+)")
    regex_match_address = regex_pattern_address.match( ip_entry )

    if regex_match_range:
        ip_list.append(IPNetwork(ip_entry))
    elif regex_match_address:
        ip_list.append(IPAddress(ip_entry))
    else:
        print("ERROR: entry: '" + ip_entry + "'")

# pprint.pprint( cidr_merge(ip_list))
#[IPNetwork('192.0.2.0/23'), IPNetwork('192.0.4.0/24'), IPNetwork('fe80::/120')]

ip_list_merged = cidr_merge(ip_list)
ip_addr_string_merged = ""
ip_addr_string_pos1 = True

for ip_entry in ip_list_merged:
    if not ip_addr_string_pos1:
        ip_addr_string_merged += " "
    else:
        ip_addr_string_pos1 = False

    if ip_entry.size == 1:
        ip_addr_string_merged += str(ip_entry.ip)
    else:
        ip_addr_string_merged += str(ip_entry.cidr)


# Hier wird die neue Liste der Adressen und Ranges ausgegeben.
print( ip_addr_string_merged )
